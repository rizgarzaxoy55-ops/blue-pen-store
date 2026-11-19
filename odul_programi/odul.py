# -*- coding: utf-8 -*-
"""
Zakho Üniversitesi - Yayın Ödülü (خەلاتکرن) yazısı otomatik hazırlama programı.

Ne yapar:
  1) Yazarların getirdiği dilekçenin (kitap) resmini/PDF'ini okur (Claude ile):
     yazar adları, paper başlığı, dergi, IF / index, bölüm, fakülte, el yazısı para.
  2) Paperi internette arar (Claude + web arama): yayınlanmış mı, DOI, dergideki link,
     yazarlar doğru mu, Clarivate IF, Scopus quartile, DOAJ, University of Zakho
     staff portal linkleri.
  3) Ödül miktarını üniversitenin kurallarına göre KENDİSİ hesaplar
     (Clarivate IF tablosu > Scopus Q > üniversite dergileri listesi).
  4) Örnek Word'ü (sablon.docx) doldurur ve yazarların adını taşıyan bir klasöre koyar:
        cikti/(yazar1) و (yazar2)/
            (yazar1) و (yazar2) - خەلاتکرن.docx
            rapor.txt          <- linkler + kontrol sonuçları + uyarılar
            veri.json          <- okunan/bulunan tüm bilgiler (elle düzeltip tekrar kullanılabilir)
            <dilekçe resmi>

Kullanım:
    python odul.py dilekce.jpg [dilekce2.pdf ...]
    python odul.py --json cikti/.../veri.json      # (elle düzeltilmiş) veriden Word'ü yeniden üret
    python odul.py --no-web dilekce.jpg             # internette arama yapmadan
API anahtarı: ANTHROPIC_API_KEY ortam değişkeni ya da ayarlar.json içindeki "anthropic_api_key".
"""
import argparse
import base64
import datetime
import difflib
import io
import json
import math
import os
import re
import shutil
import sys
import zipfile
from typing import List, Literal, Optional

from pydantic import BaseModel, Field

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "sablon.docx")
OUT_ROOT = os.path.join(HERE, "cikti")
SETTINGS = os.path.join(HERE, "ayarlar.json")
MODEL = "claude-opus-5-5"

# ─────────────────────────── Ödül kuralları ───────────────────────────
# Fermana wezarî (1692) 23/01/2025 - "onlara gore para belirleme.pdf"
SCOPUS_AMOUNT = {1: 500_000, 2: 400_000, 3: 300_000, 4: 300_000, None: 250_000}
UNI_KRG_AMOUNT = 250_000      # Kürdistan Bölgesi bakanlıkça onaylı üniversite dergileri
UNI_IRAQ_AMOUNT = 200_000     # Irak devlet üniversiteleri / Irak dışı üniversiteler
MAX_AUTHORS = {"pure": 4, "humanities": 3}

COLLEGES = {
    # anahtar: (Kürtçe, Arapça, İngilizce, bilim türü)
    "Science": ("زانست", "العلوم", "Science", "pure"),
}


def clarivate_amount(impact: float) -> int:
    """IF tablosu: 0.1-1 = 500,000 ; her +0.1 = +20,000 ; 6 ve üstü = 1,500,000."""
    if impact <= 1.0:
        return 500_000
    if impact >= 6.0:
        return 1_500_000
    tenth = math.floor(impact * 10 + 1e-9)          # 2.75 -> 27 (2.7)
    return 500_000 + 20_000 * (tenth - 10)


def decide_award(d: dict) -> dict:
    """Hangi kategori, kaç para, hangi madde (خالا). İnternetten doğrulanan değerler
    varsa onları, yoksa dilekçede yazanları kullanır."""
    v = d.get("verify") or {}
    L = d["letter"]
    impact = v.get("impact_factor") if v.get("impact_factor") else L.get("stated_impact_factor")
    clar = v.get("clarivate_listed")
    if impact and clar is not False:
        imp = float(impact)
        return {"kind": "Clarivate", "point": 1, "amount": clarivate_amount(imp),
                "index_text": f"Clarivate (Impact Factor = {imp:g})"}
    scopus = v.get("scopus_listed")
    q = v.get("scopus_quartile") or L.get("stated_scopus_quartile")
    if scopus or q or "scopus" in (L.get("stated_index") or "").lower():
        q = int(q) if q else None
        return {"kind": "Scopus", "point": 2, "amount": SCOPUS_AMOUNT[q],
                "index_text": f"Scopus Q{q}" if q else "Scopus"}
    krg = find_in_krg_list(L.get("journal", ""))
    if krg:
        return {"kind": "KRG", "point": 3, "amount": UNI_KRG_AMOUNT,
                "index_text": f"گۆڤارێن زانکۆیێن هەرێمێ ({krg['u']})"}
    return {"kind": "Iraq", "point": 3, "amount": UNI_IRAQ_AMOUNT,
            "index_text": "گۆڤارێن زانکۆیێن عیراقێ"}


# ─────────────────────────── Yardımcılar ───────────────────────────
_MAP = str.maketrans({"ی": "ي", "ى": "ي", "ێ": "ي", "ک": "ك", "ە": "ه", "ۆ": "و", "ڤ": "ف",
                      "ڕ": "ر", "ڵ": "ل", "گ": "ك", "پ": "ب", "چ": "ج", "ژ": "ز", "ة": "ه",
                      "أ": "ا", "إ": "ا", "آ": "ا"})


def norm(s):
    s = (s or "").lower().translate(_MAP)
    s = re.sub(r"[‌‍ً-ٟ]", "", s)
    s = re.sub(r"\b(the|of|and|for|journal)\b", " ", s)
    return re.sub(r"[^\w]+", " ", s).strip()


def find_in_krg_list(journal):
    try:
        with open(os.path.join(HERE, "journals_krg.json"), encoding="utf-8") as f:
            lst = json.load(f)
    except FileNotFoundError:
        return None
    for e in lst:
        for n in e["n"]:
            if difflib.SequenceMatcher(None, norm(journal), norm(n)).ratio() >= 0.85:
                return e
    return None


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|\n\r\t]+', " ", s).strip()[:150]


def api_client():
    import anthropic
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key and os.path.exists(SETTINGS):
        with open(SETTINGS, encoding="utf-8") as f:
            key = (json.load(f).get("anthropic_api_key") or "").strip()
    return anthropic.Anthropic(api_key=key) if key else anthropic.Anthropic()


# ─────────────────────────── 1) Dilekçeyi okuma ───────────────────────────
class Author(BaseModel):
    name_ku: str = Field(description="Adı dilekçede Arap/Kürt harfleriyle nasıl yazılmışsa öyle, harfi harfine")
    name_en: str = Field(description="Adı dilekçede Latin harfleriyle nasıl yazılmışsa öyle; yoksa boş")


class Letter(BaseModel):
    authors: List[Author]
    title: str = Field(description="Paperin tam başlığı, dilekçedeki gibi")
    journal: str
    stated_impact_factor: Optional[float] = None
    stated_index: str = Field(description="Dilekçede yazan index bilgisi (Clarivate/Scopus Q?/DOAJ...), yoksa boş")
    stated_scopus_quartile: Optional[int] = None
    department_ku: str = Field(description="Bölüm adı Badini Kürtçesiyle, örn. فيزيك ، ماتماتیک")
    college_key: Literal["Science", "Other"]
    college_name: str
    handwritten_amount: Optional[int] = Field(None, description="Kâğıda elle yazılmış para miktarı (dinar), yoksa null")
    letter_date: str = ""
    unclear: List[str] = Field(default_factory=list, description="Okunması zor/şüpheli kısımlar")


READ_PROMPT = """Bu, University of Zakho'da bir hocanın yayınlanmış makalesi için ödül (خەلاتکرن) \
istediği dilekçedir (Badini Kürtçesi / Arapça / İngilizce). İçindeki bilgileri çıkar.
- Yazar adlarını dilekçede yazıldığı gibi, harfi harfine kopyala; düzeltme yapma.
- Bölüm adını Kürtçe ver (örn. فيزيك). Fakülte Fen (Science / زانست / العلوم) ise college_key = "Science".
- Kâğıda yeşil/mavi kalemle elle yazılmış para varsa handwritten_amount'a sadece sayı olarak yaz.
- Emin olmadığın her şeyi unclear listesine yaz."""


def file_blocks(path):
    ext = os.path.splitext(path)[1].lower()
    with open(path, "rb") as f:
        data = f.read()
    if ext == ".pdf":
        return [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                 "data": base64.standard_b64encode(data).decode()}}]
    if ext == ".docx":
        from docx import Document
        doc = Document(path)
        txt = "\n".join(p.text for p in doc.paragraphs)
        return [{"type": "text", "text": "Dilekçe metni:\n" + txt}]
    # resim: gerekirse küçült (API sınırı ~5MB)
    from PIL import Image, ImageOps
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((2400, 2400))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=90)
    return [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                         "data": base64.standard_b64encode(buf.getvalue()).decode()}}]


def read_letter(client, path) -> dict:
    resp = client.messages.parse(
        model=MODEL, max_tokens=16000,
        output_config={"effort": "high"},
        messages=[{"role": "user", "content": file_blocks(path) + [{"type": "text", "text": READ_PROMPT}]}],
        output_format=Letter,
    )
    if resp.stop_reason == "refusal" or resp.parsed_output is None:
        raise RuntimeError(f"Dilekçe okunamadı (stop_reason={resp.stop_reason})")
    return resp.parsed_output.model_dump()


# ─────────────────────────── 2) İnternette doğrulama ───────────────────────────
class PortalLink(BaseModel):
    author: str
    url: Optional[str] = None
    paper_listed: Optional[bool] = Field(None, description="Bu paper o sayfada kayıtlı görünüyor mu")


class Verify(BaseModel):
    published: bool
    doi: Optional[str] = None
    article_url: Optional[str] = Field(None, description="Derginin sitesindeki makale linki")
    published_date: Optional[str] = None
    authors_on_paper: List[str] = Field(default_factory=list)
    authors_match: Optional[bool] = None
    uoz_affiliation: Optional[bool] = Field(None, description="Yazarların adresinde University of Zakho yazıyor mu")
    clarivate_listed: Optional[bool] = None
    impact_factor: Optional[float] = None
    impact_factor_year: Optional[str] = None
    impact_factor_source: Optional[str] = None
    scopus_listed: Optional[bool] = None
    scopus_quartile: Optional[int] = None
    doaj_listed: Optional[bool] = None
    staff_portal: List[PortalLink] = Field(default_factory=list)
    notes: List[str] = Field(default_factory=list, description="Şüpheli/doğrulanamayan her şey")


def research_prompt(L):
    authors = "; ".join(f"{a['name_en']} ({a['name_ku']})" for a in L["authors"])
    return f"""University of Zakho'da bir yayın ödülü başvurusunu doğrula. İnternette ara ve sayfaları aç.

Paper: {L['title']}
Dergi: {L['journal']}
Yazarlar: {authors}
Dilekçede yazan: IF={L.get('stated_impact_factor')} ; index={L.get('stated_index')}

Bul ve rapor et:
1. Paper gerçekten yayınlanmış mı? DOI'si ve derginin sitesindeki makale linki (doi.org değil, yayıncının sayfası).
2. Makaledeki yazar listesi ve adresleri; dilekçedeki yazarlarla aynı mı, University of Zakho adresi var mı.
3. Derginin Clarivate (Web of Science/JCR) Impact Factor'ü ve yılı; Scopus'ta mı, quartile (Q1-Q4); DOAJ'da mı.
   Kaynağını yaz. Emin değilsen tahmin etme, "doğrulanamadı" de.
4. Her yazarın University of Zakho staff portal sayfası (staffportal.uoz.edu.krd/...) ve bu paper orada kayıtlı mı.
Kısa, madde madde, linklerle yaz."""


def research(client, L) -> str:
    tools = [{"type": "web_search_20260209", "name": "web_search", "max_uses": 12},
             {"type": "web_fetch_20260209", "name": "web_fetch", "max_uses": 12}]
    messages = [{"role": "user", "content": research_prompt(L)}]
    for _ in range(6):
        with client.beta.messages.stream(
            model=MODEL, max_tokens=32000,
            output_config={"effort": "high"},
            betas=["server-side-fallback-2026-07-01"], fallbacks="default",
            tools=tools, messages=messages,
        ) as stream:
            resp = stream.get_final_message()
        if resp.stop_reason == "pause_turn":
            messages.append({"role": "assistant", "content": resp.content})
            continue
        break
    if resp.stop_reason == "refusal":
        raise RuntimeError("Arama reddedildi (refusal).")
    return "\n".join(b.text for b in resp.content if b.type == "text")


def structure_research(client, L, text) -> dict:
    resp = client.messages.parse(
        model=MODEL, max_tokens=16000,
        output_config={"effort": "medium"},
        messages=[{"role": "user", "content":
                   "Aşağıdaki araştırma raporunu alanlara dök. Raporda olmayan hiçbir şeyi uydurma; "
                   "bilinmeyen alanları null bırak.\n\nYazarlar (dilekçe): "
                   + ", ".join(a["name_en"] or a["name_ku"] for a in L["authors"])
                   + "\n\nRAPOR:\n" + text}],
        output_format=Verify,
    )
    if resp.parsed_output is None:
        raise RuntimeError("Araştırma sonucu yapılandırılamadı.")
    out = resp.parsed_output.model_dump()
    out["raw_report"] = text
    return out


# ─────────────────────────── 3) Word doldurma ───────────────────────────
def _ptext(p):
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))


def _set_para(p, new):
    """Paragrafın ilk metin parçasına tüm yeni metni koy, diğerlerini boşalt (biçim korunur)."""
    ts = list(re.finditer(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))
    out, last = [], 0
    for i, m in enumerate(ts):
        out.append(p[last:m.start()])
        out.append('<w:t xml:space="preserve">%s</w:t>' % (_esc(new) if i == 0 else ""))
        last = m.end()
    out.append(p[last:])
    return "".join(out)


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def fill_word(d, out_path, today):
    L, A = d["letter"], d["award"]
    names = [a["name_ku"] for a in L["authors"]]
    if len(names) > 1:
        names_txt = "بەرێزان " + " و ".join(f"({n})" for n in names)
    else:
        names_txt = f"بەرێز ({names[0]})"
    col = COLLEGES.get(L.get("college_key"), COLLEGES["Science"])
    dept = L["department_ku"]
    body = (f"هيڤيدارين رازيبن لسەر خەلاتكرنا دەستگێريا ڤەكۆلينا هاتيیە بەلافكرن ژ لايێ {names_txt}"
            f" دڤەكولينێدا ل پشكا {dept} ل كوليژا {col[0]} ل زانکۆیا زاخۆ ب كوژمێ({A['amount']:,}) دينار"
            f" لديڤ برگا  (1)خالا ({A['point']}) ژ فەرمانا وەزاری و توماركرى ل پشكا زانستى ل ژێر ناڤ و نيشان :   ")
    title = re.sub(r"\s+", " ", L["title"]).strip()
    LRM = "‎"
    ky = today.year + 700

    rules = [
        (lambda t: t.startswith("هيڤيدارين رازيبن"), lambda p: _set_para(p, body)),
        (lambda t: "Novel Conjugate Gradient" in t,
         lambda p: _set_para(p, f"{LRM}({title}){LRM}")),
        (lambda t: t.strip() == "(Iraqi Journal of Science)",
         lambda p: _set_para(p, f"{LRM}({L['journal']}){LRM}")),
        (lambda t: "بو زانین هاتیە" in t and "index" in t,
         lambda p: _set_para(p, f"   بو زانین هاتیە index کرن {A['index_text']} ")),
        (lambda t: t.startswith("رێكه") and "زایینى" in t,
         lambda p: _set_para(p, f"رێكه‌وت: {today.day:02d} /  {today.month:02d}/ {today.year} ى  زایینى")),
        (lambda t: "كوردى" in t and re.search(r"27\d\d", t),
         lambda p: _set_para(p, f"                  /        / {ky}ی  كوردى")),
        (lambda t: re.fullmatch(r"\s*\d{4}/\d{2}/\d{2}\s*", t) is not None,
         lambda p: _set_para(p, today.strftime("%Y/%m/%d"))),
        (lambda t: t.strip().startswith("- پشكا"), lambda p: _set_para(p, f"- پشكا {dept}.")),
    ]
    hits = set()

    def repl(m):
        p = m.group(0)
        t = _ptext(p)
        for k, (cond, fn) in enumerate(rules):
            if cond(t):
                hits.add(k)
                return fn(p)
        return p

    zin = zipfile.ZipFile(TEMPLATE)
    xml = zin.read("word/document.xml").decode("utf-8")
    xml = re.sub(r"<w:p[ >].*?</w:p>", repl, xml, flags=re.S)
    if len(hits) != len(rules):
        missing = [k for k in range(len(rules)) if k not in hits]
        raise RuntimeError(f"sablon.docx beklenen yerleri içermiyor (eksik kural: {missing}). "
                           "Şablonu değiştirdiyseniz odul.py içindeki kuralları güncelleyin.")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item)
            zout.writestr(item, data)


# ─────────────────────────── 4) Kontroller ve rapor ───────────────────────────
def checks(d):
    L, V, A = d["letter"], d.get("verify") or {}, d["award"]
    warn = []
    col = COLLEGES.get(L.get("college_key"))
    if not col:
        warn.append(f"Fakülte '{L.get('college_name')}' için şablon yok; Word Fen Fakültesi şablonuyla yapıldı. "
                    "Başlığı ve dekan adını elle düzeltin.")
    kind = (col or COLLEGES["Science"])[3]
    if len(L["authors"]) > MAX_AUTHORS[kind]:
        warn.append(f"Yazar sayısı {len(L['authors'])}; kurala göre en fazla {MAX_AUTHORS[kind]} yazara ödül verilir.")
    hw = L.get("handwritten_amount")
    if hw and hw != A["amount"]:
        warn.append(f"Elle yazılan para ({hw:,}) hesaplanan paradan ({A['amount']:,}) farklı!")
    sif, vif = L.get("stated_impact_factor"), V.get("impact_factor")
    if sif and vif and abs(float(sif) - float(vif)) > 0.05:
        warn.append(f"Dilekçedeki IF ({sif}) ile internette bulunan IF ({vif}) farklı. "
                    f"Hesap {vif} ile yapıldı. Clarivate JCR'den kontrol edin.")
    if V:
        if not V.get("published"):
            warn.append("Paper internette yayınlanmış olarak BULUNAMADI.")
        if V.get("authors_match") is False:
            warn.append("Makaledeki yazarlar dilekçedekilerle eşleşmiyor: " + ", ".join(V.get("authors_on_paper", [])))
        if V.get("uoz_affiliation") is False:
            warn.append("Makalede University of Zakho adresi görünmüyor (kural 3).")
        for s in V.get("staff_portal", []):
            if not s.get("url"):
                warn.append(f"{s['author']}: staff portal sayfası bulunamadı.")
            elif s.get("paper_listed") is False:
                warn.append(f"{s['author']}: paper staff portalda kayıtlı görünmüyor.")
        warn += V.get("notes", [])
    else:
        warn.append("İnternette doğrulama yapılmadı.")
    for u in L.get("unclear", []):
        warn.append("Okuma şüphesi: " + u)
    return warn


def write_report(d, path, warnings):
    L, V, A = d["letter"], d.get("verify") or {}, d["award"]
    lines = [
        "YAYIN ÖDÜLÜ KONTROL RAPORU",
        f"Tarih: {datetime.date.today():%Y-%m-%d}", "",
        "Yazarlar: " + " ; ".join(f"{a['name_ku']} ({a['name_en']})" for a in L["authors"]),
        "Paper   : " + L["title"],
        "Dergi   : " + L["journal"],
        f"Bölüm   : {L['department_ku']} / {L.get('college_name', '')}", "",
        f"Kategori: {A['kind']}  -> {A['index_text']}",
        f"Madde   : برگا (1) خالا ({A['point']})",
        f"ÖDÜL    : {A['amount']:,} dinar"
        + (f"   (kâğıtta elle yazılan: {L['handwritten_amount']:,})" if L.get("handwritten_amount") else ""),
        "", "── LİNKLER ──",
        "Dergideki link : " + (V.get("article_url") or "-"),
        "DOI            : " + (f"https://doi.org/{V['doi']}" if V.get("doi") else "-"),
    ]
    for s in V.get("staff_portal", []):
        lines.append(f"Staff portal   : {s['author']}: {s.get('url') or 'bulunamadı'}")
    lines += ["", "── DOĞRULAMA ──",
              f"Yayınlanmış     : {V.get('published')}",
              f"Yayın tarihi    : {V.get('published_date')}",
              f"Yazarlar (paper): {', '.join(V.get('authors_on_paper', [])) or '-'}",
              f"Yazarlar uyumlu : {V.get('authors_match')}",
              f"UoZ adresi      : {V.get('uoz_affiliation')}",
              f"Clarivate       : {V.get('clarivate_listed')}  IF={V.get('impact_factor')} "
              f"({V.get('impact_factor_year') or ''}) kaynak: {V.get('impact_factor_source') or '-'}",
              f"Scopus          : {V.get('scopus_listed')}  Q{V.get('scopus_quartile') or '-'}",
              f"DOAJ            : {V.get('doaj_listed')}",
              "", "── UYARILAR (elle kontrol edin) ──"]
    lines += [f"! {w}" for w in warnings] or ["(yok)"]
    if V.get("raw_report"):
        lines += ["", "── ARAŞTIRMA NOTLARI ──", V["raw_report"]]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ─────────────────────────── Ana akış ───────────────────────────
def build_outputs(d, src_file=None, today=None):
    today = today or datetime.date.today()
    d["award"] = decide_award(d)
    names = [a["name_ku"] for a in d["letter"]["authors"]]
    folder_name = safe_name(" و ".join(f"({n})" for n in names))
    folder = os.path.join(OUT_ROOT, folder_name)
    os.makedirs(folder, exist_ok=True)
    docx_path = os.path.join(folder, f"{folder_name} - خەلاتکرن.docx")
    fill_word(d, docx_path, today)
    warnings = checks(d)
    write_report(d, os.path.join(folder, "rapor.txt"), warnings)
    with open(os.path.join(folder, "veri.json"), "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    if src_file and os.path.abspath(os.path.dirname(src_file)) != os.path.abspath(folder):
        shutil.copy2(src_file, os.path.join(folder, "dilekce" + os.path.splitext(src_file)[1].lower()))
    return folder, docx_path, warnings


def process_file(client, path, use_web=True):
    print(f"\n=== {os.path.basename(path)}")
    print("  1/3 dilekçe okunuyor...")
    L = read_letter(client, path)
    print("      yazarlar:", ", ".join(a["name_ku"] for a in L["authors"]))
    print("      paper   :", L["title"])
    d = {"letter": L, "verify": None}
    if use_web:
        print("  2/3 paper internette aranıyor (1-3 dakika sürebilir)...")
        try:
            d["verify"] = structure_research(client, L, research(client, L))
        except Exception as e:  # arama başarısız olsa da Word yine üretilsin
            print("      ! arama başarısız:", e)
    print("  3/3 Word hazırlanıyor...")
    return build_outputs(d, path)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Zakho Üniversitesi yayın ödülü yazısı hazırlayıcı")
    ap.add_argument("files", nargs="*", help="dilekçe resmi / PDF / docx")
    ap.add_argument("--json", help="veri.json'dan Word'ü yeniden üret (API kullanmaz)")
    ap.add_argument("--no-web", action="store_true", help="internette doğrulama yapma")
    a = ap.parse_args()

    results = []
    if a.json:
        with open(a.json, encoding="utf-8") as f:
            results.append(build_outputs(json.load(f)))
    else:
        files = a.files
        if not files:
            in_dir = os.path.join(HERE, "girdi")
            os.makedirs(in_dir, exist_ok=True)
            files = [os.path.join(in_dir, x) for x in sorted(os.listdir(in_dir))
                     if x.lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".pdf", ".docx"))]
            if not files:
                print("Dosya verilmedi. Dilekçe resimlerini 'girdi' klasörüne koyun ya da run.bat üzerine sürükleyin.")
                return 1
        client = api_client()
        for p in files:
            try:
                results.append(process_file(client, p, use_web=not a.no_web))
            except Exception as e:
                print(f"  HATA ({os.path.basename(p)}): {e}")

    for folder, docx_path, warnings in results:
        print("\n✔", docx_path)
        for w in warnings:
            print("   !", w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
