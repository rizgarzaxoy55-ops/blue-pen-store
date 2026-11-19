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
    doaj = L.get("stated_doaj") or v.get("doaj_listed")
    krg = find_in_krg_list(L.get("journal", "")) or find_in_krg_list(v.get("journal") or "")
    if krg:
        return {"kind": "KRG", "point": 3, "amount": UNI_KRG_AMOUNT, "krg_university": krg["u"],
                "index_text": "DOAJ" if doaj else f"گۆڤارێن زانکۆیێن هەرێمێ ({krg['u']})"}
    return {"kind": "Iraq", "point": 3, "amount": UNI_IRAQ_AMOUNT,
            "index_text": "DOAJ" if doaj else "گۆڤارێن زانکۆیێن عیراقێ"}


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


def api_key():
    key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not key and os.path.exists(SETTINGS):
        with open(SETTINGS, encoding="utf-8") as f:
            key = (json.load(f).get("anthropic_api_key") or "").strip()
    return key


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


READ_PROMPT = """Bu (Word dosyası, PDF ya da resim), University of Zakho'da bir hocanın yayınlanmış makalesi için ödül (خەلاتکرن) \
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
        return [{"type": "text", "text": "Dilekçe (Word dosyası) metni:\n" + docx_text(path)}]
    if ext == ".doc":
        raise RuntimeError("Eski .doc biçimi desteklenmiyor. Word'de açıp 'Farklı Kaydet > .docx' yapın.")
    # resim: gerekirse küçült (API sınırı ~5MB)
    from PIL import Image, ImageOps
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((2400, 2400))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=90)
    return [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                         "data": base64.standard_b64encode(buf.getvalue()).decode()}}]


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"


def _para_text(p):
    out = []
    for el in p.iter():
        if el.tag == W + "t":
            out.append(el.text or "")
        elif el.tag in (W + "tab", W + "ptab"):
            out.append("\t")
        elif el.tag in (W + "br", W + "cr"):
            out.append("\n")
    return "".join(out)


def _block_lines(elem, lines):
    """Paragrafları ve tabloları sırayla yaz; tablo satırlarını ' | ' ile birleştir.
    Metin kutularını (textbox) da okur, Word'ün yedek kopyalarını (mc:Fallback) atlar."""
    for child in elem:
        if child.tag == MC_FALLBACK:
            continue
        if child.tag == W + "p":
            own = _para_text_no_boxes(child)
            if own.strip():
                lines.append(own)
            for box in child.iter(W + "txbxContent"):
                if not any(a.tag == MC_FALLBACK for a in box.iterancestors()):
                    _block_lines(box, lines)
        elif child.tag == W + "tbl":
            for tr in child.iter(W + "tr"):
                cells = [" ".join(_para_text(p) for p in tc.iter(W + "p")).strip() for tc in tr.findall(W + "tc")]
                if any(cells):
                    lines.append(" | ".join(cells))
        else:
            _block_lines(child, lines)


def _para_text_no_boxes(p):
    import copy
    q = copy.deepcopy(p)
    for box in list(q.iter(W + "txbxContent")):
        box.getparent().remove(box)
    return _para_text(q)


def docx_text(path) -> str:
    """Word dosyasındaki tüm metni (başlık/alt bilgi, tablolar, metin kutuları dahil) çıkar."""
    from lxml import etree
    z = zipfile.ZipFile(path)
    names = z.namelist()
    parts = ([n for n in names if re.fullmatch(r"word/header\d*\.xml", n)]
             + ["word/document.xml"]
             + [n for n in names if re.fullmatch(r"word/footer\d*\.xml", n)])
    out = []
    for n in parts:
        if n not in names:
            continue
        root = etree.fromstring(z.read(n))
        lines = []
        _block_lines(root, lines)
        if lines:
            label = "ÜST BİLGİ" if "header" in n else ("ALT BİLGİ" if "footer" in n else "METİN")
            out.append(f"[{label}]\n" + "\n".join(lines))
    return "\n\n".join(out)


TESSERACT_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                   r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                   os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")]


OCR_NOTES = []


def ocr_images(images):
    try:
        import pytesseract
    except ImportError:
        raise RuntimeError("OCR için 'pip install pytesseract' gerekli.")
    if not shutil.which("tesseract"):
        for p in TESSERACT_PATHS:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break
    try:
        langs = set(pytesseract.get_languages(config=""))
    except Exception:
        raise RuntimeError("Dilekçe taranmış (resim). Okumak için ücretsiz Tesseract OCR programını kurun: "
                           "https://github.com/UB-Mannheim/tesseract/wiki  (kurarken 'Arabic' dilini de seçin). "
                           "Ya da dilekçenin Word (.docx) halini kullanın.")
    arabic = "ara" if "ara" in langs else next(
        (l.replace("\\", "/") for l in sorted(langs) if l.replace("\\", "/").lower().endswith("arabic")), None)
    if arabic:
        lang = "eng+" + arabic
    else:
        lang = "eng"
        msg = ("Tesseract'ta Arabic dili kurulu değil: Kürtçe kısımlar (bölüm, Kürtçe adlar, başlık etiketi) "
               "okunamaz. Tesseract kurulum dosyasını tekrar çalıştırıp 'Additional language data' altından "
               "'Arabic'i işaretleyin.")
        print("      ! " + msg)
        OCR_NOTES.append(msg)
    return "\n".join(pytesseract.image_to_string(im, lang=lang, config="--psm 4") for im in images)


def free_text(path):
    """Ücretsiz mod: (metin, kaynak). Word -> doğrudan; metinli PDF -> doğrudan; taranmış PDF/resim -> OCR."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return docx_text(path), "docx"
    if ext == ".doc":
        raise RuntimeError("Eski .doc biçimi desteklenmiyor. Word'de açıp 'Farklı Kaydet > .docx' yapın.")
    if ext == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(path)
        pages = [pdf[i] for i in range(min(len(pdf), 3))]
        text = "\n".join(p.get_textpage().get_text_range() for p in pages)
        if len(re.sub(r"\s+", "", text)) > 80:
            return text, "pdf"
        print("      PDF taranmış, OCR ile okunuyor (biraz sürebilir)...")
        return ocr_images([p.render(scale=300 / 72).to_pil() for p in pages]), "ocr"
    if ext in (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"):
        from PIL import Image, ImageOps
        print("      resim OCR ile okunuyor...")
        return ocr_images([ImageOps.exif_transpose(Image.open(path)).convert("RGB")]), "ocr"
    raise RuntimeError(f"Desteklenmeyen dosya türü: {ext}")


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
    names = [a["name_ku"] or a["name_en"] for a in L["authors"]] or ["....."]
    if len(names) > 1:
        names_txt = "بەرێزان " + " و ".join(f"\u200f({n})\u200f" for n in names)
    else:
        names_txt = f"بەرێز \u200f({names[0]})\u200f"
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
    if A["kind"] == "Iraq":
        warn.append("Dergide IF/Scopus yok ve Kürdistan onaylı dergi listesinde bulunamadı. 200,000 sadece Irak devlet "
                    "üniversiteleri ya da yurt dışı üniversite dergileri için geçerli; derginin hangi kuruma ait "
                    "olduğunu kontrol edin.")
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
            warn.append("Makaledeki yazarlar dilekçedekilerle EŞLEŞMİYOR. Makaledeki yazarlar: "
                        + ", ".join(V.get("authors_on_paper", [])))
        if V.get("uoz_affiliation") is False:
            warn.append("Makalede University of Zakho adresi görünmüyor (kural 3).")
        for s in V.get("staff_portal", []):
            if s.get("candidates"):
                warn.append(f"{s['author']}: staff portalda kesin eşleşme yok, benzer kişi(ler) var; "
                            "rapordaki linklerden doğru olanı kontrol edin.")
            elif s.get("no_english_name"):
                warn.append(f"{s['author']}: İngilizce adı bilinmediği için staff portalda aranamadı; elle bakın.")
            elif not s.get("url"):
                warn.append(f"{s['author']}: staff portal sayfası bulunamadı.")
            elif s.get("paper_listed") is False:
                warn.append(f"{s['author']}: paper staff portalda kayıtlı DEĞİL (kural 5-b).")
        if V.get("mode") == "ucretsiz":
            reg = V.get("registry")
            if reg is None and "UoZ kayıt listesi indirilemedi" not in " ".join(V.get("notes", [])):
                warn.append("Paper UoZ yayın kayıt listesinde (Google Sheet) BULUNAMADI (kural 5-a).")
            elif reg and reg.get("impact_factor") and sif:
                try:
                    if abs(float(reg["impact_factor"]) - float(sif)) > 0.05:
                        warn.append(f"Kayıt listesinde IF={reg['impact_factor']}, dilekçede IF={sif}.")
                except ValueError:
                    pass
            if (A["kind"] != "Clarivate" and reg and reg.get("impact_factor")
                    and re.search(r"thomson|clarivate|web of science|wos|isi", reg.get("indexing", ""), re.I)):
                warn.append(f"Dilekçede IF yok ama UoZ kayıt listesinde bu paper için IF={reg['impact_factor']} "
                            f"({reg.get('indexing')}) yazıyor. Dergi Clarivate'te ise ödül IF tablosuna göre "
                            "hesaplanmalı; kontrol edip veri.json'a stated_impact_factor yazın.")
            if A["kind"] == "Clarivate":
                warn.append(f"IF ({sif}) dilekçeden alındı; ücretsiz modda resmi doğrulama yok. "
                            "Clarivate MJL/JCR'den kontrol edin (link raporda).")
            sjr = (V.get("sjr") or {}).get("quartile")
            if A["kind"] == "Scopus" and sjr and L.get("stated_scopus_quartile") and sjr != L["stated_scopus_quartile"]:
                warn.append(f"Dilekçede Scopus Q{L['stated_scopus_quartile']}, SCImago'da Q{sjr}. Scopus'tan kontrol edin.")
        warn += V.get("notes", [])
    else:
        warn.append("İnternette doğrulama yapılmadı.")
    no_ku = [a["name_en"] or a["name_ku"] for a in L["authors"] if not re.search(r"[\u0600-\u06FF]", a.get("name_ku", ""))]
    if no_ku:
        warn.append("Kürtçe adı bilinmeyen yazar(lar) Word'e İngilizce yazıldı: " + ", ".join(no_ku)
                    + ". veri.json'da name_ku'yu Kürtçe yazıp --json ile yeniden üretin (program bu adları öğrenir).")
    for u in L.get("unclear", []):
        if not u.startswith(("Kürtçe adı bilinmeyen", "Bazı yazarların Kürtçe adı yok")):
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
        "Başvuran: " + (L.get("applicant_ku") or "-")
        + ("   (yazarlar makaleden alındı: sadece Zakho'lu olanlar)" if L.get("authors_from_paper") else ""),
        "Dilekçedeki DOI: " + (", ".join(L.get("stated_dois") or []) or "-"),
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
        listed = {True: "  [paper KAYITLI]", False: "  [paper KAYITLI DEĞİL]"}.get(s.get("paper_listed"), "")
        lines.append(f"Staff portal   : {s['author']}: {s.get('url') or 'bulunamadı'}{listed}")
        for c in s.get("candidates", []):
            lines.append(f"                 olası: {c}")
    lines += ["", "── DOĞRULAMA ──",
              f"Yayınlanmış     : {V.get('published')}",
              f"Yayın tarihi    : {V.get('published_date')}",
              f"Yazarlar (paper): {', '.join(V.get('authors_on_paper', [])) or '-'}",
              f"Yazarlar uyumlu : {V.get('authors_match')}",
              f"UoZ adresi      : {V.get('uoz_affiliation')}",
              f"Clarivate       : {V.get('clarivate_listed')}  IF={V.get('impact_factor')} "
              f"({V.get('impact_factor_year') or ''}) kaynak: {V.get('impact_factor_source') or '-'}",
              f"Scopus          : {V.get('scopus_listed')}  Q{V.get('scopus_quartile') or '-'}",
              f"DOAJ            : {V.get('doaj_listed')}"]
    if V.get("issn"):
        lines.append(f"ISSN            : {', '.join(V['issn'])}")
    if V.get("openalex_affiliations"):
        lines.append("Adresler        : " + " | ".join(V["openalex_affiliations"]))
    if V.get("sjr"):
        sj = V["sjr"]
        lines.append(f"SCImago         : Q{sj.get('quartile') or '-'} ({sj.get('file')})")
    if V.get("mode") == "ucretsiz":
        reg = V.get("registry")
        lines.append("UoZ kayıt list. : " + (f"KAYITLI ({reg['timestamp']}) index={reg['indexing']} IF={reg['impact_factor']}"
                                             if reg else "bulunamadı"))
        if V.get("registry_journal_ifs"):
            lines.append("  aynı dergi için listede yazılan IF'ler: "
                         + ", ".join(f"{i} ({dt})" for dt, i in V["registry_journal_ifs"]))
    if V.get("links"):
        lines += ["", "── ELLE KONTROL LİNKLERİ ──"] + [f"{k}: {u}" for k, u in V["links"].items()]
    lines += ["", "── UYARILAR (elle kontrol edin) ──"]
    lines += [f"! {w}" for w in warnings] or ["(yok)"]
    if V.get("raw_report"):
        lines += ["", "── ARAŞTIRMA NOTLARI ──", V["raw_report"]]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ─────────────────────────── Ana akış ───────────────────────────
def unique_folder(folder_name, title):
    """Aynı yazarların başka bir paperi varsa üzerine yazma: '(2)', '(3)' ekle."""
    n = 1
    while True:
        cand = os.path.join(OUT_ROOT, folder_name + (f" ({n})" if n > 1 else ""))
        vj = os.path.join(cand, "veri.json")
        if not os.path.exists(vj):
            return cand
        try:
            with open(vj, encoding="utf-8") as f:
                if json.load(f)["letter"]["title"] == title:
                    return cand          # aynı paper: güncelle
        except Exception:
            pass
        n += 1


def build_outputs(d, src_file=None, today=None):
    today = today or datetime.date.today()
    d["award"] = decide_award(d)
    try:
        import ucretsiz
        ucretsiz.learn_names(d["letter"]["authors"])
    except Exception:
        pass
    names = [a["name_ku"] or a["name_en"] for a in d["letter"]["authors"]]
    folder_name = safe_name(" و ".join(f"({n})" for n in names)) or safe_name(d["letter"]["title"])[:60] or "isimsiz"
    folder = unique_folder(folder_name, d["letter"]["title"])
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
    """client=None ise ücretsiz mod (yapay zekâsız)."""
    print(f"\n=== {os.path.basename(path)}")
    print("  1/3 dilekçe okunuyor...")
    if client is None:
        import ucretsiz
        text, source = free_text(path)
        L = ucretsiz.parse_letter(text)
        L["source"] = source
        if source == "ocr":
            L["unclear"] += OCR_NOTES[-1:]
            L["unclear"].append("Dilekçe taranmış (resim); OCR ile okundu. Başlık, dergi ve adları Word'de dikkatle kontrol edin.")
    else:
        L = read_letter(client, path)
    print("      yazarlar:", ", ".join(a["name_ku"] for a in L["authors"])
          or "(dilekçede yazar listesi yok; makaleden Zakho'lu yazarlar alınacak)")
    print("      paper   :", L["title"])
    d = {"letter": L, "verify": None}
    if use_web:
        print("  2/3 paper internette aranıyor...")
        try:
            if client is None:
                d["verify"] = ucretsiz.verify(L)
                d["verify"]["mode"] = "ucretsiz"
            else:
                d["verify"] = structure_research(client, L, research(client, L))
                d["verify"]["mode"] = "ai"
        except Exception as e:  # arama başarısız olsa da Word yine üretilsin
            print("      ! arama başarısız:", e)
    V = d.get("verify") or {}
    if not L["authors"] and V.get("authors_on_paper"):
        L["authors"] = [{"name_ku": n, "name_en": n} for n in V["authors_on_paper"]]
        L.setdefault("unclear", []).append(
            "Dilekçede yazar adları okunamadı; yazarlar makaleden (Crossref) alındı ve Word'e İngilizce yazıldı. "
            "Kürtçe adları veri.json'a yazıp --json ile yeniden üretin.")
    print("  3/3 Word hazırlanıyor...")
    return build_outputs(d, path)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Zakho Üniversitesi yayın ödülü yazısı hazırlayıcı")
    ap.add_argument("files", nargs="*", help="dilekçe resmi / PDF / docx")
    ap.add_argument("--json", help="veri.json'dan Word'ü yeniden üret (API kullanmaz)")
    ap.add_argument("--no-web", action="store_true", help="internette doğrulama yapma")
    ap.add_argument("--ucretsiz", action="store_true", help="API anahtarı olsa bile ücretsiz modu kullan")
    ap.add_argument("--ai", action="store_true", help="yapay zekâ modunu zorla (API anahtarı gerekir)")
    a = ap.parse_args()

    results = []
    if a.json:
        with open(a.json, encoding="utf-8") as f:
            res = build_outputs(json.load(f))
        results.append(res)
        old = os.path.dirname(os.path.abspath(a.json))
        new = os.path.abspath(res[0])
        # isimler değiştiyse klasör adı da değişir: dilekçeyi taşı, eski klasörü sil
        if old != new and os.path.dirname(old) == os.path.abspath(OUT_ROOT):
            for fn in os.listdir(old):
                if fn.startswith("dilekce") and not os.path.exists(os.path.join(new, fn)):
                    shutil.move(os.path.join(old, fn), os.path.join(new, fn))
            shutil.rmtree(old, ignore_errors=True)
            print("Eski klasör yeni ada taşındı:", os.path.basename(new))
    else:
        files = a.files
        if not files:
            in_dir = os.path.join(HERE, "girdi")
            os.makedirs(in_dir, exist_ok=True)
            files = [os.path.join(in_dir, x) for x in sorted(os.listdir(in_dir))
                     if x.lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".pdf", ".docx", ".doc")) and not x.startswith("~$")]
            if not files:
                print("Dosya verilmedi. Dilekçe Word dosyalarını 'girdi' klasörüne koyun ya da calistir.bat üzerine sürükleyin.")
                return 1
        use_ai = a.ai or (bool(api_key()) and not a.ucretsiz)
        client = api_client() if use_ai else None
        print("Mod:", "YAPAY ZEKÂ (API anahtarı ile)" if use_ai else "ÜCRETSİZ (API anahtarı yok)")
        for p in files:
            if not os.path.isfile(p):
                print(f"\n=== {os.path.basename(p)}\n  HATA: dosya bulunamadı: {p}\n"
                      "  (İpucu: komutu yazıp dosyayı cmd penceresine sürükleyin, yol kendiliğinden yazılır.)")
                continue
            try:
                results.append(process_file(client, p, use_web=not a.no_web))
            except Exception as e:
                print(f"  HATA ({os.path.basename(p)}): {e.__class__.__name__}: {e}")

    for folder, docx_path, warnings in results:
        print("\n✔", docx_path)
        for w in warnings:
            print("   !", w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
