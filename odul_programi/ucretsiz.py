# -*- coding: utf-8 -*-
"""
Ücretsiz mod: API anahtarı ve yapay zekâ olmadan çalışır.
  - Dilekçe Word'ünü kurallarla okur.
  - Crossref + OpenAlex: paper yayınlanmış mı, DOI, makale linki, yazarlar, University of Zakho adresi.
  - UoZ yayın kayıt listesi (Google Sheet): paper üniversite listesine kaydedilmiş mi (kural 5-a).
  - UoZ staff portal: yazarın sayfası ve paper orada kayıtlı mı (kural 5-b).
  - DOAJ API: dergi DOAJ'da mı.
  - SCImago (isteğe bağlı, veriler/ klasörüne elle indirilen CSV): quartile.
Clarivate Impact Factor'ün ücretsiz resmi kaynağı yok; dilekçedeki değer kullanılır ve uyarı verilir.
"""
import csv
import difflib
import html
import io
import json
import os
import re
import time
import unicodedata
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "veriler")
UA = "UoZ-odul-programi/1.0 (yayin odulu kontrolu)"
UOZ_SHEET_CSV = ("https://docs.google.com/spreadsheets/d/"
                 "1q6_uje4RCXU9pn9JqIbDKxaJ2-doxgW095rMS2SqM4k/gviz/tq?tqx=out:csv")
UOZ_SHEET_VIEW = "https://docs.google.com/spreadsheets/d/1q6_uje4RCXU9pn9JqIbDKxaJ2-doxgW095rMS2SqM4k"
STAFF = "https://staffportal.uoz.edu.krd"

ARABIC = re.compile(r"[؀-ۿ]")
LATIN = re.compile(r"[A-Za-z]")


# ─────────────────────────── yardımcılar ───────────────────────────
def http_get(url, params=None, timeout=40, raw=False):
    if params:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    return data if raw else data.decode("utf-8", errors="replace")


def get_json(url, params=None, timeout=40):
    return json.loads(http_get(url, params, timeout))


def key(s):
    """Karşılaştırma için: KGeCl₂I -> kgecl2i, noktalama/boşluk yok."""
    s = unicodedata.normalize("NFKC", s or "").lower()
    return re.sub(r"[^0-9a-z؀-ۿ]+", "", s)


def words(s):
    s = unicodedata.normalize("NFKC", s or "").lower()
    return re.findall(r"[0-9a-z]+", s)


def title_sim(a, b):
    return difflib.SequenceMatcher(None, key(a), key(b)).ratio()


def clean_url(u):
    return re.sub(r"[?#].*$", "", (u or "").strip()) or None


# ─────────────────────────── 1) dilekçeyi kurallarla okuma ───────────────────────────
STOP_WORDS = ("kurdistan", "ministry", "university", "college", "department", "presidency",
              "impact factor", "authors", "author(s)", "road", "p.o.box", "@", "http", "regional government",
              "scopus", "clarivate", "doaj", "index")
SIGN_WORDS = ("داخاز", "دگەل", "دگه‌ل", "ڕێزگرتن", "ريزگرتن", "رێزگرتن", "سەروکێ", "سەرۆکێ", "سه‌رۆكێ", "سەرۆكێ",
              "مع التقدير", "رئيس", "signature")
DEPT_EN_KU = {"physics": "فيزيك", "mathematics": "ماتماتیک", "chemistry": "کیمیا", "biology": "بایۆلۆجی",
              "computer": "کۆمپیوتەر", "environmental": "زانستێن ژینگەهێ", "biomedical": "بایۆمێدیکال"}
DEPT_AR_KU = {"الفيزياء": "فيزيك", "الرياضيات": "ماتماتیک", "الكيمياء": "کیمیا", "علوم الحياة": "بایۆلۆجی",
              "الحاسوب": "کۆمپیوتەر", "البيئة": "زانستێن ژینگەهێ"}


def _is_latin_line(s):
    letters = re.findall(r"[^\W\d_]", s)
    return letters and sum(1 for c in letters if LATIN.match(c)) / len(letters) > 0.8


def _split_names(cell):
    cell = cell.strip().strip(",،;").strip()
    if not cell:
        return []
    parts = re.split(r"\s*(?:,|،|;|\band\b|&|\sو\s)\s*", cell)
    return [p.strip(" .-–") if not p.strip().endswith(".") else p.strip() for p in parts if p.strip(" .-–")]


def _looks_like_name(s):
    w = s.split()
    return 1 < len(w) <= 6 and not re.search(r"\d|@|http|:", s)


def parse_authors(lines):
    """'Authors:' satırından sonra gelen satırlar (tablo '|' ya da tab ile iki sütun)."""
    start = None
    for i, l in enumerate(lines):
        if re.search(r"\bauthors?\b|author\(s\)|ناڤێن ڤەکۆلەر|ناڤێن ڤەكۆلەر|الباحثين|اسماء الباحث", l, re.I):
            start = i
            break
    rows = []
    if start is not None:
        rest = re.split(r"(?i)authors?\s*(?:\(s\))?\s*[:：]?", lines[start], maxsplit=1)
        cand = ([rest[1]] if len(rest) > 1 and rest[1].strip() else []) + lines[start + 1:start + 15]
        for l in cand:
            if any(w in l for w in SIGN_WORDS) or re.search(r"\d{4}\s*/\s*\d{1,2}", l):
                break
            if not l.strip():
                continue
            rows.append(l)
    else:  # başlık yoksa: hem Latin hem Arap harfli isim içeren tablo satırları
        rows = [l for l in lines if "|" in l and LATIN.search(l) and ARABIC.search(l)
                and all(_looks_like_name(c.strip()) or not c.strip() for c in l.split("|"))]
    en, ku = [], []
    for r in rows:
        cells = [c for c in re.split(r"\s*\|\s*|\t+|\s{3,}", r) if c.strip()]
        for c in cells:
            names = [n for n in _split_names(c) if _looks_like_name(n)]
            if not names:
                continue
            (ku if ARABIC.search(c) else en).extend(names)
    n = max(len(en), len(ku))
    out = []
    for i in range(n):
        out.append({"name_ku": ku[i] if i < len(ku) else (en[i] if i < len(en) else ""),
                    "name_en": en[i] if i < len(en) else ""})
    return out


def parse_letter(text):
    lines = [l for l in text.split("\n") if l.strip() and not re.fullmatch(r"\[[^\]]+\]", l.strip())]
    full = "\n".join(lines)
    unclear = []

    m = re.search(r"(?:Impact\s*Factor|\bI\.?F\.?\b)\s*(?:\(\s*IF\s*\))?\s*[:=\-–]?\s*\(?\s*(\d+(?:[.,]\d+)?)", full, re.I)
    impact = float(m.group(1).replace(",", ".")) if m else None
    mq = re.search(r"Scopus[^\n]{0,40}?\bQ\s*([1-4])\b|\bQ\s*([1-4])\b[^\n]{0,20}Scopus", full, re.I)
    q = int(mq.group(1) or mq.group(2)) if mq else None
    idx_line = next((l.strip() for l in lines if re.search(r"impact factor|scopus|clarivate|doaj|thomson", l, re.I)), "")

    journal = ""
    for l in lines:
        if re.search(r"گوڤار|گۆڤار|مجلة|journal", l, re.I):
            mj = re.findall(r"\(([^()]*[A-Za-z][^()]*)\)", l)
            if mj:
                journal = mj[0].strip()
                break
            mj = re.search(r"journal\s*(?:name)?\s*[:：]\s*(.+)", l, re.I)
            if mj:
                journal = mj.group(1).strip()
                break
    if not journal:
        unclear.append("Dergi adı bulunamadı")

    cands = []
    for l in lines:
        s = l.strip().strip("()\"'“”«» ").strip()
        if not _is_latin_line(s) or len(s.split()) < 4 or "|" in s:
            continue
        low = s.lower()
        if any(w in low for w in STOP_WORDS) or key(s) == key(journal):
            continue
        cands.append(s)
    title = max(cands, key=len) if cands else ""
    if not title:
        unclear.append("Paper başlığı bulunamadı")

    authors = parse_authors(lines)
    if not authors:
        unclear.append("Yazar adları bulunamadı")
    elif any(not a["name_en"] for a in authors):
        unclear.append("Bazı yazarların İngilizce adı yok; internette eşleştirme zayıf olabilir")
    elif any(a["name_ku"] == a["name_en"] for a in authors):
        unclear.append("Bazı yazarların Kürtçe adı yok; Word'de İngilizce ad kullanıldı")

    dept = ""
    md = re.search(r"پ[شس][کك]ا[ \u200c]*[:/]?[ \u200c]*([\u0600-\u06FF\u200c]+(?:[ ]+[\u0600-\u06FF\u200c]+)?)", full)
    if md:
        dept = md.group(1).strip()
        dept = re.split(r"\s+(?:زۆر|زور|ل|li)\b", dept)[0]
    else:
        for ar, ku in DEPT_AR_KU.items():
            if re.search(r"قسم\s*" + ar, full):
                dept = ku
                break
        if not dept:
            me = re.search(r"Department\s+of\s+([A-Za-z]+)", full, re.I)
            if me and me.group(1).lower() in DEPT_EN_KU:
                dept = DEPT_EN_KU[me.group(1).lower()]
        if dept:
            unclear.append(f"Bölüm adı Kürtçeye çevrildi ({dept}); yazımını kontrol edin")
    if not dept:
        unclear.append("Bölüm adı bulunamadı")

    science = re.search(r"college\s+of\s+science|faculty\s+of\s+science|كلية\s*العلوم|ك[ۆو]ل[يێ]ژا\s*زانست|کۆلێژا\s*زانست",
                        full, re.I)
    mdate = re.search(r"(20\d\d)\s*/\s*(\d{1,2})\s*/\s*(\d{1,2})|(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(20\d\d)", full)
    return {
        "authors": authors, "title": title, "journal": journal,
        "stated_impact_factor": impact, "stated_index": idx_line, "stated_scopus_quartile": q,
        "department_ku": dept, "college_key": "Science" if science else "Other",
        "college_name": "College of Science" if science else "",
        "handwritten_amount": None, "letter_date": mdate.group(0) if mdate else "",
        "unclear": unclear,
    }


# ─────────────────────────── 2) internetten doğrulama ───────────────────────────
def _author_match(letter_en, paper_authors):
    """'Bewar M. Ahmad' ~ {'given': 'Bewar M.', 'family': 'Ahmad'}"""
    lw = words(letter_en)
    if not lw:
        return None
    for pa in paper_authors:
        fam = words(pa.get("family", ""))
        giv = words(pa.get("given", ""))
        if fam and giv and lw[-1] == fam[-1] and lw[0][:3] == giv[0][:3]:
            return pa
    return None


def crossref_find(L):
    items = get_json("https://api.crossref.org/works", {
        "query.bibliographic": f"{L['title']} {L['journal']}", "rows": 5,
        "select": "DOI,title,container-title,author,ISSN,URL,resource,published,issued"})["message"]["items"]
    best, score = None, 0.0
    for it in items:
        s = title_sim(L["title"], (it.get("title") or [""])[0])
        if s > score:
            best, score = it, s
    return (best, score) if score >= 0.85 else (None, score)


def openalex_work(doi):
    try:
        return get_json("https://api.openalex.org/works/https://doi.org/" + urllib.parse.quote(doi))
    except Exception:
        return None


def doaj_listed(issns):
    for i in issns:
        try:
            d = get_json("https://doaj.org/api/search/journals/issn:" + i)
            if d.get("total", 0) > 0:
                return True
        except Exception:
            return None
    return False if issns else None


def uoz_registry_rows():
    """UoZ yayın kayıt listesini indir (6 saat önbellek)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, "uoz_kayit_listesi.csv")
    if not os.path.exists(path) or time.time() - os.path.getmtime(path) > 6 * 3600:
        try:
            data = http_get(UOZ_SHEET_CSV, timeout=90, raw=True)
            with open(path, "wb") as f:
                f.write(data)
        except Exception:
            if not os.path.exists(path):
                raise
    with open(path, encoding="utf-8", errors="replace") as f:
        return list(csv.DictReader(f))


def _col(row, *names):
    for k, v in row.items():
        if k and any(n in k.lower() for n in names):
            return (v or "").strip()
    return ""


def uoz_registry(L, doi):
    rows = uoz_registry_rows()
    found = None
    same_journal = []
    for r in rows:
        rdoi = _col(r, "doi").lower()
        rtitle = _col(r, "publication title")
        if (doi and doi.lower() in rdoi) or (L["title"] and title_sim(L["title"], rtitle) >= 0.9):
            found = r
        rj = _col(r, "title of the journal")
        if L["journal"] and rj and title_sim(L["journal"], rj) >= 0.92:
            iff = _col(r, "impact factor")
            if iff:
                same_journal.append((_col(r, "publication date"), iff))
    reg = None
    if found:
        reg = {"timestamp": _col(found, "timestamp"), "authors": _col(found, "author"),
               "indexing": _col(found, "indexing"), "impact_factor": _col(found, "impact factor"),
               "department": _col(found, "department")}
    return reg, same_journal[-8:]


def staff_search(name_en):
    w = name_en.replace(".", " ").split()
    if not w:
        return []
    res = get_json(STAFF + "/api/searchstaff", {"lang": "en", "query": w[0]})
    lw = [x.lower() for x in w]
    out = []
    for p in res:
        fw = [x.lower() for x in p.get("fullname", "").split()]
        if not fw or fw[0] != lw[0]:
            continue
        # soyad eşleşmesi + ortadaki baş harf
        last_ok = lw[-1] == fw[-1] or (len(lw) > 2 and lw[-1] in fw)
        mid_ok = len(lw) < 3 or any(f.startswith(lw[1][:2]) for f in fw[1:])
        if last_ok and mid_ok:
            out.append(p)
    return out


def staff_has_paper(username, title):
    page = http_get(f"{STAFF}/en/journals/{urllib.parse.quote(username)}")
    page = html.unescape(re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", " ", page, flags=re.S)))
    return key(title)[:80] in key(page)


def scimago_quartile(issns):
    """veriler/ klasöründe scimagojr*.csv varsa (scimagojr.com > Download data) quartile döner."""
    if not os.path.isdir(DATA_DIR):
        return None
    files = sorted(f for f in os.listdir(DATA_DIR) if f.lower().startswith("scimagojr") and f.lower().endswith(".csv"))
    if not files:
        return None
    want = {i.replace("-", "") for i in issns}
    with open(os.path.join(DATA_DIR, files[-1]), encoding="utf-8", errors="replace") as f:
        for row in csv.DictReader(f, delimiter=";"):
            iss = {x.strip().replace("-", "") for x in (row.get("Issn") or "").split(",")}
            if iss & want:
                q = (row.get("SJR Best Quartile") or "").strip().upper()
                return {"quartile": int(q[1]) if re.fullmatch(r"Q[1-4]", q) else None,
                        "title": row.get("Title"), "file": files[-1]}
    return {"quartile": None, "title": None, "file": files[-1]}


def verify(L, log=print):
    v = {"published": False, "doi": None, "article_url": None, "published_date": None,
         "authors_on_paper": [], "authors_match": None, "uoz_affiliation": None,
         "clarivate_listed": None, "impact_factor": None, "impact_factor_year": None,
         "impact_factor_source": None, "scopus_listed": None, "scopus_quartile": None,
         "doaj_listed": None, "staff_portal": [], "notes": [], "registry": None,
         "registry_journal_ifs": [], "sjr": None, "issn": [], "links": {}}

    log("      Crossref...")
    try:
        it, score = crossref_find(L)
    except Exception as e:
        it, score = None, 0
        v["notes"].append(f"Crossref'e ulaşılamadı ({e.__class__.__name__}).")
    if it:
        v["published"] = True
        v["doi"] = it.get("DOI")
        v["article_url"] = clean_url((it.get("resource") or {}).get("primary", {}).get("URL")) or it.get("URL")
        dp = ((it.get("published") or it.get("issued") or {}).get("date-parts") or [[None]])[0]
        v["published_date"] = "-".join(str(x) for x in dp if x)
        v["issn"] = it.get("ISSN", [])
        pa = it.get("author", [])
        v["authors_on_paper"] = [f"{a.get('given', '')} {a.get('family', '')}".strip() for a in pa]
        if all(a["name_en"] for a in L["authors"]):
            miss = [a["name_en"] for a in L["authors"] if not _author_match(a["name_en"], pa)]
            v["authors_match"] = not miss
            if miss:
                v["notes"].append("Makalede bulunamayan yazar(lar): " + ", ".join(miss))
        cj = (it.get("container-title") or [""])[0]
        if cj and title_sim(cj, L["journal"]) < 0.8:
            v["notes"].append(f"Makalenin dergisi '{cj}', dilekçede '{L['journal']}' yazıyor.")
    elif score:
        v["notes"].append(f"Crossref'te bu başlıkla birebir paper bulunamadı (en yakın benzerlik %{score * 100:.0f}).")

    if v["doi"]:
        log("      OpenAlex...")
        w = openalex_work(v["doi"])
        if w:
            insts = [i.get("display_name", "") for a in w.get("authorships", []) for i in a.get("institutions", [])]
            raw = [s for a in w.get("authorships", []) for s in (a.get("raw_affiliation_strings") or [])]
            v["uoz_affiliation"] = any("zakho" in s.lower() for s in insts + raw)
            src = (w.get("primary_location") or {}).get("source") or {}
            v["issn"] = sorted(set(v["issn"]) | set(src.get("issn") or []))
            v["openalex_affiliations"] = sorted(set(raw))[:6]

    if v["issn"]:
        log("      DOAJ...")
        v["doaj_listed"] = doaj_listed(v["issn"])
        v["sjr"] = scimago_quartile(v["issn"])

    log("      UoZ yayın kayıt listesi...")
    try:
        v["registry"], v["registry_journal_ifs"] = uoz_registry(L, v["doi"])
    except Exception as e:
        v["notes"].append(f"UoZ kayıt listesi indirilemedi ({e.__class__.__name__}).")

    log("      Staff portal...")
    for a in L["authors"]:
        name = a["name_en"]
        entry = {"author": name or a["name_ku"], "url": None, "paper_listed": None}
        try:
            hits = staff_search(name) if name else []
            if len(hits) == 1:
                u = hits[0]["username"]
                entry["url"] = f"{STAFF}/en/{u}"
                entry["fullname"] = hits[0]["fullname"]
                if L["title"]:
                    entry["paper_listed"] = staff_has_paper(u, L["title"])
            elif len(hits) > 1:
                entry["candidates"] = [f"{h['fullname']}: {STAFF}/en/{h['username']}" for h in hits]
        except Exception as e:
            v["notes"].append(f"Staff portal ({name}): {e.__class__.__name__}")
        v["staff_portal"].append(entry)

    issn = (v["issn"] or [""])[0]
    q = urllib.parse.quote
    v["links"] = {
        "Clarivate MJL (IF/Web of Science kontrolü)": f"https://mjl.clarivate.com/search-results?issn={q(issn)}" if issn
        else f"https://mjl.clarivate.com/search-results?search={q(L['journal'])}",
        "Scopus kaynak araması": f"https://www.scopus.com/sources?query={q(L['journal'])}",
        "SCImago": f"https://www.scimagojr.com/journalsearch.php?q={q(issn or L['journal'])}",
        "UoZ yayın kayıt listesi": UOZ_SHEET_VIEW,
        "Staff portal ana sayfa (isimle arama)": STAFF + "/en",
    }
    return v
