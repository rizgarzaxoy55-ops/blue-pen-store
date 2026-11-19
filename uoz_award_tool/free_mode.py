# -*- coding: utf-8 -*-
"""
Free mode: works without an API key and without AI.
  - Reads the request letter with rules (title in “...”, journal under "ل گوڤارا:", DOI, IF/Scopus, داخازکەر).
  - Crossref + OpenAlex: is the paper published, DOI, article link, authors and each author's institution.
  - Only University of Zakho authors are written into the Word letter (OpenAlex affiliation or staff portal record).
  - UoZ publication registry (Google Sheet): is the paper registered with the university (rule 5-a).
  - UoZ staff portal: the author's page, and whether the paper is listed there (rule 5-b).
  - DOAJ API: is the journal in DOAJ.  SCImago (optional CSV in data/): quartile.
  - names.json: English name -> Kurdish name dictionary; the program learns as you correct names.
There is no free official source for the Clarivate Impact Factor: the value in the letter is used, with a warning.
"""
import csv
import difflib
import html
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
NAMES_FILE = os.path.join(HERE, "names.json")
OLD_NAMES_FILE = os.path.join(HERE, "isimler.json")
SETTINGS_FILES = [os.path.join(HERE, "settings.json"), os.path.join(HERE, "ayarlar.json")]
UA = "UoZ-award-tool/1.2 (publication award verification)"
UOZ_SHEET_CSV = ("https://docs.google.com/spreadsheets/d/"
                 "1q6_uje4RCXU9pn9JqIbDKxaJ2-doxgW095rMS2SqM4k/gviz/tq?tqx=out:csv")
UOZ_SHEET_VIEW = "https://docs.google.com/spreadsheets/d/1q6_uje4RCXU9pn9JqIbDKxaJ2-doxgW095rMS2SqM4k"
STAFF = "https://staffportal.uoz.edu.krd"

ARABIC = re.compile(r"[؀-ۿ]")
LATIN = re.compile(r"[A-Za-z]")
DOI_RE = re.compile(r"10\.\d{4,9}/[^\s\"<>“”'’]+")


# ─────────────────────────── helpers ───────────────────────────
def _setting(name):
    for path in SETTINGS_FILES:
        try:
            with open(path, encoding="utf-8") as f:
                val = (json.load(f).get(name) or "").strip()
            if val:
                return val
        except (FileNotFoundError, ValueError):
            pass
    return ""


def http_get(url, params=None, timeout=40, raw=False):
    if params:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        data = r.read()
    return data if raw else data.decode("utf-8", errors="replace")


def get_json(url, params=None, timeout=40):
    return json.loads(http_get(url, params, timeout))


def openalex(path, params=None):
    params = dict(params or {})
    k = _setting("openalex_api_key")
    if k:
        params["api_key"] = k
    return get_json("https://api.openalex.org/" + path, params)


def key(s):
    """For comparisons: KGeCl₂I -> kgecl2i, no punctuation/spaces."""
    s = unicodedata.normalize("NFKC", s or "").lower()
    return re.sub(r"[^0-9a-z؀-ۿ]+", "", s)


def words(s):
    s = unicodedata.normalize("NFKC", s or "").lower()
    return re.findall(r"[0-9a-z]+", s)


def title_sim(a, b):
    return difflib.SequenceMatcher(None, key(a), key(b)).ratio()


def clean_url(u):
    return re.sub(r"[?#].*$", "", (u or "").strip()) or None


def clean_doi(d):
    d = (d or "").strip().rstrip(".,;:)]")
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", d, flags=re.I).lower() or None


# ── Arabic-script name <-> Latin name similarity (consonant skeleton) ──
_AR = {"ب": "b", "پ": "b", "ت": "t", "ث": "s", "ج": "j", "چ": "j", "ح": "", "خ": "k", "د": "d", "ذ": "z",
       "ر": "r", "ڕ": "r", "ز": "z", "ژ": "j", "س": "s", "ش": "s", "ص": "s", "ض": "d", "ط": "t", "ظ": "z",
       "غ": "g", "ف": "f", "ڤ": "f", "ق": "k", "ك": "k", "ک": "k", "گ": "g", "ل": "l", "ڵ": "l", "م": "m",
       "ن": "n", "ه": "", "ة": "", "ھ": ""}


def _skel_ar(name):
    out = []
    for w in name.split():
        s = "".join(_AR.get(c, "") for c in w)
        if s:
            out.append(re.sub(r"(.)\1+", r"\1", s))
    return out


def _skel_lat(name):
    out = []
    for w in words(name):
        w = re.sub(r"kh|q|c", "k", w)
        w = re.sub(r"sh", "s", w)
        w = re.sub(r"th", "t", w)
        w = re.sub(r"dh", "d", w)
        w = re.sub(r"gh", "g", w)
        w = re.sub(r"ph|v", "f", w)
        w = re.sub(r"[aeiouyhw]", "", w)
        w = w.replace("p", "b")
        if w:
            out.append(re.sub(r"(.)\1+", r"\1", w))
    return out


def name_sim(ar_name, lat_name):
    a, b = _skel_ar(ar_name), _skel_lat(lat_name)
    if not a or not b:
        return 0.0
    first = difflib.SequenceMatcher(None, a[0], b[0]).ratio()
    rest = difflib.SequenceMatcher(None, "".join(a[1:]), "".join(b[1:])).ratio() if len(a) > 1 and len(b) > 1 else first
    return 0.5 * first + 0.5 * rest


# ── names.json (English -> Kurdish) ──
def _nkey(name_en):
    w = [x for x in words(name_en) if len(x) > 1]
    return " ".join(w)


def load_names():
    names = {}
    for path in (OLD_NAMES_FILE, NAMES_FILE):      # names.json wins over the old isimler.json
        try:
            with open(path, encoding="utf-8") as f:
                names.update(json.load(f))
        except (FileNotFoundError, ValueError):
            pass
    return names


def learn_names(authors):
    """Save the authors whose Kurdish name is known into names.json."""
    names = load_names()
    changed = False
    for a in authors:
        en, ku = (a.get("name_en") or "").strip(), (a.get("name_ku") or "").strip()
        if en and ku and ARABIC.search(ku) and names.get(_nkey(en)) != ku:
            names[_nkey(en)] = ku
            changed = True
    if changed:
        with open(NAMES_FILE, "w", encoding="utf-8") as f:
            json.dump(names, f, ensure_ascii=False, indent=1, sort_keys=True)


def lookup_name(name_en, names):
    k = _nkey(name_en)
    if k in names:
        return names[k]
    w = k.split()
    for kk, v in names.items():   # "Alaa L. Ibrahim" ~ "alaa luqman ibrahim"
        ww = kk.split()
        if w and ww and w[0] == ww[0] and w[-1] == ww[-1]:
            return v
    return None


# ─────────────────────────── 1) reading the letter with rules ───────────────────────────
STOP_WORDS = ("kurdistan", "ministry", "university", "college", "department", "presidency",
              "impact factor", "authors", "author(s)", "road", "p.o.box", "@", "http", "regional government",
              "scopus", "clarivate", "doaj", "issn", "web of science")
SIGN_WORDS = ("داخاز", "دگەل", "دگه‌ل", "ڕێزگرتن", "ريزگرتن", "رێزگرتن", "سەروکێ", "سەرۆکێ", "سه‌رۆكێ", "سەرۆكێ",
              "هەڤپێچ", "مع التقدير", "رئيس", "signature")
DEPT_EN_KU = {"physics": "فیزیک", "mathematics": "ماتماتیک", "chemistry": "کیمیا", "biology": "بایۆلۆجی",
              "computer": "کۆمپیوتەر", "environmental": "زانستێن ژینگەهێ", "biomedical": "بایۆمێدیکال"}
DEPT_AR_KU = {"الفيزياء": "فیزیک", "الرياضيات": "ماتماتیک", "الریاضیات": "ماتماتیک", "الكيمياء": "کیمیا",
              "علوم الحياة": "بایۆلۆجی", "الحاسوب": "کۆمپیوتەر", "البيئة": "زانستێن ژینگەهێ"}
JOURNAL_LABEL = re.compile(r"(?:ل\s*)?[گك][ۆو][ڤف]ار[اەه]?\s*[:：]|مجلة\s*[:：]|journal\s*(?:name|title)?\s*[:：]", re.I)
TITLE_LABEL = re.compile(r"ن[اإ]?[فڤةه]?\s*(?:و|9)\s*ن[يیئ]ش")          # ناڤ و نیشان (tolerates OCR errors)
AUTHORS_LABEL = re.compile(r"ن[اإ][فڤ][يێئی]?ن\b.{0,14}?[كک][ۆو]?ل[ه]|[كک][ۆو]?ل[هە]\s*را")       # ناڤێن ڤەکولەرا
LABEL_LINE = re.compile(r"^\s*(journal\s*(index|title|name)|publication\s*date|issn|doi)\b", re.I)


def parse_listed_authors(lines):
    """Biology layout: lines like '1- محمد ياسر نعيم (University of Padua)' -> Kurdish name + institution."""
    start = next((i for i, l in enumerate(lines) if AUTHORS_LABEL.search(l)), None)
    if start is None:
        return []
    out = []
    for l in lines[start + 1:start + 14]:
        if any(w in l for w in SIGN_WORDS) or re.search(r"دگه?ل|[پب]\.ه", l):
            break
        m = re.match(r"\s*(\d)\s*[-.]?\s*([؀-ۿ‌ ]{3,}?)\s*([()（）].*)?$", l)
        if m:
            name = re.sub(r"\s+", " ", m.group(2)).strip()
            if len(name.split()) >= 2:
                out.append({"name_ku": name, "aff": (m.group(3) or "").strip()})
                continue
        if out and not out[-1]["aff"] and re.match(r"\s*[(（]", l):
            out[-1]["aff"] = l.strip()
        elif out and not re.match(r"\s*\d", l):
            break
    for a in out:
        a["uoz"] = bool(re.search(r"زاخ|zakho", a["aff"], re.I))
    return out


KNOWN_DEPTS = ["ماتماتیک", "فیزیک", "کیمیا", "بایۆلۆجی", "زانستێن کۆمپیوتەری", "زانستێن ژینگەهێ",
               "جیۆلۆجی", "ئامار", "زانستێن ئاژەلان", "زانستێن ڕووەکان", "بایۆمێدیکال"]


def _dskel(s):
    s = s.translate(str.maketrans({"ك": "ک", "ي": "ی", "ى": "ی", "ة": "ە", "ۆ": "و", "ێ": "ی", "ڕ": "ر", "ڤ": "ف"}))
    return re.sub(r"[\s‌ً-ٟ]", "", s)


def canonical_dept(raw):
    """Clean OCR/spelling differences: 'ماتماتيك ا' / 'ماتماتیک زوربەرێز' -> 'ماتماتیک'."""
    w = [x for x in re.split(r"[\s/،,|:]+", raw.strip()) if x]
    if not w:
        return ""
    cands = [w[0]] + ([w[0] + " " + w[1]] if len(w) > 1 else [])
    best, score = None, 0.0
    for c in cands:
        for k in KNOWN_DEPTS:
            r = difflib.SequenceMatcher(None, _dskel(c), _dskel(k)).ratio()
            if r > score:
                best, score = k, r
    if score >= 0.8:
        return best
    return w[0].translate(str.maketrans({"ك": "ک", "ي": "ی"}))


def _is_latin_line(s):
    letters = re.findall(r"[^\W\d_]", s)
    return letters and sum(1 for c in letters if LATIN.match(c)) / len(letters) > 0.8


def _wordlike(s):
    """Share of the line made of real English-looking words (to filter OCR noise)."""
    toks = re.findall(r"[^\s]+", s)
    if not toks:
        return 0.0
    good = sum(1 for t in toks
               if re.fullmatch(r"[(\[“\"']?(?:[A-Z]?[a-z]+|[A-Z]{2,}|[A-Z][a-z]*-[A-Za-z]+|\d+[A-Za-z]*|[A-Z][a-z]*[₀-₉\d]*[A-Za-z]*)[)\]”\"'.,:;’s]*", t)
               and re.search(r"[aeiouyAEIOUY]|\d", t))
    return good / len(toks)


def _trim_garbage(s):
    """Remove OCR noise from the start/end of a title: digits, apostrophes, mid-word capitals, very short bits."""
    toks = s.split()
    bad = lambda t: (re.search(r"\d|['’‘`]", t) or re.search(r"[a-z][A-Z]", t)
                     or len(t.strip("()[],.:;")) <= 2 and t.lower() not in ("a", "an", "of", "to", "in", "on", "by")
                     or not re.search(r"[A-Za-z]", t))
    while toks and (bad(toks[0]) or len(toks[0]) <= 3 and toks[0] not in ("A", "An", "The", "On")):
        toks.pop(0)
    while toks and (bad(toks[-1]) or len(toks[-1]) <= 3 and toks[-1].islower()):
        toks.pop()
    return " ".join(toks)


def _split_names(cell):
    cell = cell.strip().strip(",،;").strip()
    if not cell:
        return []
    parts = re.split(r"\s*(?:,|،|;|\band\b|&|\sو\s)\s*", cell)
    return [p.strip(" -–") for p in parts if p.strip(" .-–")]


def _looks_like_name(s):
    w = s.split()
    return 1 < len(w) <= 6 and not re.search(r"\d|@|http|:", s)


def parse_authors(lines):
    """Lines after an 'Authors:' line (two columns via table '|' or tabs)."""
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
            if l.strip():
                rows.append(l)
    else:
        rows = [l for l in lines if "|" in l and LATIN.search(l) and ARABIC.search(l)
                and all(_looks_like_name(c.strip()) or not c.strip() for c in l.split("|"))]
    en, ku = [], []
    for r in rows:
        for c in [c for c in re.split(r"\s*\|\s*|\t+|\s{3,}", r) if c.strip()]:
            names = [n for n in _split_names(c) if _looks_like_name(n)]
            (ku if ARABIC.search(c) else en).extend(names)
    out = []
    for i in range(max(len(en), len(ku))):
        e = en[i] if i < len(en) else ""
        out.append({"name_ku": ku[i] if i < len(ku) else e, "name_en": e})
    return out


def _clean_journal(s):
    s = s.strip().strip("()“”\"' ").strip()
    s = re.split(r"\s+(?:Vol\.?|Volume|No\.|Issue)\b|\s+\d+\s*[,(]|\s*\(\d{4}\)|,\s*\d", s)[0]
    return s.strip(" .,;:-–")


def _ocr_fix(text):
    text = re.sub(r"[‎‏‪-‮⁦-⁩]", "", text)
    text = re.sub(r"(doi\.org/)\s*1\s?0\s?\.\s?", r"\g<1>10.", text, flags=re.I)
    text = re.sub(r"\b1\s0\.(\d{4,9})", r"10.\1", text)
    text = re.sub(r"(10\.\d{4,9})\s*/\s*", r"\1/", text)
    return text


def parse_letter(text):
    text = _ocr_fix(text)
    lines = [l for l in text.split("\n") if l.strip() and not re.fullmatch(r"\[[^\]]+\]", l.strip())]
    full = "\n".join(lines)
    unclear = []

    m = re.search(r"(?:Impact\s*Factor|\bI\.?F\.?\b)\s*(?:\(\s*IF\s*\))?\s*[:=\-–]?\s*\(?\s*(\d+(?:[.,]\d+)?)", full, re.I)
    impact = float(m.group(1).replace(",", ".")) if m else None
    mq = re.search(r"Scopus[^\n]{0,40}?\bQ\s*([1-4])\b|\bQ\s*([1-4])\b[^\n]{0,20}Scopus"
                   r"|Journal\s*Index\s*[:：]\s*Scopus\s*Q?\s*([1-4])\b", full, re.I)
    q = int(next(g for g in mq.groups() if g)) if mq else None
    idx_line = next((l.strip() for l in lines if re.search(r"impact factor|\bIF\s*=|scopus|clarivate|doaj|thomson", l, re.I)), "")
    doaj = bool(re.search(r"\bDOAJ\b", full, re.I))

    dois = []
    for d in DOI_RE.findall(full):
        d = clean_doi(d)
        if d and d not in dois:
            dois.append(d)
    links = [u.rstrip(".,;)") for u in re.findall(r"https?://\S+", full)
             if "doi.org" not in u and "uoz.edu.krd" not in u.lower() or "sjuoz" in u.lower()]

    # journal: "(Chemical Papers)" or the rest of / the line below "ل گوڤارا:"
    journal = ""
    for i, l in enumerate(lines):
        mj = JOURNAL_LABEL.search(l)
        if mj:
            after = l[mj.end():].strip()
            if after and LATIN.search(after):
                journal = _clean_journal(after)
            else:
                for nxt in lines[i + 1:i + 3]:
                    if LATIN.search(nxt) and not nxt.lower().startswith(("issn", "http")):
                        journal = _clean_journal(nxt)
                        break
            if journal:
                break
    if not journal:
        for l in lines:
            if re.search(r"گوڤار|گۆڤار|مجلة|journal", l, re.I):
                mj = re.findall(r"\(([^()]*[A-Za-z][^()]*)\)", l)
                if mj:
                    journal = _clean_journal(mj[0])
                    break

    # title: text inside “...”; else the English line(s) under "ناڤ و نیشان"; else the longest English line
    title = ""
    mt = re.search(r"[“\"«]([^”\"»]{15,})[”\"»]", full)
    if mt and LATIN.search(mt.group(1)):
        title = re.sub(r"\s+", " ", mt.group(1)).strip()
    if not title:
        for i, l in enumerate(lines):
            if TITLE_LABEL.search(l) and not _is_latin_line(l):
                parts = []
                for nxt in lines[i + 1:i + 5]:
                    if _is_latin_line(nxt) and not LABEL_LINE.search(nxt):
                        parts.append(nxt.strip())
                    elif parts:
                        break
                if parts:
                    title = re.sub(r"\s+", " ", " ".join(parts)).strip(" .")
                    break
    if not title:
        cands = []
        for i, l in enumerate(lines):
            s = l.strip().strip("()\"'“”«» ").strip()
            if not _is_latin_line(s) or len(s.split()) < 4 or "|" in s or _wordlike(s) < 0.75:
                continue
            if any(w in s.lower() for w in STOP_WORDS) or key(s) == key(journal) or LABEL_LINE.search(s):
                continue
            # a title broken over two lines: join with the next line ("... Molecular Targets to" + "Lifestyle Medicine")
            for nxt in lines[i + 1:i + 3]:
                n = nxt.strip().strip("()\"'“”«» ").strip()
                if (n and not re.search(r"[.?!]$", s) and _is_latin_line(n) and _wordlike(n) >= 0.75
                        and len(n.split()) <= 14 and not LABEL_LINE.search(n)
                        and not any(w in n.lower() for w in STOP_WORDS)):
                    s = s + " " + n
                else:
                    break
            s = _trim_garbage(s)
            if len(s.split()) >= 4:
                cands.append(s)
        title = max(cands, key=len) if cands else ""

    authors = parse_authors(lines)
    listed = parse_listed_authors(lines)
    if listed:
        authors = []

    applicant = ""
    for i, l in enumerate(lines):
        if "داخاز" in l:
            after = re.sub(r".*داخاز[^\s:]*\s*[:：]?", "", l).strip()
            cand = [after] if ARABIC.search(after) else lines[i + 1:i + 3]
            for c in cand:
                c = re.sub(r"^(د|م|پ\.?ی|پ\.?هـ?\.?د)\.\s*", "", c.strip())
                if ARABIC.search(c) and 1 < len(c.split()) <= 5 and not re.search(r"\d", c):
                    applicant = c
                    break
            break

    dept = ""
    md = re.search(r"[پ][شس][کك]ا[ ‌]*[:/]?[ ‌]*([؀-ۿ‌]+(?:[ ]+[؀-ۿ‌]+)?)", full)
    if md:
        dept = canonical_dept(md.group(1))
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
            unclear.append(f"Department name was translated into Kurdish ({dept}); check the spelling")
    if not dept:
        unclear.append("Department name not found")
    if not title:
        unclear.append("Paper title not found")
    if not journal:
        unclear.append("Journal name not found in the letter")

    science = re.search(r"college\s+of\s+science|faculty\s+of\s+science|كلية\s*العلوم|ك[ۆو]ل[يیێ]ژا\s*زانست|کۆلێژا\s*زانست",
                        full, re.I)
    return {
        "authors": authors, "listed_authors": listed, "applicant_ku": applicant, "title": title, "journal": journal,
        "stated_impact_factor": impact, "stated_index": idx_line, "stated_scopus_quartile": q,
        "stated_doaj": doaj, "stated_dois": dois, "stated_links": links,
        "department_ku": dept, "college_key": "Science" if science else "Other",
        "college_name": "College of Science" if science else "",
        "handwritten_amount": None, "letter_date": "", "unclear": unclear,
    }


# ─────────────────────────── 2) online verification ───────────────────────────
def _author_match(letter_en, paper_authors):
    lw = words(letter_en)
    if not lw:
        return None
    for pa in paper_authors:
        fam, giv = words(pa.get("family", "")), words(pa.get("given", ""))
        if fam and giv and lw[-1] == fam[-1] and lw[0][:3] == giv[0][:3]:
            return pa
    return None


def crossref_doi(doi):
    try:
        return get_json("https://api.crossref.org/works/" + urllib.parse.quote(doi, safe="/"))["message"]
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def crossref_find(title, journal):
    items = get_json("https://api.crossref.org/works", {
        "query.bibliographic": f"{title} {journal}", "rows": 5,
        "select": "DOI,title,container-title,author,ISSN,URL,resource,published,issued"})["message"]["items"]
    best, score = None, 0.0
    for it in items:
        s = title_sim(title, (it.get("title") or [""])[0])
        if s > score:
            best, score = it, s
    return (best, score) if score >= 0.85 else (None, score)


def openalex_by_doi(doi):
    try:
        return openalex("works/https://doi.org/" + urllib.parse.quote(doi, safe="/"))
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        raise


def openalex_by_title(title):
    r = openalex("works", {"search": title, "per-page": 5})
    best, score = None, 0.0
    for w in r.get("results", []):
        s = title_sim(title, w.get("title") or "")
        if s > score:
            best, score = w, s
    return best if score >= 0.85 else None


def doaj_listed(issns):
    for i in issns:
        try:
            if get_json("https://doaj.org/api/search/journals/issn:" + i).get("total", 0) > 0:
                return True
        except Exception:
            return None
    return False if issns else None


def uoz_registry_rows():
    """Download the UoZ publication registry (cached for 6 hours)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    path = os.path.join(DATA_DIR, "uoz_registry.csv")
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


def uoz_registry(title, journal, doi):
    rows = uoz_registry_rows()
    found, same_journal = None, []
    for r in rows:
        rdoi = _col(r, "doi").lower()
        if (doi and doi.lower() in rdoi) or (title and title_sim(title, _col(r, "publication title")) >= 0.9):
            found = r
        rj = _col(r, "title of the journal")
        if journal and rj and title_sim(journal, rj) >= 0.92 and _col(r, "impact factor"):
            same_journal.append((_col(r, "publication date"), _col(r, "impact factor")))
    reg = None
    if found:
        reg = {"timestamp": _col(found, "timestamp"), "authors": _col(found, "author"),
               "indexing": _col(found, "indexing"), "impact_factor": _col(found, "impact factor"),
               "department": _col(found, "department")}
    return reg, same_journal[-8:]


def staff_search(name_en, loose=False):
    """Exact matches; with loose=True returns (exact, possible)."""
    w = re.sub(r"[^A-Za-z\s'-]", " ", name_en).split()
    if not w:
        return ([], []) if loose else []
    res = get_json(STAFF + "/api/searchstaff", {"lang": "en", "query": w[0]})
    lw = [x.lower() for x in w]
    out, maybe = [], []
    for p in res:
        fw = [x.lower() for x in p.get("fullname", "").split()]
        if not fw or fw[0] != lw[0]:
            continue
        last_ok = lw[-1] == fw[-1] or (len(lw) > 2 and lw[-1] in fw)
        mid_ok = len(lw) < 3 or len(fw) < 3 or any(f.startswith(lw[1][:2]) for f in fw[1:])
        if last_ok and mid_ok:
            out.append(p)
        elif set(lw[1:]) & set(fw[1:]):
            maybe.append(p)
    return (out, maybe) if loose else out


def staff_has_paper(username, title):
    page = http_get(f"{STAFF}/en/journals/{urllib.parse.quote(username)}")
    page = html.unescape(re.sub(r"<[^>]+>", " ", re.sub(r"<script.*?</script>", " ", page, flags=re.S)))
    return key(title)[:80] in key(page)


def scimago_quartile(issns):
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


def _err(e):
    if isinstance(e, urllib.error.HTTPError) and e.code == 429:
        return "too many requests (free daily limit reached)"
    return e.__class__.__name__


def verify(L, log=print):
    v = {"published": False, "doi": None, "article_url": None, "published_date": None, "journal": None,
         "authors_on_paper": [], "authors_match": None, "uoz_affiliation": None, "uoz_authors": [],
         "clarivate_listed": None, "impact_factor": None, "impact_factor_year": None,
         "impact_factor_source": None, "scopus_listed": None, "scopus_quartile": None,
         "doaj_listed": None, "staff_portal": [], "notes": [], "registry": None,
         "registry_journal_ifs": [], "sjr": None, "issn": [], "links": {}}

    # ── Crossref: first the DOI from the letter, then the title
    log("      Crossref...")
    it = None
    try:
        for d in L.get("stated_dois", []):
            it = crossref_doi(d)
            if it:
                break
        if not it and L.get("stated_dois"):
            v["notes"].append(("The DOI read by OCR was not found in Crossref (it may be an OCR error): "
                               if L.get("source") == "ocr" else "The DOI in the letter is not registered in Crossref: ")
                              + ", ".join(L["stated_dois"]))
        if not it and L["title"]:
            it, score = crossref_find(L["title"], L["journal"])
            if it and L.get("stated_dois") and clean_doi(it.get("DOI")) not in L["stated_dois"]:
                v["notes"].append(f"The paper's real DOI is {it.get('DOI')}; the letter has a different one "
                                  f"({', '.join(L['stated_dois'])}).")
    except Exception as e:
        v["notes"].append(f"Could not reach Crossref ({_err(e)}).")
    pa = []
    if it:
        v["published"] = True
        v["doi"] = clean_doi(it.get("DOI"))
        v["article_url"] = clean_url((it.get("resource") or {}).get("primary", {}).get("URL")) or it.get("URL")
        dp = ((it.get("published") or it.get("issued") or {}).get("date-parts") or [[None]])[0]
        v["published_date"] = "-".join(str(x) for x in dp if x)
        v["issn"] = it.get("ISSN", [])
        v["journal"] = html.unescape((it.get("container-title") or [""])[0]) or None
        pa = [{"given": a.get("given", ""), "family": a.get("family", "")} for a in it.get("author", [])]
        real_title = html.unescape((it.get("title") or [""])[0])
        if L.get("source") == "ocr" and real_title:
            L["title"] = real_title          # use the paper's clean title instead of the OCR one
            if v["journal"]:
                L["journal"] = v["journal"]
        elif L["title"] and title_sim(L["title"], real_title) < 0.85:
            v["notes"].append("The title in the letter differs from the paper's title: " + (it.get("title") or [""])[0])

    # ── OpenAlex: the authors' institutions (+ the paper itself if Crossref did not have it)
    log("      OpenAlex...")
    w = None
    try:
        doi = v["doi"] or next(iter(L.get("stated_dois", [])), None)
        w = openalex_by_doi(doi) if doi else None
        if not w and not it and L["title"]:
            w = openalex_by_title(L["title"])
    except Exception as e:
        v["notes"].append(f"Could not reach OpenAlex ({_err(e)}); the authors' university was estimated from the staff portal.")
    oa_auth = []
    if w:
        for a in w.get("authorships", []):
            insts = [i.get("display_name", "") for i in a.get("institutions", [])]
            raw = a.get("raw_affiliation_strings") or []
            oa_auth.append({"name": a.get("author", {}).get("display_name", ""),
                            "uoz": any("zakho" in s.lower() for s in insts + raw), "aff": raw or insts})
        v["uoz_affiliation"] = any(a["uoz"] for a in oa_auth)
        v["openalex_affiliations"] = sorted({s for a in oa_auth for s in a["aff"]})[:6]
        src = (w.get("primary_location") or {}).get("source") or {}
        v["issn"] = sorted(set(v["issn"]) | set(src.get("issn") or []))
        if not it:
            v["published"] = True
            v["doi"] = clean_doi(w.get("doi"))
            v["article_url"] = (w.get("primary_location") or {}).get("landing_page_url")
            v["published_date"] = w.get("publication_date")
            v["journal"] = src.get("display_name")
        if not pa:
            pa = []
            for a in oa_auth:
                n = a["name"].split()
                pa.append({"given": " ".join(n[:-1]), "family": n[-1] if n else ""})
    if not v["published"]:
        ll = L.get("stated_links", [])[:2] + [f"https://doi.org/{d}" for d in L.get("stated_dois", [])[:1]]
        v["notes"].append("The paper was not found in Crossref/OpenAlex (the journal may not register DOIs). "
                          + ("Check it by hand with the link in the letter: " + ", ".join(ll) if ll
                             else "The letter has no link/DOI; search with the Google Scholar link in the report."))
    v["authors_on_paper"] = [f"{a['given']} {a['family']}".strip() for a in pa]
    if not v["article_url"] and L.get("stated_links"):
        v["letter_link"] = L["stated_links"][0]

    if L.get("journal"):
        L["unclear"] = [u for u in L.get("unclear", []) if not u.startswith("Journal name")]
    # ── journal: use the paper's full name if the letter has none or an abbreviation (SJUOZ)
    if v["journal"] and L["journal"] and len(L["journal"].split()) == 1 and title_sim(v["journal"], L["journal"]) < 0.8:
        L["journal"] = v["journal"]
    elif v["journal"] and not L["journal"]:
        L["journal"] = v["journal"]
        L["unclear"] = [u for u in L.get("unclear", []) if not u.startswith("Journal name")]
    elif v["journal"] and L["journal"] and title_sim(v["journal"], L["journal"]) < 0.8:
        v["notes"].append(f"The paper's journal is '{v['journal']}', the letter says '{L['journal']}'.")

    # ── authors
    staff_hits = {}

    def staff_lookup(name, maybe=False):
        if name not in staff_hits:
            try:
                staff_hits[name] = staff_search(name, loose=True)
            except Exception as e:
                staff_hits[name] = ([], [])
                v["notes"].append(f"Staff portal search ({name}): {_err(e)}")
        return staff_hits[name][1] if maybe else staff_hits[name][0]

    if L["authors"]:   # the letter has an author list: compare it with the paper
        if pa and all(a["name_en"] for a in L["authors"]):
            miss = [a["name_en"] for a in L["authors"] if not _author_match(a["name_en"], pa)]
            v["authors_match"] = not miss
            if miss:
                v["notes"].append("Author(s) not found on the paper: " + ", ".join(miss))
    elif pa:           # no authors in the letter: the University of Zakho authors of the paper
        log("      finding the University of Zakho authors...")
        names = load_names()
        listed = L.get("listed_authors") or []
        chosen = []
        for i, a in enumerate(pa):
            full = f"{a['given']} {a['family']}".strip()
            lm = max(listed, key=lambda x: name_sim(x["name_ku"], full), default=None)
            if lm and name_sim(lm["name_ku"], full) < 0.6:
                lm = None
            if oa_auth and i < len(oa_auth):
                is_uoz = oa_auth[i]["uoz"]
            elif lm and lm.get("aff"):
                is_uoz = lm["uoz"]
            else:
                is_uoz = bool(staff_lookup(full))
            if is_uoz:
                ku = (lm["name_ku"] if lm else None) or lookup_name(full, names) or ""
                chosen.append({"name_en": full, "name_ku": ku})
        if not chosen:
            v["notes"].append("None of the paper's authors appears to be from the University of Zakho! "
                              "Authors: " + ", ".join(v["authors_on_paper"]))
            chosen = [{"name_en": n, "name_ku": lookup_name(n, names) or ""} for n in v["authors_on_paper"]]
        others = [n for n in v["authors_on_paper"] if n not in [c["name_en"] for c in chosen]]
        if others:
            v["notes"].append("Not written into the Word letter because they are not from Zakho: " + ", ".join(others)
                              + ". If that is wrong, add them to data.json.")
        app = L.get("applicant_ku")
        if app:
            best = max(chosen, key=lambda c: name_sim(app, c["name_en"]))
            if name_sim(app, best["name_en"]) >= 0.6:
                best["name_ku"] = app
            else:
                v["notes"].append(f"The applicant ({app}) could not be matched to the paper's Zakho authors.")
        for c in chosen:
            if not c["name_ku"]:
                c["name_ku"] = c["name_en"]
        L["authors"] = chosen
        L["authors_from_paper"] = True
        v["authors_match"] = True
    elif L.get("listed_authors"):   # paper not found: the Zakho authors listed in the letter
        L["authors"] = [{"name_ku": a["name_ku"], "name_en": ""} for a in L["listed_authors"] if a["uoz"]] \
            or [{"name_ku": a["name_ku"], "name_en": ""} for a in L["listed_authors"]]
        L["authors_from_letter_list"] = True
    v["uoz_authors"] = [a["name_en"] or a["name_ku"] for a in L["authors"]]

    if v["issn"]:
        log("      DOAJ...")
        v["doaj_listed"] = doaj_listed(v["issn"])
        v["sjr"] = scimago_quartile(v["issn"])

    log("      UoZ publication registry...")
    try:
        v["registry"], v["registry_journal_ifs"] = uoz_registry(L["title"], L["journal"], v["doi"])
    except Exception as e:
        v["notes"].append(f"UoZ registry could not be downloaded ({_err(e)}).")

    # ── paper not found and no authors in the letter: take the authors from the UoZ registry
    reg = v.get("registry")
    if not L["authors"] and reg and reg.get("authors"):
        log("      taking the authors from the UoZ registry...")
        names = load_names()
        raw = [re.sub(r"[*†‡¹²³⁴⁵\d]+", "", n).strip(" .")
               for n in re.split(r"\s*(?:,|،|;|\band\b|&)\s*", reg["authors"])]
        raw = [n for n in raw if n and len(n.split()) >= 2]
        chosen = [{"name_en": n, "name_ku": lookup_name(n, names) or ""} for n in raw if staff_lookup(n)]
        if not chosen:
            chosen = [{"name_en": n, "name_ku": lookup_name(n, names) or ""} for n in raw]
            v["notes"].append("None of the registry's authors was found on the staff portal; all were written, check them.")
        app = L.get("applicant_ku")
        if app and chosen:
            best = max(chosen, key=lambda c: name_sim(app, c["name_en"]))
            if name_sim(app, best["name_en"]) >= 0.6:
                best["name_ku"] = app
        for c in chosen:
            c["name_ku"] = c["name_ku"] or c["name_en"]
        L["authors"] = chosen
        L["authors_from_registry"] = True
        v["notes"].append("The paper was not found online, so the authors were taken from the UoZ registry "
                          f"(registry: {reg['authors']}). Those found on the staff portal were written into the Word letter.")
        v["uoz_authors"] = [c["name_en"] for c in chosen]

    log("      UoZ staff portal...")
    for a in L["authors"]:
        name = a["name_en"]
        entry = {"author": name or a["name_ku"], "url": None, "paper_listed": None}
        if not name:
            entry["no_english_name"] = True
            v["staff_portal"].append(entry)
            continue
        hits = staff_lookup(name)
        if hits and len(hits) == 1:
            u = hits[0]["username"]
            entry["url"] = f"{STAFF}/en/{u}"
            entry["journals_url"] = f"{STAFF}/en/journals/{urllib.parse.quote(u)}"
            entry["fullname"] = hits[0]["fullname"]
            try:
                if L["title"]:
                    entry["paper_listed"] = staff_has_paper(u, L["title"])
                    if entry["paper_listed"]:
                        # text fragment link: Chrome/Edge scroll to and highlight the title
                        entry["paper_url"] = entry["journals_url"] + "#:~:text=" + urllib.parse.quote(
                            " ".join(L["title"].split()[:6]), safe="").replace("-", "%2D")
            except Exception as e:
                v["notes"].append(f"Staff portal page ({name}): {_err(e)}")
        elif hits:
            entry["candidates"] = [f"{h['fullname']}: {STAFF}/en/{h['username']}" for h in hits]
        elif name and staff_lookup(name, maybe=True):
            entry["candidates"] = [f"{h['fullname']}: {STAFF}/en/{h['username']}"
                                   for h in staff_lookup(name, maybe=True)]
        v["staff_portal"].append(entry)

    issn = (v["issn"] or [""])[0]
    q = urllib.parse.quote
    jr = L["journal"] or ""
    v["links"] = {
        "Clarivate MJL (IF / Web of Science check)": f"https://mjl.clarivate.com/search-results?issn={q(issn)}" if issn
        else f"https://mjl.clarivate.com/search-results?search={q(jr)}",
        "Scopus source search": f"https://www.scopus.com/sources?query={q(jr)}",
        "SCImago": f"https://www.scimagojr.com/journalsearch.php?q={q(issn or jr)}",
        "Google Scholar (paper search)": "https://scholar.google.com/scholar?q=" + q('"' + (L["title"] or "") + '"'),
        "UoZ publication registry": UOZ_SHEET_VIEW,
        "Staff portal home (search by name)": STAFF + "/en",
    }
    for i, u in enumerate(L.get("stated_links", [])[:2]):
        v["links"][f"Link in the letter {i + 1}"] = u
    return v
