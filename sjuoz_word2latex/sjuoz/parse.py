# -*- coding: utf-8 -*-
"""Kapak mektubu ve makale yapısının ayrıştırılması."""
import copy
import re
import unicodedata


from .common import EMAIL_RE, render_plain, warn
from .docx import Para, Table
from .render import render_tokens

# --------------------------------------------------------------------------
# Kapak mektubu
# --------------------------------------------------------------------------
PLACEHOLDER_RE = re.compile(r"click or tap here to enter text\.?|click here to enter text\.?", re.I)
TITLE_PREFIX_RE = re.compile(
    r"^(?:(?:asst|assist|assoc|associate|assistant|prof|professor|dr|lect|lecturer|mr|mrs|ms|miss|eng|engr|mrs)\.?\s+)+",
    re.I)


def strip_titles(name):
    return TITLE_PREFIX_RE.sub("", name.strip()).strip()


COVER_TITLE_END = re.compile(r"(for (its )?(consideration|possible|publication)|to be considered|this manuscript|"
                             r"in the science journal|submitted to)", re.I)


def parse_cover_letter(doc):
    info = {"authors": [], "title": None, "date": None, "editor": None}
    lines = []
    for b in doc.body_blocks():
        if isinstance(b, Para):
            for ln in b.text().split("\n"):
                lines.append(PLACEHOLDER_RE.sub("", ln).strip())
    for i, ln in enumerate(lines):
        if re.match(r"submission date\s*:", ln, re.I):
            info["date"] = ln.split(":", 1)[1].strip() or None
        if re.match(r"dear\b", ln, re.I):
            info["editor"] = re.sub(r"^dear\s*|\s*:\s*$", "", ln, flags=re.I).strip() or None
        if re.search(r"manuscript entitled", ln, re.I):
            # başlık birkaç satıra bölünmüş olabilir (PDF'ten okunan mektup): boş satıra ya da
            # "for consideration …" satırına kadar birleştirilir
            parts = []
            for nxt in lines[i + 1:]:
                if not nxt:
                    if parts:
                        break
                    continue
                if parts and (COVER_TITLE_END.match(nxt) or len(parts) >= 4):
                    break
                parts.append(nxt)
            if parts:
                info["title"] = " ".join(parts).strip()
    current = None
    role = None
    for ln in lines:
        if re.fullmatch(r"(corresponding author|co-?author|author)\s*:?", ln, re.I):
            role = ln.lower()
            continue
        mt = re.match(r"(name|contribution|orcid|e-?mail|affiliation|tel|phone|mobile)\s*:\s*(.*)$", ln, re.I)
        if not mt:
            continue
        key = mt.group(1).lower().replace("-", "")
        key = {"phone": "tel", "mobile": "tel"}.get(key, key)
        val = mt.group(2).strip()
        if key == "name":
            current = {"name": strip_titles(val), "raw_name": val,
                       "corresponding": bool(role and role.startswith("corresponding"))}
            info["authors"].append(current)
        elif current is not None:
            current[key] = val
    info["authors"] = [a for a in info["authors"] if a.get("name")]
    if not info["title"]:
        # serbest mektup: … article titled "A Lightweight …"
        joined = " ".join(lines)
        mt = re.search(r"\b(?:en)?titled\s*:?\s*[\"“](.+?)[\"”]", joined, re.I)
        if mt:
            info["title"] = mt.group(1).strip()
    if not info["authors"]:
        info["authors"] = _signature_authors(lines)
    for a in info["authors"]:
        orcid = a.get("orcid", "")
        mt = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", orcid)
        a["orcid"] = "https://orcid.org/" + mt.group(1) if mt else None
    return info


def _signature_authors(lines):
    """Alanları olmayan mektup: "Sincerely," altındaki imza grupları (ad, kurum satırları, Email: …)."""
    start = next((k for k, ln in enumerate(lines) if re.match(r"(sincerely|yours|best regards|regards)\b", ln, re.I)),
                 None)
    if start is None:
        return []
    authors, group = [], []
    for ln in lines[start + 1:]:
        if not ln:
            continue
        mt = re.match(r"e-?mail\s*:\s*(\S+@\S+)", ln, re.I)
        if mt:
            if group and len(group[0].split()) <= 6 and not re.search(r"\d|@|universit|department|college",
                                                                        group[0], re.I):
                authors.append({"name": strip_titles(group[0]), "raw_name": group[0], "email": mt.group(1),
                                "affiliation": ", ".join(group[1:]), "corresponding": not authors})
            group = []
        elif re.match(r"orcid\s*:", ln, re.I) and authors:
            authors[-1]["orcid"] = ln.split(":", 1)[1].strip()
        else:
            group.append(ln)
    return authors


def name_key(name):
    n = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().lower()
    parts = re.findall(r"[a-z]+", n)
    return parts


def fuller_name(ms_name, cl_name):
    """Makalede baş harfle kısaltılmış ad ("Ban J. Khoshaba"), kapak mektubunda tam yazılmışsa
    ("Ban Jibrael Khoshaba") tam ad; değilse makaledeki ad."""
    ms, cl = ms_name.split(), strip_titles(cl_name).split()
    if len(ms) != len(cl) or not any(re.fullmatch(r"[A-Z]\.?", w) for w in ms):
        return ms_name
    out = []
    for m, c in zip(ms, cl):
        if re.fullmatch(r"[A-Z]\.?", m):
            if not c[:1].upper() == m[0] or len(c.rstrip(".")) < 2:
                return ms_name
            out.append(c[0].upper() + c[1:])
        elif name_key(m) == name_key(c):
            out.append(m)
        else:
            return ms_name
    return " ".join(out)


def match_author(ms_name, cl_authors):
    mk = name_key(ms_name)
    if not mk:
        return None
    best, best_score = None, 0
    for a in cl_authors:
        ck = name_key(a["name"])
        if not ck:
            continue
        score = 0
        if mk[-1] == ck[-1]:
            score += 2
        if mk[0] == ck[0]:
            score += 1
        score += len(set(mk) & set(ck)) * 0.1
        if score > best_score:
            best, best_score = a, score
    return best if best_score >= 2 else None


# --------------------------------------------------------------------------
# Makale ayrıştırma
# --------------------------------------------------------------------------
BACK_MATTER = re.compile(
    r"^(acknowledg(e)?ments?|ethical (statement|approval|considerations?)|ethics statement|author'?s?'? contributions?|"
    r"contributions? of (the )?authors?|conflicts? of interests?|competing interests?|declaration of competing interest|"
    r"funding( statement| information)?|data availability( statement)?|availability of data.*|informed consent.*|"
    r"declarations?|abbreviations|appendix.*|supplementary.*)$", re.I)
REF_HEAD = re.compile(r"^(references?|bibliography|literature cited|kaynaklar)$", re.I)
ABSTRACT_HEAD = re.compile(r"^\s*a\w{0,3}stract\b\s*[:.\-–]?\s*", re.I)
KEYWORDS_HEAD = re.compile(r"^\s*(key\s*-?\s*words?|index terms)\s*[:.\-–—]?\s*", re.I)
CAPTION_RE = re.compile(r"^\s*(figure|fig\.?|table)\s*\(?\s*(\d+)\s*\)?\s*[:.\-–—]?\s*(.*)$", re.I | re.S)
# Harfle ayrılmış parça şekiller: "Figure 1-A: …", "Figure 1-B: …" (ayrı resimler, ortak numara)
SUBFIG_RE = re.compile(r"^\s*(?:figure|fig\.?)\s*\(?\s*(\d+)\s*[-–]\s*([A-Za-z])\s*\)?\s*[:.]", re.I)
ARTICLE_TYPE_RE = re.compile(r"^(original|review|research|short|case)\s+(article|paper|communication|report)$", re.I)


# Elle yazılmış bölüm numarası: "2.1 ", "1.INTRODUCTION", "6-REFERENCES", "3) Results"
# ("3D Images" gibi sayıyla başlayan sözcükler numara sayılmaz: ayraç ya da boşluk gerekir)
MANUAL_NUMBER = re.compile(r"^\s*(\d+(?:\.\d+)*)(?:\s*[.\-–)]\s*|\s+)(?=[A-Za-z])")


def clean_heading(txt):
    t = re.sub(r"\s+", " ", txt).strip()
    t = MANUAL_NUMBER.sub("", t)  # elle yazılmış numarayı kaldır
    return t.rstrip(":").strip()


def manual_number_level(txt):
    mt = MANUAL_NUMBER.match(txt)
    return len(mt.group(1).split(".")) if mt else None


class Article:
    def __init__(self):
        self.article_type = "Original Article"
        self.title = ""
        self.authors = []          # dict: name, marks(list), orcid, email
        self.affiliations = []     # list of str (sırayla)
        self.corr_email = None
        self.corr_tel = None
        self.abstract = []         # LaTeX paragrafları
        self.abstract_plain = []   # düz metin (Crossref/JATS için)
        self.keywords = ""
        self.title_tokens = []     # başlığın biçimli parçaları (italik tür adları için)
        self.keyword_tokens = []
        self.body = []             # body blokları (Para/Table)
        self.references = []       # Para listesi


def parse_front(blocks):
    art = Article()
    paras = blocks
    i = 0
    n = len(paras)

    def nxt_nonempty(k):
        while k < n and (not isinstance(paras[k], Para) or paras[k].is_empty()):
            k += 1
        return k

    # Makale türü + başlık
    i = nxt_nonempty(i)
    if i < n and ARTICLE_TYPE_RE.match(paras[i].text().strip()):
        art.article_type = re.sub(r"\s+", " ", paras[i].text().strip()).title()
        i = nxt_nonempty(i + 1)
    # Başlık ve yazarlar aynı paragrafta satır sonuyla ayrılmış olabilir
    if i < n and isinstance(paras[i], Para):
        head, tail = _split_at_break(paras[i])
        if tail is not None and _looks_like_authors(tail):
            paras[i:i + 1] = [head, tail]   # yerinde: çağıran blocks[i:] ile devam eder
            n += 1
    # Başlık: ilk paragraf. Uzun başlık birkaç satıra bölünmüşse (hepsi BÜYÜK harf ve kalın)
    # sonraki paragraflar da eklenir.
    title_parts = []
    if i < n and isinstance(paras[i], Para):
        title_parts.append(paras[i])
        i += 1
        while i < n and isinstance(paras[i], Para) and not paras[i].is_empty():
            t = paras[i].text().strip()
            if (_looks_like_authors(paras[i]) or not paras[i].all_bold() or t.upper() != t
                    or re.search(r"\d|@", t) or ABSTRACT_HEAD.match(t)):
                break
            title_parts.append(paras[i])
            i += 1
    art.title = " ".join(re.sub(r"\s+", " ", p.text()).strip() for p in title_parts)
    art.title_tokens = []
    for p in title_parts:
        if art.title_tokens:
            art.title_tokens.append(("t", " ", (False, False, None, None)))
        art.title_tokens += [t for t in p.tokens if t[0] == "t"]
    i = nxt_nonempty(i)

    # Yazarlar
    if i < n and isinstance(paras[i], Para):
        art.authors = parse_authors(paras[i])
        for a in art.authors:
            a["name"] = re.sub(r"\s+", " ", a["name"]).strip()
        i += 1

    # Affiliasyonlar, sorumlu yazar, özet, anahtar kelimeler
    abstract_mode = False
    while i < n:
        b = paras[i]
        if isinstance(b, Table):
            break
        txt = b.text().strip()
        if b.is_empty():
            i += 1
            continue
        if KEYWORDS_HEAD.match(txt):
            kw = KEYWORDS_HEAD.sub("", txt).strip()
            art.keywords = kw
            art.keyword_tokens = drop_chars([t for t in b.tokens if t[0] == "t"],
                                            len(b.text()) - len(b.text().lstrip()) + len(KEYWORDS_HEAD.match(txt).group(0)))
            i += 1
            break
        if abstract_mode:
            if b.heading_level() or (b.num_id and b.all_bold() and len(txt) < 150):
                break  # anahtar kelime yoksa ilk başlıkta dur
            art.abstract.append(render_tokens(b.tokens))
            art.abstract_plain.append(b.text())
            i += 1
            continue
        if ABSTRACT_HEAD.match(txt) and len(ABSTRACT_HEAD.sub("", txt)) < 5:
            abstract_mode = True
            i += 1
            continue
        if ABSTRACT_HEAD.match(txt):
            abstract_mode = True
            rest = ABSTRACT_HEAD.sub("", txt)
            art.abstract.append(render_plain(rest))
            art.abstract_plain.append(rest)
            i += 1
            continue
        if re.search(r"correspond", txt, re.I) or txt.startswith("*"):
            em = EMAIL_RE.search(txt)
            if em:
                art.corr_email = em.group(0)
            tel = re.search(r"(?:tel|phone|mobile)\.?\s*:?\s*(\+?[\d][\d\s\-()]{6,}\d)", txt, re.I)
            if tel:
                art.corr_tel = re.sub(r"\s+", " ", tel.group(1)).strip()
            i += 1
            continue
        # affiliasyon: üst simge numara ile ya da rakamla başlıyorsa
        first = next((t for t in b.tokens if t[0] == "t" and t[1].strip()), None)
        mt = re.match(r"^\s*(\d+|[a-z])[\s.)]*(.+)$", txt)
        if first is not None and (first[2][2] == "superscript" or mt):
            body_txt = mt.group(2) if mt else txt
            if first[2][2] == "superscript":
                sup = first[1].strip()
                body_txt = txt[len(txt) - len(txt.lstrip()):].lstrip()[len(sup):].strip()
            art.affiliations.append(re.sub(r"\s+", " ", body_txt).strip())
            i += 1
            continue
        if len(art.affiliations) == 0 and not art.abstract:
            # numarasız tek affiliasyon
            art.affiliations.append(re.sub(r"\s+", " ", txt))
            i += 1
            continue
        i += 1
    return art, i


def drop_chars(tokens, n):
    """Metin token'larının baştaki n karakterini atar."""
    out = []
    for t in tokens:
        if n > 0:
            cut = min(n, len(t[1]))
            n -= cut
            t = (t[0], t[1][cut:], t[2])
        if t[1]:
            out.append(t)
    return out


def _split_at_break(p):
    """Paragrafı ilk satır sonundan (boş olmayan iki parça olarak) böler; bölünemezse (p, None)."""
    toks = p.tokens
    for k, t in enumerate(toks):
        if t[0] == "t" and "\n" in t[1]:
            before = toks[:k] + [("t", t[1].split("\n", 1)[0], t[2])]
            after = [("t", t[1].split("\n", 1)[1].lstrip("\n"), t[2])] + toks[k + 1:]
            if any(x[0] == "t" and x[1].strip() for x in before) and any(x[0] == "t" and x[1].strip() for x in after):
                a, b = copy.copy(p), copy.copy(p)
                a.tokens, b.tokens = before, after
                return a, b
    return p, None


def _looks_like_authors(p):
    sup = [t for t in p.tokens if t[0] == "t" and t[2][2] == "superscript" and t[1].strip()]
    txt = p.text()
    return bool(sup) and ("," in txt or " and " in txt or len(sup) >= 1)


MARKS_ONLY = re.compile(r"^[\s,\d*†‡§]+$")


def parse_authors(p):
    s = ""
    has_sup = False
    for t in p.tokens:
        if t[0] != "t":
            continue
        if t[2][2] == "superscript":
            s += "{" + t[1] + "}"
            has_sup = True
        else:
            s += t[1]
    s = re.sub(r"\}\s*\{", "", s)  # bitişik üst simgeleri birleştir
    authors = []
    if has_sup:
        for mt in re.finditer(r"([^{}]*?)\{([^}]*)\}", s):
            name = re.sub(r"^\s*(,|;|&|\band\b)\s*", "", mt.group(1).strip(" ,;")).strip(" ,;")
            name = re.sub(r"^and\s+", "", name, flags=re.I)
            marks = [x for x in re.split(r"[,\s]+", mt.group(2)) if x]
            if MARKS_ONLY.match(name) and authors:  # üst simge yapılmamış "*" önceki yazara ait
                authors[-1]["marks"] += [x for x in re.split(r"[,\s]+", name) if x] + marks
            elif name:
                authors.append({"name": name, "marks": marks})
        tail = s[s.rfind("}") + 1:].strip(" ,;")
        tail = re.sub(r"^and\s+", "", tail, flags=re.I).strip()
        if tail and MARKS_ONLY.match(tail) and authors:
            authors[-1]["marks"] += [x for x in re.split(r"[,\s]+", tail) if x]
        elif tail:
            authors.append({"name": tail, "marks": []})
    else:
        for mt in re.finditer(r"([A-Z][^\d,*;&]+?)\s*((?:\d+|\*)(?:\s*,\s*(?:\d+|\*))*)\s*(?:,|;|&|\band\b|$)", s):
            name = re.sub(r"^and\s+", "", mt.group(1).strip(), flags=re.I)
            marks = [x for x in re.split(r"[,\s]+", mt.group(2)) if x]
            authors.append({"name": name, "marks": marks})
    return authors


def merge_affiliations(art):
    """Aynı affiliasyonları birleştirir ve yazar numaralarını yeniden düzenler."""
    norm = lambda s: re.sub(r"[^a-z0-9]", "", s.lower())
    old = {str(k + 1): a for k, a in enumerate(art.affiliations)}
    new_list, mapping = [], {}
    for au in art.authors:
        for mk in au["marks"]:
            if mk in old and mk not in mapping:
                key = norm(old[mk])
                for j, a in enumerate(new_list):
                    if norm(a) == key:
                        mapping[mk] = str(j + 1)
                        break
                else:
                    new_list.append(old[mk].rstrip(".") + ".")
                    mapping[mk] = str(len(new_list))
    for k, a in old.items():  # yazara bağlanmamış affiliasyonlar
        if k not in mapping:
            key = norm(a)
            if not any(norm(x) == key for x in new_list):
                new_list.append(a.rstrip(".") + ".")
            mapping[k] = str([norm(x) for x in new_list].index(key) + 1)
    for au in art.authors:
        marks = []
        for mk in au["marks"]:
            mk = mapping.get(mk, mk)
            if mk not in marks:
                marks.append(mk)
        au["marks"] = sorted([m_ for m_ in marks if m_ != "*"], key=lambda x: (not x.isdigit(), x)) + \
            (["*"] if "*" in marks else [])
    if len(new_list) != len(art.affiliations):
        print("Bilgi: %d affiliasyon -> %d (aynı olanlar birleştirildi)" % (len(art.affiliations), len(new_list)))
    art.affiliations = new_list


def apply_cover_letter(art, cl):
    if not cl:
        return
    if not art.authors and cl["authors"]:
        print("Bilgi: makalede yazar satırı bulunamadı, kapak mektubundaki yazarlar kullanılıyor.")
        for a in cl["authors"]:
            aff = a.get("affiliation", "")
            if aff and aff not in art.affiliations:
                art.affiliations.append(aff)
            marks = [str(art.affiliations.index(aff) + 1)] if aff else []
            if a.get("corresponding"):
                marks.append("*")
            art.authors.append({"name": a["name"], "marks": marks})
    for au in art.authors:
        a = match_author(au["name"], cl["authors"])
        if not a:
            warn("Kapak mektubunda '%s' için bilgi bulunamadı (ORCID eklenmedi)." % au["name"])
            continue
        full = fuller_name(au["name"], a["name"])
        if full != au["name"]:
            print("Bilgi: yazar adı kapak mektubundaki tam adla yazıldı: %s → %s" % (au["name"], full))
            au["name"] = full
        au["orcid"] = a.get("orcid")
        au["email"] = a.get("email")
        au["tel"] = a.get("tel")
        if a.get("corresponding") and "*" not in au["marks"] and not any("*" in x["marks"] for x in art.authors):
            au["marks"].append("*")
    corr = next((a for a in art.authors if "*" in a["marks"]), art.authors[0] if art.authors else None)
    if corr:
        if not art.corr_email and corr.get("email"):
            art.corr_email = corr["email"]
        elif art.corr_email and corr.get("email") and corr["email"].lower() != art.corr_email.lower():
            warn("Sorumlu yazar e-postası farklı: makale '%s', kapak mektubu '%s'. Makaledeki kullanıldı."
                 % (art.corr_email, corr["email"]))
        if not art.corr_tel and corr.get("tel"):
            art.corr_tel = corr["tel"]
    if cl.get("title") and art.title and "".join(name_key(cl["title"])) != "".join(name_key(art.title)):
        warn("Makale başlığı ile kapak mektubundaki başlık farklı.")
    if not art.title and cl.get("title"):
        art.title = cl["title"]


def parse_article(doc, cover=None):
    blocks = doc.body_blocks()
    art, i = parse_front(blocks)
    rest = blocks[i:]
    # Kaynaklar başlığını bul
    ref_idx = None
    for k, b in enumerate(rest):
        if isinstance(b, Para) and REF_HEAD.match(clean_heading(b.text())):
            ref_idx = k
    if ref_idx is not None:
        art.body = rest[:ref_idx]
        art.references = [b for b in rest[ref_idx + 1:] if isinstance(b, Para) and b.text().strip()]
    else:
        warn("'References' başlığı bulunamadı; kaynakça boş.")
        art.body = rest
    merge_affiliations(art)
    apply_cover_letter(art, cover)
    if art.authors and not any("*" in a["marks"] for a in art.authors) and art.corr_email:
        # yıldız konmamış: e-postası eşleşen yazar, yoksa ilk yazar sorumlu yazardır
        mail = art.corr_email.lower()
        corr = next((a for a in art.authors if (a.get("email") or "").lower() == mail), None) or \
            next((a for a in art.authors if name_key(a["name"])[-1] in mail), art.authors[0])
        corr["marks"].append("*")
    return art
