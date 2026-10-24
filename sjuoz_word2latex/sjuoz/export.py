# -*- coding: utf-8 -*-
"""Dışa aktarma: Crossref (DOI kaydı), JATS XML, BibTeX."""
import datetime
import os
import re
import unicodedata

from lxml import etree

from .common import DATA_DIR, DOI_DUP_RE, title_case_caps
from .docx import Para, Table
from .render import strip_eq_number_text
from .parse import BACK_MATTER, CAPTION_RE, clean_heading
from .figures import layout_table_images
from .template import MONTHS

# --------------------------------------------------------------------------
# Dışa aktarma: Crossref (DOI kaydı), JATS (indeksleme), BibTeX
# --------------------------------------------------------------------------
JOURNAL = {
    "title": "Science Journal of University of Zakho",
    "abbrev": "SJUOZ",
    "publisher": "University of Zakho",
    "issn_print": "2663-628X",
    "issn_electronic": "2663-6298",
    "doi_prefix": "10.25271",
    "article_url": "https://sjuoz.uoz.edu.krd/index.php/sjuoz/article/view/%s",
    "license_url": "https://creativecommons.org/licenses/by-nc-sa/4.0/",
    "license_text": "This is an open access article under a CC BY-NC-SA 4.0 license",
    "depositor_name": "Science Journal of University of Zakho",
    "depositor_email": "",
}
CROSSREF_NS = "http://www.crossref.org/schema/5.3.1"
JATS_NS = "http://www.ncbi.nlm.nih.gov/JATS1"
XLINK_NS = "http://www.w3.org/1999/xlink"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
DOI_RE = re.compile(r"\b(10\.\d{4,9}/[^\s\"<>]+)")


def article_doi(opts):
    return "%s/sjuoz.%s.%s.%s.%s" % (JOURNAL["doi_prefix"], opts.year, opts.volume, opts.issue, opts.article_id)


def split_name(name):
    """'Mohammed E. Safi' -> ('Mohammed E.', 'Safi')"""
    parts = name.split()
    if len(parts) < 2:
        return "", name
    return " ".join(parts[:-1]), parts[-1]


def find_doi(text):
    text = DOI_DUP_RE.sub("https://doi.org/", text or "")
    mt = DOI_RE.search(text)
    if not mt:
        return None
    return re.sub(r"[.,;)\]]+$", "", mt.group(1))


def parse_date(text):
    """'12 May 2026', '2026-05-12', '12/05/2026' -> (yıl, ay, gün) ya da None."""
    text = (text or "").strip()
    if not text or text.upper() in ("XX", "XXX"):
        return None
    for fmt in ("%d %B %Y", "%d %b %Y", "%B %d, %Y", "%Y-%m-%d", "%d/%m/%Y", "%d.%m.%Y", "%B %Y", "%b %Y"):
        try:
            d = datetime.datetime.strptime(text, fmt)
            return d.year, d.month, (d.day if "%d" in fmt else None)
        except ValueError:
            continue
    return None


def article_metadata(art, opts, first_page, last_page):
    """Dışa aktarmalarda ortak kullanılan makale bilgileri."""
    affs = art.affiliations
    authors = []
    for k, au in enumerate(art.authors):
        given, surname = split_name(au["name"])
        nums = [int(m_) - 1 for m_ in au["marks"] if m_.isdigit() and 0 < int(m_) <= len(affs)]
        if not nums and len(affs) == 1:
            nums = [0]
        authors.append({"given": given, "surname": surname, "name": au["name"], "orcid": au.get("orcid"),
                        "affiliations": [affs[i] for i in nums], "aff_ids": nums,
                        "corresponding": "*" in au["marks"], "email": au.get("email"),
                        "sequence": "first" if k == 0 else "additional"})
    month_no = MONTHS.index(opts.month) + 1 if opts.month in MONTHS else None
    title = re.sub(r"\s+", " ", art.title).strip()
    if title.upper() == title:  # TAMAMI BÜYÜK HARF başlık: meta verilerde başlık biçimi
        title = title_case_caps(title)
    return {
        "title": title, "article_type": opts.article_type or art.article_type,
        "authors": authors, "affiliations": affs, "corr_email": art.corr_email,
        "abstract": [re.sub(r"\s+", " ", t).strip() for t in art.abstract_plain if t.strip()],
        "keywords": [k.strip(" .") for k in re.split(r"[;,]", art.keywords) if k.strip(" .")],
        "doi": article_doi(opts), "url": JOURNAL["article_url"] % opts.article_id,
        "volume": str(opts.volume), "issue": str(opts.issue), "year": str(opts.year), "month": month_no,
        "first_page": first_page, "last_page": last_page, "article_id": str(opts.article_id),
        "received": parse_date(opts.received), "accepted": parse_date(opts.accepted),
        "published": parse_date(opts.published),
        "references": [re.sub(r"\s+", " ", p.text()).strip() for p in art.references],
    }


def _el(parent, tag, text=None, **attrs):
    """Crossref öğesi (öznitelik adları alt çizgili: publication_type, contributor_role)."""
    e = etree.SubElement(parent, tag, {k: str(v) for k, v in attrs.items() if v is not None})
    if text is not None:
        e.text = str(text)
    return e


def _jel(parent, tag, text=None, **attrs):
    """JATS öğesi (öznitelik adları tireli: article-type, pub-id-type)."""
    return _el(parent, tag, text, **{k.replace("_", "-"): v for k, v in attrs.items()})


def crossref_xml(articles, depositor_name=None, depositor_email=None, batch_id=None):
    """Bir veya daha çok makale için Crossref 5.3.1 'doi_batch' XML'i (aynı sayıdaki makaleler)."""
    C = "{%s}" % CROSSREF_NS
    J = "{%s}" % JATS_NS
    root = etree.Element(C + "doi_batch", nsmap={None: CROSSREF_NS, "jats": JATS_NS, "xsi": XSI_NS},
                         version="5.3.1")
    root.set("{%s}schemaLocation" % XSI_NS,
             "%s https://www.crossref.org/schemas/crossref5.3.1.xsd" % CROSSREF_NS)
    now = datetime.datetime.now()
    head = _el(root, C + "head")
    _el(head, C + "doi_batch_id", batch_id or "sjuoz-%s" % now.strftime("%Y%m%d%H%M%S"))
    _el(head, C + "timestamp", now.strftime("%Y%m%d%H%M%S"))
    dep = _el(head, C + "depositor")
    _el(dep, C + "depositor_name", depositor_name or JOURNAL["depositor_name"])
    _el(dep, C + "email_address", depositor_email or JOURNAL["depositor_email"] or "editor@example.org")
    _el(head, C + "registrant", JOURNAL["publisher"])
    body = _el(root, C + "body")
    journal = _el(body, C + "journal")
    jm = _el(journal, C + "journal_metadata", language="en")
    _el(jm, C + "full_title", JOURNAL["title"])
    _el(jm, C + "abbrev_title", JOURNAL["abbrev"])
    _el(jm, C + "issn", JOURNAL["issn_print"], media_type="print")
    _el(jm, C + "issn", JOURNAL["issn_electronic"], media_type="electronic")
    first = articles[0]
    ji = _el(journal, C + "journal_issue")
    pd = _el(ji, C + "publication_date", media_type="online")
    if first["month"]:
        _el(pd, C + "month", "%02d" % first["month"])
    _el(pd, C + "year", first["year"])
    jv = _el(ji, C + "journal_volume")
    _el(jv, C + "volume", first["volume"])
    _el(ji, C + "issue", first["issue"])
    for a in articles:
        ja = _el(journal, C + "journal_article", publication_type="full_text", language="en")
        titles = _el(ja, C + "titles")
        _el(titles, C + "title", a["title"])
        if a["authors"]:
            contribs = _el(ja, C + "contributors")
            for au in a["authors"]:
                pn = _el(contribs, C + "person_name", sequence=au["sequence"], contributor_role="author")
                if au["given"]:
                    _el(pn, C + "given_name", au["given"])
                _el(pn, C + "surname", au["surname"])
                if au["affiliations"]:
                    affs = _el(pn, C + "affiliations")
                    for aff in au["affiliations"]:
                        inst = _el(affs, C + "institution")
                        _el(inst, C + "institution_name", aff.rstrip("."))
                if au["orcid"]:
                    _el(pn, C + "ORCID", au["orcid"], authenticated="false")
        if a["abstract"]:
            ab = _el(ja, J + "abstract")
            for para in a["abstract"]:
                _el(ab, J + "p", para)
        pub = a["published"]
        pd = _el(ja, C + "publication_date", media_type="online")
        if pub and pub[1]:
            _el(pd, C + "month", "%02d" % pub[1])
            if pub[2]:
                _el(pd, C + "day", "%02d" % pub[2])
        elif a["month"]:
            _el(pd, C + "month", "%02d" % a["month"])
        _el(pd, C + "year", pub[0] if pub else a["year"])
        if a["first_page"]:
            pages = _el(ja, C + "pages")
            _el(pages, C + "first_page", a["first_page"])
            if a["last_page"]:
                _el(pages, C + "last_page", a["last_page"])
        lic = _el(ja, "{http://www.crossref.org/AccessIndicators.xsd}program",
                  name="AccessIndicators")
        lic.nsmap  # noqa  (ad alanı lxml tarafından eklenir)
        _el(lic, "{http://www.crossref.org/AccessIndicators.xsd}license_ref", JOURNAL["license_url"])
        dd = _el(ja, C + "doi_data")
        _el(dd, C + "doi", a["doi"])
        _el(dd, C + "resource", a["url"])
        if a["references"]:
            cl = _el(ja, C + "citation_list")
            for k, ref in enumerate(a["references"]):
                cit = _el(cl, C + "citation", key="ref%d" % (k + 1))
                doi = find_doi(ref)
                if doi:
                    _el(cit, C + "doi", doi)
                _el(cit, C + "unstructured_citation", ref)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True)


_CROSSREF_SCHEMA = {}


def validate_crossref(xml_bytes):
    """Crossref 5.3.1 şemasıyla doğrular. Döndürür: (geçerli mi, hata listesi) ya da (None, []) şema yoksa."""
    xsd = os.path.join(DATA_DIR, "schemas", "crossref", "crossref5.3.1.xsd")
    if not os.path.isfile(xsd):
        return None, []
    if "s" not in _CROSSREF_SCHEMA:
        _CROSSREF_SCHEMA["s"] = etree.XMLSchema(etree.parse(xsd))
    schema = _CROSSREF_SCHEMA["s"]
    ok = schema.validate(etree.fromstring(xml_bytes))
    return ok, [str(e.message) for e in schema.error_log][:10]


def jats_xml(meta, art, writer):
    """JATS 1.3 (Journal Publishing) XML: ön bilgiler, gövde, kaynakça."""
    X = "{%s}" % XLINK_NS
    root = etree.Element("article", nsmap={"xlink": XLINK_NS, "mml": "http://www.w3.org/1998/Math/MathML"})
    root.set("article-type", "review-article" if "review" in (meta["article_type"] or "").lower()
             else "research-article")
    root.set("dtd-version", "1.3")
    root.set("{http://www.w3.org/XML/1998/namespace}lang", "en")
    front = _jel(root, "front")
    jm = _jel(front, "journal-meta")
    _jel(jm, "journal-id", JOURNAL["abbrev"], journal_id_type="publisher-id")
    jtg = _jel(jm, "journal-title-group")
    _jel(jtg, "journal-title", JOURNAL["title"])
    _jel(jtg, "abbrev-journal-title", JOURNAL["abbrev"])
    _jel(jm, "issn", JOURNAL["issn_print"], pub_type="ppub")
    _jel(jm, "issn", JOURNAL["issn_electronic"], pub_type="epub")
    pub = _jel(jm, "publisher")
    _jel(pub, "publisher-name", JOURNAL["publisher"])
    am = _jel(front, "article-meta")
    _jel(am, "article-id", meta["doi"], pub_id_type="doi")
    _jel(am, "article-id", meta["article_id"], pub_id_type="publisher-id")
    tg = _jel(am, "title-group")
    _jel(tg, "article-title", meta["title"])
    cg = _jel(am, "contrib-group")
    for au in meta["authors"]:
        c = _jel(cg, "contrib", contrib_type="author", corresp="yes" if au["corresponding"] else None)
        if au["orcid"]:
            _jel(c, "contrib-id", au["orcid"], contrib_id_type="orcid", authenticated="false")
        nm = _jel(c, "name")
        _jel(nm, "surname", au["surname"])
        if au["given"]:
            _jel(nm, "given-names", au["given"])
        for i in au["aff_ids"]:
            _jel(c, "xref", str(i + 1), ref_type="aff", rid="aff%d" % (i + 1))
        if au["corresponding"]:
            _jel(c, "xref", "*", ref_type="corresp", rid="cor1")
    for i, aff in enumerate(meta["affiliations"]):
        a = _jel(am, "aff", id="aff%d" % (i + 1))
        _jel(a, "label", str(i + 1))
        _jel(a, "institution", aff.rstrip("."))
    if meta["corr_email"]:
        an = _jel(am, "author-notes")
        co = _jel(an, "corresp", "* Corresponding author: ", id="cor1")
        _jel(co, "email", meta["corr_email"])
    pd = _jel(am, "pub-date", publication_format="electronic", date_type="pub")
    pubd = meta["published"]
    if pubd and pubd[2]:
        _jel(pd, "day", "%02d" % pubd[2])
    if (pubd and pubd[1]) or meta["month"]:
        _jel(pd, "month", "%02d" % ((pubd and pubd[1]) or meta["month"]))
    _jel(pd, "year", pubd[0] if pubd else meta["year"])
    _jel(am, "volume", meta["volume"])
    _jel(am, "issue", meta["issue"])
    if meta["first_page"]:
        _jel(am, "fpage", meta["first_page"])
    if meta["last_page"]:
        _jel(am, "lpage", meta["last_page"])
    hist = [(t, d) for t, d in (("received", meta["received"]), ("accepted", meta["accepted"])) if d]
    if hist:
        h = _jel(am, "history")
        for t, d in hist:
            de = _jel(h, "date", date_type=t)
            if d[2]:
                _jel(de, "day", "%02d" % d[2])
            if d[1]:
                _jel(de, "month", "%02d" % d[1])
            _jel(de, "year", d[0])
    perm = _jel(am, "permissions")
    _jel(perm, "copyright-year", meta["year"])
    lic = _jel(perm, "license")
    lic.set(X + "href", JOURNAL["license_url"])
    _jel(lic, "license-p", JOURNAL["license_text"])
    self_uri = _jel(am, "self-uri")
    self_uri.set(X + "href", meta["url"])
    if meta["abstract"]:
        ab = _jel(am, "abstract")
        for para in meta["abstract"]:
            _jel(ab, "p", para)
    if meta["keywords"]:
        kg = _jel(am, "kwd-group", kwd_group_type="author")
        for k in meta["keywords"]:
            _jel(kg, "kwd", k)

    # gövde: bölümler, paragraflar, denklemler, şekil ve tablo başlıkları
    body = _jel(root, "body")
    back = _jel(root, "back")
    sec = None
    sec_stack = []
    n_tab = n_eq = 0
    figures = [f for f in writer.figure_log if f["caption"]]
    for b in art.body:
        if isinstance(b, Table):
            if layout_table_images(b):
                continue
            n_tab += 1
            tw = _jel(sec if sec is not None else body, "table-wrap", id="tab%d" % n_tab)
            _jel(tw, "label", "Table %d" % n_tab)
            if n_tab <= len(writer.table_log) and writer.table_log[n_tab - 1]:
                _jel(_jel(tw, "caption"), "p", writer.table_log[n_tab - 1])
            tbl = _jel(tw, "table")
            for ri, row in enumerate(b.rows):
                tr = _jel(tbl, "tr")
                for paras, span, vmerge in row:
                    if vmerge == "continue":
                        continue
                    cell = _jel(tr, "th" if ri == 0 else "td", colspan=span if span > 1 else None)
                    toks = []
                    for p in [p for p in paras if isinstance(p, Para) and not p.is_empty()]:
                        if toks:
                            toks.append(("t", " ", (False, False, None, None)))
                        toks += p.tokens
                    _jats_inline(cell, toks, strip_bold=(ri == 0))
            continue
        p = b
        if p.is_empty():
            continue
        txt = p.text().strip()
        clean = clean_heading(txt)
        lvl = p.heading_level()
        is_head = bool(lvl) or (not p.maths and len(txt) <= 150 and p.all_bold() and not txt.endswith((".", ",", ";")))
        if is_head:
            if BACK_MATTER.match(clean):
                if re.match(r"acknowledg", clean, re.I):
                    sec = _jel(back, "ack")
                else:
                    sec = _jel(back, "sec")
                _jel(sec, "title", clean)
                sec_stack = [sec]
                continue
            level = lvl or (p.ilvl + 1 if p.num_id else 1)
            while len(sec_stack) >= level:
                sec_stack.pop()
            parent = sec_stack[-1] if sec_stack else body
            sec = _jel(parent, "sec")
            _jel(sec, "title", clean)
            sec_stack.append(sec)
            continue
        container = sec if sec is not None else body
        cap = CAPTION_RE.match(txt)
        if cap and (p.jc == "center" or "caption" in p.style.lower() or p.images or
                    re.match(r"^\s*(figure|fig\.?|table)\s*\(?\s*\d+\s*\)?\s*[:.\-–—]", txt, re.I)):
            if cap.group(1).lower().startswith("fig") and figures:
                f = figures.pop(0)
                fig = _jel(container, "fig", id="fig%d" % f["number"])
                _jel(fig, "label", "Figure %d" % f["number"])
                _jel(_jel(fig, "caption"), "p", f["caption"])
                for fn in f["files"]:
                    g = _jel(fig, "graphic")
                    g.set(X + "href", fn)
            continue
        if p.maths and strip_eq_number_text(txt)[0]:
            n_eq += 1
            df = _jel(container, "disp-formula", id="eq%d" % n_eq)
            _jel(df, "tex-math", "\\[" + " ".join(m_[1] for m_ in p.maths) + "\\]")
            continue
        if p.images:
            continue
        para = _jel(container, "p")
        _jats_inline(para, p.tokens)
    if not len(back) and not art.references:
        root.remove(back)
    if art.references:
        rl = _jel(back, "ref-list")
        _jel(rl, "title", "References")
        for k, ref in enumerate(meta["references"]):
            r = _jel(rl, "ref", id="ref%d" % (k + 1))
            mc = _jel(r, "mixed-citation", ref)
            doi = find_doi(ref)
            if doi:
                mc.text = ref.replace(doi, "").rstrip()
                _jel(mc, "pub-id", doi, pub_id_type="doi")
    if not len(body):
        _jel(body, "p")
    doctype = ('<!DOCTYPE article PUBLIC "-//NLM//DTD JATS (Z39.96) Journal Publishing DTD v1.3 20210610//EN" '
               '"https://jats.nlm.nih.gov/publishing/1.3/JATS-journalpublishing1-3.dtd">')
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=True, doctype=doctype)


def _jats_inline(parent, tokens, strip_bold=False):
    """Biçimli metin parçalarını JATS satır içi öğelerine çevirir (italic, bold, sup, sub, tex-math).
    strip_bold: tablo başlık hücrelerinde kalınlık <th> ile zaten belli."""
    last = None

    def add_text(text):
        if last is None:
            parent.text = (parent.text or "") + text
        else:
            last.tail = (last.tail or "") + text

    for t in tokens:
        if t[0] == "t":
            text = t[1].replace("\n", " ")
            bold, italic, vert, _ = t[2]
            bold = bold and not strip_bold
            tag = "sup" if vert == "superscript" else "sub" if vert == "subscript" else \
                "bold" if bold else "italic" if italic else None
            if tag:
                last = etree.SubElement(parent, tag)
                last.text = text
            else:
                add_text(text)
        elif t[0] == "m" and t[1]:
            f = etree.SubElement(parent, "inline-formula")
            _jel(f, "tex-math", "\\(" + t[1] + "\\)")
            last = f
    if parent.text:
        parent.text = re.sub(r"\s+", " ", parent.text).lstrip()
    if len(parent) and parent[-1].tail:
        parent[-1].tail = parent[-1].tail.rstrip()
    elif parent.text and not len(parent):
        parent.text = parent.text.rstrip()


# ---------- BibTeX
def _bib_escape(s):
    return s.replace("\\", "/").replace("{", "(").replace("}", ")").replace("&", r"\&").replace("%", r"\%").replace("_", r"\_")


def parse_reference(text):
    """APA ya da IEEE biçimindeki bir kaynağı alanlarına ayırmaya çalışır (bulamadığını boş bırakır)."""
    ref = {"raw": text, "doi": find_doi(text)}
    t = re.sub(r"^\s*\[\d+\]\s*", "", text)
    t = DOI_DUP_RE.sub("https://doi.org/", t)
    t = re.sub(r"\s*(?:https?://(?:dx\.)?doi\.org/|doi:\s*)\S+", "", t, flags=re.I).strip()
    apa = re.match(r"^(?P<auth>.+?)\s*\((?P<year>(?:19|20)\d\d[a-z]?|n\.\s?d\.)\)\.?\s*(?P<rest>.*)$", t)
    if apa:
        ref["authors"] = apa.group("auth").strip(" .,")
        ref["year"] = apa.group("year")
        rest = apa.group("rest")
        mt = re.match(r"^(?P<title>.+?[.?!])\s+(?P<src>.*)$", rest)
        if mt:
            ref["title"] = mt.group("title").rstrip(".")
            src = mt.group("src")
            pm = re.match(r"^(?P<journal>.+?),\s*(?:pp\.\s*)?(?P<pages>\d+\s*[–\-]+\s*\d+)\s*\.?\s*$", src)
            sm = re.match(r"^(?P<journal>[^,]+?),\s*(?P<vol>\d+)(?:\((?P<num>[^)]+)\))?(?:,\s*(?P<pages>[\d–\-]+))?", src)
            vol_is_page = sm and re.match(r"\s*[–\-]", src[sm.end("vol"):])
            if pm and (not sm or vol_is_page):  # bildiri: "Kitap/konferans adı, 132–137."
                ref["journal"] = pm.group("journal").strip()
                ref["pages"] = re.sub(r"\s*[–\-]+\s*", "--", pm.group("pages"))
            elif sm:
                ref["journal"] = sm.group("journal").strip()
                ref["volume"] = sm.group("vol")
                ref["number"] = sm.group("num")
                ref["pages"] = (sm.group("pages") or "").replace("–", "--") or None
            else:
                ref["howpublished"] = src.strip(" .")
        else:
            ref["title"] = rest.strip(" .")
        return ref
    ieee = re.match(r'^(?P<auth>[^"“]+?),\s*["“](?P<title>[^"”]+?),?["”],?\s*(?P<rest>.*)$', t)
    if ieee:
        ref["authors"] = ieee.group("auth").strip(" ,")
        ref["title"] = ieee.group("title").strip(" ,")
        rest = ieee.group("rest")
        jm = re.match(r"^(?:in\s+)?(?P<journal>[^,]+)", rest)
        if jm:
            ref["journal"] = jm.group("journal").strip()
        for key, pat in (("volume", r"vol\.\s*(\d+)"), ("number", r"no\.\s*(\d+)"),
                         ("pages", r"pp\.\s*([\d–\-]+)"), ("year", r"\b((?:19|20)\d\d)\b")):
            m2 = re.search(pat, rest)
            if m2:
                ref[key] = m2.group(1).replace("–", "--")
    return ref


def bibtex_authors(text):
    """APA yazar listesi -> BibTeX: 'Safi, M. E., & Abbas, E. I.' -> 'Safi, M. E. and Abbas, E. I.'
    IEEE ('A. Author, B. Author, and C. Author') de desteklenir."""
    t = re.sub(r",?\s*(?:&|\band\b)\s*", ", ", text.strip(" ,."))
    parts = [x.strip() for x in t.split(",") if x.strip()]
    full = re.compile(r"^(?:[A-Z]\.[\s\-]*)+[A-Z][\w'’\-]+$")      # "A. Author" (IEEE)
    only_initials = re.compile(r"^(?:[A-Z]\.?[\s\-]*)+$")          # "M. E."
    names, k = [], 0
    while k < len(parts):
        cur = parts[k]
        if only_initials.match(cur) and names:   # bozuk listelerde sahipsiz baş harfler
            names[-1] += ", " + cur
            k += 1
            continue
        if k + 1 < len(parts) and not full.match(cur) and not full.match(parts[k + 1]):
            given = parts[k + 1]
            names.append("%s, %s" % (cur, given + ("." if only_initials.match(given) and not given.endswith(".") else "")))
            k += 2
        else:
            names.append(cur)
            k += 1
    return " and ".join(names)


def bibtex_key(ref, used):
    first = re.split(r"[ ,.]", (ref.get("authors") or "ref").strip())[0] or "ref"
    first = re.sub(r"[^A-Za-z]", "", unicodedata.normalize("NFKD", first).encode("ascii", "ignore").decode()) or "ref"
    key = "%s%s" % (first.lower(), re.sub(r"\D", "", ref.get("year") or "") or "nd")
    base, n = key, 0
    while key in used:
        n += 1
        key = base + "abcdefghijklmnopqrstuvwxyz"[n - 1]
    used.add(key)
    return key


def bibtex_references(art):
    used, out = set(), []
    for p in art.references:
        text = re.sub(r"\s+", " ", p.text()).strip()
        ref = parse_reference(text)
        key = bibtex_key(ref, used)
        proceedings = bool(ref.get("journal") and re.search(
            r"conference|proceedings|symposium|workshop|congress|\bICC|\bIEEE Cat\.", ref["journal"], re.I))
        if proceedings:
            kind = "inproceedings"
            ref["booktitle"] = ref.pop("journal")
        elif ref.get("journal") and ref.get("title"):
            kind = "article"
        elif ref.get("title"):
            kind = "misc"
        else:
            kind = "misc"
        fields = []
        if ref.get("authors"):
            fields.append(("author", bibtex_authors(ref["authors"])))
        for f in ("title", "journal", "booktitle", "year", "volume", "number", "pages", "howpublished", "doi"):
            if ref.get(f):
                fields.append((f, ref[f]))
        fields.append(("note", text if kind == "misc" and not ref.get("title") else None))
        body = ",\n".join("  %s = {%s}" % (k, _bib_escape(v)) for k, v in fields if v)
        out.append("@%s{%s,\n%s\n}\n" % (kind, key, body))
    return "\n".join(out)


def bibtex_self(meta):
    """Bu makaleye atıf için BibTeX kaydı."""
    first = meta["authors"][0]["surname"] if meta["authors"] else "sjuoz"
    key = re.sub(r"[^A-Za-z]", "", unicodedata.normalize("NFKD", first).encode("ascii", "ignore").decode()).lower() \
        + meta["year"]
    fields = [("author", " and ".join("%s, %s" % (a["surname"], a["given"]) if a["given"] else a["surname"]
                                      for a in meta["authors"])),
              ("title", meta["title"]), ("journal", JOURNAL["title"]), ("year", meta["year"]),
              ("volume", meta["volume"]), ("number", meta["issue"]),
              ("pages", "%s--%s" % (meta["first_page"], meta["last_page"]) if meta["last_page"]
               else (str(meta["first_page"]) if meta["first_page"] else None)),
              ("doi", meta["doi"]), ("url", meta["url"]), ("issn", JOURNAL["issn_electronic"])]
    body = ",\n".join("  %s = {%s}" % (k, _bib_escape(str(v))) for k, v in fields if v)
    return "@article{%s,\n%s\n}\n" % (key, body)
