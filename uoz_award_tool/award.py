# -*- coding: utf-8 -*-
"""
University of Zakho - Publication Award letter generator (خەلاتکرن).

What it does:
  1) Reads the award request letter submitted by the author(s) (Word, PDF or image):
     authors, paper title, journal, IF / index, department, college, DOI.
  2) Verifies the paper online: is it published, DOI, article link on the journal site,
     authors and their affiliations, UoZ publication registry (rule 5-a),
     UoZ staff portal (rule 5-b), DOAJ.
  3) Calculates the award amount itself from the university rules
     (Clarivate IF table > Scopus quartile > approved university journals).
  4) Fills the template Word letter (template.docx) and saves it in a folder named after the authors:
        output/(author1) و (author2)/
            (author1) و (author2) - خەلاتکرن.docx   <- the award letter (Kurdish, official)
            report.txt       <- links + verification results + warnings
            data.json        <- everything that was read/found (edit and re-run with --json)
            request.<ext>    <- copy of the submitted request letter
        output/summary.html  <- all papers in one table with clickable links

Two modes:
  FREE mode (default, no API key): rule-based reading + free web sources
      (Crossref, OpenAlex, DOAJ, UoZ registry, UoZ staff portal). See free_mode.py.
  AI mode (needs an Anthropic API key in settings.json): Claude reads the letter and searches the web.

Usage:
    python award.py                                 # process every file in the 'input' folder
    python award.py request.docx [request2.pdf ...]
    python award.py --json "output/.../data.json"   # rebuild the Word letter from (edited) data
    python award.py --no-web request.docx           # skip online verification
    python award.py --free request.docx             # force free mode even if an API key exists
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
TEMPLATE = os.path.join(HERE, "template.docx")
OUT_ROOT = os.path.join(HERE, "output")
IN_DIR = os.path.join(HERE, "input")
SETTINGS_FILES = [os.path.join(HERE, "settings.json"), os.path.join(HERE, "ayarlar.json")]
MODEL = "claude-opus-5-5"
VERSION = "2026-11-19-en1"

# ─────────────────────────── Award rules ───────────────────────────
# Ministerial order no. 1692 (23/01/2025)
SCOPUS_AMOUNT = {1: 500_000, 2: 400_000, 3: 300_000, 4: 300_000, None: 250_000}
UNI_KRG_AMOUNT = 250_000      # journals of KRG universities approved by the Ministry
UNI_IRAQ_AMOUNT = 200_000     # journals of Iraqi public universities / universities outside Iraq
MAX_AUTHORS = {"pure": 4, "humanities": 3}

COLLEGES = {
    # key: (Kurdish, Arabic, English, field type)
    "Science": ("زانست", "العلوم", "Science", "pure"),
}


def clarivate_amount(impact: float) -> int:
    """IF table: 0.1-1 = 500,000 ; every +0.1 = +20,000 ; 6 and above = 1,500,000."""
    if impact <= 1.0:
        return 500_000
    if impact >= 6.0:
        return 1_500_000
    tenth = math.floor(impact * 10 + 1e-9)          # 2.75 -> 27 (2.7)
    return 500_000 + 20_000 * (tenth - 10)


def decide_award(d: dict) -> dict:
    """Category, amount and clause (خالا). Uses verified values when available,
    otherwise what the request letter states."""
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


# ─────────────────────────── Helpers ───────────────────────────
_MAP = str.maketrans({"ی": "ي", "ى": "ي", "ێ": "ي", "ک": "ك", "ە": "ه", "ۆ": "و", "ڤ": "ف",
                      "ڕ": "ر", "ڵ": "ل", "گ": "ك", "پ": "ب", "چ": "ج", "ژ": "ز", "ة": "ه",
                      "أ": "ا", "إ": "ا", "آ": "ا"})


def norm(s):
    s = (s or "").lower().translate(_MAP)
    s = re.sub(r"[\u200c\u200d\u064b-\u065f]", "", s)
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


def setting(name):
    for path in SETTINGS_FILES:
        if os.path.exists(path):
            try:
                with open(path, encoding="utf-8") as f:
                    val = (json.load(f).get(name) or "").strip()
                if val:
                    return val
            except ValueError:
                pass
    return ""


def api_key():
    return os.environ.get("ANTHROPIC_API_KEY", "").strip() or setting("anthropic_api_key")


def api_client():
    import anthropic
    key = api_key()
    return anthropic.Anthropic(api_key=key) if key else anthropic.Anthropic()


# ─────────────────────────── 1) Reading the request letter (AI mode) ───────────────────────────
class Author(BaseModel):
    name_ku: str = Field(description="Name exactly as written in Arabic/Kurdish script in the letter")
    name_en: str = Field(description="Name exactly as written in Latin script in the letter; empty if absent")


class Letter(BaseModel):
    authors: List[Author]
    title: str = Field(description="Full paper title as written in the letter")
    journal: str
    stated_impact_factor: Optional[float] = None
    stated_index: str = Field(description="Index information stated in the letter (Clarivate/Scopus Q?/DOAJ...), or empty")
    stated_scopus_quartile: Optional[int] = None
    department_ku: str = Field(description="Department name in Badini Kurdish, e.g. فیزیک ، ماتماتیک")
    college_key: Literal["Science", "Other"]
    college_name: str
    handwritten_amount: Optional[int] = Field(None, description="Amount handwritten on the paper (dinar), or null")
    letter_date: str = ""
    unclear: List[str] = Field(default_factory=list, description="Parts that were hard to read or uncertain")


READ_PROMPT = """This (Word file, PDF or image) is a request letter from a University of Zakho staff member asking for \
a publication award (خەلاتکرن) for a published paper (Badini Kurdish / Arabic / English). Extract its information.
- Copy author names exactly as written in the letter; do not correct them.
- Give the department name in Kurdish (e.g. فیزیک). If the college is Science (زانست / العلوم), college_key = "Science".
- If an amount is handwritten on the paper (green/blue pen), put only the number in handwritten_amount.
- Put everything you are not sure about in the unclear list."""


def file_blocks(path):
    ext = os.path.splitext(path)[1].lower()
    with open(path, "rb") as f:
        data = f.read()
    if ext == ".pdf":
        return [{"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                                 "data": base64.standard_b64encode(data).decode()}}]
    if ext == ".docx":
        return [{"type": "text", "text": "Request letter (Word file) text:\n" + docx_text(path)}]
    if ext == ".doc":
        raise RuntimeError("The old .doc format is not supported. Open it in Word and use 'Save As > .docx'.")
    # image: shrink if needed (API limit ~5 MB)
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
    """Write paragraphs and tables in order; join table cells with ' | '.
    Also reads text boxes, and skips Word's duplicate copies (mc:Fallback)."""
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
    """Extract all text from a Word file (headers/footers, tables and text boxes included)."""
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
            label = "HEADER" if "header" in n else ("FOOTER" if "footer" in n else "BODY")
            out.append(f"[{label}]\n" + "\n".join(lines))
    return "\n\n".join(out)


# ─────────────────────────── OCR (scanned PDFs / images, free mode) ───────────────────────────
TESSERACT_PATHS = [r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                   r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                   os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")]

OCR_NOTES = []


def tesseract_langs(cmd):
    """Read 'tesseract --list-langs' ourselves: pytesseract drops names like 'script\\Arabic'."""
    import subprocess
    out = subprocess.run([cmd, "--list-langs"], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    if out.returncode not in (0, 1):
        raise RuntimeError("tesseract did not run")
    lines = out.stdout.decode("utf-8", errors="replace").splitlines()
    return {l.strip() for l in lines[1:] if l.strip() and " " not in l.strip()}


def ocr_images(images):
    try:
        import pytesseract
    except ImportError:
        raise RuntimeError("OCR needs 'pip install pytesseract'.")
    if not shutil.which("tesseract"):
        for p in TESSERACT_PATHS:
            if os.path.exists(p):
                pytesseract.pytesseract.tesseract_cmd = p
                break
    try:
        langs = tesseract_langs(pytesseract.pytesseract.tesseract_cmd)
    except Exception:
        raise RuntimeError("The request letter is scanned (an image). To read it, install the free Tesseract OCR: "
                           "https://github.com/UB-Mannheim/tesseract/wiki  (select the 'Arabic' language during setup). "
                           "Or use the Word (.docx) version of the letter.")
    arabic = "ara" if "ara" in langs else next(
        (l.replace("\\", "/") for l in sorted(langs) if l.replace("\\", "/").lower().endswith("arabic")), None)
    if arabic:
        lang = "eng+" + arabic
        print(f"      OCR languages: {lang}")
    else:
        lang = "eng"
        msg = ("Tesseract has no Arabic language installed: the Kurdish parts (department, Kurdish names, title label) "
               "cannot be read. Run the Tesseract installer again and tick 'Arabic' under 'Additional language data'.")
        print("      ! " + msg)
        OCR_NOTES.append(msg)
    return "\n".join(pytesseract.image_to_string(im, lang=lang, config="--psm 4") for im in images)


def free_text(path):
    """Free mode: (text, source). Word -> direct; PDF with text -> direct; scanned PDF/image -> OCR."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        return docx_text(path), "docx"
    if ext == ".doc":
        raise RuntimeError("The old .doc format is not supported. Open it in Word and use 'Save As > .docx'.")
    if ext == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(path)
        pages = [pdf[i] for i in range(min(len(pdf), 3))]

        def page_text(pg):
            tp = pg.get_textpage()
            return tp.get_text_bounded() if hasattr(tp, "get_text_bounded") else tp.get_text_range()
        text = "\n".join(page_text(p) for p in pages)
        if len(re.sub(r"\s+", "", text)) > 80:
            return text, "pdf"
        print("      PDF is scanned, reading it with OCR (this may take a moment)...")
        return ocr_images([p.render(scale=300 / 72).to_pil() for p in pages]), "ocr"
    if ext in (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".bmp"):
        from PIL import Image, ImageOps
        print("      reading the image with OCR...")
        return ocr_images([ImageOps.exif_transpose(Image.open(path)).convert("RGB")]), "ocr"
    raise RuntimeError(f"Unsupported file type: {ext}")


def read_letter(client, path) -> dict:
    resp = client.messages.parse(
        model=MODEL, max_tokens=16000,
        output_config={"effort": "high"},
        messages=[{"role": "user", "content": file_blocks(path) + [{"type": "text", "text": READ_PROMPT}]}],
        output_format=Letter,
    )
    if resp.stop_reason == "refusal" or resp.parsed_output is None:
        raise RuntimeError(f"Could not read the request letter (stop_reason={resp.stop_reason})")
    return resp.parsed_output.model_dump()


# ─────────────────────────── 2) Online verification (AI mode) ───────────────────────────
class PortalLink(BaseModel):
    author: str
    url: Optional[str] = None
    paper_listed: Optional[bool] = Field(None, description="Does this paper appear on that staff page")


class Verify(BaseModel):
    published: bool
    doi: Optional[str] = None
    article_url: Optional[str] = Field(None, description="Article link on the journal's website")
    published_date: Optional[str] = None
    authors_on_paper: List[str] = Field(default_factory=list)
    authors_match: Optional[bool] = None
    uoz_affiliation: Optional[bool] = Field(None, description="Does University of Zakho appear in the authors' affiliations")
    clarivate_listed: Optional[bool] = None
    impact_factor: Optional[float] = None
    impact_factor_year: Optional[str] = None
    impact_factor_source: Optional[str] = None
    scopus_listed: Optional[bool] = None
    scopus_quartile: Optional[int] = None
    doaj_listed: Optional[bool] = None
    staff_portal: List[PortalLink] = Field(default_factory=list)
    notes: List[str] = Field(default_factory=list, description="Anything doubtful or unverifiable")


def research_prompt(L):
    authors = "; ".join(f"{a['name_en']} ({a['name_ku']})" for a in L["authors"])
    return f"""Verify a publication award application at the University of Zakho. Search the web and open pages.

Paper: {L['title']}
Journal: {L['journal']}
Authors: {authors}
Stated in the letter: IF={L.get('stated_impact_factor')} ; index={L.get('stated_index')}

Find and report:
1. Is the paper really published? Its DOI and the article link on the journal's website (the publisher page, not doi.org).
2. The author list and affiliations on the paper; do they match the letter, is University of Zakho among the affiliations.
3. The journal's Clarivate (Web of Science/JCR) Impact Factor and year; is it in Scopus, quartile (Q1-Q4); is it in DOAJ.
   Give the source. If you are not sure, do not guess, say "could not verify".
4. Each author's University of Zakho staff portal page (staffportal.uoz.edu.krd/...) and whether this paper is listed there.
Write briefly, as bullet points, with links."""


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
        raise RuntimeError("The search was declined (refusal).")
    return "\n".join(b.text for b in resp.content if b.type == "text")


def structure_research(client, L, text) -> dict:
    resp = client.messages.parse(
        model=MODEL, max_tokens=16000,
        output_config={"effort": "medium"},
        messages=[{"role": "user", "content":
                   "Put the research report below into the fields. Do not invent anything that is not in the report; "
                   "leave unknown fields null.\n\nAuthors (letter): "
                   + ", ".join(a["name_en"] or a["name_ku"] for a in L["authors"])
                   + "\n\nREPORT:\n" + text}],
        output_format=Verify,
    )
    if resp.parsed_output is None:
        raise RuntimeError("Could not structure the research result.")
    out = resp.parsed_output.model_dump()
    out["raw_report"] = text
    return out


# ─────────────────────────── 3) Filling the Word letter ───────────────────────────
def _ptext(p):
    return "".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p))


def _set_para(p, new):
    """Put the whole new text into the paragraph's first text run and empty the others (keeps formatting)."""
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
    RLM, LRM = "\u200f", "\u200e"
    if len(names) > 1:
        names_txt = "بەرێزان " + " و ".join(f"{RLM}({n}){RLM}" for n in names)
    else:
        names_txt = f"بەرێز {RLM}({names[0]}){RLM}"
    col = COLLEGES.get(L.get("college_key"), COLLEGES["Science"])
    dept = L["department_ku"]
    body = (f"هيڤيدارين رازيبن لسەر خەلاتكرنا دەستگێريا ڤەكۆلينا هاتيیە بەلافكرن ژ لايێ {names_txt}"
            f" دڤەكولينێدا ل پشكا {dept} ل كوليژا {col[0]} ل زانکۆیا زاخۆ ب كوژمێ({A['amount']:,}) دينار"
            f" لديڤ برگا  (1)خالا ({A['point']}) ژ فەرمانا وەزاری و توماركرى ل پشكا زانستى ل ژێر ناڤ و نيشان :   ")
    title = re.sub(r"\s+", " ", L["title"]).strip()
    ky = today.year + 700        # Kurdish calendar year

    # Each rule finds one paragraph of template.docx (by its text) and replaces it.
    rules = [
        (lambda t: t.startswith("هيڤيدارين رازيبن"), lambda p: _set_para(p, body)),
        (lambda t: "Novel Conjugate Gradient" in t,
         lambda p: _set_para(p, f"{LRM}({title}){LRM}")),
        (lambda t: t.strip() == "(Iraqi Journal of Science)",
         lambda p: _set_para(p, f"{LRM}({L['journal']}){LRM}")),
        (lambda t: "بو زانین هاتیە" in t and "index" in t,
         lambda p: _set_para(p, f"   بو زانین هاتیە index کرن {A['index_text']} ")),
        (lambda t: t.startswith("رێكه") and "زایینى" in t,
         lambda p: _set_para(p, f"رێكه\u200cوت: {today.day:02d} /  {today.month:02d}/ {today.year} ى  زایینى")),
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
        raise RuntimeError(f"template.docx does not contain the expected paragraphs (missing rules: {missing}). "
                           "If you changed the template, update the rules in fill_word() in award.py.")
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = xml.encode("utf-8") if item.filename == "word/document.xml" else zin.read(item)
            zout.writestr(item, data)


# ─────────────────────────── 4) Checks and report ───────────────────────────
def is_free(V):
    return V.get("mode") in ("free", "ucretsiz")


def checks(d):
    L, V, A = d["letter"], d.get("verify") or {}, d["award"]
    warn = []
    col = COLLEGES.get(L.get("college_key"))
    if not col:
        warn.append(f"No template for college '{L.get('college_name')}'; the College of Science template was used. "
                    "Correct the header and the dean's name by hand.")
    kind = (col or COLLEGES["Science"])[3]
    if len(L["authors"]) > MAX_AUTHORS[kind]:
        warn.append(f"{len(L['authors'])} authors; the rules award at most {MAX_AUTHORS[kind]} authors.")
    hw = L.get("handwritten_amount")
    if A["kind"] == "Iraq":
        warn.append("The journal has no IF/Scopus and is not in the approved Kurdistan journal list. 200,000 only applies "
                    "to journals of Iraqi public universities or universities outside Iraq; check who publishes the journal.")
    if hw and hw != A["amount"]:
        warn.append(f"The handwritten amount ({hw:,}) differs from the calculated amount ({A['amount']:,})!")
    sif, vif = L.get("stated_impact_factor"), V.get("impact_factor")
    if sif and vif and abs(float(sif) - float(vif)) > 0.05:
        warn.append(f"The IF in the letter ({sif}) differs from the IF found online ({vif}). "
                    f"The amount was calculated with {vif}. Check in Clarivate JCR.")
    if V:
        if not V.get("published"):
            warn.append("The paper was NOT FOUND online as published.")
        if V.get("authors_match") is False:
            warn.append("The authors on the paper do NOT MATCH the letter. Authors on the paper: "
                        + ", ".join(V.get("authors_on_paper", [])))
        if V.get("uoz_affiliation") is False:
            warn.append("University of Zakho does not appear in the paper's affiliations (rule 3).")
        for s in V.get("staff_portal", []):
            if s.get("candidates"):
                warn.append(f"{s['author']}: no exact match on the staff portal, but similar people exist; "
                            "check the links in the report to find the right one.")
            elif s.get("no_english_name"):
                warn.append(f"{s['author']}: English name unknown, so the staff portal could not be searched; check by hand.")
            elif not s.get("url"):
                warn.append(f"{s['author']}: staff portal page not found.")
            elif s.get("paper_listed") is False:
                warn.append(f"{s['author']}: the paper is NOT listed on the staff portal (rule 5-b).")
        if is_free(V):
            reg = V.get("registry")
            if reg is None and "UoZ registry could not be downloaded" not in " ".join(V.get("notes", [])):
                warn.append("The paper was NOT FOUND in the UoZ publication registry (Google Sheet) (rule 5-a).")
            elif reg and reg.get("impact_factor") and sif:
                try:
                    if abs(float(reg["impact_factor"]) - float(sif)) > 0.05:
                        warn.append(f"The registry says IF={reg['impact_factor']}, the letter says IF={sif}.")
                except ValueError:
                    pass
            if (A["kind"] != "Clarivate" and reg and reg.get("impact_factor")
                    and re.search(r"thomson|clarivate|web of science|wos|isi", reg.get("indexing", ""), re.I)):
                warn.append(f"The letter gives no IF, but the UoZ registry lists IF={reg['impact_factor']} "
                            f"({reg.get('indexing')}) for this paper. If the journal is in Clarivate, the award should "
                            "follow the IF table; check it and put stated_impact_factor in data.json.")
            if A["kind"] == "Clarivate":
                warn.append(f"The IF ({sif}) was taken from the letter; free mode cannot verify it officially. "
                            "Check it in Clarivate MJL/JCR (link in the report).")
            sjr = (V.get("sjr") or {}).get("quartile")
            if A["kind"] == "Scopus" and sjr and L.get("stated_scopus_quartile") and sjr != L["stated_scopus_quartile"]:
                warn.append(f"The letter says Scopus Q{L['stated_scopus_quartile']}, SCImago says Q{sjr}. Check in Scopus.")
        warn += V.get("notes", [])
    else:
        warn.append("No online verification was done.")
    no_ku = [a["name_en"] or a["name_ku"] for a in L["authors"] if not re.search(r"[\u0600-\u06FF]", a.get("name_ku", ""))]
    if no_ku:
        warn.append("Author(s) with unknown Kurdish name were written in English in the Word letter: " + ", ".join(no_ku)
                    + ". Write the Kurdish name in name_ku in data.json and rebuild with --json (the program learns these names).")
    for u in L.get("unclear", []):
        if not u.startswith(("Kurdish name unknown", "Kürtçe adı bilinmeyen", "Bazı yazarların Kürtçe adı yok")):
            warn.append("Reading doubt: " + u)
    return warn


def write_report(d, path, warnings):
    L, V, A = d["letter"], d.get("verify") or {}, d["award"]
    lines = [
        "PUBLICATION AWARD VERIFICATION REPORT",
        f"Date: {datetime.date.today():%Y-%m-%d}", "",
        "Authors   : " + " ; ".join(f"{a['name_ku']} ({a['name_en']})" for a in L["authors"]),
        "Paper     : " + L["title"],
        "Journal   : " + L["journal"],
        "Applicant : " + (L.get("applicant_ku") or "-")
        + ("   (authors taken from the paper: University of Zakho authors only)" if L.get("authors_from_paper") else ""),
        "DOI in letter: " + (", ".join(L.get("stated_dois") or []) or "-"),
        f"Department: {L['department_ku']} / {L.get('college_name', '')}", "",
        f"Category  : {A['kind']}  -> {A['index_text']}",
        f"Clause    : برگا (1) خالا ({A['point']})",
        f"AWARD     : {A['amount']:,} dinar"
        + (f"   (handwritten on the letter: {L['handwritten_amount']:,})" if L.get("handwritten_amount") else ""),
        "", "── LINKS ──",
        "Article on journal site: " + (V.get("article_url") or "-"),
        "DOI                    : " + (f"https://doi.org/{V['doi']}" if V.get("doi") else "-"),
    ]
    for s in V.get("staff_portal", []):
        if s.get("paper_listed"):
            lines.append(f"Staff portal : {s['author']}: paper LISTED -> {s.get('paper_url') or s.get('journals_url')}")
        elif s.get("url"):
            lines.append(f"Staff portal : {s['author']}: paper NOT LISTED -> publication list: "
                         f"{s.get('journals_url') or s['url']}")
        else:
            lines.append(f"Staff portal : {s['author']}: page not found")
        for c in s.get("candidates", []):
            lines.append(f"               possible: {c}")
    lines += ["", "── VERIFICATION ──",
              f"Published         : {V.get('published')}",
              f"Publication date  : {V.get('published_date')}",
              f"Authors (paper)   : {', '.join(V.get('authors_on_paper', [])) or '-'}",
              f"Authors match     : {V.get('authors_match')}",
              f"UoZ affiliation   : {V.get('uoz_affiliation')}",
              f"Clarivate         : {V.get('clarivate_listed')}  IF={V.get('impact_factor')} "
              f"({V.get('impact_factor_year') or ''}) source: {V.get('impact_factor_source') or '-'}",
              f"Scopus            : {V.get('scopus_listed')}  Q{V.get('scopus_quartile') or '-'}",
              f"DOAJ              : {V.get('doaj_listed')}"]
    if V.get("issn"):
        lines.append(f"ISSN              : {', '.join(V['issn'])}")
    if V.get("openalex_affiliations"):
        lines.append("Affiliations      : " + " | ".join(V["openalex_affiliations"]))
    if V.get("sjr"):
        sj = V["sjr"]
        lines.append(f"SCImago           : Q{sj.get('quartile') or '-'} ({sj.get('file')})")
    if is_free(V):
        reg = V.get("registry")
        lines.append("UoZ registry      : " + (f"REGISTERED ({reg['timestamp']}) index={reg['indexing']} IF={reg['impact_factor']}"
                                                if reg else "not found"))
        if V.get("registry_journal_ifs"):
            lines.append("  IFs given in the registry for the same journal: "
                         + ", ".join(f"{i} ({dt})" for dt, i in V["registry_journal_ifs"]))
    if V.get("links"):
        lines += ["", "── LINKS FOR MANUAL CHECKS ──"] + [f"{k}: {u}" for k, u in V["links"].items()]
    lines += ["", "── WARNINGS (check by hand) ──"]
    lines += [f"! {w}" for w in warnings] or ["(none)"]
    if V.get("raw_report"):
        lines += ["", "── RESEARCH NOTES ──", V["raw_report"]]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ─────────────────────────── Main flow ───────────────────────────
def unique_folder(folder_name, title):
    """Do not overwrite another paper by the same authors: add '(2)', '(3)'."""
    n = 1
    while True:
        cand = os.path.join(OUT_ROOT, folder_name + (f" ({n})" if n > 1 else ""))
        dj = next((os.path.join(cand, f) for f in ("data.json", "veri.json") if os.path.exists(os.path.join(cand, f))), None)
        if not dj:
            return cand
        try:
            with open(dj, encoding="utf-8") as f:
                if json.load(f)["letter"]["title"] == title:
                    return cand          # same paper: update it
        except Exception:
            pass
        n += 1


def build_outputs(d, src_file=None, today=None):
    today = today or datetime.date.today()
    d["award"] = decide_award(d)
    try:
        import free_mode
        free_mode.learn_names(d["letter"]["authors"])
    except Exception:
        pass
    names = [a["name_ku"] or a["name_en"] for a in d["letter"]["authors"]]
    folder_name = safe_name(" و ".join(f"({n})" for n in names)) or safe_name(d["letter"]["title"])[:60] or "unnamed"
    folder = unique_folder(folder_name, d["letter"]["title"])
    os.makedirs(folder, exist_ok=True)
    docx_path = os.path.join(folder, f"{folder_name} - خەلاتکرن.docx")
    fill_word(d, docx_path, today)
    warnings = checks(d)
    d["warnings"] = warnings
    write_report(d, os.path.join(folder, "report.txt"), warnings)
    with open(os.path.join(folder, "data.json"), "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
    if src_file and os.path.abspath(os.path.dirname(src_file)) != os.path.abspath(folder):
        shutil.copy2(src_file, os.path.join(folder, "request" + os.path.splitext(src_file)[1].lower()))
    return folder, docx_path, warnings


def process_file(client, path, use_web=True):
    """client=None means free mode (no AI)."""
    print(f"\n=== {os.path.basename(path)}")
    print("  1/3 reading the request letter...")
    if client is None:
        import free_mode
        text, source = free_text(path)
        L = free_mode.parse_letter(text)
        L["source"] = source
        if source == "ocr":
            L["unclear"] += OCR_NOTES[-1:]
            L["unclear"].append("The letter is scanned (an image) and was read with OCR. "
                                "Check the title, journal and names in the Word letter carefully.")
    else:
        L = read_letter(client, path)
    print("      authors:", ", ".join(a["name_ku"] for a in L["authors"])
          or "(no author list in the letter; University of Zakho authors will be taken from the paper)")
    print("      paper  :", L["title"])
    d = {"letter": L, "verify": None}
    if use_web:
        print("  2/3 searching for the paper online...")
        try:
            if client is None:
                d["verify"] = free_mode.verify(L)
                d["verify"]["mode"] = "free"
            else:
                d["verify"] = structure_research(client, L, research(client, L))
                d["verify"]["mode"] = "ai"
        except Exception as e:  # even if the search fails, still produce the Word letter
            print("      ! search failed:", e)
    V = d.get("verify") or {}
    if not L["authors"] and V.get("authors_on_paper"):
        L["authors"] = [{"name_ku": n, "name_en": n} for n in V["authors_on_paper"]]
        L.setdefault("unclear", []).append(
            "Author names could not be read from the letter; they were taken from the paper (Crossref) "
            "and written in English. Put the Kurdish names in data.json and rebuild with --json.")
    print("  3/3 preparing the Word letter...")
    return build_outputs(d, path)


INPUT_EXT = (".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff", ".pdf", ".docx", ".doc")


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    print(f"Program version: {VERSION}")
    ap = argparse.ArgumentParser(description="University of Zakho publication award letter generator")
    ap.add_argument("files", nargs="*", help="request letter(s): Word .docx / PDF / image")
    ap.add_argument("--json", help="rebuild the Word letter from a data.json file (no internet needed)")
    ap.add_argument("--no-web", action="store_true", help="skip online verification")
    ap.add_argument("--free", "--ucretsiz", dest="free", action="store_true",
                    help="use free mode even if an API key is configured")
    ap.add_argument("--ai", action="store_true", help="force AI mode (needs an API key)")
    a = ap.parse_args()

    results = []
    if a.json:
        with open(a.json, encoding="utf-8") as f:
            res = build_outputs(json.load(f))
        results.append(res)
        old = os.path.dirname(os.path.abspath(a.json))
        new = os.path.abspath(res[0])
        # if the names changed, the folder name changes too: move the request copy, remove the old folder
        if old != new and os.path.dirname(old) == os.path.abspath(OUT_ROOT):
            for fn in os.listdir(old):
                if fn.startswith(("request", "dilekce")) and not os.path.exists(os.path.join(new, fn)):
                    shutil.move(os.path.join(old, fn), os.path.join(new, fn))
            shutil.rmtree(old, ignore_errors=True)
            print("Old folder moved to the new name:", os.path.basename(new))
    else:
        files = a.files
        if not files:
            os.makedirs(IN_DIR, exist_ok=True)
            files = [os.path.join(IN_DIR, x) for x in sorted(os.listdir(IN_DIR))
                     if x.lower().endswith(INPUT_EXT) and not x.startswith("~$")]
            if not files:
                print("No files given. Put the request letters into the 'input' folder, "
                      "or drag them onto run.bat.")
                return 1
        use_ai = a.ai or (bool(api_key()) and not a.free)
        client = api_client() if use_ai else None
        print("Mode:", "AI (with API key)" if use_ai else "FREE (no API key)")
        for p in files:
            if not os.path.isfile(p):
                print(f"\n=== {os.path.basename(p)}\n  ERROR: file not found: {p}\n"
                      "  (Tip: type the command, then drag the file into the cmd window; the path is filled in.)")
                continue
            try:
                results.append(process_file(client, p, use_web=not a.no_web))
            except Exception as e:
                print(f"  ERROR ({os.path.basename(p)}): {e.__class__.__name__}: {e}")

    for folder, docx_path, warnings in results:
        print("\n✔", docx_path)
        print_links(folder)
        for w in warnings:
            print("   !", w)
    if results:
        out = write_summary()
        print("\nSummary page (all papers, clickable links):", out)
    return 0


def _load_data(folder):
    for fn in ("data.json", "veri.json"):
        p = os.path.join(folder, fn)
        if os.path.isfile(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f), fn
    return None, None


def print_links(folder):
    try:
        d, _ = _load_data(folder)
    except Exception:
        return
    if not d:
        return
    V = d.get("verify") or {}
    print(f"   Journal: {d['letter'].get('journal') or '-'}   |   Award: {d['award']['amount']:,} dinar")
    print(f"   Article: {V.get('article_url') or (V.get('letter_link', '') + '  (link from the letter)' if V.get('letter_link') else '-')}")
    if V.get("doi"):
        print(f"   DOI    : https://doi.org/{V['doi']}")
    for s in V.get("staff_portal", []):
        u = s.get("paper_url") or s.get("journals_url") or s.get("url") or "not found"
        st = {True: "listed", False: "NOT LISTED"}.get(s.get("paper_listed"), "?")
        print(f"   Staff  : {s['author']} ({st}) {u}")
        for c in s.get("candidates", []):
            print(f"            possible: {c}")


def write_summary():
    """output/summary.html: every paper in the output folder in one table, with clickable links."""
    import html as H
    rows = []
    for name in sorted(os.listdir(OUT_ROOT)):
        folder = os.path.join(OUT_ROOT, name)
        if not os.path.isdir(folder):
            continue
        try:
            d, _ = _load_data(folder)
        except Exception:
            continue
        if not d:
            continue
        report = "report.txt" if os.path.exists(os.path.join(folder, "report.txt")) else "rapor.txt"
        L, V, A = d["letter"], d.get("verify") or {}, d.get("award") or {}
        a = lambda url, txt: f'<a href="{H.escape(url)}" target="_blank">{H.escape(txt)}</a>' if url else "-"
        staff = []
        for s in V.get("staff_portal", []):
            st = {True: '<b class="ok">✓ listed</b>', False: '<b class="no">✗ not listed</b>'}.get(
                s.get("paper_listed"), '<b class="q">?</b>')
            link = s.get("paper_url") or s.get("journals_url") or s.get("url")
            cand = "".join(f'<div class="small">possible: {a(c.split(": ", 1)[-1], c.split(": ", 1)[0])}</div>'
                           for c in s.get("candidates", []))
            staff.append(f"{a(link, s['author']) if link else H.escape(s['author']) + ' (no page)'} {st}{cand}")
        reg = V.get("registry")
        regtxt = '<b class="ok">✓</b>' if reg else ('<b class="no">✗</b>' if is_free(V) else "-")
        pub = '<b class="ok">✓</b>' if V.get("published") else '<b class="no">✗ not found</b>'
        art = (a(V.get("article_url"), "article on journal site") if V.get("article_url")
               else a(V.get("letter_link"), "link from the letter"))
        warns = d.get("warnings", [])
        rows.append(f"""<tr>
<td class="rtl">{"<br>".join(H.escape(x.get("name_ku") or x.get("name_en") or "") for x in L.get("authors", [])) or "-"}</td>
<td>{H.escape(L.get("title", ""))}<div class="small">{H.escape(L.get("journal") or "")}</div></td>
<td>{H.escape(A.get("index_text", ""))}<div class="money">{A.get("amount", 0):,}</div></td>
<td>{pub}<br>{art}<br>{a("https://doi.org/" + V["doi"] if V.get("doi") else None, "DOI")}</td>
<td>{"<br>".join(staff) or "-"}</td>
<td>{regtxt}</td>
<td><a href="{H.escape(name)}/{report}">report</a> · <a href="{H.escape(name)}">folder</a>
<div class="small">{len(warns)} warning(s)</div></td></tr>""")
    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Publication Awards – Summary</title>
<style>body{{font-family:Segoe UI,Arial,sans-serif;margin:20px;color:#222}}
table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #ccc;padding:6px 8px;vertical-align:top;font-size:14px}}
th{{background:#f0f0f0;text-align:left}}.rtl{{direction:rtl;text-align:right;font-family:Tahoma,Arial}}
.small{{color:#666;font-size:12px;margin-top:3px}}.money{{font-weight:bold;margin-top:4px}}
.ok{{color:#11823b}}.no{{color:#c62828}}.q{{color:#999}}a{{color:#0b57d0}}</style></head><body>
<h2>Publication Awards – Summary ({len(rows)} papers)</h2>
<p class="small">Generated: {datetime.datetime.now():%Y-%m-%d %H:%M} · Staff portal links jump to the paper
(Chrome/Edge highlight the title). Check every ✗ by hand; details are in each row's report.</p>
<table><tr><th>Authors (in the Word letter)</th><th>Paper / Journal</th><th>Index / Award (dinar)</th><th>Published</th>
<th>Staff portal (rule 5-b)</th><th>UoZ registry (rule 5-a)</th><th>Report</th></tr>
{"".join(rows)}</table></body></html>"""
    out = os.path.join(OUT_ROOT, "summary.html")
    with open(out, "w", encoding="utf-8") as f:
        f.write(page)
    return out


if __name__ == "__main__":
    sys.exit(main())
