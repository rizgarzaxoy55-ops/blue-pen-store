# -*- coding: utf-8 -*-
"""Ortak yardımcılar: XML ad alanları, uyarılar, LaTeX kaçışı, bağlantılar, sağdan sola metin."""
import os
import re
import unicodedata
import sys


# --------------------------------------------------------------------------
# XML ad alanları
# --------------------------------------------------------------------------
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]
R = "{%s}" % NS["r"]
MC = "{%s}" % NS["mc"]

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")  # logolar, şemalar

# Windows'ta alt süreçler (xelatex, LibreOffice, pdftotext) konsol penceresi açmasın
NO_WINDOW = {"creationflags": 0x08000000} if os.name == "nt" else {}  # CREATE_NO_WINDOW

WARNINGS = []


def warn(msg):
    WARNINGS.append(msg)
    print("UYARI: " + msg, file=sys.stderr)


def safe_console():
    """Konsolun kodlamasında olmayan karakterler (Windows) programı çökertmesin."""
    for stream in (sys.stdout, sys.stderr):
        if stream is not None and hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(errors="replace")
            except (ValueError, OSError):
                pass


def local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag


def wval(el, default=None):
    if el is None:
        return default
    v = el.get(W + "val")
    return default if v is None else v


def on_off(el):
    """w:b, w:i gibi aç/kapa özelliklerini okur."""
    if el is None:
        return False
    return wval(el, "true").lower() not in ("0", "false", "off", "none")


# --------------------------------------------------------------------------
# LaTeX kaçış (escape) yardımcıları
# --------------------------------------------------------------------------
_TEX_SPECIAL = {
    "\\": r"\textbackslash{}",
    "&": r"\&",
    "%": r"\%",
    "$": r"\$",
    "#": r"\#",
    "_": r"\_",
    "{": r"\{",
    "}": r"\}",
    "~": r"\textasciitilde{}",
    "^": r"\textasciicircum{}",
    "\u00a0": "~",
    # üst/alt simge işaretleri çoğu yazı tipinde yok ("mg g⁻¹"): matematik kipinde yazılır
    "\u207a": r"\textsuperscript{+}", "\u207b": r"$^{-}$", "\u207c": r"$^{=}$",
    "\u208a": r"\textsubscript{+}", "\u208b": r"$_{-}$",
    "\u200b": "",
    "\u00ad": "",
    "\ufeff": "",
}


def _script_letter(c):
    """Unicode üst/alt simge harfi ("ᶜ", "ᴱ", "ⁱ", "ᵢ", "ᵧ") → ("^" | "_", taban harf); değilse None.
    Bu harflerin çoğu Times New Roman'da yoktur; rakamlar (², ₀) yazı tipinde bulunduğu için dokunulmaz."""
    if c.isascii() or c in _TEX_SPECIAL:
        return None
    dec = unicodedata.decomposition(c)
    if not dec.startswith(("<super>", "<sub>")):
        return None
    base = unicodedata.normalize("NFKC", c)
    if not base.isalpha():
        return None
    return ("^" if dec.startswith("<super>") else "_"), base


def tex_escape(s):
    if s.isascii():
        return "".join(_TEX_SPECIAL.get(c, c) for c in s)
    out, run, kind = [], [], None

    def flush():
        if run:
            out.append(r"\text%s{%s}" % ("superscript" if kind == "^" else "subscript", "".join(run)))
            run.clear()
    for c in s:
        sl = _script_letter(c)
        if sl and sl[0] == kind:
            run.append(sl[1])
            continue
        flush()
        if sl:
            kind = sl[0]
            run.append(sl[1])
        else:
            kind = None
            out.append(_TEX_SPECIAL.get(c, c))
    flush()
    return "".join(out)


def url_escape(u):
    """\\url/\\href içinde sorun çıkaran karakterleri kaçırır."""
    return u.replace("\\", "/").replace("%", r"\%").replace("#", r"\#").replace("{", "").replace("}", "")


URL_RE = re.compile(r"(?:https?://|www\.)[^\s<>\"]+", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
DOI_DUP_RE = re.compile(r"https?://(?:dx\.)?doi\.org/(?:https?://(?:dx\.)?doi\.org/)+", re.I)
ETAL_RE = re.compile(r"\bet\s+al\.")


def href(link, shown):
    """Görünen metni de adres olan bağlantı: "_" gibi karakterler metin kipinde hata vermesin."""
    return r"\href{%s}{{\urlstyle{same}\nolinkurl{%s}}}" % (url_escape(link), url_escape(shown))


def clean_url(u):
    """URL sonundaki noktalama işaretlerini ayırır (parantez dengesini korur)."""
    trail = ""
    while u and u[-1] in ".,;:)]'\u201d\u2019":
        if u[-1] == ")" and u.count("(") >= u.count(")"):
            break
        trail = u[-1] + trail
        u = u[:-1]
    return u, trail


# Sağdan sola yazılar (Arapça, Kürtçe Sorani, Farsça)
RTL_CHARS = "؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿"
RTL_RUN_RE = re.compile("[%s](?:[%s\\s‌‍0-9٠-٩.,:;!?()\\-]*[%s])?" % ((RTL_CHARS,) * 3))
BIDI = {"used": False, "par": False}  # belgede RTL metin var mı / şu an RTL paragraf mı yazılıyor


def is_rtl_text(text):
    letters = [c for c in text if c.isalpha()]
    rtl = sum(1 for c in letters if re.match("[%s]" % RTL_CHARS, c))
    return bool(letters) and rtl / len(letters) > 0.5


def _bidi(text, italic):
    """Metni kaçırır; Arap harfli parçaları sağdan sola (\\rtl), RTL paragraftaki
    Latin parçaları soldan sağa (\\ltr) yazdırır."""
    if not RTL_RUN_RE.search(text) and not BIDI["par"]:
        return _etal(tex_escape(text), italic)
    out, pos = [], 0
    for mt in RTL_RUN_RE.finditer(text):
        out.append(_ltr(text[pos:mt.start()], italic))
        if BIDI["par"]:
            out.append(tex_escape(mt.group(0)))
        else:
            out.append(r"\rtl{%s}" % tex_escape(mt.group(0)))
        pos = mt.end()
        BIDI["used"] = True
    out.append(_ltr(text[pos:], italic))
    return "".join(out)


def _ltr(text, italic):
    s = _etal(tex_escape(text), italic)
    if BIDI["par"] and re.search(r"[A-Za-z0-9]", text):
        # sondaki noktalama sağdan sola akışta kalmalı ("2024." -> ".2024" değil)
        mt = re.match(r"^(\s*)(.*?)([\s.,;:!?)\]]*)$", s, re.S)
        lead, core, trail = mt.group(1), mt.group(2), mt.group(3)
        return lead + (r"\ltr{%s}" % core if core else "") + trail
    return s


def render_plain(text, italic=False, ref_mode=False, link=False):
    """Düz metni LaTeX'e çevirir: URL/e-posta bağlantıları, 'et al.' italik."""
    if link:
        return _bidi(text, italic)
    out = []
    pos = 0
    pattern = re.compile(URL_RE.pattern + "|" + EMAIL_RE.pattern, re.I)
    # dergi aralıkları yalnız kaynakçada en dash yapar; gövdede negatif sayı eksi işaretidir ("r = −0.76")
    dash = en_dash_ranges if ref_mode else (lambda x: NEG_NUMBER.sub("−", x) if "-" in x else x)
    for mt in pattern.finditer(text):
        out.append(_bidi(dash(text[pos:mt.start()]), italic))
        tok = mt.group(0)
        if "@" in tok and not tok.lower().startswith(("http", "www")):
            out.append(r"\href{mailto:%s}{%s}" % (tok, tex_escape(tok)))
        else:
            u, trail = clean_url(tok)
            link = u if u.lower().startswith("http") else "https://" + u
            if ref_mode and "doi.org/" in u.lower():
                # "DOI:" etiketi metinde zaten yoksa ekle
                prev = "".join(out).rstrip()
                label = "" if prev.lower().endswith("doi:") else "DOI:"
                out.append(label + href(link, u))
            else:
                out.append(r"\url{%s}" % url_escape(u))
            out.append(tex_escape(trail))
        pos = mt.end()
    out.append(_bidi(dash(text[pos:]), italic))
    return "".join(out)


# Kaynakçada iki sayı arasındaki kısa çizgi bir aralıktır: "pp. 546-556" → en dash (dergideki gibi).
# Tarih/ISSN/ISBN zincirleri ("2024-05-12", "2663-628X") ve telefonlar ("+964-750…") dokunulmaz.
NUM_RANGE = re.compile(r"(?<![\w.,/+\-–])(\d+(?:\.\d+)?)-(\d+(?:\.\d+)?)(?![\w/\-–]|\.\d)")


NEG_NUMBER = re.compile(r"(?<=[\s(=\[])-(?=\d)")


def en_dash_ranges(s):
    return NUM_RANGE.sub("\\1–\\2", s) if "-" in s else s


def _etal(s, italic):
    return s if italic else ETAL_RE.sub(r"\\textit{et al.}", s)


SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into", "of",
               "on", "or", "the", "to", "with", "via", "vs", "per"}
KEEP_UPPER = {"IEEE", "ACM", "IET", "SPIE", "IOP", "MDPI", "IJ", "AI", "IOT", "IT", "USA", "UK",
              "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "PLOS", "BMC", "NASA"}


_TEX_ACCENT = {"'": "\u0301", "’": "\u0301", "´": "\u0301", "`": "\u0300", '"': "\u0308", "~": "\u0303",
               "^": "\u0302"}


def fix_text(s):
    """Word metnindeki bozuk karakterler: Yunan soru işareti (;), ayrı yazılmış çengel ("c¸" → "ç"),
    metne yazılmış LaTeX aksanları ("Dom\\'\\inguez" → "Domínguez")."""
    if not s or s.isascii() and "\\" not in s:
        return s
    s = re.sub(r"([cCsStT])\u00b8", lambda m: m.group(1) + "\u0327", s)
    s = re.sub(r"\\([’'´`\"~^])\s*\{?\\?([A-Za-z])\}?", lambda m: m.group(2) + _TEX_ACCENT[m.group(1)], s)
    return unicodedata.normalize("NFC", s)


def title_case(s):
    """Başlık Biçimi: TAMAMI BÜYÜK metin çevrilir; karışık metinde yalnız sözcük başları büyütülür
    ("Ethical statement" → "Ethical Statement", "SoH" olduğu gibi kalır)."""
    if s.upper() == s:
        return title_case_caps(s)
    out, first = [], True
    for w in re.split(r"(\s+)", s):
        if w.strip():
            core = re.sub(r"[^A-Za-z]", "", w)
            if first or core.lower() not in SMALL_WORDS:
                w = "-".join(x[:1].upper() + x[1:] for x in w.split("-"))
            first = False
        out.append(w)
    return "".join(out)


def journal_case(s):
    """Kaynakçadaki dergi adı Başlık Biçiminde (APA): "Journal of clinical laboratory analysis" →
    "Journal of Clinical Laboratory Analysis", "Ieee Access" → "IEEE Access", "d’endocrinologie" →
    "d’Endocrinologie". Sözcükteki diğer büyük harflere dokunulmaz."""
    out, first = [], True
    for w in re.split(r"(\s+)", s):
        if w.strip():
            core = re.sub(r"[^A-Za-z]", "", w)
            if core.upper() in KEEP_UPPER:
                w = w.replace(core, core.upper()) if core in w else w.upper()
            elif not first and core.lower() in SMALL_WORDS:
                pass
            else:
                w = re.sub(r"(^|[-’'(/])([a-zà-ÿ])", lambda m: m.group(1) + m.group(2).upper(), w)
                w = re.sub(r"^([DdLl][’'])([A-Za-zà-ÿ])", lambda m: m.group(1).lower() + m.group(2).upper(), w)
            first = False
        out.append(w)
    return "".join(out)


def title_case_caps(s):
    """TAMAMI BÜYÜK HARF dergi adlarını Başlık Biçimine çevirir."""
    words = re.split(r"(\s+)", s)
    out = []
    first = True
    for w in words:
        if not w.strip():
            out.append(w)
            continue
        core = re.sub(r"[^A-Za-z]", "", w)
        if core.upper() in KEEP_UPPER:
            out.append(w)
        elif not first and core.lower() in SMALL_WORDS:
            out.append(w.lower())
        else:
            out.append("-".join(x[:1].upper() + x[1:].lower() for x in w.split("-")))
        first = False
    return "".join(out)
