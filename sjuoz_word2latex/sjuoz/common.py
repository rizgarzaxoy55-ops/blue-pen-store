# -*- coding: utf-8 -*-
"""Ortak yardımcılar: XML ad alanları, uyarılar, LaTeX kaçışı, bağlantılar, sağdan sola metin."""
import os
import re
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
    "\u200b": "",
    "\u00ad": "",
    "\ufeff": "",
}


def tex_escape(s):
    return "".join(_TEX_SPECIAL.get(c, c) for c in s)


def url_escape(u):
    """\\url/\\href içinde sorun çıkaran karakterleri kaçırır."""
    return u.replace("\\", "/").replace("%", r"\%").replace("#", r"\#").replace("{", "").replace("}", "")


URL_RE = re.compile(r"(?:https?://|www\.)[^\s<>\"]+", re.I)
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
DOI_DUP_RE = re.compile(r"https?://(?:dx\.)?doi\.org/(?:https?://(?:dx\.)?doi\.org/)+", re.I)
ETAL_RE = re.compile(r"\bet\s+al\.")


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
    for mt in pattern.finditer(text):
        out.append(_bidi(text[pos:mt.start()], italic))
        tok = mt.group(0)
        if "@" in tok and not tok.lower().startswith(("http", "www")):
            out.append(r"\href{mailto:%s}{%s}" % (tok, tex_escape(tok)))
        else:
            u, trail = clean_url(tok)
            href = u if u.lower().startswith("http") else "https://" + u
            if ref_mode and "doi.org/" in u.lower():
                # "DOI:" etiketi metinde zaten yoksa ekle
                prev = "".join(out).rstrip()
                label = "" if prev.lower().endswith("doi:") else "DOI:"
                out.append(r"%s\href{%s}{%s}" % (label, url_escape(href), url_escape(u)))
            else:
                out.append(r"\url{%s}" % url_escape(u))
            out.append(tex_escape(trail))
        pos = mt.end()
    out.append(_bidi(text[pos:], italic))
    return "".join(out)


def _etal(s, italic):
    return s if italic else ETAL_RE.sub(r"\\textit{et al.}", s)


SMALL_WORDS = {"a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "into", "of",
               "on", "or", "the", "to", "with", "via", "vs", "per"}
KEEP_UPPER = {"IEEE", "ACM", "IET", "SPIE", "IOP", "MDPI", "IJ", "AI", "IOT", "IT", "USA", "UK",
              "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "PLOS", "BMC", "NASA"}


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
