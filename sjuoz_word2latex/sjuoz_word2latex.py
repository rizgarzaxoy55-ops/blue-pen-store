#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SJUOZ Word -> LaTeX dönüştürücü
================================

Makale (manuscript) Word dosyasını ve isteğe bağlı kapak mektubu (cover letter)
Word dosyasını okuyup Science Journal of University of Zakho (SJUOZ) şablonunda
tek bir XeLaTeX dosyası üretir.

Makaleden alınanlar : başlık, yazarlar, affiliasyonlar, sorumlu yazar, özet,
                      anahtar kelimeler, bölümler, paragraflar (kalın/italik/
                      alt-üst simge), denklemler (OMML -> LaTeX), tablolar,
                      şekiller (resimler dışa aktarılır), son bölümler ve kaynaklar.
Kapak mektubundan   : yazarların ORCID, e-posta, telefon ve affiliasyon bilgileri.
Komut satırından    : cilt, sayı, yıl, ay, sayfa, makale no, tarihler.

Kullanım:
    python sjuoz_word2latex.py makale.docx kapak.docx -o cikti --first-page 571

Gereken paket: lxml   (pip install lxml)
İsteğe bağlı : Pillow (TIFF/GIF/BMP resimleri PNG'ye çevirmek için)
"""

import argparse
import datetime
import hashlib
import os
import posixpath
import re
import shutil
import sys
import unicodedata
import zipfile

try:
    from lxml import etree
except ImportError:  # pragma: no cover
    sys.exit("Hata: 'lxml' paketi gerekli. Kurulum: pip install lxml")

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
W = "{%s}" % NS["w"]
M = "{%s}" % NS["m"]
R = "{%s}" % NS["r"]
MC = "{%s}" % NS["mc"]

WARNINGS = []


def warn(msg):
    WARNINGS.append(msg)
    print("UYARI: " + msg, file=sys.stderr)


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


def render_plain(text, italic=False, ref_mode=False, link=False):
    """Düz metni LaTeX'e çevirir: URL/e-posta bağlantıları, 'et al.' italik."""
    if link:
        return _etal(tex_escape(text), italic)
    out = []
    pos = 0
    pattern = re.compile(URL_RE.pattern + "|" + EMAIL_RE.pattern, re.I)
    for mt in pattern.finditer(text):
        out.append(_etal(tex_escape(text[pos:mt.start()]), italic))
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
    out.append(_etal(tex_escape(text[pos:]), italic))
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
            out.append(w[:1].upper() + w[1:].lower())
        first = False
    return "".join(out)


# --------------------------------------------------------------------------
# OMML (Word denklemleri) -> LaTeX
# --------------------------------------------------------------------------
GREEK = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta", "ε": r"\varepsilon",
    "ϵ": r"\epsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta", "ϑ": r"\vartheta",
    "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu", "ν": r"\nu", "ξ": r"\xi",
    "π": r"\pi", "ϖ": r"\varpi", "ρ": r"\rho", "ϱ": r"\varrho", "σ": r"\sigma", "ς": r"\varsigma",
    "τ": r"\tau", "υ": r"\upsilon", "φ": r"\phi", "ϕ": r"\phi", "χ": r"\chi", "ψ": r"\psi",
    "ω": r"\omega", "Γ": r"\Gamma", "Δ": r"\Delta", "Θ": r"\Theta", "Λ": r"\Lambda",
    "Ξ": r"\Xi", "Π": r"\Pi", "Σ": r"\Sigma", "Υ": r"\Upsilon", "Φ": r"\Phi", "Ψ": r"\Psi",
    "Ω": r"\Omega", "µ": r"\mu",
}
MATH_SYMBOLS = {
    "−": "-", "∗": "*", "·": r"\cdot", "⋅": r"\cdot", "×": r"\times", "÷": r"\div",
    "±": r"\pm", "∓": r"\mp", "≤": r"\leq", "≥": r"\geq", "≠": r"\neq", "≈": r"\approx",
    "≅": r"\cong", "≡": r"\equiv", "∝": r"\propto", "∼": r"\sim", "~": r"\sim", "∞": r"\infty",
    "→": r"\rightarrow", "←": r"\leftarrow", "↔": r"\leftrightarrow", "⇒": r"\Rightarrow",
    "⇐": r"\Leftarrow", "⇔": r"\Leftrightarrow", "↦": r"\mapsto", "∈": r"\in", "∉": r"\notin",
    "∋": r"\ni", "⊂": r"\subset", "⊃": r"\supset", "⊆": r"\subseteq", "⊇": r"\supseteq",
    "∩": r"\cap", "∪": r"\cup", "∧": r"\wedge", "∨": r"\vee", "¬": r"\neg", "∀": r"\forall",
    "∃": r"\exists", "∅": r"\emptyset", "∂": r"\partial", "∇": r"\nabla", "…": r"\ldots",
    "⋯": r"\cdots", "⋮": r"\vdots", "⋱": r"\ddots", "′": "'", "″": "''", "°": r"^{\circ}",
    "∘": r"\circ", "∑": r"\sum", "∏": r"\prod", "∫": r"\int", "√": r"\surd", "‖": r"\|",
    "⟨": r"\langle", "⟩": r"\rangle", "⌈": r"\lceil", "⌉": r"\rceil", "⌊": r"\lfloor",
    "⌋": r"\rfloor", "{": r"\{", "}": r"\}", "%": r"\%", "#": r"\#", "$": r"\$", "_": r"\_",
    "&": r"\&", "\\": r"\backslash", "^": r"\hat{}", "≪": r"\ll", "≫": r"\gg", "⊕": r"\oplus",
    "⊗": r"\otimes", "†": r"\dagger", "ℓ": r"\ell", "ℝ": r"\mathbb{R}", "ℕ": r"\mathbb{N}",
    "ℤ": r"\mathbb{Z}", "∠": r"\angle", "⊥": r"\perp", "∥": r"\parallel", "∴": r"\therefore",
    "∵": r"\because", "\u2061": "", "\u2062": "", "\u2063": "", "\u200b": "",
}
FUNC_NAMES = {"sin", "cos", "tan", "cot", "sec", "csc", "sinh", "cosh", "tanh", "coth",
              "arcsin", "arccos", "arctan", "log", "ln", "lg", "exp", "max", "min", "sup",
              "inf", "lim", "det", "arg", "deg", "dim", "gcd", "hom", "ker", "Pr", "liminf",
              "limsup"}
NARY = {"∑": r"\sum", "∏": r"\prod", "∐": r"\coprod", "∫": r"\int", "∬": r"\iint",
        "∭": r"\iiint", "∮": r"\oint", "⋃": r"\bigcup", "⋂": r"\bigcap", "⋁": r"\bigvee",
        "⋀": r"\bigwedge", "⨁": r"\bigoplus", "⨂": r"\bigotimes"}
ACCENTS = {"\u0305": r"\bar", "\u00af": r"\bar", "\u0302": r"\hat", "^": r"\hat",
           "\u0303": r"\tilde", "~": r"\tilde", "\u0307": r"\dot", "\u0308": r"\ddot",
           "\u20d7": r"\vec", "\u20d6": r"\overleftarrow", "\u030c": r"\check",
           "\u0301": r"\acute", "\u0300": r"\grave", "\u0306": r"\breve", "\u2192": r"\vec",
           "\u20e1": r"\overleftrightarrow", "\u0304": r"\bar"}
DELIMS = {"(": "(", ")": ")", "[": "[", "]": "]", "{": r"\{", "}": r"\}", "|": "|",
          "‖": r"\|", "⟨": r"\langle", "⟩": r"\rangle", "〈": r"\langle", "〉": r"\rangle",
          "⌈": r"\lceil", "⌉": r"\rceil", "⌊": r"\lfloor", "⌋": r"\rfloor", "": "."}

EQNUM_MARK = "\x01"  # Word'de denklemle numara arasındaki uzun boşluk


class OMML:
    """Word OMML denklemini LaTeX'e çevirir. SEQ alan kodlarını (field) atlar."""

    def __init__(self):
        self.fld = []  # alan kodu yığını: 'instr' / 'result'
        self.seq_seen = False
        self.in_eqarr = False

    # ---- giriş noktası
    def convert(self, omath):
        self.fld = []
        self.seq_seen = False
        latex = self.seq(omath)
        latex = re.sub(r"[ ]{2,}", " ", latex).strip()
        number = None
        pat = r"(?:%s|\\#)?\s*\(\s*(\d+(?:\.\d+)*[a-z]?)\s*\)\s*%s?\s*$" % (EQNUM_MARK, EQNUM_MARK)
        mt = re.search(pat, latex)
        if mt and (self.seq_seen or EQNUM_MARK in mt.group(0) or "\\#" in mt.group(0)):
            number = mt.group(1)
            latex = latex[:mt.start()]
        latex = latex.replace(EQNUM_MARK, " ")
        latex = re.sub(r"\s+", " ", latex).strip()
        latex = re.sub(r"(\\[a-zA-Z]+) (?=[^a-zA-Z])", r"\1", latex)
        return latex, number

    # ---- yardımcılar
    def child(self, el, name):
        return el.find(M + name)

    def arg(self, el, name):
        c = self.child(el, name)
        return "" if c is None else self.seq(c)

    def prop(self, el, pr, name, default=None):
        p = el.find(M + pr)
        if p is None:
            return default
        c = p.find(M + name)
        if c is None:
            return default
        v = c.get(M + "val")
        return default if v is None else v

    def seq(self, el):
        """Bir elemanın çocuklarını sırayla çevirir; ardışık m:r metinlerini birleştirir."""
        out = []
        buf = []  # (text, plain)

        def flush():
            if buf:
                plain = buf[0][1]
                out.append(self.text(''.join(t for t, _ in buf), plain))
                buf.clear()

        for c in el:
            name = local(c.tag)
            if c.tag == M + "r" or c.tag == W + "r":
                txt, plain = self.run(c)
                if txt is None:
                    continue
                if buf and not txt.strip():  # boşluklar biçimden bağımsız
                    buf.append((txt, buf[0][1]))
                    continue
                if buf and buf[0][1] != plain:
                    flush()
                buf.append((txt, plain))
                continue
            if c.tag == MC + "AlternateContent":
                flush()
                ch = c.find(MC + "Choice")
                if ch is None:
                    ch = c.find(MC + "Fallback")
                if ch is not None:
                    out.append(self.seq(ch))
                continue
            flush()
            fn = getattr(self, "f_" + name, None)
            if fn is not None:
                out.append(fn(c))
            elif name in ("ctrlPr", "rPr", "argPr") or name.endswith("Pr"):
                continue
            elif name in ("bookmarkStart", "bookmarkEnd", "proofErr", "permStart", "permEnd"):
                continue
            elif name in ("sdt",):
                content = c.find(W + "sdtContent")
                if content is not None:
                    out.append(self.seq(content))
            else:
                out.append(self.seq(c))
        flush()
        return "".join(out)

    def run(self, r):
        """Bir matematik koşusunun (run) metnini döndürür; alan kodu içindeyse None."""
        plain = False
        rpr = r.find(M + "rPr")
        if rpr is not None:
            sty = rpr.find(M + "sty")
            if (sty is not None and sty.get(M + "val") == "p") or rpr.find(M + "nor") is not None:
                plain = True
        texts = []
        for c in r:
            if c.tag == W + "fldChar":
                t = c.get(W + "fldCharType")
                if t == "begin":
                    self.fld.append("instr")
                elif t == "separate" and self.fld:
                    self.fld[-1] = "result"
                elif t == "end" and self.fld:
                    self.fld.pop()
                continue
            if c.tag in (M + "t", W + "t"):
                if self.fld and self.fld[-1] == "instr":
                    if "SEQ" in (c.text or ""):
                        self.seq_seen = True
                    continue
                if self.fld:
                    self.seq_seen = self.seq_seen or bool(re.fullmatch(r"\s*\d+\s*", c.text or ""))
                texts.append(c.text or "")
            elif c.tag == W + "instrText":
                if "SEQ" in (c.text or ""):
                    self.seq_seen = True
            elif c.tag == W + "tab":
                texts.append("  ")
        if not texts:
            return None, plain
        return "".join(texts), plain

    def text(self, s, plain=False):
        out = []
        i = 0
        n = len(s)
        while i < n:
            ch = s[i]
            if ch.isspace():
                j = i
                while j < n and s[j].isspace():
                    j += 1
                prev_c = s[i - 1] if i > 0 else ""
                next_c = s[j] if j < n else ""
                if j - i >= 2:
                    out.append(EQNUM_MARK)
                elif prev_c.isalnum() and next_c.isalnum():
                    out.append(r"\ ")
                else:
                    out.append(" ")
                i = j
                continue
            if 0x1D400 <= ord(ch) <= 0x1D7FF:  # matematiksel italik/kalın harfler
                ch = unicodedata.normalize("NFKC", ch)
            if ch.isalpha() and ch.isascii():
                j = i
                while j < n and s[j].isascii() and s[j].isalpha():
                    j += 1
                word = s[i:j]
                if word in FUNC_NAMES:
                    out.append("\\" + word + " ")
                elif plain:
                    out.append(r"\mathrm{%s}" % word)
                elif len(word) > 1:
                    out.append(r"\mathit{%s}" % word)
                else:
                    out.append(word)
                i = j
                continue
            if ch in GREEK:
                out.append(GREEK[ch] + " ")
            elif ch == "&" and self.in_eqarr:
                out.append("&")
            elif ch in MATH_SYMBOLS:
                sym = MATH_SYMBOLS[ch]
                out.append(sym + (" " if sym.startswith("\\") and sym[-1].isalpha() else ""))
            elif ord(ch) < 128:
                out.append(ch)
            else:
                out.append(r"\text{%s}" % ch)
            i += 1
        return "".join(out)

    # ---- yapı elemanları
    def f_oMath(self, el):
        return self.seq(el)

    def f_oMathPara(self, el):
        return r" \\ ".join(self.seq(o) for o in el.findall(M + "oMath"))

    def f_e(self, el):
        return self.seq(el)

    def f_f(self, el):
        num, den = self.arg(el, "num"), self.arg(el, "den")
        typ = self.prop(el, "fPr", "type", "bar")
        if typ == "lin":
            return "{%s}/{%s}" % (num, den)
        if typ == "noBar":
            return r"\genfrac{}{}{0pt}{}{%s}{%s}" % (num, den)
        if typ == "skw":
            return r"{}^{%s}\!/\!{}_{%s}" % (num, den)
        return r"\frac{%s}{%s}" % (num, den)

    @staticmethod
    def _base(s):
        s = s.strip()
        if len(s) <= 1 or re.fullmatch(r"\\[a-zA-Z]+|\\mathit\{[^{}]*\}|\\mathrm\{[^{}]*\}", s):
            return s
        return "{%s}" % s

    def f_sSub(self, el):
        return "%s_{%s}" % (self._base(self.arg(el, "e")), self.arg(el, "sub"))

    def f_sSup(self, el):
        return "%s^{%s}" % (self._base(self.arg(el, "e")), self.arg(el, "sup"))

    def f_sSubSup(self, el):
        return "%s_{%s}^{%s}" % (self._base(self.arg(el, "e")), self.arg(el, "sub"), self.arg(el, "sup"))

    def f_sPre(self, el):
        return "{}_{%s}^{%s}%s" % (self.arg(el, "sub"), self.arg(el, "sup"), self._base(self.arg(el, "e")))

    def f_rad(self, el):
        deg = self.arg(el, "deg")
        hide = self.prop(el, "radPr", "degHide", "0") in ("1", "on", "true")
        if hide or not deg.strip():
            return r"\sqrt{%s}" % self.arg(el, "e")
        return r"\sqrt[%s]{%s}" % (deg, self.arg(el, "e"))

    def f_nary(self, el):
        chr_ = self.prop(el, "naryPr", "chr", "∫")
        op = NARY.get(chr_, MATH_SYMBOLS.get(chr_, chr_))
        lim = self.prop(el, "naryPr", "limLoc", "undOvr" if chr_ != "∫" else "subSup")
        sub_hide = self.prop(el, "naryPr", "subHide", "0") in ("1", "on", "true")
        sup_hide = self.prop(el, "naryPr", "supHide", "0") in ("1", "on", "true")
        sub, sup = self.arg(el, "sub"), self.arg(el, "sup")
        s = op + (r"\limits" if lim == "undOvr" and (sub or sup) and not (sub_hide and sup_hide) else "")
        if sub and not sub_hide:
            s += "_{%s}" % sub
        if sup and not sup_hide:
            s += "^{%s}" % sup
        return s + " " + self.arg(el, "e")

    def f_d(self, el):
        beg = self.prop(el, "dPr", "begChr", "(")
        end = self.prop(el, "dPr", "endChr", ")")
        sep = self.prop(el, "dPr", "sepChr", "|")
        sep_tex = DELIMS.get(sep, sep)
        parts = [self.seq(e) for e in el.findall(M + "e")]
        inner = (r" \middle%s " % sep_tex).join(parts) if len(parts) > 1 else "".join(parts)
        return r"\left%s %s \right%s" % (DELIMS.get(beg, beg), inner, DELIMS.get(end, end))

    def f_acc(self, el):
        chr_ = self.prop(el, "accPr", "chr", "\u0302")
        e = self.arg(el, "e")
        cmd = ACCENTS.get(chr_, r"\hat")
        if cmd in (r"\bar", r"\hat", r"\tilde", r"\vec") and len(re.sub(r"\\[a-z]+", "x", e)) > 1:
            cmd = {r"\bar": r"\overline", r"\hat": r"\widehat", r"\tilde": r"\widetilde",
                   r"\vec": r"\overrightarrow"}[cmd]
        return r"%s{%s}" % (cmd, e)

    def f_bar(self, el):
        pos = self.prop(el, "barPr", "pos", "bot")
        return (r"\overline{%s}" if pos == "top" else r"\underline{%s}") % self.arg(el, "e")

    def f_func(self, el):
        name = self.arg(el, "fName").strip()
        bare = re.sub(r"\\mathrm\{([^{}]*)\}|\\mathit\{([^{}]*)\}", lambda m_: m_.group(1) or m_.group(2), name)
        if bare in FUNC_NAMES:
            name = "\\" + bare
        elif re.fullmatch(r"[A-Za-z]+", bare):
            name = r"\operatorname{%s}" % bare
        return "%s{%s}" % (name, self.arg(el, "e"))

    def f_fName(self, el):
        return self.seq(el)

    def f_limLow(self, el):
        e, lim = self.arg(el, "e").strip(), self.arg(el, "lim")
        if re.fullmatch(r"\\(lim|max|min|sup|inf|limsup|liminf|arg ?max|arg ?min)\s*", e):
            return "%s_{%s}" % (e.strip(), lim)
        return r"\underset{%s}{%s}" % (lim, e)

    def f_limUpp(self, el):
        return r"\overset{%s}{%s}" % (self.arg(el, "lim"), self.arg(el, "e"))

    def f_groupChr(self, el):
        chr_ = self.prop(el, "groupChrPr", "chr", "\u23df")
        e = self.arg(el, "e")
        if chr_ in ("\u23de", "\u23b4"):
            return r"\overbrace{%s}" % e
        if chr_ in ("\u2192", "\u20d7"):
            return r"\xrightarrow{%s}" % e
        if chr_ == "\u2190":
            return r"\xleftarrow{%s}" % e
        return r"\underbrace{%s}" % e

    def f_eqArr(self, el):
        old = self.in_eqarr
        self.in_eqarr = True
        rows = [self.seq(e) for e in el.findall(M + "e")]
        self.in_eqarr = old
        env = "aligned" if any("&" in r for r in rows) else "gathered"
        return r"\begin{%s}%s\end{%s}" % (env, r" \\ ".join(rows), env)

    def f_m(self, el):
        rows = []
        for mr in el.findall(M + "mr"):
            rows.append(" & ".join(self.seq(e) for e in mr.findall(M + "e")))
        return r"\begin{matrix}%s\end{matrix}" % r" \\ ".join(rows)

    def f_box(self, el):
        return self.arg(el, "e")

    def f_borderBox(self, el):
        return r"\boxed{%s}" % self.arg(el, "e")

    def f_phant(self, el):
        return r"\phantom{%s}" % self.arg(el, "e")

    def f_num(self, el):
        return self.seq(el)

    def f_den(self, el):
        return self.seq(el)


# --------------------------------------------------------------------------
# DOCX okuyucu
# --------------------------------------------------------------------------
class Docx:
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.doc = etree.fromstring(self.z.read("word/document.xml"))
        self.rels = self._rels("word/_rels/document.xml.rels")
        self.styles = self._styles()
        self.numfmt = self._numbering()

    def _rels(self, name):
        rels = {}
        if name not in self.z.namelist():
            return rels
        root = etree.fromstring(self.z.read(name))
        for rel in root:
            rels[rel.get("Id")] = (rel.get("Target"), rel.get("TargetMode"), rel.get("Type", ""))
        return rels

    def _styles(self):
        styles = {}
        if "word/styles.xml" not in self.z.namelist():
            return styles
        root = etree.fromstring(self.z.read("word/styles.xml"))
        for s in root.findall(W + "style"):
            sid = s.get(W + "styleId")
            name = wval(s.find(W + "name"), sid) or ""
            numpr = s.find(f"{W}pPr/{W}numPr")
            outline = wval(s.find(f"{W}pPr/{W}outlineLvl"))
            styles[sid] = {"name": name, "numbered": numpr is not None,
                           "outline": int(outline) if outline and outline.isdigit() else None}
        return styles

    def _numbering(self):
        """numId -> {ilvl: numFmt} eşlemesi (madde imi mi numara mı?)."""
        res = {}
        if "word/numbering.xml" not in self.z.namelist():
            return res
        root = etree.fromstring(self.z.read("word/numbering.xml"))
        abstract = {}
        for an in root.findall(W + "abstractNum"):
            lv = {}
            for lvl in an.findall(W + "lvl"):
                lv[lvl.get(W + "ilvl")] = wval(lvl.find(W + "numFmt"), "decimal")
            abstract[an.get(W + "abstractNumId")] = lv
        for num in root.findall(W + "num"):
            aid = wval(num.find(W + "abstractNumId"))
            res[num.get(W + "numId")] = abstract.get(aid, {})
        return res

    def media(self, rid):
        target = self.rels.get(rid, (None,))[0]
        if not target:
            return None, None
        path = posixpath.normpath(posixpath.join("word", target)) if not target.startswith("/") else target[1:]
        try:
            return path, self.z.read(path)
        except KeyError:
            return path, None

    def body_blocks(self):
        body = self.doc.find(W + "body")
        blocks = []
        self._collect(body, blocks)
        return blocks

    def _collect(self, parent, blocks):
        for el in parent:
            if el.tag == W + "p":
                blocks.append(Para(el, self))
            elif el.tag == W + "tbl":
                blocks.append(Table(el, self))
            elif el.tag == W + "sdt":
                c = el.find(W + "sdtContent")
                if c is not None:
                    self._collect(c, blocks)
            elif el.tag in (W + "customXml", W + "ins"):
                self._collect(el, blocks)


class Para:
    """Bir Word paragrafı: biçimli metin parçaları (token), denklemler ve resimler."""

    def __init__(self, el, doc):
        self.el = el
        self.doc = doc
        ppr = el.find(W + "pPr")
        self.style_id = wval(ppr.find(W + "pStyle")) if ppr is not None else None
        st = doc.styles.get(self.style_id or "", {})
        self.style = (st.get("name") or self.style_id or "Normal")
        self.jc = wval(ppr.find(W + "jc")) if ppr is not None else None
        numpr = ppr.find(W + "numPr") if ppr is not None else None
        self.num_id = wval(numpr.find(W + "numId")) if numpr is not None else None
        self.ilvl = int(wval(numpr.find(W + "ilvl"), "0")) if numpr is not None else 0
        if self.num_id == "0":
            self.num_id = None
        if self.num_id is None and st.get("numbered"):
            self.num_id = "style"
        outline = wval(ppr.find(W + "outlineLvl")) if ppr is not None else None
        self.outline = int(outline) if outline and outline.isdigit() else st.get("outline")
        self.tokens = []   # ('t', text, fmt) | ('m', latex, display, number) | ('img', rid)
        self._fld = []
        self._omml = OMML()
        self._walk(el, (False, False, None, None))

    # fmt = (bold, italic, vert, link)
    def _walk(self, el, fmt):
        for c in el:
            tag = c.tag
            if tag == W + "pPr":
                continue
            if tag == W + "r":
                self._run(c, fmt)
            elif tag == W + "hyperlink":
                rid = c.get(R + "id")
                url = self.doc.rels.get(rid, (None,))[0] if rid else None
                self._walk(c, fmt[:3] + (url,))
            elif tag == W + "fldSimple":
                instr = c.get(W + "instr", "")
                mt = re.search(r'HYPERLINK\s+"([^"]+)"', instr)
                self._walk(c, fmt[:3] + ((mt.group(1),) if mt else (fmt[3],)))
            elif tag in (W + "smartTag", W + "customXml", W + "ins", W + "moveTo", W + "dir", W + "bdo"):
                self._walk(c, fmt)
            elif tag == W + "sdt":
                content = c.find(W + "sdtContent")
                if content is not None:
                    self._walk(content, fmt)
            elif tag == M + "oMathPara":
                for o in c.findall(M + "oMath"):
                    latex, num = self._omml.convert(o)
                    self.tokens.append(("m", latex, True, num))
            elif tag == M + "oMath":
                latex, num = self._omml.convert(c)
                self.tokens.append(("m", latex, False, num))
            elif tag == MC + "AlternateContent":
                ch = c.find(MC + "Choice")
                if ch is None:
                    ch = c.find(MC + "Fallback")
                if ch is not None:
                    self._walk(ch, fmt)

    def _run(self, r, fmt):
        rpr = r.find(W + "rPr")
        bold, italic, vert, link = fmt
        if rpr is not None:
            if rpr.find(W + "b") is not None:
                bold = on_off(rpr.find(W + "b"))
            if rpr.find(W + "i") is not None:
                italic = on_off(rpr.find(W + "i"))
            va = wval(rpr.find(W + "vertAlign"))
            if va in ("superscript", "subscript"):
                vert = va
            if rpr.find(W + "vanish") is not None and on_off(rpr.find(W + "vanish")):
                return  # gizli metin
        f = (bold, italic, vert, link)
        for c in r:
            tag = c.tag
            if tag == W + "fldChar":
                t = c.get(W + "fldCharType")
                if t == "begin":
                    self._fld.append({"mode": "instr", "instr": "", "link": None})
                elif t == "separate" and self._fld:
                    top = self._fld[-1]
                    top["mode"] = "result"
                    mt = re.search(r'HYPERLINK\s+"([^"]+)"', top["instr"])
                    if mt:
                        top["link"] = mt.group(1)
                elif t == "end" and self._fld:
                    self._fld.pop()
                continue
            if tag == W + "instrText":
                if self._fld:
                    self._fld[-1]["instr"] += c.text or ""
                continue
            if self._fld and self._fld[-1]["mode"] == "instr":
                continue
            ff = f
            for fl in self._fld:
                if fl["link"]:
                    ff = f[:3] + (fl["link"],)
            if tag == W + "t":
                self.tokens.append(("t", c.text or "", ff))
            elif tag == W + "tab":
                self.tokens.append(("t", " ", ff))
            elif tag in (W + "br", W + "cr"):
                if c.get(W + "type") != "page":
                    self.tokens.append(("t", "\n", ff))
            elif tag == W + "noBreakHyphen":
                self.tokens.append(("t", "-", ff))
            elif tag == W + "sym":
                code = int(c.get(W + "char", "0"), 16)
                if code >= 0xF000:
                    code -= 0xF000
                self.tokens.append(("t", chr(code) if code else "", ff))
            elif tag in (W + "drawing", W + "pict", W + "object"):
                self._images(c)
            elif tag == MC + "AlternateContent":
                ch = c.find(MC + "Choice")
                if ch is None:
                    ch = c.find(MC + "Fallback")
                if ch is not None:
                    for d in ch:
                        if d.tag in (W + "drawing", W + "pict", W + "object"):
                            self._images(d)
            elif tag in (W + "footnoteReference", W + "endnoteReference"):
                warn("Dipnot/sonnot bulundu, LaTeX'e aktarılmadı: '%s...'" % self.text()[:40])

    def _images(self, el):
        rids = el.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS)
        for rid in rids:
            self.tokens.append(("img", rid))
        if el.xpath(".//c:chart", namespaces={"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}):
            warn("Word grafiği (chart) bulundu; resim olarak kaydedip Word'e tekrar ekleyin.")

    # ---- yardımcı özellikler
    def text(self):
        return "".join(t[1] for t in self.tokens if t[0] == "t")

    @property
    def images(self):
        return [t[1] for t in self.tokens if t[0] == "img"]

    @property
    def maths(self):
        return [t for t in self.tokens if t[0] == "m"]

    def is_empty(self):
        return not self.text().strip() and not self.images and not self.maths

    def all_bold(self):
        ts = [t for t in self.tokens if t[0] == "t" and t[1].strip()]
        return bool(ts) and all(t[2][0] for t in ts)

    def heading_level(self):
        if self.style.lower().startswith("heading"):
            mt = re.search(r"(\d)", self.style)
            return int(mt.group(1)) if mt else 1
        if self.style.lower() == "title":
            return None
        return None


class Table:
    def __init__(self, el, doc):
        self.el = el
        self.doc = doc
        self.rows = []  # her satır: [(cell_paras, span, vmerge)]
        for tr in el.findall(W + "tr"):
            row = []
            for tc in tr.findall(W + "tc"):
                tcpr = tc.find(W + "tcPr")
                span = int(wval(tcpr.find(W + "gridSpan"), "1")) if tcpr is not None else 1
                vm = tcpr.find(W + "vMerge") if tcpr is not None else None
                vmerge = None if vm is None else wval(vm, "continue")
                paras = []
                Docx._collect(doc, tc, paras)
                row.append((paras, span, vmerge))
            if row:
                self.rows.append(row)

    def ncols(self):
        return max((sum(c[1] for c in r) for r in self.rows), default=0)

    def is_empty(self):
        return False


# --------------------------------------------------------------------------
# Biçimli metni LaTeX'e çevirme
# --------------------------------------------------------------------------
def norm_key(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", s.lower())


YEAR_PAT = r"(?:(?:19|20)\d\d[a-z]?|n\.\s?d\.)"
CITE_PAREN = re.compile(r"\(([^()]*?%s[^()]*?)\)" % YEAR_PAT)
CITE_PART = re.compile(r"^(?P<auth>[^\d]+?),?\s*(?P<years>%s(?:\s*,\s*%s)*)\s*$" % (YEAR_PAT, YEAR_PAT))
CITE_NARR = re.compile(r"(?P<auth>[A-Z][\w'’\-]+(?:\s+et\s+al\.|\s+(?:&|and)\s+[A-Z][\w'’\-]+)?)\s*\((?P<yr>%s)\)"
                       % YEAR_PAT)
NUM_CITE = re.compile(r"\[\s*(\d+(?:\s*[,;–\-]\s*\d+)*)\s*\]")
XREF_RE = re.compile(r"\b(Tables?|Figures?|Figs?\.|Equations?|Eqs?\.)\s*\(?(\d+)\)?")


class Linker:
    """Metindeki atıfları kaynakçaya, 'Figure 2 / Table 1 / Equation 3' ifadelerini
    ilgili şekil/tablo/denkleme tıklanabilir bağlantı yapar."""

    def __init__(self, ref_paras, n_fig, n_tab, n_eq):
        self.refs = []
        self.n = {"fig": n_fig, "tab": n_tab, "eq": n_eq}
        self.used = set()
        self.missing = set()
        numbered = 0
        for k, p in enumerate(ref_paras):
            txt = p.text().strip()
            ym = re.search(r"\((%s)\)" % YEAR_PAT, txt)
            head = txt[:ym.start()] if ym else txt
            first = re.split(r"[,(]", re.sub(r"^\s*\[?\d+[\].]\s*", "", head), maxsplit=1)[0].strip(" .")
            mnum = re.match(r"^\s*\[(\d+)\]|^\s*(\d+)[.)]\s", txt)
            if mnum:
                numbered += 1
            self.refs.append({"key": "ref:%d" % (k + 1), "first": norm_key(first),
                              "year": (ym.group(1).replace(" ", "") if ym else ""), "head": norm_key(head),
                              "num": int(mnum.group(1) or mnum.group(2)) if mnum else (k + 1),
                              "label": "%s (%s)" % (first, ym.group(1) if ym else "?")})
        self.numeric = bool(self.refs) and numbered >= 0.6 * len(self.refs)

    def find(self, auth, year):
        parts = [x for x in re.split(r"\s*(?:&|\band\b)\s*", re.sub(r",?\s*et\s*al\.?", "", auth)) if x.strip()]
        if not parts:
            return None
        words = parts[0].split()
        keys = {norm_key(parts[0]), norm_key(words[-1]), norm_key(words[0])}
        year = year.replace(" ", "")
        cands = [r for r in self.refs if r["first"] in keys and r["year"] == year]
        if not cands:
            cands = [r for r in self.refs if r["first"] in keys and r["year"][:4] == year[:4]]
        if len(cands) > 1 and len(parts) > 1:
            second = norm_key(parts[1].split()[-1])
            cands = [r for r in cands if second in r["head"]] or cands
        return cands[0] if cands else None

    def spans(self, text):
        out = []
        if self.refs:
            if self.numeric:
                bynum = {r["num"]: r for r in self.refs}
                for mt in NUM_CITE.finditer(text):
                    for nm in re.finditer(r"\d+", mt.group(1)):
                        r = bynum.get(int(nm.group(0)))
                        if r:
                            a = mt.start(1) + nm.start()
                            out.append((a, a + len(nm.group(0)), r["key"]))
                            self.used.add(r["key"])
            else:
                for mt in CITE_PAREN.finditer(text):
                    pos = mt.start(1)
                    for part in mt.group(1).split(";"):
                        lead = len(part) - len(part.lstrip())
                        pm = CITE_PART.match(part.strip())
                        if pm:
                            a0 = pos + lead
                            years = list(re.finditer(YEAR_PAT, pm.group("years")))
                            for j, y in enumerate(years):
                                r = self.find(pm.group("auth"), y.group(0))
                                ys = a0 + pm.start("years")
                                a, b = (a0, ys + y.end()) if j == 0 else (ys + y.start(), ys + y.end())
                                if r:
                                    out.append((a, b, r["key"]))
                                    self.used.add(r["key"])
                                else:
                                    self.missing.add("%s, %s" % (pm.group("auth").strip(), y.group(0)))
                        pos += len(part) + 1
                for mt in CITE_NARR.finditer(text):
                    if any(a <= mt.start() < b for a, b, _ in out):
                        continue
                    r = self.find(mt.group("auth"), mt.group("yr"))
                    if r:
                        out.append((mt.start(), mt.end(), r["key"]))
                        self.used.add(r["key"])
        for mt in XREF_RE.finditer(text):
            kind = {"t": "tab", "f": "fig", "e": "eq"}[mt.group(1)[0].lower()]
            n = int(mt.group(2))
            if 1 <= n <= self.n[kind] and not any(a <= mt.start() < b for a, b, _ in out):
                out.append((mt.start(), mt.end(), "%s:%d" % (kind, n)))
        return sorted(out)

    def apply(self, tokens):
        text = "".join(t[1] for t in tokens if t[0] == "t")
        spans = self.spans(text)
        if not spans:
            return tokens
        cuts = sorted({a for a, _, _ in spans} | {b for _, b, _ in spans})
        out, pos = [], 0
        for t in tokens:
            if t[0] != "t":
                out.append(t)
                continue
            s, start = t[1], pos
            pos += len(s)
            pieces = [start] + [c for c in cuts if start < c < pos] + [pos]
            for a, b in zip(pieces, pieces[1:]):
                piece = s[a - start:b - start]
                target = next((tg for x, y, tg in spans if x <= a and b <= y), None)
                fmt = t[2]
                if target and not fmt[3]:
                    fmt = fmt[:3] + ("@@" + target,)
                out.append(("t", piece, fmt))
        return out


LINKER = None


def render_tokens(tokens, ref_mode=False, keep_newlines=False):
    """Paragraf token'larını LaTeX'e çevirir (kalın/italik/alt-üst simge/bağlantı/denklem)."""
    if LINKER is not None and not ref_mode:
        tokens = LINKER.apply(tokens)
    # aynı biçimdeki ardışık metinleri birleştir
    merged = []
    for t in tokens:
        if t[0] == "t" and merged and merged[-1][0] == "t" and merged[-1][2] == t[2]:
            merged[-1] = ("t", merged[-1][1] + t[1], t[2])
        else:
            merged.append(t)
    out = []
    for t in merged:
        if t[0] == "t":
            txt = t[1]
            if not keep_newlines:
                txt = txt.replace("\n", " ")
            txt = re.sub(r"[ \t\u00a0]{2,}", " ", txt) if not keep_newlines else txt
            if ref_mode:
                txt = DOI_DUP_RE.sub("https://doi.org/", txt)
            if not txt:
                continue
            bold, italic, vert, link = t[2]
            if ref_mode and italic and txt.upper() == txt and len(re.findall(r"[A-Z]{2,}", txt)) >= 3:
                txt = title_case_caps(txt)
            lead = txt[: len(txt) - len(txt.lstrip(" "))]
            trail = txt[len(txt.rstrip(" ")):]
            core = txt.strip(" ")
            if not core:
                out.append(" ")
                continue
            if link and link.startswith("@@"):
                s = r"\hyperlink{%s}{%s}" % (link[2:], render_plain(core, italic, False, True))
            elif link and not link.startswith("#"):
                if core == link or URL_RE.fullmatch(core):
                    s = (r"DOI:\href{%s}{%s}" if ref_mode and "doi.org/" in link.lower() and not "".join(out).rstrip().lower().endswith("doi:")
                         else r"\href{%s}{%s}") % (url_escape(link), url_escape(DOI_DUP_RE.sub("https://doi.org/", core)))
                    if not (ref_mode and "doi.org/" in link.lower()):
                        s = r"\url{%s}" % url_escape(core)
                elif link.lower().startswith("mailto:"):
                    s = r"\href{%s}{%s}" % (url_escape(link), render_plain(core, italic, ref_mode, True))
                else:
                    s = r"\href{%s}{%s}" % (url_escape(link), render_plain(core, italic, ref_mode, True))
            else:
                s = render_plain(core, italic, ref_mode)
            if italic:
                s = r"\textit{%s}" % s
            if bold:
                s = r"\textbf{%s}" % s
            if vert == "superscript":
                s = r"\textsuperscript{%s}" % s
            elif vert == "subscript":
                s = r"\textsubscript{%s}" % s
            out.append(lead + s + trail)
        elif t[0] == "m":
            if t[1]:
                out.append("$%s$" % t[1])
    s = "".join(out)
    s = re.sub(r" {2,}", " ", s)
    # "\textit{et al.}" iki kez sarılmışsa düzelt
    s = s.replace(r"\textit{\textit{et al.}}", r"\textit{et al.}")
    return s.strip()


def strip_eq_number_text(txt):
    """Denklem paragrafındaki denklem dışı metinden '(1)' veya ')1(' numarasını çıkarır."""
    t = re.sub(r"\s+", "", txt)
    if not t:
        return True, None
    for pat in (r"\((\d+[a-z]?)\)", r"\)(\d+[a-z]?)\(", r"(\d+[a-z]?)", r"#\((\d+)\)"):
        mt = re.fullmatch(pat, t)
        if mt:
            return True, mt.group(1)
    return False, None


# --------------------------------------------------------------------------
# Kapak mektubu
# --------------------------------------------------------------------------
PLACEHOLDER_RE = re.compile(r"click or tap here to enter text\.?|click here to enter text\.?", re.I)
TITLE_PREFIX_RE = re.compile(
    r"^(?:(?:asst|assist|assoc|associate|assistant|prof|professor|dr|lect|lecturer|mr|mrs|ms|miss|eng|engr|mrs)\.?\s+)+",
    re.I)


def strip_titles(name):
    return TITLE_PREFIX_RE.sub("", name.strip()).strip()


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
            for nxt in lines[i + 1:]:
                if nxt:
                    info["title"] = nxt
                    break
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
    for a in info["authors"]:
        orcid = a.get("orcid", "")
        mt = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", orcid)
        a["orcid"] = "https://orcid.org/" + mt.group(1) if mt else None
    return info


def name_key(name):
    n = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode().lower()
    parts = re.findall(r"[a-z]+", n)
    return parts


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
KEYWORDS_HEAD = re.compile(r"^\s*(key\s*-?\s*words?|index terms)\s*[:.\-–]?\s*", re.I)
CAPTION_RE = re.compile(r"^\s*(figure|fig\.?|table)\s*(\d+)\s*[:.\-–—]?\s*(.*)$", re.I | re.S)
ARTICLE_TYPE_RE = re.compile(r"^(original|review|research|short|case)\s+(article|paper|communication|report)$", re.I)


def clean_heading(txt):
    t = re.sub(r"\s+", " ", txt).strip()
    t = re.sub(r"^\d+(\.\d+)*\.?\s+(?=[A-Za-z])", "", t)  # elle yazılmış numarayı kaldır
    return t.rstrip(":").strip()


def manual_number_level(txt):
    mt = re.match(r"^\s*(\d+(?:\.\d+)*)\.?\s+[A-Za-z]", txt)
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
        self.keywords = ""
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
    title_parts = []
    while i < n and isinstance(paras[i], Para) and not paras[i].is_empty():
        title_parts.append(paras[i])
        i += 1
        # başlık genelde tek paragraftır; sonraki paragraf yazar satırıysa dur
        nxt = nxt_nonempty(i)
        if nxt < n and _looks_like_authors(paras[nxt]):
            break
        if len(title_parts) >= 1 and (nxt != i):
            break
    art.title = " ".join(re.sub(r"\s+", " ", p.text()).strip() for p in title_parts)
    i = nxt_nonempty(i)

    # Yazarlar
    if i < n and isinstance(paras[i], Para):
        art.authors = parse_authors(paras[i])
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
            i += 1
            break
        if abstract_mode:
            if b.heading_level() or (b.num_id and b.all_bold() and len(txt) < 150):
                break  # anahtar kelime yoksa ilk başlıkta dur
            art.abstract.append(render_tokens(b.tokens))
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


def _looks_like_authors(p):
    sup = [t for t in p.tokens if t[0] == "t" and t[2][2] == "superscript" and t[1].strip()]
    txt = p.text()
    return bool(sup) and ("," in txt or " and " in txt or len(sup) >= 1)


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
            if name:
                authors.append({"name": name, "marks": marks})
        tail = s[s.rfind("}") + 1:].strip(" ,;")
        tail = re.sub(r"^and\s+", "", tail, flags=re.I).strip()
        if tail:
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
    if cl.get("title") and art.title and name_key(cl["title"]) != name_key(art.title):
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
    return art


# --------------------------------------------------------------------------
# LaTeX üretimi
# --------------------------------------------------------------------------
class Writer:
    def __init__(self, art, doc, outdir, opts):
        self.art = art
        self.doc = doc
        self.outdir = outdir
        self.opts = opts
        self.out = []
        self.fig_count = 0
        self.tab_count = 0
        self.eq_count = 0
        self.pending_imgs = []      # [[rid,...], ...]
        self.pending_fig_caption = None
        self.pending_tab_caption = None
        self.in_multicol = False
        self.list_env = None
        self.section = None         # o anki bölüm başlığı (küçük harf)
        self.contrib_items = []
        self.saved_hashes = {}

    # ---- çıktı yardımcıları
    def emit(self, s=""):
        self.out.append(s)

    def open_cols(self):
        if not self.in_multicol:
            self.emit(r"\begin{multicols}{2}")
            self.in_multicol = True

    def close_cols(self):
        self.close_list()
        if self.in_multicol:
            self.emit(r"\end{multicols}")
            self.in_multicol = False

    def close_list(self):
        if self.list_env:
            self.emit(r"\end{%s}" % self.list_env)
            self.list_env = None
        self.flush_contrib()

    def flush_contrib(self):
        if self.contrib_items:
            items = [re.sub(r"^[\s\"“”'‘’]+|[\s\"“”'‘’.,;]+$", "", x) for x in self.contrib_items]
            items = [self.abbreviate_names(x).strip(" ,;") for x in items if x]
            self.emit(r"\noindent " + "; ".join(items) + ".")
            self.emit()
            self.contrib_items = []

    # ---- resimler
    def save_image(self, rid):
        path, data = self.doc.media(rid)
        if data is None:
            warn("Resim bulunamadı: %s" % rid)
            return None
        h = hashlib.md5(data).hexdigest()
        if h in self.saved_hashes:
            return self.saved_hashes[h]
        ext = os.path.splitext(path)[1].lower()
        name = "Figure%d" % (self.fig_count + 1)
        if ext in (".png", ".jpg", ".jpeg", ".pdf"):
            fname = name + ext
            with open(os.path.join(self.outdir, fname), "wb") as f:
                f.write(data)
        elif ext in (".tif", ".tiff", ".gif", ".bmp", ".webp"):
            fname = name + ".png"
            try:
                import io
                from PIL import Image
                Image.open(io.BytesIO(data)).save(os.path.join(self.outdir, fname))
            except Exception as e:  # noqa
                warn("%s resmi PNG'ye çevrilemedi (Pillow kurun: pip install Pillow): %s" % (path, e))
                fname = name + ext
                with open(os.path.join(self.outdir, fname), "wb") as f:
                    f.write(data)
        else:
            fname = name + ext
            with open(os.path.join(self.outdir, fname), "wb") as f:
                f.write(data)
            warn("%s biçimi (%s) LaTeX'te doğrudan kullanılamaz; PNG/PDF'ye çevirin." % (fname, ext))
        self.saved_hashes[h] = fname
        return fname

    def prescan(self):
        """Bağlantılar için şekil/tablo/denklem sayılarını önceden bulur."""
        n_fig = n_tab = n_eq = 0
        for b in self.art.body:
            if isinstance(b, Table):
                n_tab += 1
            elif not b.is_empty():
                kind = self.is_caption(b)
                if kind == "figure":
                    n_fig += 1
                elif b.maths and kind is None:
                    ok, num = strip_eq_number_text(b.text().strip())
                    if ok and (num or any(m_[3] for m_ in b.maths)):
                        n_eq += 1
        return n_fig, n_tab, n_eq

    def emit_figure(self, rids, caption):
        wide = "all" in self.opts.wide_figs or str(self.fig_count + 1) in self.opts.wide_figs
        if wide:
            self.close_cols()
            return self._emit_wide_figure(rids, caption)
        self.open_cols()
        self.close_list()
        files = []
        for rid in rids:
            fn = self.save_image(rid)
            if fn and fn not in files:
                files.append(fn)
        if not files:
            return
        self.fig_count += 1
        w = r"\linewidth" if len(files) == 1 else r"%.2f\linewidth" % (0.98 / len(files))
        self.emit(r"\noindent\begin{minipage}{\linewidth}")
        self.emit(r"\centering\hypertarget{fig:%d}{}" % self.fig_count)
        for fn in files:
            self.emit(r"\includegraphics[width=%s, height=0.42\textheight, keepaspectratio]{%s}" % (w, fn))
        if caption:
            self.emit(r"\captionof{figure}{%s}" % caption)
        self.emit(r"\end{minipage}")
        self.emit(r"\par\vspace{6pt}")
        self.emit()

    def _emit_wide_figure(self, rids, caption):
        files = [f for f in (self.save_image(r) for r in rids) if f]
        if not files:
            return
        self.fig_count += 1
        w = r"\textwidth" if len(files) == 1 else r"%.2f\textwidth" % (0.98 / len(files))
        self.emit(r"\noindent\begin{minipage}{\textwidth}")
        self.emit(r"\centering\hypertarget{fig:%d}{}" % self.fig_count)
        for fn in files:
            self.emit(r"\includegraphics[width=%s, height=0.45\textheight, keepaspectratio]{%s}" % (w, fn))
        if caption:
            self.emit(r"\captionof{figure}{%s}" % caption)
        self.emit(r"\end{minipage}")
        self.emit(r"\par\vspace{6pt}")
        self.emit()

    def flush_images(self):
        for rids in self.pending_imgs:
            self.emit_figure(rids, None)
        self.pending_imgs = []

    # ---- tablolar
    def emit_table(self, tbl, caption):
        self.close_cols()
        self.tab_count += 1
        ncols = tbl.ncols()
        if ncols == 0:
            return
        body_cells = [c for r in tbl.rows[1:] for c in r]
        texts = [" ".join(p.text() for p in c[0] if isinstance(p, Para)).strip() for c in body_cells]
        numeric = sum(1 for t in texts if re.fullmatch(r"[-+]?[\d.,%\s]*", t)) / max(1, len(texts))
        if numeric > 0.6:
            colspec = r"@{}>{\raggedright\arraybackslash}X" + r">{\centering\arraybackslash}X" * (ncols - 1) + "@{}"
            size = r"\footnotesize" if ncols >= 6 else r"\small"
        else:
            colspec = r"@{}" + r">{\raggedright\arraybackslash}X" * ncols + "@{}"
            size = r"\small"
        self.emit(r"\noindent")
        self.emit(r"\begin{minipage}{\textwidth}")
        self.emit(r"\centering")
        self.emit(r"\hypertarget{tab:%d}{}" % self.tab_count)
        if caption:
            self.emit(r"\captionof{table}{%s}" % caption)
        self.emit(r"\vspace{-6pt}")
        self.emit(size)
        self.emit(r"\begin{tabularx}{\textwidth}{%s}" % colspec)
        self.emit(r"\toprule")
        for ri, row in enumerate(tbl.rows):
            cells = []
            for paras, span, vmerge in row:
                if vmerge == "continue":
                    content = ""
                else:
                    content = r" \newline ".join(
                        render_tokens(p.tokens) for p in paras if isinstance(p, Para) and not p.is_empty())
                if ri == 0 and content and not content.startswith(r"\textbf"):
                    content = r"\textbf{%s}" % content
                if span > 1:
                    align = "c" if numeric > 0.6 else "l"
                    content = r"\multicolumn{%d}{%s}{%s}" % (span, align, content)
                cells.append(content)
            line = " & ".join(cells) + r" \\"
            if ri == 0:
                line += r" \midrule"
            elif ri < len(tbl.rows) - 1:
                line += r" \addlinespace[0.3em]"
            self.emit(line)
        self.emit(r"\bottomrule")
        self.emit(r"\end{tabularx}")
        self.emit(r"\end{minipage}")
        self.emit(r"\par\vspace{6pt}")
        self.emit()

    # ---- denklemler
    def emit_equation(self, latex, number):
        self.open_cols()
        self.close_list()
        if number:
            self.eq_count += 1
            tag = "" if number == str(self.eq_count) else r" \tag{%s}" % number
            if tag:
                warn("Denklem numarası sırası farklı: beklenen %d, Word'de (%s)." % (self.eq_count, number))
            self.emit(r"\hypertarget{eq:%d}{}" % self.eq_count)
            self.emit(r"\begin{equation}")
            self.emit("  " + latex + tag)
            self.emit(r"\end{equation}")
        else:
            self.emit(r"\begin{equation*}")
            self.emit("  " + latex)
            self.emit(r"\end{equation*}")

    # ---- caption
    @staticmethod
    def caption_text(p):
        toks = list(p.tokens)
        # "Figure 1" önekini token'lardan kaldır
        full = p.text()
        mt = CAPTION_RE.match(full)
        cut = len(full) - len(mt.group(3))
        out, pos = [], 0
        for t in toks:
            if t[0] != "t":
                out.append(t)
                continue
            s = t[1]
            if pos + len(s) <= cut:
                pos += len(s)
                continue
            if pos < cut:
                s = s[cut - pos:]
                pos = cut
            pos += len(t[1])
            # caption'da kalın/italik genel biçimi kaldır (sadece alt/üst simge kalsın)
            out.append(("t", s, (False, t[2][1] and not all(x[2][1] for x in toks if x[0] == "t"), t[2][2], t[2][3])))
        cap = render_tokens(out)
        cap = re.sub(r"^[:.\-–—\s]+", "", cap).strip()
        if cap and cap[-1] not in ".?!":
            cap += "."
        return cap

    def is_caption(self, p):
        txt = p.text().strip()
        mt = CAPTION_RE.match(txt)
        if not mt or len(txt) > 400:
            return None
        rest = mt.group(3)
        style = p.style.lower()
        if ("caption" in style or p.jc == "center" or p.images or
                re.match(r"^\s*(figure|fig\.?|table)\s*\d+\s*[:.\-–—]", txt, re.I)):
            return "table" if mt.group(1).lower().startswith("tab") else "figure"
        # Gövde cümlesi mi? ("Figure 5 shows ...")
        if rest and rest[0].islower():
            return None
        return None

    # ---- ana gövde
    def run(self):
        art = self.art
        self.open_cols()
        blocks = art.body
        k = 0
        while k < len(blocks):
            b = blocks[k]
            k += 1
            if isinstance(b, Table):
                self.flush_images()
                cap = self.pending_tab_caption
                self.pending_tab_caption = None
                if cap is None:  # tablonun altındaki başlık
                    j = k
                    while j < len(blocks) and isinstance(blocks[j], Para) and blocks[j].is_empty():
                        j += 1
                    if j < len(blocks) and isinstance(blocks[j], Para) and self.is_caption(blocks[j]) == "table":
                        cap = self.caption_text(blocks[j])
                        k = j + 1
                if cap is None:
                    warn("Tablo %d için başlık (caption) bulunamadı." % (self.tab_count + 1))
                self.emit_table(b, cap)
                continue
            p = b
            if p.is_empty():
                continue
            txt = p.text().strip()

            # Resimler
            if p.images:
                rids = []
                for rid in p.images:
                    _, data = self.doc.media(rid)
                    h = hashlib.md5(data or b"").hexdigest()
                    if h not in [hashlib.md5((self.doc.media(r)[1] or b"")).hexdigest() for r in rids]:
                        rids.append(rid)
                if self.pending_fig_caption is not None:
                    self.emit_figure(rids, self.pending_fig_caption)
                    self.pending_fig_caption = None
                else:
                    self.pending_imgs.append(rids)
                if not txt:
                    continue

            # Şekil / tablo başlıkları
            kind = self.is_caption(p)
            if kind == "figure":
                cap = self.caption_text(p)
                if self.pending_imgs:
                    self.emit_figure(self.pending_imgs.pop(0), cap)
                else:
                    self.pending_fig_caption = cap
                continue
            if kind == "table":
                self.pending_tab_caption = self.caption_text(p)
                continue

            # Denklemler
            if p.maths:
                ok, num_txt = strip_eq_number_text(txt)
                if ok:
                    maths = p.maths
                    num = num_txt or next((m_[3] for m_ in maths if m_[3]), None)
                    latex = r" \quad ".join(m_[1] for m_ in maths if m_[1])
                    self.emit_equation(latex, num)
                    continue

            # Başlıklar
            if self.handle_heading(p, txt):
                continue

            # Liste maddeleri
            if p.num_id and p.num_id != "style":
                if self.section and re.search(r"contribution", self.section):
                    self.contrib_items.append(render_tokens(p.tokens))
                    continue
                fmt = self.doc.numfmt.get(p.num_id, {}).get(str(p.ilvl), "decimal")
                env = "itemize" if fmt in ("bullet", "none") else "enumerate"
                if self.list_env != env:
                    self.close_list()
                    self.emit(r"\begin{%s}" % env)
                    self.list_env = env
                self.emit(r"\item " + render_tokens(p.tokens))
                continue
            self.close_list()

            # Normal paragraf
            self.open_cols()
            s = render_tokens(p.tokens)
            if self.section and re.search(r"contribution", self.section):
                s = self.abbreviate_names(s)
            self.emit(s)
            self.emit()
        self.flush_images()
        if self.pending_fig_caption:
            warn("Resmi olmayan şekil başlığı: %s" % self.pending_fig_caption)
        self.close_list()

    def handle_heading(self, p, txt):
        level = p.heading_level()
        numbered = True
        clean = clean_heading(txt)
        if level is None and not p.maths and len(txt) <= 150 and p.all_bold():
            if p.num_id:
                level = p.ilvl + 1
            elif manual_number_level(txt):
                level = manual_number_level(txt)
            elif BACK_MATTER.match(clean) or re.match(r"^conclusions?$", clean, re.I):
                level, numbered = 2, False
            elif not txt.rstrip().endswith((".", ",", ";")):
                level, numbered = 2, False
        elif level is None and BACK_MATTER.match(clean) and len(txt) < 60:
            level, numbered = 2, False
        if level is None:
            return False
        if BACK_MATTER.match(clean):
            level, numbered = 2, False
        self.flush_images()
        self.close_list()
        self.open_cols()
        self.section = clean.lower()
        if level == 1:
            title = clean if self.opts.keep_heading_case else clean.upper()
            self.emit(r"\section%s{%s}" % ("" if numbered else "*", tex_escape(title)))
        elif level == 2:
            self.emit(r"\subsection%s{%s}" % ("" if numbered else "*", tex_escape(clean)))
        else:
            self.emit(r"\subsubsection%s{%s}" % ("" if numbered else "*", tex_escape(clean)))
        return True

    def abbreviate_names(self, s):
        for au in self.art.authors:
            s = abbreviate_in_text(s, au["name"])
        s = s.replace("“", "").replace("”", "")
        return s

    def references(self):
        self.open_cols()
        self.close_list()
        self.emit(r"\section*{REFERENCES}")
        self.emit(r"\begin{list}{}{")
        self.emit(r"    \setlength{\leftmargin}{1.5em}")
        self.emit(r"    \setlength{\itemindent}{-1.5em}")
        self.emit(r"    \setlength{\parsep}{0pt}")
        self.emit(r"    \setlength{\itemsep}{6pt}")
        self.emit(r"}")
        for k, p in enumerate(self.art.references):
            s = r"\hypertarget{ref:%d}{}" % (k + 1) + render_tokens(p.tokens, ref_mode=True)
            s = re.sub(r"^\s*\[?\d+[\].]\s+", "", s) if self.opts.strip_ref_numbers else s
            self.emit(r"\item " + s)
            self.emit()
        self.emit(r"\end{list}")


def initials(name):
    parts = name.split()
    if len(parts) < 2:
        return name
    ini = []
    for x in parts[:-1]:
        ini.append(x if re.fullmatch(r"[A-Z]\.", x) else x[0] + ".")
    return " ".join(ini + [parts[-1]])


def abbreviate_in_text(s, full):
    short = initials(full)
    parts = full.split()
    s = re.sub(re.escape(full) + r"\s*\d*(?=\W|$)", short, s)
    if len(parts) >= 3:  # "Eyad I." gibi eksik yazılmış adlar
        prefix = " ".join(parts[:-1])
        s = re.sub(re.escape(prefix) + r"(?!\s*" + re.escape(parts[-1]) + r")", short, s)
    return s


# --------------------------------------------------------------------------
# Şablon
# --------------------------------------------------------------------------
TEMPLATE = r"""%!TEX TS-program = xelatex
% Bu dosya sjuoz_word2latex.py tarafından otomatik üretildi.
\documentclass[9pt]{extarticle}
\usepackage{multicol}
\usepackage[top=0.5in, bottom=0.75in, right=0.75in, left=0.75in, headheight=0.25in, includehead]{geometry}
\usepackage{fontspec}
\usepackage{amsmath}
\usepackage{amssymb}
\usepackage{graphicx}
\usepackage{xcolor}
\usepackage{array}
\usepackage{tabularx}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{caption}
\usepackage{titlesec}
\usepackage{fancyhdr}
\usepackage{qrcode}
\usepackage{xurl}
\usepackage{lastpage}
\usepackage{indentfirst}
\usepackage{eso-pic}
\usepackage{tikz}
\usepackage[colorlinks=true, urlcolor=blue, linkcolor=black, citecolor=black]{hyperref}

\titleformat{\section}{\large\bfseries\centering}{\thesection.}{1.0em}{}
\titleformat{\subsection}{\large\bfseries}{\thesubsection.}{0.5em}{}
\titleformat{\subsubsection}{\normalsize\bfseries\itshape}{\thesubsubsection.}{0.5em}{}
\titlespacing*{\section}{0pt}{8pt}{6pt}
\titlespacing*{\subsection}{0pt}{8pt}{4pt}
\titlespacing*{\subsubsection}{0pt}{6pt}{3pt}

\newcommand{\titlesize}{\fontsize{12}{10}}
\renewcommand{\normalsize}{\fontsize{9pt}{9pt}\selectfont}
\renewcommand{\large}{\fontsize{10pt}{10pt}\selectfont}
\renewcommand{\small}{\fontsize{8pt}{8pt}\selectfont}
\captionsetup[table]{labelfont=bf, justification=raggedright, singlelinecheck=false}
\captionsetup[figure]{labelfont=bf, justification=centering, singlelinecheck=false}
\definecolor{sjuozblue}{HTML}{C1D1EB}
\definecolor{colorabs}{HTML}{E2EAF4}

% Yazı tipleri: Times New Roman / Garamond yoksa benzerleri kullanılır
\IfFontExistsTF{<<MAINFONT>>}{\setmainfont{<<MAINFONT>>}}{%
  \IfFontExistsTF{TeX Gyre Termes}{\setmainfont{TeX Gyre Termes}}{%
    \IfFontExistsTF{Liberation Serif}{\setmainfont{Liberation Serif}}{}}}
\IfFontExistsTF{<<HEADFONT>>}{\newfontfamily{\garamond}{<<HEADFONT>>}}{%
  \IfFontExistsTF{EB Garamond}{\newfontfamily{\garamond}{EB Garamond}}{\newcommand{\garamond}{\rmfamily}}}

\newcommand{\sjuozdoi}{<<DOI>>}
% İlk sayfadaki bilgi kutusunun sayfa altından yüksekliği ve metin için ayrılan boşluk
\newcommand{\firstfooterbottom}{0.4in}
\newcommand{\firstpagereserve}{1.4in}
% Sadece ilk sayfada metin alanını kısalt; ilk sayfa basıldıktan sonra eski haline döndür
\makeatletter
\newlength{\sjuozfulltextheight}
\AtBeginDocument{%
  \setlength{\sjuozfulltextheight}{\textheight}%
  \global\advance\textheight by -\firstpagereserve
  \global\@colht\textheight \global\@colroom\textheight \global\vsize\textheight
  \AddToHookNext{shipout/after}{\global\textheight\sjuozfulltextheight}}
\makeatother

% Başlık ve altlık (header/footer)
\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0pt}
\lhead{\large \textit{<<RUNHEAD>>}}
\rhead{\small{\garamond{\textbf{SJUOZ|VOL<<VOL>>|<<MON3>> <<YEAR>>|P<<FIRSTPAGE>>-<<LASTPAGE>>}}}}
\lfoot{\raisebox{8pt}{\textit{DOI:}\href{\sjuozdoi}{\textit{\small{\sjuozdoi}}}}}
\cfoot{\raisebox{8pt}{\thepage}}

\fancypagestyle{firstpage}{%
  \fancyhf{}
  \renewcommand{\headrulewidth}{0pt}
  \rhead{\IfFileExists{SJUOZ.png}{\includegraphics[width=1.04in]{SJUOZ.png}}{\textbf{SJUOZ}}}
  \lhead{\colorbox{sjuozblue}{%
      \begin{minipage}[b]{0.83\textwidth}
        \vspace{0.15cm}
        \hspace{-3pt}\ \normalsize \noindent \garamond{Science Journal of University of Zakho}\\
        \normalsize \textit{\garamond{DOI: \sjuozdoi}}
        \vspace{0.2cm}
      \end{minipage}}}
}

% İlk sayfanın altındaki bilgi kutusu (sayfaya mutlak konumla yerleştirilir)
\newcommand{\sjuozfirstfooter}{%
  \begin{minipage}[b]{\textwidth}
      \rule{\textwidth}{1.5pt} \\[0.15cm]
      \begin{minipage}[t]{0.2\textwidth}\hspace{1pt} \small Access this article online \\
        \vspace{0.1cm}
        {\hypersetup{urlcolor=black}\qrcode[height=2.6162cm]{<<ARTICLEURL>>}}
      \end{minipage}\begin{minipage}[t]{0.4\textwidth}
        \small DOI:\href{\sjuozdoi}{\sjuozdoi} \\ \\[0.5cm]
        Printed ISSN 2663-628X; \\
        Electronic ISSN 2663-6298
      \end{minipage}\begin{minipage}[t]{0.4\textwidth}\small \textbf{Science Journal of University of Zakho} \\
        \textbf{Vol. <<VOL>>, No. <<ISSUE2>>, pp. <<FIRSTPAGE>>-<<LASTPAGE>>, <<MONTH>>-<<YEAR>>} \\[0.5cm]
        This is an open access article under a CC BY-NC-SA 4.0 license \\
        (\href{https://creativecommons.org/licenses/by-nc-sa/4.0/}{https://creativecommons.org/licenses/by-nc-sa/4.0/})
      \end{minipage} \\[0.15cm]
      \rule{\textwidth}{1.5pt} \\[0.15cm]
      \centering \normalsize \thepage
  \end{minipage}}

% ORCID simgesi: orcid.png varsa o, yoksa çizilmiş yeşil "iD" simgesi
\definecolor{orcidgreen}{HTML}{A6CE39}
\newcommand{\orcidicon}{\IfFileExists{orcid.png}{\includegraphics[width=9pt]{orcid.png}}{%
  \tikz[baseline=-0.7ex]{\fill[orcidgreen] (0,0) circle (4.5pt);
  \node[text=white, font=\fontsize{4.5pt}{4.5pt}\selectfont\bfseries, inner sep=0pt] at (0,0) {iD};}}}
\newcommand{\orcidlink}[1]{\href{#1}{\orcidicon}}

\begin{document}
\setcounter{page}{<<FIRSTPAGE>>}
\thispagestyle{firstpage}
\AddToShipoutPictureFG*{\put(\LenToUnit{0.75in},\LenToUnit{\firstfooterbottom}){\sjuozfirstfooter}}

\noindent \large \textbf{<<ARTICLETYPE>>} \\[-5pt]
\begin{center} \titlesize \textbf{<<TITLE>>} \end{center}
\begin{center} \large \textbf{<<AUTHORS>>}
\end{center}
\noindent
<<AFFILIATIONS>>

\begin{center} <<CORRESPONDING>>
\end{center}
\rule{\textwidth}{1.5pt}
\begin{minipage}[t]{0.175\textwidth}
  \hspace{0cm} \vspace{0cm}
\end{minipage} \begin{minipage}[t]{0.8\textwidth}
\large \textbf{ABSTRACT}\end{minipage} \\[-6pt]
\rule{\textwidth}{1.5pt} \\
\begin{minipage}[t]{0.175\textwidth}
  \vspace{1cm} Received \\ <<RECEIVED>> \\[0.5cm] Accepted \\ <<ACCEPTED>> \\[0.5cm] Published \\ <<PUBLISHED>> \\[0.25cm]
\end{minipage} \colorbox{colorabs}{\begin{minipage}[t]{0.81\textwidth}
  \large{<<ABSTRACT>>
  \\[2pt]
  \normalsize \textbf{Keywords:} <<KEYWORDS>>
  }
\end{minipage}}
\\[-10pt]

\normalsize
\noindent \rule{\textwidth}{0.75pt}
<<BODY>>

\end{document}
"""

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


def build_tex(art, writer, opts):
    # yazarlar
    au_tex = []
    for au in art.authors:
        s = tex_escape(au["name"])
        if au["marks"]:
            s += r" \textsuperscript{%s}" % ",".join(au["marks"])
        if au.get("orcid"):
            s += r"\orcidlink{%s}" % url_escape(au["orcid"])
        au_tex.append(s)
    if len(au_tex) > 2:
        authors = ", ".join(au_tex[:-1]) + ", and " + au_tex[-1]
    else:
        authors = " and ".join(au_tex)
    if len(art.affiliations) > 1:
        affs = "\n".join(r"\large \textsuperscript{%d}%s\\" % (k + 1, tex_escape(a))
                         for k, a in enumerate(art.affiliations))
    else:
        affs = "\n".join(r"\large %s\\" % tex_escape(a) for a in art.affiliations)
    corr = r"\textsuperscript{*}\textit{Corresponding author, E-mail:}"
    if art.corr_email:
        corr += r"\href{mailto:%s}{%s}" % (art.corr_email, tex_escape(art.corr_email))
    if art.corr_tel:
        corr += r" (Tel: %s)" % tex_escape(art.corr_tel)

    surnames = [au["name"].split()[-1] for au in art.authors if au["name"].split()]
    if len(surnames) >= 3:
        runhead = tex_escape(surnames[0]) + r" \textit{et al.}"
    elif len(surnames) == 2:
        runhead = r"%s \& %s" % (tex_escape(surnames[0]), tex_escape(surnames[1]))
    else:
        runhead = tex_escape(surnames[0]) if surnames else ""

    abstract = r" \par ".join(art.abstract)
    kw = render_plain(art.keywords.strip())
    if kw and kw[-1] not in ".":
        kw += "."

    doi = "https://doi.org/10.25271/sjuoz.%s.%s.%s.%s" % (opts.year, opts.volume, opts.issue, opts.article_id)
    last = str(opts.last_page) if opts.last_page else r"\pageref*{LastPage}"
    rep = {
        "MAINFONT": opts.main_font,
        "HEADFONT": opts.header_font,
        "DOI": doi,
        "RUNHEAD": runhead,
        "VOL": str(opts.volume),
        "ISSUE2": "%02d" % int(opts.issue),
        "YEAR": str(opts.year),
        "MONTH": opts.month,
        "MON3": opts.month[:3].upper(),
        "FIRSTPAGE": str(opts.first_page),
        "LASTPAGE": last,
        "ARTICLEURL": "https://sjuoz.uoz.edu.krd/index.php/sjuoz/article/view/%s" % opts.article_id,
        "ARTICLETYPE": tex_escape(opts.article_type or art.article_type),
        "TITLE": tex_escape(art.title.upper() if not opts.keep_title_case else art.title),
        "AUTHORS": authors,
        "AFFILIATIONS": affs,
        "CORRESPONDING": corr,
        "RECEIVED": tex_escape(opts.received),
        "ACCEPTED": tex_escape(opts.accepted),
        "PUBLISHED": tex_escape(opts.published),
        "ABSTRACT": abstract,
        "KEYWORDS": kw,
        "BODY": "\n".join(writer.out),
    }
    tex = TEMPLATE
    for k, v in rep.items():
        tex = tex.replace("<<%s>>" % k, v)
    return tex


# --------------------------------------------------------------------------
# Ana program
# --------------------------------------------------------------------------
def classify(paths):
    """Verilen docx dosyalarından hangisinin kapak mektubu olduğunu bulur."""
    manuscript, cover = None, None
    for p in paths:
        d = Docx(p)
        text = " ".join(b.text() for b in d.body_blocks()[:80] if isinstance(b, Para))
        is_cover = bool(re.search(r"\bdear\b", text, re.I) and re.search(r"\bname\s*:", text, re.I)) \
            or "cover" in os.path.basename(p).lower() and len(text) < 8000
        if is_cover and cover is None:
            cover = d
        elif manuscript is None:
            manuscript = d
        else:
            cover = cover or d
    return manuscript, cover


def main(argv=None):
    today = datetime.date.today()
    ap = argparse.ArgumentParser(
        description="SJUOZ makale + kapak mektubu Word dosyalarını tek bir XeLaTeX dosyasına çevirir.")
    ap.add_argument("docx", nargs="*", help="makale .docx (ve isteğe bağlı kapak mektubu .docx); "
                    "verilmezse dosya seçme penceresi açılır")
    ap.add_argument("-o", "--outdir", default=None, help="çıktı klasörü (varsayılan: makale no)")
    ap.add_argument("--article-id", help="makale numarası (varsayılan: dosya adındaki sayı)")
    ap.add_argument("--volume", type=int, help="cilt (varsayılan: yıl-2012)")
    ap.add_argument("--issue", type=int, help="sayı (varsayılan: ayın çeyreği)")
    ap.add_argument("--year", type=int, default=today.year)
    ap.add_argument("--month", default=None, help="ay adı, ör. October (varsayılan: bu ay)")
    ap.add_argument("--first-page", type=int, default=1)
    ap.add_argument("--last-page", type=int, default=None,
                    help="son sayfa (verilmezse LaTeX otomatik hesaplar)")
    ap.add_argument("--received", default="XX")
    ap.add_argument("--accepted", default="XX")
    ap.add_argument("--published", default="XX")
    ap.add_argument("--article-type", default=None, help="ör. 'Original Article', 'Review Article'")
    ap.add_argument("--main-font", default="Times New Roman")
    ap.add_argument("--header-font", default="Garamond")
    ap.add_argument("--keep-heading-case", action="store_true", help="bölüm başlıklarını BÜYÜK harfe çevirme")
    ap.add_argument("--keep-title-case", action="store_true", help="makale başlığını BÜYÜK harfe çevirme")
    ap.add_argument("--strip-ref-numbers", action="store_true", help="kaynakların başındaki [1] numaralarını sil")
    ap.add_argument("--assets", default=None, help="SJUOZ.png ve orcid.png'nin bulunduğu klasör")
    ap.add_argument("--tex-name", default=None, help="çıktı .tex dosya adı (varsayılan: <makale no>.tex)")
    ap.add_argument("--wide-figs", default="", help="iki sütunu kaplayacak şekiller: '3,5' veya 'all'")
    ap.add_argument("--no-links", action="store_true", help="atıf/şekil/tablo bağlantılarını ekleme")
    ap.add_argument("--compile", action="store_true", help="xelatex ile PDF de üret (iki kez çalıştırır)")
    ap.add_argument("--zip", action="store_true", help="Overleaf'e yüklemek için <makale no>_latex.zip üret")
    opts = ap.parse_args(argv)
    opts.wide_figs = {"all"} if opts.wide_figs.strip() == "all" else \
        {x.strip() for x in opts.wide_figs.split(",") if x.strip()}

    if not opts.docx:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            opts.docx = list(filedialog.askopenfilenames(
                title="Makale ve kapak mektubu (.docx) seçin", filetypes=[("Word", "*.docx")]))
        except Exception:  # noqa
            pass
        if not opts.docx:
            ap.error("en az bir .docx dosyası verin")

    for p in opts.docx:
        if not os.path.isfile(p):
            ap.error("dosya bulunamadı: %s" % p)
    manuscript, cover = classify(opts.docx)
    if manuscript is None:
        ap.error("makale dosyası bulunamadı")

    if not opts.article_id:
        for p in opts.docx:
            mt = re.search(r"(\d{3,})", os.path.basename(p))
            if mt:
                opts.article_id = mt.group(1)
                break
        else:
            opts.article_id = "0000"
    if opts.month is None:
        opts.month = MONTHS[today.month - 1]
    month_no = next((k + 1 for k, mname in enumerate(MONTHS) if mname.lower().startswith(opts.month.lower()[:3])), today.month)
    opts.month = MONTHS[month_no - 1]
    if opts.volume is None:
        opts.volume = opts.year - 2012
        print("Bilgi: cilt verilmedi, %d kullanıldı (--volume)." % opts.volume)
    if opts.issue is None:
        opts.issue = (month_no - 1) // 3 + 1
        print("Bilgi: sayı verilmedi, %d kullanıldı (--issue)." % opts.issue)

    outdir = opts.outdir or opts.article_id
    os.makedirs(outdir, exist_ok=True)

    cl = parse_cover_letter(cover) if cover else None
    if cover is None:
        print("Bilgi: kapak mektubu verilmedi; ORCID bilgileri eklenmeyecek.")
    art = parse_article(manuscript, cl)

    writer = Writer(art, manuscript, outdir, opts)
    global LINKER
    if not opts.no_links:
        LINKER = Linker(art.references, *writer.prescan())
    writer.run()
    writer.references()
    writer.close_cols()
    tex = build_tex(art, writer, opts)

    tex_path = os.path.join(outdir, opts.tex_name or "%s.tex" % opts.article_id)
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex)

    if LINKER is not None:
        for m_ in sorted(LINKER.missing):
            warn("Atıf kaynakçada bulunamadı: (%s)" % m_)
        for r in LINKER.refs:
            if r["key"] not in LINKER.used:
                warn("Metinde atıf yapılmamış kaynak: %s" % r["label"])
        print("Bilgi: %d / %d kaynağa metinden bağlantı verildi." % (len(LINKER.used), len(LINKER.refs)))

    # logolar: --assets, makalenin klasörü, programın klasörü veya çalışma klasörü
    bases = [opts.assets] if opts.assets else []
    bases += [os.path.dirname(os.path.abspath(manuscript.path)),
              os.path.dirname(os.path.abspath(__file__)), os.getcwd()]
    for fn in ("SJUOZ.png", "orcid.png"):
        dst = os.path.join(outdir, fn)
        for base in bases:
            src = os.path.join(base, fn)
            if os.path.isfile(src) and os.path.abspath(src) != os.path.abspath(dst):
                shutil.copy(src, dst)
                break
        if not os.path.isfile(dst) and fn == "SJUOZ.png":
            warn("SJUOZ.png bulunamadı; logo yerine yazı gösterilecek (--assets ile verin).")

    print()
    print("Oluşturuldu: %s" % tex_path)
    print("  Başlık      : %s" % art.title)
    print("  Yazarlar    : %s" % "; ".join("%s [%s]%s" % (a["name"], ",".join(a["marks"]),
                                                          " ORCID" if a.get("orcid") else "") for a in art.authors))
    print("  Affiliasyon : %d   Şekil: %d   Tablo: %d   Denklem: %d   Kaynak: %d"
          % (len(art.affiliations), writer.fig_count, writer.tab_count, writer.eq_count, len(art.references)))
    tex_name = os.path.basename(tex_path)
    pdf_path = os.path.splitext(tex_path)[0] + ".pdf"
    if opts.compile:
        exe = shutil.which("xelatex")
        if not exe:
            warn("xelatex bulunamadı; derleme atlandı (MiKTeX / TeX Live kurun).")
        else:
            import subprocess
            ok = True
            for _ in range(2):
                r = subprocess.run([exe, "-interaction=nonstopmode", tex_name], cwd=outdir,
                                   capture_output=True, text=True, errors="replace")
                ok = r.returncode == 0
            if ok and os.path.isfile(pdf_path):
                print("PDF üretildi: %s" % pdf_path)
            else:
                warn("xelatex hata verdi; ayrıntılar: %s" % (os.path.splitext(tex_path)[0] + ".log"))
    else:
        print("Derlemek için: cd %s && xelatex %s && xelatex %s" % (outdir, tex_name, tex_name))
    if opts.zip:
        zp = os.path.join(os.path.dirname(os.path.abspath(outdir)), "%s_latex.zip" % opts.article_id)
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            for fn in sorted(os.listdir(outdir)):
                if fn == tex_name or fn in ("SJUOZ.png", "orcid.png") or re.match(r"Figure\d+\.\w+$", fn) \
                        or fn == os.path.basename(pdf_path):
                    z.write(os.path.join(outdir, fn), fn)
        print("Zip (Overleaf'e yüklenebilir): %s" % zp)
    if WARNINGS:
        print("\n%d uyarı var (yukarıya bakın)." % len(WARNINGS))
    return 0


if __name__ == "__main__":
    sys.exit(main())
