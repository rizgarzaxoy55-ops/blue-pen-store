# -*- coding: utf-8 -*-
"""Word denklemleri (OMML) -> LaTeX."""
import re
import unicodedata


from .common import (M, MC, W, local)

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
          "‖": r"\|", "∥": r"\|", "⟨": r"\langle", "⟩": r"\rangle", "〈": r"\langle", "〉": r"\rangle",
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
        if number is None:
            # Word'de numara parantez nesnesi olarak yazılmış olabilir: ... \left( 1 \right)
            mt = re.search(r"%s?\s*\\left\(\s*(\d+(?:\.\d+)*[a-z]?)\s*\\right\)\s*$" % EQNUM_MARK, latex)
            if mt and latex[:mt.start()].strip():
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
        # \left/\right yalnız bilinen ayraçları kabul eder; bilinmeyen karakter düz yazılır
        lb, rb = DELIMS.get(beg), DELIMS.get(end)
        if lb is None:
            lb, inner = ".", beg + " " + inner
        if rb is None:
            rb, inner = ".", inner + " " + end
        return r"\left%s %s \right%s" % (lb, inner, rb)

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
