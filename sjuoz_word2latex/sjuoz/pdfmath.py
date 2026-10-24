# -*- coding: utf-8 -*-
"""LaTeX ile üretilmiş PDF'lerdeki numaralı satır denklemlerini yeniden LaTeX'e çevirir.

PDF'te denklem bilgisi yoktur, ama LaTeX'in yazı tipleri ne olduğunu söyler: CMMI (matematik italik:
değişkenler, Yunan harfleri), CMR (düz: rakam, işaret, parantez), CMSY/CMEX (sembol: ≤ ∼ ∑),
metin yazı tipi (\\text{…} içindeki kelimeler). Karakterin boyutu ve taban çizgisi alt/üst simgeyi,
yatay bir çizginin üstü ve altı kesri verir. Satır sonunda "(17)" gibi numarası olan satırlar denklemdir.
Kurulamayan (çok satırlı, büyük operatör sınırlı…) denklemler düz metin kalır ve sayısı bildirilir.
"""
import re

GREEK = {"α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta", "ϵ": r"\epsilon",
         "ε": r"\varepsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta", "ϑ": r"\vartheta", "ι": r"\iota",
         "κ": r"\kappa", "λ": r"\lambda", "µ": r"\mu", "μ": r"\mu", "ν": r"\nu", "ξ": r"\xi", "π": r"\pi",
         "ρ": r"\rho", "σ": r"\sigma", "ς": r"\varsigma", "τ": r"\tau", "υ": r"\upsilon", "φ": r"\varphi",
         "ϕ": r"\phi", "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega", "Γ": r"\Gamma", "∆": r"\Delta",
         "Δ": r"\Delta", "Θ": r"\Theta", "Λ": r"\Lambda", "Ξ": r"\Xi", "Π": r"\Pi", "Σ": r"\Sigma",
         "Φ": r"\Phi", "Ψ": r"\Psi", "Ω": r"\Omega"}
SYMBOLS = {"≤": r"\leq", "≥": r"\geq", "≠": r"\neq", "∼": r"\sim", "≈": r"\approx", "≡": r"\equiv",
           "−": "-", "·": r"\cdot", "×": r"\times", "÷": r"\div", "±": r"\pm", "∓": r"\mp", "∈": r"\in",
           "∉": r"\notin", "⊂": r"\subset", "⊆": r"\subseteq", "∪": r"\cup", "∩": r"\cap", "∞": r"\infty",
           "∂": r"\partial", "∇": r"\nabla", "→": r"\to", "←": r"\leftarrow", "⇒": r"\Rightarrow",
           "⇔": r"\Leftrightarrow", "∀": r"\forall", "∃": r"\exists", "∝": r"\propto", "′": "'",
           "∑": r"\sum", "∏": r"\prod", "∫": r"\int", "√": r"\sqrt", "…": r"\dots", "⋯": r"\cdots",
           "{": r"\{", "}": r"\}", "%": r"\%", "#": r"\#", "&": r"\&", "_": r"\_", "∗": "*"}
EQ_NUMBER = re.compile(r"\((\d+[a-z]?)\)$")


def _math_font(font):
    return bool(re.match(r"(CMMI|CMR|CMSY|CMEX|CMBX|CMSL|MSAM|MSBM|LMMath|LMRoman\d+-Regular|"
                         r"Cambria\s*Math|STIX|TeXGyre.*Math|LatinModernMath)", font, re.I))


def _is_text_font(font):
    return not _math_font(font)


def page_glyphs(page):
    """Karakterler: dict(x0, y0, x1, y1, c, size, base, font)."""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                for ch in sp["chars"]:
                    if ch["c"].strip():
                        x0, y0, x1, y1 = ch["bbox"]
                        out.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1, "c": ch["c"], "size": sp["size"],
                                    "base": ch["origin"][1], "font": sp["font"]})
    return out


def _lines(glyphs):
    """Ana yazı boyutundaki karakterlerden satırlar; küçük karakterler (alt/üst simge) sonra eklenir."""
    if not glyphs:
        return []
    big = max(g["size"] for g in glyphs)
    main = [g for g in glyphs if g["size"] >= 0.8 * big or g["size"] >= 8]
    lines = []
    for g in sorted(main, key=lambda g: g["base"]):
        for ln in lines:
            if abs(ln["base"] - g["base"]) < 0.35 * g["size"]:
                ln["glyphs"].append(g)
                break
        else:
            lines.append({"base": g["base"], "glyphs": [g], "size": g["size"]})
    return lines


def find_equations(page):
    """Sayfadaki numaralı satır denklemleri: [(numara, latex, düz_metin_anahtarı)]."""
    glyphs = page_glyphs(page)
    rules = _horizontal_rules(page)
    out = []
    for ln in _lines(glyphs):
        gs = sorted(ln["glyphs"], key=lambda g: g["x0"])
        text = "".join(g["c"] for g in gs)
        mt = EQ_NUMBER.search(text)
        if not mt or len(gs) <= len(mt.group(0)) + 1:
            continue
        k = len(gs) - len(mt.group(0))
        num_x = gs[k]["x0"]
        body = gs[:k]
        if num_x - body[-1]["x1"] < 15:   # numara metne yapışık: denklem numarası değil
            continue
        if not any(_math_font(g["font"]) for g in body):
            continue
        size = max(g["size"] for g in body)
        top = min(g["y0"] for g in body) - 0.45 * size
        bottom = max(g["y1"] for g in body) + 0.45 * size
        x0, x1 = body[0]["x0"] - 2, num_x - 2  # kesir "=" ile numara arasında kalabilir
        base = _mode([g["base"] for g in body])
        # küçük karakterler (alt/üst simge) ve kesirlerin pay/paydası
        extra = [g for g in glyphs if g not in ln["glyphs"] and x0 <= (g["x0"] + g["x1"]) / 2 <= x1
                 and top - 1.2 * size <= (g["y0"] + g["y1"]) / 2 <= bottom + 1.2 * size]
        latex, complete = build_latex(body, extra, base, size, rules)
        if not complete:
            continue
        out.append((mt.group(1), latex, re.sub(r"\s+", "", "".join(g["c"] for g in body))))
    return out


def _mode(values):
    vals = sorted(round(v, 1) for v in values)
    return max(set(vals), key=vals.count)


def _horizontal_rules(page):
    rules = []
    try:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].y - it[2].y) < 0.6 and abs(it[2].x - it[1].x) > 3:
                    rules.append((min(it[1].x, it[2].x), it[1].y, max(it[1].x, it[2].x)))
                elif it[0] == "re" and it[1].height < 1.2 and it[1].width > 3:
                    rules.append((it[1].x0, (it[1].y0 + it[1].y1) / 2, it[1].x1))
    except Exception:  # noqa
        pass
    return rules


def _char_tex(g):
    c = g["c"]
    if c in GREEK:
        return GREEK[c] + " "
    if c in SYMBOLS:
        s = SYMBOLS[c]
        return s + " " if s.startswith("\\") and s[-1].isalpha() else s
    return c


def _seq_tex(glyphs, base, size):
    """Bir satır parçası (soldan sağa) → LaTeX; alt/üst simgeler ve \\text{} kelimeleri."""
    out, text_run, script = [], [], None
    prev = None

    def flush_text():
        if text_run:
            # \text{} içindeki baş/son boşluk anlamlıdır: "150,000\text{ hectares}", "\text{Maximize }Z"
            words = re.sub(r"\s+", " ", "".join(text_run))
            if out and re.search(r"(=|\\[a-z]+ |[+\-(,{])$", out[-1]):
                words = words.lstrip()   # ilişki/işaretten sonra LaTeX zaten boşluk bırakır
            if words.strip():
                out.append(r"\text{%s}" % words.replace("\\", "").replace("{", r"\{").replace("}", r"\}"))
            text_run.clear()

    for g in glyphs:
        small = g["size"] < 0.85 * size
        kind = None
        if small and g["base"] > base + 0.08 * size:
            kind = "_"
        elif small and g["base"] < base - 0.15 * size:
            kind = "^"
        elif not small and g["base"] < base - 0.3 * size:
            kind = "^"
        if kind is None and _is_text_font(g["font"]) and (g["c"].isalpha() or (text_run and g["c"] in "()-/,.:'")):
            if script:
                out.append("}")
                script = None
            if text_run and prev is not None and g["x0"] - prev["x1"] > 0.18 * g["size"]:
                text_run.append(" ")
            elif not text_run and prev is not None and g["x0"] - prev["x1"] > 0.18 * g["size"]:
                text_run.append(" ")
            text_run.append(g["c"])
            prev = g
            continue
        if text_run:
            if prev is not None and g["x0"] - prev["x1"] > 0.18 * g["size"]:
                text_run.append(" ")
            flush_text()
        if kind != script:
            if script:
                out.append("}")
            if kind:
                out.append(kind + "{")
            script = kind
        out.append(_char_tex(g))
        prev = g
    flush_text()
    if script:
        out.append("}")
    tex = "".join(out)
    tex = tex.replace("...", r"\dots ").replace(". . .", r"\dots ")
    return re.sub(r"\s+", " ", tex).strip()


def build_latex(body, extra, base, size, rules):
    """(latex, tam_mı): ana satır + alt/üst simgeler + kesirler."""
    top_band = min(g["y0"] for g in body)
    bot_band = max(g["y1"] for g in body)
    # kesir: payı ve paydası (alt simgeleriyle) bir yatay çizginin üstünde ve altında
    fracs = []
    for x0, y, x1 in rules:
        if not top_band - 1.5 * size < y < bot_band + 1.5 * size:
            continue
        inside = [g for g in extra if x0 - 1 <= (g["x0"] + g["x1"]) / 2 <= x1 + 1]
        num = [g for g in inside if g["y1"] <= y + 1]
        den = [g for g in inside if g["y0"] >= y - 1]
        if num and den:
            fracs.append((x0, x1, num, den))
    used = {id(g) for _, _, n, d in fracs for g in n + d}
    scripts = []
    for g in extra:
        if id(g) in used:
            continue
        cy = (g["y0"] + g["y1"]) / 2
        if g["size"] < 0.85 * size and top_band - 0.6 * size <= cy <= bot_band + 0.6 * size:
            scripts.append(g)          # alt/üst simge
        elif g["size"] >= 0.85 * size:
            return None, False         # büyük operatör sınırları, matris vb.: güvenle kurulamaz
    seq = sorted(body + scripts, key=lambda g: g["x0"])
    parts, k = [], 0
    for x0, x1, num, den in sorted(fracs, key=lambda f: f[0]):
        left = [g for g in seq[k:] if g["x1"] <= x0 + 1]
        parts.append(_seq_tex(left, base, size))
        k += len(left)
        nb = _mode([g["base"] for g in num if g["size"] >= 0.6 * size] or [g["base"] for g in num])
        db = _mode([g["base"] for g in den if g["size"] >= 0.6 * size] or [g["base"] for g in den])
        ns = max(g["size"] for g in num)
        ds = max(g["size"] for g in den)
        parts.append(r"\frac{%s}{%s}" % (_seq_tex(sorted(num, key=lambda g: g["x0"]), nb, ns),
                                         _seq_tex(sorted(den, key=lambda g: g["x0"]), db, ds)))
    parts.append(_seq_tex(seq[k:], base, size))
    latex = " ".join(p for p in parts if p)
    return latex, bool(latex)
