# -*- coding: utf-8 -*-
"""Uzun denklemlerin iki sütunlu sayfada sütuna sığdırılması.

Python, denklemin en üst düzeydeki ilişki (=, ≤ …) ve toplama/çıkarma işaretlerinden bölünmüş
2 ve 3 satırlık biçimlerini hazırlar; hangisinin sütuna sığdığına derleme sırasında LaTeX karar verir
(\\sjuozfit, template.py). Hiçbiri sığmazsa denklem sütun genişliğine küçültülür: taşma olmaz.
Parantez, \\left…\\right, {…} ve ortam (\\begin…\\end) içinden bölünmez.
"""
import re

RELATIONS = ("=", "<", ">", r"\le", r"\leq", r"\ge", r"\geq", r"\neq", r"\ne", r"\approx", r"\equiv",
             r"\sim", r"\simeq", r"\cong", r"\propto", r"\to", r"\rightarrow", r"\Rightarrow",
             r"\Leftrightarrow", r"\leftarrow", r"\mapsto", r"\ll", r"\gg")
BINARY = ("+", "-", r"\pm", r"\mp")
_CMD = re.compile(r"\\[A-Za-z]+")
_OPENERS = "([{"


def _tokens(s):
    """(konum, simge) dizisi: komutlar bütün, diğerleri tek karakter."""
    out, i = [], 0
    while i < len(s):
        m = _CMD.match(s, i)
        if m:
            out.append((i, m.group(0)))
            i = m.end()
        elif s[i] == "\\" and i + 1 < len(s):
            out.append((i, s[i:i + 2]))  # \{ \, \; …
            i += 2
        else:
            out.append((i, s[i]))
            i += 1
    return out


def break_points(latex):
    """En üst düzeydeki bölme yerleri: [(konum, tür)] — tür 'rel' ya da 'bin'."""
    toks = _tokens(latex)
    depth = 0
    points = []
    prev = None
    for k, (pos, t) in enumerate(toks):
        if t in ("{", r"\left", r"\begin"):
            depth += 1
        elif t in ("}", r"\right", r"\end"):
            depth -= 1
        elif t in "()[]" and len(t) == 1:
            depth += 1 if t in "([" else -1
        elif depth == 0 and t in RELATIONS:
            # "<=" gibi iki simgeli ilişkide ikinci simgeden bölme
            if not (prev and prev[1] in RELATIONS and prev[0] + len(prev[1]) == pos):
                points.append((pos, "rel"))
        elif depth == 0 and t in BINARY and prev is not None:
            # tekli eksi değil: önünde ilişki, açılış ya da virgül yoksa
            p = prev[1]
            if p not in RELATIONS and p not in BINARY and p not in _OPENERS and p not in (",", "&", r"\\"):
                points.append((pos, "bin"))
        if not t.isspace():
            prev = (pos, t)
    return points


def _lines(latex, cuts):
    """Bölme konumlarından aligned satırları üretir."""
    parts, last = [], 0
    for pos, _ in cuts:
        parts.append(latex[last:pos].strip())
        last = pos
    parts.append(latex[last:].strip())
    return parts


def candidate(latex, n):
    """n satıra bölünmüş aligned gövdesi; bölünemiyorsa None."""
    pts = break_points(latex)
    if not pts or n < 2:
        return None
    rels = [p for p in pts if p[1] == "rel"]
    first_rel = rels[0][0] if rels else None
    # ilk ilişkiden önce bölünmez: sol taraf ilk satırda kalır ("y = …" hizası)
    usable = [p for p in pts if first_rel is None or p[0] > first_rel]
    if not usable:
        return None
    total = len(latex)
    # eşit uzunlukta satırlar için her hedef konuma en yakın bölme yeri
    cuts = []
    for k in range(1, n):
        target = total * k / n
        best = min(usable, key=lambda p: (abs(p[0] - target), p[1] != "rel"))
        if best not in cuts and (not cuts or best[0] > cuts[-1][0]):
            cuts.append(best)
    if not cuts:
        return None
    lines = _lines(latex, cuts)
    if first_rel is not None:
        # "lhs &= rest" — sonraki satırlar ilişkinin hizasında, toplam satırları biraz içeride
        head = latex[:first_rel].strip()
        lines[0] = head + " &" + lines[0][len(head):].lstrip() if lines[0].startswith(head) else "&" + lines[0]
        rest = ["&" + ("" if kind == "rel" else r"\quad ") + ln for ln, (_, kind) in zip(lines[1:], cuts)]
    else:
        lines[0] = "&" + lines[0]
        rest = [r"&\quad " + ln for ln in lines[1:]]
    if any(not ln.strip("& ") for ln in [lines[0]] + rest):
        return None
    return r"\begin{aligned}" + r" \\ ".join([lines[0]] + rest) + r"\end{aligned}"


def fit(latex):
    """Denklemi sütuna sığdıran LaTeX: \\sjuozfit{tam}{2 satır}{3 satır} ya da yalnız küçültme."""
    two, three = candidate(latex, 2), candidate(latex, 3)
    if not two:
        return r"\sjuozfitone{%s}" % latex
    return r"\sjuozfit{%s}{%s}{%s}" % (latex, two, three or two)
