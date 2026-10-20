# -*- coding: utf-8 -*-
"""Biçimli metni LaTeX'e çevirme; atıf ve çapraz başvuru bağlantıları."""
import re
import unicodedata


from .common import (DOI_DUP_RE, URL_RE, render_plain, title_case_caps, url_escape)

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
                # "doi: 10.xxxx/yyy" -> bağlantı
                txt = re.sub(r"\b(?:doi|DOI)\s*:\s*(10\.\d{4,9}/\S+)", r"https://doi.org/\1", txt)
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
        elif t[0] == "fn" and len(t) > 3:
            out.append(r"\footnote{%s}" % t[3])
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
