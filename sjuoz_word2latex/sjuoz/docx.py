# -*- coding: utf-8 -*-
"""Word (.docx) okuyucu: paragraflar, koşular, tablolar, bölümler, dipnotlar."""
import posixpath
import re
import zipfile

from lxml import etree

from .common import (C_NS, M, MC, NS, R, W, on_off, wval)
from .omml import OMML
from .render import render_tokens
from .texinput import restore_math

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
        self.notes = {"footnote": self._notes("word/footnotes.xml", "footnote"),
                      "endnote": self._notes("word/endnotes.xml", "endnote")}

    def _notes(self, name, tag):
        """Dipnot/sonnot id -> paragraf elemanları."""
        res = {}
        if name in self.z.namelist():
            root = etree.fromstring(self.z.read(name))
            for fn in root.findall(W + tag):
                res[fn.get(W + "id")] = fn.findall(W + "p")
        return res

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
        if rid.startswith("chart:"):
            rid = rid[6:]
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
        # Word bölümleri: bir bölümün sayfa yönü, bölümü bitiren sectPr'de yazar
        start = 0
        for k, b in enumerate(blocks):
            sect = b.el.find("%spPr/%ssectPr" % (W, W)) if isinstance(b, Para) else None
            if sect is not None:
                for x in blocks[start:k + 1]:
                    x.landscape = is_landscape(sect)
                start = k + 1
        final = is_landscape(body.find(W + "sectPr"))
        for x in blocks[start:]:
            x.landscape = final
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


def is_landscape(sect):
    """Word bölüm ayarı (sectPr) yatay sayfa mı?"""
    if sect is None:
        return False
    pg = sect.find(W + "pgSz")
    if pg is None:
        return False
    if pg.get(W + "orient") == "landscape":
        return True
    try:
        return int(pg.get(W + "w", "0")) > int(pg.get(W + "h", "0")) > 0
    except ValueError:
        return False


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
        self.bidi = ppr is not None and on_off(ppr.find(W + "bidi"))
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
                    latex, num = restore_math(self.doc.path, *self._omml.convert(o))
                    self.tokens.append(("m", latex, True, num))
            elif tag == M + "oMath":
                latex, num = restore_math(self.doc.path, *self._omml.convert(c))
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
                kind = "footnote" if tag == W + "footnoteReference" else "endnote"
                note = " ".join(render_tokens(Para(e, self.doc).tokens)
                                for e in self.doc.notes[kind].get(c.get(W + "id"), []))
                self.tokens.append(("fn", kind, c.get(W + "id"), note.strip()))

    def _images(self, el):
        rids = el.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS)
        for rid in rids:
            self.tokens.append(("img", rid))
        # Word grafikleri (chart): resim olarak çizilip eklenir
        for rid in el.xpath(".//c:chart/@r:id", namespaces={"c": C_NS, "r": NS["r"]}):
            self.tokens.append(("img", "chart:" + rid))

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
        # sondaki kalın yazılmamış ":" / "." başlığı bozmaz ("Nitrogen percentage (N%):")
        ts = [t for t in self.tokens if t[0] == "t" and t[1].strip() and not re.fullmatch(r"[\s:.;]+", t[1])]
        return bool(ts) and all(t[2][0] for t in ts)

    def heading_level(self):
        if self.style.lower().startswith("heading"):
            if len(self.text().strip()) > 200:
                return None   # başlık biçemi verilmiş gövde paragrafı
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
