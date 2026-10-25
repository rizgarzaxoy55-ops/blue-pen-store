# -*- coding: utf-8 -*-
"""PDF girişi: PDF'i Word'e (.docx) çevirip dönüştürücünün okuyabileceği biçime toparlar.

PDF'te paragraf, başlık, denklem bilgisi yoktur; yalnızca sayfadaki satırlar vardır. Bu yüzden
sonuç yazarın Word dosyası kadar iyi olmaz: denklemler resim/düz metin olarak gelir, biçim kaybolabilir.
Çeviri pdf2docx ile yapılır (pip install pdf2docx), ardından:
  * satır/sayfa sonunda bölünmüş paragraflar birleştirilir,
  * kaynakça yeniden girdilere ayrılır (PDF'te girdiler birbirine yapışık ya da bölük gelir),
  * kaybolan kelime boşlukları PDF'in kelime sırasından geri konur,
  * numaralı denklemler LaTeX'e (pdfmath), yalnız yatay çizgili tablolar Word tablosuna geri kurulur,
  * Symbol yazı tipindeki madde işaretleri düzeltilir.
"""
import copy
import os
import re
import tempfile
import zipfile

from lxml import etree

from .common import W, warn
from .parse import ABSTRACT_HEAD, CAPTION_RE, KEYWORDS_HEAD, REF_HEAD, clean_heading, manual_number_level

PDF_NOTE = ("Kaynak dosya PDF'ti ve Word'e çevrilerek okundu: denklemler, tablolar, kalın/italik yazı ve "
            "kaynakça bölünmeleri PDF'le karşılaştırılarak elle kontrol edilmeli. Mümkünse yazardan Word dosyasını isteyin.")

# Kaynakça girdisinin başı: "[3] …", "3. …", "Safi, M. E.", "Van Hai, V.", "Al-Ghafari, A. B.", "MathWorks. (n.d.)"
REF_START = re.compile(
    r"^\s*(?:\[\d+\]|\d+\.\s"
    r"|(?:(?:[Vv]an|[Vv]on|[Dd]e|[Dd]el|[Dd]er|[Dd]i|[Dd]a|[Ll]a|[Ll]e|Al|El)\s+)?"
    r"[A-ZÀ-ÖØ-ÞŠŽ][\w'’\-]+(?:[ -][A-ZÀ-ÖØ-ÞŠŽ][\w'’\-]+)*,\s+(?:[A-ZÀ-ÖØ-ÞŠŽ]\.|[A-ZÀ-ÖØ-ÞŠŽ][a-z])"
    r"|[A-Z][\w&\- ]+?\.?\s\((?:\d{4}[a-z]?|n\.d\.)\))")  # kurum: "World Bank (2014)", "MathWorks. (n.d.)"
# Önceki girdinin bittiği yer: DOI/URL ya da sayfa numarasından sonra nokta
REF_END = re.compile(r"(?:https?://\S+|doi:\s*\S+|\d\.)\s+$", re.I)
SENT_END = (".", "?", "!", ":", ";")
# Paragraf bu sözcüklerle bitemez: "In | Albrecht dataset …" bir bölme yeri değildir
FUNCTION_WORD_END = re.compile(r"\b(?:in|on|at|of|to|for|by|with|from|the|a|an|and|or|as|is|are|was|were)\s*$", re.I)
# Bölünmemesi gereken önek: "Table 2:" etiketi ya da "3.1.1." bölüm numarası (PDF'te ayrı parça olabilir)
CAPTION_LABEL = re.compile(r"^\s*(?:(?:table|figure|fig\.?)\s*\(?\s*\d+(?:\s*[-–]\s*[A-Za-z])?\s*\)?"
                           r"|\d+(?:\.\d+)*)\s*[:.\-–—]?\s*$", re.I)


def is_pdf(path):
    return path.lower().endswith(".pdf")


def ensure_docx(path, workdir=None):
    """PDF ise .docx'e çevirip yolunu döndürür; değilse yolu aynen döndürür."""
    if not is_pdf(path):
        return path
    try:
        from pdf2docx import Converter
    except ImportError:
        raise SystemExit("HATA: PDF okumak için pdf2docx gerekli:  pip install pdf2docx")
    workdir = workdir or tempfile.mkdtemp(prefix="sjuoz_pdf_")
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".docx")
    print("Bilgi: %s PDF'ten Word'e çevriliyor (birkaç saniye sürebilir)..." % os.path.basename(path), flush=True)
    import logging
    logging.getLogger().setLevel(logging.WARNING)  # pdf2docx her sayfa için bilgi satırı yazar
    cv = Converter(path)
    try:
        # Kenarlıksız "tablo" tahmini yazar satırlarını ve kaynakçayı tabloya çeviriyor: yalnız çizgili tablolar
        cv.convert(out, parse_stream_table=False)
    finally:
        cv.close()
    tidy_docx(out, path)
    from .common import WARNINGS
    if PDF_NOTE not in WARNINGS and not re.search(r"cover|kapak", os.path.basename(path), re.I):
        # makale için bir kez; kapak mektubunda denklem/tablo olmadığından gerekmez
        warn(PDF_NOTE)
    if looks_two_column(path):
        warn("%s iki sütunlu dizilmiş bir PDF gibi görünüyor (ör. yayımlanmış dergi sayfası): sütunlar ve "
             "boşluklar karışabilir, sonuç güvenilir değil. Bu yol tek sütunlu yazar taslakları içindir."
             % os.path.basename(path))
    return out


def is_latex_pdf(path):
    """PDF LaTeX ile mi üretilmiş? (Creator/Producer: LaTeX, pdfTeX, XeTeX, LuaTeX, dvipdfmx)"""
    pm = _pymupdf()
    if pm is None:
        return False
    try:
        meta = pm.open(path).metadata or {}
    except Exception:  # noqa
        return False
    info = " ".join(str(meta.get(k) or "") for k in ("creator", "producer"))
    return bool(re.search(r"latex|pdftex|xetex|luatex|dvipdf|miktex|tex live", info, re.I))


def latex_source_for(pdf):
    """LaTeX ile üretilmiş PDF'in yanında (aynı klasörde, ya da oradaki bir zip'te) aynı adlı .tex var mı?"""
    from .inputs import zip_names
    folder = os.path.dirname(os.path.abspath(pdf))
    stem = os.path.splitext(os.path.basename(pdf))[0].lower()
    try:
        names = sorted(os.listdir(folder))
    except OSError:
        return None
    for n in names:
        if n.lower() == stem + ".tex":
            return os.path.join(folder, n)
    for n in names:
        if n.lower().endswith(".zip"):
            zp = os.path.join(folder, n)
            if any(os.path.basename(m).lower() == stem + ".tex" for m in zip_names(zp)):
                return zp
    # PDF'in klasöründe yoksa: çalışılan klasör, Masaüstü, İndirilenler, Belgeler (alt klasörleriyle)
    from .inputs import _nearby_dirs, _walk
    for d in dict.fromkeys([os.getcwd(), folder] + _nearby_dirs(folder)):
        for f in _walk(d, 3):
            low = os.path.basename(f).lower()
            if low == stem + ".tex":
                return f
            if low.endswith(".zip") and any(os.path.basename(m).lower() == stem + ".tex" for m in zip_names(f)):
                return f
    return None


def prefer_latex_source(path):
    """LaTeX ile üretilmiş PDF verildiyse yanındaki .tex kaynağını kullanır; yoksa güçlü uyarı verir."""
    if not is_pdf(path) or not is_latex_pdf(path):
        return path
    src = latex_source_for(path)
    if src:
        print("Bilgi: %s LaTeX ile üretilmiş; PDF yerine LaTeX kaynağı kullanılıyor: %s"
              % (os.path.basename(path), os.path.abspath(src)))
        return src
    warn("%s LaTeX ile üretilmiş bir PDF ve aynı adlı .tex kaynağı (klasörde, Masaüstü/İndirilenler/Belgeler'de "
         "ya da oradaki zip'lerde) bulunamadı. PDF'ten denklemler, tablolar ve yazar numaraları doğru çıkmaz: "
         "yazarın .tex kaynağını (resimler ve .bib ile birlikte .zip) verin." % os.path.basename(path))
    return path


# --------------------------------------------------------------------------
# Tablolar: LaTeX (booktabs) tabloları yalnız yatay çizgilidir; pdf2docx onları düz metne çevirir.
# PyMuPDF'in çizgi tabanlı tablo bulucusu yapıyı bulur, hücre metni karakter konumlarından okunur.
# --------------------------------------------------------------------------
def page_chars(page):
    """(x0, y0, x1, y1, karakter, boyut, taban çizgisi) listesi."""
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        for ln in b.get("lines", []):
            for sp in ln["spans"]:
                for c in sp["chars"]:
                    if c["c"].strip():
                        out.append(tuple(c["bbox"]) + (c["c"], sp["size"], c["origin"][1]))
    return out


def text_in(chars, rect):
    """Dikdörtgendeki metin: satırlara ayrılır, aralıklardan boşluk konur ("25,700" virgülü biraz aşağıda
    durduğu için PyMuPDF'in kendi hücre metni "25700\n," olur)."""
    inside = [c for c in chars if rect[0] - 1 <= (c[0] + c[2]) / 2 <= rect[2] + 1
              and rect[1] - 1 <= (c[1] + c[3]) / 2 <= rect[3] + 1]
    lines = []
    for c in sorted(inside, key=lambda c: c[6]):
        for ln in lines:
            if abs(ln[0] - c[6]) < 0.45 * c[5]:
                ln[1].append(c)
                break
        else:
            lines.append([c[6], [c]])
    out = []
    for _, cs in sorted(lines, key=lambda ln: ln[0]):
        cs.sort(key=lambda c: c[0])
        t, prev = "", None
        for c in cs:
            if prev and c[0] - prev[2] > 0.18 * c[5]:
                t += " "
            t += c[4]
            prev = c
        out.append(t)
    return " ".join(out)


def _hrules(page):
    rules = []
    try:
        for d in page.get_drawings():
            for it in d["items"]:
                if it[0] == "l" and abs(it[1].y - it[2].y) < 0.8 and abs(it[2].x - it[1].x) > 40:
                    rules.append((min(it[1].x, it[2].x), it[1].y, max(it[1].x, it[2].x)))
                elif it[0] == "re" and it[1].height < 1.5 and it[1].width > 40:
                    rules.append((it[1].x0, (it[1].y0 + it[1].y1) / 2, it[1].x1))
    except Exception:  # noqa
        pass
    return sorted(rules, key=lambda r: r[1])


def rule_tables(page, chars):
    """Yalnız yatay çizgili tablolar (LaTeX booktabs / \\hline): aynı genişlikteki çizgiler tablonun
    üstü ve altıdır; satırlar metin satırlarından, sütunlar geniş boşluklardan bulunur."""
    rules = _hrules(page)
    groups = []
    for r in rules:
        for g in groups:
            if abs(g[-1][0] - r[0]) < 3 and abs(g[-1][2] - r[2]) < 3 and r[1] - g[-1][1] < 350:
                g.append(r)
                break
        else:
            groups.append([r])
    out = []
    for g in groups:
        if len(g) < 2:
            continue
        x0, x1, top, bottom = min(r[0] for r in g), max(r[2] for r in g), g[0][1], g[-1][1]
        inside = [c for c in chars if x0 - 2 <= (c[0] + c[2]) / 2 <= x1 + 2 and top < (c[1] + c[3]) / 2 < bottom]
        if not inside:
            continue
        size = max(c[5] for c in inside)
        lines = []
        for c in sorted(inside, key=lambda c: c[6]):
            for ln in lines:
                if abs(ln[0] - c[6]) < 0.5 * size:
                    ln[1].append(c)
                    break
            else:
                lines.append([c[6], [c]])
        rows = []
        for _, cs in sorted(lines, key=lambda ln: ln[0]):
            cs.sort(key=lambda c: c[0])
            segs, cur = [], [cs[0]]
            for c in cs[1:]:
                if c[0] - cur[-1][2] > 0.9 * size:   # hücre aralığı (kelime aralığından geniş)
                    segs.append(cur)
                    cur = [c]
                else:
                    cur.append(c)
            segs.append(cur)
            rows.append([(s[0][0], s[-1][2], text_in(s, (s[0][0], min(c[1] for c in s), s[-1][2],
                                                           max(c[3] for c in s)))) for s in segs])
        if len(rows) < 2:
            continue
        # sütunlar: gövde satırlarındaki parçaların x aralıklarının birleşimi
        cols = []
        for seg in sorted((sg for r in (rows[1:] or rows) for sg in r), key=lambda sg: sg[0]):
            if cols and seg[0] <= cols[-1][1] + 2:
                cols[-1][1] = max(cols[-1][1], seg[1])
            else:
                cols.append([seg[0], seg[1]])
        if len(cols) < 2:
            continue
        table = []
        for r in rows:
            cells = [""] * len(cols)
            for a, b, t in r:
                mid = (a + b) / 2
                k = min(range(len(cols)), key=lambda i: 0 if cols[i][0] - 2 <= mid <= cols[i][1] + 2
                        else min(abs(mid - cols[i][0]), abs(mid - cols[i][1])))
                cells[k] = (cells[k] + " " + t).strip()
            table.append(cells)
        out.append(((x0, top, x1, bottom), table))
    return out


def pdf_tables(path):
    """PDF'teki tablolar (sırayla): [satırlar] — denklem blokları ("… (17)") hariç."""
    pm = _pymupdf()
    if pm is None:
        return []
    out = []
    try:
        doc = pm.open(path)
        for page in doc:
            try:
                found = page.find_tables(strategy="lines").tables
            except Exception:  # noqa
                found = []
            chars = page_chars(page)
            page_tabs = []
            for t in found:
                rows = [[text_in(chars, c) if c else "" for c in r.cells] for r in t.rows]
                page_tabs.append((tuple(t.bbox), rows))
            # yalnız yatay çizgili tablolar (PyMuPDF 1.x bunları bulamaz)
            for bbox, rows in rule_tables(page, chars):
                if not any(bbox[0] < b[2] and b[0] < bbox[2] and bbox[1] < b[3] and b[1] < bbox[3]
                           for b, _ in page_tabs):
                    page_tabs.append((bbox, rows))
            for bbox, rows in sorted(page_tabs, key=lambda t: t[0][1]):
                rows = [r for r in rows if any(x.strip() for x in r)]
                last = [r[-1].strip() for r in rows if r[-1].strip()]
                if len(rows) < 2 or (last and all(re.fullmatch(r"\(\d+\)", x) for x in last)):
                    continue  # numaralı denklemler tablo değildir
                out.append(rows)
        doc.close()
    except Exception:  # noqa
        return []
    return out


def _letters(s):
    return re.sub(r"[^0-9a-zà-ÿµσϵα-ω]", "", s.lower())


def _table_el(rows):
    tbl = etree.Element(W + "tbl")
    pr = etree.SubElement(tbl, W + "tblPr")
    borders = etree.SubElement(pr, W + "tblBorders")
    for side in ("top", "bottom", "insideH"):
        b = etree.SubElement(borders, W + side)
        b.set(W + "val", "single")
        b.set(W + "sz", "4")
    grid = etree.SubElement(tbl, W + "tblGrid")
    ncol = max(len(r) for r in rows)
    for _ in range(ncol):
        etree.SubElement(grid, W + "gridCol")
    for k, row in enumerate(rows):
        tr = etree.SubElement(tbl, W + "tr")
        for cell in row + [""] * (ncol - len(row)):
            tc = etree.SubElement(tr, W + "tc")
            p = etree.SubElement(tc, W + "p")
            r = etree.SubElement(p, W + "r")
            if k == 0:  # başlık satırı kalın
                etree.SubElement(etree.SubElement(r, W + "rPr"), W + "b")
            t = etree.SubElement(r, W + "t")
            t.text = cell
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return tbl


TABLE_CAPTION = re.compile(r"^\s*Table\s*\(?\s*\d+\s*\)?\s*[:.]", re.I)


def rebuild_tables(body, tables):
    """Her "Table N:" başlığından sonra düz metne dönmüş tablo satırlarını gerçek tabloyla değiştirir."""
    if not tables:
        return 0
    used, done = set(), 0
    for cap in [p for p in body.findall(W + "p") if TABLE_CAPTION.match(_text(p))]:
        # tablonun ilk satırı başlıkla aynı paragrafa yapışmış olabilir: "Table 1: … crop Crops Yield …"
        ctext = _text(cap)
        norm, idx = _norm_map(ctext)
        for k, rows in enumerate(tables):
            key = re.sub(r"\s+", "", " ".join(rows[0]))[:15]
            pos = norm.find(key, 8) if k not in used and len(key) >= 4 else -1
            if pos > 0:
                for part in _split(cap, [idx[pos]])[1:]:
                    _strip_lead(part)
                break
        # başlıktan sonraki metinle en iyi eşleşen (henüz kullanılmamış) tablo
        follow, nxt = [], cap.getnext()
        while nxt is not None and len(follow) < 60:
            if nxt.tag == W + "p":
                if TABLE_CAPTION.match(_text(nxt)) or nxt.find(W + "pPr/" + W + "pStyle") is not None:
                    break
                follow.append(nxt)
            elif nxt.tag == W + "tbl":
                break
            nxt = nxt.getnext()
        text_after = _letters(" ".join(_text(p) for p in follow[:6]))
        best = None
        for k, rows in enumerate(tables):
            head = _letters(" ".join(rows[0]))
            if k not in used and head and head[:20] in text_after:
                best = k
                break
        if best is None:
            continue
        used.add(best)
        whole = _letters(" ".join(" ".join(r) for r in tables[best]))
        consumed = []
        for p in follow:
            t = _letters(_text(p))
            if not t:
                consumed.append(p)
                continue
            if t in whole:
                consumed.append(p)
            else:
                break
        for p in consumed:
            body.remove(p)
        cap.addnext(_table_el(tables[best]))
        done += 1
    return done


# --------------------------------------------------------------------------
# Denklemler: PDF'ten yeniden kurulan LaTeX (pdfmath) .docx'te denklem paragrafı olur
# --------------------------------------------------------------------------
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def pdf_equations(path):
    pm = _pymupdf()
    if pm is None:
        return []
    from .pdfmath import find_equations
    out = []
    try:
        doc = pm.open(path)
        for page in doc:
            out += find_equations(page)
        doc.close()
    except Exception:  # noqa
        return out
    return out


def _locate(text, key, num):
    """Paragraf metninde denklemin yeri: (başlangıç, bitiş) ya da None. Anahtar ana boyuttaki
    karakterlerdir; metinde arada alt/üst simge karakterleri (en çok 3) bulunabilir."""
    norm, idx = _norm_map(text)
    tag = "(%s)" % num
    for mt in re.finditer(re.escape(tag), norm):
        j = mt.start() - 1
        k = len(key) - 1
        skipped = 0
        while k >= 0 and j >= 0:
            if norm[j] == key[k]:
                k -= 1
                j -= 1
                skipped = 0
            elif skipped < 3:
                j -= 1
                skipped += 1
            else:
                break
        if k < 0:
            return idx[j + 1], idx[mt.end() - 1] + 1
    return None


def _math_par(like, mark):
    p = _new_para_like(like)
    omp = etree.SubElement(p, "{%s}oMathPara" % M_NS)
    om = etree.SubElement(omp, "{%s}oMath" % M_NS)
    r = etree.SubElement(om, "{%s}r" % M_NS)
    t = etree.SubElement(r, "{%s}t" % M_NS)
    t.text = mark
    return p


def insert_equations(body, equations, docx_path):
    """Bulunan denklemleri paragraflardan ayırıp denklem paragrafına çevirir; kaç tanesi yerleşti."""
    from .texinput import TEX_MATH, math_mark
    table = TEX_MATH.setdefault(os.path.abspath(docx_path), {})
    placed = 0
    paras = list(body.iter(W + "p"))
    start = 0
    for num, latex, key in equations:
        for i in range(start, len(paras)):
            p = paras[i]
            if p.getparent() is None:
                continue
            loc = _locate(_text(p), key, num)
            if not loc and "(%s)" % num in re.sub(r"\s+", "", _text(p)):
                # numara ayrı paragrafa düşmüş olabilir: önceki kardeş paragraflarla birlikte dene
                group = [p]
                prev = p.getprevious()
                while prev is not None and prev.tag == W + "p" and len(group) < 4:
                    group.insert(0, prev)
                    if _locate("".join(_text(x) for x in group), key, num):
                        first = group[0]
                        for x in group[1:]:
                            _append(first, x, space=False)
                            x.getparent().remove(x)
                        p = first
                        loc = _locate(_text(p), key, num)
                        break
                    prev = prev.getprevious()
            if not loc:
                continue
            a, b = loc
            parts = _split(p, [x for x in (a, b) if 0 < x < len(_text(p))])
            # [önce][denklem][sonra] — kenarlarda olmayanlar yoktur
            eq = parts[1] if a > 0 else parts[0]
            n = len(table)
            table[n] = (latex, num)
            new = _math_par(eq, math_mark(n))
            eq.addprevious(new)
            eq.getparent().remove(eq)
            for part in parts:
                if part is not eq and part.getparent() is not None and not _text(part).strip():
                    if not any(x.tag.endswith(("}drawing", "}pict")) for x in part.iter()):
                        part.getparent().remove(part)
            paras = list(body.iter(W + "p"))
            start = paras.index(new) + 1
            placed += 1
            break
    return placed


def pdf_word_pairs(path):
    """PDF'teki ardışık kelime çiftleri ve kelimeler (PyMuPDF boşlukları doğru çıkarır)."""
    pm = _pymupdf()
    pairs, words = set(), set()
    if pm is None:
        return pairs, words
    try:
        doc = pm.open(path)
    except Exception:  # noqa
        return pairs, words
    for page in doc:
        ws = [w[4] for w in page.get_text("words")]
        words.update(ws)
        pairs.update(zip(ws, ws[1:]))
    doc.close()
    return pairs, words


def restore_spaces(body, pairs, words):
    """pdf2docx iki yana yaslı satırlarda kelimeleri ayrı parçalara bölüp boşlukları atıyor
    ("InPapuaNewGuinea"): PDF'te ardışık iki ayrı kelime olan parçaların arasına boşluk koyar."""
    if not pairs:
        return
    for p in body.iter(W + "p"):
        runs = [r for r in p.iter(W + "r") if r.find(W + "t") is not None]
        for a, b in zip(runs, runs[1:]):
            ta = "".join(t.text or "" for t in a.iter(W + "t"))
            tb = "".join(t.text or "" for t in b.iter(W + "t"))
            if not ta.strip() or not tb.strip() or ta[-1].isspace() or tb[0].isspace():
                continue
            left, right = ta.split()[-1], tb.split()[0]
            if (left, right) in pairs and left + right not in words:
                last = list(a.iter(W + "t"))[-1]
                last.text = (last.text or "") + " "
                last.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    # aynı parçanın içinde yapışık kalanlar: "significantlyimprove" → PDF'te "significantly" "improve"
    following = {}
    for x, y in pairs:
        following.setdefault(x, set()).add(y)
    for t in body.iter(W + "t"):
        if t.text and t.text.strip():
            new = " ".join(_split_glued(tok, following, words) for tok in t.text.split(" "))
            if new != t.text:
                t.text = new
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _split_glued(tok, following, words, depth=0):
    """PDF'te kelime olmayan parçayı PDF'te ardışık duran kelimelere böler (en çok 12 kelime)."""
    if not tok or tok in words or len(tok) < 2 or depth > 11:
        return tok
    for k in range(len(tok) - 1, 0, -1):
        a, b = tok[:k], tok[k:]
        nxt = following.get(a)
        if not nxt:
            continue
        if b in nxt:
            return a + " " + b
        if any(b.startswith(y) and len(y) > 1 for y in nxt):
            rest = _split_glued(b, following, words, depth + 1)
            if rest != b:
                return a + " " + rest
    # PDF'te ardışık değiller (ör. satır sonu) ama ikisi de PDF'te ayrı kelime: "significantly|improve"
    if len(tok) >= 10 and tok.isalpha():
        for k in range(len(tok) - 3, 2, -1):
            a, b = tok[:k], tok[k:]
            if a in words and (b in words or _split_glued(b, following, words, depth + 1) != b):
                rest = b if b in words else _split_glued(b, following, words, depth + 1)
                return a + " " + rest
    return tok


def looks_two_column(path):
    """Sayfaların çoğunda metin blokları sayfanın iki yarısında ayrı ayrı duruyorsa iki sütunludur."""
    pymupdf = _pymupdf()
    if pymupdf is None:
        return False
    try:
        doc = pymupdf.open(path)
    except Exception:  # noqa
        return False
    votes = 0
    pages = list(doc)[:6]
    for page in pages:
        mid = page.rect.width / 2
        blocks = [b for b in page.get_text("blocks") if b[6] == 0 and len(b[4].split()) >= 8]
        left = sum(1 for b in blocks if b[2] < mid + 5)
        right = sum(1 for b in blocks if b[0] > mid - 5)
        if left >= 2 and right >= 2:
            votes += 1
    doc.close()
    return votes * 2 > len(pages)


# --------------------------------------------------------------------------
# .docx toparlama
# --------------------------------------------------------------------------
def _text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def _is_blank(p):
    """Metni, resmi, denklemi olmayan paragraf (birleştirmede atlanır)."""
    if _text(p).strip():
        return False
    return not any(el.tag.endswith(("}drawing", "}pict", "}oMath", "}oMathPara", "}sectPr")) for el in p.iter())


def _runs(p):
    return [el for el in p if el.tag != W + "pPr"]


def _append(dst, src, space=True):
    """src paragrafının içeriğini dst'nin sonuna ekler (biçim korunur)."""
    if space and not _text(dst).endswith((" ", "-", "/")):
        r = etree.SubElement(dst, W + "r")
        t = etree.SubElement(r, W + "t")
        t.text = " "
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for el in _runs(src):
        dst.append(el)


def _split(p, offsets):
    """Paragrafı verilen karakter konumlarından böler; biçimli koşular (run) korunur.
    Sondan başa bölünür: her adımda yalnız ilk paragraf kısalır."""
    parts = []
    for cut in sorted(offsets, reverse=True):
        pos, at = 0, None
        for run in _runs(p):
            n = sum(len(t.text or "") for t in run.iter(W + "t"))
            if pos <= cut < pos + n or (cut == pos + n and run.getnext() is not None and n):
                at = (run, cut - pos)
                break
            pos += n
        if at is None:
            continue
        run, k = at
        texts = list(run.iter(W + "t"))
        if 0 < k < sum(len(t.text or "") for t in texts) and len(texts) == 1:
            tail = copy.deepcopy(run)
            tt = next(tail.iter(W + "t"))
            tt.text = texts[0].text[k:]
            tt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            texts[0].text = texts[0].text[:k]
            run.addnext(tail)
            first = tail
        else:
            first = run.getnext() if k else run
        if first is None:
            continue
        new = _new_para_like(p)
        sib = first
        while sib is not None:
            nxt = sib.getnext()
            new.append(sib)
            sib = nxt
        p.addnext(new)
        parts.insert(0, new)
    return [p] + parts


def _new_para_like(p):
    new = etree.Element(W + "p")
    ppr = p.find(W + "pPr")
    if ppr is not None:
        new.append(copy.deepcopy(ppr))
    return new


def _strip_lead(p):
    """Paragrafın başındaki boşlukları siler (bölmeden sonra kalan)."""
    for t in p.iter(W + "t"):
        if not t.text:
            continue
        t.text = t.text.lstrip()
        if t.text:
            break


def should_join(prev, nxt):
    """Gövdede: önceki paragraf cümle ortasında bitiyor ve sonraki küçük harfle sürüyorsa aynı paragraftır."""
    a, b = prev.rstrip(), nxt.lstrip()
    if not a or not b:
        return False
    if b[0].islower():
        return True
    return a.endswith(("-", ",", "(", "&")) or re.search(r"\b(and|or|of|the|in|to|with|et al\.)$", a) is not None


def caption_continues(prev, el):
    """Tablo başlığının ikinci satırı: önceki paragraf başlık, bu kısa ve hemen ardından tablo geliyor."""
    ptxt, txt = _text(prev).strip(), _text(el).strip()
    if not CAPTION_RE.match(ptxt) or ptxt.endswith((".", "!", "?")) or len(txt) > 250:
        return False
    nxt = el.getnext()
    while nxt is not None and nxt.tag == W + "p" and _is_blank(nxt):
        nxt = nxt.getnext()
    return nxt is not None and nxt.tag == W + "tbl"


YEAR_SOON = re.compile(r".{0,160}?\((?:\d{4}[a-z]?|n\.d\.)\)", re.S)


def split_refs_text(text):
    """Bir kaynakça paragrafında yeni girdinin başladığı karakter konumları. Önceki girdi DOI/URL ya da
    sayfa numarasıyla bitiyorsa girdi başı yeter; düz noktayla bitiyorsa yeni girdide yıl da olmalı."""
    cuts = []
    for mt in re.finditer(r"\s+(?=\S)", text):
        start = mt.end()
        rest = text[start:]
        if not REF_START.match(rest):
            continue
        if REF_END.search(text[:start]) or (text[:start].rstrip().endswith(".") and YEAR_SOON.match(rest)):
            cuts.append(start)
    return cuts


XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
LINE_START = re.compile(r"^\s*(?:[•\u2022\uf0b7▪\-–]\s|\d+(?:\.\d+)*\.?\s+[A-Z]|\(?[a-z]\)\s)")


def _lines(p):
    """Paragraftaki satır sonları (w:br): [(satır metni, run, br), …] ve son satır."""
    out, cur = [], []
    for run in _runs(p):
        for el in run.iter(W + "t", W + "br", W + "tab"):
            if el.tag == W + "t":
                cur.append(el.text or "")
            elif el.tag == W + "tab":
                cur.append(" ")
            elif el.get(W + "type") in (None, "textWrapping"):
                out.append(("".join(cur), run, el))
                cur = []
    return out, "".join(cur)


def hard_break(line, nxt, longest):
    """Satır sonu gerçek paragraf sonu mu? (PDF'te ikisi de yalnızca yeni satırdır.)"""
    a, b = line.strip(), nxt.strip()
    if not a or not b:
        return True
    if (REF_HEAD.match(clean_heading(a)) or ABSTRACT_HEAD.match(b) or KEYWORDS_HEAD.match(b)
            or CAPTION_RE.match(b) or LINE_START.match(b) or REF_HEAD.match(clean_heading(b))):
        return True
    if b[0].islower():
        return False
    # kısa satır: paragrafın son satırı (ya da başlık/yazar satırı)
    return len(a) < 0.8 * longest or (a.endswith(SENT_END) and len(a) < 0.92 * longest)


def _split_at_br(p, run, br):
    """Paragrafı w:br'den böler; br'den sonrası yeni paragrafa taşınır."""
    tail = etree.Element(W + "r")
    rpr = run.find(W + "rPr")
    if rpr is not None:
        tail.append(copy.deepcopy(rpr))
    sib = br.getnext()
    while sib is not None:
        nxt = sib.getnext()
        tail.append(sib)
        sib = nxt
    run.remove(br)
    run.addnext(tail)
    new = _new_para_like(p)
    sib = tail
    while sib is not None:
        nxt = sib.getnext()
        new.append(sib)
        sib = nxt
    p.addnext(new)
    return new


def _pymupdf():
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            return None
    return pymupdf


def pdf_line_starts(path):
    """PDF'teki satırların geometrisinden yeni paragrafın başladığı satırları bulur.

    Ölçütler: önceki satır kısa (paragraf sonu ya da başlık), yazı boyutu/kalınlığı değişiyor,
    satırlar arasında boşluk var, girinti var. Döndürülen satır başları (metin) .docx paragraflarını
    bölmek için kullanılır: pdf2docx ortalanmış başlık/yazar satırlarını bazen tek paragrafa koyuyor."""
    pm = _pymupdf()
    if pm is None:
        return []
    starts = []
    try:
        doc = pm.open(path)
    except Exception:  # noqa
        return []
    for page in doc:
        lines = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for ln in b["lines"]:
                spans = [sp for sp in ln["spans"] if sp["text"].strip()]
                if not spans:
                    continue
                main = max(spans, key=lambda sp: len(sp["text"]))
                text = "".join(sp["text"] for sp in ln["spans"]).strip()
                lines.append({"text": text, "x0": ln["bbox"][0], "x1": ln["bbox"][2], "y0": ln["bbox"][1],
                              "y1": ln["bbox"][3], "size": round(main["size"], 1),
                              "bold": bool(main["flags"] & 16) or "Bold" in main["font"]})
        # aynı satırdaki parçaları (üst simge, ayrı yazı tipi) birleştir
        lines.sort(key=lambda d: (round(d["y1"]), d["x0"]))
        merged = []
        for d in lines:
            if merged and abs(merged[-1]["y1"] - d["y1"]) < 3 and d["x0"] >= merged[-1]["x1"] - 1:
                m = merged[-1]
                m["text"] += ("" if d["x0"] - m["x1"] < 1.5 else " ") + d["text"]
                m["x1"] = d["x1"]
                continue
            merged.append(dict(d))
        if not merged:
            continue
        left = min(d["x0"] for d in merged)
        right = max(d["x1"] for d in merged)
        width = right - left
        for prev, d in zip(merged, merged[1:]):
            height = max(prev["y1"] - prev["y0"], 1)
            gap = d["y0"] - prev["y1"]
            mid = (left + right) / 2
            centered = all(abs((x["x0"] + x["x1"]) / 2 - mid) < 4 and x["x0"] > left + 6 for x in (prev, d))
            hard = (d["size"] != prev["size"] or d["bold"] != prev["bold"]
                    or gap > 0.6 * height                       # paragraf aralığı
                    or not centered and (
                        prev["x1"] < right - 0.15 * width       # önceki satır kısa
                        or d["x0"] > left + 0.02 * width and d["x0"] - prev["x0"] > 4))  # ilk satır girintisi
            if hard and len(d["text"]) >= 4:
                starts.append(d["text"])
    doc.close()
    return starts


def _norm_map(text):
    """Boşlukları atılmış metin ve her karakterin özgün konumu."""
    chars, idx = [], []
    for i, c in enumerate(text):
        if not c.isspace():
            chars.append(c)
            idx.append(i)
    return "".join(chars), idx


def start_keys(starts):
    keys = {re.sub(r"\s+", "", s)[:24] for s in starts}
    return {k for k in keys if len(k) >= 4}


def split_at_line_starts(body, keys):
    """Paragraf içinde ortada kalan PDF paragraf başlarından böl."""
    if not keys:
        return
    for p in list(body.findall(W + "p")):
        text = _text(p)
        if REF_HEAD.match(clean_heading(text)):
            break  # kaynakçada asılı girinti satır başı gibi görünür; orası ayrıca bölünür
        norm, idx = _norm_map(text)
        cuts = set()
        for k in keys:
            pos = norm.find(k, 1)
            while pos > 0:
                i = idx[pos]
                # "Table 2:" (kalın etiket) ile başlık metni PDF'te ayrı satır parçası olabilir: bölme
                before = text[:i].rstrip()
                # kısa satır başı ("Abstract", "Albrecht") metnin içinde de geçer: yalnız cümle sonundan sonra
                short_ok = len(k) >= 10 or before.endswith((".", ")", ":", ";", "!", "?"))
                if ((text[i - 1:i].isspace() or text[i - 1:i] in ".,;:") and short_ok
                        and not CAPTION_LABEL.match(text[:i]) and not FUNCTION_WORD_END.search(text[:i])):
                    cuts.add(i)
                pos = norm.find(k, pos + 1)
        if cuts:
            for part in _split(p, sorted(cuts))[1:]:
                _strip_lead(part)


def split_lines(body):
    """pdf2docx aynı girintideki satırları w:br ile tek paragrafa koyabiliyor: gerçek paragraf
    sonlarından böl, satır kaydırmalarını boşluğa çevir."""
    for p in list(body.findall(W + "p")):
        lines, last = _lines(p)
        if not lines:
            continue
        texts = [ln for ln, _, _ in lines] + [last]
        longest = max(len(t.strip()) for t in texts)
        for k in range(len(lines) - 1, -1, -1):
            line, run, br = lines[k]
            if hard_break(line, texts[k + 1], longest):
                _split_at_br(p, run, br)
            else:
                t = etree.Element(W + "t")
                t.text = " "
                t.set(XML_SPACE, "preserve")
                br.addprevious(t)
                run.remove(br)


def _size(run):
    sz = run.find(W + "rPr/" + W + "sz")
    return int(sz.get(W + "val")) if sz is not None and (sz.get(W + "val") or "").isdigit() else None


def _bold(run):
    b = run.find(W + "rPr/" + W + "b")
    return b is not None and b.get(W + "val") not in ("0", "false")


def split_affiliations(body):
    """Özetten önce: "…Guinea. 2School of …" — numaralı kurum satırları ayrı paragraf olmalı."""
    for p in list(body.findall(W + "p")):
        txt = _text(p)
        if ABSTRACT_HEAD.match(txt) or KEYWORDS_HEAD.match(txt):
            break
        cuts = [m.end() for m in re.finditer(r"[.;]\s*(?=[1-9]\d?\s?[A-Z][a-z])", txt)]
        if len(cuts) >= 1 and re.match(r"^\s*[1-9*]", txt):
            for part in _split(p, cuts)[1:]:
                _strip_lead(part)


def mark_headings(body):
    """PDF'te başlık stili yoktur: gövde yazısından büyük, kalın, kısa satırları yazı boyutuna göre
    Heading 1/2/3 yapar (en büyük boyut 1). Özet/anahtar kelimelerden önceki kısım (başlık, yazarlar) atlanır."""
    from collections import Counter
    count = Counter()
    for r in body.iter(W + "r"):
        n = len("".join(t.text or "" for t in r.iter(W + "t")).strip())
        if n and _size(r):
            count[_size(r)] += n
    if not count:
        return
    body_size = count.most_common(1)[0][0]
    paras = [p for p in body if p.tag == W + "p"]
    start = next((k for k, p in enumerate(paras) if KEYWORDS_HEAD.match(_text(p))), None)
    if start is None:
        start = next((k for k, p in enumerate(paras) if ABSTRACT_HEAD.match(_text(p))), -1)
    cands = []
    for p in paras[start + 1:]:
        txt = _text(p).strip()
        runs = [r for r in p.iter(W + "r") if "".join(t.text or "" for t in r.iter(W + "t")).strip()]
        if (not runs or len(txt) > 120 or txt.endswith((".", ",", ";")) or manual_number_level(txt)
                or p.find(W + "pPr/" + W + "pStyle") is not None):
            continue  # elle numaralı başlığın düzeyini numarası belirler
        sizes = {_size(r) for r in runs}
        if len(sizes) == 1 and None not in sizes and all(_bold(r) for r in runs) and min(sizes) >= body_size + 2:
            cands.append((p, sizes.pop()))
    levels = {sz: k + 1 for k, sz in enumerate(sorted({sz for _, sz in cands}, reverse=True)[:3])}
    for p, sz in cands:
        ppr = p.find(W + "pPr")
        if ppr is None:
            ppr = etree.Element(W + "pPr")
            p.insert(0, ppr)
        st = etree.Element(W + "pStyle")
        st.set(W + "val", "Heading%d" % levels[sz])
        ppr.insert(0, st)


def tidy_body(body, starts=(), pairs=None, words=None):
    """document.xml gövdesini yerinde toparlar. starts: PDF'ten bulunan paragraf başları;
    pairs/words: PDF'teki kelime çiftleri (yapışık kelimeleri ayırmak için)."""
    if pairs:
        restore_spaces(body, pairs, words or set())
    # yalnız sayfa numarası olan satırlar ("3", "- 3 -")
    for p in list(body.findall(W + "p")):
        if re.fullmatch(r"[\s\-–]*\d{1,4}[\s\-–]*", _text(p) or "x") and not any(
                x.tag.endswith(("}drawing", "}pict", "}oMath")) for x in p.iter()):
            body.remove(p)
    keys = start_keys(starts)
    split_lines(body)
    split_at_line_starts(body, keys)
    # pdf2docx bağlantıları (e-posta, DOI) w:r'nin içine koyuyor; Word'de w:hyperlink w:r'nin kardeşidir
    for hl in list(body.iter(W + "hyperlink")):
        run = hl.getparent()
        if run.tag == W + "r":
            run.addnext(hl)
    children = list(body)
    in_refs = False
    prev = None  # son metinli üst düzey paragraf
    for el in children:
        if el.tag != W + "p":
            prev = None  # tablo/başka öğe: birleştirme zinciri kopar
            continue
        if _is_blank(el):
            continue
        txt = _text(el)
        if REF_HEAD.match(clean_heading(txt)):
            in_refs, prev = True, None
            continue
        has_media = any(x.tag.endswith(("}drawing", "}pict", "}oMath")) for x in el.iter())
        if prev is not None and not has_media and caption_continues(prev, el):
            _append(prev, el)
            body.remove(el)
            continue
        if prev is not None and not has_media:
            ptxt = _text(prev)
            starts_par = re.sub(r"\s+", "", txt)[:24] in keys or LINE_START.match(txt)
            join = (not REF_START.match(txt)) if in_refs else (not starts_par and should_join(ptxt, txt))
            if join and not (in_refs and not ptxt.strip()):
                _append(prev, el)
                body.remove(el)
                continue
        prev = None if has_media else el
    if in_refs:
        # yapışık girdileri ayır ("…sjmms_357_19 Cayres, L. C. …")
        seen = False
        for el in list(body):
            if el.tag != W + "p":
                continue
            if REF_HEAD.match(clean_heading(_text(el))):
                seen = True
                continue
            if seen:
                cuts = split_refs_text(_text(el))
                if cuts:
                    for part in _split(el, cuts)[1:]:
                        _strip_lead(part)
    split_affiliations(body)
    mark_headings(body)
    # Symbol yazı tipindeki madde işareti (U+F0B7); LaTeX'in yazar yıldızı "∗" (U+2217)
    for t in body.iter(W + "t"):
        if t.text and ("\uf0b7" in t.text or "\u2217" in t.text):
            t.text = t.text.replace("\uf0b7", "\u2022").replace("\u2217", "*")


def tidy_docx(path, pdf=None):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    root = etree.fromstring(items["word/document.xml"])
    pairs, words = pdf_word_pairs(pdf) if pdf else (set(), set())
    body = root.find(W + "body")
    tidy_body(body, pdf_line_starts(pdf) if pdf else (), pairs, words)
    if pdf:
        eqs = pdf_equations(pdf)
        placed = insert_equations(body, eqs, path) if eqs else 0
        if placed:
            print("Bilgi: PDF'teki %d numaralı denklem LaTeX olarak yeniden kuruldu." % placed)
        n = rebuild_tables(body, pdf_tables(pdf))
        if n:
            print("Bilgi: PDF'teki %d tablo yeniden kuruldu." % n)
    items["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)
