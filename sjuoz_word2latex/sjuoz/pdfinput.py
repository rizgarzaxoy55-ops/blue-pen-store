# -*- coding: utf-8 -*-
"""PDF girişi: PDF'i Word'e (.docx) çevirip dönüştürücünün okuyabileceği biçime toparlar.

PDF'te paragraf, başlık, denklem bilgisi yoktur; yalnızca sayfadaki satırlar vardır. Bu yüzden
sonuç yazarın Word dosyası kadar iyi olmaz: denklemler resim/düz metin olarak gelir, biçim kaybolabilir.
Çeviri pdf2docx ile yapılır (pip install pdf2docx), ardından:
  * satır/sayfa sonunda bölünmüş paragraflar birleştirilir,
  * kaynakça yeniden girdilere ayrılır (PDF'te girdiler birbirine yapışık ya da bölük gelir),
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
    r"|[A-Z][\w&\- ]+\.\s\((?:\d{4}|n\.d\.))")
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
    warn(PDF_NOTE)
    if looks_two_column(path):
        warn("%s iki sütunlu dizilmiş bir PDF gibi görünüyor (ör. yayımlanmış dergi sayfası): sütunlar ve "
             "boşluklar karışabilir, sonuç güvenilir değil. Bu yol tek sütunlu yazar taslakları içindir."
             % os.path.basename(path))
    return out


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


def split_refs_text(text):
    """Bir kaynakça paragrafında yeni girdinin başladığı karakter konumları."""
    cuts = []
    for mt in re.finditer(r"\s+(?=\S)", text):
        start = mt.end()
        if REF_END.search(text[:start]) and REF_START.match(text[start:]):
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


def tidy_body(body, starts=()):
    """document.xml gövdesini yerinde toparlar. starts: PDF'ten bulunan paragraf başları."""
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
    mark_headings(body)
    # Symbol yazı tipindeki madde işareti (U+F0B7) ve benzerleri
    for t in body.iter(W + "t"):
        if t.text and "" in t.text:
            t.text = t.text.replace("", "•")


def tidy_docx(path, pdf=None):
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    root = etree.fromstring(items["word/document.xml"])
    tidy_body(root.find(W + "body"), pdf_line_starts(pdf) if pdf else ())
    items["word/document.xml"] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in items.items():
            z.writestr(n, data)
