# -*- coding: utf-8 -*-
"""Kontrol raporu: Word ile çıktının karşılaştırılması, LaTeX log analizi."""
import hashlib
import os
import re
import shutil
import unicodedata


from .common import (C_NS, M, MC, NO_WINDOW, NS, RTL_RUN_RE, W)
from .docx import Para, Table
from .render import strip_eq_number_text
from .parse import CAPTION_RE, SUBFIG_RE
from .figures import layout_table_images
from .export import JOURNAL
from . import render

# --------------------------------------------------------------------------
# Kontrol raporu
# --------------------------------------------------------------------------
REPORT_ERROR, REPORT_CHECK, REPORT_OK = "HATA", "KONTROL", "TAMAM"


def _norm_words(text):
    text = unicodedata.normalize("NFKC", text).lower()
    text = re.sub(r"[-\u00ad]\s*", "", text)  # tire ve satır sonu bölünmeleri: "algo-\nrithm" = "algorithm"
    return re.findall(r"[0-9a-z\u00c0-\u024f]+", text)


def _shingles(words, n=5):
    return {" ".join(words[i:i + n]) for i in range(max(0, len(words) - n + 1))}


def _letters(text):
    """Sadece harf ve rakamlar (boşluk/biçim farklarından bağımsız karşılaştırma için).
    URL'ler ve 'DOI:' etiketleri çıkarılır; dönüştürücü bunları bilerek düzenler."""
    text = re.sub(r"(?:https?://|www\.)\S+|\bdoi\s*:|\b10\.\d{4,9}/\S+", " ", text, flags=re.I)
    return "".join(_norm_words(text))


def found_ratio(src_text, out_letters, size=24, step=12):
    """Kaynak paragrafın 24 harflik parçalarının çıktıda bulunma oranı
    (sütun/sayfa sonunda bölünen paragraflar için kısa parçalar)."""
    s = _letters(src_text)
    if len(s) <= size:
        return 1.0 if s in out_letters else 0.0
    chunks = [s[i:i + size] for i in range(0, len(s) - size + 1, step)]
    return sum(1 for c in chunks if c in out_letters) / len(chunks)


def text_segments(p):
    """Paragraf metnini satır içi denklemlerden böler: "Where [y] is …" Word metninde "Where is …"
    olur ama PDF'te arada denklem vardır; sınırı aşan parçalar yanlış "kayıp" alarmı verir."""
    segs, cur = [], []
    for t in p.tokens:
        if t[0] == "t":
            cur.append(t[1])
        elif t[0] == "m":
            segs.append("".join(cur))
            cur = []
    segs.append("".join(cur))
    return segs


def found_ratio_segments(segs, out_letters):
    """Parçaların bulunma oranının harf sayısına göre ağırlıklı ortalaması."""
    total = found = 0
    for seg in segs:
        n = len(_letters(seg))
        if n < 8:  # denklem arasındaki "and", "where" gibi kısa parçalar
            continue
        total += n
        found += n * found_ratio(seg, out_letters)
    return found / total if total else 1.0


def tex_to_text(tex):
    """PDF metni yoksa .tex'ten kaba düz metin çıkarır (kayıp metin kontrolü için)."""
    body = tex[tex.find(r"\begin{document}"):]
    body = re.sub(r"(?<!\\)%.*", "", body)
    body = re.sub(r"\\(?:href|hyperlink)\{[^{}]*\}", "", body)
    body = re.sub(r"\\url\{([^{}]*)\}", r" \1 ", body)
    body = re.sub(r"\\(?:includegraphics(?:\[[^\]]*\])?|hypertarget|qrcode(?:\[[^\]]*\])?)\{[^{}]*\}", "", body)
    body = re.sub(r"\\[a-zA-Z@]+\*?(?:\[[^\]]*\])?", " ", body)
    body = body.replace(r"\&", "&").replace(r"\%", "%").replace("~", " ")
    return re.sub(r"[{}$\\^_]", " ", body)


def pdf_to_text(pdf_path):
    exe = shutil.which("pdftotext")
    if exe and os.path.isfile(pdf_path):
        import subprocess
        try:
            r = subprocess.run([exe, "-enc", "UTF-8", pdf_path, "-"], capture_output=True, **NO_WINDOW, timeout=120)
            return r.stdout.decode("utf-8", "replace")
        except Exception:  # noqa
            return None
    try:  # isteğe bağlı: pip install pypdf
        from pypdf import PdfReader
        return "\n".join(pg.extract_text() or "" for pg in PdfReader(pdf_path).pages)
    except Exception:  # noqa
        return None


WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
ICON_EMU = 300000  # ~0,8 cm: yazar satırındaki ORCID/e-posta simgeleri şekil sayılmaz


def is_icon(blip):
    for anc in blip.iterancestors(WP + "inline", WP + "anchor"):
        ext = anc.find(WP + "extent")
        if ext is not None:
            return int(ext.get("cx", "0")) < ICON_EMU and int(ext.get("cy", "0")) < ICON_EMU
        break
    return False


def word_source_stats(doc, art):
    """Word dosyasının kendisinden (dönüştürücüden bağımsız) sayımlar."""
    body = doc.doc.find(W + "body")
    text = "".join(t.text or "" for t in body.iter(W + "t"))
    hashes = set()
    for att in body.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS):
        # mc:Fallback içindeki yedek resimleri sayma
        if is_icon(att.getparent()):
            continue
        _, data = doc.media(str(att))
        if data:
            hashes.add(hashlib.md5(data).hexdigest())
    charts = len(body.xpath(".//c:chart", namespaces={"c": C_NS}))
    fallback = set()
    for fb in body.iter(MC + "Fallback"):
        for rid in fb.xpath(".//a:blip/@r:embed | .//v:imagedata/@r:id", namespaces=NS):
            _, data = doc.media(rid)
            if data:
                fallback.add(hashlib.md5(data).hexdigest())
    in_choice = set()
    for ch in body.iter(MC + "Choice"):
        for rid in ch.xpath(".//a:blip/@r:embed", namespaces=NS):
            _, data = doc.media(rid)
            if data:
                in_choice.add(hashlib.md5(data).hexdigest())
    images = len(hashes - (fallback - in_choice)) + charts
    tables = sum(1 for t in body.findall(W + "tbl") if not layout_table_images(Table(t, doc)))
    max_num = lambda pat: max([int(x) for x in re.findall(pat, text)] or [0])
    return {
        "images": images,
        "charts": charts,
        "tables": tables,
        "max_figure_ref": max_num(r"\bFig(?:ure|\.)\s*(\d+)"),
        "max_table_ref": max_num(r"\bTable\s*(\d+)"),
        "display_equations": sum(1 for p in body.iter(W + "p")
                                 if p.find(M + "oMathPara") is not None or (
                                     p.find(M + "oMath") is not None and
                                     strip_eq_number_text("".join(t.text or "" for t in p.findall(".//" + W + "t")))[0])),
        "references": len(art.references),
    }


def analyze_log(log_path, tex_lines):
    res = {"errors": [], "overfull_h": [], "overfull_v": 0, "missing_chars": [], "missing_files": [],
           "undefined_links": [], "pages": None}
    if not os.path.isfile(log_path):
        return res
    log = open(log_path, encoding="utf-8", errors="replace").read()
    lines = log.splitlines()
    for i, ln in enumerate(lines):
        if ln.startswith("! "):
            ctx = " ".join(x.strip() for x in lines[i + 1:i + 3] if x.strip())
            res["errors"].append((ln[2:] + " " + ctx).strip()[:300])
    for mt in re.finditer(r"Overfull \\hbox \(([\d.]+)pt too wide\) (?:in paragraph|detected|in alignment) at lines? (\d+)", log):
        amount, line = float(mt.group(1)), int(mt.group(2))
        if amount > 2.0:
            snippet = tex_lines[line - 1].strip() if 0 < line <= len(tex_lines) else ""
            res["overfull_h"].append({"pt": amount, "line": line, "tex": snippet[:160]})
    res["overfull_v"] = len(re.findall(r"Overfull \\vbox \(([\d.]+)pt too high\)", log))
    res["missing_chars"] = sorted(set(re.findall(r"Missing character: There is no (.+?) in font", log)))
    # sütuna sığmayıp küçültülen denklemler: (numara, oran)
    res["eq_scaled"] = [(num, float(a) / float(b)) for num, a, b in
                        re.findall(r"SJUOZ-EQ-SCALED: ([^;\n]*); ([\d.]+)pt; ([\d.]+)pt", log) if float(b)]
    res["missing_files"] = sorted(set(re.findall(r"File `([^']+)' not found", log)))
    res["undefined_links"] = sorted(set(re.findall(r"name\{([^}]+)\} has been referenced but does not exist", log)))
    mt = re.search(r"Output written on .*?\((\d+) pages?", log)
    res["pages"] = int(mt.group(1)) if mt else None
    return res


def build_report(ctx):
    """Dönüşümü denetler; maddeleri (seviye, başlık, ayrıntı) olarak döndürür."""
    art, writer, src = ctx["art"], ctx["writer"], ctx["src"]
    items = []

    def add(level, title, detail=""):
        items.append({"level": level, "title": title, "detail": detail})

    # 1) Sayımlar: Word ↔ LaTeX
    n_img_out = len([v for v in writer.saved_hashes.values() if "/Table" not in v])
    n_img_tab = len(writer.saved_hashes) - n_img_out
    if n_img_out + n_img_tab < src["images"]:
        add(REPORT_ERROR, "Eksik resim",
            "Word'de %d resim/grafik var, LaTeX'e %d tanesi aktarıldı." % (src["images"], n_img_out + n_img_tab))
    else:
        add(REPORT_OK, "Resimler", "%d resim/grafik aktarıldı%s." % (
            n_img_out + n_img_tab, " (%d Word grafiği)" % src["charts"] if src["charts"] else ""))
    if src["max_figure_ref"] > writer.fig_count:
        add(REPORT_ERROR, "Şekil sayısı tutmuyor",
            "Metinde 'Figure %d'e kadar atıf var, PDF'te %d şekil var." % (src["max_figure_ref"], writer.fig_count))
    n_figs = len(writer.figure_log)  # "Figure 1-A/1-B" iki ayrı şekil kutusudur
    if n_figs != writer.fig_captioned:
        add(REPORT_CHECK, "Başlıksız şekil",
            "%d şeklin %d tanesinde başlık (caption) bulunamadı." % (n_figs, n_figs - writer.fig_captioned))
    if writer.tab_count != src["tables"]:
        add(REPORT_ERROR, "Tablo sayısı tutmuyor",
            "Word'de %d tablo var, LaTeX'te %d." % (src["tables"], writer.tab_count))
    elif src["max_table_ref"] > writer.tab_count:
        add(REPORT_CHECK, "Tablo atfı",
            "Metinde 'Table %d'ye atıf var ama %d tablo var." % (src["max_table_ref"], writer.tab_count))
    else:
        add(REPORT_OK, "Tablolar", "%d tablo aktarıldı." % writer.tab_count)
    if writer.eq_display != src["display_equations"]:
        add(REPORT_CHECK, "Denklem sayısı",
            "Word'de %d ayrı satır denklem var, LaTeX'te %d." % (src["display_equations"], writer.eq_display))
    else:
        add(REPORT_OK, "Denklemler", "%d denklem (%d numaralı)." % (writer.eq_display, writer.eq_count))
    if not art.references:
        add(REPORT_ERROR, "Kaynakça yok", "'References' başlığı bulunamadı.")
    else:
        add(REPORT_OK, "Kaynakça", "%d kaynak." % len(art.references))

    # 2) Ön bilgiler
    if not art.abstract:
        add(REPORT_ERROR, "Özet bulunamadı", "ABSTRACT kutusu boş kalacak.")
    if not art.keywords:
        add(REPORT_CHECK, "Anahtar kelime yok", "'Keywords' satırı bulunamadı.")
    if not art.authors:
        add(REPORT_ERROR, "Yazar bulunamadı", "")
    else:
        no_orcid = [a["name"] for a in art.authors if not a.get("orcid")]
        if no_orcid:
            add(REPORT_CHECK, "ORCID eksik", ", ".join(no_orcid))
        else:
            add(REPORT_OK, "Yazarlar", "%d yazar, hepsinin ORCID'i var." % len(art.authors))
    if not art.corr_email:
        add(REPORT_CHECK, "Sorumlu yazar e-postası yok", "")

    # 3) Atıflar
    if render.LINKER is not None and render.LINKER.refs:
        if render.LINKER.missing:
            add(REPORT_CHECK, "Kaynakçada bulunamayan atıflar", "; ".join(sorted(render.LINKER.missing)))
        unused = [r["label"] for r in render.LINKER.refs if r["key"] not in render.LINKER.used]
        if unused:
            add(REPORT_CHECK, "Metinde atıf yapılmamış kaynaklar", "; ".join(unused))
        if not render.LINKER.missing and not unused:
            add(REPORT_OK, "Atıflar", "%d kaynağın hepsine metinden bağlantı verildi." % len(render.LINKER.refs))

    # 4) Kayıp metin: Word'deki her paragraf PDF'te (yoksa .tex'te) var mı?
    out_text = ctx.get("pdf_text")
    where = "PDF"
    if not out_text:
        out_text, where = tex_to_text(ctx["tex"]), ".tex"
    out_letters = _letters(out_text)
    # pdftotext sağdan sola/soldan sağa karışık satırları bozuk çıkarabilir: bu paragraflar .tex'le de denenir
    tex_letters = _letters(tex_to_text(ctx["tex"])) if where == "PDF" else out_letters
    lost = []
    blocks = list(art.body) + list(art.references)
    n_tab = 0
    for b in blocks:
        if isinstance(b, Table) and not layout_table_images(b):
            # tablo hücreleri kısa olduğu için hücre hücre bakılır (sayfadan taşan satırlar)
            n_tab += 1
            cells = [" ".join(p.text() for p in c[0] if isinstance(p, Para)).strip() for r in b.rows for c in r]
            missing = [c for c in cells if len(_letters(c)) >= 3 and _letters(c) not in out_letters]
            if missing:
                lost.append("Tablo %d: %d / %d hücre bulunamadı (ör. %s)" % (
                    n_tab, len(missing), len(cells), ", ".join("'%s'" % m_[:30] for m_ in missing[:5])))
        paras = [b] if isinstance(b, Para) else [p for r in b.rows for c in r for p in c[0]]
        for p in paras:
            if not isinstance(p, Para):
                continue
            words = _norm_words(p.text())
            if len(words) < 8:
                continue
            # yazar adları bilerek kısaltılmış olabilir (M. E. Safi): o biçimi de dene
            src_text = p.text()
            cap = CAPTION_RE.match(src_text)
            sub = SUBFIG_RE.match(src_text)
            if cap:  # "Figure 8" / "Figure 2-A:" numarasını LaTeX kendisi yazar
                src_text = src_text[sub.end():] if sub else cap.group(3)
            found = max(found_ratio(v, out_letters) for v in (src_text, writer.abbreviate_names(src_text)))
            if found < 0.7 and any(t[0] == "m" for t in p.tokens):
                found = max(found, found_ratio_segments(text_segments(p), out_letters))
            if found < 0.7 and RTL_RUN_RE.search(src_text):
                found = max(found, found_ratio(src_text, tex_letters))
            if found < 0.7:
                lost.append("%s… (%%%d bulundu)" % (" ".join(p.text().split()[:12]), round(found * 100)))
    if lost:
        add(REPORT_ERROR, "Kaybolmuş olabilecek metin (%s'te bulunamadı)" % where, "<br>".join(lost[:15]))
    else:
        add(REPORT_OK, "Metin bütünlüğü", "Word'deki bütün paragraflar %s'te bulundu." % where)

    # 5) Derleme
    lg = ctx.get("log")
    if lg is None:
        add(REPORT_CHECK, "PDF derlenmedi", "xelatex bulunamadı veya --no-compile verildi.")
    else:
        if lg["errors"]:
            add(REPORT_ERROR, "LaTeX hataları", "<br>".join(lg["errors"][:10]))
        if lg["missing_files"]:
            add(REPORT_ERROR, "Bulunamayan dosyalar", ", ".join(lg["missing_files"]))
        if lg["missing_chars"]:
            add(REPORT_ERROR, "Yazı tipinde olmayan karakterler (PDF'te görünmez)", " ".join(lg["missing_chars"]))
        if lg["undefined_links"]:
            add(REPORT_CHECK, "Hedefi olmayan bağlantılar", ", ".join(lg["undefined_links"]))
        if lg["overfull_h"]:
            add(REPORT_CHECK, "Sütundan taşan satırlar",
                "<br>".join("%.1fpt — .tex satır %d: <code>%s</code>" % (o["pt"], o["line"], tex_escape_html(o["tex"]))
                            for o in sorted(lg["overfull_h"], key=lambda o: -o["pt"])[:10]))
        small = [(n, r) for n, r in lg.get("eq_scaled", []) if r < 0.8]
        if small:
            add(REPORT_CHECK, "Çok küçültülen denklemler",
                "Sütuna sığması için küçültüldü (bölünecek yer yok): " +
                ", ".join("denklem %s yakınında %%%d" % (n, round(r * 100)) for n, r in small) +
                ". Okunurluğu kontrol edin; iki sütunu kaplaması için --wide-eqs %s verin "
                "ya da Word'de denklemi satırlara bölün." % ",".join(n for n, _ in small))
        if lg["overfull_v"]:
            add(REPORT_CHECK, "Sayfadan taşan içerik", "%d yerde sayfa/sütun yüksekliği aşıldı." % lg["overfull_v"])
        if not (lg["errors"] or lg["missing_files"] or lg["missing_chars"]):
            add(REPORT_OK, "Derleme", "PDF hatasız derlendi (%s sayfa)." % (lg["pages"] or "?"))

    # 6) Crossref (DOI kaydı)
    if ctx.get("export"):
        check = ctx.get("crossref_check")
        meta = ctx.get("meta") or {}
        if check and check[0] is False:
            add(REPORT_ERROR, "Crossref XML şemaya uymuyor", "<br>".join(tex_escape_html(e) for e in check[1]))
        elif check and check[0]:
            add(REPORT_OK, "Crossref XML", "Crossref 5.3.1 şemasına uygun (DOI: %s)." % meta.get("doi", ""))
        missing = []
        if not (ctx["opts"].depositor_email or JOURNAL["depositor_email"]):
            missing.append("gönderen e-postası (--depositor-email)")
        if not meta.get("last_page"):
            missing.append("son sayfa (PDF derlenmedi)")
        if not meta.get("published"):
            missing.append("yayın tarihi (--published)")
        if missing:
            add(REPORT_CHECK, "Crossref için eksik bilgi", ", ".join(missing))

    # 7) Dönüştürücü uyarıları
    for w_ in ctx["warnings"]:
        if not any(w_ in it["detail"] or it["detail"] in w_ for it in items if it["detail"]):
            add(REPORT_CHECK, "Uyarı", w_)
    return items


def tex_escape_html(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def report_status(items):
    levels = {it["level"] for it in items}
    return REPORT_ERROR if REPORT_ERROR in levels else REPORT_CHECK if REPORT_CHECK in levels else REPORT_OK


def write_report_html(path, items, meta):
    color = {REPORT_ERROR: "#b42318", REPORT_CHECK: "#b54708", REPORT_OK: "#067647"}
    bg = {REPORT_ERROR: "#fef3f2", REPORT_CHECK: "#fffaeb", REPORT_OK: "#ecfdf3"}
    status = report_status(items)
    order = {REPORT_ERROR: 0, REPORT_CHECK: 1, REPORT_OK: 2}
    rows = []
    for it in sorted(items, key=lambda x: order[x["level"]]):
        detail = it["detail"] if "<br>" in it["detail"] or "<code>" in it["detail"] else tex_escape_html(it["detail"])
        rows.append('<tr><td><span class="b" style="color:%s;background:%s">%s</span></td><td><b>%s</b>'
                    '<div class="d">%s</div></td></tr>' % (color[it["level"]], bg[it["level"]], it["level"],
                                                         tex_escape_html(it["title"]), detail))
    html = """<!doctype html><html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Kontrol raporu %(id)s</title>
<style>body{font:14px/1.5 system-ui,sans-serif;margin:24px auto;max-width:900px;padding:0 16px;color:#101828;background:#fff}
h1{font-size:20px;margin:0 0 4px}.m{color:#475467;margin-bottom:16px}
.s{display:inline-block;padding:6px 14px;border-radius:6px;font-weight:700;color:%(c)s;background:%(bg)s;margin-bottom:16px}
table{border-collapse:collapse;width:100%%}td{border-top:1px solid #eaecf0;padding:8px;vertical-align:top}
.b{font-size:12px;font-weight:700;padding:2px 8px;border-radius:4px;white-space:nowrap}.d{color:#344054;margin-top:2px;word-break:break-word}
code{background:#f2f4f7;padding:1px 4px;border-radius:3px;font-size:12px}</style></head><body>
<h1>Kontrol raporu — %(id)s</h1><div class="m">%(title)s<br>%(when)s</div>
<div class="s">Genel durum: %(status)s</div><table>%(rows)s</table></body></html>""" % {
        "id": tex_escape_html(meta["id"]), "title": tex_escape_html(meta["title"]), "when": meta["when"],
        "status": status, "c": color[status], "bg": bg[status], "rows": "\n".join(rows)}
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
