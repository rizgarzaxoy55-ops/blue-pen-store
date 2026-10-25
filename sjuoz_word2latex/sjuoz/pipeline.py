# -*- coding: utf-8 -*-
"""Dönüştürme akışı ve komut satırı seçenekleri."""
import argparse
import datetime
import os
import re
import shutil
import zipfile


from .common import (BIDI, DATA_DIR, NO_WINDOW, warn, WARNINGS)
from .inputs import expand_inputs
from .pdfinput import ensure_docx, prefer_latex_source
from .texinput import is_tex_input, tex_to_docx
from .docx import Docx, Para
from .render import Linker
from .parse import parse_article, parse_cover_letter
from .template import (FIGDIR, LANDSCAPE_COLS, MONTHS, build_tex)
from .writer import Writer
from .wordout import make_docx, qr_from_pdf, running_head
from .export import (article_metadata, bibtex_references, bibtex_self, crossref_xml, jats_xml, validate_crossref)
from .report import (REPORT_OK, analyze_log, build_report, pdf_to_text, report_status, word_source_stats, write_report_html)
from . import render

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


def build_parser():
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
    ap.add_argument("--first-page", type=int, default=None,
                    help="ilk sayfa (verilmezse dergideki gibi 'pp. xxx-xxx' yazılır)")
    ap.add_argument("--last-page", type=int, default=None,
                    help="son sayfa (verilmezse LaTeX otomatik hesaplar)")
    ap.add_argument("--received", default="xxx")
    ap.add_argument("--accepted", default="xxx")
    ap.add_argument("--published", default="xxx")
    ap.add_argument("--article-type", default=None, help="ör. 'Original Article', 'Review Article'")
    ap.add_argument("--main-font", default="Times New Roman")
    ap.add_argument("--header-font", default="Garamond")
    ap.add_argument("--keep-heading-case", action="store_true", help="bölüm başlıklarını BÜYÜK harfe çevirme")
    ap.add_argument("--keep-title-case", action="store_true", help="makale başlığını BÜYÜK harfe çevirme")
    ap.add_argument("--strip-ref-numbers", action="store_true", help="kaynakların başındaki [1] numaralarını sil")
    ap.add_argument("--assets", default=None, help="SJUOZ.png ve orcid.png'nin bulunduğu klasör")
    ap.add_argument("--tex-name", default=None, help="çıktı .tex dosya adı (varsayılan: <makale no>.tex)")
    ap.add_argument("--wide-figs", default="", help="iki sütunu kaplayacak şekiller: '3,5' veya 'all'")
    ap.add_argument("--wide-eqs", default="", help="iki sütunu kaplayacak (çok uzun, bölünemeyen) denklemler: '5,7'")
    ap.add_argument("--landscape-tables", default="",
                    help="yatay sayfaya konacak tablolar, ör. '2,3' (Word'de yatay bölümdekiler ve "
                         "%d+ sütunlu tablolar zaten yatay konur)" % LANDSCAPE_COLS)
    ap.add_argument("--no-links", action="store_true", help="atıf/şekil/tablo bağlantılarını ekleme")
    ap.add_argument("--link-color", default="black",
                    help="metindeki şekil/tablo/denklem/atıf bağlantılarının rengi (varsayılan: black — dergi "
                         "görünümü; tıklanabilirlik görünsün diye ör. blue)")
    ap.add_argument("--no-word", dest="word", action="store_false",
                    help="dergi biçimindeki Word (.docx) sürümünü üretme")
    ap.add_argument("--no-export", dest="export", action="store_false",
                    help="Crossref XML, JATS XML ve BibTeX dosyalarını üretme")
    ap.add_argument("--depositor-name", default=None, help="Crossref gönderen adı (varsayılan: dergi adı)")
    ap.add_argument("--depositor-email", default=None, help="Crossref gönderen e-postası")
    ap.add_argument("--no-compile", dest="compile", action="store_false",
                    help="PDF derleme (varsayılan: xelatex varsa iki kez derlenir)")
    ap.add_argument("--no-zip", dest="zip", action="store_false",
                    help="zip üretme (varsayılan: .tex + figures/ + logolar + PDF tek zip'e konur)")
    ap.add_argument("--no-report", dest="report", action="store_false",
                    help="kontrol raporu üretme (varsayılan: <makale no>_kontrol_raporu.html)")
    ap.add_argument("--strict", action="store_true",
                    help="raporda HATA varsa çıkış kodu 2 olsun (toplu işlemler/CI için)")
    ap.add_argument("--sayi", metavar="KLASOR",
                    help="bir sayının bütün makalelerini işle: KLASOR içindeki her alt klasör bir makale; "
                         "isteğe bağlı KLASOR/sayi.json (cilt, sayı, ay, ilk sayfa, sıra, tarihler)")
    ap.add_argument("--compile", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--zip", action="store_true", help=argparse.SUPPRESS)
    return ap


XELATEX_TIMEOUT = 600  # saniye, her geçiş için


def xelatex_command(exe, tex_name):
    """MiKTeX eksik paketleri sormadan kursun: aksi halde gizli bir pencerede onay bekleyip takılır."""
    import subprocess
    cmd = [exe, "-interaction=nonstopmode"]
    try:
        ver = subprocess.run([exe, "--version"], capture_output=True, text=True, errors="replace",
                             timeout=60, **NO_WINDOW).stdout
    except Exception:  # noqa
        ver = ""
    if "MiKTeX" in ver:
        cmd.append("--enable-installer")
    return cmd + [tex_name]


def run_xelatex(exe, tex_name, outdir, passes=2):
    """xelatex'i iki kez çalıştırır; süre aşılırsa False döner."""
    import subprocess
    cmd = xelatex_command(exe, tex_name)
    print("PDF derleniyor (xelatex; ilk seferde MiKTeX eksik paketleri indirirse birkaç dakika sürebilir)...",
          flush=True)
    for _ in range(passes):
        try:
            subprocess.run(cmd, cwd=outdir, capture_output=True, text=True, errors="replace",
                           timeout=XELATEX_TIMEOUT, **NO_WINDOW)
        except subprocess.TimeoutExpired:
            return False
    return True


def convert(opts):
    """Dönüşümü yapar. Sonuçları (dosya yolları, sayımlar, rapor) sözlük olarak döndürür."""
    del WARNINGS[:]
    BIDI.update(used=False, par=False)
    render.LINKER = None
    today = datetime.date.today()
    if isinstance(getattr(opts, "landscape_tables", ""), str):
        opts.landscape_tables = {x.strip() for x in opts.landscape_tables.split(",") if x.strip()}
    if isinstance(getattr(opts, "wide_eqs", ""), str):
        opts.wide_eqs = {x.strip() for x in opts.wide_eqs.split(",") if x.strip()}
    if isinstance(opts.wide_figs, str):
        opts.wide_figs = {"all"} if opts.wide_figs.strip() == "all" else \
            {x.strip() for x in opts.wide_figs.split(",") if x.strip()}

    # PDF ya da LaTeX kaynağı (.tex / .zip) verildiyse önce Word'e çevrilir
    given = list(opts.docx)
    opts.docx = expand_inputs(opts.docx)  # zip açılır, .doc/.odt/.rtf → .docx
    opts.docx = [prefer_latex_source(p) for p in opts.docx]
    opts.docx = [tex_to_docx(p) if is_tex_input(p) else ensure_docx(p) for p in opts.docx]
    manuscript, cover = classify(opts.docx)
    if manuscript is None:
        return None

    if not opts.article_id:
        for p in given + opts.docx:  # önce kullanıcının verdiği adlar (zip / .tex adı da olabilir)
            mt = re.search(r"(\d{3,})", os.path.basename(p))
            if mt:
                opts.article_id = mt.group(1)
                break
        else:
            opts.article_id = "0000"
            warn("Makale numarası dosya adlarında bulunamadı; 0000 kullanıldı (DOI ve dosya adları buna göre). "
                 "--article-id 1993 gibi verin.")
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
    if not opts.no_links:
        render.LINKER = Linker(art.references, *writer.prescan())
        render.LINKER.sub_targets = writer.sub_targets
    writer.run()
    writer.references()
    writer.close_cols()
    tex = build_tex(art, writer, opts)

    tex_path = os.path.join(outdir, opts.tex_name or "%s.tex" % opts.article_id)
    with open(tex_path, "w", encoding="utf-8") as f:
        f.write(tex)

    if render.LINKER is not None:
        for m_ in sorted(render.LINKER.missing):
            warn("Atıf kaynakçada bulunamadı: (%s)" % m_)
        for r in render.LINKER.refs:
            if r["key"] not in render.LINKER.used:
                warn("Metinde atıf yapılmamış kaynak: %s" % r["label"])
        print("Bilgi: %d / %d kaynağa metinden bağlantı verildi." % (len(render.LINKER.used), len(render.LINKER.refs)))

    # logolar: --assets, makalenin klasörü, programın klasörü veya çalışma klasörü
    bases = [opts.assets] if opts.assets else []
    bases += [os.path.dirname(os.path.abspath(manuscript.path)),
              DATA_DIR, os.getcwd()]
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
    stem = os.path.splitext(tex_path)[0]
    pdf_path = stem + ".pdf"
    log_info = None
    compiled = False
    if opts.compile:
        exe = shutil.which("xelatex")
        if not exe:
            print("Bilgi: xelatex bulunamadı; PDF derlenmedi (MiKTeX / TeX Live kurun ya da zip'i Overleaf'e yükleyin).")
        else:
            if os.path.isfile(pdf_path):
                os.remove(pdf_path)
            timed_out = not run_xelatex(exe, tex_name, outdir)
            compiled = os.path.isfile(pdf_path)
            log_info = analyze_log(stem + ".log", tex.splitlines())
            if timed_out:
                warn("xelatex %d dakikada bitmedi ve durduruldu; MiKTeX Console'dan eksik paketleri kurup"
                     " tekrar deneyin ya da zip'i Overleaf'e yükleyin." % (XELATEX_TIMEOUT // 60))
            elif compiled and not log_info["errors"]:
                print("PDF üretildi: %s" % pdf_path)
            else:
                warn("xelatex hata verdi; ayrıntılar: %s" % (stem + ".log"))
    else:
        print("Derlemek için: cd %s && xelatex %s && xelatex %s" % (outdir, tex_name, tex_name))

    # Dışa aktarma: Crossref (DOI kaydı), JATS (indeksleme), BibTeX
    pages = log_info["pages"] if log_info else None
    last_page = opts.last_page or (opts.first_page + pages - 1 if pages and opts.first_page else None)
    meta = article_metadata(art, opts, opts.first_page, last_page)
    exports = {}
    crossref_check = None
    if opts.export:
        xml = crossref_xml([meta], opts.depositor_name, opts.depositor_email,
                           batch_id="sjuoz-%s-%s-%s-%s" % (opts.volume, opts.issue, opts.article_id,
                                                            datetime.datetime.now().strftime("%Y%m%d%H%M%S")))
        exports["crossref"] = stem + "_crossref.xml"
        exports["jats"] = stem + "_jats.xml"
        exports["bib_refs"] = stem + "_kaynaklar.bib"
        exports["bib_self"] = stem + "_atif.bib"
        with open(exports["crossref"], "wb") as f:
            f.write(xml)
        with open(exports["jats"], "wb") as f:
            f.write(jats_xml(meta, art, writer))
        with open(exports["bib_refs"], "w", encoding="utf-8") as f:
            f.write(bibtex_references(art))
        with open(exports["bib_self"], "w", encoding="utf-8") as f:
            f.write(bibtex_self(meta))
        crossref_check = validate_crossref(xml)
        print("Dışa aktarma: %s" % ", ".join(os.path.basename(v) for v in exports.values()))

    word_path = None
    if getattr(opts, "word", True):
        word_path = os.path.join(outdir, "%s.docx" % opts.article_id)
        if os.path.abspath(word_path) == os.path.abspath(manuscript.path):
            word_path = os.path.join(outdir, "%s_sjuoz.docx" % opts.article_id)
        pages = "%s-%s" % (opts.first_page, last_page or "xxx") if opts.first_page else "xxx-xxx"
        info = {"doi": "https://doi.org/10.25271/sjuoz.%s.%s.%s.%s" % (opts.year, opts.volume, opts.issue,
                                                                       opts.article_id),
                "pages": pages, "runhead": running_head(art),
                "logo": os.path.join(outdir, "SJUOZ.png"), "orcid_png": os.path.join(outdir, "orcid.png"),
                "qr": qr_from_pdf(pdf_path, os.path.join(outdir, "qr.png")) if compiled else None}
        try:
            word_path = make_docx(art, writer, opts, outdir, word_path, info)
        except Exception as e:  # noqa  Word sürümü yan üründür: hata dönüşümü durdurmaz
            warn("Word çıktısı üretilemedi: %s" % e)
            word_path = None
        if info["qr"] and os.path.isfile(info["qr"]):
            os.remove(info["qr"])
        if word_path:
            print("Word (dergi biçiminde): %s" % word_path)

    zp = None
    if opts.zip:
        zp = os.path.join(os.path.dirname(os.path.abspath(outdir)), "%s_latex.zip" % opts.article_id)
        n_img = 0
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            for fn in sorted(os.listdir(outdir)):
                if fn == tex_name or fn in ("SJUOZ.png", "orcid.png") or fn == os.path.basename(pdf_path):
                    z.write(os.path.join(outdir, fn), fn)
            figdir = os.path.join(outdir, FIGDIR)
            if os.path.isdir(figdir):
                for fn in sorted(os.listdir(figdir)):
                    z.write(os.path.join(figdir, fn), FIGDIR + "/" + fn)
                    n_img += 1
        print("Zip (Overleaf'e yüklenebilir): %s  [%s + %d resim/grafik + logolar%s]"
              % (zp, tex_name, n_img, " + PDF" if os.path.isfile(pdf_path) else ""))

    # Kontrol raporu
    src = word_source_stats(manuscript, art)
    report_ctx = {"art": art, "writer": writer, "src": src, "opts": opts, "tex": tex,
                  "pdf_text": pdf_to_text(pdf_path) if compiled else None,
                  "log": log_info, "warnings": list(WARNINGS), "crossref_check": crossref_check,
                  "meta": meta, "export": opts.export}
    items = build_report(report_ctx)
    status = report_status(items)
    report_path = None
    if opts.report:
        report_path = stem + "_kontrol_raporu.html"
        write_report_html(report_path, items, {"id": opts.article_id, "title": art.title,
                                               "when": datetime.datetime.now().strftime("%d.%m.%Y %H:%M")})
    print()
    print("KONTROL RAPORU: %s" % status)
    for it in items:
        if it["level"] != REPORT_OK:
            print("  [%s] %s: %s" % (it["level"], it["title"],
                                     re.sub(r"<[^>]+>", " ", it["detail"].replace("<br>", " | "))[:200]))
    if report_path:
        print("  Ayrıntılar: %s" % report_path)

    return {
        "tex_path": tex_path, "pdf_path": pdf_path if compiled else None, "zip_path": zp, "word_path": word_path,
        "report_path": report_path, "status": status, "items": items, "warnings": list(WARNINGS),
        "stats": {"authors": len(art.authors), "orcids": sum(1 for a in art.authors if a.get("orcid")),
                  "affiliations": len(art.affiliations), "figures": writer.fig_count,
                  "figures_captioned": writer.fig_captioned, "tables": writer.tab_count,
                  "equations": writer.eq_display, "numbered_equations": writer.eq_count,
                  "references": len(art.references),
                  "linked_references": len(render.LINKER.used) if render.LINKER else 0,
                  "pages": log_info["pages"] if log_info else None},
        "source": src,
        "context": report_ctx,
        "meta": meta,
        "exports": exports,
    }
