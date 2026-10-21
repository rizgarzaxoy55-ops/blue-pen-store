# -*- coding: utf-8 -*-
"""Bir sayının bütün makalelerini tek seferde işleme."""
import datetime
import os
import re
import shutil


from .common import NO_WINDOW, warn
from .template import MONTHS
from .export import crossref_xml, validate_crossref
from .report import (REPORT_CHECK, REPORT_ERROR, REPORT_OK, tex_escape_html)
from .pipeline import build_parser, convert
from .texinput import is_tex_input

# --------------------------------------------------------------------------
# Sayı: bir sayının bütün makalelerini tek seferde işleme
# --------------------------------------------------------------------------
ISSUE_CONFIG = "sayi.json"
ISSUE_KEYS = {  # sayi.json anahtarı -> komut satırı seçeneği
    "gelis": "--received", "kabul": "--accepted", "yayin": "--published", "tur": "--article-type",
}


def load_issue_config(folder):
    path = os.path.join(folder, ISSUE_CONFIG)
    if not os.path.isfile(path):
        return {}
    import json
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def issue_articles(folder, cfg):
    """Sayıdaki makaleler: sayi.json'daki sırayla, yoksa alt klasör adına göre."""
    items = cfg.get("makaleler")
    if not items:
        items = [{"klasor": d} for d in sorted(os.listdir(folder))
                 if os.path.isdir(os.path.join(folder, d)) and not d.startswith((".", "_"))]
    arts = []
    for it in items:
        sub = os.path.join(folder, it["klasor"])
        names = sorted(f for f in os.listdir(sub) if not f.startswith("~$")) if os.path.isdir(sub) else []
        stems = {os.path.splitext(f)[0].lower() for f in names if f.lower().endswith(".docx")}
        # Word dosyası olmayanlar için PDF de kabul edilir (Word'ü varsa o kullanılır)
        docs = [os.path.join(sub, f) for f in names if f.lower().endswith(".docx")
                or (f.lower().endswith(".pdf") and os.path.splitext(f)[0].lower() not in stems)]
        if not any(d.lower().endswith(".docx") for d in docs):
            # .docx yoksa: LaTeX kaynağı (.tex ya da .tex içeren zip), eski Word (.doc/.odt/.rtf),
            # içinde Word/PDF olan zip — hepsi PDF'ten önce gelir (convert() açar/çevirir)
            tex = [os.path.join(sub, f) for f in names if f.lower().endswith(".tex")
                   or (f.lower().endswith(".zip") and is_tex_input(os.path.join(sub, f)))]
            other = [os.path.join(sub, f) for f in names if f.lower().endswith((".doc", ".odt", ".rtf", ".zip"))
                     and os.path.join(sub, f) not in tex]
            if tex or other:
                docs = tex[:1] + other
        if not docs:
            warn("Sayı: '%s' klasöründe .docx / .doc / .tex / .zip / .pdf yok, atlandı." % it["klasor"])
            continue
        aid = it.get("makale_no") or (re.findall(r"\d{3,}", it["klasor"]) or [it["klasor"]])[0]
        arts.append(dict(it, docs=docs, article_id=str(aid)))
    return arts


def merge_pdfs(paths, out):
    try:
        from pypdf import PdfWriter
        w = PdfWriter()
        for p_ in paths:
            w.append(p_)
        with open(out, "wb") as f:
            w.write(f)
        return True
    except ImportError:
        pass
    exe = shutil.which("pdfunite")
    if exe:
        import subprocess
        return subprocess.run([exe] + list(paths) + [out], capture_output=True, **NO_WINDOW).returncode == 0
    return False


def convert_issue(folder, opts):
    """Sayı klasöründeki bütün makaleleri sırayla dönüştürür; sayfa numaralarını zincirler."""
    import csv
    cfg = load_issue_config(folder)
    arts = issue_articles(folder, cfg)
    if not arts:
        print("Sayı klasöründe makale bulunamadı: %s" % folder)
        return {"status": REPORT_ERROR, "articles": []}
    year = int(cfg.get("yil", opts.year))
    month = cfg.get("ay", opts.month) or MONTHS[datetime.date.today().month - 1]
    volume = cfg.get("cilt", opts.volume)
    issue = cfg.get("sayi", opts.issue)
    page = int(cfg.get("ilk_sayfa", opts.first_page))
    outroot = opts.outdir or os.path.join(folder, "_cikti")
    os.makedirs(outroot, exist_ok=True)
    can_chain = bool(shutil.which("xelatex")) and opts.compile
    if not can_chain:
        print("Bilgi: xelatex yok veya --no-compile verildi; sayfa numaraları zincirlenemez, "
              "her makale %d. sayfadan başlar." % page)
    results = []
    for k, art in enumerate(arts):
        print("\n" + "=" * 70)
        print("[%d/%d] Makale %s — ilk sayfa %d" % (k + 1, len(arts), art["article_id"], page))
        print("=" * 70)
        args = list(art["docs"]) + ["-o", os.path.join(outroot, art["article_id"]),
                                    "--article-id", art["article_id"], "--year", str(year), "--month", str(month),
                                    "--first-page", str(page)]
        if volume:
            args += ["--volume", str(volume)]
        if issue:
            args += ["--issue", str(issue)]
        for key, flag in ISSUE_KEYS.items():
            if art.get(key):
                args += [flag, str(art[key])]
        args += list(art.get("secenekler", []))
        if not opts.compile:
            args.append("--no-compile")
        a_opts = build_parser().parse_args(args)
        res = convert(a_opts)
        if res is None:
            results.append({"id": art["article_id"], "status": REPORT_ERROR, "title": "(makale bulunamadı)",
                            "first": page, "last": None, "res": None})
            continue
        pages = res["stats"].get("pages")
        last = page + pages - 1 if pages else None
        results.append({"id": art["article_id"], "status": res["status"], "title": res["context"]["art"].title,
                        "authors": ", ".join(a["name"] for a in res["context"]["art"].authors),
                        "first": page, "last": last, "res": res})
        if can_chain and pages:
            page = last + 1

    # İçindekiler (CSV) ve birleşik PDF
    toc_path = os.path.join(outroot, "icindekiler.csv")
    with open(toc_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["Makale no", "Başlık", "Yazarlar", "İlk sayfa", "Son sayfa", "DOI", "Durum"])
        for r in results:
            w.writerow([r["id"], r["title"], r.get("authors", ""), r["first"], r["last"] or "",
                        "https://doi.org/10.25271/sjuoz.%s.%s.%s.%s" % (year, volume or "", issue or "", r["id"]),
                        r["status"]])
    pdfs = [r["res"]["pdf_path"] for r in results if r["res"] and r["res"]["pdf_path"]]
    full_pdf = None
    if pdfs:
        full_pdf = os.path.join(outroot, "sayi_tam.pdf")
        if not merge_pdfs(pdfs, full_pdf):
            full_pdf = None
            print("Bilgi: PDF'ler birleştirilemedi (pip install pypdf veya pdfunite kurun).")

    # Bütün sayı için tek Crossref dosyası (DOI kaydı tek seferde yapılır)
    crossref_path = None
    metas = [r["res"]["meta"] for r in results if r["res"]]
    if metas and opts.export:
        for r, m_ in zip([r for r in results if r["res"]], metas):
            m_["first_page"], m_["last_page"] = r["first"], r["last"]
        xml = crossref_xml(metas, opts.depositor_name, opts.depositor_email,
                           batch_id="sjuoz-%s-%s-%s" % (volume, issue, datetime.datetime.now().strftime("%Y%m%d%H%M%S")))
        crossref_path = os.path.join(outroot, "sayi_crossref.xml")
        with open(crossref_path, "wb") as f:
            f.write(xml)
        ok, errs = validate_crossref(xml)
        if ok is False:
            warn("Sayı Crossref XML'i şemaya uymuyor: %s" % "; ".join(errs[:3]))
            results.append({"id": "Crossref", "status": REPORT_ERROR, "title": "sayi_crossref.xml şemaya uymuyor",
                            "first": "", "last": "", "res": None})

    status = REPORT_ERROR if any(r["status"] == REPORT_ERROR for r in results) else \
        REPORT_CHECK if any(r["status"] == REPORT_CHECK for r in results) else REPORT_OK
    summary = os.path.join(outroot, "sayi_raporu.html")
    write_issue_html(summary, results, {"volume": volume, "issue": issue, "month": month, "year": year,
                                        "status": status, "full_pdf": full_pdf, "toc": toc_path, "root": outroot,
                                        "crossref": crossref_path})
    print("\n" + "=" * 70)
    print("SAYI ÖZETİ: %d makale, durum %s" % (len(results), status))
    for r in results:
        print("  %-8s s. %s-%s  %-8s %s" % (r["id"], r["first"], r["last"] or "?", r["status"], r["title"][:60]))
    print("  Özet rapor   : %s" % summary)
    print("  İçindekiler  : %s" % toc_path)
    if full_pdf:
        print("  Bütün sayı   : %s" % full_pdf)
    if crossref_path:
        print("  Crossref XML : %s" % crossref_path)
    return {"status": status, "articles": results, "summary": summary, "toc": toc_path, "full_pdf": full_pdf,
            "crossref": crossref_path}


def write_issue_html(path, results, meta):
    color = {REPORT_ERROR: "#b42318", REPORT_CHECK: "#b54708", REPORT_OK: "#067647"}
    bg = {REPORT_ERROR: "#fef3f2", REPORT_CHECK: "#fffaeb", REPORT_OK: "#ecfdf3"}
    root = meta["root"]
    rel = lambda p_: os.path.relpath(p_, root).replace(os.sep, "/") if p_ else None
    rows = []
    for r in results:
        res = r["res"] or {}
        issues = [it for it in res.get("items", []) if it["level"] != REPORT_OK]
        notes = "<br>".join("<b>%s</b>: %s" % (tex_escape_html(it["title"]),
                                                re.sub(r"<[^>]+>", " ", it["detail"])[:160]) for it in issues[:4])
        links = []
        if res.get("report_path"):
            links.append('<a href="%s">rapor</a>' % rel(res["report_path"]))
        if res.get("pdf_path"):
            links.append('<a href="%s">PDF</a>' % rel(res["pdf_path"]))
        if res.get("zip_path"):
            links.append('<a href="%s">zip</a>' % rel(res["zip_path"]))
        rows.append("<tr><td>%s</td><td>%s</td><td>%s–%s</td><td><span class=\"b\" style=\"color:%s;background:%s\">"
                    "%s</span></td><td class=\"d\">%s</td><td>%s</td></tr>" % (
                        tex_escape_html(r["id"]), tex_escape_html(r["title"]), r["first"], r["last"] or "?",
                        color[r["status"]], bg[r["status"]], r["status"], notes, " · ".join(links)))
    extra = []
    if meta.get("full_pdf"):
        extra.append('<a href="%s">Bütün sayı (tek PDF)</a>' % rel(meta["full_pdf"]))
    extra.append('<a href="%s">İçindekiler (CSV)</a>' % rel(meta["toc"]))
    if meta.get("crossref"):
        extra.append('<a href="%s">Crossref XML (bütün sayı)</a>' % rel(meta["crossref"]))
    html = """<!doctype html><html lang="tr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Sayı raporu</title>
<style>body{font:14px/1.5 system-ui,sans-serif;margin:24px auto;max-width:1100px;padding:0 16px;color:#101828;background:#fff}
h1{font-size:20px;margin:0 0 8px}.s{display:inline-block;padding:6px 14px;border-radius:6px;font-weight:700;margin:8px 0 16px}
table{border-collapse:collapse;width:100%%}th,td{border-top:1px solid #eaecf0;padding:8px;vertical-align:top;text-align:left}
th{font-size:12px;color:#475467}.b{font-size:12px;font-weight:700;padding:2px 8px;border-radius:4px;white-space:nowrap}
.d{color:#344054;font-size:13px}a{color:#175cd3}.wrap{overflow-x:auto}</style></head><body>
<h1>SJUOZ Cilt %(vol)s, Sayı %(iss)s — %(month)s %(year)s</h1>
<div>%(extra)s</div><div class="s" style="color:%(c)s;background:%(bg)s">Genel durum: %(status)s</div>
<div class="wrap"><table><tr><th>Makale</th><th>Başlık</th><th>Sayfalar</th><th>Durum</th><th>Bakılacaklar</th><th>Dosyalar</th></tr>
%(rows)s</table></div></body></html>""" % {
        "vol": meta["volume"] or "?", "iss": meta["issue"] or "?", "month": meta["month"], "year": meta["year"],
        "extra": " · ".join(extra), "c": color[meta["status"]], "bg": bg[meta["status"]], "status": meta["status"],
        "rows": "\n".join(rows)}
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
