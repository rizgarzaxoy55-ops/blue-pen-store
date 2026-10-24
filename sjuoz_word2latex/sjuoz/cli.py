# -*- coding: utf-8 -*-
"""Komut satırı giriş noktası."""
import os
import re
import sys

from .common import safe_console
from .inputs import INPUT_EXTS, extract_zip, zip_names
from .report import REPORT_ERROR
from .pipeline import build_parser, convert
from .issue import convert_issue


def _key(name):
    """Karşılaştırma anahtarı: küçük harf, yalnız harf/rakam, sondaki uzantı(lar) olmadan
    ("makale.docx.docx", "Kapak+Letterdocx" ve "makale" aynı anahtarı verir)."""
    return re.sub(r"(docx|doc|pdf|tex|zip|odt|rtf)+$", "", re.sub(r"[^0-9a-z]", "", name.lower()))


# Aynı ada uyan dosyalardan hangisi kullanılır: Word, LaTeX, eski Word, zip, en son PDF
_RANK = {".docx": 0, ".tex": 1, ".doc": 2, ".odt": 2, ".rtf": 2, ".zip": 3, ".pdf": 4}


def _best(hits):
    """En iyi biçimdeki tek dosya (varsa); yoksa None."""
    if not hits:
        return None
    rank = lambda h: _RANK.get(os.path.splitext(h)[1].lower(), 9)
    top = min(rank(h) for h in hits)
    best = [h for h in hits if rank(h) == top]
    return best[0] if len(best) == 1 else None


def _nearby_dirs(folder):
    """Dosyaların sık bulunduğu yerler: üst klasör, Masaüstü, İndirilenler, Belgeler (OneDrive dahil)."""
    home = os.path.expanduser("~")
    dirs = [os.path.dirname(os.path.abspath(folder))]
    for base in (home, os.path.join(home, "OneDrive"), os.environ.get("OneDrive", "")):
        if base:
            dirs += [os.path.join(base, d) for d in ("Desktop", "Masaüstü", "Downloads", "Documents", "Belgeler")]
    seen, out = set(), []
    for d in dirs:
        key = os.path.normcase(os.path.abspath(d))
        if key not in seen and key != os.path.normcase(os.path.abspath(folder)) and os.path.isdir(d):
            seen.add(key)
            out.append(d)
    return out


def _walk(folder, max_depth=3, limit=20000):
    """Klasör ve alt klasörlerindeki girdi dosyaları (gizli/sistem klasörleri atlanır)."""
    count = 0
    root_depth = os.path.abspath(folder).count(os.sep)
    for d, subdirs, files in os.walk(folder):
        if os.path.abspath(d).count(os.sep) - root_depth >= max_depth:
            subdirs[:] = []
        subdirs[:] = [s for s in subdirs if not s.startswith((".", "$", "_cikti", "node_modules", "AppData"))]
        for f in files:
            count += 1
            if f.lower().endswith(INPUT_EXTS) and not f.startswith("~$"):
                yield os.path.join(d, f)
        if count > limit:
            return


def _matches(folder, want, depth=3):
    return [p for p in _walk(folder, depth) if _key(os.path.basename(p)) == want]


def _from_zips(folders, want, workdir):
    """Aranan dosya bir zip'in içindeyse çıkarılır."""
    hits = []
    for folder in folders:
        try:
            zips = [os.path.join(folder, n) for n in os.listdir(folder) if n.lower().endswith(".zip")]
        except OSError:
            continue
        for zp in zips:
            for name in zip_names(zp):
                if _key(os.path.basename(name)) == want and name.lower().endswith(INPUT_EXTS):
                    hits.append((zp, name))
    if len(hits) != 1:
        return None, hits
    zp, name = hits[0]
    files = extract_zip(zp, os.path.join(workdir, os.path.splitext(os.path.basename(zp))[0]))
    target = next((f for f in files if os.path.basename(f) == os.path.basename(name)), None)
    return target, hits


def resolve_docx(path):
    """Verilen yolu bulmaya çalışır; (yol, None) ya da (None, hata mesajı) döner.

    Windows uzantıları gizlediği için sık görülen yazım hataları düzeltilir:
    "makale.docx.docx", "Kapak+Letterdocx" (nokta unutulmuş), "makale" (uzantısız).
    Yalnız ad verildiyse bu klasör ve alt klasörleri, Masaüstü/İndirilenler/Belgeler ve oralardaki
    zip'lerin içi de aranır.
    """
    if os.path.isfile(path):
        return path, None
    folder = os.path.dirname(path) or "."
    if not os.path.isdir(folder):
        return None, "klasör bulunamadı: %s" % folder
    want = _key(os.path.basename(path))
    here = [p for p in _walk(folder, 1)]
    hits = [p for p in here if _key(os.path.basename(p)) == want]
    if _best(hits):
        return (os.path.relpath(_best(hits)) if folder == "." else _best(hits)), None
    if not hits and not os.path.dirname(path):  # yalnız ad verildiyse daha geniş ara
        hits = _matches(folder, want) + [h for d in _nearby_dirs(folder) for h in _matches(d, want)]
        hits = list(dict.fromkeys(os.path.abspath(h) for h in hits))
        if _best(hits):
            return _best(hits), None
        if not hits:
            import tempfile
            target, zhits = _from_zips([folder] + _nearby_dirs(folder), want, tempfile.mkdtemp(prefix="sjuoz_zip_"))
            if target:
                print("Not: '%s' %s içinden çıkarıldı." % (os.path.basename(path), os.path.basename(zhits[0][0])))
                return target, None
            hits = ["%s → %s" % z for z in zhits]
    msg = ["dosya bulunamadı: %s" % path]
    if len(hits) > 1:
        msg.append("  Bu adla birden çok dosya var; hangisi olduğunu tam yoluyla yazın:")
        msg += ['    "%s"' % h for h in hits]
    if here:
        msg.append("  Bu klasördeki dosyalar (%s):" % os.path.abspath(folder))
        msg += ['    "%s"' % os.path.basename(n) for n in here[:30]]
        msg.append("  Ad boşluk içeriyorsa tırnak içinde yazın; dosya adını yazarken Tab tuşu tamamlar.")
    else:
        msg.append("  %s klasöründe (alt klasörler, Masaüstü, İndirilenler, Belgeler ve oradaki zip'ler dahil)"
                   " bu adla bir .docx / .doc / .pdf / .tex / .zip bulunamadı." % os.path.abspath(folder))
        msg.append("  Dosyaları bu klasöre kopyalayın, tam yolunu verin ya da programı dosya adı"
                   " vermeden çalıştırıp pencereden seçin.")
    return None, "\n".join(msg)


def main(argv=None):
    safe_console()
    ap = build_parser()
    opts = ap.parse_args(argv)
    if opts.sayi:
        if not os.path.isdir(opts.sayi):
            ap.error("sayı klasörü bulunamadı: %s" % opts.sayi)
        res = convert_issue(opts.sayi, opts)
        return 2 if (opts.strict and res["status"] == REPORT_ERROR) else 0
    if not opts.docx:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            opts.docx = list(filedialog.askopenfilenames(
                title="Makale ve kapak mektubu seçin (.docx / .doc / .pdf / .tex / .zip)",
                filetypes=[("Word, PDF, LaTeX, zip", "*.docx *.doc *.odt *.rtf *.pdf *.tex *.zip"),
                           ("Word", "*.docx *.doc *.odt *.rtf"), ("PDF", "*.pdf"), ("LaTeX", "*.tex"), ("Zip", "*.zip")]))
        except Exception:  # noqa
            pass
        if not opts.docx:
            ap.error("en az bir dosya verin (.docx / .doc / .pdf / .tex / .zip)")
    found, errors = [], []
    for p in opts.docx:
        real, err = resolve_docx(p)
        if err:
            errors.append(err)
        else:
            if real != p:
                print("Not: '%s' yerine '%s' kullanıldı." % (p, real))
            found.append(real)
    if errors:
        print("HATA: " + "\n\nHATA: ".join(errors), file=sys.stderr)
        return 2
    opts.docx = found
    result = convert(opts)
    if result is None:
        ap.error("makale dosyası bulunamadı")
    return 2 if (opts.strict and result["status"] == REPORT_ERROR) else 0
