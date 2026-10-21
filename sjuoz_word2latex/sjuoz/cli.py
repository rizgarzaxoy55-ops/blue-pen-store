# -*- coding: utf-8 -*-
"""Komut satırı giriş noktası."""
import os
import re
import sys

from .common import safe_console
from .report import REPORT_ERROR
from .pipeline import build_parser, convert
from .issue import convert_issue


def _key(name):
    """Karşılaştırma anahtarı: küçük harf, yalnız harf/rakam, sondaki docx'ler olmadan."""
    return re.sub(r"(docx)+$", "", re.sub(r"[^0-9a-z]", "", name.lower()))


def _nearby_dirs(folder):
    """Word dosyalarının sık bulunduğu yerler: üst klasör, Masaüstü, İndirilenler, Belgeler."""
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


def _matches(folder, want):
    try:
        return [os.path.join(folder, n) for n in sorted(os.listdir(folder))
                if n.lower().endswith(".docx") and not n.startswith("~$") and _key(n) == want]
    except OSError:
        return []


def resolve_docx(path):
    """Verilen yolu bulmaya çalışır; (yol, None) ya da (None, hata mesajı) döner.

    Windows uzantıları gizlediği için sık görülen yazım hataları düzeltilir:
    "makale.docx.docx", "Kapak+Letterdocx" (nokta unutulmuş), "makale" (uzantısız).
    """
    if os.path.isfile(path):
        return path, None
    folder = os.path.dirname(path) or "."
    try:
        names = sorted(n for n in os.listdir(folder) if not n.startswith("~$"))
    except OSError:
        return None, "klasör bulunamadı: %s" % folder
    want = _key(os.path.basename(path))
    docx = [n for n in names if n.lower().endswith(".docx")]
    hits = [n for n in docx if _key(n) == want]
    if len(hits) == 1:
        return os.path.join(folder, hits[0]) if folder != "." else hits[0], None
    if not os.path.dirname(path):  # yalnız ad verildiyse yakın klasörlere de bak
        hits = [h for d in _nearby_dirs(folder) for h in _matches(d, want)]
        if len(hits) == 1:
            return hits[0], None
    msg = ["dosya bulunamadı: %s" % path]
    if len(hits) > 1:
        msg.append("  Bu adla birden çok dosya var; hangisi olduğunu tam yoluyla yazın:")
        msg += ['    "%s"' % h for h in hits]
    old = [n for n in names if n.lower().endswith(".doc") and _key(n[:-4]) == want]
    if old:
        msg.append("  '%s' eski .doc biçiminde; Word'de Farklı Kaydet → .docx yapın." % old[0])
    if docx:
        msg.append("  Bu klasördeki Word dosyaları (%s):" % os.path.abspath(folder))
        msg += ['    "%s"' % n for n in docx]
        msg.append("  Ad boşluk içeriyorsa tırnak içinde yazın; dosya adını yazarken Tab tuşu tamamlar.")
    else:
        msg.append("  %s klasöründe hiç .docx dosyası yok (Masaüstü, İndirilenler ve Belgeler'de de"
                   " bulunamadı)." % os.path.abspath(folder))
        msg.append("  Word dosyalarını bu klasöre kopyalayın, tam yolunu verin ya da programı dosya adı"
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
                title="Makale ve kapak mektubu (.docx / .pdf) seçin",
                filetypes=[("Word, PDF, LaTeX", "*.docx *.pdf *.tex *.zip"), ("Word", "*.docx"), ("PDF", "*.pdf"),
                           ("LaTeX (.tex ya da .tex içeren .zip)", "*.tex *.zip")]))
        except Exception:  # noqa
            pass
        if not opts.docx:
            ap.error("en az bir .docx dosyası verin")
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
