# -*- coding: utf-8 -*-
"""Komut satırı giriş noktası."""
import os


from .report import REPORT_ERROR
from .pipeline import build_parser, convert
from .issue import convert_issue


def main(argv=None):
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
                title="Makale ve kapak mektubu (.docx) seçin", filetypes=[("Word", "*.docx")]))
        except Exception:  # noqa
            pass
        if not opts.docx:
            ap.error("en az bir .docx dosyası verin")
    for p in opts.docx:
        if not os.path.isfile(p):
            ap.error("dosya bulunamadı: %s" % p)
    result = convert(opts)
    if result is None:
        ap.error("makale dosyası bulunamadı")
    return 2 if (opts.strict and result["status"] == REPORT_ERROR) else 0
