#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SJUOZ Word -> LaTeX dönüştürücü — eski tek dosya kullanımı için giriş noktası.

Asıl kod `sjuoz/` paketindedir. Bu dosya sayesinde eskisi gibi çalışır:

    python sjuoz_word2latex.py makale.docx kapak.docx --first-page 571

Paket kurulduysa (`pip install .`) aynı iş şu komutlarla da yapılır:

    sjuoz-word2latex makale.docx kapak.docx --first-page 571
    sjuoz-gui                     # pencere arayüzü
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from sjuoz.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
