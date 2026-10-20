# -*- coding: utf-8 -*-
"""Pencere arayüzünün ekran gerektirmeyen mantığı: formdaki değerlerden seçenek üretimi."""
import pytest

from conftest import conv
from sjuoz import gui


def test_build_args_single_article():
    args = gui.build_args({"manuscript": "m.docx", "cover": "c.docx", "article_id": "1993", "volume": "14",
                           "issue": "4", "year": "2026", "month": "October", "first_page": "571",
                           "received": "12 May 2026", "accepted": "", "published": "", "article_type": "",
                           "outdir": "out", "depositor_email": "e@x.org", "compile": False, "zip": True,
                           "export": True, "links": False})
    opts = conv.build_parser().parse_args(args)
    assert opts.docx == ["m.docx", "c.docx"] and opts.article_id == "1993" and opts.volume == 14
    assert opts.first_page == 571 and opts.received == "12 May 2026" and opts.outdir == "out"
    assert opts.compile is False and opts.zip is True and opts.no_links is True
    assert opts.depositor_email == "e@x.org"


def test_build_args_issue_and_validation():
    args = gui.build_args({"issue_folder": "sayi", "volume": "14", "issue": "4", "year": "2026",
                           "month": "October", "first_page": "571"}, mode="sayi")
    opts = conv.build_parser().parse_args(args)
    assert opts.sayi == "sayi" and opts.issue == 4
    with pytest.raises(ValueError):
        gui.build_args({"manuscript": ""})
    with pytest.raises(ValueError):
        gui.build_args({"manuscript": "m.docx", "volume": "on dört"})


def test_guess_article_id():
    assert gui.guess_article_id("/x/1993_Finalversion.docx") == "1993"
    assert gui.guess_article_id("makale.docx") == ""


def test_cli_survives_non_utf8_console(tmp_path):
    """Windows konsolu (cp1252) Türkçe/Arapça karakterleri gösteremese de program çökmemeli."""
    import os
    import subprocess
    import sys
    from conftest import FIXTURES, ROOT
    env = dict(os.environ, PYTHONIOENCODING="cp1252")
    docx = os.path.join(FIXTURES, "sentetik_rtl", "sentetik_rtl.docx")
    r = subprocess.run([sys.executable, os.path.join(ROOT, "sjuoz_word2latex.py"), docx, "-o", str(tmp_path / "o"),
                        "--no-compile", "--no-zip"], env=env, capture_output=True, timeout=120)
    assert r.returncode == 0, r.stderr.decode("cp1252", "replace")[-500:]
    assert (tmp_path / "o" / "0000.tex").exists()
