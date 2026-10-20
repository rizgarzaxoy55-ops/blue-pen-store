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


def test_build_args_blue_links():
    base = {"manuscript": "m.docx", "links": True}
    assert conv.build_parser().parse_args(gui.build_args(base)).link_color == "black"
    opts = conv.build_parser().parse_args(gui.build_args(dict(base, link_blue=True)))
    assert opts.link_color == "blue"


def test_link_color_sanitized():
    from sjuoz.template import link_color
    assert link_color("blue!70!black") == "blue!70!black"
    assert link_color(None) == "black"
    assert link_color("red}\\evil") == "black"


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


def test_resolve_docx_forgives_windows_typos(tmp_path):
    from sjuoz.cli import resolve_docx
    for n in ("paper1995.docx.docx", "1995+Cover+Letter.docx", "eski.doc"):
        (tmp_path / n).write_bytes(b"x")
    assert resolve_docx(str(tmp_path / "paper1995.docx")) == (str(tmp_path / "paper1995.docx.docx"), None)
    assert resolve_docx(str(tmp_path / "1995+Cover+Letterdocx"))[0] == str(tmp_path / "1995+Cover+Letter.docx")
    real, err = resolve_docx(str(tmp_path / "eski.docx"))
    assert real is None and "Farklı Kaydet" in err and '"paper1995.docx.docx"' in err


def test_resolve_docx_looks_in_desktop(tmp_path, monkeypatch):
    from sjuoz.cli import resolve_docx
    home, work = tmp_path / "home", tmp_path / "home" / "OneDrive" / "Desktop" / "kod"
    work.mkdir(parents=True)
    (home / "Downloads").mkdir()
    (home / "Downloads" / "paper1995.docx").write_bytes(b"x")
    monkeypatch.setenv("HOME", str(home))
    monkeypatch.setenv("USERPROFILE", str(home))
    monkeypatch.delenv("OneDrive", raising=False)
    monkeypatch.chdir(work)
    assert resolve_docx("paper1995.docx") == (str(home / "Downloads" / "paper1995.docx"), None)
    real, err = resolve_docx("yok.docx")
    assert real is None and "pencereden seçin" in err
