# -*- coding: utf-8 -*-
"""Derleme testleri: her test makalesi XeLaTeX ile hatasız PDF'e dönüşmeli.

xelatex kurulu değilse atlanır. Kayıp metin kontrolü bu testte PDF metniyle yapılır.
"""
import os
import shutil

import pytest

from conftest import CASES, conv, run_case

needs_xelatex = pytest.mark.skipif(shutil.which("xelatex") is None, reason="xelatex kurulu değil")


@needs_xelatex
@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_compiles_cleanly(case, tmp_path):
    res = run_case(case, tmp_path / "out", compile_pdf=True)
    assert res["pdf_path"], "PDF üretilmedi"
    bad = [it for it in res["items"] if it["level"] == conv.REPORT_ERROR]
    assert not bad, "kontrol raporunda HATA var: %s" % bad
    if case.get("expect_pages"):
        assert res["stats"]["pages"] == case["expect_pages"]


@needs_xelatex
@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_nothing_overflows_the_page(case, tmp_path):
    """Şekil/metin sayfadan veya ilk sayfadaki bilgi kutusunun üstüne taşmamalı."""
    res = run_case(case, tmp_path / "out", compile_pdf=True)
    titles = [it["title"] for it in res["items"]]
    assert "Sayfadan taşan içerik" not in titles


@needs_xelatex
def test_math_equations_stay_inside_columns(tmp_path):
    """İki sütunda hiçbir denklem sütundan taşmaz (LaTeX 'Overfull \\hbox' vermez)."""
    import re
    case = next(c for c in CASES if c["name"] == "sentetik_matematik")
    res = run_case(case, tmp_path / "out", compile_pdf=True)
    with open(os.path.join(str(tmp_path / "out"), "2012.log"), encoding="utf-8", errors="replace") as f:
        log = f.read()
    assert not re.findall(r"Overfull \\hbox \(([\d.]+)pt too wide\) (?:in paragraph|detected)", log)
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    assert tex.count(r"\sjuozfit{") >= 3  # bölünebilen denklemlerin satırlı biçimleri hazır
