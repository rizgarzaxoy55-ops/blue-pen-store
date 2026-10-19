# -*- coding: utf-8 -*-
"""Derleme testleri: her test makalesi XeLaTeX ile hatasız PDF'e dönüşmeli.

xelatex kurulu değilse atlanır. Kayıp metin kontrolü bu testte PDF metniyle yapılır.
"""
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
