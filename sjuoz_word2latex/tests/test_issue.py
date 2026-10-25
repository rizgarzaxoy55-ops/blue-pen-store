# -*- coding: utf-8 -*-
"""Sayı (toplu işlem) testleri: makaleler sırayla işlenir, sayfa numaraları zincirlenir."""
import csv
import json
import os
import shutil

import pytest

from conftest import CASES, conv


def make_issue(tmp_path, names):
    folder = tmp_path / "sayi"
    order = []
    for name in names:
        case = next(c for c in CASES if c["name"] == name)
        aid = [a for i, a in enumerate(case["args"]) if case["args"][i - 1] == "--article-id"][0]
        sub = folder / aid
        sub.mkdir(parents=True)
        for d in case["docx"]:
            shutil.copy(d, sub)
        order.append({"klasor": aid})
    cfg = {"cilt": 14, "sayi": 4, "yil": 2026, "ay": "October", "ilk_sayfa": 571, "makaleler": order}
    (folder / "sayi.json").write_text(json.dumps(cfg), encoding="utf-8")
    return folder


def run_issue(folder, *extra):
    opts = conv.build_parser().parse_args(["--sayi", str(folder), "--no-zip"] + list(extra))
    return conv.convert_issue(str(folder), opts)


def read_toc(res):
    with open(res["toc"], encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def test_issue_without_compile(tmp_path):
    folder = make_issue(tmp_path, ["sentetik_numerik", "sentetik_temel"])
    res = run_issue(folder, "--no-compile")
    toc = read_toc(res)
    assert [r["Makale no"] for r in toc] == ["2003", "2001"]  # sayi.json'daki sıra
    assert all(r["İlk sayfa"] == "571" for r in toc)  # derleme yoksa zincirlenemez
    assert os.path.isfile(res["summary"])
    assert res["full_pdf"] is None


@pytest.mark.skipif(shutil.which("xelatex") is None, reason="xelatex kurulu değil")
def test_issue_chains_pages_and_merges_pdf(tmp_path):
    folder = make_issue(tmp_path, ["1993", "sentetik_numerik", "sentetik_temel"])
    res = run_issue(folder)
    toc = read_toc(res)
    pages = [(int(r["İlk sayfa"]), int(r["Son sayfa"])) for r in toc]
    expected = [c["expect_pages"] for n in ("1993", "sentetik_numerik", "sentetik_temel")
                for c in CASES if c["name"] == n]
    first = 571
    for (a, b), n in zip(pages, expected):
        assert (a, b) == (first, first + n - 1)
        first = b + 1
    # her makalenin .tex'i kendi ilk sayfasıyla başlamalı
    for r in res["articles"]:
        with open(r["res"]["tex_path"], encoding="utf-8") as f:
            assert r"\setcounter{page}{%d}" % r["first"] in f.read()
    if res["full_pdf"]:
        assert os.path.getsize(res["full_pdf"]) > 0
    assert res["status"] != conv.REPORT_ERROR
