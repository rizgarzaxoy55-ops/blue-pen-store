# -*- coding: utf-8 -*-
"""Gerileme (regression) testleri.

Her test makalesi dönüştürülür ve
  1) üretilen .tex, daha önce kontrol edilip saklanmış .tex ile (tests/golden) karşılaştırılır,
  2) sayımlar (yazar, ORCID, şekil, tablo, denklem, kaynak...) case.json'daki beklenenle karşılaştırılır,
  3) kontrol raporunun durumu case.json'daki en kötü izin verilen durumdan kötü olmamalıdır.

Kodda bilerek yapılan bir değişiklik çıktıyı değiştiriyorsa farkı inceleyin ve
`pytest --update-golden` ile beklenen çıktıları güncelleyin.
"""
import difflib
import os
import re

import pytest

from conftest import CASES, GOLDEN, conv, run_case

RANK = {conv.REPORT_OK: 0, conv.REPORT_CHECK: 1, conv.REPORT_ERROR: 2}


def normalize(tex):
    # EMF/WMF dönüştürücüsü olmayan makinelerde uzantı farklı olabilir
    return re.sub(r"(figures/[\w-]+)\.(png|jpe?g|pdf|emf|wmf|svg|eps)\b", r"\1.IMG", tex)


@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_golden_tex(case, tmp_path, update_golden):
    res = run_case(case, tmp_path / "out")
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = normalize(f.read())
    golden_path = os.path.join(GOLDEN, case["name"] + ".tex")
    if update_golden or not os.path.isfile(golden_path):
        os.makedirs(GOLDEN, exist_ok=True)
        with open(golden_path, "w", encoding="utf-8") as f:
            f.write(tex)
        if not update_golden:
            pytest.skip("beklenen çıktı ilk kez yazıldı: %s" % golden_path)
        return
    with open(golden_path, encoding="utf-8") as f:
        golden = f.read()
    if tex != golden:
        diff = "".join(difflib.unified_diff(golden.splitlines(True), tex.splitlines(True),
                                            "beklenen/" + case["name"], "yeni/" + case["name"], n=2))
        pytest.fail("LaTeX çıktısı beklenenden farklı (bilerek değiştirildiyse "
                    "`pytest --update-golden`):\n" + diff[:6000])


@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_counts_and_report(case, tmp_path):
    res = run_case(case, tmp_path / "out")
    stats = res["stats"]
    wrong = {k: (v, stats.get(k)) for k, v in case["expect"].items() if stats.get(k) != v}
    assert not wrong, "beklenen ≠ bulunan: %s" % wrong
    max_status = case.get("max_status", conv.REPORT_CHECK)
    errors = [it for it in res["items"] if it["level"] == conv.REPORT_ERROR]
    assert RANK[res["status"]] <= RANK[max_status], "rapor durumu %s: %s" % (res["status"], errors)
    assert os.path.isfile(res["report_path"])


def _titles(items, level):
    return [it["title"] for it in items if it["level"] == level]


def test_report_catches_dropped_paragraph_and_figure(tmp_path):
    """Rapor gerçekten kaybı yakalıyor mu? Çıktıdan bir paragraf ve bir resim bilerek silinir."""
    case = next(c for c in CASES if c["name"] == "1993")
    res = run_case(case, tmp_path / "out")
    ctx = dict(res["context"])
    sentence = "For the second criterion (detection ratio), a practical experiment"
    assert sentence in ctx["tex"]
    start = ctx["tex"].index(sentence)
    ctx["tex"] = ctx["tex"][:start] + ctx["tex"][ctx["tex"].index("\n", start):]
    writer = ctx["writer"]
    removed = next(iter(writer.saved_hashes))
    saved = dict(writer.saved_hashes)
    del writer.saved_hashes[removed]
    try:
        items = conv.build_report(ctx)
    finally:
        writer.saved_hashes.update(saved)
    errors = _titles(items, conv.REPORT_ERROR)
    assert any(t.startswith("Kaybolmuş") for t in errors), errors
    assert "Eksik resim" in errors, errors


def test_report_clean_for_untouched_output(tmp_path):
    case = next(c for c in CASES if c["name"] == "1993")
    res = run_case(case, tmp_path / "out")
    assert not _titles(res["items"], conv.REPORT_ERROR)


def test_report_catches_lost_table_rows(tmp_path):
    """Sayfadan taşıp kaybolan tablo satırları (kısa hücreler) HATA olarak raporlanmalı."""
    case = next(c for c in CASES if c["name"] == "sentetik_uzun_tablo")
    res = run_case(case, tmp_path / "out")
    ctx = dict(res["context"])
    lines = ctx["tex"].splitlines(True)
    ctx["tex"] = "".join(l for l in lines if not l.startswith(("S77 &", "S78 &", "S79 &")))
    errors = [it for it in conv.build_report(ctx) if it["level"] == conv.REPORT_ERROR]
    assert errors and "Tablo 1" in errors[0]["detail"], errors


def test_long_table_is_breakable(tmp_path):
    case = next(c for c in CASES if c["name"] == "sentetik_uzun_tablo")
    res = run_case(case, tmp_path / "out")
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    assert r"\begin{xltabular}" in tex and r"\endhead" in tex
    assert all("S%02d &" % i in tex for i in range(1, 81))


def test_landscape_sections_and_wide_tables(tmp_path):
    case = next(c for c in CASES if c["name"] == "sentetik_yatay")
    res = run_case(case, tmp_path / "out")
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    body = tex[tex.index(r"\begin{document}"):]
    assert body.count(r"\begin{landscape}") == 2 == body.count(r"\end{landscape}")
    first = body[body.index(r"\begin{landscape}"):body.index(r"\end{landscape}")]
    assert "tab:1" in first and "fig:1" in first  # Word'deki yatay bölüm
    second = body[body.rindex(r"\begin{landscape}"):]
    assert "tab:2" in second  # 12 sütunlu tablo


def test_landscape_option_for_any_table(tmp_path):
    case = next(c for c in CASES if c["name"] == "sentetik_temel")
    args = list(case["docx"]) + case["args"] + ["-o", str(tmp_path / "o"), "--no-zip", "--no-compile",
                                                "--landscape-tables", "1"]
    res = conv.convert(conv.build_parser().parse_args(args))
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    assert tex.count(r"\begin{landscape}") == 1


def test_right_to_left_text(tmp_path):
    case = next(c for c in CASES if c["name"] == "sentetik_rtl")
    res = run_case(case, tmp_path / "out")
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    assert r"\TeXXeTstate=1" in tex
    assert r"\rtl{زانکۆ}" in tex and r"\rtl{جامعة}" in tex  # cümle içindeki kelimeler
    assert tex.count(r"\begin{rtlpar}") == 2  # Kürtçe (Word'de RTL) + Arapça (oranla bulunan)
    assert r"\ltr{2024}." in tex  # sayı soldan sağa, nokta sağdan sola akışta


def test_rtl_preamble_only_when_needed(tmp_path):
    case = next(c for c in CASES if c["name"] == "sentetik_temel")
    res = run_case(case, tmp_path / "out")
    with open(res["tex_path"], encoding="utf-8") as f:
        assert "TeXXeTstate" not in f.read()
