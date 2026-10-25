# -*- coding: utf-8 -*-
"""Dışa aktarma testleri: Crossref XML resmi şemaya uymalı, JATS ve BibTeX makaleyle tutarlı olmalı."""
import re

import pytest
from lxml import etree

from conftest import CASES, conv

JATS_DATES = ["--received", "12 May 2026", "--accepted", "1 Aug 2026", "--published", "15 Oct 2026",
              "--depositor-email", "editor@uoz.edu.krd"]


def run_with_dates(case, tmp_path):
    args = list(case["docx"]) + case["args"] + JATS_DATES + ["-o", str(tmp_path / "o"), "--no-zip", "--no-compile"]
    return conv.convert(conv.build_parser().parse_args(args))


@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_crossref_xml_is_schema_valid(case, tmp_path):
    res = run_with_dates(case, tmp_path)
    with open(res["exports"]["crossref"], "rb") as f:
        xml = f.read()
    ok, errors = conv.validate_crossref(xml)
    assert ok, errors
    root = etree.fromstring(xml)
    ns = {"c": conv.CROSSREF_NS}
    assert root.findtext(".//c:doi_data/c:doi", namespaces=ns) == res["meta"]["doi"]
    assert len(root.findall(".//c:person_name", ns)) == res["stats"]["authors"]
    assert len(root.findall(".//c:citation", ns)) == res["stats"]["references"]
    assert len(root.findall(".//c:ORCID", ns)) == res["stats"]["orcids"]


@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_jats_matches_article(case, tmp_path):
    res = run_with_dates(case, tmp_path)
    with open(res["exports"]["jats"], "rb") as f:
        root = etree.fromstring(f.read())
    stats = res["stats"]
    assert root.findtext(".//article-meta/article-id[@pub-id-type='doi']") == res["meta"]["doi"]
    assert len(root.findall(".//contrib")) == stats["authors"]
    assert len(root.findall(".//body//fig")) == stats["figures_captioned"]
    assert len(root.findall(".//table-wrap")) == stats["tables"]
    assert len(root.findall(".//disp-formula")) == stats["equations"]
    assert len(root.findall(".//ref-list/ref")) == stats["references"]
    hrefs = [g.get("{http://www.w3.org/1999/xlink}href") for g in root.iter("graphic")]
    assert all(re.match(r"figures/Figure\d+(?:-[A-Z])?[a-z]?\.\w+$", h) for h in hrefs), hrefs
    hist = root.find(".//history")
    assert hist is not None and hist.find("date[@date-type='accepted']/month").text == "08"


def test_bibtex_files(tmp_path):
    case = next(c for c in CASES if c["name"] == "1993")
    res = run_with_dates(case, tmp_path)
    with open(res["exports"]["bib_refs"], encoding="utf-8") as f:
        bib = f.read()
    entries = re.findall(r"^@(\w+)\{([^,]+),", bib, re.M)
    assert len(entries) == 21 and len({k for _, k in entries}) == 21  # anahtarlar benzersiz
    assert bib.count("{") == bib.count("}")
    assert "@inproceedings{ali2020" in bib and "pages = {132--137}" in bib
    with open(res["exports"]["bib_self"], encoding="utf-8") as f:
        own = f.read()
    assert "doi = {10.25271/sjuoz.2026.14.4.1993}" in own
    assert "Performance Enhancement for Mouth Detection" in own


def test_issue_crossref_single_file(tmp_path):
    import json
    import shutil
    folder = tmp_path / "sayi"
    for name, aid in (("sentetik_temel", "2001"), ("sentetik_numerik", "2003")):
        case = next(c for c in CASES if c["name"] == name)
        (folder / aid).mkdir(parents=True)
        for d in case["docx"]:
            shutil.copy(d, folder / aid)
    (folder / "sayi.json").write_text(json.dumps({"cilt": 14, "sayi": 4, "yil": 2026, "ay": "October",
                                                  "ilk_sayfa": 1}), encoding="utf-8")
    opts = conv.build_parser().parse_args(["--sayi", str(folder), "--no-zip", "--no-compile"])
    res = conv.convert_issue(str(folder), opts)
    with open(res["crossref"], "rb") as f:
        xml = f.read()
    ok, errors = conv.validate_crossref(xml)
    assert ok, errors
    assert len(etree.fromstring(xml).findall(".//{%s}journal_article" % conv.CROSSREF_NS)) == 2


@pytest.mark.parametrize("case", CASES, ids=[c["name"] for c in CASES])
def test_word_version_matches_article(case, tmp_path):
    """Dergi biçimindeki Word sürümü: ön kısım, iki sütun, bütün denklemler Word denklemi, bütün tablolar."""
    from conftest import has_pandoc
    if not has_pandoc():
        pytest.skip("Word çıktısı için pandoc yok (pip install pypandoc_binary)")
    import zipfile
    res = run_with_dates(case, tmp_path)
    assert res["word_path"], "Word dosyası üretilmedi"
    with zipfile.ZipFile(res["word_path"]) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        names = z.namelist()
    stats = res["stats"]
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = etree.fromstring(xml.encode("utf-8"))
    text = "".join(t.text or "" for t in root.iter(W + "t"))
    for au in res["context"]["art"].authors:
        assert au["name"] in text
    assert "ABSTRACT" in text and "Keywords:" in text and "SJUOZTEKSUTUN" not in text and "SJUOZIKISUTUN" not in text
    cols = [c.get(W + "num") for c in root.iter(W + "cols")]
    assert cols[0] == "1" and "2" in cols                      # ön kısım tek, gövde iki sütun
    display = len(re.findall(r"<m:oMathPara", xml))
    assert display >= stats["equations"]
    # ön kısımdaki özet kutusu + gövde tabloları (hücre içi alt tablolar hariç)
    body_tables = [t for t in root.iter(W + "tbl") if not any(a.tag == W + "tbl" for a in t.iterancestors())]
    assert len(body_tables) >= stats["tables"] + 1
    assert sum(1 for n in names if n.startswith("word/media/")) >= min(1, stats["figures"])
    assert any(n.startswith("word/header") for n in names) and any(n.startswith("word/footer") for n in names)


def test_pandoc_latex_simplifies_writer_output():
    from sjuoz.wordout import pandoc_latex
    body = (r"\begin{multicols}{2}" "\n"
            r"\section{INTRODUCTION}" "\nText with \\hyperlink{ref:1}{Smith, 2020}.\n"
            r"\begin{equation}" "\n" r"  \sjuozfit{a=b}{a\\=b}{a}" "\n" r"\end{equation}" "\n"
            r"\end{multicols}" "\n" r"\noindent" "\n" r"\begin{minipage}{\linewidth}" "\n" r"\centering" "\n"
            r"\hypertarget{tab:1}{}" "\n" r"\captionof{table}{Values.}" "\n"
            r"\begin{tabularx}{\linewidth}{@{}>{\raggedright\arraybackslash\hsize=0.50\hsize}X"
            r">{\raggedright\arraybackslash\hsize=1.50\hsize}X@{}}" "\n"
            r"\toprule" "\n" r"\textbf{A} & \textbf{B} \\ \midrule" "\n" r"1 & 2 \\" "\n" r"\bottomrule" "\n"
            r"\end{tabularx}" "\n" r"\end{minipage}" "\n" r"\par\vspace{6pt}" "\n")
    out = pandoc_latex(body)
    assert r"\section{1. INTRODUCTION}" in out
    assert "Smith, 2020" in out and "hyperlink" not in out
    assert r"\[ a=b \qquad (1)\]" in out
    assert r"\textbf{Table 1:} Values." in out and r"\begin{tabular}{p{" in out
    assert "SJUOZTEKSUTUN" in out and "SJUOZIKISUTUN" in out and "multicols" not in out
