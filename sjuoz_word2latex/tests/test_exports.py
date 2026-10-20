# -*- coding: utf-8 -*-
"""Dışa aktarma testleri: Crossref XML resmi şemaya uymalı, JATS ve BibTeX makaleyle tutarlı olmalı."""
import re

import pytest
from lxml import etree

from conftest import CASES, conv, run_case

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
    assert all(re.match(r"figures/Figure\d+[a-z]?\.\w+$", h) for h in hrefs), hrefs
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
