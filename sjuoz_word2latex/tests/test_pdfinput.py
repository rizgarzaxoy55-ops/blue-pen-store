# -*- coding: utf-8 -*-
"""PDF girişi: PDF'ten çevrilen .docx'in toparlanması ve uçtan uca dönüşüm."""
import os

import pytest
from lxml import etree

from conftest import FIXTURES, conv
from sjuoz import pdfinput
from sjuoz.common import W

WNS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def body(*paras):
    """Her öğe bir paragraf: str ya da [(metin, kalın mı), …]; '\n' satır sonudur (w:br)."""
    xml = []
    for p in paras:
        runs = [(p, False)] if isinstance(p, str) else p
        rx = []
        for text, bold in runs:
            parts = text.split("\n")
            inner = '<w:br/>'.join('<w:t xml:space="preserve">%s</w:t>' % t for t in parts)
            rx.append("<w:r>%s%s</w:r>" % ("<w:rPr><w:b/></w:rPr>" if bold else "", inner))
        xml.append("<w:p>%s</w:p>" % "".join(rx))
    return etree.fromstring("<w:body %s>%s</w:body>" % (WNS, "".join(xml)))


def texts(b):
    return [pdfinput._text(p) for p in b.findall(W + "p")]


def test_reference_entry_starts():
    for s in ("Safi, M. E., & Abbas, E. I. (2016).", "Van Hai, V., Nhung, H. (2022).", "[3] C. Author,",
              "Al-Ghafari, A. B. (2020).", "MathWorks. (n.d.). Docs"):
        assert pdfinput.REF_START.match(s), s
    for s in ("International Journal of Immunogenetics, 42(3)", "Endocrine, 55(2), 599", "(2019). Study of",
              "Egyptian patients. Journal of"):
        assert not pdfinput.REF_START.match(s), s


def test_glued_references_are_split():
    t = ("Ali, W. (2020). One. J, 8(3), 188-195. https://doi.org/10.4103/x_1 Cayres, L. C., & Lee, B. (2021). "
         "Two. Endocrine, 55(2), 599-606. Guleryuz, B., Akin, F. (2016). Three.")
    cuts = pdfinput.split_refs_text(t)
    assert [t[c:c + 6] for c in cuts] == ["Cayres", "Gulery"]  # yazar listesindeki "Lee, B." bölünmez


def test_body_paragraph_rules():
    assert pdfinput.should_join("The consequences of inaccurate estimation are", "projects went over budget.")
    assert not pdfinput.should_join("This is a sentence.", "Another paragraph starts here.")
    assert pdfinput.FUNCTION_WORD_END.search("In ")
    assert pdfinput.CAPTION_LABEL.match("Table 2: ") and not pdfinput.CAPTION_LABEL.match("Table 2: Values")


def test_split_keeps_formatting_and_fixes_links():
    b = body([("Smith, J. (2019). A paper. ", False), ("Journal A", True),
              (", 1, 2-3. https://doi.org/10.1/a Brown, A. (2020). B.", False)])
    p = b.find(W + "p")
    pdfinput._split(p, [pdfinput._text(p).index("Brown")])
    assert texts(b) == ["Smith, J. (2019). A paper. Journal A, 1, 2-3. https://doi.org/10.1/a ", "Brown, A. (2020). B."]
    assert b.find(".//" + W + "b") is not None  # kalın dergi adı korunur
    # pdf2docx w:hyperlink'i w:r içine koyar; çıkarılmalı
    x = etree.fromstring('<w:body %s><w:p><w:r><w:t>E-mail: </w:t><w:hyperlink><w:r><w:t>a@b.org</w:t></w:r>'
                         '</w:hyperlink></w:r></w:p></w:body>' % WNS)
    pdfinput.tidy_body(x)
    assert x.find(".//" + W + "hyperlink").getparent().tag == W + "p"


def test_line_breaks_and_page_breaks():
    b = body("Introduction\nThe method works well on the first data set and it was\ntested again on the second data set.",
             "The consequences of inaccurate estimation are documented and can lead to",
             "projects that went over budget.",
             "References",
             "Ali, W. (2020). One. J, 8(3), 188-195.",
             "https://doi.org/10.1/x Brown, A. (2019). Two. K, 3, 10-20.")
    pdfinput.tidy_body(b)
    assert [t.strip() for t in texts(b)] == [
        "Introduction",
        "The method works well on the first data set and it was tested again on the second data set.",
        "The consequences of inaccurate estimation are documented and can lead to projects that went over budget.",
        "References",
        "Ali, W. (2020). One. J, 8(3), 188-195. https://doi.org/10.1/x",
        "Brown, A. (2019). Two. K, 3, 10-20."]


def test_pdf_line_starts_split_title_from_authors():
    b = body("A STUDY OF SOMETHING Ali Hassan1,*, Sara Ahmed2 1Department of Physics, Zakho, Iraq.")
    pdfinput.tidy_body(b, ["Ali Hassan1,*, Sara Ahmed2", "1Department of Physics, Zakho, Iraq."])
    assert texts(b) == ["A STUDY OF SOMETHING ", "Ali Hassan1,*, Sara Ahmed2 ", "1Department of Physics, Zakho, Iraq."]


PDF = os.path.join(FIXTURES, "sentetik_pdf", "sentetik_pdf.pdf")


def test_pdf_end_to_end(tmp_path):
    pytest.importorskip("pdf2docx")
    opts = conv.build_parser().parse_args([PDF, "--article-id", "2011", "--volume", "14", "--issue", "4",
                                          "--year", "2026", "--month", "October", "-o", str(tmp_path / "o"),
                                          "--no-zip", "--no-compile"])
    res = conv.convert(opts)
    st = res["stats"]
    assert (st["authors"], st["affiliations"], st["references"], st["linked_references"]) == (3, 2, 2, 2)
    assert st["figures_captioned"] == 1
    assert any("PDF'ti" in w for w in res["warnings"])
    assert not [i for i in res["items"] if i["level"] == "HATA"], res["items"]
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    assert r"\section{INTRODUCTION}" in tex and "A STUDY OF SOMETHING" in tex


def test_two_column_detection():
    pytest.importorskip("pdf2docx")
    sample = os.path.join(FIXTURES, "..", "..", "ornek_cikti_1993", "1993.pdf")
    assert pdfinput.looks_two_column(sample)
    assert not pdfinput.looks_two_column(PDF)
