# -*- coding: utf-8 -*-
"""Küçük parçaların birim testleri: denklem çevirici, yazar ayrıştırma, atıf eşleştirme, rapor."""
from lxml import etree

from conftest import conv

MNS = ('xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
       'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"')


def omml(inner):
    el = etree.fromstring('<m:oMath %s>%s</m:oMath>' % (MNS, inner))
    return conv.OMML().convert(el)


def r(text):
    return '<m:r><m:t xml:space="preserve">%s</m:t></m:r>' % text


def fld(instr, result):
    """Word'ün denklem içine koyduğu SEQ alan kodu."""
    return ('<m:r><w:fldChar w:fldCharType="begin"/></m:r>' + r(instr) +
            '<m:r><w:fldChar w:fldCharType="separate"/></m:r>' + r(result) +
            '<m:r><w:fldChar w:fldCharType="end"/></m:r>')


# ---------------------------------------------------------------- denklemler
def test_seq_field_number_is_removed_and_returned():
    latex, num = omml(r("v") + r("=B-A") + r("        (") + fld(" SEQ Equation \\* ARABIC ", "3") + r(")"))
    assert num == "3"
    assert "SEQ" not in latex and "ARABIC" not in latex
    assert latex == "v=B-A"


def test_plain_trailing_number_needs_wide_gap():
    # "f(2)" bir denklem numarası değildir
    latex, num = omml(r("y=f(2)"))
    assert num is None and latex == "y=f(2)"
    latex, num = omml(r("y=x") + r("      (7)"))
    assert num == "7" and latex == "y=x"


def test_bracket_object_number():
    latex, num = omml(r("L=x") + '<m:d><m:e>%s</m:e></m:d>' % r("1"))
    assert num == "1" and latex == "L=x"


def test_overline_accent_and_word_spaces():
    latex, _ = omml('<m:acc><m:accPr><m:chr m:val="\u0305"/></m:accPr><m:e>%s</m:e></m:acc>' % r("i"))
    assert latex == r"\bar{i}"
    latex, _ = omml('<m:sSub><m:e>%s</m:e><m:sub>%s%s%s</m:sub></m:sSub>' % (r("P"), r("white"), r(" "), r("region")))
    assert latex == r"P_{\mathit{white}\ \mathit{region}}"


def test_fraction_sum_and_greek():
    latex, _ = omml('<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (r("α"), r("β")))
    assert latex == r"\frac{\alpha}{\beta}"
    latex, _ = omml('<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>%s</m:sub><m:sup>%s</m:sup>'
                    '<m:e>%s</m:e></m:nary>' % (r("i=1"), r("N"), r("x")))
    assert latex == r"\sum\limits_{i=1}^{N} x"


def test_equation_number_outside_math():
    assert conv.strip_eq_number_text("       )1(") == (True, "1")
    assert conv.strip_eq_number_text(" (12) ") == (True, "12")
    assert conv.strip_eq_number_text("where x is") == (False, None)


# ---------------------------------------------------------------- metin
def test_tex_escape():
    assert conv.tex_escape("50% & $x_1#{}") == r"50\% \& \$x\_1\#\{\}"


def test_render_plain_links_and_etal():
    s = conv.render_plain("see (Smith et al., 2020) at https://a.org/x_y.")
    assert r"\textit{et al.}" in s
    assert r"\url{https://a.org/x_y}." in s


def test_title_case_caps_keeps_acronyms():
    assert conv.title_case_caps("IEEE TRANSACTIONS ON SYSTEMS, MAN, AND CYBERNETICS") == \
        "IEEE Transactions on Systems, Man, and Cybernetics"


def test_clean_heading_and_manual_number():
    assert conv.clean_heading("2.1. Data Set ") == "Data Set"
    assert conv.manual_number_level("2.1 Data") == 2


# ---------------------------------------------------------------- yazarlar
def test_initials_and_abbreviation():
    assert conv.initials("Mohammed E. Safi") == "M. E. Safi"
    s = conv.abbreviate_in_text("design: Mohammed E. Safi 1, Eyad I., Alaa H. Ali", "Eyad I. Abbas")
    assert "E. I. Abbas" in s
    assert conv.abbreviate_in_text("by Mohammed E. Safi 1,", "Mohammed E. Safi") == "by M. E. Safi,"


def test_strip_titles_and_match():
    assert conv.strip_titles("Asst. Prof. Dr. Mohammed E. Safi") == "Mohammed E. Safi"
    cl = [{"name": "Eyad I. Abbas"}, {"name": "Alaa H. Ali"}]
    assert conv.match_author("Alaa H. Ali", cl)["name"] == "Alaa H. Ali"
    assert conv.match_author("John Smith", cl) is None


def test_merge_affiliations_renumbers():
    art = conv.Article()
    art.affiliations = ["Dept A, Univ X.", "Dept A, Univ X.", "Dept B, Univ Y."]
    art.authors = [{"name": "A", "marks": ["1", "*"]}, {"name": "B", "marks": ["2"]},
                   {"name": "C", "marks": ["3"]}]
    conv.merge_affiliations(art)
    assert art.affiliations == ["Dept A, Univ X.", "Dept B, Univ Y."]
    assert [a["marks"] for a in art.authors] == [["1", "*"], ["1"], ["2"]]


# ---------------------------------------------------------------- atıflar
class FakePara:
    def __init__(self, text):
        self._t = text

    def text(self):
        return self._t


def test_linker_author_year_variants():
    refs = [FakePara("Ali, W., Saeed, T. (2020). Paper one."),
            FakePara("Safi, M. E., & Abbas, E. I. (2016). Paper two."),
            FakePara("Safi, M. E., & Abbas, E. I. (2022). Paper three."),
            FakePara("MathWorks. (n.d.). Docs.")]
    lk = conv.Linker(refs, 0, 0, 0)
    spans = lk.spans("as in (Wisam H Ali et al., 2020; Mohammed E. Safi & Abbas, 2016) and (MathWorks, n.d.)")
    assert [t for _, _, t in spans] == ["ref:1", "ref:2", "ref:4"]
    assert not lk.missing
    assert "ref:3" not in lk.used


def test_linker_numeric_and_xrefs():
    refs = [FakePara("[1] A. B, x, 2020."), FakePara("[2] C. D, y, 2021."), FakePara("[3] E. F, z, 2022.")]
    lk = conv.Linker(refs, 2, 1, 0)
    assert lk.numeric
    targets = [t for _, _, t in lk.spans("see [1], [2, 3] and Figure 2, Table 1, Figure 5.")]
    assert targets == ["ref:1", "ref:2", "ref:3", "fig:2", "tab:1"]  # Figure 5 yok: bağlanmaz


# ---------------------------------------------------------------- rapor
def test_lost_text_detection_and_hyphenation():
    words = conv._norm_words("The Viola-Jones algo-\nrithm detects mouths in real time with high speed.")
    assert "violajones" in words and "algorithm" in words
    src = conv._shingles(conv._norm_words("a b c d e f g h i j"))
    out = conv._shingles(conv._norm_words("x a b c d e f g h i j y"))
    assert src <= out


def test_analyze_log(tmp_path):
    log = tmp_path / "a.log"
    log.write_text("! Undefined control sequence.\nl.12 \\foo\n\n"
                   "Overfull \\hbox (15.2pt too wide) in paragraph at lines 2--3\n"
                   "Overfull \\hbox (0.5pt too wide) in paragraph at lines 1--1\n"
                   "Missing character: There is no ş in font x!\n"
                   "Output written on a.pdf (3 pages, 100 bytes).\n", encoding="utf-8")
    res = conv.analyze_log(str(log), ["l1", "long line here", "l3"])
    assert res["errors"] and res["pages"] == 3
    assert [o["line"] for o in res["overfull_h"]] == [2]  # 2pt altı sayılmaz
    assert res["missing_chars"] == ["ş"]
