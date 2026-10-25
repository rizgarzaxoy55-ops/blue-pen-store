# -*- coding: utf-8 -*-
"""Küçük parçaların birim testleri: denklem çevirici, yazar ayrıştırma, atıf eşleştirme, rapor."""
from lxml import etree

from conftest import conv
from sjuoz import report

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
    words = report._norm_words("The Viola-Jones algo-\nrithm detects mouths in real time with high speed.")
    assert "violajones" in words and "algorithm" in words
    src = report._shingles(report._norm_words("a b c d e f g h i j"))
    out = report._shingles(report._norm_words("x a b c d e f g h i j y"))
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


def test_clean_sub_label():
    assert conv.clean_sub_label("(a) Original image") == "Original image"
    assert conv.clean_sub_label("b) Detected") == "Detected"
    assert conv.clean_sub_label("(c)") == ""


def test_bidi_inline_and_rtl_detection():
    conv.BIDI.update(used=False, par=False)
    s = conv.render_plain("the word زانکۆ means university")
    assert s == r"the word \rtl{زانکۆ} means university" and conv.BIDI["used"]
    assert conv.is_rtl_text("ئەم توێژینەوەیە باس لە 2024")
    assert not conv.is_rtl_text("Mostly English text with one word زانکۆ")


# ---------------------------------------------------------------- dışa aktarma
def test_bibtex_authors():
    assert conv.bibtex_authors("Safi, M. E., & Abbas, E. I.") == "Safi, M. E. and Abbas, E. I."
    assert conv.bibtex_authors("Ali, Wissam Hassan.") == "Ali, Wissam Hassan"
    assert conv.bibtex_authors("A. Author, B. Author, and C. Author") == "A. Author and B. Author and C. Author"
    assert conv.bibtex_authors("Indrabayu, Hasanuddin, & Nurtanio, I.") == "Indrabayu, Hasanuddin and Nurtanio, I."


def test_parse_reference_apa_journal_and_proceedings():
    r = conv.parse_reference("Balasubramanian, M., & Ramalingam, V. (2009). Real time face recognition. "
                             "Expert Systems with Applications, 36(3, Part 2), 6879–6888. "
                             "https://doi.org/https://doi.org/10.1016/j.eswa.2008.08.001")
    assert (r["journal"], r["volume"], r["number"], r["pages"], r["year"]) == \
        ("Expert Systems with Applications", "36", "3, Part 2", "6879--6888", "2009")
    assert r["doi"] == "10.1016/j.eswa.2008.08.001" and r["title"] == "Real time face recognition"
    p = conv.parse_reference("Ali, W. (2020). FPGA system. 2020 International Conference on Computer "
                             "Science (CSASE), 132–137. https://doi.org/10.1109/CSASE48920.2020.9142095")
    assert p["pages"] == "132--137" and "volume" not in p


def test_parse_reference_ieee():
    r = conv.parse_reference('[3] C. Author, "Third paper," Journal C, vol. 3, pp. 21–30, 2022. doi: 10.1000/ccc')
    assert (r["title"], r["journal"], r["volume"], r["pages"], r["year"], r["doi"]) == \
        ("Third paper", "Journal C", "3", "21--30", "2022", "10.1000/ccc")


def test_parse_date_and_title_case():
    assert conv.parse_date("12 May 2026") == (2026, 5, 12)
    assert conv.parse_date("2026-05-12") == (2026, 5, 12)
    assert conv.parse_date("XX") is None
    assert conv.title_case_caps("MOUTH DETECTION BASED ON MODIFIED VIOLA-JONES ALGORITHM") == \
        "Mouth Detection Based on Modified Viola-Jones Algorithm"


def test_placeholder_png_is_valid(tmp_path, monkeypatch):
    from PIL import Image
    from sjuoz import figures
    figures.placeholder_png(str(tmp_path / "a.png"), "Figure3.emf")
    assert Image.open(tmp_path / "a.png").size == (1200, 600)
    Image.open(__import__("io").BytesIO(figures._GRAY_PNG)).load()  # Pillow'suz yedek de geçerli PNG


def test_unmarked_asterisk_is_not_an_author():
    class P:
        tokens = [("t", "Ban J. Khoshaba", (0, 0, None)), ("t", "1", (0, 0, "superscript")),
                  ("t", ", Layla S. Mohammed", (0, 0, None)), ("t", "1", (0, 0, "superscript")), ("t", "*", (0, 0, None))]
    from sjuoz.parse import parse_authors
    authors = parse_authors(P())
    assert [a["name"] for a in authors] == ["Ban J. Khoshaba", "Layla S. Mohammed"]
    assert authors[1]["marks"] == ["1", "*"]


def test_xelatex_command_enables_miktex_installer(tmp_path):
    import os
    import stat
    import sys
    from sjuoz.pipeline import xelatex_command
    fake = tmp_path / "xelatex"
    fake.write_text("#!%s\nprint('MiKTeX-XeTeX 4.10 (MiKTeX 23.10)')\n" % sys.executable)
    fake.chmod(fake.stat().st_mode | stat.S_IEXEC)
    if os.name != "nt":
        assert "--enable-installer" in xelatex_command(str(fake), "a.tex")


def test_glued_heading_numbers():
    assert conv.clean_heading("1.INTRODUCTION") == "INTRODUCTION"
    assert conv.clean_heading("6-REFERENCES") == "REFERENCES"
    assert conv.manual_number_level("2.MATERIAL AND METHOD") == 1
    assert conv.clean_heading("3D Images") == "3D Images" and conv.manual_number_level("3D Images") is None


def test_caption_forms():
    from sjuoz.parse import CAPTION_RE, SUBFIG_RE
    assert CAPTION_RE.match("Fig (1): Workflow").group(2, 3) == ("1", "Workflow")
    assert CAPTION_RE.match("Table (2): Values").group(2) == "2"
    assert SUBFIG_RE.match("Figure 1-A: Gel").group(1, 2) == ("1", "A")
    assert SUBFIG_RE.match("Figure 1- Overview of the method") is None


def test_superscript_minus_is_typeset():
    assert conv.tex_escape("mg g⁻¹") == "mg g$^{-}$¹"


def test_inline_math_does_not_look_lost():
    class P:
        tokens = [("t", "Where ", None), ("m", "y", False, None), ("t", " is the final prediction from the model.", None)]
    segs = report.text_segments(P())
    out = report._letters("Where y is the final prediction from the model.")
    assert report.found_ratio("Where  is the final prediction from the model.", out) < 0.7
    assert report.found_ratio_segments(segs, out) == 1.0


# ---------------------------------------------------------------- uzun denklemler
def test_equation_break_points_respect_groups():
    from sjuoz.eqbreak import break_points
    e = r"f\left( a+b \right)=\frac{x+y}{2}+c-d"
    kinds = [(e[p:p + 1], k) for p, k in break_points(e)]
    assert kinds == [("=", "rel"), ("+", "bin"), ("-", "bin")]  # parantez/kesir içindekiler hariç
    assert break_points("-x=-y") == [(2, "rel")]  # tekli eksi bölme yeri değil


def test_equation_candidates_align_on_relation():
    from sjuoz.eqbreak import candidate, fit
    two = candidate(r"F=a+b+c+d", 2)
    assert two.startswith(r"\begin{aligned}F &=") and r"\\ &\quad +" in two
    assert candidate(r"a=b+c=d-e", 2) == r"\begin{aligned}a &=b+c \\ &=d-e\end{aligned}"
    assert fit(r"P=\frac{a+b}{c}") == r"\sjuozfitone{P=\frac{a+b}{c}}"  # bölünemez: yalnız küçültme
    assert fit(r"F=a+b").startswith(r"\sjuozfit{F=a+b}")


def test_scaled_equations_in_log(tmp_path):
    log = tmp_path / "a.log"
    log.write_text("SJUOZ-EQ-SCALED: 5; 142.5pt; 613.4pt\nSJUOZ-EQ-SCALED: 6; 142.5pt; 150.0pt\n", encoding="utf-8")
    res = conv.analyze_log(str(log), [])
    assert [(n, round(r, 2)) for n, r in res["eq_scaled"]] == [("5", 0.23), ("6", 0.95)]


def test_corporate_narrative_citation_links():
    class P:
        def __init__(self, t):
            self._t = t

        def text(self):
            return self._t
    lk = conv.Linker([P("Food and Agriculture Organization of the United Nations (2023). Stats."),
                      P("Smith, J. (2020). X.")], 0, 0, 0)
    t = "as described by Food and Agriculture Organization of the United Nations (2023) and Smith (2020)."
    assert [(t[a:b], k) for a, b, k in lk.spans(t)] == [
        ("Food and Agriculture Organization of the United Nations (2023)", "ref:1"), ("Smith (2020)", "ref:2")]


def test_caption_sentence_is_not_a_caption():
    from sjuoz.writer import Writer

    class P:
        style, jc, images = "Normal", "center", []

        def __init__(self, t):
            self._t = t

        def text(self):
            return self._t
    w = Writer.__new__(Writer)
    assert w.is_caption(P("Table 1 reports total annual production per crop.")) is None
    assert w.is_caption(P("Table 1: Estimated production values")) == "table"


# ---------------------------------------------------------------- dergi biçimi (bitmiş makalelerle karşılaştırma)
def _para(tokens, style="Normal"):
    from sjuoz.docx import Para
    p = Para.__new__(Para)
    p.tokens, p.style, p.num_id = tokens, style, None
    return p


def test_full_names_from_cover_letter():
    from sjuoz.parse import fuller_name
    assert fuller_name("Ban J. Khoshaba", "Ban Jibrael Khoshaba") == "Ban Jibrael Khoshaba"
    assert fuller_name("Shamil K. Ramadhan", "shamil Khadr Ramadhan") == "Shamil Khadr Ramadhan"
    assert fuller_name("Arzu Çığ", "Arzu Çığ") == "Arzu Çığ"
    assert fuller_name("A. B. Smith", "John Smith") == "A. B. Smith"          # sözcük sayısı farklı
    assert fuller_name("Ban J. Khoshaba", "Ban Karim Khoshaba") == "Ban J. Khoshaba"  # baş harf tutmuyor


def test_title_keeps_italic_species_name():
    from sjuoz.template import italic_tex
    b, bi = (True, False, None, None), (True, True, None, None)
    toks = [("t", "Properties of Zinnia Cultivars (", b), ("t", "Zinnia elegans", bi), ("t", ")", b)]
    assert italic_tex(toks, upper=True) == r"PROPERTIES OF ZINNIA CULTIVARS (\textit{Zinnia elegans})"
    assert italic_tex([("t", "Plain title", b)], upper=True) is None   # italik yoksa eski yol


def test_title_and_authors_in_one_paragraph_are_split():
    from sjuoz.parse import _split_at_break, _looks_like_authors
    f, sup = (True, False, None, None), (True, False, "superscript", None)
    p = _para([("t", "A TITLE", f), ("t", "\n", f), ("t", "\n", f), ("t", "Linus Labik", f), ("t", "1,*", sup),
               ("t", ", Felix Asare", f), ("t", "1", sup)])
    head, tail = _split_at_break(p)
    assert head.text().strip() == "A TITLE" and _looks_like_authors(tail)
    assert _split_at_break(_para([("t", "No break", f)]))[1] is None


def test_long_heading_styled_paragraph_is_body_text():
    assert _para([("t", "Charging Protocol:", (True, False, None, None))], "heading 2").heading_level() == 2
    assert _para([("t", "x " * 150, (False, False, None, None))], "heading 2").heading_level() is None


def test_unbolded_colon_keeps_heading_bold():
    p = _para([("t", "Nitrogen percentage (N%)", (True, False, None, None)), ("t", ":", (False, False, None, None)),
               ("t", " ", (False, False, None, None))])
    assert p.all_bold()


def test_cover_letter_title_over_two_lines():
    from sjuoz.parse import parse_cover_letter
    f = (False, False, None, None)
    lines = ["I am pleased to submit an original research manuscript entitled:",
             "Second-life Assessment of Salvaged 18650 Lithium-Ion Cells Using Capacity and Internal",
             "Resistance Metrics.", "for consideration for publication in the Science Journal of University of Zakho."]

    class Doc:
        def body_blocks(self):
            return [_para([("t", ln, f)]) for ln in lines]
    assert parse_cover_letter(Doc())["title"].endswith("Internal Resistance Metrics.")


def test_free_form_cover_letter_signature():
    from sjuoz.parse import parse_cover_letter
    f = (False, False, None, None)
    lines = ['We are pleased to submit our original research article titled "A Lightweight CNN" for publication.',
             "Sincerely,", "Kazheen Sileman Muhammad", "Department of Information Technology", "Kurdistan Region, Iraq",
             "Email: kazheen.sileman@auas.edu.krd", "Araz Rajab Abrahim", "Duhok Polytechnic University",
             "Email: araz.abrahim@dpu.edu.krd"]

    class Doc:
        def body_blocks(self):
            return [_para([("t", ln, f)]) for ln in lines]
    info = parse_cover_letter(Doc())
    assert info["title"] == "A Lightweight CNN"
    assert [(a["name"], a["email"], a["corresponding"]) for a in info["authors"]] == [
        ("Kazheen Sileman Muhammad", "kazheen.sileman@auas.edu.krd", True),
        ("Araz Rajab Abrahim", "araz.abrahim@dpu.edu.krd", False)]


# ---------------------------------------------------------------- derginin editör kuralları
def test_journal_name_case_in_references():
    from sjuoz.common import journal_case
    assert journal_case("Journal of clinical laboratory analysis") == "Journal of Clinical Laboratory Analysis"
    assert journal_case("Ieee Access") == "IEEE Access"
    assert journal_case("Annales d’endocrinologie") == "Annales d’Endocrinologie"


def test_reference_page_ranges_get_en_dash_body_does_not():
    from sjuoz.common import en_dash_ranges, render_plain
    assert en_dash_ranges("95(6), 546-556.") == "95(6), 546–556."
    for keep in ("2024-05-12", "ISSN 2663-628X", "+964-7719698865", "COVID-19", "978-3-032-15398-2"):
        assert en_dash_ranges(keep) == keep
    assert "24-59" in render_plain("(24-59 years)")              # gövdede dergi kısa çizgiyi korur
    assert "546–556" in render_plain("8(3), 546-556.", ref_mode=True)


def test_all_caps_keywords_title_case_with_acronyms():
    from sjuoz.template import keywords_case

    class A:
        keywords = "LITHIUM-ION BATTERIES, STATE OF HEALTH (SOH), INTERNAL RESISTANCE."
        abstract_plain = ["the State of Health (SoH) and DC internal resistance (DCIR)"]
    assert keywords_case(A()) == "Lithium-Ion Batteries, State of Health (SoH), Internal Resistance."


def test_duplicated_author_becomes_narrative_citation():
    from sjuoz.render import narrative_citations, EQ_REF
    f, it = (False, False, None, None), (False, True, None, None)
    out = narrative_citations([("t", "Hanna (Hanna ", f), ("t", "et al.,", it), ("t", " 2021) reported", f)])
    assert "".join(t[1] for t in out) == "Hanna et al., (2021) reported"   # dergideki gibi
    assert ("t", "et al.,", it) in out
    assert EQ_REF.sub(r"\1 (\2)", "given in Equation 2 and Eq. 3.") == "given in Equation (2) and Eq. (3)."
    assert EQ_REF.sub(r"\1 (\2)", "Equation (1) and Equation 2.5 m") == "Equation (1) and Equation 2.5 m"


def test_apa_initials_and_year_period():
    from sjuoz.writer import apa_authors
    f = (False, False, None, None)
    out = apa_authors([("t", "Philipps, P.J., & Strange, R.E. (1965) Methods A.B. 5(8), 209-344.", f)])
    assert out[0][1] == "Philipps, P. J., & Strange, R. E. (1965). Methods A.B. 5(8), 209-344."


def test_iraqi_phone_format_and_cover_letter_spelling():
    from sjuoz.parse import iraq_phone, fuller_name
    assert iraq_phone("+9647719698865") == "+964-7719698865"
    assert iraq_phone("07510344636") == "+964-7510344636"
    assert iraq_phone("+233 26 232 6483") is None
    assert fuller_name("Hêro Farhad Salah", "Hero Farhad Salah") == "Hero Farhad Salah"
    assert fuller_name("Felix Asirifi Asare", "Felix Asiriﬁ Asare") == "Felix Asirifi Asare"


def test_broken_characters_from_word():
    from sjuoz.common import fix_text, title_case
    assert fix_text("Quintanilla-Dom\\’\\inguez") == "Quintanilla-Domínguez"
    assert fix_text("Gonc¸alves") == "Gonçalves"
    assert fix_text("a; b") == "a; b"
    assert title_case("Ethical statement") == "Ethical Statement"
    assert title_case("CONFLICT OF INTEREST") == "Conflict of Interest"


def test_conflict_statement_from_cover_letter():
    from sjuoz.parse import Article, add_conflict_statement
    b = (True, False, None, None)
    art = Article()
    art.body = [_para([("t", "Funding:", b)]), _para([("t", "None.", (False, False, None, None))]),
                _para([("t", "Author Contributions:", b)]), _para([("t", "All authors.", (False, False, None, None))])]
    for p in art.body:
        p.style_id = p.jc = p.outline = None
        p.ilvl, p.bidi = 0, False
    add_conflict_statement(art)
    assert [p.text() for p in art.body][2:4] == ["Conflict of Interest:", COI]
    add_conflict_statement(art)                 # ikinci kez eklenmez
    assert len(art.body) == 6


COI = "The authors have no conflict of interest to declare concerning the publication of this paper."


def test_mendeley_cell_and_glued_star_mark():
    from lxml import etree
    from sjuoz.docx import _children, W
    from sjuoz.parse import split_marks
    tr = etree.fromstring('<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                          '<w:sdt><w:sdtContent><w:tc/></w:sdtContent></w:sdt><w:tc/><w:tc/></w:tr>')
    assert len(_children(tr, W + "tc")) == 3        # Mendeley atıf hücresi atlanmaz
    assert split_marks("1,2*") == ["1", "2", "*"]
    assert split_marks("1, *") == ["1", "*"]


def test_unicode_script_letters_and_broken_image():
    from sjuoz.common import tex_escape
    from sjuoz.writer import image_ok
    assert tex_escape("Lᵢ = wᵧᵢ, R²") == r"L\textsubscript{i} = w\textsubscript{γi}, R²"
    assert tex_escape("e⁻ᴱⁱ") == r"e$^{-}$\textsuperscript{Ei}"
    assert not image_ok(b"\x89PNG\r\n\x1a\n" + b"\x00" * 40)
