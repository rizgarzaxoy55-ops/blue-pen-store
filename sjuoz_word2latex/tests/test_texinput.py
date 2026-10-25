# -*- coding: utf-8 -*-
"""LaTeX girişi: ön işleme birimleri ve .tex / .zip'ten uçtan uca dönüşüm."""
import os
import re
import zipfile

import pytest

from conftest import CASES, FIXTURES, has_pandoc, run_case
from sjuoz import texinput as tx

needs_pandoc = pytest.mark.skipif(not has_pandoc(), reason="pandoc yok (pip install pypandoc_binary)")


def test_macros_and_revision_marks():
    pre = r"\newcommand{\mycite}[1]{(\citeauthor{#1}, \citeyear{#1})}\newcommand{\rev}[1]{\textcolor{red}{#1}}" \
          r"\newcommand{\R}{\mathbb{R}}"
    body = r"see \mycite{Bank2014} and \rev{Papua \textbf{New}} in $\R^n$, not \Rbig"
    out = tx.strip_revision_marks(tx.expand_macros(body, tx.user_macros(pre)))
    assert out == r"see \citep{Bank2014} and Papua \textbf{New} in $\mathbb{R}^n$, not \Rbig"


def test_front_matter_authblk_and_and():
    pre = r"""\title{\vspace{-1cm}\textbf{\uppercase{A Study}}}
\author[1]{Ali Hassan}\author[2, *]{Sara Ahmed}
\affil[1]{Dept A, Univ X.}\affil[2]{Dept B, Univ Y.}"""
    title, authors, affils, block = tx.front_matter(pre)
    assert title == "A Study" and authors == [("Ali Hassan", ["1"]), ("Sara Ahmed", ["2", "*"])]
    assert r"Ali Hassan\textsuperscript{1} and Sara Ahmed\textsuperscript{2,*}" in block
    assert r"\textsuperscript{2}Dept B, Univ Y." in block
    _, authors, _, _ = tx.front_matter(r"\title{T}\author{John Smith\inst{1} \and Jane Doe\inst{2}}")
    assert authors == [("John Smith", ["1"]), ("Jane Doe", ["2"])]


def test_numbering_and_references():
    body = r"""\section{Intro}\label{sec:a}
\begin{equation} a = b \label{eq:a} \end{equation}
\begin{align} x &= 1 \label{eq:x} \\ y &= 2 \nonumber \\ z &= 3 \label{eq:z} \end{align}
\begin{equation*} q = r \end{equation*}
\begin{table}[h]\centering\caption{Data}\label{tab:d}\begin{tabular}{l}A\end{tabular}\end{table}
\begin{figure}[h]\centering\includegraphics{p.png}\caption{Plot}\label{fig:p}\end{figure}
See \eqref{eq:z}, Table~\ref{tab:d}, Figure~\ref{fig:p}, Section~\ref{sec:a}, \autoref{eq:a}."""
    labels = {}
    tx.number_sections(body, labels)
    body = tx.number_floats(body, labels)
    body = tx.number_equations(body, labels)
    body = tx.resolve_refs(body, labels)
    flat = re.sub(r"\s+", " ", body)
    assert r"$$a = b \tag{1}$$" in flat and r"$$x = 1 \tag{2}$$" in flat
    assert "$$y = 2$$" in flat  # \nonumber: numarasız
    assert r"$$z = 3 \tag{3}$$" in flat and r"$$q = r$$" in flat
    assert r"\noindent Table 1: Data" in body and r"\noindent Figure 1: Plot" in body
    assert body.index("Table 1: Data") < body.index(r"\begin{tabular}")  # tablo başlığı üstte
    assert r"See \href{@@eq:3}{(3)}, Table~1, Figure~1, Section~1, Equation 1." in body


def test_bbl_fallback_writes_natbib_citations():
    items = tx.bbl_items(r"""\begin{thebibliography}{2}
\bibitem[Smith and Doe(2020)]{s20} Smith, J. and Doe, J. (2020). \newblock A paper. \newblock J, 1.
\bibitem[{Van Hai et~al.(2022)}]{v22} Van Hai, V. (2022). Clusters.
\end{thebibliography}""")
    assert [k for k, _, _ in items] == ["s20", "v22"] and items[0][2].startswith("Smith, J. and Doe")
    out = tx.natbib_cites(r"\citep{s20} and \citet{v22}; \citep{s20,v22}", items)
    assert out == "(Smith and Doe, 2020) and Van Hai et al. (2022); (Smith and Doe, 2020; Van Hai et al., 2022)"


def test_math_marker_survives_word_splitting():
    tx.TEX_MATH["/m.docx"] = {12: (r"\text{orig}", "7")}
    # Word denkleminde işaret parçalara bölünebilir
    assert tx.restore_math("/m.docx", r"\mathrm{sjuoztex}\mathrm{bc}q", None) == (r"\text{orig}", "7")
    assert tx.restore_math("/other.docx", "x", None) == ("x", None)


LATEX = next(c for c in CASES if c["name"] == "sentetik_latex")


@needs_pandoc
def test_tex_end_to_end(tmp_path):
    res = run_case(LATEX, tmp_path / "o")
    with open(res["tex_path"], encoding="utf-8") as f:
        tex = f.read()
    # yazarın denklemleri birebir (Word denklemine çevrilip bozulmadan)
    assert r"\epsilon_i \sim N(0, \sigma^2), \quad \text{with } \sigma = 0.11\,\mu_i" in tex
    assert r"\hyperlink{eq:2}{(2)}" in tex and r"\hyperlink{tab:1}{Table~1}" in tex
    assert "textcolor" not in tex and "Van Hai \\& Silhavy (2022)" in tex
    assert not [i for i in res["items"] if i["level"] == "HATA"], res["items"]


@needs_pandoc
def test_zip_with_nested_zip(tmp_path):
    folder = os.path.join(FIXTURES, "sentetik_latex")
    inner = tmp_path / "kaynak.zip"
    with zipfile.ZipFile(inner, "w") as z:
        for n in ("main.tex", "refs.bib", "plot.png"):
            z.write(os.path.join(folder, n), "proje/" + n)
    outer = tmp_path / "gonderim.zip"
    with zipfile.ZipFile(outer, "w") as z:
        z.write(inner, "kaynak.zip")
        z.writestr("not.txt", "PDF yalnız bilgi için")
    assert tx.is_tex_input(str(outer))
    case = dict(LATEX, docx=[str(outer)])
    res = run_case(case, tmp_path / "o")
    assert res["stats"]["references"] == 2 and res["stats"]["numbered_equations"] == 4
