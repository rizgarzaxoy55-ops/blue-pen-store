# -*- coding: utf-8 -*-
"""SJUOZ Word -> LaTeX dönüştürücü.

Makale ve kapak mektubu Word dosyalarını Science Journal of University of Zakho
(SJUOZ) şablonunda XeLaTeX'e çevirir; PDF, kontrol raporu, Crossref/JATS XML ve
BibTeX üretir.

Kullanım (Python'dan):

    from sjuoz import build_parser, convert
    sonuc = convert(build_parser().parse_args(["makale.docx", "kapak.docx", "--first-page", "571"]))
    print(sonuc["status"], sonuc["tex_path"])
"""
__version__ = "1.0.0"

import importlib.util as _ilu

if _ilu.find_spec("lxml") is None:  # pragma: no cover
    raise SystemExit("Hata: 'lxml' paketi gerekli. Kurulum: pip install lxml")

from .common import BIDI, DATA_DIR, WARNINGS, is_rtl_text, render_plain, tex_escape, title_case_caps, warn
from .omml import OMML
from .docx import Docx, Para, Table
from .render import Linker, render_tokens, strip_eq_number_text
from .parse import (Article, clean_heading, manual_number_level, match_author, merge_affiliations,
                    parse_article, parse_cover_letter, strip_titles)
from .writer import Writer, abbreviate_in_text, initials
from .figures import clean_sub_label, layout_table_images, render_chart
from .template import build_tex
from .export import (CROSSREF_NS, JOURNAL, article_metadata, bibtex_authors, bibtex_references,
                     bibtex_self, crossref_xml, jats_xml, parse_date, parse_reference, validate_crossref)
from .report import REPORT_CHECK, REPORT_ERROR, REPORT_OK, analyze_log, build_report, report_status
from .pipeline import build_parser, classify, convert
from .issue import convert_issue
from .cli import main

__all__ = [
    "__version__",
    'BIDI',
    'DATA_DIR',
    'WARNINGS',
    'is_rtl_text',
    'render_plain',
    'tex_escape',
    'title_case_caps',
    'warn',
    'OMML',
    'Docx',
    'Para',
    'Table',
    'Linker',
    'render_tokens',
    'strip_eq_number_text',
    'Article',
    'clean_heading',
    'manual_number_level',
    'match_author',
    'merge_affiliations',
    'parse_article',
    'parse_cover_letter',
    'strip_titles',
    'Writer',
    'abbreviate_in_text',
    'initials',
    'clean_sub_label',
    'layout_table_images',
    'render_chart',
    'build_tex',
    'CROSSREF_NS',
    'JOURNAL',
    'article_metadata',
    'bibtex_authors',
    'bibtex_references',
    'bibtex_self',
    'crossref_xml',
    'jats_xml',
    'parse_date',
    'parse_reference',
    'validate_crossref',
    'REPORT_CHECK',
    'REPORT_ERROR',
    'REPORT_OK',
    'analyze_log',
    'build_report',
    'report_status',
    'build_parser',
    'classify',
    'convert',
    'convert_issue',
    'main',
]
