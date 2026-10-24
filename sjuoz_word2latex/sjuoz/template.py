# -*- coding: utf-8 -*-
"""SJUOZ LaTeX şablonu ve ön bilgilerin yazılması."""

import re

from .common import (BIDI, render_plain, tex_escape, title_case_caps, url_escape)

# --------------------------------------------------------------------------
# Şablon
# --------------------------------------------------------------------------
TEMPLATE = r"""%!TEX TS-program = xelatex
% Bu dosya sjuoz_word2latex.py tarafından otomatik üretildi.
\documentclass[9pt]{extarticle}
\usepackage{multicol}
\raggedcolumns  % sütun boyu doldurulmaz: başlık ve paragraf aralarında büyük boşluk açılmaz
\usepackage[top=0.5in, bottom=0.75in, right=0.75in, left=0.75in, headheight=0.25in, includehead]{geometry}
\usepackage{fontspec}
\usepackage{amsmath}
\usepackage{amssymb}
\usepackage{graphicx}
% Denklem sütuna sığmazsa: önce 2, sonra 3 satıra bölünmüş biçimi; o da sığmazsa küçültülür
\newsavebox{\sjuozeqbox}
\newlength{\sjuozeqmax}
\newcommand{\sjuozeqset}{\setlength{\sjuozeqmax}{\linewidth}\addtolength{\sjuozeqmax}{-2.5em}}
\newcommand{\sjuozeqput}{\ifdim\wd\sjuozeqbox>\sjuozeqmax
  \typeout{SJUOZ-EQ-SCALED: \theequation; \the\sjuozeqmax; \the\wd\sjuozeqbox}% kontrol raporu için
  \resizebox{\sjuozeqmax}{!}{\usebox{\sjuozeqbox}}\else\usebox{\sjuozeqbox}\fi}
\newcommand{\sjuozfit}[3]{\sjuozeqset\sbox{\sjuozeqbox}{$\displaystyle #1$}%
  \ifdim\wd\sjuozeqbox>\sjuozeqmax\sbox{\sjuozeqbox}{$\displaystyle #2$}\fi
  \ifdim\wd\sjuozeqbox>\sjuozeqmax\sbox{\sjuozeqbox}{$\displaystyle #3$}\fi\sjuozeqput}
\newcommand{\sjuozfitone}[1]{\sjuozeqset\sbox{\sjuozeqbox}{$\displaystyle #1$}\sjuozeqput}
\usepackage{xcolor}
\usepackage{array}
\usepackage{tabularx}
\usepackage{xltabular}
\usepackage{pdflscape}
\usepackage{booktabs}
\usepackage{multirow}
\usepackage{caption}
\usepackage{subcaption}
\usepackage{titlesec}
\usepackage{fancyhdr}
\usepackage{qrcode}
\usepackage{xurl}
\usepackage{lastpage}
\usepackage{indentfirst}
\usepackage{eso-pic}
\usepackage{tikz}
\usepackage[colorlinks=true, urlcolor=blue, linkcolor=<<LINKCOLOR>>, citecolor=<<LINKCOLOR>>]{hyperref}

\titleformat{\section}{\large\bfseries\centering}{\thesection.}{1.0em}{}
\titleformat{\subsection}{\large\bfseries}{\thesubsection.}{0.5em}{}
\titleformat{\subsubsection}{\large\bfseries}{\thesubsubsection.}{0.5em}{}
\titlespacing*{\section}{0pt}{8pt}{6pt}
\titlespacing*{\subsection}{0pt}{8pt}{4pt}
\titlespacing*{\subsubsection}{0pt}{6pt}{3pt}

\newcommand{\titlesize}{\fontsize{12}{10}}
\renewcommand{\normalsize}{\fontsize{9pt}{9pt}\selectfont}
\renewcommand{\large}{\fontsize{10pt}{10pt}\selectfont}
\renewcommand{\small}{\fontsize{8pt}{8pt}\selectfont}
\DeclareCaptionLabelFormat{sjuozparen}{#1 (#2)}
\captionsetup[table]{labelfont=bf, justification=raggedright, singlelinecheck=false}
\captionsetup[figure]{labelfont=bf, justification=centering, singlelinecheck=false}
\captionsetup[sub]{labelformat=parens, labelfont=bf, font=small, justification=centering}
\definecolor{sjuozblue}{HTML}{C1D1EB}
\definecolor{colorabs}{HTML}{E2EAF4}

% Yazı tipleri: Times New Roman / Garamond yoksa benzerleri kullanılır
\IfFontExistsTF{<<MAINFONT>>}{\setmainfont{<<MAINFONT>>}}{%
  \IfFontExistsTF{TeX Gyre Termes}{\setmainfont{TeX Gyre Termes}}{%
    \IfFontExistsTF{Liberation Serif}{\setmainfont{Liberation Serif}}{}}}
\IfFontExistsTF{<<HEADFONT>>}{\newfontfamily{\garamond}{<<HEADFONT>>}}{%
  \IfFontExistsTF{EB Garamond}{\newfontfamily{\garamond}{EB Garamond}}{\newcommand{\garamond}{\rmfamily}}}

<<RTL>>\newcommand{\sjuozdoi}{<<DOI>>}
% İlk sayfadaki bilgi kutusunun sayfa altından yüksekliği ve metin için ayrılan boşluk
\newcommand{\firstfooterbottom}{0.4in}
\newcommand{\firstpagereserve}{1.4in}
% Sadece ilk sayfada metin alanını kısalt; ilk sayfa basıldıktan sonra eski haline döndür
\makeatletter
\newlength{\sjuozfulltextheight}
\AtBeginDocument{%
  \setlength{\sjuozfulltextheight}{\textheight}%
  \global\advance\textheight by -\firstpagereserve
  \global\@colht\textheight \global\@colroom\textheight \global\vsize\textheight
  \AddToHookNext{shipout/after}{\global\textheight\sjuozfulltextheight}}
\makeatother

% Başlık ve altlık (header/footer)
\pagestyle{fancy}
\fancyhf{}
\renewcommand{\headrulewidth}{0pt}
\lhead{\large \textit{<<RUNHEAD>>}}
\rhead{\small{\garamond{\textbf{SJUOZ|VOL<<VOL>>|<<MON3>> <<YEAR>>|P<<FIRSTPAGE>>-<<LASTPAGE>>}}}}
\lfoot{\raisebox{8pt}{\textit{DOI:}\href{\sjuozdoi}{\textit{\small{\sjuozdoi}}}}}
\cfoot{\raisebox{8pt}{\thepage}}

\fancypagestyle{firstpage}{%
  \fancyhf{}
  \renewcommand{\headrulewidth}{0pt}
  \rhead{\IfFileExists{SJUOZ.png}{\includegraphics[width=1.04in]{SJUOZ.png}}{\textbf{SJUOZ}}}
  \lhead{\colorbox{sjuozblue}{%
      \begin{minipage}[b]{0.83\textwidth}
        \vspace{0.15cm}
        \hspace{-3pt}\ \normalsize \noindent \garamond{Science Journal of University of Zakho}\\
        \normalsize \textit{\garamond{DOI: \sjuozdoi}}
        \vspace{0.2cm}
      \end{minipage}}}
}

% İlk sayfanın altındaki bilgi kutusu (sayfaya mutlak konumla yerleştirilir)
\newcommand{\sjuozfirstfooter}{%
  \begin{minipage}[b]{\textwidth}
      \rule{\textwidth}{1.5pt} \\[0.15cm]
      \begin{minipage}[t]{0.2\textwidth}\hspace{1pt} \small Access this article online \\
        \vspace{0.1cm}
        {\hypersetup{urlcolor=black}\qrcode[height=2.6162cm]{<<ARTICLEURL>>}}
      \end{minipage}\begin{minipage}[t]{0.4\textwidth}
        \small DOI:\href{\sjuozdoi}{\sjuozdoi} \\ \\[0.5cm]
        Printed ISSN 2663-628X; \\
        Electronic ISSN 2663-6298
      \end{minipage}\begin{minipage}[t]{0.4\textwidth}\small \textbf{Science Journal of University of Zakho} \\
        \textbf{Vol. <<VOL>>, No. <<ISSUE2>>, pp. <<FIRSTPAGE>>-<<LASTPAGE>>, <<MONTH>>-<<YEAR>>} \\[0.5cm]
        This is an open access article under a CC BY-NC-SA 4.0 license \\
        (\href{https://creativecommons.org/licenses/by-nc-sa/4.0/}{https://creativecommons.org/licenses/by-nc-sa/4.0/})
      \end{minipage} \\[0.15cm]
      \rule{\textwidth}{1.5pt} \\[0.15cm]
      \centering \normalsize \thepage
  \end{minipage}}

% ORCID simgesi: orcid.png varsa o, yoksa çizilmiş yeşil "iD" simgesi
\definecolor{orcidgreen}{HTML}{A6CE39}
\newcommand{\orcidicon}{\IfFileExists{orcid.png}{\includegraphics[width=9pt]{orcid.png}}{%
  \tikz[baseline=-0.7ex]{\fill[orcidgreen] (0,0) circle (4.5pt);
  \node[text=white, font=\fontsize{4.5pt}{4.5pt}\selectfont\bfseries, inner sep=0pt] at (0,0) {iD};}}}
\newcommand{\orcidlink}[1]{\href{#1}{\orcidicon}}

\begin{document}
\setcounter{page}{<<PAGESTART>>}
\thispagestyle{firstpage}
\AddToShipoutPictureFG*{\put(\LenToUnit{0.75in},\LenToUnit{\firstfooterbottom}){\sjuozfirstfooter}}

\noindent \large \textbf{<<ARTICLETYPE>>} \\[-5pt]
\begin{center} \titlesize \textbf{<<TITLE>>} \end{center}
\begin{center} \large \textbf{<<AUTHORS>>}
\end{center}
\noindent
<<AFFILIATIONS>>

\begin{center} <<CORRESPONDING>>
\end{center}
\rule{\textwidth}{1.5pt}
\begin{minipage}[t]{0.175\textwidth}
  \hspace{0cm} \vspace{0cm}
\end{minipage} \begin{minipage}[t]{0.8\textwidth}
\large \textbf{ABSTRACT}\end{minipage} \\[-6pt]
\rule{\textwidth}{1.5pt} \\
\begin{minipage}[t]{0.175\textwidth}
  \vspace{1cm} Received \\ <<RECEIVED>> \\[0.5cm] Accepted \\ <<ACCEPTED>> \\[0.5cm] Published \\ <<PUBLISHED>> \\[0.25cm]
\end{minipage} \colorbox{colorabs}{\begin{minipage}[t]{0.81\textwidth}
  \large{<<ABSTRACT>>
  \\[2pt]
  \normalsize \textbf{Keywords:} <<KEYWORDS>>
  }
\end{minipage}}
\\[-10pt]

\normalsize
\noindent \rule{\textwidth}{0.75pt}
<<BODY>>

\end{document}
"""

FIGDIR = "figures"  # resimlerin konduğu alt klasör
LONG_TABLE_ROWS = 34  # bundan uzun tablolar sayfalara bölünür (xltabular)
LANDSCAPE_COLS = 10   # bu kadar veya daha çok sütunlu tablolar yatay sayfaya konur

RTL_PREAMBLE = r"""% Sağdan sola yazılar (Arapça, Kürtçe Sorani, Farsça): XeTeX TeX--XeT
\TeXXeTstate=1
\IfFontExistsTF{Amiri}{\newfontfamily\arabicfont[Script=Arabic]{Amiri}}{%
  \IfFontExistsTF{Noto Naskh Arabic}{\newfontfamily\arabicfont[Script=Arabic]{Noto Naskh Arabic}}{%
    \IfFontExistsTF{Times New Roman}{\newfontfamily\arabicfont[Script=Arabic]{Times New Roman}}{%
      \newfontfamily\arabicfont[Script=Arabic]{DejaVu Sans}}}}
\newcommand{\rtl}[1]{{\arabicfont\beginR #1\endR}}
\newcommand{\ltr}[1]{{\rmfamily\beginL #1\endL}}
\newenvironment{rtlpar}{\par\raggedleft\arabicfont\beginR\ignorespaces}{\unskip\endR\par}

"""

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August",
          "September", "October", "November", "December"]


def link_color(name):
    """xcolor renk adı (blue, red, teal, blue!70!black …); geçersizse siyah."""
    name = (name or "black").strip()
    return name if re.fullmatch(r"[A-Za-z]+(?:![0-9]{1,3}(?:![A-Za-z]+)?)*", name) else "black"


def keywords_case(art):
    """TAMAMI BÜYÜK HARF anahtar kelimeler dergide Başlık Biçimindedir; kısaltmalar özetteki yazılışını
    alır: "STATE OF HEALTH (SOH)" → "State of Health (SoH)"."""
    kw = art.keywords.strip()
    if not kw or kw.upper() != kw or not re.search(r"[A-Z]{2}", kw):
        return kw
    forms = {}
    for w in re.findall(r"[A-Za-z][A-Za-z0-9]*", " ".join(getattr(art, "abstract_plain", []))):
        if re.search(r"[A-Z]", w[1:]):
            forms.setdefault(w.lower(), w)
    return re.sub(r"[A-Za-z][A-Za-z0-9]*", lambda m: forms.get(m.group(0).lower(), m.group(0)), title_case_caps(kw))


def italic_tex(tokens, upper=False):
    """Düz metin; yalnız italik parçalar (ör. tür adları "Zinnia elegans") \\textit{} olarak ve
    büyük harfe çevrilmeden korunur. İçinde küçük harfli italik parça yoksa None."""
    runs = []
    for t in tokens:
        if t[0] != "t":
            continue
        txt = re.sub(r"\s+", " ", t[1])
        it = bool(t[2][1]) and bool(txt.strip())
        if runs and runs[-1][1] == it:
            runs[-1][0] += txt
        else:
            runs.append([txt, it])
    if not any(it and re.search(r"[a-z]", txt) for txt, it in runs):
        return None
    out = []
    for txt, it in runs:
        if it:
            core = txt.strip(" ")
            out.append(txt[:len(txt) - len(txt.lstrip(" "))] + r"\textit{%s}" % tex_escape(core)
                       + txt[len(txt.rstrip(" ")):])
        else:
            out.append(tex_escape(txt.upper() if upper else txt))
    return re.sub(r"\s+", " ", "".join(out)).strip()


def build_tex(art, writer, opts):
    # yazarlar
    au_tex = []
    for au in art.authors:
        s = tex_escape(au["name"])
        marks = au["marks"]   # dergi tek affiliasyonda da numarayı yazar ("Muhammed 1,*")
        if marks:
            s += r" \textsuperscript{%s}" % ",".join(marks)
        if au.get("orcid"):
            s += r"\orcidlink{%s}" % url_escape(au["orcid"])
        au_tex.append(s)
    if len(au_tex) > 2:
        authors = ", ".join(au_tex[:-1]) + ", and " + au_tex[-1]
    else:
        authors = " and ".join(au_tex)
    affs = "\n".join(r"\large \textsuperscript{%d}%s\\" % (k + 1, tex_escape(a))
                     for k, a in enumerate(art.affiliations))
    corr = r"\textsuperscript{*}\textit{Corresponding author, E-mail:}"
    if art.corr_email:
        corr += r"\href{mailto:%s}{%s}" % (art.corr_email, tex_escape(art.corr_email))
    if art.corr_tel:
        corr += r" (%s: %s)" % (getattr(art, "corr_tel_label", "Tel"), tex_escape(art.corr_tel))

    surnames = [au["name"].split()[-1] for au in art.authors if au["name"].split()]
    if len(surnames) >= 3:
        runhead = tex_escape(surnames[0]) + r" \textit{et al.}"
    elif len(surnames) == 2:
        runhead = r"%s \& %s" % (tex_escape(surnames[0]), tex_escape(surnames[1]))
    else:
        runhead = tex_escape(surnames[0]) if surnames else ""

    abstract = r" \par ".join(art.abstract) or "~"
    kw = italic_tex(getattr(art, "keyword_tokens", [])) or render_plain(keywords_case(art))
    if kw and kw[-1] not in ".":
        kw += "."

    doi = "https://doi.org/10.25271/sjuoz.%s.%s.%s.%s" % (opts.year, opts.volume, opts.issue, opts.article_id)
    if not opts.first_page:
        last = "xxx"            # ilk sayfa bilinmiyorsa dergideki gibi "pp. xxx-xxx"
    else:
        last = str(opts.last_page) if opts.last_page else r"\pageref*{LastPage}"
    rep = {
        "MAINFONT": opts.main_font,
        "LINKCOLOR": link_color(getattr(opts, "link_color", None)),
        "HEADFONT": opts.header_font,
        "DOI": doi,
        "RUNHEAD": runhead,
        "VOL": str(opts.volume),
        "ISSUE2": "%02d" % int(opts.issue),
        "YEAR": str(opts.year),
        "MONTH": opts.month,
        "MON3": opts.month[:3].upper(),
        "FIRSTPAGE": str(opts.first_page) if opts.first_page else "xxx",
        "PAGESTART": str(opts.first_page or 1),
        "LASTPAGE": last,
        "ARTICLEURL": "https://sjuoz.uoz.edu.krd/index.php/sjuoz/article/view/%s" % opts.article_id,
        "ARTICLETYPE": tex_escape(opts.article_type or art.article_type),
        "TITLE": italic_tex(getattr(art, "title_tokens", []), upper=not opts.keep_title_case)
        or tex_escape(art.title.upper() if not opts.keep_title_case else art.title),
        "AUTHORS": authors,
        "AFFILIATIONS": affs,
        "CORRESPONDING": corr,
        "RECEIVED": tex_escape(opts.received),
        "ACCEPTED": tex_escape(opts.accepted),
        "PUBLISHED": tex_escape(opts.published),
        "ABSTRACT": abstract,
        "KEYWORDS": kw,
        "BODY": "\n".join(writer.out),
        "RTL": RTL_PREAMBLE if BIDI["used"] else "",
    }
    tex = TEMPLATE
    for k, v in rep.items():
        tex = tex.replace("<<%s>>" % k, v)
    return tex
