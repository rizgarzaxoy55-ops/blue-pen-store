#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Yapay test makalelerini (.docx) üretir. Sadece test setini yeniden oluşturmak için
gerekir; üretilen .docx dosyaları depoda hazır durur. Gerekenler: pandoc,
(EMF için) LibreOffice.

    python tests/fixtures/make_synthetic.py
"""
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, "..", "..", "ornek_cikti_1993", "figures", "Figure1.png")

TEMEL = r"""
**A STUDY OF SOMETHING\_WITH & SPECIAL % CHARS**

**Ali Hassan^1,\*^, Sara Ahmed^2^ and John Smith^1^**

^1^Department of Physics, University of Zakho, Duhok, Iraq.

^2^Department of Math, University of Duhok, Duhok, Iraq.

\* Corresponding author. E-mail: ali.hassan@uoz.edu.krd (Tel.: +964 750 123 4567)

**Abstract**

We study $E=mc^2$ and the fraction $\frac{a}{b}$ with 50% accuracy & more. The method is
evaluated on two datasets and compared with earlier work in detail.

**Keywords:** physics; math.

# Introduction

Intro text with a citation (Smith et al., 2020) and a narrative citation Brown and Lee (2019),
a URL https://example.com/test_page and an e-mail info@example.org. See Table 1 and Figure 1.

$$\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}$$

## Methods subsection

- first bullet with $\alpha \le \beta$
- second bullet

1. numbered one
2. numbered two

: Table 1 Sample results

| Name | Value | Error |
|------|------:|------:|
| A    | 1.5   | 0.1   |
| B    | 2.5   | 0.2   |

![](@@FIG@@)

Figure 1: A figure caption with $x_i$.

# Conclusions

Done. The results agree with (Brown & Lee, 2019; Smith et al., 2020).

**Acknowledgements**

Thanks to everyone who helped with the experiments in this study.

**References**

Brown, A., & Lee, B. (2019). Another paper on things. *Physics Letters*, 3, 10–20. https://doi.org/10.1000/abc

Smith, J., Doe, J., & Roe, R. (2020). A paper. *JOURNAL OF THE BEST THINGS EVER*, 1(2), 3–4. https://doi.org/https://doi.org/10.1000/xyz
"""

NUMERIK = r"""
Review Article

**DEEP LEARNING FOR CROP YIELD PREDICTION**

**Karwan Ali 1, Dilan Omer 1**

Department of Computer Science, Faculty of Science, University of Zakho, Duhok, Iraq.

\* Corresponding author: karwan.ali@uoz.edu.krd

Abstract: Crop yield prediction is studied with deep networks [1]. We compare three models on
public data and report the error of each model in a table with merged header cells.

Index Terms— deep learning; agriculture; regression

# INTRODUCTION

Prior work used classical regression [1], [2] and tree models [3]. Recent studies [2–3] show gains.
The model is $\hat{y} = W x + b$ and the loss is given in Equation 1.

$$L = \frac{1}{N}\sum_{i=1}^{N} \left( y_i - \hat{y}_i \right)^2 \qquad (1)$$

# METHODOLOGY

## Data

The data set has 1,200 samples.[^1]

[^1]: Collected in 2024.

## Models

+-----------+-----------------+
| Model     | Error           |
|           +--------+--------+
|           | MAE    | RMSE   |
+===========+========+========+
| CNN       | 0.21   | 0.30   |
+-----------+--------+--------+
| LSTM      | 0.19   | 0.27   |
+-----------+--------+--------+

: Table 1. Errors of the models.

# CONCLUSION

The LSTM model is best [3].

# REFERENCES

[1] A. Author, "First paper," *Journal A*, vol. 1, pp. 1–10, 2020.

[2] B. Author, "Second paper," *Journal B*, vol. 2, pp. 11–20, 2021.

[3] C. Author, "Third paper," *Journal C*, vol. 3, pp. 21–30, 2022. doi: 10.1000/ccc
"""

def uzun_tablo_md():
    rows = "\n".join("| S%02d | %d | %.3f | %.1f |" % (i, 700 + 13 * i, 0.1 + i / 97.0, 90 + (i % 10) / 2.0)
                     for i in range(1, 81))
    return r"""
**A LONG TABLE THAT SPANS SEVERAL PAGES**

**Rojin Saleh^1,\*^ and Hawar Aziz^1^**

^1^Department of Statistics, University of Zakho, Duhok, Iraq.

\* Corresponding author. E-mail: rojin.saleh@uoz.edu.krd

**Abstract**

This synthetic manuscript contains a table with eighty rows that cannot fit on one page and
must be split across pages without losing any row.

**Keywords:** long table; pagination.

# Introduction

The full results are listed in Table 1 (Saleh, 2024).

: Table 1 Results for all eighty speakers

| Speaker | Frames | Time (s) | Accuracy (%) |
|---------|-------:|---------:|-------------:|
""" + rows + r"""

# Conclusions

Every row of Table 1 must appear in the PDF.

**References**

Saleh, R. (2024). Long tables in journals. *Journal of Layout*, 2(1), 1–9.
"""


ALT_SEKIL = r"""
**SUBFIGURES PLACED IN DIFFERENT WAYS**

**Shilan Khalid^1,\*^ and Bawar Yousif^1^**

^1^Department of Computer Science, University of Zakho, Duhok, Iraq.

\* Corresponding author. E-mail: shilan.khalid@uoz.edu.krd

**Abstract**

Authors place sub-figures in Word in several ways: inside a borderless table with labels under
the pictures, as several pictures in one paragraph, or as pictures in separate paragraphs.

**Keywords:** subfigures; layout.

# Introduction

Figure 1 uses a layout table, Figure 2 has two pictures in one paragraph and Figure 3 has two
pictures in separate paragraphs.

| ![](IMG_A){width=1.4in} | ![](IMG_B){width=1.4in} |
|:---:|:---:|
| (a) Original image | (b) Detected mouth |

Figure 1: Pictures placed with a layout table.

![](IMG_C){width=1.3in} ![](IMG_D){width=1.3in}

Figure 2: (a) Left picture. (b) Right picture.

![](IMG_E){width=1.5in}

![](IMG_F){width=1.5in}

Figure 3: Two pictures in separate paragraphs.

# Conclusions

All pictures must stay with their captions (Khalid, 2025).

**References**

Khalid, S. (2025). Placing pictures in Word. *Journal of Layout*, 3(1), 1–5.
"""


def make_pictures(workdir):
    """Altı farklı, küçük test resmi (içerikleri farklı olmalı; aynı resimler birleştirilir)."""
    from PIL import Image, ImageDraw
    paths = {}
    colors = ["#d92d20", "#1570ef", "#079455", "#dc6803", "#7a5af8", "#475467"]
    for k, name in enumerate("ABCDEF"):
        im = Image.new("RGB", (240, 160), "white")
        d = ImageDraw.Draw(im)
        d.rectangle([10, 10, 230, 150], outline="black", width=3)
        d.ellipse([60 + 10 * k, 40, 180 - 5 * k, 120], fill=colors[k])
        d.text((20, 18), name, fill="black")
        path = os.path.join(workdir, "img_%s.png" % name.lower())
        im.save(path)
        paths["IMG_" + name] = path
    return paths


CHART_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>
<c:title><c:tx><c:rich><a:p><a:r><a:t>Accuracy by method</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>%s
</c:barChart><c:catAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Dataset</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx>
<c:valAx><c:axPos val="l"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Accuracy (%%)</a:t></a:r></a:p></c:rich></c:tx></c:title></c:valAx>
</c:plotArea><c:legend/></c:chart></c:chartSpace>"""


def chart_series(name, vals):
    cat = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, c) for i, c in enumerate("ABC"))
    val = "".join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(vals))
    return ('<c:ser><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>%s</c:v></c:pt></c:strCache></c:strRef></c:tx>'
            '<c:cat><c:strRef><c:strCache>%s</c:strCache></c:strRef></c:cat>'
            '<c:val><c:numRef><c:numCache>%s</c:numCache></c:numRef></c:val></c:ser>' % (name, cat, val))


A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def inline(graphic_data):
    return ('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>'
            '<wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
            '<wp:extent cx="3000000" cy="2000000"/><wp:docPr id="90" name="x"/><a:graphic %s>%s</a:graphic>'
            '</wp:inline></w:drawing></w:r></w:p>' % (A_NS, graphic_data))


def caption(text):
    return '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>' % text


def add_chart_and_emf(src, dst, emf):
    """pandoc çıktısına bir Word grafiği ve bir EMF resmi ekler (Conclusions başlığından önce)."""
    paras = (inline('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>'
                    '</a:graphicData>')
             + caption("Figure 2 Accuracy comparison chart.")
             + inline('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                      '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                      '<pic:nvPicPr><pic:cNvPr id="0" name="e"/><pic:cNvPicPr/></pic:nvPicPr>'
                      '<pic:blipFill><a:blip r:embed="rIdEmf1"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData>')
             + caption("Figure 3 An EMF picture."))
    chart = CHART_XML % (chart_series("Baseline", [90, 85, 70]) + chart_series("Proposed", [95, 92, 88]))
    zin = zipfile.ZipFile(src)
    zout = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    for it in zin.infolist():
        d = zin.read(it.filename)
        if it.filename == "word/document.xml":
            x = d.decode("utf-8")
            i = x.index("Conclusions")
            j = max(x.rfind("<w:p>", 0, i), x.rfind("<w:p ", 0, i))
            d = (x[:j] + paras + x[j:]).encode("utf-8")
        elif it.filename == "word/_rels/document.xml.rels":
            d = d.decode("utf-8").replace("</Relationships>", (
                '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/chart" Target="charts/chart1.xml"/>'
                '<Relationship Id="rIdEmf1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/image" Target="media/pic.emf"/></Relationships>')).encode("utf-8")
        elif it.filename == "[Content_Types].xml":
            d = d.decode("utf-8").replace("</Types>", (
                '<Override PartName="/word/charts/chart1.xml" ContentType="application/'
                'vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
                '<Default Extension="emf" ContentType="image/x-emf"/></Types>')).encode("utf-8")
        zout.writestr(it, d)
    zout.writestr("word/charts/chart1.xml", chart)
    zout.write(emf, "word/media/pic.emf")
    zout.close()


def pandoc(md, out, workdir, images=None):
    path = os.path.join(workdir, "in.md")
    for key, img in (images or {}).items():
        md = md.replace(key, img)
    with open(path, "w", encoding="utf-8") as f:
        f.write(md.replace("@@FIG@@", os.path.abspath(FIG)))
    subprocess.run(["pandoc", path, "-f", "markdown+grid_tables", "-o", out], check=True)


def make_emf(workdir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("EMF üretmek için LibreOffice gerekli.")
    src = os.path.join(workdir, "pic.png")
    shutil.copy(FIG, src)
    subprocess.run([soffice, "--headless", "--convert-to", "emf", "--outdir", workdir, src],
                   check=True, capture_output=True, timeout=180)
    return os.path.join(workdir, "pic.emf")


def main(only=None):
    tmp = tempfile.mkdtemp()
    try:
        for case, md in (("sentetik_temel", TEMEL), ("sentetik_numerik", NUMERIK),
                         ("sentetik_uzun_tablo", uzun_tablo_md())):
            os.makedirs(os.path.join(HERE, case), exist_ok=True)
            pandoc(md, os.path.join(HERE, case, "%s.docx" % case), tmp)
        os.makedirs(os.path.join(HERE, "sentetik_alt_sekil"), exist_ok=True)
        pandoc(ALT_SEKIL, os.path.join(HERE, "sentetik_alt_sekil", "sentetik_alt_sekil.docx"), tmp,
               make_pictures(tmp))
        os.makedirs(os.path.join(HERE, "sentetik_grafik"), exist_ok=True)
        base = os.path.join(tmp, "base.docx")
        pandoc(TEMEL, base, tmp)
        add_chart_and_emf(base, os.path.join(HERE, "sentetik_grafik", "sentetik_grafik.docx"), make_emf(tmp))
        print("Yapay test makaleleri üretildi.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
