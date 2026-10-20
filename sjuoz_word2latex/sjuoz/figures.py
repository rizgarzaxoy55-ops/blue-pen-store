# -*- coding: utf-8 -*-
"""Resimler: düzen tablosu/alt şekiller, biçim dönüştürme, Word grafikleri."""
import os
import re
import shutil

from lxml import etree

from .common import (C_NS, local, NO_WINDOW, NS, render_plain, warn)
from .docx import Para, Table


def layout_table_images(tbl):
    """Tablo, resimleri yan yana dizmek için kullanılmışsa [(rid, etiket), ...] döndürür; değilse None.
    Etiket, resmin bulunduğu hücredeki ya da aynı sütunda altındaki hücredeki kısa metindir."""
    if not isinstance(tbl, Table):
        return None
    entries = []  # (sütun, [rid], metin)
    for row in tbl.rows:
        col = 0
        for paras, span, vmerge in row:
            ps = [p for p in paras if isinstance(p, Para)]
            rids = [r for p in ps for r in p.images]
            text = " ".join(p.text().strip() for p in ps).strip()
            entries.append((col, rids, text))
            col += span
    n_img = sum(1 for _, r, _ in entries if r)
    if not n_img or any(len(t) > 120 for _, _, t in entries):
        return None
    if sum(1 for _, r, t in entries if t and not r) > n_img:
        return None
    items = []  # [rid, etiket, sütun]
    for col, rids, text in entries:
        if rids:
            for k, rid in enumerate(rids):
                items.append([rid, text if k == 0 else "", col])
        elif text:
            for it in reversed(items):
                if it[2] == col and not it[1]:
                    it[1] = text
                    break
    return [(rid, label) for rid, label, _ in items]


def clean_sub_label(label):
    """'(a) Original image' -> 'Original image'; '(b)' -> ''."""
    label = re.sub(r"^\s*\(?\s*[a-hA-H]\s*[).:]\s*", "", label or "").strip()
    return render_plain(label) if label else ""


# --------------------------------------------------------------------------
# Resim dönüştürme ve Word grafikleri (chart)
# --------------------------------------------------------------------------
def convert_external(src, dst_png):
    """EMF/WMF/SVG/EPS gibi biçimleri PNG'ye çevirir (LibreOffice, Inkscape veya ImageMagick)."""
    import subprocess
    import tempfile
    tools = []
    for exe in ("soffice", "libreoffice"):
        if shutil.which(exe):
            tools.append(("lo", [shutil.which(exe), "--headless", "--convert-to", "png", "--outdir"]))
            break
    if shutil.which("inkscape"):
        tools.append(("ink", [shutil.which("inkscape"), src, "-o", dst_png]))
    for exe in ("magick", "convert"):
        if shutil.which(exe):
            tools.append(("im", [shutil.which(exe), "-density", "300", src, dst_png]))
            break
    for kind, cmd in tools:
        try:
            if kind == "lo":
                tmp = tempfile.mkdtemp()
                subprocess.run(cmd + [tmp, src], check=True, capture_output=True, **NO_WINDOW, timeout=180)
                out = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + ".png")
                if os.path.isfile(out):
                    shutil.move(out, dst_png)
                shutil.rmtree(tmp, ignore_errors=True)
            else:
                subprocess.run(cmd, check=True, capture_output=True, **NO_WINDOW, timeout=180)
            if os.path.isfile(dst_png) and os.path.getsize(dst_png) > 0:
                trim_whitespace(dst_png)
                return True
        except Exception:  # noqa
            continue
    return False


def trim_whitespace(png):
    """Dönüştürülen resmin etrafındaki boş beyaz alanı kırpar (Pillow varsa)."""
    try:
        from PIL import Image, ImageChops
        im = Image.open(png)
        rgb = im.convert("RGB")
        bbox = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255))).getbbox()
        if bbox and (bbox[2] - bbox[0]) * (bbox[3] - bbox[1]) < 0.95 * rgb.size[0] * rgb.size[1]:
            pad = 10
            bbox = (max(0, bbox[0] - pad), max(0, bbox[1] - pad),
                    min(rgb.size[0], bbox[2] + pad), min(rgb.size[1], bbox[3] + pad))
            im.crop(bbox).save(png)
    except Exception:  # noqa
        pass


def _chart_text(el):
    if el is None:
        return ""
    return "".join(el.itertext()).strip() if el.find(".//{%s}t" % NS["a"]) is None else \
        "".join(t.text or "" for t in el.iter("{%s}t" % NS["a"])).strip()


def _chart_cache(el):
    """c:cat / c:val / c:tx içindeki önbellek değerlerini (sırayla) döndürür."""
    if el is None:
        return []
    C = "{%s}" % C_NS
    pts = {}
    for pt in el.iter(C + "pt"):
        v = pt.find(C + "v")
        pts[int(pt.get("idx", len(pts)))] = v.text if v is not None else ""
    if pts:
        return [pts[k] for k in sorted(pts)]
    v = el.find(".//" + C + "v")
    return [v.text] if v is not None else []


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return float("nan")


def render_chart(xml_bytes, dst_png):
    """Word grafiğini (DrawingML chart) matplotlib ile PNG olarak çizer."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        warn("Word grafiği çizilemedi: matplotlib gerekli (pip install matplotlib).")
        return False
    C = "{%s}" % C_NS
    root = etree.fromstring(xml_bytes)
    chart = root.find(C + "chart")
    plot = chart.find(C + "plotArea") if chart is not None else None
    if plot is None:
        return False
    plt.rcParams.update({"font.family": "serif", "font.size": 9})
    fig, ax = plt.subplots(figsize=(5, 3.2), dpi=200)
    drawn = False
    for ctype in plot:
        name = local(ctype.tag)
        if not name.endswith("Chart"):
            continue
        sers = ctype.findall(C + "ser")
        if not sers:
            continue
        if name in ("pieChart", "pie3DChart", "doughnutChart"):
            s0 = sers[0]
            cats = _chart_cache(s0.find(C + "cat"))
            vals = [_num(v) for v in _chart_cache(s0.find(C + "val"))]
            ax.pie(vals, labels=cats or None, autopct="%1.1f%%", startangle=90,
                   wedgeprops={"width": 0.45} if name == "doughnutChart" else None)
            ax.axis("equal")
            drawn = True
            continue
        bar_dir = wval_c(ctype.find(C + "barDir"), "col")
        grouping = wval_c(ctype.find(C + "grouping"), "clustered")
        n = len(sers)
        bottoms = None
        for k, ser in enumerate(sers):
            label = (_chart_cache(ser.find(C + "tx")) or ["Series %d" % (k + 1)])[0]
            if name in ("scatterChart", "bubbleChart"):
                xs = [_num(v) for v in _chart_cache(ser.find(C + "xVal"))]
                ys = [_num(v) for v in _chart_cache(ser.find(C + "yVal"))]
                style = wval_c(ctype.find(C + "scatterStyle"), "marker")
                line = "-" if "line" in style.lower() or "smooth" in style.lower() else ""
                ax.plot(xs, ys, line + "o", markersize=3, label=label)
                drawn = True
                continue
            cats = _chart_cache(ser.find(C + "cat"))
            vals = [_num(v) for v in _chart_cache(ser.find(C + "val"))]
            if not cats:
                cats = [str(i + 1) for i in range(len(vals))]
            xs = list(range(len(vals)))
            if name in ("barChart", "bar3DChart"):
                stacked = grouping in ("stacked", "percentStacked")
                width = 0.8 if stacked else 0.8 / n
                pos = xs if stacked else [x - 0.4 + width * (k + 0.5) for x in xs]
                if bottoms is None or not stacked:
                    bottoms_k = [0] * len(vals)
                else:
                    bottoms_k = bottoms
                if bar_dir == "bar":
                    ax.barh(pos, vals, height=width, left=bottoms_k, label=label)
                    ax.set_yticks(xs)
                    ax.set_yticklabels(cats)
                else:
                    ax.bar(pos, vals, width=width, bottom=bottoms_k, label=label)
                    ax.set_xticks(xs)
                    ax.set_xticklabels(cats, rotation=0 if len(cats) < 8 else 45, ha="center" if len(cats) < 8 else "right")
                if stacked:
                    bottoms = [a + b for a, b in zip(bottoms_k, vals)]
            elif name in ("areaChart", "area3DChart"):
                ax.fill_between(xs, vals, alpha=0.5, label=label)
                ax.set_xticks(xs)
                ax.set_xticklabels(cats)
            else:  # lineChart, radarChart vb.
                marker = "o" if ser.find(C + "marker") is None or \
                    wval_c(ser.find("%smarker/%ssymbol" % (C, C)), "auto") != "none" else ""
                ax.plot(xs, vals, "-" + marker, markersize=3, label=label)
                ax.set_xticks(xs)
                ax.set_xticklabels(cats, rotation=0 if len(cats) < 8 else 45, ha="center" if len(cats) < 8 else "right")
            drawn = True
        if (n > 1 or chart.find(C + "legend") is not None) and name not in ("pieChart", "pie3DChart"):
            ax.legend(frameon=False, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.18),
                      ncol=min(n, 4))
    if not drawn:
        plt.close(fig)
        return False
    title = chart.find(C + "title")
    auto_del = wval_c(chart.find(C + "autoTitleDeleted"), "0") in ("1", "true")
    if title is not None and not auto_del:
        t = _chart_text(title.find(C + "tx"))
        if t:
            ax.set_title(t, fontsize=10)
    axes = {wval_c(a.find(C + "axPos"), ""): _chart_text(a.find("%stitle/%stx" % (C, C)))
            for a in list(plot.findall(C + "catAx")) + list(plot.findall(C + "valAx"))}
    if axes.get("b"):
        ax.set_xlabel(axes["b"])
    if axes.get("l"):
        ax.set_ylabel(axes["l"])
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    fig.tight_layout()
    fig.savefig(dst_png)
    plt.close(fig)
    return True


def wval_c(el, default=None):
    if el is None:
        return default
    v = el.get("val")
    return default if v is None else v
