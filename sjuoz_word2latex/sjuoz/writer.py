# -*- coding: utf-8 -*-
"""Makale gövdesinin LaTeX'e yazılması (bölümler, şekiller, tablolar, denklemler)."""
import hashlib
import os
import re


from .common import (BIDI, is_rtl_text, tex_escape, warn)
from .docx import Para, Table
from .render import render_tokens, strip_eq_number_text
from .parse import (BACK_MATTER, CAPTION_RE, SUBFIG_RE, clean_heading, manual_number_level)
from .figures import (clean_sub_label, convert_external, layout_table_images, placeholder_png,
                      render_chart)
from .template import FIGDIR, LANDSCAPE_COLS, LONG_TABLE_ROWS

# --------------------------------------------------------------------------
# LaTeX üretimi
# --------------------------------------------------------------------------


class Writer:
    def __init__(self, art, doc, outdir, opts):
        self.art = art
        self.doc = doc
        self.outdir = outdir
        self.opts = opts
        self.out = []
        self.fig_count = 0
        self.caption_sub = {}      # başlık -> parça harfi ("Figure 1-A" -> "A")
        self.sub_targets = set()   # bağlanabilir parça şekiller: "1-A", "1-B"
        self.tab_count = 0
        self.eq_count = 0
        self.eq_display = 0         # numaralı + numarasız ayrı satır denklemler
        self.fig_captioned = 0
        self.caption_plain = {}     # LaTeX başlık -> düz metin
        self.figure_log = []        # [{"number", "files", "caption", "labels"}]
        self.table_log = []         # [düz başlık]
        self.pending_imgs = []      # [[rid,...], ...]
        self.pending_landscape = False
        self.pending_fig_caption = None
        self.pending_fig_para = None
        self.pending_tab_caption = None
        self.pending_tab_para = None
        self.in_multicol = False
        self.in_landscape = False
        self.list_env = None
        self.section = None         # o anki bölüm başlığı (küçük harf)
        self.contrib_items = []
        self.saved_hashes = {}

    # ---- çıktı yardımcıları
    def emit(self, s=""):
        self.out.append(s)

    def begin_landscape(self):
        if not self.in_landscape:
            self.close_cols()
            self.emit(r"\begin{landscape}")
            self.in_landscape = True

    def end_landscape(self):
        if self.in_landscape:
            self.emit(r"\end{landscape}")
            self.in_landscape = False

    def wants_landscape(self, block, ncols=0, kind="table"):
        number = str((self.tab_count if kind == "table" else self.fig_count) + 1)
        manual = self.opts.landscape_tables if kind == "table" else set()
        return bool(getattr(block, "landscape", False) or number in manual
                    or (kind == "table" and ncols >= LANDSCAPE_COLS))

    def open_cols(self):
        self.end_landscape()
        if not self.in_multicol:
            self.emit(r"\begin{multicols}{2}")
            self.in_multicol = True

    def close_cols(self):
        self.close_list()
        if self.in_multicol:
            self.emit(r"\end{multicols}")
            self.in_multicol = False

    def close_list(self):
        if self.list_env:
            self.emit(r"\end{%s}" % self.list_env)
            self.list_env = None
        self.flush_contrib()

    def flush_contrib(self):
        if self.contrib_items:
            items = [re.sub(r"^[\s\"“”'‘’]+|[\s\"“”'‘’.,;]+$", "", x) for x in self.contrib_items]
            items = [self.abbreviate_names(x).strip(" ,;") for x in items if x]
            self.emit(r"\noindent " + "; ".join(items) + ".")
            self.emit()
            self.contrib_items = []

    # ---- resimler
    def save_image(self, rid, name=None):
        """Resmi figures/ klasörüne kaydeder; LaTeX'te kullanılacak yolu döndürür."""
        path, data = self.doc.media(rid)
        if data is None:
            warn("Resim bulunamadı: %s" % rid)
            return None
        h = hashlib.md5(data).hexdigest()
        if h in self.saved_hashes:
            return self.saved_hashes[h]
        figdir = os.path.join(self.outdir, FIGDIR)
        os.makedirs(figdir, exist_ok=True)
        ext = os.path.splitext(path)[1].lower()
        name = name or "Figure%d" % (self.fig_count + 1)
        if any(v == FIGDIR + "/" + name + e for v in self.saved_hashes.values()
               for e in (".png", ".jpg", ".jpeg", ".pdf")):
            name += "_%d" % (len(self.saved_hashes) + 1)  # aynı şekilde birden fazla resim
        fname = None
        if rid.startswith("chart:"):
            if render_chart(data, os.path.join(figdir, name + ".png")):
                fname = name + ".png"
                print("Bilgi: Word grafiği resme çevrildi: %s" % fname)
            else:
                warn("Word grafiği çizilemedi (%s)." % path)
                return None
        elif ext in (".png", ".jpg", ".jpeg", ".pdf"):
            fname = name + ext
            with open(os.path.join(figdir, fname), "wb") as f:
                f.write(data)
        elif ext in (".tif", ".tiff", ".gif", ".bmp", ".webp", ".jfif"):
            fname = name + ".png"
            try:
                import io
                from PIL import Image
                im = Image.open(io.BytesIO(data))
                if im.mode not in ("RGB", "RGBA", "L"):
                    im = im.convert("RGBA")
                im.save(os.path.join(figdir, fname))
            except Exception as e:  # noqa
                fname = None
                src = os.path.join(figdir, name + ext)
                with open(src, "wb") as f:
                    f.write(data)
                if convert_external(src, os.path.join(figdir, name + ".png")):
                    os.remove(src)
                    fname = name + ".png"
                else:
                    warn("%s resmi PNG'ye çevrilemedi (pip install Pillow): %s" % (path, e))
                    fname = name + ext
        else:  # .emf, .wmf, .svg, .eps ...
            src = os.path.join(figdir, name + ext)
            with open(src, "wb") as f:
                f.write(data)
            if convert_external(src, os.path.join(figdir, name + ".png")):
                os.remove(src)
                fname = name + ".png"
                print("Bilgi: %s resmi PNG'ye çevrildi: %s" % (ext, fname))
            else:
                # Orijinali klasörde kalır (elle çevrilebilir); LaTeX takılmasın diye yer tutucu konur
                placeholder_png(os.path.join(figdir, name + ".png"), name + ext)
                fname = name + ".png"
                warn("%s%s PNG'ye çevrilemedi, yerine yer tutucu kondu; LibreOffice kurun ya da %s/%s%s"
                     " dosyasını PNG olarak kaydedip %s/%s yerine koyun." % (name, ext, FIGDIR, name, ext, FIGDIR, fname))
        rel = FIGDIR + "/" + fname
        self.saved_hashes[h] = rel
        return rel

    def prescan(self):
        """Bağlantılar için şekil/tablo/denklem sayılarını önceden bulur."""
        n_fig = n_tab = n_eq = 0
        for b in self.art.body:
            if isinstance(b, Table):
                n_tab += 0 if layout_table_images(b) else 1
            elif not b.is_empty():
                kind = self.is_caption(b)
                sub = SUBFIG_RE.match(b.text().strip())
                if kind == "figure" and sub and sub.group(2).upper() != "A":
                    self.sub_targets.add("%d-%s" % (n_fig, sub.group(2).upper()))  # 1-B: numara artmaz
                elif kind == "figure":
                    n_fig += 1
                    if sub:
                        self.sub_targets.add("%d-A" % n_fig)
                elif b.maths and kind is None:
                    ok, num = strip_eq_number_text(b.text().strip())
                    if ok and (num or any(m_[3] for m_ in b.maths)):
                        n_eq += 1
        return n_fig, n_tab, n_eq

    def emit_figure(self, items, caption, landscape=False):
        """items: [(rid, etiket), ...] — bir şekle ait bir veya birden çok resim (alt şekiller)."""
        items = [it if isinstance(it, tuple) else (it, "") for it in items]
        wide = "all" in self.opts.wide_figs or str(self.fig_count + 1) in self.opts.wide_figs
        if landscape:
            self.begin_landscape()
            wide = True
        elif wide:
            self.end_landscape()
            self.close_cols()
        else:
            self.open_cols()
            self.close_list()
            if self.out and self.out[-1] == r"\begin{multicols}{2}" and len(self.out) > 1:
                # İki sütunlu bölüm bir şekille başlıyor (genelde tablodan sonra): sayfada şekle
                # yetecek yer yoksa bölüm yeni sayfada başlasın; yoksa şekil sayfadan taşar.
                self.out[-1] = (r"\setlength{\premulticols}{0.45\textheight}\begin{multicols}{2}"
                                r"\setlength{\premulticols}{50pt}")
        files = []
        letters = "abcdefghijklmnopqrstuvwxyz"
        part = self.caption_sub.get(caption) if caption else None  # "Figure 1-B" gibi parça şekil
        same_number = bool(part and part != "A" and self.fig_count)
        number = self.fig_count if same_number else self.fig_count + 1
        for rid, label in items:
            name = "Figure%d-%s" % (number, part) if part else None
            if len(items) > 1:  # alt şekiller: Figure3a.png, Figure3b.png ...
                name = "Figure%d%s%s" % (number, "-" + part if part else "", letters[min(len(files), 25)])
            fn = self.save_image(rid, name)
            if fn and fn not in [f for f, _ in files]:
                files.append((fn, label))
        if not files:
            return
        self.fig_count = number
        box = r"\linewidth"  # sütunda sütun, geniş/yatay şekilde sayfa genişliği
        self.emit(r"\noindent\begin{minipage}{%s}" % box)
        if part:
            # "Figure 1-A": numara Word'deki gibi kalır; 1-A ile 1-B aynı numarayı paylaşır
            self.emit(r"\setcounter{figure}{%d}\renewcommand{\thefigure}{%d-%s}\renewcommand{\theHfigure}{%d-%s}"
                      % (number - 1, number, part, number, part))
        targets = ([] if same_number else ["fig:%d" % number]) + (["fig:%d-%s" % (number, part)] if part else [])
        self.emit(r"\centering" + "".join(r"\hypertarget{%s}{}" % t for t in targets))
        if len(files) == 1:
            self.emit(r"\includegraphics[width=%s, height=%s\textheight, keepaspectratio]{%s}"
                      % (box, "0.45" if wide else "0.42", files[0][0]))
        else:
            self.emit(r"\captionsetup{type=figure}")
            labels = [clean_sub_label(lb) for _, lb in files]
            # (a)(b) etiketleri: Word'de etiket varsa veya başlık "(a) … (b) …" diyorsa
            use_sub = any(labels) or bool(caption and re.search(r"\(\s*[a-h]\s*\)|(?:^|\s)[a-h]\)\s", caption))
            per_row = min(len(files), 4) if wide else 2
            width = 0.98 / per_row - 0.02
            height = "0.3" if per_row <= 2 else "0.22"
            for k, (fn, _) in enumerate(files):
                img = r"\includegraphics[width=\linewidth, height=%s\textheight, keepaspectratio]{%s}" % (height, fn)
                if use_sub:
                    self.emit(r"\begin{subfigure}[t]{%.2f%s}\centering" % (width, box))
                    self.emit(img)
                    self.emit(r"\caption{%s}" % labels[k])
                    self.emit(r"\end{subfigure}%")
                else:
                    self.emit(r"\begin{minipage}[t]{%.2f%s}\centering" % (width, box))
                    self.emit(img)
                    self.emit(r"\end{minipage}%")
                if k + 1 == len(files):
                    self.emit(r"\par\vspace{2pt}")
                elif (k + 1) % per_row == 0:
                    self.emit(r"\par\vspace{4pt}")
                else:
                    self.emit(r"\hfill")
        self.figure_log.append({"number": self.fig_count, "files": [f for f, _ in files],
                                "caption": self.caption_plain.get(caption, "") if caption else "",
                                "labels": [clean_sub_label(lb) for _, lb in files]})
        if caption:
            self.fig_captioned += 1
            # alt şekilli kutuda tür zaten 'figure' olarak ayarlı: \caption kullanılır
            self.emit((r"\caption{%s}" if len(files) > 1 else r"\captionof{figure}{%s}") % caption)
        elif len(files) > 1:
            self.emit(r"\addtocounter{figure}{1}")  # başlıksız alt şekilli şekil de numara alır
        self.emit(r"\end{minipage}")
        if part:
            self.emit(r"\renewcommand{\thefigure}{\arabic{figure}}\renewcommand{\theHfigure}{\arabic{figure}}")
        self.emit(r"\par\vspace{6pt}")
        self.emit()

    def flush_images(self):
        for rids in self.pending_imgs:
            self.emit_figure(rids, None, self.pending_landscape)
        self.pending_imgs = []
        self.pending_landscape = False

    # ---- tablolar
    def emit_table(self, tbl, caption):
        ncols = tbl.ncols()
        if self.wants_landscape(tbl, ncols):
            self.begin_landscape()  # yatay sayfa: genişlik = sayfanın uzun kenarı
        else:
            self.end_landscape()
            self.close_cols()
        self.tab_count += 1
        self.table_log.append(self.caption_plain.get(caption, "") if caption else "")
        if ncols == 0:
            return
        body_cells = [c for r in tbl.rows[1:] for c in r]
        texts = [" ".join(p.text() for p in c[0] if isinstance(p, Para)).strip() for c in body_cells]
        numeric = sum(1 for t in texts if re.fullmatch(r"[-+]?[\d.,%\s]*", t)) / max(1, len(texts))
        if numeric > 0.6:
            colspec = r"@{}>{\raggedright\arraybackslash}X" + r">{\centering\arraybackslash}X" * (ncols - 1) + "@{}"
            size = r"\footnotesize" if ncols >= 6 else r"\small"
        else:
            colspec = r"@{}" + r">{\raggedright\arraybackslash}X" * ncols + "@{}"
            size = r"\small"
        lines = []
        # başlık satırları: ilk satır + dikey birleştirilmiş hücresi devam eden satırlar
        n_head = 1
        while n_head < len(tbl.rows) - 1 and any(c[2] == "continue" for c in tbl.rows[n_head]):
            n_head += 1
        for ri, row in enumerate(tbl.rows):
            cells = []
            for paras, span, vmerge in row:
                if vmerge == "continue":
                    content = ""
                else:
                    parts = []
                    for p in paras:
                        if not isinstance(p, Para) or p.is_empty():
                            continue
                        for rid in p.images:  # hücre içindeki resimler
                            fn = self.save_image(rid, "Table%d_img%d" % (self.tab_count, len(self.saved_hashes) + 1))
                            if fn:
                                parts.append(r"\includegraphics[width=\linewidth, keepaspectratio]{%s}" % fn)
                        if p.text().strip() or p.maths:
                            parts.append(render_tokens(p.tokens))
                    content = r" \newline ".join(parts)
                if ri < n_head and content and not content.startswith(r"\textbf"):
                    content = r"\textbf{%s}" % content
                if span > 1:
                    align = "c" if (numeric > 0.6 or ri < n_head) else "l"
                    content = r"\multicolumn{%d}{%s}{%s}" % (span, align, content)
                cells.append(content)
            line = " & ".join(cells) + r" \\"
            if ri == n_head - 1:
                line += r" \midrule"
            elif ri < len(tbl.rows) - 1 and ri >= n_head:
                line += r" \addlinespace[0.3em]"
            lines.append(line)
        long_table = len(tbl.rows) > LONG_TABLE_ROWS
        if long_table:
            # sayfalara bölünebilen tablo: başlık her sayfada tekrarlanır
            head = lines[:n_head]
            head[-1] = head[-1].replace(r" \midrule", "")
            self.emit(r"\begingroup")
            self.emit(size)
            self.emit(r"\setlength{\LTcapwidth}{\linewidth}")
            self.emit(r"\begin{xltabular}{\linewidth}{%s}" % colspec)
            self.emit((r"\caption{%s}\hypertarget{tab:%d}{}\\" % (caption, self.tab_count)) if caption
                      else r"\hypertarget{tab:%d}{}\\" % self.tab_count)
            self.emit(r"\toprule")
            self.emit("\n".join(head) + r" \midrule")
            self.emit(r"\endfirsthead")
            self.emit(r"\multicolumn{%d}{@{}l}{\textbf{Table \thetable:} \textit{(continued)}} \\[2pt]" % ncols)
            self.emit(r"\toprule")
            self.emit("\n".join(head) + r" \midrule")
            self.emit(r"\endhead")
            self.emit(r"\midrule \multicolumn{%d}{@{}r@{}}{\textit{(continued on next page)}} \\" % ncols)
            self.emit(r"\endfoot")
            self.emit(r"\bottomrule")
            self.emit(r"\endlastfoot")
            self.emit("\n".join(lines[n_head:]))
            self.emit(r"\end{xltabular}")
            self.emit(r"\endgroup")
            self.emit(r"\par\vspace{6pt}")
            self.emit()
            return
        self.emit(r"\noindent")
        self.emit(r"\begin{minipage}{\linewidth}")
        self.emit(r"\centering")
        self.emit(r"\hypertarget{tab:%d}{}" % self.tab_count)
        if caption:
            self.emit(r"\captionof{table}{%s}" % caption)
        self.emit(r"\vspace{-6pt}")
        self.emit(size)
        self.emit(r"\begin{tabularx}{\linewidth}{%s}" % colspec)
        self.emit(r"\toprule")
        self.emit("\n".join(lines))
        self.emit(r"\bottomrule")
        self.emit(r"\end{tabularx}")
        self.emit(r"\end{minipage}")
        self.emit(r"\par\vspace{6pt}")
        self.emit()

    # ---- denklemler
    def emit_equation(self, latex, number):
        self.open_cols()
        self.close_list()
        self.eq_display += 1
        if number:
            self.eq_count += 1
            tag = "" if number == str(self.eq_count) else r" \tag{%s}" % number
            if tag:
                warn("Denklem numarası sırası farklı: beklenen %d, Word'de (%s)." % (self.eq_count, number))
            self.emit(r"\hypertarget{eq:%d}{}" % self.eq_count)
            self.emit(r"\begin{equation}")
            self.emit("  " + latex + tag)
            self.emit(r"\end{equation}")
        else:
            self.emit(r"\begin{equation*}")
            self.emit("  " + latex)
            self.emit(r"\end{equation*}")

    # ---- caption
    def cap_of(self, p):
        """LaTeX başlığı döndürür; düz metnini (JATS/Crossref için) saklar."""
        latex = self.caption_text(p)
        mt = CAPTION_RE.match(p.text().strip())
        sub = SUBFIG_RE.match(p.text().strip())
        plain = p.text().strip()[sub.end():] if sub else (mt.group(3) if mt else "")
        self.caption_plain[latex] = re.sub(r"\s+", " ", plain).strip(" .:") + "." if mt else ""
        if sub:
            self.caption_sub[latex] = sub.group(2).upper()
        return latex

    @staticmethod
    def caption_text(p):
        toks = list(p.tokens)
        # "Figure 1" önekini token'lardan kaldır
        full = p.text()
        mt = CAPTION_RE.match(full)
        sub = SUBFIG_RE.match(full)
        cut = sub.end() if sub else len(full) - len(mt.group(3))
        out, pos = [], 0
        for t in toks:
            if t[0] != "t":
                out.append(t)
                continue
            s = t[1]
            if pos + len(s) <= cut:
                pos += len(s)
                continue
            if pos < cut:
                s = s[cut - pos:]
                pos = cut
            pos += len(t[1])
            # caption'da kalın/italik genel biçimi kaldır (sadece alt/üst simge kalsın)
            out.append(("t", s, (False, t[2][1] and not all(x[2][1] for x in toks if x[0] == "t"), t[2][2], t[2][3])))
        cap = render_tokens(out)
        cap = re.sub(r"^[:.\-–—\s]+", "", cap).strip()
        if cap and cap[-1] not in ".?!":
            cap += "."
        return cap

    def flush_orphan_figure_caption(self):
        """Resmin üstüne yazılmış sanılan başlıktan sonra resim değil metin geldi: başlık sonraki bir
        resme yapışıp bütün şekil numaralarını kaydırmasın (ör. PDF'te tabloya dönüşmüş resim)."""
        if self.pending_fig_caption is None:
            return
        p = self.pending_fig_para
        warn("Şekil başlığından önce/sonra resim bulunamadı, başlık metin olarak bırakıldı: %s" % p.text().strip()[:80])
        self.pending_fig_caption = self.pending_fig_para = None
        self.open_cols()
        self.emit(r"\noindent\textbf{%s}" % render_tokens(p.tokens))
        self.emit()

    def flush_orphan_table_caption(self):
        """Başlıktan sonra tablo gelmedi (ör. PDF'te çizgisiz tablo düz metne dönüştü): başlık
        sonraki bir tabloya yanlışlıkla bağlanmasın, metin olarak yazılsın."""
        if self.pending_tab_caption is None:
            return
        p = self.pending_tab_para
        warn("Tablo başlığından sonra tablo bulunamadı, başlık metin olarak bırakıldı: %s" % p.text().strip()[:80])
        self.pending_tab_caption = self.pending_tab_para = None
        self.open_cols()
        self.emit(r"\noindent\textbf{%s}" % render_tokens(p.tokens))
        self.emit()

    def caption_in_first_row(self, tbl):
        """Başlık tablonun içinde, bütün sütunları kaplayan ilk satırda olabilir ("Table (1): …").
        Bulunursa o satır tablodan çıkarılır ve başlık olarak döner."""
        if len(tbl.rows) < 2 or len(tbl.rows[0]) != 1:
            return None
        paras = [p for p in tbl.rows[0][0][0] if isinstance(p, Para) and not p.is_empty()]
        if len(paras) != 1 or self.is_caption(paras[0]) != "table":
            return None
        tbl.rows = tbl.rows[1:]
        return self.cap_of(paras[0])

    def is_caption(self, p):
        txt = p.text().strip()
        mt = CAPTION_RE.match(txt)
        if not mt or len(txt) > 400:
            return None
        rest = mt.group(3)
        style = p.style.lower()
        if ("caption" in style or p.jc == "center" or p.images or
                re.match(r"^\s*(figure|fig\.?|table)\s*\(?\s*\d+\s*\)?\s*[:.\-–—]", txt, re.I)):
            return "table" if mt.group(1).lower().startswith("tab") else "figure"
        # Gövde cümlesi mi? ("Figure 5 shows ...")
        if rest and rest[0].islower():
            return None
        return None

    # ---- ana gövde
    def run(self):
        art = self.art
        self.open_cols()
        blocks = art.body
        k = 0
        while k < len(blocks):
            b = blocks[k]
            k += 1
            if isinstance(b, Table) and layout_table_images(b):
                # resim yerleştirmek için kullanılan kenarlıksız tablo: tablo değil, şekildir
                items = layout_table_images(b)
                if self.pending_fig_caption is not None:
                    self.emit_figure(items, self.pending_fig_caption, getattr(b, "landscape", False))
                    self.pending_fig_caption = None
                else:
                    self.pending_imgs.append(items)
                    self.pending_landscape |= getattr(b, "landscape", False)
                continue
            if isinstance(b, Table):
                self.flush_images()
                cap = self.pending_tab_caption
                self.pending_tab_caption = None
                if cap is None:  # tablonun altındaki başlık
                    j = k
                    while j < len(blocks) and isinstance(blocks[j], Para) and blocks[j].is_empty():
                        j += 1
                    if j < len(blocks) and isinstance(blocks[j], Para) and self.is_caption(blocks[j]) == "table":
                        cap = self.cap_of(blocks[j])
                        k = j + 1
                if cap is None:
                    cap = self.caption_in_first_row(b)
                if cap is None:
                    warn("Tablo %d için başlık (caption) bulunamadı." % (self.tab_count + 1))
                self.emit_table(b, cap)
                continue
            p = b
            if p.is_empty():
                continue
            txt = p.text().strip()

            # Resimler
            if p.images:
                rids = []
                for rid in p.images:
                    _, data = self.doc.media(rid)
                    h = hashlib.md5(data or b"").hexdigest()
                    if h not in [hashlib.md5((self.doc.media(r)[1] or b"")).hexdigest() for r in rids]:
                        rids.append(rid)
                if self.pending_fig_caption is not None:
                    self.emit_figure(rids, self.pending_fig_caption, getattr(p, "landscape", False))
                    self.pending_fig_caption = None
                else:
                    self.pending_imgs.append(rids)
                    self.pending_landscape |= getattr(p, "landscape", False)
                if not txt:
                    continue

            # Şekil / tablo başlıkları
            kind = self.is_caption(p)
            if kind == "figure":
                cap = self.cap_of(p)
                if self.pending_imgs:
                    # başlıktan önceki bütün resimler bu şekle aittir (alt şekiller)
                    self.emit_figure([it for grp in self.pending_imgs for it in grp], cap, self.pending_landscape)
                    self.pending_imgs = []
                    self.pending_landscape = False
                else:
                    self.flush_orphan_figure_caption()
                    self.pending_fig_caption = cap  # başlık resmin üstünde olabilir
                    self.pending_fig_para = p
                continue
            if kind == "table":
                self.flush_orphan_table_caption()
                self.pending_tab_caption = self.cap_of(p)
                self.pending_tab_para = p
                continue
            self.flush_orphan_table_caption()
            self.flush_orphan_figure_caption()

            # Denklemler
            if p.maths:
                ok, num_txt = strip_eq_number_text(txt)
                if ok:
                    maths = p.maths
                    num = num_txt or next((m_[3] for m_ in maths if m_[3]), None)
                    latex = r" \quad ".join(m_[1] for m_ in maths if m_[1])
                    self.emit_equation(latex, num)
                    continue

            # Başlıklar
            if self.handle_heading(p, txt):
                continue

            # Liste maddeleri
            if p.num_id and p.num_id != "style":
                if self.section and re.search(r"contribution", self.section):
                    self.contrib_items.append(render_tokens(p.tokens))
                    continue
                fmt = self.doc.numfmt.get(p.num_id, {}).get(str(p.ilvl), "decimal")
                env = "itemize" if fmt in ("bullet", "none") else "enumerate"
                if self.list_env != env:
                    self.close_list()
                    self.emit(r"\begin{%s}" % env)
                    self.list_env = env
                self.emit(r"\item " + render_tokens(p.tokens))
                continue
            self.close_list()

            # Normal paragraf
            self.open_cols()
            if p.bidi or is_rtl_text(txt):
                # sağdan sola paragraf (Arapça / Kürtçe Sorani)
                BIDI["par"] = True
                try:
                    s = render_tokens(p.tokens)
                finally:
                    BIDI["par"] = False
                BIDI["used"] = True
                self.emit(r"\begin{rtlpar}%s\end{rtlpar}" % s)
                self.emit()
                continue
            s = render_tokens(p.tokens)
            if self.section and re.search(r"contribution", self.section):
                s = self.abbreviate_names(s)
            self.emit(s)
            self.emit()
        self.flush_images()
        self.flush_orphan_table_caption()
        self.flush_orphan_figure_caption()
        self.close_list()

    def handle_heading(self, p, txt):
        level = p.heading_level()
        numbered = True
        clean = clean_heading(txt)
        if level is None and not p.maths and len(txt) <= 150 and p.all_bold():
            if p.num_id:
                level = p.ilvl + 1
            elif manual_number_level(txt):
                level = manual_number_level(txt)
            elif BACK_MATTER.match(clean) or re.match(r"^conclusions?$", clean, re.I):
                level, numbered = 2, False
            elif not txt.rstrip().endswith((".", ",", ";")):
                level, numbered = 2, False
        elif level is None and BACK_MATTER.match(clean) and len(txt) < 60:
            level, numbered = 2, False
        if level is None:
            return False
        if BACK_MATTER.match(clean):
            level, numbered = 2, False
        self.flush_images()
        self.close_list()
        self.open_cols()
        self.section = clean.lower()
        if level == 1:
            title = clean if self.opts.keep_heading_case else clean.upper()
            self.emit(r"\section%s{%s}" % ("" if numbered else "*", tex_escape(title)))
        elif level == 2:
            self.emit(r"\subsection%s{%s}" % ("" if numbered else "*", tex_escape(clean)))
        else:
            self.emit(r"\subsubsection%s{%s}" % ("" if numbered else "*", tex_escape(clean)))
        return True

    def abbreviate_names(self, s):
        for au in self.art.authors:
            s = abbreviate_in_text(s, au["name"])
        s = s.replace("“", "").replace("”", "")
        return s

    def references(self):
        if not self.art.references:  # boş list ortamı LaTeX'te "missing \item" hatası verir
            return
        self.open_cols()
        self.close_list()
        self.emit(r"\section*{REFERENCES}")
        self.emit(r"\begin{list}{}{")
        self.emit(r"    \setlength{\leftmargin}{1.5em}")
        self.emit(r"    \setlength{\itemindent}{-1.5em}")
        self.emit(r"    \setlength{\parsep}{0pt}")
        self.emit(r"    \setlength{\itemsep}{6pt}")
        self.emit(r"}")
        for k, p in enumerate(self.art.references):
            s = r"\hypertarget{ref:%d}{}" % (k + 1) + render_tokens(p.tokens, ref_mode=True)
            s = re.sub(r"^\s*\[?\d+[\].]\s+", "", s) if self.opts.strip_ref_numbers else s
            self.emit(r"\item " + s)
            self.emit()
        self.emit(r"\end{list}")


def initials(name):
    parts = name.split()
    if len(parts) < 2:
        return name
    ini = []
    for x in parts[:-1]:
        ini.append(x if re.fullmatch(r"[A-Z]\.", x) else x[0] + ".")
    return " ".join(ini + [parts[-1]])


def abbreviate_in_text(s, full):
    short = initials(full)
    parts = full.split()
    s = re.sub(re.escape(full) + r"\s*\d*(?=\W|$)", short, s)
    if len(parts) >= 3:  # "Eyad I." gibi eksik yazılmış adlar
        prefix = " ".join(parts[:-1])
        s = re.sub(re.escape(prefix) + r"(?!\s*" + re.escape(parts[-1]) + r")", short, s)
    return s
