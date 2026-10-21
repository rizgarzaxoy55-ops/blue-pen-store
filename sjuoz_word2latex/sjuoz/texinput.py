# -*- coding: utf-8 -*-
"""LaTeX girişi: yazarın .tex kaynağını (ya da onu içeren .zip'i) dönüştürücünün okuyabileceği .docx'e çevirir.

Yol: .tex → (ön işleme) → pandoc AST (JSON) → .docx → olağan Word hattı. Denklemler Word denklemine
çevrilip geri okunursa bozulur (\\text{…} → \\mathrm{…}\\ \\mathrm{…}); bu yüzden AST'deki her denklemin
yerine benzersiz bir işaret konur ve .docx okunurken (Docx.restore_math) yazarın özgün LaTeX'i geri yazılır.

Ön işlemede:
  * \\input/\\include içerikleri birleştirilir, yorumlar silinir;
  * başlık, yazarlar (authblk \\author[1]{…}/\\affil[1]{…} ya da \\author{A \\and B}) ve kurumlar
    dergi biçimindeki paragraflara çevrilir (\\maketitle yerine);
  * numaralı denklem/tablo/şekil/bölümler sayılır, \\label'lar numaralara, \\ref/\\eqref/\\autoref/\\cref
    metne çevrilir; çok satırlı numaralı align'ın her satırı ayrı numaralı denklem olur;
  * figure/table ortamları "Figure N: …"/"Table N: …" başlıklı düz bloklara çevrilir;
  * \\bibliography{x} → "References" başlığı; kaynaklar x.bib'den APA (CSL) ile pandoc citeproc'la yazılır;
    .bib yoksa .bbl/thebibliography girdileri kullanılır ve natbib atıfları elle yazılır.
Pandoc gerekir: pip install pypandoc_binary (pandoc'u içinde getirir) ya da sistemde pandoc.
"""
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

from .common import DATA_DIR, NO_WINDOW, warn

# Docx yolu -> {numara: özgün LaTeX}; Docx bu tabloyla denklemleri geri yükler
TEX_MATH = {}
# Yalnız harflerden işaret: rakamlar Word denkleminde ayrı parçaya düşüyor ("SJUOZTEX" "0")
MATH_MARK_RE = re.compile(r"sjuoztex([a-j]+)q")


def math_mark(n):
    return "sjuoztex" + "".join(chr(97 + int(d)) for d in str(n)) + "q"


def _mark_number(letters):
    return int("".join(str(ord(c) - 97) for c in letters))

NUMBERED_EQ_ENVS = ("equation", "align", "gather", "multline", "eqnarray", "flalign", "alignat")


def is_tex_input(path):
    low = path.lower()
    if low.endswith(".tex"):
        return True
    if low.endswith(".zip"):
        try:
            with zipfile.ZipFile(path) as z:
                return any(n.lower().endswith(".tex") for n in z.namelist()) or \
                    any(n.lower().endswith(".zip") for n in z.namelist())
        except zipfile.BadZipFile:
            return False
    return False


# --------------------------------------------------------------------------
# Yardımcılar: dengeli parantez, komut argümanları
# --------------------------------------------------------------------------
def group_end(s, i, open_ch="{", close_ch="}"):
    """s[i] == open_ch; eşleşen kapanışın konumu (yoksa -1)."""
    depth = 0
    k = i
    while k < len(s):
        c = s[k]
        if c == "\\":
            k += 2
            continue
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return k
        k += 1
    return -1


def command_args(s, pos, n_req=1, opt=True):
    """pos'tan itibaren [isteğe bağlı] ve {zorunlu} argümanları okur: (opts, args, son)."""
    opts, args = [], []
    k = last = pos  # last: son argümanın bittiği yer (sonraki boşluk metne aittir)
    while True:
        while k < len(s) and s[k] in " \t\n":
            k += 1
        if opt and k < len(s) and s[k] == "[" and not args:
            e = group_end(s, k, "[", "]")
            if e < 0:
                break
            opts.append(s[k + 1:e])
            k = last = e + 1
            continue
        if len(args) < n_req and k < len(s) and s[k] == "{":
            e = group_end(s, k)
            if e < 0:
                break
            args.append(s[k + 1:e])
            k = last = e + 1
            continue
        break
    return opts, args, last


def find_commands(s, name, n_req=1):
    """\\name[...]{...} geçişleri: [(başlangıç, son, opts, args)]."""
    out = []
    for mt in re.finditer(r"\\%s(?![A-Za-z@])\*?" % re.escape(name), s):
        opts, args, end = command_args(s, mt.end(), n_req)
        if len(args) == n_req:
            out.append((mt.start(), end, opts, args))
    return out


def remove_commands(s, name, n_req=1):
    for start, end, _, _ in reversed(find_commands(s, name, n_req)):
        s = s[:start] + s[end:]
    return s


def strip_comments(s):
    return "\n".join(re.sub(r"(?<!\\)%.*", "", ln) for ln in s.split("\n"))


def clean_inline(s):
    """Başlık/yazar/kurum metni için biçim komutlarını sadeleştirir (LaTeX olarak kalır)."""
    s = re.sub(r"\\(?:vspace|hspace)\*?\{[^{}]*\}", " ", s)
    s = remove_commands(s, "thanks")
    s = re.sub(r"\\(?:uppercase|MakeUppercase|textbf|bfseries|centering|large|Large|LARGE|huge|normalsize|small)\b",
               "", s)
    s = s.replace("\\\\", " ").replace("\\newline", " ")
    s = re.sub(r"[{}]", "", s)
    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------------------
# Okuma: zip, \input
# --------------------------------------------------------------------------
def locate_main_tex(path, workdir):
    """.tex ya da .zip (iç içe zip olabilir) → (ana .tex yolu, klasör)."""
    if path.lower().endswith(".tex"):
        return path, os.path.dirname(os.path.abspath(path))
    root = os.path.join(workdir, "kaynak")
    _extract(path, root)
    cands = []
    for d, _, files in os.walk(root):
        for f in files:
            if f.lower().endswith(".tex"):
                p = os.path.join(d, f)
                with open(p, encoding="utf-8", errors="replace") as fh:
                    t = fh.read()
                if "\\begin{document}" in t:
                    cands.append((len(t), p))
    if not cands:
        raise SystemExit("HATA: %s içinde \\begin{document} içeren .tex bulunamadı." % os.path.basename(path))
    main = max(cands)[1]
    if len(cands) > 1:
        print("Bilgi: zip'te birden çok ana .tex var; en büyüğü kullanıldı: %s" % os.path.basename(main))
    return main, os.path.dirname(main)


def _extract(zpath, dest, depth=0):
    with zipfile.ZipFile(zpath) as z:
        for n in z.namelist():
            target = os.path.normpath(os.path.join(dest, n))
            if not target.startswith(os.path.abspath(dest)) and not target.startswith(dest):
                continue  # zip dışına yazma girişimi
            if n.endswith("/"):
                os.makedirs(target, exist_ok=True)
                continue
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with z.open(n) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            if n.lower().endswith(".zip") and depth < 2:
                _extract(target, os.path.splitext(target)[0], depth + 1)


def read_with_inputs(path, base, seen=None):
    seen = seen or set()
    with open(path, encoding="utf-8", errors="replace") as f:
        s = strip_comments(f.read())

    def repl(mt):
        name = mt.group(2).strip()
        cand = os.path.join(base, name if name.endswith(".tex") else name + ".tex")
        if os.path.isfile(cand) and cand not in seen:
            seen.add(cand)
            return read_with_inputs(cand, base, seen)
        return ""
    return re.sub(r"\\(input|include)\{([^{}]+)\}", repl, s)


# --------------------------------------------------------------------------
# Kullanıcı makroları ve revizyon işaretleri
# --------------------------------------------------------------------------
def user_macros(pre):
    r"""\newcommand{\ad}[n]{gövde} ve \def\ad{gövde} tanımları (isteğe bağlı argümansız olanlar)."""
    macros = {}
    for mt in re.finditer(r"\\(?:re)?newcommand\*?\s*\{?\\([A-Za-z@]+)\}?\s*(?:\[(\d)\])?\s*(?=\{)", pre):
        e = group_end(pre, mt.end())
        if e > 0:
            macros[mt.group(1)] = (int(mt.group(2) or 0), pre[mt.end() + 1:e])
    for mt in re.finditer(r"\\def\\([A-Za-z@]+)\s*(?=\{)", pre):
        e = group_end(pre, mt.end())
        if e > 0:
            macros[mt.group(1)] = (0, pre[mt.end() + 1:e])
    return macros


def expand_macros(body, macros, passes=4):
    for _ in range(passes):
        changed = False
        for name, (n, tpl) in macros.items():
            for start, end, _, args in reversed(find_commands(body, name, n) if n else
                                                [(m.start(), m.end(), [], []) for m in
                                                 re.finditer(r"\\%s(?![A-Za-z@])" % re.escape(name), body)]):
                out = tpl
                for k, a in enumerate(args, 1):
                    out = out.replace("#%d" % k, a)
                body = body[:start] + out + body[end:]
                changed = True
        if not changed:
            break
    return body


def strip_revision_marks(body):
    r"""Hakem turlarının renkli işaretleri yayımlanan metinde olmaz: \textcolor{red}{x} → x."""
    for _ in range(3):
        for start, end, _, args in reversed(find_commands(body, "textcolor", 2)):
            body = body[:start] + args[1] + body[end:]
    body = re.sub(r"\\color\{[^{}]*\}", "", body)
    # (\citeauthor{x}, \citeyear{x}) elle yazılmış parantezli atıf → \citep{x}
    body = re.sub(r"\(\s*\\citeauthor\{([^{}]+)\}\s*,\s*\\citeyear\{\1\}\s*\)", r"\\citep{\1}", body)
    return body


# --------------------------------------------------------------------------
# Ön bilgi: başlık, yazarlar, kurumlar
# --------------------------------------------------------------------------
def front_matter(pre):
    """Önsözden (preamble) başlık/yazar/kurum bilgisi ve dergi biçiminde LaTeX bloğu."""
    title = ""
    t = find_commands(pre, "title")
    if t:
        title = clean_inline(t[-1][3][0])
    authors, affils = [], {}
    for _, _, opts, args in find_commands(pre, "affil"):
        affils[(opts[0] if opts else str(len(affils) + 1)).strip()] = clean_inline(args[0])
    for _, _, opts, args in find_commands(pre, "author"):
        marks = [m.strip() for m in (opts[0].split(",") if opts else []) if m.strip()]
        raw = args[0]
        if "\\and" in raw or not opts:
            for part in re.split(r"\\and\b", raw):
                name = clean_inline(re.sub(r"\\(?:inst|textsuperscript)\{([^{}]*)\}", "", part))
                sup = re.findall(r"\\(?:inst|textsuperscript)\{([^{}]*)\}", part)
                pm = [m.strip() for s in sup for m in s.split(",") if m.strip()] or marks
                if name:
                    authors.append((name, pm))
        else:
            authors.append((clean_inline(raw), marks))
    lines = []
    if title:
        lines.append(r"\noindent\textbf{%s}" % title)
    if authors:
        parts = []
        for name, marks in authors:
            sup = ",".join(m.replace("*", r"*") for m in marks)
            parts.append(name + (r"\textsuperscript{%s}" % sup if sup else ""))
        joined = ", ".join(parts[:-1]) + (" and " if len(parts) > 1 else "") + parts[-1]
        lines.append(r"\noindent\textbf{%s}" % joined)
    for key, text in affils.items():
        lines.append(r"\noindent\textsuperscript{%s}%s" % (key, text))
    return title, authors, affils, "\n\n".join(lines)


# --------------------------------------------------------------------------
# Numaralar ve çapraz atıflar
# --------------------------------------------------------------------------
def _env_spans(s, names):
    """İç içe olmayan \\begin{name}…\\end{name} ortamları (sırayla)."""
    pat = re.compile(r"\\begin\{(%s)(\*?)\}" % "|".join(re.escape(n) for n in names))
    out, pos = [], 0
    while True:
        mt = pat.search(s, pos)
        if not mt:
            return out
        name, star = mt.group(1), mt.group(2)
        end_tag = "\\end{%s%s}" % (name, star)
        e = s.find(end_tag, mt.end())
        if e < 0:
            return out
        out.append((mt.start(), e + len(end_tag), name, bool(star), s[mt.end():e]))
        pos = e + len(end_tag)


def _rows(body):
    """align gövdesini en üst düzeydeki \\\\ satırlarına böler."""
    rows, depth, cur, k = [], 0, [], 0
    while k < len(body):
        c = body[k]
        if body.startswith("\\\\", k) and depth == 0:
            rows.append("".join(cur))
            cur = []
            k += 2
            continue
        if c == "\\":
            cur.append(body[k:k + 2])
            k += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        cur.append(c)
        k += 1
    rows.append("".join(cur))
    return [r for r in rows if r.strip()]


def number_equations(body, labels):
    """Denklem ortamlarını numaralı $$…$$ (N) paragraflarına çevirir; etiketleri kaydeder."""
    out, last, n = [], 0, 0
    for start, end, name, star, inner in _env_spans(body, NUMBERED_EQ_ENVS):
        out.append(body[last:start])
        last = end
        rows = _rows(inner) if name in ("align", "gather", "eqnarray", "flalign", "alignat") else [inner]
        if name == "alignat":
            rows = _rows(re.sub(r"^\s*\{[^{}]*\}", "", inner))
        pieces = []
        for row in rows:
            lab = re.findall(r"\\label\{([^{}]+)\}", row)
            tag = re.search(r"\\tag\*?\{([^{}]*)\}", row)
            nonum = star or re.search(r"\\(nonumber|notag)\b", row)
            if name == "multline":
                row = row.replace("\\\\", " ")
            tex = re.sub(r"\\label\{[^{}]*\}|\\tag\*?\{[^{}]*\}|\\(?:nonumber|notag)\b", "", row)
            tex = tex.replace("&", " ").strip() if name in ("align", "eqnarray", "flalign", "alignat") else tex.strip()
            if not tex:
                continue
            if tag:
                num = tag.group(1)
            elif nonum:
                num = None
            else:
                n += 1
                num = str(n)
            for lb in lab:
                labels[lb] = ("Equation", num or "")
            # numara denklemin içinde taşınır: pandoc "(1)" metnini ayrı paragrafa koyuyor
            pieces.append("\n\n$$%s%s$$\n\n" % (tex, r" \tag{%s}" % num if num else ""))
        out.append("".join(pieces))
    out.append(body[last:])
    return "".join(out)


def number_floats(body, labels):
    """figure/table ortamlarını "Figure N: …" başlıklı düz bloklara çevirir; etiketleri kaydeder."""
    out, last = [], 0
    count = {"figure": 0, "table": 0}
    for start, end, name, star, inner in _env_spans(body, ("figure", "table", "SCfigure", "wrapfigure")):
        out.append(body[last:start])
        last = end
        kind = "table" if name == "table" else "figure"
        inner = re.sub(r"^\s*\[[^\]]*\]", "", inner)  # [h], [htbp]
        if name == "wrapfigure":
            inner = re.sub(r"^\s*\{[^{}]*\}\s*\{[^{}]*\}", "", inner)
        caps = find_commands(inner, "caption")
        caption = caps[0][3][0].strip() if caps else ""
        for lb in re.findall(r"\\label\{([^{}]+)\}", inner):
            labels[lb] = ("Figure" if kind == "figure" else "Table", str(count[kind] + 1) if caps else "")
        content = remove_commands(inner, "caption")
        content = re.sub(r"\\label\{[^{}]*\}|\\centering\b|\\(?:small|footnotesize|scriptsize)\b", "", content)
        if caps:
            count[kind] += 1
            label = "Figure" if kind == "figure" else "Table"
            cap = "\n\n\\noindent %s %d: %s\n\n" % (label, count[kind], caption)
        else:
            cap = ""
        # tablo başlığı tablonun üstünde, şekil başlığı resmin altında
        out.append(cap + content + "\n\n" if kind == "table" else "\n\n" + content + cap)
    out.append(body[last:])
    return "".join(out)


def number_sections(body, labels):
    """Numaralı bölümlerin etiketleri ("Section 2.1")."""
    sec = [0, 0, 0]
    for mt in re.finditer(r"\\(section|subsection|subsubsection)(\*?)\s*(?:\[[^\]]*\])?\s*\{", body):
        if mt.group(2):
            continue
        lvl = ("section", "subsection", "subsubsection").index(mt.group(1))
        sec[lvl] += 1
        for k in range(lvl + 1, 3):
            sec[k] = 0
        num = ".".join(str(x) for x in sec[:lvl + 1])
        e = group_end(body, mt.end() - 1)
        tail = body[e + 1:e + 120] if e > 0 else ""
        lab = re.match(r"\s*\\label\{([^{}]+)\}", tail)
        if lab:
            labels[lab.group(1)] = ("Section", num)


def resolve_refs(body, labels):
    def one(key):
        return labels.get(key.strip())

    def ref(mt):
        cmd, keys = mt.group(1), mt.group(2).split(",")
        res = []
        for key in keys:
            v = one(key)
            if v is None:
                warn("LaTeX: \\%s{%s} için etiket bulunamadı." % (cmd, key.strip()))
                res.append("??")
                continue
            kind, num = v
            if cmd == "eqref":
                # çıplak "(2)" metinden tanınamaz: doğrudan iç bağlantı (render.py "@@" → \hyperlink)
                res.append(r"\href{@@eq:%s}{(%s)}" % (num, num) if num.isdigit() else "(%s)" % num)
            elif cmd in ("autoref", "cref", "Cref"):
                res.append("%s %s" % (kind, num))
            else:
                res.append(num)
        return ", ".join(res)
    body = re.sub(r"\\(eqref|ref|autoref|cref|Cref|pageref)\*?\{([^{}]+)\}", ref, body)
    return re.sub(r"\\label\{[^{}]*\}", "", body)


# --------------------------------------------------------------------------
# Kaynakça
# --------------------------------------------------------------------------
def bib_files(body, base):
    out = []
    for mt in re.finditer(r"\\(?:bibliography|addbibresource)\{([^{}]+)\}", body):
        for name in mt.group(1).split(","):
            p = os.path.join(base, name.strip())
            p = p if p.endswith(".bib") else p + ".bib"
            if os.path.isfile(p):
                out.append(p)
    return out


def bbl_items(text):
    """thebibliography / .bbl: [(anahtar, natbib etiketi, metin)]."""
    items = []
    for mt in re.finditer(r"\\bibitem(?:\[((?:[^\[\]]|\[[^\]]*\])*)\])?\{([^{}]+)\}(.*?)(?=\\bibitem|\\end\{thebibliography\}|$)",
                          text, re.S):
        body = re.sub(r"\\newblock\b", " ", mt.group(3))
        items.append((mt.group(2).strip(), mt.group(1) or "", re.sub(r"\s+", " ", body).strip()))
    return items


def natbib_cites(body, items):
    """.bib yokken natbib atıflarını etiketlerden yazar: \\citep → (Smith et al., 2020)."""
    info = {}
    for key, label, _ in items:
        mt = re.match(r"(.*?)\((\d{4}[a-z]?|n\.d\.)\)(.*)", re.sub(r"[{}~]", " ", label))
        if mt:
            info[key] = (re.sub(r"\s+", " ", mt.group(1)).strip(), mt.group(2))
    order = {key: k + 1 for k, (key, _, _) in enumerate(items)}

    def cite(mt):
        cmd, keys = mt.group(1), [k.strip() for k in mt.group(3).split(",")]
        if not all(k in info for k in keys):
            nums = [str(order[k]) for k in keys if k in order]
            return "[%s]" % ", ".join(nums) if nums else ""
        if cmd in ("citet", "cite", "Citet"):
            return "; ".join("%s (%s)" % info[k] for k in keys)
        if cmd == "citeauthor":
            return "; ".join(info[k][0] for k in keys)
        if cmd == "citeyear":
            return ", ".join(info[k][1] for k in keys)
        return "(%s)" % "; ".join("%s, %s" % info[k] for k in keys)
    return re.sub(r"\\(citep|citet|cite|Citet|Citep|citeauthor|citeyear|citealp)\*?(\[[^\]]*\])?\{([^{}]+)\}", cite, body)


# --------------------------------------------------------------------------
# Ana dönüşüm
# --------------------------------------------------------------------------
def pandoc_exe():
    """Önce pypandoc_binary'nin getirdiği pandoc (sürüm her yerde aynı), yoksa sistemdeki."""
    try:
        import pypandoc
        exe = pypandoc.get_pandoc_path()
        if exe and os.path.isfile(exe):
            return exe
    except Exception:  # noqa
        pass
    exe = shutil.which("pandoc")
    if exe:
        return exe
    raise SystemExit("HATA: LaTeX okumak için pandoc gerekli:  pip install pypandoc_binary")


def prepare_tex(src, base):
    """Ön işlenmiş LaTeX metni, .bib listesi ve bilgi sözlüğü."""
    if "\\begin{document}" not in src:
        raise SystemExit("HATA: .tex dosyasında \\begin{document} yok.")
    pre, body = src.split("\\begin{document}", 1)
    body = body.split("\\end{document}", 1)[0]
    body = strip_revision_marks(expand_macros(body, user_macros(pre)))
    title, authors, affils, front = front_matter(pre)
    # kullanıcı makroları (\newcommand) pandoc'a kalsın; başlık/yazar komutları çıkarılsın
    for name, n in (("title", 1), ("author", 1), ("affil", 1), ("date", 1), ("address", 1), ("institute", 1)):
        pre = remove_commands(pre, name, n)
    pre = re.sub(r"\\usepackage(\[[^\]]*\])?\{(authblk|titlesec|ragged2e|natbib|setspace)\}", "", pre)
    labels = {}
    number_sections(body, labels)
    body = number_floats(body, labels)
    body = number_equations(body, labels)
    body = resolve_refs(body, labels)
    if "\\maketitle" in body:
        body = body.replace("\\maketitle", "\n\n" + front + "\n\n", 1)
    else:
        body = front + "\n\n" + body
    body = re.sub(r"\\(justifying|raggedright|centering|selectfont|onecolumn|twocolumn)\b", "", body)
    body = re.sub(r"\\fontsize\{[^{}]*\}\{[^{}]*\}", "", body)
    body = re.sub(r"\\vspace\*?\{[^{}]*\}", "", body)
    bibs = bib_files(body, base)
    bbl = None
    if not bibs:
        tb = re.search(r"\\begin\{thebibliography\}.*?\\end\{thebibliography\}", body, re.S)
        if tb:
            bbl = tb.group(0)
            body = body.replace(bbl, "\\bibliography{x}")
        else:
            cand = [f for f in os.listdir(base) if f.lower().endswith(".bbl")]
            if cand:
                with open(os.path.join(base, cand[0]), encoding="utf-8", errors="replace") as f:
                    bbl = f.read()
    refs_tex = ""
    if bbl:
        items = bbl_items(bbl)
        body = natbib_cites(body, items)
        refs_tex = "\n\n".join(re.sub(r"\\url\{([^{}]*)\}", r"\1", it[2]) for it in items)
    body = re.sub(r"\\bibliographystyle\{[^{}]*\}", "", body)
    head = "\n\n\\section*{References}\n\n" + refs_tex + "\n\n"
    if re.search(r"\\(bibliography|printbibliography)\b", body):
        body = re.sub(r"\\(?:bibliography\{[^{}]*\}|printbibliography(\[[^\]]*\])?)", lambda m: head, body, count=1)
    elif bibs or bbl:
        body += head
    info = {"title": title, "authors": authors, "affiliations": affils, "bib": bibs,
            "labels": labels, "bbl": bool(bbl)}
    return pre + "\\begin{document}\n" + body + "\n\\end{document}\n", bibs, info


def _walk_math(node, store):
    """AST'deki her Math düğümünün yerine işaret koyar, özgün TeX'i saklar."""
    if isinstance(node, dict):
        if node.get("t") == "Math":
            kind, tex = node["c"]
            tag = re.search(r"\\tag\*?\{([^{}]*)\}", tex)
            store.append((re.sub(r"\\tag\*?\{[^{}]*\}", "", tex).strip(), tag.group(1) if tag else None))
            node["c"] = [kind, r"\mathrm{%s}" % math_mark(len(store) - 1)]
            return
        for v in node.values():
            _walk_math(v, store)
    elif isinstance(node, list):
        for v in node:
            _walk_math(v, store)


def figures_to_png(tex, base, workdir):
    """\\includegraphics{x.pdf|eps} → PNG (PyMuPDF varsa); Word PDF resim gömemez."""
    def repl(mt):
        name = mt.group(2)
        path = None
        for ext in ("", ".pdf", ".eps", ".png", ".jpg", ".jpeg"):
            p = os.path.join(base, name + ext)
            if os.path.isfile(p):
                path = p
                break
        if not path or not path.lower().endswith((".pdf", ".eps")):
            return mt.group(0)
        try:
            try:
                import pymupdf
            except ImportError:
                import fitz as pymupdf
            doc = pymupdf.open(path)
            out = os.path.join(workdir, re.sub(r"[^\w.-]", "_", name) + ".png")
            doc[0].get_pixmap(dpi=200).save(out)
            return "\\includegraphics%s{%s}" % (mt.group(1) or "", out.replace("\\", "/"))
        except Exception:  # noqa
            warn("LaTeX: %s resmi PNG'ye çevrilemedi (pip install pymupdf)." % os.path.basename(path))
            return mt.group(0)
    return re.sub(r"\\includegraphics(\[[^\]]*\])?\{([^{}]+)\}", repl, tex)


def tex_to_docx(path, workdir=None):
    """LaTeX kaynağını .docx'e çevirir; yolunu döndürür (denklemler TEX_MATH'te)."""
    workdir = workdir or tempfile.mkdtemp(prefix="sjuoz_tex_")
    main, base = locate_main_tex(path, workdir)
    print("Bilgi: LaTeX kaynağı okunuyor: %s" % os.path.basename(main), flush=True)
    src = read_with_inputs(main, base)
    tex, bibs, info = prepare_tex(src, base)
    tex = figures_to_png(tex, base, workdir)
    prepped = os.path.join(workdir, "_sjuoz_hazir.tex")  # yazarın klasörüne dosya yazılmaz
    with open(prepped, "w", encoding="utf-8") as f:
        f.write(tex)
    exe = pandoc_exe()
    common = ["--resource-path", base + os.pathsep + workdir]
    cmd = [exe, prepped, "-f", "latex", "-t", "json"] + common
    if bibs:
        cmd += ["--citeproc", "--csl", os.path.join(DATA_DIR, "apa.csl")]
        for b in bibs:
            cmd += ["--bibliography", b]
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=base,
                         timeout=300, **NO_WINDOW)
    if res.returncode != 0:
        raise SystemExit("HATA: pandoc LaTeX'i okuyamadı:\n" + res.stderr[-2000:])
    for ln in res.stderr.splitlines():
        if "WARNING" in ln and ("Could not convert TeX math" in ln or "Citeproc" in ln or "not found" in ln):
            warn("pandoc: " + ln.split("]", 1)[-1].strip()[:200])
    ast = json.loads(res.stdout)
    store = []
    _walk_math(ast, store)
    out = os.path.join(workdir, os.path.splitext(os.path.basename(main))[0] + ".docx")
    res2 = subprocess.run([exe, "-f", "json", "-t", "docx", "-o", out] + common, input=json.dumps(ast),
                          capture_output=True, text=True, encoding="utf-8", errors="replace", cwd=base,
                          timeout=300, **NO_WINDOW)
    if res2.returncode != 0:
        raise SystemExit("HATA: pandoc .docx yazamadı:\n" + res2.stderr[-2000:])
    TEX_MATH[os.path.abspath(out)] = dict(enumerate(store))
    print("Bilgi: LaTeX'ten %d denklem, %d kaynak dosyası okundu." % (len(store), len(bibs)))
    return out


def restore_math(doc_path, latex, num=None):
    """Docx okunurken: işaretli denklem yerine yazarın özgün LaTeX'i ve numarası: (latex, num)."""
    table = TEX_MATH.get(os.path.abspath(doc_path)) if doc_path else None
    if not table:
        return latex, num
    mt = MATH_MARK_RE.search(re.sub(r"\\[A-Za-z]+|[^a-z]", "", latex))
    if not mt or _mark_number(mt.group(1)) not in table:
        return latex, num
    tex, tag = table[_mark_number(mt.group(1))]
    return tex, tag or num
