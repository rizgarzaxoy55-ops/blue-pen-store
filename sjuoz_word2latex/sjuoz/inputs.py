# -*- coding: utf-8 -*-
"""Girdi dosyalarını dönüştürülebilir dosyalara açar.

  * .docx                      → olduğu gibi
  * .doc / .odt / .rtf         → LibreOffice ile .docx'e çevrilir
  * .zip                       → içinde .tex varsa LaTeX kaynağı olarak (texinput); yoksa içindeki Word
                                 dosyaları, onlar da yoksa PDF'ler (iç içe zip'ler de açılır)
  * .pdf / .tex                → olduğu gibi (pipeline'da pdfinput / texinput çevirir)
"""
import io
import os
import shutil
import subprocess
import tempfile
import zipfile

from .common import NO_WINDOW, warn

WORD_EXTS = (".docx", ".doc", ".odt", ".rtf")
INPUT_EXTS = WORD_EXTS + (".pdf", ".tex", ".zip")


def _skip(name):
    base = os.path.basename(name)
    return base.startswith(("~$", "._")) or "__MACOSX" in name or name.endswith("/")


def zip_names(source, depth=0):
    """Zip'teki (iç içe zip'ler dahil) dosya adları."""
    out = []
    try:
        with zipfile.ZipFile(source) as z:
            for n in z.namelist():
                if _skip(n):
                    continue
                out.append(n)
                if n.lower().endswith(".zip") and depth < 2:
                    try:
                        out += [n + "/" + m for m in zip_names(io.BytesIO(z.read(n)), depth + 1)]
                    except zipfile.BadZipFile:
                        pass
    except zipfile.BadZipFile:
        pass
    return out


def zip_has_tex(path):
    return any(n.lower().endswith(".tex") for n in zip_names(path))


def extract_zip(path, dest, depth=0):
    """Zip'i (iç içe olanlarla) güvenli biçimde açar; açılan dosyaların yolları."""
    files = []
    root = os.path.abspath(dest)
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if _skip(n):
                continue
            target = os.path.abspath(os.path.join(root, n))
            if not target.startswith(root + os.sep):
                continue  # zip dışına yazma girişimi
            os.makedirs(os.path.dirname(target), exist_ok=True)
            with z.open(n) as src, open(target, "wb") as out:
                shutil.copyfileobj(src, out)
            if n.lower().endswith(".zip") and depth < 2:
                try:
                    files += extract_zip(target, os.path.splitext(target)[0], depth + 1)
                    continue
                except zipfile.BadZipFile:
                    pass
            files.append(target)
    return files


def soffice_path():
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if exe:
        return exe
    for v in ("ProgramFiles", "ProgramFiles(x86)"):
        p = os.path.join(os.environ.get(v, ""), "LibreOffice", "program", "soffice.exe")
        if os.path.isfile(p):
            return p
    return None


def to_docx(path, workdir):
    """.doc/.odt/.rtf → .docx (LibreOffice)."""
    exe = soffice_path()
    if not exe:
        raise SystemExit("HATA: %s eski/başka bir Word biçiminde. LibreOffice kurun ya da Word'de "
                         "Farklı Kaydet → Word Belgesi (.docx) yapın." % os.path.basename(path))
    print("Bilgi: %s LibreOffice ile .docx'e çevriliyor..." % os.path.basename(path), flush=True)
    subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", workdir, path],
                   capture_output=True, timeout=300, **NO_WINDOW)
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".docx")
    if not os.path.isfile(out):
        raise SystemExit("HATA: %s .docx'e çevrilemedi; Word'de Farklı Kaydet → .docx yapın."
                         % os.path.basename(path))
    return out


def expand_inputs(paths, workdir=None):
    """Kullanıcının verdiği dosyalar → dönüştürücüye verilecek dosyalar (sırayla)."""
    workdir = workdir or tempfile.mkdtemp(prefix="sjuoz_girdi_")
    out = []
    for p in paths:
        low = p.lower()
        if low.endswith(".zip") and not zip_has_tex(p):
            dest = os.path.join(workdir, os.path.splitext(os.path.basename(p))[0])
            files = extract_zip(p, dest)
            words = sorted(f for f in files if f.lower().endswith(WORD_EXTS))
            pdfs = sorted(f for f in files if f.lower().endswith(".pdf"))
            # aynı adlı Word ve PDF varsa Word kullanılır
            stems = {os.path.splitext(os.path.basename(f))[0].lower() for f in words}
            pdfs = [f for f in pdfs if os.path.splitext(os.path.basename(f))[0].lower() not in stems]
            chosen = words or pdfs
            if not chosen:
                raise SystemExit("HATA: %s içinde .docx, .doc, .tex ya da .pdf yok." % os.path.basename(p))
            if words and pdfs:
                print("Bilgi: zip'teki PDF'ler (%s) Word dosyaları olduğu için kullanılmadı."
                      % ", ".join(os.path.basename(f) for f in pdfs))
                chosen = words
            print("Bilgi: %s içinden kullanılan: %s" % (os.path.basename(p),
                                                        ", ".join(os.path.basename(f) for f in chosen)))
            out += [to_docx(f, workdir) if f.lower().endswith((".doc", ".odt", ".rtf")) else f for f in chosen]
        elif low.endswith((".doc", ".odt", ".rtf")):
            out.append(to_docx(p, workdir))
        else:
            out.append(p)
    if len(out) > 2:
        warn("%d dosya verildi; makale ve kapak mektubu dışındakiler yok sayılabilir." % len(out))
    return out


def _nearby_dirs(folder):
    """Dosyaların sık bulunduğu yerler: üst klasör, Masaüstü, İndirilenler, Belgeler (OneDrive dahil)."""
    home = os.path.expanduser("~")
    dirs = [os.path.dirname(os.path.abspath(folder))]
    for base in (home, os.path.join(home, "OneDrive"), os.environ.get("OneDrive", "")):
        if base:
            dirs += [os.path.join(base, d) for d in ("Desktop", "Masaüstü", "Downloads", "Documents", "Belgeler")]
    seen, out = set(), []
    for d in dirs:
        key = os.path.normcase(os.path.abspath(d))
        if key not in seen and key != os.path.normcase(os.path.abspath(folder)) and os.path.isdir(d):
            seen.add(key)
            out.append(d)
    return out


def _walk(folder, max_depth=3, limit=20000):
    """Klasör ve alt klasörlerindeki girdi dosyaları (gizli/sistem klasörleri atlanır)."""
    count = 0
    root_depth = os.path.abspath(folder).count(os.sep)
    for d, subdirs, files in os.walk(folder):
        if os.path.abspath(d).count(os.sep) - root_depth >= max_depth:
            subdirs[:] = []
        subdirs[:] = [s for s in subdirs if not s.startswith((".", "$", "_cikti", "node_modules", "AppData"))]
        for f in files:
            count += 1
            if f.lower().endswith(INPUT_EXTS) and not f.startswith("~$"):
                yield os.path.join(d, f)
        if count > limit:
            return
