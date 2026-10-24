# -*- coding: utf-8 -*-
"""Girdi açma: zip (Word / PDF / LaTeX), .doc çevirisi ve dosya arama."""
import os
import shutil
import zipfile

import pytest

from conftest import CASES, FIXTURES
from sjuoz import inputs
from sjuoz.cli import resolve_docx

C1993 = next(c for c in CASES if c["name"] == "1993")
MAN, COV = sorted(C1993["docx"], key=lambda p: "cover" in p)


def make_zip(path, files):
    with zipfile.ZipFile(path, "w") as z:
        for arc, src in files.items():
            z.write(src, arc)
    return str(path)


def test_zip_with_word_files_prefers_word_over_same_pdf(tmp_path):
    pdf = os.path.join(FIXTURES, "sentetik_pdf", "sentetik_pdf.pdf")
    zp = make_zip(tmp_path / "g.zip", {"a/1993_manuscript.docx": MAN, "a/1993_cover_letter.docx": COV,
                                       "a/1993_manuscript.pdf": pdf, "__MACOSX/a/._x.docx": MAN})
    assert not inputs.zip_has_tex(zp)
    out = inputs.expand_inputs([zp], str(tmp_path / "w"))
    assert sorted(os.path.basename(p) for p in out) == ["1993_cover_letter.docx", "1993_manuscript.docx"]


def test_zip_with_only_pdf_and_nested_zip(tmp_path):
    pdf = os.path.join(FIXTURES, "sentetik_pdf", "sentetik_pdf.pdf")
    inner = make_zip(tmp_path / "inner.zip", {"makale.pdf": pdf})
    outer = make_zip(tmp_path / "outer.zip", {"inner.zip": inner, "oku.txt": pdf})
    assert [os.path.basename(p) for p in inputs.expand_inputs([outer], str(tmp_path / "w"))] == ["makale.pdf"]
    tex = os.path.join(FIXTURES, "sentetik_latex", "main.tex")
    texzip = make_zip(tmp_path / "t.zip", {"src.zip": make_zip(tmp_path / "src.zip", {"main.tex": tex})})
    assert inputs.zip_has_tex(texzip) and inputs.expand_inputs([texzip]) == [texzip]  # LaTeX yoluna kalır


def test_zip_slip_is_blocked(tmp_path):
    zp = tmp_path / "kotu.zip"
    with zipfile.ZipFile(zp, "w") as z:
        z.writestr("../../kacak.docx", "x")
        z.writestr("ok.docx", "x")
    files = inputs.extract_zip(str(zp), str(tmp_path / "d"))
    assert [os.path.basename(f) for f in files] == ["ok.docx"]
    assert not (tmp_path / "kacak.docx").exists()


def test_doc_without_libreoffice_gives_clear_error(tmp_path, monkeypatch):
    monkeypatch.setattr(inputs, "soffice_path", lambda: None)
    doc = tmp_path / "eski.doc"
    doc.write_bytes(b"x")
    with pytest.raises(SystemExit) as e:
        inputs.expand_inputs([str(doc)], str(tmp_path))
    assert "Farklı Kaydet" in str(e.value)


@pytest.mark.skipif(inputs.soffice_path() is None, reason="LibreOffice yok")
def test_doc_is_converted(tmp_path):
    import subprocess
    shutil.copy(COV, tmp_path / "kapak.docx")
    subprocess.run([inputs.soffice_path(), "--headless", "--convert-to", "doc", "--outdir", str(tmp_path),
                    str(tmp_path / "kapak.docx")], capture_output=True, timeout=180)
    if not (tmp_path / "kapak.doc").exists():
        pytest.skip("LibreOffice Writer yok (yalnız Draw kurulu): .doc örneği üretilemedi")
    out = inputs.expand_inputs([str(tmp_path / "kapak.doc")], str(tmp_path / "w"))
    assert out[0].endswith("kapak.docx") and os.path.getsize(out[0]) > 0


def test_resolve_searches_subfolders_and_zips(tmp_path, monkeypatch):
    home = tmp_path / "home"
    work = home / "OneDrive" / "Desktop" / "kod"
    (work / "gelen" / "jordan").mkdir(parents=True)
    shutil.copy(MAN, work / "gelen" / "jordan" / "jORDAN.docx")
    (home / "Downloads").mkdir()
    make_zip(home / "Downloads" / "gonderim.zip", {"x/SJUOZ+Cover+Letter-JORDAN (2).docx": COV})
    for k, v in (("HOME", home), ("USERPROFILE", home)):
        monkeypatch.setenv(k, str(v))
    monkeypatch.delenv("OneDrive", raising=False)
    monkeypatch.chdir(work)
    real, err = resolve_docx("jORDAN.docx")
    assert err is None and real.endswith(os.path.join("gelen", "jordan", "jORDAN.docx"))
    real, err = resolve_docx("SJUOZ+Cover+Letter-JORDAN (2).docx")
    assert err is None and os.path.basename(real) == "SJUOZ+Cover+Letter-JORDAN (2).docx" and os.path.isfile(real)
    real, err = resolve_docx("yok.pdf")
    assert real is None and "zip" in err


def test_resolve_prefers_word_and_latex_over_pdf(tmp_path, monkeypatch):
    for n in ("makale.pdf", "makale.docx", "rapor.pdf", "rapor.tex"):
        (tmp_path / n).write_bytes(b"x")
    monkeypatch.chdir(tmp_path)
    assert resolve_docx("makale")[0] == "makale.docx"
    assert resolve_docx("rapor.docx")[0] == "rapor.tex"  # PDF'ten önce LaTeX kaynağı


def test_report_flags_glued_words():
    from sjuoz.report import glued_words

    class P:
        def __init__(self, t):
            self._t = t

        def text(self):
            return self._t

    class A:
        pass
    import sjuoz.docx as dx
    art = A()
    art.body = [dx.Para.__new__(dx.Para) for _ in range(2)]
    art.body[0].tokens = [("t", "InPapuaNewGuinea,agricultureisoneofthemost important " * 3, None)]
    art.body[1].tokens = [("t", "Agricultureisthemainsourceofliving for many; Thecoefficientpiisdefinedstrictly x", None)]
    assert len(glued_words(art)) >= 5
    art.body[0].tokens = [("t", "A normal sentence with optimisation-based https://doi.org/10.1000/abcdefghijklmnop", None)]
    art.body[1].tokens = [("t", "Another normal sentence.", None)]
    assert glued_words(art) == []
