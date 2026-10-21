# Windows'ta .exe üretir (GitHub Actions bunu kendisi çalıştırır; elle de çalıştırılabilir).
#   pip install pyinstaller lxml Pillow matplotlib pypdf pdf2docx "pymupdf<2" pypandoc_binary
#   powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)
# pandoc.exe (pypandoc_binary) PyInstaller'ın veri toplamasına girmiyor: ikili dosya olarak açıkça eklenir
$pandocDir = python -c "import os, pypandoc; print(os.path.join(os.path.dirname(pypandoc.__file__), 'files'))"
if (-not (Test-Path "$pandocDir\pandoc.exe")) { throw "pandoc.exe bulunamadı: pip install pypandoc_binary" }
$common = @("--noconfirm", "--onefile", "--add-data", "sjuoz/data;sjuoz/data", "--paths", ".",
            "--hidden-import", "pdf2docx", "--collect-submodules", "pdf2docx", "--collect-data", "pymupdf",
            "--collect-all", "pypandoc", "--add-binary", "$pandocDir\pandoc.exe;pypandoc/files")
pyinstaller @common --windowed --name "SJUOZ-Word2LaTeX" packaging/gui_entry.py
pyinstaller @common --console --name "sjuoz-word2latex" packaging/cli_entry.py
Get-ChildItem dist
