# Windows'ta .exe üretir (GitHub Actions bunu kendisi çalıştırır; elle de çalıştırılabilir).
#   pip install pyinstaller lxml Pillow matplotlib pypdf pdf2docx
#   powershell -ExecutionPolicy Bypass -File packaging\build_windows.ps1
$ErrorActionPreference = "Stop"
Set-Location (Split-Path $PSScriptRoot -Parent)
$common = @("--noconfirm", "--onefile", "--add-data", "sjuoz/data;sjuoz/data", "--paths", ".",
            "--hidden-import", "pdf2docx", "--collect-submodules", "pdf2docx")
pyinstaller @common --windowed --name "SJUOZ-Word2LaTeX" packaging/gui_entry.py
pyinstaller @common --console --name "sjuoz-word2latex" packaging/cli_entry.py
Get-ChildItem dist
