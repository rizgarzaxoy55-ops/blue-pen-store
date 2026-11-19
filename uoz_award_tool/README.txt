UoZ PUBLICATION AWARD TOOL
==========================
Turns a staff member's award request letter into the Kurdish award letter (خەلاتکرن)
and verifies the paper (published, DOI, authors, UoZ registry, staff portal).
The full user guide is a separate document (USER GUIDE).

SETUP (once)
  1. Install Python 3.10+ (python.org), tick "Add python.exe to PATH".
  2. In cmd, inside this folder:   pip install -r requirements.txt
  3. For scanned PDFs/images only: install Tesseract OCR
     https://github.com/UB-Mannheim/tesseract/wiki   (tick "Arabic" during setup)
  4. Optional: copy settings.example.json to settings.json and add keys
     (anthropic_api_key = paid AI mode, openalex_api_key = free, avoids daily limits).

USE
  Put request letters (Word / PDF / image) into the "input" folder, then double-click run.bat
  or run:   python award.py
  One file:              python award.py "C:\...\letter.docx"
  Rebuild after editing: python award.py --json "output\(authors)\data.json"
  Offline / no checks:   python award.py --no-web "..."
  All options:           python award.py --help

RESULTS (output folder)
  output\summary.html                  all papers, award, clickable article/DOI/staff portal links
  output\(authors)\... - خەلاتکرن.docx  the award letter
  output\(authors)\report.txt          links, verification, WARNINGS (read before signing)
  output\(authors)\data.json           edit to correct names/IF, then rebuild with --json

AWARD RULES (Ministerial order 1692, 23/01/2025)
  Clarivate IF: 0.1-1 = 500,000; +20,000 per 0.1; 6+ = 1,500,000      clause (1)
  Scopus (no IF): Q1 500,000; Q2 400,000; Q3/Q4 300,000; none 250,000   clause (2)
  University journals: approved KRG 250,000; Iraqi/foreign 200,000      clause (3)

ALWAYS CHECK BY HAND: the IF in Clarivate, every ✗ in summary.html, the WARNINGS,
and the Kurdish spelling of names in the Word letter.
