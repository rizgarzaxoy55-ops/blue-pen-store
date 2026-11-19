@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv" (
  echo Ilk calistirma: paketler kuruluyor...
  py -m venv .venv
  call .venv\Scripts\activate
  pip install -r requirements.txt
) else (
  call .venv\Scripts\activate
)
python odul.py %*
echo.
echo Bitti. Sonuclar "cikti" klasorunde.
start "" "%~dp0cikti"
pause
