@echo off
chcp 65001 >nul
set PYTHONDONTWRITEBYTECODE=1
cd /d "%~dp0"

set PY=python
%PY% --version >nul 2>&1
if errorlevel 1 (
  set PY=py -3
  py -3 --version >nul 2>&1
  if errorlevel 1 (
    echo Chua co Python tren may nay. Tai tai python.org, nho tick Add Python to PATH, roi chay lai file nay.
    pause
    exit /b 1
  )
)

echo Dang cai thu vien...
%PY% -m pip install -q -r requirements.txt
if errorlevel 1 (
  echo Cai bang quyen may khong duoc, thu cai rieng cho tai khoan nay...
  %PY% -m pip install -q --user -r requirements.txt
)

%PY% tools\setup_wizard.py
pause
