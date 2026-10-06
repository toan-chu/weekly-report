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
    echo Chua co Python tren may nay. Chay 1_Khoi_tao_workspace.bat truoc.
    pause
    exit /b 1
  )
)

echo Dang kiem tra file ghep khach... (Excel se tu mo file ket qua)
%PY% runner.py --kiem-tra-ghep-khach
if errorlevel 1 (
  echo.
  echo Co loi o tren. Doc dong cuoi cung de biet can lam gi.
  pause
  exit /b 1
)
timeout /t 5 >nul
