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

echo Dang chay bao cao tuan...
%PY% runner.py --mo-ket-qua
echo.
echo Xong. Thu muc 02_Output da duoc mo. Neu khong thay bao cao moi, doc cac dong o tren.
pause
