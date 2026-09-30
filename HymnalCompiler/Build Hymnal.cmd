@echo off
setlocal
cd /d "%~dp0"
set "PATH=%~dp0tools\typst;%PATH%"
set "TYPST_FONT_PATHS=%~dp0fonts"
set "PYTHONUTF8=1"

if "%~1"=="" (
  set /p SRC="Workbook path or Google Sheets URL: "
) else (
  set "SRC=%~1"
)
set SRC=%SRC:"=%

"%~dp0tools\python\python.exe" src\build.py "%SRC%"
if errorlevel 1 (
  echo.
  echo BUILD FAILED - see messages above.
  pause
  exit /b 1
)
start "" "src\hymnal.pdf"
pause