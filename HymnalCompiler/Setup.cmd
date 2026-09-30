@echo off
setlocal
cd /d "%~dp0"

rem ============================================================
rem  Setup.cmd - one-time setup for the Hymnal Compiler.
rem  Downloads Typst and portable Python into tools\ and installs
rem  openpyxl. Safe to re-run: finished steps are skipped.
rem ============================================================

rem --- Settings you may want to change ---------------------------------
rem Leave TYPST_VERSION blank for the newest release, or pin one, e.g. 0.13.1
rem (no leading "v"). Pinning is recommended once your pagination is tuned.
set "TYPST_VERSION=0.15.1"
set "PY_VER=3.14.7"
rem ---------------------------------------------------------------------

set "PY_ZIP=python-%PY_VER%-embed-amd64.zip"
set "PY_URL=https://www.python.org/ftp/python/%PY_VER%/%PY_ZIP%"
set "PIP_URL=https://bootstrap.pypa.io/get-pip.py"
if "%TYPST_VERSION%"=="" (
  set "TYPST_URL=https://github.com/typst/typst/releases/latest/download/typst-x86_64-pc-windows-msvc.zip"
) else (
  set "TYPST_URL=https://github.com/typst/typst/releases/download/v%TYPST_VERSION%/typst-x86_64-pc-windows-msvc.zip"
)
set "PS=powershell -NoProfile -ExecutionPolicy Bypass -Command"
set "DL=%~dp0tools\_download"
set "PYEXE=%~dp0tools\python\python.exe"

echo.
echo === Hymnal Compiler setup ===
echo.

if not exist "%~dp0tools" mkdir "%~dp0tools"
if not exist "%DL%" mkdir "%DL%"

rem ---------- 1. Typst ----------
if exist "%~dp0tools\typst\typst.exe" (
  echo [1/4] Typst already installed - skipping.
  goto :python
)
echo [1/4] Downloading Typst ...
%PS% "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%TYPST_URL%' -OutFile '%DL%\typst.zip'"
if errorlevel 1 goto :fail
%PS% "$ErrorActionPreference='Stop'; Expand-Archive -Path '%DL%\typst.zip' -DestinationPath '%DL%\typst_x' -Force; New-Item -ItemType Directory -Force -Path '%~dp0tools\typst' | Out-Null; Get-ChildItem -Path '%DL%\typst_x' -Recurse -Filter typst.exe | Select-Object -First 1 | Copy-Item -Destination '%~dp0tools\typst\typst.exe' -Force"
if errorlevel 1 goto :fail
if not exist "%~dp0tools\typst\typst.exe" goto :fail

:python
rem ---------- 2. Portable Python ----------
if exist "%PYEXE%" (
  echo [2/4] Python already installed - skipping.
  goto :pip
)
echo [2/4] Downloading Python %PY_VER% ...
%PS% "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PY_URL%' -OutFile '%DL%\%PY_ZIP%'"
if errorlevel 1 goto :fail
%PS% "$ErrorActionPreference='Stop'; Expand-Archive -Path '%DL%\%PY_ZIP%' -DestinationPath '%~dp0tools\python' -Force"
if errorlevel 1 goto :fail
if not exist "%PYEXE%" goto :fail

rem Rewrite the ._pth file: enable site-packages and add the src folder.
for %%F in ("%~dp0tools\python\python*._pth") do (
  > "%%F" (
    echo %%~nF.zip
    echo .
    echo ..\..\src
    echo import site
  )
)

:pip
rem ---------- 3. pip ----------
"%PYEXE%" -m pip --version >nul 2>&1
if not errorlevel 1 (
  echo [3/4] pip already installed - skipping.
  goto :openpyxl
)
echo [3/4] Installing pip ...
%PS% "$ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%PIP_URL%' -OutFile '%DL%\get-pip.py'"
if errorlevel 1 goto :fail
"%PYEXE%" "%DL%\get-pip.py" --no-warn-script-location
if errorlevel 1 goto :fail

:openpyxl
rem ---------- 4. openpyxl ----------
"%PYEXE%" -c "import openpyxl" >nul 2>&1
if not errorlevel 1 (
  echo [4/4] openpyxl already installed - skipping.
  goto :verify
)
echo [4/4] Installing openpyxl ...
"%PYEXE%" -m pip install openpyxl --no-warn-script-location
if errorlevel 1 goto :fail

:verify
if exist "%DL%" rmdir /s /q "%DL%"

echo.
echo === Verification ===
"%~dp0tools\typst\typst.exe" --version
"%PYEXE%" --version
"%PYEXE%" -c "import openpyxl; print('openpyxl', openpyxl.__version__)"

dir /b "%~dp0fonts\*.ttf" >nul 2>&1
if errorlevel 1 (
  echo.
  echo WARNING: no .ttf files found in the fonts folder.
  echo Add the PT Sans files ^(Regular, Bold, Italic, Bold Italic^) before building,
  echo or Typst will substitute another font and pagination will be wrong.
)

echo.
echo Setup complete. Use "Build Hymnal.cmd" to build the hymnal.
echo.
pause
exit /b 0

:fail
echo.
echo SETUP FAILED. Check the messages above, your internet connection,
echo and that no antivirus is blocking the downloads. Then run Setup.cmd again;
echo completed steps will be skipped.
echo.
pause
exit /b 1