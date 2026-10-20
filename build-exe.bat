@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
set "APP=Lectern"
set "LOG=%~dp0build-win-log.txt"
set "VENV=%~dp0.venv-build"

echo ==== %APP% Windows build  %DATE% %TIME% ==== > "%LOG%"
echo.
echo  %APP% - Windows build
echo  (full log: build-win-log.txt)
echo.

echo [1/5] Finding or installing Python...
set "PY="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ensure_python.ps1"`) do if /i "%%~xP"==".exe" set "PY=%%P"
if not defined PY (
    echo Could not find or install Python. See build-win-log.txt.
    echo Could not find or install Python >> "%LOG%"
    goto :fail
)
echo       using %PY%
echo Python: %PY% >> "%LOG%"

echo [2/5] Creating build environment...
if exist "%VENV%\Scripts\python.exe" (
    "%VENV%\Scripts\python.exe" -c "import sys; sys.exit(0)" >nul 2>&1 || rmdir /s /q "%VENV%"
)
if not exist "%VENV%\Scripts\python.exe" (
    "%PY%" -m venv "%VENV%" >> "%LOG%" 2>&1 || goto :fail
)
set "VPY=%VENV%\Scripts\python.exe"
"%VPY%" -m pip install --upgrade pip --only-binary :all: >> "%LOG%" 2>&1

echo [3/5] Installing PyInstaller, PyMuPDF (PDF pages) and the voices...
"%VPY%" -m pip install --only-binary :all: pyinstaller >> "%LOG%" 2>&1 || goto :fail
"%VPY%" -m pip install --only-binary :all: pymupdf >> "%LOG%" 2>&1 || echo       (PyMuPDF did not install - PDFs will not open; everything else works)
rem Read aloud: miniaudio = sound output + MP3 decode; kokoro-onnx = offline neural
rem voices (needs Python 3.10-3.13); edge-tts = Microsoft's online voices. All
rem optional - the app lists whatever installed, and the Windows voices always work.
"%VPY%" -m pip install --only-binary :all: miniaudio >> "%LOG%" 2>&1 || echo       (miniaudio did not install - read aloud uses winsound instead; online voices off)
"%VPY%" -m pip install --only-binary :all: kokoro-onnx >> "%LOG%" 2>&1 || echo       (kokoro-onnx did not install - no offline neural voices)
"%VPY%" -m pip install --only-binary :all: edge-tts >> "%LOG%" 2>&1 || echo       (edge-tts did not install - no online voices)
set "EXTRA="
for /f "usebackq delims=" %%F in (`"%VPY%" build_flags.py`) do set "EXTRA=%%F"
echo       optional packages:
"%VPY%" build_flags.py --list
"%VPY%" build_flags.py --list >> "%LOG%"
"%VPY%" -c "import kokoro_onnx" >nul 2>&1 || goto :build
echo       fetching the Kokoro voice files (about 140 MB, once, into %%LOCALAPPDATA%%\ReadAloud)...
"%VPY%" -c "import readaloud as r; print('      already there' if r.kokoro_models_present() else ('      ok' if r.download_kokoro_models(lambda d,t,n: None) else '      incomplete'))" || echo       (download failed - the app offers it again from its Voice list)

:build

echo [4/5] Building dist\%APP%.exe...
if exist "dist\%APP%.exe" del /q "dist\%APP%.exe"
"%VPY%" -m PyInstaller --noconfirm --clean --onefile --windowed --name "%APP%" ^
  --hidden-import lectern --hidden-import lectern_gui --hidden-import theme --hidden-import readaloud ^
  --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.font ^
  --hidden-import zipfile --hidden-import xml.etree.ElementTree --hidden-import html.parser %EXTRA% ^
  lectern_app.py >> "%LOG%" 2>&1 || goto :fail

echo [5/5] Self-test...
if exist "dist\lectern-selftest.txt" del /q "dist\lectern-selftest.txt"
start "selftest" /wait "dist\%APP%.exe" selftest
echo selftest exit code: %ERRORLEVEL% >> "%LOG%"
set "RC=1"
if exist "dist\lectern-selftest.txt" (
    type "dist\lectern-selftest.txt"
    type "dist\lectern-selftest.txt" >> "%LOG%"
    findstr /C:"RESULT: OK" "dist\lectern-selftest.txt" >nul && set "RC=0"
) else (
    echo (no selftest output written)
)
echo.
if not "%RC%"=="0" (
    echo PROBLEMS FOUND - see above and build-win-log.txt
    pause
    exit /b 1
)
echo Built: dist\%APP%.exe   (settings and the log are written beside the exe)
echo Build OK >> "%LOG%"
pause
exit /b 0

:fail
echo.
echo BUILD FAILED. Last lines of the log:
echo ----------------------------------------
powershell -NoProfile -Command "Get-Content -LiteralPath '%LOG%' -Tail 40"
echo ----------------------------------------
pause
exit /b 1
