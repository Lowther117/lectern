@echo off
setlocal EnableExtensions
cd /d "%~dp0"
rem run.bat - run Lectern from source on Windows. Double-click it.
rem First run finds (or installs) Python, makes .venv-win and installs PyMuPDF
rem (PDF pages) and the voices. Everything but PDFs opens without them.
set "VENV=%~dp0.venv-win"

if exist "%VENV%\Scripts\pythonw.exe" goto :run

echo First run - setting up (finds/installs Python, creates .venv-win, installs PyMuPDF + voices)...
set "PY="
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ensure_python.ps1"`) do if /i "%%~xP"==".exe" set "PY=%%P"
if not defined PY (
    echo Could not find or install Python.
    pause
    exit /b 1
)
"%PY%" -m venv "%VENV%" || (pause & exit /b 1)
"%VENV%\Scripts\python.exe" -m pip install --upgrade pip --only-binary :all: >nul 2>&1
"%VENV%\Scripts\python.exe" -m pip install --only-binary :all: pymupdf || echo (PyMuPDF did not install - PDFs will not open; everything else works)
"%VENV%\Scripts\python.exe" -m pip install --only-binary :all: miniaudio || echo (miniaudio did not install - read aloud uses winsound; online voices off)
"%VENV%\Scripts\python.exe" -m pip install --only-binary :all: kokoro-onnx || echo (kokoro-onnx did not install - no offline neural voices)
"%VENV%\Scripts\python.exe" -m pip install --only-binary :all: edge-tts || echo (edge-tts did not install - no online voices)

:run
start "" "%VENV%\Scripts\pythonw.exe" "%~dp0lectern_app.py" %*
exit /b 0
