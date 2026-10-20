#!/bin/bash
# build-app.command - build dist/Lectern.app on macOS.  Double-click it.
# Installs whatever is missing itself (Homebrew, a bundle-able Python with Tk 8.6+).
# Full log: build-mac-log.txt beside this file.

cd "$(dirname "$0")" || exit 1
chmod +x build-app.command run.command 2>/dev/null
xattr -d com.apple.quarantine build-app.command run.command 2>/dev/null
APP="Lectern"
LOG="$PWD/build-mac-log.txt"
VENV="$PWD/.venv-build-mac"
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
export HOMEBREW_NO_AUTO_UPDATE=1
: > "$LOG"
exec > >(tee -a "$LOG") 2>&1

echo "==== $APP macOS build  $(date) ===="

ensure_brew() {
    if command -v brew >/dev/null 2>&1; then return 0; fi
    echo "Homebrew not found - installing it (you will be asked for your password once)..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" < /dev/tty
    export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
    command -v brew >/dev/null 2>&1
}

brew_install() {
    ensure_brew || { echo "Homebrew unavailable - cannot install $*"; return 1; }
    echo "brew install $* ..."
    brew install "$@" < /dev/null
}

# A Python that can be bundled: NOT Apple's /usr/bin/python3 (Tk 8.5 - the app
# builds but never opens), 3.9+ with tkinter.TkVersion >= 8.6. Resolved with
# readlink, never by running "python3" (a fresh Mac pops the Xcode CLT installer).
pick_python() {
    local cands=()
    if [ -n "$LECTERN_BUILD_PYTHON" ]; then cands+=("$LECTERN_BUILD_PYTHON"); fi
    # 3.13 first: the offline voices (kokoro-onnx) need Python < 3.14
    for v in 3.13 3.12 3.11 3.10 3.14 3.9; do
        cands+=("/opt/homebrew/opt/python@$v/bin/python$v" "/usr/local/opt/python@$v/bin/python$v"
                "/Library/Frameworks/Python.framework/Versions/$v/bin/python$v")
    done
    cands+=("/opt/homebrew/bin/python3" "/usr/local/bin/python3")
    for c in "${cands[@]}"; do
        [ -x "$c" ] || continue
        local real
        real="$(readlink -f "$c" 2>/dev/null || echo "$c")"
        case "$real" in /usr/bin/*|/System/*|/Library/Developer/CommandLineTools/*|/Applications/Xcode.app/*) continue;; esac
        local out
        out="$("$c" -c 'import sys, tkinter; print(sys.version_info[0], sys.version_info[1], tkinter.TkVersion)' 2>/dev/null)" || continue
        set -- $out
        if [ "$1" -eq 3 ] && [ "$2" -ge 9 ] && awk "BEGIN{exit !($3 >= 8.6)}"; then
            echo "$c"; return 0
        fi
    done
    return 1
}

echo "[1/6] Finding a bundle-able Python..."
PY="$(pick_python)"
if [ -z "$PY" ]; then
    echo "      none found - installing python + python-tk via Homebrew..."
    brew_install python python-tk
    PY="$(pick_python)"
fi
if [ -z "$PY" ]; then
    echo "PROBLEMS FOUND: no Python with Tk 8.6+ could be found or installed. See $LOG"
    read -r -p "Press Enter to close." < /dev/tty
    exit 1
fi
echo "      using $PY ($("$PY" -c 'import sys,tkinter;print(sys.version.split()[0],"Tk",tkinter.TkVersion)'))"

echo "[2/6] Creating build environment..."
if [ -x "$VENV/bin/python" ] && ! "$VENV/bin/python" -c 'import sys' >/dev/null 2>&1; then rm -rf "$VENV"; fi
[ -x "$VENV/bin/python" ] || "$PY" -m venv "$VENV" || exit 1
VPY="$VENV/bin/python"
"$VPY" -m pip install --upgrade pip --only-binary :all: >/dev/null 2>&1

echo "[3/6] Installing PyInstaller, PyMuPDF (PDF pages) and the voices..."
"$VPY" -m pip install --only-binary :all: pymupdf || echo "      (PyMuPDF did not install - PDFs will not open; everything else works)"
# Read aloud: miniaudio = sound output + MP3 decode; kokoro-onnx = offline neural
# voices; edge-tts = Microsoft's online voices. Each is optional - the app
# lists whatever installed, and the OS voices always work.
"$VPY" -m pip install --only-binary :all: miniaudio || echo "      (miniaudio did not install - read aloud uses afplay instead; online voices off)"
"$VPY" -m pip install --only-binary :all: kokoro-onnx || echo "      (kokoro-onnx did not install - no offline neural voices)"
"$VPY" -m pip install --only-binary :all: edge-tts || echo "      (edge-tts did not install - no online voices)"
"$VPY" -m pip install --only-binary :all: pyinstaller || {
    echo "PROBLEMS FOUND: PyInstaller install failed (see $LOG)"; read -r -p "Press Enter to close." < /dev/tty; exit 1; }
EXTRA="$("$VPY" build_flags.py)"
echo "      optional packages:"; "$VPY" build_flags.py --list | sed 's/^/        /'
if "$VPY" -c "import kokoro_onnx" >/dev/null 2>&1; then
    echo "      fetching the Kokoro voice files (about 140 MB, once, into ~/Library/Application Support/ReadAloud)..."
    "$VPY" -c "import readaloud as r; print('      already there' if r.kokoro_models_present() else ('      ok' if r.download_kokoro_models(lambda d,t,n: None) else '      incomplete'))" \
        || echo "      (download failed - the app offers it again from its Voice list)"
fi

echo "[4/6] Building dist/$APP.app..."
rm -rf "dist/$APP.app" "dist/$APP"
"$VPY" -m PyInstaller --noconfirm --clean --windowed --name "$APP" \
    --hidden-import lectern --hidden-import lectern_gui --hidden-import theme --hidden-import readaloud \
    --hidden-import tkinter --hidden-import tkinter.ttk --hidden-import tkinter.font \
    --hidden-import zipfile --hidden-import xml.etree.ElementTree --hidden-import html.parser $EXTRA \
    --osx-bundle-identifier uk.lowther.lectern \
    lectern_app.py || {
    echo "PROBLEMS FOUND: PyInstaller failed (see $LOG)"; read -r -p "Press Enter to close." < /dev/tty; exit 1; }

echo "[5/6] Clearing quarantine + ad-hoc signing (needed on Apple silicon)..."
xattr -cr "dist/$APP.app" 2>/dev/null
codesign --force --deep --sign - "dist/$APP.app" 2>&1 | grep -v "replacing existing signature" || true

echo "[6/6] Self-test..."
rm -f "dist/lectern-selftest.txt"
"dist/$APP.app/Contents/MacOS/$APP" selftest
echo "selftest exit code: $?"
RC=1
if [ -f "dist/lectern-selftest.txt" ]; then
    cat "dist/lectern-selftest.txt"
    grep -q "RESULT: OK" "dist/lectern-selftest.txt" && RC=0
else
    echo "(no selftest output written)"
fi
echo
if [ "$RC" -ne 0 ]; then
    echo "PROBLEMS FOUND - see above and $LOG"
    read -r -p "Press Enter to close." < /dev/tty
    exit 1
fi
echo "Built: dist/$APP.app   (settings and the log are written in the folder next to the .app)"
read -r -p "Done. Press Enter to close." < /dev/tty
