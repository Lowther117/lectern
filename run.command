#!/bin/bash
# run.command - run Lectern from source on macOS. Double-click it.
# Needs Python 3.9+ with Tk 8.6+; pip packages go in a venv it makes itself. Installs Homebrew
# python + python-tk first if the Mac has no usable Python.

cd "$(dirname "$0")" || exit 1
chmod +x build-app.command run.command 2>/dev/null
xattr -d com.apple.quarantine build-app.command run.command 2>/dev/null
export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
export HOMEBREW_NO_AUTO_UPDATE=1

pick_python() {
    local cands=()
    [ -n "$LECTERN_BUILD_PYTHON" ] && cands+=("$LECTERN_BUILD_PYTHON")
    for v in 3.13 3.12 3.11 3.10 3.14 3.9; do
        cands+=("/opt/homebrew/opt/python@$v/bin/python$v" "/usr/local/opt/python@$v/bin/python$v"
                "/Library/Frameworks/Python.framework/Versions/$v/bin/python$v")
    done
    cands+=("/opt/homebrew/bin/python3" "/usr/local/bin/python3")
    for c in "${cands[@]}"; do
        [ -x "$c" ] || continue
        local real out
        real="$(readlink -f "$c" 2>/dev/null || echo "$c")"
        case "$real" in /usr/bin/*|/System/*|/Library/Developer/CommandLineTools/*|/Applications/Xcode.app/*) continue;; esac
        out="$("$c" -c 'import sys, tkinter; print(sys.version_info[0], sys.version_info[1], tkinter.TkVersion)' 2>/dev/null)" || continue
        set -- $out
        if [ "$1" -eq 3 ] && [ "$2" -ge 9 ] && awk "BEGIN{exit !($3 >= 8.6)}"; then echo "$c"; return 0; fi
    done
    return 1
}

PY="$(pick_python)"
if [ -z "$PY" ]; then
    echo "No Python with Tk 8.6+ found - installing one with Homebrew..."
    if ! command -v brew >/dev/null 2>&1; then
        echo "Installing Homebrew (password asked once)..."
        /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)" < /dev/tty
        export PATH="/opt/homebrew/bin:/usr/local/bin:$PATH"
    fi
    brew install python python-tk < /dev/null
    PY="$(pick_python)"
fi
[ -n "$PY" ] || { echo "No Python with Tk 8.6+ available."; read -r -p "Press Enter." < /dev/tty; exit 1; }

# A small venv for PyMuPDF (PDF pages) and the read-aloud voices (miniaudio,
# kokoro-onnx, edge-tts). Lectern opens everything but PDFs without them.
VENV="$PWD/.venv-mac"
if [ ! -x "$VENV/bin/python" ]; then
    echo "First run - creating .venv-mac and installing PyMuPDF + voices..."
    "$PY" -m venv "$VENV" || exit 1
    "$VENV/bin/python" -m pip install --upgrade pip --only-binary :all: >/dev/null 2>&1
    "$VENV/bin/python" -m pip install --only-binary :all: pymupdf || echo "(PyMuPDF did not install - PDFs will not open; everything else works)"
    "$VENV/bin/python" -m pip install --only-binary :all: miniaudio || echo "(miniaudio did not install - read aloud uses afplay; online voices off)"
    "$VENV/bin/python" -m pip install --only-binary :all: kokoro-onnx || echo "(kokoro-onnx did not install - no offline neural voices)"
    "$VENV/bin/python" -m pip install --only-binary :all: edge-tts || echo "(edge-tts did not install - no online voices)"
fi
exec "$VENV/bin/python" "$PWD/lectern_app.py" "$@"
