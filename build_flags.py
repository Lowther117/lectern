"""build_flags.py - print the PyInstaller flags for the optional packages that
are actually installed in the build venv (one line, space separated).

Used by build-exe.bat / build-app.command. A --collect-all for a package that
is not installed fails the whole build, so the flags are worked out here.

    python build_flags.py            -> "--collect-all kokoro_onnx --collect-all ... "
    python build_flags.py --list     -> one "name: yes/no" line per package
"""
import importlib
import sys

# module name -> what to pass. collect-all carries package data (espeak-ng's
# dictionaries, onnxruntime's DLLs, certifi's bundle); a plain hidden-import
# is enough for single-module packages.
OPTIONAL = [
    ("miniaudio", "--hidden-import miniaudio"),
    ("kokoro_onnx", "--collect-all kokoro_onnx"),
    ("onnxruntime", "--collect-all onnxruntime"),
    ("espeakng_loader", "--collect-all espeakng_loader"),
    ("phonemizer", "--collect-all phonemizer"),
    ("segments", "--collect-all segments"),
    ("csvw", "--collect-all csvw"),
    ("clldutils", "--collect-all clldutils"),
    ("language_tags", "--collect-all language_tags"),
    ("joblib", "--collect-submodules joblib"),
    ("edge_tts", "--collect-all edge_tts"),
    ("aiohttp", "--collect-submodules aiohttp"),
    ("certifi", "--collect-all certifi"),
    ("tabulate", "--hidden-import tabulate"),
    ("curl_cffi", "--collect-all curl_cffi"),
    ("pymupdf", "--collect-all pymupdf"),
    ("fitz", "--hidden-import fitz"),
]


def main():
    flags, rows = [], []
    for mod, flag in OPTIONAL:
        try:
            importlib.import_module(mod)
            ok = True
        except Exception:
            ok = False
        rows.append(f"{mod}: {'yes' if ok else 'no'}")
        if ok:
            flags.append(flag)
    if "--list" in sys.argv:
        print("\n".join(rows))
    else:
        print(" ".join(flags))


if __name__ == "__main__":
    main()
