"""Frozen entry point for Lectern (PyInstaller builds) - also fine from source.

    Lectern [file]        open the window (and that document)
    Lectern selftest      print + write lectern-selftest.txt beside the app, exit 0/1

Anything that goes wrong at start-up lands in lectern-crash.log beside the
app and in a dialog - never silence.
"""

import multiprocessing
import os
import sys
import traceback


def _strip_finder_args(argv):
    # Finder launches apps with "-psn_0_12345"; keep argument handling happy.
    return [a for a in argv if not a.startswith("-psn_")]


def _crash(text):
    try:
        import lectern
        p = os.path.join(lectern.HERE, "lectern-crash.log")
    except Exception:
        p = os.path.join(os.path.dirname(os.path.realpath(sys.executable)), "lectern-crash.log")
    try:
        with open(p, "a", encoding="utf-8") as f:
            f.write(text + "\n\n")
    except Exception:
        pass
    try:
        import tkinter
        from tkinter import messagebox
        r = tkinter.Tk()
        r.withdraw()
        messagebox.showerror("Lectern failed to start", text[-2000:] + "\n\nDetails: " + p)
        r.destroy()
    except Exception:
        pass


def main():
    multiprocessing.freeze_support()
    sys.argv = _strip_finder_args(sys.argv)
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.reconfigure(errors="replace")
        except Exception:
            pass
    try:
        import lectern
        mode = sys.argv[1] if len(sys.argv) > 1 else ""
        if mode == "selftest":
            ok, text = lectern.selftest(with_gui="--no-gui" not in sys.argv)
            try:
                print(text)
            except Exception:
                pass
            try:
                with open(os.path.join(lectern.HERE, "lectern-selftest.txt"), "w", encoding="utf-8") as f:
                    f.write(text + "\n")
            except Exception:
                pass
            sys.exit(0 if ok else 1)
        import lectern_gui
        lectern_gui.main(sys.argv[1:])
    except SystemExit:
        raise
    except Exception:
        _crash(traceback.format_exc())
        sys.exit(1)


if __name__ == "__main__":
    main()
