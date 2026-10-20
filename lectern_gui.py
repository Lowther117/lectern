"""Lectern - the window: contents pane, text view, page view (PDF), the
read-aloud strip and the RSVP panel. Built to get through a document fast:
read it aloud with the words lit up, flash it one word at a time, or both.

Run from source:  python lectern_gui.py [file]
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import lectern as core
import readaloud
import theme


def _open_path(path):
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", path])
        elif os.name == "nt":
            os.startfile(path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showerror(core.APP_NAME, f"Could not open {path}\n{e}")


class App:
    def __init__(self, root, settings=None, autostart=True, open_file=None):
        self.root = root
        self.s = settings if settings is not None else core.load_settings()
        for k, v in readaloud.SETTINGS_DEFAULTS.items():
            self.s.setdefault(k, v)
        self.doc = None                 # core.Document
        self.rdoc = None                # readaloud.Doc over doc.text
        self.hl = None                  # readaloud.TextHighlighter
        self.view = "text"
        self.page_no = 0
        self._page_cache = {}
        self._page_img = None
        self._page_scale = 1.0          # canvas px per PDF point
        self._page_items = []
        self._sent_items = []
        self._sent_of_items = None
        self._after = []
        self._loading = False
        self.palette = theme.apply(root, theme.from_settings(self.s))
        root.title(core.APP_NAME)
        try:
            root.geometry(self.s.get("window") or "1240x820")
        except tk.TclError:
            pass
        root.minsize(760, 480)
        self.ui = theme.ui_family(root)
        self.mono = theme.mono_family(root)
        self._build_menu()
        self._build()
        self.set_theme(self.palette["name"])
        root.protocol("WM_DELETE_WINDOW", self.quit)
        root.bind_all("<Control-d>", self._cycle_theme)
        root.bind_class("Entry", "<Control-d>", lambda e: "break")
        root.bind_class("TEntry", "<Control-d>", lambda e: "break")
        root.bind_all("<Control-o>", lambda e: self.open_dialog())
        root.bind_all("<Control-f>", lambda e: (self.ent_find.focus_set(), self.ent_find.select_range(0, "end"), "break")[2])
        root.bind_all("<Control-plus>", lambda e: self.zoom(1))
        root.bind_all("<Control-equal>", lambda e: self.zoom(1))
        root.bind_all("<Control-minus>", lambda e: self.zoom(-1))
        root.bind_all("<F1>", lambda e: self.show_help())
        self.readbar.bind_keys(root)
        if autostart:
            if open_file:
                self._schedule(50, lambda: self.open_path(open_file))
            elif self.s.get("recent"):
                self._schedule(50, self._offer_last)

    # ------------------------------------------------------------ plumbing
    def _schedule(self, ms, fn):
        job = self.root.after(ms, fn)
        self._after.append(job)
        return job

    def log(self, text):
        core.log_line(text)

    def save_settings(self):
        try:
            core.save_settings(self.s)
        except OSError:
            pass

    def status(self, text):
        self.lbl_status.configure(text=text)

    # --------------------------------------------------------------- menus
    def _build_menu(self):
        m = tk.Menu(self.root)
        f = tk.Menu(m, tearoff=0)
        f.add_command(label="Open…   Ctrl+O", command=self.open_dialog)
        self.menu_recent = tk.Menu(f, tearoff=0)
        f.add_cascade(label="Open recent", menu=self.menu_recent)
        f.add_separator()
        f.add_command(label="Save as plain text…", command=self.save_text)
        f.add_command(label="Default save folder…", command=self.choose_save_dir)
        f.add_command(label="Reset save folder to Downloads", command=self.reset_save_dir)
        f.add_separator()
        f.add_command(label="Quit", command=self.quit)
        m.add_cascade(label="File", menu=f)
        self._fill_recent_menu()

        r = tk.Menu(m, tearoff=0)
        r.add_command(label="Read / pause   Ctrl+Space", command=lambda: self.readbar.toggle())
        r.add_command(label="Stop", command=lambda: self.readbar.stop())
        r.add_command(label="Previous sentence   Ctrl+Left", command=lambda: self.readbar.skip(-1))
        r.add_command(label="Next sentence   Ctrl+Right", command=lambda: self.readbar.skip(1))
        r.add_separator()
        self.var_autoadv = tk.BooleanVar(value=bool(self.s.get("read_autoadvance", True)))
        r.add_checkbutton(label="Carry on past the end of the page", variable=self.var_autoadv,
                          command=lambda: (self.s.__setitem__("read_autoadvance", bool(self.var_autoadv.get())), self.save_settings()))
        r.add_command(label="Bigger RSVP word", command=lambda: self.rsvp_size(4))
        r.add_command(label="Smaller RSVP word", command=lambda: self.rsvp_size(-4))
        m.add_cascade(label="Read", menu=r)

        v = tk.Menu(m, tearoff=0)
        self.var_view = tk.StringVar(value="text")
        v.add_radiobutton(label="Text view", value="text", variable=self.var_view, command=lambda: self.set_view("text"))
        v.add_radiobutton(label="Page view (PDF)", value="page", variable=self.var_view, command=lambda: self.set_view("page"))
        v.add_separator()
        v.add_command(label="Bigger   Ctrl++", command=lambda: self.zoom(1))
        v.add_command(label="Smaller   Ctrl+-", command=lambda: self.zoom(-1))
        v.add_separator()
        self.var_theme = tk.StringVar(value=self.palette["name"])
        for name in theme.ORDER:
            v.add_radiobutton(label=theme.label(name), value=name, variable=self.var_theme,
                              command=lambda n=name: self.set_theme(n))
        v.add_command(label="Next theme   Ctrl+D", command=self._cycle_theme)
        m.add_cascade(label="View", menu=v)

        h = tk.Menu(m, tearoff=0)
        h.add_command(label="Keys and how it works   F1", command=self.show_help)
        h.add_command(label="Open voice model folder", command=lambda: _open_path(readaloud.models_dir()))
        h.add_command(label="Open log file", command=lambda: _open_path(os.path.join(core.HERE, core.LOG_FILE)))
        h.add_command(label="About", command=lambda: messagebox.showinfo(
            core.APP_NAME, f"{core.APP_NAME} {core.VERSION}\nRead a document fast: aloud, highlighted, or one word "
                           f"at a time.\n\nSettings folder:\n{core.HERE}\n\nVoice models:\n{readaloud.models_dir()}"))
        m.add_cascade(label="Help", menu=h)
        self.root.config(menu=m)
        self.menubar = m

    def _fill_recent_menu(self):
        self.menu_recent.delete(0, "end")
        rec = [r for r in self.s.get("recent", []) if isinstance(r, dict) and r.get("path")]
        if not rec:
            self.menu_recent.add_command(label="(nothing yet)", state="disabled")
            return
        for r in rec:
            self.menu_recent.add_command(label=f"{r.get('title') or os.path.basename(r['path'])}   -   {r['path']}",
                                         command=lambda p=r["path"]: self.open_path(p))

    # --------------------------------------------------------------- build
    def _build(self):
        r = self.root
        top = ttk.Frame(r, padding=(10, 8, 10, 4))
        top.pack(fill="x")
        ttk.Button(top, text="Open…", style="Accent.TButton", command=self.open_dialog).pack(side="left")
        self.lbl_title = ttk.Label(top, text="Open a PDF, Word, PowerPoint, Excel, OpenDocument, EPUB/MOBI, e-mail, web, Markdown or text file", style="Dim.TLabel")
        self.lbl_title.pack(side="left", padx=12)
        self.btn_theme = ttk.Button(top, text="Theme", command=self._cycle_theme)
        self.btn_theme.pack(side="right")
        ttk.Button(top, text="A+", width=4, command=lambda: self.zoom(1)).pack(side="right", padx=(0, 10))
        ttk.Button(top, text="A−", width=4, command=lambda: self.zoom(-1)).pack(side="right", padx=(0, 4))
        self.btn_view = ttk.Button(top, text="Page view", width=10, command=self.toggle_view)
        self.btn_view.pack(side="right", padx=(0, 10))
        self.var_find = tk.StringVar()
        self.ent_find = ttk.Entry(top, textvariable=self.var_find, width=22)
        self.ent_find.pack(side="right", padx=(0, 10))
        self.ent_find.bind("<Return>", lambda e: self.find(1))
        self.ent_find.bind("<Shift-Return>", lambda e: self.find(-1))
        self.ent_find.bind("<Escape>", lambda e: self.root.focus_set())
        ttk.Label(top, text="Find", style="Dim.TLabel").pack(side="right", padx=(0, 4))

        # read-aloud strip; the RSVP panel lives in a holder above whichever
        # view is showing (an empty holder takes no room)
        self.body = ttk.Frame(r)
        self.rsvp_holder = ttk.Frame(self.body)
        self.rsvp_holder.pack(fill="x")
        self.readbar = readaloud.ReadBar(r, self, get_doc=self._rdoc, on_word=self._on_word, on_end=self._on_end,
                                         rsvp_parent=self.rsvp_holder)
        self.readbar.pack(fill="x")
        self.readbar.rsvp.set_size(int(self.s.get("rsvp_font") or 30))

        pan = ttk.Panedwindow(r, orient="horizontal")
        pan.pack(fill="both", expand=True, padx=10, pady=(0, 4))
        left = ttk.Frame(pan)
        pan.add(left, weight=1)
        ttk.Label(left, text="Contents", style="Dim.TLabel").pack(anchor="w", pady=(0, 2))
        self.toc_tree = ttk.Treeview(left, show="tree", selectmode="browse", style="Side.Treeview")
        tsb = ttk.Scrollbar(left, orient="vertical", command=self.toc_tree.yview)
        self.toc_tree.configure(yscrollcommand=tsb.set)
        self.toc_tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="left", fill="y")
        self.toc_tree.bind("<<TreeviewSelect>>", self._toc_pick)

        pan.add(self.body, weight=5)

        # text view
        self.text_frame = ttk.Frame(self.body)
        self.text = tk.Text(self.text_frame, wrap="word", relief="flat", padx=48, pady=28, spacing1=2, spacing3=8,
                            cursor="arrow", highlightthickness=0)
        self.text_sb = ttk.Scrollbar(self.text_frame, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=self.text_sb.set)
        self.text.pack(side="left", fill="both", expand=True)
        self.text_sb.pack(side="left", fill="y")
        self.text.bind("<Configure>", lambda e: self.text.configure(padx=max(48, (e.width - 960) // 2)))
        self.text.bind("<ButtonRelease-1>", self._text_clicked)
        self.text.tag_configure("ra_sent")
        self.text.tag_configure("ra_word")

        # page view
        self.page_frame = ttk.Frame(self.body)
        pbar = ttk.Frame(self.page_frame)
        pbar.pack(fill="x", pady=(0, 4))
        ttk.Button(pbar, text="◂", width=3, command=lambda: self.show_page(self.page_no - 1)).pack(side="left")
        ttk.Button(pbar, text="▸", width=3, command=lambda: self.show_page(self.page_no + 1)).pack(side="left", padx=(2, 0))
        self.var_page = tk.StringVar()
        self.ent_page = ttk.Entry(pbar, textvariable=self.var_page, width=5)
        self.ent_page.pack(side="left", padx=(8, 0))
        self.ent_page.bind("<Return>", self._page_entered)
        self.lbl_pages = ttk.Label(pbar, text="", style="Dim.TLabel")
        self.lbl_pages.pack(side="left", padx=6)
        ttk.Button(pbar, text="Fit width", command=self.fit_width).pack(side="right")
        self.canvas = tk.Canvas(self.page_frame, highlightthickness=0, bd=0)
        self.cv_sb = ttk.Scrollbar(self.page_frame, orient="vertical", command=self.canvas.yview)
        self.cv_hsb = ttk.Scrollbar(self.page_frame, orient="horizontal", command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.cv_sb.set, xscrollcommand=self.cv_hsb.set)
        self.cv_hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.cv_sb.pack(side="left", fill="y")
        self.canvas.bind("<ButtonRelease-1>", self._page_clicked)
        self.canvas.bind("<Configure>", self._canvas_resized)
        self._bind_wheel(self.canvas)
        self.text_frame.pack(fill="both", expand=True)
        self.btn_view.state(["disabled"])

        bottom = ttk.Frame(r, padding=(10, 0, 10, 6))
        bottom.pack(fill="x")
        self.lbl_status = ttk.Label(bottom, text="", style="Dim.TLabel")
        self.lbl_status.pack(side="left")
        self.lbl_est = ttk.Label(bottom, text="", style="Dim.TLabel")
        self.lbl_est.pack(side="right")
        self._fonts()

        for w in (self.root, self.text, self.canvas):
            w.bind("<Prior>", lambda e: self._page_key(-1))
            w.bind("<Next>", lambda e: self._page_key(1))

    def _bind_wheel(self, w):
        def wheel(e):
            d = -1 if (e.delta > 0 or getattr(e, "num", 0) == 4) else 1
            w.yview_scroll(d * (3 if sys.platform != "darwin" else 1), "units")
        w.bind("<MouseWheel>", wheel)
        w.bind("<Button-4>", wheel)
        w.bind("<Button-5>", wheel)

    def _fonts(self):
        from tkinter import font as tkfont
        fam = "Georgia"
        try:
            fams = set(tkfont.families(self.root))
            for cand in ("Georgia", "Palatino", "Palatino Linotype", "Book Antiqua", "Charter", "Baskerville",
                         "Times New Roman", "DejaVu Serif"):
                if cand in fams:
                    fam = cand
                    break
        except tk.TclError:
            pass
        s = int(self.s.get("font_size") or 14)
        t = self.text
        t.configure(font=(fam, s))
        t.tag_configure("h1", font=(fam, s + 9, "bold"), spacing1=18, spacing3=10)
        t.tag_configure("h2", font=(fam, s + 5, "bold"), spacing1=14, spacing3=8)
        t.tag_configure("h3", font=(fam, s + 2, "bold"), spacing1=10, spacing3=6)
        t.tag_configure("li", lmargin1=36, lmargin2=52)
        t.tag_configure("quote", lmargin1=40, lmargin2=40, rmargin=40)
        t.tag_configure("mono", font=(self.mono, max(8, s - 2)))
        t.tag_configure("pagebreak", font=(self.ui, max(8, s - 4)), spacing1=14, spacing3=14, justify="center")
        for name in ("ra_sent", "ra_word"):
            t.tag_raise(name)

    def set_theme(self, name):
        self.palette = c = theme.apply(self.root, name)
        self.s["theme"] = c["name"]
        self.s["dark"] = c["is_dark"]
        self.var_theme.set(c["name"])
        self.btn_theme.configure(text=f"Theme: {theme.short_label(theme.next_theme(c['name']))}")
        self.text.configure(bg=c["panel"], fg=c["text"], insertbackground=c["accent"],
                            selectbackground=c["sel"], selectforeground=c["text"])
        self.text.tag_configure("pagebreak", foreground=c["dim"])
        self.text.tag_configure("quote", foreground=c["dim"] if c["is_dark"] else c["text"])
        self.canvas.configure(bg=c["sidebar"])
        if self.hl:
            self.hl.apply_palette(c)
        self.readbar.apply_palette(c)
        self._redraw_page_marks()
        self.save_settings()

    def _cycle_theme(self, event=None):
        self.set_theme(theme.next_theme(self.palette["name"]))
        return "break"

    def zoom(self, d):
        if self.view == "page" and self.doc and self.doc.pages:
            z = max(0.4, min(4.0, float(self.s.get("page_zoom") or 1.0) * (1.15 if d > 0 else 1 / 1.15)))
            self.s["page_zoom"] = round(z, 3)
            self.show_page(self.page_no, keep=True)
        else:
            self.s["font_size"] = max(8, min(36, int(self.s.get("font_size") or 14) + d))
            self._fonts()
        self.save_settings()
        return "break"

    def rsvp_size(self, d):
        v = max(14, min(72, int(self.s.get("rsvp_font") or 30) + d))
        self.s["rsvp_font"] = v
        self.readbar.rsvp.set_size(v)
        self.save_settings()

    # ---------------------------------------------------------------- files
    def _offer_last(self):
        rec = self.s.get("recent") or []
        if rec and isinstance(rec[0], dict) and os.path.isfile(rec[0].get("path", "")):
            self.status(f"Last time: {rec[0].get('title') or rec[0]['path']} (File > Open recent)")

    def open_dialog(self):
        p = filedialog.askopenfilename(
            title="Open a document", parent=self.root,
            initialdir=os.path.dirname(self.doc.path) if self.doc else core.default_save_dir(self.s),
            filetypes=[("All documents", " ".join("*" + e for e in core.SUPPORTED)),
                       ("PDF", "*.pdf"), ("Word", "*.docx *.docm *.dotx *.doc *.rtf *.odt"),
                       ("Presentations", "*.pptx *.ppt *.odp"), ("Spreadsheets", "*.xlsx *.xls *.ods *.csv *.tsv"),
                       ("E-books", "*.epub *.mobi *.azw *.azw3 *.fb2"), ("E-mail", "*.eml *.msg"),
                       ("Web / Markdown / text", "*.html *.htm *.xhtml *.md *.markdown *.txt *.srt *.vtt"),
                       ("All files", "*.*")])
        if p:
            self.open_path(p)

    def open_path(self, path, resume=True):
        if self._loading:
            return
        path = os.path.abspath(path)
        if not os.path.isfile(path):
            messagebox.showerror(core.APP_NAME, f"Not found:\n{path}", parent=self.root)
            return
        self._remember_position()
        self.readbar.document_changed()
        self._loading = True
        self.status(f"Opening {os.path.basename(path)}…")
        self.root.update_idletasks()
        try:
            doc = core.load_document(path)
        except Exception as e:
            self._loading = False
            self.status("")
            messagebox.showerror(core.APP_NAME, f"Could not open {os.path.basename(path)}\n\n{e}", parent=self.root)
            self.log(f"open failed: {path}: {e!r}")
            return
        self._loading = False
        self._install(doc)
        rec = core.recent_entry(self.s, path)
        word = int(rec.get("word") or 0) if (rec and resume) else 0
        core.remember_recent(self.s, path, doc.title, word)
        self._fill_recent_menu()
        self.save_settings()
        if word and self.rdoc and word < len(self.rdoc.words):
            self.readbar.set_cursor(word, restart=False)
            self.status(f"Back where you were - word {word + 1:,} of {len(self.rdoc.words):,}")
        else:
            self.status(f"{len(self.rdoc.words):,} words" if self.rdoc else "")
        for note in doc.notes:
            messagebox.showwarning(core.APP_NAME, note, parent=self.root)

    def _install(self, doc):
        self.doc = doc
        self.rdoc = readaloud.tokenize(doc.text)
        self.lbl_title.configure(text=doc.title or os.path.basename(doc.path))
        self.root.title(f"{doc.title or os.path.basename(doc.path)} - {core.APP_NAME}")
        self._page_cache = {}
        self.page_no = 0
        # text view
        t = self.text
        t.configure(state="normal")
        t.delete("1.0", "end")
        pos = 0
        page_starts = {p.start: p.number for p in doc.pages} if doc.pages else {}
        for a, b, st in doc.blocks:
            if a > pos:
                t.insert("end", doc.text[pos:a])
            t.insert("end", doc.text[a:b], (st,))
            pos = b
        if pos < len(doc.text):
            t.insert("end", doc.text[pos:])
        t.configure(state="disabled")
        self.hl = readaloud.TextHighlighter(t, self.rdoc, self.palette)
        self._fonts()
        # contents
        tr = self.toc_tree
        tr.delete(*tr.get_children())
        self._toc_targets = {}
        if doc.toc:
            parents = {0: ""}
            for lvl, title, off in doc.toc:
                lvl = max(1, int(lvl))
                parent = parents.get(lvl - 1, "")
                for k in range(lvl - 1, 0, -1):
                    if k in parents:
                        parent = parents[k]
                        break
                iid = tr.insert(parent, "end", text=title, open=True)
                parents[lvl] = iid
                for k in list(parents):
                    if k > lvl:
                        del parents[k]
                self._toc_targets[iid] = off
        elif doc.pages:
            for p in doc.pages:
                iid = tr.insert("", "end", text=f"Page {p.number + 1}")
                self._toc_targets[iid] = p.start
        else:
            for a, b, st in doc.blocks[:400]:
                iid = tr.insert("", "end", text=doc.text[a:min(b, a + 60)] + ("…" if b - a > 60 else ""))
                self._toc_targets[iid] = a
        # which view
        want = self.s.get("view") or "auto"
        if doc.pages and want != "text":
            self.set_view("page", remember=False)
        else:
            self.set_view("text", remember=False)
        self.btn_view.state(["!disabled"] if doc.pages else ["disabled"])
        self._estimate()

    def _remember_position(self):
        if self.doc and self.rdoc:
            core.remember_recent(self.s, self.doc.path, self.doc.title, self.readbar.cursor)

    def save_text(self):
        if not self.doc:
            return
        p = filedialog.asksaveasfilename(parent=self.root, title="Save as plain text", defaultextension=".txt",
                                         initialdir=core.default_save_dir(self.s),
                                         initialfile=(self.doc.title or "document") + ".txt",
                                         filetypes=[("Text", "*.txt")])
        if p:
            with open(p, "w", encoding="utf-8") as fh:
                fh.write(self.doc.text.replace("\n", "\n\n"))
            self.status(f"Saved {p}")

    def choose_save_dir(self):
        d = filedialog.askdirectory(parent=self.root, title="Default save folder", initialdir=core.default_save_dir(self.s))
        if d:
            self.s["save_dir"] = d
            self.save_settings()

    def reset_save_dir(self):
        self.s["save_dir"] = ""
        self.save_settings()

    # ---------------------------------------------------------------- views
    def toggle_view(self):
        self.set_view("page" if self.view == "text" else "text")

    def set_view(self, view, remember=True):
        if view == "page" and not (self.doc and self.doc.pages):
            view = "text"
        self.view = view
        self.var_view.set(view)
        self.btn_view.configure(text="Text view" if view == "page" else "Page view")
        if view == "page":
            self.text_frame.pack_forget()
            self.page_frame.pack(fill="both", expand=True)
            w = self.rdoc.words[self.readbar.cursor] if self.rdoc and self.rdoc.words else None
            pg = self.doc.page_of_char(w.start) if w else None
            self.show_page(pg.number if pg else 0)
            self._mark_word(self.readbar._last_word)
        else:
            self.page_frame.pack_forget()
            self.text_frame.pack(fill="both", expand=True)
            if self.hl and self.readbar._last_word is not None:
                self.hl.cur = None
                self.hl.show(self.readbar._last_word)
        if remember:
            self.s["view"] = view
            self.save_settings()

    def _page_key(self, d):
        if self.view == "page":
            self.show_page(self.page_no + d)
            return "break"
        self.text.yview_scroll(d, "pages")
        return "break"

    def _page_entered(self, event=None):
        try:
            self.show_page(int(self.var_page.get()) - 1)
        except ValueError:
            pass
        self.canvas.focus_set()

    def fit_width(self):
        if not (self.doc and self.doc.pages):
            return
        pg = self.doc.pages[self.page_no]
        avail = max(200, self.canvas.winfo_width() - 24)
        self.s["page_zoom"] = round(avail / max(1.0, pg.width) / self._dpi_scale(), 3)
        self.show_page(self.page_no, keep=True)
        self.save_settings()

    def _canvas_resized(self, event):
        pass

    def show_page(self, n, keep=False):
        if not (self.doc and self.doc.pages):
            return
        n = max(0, min(len(self.doc.pages) - 1, n))
        pg = self.doc.pages[n]
        if n != self.page_no or not keep:
            self.canvas.yview_moveto(0)
        self.page_no = n
        zoom = float(self.s.get("page_zoom") or 1.0)
        try:
            data, w, h = core.render_pdf_page(self.doc.path, n, zoom * self._dpi_scale(), self._page_cache)
        except Exception as e:
            self.status(f"Could not draw page {n + 1}: {e}")
            return
        self._page_img = tk.PhotoImage(data=data)
        self._page_scale = w / max(1.0, pg.width)
        cv = self.canvas
        cv.delete("all")
        self._page_items, self._sent_items, self._sent_of_items = [], [], None
        cv.create_image(12, 12, image=self._page_img, anchor="nw", tags=("page",))
        cv.configure(scrollregion=(0, 0, w + 24, h + 24))
        self.var_page.set(str(n + 1))
        self.lbl_pages.configure(text=f"of {len(self.doc.pages)}")
        self._mark_word(self.readbar._last_word)

    def _dpi_scale(self):
        try:
            return max(1.0, float(self.root.tk.call("tk", "scaling")) / 1.3333)
        except Exception:
            return 1.0

    def _rect_px(self, r):
        s = self._page_scale
        return (12 + r[0] * s, 12 + r[1] * s, 12 + r[2] * s, 12 + r[3] * s)

    def _mark_word(self, i):
        """Outline the current word (and its sentence) on the page view."""
        cv = self.canvas
        for it in self._page_items:
            cv.delete(it)
        self._page_items = []
        if i is None or not (self.doc and self.doc.pages and self.rdoc):
            for it in self._sent_items:
                cv.delete(it)
            self._sent_items, self._sent_of_items = [], None
            return
        w = self.rdoc.words[i]
        pg = self.doc.page_of_char(w.start)
        if pg is None:
            return
        if pg.number != self.page_no:
            self.show_page(pg.number)
            return
        c = self.palette
        s = self.rdoc.sentences[w.sent]
        if self._sent_of_items != s.i:
            for it in self._sent_items:
                cv.delete(it)
            self._sent_items = []
            for r in pg.rects_for(s.start, s.end):
                x0, y0, x1, y1 = self._rect_px(r)
                self._sent_items.append(cv.create_rectangle(x0 - 1, y0 - 1, x1 + 1, y1 + 1, outline=c["accent"],
                                                            width=1, dash=(2, 3)))
            self._sent_of_items = s.i
        rects = pg.rects_for(w.start, w.end)
        for r in rects:
            x0, y0, x1, y1 = self._rect_px(r)
            self._page_items.append(cv.create_rectangle(x0 - 2, y0 - 2, x1 + 2, y1 + 2, outline=c["accent"], width=3))
            self._page_items.append(cv.create_rectangle(x0 - 2, y1 + 2, x1 + 2, y1 + 5, fill=c["accent"], outline=""))
        if rects:
            self._scroll_canvas_to(rects[0])

    def _scroll_canvas_to(self, r):
        cv = self.canvas
        x0, y0, x1, y1 = self._rect_px(r)
        try:
            total = float(cv.cget("scrollregion").split()[3])
        except Exception:
            return
        top, bottom = cv.yview()
        h = cv.winfo_height() or 1
        vis_top, vis_bot = top * total, top * total + h
        if y0 < vis_top + h * 0.1 or y1 > vis_bot - h * 0.15:
            cv.yview_moveto(max(0.0, (y0 - h * 0.3) / total))

    def _redraw_page_marks(self):
        if self.view == "page" and self.doc and self.doc.pages:
            self._sent_of_items = None
            self._mark_word(self.readbar._last_word)

    # ------------------------------------------------------- reading hooks
    def _rdoc(self):
        return self.rdoc

    def _on_word(self, i):
        if self.hl:
            if self.view == "text":
                self.hl.show(i)
            else:
                self.hl.cur = i      # keep in step without scrolling the hidden widget
        if self.view == "page":
            self._mark_word(i)
        if i is not None and self.rdoc and (i % 25 == 0):
            self._estimate(i)

    def _on_end(self):
        return False            # one document: the end is the end

    def _text_clicked(self, event):
        if not self.rdoc or not self.rdoc.words or self.hl is None:
            return
        try:
            if self.text.tag_ranges("sel"):
                return
        except tk.TclError:
            pass
        i = self.hl.word_at(f"@{event.x},{event.y}")
        if i is not None:
            self.readbar.set_cursor(i)

    def _page_clicked(self, event):
        if not (self.doc and self.doc.pages and self.rdoc):
            return
        cv = self.canvas
        x = (cv.canvasx(event.x) - 12) / self._page_scale
        y = (cv.canvasy(event.y) - 12) / self._page_scale
        off = self.doc.pages[self.page_no].word_at(x, y)
        if off is None:
            return
        self.readbar.set_cursor(self.rdoc.word_at_char(off))

    def _toc_pick(self, event=None):
        sel = self.toc_tree.selection()
        if not sel or not self.rdoc or not self.rdoc.words:
            return
        off = self._toc_targets.get(sel[0])
        if off is None:
            return
        i = self.rdoc.word_at_char(off)
        self.readbar.set_cursor(i)
        if self.view == "text":
            self.text.see(f"1.0+{off}c")

    def find(self, direction=1):
        q = self.var_find.get().strip().lower()
        if not q or not self.rdoc or not self.rdoc.words:
            return
        text = self.doc.text.lower()
        cur = self.rdoc.words[self.readbar.cursor]
        if direction > 0:
            pos = text.find(q, cur.end)
            if pos < 0:
                pos = text.find(q)
        else:
            pos = text.rfind(q, 0, cur.start)
            if pos < 0:
                pos = text.rfind(q)
        if pos < 0:
            self.status(f"'{self.var_find.get()}' not found")
            return
        i = self.rdoc.word_at_char(pos)
        self.readbar.set_cursor(i, restart=False)
        if self.view == "text":
            self.text.see(f"1.0+{pos}c")
        self.status(f"Found at word {i + 1:,}")

    def _estimate(self, i=None):
        if not self.rdoc:
            self.lbl_est.configure(text="")
            return
        n = len(self.rdoc.words)
        i = self.readbar.cursor if i is None else i
        left = max(0, n - i)
        if self.readbar.var_audio.get():
            wpm = 165 * float(self.s.get("tts_speed") or 1.0)
            how = f"{self.s.get('tts_speed') or 1.0:g}× speech"
        else:
            wpm = int(self.s.get("rsvp_wpm") or 350)
            how = f"{wpm} wpm"
        mins = left / max(60, wpm)
        self.lbl_est.configure(text=f"{left:,} words left - about {int(mins) if mins >= 1 else '<1'} min at {how}")

    # ---------------------------------------------------------------- misc
    def show_help(self):
        messagebox.showinfo(core.APP_NAME + " - keys", (
            "Ctrl+O open    Ctrl+F find (Enter next, Shift+Enter previous)\n"
            "Ctrl+Space read / pause    Ctrl+Left / Ctrl+Right previous / next sentence\n"
            "Ctrl+Up / Ctrl+Down speed up / down (WPM when Audio is off)\n"
            "Ctrl+R RSVP panel    Ctrl++ / Ctrl+- bigger / smaller    Ctrl+D theme\n"
            "PageUp / PageDown turn the page (page view) or scroll\n\n"
            "Click any word to start there, or jump there while it is reading.\n"
            "Audio off = silent speed-reading on the WPM clock; Audio on = the voice is the clock.\n"
            "Page view (PDF) outlines the spoken word on the page itself; Text view reflows everything."),
            parent=self.root)

    def quit(self):
        try:
            self.s["window"] = self.root.geometry()
            self._remember_position()
            self.save_settings()
        except Exception:
            pass
        try:
            self.readbar.close()
        except Exception:
            pass
        self.root.destroy()

    def destroy_for_test(self):
        for j in self._after:
            try:
                self.root.after_cancel(j)
            except tk.TclError:
                pass
        try:
            self.readbar.close()
        except Exception:
            pass


# ---------------------------------------------------------------------- main
def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    files = [a for a in argv if not a.startswith("-") and os.path.isfile(a)]
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream is not None:
                stream.reconfigure(errors="replace")
        except Exception:
            pass
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", root.tk.call("tk", "scaling"))
    except tk.TclError:
        pass
    App(root, open_file=files[0] if files else None)
    root.mainloop()


if __name__ == "__main__":
    main()
