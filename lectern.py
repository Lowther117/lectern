"""Lectern - core: open a Word document, PDF, EPUB, HTML, Markdown or text
file as one stream of words with structure (headings, paragraphs, pages and,
for PDFs, the box every word sits in on the page).

    doc = load_document(path)      -> Document
    doc.text                        the whole document as plain text, one
                                    paragraph per line
    doc.blocks                      [(start, end, style)] over doc.text
    doc.toc                         [(level, title, char offset)]
    doc.pages                       PDF only: Page objects with word boxes

The reading itself (voices, highlighting, RSVP) is readaloud.py, shared
with Hayloft. Settings live in lectern.json beside the app; anything you
save goes to Downloads unless you pick another default folder.

    python lectern.py selftest [--no-gui]
"""
from __future__ import annotations

import html
import json
import os
import re
import shutil
import sys
import tempfile
import time
import zipfile
from html.parser import HTMLParser
from xml.etree import ElementTree as ET

APP_NAME = "Lectern"
VERSION = "1.0.0"
SETTINGS_FILE = "lectern.json"
LOG_FILE = "lectern.log"
SUPPORTED = (".pdf", ".docx", ".docm", ".dotx", ".doc", ".rtf", ".odt", ".fodt", ".epub", ".mobi", ".azw", ".azw3", ".fb2",
             ".pptx", ".ppt", ".odp", ".xlsx", ".xls", ".ods", ".csv", ".tsv", ".eml", ".msg", ".srt", ".vtt",
             ".txt", ".md", ".markdown", ".html", ".htm", ".xhtml", ".json", ".xml", ".log")


# ------------------------------------------------------------------ folders
def writable(d):
    try:
        os.makedirs(d, exist_ok=True)
        p = os.path.join(d, ".write-test")
        with open(p, "w") as fh:
            fh.write("x")
        os.remove(p)
        return True
    except OSError:
        return False


def _user_data_dir():
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
    return os.path.join(base, APP_NAME)


def _app_dir():
    """Settings and the log live beside the scripts / the exe; on macOS in the
    folder containing the .app. Per-user data folder if that is read-only."""
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.realpath(__file__))
    exe_dir = os.path.dirname(os.path.realpath(sys.executable))
    parts = exe_dir.split(os.sep)
    if (sys.platform == "darwin" and len(parts) >= 3 and parts[-1] == "MacOS"
            and parts[-2] == "Contents" and parts[-3].endswith(".app")):
        exe_dir = os.path.dirname(os.path.dirname(os.path.dirname(exe_dir)))
    if writable(exe_dir):
        return exe_dir
    fallback = _user_data_dir()
    return fallback if writable(fallback) else exe_dir


HERE = _app_dir()


def downloads_dir():
    if os.name == "nt":
        try:
            import ctypes
            from ctypes import wintypes
            buf = ctypes.c_wchar_p()

            class GUID(ctypes.Structure):
                _fields_ = [("Data1", wintypes.DWORD), ("Data2", wintypes.WORD), ("Data3", wintypes.WORD),
                            ("Data4", wintypes.BYTE * 8)]
            g = GUID(0x374DE290, 0x123F, 0x4565, (wintypes.BYTE * 8)(0x91, 0x64, 0x39, 0xC4, 0x92, 0x5E, 0x46, 0x7B))
            if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(g), 0, None, ctypes.byref(buf)) == 0:
                p = buf.value
                ctypes.windll.ole32.CoTaskMemFree(buf)
                if p and os.path.isdir(p):
                    return p
        except Exception:
            pass
    p = os.path.join(os.path.expanduser("~"), "Downloads")
    return p if os.path.isdir(p) else os.path.expanduser("~")


DEFAULTS = {
    "save_dir": "",
    "theme": "phosphor",
    "dark": True,
    "font_size": 14,
    "page_zoom": 1.0,
    "view": "auto",            # auto = page view for PDFs, text view for the rest
    "window": "",
    "recent": [],              # [{"path", "title", "word", "when"}]
    # read aloud / RSVP (shared keys with readaloud.py)
    "tts_voice": "",
    "tts_speed": 1.0,
    "tts_audio": True,
    "rsvp_wpm": 350,
    "rsvp_show": False,
    "read_autoadvance": True,
    "rsvp_font": 30,
}


def settings_path():
    return os.path.join(HERE, SETTINGS_FILE)


def load_settings():
    s = dict(DEFAULTS)
    try:
        with open(settings_path(), encoding="utf-8") as fh:
            data = json.load(fh)
        if isinstance(data, dict):
            s.update(data)
    except (OSError, ValueError):
        pass
    if not isinstance(s.get("recent"), list):
        s["recent"] = []
    return s


def save_settings(s):
    tmp = settings_path() + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(s, fh, indent=2)
    os.replace(tmp, settings_path())


def default_save_dir(s):
    d = s.get("save_dir") or ""
    return d if d and os.path.isdir(d) else downloads_dir()


def log_line(text):
    try:
        with open(os.path.join(HERE, LOG_FILE), "a", encoding="utf-8") as fh:
            fh.write(time.strftime("%Y-%m-%d %H:%M:%S  ") + text + "\n")
    except OSError:
        pass


def remember_recent(s, path, title, word):
    rec = [r for r in s.get("recent", []) if isinstance(r, dict) and r.get("path") != path]
    rec.insert(0, {"path": path, "title": title, "word": int(word or 0), "when": int(time.time())})
    s["recent"] = rec[:20]


def recent_entry(s, path):
    for r in s.get("recent", []):
        if isinstance(r, dict) and r.get("path") == path:
            return r
    return None


# ---------------------------------------------------------------- document
class Page:
    """One PDF page: size in points and every word's box, tied to doc.text."""
    __slots__ = ("number", "width", "height", "words", "start", "end")

    def __init__(self, number, width, height):
        self.number = number
        self.width = width
        self.height = height
        self.words = []         # [(rects [(x0, y0, x1, y1), ...], char start, char end)]
        self.start = 0          # char range of this page in doc.text
        self.end = 0

    def word_at(self, x, y):
        """Char offset of the word under a point in page coordinates."""
        for rects, a, b in self.words:
            for x0, y0, x1, y1 in rects:
                if x0 - 1 <= x <= x1 + 1 and y0 - 1 <= y <= y1 + 1:
                    return a
        return None

    def rects_for(self, start, end):
        """Boxes of the page words overlapping a char range."""
        out = []
        for rects, a, b in self.words:
            if a < end and b > start:
                out.extend(rects)
        return out


class Document:
    def __init__(self, path, kind):
        self.path = path
        self.kind = kind
        self.title = os.path.splitext(os.path.basename(path))[0] if path else ""
        self.text = ""
        self.blocks = []        # [(start, end, style)]
        self.toc = []           # [(level, title, char offset)]
        self.pages = []         # PDF only
        self.notes = []         # things worth telling the user (no text layer...)

    # building
    def add_block(self, text, style="p"):
        text = re.sub(r"[ \t\r\f\v]+", " ", text.replace("\n", " ")).strip()
        if not text:
            return None
        if self.text:
            self.text += "\n"
        start = len(self.text)
        self.text += text
        self.blocks.append((start, len(self.text), style))
        return start

    def add_heading(self, text, level):
        start = self.add_block(text, f"h{min(3, max(1, level))}")
        if start is not None:
            self.toc.append((level, text.strip()[:120], start))
        return start

    # queries
    def page_of_char(self, off):
        if not self.pages:
            return None
        lo, hi = 0, len(self.pages)
        while lo < hi:
            mid = (lo + hi) // 2
            if self.pages[mid].end <= off:
                lo = mid + 1
            else:
                hi = mid
        return self.pages[min(lo, len(self.pages) - 1)]

    def block_style_at(self, off):
        for a, b, st in self.blocks:
            if a <= off < b:
                return st
        return "p"

    @property
    def word_count(self):
        return len(re.findall(r"\S+", self.text))


def load_document(path):
    """Open anything: formats.open_any sniffs the first bytes, then falls back
    to the extension, and calls the loaders here or in formats.py."""
    import formats
    return formats.open_any(path)


# ---------------------------------------------------------------- text / md
def _read_text(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def load_text(path):
    d = Document(path, "txt")
    for para in re.split(r"\n\s*\n", _read_text(path)):
        d.add_block(para)
    return d


def load_markdown(path):
    d = Document(path, "md")
    src = _read_text(path)
    in_code = False
    code = []
    for para in re.split(r"\n\s*\n", src):
        lines = para.split("\n")
        if in_code or para.lstrip().startswith("```"):
            for ln in lines:
                if ln.strip().startswith("```"):
                    if in_code:
                        d.add_block(" ".join(code), "mono")
                        code = []
                    in_code = not in_code
                else:
                    code.append(ln.strip())
            continue
        first = lines[0].lstrip()
        m = re.match(r"^(#{1,6})\s+(.*)$", first)
        if m and len(lines) == 1:
            d.add_heading(_md_inline(m.group(2)), len(m.group(1)))
            continue
        if len(lines) >= 2 and re.fullmatch(r"=+\s*", lines[1]):
            d.add_heading(_md_inline(first), 1)
            continue
        if len(lines) >= 2 and re.fullmatch(r"-+\s*", lines[1]):
            d.add_heading(_md_inline(first), 2)
            continue
        if all(re.match(r"^\s*([-*+]|\d+[.)])\s+", ln) for ln in lines if ln.strip()):
            for ln in lines:
                if ln.strip():
                    d.add_block(_md_inline(re.sub(r"^\s*([-*+]|\d+[.)])\s+", "", ln)), "li")
            continue
        if first.startswith(">"):
            d.add_block(_md_inline(" ".join(ln.lstrip().lstrip(">").strip() for ln in lines)), "quote")
            continue
        d.add_block(_md_inline(" ".join(ln.strip() for ln in lines)))
    if code:
        d.add_block(" ".join(code), "mono")
    return d


def _md_inline(t):
    t = re.sub(r"!\[([^\]]*)\]\([^)]*\)", r"\1", t)
    t = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", t)
    t = re.sub(r"(\*\*|__)(.+?)\1", r"\2", t)
    t = re.sub(r"(\*|_)(.+?)\1", r"\2", t)
    t = re.sub(r"`([^`]*)`", r"\1", t)
    return t.strip()


# ------------------------------------------------------------------- html
class _HtmlBlocks(HTMLParser):
    BLOCK = {"p", "div", "li", "blockquote", "pre", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "td", "th", "dt", "dd",
             "section", "article", "header", "footer", "figcaption", "br", "hr", "table", "ul", "ol", "aside", "nav"}
    SKIP = {"script", "style", "head", "title", "noscript", "svg", "template"}

    def __init__(self, doc):
        super().__init__(convert_charrefs=True)
        self.doc = doc
        self.buf = []
        self.style = "p"
        self.skip = 0
        self.cell = []

    def _flush(self):
        text = "".join(self.buf)
        self.buf = []
        if not text.strip():
            return
        st = self.style
        if st.startswith("h"):
            self.doc.add_heading(text, int(st[1]))
        else:
            self.doc.add_block(text, st)

    def handle_starttag(self, tag, attrs):
        if tag in self.SKIP:
            self.skip += 1
            return
        if self.skip:
            return
        if tag in self.BLOCK:
            if tag == "td" or tag == "th":
                self.buf.append("  ")
                return
            if tag == "br":
                self.buf.append(" ")
                return
            self._flush()
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                self.style = tag
            elif tag == "blockquote":
                self.style = "quote"
            elif tag == "pre":
                self.style = "mono"
            elif tag == "li":
                self.style = "li"
            elif tag in ("p", "div", "section", "article", "tr"):
                self.style = "p"

    def handle_endtag(self, tag):
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
            return
        if self.skip:
            return
        if tag in self.BLOCK and tag not in ("td", "th", "br"):
            self._flush()
            self.style = "p"

    def handle_data(self, data):
        if not self.skip:
            self.buf.append(data)

    def close(self):
        super().close()
        self._flush()


def _html_into(doc, src):
    p = _HtmlBlocks(doc)
    p.feed(src)
    p.close()


def load_html(path):
    d = Document(path, "html")
    src = _read_text(path)
    m = re.search(r"<title[^>]*>(.*?)</title>", src, re.S | re.I)
    if m:
        d.title = html.unescape(re.sub(r"\s+", " ", m.group(1))).strip() or d.title
    _html_into(d, src)
    return d


# ------------------------------------------------------------------- epub
def load_epub(path):
    d = Document(path, "epub")
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        opf_path = None
        try:
            root = ET.fromstring(z.read("META-INF/container.xml"))
            for el in root.iter():
                if el.tag.endswith("rootfile"):
                    opf_path = el.get("full-path")
                    break
        except Exception:
            pass
        if not opf_path:
            opf_path = next((n for n in names if n.lower().endswith(".opf")), None)
        if not opf_path:
            raise ValueError("not an EPUB (no .opf)")
        opf = ET.fromstring(z.read(opf_path))
        base = os.path.dirname(opf_path)
        ns = {"o": "http://www.idpf.org/2007/opf", "dc": "http://purl.org/dc/elements/1.1/"}
        t = opf.find(".//dc:title", ns)
        if t is not None and (t.text or "").strip():
            d.title = t.text.strip()
        items = {}
        for it in opf.iter():
            if it.tag.endswith("item") and it.get("id"):
                items[it.get("id")] = (it.get("href") or "", it.get("media-type") or "")
        spine = [it.get("idref") for it in opf.iter() if it.tag.endswith("itemref")]
        for idref in spine:
            href, mt = items.get(idref, ("", ""))
            if not href:
                continue
            full = (base + "/" + href if base else href).split("#")[0]
            full = os.path.normpath(full).replace("\\", "/")
            if full not in names:
                continue
            if "html" not in mt and not full.lower().endswith((".xhtml", ".html", ".htm")):
                continue
            try:
                src = z.read(full).decode("utf-8", errors="replace")
            except KeyError:
                continue
            _html_into(d, src)
    return d


# ------------------------------------------------------------------- docx
_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_para_text(p):
    """Text of a paragraph, honouring tabs/breaks, skipping deleted runs."""
    out = []

    def walk(el):
        tag = el.tag
        if tag == _W + "del" or tag == _W + "moveFrom":
            return
        if tag == _W + "t":
            out.append(el.text or "")
        elif tag == _W + "tab":
            out.append("\t")
        elif tag in (_W + "br", _W + "cr"):
            out.append(" ")
        elif tag == _W + "sym":
            out.append(" ")
        elif tag == _W + "footnoteReference" or tag == _W + "endnoteReference":
            return
        for ch in el:
            walk(ch)
    walk(p)
    return "".join(out)


def _docx_style(p):
    ppr = p.find(_W + "pPr")
    if ppr is None:
        return "p", 0
    ps = ppr.find(_W + "pStyle")
    name = (ps.get(_W + "val") if ps is not None else "") or ""
    low = name.lower().replace(" ", "")
    m = re.match(r"^(heading|berschrift|titre|titolo)(\d)$", low)
    if m:
        return "h", int(m.group(2))
    if low in ("title", "subtitle"):
        return "h", 1 if low == "title" else 2
    if "quote" in low:
        return "quote", 0
    if ppr.find(_W + "numPr") is not None or "list" in low:
        return "li", 0
    if low in ("code", "sourcecode", "htmlpreformatted", "plaintext"):
        return "mono", 0
    return "p", 0


def load_docx(path):
    d = Document(path, "docx")
    with zipfile.ZipFile(path) as z:
        try:
            core = ET.fromstring(z.read("docProps/core.xml"))
            for el in core.iter():
                if el.tag.endswith("}title") and (el.text or "").strip():
                    d.title = el.text.strip()
        except Exception:
            pass
        try:
            root = ET.fromstring(z.read("word/document.xml"))
        except KeyError:
            raise ValueError("not a Word .docx (word/document.xml missing)")
    body = root.find(_W + "body")
    if body is None:
        return d

    def do_para(p):
        text = _docx_para_text(p)
        style, lvl = _docx_style(p)
        if style == "h":
            d.add_heading(text, lvl)
        else:
            d.add_block(text, style)

    def do_table(tbl):
        for tr in tbl.iter(_W + "tr"):
            cells = []
            for tc in tr.findall(_W + "tc"):
                ct = " ".join(_docx_para_text(p) for p in tc.iter(_W + "p"))
                ct = re.sub(r"\s+", " ", ct).strip()
                if ct:
                    cells.append(ct)
            if cells:
                d.add_block(" | ".join(cells), "p")

    for el in body:
        if el.tag == _W + "p":
            do_para(el)
        elif el.tag == _W + "tbl":
            do_table(el)
        elif el.tag == _W + "sdt":          # content controls wrap paragraphs
            for sub in el.iter():
                if sub.tag == _W + "p":
                    do_para(sub)
                elif sub.tag == _W + "tbl":
                    do_table(sub)
    return d


# -------------------------------------------------------------------- pdf
def have_pdf():
    try:
        import pymupdf  # noqa: F401
        return True
    except Exception:
        try:
            import fitz  # noqa: F401
            return True
        except Exception:
            return False


def _fitz():
    try:
        import pymupdf
        return pymupdf
    except Exception:
        import fitz
        return fitz


def load_pdf(path):
    if not have_pdf():
        raise RuntimeError("PyMuPDF is not installed, so PDFs cannot be opened (the run/build scripts install it)")
    fitz = _fitz()
    d = Document(path, "pdf")
    pdf = fitz.open(path)
    try:
        meta = pdf.metadata or {}
        if (meta.get("title") or "").strip():
            d.title = meta["title"].strip()
        toc_pages = {}
        try:
            for lvl, title, pno in pdf.get_toc():
                if pno >= 1:
                    toc_pages.setdefault(pno - 1, []).append((lvl, title))
        except Exception:
            pass
        empty_pages = 0
        for pno in range(len(pdf)):
            page = pdf[pno]
            rect = page.rect
            pg = Page(pno, rect.width, rect.height)
            rot = page.rotation_matrix if getattr(page, "rotation", 0) else None
            pg.start = len(d.text) + (1 if d.text else 0)
            for lvl, title in toc_pages.get(pno, []):
                d.toc.append((lvl, title.strip()[:120], pg.start))
            words = page.get_text("words")      # x0, y0, x1, y1, word, block, line, wordno
            if not words:
                empty_pages += 1
            # group into blocks and lines, in reading order
            words.sort(key=lambda w: (w[5], w[6], w[7]))
            cur_block = None
            cur_line = None
            line_buf = []       # [(rects, token)] of the current paragraph (block)
            para = []           # tokens of the current block
            pending = []        # per-token rects for the block

            def flush_block():
                if not para:
                    return
                start = d.add_block(" ".join(para), "p")
                if start is None:
                    return
                # tie every token back to its char range inside the block
                off = start
                for tok, rects in zip(para, pending):
                    pg.words.append((rects, off, off + len(tok)))
                    off += len(tok) + 1
                para.clear()
                pending.clear()

            for x0, y0, x1, y1, w, b, ln, wn in words:
                w = w.strip()
                if not w:
                    continue
                new_line = (ln != cur_line) or (b != cur_block)
                if rot is not None:           # word boxes come unrotated; the picture is rotated
                    rr = fitz.Rect(x0, y0, x1, y1) * rot
                    x0, y0, x1, y1 = rr.x0, rr.y0, rr.x1, rr.y1
                r = (float(x0), float(y0), float(x1), float(y1))
                if new_line and para and para[-1].endswith("-") and len(para[-1]) > 2 and w[:1].islower():
                    # hyphenated across lines: "exam-" + "ple" -> "example", two boxes
                    para[-1] = para[-1][:-1] + w
                    pending[-1] = pending[-1] + [r]
                    cur_block, cur_line = b, ln
                    continue
                if cur_block is not None and b != cur_block:
                    flush_block()
                cur_block, cur_line = b, ln
                para.append(w)
                pending.append([r])
            flush_block()
            pg.end = len(d.text)
            d.pages.append(pg)
        if empty_pages and empty_pages == len(pdf):
            d.notes.append("This PDF has no text layer (scanned pages). Run it through an OCR tool first "
                           "(the OCR PDF Converter), then open the result.")
        elif empty_pages:
            d.notes.append(f"{empty_pages} of {len(pdf)} pages have no text layer and will be skipped when reading.")
    finally:
        pdf.close()
    return d


def render_pdf_page(path, pno, zoom, cache=None):
    """PPM bytes for tk.PhotoImage (no Pillow needed) and the pixel size."""
    fitz = _fitz()
    key = (path, pno, round(zoom, 3))
    if cache is not None and key in cache:
        return cache[key]
    pdf = fitz.open(path)
    try:
        page = pdf[pno]
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        data = pix.tobytes("ppm")
        out = (data, pix.width, pix.height)
    finally:
        pdf.close()
    if cache is not None:
        if len(cache) > 12:
            cache.pop(next(iter(cache)))
        cache[key] = out
    return out


# --------------------------------------------------------------- self-test
def _sample_docx(path):
    doc_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
               '<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Sample Report</w:t></w:r></w:p>'
               '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>'
               '<w:p><w:r><w:t xml:space="preserve">This is the first </w:t></w:r><w:r><w:rPr><w:b/></w:rPr>'
               '<w:t>paragraph.</w:t></w:r><w:r><w:t xml:space="preserve"> It has two sentences.</w:t></w:r></w:p>'
               '<w:p><w:r><w:t>Deleted text follows: </w:t></w:r><w:del><w:r><w:delText>gone</w:delText></w:r></w:del>'
               '<w:ins><w:r><w:t>kept</w:t></w:r></w:ins></w:p>'
               '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Bullet one</w:t></w:r></w:p>'
               '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
               '<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Conclusion</w:t></w:r></w:p>'
               '<w:p><w:r><w:t>The end.</w:t></w:r></w:p>'
               '</w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc_xml)


def _sample_epub(path):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/epub+zip")
        z.writestr("META-INF/container.xml", '<?xml version="1.0"?><container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">'
                   '<rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>')
        z.writestr("OEBPS/content.opf", '<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="id">'
                   '<metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Tiny Book</dc:title></metadata>'
                   '<manifest><item id="c1" href="c1.xhtml" media-type="application/xhtml+xml"/><item id="c2" href="c2.xhtml" media-type="application/xhtml+xml"/></manifest>'
                   '<spine><itemref idref="c1"/><itemref idref="c2"/></spine></package>')
        z.writestr("OEBPS/c1.xhtml", "<html><head><style>p{}</style></head><body><h1>Chapter One</h1><p>Hello <b>bold</b> world.</p></body></html>")
        z.writestr("OEBPS/c2.xhtml", "<html><body><h2>Two</h2><p>Second chapter here.</p><ul><li>item</li></ul></body></html>")


def selftest(with_gui=True):
    lines = [f"{APP_NAME} {VERSION} self-test", f"python {sys.version.split()[0]} on {sys.platform}",
             f"app folder: {HERE}"]
    problems = []

    def check(name, cond, detail=""):
        lines.append(f"  [{'ok' if cond else 'FAIL'}] {name}" + (f" - {detail}" if detail and not cond else ""))
        if not cond:
            problems.append(name)

    tmp = tempfile.mkdtemp(prefix="lectern-selftest-")
    try:
        s = dict(DEFAULTS)
        remember_recent(s, "/a.pdf", "A", 12)
        remember_recent(s, "/b.docx", "B", 0)
        remember_recent(s, "/a.pdf", "A", 30)
        check("recent list: dedupe, newest first", [r["path"] for r in s["recent"]] == ["/a.pdf", "/b.docx"] and s["recent"][0]["word"] == 30)

        # docx
        dp = os.path.join(tmp, "s.docx")
        _sample_docx(dp)
        d = load_docx(dp)
        check("docx: title/headings in toc", [t for _, t, _ in d.toc] == ["Sample Report", "Introduction", "Conclusion"], str(d.toc))
        check("docx: runs joined, tracked deletion dropped, insertion kept",
              "This is the first paragraph. It has two sentences." in d.text and "Deleted text follows: kept" in d.text and "gone" not in d.text, repr(d.text))
        check("docx: list + table rows", any(st == "li" for _, _, st in d.blocks) and "Cell A | Cell B" in d.text, repr(d.text))
        check("docx: one paragraph per line", d.text.count("\n") == 7, str(d.text.count("\n")))
        check("docx: generic loader dispatch", load_document(dp).text == d.text)

        # text + markdown
        tp = os.path.join(tmp, "t.md")
        with open(tp, "w", encoding="utf-8") as fh:
            fh.write("# Heading\n\nSome *emphasis* and a [link](http://x).\n\n- one\n- two\n\n> quoted\n\n```\ncode here\n```\n")
        d = load_document(tp)
        check("markdown: heading, inline stripped, list, quote, code",
              d.toc and d.toc[0][1] == "Heading" and "Some emphasis and a link." in d.text and
              [st for _, _, st in d.blocks] == ["h1", "p", "li", "li", "quote", "mono"], str(d.blocks) + repr(d.text))
        tp2 = os.path.join(tmp, "t.txt")
        with open(tp2, "wb") as fh:
            fh.write(b"Para one\nstill one.\n\nPara two \x93quoted\x94.")
        d = load_document(tp2)
        check("text: paragraphs by blank line, cp1252 fallback", d.text == "Para one still one.\nPara two “quoted”.", repr(d.text))

        # epub + html
        ep = os.path.join(tmp, "b.epub")
        _sample_epub(ep)
        d = load_document(ep)
        check("epub: title, spine order, headings, style skipped",
              d.title == "Tiny Book" and d.text == "Chapter One\nHello bold world.\nTwo\nSecond chapter here.\nitem"
              and [t for _, t, _ in d.toc] == ["Chapter One", "Two"], repr(d.text))
        hp = os.path.join(tmp, "p.html")
        with open(hp, "w", encoding="utf-8") as fh:
            fh.write("<html><head><title>Page &amp; Title</title><script>var x=1;</script></head><body><h2>H</h2><p>One<br>two</p><div>Three</div></body></html>")
        d = load_document(hp)
        check("html: title unescaped, script skipped, blocks", d.title == "Page & Title" and d.text == "H\nOne two\nThree", repr(d.text))

        # pdf
        if have_pdf():
            fitz = _fitz()
            pp = os.path.join(tmp, "p.pdf")
            pdf = fitz.open()
            page = pdf.new_page()
            page.insert_text((72, 100), "Hello wonderful world of PDF read-", fontsize=14)
            page.insert_text((72, 120), "ing. Second sentence here.", fontsize=14)
            page2 = pdf.new_page()
            page2.insert_text((72, 100), "Page two text.", fontsize=14)
            pdf.set_toc([[1, "Start", 1], [1, "Second page", 2]])
            pdf.save(pp)
            pdf.close()
            d = load_document(pp)
            check("pdf: two pages, text in order", len(d.pages) == 2 and d.text.startswith("Hello wonderful world of PDF") and "Page two text." in d.text, repr(d.text))
            check("pdf: hyphen across lines joined with two boxes",
                  "reading." in d.text and any(len(r) == 2 and d.text[a:b] == "reading." for r, a, b in d.pages[0].words), repr(d.text))
            check("pdf: outline -> toc with page offsets", [t for _, t, _ in d.toc] == ["Start", "Second page"] and d.toc[1][2] == d.pages[1].start, str(d.toc))
            off = d.text.index("wonderful")
            pg = d.page_of_char(off)
            rects = pg.rects_for(off, off + 9)
            check("pdf: word box lookup both ways", pg.number == 0 and len(rects) == 1 and pg.word_at((rects[0][0] + rects[0][2]) / 2, (rects[0][1] + rects[0][3]) / 2) == off, str(rects))
            data, w, h = render_pdf_page(pp, 0, 1.0, {})
            check("pdf: page renders to PPM", data.startswith(b"P6") and w > 500 and h > 700, f"{w}x{h}")
        else:
            lines.append("  [info] PyMuPDF not installed - PDF checks skipped")

        # the long tail of formats
        import formats
        formats.selftest_checks(check, tmp)

        # the shared reading module
        import readaloud
        for k, (okk, why) in readaloud.selftest_checks(check).items():
            lines.append(f"  [info] voice engine {k}: {'available' if okk else 'not available - ' + why}")

        if with_gui:
            real_here = HERE
            globals()["HERE"] = tmp
            try:
                import tkinter as tk
            except Exception as ex:
                lines.append(f"  [skip] GUI not tested: tkinter missing ({ex})")
                with_gui = False
        if with_gui:
            try:
                import theme
                root = tk.Tk()
                root.withdraw()
                import lectern_gui
                app = lectern_gui.App(root, settings=dict(DEFAULTS), autostart=False)
                for name in theme.ORDER:
                    app.set_theme(name)
                app.open_path(dp)
                def _count(parent=""):
                    return sum(1 + _count(c) for c in app.toc_tree.get_children(parent))
                check("gui: docx opens, text view, toc rows (Heading 2 nested)", app.doc is not None and app.view == "text"
                      and len(app.toc_tree.get_children()) == 2 and _count() == 3, f"{app.view} {_count()}")
                app.readbar.var_audio.set(False)
                app.readbar.var_rsvp.set(True)
                app.readbar._rsvp_toggled()
                app.readbar.play(2)
                root.update()
                check("gui: timer mode highlights word 3", app.hl is not None and app.hl.cur == 2, str(app.hl.cur if app.hl else None))
                app.readbar.stop()
                if have_pdf():
                    app.open_path(pp)
                    root.update()
                    check("gui: pdf opens in page view", app.view == "page" and app.page_no == 0, app.view)
                    app.readbar.play(0)
                    root.update()
                    app.readbar.stop()
                    app.set_view("text")
                    check("gui: pdf text view", app.view == "text")
                app.destroy_for_test()
                root.destroy()
                check("gui: builds in all three themes", True)
            except tk.TclError as ex:
                lines.append(f"  [skip] GUI not tested: no display ({ex})")
            except Exception:
                import traceback
                check("gui: builds and opens documents", False, traceback.format_exc()[-800:])
            finally:
                globals()["HERE"] = real_here
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    ok = not problems
    lines.append("RESULT: OK" if ok else "RESULT: PROBLEMS FOUND - " + ", ".join(problems))
    return ok, "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "selftest":
        ok, text = selftest(with_gui="--no-gui" not in sys.argv)
        print(text)
        sys.exit(0 if ok else 1)
    print(__doc__)
