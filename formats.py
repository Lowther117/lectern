"""formats - the long tail of file types Lectern opens, all in plain Python.

Everything here builds on lectern.Document and feeds the same word stream:

    Office (old binary)   .doc .xls .ppt            via a small CFB/OLE reader
    Office (OOXML extras) .pptx .pptm .potx .xlsx .xlsm .xltx .docm .dotx .dotm
    OpenDocument          .odt .ods .odp .fodt .fodp
    Rich text             .rtf
    Tables                .csv .tsv
    E-mail                .eml .msg
    Subtitles             .srt .vtt .sbv
    E-books               .mobi .prc .azw .azw3 (PalmDOC only, no DRM), .fb2
    Anything text-like    .json .xml .yaml .ini .log .py .js .sh ... and any
                          file that turns out to be UTF-8/UTF-16/cp1252 text

Files are sniffed by their first bytes as well, so a PDF called .txt or a
.docx called .doc still opens. LibreOffice, when installed, is the last
resort for whatever is left (WordPerfect, Works, StarOffice...).
"""
from __future__ import annotations

import bisect
import csv
import email
import email.policy
import io
import os
import re
import shutil
import struct
import subprocess
import tempfile
import zipfile
from xml.etree import ElementTree as ET

import lectern as core

TEXT_EXTS = {".txt", ".text", ".log", ".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".toml", ".csv",
             ".tsv", ".py", ".js", ".ts", ".css", ".sh", ".bat", ".ps1", ".cmd", ".command", ".c", ".h", ".cpp",
             ".cs", ".java", ".go", ".rs", ".rb", ".php", ".lua", ".sql", ".tex", ".rst", ".nfo", ".asc"}
CODE_EXTS = {".json", ".xml", ".yaml", ".yml", ".ini", ".cfg", ".conf", ".toml", ".py", ".js", ".ts", ".css", ".sh",
             ".bat", ".ps1", ".cmd", ".command", ".c", ".h", ".cpp", ".cs", ".java", ".go", ".rs", ".rb", ".php",
             ".lua", ".sql", ".tex"}
SOFFICE_EXTS = {".wpd", ".wps", ".sxw", ".sxc", ".sxi", ".abw", ".dot", ".xlt", ".pot", ".pages", ".numbers", ".key",
                ".uot", ".602", ".hwp", ".lwp", ".mcw", ".wri"}

EXTENSIONS = {
    ".doc": "doc", ".xls": "xls", ".ppt": "ppt",
    ".pptx": "pptx", ".pptm": "pptx", ".potx": "pptx", ".potm": "pptx", ".ppsx": "pptx",
    ".xlsx": "xlsx", ".xlsm": "xlsx", ".xltx": "xlsx", ".xltm": "xlsx",
    ".docm": "docx", ".dotx": "docx", ".dotm": "docx",
    ".odt": "odf", ".ods": "odf", ".odp": "odf", ".fodt": "fodf", ".fodp": "fodf", ".fods": "fodf",
    ".rtf": "rtf", ".csv": "csv", ".tsv": "csv", ".eml": "eml", ".msg": "msg",
    ".srt": "subs", ".vtt": "subs", ".sbv": "subs",
    ".mobi": "mobi", ".prc": "mobi", ".azw": "mobi", ".azw3": "mobi", ".fb2": "fb2",
}


# ------------------------------------------------------------------ sniffing
def sniff(path):
    """What a file really is, from its first bytes: 'pdf', 'zip:<kind>', 'cfb:<kind>', 'rtf', 'mobi', 'text', or None."""
    with open(path, "rb") as fh:
        head = fh.read(4096)
    if head.startswith(b"%PDF"):
        return "pdf"
    if head.startswith(b"{\\rtf"):
        return "rtf"
    if head.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as z:
                names = set(z.namelist())
            if "word/document.xml" in names:
                return "zip:docx"
            if "ppt/presentation.xml" in names:
                return "zip:pptx"
            if "xl/workbook.xml" in names:
                return "zip:xlsx"
            if "META-INF/container.xml" in names:
                return "zip:epub"
            if "content.xml" in names:
                return "zip:odf"
        except zipfile.BadZipFile:
            pass
        return "zip"
    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        try:
            names = {n.lower() for n in Cfb(path).names()}
        except Exception:
            return "cfb"
        if "worddocument" in names:
            return "cfb:doc"
        if "powerpoint document" in names:
            return "cfb:ppt"
        if "workbook" in names or "book" in names:
            return "cfb:xls"
        if any(n.startswith("__substg1.0_") for n in names):
            return "cfb:msg"
        return "cfb"
    if len(head) > 68 and head[60:68] in (b"BOOKMOBI", b"TEXtREAd"):
        return "mobi"
    if b"\x00" not in head[:1024] or head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "text"
    return None


def open_any(path):
    """The dispatcher lectern.load_document hands anything it does not own to."""
    ext = os.path.splitext(path)[1].lower()
    kind = EXTENSIONS.get(ext)
    real = sniff(path)
    # trust the bytes over the name when they disagree
    if real == "pdf":
        return core.load_pdf(path)
    if real == "zip:docx":
        return core.load_docx(path)
    if real == "zip:epub":
        return core.load_epub(path)
    if real == "zip:pptx":
        return load_pptx(path)
    if real == "zip:xlsx":
        return load_xlsx(path)
    if real == "zip:odf":
        return load_odf(path)
    if real == "rtf":
        return load_rtf(path)
    if real == "cfb:doc":
        return load_doc(path)
    if real == "cfb:ppt":
        return load_ppt(path)
    if real == "cfb:xls":
        return load_xls(path)
    if real == "cfb:msg":
        return load_msg(path)
    if real == "mobi" or kind == "mobi":
        return load_mobi(path)
    if kind == "fodf":
        return load_odf(path, flat=True)
    if kind == "csv":
        return load_csv(path)
    if kind == "eml":
        return load_eml(path)
    if kind == "subs":
        return load_subtitles(path)
    if kind == "fb2":
        return load_fb2(path)
    if ext in (".html", ".htm", ".xhtml"):
        return core.load_html(path)
    if ext in (".md", ".markdown"):
        return core.load_markdown(path)
    if ext in (".txt", ".text"):
        return core.load_text(path)
    if ext in SOFFICE_EXTS or (kind in ("doc", "xls", "ppt") and real in ("cfb", None)):
        conv = soffice_convert(path)
        if conv:
            return core.load_document(conv)
        raise ValueError(f"{ext} needs LibreOffice installed to convert it (free, libreoffice.org); or export it as PDF or Word first")
    if real == "text" or ext in TEXT_EXTS:
        return load_plain(path, mono=ext in CODE_EXTS)
    if real == "zip":
        raise ValueError("That is a zip archive, not a document")
    raise ValueError(f"Lectern does not know how to read {ext or 'that'} files")


# ------------------------------------------------------------ text-likes
def load_plain(path, mono=False):
    d = core.Document(path, "txt")
    src = core._read_text(path)
    if mono:
        for ln in src.split("\n"):
            d.add_block(ln, "mono")
    else:
        for para in re.split(r"\n\s*\n", src):
            d.add_block(para)
    return d


def load_csv(path):
    d = core.Document(path, "csv")
    src = core._read_text(path)
    try:
        dialect = csv.Sniffer().sniff(src[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel_tab if path.lower().endswith(".tsv") else csv.excel
    for row in csv.reader(io.StringIO(src), dialect):
        cells = [c.strip() for c in row if c.strip()]
        if cells:
            d.add_block(" | ".join(cells))
    return d


def load_subtitles(path):
    d = core.Document(path, "subs")
    src = core._read_text(path)
    timing = re.compile(r"^\s*(\d+:)?\d{1,2}:\d{2}[.,]\d{3}\s*-->|^\d+:\d{2}:\d{2}\.\d{3},\d+:\d{2}:\d{2}\.\d{3}\s*$|^WEBVTT|^NOTE\b|^STYLE\b")
    buf = []
    for block in re.split(r"\n\s*\n", src):
        lines = [ln for ln in block.split("\n") if ln.strip() and not timing.match(ln) and not re.fullmatch(r"\s*\d+\s*", ln)]
        txt = " ".join(re.sub(r"<[^>]+>", "", ln).strip() for ln in lines)
        if txt:
            buf.append(txt)
    # cues are fragments of sentences: join them, start a new paragraph at a clear sentence end
    para = []
    for t in buf:
        para.append(t)
        if t.rstrip()[-1:] in ".!?" and len(" ".join(para)) > 300:
            d.add_block(" ".join(para))
            para = []
    if para:
        d.add_block(" ".join(para))
    return d


# ----------------------------------------------------------------- rtf
_RTF_SKIP = {"fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "header", "footer", "headerl", "headerr",
             "headerf", "footerl", "footerr", "footerf", "xe", "tc", "fldinst", "datafield", "themedata",
             "colorschememapping", "latentstyles", "listtable", "listoverridetable", "rsidtbl", "generator",
             "mmathPr", "wgrffmtfilter", "pntext", "bkmkstart", "bkmkend", "annotation", "atnid", "atnauthor"}
_RTF_SPECIAL = {"par": "\n", "line": " ", "tab": "\t", "cell": " | ", "row": "\n", "sect": "\n", "page": "\n",
                "emdash": "\u2014", "endash": "\u2013", "lquote": "\u2018", "rquote": "\u2019", "ldblquote": "\u201c",
                "rdblquote": "\u201d", "bullet": "\u2022", "~": "\u00a0", "-": "", "_": "-", "{": "{", "}": "}",
                "\\": "\\", "\n": "\n", "\r": "\n"}
_RTF_TOKEN = re.compile(rb"\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}]+)", re.S)


def rtf_to_text(raw):
    """RTF bytes -> text with '\\n' paragraphs. Handles \\uN unicode, \\'hh, nested skip destinations."""
    out = []
    stack = []
    skip = 0           # depth at which an ignored destination started (0 = not skipping)
    depth = 0
    ucskip = 1
    pending_skip = 0   # bytes to swallow after \uN
    codepage = "cp1252"
    for m in _RTF_TOKEN.finditer(raw):
        word, arg, hexv, ctrl, brace, text = m.groups()
        if brace == b"{":
            depth += 1
            stack.append((ucskip,))
            continue
        if brace == b"}":
            if skip and depth <= skip:
                skip = 0
            depth -= 1
            if stack:
                (ucskip,) = stack.pop()
            continue
        if skip:
            continue
        if word is not None:
            w = word.decode("ascii")
            if w == "u" and arg is not None:
                cp = int(arg)
                if cp < 0:
                    cp += 65536
                out.append(chr(cp))
                pending_skip = ucskip
            elif w == "uc" and arg is not None:
                ucskip = int(arg)
            elif w in _RTF_SKIP:
                skip = depth
            elif w == "ansicpg" and arg is not None:
                try:
                    codepage = "cp" + arg.decode()
                    "x".encode(codepage)
                except Exception:
                    codepage = "cp1252"
            elif w in _RTF_SPECIAL:
                out.append(_RTF_SPECIAL[w])
            continue
        if hexv is not None:
            if pending_skip:
                pending_skip -= 1
                continue
            try:
                out.append(bytes.fromhex(hexv.decode()).decode(codepage, errors="replace"))
            except Exception:
                pass
            continue
        if ctrl is not None:
            c = ctrl.decode("latin-1")
            if c == "*":
                skip = depth       # \* means "ignore this destination if unknown" - and we know none of them
            elif c in _RTF_SPECIAL:
                out.append(_RTF_SPECIAL[c])
            continue
        if text is not None:
            t = text.replace(b"\r", b"").replace(b"\n", b"")
            if pending_skip:
                cut = min(pending_skip, len(t))
                t = t[cut:]
                pending_skip -= cut
            if t:
                out.append(t.decode(codepage, errors="replace"))
    return "".join(out)


def load_rtf(path):
    d = core.Document(path, "rtf")
    with open(path, "rb") as fh:
        raw = fh.read()
    for para in rtf_to_text(raw).split("\n"):
        d.add_block(para)
    return d


# --------------------------------------------------------- OpenDocument
_T = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"
_TB = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
_DR = "{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}"
_OF = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"


def _odf_text(el):
    out = []

    def walk(e):
        if e.tag == _T + "s":
            out.append(" " * int(e.get(_T + "c") or 1))
        elif e.tag == _T + "tab":
            out.append("\t")
        elif e.tag == _T + "line-break":
            out.append(" ")
        elif e.tag in (_T + "note", _T + "tracked-changes", _OF + "annotation"):
            return
        else:
            if e.text:
                out.append(e.text)
            for ch in e:
                walk(ch)
        if e.tail and e is not el:
            out.append(e.tail)
    walk(el)
    return "".join(out)


def load_odf(path, flat=False):
    d = core.Document(path, "odf")
    if flat:
        root = ET.parse(path).getroot()
        body = root
    else:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read("content.xml"))
            try:
                meta = ET.fromstring(z.read("meta.xml"))
                for el in meta.iter():
                    if el.tag.endswith("}title") and (el.text or "").strip():
                        d.title = el.text.strip()
            except Exception:
                pass
        body = root
    kinds = (_T + "h", _T + "p", _TB + "table-row", _DR + "page")

    def emit(el):
        if el.tag == _DR + "page":
            d.add_heading(el.get(_DR + "name") or f"Slide {len(d.toc) + 1}", 2)
            for sub in el.iter():
                if sub.tag in (_T + "p", _T + "h"):
                    d.add_block(_odf_text(sub))
            return
        if el.tag == _T + "h":
            d.add_heading(_odf_text(el), int(el.get(_T + "outline-level") or 1))
        elif el.tag == _TB + "table-row":
            cells = [re.sub(r"\s+", " ", _odf_text(c)).strip() for c in el if c.tag == _TB + "table-cell"]
            cells = [c for c in cells if c]
            if cells:
                d.add_block(" | ".join(cells))
        elif el.tag == _T + "p":
            d.add_block(_odf_text(el), "p")

    sheet = []

    def is_sheet():     # _spreadsheet(root) scans the whole tree: once per document, not once per table
        if not sheet:
            sheet.append(_spreadsheet(root))
        return sheet[0]

    # walk in document order but do not descend into an element we already emitted
    def walk(el):
        for ch in el:
            if ch.tag in kinds:
                emit(ch)
                continue
            if ch.tag == _TB + "table":
                name = ch.get(_TB + "name")
                if name and root.tag.endswith("document-content") and is_sheet():
                    d.add_heading(name, 2)
            walk(ch)
    walk(body)
    return d


def _spreadsheet(root):
    return any(e.tag == _OF + "spreadsheet" for e in root.iter())


# ---------------------------------------------------------- OOXML extras
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _rels(z, part):
    """{rId: target path} for a part, resolved against the part's folder."""
    folder, name = os.path.split(part)
    rp = (folder + "/" if folder else "") + "_rels/" + name + ".rels"
    out = {}
    try:
        root = ET.fromstring(z.read(rp))
    except KeyError:
        return out
    for rel in root:
        tgt = rel.get("Target") or ""
        if rel.get("TargetMode") == "External":
            continue
        if tgt.startswith("/"):
            full = tgt[1:]
        else:
            full = os.path.normpath(os.path.join(folder, tgt)).replace("\\", "/")
        out[rel.get("Id")] = full
    return out


def load_pptx(path):
    d = core.Document(path, "pptx")
    with zipfile.ZipFile(path) as z:
        try:
            core_xml = ET.fromstring(z.read("docProps/core.xml"))
            for el in core_xml.iter():
                if el.tag.endswith("}title") and (el.text or "").strip():
                    d.title = el.text.strip()
        except Exception:
            pass
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        rels = _rels(z, "ppt/presentation.xml")
        slides = []
        for sid in pres.iter():
            if sid.tag.endswith("}sldId"):
                p = rels.get(sid.get(_R + "id"))
                if p:
                    slides.append(p)
        if not slides:
            slides = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                            key=lambda n: int(re.findall(r"\d+", n)[-1]))
        for n, part in enumerate(slides, 1):
            try:
                root = ET.fromstring(z.read(part))
            except KeyError:
                continue
            paras = []
            title = None
            for sp in root.iter():
                if not sp.tag.endswith("}sp"):
                    continue
                is_title = any(ph.get("type") in ("title", "ctrTitle") for ph in sp.iter() if ph.tag.endswith("}ph"))
                for p in sp.iter(_A + "p"):
                    t = "".join((x.text or "") for x in p.iter() if x.tag in (_A + "t",))
                    t = re.sub(r"\s+", " ", t).strip()
                    if not t:
                        continue
                    if is_title and title is None:
                        title = t
                    else:
                        paras.append(t)
            d.add_heading(f"{n}. {title}" if title else f"Slide {n}", 2)
            for t in paras:
                d.add_block(t)
    return d


_S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _col_index(ref):
    n = 0
    for ch in ref:
        if ch.isalpha():
            n = n * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return n


def load_xlsx(path):
    d = core.Document(path, "xlsx")
    with zipfile.ZipFile(path) as z:
        shared = []
        try:
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst:
                if si.tag == _S + "si":
                    shared.append("".join((t.text or "") for t in si.iter(_S + "t")))
        except KeyError:
            pass
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = _rels(z, "xl/workbook.xml")
        sheets = []
        for sh in wb.iter(_S + "sheet"):
            part = rels.get(sh.get(_R + "id"))
            if part and sh.get("state", "visible") == "visible":
                sheets.append((sh.get("name") or "Sheet", part))
        for name, part in sheets:
            try:
                root = ET.fromstring(z.read(part))
            except KeyError:
                continue
            if len(sheets) > 1 or name != "Sheet1":
                d.add_heading(name, 2)
            for row in root.iter(_S + "row"):
                cells = []
                for c in row:
                    if c.tag != _S + "c":
                        continue
                    t = c.get("t", "n")
                    v = c.find(_S + "v")
                    val = ""
                    if t == "s" and v is not None and v.text is not None:
                        try:
                            val = shared[int(v.text)]
                        except (ValueError, IndexError):
                            val = ""
                    elif t == "inlineStr":
                        val = "".join((x.text or "") for x in c.iter(_S + "t"))
                    elif t == "b" and v is not None:
                        val = "TRUE" if v.text == "1" else "FALSE"
                    elif v is not None and v.text is not None:
                        val = v.text
                        try:
                            f = float(val)
                            val = str(int(f)) if f == int(f) and abs(f) < 1e15 else f"{f:g}"
                        except ValueError:
                            pass
                    val = re.sub(r"\s+", " ", val).strip()
                    if val:
                        cells.append(val)
                if cells:
                    d.add_block(" | ".join(cells))
    return d


# ---------------------------------------------------------------- CFB / OLE
class Cfb:
    """Just enough of the Compound File Binary format to pull streams out of
    .doc, .xls, .ppt and .msg files (MS-CFB)."""

    def __init__(self, path):
        with open(path, "rb") as fh:
            self.data = fh.read()
        h = self.data
        if h[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise ValueError("not an OLE compound file")
        self.ssz = 1 << struct.unpack_from("<H", h, 0x1E)[0]
        self.mssz = 1 << struct.unpack_from("<H", h, 0x20)[0]
        n_fat = struct.unpack_from("<I", h, 0x2C)[0]
        self.dir_start = struct.unpack_from("<I", h, 0x30)[0]
        self.mini_cutoff = struct.unpack_from("<I", h, 0x38)[0]
        self.minifat_start = struct.unpack_from("<I", h, 0x3C)[0]
        self.n_minifat = struct.unpack_from("<I", h, 0x40)[0]
        difat_start = struct.unpack_from("<I", h, 0x44)[0]
        n_difat = struct.unpack_from("<I", h, 0x48)[0]
        difat = list(struct.unpack_from("<109I", h, 0x4C))
        sec = difat_start
        for _ in range(n_difat):
            if sec >= 0xFFFFFFFA:
                break
            blk = self._sector(sec)
            ents = struct.unpack_from(f"<{self.ssz // 4}I", blk)
            difat.extend(ents[:-1])
            sec = ents[-1]
        fat = []
        for sec in difat[:n_fat]:
            if sec >= 0xFFFFFFFA:
                continue
            fat.extend(struct.unpack_from(f"<{self.ssz // 4}I", self._sector(sec)))
        self.fat = fat
        self.dir = self._chain_bytes(self.dir_start)
        self.entries = []
        for off in range(0, len(self.dir), 128):
            e = self.dir[off:off + 128]
            if len(e) < 128:
                break
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, nlen - 2)].decode("utf-16-le", errors="replace") if nlen >= 2 else ""
            typ = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0] & 0xFFFFFFFF
            self.entries.append((name, typ, left, right, child, start, size))
        self.minifat = []
        if self.n_minifat and self.minifat_start < 0xFFFFFFFA:
            mf = self._chain_bytes(self.minifat_start)
            self.minifat = list(struct.unpack_from(f"<{len(mf) // 4}I", mf))
        root = self.entries[0] if self.entries else None
        self.ministream = self._chain_bytes(root[5], root[6]) if root and root[6] else b""
        self.paths = {}
        if root:
            self._walk(root[4], "")

    def _sector(self, n):
        off = (n + 1) * self.ssz
        return self.data[off:off + self.ssz]

    def _chain_bytes(self, start, size=None):
        out = bytearray()
        sec = start
        seen = 0
        while sec < 0xFFFFFFFA and seen < len(self.fat) + 2:
            out += self._sector(sec)
            if sec >= len(self.fat):
                break
            sec = self.fat[sec]
            seen += 1
        return bytes(out[:size]) if size is not None else bytes(out)

    def _mini_bytes(self, start, size):
        out = bytearray()
        sec = start
        seen = 0
        while sec < 0xFFFFFFFA and seen < len(self.minifat) + 2:
            off = sec * self.mssz
            out += self.ministream[off:off + self.mssz]
            if sec >= len(self.minifat):
                break
            sec = self.minifat[sec]
            seen += 1
        return bytes(out[:size])

    def _walk(self, idx, prefix, depth=0):
        if idx >= len(self.entries) or idx == 0xFFFFFFFF or depth > 64:
            return
        name, typ, left, right, child, start, size = self.entries[idx]
        self._walk(left, prefix, depth + 1)
        p = prefix + name
        if typ == 2:
            self.paths[p] = idx
        elif typ == 1:
            self.paths[p + "/"] = idx
            self._walk(child, p + "/", depth + 1)
        self._walk(right, prefix, depth + 1)

    def names(self):
        return list(self.paths)

    def read(self, name):
        idx = self.paths.get(name)
        if idx is None:
            for k, v in self.paths.items():
                if k.lower() == name.lower():
                    idx = v
                    break
        if idx is None:
            raise KeyError(name)
        _, typ, _, _, _, start, size = self.entries[idx]
        if size < self.mini_cutoff:
            return self._mini_bytes(start, size)
        return self._chain_bytes(start, size)


def _word_text_clean(t):
    """Word's control characters -> plain text, one paragraph per line."""
    out = []
    in_field_code = False
    for ch in t:
        if ch == "\x13":
            in_field_code = True
            continue
        if ch == "\x14":
            in_field_code = False
            continue
        if ch == "\x15":
            in_field_code = False
            continue
        if in_field_code:
            continue
        if ch in "\r\x0c":
            out.append("\n")
        elif ch == "\x07":
            out.append(" | ")
        elif ch in "\x0b":
            out.append(" ")
        elif ch < " " and ch != "\t":
            continue
        else:
            out.append(ch)
    return "".join(out)


def load_doc(path):
    """Word 97-2003 binary: the main text through the piece table (MS-DOC)."""
    d = core.Document(path, "doc")
    cfb = Cfb(path)
    wd = cfb.read("WordDocument")
    if struct.unpack_from("<H", wd, 0)[0] != 0xA5EC:
        raise ValueError("not a Word binary document")
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    table_name = "1Table" if flags & 0x0200 else "0Table"
    try:
        tbl = cfb.read(table_name)
    except KeyError:
        tbl = b""
    ccp_text = struct.unpack_from("<I", wd, 0x4C)[0]
    fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x01A2)
    text = ""
    pieces = []
    if tbl and lcb_clx and fc_clx + lcb_clx <= len(tbl):
        clx = tbl[fc_clx:fc_clx + lcb_clx]
        pos = 0
        while pos < len(clx):
            tag = clx[pos]
            if tag == 1:                     # Prc: grpprl to skip
                cb = struct.unpack_from("<H", clx, pos + 1)[0]
                pos += 3 + cb
            elif tag == 2:                   # Pcdt
                lcb = struct.unpack_from("<I", clx, pos + 1)[0]
                plc = clx[pos + 5:pos + 5 + lcb]
                n = (lcb - 4) // 12
                cps = struct.unpack_from(f"<{n + 1}I", plc, 0)
                for k in range(n):
                    pcd = plc[4 * (n + 1) + 8 * k: 4 * (n + 1) + 8 * k + 8]
                    fc = struct.unpack_from("<I", pcd, 2)[0]
                    compressed = bool(fc & 0x40000000)
                    fc &= 0x3FFFFFFF
                    pieces.append((cps[k], cps[k + 1], fc, compressed))
                break
            else:
                break
    if pieces:
        parts = []
        for cp0, cp1, fc, compressed in pieces:
            n = cp1 - cp0
            if compressed:
                parts.append(wd[fc // 2: fc // 2 + n].decode("cp1252", errors="replace"))
            else:
                parts.append(wd[fc: fc + 2 * n].decode("utf-16-le", errors="replace"))
        text = "".join(parts)
    else:
        # Word 6/95 or a damaged piece table: the text run between fcMin and fcMac
        fc_min, fc_mac = struct.unpack_from("<II", wd, 0x18)
        text = wd[fc_min:fc_mac].decode("cp1252", errors="replace")
    if ccp_text and pieces:
        text = text[:ccp_text]
    for para in _word_text_clean(text).split("\n"):
        d.add_block(para)
    try:
        si = cfb.read("\x05SummaryInformation")
        t = _summary_title(si)
        if t:
            d.title = t
    except Exception:
        pass
    return d


def _summary_title(si):
    """PIDSI_TITLE (2) out of a SummaryInformation property set."""
    try:
        n_sections = struct.unpack_from("<I", si, 0x18)[0]
        off = struct.unpack_from("<I", si, 0x2C)[0]
        cnt = struct.unpack_from("<I", si, off + 4)[0]
        cp = "cp1252"
        props = {}
        for k in range(cnt):
            pid, poff = struct.unpack_from("<II", si, off + 8 + 8 * k)
            props[pid] = off + poff
        if 1 in props:
            t, v = struct.unpack_from("<HxxI", si, props[1])
            if t == 2:
                cp = {1200: "utf-16-le", 65001: "utf-8"}.get(v, f"cp{v}")
        if 2 in props:
            t = struct.unpack_from("<H", si, props[2])[0]
            if t == 0x1E:
                n = struct.unpack_from("<I", si, props[2] + 4)[0]
                return si[props[2] + 8: props[2] + 8 + n].decode(cp, errors="replace").rstrip("\x00").strip()
            if t == 0x1F:
                n = struct.unpack_from("<I", si, props[2] + 4)[0]
                return si[props[2] + 8: props[2] + 8 + 2 * n].decode("utf-16-le", errors="replace").rstrip("\x00").strip()
    except Exception:
        return ""
    return ""


def _ppt_records(data, pos, end, out):
    while pos + 8 <= end:
        ver_inst, rtype, rlen = struct.unpack_from("<HHI", data, pos)
        body = pos + 8
        if rlen > end - body:
            rlen = end - body
        if (ver_inst & 0x0F) == 0x0F:              # container
            _ppt_records(data, body, body + rlen, out)
        else:
            out.append((rtype, data[body:body + rlen]))
        pos = body + rlen
        if rlen == 0 and rtype == 0:
            break


def load_ppt(path):
    """PowerPoint 97-2003: text atoms out of the PowerPoint Document stream, by slide."""
    d = core.Document(path, "ppt")
    cfb = Cfb(path)
    data = cfb.read("PowerPoint Document")
    recs = []
    _ppt_records(data, 0, len(data), recs)
    slide = 0
    seen_persist = False
    first_line = True
    for rtype, body in recs:
        if rtype == 0x03F3:                       # SlidePersistAtom: a new slide in SlideListWithText
            seen_persist = True
            slide += 1
            first_line = True
            continue
        if rtype in (0x0FA0, 0x0FA8):             # TextCharsAtom (UTF-16) / TextBytesAtom (cp1252)
            if rtype == 0x0FA0:
                t = body.decode("utf-16-le", errors="replace")
            else:
                t = body.decode("cp1252", errors="replace")
            for line in re.split(r"[\r\x0b]+", t):
                line = line.strip()
                if not line:
                    continue
                if seen_persist and first_line:
                    d.add_heading(f"{slide}. {line}", 2)
                    first_line = False
                else:
                    d.add_block(line)
    if not d.text:
        raise ValueError("no text found in this presentation")
    return d


def _biff_unicode(data, pos, cch, segs=None):
    """Read an XLUnicodeRichExtendedString starting at pos (after cch). Returns (text, new pos).
    `segs` lists CONTINUE boundaries inside data; crossing one re-reads the option byte."""
    flags = data[pos]
    pos += 1
    rich = flags & 0x08
    ext = flags & 0x04
    high = flags & 0x01
    n_runs = cb_ext = 0
    if rich:
        n_runs = struct.unpack_from("<H", data, pos)[0]
        pos += 2
    if ext:
        cb_ext = struct.unpack_from("<I", data, pos)[0]
        pos += 4
    out = []
    left = cch
    while left > 0:
        if segs:
            k = bisect.bisect_right(segs, pos)          # segs is ascending: the first boundary past pos
            nxt = segs[k] if k < len(segs) else None
            room = (nxt - pos) if nxt is not None else len(data) - pos
        else:
            room = len(data) - pos
        take = min(left, room // 2 if high else room)
        if take <= 0:
            if segs and nxt is not None and pos == nxt:
                high = data[pos] & 0x01
                pos += 1
                continue
            break
        if high:
            out.append(data[pos:pos + 2 * take].decode("utf-16-le", errors="replace"))
            pos += 2 * take
        else:
            out.append(data[pos:pos + take].decode("cp1252", errors="replace"))
            pos += take
        left -= take
        if left > 0 and segs and nxt is not None and pos == nxt:
            high = data[pos] & 0x01
            pos += 1
    pos += 4 * n_runs + cb_ext
    return "".join(out), pos


def _rk_value(rk):
    if rk & 0x02:
        v = float(rk >> 2)
    else:
        v = struct.unpack("<d", struct.pack("<Q", (rk & 0xFFFFFFFC) << 32))[0]
    if rk & 0x01:
        v /= 100
    return v


def _num(v):
    return str(int(v)) if v == int(v) and abs(v) < 1e15 else f"{v:g}"


def load_xls(path):
    """Excel 97-2003 BIFF8: shared strings + cell values, one line per row per sheet."""
    d = core.Document(path, "xls")
    cfb = Cfb(path)
    try:
        data = cfb.read("Workbook")
    except KeyError:
        data = cfb.read("Book")
    # pass 1: records with CONTINUE merged (for SST)
    recs = []
    pos = 0
    while pos + 4 <= len(data):
        rtype, rlen = struct.unpack_from("<HH", data, pos)
        body = data[pos + 4:pos + 4 + rlen]
        pos += 4 + rlen
        if rtype == 0x003C and recs:
            prev = recs[-1]
            prev[1].append(body)
            continue
        recs.append((rtype, [body]))
    sst = []
    sheets = []
    for rtype, bodies in recs:
        if rtype == 0x0085:                       # BOUNDSHEET
            b = bodies[0]
            cch = b[6]
            flags = b[7]
            name = b[8:8 + 2 * cch].decode("utf-16-le", errors="replace") if flags & 1 else b[8:8 + cch].decode("cp1252", errors="replace")
            sheets.append((struct.unpack_from("<I", b, 0)[0], name, b[4] & 0x03 == 0))
        elif rtype == 0x00FC:                     # SST
            buf = b"".join(bodies)
            segs = []
            acc = 0
            for b in bodies[:-1]:
                acc += len(b)
                segs.append(acc)
            total = struct.unpack_from("<I", buf, 4)[0]
            p = 8
            for _ in range(total):
                if p + 3 > len(buf):
                    break
                cch = struct.unpack_from("<H", buf, p)[0]
                p += 2
                s, p = _biff_unicode(buf, p, cch, segs)
                sst.append(s)
    # pass 2: per sheet, from its BOF offset to EOF
    for off, name, visible in sheets:
        if not visible:
            continue
        rows = {}
        p = off
        pending_formula = None
        while p + 4 <= len(data):
            rtype, rlen = struct.unpack_from("<HH", data, p)
            b = data[p + 4:p + 4 + rlen]
            p += 4 + rlen
            if rtype == 0x000A:                   # EOF
                break
            try:
                if rtype == 0x00FD:               # LABELSST
                    r, c, _, i = struct.unpack_from("<HHHI", b)
                    rows.setdefault(r, {})[c] = sst[i] if i < len(sst) else ""
                elif rtype == 0x0204:             # LABEL
                    r, c, _, cch = struct.unpack_from("<HHHH", b)
                    s, _ = _biff_unicode(b, 8, cch)
                    rows.setdefault(r, {})[c] = s
                elif rtype == 0x0203:             # NUMBER
                    r, c, _, v = struct.unpack_from("<HHHd", b)
                    rows.setdefault(r, {})[c] = _num(v)
                elif rtype == 0x027E:             # RK
                    r, c, _, rk = struct.unpack_from("<HHHI", b)
                    rows.setdefault(r, {})[c] = _num(_rk_value(rk))
                elif rtype == 0x00BD:             # MULRK
                    r, c0 = struct.unpack_from("<HH", b)
                    n = (len(b) - 6) // 6
                    for k in range(n):
                        rk = struct.unpack_from("<I", b, 4 + 6 * k + 2)[0]
                        rows.setdefault(r, {})[c0 + k] = _num(_rk_value(rk))
                elif rtype == 0x0205:             # BOOLERR
                    r, c, _, v, err = struct.unpack_from("<HHHBB", b)
                    rows.setdefault(r, {})[c] = ("#ERR" if err else ("TRUE" if v else "FALSE"))
                elif rtype == 0x0006:             # FORMULA: cached result
                    r, c, _ = struct.unpack_from("<HHH", b)
                    res = b[6:14]
                    if res[6:8] == b"\xff\xff":
                        if res[0] == 0:
                            pending_formula = (r, c)       # string follows in a STRING record
                        elif res[0] == 1:
                            rows.setdefault(r, {})[c] = "TRUE" if res[2] else "FALSE"
                        elif res[0] == 2:
                            rows.setdefault(r, {})[c] = "#ERR"
                    else:
                        rows.setdefault(r, {})[c] = _num(struct.unpack("<d", res)[0])
                elif rtype == 0x0207 and pending_formula:   # STRING
                    cch = struct.unpack_from("<H", b)[0]
                    s, _ = _biff_unicode(b, 2, cch)
                    rows.setdefault(pending_formula[0], {})[pending_formula[1]] = s
                    pending_formula = None
            except (struct.error, IndexError):
                continue
        if rows and (len(sheets) > 1 or name != "Sheet1"):
            d.add_heading(name, 2)
        for r in sorted(rows):
            cells = [re.sub(r"\s+", " ", str(rows[r][c])).strip() for c in sorted(rows[r])]
            cells = [c for c in cells if c]
            if cells:
                d.add_block(" | ".join(cells))
    return d


# ------------------------------------------------------------------ e-mail
def _email_blocks(d, hdrs, body_text, body_html):
    for label, val in hdrs:
        if val:
            d.add_block(f"{label}: {val}")
    if body_text and body_text.strip():
        for para in re.split(r"\n\s*\n", body_text.replace("\r\n", "\n")):
            d.add_block(para)
    elif body_html:
        core._html_into(d, body_html)


def load_eml(path):
    d = core.Document(path, "eml")
    with open(path, "rb") as fh:
        msg = email.message_from_bytes(fh.read(), policy=email.policy.default)
    d.title = str(msg.get("Subject") or d.title)
    text = html_ = None
    try:
        part = msg.get_body(preferencelist=("plain",))
        if part is not None:
            text = part.get_content()
        part = msg.get_body(preferencelist=("html",))
        if part is not None:
            html_ = part.get_content()
    except Exception:
        pass
    _email_blocks(d, [("From", msg.get("From")), ("To", msg.get("To")), ("Date", msg.get("Date")),
                      ("Subject", msg.get("Subject"))], text, html_)
    atts = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
    if atts:
        d.add_block("Attachments: " + ", ".join(atts), "mono")
    return d


def _msg_prop(cfb, tag):
    for suffix, enc in (("001F", "utf-16-le"), ("001E", "cp1252")):
        try:
            return cfb.read(f"__substg1.0_{tag}{suffix}").decode(enc, errors="replace").rstrip("\x00")
        except KeyError:
            continue
    return ""


def load_msg(path):
    """Outlook .msg (MS-OXMSG): headers and body out of the property streams."""
    d = core.Document(path, "msg")
    cfb = Cfb(path)
    subject = _msg_prop(cfb, "0037")
    if subject:
        d.title = subject
    body = _msg_prop(cfb, "1000")
    html_ = None
    if not body.strip():
        try:
            html_ = cfb.read("__substg1.0_10130102").decode("utf-8", errors="replace")
        except KeyError:
            html_ = None
    _email_blocks(d, [("From", _msg_prop(cfb, "0C1A") or _msg_prop(cfb, "0042")), ("To", _msg_prop(cfb, "0E04")),
                      ("Cc", _msg_prop(cfb, "0E03")), ("Subject", subject)], body, html_)
    atts = []
    for name in cfb.names():
        m = re.match(r"^(__attach_version1\.0_#\d+/)__substg1\.0_3707001F$", name)
        if m:
            atts.append(cfb.read(name).decode("utf-16-le", errors="replace").rstrip("\x00"))
    if atts:
        d.add_block("Attachments: " + ", ".join(atts), "mono")
    return d


# ------------------------------------------------------------------ e-books
def _palmdoc_decompress(data):
    out = bytearray()
    i = 0
    n = len(data)
    while i < n:
        c = data[i]
        i += 1
        if c == 0:
            out.append(0)
        elif c <= 8:
            out += data[i:i + c]
            i += c
        elif c <= 0x7F:
            out.append(c)
        elif c <= 0xBF:
            if i >= n:
                break
            w = (c << 8) | data[i]
            i += 1
            dist = (w >> 3) & 0x7FF
            length = (w & 7) + 3
            for _ in range(length):
                out.append(out[-dist] if dist <= len(out) else 32)
        else:
            out += b" " + bytes([c ^ 0x80])
    return bytes(out)


def load_mobi(path):
    """MOBI / AZW (PalmDOC-compressed or uncompressed, no DRM)."""
    d = core.Document(path, "mobi")
    with open(path, "rb") as fh:
        data = fh.read()
    n_rec = struct.unpack_from(">H", data, 76)[0]
    offs = [struct.unpack_from(">I", data, 78 + 8 * k)[0] for k in range(n_rec)]
    offs.append(len(data))
    rec0 = data[offs[0]:offs[1]]
    compression, _, text_len, text_recs, rec_size, enc_type = struct.unpack_from(">HHIHHH", rec0, 0)
    if enc_type != 0:
        raise ValueError("this book is DRM-protected, so its text cannot be read")
    if compression == 17480:
        raise ValueError("this book uses HUFF/CDIC compression, which Lectern does not decode - convert it with Calibre first")
    codepage = "cp1252"
    trailing = 0
    if rec0[16:20] == b"MOBI":
        hdr_len = struct.unpack_from(">I", rec0, 20)[0]
        cp = struct.unpack_from(">I", rec0, 28)[0]
        codepage = {65001: "utf-8", 1252: "cp1252"}.get(cp, "utf-8")
        if hdr_len >= 0xE4 + 2:
            flags = struct.unpack_from(">H", rec0, 0xF2)[0] if hdr_len >= 0xF4 else 0
            trailing = flags
        # title
        try:
            t_off, t_len = struct.unpack_from(">II", rec0, 0x54)
            title = rec0[t_off:t_off + t_len].decode(codepage, errors="replace").strip()
            if title:
                d.title = title
        except Exception:
            pass
    chunks = []
    for k in range(1, min(text_recs + 1, n_rec)):
        rec = data[offs[k]:offs[k + 1]]
        # strip trailing entries (multibyte / TBS), each sized by a backwards varint
        f = trailing
        bit = 1
        while f:
            if f & 1 and bit > 1:
                size = 0
                shift = 0
                for j in range(1, 5):
                    b = rec[-j]
                    size |= (b & 0x7F) << shift
                    shift += 7
                    if b & 0x80:
                        break
                rec = rec[:-size] if size < len(rec) else b""
            f >>= 1
            bit <<= 1
        if trailing & 1:
            n = (rec[-1] & 3) + 1 if rec else 0
            rec = rec[:-n]
        chunks.append(_palmdoc_decompress(rec) if compression == 2 else rec)
    html_ = b"".join(chunks).decode(codepage, errors="replace")
    core._html_into(d, html_)
    if not d.text:
        raise ValueError("no readable text found in this book")
    return d


def load_fb2(path):
    d = core.Document(path, "fb2")
    root = ET.parse(path).getroot()

    def strip(tag):
        return tag.split("}")[-1]
    for el in root.iter():
        if strip(el.tag) == "book-title" and (el.text or "").strip():
            d.title = el.text.strip()
            break
    for body in root:
        if strip(body.tag) == "body":
            _fb2_walk(body, d, strip)
    return d


def _fb2_walk(el, d, strip):
    for ch in el:
        t = strip(ch.tag)
        if t == "title":
            d.add_heading(" ".join("".join(x.itertext()).strip() for x in ch), 2)
        elif t in ("p", "v", "subtitle", "text-author"):
            d.add_block("".join(ch.itertext()), "quote" if t == "v" else "p")
        elif t == "empty-line":
            continue
        else:
            _fb2_walk(ch, d, strip)


# ----------------------------------------------------------- LibreOffice
def find_soffice():
    cands = []
    if os.name == "nt":
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if base:
                cands.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    else:
        cands += ["/Applications/LibreOffice.app/Contents/MacOS/soffice", os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    for c in cands:
        if os.path.isfile(c):
            return c
    return shutil.which("soffice") or shutil.which("libreoffice")


def soffice_convert(path, to="docx"):
    """Convert with LibreOffice (headless) into a temp folder; returns the new path or None."""
    exe = find_soffice()
    if not exe:
        return None
    out = tempfile.mkdtemp(prefix="lectern-conv-")
    try:
        subprocess.run([exe, "--headless", "--norestore", "--convert-to", to, "--outdir", out, path],
                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
    except Exception:
        return None
    for n in os.listdir(out):
        if n.lower().endswith("." + to):
            return os.path.join(out, n)
    return None


# ------------------------------------------------------------- self-test
def _cfb_write(path, streams):
    """A minimal CFB writer for the self-test: every stream padded to >= 4096
    bytes so nothing needs the mini stream; one FAT sector; v3, 512-byte sectors."""
    ssz = 512
    blobs = []
    for name, data in streams:
        data = data + b"\x00" * max(0, 4096 - len(data))
        blobs.append((name, data))
    # sector layout: 0 = FAT, then the directory sectors, then the streams
    sectors = []
    fat = []
    sectors.append(None)       # FAT placeholder
    fat.append(0xFFFFFFFD)
    n_dir = ((1 + len(blobs)) * 128 + ssz - 1) // ssz
    dir_first = len(sectors)
    for k in range(n_dir):
        sectors.append(None)
        fat.append(len(sectors) if k < n_dir - 1 else 0xFFFFFFFE)
    starts = []
    for name, data in blobs:
        starts.append(len(sectors))
        n = (len(data) + ssz - 1) // ssz
        for k in range(n):
            sectors.append(data[k * ssz:(k + 1) * ssz].ljust(ssz, b"\x00"))
            fat.append(len(sectors) if k < n - 1 else 0xFFFFFFFE)
    # directory: root + one entry per stream, chained as a simple right-linked list
    def entry(name, typ, left, right, child, start, size):
        nm = name.encode("utf-16-le") + b"\x00\x00"
        e = bytearray(128)
        e[:len(nm)] = nm
        struct.pack_into("<H", e, 64, len(nm))
        e[66] = typ
        e[67] = 1
        struct.pack_into("<III", e, 68, left, right, child)
        struct.pack_into("<I", e, 116, start)
        struct.pack_into("<Q", e, 120, size)
        return bytes(e)
    NO = 0xFFFFFFFF
    ents = [entry("Root Entry", 5, NO, NO, 1 if blobs else NO, 0xFFFFFFFE, 0)]
    for k, (name, data) in enumerate(blobs):
        ents.append(entry(name, 2, NO, k + 2 if k + 1 < len(blobs) else NO, NO, starts[k], len(data)))
    dirdata = b"".join(ents).ljust(n_dir * ssz, b"\x00")
    while len(fat) < ssz // 4:
        fat.append(0xFFFFFFFF)
    sectors[0] = struct.pack(f"<{ssz // 4}I", *fat[:ssz // 4])
    for k in range(n_dir):
        sectors[dir_first + k] = dirdata[k * ssz:(k + 1) * ssz]
    h = bytearray(ssz)
    h[:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", h, 0x18, 0x3E)
    struct.pack_into("<H", h, 0x1A, 0x03)
    struct.pack_into("<H", h, 0x1C, 0xFFFE)
    struct.pack_into("<H", h, 0x1E, 9)
    struct.pack_into("<H", h, 0x20, 6)
    struct.pack_into("<I", h, 0x2C, 1)
    struct.pack_into("<I", h, 0x30, dir_first)
    struct.pack_into("<I", h, 0x38, 4096)
    struct.pack_into("<I", h, 0x3C, 0xFFFFFFFE)
    struct.pack_into("<I", h, 0x40, 0)
    struct.pack_into("<I", h, 0x44, 0xFFFFFFFE)
    struct.pack_into("<I", h, 0x48, 0)
    difat = [0] + [0xFFFFFFFF] * 108
    struct.pack_into("<109I", h, 0x4C, *difat)
    with open(path, "wb") as fh:
        fh.write(bytes(h))
        for s in sectors:
            fh.write(s)


def _sample_doc(path):
    text = "Heading line\rFirst paragraph of the old Word file.\rCell A\x07Cell B\x07\r\x13 HYPERLINK x \x14shown\x15 end\r"
    utext = text.encode("utf-16-le")
    wd = bytearray(2048)
    struct.pack_into("<H", wd, 0, 0xA5EC)
    struct.pack_into("<H", wd, 0x0A, 0x0200)          # fWhichTblStm -> 1Table
    struct.pack_into("<I", wd, 0x4C, len(text))       # ccpText
    text_fc = 1024
    wd[text_fc:text_fc + len(utext)] = utext
    # Clx: Pcdt with one uncompressed piece
    n = 1
    plc = struct.pack("<II", 0, len(text)) + struct.pack("<H", 0) + struct.pack("<I", text_fc) + struct.pack("<H", 0)
    clx = b"\x02" + struct.pack("<I", len(plc)) + plc
    tbl = bytearray(512)
    tbl[64:64 + len(clx)] = clx
    struct.pack_into("<II", wd, 0x01A2, 64, len(clx))
    _cfb_write(path, [("WordDocument", bytes(wd)), ("1Table", bytes(tbl))])


def _ppt_atom(rtype, body, ver=0):
    return struct.pack("<HHI", ver, rtype, len(body)) + body


def _sample_ppt(path):
    slide1 = _ppt_atom(0x03F3, b"\x00" * 20) + _ppt_atom(0x0FA0, "Title One".encode("utf-16-le")) + _ppt_atom(0x0FA8, b"Body of slide one.")
    slide2 = _ppt_atom(0x03F3, b"\x00" * 20) + _ppt_atom(0x0FA8, b"Second slide\rwith two lines")
    container = _ppt_atom(0x0FF0, slide1 + slide2, ver=0x0F)
    _cfb_write(path, [("PowerPoint Document", container)])


def _sample_xls(path):
    def rec(t, b):
        return struct.pack("<HH", t, len(b)) + b
    # globals
    def ustr(s):
        return struct.pack("<H", len(s)) + b"\x00" + s.encode("cp1252")
    g = rec(0x0809, b"\x00" * 16)
    sst = struct.pack("<II", 2, 2) + ustr("Name") + struct.pack("<H", 5) + b"\x01" + "Wïdth".encode("utf-16-le")
    g += rec(0x00FC, sst)
    bs_off_pos = None
    g_len_before_bs = len(g)
    # BOUNDSHEET: offset filled later
    bs = struct.pack("<I", 0) + b"\x00\x00" + bytes([5]) + b"\x00" + b"Sales"
    g += rec(0x0085, bs)
    g += rec(0x000A, b"")
    sheet = rec(0x0809, b"\x00" * 16)
    sheet += rec(0x00FD, struct.pack("<HHHI", 0, 0, 0, 0))
    sheet += rec(0x00FD, struct.pack("<HHHI", 0, 1, 0, 1))
    sheet += rec(0x0203, struct.pack("<HHHd", 1, 0, 0, 12.5))
    sheet += rec(0x027E, struct.pack("<HHHI", 1, 1, 0, (7 << 2) | 2))
    sheet += rec(0x0204, struct.pack("<HHHH", 2, 0, 0, 4) + b"\x00Last")
    sheet += rec(0x000A, b"")
    data = bytearray(g + sheet)
    struct.pack_into("<I", data, g_len_before_bs + 4, len(g))     # BOUNDSHEET lbPlyPos
    _cfb_write(path, [("Workbook", bytes(data))])


def _sample_msg(path):
    streams = [("__substg1.0_0037001F", "Hello subject".encode("utf-16-le")),
               ("__substg1.0_0C1A001F", "Sender Person".encode("utf-16-le")),
               ("__substg1.0_0E04001F", "dan@example.org".encode("utf-16-le")),
               ("__substg1.0_1000001F", "Body line one.\r\n\r\nBody line two.".encode("utf-16-le"))]
    _cfb_write(path, streams)


def _sample_mobi(path):
    html_ = b"<html><body><h1>Mobi Book</h1><p>Paragraph one of the book, with words words words words.</p><p>Second paragraph.</p></body></html>"
    rec0 = bytearray(16)
    struct.pack_into(">HHIHHH", rec0, 0, 1, 0, len(html_), 1, 4096, 0)
    recs = [bytes(rec0), html_]
    hdr = bytearray(78)
    hdr[0:4] = b"Test"
    hdr[60:68] = b"BOOKMOBI"
    struct.pack_into(">H", hdr, 76, len(recs))
    off = 78 + 8 * len(recs) + 2
    table = b""
    for k, r in enumerate(recs):
        table += struct.pack(">IBBH", off, 0, 0, k)
        off += len(r)
    with open(path, "wb") as fh:
        fh.write(bytes(hdr) + table + b"\x00\x00" + b"".join(recs))


def selftest_checks(check, tmp):
    import lectern
    p = os.path.join(tmp, "t.rtf")
    with open(p, "wb") as fh:
        fh.write(rb"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Arial;}}{\colortbl;\red0\green0\blue0;}\pard First \b bold\b0  line.\par Second \'e9 and \u8212? dash\par}")
    d = lectern.load_document(p)
    check("rtf: text, hex and unicode escapes, font table skipped", d.text == "First bold line.\nSecond \u00e9 and \u2014 dash", repr(d.text))

    p = os.path.join(tmp, "t.odt")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        z.writestr("content.xml", '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
                   'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0">'
                   '<office:body><office:text><text:h text:outline-level="1">Odt Title</text:h><text:p>Para <text:span>one</text:span>.</text:p>'
                   '<text:list><text:list-item><text:p>Item</text:p></text:list-item></text:list>'
                   '<table:table><table:table-row><table:table-cell><text:p>A1</text:p></table:table-cell><table:table-cell><text:p>B1</text:p></table:table-cell></table:table-row></table:table>'
                   '</office:text></office:body></office:document-content>')
    d = lectern.load_document(p)
    check("odt: heading, span, list, table row", d.text == "Odt Title\nPara one.\nItem\nA1 | B1" and d.toc[0][1] == "Odt Title", repr(d.text))

    p = os.path.join(tmp, "t.pptx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("ppt/presentation.xml", '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId1"/></p:sldIdLst></p:presentation>')
        z.writestr("ppt/_rels/presentation.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="x" Target="slides/slide2.xml"/></Relationships>')
        sl = ('<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree>'
              '<p:sp><p:nvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>'
              '<p:sp><p:txBody><a:p><a:r><a:t>Point </a:t></a:r><a:r><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>')
        z.writestr("ppt/slides/slide1.xml", sl % ("Slide B", "b"))
        z.writestr("ppt/slides/slide2.xml", sl % ("Slide A", "a"))
    d = lectern.load_document(p)
    check("pptx: slides in sldIdLst order, title + body runs joined", d.text == "1. Slide A\nPoint a\n2. Slide B\nPoint b", repr(d.text))

    p = os.path.join(tmp, "t.xlsx")
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/sharedStrings.xml", '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Name</t></si><si><r><t>Wi</t></r><r><t>dth</t></r></si></sst>')
        z.writestr("xl/worksheets/sheet1.xml", '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2"><v>12.5</v></c><c r="B2"><v>7</v></c><c r="C2" t="inlineStr"><is><t>inline</t></is></c></row></sheetData></worksheet>')
    d = lectern.load_document(p)
    check("xlsx: shared strings, numbers, inline strings", d.text == "Data\nName | Width\n12.5 | 7 | inline", repr(d.text))

    p = os.path.join(tmp, "t.csv")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write('name,qty\n"Smith, J",3\n')
    d = lectern.load_document(p)
    check("csv: quoted comma kept", d.text == "name | qty\nSmith, J | 3", repr(d.text))

    p = os.path.join(tmp, "t.srt")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write("1\n00:00:01,000 --> 00:00:02,000\nHello there,\n\n2\n00:00:02,000 --> 00:00:03,000\n<i>general</i> Kenobi.\n")
    d = lectern.load_document(p)
    check("srt: timings and tags dropped, cues joined", d.text == "Hello there, general Kenobi.", repr(d.text))

    p = os.path.join(tmp, "t.eml")
    with open(p, "wb") as fh:
        fh.write(b"From: A <a@x.org>\r\nTo: b@x.org\r\nSubject: Test mail\r\nContent-Type: text/plain\r\n\r\nFirst line.\r\n\r\nSecond para.\r\n")
    d = lectern.load_document(p)
    check("eml: headers + body paragraphs", d.title == "Test mail" and d.text == "From: A <a@x.org>\nTo: b@x.org\nSubject: Test mail\nFirst line.\nSecond para.", repr(d.text))

    p = os.path.join(tmp, "t.doc")
    _sample_doc(p)
    check("doc: sniffed as cfb:doc", sniff(p) == "cfb:doc", str(sniff(p)))
    d = lectern.load_document(p)
    check("doc: piece table text, table cells, field result only",
          d.text == "Heading line\nFirst paragraph of the old Word file.\nCell A | Cell B |\nshown end", repr(d.text))
    p2 = os.path.join(tmp, "renamed.txt")
    shutil.copy(p, p2)
    check("doc: wrong extension still opens by magic", lectern.load_document(p2).text == d.text)

    p = os.path.join(tmp, "t.ppt")
    _sample_ppt(p)
    d = lectern.load_document(p)
    check("ppt: slides from SlidePersistAtom, both atom types", d.text == "1. Title One\nBody of slide one.\n2. Second slide\nwith two lines", repr(d.text))

    p = os.path.join(tmp, "t.xls")
    _sample_xls(p)
    d = lectern.load_document(p)
    check("xls: SST (ansi + unicode), NUMBER, RK, LABEL", d.text == "Sales\nName | W\u00efdth\n12.5 | 7\nLast", repr(d.text))

    p = os.path.join(tmp, "t.msg")
    _sample_msg(p)
    d = lectern.load_document(p)
    check("msg: subject, from, to, body", d.title == "Hello subject" and d.text == "From: Sender Person\nTo: dan@example.org\nSubject: Hello subject\nBody line one.\nBody line two.", repr(d.text))

    p = os.path.join(tmp, "t.mobi")
    _sample_mobi(p)
    d = lectern.load_document(p)
    check("mobi: uncompressed record -> html -> blocks", d.text == "Mobi Book\nParagraph one of the book, with words words words words.\nSecond paragraph.", repr(d.text))
    comp = _palmdoc_decompress(b"\x06Hello \x80\x32\xC1")      # literal run, back-copy (dist 6, len 5), packed " A"
    check("palmdoc: literal, back-reference and packed space", comp == b"Hello Hello A", repr(comp))

    p = os.path.join(tmp, "t.fb2")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write('<?xml version="1.0"?><FictionBook xmlns="http://www.gribuser.ru/xml/fictionbook/2.0"><description><title-info><book-title>FB Title</book-title></title-info></description>'
                 '<body><section><title><p>Chapter 1</p></title><p>Text <emphasis>here</emphasis>.</p></section></body></FictionBook>')
    d = lectern.load_document(p)
    check("fb2: title, chapter heading, paragraph", d.title == "FB Title" and d.text == "Chapter 1\nText here.", repr(d.text))

    p = os.path.join(tmp, "t.json")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write('{"a": 1,\n "b": 2}\n')
    d = lectern.load_document(p)
    check("json: opened as code lines", d.text == '{"a": 1,\n"b": 2}' and all(st == "mono" for _, _, st in d.blocks), repr(d.text))
    p = os.path.join(tmp, "noext")
    with open(p, "w", encoding="utf-8") as fh:
        fh.write("plain words here")
    check("unknown extension, text bytes -> text", lectern.load_document(p).text == "plain words here")
    check("libreoffice lookup does not crash", find_soffice() is None or os.path.isfile(find_soffice()))
