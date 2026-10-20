# Lectern

A reader built to get you through a document fast. Open a PDF, Word,
PowerPoint or Excel file (new or old), an OpenDocument, EPUB, MOBI, e-mail,
web page, Markdown or text file and Lectern reads it out loud with
each word lit up as it is spoken, flashes it one word at a time (RSVP -
Rapid Serial Visual Presentation) at whatever words-per-minute you like, or
does both at once from the same clock. For PDFs it works on the page itself,
outlining the spoken word where it sits. Python 3.9+ with Tk; PyMuPDF and
the voice packages are installed by the scripts.

## Which file do I use?

| Want to... | Windows | macOS |
|---|---|---|
| Run from source | `run.bat` (drop a file on it, or open from the window) | `run.command` |
| Build a standalone app | `build-exe.bat` -> `dist\Lectern.exe` | `build-app.command` -> `dist/Lectern.app` |
| Check everything works | `Lectern.exe selftest` | `Lectern selftest` |

Both build scripts install what they need themselves (a real Python with Tk,
PyInstaller, PyMuPDF, the voices), fetch the Kokoro voice files once, log to
`build-win-log.txt` / `build-mac-log.txt`, and finish by running the app's
self-test. Settings (`lectern.json`) and the log live beside the exe / next
to the .app.

## Using it

1. **Open** a document (button, Ctrl+O, File > Open recent, or pass the file
   on the command line / drop it on `run.bat`). Where you were in it is
   remembered. The full list of formats is at the end.
2. The **Contents** pane lists headings (Word, EPUB, HTML, Markdown), the
   outline or pages (PDF); click one to jump.
3. **▶ Read** starts from the top, or from any word you click. While it is
   reading, clicking a word jumps there. ■ stops; ⏮ ⏭ step a sentence.
4. **Voice** - everything installed is listed:
   *Kokoro* - offline neural voices, the best sounding, British and American.
   `kokoro-onnx` (Python 3.10-3.13) plus ~140 MB of model files in
   `%LOCALAPPDATA%\ReadAloud\models` / `~/Library/Application
   Support/ReadAloud/models`, shared with Hayloft; the build scripts fetch
   them, and the Voice list offers the download if they are missing.
   *Online* - Microsoft's neural voices through `edge-tts`; needs the
   internet, gives exact word timing. *Fetch the full online voice list…*
   adds the other 300-odd voices.
   *System* - Windows SAPI voices or macOS `say` voices. Always there.
5. **Speed** 0.5×-3× of the voice's natural rate.
6. **Audio** off turns the same controls into a silent speed-reader: the
   highlight and the RSVP panel run on their own clock at **WPM**
   (100-1000). That is how RSVP goes past what speech can do.
7. **RSVP** shows the panel: each word pinned on its recognition letter (the
   accent-coloured one) so your eye never moves. Read > Bigger/Smaller RSVP
   word sets the size; Ctrl+R toggles the panel.
8. **Page view / Text view** (PDF only): page view shows the real page with
   the spoken word outlined and underlined and the rest of its sentence
   dotted; A+/A-/Fit width zoom it. Text view reflows the words like any
   other document. The choice is remembered.
9. **Find** (Ctrl+F; Enter next, Shift+Enter previous) moves the reading
   cursor to the match, so ▶ carries on from there.

The status bar shows how many words are left and roughly how long they will
take at the current speed.

### Keys

| Key | Does |
|---|---|
| Ctrl+Space | read / pause |
| Ctrl+Left / Ctrl+Right | previous / next sentence |
| Ctrl+Up / Ctrl+Down | speed up / down (WPM when Audio is off) |
| Ctrl+R | RSVP panel |
| Ctrl+F | find |
| Ctrl++ / Ctrl+- | bigger / smaller (page zoom in page view) |
| PageUp / PageDown | turn the page (page view) or scroll |
| Ctrl+D | next theme |
| F1 | this list |

## How the sync works

One word stream drives everything. The text is split into words and
sentences with their character offsets; the voice engine returns one audio
clip per sentence with a start time for every word; the playback position
is the clock, and the page highlight, the PDF outline and the RSVP flash all
read it. With Audio off the clock is a timer running at the chosen WPM, with
longer dwell on long words, numbers, commas and sentence ends.

The online voices report exact word boundaries. Kokoro and the system
voices give sentence-exact clips with the words inside placed by phoneme
count (Kokoro) or letter count, which lands within about a tenth of a
second. Sound goes through `miniaudio` when installed (it also decodes the
online voices' MP3); without it, one `winsound`/`afplay` call per sentence.

## What it opens

Files are recognised by their first bytes as well as their name, so a PDF
called `.txt` or a `.docx` called `.doc` still opens. Everything is read in
plain Python except PDFs (PyMuPDF).

| Kind | How |
|---|---|
| `.pdf` | PyMuPDF: text in reading order with every word's box; hyphenated line breaks re-joined; outline -> Contents; rotated pages handled. A scanned PDF with no text layer gets a warning (run it through the OCR PDF Converter first). |
| `.docx` `.docm` `.dotx` `.dotm` | The document XML directly: paragraphs, heading styles -> Contents, lists, tables (one row per line), tracked insertions kept and deletions dropped, content controls. |
| `.doc` (Word 97-2003) | The binary format through its piece table: text, tables, field results (not field codes), Word 6/95 fallback. |
| `.rtf` | Own parser: text, `\'hh` and `\uN` escapes, tables, with font/colour/style tables and pictures skipped. |
| `.odt` `.ods` `.odp` `.fodt` `.fodp` | OpenDocument content.xml: headings -> Contents, lists, table rows, slides. |
| `.pptx` `.pptm` `.potx` `.ppsx` | Slides in presentation order, each titled from its title placeholder, body text as paragraphs. |
| `.ppt` (PowerPoint 97-2003) | Text atoms out of the binary stream, grouped by slide. |
| `.xlsx` `.xlsm` `.xltx` | Visible sheets, one line per row, cells separated by `\|`; shared and inline strings, numbers, booleans. |
| `.xls` (Excel 97-2003) | BIFF8: shared strings, numbers, RK/MULRK, labels, cached formula results. |
| `.csv` `.tsv` | One line per row; the delimiter is sniffed. |
| `.epub` | Spine order, headings -> Contents. |
| `.mobi` `.prc` `.azw` `.azw3` | PalmDOC or uncompressed text records -> HTML -> blocks. DRM'd books and HUFF/CDIC compression are reported, not read (convert with Calibre). |
| `.fb2` | Sections and titles -> Contents, paragraphs, verse. |
| `.eml` | From / To / Date / Subject lines, then the plain-text body (or the HTML body when that is all there is), attachment names listed. |
| `.msg` (Outlook) | Subject, from, to, cc and body out of the property streams; attachment names listed. |
| `.srt` `.vtt` `.sbv` | Cue numbers, timings and tags dropped; cues joined back into paragraphs. |
| `.html` `.htm` `.xhtml` | Block structure, scripts and styles skipped. |
| `.md` `.markdown` | Headings, lists, quotes, code fences; inline marks stripped. |
| `.txt` and anything else that is text | Paragraphs by blank line; UTF-8, UTF-16 or Windows-1252. Code and data files (`.json` `.xml` `.yaml` `.ini` `.log` `.py` `.js` `.sh` ...) are shown line by line in a monospace face. |
| Everything else | If LibreOffice is installed it converts the file first (WordPerfect, Works, StarOffice, old templates...); otherwise Lectern says so. |

## Layout

| File | Purpose |
|---|---|
| `lectern.py` | Core: Document model, PDF / DOCX / EPUB / HTML / Markdown / text loaders, PDF word boxes and page rendering, settings, recent files, self-test |
| `formats.py` | The rest of the formats: a CFB/OLE reader for .doc .xls .ppt .msg, RTF, OpenDocument, .pptx/.xlsx, CSV, e-mail, subtitles, MOBI, FB2, magic-byte sniffing, LibreOffice fallback |
| `lectern_gui.py` | The window: contents pane, text view, page view, find |
| `lectern_app.py` | Entry point (`selftest` sub-mode, crash log) |
| `readaloud.py` | Shared read-aloud module (identical copy in Hayloft): tokenizer, voice engines, word clock, highlighter, RSVP panel, control strip |
| `theme.py` | House theme (Phosphor / Dark / Light, Ctrl+D cycles) |
| `build_flags.py` | Build helper: PyInstaller flags for whichever optional packages installed |
| `ensure_python.ps1` | Windows helper: find or install a real Python |

Repos stay flat: the extension says which OS a file is for.
