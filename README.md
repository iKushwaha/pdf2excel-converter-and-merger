# PDF → Excel Converter & Merger

A native desktop application that converts **multiple PDF files into Excel
workbooks with zero data loss**, then lets you **merge** the results into a
single `.xlsx` and **save** everything to your local machine.

Built with Python + Tkinter, [PyMuPDF](https://pymupdf.readthedocs.io/)
(PDF parsing) and [openpyxl](https://openpyxl.readthedocs.io/) (Excel
writing).

Two ways to use it:

- **Desktop app** (100% local) — a native Tkinter app; everything runs on
  your machine, no network calls.
- **Web app** (self-hosted) — the same conversion engine behind a small
  Flask server; PDFs are processed **on the server that runs it**, and the
  results are downloaded back through the browser.

---

## Features

### 1. Multi-file selection
- **Select PDFs…** button (or press `⌘O`).
- **Drag & drop** PDFs straight onto the window (requires `tkinterdnd2`).
- File list shows name, size, page count and live per-file status/progress.
- Remove individual files (`⌫`/`Delete` on the selection) or clear the whole list.

### 2. PDF → Excel conversion (zero data loss)
Each selected PDF becomes its own `.xlsx` (same base name, auto-suffixed
`file (1).xlsx` if the name already exists). Per page:

- **Tables** are extracted accurately and written to their own worksheets,
  named `Page N – Table M`, with bold headers, frozen header row,
  auto-fitted column widths, borders and numeric cells (values like `0100`
  stay text so nothing is lost).
- **Tables spanning multiple pages** are detected and re-joined into one
  worksheet (matching column count and alignment; repeated header rows on the
  continuation page are skipped).
- **Free text** (paragraphs, headers, footers, page breaks) is preserved in
  reading order on `Page N – Text & Images` worksheets, keeping line breaks.
- **Embedded images** are extracted and embedded as pictures with a caption.
- **Scanned / image-only pages** are embedded as images and flagged with a
  clear notice that OCR is not bundled — content is never silently dropped.
- **Password-protected PDFs** trigger a password dialog; wrong passwords are
  reported instead of crashing.
- Large documents are processed **page-by-page** (streaming), with per-file
  and overall progress bars and a **Cancel** button.

### 3. Merge converted Excel files
- **Merge Into One Excel File** (`⌘M`) combines every converted workbook —
  or only the **selected** files in the results list.
- Two modes:
  - *Merge as separate sheets* — every source sheet becomes its own sheet
    (`source – sheet`), preserving formatting, merges, widths **and images**.
  - *Auto-append matching sheets* — sheets with the same schema (same columns
    and header) are stacked vertically; incompatible ones are kept as their
    own sheets. Append actions are reported in the summary dialog.
- Tick **Merge automatically after conversion** to skip the extra click and
  merge as soon as the conversion finishes.

### 4. Extract selected fields
Pull a standard set of columns out of the converted workbooks and consolidate
them into one spreadsheet:
- Checkboxes for **Material Code**, **Item Description**, **EAN No.**,
  **Quantity** and **Unit Base Cost** (any subset works; `⌘E`).
- **Extract Selected Fields** scans every table sheet, matches the headers
  tolerantly (synonyms like *Qty*, *Item Desc*, *Unit Base Price* are found),
  and writes one `Extracted_Fields_….xlsx` with a bold header and auto-fitted
  columns.
- EAN codes and material codes are kept as **text** so leading zeros and full
  13-digit codes survive intact; notes report any sheet whose columns were
  not fully matched.

### 4. Save / Download locally
- Conversion writes directly into the chosen **output folder** (default:
  a `Converted` folder next to your PDFs; change it with **Output Folder…**
  or `⌘F`).
- **Save All** (`⌘S`) exports every converted file (and the merged file) to
  the output folder.
- **Save Selected As…** exports a single file via the native save dialog.
- **Open Output Folder** reveals the results in your file manager.

### Shortcuts
| Action | macOS | Windows/Linux |
|---|---|---|
| Select PDFs | `⌘O` | `Ctrl+O` |
| Choose output folder | `⌘F` | `Ctrl+F` |
| Merge | `⌘M` | `Ctrl+M` |
| Extract fields | `⌘E` | `Ctrl+E` |
| Save All | `⌘S` | `Ctrl+S` |
| Remove selected | `⌫` | `Delete` |
| Cancel job | `Esc` | `Esc` |

---

## Install & run — desktop app

Requires **Python 3.10+**.

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

> On Linux you may also need the system Tk package, e.g.
> `sudo apt install python3-tk`. Drag & drop (`tkinterdnd2`) is optional —
> the app runs without it, you just use **Select PDFs…**.

---

## Web app (server-side, browser-based)

An alternative to the desktop app for users who want to convert in the
browser. It reuses the exact same conversion engine, so output is identical.
**Important:** files are processed on the machine running the server, not in
the client's browser.

> **Launch the app →**
> - **Live hosted site (GitHub Pages):** https://ikushwaha.github.io/pdf2excel-converter-and-merger/
> - **Local web app:** http://127.0.0.1:8000 (after starting the server below)

### Run the web app

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python -m pdf2xlsx.web.server          # or, after `pip install -e .`:
pdf2excel-web
```

Then open **[http://127.0.0.1:8000](http://127.0.0.1:8000)**.

Environment variables (all optional):

| Variable | Default | Purpose |
|---|---|---|
| `PDF2EXCEL_HOST` | `127.0.0.1` | bind address |
| `PDF2EXCEL_PORT` | `8000` | listen port |
| `PDF2EXCEL_DATA_DIR` | system temp `pdf2excel-web/` | where per-session uploads/outputs live |
| `PDF2EXCEL_MAX_UPLOAD` | `104857600` (100 MB) | per-upload size cap (bytes) |
| `PDF2EXCEL_SESSION_TTL` | `86400` | session lifetime in seconds |
| `PDF2EXCEL_SECRET_KEY` | dev key | Flask session signing key |
| `PDF2EXCEL_LANDING_URL` | GitHub Pages site | URL used by the header's **Landing Page** link |

### What the web UI does

- Drag & drop or browse for multiple PDFs (upload progress shown).
- Optional password field for encrypted PDFs.
- Converts each PDF to its own `.xlsx` with zero data loss (same engine).
- Lists every generated workbook with a **Download** link.
- **Merge** converts into one workbook — separate sheets or auto-append.
- **Extract** consolidates the standard columns (Material Code, Item
  Description, EAN No., Quantity, Unit Base Cost) into one workbook.
- **Reset session** deletes the current session's uploaded and generated
  files (files also expire automatically after `SESSION_TTL_SECONDS`).

### API reference

| Method | Path | Body / Query | Returns |
|---|---|---|---|
| `POST` | `/api/upload` | multipart `pdfs` (+ optional `password`) | per-file `{ok, name, sheets, warnings}` + `outputs` |
| `GET` | `/api/fields` | — | available extraction fields |
| `GET` | `/api/outputs` | — | generated workbooks |
| `GET` | `/api/download/<name>` | — | the `.xlsx` file |
| `POST` | `/api/merge` | JSON `{mode, files?}` | merged workbook + `outputs` |
| `POST` | `/api/extract` | JSON `{fields}` | extracted workbook + `outputs` |
| `POST` | `/api/reset` | — | clears current session |
| `GET` | `/health` | — | `{ok: true, name, version}` |

### Deploy the web app

Run it behind a reverse proxy for production (e.g. `nginx`/`caddy` proxying
to `127.0.0.1:8000`), set a real `PDF2EXCEL_SECRET_KEY`, and point
`PDF2EXCEL_DATA_DIR` at a persistent volume. A WSGI server such as gunicorn
(`gunicorn 'pdf2xlsx.web.server:create_app()'`) or waitress works for
serving it.

---

## Build for Windows (distribute a .exe)

The desktop app runs cross-platform from source; for Windows end-users you
package it into a standalone executable with [PyInstaller](https://pyinstaller.org/).

### Option A — single-file `.exe` (no Python required)

On a Windows machine (or VM), from the project root:

```bat
scripts\build_windows.bat
```

This installs dependencies + PyInstaller and produces:

```
dist\PDF2Excel.exe
```

A single-file, windowed executable that runs without Python installed.

### Option B — proper `Setup.exe` installer

1. Run `scripts\build_windows.bat` (or `scripts\build.sh`) to get
   `dist\PDF2Excel.exe`.
2. Wrap it with **NSIS** or **Inno Setup** to add a Start Menu shortcut,
   an icon and a guided install.

### Cross-platform builds

```bash
scripts/build.sh
```

Works on macOS (produces `dist/PDF2Excel.app`), Linux (`dist/PDF2Excel`) and
Windows Git Bash (`dist/PDF2Excel.exe`).

### Publish the installer on the web page

Copy the Windows build into the site and deploy the `web/` folder:

```bash
cp dist/PDF2Excel.exe web/downloads/PDF2Excel-win64.exe
```

The landing page's **Download for Windows** button already points at
`downloads/PDF2Excel-win64.exe`.

---

## Web landing page (download site)

**Live at:** https://ikushwaha.github.io/pdf2excel-converter-and-merger/

`web/` is a fully **static** site — no backend, no build step — that can be
hosted on GitHub Pages, Netlify, or any static host as-is.

```
web/
  index.html                 landing page (hero, features, screenshots, FAQ, CTA)
  style.css                  dark theme matching the desktop app
  assets/                    app-preview.svg, app icons (ico/png)
  downloads/                 put PDF2Excel-win64.exe here
```

Deploying:

- **GitHub Pages** — push the `web/` folder to a branch/pages source.
- **Netlify** — drag the `web/` folder into the Netlify drop zone.
- **Any static host** — upload `web/` as the site root.

Before publishing, replace `assets/app-preview.svg` with real screenshots.

---

## Project layout

```
main.py                 desktop entry point
src/pdf2xlsx/
  core/
    converter.py        PDF → Excel extraction (PyMuPDF)
    merger.py           merge workbooks + direct OOXML image extraction
    extractor.py        consolidate selected columns (5 catalogue fields)
    excelio.py          openpyxl helpers (styling, sheet copying)
    utils.py            paths, dedup, platform commands
  desktop/
    ui.py               Tkinter UI, threading, progress, drag & drop
    theme.py            dark ttk theme
    main.py             desktop entry (console script: pdf2excel-desktop)
  web/
    server.py           Flask app + routes (create_app factory)
    config.py           env-driven configuration
    services/           session storage + core orchestration
    static/             index.html, css/app.css, js/app.js
scripts/
  build_windows.bat     Windows PyInstaller build (single .exe)
  build.sh              cross-platform PyInstaller build (macOS/Linux/Windows)
web/
  index.html            static landing page with "Download for Windows"
  style.css             landing page stylesheet
  assets/               app-preview.svg + app icons
  downloads/            drop the Windows installer here for the download button
tests/
  make_samples.py       generates sample PDFs (tables, text, image, encrypted)
  verify_samples.py     headless end-to-end verification
  test_web.py           headless Flask web-app tests
pyproject.toml          packaging (src layout) — replaces requirements.txt
```

---

## How the output is organised

Given `report.pdf` with 3 pages, 2 tables on page 1 and a table that runs
over pages 1–2, the workbook contains:

```
report.xlsx
├── Page 1 – Text & Images      free text + images in reading order
├── Page 1 – Table 1            table (rows continue onto page 2)
├── Page 2 – Text & Images
└── Page 3 – Text & Images
```

## Running the tests

```bash
.venv/bin/python tests/make_samples.py      # (re)create sample PDFs
.venv/bin/python tests/verify_samples.py    # convert, inspect, merge, assert
.venv/bin/python tests/test_web.py          # Flask web-app end-to-end checks
```

The suite verifies table extraction, cross-page table continuation, free-text
preservation, image embedding, encrypted-PDF handling, unicode file names,
both merge modes (including image preservation across merges), field
extraction of the five catalogue columns, and the web app's upload, download,
merge, extract and session lifecycle.

---

## Known limitations

- **Borderless tables** (no ruled lines, columns separated only by spaces)
  are not auto-detected. Their text is preserved as plain text on the
  `Text & Images` sheet instead, so nothing is lost, but they are not
  reconstructed as grid tables. The `text` table-detection strategy was
  tested and rejected because it produced many false positives on prose.
- **OCR is not bundled.** Scanned/image-only PDFs have their pictures
  embedded and are flagged with a notice; run OCR on the PDF first if you
  need searchable text.
- Table **cell-by-cell formatting** (fonts, colours inside cells) is
  normalised to a consistent style; the *values* are always preserved.
- Interactive annotations, form fields and JavaScript in PDFs are ignored.
- Merging preserves values, styles, merges, widths and images, but complex
  charts and conditional formatting rules are not copied between workbooks.
