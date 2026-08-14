# PDF → Excel Converter & Merger

A native desktop application that converts **multiple PDF files into Excel
workbooks with zero data loss**, then lets you **merge** the results into a
single `.xlsx` and **save** everything to your local machine.

Built with Python + Tkinter, [PyMuPDF](https://pymupdf.readthedocs.io/)
(PDF parsing) and [openpyxl](https://openpyxl.readthedocs.io/) (Excel
writing). No web service is involved — everything runs locally.

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

## Install & run

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
main.py                 entry point
pdf2xlsx/
  ui.py                 Tkinter UI, threading, progress, drag & drop
  converter.py          PDF → Excel extraction (PyMuPDF)
  merger.py             merge workbooks + direct OOXML image extraction
  extractor.py          consolidate selected columns (5 catalogue fields)
  excelio.py            openpyxl helpers (styling, sheet copying)
  utils.py              paths, dedup, platform commands
  theme.py              dark ttk theme
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
```

The suite verifies table extraction, cross-page table continuation, free-text
preservation, image embedding, encrypted-PDF handling, unicode file names,
both merge modes (including image preservation across merges), and field
extraction of the five catalogue columns.

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
