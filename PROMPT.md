# PROMPT — Build "PDF → Excel Converter & Merger" Desktop App + Windows Download Page

> This file is the exact spec used to build the app in this repository. Paste
> it into any AI assistant to regenerate the project from scratch.

---

You are a senior software engineer. Build a complete, production-ready desktop application plus a static web landing page, exactly to the specification below. Do not simplify, omit, or "improve" any listed behavior — the spec is the contract. All processing must happen locally on the user's machine; no cloud service.

## Part 1 — Product overview

Build two deliverables:

### 1a. Desktop app (primary)
A cross-platform native desktop application (macOS / Windows / Linux) that converts **multiple PDF files into Excel workbooks with zero data loss**, then lets the user **merge** the results into a single `.xlsx` and **extract** a standard set of columns into one consolidated workbook.

- **Tech stack (fixed):** Python 3.10+, Tkinter for the GUI, [PyMuPDF](https://pymupdf.readthedocs.io/) for PDF parsing (import as `import pymupdf as fitz`, with a `import fitz` fallback), [openpyxl](https://openpyxl.readthedocs.io/) for Excel writing, Pillow for image handling, and `tkinterdnd2` for drag & drop (**optional** — the app must run without it, falling back to a plain `tk.Tk()` root and a status note).
- Threading model: a background worker thread per job communicating with the UI through a `queue.Queue`, drained by `root.after(80, poll)`.
- Dark theme throughout the UI.
- App version string `1.0.0` in `pdf2xlsx/__init__.py`; entry point `main.py`.

### 1b. Web landing page (secondary)
A **static** landing page (`index.html` + one CSS file, optionally a tiny JS file; must be hostable as-is on GitHub Pages, Netlify, or any static host). It presents the app (name, tagline, feature list, screenshots/placeholders, system requirements) and prominently features a **"Download for Windows"** button that links to a Windows `.exe` installer artifact (see Part 7). No backend, no build step required to host.

## Part 2 — Complete UI specification

### Window
- Title: `PDF → Excel Converter & Merger`
- Geometry `980x780`; `minsize(820, 700)`.
- Header: title label `PDF → Excel Converter & Merger` (style `Title.TLabel`, Helvetica Neue 15 bold), subtitle `Convert PDFs to Excel with zero data loss, then merge the results.` (style `Dim.TLabel`).

### Toolbar (left to right)
- `Select PDFs...` (style `Accent.TButton`)
- `Output Folder...`
- `Remove Selected`
- `Clear List`
- `Cancel` (style `Danger.TButton`, packed right)

### Output-folder status label
Initial text: `Output folder: not set — you will be asked when converting.` After choosing: `Output folder: {path}`.

### Input file list (ttk.Treeview, `show="headings"`, `selectmode="extended"`, `height=5`)
- Columns and exact headings: `file` → `PDF file` (width 430, left-aligned), `size` → `Size` (90, center), `pages` → `Pages` (60, center), `status` → `Status` (150, center), `progress` → `Progress` (140, center).
- Vertical scrollbar. Tree item `iid` = absolute file path.
- Row tags → foreground colors: `pending`→default, `converting`→warning, `done`→success, `failed`→danger, `cancelled`→dim, `need_password`→warning, `invalid`→danger.
- `Pages` column computed **at add time** (`pymupdf.open(path).page_count`); show `—` if unreadable.
- `<<TreeviewSelect>>` triggers button-state refresh.

### File-adding rules
- Only `.pdf` files (case-insensitive) are accepted; skip names starting with `._` (macOS AppleDouble); skip exact-path duplicates.
- Entry fields: `path, size, pages, state, message, result`.
- If a file cannot be opened as a PDF: state `invalid`, status `Invalid PDF`, message `Cannot open file as a PDF`.
- If `output_dir` is not yet set, the first added PDF sets it to `dirname(first_pdf)/Converted`.

### Convert row
- Button `Convert All to Excel` (style `Accent.TButton`).
- `ttk.Progressbar` (determinate) beside it, then an overall label formatted `{completed} / {total}` (e.g. `2 / 5`).
- Clicking Convert re-runs only entries in states `PENDING`, `FAILED`, `NEED_PASSWORD` (done files are skipped). If none qualify: `showinfo("Nothing to do", "Add PDF files or reset failed entries first.")`. If no output folder: prompt `askdirectory` titled `Choose output folder`, else `os.makedirs(output_dir, exist_ok=True)`.
- On start: clear stored passwords, clear cancel event, reset completed counter, progress max = number of targets.

### Results list (ttk.Treeview, `height=4`)
- Section header label `Converted Excel files` (style `Panel.TLabel`).
- Columns/headings: `file` → `Excel file` (460, left), `size` → `Size` (90, center), `status` → `Status` (150, center). Vertical scrollbar, `selectmode="extended"`.
- Status text `Done`; `Extracted` or `Merged` for outputs produced by those tools. The merged and extracted files are themselves added to this list (so they can be saved, re-merged, or extracted again).

### Merge row
- Radio `Merge as separate sheets` (value `sheets`, default).
- Radio `Auto-append matching sheets` (value `append`).
- Check `Merge automatically after conversion`.
- Button `Merge Into One Excel File` (style `Accent.TButton`, packed right).

### Extraction panel
- Section label `Extract columns from converted files:` (style `Panel.TLabel`).
- Five checkbuttons, all **checked by default**, spaced `padx=(0, 12)`:
  `Material Code`, `Item Description`, `EAN No.`, `Quantity`, `Unit Base Cost`.
- Button `Extract Selected Fields` (style `Accent.TButton`, packed right).
- If none checked: `showinfo("No fields selected", "Tick at least one checkbox (Material Code, Item Description, EAN No., Quantity, Unit Base Cost).")`.

### Save row
- Buttons: `Save All to Output Folder`, `Save Selected As...`, `Open Output Folder`.
- Status label (packed right), initial `Ready.`. If drag & drop unavailable: `Ready. (Drag & drop unavailable — use Select PDFs.)`.

### Keyboard shortcuts (bind on both platforms)
| Action | macOS | Windows/Linux |
|---|---|---|
| Select PDFs | `⌘O` / `<Command-o>` | `Ctrl+O` / `<Control-o>` |
| Choose output folder | `⌘F` | `Ctrl+F` |
| Merge | `⌘M` | `Ctrl+M` |
| Extract fields | `⌘E` | `Ctrl+E` |
| Save All | `⌘S` | `Ctrl+S` |
| Remove selected | `⌫` / `<BackSpace>` | `Delete` |
| Cancel job | `Esc` | `Esc` |

### Button enable/disable rules
- Convert/Merge/Extract/Save All/Save Selected: disabled while a worker is alive, and while the results list is empty.
- Open Output Folder: disabled only while busy.
- Remove Selected / Clear List / Convert All silently no-op while busy.

### Status/progress texts (exact)
- Input progress cell: `"{pct}%  ({cur}/{tot})"` with two spaces before `(`; `pct = int(cur*100/tot)`.
- Overall: `"0 / {n}"` → `"{done} / {n}"` during convert; `"Merge {cur}/{tot}"` during merge; `"Extract {cur}/{tot}"` during extract.
- Status label: `"Added {n} PDF file(s)."`, `"Output folder changed."`, `"Converted: {basename}"`, `"Merging {scope}…"`, `"Extracting fields…"`, `"Cancelling…"`, `"Extraction complete."`, `"Merge complete."`, `"Saved to {target}"`.

### Message dialogs (exact titles/texts)
- Busy: `showinfo("Busy", "Wait for the current job to finish first.")` (when adding files mid-job).
- Clear list: `askyesno("Clear list", "Remove all files from the list?")`.
- Merge/extract with no results: `showinfo("No files", "Convert some PDFs first.")`.
- Missing selection files: `showinfo("No valid selection", "The selected files are missing.")`.
- Save complete: `showinfo("Save complete", f"Saved {len(saved)} file(s) to:\n{self.output_dir}")`.
- Nothing selected: `showinfo("Nothing selected", "Select a converted file in the list.")`.
- Missing file: `showerror("Missing file", f"File not found:\n{src}")`.
- No output folder: `showinfo("No output folder", "Choose an output folder first.")`.
- Open failed: `showerror("Error", "Could not open the folder.")`.
- Merge complete: `showinfo("Merge complete", "Merged file created:\n{path}" + notes…)`.
- Merge failed: `showerror("Merge failed", err)`.
- Extraction complete: `showinfo("Extraction complete", f"Extracted {rows} row(s)\nColumns: {columns}\n\n{path}" + notes…)`.
- Extraction failed: `showerror("Extraction failed", err)`.

### Password dialog
- `Toplevel`, title `PDF is password protected`, `resizable(False, False)`, `transient(root)`, modal (`grab_set` + `wait_window`).
- Message: `"'{basename}' is encrypted.\nEnter the password to continue:"`.
- `ttk.Entry(show="*", width=40)` auto-focused; Enter = unlock, Escape = skip.
- Buttons `Unlock` (Accent) and `Skip`.
- Worker blocks on a `threading.Event` while the dialog is shown (modal to the main window, runs on the UI thread).

## Part 3 — Conversion engine specification

`convert_pdf_to_excel(pdf_path, out_dir, password=None, cancel_event=None, progress_cb=None)` → `{"path", "sheets", "warnings"}`. Raises `EncryptedPDF`, `ConversionCancelled`, `ConversionError`.

### Constants
- `CONTINUATION_MARGIN = 60.0` (points from page bottom for continuation detection)
- `X_TOLERANCE = 14.0` (column-alignment tolerance in points)
- `MAX_IMG_WIDTH_PX = 620.0`
- Warning color hex `E05656`

### Table detection
- Ruled-line strategy only: `page.find_tables().tables` (PyMuPDF). Borderless (whitespace-separated) tables are **not** auto-detected; their text is preserved on the `Text & Images` sheet instead. A `text` alignment strategy was tested and deliberately rejected (false positives on prose) — do not reintroduce it.
- Table rows from `table.extract()`, cleaned: drop fully-empty rows, `None`→`""`, pad short rows to max column count.

### Cross-page table continuation
A table continues the previous table iff **all** of:
1. Previous table `ended_at_bottom` (`bbox.y1 >= page_height - 60.0`).
2. Column counts match.
3. Left and right edges within `X_TOLERANCE` (14 pt).

On continuation: append to the **same worksheet** (keeping the original `Page N – Table M` name, no second sheet); drop a repeated header row if the first continuation row equals the original header; re-run column autofit afterwards. Continuation rows are **not** written to the text sheet and no pointer note is repeated. The table counter is **document-wide**, incrementing only for genuinely new tables.

### Per-page worksheets
Every page gets a `Page {N} – Text & Images` sheet (N = 1-based):
- `A1` = `Page {N}` (bold, 13 pt); column A width preset to 110.
- A single ordered event stream of text blocks, tables, and images sorted by `(top-y, kind)` with kind order at equal y: `image` < `table` < `text`.
- Text blocks whose center falls inside a detected table bbox are skipped (already captured); same for images inside tables.
- Free text: one cell per block, 10 pt, `wrap_text=True`, `vertical="top"`, row height auto-fit afterwards.

### Table worksheet styling
- Title `Page {N} – Table {run_id}` (validated/deduped).
- Header row: fill `1F4E78`, white bold 11 pt, centered.
- Body: 10 pt, thin borders `8A8F98`; `freeze_panes = "A2"`; `autofit_columns` (min 8, max 60, width = longest line + 2).
- A pointer note on the text sheet at the table's reading position: `→ Table {run_id} is on the '{title}' worksheet.` (italic 9 pt, color `4C9AFF`). Not repeated on continuation pages.

### Numeric cell typing
- `^[+-]?\d+$` → `int`; `^[+-]?\d+\.\d+(?:[eE][+-]?\d+)?$` → `float`; anything else stays a **string** — so `0100`, `0850`, `4002355021235` remain text; `0.02`→`0.02`, `3400.50`→`3400.5`; empty→`None`.
- Body cells: `wrap=True` only for strings containing `\n`.

### Images
- `page.get_image_info(xrefs=True)`; `doc.extract_image(xref)` → blob → temp file → `openpyxl.drawing.image.Image` anchored at `A{row}` on the text sheet; width capped at 620 px (`scale = min(1.0, 620/img_w)`).
- Caption cell: `[Image embedded – {width}x{height} px]` (italic 9 pt, color `888888`).
- Embed failure → appended to `warnings`.

### Scanned / image-only pages
- If a page has no text blocks and no tables but has images: embed pictures plus an italic warning note (9 pt, color `E05656`):
  `[This page appears to be a scanned image – the picture was embedded as-is. No OCR is bundled, so the page has no searchable text. Run OCR on the PDF first if you need it.]`
- Also append warning: `Page {N} contains only images (scanned page); no text extracted – OCR not enabled.`

### Encrypted PDFs
- Open failure → `ConversionError("Cannot open PDF – file is corrupt or not a PDF: {exc}")`.
- `doc.is_encrypted` and no password → raise `EncryptedPDF` (UI shows `Password required`, then the password dialog; retry loop on a `threading.Event`).
- Wrong password → `ConversionError("Incorrect password – the PDF could not be unlocked.")`.
- Skip → state `failed` with `Password not provided.`.

### Streaming, cancel, atomic writes
- Pages processed in a loop; `cancel_event` checked per page → `ConversionCancelled`.
- `progress_cb(pno, page_count, "Page {pno+1} of {page_count}")`.
- Zero pages → `ConversionError("The PDF contains no pages.")`.
- Output written to `out_path + ".part"` then `os.replace(part, out_path)`; `.part` cleaned up on all paths; workbook always closed.
- Output name: PDF base name + `.xlsx`, deduped by `unique_path` (never overwrite).

## Part 4 — Merge specification

`merge_excels(input_paths, out_path, mode="sheets", cancel_event=None, progress_cb=None)` → `{"path", "sheets", "notes"}`. Raises `MergeCancelled`. Atomic `.part` + `os.replace`.

### Scope (UI)
- If files are selected in the results list → merge **only selected** existing files (status `Merging {n} selected file(s)…`); otherwise all converted files that still exist (`Merging {n} file(s)…`).
- Output: `Merged_{YYYYmmdd-HHMMSS}.xlsx` in the output folder.

### Mode "sheets" (Merge as separate sheets)
- Every source worksheet becomes `{source-stem} – {sheet-title}` (en dash separator). Nothing dropped: values, styles, column widths, row heights, merged cells, freeze panes, auto-filter, and images.
- Sheet title sanitizer: `[]:*?/\` → `_`, strip, cap at 31 chars, dedup with ` (N)`; when prefixing, truncate the **prefix** (not the sheet name) to fit 31 chars.

### Mode "append" (Auto-append matching sheets)
- Group sheets across all workbooks by normalized title (`re.sub(r"[^0-9a-z]+", "", name.lower())`).
- First sheet of each group copied fully; later sheets appended **only if schemas match** (equal `max_column`, both nonzero, identical row-1 values string-stripped).
- Match → append from row 2 (skipping the repeated header), preserving values, styles, merged ranges (offset), and images (offset by appended rows).
- Mismatch → kept as its own prefixed sheet; note `"'{stem} – {sheet}' had a different schema; kept as its own sheet '{alt}'."`
- Append note: `"Appended {n} row(s) into '{title}' from '{stem} – {sheet}'."`
- Progress: `progress_cb(processed, len(groups))`.

### Image preservation across merges
openpyxl cannot read pictures back from existing workbooks, so images are pulled directly from source `.xlsx` **zip packages**:
- Parse `xl/workbook.xml` → sheet name → relationship id; `xl/_rels/workbook.xml.rels` (fallback `xl/workbook.xml.rels`) → worksheet part paths.
- For each sheet, follow the drawing relationship → `xl/drawings/drawingN.xml`, then drawing rels → media targets.
- Handle `oneCellAnchor` anchors only; extract `from/col`, `from/row` (zero-based), `ext cx/cy` (EMU → points via `/9525`), and the image blob via `blip embed`.
- Record `{data, ext, col, row, cx, cy}`; unknown/corrupt packages silently yield `{}`.
- Re-embed via BytesIO → `XLImage` → `AnchorMarker` (zero offsets) → `XDRPositiveSize2D`. Do **not** use `XDRPoint2D` (throws `TypeError`).

## Part 5 — Extraction specification

`extract_fields(input_paths, out_path, fields, cancel_event=None, progress_cb=None)` → `{"path", "rows", "columns", "notes"}`. Raises `ExtractionCancelled`; empty `fields` → `ValueError("Select at least one field to extract.")`.

### The five fields (order = output column order)
1. `material_code` → `Material Code`
2. `item_description` → `Item Description`
3. `ean` → `EAN No.`
4. `quantity` → `Quantity`
5. `unit_cost` → `Unit Base Cost`

### Tolerant header matching
- `_norm`: lowercase + reduce non-alphanumerics to single spaces.
- Match = exact synonym equality, else one-way containment when the synonym is ≥4 normalized chars.
- Synonym tables:
  - `material_code`: material code, material, material no, item code, item no, product code, article no, artikel code, sku, ref no, reference, matcode, code of material
  - `item_description`: item description, item desc, description, description of goods, description of item, description of material, product description, desc, item
  - `ean`: ean, ean no, ean number, ean-13, ean13, barcode, bar code, gtin, gtin no, upc, upc no, barcode ean
  - `quantity`: quantity, qty, qty no, qty pcs, qty(pcs), no of pcs, no of pieces, number of pieces, pieces, quantity of items, total qty
  - `unit_cost`: unit base cost, unit cost, base cost, unit base price, unit price, base price, cost, price, unit base
- `HEADER_LIMIT = 60` rows scanned; the header row matching the most requested fields wins (first on tie); early exit when all fields found.
- Each field maps to the first column whose cell matches.

### Sheet selection heuristic
All matching sheets collected; **if any matched sheet title contains `table` (case-insensitive), only table sheets are used** (prose sheets that merely mention column names in a sentence are ignored).

### Data extraction
- Rows read from `header_row + 1` down; the **first fully-empty row ends the data block**.
- `ean` and `material_code` are forced to **text** (`STRING_FIELDS`; float with integer value → `str(int(...))`) so leading zeros and 13-digit codes survive.
- Missing columns → per-sheet note: `"'{basename}' → '{sheet}': column '{label}' not found"`.
- Per-file notes: `"{basename}: could not be read ({exc})"`, `"{basename}: no matching table found"`.

### Output workbook
- Single sheet `Extracted`; header row styled (dark blue fill `1F4E78`, white bold), `freeze_panes = "A2"`, `autofit_columns`; rows appended in file order, fields in the user's chosen order.
- Output: `Extracted_Fields_{YYYYmmdd-HHMMSS}.xlsx`, atomic write; result dialog shows row count, columns, path, and first 12 notes.
- UI progress: `progress_cb(i+1, len(paths), basename)` per file.

## Part 6 — Output organization, save operations, utilities

- Default output folder: `dirname(first_added_pdf)/Converted`; overridable via `Output Folder...` (native `askdirectory`). If unset at Convert/Merge/Extract/Save time, prompt with `askdirectory`.
- `unique_path(path)`: `file.xlsx` → `file (1).xlsx` → … (never overwrites).
- Sheet naming: text sheet `Page {N} – Text & Images`; table sheet `Page {N} – Table {M}`; extraction `Extracted`; merge-separate `{stem} – {sheet}`.
- Save All: copy every existing converted file into the output folder (dedup via `unique_path`), skip missing files, then show `Save complete` and open the folder.
- Save Selected As: native `asksaveasfilename` (title `Save Excel file as…`, `defaultextension=".xlsx"`, filetypes `[("Excel workbook", "*.xlsx")]`).
- Open Output Folder / reveal: `open` (macOS), `os.startfile` (Windows), `xdg-open` (Linux).
- Styling helpers: header fill `1F4E78`, borders `8A8F98`, autofit min 8 / max 60, row-height autofit `max(15.0, line_count * 15.0)`.

## Part 7 — Windows web download page + distribution

### Landing page (static)
- `index.html` + `style.css` (a little vanilla JS allowed). Must work by opening the file locally and by hosting on GitHub Pages / Netlify / any static host with zero configuration.
- Sections: hero (app name + tagline + primary CTA), feature list, screenshots/placeholder images, system requirements (Windows 10/11 64-bit; no Python needed to run the app), FAQ-ish notes (all local, privacy, supported formats), footer with version + release date.
- **Prominent `Download for Windows` button** linking to the installer artifact (e.g. `downloads/PDF2Excel-Setup.exe` or `downloads/PDF2Excel-win64.exe`). Placeholder path documented in a README section so the built file can be dropped in.

### Desktop app Windows distribution
- **Packaging tool is the implementing AI's choice** (PyInstaller single-file, PyInstaller + NSIS, etc.), but these requirements are fixed:
  - Ships a **working Windows installer artifact** (a single-file `.exe` that runs without Python installed is acceptable; a proper `Setup.exe` installer with shortcuts is better).
  - Must be reproducible via a documented command or script (e.g. `scripts/build_windows.bat` / `build.sh`).
  - Include cross-platform build instructions in the README (macOS and Linux can run from source; Windows gets the packaged artifact).
  - The packaged app must show the normal Tkinter window with full functionality (conversion, merge, extraction, drag & drop), not a console-only stub.
- Note: the desktop app itself does **not** need an in-app "download Windows version" button; the download button lives on the web page.

## Part 8 — Verification (required test suite)

Provide `tests/make_samples.py` (generates sample PDFs with PyMuPDF + Pillow) and `tests/verify_samples.py` (headless end-to-end verification, no GUI). Sample PDFs:
- `report_tables.pdf` — paragraphs + a table split across two pages (tests continuation).
- `notes_text.pdf` — plain text, multi-paragraph, page break.
- `scan_image.pdf` — embedded image + paragraph + small ruled table.
- `confidential.pdf` — AES-256 encrypted, password `secret`.
- `catalog.pdf` — a ruled table with the exact headers `Material Code, Item Description, EAN No., Quantity, Unit Base Cost` and 4 data rows (first row `MC-1001 / Aluminium Screw 3mm / 4002355021235 / 1500 / 0.02`).

The suite must pass **all** of the following checks (29):

1. `report: output exists`
2. `report: has a Table worksheet`
3. `report: table continued across pages (single sheet)` — exactly 1 table sheet
4. `report: header preserved` (`Item`, `Amount` present)
5. `report: page-1 rows preserved` (`Widget A`, `Gadget C`)
6. `report: continuation rows preserved` (`Tool N`, `Gadget Q`)
7. `report: free text preserved` (sentence containing `report summarises product sales`)
8. `report: page-2 text sheet exists`
9. `notes: paragraph text preserved` (`Meeting Notes`)
10. `notes: multi-line action item preserved` (`Alice to share`)
11. `notes: page 2 present`
12. `scan: image embedded` (≥1 file in `xl/media/`)
13. `scan: text preserved` (cell containing `embedded as a picture`)
14. `scan: small table captured` (exactly 1 table sheet)
15. `encrypted: raises EncryptedPDF without password`
16. `encrypted: converts with correct password` (`secret`; contains `Confidential Monthly Summary`)
17. `unicode: output created` (from filename `München Bérlin report (final).pdf`)
18. `unicode: base name preserved`
19. `merge-sheets: all source sheets present` (≥7 sheets)
20. `merge-sheets: naming kept source prefix` (names start with `report` and `notes`)
21. `merge-sheets: images preserved`
22. `merge-append: matching tables appended into one sheet` (exactly 1 sheet with `Table 1`)
23. `merge-append: rows from both files present` (`Widget A` and `Gadget Q`)
24. `extract: output exists`
25. `extract: all four rows captured` (`rows == 4`)
26. `extract: all five columns mapped`
27. `extract: headers written in order` (`["Material Code", "Item Description", "EAN No.", "Quantity", "Unit Base Cost"]`)
28. `extract: first data row correct` (`["MC-1001", "Aluminium Screw 3mm", "4002355021235", 1500, 0.02]` — EAN stays text, quantity/cost numeric)
29. `extract: all EANs present` (all four 13-digit EANs)

Test hygiene: before running, wipe the output dir but **skip `._*` AppleDouble files**; also skip such files when listing samples.

## Part 9 — Constraints & quality bar

- **Zero data loss** is the core principle: values (including leading zeros), line breaks, reading order, and embedded images must survive conversion and merging.
- Borderless tables are preserved as text on the `Text & Images` sheet (documented limitation — never fabricate grids from prose).
- **No OCR** is bundled; scanned pages are embedded as images plus a clear notice. Never silently drop content.
- Python 3.10+; `tkinterdnd2` optional (marker `platform_system != "Linux"` in requirements); macOS AppleDouble files ignored.
- Keep the exact visual style: dark theme, `Accent.TButton`, `Danger.TButton`, `Title.TLabel`, `Dim.TLabel`, `Panel.TLabel`.
- Use unicode text as specified (`→`, en dash `–` in sheet names, `…` in some labels and ASCII `...` in others exactly as listed).
- Deliver clean, documented, idiomatic Python; run the full 29-check suite and report results. Include a README covering install, run, build-for-Windows, hosting the web page, and known limitations (borderless tables, no OCR, oneCellAnchor-only image merge, no charts/conditional-format copying across merges).
