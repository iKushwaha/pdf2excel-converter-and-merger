"""Tkinter desktop UI for the PDF → Excel converter and merger."""

import os
import queue
import shutil
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from ..core import converter, extractor, merger, utils
from . import theme

# File states
PENDING = "pending"
CONVERTING = "converting"
DONE = "done"
FAILED = "failed"
CANCELLED = "cancelled"
NEED_PASSWORD = "need_password"
INVALID = "invalid"

# Fields offered by the extraction panel (order = output column order)
EXTRACT_FIELDS = ["material_code", "item_description", "ean", "quantity", "unit_cost"]
EXTRACT_LABELS = {
    "material_code": "Material Code",
    "item_description": "Item Description",
    "ean": "EAN No.",
    "quantity": "Quantity",
    "unit_cost": "Unit Base Cost",
}

STATUS_TEXT = {
    PENDING: "Pending",
    CONVERTING: "Converting",
    DONE: "Done",
    FAILED: "Failed",
    CANCELLED: "Cancelled",
    NEED_PASSWORD: "Password required",
    INVALID: "Invalid PDF",
}


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF → Excel Converter & Merger")
        self.root.geometry("980x780")
        self.root.minsize(820, 700)

        self.style = ttk.Style()
        theme.apply_theme(root, self.style)

        self._dnd = False
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD  # noqa: PLC0415
            self._dnd = isinstance(root, TkinterDnD.Tk)
            self._dnd_files = DND_FILES
        except Exception:
            pass

        self.inputs = []              # list[dict] for selected PDFs
        self.converted = []           # list[dict] for generated xlsx files
        self.output_dir = None
        self.passwords = {}
        self._completed = 0
        self._job_total = 0
        self._password_event = threading.Event()
        self._cancel = threading.Event()
        self._queue = queue.Queue()
        self._worker = None
        self.merged_path = None
        self.auto_merge = tk.BooleanVar(value=False)
        self.extract_vars = {key: tk.BooleanVar(value=True) for key in EXTRACT_FIELDS}

        self._build_ui()
        self._build_shortcuts()
        self.root.after(80, self._poll)
        if self._dnd:
            self.root.drop_target_register(self._dnd_files)
            self.root.dnd_bind("<<Drop>>", self._on_drop)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=12)
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PDF → Excel Converter & Merger",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Convert PDFs to Excel with zero data loss, then merge the results.",
                  style="Dim.TLabel").pack(anchor="w", pady=(0, 10))

        # --- toolbar ------------------------------------------------------
        bar = ttk.Frame(outer, style="Toolbar.TFrame")
        bar.pack(fill="x", pady=(0, 6))
        ttk.Button(bar, text="Select PDFs...", command=self.select_pdfs,
                   style="Accent.TButton").pack(side="left")
        ttk.Button(bar, text="Output Folder...", command=self.choose_output_dir).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Clear List", command=self.clear_list).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Cancel", command=self.cancel_work, style="Danger.TButton").pack(side="right")

        self.out_dir_label = ttk.Label(outer, text="Output folder: not set — you will be asked when converting.",
                                       style="Dim.TLabel")
        self.out_dir_label.pack(anchor="w", pady=(0, 4))

        # --- input list ---------------------------------------------------
        list_frame = ttk.Frame(outer)
        list_frame.pack(fill="x")
        columns = ("file", "size", "pages", "status", "progress")
        self.input_tree = ttk.Treeview(list_frame, columns=columns, show="headings",
                                       selectmode="extended", height=5)
        headings = {"file": "PDF file", "size": "Size", "pages": "Pages",
                    "status": "Status", "progress": "Progress"}
        widths = {"file": 430, "size": 90, "pages": 60, "status": 150, "progress": 140}
        for col in columns:
            self.input_tree.heading(col, text=headings[col])
            self.input_tree.column(col, width=widths[col], anchor="w" if col == "file" else "center")
        self.input_tree.tag_configure(PENDING, foreground=theme.FG)
        self.input_tree.tag_configure(CONVERTING, foreground=theme.WARNING)
        self.input_tree.tag_configure(DONE, foreground=theme.SUCCESS)
        self.input_tree.tag_configure(FAILED, foreground=theme.DANGER)
        self.input_tree.tag_configure(CANCELLED, foreground=theme.FG_DIM)
        self.input_tree.tag_configure(NEED_PASSWORD, foreground=theme.WARNING)
        self.input_tree.tag_configure(INVALID, foreground=theme.DANGER)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.input_tree.yview)
        self.input_tree.configure(yscrollcommand=vsb.set)
        self.input_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.input_tree.bind("<<TreeviewSelect>>", lambda e: self._refresh_buttons())

        # --- conversion controls ------------------------------------------
        conv = ttk.Frame(outer)
        conv.pack(fill="x", pady=(8, 4))
        self.convert_btn = ttk.Button(conv, text="Convert All to Excel",
                                      command=self.convert_all, style="Accent.TButton")
        self.convert_btn.pack(side="left")
        self.progress = ttk.Progressbar(conv, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=12)
        self.overall_label = ttk.Label(conv, text="", style="Dim.TLabel")
        self.overall_label.pack(side="right")

        # --- results ------------------------------------------------------
        ttk.Label(outer, text="Converted Excel files", style="Panel.TLabel").pack(anchor="w", pady=(12, 2))
        res_frame = ttk.Frame(outer)
        res_frame.pack(fill="both", expand=True)
        rcols = ("file", "size", "status")
        self.result_tree = ttk.Treeview(res_frame, columns=rcols, show="headings",
                                        selectmode="extended", height=4)
        for col, text, width in (("file", "Excel file", 460), ("size", "Size", 90), ("status", "Status", 150)):
            self.result_tree.heading(col, text=text)
            self.result_tree.column(col, width=width, anchor="w" if col == "file" else "center")
        rvsb = ttk.Scrollbar(res_frame, orient="vertical", command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=rvsb.set)
        self.result_tree.pack(side="left", fill="both", expand=True)
        rvsb.pack(side="right", fill="y")

        # --- merge + save controls -----------------------------------------
        controls = ttk.Frame(outer)
        controls.pack(fill="x", pady=(8, 0))
        self.merge_mode = tk.StringVar(value="sheets")
        ttk.Radiobutton(controls, text="Merge as separate sheets", value="sheets",
                        variable=self.merge_mode).pack(side="left")
        ttk.Radiobutton(controls, text="Auto-append matching sheets", value="append",
                        variable=self.merge_mode).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(controls, text="Merge automatically after conversion",
                        variable=self.auto_merge).pack(side="left", padx=(12, 0))
        self.merge_btn = ttk.Button(controls, text="Merge Into One Excel File",
                                    command=self.merge_files, style="Accent.TButton")
        self.merge_btn.pack(side="right")

        # --- extract selected fields ----------------------------------------
        extract_frame = ttk.Frame(outer)
        extract_frame.pack(fill="x", pady=(8, 0))
        ttk.Label(extract_frame, text="Extract columns from converted files:",
                  style="Panel.TLabel").pack(anchor="w")
        cb_row = ttk.Frame(extract_frame)
        cb_row.pack(anchor="w", pady=(3, 0))
        for key in EXTRACT_FIELDS:
            ttk.Checkbutton(cb_row, text=EXTRACT_LABELS[key],
                            variable=self.extract_vars[key]).pack(side="left", padx=(0, 12))
        self.extract_btn = ttk.Button(extract_frame, text="Extract Selected Fields",
                                      command=self.extract_fields_ui, style="Accent.TButton")
        self.extract_btn.pack(side="right")

        save_row = ttk.Frame(outer)
        save_row.pack(fill="x", pady=(6, 0))
        self.save_all_btn = ttk.Button(save_row, text="Save All to Output Folder", command=self.save_all)
        self.save_all_btn.pack(side="left")
        self.save_as_btn = ttk.Button(save_row, text="Save Selected As...", command=self.save_selected_as)
        self.save_as_btn.pack(side="left", padx=(6, 0))
        self.open_dir_btn = ttk.Button(save_row, text="Open Output Folder", command=self.open_output_dir)
        self.open_dir_btn.pack(side="left", padx=(6, 0))
        self.status_label = ttk.Label(save_row, text="Ready.", style="Dim.TLabel")
        self.status_label.pack(side="right")

        if not self._dnd:
            self.status_label.config(text="Ready. (Drag & drop unavailable — use Select PDFs.)")

        self._refresh_buttons()

    def _build_shortcuts(self):
        root = self.root
        root.bind_all("<Command-o>", lambda e: self.select_pdfs())
        root.bind_all("<Control-o>", lambda e: self.select_pdfs())
        root.bind_all("<Command-f>", lambda e: self.choose_output_dir())
        root.bind_all("<Control-f>", lambda e: self.choose_output_dir())
        root.bind_all("<Command-m>", lambda e: self.merge_files())
        root.bind_all("<Control-m>", lambda e: self.merge_files())
        root.bind_all("<Command-e>", lambda e: self.extract_fields_ui())
        root.bind_all("<Control-e>", lambda e: self.extract_fields_ui())
        root.bind_all("<Command-s>", lambda e: self.save_all())
        root.bind_all("<Control-s>", lambda e: self.save_all())
        root.bind_all("<Delete>", lambda e: self.remove_selected())
        root.bind_all("<BackSpace>", lambda e: self.remove_selected())
        root.bind_all("<Escape>", lambda e: self.cancel_work())

    # ------------------------------------------------------------ selection
    def select_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if paths:
            self.add_files(list(paths))

    def _on_drop(self, event):
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:
            return
        self.add_files([p for p in paths if os.path.isfile(p)])

    def add_files(self, paths):
        if self._worker is not None and self._worker.is_alive():
            messagebox.showinfo("Busy", "Wait for the current job to finish first.")
            return
        added = 0
        for path in paths:
            name = os.path.basename(path)
            if not path.lower().endswith(".pdf"):
                continue
            if name.startswith("._"):  # macOS AppleDouble metadata file
                continue
            if any(f["path"] == path for f in self.inputs):
                continue
            entry = {"path": path, "size": os.path.getsize(path), "pages": None,
                     "state": PENDING, "message": "", "result": None}
            try:
                import pymupdf  # noqa: PLC0415
                entry["pages"] = pymupdf.open(path).page_count
            except Exception:
                entry["state"] = INVALID
                entry["message"] = "Cannot open file as a PDF"
            self.inputs.append(entry)
            self.input_tree.insert("", "end", iid=path,
                                   values=(path, utils.format_size(entry["size"]),
                                           entry["pages"] if entry["pages"] is not None else "—",
                                           STATUS_TEXT[entry["state"]], ""),
                                   tags=(entry["state"],))
            added += 1
        if added == 0:
            return
        if self.output_dir is None:
            first_dir = os.path.dirname(self.inputs[0]["path"])
            self.output_dir = os.path.join(first_dir, "Converted")
            self._set_out_dir_label()
        self.status_label.config(text=f"Added {added} PDF file(s).")
        self._refresh_buttons()

    def remove_selected(self):
        if self._worker is not None and self._worker.is_alive():
            return
        sel = set(self.input_tree.selection())
        if not sel:
            return
        self.inputs = [f for f in self.inputs if f["path"] not in sel]
        for iid in sel:
            self.input_tree.delete(iid)
        self._refresh_buttons()

    def clear_list(self):
        if self._worker is not None and self._worker.is_alive():
            return
        if self.inputs and not messagebox.askyesno("Clear list", "Remove all files from the list?"):
            return
        self.input_tree.delete(*self.input_tree.get_children())
        self.inputs = []
        self._refresh_buttons()

    # ------------------------------------------------------------ output dir
    def choose_output_dir(self):
        initial = self.output_dir or os.path.expanduser("~")
        folder = filedialog.askdirectory(title="Choose output folder", initialdir=initial)
        if folder:
            self.output_dir = folder
            self._set_out_dir_label()
            self.status_label.config(text="Output folder changed.")
        self._refresh_buttons()

    def _set_out_dir_label(self):
        self.out_dir_label.config(text=f"Output folder: {self.output_dir}")

    # -------------------------------------------------------------- convert
    def convert_all(self):
        if self._worker is not None and self._worker.is_alive():
            return
        targets = [f for f in self.inputs if f["state"] in (PENDING, FAILED, NEED_PASSWORD)]
        if not targets:
            messagebox.showinfo("Nothing to do", "Add PDF files or reset failed entries first.")
            return
        if self.output_dir is None:
            folder = filedialog.askdirectory(title="Choose output folder")
            if not folder:
                return
            self.output_dir = folder
            self._set_out_dir_label()
        else:
            os.makedirs(self.output_dir, exist_ok=True)
        self.passwords.clear()
        self._password_event.clear()
        self._cancel.clear()
        self._completed = 0
        self._job_total = len(targets)
        self.progress.config(value=0, maximum=self._job_total)
        self.overall_label.config(text="0 / %d" % self._job_total)
        self._worker = threading.Thread(target=self._convert_worker, args=(targets,), daemon=True)
        self._worker.start()
        self._refresh_buttons()

    def _post(self, msg):
        self._queue.put(msg)

    def _convert_worker(self, targets):
        try:
            for entry in targets:
                if self._cancel.is_set():
                    break
                entry["state"] = CONVERTING
                self._post(("status", entry["path"], CONVERTING, ""))
                self._process_one(entry)
        finally:
            self._post(("finished",))

    def _process_one(self, entry):
        attempts = 0
        while True:
            if self._cancel.is_set():
                entry["state"] = CANCELLED
                self._post(("status", entry["path"], CANCELLED, ""))
                return
            try:
                password = self.passwords.get(entry["path"])
                result = converter.convert_pdf_to_excel(
                    entry["path"], self.output_dir, password=password,
                    cancel_event=self._cancel,
                    progress_cb=lambda cur, tot, msg, p=entry["path"]:
                        self._post(("progress", p, cur, tot, msg)))
                entry["state"] = DONE
                entry["result"] = result
                self._post(("done", entry["path"], result))
                return
            except converter.EncryptedPDF:
                entry["state"] = NEED_PASSWORD
                self._post(("need_password", entry["path"]))
                self._password_event.wait()
                self._password_event.clear()
                pw = self.passwords.get(entry["path"])
                if pw is None:
                    entry["state"] = FAILED
                    entry["message"] = "Password not provided."
                    self._post(("status", entry["path"], FAILED, "Password not provided"))
                    return
                attempts += 1
                if attempts > 3:
                    entry["state"] = FAILED
                    entry["message"] = "Incorrect password."
                    self._post(("status", entry["path"], FAILED, "Incorrect password"))
                    return
                continue
            except converter.ConversionCancelled:
                entry["state"] = CANCELLED
                self._post(("status", entry["path"], CANCELLED, ""))
                return
            except converter.ConversionError as exc:
                entry["state"] = FAILED
                entry["message"] = str(exc)
                self._post(("status", entry["path"], FAILED, str(exc)))
                return
            except Exception as exc:  # defensive
                entry["state"] = FAILED
                entry["message"] = f"Unexpected error: {exc}"
                self._post(("status", entry["path"], FAILED, entry["message"]))
                return

    # --------------------------------------------------------------- merging
    def merge_files(self):
        if self._worker is not None and self._worker.is_alive():
            return
        if not self.converted:
            messagebox.showinfo("No files", "Convert some PDFs first.")
            return
        if self.output_dir is None:
            folder = filedialog.askdirectory(title="Choose output folder for the merged file")
            if not folder:
                return
            self.output_dir = folder
            self._set_out_dir_label()
        else:
            os.makedirs(self.output_dir, exist_ok=True)

        # Merge the selected files when there is a selection, otherwise all.
        selected = set(self.result_tree.selection())
        if selected:
            paths = [f["path"] for f in self.converted
                     if f["path"] in selected and os.path.exists(f["path"])]
            if not paths:
                messagebox.showinfo("No valid selection",
                                    "The selected files are missing.")
                return
            scope = f"{len(paths)} selected file(s)"
        else:
            paths = [f["path"] for f in self.converted if os.path.exists(f["path"])]
            scope = f"{len(paths)} file(s)"

        import time  # noqa: PLC0415
        name = f"Merged_{time.strftime('%Y%m%d-%H%M%S')}.xlsx"
        out_path = os.path.join(self.output_dir, name)
        self._cancel.clear()
        self.merge_btn.config(state="disabled")
        self.status_label.config(text=f"Merging {scope}…")
        self._worker = threading.Thread(
            target=self._merge_worker, args=(paths, out_path, self.merge_mode.get()),
            daemon=True)
        self._worker.start()
        self._refresh_buttons()

    def _merge_worker(self, paths, out_path, mode):
        try:
            result = merger.merge_excels(
                paths, out_path, mode=mode, cancel_event=self._cancel,
                progress_cb=lambda cur, tot: self._post(("merge_progress", cur, tot)))
            self._post(("merge_done", result))
        except merger.MergeCancelled:
            self._post(("merge_cancelled",))
        except Exception as exc:
            self._post(("merge_failed", str(exc)))

    # ------------------------------------------------------------- extraction
    def extract_fields_ui(self):
        if self._worker is not None and self._worker.is_alive():
            return
        if not self.converted:
            messagebox.showinfo("No files", "Convert some PDFs first.")
            return
        fields = [key for key in EXTRACT_FIELDS if self.extract_vars[key].get()]
        if not fields:
            messagebox.showinfo("No fields selected",
                                "Tick at least one checkbox (Material Code, "
                                "Item Description, EAN No., Quantity, Unit Base Cost).")
            return
        if self.output_dir is None:
            folder = filedialog.askdirectory(title="Choose output folder for the extracted file")
            if not folder:
                return
            self.output_dir = folder
            self._set_out_dir_label()
        else:
            os.makedirs(self.output_dir, exist_ok=True)

        import time  # noqa: PLC0415
        name = f"Extracted_Fields_{time.strftime('%Y%m%d-%H%M%S')}.xlsx"
        out_path = os.path.join(self.output_dir, name)
        paths = [f["path"] for f in self.converted if os.path.exists(f["path"])]
        self._cancel.clear()
        self.extract_btn.config(state="disabled")
        self.status_label.config(text="Extracting fields…")
        self._worker = threading.Thread(
            target=self._extract_worker, args=(paths, out_path, fields), daemon=True)
        self._worker.start()
        self._refresh_buttons()

    def _extract_worker(self, paths, out_path, fields):
        try:
            result = extractor.extract_fields(
                paths, out_path, fields, cancel_event=self._cancel,
                progress_cb=lambda cur, tot, msg: self._post(("extract_progress", cur, tot, msg)))
            self._post(("extract_done", result))
        except extractor.ExtractionCancelled:
            self._post(("extract_cancelled",))
        except Exception as exc:
            self._post(("extract_failed", str(exc)))

    # ---------------------------------------------------------------- saving
    def save_all(self):
        if not self.converted:
            messagebox.showinfo("No files", "Nothing to save yet.")
            return
        if self.output_dir is None:
            folder = filedialog.askdirectory(title="Choose output folder")
            if not folder:
                return
            self.output_dir = folder
            self._set_out_dir_label()
        os.makedirs(self.output_dir, exist_ok=True)
        saved = []
        for f in self.converted:
            src = f["path"]
            if not os.path.exists(src):
                continue
            dst = os.path.join(self.output_dir, os.path.basename(src))
            if os.path.abspath(src) == os.path.abspath(dst):
                saved.append(dst)
                continue
            dst = utils.unique_path(dst)
            shutil.copy2(src, dst)
            saved.append(dst)
        messagebox.showinfo("Save complete",
                            f"Saved {len(saved)} file(s) to:\n{self.output_dir}")
        utils.open_folder(self.output_dir)

    def save_selected_as(self):
        sel = self.result_tree.selection()
        if not sel:
            messagebox.showinfo("Nothing selected", "Select a converted file in the list.")
            return
        src = sel[0]
        if not os.path.exists(src):
            messagebox.showerror("Missing file", f"File not found:\n{src}")
            return
        target = filedialog.asksaveasfilename(
            title="Save Excel file as…",
            initialdir=self.output_dir or os.path.expanduser("~"),
            initialfile=os.path.basename(src),
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")])
        if not target:
            return
        shutil.copy2(src, target)
        self.status_label.config(text=f"Saved to {target}")

    def open_output_dir(self):
        folder = self.output_dir
        if not folder or not os.path.isdir(folder):
            messagebox.showinfo("No output folder", "Choose an output folder first.")
            return
        if not utils.open_folder(folder):
            messagebox.showerror("Error", "Could not open the folder.")

    def cancel_work(self):
        if self._worker is not None and self._worker.is_alive():
            self._cancel.set()
            self.status_label.config(text="Cancelling…")

    # ------------------------------------------------------------ queue pump
    def _poll(self):
        try:
            while True:
                self._handle(self._queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _handle(self, msg):
        kind = msg[0]
        if kind == "progress":
            _, path, cur, tot, text = msg
            if self.input_tree.exists(path):
                pct = int(cur * 100 / tot) if tot else 0
                self.input_tree.set(path, "progress", f"{pct}%  ({cur}/{tot})")
        elif kind == "status":
            _, path, state, message = msg
            self._set_input_status(path, state, message)
            if state in (FAILED, CANCELLED):
                self._completed += 1
                self._update_overall()
        elif kind == "done":
            _, path, result = msg
            self._set_input_status(path, DONE, "")
            self._add_converted(result["path"])
            self._completed += 1
            self._update_overall()
            self.status_label.config(text=f"Converted: {os.path.basename(result['path'])}")
        elif kind == "need_password":
            _, path = msg
            self._prompt_password(path)
        elif kind == "finished":
            self._on_worker_finished("convert")
        elif kind == "extract_progress":
            _, cur, tot, msg = msg
            self.progress.config(value=cur, maximum=max(1, tot))
            self.overall_label.config(text=f"Extract {cur}/{tot}")
            self.status_label.config(text=f"Extracting from {msg}…")
        elif kind == "extract_done":
            _, result = msg
            self.progress.config(value=0, maximum=1)
            self.overall_label.config(text="")
            self._add_converted(result["path"], note="Extracted")
            info = (f"Extracted {result['rows']} row(s)\nColumns: "
                    f"{', '.join(result['columns'])}\n\n{result['path']}")
            notes = result.get("notes")
            if notes:
                info += "\n\n" + "\n".join(notes[:12])
            messagebox.showinfo("Extraction complete", info)
            self.status_label.config(text="Extraction complete.")
            self._on_worker_finished("extract")
        elif kind == "extract_cancelled":
            self.status_label.config(text="Extraction cancelled.")
            self._on_worker_finished("extract")
        elif kind == "extract_failed":
            _, err = msg
            messagebox.showerror("Extraction failed", err)
            self.status_label.config(text="Extraction failed.")
            self._on_worker_finished("extract")
        elif kind == "merge_progress":
            _, cur, tot = msg
            self.progress.config(value=cur, maximum=max(1, tot))
            self.overall_label.config(text=f"Merge {cur}/{tot}")
        elif kind == "merge_done":
            _, result = msg
            self.merged_path = result["path"]
            self.progress.config(value=0, maximum=1)
            self.overall_label.config(text="")
            self._add_converted(result["path"], note="Merged")
            notes = result.get("notes")
            info = f"Merged file created:\n{result['path']}"
            if notes:
                info += "\n\n" + "\n".join(notes[:12])
            messagebox.showinfo("Merge complete", info)
            self.status_label.config(text="Merge complete.")
            self._on_worker_finished("merge")
        elif kind == "merge_cancelled":
            self.status_label.config(text="Merge cancelled.")
            self._on_worker_finished("merge")
        elif kind == "merge_failed":
            _, err = msg
            messagebox.showerror("Merge failed", err)
            self.status_label.config(text="Merge failed.")
            self._on_worker_finished("merge")

    def _set_input_status(self, path, state, message):
        if not self.input_tree.exists(path):
            return
        self.input_tree.item(path, tags=(state,))
        text = message or STATUS_TEXT.get(state, state)
        self.input_tree.set(path, "status", text)
        for entry in self.inputs:
            if entry["path"] == path:
                entry["state"] = state
                if message:
                    entry["message"] = message

    def _add_converted(self, path, note=""):
        for f in self.converted:
            if f["path"] == path:
                return
        self.converted.append({"path": path, "status": note or "Done"})
        size = utils.format_size(os.path.getsize(path))
        self.result_tree.insert("", "end", iid=path, values=(path, size, note or "Done"))

    # ----------------------------------------------------------- password dlg
    def _prompt_password(self, path):
        dlg = tk.Toplevel(self.root)
        dlg.title("PDF is password protected")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.configure(bg=theme.BG)
        frame = ttk.Frame(dlg, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"'{os.path.basename(path)}' is encrypted.\nEnter the password to continue:",
                  style="Panel.TLabel").pack(anchor="w")
        entry = ttk.Entry(frame, show="*", width=40)
        entry.pack(fill="x", pady=(8, 12))
        entry.focus_set()
        result = {"pw": None}

        def ok():
            result["pw"] = entry.get()
            dlg.destroy()

        def cancel():
            dlg.destroy()

        btns = ttk.Frame(frame)
        btns.pack(fill="x")
        ttk.Button(btns, text="Unlock", command=ok, style="Accent.TButton").pack(side="right")
        ttk.Button(btns, text="Skip", command=cancel).pack(side="right", padx=(0, 6))
        entry.bind("<Return>", lambda e: ok())
        entry.bind("<Escape>", lambda e: cancel())
        dlg.grab_set()
        dlg.wait_window()
        self.passwords[path] = result["pw"]
        self._password_event.set()

    # ------------------------------------------------------------------ misc
    def _update_overall(self):
        if self._job_total:
            self.progress.config(value=self._completed)
            self.overall_label.config(text=f"{self._completed} / {self._job_total}")

    def _on_worker_finished(self, kind):
        self._worker = None
        self._cancel.clear()
        self._refresh_buttons()
        if kind == "convert" and self.auto_merge.get() and self.converted:
            self.root.after(120, self.merge_files)

    def _refresh_buttons(self):
        busy = self._worker is not None and self._worker.is_alive()
        has_result = bool(self.converted)
        self.convert_btn.config(state="disabled" if busy else "normal")
        self.merge_btn.config(state="disabled" if (busy or not has_result) else "normal")
        self.extract_btn.config(state="disabled" if (busy or not has_result) else "normal")
        self.save_all_btn.config(state="disabled" if (busy or not has_result) else "normal")
        self.save_as_btn.config(state="disabled" if (busy or not has_result) else "normal")
        self.open_dir_btn.config(state="disabled" if busy else "normal")
