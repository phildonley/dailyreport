#!/usr/bin/env python3
"""
PhotoPull — Batch image downloader driven by Excel / CSV spreadsheets.
Replaces partial URL patterns (e.g. media1:/products/…) with full URLs,
then downloads every image into a single flat output folder.
"""

import csv
import json
import os
import platform
import re
import subprocess
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext
import tkinter as tk
from tkinter import ttk

import requests

try:
    import openpyxl
    EXCEL_SUPPORT = True
except ImportError:
    EXCEL_SUPPORT = False

# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------

SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "photopull_settings.json"
)

DEFAULT_SETTINGS = {
    "output_folder": str(Path.home() / "Downloads" / "PhotoPull"),
    "url_patterns": [
        {
            "find": r"media\d+:/products/",
            "replace": (
                "https://myparts.terex.com/ccstore/v1/images/"
                "?source=/file/products/"
            ),
        }
    ],
    "separator": "|",
    "has_headers": True,
    "key_column": "",
    "image_column": "",
    "last_file": "",
}


# ---------------------------------------------------------------------------
# Rule editor dialog
# ---------------------------------------------------------------------------

class RuleDialog(tk.Toplevel):
    """Modal dialog for adding or editing a single URL replacement rule."""

    def __init__(self, parent, find="", replace=""):
        super().__init__(parent)
        self.result = None
        self.title("URL Replacement Rule")
        self.geometry("640x210")
        self.resizable(True, False)
        self.transient(parent)
        self.grab_set()
        self._build(find, replace)
        self.wait_window()

    def _build(self, find, replace):
        pad = {"padx": 12, "pady": 6}

        ttk.Label(self, text="Find pattern (supports regex):").grid(
            row=0, column=0, sticky=tk.W, **pad
        )
        self.find_var = tk.StringVar(value=find)
        ttk.Entry(self, textvariable=self.find_var, width=58).grid(
            row=0, column=1, sticky=tk.EW, **pad
        )

        ttk.Label(self, text="Replace with (URL prefix):").grid(
            row=1, column=0, sticky=tk.W, **pad
        )
        self.replace_var = tk.StringVar(value=replace)
        ttk.Entry(self, textvariable=self.replace_var, width=58).grid(
            row=1, column=1, sticky=tk.EW, **pad
        )

        ttk.Label(
            self,
            text=(
                "Tip:  \\d+  matches any number, so  media\\d+:/products/  covers "
                "media1:/products/,  media2:/products/,  media3:/products/ …"
            ),
            foreground="gray",
            wraplength=580,
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=12)

        bf = ttk.Frame(self)
        bf.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="Save", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)

        self.columnconfigure(1, weight=1)

    def _save(self):
        find = self.find_var.get().strip()
        replace = self.replace_var.get().strip()
        if not find or not replace:
            messagebox.showwarning("Required", "Both fields must be filled in.", parent=self)
            return
        self.result = (find, replace)
        self.destroy()


# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class PhotoPullApp:
    """Main application window."""

    def __init__(self, root):
        self.root = root
        self.root.title("PhotoPull — Image Downloader")
        self.root.geometry("920x700")
        self.root.minsize(740, 560)

        self.settings = {}
        self.columns = []
        self.data_rows = []
        self.stop_flag = False
        self._last_report_path = None

        self._load_settings()
        self._build_ui()
        self._apply_settings_to_ui()

    # ── Settings ──────────────────────────────────────────────────────────────

    def _load_settings(self):
        self.settings = {k: v for k, v in DEFAULT_SETTINGS.items()}
        # Deep-copy the list so defaults aren't mutated
        self.settings["url_patterns"] = [
            dict(r) for r in DEFAULT_SETTINGS["url_patterns"]
        ]
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, encoding="utf-8") as fh:
                    saved = json.load(fh)
                self.settings.update(saved)
            except Exception:
                pass

    def _save_settings(self, notify=False):
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
                json.dump(self.settings, fh, indent=2)
            if notify:
                messagebox.showinfo("Saved", "Settings saved successfully.")
        except Exception as exc:
            messagebox.showwarning("Save Error", f"Could not save settings:\n{exc}")

    def _collect_settings_from_ui(self):
        """Pull current UI values into self.settings."""
        self.settings["output_folder"] = self.output_var.get().strip()
        self.settings["separator"] = self.sep_var.get() or "|"
        self.settings["has_headers"] = self.has_headers_var.get()
        self.settings["key_column"] = self.key_col_var.get()
        self.settings["image_column"] = self.img_col_var.get()
        self.settings["last_file"] = self.file_var.get().strip()
        self.settings["url_patterns"] = [
            {
                "find": self.rules_tree.item(i, "values")[0],
                "replace": self.rules_tree.item(i, "values")[1],
            }
            for i in self.rules_tree.get_children()
        ]

    def _apply_settings_to_ui(self):
        self.output_var.set(self.settings.get("output_folder", ""))
        self.sep_var.set(self.settings.get("separator", "|"))
        self.has_headers_var.set(self.settings.get("has_headers", True))
        self._reload_rules_tree()
        last = self.settings.get("last_file", "")
        if last:
            self.file_var.set(last)

    # ── Top-level UI assembly ─────────────────────────────────────────────────

    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self.root, bg="#1a252f", pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(
            hdr, text="PhotoPull",
            font=("Helvetica", 20, "bold"), bg="#1a252f", fg="#ecf0f1",
        ).pack()
        tk.Label(
            hdr, text="Batch Image Downloader  •  Excel & CSV",
            font=("Helvetica", 9), bg="#1a252f", fg="#95a5a6",
        ).pack()

        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)

        self._tab_file()
        self._tab_settings()
        self._tab_run()

    # ── Tab 1 — File & Columns ────────────────────────────────────────────────

    def _tab_file(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 1 · File & Columns  ")

        ttk.Label(
            f,
            text=(
                "Select your Excel or CSV file, confirm whether it has a header row, "
                "click Load File, then choose which columns hold the part number and the image paths."
            ),
            foreground="gray", wraplength=860,
        ).pack(anchor=tk.W, pady=(0, 8))

        # ── File chooser ──────────────────────────────────────────────────────
        fg = ttk.LabelFrame(f, text="Input File", padding=10)
        fg.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(fg)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="File:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.file_var).pack(
            side=tk.LEFT, padx=6, fill=tk.X, expand=True
        )
        ttk.Button(row1, text="Browse…", command=self._browse_input).pack(side=tk.LEFT)

        row2 = ttk.Frame(fg)
        row2.pack(fill=tk.X, pady=(8, 0))
        self.has_headers_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            row2,
            text="First row contains column headers (column names)",
            variable=self.has_headers_var,
            command=self._on_headers_toggle,
        ).pack(side=tk.LEFT)
        ttk.Button(row2, text="Load File  ▸", command=self.load_file).pack(side=tk.RIGHT)

        # ── Column pickers ────────────────────────────────────────────────────
        cp = ttk.LabelFrame(f, text="Column Selection", padding=10)
        cp.pack(fill=tk.X, pady=(0, 8))

        for label_text, attr, hint in [
            ("Part Number / Key Column:", "key_col",
             "Optional — appears in the download report so you can trace errors by part number"),
            ("Image URL Column:", "img_col",
             "Required — the column containing image paths or partial URLs"),
        ]:
            r = ttk.Frame(cp)
            r.pack(fill=tk.X, pady=3)
            ttk.Label(r, text=label_text, width=30).pack(side=tk.LEFT)
            var = tk.StringVar()
            combo = ttk.Combobox(r, textvariable=var, state="readonly", width=34)
            combo.pack(side=tk.LEFT, padx=4)
            ttk.Label(r, text=f"← {hint}", foreground="gray").pack(side=tk.LEFT)
            setattr(self, f"{attr}_var", var)
            setattr(self, f"{attr}_combo", combo)

        # ── Preview ───────────────────────────────────────────────────────────
        pv = ttk.LabelFrame(f, text="Preview — first 5 rows", padding=5)
        pv.pack(fill=tk.BOTH, expand=True)

        self.preview_tree = ttk.Treeview(pv, show="headings", height=5)
        sx = ttk.Scrollbar(pv, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        sy = ttk.Scrollbar(pv, orient=tk.VERTICAL,   command=self.preview_tree.yview)
        self.preview_tree.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)
        sx.pack(side=tk.BOTTOM, fill=tk.X)
        sy.pack(side=tk.RIGHT,  fill=tk.Y)
        self.preview_tree.pack(fill=tk.BOTH, expand=True)

    # ── Tab 2 — Settings ──────────────────────────────────────────────────────

    def _tab_settings(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 2 · Settings  ")

        ttk.Label(
            f,
            text=(
                "Set where images are saved and how the partial paths in your spreadsheet "
                "are converted into full download URLs.  Click Save Settings when done."
            ),
            foreground="gray", wraplength=860,
        ).pack(anchor=tk.W, pady=(0, 8))

        # ── Output folder ─────────────────────────────────────────────────────
        of = ttk.LabelFrame(f, text="Output Folder", padding=10)
        of.pack(fill=tk.X, pady=(0, 8))

        row = ttk.Frame(of)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Save images to:").pack(side=tk.LEFT)
        self.output_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.output_var).pack(
            side=tk.LEFT, padx=6, fill=tk.X, expand=True
        )
        ttk.Button(row, text="Browse…", command=self._browse_output).pack(side=tk.LEFT)
        ttk.Label(
            of,
            text="All images go into this one folder — no sub-folders are created.",
            foreground="gray",
        ).pack(anchor=tk.W, pady=(5, 0))

        # ── Separator ─────────────────────────────────────────────────────────
        sf = ttk.LabelFrame(f, text="Multiple-Image Separator", padding=10)
        sf.pack(fill=tk.X, pady=(0, 8))

        row2 = ttk.Frame(sf)
        row2.pack(fill=tk.X)
        ttk.Label(row2, text="Character that separates multiple images in one cell:").pack(
            side=tk.LEFT
        )
        self.sep_var = tk.StringVar(value="|")
        ttk.Entry(row2, textvariable=self.sep_var, width=4).pack(side=tk.LEFT, padx=6)
        ttk.Label(
            row2,
            text="Default: |   Example cell:  img1.jpg | img2.jpg | img3.jpg",
            foreground="gray",
        ).pack(side=tk.LEFT)

        # ── URL replacement rules ─────────────────────────────────────────────
        rl = ttk.LabelFrame(f, text="URL Replacement Rules", padding=10)
        rl.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            rl,
            text=(
                "Each rule converts a partial path from the spreadsheet into a full download URL.\n"
                "The Find field is a regular expression.  "
                "Use \\d+ to match any digit sequence, so one rule covers media1, media2, media3 …"
            ),
            foreground="gray", justify=tk.LEFT, wraplength=840,
        ).pack(anchor=tk.W, pady=(0, 6))

        cols = ("find", "replace")
        self.rules_tree = ttk.Treeview(rl, columns=cols, show="headings", height=5)
        self.rules_tree.heading("find",    text="Find  (regex pattern)")
        self.rules_tree.heading("replace", text="Replace With  (full URL prefix)")
        self.rules_tree.column("find",    width=240)
        self.rules_tree.column("replace", width=520)

        rsy = ttk.Scrollbar(rl, orient=tk.VERTICAL, command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=rsy.set)
        rsy.pack(side=tk.RIGHT, fill=tk.Y)
        self.rules_tree.pack(fill=tk.BOTH, expand=True)

        btns = ttk.Frame(rl)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="Add Rule",       command=self._rule_add).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Edit Rule",      command=self._rule_edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Remove Rule",    command=self._rule_remove).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Reset to Defaults", command=self._rule_reset).pack(side=tk.LEFT, padx=14)
        ttk.Button(btns, text="Save Settings", command=self._save_settings_ui).pack(side=tk.RIGHT, padx=2)

    # ── Tab 3 — Run ───────────────────────────────────────────────────────────

    def _tab_run(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 3 · Run  ")

        ttk.Label(
            f,
            text=(
                "Review the summary below, then click Start Download.  "
                "Settings are saved automatically each time you start a download."
            ),
            foreground="gray", wraplength=860,
        ).pack(anchor=tk.W, pady=(0, 8))

        # ── Run Options ───────────────────────────────────────────────────────
        ro = ttk.LabelFrame(f, text="Run Options", padding=10)
        ro.pack(fill=tk.X, pady=(0, 8))

        opt_row = ttk.Frame(ro)
        opt_row.pack(fill=tk.X)
        self.test_run_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            opt_row,
            text="Test run — process only the first",
            variable=self.test_run_var,
        ).pack(side=tk.LEFT)
        self.test_limit_var = tk.IntVar(value=10)
        ttk.Spinbox(
            opt_row, from_=1, to=9999,
            textvariable=self.test_limit_var, width=6,
        ).pack(side=tk.LEFT, padx=4)
        ttk.Label(
            opt_row,
            text="rows   (verify your settings before committing to a full run)",
            foreground="gray",
        ).pack(side=tk.LEFT)

        # ── Summary ───────────────────────────────────────────────────────────
        sf = ttk.LabelFrame(f, text="Summary", padding=10)
        sf.pack(fill=tk.X, pady=(0, 8))
        self.summary_lbl = ttk.Label(
            sf,
            text="Load a file in Step 1 to see a summary here.",
            justify=tk.LEFT,
        )
        self.summary_lbl.pack(anchor=tk.W)

        # ── Progress ──────────────────────────────────────────────────────────
        pf = ttk.LabelFrame(f, text="Progress", padding=10)
        pf.pack(fill=tk.X, pady=(0, 8))
        self.prog_lbl = ttk.Label(pf, text="Idle — ready to start.")
        self.prog_lbl.pack(anchor=tk.W, pady=(0, 4))
        self.prog_bar = ttk.Progressbar(pf, mode="determinate")
        self.prog_bar.pack(fill=tk.X)
        self.prog_count = ttk.Label(pf, text="")
        self.prog_count.pack(anchor=tk.W, pady=(3, 0))

        # ── Control buttons ───────────────────────────────────────────────────
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(0, 8))
        self.start_btn = ttk.Button(bf, text="▶  Start Download", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn = ttk.Button(
            bf, text="◼  Stop", command=self._stop, state=tk.DISABLED
        )
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        self.open_btn = ttk.Button(
            bf, text="Open Output Folder", command=self._open_folder, state=tk.DISABLED
        )
        self.open_btn.pack(side=tk.RIGHT, padx=2)
        self.report_btn = ttk.Button(
            bf, text="View Report", command=self._view_report, state=tk.DISABLED
        )
        self.report_btn.pack(side=tk.RIGHT, padx=2)

        # ── Activity log ──────────────────────────────────────────────────────
        lf = ttk.LabelFrame(f, text="Activity Log", padding=5)
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_box = scrolledtext.ScrolledText(
            lf, height=14, state=tk.DISABLED,
            font=("Courier", 9), wrap=tk.WORD,
        )
        self.log_box.pack(fill=tk.BOTH, expand=True)

    # ── File loading ──────────────────────────────────────────────────────────

    def _browse_input(self):
        if EXCEL_SUPPORT:
            types = [
                ("Spreadsheets", "*.xlsx *.xls *.csv"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv"),
                ("All files", "*.*"),
            ]
        else:
            types = [("CSV", "*.csv"), ("All files", "*.*")]
        path = filedialog.askopenfilename(title="Select Input File", filetypes=types)
        if path:
            self.file_var.set(path)
            self.load_file()

    def _on_headers_toggle(self):
        if self.data_rows:
            self.load_file()

    def load_file(self):
        path = self.file_var.get().strip()
        if not path:
            messagebox.showwarning("No File", "Please select a file first.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Not Found", f"File not found:\n{path}")
            return
        try:
            ext = os.path.splitext(path)[1].lower()
            hdrs = self.has_headers_var.get()
            if ext in (".xlsx", ".xls"):
                if not EXCEL_SUPPORT:
                    messagebox.showerror(
                        "Missing Library",
                        "Install openpyxl to read Excel files:\n\n  pip install openpyxl",
                    )
                    return
                self.columns, self.data_rows = self._read_excel(path, hdrs)
            elif ext == ".csv":
                self.columns, self.data_rows = self._read_csv(path, hdrs)
            else:
                messagebox.showerror("Unsupported", "Please choose an Excel (.xlsx) or CSV file.")
                return
            self._refresh_columns()
            self._refresh_preview()
            self.settings["last_file"] = path
            self._update_summary()
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not load file:\n{exc}")

    @staticmethod
    def _read_excel(path, has_headers):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        if not rows:
            return [], []
        if has_headers:
            headers = [
                str(c) if c is not None else f"Col{i + 1}"
                for i, c in enumerate(rows[0])
            ]
            return headers, rows[1:]
        return [f"Column {i + 1}" for i in range(len(rows[0]))], rows

    @staticmethod
    def _read_csv(path, has_headers):
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        if not rows:
            return [], []
        if has_headers:
            return [c.strip() for c in rows[0]], rows[1:]
        return [f"Column {i + 1}" for i in range(len(rows[0]))], rows

    def _refresh_columns(self):
        saved_key = self.settings.get("key_column", "")
        saved_img = self.settings.get("image_column", "")
        for combo, var, saved in [
            (self.key_col_combo, self.key_col_var, saved_key),
            (self.img_col_combo, self.img_col_var, saved_img),
        ]:
            combo["values"] = self.columns
            var.set(
                saved if saved in self.columns
                else (self.columns[0] if self.columns else "")
            )

    def _refresh_preview(self):
        t = self.preview_tree
        t.delete(*t.get_children())
        t["columns"] = self.columns
        for col in self.columns:
            t.heading(col, text=col)
            t.column(col, width=max(80, min(220, len(col) * 9 + 24)))
        for row in self.data_rows[:5]:
            t.insert("", tk.END, values=[str(v) if v is not None else "" for v in row])

    # ── Settings-tab helpers ──────────────────────────────────────────────────

    def _browse_output(self):
        d = filedialog.askdirectory(title="Select Output Folder")
        if d:
            self.output_var.set(d)

    def _reload_rules_tree(self):
        self.rules_tree.delete(*self.rules_tree.get_children())
        for r in self.settings.get("url_patterns", []):
            self.rules_tree.insert("", tk.END, values=(r["find"], r["replace"]))

    def _rule_add(self):
        dlg = RuleDialog(self.root)
        if dlg.result:
            self.rules_tree.insert("", tk.END, values=dlg.result)

    def _rule_edit(self):
        sel = self.rules_tree.selection()
        if not sel:
            messagebox.showinfo("Edit Rule", "Select a rule first.")
            return
        vals = self.rules_tree.item(sel[0], "values")
        dlg = RuleDialog(self.root, find=vals[0], replace=vals[1])
        if dlg.result:
            self.rules_tree.item(sel[0], values=dlg.result)

    def _rule_remove(self):
        sel = self.rules_tree.selection()
        if not sel:
            messagebox.showinfo("Remove Rule", "Select a rule first.")
            return
        if messagebox.askyesno("Remove Rule", "Delete the selected rule?"):
            self.rules_tree.delete(sel[0])

    def _rule_reset(self):
        if messagebox.askyesno("Reset Rules", "Reset URL rules to factory defaults?"):
            self.settings["url_patterns"] = [
                dict(r) for r in DEFAULT_SETTINGS["url_patterns"]
            ]
            self._reload_rules_tree()

    def _save_settings_ui(self):
        self._collect_settings_from_ui()
        self._save_settings(notify=True)

    # ── Run-tab helpers ───────────────────────────────────────────────────────

    def _update_summary(self):
        if not self.data_rows:
            self.summary_lbl.config(text="No data loaded.")
            return
        img = self.img_col_var.get()
        if not img or img not in self.columns:
            self.summary_lbl.config(text="Select the image URL column in Step 1.")
            return
        idx = self.columns.index(img)
        sep = self.sep_var.get() or "|"
        total_img = sum(
            len([p for p in str(r[idx]).split(sep) if p.strip()])
            for r in self.data_rows
            if r[idx] is not None and str(r[idx]).strip()
        )
        total_rows = sum(
            1 for r in self.data_rows
            if r[idx] is not None and str(r[idx]).strip()
        )
        out = self.output_var.get() or "(not set — see Step 2)"
        self.summary_lbl.config(
            text=(
                f"File:                   {os.path.basename(self.file_var.get())}\n"
                f"Rows with images:       {total_rows} of {len(self.data_rows)}\n"
                f"Total images to fetch:  {total_img}\n"
                f"Output folder:          {out}"
            )
        )

    def _start(self):
        self._collect_settings_from_ui()

        img = self.settings.get("image_column")

        if not self.data_rows:
            messagebox.showwarning("No Data", "Load a file in Step 1 first.")
            return
        if not img:
            messagebox.showwarning("Column Missing", "Select the Image URL column in Step 1.")
            return
        if img not in self.columns:
            messagebox.showwarning("Column Error", "Selected image column not found in the loaded file.")
            return
        out = self.settings.get("output_folder", "").strip()
        if not out:
            messagebox.showwarning("Output Folder", "Set an output folder in Step 2.")
            return
        if not self.settings.get("url_patterns"):
            messagebox.showwarning("No Rules", "Add at least one URL replacement rule in Step 2.")
            return

        try:
            os.makedirs(out, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Folder Error", f"Cannot create output folder:\n{exc}")
            return

        test_run   = self.test_run_var.get()
        test_limit = max(1, self.test_limit_var.get())

        self._save_settings()
        self.nb.select(2)
        self._update_summary()
        self._log_clear()
        self.prog_bar["value"] = 0
        self.prog_lbl.config(text="Starting…")
        self.prog_count.config(text="")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.DISABLED)
        self.report_btn.config(state=tk.DISABLED)
        self._last_report_path = None
        self.stop_flag = False

        threading.Thread(
            target=self._worker, args=(test_run, test_limit), daemon=True
        ).start()

    def _stop(self):
        self.stop_flag = True
        self.prog_lbl.config(text="Stopping after current download…")

    def _open_folder(self):
        folder = self.settings.get("output_folder", "")
        if not os.path.isdir(folder):
            messagebox.showwarning("Not Found", "Output folder does not exist yet.")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            messagebox.showwarning("Open Folder", str(exc))

    # ── Download worker ───────────────────────────────────────────────────────

    def _worker(self, test_run=False, test_limit=10):
        key_col = self.settings.get("key_column", "")
        img_col = self.settings["image_column"]
        out     = self.settings["output_folder"]
        sep     = self.settings.get("separator", "|")
        rules   = self.settings.get("url_patterns", [])

        ki = self.columns.index(key_col) if key_col and key_col in self.columns else None
        ii = self.columns.index(img_col)

        source_rows = self.data_rows[:test_limit] if test_run else self.data_rows

        if test_run:
            self.root.after(0, lambda n=test_limit: self._log(
                f"[TEST] Running on first {n} rows only."
            ))

        # Build the full job list up front so we know the total
        # Each job: (key_val, raw_part, full_url)
        jobs = []
        for row_num, row in enumerate(source_rows, 1):
            cell = str(row[ii]) if row[ii] is not None else ""
            if not cell.strip():
                continue
            if ki is not None:
                key_val = str(row[ki]) if row[ki] is not None else "unknown"
            else:
                key_val = f"Row {row_num}"
            for part in cell.split(sep):
                part = part.strip()
                if part:
                    jobs.append((key_val, part, self._apply_rules(part, rules)))

        total = len(jobs)
        if total == 0:
            self.root.after(0, lambda: self._done([], False))
            return

        self.root.after(0, lambda t=total: self.prog_bar.config(maximum=t, value=0))

        downloaded = skipped = errors = 0
        seen_files = set()
        results = []

        for idx, (key_val, raw_part, url) in enumerate(jobs, 1):
            if self.stop_flag:
                self.root.after(
                    0, lambda r=results: self._done(r, True)
                )
                return

            # Skip non-HTTP URLs (rule didn't transform them)
            if not url.startswith(("http://", "https://")):
                msg = f"[SKIP] Not a valid URL: {url}"
                self.root.after(0, lambda m=msg: self._log(m))
                results.append({
                    "key": key_val, "original": raw_part, "url": url,
                    "filename": "", "status": "Invalid URL",
                    "note": "Did not match any replacement rule",
                })
                skipped += 1
                self.root.after(0, lambda v=idx, t=total: self._tick(v, t))
                continue

            fname = self._filename_from_url(url)
            # Deduplicate within this session
            base, ext = os.path.splitext(fname)
            counter = 1
            while fname in seen_files:
                fname = f"{base}_{counter}{ext}"
                counter += 1
            seen_files.add(fname)
            dest = os.path.join(out, fname)

            if os.path.exists(dest):
                msg = f"[SKIP] Already exists: {fname}"
                self.root.after(0, lambda m=msg: self._log(m))
                results.append({
                    "key": key_val, "original": raw_part, "url": url,
                    "filename": fname, "status": "Skipped",
                    "note": "File already exists in output folder",
                })
                skipped += 1
            else:
                ok, note = self._fetch(url, dest)
                if ok:
                    downloaded += 1
                    results.append({
                        "key": key_val, "original": raw_part, "url": url,
                        "filename": fname, "status": "Downloaded", "note": "",
                    })
                else:
                    errors += 1
                    results.append({
                        "key": key_val, "original": raw_part, "url": url,
                        "filename": fname, "status": "Error", "note": note,
                    })

            self.root.after(0, lambda v=idx, t=total: self._tick(v, t))

        self.root.after(
            0, lambda r=results: self._done(r, False)
        )

    @staticmethod
    def _apply_rules(raw, rules):
        result = raw
        for r in rules:
            try:
                result = re.sub(r["find"], r["replace"], result)
            except re.error:
                result = result.replace(r["find"], r["replace"])
        return result

    @staticmethod
    def _filename_from_url(url):
        # Handle  ?source=/path/to/file.jpg  style URLs (Oracle Commerce / ATG)
        if "source=" in url:
            for part in url.split("?", 1)[-1].split("&"):
                if part.startswith("source="):
                    fname = part[7:].rstrip("/").split("/")[-1]
                    if fname:
                        return fname
        # Fallback: last path segment before any query string
        fname = url.split("?")[0].rstrip("/").split("/")[-1]
        return fname or "image.jpg"

    @staticmethod
    def _cleanup(path):
        if os.path.exists(path):
            try:
                os.remove(path)
            except OSError:
                pass

    def _fetch(self, url, dest):
        """Download url to dest.  Returns (success, error_message)."""
        self.root.after(0, lambda: self._log(f"[DOWN] {url}"))
        error = ""
        try:
            resp = requests.get(url, timeout=30, stream=True)
            resp.raise_for_status()
            interrupted = False
            with open(dest, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=16384):
                    if self.stop_flag:
                        interrupted = True
                        break
                    fh.write(chunk)
            if interrupted:
                self._cleanup(dest)
                return False, "Interrupted by user"
            fname = os.path.basename(dest)
            self.root.after(0, lambda fn=fname: self._log(f"[ OK ] Saved: {fn}"))
            return True, ""
        except requests.exceptions.HTTPError as exc:
            error = f"HTTP error: {exc}"
            self.root.after(0, lambda m=error: self._log(f"[ERR] {m}"))
        except requests.exceptions.ConnectionError:
            error = "Connection failed"
            self.root.after(0, lambda u=url: self._log(f"[ERR] Connection failed: {u}"))
        except requests.exceptions.Timeout:
            error = "Request timed out"
            self.root.after(0, lambda u=url: self._log(f"[ERR] Timeout: {u}"))
        except OSError as exc:
            error = f"File write error: {exc}"
            self.root.after(0, lambda m=str(exc): self._log(f"[ERR] File write: {m}"))
        self._cleanup(dest)
        return False, error

    def _tick(self, value, total):
        self.prog_bar["value"] = value
        text = f"{value} / {total}"
        self.prog_count.config(text=text)
        self.prog_lbl.config(text=f"Downloading…  ({text})")

    def _done(self, results, stopped):
        self.stop_flag = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.NORMAL)

        downloaded = sum(1 for r in results if r["status"] == "Downloaded")
        skipped    = sum(1 for r in results if r["status"] in ("Skipped", "Invalid URL"))
        errors     = sum(1 for r in results if r["status"] == "Error")

        verb = "Stopped" if stopped else "Complete"
        self.prog_lbl.config(
            text=f"{verb} — Downloaded: {downloaded}  |  Skipped: {skipped}  |  Errors: {errors}"
        )
        self._log("─" * 52)
        self._log(f"{verb}!   Downloaded: {downloaded}   Skipped: {skipped}   Errors: {errors}")
        if not stopped:
            self._log(f"Output folder: {self.settings.get('output_folder', '')}")

        if results:
            report_path = self._save_report(results)
            if report_path:
                self._last_report_path = report_path
                self.report_btn.config(state=tk.NORMAL)
                self._log(f"Report saved: {os.path.basename(report_path)}")

        if not stopped and results:
            self._show_summary_popup(downloaded, skipped, errors, results)

    # ── Report helpers ────────────────────────────────────────────────────────

    def _save_report(self, results):
        """Write a CSV report to the output folder.  Returns the file path or None."""
        import datetime
        out = self.settings.get("output_folder", "")
        if not out:
            return None
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(out, f"photopull_report_{timestamp}.csv")
        try:
            with open(path, "w", newline="", encoding="utf-8") as fh:
                writer = csv.DictWriter(
                    fh,
                    fieldnames=["key", "original", "url", "filename", "status", "note"],
                )
                writer.writeheader()
                writer.writerows(results)
            return path
        except Exception as exc:
            self._log(f"[WARN] Could not save report: {exc}")
            return None

    def _view_report(self):
        if not self._last_report_path or not os.path.exists(self._last_report_path):
            messagebox.showwarning("Report", "No report file found.")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(self._last_report_path)
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", self._last_report_path])
            else:
                subprocess.Popen(["xdg-open", self._last_report_path])
        except Exception as exc:
            messagebox.showwarning("View Report", str(exc))

    def _show_summary_popup(self, downloaded, skipped, errors, results):
        dlg = tk.Toplevel(self.root)
        dlg.title("Download Complete — Summary")
        dlg.geometry("480x320")
        dlg.resizable(True, True)
        dlg.transient(self.root)

        total = len(results)
        pct = f"{downloaded / total * 100:.1f}%" if total else "—"

        ttk.Label(dlg, text="Download Summary", font=("Helvetica", 13, "bold")).pack(pady=(14, 4))

        grid = ttk.Frame(dlg, padding=10)
        grid.pack(fill=tk.X)
        for row_idx, (label, value, color) in enumerate([
            ("Total images processed:", str(total),       ""),
            ("Successfully downloaded:", str(downloaded), "green" if downloaded else ""),
            ("Skipped (already existed / invalid):", str(skipped), ""),
            ("Errors:", str(errors), "red" if errors else ""),
            ("Success rate:", pct, ""),
        ]):
            ttk.Label(grid, text=label, anchor=tk.W).grid(row=row_idx, column=0, sticky=tk.W, padx=6, pady=2)
            lbl = ttk.Label(grid, text=value, anchor=tk.W, font=("Helvetica", 10, "bold"))
            lbl.grid(row=row_idx, column=1, sticky=tk.W, padx=6, pady=2)
            if color:
                lbl.configure(foreground=color)

        if errors:
            ttk.Separator(dlg).pack(fill=tk.X, padx=10, pady=4)
            ttk.Label(dlg, text="Failed images:", anchor=tk.W).pack(anchor=tk.W, padx=12)
            err_box = scrolledtext.ScrolledText(
                dlg, height=6, state=tk.NORMAL, font=("Courier", 8), wrap=tk.WORD
            )
            err_box.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
            for r in results:
                if r["status"] == "Error":
                    err_box.insert(tk.END, f"{r['key']}  →  {r['filename'] or r['original']}\n")
                    if r["note"]:
                        err_box.insert(tk.END, f"   {r['note']}\n")
            err_box.config(state=tk.DISABLED)

        ttk.Button(dlg, text="Close", command=dlg.destroy).pack(pady=8)

    # ── Log helpers ───────────────────────────────────────────────────────────

    def _log(self, message):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def _log_clear(self):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.delete("1.0", tk.END)
        self.log_box.config(state=tk.DISABLED)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    import sys
    if not EXCEL_SUPPORT:
        print(
            "Note: openpyxl is not installed — Excel support is disabled.\n"
            "Install it with:  pip install openpyxl\n",
            file=sys.stderr,
        )
    root = tk.Tk()
    PhotoPullApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
