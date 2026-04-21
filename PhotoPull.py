#!/usr/bin/env python3
"""
PhotoPull — Batch image downloader driven by Excel / CSV spreadsheets.
"""

import csv
import datetime
import json
import os
import platform
import re
import subprocess
import sys
import threading
import zipfile
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
# Path helpers
# ---------------------------------------------------------------------------

def resource_path(relative):
    """Resolve path for both dev and PyInstaller --onefile bundles."""
    try:
        base = sys._MEIPASS
    except AttributeError:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative)

SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "photopull_settings.json"
)

# ---------------------------------------------------------------------------
# Filename helpers
# ---------------------------------------------------------------------------

_UNSAFE_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

def safe_filename(text):
    """Make arbitrary text safe for use inside a filename."""
    text = str(text).replace(" ", "_")
    text = _UNSAFE_RE.sub("", text)
    return text.strip("._") or "unknown"

def format_seq(n, digits=3, pad=True):
    return str(n).zfill(digits) if pad else str(n)

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_SETTINGS = {
    "output_folder":  str(Path.home() / "Downloads" / "PhotoPull"),
    "url_patterns": [
        {
            "find":    r"media\d+:/products/",
            "replace": ("https://myparts.terex.com/ccstore/v1/images/"
                        "?source=/file/products/"),
        }
    ],
    "separator":   "|",
    "has_headers": True,
    "id_column":   "",
    "sku_column":  "",
    "desc_column": "",
    "image_column":"",
    "last_file":   "",
    # Naming
    "naming_mode": "original",
    "pattern_tokens": [
        {"type": "field", "value": "id"},
        {"type": "sep",   "value": "."},
        {"type": "field", "value": "sku"},
        {"type": "sep",   "value": "."},
        {"type": "sequence"},
    ],
    "seq_digits":    3,
    "seq_start":     103,
    "seq_increment": 2,
    "seq_pad":       True,
    # ZIP
    "zip_enabled":    False,
    "zip_batch_size": 500,
}

# ---------------------------------------------------------------------------
# Splash screen  (500×500, 5 s or click/keypress to dismiss)
# ---------------------------------------------------------------------------

class SplashScreen:
    DURATION_MS = 5000

    def __init__(self, root):
        self.root = root
        self.win  = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.resizable(False, False)
        self._center(500, 500)
        self._build()
        for seq in ("<Button-1>", "<Key>"):
            self.win.bind(seq, self._dismiss)
        self.win.focus_force()
        self._job = self.win.after(self.DURATION_MS, self._dismiss)

    def _center(self, w, h):
        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        self.win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build(self):
        img_path = resource_path(os.path.join("assets", "splash.png"))
        if os.path.exists(img_path):
            try:
                img = tk.PhotoImage(file=img_path)
                lbl = tk.Label(self.win, image=img, cursor="hand2")
                lbl.image = img
                lbl.pack(fill=tk.BOTH, expand=True)
                lbl.bind("<Button-1>", self._dismiss)
                return
            except Exception:
                pass
        # Fallback canvas when image not yet available
        c = tk.Canvas(self.win, width=500, height=500,
                      bg="#1a252f", highlightthickness=0)
        c.pack(fill=tk.BOTH, expand=True)
        c.create_text(250, 200, text="PhotoPull",
                      font=("Helvetica", 52, "bold"), fill="#ecf0f1")
        c.create_text(250, 268, text="Batch Image Downloader",
                      font=("Helvetica", 15), fill="#95a5a6")
        c.create_text(250, 465, text="Click anywhere or wait 5 seconds…",
                      font=("Helvetica", 10), fill="#7f8c8d")
        c.bind("<Button-1>", self._dismiss)

    def _dismiss(self, _=None):
        try:
            self.win.after_cancel(self._job)
        except Exception:
            pass
        self.win.destroy()

# ---------------------------------------------------------------------------
# URL rule editor dialog
# ---------------------------------------------------------------------------

class RuleDialog(tk.Toplevel):
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
        ttk.Label(self, text="Find pattern (regex):").grid(
            row=0, column=0, sticky=tk.W, **pad)
        self.find_var = tk.StringVar(value=find)
        ttk.Entry(self, textvariable=self.find_var, width=58).grid(
            row=0, column=1, sticky=tk.EW, **pad)
        ttk.Label(self, text="Replace with (URL prefix):").grid(
            row=1, column=0, sticky=tk.W, **pad)
        self.replace_var = tk.StringVar(value=replace)
        ttk.Entry(self, textvariable=self.replace_var, width=58).grid(
            row=1, column=1, sticky=tk.EW, **pad)
        ttk.Label(self,
            text="Tip: \\d+ matches any number — media\\d+:/products/ covers media1, media2, media3 …",
            foreground="gray", wraplength=560).grid(
            row=2, column=0, columnspan=2, sticky=tk.W, padx=12)
        bf = ttk.Frame(self)
        bf.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(bf, text="Save",   command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(bf, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)
        self.columnconfigure(1, weight=1)

    def _save(self):
        f, r = self.find_var.get().strip(), self.replace_var.get().strip()
        if not f or not r:
            messagebox.showwarning("Required", "Both fields are required.", parent=self)
            return
        self.result = (f, r)
        self.destroy()

# ---------------------------------------------------------------------------
# Main application
# ---------------------------------------------------------------------------

class PhotoPullApp:

    def __init__(self, root):
        self.root = root
        self.root.title("PhotoPull — Image Downloader")
        self.root.geometry("960x720")
        self.root.minsize(780, 580)

        self.settings        = {}
        self.columns         = []
        self.data_rows       = []
        self.stop_flag       = False
        self._last_report    = None
        self._preview_offset = 0

        self._load_settings()
        self._build_ui()
        self._apply_settings_to_ui()

    # ── Settings ──────────────────────────────────────────────────────────────

    def _load_settings(self):
        self.settings = {k: v for k, v in DEFAULT_SETTINGS.items()}
        self.settings["url_patterns"]   = [dict(r) for r in DEFAULT_SETTINGS["url_patterns"]]
        self.settings["pattern_tokens"] = [dict(t) for t in DEFAULT_SETTINGS["pattern_tokens"]]
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, encoding="utf-8") as fh:
                    self.settings.update(json.load(fh))
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
        self.settings["output_folder"] = self.output_var.get().strip()
        self.settings["separator"]     = self.sep_var.get() or "|"
        self.settings["has_headers"]   = self.has_headers_var.get()
        self.settings["id_column"]     = self.id_col_var.get()
        self.settings["sku_column"]    = self.sku_col_var.get()
        self.settings["desc_column"]   = self.desc_col_var.get()
        self.settings["image_column"]  = self.img_col_var.get()
        self.settings["last_file"]     = self.file_var.get().strip()
        self.settings["url_patterns"]  = [
            {"find": self.rules_tree.item(i, "values")[0],
             "replace": self.rules_tree.item(i, "values")[1]}
            for i in self.rules_tree.get_children()
        ]
        self.settings["naming_mode"]    = self.naming_mode_var.get()
        self.settings["pattern_tokens"] = list(self._pattern_tokens)
        self.settings["seq_digits"]     = self.seq_digits_var.get()
        self.settings["seq_start"]      = self.seq_start_var.get()
        self.settings["seq_increment"]  = self.seq_inc_var.get()
        self.settings["seq_pad"]        = self.seq_pad_var.get()
        self.settings["zip_enabled"]    = self.zip_enabled_var.get()
        self.settings["zip_batch_size"] = self.zip_batch_var.get()

    def _apply_settings_to_ui(self):
        self.output_var.set(self.settings.get("output_folder", ""))
        self.sep_var.set(self.settings.get("separator", "|"))
        self.has_headers_var.set(self.settings.get("has_headers", True))
        self._reload_rules_tree()
        self.naming_mode_var.set(self.settings.get("naming_mode", "original"))
        self._pattern_tokens = [dict(t) for t in self.settings.get("pattern_tokens", [])]
        self._rebuild_pattern_list()
        self.seq_digits_var.set(self.settings.get("seq_digits", 3))
        self.seq_start_var.set(self.settings.get("seq_start", 103))
        self.seq_inc_var.set(self.settings.get("seq_increment", 2))
        self.seq_pad_var.set(self.settings.get("seq_pad", True))
        self.zip_enabled_var.set(self.settings.get("zip_enabled", False))
        self.zip_batch_var.set(self.settings.get("zip_batch_size", 500))
        self._update_seq_preview()
        self._update_pattern_preview()
        self._on_naming_mode_change()
        last = self.settings.get("last_file", "")
        if last:
            self.file_var.set(last)

    # ── UI shell ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        hdr = tk.Frame(self.root, bg="#1a252f", pady=10)
        hdr.pack(fill=tk.X)
        tk.Label(hdr, text="PhotoPull",
                 font=("Helvetica", 20, "bold"), bg="#1a252f", fg="#ecf0f1").pack()
        tk.Label(hdr, text="Batch Image Downloader  •  Excel & CSV",
                 font=("Helvetica", 9), bg="#1a252f", fg="#95a5a6").pack()
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=10)
        self._tab_file()
        self._tab_settings()
        self._tab_naming()
        self._tab_run()

    # ── Tab 1 — File & Columns ────────────────────────────────────────────────

    def _tab_file(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 1 · File & Columns  ")

        ttk.Label(f, text=(
            "Select your Excel or CSV file, confirm headers, load it, "
            "then map each column.  Only the Image column is required."),
            foreground="gray", wraplength=880).pack(anchor=tk.W, pady=(0, 8))

        # File row
        fg = ttk.LabelFrame(f, text="Input File", padding=10)
        fg.pack(fill=tk.X, pady=(0, 8))
        r1 = ttk.Frame(fg)
        r1.pack(fill=tk.X)
        ttk.Label(r1, text="File:").pack(side=tk.LEFT)
        self.file_var = tk.StringVar()
        ttk.Entry(r1, textvariable=self.file_var).pack(
            side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(r1, text="Browse…", command=self._browse_input).pack(side=tk.LEFT)
        r2 = ttk.Frame(fg)
        r2.pack(fill=tk.X, pady=(8, 0))
        self.has_headers_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(r2, text="First row contains column headers",
                        variable=self.has_headers_var,
                        command=self._on_headers_toggle).pack(side=tk.LEFT)
        ttk.Button(r2, text="Load File  ▸", command=self.load_file).pack(side=tk.RIGHT)

        # Column pickers (4 rows)
        cp = ttk.LabelFrame(f, text="Column Mapping", padding=10)
        cp.pack(fill=tk.X, pady=(0, 8))

        col_defs = [
            ("ID / Catalog Number:",  "id_col",   "Optional — online catalog ID used in renamed filenames & reports"),
            ("SKU / Part Number:",     "sku_col",  "Optional — part number used in renamed filenames & reports"),
            ("Description:",           "desc_col", "Optional — text used in renamed filenames (spaces → underscores)"),
            ("Image URL Column:",      "img_col",  "Required — the column containing image paths or partial URLs"),
        ]
        for label_text, attr, hint in col_defs:
            row = ttk.Frame(cp)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=label_text, width=24).pack(side=tk.LEFT)
            var   = tk.StringVar()
            combo = ttk.Combobox(row, textvariable=var, state="readonly", width=32)
            combo.pack(side=tk.LEFT, padx=4)
            ttk.Label(row, text=f"← {hint}", foreground="gray").pack(side=tk.LEFT)
            setattr(self, f"{attr}_var",   var)
            setattr(self, f"{attr}_combo", combo)

        # Preview with navigation
        pv = ttk.LabelFrame(f, text="Preview", padding=5)
        pv.pack(fill=tk.BOTH, expand=True)

        nav = ttk.Frame(pv)
        nav.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(nav, text="◀ Prev 5", command=self._preview_prev).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav, text="Next 5 ▶", command=self._preview_next).pack(side=tk.LEFT, padx=2)
        self.preview_range_lbl = ttk.Label(nav, text="", foreground="gray")
        self.preview_range_lbl.pack(side=tk.LEFT, padx=8)

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
        ttk.Label(f, text="Configure output folder, URL replacement rules, and ZIP options.",
                  foreground="gray", wraplength=880).pack(anchor=tk.W, pady=(0, 8))

        # Output folder
        of = ttk.LabelFrame(f, text="Output Folder", padding=10)
        of.pack(fill=tk.X, pady=(0, 8))
        r = ttk.Frame(of)
        r.pack(fill=tk.X)
        ttk.Label(r, text="Save images to:").pack(side=tk.LEFT)
        self.output_var = tk.StringVar()
        ttk.Entry(r, textvariable=self.output_var).pack(
            side=tk.LEFT, padx=6, fill=tk.X, expand=True)
        ttk.Button(r, text="Browse…", command=self._browse_output).pack(side=tk.LEFT)
        ttk.Label(of, text="All images saved into one flat folder — no sub-folders.",
                  foreground="gray").pack(anchor=tk.W, pady=(5, 0))

        # Separator + ZIP side by side
        mid = ttk.Frame(f)
        mid.pack(fill=tk.X, pady=(0, 8))

        sf = ttk.LabelFrame(mid, text="Multiple-Image Separator", padding=10)
        sf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        sr = ttk.Frame(sf)
        sr.pack(fill=tk.X)
        ttk.Label(sr, text="Separator:").pack(side=tk.LEFT)
        self.sep_var = tk.StringVar(value="|")
        ttk.Entry(sr, textvariable=self.sep_var, width=4).pack(side=tk.LEFT, padx=6)
        ttk.Label(sr, text="Default: |  e.g. img1.jpg | img2.jpg",
                  foreground="gray").pack(side=tk.LEFT)

        zf = ttk.LabelFrame(mid, text="ZIP Output", padding=10)
        zf.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.zip_enabled_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(zf, text="Zip images after download (originals deleted after zipping)",
                        variable=self.zip_enabled_var).pack(anchor=tk.W)
        zr = ttk.Frame(zf)
        zr.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(zr, text="Images per ZIP:").pack(side=tk.LEFT)
        self.zip_batch_var = tk.IntVar(value=500)
        ttk.Spinbox(zr, from_=10, to=10000, textvariable=self.zip_batch_var,
                    width=7).pack(side=tk.LEFT, padx=6)
        ttk.Label(zr, text="(named photopull_batch_001.zip, 002.zip …)",
                  foreground="gray").pack(side=tk.LEFT)

        # URL rules
        rl = ttk.LabelFrame(f, text="URL Replacement Rules", padding=10)
        rl.pack(fill=tk.BOTH, expand=True)
        ttk.Label(rl, text=(
            "Each rule converts a partial path into a full download URL.  "
            "Find is a regex — \\d+ matches any digit, so one rule covers media1, media2, media3 …"),
            foreground="gray", wraplength=860).pack(anchor=tk.W, pady=(0, 6))

        cols = ("find", "replace")
        self.rules_tree = ttk.Treeview(rl, columns=cols, show="headings", height=4)
        self.rules_tree.heading("find",    text="Find  (regex)")
        self.rules_tree.heading("replace", text="Replace With")
        self.rules_tree.column("find",    width=220)
        self.rules_tree.column("replace", width=560)
        rsy = ttk.Scrollbar(rl, orient=tk.VERTICAL, command=self.rules_tree.yview)
        self.rules_tree.configure(yscrollcommand=rsy.set)
        rsy.pack(side=tk.RIGHT, fill=tk.Y)
        self.rules_tree.pack(fill=tk.BOTH, expand=True)

        btns = ttk.Frame(rl)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="Add",          command=self._rule_add).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Edit",         command=self._rule_edit).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Remove",       command=self._rule_remove).pack(side=tk.LEFT, padx=2)
        ttk.Button(btns, text="Reset to Defaults", command=self._rule_reset).pack(side=tk.LEFT, padx=12)
        ttk.Button(btns, text="Save Settings",command=self._save_settings_ui).pack(side=tk.RIGHT, padx=2)

    # ── Tab 3 — Naming ────────────────────────────────────────────────────────

    def _tab_naming(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 3 · Naming  ")
        ttk.Label(f, text=(
            "Choose how downloaded images are named.  "
            "Pattern mode builds the filename from your spreadsheet columns + a sequence number."),
            foreground="gray", wraplength=880).pack(anchor=tk.W, pady=(0, 8))

        # Mode toggle
        mf = ttk.LabelFrame(f, text="Naming Mode", padding=10)
        mf.pack(fill=tk.X, pady=(0, 8))
        self.naming_mode_var = tk.StringVar(value="original")
        ttk.Radiobutton(mf, text="Keep original filename (as downloaded — no renaming)",
                        variable=self.naming_mode_var, value="original",
                        command=self._on_naming_mode_change).pack(anchor=tk.W)
        ttk.Radiobutton(mf, text="Use custom naming pattern  (e.g. ID.SKU.103.jpg)",
                        variable=self.naming_mode_var, value="pattern",
                        command=self._on_naming_mode_change).pack(anchor=tk.W)

        # Pattern builder
        self.builder_frame = ttk.LabelFrame(f, text="Filename Pattern Builder", padding=10)
        self.builder_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        left = ttk.Frame(self.builder_frame)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))
        ttk.Label(left, text="FIELDS", foreground="gray",
                  font=("Helvetica", 8, "bold")).pack(anchor=tk.W)
        for label, fval in [("[ID]", "id"), ("[SKU]", "sku"), ("[Description]", "desc")]:
            ttk.Button(left, text=label, width=16,
                       command=lambda v=fval: self._add_token({"type":"field","value":v})
                       ).pack(pady=1, anchor=tk.W)
        ttk.Label(left, text="SEQUENCE", foreground="gray",
                  font=("Helvetica", 8, "bold")).pack(anchor=tk.W, pady=(8, 0))
        ttk.Button(left, text="[Sequence]", width=16,
                   command=lambda: self._add_token({"type":"sequence"})
                   ).pack(pady=1, anchor=tk.W)
        ttk.Label(left, text="SEPARATORS", foreground="gray",
                  font=("Helvetica", 8, "bold")).pack(anchor=tk.W, pady=(8, 0))
        sr = ttk.Frame(left)
        sr.pack(anchor=tk.W)
        for sep in [".", "-", "_"]:
            ttk.Button(sr, text=sep, width=4,
                       command=lambda s=sep: self._add_token({"type":"sep","value":s})
                       ).pack(side=tk.LEFT, padx=1)
        ttk.Label(left, text="STATIC TEXT", foreground="gray",
                  font=("Helvetica", 8, "bold")).pack(anchor=tk.W, pady=(8, 0))
        self.static_var = tk.StringVar()
        ttk.Entry(left, textvariable=self.static_var, width=16).pack(anchor=tk.W)
        ttk.Button(left, text="Add Text", width=16,
                   command=self._add_static_token).pack(pady=2, anchor=tk.W)

        right = ttk.Frame(self.builder_frame)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ttk.Label(right, text="Current Pattern  (select then move or remove)",
                  font=("Helvetica", 9, "bold")).pack(anchor=tk.W, pady=(0, 4))
        lf2 = ttk.Frame(right)
        lf2.pack(fill=tk.BOTH, expand=True)
        self.pattern_list = tk.Listbox(lf2, height=7, selectmode=tk.SINGLE,
                                       font=("Courier", 12))
        pls = ttk.Scrollbar(lf2, orient=tk.VERTICAL, command=self.pattern_list.yview)
        self.pattern_list.configure(yscrollcommand=pls.set)
        pls.pack(side=tk.RIGHT, fill=tk.Y)
        self.pattern_list.pack(fill=tk.BOTH, expand=True)

        pb = ttk.Frame(right)
        pb.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(pb, text="↑ Up",      command=self._token_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(pb, text="↓ Down",    command=self._token_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(pb, text="Remove",    command=self._token_remove).pack(side=tk.LEFT, padx=2)
        ttk.Button(pb, text="Clear All", command=self._token_clear).pack(side=tk.LEFT, padx=10)
        ttk.Button(pb, text="Reset Default", command=self._token_reset).pack(side=tk.LEFT, padx=2)

        pr = ttk.Frame(right)
        pr.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(pr, text="Live preview:").pack(side=tk.LEFT)
        self.pattern_preview_lbl = ttk.Label(pr, text="",
                                              foreground="#27ae60",
                                              font=("Courier", 10, "bold"))
        self.pattern_preview_lbl.pack(side=tk.LEFT, padx=6)

        # Sequence settings
        self.seq_frame = ttk.LabelFrame(f, text="Sequence Settings", padding=10)
        self.seq_frame.pack(fill=tk.X)
        sg = ttk.Frame(self.seq_frame)
        sg.pack(fill=tk.X)
        fields = [("Digits:", "seq_digits_var", 1, 6, 3, 5),
                  ("Start:",  "seq_start_var",  1, 99999, 103, 7),
                  ("Increment:", "seq_inc_var", 1, 100, 2, 5)]
        for col, (lbl, attr, lo, hi, default, w) in enumerate(fields):
            ttk.Label(sg, text=lbl).grid(row=0, column=col*2, sticky=tk.W, padx=(10,2), pady=4)
            var = tk.IntVar(value=default)
            setattr(self, attr, var)
            ttk.Spinbox(sg, from_=lo, to=hi, textvariable=var, width=w,
                        command=self._update_seq_preview).grid(
                row=0, column=col*2+1, sticky=tk.W, padx=(0, 10))
        ttk.Label(sg, text="Pad with zeros:").grid(row=0, column=6, sticky=tk.W, padx=(10,2))
        self.seq_pad_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(sg, variable=self.seq_pad_var,
                        command=self._update_seq_preview).grid(row=0, column=7, sticky=tk.W)
        self.seq_preview_lbl = ttk.Label(self.seq_frame, text="",
                                          foreground="#27ae60", font=("Courier", 10))
        self.seq_preview_lbl.pack(anchor=tk.W, padx=10, pady=(2, 4))

        self._pattern_tokens = [dict(t) for t in DEFAULT_SETTINGS["pattern_tokens"]]
        self._rebuild_pattern_list()
        self._update_seq_preview()
        self._on_naming_mode_change()

    # ── Tab 4 — Run ───────────────────────────────────────────────────────────

    def _tab_run(self):
        f = ttk.Frame(self.nb, padding=12)
        self.nb.add(f, text="  Step 4 · Run  ")
        ttk.Label(f, text="Review the summary, then start.  Settings save automatically on each run.",
                  foreground="gray", wraplength=880).pack(anchor=tk.W, pady=(0, 8))

        # Run options (test mode)
        ro = ttk.LabelFrame(f, text="Run Options", padding=10)
        ro.pack(fill=tk.X, pady=(0, 8))
        opt = ttk.Frame(ro)
        opt.pack(fill=tk.X)
        self.test_run_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="Test run — process only the first",
                        variable=self.test_run_var).pack(side=tk.LEFT)
        self.test_limit_var = tk.IntVar(value=10)
        ttk.Spinbox(opt, from_=1, to=9999, textvariable=self.test_limit_var,
                    width=6).pack(side=tk.LEFT, padx=4)
        ttk.Label(opt, text="rows   (verify settings before a full run)",
                  foreground="gray").pack(side=tk.LEFT)

        # Summary
        sf = ttk.LabelFrame(f, text="Summary", padding=10)
        sf.pack(fill=tk.X, pady=(0, 8))
        self.summary_lbl = ttk.Label(sf, text="Load a file in Step 1 to see a summary.",
                                     justify=tk.LEFT)
        self.summary_lbl.pack(anchor=tk.W)

        # Progress
        pf = ttk.LabelFrame(f, text="Progress", padding=10)
        pf.pack(fill=tk.X, pady=(0, 8))
        self.prog_lbl   = ttk.Label(pf, text="Idle — ready to start.")
        self.prog_lbl.pack(anchor=tk.W, pady=(0, 4))
        self.prog_bar   = ttk.Progressbar(pf, mode="determinate")
        self.prog_bar.pack(fill=tk.X)
        self.prog_count = ttk.Label(pf, text="")
        self.prog_count.pack(anchor=tk.W, pady=(3, 0))

        # Buttons
        bf = ttk.Frame(f)
        bf.pack(fill=tk.X, pady=(0, 8))
        self.start_btn  = ttk.Button(bf, text="▶  Start Download",  command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn   = ttk.Button(bf, text="◼  Stop", command=self._stop,
                                     state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        self.open_btn   = ttk.Button(bf, text="Open Output Folder",
                                     command=self._open_folder, state=tk.DISABLED)
        self.open_btn.pack(side=tk.RIGHT, padx=2)
        self.report_btn = ttk.Button(bf, text="View Report",
                                     command=self._view_report, state=tk.DISABLED)
        self.report_btn.pack(side=tk.RIGHT, padx=2)

        # Log
        lf = ttk.LabelFrame(f, text="Activity Log", padding=5)
        lf.pack(fill=tk.BOTH, expand=True)
        self.log_box = scrolledtext.ScrolledText(
            lf, height=12, state=tk.DISABLED, font=("Courier", 9), wrap=tk.WORD)
        self.log_box.pack(fill=tk.BOTH, expand=True)

    # ── File loading ──────────────────────────────────────────────────────────

    def _browse_input(self):
        types = ([("Spreadsheets","*.xlsx *.xls *.csv"),
                  ("Excel","*.xlsx *.xls"), ("CSV","*.csv"), ("All","*.*")]
                 if EXCEL_SUPPORT else
                 [("CSV","*.csv"), ("All","*.*")])
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
            ext  = os.path.splitext(path)[1].lower()
            hdrs = self.has_headers_var.get()
            if ext in (".xlsx", ".xls"):
                if not EXCEL_SUPPORT:
                    messagebox.showerror("Missing Library",
                                         "Install openpyxl:\n  pip install openpyxl")
                    return
                self.columns, self.data_rows = self._read_excel(path, hdrs)
            elif ext == ".csv":
                self.columns, self.data_rows = self._read_csv(path, hdrs)
            else:
                messagebox.showerror("Unsupported", "Please choose an Excel or CSV file.")
                return
            self._preview_offset = 0
            self._refresh_columns()
            self._refresh_preview()
            self.settings["last_file"] = path
            self._update_summary()
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not load file:\n{exc}")

    @staticmethod
    def _read_excel(path, has_headers):
        wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws   = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        if not rows:
            return [], []
        if has_headers:
            headers = [str(c) if c is not None else f"Col{i+1}"
                       for i, c in enumerate(rows[0])]
            return headers, rows[1:]
        return [f"Column {i+1}" for i in range(len(rows[0]))], rows

    @staticmethod
    def _read_csv(path, has_headers):
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
        if not rows:
            return [], []
        if has_headers:
            return [c.strip() for c in rows[0]], rows[1:]
        return [f"Column {i+1}" for i in range(len(rows[0]))], rows

    def _refresh_columns(self):
        saved = {
            "id_col":   self.settings.get("id_column",   ""),
            "sku_col":  self.settings.get("sku_column",  ""),
            "desc_col": self.settings.get("desc_column", ""),
            "img_col":  self.settings.get("image_column",""),
        }
        for attr, key in [("id_col","id_col"), ("sku_col","sku_col"),
                           ("desc_col","desc_col"), ("img_col","img_col")]:
            combo = getattr(self, f"{attr}_combo")
            var   = getattr(self, f"{attr}_var")
            combo["values"] = self.columns
            sv = saved[key]
            var.set(sv if sv in self.columns else (self.columns[0] if self.columns else ""))

    def _refresh_preview(self):
        t = self.preview_tree
        t.delete(*t.get_children())
        if not self.columns:
            return
        t["columns"] = self.columns
        for col in self.columns:
            t.heading(col, text=col)
            width = max(90, min(240, len(col) * 9 + 28))
            t.column(col, width=width, minwidth=70)
        start = self._preview_offset
        end   = start + 5
        for row in self.data_rows[start:end]:
            t.insert("", tk.END,
                     values=[str(v) if v is not None else "" for v in row])
        total = len(self.data_rows)
        self.preview_range_lbl.config(
            text=f"Rows {start+1}–{min(end,total)} of {total}")

    def _preview_prev(self):
        self._preview_offset = max(0, self._preview_offset - 5)
        self._refresh_preview()

    def _preview_next(self):
        if self._preview_offset + 5 < len(self.data_rows):
            self._preview_offset += 5
        self._refresh_preview()

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
            messagebox.showinfo("Edit", "Select a rule first.")
            return
        v = self.rules_tree.item(sel[0], "values")
        dlg = RuleDialog(self.root, find=v[0], replace=v[1])
        if dlg.result:
            self.rules_tree.item(sel[0], values=dlg.result)

    def _rule_remove(self):
        sel = self.rules_tree.selection()
        if not sel:
            messagebox.showinfo("Remove", "Select a rule first.")
            return
        if messagebox.askyesno("Remove Rule", "Delete selected rule?"):
            self.rules_tree.delete(sel[0])

    def _rule_reset(self):
        if messagebox.askyesno("Reset", "Reset URL rules to defaults?"):
            self.settings["url_patterns"] = [dict(r) for r in DEFAULT_SETTINGS["url_patterns"]]
            self._reload_rules_tree()

    def _save_settings_ui(self):
        self._collect_settings_from_ui()
        self._save_settings(notify=True)

    # ── Naming-tab helpers ────────────────────────────────────────────────────

    def _token_display(self, token):
        t = token["type"]
        if t == "field":
            return {"id":"[ID]","sku":"[SKU]","desc":"[Description]"}.get(token["value"], token["value"])
        if t == "sequence":
            return "[Sequence]"
        if t == "sep":
            return f'  {token["value"]}  '
        if t == "static":
            return f'"{token["value"]}"'
        return str(token)

    def _rebuild_pattern_list(self):
        self.pattern_list.delete(0, tk.END)
        for token in self._pattern_tokens:
            self.pattern_list.insert(tk.END, self._token_display(token))
        self._update_pattern_preview()
        self._update_seq_preview()

    def _add_token(self, token):
        self._pattern_tokens.append(dict(token))
        self._rebuild_pattern_list()

    def _add_static_token(self):
        text = self.static_var.get().strip()
        if not text:
            messagebox.showwarning("Static Text", "Enter some text first.")
            return
        self._add_token({"type": "static", "value": text})
        self.static_var.set("")

    def _token_up(self):
        sel = self.pattern_list.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self._pattern_tokens[i-1], self._pattern_tokens[i] = \
            self._pattern_tokens[i], self._pattern_tokens[i-1]
        self._rebuild_pattern_list()
        self.pattern_list.selection_set(i-1)

    def _token_down(self):
        sel = self.pattern_list.curselection()
        if not sel or sel[0] >= len(self._pattern_tokens)-1:
            return
        i = sel[0]
        self._pattern_tokens[i], self._pattern_tokens[i+1] = \
            self._pattern_tokens[i+1], self._pattern_tokens[i]
        self._rebuild_pattern_list()
        self.pattern_list.selection_set(i+1)

    def _token_remove(self):
        sel = self.pattern_list.curselection()
        if not sel:
            return
        del self._pattern_tokens[sel[0]]
        self._rebuild_pattern_list()

    def _token_clear(self):
        if messagebox.askyesno("Clear Pattern", "Remove all tokens?"):
            self._pattern_tokens.clear()
            self._rebuild_pattern_list()

    def _token_reset(self):
        if messagebox.askyesno("Reset", "Reset pattern to default?"):
            self._pattern_tokens = [dict(t) for t in DEFAULT_SETTINGS["pattern_tokens"]]
            self._rebuild_pattern_list()

    def _on_naming_mode_change(self):
        state = tk.NORMAL if self.naming_mode_var.get() == "pattern" else tk.DISABLED
        for child in self.builder_frame.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass
        for child in self.seq_frame.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass

    def _update_seq_preview(self, *_):
        try:
            start = self.seq_start_var.get()
            inc   = self.seq_inc_var.get()
            d     = self.seq_digits_var.get()
            pad   = self.seq_pad_var.get()
            nums  = [format_seq(start + i*inc, d, pad) for i in range(5)]
            self.seq_preview_lbl.config(text="Next 5: " + ",  ".join(nums))
        except Exception:
            pass

    def _update_pattern_preview(self, *_):
        try:
            start = self.seq_start_var.get()
            d     = self.seq_digits_var.get()
            pad   = self.seq_pad_var.get()
            sample = self._build_filename(
                id_val="10025142", sku_val="0090-193", desc_val="Sample_Part",
                seq_num=start, ext=".jpg",
                tokens=self._pattern_tokens,
                digits=d, pad=pad)
            self.pattern_preview_lbl.config(text=sample)
        except Exception:
            self.pattern_preview_lbl.config(text="")

    def _build_filename(self, id_val, sku_val, desc_val, seq_num, ext, tokens, digits, pad):
        parts = []
        for tok in tokens:
            t = tok["type"]
            if t == "field":
                raw = {"id": id_val, "sku": sku_val,
                       "desc": desc_val}.get(tok["value"], "")
                parts.append(safe_filename(raw))
            elif t == "sep":
                parts.append(tok["value"])
            elif t == "static":
                parts.append(safe_filename(tok["value"]))
            elif t == "sequence":
                parts.append(format_seq(seq_num, digits, pad))
        name = "".join(parts)
        return (name or "image") + ext

    @staticmethod
    def _apply_rules(raw, rules):
        result = raw
        for r in rules:
            try:
                result = re.sub(r["find"], r["replace"], result)
            except re.error:
                result = result.replace(r["find"], r["replace"])
        return result

    # ── Run-tab helpers ───────────────────────────────────────────────────────

    def _update_summary(self):
        if not self.data_rows:
            self.summary_lbl.config(text="No data loaded.")
            return
        img = self.img_col_var.get()
        if not img or img not in self.columns:
            self.summary_lbl.config(text="Select the Image URL column in Step 1.")
            return
        ii  = self.columns.index(img)
        sep = self.sep_var.get() or "|"
        total_imgs = sum(
            len([p for p in str(r[ii]).split(sep) if p.strip()])
            for r in self.data_rows
            if r[ii] is not None and str(r[ii]).strip())
        total_rows = sum(
            1 for r in self.data_rows
            if r[ii] is not None and str(r[ii]).strip())
        out  = self.output_var.get() or "(not set — see Step 2)"
        mode = self.naming_mode_var.get()
        self.summary_lbl.config(text=(
            f"File:              {os.path.basename(self.file_var.get())}\n"
            f"Rows with images:  {total_rows} of {len(self.data_rows)}\n"
            f"Total images:      {total_imgs}\n"
            f"Naming mode:       {'Original filenames' if mode=='original' else 'Custom pattern'}\n"
            f"Output folder:     {out}"))

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
            messagebox.showwarning("Column Error", "Image column not found in loaded file.")
            return
        out = self.settings.get("output_folder", "").strip()
        if not out:
            messagebox.showwarning("Output Folder", "Set an output folder in Step 2.")
            return
        if not self.settings.get("url_patterns"):
            messagebox.showwarning("No Rules", "Add a URL replacement rule in Step 2.")
            return
        try:
            os.makedirs(out, exist_ok=True)
        except Exception as exc:
            messagebox.showerror("Folder Error", f"Cannot create output folder:\n{exc}")
            return
        test_run   = self.test_run_var.get()
        test_limit = max(1, self.test_limit_var.get())
        self._save_settings()
        self.nb.select(3)
        self._update_summary()
        self._log_clear()
        self.prog_bar["value"] = 0
        self.prog_lbl.config(text="Starting…")
        self.prog_count.config(text="")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.DISABLED)
        self.report_btn.config(state=tk.DISABLED)
        self._last_report = None
        self.stop_flag = False
        threading.Thread(target=self._worker,
                         args=(test_run, test_limit), daemon=True).start()

    def _stop(self):
        self.stop_flag = True
        self.prog_lbl.config(text="Stopping after current download…")

    def _open_folder(self):
        folder = self.settings.get("output_folder", "")
        if not os.path.isdir(folder):
            messagebox.showwarning("Not Found", "Output folder does not exist.")
            return
        try:
            if   platform.system() == "Windows": os.startfile(folder)
            elif platform.system() == "Darwin":  subprocess.Popen(["open", folder])
            else:                                subprocess.Popen(["xdg-open", folder])
        except Exception as exc:
            messagebox.showwarning("Open Folder", str(exc))

    # ── Download worker ───────────────────────────────────────────────────────

    def _worker(self, test_run=False, test_limit=10):
        s       = self.settings
        img_col = s["image_column"]
        out     = s["output_folder"]
        sep     = s.get("separator", "|")
        rules   = s.get("url_patterns", [])
        mode    = s.get("naming_mode", "original")
        tokens  = s.get("pattern_tokens", [])
        digits  = s.get("seq_digits", 3)
        start   = s.get("seq_start", 103)
        inc     = s.get("seq_increment", 2)
        pad     = s.get("seq_pad", True)

        def col_idx(key):
            v = s.get(key, "")
            return self.columns.index(v) if v and v in self.columns else None

        ii   = self.columns.index(img_col)
        id_i = col_idx("id_column")
        sk_i = col_idx("sku_column")
        ds_i = col_idx("desc_column")

        rows = self.data_rows[:test_limit] if test_run else self.data_rows
        if test_run:
            self.root.after(0, lambda n=test_limit:
                self._log(f"[TEST] First {n} rows only."))

        # Build job list: (id_val, sku_val, desc_val, raw_part, url, seq_within_row)
        jobs = []
        for row_num, row in enumerate(rows, 1):
            cell = str(row[ii]) if row[ii] is not None else ""
            if not cell.strip():
                continue
            id_v  = str(row[id_i])  if id_i  is not None and row[id_i]  is not None else ""
            sk_v  = str(row[sk_i])  if sk_i  is not None and row[sk_i]  is not None else ""
            ds_v  = str(row[ds_i])  if ds_i  is not None and row[ds_i]  is not None else ""
            parts = [p.strip() for p in cell.split(sep) if p.strip()]
            for seq_idx, part in enumerate(parts):
                seq_num = start + seq_idx * inc
                jobs.append((id_v, sk_v, ds_v, part,
                              self._apply_rules(part, rules), seq_num))

        total = len(jobs)
        if total == 0:
            self.root.after(0, lambda: self._done([], False))
            return

        self.root.after(0, lambda t=total: self.prog_bar.config(maximum=t, value=0))

        downloaded_paths = []
        results          = []
        seen_files       = set()
        dl = sk = er = 0

        for idx, (id_v, sk_v, ds_v, raw, url, seq_num) in enumerate(jobs, 1):
            if self.stop_flag:
                self.root.after(0, lambda r=results: self._done(r, True))
                return

            if not url.startswith(("http://", "https://")):
                msg = f"[SKIP] Invalid URL: {url}"
                self.root.after(0, lambda m=msg: self._log(m))
                results.append({"id":id_v,"sku":sk_v,"original":raw,"url":url,
                                 "filename":"","status":"Invalid URL","note":"No rule matched"})
                sk += 1
                self.root.after(0, lambda v=idx, t=total: self._tick(v, t))
                continue

            ext  = os.path.splitext(self._filename_from_url(url))[1] or ".jpg"
            if mode == "pattern" and tokens:
                fname = self._build_filename(id_v, sk_v, ds_v, seq_num, ext,
                                             tokens, digits, pad)
            else:
                fname = self._filename_from_url(url)

            base, fext = os.path.splitext(fname)
            ctr = 1
            while fname in seen_files:
                fname = f"{base}_{ctr}{fext}"
                ctr  += 1
            seen_files.add(fname)
            dest = os.path.join(out, fname)

            if os.path.exists(dest):
                msg = f"[SKIP] Already exists: {fname}"
                self.root.after(0, lambda m=msg: self._log(m))
                results.append({"id":id_v,"sku":sk_v,"original":raw,"url":url,
                                 "filename":fname,"status":"Skipped",
                                 "note":"Already exists"})
                sk += 1
            else:
                ok, note = self._fetch(url, dest)
                if ok:
                    dl += 1
                    downloaded_paths.append(dest)
                    results.append({"id":id_v,"sku":sk_v,"original":raw,"url":url,
                                    "filename":fname,"status":"Downloaded","note":""})
                else:
                    er += 1
                    results.append({"id":id_v,"sku":sk_v,"original":raw,"url":url,
                                    "filename":fname,"status":"Error","note":note})

            self.root.after(0, lambda v=idx, t=total: self._tick(v, t))

        # ZIP phase
        if downloaded_paths and s.get("zip_enabled") and not self.stop_flag:
            batch_size = max(1, s.get("zip_batch_size", 500))
            self.root.after(0, lambda: self._log("Zipping files…"))
            self._zip_files(downloaded_paths, out, batch_size)

        self.root.after(0, lambda r=results: self._done(r, False))

    # ── Download helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _filename_from_url(url):
        if "source=" in url:
            for part in url.split("?",1)[-1].split("&"):
                if part.startswith("source="):
                    fn = part[7:].rstrip("/").split("/")[-1]
                    if fn:
                        return fn
        fn = url.split("?")[0].rstrip("/").split("/")[-1]
        return fn or "image.jpg"

    @staticmethod
    def _cleanup(path):
        if os.path.exists(path):
            try: os.remove(path)
            except OSError: pass

    def _fetch(self, url, dest):
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
                return False, "Interrupted"
            fn = os.path.basename(dest)
            self.root.after(0, lambda f=fn: self._log(f"[ OK ] Saved: {f}"))
            return True, ""
        except requests.exceptions.HTTPError as exc:
            error = f"HTTP error: {exc}"
            self.root.after(0, lambda m=error: self._log(f"[ERR] {m}"))
        except requests.exceptions.ConnectionError:
            error = "Connection failed"
            self.root.after(0, lambda u=url: self._log(f"[ERR] Connection failed: {u}"))
        except requests.exceptions.Timeout:
            error = "Timed out"
            self.root.after(0, lambda u=url: self._log(f"[ERR] Timeout: {u}"))
        except OSError as exc:
            error = f"File write: {exc}"
            self.root.after(0, lambda m=str(exc): self._log(f"[ERR] {m}"))
        self._cleanup(dest)
        return False, error

    def _zip_files(self, paths, out_folder, batch_size):
        batches = [paths[i:i+batch_size] for i in range(0, len(paths), batch_size)]
        for batch_num, batch in enumerate(batches, 1):
            zip_name = f"photopull_batch_{batch_num:03d}.zip"
            zip_path = os.path.join(out_folder, zip_name)
            try:
                with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fp in batch:
                        if os.path.exists(fp):
                            zf.write(fp, os.path.basename(fp))
                for fp in batch:
                    self._cleanup(fp)
                msg = f"[ZIP ] Created: {zip_name}  ({len(batch)} files)"
                self.root.after(0, lambda m=msg: self._log(m))
            except Exception as exc:
                msg = f"[ERR] Zip failed: {exc}"
                self.root.after(0, lambda m=msg: self._log(m))

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

        dl = sum(1 for r in results if r["status"] == "Downloaded")
        sk = sum(1 for r in results if r["status"] in ("Skipped","Invalid URL"))
        er = sum(1 for r in results if r["status"] == "Error")
        verb = "Stopped" if stopped else "Complete"
        self.prog_lbl.config(
            text=f"{verb} — Downloaded: {dl}  |  Skipped: {sk}  |  Errors: {er}")
        self._log("─" * 52)
        self._log(f"{verb}!  Downloaded: {dl}   Skipped: {sk}   Errors: {er}")
        if not stopped:
            self._log(f"Output: {self.settings.get('output_folder','')}")

        if results:
            rp = self._save_report(results)
            if rp:
                self._last_report = rp
                self.report_btn.config(state=tk.NORMAL)
                self._log(f"Report: {os.path.basename(rp)}")

        if not stopped and results:
            self._show_summary_popup(dl, sk, er, results)

    # ── Report helpers ────────────────────────────────────────────────────────

    def _save_report(self, results):
        out = self.settings.get("output_folder", "")
        if not out:
            return None
        ts   = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(out, f"photopull_report_{ts}.csv")
        try:
            with open(path, "w", newline="", encoding="utf-8") as fh:
                w = csv.DictWriter(fh, fieldnames=[
                    "id","sku","original","url","filename","status","note"])
                w.writeheader()
                w.writerows(results)
            return path
        except Exception as exc:
            self._log(f"[WARN] Could not save report: {exc}")
            return None

    def _view_report(self):
        if not self._last_report or not os.path.exists(self._last_report):
            messagebox.showwarning("Report", "No report file found.")
            return
        try:
            if   platform.system() == "Windows": os.startfile(self._last_report)
            elif platform.system() == "Darwin":  subprocess.Popen(["open", self._last_report])
            else:                                subprocess.Popen(["xdg-open", self._last_report])
        except Exception as exc:
            messagebox.showwarning("View Report", str(exc))

    def _show_summary_popup(self, dl, sk, er, results):
        dlg = tk.Toplevel(self.root)
        dlg.title("Download Complete")
        dlg.geometry("480x300")
        dlg.resizable(True, True)
        dlg.transient(self.root)
        total = len(results)
        pct   = f"{dl/total*100:.1f}%" if total else "—"
        ttk.Label(dlg, text="Download Summary",
                  font=("Helvetica", 13, "bold")).pack(pady=(14, 4))
        g = ttk.Frame(dlg, padding=10)
        g.pack(fill=tk.X)
        for ri, (lbl, val, color) in enumerate([
            ("Total processed:",        str(total), ""),
            ("Successfully downloaded:",str(dl),    "green" if dl else ""),
            ("Skipped:",                str(sk),    ""),
            ("Errors:",                 str(er),    "red"   if er else ""),
            ("Success rate:",           pct,        ""),
        ]):
            ttk.Label(g, text=lbl, anchor=tk.W).grid(
                row=ri, column=0, sticky=tk.W, padx=6, pady=2)
            lb = ttk.Label(g, text=val, font=("Helvetica", 10, "bold"), anchor=tk.W)
            lb.grid(row=ri, column=1, sticky=tk.W, padx=6, pady=2)
            if color:
                lb.configure(foreground=color)
        if er:
            ttk.Separator(dlg).pack(fill=tk.X, padx=10, pady=4)
            ttk.Label(dlg, text="Failed images:", anchor=tk.W).pack(anchor=tk.W, padx=12)
            eb = scrolledtext.ScrolledText(dlg, height=5, font=("Courier", 8), wrap=tk.WORD)
            eb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))
            for r in results:
                if r["status"] == "Error":
                    eb.insert(tk.END,
                              f"{r['id'] or r['sku'] or 'row'}  →  "
                              f"{r['filename'] or r['original']}\n")
                    if r["note"]:
                        eb.insert(tk.END, f"   {r['note']}\n")
            eb.config(state=tk.DISABLED)
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
    root = tk.Tk()
    root.withdraw()          # hide main window during splash
    splash = SplashScreen(root)
    root.wait_window(splash.win)
    root.deiconify()         # show main window after splash closes
    PhotoPullApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()






