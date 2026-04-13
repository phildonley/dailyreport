"""
settings_manager.py
-------------------
Reads and writes app configuration to config/config.json
(located in the same directory as the app).

Also provides the Settings dialog (a Tkinter modal window).

Config keys
-----------
  db_path         - Full path to the WorkLog .db file
  html_path       - Full path where the HTML report is written
  author_name     - Shown in the HTML report header
  author_team     - Shown in the HTML report header
  author_org      - Shown in the HTML report header
  backup_enabled  - Boolean; if True, backup before every save
"""

import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from logger_setup import get_logger

log = get_logger(__name__)

# Config file lives in config/config.json next to the app scripts
_APP_DIR = os.path.dirname(os.path.abspath(__file__))
_CONFIG_DIR = os.path.join(_APP_DIR, "config")
_CONFIG_FILE = os.path.join(_CONFIG_DIR, "config.json")

# Default values for every key.  New keys added in future versions should
# also be added here so existing config files stay backward-compatible.
_DEFAULTS = {
    "db_path":        "",
    "html_path":      "",
    "author_name":    "",
    "author_team":    "",
    "author_org":     "",
    "backup_enabled": True,
}


# ── Config I/O ───────────────────────────────────────────────────────────────────

def load_config() -> dict:
    """
    Load config from disk.  Returns a fresh copy of _DEFAULTS if the file
    does not exist yet (first run).  Missing keys are filled from _DEFAULTS
    so the rest of the app can always rely on every key being present.
    """
    os.makedirs(_CONFIG_DIR, exist_ok=True)

    if not os.path.isfile(_CONFIG_FILE):
        log.info("No config file found — using defaults (first run).")
        return dict(_DEFAULTS)

    try:
        with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # Merge: start with defaults, overlay saved values
        merged = dict(_DEFAULTS)
        merged.update(data)
        log.debug("Config loaded from %s", _CONFIG_FILE)
        return merged
    except (json.JSONDecodeError, OSError) as exc:
        log.error("Failed to load config (%s); using defaults.", exc)
        return dict(_DEFAULTS)


def save_config(config: dict):
    """
    Write the config dict to disk.
    Raises OSError if the write fails (caller should show an error dialog).
    """
    os.makedirs(_CONFIG_DIR, exist_ok=True)
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2)
    log.info("Config saved to %s", _CONFIG_FILE)


# ── Settings dialog ──────────────────────────────────────────────────────────────

class SettingsDialog(tk.Toplevel):
    """
    Modal settings window.

    Usage:
        SettingsDialog(parent, current_config, on_save_callback=my_fn)

    on_save_callback(new_config) is called only when the user clicks Save
    and the config is written successfully.
    """

    def __init__(self, parent: tk.Widget, config: dict, on_save_callback=None):
        super().__init__(parent)
        self.title("Settings — WorkLog")
        self.resizable(False, False)
        self.transient(parent)   # Stays on top of parent
        self.grab_set()          # Modal: block parent interaction

        self._config = dict(config)          # Working copy
        self._on_save_callback = on_save_callback

        self._build_ui()
        self._center(parent)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.wait_window()  # Block until this window closes

    # ── Layout ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = ttk.Frame(self, padding=20)
        outer.pack(fill="both", expand=True)

        # ── Section: Database ────────────────────────────────────────────────────
        self._section_label(outer, "Database File", row=0)

        self._db_var = tk.StringVar(value=self._config.get("db_path", ""))
        db_row = ttk.Frame(outer)
        db_row.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 10))
        ttk.Entry(db_row, textvariable=self._db_var, width=54).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(db_row, text="Browse…", command=self._browse_db).pack(
            side="left", padx=(6, 0)
        )

        ttk.Label(
            outer,
            text="If the file does not exist you will be prompted to create it.",
            foreground="#777",
            font=("", 8),
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 8))

        # ── Section: HTML output ─────────────────────────────────────────────────
        self._section_label(outer, "HTML Report Output File", row=3)

        self._html_var = tk.StringVar(value=self._config.get("html_path", ""))
        html_row = ttk.Frame(outer)
        html_row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(2, 10))
        ttk.Entry(html_row, textvariable=self._html_var, width=54).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(html_row, text="Browse…", command=self._browse_html).pack(
            side="left", padx=(6, 0)
        )

        ttk.Label(
            outer,
            text="The report is regenerated automatically every time you save an entry.",
            foreground="#777",
            font=("", 8),
        ).grid(row=5, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 8))

        # ── Section: Branding ────────────────────────────────────────────────────
        ttk.Separator(outer, orient="horizontal").grid(
            row=6, column=0, columnspan=2, sticky="ew", pady=8
        )
        self._section_label(outer, "HTML Report Branding", row=7)
        ttk.Label(
            outer,
            text="These values appear in the report header shared with your manager.",
            foreground="#777",
            font=("", 8),
        ).grid(row=8, column=0, columnspan=2, sticky="w", padx=2, pady=(0, 6))

        self._author_name_var = tk.StringVar(value=self._config.get("author_name", ""))
        self._author_team_var = tk.StringVar(value=self._config.get("author_team", ""))
        self._author_org_var  = tk.StringVar(value=self._config.get("author_org",  ""))

        for row_i, (label_text, var) in enumerate([
            ("Your name:",     self._author_name_var),
            ("Team / group:",  self._author_team_var),
            ("Organization:",  self._author_org_var),
        ]):
            ttk.Label(outer, text=label_text).grid(
                row=9 + row_i, column=0, sticky="w", padx=(8, 4), pady=3
            )
            ttk.Entry(outer, textvariable=var, width=42).grid(
                row=9 + row_i, column=1, sticky="w", pady=3
            )

        # ── Section: Options ─────────────────────────────────────────────────────
        ttk.Separator(outer, orient="horizontal").grid(
            row=12, column=0, columnspan=2, sticky="ew", pady=8
        )
        self._backup_var = tk.BooleanVar(
            value=self._config.get("backup_enabled", True)
        )
        ttk.Checkbutton(
            outer,
            text="Create a daily backup of the database before each save",
            variable=self._backup_var,
        ).grid(row=13, column=0, columnspan=2, sticky="w", padx=4, pady=4)

        # ── Buttons ──────────────────────────────────────────────────────────────
        btn_frame = ttk.Frame(outer)
        btn_frame.grid(row=14, column=0, columnspan=2, pady=(16, 0))
        ttk.Button(btn_frame, text="Save",   command=self._on_save,   width=12).pack(side="left", padx=6)
        ttk.Button(btn_frame, text="Cancel", command=self._on_cancel, width=10).pack(side="left", padx=6)

        outer.columnconfigure(1, weight=1)

    def _section_label(self, parent, text, row):
        ttk.Label(parent, text=text, font=("", 10, "bold")).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(4, 0)
        )

    # ── Actions ───────────────────────────────────────────────────────────────────

    def _browse_db(self):
        path = filedialog.askopenfilename(
            title="Select WorkLog database file",
            filetypes=[("SQLite database", "*.db"), ("All files", "*.*")],
        )
        if not path:
            return  # User cancelled

        if os.path.isfile(path):
            # Existing file selected — use it directly
            self._db_var.set(path)
            log.debug("DB path chosen (existing): %s", path)
        else:
            # User typed a name that doesn't exist yet
            answer = messagebox.askyesno(
                "Database Not Found",
                f"No database file was found at:\n{path}\n\n"
                "Would you like to create a new database here?\n\n"
                "Click No to browse a different folder.",
                parent=self,
            )
            if answer:
                self._db_var.set(path)
                log.debug("DB path chosen (new): %s", path)
            # If No: dialog closes, Browse button is still available to try again

    def _browse_html(self):
        path = filedialog.asksaveasfilename(
            title="Select HTML report output location",
            filetypes=[("HTML file", "*.html")],
            defaultextension=".html",
            initialfile="worklog_report.html",
        )
        if path:
            self._html_var.set(path)
            log.debug("HTML path chosen: %s", path)

    def _on_save(self):
        self._config["db_path"]        = self._db_var.get().strip()
        self._config["html_path"]      = self._html_var.get().strip()
        self._config["author_name"]    = self._author_name_var.get().strip()
        self._config["author_team"]    = self._author_team_var.get().strip()
        self._config["author_org"]     = self._author_org_var.get().strip()
        self._config["backup_enabled"] = self._backup_var.get()

        try:
            save_config(self._config)
        except OSError as exc:
            messagebox.showerror(
                "Settings Error",
                f"Could not save settings:\n{exc}",
                parent=self,
            )
            return

        if self._on_save_callback:
            self._on_save_callback(self._config)
        self.destroy()

    def _on_cancel(self):
        self.destroy()

    # ── Utility ───────────────────────────────────────────────────────────────────

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
