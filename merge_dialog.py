"""
merge_dialog.py
---------------
Database merge dialog and execution logic.

Workflow
--------
  1. User browses to a source .db file
  2. App validates that the source is a genuine WorkLog database
  3. App shows a preview: new dates vs conflicting dates
  4. User chooses a conflict-resolution strategy
  5. Merge executes inside a single SQLite transaction
     (rolled back entirely on any error — current DB is never partially modified)
  6. Summary dialog shown on completion

Conflict strategies
-------------------
  append   — Add incoming work items as extra rows on the existing day
             (non-destructive; your data is never removed)
  skip     — Keep your existing entry; discard the incoming one
  replace  — Overwrite your entry with the incoming one (warned clearly)
"""

import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

from logger_setup import get_logger
from database import Database, APP_NAME

log = get_logger(__name__)


def open_merge_dialog(parent: tk.Widget, current_db: Database, on_complete=None):
    """
    Convenience function — open the MergeDialog.
    on_complete() is called (with no arguments) after a successful merge.
    """
    MergeDialog(parent, current_db, on_complete)


class MergeDialog(tk.Toplevel):
    """
    Modal dialog that guides the user through merging a second WorkLog
    database into the currently open one.
    """

    def __init__(self, parent: tk.Widget, current_db: Database, on_complete=None):
        super().__init__(parent)
        self.title("Merge Database — WorkLog")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._current_db = current_db
        self._on_complete = on_complete
        self._preview_data = None   # Populated after the Preview step

        self._build_ui()
        self._center(parent)
        self.wait_window()

    # ── Layout ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = ttk.Frame(self, padding=22)
        outer.pack(fill="both", expand=True)

        ttk.Label(
            outer,
            text="Merge another WorkLog database into the currently open one.",
            font=("", 10),
        ).pack(anchor="w", pady=(0, 14))

        # ── Source file row ──────────────────────────────────────────────────────
        src_frame = ttk.Frame(outer)
        src_frame.pack(fill="x")

        ttk.Label(src_frame, text="Source database:").pack(side="left")
        self._src_var = tk.StringVar()
        ttk.Entry(src_frame, textvariable=self._src_var, width=46).pack(
            side="left", padx=8, fill="x", expand=True
        )
        ttk.Button(src_frame, text="Browse\u2026", command=self._browse_source).pack(
            side="left"
        )

        # ── Preview result area ──────────────────────────────────────────────────
        self._preview_var = tk.StringVar(value="")
        preview_label = ttk.Label(
            outer,
            textvariable=self._preview_var,
            foreground="#2a4080",
            font=("", 10),
            justify="left",
        )
        preview_label.pack(anchor="w", pady=(14, 4))

        # ── Conflict strategy ────────────────────────────────────────────────────
        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(
            outer,
            text="When the same date exists in BOTH databases:",
            font=("", 10, "bold"),
        ).pack(anchor="w")

        self._strategy_var = tk.StringVar(value="append")
        strategies = [
            (
                "append",
                "Append \u2014 add incoming rows to the existing day (safe, non-destructive)",
            ),
            (
                "skip",
                "Skip \u2014 keep my existing entry, ignore the incoming one",
            ),
            (
                "replace",
                "Replace \u2014 overwrite my entry with the incoming one",
            ),
        ]
        for val, text in strategies:
            ttk.Radiobutton(
                outer,
                text=text,
                variable=self._strategy_var,
                value=val,
            ).pack(anchor="w", padx=14, pady=3)

        ttk.Label(
            outer,
            text="Tip: 'Append' is the safest choice if you are unsure.",
            foreground="#888",
            font=("", 8),
        ).pack(anchor="w", padx=14, pady=(0, 6))

        # ── Action buttons ────────────────────────────────────────────────────────
        ttk.Separator(outer, orient="horizontal").pack(fill="x", pady=10)
        btn_frame = ttk.Frame(outer)
        btn_frame.pack()

        self._preview_btn = ttk.Button(
            btn_frame, text="Preview", command=self._do_preview, width=12
        )
        self._preview_btn.pack(side="left", padx=4)

        self._merge_btn = ttk.Button(
            btn_frame,
            text="Merge",
            command=self._do_merge,
            width=12,
            state="disabled",   # Enabled only after a successful preview
        )
        self._merge_btn.pack(side="left", padx=4)

        ttk.Button(
            btn_frame, text="Cancel", command=self.destroy, width=10
        ).pack(side="left", padx=4)

    # ── Actions ───────────────────────────────────────────────────────────────────

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select source WorkLog database",
            filetypes=[("SQLite database", "*.db"), ("All files", "*.*")],
        )
        if path:
            self._src_var.set(path)
            # Reset state so user must preview again
            self._preview_var.set("")
            self._merge_btn.config(state="disabled")
            self._preview_data = None

    def _do_preview(self):
        """Validate the source file and show a preview summary."""
        source_path = self._src_var.get().strip()
        if not source_path:
            messagebox.showwarning(
                "No File Selected",
                "Please browse to a source database file first.",
                parent=self,
            )
            return

        # Validate
        ok, reason = Database.validate_file(source_path)
        if not ok:
            messagebox.showerror(
                "Invalid Database",
                f"The selected file is not a valid WorkLog database:\n\n{reason}",
                parent=self,
            )
            return

        if source_path == self._current_db.db_path:
            messagebox.showerror(
                "Same File",
                "The source and destination are the same file.",
                parent=self,
            )
            return

        # Read the source dates
        try:
            conn = sqlite3.connect(source_path, timeout=5)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("SELECT entry_date FROM daily_entries ORDER BY entry_date")
            source_dates = {row["entry_date"] for row in cur.fetchall()}
            conn.close()
        except sqlite3.Error as exc:
            messagebox.showerror(
                "Read Error",
                f"Could not read the source database:\n{exc}",
                parent=self,
            )
            return

        current_dates  = set(self._current_db.get_all_dates())
        new_dates      = source_dates - current_dates
        conflict_dates = source_dates & current_dates

        self._preview_data = {
            "source_path":    source_path,
            "source_dates":   source_dates,
            "new_dates":      new_dates,
            "conflict_dates": conflict_dates,
        }

        self._preview_var.set(
            f"Source database has {len(source_dates)} entr{'y' if len(source_dates)==1 else 'ies'}.\n"
            f"  \u2022 {len(new_dates)} new date(s) \u2014 will be imported\n"
            f"  \u2022 {len(conflict_dates)} conflict(s) \u2014 date(s) present in both databases"
        )
        self._merge_btn.config(state="normal")
        log.info(
            "Merge preview: %d new, %d conflict(s) from %s",
            len(new_dates), len(conflict_dates), source_path,
        )

    def _do_merge(self):
        """Execute the merge after user confirmation."""
        if not self._preview_data:
            return

        strategy       = self._strategy_var.get()
        source_path    = self._preview_data["source_path"]
        new_count      = len(self._preview_data["new_dates"])
        conflict_count = len(self._preview_data["conflict_dates"])

        confirm_msg = (
            f"Ready to merge into the current database:\n\n"
            f"  \u2022 {new_count} new entr{'y' if new_count==1 else 'ies'} will be added\n"
            f"  \u2022 {conflict_count} conflict(s) will be handled with: "
            f"{strategy.upper()}\n\n"
            f"This operation modifies your current database.\n"
            f"Your daily backup will be created before the merge.\n\n"
            f"Continue?"
        )
        if not messagebox.askyesno("Confirm Merge", confirm_msg, icon="warning", parent=self):
            return

        # Backup before modifying
        self._current_db.backup()

        try:
            result = _execute_merge(
                current_db=self._current_db,
                source_path=source_path,
                conflict_strategy=strategy,
            )
        except Exception as exc:
            messagebox.showerror(
                "Merge Failed",
                f"The merge failed and your database was NOT modified.\n\nError:\n{exc}",
                parent=self,
            )
            log.error("Merge failed: %s", exc)
            return

        summary = (
            f"Merge completed successfully.\n\n"
            f"  \u2022 {result['added']}    entr{'y' if result['added']==1 else 'ies'} added\n"
            f"  \u2022 {result['skipped']}  entr{'y' if result['skipped']==1 else 'ies'} skipped\n"
            f"  \u2022 {result['replaced']} entr{'y' if result['replaced']==1 else 'ies'} replaced\n"
            f"  \u2022 {result['appended']} entr{'y' if result['appended']==1 else 'ies'} had rows appended"
        )
        messagebox.showinfo("Merge Complete", summary, parent=self)
        log.info("Merge complete: %s", result)

        if self._on_complete:
            self._on_complete()
        self.destroy()

    # ── Utility ───────────────────────────────────────────────────────────────────

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")


# ── Merge execution (pure logic, no UI) ─────────────────────────────────────────

def _execute_merge(
    current_db: Database,
    source_path: str,
    conflict_strategy: str,
) -> dict:
    """
    Read all data from source_path and write it into current_db.

    Runs entirely inside a single SQLite transaction on current_db.
    If anything raises an exception the ROLLBACK is issued and the
    exception is re-raised to the caller.

    Returns a summary dict:
        { 'added': int, 'skipped': int, 'replaced': int, 'appended': int }
    """

    def now():
        return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S")

    result = {"added": 0, "skipped": 0, "replaced": 0, "appended": 0}

    # ── Read ALL source data first (before touching current_db) ──────────────────
    src_conn = sqlite3.connect(source_path, timeout=5)
    src_conn.row_factory = sqlite3.Row
    src_cur = src_conn.cursor()

    src_cur.execute("SELECT * FROM daily_entries ORDER BY entry_date")
    source_entries = src_cur.fetchall()

    # Map entry_id → list of work_item rows
    source_items: dict = {}
    for entry in source_entries:
        src_cur.execute(
            "SELECT * FROM work_items WHERE entry_id = ? ORDER BY sort_order",
            (entry["id"],),
        )
        source_items[entry["id"]] = src_cur.fetchall()

    src_conn.close()
    log.info("Source read: %d entries, strategies=%s", len(source_entries), conflict_strategy)

    # ── Execute inside one transaction on current_db ─────────────────────────────
    current_dates = set(current_db.get_all_dates())
    conn = current_db._conn

    try:
        conn.execute("BEGIN")

        for entry in source_entries:
            date_str = entry["entry_date"]
            items    = source_items[entry["id"]]

            if date_str not in current_dates:
                # ── New date: insert entry + all work items ──────────────────────
                cur = conn.execute(
                    "INSERT INTO daily_entries "
                    "(entry_date, notes, created_at, updated_at) VALUES (?, ?, ?, ?)",
                    (date_str, entry["notes"], entry["created_at"], now()),
                )
                new_id = cur.lastrowid
                for order, item in enumerate(items):
                    conn.execute(
                        "INSERT INTO work_items "
                        "(entry_id, sort_order, task, reason, hours, created_at, updated_at) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (new_id, order, item["task"], item["reason"],
                         item["hours"], item["created_at"], now()),
                    )
                result["added"] += 1
                log.debug("Merge: added %s", date_str)

            else:
                # ── Conflict: apply chosen strategy ─────────────────────────────
                existing = current_db.get_entry_by_date(date_str)
                existing_id = existing["id"]

                if conflict_strategy == "skip":
                    result["skipped"] += 1
                    log.debug("Merge: skipped %s", date_str)

                elif conflict_strategy == "replace":
                    conn.execute(
                        "UPDATE daily_entries SET notes=?, updated_at=? WHERE id=?",
                        (entry["notes"], now(), existing_id),
                    )
                    conn.execute(
                        "DELETE FROM work_items WHERE entry_id=?", (existing_id,)
                    )
                    for order, item in enumerate(items):
                        conn.execute(
                            "INSERT INTO work_items "
                            "(entry_id, sort_order, task, reason, hours, created_at, updated_at) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (existing_id, order, item["task"], item["reason"],
                             item["hours"], item["created_at"], now()),
                        )
                    result["replaced"] += 1
                    log.debug("Merge: replaced %s", date_str)

                elif conflict_strategy == "append":
                    # Find the current highest sort_order so we don't collide
                    cur = conn.execute(
                        "SELECT COALESCE(MAX(sort_order), -1) AS max_order "
                        "FROM work_items WHERE entry_id=?",
                        (existing_id,),
                    )
                    max_order = cur.fetchone()[0]
                    for i, item in enumerate(items):
                        conn.execute(
                            "INSERT INTO work_items "
                            "(entry_id, sort_order, task, reason, hours, created_at, updated_at) "
                            "VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (existing_id, max_order + 1 + i, item["task"],
                             item["reason"], item["hours"], item["created_at"], now()),
                        )
                    conn.execute(
                        "UPDATE daily_entries SET updated_at=? WHERE id=?",
                        (now(), existing_id),
                    )
                    result["appended"] += 1
                    log.debug("Merge: appended rows to %s", date_str)

        conn.execute("COMMIT")
        log.info("Merge transaction committed: %s", result)

    except Exception as exc:
        conn.execute("ROLLBACK")
        log.error("Merge rolled back due to error: %s", exc)
        raise

    return result
