"""
main.py — WorkLog Manager entry point and main window.
"""

import os
import sys
import tkinter as tk
import webbrowser
from datetime import date, datetime
from tkinter import messagebox, ttk
from typing import Optional

from logger_setup import setup_logging, get_logger
from database import Database
from html_generator import generate_html
from settings_manager import SettingsDialog, load_config, save_config
from calendar_popup import CalendarPopup
from merge_dialog import open_merge_dialog

setup_logging()
log = get_logger(__name__)

APP_TITLE   = "WorkLog Manager"
APP_VERSION = "1.0.0"
MIN_WIDTH   = 920
MIN_HEIGHT  = 600

# Column definitions: (id, heading, width, stretch)
COLUMNS = [
    ("task",   "Task Performed",   370, True),
    ("reason", "Reason / Context", 280, True),
    ("hours",  "Hours",             80, False),
]


class WorkLogApp(tk.Tk):

    def __init__(self):
        super().__init__()
        log.info("Starting %s v%s", APP_TITLE, APP_VERSION)
        self.title(APP_TITLE)
        self.minsize(MIN_WIDTH, MIN_HEIGHT)
        self.geometry(f"{MIN_WIDTH}x{MIN_HEIGHT}")

        self._db: Optional[Database] = None
        self._config: dict = load_config()
        self._current_entry_id: Optional[int] = None
        self._unsaved_changes = False
        self._loading = False          # suppresses notes-trace during load
        self._edit_widgets: dict = {}  # active inline-edit overlay

        self._build_menu()
        self._build_layout()
        self._apply_style()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self._startup_connect)

    # ── Style ────────────────────────────────────────────────────────────────

    def _apply_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Toolbar.TButton", padding=(8, 4))
        s.configure("Accent.TButton", padding=(12, 6), font=("", 10, "bold"))
        s.map("Accent.TButton",
              background=[("active", "#2a4080"), ("!active", "#1a2744")],
              foreground=[("active", "white"),   ("!active", "white")])
        s.configure("Header.TFrame", background="#1a2744")
        s.configure("Header.TLabel",
                    background="#1a2744", foreground="white",
                    font=("", 14, "bold"))
        s.configure("HeaderSub.TLabel",
                    background="#1a2744", foreground="#c8a951", font=("", 9))
        s.configure("WorkItems.Treeview",        rowheight=28, font=("", 10))
        s.configure("WorkItems.Treeview.Heading", font=("", 9, "bold"))

    # ── Menu ─────────────────────────────────────────────────────────────────

    def _build_menu(self):
        bar = tk.Menu(self)
        fm = tk.Menu(bar, tearoff=0)
        fm.add_command(label="Settings\u2026",        command=self._open_settings)
        fm.add_command(label="Merge Database\u2026",  command=self._open_merge)
        fm.add_separator()
        fm.add_command(label="Exit",                  command=self._on_close)
        bar.add_cascade(label="File", menu=fm)

        vm = tk.Menu(bar, tearoff=0)
        vm.add_command(label="Regenerate HTML Report",     command=self._regenerate_html)
        vm.add_command(label="Open HTML Report in Browser", command=self._open_html_in_browser)
        bar.add_cascade(label="View", menu=vm)

        hm = tk.Menu(bar, tearoff=0)
        hm.add_command(label=f"About {APP_TITLE}", command=self._show_about)
        bar.add_cascade(label="Help", menu=hm)
        self.config(menu=bar)

    # ── Layout ────────────────────────────────────────────────────────────────

    def _build_layout(self):
        # Header bar
        header = ttk.Frame(self, style="Header.TFrame", height=48)
        header.pack(side="top", fill="x")
        header.pack_propagate(False)
        ttk.Label(header, text=APP_TITLE, style="Header.TLabel").pack(
            side="left", padx=16, pady=10)
        ttk.Label(header, text=f"v{APP_VERSION}", style="HeaderSub.TLabel").pack(
            side="left", pady=14)
        ttk.Button(header, text="\u2699  Settings",
                   command=self._open_settings,
                   style="Toolbar.TButton").pack(side="right", padx=12, pady=8)

        # Split pane
        pane = tk.PanedWindow(self, orient="horizontal",
                              sashrelief="flat", sashwidth=5,
                              background="#dde1ea")
        pane.pack(fill="both", expand=True)

        left = ttk.Frame(pane, width=200)
        pane.add(left, minsize=160)
        self._build_left_panel(left)

        right = ttk.Frame(pane)
        pane.add(right, minsize=500)
        self._build_right_panel(right)

    def _build_left_panel(self, parent):
        tb = ttk.Frame(parent)
        tb.pack(fill="x", padx=6, pady=(8, 4))

        ttk.Button(tb, text="+ Add",   command=self._on_add_entry,
                   style="Toolbar.TButton", width=6).pack(side="left", padx=(0, 2))
        ttk.Button(tb, text="Delete",  command=self._on_delete_entry,
                   style="Toolbar.TButton", width=6).pack(side="left", padx=2)
        ttk.Button(tb, text="\U0001f4c5", command=self._open_calendar,
                   style="Toolbar.TButton", width=3).pack(side="right")

        lf = ttk.Frame(parent)
        lf.pack(fill="both", expand=True, padx=6, pady=(0, 8))

        self._date_listbox = tk.Listbox(
            lf, selectmode="browse", activestyle="none",
            font=("", 10), relief="flat", borderwidth=1,
            highlightthickness=1, highlightcolor="#2a4080",
            selectbackground="#1a2744", selectforeground="white", bg="white")
        sb = ttk.Scrollbar(lf, orient="vertical",
                           command=self._date_listbox.yview)
        self._date_listbox.config(yscrollcommand=sb.set)
        self._date_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._date_listbox.bind("<<ListboxSelect>>", self._on_date_selected)
        self._date_listbox.bind("<Up>",
            lambda e: self.after(1, self._on_date_selected))
        self._date_listbox.bind("<Down>",
            lambda e: self.after(1, self._on_date_selected))

        self._no_db_label = ttk.Label(parent,
            text="No database.\nOpen Settings \u2192",
            foreground="#999", justify="center", font=("", 9))
        self._no_db_label.pack(pady=6)

    def _build_right_panel(self, parent):
        # Date + unsaved indicator
        date_bar = ttk.Frame(parent)
        date_bar.pack(fill="x", padx=16, pady=(12, 4))
        ttk.Label(date_bar, text="Entry Date:", font=("", 10)).pack(side="left")
        self._date_display_var = tk.StringVar(value="\u2014 no entry selected \u2014")
        ttk.Label(date_bar, textvariable=self._date_display_var,
                  font=("", 11, "bold"), foreground="#1a2744").pack(
                  side="left", padx=8)
        self._unsaved_label = ttk.Label(date_bar,
            text="\u25cf Unsaved changes",
            foreground="#c0392b", font=("", 9))
        # packed/unpacked dynamically

        # Notes
        nf = ttk.Frame(parent)
        nf.pack(fill="x", padx=16, pady=(0, 6))
        ttk.Label(nf, text="Daily note (optional):").pack(side="left")
        self._notes_var = tk.StringVar()
        self._notes_var.trace_add("write", self._on_field_change)
        ttk.Entry(nf, textvariable=self._notes_var, width=58).pack(
            side="left", padx=8)

        # Work-items treeview
        tf = ttk.Frame(parent)
        tf.pack(fill="both", expand=True, padx=16, pady=(0, 4))
        cols = [c[0] for c in COLUMNS]
        self._tree = ttk.Treeview(tf, columns=cols, show="headings",
                                  selectmode="browse",
                                  style="WorkItems.Treeview")
        for cid, cname, width, stretch in COLUMNS:
            self._tree.heading(cid, text=cname)
            self._tree.column(cid, width=width, stretch=stretch, minwidth=50)
        tsb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=tsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")
        self._tree.bind("<Double-1>", self._on_tree_double_click)
        self._tree.bind("<Return>",   self._on_tree_double_click)
        self._tree.bind("<Delete>",   lambda e: self._on_remove_row())

        # Row toolbar
        rf = ttk.Frame(parent)
        rf.pack(fill="x", padx=16, pady=(2, 6))
        ttk.Button(rf, text="+ Add Row",    command=self._on_add_row,
                   style="Toolbar.TButton").pack(side="left", padx=(0, 6))
        ttk.Button(rf, text="\u2212 Remove Row", command=self._on_remove_row,
                   style="Toolbar.TButton").pack(side="left")
        self._total_var = tk.StringVar(value="Daily total: 0.0 hrs")
        ttk.Label(rf, textvariable=self._total_var,
                  font=("", 10), foreground="#2a4080").pack(side="right")

        # Save button
        sf = ttk.Frame(parent)
        sf.pack(fill="x", padx=16, pady=(0, 6))
        self._save_btn = ttk.Button(sf, text="\U0001f4be  Save Entry",
                                    command=self._on_save,
                                    style="Accent.TButton", state="disabled")
        self._save_btn.pack(side="right")

        # Status bar — shows result of last save operation
        self._status_var = tk.StringVar(value="")
        self._status_label = ttk.Label(
            parent, textvariable=self._status_var,
            font=("", 9), foreground="#555")
        self._status_label.pack(fill="x", padx=16, pady=(0, 10))

    # ── Startup / DB connection ───────────────────────────────────────────────

    def _startup_connect(self):
        db_path = self._config.get("db_path", "").strip()
        if db_path:
            self._connect_database(db_path)
        else:
            self._show_no_db_state()
            if messagebox.askyesno("Welcome to WorkLog",
                    "No database configured.\nOpen Settings to get started?",
                    icon="info"):
                self._open_settings()

    def _connect_database(self, db_path: str):
        if not os.path.isfile(db_path):
            if not messagebox.askyesno("Create Database?",
                    f"No file found at:\n{db_path}\n\nCreate a new database here?"):
                self._show_no_db_state()
                return
        try:
            if self._db:
                self._db.close()
            self._db = Database(db_path)
            log.info("Connected: %s", db_path)
            self._show_db_state()
            self._refresh_date_list()
        except Exception as exc:
            log.error("DB open failed: %s", exc)
            messagebox.showerror("Database Error",
                f"Could not open:\n{db_path}\n\n{exc}")
            self._show_no_db_state()

    def _show_no_db_state(self):
        self._no_db_label.pack(pady=6)
        self._save_btn.config(state="disabled")
        self.title(APP_TITLE)

    def _show_db_state(self):
        self._no_db_label.pack_forget()
        self.title(f"{APP_TITLE} \u2014 {os.path.basename(self._db.db_path)}")

    # ── Date list ─────────────────────────────────────────────────────────────

    def _refresh_date_list(self, select_date: str = None):
        if not self._db:
            return
        self._date_listbox.delete(0, "end")
        dates = self._db.get_all_dates()
        for d in dates:
            self._date_listbox.insert("end", d)

        if not dates:
            return

        if select_date and select_date in dates:
            idx = dates.index(select_date)
        elif self._current_entry_id:
            row = self._db.get_entry_by_id(self._current_entry_id)
            idx = dates.index(row["entry_date"]) if row and row["entry_date"] in dates else 0
        else:
            idx = 0

        self._date_listbox.selection_set(idx)
        self._date_listbox.see(idx)
        self._load_entry_by_date(dates[idx])

    # ── Entry loading ─────────────────────────────────────────────────────────

    def _load_entry_by_date(self, date_str: str):
        if not self._db:
            return
        if self._unsaved_changes:
            if not messagebox.askyesno("Unsaved Changes",
                    "Discard unsaved changes and load a different entry?"):
                return
        entry = self._db.get_entry_by_date(date_str)
        if not entry:
            return

        self._current_entry_id = entry["id"]
        self._date_display_var.set(self._fmt_date(date_str))

        self._loading = True
        self._notes_var.set(entry["notes"] or "")
        self._loading = False

        self._tree.delete(*self._tree.get_children())
        for item in self._db.get_work_items(entry["id"]):
            h = item["hours"]
            self._tree.insert("", "end",
                values=(item["task"], item["reason"],
                        str(int(h)) if h == int(h) else f"{h:.1f}"))

        self._unsaved_changes = False
        self._unsaved_label.pack_forget()
        self._save_btn.config(state="normal")
        self._update_total()
        log.debug("Loaded entry %s (id=%d)", date_str, entry["id"])

    # ── Left-panel events ─────────────────────────────────────────────────────

    def _on_date_selected(self, event=None):
        sel = self._date_listbox.curselection()
        if not sel:
            return
        date_str = self._date_listbox.get(sel[0])
        if self._current_entry_id and self._db:
            row = self._db.get_entry_by_id(self._current_entry_id)
            if row and row["entry_date"] == date_str:
                return
        self._load_entry_by_date(date_str)

    def _on_add_entry(self):
        if not self._db:
            messagebox.showinfo("No Database",
                "Configure a database in Settings first.")
            return
        dlg = _DateInputDialog(self, "New Entry \u2014 Pick a Date")
        new_date = dlg.result
        if not new_date:
            return
        if self._db.get_entry_by_date(new_date):
            messagebox.showinfo("Entry Exists",
                f"An entry for {new_date} already exists.")
            self._select_date(new_date)
            return
        try:
            self._db.create_entry(new_date)
            self._refresh_date_list(select_date=new_date)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _on_delete_entry(self):
        sel = self._date_listbox.curselection()
        if not sel:
            return
        date_str = self._date_listbox.get(sel[0])
        if not messagebox.askyesno("Confirm Delete",
                f"Delete the entry for {date_str} and all its work items?\n\nThis cannot be undone.",
                icon="warning"):
            return
        entry = self._db.get_entry_by_date(date_str)
        if entry:
            self._db.delete_entry(entry["id"])
        self._current_entry_id = None
        self._unsaved_changes = False
        self._unsaved_label.pack_forget()
        self._date_display_var.set("\u2014 no entry selected \u2014")
        self._notes_var.set("")
        self._tree.delete(*self._tree.get_children())
        self._save_btn.config(state="disabled")
        self._update_total()
        self._refresh_date_list()
        self._regenerate_html()  # Silent on delete — no entry to confirm saved
        self._set_status("\u2713 Entry deleted \u2014 HTML report updated.", color="#2a7a2a")

    # ── Work-item row editing ─────────────────────────────────────────────────

    def _on_add_row(self):
        if not self._current_entry_id:
            return
        self._commit_inline_edit()
        iid = self._tree.insert("", "end", values=("", "", "0"))
        self._mark_unsaved()
        self._update_total()
        self._tree.selection_set(iid)
        self._tree.see(iid)
        self._start_inline_edit(iid, 0)

    def _on_remove_row(self):
        sel = self._tree.selection()
        if not sel:
            return
        self._commit_inline_edit()
        self._tree.delete(sel[0])
        self._mark_unsaved()
        self._update_total()

    def _on_tree_double_click(self, event=None):
        if event and hasattr(event, "x"):
            if self._tree.identify_region(event.x, event.y) != "cell":
                return
            iid = self._tree.identify_row(event.y)
            col_idx = int(self._tree.identify_column(event.x).replace("#", "")) - 1
        else:
            sel = self._tree.selection()
            if not sel:
                return
            iid, col_idx = sel[0], 0
        if iid:
            self._start_inline_edit(iid, col_idx)

    def _start_inline_edit(self, iid: str, col_idx: int):
        self._commit_inline_edit()
        col_id = COLUMNS[col_idx][0]
        bbox = self._tree.bbox(iid, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        vals = list(self._tree.item(iid, "values"))
        var = tk.StringVar(value=vals[col_idx] if col_idx < len(vals) else "")
        widget = ttk.Entry(self._tree, textvariable=var, font=("", 10))
        widget.place(x=x, y=y, width=w, height=h)
        widget.focus_set()
        widget.select_range(0, "end")
        self._edit_widgets = {"widget": widget, "var": var,
                              "iid": iid, "col_idx": col_idx}

        def commit(e=None): self._commit_inline_edit()
        def tab_next(e=None):
            self._commit_inline_edit()
            next_col = (col_idx + 1) % len(COLUMNS)
            if next_col == 0:
                children = self._tree.get_children()
                idx = list(children).index(iid)
                if idx + 1 < len(children):
                    self._start_inline_edit(children[idx + 1], 0)
                else:
                    self._on_add_row()
            else:
                self._start_inline_edit(iid, next_col)
            return "break"

        widget.bind("<Return>",   commit)
        widget.bind("<Tab>",      tab_next)
        widget.bind("<FocusOut>", commit)
        widget.bind("<Escape>",   lambda e: self._cancel_inline_edit())

    def _commit_inline_edit(self):
        if not self._edit_widgets:
            return
        w = self._edit_widgets.get("widget")
        if not w or not w.winfo_exists():
            self._edit_widgets = {}
            return
        iid     = self._edit_widgets["iid"]
        col_idx = self._edit_widgets["col_idx"]
        val     = self._edit_widgets["var"].get()
        w.destroy()
        self._edit_widgets = {}

        if col_idx == 2:  # hours
            try:
                parsed = max(0.0, float(val))
                val = str(int(parsed)) if parsed == int(parsed) else f"{parsed:.1f}"
            except ValueError:
                val = "0"

        vals = list(self._tree.item(iid, "values"))
        while len(vals) < len(COLUMNS):
            vals.append("")
        vals[col_idx] = val
        self._tree.item(iid, values=vals)
        self._mark_unsaved()
        self._update_total()

    def _cancel_inline_edit(self):
        w = self._edit_widgets.get("widget")
        if w and w.winfo_exists():
            w.destroy()
        self._edit_widgets = {}

    # ── Save & HTML ───────────────────────────────────────────────────────────

    def _on_save(self):
        if not self._db or not self._current_entry_id:
            return
        self._commit_inline_edit()
        items = []
        for iid in self._tree.get_children():
            v = self._tree.item(iid, "values")
            try:
                hours = float(v[2]) if len(v) > 2 else 0.0
            except ValueError:
                hours = 0.0
            items.append({
                "task":   v[0] if len(v) > 0 else "",
                "reason": v[1] if len(v) > 1 else "",
                "hours":  hours,
            })
        if self._config.get("backup_enabled", True):
            self._db.backup()
        try:
            self._db.update_entry_notes(self._current_entry_id,
                                        self._notes_var.get().strip())
            self._db.replace_work_items(self._current_entry_id, items)
        except Exception as exc:
            messagebox.showerror("Save Error", f"Could not save:\n{exc}")
            log.error("Save failed: %s", exc)
            return
        self._unsaved_changes = False
        self._unsaved_label.pack_forget()
        self._refresh_date_list()
        html_ok, html_msg = self._regenerate_html()
        if html_ok:
            self._set_status(f"\u2713 Entry saved \u2014 HTML report updated.", color="#2a7a2a")
        else:
            self._set_status(f"\u2713 Entry saved \u2014 HTML: {html_msg}", color="#b05000")
        log.info("Entry saved: id=%d  html=%s", self._current_entry_id, html_msg)

    def _regenerate_html(self):
        """Regenerate the HTML report. Returns (success: bool, message: str)."""
        if not self._db:
            return False, "no database"
        html_path = self._config.get("html_path", "").strip()
        if not html_path:
            return False, "HTML path not set — configure it in Settings"
        try:
            entries = self._db.get_all_entries_with_items()
            ok = generate_html(entries, html_path,
                               author_name=self._config.get("author_name", ""),
                               author_team=self._config.get("author_team", ""),
                               author_org=self._config.get("author_org", ""))
            if ok:
                return True, html_path
            else:
                return False, "write failed — check logs/worklog.log"
        except Exception as exc:
            log.error("HTML generation failed: %s", exc)
            return False, str(exc)

    def _set_status(self, message: str, color: str = "#555", timeout_ms: int = 6000):
        """Show a temporary status message below the Save button."""
        self._status_var.set(message)
        self._status_label.configure(foreground=color)
        # Clear after timeout_ms so it doesn't linger forever
        self.after(timeout_ms, lambda: self._status_var.set(""))

    def _open_html_in_browser(self):
        html_path = self._config.get("html_path", "").strip()
        if html_path and os.path.isfile(html_path):
            webbrowser.open(f"file://{os.path.abspath(html_path)}")
        else:
            messagebox.showinfo("No Report",
                "HTML report not found. Save an entry first.")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _mark_unsaved(self):
        if not self._unsaved_changes:
            self._unsaved_changes = True
            self._unsaved_label.pack(side="left", padx=(12, 0))

    def _on_field_change(self, *_):
        if self._current_entry_id and not self._loading:
            self._mark_unsaved()

    def _update_total(self):
        total = 0.0
        for iid in self._tree.get_children():
            v = self._tree.item(iid, "values")
            try:
                total += float(v[2]) if len(v) > 2 else 0.0
            except (ValueError, IndexError):
                pass
        self._total_var.set(f"Daily total: {total:.1f} hrs")

    def _fmt_date(self, iso: str) -> str:
        try:
            d = datetime.strptime(iso, "%Y-%m-%d")
            return f"{d.strftime('%A, %B')} {d.day}, {d.year}"
        except ValueError:
            return iso

    def _select_date(self, date_str: str):
        for i in range(self._date_listbox.size()):
            if self._date_listbox.get(i) == date_str:
                self._date_listbox.selection_clear(0, "end")
                self._date_listbox.selection_set(i)
                self._date_listbox.see(i)
                self._load_entry_by_date(date_str)
                return

    # ── Dialog launchers ──────────────────────────────────────────────────────

    def _open_settings(self):
        def on_save(cfg):
            self._config = cfg
            new_path = cfg.get("db_path", "")
            cur_path = self._db.db_path if self._db else ""
            if new_path and new_path != cur_path:
                self._connect_database(new_path)
        SettingsDialog(self, self._config, on_save_callback=on_save)

    def _open_calendar(self):
        if not self._db:
            return
        entry_dates = set(self._db.get_all_dates())
        sel = self._date_listbox.curselection()
        initial = self._date_listbox.get(sel[0]) if sel else None

        def on_pick(date_str):
            if date_str in entry_dates:
                self._select_date(date_str)
            elif messagebox.askyesno("No Entry",
                    f"No entry for {date_str}.\nCreate one?"):
                try:
                    self._db.create_entry(date_str)
                    self._refresh_date_list(select_date=date_str)
                except Exception as exc:
                    messagebox.showerror("Error", str(exc))

        CalendarPopup(self, entry_dates, on_pick, initial_date=initial)

    def _open_merge(self):
        if not self._db:
            messagebox.showinfo("No Database",
                "Open a database before merging.")
            return
        open_merge_dialog(self, self._db,
                          on_complete=lambda: self._refresh_date_list())

    def _show_about(self):
        messagebox.showinfo(f"About {APP_TITLE}",
            f"{APP_TITLE} v{APP_VERSION}\n\n"
            "A daily work logging tool that stores entries in SQLite\n"
            "and auto-generates a shareable HTML report on every save.\n\n"
            "Check logs/worklog.log if you experience any issues.")

    def _on_close(self):
        if self._unsaved_changes:
            if not messagebox.askyesno("Unsaved Changes",
                    "Exit without saving?"):
                return
        self._commit_inline_edit()
        if self._db:
            self._db.close()
        log.info("Application closed.")
        self.destroy()
        sys.exit(0)


# ── Date-picker dialog ────────────────────────────────────────────────────────

class _DateInputDialog(tk.Toplevel):
    """Simple modal date-picker. Uses tkcalendar if available."""

    def __init__(self, parent, title="Select Date"):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.result: Optional[str] = None
        self._build(parent)
        self._center(parent)
        self.wait_window()

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _build(self, parent):
        f = ttk.Frame(self, padding=16)
        f.pack()
        try:
            from tkcalendar import Calendar as Cal
            today = date.today()
            self._cal = Cal(f, selectmode="day",
                            year=today.year, month=today.month, day=today.day,
                            date_pattern="yyyy-mm-dd",
                            background="#1a2744", foreground="white",
                            selectbackground="#c8a951", selectforeground="#1a2744")
            self._cal.pack(padx=6, pady=6)
            self._get = self._cal.get_date
        except ImportError:
            ttk.Label(f, text="Date (YYYY-MM-DD):").pack()
            self._var = tk.StringVar(value=date.today().isoformat())
            ttk.Entry(f, textvariable=self._var, width=14,
                      justify="center").pack(pady=6)
            self._get = self._var.get

        bf = ttk.Frame(f)
        bf.pack(pady=(8, 0))
        ttk.Button(bf, text="Select", command=self._ok, width=10).pack(
            side="left", padx=4)
        ttk.Button(bf, text="Cancel", command=self.destroy, width=8).pack(
            side="left", padx=4)

    def _ok(self):
        val = self._get().strip()
        try:
            parts = val.split("-")
            date(int(parts[0]), int(parts[1]), int(parts[2]))
            self.result = val
            self.destroy()
        except Exception:
            messagebox.showerror("Invalid Date",
                "Please enter a valid date.", parent=self)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = WorkLogApp()
    app.mainloop()
