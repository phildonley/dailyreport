"""
calendar_popup.py
-----------------
Popup calendar for jumping to a specific date.

Requires the external package tkcalendar (listed in requirements.txt).
If tkcalendar is not installed a minimal text-input fallback is shown
so the app still works, just less conveniently.

Dates that HAVE entries in the database are highlighted in blue.
All other dates are still selectable (to create a new entry on that date).

Usage
-----
    CalendarPopup(
        parent=self,
        entry_dates={"2026-04-13", "2026-04-10", ...},
        on_date_selected=my_callback,   # called with a 'YYYY-MM-DD' string
        initial_date="2026-04-13",      # optional — where the calendar opens
    )
"""

import tkinter as tk
from tkinter import messagebox, ttk
from datetime import date

from logger_setup import get_logger

log = get_logger(__name__)

# Try to import tkcalendar; fall back gracefully if missing
try:
    from tkcalendar import Calendar as TkCalendar
    _HAS_TKCALENDAR = True
    log.debug("tkcalendar is available.")
except ImportError:
    _HAS_TKCALENDAR = False
    log.warning(
        "tkcalendar is not installed. "
        "Calendar popup will use a plain text-entry fallback. "
        "Run:  pip install tkcalendar"
    )


class CalendarPopup(tk.Toplevel):
    """
    Modal popup calendar.

    Parameters
    ----------
    parent            : parent Tkinter widget
    entry_dates       : iterable of 'YYYY-MM-DD' strings that have entries
    on_date_selected  : callable(date_str: str) — called when the user picks a date
    initial_date      : 'YYYY-MM-DD' string — where the calendar opens (today if None)
    """

    def __init__(
        self,
        parent: tk.Widget,
        entry_dates,
        on_date_selected,
        initial_date: str = None,
    ):
        super().__init__(parent)
        self.title("Jump to Date — WorkLog")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._entry_dates = set(entry_dates)
        self._on_date_selected = on_date_selected

        if _HAS_TKCALENDAR:
            self._build_full_calendar(initial_date)
        else:
            self._build_fallback(initial_date)

        self._center(parent)
        self.wait_window()

    # ── Full calendar (tkcalendar) ────────────────────────────────────────────

    def _build_full_calendar(self, initial_date: str):
        today = date.today()
        year, month, day = today.year, today.month, today.day

        if initial_date:
            try:
                parts = initial_date.split("-")
                year, month, day = int(parts[0]), int(parts[1]), int(parts[2])
            except (ValueError, IndexError):
                pass  # Fall back to today

        frame = ttk.Frame(self, padding=14)
        frame.pack(fill="both", expand=True)

        self._cal = TkCalendar(
            frame,
            selectmode="day",
            year=year, month=month, day=day,
            date_pattern="yyyy-mm-dd",   # get_date() returns YYYY-MM-DD
            # Colour scheme to match the app's navy/gold palette
            background="#1a2744",
            foreground="white",
            headersbackground="#1a2744",
            headersforeground="#c8a951",
            selectbackground="#c8a951",
            selectforeground="#1a2744",
            normalbackground="white",
            normalforeground="#222",
            weekendbackground="#f4f5f8",
            weekendforeground="#444",
            othermonthbackground="#e8e8e8",
            othermonthforeground="#aaa",
            font=("Segoe UI", 10),
        )
        self._cal.pack(padx=6, pady=6)

        # Highlight dates that already have entries
        for date_str in self._entry_dates:
            try:
                p = date_str.split("-")
                self._cal.calevent_create(
                    date(int(p[0]), int(p[1]), int(p[2])),
                    "Entry",
                    tags="has_entry",
                )
            except (ValueError, IndexError):
                pass  # Ignore malformed dates

        self._cal.tag_config(
            "has_entry",
            background="#d0e4ff",
            foreground="#1a2744",
        )

        # Legend
        legend = ttk.Frame(frame)
        legend.pack(pady=(0, 8))
        tk.Label(legend, bg="#d0e4ff", width=2, relief="solid",
                 borderwidth=1).pack(side="left", padx=(0, 5))
        ttk.Label(legend, text="= date has an entry").pack(side="left")

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(6, 2))
        ttk.Button(
            btn_frame, text="Go to Date", command=self._on_select_full, width=14
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame, text="Cancel", command=self.destroy, width=10
        ).pack(side="left", padx=4)

    def _on_select_full(self):
        selected = self._cal.get_date()   # 'YYYY-MM-DD' per date_pattern
        log.info("Calendar: user selected %s", selected)
        self.destroy()
        self._on_date_selected(selected)

    # ── Fallback (no tkcalendar) ─────────────────────────────────────────────

    def _build_fallback(self, initial_date: str):
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="tkcalendar is not installed.\nEnter the date manually:",
            justify="center",
            font=("", 10),
        ).pack(pady=(0, 12))

        default = initial_date or date.today().isoformat()
        self._fallback_var = tk.StringVar(value=default)
        entry = ttk.Entry(frame, textvariable=self._fallback_var, width=16,
                          font=("", 12), justify="center")
        entry.pack()
        entry.bind("<Return>", lambda e: self._on_select_fallback())

        ttk.Label(frame, text="Format: YYYY-MM-DD", foreground="#888",
                  font=("", 9)).pack(pady=4)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=14)
        ttk.Button(
            btn_frame, text="Go", command=self._on_select_fallback, width=10
        ).pack(side="left", padx=4)
        ttk.Button(
            btn_frame, text="Cancel", command=self.destroy, width=10
        ).pack(side="left", padx=4)

    def _on_select_fallback(self):
        val = self._fallback_var.get().strip()
        try:
            parts = val.split("-")
            if len(parts) != 3:
                raise ValueError
            date(int(parts[0]), int(parts[1]), int(parts[2]))  # Validate
        except (ValueError, IndexError):
            messagebox.showerror(
                "Invalid Date",
                "Please enter a date in YYYY-MM-DD format.\nExample: 2026-04-13",
                parent=self,
            )
            return
        log.info("Fallback calendar: user entered %s", val)
        self.destroy()
        self._on_date_selected(val)

    # ── Utility ───────────────────────────────────────────────────────────────

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
