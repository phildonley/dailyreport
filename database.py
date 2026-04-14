"""
database.py
-----------
All SQLite operations for WorkLog.

Responsibilities
----------------
- Create and validate the schema
- CRUD for daily_entries and work_items
- Daily backup before save
- WAL journal mode for safer writes on network drives
- Static schema-validation helper used by the merge dialog

Schema summary
--------------
  app_metadata    key/value store — identifies this as a WorkLog DB
  daily_entries   one row per calendar day (entry_date is UNIQUE)
  work_items      multiple rows per daily entry, linked by entry_id FK
"""

import os
import shutil
import sqlite3
import time
from datetime import datetime, date
from typing import List, Optional

from logger_setup import get_logger

log = get_logger(__name__)

# ── Identity constants ──────────────────────────────────────────────────────────
SCHEMA_VERSION = "1"
APP_NAME = "WorkLog"

# ── DDL statements ──────────────────────────────────────────────────────────────
_SQL_CREATE_METADATA = """
CREATE TABLE IF NOT EXISTS app_metadata (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);
"""

_SQL_CREATE_DAILY_ENTRIES = """
CREATE TABLE IF NOT EXISTS daily_entries (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_date   TEXT    NOT NULL UNIQUE,   -- ISO-8601: YYYY-MM-DD
    notes        TEXT    NOT NULL DEFAULT '',  -- Optional daily note
    created_at   TEXT    NOT NULL,             -- ISO-8601 UTC datetime
    updated_at   TEXT    NOT NULL              -- Bumped on every save
);
"""

_SQL_CREATE_WORK_ITEMS = """
CREATE TABLE IF NOT EXISTS work_items (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    entry_id     INTEGER NOT NULL,            -- FK -> daily_entries.id
    sort_order   INTEGER NOT NULL DEFAULT 0,  -- Preserves UI row order
    task         TEXT    NOT NULL DEFAULT '',
    reason       TEXT    NOT NULL DEFAULT '',
    hours        REAL    NOT NULL DEFAULT 0.0, -- Supports 0.5, 1.25, etc.
    created_at   TEXT    NOT NULL,
    updated_at   TEXT    NOT NULL,
    FOREIGN KEY (entry_id) REFERENCES daily_entries(id) ON DELETE CASCADE
);
"""

_SQL_CREATE_INDEX = """
CREATE INDEX IF NOT EXISTS idx_work_items_entry_id
    ON work_items (entry_id);
"""


# ── Helpers ─────────────────────────────────────────────────────────────────────

def _now() -> str:
    """Current UTC time as ISO-8601 string (no microseconds)."""
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S")


def _today() -> str:
    """Today's local date as YYYY-MM-DD."""
    return date.today().isoformat()


# ── Database class ───────────────────────────────────────────────────────────────

class Database:
    """
    Manages a single SQLite connection for the app's lifetime.

    Usage:
        db = Database("/path/to/worklog.db")  # opens or creates
        ...                                    # use its methods
        db.close()                             # call on app exit

    The constructor calls _initialize_schema(), which is safe to call on
    both new and existing databases (uses CREATE TABLE IF NOT EXISTS).
    """

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._conn: Optional[sqlite3.Connection] = None
        log.info("Opening database: %s", db_path)
        self._connect()
        self._cleanup_wal_files()
        self._initialize_schema()

    # ── Internal helpers ────────────────────────────────────────────────────────

    def _connect(self):
        """Open the SQLite connection with safe settings."""
        self._conn = sqlite3.connect(
            self.db_path,
            timeout=10,              # Wait up to 10 s if file is locked (network drives)
            check_same_thread=False, # We only ever use one thread, but this is safer
        )
        self._conn.row_factory = sqlite3.Row  # Rows accessible by column name

        # Use DELETE journal mode (SQLite default) — most compatible with
        # network drives. WAL mode requires file locking support that many
        # SMB/CIFS shares do not provide reliably, causing "locking protocol"
        # errors. DELETE mode works correctly on all network filesystems.
        self._conn.execute("PRAGMA journal_mode=DELETE;")

        # Enforce ON DELETE CASCADE and other FK constraints.
        self._conn.execute("PRAGMA foreign_keys=ON;")

        # Increase busy timeout — on a network drive, another process or the
        # OS itself may briefly hold a lock. Wait up to 15 seconds before
        # giving up with "database is locked".
        self._conn.execute("PRAGMA busy_timeout=15000;")

        self._conn.commit()
        log.debug("SQLite opened (DELETE journal mode, 15s busy timeout).")

    def _cleanup_wal_files(self):
        """
        Remove stale WAL/SHM sidecar files left by a previous WAL-mode session.
        These cause locking errors if the journal mode is now DELETE.
        Safe to call on any database — does nothing if the files don't exist.
        """
        for suffix in ("-wal", "-shm"):
            path = self.db_path + suffix
            if os.path.exists(path):
                try:
                    os.remove(path)
                    log.info("Removed stale WAL sidecar file: %s", path)
                except OSError as exc:
                    log.warning("Could not remove %s: %s", path, exc)

    def _initialize_schema(self):
        """Create tables and seed metadata on first run."""
        cur = self._conn.cursor()
        cur.execute(_SQL_CREATE_METADATA)
        cur.execute(_SQL_CREATE_DAILY_ENTRIES)
        cur.execute(_SQL_CREATE_WORK_ITEMS)
        cur.execute(_SQL_CREATE_INDEX)

        # Only seed metadata when the table is brand new
        cur.execute("SELECT value FROM app_metadata WHERE key = 'app_name'")
        if cur.fetchone() is None:
            cur.executemany(
                "INSERT OR IGNORE INTO app_metadata (key, value) VALUES (?, ?)",
                [
                    ("app_name",       APP_NAME),
                    ("schema_version", SCHEMA_VERSION),
                    ("created_at",     _now()),
                ],
            )
            log.info("New database: metadata seeded.")

        self._conn.commit()
        log.debug("Schema initialization complete.")

    def _write_with_retry(self, fn, max_attempts: int = 4, base_delay: float = 0.5):
        """
        Call fn() (a zero-argument callable that performs SQLite writes).
        On transient OperationalError (disk I/O error, locking protocol) the
        call is retried up to max_attempts times with exponential back-off
        (0.5 s, 1 s, 2 s, 4 s).  Non-transient errors are re-raised immediately.
        """
        _TRANSIENT = ("i/o error", "disk i/o", "locking protocol", "unable to open")
        last_exc: Optional[Exception] = None
        for attempt in range(max_attempts):
            try:
                return fn()
            except sqlite3.OperationalError as exc:
                msg = str(exc).lower()
                if any(k in msg for k in _TRANSIENT):
                    last_exc = exc
                    delay = base_delay * (2 ** attempt)
                    log.warning(
                        "Transient DB error (attempt %d/%d): %s — retrying in %.1f s",
                        attempt + 1, max_attempts, exc, delay,
                    )
                    time.sleep(delay)
                else:
                    raise  # Non-transient — fail immediately
        log.error("All %d write attempts failed: %s", max_attempts, last_exc)
        raise last_exc  # type: ignore[misc]

    # ── Connection management ────────────────────────────────────────────────────

    def close(self):
        """Close the database connection cleanly."""
        if self._conn:
            self._conn.close()
            self._conn = None
            log.info("Database connection closed.")

    # ── Schema validation (used by merge dialog) ─────────────────────────────────

    @staticmethod
    def validate_file(db_path: str):
        """
        Check whether a file on disk is a valid WorkLog database.
        Opens a read-only connection; does not modify anything.

        Returns:
            (True, "")              — valid WorkLog database
            (False, "reason string") — not valid, with explanation
        """
        if not os.path.isfile(db_path):
            return False, "File does not exist."
        try:
            conn = sqlite3.connect(db_path, timeout=5)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()

            # Check all required tables are present
            cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
            tables = {row["name"] for row in cur.fetchall()}
            required = {"app_metadata", "daily_entries", "work_items"}
            missing = required - tables
            if missing:
                conn.close()
                return False, f"Missing tables: {missing}"

            # Check the app_name metadata value
            cur.execute("SELECT value FROM app_metadata WHERE key='app_name'")
            row = cur.fetchone()
            conn.close()

            if row is None or row["value"] != APP_NAME:
                return False, "Not a WorkLog database (app_name mismatch)."
            return True, ""

        except sqlite3.DatabaseError as exc:
            return False, f"SQLite error: {exc}"

    # ── Backup ──────────────────────────────────────────────────────────────────

    def backup(self) -> Optional[str]:
        """
        Copy the .db file to worklog_backup_YYYYMMDD.db in the same directory.
        At most one backup per calendar day (same-day backup is overwritten).

        Returns the backup path on success, None on failure.
        Failure is non-fatal — the save will proceed regardless.
        """
        today_str = _today().replace("-", "")  # e.g. 20260413
        backup_name = f"worklog_backup_{today_str}.db"
        backup_path = os.path.join(os.path.dirname(self.db_path), backup_name)
        try:
            shutil.copy2(self.db_path, backup_path)
            log.info("Backup created: %s", backup_path)
            return backup_path
        except Exception as exc:
            log.warning("Backup failed (save will continue): %s", exc)
            return None

    # ── Daily entries ────────────────────────────────────────────────────────────

    def get_all_dates(self) -> List[str]:
        """Return every entry_date (YYYY-MM-DD), newest first."""
        cur = self._conn.cursor()
        cur.execute(
            "SELECT entry_date FROM daily_entries ORDER BY entry_date DESC"
        )
        return [row["entry_date"] for row in cur.fetchall()]

    def get_entry_by_date(self, entry_date: str) -> Optional[sqlite3.Row]:
        """Return the daily_entries row for the given date, or None."""
        cur = self._conn.cursor()
        cur.execute(
            "SELECT * FROM daily_entries WHERE entry_date = ?", (entry_date,)
        )
        return cur.fetchone()

    def get_entry_by_id(self, entry_id: int) -> Optional[sqlite3.Row]:
        """Return the daily_entries row for the given ID, or None."""
        cur = self._conn.cursor()
        cur.execute("SELECT * FROM daily_entries WHERE id = ?", (entry_id,))
        return cur.fetchone()

    def create_entry(self, entry_date: str, notes: str = "") -> int:
        """
        Insert a new daily_entries row.
        Raises sqlite3.IntegrityError if entry_date already exists.
        Returns the new row ID.
        """
        now = _now()
        cur = self._conn.cursor()
        cur.execute(
            "INSERT INTO daily_entries (entry_date, notes, created_at, updated_at) "
            "VALUES (?, ?, ?, ?)",
            (entry_date, notes, now, now),
        )
        self._conn.commit()
        entry_id = cur.lastrowid
        log.info("Created entry id=%d date=%s", entry_id, entry_date)
        return entry_id

    def update_entry_notes(self, entry_id: int, notes: str):
        """Update the daily note and bump updated_at (retries on transient I/O errors)."""
        def _do():
            self._conn.execute(
                "UPDATE daily_entries SET notes = ?, updated_at = ? WHERE id = ?",
                (notes, _now(), entry_id),
            )
            self._conn.commit()
        self._write_with_retry(_do)

    def delete_entry(self, entry_id: int):
        """
        Delete a daily entry.  ON DELETE CASCADE removes all its work_items.
        """
        self._conn.execute(
            "DELETE FROM daily_entries WHERE id = ?", (entry_id,)
        )
        self._conn.commit()
        log.info("Deleted entry id=%d (and its work items)", entry_id)

    # ── Work items ───────────────────────────────────────────────────────────────

    def get_work_items(self, entry_id: int) -> List[sqlite3.Row]:
        """Return all work items for an entry, ordered by sort_order."""
        cur = self._conn.cursor()
        cur.execute(
            "SELECT * FROM work_items WHERE entry_id = ? ORDER BY sort_order ASC",
            (entry_id,),
        )
        return cur.fetchall()

    def replace_work_items(self, entry_id: int, items: List[dict]):
        """
        Atomically replace ALL work items for an entry.

        items: list of dicts, each with keys 'task', 'reason', 'hours'.

        Uses a context-manager transaction — if anything fails the entire
        operation is rolled back and the database is unchanged.
        """
        now = _now()

        def _do():
            with self._conn:  # auto-COMMIT on exit, auto-ROLLBACK on exception
                # Delete all existing rows for this entry
                self._conn.execute(
                    "DELETE FROM work_items WHERE entry_id = ?", (entry_id,)
                )
                # Re-insert in the order provided (sort_order = list index)
                for order, item in enumerate(items):
                    self._conn.execute(
                        "INSERT INTO work_items "
                        "(entry_id, sort_order, task, reason, hours, created_at, updated_at) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        (
                            entry_id,
                            order,
                            item.get("task",   "").strip(),
                            item.get("reason", "").strip(),
                            float(item.get("hours", 0.0)),
                            now,
                            now,
                        ),
                    )
                # Bump the parent entry's updated_at timestamp
                self._conn.execute(
                    "UPDATE daily_entries SET updated_at = ? WHERE id = ?",
                    (now, entry_id),
                )

        try:
            self._write_with_retry(_do)
            log.info("Saved %d work item(s) for entry_id=%d", len(items), entry_id)
        except sqlite3.Error as exc:
            log.error("Failed to save work items for entry_id=%d: %s", entry_id, exc)
            raise

    # ── Combined read for HTML generation ────────────────────────────────────────

    def get_all_entries_with_items(self) -> List[dict]:
        """
        Return every entry together with its work items.
        This is the data source for the HTML generator.

        Returns a list of dicts (newest entry first):
          {
            'id':         int,
            'date':       'YYYY-MM-DD',
            'notes':      str,
            'created_at': str,
            'updated_at': str,
            'items': [
                {'task': str, 'reason': str, 'hours': float}, ...
            ]
          }
        """
        cur = self._conn.cursor()
        cur.execute("SELECT * FROM daily_entries ORDER BY entry_date DESC")
        entries = cur.fetchall()

        result = []
        for entry in entries:
            items = self.get_work_items(entry["id"])
            result.append({
                "id":         entry["id"],
                "date":       entry["entry_date"],
                "notes":      entry["notes"],
                "created_at": entry["created_at"],
                "updated_at": entry["updated_at"],
                "items": [
                    {
                        "task":   item["task"],
                        "reason": item["reason"],
                        "hours":  item["hours"],
                    }
                    for item in items
                ],
            })
        return result
