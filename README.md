# WorkLog Manager

WorkLog Manager is a lightweight Tkinter desktop application for logging your daily work activities. Each day you record what tasks you performed, why you did them, and how many hours each took. The app stores everything in a local SQLite database and automatically generates a polished, self-contained HTML report that you can open in any browser, email to your manager, or save to a shared network folder. The interface keeps a dated list of all your entries on the left, with an editable work-item grid on the right, making it quick to review past days or add new ones.

---

## Requirements

- **Python 3.8 or later**
- **tkcalendar** (for the calendar date-picker popup)

```
pip install tkcalendar
```

Tkinter is included with most Python distributions. If it is missing on Linux, install it via your package manager (e.g. `sudo apt install python3-tk`).

---

## File Structure

| File | Purpose |
|---|---|
| `main.py` | Application entry point. Defines `WorkLogApp` (the main window) and `_DateInputDialog`. Run this file to start the app. |
| `database.py` | All SQLite operations — schema creation, CRUD for entries and work items, daily backup, and the static schema-validation helper used by the merge dialog. |
| `html_generator.py` | Generates the self-contained HTML report from all database entries. Includes embedded CSS and JavaScript — no internet connection required. |
| `settings_manager.py` | Reads and writes `config/config.json`. Also provides the Settings dialog where you set your database path, HTML output path, and branding fields. |
| `logger_setup.py` | Configures the Python `logging` module. Writes logs to `logs/worklog.log` and also streams them to the console during development. |
| `calendar_popup.py` | Modal calendar popup (requires tkcalendar). Highlights dates that already have entries. Falls back gracefully to a plain text-entry field if tkcalendar is not installed. |
| `merge_dialog.py` | Dialog and logic for merging a second WorkLog database into the currently open one. Supports three conflict strategies: append, skip, or replace. |
| `requirements.txt` | Lists the single third-party dependency: `tkcalendar>=1.6.1`. |

---

## How to Run

```
python main.py
```

The application window will open. On the very first run it will detect that no database path has been configured and prompt you to open Settings.

---

## First-Run Steps

1. Click **Settings** in the top-right corner (or go to **File -> Settings**).
2. Under **Database File**, click **Browse...** and choose a location for your `.db` file (e.g. `C:\Users\You\Documents\worklog.db` or `/home/you/worklog.db`). The file will be created automatically if it does not exist yet.
3. Under **HTML Report Output File**, click **Browse...** and choose where the HTML report should be written (e.g. `worklog_report.html` on a shared drive, or anywhere convenient).
4. Fill in your **Name**, **Team / Group**, and **Organization** so the HTML report is properly branded.
5. Click **Save**. The app will connect to the database immediately.

---

## How the Database Works

WorkLog uses **SQLite** with **WAL (Write-Ahead Logging)** journal mode. WAL reduces the risk of "database is locked" errors and allows safe concurrent reads, which is important when the `.db` file is stored on a network drive.

The database has three tables:

| Table | Purpose |
|---|---|
| `app_metadata` | Key/value store that identifies the file as a WorkLog database and stores the schema version. |
| `daily_entries` | One row per calendar day. Stores the date (`YYYY-MM-DD`), an optional notes string, and timestamps. |
| `work_items` | Multiple rows per daily entry. Each row stores a task description, reason/context, and hours. Linked to `daily_entries` by a foreign key with `ON DELETE CASCADE`. |

**Network drive note:** SQLite works on network drives but is sensitive to network interruptions. The built-in daily backup (see below) and WAL mode reduce the risk of data loss, but storing the `.db` file locally and syncing it via OneDrive/Dropbox/etc. is generally more reliable than a raw network share (UNC path).

---

## How the HTML Report Works

Every time you click **Save**, the app automatically regenerates the HTML report at the path you configured in Settings. The report is a single `.html` file with all CSS and JavaScript embedded -- no internet connection is needed to view it.

Features of the report:
- Entries sorted newest first
- Collapsible day sections (click a date header to expand/collapse)
- Live search bar to filter by keyword
- Hours total per day shown in each section header
- Grand total hours shown in the page header
- Print-friendly via `@media print` CSS

To open the report in your default browser, go to **View -> Open HTML in Browser**.

To manually regenerate the report without saving, go to **View -> Regenerate HTML**.

---

## How to Use the Calendar Button

Click the **Calendar** button in the left-panel toolbar to open a month-view date picker. Dates that already have entries are highlighted in blue. Click any date and then **Go to Date** to jump to it. If the selected date does not yet have an entry, the app will offer to create one for you.

---

## How to Merge Databases

If you have been running WorkLog on two different machines (e.g. a laptop and a desktop), you can combine the two databases:

1. Go to **File -> Merge Database**.
2. Browse to the **source** `.db` file (the one you want to import from).
3. Click **Preview** to see how many new dates will be imported and how many dates conflict (exist in both databases).
4. Choose a conflict strategy:
   - **Append** -- add the incoming work items as extra rows on the existing day (safe, non-destructive; recommended)
   - **Skip** -- keep your existing entry and ignore the incoming one
   - **Replace** -- overwrite your existing entry with the incoming one
5. Click **Merge**. A backup of your current database is created before any changes are made. The entire merge runs in a single transaction -- if anything goes wrong your database is left completely unchanged.

---

## How to Package as an Executable (.exe)

To distribute WorkLog as a standalone Windows executable that does not require Python to be installed:

```
pip install pyinstaller
pyinstaller --onefile --windowed main.py
```

The executable will be created in the `dist/` folder. Note that the `config/` folder and the `logs/` folder will be created at runtime in the same directory as the executable.

---

## Troubleshooting

| Symptom | What to check |
|---|---|
| App does not start / crashes immediately | Check `logs/worklog.log` for a Python traceback. |
| "Database is locked" error | The `.db` file may be open in another process, or the network share is slow. Try closing other instances of the app. WAL mode should reduce this. |
| Calendar popup shows a plain text field instead of a calendar | `tkcalendar` is not installed. Run `pip install tkcalendar`. |
| HTML report is not updated after Save | Check that the HTML path in Settings points to a location you have write access to. Look for errors in `logs/worklog.log`. |
| Backup files filling up disk | Each save creates at most one backup file per calendar day (named `worklog_backup_YYYYMMDD.db`). You can disable backups in Settings or manually delete old backup files. |

All errors and key events are written to **`logs/worklog.log`** in the same directory as `main.py`. This is the first place to look when something goes wrong.

---

## Database Schema Summary

### `app_metadata`

| Column | Type | Notes |
|---|---|---|
| `key` | TEXT (PK) | e.g. `app_name`, `schema_version`, `created_at` |
| `value` | TEXT | Corresponding value |

### `daily_entries`

| Column | Type | Notes |
|---|---|---|
| `id` | INTEGER (PK) | Auto-incremented |
| `entry_date` | TEXT UNIQUE | ISO-8601: `YYYY-MM-DD` |
| `notes` | TEXT | Optional daily note |
| `created_at` | TEXT | ISO-8601 UTC datetime |
| `updated_at` | TEXT | Bumped on every save |

### `work_items`

| Column | Type | Notes |
|---|---|---|
| `id` | INTEGER (PK) | Auto-incremented |
| `entry_id` | INTEGER (FK) | References `daily_entries.id` -- CASCADE delete |
| `sort_order` | INTEGER | Preserves UI row order |
| `task` | TEXT | Task performed |
| `reason` | TEXT | Reason / context |
| `hours` | REAL | Supports fractional values (e.g. 0.5, 1.25) |
| `created_at` | TEXT | ISO-8601 UTC datetime |
| `updated_at` | TEXT | Bumped on every save |
