"""
html_generator.py
-----------------
Generates a self-contained HTML report from WorkLog database entries.

The output is a single .html file with all CSS and JavaScript embedded —
no internet connection required, no external fonts or CDNs.  It can be
opened in any browser and emailed or linked from a shared folder.

Features
--------
  - Entries sorted newest first
  - Collapsible date sections (click the header to expand/collapse)
  - In-browser live search (pure JavaScript)
  - Hours total per day shown in the section header
  - Grand total hours in the page header
  - Branding header (name, team, org — customisable in Settings)
  - Last-generated timestamp in the footer
  - Print-friendly via @media print CSS
  - Atomic write (temp file + rename) to avoid partial-file corruption
"""

import os
from datetime import datetime
from typing import List

from logger_setup import get_logger

log = get_logger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Embedded CSS
# ─────────────────────────────────────────────────────────────────────────────

_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto,
                 'Helvetica Neue', Arial, sans-serif;
    font-size: 14px;
    background: #f0f2f5;
    color: #1a1a2e;
    line-height: 1.55;
}

/* ── Header bar ──────────────────────────────────────────────────────── */
.site-header {
    background: linear-gradient(135deg, #1a2744 0%, #2a4080 100%);
    color: white;
    padding: 28px 40px 22px;
    border-bottom: 4px solid #c8a951;
}
.site-header h1 {
    font-size: 24px;
    font-weight: 700;
    letter-spacing: -0.2px;
    margin-bottom: 4px;
}
.site-header .byline {
    font-size: 13px;
    opacity: 0.75;
    margin-bottom: 16px;
}
.header-stats {
    display: flex;
    gap: 28px;
    flex-wrap: wrap;
    font-size: 13px;
}
.header-stats span { opacity: 0.85; }
.header-stats strong { opacity: 1; font-size: 15px; }

/* ── Sticky controls bar ─────────────────────────────────────────────── */
.controls {
    background: white;
    padding: 12px 40px;
    border-bottom: 1px solid #dde1ea;
    display: flex;
    align-items: center;
    gap: 14px;
    position: sticky;
    top: 0;
    z-index: 100;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
.search-icon { font-size: 16px; color: #888; }
.controls input[type=text] {
    flex: 1;
    max-width: 420px;
    padding: 8px 14px;
    border: 1px solid #ccd0da;
    border-radius: 6px;
    font-size: 14px;
    outline: none;
    transition: border-color 0.18s;
}
.controls input[type=text]:focus { border-color: #2a4080; }
.grand-total-pill {
    margin-left: auto;
    background: #eef2ff;
    color: #2a4080;
    border: 1px solid #c5cfee;
    border-radius: 20px;
    padding: 4px 16px;
    font-size: 13px;
    font-weight: 600;
    white-space: nowrap;
}

/* ── Page content wrapper ────────────────────────────────────────────── */
.content {
    max-width: 1020px;
    margin: 0 auto;
    padding: 28px 20px 64px;
}

/* ── Day section card ────────────────────────────────────────────────── */
.day-section {
    background: white;
    border: 1px solid #dde1ea;
    border-radius: 10px;
    margin-bottom: 16px;
    overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    transition: box-shadow 0.15s;
}
.day-section:hover { box-shadow: 0 3px 10px rgba(0,0,0,0.09); }

/* Clickable section header */
.day-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 13px 20px;
    cursor: pointer;
    user-select: none;
    background: #f8f9fc;
    border-bottom: 1px solid #dde1ea;
    transition: background 0.12s;
}
.day-header:hover { background: #eef1f9; }
.day-header h2 {
    font-size: 15px;
    font-weight: 600;
    color: #1a2744;
}
.day-meta { display: flex; align-items: center; gap: 14px; }
.hours-badge {
    background: #1a2744;
    color: white;
    border-radius: 20px;
    padding: 3px 13px;
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.2px;
}
.chevron {
    color: #999;
    font-size: 11px;
    transition: transform 0.18s;
}
/* Collapsed state */
.day-section.collapsed .chevron    { transform: rotate(-90deg); }
.day-section.collapsed .day-body  { display: none; }

/* ── Work items table ────────────────────────────────────────────────── */
.day-body { padding: 0; }

table { width: 100%; border-collapse: collapse; }

thead th {
    padding: 9px 16px;
    text-align: left;
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    color: #777;
    background: #fafbfd;
    border-bottom: 1px solid #eaecf2;
}
.col-task   { width: 38%; }
.col-reason { width: 44%; }
.col-hours  { width: 18%; text-align: right; }

tbody tr {
    border-bottom: 1px solid #f0f2f6;
    transition: background 0.1s;
}
tbody tr:last-child { border-bottom: none; }
tbody tr:hover { background: #f7f9ff; }

tbody td {
    padding: 10px 16px;
    vertical-align: top;
    font-size: 13.5px;
    color: #222;
}
.hours-val {
    color: #2a4080;
    font-weight: 600;
    font-variant-numeric: tabular-nums;
    text-align: right;
    display: block;
}

/* Optional daily note row */
.notes-row td {
    background: #fffdf0;
    font-style: italic;
    color: #777;
    font-size: 12.5px;
    padding: 7px 16px;
    border-top: 1px dashed #e8dec8;
}

/* No items placeholder */
.no-items td {
    color: #bbb;
    font-style: italic;
    padding: 12px 16px;
}

/* Empty database message */
.no-entries {
    text-align: center;
    padding: 72px 20px;
    color: #aaa;
    font-size: 16px;
}

/* ── Footer ──────────────────────────────────────────────────────────── */
.site-footer {
    text-align: center;
    padding: 18px 20px;
    font-size: 11.5px;
    color: #bbb;
    border-top: 1px solid #e0e3ea;
    background: white;
}

/* ── Search filter helper ────────────────────────────────────────────── */
.day-section.hidden { display: none; }

/* ── Print styles ────────────────────────────────────────────────────── */
@media print {
    .controls                          { display: none; }
    .site-header                       { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .day-section                       { break-inside: avoid; box-shadow: none; border: 1px solid #ccc; margin-bottom: 10px; }
    .day-section.collapsed .day-body   { display: block !important; }
    .chevron                           { display: none; }
    body                               { background: white; }
}
"""

# ─────────────────────────────────────────────────────────────────────────────
# Embedded JavaScript
# ─────────────────────────────────────────────────────────────────────────────

_JS = """
// ── Collapse / expand day sections ───────────────────────────────────────────
document.querySelectorAll('.day-header').forEach(function(header) {
    header.addEventListener('click', function() {
        this.closest('.day-section').classList.toggle('collapsed');
    });
});

// ── Live search ───────────────────────────────────────────────────────────────
var searchInput = document.getElementById('search-input');
if (searchInput) {
    searchInput.addEventListener('input', function() {
        var query = this.value.trim().toLowerCase();
        document.querySelectorAll('.day-section').forEach(function(section) {
            if (!query) {
                section.classList.remove('hidden', 'collapsed');
            } else if (section.innerText.toLowerCase().includes(query)) {
                section.classList.remove('hidden', 'collapsed');
            } else {
                section.classList.add('hidden');
            }
        });
    });
}
"""


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def generate_html(
    entries: List[dict],
    output_path: str,
    author_name: str = "",
    author_team: str = "",
    author_org: str = "",
) -> bool:
    """
    Generate the HTML report and write it to output_path.

    Uses an atomic write (write to .tmp file, then os.replace) so that
    the existing report is never left in a partial/corrupt state.

    Parameters
    ----------
    entries     : list of dicts from Database.get_all_entries_with_items()
    output_path : full path to the target .html file
    author_*    : branding strings shown in the report header

    Returns True on success, False on failure (errors are logged).
    """
    log.info("Generating HTML report → %s  (%d entries)", output_path, len(entries))
    try:
        html_content = _build_html(entries, author_name, author_team, author_org)

        # Ensure the output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        # If the file already exists, remove any read-only attribute before writing.
        # This handles cases where antivirus, OneDrive, or Windows sets the file
        # read-only after the first write.
        if os.path.exists(output_path):
            try:
                import stat
                os.chmod(output_path, stat.S_IWRITE | stat.S_IREAD)
                log.debug("Cleared read-only attribute on %s", output_path)
            except OSError as exc:
                log.warning("Could not clear read-only attribute: %s", exc)

        # Also clear any leftover .tmp file from a previous failed attempt
        tmp_path = output_path + ".tmp"
        if os.path.exists(tmp_path):
            try:
                import stat
                os.chmod(tmp_path, stat.S_IWRITE | stat.S_IREAD)
                os.remove(tmp_path)
            except OSError:
                pass

        # Try atomic write (temp + rename), fall back to direct write
        try:
            with open(tmp_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            os.replace(tmp_path, output_path)
            log.info("HTML report written (atomic).")
        except OSError:
            log.warning("Atomic write failed; writing directly to %s", output_path)
            with open(output_path, "w", encoding="utf-8") as f:
                f.write(html_content)
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
            except OSError:
                pass
            log.info("HTML report written (direct).")

        return True

    except OSError as exc:
        log.error("Failed to write HTML report: %s", exc)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# Internal builders
# ─────────────────────────────────────────────────────────────────────────────

def _escape(text: str) -> str:
    """Minimal HTML escaping for user-supplied content."""
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _format_date(iso_date: str) -> str:
    """
    Convert 'YYYY-MM-DD' → 'Monday, April 13, 2026'.
    Cross-platform: uses d.day (int) instead of %-d (Linux-only).
    """
    try:
        d = datetime.strptime(iso_date, "%Y-%m-%d")
        return f"{d.strftime('%A, %B')} {d.day}, {d.year}"
    except ValueError:
        return iso_date


def _now_human() -> str:
    """Return a human-readable 'last generated' string, cross-platform."""
    n = datetime.now()
    hour = n.hour % 12 or 12
    ampm = "AM" if n.hour < 12 else "PM"
    return f"{n.strftime('%B')} {n.day}, {n.year} at {hour}:{n.strftime('%M')} {ampm}"


def _build_html(
    entries: List[dict],
    author_name: str,
    author_team: str,
    author_org: str,
) -> str:
    """Assemble the full HTML document string."""
    total_entries = len(entries)
    grand_total   = sum(sum(i["hours"] for i in e["items"]) for e in entries)

    # Build the "Name | Team | Org" subtitle line (skip empty parts)
    parts    = [p for p in [author_name, author_team, author_org] if p]
    byline   = " &nbsp;|&nbsp; ".join(_escape(p) for p in parts) or "WorkLog Report"

    if entries:
        sections_html = "\n".join(_build_day_section(e) for e in entries)
    else:
        sections_html = (
            '<div class="no-entries">'
            "No entries yet. Open WorkLog and start logging your work!"
            "</div>"
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>WorkLog Report</title>
  <style>{_CSS}</style>
</head>
<body>

<header class="site-header">
  <h1>WorkLog Report</h1>
  <p class="byline">{byline}</p>
  <div class="header-stats">
    <span><strong>{total_entries}</strong> entries</span>
    <span><strong>{grand_total:.1f}</strong> total hours logged</span>
  </div>
</header>

<div class="controls">
  <span class="search-icon">&#128269;</span>
  <input type="text" id="search-input"
         placeholder="Search tasks, reasons, or dates&#8230;" autocomplete="off">
  <span class="grand-total-pill">{total_entries} entries &nbsp;&middot;&nbsp; {grand_total:.1f} hrs</span>
</div>

<main class="content">
{sections_html}
</main>

<footer class="site-footer">
  Generated by <strong>WorkLog</strong> &nbsp;&middot;&nbsp;
  Last updated: {_escape(_now_human())}
</footer>

<script>{_JS}</script>
</body>
</html>"""


def _build_day_section(entry: dict) -> str:
    """Build the HTML block for one day (header + table)."""
    day_total  = sum(item["hours"] for item in entry["items"])
    date_label = _escape(_format_date(entry["date"]))
    rows_html  = _build_table_rows(entry["items"])

    notes_html = ""
    if entry.get("notes", "").strip():
        notes_html = (
            f'<tr class="notes-row">'
            f'<td colspan="3"><em>Note:</em> {_escape(entry["notes"])}</td>'
            f"</tr>\n"
        )

    return f"""<section class="day-section" data-date="{_escape(entry['date'])}">
  <div class="day-header">
    <h2>{date_label}</h2>
    <div class="day-meta">
      <span class="hours-badge">{day_total:.1f} hrs</span>
      <span class="chevron">&#9660;</span>
    </div>
  </div>
  <div class="day-body">
    <table>
      <thead>
        <tr>
          <th class="col-task">Task Performed</th>
          <th class="col-reason">Reason / Context</th>
          <th class="col-hours">Hours</th>
        </tr>
      </thead>
      <tbody>
{rows_html}{notes_html}      </tbody>
    </table>
  </div>
</section>"""


def _build_table_rows(items: List[dict]) -> str:
    """Build <tr> elements for each work item."""
    if not items:
        return (
            '        <tr class="no-items">'
            '<td colspan="3">No work items recorded for this day.</td>'
            "</tr>\n"
        )
    rows = []
    for item in items:
        h = item["hours"]
        # Show "2" for whole numbers, "1.5" for fractional
        h_str = str(int(h)) if h == int(h) else f"{h:.1f}"
        rows.append(
            "        <tr>\n"
            f'          <td class="col-task">{_escape(item["task"])}</td>\n'
            f'          <td class="col-reason">{_escape(item["reason"])}</td>\n'
            f'          <td class="col-hours"><span class="hours-val">{h_str}</span></td>\n'
            "        </tr>"
        )
    return "\n".join(rows) + "\n"
