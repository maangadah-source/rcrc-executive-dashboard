"""
RCRC Executive Dashboard Auto-Builder
--------------------------------------
Reads Stakeholder Engagement Plan (4).xlsx and injects fresh data into
window.DASH_DATA inside index.html, WITHOUT touching any HTML, CSS, or
JavaScript function outside that single data object.

This script is designed to run inside GitHub Actions on every push,
so any Excel update automatically produces an updated index.html that
Netlify then deploys.
"""

import json
import re
import sys
from datetime import datetime, date
from pathlib import Path

from openpyxl import load_workbook

EXCEL_FILE = "Stakeholder Engagement Plan (4).xlsx"
HTML_FILE = "index.html"


def fmt(v):
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def table_from_named_table(wb, sheet_name, table_name):
    """Read an Excel Table object by name and return list[dict] of its rows."""
    ws = wb[sheet_name]
    tbl = ws.tables.get(table_name)
    if tbl is None:
        raise ValueError(f"Table '{table_name}' not found in sheet '{sheet_name}'")
    ref = tbl.ref  # e.g. "A7:Q18"
    cells = ws[ref]
    rows = [[c.value for c in row] for row in cells]
    headers = [fmt(h) for h in rows[0]]
    # de-duplicate headers (Excel appends 2, 3... to repeats)
    seen = {}
    clean_headers = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            clean_headers.append(f"{h}{seen[h]}")
        else:
            seen[h] = 1
            clean_headers.append(h)

    records = []
    for row in rows[1:]:
        if all(v is None or str(v).strip() == "" for v in row):
            continue
        rec = {}
        for h, v in zip(clean_headers, row):
            if not h:
                continue
            rec[h] = fmt(v)
        records.append(rec)
    return records


def find_table_sheet(wb, table_name):
    for ws in wb.worksheets:
        if table_name in ws.tables:
            return ws.title
    raise ValueError(f"Table '{table_name}' not found in any sheet")


def main():
    wb = load_workbook(EXCEL_FILE, data_only=True)

    tables = {
        "schedulePriority": "tblSchedulingPriority",
        "risks": "tblRiskRegister",
        "previous": "tblPreviousTracker",
        "followUpStatus": "tblFollowUpStatus",
        "questions": "tblQuestionBank",
        "stakeholders": "tblEngagementStrategy",
    }

    data = {}
    for key, table_name in tables.items():
        sheet = find_table_sheet(wb, table_name)
        data[key] = table_from_named_table(wb, sheet, table_name)
        print(f"[ok] {key}: {len(data[key])} rows from '{table_name}' ({sheet})")

    # ---- validation guardrails: never publish empty/broken data ----
    if len(data["schedulePriority"]) == 0:
        print("[error] schedulePriority table is empty, aborting to protect dashboard")
        sys.exit(1)
    if len(data["risks"]) == 0:
        print("[error] risks table is empty, aborting to protect dashboard")
        sys.exit(1)

    html_path = Path(HTML_FILE)
    html = html_path.read_text(encoding="utf-8")

    pattern = re.compile(r"window\.DASH_DATA\s*=\s*(\{.*?\});", re.S)
    match = pattern.search(html)
    if not match:
        print("[error] Could not locate window.DASH_DATA block in index.html")
        sys.exit(1)

    try:
        existing = json.loads(match.group(1))
    except Exception as e:
        print(f"[error] Existing DASH_DATA is not valid JSON, aborting: {e}")
        sys.exit(1)

    # Only overwrite the six keys we manage; leave any other existing key untouched.
    existing.update(data)

    new_json = json.dumps(existing, ensure_ascii=False, separators=(",", ":"))
    new_block = "window.DASH_DATA = " + new_json + ";"

    new_html = html[: match.start()] + new_block + html[match.end():]

    # Basic safety check: new file should not be drastically smaller than before
    if len(new_html) < len(html) * 0.5:
        print("[error] Resulting HTML looks truncated, aborting to protect dashboard")
        sys.exit(1)

    html_path.write_text(new_html, encoding="utf-8")
    print(f"[done] index.html updated successfully ({len(new_html)} bytes)")


if __name__ == "__main__":
    main()
