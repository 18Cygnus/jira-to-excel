import os
import math
import time
from typing import List, Dict, Any, Optional

import requests
import pandas as pd
from dotenv import load_dotenv
from datetime import datetime
from pathlib import Path

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

load_dotenv()

JIRA_BASE_URL = os.getenv("JIRA_BASE_URL", "").rstrip("/")
JIRA_EMAIL = os.getenv("JIRA_EMAIL")
JIRA_API_TOKEN = os.getenv("JIRA_API_TOKEN")
JIRA_FILTER_ID = os.getenv("JIRA_FILTER_ID")
JIRA_FILTER_NAME = os.getenv("JIRA_FILTER_NAME")

PROJECT_ROOT = Path(__file__).resolve().parent
EXCEL_DIR = PROJECT_ROOT / "excels"
EXCEL_DIR.mkdir(parents=True, exist_ok=True)  

# Tuning

MAX_RESULTS = 100
REQUEST_TIMEOUT = 30
RETRY_429_SECONDS = 10
OUTPUT_FILE = EXCEL_DIR / "jira_export.xlsx"

SESSION = requests.Session()
SESSION.auth = (JIRA_EMAIL, JIRA_API_TOKEN)
SESSION.headers.update({"Accept": "application/json"})

def _require_env():
    missing = [k for k, v in {
        "JIRA_BASE_URL": JIRA_BASE_URL,
        "JIRA_EMAIL": JIRA_EMAIL,
        "JIRA_API_TOKEN": JIRA_API_TOKEN
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Missing env vars: {', '.join(missing)}")
    
def find_filter_id_by_name(name: str) -> Optional[int]:
    """
    Uses /rest/api/3/filter/search to find a filter ID by name (exact match preferred).
    """
    url = f"{JIRA_BASE_URL}/rest/api/3/filter/search"
    params = {"filterName": name}
    r = SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
    if r.status_code != 200:
        raise RuntimeError(f"Failed to find filter ID for '{name}': {r.status_code} {r.text}")
    data = r.json()
    candidates = data.get("values", [])
    # prefer exact match, else first
    for f in candidates:
        if f.get("name") == name:
            return int(f.get("id"))
    if candidates:
        return int(candidates[0].get("id"))
    return None

def fetch_all_issues_for_filter(filter_id: int, fields: List[str]) -> List[Dict[str, Any]]:
    """
    Paginates /rest/api/3/search with jql=filter<id>
    """
    start_at = 0
    issues: List[Dict[str, Any]] = []

    while True:
        params = {
            "jql": f"filter={filter_id}",
            "startAt": start_at,
            "maxResults": MAX_RESULTS,
            "fields": ",".join(fields)
        }
        url = f"{JIRA_BASE_URL}/rest/api/3/search"

        r = SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
        if r.status_code == 429:
            time.sleep(RETRY_429_SECONDS)
            continue
        if r.status_code != 200:
            raise RuntimeError(f"Failed to fetch issues for filter {filter_id}: {r.status_code} {r.text}")
        
        data = r.json()
        issues_page = data.get("issues", [])
        issues.extend(issues_page)

        total = data.get("total", 0)
        start_at += len(issues_page)
        if start_at >= total or not issues_page:
            break

    return issues

def dt_obj(dt_str):
    """
    Converts a datetime string from Jira to a datetime object.
    """
    if not dt_str:
        return None
    # contoh input: 2025-05-27T11:56:17.563+0700
    dt = datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%S.%f%z")
    # Buang timezone -> Excel-friendly (tetap jam lokal sesuai string)
    return dt.replace(tzinfo=None)

def d_obj(d_str):
    """
    Converts a date string from Jira to a date object.
    """
    return datetime.strptime(d_str, "%Y-%m-%d").date() if d_str else None

def flatten_issue(issue: Dict[str, Any]) -> Dict[str, Any]:
    """
    Converts Jira's nested issue JSON to flat columns
    add or remove fields as necessary
    """
    fields = issue.get("fields", {})
    def g(*keys, default=None):
        cur = fields
        for k in keys:
            cur = (cur or {}).get(k)
        return cur if cur is not None else default
    
    res_obj = fields.get("resolution")
    resolution_name = (
        res_obj.get("name") if isinstance(res_obj, dict) and res_obj.get("name") else "Unresolved"
    )
    
    return {
        "key": issue.get("key"),
        "work": g("summary"),
        "assignee": g("assignee", "displayName"),
        "reporter": g("reporter", "displayName"),
        "priority": g("priority", "name"),
        "status": g("status", "name"),
        "resolution": resolution_name,
        "created": dt_obj(g("created")),
        "updated": dt_obj(g("updated")),
        "due_date": d_obj(g("duedate"))
    }

def format_duration(seconds: float) -> str:
    total = int(round(seconds))
    m, s = divmod(total, 60)
    if m >= 60:
        h, m = divmod(m, 60)
        return f"{h} h {m} min {s} sec"
    return f"{m} min {s} sec"

def main():
    _require_env()

    # filter ID
    filter_id = None
    if JIRA_FILTER_ID:
        try:
            filter_id = int(JIRA_FILTER_ID)
        except ValueError:
            raise RuntimeError("JIRA_FILTER_ID must be a number")
    elif JIRA_FILTER_NAME:
        filter_id = find_filter_id_by_name(JIRA_FILTER_NAME)
        if filter_id is None:
            raise RuntimeError(f"Filter '{JIRA_FILTER_NAME}' not found")
    else:
        raise RuntimeError("Provide JIRA_FILTER_ID or JIRA_FILTER_NAME in .env")
    
    # choose the fields you want from Jira
    # include everything you use in flatten_issue()
    fields = [
        "summary", "assignee", "reporter", "priority", "status", "resolution",
        "created", "updated", "duedate", # change/remove if your instance differs
    ]

    print(f"Fetching issues for filter ID {filter_id} ...")
    t0 = time.perf_counter() 
    issues = fetch_all_issues_for_filter(filter_id, fields=fields)
    print(f"Fetched {len(issues)} issues")
    t_end = time.perf_counter()
    print(f"Total run time   : {format_duration(t_end - t0)}")

    rows = [flatten_issue(i) for i in issues]
    df = pd.DataFrame(rows)

    # sort & write to excel (sort by "created" column)
    if not df.empty:
        df.sort_values(by=["created"], ascending=False, inplace=True)

    with pd.ExcelWriter(str(OUTPUT_FILE), engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Issues")
        
        # Auto-fit columns
        ws = writer.sheets["Issues"]
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)) if cell.value else 0)
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

        # Add filtering dropdown for each column
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.table import Table, TableStyleInfo

        # Tentukan range data (dari A1 sampai kolom terakhir & baris terakhir)
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        table = Table(displayName="IssuesTable", ref=ref)

        # # Configure column style
        # style = TableStyleInfo(
        #     name="TableStyleMedium9",
        #     showRowStripes=True,
        #     showColumnStripes=False
        # )
        # table.tableStyleInfo = style

        # Tambahkan table ke worksheet
        ws.add_table(table)

        # Freeze header row (opsional, supaya header tetap terlihat saat scroll)
        ws.freeze_panes = "A2"

        # === Format tanggal ===
        from openpyxl.styles import numbers

        # Cari index kolom (1-based untuk Excel)
        cols = {name: i+1 for i, name in enumerate(df.columns)}
        # created/updated: "Aug 15, 2025, 6:42 PM"
        fmt_dt = 'mmm d, yyyy, h:mm AM/PM'
        for row in range(2, ws.max_row + 1):
            if "created" in cols:
                ws.cell(row=row, column=cols["created"]).number_format = fmt_dt
            if "updated" in cols:
                ws.cell(row=row, column=cols["updated"]).number_format = fmt_dt
        # due_date: "Aug 15, 2025"
        if "due_date" in cols:
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=cols["due_date"]).number_format = 'mmm d, yyyy'

    print(f"Exported {len(issues)} issues to {OUTPUT_FILE}")

if __name__ == "__main__":
    main()