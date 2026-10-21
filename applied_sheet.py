#!/usr/bin/env python3
"""
Maintain the 'Applied' tab in the same Google Sheet as sheets_sync.py.

One row per application: when it went in, which resume and cover letter were
used, who was contacted about it, and what happens next. Seed rows live in
applied_seed.json; any row whose (Company, Role) is not in the tab yet is
appended. Existing rows are never edited or deleted, so statuses and notes
typed by hand in the Sheet always win.

Run by hand after adding an application to applied_seed.json:
    /usr/local/bin/python3 applied_sheet.py

Same quiet-no-op contract as the other sheet scripts: missing credentials or
config just skip, and any error exits 0.
"""
import json
import os
import socket
import sys

# See rotational_sheet.py: never let a Google API call block forever.
socket.setdefaulttimeout(60)

HERE = os.path.dirname(os.path.abspath(__file__))
KEY_PATH = os.path.join(HERE, "gcp-service-account.json")
CONFIG_PATH = os.path.join(HERE, "sheet_config.json")
SEED_PATH = os.path.join(HERE, "applied_seed.json")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TAB = "Applied"
HEADER = ["Company", "Role", "Location", "Date Applied", "Status",
          "Resume Used", "Cover Letter", "Contacts", "Contact Status",
          "Next Step", "Posting URL", "Notes"]


def log(msg):
    print(f"[applied_sheet] {msg}")


def main():
    if not (os.path.exists(KEY_PATH) and os.path.exists(CONFIG_PATH)):
        log("no credentials/config; skipping")
        return
    import gspread
    from google.oauth2.service_account import Credentials

    with open(CONFIG_PATH) as f:
        sheet_id = json.load(f)["sheet_id"]
    with open(SEED_PATH) as f:
        seed = json.load(f)

    creds = Credentials.from_service_account_file(KEY_PATH, scopes=SCOPES)
    book = gspread.authorize(creds).open_by_key(sheet_id)
    try:
        ws = book.worksheet(TAB)
    except gspread.WorksheetNotFound:
        ws = book.add_worksheet(title=TAB, rows=200, cols=len(HEADER))
        ws.append_row(HEADER, value_input_option="RAW")
        ws.format("A1:L1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        log("created tab")

    rows = ws.get_all_values()
    have = {(r[0].strip().lower(), r[1].strip().lower()) for r in rows[1:] if len(r) > 1}
    new = [[e.get(h, "") for h in HEADER] for e in seed
           if (e["Company"].strip().lower(), e["Role"].strip().lower()) not in have]
    if new:
        ws.append_rows(new, value_input_option="USER_ENTERED")
    log(f"{len(new)} new row(s), {len(rows) - 1} already present")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # never break the pipeline
        log(f"error: {e}")
        sys.exit(0)
