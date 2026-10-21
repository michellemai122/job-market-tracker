#!/usr/bin/env python3
"""
Maintain the 'Rotational Programs' tab in the same Google Sheet as sheets_sync.py.

Two jobs each cycle:
  1. Seed: any program in rotational_programs.json that isn't in the tab yet gets
     a row (Program, Company, Typical Window, program link). Rows are never
     deleted, so programs added by hand in the Sheet survive.
  2. Cross-reference: scan data/jobs.xlsx for live APM/rotational postings at the
     watchlist companies. On a match, fill in Latest Posting Seen / Posting Title /
     Posting URL and flip Status to "OPEN" — but only if Status is blank or
     "Watch", so hand-entered statuses like "Applied" are never clobbered.
     The Notes column is never touched.

Same quiet-no-op contract as sheets_sync.py: missing credentials or config just
skip, and any error exits 0 so the poller pipeline never breaks.
"""
import json
import os
import re
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

import socket

# Never let a Google API call block forever. gspread -> requests -> urllib3 sets
# no timeout by default, so a half-open socket (laptop sleep, WiFi change, VPN)
# hangs the read indefinitely. That stalls run.sh, and launchd will not start a
# new cycle while the old one is alive -- one hang silently freezes the tracker.
socket.setdefaulttimeout(60)

HERE = os.path.dirname(os.path.abspath(__file__))
KEY_PATH = os.path.join(HERE, "gcp-service-account.json")
CONFIG_PATH = os.path.join(HERE, "sheet_config.json")
PROGRAMS_PATH = os.path.join(HERE, "rotational_programs.json")
XLSX_PATH = os.path.join(HERE, "data", "jobs.xlsx")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

TAB = "Rotational Programs"
HEADER = ["Program", "Company", "Typical Window", "Status",
          "Latest Posting Seen (ET)", "Posting Title", "Posting URL",
          "Program Link", "Notes"]
COL = {name: i + 1 for i, name in enumerate(HEADER)}  # 1-indexed

# A posting counts for the watchlist if the poller filed it as APM, or its title
# carries a rotational/APM signal.
ROTATIONAL = re.compile(r"\b(rotational|rotation\s*program|associate\s+product\s+manager"
                        r"|\bapm\b|product\s+development\s+program)\b", re.I)

# A real new-grad PM *program* (vs. a one-off "Associate PM - Pricing" job): the
# title says so. Used to count New-Grad PM rows (e.g. TikTok "Product Manager
# Graduate") and to auto-add programs at companies not on the watchlist yet.
PROGRAM = re.compile(r"\b(new\s*grad|new\s+college\s+grad|university|early\s+career|"
                     r"graduate|program|launchpad|accelerator|rotational|campus|"
                     r"20\d\d\s+start|\[20\d\d\])", re.I)
NOT_PROGRAM = re.compile(r"\b(mba|intern|internship)\b", re.I)


def is_program(category, title):
    t = str(title)
    if NOT_PROGRAM.search(t):
        return False
    return category in ("APM", "New-Grad PM") and bool(PROGRAM.search(t))


def names(p):
    """All lowercase company names a watchlist entry answers to."""
    return {n.strip().lower() for n in [p["company"], *p.get("aliases", [])]}


def log(msg):
    print(f"[rotational_sheet] {msg}")


def latest_postings_by_company():
    """company(lower) -> (first_seen, title, url, is_program, company) for the newest
    matching posting, plus the set of every URL still in jobs.xlsx."""
    if not os.path.exists(XLSX_PATH):
        return {}, set()
    from openpyxl import load_workbook
    wb = load_workbook(XLSX_PATH, read_only=True)
    ws = wb["Jobs"]
    out, urls = {}, set()
    # rows are newest-first, so the first hit per company is the latest
    for row in ws.iter_rows(min_row=2, values_only=True):
        first_seen, company, title, category = (row + ("",) * 4)[:4]
        url = (row[8] if len(row) > 8 else "") or ""
        urls.add(str(url).strip())
        if not company or not title:
            continue
        prog = is_program(category, title)
        if category != "APM" and not ROTATIONAL.search(str(title)) and not prog:
            continue
        key = str(company).strip().lower()
        if key not in out or (prog and not out[key][3]):  # prefer a real program
            out[key] = (str(first_seen or ""), str(title), str(url), prog, str(company).strip())
    return out, urls


def add_new_programs(programs, latest):
    """Append a watchlist entry for every company with a live new-grad PM program
    that isn't tracked yet. Saved to rotational_programs.json so it's permanent."""
    known = set().union(*(names(p) for p in programs))
    added = []
    for key, (first_seen, title, url, prog, company) in latest.items():
        if not prog or key in known:
            continue
        added.append({"program": f"{company} {'APM' if 'associate' in title.lower() else 'New-Grad PM'}",
                      "company": company, "window": f"First seen {first_seen[:10]}",
                      "link": url, "auto": True})
    if added:
        data = json.load(open(PROGRAMS_PATH))
        data["programs"].extend(added)
        with open(PROGRAMS_PATH, "w") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            f.write("\n")
        programs.extend(added)
    return added


def notify(title, body):
    if sys.platform != "darwin":
        return
    import subprocess
    body = body.replace('"', "'")
    subprocess.run(["osascript", "-e", f'display notification "{body}" with title "{title}" '
                    f'sound name "Glass"'], timeout=10, check=False)


def main():
    if not os.path.exists(KEY_PATH) or not os.path.exists(CONFIG_PATH):
        log("no credentials/config yet — skipping.")
        return
    if not os.path.exists(PROGRAMS_PATH):
        log("no rotational_programs.json — skipping.")
        return

    import gspread
    from google.oauth2.service_account import Credentials

    programs = json.load(open(PROGRAMS_PATH))["programs"]
    latest, live_urls = latest_postings_by_company()
    new_programs = add_new_programs(programs, latest)
    sheet_id = json.load(open(CONFIG_PATH))["sheet_id"]
    creds = Credentials.from_service_account_file(KEY_PATH, scopes=SCOPES)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(sheet_id)

    try:
        ws = sh.worksheet(TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(TAB, rows=100, cols=len(HEADER))
        ws.update([HEADER], "A1")
        ws.format("A1:I1", {"textFormat": {"bold": True}})
        ws.freeze(rows=1)
        log(f"created tab '{TAB}'.")

    rows = ws.get_all_values()
    existing = {r[0].strip().lower(): i + 1 for i, r in enumerate(rows) if i > 0 and r and r[0].strip()}

    # 1. seed missing programs
    to_add = [[p["program"], p["company"], p.get("window", ""), "Watch",
               "", "", "", p.get("link", ""), ""]
              for p in programs if p["program"].strip().lower() not in existing]
    if to_add:
        ws.append_rows(to_add, value_input_option="USER_ENTERED")
        log(f"seeded {len(to_add)} programs.")
        rows = ws.get_all_values()
        existing = {r[0].strip().lower(): i + 1 for i, r in enumerate(rows) if i > 0 and r and r[0].strip()}

    # 2. cross-reference live postings. Status is only ever moved between
    # Watch / OPEN / Closed -- anything you typed (e.g. "Applied") is left alone.
    updates, opened, closed = [], [], []
    for p in programs:
        rownum = existing.get(p["program"].strip().lower())
        if not rownum:
            continue
        hits = [latest[n] for n in names(p) if n in latest]
        hit = max(hits, key=lambda h: (h[3], h[0])) if hits else None
        row = rows[rownum - 1] + [""] * len(HEADER)
        status = row[COL["Status"] - 1].strip().lower()
        cell = lambda name: gspread.utils.rowcol_to_a1(rownum, COL[name])
        if not hit:
            old_url = row[COL["Posting URL"] - 1].strip()
            if status == "open" and old_url and old_url not in live_urls:
                updates.append({"range": cell("Status"), "values": [["Closed"]]})
                closed.append(p["program"])
            continue
        first_seen, title, url = hit[:3]
        if row[COL["Posting URL"] - 1].strip() == url.strip() and url and status != "closed":
            continue  # already recorded
        updates.append({"range": cell("Latest Posting Seen (ET)") + ":" + cell("Posting URL"),
                        "values": [[first_seen, title, url]]})
        if status in ("", "watch", "closed"):
            updates.append({"range": cell("Status"), "values": [["OPEN"]]})
            opened.append(f"{p['program']}: {title}")
    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        log(f"updated {len(updates)} cells from live postings.")
    else:
        log("no new rotational postings matched.")
    for o in opened:
        log(f"OPEN: {o}")
    for c in closed:
        log(f"Closed (posting gone): {c}")
    if opened:
        notify(f"Rotational — {len(opened)} program(s) OPEN",
               opened[0][:90] + (f" (+{len(opened) - 1} more)" if len(opened) > 1 else ""))

    if ws.col_count < 12:
        ws.add_cols(12 - ws.col_count)
    stamp = datetime.now(ZoneInfo("America/New_York")).strftime("%Y-%m-%d %H:%M ET")
    ws.update([["Last checked:", stamp]], "K1", value_input_option="USER_ENTERED")
    ws.format("K1", {"textFormat": {"bold": True}})


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log(f"ERROR: {e}")
        sys.exit(0)  # never break the poller pipeline
