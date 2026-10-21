# Job Market Tracker

A personal job-hunt pipeline. Every 10 minutes it polls the public job boards of 57 companies (about 13,000 live postings), keeps only the new-grad, APM, AI-engineer and forward-deployed roles in the US, and adds anything new to a spreadsheet and a Google Sheet, with a desktop alert.

```
registry.json ──► poller.py ──► data/jobs.xlsx ──► sheets_sync.py ──► Google Sheet "Jobs"
 (57 companies)     │  dedupe on data/seen.json                        ▲
                    │                                                  │
custom_sources.py ──┘                        prune_closed.py ──────────┘
 (Amazon)                                     (removes closed postings, every 6h)
```

## How it works

**Finding each company's feed.** `discover.py` guesses slug variants for each company and probes the public JSON endpoints of Greenhouse, Lever, Ashby and SmartRecruiters. The first one that returns jobs is saved to `registry.json`. Companies with no standard feed get a custom fetcher in `custom_sources.py` (Amazon uses its public `search.json` API).

**Polling.** `poller.py` fetches every feed in parallel and normalizes each posting into one shape. It filters by level (new grad, early career, APM), role family and US location. It dedupes on the posting URL against `data/seen.json`, so a role is only ever reported once.

**Syncing to Google Sheets.** `sheets_sync.py` mirrors new rows into a Google Sheet through a service account. The sync is **append-only**: it reads the URLs already in the Sheet and adds only missing ones. Edits made by hand in the Sheet, such as the Status column, are never overwritten.

**Removing closed postings.** `prune_closed.py` deletes a row only on positive evidence that the posting closed. It checks evidence in this order:
1. The posting is still in a live feed, so it's kept.
2. Otherwise it asks the posting's own site. A Workday `cxs` API 403 means closed (not bot-blocking). It also checks the Greenhouse, Ashby and SmartRecruiters APIs.
3. Only if the site can't answer, it falls back to "missing from a healthy feed."

Rows that have a manual Status are never deleted. The first run removed 725 of 1,712 stale rows.

**Other tabs.** `rotational_sheet.py` tracks rotational PM and engineering programs and flips a program to OPEN when the poller sees a matching posting. `startups_sheet.py` and `radar.py` build a startup digest with live open-role counts. `swe_fair.py` sweeps a career fair's company list for early-career roles and pay ranges.

## Reliability: the six-day silent outage

launchd runs `run.sh` every 600 seconds, but **it will not start a new cycle while the previous one is still running**. In September 2026, `sheets_sync.py` blocked forever on a half-open socket to Google, because gspread and requests set no timeout by default. The tracker stopped for six days, and nothing appeared in any log.

The fix has two layers:
- `socket.setdefaulttimeout(60)` in every Google-facing script that `run.sh` runs each cycle, so a single stuck connection can't block forever.
- A wall-clock watchdog in `run.sh`. Each step runs under `run_capped <seconds>`, which kills it if it overruns and logs a `WATCHDOG` line. A run can fail, but it can no longer hang.

To check the tracker is healthy, compare `tail data/poller.log` against `date`. A gap longer than about 30 minutes while the laptop is awake means something is wrong.

## Setup

```bash
pip install -r requirements.txt
python discover.py            # builds registry.json from companies.txt
python poller.py              # first run seeds data/seen.json and data/jobs.xlsx
```

To sync to Google Sheets (optional):
1. Create a Google Cloud service account, download its key as `gcp-service-account.json` in this folder, and enable the Sheets API.
2. Create a blank Google Sheet and share it (Editor) with the service account's email.
3. Copy `sheet_config.example.json` to `sheet_config.json` and fill in the Sheet ID. Set `alert_email` if you want a Mail.app email on new roles, or leave it empty.

Both `gcp-service-account.json` and `sheet_config.json` are gitignored.

To run every 10 minutes on macOS, point a launchd agent at `run.sh` with `StartInterval` set to `600`. Set `PYTHON` in the agent's environment if your Python isn't at `/usr/local/bin/python3`.

## Files

| File | What it does |
|---|---|
| `discover.py` | Finds each company's ATS provider and slug, then writes `registry.json` |
| `poller.py` | Fetches, filters and dedupes postings, then appends to `data/jobs.xlsx` |
| `custom_sources.py` | Fetchers for employers without a standard ATS feed |
| `sheets_sync.py` | Append-only mirror of `jobs.xlsx` into the Google Sheet |
| `prune_closed.py` | Evidence-based removal of closed postings from both copies |
| `rotational_sheet.py` | Tracks rotational programs and auto-flips them to OPEN |
| `applied_sheet.py` | Keeps an "Applied" tab from a local seed file |
| `startups_sheet.py`, `radar.py` | Startup digest with live open-role counts |
| `swe_fair.py` | Career-fair sweep of early-career roles and pay ranges |
| `run.sh` | launchd entry point, with a watchdog on every step |
