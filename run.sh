#!/bin/bash
# Wrapper launchd calls every 10 minutes (StartInterval 600).
#
# Every step runs under a hard wall-clock watchdog. This matters because launchd
# will NOT start a new cycle while the previous one is still alive: on 2026-09-04
# sheets_sync.py blocked forever on a half-open socket to Google and silently
# froze the whole tracker for six days. A bounded run can fail; it can never hang.
cd "$(dirname "$0")" || exit 1

PY="${PYTHON:-/usr/local/bin/python3}"

# run_capped <seconds> <script> <logfile>
run_capped() {
    local secs="$1" script="$2" log="$3"
    shift 3
    "$PY" "$script" "$@" >> "$log" 2>&1 &
    local pid=$!
    ( sleep "$secs"; kill -9 "$pid" 2>/dev/null ) &
    local watcher=$!
    wait "$pid" 2>/dev/null
    local rc=$?
    kill -9 "$watcher" 2>/dev/null
    wait "$watcher" 2>/dev/null
    if [ "$rc" -ge 128 ]; then
        echo "[run.sh] WATCHDOG: $script exceeded ${secs}s and was killed ($(date '+%Y-%m-%d %H:%M'))" >> "$log"
    fi
    return 0
}

run_capped 480 poller.py           data/poller.log
# every 6h: drop closed postings from jobs.xlsx + the Sheet (refetches all feeds)
run_capped 300 prune_closed.py     data/poller.log --if-due --apply
run_capped 180 sheets_sync.py      data/sheets.log
run_capped 180 rotational_sheet.py data/sheets.log
