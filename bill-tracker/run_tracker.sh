#!/bin/bash
# run_tracker.sh — entry point for the bill tracker sync (sync-ap).
#
# Launches the visual viewer (sync_view.py), which runs excel_bill_sync.py,
# renders colorful emoji phases to the terminal, AND appends the raw stream to
# the log. Mirrors the invoice sync's run_invoice_sync.sh.
#
# Used by:
#   - manual runs from terminal (also via the `sync-ap` alias)
#     (there is no scheduler — the launchd auto-run was scrapped)
#
# Output workbook:
#   /Volumes/Accounting/Accounts Payable/Bill Tracker.xlsx (the Accounting share - since 2026-09-16)
#
# Logs live OUTSIDE the project folder (sync_view.py owns run.log):
#   ~/Library/Logs/Proficient/bill-tracker/
#
# History:
#   2026-05-13 — pivoted notion_bill_sync.py → excel_bill_sync.py.
#   2026-05-29 — output moved to OneDrive.
#   2026-06-18 — visual viewer (sync_view.py) added; logging moved into it.

set -u
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # self-locating — works from any clone, any user
cd "$DIR"

base="$(cd "$DIR/.." && pwd)"
. "$base/python-env/python.sh"   # THE interpreter -> $ACB_PY (never a bare python3)
# A missing drive PAUSES before any work: reconnect, Enter to retry, q to close (owner 2026-09-23)
"$ACB_PY" "$base/shared/paths.py" --require accounting --for "sync-ap (the Bill Tracker)" || exit 2

# exec so the viewer's exit code propagates straight through to the caller.
exec "$ACB_PY" "$DIR/sync_view.py" "$@"
