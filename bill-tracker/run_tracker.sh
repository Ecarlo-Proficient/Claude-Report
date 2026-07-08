#!/bin/bash
# run_tracker.sh — entry point for the bill tracker sync (sync-ap).
#
# Launches the visual viewer (sync_view.py), which runs excel_bill_sync.py,
# renders colorful emoji phases to the terminal, AND appends the raw stream to
# the log. Mirrors the invoice sync's run_invoice_sync.sh.
#
# Used by:
#   - manual runs from terminal (also via the `sync-ap` alias)
#   - launchd schedule (com.proficient.billtracker.plist)
#
# Output workbook:
#   /Users/sebas/Library/CloudStorage/OneDrive-ProficientConcrete,LLC/Automations-/Bill Tracker.xlsx
#
# Logs live OUTSIDE the project folder (sync_view.py owns run.log):
#   ~/Library/Logs/Proficient/bill-tracker/
#
# History:
#   2026-05-13 — pivoted notion_bill_sync.py → excel_bill_sync.py.
#   2026-05-29 — output moved to OneDrive.
#   2026-06-18 — visual viewer (sync_view.py) added; logging moved into it.

set -u
DIR="/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/bill-tracker"
cd "$DIR"

# exec so the viewer's exit code propagates straight through to the caller.
exec /usr/bin/env python3 "$DIR/sync_view.py" "$@"
