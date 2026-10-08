#!/bin/bash
# run_recheck.sh - the Notion "Re-check" poller (launchd, every 5 minutes).
# Re-runs the vendors whose Re-check box is ticked on the Notion Vendor Statements
# board, then unticks them. Silent and free when nothing is ticked; never prompts
# (a locked Key Helper or an unmounted Accounting share just waits for the next poll).
set -u
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # self-locating - any clone, any user
base="$(cd "$DIR/.." && pwd)"
. "$base/python-env/python.sh"   # THE interpreter -> $ACB_PY (never a bare python3)
LOG_DIR="$HOME/Library/Logs/Proficient/statement-reconciler"
mkdir -p "$LOG_DIR"
exec "$ACB_PY" "$DIR/statement_reconciler.py" --from-notion --no-color >> "$LOG_DIR/recheck.log" 2>&1
