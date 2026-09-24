#!/bin/bash
# run_invoice_sync.sh — wrapper for the invoice sync.
#
# Interactive (a person ran it in a terminal):
#   launch the VISUAL viewer (sync_view.py) — phases, live progress bar, color,
#   and a crash report on failure. The Python sync writes its own rotating
#   sync.log regardless, so file logging is preserved.
#
# Non-interactive (launchd schedule / piped / redirected to a file):
#   run plain with a banner and tee to invoice_sync.log, same as before.
set -u
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # self-locating — works from any clone, any user
cd "$DIR"

# stdout is a terminal → show the visual front-end.
base="$(cd "$DIR/.." && pwd)"
. "$base/python-env/python.sh"   # THE interpreter -> $ACB_PY (never a bare python3)
# A missing drive PAUSES before any work: reconnect, Enter to retry, q to close (owner 2026-09-23)
"$ACB_PY" "$base/shared/paths.py" --require onedrive,accounting --for "sync-ar (Open_Invoices.xlsx on OneDrive; the AR Aging tab reads the Bill Tracker on the Accounting share)" || exit 2

if [ -t 1 ]; then
  exec "$ACB_PY" "$DIR/sync_view.py" "$@"
fi

# Otherwise (scheduled / piped) → plain run, banner + tee to the log file.
LOG_DIR="$HOME/Library/Logs/Proficient/automation-worker"
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/invoice_sync.log"
TS="$(date +'%Y-%m-%d %H:%M:%S')"

{
  echo ""
  echo "========================================================================"
  echo "  INVOICE SYNC RUN @ $TS"
  echo "========================================================================"
} | tee -a "$LOG_FILE"

"$ACB_PY" "$DIR/run_invoice_sync.py" "$@" 2>&1 | tee -a "$LOG_FILE"
EXIT=${PIPESTATUS[0]}

{
  echo ""
  echo "  EXIT $EXIT @ $(date +'%Y-%m-%d %H:%M:%S')"
} | tee -a "$LOG_FILE"

exit "$EXIT"
