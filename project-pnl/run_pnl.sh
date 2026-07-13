#!/bin/bash
# run_pnl.sh — wrapper for the per-project P&L export (mirrors run_invoice_sync.sh).
#
# Interactive (a person ran it in a terminal):
#   run the export directly — its own output is already styled (banner, phases,
#   color-coded counts, per-project sections, final summary panel).
#
# Non-interactive (piped / redirected / scheduled):
#   plain run with a banner and tee to a rotating log OUTSIDE the project folder
#   (~/Library/Logs/Proficient/project-pnl/), so nothing is written into the
#   Claude-visible project folder.
#
# Usage (pass project #s and/or a dropped report straight through):
#   project-pnl MFD177 RP7186 CP672
#   project-pnl --dry-run MFD177
set -u
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # self-locating — works from any clone, any user
cd "$DIR"

# stdout is a terminal → run with the script's built-in styled output.
if [ -t 1 ]; then
  exec /usr/bin/env python3 "$DIR/project_pnl_export.py" "$@"
fi

# Otherwise (scheduled / piped) → plain run, banner + tee to the log file.
LOG_DIR="$HOME/Library/Logs/Proficient/project-pnl"
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/project_pnl.log"
TS="$(date +'%Y-%m-%d %H:%M:%S')"

{
  echo ""
  echo "========================================================================"
  echo "  PROJECT P&L EXPORT RUN @ $TS"
  echo "========================================================================"
} | tee -a "$LOG_FILE"

/usr/bin/env python3 "$DIR/project_pnl_export.py" "$@" 2>&1 | tee -a "$LOG_FILE"
EXIT=${PIPESTATUS[0]}

{
  echo ""
  echo "  EXIT $EXIT @ $(date +'%Y-%m-%d %H:%M:%S')"
} | tee -a "$LOG_FILE"

exit "$EXIT"
