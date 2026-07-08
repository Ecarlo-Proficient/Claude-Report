#!/bin/bash
# run_wip_sync.sh — wrapper for wip_sync.py
#
# Used by:
#   - manual runs from terminal
#   - launchd schedule (com.proficient.wip-sync.plist, daily 5:30am CT)
#
# Behavior:
#   - cd's into automation-worker/ so config/state paths resolve correctly
#   - ensures logs/ exists
#   - runs python with `tee` so output is live in terminal AND logged to logs/wip_sync.log

set -u
DIR="/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker"
# Logs live OUTSIDE the project folder (privacy: project folder is AI-session-visible)
LOG_DIR="$HOME/Library/Logs/Proficient/automation-worker"
mkdir -p "$LOG_DIR"

LOG_FILE="$LOG_DIR/wip_sync.log"
TS="$(date +'%Y-%m-%d %H:%M:%S')"

cd "$DIR"
{
  echo ""
  echo "========================================================================"
  echo "  WIP SYNC RUN @ $TS"
  echo "========================================================================"
} | tee -a "$LOG_FILE"

/usr/bin/env python3 "$DIR/wip_sync.py" "$@" 2>&1 | tee -a "$LOG_FILE"
EXIT=${PIPESTATUS[0]}

{
  echo ""
  echo "  EXIT $EXIT @ $(date +'%Y-%m-%d %H:%M:%S')"
} | tee -a "$LOG_FILE"

exit "$EXIT"
