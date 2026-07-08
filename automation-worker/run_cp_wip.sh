#!/bin/bash
# run_cp_wip.sh — wrapper for the CP WIP reader.
# Run from ANYWHERE via the `cp-wip` alias; this cd's into the folder for you
# so the module imports resolve — you never have to cd yourself.
#
# Interactive (a person ran it): the reader's own styled output is the viewer.
# Non-interactive (piped / redirected): banner + tee to a rotating log.
set -u
DIR="/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker"
cd "$DIR"

if [ -t 1 ]; then
  exec /usr/bin/env python3 "$DIR/cp_wip_reader.py" "$@"
fi

LOG_DIR="$HOME/Library/Logs/Proficient/automation-worker"
mkdir -p "$LOG_DIR"
LOG_FILE="$LOG_DIR/cp_wip.log"
TS="$(date +'%Y-%m-%d %H:%M:%S')"
{
  echo ""
  echo "========================================================================"
  echo "  CP WIP READER RUN @ $TS"
  echo "========================================================================"
} | tee -a "$LOG_FILE"
/usr/bin/env python3 "$DIR/cp_wip_reader.py" "$@" 2>&1 | tee -a "$LOG_FILE"
EXIT=${PIPESTATUS[0]}
{
  echo ""
  echo "  EXIT $EXIT @ $(date +'%Y-%m-%d %H:%M:%S')"
} | tee -a "$LOG_FILE"
exit "$EXIT"
