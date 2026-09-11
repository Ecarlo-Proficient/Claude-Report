#!/bin/bash
# Open Project Ledger — the one-click entry point.
# Starts the local read-only dashboard server if it isn't already running,
# then opens it in your browser. Double-click me (or the copy in CompanyHealth).
# Self-locating: works from wherever the repo lives, no hard-coded paths.
set -e
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PORT="${ACB_LEDGER_PORT:-8787}"
URL="http://127.0.0.1:${PORT}"

if ! curl -s -o /dev/null "${URL}/api/health" 2>/dev/null; then
  echo "Starting the ledger dashboard…"
  # --background makes dashboard.py double-fork into its OWN session, so it
  # survives when this is invoked from a GUI app's `do shell script` (Project
  # Ledger.app) — otherwise macOS reaps the child when the app returns. The call
  # returns on its own (the pre-fork parent exits); no nohup/& needed. Output goes
  # to a log outside the repo so a failed start is never silent.
  LOG_DIR="$HOME/Library/Logs/Proficient/ledger-dashboard"
  mkdir -p "$LOG_DIR"
  python3 "${HERE}/dashboard.py" --no-open --background --port "${PORT}" </dev/null >>"$LOG_DIR/server.log" 2>&1
  for _ in $(seq 1 25); do
    curl -s -o /dev/null "${URL}/api/health" 2>/dev/null && break
    sleep 0.3
  done
fi
open "${URL}"
