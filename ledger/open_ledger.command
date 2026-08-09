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
  nohup python3 "${HERE}/dashboard.py" --no-open --port "${PORT}" >/dev/null 2>&1 &
  for _ in $(seq 1 25); do
    curl -s -o /dev/null "${URL}/api/health" 2>/dev/null && break
    sleep 0.3
  done
fi
open "${URL}"
