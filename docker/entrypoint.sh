#!/usr/bin/env bash
# entrypoint.sh — runs the invoice sync on a loop with graceful shutdown.
#
# Why a shell loop (vs. cron in the container):
#  - Simpler — no extra process supervisor or cron daemon to debug.
#  - Logs are stdout/stderr-native, picked up by Docker's log driver.
#  - SIGTERM (docker stop) terminates the current python run cleanly via `wait`.
#
# Interval is controlled by SYNC_INTERVAL_SECONDS (default 900 = 15 min).
# To run once and exit instead of looping, set SYNC_RUN_ONCE=1.
set -euo pipefail

INTERVAL="${SYNC_INTERVAL_SECONDS:-900}"
RUN_ONCE="${SYNC_RUN_ONCE:-0}"

# Forward SIGTERM/SIGINT to the python child for clean shutdown.
shutdown() {
    echo "[entrypoint] received shutdown signal, exiting after current run…"
    if [[ -n "${child:-}" ]] && kill -0 "$child" 2>/dev/null; then
        kill -TERM "$child"
        wait "$child" 2>/dev/null || true
    fi
    exit 0
}
trap shutdown SIGTERM SIGINT

cd /app/automation-worker

if [[ "$RUN_ONCE" == "1" ]]; then
    echo "[entrypoint] one-shot mode (SYNC_RUN_ONCE=1)"
    exec python3 run_invoice_sync.py
fi

echo "[entrypoint] loop mode — interval ${INTERVAL}s"
while true; do
    python3 run_invoice_sync.py &
    child=$!
    wait "$child" || true
    child=""
    echo "[entrypoint] sleep ${INTERVAL}s before next run"
    # Sleep in foreground so SIGTERM interrupts immediately
    sleep "$INTERVAL" &
    child=$!
    wait "$child" || true
    child=""
done
