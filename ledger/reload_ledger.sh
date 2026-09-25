#!/usr/bin/env bash
# reload_ledger.sh - the read-only LOADER half of a sync.
#
# `sync-all` runs the PRODUCERS (QBO -> Bill Tracker.xlsx, QBO -> Notion invoices). Those write the
# systems, NOT the ledger the dashboard reads. This script loads those just-written sources INTO the
# ledger, so `sync-all` becomes ONE command that leaves the ledger fully fresh - the same loader set
# the dashboard's Resync runs.
#
# COMPLETE by design (owner 2026-08-27: "i need this to be simple to sync ... the P&L is not
# working due to it needing data ... payments not showing recent payments"). It used to skip the two
# QBO-direct loaders (costs, payments) to stay quick, which is exactly why the P&L and Payments went
# stale while everything else was fresh. They are IN now; the incremental windows keep them cheap:
#   - load_costs   --active --changed-since <90 days>   (active jobs: every bill entered or edited in 90 days, any bill date)
#   - load_payments --months 12                 (rolling year; load_payments DELETE+reloads its
#                                                window, so the window IS the Payments history depth)
# For a fuller/shorter view run either loader by hand with a different window. Continues past a single
# loader's failure and reports a summary + a non-zero exit if anything failed.
set -uo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

base="$root"
. "$base/python-env/python.sh"   # THE interpreter -> $ACB_PY (never a bare python3)
# In a terminal, show the progress view (ledger/sync_view.py - one line per loader, the
# same look as sync-ap / sync-ar); it runs this script again in plain mode and reads it.
# ACB_PLAIN=1 or a pipe/file = the plain full output, as before.
if [ -t 1 ] && [ -z "${ACB_PLAIN:-}" ]; then
  exec "$ACB_PY" "$base/ledger/sync_view.py" reload
fi

# A missing drive PAUSES before any work: reconnect, Enter to retry, q to close (owner 2026-09-23).
# Skipped when the caller (sync-all / the view) already checked, so no prompt is ever hidden.
if [ -z "${ACB_DRIVES_CHECKED:-}" ]; then
  "$ACB_PY" "$base/shared/paths.py" --require accounting,onedrive --for "the ledger reload (Bill Tracker on the Accounting share, the WIP master on OneDrive)" || exit 2
fi
rc=0
since="$(date -v-90d +%F)"   # macOS/BSD date: 90 days ago, for the incremental cost pull
# run <label> <cmd...>: the "-- label --" start line and the "[step-exit N]" end line are
# what sync_view.py reads - keep both formats.
run() {
  local label="$1"; shift
  printf '\n\033[36m-- %s --\033[0m\n' "$label"
  "$@"; local r=$?
  printf '[step-exit %s]\n' "$r"
  [ "$r" -eq 0 ] || rc=1
}
run "WIP master -> ledger"    "$ACB_PY" ledger/load_wip_master.py
run "Bills -> ledger"         "$ACB_PY" ledger/load_bill_tracker.py
run "Invoices -> ledger"      "$ACB_PY" ledger/load_invoices.py --no-qbo
run "Customers -> ledger"     "$ACB_PY" ledger/load_customers.py
run "Costs -> ledger"         "$ACB_PY" ledger/load_costs.py --active --changed-since "$since"
run "Payments -> ledger"      "$ACB_PY" ledger/load_payments.py --months 12
run "Bill payments -> ledger" "$ACB_PY" ledger/load_bill_payments.py
run "Sub LOC -> ledger"       "$ACB_PY" ledger/load_sub_loc.py
run "Health -> ledger"        "$ACB_PY" ledger/load_health.py
run "Attachments -> ledger"   "$ACB_PY" ledger/load_attachments.py --refresh

if [ "$rc" -eq 0 ]; then
  printf '\n\033[32mledger reload: OK\033[0m\n'
else
  printf '\n\033[31mledger reload: one or more loaders failed (see above)\033[0m\n'
fi
exit "$rc"
