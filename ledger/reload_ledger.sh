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
#   - load_costs   --active --since <90 days>   (only active jobs, last 90 days)
#   - load_payments --months 12                 (rolling year; load_payments DELETE+reloads its
#                                                window, so the window IS the Payments history depth)
# For a fuller/shorter view run either loader by hand with a different window. Continues past a single
# loader's failure and reports a summary + a non-zero exit if anything failed.
set -uo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

rc=0
since="$(date -v-90d +%F)"   # macOS/BSD date: 90 days ago, for the incremental cost pull
step() { printf '\n\033[36m-- %s --\033[0m\n' "$1"; }
step "WIP master -> ledger";  python3 ledger/load_wip_master.py             || rc=1
step "Bills -> ledger";       python3 ledger/load_bill_tracker.py           || rc=1
step "Invoices -> ledger";    python3 ledger/load_invoices.py --no-qbo      || rc=1
step "Customers -> ledger";   python3 ledger/load_customers.py              || rc=1
step "Costs -> ledger";       python3 ledger/load_costs.py --active --since "$since" || rc=1
step "Payments -> ledger";    python3 ledger/load_payments.py --months 12   || rc=1
step "Bill payments -> ledger"; python3 ledger/load_bill_payments.py         || rc=1
step "Sub LOC -> ledger";      python3 ledger/load_sub_loc.py               || rc=1
step "Health -> ledger";       python3 ledger/load_health.py                || rc=1
step "Attachments -> ledger";  python3 ledger/load_attachments.py --refresh || rc=1

if [ "$rc" -eq 0 ]; then
  printf '\n\033[32mledger reload: OK\033[0m\n'
else
  printf '\n\033[31mledger reload: one or more loaders failed (see above)\033[0m\n'
fi
exit "$rc"
