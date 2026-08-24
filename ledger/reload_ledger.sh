#!/usr/bin/env bash
# reload_ledger.sh - the read-only LOADER half of a sync.
#
# `sync-all` runs the PRODUCERS (QBO -> Bill Tracker.xlsx, QBO -> Notion invoices). Those write the
# systems, NOT the ledger the dashboard reads. This script loads those just-written sources INTO the
# ledger, so `sync-all` becomes one command instead of "sync, then remember to reload the ledger".
# It is the same loader set the dashboard's Resync ("reload") runs, minus costs.
#
# Fast + no Touch ID: WIP master Excel, Bill Tracker.xlsx, Notion AR (--no-qbo), Notion CRM.
# COSTS ARE NOT HERE: `load_costs.py` is a 90-day QBO pull (slower); keep it on the dashboard Resync
# (or run it yourself) so a routine sync-all stays quick. Continues past a single loader's failure
# and reports a summary + a non-zero exit if anything failed.
set -uo pipefail
root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root"

rc=0
step() { printf '\n\033[36m-- %s --\033[0m\n' "$1"; }
step "WIP master -> ledger";  python3 ledger/load_wip_master.py           || rc=1
step "Bills -> ledger";       python3 ledger/load_bill_tracker.py         || rc=1
step "Invoices -> ledger";    python3 ledger/load_invoices.py --no-qbo    || rc=1
step "Customers -> ledger";   python3 ledger/load_customers.py            || rc=1

if [ "$rc" -eq 0 ]; then
  printf '\n\033[32mledger reload: OK\033[0m\n'
else
  printf '\n\033[31mledger reload: one or more loaders failed (see above)\033[0m\n'
fi
exit "$rc"
