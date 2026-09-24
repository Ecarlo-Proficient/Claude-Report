#!/usr/bin/env bash
# sync_all.sh - THE sync. One alias (`sync-all`), four steps, in order:
#
#   0/4  QBO mirror refresh   the ONE read of QBO (shared/qbo_mirror): the change
#                             feed since the last stamp. Sundays also reconcile
#                             (COUNT(*) per entity vs the mirror). Every tool
#                             below reads the mirror, so this is the only pull.
#   1/4  AP  bill tracker     Bill Tracker.xlsx (must precede AR: the AR Aging
#                             tab reads it for the vendor column)
#   2/4  AR  invoice sync     Notion + Open_Invoices.xlsx
#   3/4  Ledger reload        the loaders, so the dashboard matches
#
# Args pass to AP and AR (e.g. --dry-run, which also skips the ledger reload).
# Each step runs even if the one before failed; the summary says which failed
# and the exit code is non-zero if any did. Lives in the repo so the sequence is
# versioned; the shell alias is one line pointing here (owner 2026-09-17:
# "concise and deliberately made aliases that do multiple scripts in sequence").
set -uo pipefail
base="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$base"
. "$base/python-env/python.sh"   # THE interpreter -> $ACB_PY (never a bare python3)
c() { printf '\n\033[36m\033[1m== %s ==\033[0m\n' "$1"; }
ok() { printf '  \033[32m✓\033[0m %s\n' "$1"; }
bad() { printf '  \033[31m✗\033[0m %s (exit %s)\n' "$1" "$2"; }

# A missing drive PAUSES before any work: reconnect, Enter to retry, q to close (owner 2026-09-23)
"$ACB_PY" "$base/shared/paths.py" --require accounting,onedrive --for "sync-all" || exit 2

c "0/4  QBO mirror - refresh (the one read of QBO)"
"$ACB_PY" ledger/refresh_mirror.py; mir=$?
if [[ "$(date +%u)" == "7" && $mir -eq 0 ]]; then
  c "0/4  QBO mirror - Sunday reconcile"
  "$ACB_PY" ledger/refresh_mirror.py --reconcile || mir=$?
fi

c "1/4  AP - bill tracker (must precede AR)"
bash "$base/bill-tracker/run_tracker.sh" "$@"; ap=$?

c "2/4  AR - invoice sync + AR Aging"
bash "$base/invoice-sync/run_invoice_sync.sh" "$@"; ar=$?

c "3/4  Ledger - reload the spine (so the dashboard matches)"
if [[ "$*" == *dry-run* ]]; then
  echo "  (dry-run: skipping ledger reload)"; led=0
else
  bash "$base/ledger/reload_ledger.sh"; led=$?
fi

c "summary"
(( mir == 0 )) && ok "MIRROR  QBO mirror refresh"  || bad "MIRROR  QBO mirror refresh" "$mir"
(( ap  == 0 )) && ok "AP      bill tracker"        || bad "AP      bill tracker" "$ap"
(( ar  == 0 )) && ok "AR      invoice sync"        || bad "AR      invoice sync" "$ar"
(( led == 0 )) && ok "LEDGER  reload -> dashboard" || bad "LEDGER  reload" "$led"
"$ACB_PY" ledger/refresh_mirror.py --status 2>/dev/null | sed -n 2p
(( mir || ap || ar || led )) && exit 1
exit 0
