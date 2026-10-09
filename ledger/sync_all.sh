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
# The progress view (ledger/sync_view.py): one line per step, the same look as AP / AR.
# Sundays it also reconciles (COUNT(*) per entity vs the mirror), never after a failed refresh.
recon=""; [[ "$(date +%u)" == "7" ]] && recon="--reconcile"
"$ACB_PY" ledger/sync_view.py mirror $recon; mir=$?

# The office server owns AP/AR once the writer file on the Accounting share
# says {"writer": "server"} (docker/README) - the Mac then stands down on
# both so the two never write the trackers / Notion at once; the mirror and
# the ledger reload stay here (the ledger lives on this Mac).
writer="$("$ACB_PY" -c 'import json
from pathlib import Path
p = Path("/Volumes/Accounting/_automation/writer.json")
try: print(str(json.loads(p.read_text()).get("writer") or "").strip().lower())
except Exception: print("")' 2>/dev/null)"
if [[ "$writer" == "server" ]]; then
  c "1-2/4  AP + AR - the office server runs these (writer file says server)"
  echo "  skipped here - Accounting/_automation/server-status.json has its last run"
  ap=0; ar=0
else
  c "1/4  AP - bill tracker (must precede AR)"
  bash "$base/bill-tracker/run_tracker.sh" "$@"; ap=$?

  c "2/4  AR - invoice sync + AR Aging"
  bash "$base/invoice-sync/run_invoice_sync.sh" "$@"; ar=$?
fi

c "3/4  Ledger - reload the spine (so the dashboard matches)"
if [[ "$*" == *dry-run* ]]; then
  echo "  (dry-run: skipping ledger reload)"; led=0
else
  ACB_DRIVES_CHECKED=1 bash "$base/ledger/reload_ledger.sh"; led=$?
fi

c "summary"
(( mir == 0 )) && ok "MIRROR  QBO mirror refresh"  || bad "MIRROR  QBO mirror refresh" "$mir"
(( ap  == 0 )) && ok "AP      bill tracker"        || bad "AP      bill tracker" "$ap"
(( ar  == 0 )) && ok "AR      invoice sync"        || bad "AR      invoice sync" "$ar"
(( led == 0 )) && ok "LEDGER  reload -> dashboard" || bad "LEDGER  reload" "$led"
(( mir || ap || ar || led )) && exit 1
exit 0
