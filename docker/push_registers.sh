#!/usr/bin/env bash
# push_registers.sh - send the owner's rule files from the Mac to the office server (2026-09-30).
#
# CompanyHealth stays on the owner's Mac (owner: "my company health on my mac keep on my mac, the synology should be
# the data center"). The AP job on the server needs the rules the owner keeps there, so this copies them ONE WAY,
# Mac -> Accounting/_automation/registers/, which the server mounts. Run it after changing a rule (a draw move, an
# audit exclusion ...); sync-all runs it too. It never copies cost_code_history.json - the server's AP job writes
# that one (server data; the Mac reads it from the share).
set -euo pipefail
src="${ACB_COMPANYHEALTH_DIR:-$HOME/Documents/CompanyHealth}/Registers"
dst="${ACB_ACCOUNTING_BASE:-/Volumes/Accounting}/_automation/registers"
files=(audit_exclusions.json bizdev_cut.json concrete_suppliers.json draw_moves.json job_rulings.json vendor_types.json)
if [ ! -d "$(dirname "$dst")" ] && [ ! -d "${ACB_ACCOUNTING_BASE:-/Volumes/Accounting}" ]; then
  echo "push_registers: the Accounting share is not mounted - connect it and run again"; exit 2
fi
mkdir -p "$dst"
n=0
for f in "${files[@]}"; do
  if [ -f "$src/$f" ]; then
    if ! cmp -s "$src/$f" "$dst/$f"; then cp -p "$src/$f" "$dst/$f.tmp" && mv -f "$dst/$f.tmp" "$dst/$f"; n=$((n+1)); fi
  fi
done
echo "push_registers: $n rule file(s) updated on the server share ($dst)"
