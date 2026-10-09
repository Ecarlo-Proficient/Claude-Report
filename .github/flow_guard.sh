#!/usr/bin/env bash
# flow_guard.sh - every tool carries a flow chart (<tool>/FLOW.md, Mermaid) that
# shows how it works, and the chart moves with the code: a commit that changes a
# tool must change that tool's FLOW.md too (the owner 2026-09-30: "every tool
# should have a flow chart that shows how it works and that chart should be
# updated with every commit to dev"). Same standard as STATUS.md and
# docs/ARCHITECTURE.md - a tool change without its chart is an incomplete commit.
#
#   bash .github/flow_guard.sh                 working tree vs HEAD (a manual preflight)
#   bash .github/flow_guard.sh <from>..<to>    every commit in the range (pre-push, CI)
#
# The ONE copy of this rule - preflight.sh and ci.yml both call it.
# A tool with no FLOW.md yet is listed as MISSING; STRICT_MISSING=1 (once every
# tool has its chart) turns that into a failure too.
set -uo pipefail

TOOLS="bill-tracker debt-schedule docker health-dashboard invoice-sync job-auditor keyhelper ledger project-pnl python-env qbo-export shared statement-reconciler synology wip"
STRICT_MISSING="${STRICT_MISSING:-0}"
fail=0
missing=""

check_files() {   # $1 = label (a commit sha or "working tree"), $2 = newline list of changed paths, $3 = tree-ish for "does FLOW.md exist" ("" = the working tree)
  local label="$1" files="$2" at="$3" t
  for t in $TOOLS; do
    printf '%s\n' "$files" | grep -q "^$t/" || continue
    local has=0
    if [ -z "$at" ]; then [ -f "$t/FLOW.md" ] && has=1
    else git cat-file -e "$at:$t/FLOW.md" 2>/dev/null && has=1; fi
    if [ "$has" -eq 0 ]; then
      case " $missing " in *" $t "*) ;; *) missing="$missing $t" ;; esac
      continue
    fi
    if ! printf '%s\n' "$files" | grep -qx "$t/FLOW.md"; then
      echo "   FAIL ($label): $t/ changed but $t/FLOW.md did not - update the flow chart (and its 'Last changed' line) in the same commit"
      fail=1
    fi
  done
}

if [ "$#" -ge 1 ] && [ -n "$1" ]; then
  range="$1"
  commits="$(git rev-list --no-merges "$range" 2>/dev/null)" || { echo "   flow guard: cannot read range $range - skipped"; exit 0; }
  for c in $commits; do
    check_files "${c:0:10}" "$(git diff-tree --no-commit-id --name-only -r "$c")" "$c"
  done
else
  check_files "working tree" "$( { git diff --name-only HEAD; git ls-files --others --exclude-standard; } 2>/dev/null)" ""
fi

if [ -n "$missing" ]; then
  if [ "$STRICT_MISSING" = "1" ]; then
    echo "   FAIL: no flow chart yet for:$missing - add <tool>/FLOW.md"
    fail=1
  else
    echo "   note: no flow chart yet for:$missing (add <tool>/FLOW.md - becomes a failure once every tool has one)"
  fi
fi
[ "$fail" -eq 0 ] && echo "   ok (flow charts move with their tools)"
exit "$fail"
