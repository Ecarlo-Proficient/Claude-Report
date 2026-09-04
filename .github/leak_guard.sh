#!/usr/bin/env bash
# leak_guard.sh - THE data-leak gate. Single source of truth for the patterns
# and exclusions; ci.yml and preflight.sh both call this file, so the two can
# never drift. Edit HERE, never fork the regex into another script.
#
# Blocks: dollar amounts with cents (>= $1,000), millions-scale dollars,
# NON-ROUND six-figure dollars, FEIN-shaped ids, and street addresses.
# Philosophy: real figures have cents or land on odd dollars; examples are
# round ($100,000 / $250,000 stay legal) or written as ~$Nk / $1.2M.
# Business findings belong in the owner's vault, data dumps in
# ~/Library/Logs/Proficient - never in this repo, and never in a STATUS.md.
#
# git grep -P, never plain grep -P: macOS /usr/bin/grep has no -P and exits 2,
# which `if grep ...` reads as "no match" - a silently dead gate. git bundles
# PCRE on macOS and Linux both. And -P, not -E: \b is dead in git grep's ERE.
#
# Usage: leak_guard.sh            scan the working tree (tracked files)
#        leak_guard.sh --cached   scan the index (staged files)
#        leak_guard.sh HEAD       scan a committed tree
set -uo pipefail

# Six-figure term: $NNN,NNN with no cents/continuation, EXCEPT exact thousands
# ($NNN,000) - round numbers are how examples are written, odd ones are real.
PATTERNS='\$[0-9]{1,3}(,[0-9]{3})+\.[0-9]{2}|\$[0-9]{1,3},[0-9]{3},[0-9]{3}|\$[0-9]{3},(?!000(?![0-9,.]))[0-9]{3}(?![,.0-9])|\b[0-9]{2}-[0-9]{7}\b|\b[0-9]{3,5} [A-Z]{3,}( [A-Z]{3,})* (ROAD|STREET|DRIVE|AVENUE|TRAIL|COURT|LANE|CIRCLE|BOULEVARD)\b'
EXCLUDES=(':!.github' ':!*.example.json')

# git grep wants OPTIONS before the pattern and REVISIONS after it, so the
# three call forms are spelled out rather than passing "$@" through one slot.
case "${1:-}" in
  "")       git grep -nP "$PATTERNS" -- "${EXCLUDES[@]}" ;;
  --cached) git grep --cached -nP "$PATTERNS" -- "${EXCLUDES[@]}" ;;
  *)        git grep -nP "$PATTERNS" "$1" -- "${EXCLUDES[@]}" ;;
esac
rc=$?
case "$rc" in
  0) echo "::error::Real-looking dollar figures, FEIN-shaped ids, or street addresses in tracked files. Genericize them (round the figure or write ~\$Nk) or move the data to the vault / ~/Library/Logs/Proficient."
     exit 1 ;;
  1) echo "data-leak guard: clean"
     exit 0 ;;
  *) echo "::error::git grep -P itself failed (rc=$rc) - the guard did NOT run; treating as a failure."
     exit 1 ;;
esac
