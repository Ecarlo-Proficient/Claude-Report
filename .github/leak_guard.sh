#!/usr/bin/env bash
# leak_guard.sh - THE data-leak gate. Single source of truth for the patterns
# and exclusions; ci.yml and preflight.sh both call this file, so the two can
# never drift. Edit HERE, never fork the regex into another script.
#
# Blocks: keys and tokens (see KEY_PATTERNS), dollar amounts with cents (>= $1,000), millions-scale dollars,
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

# Keys and tokens (security review 09/29/2026) - a second pattern set with its own error: our key names
# assigned a long value, Notion tokens, Teams / Power Automate webhook signatures, QuickBooks refresh tokens,
# private keys, GitHub / AWS / Slack / AI-service tokens. Keys live in the Keychain key library behind Key
# Helper (keyhelper/), never in a file. The fake values in tests are built at runtime, so they never match.
KEY_PATTERNS='\b(ntn|secret)_[A-Za-z0-9]{40,}\b|webhook\.office\.com/webhookb2/[A-Za-z0-9@/._-]{20,}|\.logic\.azure\.com[^\s"'"'"'<>]*[?&]sig=[A-Za-z0-9%_-]{20,}|powerplatform\.com[^\s"'"'"'<>]*[?&]sig=[A-Za-z0-9%_-]{20,}|-----BEGIN [A-Z ]*PRIVATE KEY-----|\bgh[pousr]_[A-Za-z0-9]{36,}\b|\bgithub_pat_[A-Za-z0-9_]{50,}|\bAKIA[0-9A-Z]{16}\b|\bxox[abprs]-[A-Za-z0-9-]{10,}|\bsk-ant-[A-Za-z0-9_-]{20,}|\bsk-[A-Za-z0-9]{40,}\b|\bAB11[0-9]{10}[A-Za-z0-9]{25,}\b|\b(QBO_CLIENT_SECRET|QBO_CLIENT_ID|QBO_(?:[A-Z]+_)?REFRESH_TOKEN|NOTION_SECRET|JT_GRANT_KEY|MIRROR_KEY|GRAPH_CLIENT_SECRET|TEAMS_[A-Z_]*WEBHOOK[A-Z_]*)\s*[:=]\s*["'"'"']?[A-Za-z0-9+/_.%-]{20,}'

scan() {   # scan <pattern> [--cached|<rev>]
  case "${2:-}" in
    "")       git grep -nP "$1" -- "${EXCLUDES[@]}" ;;
    --cached) git grep --cached -nP "$1" -- "${EXCLUDES[@]}" ;;
    *)        git grep -nP "$1" "$2" -- "${EXCLUDES[@]}" ;;
  esac
}

scan "$KEY_PATTERNS" "${1:-}" | sed -E 's/^([^:]+:[0-9]+:).*/\1 <a key or token - value not shown>/'
krc=${PIPESTATUS[0]}
case "$krc" in
  0) echo "::error::A key or token is in a tracked file (lines above; the value is not printed). Remove it, rotate that key (shared/setup_qbo.py --rotate NAME), and keep keys in the key library - never in a file."
     exit 1 ;;
  1) echo "key/token guard: clean" ;;
  *) echo "::error::git grep -P itself failed (rc=$krc) - the key guard did NOT run; treating as a failure."
     exit 1 ;;
esac

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
