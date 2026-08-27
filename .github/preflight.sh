#!/usr/bin/env bash
# preflight.sh - run the CI gates locally, BEFORE pushing.
#
# Mirrors .github/workflows/ci.yml step for step. Every red CI run so far has
# been one of these three catching something a local run would have caught in
# five seconds - the gates only ever ran on GitHub, after the push.
#
# The gates are SEQUENTIAL on GitHub: the first failure hides the rest. This
# script runs ALL of them every time, so one pass shows everything.
#
# Install as a pre-push hook (this clone only - hooks are not tracked by git):
#   ln -sf ../../.github/preflight.sh "$(git rev-parse --git-dir)/hooks/pre-push"
# Bypass once:  git push --no-verify
# Uninstall:    rm "$(git rev-parse --git-dir)/hooks/pre-push"
set -uo pipefail

# Git runs hooks with GIT_DIR set and the cwd wherever it likes; clear those so
# the git calls below see a normal work tree.
unset GIT_DIR GIT_WORK_TREE GIT_INDEX_FILE

# This file is SYMLINKED to .git/hooks/pre-push, so ${BASH_SOURCE[0]} is the
# symlink, not the target - a plain `dirname .. ` lands in .git/ and every gate
# then passes vacuously (no .py files found, git grep with no work tree).
# Follow the symlink chain to the real file first.
src="${BASH_SOURCE[0]}"
while [ -L "$src" ]; do
  dir="$(cd -P "$(dirname "$src")" && pwd)"
  src="$(readlink "$src")"
  case "$src" in /*) ;; *) src="$dir/$src" ;; esac
done
repo_root="$(cd -P "$(dirname "$src")/.." && pwd)" || exit 1
cd "$repo_root" || exit 1

# Refuse to run from the wrong place rather than report a vacuous "clean".
if [ ! -f .github/workflows/ci.yml ]; then
  echo "preflight: ABORT - cwd '$PWD' is not the repo root (no .github/workflows/ci.yml)."
  exit 1
fi

fail=0

echo "== 1/3 syntax (compileall) =="
if python3 -m compileall -q . >/dev/null 2>&1; then
  echo "   ok"
else
  echo "   FAIL - a .py file does not parse:"
  python3 -m compileall -q . 2>&1 | head -20
  fail=1
fi

echo "== 2/3 critical lint (ruff E9,F63,F7,F82) =="
if ! python3 -c "import ruff" >/dev/null 2>&1; then
  echo "   NOT INSTALLED - and this is the gate that fails most often."
  echo "   Install it:  pip3 install --break-system-packages ruff"
  fail=1
else
  out="$(python3 -m ruff check --select E9,F63,F7,F82 . 2>&1)"
  rc=$?
  echo "$out"
  if [ "$rc" -eq 0 ] && printf '%s' "$out" | grep -q "No Python files found"; then
    # A clean exit over zero files is not a pass, it is a misfire.
    echo "   ERROR - ruff scanned NO files; refusing to call that clean"
    fail=1
  elif [ "$rc" -eq 0 ]; then
    :
  elif [ "$rc" -eq 1 ]; then
    echo "   FAIL - undefined names / unreachable code / bad comparisons"
    fail=1
  else
    echo "   ERROR - ruff itself failed (rc=$rc); treating as a failure"
    fail=1
  fi
fi

echo "== 3/3 data-leak guard =="
# Use `git grep -P`, never plain grep: /usr/bin/grep on macOS has NO -P and
# exits 2, which a naive `if grep ...` reads as "no match" - a silently dead
# gate. git bundles PCRE on both macOS and the Linux runner. And -P not -E:
# \b is dead in git grep's ERE, so -E passes locally while failing on CI.
# Keep this regex identical to ci.yml.
PATTERNS='\$[0-9]{1,3}(,[0-9]{3})+\.[0-9]{2}|\$[0-9]{1,3},[0-9]{3},[0-9]{3}|\b[0-9]{2}-[0-9]{7}\b|\b[0-9]{3,5} [A-Z]{3,}( [A-Z]{3,})* (ROAD|STREET|DRIVE|AVENUE|TRAIL|COURT|LANE|CIRCLE|BOULEVARD)\b'
EXCLUDES=(':!.github' ':!*.example.json' ':!project-pnl/project_pnl_export.py')

leak_scan() {  # $1 = label, rest = extra git-grep args
  local label="$1"; shift
  git grep -nP "$PATTERNS" "$@" -- "${EXCLUDES[@]}"
  local rc=$?
  case "$rc" in
    0) echo "   FAIL ($label) - real dollar figures, FEIN-shaped ids, or street"
       echo "   addresses above. Genericize them, or move the finding to the"
       echo "   vault / ~/Library/Logs/Proficient. Repo rule: dollar exposures"
       echo "   never live in a STATUS.md."
       fail=1 ;;
    1) echo "   ok ($label)" ;;
    *) echo "   ERROR ($label) - git grep -P failed (rc=$rc); treating as a failure"
       fail=1 ;;
  esac
}

# Working tree covers tracked files as they stand; --cached covers the index,
# which is the only place a brand-new `git add`ed file shows up.
leak_scan "working tree"
if ! git diff --cached --quiet 2>/dev/null; then
  leak_scan "staged" --cached
fi

if [ "$fail" -eq 0 ]; then
  echo "preflight: clean"
else
  echo "preflight: FAILED - fix the above; CI will reject this push otherwise."
fi
exit "$fail"
