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
# -x mirrors ci.yml: skip venvs/node_modules so a local env can't fail a gate
# CI never sees. Note local python may be older than CI's 3.11 - that is a
# FEATURE: the repo floor is 3.9+, and 3.9 parsing here enforces it.
CA_SKIP='/(\.venv|venv|env|node_modules|\.git)(/|$)'
if python3 -m compileall -q -x "$CA_SKIP" . >/dev/null 2>&1; then
  echo "   ok"
else
  echo "   FAIL - a .py file does not parse:"
  python3 -m compileall -q -x "$CA_SKIP" . 2>&1 | head -20
  fail=1
fi

echo "== 2/3 critical lint (ruff E9,F63,F7,F82) =="
# CI pins ruff (see ci.yml) so a new release can't newly flag old code. Warn
# when the local copy drifts from the pin - a newer local ruff can block a
# push CI would pass, an older one can miss what CI will catch.
pin="$(sed -n 's/.*ruff==\([0-9][0-9.]*\).*/\1/p' .github/workflows/ci.yml | head -1)"
have="$(python3 -m ruff --version 2>/dev/null | awk '{print $2}')"
if [ -n "$pin" ] && [ -n "$have" ] && [ "$pin" != "$have" ]; then
  echo "   note: local ruff $have != CI pin $pin - align with:"
  echo "         pip3 install --break-system-packages ruff==$pin"
fi
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
# The patterns live in .github/leak_guard.sh - the ONE copy, shared with
# ci.yml, so this gate and CI can never drift. Edit the script, never here.
leak_scan() {  # $1 = label, rest = args passed through to leak_guard.sh
  local label="$1"; shift
  if bash .github/leak_guard.sh "$@"; then
    echo "   ok ($label)"
  else
    echo "   FAIL ($label) - see above. Genericize the figure (round it or"
    echo "   write ~\$Nk), or move the finding to the vault /"
    echo "   ~/Library/Logs/Proficient. Dollar exposures never live in a"
    echo "   STATUS.md."
    fail=1
  fi
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
