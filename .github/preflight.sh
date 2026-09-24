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

# THE interpreter (python-env): the gates run on the same Python and the same
# pinned ruff as the tools and CI - never whatever python3 is first on PATH.
. "$repo_root/python-env/python.sh"

# ---- what to scan: the working tree, or the commit being pushed? ----
# git invokes pre-push as `pre-push <remote> <url>` with the pushed refs on
# stdin ("<local-ref> <local-sha> <remote-ref> <remote-sha>"). In that mode,
# gate THE PUSH CONTENT, not the working tree: concurrent sessions edit this
# clone side by side (2026-08-27: another session's half-done edit blocked a
# clean push), and CI will judge the commit, never the tree. Run by hand
# (no args / a tty on stdin) it scans the working tree as before - that is
# the mode that catches problems BEFORE they are committed.
MODE=tree
if [ "$#" -ge 1 ] && [ ! -t 0 ]; then
  sha=""
  while read -r _lref lsha _rref _rsha; do
    case "$lsha" in
      "" ) ;;
      *[!0]* ) sha="$lsha"; break ;;   # skip branch deletions (all-zero sha)
    esac
  done
  if [ -n "$sha" ]; then
    WT="$(mktemp -d "${TMPDIR:-/tmp}/preflight.XXXXXX")/wt"
    if git worktree add --detach --quiet "$WT" "$sha" 2>/dev/null; then
      MODE=push
      trap 'cd "$repo_root"; git worktree remove --force "$WT" >/dev/null 2>&1; rm -rf "$(dirname "$WT")"' EXIT
      cd "$WT" || exit 1
      echo "preflight: gating pushed commit ${sha:0:10} (not the working tree)"
    else
      rm -rf "$(dirname "$WT")"
      echo "preflight: could not materialize $sha - falling back to a working-tree scan"
    fi
  fi
fi

fail=0

echo "== 1/5 syntax (compileall) =="
# -x mirrors ci.yml: skip venvs/node_modules so a local env can't fail a gate
# CI never sees. python-env and CI run the SAME version (python-env/PYTHON_VERSION),
# so a file that parses here parses there.
CA_SKIP='/(\.venv|venv|env|node_modules|\.git)(/|$)'
if "$ACB_PY" -m compileall -q -x "$CA_SKIP" . >/dev/null 2>&1; then
  echo "   ok"
else
  echo "   FAIL - a .py file does not parse:"
  "$ACB_PY" -m compileall -q -x "$CA_SKIP" . 2>&1 | head -20
  fail=1
fi

echo "== 2/5 critical lint (ruff E9,F63,F7,F82) =="
# ruff is pinned in python-env/requirements.txt, which CI installs too, so the
# local copy and CI's are the same release by construction.
if ! "$ACB_PY" -c "import ruff" >/dev/null 2>&1; then
  echo "   NOT INSTALLED - rebuild the environment:  bash python-env/setup.sh --rebuild"
  fail=1
else
  out="$("$ACB_PY" -m ruff check --select E9,F63,F7,F82 . 2>&1)"
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

echo "== 3/5 data-leak guard =="
# The patterns live in .github/leak_guard.sh - the ONE copy, shared with
# ci.yml, so this gate and CI can never drift. Edit the script, never here.
GUARD=.github/leak_guard.sh
[ -f "$GUARD" ] || GUARD="$repo_root/.github/leak_guard.sh"
leak_scan() {  # $1 = label, rest = args passed through to leak_guard.sh
  local label="$1"; shift
  if bash "$GUARD" "$@"; then
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

echo "== 4/5 interpreter guard (no bare python3) =="
# The rules live in .github/interpreter_guard.sh - the ONE copy, shared with ci.yml.
bash .github/interpreter_guard.sh || fail=1

echo "== 5/5 tests (pytest) =="
# Same rule as ci.yml's "Run tests" step: run tests/ when it holds any. Runs on
# python-env's pinned pytest, so a red test is caught here, before the push.
if [ -d tests ] && find tests -name 'test_*.py' -o -name '*_test.py' | grep -q .; then
  if out="$("$ACB_PY" -m pytest -q tests 2>&1)"; then
    echo "   ok ($(printf '%s\n' "$out" | tail -1))"
  else
    printf '%s\n' "$out" | tail -30
    echo "   FAIL - a test is red"
    fail=1
  fi
else
  echo "   ok (no tests yet)"
fi

if [ "$fail" -eq 0 ]; then
  echo "preflight: clean"
else
  echo "preflight: FAILED - fix the above; CI will reject this push otherwise."
fi
exit "$fail"
