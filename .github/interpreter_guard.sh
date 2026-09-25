#!/usr/bin/env bash
# interpreter_guard.sh - no entry point may start a bare `python3` (2026-09-24).
#
# A bare python3 means "whatever python3 is first on PATH today". On 09/23 a
# `brew install ffmpeg` swapped that to a fresh Homebrew Python with no packages
# and every step of sync-all died. Every tool now starts through
# python-env/python.sh ("$ACB_PY"), and this gate keeps it that way. It runs in
# CI and in .github/preflight.sh (the pre-push hook) - the ONE copy of the rules.
#
# Fails on, in tracked files:
#   *.sh / *.command   a non-comment line that RUNS python / python3 / python3.X,
#                      bare or as /usr/bin/python3, /usr/bin/env python3,
#                      /opt/homebrew/bin/python3 ... (use "$ACB_PY")
#   *.py               a subprocess argv that starts with such an interpreter,
#                      e.g. ["python3", ...] or ["/usr/bin/python3", ...]
#                      (use sys.executable, or bash + python-env/python.sh)
#   *.plist            a ProgramArguments <string> naming python directly
#                      (point it at the tool's .sh wrapper instead)
#   invoice-sync/requirements.txt  a pin that disagrees with python-env's
#
# Exempt: python-env/ (it builds the environment from the pinned Homebrew
# interpreter) and docker/ (the image's own python IS the pinned interpreter).
# Usage/help text inside a .py docstring is not an invocation and is not scanned.
set -uo pipefail
cd "$(git rev-parse --show-toplevel)" || exit 2

fail=0
interp='(/usr/bin/env[[:space:]]+|/usr/bin/|/usr/local/bin/|/opt/homebrew/bin/)?python(3(\.[0-9]+)?)?'

# --- shell entry points ---
while IFS= read -r f; do
  case "$f" in python-env/*|docker/*|.github/interpreter_guard.sh) continue ;; esac
  hits="$(grep -nE "(^|[[:space:];&|(\`])${interp}([[:space:]]|\$)" "$f" \
          | grep -vE '^[0-9]+:[[:space:]]*#' \
          | grep -vE '^[0-9]+:[^"'"'"']*#[^"'"'"']*python' \
          | grep -vE '^[0-9]+:[[:space:]]*(echo|printf)[[:space:]]' || true)"
  if [ -n "$hits" ]; then
    echo "   $f runs a bare python - use \"\$ACB_PY\" (source python-env/python.sh first):"
    printf '%s\n' "$hits" | sed 's/^/      /'
    fail=1
  fi
done < <(git ls-files '*.sh' '*.command')

# --- python code that spawns an interpreter by name ---
py_hits="$(git grep -nE "\[[[:space:]]*[\"']${interp}[\"'][[:space:]]*," -- '*.py' ':!python-env/*' ':!docker/*' || true)"
if [ -n "$py_hits" ]; then
  echo "   a subprocess starts python by name - use sys.executable:"
  printf '%s\n' "$py_hits" | sed 's/^/      /'
  fail=1
fi

# --- launchd templates ---
pl_hits="$(git grep -nE "<string>${interp}</string>" -- '*.plist' || true)"
if [ -n "$pl_hits" ]; then
  echo "   a launchd plist runs python directly - point it at the tool's .sh wrapper:"
  printf '%s\n' "$pl_hits" | sed 's/^/      /'
  fail=1
fi

# --- the container's subset must pin what python-env pins ---
lock=python-env/requirements.txt
sub=invoice-sync/requirements.txt
if [ -f "$lock" ] && [ -f "$sub" ]; then
  while IFS= read -r line; do
    name="$(printf '%s' "$line" | sed -E 's/[[:space:]]*[=<>!~].*//' | tr 'A-Z_' 'a-z-')"
    [ -n "$name" ] || continue
    want="$(sed -E 's/[[:space:]]*#.*//' "$lock" | grep -iE "^${name//-/[-_]}[[:space:]]*==" | head -1 | sed -E 's/;.*//; s/[[:space:]]+$//')"
    have="$(printf '%s' "$line" | sed -E 's/;.*//; s/[[:space:]]+$//')"
    if [ -z "$want" ]; then
      echo "   $sub pins $name, which $lock does not list - add it there first"
      fail=1
    elif [ "$(printf '%s' "$want" | tr 'A-Z_' 'a-z-')" != "$(printf '%s' "$have" | tr 'A-Z_' 'a-z-')" ]; then
      echo "   $sub has '$have' but $lock has '$want' - make them match"
      fail=1
    fi
  done < <(sed -E 's/[[:space:]]*#.*//' "$sub" | grep -vE '^[[:space:]]*$')
fi

[ "$fail" -eq 0 ] && echo "   ok (every entry point runs through python-env)"
exit "$fail"
