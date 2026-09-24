#!/usr/bin/env bash
# python-env/setup.sh - build, check, or repair THE Python environment.
#
# Why this exists (2026-09-24): every tool used to start with a bare `python3`,
# which means "whatever python3 comes first on PATH". A `brew install ffmpeg`
# pulled in Homebrew's Python 3.14 as a dependency, it took over `python3`, and
# sync-all died on every step with ModuleNotFoundError - the packages lived under
# a different Python. The fix is structural: one environment, built from ONE
# named interpreter, and every entry point goes through python.sh - so what
# Homebrew (or anything else) does to `python3` can no longer reach the tools.
#
#   bash python-env/setup.sh             build it if missing, sync it if the
#                                        package list changed (same as --ensure)
#   bash python-env/setup.sh --ensure    quiet when healthy; heals otherwise
#   bash python-env/setup.sh --check     report only - exit 1 if not healthy
#   bash python-env/setup.sh --rebuild   delete the environment and build fresh
#
# Location: ~/.venvs/proficient (override with ACB_VENV). Outside the repo so
# every clone and every worktree on this machine shares it.
set -uo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
want="$(tr -d '[:space:]' < "$here/PYTHON_VERSION")"
lock="$here/requirements.txt"
venv="${ACB_VENV:-$HOME/.venvs/proficient}"
py="$venv/bin/python"
stamp="$venv/.acb-stamp"
mode="${1:---ensure}"

say() { printf '  \033[36mpython-env:\033[0m %s\n' "$*" >&2; }
die() { printf '  \033[31mpython-env: %s\033[0m\n' "$*" >&2; exit 1; }

# The stamp = the interpreter version + the exact package list the environment
# was last synced to. Any edit to requirements.txt or PYTHON_VERSION changes it.
want_stamp() { { echo "$want"; cat "$lock"; } | shasum -a 256 | awk '{print $1}'; }

# The interpreter we build FROM: Homebrew's versioned python@X.Y, through its
# opt/ link so patch upgrades (3.14.7 -> 3.14.8) carry the environment along.
# Never the bare `python3` - that is exactly what broke.
base_python() {
  local prefix cand
  for prefix in "$(brew --prefix 2>/dev/null)" /opt/homebrew /usr/local; do
    [ -n "$prefix" ] || continue
    cand="$prefix/opt/python@$want/bin/python$want"
    [ -x "$cand" ] && { echo "$cand"; return 0; }
  done
  return 1
}

have_version() { "$py" -c 'import sys; print("%d.%d" % sys.version_info[:2])' 2>/dev/null; }

healthy() {
  [ -x "$py" ] || return 1
  [ "$(have_version)" = "$want" ] || return 1
  [ -f "$stamp" ] && [ "$(cat "$stamp")" = "$(want_stamp)" ]
}

build() {
  local base
  base="$(base_python)" || die "Python $want is not installed. Run:  brew install python@$want   then rerun."
  if [ -e "$venv" ]; then
    # Only ever delete something that is recognisably a venv.
    [ -f "$venv/pyvenv.cfg" ] || die "$venv exists but is not a Python environment - move it aside first."
    say "rebuilding $venv on Python $want"
    rm -rf "$venv"
  else
    say "building $venv on Python $want (one time, about a minute)"
  fi
  mkdir -p "$(dirname "$venv")"
  "$base" -m venv "$venv" || die "could not create the environment"
  "$py" -m pip install --quiet --disable-pip-version-check --upgrade pip || die "pip upgrade failed"
}

sync_packages() {
  say "installing the package list (python-env/requirements.txt)"
  "$py" -m pip install --quiet --disable-pip-version-check -r "$lock" \
    || die "package install failed - see the pip output above"
  "$py" -m pip check >/dev/null 2>&1 || say "note: pip check reports a dependency conflict (run: $py -m pip check)"
  want_stamp > "$stamp"
  say "ready - $("$py" --version 2>&1) at $venv"
}

case "$mode" in
  --check)
    if healthy; then echo "python-env: ok - $("$py" --version 2>&1) at $venv"; exit 0; fi
    [ -x "$py" ] || { echo "python-env: MISSING - $venv (run: bash python-env/setup.sh)"; exit 1; }
    [ "$(have_version)" = "$want" ] || { echo "python-env: WRONG VERSION - $(have_version), want $want (run: bash python-env/setup.sh)"; exit 1; }
    echo "python-env: OUT OF DATE - requirements.txt changed since the last install (run: bash python-env/setup.sh)"
    exit 1 ;;
  --rebuild)
    build; sync_packages ;;
  --ensure|"")
    healthy && exit 0
    if [ ! -x "$py" ] || [ "$(have_version)" != "$want" ]; then build; fi
    sync_packages ;;
  *)
    die "unknown option $mode (use --ensure, --check or --rebuild)" ;;
esac
