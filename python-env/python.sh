# python-env/python.sh - THE interpreter resolver. SOURCE it, never run it:
#
#   . "$base/python-env/python.sh"      then:   "$ACB_PY" some_script.py
#
# Sets ACB_PY to the suite's own Python (~/.venvs/proficient/bin/python). On
# every call it confirms the environment is the right version with the current
# package list, and builds or updates it if not (python-env/setup.sh). If it
# cannot, the calling script stops HERE with one clear message - never a wall
# of ModuleNotFoundError tracebacks from every step (2026-09-24).
#
# Every shell entry point in this repo must use "$ACB_PY" - never a bare python3.
# .github/interpreter_guard.sh enforces that on every push and in CI.

_acb_env_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ACB_VENV="${ACB_VENV:-$HOME/.venvs/proficient}"
ACB_PY="$ACB_VENV/bin/python"
export ACB_VENV ACB_PY
if ! bash "$_acb_env_dir/setup.sh" --ensure; then
  printf '\n  \033[31mStopped: the Python environment is not ready (see above).\033[0m\n' >&2
  printf '  Fix it, then rerun:  bash "%s/setup.sh"\n\n' "$_acb_env_dir" >&2
  exit 3
fi
unset _acb_env_dir
