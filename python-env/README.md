# python-env - the one Python every tool runs on

One interpreter, one package list, one environment, for the whole suite.

| File | What it is |
|---|---|
| `PYTHON_VERSION` | The Python the suite runs on (`3.14`). Homebrew's `python@3.14` builds it; CI reads the same file. |
| `requirements.txt` | Every package, pinned exactly. CI installs this same file. |
| `setup.sh` | Builds, checks or repairs the environment at `~/.venvs/proficient` (override: `ACB_VENV`). |
| `python.sh` | Sourced by every shell entry point: checks the environment (about 0.06 s), heals it if needed, sets `$ACB_PY`. |

## Why (2026-09-24)

Every tool used to start with a bare `python3`, meaning "whichever `python3` comes
first on PATH today". On 09/23 a `brew install ffmpeg` pulled in Homebrew's Python
3.14 as a dependency. It took over `python3` with none of the packages installed,
and every step of `sync-all` died with `ModuleNotFoundError`. Nothing in the repo
had changed. The machine had.

The fix is structural, so it cannot recur:

1. **A named interpreter, never the ambient one.** The environment is built from
   `/opt/homebrew/opt/python@3.14`, Homebrew's versioned formula, through its `opt/`
   link. Patch updates (3.14.7 -> 3.14.8) carry the environment along. Nothing that
   happens to `python3` on PATH can reach it. `python@3.14` is marked
   installed-on-request, so removing ffmpeg can't auto-remove it.
2. **Every entry point goes through `python.sh`.** `sync-all`, `sync-ap`, `sync-ar`,
   `project-pnl`, `cp-wip`, the ledger launcher, the Project Ledger app and the
   pre-push hook all run `"$ACB_PY"`. Python code starts child processes with
   `sys.executable`.
3. **It heals itself or stops with ONE clear line.** If the environment is missing,
   on the wrong version, or behind a changed `requirements.txt`, `python.sh` rebuilds
   or updates it before the tool runs. If that's impossible (for example Homebrew's
   `python@3.14` is gone), the run stops before doing any work and prints the exact
   fix. No more walls of tracebacks.
4. **A gate keeps it that way.** `.github/interpreter_guard.sh` runs in CI and in the
   pre-push hook. It fails any shell entry point that starts a bare `python`/`python3`,
   any Python subprocess that starts one by name, any launchd plist that runs Python
   directly, and a container pin that disagrees with this list.
5. **CI tests what the Mac runs:** the same `PYTHON_VERSION` and the same
   `requirements.txt`.

## Everyday use

Nothing. The wrappers handle it.

```bash
bash python-env/setup.sh --check
```

```bash
bash python-env/setup.sh --rebuild
```

`--check` reports health without changing anything. `--rebuild` deletes the
environment and builds it fresh.

**Adding a package:** add an exact pin (`name==x.y.z`) to `requirements.txt` and
commit. The next run of any tool installs it on every machine that pulls.
If `invoice-sync` needs it inside the container too, add the same pin to
`invoice-sync/requirements.txt`. The guard checks that the two agree.

**Changing the Python version:** change `PYTHON_VERSION`, run
`brew install python@<new>`, then `bash python-env/setup.sh`. Test with the dry
runs listed in `STATUS.md` before landing. CI follows the same file automatically.

**Ad-hoc commands** (`python3 shared/paths.py`, the usage lines in each script's
docstring) resolve to this environment because `~/.zshrc` puts
`~/.venvs/proficient/bin` first on PATH. That is a convenience only. No tool
depends on it.

## A new machine (the developer's clone)

```bash
brew install python@3.14
```

```bash
bash python-env/setup.sh
```

Then add the PATH line to `~/.zshrc` for ad-hoc commands:

```bash
export PATH="$HOME/.venvs/proficient/bin:$PATH"
```
