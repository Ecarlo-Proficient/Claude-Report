# STATUS - python-env

Shared progression record for the suite's one Python environment. Update it in
the SAME commit as any change to this folder.

---

## DONE / FINALIZED

- **Built and adopted (2026-09-24).** Trigger: a `brew install ffmpeg` on 09/23 made
  Homebrew's Python 3.14 the `python3` on PATH. The packages lived under the system
  Python 3.9 (`~/Library/Python/3.9`), so every sync step died on import.
  Immediate fix that day: `brew unlink python@3.14`. Permanent fix, this folder:
  - Python **3.14** (Homebrew `python@3.14`, marked installed-on-request), environment
    at `~/.venvs/proficient`, every package pinned in `requirements.txt`. That
    replaces the root `requirements.txt` and `ledger/requirements.txt`, which were
    unpinned. `invoice-sync/requirements.txt` stays as the container's subset, now
    pinned to match.
  - Every entry point converted to `"$ACB_PY"` via `python.sh`: `ledger/sync_all.sh`,
    `ledger/reload_ledger.sh`, `ledger/open_ledger.command`,
    `ledger/build_ledger_app.command`, `ledger/app/ledger_app.py` (was a hard-coded
    `/usr/bin/python3`), `bill-tracker/run_tracker.sh`,
    `invoice-sync/run_invoice_sync.sh`, `project-pnl/run_pnl.sh`, `wip/run_cp_wip.sh`,
    `.github/preflight.sh`.
  - `.github/interpreter_guard.sh` added to CI and the pre-push hook (gate 4/4).
  - CI moved from 3.11 to `python-version-file: python-env/PYTHON_VERSION`. The Docker
    image moved from `python:3.11-slim` to `python:3.14-slim` (the image build itself
    is not yet re-tested, see OPEN ISSUES).
  - Every "missing package, run pip3 install --break-system-packages ..." hint in the
    tools now says `bash python-env/setup.sh`.
  - Machine side (not in the repo): `~/.zshrc` puts `~/.venvs/proficient/bin` first on
    PATH, and the `statement-reconcile` / `audit-job` helpers run on the environment.
  - **Verified on 3.14 before landing:** all `.py` files compile; ruff gate clean;
    `sync-all --dry-run` with `/usr/bin` as the ONLY PATH entry (mirror refresh +
    Bill Tracker green); invoice sync `--dry-run` (118 invoices, 0 errors); the full
    ledger reload (all 10 loaders) against a scratch copy of the ledger DB; the
    dashboard (`/`, `/api/healthtab`, `/api/qboaudit`, `/api/graph` all 200);
    `project-pnl --dry-run CP672`; `--help` on the WIP readers, statement reconciler,
    job audit and loan sync; every `shared/` module imports.
- **Old 3.9 leftovers removed (2026-09-24, owner OK).** `~/Library/Python/3.9` (the system-Python
  packages) and `statement-reconciler/.venv` moved to the Trash; `sync-all --dry-run` green after.
- **The developer's clone:** the owner is advising them. Two commands after pulling dev:
  `brew install python@3.14`, then `bash python-env/setup.sh`.

- **Pre-push gate 5/5 = the tests (2026-09-24).** `.github/preflight.sh` now runs `pytest tests` on
  python-env's pinned pytest, the same rule as CI's "Run tests" step, so a red test stops the push
  instead of turning GitHub red.

## OPEN ISSUES

- (none on the Mac side)

## TO DO

- **Container (the developer):** `docker/` moved to `python:3.14-slim` on paper but the image
  has not been rebuilt. It is being reworked for a new architecture that also has to bring the
  QBO mirror into the container. That is the developer's lane and a separate piece of work.
