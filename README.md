# Automation Suite

Python tools wired around QuickBooks Online (QBO), Notion, Teams,
Excel/SharePoint, and Synology. **One folder = one tool**; the only shared
code lives in `shared/`. Restructured 2026-07-13 — the map below is the law
of the land (full rules in `CLAUDE.md`, live diagram in
`docs/ARCHITECTURE.md`).

Working on this repo across `dev`/`main`? See
[`docs/GIT_WORKFLOW.md`](docs/GIT_WORKFLOW.md) for the exact pull/push/PR
commands and what branch protection blocks.

## The map

| Folder | Tool | Entry points |
|---|---|---|
| `shared/` | **Common code (importable package)** — QBO Keychain vault, per-machine paths, QBO API helpers | `setup_qbo.py` (vault admin CLI), `paths.py` (self-check) |
| `invoice-sync/` | QBO → Notion AR invoice sync + Teams cards (was `automation-worker/`) | `run_invoice_sync.py` / `sync-ar` alias; `doctor.py`; verifiers |
| `bill-tracker/` | AP bills → GC-invoice match → Excel | `excel_bill_sync.py` / `sync-ap` alias; 4 audit scripts |
| `statement-reconciler/` | Vendor statement PDFs ↔ QBO open bills | `statement_reconciler.py` |
| `wip/` | ALL WIP tooling: CP/RP Excel readers + gated invoice close scripts | `cp_wip_reader.py`, `rp_wip_reader.py`, `qbo_close_list.py`, `qbo_bulk_close.py` |
| `project-pnl/` | Per-project P&L workbooks → OneDrive | `project_pnl_export.py` / `project-pnl` alias |
| `debt-schedule/` | Equipment debt schedule workbook + QBO loan sync | `loan_sync.py`, `build_workbook_v2.py` |
| `health-dashboard/` | Local company-health xlsx (private, chmod 600) | `qbo_health.py` |
| `qbo-export/` | One-row-per-line-item transaction export → OneDrive inbox | `qbo_export.py` |
| `one-offs/` | Occasional & not-yet-developed tools (never the repo root) | see its README |
| `synology/` | NAS file-tree audit (always `--exclude` the sensitive path) | `synology_audit.py` |
| `docker/` | Synology container package for the invoice sync (v1.1.0) | `docker compose up -d --build` |
| `docs/` | System references + the living `ARCHITECTURE.md` diagram | — |

**The import rule:** tools never import tools. Entry scripts put the repo
root on `sys.path` and use `from shared import qbo_vault / paths / qbo_api`.
If a second tool wants a file, the file moves to `shared/` — that's the
only trigger.

## First time on a new machine

**1. Paths.** Output locations (OneDrive mirror, CompanyHealth folder) differ
per machine, so they come from config, not code:

```bash
cp machine.env.example machine.env
python3 shared/paths.py
```

Edit `machine.env` (repo root — it stays there, gitignored) to point the two
roots at real folders, then re-run `python3 shared/paths.py` — it prints
where each path resolved from and flags anything missing or not writable.

**2. QBO credentials.** All 4 QBO keys live in **one** encrypted Keychain
blob (service `automation-qbo`) — one Touch ID per run unlocks everything:

```bash
pip3 install --break-system-packages -r requirements.txt
python3 shared/setup_qbo.py            # interactive setup + auth test
python3 shared/setup_qbo.py --status   # what's stored
python3 shared/setup_qbo.py --test     # auth test only
```

Production only — no sandbox, no env selector (`quickbooks.api.intuit.com`
is hardcoded by design). Notion/Teams secrets are separate blobs owned by
the invoice sync — see `invoice-sync/README.md`.

**3. Per-tool setup.** Each tool folder has its own README and (where needed)
`requirements.txt` / venv. Start with the README of the tool you're touching.
