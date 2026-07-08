# CLAUDE.md — Automate Concrete Business (repo operating manual)

This folder is the **automation suite** for the business: a set of Python tools wired around
QuickBooks Online (QBO), Notion, Teams, Excel/SharePoint, and Synology. It is not one app —
each subsystem has its own README; read that subsystem's README before working in it.

Company identity, divisions, and finance vocabulary live in the **global CLAUDE.md** — don't
restate them here. Business/strategic context lives in session memory, not in this file.

---

## Safety rails (read first — this is where a session can do real damage)

1. **QBO is the source of truth.** Never guess costs-to-date or billed-to-date; pull and verify against QBO.
2. **Never read, ask for, or hard-code secrets.** Auth is a single encrypted Keychain blob via
   `qbo_vault.py` (service `automation-qbo`, label `credentials`) — one Touch ID per run unlocks all
   keys. Notion/Teams get their own blobs (`automation-notion`, `automation-teams`). Use the setup
   scripts (`setup_qbo.py`, `automation-worker/setup_keychain.py`) and metadata-only diagnostics.
3. **QBO is production-only.** No sandbox/env toggle. `quickbooks.api.intuit.com` is hardcoded by design.
4. **Logs and data dumps go to `~/Library/Logs/Proficient/`** — never inside this folder (it's Claude-visible/synced).
5. **Excel outputs are plain:** white/black only, no fills, no hidden rows, label + amount on the same row,
   split into separate sheets rather than crowding one. (Binding — don't re-litigate.)
6. **Shell commands must be complete and copy-paste-ready, with NO inline `#` comments** (they break zsh paste).
   Put explanations in prose outside the code block.
7. **Never overwrite a data file or write to QBO without an explicit confirm/dry-run gate.** Writers
   default to dry-run; QBO writes are gated (e.g. `--commit`, `CONFIRM=Y`).

---

## Subsystem map (detail in each README)

- **QBO export** (`qbo_vault.py`, `setup_qbo.py`, `qbo_export.py`) — one-row-per-line-item txn export →
  OneDrive `-Inbox- Project Report Exports`. `setup_qbo.py --status/--test/--rotate/--purge`.
- **automation-worker/** — Notion + AR. Bid List → RP/CP Field Logs and Project Plans (UPDATE-only,
  gated by the `Send to Field Log` checkbox, ~5-min). Invoice sync (`run_invoice_sync.py`): QBO open
  invoices → two Notion DBs (MFD isolated; Res/Com combined) routed by project-# prefix; sweeps paid;
  archives QBO-deleted (CDC); posts MFD pay events to Teams. `doctor.py` for diagnostics. launchd
  schedules exist but some died after a macOS update → Ted runs `sync-ar` manually. Dockerized at
  **v1.0.0** for Synology (coexistence: `SKIP_EXCEL_EXPORT=1` so the Mac keeps the Excel mirror).
- **bill-tracker/** — AP bills → matched to the GC invoice that authorizes payment → Excel
  (`~/Documents/CompanyHealth/Bill Tracker.xlsx`); launchd Mon–Fri 15:00. Also `statement_reconciler.py`
  (vendor statement PDF ↔ QBO open bills) and `job_coding_audit.py`.
- **wip/** — `build_wip_master.py`, `wip_sync.py` (2 Notion WIP DBs MFD/RP-CP + snapshots; $25K job-borrow
  Teams alert), `qbo_close_list.py` / `qbo_bulk_close.py` (**always exclude MFD — those close by hand**).
- **debt-schedule/ + loan_sync.py** — QBO → `Equipment_Debt_Schedule_v2.xlsx`. Balance = QBO actual
  (no P/I split); QBO mapping gated by `CONFIRM=Y`; ledger idempotent by (TxnId, AcctId). Mac-only.
- **health-dashboard/qbo_health.py** — local company-health xlsx; reuses `qbo_vault`; private path + chmod 600.
- **project-pnl/project_pnl_export.py** — per-project P&L (CP/MFD + RP × budgeted/unbudgeted) → OneDrive
  PROJECT P&Ls. Overhead shown as a final row at **11% of revenue**; cost-code → name mapping.
- **synology/** — file-tree audit. **Always pass `--exclude /Volumes/Proinfo/Items/`** (sensitive).
- **qbo_recode_review.py** — audit-gated job-cost recoder: `--export` xlsx → Ted audits with QBO-name
  dropdowns → `--apply` (validate) then `--apply --commit`. Only `Approved=Y` rows; exact-spelling +
  stale-SyncToken + closed-period guards.
- **docs/** — Notion architecture, the Invoice Tracker system reference, Field Log stage templates.

---

## QBO API gotchas (learned the hard way)

- **Query parser is AND-only** — split OR conditions into multiple calls, merge in Python.
- **Retry transient 5xx + 429** with backoff (the `_api_get` pattern: 1/2/4/8/16s); let 4xx pass through.
- **Class names are spelled out** — `Residential` / `Commercial` / `Multi Family` (not codes). Run
  `_normalize_class()` before any equality check.
- **`CustomerRef.name` on lines is `Parent:Project # Project Name`** — search for the project #, don't
  match from the start of the string.
- **Custom fields:** the API only returns `DefinitionId=1` regardless of approach/minor version, so the
  Draw Period field is unreachable — use **PrivateNote** as the workaround.
- **Bills carry the project # in memo / PrivateNote** (not a period tag); subs are flagged by `"sub"` in
  the bill memo.
- **Claude's QBO connector P&L (for analysis) caps at 100 rows and doubles monthly totals, and name-keyed
  maps collide on duplicate account names** — use the id-keyed row tree (direct children, exclude
  TOTAL-type rows). The repo's own `qbo_export.py` hits the API directly and isn't subject to this.

---

## Division & matching rules

- **MFD = parent customer "Multi Family"** (or `MFD####` project-# prefix). **Do not trust the Class
  field** for division. **MFD closures are always manual** — bulk-close scripts must exclude all MFD.
- **`-FTW` is a separate project** — strict project-# matching only; never family/base-match
  (e.g. `RP7186` and `RP7186-FTW` are two distinct projects).
- **WIP** lives in the Company Files Teams channel / SharePoint as monthly `.xlsb` snapshots (the MS365
  connector can't open `.xlsb` — use the `.xlsx` adjustment or a PDF). QBO shows **billed** margin, not
  **earned** — earned/over-under-billing is computed in the WIP spreadsheet, not the GL.
- **Cost codes:** job-type prefix (SL/PV/FW/PR/WL/CS/MS) + cost number; small consumables roll into a
  single **Supplies** account (no item-level split).

---

## Boundaries

- **Synology `/Volumes/Proinfo/Items/` is off-limits** — never tree-extract or output its filenames;
  pass it via `--exclude` on every Synology run.
- **Obsidian Main Vault is read-only** (`~/Library/Mobile Documents/com~apple~CloudDocs/Documents/Main Vault`
  — note the global CLAUDE.md path is stale). Draft vault content into `AI Brain_Vault/drafts/vault-fills/`
  mirroring vault structure; never write into the Main Vault.
- **Scope is this company only.** Don't reference affiliated entities (pump co., custom homes, etc.) in
  outputs, and keep artifact titles/filenames neutral (no company name).

---

## Conventions

- Python 3.9+; deps via `pip3 install --break-system-packages -r requirements.txt` (per subfolder).
- Each subsystem keeps its own README, `requirements.txt`/venv, and `launchd/` plist where scheduled.
- Build the core happy-path first; don't pre-add heartbeats/fallback monitors before the core is proven.
