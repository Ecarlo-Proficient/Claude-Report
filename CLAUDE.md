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
   scripts (`shared/setup_qbo.py`, `invoice-sync/setup_keychain.py`) and metadata-only diagnostics.
3. **QBO is production-only.** No sandbox/env toggle. `quickbooks.api.intuit.com` is hardcoded by design.
4. **Logs and data dumps go to `~/Library/Logs/Proficient/`** — never inside this folder (it's Claude-visible/synced).
5. **Excel outputs are plain:** white/black only, no fills, no hidden rows, label + amount on the same row,
   split into separate sheets rather than crowding one. (Binding — don't re-litigate.)
6. **Shell commands must be complete and copy-paste-ready, with NO inline `#` comments** (they break zsh paste).
   Put explanations in prose outside the code block.
7. **Never overwrite a data file or write to QBO without an explicit confirm/dry-run gate.** Writers
   default to dry-run; QBO writes are gated (e.g. `--commit`, `CONFIRM=Y`).

---

## Repo structure rules (binding — locked in 2026-07-13)

1. **One folder = one tool.** A tool's scripts, verifiers, launchd plist, wrapper, and README
   live together in its folder. The repo root holds NO loose Python, ever.
2. **`shared/` is the ONLY importable common code** (underscore-named = a real package:
   `qbo_vault.py`, `paths.py`, `qbo_api.py`, `setup_qbo.py`). Entry scripts bootstrap with
   `sys.path.insert(0, <repo root>)` then `from shared import …`. That is the only path hack
   allowed — one hack, one direction, one target.
3. **Tools never import tools.** The moment a second tool needs a file, that file moves to
   `shared/` — never a cross-folder import, never an importlib file-path load.
4. **One-offs live in `one-offs/`** and graduate by earning their own folder — never to the root.
5. **`machine.env` stays at the repo root** (per-machine paths; gitignored). `shared/paths.py`
   resolves it there — new machines: `cp machine.env.example machine.env && python3 shared/paths.py`.
6. **Field Log is GONE** (erased 2026-07-13, the user's decision) — sync code, templates, config
   fields all removed. Don't rebuild it without an explicit ask.

## Subsystem map (detail in each README)

- **shared/** — the common package: `qbo_vault.py` (Keychain blob, one Touch ID per run),
  `paths.py` (per-machine path resolution), `qbo_api.py` (QBO auth + retrying GET, `query_all`,
  report walkers, `PROJ_RE` — used by project-pnl and the WIP readers), `setup_qbo.py`
  (`--status/--test/--rotate/--purge`).
- **invoice-sync/** — the QBO → Notion AR invoice sync (was `automation-worker/`). Open invoices
  → two Notion DBs (MFD isolated; Res/Com combined) routed by project-# prefix; sweeps paid;
  archives QBO-deleted (CDC); posts MFD pay events to Teams. Manual via `sync-ar` (launchd plists
  exist but are .disabled). Its config/clients (`config.py`, `qbo_client.py`, `notion_client.py`,
  `teams_notify.py`, `logger.py`, `state.py`, `version.py`, `sync_view.py`, `doctor.py`,
  `setup_keychain.py`) are tool-local — nothing else may import them. `.env` + `state/` live here
  (legacy fallback reads `../automation-worker/` until old clones move them). Dockerized at
  **v1.1.0** for Synology (`SKIP_EXCEL_EXPORT=1` so the Mac keeps the Excel mirror). Keychain
  service (`proficient-automation-worker`) and log dir (`~/Library/Logs/Proficient/automation-worker/`)
  keep their historical names on purpose.
- **bill-tracker/** — AP bills → matched to the GC invoice that authorizes payment → Excel
  (`~/Documents/CompanyHealth/Bill Tracker.xlsx`); manual via `sync-ap` (launchd scrapped).
  Plus 4 audit scripts (`job_coding_audit.py`, `sub_bill_audit.py`, `item_no_project_audit.py`,
  `duplicate_bill_audit.py`).
- **statement-reconciler/** — vendor statement PDF ↔ QBO open bills.
- **wip/** — ALL WIP tooling. Readers: `cp_wip_reader.py` / `rp_wip_reader.py` write ONLY the
  Test tabs of `WIP - MASTER new.xlsx` on SharePoint (guarded by `wip_excel_guard.py`);
  over/under-billing and job-borrow are computed columns in Excel. Close scripts:
  `qbo_close_list.py` / `qbo_bulk_close.py` (**always exclude MFD — those close by hand**).
  The old QBO→Notion WIP sync is fully deleted (stub + plist gone 2026-07-13).
- **project-pnl/** — per-project P&L (CP/MFD + RP × budgeted/unbudgeted) → OneDrive PROJECT
  P&Ls. Overhead shown as a final row at **10% of revenue** (was 11%, the user 2026-07-16;
  MFD alt view stays 9% on costs); QBO helpers come from `shared/qbo_api.py`. Batch mode:
  `project-pnl active cp|rp|mfd` regenerates every Active project of a division (Active =
  the WIP master's Test-Master STATUS).
- **debt-schedule/** — `loan_sync.py` (QBO → `Equipment_Debt_Schedule_v2.xlsx`, beside it) +
  workbook builders. Balance = QBO actual (no P/I split); QBO mapping gated by `CONFIRM=Y`;
  ledger idempotent by (TxnId, AcctId). Mac-only.
- **health-dashboard/** — `qbo_health.py` local company-health xlsx; private path + chmod 600.
- **qbo-export/** — `qbo_export.py` one-row-per-line-item txn export → OneDrive
  `-Inbox- Project Report Exports`.
- **one-offs/** — occasional / not-yet-developed tools. Currently: `qbo_recode_review.py`
  (audit-gated job-cost recoder: `--export` xlsx → the user audits with QBO-name dropdowns →
  `--apply` then `--apply --commit`; only `Approved=Y` rows; exact-spelling + stale-SyncToken +
  closed-period guards; `get_auth()` still an env stub).
- **synology/** — file-tree audit. **Always pass `--exclude /Volumes/Proinfo/Items/`** (sensitive).
- **docker/** — the invoice-sync container package (build context = repo root; copies `shared/`
  + `invoice-sync/`).
- **docs/** — Notion architecture, the Invoice Tracker system reference, `ARCHITECTURE.md`
  (the living diagram).

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
- **Cost codes live in the QBO ITEM name, NOT the account.** An item-based expense line
  (`ItemBasedExpenseLineDetail`) carries an `ItemRef` whose `name` is our cost code (SL1, PV6, CS1…)
  and has NO line-level `AccountRef`; an account-based line resolves to its account
  (`Job Materials: Concrete`, `Subcontractors Expense: Labor`). To key costs BY cost code (accumulating
  costs, Budget vs Actual) use the one resolver **`cost_leaf(det, account_names)`** in
  `project-pnl/project_pnl_export.py` — account name → `AccountRef.name` tail → **`ItemRef.name` (the
  cost code)** → fallback. NEVER resolve the item to its posting account for cost-code work — that
  collapses every SL#/PV# into one account and breaks the join to the takeoff budget.
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

- **No personal names anywhere in the repo** — code, comments, docs, commit messages (swept
  2026-07-13). People are roles: **"the user"** (owner — pulls, reviews, runs against QBO) and
  **"the developer"** (develops from their own clone). Pre-sweep git history keeps old names —
  left intentionally; rewriting shared history costs more than it's worth.
- **No `/Users/<name>` paths in tracked files.** Shell wrappers self-locate via
  `dirname "${BASH_SOURCE[0]}"`; Python derives paths from `__file__` / `shared/paths.py`;
  launchd plists are TEMPLATES with `/ABSOLUTE/PATH/TO/...` placeholders and a documented
  sed install one-liner (see the plist header comment).
- **Vault drafts never live in this repo** — stage them in the AI working vault's
  `drafts/vault-fills/` (outside the repo), mirroring the Main Vault structure.
- Python 3.9+; deps via `pip3 install --break-system-packages -r requirements.txt` (per subfolder).
- Each subsystem keeps its own README, `requirements.txt`/venv, and `launchd/` plist where scheduled.
- Build the core happy-path first; don't pre-add heartbeats/fallback monitors before the core is proven.

---

## Git / GitHub sync (this folder is a live repo — added 2026-07-09)

This folder is a git repo pushed to a **private GitHub remote** shared with the user's assistant
(the developer), who develops from his own clone; the user pulls and runs against QBO.

**Branch model (simplified 2026-07-10):**
- **`main` — released, stable.** Every change reaches it through a PR with **1 approving review
  and a passing `test` CI check**. No direct pushes, no force pushes, no branch deletion —
  enforced even for repo admins. This is where the user pulls production-ready code from.
- **`dev` — the working branch, direct-pushable by everyone** (the user and the developer). Force pushes
  and deletion are still blocked. No PR or pre-merge check is required to land on `dev`; the
  `test` CI workflow runs on **every push to `dev`** and flags breakage after the fact (red ✗ on
  the commit on GitHub). A red `dev` must be fixed before opening a release PR — `main` still
  hard-requires the check. Feature branches + PRs into `dev` remain fine for larger or riskier
  changes, just no longer mandatory.

**Session flow:**
- **Start of any session that will modify files: `git pull` on `dev` first** — the developer's pushes
  may have changed things. Working on a stale copy creates divergence.
- **End of any session where files changed: run `git status`, then PROMPT THE USER** to land the work
  before wrapping up. Do not let a session end with uncommitted/unpushed changes without explicitly
  asking. Unpushed work is invisible to the developer and is how work gets overwritten.
- Landing work on `dev` (complete, paste-ready — the default):
  `git switch dev` → `git pull` → `git add -A` → `git commit -m "<what changed>"` → `git push`,
  then check the commit shows a green ✓ on GitHub.
- Releasing `dev` → `main`: open a PR from `dev` into `main`; it needs 1 approving review + passing
  `test` CI before it can merge (the author can't approve their own PR — the other person reviews).
  Never try to push `main` directly — branch protection rejects it. Auto-merge is enabled: you can
  flag the release PR to merge itself once review + CI complete.
- Never commit `.env`, data files, logs, or venvs — `.gitignore` enforces this; don't override it.
- **`docs/ARCHITECTURE.md` is the living system map (Mermaid — renders on GitHub).** Any commit
  that adds/removes a script or changes a data flow must update the diagram in the same commit.
  Include this in the end-of-session `git status` check: script changes without a diagram update
  are an incomplete commit.
- The developer's QBO auth uses a **separate Intuit app ("EC-Data Export")** in the user's team workspace.
  Never authorize the user's original app from any other machine — one connection per app+realm;
  re-auth kills the existing token and breaks the user's production runs.
