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
   keys. **That blob is THE key library (the user 2026-07-17): every new integration's key goes in it**
   (QBO keys + `JT_GRANT_KEY` for JobTread so far) — never create a new per-service blob. Notion/Teams
   blobs (`automation-notion`, `automation-teams`) predate the rule and stay as historical exceptions.
   Use the setup scripts (`shared/setup_qbo.py` — `--rotate <KEY>` adds/updates any library key —
   `invoice-sync/setup_keychain.py`) and metadata-only diagnostics.
3. **QBO is production-only.** No sandbox/env toggle. `quickbooks.api.intuit.com` is hardcoded by design.
4. **Logs and data dumps go to `~/Library/Logs/Proficient/`** — never inside this folder (it's Claude-visible/synced).
5. **Excel outputs are plain:** white/black only, no fills, no hidden rows, label + amount on the same row,
   split into separate sheets rather than crowding one. (Binding — don't re-litigate.)
   **Named exception — the `AR Aging` tab of `Open_Invoices.xlsx` (the user 2026-08-05, "add more
   color").** Colour and collapsed rows were asked for there explicitly: green→red aging buckets,
   blue client banding, red/green vendor status, and a grey `n/a` block on RP vendor cells. Colour
   on that tab encodes age or state only — never decoration. **Do not restyle it back to plain**;
   it is the one place the owner overrode this rule on purpose. Every other Excel output stays plain.
5a. **WIP REPORT FORMATTING IS FROZEN (binding, the user 2026-07-31 — "you cannot keep messing with
   formatting").** Every Test tab in `WIP - MASTER new.xlsx` must look like the **original `WIP Master`
   sheet**. That sheet is the ONLY reference — read it, copy it, never invent:
   - **Font:** Tahoma 8 everywhere (headers bold, data regular). No Calibri, no size bumps.
   - **Numbers:** `"$"#,##0_);[Red]("$"#,##0)` (no cents) · percents `0.00%`.
   - **Title block:** B1 = `<company> - <REPORT NAME>`, B2 = `REPORT DATE: <MON DD, YYYY>`, both bold
     and LEFT-aligned, medium rule above row 1 and below row 2. **NO merge-and-center, no big banner
     font, no custom row heights.** The company prefix is read from `WIP Master`!B1 at runtime —
     never hard-code it here.
   - **Header row:** gray `D9D9D9` fill, bold, centered, wrapped, thin borders; freeze below it.
   - **Columns:** ONE commentary column (`NOTES`) — never a NOTES *and* a FLAGS column.
     `TYPE` means Tract/Custom and must never be repurposed.
   - **Links:** QBO deep links on Billed/Costs only — no file/Synology hyperlinks (they bloat the file).
   - **Default view:** Active only; Closed rows filtered AND hidden on open.
   A formatting change to these tabs needs the user's explicit say-so first. Restyling on your own
   initiative is a defect, not an improvement.
5b. **NEVER hand the user an Excel file that hasn't passed `shared/xlsx_verify.assert_clean(path)` as its
   LAST step** — this is the standing guard against Excel's "we found a problem with some content" repair
   prompt (the user 2026-08-17: "same errors still producing"). **Before writing ANY Excel generator, read
   the header of `shared/xlsx_verify.py`** — it is the hard-won list of what NOT to do. The recurring killer:
   **an Excel Table whose `ref`/`autoFilter` still points past the last row after you insert/delete rows.**
   openpyxl `delete_rows` does NOT adjust table ranges, merged cells, or formulas, and `copy_worksheet` does
   NOT copy tables — so after any row surgery, reset every table's `ref` to its new data range (or drop the
   table), fix merges/formulas by hand, then let `assert_clean` catch what you missed. Kill rich-text/multi-run
   cells too. A repair prompt reaching the user is a defect, not a warning.
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
7. **`STATUS.md` beside a tool's README is its shared progression record** (the user 2026-07-21):
   TO DO · IN PROGRESS · DONE/FINALIZED · OPEN ISSUES, so either side (the user's sessions or the
   developer's) can pick up where the other left off. Update it in the SAME commit as any change
   to that tool — a tool change without its STATUS.md update is an incomplete commit (same standard
   as the ARCHITECTURE.md rule). **Scope filter (binding):** only what pertains to the tool itself —
   business/management findings, dollar exposures, and owner-only analyses NEVER go in a STATUS.md
   or anywhere else in this repo; those live in the user's local vault only.

## Subsystem map (detail in each README)

- **shared/** — the common package: `qbo_vault.py` (Keychain blob, one Touch ID per run),
  `paths.py` (per-machine path resolution), `qbo_api.py` (QBO auth + retrying GET, `query_all`,
  report walkers, `PROJ_RE` — used by project-pnl and the WIP readers), `job_lines.py` (the ONE
  "does this line belong to this job" test — strict by default; project-pnl's `--legacy`
  mode adds line-text and bill-memo rules for pre-2025 jobs, with a memo naming 2+ jobs
  SKIPPED not split; shared with `one-offs/legacy_job_cost_pull.py`), `qbo_costs.py` (the ONE
  cost-code resolver `cost_leaf` + the `iter_cost_lines` pull-and-resolve engine — shared by
  project-pnl and the ledger's `load_costs.py`), `notion_client.py` (thin Notion API client —
  create/query/update pages; used by `ledger/sync_actions.py`; invoice-sync keeps its own tool-local
  copy on purpose), `lien_status.py` (the ONE Notion Lien Tracker -> Status resolver: index the
  tracker once, most-escalated pick per invoice by its `Lien` relation; shared by the ledger's
  `load_invoices.py` and invoice-sync's AR Aging Excel so the site and workbook never drift),
  `notion_customers.py` (the ONE parent-client resolver: `{page_id -> title}` cache of a customer
  DB + `relation_title`, so the invoice `Customer` relation resolves to the GC, not `Customer (raw)`;
  shared by the same two so both name the client identically), `setup_qbo.py` (`--status/--test/--rotate/--purge`).
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
  (`Bill Tracker.xlsx` on OneDrive `Automations-/`); manual via `sync-ap` (launchd scrapped).
  **FULL pull incl. subs (2026-08-06):** every bill is fetched; subs are kept off the
  Bills/Inventory/Liens display sheets but flow to the audit — now **six `Audit - …` sheets,
  each a proper Excel Table** (Not Approved · Data Entry · Missing Project · Duplicates ·
  **FW Misplaced** = FW code on a CP/MFD/base-RP slab · Sub No Project). Cost codes (QBO Item
  name) are captured for the audit only, never a display column. The old `sub_bill_audit.py`,
  `item_no_project_audit.py`, `duplicate_bill_audit.py` were folded in and retired;
  `job_coding_audit.py` remains as the interactive `audit-job` per-job drill.
- **statement-reconciler/** — vendor statement PDF ↔ QBO open bills.
- **wip/** — ALL WIP tooling. Readers: `cp_wip_reader.py` / `rp_wip_reader.py` write ONLY the
  Test tabs of `WIP - MASTER new.xlsx` on SharePoint (guarded by `wip_excel_guard.py`);
  over/under-billing and job-borrow are computed columns in Excel. Close scripts:
  `qbo_close_list.py` / `qbo_bulk_close.py` (**always exclude MFD — those close by hand**).
  The old QBO→Notion WIP sync is fully deleted (stub + plist gone 2026-07-13).
  **WIP Review accept/merge (2026-08-25):** each reader + `master_wip_test` gained `--emit-review`
  (compute as usual, diff the tab, dump JSON, NO write) and `--apply-review` (revert disapproved
  fields, then write) modes, backed by `wip_review_common.py` (the ONE diff/revert core — QBO
  fields accepted, PM fields answered). The ledger's WIP Review tab orchestrates them; see the
  ledger bullet. `wip_writer.write_test_cp` is still the ONE guarded, frozen-format tab writer.
- **ledger/** — the canonical project database (Phase 1 of "own the spine, keep the systems as
  peripherals"). `schema.sql` = the portable spine (SQLite + Postgres): `project` · `cost_code` ·
  `budget_line` · `cost_line` · `billing_event` · `wip_snapshot` · `ap_bill_line` · `waiver` ·
  `action` · `customer` · `sales_touch` + views. Loaders (read-only on their sources, idempotent):
  `load_wip_master.py` (WIP master Test tabs → project + wip_snapshot), `load_bill_tracker.py`
  (Bill Tracker → `ap_bill_line`, AP + lien clock + draw/invoice cols — NOT cost truth, subs
  excluded), `load_costs.py` (QBO pull via `shared/qbo_costs` → complete `cost_line` by cost code,
  incl. subs; reconciles to wip_snapshot), `load_customers.py` (Notion Customer List → `customer` +
  `sales_touch`: CRM leads/clients + outreach touch log, per-rep attribution via Notion
  Created/Last-edited-by, needs `ACB_CUSTOMER_LIST_DS_ID`; read-only, `--selftest`).
  `sync_actions.py` mirrors action items (draws-ready MVP) to the Notion "Ledger Actions" DB via
  `shared/notion_client` (needs `ACB_ACTIONS_DS_ID` + the DB shared with the "Automation Integrator"
  integration). `dashboard.py` + `static/` = a local web UI (127.0.0.1) — tabs (My view · Overview ·
  Costs · Draws · Liens · Vendors), READ-ONLY except the ONE write (the owner's waiver mark →
  `waiver`); `open_ledger.command` launcher (co-located in `~/Documents/CompanyHealth/`). **`registry_view.py` + the `Systems` tab** render the vault's systems & process registry (`AI Brain_Vault/02_processes/*.md`) LIVE - parsed per request, never cached, never written back, no ledger table; vault path via `shared/paths.vault_dir()` (`ACB_VAULT_DIR`), read-only. It **replaced the daily markdown digest** (disabled 2026-08-19). **`vault_graph.py` + the `Graph` tab** render the org as a map the SAME way (live, no cache): the whole vault's `[[wikilinks]]` as a force-directed org graph (`ROSTER.md` excluded - no names) plus the mermaid system diagrams IMPORTED from `docs/ARCHITECTURE.md`, all in one self-contained canvas viewer (no JS libraries); `/api/graph` serves it. **The `WIP Review`
  tab** is the WIP update as accept/merge: `/api/wip/review` runs each wip tool's `--emit-review`
  (diff each Test tab, no write), the tab shows every change as WAS→NOW split into **Accept·QBO**
  (costs/billed/retainage, checked) and **PM answers** (contract/COs/ETC, unchecked), and
  `/api/wip/merge` runs `--apply-review` to write only the approved values to Test - CP / Test - RP /
  Test-Master (guarded). Dashboard↔wip is subprocess+JSON only (never an import); JSON in
  `~/Library/Application Support/Proficient/wip-review/`. DB lives
  OUTSIDE the repo (`~/Library/Application Support/Proficient/ledger.sqlite3`; override `ACB_LEDGER_DB`).
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
- **Ref fields come back as an ID ONLY (`{value: "65"}`, no `name`) - resolve the name yourself** by
  pulling that entity once into an `{id -> name}` map. Bit us on `PaymentMethodRef` (all 846 payments
  showed a blank "Payment Type" until we pulled the `PaymentMethod` entity -> Check/ACH/Wire/…). Same
  shape for a Payment's `CustomerRef`, which is the **bare leaf** (the project sub-customer, e.g.
  `RP6676-FTW`) - pull `Customer` and walk `ParentRef` to the top parent for the GC (`LONESTAR GREEN
  HOMES`); stop at the deepest KNOWN ancestor since inactive parents are absent from the active-only
  pull. Both resolvers live in `ledger/load_payments.py` (`_payment_method_map`, `_customer_gc_map`).
- **A `Payment` = money IN as a transaction.** `TotalAmt` + `TxnDate` + `CustomerRef`; the invoices it
  paid are in `Line[].LinkedTxn` where `TxnType == "Invoice"` (the line `Amount` is the slice applied to
  that invoice). It also carries `PaymentRefNum` (check #), `ProjectRef`, `DepositToAccountRef`,
  `UnappliedAmt`, `TxnSource`. `Balance` on the linked `Invoice` = that invoice's open amount.
- **Cost codes live in the QBO ITEM name, NOT the account.** An item-based expense line
  (`ItemBasedExpenseLineDetail`) carries an `ItemRef` whose `name` is our cost code (SL1, PV6, CS1…)
  and has NO line-level `AccountRef`; an account-based line resolves to its account
  (`Job Materials: Concrete`, `Subcontractors Expense: Labor`). To key costs BY cost code (accumulating
  costs, Budget vs Actual) use the one resolver **`cost_leaf(det, account_names)`** in
  **`shared/qbo_costs.py`** (moved out of project-pnl 2026-08-08 when the ledger's `load_costs.py`
  needed the SAME resolver; project-pnl imports it back) — account name → `AccountRef.name` tail →
  **`ItemRef.name` (the cost code)** → fallback. NEVER resolve the item to its posting account for
  cost-code work — that collapses every SL#/PV# into one account and breaks the join to the takeoff
  budget. `shared/qbo_costs.py` also has `is_cost_code`, `cost_code_meta`, and the
  `iter_cost_lines` / `cost_lines_from_txns` pull-and-resolve engine both tools share.
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
- **The Obsidian Main Vault is RETIRED** (merged into `AI Brain_Vault/` 2026-08-04). Do not read
  it, write to it, or re-create the `drafts/vault-fills/` staging folder — there is nothing to
  stage for. Business context goes straight into the vault; the process registry is
  `AI Brain_Vault/02_processes/`.
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
- **Business/vault content never lives in this repo** — it goes in `AI Brain_Vault/` (outside the
  repo). Names never appear anywhere here: people are role handles, and the name↔handle roster is
  `AI Brain_Vault/01_company/ROSTER.md`, which is gitignored.
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

**Multi-session protocol (binding, the user 2026-08-27 - "I need multiple sessions open
running, and after, for it to push and commit flawlessly"):**
- **One session = one worktree.** A repo-MODIFYING session whose main clone already carries
  another session's uncommitted work does not wait and does not tiptoe - it moves itself into
  its own worktree: `bash .github/worktree.sh new <topic>` (or the session's native worktree
  isolation), work there, then land with `bash .github/worktree.sh done` - which rebases onto
  fresh dev, pushes through the gates, retries the push race, and removes the worktree.
  Worktrees inherit this machine's `machine.env` automatically (`shared/paths.py` resolves the
  main clone through the git common dir). Land early, land often - small landings can't conflict.
- **The main clone carries at most ONE modifying session at a time.** Read-only sessions are
  unrestricted. Work that must stay in the main clone: invoice-sync (tool-local `.env`/`state/`),
  anything launchd-run, and any deliverable regen - one writer per output file, always.
- **Never touch a sibling session's dirty files, never `git add -A` on a shared tree** - stage
  your own files by name. Blocked on their file anyway? Message that session
  (ListAgents -> SendMessage) instead of stalling or waiting blind.
- A red MANUAL preflight while pushes go green means a sibling is mid-edit somewhere in the
  tree, not that the gates disagree - the push hook judges only the pushed commit.

**Session flow:**
- **One-time per clone: install the CI gates as a pre-push hook** -
  `ln -sf ../../.github/preflight.sh "$(git rev-parse --git-dir)/hooks/pre-push"`.
  `preflight.sh` runs the same gates as CI (syntax, ruff, the data-leak guard in
  `.github/leak_guard.sh` - the ONE copy of those patterns, never fork it), all of
  them every run, so a red CI never starts on GitHub. Bypass once: `git push --no-verify`.
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
