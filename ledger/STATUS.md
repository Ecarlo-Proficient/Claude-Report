# ledger/ — STATUS

Progression record for the canonical project database. Update in the SAME commit as any
change to this tool (repo rule). Tool-scope only — business/dollar analyses live in the vault.

## DONE / FINALIZED
- **`schema.sql`** — the 6-table spine (`project`, `cost_code`, `budget_line`, `cost_line`,
  `billing_event`, `wip_snapshot`) + `v_wip_latest` view. Portable across SQLite and Postgres
  (natural keys, ISO-text timestamps, 0/1 booleans, `DROP VIEW`+`CREATE VIEW`, `ON CONFLICT`).
- **`load_wip_master.py`** — lands the FINAL WIP master Test tabs into `project` + `wip_snapshot`.
  - CP←`Test - CP`, RP←`Test - RP`, MFD←`Test-Master`; each project read once from its richest tab.
  - Rows filtered to `^(MFD|CP|RP)\d+(-FTW)?$` → all legend/total/section rows excluded.
  - Excel opened **read-only**; upserts idempotent by `project_no` / `(project_no, report_date)`.
  - `--dry-run` (write nothing) and `--show N` (sample after load).
- **Verified against the 2026-08-07 master:** 170 projects (Commercial 48, Residential 119,
  Multi Family 3), report date parsed, re-run leaves 0 duplicate snapshot keys, `v_wip_latest`
  and division/category rollups query correctly.
- `docs/ARCHITECTURE.md` updated (new "Ledger" section + folder map + folder-map diagram).
- **`dashboard.py` + `static/`** — local web dashboard over the ledger (Phase-1 UI / "Rung 1").
  - Portfolio KPIs, division rollup, searchable/filterable/sortable projects table, click-into-job
    detail, click-to-copy cells, CSV export.
  - **Customize panel:** theme (auto/light/dark), accent, font, text size, density, width, widget
    toggles, per-column visibility — saved per person in `localStorage`.
  - Read-only on the DB; binds 127.0.0.1 only; stdlib server (no Flask). No new dependencies.
  - Verified live in the browser (light + dark), 170 projects, detail + settings + sort all working.
- **"Needs attention" widget** (generic, data-driven exposure rules — not stored findings):
  Underbilled / Overbilled / Over budget (costs>ETC) / Borrowing cash. Each chip shows a count +
  subtotal and click-filters the projects table (composes with the dropdown filters); "Clear filter"
  resets it. Toggleable like the other widgets. Table cues: % complete as an inline bar (red past
  100%), pure-job-borrow in red, underbillings in green. Verified live (Underbilled → 10 rows).
  - NOTE: the app-preview sandbox can't run this server (it needs the DB + `shared/` outside
    `.preview`); run it directly with `python3 ledger/dashboard.py`. A `.claude/launch.json` entry
    (`ledger-dashboard`) exists but launch.json is untracked/local.

- **`ap_bill_line` + `load_bill_tracker.py`** — AP + lien feed from `Bill Tracker.xlsx`
  (Bills + Inventory display sheets → 2,814 lines, $5.0M open AP, 448 on the lien clock).
  - Read-only on Excel; full-replace by `source='bill_tracker'` (idempotent, mirrors the file).
  - `v_ap_by_project` view; dashboard **AP & liens** widget (open-AP stats + lien watchlist ordered
    by urgency, red past-due pills) + AP line in each job's detail. Verified live.
  - **Deliberately NOT cost_line:** Bill Tracker excludes subs (measured 25–98% short of WIP cost
    per job), so it can't state job cost. Job cost stays in wip_snapshot; complete cost_line waits
    for the qbo-export pull. 284 AP project#s are off-WIP (closed/older) — kept, no FK on project_no.

- **`shared/qbo_costs.py` + `load_costs.py`** — complete cost load, by cost code, incl. subs.
  - `cost_leaf` MOVED out of project-pnl into `shared/qbo_costs.py` (project-pnl imports it back —
    byte-compatible, compiles clean). The ONE resolver both tools share, so they can't drift.
    Engine adds `is_cost_code`, `cost_code_meta`, `build_account_map`, `pull_expense_txns`,
    `cost_lines_from_txns` (network-free, unit-testable), `iter_cost_lines`.
  - `schema.sql`: `cost_line` fleshed out (txn_type, account, vendor, description, source, loaded_at)
    + `v_cost_by_project` / `v_cost_by_code` views.
  - `load_costs.py`: pulls QBO Bills + Purchases → `cost_line` keyed by cost code; scoped full-replace
    (idempotent); `--active/--division/--project/--since/--dry-run`; **`--selftest` proves the whole
    pipeline OFFLINE** (fabricated txns → codes resolved → cost_line written → reconciles $25k=$25k).
    Reconciles loaded cost vs `wip_snapshot.costs_to_date` per project after each load.
  - CLAUDE.md updated (cost_leaf now in shared/qbo_costs; ledger subsystem bullet added).
  - **Run against live QBO 2026-08-08** (owner, `--active`): cost_line populated across active
    projects; **90 of 96 active projects with a WIP cost reconcile within 5%** of
    wip_snapshot.costs_to_date. Residual = per-job attribution differences (a handful of RP/CP
    jobs where QBO-sourced cost ≠ the WIP figure) — surfaced to the owner for review; specific
    dollar findings stay OUT of the repo (scope rule).
- **Dashboard cost-code drill** — `/api/data` now carries a `cost` section (`_fetch_costs`):
  portfolio by-code, per-project by-code, and per-project rollup attached to each project row.
  New **"Costs by code"** widget (portfolio table with % bars), a **QBO Costs / Subs** toggleable
  column pair, and a **"Costs (QBO, by code)"** group in each job's detail showing total loaded /
  subs / WIP costs_to_date (the reconciliation) + the full code breakdown. Verified live — the
  biggest MFD job loads within ~0.3% of its WIP costs_to_date; portfolio spans ~80 cost codes.
- **Costs by code grouped as cost TYPE (parent) → job TYPE (sub)** — the JobTread model: the number
  meaning (Concrete/Labor/Rebar…) is the parent that ALL material rolls up to, the prefix
  (Slab/Paving/Flatwork…) is the collapsible sub, with the cost code shown. Account-based lines land
  under their cost-type parent as an "(account)" sub. Grouping computed server-side in `_fetch_costs`
  via `shared/qbo_costs.job_type_name` + `cost_code_meta`. Verified live. **This mirrors an intended
  QBO restructure** (today QBO cost codes are standalone items routing to categories) — that future
  change is an owner/ops decision, tracked in the vault, not here.
- **Layout: metrics up, bills down (owner).** Widget order is now KPIs → attention → costs → margins
  → division → **AP & liens (bills) moved to position 6, off the top** → projects. The Costs widget
  leads with a **cost-mix** proportional bar + legend (how much each cost type takes, % wise —
  Concrete/Labor/Rebar… as one glance) above the grouped tree.
- **Multi-tab app (owner: "deliberate tabs only a page would contain").** A tab bar splits the app
  into **Overview** (the glance: KPIs · attention · cost-mix · margins · division · projects) plus
  three deep pages: **Costs** (the cost-type→job-type tree + a **code→jobs pivot** — click a code,
  see every job that spent on it), **Liens** (the full collections worklist — every bill on the
  clock grouped into urgency buckets Past-due/≤7d/≤15d/≤30d/…, not a top-N teaser), and **Vendors**
  (spend by vendor, jobs, of-which-subs — from `cost_line.vendor`; new `by_vendor` in the API).
  Active tab persists in localStorage; bills now live only on the Liens tab.
- **Markup + margin (owner).** Derived per job: **planned markup** (contract÷ETC, on cost),
  **planned margin** (GP÷contract, on revenue), **actual markup** (billed÷QBO cost) — as toggleable
  columns, in the Margins widget stats, and in each job's detail. Markup and margin are kept
  distinct on purpose (markup on cost ≠ margin on revenue).
- **Margins & burn** — derived from the loaded QBO costs (client-side; the WIP only ever showed
  *billed* margin, never actual-cost margin). Budget burn (cost ÷ ETC), margin-to-date (billed −
  cost) + margin %, and subs-share as toggleable columns; a portfolio "Margins & burn" widget
  (stats + an over-budget watchlist ordered by burn, worst first); and a "Margin (QBO actual)"
  group in each job's detail. Verified live.
- **Budget-adherence rule (`isOverBudget`)** — the ONE thing the dashboard flags "over budget" with,
  encoding the owner/ops-manager ruling: **flatwork (`-FTW`) budgets are a soft reference, not a
  strict target like slab** (flatwork = sub sent, charged by labor). A `-FTW` job is flagged over
  budget only when its size (max contract/ETC) ≥ ~$15k; slab / CP / MFD stay strict. Used by the
  "Over budget" attention chip and the Margins over-budget watchlist — cut the over-budget list from
  the raw burn>1 count down to the genuinely-actionable jobs (small flatwork false-alarms removed,
  big flatwork + slab + CP kept). Threshold `FTW_BUDGET_FLOOR` is a documented judgment knob. Apply
  the same tolerance when `budget_line` (budget-vs-actual by code) is built.

## IN PROGRESS
- (none)

## TO DO
- **Investigate the ~6 active reconcile mismatches** (QBO cost ≠ WIP figure, e.g. RP6901/RP6440):
  likely a base/-FTW split or a stale WIP cost — owner review; keep dollar specifics out of the repo.
- `budget_line` from the takeoff/ETC extractor by cost code (`shared/takeoff_etc.py` is project-total
  today — needs per-code) → enables budget-vs-actual from the spine.
- `billing_event` from `invoice-sync` (draw period from PrivateNote).
- Postgres deployment decision (Synology container vs. small cloud box) — schema is ready either way.
- Optional: a read-only dashboard over `v_wip_latest` (Phase 3) — DB first, UI later.

## OPEN ISSUES / NOTES
- MFD rows come from `Test-Master`, which has **no STATUS column** → MFD `status` loads as NULL
  (MFD closures are manual anyway). RP/CP status comes from their own tabs.
- `wip_snapshot` stores the master's already-computed figures verbatim (source of truth = the
  sheet, per the "do not generate" instruction). When Phase-2 granular data exists, snapshots can
  be reconciled against a spine-computed WIP as a cross-check.
- The RP tab has no retainage / over-under / earned columns (those are computed on the master);
  for RP those snapshot fields load as NULL by design.
