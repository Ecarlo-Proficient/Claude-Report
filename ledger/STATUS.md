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

- **`customer` + `sales_touch` + `load_customers.py`** — the CRM / sales-pipeline fold (2026-08-09).
  - `schema.sql`: `customer` (one row per Notion Customer List page — identity + pipeline stage +
    Notion `Created by` / `Last edited by` = who sourced / who worked it last, the honest per-rep
    attribution, NO manual Owner property) and `sales_touch` (one row per "History of interactions"
    body line, touch_date parsed) + `v_sales_pipeline` / `v_sales_by_rep` views. Portable (no
    SQLite-only date funcs in views).
  - `shared/notion_client.py`: added `block_children()` (page-body reader) — the shared client now
    reads bodies, not just query/create/update.
  - `load_customers.py`: reads the Customer List **read-only** via `shared.notion_client`; parses
    props + the touch log; idempotent full-replace by `source='notion_customer_list'`;
    `--dry-run/--show/--limit/--all-notes`; **`--selftest` proves parse+load OFFLINE** (no Notion).
    Body fetched only for worked rows (status past Lead/Follow up) by default.
  - `machine.env`: `ACB_CUSTOMER_LIST_DS_ID` added (local, gitignored). Auth reuses the shared
    Notion token (Keychain `proficient-automation-worker/notion`) — verified it can read the list.
  - **Run live 2026-08-09:** 622 customers + 168 touches landed; spine untouched (project 170,
    wip 170, ap 2814, cost 6009). `v_sales_by_rep` for the outreach rep = 141 worked / 112 contacted
    / 12 interested; touch-log dates parse (e.g. "Quote sent 07/15/26" → 2026-07-15).
  - `docs/ARCHITECTURE.md` + README + CLAUDE.md ledger bullet updated in the same commit.
  - Not joined to `project` yet (leads→jobs downstream).

- **Dashboard Sales tab** (2026-08-09) — `/api/data` now carries a `sales` section (`_fetch_sales`):
  pipeline funnel, activity-by-rep (last-editor attribution; the invoice-sync bot relabeled
  "Automation (sync)" via `_rep_label`), warm-account cards with each account's full touch log +
  a stale flag (>21d), and a searchable/filterable all-customers table linking out to Notion. New
  **Sales** tab in `index.html` + `renderSales()` in `app.js` (reuses the existing kpi/bar/table
  helpers, no new deps) + warm-card CSS. **Verified live** against the loaded DB (622 customers /
  168 touches): pipeline, rep table, all 30 warm accounts with touch logs, and the customer table
  all render in the browser (dark + light). Read-only — the tab never writes; edits stay in Notion.
  - QoL (2026-08-09): **Set as default** button + **boxed default** — `baseDefaults()` baseline the
    user snapshots via "Set as default"; Reset + fresh browsers restore to it (shipped default width
    now `boxed`). **Automation accounts kept out of the sales scoreboard** — Notion bots (bare UUID)
    and any name in `ACB_SALES_AUTOMATION_REPS` (machine.env, gitignored — no names in the repo) are
    excluded from Activity-by-rep and shown as "Automation" in the customer table (fixes the raw-UUID
    + the import-account-as-rep). Clickable KPI tiles + pipeline rows → filter the customer table;
    client names are Notion links. Verified live: reps = 5 real people, Interested tile → 30 shown.
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
  see every job that spent on it), **Liens** (the full collections worklist — see the redesign below),
  and **Vendors** (spend by vendor, jobs, of-which-subs — from `cost_line.vendor`; new `by_vendor`
  in the API).
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

- **`open_ledger.command` launcher + CompanyHealth co-location (owner).** Double-click launcher
  (self-locating) that starts the server if down, then opens the browser — one-click, no terminal.
  A thin copy lives in the owner-private `~/Documents/CompanyHealth/` ("Open Project Ledger.command")
  so the ledger opens from the same private cockpit as Company Dashboard + Tracker. Tool code stays
  in the repo; only the front-door launcher sits in CompanyHealth.

- **Draws tab — the draw race-through + the waiver input (owner).** `load_bill_tracker.py` now also
  captures `matched_invoice` / `invoice_status` / `gc_paid_date` / `pay_date` / `bt_key` (ap_bill_line
  migrated in place — rows reload from Excel, lossless). `dashboard.py _fetch_draws` rolls bills up BY
  DRAW (dedup lines → bills), computes the stage — **Awaiting GC funding → Fund in, pay vendors →
  Paid, collect waivers → Ready to turn in** — sorted worklist, 40 most-recent shown of 337. The
  **Draws** tab renders each draw with its bills and a **waiver checkbox** per bill.
- **The ledger's FIRST write surface: `waiver` table + `POST /api/waiver`.** The one place the app
  writes — the owner marks "unconditional waiver in hand." Everything else stays read-only (`mode=ro`);
  the waiver write opens a scoped writable connection, binds 127.0.0.1 only, keyed by
  hash(matched_invoice+vendor+bill) so it survives Bill Tracker reloads. Verified end-to-end: UI
  checkbox → POST → DB persisted (received + timestamp), reverts on failure.

- **"My view" home tab (default) + big-picture→zoom UX pass (owner).**
  - **My view** (default landing): a **data-freshness** strip (sync-ap / sync-ar / WIP-master file
    mtimes + ledger loaded_at, via `_freshness`), clickable **action items** (liens past-due / due-≤7d,
    draws collect-waivers / ready, over-budget, underbilled → each jumps to the right tab+filter), and
    **Working on** = active projects with a division filter. (sync-ar mtime shows "not found" until the
    Open_Invoices path is confirmed.)
  - **Collapse everything by default** — cost tree parents + draw cards start collapsed; expand to zoom.
    Preserved across the live auto-refresh.
  - **Search** on Draws (draw/project/vendor) and Vendors; **division filter** on Draws (MFD/CP only).
  - **Division drill**: click a By-division rollup row → that division's active projects.
  - **Live**: soft auto-refresh every 90s (preserves expand state + active tab).
  - **Vendor TYPE** column (replaces the subs-$ column): "Sub" vs "Supplier: <material>" (Concrete /
    Rebar / …), derived from each vendor's cost mix.
  - **RP is not draws** — the Draws rollup excludes RP (residential bills at completion/milestones, not
    formal draws): 337 → 49 draws (MFD/CP only). Owner ruling.

- **Notion link-out — action items → Notion pages (owner: "folder memory like Notion").**
  MVP wired end-to-end for **draws ready to turn in**. `shared/notion_client.py` (the clean
  invoice-sync client, graduated to shared — invoice-sync keeps its tool-local copy on purpose);
  `ledger/sync_actions.py` finds ready draws (funded + all paid + all waivers) and upserts a page in
  the Notion **"Ledger Actions"** DB keyed by a stable Action Key, reads Status back → the local
  `action` table (new). The dashboard shows a **📄 Notion · <status>** link on tracked draws; the
  ledger stays the RADAR, the thread/notes/done live in the Notion page. `--dry-run` proves it
  offline. `ACB_ACTIONS_DS_ID` in machine.env (gitignored). Ledger stays read-only; Notion writes
  are scoped to the Actions DB.
  - **ONE manual step for the automated path:** share the "Ledger Actions" DB with the Notion
    integration ("Automation Integrator") — open the DB → ••• → Connections → add it — so
    `sync_actions.py` (keychain token) can write unattended. Until then the 404 is expected.
  - Proof page created via the connected workspace (CP585 Draw #4). **Demo data to clean up:** the
    CP585 draw's 2 waivers were test-marked to make it "ready"; uncheck them (and delete/close the
    demo Notion page) — they are not real.

- **Liens tab redesign — clickable stage tiles → one filtered table (owner).** Replaced the stacked
  per-status bucket sections with **clickable stage widgets** ("All on the clock" + Past-due / ≤7d /
  ≤15d / ≤30d / Notice-sent / Lien-filed, each count + $ open, urgency-coloured edge) that filter a
  **single table below** (the `.attn`-tile pattern, mirroring Overview's "Needs attention"). Columns
  reordered to the owner's spec — **CP # · Draw # · Name/Address · Invoice # · Amount** — with Vendor
  trailing and urgency shown as the row's left edge so **CP # stays first**. Invoice # (the vendor
  bill_ref) is a mono chip and Amount is bold — the two the owner said "get lost". Added a search box
  (CP #/draw/name/vendor/invoice). Backend: `_fetch_ap` now also selects `matched_invoice` +
  `invoice_no`; front-end derives Draw # from `invoice_no` (falls back to parsing the draw label) and
  Name from the WIP name (falls back to the draw label). Row-click still opens the job detail. No new
  scripts / data-flow (same `ap_bill_line` → dashboard), so ARCHITECTURE.md unchanged. Verified live:
  header order, tile filtering (Past-due → 110 rows all past-due), search (7 SUNRISE rows), row-click
  detail, no console errors.
- **P&L link — job detail ↔ project-pnl (owner picked A+B: "open + generate, show when last pulled").**
  New `shared/pnl_paths.py` resolves a project's `Project_PnL_<proj>.xlsx` with the SAME rules
  project-pnl writes with (CP → Common-drive awarded folder's `Profit and Loss/`, else OneDrive
  `PROJECT P&Ls/<proj>/`) and returns `{exists, path, mtime, note}` — `mtime` = the **"last pulled"**
  time. Dashboard endpoints (all guarded by `_PROJ_RE`): **`GET /api/pnl`** (find + mtime),
  **`POST /api/pnl/open`** (macOS `open` on the resolved workbook — only ever `Project_PnL_<proj>.xlsx`),
  **`POST /api/pnl/generate`** (runs `project-pnl/run_pnl.sh <proj>` as a **subprocess**, not an import —
  gated behind a `confirm` flag; logs to `~/Library/Logs/Proficient/ledger-pnl/`), **`GET /api/pnl/status`**
  (running/done/error + elapsed; a daemon thread reaps the process). Job detail shows a **P&L
  (project-pnl)** group: *Last pulled <ago · timestamp>* + **Open** + **Generate / Refresh** (confirm
  dialog warns about the QBO pull + Touch ID). Generate is the ONE place the dashboard triggers a QBO
  pull + a file write — QBO stays read-only inside project-pnl; the .xlsx write is the gated action.
  Verified live: find (MFD325 → mtime; CP745 → "not generated yet" + CP note), the `confirm`-required
  gate (400 without it), status idle, and the panel rendering both states. **Generate's first LIVE run
  is owner-driven** (Touch ID on the Mac) — the non-QBO plumbing is verified; the QBO run was not
  triggered from here.
  - **Follow-up — dedupe the resolver:** project-pnl still has its own
    `_resolve_project_out_dir`/`_find_awarded_cp_folder`; it should import `shared/pnl_paths.py`
    (same move as `cost_leaf`→`shared/qbo_costs.py`). Deferred to avoid editing the 326 KB export
    script while a concurrent session is in `shared/`.
  - **Not built — option C** (project-pnl reads cost_line from the ledger instead of re-pulling QBO,
    the "own-the-spine" data-source refactor). Still the strategic direction; larger, separate job.

- **Dock on/off switch — `build_ledger_app.command` (owner: on-demand, "no always-on anything").**
  The owner rejected an always-on launchd agent. Instead a self-locating builder osacompiles a tiny
  stay-open AppleScript applet → **`~/Applications/Project Ledger.app`**: launch → runs
  `open_ledger.command` (start server if down + open browser); **the Dock icon present = ON, gone =
  OFF** (the indicator); **Quit / log out / shut down → stops the server** (`on quit` → `pkill -f
  ledger/dashboard.py`); an **`on idle` watchdog stops it on sleep** (a >90 s gap between 15–20 s idle
  ticks = the Mac slept → stop + quit) and quits if the server dies (keeps the indicator honest).
  Repo-rule clean: the tracked builder self-locates (no `/Users` path); the real path is baked into
  the generated app on the owner's machine only. Never runs in the background. Verified: builds +
  compiles, foreground/Dock-visible (no LSUIElement), launcher path baked, handlers present. First
  GUI launch is the owner's.
- **"Recommended to sync" freshness flag (owner, weekend-aware).** The My-view data-freshness cards now
  flag a source with a **⟳ Sync recommended** badge (+ amber border, + a "N recommended to sync" note)
  when it is stale **> 48 business-hours** — weekends don't age the data (a Friday load isn't "stale"
  Monday). New client-side `businessHoursSince()` sums only Mon–Fri slices; threshold `STALE_BUSINESS_H
  = 48`. Purely front-end (uses the existing `meta.freshness`); no server change. Verified: unit tests
  (Fri→Mon = 16 business-h vs 64 raw → not flagged; Wed→Fri = 52 → flagged) + the badge/note render.

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
