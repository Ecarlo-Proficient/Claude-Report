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
  Ready to turn in** (green once every bill is paid; see the later entry — the collect-waivers stage
  was removed) — sorted worklist, 40 most-recent shown of 337. The
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

- **Dock on/off switch — `Project Ledger.app` (owner: on-demand, "no always-on", real indicator + clean Quit).**
  The owner rejected an always-on launchd agent and wanted a Dock indicator + one-click open + a clean
  off switch. **`ledger/app/`** holds a real Cocoa app (`ledger_app.py` via **PyObjC + py2app**);
  `build_ledger_app.command` installs the toolkit once, optionally makes an icon from `app_icon.png`,
  and py2app-builds (alias mode) → **`~/Applications/Project Ledger.app`**. The app runs the SAME
  `dashboard.py` server as a **child process**; its **Dock icon present = ON, gone = OFF** (the
  indicator); **Cmd-Q / Quit / log out / shut down / real system sleep** all stop it cleanly
  (`applicationShouldTerminate` + `NSWorkspaceWillSleepNotification` → `pkill -f ledger/dashboard.py`).
  Repo path is baked via the plist `LSEnvironment` (works wherever the .app lives); the child is
  spawned with a **cleaned env** (strip py2app's `PYTHONPATH`/`PYTHONHOME` so the server sees its own
  site-packages, e.g. `requests`). Never runs at login.
  - **Why not the simpler routes (learned the hard way):** an `osacompile` applet **can't stay open**
    from the CLI (it quit instantly and killed the server); a hand-built foreground app **can't Quit
    cleanly** (⌘Q hangs). A proper PyObjC/py2app Cocoa app is the only one that does **both**. Isolation-
    tested each claim before building.
  - **Server hardening (used by the app + the terminal launcher):** `dashboard.py --background`
    double-forks + `setsid` (detaches so a launcher can't reap it; PPID→launchd), and it now quits
    cleanly on **SIGTERM** (`signal.default_int_handler`). `open_ledger.command` uses `--background`.
  - Verified live end-to-end: launch → server 200 + Dock name "Project Ledger" + browser opens; Quit →
    app gone, server 000, child gone. Sleep-stop uses the documented sleep notification (not triggered
    from here). Build artifacts (`ledger/app/{build,dist}`, `app_icon.icns`) are gitignored.
- **"Recommended to sync" freshness flag (owner, weekend-aware).** The My-view data-freshness cards now
  flag a source with a **⟳ Sync recommended** badge (+ amber border, + a "N recommended to sync" note)
  when it is stale **> 48 business-hours** — weekends don't age the data (a Friday load isn't "stale"
  Monday). New client-side `businessHoursSince()` sums only Mon–Fri slices; threshold `STALE_BUSINESS_H
  = 48`. Purely front-end (uses the existing `meta.freshness`); no server change. Verified: unit tests
  (Fri→Mon = 16 business-h vs 64 raw → not flagged; Wed→Fri = 52 → flagged) + the badge/note render.

- **Draws tab — clickable stage tiles + unambiguous "who paid whom" wording (owner).** The 3 stat
  tiles (Ready to turn in / Collect waivers / Pay vendors) are now **clickable → filter the draw list**
  by stage (toggle; "Show all" clears; active tile outlined), mirroring the Liens tab. The owner asked
  what "Paid" meant — so pill text is now direction-explicit via a display map `DRAW_STAGE_LABEL`
  ("Fund in — pay vendors" → **"GC funded → pay vendors"**, "Paid — collect waivers" → **"Vendors paid →
  collect waivers"**), tile subs spell it out ("GC funded — vendors not paid yet" / "vendors paid —
  waivers pending"), and the hint states the one-way flow (GC funds you IN → you pay vendors OUT →
  waivers back). **Internal stage keys are unchanged** (matched in `_STAGE_ORDER`, `DRAW_STAGE_CLASS`,
  `renderHome`, the waiver-toggle recompute) — display-only. Front-end only; verified live (Pay vendors
  → 49→15, active highlight, clearer pills, no console errors). New `kpi-click` CSS.

- **Owner's date format everywhere + Refresh feedback + a judgeable % bar (owner QoL).**
  - **Dates NEVER year-first** (binding owner preference): new `fmtDate(v, withTime)` renders every
    displayed date as **weekday, abbr-month day, year** — "Mon, Aug 10, 2026" (+ 12h time) — applied to
    the meta line (report/loaded), Data-freshness cards, P&L "last pulled", detail report_date, and the
    draw pay/GC-funded cells. Parses ISO components into a LOCAL date (no `new Date("YYYY-MM-DD")` UTC
    off-by-one). ISO stays only in filenames/keys. (Memory: dates-never-year-first.)
  - **Refresh** now gives feedback (button → "Refreshing…", then a toast "Refreshed · ledger loaded
    <date>") + honest tooltip — it re-reads the ledger instantly; it does NOT re-pull QBO (that's a
    sync). Fixed the wiring (onclick had passed the MouseEvent as `isAuto`).
  - **`% complete` bar** is now full-width (`.pct-bar`, min 130px) with a **visible track**, so the fill
    level is judgeable at a glance; over-100% caps full + red (163.1% reads clearly). Verified live.

- **AR money-IN per draw + the Draws-tab table redesign (owner: "connect systems, don't re-pull QBO").**
  Owner wanted the Draws view to show what the GC pays HIM (in) next to what he pays vendors (out) — and
  corrected an initial QBO-pull approach: the **Invoice Tracker** (Notion, `invoice-sync`) already mirrors
  every QBO invoice and keeps paid ones 12 months, so the ledger reads THAT, not QBO again.
  `QBO → invoice-sync → Invoice Tracker (Notion) → ledger.billing_event`.
  - **`load_invoices.py`** reads both Invoice Tracker DBs (Res/Com `265b…`, MFD `0f8e…`) via
    `shared/notion_client` (the shared token — **no QBO, no Touch ID**) → `billing_event`, keyed by
    **Invoice #** = `ap_bill_line.invoice_no`, so the draw↔invoice join is exact. `billing_event` schema
    fleshed out (doc_number, project, division, customer, memo, amount=TotalAmt, balance, status
    Unpaid/Partially Paid/Paid, source, loaded_at); empty-table migration = safe drop+recreate. Read-only
    on Notion; full-replace by `source='invoice_tracker'`; `--dry-run` coverage; `--selftest` proves the
    parse + status + draw↔invoice join OFFLINE. **Run live: 310 invoices → 41 of 49 draws matched;
    $14.7M billed · $4.5M still open.** Also fixed the freshness sync-ar path (`Collections/Open_Invoices.xlsx`).
  - **`_fetch_draws`** now joins `billing_event` by Invoice # → each draw carries `billed` (net in),
    `ar_status`, `ar_open`, `ar_date`, `customer`.
  - **Draws tab is now a TABLE** (`renderDraws` rewrite): one row per draw — **caret · Project # · Draw
    memo · Billed (in, green) · Invoice # · Date · Paid out (+paid/N) · Stage** — **green row when fully
    done** (Ready to turn in); **click a row → its bills open underneath** (vendor · bill # · amount ·
    paid · GC-funded · waiver checkbox) with a caption. `drawsExpanded` state; waiver POST preserved.
    Verified live: 49 rows, 41 with billed-in, expand works, dates month-first, no console errors.

- **Draws/Liens UI polish batch (owner).** Draws **grouped by project #** (project header row with
  name + N draws + $in/$out totals; dropped the redundant per-row Project # column; memo strips the
  project #). Tables **fit the boxed width** (default is boxed — owner dislikes full width): Name/memo
  columns truncate with ellipsis, and the draw **stage pill shows the short action** ("Pay vendors" /
  "Collect waivers", full "GC funded → …" on hover) so the Stage column stops clipping. **Per-field
  search boxes** on Draws (Project # · Vendor · Invoice # · Division) and Liens (CP # · Vendor · Invoice
  # · Name/address) — a row must match every filled field (AND); Vendors/Sales stay single. **QBO deep
  links on invoices:** the draws' Invoice # links to `app.qbo.intuit.com/app/invoice?txnId=<id>` (the
  Invoice Id comes from the Invoice Tracker load, on `billing_event`). Verified live.
  - **QBO deep links on BILLS (AP) — DONE.** The earlier "no bill id" assumption was wrong: the Bill
    Tracker's **"Open" column is `=HYPERLINK("…/app/bill?txnId=<id>","↗")`** for every bill. `data_only`
    reads only the cached "↗" glyph, so `load_bill_tracker` now does a **second read-only, formula pass**
    (`read_bill_links`, scanning each row for `app/bill?txnId=`) and stores the link in the new
    `ap_bill_line.qbo_link` column (migration drops+recreates on the missing column; matched **2825/2825**).
    `_fetch_ap` and `_fetch_draws` carry `qbo_link`; the Liens **Invoice #** chip (448/448) and the draw
    **Bill #** cell are now `app/bill` links (new-tab, click-stops-row-expand). A shared `qboLinkCell()`
    helper + `a.invno.qbo-link` hover style. Verified live on both tabs, no console errors.
  - **Draw "done" is now PAID, not paid+waivers (owner).** The `waiver` table is empty (the owner
    doesn't mark waivers in-tool), so the old `waivers == n` gate made **"Ready to turn in" (green)
    unreachable** — every fully-paid draw sat on amber **"Paid — collect waivers"** forever. Per the
    owner ("don't say collect waiver if all bills have been paid, just make it green"), the stage is
    now **`paid == n` → green "Ready to turn in"** and the "Paid — collect waivers" stage is **removed**
    (backend stage logic + `_STAGE_ORDER`; frontend `DRAW_STAGE_CLASS/LABEL/SHORT`, the 3 stat tiles
    now **Ready to turn in · Pay vendors · Awaiting GC**, the My-view action item, and the explainer).
    Per-bill **waiver checkboxes stay** (they persist to `waiver` and update the expanded caption for
    the owner's records) but **no longer gate the color**. Live: 18 ready / 15 pay / 16 awaiting = 49.
  - **NOTE — money-IN blank on some draws is a source gap, not a bug.** Billed-(in) comes from the
    Invoice Tracker (`billing_event`, by Invoice #); a draw whose AR invoice isn't in the tracker shows
    "—" (e.g. Briarwood invoices 33942/34103 are absent from the tracker though older ones are present —
    a data-entry gap in the Invoice Tracker, not a recency cutoff). No QBO fallback by design (owner).
  - **UI fit + readability polish (owner).** (a) The draws STAGE pill was clipping off the right edge —
    the short label was reverted to **"Ready to turn in"** (the verbose "Paid — ready to turn in" pushed
    the column out of view; the full text stays on hover). (b) The Sales **All-customers** table clipped
    its last column (Touches) because long client names blew out col 1 — the Client name now truncates
    (`#salesTable` col 1 `max-width:240px`, ellipsis), so all six columns fit. (c) Pipeline **bars were
    near-invisible** (faint fill, sub-pixel for small counts) — added a **border** to `.bar .bar-fill`
    (+`.over`) and a **7px min-width** on the funnel fill so counts like 1/2/9 show. (d) Sales dates
    (customers table + warm list) now use `fmtDate` — no more ISO `2026-05-29`. Verified live both themes.
  - **Per-rep activity drill on the Sales tab (owner: "Devan report — weekly/daily").** Click a rep in
    "Activity by rep" → a drill (defaults to the busiest-by-touches rep, i.e. the outreach person) with:
    This-week/Last-week/Today/All-time touch tiles (with a vs-last-week trend), a **12-week touch
    timeline** (bordered bars — surfaces the drop-off), a **recent-touch log** (dated notes), **follow-ups
    due**, **going stale (21d+ no contact)**, and their **pipeline by stage**. Backend: `_fetch_sales`
    now sends `touch_log` (every dated touch joined to its customer, rep = `_rep_label(last_edited_by)`)
    and enriches `customers` with `follow_up_date` + `main_status`; all bucketing (Monday-anchored weeks,
    local dates) is client-side in `renderRepActivity()`. Rep is a **runtime value, never hard-coded** —
    the drill works for any rep and the code carries no personal names (names come live from Notion).
    Verified live: auto-features the outreach rep, weekly trend renders, rep-switch works, no console errors.
  - **P&L folded INTO the dashboard — live compute (owner: "super database, one place").** The job
    detail now SHOWS the P&L (was only a link out to the Excel): `_project_pnl(con, proj)` assembles it
    from the spine — **Earned Revenue = contract × %complete** (WIP), **costs from `cost_line`** (QBO
    truth, incl subs, itemized by cost code), **overhead 10% of revenue** (MFD alt = 9% of costs), net
    margin + %; **billed (AR)** shown alongside. Conventions match `project_pnl_export.py` so they
    reconcile. `GET /api/pnl/pl?proj=`; `buildPnlGroup` renders the numbers + a "Costs by code" list
    (uncoded flagged red, subs marked) with the project-pnl **Excel demoted to "Detailed export"**.
    **Cross-platform open (owner's "trick"):** `_os_open()` opens files/folders with the host OS command
    (`open`/`os.startfile`/`xdg-open`) so the same dashboard works on Mac OR Windows; `_pnl_open` gained
    an **Open folder** action (`?folder=1`) — CP resolves onto the Synology Common drive, RP/MFD onto
    OneDrive, per `pnl_paths`. Verified live (CP800 net 2.8%, MFD 9%-on-costs), no console errors.
    See [[ledger-super-database]] memory. **NEXT:** portfolio "P&L" tab; RP/CP **source**-folder links
    (RP is address-matched under `/Volumes/Common/CURRENT PROJECTS/Residential` via `rp_wip_reader`).

## IN PROGRESS
- **P&L super-database (Phase 2/3):** portfolio "P&L" tab (all active jobs + company/division totals);
  Synology **source-folder** links everywhere (CP `Awarded Projects Commercial projects`, RP
  `Residential` address-matched, MFD → OneDrive) via the cross-platform `/api/pnl/open` mechanism.

## TO DO
- **Investigate the ~6 active reconcile mismatches** (QBO cost ≠ WIP figure, e.g. RP6901/RP6440):
  likely a base/-FTW split or a stale WIP cost — owner review; keep dollar specifics out of the repo.
- `budget_line` from the takeoff/ETC extractor by cost code (`shared/takeoff_etc.py` is project-total
  today — needs per-code) → enables budget-vs-actual from the spine.
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
