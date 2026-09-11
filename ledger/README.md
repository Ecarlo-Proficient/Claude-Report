# ledger/ — the canonical project database

**The idea:** own the spine, keep the systems as peripherals. QBO stays the books,
JobTread stays the ops shell, Excel goes back to being an export — but the *reconciled
shape of a JOB* (identity + budget + costs + billing + the computed WIP position) lives
here, in one database we own instead of in six vendor data models.

This is **Phase 1** of that: a real schema, and the final WIP master sheet landed into it
so you can watch your actual data live in a database instead of a spreadsheet.

---

## Files

| File | What it is |
|------|-----------|
| `schema.sql` | The whole 6-table spine. **Portable — runs on SQLite and PostgreSQL unchanged.** |
| `load_wip_master.py` | Reads the FINAL WIP master (the Test tabs) → fills `project` + `wip_snapshot`. |
| `load_bill_tracker.py` | Reads `Bill Tracker.xlsx` (Bills + Inventory) → fills `ap_bill_line` (AP + liens). |
| `load_costs.py` | QBO pull → `cost_line` by cost code (incl. subs), via `shared/qbo_costs`. |
| `load_customers.py` | Notion "Customer List" → `customer` + `sales_touch` (CRM leads/clients + outreach touch log, read-only). |
| `load_health.py` | QBO pull → `health_snapshot`: bank cash, retainage GL, P&L blocks, 13-wk cash flow, recurring register - the Health tab's QBO-only layer. |
| `dashboard.py` | Local web dashboard over the ledger — the browser UI (read-only). |

The cost engine itself lives in **`shared/qbo_costs.py`** (`cost_leaf` + `iter_cost_lines`) — the
SAME resolver project-pnl uses, so the ledger and the P&L can never drift.
| `static/` | The dashboard front-end (`index.html`, `style.css`, `app.js`) — no build step. |
| `requirements.txt` | `openpyxl` (SQLite + the web server are stdlib — nothing else to install). |

## The schema (6 tables + 1 view)

```
project          the aggregate root — one row per job. project_no is THE join key.
cost_code        job-type prefix + number, a first-class dimension (not a QBO item-name string)
budget_line      the plan: ETC by cost code
cost_line        COMPLETE spend, one row per QBO expense line, keyed by cost code (incl. subs)
billing_event    AR invoices / draws (append-only, idempotent by invoice id)
wip_snapshot     the COMPUTED WIP position — one row per (project, report_date)
ap_bill_line     vendor bills from Bill Tracker — AP pay status + the lien clock (NOT cost truth)
customer         CRM master — one row per Notion Customer List page (identity + pipeline stage + created/last-edited-by)
sales_touch      outreach touch log — one row per "History of interactions" line (date parsed when present)
v_wip_latest     view: each project joined to its most-recent snapshot
v_ap_by_project  view: open AP + bill counts per project
v_cost_by_project view: loaded QBO cost per project (reconcile vs WIP)
v_cost_by_code   view: per-project cost-code drill (budget-vs-actual base)
v_sales_pipeline view: customer counts by pipeline stage
v_sales_by_rep   view: outreach activity by last editor (per-rep attribution)
```

**What fills what:**
- `project`, `wip_snapshot` ← `load_wip_master.py` (**today** — from the WIP master sheet)
- `ap_bill_line` ← `load_bill_tracker.py` (**today** — from Bill Tracker.xlsx)
- `cost_line` + `cost_code` ← `load_costs.py` (**today** — one QBO pull, incl. subs)
- `customer` + `sales_touch` ← `load_customers.py` (**today** — from the Notion Customer List)
- `health_snapshot` ← `load_health.py` (**today** — the Health tab's QBO-only metric layer)
- `budget_line`, `billing_event` ← later (takeoff-by-code budget; AR/draws)

## Run it

```bash
python3 ledger/load_wip_master.py --dry-run --show 6
```
```bash
python3 ledger/load_wip_master.py --show 8
```

- Default source: the same `WIP - MASTER new.xlsx` the WIP readers write (`WIP_EXCEL_PATH`).
- Default database: `~/Library/Application Support/Proficient/ledger.sqlite3` (outside the repo,
  created on demand). Override with `--db`.
- `--dry-run` parses and reports counts, writes nothing.

Each project is read from its richest tab exactly once — CP from `Test - CP`, RP from
`Test - RP`, MFD from `Test-Master` (MFD has no own tab). Rows are filtered to real project
numbers (`^(MFD|CP|RP)\d+(-FTW)?$`), so every legend / totals / section-break row drops out.

## AP & liens (Bill Tracker)

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/load_bill_tracker.py --show 8
```

Reads the line-level `Bills` + `Inventory` sheets of `Bill Tracker.xlsx` (override with
`ACB_BILL_TRACKER_XLSX`) into `ap_bill_line` — vendor, project, account, amount, open balance,
pay status, and the Texas lien clock per bill. Read-only on Excel; each run **full-replaces**
`source='bill_tracker'` so it mirrors the current file.

The tracker prefixes `Matched Invoice` with a `[TAG] ` on special matches (`[DRAW]`, `[FULLY BILLED]`, `[PUSHED from Draw #3]` - the last is a bill the supplier agreed to carry into a later draw, rule in `<CompanyHealth>/draw_moves.json` via `shared/draw_moves.py`). The loader splits that tag into `match_tag` and keeps `matched_invoice` as the bare "invoice — memo", so every bill on one draw shares the same draw key.

> **Not the cost ledger.** Bill Tracker's display sheets EXCLUDE subs, and for a sub-based labor
> company subs are most of the cost — measured 25–98% short of the QBO WIP truth per job. Job cost
> stays in `wip_snapshot`; the complete `cost_line` (incl. subs + true SL/PV cost codes) comes later
> from a `qbo-export` pull. What this feed uniquely adds is **AP pay status + lien deadlines** the
> WIP snapshot lacks. The dashboard surfaces it as the **AP & liens** widget (open AP, a lien
> watchlist ordered by urgency) and an AP line in each job's detail.

## Costs — complete, by cost code (QBO)

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/load_costs.py --selftest
```
```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/load_costs.py --active --show 20
```

`load_costs.py` pulls QBO expense transactions (Bills + Purchases) and writes one `cost_line` per
expense line — attributed to a project by its line `CustomerRef`, keyed to the cost code by the
shared `cost_leaf()` resolver. This is the **complete** cost source (subs included), and it
**reconciles** to `wip_snapshot.costs_to_date` (printed per project after the load).

- **`--selftest`** runs the whole pipeline **offline** on a throwaway DB — no QBO, no touch to your
  real ledger. Run it any time to prove the wiring.
- A real load needs **one Touch ID** (QBO read-only). Scope with `--active`, `--division cp|rp|mfd`,
  `--project MFD177 CP800`, or `--since 2025-01-01`. `--dry-run` pulls + reconciles without writing.
- Idempotent: a scoped full-replace of `source='qbo'` cost_line for the target projects each run
  (mirrors QBO, drops deleted txns).

The cost engine (`shared/qbo_costs.py`) is shared with project-pnl, so cost-code figures tie between
the ledger and the P&L export by construction.

## The money trail - every line behind Costs / Billed (2026-09-02)

`GET /api/trail?project=CP800&kind=costs|billed|both` (`ledger/trail.py`) returns every
`cost_line` (one QBO Bill / Expense line: date, vendor, bill #, memo, description, code, amount,
"part of a $X bill", qb deep link, scan flag) and every `billing_event` invoice for a project,
chronological, with a **running total per kind** - the red line the page draws against the ETC and
the contract. `&csv=1` gives the same as an Excel-ready CSV. The totals block carries the WIP
report's figures beside QuickBooks' and the two honest comparisons: invoices + retainage held vs
WIP billed (WIP is gross), and costs split into "dated after the report" vs "not explained by
dates". The page (`static/trail.js`, "Show every dollar" in the project drawer) is read-only.
`python3 ledger/trail.py --selftest` proves it offline. The extra `cost_line` columns are additive
and NULL until the next full `load_costs.py` run. No created / edited stamps (parked by the owner).

## Attachments - every scan, on every row (2026-09-03)

`load_attachments.py` indexes every QBO attachment (`Attachable` -> `attachment(etype, txn_id,
attachable_id, file_name)`; Bill, Purchase, Invoice, Payment, BillPayment, ...). Every transaction row in
the dashboard carries a 📎 with its count; the click opens the in-app viewer, which asks
`/api/attachment?id=&type=` for FRESH download links (QBO's expire in minutes, so none is stored) and
previews PDFs / images inline with Open-in-new-tab and Download. `--refresh` forces a new sweep instead of
the week-old disk cache the P&L export also uses. `--selftest` offline.

## CRM — customers & sales pipeline (Notion)

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/load_customers.py --selftest
```
```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/load_customers.py --dry-run --show 10
```

`load_customers.py` reads the Notion **Customer List** (read-only) into `customer` (one row per
lead/client: identity, current pipeline stage, and Notion's own **Created by** / **Last edited by**
system fields — who sourced it / who worked it last, the honest per-rep attribution with no manual
Owner property to maintain) and `sales_touch` (one row per "History of interactions" line in the page
body — the outreach touch log, with the date parsed when the line carries one). `v_sales_pipeline` and
`v_sales_by_rep` make "what has the outreach rep done" a query. This is the pre-project (CRM) half of
"own the spine": leads become jobs downstream, but sales activity now lives in the same database as WIP.

- **`--selftest`** runs the whole parse+load **offline** on a throwaway DB — no Notion, no touch to
  your real ledger. Run it to prove the wiring.
- Idempotent: a full-replace of `source='notion_customer_list'` (customer + sales_touch) each run,
  mirroring the current list. `--dry-run` pulls + reports without writing; `--show N` prints the warm
  (Interested) accounts with their last touches.
- By default the touch-log body is fetched only for **worked** rows (status past `Lead`/`Follow up`);
  `--all-notes` fetches every page's body.
- **Setup:** `ACB_CUSTOMER_LIST_DS_ID` in `machine.env` (the Customer List data-source id), and the
  Notion integration (Keychain `proficient-automation-worker/notion`, the same token `sync_actions.py`
  uses) must have the Customer List shared with it.

## The dashboard (browser UI)

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/dashboard.py
```

Opens `http://127.0.0.1:8787` in your browser (add `--no-open` to skip that, `--port N` to
change the port). It reads the ledger **read-only** and binds to `127.0.0.1` only (not exposed
on the network). What it shows:

- **Portfolio KPIs** — total contract, costs, billed, left-to-bill, net over/(under), active jobs.
- **Health** — the company-health metric layer (the fold-in of the retired Company
  Tracker/Dashboard): **Money In / Money Out / Position / Break-Even** + the **Recurring & Debt**
  register (FIN-12). Sections come preformatted from `/api/healthtab` (one server-side model, no
  drift); most rows derive live from the ledger tables, the QBO-only numbers (cash with its as-of
  stamp, retainage GL, margins, burn/runway, the register) come from the last `load_health.py`
  pull - the tab's **Pull QBO metrics** button, or Console → "Health metrics (QBO)", also in the
  Resync chain. Every metric row click-jumps to the tab holding its detail; the break-even audit
  trail (where every figure came from) is a collapsible table.
- **Needs attention / AP & liens** — exposure chips and the lien watchlist.
- **Costs by code** — the QBO cost ledger by cost code (portfolio table + a per-job breakdown in the
  detail panel showing loaded vs subs vs WIP costs_to_date — the reconciliation, in the UI). Plus
  toggleable **QBO Costs / Subs** columns in the projects table.
- **Margins & burn** — real margin from the QBO costs: budget burn (cost ÷ ETC), margin-to-date
  (billed − cost), margin %, subs-share — as a widget (with an over-budget watchlist), toggleable
  columns, and a per-job detail group. The margin the WIP never showed.
- **By division** — the CP / RP / MFD rollup.
- **Sales** (CRM pipeline, read-only from the Notion Customer List) — pipeline funnel, activity by
  rep (last-editor attribution), warm-account cards with each account's full touch log, and a
  searchable/filterable all-customers table that links out to Notion. Edits stay in Notion; the tab
  never writes.
- **Projects** — searchable, filterable (division / status / category / active-only), sortable;
  click any row for the full job detail (Contract / Budget / Costs / Earned / Billing / Notes).
- **Copy & export** — click any number to copy it; **Export CSV** downloads the current view.
- **Customize** (⚙) — theme (auto/light/dark), accent color, font, text size, density, width
  (**boxed by default**), which widgets show, and which table columns show. Saved per person in the
  browser; **Set as default** snapshots the current view as the baseline that Reset (and a fresh
  browser) restores to.

The dashboard is a view. It never writes the database or the Excel sheet.

## Safety

- The Excel workbook is opened **read-only** — this tool never writes the sheet.
- Upserts are **idempotent**: re-running replaces the same `project` / `(project_no,
  report_date)` rows, never duplicates them.
- The database is local and disposable; delete the file to start clean.

## Postgres

`schema.sql` deploys to Postgres unchanged:

```bash
createdb proficient_ledger
psql proficient_ledger -f ledger/schema.sql
```

The loader targets SQLite for the zero-install spike. Pointing the load at Postgres is a
driver swap (`psycopg`) with the identical `INSERT ... ON CONFLICT` statements — the SQL is
already portable.

## What this replaces (eventually)

`v_wip_latest` is the portfolio rollup that's rebuilt in Excel every month. Once the Phase-2
connectors fill the granular tables, over/under-billing and per-cost-code budget-vs-actual
become queries against the spine instead of computed spreadsheet columns — one definition,
computed once, correct everywhere.

See `../docs/ARCHITECTURE.md` (the "Ledger" section) for the data-flow diagram, and
`STATUS.md` for the current state and open items.
