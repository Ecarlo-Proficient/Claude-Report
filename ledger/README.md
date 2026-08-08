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
v_wip_latest     view: each project joined to its most-recent snapshot
v_ap_by_project  view: open AP + bill counts per project
v_cost_by_project view: loaded QBO cost per project (reconcile vs WIP)
v_cost_by_code   view: per-project cost-code drill (budget-vs-actual base)
```

**What fills what:**
- `project`, `wip_snapshot` ← `load_wip_master.py` (**today** — from the WIP master sheet)
- `ap_bill_line` ← `load_bill_tracker.py` (**today** — from Bill Tracker.xlsx)
- `cost_line` + `cost_code` ← `load_costs.py` (**today** — one QBO pull, incl. subs)
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

## The dashboard (browser UI)

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && python3 ledger/dashboard.py
```

Opens `http://127.0.0.1:8787` in your browser (add `--no-open` to skip that, `--port N` to
change the port). It reads the ledger **read-only** and binds to `127.0.0.1` only (not exposed
on the network). What it shows:

- **Portfolio KPIs** — total contract, costs, billed, left-to-bill, net over/(under), active jobs.
- **By division** — the CP / RP / MFD rollup.
- **Projects** — searchable, filterable (division / status / category / active-only), sortable;
  click any row for the full job detail (Contract / Budget / Costs / Earned / Billing / Notes).
- **Copy & export** — click any number to copy it; **Export CSV** downloads the current view.
- **Customize** (⚙) — theme (auto/light/dark), accent color, font, text size, density, width,
  which widgets show, and which table columns show. Saved per person in the browser.

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
