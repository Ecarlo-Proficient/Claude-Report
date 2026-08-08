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
| `dashboard.py` | Local web dashboard over the ledger — the browser UI (read-only). |
| `static/` | The dashboard front-end (`index.html`, `style.css`, `app.js`) — no build step. |
| `requirements.txt` | `openpyxl` (SQLite + the web server are stdlib — nothing else to install). |

## The schema (6 tables + 1 view)

```
project          the aggregate root — one row per job. project_no is THE join key.
cost_code        job-type prefix + number, a first-class dimension (not a QBO item-name string)
budget_line      the plan: ETC by cost code
cost_line        actual spend from QBO bills (append-only, idempotent by bill+line id)
billing_event    AR invoices / draws (append-only, idempotent by invoice id)
wip_snapshot     the COMPUTED WIP position — one row per (project, report_date)
v_wip_latest     view: each project joined to its most-recent snapshot
```

**What fills what:**
- `project`, `wip_snapshot` ← `load_wip_master.py` (**today** — from the WIP master sheet)
- `cost_code`, `budget_line`, `cost_line`, `billing_event` ← the QBO connectors (**Phase 2**,
  one connector at a time — not built yet)

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
