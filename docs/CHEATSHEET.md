# Cheat Sheet / FAQ - running the tools

Quick reference for running the automation suite from the terminal.

## First: you may not need the terminal anymore

Most **data syncs are now one click in the ledger Console**. Open the ledger, go to the
**Console** tab, and hit **Run** on a pipeline (AR / AP / Costs / CRM / WIP) or **Full
refresh**. **Reload** on My view refreshes just the ledger (read-only, fast). The tools
below are the ones you still run from a terminal (the statement reconciler, P&L export,
and the raw scripts behind the Console).

## Setup (once per terminal window)

Everything runs from the repo folder - cd there first:

```bash
cd ~/Documents/Claude/Projects/"Automate Concrete Business"
```

- QBO tools prompt **Touch ID** (one per run). Notion-only tools do not.
- Any script's own flags: `python3 <tool>/<script>.py --help`.
- `--dry-run` previews without writing (most writers). `--yes` skips confirm prompts.

---

## Vendor statement reconciler

Reconciles a vendor statement against QBO open bills and writes an Excel report
(Matched / vendor-tax-violation / clerk-mismatch / missing-in-QBO / missing-on-statement).
Reports land in OneDrive `Automations-/statement reconciles/`.

**Scan the whole Synology inbox** (`/Volumes/Accounting/Automations/Vendor Statements` -
needs the **Accounting** share mounted). It reconciles every file in the inbox, then moves
each into the `DONE` subfolder:

```bash
python3 statement-reconciler/statement_reconciler.py --inbox
```

Unattended sweep (auto-processes only vendors already confirmed/cached; a first-time
vendor is left in the inbox to run once by hand so it can be confirmed):

```bash
python3 statement-reconciler/statement_reconciler.py --inbox --yes
```

**One statement file** (PDF, image, or xlsx):

```bash
python3 statement-reconciler/statement_reconciler.py "/path/to/statement.pdf"
```

Force the vendor if auto-detect is unsure (use the exact QBO display name):

```bash
python3 statement-reconciler/statement_reconciler.py "/path/to/statement.pdf" --vendor "Exact QBO Display Name"
```

Preview without writing / point at a different inbox:

```bash
python3 statement-reconciler/statement_reconciler.py --inbox --dry-run
python3 statement-reconciler/statement_reconciler.py --inbox --inbox-root "/Volumes/Accounting/Automations/Vendor Statements"
```

---

## The ledger app (the super-database)

Open the dashboard:

```bash
python3 ledger/dashboard.py
```

Or double-click **Open Project Ledger.command** (in `~/Documents/CompanyHealth/`) or the
Project Ledger Dock app. Inside: the **Console** tab runs every pipeline; **Resync**
(My view) reloads the ledger data; the **P&L** tab shows the live company + per-job P&L.

### How a sync flows (fresh sync, in order)

The dashboard reads a local SQLite DB (`~/Library/Application Support/Proficient/ledger.sqlite3`).
**Producers** pull from QBO/Notion out to files/Notion; **loaders** read those into the DB;
the app reads the DB. Order matters - WIP builds the `project` table everything else joins to.
The DB is created automatically by the first loader (no setup step).

1. **WIP → ledger** - `load_wip_master.py` builds the project spine + snapshots. **Runs first.**
2. **Costs → ledger** - `load_costs.py --active` pulls QBO job costs by cost code (Touch ID).
3. **AP** - `excel_bill_sync.py` (QBO → `Bill Tracker.xlsx`, Touch ID) → `load_bill_tracker.py` (→ the **Bills** tab).
4. **AR** - `run_invoice_sync.py` (QBO → Notion + Teams, Touch ID) → `load_invoices.py` (→ the **Draws** tab).
5. **CRM** - `load_customers.py` pulls the Notion Customer List (→ the **Sales** tab).

**Full fresh sync** (producers + loaders, in order; stops on any failure; expect several Touch ID prompts):

```bash
python3 ledger/load_wip_master.py && python3 ledger/load_costs.py --active && python3 bill-tracker/excel_bill_sync.py && python3 ledger/load_bill_tracker.py && python3 invoice-sync/run_invoice_sync.py && python3 ledger/load_invoices.py && python3 ledger/load_customers.py
```

**Loaders-only reload** (fast, read-only - use when `Bill Tracker.xlsx` / Notion are already current; skips the two producers):

```bash
python3 ledger/load_wip_master.py && python3 ledger/load_costs.py --active --since 2026-05-16 && python3 ledger/load_bill_tracker.py && python3 ledger/load_invoices.py && python3 ledger/load_customers.py
```

The `--since` date makes the cost pull incremental (any date ~90 days back); drop it to pull the
full history (slower, but also reaps txns deleted in QBO). In the app, **Console → Full refresh**
is the full sync above and **Resync** (My view) is the loaders-only reload. `sync-ap`, `sync-ar`,
and `sync-all` (AP→AR) are your shell aliases for just the producer steps.

---

## AR - invoice sync (QBO to Notion + Teams)

```bash
python3 invoice-sync/run_invoice_sync.py
```

Visual runner (phases + a live progress bar):

```bash
./invoice-sync/run_invoice_sync.sh
```

---

## AP - bill tracker (QBO to `Bill Tracker.xlsx` on OneDrive)

```bash
python3 bill-tracker/excel_bill_sync.py
```

Visual runner:

```bash
./bill-tracker/run_tracker.sh
```

Per-job coding drill (interactive):

```bash
python3 bill-tracker/job_coding_audit.py
```

Note: `qbo_bill_tracker.py` is a legacy tracker - use `excel_bill_sync.py`, which is what the ledger reads.

---

## WIP master (draft Test tabs for PM review)

```bash
python3 wip/cp_wip_reader.py
python3 wip/rp_wip_reader.py
python3 wip/master_wip_test.py
```

Inspect before writing (no WIP write): `python3 wip/master_wip_test.py --audit`.
Close lists (these always exclude MFD): `python3 wip/qbo_close_list.py`.

---

## Project P&L (to OneDrive `PROJECT P&Ls`)

One project:

```bash
python3 project-pnl/project_pnl_export.py CP800
```

Every Active project of a division (batch):

```bash
python3 project-pnl/project_pnl_export.py active cp
python3 project-pnl/project_pnl_export.py active rp mfd
```

---

## Other tools

```bash
python3 debt-schedule/loan_sync.py
python3 health-dashboard/qbo_health.py
python3 qbo-export/qbo_export.py
python3 synology/synology_tree.py --root "/Volumes/Common" --exclude /Volumes/Proinfo/Items/
```

(The Synology audit must ALWAYS pass `--exclude /Volumes/Proinfo/Items/`.)

---

Each tool has its own `README.md` with the full detail; this page is just the quick "how do I run it."
