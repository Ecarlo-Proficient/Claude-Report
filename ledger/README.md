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
| `sync_all.sh` | **THE sync** behind the `sync-all` alias: mirror refresh (Sundays + reconcile) -> AP -> AR -> ledger reload. |
| `refresh_mirror.py` | **The raw QBO mirror** (`shared/qbo_mirror`): `--seed` once, then the change feed; `--reconcile`, `--status`. Step 0 of sync-all once tools read it. |
| `load_wip_master.py` | Reads the FINAL WIP master (the Test tabs) → fills `project` + `wip_snapshot`. |
| `load_bill_tracker.py` | Reads `Bill Tracker.xlsx` (Bills + Inventory) → fills `ap_bill_line` (AP + liens). |
| `load_costs.py` | QBO pull → `cost_line` by cost code (incl. subs), via `shared/qbo_costs`. |
| `load_customers.py` | Notion "Customer List" → `customer` + `sales_touch` (CRM leads/clients + outreach touch log, read-only). |
| `load_health.py` | QBO pull → `health_snapshot`: bank cash, retainage GL, P&L blocks, 13-wk cash flow, recurring register - the Health tab's QBO-only layer. |
| `dashboard.py` | Local web dashboard over the ledger — the browser UI (read-only). |
| `registry_view.py` | Parses the vault's systems & process registry (`02_processes/*.md`) for the Systems tab, and links each row to its one-page guide in the vault's `assets/processes/` (`<ID>_<slug>.pdf`/`.html`, served by `/api/process-guide`). Read-only, no DB. |

The cost engine itself lives in **`shared/qbo_costs.py`** (`cost_leaf` + `iter_cost_lines`) — the
SAME resolver project-pnl uses, so the ledger and the P&L can never drift.
| `static/` | The dashboard front-end (`index.html`, `style.css`, `app.js`) — no build step. |

## The raw QBO mirror (2026-09-17)

**The change log (2026-09-23).** Every refresh (never the seed) diffs each record QBO hands back against the copy the
mirror holds and writes one `mirror_change` row per record that moved - `created` · `edited` · `deleted` · `restored` -
stamped with QBO's own time (`MetaData.LastUpdatedTime`, or the change feed's deletion time), the before / after
summary (total, balance, money applied, line count, party), plain-word flags from `change_flags()` and the BEFORE
record encrypted like the rest (`change_before(con, id)` = the repair material). `changes(con, since=, entity=)`
reads it; `backfill_changes()` runs once for the deletions flagged before the log existed;
`refresh_mirror.py --changes [N]` prints the last N days. Serves the ledger's QBO Audit page.

`qbo_mirror.sqlite3` beside the ledger: every QBO entity (21 - all transaction types plus the
name lists), one row per record with the full JSON as QBO returned it, deletes flagged with a date
and never dropped. `python3 ledger/refresh_mirror.py --seed` once (~20 min, 296k records);
`python3 ledger/refresh_mirror.py` after that reads QBO's change feed since the last stamp
(seconds). `--reconcile` checks a count per entity; `--status` shows what the file holds.
Every tool reads it already: `shared.qbo_api.query_all` is answered from the mirror (QBO WHERE
evaluated in Python; `ACB_QBO_LIVE=1` forces QBO). `one-offs/mirror_parity.py` proves the two
agree shape by shape. What is left per tool is in `STATUS.md`. QBO stays the source of truth; nothing writes to QBO from here.

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

### UI rules - the register (owner 2026-09-22: "if i tell you something about the project ledger, i expect it to save across the whole thing")

The ledger is ONE app: a behaviour the owner approved on one page is the standard on every page. Reuse the
component below; never rebuild one from scratch, and add to this list when a new rule is settled.

- **Date filter = `dateFilter()`** (Month multi-select | Date from-to, Reset). The month list STAYS OPEN while
  ticking; unticking the last month means all; newest month first. On a page that re-renders its body, pass
  the previous state (`prev`) and reopen the menu after the render (`_dateMenuWasOpen` in `renderVendorPage`).
- **Every multi-select stays open while selecting** (the filter-bar menus, the header funnels, the months).
  Close on an outside click or Esc, never on a tick.
- **Column filters = `hfDecorate()`**: an Excel-style funnel in each header - distinct values with counts over
  what the other columns leave, a value search, Select all / None / Clear. New tables get these, not bands.
- **Bill tables are flat lists like the Excel** (Group by = None); the funnels and the search do the narrowing.
- **Search on every view of a page**: a broad box (⌘F / Ctrl+F lands in it, Esc clears, every word must match
  somewhere on the row, memo included) and a standalone Project # box that suggests the page's projects
  (ArrowDown / ArrowUp, Enter picks).
- **Payments are groups**: collapsed = ref, date, type, client, amount, stub; the bills paid are the expansion;
  a checkbox per row and a select-all for bulk actions; a voided payment says VOIDED.
- Groups elsewhere follow the central `GRP_KINDS` mechanism.
- **Four views** (owner 2026-09-23): Projects · Vendors (Bill Tracker · Vendor Center) · Customers (Invoice Tracker ·
  Customer Center · Payments received · Sales pipeline) · Company (Money · QBO Audit) + the gear. A new company-wide
  list is a sub-tab of the view it belongs to, never a fifth view.
- **Stats are one line each** (label · value · note · source), never boxed tiles - "no big waste spacer blocks".
- **A click never scrolls the page** (a draw, a group, a filter): re-render in place, restore `scrollY`.

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
- **QBO Audit** (Company) — what changed in QuickBooks: the mirror's change log (`/api/qboaudit?days=`), every
  record deleted / edited / restored / added since the previous refresh with QuickBooks' own time and plain-word
  flags (deleted paid bill, payment unapplied, reopened, voided, class dropped…), tiles + chips + search, **Refresh
  from QuickBooks** (the `mirror` pipeline); the Bill Tracker audits sit underneath. QuickBooks' API never says who -
  the audit log inside QuickBooks does. Born 2026-09-18 (a connected app's deletions were logged as the owner).
- **Uncleared checks** (Company) - every check QuickBooks still shows as uncleared (not matched = not deposited),
  from QBO's TransactionList `cleared=Uncleared` filter via `ledger/load_uncleared_checks.py`; each bank account's
  matched-through date; accounts never matched to a bank feed (Joint Checks) kept apart. `/api/uncleared`.
- **Checks QBO changed** (Company) - checks QuickBooks rewrote on its own after they were paid (edit or delete a
  paid bill and QBO takes the check off it without a warning): money floating, the paid bill or its re-entered copy
  reading as owed again, the freed loan credit, floating money put on a later bill. Every year on file, subs first,
  one bill per check, the fix per check. `ledger/check_drift.py` (`/api/checkdrift`; `--check <#>` on the command line).
  Chips in ONE row by what happened: **To fix · Unapplied · Moved to a newer bill · Credit · Resolved · History**, plus
  **All vendors / Subs / Suppliers**; Urgent / Low is the row's tag. **Re-apply in QuickBooks** on a check's card puts it
  back on the bills it paid - **the one QBO write in the ledger**: a dry run first (live QBO, the bills listed, matched
  against the card), then "Are you sure?", then `ledger/reapply_check.py` writes only that exact plan (same check
  version, bill count and amount) and refreshes the mirror. What it paid = the change log's before copy, or for a check
  stripped before the log (09/23/2026) a saved ledger copy (`ledger*.sqlite3*` `bill_payment_line`) adding up to the check
  to the cent. Each bill re-read live (exists, same vendor, open balance covers it), never over the check total, the live
  check backed up to `~/Library/Logs/Proficient/reapply-check/` first. `/api/checkdrift/reapply` (GET = dry run, POST =
  write); CLI `python ledger/reapply_check.py <check#> [--list] [--commit]`. **Mark resolved** = a local mark (settled
  outside the ledger, e.g. the vendor applied it to another bill).
  **History** keeps every check QuickBooks left with no bills, fixed or not: when, the trigger (a paid bill deleted /
  check saved with no bill edited / before the change log), bills before, saves, status with the re-applied date, and
  any bill paid a second time. **QBO report (PDF)** rebuilds the support report from it on every click, into
  CompanyHealth. `ledger/strip_history.py` (`/api/checkstrips`, `/api/checkstrips/pdf`; `--pdf` on the command line).
  Strips from before the change log (09/23/2026) and the report's narrative live in
  `CompanyHealth/check_strip_case.json` (business data, never in the repo).
- **Copy & export** — click any number to copy it; **Export CSV** downloads the current view.
- **Customize** (⚙) — theme (auto/light/dark), accent color, font, text size, density, width
  (**boxed by default**), which widgets show, and which table columns show. Saved per person in the
  browser; **Set as default** snapshots the current view as the baseline that Reset (and a fresh
  browser) restores to.

- **Systems** (the systems & process registry, read-only from the vault) — every process the
  business runs, one row each, read **live** from `AI Brain_Vault/02_processes/*.md`. Three axes are
  kept separate because the registry keeps them separate: a **health dot** (running / fragile /
  broken / nothing to fail yet), a **state pill** (how sure the description is — confirmed ·
  inferred · proposed), and a **life tag** shown only when a row is not live (idea · agreed ·
  building), which is how "agreed but never built" stops being invisible. Filter by domain, owner,
  health, state or life; **Reload** re-reads the files. Owners show as role handles exactly as
  stored — the roster is never read and no name enters the UI.

  It is a window, not a copy: `registry_view.py` re-parses the markdown on every request, so the
  update loop is *edit the vault file, hit Reload*. Nothing is stored in `ledger.sqlite3` and
  nothing is written back — the vault stays the source of truth. Point it elsewhere with
  `ACB_VAULT_DIR` in `machine.env`; a machine with no vault just shows a one-line note on that tab.
  **This replaced the daily 06:38 markdown digest** (disabled 2026-08-19; `digest-log.md` kept as
  history).

The dashboard is a view. It never writes the database or the Excel sheet.

### WIP review (weekly, with the division PM) - 2026-09-15/16

Gear -> **WIP review · with the PM**, or the **WIP review →** button on a division band of the Projects
page. One page per division - **RP · CP · MFD** - the same system for all three. Every line on the WIP
is a card: the numbers, **where each one was grabbed from** (a source pill + the file, with a Finder
"Show file in folder" button), and **Our numbers** - the owner's / PM's own space (prepopulated,
a check or X per number, a note, Confirmed / Needs a fix, an explicit Save that reads the answer back).
Every answer is stamped **Me** (the owner alone) or **PM** (the division PM is here, deciding together)
plus the time, in `rp_review_mark` (+ `rp_review_mark_log`); a line with no answer in 7 days is due again.

- **RP** - `ledger/rp_review.py` builds `rp-review/rp_review.json`: every RP line on the master's RP tab,
  the crew schedule (every day the job was on it), the proposal PDF / takeoff bid sheet / cost sheet the
  number came from as a **picture of the real file**, JobTread's approved price / cost, and the finished
  lines from the RP file's Removed log. **Finalize to WIP Report** turns the answers into the package the
  guarded WIP writer applies (nothing is written from here).
- **CP / MFD** - no schedule (the awarded-projects folders are kept current): `rp_review.division_cards()`
  adapts the WIP tool's `--emit-review` JSON (`wip-review/cp.json`, `master.json`) - the contract and the
  **approved change orders from the draw G702**, the **ETC from the takeoff**, costs and billed from
  QuickBooks, and the pending update's was → now per field with the document the new number came from.
  MFD's contract / ETC are typed on the master by design and say so. Approved changes are written from the
  **WIP Review** tab as before.

```
python3 ledger/rp_review.py                 # full RP build (JobTread = one Touch ID), a few minutes
python3 ledger/rp_review.py --no-jobtread   # without JobTread
python3 ledger/rp_review.py --answers       # what the owner / PM answered, every division
```

Endpoints: `/api/review?div=RP|CP|MFD`, `/api/review/mark`, `/api/review/refresh` (`{div}`); the old
`/api/rp/*` routes stay as aliases.

### Vendor page and client page (2026-09-16)

**2026-09-22 - flat like the Excel, Excel-style filters.** The Bill Tracker tab and the vendor page's Bills view are
flat lists by default (Group by = None; the vendor page has no bands at all). Every column header carries a funnel
that opens the column's distinct values with counts, a value search, Select all / None / Clear - one mechanism
(`hfDecorate` in app.js), the values reflecting what the other columns leave, like an Excel table. A search box
narrows either list by any word on the row; the vendor page's box works on Payments too.
Every bill row carries its QuickBooks **memo** (from the mirror, `_bill_memos`) as a column and in the search; the
Date funnel filters by month; the vendor page adds a standalone Project # box beside the broad search, and ⌘F lands
in the broad search (vendor page when open, else the Bill Tracker).
The vendor page also carries the Bill Tracker's Date control (Month multi-select or a from-to range) on both views.
Its Project # box suggests the vendor's projects as you type - ArrowDown, Enter to pick.

A vendor's page IS the Bill Tracker filtered to that vendor - the same rows and columns, every bill with
its project, client, the **invoice (draw) it is matched to and whether the GC has paid it**, our pay
status, lien and approval; All / Unpaid / Paid, an invoice filter (GC paid · GC owes · no invoice yet),
project bands, and **Open in Bill Tracker** to carry the vendor into the tracker's own filters. A client's
page IS the Open invoices grid filtered to that client (Open · All incl. paid, project bands, bucket
totals) with **Open in Invoices**. On the trackers, the vendor name opens the vendor page and the
client name opens the client page.

### Bill payment stubs (2026-09-22)

`ledger/bill_payment_stub.py` prints a check / bill-payment stub the way QuickBooks lays out its
"Bills and Applied Payments" report, with the PAYMENT on top and the bills it paid below - one
payment per Letter page. `--vendor COWTOWN --date 2026-09-21` (or `--from/--to`, or `--payment <id>`)
reads the mirror (BillPayment -> the bills it linked, the vendor's print name and address) and writes
`Accounting/Accounts Payable/Bill Payment Stubs/<QBO vendor name>/<Vendor> - Bill Payment Stub <check # or card reference> <mm-dd-yyyy>.pdf`
(`shared/paths.bill_payment_stubs_dir()`; **one payment = one file**, two payments on a day are two files; the share must be mounted) through Chrome headless. Columns are a
registry: the default six are QBO's (date · type · number · memo · amount · open balance);
`--add project,bill_total,due_date` / `--drop type` / `--columns ...` change the set; `--list-columns`
shows it. `amount` is what that payment applied to the bill, so the column ties to the payment; a bill
covered only in part says so under its memo. The company name on the stub is `ACB_COMPANY_NAME` in
`machine.env`. `--selftest` proves it offline.

**In the ledger:** vendor page > Payments view, the **Stub** column - `Print stub` per payment (POST
`/api/bill-payment/stub`), a **Stub columns** picker (the registry as checkboxes, QBO's six by default,
remembered per browser), and the **print history** under each payment. Every print is a row in
`bill_payment_stub` with the stub as printed (snapshot), the columns, the QBO SyncToken at print time and
its own archived PDF copy beside the ledger DB - `GET /api/bill-payment/stub/file?id=` serves that copy, so
a print opens as it went out even after a re-print. The status pill is judged live against the mirror:
current · changed in QBO · voided in QBO · deleted in QBO. Payments the ledger's list no longer carries
(deleted, voided, outside the year window) but which have prints are listed under "Printed stubs for
payments no longer in this list", each re-printable from the mirror's copy. `--history` shows the same
on the command line.
The Payments view itself is grouped (owner 2026-09-22): a payment is the row - ref, date, type, client,
amount, stub - and the bills it paid open under it (the central group mechanism, `tr.vp-pay` in
`GRP_KINDS`). A checkbox per payment and a select-all feed **Print N stubs** for a multi-print. A payment
voided in QuickBooks (zeroed, memo "Voided…") shows a red VOIDED pill with its memo and cannot be printed.

### The project page's funding section (2026-09-16)

Every invoice on the job is income. Invoices that name the same draw month (MFD192 bills its base,
HUDSONWOOD and OFFSITE contracts every month) or, failing a named month, that carry the same date are
**one draw** - income is their sum, the bills are what the Bill Tracker matched to any of them. **An RP
job has no draws: each invoice is a scope** and its costs are the bills (tracker and sub) dated after
the previous invoice up to its own date; bills after the last invoice sit in "Not yet invoiced". The
Coverage table and the draw / scope boxes stay on top; a click opens that one underneath as the
**equation** the owner reads (invoices = Income − Materials − Labor = Gross profit − Overhead = Net)
over the bills - all bills by default, plain section rows with Materials total / Labor total / Total at
the bottom, sortable by vendor name, total amount, cost code or date. Above it, "How it's doing" is one
table: **Projected (WIP master) next to Actual (QuickBooks)**, row by row. **Add a note** keeps the
owner's notes on the job (`project_note`). **Pay bills** (off to the right) reveals the pay-run
controls; ticks are a draft until **Save** (Discard drops them), and the page will not let you leave
with unsaved ticks.

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
