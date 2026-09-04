# Bill Tracker

Pulls AP bills from QBO, matches each line to the GC invoice that authorizes its
payment, and writes `Bill Tracker.xlsx` (OneDrive `Automations-/`). One tool, one
workbook: the display sheets are the daily AP view; the **QBO Audit** sheet is the
consolidated coding-error catcher.

## Pipeline

`run_tracker.sh` → `sync_view.py` → **`excel_bill_sync.py`** → `Bill Tracker.xlsx`

- `qbo_bill_tracker.py` — the auth/paging/parsing **library** (QBO fetch, project #
  parsing, sub detection, invoice matching). Not run directly anymore.
- `bill_rows.py` — pure QBO → row-dict assembly + collapse helpers.
- `general_list.py` — read-only General List loader (RP draw semantics).
- `job_coding_audit.py` — the interactive per-job drill (see **Audits** below).

**Runs are manual** (`sync-ap`, or `sync-all` for AP→AR). No scheduler — the QBO
Keychain entry needs a Touch ID per run, so someone has to be at the desk.

## Manual run

```bash
bill-tracker/run_tracker.sh
```

Or directly (from the repo root):

```bash
python3 bill-tracker/excel_bill_sync.py
python3 bill-tracker/excel_bill_sync.py --dry-run      # build rows, write nothing
python3 bill-tracker/excel_bill_sync.py --limit 200    # cap rows (smoke test)
```

## The full pull (the user 2026-08-06)

The tracker pulls **every** bill — open (any date) + paid (since `PAID_CUTOFF_DATE`),
**including sub bills**. Subs are then split off:

- **Bills / Inventory / Liens** (the display sheets) exclude subs — unchanged.
- **The QBO Audit sheet** sees the full population incl. subs, so sub-bill findings
  surface there and nowhere else.

Cost codes (the QBO **Item** name — `SL1`, `PV6`, `FW2` …) are captured at ingest for
the audit only; they are **never** shown as a column on the display sheets.

## Output sheets

| Sheet | Grain | Notes |
|---|---|---|
| `Bills` | bill | one row per bill; multi-project bills show `(multiple)`, drill to Inventory |
| `Inventory` | line | per-line drill-down for multi-project bills |
| `Liens` | — | live Excel FILTER over the Bills table (Lien set) |
| `Audit - Coding` / `Audit - PO` / `Audit - Bills` | mixed | THREE themed Excel Tables, each with an `Issue` filter — see below |
| `Audit - History` | miscode | persistent cost-code miscode log — how often + what got fixed (see below) |

`Lien` and `Notes` on the Bills sheet are user-editable and **preserved** across runs
(keyed by `Bill.Id-Line.Id`).

## The audit — THREE themed sheets (the user 2026-08-25, de-bloat from 9 tabs)

The audit checks are grouped into **three filterable Excel Tables**, each with an `Issue`
column (filter to one check) plus a `Detail` column and an `Open` QBO deep-link. Every row's
finding logic is unchanged — only the rendering is consolidated.

**Acknowledge & keep** — `Audit - Coding` has an editable **`Status`** column. Type a mark
(e.g. `KEEP - prepaid, allocate at closeout`) on a one-off you're deliberately keeping: it
**persists** across runs (keyed by bill #) and **mirrors into the Bills `Notes`** so both agree.
Filter `Status`: blank = still needs review, marked = acknowledged / on the record. This differs
from `audit_exclusions.json`, which *hides* an item — the mark keeps it visible so nothing is lost.

| Sheet | `Issue` values it merges |
|---|---|
| `Audit - Coding` | **Data Entry** (empty/mismatched Class, line-desc project mismatch) · **Missing Project** (uncoded job cost, no project #) · **FW Misplaced** (FW code on a CP/MFD/base-`RP####` slab) · **Sub No Project** (sub cost-code line, no project #) · **Cost Code** (wrong cost-code FAMILY for the vendor's type; `Detail` carries the type + PO origin) |
| `Audit - PO` | **Unused PO** (`Open, no bill` · `Stale >60d` · `On tracker, not in QBO`) · **Missing PO** (a real COGS bill — not a sub, not expense-only — with NO PO, last 90 days). PO rows link to the QBO PO, bill rows to the bill |
| `Audit - Bills` | **Not Approved** (stale NOT APPROVED bills) · **Duplicate** (same bill # within a vendor tree) |

**Cost-code detail:** the cost-code NUMBER is the family (1 concrete · 2/3/4 material · 5/51/52
equip · 6 labor). Each vendor's TYPE is auto-captured from its `*1`-vs-`*2/3/4` split —
**concrete** (→ all `*1`; a `*4` aggregate line - pea gravel / sand / base - is fine WHEN the
memo says so, memo-gated not a blanket pass), **material** (e.g. RCI → `*2/*3/*4`, never `*1`/`*5`/`*6`), **both**
(e.g. Preferred Materials → a yardage/ready-mix MEMO line must be `*1`), **hauler** (trucking/haul
→ material plus haul-off `*5` OK, only `*1`/`*6` flag; override-only). Credit-card / finance / bank
/ late fees post to an expense account, not a cost code — never flagged. Each flag's `Detail`
cross-references the bill's PO: **upstream** (PO also carried the wrong code), **bill deviated**, or
**no PO**. Logic in `shared/cost_code_audit.py`; type overrides via
`<companyhealth>/concrete_suppliers.json` (`{concrete/material/both/hauler/exclude}`).

`Audit - PO` reads a second source: the office **PO tracker** workbook (`Orders` tab,
READ-ONLY) via `po_tracker.py`. Path = `ACB_PO_TRACKER_XLSX` (default OneDrive
`Purchase Orders/Copy 05 dic.xlsx` — the CURRENT file; the older `1.0purchase-order-tracker.xlsx`
is ~15 months stale). If the tracker is unreadable the sheet degrades to a QBO-only view. The
manual tracker lags QBO by weeks/months, so its last-data-date is printed on the sheet — never
treat a blank tracker side as truth.

The Duplicate / FW / Sub-No-Project checks fold in what the standalone `duplicate_bill_audit.py`,
`item_no_project_audit.py`, and `sub_bill_audit.py` did — retired 2026-08-06 (one tool, one
workbook). History note: the audit covers the tracker's population (open +
paid-since-`PAID_CUTOFF_DATE`); the retired scripts could scan all-time, so pre-cutoff
**paid** bills are out of the audit's window.

## `Audit - History` — the cost-code miscode log (the user 2026-09-01)

"I need this logged in the system — how often the clerk is making the mistakes, and what got
fixed after refreshing." The `Cost Code` findings are folded into a **persistent log** that
survives across runs, rendered as a fourth filterable sheet.

- **State** = `<companyhealth>/cost_code_history.json` — OUTSIDE the repo, next to the other
  audit configs, never committed. Logic in **`cost_code_history.py`** (pure state; no QBO).
- Each **real** run (dry-run bails before the audit, so it never mutates the log): every current
  miscode opens/updates an entry (`First Seen` / `Last Seen` / `Times`); any previously-**OPEN**
  entry gone from this run flips to **FIXED** (with the date) — that is "what got fixed"; a FIXED
  entry that reappears re-opens.
- Key = `bill_id + cost_code`, so recoding `SL51 → SL1` next run reads as one FIXED. Recoding to
  *another* wrong code reads as one fixed + one new (a fresh mistake — which is the point).
- The sheet shows every **OPEN** entry plus **FIXED** entries from the last 60 days (older fixes
  stay in the JSON, off-sheet, so it stays lean). Filter `Status` = OPEN / FIXED. The caption
  carries the recap: open / new-this-run / fixed-this-run and a rolling **new/run** rate — the
  bill clerk's coding-error frequency (one clerk does data entry, so the aggregate is the rate).

## Audits — the remaining standalone drill

`job_coding_audit.py` stays as an on-demand, per-job investigation (via the `audit-job`
shell alias) — a different workflow from the always-on `Audit - …` table sheets. It writes
a plain xlsx and never edits QBO.

| Script | Flags | Catches |
|---|---|---|
| `job_coding_audit.py` | `--job <code>` | Lines that mention a job but are coded to the wrong/missing project or a non-parent class |

```bash
python3 bill-tracker/job_coding_audit.py --job MFD281
python3 bill-tracker/job_coding_audit.py --job CP142 --checks desc --dry-run
```

## How matching works

For each non-sub bill line:

1. Read project # from the line's CustomerRef name (`MFD###`, `CP###`, `RP####`).
2. Division-specific rule:
   - **RP** → amount-aware: the largest same-project invoice on/after the bill date that
     covers the bill amount.
   - **MFD / CP** → invoice whose Draw Period (parsed from `PrivateNote`) contains the bill date.
3. No match → `Awaiting Invoice`; matched + invoice paid → `OK TO PAY`; matched + invoice
   open → `Awaiting Payment`; no project # → `No project #`; bill paid → `Bill paid`.

## Exclusions

- **Sub bills** (`PrivateNote` contains the word "sub") — off the display sheets, but
  **included in the audit**.
- **Retainage invoices** dropped from the matching pool (`PrivateNote` "retainage not
  billed", or every line description contains "retainage"). Regular draws with a
  retainage withholding line stay in.

## Touch ID

The QBO Keychain entry (`automation-qbo`) has a biometric ACL, so every run prompts
Touch ID once. Expected — the tracker is run by hand from your desk.

## Troubleshooting

- **`✗ no credentials`** — Keychain entry wiped/never created. Run `shared/setup_qbo.py`.
- **`token refresh status=400`** — refresh token expired or rotated by another process.
  Re-run `shared/setup_qbo.py --rotate QBO_REFRESH_TOKEN`.
- **validation FAILED (Excel-strict preflight)** — a known-bad openpyxl pattern survived;
  the workbook is saved but should not be trusted until the reported issues are fixed.
- **file open in Excel** — `safe_save` writes to a temp file and swaps; close the workbook
  and re-run.
