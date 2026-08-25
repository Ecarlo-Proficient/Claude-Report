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
| `Audit - …` (×8) | mixed | one proper Excel Table per audit section — see below |

`Lien` and `Notes` on the Bills sheet are user-editable and **preserved** across runs
(keyed by `Bill.Id-Line.Id`).

## The audit — one Excel Table per section (the user 2026-08-06)

Each section is its own **`Audit - …` sheet** wrapped in a proper Excel Table, so every
column filters and sorts natively (banners can't live inside a table — that's why they're
separate sheets). Every data row deep-links to the bill in QBO via the `Open` column. An
empty section still renders a valid one-row table (`✓ none found`).

| Sheet | Catches |
|---|---|
| `Audit - Not Approved` | Stale NOT APPROVED bills aged past `NOT_APPROVED_BUFFER_DAYS` (Aging + Days Old columns) |
| `Audit - Data Entry` | Empty/mismatched Class, line-desc project mismatch |
| `Audit - Missing Project` | High-confidence uncoded job costs with no project # (non-sub) |
| `Audit - Duplicates` | Same bill # within a vendor tree — double-entry / double-pay risk (all bills) |
| `Audit - FW Misplaced` | **FW flatwork code on any CP job, MFD job, or base `RP####` slab** — legit only on `-FTW`; division/slab from the project #, never Class. Has Cost Code + Sub? columns |
| `Audit - Sub No Project` | A sub **cost-code (item) line** with no project # — account-based category lines (reimbursements, fees, overhead) are excluded |
| `Audit - Unused PO` | **QBO POs × the office PO tracker, one story.** Flags: `Open, no bill` (QBO PO Open, nothing billed) · `Stale >60d` (that, aged) · `On tracker, not in QBO` (recent UNBILLED tracker PO never issued). QBO + tracker columns side by side; `Open` deep-links the QBO PO. Freshness of the manual tracker is stamped by the header |
| `Audit - Cost Code` | **Vendors coding to the wrong cost-code FAMILY** (the number: 1 concrete · 2/3/4 material · 5/51/52 equip · 6 labor). Each vendor's TYPE is auto-captured from its `*1`-vs-`*2/3/4` split: **concrete** (→ all `*1`), **material** (e.g. RCI → `*2/*3/*4`, never `*1`/`*5`/`*6`), **both** (e.g. Preferred Materials → a yardage/ready-mix MEMO line must be `*1`). Flags every line that breaks its type's rule. **Each flag is cross-referenced to the bill's PO** (`PO #` / `PO Cost Code` / `Origin`): whether the PO ALSO carried the wrong code (**upstream** super/PM), the bill **deviated** from a correct PO, or there was **no PO** (clerk-coded). PO codes come from QBO; the PO tracker recovers the PO# when QBO left the bill unlinked. Logic in `shared/cost_code_audit.py`; override types via `<companyhealth>/concrete_suppliers.json` |

`Audit - Unused PO` reads a second source: the office **PO tracker** workbook (`Orders` tab,
READ-ONLY) via `po_tracker.py`. Path = `ACB_PO_TRACKER_XLSX` (default OneDrive
`Purchase Orders/Copy 05 dic.xlsx` — the CURRENT file; the older `1.0purchase-order-tracker.xlsx`
is ~15 months stale). If the tracker is unreadable the sheet degrades to a QBO-only view. The
manual tracker lags QBO by weeks/months, so its last-data-date is printed on the sheet — never
treat a blank tracker side as truth.

`Audit - Duplicates`, `Audit - FW Misplaced`, and `Audit - Sub No Project` fold in what the
standalone `duplicate_bill_audit.py`, `item_no_project_audit.py`, and `sub_bill_audit.py` did — retired 2026-08-06
(one tool, one workbook). History note: the audit covers the tracker's population (open +
paid-since-`PAID_CUTOFF_DATE`); the retired scripts could scan all-time, so pre-cutoff
**paid** bills are out of the audit's window.

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
