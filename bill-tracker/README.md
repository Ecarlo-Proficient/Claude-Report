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
| `QBO Audit` | mixed | six collapsible sections — see below |
| `Liens` | — | live Excel FILTER over the Bills table (Lien set) |

`Lien` and `Notes` on the Bills sheet are user-editable and **preserved** across runs
(keyed by `Bill.Id-Line.Id`).

## QBO Audit sheet — six sections

Each renders its banner + count even at zero (a `✓ none` row), so no check silently
disappears. All share one 11-column layout; every data row deep-links to the bill in QBO.

1. **Stale NOT APPROVED bills** — aged past `NOT_APPROVED_BUFFER_DAYS`, bucketed by age.
2. **Data entry issues** — empty/mismatched Class, line-desc project mismatch.
3. **Likely job cost — MISSING project** — high-confidence uncoded job costs (non-sub).
4. **Duplicate bill #** in a vendor tree — double-entry / double-pay risk (all bills).
5. **FW code on a CP / MFD / RP-slab job** — flatwork miscode. FW is legitimate ONLY on
   the `-FTW` project; it is flagged on any CP job, any MFD job, or a base `RP####` slab.
   Division/slab are read from the project #, never the Class field.
6. **SUB bill line — MISSING project** — a sub bill line with no project #.

Sections 4–6 fold in what the standalone `duplicate_bill_audit.py`,
`item_no_project_audit.py`, and `sub_bill_audit.py` scripts used to do — they were
retired 2026-08-06 (one tool, one workbook). History note: the audit covers the
tracker's population (open + paid-since-`PAID_CUTOFF_DATE`); the retired scripts could
scan all-time, so pre-cutoff **paid** bills are out of the audit's window.

## Audits — the remaining standalone drill

`job_coding_audit.py` stays as an on-demand, per-job investigation (via the `audit-job`
shell alias) — a different workflow from the always-on QBO Audit sheet. It writes a
plain xlsx and never edits QBO.

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
