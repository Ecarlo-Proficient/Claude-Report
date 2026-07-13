# Bill Payment Tracker

Pulls open AP bills from QBO, matches each line to the GC invoice that
authorizes its payment, and writes a 4-sheet Excel report.

## Files

| File | Purpose |
|---|---|
| `qbo_bill_tracker.py` | Main script — auth, fetch, match, write |
| `run_tracker.sh` | Wrapper for manual runs |
| `probe_draw_period.py` | Diagnostic — confirms whether QBO API returns a custom field |
| `read_private_note.py` | Diagnostic — reads PrivateNote on a single invoice |
| `Bill_Payment_Tracker.xlsx` | Output (created on first run) |
| `logs/run.log` | Run history |

**Runs are manual** — there is no scheduler. Run it from your desk when you
want fresh data (see below). The old launchd auto-run was scrapped.

## Manual run

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business"
bill-tracker/run_tracker.sh
```

Or directly:

```bash
python3 bill-tracker/qbo_bill_tracker.py
python3 bill-tracker/qbo_bill_tracker.py --dry-run        # build but don't write
python3 bill-tracker/qbo_bill_tracker.py --out /path/x.xlsx
```

## Output structure

**Sheet 1 — `Bill Tracker`** (open bills, EDITABLE on Override + Notes only)

Columns: Bill Date · Vendor · Bill Ref # · Bill Total · Bill Open Bal ·
Customer/Project · Division · Line Amount · Line Description ·
Matched Invoice # · Invoice Date · Draw Period · Invoice Total ·
Invoice Open Bal · **Status (Auto)** · **Status Override** ·
Status (Final) · **Notes** · _Key (hidden)

Sheet protection is on but the **Status Override** and **Notes** cells
are unlocked, so users can edit those without unprotecting the sheet.
The Override column has a dropdown: `OK TO PAY`, `ON HOLD`, `REVIEW`,
`DISPUTED` — but free text is allowed.

**Sheet 2 — `By Vendor`** (read-only, grouped, subtotals on Line Amount + Bill Open Bal)

**Sheet 3 — `By Project`** (read-only, grouped by Division → Project, subtotals)

**Sheet 4 — `Archive`** (read-only) — bills that were once on Sheet 1 but
are now closed. Notes / Status Override survive the move.

## How matching works

For each non-sub bill line:

1. Read project # from the line's CustomerRef name (`MFD###`, `CP###`, `RP####`)
2. Apply division-specific rule:
   - **RP** → earliest invoice on/after bill date for the same project
   - **MFD / CP** → invoice whose Draw Period (parsed from `PrivateNote`)
     contains the bill date
3. If no match: `AWAITING INVOICE`
4. If matched and invoice Balance == 0: `OK TO PAY`
5. If matched but invoice Balance > 0: `AWAITING INVOICE PAYMENT`
6. If line has no project #: `NO PROJECT #`
7. If bill Balance == 0 (paid): `PAID` — only seen on Archive sheet

## Exclusions

- **Sub bills** — any Bill where `PrivateNote` (memo) contains "sub"
- **Retainage invoices** (excluded from matching pool):
  - `PrivateNote` contains "retainage not billed", OR
  - every non-empty line description contains "retainage"
- **Regular draws with retainage withholding lines stay included** — only
  invoices that are 100% retainage are dropped.

## Manual edits — what's preserved

The script preserves your **Status Override** and **Notes** entries
across every run, keyed by `Bill.Id-Line.Id`. So you can:

- Set an Override → it sticks
- Type a Note → it sticks
- A bill goes paid → row moves to Archive with notes intact
- A check bounces, bill reopens → row comes back to Sheet 1 with notes

You can edit **only on Sheet 1**. Sheets 2-4 are locked. If you do
unprotect Sheet 2 or 3 to edit, those edits are lost on the next run.

## File-lock handling

If `Bill_Payment_Tracker.xlsx` is open in Excel when the script runs,
the script writes to a sidecar `Bill_Payment_Tracker.PENDING.xlsx`
instead. Close the canonical file and the next run will overwrite it
with fresh data.

## Audits (read-only, one-off)

Standalone catchers that share the tracker's QBO auth/paging. Each writes a
plain xlsx to the OneDrive `QBO Audits` folder (override with `--out`) and
never edits QBO. All accept an optional `since [until]` date window and
`--dry-run`.

| Script | Flags | Catches |
|---|---|---|
| `sub_bill_audit.py` | `[since] [until]` | Sub bills (memo starts "Sub") with a line missing a project # |
| `item_no_project_audit.py` | `[since] [until]` | Item-based bill lines (job costs) missing a project # |
| `duplicate_bill_audit.py` | `[since] [until]` | Same Bill Ref # entered on 2+ bills within one **vendor tree** (root + sub-vendors) — double-entry / double-pay risk |
| `job_coding_audit.py` | `--job <code>` | Lines that mention a job but are coded to the wrong/missing project or a non-parent class |

```bash
python3 bill-tracker/duplicate_bill_audit.py                 # all bills
python3 bill-tracker/duplicate_bill_audit.py 2026-01-01      # dated on/after
python3 bill-tracker/duplicate_bill_audit.py --dry-run       # counts only
```

`duplicate_bill_audit.py` groups bills by (top-level vendor, ref #),
case-insensitive and whitespace-trimmed; blank ref #s are skipped. Output is
grouped by vendor tree with a `Same amount?` flag — matching totals across
copies are the strongest double-entry signal.

## Touch ID

The QBO Keychain entry (`automation-qbo`) was created with a biometric
ACL, so every run prompts Touch ID once. That's expected — the tracker is
run manually from your desk, so someone is there to approve it.

## Running diagnostics

```bash
# verify QBO returns the custom field "Draw Period" (it doesn't — confirms our PrivateNote workaround is needed)
python3 bill-tracker/probe_draw_period.py

# read PrivateNote on a specific invoice to confirm draw period text is reachable
python3 bill-tracker/read_private_note.py 34099
```

## Troubleshooting

**`✗ no credentials. Run: python3 setup_qbo.py`** — the Keychain entry
was wiped or never created.

**`token refresh status=400`** — refresh token expired (100-day limit
of inactivity) or rotated by another process. Re-run `setup_qbo.py
--rotate QBO_REFRESH_TOKEN`.

**`could not write canonical file`** — the file is open in Excel.
Sidecar `Bill_Payment_Tracker.PENDING.xlsx` was written instead.

**Override edits disappear** — you edited on Sheet 2/3/4 instead of
Sheet 1. All edits must happen on Sheet 1.
