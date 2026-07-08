# Bill Payment Tracker

Pulls open AP bills from QBO, matches each line to the GC invoice that
authorizes its payment, and writes a 4-sheet Excel report.

## Files

| File | Purpose |
|---|---|
| `qbo_bill_tracker.py` | Main script — auth, fetch, match, write |
| `run_tracker.sh` | Wrapper used by both manual runs and launchd |
| `launchd/com.proficient.billtracker.plist` | macOS scheduler — Mon-Fri 15:00 |
| `probe_draw_period.py` | Diagnostic — confirms whether QBO API returns a custom field |
| `read_private_note.py` | Diagnostic — reads PrivateNote on a single invoice |
| `Bill_Payment_Tracker.xlsx` | Output (created on first run) |
| `logs/run.log` | Run history (manual + scheduled) |
| `logs/launchd_stdout.log` | Scheduled-run stdout |
| `logs/launchd_stderr.log` | Scheduled-run stderr |

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

## Scheduling — install once

The launchd plist runs the wrapper Mon-Fri at 15:00 (3 PM, after
checks have been deposited).

### One-time install

```bash
PLIST=~/Library/LaunchAgents/com.proficient.billtracker.plist
cp "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/bill-tracker/launchd/com.proficient.billtracker.plist" "$PLIST"
launchctl load "$PLIST"
```

### Verify

```bash
launchctl list | grep billtracker        # should show the agent
tail -f "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/bill-tracker/logs/run.log"
```

### Disable temporarily

```bash
launchctl unload ~/Library/LaunchAgents/com.proficient.billtracker.plist
```

### Re-enable

```bash
launchctl load ~/Library/LaunchAgents/com.proficient.billtracker.plist
```

### Uninstall

```bash
launchctl unload ~/Library/LaunchAgents/com.proficient.billtracker.plist
rm ~/Library/LaunchAgents/com.proficient.billtracker.plist
```

## Touch ID and unattended runs

The QBO Keychain entry (`automation-qbo`) was created with biometric
ACL — every read prompts Touch ID. That's fine for manual runs, but a
3 PM scheduled run with no one at the keyboard will **wait for Touch
ID and time out**.

You have two options for unattended scheduling:

**(A) Be at the computer at 3 PM.** Easy if you're at your desk.

**(B) Re-key the Keychain entry to allow specific binaries without
prompting.** Run `setup_qbo.py` once with a flag (or rotate manually)
and add Python and the bill tracker script to the trusted apps list:

```bash
# show current ACL
security find-generic-password -a "$USER" -s automation-qbo -l credentials -g

# wipe and re-add allowing python3 + bash to read without confirmation
python3 setup_qbo.py --purge
python3 setup_qbo.py
# during setup, when prompted, answer that you want unattended access
# (or run `security add-generic-password ... -A` manually — but that's
#  less secure since ANY app can read)
```

For now, the safest path is **(A)** — start with manual + scheduled
that may need approval, see how often you're not at your desk, then
decide if (B) is worth the security trade-off.

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
