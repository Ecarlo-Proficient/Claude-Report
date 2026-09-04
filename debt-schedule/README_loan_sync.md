# Loan Sync — one-click QBO → Equipment Debt Schedule

Pulls loan activity from QuickBooks into
`debt-schedule/Equipment_Debt_Schedule_v2.xlsx`:

- **Current Balance** (Master col **G**) ← the loan's QBO liability-account *actual*
  balance. No principal-vs-interest split is computed or needed.
- **Payments** hitting each loan account → a **QBO Payment Ledger** sheet
  (one row per payment, never duplicated on re-run).
- **As-of Date** (Master col **H**) ← the day you synced.

The amortization tabs and every formula are left untouched.

---

## The merged workbook (`Equipment_Debt_Schedule_v2.xlsx`)

Built from your `Copy of Monthly payments` file + the prior schedule. 28 loans,
each with an amortization tab. New on the Master:

- **Company** (col A) — `Proficient` or `L&A Holdings`. The two American National
  Bank loans (CF Hawn office, Balch Springs) are L&A.
- **Account #** (col E) — the lender loan/account number, used to match each loan
  to its QuickBooks account.

### Highlight colors (what needs your attention)

- **Yellow** cell = a term the amortization needs is still blank.
- **Orange** Account # cell = no account number → can't look up QBO yet.
- **Light-green** Company cell = L&A Holdings (separate QuickBooks file).

### Where to enter terms (for amortization)

On each Master row, fill:

| Col | Field | Notes |
|-----|-------|-------|
| F | Original Loan Balance | yellow if missing |
| J | Term (months) | yellow if missing |
| K | Annual Rate % | leave blank (grey) to auto-compute via RATE() |
| L | Start Date | yellow — needed for exact month scheduling |
| M | Pmt Day | day of month the draft hits (default 1) |

`Current Balance` (G) and `As-of` (H) are maintained by the QBO sync — you don't
type those. **9 loans still need an Original**, **3 need a Term** (the new/L&A
loans) — they're highlighted yellow.

---

## L&A Holdings = separate QuickBooks file

L&A is its own QBO company, so the Proficient login can't see its transactions.
By default the sync only touches **Proficient** rows and skips L&A (shown green in
QBO Setup). To sync L&A too, set up its credentials and run:

```
python3 loan_sync.py --discover --company "L&A Holdings"
python3 loan_sync.py --company "L&A Holdings"
```

Until then, keep the two L&A balances current by hand.

---

## First-time setup (≈5 min, once)

1. **Map loans to QBO accounts:**
   ```
   cd "/ABSOLUTE/PATH/TO/Automate Concrete Business"
   python3 loan_sync.py --discover
   ```
   One Touch ID. Pulls your QBO liability accounts and writes a **`QBO Setup`**
   sheet, auto-matching by lender account number (col E) — usually exact.

2. **Confirm the mapping** — open the workbook → `QBO Setup`:
   - Yellow rows = low-confidence, check first. Green rows = L&A (leave them).
   - Fix a wrong match: pick the account name in **col E**, put its **Id in col F**
     (re-run `--discover` to auto-fill the Id).
   - Set **CONFIRM = Y** (col J) on every Proficient loan. **Nothing syncs until
     CONFIRM = Y.**
   - Save and close.

3. **Preview:**
   ```
   python3 loan_sync.py --dry-run
   ```
   Writes `..._PREVIEW.xlsx`, changes nothing live, prints every balance change +
   new payment.

---

## The one click (ongoing)

```
python3 loan_sync.py
```
Touch ID → backup → refresh balances → append new payments → save. Re-run anytime;
recorded payments are skipped. Widen the window with `--since 2024-01-01`.

---

## Confirm on your first live run

- **[V1] Balance sign** — we show `abs(CurrentBalance)`. Confirm they match lender
  statements.
- **[V2] Payment type** — we scan Purchases (checks/expenses), Bills, Journal
  Entries. If a loan's payments don't appear, send one example txn.
- **[V3] Mapping** — every balance in QBO Setup lines up with the right asset.

## Safety

- Live workbook is copied to a timestamped backup before any write
  (`~/Library/Application Support/proficient-automation/debt-schedule-backups/`).
- `--dry-run` never touches the live file. Re-runs never duplicate payments.
- Logs: `~/Library/Logs/Proficient/loan_sync.log`. Close the workbook in Excel
  before running.
