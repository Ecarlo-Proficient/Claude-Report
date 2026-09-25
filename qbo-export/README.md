# QBO Transaction Export

Export QuickBooks Online transactions to a flat xlsx table — one row per
line item, with QBO account name as the Category column. Output lands in
the OneDrive inbox (`-Inbox- Project Report Exports`).

Auth comes from the shared vault (`shared/qbo_vault.py`); set it up once
with `python3 shared/setup_qbo.py` (see the root README). Commands below
run from the repo root.

## Everyday commands

```bash
python3 shared/setup_qbo.py --status          # what's stored (Touch ID to see keys)
python3 shared/setup_qbo.py --test            # auth test only, no prompts
python3 shared/setup_qbo.py --rotate QBO_CLIENT_ID   # rotate one key — others untouched
python3 shared/setup_qbo.py --purge           # wipe the blob entirely
python3 qbo-export/qbo_export.py --since 2025-01-01      # custom start date
python3 qbo-export/qbo_export.py --all                    # no date filter
```

## What you need (4 values)

**From Intuit Developer — intuit.com → your app → Keys & OAuth → Production Keys tab:**

- `QBO_CLIENT_ID` — from the Production Keys tab
- `QBO_CLIENT_SECRET` — from the Production Keys tab (same view)

**From QBO — gear → Account and Settings → Billing & Subscription:**

- `QBO_COMPANY_ID` — the long digit string

**From your OAuth authorization flow:**

- `QBO_REFRESH_TOKEN` — captured when you authorized your production app

That's it. No environment setting, no dev/prod toggle.

## If auth fails

`shared/setup_qbo.py --test` prints the specific failure and the exact `--rotate` command:

| Error | What it means | Fix |
|-------|---------------|-----|
| `invalid_client` | Client ID came from Development Keys tab instead of Production | Re-copy both from Production Keys tab, `--rotate QBO_CLIENT_ID` and `--rotate QBO_CLIENT_SECRET` |
| `invalid_grant` | Refresh token expired (100-day max) | Re-authorize your production app, `--rotate QBO_REFRESH_TOKEN` |
| Company probe 401/404 | Company ID doesn't match app's authorized realm | `--rotate QBO_COMPANY_ID` |

## Output

`qbexp_transactions_<YYYYMMDD_HHMMSS>.xlsx` → OneDrive inbox (`-Inbox- Project Report Exports`).

Columns: Txn Date, Type, Doc #, Name, Project #, Account (Category), Description, Amount, Memo, Txn ID.

Sources pulled: Bill, Purchase (cash/cc/check), JournalEntry, Invoice. One row per line item.

