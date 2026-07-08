# QBO Transaction Export

Export QuickBooks Online transactions to a flat xlsx table — one row per line item, with QBO account name as the Category column.

## How the Keychain works

All 4 QBO keys live inside **one** Keychain entry (service `automation-qbo`, label `credentials`) as an encrypted JSON blob. **One Touch ID prompt per script run unlocks everything.** After that first approval in a given process, keys are cached in memory — no re-prompting per key.

When Notion / Teams get added later, each gets its **own** Keychain blob (`automation-notion`, `automation-teams`). A bad paste in one service can never corrupt another, and rotating one service's keys never touches another's.

## Production only — no env selector

This tool is hardcoded to call QBO's production API (`quickbooks.api.intuit.com`). Sandbox has been intentionally removed. One less thing to configure, one less class of mistake.

## Files

| File | What it does |
|------|--------------|
| `qbo_vault.py` | Keychain blob helper — QBO only |
| `setup_qbo.py` | Interactive setup + built-in auth test |
| `qbo_export.py` | The actual export — writes xlsx to OneDrive inbox |
| `qbo_health.py` | Local company health dashboard — AR/AP coverage, P&L, cash, anomalies |
| `requirements.txt` | `requests` + `openpyxl` |

## Run order

```bash
cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business"

# 1. install deps (one time)
pip3 install --break-system-packages -r requirements.txt

# 2. store credentials, auth test runs automatically at the end
python3 setup_qbo.py

# 3. export — one Touch ID prompt, then the xlsx lands in OneDrive inbox
python3 qbo_export.py
```

## Everyday commands

```bash
python3 setup_qbo.py --status          # what's stored (Touch ID to see keys)
python3 setup_qbo.py --test            # auth test only, no prompts
python3 setup_qbo.py --rotate QBO_CLIENT_ID   # rotate one key — others untouched
python3 setup_qbo.py --purge           # wipe the blob entirely
python3 qbo_export.py --since 2025-01-01      # custom start date
python3 qbo_export.py --all                    # no date filter
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

`setup_qbo.py --test` prints the specific failure and the exact `--rotate` command:

| Error | What it means | Fix |
|-------|---------------|-----|
| `invalid_client` | Client ID came from Development Keys tab instead of Production | Re-copy both from Production Keys tab, `--rotate QBO_CLIENT_ID` and `--rotate QBO_CLIENT_SECRET` |
| `invalid_grant` | Refresh token expired (100-day max) | Re-authorize your production app, `--rotate QBO_REFRESH_TOKEN` |
| Company probe 401/404 | Company ID doesn't match app's authorized realm | `--rotate QBO_COMPANY_ID` |

## Output

`qbexp_transactions_<YYYYMMDD_HHMMSS>.xlsx` → OneDrive inbox (`-Inbox- Project Report Exports`).

Columns: Txn Date, Type, Doc #, Name, Project #, Account (Category), Description, Amount, Memo, Txn ID.

Sources pulled: Bill, Purchase (cash/cc/check), JournalEntry, Invoice. One row per line item.

---

# Company Health Dashboard (`qbo_health.py`)

Multi-sheet local Excel dashboard built from live QBO data — designed to answer **"where should I be looking today?"** at a glance. Reuses the same Keychain blob as the transaction export: one Touch ID per run.

## Run

```bash
python3 qbo_health.py                   # refresh to default path
python3 qbo_health.py --out /path/x.xlsx  # override
python3 qbo_health.py --anomaly-sigma 2.5 # tighten the spike threshold
```

## Default output (private, not synced)

```
~/Documents/CompanyHealth/health_dashboard.xlsx   (chmod 600 after every write)
```

This path is intentionally outside OneDrive, iCloud Drive, and the project folder. `chmod 600` means only your Mac user can read it. Combined with FileVault, the file is encrypted on disk while the laptop is locked.

## Sheets

| Sheet | What it shows |
|-------|---------------|
| **Dashboard** | Single-view KPIs + flagged "where to look" callouts |
| **Cash** | Accounts with balances, cash runway (weeks at current burn), last 13 weekly net cash flows |
| **AR Aging** | Every open invoice with bucket (Current / 1-30 / 31-60 / 61-90 / 90+) |
| **AP Aging** | Every open bill with the same buckets |
| **Coverage** | AR vs AP by bucket with Net and Coverage Ratio (>1 = inflows cover outflows) |
| **Relationships** | Top 10 customers by YTD revenue + top 10 vendors by YTD spend; concentration risk callout if any customer >25% |
| **Collections** | DSO and DPO trended monthly over 12 months; 3-month averages summary |
| **P&L** | MTD / YTD / prior-year YTD with YoY delta; Gross Margin % computed live |
| **Anomalies** | Overhead accounts spending ≥2σ above their 90-day weekly average + top 10 largest recent line items |
| `_Meta` (hidden) | Generated timestamp, sigma threshold, output path |

Dashboard flags include: AR aged 60+, AP aged 60+, AR < AP overall, overhead spikes, largest single AR/AP balance, YTD Net Income behind prior year, cash runway <16/<8 weeks, customer concentration risk (any customer >25% of YTD revenue), DSO trending up (last-3-mo vs prior-3-mo delta >5 days when >45). If none trigger, you get "All green."

## Why Python can't set Excel's "password on open"

Neither `openpyxl` nor `xlsxwriter` can write the AES-encrypted OOXML format that Excel's **Protect Workbook → Encrypt with Password** feature produces. That dialog ONLY works interactively in Excel itself. So this dashboard uses a different privacy stack:

1. **Private path** — outside any sync folder (default `~/Documents/CompanyHealth/`)
2. **File permissions** — `chmod 600` after every write, owner-only
3. **FileVault** — whole-disk encryption when the laptop is locked
4. **(Optional) Encrypted .dmg wrapper** — see below for maximum at-rest protection

## Optional: point output at an encrypted .dmg

One-time setup:

```bash
# Create a 100 MB encrypted disk image. You'll be prompted for a password.
hdiutil create -encryption AES-256 -stdinpass -size 100m \
  -volname Health -fs HFS+ ~/CompanyHealth.dmg
```

Daily use:

```bash
# Mount the image (prompts for password), then run the dashboard into it:
hdiutil attach ~/CompanyHealth.dmg
python3 qbo_health.py --out /Volumes/Health/health_dashboard.xlsx

# When you're done, eject — file is encrypted at rest until next mount:
hdiutil detach /Volumes/Health
```

## Schedule it

### Mac (launchd is the modern choice; cron works too)

Create `~/Library/LaunchAgents/com.proficient.qbo-health.plist`:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>com.proficient.qbo-health</string>
  <key>ProgramArguments</key>
  <array>
    <string>/usr/bin/python3</string>
    <string>/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/qbo_health.py</string>
  </array>
  <key>StartCalendarInterval</key>
  <dict>
    <key>Hour</key><integer>6</integer>
    <key>Minute</key><integer>0</integer>
  </dict>
  <key>StandardOutPath</key><string>/tmp/qbo_health.out</string>
  <key>StandardErrorPath</key><string>/tmp/qbo_health.err</string>
</dict>
</plist>
```

```bash
launchctl load ~/Library/LaunchAgents/com.proficient.qbo-health.plist
```

Daily at 06:00. Touch ID will prompt the first time the agent fires after a reboot — approve with **Always Allow** to stop future prompts.

### Or cron, if you prefer

```
0 6 * * * cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business" && /usr/bin/python3 qbo_health.py >> /tmp/qbo_health.log 2>&1
```

## Tuning overhead spike detection

The "Anomalies" sheet flags overhead accounts whose last-7-days spend is ≥ N standard deviations above their prior 90 days of weekly averages. Default is 2σ. To change:

```bash
python3 qbo_health.py --anomaly-sigma 2.5   # stricter (fewer flags)
python3 qbo_health.py --anomaly-sigma 1.5   # looser (more flags)
```

What counts as "overhead" is matched by substring against a hint list in `qbo_health.py` (`OVERHEAD_HINTS`) — office, insurance, utilities, rent, fuel, dues, meals, travel, legal, software, advertising, repairs, maintenance, etc. Edit that list if your QBO Chart of Accounts uses different names.

