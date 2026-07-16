# Company Health (`money_bleeds.py` + legacy `qbo_health.py`)

> Commands below run from the repo root. Uses the shared QBO vault
> (`shared/qbo_vault.py`) — one Touch ID per run.

## Money Bleeds (`money_bleeds.py`) — the current company-health report

The KPI dashboard was retired as "doing too much / not accurate enough"
(2026-07-16). Company health is now an **exceptions report**: a short list of
things that are provably wrong and cost money.

```bash
python3 health-dashboard/money_bleeds.py
python3 health-dashboard/money_bleeds.py --out /path/x.xlsx
```

Output: `~/Documents/CompanyHealth/Money Bleeds.xlsx` (chmod 600). Read-only
against QBO, the WIP workbook, and both volumes. Hard-fails up front if the
`Multi Family` / `Common` volumes aren't mounted or the WIP workbook isn't
synced.

| Sheet | Bleed it catches |
|-------|------------------|
| **Dashboard** | $ totals per bleed + the assumptions used |
| **Draws MFD** | Active MFD projects (from the 'WIP Master' tab) whose **latest** numbered draw folder (`…/PM MISC/DRAWS/N- MONTH YEAR DRAW`) has no QBO invoice in/after the draw month. Latest draw only — history is assumed billed; this is the future-case tripwire. |
| **Draws CP** | CP projects whose latest draw's G702 earned-less-retainage exceeds cumulative QBO invoiced — a draw that never became an invoice. Draw discovery/G702 parsing: `shared/draws.py`. |
| **Lien Clock** | Every open construction invoice with its Texas lien-notice deadline (commercial CP/MFD = 15th of the 3rd month after the work month; residential RP = 15th of the 2nd). Parent customer shown per project row. Sorted by days left: PAST / URGENT ≤15d / WATCH ≤45d. |
| **Lien Retainage** | Retainage invoices — separate statutory track (§ 53.057, completion-based), never mixed into the monthly clock. |
| **Leases (excluded)** | Equipment-lease / note-payment invoices to subs — not construction income, so they carry no lien clock. Listed so the money doesn't silently disappear from the report. |
| **RP Wrap-Up** | SLAB lines 100% complete in the General List but not fully billed (waiting on punch). FTW lines are ignored — the list's 100% column is slab-only. Read from the WIP workbook's `Test - RP` tab, so run the WIP readers first. |

**The lien clock's work-month rule.** Texas deadlines run from the month the
work was performed. Here work month = **invoice month**, by the user's ruling
(2026-07-16): RP invoices go out the day the job finishes, and draws bill
their own work month. Deadlines roll **backward** off weekends (holidays not
modeled). The report is a tripwire, not legal advice — verify the actual work
period before sending any notice.

Formatting note: this workbook uses color (status fills, red row washes) at
the user's explicit request — it is exempt from the repo's plain-Excel rule.

---

## Legacy: KPI dashboard (`qbo_health.py`)

Multi-sheet local Excel dashboard built from live QBO data — designed to answer **"where should I be looking today?"** at a glance. Reuses the same Keychain blob as the transaction export: one Touch ID per run. **Superseded by Money Bleeds** — kept runnable for the hard-fact sheets (AR/AP aging, cash) until formally retired.

## Run

```bash
python3 health-dashboard/qbo_health.py                   # refresh to default path
python3 health-dashboard/qbo_health.py --out /path/x.xlsx  # override
python3 health-dashboard/qbo_health.py --anomaly-sigma 2.5 # tighten the spike threshold
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
python3 health-dashboard/qbo_health.py --out /Volumes/Health/health_dashboard.xlsx

# When you're done, eject — file is encrypted at rest until next mount:
hdiutil detach /Volumes/Health
```

## Schedule it

The launchd and cron examples below use absolute paths (neither has a working
directory) — replace `/ABSOLUTE/PATH/TO/Automate Concrete Business` with this
machine's clone location before loading them.

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
    <string>/ABSOLUTE/PATH/TO/Automate Concrete Business/health-dashboard/qbo_health.py</string>
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
0 6 * * * cd "/ABSOLUTE/PATH/TO/Automate Concrete Business" && /usr/bin/python3 health-dashboard/qbo_health.py >> /tmp/qbo_health.log 2>&1
```

## Tuning overhead spike detection

The "Anomalies" sheet flags overhead accounts whose last-7-days spend is ≥ N standard deviations above their prior 90 days of weekly averages. Default is 2σ. To change:

```bash
python3 health-dashboard/qbo_health.py --anomaly-sigma 2.5   # stricter (fewer flags)
python3 health-dashboard/qbo_health.py --anomaly-sigma 1.5   # looser (more flags)
```

What counts as "overhead" is matched by substring against a hint list in `qbo_health.py` (`OVERHEAD_HINTS`) — office, insurance, utilities, rent, fuel, dues, meals, travel, legal, software, advertising, repairs, maintenance, etc. Edit that list if your QBO Chart of Accounts uses different names.

