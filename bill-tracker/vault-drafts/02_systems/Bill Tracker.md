---
tags: [system, ap, qbo, excel, automation]
status: live
type: system
created: 2026-04-28
moved-to-systems: 2026-05-27
source: QBO Enterprise
output: /Users/sebas/Documents/CompanyHealth/Bill Tracker.xlsx
schedule: Mon–Fri 15:00 via launchd
---

# Bill Tracker

Daily AP workflow tracker. Pulls open + recently-paid bills from QBO, writes a 5-sheet Excel workbook to `CompanyHealth/Bill Tracker.xlsx`. Single source of truth for what to pay, what's blocked, and bill status by project / client / division.

## Purpose

Four questions, one workbook:

1. **What can I cut a check on today?** → Pay List
2. **Why isn't this other bill ready to pay yet?** → Open Bills
3. **How much AP do I owe per project (open + paid YTD)?** → Project Ledger / Client Ledger / By Division
4. **Client calls about Bill X — what's its status?** → any ledger sheet, AutoFilter by Bill #

## Workbook structure

Five visible sheets + one hidden audit sheet. Every visible sheet is an Excel Table with AutoFilter on every column.

| Sheet | Filter | Sort |
|---|---|---|
| **Pay List** | Pay = x AND open | Vendor → Bill Date |
| **Open Bills** | Open only (Balance > 0) | Vendor → Bill Date → Line |
| **Project Ledger** | Open + paid since `2026-01-01` | Project # → Vendor → Bill Date |
| **Client Ledger** | Open + paid since `2026-01-01` | GC → Project → Bill Date |
| **By Division** | Open + paid since `2026-01-01` | Division → Project → Invoice → Bill |
| `_Audit` | (hidden) | last-sync Pay/Notes snapshot |

### Unified column layout

Same 25 columns on every visible sheet:

```
Vendor · Bill # · Bill Date · Days Open · Bucket · Project # · Division · 
Bill Type · Account · Line Description · Line Amount │ Matched Invoice · 
Invoice Date · Invoice Total · Payment Date │ Status (Auto) · Approved · 
Pay · Notes · Bill Open Bal · Bill Total · Open · _Key (hidden)
```

Two dark divider columns (1.5-wide) split **Bill info │ Invoice info │ Status & amounts**.

### Conditional formatting (only two rules)

- 🔴 **RED row tint** — Approved = `NOT APPROVED`. Don't pay.
- 🟢 **GREEN cell tint** — Pay = `x`. Visual confirmation of check-run intent.

No other coloring. Color is reserved for these two signals.

## Workflow

1. **Open the workbook** — 3 PM sync has run; data is fresh.
2. **Review Open Bills** — see what's open. Filter by Status, Vendor, Project, anything. Red rows = NOT APPROVED, skip them.
3. **Mark Pay = x** on bills you want to cut checks on. Add Notes if needed. (Edits on any sheet propagate to all 5 via 4-way merge.)
4. **Go to Pay List** — filtered view of your Pay = x marks, vendor-grouped. Print or copy to whatever check-writing tool.
5. **Cut checks** — pay vendors.
6. **Next 3 PM sync** — bills that QBO now shows as paid auto-drop the Pay mark. Loop closes.

## Editable fields

- **Pay** — type `x` to mark for check run. Auto-clears when QBO shows bill paid.
- **Notes** — free text. Always preserved across syncs.

No status override, no other editable columns. Pay is the only "do this" signal.

## Approval rule

Bill is **APPROVED** unless its memo (QBO `PrivateNote`) starts with `NOT APPROVED` (case-insensitive). Tag is set in QBO at bill entry. The rule applies to every bill type — no exceptions.

NOT APPROVED bills:
- Show with red row tint everywhere
- Still appear on Pay List if you mark Pay = x (system never blocks you — your override)

## Auto-uncheck on payment

If a bill flips from open → paid between syncs:
- Bill row stays visible on the 3 ledger sheets (Project / Client / Division) so you can confirm
- Pay column is force-cleared (intent already executed, no need for the mark)
- Notes stay (commentary, not action)

## 4-way merge for Pay / Notes

Edits made on any of the 5 sheets propagate to all 5 on next sync. Mechanism:

1. Read existing workbook's `_Audit` sheet → baseline `{_Key → {Pay, Notes}}`
2. Read each of the 5 visible sheets → current values per `_Key`
3. For each `_Key`: if exactly one sheet differs from baseline, that's the intentional edit
4. Apply to all 5 sheets in the new build, snapshot the new state to `_Audit`

If you uncheck Pay on Pay List, that unmark wins across all sheets on next sync.

## Data sources

- **Open bills** — QBO Bill query: `Balance > 0` (any date), excluding sub-bills (`PrivateNote` contains "sub")
- **Paid bills** — QBO Bill query: `Balance = 0 AND TxnDate >= 2026-01-01`. Adjust cutoff in code when crossing fiscal year.
- **Invoices** — QBO Invoice query, excluding retainage-only invoices
- **Payment dates** — QBO Payment objects joined to invoices via `LinkedTxn`

Invoice → Bill matching uses Customer ID + Division + draw-period regex (see `qbo_bill_tracker.py`).

## Files

```
bill-tracker/
├── excel_bill_sync.py          # main entry — 825 lines
├── bill_rows.py                # shared QBO → row-dict builder
├── qbo_bill_tracker.py         # QBO extraction layer
├── run_tracker.sh              # launchd wrapper
├── launchd/
│   └── com.proficient.billtracker.plist   # Mon–Fri 15:00
├── logs/
│   └── run.log                 # tee'd output of each sync
├── _backups/                   # daily snapshots (14-day retention)
└── _archive/                   # Notion-era files (read-only reference)
```

## Operations

### Manual run

```bash
bash "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/bill-tracker/run_tracker.sh"
```

Touch ID prompts once for QBO Keychain unlock.

### Schedule

`launchctl load ~/Library/LaunchAgents/com.proficient.billtracker.plist` — runs the wrapper at 15:00 Mon-Fri.

### Backups

Daily snapshot to `_backups/Bill Tracker — YYYY-MM-DD.xlsx`. Pruned after 14 days. One snapshot per day max.

### File location + permissions

- `~/Documents/CompanyHealth/Bill Tracker.xlsx`
- `chmod 600` — sensitive AP data, same convention as `qbo_health.xlsx`
- Single-user. For co-authoring, move to OneDrive/Teams (path TBD).

## Known constraints

- **No outline grouping** — conflicts with Excel Tables. AutoFilter is the grouping mechanism.
- **Per-line rows** — multi-line bills show N rows. AutoFilter to a vendor or project naturally clusters them.
- **First run after rewrite has no preserved edits** — old workbook's sheet names don't match the new layout. Subsequent runs preserve perfectly via `_Audit` snapshot.

## Related

- [[QBO Bill Tracker logic]] — extraction + matching algorithms
- [[Project P&L System]] — where deep paid-bill analysis lives (not this tracker)
- [[Statement Reconciler]] — separate workflow for vendor statement matching
- [[QBO Vault]] — Keychain auth used by all sync scripts

## History

- **2026-04-28** — built as Notion DB in new Accounting teamspace
- **2026-05-13** — pivoted from Notion to xlsx (Notion grouping choked at 1,800+ rows)
- **2026-05-27** — full rewrite from 8 sheets to 5; added 4-way merge with `_Audit`; added paid bills since YTD; removed Status Override; moved from `04_projects` to `02_systems` (now operational, not in-progress)
