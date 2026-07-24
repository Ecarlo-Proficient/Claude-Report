# health-dashboard — STATUS

Shared progression record (see repo CLAUDE.md). Tool-scoped only — no dollar
exposures or business analysis here (those live in the owner's vault).

## DONE / FINALIZED
- **`money_bleeds.py`** — the current company-health report (an exceptions
  "watchboard", supersedes the KPI idea). Read-only. Output
  `~/Documents/CompanyHealth/Money Bleeds.xlsx`, chmod 600. Sheets:
  Dashboard (colored KPI cards), Draws MFD, Draws CP, Lien Clock (grouped by
  status), Lien Retainage, Leases (excluded), RP Wrap-Up, Unused POs 30d+,
  Open Bills (AP).
- Checks: (1) draws billed-not-invoiced — MFD (latest draw folder vs latest
  QBO invoice) + CP (latest G702 earned-less-retainage vs cumulative QBO
  invoiced); (2) Texas lien-notice clock on open construction invoices
  (work month = invoice month; OK rows hidden; retainage + lease/note split
  off); (3) RP slab wrap-up from the Test - RP tab; (4) unused QBO POs ≥30d;
  (5) open bills (AP) read from the Bill Tracker, grouped by AR state × the
  sub's lien clock.
- Rich Excel formatting (colored tabs, zebra, data bars, color scales,
  grouped bands, subtotals) — a deliberate, owner-requested exception to the
  repo's plain-Excel rule.
- Draw discovery + G702 parsing extracted to `shared/draws.py` (shared with
  `wip/cp_wip_reader.py`).

## OPEN ISSUES
- **SMB volumes + Touch ID**: hard-fails unless `Multi Family` and `Common`
  are mounted; one Touch ID per run. Blocks a fully unattended morning run.
- **Bill Tracker freshness**: `Open Bills (AP)` is only as current as the last
  `sync-ap`; the dashboard surfaces the tracker's own sync date.
- **`qbo_health.py`** (legacy KPI dashboard) still present, superseded — decide
  whether to retire or keep for its AR/AP-aging + cash sheets.

## DONE (cont.)
- **`company_dashboard.py`** — consolidated HTML view. Reads the tracker
  workbooks (Money Bleeds, Sub LOC) and emits one self-contained
  `Company Dashboard.html` (chmod 600): Money Bleeds KPI tiles, Sub LOC peak +
  by-division bars, per-source freshness badges. No QBO/Touch ID. First step of
  the "one dashboard reading all trackers" goal.

## TO DO (requested, not yet built)
- **One `Company Tracker.xlsx`** — fold every tracker into a single workbook (Money In /
  Money Out / Position tabs) that the HTML breaks down (the user wants one file, not many).
  Next focused build.
- Bill Tracker detail into the dashboard (currently freshness-badge only).
- ~~Money-out check register~~ DONE — `one-offs/money_out_register.py` builds a
  stateful, self-reconciled register (user-owned CLEARED? column, marks preserved across
  runs); company_dashboard surfaces the aged->30d unmarked checks as the chase list. QBO
  still can't auto-detect cleared status — that's why the flag is manual.
- Fold the legacy hard-fact sheets (AR/AP aging, cash) into the watchboard.
- Auto morning refresh (launchd on this Mac) + a routine that reads the
  workbook/HTML and messages a summary. Needs the Touch-ID/volume constraints solved.
- **Bid List reconciliation** (Notion Bid List ↔ `Bid List
  Residential-Commercial.xlsx`, match on project #): needs a SHARED Notion
  client for headless runs (the interactive MCP won't run under a scheduler).
