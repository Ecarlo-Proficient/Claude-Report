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
- Checks: (1) draws billed-not-invoiced — MFD (latest BUILT draw folder vs
  latest QBO invoice) + CP (latest G702 earned-less-retainage vs cumulative
  QBO invoiced); (2) Texas lien-notice clock on open construction invoices
  (work month = invoice month; OK rows hidden; retainage + lease/note split
  off); (3) RP slab wrap-up from the Test - RP tab; (4) unused QBO POs ≥30d;
  (5) open bills (AP) read from the Bill Tracker, grouped by AR state × the
  sub's lien clock.
- Rich Excel formatting (colored tabs, zebra, data bars, color scales,
  grouped bands, subtotals) — a deliberate, owner-requested exception to the
  repo's plain-Excel rule.
- Draw discovery + G702 parsing extracted to `shared/draws.py` (shared with
  `wip/cp_wip_reader.py`).
- **Revised draws supersede originals** (the user 2026-08-04). CP765 filed
  `Revised LP Draw Excel #4` beside `Draw Excel #4` — same draw number, so the
  winner fell out of filesystem order and the dead original won, inventing a
  $13,552 shortfall against an invoice that was correct to the dollar.
  `shared/draws._supersedes()` now ranks `revis*` over the original, then newest
  mtime. CP765 is the only project whose file changes; `Draws CP` went 1 RED →
  0 RED, and CP765's WIP row moves too (see `wip/STATUS.md`).
- **MFD built-vs-staged draw detection** (the user 2026-08-04). MFD draws are
  PDFs — no G702, so no amount to match. `mfd_draw_documents()` reads the draw
  folder and calls it BUILT only if it holds a pay application; the lien-release
  test runs first so a "Conditional Release … August Draw 2026.pdf" can't
  masquerade as the draw. A higher-numbered folder with no pay application is a
  placeholder — reported in the new `STAGED NEXT` column, never judged, because
  there is nothing to invoice until the draw is built. Recurses so a job billing
  several contracts (MFD192: Mayhill + Offsite) counts each contract's pay app.
  Scope is deliberately the latest draw only — the user 2026-08-04, older folders
  predate this naming and enumerating them all would break.

## SUPERSEDED (2026-08-31) - the ledger's Health tab
The Money In / Money Out / Position / Break-Even model + the Recurring & Debt register now
render LIVE in the ledger (`ledger/load_health.py` -> `health_snapshot`; `/api/healthtab`;
`shared/qbo_pl.py` for the P&L blocks; `shared/breakeven.build_from_blocks` - the xlsx-reading
`build()` still works for this stack). The tracker outputs here had been unopened since
2026-08-12. Everything stays runnable until the owner calls its retirement; nothing schedules
it, and the ledger reads none of its files. What the ledger does NOT yet replace:
(a) `qbo_health.py --as-of` reproducible month-end closes, and (b) Money Bleeds' filesystem
draw checks (MFD built-draw / CP G702 shortfall) - the ledger's Draws tab judges draws from
the Bill Tracker, not the draw folders.

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

## DONE (cont.)
- **`company_tracker.py`** — the ONE consolidated workbook (Summary + Money In / Money Out
  / Position tabs) + the HTML, both from one metric model (`company_dashboard.build_sections`).
  Reads all source trackers; Company Tracker.xlsx + Company Dashboard.html, chmod 600.

- **RP Billing Status** (2026-07-28) — `shared/rp_billing.py`; poured-but-unbilled vs
  backlog, folded into Company Tracker.xlsx as its own tab. Replaced the old "RP slabs
  waiting on punch" check (it measured the wrong thing and returned 0).

## TO DO (requested, not yet built)
- Pull transaction-level detail (lien chase list, checks >30d, top over/under jobs) into
  Company Tracker.xlsx as drill-down tabs — currently summary/metric level.
- Bill Tracker detail into the dashboard (currently freshness-badge only).
- Morning auto-refresh (launchd runs the source trackers → company_tracker) once the
  Touch-ID / volume-mount constraints are solved.
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
