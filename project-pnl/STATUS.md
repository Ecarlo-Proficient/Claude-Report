# project-pnl — STATUS

Per-project P&L workbooks from QBO. One tool, one script
(`project_pnl_export.py`), three templates: CP (draw-based), MFD (draw-based,
manual close), RP (no draws — expenses → invoice → profit).

---

## DONE / FINALIZED

- **P&L + Transactions + Draws + POs + Cash Flow + Reconciliations** sheets.
- **Batch mode** — `project-pnl active cp|rp|mfd` regenerates every Active
  project of a division (Active = the WIP master's Test-Master STATUS).
- **Budget vs Actual** (CP + RP) — takeoff cost-code budget vs QBO cost-code
  actuals, every transaction listed under its code with a QBO link, job-type
  color bands, class-mismatch flags.
- **Labor + Concrete sheets** (CP, reworked 2026-07-29 pm) — two blocks at
  different altitudes, per the user: metrics are a top-level data point, and
  one grid trying to be scoreboard AND ledger is what produced empty cells
  and hidden rows.
  **SCOREBOARD** (frozen top): one row per cost code — BUDGET · ACTUAL ·
  BALANCE $ · BALANCE % · one total per draw (header carries the period);
  over-budget red. Concrete adds SALES TAX and ACTUAL INCL. columns, and the
  grand total names the P&L line it ties to ('Job Materials: Concrete' /
  'Subcontractors Expense: Labor'). Concrete also gets the horizontal
  yards/$-per-yd strip (takeoff implied vs paid, lump bills excluded from the
  rate and flagged).
  **LEDGER** (below, fully expanded, nothing collapsed): QBO # (linked,
  the user's identifier) · DATE · VENDOR · DESCRIPTION · QTY · RATE ·
  AMOUNT · [SALES TAX] · DRAW label. A bill lands in exactly one draw, so a
  label column replaces the per-draw matrix that guaranteed blank cells on
  every bill row. Tax folds onto the bill row it came from (joined by bill
  #). Tax/fuel columns appear ONLY on a trade that has such lines — labor
  subs bill neither (auto-omit, disclosed). No fuel lines exist anywhere yet
  (AP folds the surcharge into the rate); Concrete says so on the sheet.
  Font 12 flat; uniform row heights. DESCRIPTION is the ledger's LAST
  column and spills right over empty space — scoreboard and ledger share
  physical columns, so a wide mid-table description column was inflating
  BALANCE $ above it (the user 2026-07-29). Draw headers are wide enough for
  the full period. NO BUDGET → NO SHEET: with the takeoff unreadable (e.g.
  Common drive unmounted) the sheets are skipped with a warning, because a
  scoreboard of $0 budgets reads as wildly over budget.
- **Contract price + approved COs from the G702** (CP, 2026-07-29) — the
  signed pay application beats the WIP master AND any hand-typed cell; the
  P&L prints the source on the contract line itself, and that cell is no
  longer a yellow input (yellow means the user typed it). Reader is `shared/draws.py::read_pay_app`
  (handles the legacy .xls template whose sheets are named 'A'/'B', which the
  existing `G702`-sheet reader can't see). **Needs `xlrd`** — without it the
  run warns and falls back to the WIP master.
- **Draw sheets lead with a horizontal KPI strip** (the user 2026-07-29):
  income → retainage held → net draw → costs → gross profit → gross margin %
  → overhead → REAL net profit → REAL net %, big type, $-formatted, profit
  cells colored by sign, and the strip is the freeze pane. MFD keeps its
  PM-vs-QBO comparison as a second strip. Replaces the old vertical summary
  box. Draw-sheet body font is 12 (was 11).
- Overhead: 10% of revenue (MFD alt view 9% on costs).

## OPEN ISSUES

- **Only 3 of 17 Active CP jobs have a readable cost-code budget** (CP585,
  CP672, CP745). Of the rest: 9 use a descriptive `Cost Codes` sheet
  (`DRILLED PIERS / Concrete / Rebar / Labor` down col A — no SL#/PV# codes,
  so nothing joins to QBO), 4 have no cost-code sheet at all, 1 (CP865) has no
  folder under the awarded-projects root. Blocked on a decision: teach the
  reader the second template, or have the estimators convert those jobs.
  CP910 carries its codes on a `COST QB` sheet — possible lead.
- **Two CP745 budget-gap findings** — (a) the labor budget imported into QBO
  runs short of the takeoff's Labor Report (bollards + the dumpster beam never
  made it into the cost codes); (b) the implied concrete $/yd from cost codes
  runs well above the rate actually paid (curb + bollards never coded, so the
  two bases differ). Worth confirming with the estimators what the CONCRETE
  report is measuring. Dollar detail lives in the owner's vault — scrubbed
  from the repo 2026-07-30 per the STATUS scope filter.
- **Fuel surcharge is not reported** (the user 2026-07-29) — AP clerks folded
  it into the per-yard rate instead of coding it separately. Deliberately
  omitted until AP re-enters those bills correctly.
- CO costs are still a manual yellow input (no CO cost template in QBO yet).

## TO DO

- Extend the Labor/Concrete sheets to RP and MFD if the PMs want them there.
- Roll the G702 contract source into the WIP readers so the master and the
  P&L can't disagree.
