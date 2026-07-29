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
- **Labor + Concrete sheets** (CP, 2026-07-29) — the PM/ops manager's main
  view. Rows = that trade's cost codes (`*6` labor, `*1` concrete), columns =
  draw windows, each code expands to the bills behind it (amount in the draw
  column it landed in, so the code row is the sum of the rows beneath).
  Each draw is a **QTY · RATE · TOTAL** group under one merged header
  carrying the period, with a thick rule down both sides of the group. QTY
  and RATE only count quantity-bearing lines — a lump bill booked as
  "1 × $8,901.84" is a price, not a quantity, and would wreck the blended
  rate; TOTAL still carries every dollar. SALES TAX and FUEL SURCHARGE are
  their own columns and are folded onto the bill row they came from (joined
  by bill #), never into the budget comparison — the takeoff budget is
  pre-tax. Fuel is empty on every job today and the sheet says why (AP folds
  it into the rate). BUDGET / BALANCE $ / BALANCE % sit together on the LEFT
  and the bills open collapsed, so the sheet shows no blank cells at rest
  (the user's call, 2026-07-29). Bill rows lead with the QBO bill # — that's
  the identifier, not the date. Font 12, uniform row height, long
  descriptions clip rather than grow the row. Concrete adds a horizontal
  yards / $/yd strip vs the takeoff's implied rate, with no-yardage lump
  bills excluded and listed by bill #.
- **Contract price + approved COs from the G702** (CP, 2026-07-29) — the
  signed pay application beats the WIP master AND any hand-typed cell; the
  P&L prints the source on the contract line itself, and that cell is no
  longer a yellow input (yellow means the user typed it). Reader is `shared/draws.py::read_pay_app`
  (handles the legacy .xls template whose sheets are named 'A'/'B', which the
  existing `G702`-sheet reader can't see). **Needs `xlrd`** — without it the
  run warns and falls back to the WIP master.
- **Draw sheets lead with a horizontal KPI strip** (the user 2026-07-29):
  income → retainage held → net draw → costs → gross profit → gross margin %
  → overhead → REAL net profit → REAL net %, in big type, profit cells
  colored by sign via conditional formatting, and the strip is the freeze
  pane so everything below scrolls under it. MFD keeps its PM-vs-QBO
  comparison as a second strip. Replaces the old vertical summary box.
- Overhead: 10% of revenue (MFD alt view 9% on costs).

## OPEN ISSUES

- **Only 3 of 17 Active CP jobs have a readable cost-code budget** (CP585,
  CP672, CP745). Of the rest: 9 use a descriptive `Cost Codes` sheet
  (`DRILLED PIERS / Concrete / Rebar / Labor` down col A — no SL#/PV# codes,
  so nothing joins to QBO), 4 have no cost-code sheet at all, 1 (CP865) has no
  folder under the awarded-projects root. Blocked on a decision: teach the
  reader the second template, or have the estimators convert those jobs.
  CP910 carries its codes on a `COST QB` sheet — possible lead.
- **CP745 labor budget is short $3,735** — the takeoff's Labor Report totals
  $82,937.80 but the QBO budget import carries $79,202.80; bollards ($2,375)
  and the dumpster beam ($1,360) never made it into the cost codes.
- **Concrete $/yd budget looks high** — CP745's implied rate is $168.12/yd
  (cost-code $ ÷ takeoff yards) against $100.42/yd actually paid. The
  takeoff's CONCRETE report totals $143,718.72 vs $139,674.34 in the cost
  codes ($4,044.38 of curb + bollards never coded), so the two bases differ.
  Worth confirming with the estimators what the CONCRETE report is measuring.
- **Fuel surcharge is not reported** (the user 2026-07-29) — AP clerks folded
  it into the per-yard rate instead of coding it separately. Deliberately
  omitted until AP re-enters those bills correctly.
- CO costs are still a manual yellow input (no CO cost template in QBO yet).

## TO DO

- Extend the Labor/Concrete sheets to RP and MFD if the PMs want them there.
- Roll the G702 contract source into the WIP readers so the master and the
  P&L can't disagree.
