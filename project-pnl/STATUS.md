# project-pnl — STATUS

Per-project P&L workbooks from QBO. One tool, one script
(`project_pnl_export.py`), three templates: CP (draw-based), MFD (draw-based,
manual close), RP (no draws — expenses → invoice → profit).

---

## DONE / FINALIZED
- 2026-10-06 · **To date row, expenses only when real, Progress out, re-bill guard** (owner: "remove Progress";
  "there should be no expenses ... show if there are so we can remove ... no expenses = that block removed entirely";
  next draw under the total, "a blank thicker line ... remove the income there since it's not known yet"). Coverage:
  TOTAL, a thick rule, Next draw (forming) - costs only - and To date (draws + next draw + outside the windows - office
  expenses), both yellow italic; To date gross profit = D6 on all 21 jobs with draws. ② drops the EXPENSES group when
  a job has none (11 jobs) and GROSS / NET move left. Progress (earned revenue, over / under billing, cost to complete)
  removed - E5's variance already says cost to complete. A PC00 retainage invoice is written to the workbook name
  RetainageIncomeDocs; completed_pnl.retainage_booked counts it, so CP610 / CP595's re-bill releases do not bill twice.
  completed_pnl reads a combined draw by outline DEPTH (the invoices now fold one level in; MFD192 read double for a
  moment). Overviews: CP 6,955,359 (reconciled, 0 problems; QBO ties but CP790's known duplicate); MFD 54,879,453.
- 2026-10-06 · **Transactions sheet: wrapped invoice memos, collapsible invoices, stronger cost dividers** (owner:
  "wrap texts for transactions invoices but only that part, make it groupable"; "keep the costs collapsed, but the
  header/dividers are not pronounced enough"). Invoice memos wrap; every invoice folds under the INCOME header (a
  combined draw's invoices one level deeper); cost lines open folded under their vendor; each cost category is a
  tinted band with a thick navy rule above it, closed by a grey subtotal with a medium rule and a blank row. CP785
  rebuilt and opened in real Excel by script (clean, memo rows auto-sized).
- 2026-10-06 · **Credits reduce job cost; retainage told apart by item account; MFD192 cut fixed** (owner: "credits
  reduce job cost, but don't add a credits line - keep it in the totals and show it in the transaction crediting it").
  `shared/qbo_costs.with_credits`: vendor credits pulled, card credits negated - each a negative transaction under its
  vendor / account, Paid? = CREDIT, QBO link to the credit; the P&L, Overview and ledger cost pull all use it. Invoice
  lines: an item posting to Retainage Receivable (99 - Retainage) is withheld / released, never billed; every other
  line is billed even when its text says retainage (PC00) - P&L and ledger. MFD192's 23,181.91 was the Reconciliations
  date cut reading a combined draw by its first invoice (#34638, 09/07); fixed. Regenerated CP + MFD: CP785, CP961,
  MFD177, MFD192, MFD325 tie to QuickBooks per document (pnl_qbo_gap_trace.py). CP790 skipped (open in Excel). MFD
  Overview billed +6,363 (MFD295 #32575 now income, as QuickBooks books it). Open: a future re-billed release on the
  receivable item (CP610 / CP595) still counts past held + not billed - needs the item on the workbook.
- 2026-10-06 · **The 8 Reconciliations gaps traced** (`one-offs/pnl_qbo_gap_trace.py`, read-only: QuickBooks'
  ProfitAndLossDetail vs the Transactions sheet, per document). Vendor credits are not pulled (bills + expenses only):
  CP785 1,483.03, MFD325 1,540.79, MFD177 1,097.26 + a Home Depot card credit 143.14, CP790 295.30 (net of an AA
  Rental expense 235.50 the sheet misses); CP961 56.28 = one AA Rental expense with a credit line. MFD192 23,181.91 =
  "99 - Retainage" lines counted twice (the retainage audit session's fix, waiting on the owner). MFD295: the legacy
  "QuickBooks" income was net billed with no date cut - fixed; the real gap is release #32575 (6,363.00) that
  QuickBooks books as income and the release rule leaves out. CP790 -5,401.82 = the known #33995 duplicate.
- 2026-10-02 · **P&L checks count the billing outside the draw table** - the redesign's draw checks (billed,
  retained, overhead) only summed the draw rows, so every job with untagged pre-period invoices, a release outside the
  draw windows or unplaced not-billed retainage read ✗ (CP585 / 672 / 765 / 803 / 861 / 961, MFD177 / 295).
  `_outside_draws` adds those groups, named in the check label ("+ X outside the draws"); `_forming_invoices` keeps
  pre-period untagged invoices off the Next Draw sheet (MFD295 listed nine 2024-25 invoices as next-draw income). CP +
  MFD regenerated: every P&L check ties on all 23 workbooks; the ✗ left are QuickBooks-vs-Transactions gaps only.
  Overviews rebuilt (MFD after a OneDrive "not a zip file" read on the first try).
- 2026-10-02 · **CP Overview: releases were taken out of billed twice** - `completed_pnl._totals` stopped adding
  retainage releases to billed (3c1fe63) but `cp_overview` still subtracted them when it balanced each job sheet, so
  CP585 / CP672 / CP861 read short by exactly their releases. Fixed in cp_overview (`want_b`); after the CP
  regeneration `one-offs/cp_overview_reconcile.py` = 0 problems (1,511 formulas, 91 links) and
  `one-offs/cp_overview_verify.py` ties billed and costs on all 17 active jobs, the only gap CP790's known double
  count removed on purpose.
- 2026-10-02 · **APPROVED, round 23 - landed** (owner: "make sure to put $ formatting where applicable. commit push";
  "update all cp project"). Every dollar format carries the $ (CURR_FMT, RET_FMT, OVER_FMT, pnl_style.MONEY, the "off
  by" / "after X expenses" / "takeoff X under ETC" texts); `safe_save` widens each $ column by 1.5 (`_widen_for_dollar`)
  and the draw sheets' figure columns start at 13. The corrections this review made, all kept: billed is GROSS
  (retainage included); ② P&L by account took the overhead on the full CONTRACT and overstated the loss - every to-date
  overhead (① Actual, ② by account, every draw) is now rate x billed, only the projection uses the contract; ② Net
  ties to the draws' overhead; Reconciliations compares QuickBooks and Transactions through the same date; the
  gutter pass shifts only references into sheets that move (CF formulas too). Pinned by tests/test_pnl_layout_rules.py;
  project-pnl/FLOW.md added; docs/ARCHITECTURE.md updated. Copy: round 23. CP regenerated; MFD waits.
- 2026-10-02 · **IN REVIEW, round 22** (owner: "move checks to reconciliations and combine that info together";
  "move back draw coverage to P&L and put back the individual draw sheets but keep the new format"; "divider between the
  P&L and draw coverage"). The Draws sheet is gone: DRAW COVERAGE sits on the P&L under ① / ② behind a divider (white
  space, thick navy rule from column A - the gutter pass now carries an empty thick rule into the gutter), each name
  linking to its own sheet. One sheet per draw again - Next Draw, Draw 10 ... Draw 1, at the END before Draw Data - each
  in the section format with "↑ Draw coverage" on its band (write_draw_coverage / build_sheets_draws / set_back_links).
  The P&L CHECKS moved into the Reconciliations table under its QuickBooks tie-out ("P&L ties to its support"), the
  status line covers both, ✓/✗ coloured alike; Reconciliations is built after the P&L now. cp_overview reads the same
  (10 draws, 791,905.60, 726,547.23, awaiting 52,171.76, 16 next lines, 10 invoices, 347 bills). Copy: round 22.
- 2026-10-02 · **IN REVIEW, round 21** (owner: "last fix, just do costs by cost type, and just put Costs"). Each
  draw section has ONE cost block, titled "COSTS (n bills)", grouped by cost type; the by-vendor cut is gone (CP and PM
  branches). Every section's COSTS total = its coverage Costs (11/11). Copy: round 21.
- 2026-10-02 · **IN REVIEW, round 20** (owner: "put the draw # first, put Paid or Unpaid on the end"; "remove the
  individual big lines vertical in after 10% overhead, the big box is covering the three columns"). Draws bands read
  "Draw 5   ·   04/01/26–04/30/26   ·   PAID" (cp_overview's band regex still matches; invoices read back per draw).
  Coverage: Net Profit's inner box removed - only the three group boxes. Copy: round 20.
- 2026-10-02 · **IN REVIEW, round 19** (owner: "it was better before ... gross one box, after one box, billed and
  costs one box"). Coverage: the round-18 per-column rules and the Gross Profit box are gone; ONE medium black box per
  group from its band to TOTAL - BILLED + COSTS, GROSS, AFTER OVERHEAD - with Net Profit's own box kept. Copy: round 19.
- 2026-10-02 · **IN REVIEW, round 18** (owner: "divider needs to be a bit longer"; "the lines around gross and net
  ... need to be more pronounced"). The divider runs A to one column past the coverage table (A:N on CP). Coverage:
  GROSS and AFTER-OVERHEAD groups framed in medium black, a thin black rule after Gross Billed, and Gross Profit
  boxed like Net Profit. Ties unchanged. Copy: round 18.
- 2026-10-02 · **IN REVIEW, round 17** (owner, on a screenshot: "make the divider start from 1 column"; "i clicked
  draw 5 and it put it there it should be on the top"; "you remove the vertical lines from the draw coverage on gross and
  net with oh"). The divider rule now runs from column A. Coverage links reach 120 VISIBLE rows past the band (collapsed
  detail rows take no height, so the fixed 60-row range was shorter than the window and Excel left the band low). The
  coverage table's vertical rules are back (after Gross Billed, Costs, Coverage %, between OH views). Ties unchanged.
  Copy: round 17.
- 2026-10-02 · **IN REVIEW, round 16** (owner, on screenshots: "there's still no real dividers for the draws"; "when
  i click a draw in coverage it puts the location so the draw row is in the bottom, it should take me to top of it";
  new tab order; the ① Net Profit row "make it like this to bring eyes here"). Draws: between sections an 18-px gap, a
  thick navy rule the section's width (B..J), an 18-px gap; each coverage link targets B<band>:B<band+60> so Excel
  scrolls the band to the TOP (a one-cell link only scrolls it into view, at the bottom). Tabs: P&L, Draws, Budget vs
  Actual, Transactions, By Account, Labor, Concrete, Cash Flow, POs, Reconciliations, Draw Data. ① Net Profit row:
  bright yellow across A:E, the Actual figure in a medium black box. Ties unchanged. Copy: round 16.
- 2026-10-02 · **IN REVIEW, round 15** (owner: "option 1, total cost with the note"; "draws put dividers, put costs
  together meaning no extra space, remove by cost code and make by vendor first and by cost type 2nd. the top button
  should live on row 1, remove row 2 that's extra space"). Budget vs Actual = Budget | Actual (total cost, tax and fuel
  in - it IS the P&L) | Variance | Used, the note "Actual = total cost, incl. 13,680.16 tax & fuel - the detail is on
  the Labor and Concrete sheets" (supersedes the 07-29 pre-tax comparison on this sheet). Draws: one "↑ Top" in the
  frozen title row (freeze A2), coverage from row 2, a navy divider row above every section, no blank rows inside a
  section, cuts BY VENDOR then BY COST TYPE (cost-code cut removed - Budget vs Actual is the by-code view). Rebased on
  dev 3c1fe63: ① actuals overhead on billed for MFD too (10% and 9%). Ties: P&L 5, every section (11), BvA; 0 cut.
  Copy: round 15.
- 2026-10-02 · **IN REVIEW, round 14 - ONE Draws sheet** (owner: "the draws consolidate into one sheet like it was
  originally. move the coverage to draws so it will be the table with hyperlink to the section. freeze pane and have
  arrow up with 'top' ... make it a P&L style with income, retained, cogs, expenses, gross and net profit for each
  draw with next draw first before the latest draw, new to old"). `build_sheet_draws`: DRAW COVERAGE on top (Next draw
  (forming) row = costs so far, then Draw 10 .. Draw 1, TOTAL (draws), % Billed cumulative from the oldest), every
  name a link to its section; one section per draw via build_sheet_one_draw(ws=..., start_row=...) - band with
  "↑ Top", P&L-style summary (Income gross, Retained by the GC, less COGS, less Expenses, Gross profit, less OH 10% of
  income, Net profit, Net %), invoices (skipped when none), the three cuts; columns fitted once; frozen under the
  title. The Next Draw sheet and the per-draw tabs are gone; the coverage left the P&L (its checks read
  'Draws'!...); Draw Data carries the next draw's bills. Gutter pass: `_shift_a1_aware` moves a reference only when
  the sheet it points at moves (the Draws sheet keeps its own gutter) - and a name clash (`shifted` reused by the
  print-area step) was caught on the way. cp_overview readers take the Draws sheet (coverage, section invoices,
  next-draw lines from Draw Data) and still read the old layout: identical reads on CP790 old vs new. Ties: every
  section's COGS + expenses = its coverage costs = its three cut TOTALs, income = coverage gross (11/11); P&L 5 ties.
  OPEN: Budget vs Actual on total cost only (owner asked my take). Copy: round 14.
- 2026-10-02 · **IN REVIEW, round 13** (owner: "yes, move ① actual overhead to billed too"; "budget vs actuals,
  reverse the grouping, do the costs first then the phases as the sub grouping"; "do the variations for the actual
  draw sheets and see which one scores the highest for visibility, professional and usability ... numbers are
  floating ... remove the cell colors from the top strips where the numbers are, not the headers"). ① Actual overhead
  = rate x billed (CP790 -79,190.56; Net Profit -65,821.83), projection stays on the contract, label "less: Overhead
  10% (contract | billed)"; ② now differs from ① only by the office expenses. BVA_LAYOUT costtype (Concrete, Labor ...
  then SL1 Slab ...). Draw sheets: scorecard `score_draw.py` (floating numbers, coloured number cells, grey on EMPTY
  cells, width, rows, cut, heavy); 6 variants; winner V5 = DRAW_SUMMARY block (the nine figures as one label | value
  column, number cells uncoloured, sign by font colour), DRAW_LEAN (no Cost type / Note columns on CP, group rows grey
  only on label + amount): width 200 -> 167, grey-on-empty 115-213 -> 31, nothing cut. Draws: summary costs = the
  three cut TOTALs = the P&L draw row, 10/10. Copy: round 13.
- 2026-10-02 · **IN REVIEW, round 12** (owner: "where are the horizontal lines??? i need to separate the text from
  the numbers on the top"; "P&L by account renamed to P&L by account (current)"). The row rules were Excel "hair"
  (near-invisible) - now thin grey (BFBFBF) under every row of ① and the ② groups, and the yellow input cells too; a
  thin rule parts labels from numbers down ① (right of the label column) and inside each ② group (left of its
  value column); banner "② P&L BY ACCOUNT (CURRENT)". Scores: 0 cut, 5 ties. Copy: round 12.
- 2026-10-02 · **IN REVIEW, round 11 - ② overhead is on BILLED** (owner: "why are you taking off the full oh% on the by
  account if it's based on what we billed so far? ... overstating the loss ... do P&L by account (current)"). ② is a
  to-date view, so − OH = rate x income billed (CP790 79,190.56, was 97,701.06 of the contract; NET -66,153.95, was
  -84,664.45) - the draw table's basis, and a new check ties the two. The note right of GROSS / NET names only the
  reasons that apply ("after 332.12 expenses; OH on billed, ① on contract"). Audit of every figure on the sheet for
  basis: the coverage "% Compl" was billed ÷ contract - renamed "% Billed" (% complete is costs ÷ ETC). OPEN: the ①
  ACTUAL column still charges 10% of the contract (09-03 ruling) - asked the owner. Copy: round 11.
- 2026-10-02 · **IN REVIEW, round 10** (owner, on screenshots: "move the note to the right ... right borders for all
  the tops, need to divide"). The "after 332.12 expenses" note is now a formula in the column right of the = NET group
  (22 wide), on the GROSS and NET rows; every ② group closes with a thin right border on all its rows, header
  included; the ① header cells are divided by a white rule. Scores unchanged: 0 cut, all ties. Copy: round 10.
- 2026-10-02 · **IN REVIEW, round 9 - draw sheets scored; ② notes the expense gap** (owner: "the draw sheets do that
  after too ... the bottom two don't expand the same way the first one does as default grouping"; "if gross doesn't
  match add a note that it's taking off expenses next to it" - so BOTH bases stay: ① stops at COGS, ② also takes off
  the office expenses, and a formula note "after 332.12 expenses" sits under ② GROSS and NET whenever they differ
  from ①). Draw sheets: the by-vendor and by-cost-code cuts open like the by-cost-type one (band, header, group
  totals, TOTAL; bills on [+]) instead of folding whole; titles BY COST TYPE / BY VENDOR / BY COST CODE (the long
  ones were cut); section frames thin (`DRAW_FRAME`, medium -> thin: heavy cells 66-132 -> 0 per sheet); the label
  column fits its longest group label (cap 52; long vendor names were cut when opened); invoice block "INVOICES THIS
  DRAW (n)" / "TOTAL". Every draw: the three cut TOTALs = the COSTS tile = the P&L draw-table Costs (10/10). Scores,
  folded and fully opened: P&L 0 cut, BvA 0 cut, all 10 draws 0 cut / 0 heavy. Copy: round 9.
- 2026-10-02 · **IN REVIEW, round 8 - true net on ②, Budget vs Actual scored** (owner: "add the OH% after net profit to
  show that final true net and rename that first net as Gross, now do the same scoring for budget vs actual"). ② ends
  = GROSS (income - COGS - expenses), Gross %, − OH 10% (of contract), = NET, Net %. CP790: Gross 13,036.61, Net
  -84,664.45 - NOTE ① Net Profit is -84,332.33 because ① subtracts COGS only and ② also the 332.12 office expenses
  (asked the owner which basis). Budget vs Actual scorecard (`score_bva.py`, folded AND fully opened): 5 variants;
  heavy borders 129 -> 7 (the TOTAL double rule), width 210 -> 181, rows on open 64 -> 25. Defaults now
  BVA_FRAME thin, BVA_DETAIL compact (Class, then one "#bill · date · description" column that runs free - beside a
  Class flag it was cut), BVA_FOLD groups (job types open, codes on [+]); office block heading short; group bands
  stop at Total cost; text headers left. Copy: round 8.
- 2026-10-02 · **IN REVIEW, round 7 - scored, not eyeballed** (owner: "too many thick boxes in the top ... keep redoing
  the whole thing and finding the best scorecard"). A scorecard over the built file (scratch `score.py`): every check
  green, heavy-border cells in the top, cut-off text (the whole P&L; bold counted wider), fill count, font sizes, top
  width. Four framings built and scored (`PL_TOP_STYLE`): A boxed 101 heavy cells; B/C/D thin 17 (the one divider);
  D (thin outline + a hairline under each row) won on row tracking and is the default. Then: Variance / % headers
  navy like the rest (fills 9 -> 7); ② columns 15 -> 11 and the coverage grid matched to them (top ~2,200 px -> fits
  a 1920 screen at 110%); group headers in the owner's words INCOME | − COGS | − EXPENSES | = NET (the long ones
  were cut); "AFTER 10% OVERHEAD" band; account names un-indented (an indent pushed two past their cell); checks box
  thin. CP, MFD 9%, no-projection and finished builds all score 0 cut / all ties. Copy: round 7.
- 2026-10-02 · **IN REVIEW, round 6 - page-by-page refine** (owner: "go through each page and see what can be fixed
  and refined"; mid-way: "by account ... flow horizontal better with an equals to the net, so income > cogs > expense
  > net"). P&L: ② P&L BY ACCOUNT reads INCOME | − COST OF GOODS SOLD | − OPERATING EXPENSES | = NET OPERATING INCOME
  (Net %, Gross Profit, GP% under it); block labels are forced to text (a label starting "=" had become a formula,
  #REF!). Every tab: the long instruction sentences cut to one short line (Transactions, By Account, Labor, Concrete,
  POs, Reconciliations, Cash Flow, Next Draw); `_tidy_text` turns every em dash into a hyphen before save, except the
  "PROJECT P&L — <job>" title three readers parse; Next Draw's pre-period note was year-first, now mm/dd/yyyy.
  Reconciliations compared QuickBooks THROUGH THE LAST DRAW END with Transactions TO TODAY, so every later bill read
  as an error (CP790 "off by 36,571.61"); both sides now stop at the same date (`_after_cut`), later lines listed
  under the table. CP790 left: income -5,401.82 (the known #33995 duplicate) and COGS 295.30 (no single line equals
  it - needs a QuickBooks look). Copy: round 6.
- 2026-10-02 · **IN REVIEW, round 5 - P&L by account beside ①, coverage under both** (owner: "merge 3 and 4, billed
  gross, retained, net billed, retained still owed can all be under income; 2 progress, put it in 1 without distorting
  what's there; move P&L by account to the right of 1 and make it horizontal ... then move the draw coverage underneath
  those blocks with a divider"). ① PROFIT & LOSS (A:E) gains a grey "Progress" divider row + Earned Revenue / Over-
  (Under) Billing / Cost to Complete in the ACTUAL column. ② P&L BY ACCOUNT (F:N) = three side-by-side groups, label
  over 2 merged columns + value: INCOME (billed gross, retainage held, net billed, released, still owed, QBO-fix
  note), COST OF GOODS SOLD (header carries the total, child accounts link to By Account), OPERATING EXPENSES + Gross
  Profit / GP% / Net Operating Income. The invoice-by-invoice list left the P&L (a row outline cannot fold one
  group without its neighbours) - it is on Transactions. A thick navy divider, then DRAW COVERAGE from column A,
  then the rulings, then CHECKS. Root-cause read PINNED -> vault tasks/2026-10-02_pnl-one-engine.md. Copy: round 5.
- 2026-10-02 · **IN REVIEW, round 4 - linear again, draw coverage on the right** (owner: "make projects and actuals
  next to each other and go back to original formatting of having it linear, with draw coverage to the right; remove
  2 from actual, now it's just one box; the yellow your input remove and rename that block ... and put the 1 there").
  Left, down A:E/A:B: ① PROFIT & LOSS (PROJECTION | ACTUAL | VARIANCE | ACTUAL % OF PROJ., one box, actuals on the
  contract/ETC rows), ② PROGRESS (Earned Revenue, Over/(Under) Billing, Cost to Complete), ③ BILLING & RETAINAGE,
  ④ P&L BY ACCOUNT, CHECKS last. Right, from G at the top: DRAW COVERAGE (`_write_coverage(x0=7)`, F a gap). Branches
  CP / MFD 9% / no projection / finished all build; CP Overview reads the same draws and projection. Copy: round 4.
- 2026-10-02 · **IN REVIEW, round 3 - one band across the top, coverage under it** (owner marked up round 2: delete
  the Contract / Billed and ETC / Costs rows - the actuals go on the Original Contract / Original ETC rows - "move it
  horizontal ... and put draw coverage below ... redo the whole P&L ... simplified, easy to read"). Band: A:E
  ① PROJECTION | ② ACTUAL | VARIANCE | ACTUAL % OF PROJ. (actual billed on the contract row, costs on the ETC row - the
  Revised rows when a job has COs; % column = % billed / % of ETC spent, red past 100%), F:H ③ PROGRESS (Earned
  Revenue, Over/(Under) Billing, Cost to Complete - the two %s are not repeated), I:K ④ BILLING & RETAINAGE (short
  labels). Then ⑤ DRAW COVERAGE, ⑥ P&L BY ACCOUNT, CHECKS. Subtitle no longer says "P&L through <last draw end>"
  (the sheet runs to today). Fixed: `read_back_inputs` looked for its labels in column A only, so since the 08/31
  gutter a typed contract / ETC was NOT carried into the next run (and the ETC label had drifted); it reads A or B now.
  Not changed (other readers parse them): the em dash in the "PROJECT P&L — ..." title, the year-first "Generated"
  stamp. Review copy "Project_PnL_CP790 - round 3.xlsx".
- 2026-10-02 · **IN REVIEW, round 2 - the P&L is four blocks then coverage** (owner on round 1: "i dont see any changes
  between the three. i dont want to see checks there put it in the bottom. make a projections and actuals side by
  side with variance the other 2 below that and then the coverage below those 4 blocks"). Top: one table, label ·
  ① PROJECTION · ② ACTUAL - QBO TO DATE · VARIANCE (actual - projection) · ACTUAL AS % OF PROJECTION, rows Contract /
  Billed to Date, ETC / Costs to Date, Gross Profit (+%), less Overhead (10% of contract both sides), Net Profit (+%),
  MFD's 9% rows when MFD; the yellow inputs (Original Contract, Original ETC, COs) sit above it in the projection
  column. Then ③ PROGRESS (A:B) beside ④ BILLING & RETAINAGE (C:E), rulings, ⑤ DRAW COVERAGE across the sheet from
  column A (was top right from D), ⑥ P&L BY ACCOUNT, and the CHECKS at the very bottom (no ✓ marks up top).
  `build_sheet_pl` returns the cells the checks need; `_wire_pl_support` writes them. `cp_overview.read_coverage`
  looks down the sheet for the table (verified: same draws / totals / projection off old and new layouts). Branches
  run: CP, MFD 9% view, no projection on file, finished job (`simple`). Review copy: "Project_PnL_CP790 - round 2.xlsx".
- 2026-10-02 · **IN REVIEW (branch wt/pnl-redesign, not landed) - P&L + Budget vs Actual tie to each other.** The
  owner: the workbook is "disconnected and untrustworthy"; scope (relayed 10-02): the P&L and Budget vs Actual only,
  draw sheets untouched, CP only, MFD waits. Built so far: `pnl_style.py` (one bar / band / grid / money / tie-out
  look for rebuilt sheets); Budget vs Actual rebuilt for CP/MFD (RP keeps the old one as
  `build_sheet_budget_vs_actual_rp`): opens with "② Costs to Date - P&L vs this sheet - ✓ ties" and "① ETC vs takeoff
  budget"; job type (or cost type) > code > bill, bills folded, one row per bill per code; Actual is pre-tax against
  the pre-tax takeoff, Tax & fuel its own column, Total cost = the P&L to the cent; office accounts in their own block
  under the total; a code the takeoff never priced says "no budget". The P&L gets ✓ marks beside Billed / Costs to
  Date with the labels linking down, and/or a CHECKS box under the draw table (BvA = Costs to Date; draw gross =
  Billed; draws + next draw + pre-Draw-1 = costs + office; draw retained = retainage receivable) - `_wire_pl_support`,
  run after the P&L, finds its rows by label.
  Review copies: Analysis/P&L layout review (10-02-2026)/Project_PnL_CP790 - round 1 A/B/C.xlsx (A = job type +
  inline ✓; B = cost type + checks box; C = job type with a column per draw tying to the draw table + both).
  Built from one captured QBO pull (scratch replay), all tie-outs green, assert_clean green. Pending: the owner's
  picks, then drop the unpicked layouts, a test, ARCHITECTURE/FLOW, land, `active cp`.
- 2026-10-02 · **CP actuals charge overhead on billed to date, projections on the contract - never mixed** (owner:
  "make sure you are getting the oh% of the actual based on the current billing and not overstating the oh amount
  with using the contract oh total. we still need that in the projections"). ② Actuals on the CP P&L:
  "less: Overhead (10% of billed to date)" = rate x Billed to Date (was rate x contract - CP790 97,701 vs 79,191);
  ① Projection keeps rate x contract; the draw table was already per-draw billed. The CP Overview / Completed were
  already split (projection OH on contract, actual OH on billed).
  Same day, MFD + RP follow (owner: "yes the actual should also follow to mfd and rp, it only makes sense"): MFD
  P&L actuals (10% and 9%), the RP "True Net Profit", and the MFD / RP division Overviews (`completed_pnl._totals`,
  formulas, overhead tiles) take overhead on billed. `_totals` also stops adding retainage RELEASES to billed (they
  collect retainage already in gross) except past the retainage booked - MFD177 now ties to QBO income to the cent
  (was +63,739.33). The director's cut page follows (owner: "director's cut page fix ... don't confuse actual
  with projected vs completed oh"): overhead on billed, releases not counted twice, ACTIVE labelled "actuals to
  date, NOT final" and COMPLETED "final". MFD real net 10% moves -52,291 -> -127,081 before his cut (the cut paid
  is unchanged, 1,065,072.54 on these jobs).
- 2026-10-02 · **Retainage release invoices carry their QBO balance and id** - they read PAID while QBO still had them
  open (CP582, CP689, CP786) and had no QBO link.
- 2026-10-02 · **Retainage releases are not billing; Net Profit header colour fixed** - Billed to Date = gross invoice
  lines + any not-billed retainage still unplaced; a release (positive retainage line) is no longer added on top (CP585
  read 139,293 on a 126,630 contract; CP672 +30,870, CP861 +65,761) - same rule as the CP Overview and the ledger
  (owner 2026-09-23). Billing & retainage now reads held back / Net Billed / released / still receivable; the
  Reconciliations income check leaves releases out too, as QBO does. `_apply_left_gutter` now shifts conditional-format
  FORMULAS as well as ranges - the draw table's Net Profit header was keyed to the overhead column (always green).
- 2026-10-02 · **CP Overview = projections vs actuals, a draw-coverage sheet per job** (owner: "a cp overview to review
  projects projections and their actuals with the draw coverage in separate sheets ... simple, good colour"; then "the
  project sheets ... just give us the raw data with color well put"). New `cp_overview.py`, routed from
  `completed_pnl._build` for CP only. Overview: Projection (blue) / Actual (green) / Progress (amber) column groups, a
  KPI strip off the ALL row, coverage traffic-lit (under 100% red, to 111% amber = covers cost not 10% OH, past it
  green), % complete red past 100%. Job sheet: no cards or notes - the coverage table (each draw links to its lines)
  and the transactions by draw, newest first; bills outside the draw windows are matched by date (the draw sheets join
  split bill #s), and a "difference to the P&L total" line appears only if the lines do not tie (none on the 17
  jobs built 10/02). MFD / RP Overviews unchanged.
  Same day (owner: "group the project vendor costs by vendor. remove the big kpi strip ... and you forgot net with
  oh %"): KPI strip gone (the ALL row carries the totals); bills inside each draw grouped under a vendor subtotal
  row; NET added with the rate in every label - projected net = contract - ETC - 10% of CONTRACT, net to date =
  billed - costs - 10% of GROSS BILLED (the per-draw overhead of the P&L's draw table, 10/01), and each draw row on
  the job sheet carries its own 10% OH and net. Net vs plan compares net % to projected net %.
  Round 3 (owner: gross % + a divider between gross and net, $ formats, outline level 2 by default, a thick black box
  per draw, "verify cp790 ... not reconciled right with the retainage"): projected gross % added; a navy rule splits
  gross from net in both groups; money is `$` with a minus sign on negatives (Excel's accessibility check flagged
  colour-only negatives and the light grey text - grey is now 4B5563); bill lines start folded (vendor totals show);
  each draw sits in a thick black box whose bottom edge is the spacer row's top, so it closes when folded.
  **Retainage not billed is spread, not parked** (`_spread_retainage`): the "Retainage not Billed" invoice / its
  journal-entry retainage goes back onto the early NET-entered draws at net / 9; a remainder equal to a later draw's
  own retainage line is a QBO double count and is NOT counted. CP790: 22,720.10 -> Draw 1 3,243.38 + Draw 2
  14,074.90, 5,401.82 = Draw 3's retainage (not counted); billed 791,905.60, retainage to date 79,190.56 = 10% of
  it exactly. The P&L workbook itself still carries 797,307.42 - a project_pnl_export fix, not this file's.
  Round 4 (owner: "pull info from qbo to update these projects", "the completed needs a full run", "show me active
  excel and work on completed excel separately so it's two excels", "verify ... each number"): every CP P&L
  regenerated from QBO 10/02; `CP Overview.xlsx` = ACTIVE jobs only, finished jobs get their own workbook after a
  full run confirms they are done (their P&Ls fall back to the 12-31-25 WIP report for contract / ETC - unconfirmed,
  so projections are blank on finished rows). A **retainage RELEASE is not billing** (owner 09/23 rule): the release
  invoice is listed at $0 with its amount in the description - CP585 / CP672 / CP861 were over by 12,663 / 30,870 /
  65,760.60. Retainage-not-billed older than the draw table stays its own line, tested against the net invoices / 9
  (CP861 ties to its pay apps instead: 38,940 + 2,200). QBO data stamp on the title line of every sheet. Verified by
  `one-offs/cp_overview_verify.py`: billed and costs re-derived from QBO (strict: lines CODED to the project) tie on
  all 17 active jobs; the only gap is CP790's 5,401.82 double count, removed on purpose.
  TO DO: the Completed workbook (CP582, CP656, CP689, CP697, CP714, CP742, CP786 regenerated; CP510 is not a QBO
  customer); CP783 / CP885 have no costs yet so are not on the Overview.
  Round 5 - **the owner's own layout, copied from the file he edited on Common**: no ACTIVE band (jobs start under the
  headers), one total row ("Subtotal · N job(s)", no ALL row when there is one group), the coverage legend under
  PROGRESS & COVERAGE on two rows, freeze at C6 (headers + job column), ETC / projected GP / costs / GP / % complete /
  coverage / net vs plan at width 13. Values matched his file cell for cell before the live copy was rewritten.
  Round 6 - **CP Completed workbook** (`cp_completed.py`, owner: "work on completed excel separately" / "be wary of
  projects not being done"): `Completed Projects/CP Completed <year>.xlsx`, the finished jobs of the year with no
  projection (contract / ETC on a finished job came from an old WIP report). Each job is sorted by what QBO still
  has open: PAID IN FULL · RETAINAGE STILL OWED (only retainage-release invoices open) · CHECK - MAY NOT BE DONE
  (a short-paid regular draw, or costs > 30 days after the last invoice). The finished jobs' P&Ls are built
  `--legacy` - before project coding, their costs sit on lines that only NAME the job (CP582 +138K, CP697 +126K,
  CP714 +118K, CP786 +66K, CP689 +38K) and CP714's Draw #1 was invoiced on the GC. A retainage release counts as
  billing only past the retainage already booked (CP697 booked none: its 21,674.27 release is its only record of
  that billing; CP786 released 258.95 more than booked). Retainage-not-billed is cleared as a possible double count
  when no invoice carries a retainage line, or the GC released everything booked. `cp_overview_verify.py --legacy`
  checks the book against QBO with the same attribution. CP510 is not a QBO project customer (named
  "CP510-GRACE CHAPEL - PROSPER") - not in the book. CP Overview build takes a column set (`COMPLETED_COLS`);
  `read_coverage` searches rows 1-200 (the new P&L layout moves the draw table down).
  Checks (owner: "verify the overviews reconcile and are correct", 10/02): `one-offs/cp_overview_verify.py` ties
  billed / costs to QBO (refresh the mirror first - `ledger/refresh_mirror.py`; on 10/02 three jobs had bills
  entered after the morning pull), `one-offs/cp_overview_reconcile.py` proves the workbook adds up inside
  (1,402 + 676 formulas, 0 problems). CP831's contract carries a 33,361 CO from the Draw #2 G702 that the WIP
  master does not have yet.
- 2026-10-02 · **Layout round 2 (the owner's 7 items on the CP790 review copy)** - ④ shows child account names only
  (full name kept when two parents share one); the P&L label column is sized to its widest visible label;
  draw-coverage Net Profit in a thick black box with a green/red header off the TOTAL; Transactions section bars
  merged + centered and long labels merged instead of wrapped (`_merge_label_rows`); Budget vs Actual headers
  Cost code / Budget / Actuals / Variance / Used % with Budget and Actuals each boxed; Draw Data is the last tab;
  draw sheets boxed B..J with row rules, bill # under its vendor, amount beside each group, TOTAL row per costs
  section, no "in this draw" filler on CP, tighter tiles. The not-billed retainage header now says it is placed back
  on the draws (it is in income).
- 2026-10-02 · **Layout pass (the owner's picks from four mocked variations)** - P&L sheet: sections number themselves
  and are merged + centered; ① Projections / WIP, ② Actuals - QBO to date (its own GREEN bar), ③ Progress,
  ④ P&L by account, ⑤ Billing & retainage; the ③ Snapshot is gone (MFD keeps its 9% net profit inside Actuals). Draw
  coverage: Net Billed | Retained | Gross Billed | Costs, coloured group bands, Gross Profit / Net Profit highlighted on
  the TOTAL row only. Labor/Concrete ledger: no VENDOR column (the vendor is the group row; the PM-mark read-back takes
  it from there), AMOUNT under ACTUAL, bare cost codes. Draw pages: tiles coloured by section, green invoices band,
  orange costs band with light grey sub-headers, by cost type open and the other cuts folded. Label wrap now
  estimates rendered width (`_text_units`) - the blank line under Billed to Date is gone.
- 2026-10-02 · **"Retainage not billed" invoices go back on their draws** (`_spread_not_billed_retainage`): placed
  oldest-first on the draws entered net, at net/9 each; a leftover equal to a retainage line already on another
  invoice is a duplicate - flagged on the Transactions sheet and in ⑤ as "⚑ QuickBooks fix pending", never counted;
  any other leftover still counts. P&L and draw table now agree. `tests/test_pnl_retainage_spread.py`.
- 2026-10-01 · **Draw coverage runs on GROSS billed (CP + MFD)** - the P&L sheet's draw table now shows Gross Billed,
  Retained and Net Billed side by side (the retainage columns were collapsed and hidden), and Gross Profit, Net Profit,
  Coverage %, Net Cov % and % Compl all run off gross. Overhead per draw = rate x the draw's gross billed (was
  contract x costs / ETC, which overcharged a job running past its ETC) - the ledger's draw basis since 09/23.
- 2026-09-30 · **No false "no draw #" warning on month-named draws** - a memo like "September Draw 2026 (Period: …)"
  already names the sheet by its month and the period drives the window; the yellow warning now fires only when the memo
  names neither a draw number nor a month.
- 2026-09-30 · **CP Overview reads year-filed finished jobs** - `Completed Projects/<year>/<job>/Profit and Loss/`
  (CP610 is filed under 2025); the walk descends into 4-digit year folders under an archive folder.
- 2026-09-30 · **CP Overview lives in the Active Awarded Projects folder on Common** (owner: "cp overview should live in the
  active awarded projects folder in common", after `✗ CP Overview not rebuilt: CP has no home mapped beyond the retired
  Automations- folder`). `shared/pnl_paths.overview_dir()` - CP = `/Volumes/Common/CURRENT PROJECTS/Awarded Projects
  Commercial projects` (top level), not mounted = HomeNotMounted unless `--to-automations`; MFD / RP unchanged.
  `completed_pnl.rebuild_overview` / the CLI write there; the job links are relative (`<job>/Profit and Loss/...`).
- 2026-09-30 · **Finished CP jobs find their folder, and every P&L carries its bill scans** (owner on CP610: "doesn't have the new features of downloading the attachments and letting me open the folder"). (1) `shared/pnl_paths._find_awarded_cp_folder` searched only the top of the awarded tree; finished jobs sit in `Completed Projects/` and `Completed Projects/<year>/`, so they lost their home, pay app (contract) and takeoff. It now falls back there - active folders first, and a FULL job-# match only in the archive (`CP592` never picks `CP5921`); `tests/test_cp_folder_lookup.py`. (2) Scans used to be downloaded only for the Labor/Concrete sheets, which need a coded takeoff budget - old jobs have none, so they had no scans at all. The CP/MFD template now downloads every Transactions bill/expense scan into `attachments/` beside the workbook, adds a **Scan** column (`Open scan`, or `Open folder (N files)` for a multi-scan bill) next to the QBO-linked Ref #, and an **Open scans folder** link at the top of the sheet. Same stored relative links as the Labor sheets. RP not wired yet (its folder is chosen after the sheets are built). CP610 rebuilt: 77 of 82 txns had scans, totals unchanged.
  Open: CP610 itself is not a QBO project (costs on its deleted class, invoices on the GC) - it was built by supplying the GC customer by hand; a flag for that shape is still to do.
- 2026-09-28 · **CompanyHealth is organized, not a dump** (owner: "a more organized automated folder system rather than a dump"). Every read/write there goes through a named folder in `shared/paths.py`: `register_file()` -> `Registers/` (the JSON rules), `reports_dir()` -> `Reports/` (every generated report), `analysis_dir(topic)` -> `Analysis/<topic> (mm-dd-yyyy)/` (session one-offs). `tests/test_companyhealth_layout.py` fails any module that joins a file onto `companyhealth_dir()`; `python3 shared/paths.py --organize [--apply]` sorts a root that filled up anyway. This tool: job rulings / draw moves / biz-dev cut read from `Registers/`; the director's cut workbook writes to `Reports/`.
- **Runs on python-env (2026-09-24).** `run_pnl.sh` now starts through `python-env/python.sh` (`"$ACB_PY"`, Python 3.14, every package pinned), never a bare `python3` - a `brew install ffmpeg` on 09/23 swapped `python3` and broke every sync step. Missing-package hints now say `bash python-env/setup.sh`. Verified by a dry run on 3.14. See `python-env/STATUS.md`.

- 2026-09-23 · **A `costs` class ruling switches the class/project lookup on by itself** - the job's costs = its QBO project + every line on its OWN class with no project, via a `costs` ruling
  (`kind: costs, rule: class`) in the job rulings register (owner 2026-09-23 on MFD295: "why aren't you combining all
  the real costs?"). MFD295: most of its cost sits on `Elite Construction:MFD295` with no project (129 lines, Dec 2024 -
  Jul 2025, entered before project coding) - none of those lines carries any project, so nothing counts twice.
  So an `active mfd` batch rebuild of MFD295 can never fall back to project-only (the 09/11 FINAL was built with the lookup by hand).

- **Overview reader follows the combined draws (2026-09-16).** `completed_pnl._read_invoices`
  read every row of the Transactions INCOME list, so a job with a `draws` ruling (MFD192) would
  have counted each draw twice on the next rebuild - the bold draw line plus the invoices folded
  under it (the reader returned double the job's billed). It now treats a level-0 row followed
  by outline-level-1 rows as the draw head: skips it, reads the invoices, and tags each with its
  draw. The MFD192 job sheet in `MFD Overview.xlsx` shows the same shape as the P&L - one bold
  line per draw (its total, invoice count, PAID only when every invoice is) with the invoices
  indented under it, each still linked; every other job renders as before. MFD Overview rebuilt
  from the worktree: MFD192 billed ties to its workbook, `assert_clean` green.

- **Draw sheets reworked on the owner's MFD192 review (2026-09-16).** Four asks, one
  sheet: (1) **the KPI tiles are one cell each** - no more merged pairs; nine tiles sit on
  B..J, the same columns the bills table uses, so nothing merges and any column selects.
  (2) **INVOICES THIS DRAW sits under the strip, above the costs** - one linked row per
  QBO invoice (MFD192's draws are 2-3), Gross / Retainage withheld / Retainage billed /
  Net (cash) / Paid? / Memo, with a TOTAL row; the INCOME, RETAINAGE and OVERHEAD tiles
  are now FORMULAS off that row, so the strip traces to the invoices on its face.
  Caught in the process: INCOME had been the net-of-retainage total (TotalAmt), so the
  strip held retainage back a second time and NET DRAW read low by exactly the
  retainage on every draw; `rev` is now gross + retainage billed = net + withheld, and
  gross profit / overhead ride the gross the way the P&L sheet does. (3) **Cost code and Cost type columns on every
  bill row, and three grouped cuts** of the same bills: by vendor; by cost code, then
  vendor; by cost type (Concrete / Materials / Labor, then the cost family, then vendor
  - a family that only repeats its category, Concrete or Labor, gets no extra row).
  Each cut is an outline collapsed to its top rows; openpyxl cannot author a PivotTable,
  so the cuts are written out (the Draw Data sheet stays the flat pivot source). The
  MFD reconciliation blocks (MATCHED / QBO ONLY / PM ONLY) keep their vendor grouping
  and gain the two extra cuts over all QBO bills of the draw. One writer, `detail()`,
  takes the level spec; `_bill_row` is the one bill line. (4) **Auto-fit at close-out**
  (the owner: "look how much space bill # has ... remember to auto size when closing
  out"): `_autofit` over B..I measuring only the table rows, floor 16 for the tiles,
  cap 40; Description (J) stays at 16 and spills. Bill # went from a fixed 30 to what
  its ids need. Bill and invoice numbers that are all digits are written as real
  numbers (`idc`, format `0`, left-aligned with an indent) - no more green
  "number stored as text" triangles; `_autofit` measures such ids as plain digits. An
  account-based line shows a blank Cost code and its account name as the cost type
  (it was landing the account name in the Cost code column, 29 wide). Verified on a
  scratch build of MFD192 (the live workbook was open in Excel, so it was not
  rewritten): September draw = 2 invoices totalling to the tiles, 106 bills in each of
  the three cuts with equal totals, no merged cells but the title, no text ids,
  `assert_clean` green. Print area A..J, freeze under the strip unchanged.

- **ONE income line per draw for a job whose standing `draws` ruling says its invoices
  combine (2026-09-16).** MFD192 bills three contracts (main / HUDSONWOOD / OFFSITE) as
  2-3 invoices dated together every month; the draw sheets already summed them, but the
  P&L sheet's Income detail and the Transactions INCOME list ran one line per invoice
  (20 lines for 8 draws). The owner: "combine the income into one income only for this
  project" - so it is gated by the job rulings register, kind `draws` (`combine: month`,
  `shared/job_rulings.draw_combine`), never by a generic same-date rule.
  `gather_transactions(combine_draws=, draw_names=)` emits one record per draw group with
  the invoices riding along as `docs`; the Transactions sheet writes the draw as a bold
  line (Inv # = the invoice numbers, memo = the draw name, summed columns, Paid? from the
  combined balance) with its invoices collapsed under it on the [+], each still a QBO
  link, and TOTAL INCOME sums the draw lines only (`=SUM(E7,E10,...)`); the P&L sheet's
  Income detail writes the draw at outline level 1 and its invoices at level 2. A
  one-invoice draw, the retainage-billed and untagged buckets, and every job without the
  ruling render exactly as before. `_write_job_rulings` reads `job_rulings.findings()`
  (new - every kind but `draws`), so a billing convention never prints in KNOWN LOSSES /
  RULINGS. **Draw names now come from the memo's own month**
  (`shared/draws.draw_month_from_memo`); the period END month only when no invoice names
  one. MFD192's January window is its "February Draw", so the old rule filed it under
  "Draw – January 2026" and left no May sheet; the sheets now run February..September.
  Jobs with "Draw #N" memos, or whose memo month equals the period-end month, are
  unchanged. Verified on the regenerated MFD192 workbook: 22 invoices (2 voided dropped)
  -> 8 draw lines, Transactions TOTAL INCOME ties to the per-invoice sum, the QBO P&L
  through 09/01 ties once the 09/07 invoice is added back, `assert_clean` passes. The
  ledger was deliberately not changed (the owner: "leave the ledger alone").

- **Draw period tag parser moved to `shared/draws` (2026-09-15).** `DRAW_PERIOD_RE` is now an
  alias of `shared.draws.PERIOD_TAG_RE`, which accepts the tag with OR without parentheses
  (`Draw #1 - Period:05/21/2026 - 06/20/2026` as typed on the CP785/CP997 draws). Draw grouping
  and the memo-strip on the invoice lines behave the same; the lenient mistyped-date fallback is
  unchanged. Same fix as the bill-tracker's (a CP785 bill showed Awaiting Invoice).

- **NO-CUT JOBS SIT APART, LINES GROUP BY CATEGORY, NO BOX ON THE STRIP (2026-09-11).** The
  owner: a job he was never paid on "we can't judge due to not knowing his cost, so it should be
  as a pure job performance" - those jobs (MFD281, MFD295, MFD183 today) sit in their own section
  under COMPLETED, out of the judged `ALL <DIV> WITH HIS CUT` total the strip reads. Each job
  sheet groups its lines by category (header, lines newest first, subtotal) instead of a flat
  list; the sheet's DIRECTOR CUT total is the SUM of the three exact-label GOES TO splits, never
  a wildcard SUMIF (the `formulas` engine counted blank cells as "?*" matches and doubled every
  job - caught by the check cells before shipping). The heavy box around the KPI strip is gone;
  a hairline separates the two overhead views.

- **CATEGORIES OF THE CUT, AND WHO IT GOES TO (2026-09-11).** The owner: "how do you suggest we
  create categories of his cut? that way when we open the MFD report we see how it's really
  getting split off, either to him or to his junior estimator." The register gained
  `categories` (name, who, pattern, memo) and `who_labels`; `shared/bizdev_cut.categorize`
  is the one test (first match on the line text, then memo-only rules; a fronted line is job
  cost before any category; no rule = Other, blank = Blank line, both "ask the director" so
  nothing files silently). The PAGE stays combined - one DIRECTOR CUT column (the owner: "in the actual
  project sheets, not the overview ... once he wants to see it we pull the details"); each
  JOB SHEET lists CATEGORY and GOES TO per line and closes with the split (his / junior
  estimator / to confirm) and a by-category block, all SUMIF formulas over the lines. Burden, Sub
  service and Hourly help are deliberately NOT tied to anyone - the owner is asking the
  director. Validation: 2025 draws tie to the director's own billing tally within 174 on
  185,625 and estimating within 120 on 11,260; the residual bucket is 750 and was read line
  by line. Petty cash, reimbursements, notary and registration fees joined `fronted`.

- **THE AUTOMATIONS- TREE IS RETIRED - A MISSING HOME STOPS THE RUN (2026-09-11).** The
  owner: "Automations folder I'm retiring, only should be used if no mount exists and i
  deliberately say bypass by sending to automations." `shared/pnl_paths` now refuses to
  resolve a P&L, Overview or cut page INTO the OneDrive `Automations-` tree
  (`HomeNotMounted`, message names the reason and the rule) unless the run carries
  **`--to-automations`** (project_pnl_export, completed_pnl) - an explicit `--out` still
  wins, an `ACB_PNL_DIR_<DIV>` override still wins, and an `ACB_PNL_OUT_DIR` outside that
  tree is untouched. The real home is tried FIRST (Teams channel for MFD, the awarded /
  address folder on the Common drive for CP / RP) and the division folder only consulted
  when it is missing. READS are unaffected: `find_pnl`, `_candidates`, `_archive_dirs`
  pass `bypass=True` so old files there are still found. `bizdev_cut_view` prints
  "stopped" and exits 1 rather than building the owner's page from stale copies.
  **P&L ONLY** (the owner, same day: "i said only for the P&L never said everything else") -
  Bill Tracker, Open_Invoices and the statement reconciles keep their Automations- homes.

- **`bizdev_cut_view.py` - THE OWNER'S INTERNAL P&L WITH THE DIRECTOR'S CUT AT THE END
  (2026-09-11, reworked the same day).** The owner: "take mfd overview and make sure the
  director cut looks like that ... add both 9 and 10% ... add the director's cut at the end
  to see the final net profit. then have it hyper link to those transactions." Then: "remove
  any OH calculations and make it a separate excel ... put jordan's cut in the costs of the
  project instead, we want straight director only." One page in the Overview's exact shape
  (same reader, `completed_pnl.read_source` + `_totals`, so it cannot disagree with it), then
  ONE cut column - the register's `director` - and REAL NET at both rates. Every other
  registered vendor's cut is charged into that job's COST (the COST cell is `<workbook cost>
  + '<job>'!<in-job-cost total>`), so it is visible on the job sheet but never a column.
  Every cut cell is a formula pointing at the job's own sheet, where each line carries its
  QBO deep link, the account, and how it was tied to the job (project on the line, draw code,
  bill memo, split N ways). Below the table only a reconciliation: every dollar paid to the
  director ties to the page (14 jobs + older jobs + no job + fronted = total paid, check = 0).
  No rate ladder, no overhead model - the overhead review lives in its own local workbook
  (`<DIV> OH Calculations.xlsx`). The cut comes from ONE pull per vendor - the TransactionList
  report by vendor names the checks and expenses a Purchase query cannot find. Writes
  `<CompanyHealth>/Reports/<DIV> PnL - Internal - Director Cut.xlsx` from scratch, local only, never
  on the share. Verified with the `formulas` engine (zero errors) and `assert_clean`. The
  register gained `labels` (column headers) and `director` (`bizdev_cut.is_director`);
  shared reports still use `note()` and name nobody.

- **A RUN PULLS QBO ONCE, NOT ONCE PER JOB (2026-09-10).** 13 regenerations took
  39 minutes because `generate_project_pnl` runs per project and every call
  re-downloaded the whole Bill and Purchase universe, plus Account and Item.
  Batching into one process saved nothing. Each job narrows the pull to its OWN
  activity window, so a cache keyed on the query never hits - `_txn_pull` caches a
  SUPERSET instead: the first job pulls its window, a job inside that window is
  served from memory, and a wider window re-pulls the union ONCE and replaces the
  cache. Peak memory is one entity list for the widest window, no worse than a single
  job costs today. `_list_pull` does Account and Item once per run (4 and 3 call sites).
  `_reset_txn_cache()` exists for tests only. **A batch shares the cache only within
  one process** - the per-job flag recipe still forces separate invocations for
  MFD172, MFD228 and MFD295.

- **`_wrap_long_labels` HAS THREE GUARDS NOW - it shipped broken (2026-09-10).**
  It wrapped a 50-character title inside the ~3-wide GUTTER column, which renders one
  letter per line down forty rows, on every sheet of all 14 workbooks. Three causes,
  three guards: `min_width` (draw sheets already carry their own narrow first column
  so `_apply_left_gutter` skips them and the pass met the real gutter head-on), row 1
  is never wrapped (`_apply_left_gutter` deliberately hangs the title back into the
  gutter afterwards and carried the wrap with it), and merged cells are never wrapped
  (the text already spans columns, so width was never the constraint).
  **It got through because the test checked the four properties that had been changed
  instead of opening the workbook** - see the standing rule that checking a property
  is not checking the file.

- **TRANSACTIONS READS AT 14, NOTHING GOES OVER 16 (2026-09-10).** The owner:
  "default pnl to font size 14 for transactions. don't let it go over 16 anywhere."
  `BODY_SIZE_BY_SHEET = {"Transactions": 14}` in `_normalise_body_font`, which
  already scales that sheet's column widths by the same ratio - a size bump without
  the widths clips the text. Every other sheet stays at 12. `MAX_FONT_PT = 16` is a
  CLAMP that runs last over every cell, so a call site cannot reintroduce an 18 or 20
  point banner later. Verified offline against the delivered MFD177 workbook (8,489
  Transactions cells 12 -> 14, widths scaled, P&L and By Account byte-identical at
  12, nothing above 16) before spending a QBO run on it.

- **MFD IS 9%, AND THE P&L BLOCK READS TOP TO BOTTOM (2026-09-10).** The owner:
  "do 9% moving forward for MFD and remove 10% everywhere. let's just use 9%."
  An MFD workbook used to carry BOTH a 9% MFD view and a 10% company view, plus a
  `SNAPSHOT - MFD vs COMPANY` block, so every figure had two answers and the reader
  picked one. `_alt_oh` is gone: MFD sets `overhead_pct = 9.0` and there is no second
  view on any division. **CP and RP still run the 10% company rate** - this moved MFD
  only, confirmed with the owner.
  ACTUALS now reads **Billed -> Costs -> Gross Profit -> Gross Profit % -> less
  Overhead -> Net Profit -> Net Profit %**: the percentage sits with the number it
  belongs to. "REAL" is gone from the P&L rows and the draw KPI cards - with one rate
  there is nothing for it to distinguish.

- **ONE ZOOM, ONE LABEL WIDTH - both post-passes in `safe_save` (2026-09-10).**
  `_apply_zoom` sets 110 on every sheet. Set per call-site it drifted: Transactions,
  By Account, Reconciliations and Draw Data never set one at all (Excel opened them
  at whatever it liked) and a draw tab had wandered to 138. `_wrap_long_labels` wraps
  a label rather than widening the column for it, and runs BEFORE `_apply_left_gutter`
  while the label column is still A. A column has ONE width, so the longest string in
  it sets the width for every row below - the P&L label column was 52 wide because of
  a few long COGS account names while the block the owner reads tops out at 36
  ("do they really have to be this wide? does it mess with the below cells?").
  The column is now 40 in code (43.6 delivered) and 11 labels wrap instead. Rows with
  an explicit height are skipped so nothing is squashed.

- **A BUSINESS-DEVELOPMENT CUT IS NOT A JOB COST (2026-09-09).** Some outside
  parties are paid a cut of the work they bring in and invoice it against
  whichever job is open, writing the job number on the line themselves in their
  own coding. Those lines were landing in job cost, so the jobs that happened to
  be open read as losers and the jobs that were not read as winners - and which
  jobs got hit depended on the FLAG the run used (`--legacy` reads line text and
  picks them up, `--class-project` does not), not on anything real. The owner:
  "is it a job cost or not. it's not it's overhead."
  `shared/bizdev_cut.is_cut(vendor, account, text)` is the ONE test, and
  `gather_transactions.take()` diverts a matching line before it can reach
  `cogs`/`exp`, so the Job P&L totals AND the division Overview built from those
  workbooks are both clean and nobody's pay is printed on a workbook they can
  open. The Job P&L carries one grey line under Costs to Date - the amount
  excluded and nothing else, no vendor, no breakdown. The test is one question,
  "did this line build the job?": their draw, commission, allowance, estimating
  fee, software and their own staff's hours are overhead even with a job number
  and a job cost code on the line; material, fuel, safety, rentals and forms/
  batters they laid out for the job stay in cost. **The ACCOUNT is not the test** -
  the same cut has been booked to a COGS account and an operating account on the
  same job, and real fronted material has been booked to the COGS one. The
  register (`<CompanyHealth>/Registers/bizdev_cut.json`, `ACB_BIZDEV_CUT_FILE`) lives
  outside the repo because it names vendors; **with no register `is_cut()` is
  always False and no report changes**, deliberately - a missing file must never
  silently move money. Self-test: `python3 shared/bizdev_cut.py`.

- **SCOPE-BASED RP JOBS GET A SHEET PER INVOICE (2026-09-08, RP6586 the test
  subject).** The owner: "we have two different types, a one invoice one job vs a
  scope based invoices for one job like this one. the P&L should have sheets like
  draws but with the invoice # and using the invoice date as the cutoff for costs."
  The kind is read off the job's own invoices by the ONE rule in
  `shared/rp_invoicing.classify` (the scan uses it too). What the 2,464
  invoices since 2025 taught: EVERY memo reads "RP#### - <address> - <what was
  billed>", so the third segment is the invoice's description, not a kind
  marker (a tract job says "Foundation" there), and a tract job is routinely
  billed AGAIN for extras (pump, rock saw, brickledge, HTT5 bolts, "... Extra")
  days later - still a one-invoice job. A STAGE invoice names a piece of the
  slab build (lot prep, piers, mud slab, grade beams, walls, foundation) and is
  not an extra; 2+ stage invoices = scope-based (75 jobs since 2025); one
  PARTIAL stage (piers, grade beams ...) = scope-based with one stage billed so
  far (31); one "Foundation" invoice, extras or not = one-invoice (770). The
  Job P&L subtitle states the kind. **One-invoice job:** unchanged - the invoice date
  ends the job, later non-wreck bills go to Pending Review. **Scope-based job:**
  no cutoff at all (the latest invoice is just the latest stage), every bill is
  a job cost, open POs count as committed; then ONE SHEET PER INVOICE
  (`Inv #34114`, draw-blue) holding the invoice (QBO link), date, scope, the
  stage's billed / costs / GP / margin, the running totals THROUGH that invoice
  (billed, costs, GP, margin, % of bid billed, bid left to bill), and the
  stage's costs account -> vendor -> bill - bills dated after the previous
  invoice and on/before this one (`_slice_rp_groups`, start exclusive, end
  inclusive). A final **`Next Invoice`** sheet (the CP Next Draw shade) holds
  what accumulated since the latest invoice. The Job P&L's invoice rows link
  to their stage sheets. The cost detail walk left `build_sheet_job_rp` for
  `_write_rp_cost_groups`, the ONE writer both the job sheet and the stage
  sheets use, so a stage reads exactly like the job. Sheets sit between Job
  P&L and Transactions. Companion scan: `one-offs/rp_stage_scan.py` lists every
  RP job's kind off QBO (count, scope-named memos, span, WIP TYPE) to
  `<CompanyHealth>/Reports/RP Invoicing Stages.xlsx`.

- **RP P&Ls LIVE IN THE JOB FOLDER ON THE COMMON DRIVE, LIKE CP (2026-09-08).** The
  owner, on seeing RP6586 land in OneDrive: "it should be in the current projects
  folder not the automation folder, common was mounted." An RP job folder is
  `<Residential>/<builder>/<address>/` and the address never carries the job #, so
  the only key is the takeoff inside (`RP####_<ADDRESS>.xlsm`). The takeoff walk
  moved out of this tool into `shared/pnl_paths.rp_takeoff_index` (one parallel
  scan, cached) and grew `rp_job_folder` / `rp_pnl_dir`: the workbook now goes to
  `<job folder>/Profit and Loss/<RP#### - Customer>.xlsx`; the OneDrive division
  folder is the fallback only (drive not mounted, no takeoff names the job, or an
  explicit `--out`), and the run says which. `resolve_project_out_dir` answers the
  same for RP, `_candidates` now also globs the RP name in both homes (so the
  ledger's "last pulled" finds RP workbooks - the open issue below is closed), and
  the RP Overview's `_iter_jobs` walks the share through the same index with the
  CP replace rule (the share's copy is the live one). A takeoff filed under an
  ARCHIVE/OLD/COPY folder never decides the home. RP6586 regenerated into its
  job folder's `Profit and Loss/`; the OneDrive copy from earlier the same day
  was deleted (one job, one P&L).

- **KNOWN LOSSES / RULINGS BLOCK (2026-09-08, first case RP6586).** The owner:
  "this needs to be a note for that job permanently so when we do the WIP
  report we don't flag ... we also need to put it in their P&L somewhere." A
  job-specific fact had nowhere durable to live - the vault refuses job notes
  by design, session memory dies, and the WIP NOTES column is regenerated -
  so every report re-discovered the same overrun. Now ONE register,
  `<CompanyHealth>/Registers/job_rulings.json` read through `shared/job_rulings.py`
  (same pattern as `draw_moves.json`), feeds every reader: this tool writes a
  **KNOWN LOSSES / RULINGS** box on the Job P&L (RP card: after WIP /
  PROJECTION, before INVOICE; CP/MFD P&L sheet: between ① WIP and ② TOTALS) -
  the why in col A (red when `kind` is `loss`), the $ the ruling concerns in
  the value column, the bid line + document + ruling date in grey under it.
  Nothing is written for a job with no ruling. `_write_job_rulings` is the
  one writer for both templates. The register is owner-edited business data
  and never enters the repo; `python3 shared/job_rulings.py RP6586` prints it.
  **Same run, second defect:** RP6586 sat on `Test - RP` but not on `Test-Master`
  (the master had not been re-run with it), and `load_wip_master` reads ONLY
  Test-Master when it exists - so the card's Bid Proposal and ETC were blank
  and the projection block was empty. The division-tab overlay (which already
  supplied STATUS) now also carries a project that is ONLY on `Test - CP` /
  `Test - RP`: ORIGINAL CONTRACT + APPROVED COs and ORIGINAL ESTIMATED COST +
  CO COSTS (the tab's TOTAL columns are live formulas and read None from a
  workbook saved without a recalc, so the parts are the fallback), with
  `wip_source` = the tab and row, printed under the ETC input in grey
  ("contract + ETC from the WIP master · 'Test - RP' row 25"). A job on
  Test-Master is untouched; a typed value on the prior sheet still wins.

- **THE CP AND RP OVERVIEWS WERE DEAD ON ARRIVAL (2026-09-04, caught by the CP
  batch).** `_formula` in `completed_pnl.build_bundle` returned a DICT LITERAL
  and indexed it - so every branch was evaluated, including `L['moh']`, the
  9%-of-contract column that exists for **MFD only**. On CP and RP that raised
  `KeyError: 'moh'` and the Overview never built; the per-job workbooks were
  all fine, so the failure only showed at the very end of a 19-job run. Now
  one `if` per key. **A dict literal is not a switch** - if any branch can be
  invalid, build the one you need.

- **ONE FILE PER JOB, AND `FINAL` IS THE FINISHED NAME (2026-09-04).** A
  finished job folder still held THREE workbooks - `Project_PnL_<job>.xlsx`,
  `<job> Job Result.xlsx` and `<job> FINAL Closeout.xlsx`. The generators were
  retired 2026-09-03 but their output was never swept, and the leftovers rode
  the folder move into the Teams channel. The owner: "why is there three? for
  completed just rename the P&L to final, for actives just use the original
  name of Project Pnl. delete job result everywhere and clean this up."
  **25 stale workbooks deleted** (14 Job Result + 11 FINAL Closeout).
  **Naming now:** live job -> `Project_PnL_<job>.xlsx`; a job filed under a
  `completed …` archive -> **`<job> FINAL.xlsx`**. `shared/pnl_paths` owns both
  (`pnl_filename` / `is_archived_dir`), the export writes whichever the output
  folder implies, and both readers know both names - `find_pnl`'s candidate
  sweep and `completed_pnl._find_workbook`. The 11 finished MFD workbooks were
  renamed in place.

- **`active mfd` FOUND NOTHING - A BLANK STATUS IS ACTIVE (2026-09-04).**
  `project-pnl active mfd` reported "no Active projects in the WIP master" and
  did nothing. `expand_active_projects` required the literal word `Active` in
  Test-Master STATUS, but **MFD rows carry a BLANK status** - MFD closes by
  hand, so nothing ever writes one. Measured: 3 of 3 MFD rows blank, 19 of 19
  CP `Active`, 119 of 119 RP `Active`. Blank now counts as Active, which is
  what the ledger already did (`dashboard.py`: "blank = MFD, active by
  construction"). Two readers of the same column had disagreed for months and
  only the batch shortcut was wrong.

- **THE ACCESS TOKEN EXPIRES MID-RUN AND NOTHING REFRESHED IT (2026-09-03/04).**
  An Intuit access token lives ONE HOUR. Every tool minted one at startup and
  passed that string around for the rest of the run, so any batch longer than
  an hour died on `401 AuthenticationFailed` partway through. It cost three
  overnight MFD regens: a `--legacy --class` batch of 8 finished jobs takes
  60-90 minutes, every job past the hour mark failed, and each run reported
  **"Done - 0 workbook(s)"** with rc=0 - a silent, total loss that looked like
  a successful run in the log's tail.
  **Fixed in `shared/qbo_api.py`** (so every tool gets it): a process-wide
  `_CURRENT_ACCESS`, a `refresh_access()` that re-does the bearer exchange, and
  `_api_get` minting a fresh token on a 401 and retrying the same call (twice
  at most, so a genuinely dead refresh token still fails fast). `_api_get`
  reads the process-wide token in preference to whatever string the caller
  holds, so a stale token captured before the refresh cannot outlive it. The
  Keychain read needs no prompt, so it is silent inside an unattended run.
  Verified live: refresh mints a new token and a call made with the OLD string
  still succeeds.
  **Two lessons:** rc=0 with "0 workbook(s)" is a FAILURE - the batch summary
  must be read, not just the exit code. And a long unattended batch has to be
  assumed to outlive its credentials.

- **ARCHIVE ORDER DEFECT (2026-09-03, caught by the post-regen diff).** A
  finished job regenerates into the first archive folder that already holds it
  (`_resolve_project_out_dir`), and `pnl_paths._archive_dirs` listed the OLD
  OneDrive tree before the routed Teams channel - so the 11 completed MFD jobs
  quietly regenerated into `Automations-/PROJECT P&Ls/Multi-Family/completed …`
  while the channel kept the morning's copies, and the channel's Overview was
  assembled from stale workbooks. The log said "wrote MFD172/…" and looked fine;
  only the written-time column in the verification diff showed 10:55 on 11 of
  14. Fixed: the routed division folders come first in `_archive_dirs`. **Always
  print the workbook's mtime in the post-regen diff**, not just the figures - it
  is the one column that catches a right file in the wrong place.

- **P&L SHEET: FOUR OWNER FIXES IN ONE PASS (2026-09-03, MFD test subject).**
  1. **Change Orders / CO Costs / Revised rows appear only when there is a CO
     or a CO cost** ("it just takes up space"); with none, the Original rows are
     the cells everything measures against, and the rule above Revised Contract
     is gone. The no-contract fallback then lands IN the Original Contract input
     (`=<total billed>`, still yellow, still overridable - the read-back skips
     formulas) and the ETC input gets `=<costs to date>`; both carry a small
     grey note naming the stand-in.
  2. **"Billed to Date (incl. retainage)"** - the total always is (gross work +
     retainage billed back + retainage moved by JE).
  3. **ACCUMULATING COSTS left the P&L for the Next Draw sheet**, which now IS
     that block: Job Type > cost code rows open (as the P&L block read), vendors
     and bills collapsed under them, bills newest first, every total a SUM of the
     rows beneath, then Total / Labor already paid / "Draw needed" under each
     overhead view. The P&L keeps one linked line under the coverage table
     ("Accumulating toward the next draw ... ➜ Next Draw" + the amount). The tab
     sits **right before the newest draw tab** in a lighter blue (`9DC3E6`) than
     the draws (`2E75B6`) - the forming draw heads the run (the owner: "not grey,
     right next to the latest draw sheet, a different shade of blue").
  4. **DRAW COVERAGE shows the overhead $ per view, not just the net**, and MFD
     gets BOTH views (9% then 10%; CP the 10% alone): per view `OH $` · `Net
     Profit` · `Net Cov %`, then `% Compl`. **Overhead per draw is on
     COMPLETION** (the owner: "this is based on completion, make sure to use the
     real metrics"): `rate x contract x (this draw's costs / ETC)` - the share of
     the contract's overhead the draw's WORK carried, off real costs against the
     ETC, not what the draw happened to bill; the shares sum to rate x contract x
     % complete. Falls back to rate x draw income only when ETC is 0. Sign
     colours on the overhead columns are conditional formats (Excel knows the
     value, Python does not). Column widths, the % Compl column and the vertical
     rules follow the view count.
  Verified on MFD325 (12 draws, COs absent, accumulating 78,571.61) and MFD133
  (completed, `--simple`, blank inputs): every formula inspected after the
  gutter pass; `assert_clean` passed in `safe_save`.

- **OVERHEAD IS A % OF THE CONTRACT - ONE RULE, P&L AND OVERVIEW (2026-09-03,
  MFD as the test subject).** Three sessions hit the same defect from three
  sides in one afternoon and pulled in different directions: one fixed only the
  Overview (`dbb594e`), one wrote an earned-revenue proration on a worktree
  (`wt/overhead-on-contract`, branch kept, worktree removed - **do not land it**),
  one wrote the spec. The owner ruled, three times in plain words: "it's
  contract 10%", "completed jobs use the total billed as contract ... take 10%
  or 9% of the total contract", and against proration - "why are you reverting
  to wip stuff when we are looking at actuals". This is that rule, everywhere:
  - **Per-job P&L:** ACTUALS overhead = `pct x Revised Contract` (was `x Billed
    to Date` - the one line that moved every draw); snapshot (3) company AND MFD
    lines = % of contract (MFD was 9% of COSTS, a different base from its own
    Overview); projection label says "of contract" (its formula already was).
    Gross profit stays `billed - cost`: the P&L is actuals, not WIP.
  - **No contract on file => total billed stands in**, as an `IF` on the Revised
    Contract row, and **Revised ETC falls back to costs to date the same way**;
    the pair MUST move together or a filled contract over a zero ETC shows the
    whole contract as profit. A typed value wins the moment a PM enters one. A
    small grey note names the stand-in. This is what makes a finished job's
    projection block live instead of zero (11 of 14 MFD jobs are off the WIP
    master and had blank inputs).
  - **Per draw:** both views net overhead on the draw's INCOME (draw income sums
    to the contract); MFD's draw coverage / KPI / "draw needed" used costs.
  - **RP job card:** `% of contract` (the bid; billed stands in), rate from
    `--overhead-pct` (the old card hard-read "of billed").
  - **Overview (`completed_pnl.py`):** `_read_contract` reads Original Contract +
    Change Orders off each P&L (the Revised cell is a formula with no cached
    value), billed stands in when blank; `oh = 10% x contract`, `moh = 9% x
    contract`; a **CONTRACT column** leads the table and **`10% OH` / `9% OH`
    columns** sit beside the net each produces (the owner: "how much OH does the
    10% and 9% make up?"). Layout as the owner drew it (`dbb594e`): `10% OH` /
    `FINAL NET PROFIT`, heavy rule, `9% OH` / `FINAL NET PROFIT`.
    **Every derived cell is a FORMULA** (the owner: "make formulas instead of
    putting straight numbers"): GP `=billed-cost`, OH `=contract*rate`, net
    `=GP-OH`, subtotals `SUM()` over the section's job rows, the ALL row the
    sum of both, and the KPI strip reads the ALL row. Job sheets: the strip
    reads the sheet's own INVOICED / COSTS totals plus `Summary!<CONTRACT cell>`;
    INVOICED is a `SUM` of the invoice rows, each vendor a `SUM` of its lines,
    account = its vendors, section = its accounts, job cost = its sections.
    Only the three source figures (contract, billed, cost) are values. Written
    with openpyxl so there are no cached results - Excel computes on open;
    nothing reads the Overview back with `data_only`.
  - **MFD's folder IS the synced Teams channel now**: `Multi-Family-Project
    Financials - Documents` under the personal OneDrive root (the owner moved
    the tree 2026-09-03 10:55). `channel_dir` resolves it as designed - it just
    was not LISTABLE by this process until mid-afternoon, so the first regen of
    the day wrote to the old `Automations-/PROJECT P&Ls/Multi-Family` folder and
    was stopped and re-run. The old folder still holds today's earlier copies;
    it is not written any more.
  - **CP Overview reads the Common drive** (`_iter_jobs`): CP P&Ls are written to
    the awarded folders when the share is mounted, so the OneDrive Commercial
    folder held stale copies and the Overview summarised them. Not regenerated
    today - the owner scoped this to MFD.
  - The ledger's `dashboard.py` overhead constants follow (contract base, MFD 9%
    of contract; per-draw = % of the draw's income) so the site cannot disagree.
  **Verified:** scratch Overview built from the 14 MFD workbooks, `assert_clean`
  passes, every derived cell inspected as a formula; the 14 MFD P&Ls regenerated
  into the Teams folder with the recipe below (the last run rebuilds
  `MFD Overview.xlsx` there); MFD133's regenerated P&L inspected row by row.

- **HISTORY CORRECTION — commit `d436ed2` contains work its message does not
  describe (2026-09-03).** That commit is titled "gutter: shift WHOLE-COLUMN
  references too" and does contain that fix, but roughly 25 of its 45 added
  lines are ANOTHER session's auto-Overview feature (`--no-overview`,
  `touched_divs`, the `rebuild_overview` call at the end of a run). Two sessions
  were editing `project_pnl_export.py` in the same clone at the same time, and
  a `git commit -- <path>` swept in whatever was in the working tree.
  Verified: `git show d436ed2^:project-pnl/project_pnl_export.py` matches
  `rebuild_overview|no-overview` **0** times, `d436ed2` matches **2**.
  The feature itself is owner-requested and correct - any run that writes a P&L
  rebuilds that division's Overview, `--no-overview` opts out. Not amended: the
  commit is pushed to a shared branch and rewriting that history costs more than
  the misattribution does. Recorded here instead so nobody trusts the message
  over the diff.

  **The rule this broke:** a pathspec commit is only as safe as the working
  tree. `git commit -- <file>` commits the file's CURRENT CONTENT, not just
  your own edits to it. Committing by path protects you from OTHER files a
  teammate has staged; it does not protect you from their edits inside YOUR
  file. In a shared clone, check `git diff <path>` before committing it, not
  just `git status`.


- **THE REPAIR PROMPT THAT GOT THROUGH (CP800, 2026-09-03) - read this before
  adding a table to any sheet.** The gutter pass hangs row 1's title back into
  column A. On `Draw Data`, row 1 is not a title, it is the **table header** -
  so the hang moved "Draw" into A1 and left B1 empty, the table ref still
  started at B1, and openpyxl wrote that blank header out as a tableColumn
  literally named **`"None"`**. Excel opens that with "we found a problem with
  some content".

  **`assert_clean` passed it**, and that is the lesson: its table check
  compares a table's ref to the sheet's ROW COUNT (the stale-ref case from
  2026-08-17) and validates the displayName and duplicate column names - but it
  never compares a tableColumn's NAME to the header cell underneath it, and
  `"None"` is a non-blank string so the blank-name check does not fire either.
  A table can therefore be internally valid and still disagree with its sheet.
  Fixed at the source: the title-hang is skipped when any table on the sheet
  starts at row 1 (`_ref_first_row`).

  **Checklist when a generated sheet carries an Excel Table:**
  1. The ref's column span must equal the number of tableColumns.
  2. Every header cell under the ref must be non-empty AND match its
     tableColumn name.
  3. The ref must not end past the sheet's last row (this one `assert_clean`
     does cover).
  4. Anything that MOVES cells afterwards - a column insert, a title hang, a
     row delete - has to move the table ref with them, and must not blank a
     cell the ref still covers.


- **Labor / Concrete LEDGER groups by VENDOR and filters (2026-09-02).**
  The ledger used to group by cost code, fully expanded. It now opens as one
  collapsed block per vendor - "where did it go" first - with the COST CODE
  moved into its own COLUMN so the block can be cut by it, an AUTOFILTER over
  the ledger only (never the scoreboard above it), and every subtotal a
  `SUBTOTAL()` formula, so filtering to a single draw re-totals the sheet.
  DRAW values on CP800 filter to: Before draw 1 · Draw 1..4 · "Draw 4 · pushed
  from Draw #3".

  **`SUBTOTAL(9)`, never 109.** 109 also ignores MANUALLY hidden rows, and a
  collapsed outline group IS manually hidden - every vendor total would read 0
  the moment it was collapsed. 9 ignores filtered rows only.

  The PM's green marks survive this: `read_back_ledger_marks` keys on
  (bill #, date, vendor, amount) and finds its columns from the header row, so
  re-ordering and re-grouping rows cannot break it.

- **`--out` is now AUTHORITATIVE (2026-09-03).** It was silently ignored for CP
  (which routes to the Common-drive awarded folder whenever the drive is
  mounted) and for any job already filed under an archive, so a run aimed at a
  scratch directory overwrote the live workbook instead - which is exactly what
  happened to CP800 during this change. If someone names an output folder, that
  is the output folder, and the run says so.


- **`Draw Data` sheet - the flat table a PivotTable sits on (2026-09-02).**
  One row per draw transaction (Draw · Period · Date · Month · Vendor · Cost
  code · Category · Bill # · Amount · Description · Pushed from) as a real
  Excel Table, so the owner can re-cut a draw by vendor AND by cost code the
  way a pivot does. **openpyxl cannot AUTHOR a PivotTable** - `add_pivot`
  exists to preserve one across a read/write, and building one from nothing
  means hand-assembling the cache definition and records, which is exactly the
  hand-written XML rule 5b exists to prevent. So the workbook ships the tidy
  rectangular source instead and Excel makes the pivot in three clicks. Sits
  with Cash Flow AHEAD of the draw sheets, never behind them.

- **Budget vs Actual opens COLLAPSED** (2026-09-02) - the cost-code scoreboard
  first, transactions on demand, same as By Account and the draw sheets.
  Draw transactions now read in DATE order within a vendor, not
  biggest-amount-first: a draw is a period, so its bills should read as a run
  of dates.

- **Gutter pass: two defects the Draw Data table exposed.** A row-only freeze
  ("A2") was being shifted to "B2", which freezes the gutter column too and
  turns a 2-pane view into a 4-pane one - `xlsx_verify` correctly called that
  corrupt and blocked the write. And the pass never moved **Excel Table refs**,
  the single failure rule 5b names by name; there had been no table in a P&L
  until now. Both fixed, and the corruption gate is what caught them.


- **The PUSH - a bill carried into a later draw by agreement (2026-09-02).**
  Bills land in a draw by date (TxnDate inside the invoice's Period tag). When the
  user agrees with a supplier to carry end-of-period bills into the next draw, a
  rule in `<CompanyHealth>/Registers/draw_moves.json` (read by **`shared/draw_moves.py`** -
  project + vendor substring + cutoff `after` / `through` + `move_to` date +
  draw numbers + why) makes `bucket_costs_by_draw_window` and `code_costs_by_draw`
  bucket those bills AS OF the rule's date. The bill keeps its real date; the
  receiving draw sheet says "N bills · $X pushed in: <vendor> bills dated after
  <cutoff> were pushed in from Draw #a - <why>", the draw they left says
  "pushed out", every moved row's Where/status reads "pushed from Draw #a", and
  the Labor/Concrete ledger's DRAW column carries the same mark. The console
  logs each move per draw. The same rule file drives the Bill Tracker match and
  the ledger's draw bands, so all three agree. No rule file → nothing moves.
  First use: CP800 Preferred Materials after 07/20/26, Draw #3 → Draw #4.

- **MFD ROUTES TO THE TEAMS 'Project Financials' CHANNEL (2026-09-03).**
  `shared/pnl_paths.DIVISION_CHANNELS` maps a division to a Teams channel, and
  `division_dir` / `division_dir_note` resolve it: explicit `--out` ·
  `ACB_PNL_DIR_<DIV>` · the synced channel · the OneDrive division folder.
  A Teams channel's Files tab IS a SharePoint folder, so once the channel is
  synced on the Mac it is an ordinary path — the same shape as
  `Company Files - WIP Report`. No Graph API and no new key.
  **It is a MOVE, not a mirror** (the user 2026-09-03): two copies of
  `MFD Overview.xlsx` would drift, and the folder link the owner shares has to
  be the one with the live numbers. The `completed mfd project p&l` archive
  travels with the division — `_archive_dirs()` resolves division folders
  instead of hard-coding them, so a finished job still regenerates into its
  archive rather than spawning a second copy at the top level.
  **A run whose channel is not synced falls back to OneDrive and SAYS SO** in
  its note; it does not pretend to have routed. `find_pnl` still searches the
  pre-move OneDrive folder, so a P&L generated before the move is found.
  **Setup is one click, once:** Teams → the channel's Files tab → Sync.

- **OVERHEAD IS A % OF THE CONTRACT, BOTH VIEWS (2026-09-03).** The owner:
  "completed jobs use the total billed as contract and make the oh correct then
  take 10% or 9% of the total contract." On a finished job the contract IS the
  total billed - what it ultimately sold for - so the Overview now charges
  **10% of billed** and **9% of billed**, shown as two stacked views split by a
  heavy rule: `10% OH` / `FINAL NET PROFIT`, then the rule, then `9% OH` /
  `FINAL NET PROFIT`. The per-job columns read `FINAL NET (10% OH)` and
  `FINAL NET (9% OH)`.
  **This REPLACES the 9%-of-COSTS basis** for the Overview. A cost-based
  overhead rose with the overrun, so the worse a job went the bigger its
  overhead charge - backwards, and it flattered nothing.
  **CLOSED the same day** - the per-job P&L now uses the same contract base
  (see the entry at the top of this list); the owner's word was "it's contract
  10%", and all 14 MFD P&Ls were regenerated.

- **RETIRED 2026-09-03 — ONE JOB, ONE P&L.** A finished MFD job folder held
  THREE workbooks of the same figures: `Project_PnL_<job>.xlsx`, `<job> Job
  Result.xlsx` (`completed_pnl.py`) and `<job> FINAL Closeout.xlsx`
  (`closeout.py`) — plus `Closeout Index.xlsx` and a never-once-run
  `completed_rollup.py` → `Completed MFD P&L.xlsx`. Five outputs, one set of
  numbers, every one of them re-derived from the P&L. The owner called it:
  "why is there a job result excel? shouldn't this be merged with the P&L? i
  feel that we are confused and all over the place."
  **The simplified finished-job report was already a FLAG on the real P&L** —
  `--simple` drops the draw sheets, the Next Draw sheet and the coverage
  blocks, and the MFD recipe above already runs completed jobs with it. The
  second file bought nothing and drifted: MFD177/192/325 were carrying a
  "finished job" report on an ACTIVE job, dated a day BEFORE the P&L beside it.
  **Deleted:** `closeout.py`, `completed_rollup.py`, `completed_pnl.build()`
  (the Job Result writer, 165 lines) and 26 stale workbooks on OneDrive.
  **Kept:** `Project_PnL_<job>.xlsx` per job · `<DIV> Overview.xlsx` per
  division. Before deleting a generator again, check whether the shape it makes
  is already a flag on the P&L.

- **THE OVERVIEW REBUILDS WITH THE P&L (2026-09-03).** The owner: "make sure
  now if we update any mfd p&l it will get updated on the overview." The
  Overview is assembled FROM the division's workbooks, so a run that rewrote
  one left it describing figures that no longer existed. `project_pnl_export`
  now calls `completed_pnl.rebuild_overview()` for every division a run
  touched, before it prints the summary; `--no-overview` opts out, and a
  failure there is reported but never fails an otherwise good P&L run. It reads
  workbooks only — no QBO, no credential unlock — so it is cheap enough to do
  every time. `load_division()` is the ONE reader the CLI and the auto-rebuild
  share, so the same workbook can never be assembled two different ways.


- **THE MFD REGENERATION RECIPE (2026-08-31) — stop guessing the flags.**
  All 14 MFD P&Ls were rebuilt onto the gutter layout. The flags are NOT
  interchangeable; run them exactly like this:

  | Jobs | Command |
  |---|---|
  | MFD133 160 166 182 183 186 231 281 | `--class --simple` (one batch is fine) |
  | MFD295 | `--class --simple --infer-periods` |
  | MFD172 | `--legacy --alias 'BONDS RANCH' --class --simple` — **alone** |
  | MFD228 | `--legacy --class --simple` — **alone** |
  | MFD177 MFD192 MFD325 | no flags (live jobs, full draw template) |

  `--class` alone is NOT enough for MFD172 (short by **192,526**) or MFD228
  (short by 11,381) - both need `--legacy`'s bill-memo rule, and MFD172 needs
  its street alias too. `--alias` is refused on a multi-project run, so those
  two must run on their own.

  **Always snapshot billed/cost per job BEFORE regenerating and diff after.**
  That is what caught the two wrong-flag jobs, and it is the only way to tell a
  flag mistake from a real QBO change: with the right flags 12 of 14 jobs
  reproduced to the cent, and the two that moved were traced to edited bills
  (MFD177 -10,933: 16 lines recoded off the job, several of whose own
  descriptions name MFD186/MFD172, plus bill D78027 edited 4,923.28 → 4,623.28;
  MFD228 -484.50: bill JMP07252024MFD lost a line). A silent drop with no
  explanation is a flag bug, not a data change.


- **THE LEFT GUTTER IS THE STANDARD FOR EVERY P&L (2026-08-31, signed off).**
  Column A is a narrow gutter on every sheet and content starts in B; only the
  row-1/row-2 titles stay in A, hanging into it. On the draw sheets the KPI band
  and its `DRAW SUMMARY` banner start in B too, so the strip lines up with the
  bills table under it. Body text is size 12 everywhere, with each sheet's
  column widths scaled by the same ratio so a bump cannot clip anything.
  **THE DRAW SHEETS GO LAST and nothing sits behind them** (the user
  2026-09-01) - a job with a dozen monthly draws pushes anything after them off
  the end of the tab bar. POs, Reconciliations and Cash Flow all moved ahead of
  the draws; every one of them is read more often than any single old draw.

  **Both are POST-PASSES in `safe_save` (`_normalise_body_font`,
  `_apply_left_gutter`), not edits at the call sites, and that is deliberate:**
  the builders hard-code column numbers in ~310 places AND build 270+ formulas
  out of literal column letters (`"=D7-E7+G7"`, `"=Transactions!E568"`).
  Re-indexing that by hand is how you ship a workbook whose formulas quietly
  point one column off. A uniform one-column shift, applied once, rewrites every
  reference mechanically - cross-sheet references included, because every sheet
  moves by the same amount. Repo rule 5b in full: `insert_cols` moves cells and
  styles and NOTHING else, so merges, widths, freeze panes, hyperlink anchors,
  conditional-format ranges, the print area and every formula are re-derived by
  hand, and `assert_clean` still runs last. A sheet that ALREADY reads
  gutter-first (the draw sheets, `By Account`) is skipped rather than shifted
  twice. Verified on MFD325 cell-by-cell: every value, format, merge, width,
  freeze pane, print area and outline landed exactly one column right, nothing
  else moved.
  **Do not "fix" a font size or a column index at a call site** - the delivered
  size and the delivered left edge are decided in those two passes, and editing
  both places double-applies.


- **One overview PER DIVISION, in the division folder (2026-08-31).**
  `completed_pnl.py --division mfd|cp|rp --bundle` writes
  `<division folder>/<DIV> Overview.xlsx`. The MFD one moved OUT of the archive
  subfolder - it covers live and finished jobs alike, so filing it under
  "completed" put it somewhere it did not belong. CP and RP default to the
  CURRENT YEAR only (`--year 2026` / `--year all` to override): a job is kept
  when it has an invoice or a cost dated in that year, which keeps a job still
  running from last year and drops one that finished before the year started
  (the user 2026-08-31: "just for this year projects, don't go further back").
  The MFD 9%-of-cost column and its overhead row appear only for MFD; CP and RP
  get the company 10%-of-revenue view alone, and the table's column count now
  drives the metric-strip spans instead of being hard-coded.

- **The reader reads the TRANSACTIONS sheet, not `By Account` (2026-08-31).**
  Only the current CP/MFD template has a `By Account` sheet - CP672 and every RP
  workbook have none, so a CP or RP overview could never have been built from
  it. Every template writes the Transactions cost block the same way (vendor
  row, then its lines carrying an Account), so the account tree is regrouped
  from there and one reader now serves MFD, CP and RP. Verified equal **to the
  cent on all 14 MFD jobs** against the old `By Account` path before switching,
  and it also retires the old-layout problem: the 11 archived MFD workbooks no
  longer need regenerating to be readable.


- **P&Ls are SORTED INTO THE DIVISION FOLDER (2026-08-31, binding).** The
  OneDrive folder LINK is the unit of sharing: the owner sends a PM the link to
  their division, so a P&L landing at the `PROJECT P&Ls` root would put every
  other PM's numbers behind that same link. One rule, in
  `shared/pnl_paths.division_dir()` -> `Commercial` / `Multi-Family` /
  `Residential` (folder names already on OneDrive; an unrecognised project #
  stays at the root rather than risk being misfiled into the wrong division).
  Applied at every writer: `_resolve_project_out_dir` (CP/MFD), the RP
  `<proj> - <client>` folder, and the draw cross-check workbook - which used to
  land loose at the root where nobody it was written for could see it.
  `pnl_paths._archive_dirs()` now sweeps each division folder as well as the
  root, and `find_pnl` keeps the pre-division root path as a candidate so
  nothing already written goes missing. The disk was ALREADY sorted by hand;
  this is the code catching up, so no files moved. Fixed the same stale
  assumption in `completed_pnl.py`, `completed_rollup.py` and
  `one-offs/pnl_line_level_audit.py`, whose default archive path had silently
  stopped resolving.

- **Overview: column A is a narrow gutter, content starts in B (2026-08-31).**
  `completed_pnl.py --bundle`, both the Summary and the per-job sheets (the
  user: "goal: have ability to move info away from left side"). Only the big
  title stays in A, hanging into the gutter and spilling across. `lint_layout`
  takes a `first_col` so the deliberate gutter doesn't trip its
  empty-column/ragged-left-edge checks. Overview workbook links are RELATIVE
  and forced to `/` separators (`_link_target`) so they resolve on Windows and
  Mac and survive the tree being re-shared under a different OneDrive root.
  Totals unchanged by the shift: 14 jobs, billed 54.6M, cost 49.2M, GP 9.74%.

- **Dollar figures genericized to `~$Nk` form (2026-08-27)** in this file and in
  `project_pnl_export.py` comments; the CI leak guard's TEMP exclusion for the
  export is retired and the guard now also catches non-round six-figure amounts
  (patterns live in `.github/leak_guard.sh`). Comment/doc wording only - no
  behavior change.

- **`cost_leaf` moved to `shared/qbo_costs.py`** (2026-08-08) — the ledger's `load_costs.py` needs
  the SAME cost-code resolver, so it graduated to shared/. This tool imports it back
  (`from shared.qbo_costs import cost_leaf`); byte-compatible, no behavior change (imports + compiles
  clean, verified). Do not re-add a local copy.
- **P&L + Transactions + Draws + POs + Cash Flow + Reconciliations** sheets.
- **Batch mode** — `project-pnl active cp|rp|mfd` regenerates every Active
  project of a division (Active = the WIP master's Test-Master STATUS).
- **Budget vs Actual** (CP + RP) — takeoff cost-code budget vs QBO cost-code
  actuals, every transaction listed under its code with a QBO link, job-type
  color bands, class-mismatch flags.
- **QBO deep links (2026-08-06).** Every transaction row already deep-links to its
  QBO txn (Transactions, draws, Next Draw, Budget vs Actual, Labor/Concrete, POs,
  RP Job P&L, Pending Review, Reconciliations). Added a header **"Open Project in
  QBO"** link → the project HOME page (`customerdetail` via `_qbo_customer_url`,
  NOT the P&L report) on the CP/MFD `P&L` sheet (**I2**) and the RP `Job P&L`
  (**E2**), distinct from the per-figure Billed/Costs links. Stored
  `cell.hyperlink`, never `=HYPERLINK()`.
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
  **LEDGER** (below, fully expanded, nothing collapsed): **↗ (QBO bill
  page)** · QBO # (linked to the SCAN, the user's identifier) · DATE ·
  VENDOR · QTY · RATE · AMOUNT · [SALES TAX] · DRAW label · DESCRIPTION.
  The ↗ first column is the direct QBO link (the user 2026-08-10 — it went
  missing when QBO # became the attachment link). **Column A is a DEDICATED
  4.5-wide ↗ lane and the scoreboard starts at column B** (the user
  2026-08-10, rejecting the first cut: arrows floating in the 34-wide ITEM
  column read as slop). Every scoreboard formula, band fill, conditional
  format and autofit range is one column right of where it used to be.
  Verified by LOOKING at the rendered CP585 Labor + Concrete sheets in Excel,
  not just by reading cells back — that is now the standing bar for any
  layout change here. The mark readback is
  HEADER-DRIVEN (finds the QBO #/DATE/VENDOR/AMOUNT columns by name), so it
  reads both pre- and post-↗ layouts and survives future column moves. The QBO # link opens the UPLOADED BILL
  FILE, not the QBO bill page (the user 2026-07-31): QBO's attachment URLs
  expire in minutes, so the exporter downloads each scan into `attachments/`
  beside the workbook and links the local copy (offline, no QBO login);
  bills with no attachment keep the QBO https link. Downloads are idempotent.
  **The link is a STORED RELATIVE target — settled empirically on the Mac
  2026-07-31** after a formula detour: `=HYPERLINK()` formulas hard-fail in
  Mac Excel's sandbox ("Cannot open the specified file", every URL form
  tested — bare path, file://, %20-encoded), and a HyperlinkBase property
  breaks resolution outright. A stored relative target opens the file (after
  a ONE-TIME macOS "Grant File Access" per file — that's the click friction,
  not a broken link) and **survives a Mac Excel save unrewritten** (verified
  by saving in Excel and re-reading the sheet rels). Windows resolves the
  same relative target against the share path it opened from — **CONFIRMED
  working by an estimator on Windows, 2026-08-05**. Cross-platform question
  closed. Requires opening the workbook from the share (an
  emailed copy has no attachments/ beside it). Multi-scan bills download
  into their OWN `attachments/<bill #>/` subfolder and the "(N files)" cell
  opens that folder — not the whole attachments library (the user
  2026-07-31); single scans stay flat and open directly. Legacy flat files
  are moved into the subfolder on the next run, not re-downloaded.
  The company-wide Attachable sweep (~10 min) is cached for 7 days in
  ~/Library/Logs/Proficient/project-pnl/ — attachments uploaded since the
  cache was built appear after the TTL, or delete the cache file to re-sweep.
  Downloads always fetch a fresh TempDownloadUri per file (one GET each). A bill lands in exactly one draw, so a
  label column replaces the per-draw matrix that guaranteed blank cells on
  every bill row. Tax folds onto the bill row it came from (joined by bill
  #). Tax/fuel columns appear ONLY on a trade that has such lines — labor
  subs bill neither (auto-omit, disclosed). No fuel lines exist anywhere yet
  (AP folds the surcharge into the rate); Concrete says so on the sheet.
  Martin Marietta bills it as "SERVICE CHARGE" — classified into the same
  bucket (column reads FUEL / SVC CHARGE) and folded onto the bill's row,
  so a ready-mix bill is ONE ledger line: qty · rate · amount · tax · svc
  (the user 2026-08-01).
  **PM-confirmation marks survive re-syncs** (the user 2026-07-31): an
  estimator marks a ledger row GREEN when the PM confirms the bill; before
  each regeneration `read_back_ledger_marks` lifts every manual row fill
  from the prior workbook (keyed bill # + date + vendor + amount) and the
  builder re-applies the exact color — so green means confirmed, and any
  other color convention survives too. A mark whose bill changed in QBO
  (amount/date edited) no longer matches and the run flags it — deliberate:
  a changed bill needs re-confirmation. Scoreboard band fills are excluded
  (only rows with a real DATE are read). **INVARIANT (binding): the script
  never writes a direct cell fill on a dated bill row** — that is the whole
  ownership model: colors are never interpreted or matched, so the
  estimator's palette can be anything; script coloring on data rows must use
  CONDITIONAL FORMATTING (a separate xlsx layer readback cannot see).
  Limits: a white fill is not a mark; a single painted cell is read as a
  row mark and re-applied to the whole row.
  **Sheets arrive auto-fitted** (the user 2026-07-31): `_autofit` computes
  what Excel's double-click would — every column sized to its longest
  display line at font 12, wrapped header rows sized to their line count —
  measured ONLY over the scoreboard, yards strip and ledger rows, never the
  long note lines that spill by design (DESCRIPTION also excluded — it
  spills). No more clipped draw periods. 2026-08-04 layout pass (the user):
  NO freeze pane; money cells in accounting $ format (ACC_FMT — $ pinned
  left, zeros as "-", red parens); the tie-note and columns-omitted filler
  lines removed; Concrete's yards/$-per-yd strip parked TOP-RIGHT beside the
  title (cols J+, rows 1-2, lump note beneath) instead of a band between
  scoreboard and ledger.
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
- **Payment state everywhere it's asked** (the user 2026-08-05): each draw
  sheet's title leads with PAID (green) / UNPAID (red) — PAID when every
  invoice in the draw has a zero open Balance in QBO. The draw bill tables
  and the Transactions sheet (income rows AND every bill line) carry a
  "Paid?" column from the same Balance test; purchases count as paid by
  nature. PM-report rows have no QBO bill and show nothing.
- **Draw sheets lead with a horizontal KPI strip** (the user 2026-07-29):
  income → retainage held → net draw → costs → gross profit → gross margin %
  → overhead → REAL net profit → REAL net %, big type, $-formatted, profit
  cells colored by sign, and the strip is the freeze pane. MFD keeps its
  PM-vs-QBO comparison as a second strip. Replaces the old vertical summary
  box. Draw-sheet body font is 12 (was 11).
- **Voided invoices are dropped everywhere** (the user 2026-08-05, found on
  MFD192): QBO zeroes a voided invoice and prefixes the memo "Voided - ";
  those never belong on a P&L and used to clutter the untagged block. Note:
  MFD192's three contracts (main / HUDSONWOOD / OFFSITE) already combine
  into one draw per month by shared period tag — verified with the user, no
  structural change was needed.
- Overhead: a % of the CONTRACT - 10% company, 9% MFD view (total billed stands in when no contract is on file).

- **RP template fixed 2026-08-06**: the "Open Project in QBO" header link
  (added 7fa2b40) wrote E2 on the RP Job P&L — inside the meta block's A2:H2
  subtitle merge → 'MergedCell.value is read-only' crash on every RP run.
  Moved to I2, matching the CP/MFD template. Also: the run no longer echoes
  the QBO company/realm id (same convention as cc2035f), and ACB_DEBUG=1 now
  prints full tracebacks behind the per-project ✗ lines.

- **WIP master resilience (2026-08-07):** someone restructured the
  Test-Master tab into a bonding-style report (TYPE/BONDED/PROFIT, no STATUS
  column) outside the repo's readers — `active cp|rp` went blind and Closed
  handling dark. `load_wip_master` now overlays STATUS from the per-division
  `Test - CP` / `Test - RP` tabs when Test-Master carries none. NOTE: MFD
  rows have no division tab, so MFD status is gone until Test-Master carries
  STATUS again — `active mfd` will find nothing.

- **LEGACY-JOB attribution (`--legacy`, 2026-08-24).** Jobs that predate
  consistent project coding carry only PART of their cost on the project
  customer; the rest is named in the line description or the bill memo, and
  their invoices sit on the PARENT customer. QBO's own project P&L report
  cannot see any of it, so those jobs used to export millions short. `--legacy`
  (plus `--alias "<street name>"`) routes every cost-line test through
  **`shared/job_lines.JobMatcher`** — project customer → line text → bill memo,
  first rule wins, and a memo naming MORE THAN ONE job number is skipped, never
  split. It also pulls the parent customer's invoices (memo-filtered) and
  SYNTHESIZES the P&L totals from the same attributed lines
  (`_synth_pl_totals`) instead of asking QBO for a report it cannot answer.
  Opt-in and scoped per project (`_set_legacy_matcher` is called per project so
  a batch can't leak one job's aliases into the next); with the flag off,
  `_line_belongs` is byte-identical to the old `CustomerRef == customer_id`
  test — verified on CP585 (identical six-figure COGS both ways). Same matcher backs
  `one-offs/legacy_job_cost_pull.py`, so the P&L and that pull can never
  disagree. First use: MFD172, reproducing its known figures to the cent.
- **CLASS/PROJECT LOOKUP — `--class-project` (the user 2026-08-25, MFD295).**
  The owner's name for it: the OLD method (class) plus the NEW one (project),
  and nothing else. It is the right method for a job that ran straight across
  the coding switchover. On MFD295 the two are perfectly disjoint - 163
  project-coded lines (Dec 2024 → Aug 2026) and 127 class-coded lines
  (Sep 2024 → Aug 2025), with **ZERO lines carrying both**. Either source
  alone reports a fraction of the job. The flag implies `--legacy`, REQUIRES
  `--job-class`, and switches the line-text and bill-memo rules OFF
  (`JobMatcher(text_rules=False)`) so the answer is exactly class ∪ project -
  on MFD295 the text rules would have pulled in a further block of ambiguous
  lines the owner did not ask to include.
- **`--infer-periods` — retroactive draw windows (the user 2026-08-25).**
  Older invoices carry no `(Period:…)` tag, and the untagged fallback is the
  CALENDAR month, which is wrong whenever the GC's window straddles month end.
  MFD295 bills the 21st through the 20th, so the fallback pushed three weeks
  of cost into the wrong draw. `shared/draws.learn_period_shape` reads the
  window SHAPE off the invoices that ARE tagged (MFD295: start day 21, end day
  20, span 1 month, learned from 3) and `infer_period_tag` writes the matching
  tag onto the untagged ones before grouping, so the existing parser handles
  them natively. The draw's MONTH is never guessed from the invoice date when
  the memo names one - MFD295's June 2025 draw was billed on the 23rd and
  still lands in 05/21–06/20. Retainage-only invoices are deliberately left
  untagged; they bill no work window and belong to the retainage blocks.
  Ties broken toward the LATER day, so one 04/20 typo among 21sts loses.
- **`--job-class` (2026-08-25, MFD228).** A fourth legacy rule: the line's
  `ClassRef` sits under the job's OWN class branch, matched as a PREFIX so the
  live parent and its deleted per-job leaf both count. Two traps it exists to
  survive: (a) **the job's class is usually INACTIVE** - a plain
  `SELECT * FROM Class` returns active only, so on MFD228 the query showed
  `MULTI FAMILY:MARKER LAPIZ` while every cost line actually carried
  `…:MFD228 (deleted)`; (b) **a division class is not a job class** -
  `JobMatcher` REFUSES a bare `MULTI FAMILY` / `Residential` / `Commercial`
  prefix, which would claim every job in the division.
- **Job numbers now match separator-tolerantly and suffix-exactly
  (2026-08-25).** `job_number_pattern` accepts `MFD228`, `MFD 228`, `MFD-228`
  (clerks write all three) while still refusing `MFD2281` and, critically,
  keeping a base job and its `-FTW` sibling apart. The suffix guard fires only
  on a hyphen-attached token (`RP7186-FTW`) or FTW in any spacing - NOT on the
  ordinary memo form `MFD172 - 1392 E Bonds Ranch Rd`, where the spaced hyphen
  separates fields. Getting that wrong dropped 48 real lines in testing.
  Effect on live numbers: MFD228 gained $6,680 (9 lines written `MFD 228`),
  and MFD172 gained **~$105k** across 18 `MFD 172-0-20-1` sub-service draws
  that the original hand-built pull never saw.
- **`+class` — the short form, and the class is FOUND not typed (2026-08-25).**
  `project-pnl MFD228 +class` is the whole command. `discover_job_classes`
  matches the job number against each class's LEAF segment across ACTIVE and
  INACTIVE classes, then keys on the class **ID**, because QBO renames a class
  when you deactivate or reactivate it (MFD228's went from
  `…:MFD228 (deleted)` to `…:MFD228` mid-session; the id never moved). Leaf
  matching also means a division or builder branch can never be selected by
  accident, and lookalikes are safe — MFD295 does not match RP5295/RP4295.
  `+class` alone = **project ∪ class**; add `--legacy`/`--alias` to turn the
  line-text and bill-memo rules back on too. `--job-class` survives as an
  explicit override. The class list is pulled once per run, not per project.
- **P&L reads total-first, and PARTIAL is a real state (2026-08-27, all
  templates).** Three changes, asked for on MFD295 but applied everywhere:
  * **`PARTIAL — <amount> open`** replaces a bare UNPAID wherever a balance is
    known (Transactions income rows, Transactions bill lines, the draw sheets).
    One resolver, `_pay_state(balance, total)`. `paid_map` now carries
    `(balance, total)` instead of a bool, which is what makes the open amount
    available. It was calling a 280,838 invoice with 389.70 left UNPAID, which
    reads as a collection problem rather than a rounding tail.
  * **Section totals sit ON the header bar**, not in a total row underneath —
    `Cost of Goods Sold` and `Operating Expenses` carry their own sum and the
    accounts detail them below. `acct_lines` returns the HEADER row now, so
    every downstream ref (Costs to Date, Gross Profit) still points at the
    total. `total_label` stays in the signature but is no longer written.
  * **`Income (incl. retainage)` lists every invoice behind it** — number,
    memo, amount, newest first, each linked to its QBO invoice. Labels are
    flattened to one line: the Period tag is stripped (it is the row above's
    identity) and the project name dropped via `_project_name_words`, so a
    memo does not repeat the client on all 14 rows. Verified to tie: the
    listed invoices sum to the bar exactly.
  NOT changed: the RP `Job P&L` keeps its flatter shape (no COGS account
  block, so nothing to move) but inherits PARTIAL through the shared
  Transactions builder. Draw sheets are still generated for every template —
  the owner deletes them in his own copy, which is his edit, not the tool's.
- **`+simple` — the stripped-back P&L for a COMPLETED job (2026-08-27).**
  Drops every forward-looking surface: the per-draw sheets, the `Next Draw`
  sheet, the `DRAW COVERAGE` table and the `ACCUMULATING COSTS — NEXT DRAW`
  block. What remains is P&L · Transactions · POs · Reconciliations · Cash
  Flow. "What do we bill next" is a settled question on a finished job.
  **NOT DONE — laying blocks ① and ② side by side.** It was attempted and
  reverted: the approach gave `row()` a `_Ref(int)` handle carrying its own
  column so a formula could render `B12`/`E12` from the same f-string. That
  works for formulas and breaks openpyxl, which builds a cell coordinate from
  `str(row)` — so `ws.cell(row=_Ref(7,'B'), column=2)` produced **`BB7`**
  (column 54), and the corruption gate caught it. Any retry must NOT override
  `__str__` on a value that is ever passed as a row: give the handle an
  explicit `.ref` property and change the ~23 formula sites to use it.
- **Completed jobs get filed, and roll up (2026-08-27).** Finished jobs live
  in an ARCHIVE subfolder of the P&L root — `completed mfd project p&l` — so
  the top level stays the live work. Both sides know about it:
  `pnl_paths._archive_dirs()` matches any subfolder starting `completed` /
  `closed` / `archive`, `find_pnl` searches inside them (else the dashboard
  reports a filed job as never generated), and `_resolve_project_out_dir`
  REGENERATES a filed job back into its archive folder instead of quietly
  creating a second copy at the top level.
  **`completed_rollup.py`** builds `Completed MFD P&L.xlsx` beside those
  folders: one row per job (contract · ETC · billed · cost · GP · GP% · cost
  vs ETC), a portfolio total, and an `OPEN ↗` link per row. It reads each
  job's **Transactions** sheet, never QBO — offline, seconds, and it cannot
  disagree with the workbook it links to. (Not the P&L sheet: those totals are
  live formulas, and openpyxl returns formula TEXT unless Excel has cached a
  value.) Links are stored relative targets, so the rollup must sit beside the
  job folders.
- **`--alias` is REFUSED on a multi-project run (2026-08-27).** It is global to
  the run, so a batch applied every job's street name to every job: rebuilding
  MFD172 and MFD228 together with both `BONDS RANCH` and `LAPIZ` made MFD228
  report **4,213,532** of cost instead of 879,732, and inflated MFD172 to
  5,422,266. The attribution was fine; the invocation was not. A wrong number
  that looks plausible is worse than an error, so the run now stops and says
  to do one job at a time. `+class` stays safe in a batch — the class is
  discovered per job.
- **`completed_pnl.py` — the SIMPLE report for a finished job (2026-08-27).**
  A separate template, not more surgery on the main exporter (the user:
  "made simply, not small font and easy to follow so we can get a birds eye
  view and swoop into the details when needed"). Three sheets:
  **Summary** — a metrics strip ACROSS the top (billed · cost · gross profit ·
  margin · overhead · net · net margin), then cost-by-account beside the
  invoices that paid for it, every account name linking into the detail;
  **Costs** — account → vendor → line, collapsed to accounts by default.
  Base font 14 (18 for the KPI figures), no cents on the birds-eye view.
  **It READS the generated workbook, not QBO** — it re-shapes
  `Project_PnL_<job>.xlsx`, whose numbers are already proven line-level by
  `one-offs/pnl_line_level_audit.py`. So it cannot introduce an attribution
  bug, needs no credentials, and runs in a second. Requires the source
  workbook to carry a `By Account` sheet (regenerate older ones first).
  This is why the side-by-side layout did NOT need the `_Ref` refactor that
  corrupted the main sheet: a fresh template owns its own geometry.
- **Bundle visual pass (2026-08-27).** The first cut was called amateurish and
  the diagnosis was competing treatments: solid navy on the tile headers AND
  section headers AND table headers, borders on every cell, row banding, and
  red on every negative — four things fighting for attention. Now ONE navy band
  anchors each table, rules separate instead of boxes, banding is barely-there
  (`F4F6F9`), and red is reserved for profit/net figures. Measured on MFD172:
  navy fills 14 cells (was every header row), bordered cells 44 (was every
  cell), red text 5. Tiles sit on white with a hairline under the label.
  Also: identifiers align LEFT (an invoice # right-aligned floated to the far
  edge of its column, away from its own header), description spills across the
  tile columns instead of needing a 74-wide column, explicit row heights, and
  landscape fit-to-width page setup so it prints/PDFs as one page wide.
  **`lint_layout()` expands merged ranges** before calling a column empty — it
  was flagging the tile columns as gutters.
- **"Funded but unpaid" flag on every draw sheet (2026-08-28).** When the GC
  has PAID a draw but bills inside it are still open, the sheet says so with
  the count, the total and the bill numbers. That state is what earns a
  supplier notice and it is invisible everywhere else — the draw reads
  collected and the job reads covered. It came from MFD325's July 2026 draw:
  Estrada 598125 sat open on a draw that had already funded it, so the PM
  widened his report to 05/30 to surface it, which dragged June's cost onto
  July's income and made a +18k draw read as -100k. The flag means nobody has
  to widen a window to find one. Also fixed here: the draw sheets painted
  every payment status GREEN, because `paid_map` became a `(balance, total)`
  tuple with PARTIAL and `GREEN if _pd else RED` is always true on a tuple.
- **Rich text is BANNED in this exporter, and every save is now gated on the
  corruption check (2026-08-24).** `_cost_code_value` / `_cost_name_value` were
  returning `CellRichText` (bold code token + regular description, the user
  2026-06-09) — multi-run inline strings, exactly what `shared/xlsx_verify`
  refuses and what makes Mac Excel offer to "repair" the file. It only ever
  showed up when the accumulating-costs block contained a cost code, so most
  P&Ls were clean by luck; MFD172 tripped it. Both helpers now return plain
  strings (style the CELL, never runs inside it) and the rich-text imports are
  gone. `safe_save` runs `assert_clean` on the TEMP file and REFUSES to publish
  a workbook that fails — rule 5b was never wired into this tool before.

## OPEN ISSUES


- **A DYING SMB SHARE HANGS A CP RUN SILENTLY (2026-09-03).** `active cp`
  ran 71 minutes with 6 seconds of CPU and three log lines: the Synology
  `Common` share dropped mid-run and a filesystem call on `/Volumes/Common`
  blocked in the kernel (even `ls /Volumes` hung). No error, no timeout, no
  output - the QBO calls all carry timeouts, the mount does not. Before a CP
  batch, confirm the share answers (`ls "/Volumes/Common/CURRENT PROJECTS"`
  returns promptly); if it is gone, CP P&Ls would fall back to OneDrive, which
  is the wrong home (one writer per file). The 19 active CP P&Ls are STILL
  PENDING for that reason - rerun `active cp` once the share is remounted.


- **6 of 17 Active CP jobs now have a readable cost-code budget.** The newer
  jobs keep the coded budget in a ROOT `Cost Codes.xlsx` on a `Cost Codes V2`
  sheet (same col A = code / col C = $ layout) — the reader learned that as a
  fallback 2026-07-31 (the takeoff's own sheet still wins when coded). Reads
  now: CP585, CP672, CP745 (takeoff) + CP785, CP831, CP961 (root workbook).
  **11 still have NO `Cost Codes.xlsx` in their folder root** — per the user
  every job should have one, so this is a work-list for the estimators:
  CP765, CP783, CP790, CP794, CP800, CP803, CP821, CP861, CP885, CP910,
  CP961→done, CP865 (no folder at all). The moment the file lands in a
  folder, the P&L picks it up with no code change.
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

- **CP790 retainage, back end (the owner 2026-10-02):** the owner is correcting the duplicated retainage in
  QuickBooks. After that, regenerate CP790 and confirm the "⚑ QuickBooks fix pending" row is gone and Billed to Date
  did not move. If QBO is fixed by shrinking the not-billed invoice, the spread places it all and nothing changes.

- Extend the Labor/Concrete sheets to RP and MFD if the PMs want them there.
- Roll the G702 contract source into the WIP readers so the master and the
  P&L can't disagree.
