# project-pnl — STATUS

Per-project P&L workbooks from QBO. One tool, one script
(`project_pnl_export.py`), three templates: CP (draw-based), MFD (draw-based,
manual close), RP (no draws — expenses → invoice → profit).

---

## DONE / FINALIZED

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
- Overhead: 10% of revenue (MFD alt view 9% on costs).

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

- Extend the Labor/Concrete sheets to RP and MFD if the PMs want them there.
- Roll the G702 contract source into the WIP readers so the master and the
  P&L can't disagree.
