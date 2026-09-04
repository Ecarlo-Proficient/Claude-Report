# project-pnl — STATUS

Per-project P&L workbooks from QBO. One tool, one script
(`project_pnl_export.py`), three templates: CP (draw-based), MFD (draw-based,
manual close), RP (no draws — expenses → invoice → profit).

---

## DONE / FINALIZED

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
  rule in `<CompanyHealth>/draw_moves.json` (read by **`shared/draw_moves.py`** -
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

- Extend the Labor/Concrete sheets to RP and MFD if the PMs want them there.
- Roll the G702 contract source into the WIP readers so the master and the
  P&L can't disagree.
