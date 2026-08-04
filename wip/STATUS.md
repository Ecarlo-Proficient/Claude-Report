# wip/ — STATUS (shared progression record)

> Rule: update this file in the SAME commit as any change to this tool
> (CLAUDE.md/AGENTS.md structure rule 7). Tool matters only — no business
> findings, no dollar exposures, no owner-only analysis.

Last updated: 2026-07-29

## DONE / FINALIZED

- **DIVISION SOURCES OF TRUTH (the user 2026-07-31 — binding):**
  - **CP** — the G702 **draws** in each project folder (latest draw = contract,
    billed, retainage). Folder scan finds them; QBO only fills pre-Draw-#1 jobs.
  - **MFD** — whatever is on the **`WIP Master` tab** (contract col E, ETC col F).
  - **RP** — the **Schedule** derives the active jobs; the estimators' final
    numbers for that snapshot live in the owner's
    `OneDrive/RP WIP TO FIX_Final.xlsx` ('RP WIP' sheet), read via
    `--rp-from-file`. Billed/Costs always re-pulled from QBO.
- **CHANGE AUDIT ON EVERY SYNC (the user 2026-07-31 — never optional):** each
  run prints, **split by division**, (1) jobs added, (2) jobs removed,
  (3) ORIGINAL contract/ETC changes, (4) REVISED contract/ETC changes — and
  writes the same to `~/Downloads/WIP Changes.xlsx` (one file, overwritten in
  place). Baseline = the tab's own previous contents, read before the wipe, so
  it reflects what the owner last saw. Original contract = TOTAL CONTRACT PRICE
  − APPROVED COs; original ETC == revised ETC until CO costs have a source.
  Run on the Test-Master write (it carries all three divisions).
- **Owner-typed NOTES survive the sync** (2026-07-31): NOTES segments that
  don't match the script's own note/flag vocabulary (`_SCRIPT_NOTE_RE`) are
  harvested before the full-replace and re-attached to the same job, alongside
  the existing cell-comment preservation.

- **WIP TAB FORMATTING IS FROZEN** (2026-07-31, the user: "you cannot keep
  messing with formatting") — the original `WIP Master` sheet is the ONLY
  reference; read it, copy it, never invent. Title block is now its two-line
  LEFT-aligned form (B1 `<company> - <REPORT NAME>` read from `WIP Master`!B1
  at runtime, B2 `REPORT DATE: …`, medium rule above/below) — the
  merge-and-center 18pt banner is GONE. Full rule: repo CLAUDE.md rail 5a.
  Changing these tabs' look needs the user's explicit ask first.
- **One commentary column** (2026-07-31): NOTES and FLAGS merged into a single
  `NOTES` column (`_notes_all` = the owner's ACTION text · script notes ·
  must-fix flags, de-duplicated; yellow/italic when it carries a flag) on
  every tab.
- **RP TYPE (Tract/Custom) restored** (2026-07-31): `read_rp_from_file` derives
  it from the builder against `TRACT_CLIENTS`/`TRACT_CODES` (the file carries
  both full names and GL codes); CAMDEN HOMES added to `TRACT_CLIENTS` per the
  team's 2026-07-22 finding. The row-description column is now `CATEGORY`, so
  `TYPE` keeps its Tract/Custom meaning.
- **CATEGORY is decided from the DATA, not the band** (2026-07-31, the user:
  "rp7234-ftw is not good … there are no costs"): GOOD requires QBO costs or
  billing; no activity + on the schedule = NOT STARTED; no activity + not
  scheduled (FTW) = FTW BACKLOG. A line with no ETC gets a "No budget (ETC)"
  note.
- **Test - CP opens Active-only** (2026-07-31): `cp_wip_reader` now writes with
  `default_filter_active=True` (+ title block and TOTALS block, like the other
  tabs) so Closed rows are filtered and hidden on open.

- **`master_wip_test --rp-from-file <xlsx>`** (2026-07-29): the RP section of
  Test-Master comes from the owner's verified RP WIP workbook ('RP WIP' sheet)
  instead of the General List pipeline. Sections map from the file's band rows
  (main table → RP SLAB / FTW — ACTIVE by suffix; DROPPED, UNBILLED;
  FTW — OFF-SCHEDULE (COSTS); FTW BACKLOG). CP lines in the file are EXCLUDED
  (CP comes from the folder scan). Duplicate job lines dedupe to the first
  copy — flagged red when the copies disagree. Billed/Costs still refresh
  from QBO per line (file value survives only a failed lookup); the locked
  backlog rule applies post-QBO (any activity ⇒ out of FTW BACKLOG).
- **Test tabs match the real 'WIP Master' formatting** (2026-07-29): Tahoma 8
  everywhere (headers, data, TOTALS, cash-flow), master currency format
  `"$"#,##0_);[Red]("$"#,##0)` (no cents), `0.00%` percents.
- **QBO-only links are now the DEFAULT everywhere** (2026-07-29 pm, the user:
  file hyperlinks weigh the workbook down): `write_test_cp(qbo_links_only=True)`
  is the default; PROJECT FOLDER + DATA SOURCE columns removed from CP.COLS
  itself. The only hyperlinks on any test tab are Billed → QBO customer page
  and Costs → QBO project P&L. Pass `qbo_links_only=False` to restore the
  full click-to-verify link set.
- **'Test - RP' = the RP rows of Test-Master, same layout** (2026-07-29 pm):
  with `--rp-from-file`, master_wip_test also writes the RP rows to
  'Test - RP' in the master column set — revised TOTAL CONTRACT PRICE
  (contract + COs) and ESTIMATED TOTAL COSTS on every row, consistent across
  the board. Supersedes the rp_wip_simple layout on that tab.
- **'Test - RP' typed + legend + owner marks** (2026-07-31): the RP source is
  the owner's LIVE OneDrive file (`RP WIP TO FIX_Final.xlsx`); every row gets
  a TYPE column (GOOD / FTW WITH COSTS / DROPPED OFF SCHEDULE / FTW BACKLOG,
  from the file's bands; backlog reclasses to FTW WITH COSTS on QBO activity)
  with a legend block under the banner (types + colour meanings, single-font
  cells — no rich text). The owner's ACTION notes land in NOTES; his colour
  marks (green=verified / red=changed / orange=verify) are re-applied to the
  $ cells on BOTH RP-carrying tabs, and an owner-marked Billed/Costs value
  survives the QBO refresh (`qbo_protect`). BUILDER column restored.
  `write_test_cp` grew `legend=` and per-row `cell_marks` support;
  `_find_header_row` scans 15 rows (legend pushes the header down).

- **Test-Master is the deliverable WIP report** (2026-07-16): "WIP REPORT as of
  <date>" banner, rows 1–2 reserved as logo space (embedded images survive every
  sync), TOTALS row on live `SUBTOTAL(109,…)` (re-totals with the table filter),
  FUTURE WIP CASH FLOW block derived from the TOTALS row.
- **Identifiers are grab-able** (2026-07-16): PROJECT #/NAME cells are plain
  text; folder + data-source links moved to their own PROJECT FOLDER /
  DATA SOURCE columns. OVERBILLINGS/UNDERBILLINGS headers shortened to fit.
  FLAGS states the classify reason on red rows (never "OK" on red).
- **Excel 'repair on open' — THIRD cause found & fixed: stale sheet AutoFilter**
  (2026-07-31): a worksheet-level `<autoFilter>` cannot coexist with the Table's
  own filter — Excel calls the workbook damaged. 'Test - RP' still carried
  `A2:L69` from the old 12-column rp_wip_simple layout; the cell wipe in
  `write_test_cp` never reset it, and openpyxl re-derived the hidden
  `_xlnm._FilterDatabase` defined name from it on every save. Fix:
  `ws.auto_filter.ref = None` before the write (kills both), plus a `_qc_check`
  tripwire that fails loudly if the two ever coexist again. Verified on the
  saved file: no sheet AutoFilter, no `_FilterDatabase`, table headers match
  their tableColumn names, no merge overlaps, no rich text (inline strings).
- **Excel 'repair on open' eliminated — no rich text, clean links** (2026-07-21):
  two causes, both fixed & verified by opening in Mac Excel (no dialog, no
  'Repaired' title). (1) Hyperlink sheet-jumps (`'Small Jobs'!C7`) were
  appended to the file URI, putting raw spaces in the .rels target — invalid
  XML; `_apply_hyperlink`/`_link` now write a clean URI + the jump as
  `location=`. (2) openpyxl's multi-run inline rich text ('String properties'
  repair) is gone: the preview carries NO rich text at all — AR (orange) and
  JR (blue) NEEDS are now SEPARATE single-font columns, the legend is two
  plain cells, breadcrumbs are plain hyperlinked text.
- **User cell comments survive every sync** (2026-07-16): harvested by
  (PROJECT #, header) before the full-replace, re-attached after; a comment
  whose line left the tab prints loudly instead of vanishing.
- **General List AF = OTHER excludes the flatwork scope** (2026-07-16): POUR
  FLATWORK col AF "OTHER" ⇒ another contractor won it — no -FTW line even when
  priced, no flat $ in CP-standalone sums; slab line stays with a note.
- **`master_wip_test --rp-existing-only`** (2026-07-22): MFD (WIP Master tab)
  and CP (folder scan) refresh billed/costs fully from QBO; RP is locked to
  PROJECT #s already on Test-Master so a sync updates their numbers without
  adding new RP jobs from the General List. `existing_project_nums()` reads
  the tab first.
- **`rp_schedule_wip_preview --commit`** (2026-07-23): splits schedule-active jobs —
  READY (contract+budget both known) are written to the **'Test - RP'** tab as the new
  RP WIP report (QBO-enriched, banner + TOTALS/cash-flow, FTW backlog appendix); MISSING
  (lacking contract and/or budget) go to a clean ~/Downloads/RP WIP - Missing.xlsx so the
  gaps are obvious. RP7535 budget taken from the General Lista (GL_ETC_JOBS). 7-23 run:
  READY 35 · MISSING 31 · backlog 38. Verified on a scratch WIP copy before the prod write.
  Missing list also surfaces the RED 'in General Lista, in progress, NOT on schedule'
  jobs at the top (was silently dropped in --commit; RP7613 caught 2026-07-24).
  Consolidated to ONE Downloads file 'RP WIP.xlsx' (the user 2026-07-24): sheet
  'WIP (AUDIT)' = jobs in the WIP with original contract / budget / billed / costs to
  date, $ cells link to the source file, FROM columns to its folder; sheet 'MISSING' =
  the red not-on-schedule catch + schedule jobs missing contract/budget. Retired the
  separate Missing + Justification downloads.
- **RP done-rule + FTW backlog model** (2026-07-14, unchanged): billing is the
  truth; backlog = -FTW with no QBO activity and not on today's schedule.

## IN PROGRESS

- **Schedule-driven RP method — preview stage** (`one-offs/rp_schedule_wip_preview.py`):
  Main Schedule tab = active-jobs truth (the General List lags it); contract =
  bid proposal **PDF only** (signed doc; no takeoff bid-sheet substitution);
  ETC = takeoff cost sheet's own subtotal cells (side-scope files whole-sheet,
  base files SL+PR vs FW; items-vs-subtotal mismatches flagged). Output: one
  audit xlsx in Downloads (NEW / CHANGED / MATCHES vs the GL), yellow = GL
  numbers, green = source-doc numbers. Estimator-facing (2026-07-21 pm):
  "Budget" wording throughout (not ETC); NEEDS items are one-per-line
  instructions color-coded ORANGE = bid-proposal/contract actions (AR) and
  BLUE = budget-takeoff actions (JR), legend on row 2; $ cells open the
  source, file cells open the folder with a `CURRENT PROJECTS > …`
  breadcrumb for Windows users. AR/JR fix a wrong/missing file by
  editing the PROPOSAL PDF / TAKEOFF FILE cell in place — delete it, paste
  the correct path (the user 2026-07-22). GL is labeled 'General Lista';
  Δ CONTRACT/Δ BUDGET columns dropped as clutter. Commercial-takeoff
  fallback: a workbook with a 'BID' sheet (CP PM template, budget in
  AP1948/AP1961) wins over 'JobTread Cost Gral' for slab scope — RP6586
  pattern, incl. takeoffs not named with the RP#.
  **NOT wired into rp_wip_reader yet — awaiting the user's approval.**
  Round-2 (2026-07-22): (a) team-corrected file paths harvested to
  `one-offs/rp_source_overrides.json` (30 jobs; Windows I:/Z: → /Volumes/Common;
  fixes wrong folders/builders/typos) — overrides win over the folder guess;
  (b) TRACT builders (Camden/Grand Homes/Habitat) no longer false-flag 'no
  proposal' — contract from P.O.'s / General Lista; (c) CLEANUP CHECKLIST sheet
  lists takeoffs missing the 'JobTread Cost Gral' sheet (17). Two General-Lista
  cross-checks added: RED 'in General Lista, not on schedule' (0 today — all 27
  in-progress slabs are scheduled) and 'FTW BACKLOG (GL, not scheduled)' (38) +
  'FLATWORK TAKEN BY OTHER' (AF=OTHER).
- **JobTread as a budget source — verified, not wired**: Pave API grant key in
  the shared vault (`JT_GRANT_KEY`); schema proven (`one-offs/jobtread_probe.py`):
  approved `customerOrder` document → price = CONTRACT, cost = ETC; line names
  are the QBO item cost codes (SL1…) so budget-vs-actual per cost code is a
  direct join. Coverage gap: newest RP74xx–75xx cohort mostly absent/empty in
  JT. Planned source chain: JT approved proposal → proposal PDF/takeoff → GL.

## TO DO

- [ ] Wire the source chain (JobTread → PDF/takeoff → GL) into the preview so
      the sources can be compared per row; then, on approval, into
      `rp_wip_reader.py` itself.
- [ ] Coverage nag: flag schedule-active jobs missing a JT approved proposal
      and/or a priced proposal PDF (the orange NEEDS flags are the seed).
- [ ] Guard `read_general_list()` against suffixed `RP####-FTW` rows in col C
      (latent: regex strips the suffix and dedupe would silently merge — zero
      such rows exist today, verified 2026-07-21).
- [ ] Change-order staleness: a JT/PDF contract with QBO billed above it should
      flag "CO not entered at the source" (billed-over-contract red already
      exists; wants a source-aware message).
- [ ] Decide the fate of the WHY (TEMP) column on Test - RP once the
      justification workbook workflow settles.

## OPEN ISSUES

- **Running `rp_wip_reader.py` standalone REGRESSES `Test - RP` (2026-08-04).** A daily-list run
  of this reader rewrote the tab in its own older layout — header back on **row 1**, `WHY (TEMP)`
  and `CLIENT` columns returned, **no `APPROVED COs` column at all**, no title block, and the
  `TestRP` table ref left at `A1:Z90`. `Test - CP` (`A3:Y51`) and `Test-Master` (`A3:X132`) were
  untouched and correct, which is what made it look like an RP-only table-range bug. It isn't —
  it is a **layout collision**: several tools can write `Test - RP`, each with its own column
  definition, and whichever runs last wins.
  - Root fix in flight: `_rp_cols()` realigned to the standard `CP.COLS` set (keeps
    `APPROVED COs`, drops the `WHY (TEMP)` file:// link the QBO-links-only rule disallows).
  - Still open: the two `one-offs/` scripts that also write this tab can regress it the same
    way. Retire them or point them at a scratch tab.
  - **The daily checklist in the vault still names this script as the item-2 command**, so the
    morning run reproduces the regression until the realignment lands.
- **`_qc()` miscounts rows on any tab with a bottom legend** (found 2026-08-04). It counts every
  row with a value in the project column, so the COLUMN GUIDE block is counted as data — that is
  the `rows 135 ≠ expected 117` half of the RP warning, and it pushes `last_data` past the table
  end so the span check fails independently of the real layout problem. Stop the count at the
  first blank row after the data block, or exclude the guide range explicitly.
- Schedule→General List lag is real and recurring (26 schedule-active lines
  missing from the WIP as of the 7-20-26 schedule run; was 20 on 7-17-26).
  Process owner needed for entering new jobs in the GL (or the schedule method
  replaces the GL as the active-list source).
- Several active jobs have takeoff cost > proposal price (negative implied
  margin flags in the preview) — usually a stale proposal PDF vs a grown
  takeoff; estimator review needed per orange/blue NEEDS flags.
- `-FTW` scopes on the flatwork schedule with no flatwork-named proposal PDF or
  takeoff anywhere (pricing exists in no readable source).
- **Team fixes round 1 (RP WIP Fixes.xlsx, 2026-07-22)** — the estimator/OM worked every
  preview row. Findings that reshape the RP model:
  (a) TRACT builders (Camden, most Grand Homes, Dallas Area Habitat) have NO bid proposal —
      contract+cost come from **P.O.'s**. The 'no proposal PDF' flag is a FALSE ALARM for
      tract; tract needs a PO-based price source (JobTread won't cover them either).
  (b) Missing 'JobTread Cost Gral' budget sheet is the #1 data gap (~16 jobs) — blocks both
      takeoff extraction and JobTread coverage. Top cleanup item.
  (c) Some prices are un-fileable by design: budget-in-head, priced-by-sibling-street,
      per-sqft, or 20%-profit allocation.
  (d) RP7083 split across two customers in the books — will scramble billed/costs.
  Validation: ~16 negative/low-margin flags → 'Takeoff formulas updated' (real errors caught).
  The AR/JR columns + edit-path-in-place workflow were used exactly as designed.
- Excel cannot gate hyperlinks behind ctrl/cmd-click (asked 2026-07-16;
  answered with the plain-text identifier columns + click-and-hold).

## SIDE TOOL — JobTread coverage (2026-07-22)
- `one-offs/rp_jobtread_coverage.py` — read-only estimator to-do: every schedule-active
  job × JobTread status (MISSING / NEEDS PROPOSAL / COVERED, with contract+budget on the
  covered ones) → ~/Downloads/RP JobTread Coverage.xlsx. Nothing pushed. First run
  (schedule 7-22): 15 covered (23%) · 17 needs proposal · 33 missing. Closing this list =
  closing the JobTread coverage gap (the sole requirement for JT-as-pricing-source).

## NOTES

- 2026-07-30 — data-risk scrub: real dollar figures removed from code comments
  (`cp_wip_reader.py` retainage note) — no behavior change; the dollar detail
  lives in the owner's vault, per the STATUS scope filter.

## FIXED 2026-08-03 (pm) — layout + NOTES defects the owner caught

- **Stale row HEIGHTS** left tall blank rows: the header/banner row moves
  between layouts (Test - CP row 1 kept 30pt from when it WAS the header;
  Test - RP row 1 kept the retired banner's 34pt and row 11 the old header's
  30pt). The rewrite reset `hidden` but never `height`. Now both reset, so
  only the current header row carries a height.
- **NOTES duplicated every sync**: the owner's ACTION text is ONE string that
  already contains ' · ', while the carried-forward prior cell was split into
  segments — whole-string de-dup never matched, so each sentence came back
  twice. De-dup is now per SEGMENT on a whitespace/case-normalised key.
  The redundant leading `note:` label is dropped (the column is called NOTES);
  wording is otherwise never altered.
- **Notes no longer resurrect**: rows whose NOTES come from a source file the
  owner edits (RP, `notes_from_source=True`) are excluded from the
  carry-forward — otherwise deleting a note from his RP file would put it
  straight back. Carry-forward still protects hand-typed CP/MFD notes.

## 2026-08-03 (pm 2) — report layout brought up to the owner's reference WIP

The owner supplied a reference WIP (another GC's) and asked for that level of
finish. Delivered, WITHOUT renaming any header (five tools read this tab by
header name — `shared/schedule.py`, `shared/breakeven.py`,
`health-dashboard/company_dashboard.py`, `money_bleeds.py`,
`project-pnl/project_pnl_export.py`):

- **Columns grouped and reordered** CONTRACT → BUDGET → COSTS → PROFIT →
  BILLING → REMAINING → ANALYSIS (`_COL_GROUPS`). New `ORIGINAL CONTRACT`
  (A) and `PERCENT BILLED` (K) columns; `APPROVED COs` (B) now sits directly
  after the contract it changes, with `TOTAL CONTRACT PRICE` (C = A+B) next.
  A medium vertical rule opens each group so they read as boxes.
- **COLUMN GUIDE block at the bottom** — every money column with its letter and
  derivation (C = A+B, G = F÷D, …), grouped like the columns.
- **Legends moved to the BOTTOM** (`legend=` now renders under the report, not
  above the header) — keeps the header at row 3 where the readers expect it.
- **Cash-flow block reformatted**: label merged across A:B, amount across C:D
  immediately beside it, every cell bordered, title bar filled. Previously the
  amounts sat in the contract column, stranded far from their labels with no
  rules at all.
- **Divisions spelled out** in SECTION: `Multi-Family`, `Commercial`,
  `Residential — Slab / Flatwork / Flatwork (off-schedule) / Dropped, Unbilled
  / Flatwork Backlog` (`_SECTION_LABEL`).
- **NOT DONE — the group header BAND across the top.** `shared/schedule.py`
  finds the header by scanning rows 1–5 for any cell containing "CONTRACT"; a
  band row saying CONTRACT would hijack that and silently break the schedule
  linkage (and company_dashboard). Adding the band needs those readers to
  match `PROJECT #` exactly first — do that before revisiting.

### Fixed here
- `money_bleeds.check_rp_wrapup` read 'Test - RP' headers from **row 1**, so it
  silently returned nothing once that tab gained a title block. It now finds
  the header row (scans rows 1–15 for `PROJECT #`).

## 2026-08-03 (pm 3) — symmetric contract/budget columns; MFD budget-loss bug

Researched the real WIP standard (Foundation Software, surety template,
Northstar, Dean Dorton). A change order moves BOTH sides, so the columns now
mirror each other:

    ORIGINAL CONTRACT A · APPROVED COs B · TOTAL CONTRACT PRICE C = A+B
    ORIGINAL ESTIMATED COST D · CO COSTS E · ESTIMATED TOTAL COSTS F = D+E

`TOTAL CONTRACT PRICE` IS the revised contract and `ESTIMATED TOTAL COSTS` IS
the revised estimated cost — those names are read by five other tools and must
never be renamed. Order follows the researched standard: CONTRACT → BUDGET →
PROFIT → COSTS → EARNED → BILLING → POSITION → ANALYSIS. `PERCENT BILLED` was
dropped (it appears in no standard; it was invented here).

**CO COSTS has no source and is written EMPTY on purpose** (the user
2026-08-03: "the empty cell is loud enough" — no flag). It is styled as a
yellow INPUT cell, so a blank reads as "nobody costed this CO". Five jobs
currently carry approved COs with no CO cost, which overstates their profit;
closing that needs a cost line on the CO template, or a PM cost-to-complete
re-forecast (the industry answer, which also absorbs overruns).

### Bug the change audit caught on its first live run
`read_mfd_from_master` read the ETC cell with `data_only=True`. That cell is a
FORMULA on 'WIP Master' (`=(E4/1.17)` — contract ÷ markup), and **openpyxl
strips every cached formula result workbook-wide on save**, so from our first
write onwards MFD's entire budget read as None — silently blanking the budget,
percent complete, earned revenue and profit on ~$16.8M of MFD contract. The
reader now falls back to evaluating the sheet's own divisor formula; recovered
values match the pre-loss figures exactly. Any future WIP-Master-sourced cell
that is a formula needs the same treatment.

## 2026-08-03 (pm 4) — Test-Master is the finished report

- **LEFT TO BILL joined the BILLING box** — it is a billing position, not its
  own section, so the thick rule after UNDERBILLINGS is gone and the next one
  opens at FUTURE PROFIT TO EARN. `_COL_GROUPS` lost the POSITION group.
- **Test-Master drops LAST SYNCED and NOTES** (the user 2026-08-03: the master
  "should be the final visually perfected wip report as is"). The report date
  on row 2 already carries the timestamp. BOTH stay on 'Test - CP' and
  'Test - RP', which are the working views — and `money_bleeds` reads both of
  them from 'Test - RP', so removing them there would break it.
- **Known/accepted notes are muted, not shown** (`_MUTED_NOTE_RE`): the first
  is "proposal quotes PIERS but no PR cost in the takeoff" — pier costs sit in
  the Piers takeoff sheet's overall costs and were never broken out per code,
  so it is a description of how the takeoffs are built, not a finding. Add to
  that regex when another accepted condition is identified; genuine takeoff
  errors (e.g. "pier cost row(s) are #N/A") are NOT muted.

## 2026-08-03 (pm 5) — owner edits auto-colour, survive syncs; Test-Master locked

The owner edits the WIP directly. Three guarantees, all verified end-to-end
(edit → sync → sync → still there):

- **Auto-colour on edit, with no macro.** Every editable input is mirrored into
  a HIDDEN baseline column (`«base» <field>`) holding what the DATA SOURCES
  say. Excel conditional formatting reddens any cell that differs from its
  baseline, so the cell colours itself the instant he types. `_OVERRIDE_FIELDS`
  lists the mirrored inputs.
- **His value wins and keeps winning.** `read_owner_edits()` uses the SAME
  comparison Excel used to colour the cell, so "what he sees marked" and "what
  the script keeps" can never drift. The baseline must stay the SOURCE value —
  baselining to his override made the cell match next run, the red cleared and
  the script silently restored the source value (found by testing two
  consecutive syncs, not one). Overrides are limited to sourced inputs; every
  derived column is an Excel formula off them, so `ORIGINAL CONTRACT` +
  `APPROVED COs` and `ORIGINAL ESTIMATED COST` + `CO COSTS` stay consistent.
  This also gives `CO COSTS` its first real source: the owner typing it.
- **Notes and comments never lost.** Cell comments were already harvested and
  re-attached. NOTES now has exact provenance: the baseline stores the
  SCRIPT-generated text only (captured before his lines are folded back in), so
  anything else in the cell is his and is carried. A note deleted at its source
  stays deleted (it was in the baseline) — the old `notes_from_source` blanket
  skip is gone, since it discarded his RP notes wholesale.
- **Test-Master is protected** (`protect=True`): it is a read-only roll-up of
  the CP/RP tabs and the WIP Master MFD section. No password — Review ▸
  Unprotect Sheet is one click — and filtering, sorting and selection stay on.
  The working tabs ('Test - CP', 'Test - RP') stay editable.

## 2026-08-03 (pm 6) — 'Test - RP' has FOUR possible writers; two guards added

A run of `rp_wip_reader.py` overwrote 'Test - RP' with its own stale layout —
header back on row 1, `WHY (TEMP)` + `CLIENT` columns, **no APPROVED COs column
at all**, legend gone. The edit-tracking machinery still applied (it lives in
the shared writer), but the tab's format regressed. Four tools can write that
tab: `master_wip_test --rp-from-file`, `rp_wip_reader`,
`one-offs/rp_schedule_wip_preview` and `one-offs/rp_wip_simple` — last run wins.

- **`_rp_cols()` realigned** to the standard set: keeps `APPROVED COs` (a change
  order is half the contract story), drops `WHY (TEMP)` (a file:// link into a
  Downloads workbook, which the QBO-links-only rule no longer allows), and
  labels the client column `BUILDER` like every other tab. Running the reader
  can no longer regress the format.
- **Baseline TRUST CHECK in `read_owner_edits()`**: the baselines are believed
  only when the tab was last written by the SAME layout. After a foreign write
  the numbers under our headers came from another pipeline, so every difference
  looked like an owner edit — the restore run reported **28 phantom edits** and
  applied them as overrides. They happened to match source values so nothing
  was corrupted, but that was luck. Detection now skips the run and says so.
- **STILL OPEN:** the two `one-offs/` RP writers can each still clobber the tab.
  Retire them or point them at a scratch sheet — owner's call.

## 2026-08-04 — CP contract for a pre-draw job: PDF → takeoff

`CP910` showed a blank contract: no first draw, and its takeoff has BOTH a
Commercial and a Residential proposal tab, so `_select_proposal_sheet` saw two
competing proposals and refused to guess. Source order for a no-draw CP job is
now (the user 2026-08-04): **draw → signed proposal PDF → takeoff**.

- **`shared/proposals.py`** (new, in `shared/` because CP and RP both need it —
  tools never import tools): finds the folder's proposal PDF and reads its
  overall total. pdfplumber splits digits ("$ 1 05,815.00"), so whitespace
  inside a number is stripped before parsing.
- **Residential proposal tabs are NEVER read for CP** (the user: "if CP use the
  commercial 100%"), and with several non-final proposal tabs the Commercial
  one wins outright. That alone un-blanks CP910's takeoff path.
- **Cross-verified both ways:** agreement between PDF and Commercial tab is
  stated in NOTES ("matches the Commercial Proposal tab ✓"); disagreement is a
  must-fix flag naming both figures.

### Two guards, both added after the change audit caught a regression
The first cut of this dropped **CP783 from $364,032 to $56,544** — the audit
caught it on the very next run:
- **Several proposal PDFs ⇒ don't guess.** CP783 holds a main proposal, a
  breakout, a dirt/utilities proposal and a revised dirt proposal. They are
  different SCOPES, not revisions of one price, and "newest wins" picked the
  $56k dirt scope. With more than one candidate the PDF path is abandoned and
  the takeoff is used, with the candidates named in NOTES.
- **A section subtotal is never the contract.** That dirt PDF's only figure was
  `Eartwork Total: $56,544` with no overall `TOTAL:` line. `grand_total()` now
  returns None when it finds only qualified section totals.
