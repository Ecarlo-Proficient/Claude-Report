# wip/ — STATUS (shared progression record)

> Rule: update this file in the SAME commit as any change to this tool
> (CLAUDE.md/AGENTS.md structure rule 7). Tool matters only — no business
> findings, no dollar exposures, no owner-only analysis.

Last updated: 2026-08-25

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
  - Format drift fixed: `_rp_cols()` realigned to the standard `CP.COLS` set (keeps
    `APPROVED COs`, drops the `WHY (TEMP)` file:// link the QBO-links-only rule disallows).
  - **The script is still unsafe to run, and now for a worse reason (the user 2026-08-04).**
    Run as `__main__` it resolves RP from `ALPHA_PATH` — the **General List** — while the
    binding RP source of truth is the estimators' verified `RP WIP TO FIX_Final.xlsx`, which
    only `master_wip_test --rp-from-file` reads. So it would overwrite verified numbers with
    General-List ones **in a now-correct-looking format**: a silent data regression that shows
    no visible symptom. Fixing the layout removed the warning sign, not the hazard.
  - Removed from the vault's daily checklist entirely; the daily WIP command is
    `master_wip_test.py --rp-from-file …` **plus** `cp_wip_reader.py` (the former does not write
    `Test - CP`; the mandatory change audit only fires on the Test-Master write).
  - Still open: the two `one-offs/` scripts that also write this tab can regress it the same
    way. Retire them or point them at a scratch tab.
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

## 2026-08-04 (later) — per-division runs restored: ONE RP implementation

The per-division split is deliberate (run one division on its own), but RP's
reader was on the WRONG SOURCE — the General List, not the estimators' verified
file — so it could not be used daily and was clobbering `Test - RP`. Realigning
its columns fixed the FORMAT only; the DATA source was still wrong.

Fixed by moving the RP source-of-truth logic OUT of `master_wip_test` and INTO
`rp_wip_reader`, where it belongs, so there is exactly one implementation:
- `RP_WIP_FILE` — the owner's file, resolved via `shared/paths` (env-overridable).
- `read_rp_from_file()`, `classify_from_file()`, `rp_tab_cols()`,
  `RP_TAB_LEGEND`, `write_rp_tab()`, `SECTION_LABEL`, `_rp_category`,
  `_owner_mark` — all now live in `rp_wip_reader`.
- `master_wip_test` calls those; `write_rp_tab()` is the ONLY writer of
  'Test - RP', so a standalone RP run and the unified run cannot drift.
- `rp_wip_reader.py` now DEFAULTS to the owner's file. The General List
  pipeline is still there behind `--general-list` for inspecting the legacy
  path; `--file <xlsx>` overrides the source.

**Verified:** ran both entry points back to back and diffed 'Test - RP'
cell-for-cell (162 rows × every column) — IDENTICAL apart from the LAST SYNCED
timestamp.

Daily commands (either shape works now):
  `master_wip_test.py --rp-from-file <xlsx>` + `cp_wip_reader.py`   (all tabs + audit)
  or per division: `cp_wip_reader.py` / `rp_wip_reader.py`
Only `master_wip_test` writes Test-Master and runs the change audit.

## 2026-08-04 — CP861 had no change orders: a STALE DRAW, not a dropped check

The CO check was never removed. `co_revenue` still reads G702 line 2 from
whatever draw is found — the bug was that the wrong draw was found.
`find_latest_draw()` descended exactly ONE level into a numbered `Draw #N`
folder. CP861 files its workbook a level deeper, under a per-company subfolder
(ours and the GC's sit side by side), so draws #4 and #5 produced no candidate
and the reader silently settled on **draw #3**:

    draw #3 (what we read)  net CO $0        billed $410,520
    draw #5 (actual latest) net CO $52,576   billed $605,426

So the job was understated by **$52,576 of change orders and $194,906 of
billing**, with nothing on the report to say so — the reader had a valid draw,
just an old one.

- `_draw_workbooks()` now walks DOWN to depth 3 inside a numbered draw folder,
  skipping `DO NOT USE` / SUPERSEDED / OLD / ARCHIVE / VOID / BACKUP folders
  (the team parks retired paperwork there — CP861 draw #5 has one). Shallowest
  match ranks first, and `has_g702()` still gates every candidate, so the GC's
  own spreadsheets can never win.
- Re-checked every CP job: all still resolve, several now to a deeper (correct)
  draw. CP861 is on draw #5 and the audit recorded the correction.

### CP765 — a REVISED draw sits beside the original it replaces (2026-08-04)
Second stale-draw failure the same day, different cause: **two workbooks for the
SAME draw number** in one folder. `Draw #4/` holds both `Draw Excel #4` and
`Revised LP Draw Excel #4`. Draw # tied, so the winner fell out of filesystem
iteration order — arbitrary — and the superseded original won:

    original draw #4   net CO $79,752   billed $529,054
    revised  draw #4   net CO $66,200   billed $515,502   ← the live one

The revision backs a **$13,552** change order out. Reading the dead file invented
a $13,552 shortfall against QBO invoice 34288, whose cumulative $515,502 ties to
the revised draw **to the dollar** (the user 2026-08-04 — flagged the file and
the invoice).

- `_supersedes()` now breaks ties WITHIN a draw number in `shared/draws.py`:
  a name matching `revis` (Revised/Revision) beats the original, then newest
  mtime. `has_g702()` still gates every candidate.
- CP765 is the only project whose chosen workbook changes. **Its WIP row moves
  when the reader next runs**: contract $635,578 → $622,026, CO $79,752 →
  $66,200, billed $529,054 → $515,502. Numbers going DOWN here is the correction,
  not a regression.
- Money Bleeds `Draws CP` goes 1 RED → 0 RED as a result.

### CP800 — both takeoffs ARE summed; FDT has no cost in it
`_select_takeoffs` already includes and sums both (`FDT TAKEOFF - WIP.xlsx` +
`PAVING TAKEOFF - WIP.xlsx`) — both contracts add into the row. The ETC is
PAVING's alone because **FDT's cost cells are empty**: `BID!AP1948` and
`BID!AP1961` are both literally 0, and no other cost roll-up exists in that
file. Nothing to add until an estimator fills it; picking another number out of
that workbook would be a guess.

### shared/proposals: 'GRAND TOTAL' is the overall figure
It was being classified as a section subtotal (the leading word looked like a
scope), so a PDF whose total is labelled `GRAND TOTAL` returned nothing.
`_OVERALL_WORDS` = GRAND / PROJECT / CONTRACT / OVERALL / BID.

## 2026-08-04 — the two roll-ups are LIVE formulas on the tabs the owner edits

`TOTAL CONTRACT PRICE` and `ESTIMATED TOTAL COSTS` were written as VALUES, so
typing a CO cost into `CO COSTS` left the revised total unchanged until the next
sync (the user: "etc total are not formulas so if i edit the og etc and the
revised co etc it doesn't add up").

They cannot simply become formulas everywhere: an openpyxl-written formula
carries **no cached value**, so every `data_only=True` reader gets None — the
same failure that blanked MFD's budget. Four tools read Test-Master that way
(`shared/schedule`, `company_dashboard`, `money_bleeds`,
`project_pnl_export`). So the behaviour is split by what the tab is FOR:

- **Test - CP / Test - RP** (`live_formulas=True`) — the tabs he types into:
  `C = IF(AND(A="",B=""),"",A+B)` and `F = IF(AND(D="",E=""),"",D+E)`. Edit an
  input and the roll-up updates on the spot; both blank still renders blank,
  not a false $0. Being calculated, they lose the yellow input fill.
- **Test-Master** — values, unchanged. It is LOCKED, nobody types in it, and it
  is the tab every other tool reads.
- `money_bleeds.check_rp_wrapup` reads TOTAL CONTRACT PRICE off 'Test - RP', so
  it gained a component fallback (ORIGINAL CONTRACT + APPROVED COs), mirroring
  the one it already had for LEFT TO BILL. Verified: the fallback fires on all
  110 RP rows and the check still reads the tab.

## 2026-08-04 — the SHARED ENGINE split out of cp_wip_reader → wip_writer.py

Pure structural cleanup, NO behaviour change. The report engine — CpRow,
COLS, write_test_cp, all formatting, the change audit, the owner-edit
baselines, the QC check — used to live inside `cp_wip_reader.py`, so every
division tool did `import cp_wip_reader as CP` to reach it. That is a tool
importing a tool (repo rule forbids it), it buried MFD/RP logic in a file
named "cp", and it let the layout drift (a stray rp_wip_reader run once
regressed the tab).

- **`wip/wip_writer.py` (NEW)** — the engine, its own honest home. Pure output
  machinery: it never reads a division's source. Owns `QBO_REALM` (a reader's
  enrich sets `wip_writer.QBO_REALM`).
- **`cp_wip_reader.py`** — now ONLY the CP reader (folder scan / draws /
  proposal PDF). Imports the model + helpers from wip_writer; calls the writer
  as `W.write_test_cp`. It does NOT re-export the writer, so no tool can reach
  the writer through cp_wip_reader again — the shortcut that caused the tangle
  is structurally closed.
- **`rp_wip_reader.py`** — imports `wip_writer`, no longer `cp_wip_reader`. A
  reader importing the engine, never another reader.
- **`master_wip_test.py`** — the orchestrator; imports the engine (`W`) for the
  writer/audit AND `cp_wip_reader` (`CP`) for the folder scan. Legitimate: it
  is the orchestrator, not a peer tool.
- Two one-offs (`rp_wip_simple`, `rp_schedule_wip_preview`) repointed to
  `wip_writer`. Only `master` still imports `cp_wip_reader`.

VERIFIED behaviour-preserving: snapshotted all three tabs (value/formula/number
format/fill/font per cell) before, regenerated after, diffed — Test-Master and
Test - RP are byte-identical (0 differing cells). pyflakes: no undefined names,
no unused imports across all four files. Runtime NameErrors the compiler
couldn't see (ALLOWED_WRITE_SHEETS, get_column_letter) were caught by the
regenerate run and fixed. Test - CP verified after the share remounted: header at row 3, full
symmetric column set (incl. CP-only RETAINAGE HELD) + the «base»
edit-tracking columns, no corruption triggers on any tab.

## 2026-08-06 — RP reader resolves columns BY HEADER (owner rearranges his file)

The owner added JOBTREAD + STATUS columns to `RP WIP TO FIX_Final.xlsx` (next to
GENERAL LIST), which shifted ACTION/CO right. `read_rp_from_file` read by FIXED
column position, so a column move silently misaligned contract/ETC/action — the
same fragility that has now bitten twice. Fixed: `_resolve_rp_cols()` locates
each field BY HEADER NAME (row 2, case-insensitive, fixed-position fallback),
so future column adds/moves in the owner's file no longer break the sync.
Verified: the reader auto-resolves ACTION→13, CO→14 on the new layout and parses
all 119 lines (110 + 9 new jobs).

Note: the JOBTREAD (✓ in JobTread / ✗ not) and STATUS (Needs ETC · Pull ETC from
JobTread · Add JT proposal · Add to JobTread · Ready) columns are an estimator
aid the reader ignores; they are a point-in-time snapshot, refreshed by re-running
the build, not live.

## 2026-08-06 — Test-Master is the clean BANK report (`plain_report`)

Test-Master is what the owner sends to banks, so it is now the lean, plain view
(the user 2026-08-06). `write_test_cp(plain_report=True)`, used ONLY for the
Test-Master write, drops everything colour/working-tool: no coloured fonts (no
red review, no blue links, no owner green/red marks), no QBO hyperlinks, no
yellow input fills, no medium group rules, no edit-tracking baselines, no bottom
COLUMN GUIDE. Kept: grey header, thin grid, TOTALS + cash-flow summary.

`master_cols()` for the bank report: the first column is the **division**,
labelled **TYPE** (was SECTION); the old Tract/Custom TYPE is dropped; the
change-order breakout is collapsed to one **TOTAL CONTRACT PRICE** + one
**ESTIMATED TOTAL COSTS** (ORIGINAL/CO columns gone). 19 columns.

`_BANK_EXCLUDE`: the bank report shows only clean active WIP — the
`FTW — OFF-SCHEDULE (COSTS)`, `FTW BACKLOG`, and `RP — DROPPED, UNBILLED`
sections are filtered OUT of Test-Master (they stay on the working 'Test - RP'
tab). Test-Master went 138 → 96 rows; the excluded rows show as a one-time
REMOVED block in the change audit.

Working tabs unchanged: 'Test - CP' and 'Test - RP' keep colours, the CO
columns, edit-tracking, and every section. Verified: Test-Master data cells
have ZERO coloured fonts / fills; Test - RP still has its 588 coloured cells.

## 2026-08-06 (later) — bank report: BONDED column, closed jobs excluded

- The Test-Master STATUS column is now **BONDED** = "N" on every job (the user
  2026-08-06). The bank report carries only active WIP, so the old Active/Closed
  STATUS was redundant; the slot now shows bonding, "N" for all (none bonded —
  a per-job source would replace the constant).
- Closed jobs are excluded from the bank report OUTRIGHT (not just filter-hidden)
  and `default_filter_active=False` — every row is active by construction.
- `_BANK_EXCLUDE_JOBS = {"RP6901"}` — billed out / done, dropped from the bank
  report (the user: "already billed out"). It stays on the working 'Test - RP'.
  Test-Master 96 → 95 rows. Verified: BONDED = {'N'} across all jobs, RP6901
  absent from Test-Master but present on Test - RP.

## 2026-08-07 — Test - RP is the LEAN "where is it at?" working view

The owner wants Test - RP leaner so project numbers are easy to spot (2026-08-07):
- **Dropped** REVENUES/PROFIT EARNED, OVER/UNDERBILLINGS, FUTURE PROFIT, PURE
  JOB BORROW, LAST SYNCED (the row-2 report date IS the sync date), and NOTES.
- **Added** three "where is it at?" columns pulled straight from the owner's
  RP file — SCHEDULE / GENERAL LIST / JOBTREAD (✓ green / ✗ red) — placed right
  after STATUS. `read_rp_from_file` reads the `gl`/`jt` marks; `_resolve_rp_cols`
  gained GENERAL LIST + JOBTREAD headers.
- **GP% over 30% is conditional-formatted** amber (`gp_highlight_over=0.30` →
  CellIsRule on the GROSS PROFIT % column) — a too-good margin usually means a
  missing cost.

Two lean-layout robustness fixes in the shared writer (a dropped column must
not crash): `_build_formula`'s eager refs use `.get` → `#REF!` for a missing
column (the RETURNED formulas never reference a dropped column); `_write_summary`
skips the FUTURE WIP CASH FLOW block when its earned/billing columns are absent
(TOTALS row still written). `money_bleeds.check_rp_wrapup` falls back to the
row-2 REPORT DATE now that LAST SYNCED left the tab.

Isolated to Test - RP: Test-Master (bank) and Test - CP unaffected.

## 2026-08-07 — Blank ETC now falls back to the takeoff (orange), manual = blue

Follow-up to the lean Test - RP. The GP%>30% highlight can't flag the rows that
matter most — a blank ETC makes ORIGINAL PROFIT (and so GP%) compute to blank,
so a missing-budget job never trips the rule. Fixed at the source:

- **`shared/takeoff_etc.py` (new).** The verified takeoff→ETC extractor
  (`find_takeoff_etc` + `_cost_sheet_totals` + name-scoring helpers) MOVED out of
  `one-offs/rp_schedule_wip_preview.py` the moment the WIP reader needed it — the
  one-off now re-imports the same names from shared (so `P.find_takeoff_etc` /
  `P._norm` still resolve for the job-auditor prototypes). No logic change.
- **`rp_wip_reader.classify_from_file`** calls `fill_missing_etc_from_takeoff`
  first: for every row the estimator left ETC blank, it resolves the job folder
  (this module's own `index_residential` / `match_by_address`) and reads the
  budget from the takeoff cost sheet (SL+PR slab / FW flatwork / commercial BID).
  The estimator's manual entry ALWAYS wins — the fallback runs only on blanks.
- **ETC provenance colour on 'Test - RP' (font):** BLUE `0070C0` = estimator's
  manual entry, ORANGE `ED7D31` = machine-read from the takeoff (verify). Legend
  updated. Test-Master (bank report) suppresses these marks as before.
- Field key is `base_etc` (the writer column), not `etc` — the old owner ETC
  mark used `etc` and had been a silent no-op.

Live result on the 11 blank ETCs: **8 filled from takeoffs** (RP7507 $247,988,
RP7612 $111,160, RP7610, RP7621, RP7623, RP7622, RP7607, RP7118-FTW); **3 stay
blank + flagged** because their folder has only a proposal / no cost sheet
(RP7234-FTW, RP6766-FTW, RP6901 — a budget sheet must be added to the takeoff).

## OPEN ISSUES
- Test-Master carries the same fallback (shared `classify_from_file`) but was not
  re-run this session — regenerate it to pick up the 8 filled RP ETCs.
- 3 RP jobs need a `JobTread Cost Gral` budget sheet added to their takeoff before
  their ETC can be read.

## 2026-08-07 — Test - RP: marks moved to the end + stale columns cleaned

Meeting-prep QC (the user 2026-08-07):
- **SCHEDULE / GENERAL LIST / JOBTREAD moved to the VERY END** of Test - RP
  (after LEFT TO BILL), not after STATUS — `rp_tab_cols` appends them last.
- **Orphan tiny columns cleaned.** A narrower rewrite (the lean Test - RP, and
  the plain Test-Master) left old column widths floating to the right of the
  data — Test - RP had AF–AI (width 4–7), Test-Master had W–AA (one at width
  60). New `wip_writer._clear_stale_columns(ws, last_used)` drops every column
  dimension past the last real column on every tab; the spacer between the data
  and the hidden `«base»` baseline block is now hidden too. The `«base»` columns
  themselves are the edit-tracking baseline (hidden by design) — left as-is.
- QC verified: Test - RP / Test-Master / Test - CP all clean past their last
  column; `«base»` block + spacer hidden; workbook reopens without repair.

ETC takeoff fallback now fills 9 of 11 (RP6766-FTW resolved a takeoff this run);
2 still blank (RP7234-FTW, RP6901 — no cost sheet in the folder).

## 2026-08-07 — Test-Master: home type in TYPE + two more bank drops

- **Custom/Tract folded into the residential TYPE** (the user 2026-08-07: "add
  custom/tract before slab"): the bank label now reads "Residential — Custom —
  Slab" / "Residential — Tract — Slab" (and Flatwork). Done in the bank_rows
  relabel loop from `row.home_type`; MFD/CP unchanged. TYPE width 22 → 32.
- **RP6586 and CP585 dropped from the bank report** (added to
  `_BANK_EXCLUDE_JOBS` with RP6901). They remain on the working Test - RP /
  Test - CP tabs. Test-Master now 93 rows.

## OPEN ISSUES
- **RP6901-FTW still on Test-Master** with broken numbers (billed $46,662 on an
  $8,342 contract, cost 4.6× ETC, 465% "complete"). RP6901 (slab) is excluded
  but the -FTW twin is a distinct project # — exclude it too if the owner agrees.
- RP-file duplicate lines flagged: RP6938-FTW (numbers differ) and RP6858-FTW.
- RP7234-FTW still blank ETC (no takeoff cost sheet).

## 2026-08-07 — Fix: frozen-pane view corruption ("Repaired Records: View")

Excel repaired 'Test - RP' on open ("Repaired Records: View from
/xl/worksheets/sheet11.xml"). Cause: the reused tab accumulated INVALID
sheet-view selections — a spurious `topRight` (no vertical split) plus a
duplicate `bottomLeft` — that survived the cell wipe exactly like the stale
auto-filter the writer already clears. `write_test_cp` now forces a single
valid `bottomLeft` selection right after `freeze_panes`, so every tab writes a
clean view. The live file was repaired in place surgically (only the offending
`<selection>` XML rewritten; all other bytes preserved) and passed a full
corruption sweep (view / rich-text / merges / tables / style+dxf indices).

Note: the 'custom XML no longer supported' Excel notice is a pre-existing,
benign SharePoint leftover in the source workbook — not corruption.

## 2026-08-07 — Pre-write AUDIT (--audit): inspect before the WIP updates

The owner wants to verify the non-QBO parts before the report is written —
mainly the automatic add/remove-jobs logic, and where each contract/ETC comes
from (QBO billed/costs are trusted). New:

- **`wip/wip_audit.py`** (new, READ-ONLY): builds one plain workbook, one row
  per job — Δ vs the current report (ADDED / REMOVED / SAME) + the reason, and
  CONTRACT / ETC each with its source. Never touches the WIP file.
- **`master_wip_test --audit [path]`**: runs the full pipeline, writes the audit
  (`~/Downloads/WIP Audit.xlsx` default), and STOPS — no WIP write. Reads the
  current Test-Master for the prior job set (`W._snapshot_tab`), so REMOVED jobs
  say whether they hit a rule (`_BANK_EXCLUDE_JOBS`), left the source, or just
  dropped off the bank cut. Safe even with the report open in Excel.
- **Provenance captured at read time** in `rp_wip_reader`: each RP row now tags
  `audit_contract_src` ("RP file 'RP WIP'!row N · CONTRACT $") and
  `audit_etc_src` (estimator cell / takeoff file+cell / BLANK). CP/MFD fall back
  to a source label by division.

Verified with a synthetic set (ADDED/REMOVED-by-rule/left-source/SAME all render
with correct reasons + source cells). A live run needs the Common volume mounted
(CP folder scan) — it was unmounted at build time.

Workflow going forward: `--audit` → owner inspects the workbook → approve → real
run writes the report.

---

## 2026-08-25 — `Test - MFD`: an entry tab MFD owns, with a QBO block they can't edit

**New tool: `mfd_wip_test.py`.** MFD needed an easier way to put their numbers into the
WIP. The live `WIP - MFD` tab stays untouched; the script seeds an allow-listed
`Test - MFD` copy of it and adds six columns. The owner's intent (2026-08-25) is that
`Test - MFD` **takes over from `WIP - MFD`** once it is signed off.

**Layout — B..M are copied verbatim and never written again.** The script owns N..S only:

| col | header | who fills it | how |
|-----|--------|--------------|-----|
| N | `ETC` | MFD types it | grey/orange input style lifted from the live tab |
| O | `REVISED ETC` | MFD types it | seeded `=N<row>`; typed over when a CO moves the budget |
| P | `GP %` | formula | `(REV. CONTRACT - REVISED ETC) / REV. CONTRACT` |
| Q | `COSTS TO DATE` | QBO | green header, tinted cell, comment |
| R | `BILLED TO DATE` | QBO | green header, tinted cell, comment |
| S | `RETAINAGE (QBO)` | QBO | green header; comment carries the variance vs col M |
| T | `COST TO COMPLETE` | formula | `REVISED ETC - COSTS TO DATE` |

`Q5:S5` is a merged, centered banner carrying `QBO - LAST SYNC mm/dd/yyyy h:mm AM`.
GP% mirrors `WIP Master`!Q and cost-to-complete mirrors `WIP Master`!I so the two sheets
agree. The TOTALS block is extended across N..S, and a two-cell key sits below the table.

**FORMATTING — this tab is the rail-5a exception, on purpose.** Rail 5a freezes the
GENERATED Test tabs to the `WIP Master` Tahoma-8 look. `Test - MFD` is a DATA-ENTRY tab
that replaces `WIP - MFD`, so it mirrors `WIP - MFD`'s Calibri look instead, including
MFD's existing bold-orange-on-grey input convention. **Do not restyle it to Tahoma 8** —
that would make it stop looking like the tab it is replacing.

**MFD192 anchoring (the owner's ruling, 2026-08-25).** `WIP - MFD` carries three contract
rows for job 192 (Hudsonwood 009 / Offsite 010 / base 008); QBO has ONE project MFD192.
Costs cannot be split — 455 of 460 cost lines carry no contract marker at all, and the five
that do account for well under 1% of the job. So QBO figures land on the
**largest-contract row** of each job
group (row 10, the 008 base, for MFD192) and sibling rows get a muted `see MFD192` marker.
`SUM()` ignores text, so the totals row still counts each job exactly once.

**Idempotency.** `--seed` builds the tab (gated behind `CONFIRM=Y` if it already exists,
since re-seeding discards whatever MFD typed). The default run refreshes **only** Q, R and
the banner; `build_columns` runs every pass but never overwrites an ETC or REVISED ETC that
already has a value, so a row MFD adds later picks up the styling and formulas on the next
sync. Verified: typed values in N and O survive a refresh, stale QBO cells get replaced.

**Cells are visually locked, not protected** (owner's choice). No sheet protection is
applied — the QBO look is carried by the green header, the tinted fill and a cell comment.

**Cross-check the tab now gives for free** (first live run, 2026-08-25): QBO billed vs the
tab's own hand-entered `COMPLETED TO DATE` matched to the penny on MFD192 and MFD325, and to
within a dollar on MFD177. MFD295 differed materially — see OPEN ISSUES.

## OPEN ISSUES

- **MFD295 (ELITE ROCK CREEK) billed disagrees materially with QBO.** The tab's
  `COMPLETED TO DATE` runs well above the QBO project P&L, and the job shows 100% complete
  on the tab. Someone has to reconcile which is right — the script only surfaces the gap, it
  does not resolve it. Figures are in the vault log, not here (repo rule 7).
- **MFD295 is on `WIP - MFD` but not on `WIP Master`.** The two tabs do not carry the same
  job list. Not a defect of this script, but it means the tabs will not tie out.
- **Per-contract billed is available but not implemented.** MFD192's invoice memos DO name
  the contract ("HUDSONWOOD CONTRACT", "OFFSITE CONTRACT", bare = 008), so BILLED could be
  split three ways even though COSTS cannot. Left out of v1 deliberately: splitting billed
  while costs stay anchored would put a billed figure next to an empty cost cell and make
  the margin on those rows read as real. Revisit only if the owner wants it.

## 2026-08-25 (later) — `RETAINAGE (QBO)` added to the green block

The owner asked to see QBO's retainage next to MFD's own, so the green QBO block grew from
two columns to three: `COSTS TO DATE` · `BILLED TO DATE` · **`RETAINAGE (QBO)`**, with
`COST TO COMPLETE` moving from S to T and the sync banner widening from `Q5:R5` to `Q5:S5`.

**Where the number comes from — do not re-derive.** QBO tracks retainage properly: the
invoice item `99 - Retainage` posts to a real Other Current Asset account,
**`Retainage Receivable`**. A negative retainage line on a draw DEBITS that account
(retainage moves out of AR); billing the retainage later CREDITS it back out. So the
per-job balance of that account IS "what QBO has", pulled from the `GeneralLedger` report
filtered to that account. Verified: the per-name sum ties to the account's `CurrentBalance`
exactly, which is how the pull is known to be complete.

**Two traps, both now handled in code:**

- **The GL report's account filter is `account`, SINGULAR.** Passing `accounts` is accepted
  and then silently ignored — you get the entire 66k-row general ledger back, truncated to
  its first 11 accounts, with no error and no retainage section. It looks like a clean empty
  result. `account=<id>` returns 101 rows and one section.
- **Do NOT reuse `cp_wip_reader`'s retainage heuristic** (gross P&L income minus the sum of
  non-retainage invoice totals). It is built for CP and is badly wrong on MFD — on the
  largest job it missed by more than twice the retainage actually at stake — because
  retainage that has since been BILLED still sits in the invoice history. The GL balance
  nets it out; the invoice scan does not.

**The column is expected to disagree with col M, and that is the point.** QBO stops counting
retainage once it has been billed to the GC; the WIP tab keeps carrying it. The cell comment
states the variance per job, summed across the job's contract rows (MFD192 spreads retainage
over three rows against one QBO balance).

First live run: **MFD192 agrees to the penny**. The other three do not — see OPEN ISSUES.

**Migration guard.** A tab built by the earlier two-column version carries a stale `Q5:R5`
merge and a `COST TO COMPLETE` formula sitting in S. `build_columns` now drops any
banner-row merge that is not the current span before re-merging (overlapping merges are an
Excel repair prompt), and clears any formula found in a QBO column. Upgrade path tested on a
copy of the already-built tab: `assert_clean` passes, B..M still byte-identical.

## OPEN ISSUES

- **Retainage: three of four MFD jobs disagree with QBO, all in the same direction** — QBO
  lower, i.e. QBO has already released retainage the WIP still carries. MFD192 agrees
  exactly; MFD177, MFD295 and MFD325 do not. On MFD177 the QBO balance is NEGATIVE, meaning
  more retainage was invoiced than ever accrued to that account. Each gap lines up with an
  amount the hidden `RETAINAGE MFD` tab records as last billed, so these read as retainage
  that was invoiced and never taken off the WIP — but that is a call for a person, not the
  script. Per-job figures are in the vault log; repo rule 7 keeps dollar exposures out of
  this file.

## 2026-08-25 (later still) — replacement audit: the seed was missing sheet CHROME

Before retiring `WIP - MFD`, both sheets were diffed attribute by attribute. **The data copy
was complete** — cell values, formulas, every style facet, merges, column widths, row
heights, hidden state, comments, conditional formatting, data validation, hyperlinks,
images/charts, autofilter and freeze panes were all identical or absent in both.

**What the seed had NOT copied: sheet-level chrome, all of it print behaviour.** Tab colour,
page orientation (landscape), `fitToHeight=0`, `fitToPage`, page margins, zoom, and the
**print area**. openpyxl does not carry any of these with a cell-by-cell copy. Nothing here
shows up on screen, which is why it survived the first build unnoticed — it only appears
when the report is printed or PDF'd, and MFD's WIP goes to banks.

Fixed by `copy_sheet_chrome(src, ws, force=False)`, called from `seed()` and from
`build_columns()` so a tab built by an earlier version **self-heals on the next sync**
rather than needing a re-seed. With `force=False` it fills in only what is UNSET on the
target, so it never fights a later hand adjustment.

**Trap inside the fix:** test `pageSetUpPr.fitToPage`, NOT `pageSetUpPr is None`. openpyxl
returns an empty `PageSetupProperties` object rather than `None`, so a container check
silently skips it — leaving orientation and `fitToHeight` set but not actually honoured,
which prints across pages while every attribute reads correct in code.

**Print area was deliberately WIDENED, not copied verbatim** — `$B$2:$L$15` → `$B$2:$T$15`.
The source's area stops at column L: it predates both the Total Retainage column (M) and
everything this script adds, so a verbatim copy would print a report missing the new
columns. Same first cell and same last row as theirs; only the column span changed. To go
back to the original span, set `ws.print_area` in `copy_sheet_chrome`.

**Nothing else in the workbook references `'WIP - MFD'`** — checked every sheet XML part.
The only reference is its own print-area defined name, which retires with the sheet. So
deleting it breaks no formula anywhere.

**Verdict: `Test - MFD` is a faithful superset of `WIP - MFD` and safe to swap in.**

## 2026-08-25 (final) — MERGED: `Test - MFD` retired, `WIP - MFD` graduated

The owner's call after the replacement audit came back clean: **merge the two tabs into one
under the original name.** `WIP - MFD` now carries the entry columns and the QBO block;
`Test - MFD` is deleted.

**`wip_excel_guard` graduated `'WIP - MFD'`** — the documented path (owner's explicit
instruction, name added deliberately to `ALLOWED_WRITE_SHEETS`, never a config flag). It is
the first live division tab any script may write.

**The contract that makes writing a live tab safe:** columns B..M are MFD's. The script reads
them only for the job number, contract, change orders and the retainage variance, and NEVER
writes them. It owns N..T and nothing else. Verified against the pre-merge backup: B..M values
AND formatting are byte-identical after the live write.

**Script simplified with the staging tab gone.** `seed()` and `--seed` deleted (source and
target are the same sheet now), `copy_sheet_chrome()` deleted (writing in place, the chrome is
already correct and must not be touched) — replaced by `widen_print_area()`, which only pushes
the print area out to cover the owned columns and leaves it alone once wide enough.
`main()` deletes a leftover `Test - MFD` on sight so the two can never drift apart again.

**Key/legend block removed** (owner: clutter). `_clear_legend()` wipes it wherever an earlier
run parked it, so an already-built tab is cleaned on the next sync rather than needing a
rebuild.

Backups before each step are in `WIP History/`: `… (pre Test-MFD 08-25).xlsx` is the tab as it
stood before any of this work, `… (pre MFD merge 08-25).xlsx` is immediately before the merge.

## OPEN ISSUES

- ~~`Test - MFD` stays in `ALLOWED_WRITE_SHEETS`~~ — **stripped 2026-08-25.** The allow-list
  is now `Test`, `Test-Master`, `Test - CP`, `Test - RP`, `WIP - MFD`. `mfd_wip_test` still
  deletes a stray `Test - MFD` if a workbook carries one, but a sheet DELETE does not go
  through `assert_write_allowed`, so it needs no entry.
- The other live division tabs (`WIP - CP`, `WIP Master`) remain code-locked. This graduation
  covers `WIP - MFD` only, and deliberately so.
