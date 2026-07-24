# wip/ — STATUS (shared progression record)

> Rule: update this file in the SAME commit as any change to this tool
> (CLAUDE.md/AGENTS.md structure rule 7). Tool matters only — no business
> findings, no dollar exposures, no owner-only analysis.

Last updated: 2026-07-21 (pm)

## DONE / FINALIZED

- **Test-Master is the deliverable WIP report** (2026-07-16): "WIP REPORT as of
  <date>" banner, rows 1–2 reserved as logo space (embedded images survive every
  sync), TOTALS row on live `SUBTOTAL(109,…)` (re-totals with the table filter),
  FUTURE WIP CASH FLOW block derived from the TOTALS row.
- **Identifiers are grab-able** (2026-07-16): PROJECT #/NAME cells are plain
  text; folder + data-source links moved to their own PROJECT FOLDER /
  DATA SOURCE columns. OVERBILLINGS/UNDERBILLINGS headers shortened to fit.
  FLAGS states the classify reason on red rows (never "OK" on red).
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
