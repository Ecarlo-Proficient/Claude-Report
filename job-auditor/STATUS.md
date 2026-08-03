# job-auditor/ — STATUS (shared progression record)

> Rule: update this file in the SAME commit as any change to this tool
> (CLAUDE.md/AGENTS.md structure rule 7). Tool matters only — no business
> findings, no dollar exposures, no owner-only analysis.

Last updated: 2026-07-28

## DONE / FINALIZED

- **`vault-steward/` scrapped and replaced by this tool.** v1 watched the code
  repo and could only report file hygiene. The tree that matters is
  `/Volumes/Common/CURRENT PROJECTS/` and the errors that matter are *inside*
  the workbooks and PDFs. Only runtime plumbing survived. `SPEC.md` §0.
- **`SPEC.md` v3, calibrated on two live runs** over the active board (48 slab
  lines, 45 auditable). Hit rates in §2; exposures live in the vault.
- **The bid proposal is a parseable line-item table** (`DESCRIPTION · UNIT ·
  PRICE/UNIT · TOTAL`). Working prototype parser reconciles Σ(sold lines) to the
  proposal's own SUB TOTAL to the cent on a hand-checked fixture.
- **Ruling REVERSED from v2: piers ARE sold as their own priced line.** v2
  concluded "described but never priced" — an artifact of descriptions wrapping
  across lines. Consequently the valuable rule is **scope-level margin** (sold
  line vs cost band), not presence/absence.
- **Ruling: the `#N/A` band drop is a separate, silent ETC defect.** Both live
  ETC paths already exclude FW from a slab ETC, so FW on slab takeoffs is data
  noise. A band whose subtotal errors is dropped whole with no notice.
- **Detector discipline is binding** (`SPEC.md` §5), earned from FIVE wrong
  detectors in one afternoon: blind to `#N/A`; regex missing the plural
  "PIERS"; line-regex defeated by wrapped descriptions; a capture-group
  off-by-one that shifted every unit and total one position; and stopping at
  the derived cost sheet instead of the source sheet. None was visible in the
  output. Required now: **reconciliation gate**, **fixture first**,
  **signal-quality line**, **trace to the source sheet**.

- **VERIFIED (2026-07-28): our ETC reader mis-reads the cost-sheet variant.**
  The `Cost Gral` sheet has **two scopes, not three bands** — FOUNDATION
  (slab + piers, its SUB TOTAL summing from the slab subtotal down) and
  FLATWORK (its own scope). No piers ⇒ pier rows read 0 and you read the slab
  subtotal. Our `ETC = SL_sub + PR_sub` double-counts the slab on that variant.
  **The workbooks are correct; the bug is ours. Never edit the takeoffs.**
- **Both variants are live** — sweep of the active board: 21 CUMULATIVE
  (reader double-counts), 14 BAND_ONLY (reader correct by accident), 21
  unclassified, 10 skipped. **Wrong on 19 of 35 classifiable jobs, in both
  directions** — overstating where the subtotal reads, understating where it is
  `#N/A` (band dropped, piers lost). Classification is rule 1.
- **An earlier verdict ("the formula is one row too tall — intentional
  rejected") was WRONG and is retracted.** It was defended with four checks
  that were all equally consistent with the opposite conclusion. **A test that
  cannot fail is not evidence** — state the refuting result before treating a
  cross-check as confirmation, and ask whoever built the artifact before
  concluding it is broken.
- **The estimator's red-flag rule is endorsed — and tracing it inverted the
  conclusion.** Pier cost on a job with no piers is **real cost, misfiled**:
  the `Piers takeoff` sheet's subtotal ranges sweep in site/general lines (tie
  wire, scrape lot, reset forms) that map to `PR3`/`PR6`. The slab sheet carries
  no scrape-lot or reset-forms line, so excluding the band LOSES real money.
  **Ruling: count it, flag the miscoding** (`SPEC.md` rule 5). Carve-out: tie
  wire appears on both sheets, sometimes at identical amounts — report those as
  possible duplicates rather than summing them.
- **Three findings inverted on tracing one level deeper today** — takeoff bug →
  reader bug; phantom cost → misfiled cost. Added as detector lesson #5: stop at
  the derived sheet and you will name the wrong culprit.

## IN PROGRESS

- Nothing. Three read-only prototypes, one per live rule:
  `variant_sweep_prototype.py` (rule 1 — cost-sheet variant classification),
  `proposal_parser_prototype.py` (line-item parser + reconciliation gate), and
  `piers_sheet_trace_prototype.py` (rule 5 — traces the source Piers sheet).
  None is a shipped tool — stdout only, no watermarks, no xlsx, no launchd.

## TO DO

- **Phase 1 — rule 1 (cost-sheet variant classification) alone.** Mechanical,
  needs no PDF parsing and no proposal: read the piers subtotal's formula range,
  classify CUMULATIVE vs BAND_ONLY, compute the ETC accordingly. Prerequisite
  for every value-level rule. Ship with one workbook of each variant as fixtures.
- Phase 1b — reconciling parser + rules 2–4 (scope-vs-cost margin, `#N/A` band
  drop, missing slab), with the fixture test committed alongside.
- Phase 2 — close the parser gap: 13 proposals fail the reconciliation gate;
  7 are a tract price list with no SUB TOTAL (exempt by design — detect and
  skip), 6 need the grammar extended.
- Phase 3 — rules 5–6 (cost band with no sold line → reattribute, and the
  inverse), including the tie-wire duplicate carve-out.
- Phase 4 — change detection (`SPEC.md` §7), then launchd weekly, `.disabled`.
- Fold `one-offs/rp_wip_simple.py`'s `proposal_has_piers` into the rule catalog
  instead of maintaining two implementations — it has detector bug #2 (its
  `"PIER" in txt` test is substring-based, so it escapes the plural
  bug, but it is still presence-only and superseded by the parser).
- Update `docs/ARCHITECTURE.md` when Phase 1 ships (no tool shipped yet, so the
  diagram is intentionally untouched by this commit).

## OPEN ISSUES

- **Three active takeoffs have no `Cost Gral` sheet** — unauditable. Different
  template or genuinely incomplete? Needs a human.
- **Plans are unreadable to the tool.** Pier counts/depths live in drawings,
  often as images. Scope existing only on the plans is invisible — a permanent
  bound on rules 4–5.
- **21 of 56 takeoffs still unclassified** — no readable piers-subtotal
  formula, so the variant is unknown and the ETC for them cannot be trusted
  either way. (The earlier "40 unclassifiable" figure came from the superseded
  value-based test; the formula test resolved 35.)
- **Cost bands do not map 1:1 to sold lines.** A rate line can absorb cost
  sitting in another band, so single-line margin is a flag to verify, never a
  verdict. Rule 2 must present it that way.
- **Template drift (rule 9) is unwritten** — the failure mode that would
  silently disable every other rule.
- Prototypes depend on `rp_schedule_wip_preview` internals
  (`_cost_sheet_totals`, `find_proposal`, `find_takeoff_etc`). Per repo rule 3,
  graduating this tool means moving those to `shared/`, not cross-importing
  from `one-offs/`.
