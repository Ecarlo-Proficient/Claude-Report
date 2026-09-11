# statement-reconciler — STATUS

Shared progression record (repo rule 7). Update in the SAME commit as any change
to this tool. Tool-only scope: no business/owner analyses or dollar-exposure
findings here — those live in the owner's vault.

## DONE / FINALIZED
- **Parse tie-out gate (2026-09-11).** Every report now shows a `Sum of parsed
  lines` row and a `Parse gap (lines − statement total)` row in the TIE-OUT
  block. When the parsed lines do not sum to the statement's own Amount Due
  (tolerance $0.50), the Summary gets a red **TIE-OUT FAILED** banner and the
  reconciliation is treated as unreliable:
  - `process_pdf` returns `(ok, counts, tieout_ok)`.
  - `run_inbox` keeps a failed-tie-out report **out of DONE** (new `held`
    bucket in the INBOX SUMMARY) so a human re-parses the source; the sweep
    exit code is nonzero when anything is held.
  - Unattended (`--inbox --yes`) runs no longer silently emit-and-archive a
    broken parse — the hole that shipped the Cowtown 09-01 report.
- **`assert_clean` wired in (2026-09-11).** `write_excel` now runs
  `shared/xlsx_verify.assert_clean` as its last step (repo rule 5b). Was
  previously missing from this tool.
- **Balance-forward / payment reconciliation (2026-09-11).** Running-balance
  statements (e.g. CowTown 09-01) carry a `Balance forward` lump (parsed with
  `ref == ""`) and `PMT` payment rows (negative amounts). These are NOT bills to
  match against QBO — `reconcile_iter` now excludes them from per-invoice
  matching (they no longer false-flag as MISSING_IN_QBO) while they stay in the
  tie-out. QBO bills dated on/before the balance-forward date are folded into the
  forward, so they're excluded from MISSING_ON_STATEMENT (with a count/$ warning).
  The Summary tie-out decomposes the total into itemized invoices + balance
  forward + payments. Verified on CowTown 09-01 (was 20 false MISSING_IN_QBO +
  76 false MISSING_ON_STATEMENT → 1 + 4 genuinely actionable) and RCI 08-26
  (payments pulled out of MISSING_IN_QBO 20→15; still ties). Vendors without a
  balance-forward or payment rows are unaffected (verified: plain-vendor reconcile
  unchanged).
  - **Root cause of the original CowTown "under-parse"**: it was NOT a parser
    bug. `detect_template` already routes CowTown to `qbo_statement` (its
    statements carry `INV #N. Due N` lines), and that parser already ties out
    (parsed lines net to the printed Amount Due, including the balance-forward
    lump). The broken 09-01 report on the share was produced by an OLDER build;
    regenerating it with current code is the fix. The legacy tabular
    `parse_statement_cowtown` / `COWTOWN_SIG` are left in place as a fallback for
    the older single-line CowTown layout.

## OPEN ISSUES
- (none)

## VERIFICATION NOTES
- The 12 current reports (statements 08-26 → 09-08) were audited 2026-09-10/11:
  each report's Statement total equals the exact grand total printed on its
  source PDF; 11 of 12 tie their parsed lines to that total. Cowtown 09-01 was the
  lone failure and was a STALE report (old build), not a current-code parse bug —
  regenerate it. No misclassifications; no fabricated lines. The Preferred
  running-balance parse (balance-delta amounts, not gross charges) is correct.
