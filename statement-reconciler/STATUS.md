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

## OPEN ISSUES
- **Cowtown Redi Mix parser is broken (found 2026-09-11).** `parse_statement_cowtown`
  expects one line per invoice (`<6-digit ref> <date> <date> $inv $amt`), but the
  current Cowtown statements are a multi-page, per-job aging layout. Most lines
  are dropped, so the parsed line-sum falls far short of the printed Amount Due
  and open QBO bills false-flag as MISSING_ON_STATEMENT. Reproduces on the
  09-01-2026 statement (tie-out fails) and the earlier 07-02 statement. The new
  tie-out gate now catches this instead of shipping it silently, but the parser
  itself still needs a rewrite for the aging layout. Fix requires the source PDF
  on the Synology `Vendor Statements` share (mount before working).

## VERIFICATION NOTES
- The 12 current reports (statements 08-26 → 09-08) were audited 2026-09-10/11:
  each report's Statement total equals the exact grand total printed on its
  source PDF; 11 of 12 tie their parsed lines to that total; Cowtown 09-01 is the
  lone failure. No misclassifications; no fabricated lines. The Preferred
  running-balance parse (balance-delta amounts, not gross charges) is correct.
