# statement-reconciler — STATUS

Shared progression record (repo rule 7). Update in the SAME commit as any change
to this tool. Tool-only scope: no business/owner analyses or dollar-exposure
findings here — those live in the owner's vault.

## DONE / FINALIZED
- **Print Status tab — print-status verification (2026-09-14, opt-in).**
  `print_status.py` ties AP-03 → AP-01: for every parsed statement invoice it asks
  the billings mailbox "was this bill ever received-and-printed?" (an untagged
  email is an invisible bill). Findings, all verified by a full inbox+DONE sweep:
  - **Source per vendor, not RCI-only.** The invoice # is pulled from the email
    subject, the FULL body (HTML stripped), AND attachment filenames - because
    each vendor encodes it differently. Verified: White Cap **0 → 100%** (its #
    sits in the body table, past the 255-char bodyPreview - full body was the fix),
    CMC 100% + Sunbelt 88–100% (filename decode), Cowtown 94% (QBO-desktop body).
    No vendor regressed. Moderate rates (Sunrise/Croell) are the OLDEST statements
    only (pre-tagging invoices); every recent statement hits ~100%.
  - **Matching:** three keys per token (full normalized, longest digit run,
    all-digits-0-stripped), most-precise-first, so a ref matches however it's
    written. `_is_printed_category` = whole-word "printed" (catches ITZ/GISELLE
    PRINTED, excludes PRINT PENDING); no personal name hard-coded.
  - **Auth:** MS Graph client-credentials, keys in the ONE automation-qbo blob
    (GRAPH_TENANT_ID/CLIENT_ID/CLIENT_SECRET/BILLING_MAILBOX, all `required=False`
    in `shared/setup_qbo.py` so the QBO auth test never demands them).
  - **Perf:** first build ~514s (full body, all messages); then a disk cache
    (`~/Library/Application Support/Proficient/print-status/`, override
    `PRINT_STATUS_CACHE_DIR`) makes each run an **~11s incremental** pull, floored
    on `lastModifiedDateTime` so a newly-printed OLD bill is still caught and an
    un-tagged email is dropped.
  - **Wiring:** `write_excel` gained `stmt_lines=` and writes a plain "Print Status"
    tab (Date · Invoice # · Amount · Printed? · Email date · Email subject),
    passing `assert_clean`. **OFF by default; set `PRINT_STATUS=1` to enable** while
    we test on a vendor or two. Any Graph failure skips the tab, never breaks the
    reconcile. A whole-statement 0% match (e.g. Ellis, invoice # only inside the
    PDF) shows "unverified" + a caveat instead of a wall of false "NOT PRINTED".
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
- **Print Status is opt-in (`PRINT_STATUS=1`) pending a test-and-see on 1–2
  vendors** (the owner 2026-09-14). Flip the default on once proven.
- **"Who's done" workflow script HELD (the owner 2026-09-14, "hold - don't build
  yet").** Design agreed: a script that DERIVES each statement's state - the bill
  clerk's fix-the-bad-bills side from the reconciliation xlsx (tie-out held /
  MISSING_IN_QBO) + folder location (`Reconciliations/waiting for actions` →
  `Old-Done`), and the print side from the Print Status scan - with NOBODY
  maintaining a tracker. Not started.
- **Least-privilege lockdown not yet applied.** The Graph app can currently read
  ALL mailboxes; restrict it to the billings mailbox with an Exchange
  `New-ApplicationAccessPolicy` (owner/IT step, provided, not run).
- **Ellis-type vendors (invoice # only inside the PDF)** are shown "unverified",
  not confirmed. Reading the PDF bytes to close them is out of scope for now
  (the owner: statement reconcile should avoid opening each PDF).

## VERIFICATION NOTES
- The 12 current reports (statements 08-26 → 09-08) were audited 2026-09-10/11:
  each report's Statement total equals the exact grand total printed on its
  source PDF; 11 of 12 tie their parsed lines to that total. Cowtown 09-01 was the
  lone failure and was a STALE report (old build), not a current-code parse bug —
  regenerate it. No misclassifications; no fabricated lines. The Preferred
  running-balance parse (balance-delta amounts, not gross charges) is correct.
