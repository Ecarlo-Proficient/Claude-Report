# statement-reconciler — STATUS

Shared progression record (repo rule 7). Update in the SAME commit as any change
to this tool. Tool-only scope: no business/owner analyses or dollar-exposure
findings here — those live in the owner's vault.

## DONE / FINALIZED
- **New folder layout + filing + DONE workflow (2026-09-16, owner).** The whole
  Vendor Statements folder moved into the Accounting share's new `Accounts Payable/`
  parent: `INBOX_ROOT = /Volumes/Accounting/Accounts Payable/Vendor Statements`.
  - **Filing:** the Inbox is a pure dump. `process_pdf` now files BOTH the Excel
    and the source statement together under `<root>/<Vendor>/<MM-YYYY>/` (vendor =
    resolved QBO name via `_vendor_folder`; month = `MM-YYYY`, month-first per the
    owner's no-year-first rule). The old flat `Reconciliations/` + inbox `DONE/`
    are retired; `run_inbox` no longer moves to DONE (filing is centralized in
    `process_pdf`), and `_resolve_workflow_dirs` now returns `(inbox, root)`.
  - **Finalize = the clerk marks the MONTH FOLDER done** (owner 2026-09-16, "if a
    excel/statement was put in DONE folder mark the month folder done"). When a
    vendor's `<MM-YYYY>` folder is renamed to `<MM-YYYY> DONE`, the whole month is
    left untouched on re-runs (`_month_done_dir`, checked right after vendor
    resolve, before the QBO bill pull - a re-drop of a finished statement costs
    nothing and stays in the Inbox for a human to notice). Un-done months
    regenerate in place. Tie-out failures leave the source in the Inbox (held).
  - **Migration:** one-offs (scratchpad, not committed) routed all 125 existing
    files into the new tree, then marked the 35 fully-reconciled month folders
    `<MM-YYYY> DONE` (from the old Old-Done / DONE FOR REAL). Verified live: an
    inbox sweep filed a fresh Cowtown 09-01, and re-dropping a Cowtown 07 or a
    White Cap 06 statement was skipped ("month already marked DONE").
  - **Bill Tracker moved too (same day):** out of OneDrive `Automations-/` into
    `<Accounting>/Accounts Payable/Bill Tracker.xlsx`, via the new ONE resolver
    `shared/paths.bill_tracker_xlsx()` that every reader/writer shares (bill-tracker,
    ledger, invoice-sync, health-dashboard). **Dockerized invoice-sync must set
    `ACB_BILL_TRACKER_XLSX` to its in-container mount** or the AR-aging tab won't
    find it.
- **Print-status future-proofing + Summary restructure (2026-09-14).**
  - **QBO↔printed agreement gate (the tie-out analog).** Each statement invoice
    carries two independent facts: is it in QBO (the reconcile matched a bill) and
    did the reader find a printed email? A "NOT PRINTED" that is ALSO in QBO is a
    *reader blind spot* (a format we don't handle), not a genuine gap - it raises a
    red banner in the Summary, exactly like TIE-OUT FAILED. `build_print_rows`
    takes `qbo_refs`; `PrintRow.reader_suspect` flags the disagreement. This would
    have caught the Croell bundled-PDF issue on its own.
  - **Index-health sanity.** The disk cache stores the printed-email count; a fresh
    pull that is empty or drops >40% raises a health warning (catches a truncated
    mailbox pull or a printer-category label that stopped matching "printed").
  - **Backup telemetry, error vs not-found.** `live_search_printed` now returns a
    `SEARCH_ERROR` sentinel on a Graph error (never silently "not printed"); those
    rows show "unverified". `search_stats()` tracks calls/hits/notfound/errors.
  - **Self-audit mode** `--audit-print-status`: sweeps every statement (inbox +
    DONE) and reports reader coverage per statement + backup telemetry + index
    health. Run after ANY matcher change (the print-status analog of the parser's
    "sweep ALL PDFs" rule). 2026-09-14 baseline: 2405/2465 matched, 0 errors.
  - **Summary restructure (the owner 2026-09-14).** Print status moved OUT of a
    separate tab and INTO the Summary as the FIRST section, split into two grouped
    sub-lists: "Printed bills check" (collapsed) and "Unprinted bills" (open by
    default, with an "In QBO?" column and a yellow "Printed ✓" box the bill clerk
    marks - the workbook doubles as her worklist; the invoice clerk works the QBO
    buckets below). Empty QBO buckets (0 bills) are no longer rendered at all.
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
  - **Multi-bill / PDF-only backup (`live_search_printed`, 2026-09-14).** For an
    invoice the pre-built index misses, a fallback runs a Graph `$search` for the
    invoice #'s longest alphanumeric SEGMENT before flagging it. `$search` reads
    INSIDE PDF attachments server-side, so a bill that arrived bundled with others
    in ONE printed email whose PDF has a generic filename (Croell "Croell Invocies"
    -> "Proficient Concrete Invoices 1.pdf") is still found - no PDF download. The
    segment (not digits-only) term is required because Graph tokenizes "RW786083"
    and "188673772-0001" as whole words. Proven: Croell 09-08 9/14 -> 12/14 (3
    recovered from the bundle); a "NOT PRINTED" that survives the backup means the
    # is in no printed email at all. Per-# cached; only misses trigger it (a few
    calls per statement).
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
- **Ellis-type vendors (invoice # only inside the PDF)** are now largely handled
  by the `live_search_printed` backup (Graph reads inside the PDF), PROVIDED the
  invoice email is printed-tagged and the PDF has a real text layer. A pure
  scanned-image PDF (no text) is not full-text-indexed by Graph, so it could still
  show a false "NOT PRINTED"; the whole-statement-0% "unverified" caveat remains
  the safety net for that case.
- **First-run sweep (2026-09-14) behaves as designed:** after the backup, the
  surviving "NOT PRINTED" rows are genuine (no printed email at all) and split into
  old invoices on old statements (pre-tagging / pre-mailbox-horizon) and a small
  number on current statements that are real AP-01 intake gaps, not tool bugs. The
  specific bills go to the owner, never into this repo (STATUS scope rule).

## VERIFICATION NOTES
- The 12 current reports (statements 08-26 → 09-08) were audited 2026-09-10/11:
  each report's Statement total equals the exact grand total printed on its
  source PDF; 11 of 12 tie their parsed lines to that total. Cowtown 09-01 was the
  lone failure and was a STALE report (old build), not a current-code parse bug —
  regenerate it. No misclassifications; no fabricated lines. The Preferred
  running-balance parse (balance-delta amounts, not gross charges) is correct.
