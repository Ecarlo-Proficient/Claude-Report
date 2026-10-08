# statement-reconciler — STATUS

Shared progression record (repo rule 7). Update in the SAME commit as any change
to this tool. Tool-only scope: no business/owner analyses or dollar-exposure
findings here — those live in the owner's vault.

## DONE / FINALIZED
- 2026-10-08 (night) · **New approval process in the reconciler (owner: "bills are approved in QBO now, but there are
  still some that use NOT APPROVED").** The same three states as the Bill Tracker, from `shared/bill_approval`: NOT
  APPROVED memo (paper era - chase the PM) · **check QBO** (entered since 09/16 and unpaid - its approval is in QBO's
  queue, which the API can't see) · approved. Excel: a "MATCHED - approval is in QBO" section, the Approved? column
  says Check QBO (teal), the banner counts both kinds. Marked statement: a teal "approval is in QBO" colour. Notion /
  Teams: the same states (the board's check-QBO item already existed; now from the shared rule).
  `tests/test_bill_approval.py`.
- 2026-10-08 (evening) · **Excel layout per the owner's spot-check + the last three vendors refreshed.**
  - Marked sheet: the colour key moved to the RIGHT of the statement (column P, from the first page's row - the
    pages fill A:N); payment / credit / balance forward is purple, "another customer's line" is grey (the two
    greys were confused); prints landscape, one page wide.
  - No Statements or Changes sheets - the Notion page shows both. The workbook is Summary + Statement (marked).
  - Columns: the tool's "Note" is now **Finding**, the clerk's "Clerk notes" is now **Notes** (her notes are still
    read from any column M header that says NOTES).
  - Summary: no blank row 3 under the frozen title (it showed as a white band) - the tie-out starts on row 3,
    frozen at A3. Em dashes removed from everything the tool prints or writes (owner rule).
  - Void Forms, Power Jack, Post-Tension refreshed live - QBO vendors confirmed by their bill numbers (VOID FORM
    PRODUCTS, POWER JACK FOUNDATION REPAIR; Post-Tension already cached); Power Jack pinned by its statement text.
    Every vendor Excel rebuilt in the new layout; Cowtown skipped while its Excel was open (re-run after closing).
- 2026-10-08 (afternoon) · **Marked-up statement in every Excel + the whole Inbox run live (owner: "highlight each bill row with the color of its bucket so the clerk can verify you grabbed everything").**
  - `statement_markup.py` (new): every page of each current statement rendered; every line the parser read is found
    on its page and its row banded in its bucket colour (matched / not approved / amount / tax / paid-vendor-shows-open
    / not in QBO / payment-credit / another customer's). The check runs both ways: a line read but not found on the
    page, and a row on the page that looks like a bill (date + amount) with no band - outlined red. The statement
    file is only read. Text PDFs use their own word boxes; scanned pages / images use OCR.
  - Excel: a 2nd sheet **Statement (marked)** (legend, "N of N lines found on the page", the misses, the pages; prints
    one page wide - a named print area, Excel clips pictures outside the cells that hold values). Pages as 64-colour
    PNGs (RCI 6 pages: 3.5 MB -> 0.9 MB). Notion: misses become "unread" to-dos + a **Not read** column.
  - Trained on every statement on the share (97 PDFs) and the page images inside 69 old reconciliation Excels.
    Fixes learned: a bill printed twice (tear-off stub) banded twice; $0 rows are not bills (2+ amounts: the last is
    the running balance); exact refs placed first, OCR look-alikes (O/0, S/5, one lost/extra char, an amount off one
    digit) ranked globally so a sequential neighbour never takes a row; no single OCR setting reads every page - four
    readings tried, the one placing the most lines kept; transparent screenshots flattened on white, small ones
    enlarged. Result: PDFs 96/97 clean, old-Excel images 65/69 (the rest: no source PDF, one line OCR can't see, and
    CMC 08-03 - a real miss the check catches every time).
  - Parser fixes found by the markup: Cowtown "Discount -0.01" rows (DISC lines) - the 07-02 statement was off 0.13.
    New layout: VoidForm customer statement. Scanned statements (no text in the PDF) are read with OCR and still must
    tie out; the Amount Due box unreadable -> the last running balance (Power Jack 07-28 now reads).
  - Fixes: a History twin no longer knocks out the live statement (RCI refresh read 'no readable statement'); the
    clerk's own notes column ('<name>'S NOTES', column M, pre-10/07 workbooks) is picked up.
  - Live 10/08: `--inbox --yes` - 18 vendors, 17 Inbox statements, all lines found on their pages; Inbox empty; all 20
    vendor Excels opened clean in real Excel. New vendors confirmed by their QBO bill numbers (Quikrete, CONSOLIDATED
    REINFORCEMENT) and cached.
  - Not migrated yet (no Inbox statement this run; `--refresh` picks them up): Void Forms, Power Jack, Post-Tension.
    Not statements, left unread on purpose: Bobcat 06 payment receipt, Texas Materials invoice, Sunbelt / Core Supply
    printed lien-notice emails.
  - Tests: +6 (VoidForm, scanned fallback, markup: printed twice, look-alike ranking, $0 / total rows); 42 in all.
- 2026-10-08 · **Inbox sweep: every statement on the share read through the parser, misses fixed, safety nets added (owner: "make sure it catches all and has safety nets so we know it didn't read it right").**
  - New layouts: Abatix online statement (`vendor_abatix_online`), Croell "AR Customer Balance" (`vendor_croell_ar` -
    the 7-digit invoice # runs into the PO in the text), Quikrete Ready Mix account statement (`vendor_quikrete`),
    a vendor's QuickBooks "Invoices for <us>" list (`qbo_invoice_list`, CRI-DFW; only open lines count).
  - Fixes: a QBO Open Invoices report covering ALL the vendor's customers (Core Concrete Pumping) - only our section
    counts, a page header mid-section never splits it; QBO statement invoice # with a letter (`INV #R16738`, a revised
    invoice); the statement date when its "Date" label is far from the value (Carder read the first invoice date);
    vendor pinned by its own statement text for Gonzalez Brothers and Core Concrete Pumping (+ Core alias); the
    "paid in QBO" note now says the BILL's date (it never was the payment date).
  - Safety nets: a statement dated before its own newest invoice, or in the future, or with no date, is never used
    (`Doc.problem`, shown in the run); an undated Excel/image statement takes the date in its file name, never today;
    a vendor name search with several QBO matches refuses instead of taking the first (Gonzalez matched seven, the
    first a person); a vendor folder differing only by LLC / Inc / LP is the same folder; an activity statement's
    balance forward is itemized only by a statement of the SAME batch (within 7 days) - a month-old list is stale
    (Preferred 09/01 vs 10/05); old unreadable files from month folders go to History as "Not read" (printed emails,
    scanned pages), a NEW one stays in front of a person and the vendor shows Unreadable.
  - `--audit-parsing` (new): every statement on the share through the parser + safety nets, read-only, no QBO; lists
    the ones that can't be trusted (new ones first) and a count per layout. Run it after any parser change.
  - Known, left on purpose: CMC 08-03 (a watermark printed across one row) and scanned / unsupported vendor files
    (Void Forms, Texas Materials, Power Jack, Core Supply, Bobcat 06) - the tie-out flags them; they are not used.
  - Tests: `tests/test_statement_parsers.py` (12 - one sample of the text the parser sees per layout learned, plus the
    safety nets). Rule: a new or fixed layout adds its sample there.
- 2026-10-07 · **One running record per vendor - no more months (owner: "all entered AS OF the last date the statement shows").**
  The clerk dumps statements in any order; a vendor often sends two close together (Cowtown: a past-due letter
  10/01 + its monthly activity statement 10/02) or the same file twice (RCI 10/07 x2), and month folders put a
  statement dropped on the 30th in the wrong month.
  - `statement_set.py` (new): each statement covers a RANGE of invoice dates - a full open list = everything up to
    its as-of date; Cowtown's past-due letter = up to the month of its newest invoice; an activity statement
    (balance forward) = only the invoices after the forward. Walked oldest first, a statement wipes its range and adds
    its lines; the rest is the vendor's ONE open list, as of the newest date. Same file twice (checksum) = Duplicate;
    a statement that does not tie out on its own is kept out (Unreadable). When the vendor's own documents disagree
    (Cowtown 10/01 vs 10/02) the page says so instead of calling the parse broken.
  - Folders: `<Vendor>/Current` (statements in use + ONE Excel `Reconciliation - <Vendor> - as of mm-dd-yyyy.xlsx`),
    `<Vendor>/History` (`mm-dd-yyyy DONE|Replaced|Duplicate - <file>`), `<Vendor>/.reconciler.json` (last as-of,
    clean flag, changes, notes). Old `MM-YYYY` / `MM-YYYY DONE` folders are folded in on a vendor's first run.
  - Clerk notes (her complaint: a re-run wiped her notes): Excel column M "Clerk notes" (yellow). Every run reads her
    notes out of the Excel before rewriting it and keeps them by QBO bill id (else invoice #), so a note follows its
    bill when rows move. Her note also shows on the bill's line in Notion. An Excel left open (a `~$` lock file) skips
    that vendor - nothing written over her work. No pre-10/07 workbook had notes left to carry.
  - Notion: a count column per bucket (Not entered, Not approved, Amount mismatch, Tax charged, Paid vendor shows open,
    Not on statement, Not printed, Check QBO, Matched), "Checked against" (which statements vs which QBO pull), column
    descriptions (added on the first live run). Page: Pay-run check · To do (as of) · Changed since <prev> (new / gone /
    amount changed) · Statements (every one, newest first) · Cleared. Pre-10/07 pages (month headings) are read as one
    bucket so the clerk's ticks carry. The "Month done" tick and `--sync-done` are gone (nothing to close).
  - CLI: `--only cowtown,rci` limits any run to those vendors; `--inbox --dry-run` now reads QBO + Notion and prints
    the full plan (it used to stop before QBO). `--refresh` re-checks every vendor with a current statement.
  - Parser fixes found on the way: Cowtown past-due letter invoice numbers with a letter suffix (`399238R`); Cowtown
    activity statement payments with no check # (`PMT -3,210.00`) - the 10-02 statement now ties out alone; QBO open
    balance "Credit M..." truncation (RCI 08/31 did not tie out).
  - First live run 10/07 (`--inbox --only cowtown,rci --yes --no-teams`): Cowtown 2 Current / 4 DONE / 1 Replaced,
    RCI 1 Current / 2 DONE / 1 Duplicate / 3 Replaced; both Excels opened clean in real Excel; Notion pages verified.
  - Tests: `tests/test_statement_set.py` (merge kinds, duplicates, unreadable, changes, notes surviving a re-run in a real
    workbook), `tests/test_statement_board.py` rewritten for the vendor page (15).
- 2026-09-30 · **Notion board rebuilt as ONE page per VENDOR and run live for the first time (owner).**
  The owner's question for the board is the pay-run check: "are all the bills from the statement ENTERED" - the clerk
  confirms it before sending the pay run to the manager; approvals, amount/tax fixes and printing are follow-ups done
  after paying, so they never change the colour.
  - Row = vendor (key = its folder name on the share, so an unreadable statement still lands on the right vendor).
    Status (colour) + Standing text: **All entered as of mm/dd/yyyy** (the last check) · **N not entered (oldest MM-YYYY)**
    · **Unreadable** (tie-out failed or the parser could not read a statement in its folder) · **No statement** (none in
    60+ days). Columns: Not entered · Follow-ups · Last statement · Last checked · QBO vendor.
  - Page = a Pay-run check callout, then one heading per OPEN month (newest first; a "Month done" tick, To enter in QBO,
    then the follow-ups), Cleared and History. Each bill is listed once per vendor, under the newest statement that shows
    it (running-balance statements repeat unpaid bills every month).
  - A month closes itself when it ties out with nothing left (owner: yes, auto-done), when the clerk ticks "Month done", or
    when its folder was already renamed DONE; closing renames the folder `<MM-YYYY> DONE` and moves it to History. A
    rename that fails keeps the month open. `--sync-done` is retired (closing happens in the run).
  - `--refresh --dry-run` now reads QBO + Notion and prints one line per vendor + the folders it would file (writes no
    Excel, nothing to Notion, never posts to Teams). New `--no-teams` switch. Teams digest lists only vendors not ready
    to pay.
  - Notion schema reshaped to match (database was empty). Views: "Pay-run check" (blocked vendors first), "Not ready to
    pay", "Board" by colour.
  - First live run 09/30/2026 (`--refresh --yes --no-teams`): 17 vendor pages, 20 clean month folders filed DONE, no
    Notion or rename errors; verified in Notion (rows, colours, a page body with QBO links). Run log in
    ~/Library/Logs/Proficient/.
  - Refresh is MANUAL before each pay run for now; a weekday-morning schedule comes with the Synology migration.
  - Tests: `tests/test_statement_board.py` (21: merge, month closing, one-bill-once, unreadable months, 60-day stale,
    grouping, the digest, a fake-Notion round trip + a read-only preview that fails on any write).
- 2026-09-29 · **(Replaced 09/30 by the one-page-per-vendor board above.) Notion Vendor Statements board replaces the Teams card per vendor (owner: "one card that keeps a live record of that statement").**
  The Teams Workflows webhook is post-only - it can't read a ✅ reaction or edit a card - so the live record moved to Notion.
  - `notion_board.py`: ONE page per vendor-month in the "Vendor Statements" database (`ACB_STATEMENTS_DS_ID` in machine.env; the database must be shared with the Notion integration). Body = one collapsible heading per kind of work (To enter · Not approved · Amount mismatch · Tax charged · Not printed · Approval pending? check QBO), one to-do per bill: ref linked to QBO, date mm/dd/yyyy, $, job. Plus a Cleared section.
  - Every run (inbox, single file, `--refresh`) rewrites the page and merges it with the clerk's ticks: gone from QBO -> Cleared (dated); ticked but QBO still shows it -> unticked "still open in QBO"; unverifiable (approval pending in QBO's workflow, bills entered since 09/16) -> her tick is kept. Print items are carried forward when print status didn't run. Her own notes on the page are left alone.
  - Status computed: Open / In progress / Clean / Tie-out failed; Done is set by a human and never overwritten. `--sync-done` mirrors Done both ways with the month folder (asks before renaming); every run prints what it would do.
  - Teams: one digest card per run (`shared/teams_notify.post_statement_digest`), a line per vendor-month linked to its page + an "Open the board" button. The per-vendor cards are retired.
  - Built full, then simplified the same day (owner: "build the max, then review to cut"): dropped per-kind count columns, $ totals/difference (noise - the statement total includes already-paid lines), tie-out checkbox, summary callout, a "for reference" section (158 already-paid lines on one RCI statement), run history, and PDF/Excel upload to Notion (copies of Synology). The full build is commit 822be06 if any of it is wanted back.
  - Tests: `tests/test_statement_board.py` (merge rules, status, grouping, a fake-Notion create -> tick -> refresh round trip, the digest card).
- 2026-09-29 · **Formula-injection guard (security review).** Outside text (vendor / QBO / Notion / email) this tool writes goes through `shared/xlsx_guard` - stored as text, never a live formula (statement refs, email-subject refs; the clerk-performance CSV via `csv_cell`); every workbook save also passes the shared guard, and `xlsx_verify` fails a risky formula.
- 2026-09-29 · **QuickBooks login moved to the shared one (security review).** This tool's own copy of the refresh-token exchange is retired; its login is `shared/qbo_api` (`load_credentials` / `get_pass`), which asks Key Helper (`keyhelper/`) once the key library is adopted - the tool never holds the refresh token.
- **`datetime.utcnow()` retired (2026-09-24).** `print_status.py`'s two date floors now use `datetime.now(UTC)`: same output, no deprecation warning (Python is removing `utcnow`).

- **Runs on python-env (2026-09-24).** The reconciler (the `statement-reconcile` shell helper) now starts through `python-env/python.sh` (`"$ACB_PY"`, Python 3.14, every package pinned), never a bare `python3` - a `brew install ffmpeg` on 09/23 swapped `python3` and broke every sync step. Missing-package hints now say `bash python-env/setup.sh`. Verified by a dry run on 3.14. See `python-env/STATUS.md`.

- **QBO Customer Open Balance date fallback + batch robustness (2026-09-17).**
  The newer Customer Open Balance layout prints the as-of date as a bare MM/DD/YY
  under the title (no "As of" wording) - Estrada 09-03 and Post-Tension 08-18
  parsed lines + tied out but returned no date and were skipped. Added a fallback
  that reads the MM/DD/YY under "Customer Open Balance"; both now date + tie out.
  Also: `find_vendor_id(strict=)` so an unresolvable vendor SKIPS instead of
  sys.exit-aborting a whole `--refresh` (a SystemExit slipped past run_inbox).
  Vendor aliases cached for Abatix (30145), Carder Concrete (30733), Gonzalez
  Brothers Batch Plant LP (28894, incl. the 'GONZALES' misspelling) - local cache.
- **Abatix Corp parser added (2026-09-16).** New template `vendor_abatix`
  (`ABATIX_SIG` = the doubled "Invoice Amount" header; parser reads Invoice Date |
  Due Date | Invoice No | PO | Amount Due | Enclosed No). All 3 filed Abatix
  statements tie out exactly (452.49 / 1,691.42 / 2,069.75); no other vendor
  mis-routes to it. `--refresh` now includes a no-recon month when its PDF matches
  a supported template (so a newly-added parser is picked up), and still skips
  unsupported source-only months (Void Forms etc.) rather than loudly re-failing.
- **Refresh-in-place + Teams task cards - the fix/refresh loop (2026-09-16, owner).**
  Closes the "who re-runs, and how does nobody forget" gap without migrating to
  Synology (deferred). The loop: clerk pings the owner -> owner runs it -> the
  script posts a Teams task card PER open vendor-month -> clerk reacts ✅ on each
  as she finishes -> owner runs `--refresh` -> re-checks and posts "clean -> mark
  DONE" or the still-open cards.
  - **`--refresh`**: re-reconciles every OPEN (non-DONE) vendor-month IN PLACE
    (the source statement is already filed there) - nothing is moved back to the
    Inbox. `_gather_open_sources` skips DONE months, the Inbox, the recon
    Excels, and any source file whose OWN name carries a `DONE` marker (the clerk
    can finalize a single statement inside a still-open month by renaming the
    file, not just the folder - QC 2026-09-23). So after the clerk fixes bills /
    prints, one `--refresh` re-derives every open month against current QBO;
    fixed items leave their buckets.
  - **Teams cards, ONE per vendor-month** (owner: "needs to be multiple messages
    so she can react to each"). `shared/teams_notify.py` (the ONE shared poster;
    invoice-sync keeps its own) posts an Adaptive Card per open vendor-month with
    the fix-list (to enter / not approved / mismatch / unprinted). Grouped by
    (vendor, month) so two statements in a month = one card. On `--refresh`, a
    month that comes back with nothing open gets a green "clean -> mark DONE" card.
    Webhook = `TEAMS_STMT_WEBHOOK` in the automation-qbo keychain blob (key-library
    rule; `setup_qbo` KeySpec added, required=False). No webhook -> silent no-op,
    so nothing breaks before the channel exists.
  - Clerk marks a card done with a ✅ reaction (zero Teams setup); a real Submit
    button is a later Power Automate add-on. Verified: `--refresh --dry-run` gathers
    the open months; the card builder + no-op path unit-tested. Live Teams post is
    untested until the owner creates the channel Workflows webhook.
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
- **Statements the board cannot place:** a statement filed in a folder not named `MM-YYYY` (e.g. `Core Concrete Pumping/SEPT 2026/`) gets no vendor row - rename the folder to `09-2026`. Unsupported layouts (Core Pumping, Sunbelt past-due / due-notice PDFs) show their vendor as Unreadable until a parser exists.
- **Teams digest** posts only if `TEAMS_STMT_WEBHOOK` is set; first live run used `--no-teams` while the owner decides its use.
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
