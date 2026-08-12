# STATUS — invoice-sync

Shared progression record for the QBO → Notion AR invoice sync and its Excel
mirror. Update this in the SAME commit as any change to this tool.

---

## DONE / FINALIZED

- **Aging tab fixes (2026-08-12, the user):**
  - **Client → project sub-grouping.** A client with more than one project now
    gets a project sub-group row under it (outline: client 0 → project 1 → invoice
    2); a single-project client stays flat (client 0 → invoice 1). "See it based
    on client, then project." Roll-up sums + the Open/Total bar apply at each level.
  - **"First draw" no longer over-claims.** The draw chain is built only from
    invoices that passed through Notion's open-invoice sync, so a paid `Draw #1`
    that was never synced is invisible and its `Draw #2` looked like idx-0 "first."
    Now `CHAIN_FIRST_DRAW` fires ONLY when it's provably `Draw #1`; any other
    earliest-seen draw returns `CHAIN_PREV_UNKNOWN` → verdict **"Prev not synced"**
    (greyed, honest) instead of a false "First draw." Also fixed `contract_label`
    to strip a trailing `- Retainage` AND the draw tail in either order, so a
    `Draw #8 - Retainage` chains onto its base contract instead of splitting off.
  - **Retainage is on the lien clock.** Notice due 30 days after the invoice date
    (was "own track, undated"). See shared/lien_clock + the retainage memory. The
    lien column shows `RET DUE <date> · Nd` etc.; money_bleeds updated to match.
- **Collections tab upgrades (2026-08-11, the user):**
  - **QBO links fixed — company-scoped deep links.** `qbo_client.invoice_deep_link`
    now builds Intuit's own `/app/login?pagereq=invoice%3FtxnId%3D<id>&deeplinkcompanyid=<realm>`
    form. The old bare `app/invoice?txnId=` link carried no company, so with more
    than one Intuit company on the login it opened a *different* company's
    transaction — the "random invoice" the owner hit. Realm comes from the loaded
    creds (never source/logs). Fixed at all writers: sync props, MFD Teams card,
    paid-flip fallback, `resend_mfd_paid`. Notion + Excel + Teams all inherit it
    on the next sync. **Verified** invoice 34431 → txnId 1313157 against the QBO API.
  - **Aging tabs: `Open Balance` then `Total Amount`, with a per-row data bar.**
    Open balance first, invoice total beside it; a blue data bar on Open Balance
    scaled by FORMULA to that row's Total Amount cell, so its fill = open ÷ total
    (full = untouched, short = partly collected). Summary + grand rows scale to
    their own sums. `C_INVTOTAL` added to the grid; `total_amount` added to the
    aging record (falls back to open balance so the bar never divides by zero).
  - **Excel Notes are a two-way status channel** (`notes_preserve.py`). The clerk
    writes **Notes** (legacy yellow sticky, author attached — NOT threaded
    Comments, which openpyxl can't read; verified the live file has zero threaded
    comments). Two scopes, both anchored on the Invoice # column: per-invoice on a
    detail row, per-client on a summary `N inv` cell (covers all that client's open
    invoices; a per-invoice Note wins its own row). Two modes:
    - **PRESERVE (default)** — re-attach every Note verbatim so nothing is lost on
      the rebuild. No Notion writes. Round-trip tested on the live file: 11
      per-invoice + 3 per-client, exact.
    - **ABSORB (default for sync-ar; `ABSORB_NOTES=0` disables; the user 2026-08-11)**
      - a Note IS the status. Its text (stamped ` – Name, M/D`, en dash never em
      dash) replaces the Notes column, the cell Note is dropped, and it's **pushed
      to Notion `Quick Status`**; the prior Quick Status is archived (dated) to the
      page body (the documented **Collection Log**). This is the first time the sync
      writes a human-owned field: deliberate, and **idempotent** (absorb only emits
      a change when text differs, so re-runs push nothing / no duplicate log lines).
      **Safety: a Note absorbs only if its Notion push SUCCEEDS** - a failed push
      keeps the cell Note (re-attached, never lost) and retries next sync.
    - `preview_export.py` runs **absorb + dry-run** against a throwaway path
      (seeded from the live file): shows the absorbed Notes column, the removed
      cell Notes, and logs the exact `Quick Status old → new` it WOULD push —
      writing nothing to Notion or the live file. Absorb + push verified by
      round-trip and mock-client payload tests.
- **QBO → Notion AR sync** — open invoices to two Notion DBs (MFD isolated;
  Res/Com combined) routed by project-# prefix; sweeps paid; archives
  QBO-deleted via CDC; posts MFD pay events to Teams. Manual via `sync-ar`.
- **Excel mirror** — `export_invoices_xlsx.py` writes `Open_Invoices.xlsx` to
  OneDrive `Collections/`. Skips cleanly when Excel has the file open
  (`~$` lock check) so AutoSave can't clobber the update.
- **Aging tabs (2026-08-05, split per division 2026-08-10)** — built by
  `aging_sheet.py`. The owner's ask: Notion reads fine one page at a time but
  can't be scanned as a hundred rows, so the aging view lives in Excel.
  - **One tab per division: `CP Aging`, `MFD Aging`, `RP Aging`** (the user
    2026-08-10 — "keep cp and mfd separated"). **No Division column** — the tab
    is the division. The combined `AR Aging` tab is gone with it.
  - **`Lien` column replaced `Days Past Due`** (the user 2026-08-10). Days past
    due was already legible from which bucket the money sits in; the lien
    deadline is not, and it is the one that EXPIRES. Shows the date a Ch. 53
    notice must be MAILED by, with urgency banding: `PAST DUE` (reversed white
    on dark red), `DUE <date> · Nd` (≤15 days), `<date> · Nd` (≤45), plain date
    otherwise, `Notice sent`, and `Retainage — own track`.
  - Buckets Current / 1-30 / 31-60 / 61-90 / 90+, aged **by due date** (same
    rule `invoice_sync.py` uses for the Notion `Aging Bucket` select, and the
    same basis as QBO's AR aging).
  - RP + CP + MFD in one table with a `Division` column for filtering. Parent
    rows carry a division too (`(mixed)` when a client spans divisions) so the
    Division filter never orphans a group header.
  - Invoices grouped under the **parent client** (Notion `Customer` relation →
    Customer List title, not the project-level `Customer (raw)`), **collapsed
    by default** via row outline with `summaryBelow = False`.
  - `Notes` = the collections clerk's Notion `Quick Status`, plus their
    `Last Action Date`.
  - **Litigation invoices excluded** (the `Litigation` checkbox on both
    trackers) — legal work, not collections work, and leaving them in inflates
    every bucket. The count of what was dropped prints in the subtitle.
  - **`Invoice #` is a hyperlink to the invoice in QBO** (Notion's `QBO Link`
    url property, written by the sync). On the number rather than in its own
    column — keeps the sheet narrow and puts the click where the eye already is.
    A handful of MFD invoices have no link in Notion; those stay plain text.
  - **Type sizes: body 12pt, client rows 13pt, title 14pt** (the user
    2026-08-05). Client rows carry the extra point because the sheet is read
    COLLAPSED — the client name is the line that has to land first. Column
    widths are computed by `_autofit`, which measures the strings actually
    written and scales for 12pt (openpyxl can't autofit: Excel sizes columns at
    render time and openpyxl never renders). Capped at `MAX_COL_WIDTH` so a long
    memo can't create a column you have to scroll past.
  - **Colour (2026-08-05, the user's explicit ask — a named exception to the
    repo's plain-Excel rule, see CLAUDE.md rule 5).** It encodes age or state
    only: bucket headers green→red; each detail row tints the one bucket cell
    holding its balance, so colour drifting rightward = money getting older;
    blue banding on client summary + all-clients rows; `Vendor Unpaid Bills`
    red, `Vendors Paid` green. A KEY row and a one-line note sit below the data.
    **Do not restyle this tab back to plain.**
  - **RP vendor cells read `n/a` on grey with darker grey italic text** — a
    blank read as "not looked up yet" when it means "nothing to look up".
    Applied to all-RP client summary rows too, or an RP-only client looked like
    it had cleared its vendors.
  - **`Prev Draw` / `Prev Draw Status` / `Prev Bills Open` / `Prev $ Open`
    (2026-08-05) — the PREVIOUS draw, not this one.** The funding chain: the GC
    funds draw N → we pay draw N's vendor bills → those vendors issue
    unconditional waivers → the GC releases draw N+1. So an unpaid draw is
    gated by its predecessor, and the verdict separates the two holds:
    `PAY BILLS → unlock` (prev draw funded, our vendors still owed — **ours** to
    fix, red) vs `Waiting GC on prev` (prev draw unfunded too — upstream).
    Chain built by `draw_chain.py`; `Prev Draw` names the invoice used so a bad
    pick is visible rather than hidden. `This Draw $ Open` keeps the original
    same-draw figure alongside.
  - Bill data read from the bill-tracker's `Bill Tracker.xlsx` output file
    (never its code — repo rule 3, tools never import tools), deduped to bill
    grain because that sheet is line-level and repeats `Bill Open Bal` on
    every line of a bill.

## IN PROGRESS

- **Absorb is LIVE by default (the user 2026-08-11).** `sync-ar` now absorbs Notes
  and pushes Quick Status. Watch the FIRST live run: confirm the prior Quick Status
  lands in the page Collection Log and the note text becomes Quick Status, and that
  no push-failure warnings appear (which would mean Quick Status is not the rich_text
  type this assumes). `python3 preview_export.py` is the always-safe dry-run preview.
  `ABSORB_NOTES=0 sync-ar` falls back to preserve mode if needed.

## TO DO

- Collection Log lines append oldest-first (Notion has no prepend). If the owner
  wants newest-first, that needs a read-reorder-rewrite of the page body.

## OPEN ISSUES

- **THE LIEN COLUMN IS A WATCHLIST, NOT LEGAL ADVICE — and its weakest link is
  the work month.** Every Ch. 53 deadline runs from the month the labor was
  furnished. The clock here uses the **invoice month**, which is the owner's
  settled ruling (2026-07-16: RP invoices go out the day the job finishes,
  draws bill their own work month; "never re-add a conservative offset" after
  the first build produced month-early false alarms). **But MFD/CP draw memos
  routinely state a period that starts in the PRIOR month** — e.g. MFD177 inv
  34318, "May Draw 2026 (Period: 04/02/2026 - 05/01/2026)". Read as April work,
  its notice deadline was Jul 15, not Aug 14. The stated period is sitting right
  there in the memo and is not currently used. Raised with the owner 2026-08-10;
  the ruling stands until they say otherwise.
- **Retainage and lease exclusions are memo-text only here.** `money_bleeds`
  detects both off QBO line items; this tab only has the Notion memo and the
  clerk's note, so a retainage invoice whose memo doesn't say "retainage" gets a
  monthly deadline it may not be governed by (one live row reads "Retaiange").
  Equipment-lease/note invoices to subs — excluded from the clock by the
  2026-07-16 ruling — are **not detectable at all** from memo text; none are in
  the current open set, but that is luck, not a guarantee.
- **"Notice sent" is inferred from the clerk's free-text `Quick Status`.** The
  pattern is deliberately narrow (it must name the notice/affidavit, so
  "waiting vendor unconditional" doesn't count), but the real record is the
  Notion `Lien Tracker` (Admin) DB, which nothing feeds automatically. Wiring
  that in would replace a heuristic with a fact.

- **RUN ORDER: AP BEFORE AR (2026-08-05).** The aging tab's vendor columns read
  `Bill Tracker.xlsx`, so `sync-ap` must finish before `sync-ar` or those
  columns silently report the *previous* AP run. `sync-all` in `~/.zshrc` was
  written AR-then-AP back when the two were independent; it has been swapped to
  **AP → AR** and is now the command to use daily (~5 min: AP ≈ 3.5 min,
  AR ≈ 1-2 min). **There is no cycle** — AP pulls its bills *and* its invoices
  straight from QBO and never reads Notion or `Open_Invoices.xlsx`, so the
  dependency is one-way.
- **Staleness is surfaced, not assumed.** If the tracker predates today,
  `aging_sheet.py` logs a warning and the tab subtitle turns red with
  "⚠ VENDOR COLUMNS ARE N HOURS OLD". A missing/unreadable file yields `?` per
  invoice rather than a false "Vendors Paid". AR never auto-runs AP — a sync
  with hidden side effects is worse than a stale column that announces itself.
- **RP has no previous-draw block by design.** RP doesn't bill in draws, and
  the bill-tracker matches RP bills on "earliest invoice on/after bill date" —
  not a draw period. Rendered as a grey `n/a` block on the combined tab, and
  **omitted entirely on the `RP Aging` tab** (the user 2026-08-05 — spreadsheet
  columns N through R). Both tabs come from one `build_aging_sheet`; the RP one
  passes `drop_columns=RP_DROP_COLUMNS` and `_Grid` handles the projection, so
  rows are still built at full width against the `C_*` constants and no caller
  has to know which physical column a field landed in.
- **MFD192 and CP861 report `Multi-contract` instead of a previous draw.**
  MFD192 runs three contracts in parallel (base, HUDSONWOOD, OFFSITE) and CP861
  carries two different jobs (a 7-Eleven and a BP) under one project #; their
  draws interleave by date. **Bills carry a project #, not a contract**, so
  which bills belong to which contract can't be determined — the user
  2026-08-05: *"MFD192 - oddball, i have no definite way to see what bills
  belong to which contract yet."* Reported as unattributable rather than split
  on a guess. If bills ever become contract-attributable, `draw_chain.py`
  already keys chains by (project, contract) and only the guard needs removing.
- **Draw sequencing is memo-parsed, which is the fragile part.** `draw_chain.py`
  strips the draw designator to identify the contract, and the memos are not
  uniform: `May Draw 2026`, `Draw #2`, `Draw # 6`, `Draw #3 December 2024`,
  `March 2025 Draw`, `- Retainage - Draw #5`, and periods written both
  `(Period: …)` and `- Period:…`. All are covered and unit-checked against live
  memos; a new spelling would show up as a project splitting into phantom
  contracts. Invoices whose memo names no draw (retainage releases, turnkey
  flatwork) are excluded from chains — MFD177's `City Retainage` sits between
  two draws by date and would otherwise be picked as "the previous draw".
- **Zero-balance rows survive** into the aging tab when Notion still has an
  invoice marked Unpaid with a $0 open balance (QBO/Notion drift between syncs).
  Kept deliberately — the row is visible drift the collections clerk should
  see, and it adds $0 to every bucket.
- The 15-min launchd schedule is still paused (macOS update broke it); plists
  live in `launchd/` as `.disabled`.
