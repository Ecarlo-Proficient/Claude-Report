# STATUS — invoice-sync

Shared progression record for the QBO → Notion AR invoice sync and its Excel
mirror. Update this in the SAME commit as any change to this tool.

---

## DONE / FINALIZED

- **QBO → Notion AR sync** — open invoices to two Notion DBs (MFD isolated;
  Res/Com combined) routed by project-# prefix; sweeps paid; archives
  QBO-deleted via CDC; posts MFD pay events to Teams. Manual via `sync-ar`.
- **Excel mirror** — `export_invoices_xlsx.py` writes `Open_Invoices.xlsx` to
  OneDrive `Collections/`. Skips cleanly when Excel has the file open
  (`~$` lock check) so AutoSave can't clobber the update.
- **AR Aging tab (2026-08-05)** — second sheet in the same workbook, built by
  `aging_sheet.py`. The owner's ask: Notion reads fine one page at a time but
  can't be scanned as a hundred rows, so the aging view lives in Excel.
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

- Nothing open.

## TO DO

- Nothing queued.

## OPEN ISSUES

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
