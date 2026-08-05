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
  - `Vendor Status` / `Open Bills` / `Vendor $ Open` — MFD/CP only. Read from
    the bill-tracker's `Bill Tracker.xlsx` output file (never its code — repo
    rule 3, tools never import tools), deduped to bill grain because that
    sheet is line-level and repeats `Bill Open Bal` on every line of a bill.

## IN PROGRESS

- Nothing open.

## TO DO

- Nothing queued.

## OPEN ISSUES

- **`Vendor Status` freshness is bounded by `sync-ap`, not `sync-ar`.** The
  aging tab reads whatever `Bill Tracker.xlsx` last contained; its timestamp is
  printed in the tab subtitle, and the column shows `?` when the file is
  missing or unreadable rather than falsely reporting vendors paid. Run
  `sync-ap` before `sync-ar` when the vendor column needs to be current.
- **RP has no vendor column by design.** The bill-tracker matches RP bills on
  "earliest invoice on/after bill date", which is not a draw-period statement,
  so an RP flag there would imply a match the data doesn't support. Left blank.
- **Zero-balance rows survive** into the aging tab when Notion still has an
  invoice marked Unpaid with a $0 open balance (QBO/Notion drift between syncs).
  Kept deliberately — the row is visible drift the collections clerk should
  see, and it adds $0 to every bucket.
- The 15-min launchd schedule is still paused (macOS update broke it); plists
  live in `launchd/` as `.disabled`.
