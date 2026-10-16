# bill-tracker/ - how the AP run works

Last changed: 10/16/2026 - **the clerks' typing can no longer be wiped**: the tracker is opened ONCE (`read_existing`, shared/xlsx_read); a tracker that exists but cannot be opened stops the run (exit 4, nothing written, the server alerts) instead of reading as "no Lien / Notes / marks" and being saved over. Backups copy only a tracker that just opened: `_backups/Bill Tracker - previous.xlsx` (just before this run) + the day's FIRST dated copy (no longer overwritten every run). The save goes through `shared/xlsx_save.save_verified` (built aside, `validate_xlsx` + `xlsx_verify`, then written into the file, permissions kept); a failed check writes nothing (exit 2) - it used to save the bad file anyway. Earlier: 10/14/2026 - **RP stage-invoice match**: on a job billed in stages (`shared/rp_invoicing`) a bill goes to the FIRST stage invoice dated on/after it that covers it (basis `stage`), before the largest-cover rule (RP7597 K783197 case). Same day: **Approved is one icon** on Bills + Inventory (green check / red cross / amber `?` for check QBO; `shared/bill_approval.state_icon`, the ledger loader maps it back to the word). Same day: the Bills and Inventory sheets gain a **Bill Memo** column right after Line Description (the bill's PrivateNote, one line, the same on every line of the bill; the bill-grain collapse now carries it; the Bill List already had it); column widths are keyed by header name so an inserted column never shifts them. Earlier: 10/13/2026 - the Bill List's Approved column gains **pending approval / requires approval** from QuickBooks' approval report (`shared/bill_approval.current()`, read once per run; the run prints when the report synced); a bill the fresh report does not list is approved; the stale-approval audit counts the queued bills too. Earlier: 10/12/2026 - first chart for this tool. The REVISED-bill follow-up built here this morning moved the same day into the statement reconciler (owner: "in the same statement Excel, not a mix"); `bill_rows.revised_reason` now lives in `shared/bill_approval` (re-exported) so both tools read the REVISED memo line the same way.

Update this chart - and the line above - in the same commit as any change to `bill-tracker/`
(`.github/flow_guard.sh`). The full system map is `docs/ARCHITECTURE.md`.

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef stop fill:#fdecec,stroke:#c0392b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111

    QBO[("QuickBooks<br/>every bill incl. subs, invoices, POs")]:::src
    GL[("General List xlsx<br/>read-only")]:::src
    POT[("PO tracker xlsx<br/>read-only")]:::src
    LOCK{"Bill Tracker open in Excel?<br/>~$ lock beside it"}:::gate

    BT["excel_bill_sync.py<br/>Bills · Liens · Inventory · Bill List<br/>+ Audit - Coding / History / PO / Bills<br/>(Not Approved · Revised · Duplicate)"]:::tool
    BR["bill_rows.py<br/>the ONE reading of a bill:<br/>approval state (memo + QuickBooks' approval report) + REVISED reason (shared/bill_approval) · project"]:::tool
    JCA["job_coding_audit.py<br/>audit-job: one job's bills"]:::tool

    BX[("Bill Tracker.xlsx<br/>Accounting / Accounts Payable<br/>typed Lien tags + Notes carried forward")]:::out
    CCH[("cost_code_history.json<br/>CompanyHealth - miscode log")]:::out

    QBO --> BT
    BT --> BR
    MATCH{"RP bill -> client invoice<br/>billed in stages? FIRST stage invoice on/after the bill that covers it<br/>else the largest covering invoice within 60 days, then the General List draw rule<br/>CP / MFD: the draw whose period holds the bill date"}:::gate
    BR --> MATCH --> BT
    GL -- "RP draw matching" --> MATCH
    POT -- "unused PO reconcile" --> BT
    BT --> LOCK -- "closed" --> RD{"read_existing<br/>the tracker opens?"}:::gate
    RD -- "cannot open: stop, write nothing, alert" --> STOP1["exit 4"]:::stop
    RD -- "opens: Lien / Notes / marks read, backups taken" --> SV{"save_verified<br/>built aside · validate_xlsx + xlsx_verify"}:::gate
    SV -- "fails: tracker untouched" --> STOP2["exit 2"]:::stop
    SV -- "passes: written into the file" --> BX
    LOCK -- "open: write nothing, next run" --> BT
    BT -. "miscode log, read + write" .-> CCH
    QBO -. "audit-job" .-> JCA
```

Runs on the office server every 15 minutes (docker/scheduler.py, AP then AR) while the writer file names the
server; the Mac's `sync-ap` stands down.
