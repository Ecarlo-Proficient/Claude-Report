# bill-tracker/ - how the AP run works

Last changed: 10/13/2026 - the Bill List's Approved column gains **pending approval / requires approval** from QuickBooks' approval report (`shared/bill_approval.current()`, read once per run; the run prints when the report synced); a bill the fresh report does not list is approved; the stale-approval audit counts the queued bills too. Earlier: 10/12/2026 - first chart for this tool. The REVISED-bill follow-up built here this morning moved the same day into the statement reconciler (owner: "in the same statement Excel, not a mix"); `bill_rows.revised_reason` now lives in `shared/bill_approval` (re-exported) so both tools read the REVISED memo line the same way.

Update this chart - and the line above - in the same commit as any change to `bill-tracker/`
(`.github/flow_guard.sh`). The full system map is `docs/ARCHITECTURE.md`.

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
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
    GL -- "RP draw matching" --> BT
    POT -- "unused PO reconcile" --> BT
    BT --> LOCK -- "closed" --> BX
    LOCK -- "open: write nothing, next run" --> BT
    BT -. "miscode log, read + write" .-> CCH
    QBO -. "audit-job" .-> JCA
```

Runs on the office server every 15 minutes (docker/scheduler.py, AP then AR) while the writer file names the
server; the Mac's `sync-ap` stands down.
