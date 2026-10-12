# bill-tracker/ - how the AP run works

Last changed: 10/12/2026 - NEW `revised_bills.py`, the REVISED-bill follow-up as a process (AP-13): every AP run finds each PM-revised bill in the QuickBooks copy (the `REVISED - <reason>` memo line), its original total and when the PM changed it (the change log), what the vendor's latest statement says (the reconciler's record), the pickles (entered twice · amount changed again after the clerk sent it · paid before re-issued · gone from QuickBooks), and keeps one row per bill on the Notion **Revised Bills** board (`shared/revised_board`); the clerk sets Stage + Sent / Replied / Received / Entered on Notion; the ledger reads the board (Vendors -> Revised bills). First chart for this tool.

Update this chart - and the line above - in the same commit as any change to `bill-tracker/`
(`.github/flow_guard.sh`). The full system map is `docs/ARCHITECTURE.md`.

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111

    QBO[("QuickBooks<br/>every bill incl. subs, invoices, POs")]:::src
    MIR[("QuickBooks copy + change log<br/>shared/qbo_mirror")]:::src
    GL[("General List xlsx<br/>read-only")]:::src
    POT[("PO tracker xlsx<br/>read-only")]:::src
    SREC[("&lt;Vendor&gt;/.statement-record.json<br/>statement-reconciler's record")]:::src
    LOCK{"Bill Tracker open in Excel?<br/>~$ lock beside it"}:::gate

    BT["excel_bill_sync.py<br/>Bills · Liens · Inventory · Bill List<br/>+ Audit - Coding / History / PO / Bills<br/>(Not Approved · Revised · Duplicate)"]:::tool
    BR["bill_rows.py<br/>the ONE reading of a bill:<br/>approval state · REVISED reason · project"]:::tool
    RV["revised_bills.py<br/>find revised bills in the copy ·<br/>original total + revised date from the change log ·<br/>what the statement says · flags ·<br/>upsert the board, Stage forward only"]:::tool
    JCA["job_coding_audit.py<br/>audit-job: one job's bills"]:::tool

    BX[("Bill Tracker.xlsx<br/>Accounting / Accounts Payable<br/>typed Lien tags + Notes carried forward")]:::out
    CCH[("cost_code_history.json<br/>CompanyHealth - miscode log")]:::out
    NB[("Notion · Revised Bills board<br/>1 row per revised bill: Stage (the clerk) ·<br/>Sent / Replied / Received / Entered on (the clerk) ·<br/>QBO total vs vendor original · Statement says · Flags")]:::out
    LEDGER[("the Project Ledger<br/>Vendors -> Revised bills: where each stands, for how long")]:::out

    QBO --> BT
    BT --> BR
    GL -- "RP draw matching" --> BT
    POT -- "unused PO reconcile" --> BT
    BT --> LOCK -- "closed" --> BX
    LOCK -- "open: write nothing, next run" --> BT
    BT -. "miscode log, read + write" .-> CCH
    QBO -. "audit-job" .-> JCA
    BT -- "after the save (ACB_REVISED_DS_ID set)" --> RV
    MIR --> RV
    SREC --> RV
    BR --> RV
    RV --> NB
    NB -. "/api/revisions, read live" .-> LEDGER
```

Runs on the office server every 15 minutes (docker/scheduler.py, AP then AR) while the writer file names the
server; the Mac's `sync-ap` stands down. `python3 bill-tracker/revised_bills.py` previews the board plan by hand
(`--write` writes; `--since` widens the scan).
