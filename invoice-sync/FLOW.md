# invoice-sync/ - how the AR invoice sync and its Excel mirror work

Last changed: 10/01/2026 - the aging tabs show Open Balance + Total Amount after Due Date, ONE Aging
column and the lien columns last, client -> invoices (JPI on MFD keeps project rows); new Lease
Invoices tab from the QBO mirror (open invoices with no project #).

Update this chart - and the line above - in the same commit as any change to `invoice-sync/`
(`.github/flow_guard.sh`).

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111

    QBO[("QuickBooks<br/>open invoices + change feed")]:::src
    MIR[("QBO mirror<br/>shared/qbo_mirror")]:::src
    BTX[("Bill Tracker.xlsx<br/>read-only")]:::src
    OLD[("current Open_Invoices.xlsx<br/>clerk's cell Notes")]:::src

    SYNC["invoice_sync.py<br/>parse each open invoice"]:::tool
    ROUTE{"project # in customer<br/>or memo?"}:::gate
    SKIP["skipped - no project #<br/>(lease / note / interest)"]:::tool
    LOCK{"Excel open?<br/>~$ lock file"}:::gate
    EXP["export_invoices_xlsx.py<br/>pull both trackers (all statuses)"]:::tool
    CHAIN["draw_chain.py<br/>previous draw per MFD/CP invoice"]:::tool
    AGE["aging_sheet.py<br/>client -> invoices (JPI on MFD: + project rows)<br/>Open Bal · Total · Aging · draws · Notes · Lien last"]:::tool
    LEASE["_lease_records<br/>open invoices with no project #"]:::tool
    NOTES["notes_preserve.py<br/>absorb Note -> Quick Status, or re-attach"]:::tool
    VER{"xlsx_verify.assert_clean"}:::gate

    MFD[("Notion<br/>MFD Invoice Tracker")]:::out
    RES[("Notion<br/>Res/Com Invoice Tracker")]:::out
    TEAMS[("Teams<br/>MFD paid / short-pay")]:::out
    OD[("Open_Invoices.xlsx<br/>Open Invoices · CP / MFD / RP Aging ·<br/>Lease Invoices · Cash Flow · Pay Calendar")]:::out
    HOLD[("skipped this run<br/>close the file, re-run sync-ar")]:::out

    QBO --> SYNC --> ROUTE
    ROUTE -- "MFD" --> MFD
    ROUTE -- "CP / RP" --> RES
    ROUTE -- "none" --> SKIP
    SYNC --> TEAMS
    MFD --> LOCK
    RES --> LOCK
    LOCK -- "open" --> HOLD
    LOCK -- "closed" --> EXP
    OLD --> NOTES
    EXP --> CHAIN --> AGE
    BTX --> AGE
    MIR --> LEASE --> AGE
    NOTES --> AGE
    AGE --> VER --> OD
```
