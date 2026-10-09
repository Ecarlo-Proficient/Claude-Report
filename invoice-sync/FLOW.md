# invoice-sync/ - how the AR invoice sync and its Excel mirror work

Last changed: 10/08/2026 - Notion: an open invoice is written only when a QBO field differs from the page or it was not stamped Last Synced today (was: every run rewrote every open invoice); the run counts `unchanged`, a dry run names the changed fields. Earlier: 10/07/2026 - `qbo_payment_notify.py` runs standalone too (`--test` = "TEST" cards to TEAMS_WEBHOOK_PAYMENTS_TEST with their own posted record - the office server's test step; a dry run with no record lists the last 72h). Earlier the same day: QuickBooks payment received card: every payment a client makes through the
e-invoice link (bank / card), once, to the payments channel - read from the mirror, first run seeds.

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
    CHG{"page already matches QBO<br/>and Last Synced = today?"}:::gate
    SAME["unchanged - no Notion write<br/>(counted, at most one write per invoice a day)"]:::tool
    SKIP["skipped - no project #<br/>(lease / note / interest)"]:::tool
    LOCK{"Excel open?<br/>~$ lock file"}:::gate
    EXP["export_invoices_xlsx.py<br/>pull both trackers (all statuses)"]:::tool
    CHAIN["draw_chain.py<br/>previous draw per MFD/CP invoice"]:::tool
    AGE["aging_sheet.py<br/>client name -> invoices -> Total client (JPI on MFD: + project rows)<br/>Open Bal · Total · Due · Aging · Notion Notes · Lien · Litigation"]:::tool
    ALL["All Open<br/>every division + litigation + lease, Division filter"]:::tool
    LEASE["_lease_records<br/>open invoices with no project #"]:::tool
    NOTES["notes_preserve.py<br/>absorb Note -> Quick Status, or re-attach"]:::tool
    VER{"xlsx_verify.assert_clean"}:::gate

    MFD[("Notion<br/>MFD Invoice Tracker")]:::out
    RES[("Notion<br/>Res/Com Invoice Tracker")]:::out
    TEAMS[("Teams<br/>MFD paid / short-pay")]:::out
    PAYN["qbo_payment_notify.py<br/>QuickBooks Payments only, not seen before, under 72h<br/>(--test on the office server: TEST cards)"]:::tool
    TPAY[("Teams payments channel<br/>QuickBooks payment received")]:::out
    OD[("Invoice Tracker.xlsx<br/>All Open · CP / MFD / RP Aging · Lease Invoices ·<br/>Open Invoices · Cash Flow · Pay Calendar")]:::out
    HOLD[("skipped this run<br/>close the file, re-run sync-ar")]:::out

    QBO --> SYNC --> ROUTE
    ROUTE -- "MFD" --> CHG
    ROUTE -- "CP / RP" --> CHG
    CHG -- "yes" --> SAME
    CHG -- "no, MFD" --> MFD
    CHG -- "no, CP / RP" --> RES
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
    MIR --> PAYN --> TPAY
    NOTES --> AGE
    AGE --> ALL --> VER --> OD
    AGE --> VER
```
