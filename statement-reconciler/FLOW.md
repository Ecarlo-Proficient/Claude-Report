# statement-reconciler/ - how the vendor statement reconciler works

Last changed: 09/30/2026 - follow-up moved from a Teams card per vendor-month to the Notion
"Vendor Statements" board: one live checklist page per vendor-month, rewritten every run, and
one Teams digest per run.

Update this chart - and the line above - in the same commit as any change to
`statement-reconciler/` (`.github/flow_guard.sh`).

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111

    INBOX[("Accounting share<br/>Vendor Statements / Statement Inbox<br/>PDF · Excel · image")]:::src
    QBO[("QuickBooks<br/>open + recently paid bills")]:::src
    MBX[("billings mailbox<br/>printed bills")]:::src

    PARSE["parse_statement<br/>~8 vendor templates + OCR"]:::tool
    VEND["find_vendor_id<br/>alias cache, then QBO search"]:::tool
    TIE{"tie-out<br/>lines sum to Amount Due?"}:::gate
    REC["reconcile_iter<br/>matched · missing in QBO · amount mismatch ·<br/>tax 8.25% · vendor lag · not on statement"]:::tool
    PS["print_status.py<br/>opt-in PRINT_STATUS=1"]:::tool
    ITEMS["_board_items<br/>one item per bill: enter · approve · check QBO ·<br/>amount · tax · print"]:::tool
    NB["notion_board.py<br/>merge with the clerk's ticks:<br/>fixed in QBO -> Cleared · ticked but still open -> unticked ·<br/>check QBO ticks kept"]:::tool

    XLSX[("Excel + source statement<br/>filed under Vendor / MM-YYYY")]:::out
    HELD[("tie-out failed:<br/>Excel banded, source stays in Inbox")]:::out
    PAGE[("Notion Vendor Statements<br/>1 page per vendor-month<br/>Open · In progress · Clean · Done")]:::out
    TEAMS[("Teams<br/>1 digest card per run")]:::out

    INBOX --> PARSE --> VEND --> REC
    QBO --> VEND
    QBO --> REC
    PARSE --> TIE
    TIE -- yes --> XLSX
    TIE -- no --> HELD
    REC --> XLSX
    MBX --> PS --> ITEMS
    REC --> ITEMS --> NB --> PAGE
    NB --> TEAMS
    PAGE -. "--sync-done: Done -> folder '<MM-YYYY> DONE'<br/>folder DONE -> page Done" .-> XLSX
    XLSX -. "--refresh re-checks every open month in place" .-> PARSE
```
