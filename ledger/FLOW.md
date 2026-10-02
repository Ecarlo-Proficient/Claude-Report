# ledger/ - how the Project Ledger works

Last changed: 10/01/2026 - vendor page round 2: payment bills carry their scan count (📎); stubs open in the
ledger; bills tick to copy. Earlier the same day: every bill row carries its QuickBooks memo AND due date from the copy (`MIR -> SRV`); the
Vendor Center's Unpaid view shows Due; its columns drag to a remembered order; the bill viewer's scan fills its pane.
Earlier the same day: the process registry moved out of the gear: Company -> **Processes**, grouped by area
(Accounts Payable, ...), plain process names, no registry codes. Earlier, 09/30/2026 - `refresh_mirror.py` also runs on the office server (docker/), where there is no Pay run to
watch. Earlier the same day: the bill viewer's pay status from the mirror; the viewer's job share + zoomable scan.

Update this chart - and the line above - in the same commit as any change to `ledger/`
(`.github/flow_guard.sh` fails the push otherwise). The full system map is `docs/ARCHITECTURE.md`.

```mermaid
flowchart TD
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef db fill:#f3ecfb,stroke:#8e5cc4,color:#111
    classDef ui fill:#eaf7ea,stroke:#3c8d40,color:#111
    classDef write fill:#fdecec,stroke:#c0392b,color:#111

    subgraph SRC["Sources - read only"]
        MIR[("QuickBooks copy<br/>shared/qbo_mirror")]:::src
        QBO[("QuickBooks Online")]:::src
        WIPX[("WIP - MASTER.xlsx<br/>Test tabs")]:::src
        BTX[("Bill Tracker.xlsx")]:::src
        NOT[("Notion<br/>invoices, customers, liens")]:::src
        VAULT[("AI Brain_Vault<br/>process registry")]:::src
    end

    subgraph LOAD["Loaders - sync_all.sh / the gear's Resync"]
        RM["refresh_mirror.py"]:::tool
        LW["load_wip_master.py"]:::tool
        LB["load_bill_tracker.py"]:::tool
        LC["load_costs.py - every job, 90-day window<br/>load_invoices.py - + QuickBooks billed history<br/>load_payments.py · load_bill_payments.py"]:::tool
        LH["load_health.py · load_customers.py<br/>load_uncleared_checks.py · load_sub_loc.py"]:::tool
    end
    QBO --> RM --> MIR
    WIPX --> LW
    BTX --> LB
    MIR --> LC
    QBO --> LH
    NOT --> LH

    DB[("ledger.sqlite3<br/>Application Support/Proficient")]:::db
    LW & LB & LC & LH --> DB

    SRV["dashboard.py<br/>127.0.0.1 only · /api/*"]:::tool
    DB --> SRV

    subgraph AUD["Audit pages - live over the QuickBooks copy"]
        QA["QBO changes<br/>the change log"]:::tool
        CD["check_drift.py<br/>Checks QBO changed"]:::tool
        SH["strip_history.py<br/>History + QBO report"]:::tool
        DC["dup_customers.py<br/>Duplicate customers"]:::tool
    end
    MIR --> QA & CD & SH & DC
    MIR -->|"bill memo + due date"| SRV
    QA & CD & SH & DC --> SRV

    subgraph UI["static/ - four views + the gear"]
        PJ["Projects<br/>the WIP as the page + project page<br/>(P&L, Draw overview + fx bar, bills)"]:::ui
        VD["Vendors<br/>Vendor Center · Bill Tracker · Pay bills"]:::ui
        CU["Customers<br/>Invoices · Customer Center · Payments · Sales"]:::ui
        CO["Company<br/>Money · Bills to fix · QBO changes ·<br/>Checks QBO changed · Uncleared · Duplicate customers ·<br/>Processes"]:::ui
        GR["Gear<br/>sync · WIP Review · Console"]:::ui
    end
    SRV --> PJ & VD & CU & CO & GR
    VAULT -->|"registry_view.py · vault_graph.py"| CO

    subgraph WR["WIP Review - subprocess + JSON, never an import"]
        EM["wip tools --emit-review<br/>diff, no write"]:::tool
        RV["Passed · Check · Blocked · Held<br/>the owner approves"]:::ui
        AP["wip tools --apply-review<br/>trust gate + guarded writer"]:::write
    end
    GR --> EM --> RV --> AP --> WIPX

    subgraph WRITES["The only writes"]
        MK["owner marks<br/>waiver · lien · pay run · notes"]:::write
        RA["reapply_check.py<br/>Re-apply in QuickBooks"]:::write
        PB["pay_bills.py<br/>Pay in QuickBooks"]:::write
        PR["presence.py<br/>Touch ID every write"]:::write
    end
    PJ & VD --> MK --> DB
    CO --> RA
    VD --> PB
    RA & PB --> PR
    PB -->|"then the billpay refresh: mirror + bill payments + open AP"| RM
    PR -->|"dry run, then Are you sure"| QBO
```
