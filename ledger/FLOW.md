# ledger/ - how the Project Ledger works

Last changed: 10/07/2026 - process guides show people's names from `names.js` beside them in the vault's assets/processes/ (gitignored, the owner types the names once; every guide fills each role with its name). The ledger inlines it when it serves a guide (`registry_view.inline_names`); the guide on disk stays roles-only; Guide opens the html first when it has name slots. Earlier the same day: the bill payment stub drops the date under its title (Payment date is top right) and both stubs carry a divider above a full-size Amount. Earlier the same day: the stub column picker shows the stub's own labels (Open balance · Paid · New balance) and the vendor stub prints exactly the ticked columns. Earlier the same day: vendor page **Print stub** offers a **Vendor stub** (partials as columns: Open balance - Paid = New balance, no sentence) or an **Internal stub** (`bill_payment_stub.py` kind=internal: client + Joint check in the header, the client invoice(s) the money came from on top - a joint check matched to the client's Payment by date + amount - then the bills as Open balance - Paid = New balance; filed under `<vendor>/Internal/`); every print refreshes the mirror first, and a payment the mirror lacks also reloads bill payments and retries. Earlier the same day: process guides show people's names: a guide's `data-role` name slot is filled from the gitignored ROSTER.md when the ledger serves the page (`registry_view.roster_names` / `fill_names`; the file on disk stays roles-only; Guide opens the html first when it has slots). Earlier the same day: Company -> **Reclassify transactions** (`reclassify.py`, was one-offs/move_project_lines.py): preview every cost line on a project, tick lines, optional cost-code swap, Are you sure + Touch ID, then line-level writes (customer + the line's Projects tag), each document re-read live, backed up and proven after the write. Earlier: 10/06/2026 - sync-all stands down on AP + AR when Accounting/_automation/writer.json says "server" (the office server then owns them); the mirror refresh and the ledger reload still run here. Earlier: 10/06/2026 - Checks QBO changed splits floating checks into **Clear copy (mirror)** (tick and write in bulk) and **Hunt down**. **Clear copy**: the checks with a proven copy from before QuickBooks stripped them, fixed in bulk (`reapply_check.annotate` / `bulk_dry_run` / `bulk_commit`, one Touch ID per batch). Earlier the same day: gross billed tells retainage lines by the account the item posts to (Retainage Receivable), so a PC00 'Retainage' invoice is income, as in the project P&L. Earlier: 10/01/2026 - vendor page round 2: payment bills carry their scan count (📎); stubs open in the
Also, 10/01/2026 - `load_costs.py` maps every QuickBooks customer of a split-billed job (job rulings
`customers: all`, RP7401-FTW) to that one job; the Duplicate customers page no longer lists it.
Earlier the same day: vendor page round 2: payment bills carry their scan count (📎); stubs open in the
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
        VD["Vendors<br/>Vendor Center · Bill Tracker · Pay bills<br/>Print stub: vendor / internal (client invoice per bill)"]:::ui
        CU["Customers<br/>Invoices · Customer Center · Payments · Sales"]:::ui
        CO["Company<br/>Money · Bills to fix · QBO changes ·<br/>Checks QBO changed · Uncleared · Duplicate customers ·<br/>Processes"]:::ui
        GR["Gear<br/>sync · WIP Review · Console"]:::ui
    end
    SRV --> PJ & VD & CU & CO & GR
    VAULT -->|"registry_view.py · vault_graph.py"| CO
    NAMESJS[("assets/processes/names.js (gitignored)<br/>names by role, inlined at serve time")] -->|"guide names"| CO

    subgraph WR["WIP Review - subprocess + JSON, never an import"]
        EM["wip tools --emit-review<br/>diff, no write"]:::tool
        RV["Passed · Check · Blocked · Held<br/>the owner approves"]:::ui
        AP["wip tools --apply-review<br/>trust gate + guarded writer"]:::write
    end
    GR --> EM --> RV --> AP --> WIPX

    subgraph WRITES["The only writes"]
        MK["owner marks<br/>waiver · lien · pay run · notes"]:::write
        RA["reapply_check.py<br/>Re-apply in QuickBooks<br/>+ Clear copy bulk fix: proven copy only"]:::write
        PB["pay_bills.py<br/>Pay in QuickBooks"]:::write
        RC["reclassify.py<br/>Reclassify transactions<br/>cost lines project -> project, line level"]:::write
        PR["presence.py<br/>Touch ID every write"]:::write
    end
    PJ & VD --> MK --> DB
    CO --> RA
    CO --> RC
    VD --> PB
    RA & PB & RC --> PR
    PB -->|"then the billpay refresh: mirror + bill payments + open AP"| RM
    PR -->|"dry run, then Are you sure"| QBO
```
