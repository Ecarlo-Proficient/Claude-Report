# shared/ - how the common package works

Last changed: 10/02/2026 - `pnl_paths._find_awarded_cp_folder` matches a finished job's folder on its name as
written ("CP656 - 77 PAUL WILSON" was lost when the spaces were squeezed out and the street number glued onto the
job #). Earlier, 10/01/2026 - `qbo_vault.has_credentials()` on Linux checks only the four QBO keys (JT_GRANT_KEY is
optional; the office server has none). Earlier, 09/30/2026: NEW `qbo_trust.py`, the QuickBooks trust gate every WIP write passes
(duplicate customers, name typos, costs with no project, flatwork on the slab, a stale copy). It
also owns `duplicate_groups`, which the ledger's Duplicate customers page reads. Same day: `pnl_paths`
finds a finished CP job's folder under `Completed Projects/<year>/` (active awarded folders first). Same day: `notion_client.py` keeps only the block edits the statement reconciler's Notion board uses (append / delete children, read a data source); `teams_notify.py` posts ONE statement digest card per run - since the evening of 09/30 it lists only the vendors NOT ready for a pay run (a line each + a count of the ready ones).

`shared/` is the ONLY importable common code (repo rule): a tool imports from here, never from
another tool. Update this chart - and the line above - in the same commit as any change to
`shared/` (`.github/flow_guard.sh` fails the push otherwise).

```mermaid
flowchart LR
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef mod fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef use fill:#eaf7ea,stroke:#3c8d40,color:#111

    KH[("Key Helper<br/>keyhelper/ - the key broker")]:::src
    QBO[("QuickBooks Online<br/>production")]:::src
    NOT[("Notion")]:::src
    FILES[("Synology / OneDrive / SharePoint<br/>takeoffs, draws, schedules")]:::src

    subgraph KEYS["Keys and QuickBooks access"]
        KB["key_broker.py<br/>one-hour QBO pass"]:::mod
        VAULT["qbo_vault.py<br/>the key library"]:::mod
        API["qbo_api.py<br/>auth, retrying GET, query_all,<br/>report walkers, PROJ_RE,<br/>build_project_customer_map"]:::mod
        SETUP["setup_qbo.py<br/>--status / --rotate"]:::mod
    end
    KH --> KB --> API
    KH --> VAULT
    QBO --> API

    subgraph MIRROR["The one read of QuickBooks"]
        MIR["qbo_mirror.py<br/>raw copy + change log"]:::mod
        CACHE["qbo_cache.py"]:::mod
        ATT["qbo_attachments.py"]:::mod
    end
    API --> MIR
    MIR -->|"query_all answers from the copy"| API

    subgraph COSTS["Job cost engine"]
        JL["job_lines.py<br/>does this line belong to this job"]:::mod
        QC["qbo_costs.py<br/>cost_leaf + iter_cost_lines"]:::mod
        CL["cost_lines.py · cost_code_audit.py"]:::mod
        JR["job_rulings.py<br/>per-job rulings register"]:::mod
        PL["qbo_pl.py<br/>company P and L totals"]:::mod
    end
    MIR --> QC
    JL --> QC

    subgraph TRUST["QuickBooks trust gate"]
        QT{"qbo_trust.py<br/>stale copy · duplicate customer ·<br/>name typo · cost with no project ·<br/>FW on the slab"}:::gate
    end
    MIR --> QT
    JR --> QT

    subgraph WIPH["WIP helpers"]
        WC["wip_contracts.py · wip_audit.py<br/>takeoff_etc.py · draws.py · draw_moves.py"]:::mod
        RPP["rp_price.py · rp_billing.py · rp_invoicing.py<br/>proposals.py · schedule.py · schedule_index.py"]:::mod
        JT["jobtread.py"]:::mod
    end
    FILES --> WC & RPP

    subgraph LIEN["AP / AR helpers"]
        LC["lien_clock.py · lien_status.py<br/>bill_marks.py · sub_loc.py"]:::mod
        NC["notion_client.py · notion_customers.py<br/>pages, block edits"]:::mod
        TN["teams_notify.py<br/>statement digest: vendors not ready to pay"]:::mod
    end
    NOT --> NC

    subgraph MODELS["Finance models"]
        FM["breakeven.py · recurring.py · bizdev_cut.py"]:::mod
    end

    subgraph SAFE["Every Excel file"]
        XG["xlsx_guard.py<br/>outside text never a formula"]:::mod
        XV["xlsx_verify.py<br/>assert_clean - no repair prompt"]:::mod
        PATHS["paths.py · pnl_paths.py<br/>per-machine paths, CompanyHealth layout,<br/>CP job folders incl. Completed Projects"]:::mod
    end

    WIPT["wip/<br/>the WIP update"]:::use
    LEDG["ledger/<br/>the project database + UI"]:::use
    PNL["project-pnl/"]:::use
    BT["bill-tracker/ · invoice-sync/"]:::use

    API --> WIPT
    QT -->|"holds costs / billed"| WIPT
    QT -->|"duplicate_groups"| LEDG
    QC --> PNL & LEDG
    JR --> WIPT & PNL & LEDG
    PL --> LEDG
    WC & RPP --> WIPT
    LC & NC --> LEDG & BT
    XG & XV --> WIPT & PNL & BT & LEDG
```
