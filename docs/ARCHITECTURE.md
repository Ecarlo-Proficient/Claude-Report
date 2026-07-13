# Automation Suite — System Map

> **Maintenance rule (binding):** any commit that adds/removes a script, changes what a script
> reads or writes, or rewires a data flow MUST update this diagram in the same commit.
> Claude sessions: check this file whenever `git status` shows script changes at session end.
> GitHub renders the Mermaid blocks below natively — view this file on github.com to see the diagrams.

Last updated: 2026-07-13 (**the restructure**: `shared/` package created; `automation-worker/`
split into `invoice-sync/` + WIP readers joined `wip/`; Field Log subsystem erased; loose root
scripts rehomed to `debt-schedule/`, `qbo-export/`, `one-offs/`; Docker → v1.1.0)

---

## The structure rules (see CLAUDE.md for the full text)

**One folder = one tool. `shared/` is the only importable common code. Tools never import
tools. One-offs live in `one-offs/`, never the repo root. `machine.env` stays at the repo root.**

```
shared/          qbo_vault · paths · qbo_api · setup_qbo      ← the ONLY shared code
invoice-sync/    QBO → Notion AR sync (was automation-worker/)
bill-tracker/    AP bills → Excel tracker + 4 audits
statement-reconciler/  vendor statement PDF ↔ QBO bills
wip/             CP/RP WIP readers + gated close scripts      ← ALL WIP tooling
project-pnl/     per-project P&L workbooks
debt-schedule/   equipment debt workbook + loan_sync
health-dashboard/ local health xlsx
qbo-export/      transaction export
one-offs/        occasional tools (qbo_recode_review)
synology/        NAS file-tree audit
docker/          invoice-sync container (v1.1.0)
```

---

## Big picture

Every QBO script authenticates through the one shared Keychain vault (single Touch ID per run).
QBO is production-only and the source of truth for costs/billed. Writers are gated
(`--commit` / `CONFIRM=Y`); everything else is read-only against QBO.

```mermaid
flowchart TB
    subgraph SHARED["shared/ — the only common code"]
        VAULT["qbo_vault.py + setup_qbo.py<br/>(automation-qbo Keychain blob)"]
        PATHS["paths.py — per-machine output paths<br/>machine.env at REPO ROOT (gitignored)"]
        QAPI["qbo_api.py — auth, retrying GET,<br/>query_all, P&L walkers, PROJ_RE"]
    end

    QBO[("QBO API<br/>(production)")]
    NOTION[("Notion")]
    TEAMS[("Teams")]
    ODRIVE[("OneDrive / SharePoint")]
    NAS[("Synology NAS")]

    VAULT --> QBO

    subgraph INV["invoice-sync/ (was automation-worker/)"]
        INVSYNC["run_invoice_sync.py<br/>(manual alias: sync-ar; Docker v1.1.0)"]
        XLMIRROR["export_invoices_xlsx.py"]
        DOCTOR["doctor.py (diagnostics)"]
    end
    QBO --> INVSYNC
    INVSYNC --> NOTION
    INVSYNC --> XLMIRROR --> ODRIVE
    INVSYNC -- "MFD paid / short-pay cards" --> TEAMS

    subgraph BT["bill-tracker/"]
        BILLS["excel_bill_sync.py<br/>(manual alias: sync-ap)"]
        AUDIT["job_coding_audit + sub_bill_audit<br/>+ item_no_project_audit + duplicate_bill_audit"]
    end
    QBO --> BILLS --> ODRIVE
    QBO --> AUDIT --> ODRIVE

    subgraph SR["statement-reconciler/"]
        RECON["statement_reconciler.py<br/>+ vendor_aliases.json cache"]
    end
    NAS -- "vendor statement PDFs (Inbox)" --> RECON
    QBO --> RECON
    RECON -- "reconciliation xlsx" --> NAS

    subgraph WIP["wip/ — ALL WIP tooling"]
        WIPXL["cp_wip_reader.py / rp_wip_reader.py<br/>(write ONLY Test tabs of WIP - MASTER new.xlsx,<br/>guarded by wip_excel_guard.py)"]
        CLOSE["qbo_close_list.py / qbo_bulk_close.py<br/>(WRITER — gated, ALWAYS excludes MFD)"]
    end
    QBO --> WIPXL --> ODRIVE
    QAPI --> WIPXL
    QBO <--> CLOSE

    subgraph PNL["project-pnl/"]
        PNLX["project_pnl_export.py<br/>(QBO helpers from shared/qbo_api)"]
    end
    QBO --> PNLX --> ODRIVE
    QAPI --> PNLX

    subgraph FIN["Finance tools"]
        LOAN["debt-schedule/loan_sync.py<br/>→ equipment debt xlsx (beside it)"]
        HEALTH["health-dashboard/qbo_health.py → local xlsx"]
        EXPORT["qbo-export/qbo_export.py → OneDrive inbox"]
    end
    QBO --> LOAN
    QBO --> HEALTH
    QBO --> EXPORT --> ODRIVE

    subgraph OO["one-offs/"]
        RECODE["qbo_recode_review.py<br/>(WRITER — audit-gated)"]
    end
    QBO <--> RECODE

    subgraph SYN["synology/"]
        TREE["file-tree audit<br/>(always --exclude sensitive path)"]
    end
    NAS --> TREE
```

**Removed 2026-07-13:** the Field Log subsystem (Bid List → RP/CP Field Logs / Project Plans
sync) — erased by decision, config fields dropped. The retired `wip_sync.py` stub and its
launchd plist are also gone.

---

## Invoice sync detail (the busiest pipeline)

```mermaid
flowchart LR
    QBO[("QBO API")] -- "open invoices" --> SYNC["run_invoice_sync.py"]
    SYNC -- "route by project # prefix" --> MFD[("Notion: MFD Invoice DB")]
    SYNC --> RESCOM[("Notion: Res/Com Invoice DB")]
    SYNC -- "sweep paid" --> MFD
    SYNC -- "sweep paid" --> RESCOM
    QBO -- "CDC changedSince" --> SYNC
    SYNC -- "archive QBO-deleted" --> MFD
    SYNC -- "MFD paid + short-pay Adaptive Cards" --> TEAMS[("Teams webhook")]
    SYNC --> XL["export_invoices_xlsx.py"] -- "full-overwrite mirror<br/>Open_Invoices.xlsx" --> OD[("OneDrive")]
```

---

## Who writes to QBO (the short list)

| Script | What it writes | Gate |
|---|---|---|
| `one-offs/qbo_recode_review.py --apply` | line Customer:Project + Class on job-cost lines | xlsx audit, `Approved=Y` rows only, then `--commit` |
| `wip/qbo_bulk_close.py` | closes customers/projects | `CONFIRM=Y`; always excludes MFD (manual close) |

Everything else is read-only against QBO. All other "writes" land in Excel, Notion, or Teams.

---

## Machine notes

- Output paths (OneDrive folders, `~/Documents/CompanyHealth/`) resolve through
  **`shared/paths.py`**: process env > `machine.env` (REPO ROOT, gitignored, per-machine) >
  owner's original defaults. New machines copy `machine.env.example` → `machine.env` and point
  the two roots (`ACB_ONEDRIVE_BASE`, `ACB_COMPANYHEALTH_DIR`) at their own folders. Never
  hardcode a machine path in a script — add a key through `shared/paths.py` instead.
  Run **`python3 shared/paths.py`** on a new machine: it prints where each path resolved from
  and flags any target that is missing or not writable before you run a real script.
- **Migration note (old clones):** after pulling the restructure, untracked files linger in the
  old `automation-worker/` folder. Move `.env` and `state/` into `invoice-sync/`, then delete
  the leftover folder. Until then, `invoice-sync/config.py` falls back to the legacy location
  automatically (the CDC watermark is never lost).
- Logs stay at `~/Library/Logs/Proficient/automation-worker/` (historical name kept on purpose)
  and always outside the repo. The Keychain service `proficient-automation-worker` also keeps
  its name — renaming it would orphan stored secrets.
- Each developer has their own Intuit app + own Keychain vault. One app connection = one machine.
