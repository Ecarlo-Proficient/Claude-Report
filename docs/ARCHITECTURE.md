# Automation Suite — System Map

> **Maintenance rule (binding):** any commit that adds/removes a script, changes what a script
> reads or writes, or rewires a data flow MUST update this diagram in the same commit.
> Claude sessions: check this file whenever `git status` shows script changes at session end.
> GitHub renders the Mermaid blocks below natively — view this file on github.com to see the diagrams.

Last updated: 2026-07-09 (added paths.py per-machine config layer)

---

## Big picture

Every QBO script authenticates through one shared Keychain vault (single Touch ID per run).
QBO is production-only and the source of truth for costs/billed. Writers are gated
(`--commit` / `CONFIRM=Y`); everything else is read-only against QBO.

```mermaid
flowchart TB
    subgraph AUTH["Auth (macOS Keychain)"]
        VAULT["automation-qbo blob<br/>qbo_vault.py + setup_qbo.py"]
        NVAULT["automation-notion blob<br/>setup_keychain.py"]
        TVAULT["automation-teams blob<br/>setup_keychain.py"]
    end

    subgraph CFG["Per-machine config (repo root)"]
        PATHS["paths.py — path lookups for ALL output scripts<br/>machine.env (gitignored) overrides;<br/>defaults = owner's original paths"]
    end

    QBO[("QBO API<br/>(production)")]
    NOTION[("Notion")]
    TEAMS[("Teams")]
    ODRIVE[("OneDrive / SharePoint")]
    NAS[("Synology NAS")]

    VAULT --> QBO

    subgraph ROOT["Root exports"]
        EXPORT["qbo_export.py"]
        RECODE["qbo_recode_review.py<br/>(WRITER — audit-gated)"]
    end
    QBO --> EXPORT --> ODRIVE
    QBO <--> RECODE

    subgraph AW["automation-worker/"]
        INVSYNC["run_invoice_sync.py<br/>(manual alias: sync-ar)"]
        XLMIRROR["export_invoices_xlsx.py"]
        FIELDLOG["field_log_sync.py<br/>(checkbox-gated, UPDATE-only)"]
        PLANS["project_plans_sync.py"]
        DOCTOR["doctor.py (diagnostics)"]
    end
    QBO --> INVSYNC
    INVSYNC --> NOTION
    INVSYNC --> XLMIRROR --> ODRIVE
    INVSYNC -- "MFD paid / short-pay cards" --> TEAMS
    NOTION --> FIELDLOG --> NOTION
    NOTION --> PLANS --> NOTION
    NVAULT --> INVSYNC
    TVAULT --> INVSYNC

    subgraph BT["bill-tracker/"]
        BILLS["qbo_bill_tracker.py<br/>(launchd Mon–Fri 15:00, alias: sync-ap)"]
        AUDIT["job_coding_audit.py + sub_bill_audit"]
    end
    QBO --> BILLS --> ODRIVE
    QBO --> AUDIT --> ODRIVE

    subgraph SR["statement-reconciler/"]
        RECON["statement_reconciler.py<br/>+ vendor_aliases.json cache"]
    end
    NAS -- "vendor statement PDFs (Inbox)" --> RECON
    QBO --> RECON
    RECON -- "reconciliation xlsx" --> NAS

    subgraph WIP["wip/"]
        WIPSYNC["build_wip_master.py / wip_sync.py<br/>(writes ONLY the Test tab)"]
        CLOSE["qbo_close_list.py / qbo_bulk_close.py<br/>(WRITER — gated, ALWAYS excludes MFD)"]
    end
    QBO --> WIPSYNC --> ODRIVE
    WIPSYNC -- "$25K job-borrow alert" --> TEAMS
    QBO <--> CLOSE

    subgraph FIN["Finance one-shots"]
        LOAN["loan_sync.py → equipment debt schedule xlsx"]
        HEALTH["health-dashboard/qbo_health.py → local xlsx"]
        PNL["project-pnl/project_pnl_export.py → per-project P&L xlsx"]
    end
    QBO --> LOAN
    QBO --> HEALTH
    QBO --> PNL --> ODRIVE

    subgraph SYN["synology/"]
        TREE["file-tree audit<br/>(always --exclude sensitive path)"]
    end
    NAS --> TREE
```

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
| `qbo_recode_review.py --apply` | line Customer:Project + Class on job-cost lines | xlsx audit, `Approved=Y` rows only, then `--commit` |
| `wip/qbo_bulk_close.py` | closes customers/projects | `CONFIRM=Y`; always excludes MFD (manual close) |

Everything else is read-only against QBO. All other "writes" land in Excel, Notion, or Teams.

---

## Machine notes

- Output paths (OneDrive folders, `~/Documents/CompanyHealth/`) resolve through **`paths.py`**
  (repo root): process env > `machine.env` (gitignored, per-machine) > owner's original defaults.
  New machines copy `machine.env.example` → `machine.env` and point the two roots
  (`ACB_ONEDRIVE_BASE`, `ACB_COMPANYHEALTH_DIR`) at their own folders. Never hardcode a
  machine path in a script again — add a key through `paths.py` instead.
  Run **`python3 paths.py`** to confirm config on a new machine: it prints where each
  path resolved from and flags any target that is missing or not writable before you
  run a real script.
- Logs always go to `~/Library/Logs/` (external log dir per CLAUDE.md), never inside the repo.
- Each developer has their own Intuit app + own Keychain vault. One app connection = one machine.
