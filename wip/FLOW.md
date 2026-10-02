# wip/ - how the WIP update works

Last changed: 10/02/2026 - the CP reader falls back to the signed pay app (shared.draws.read_pay_app, .xls included) when a job has no .xlsx draw workbook; CP831's Draw #2 change order (33,361) now reaches the WIP Review. Earlier: 09/30/2026 - the **QuickBooks trust gate**: before any costs / billed number is
overwritten, `shared/qbo_trust` checks the job is booked cleanly in QuickBooks; a job that is not
is HELD (the tab keeps its number, a new job is not added) until the reason is fixed.

Update this chart - and the line above - in the same commit as any change to `wip/`
(`.github/flow_guard.sh` fails the push otherwise).

```mermaid
flowchart TD
    classDef src fill:#eef4ff,stroke:#5b7fc7,color:#111
    classDef tool fill:#f4f4f4,stroke:#777,color:#111
    classDef gate fill:#fff3e0,stroke:#d9822b,color:#111
    classDef out fill:#eaf7ea,stroke:#3c8d40,color:#111
    classDef stop fill:#fdecec,stroke:#c0392b,color:#111

    SRC[("1 · Documents<br/>CP draws + takeoffs · RP WIP file + takeoffs · WIP Master sheet (MFD)")]:::src
    READ["2 · The readers<br/>cp_wip_reader · rp_wip_reader · master_wip_test · mfd_wip_test<br/>contract, COs, ETC per job"]:::tool
    PULL["3 · QuickBooks numbers<br/>build_project_customer_map picks ONE customer per project #<br/>costs + billed + retainage from that customer's report"]:::tool
    SRC --> READ --> PULL

    subgraph GATE["4 · QuickBooks trust gate - shared/qbo_trust.assess"]
        G0{"QuickBooks copy readable<br/>and under 12 h old?"}:::gate
        G1{"Project # on two active customers<br/>with money on the skipped one?"}:::gate
        G2{"Its customer hidden by a<br/>name typo - RP7340 -FTW?"}:::gate
        G3{"A cost naming the job<br/>with no project on the line?"}:::gate
        G4{"Flatwork FW cost on the slab<br/>while RP####-FTW exists?"}:::gate
        G0 -->|yes| G1
        G1 -->|no| G2
        G2 -->|no| G3
        G3 -->|no| G4
    end
    PULL --> G0
    G0 -->|no| ALL["EVERY job held"]:::stop
    G1 -->|yes| HB["HELD - costs + billed"]:::stop
    G2 -->|yes| HB
    G3 -->|"$1,000 or more"| HC["HELD - costs"]:::stop
    G4 -->|"$1,000 or more"| HC
    G3 -->|"under $1,000"| CK["CHECK note - still updates"]:::gate
    G4 -->|"under $1,000"| CK
    G4 -->|no| OK["Trusted"]:::out

    DIFF["5 · --emit-review: wip_review_common.diff_rows<br/>carry · reversed · up-only BLOCKED · HELD"]:::tool
    ALL & HB & HC & CK & OK --> DIFF
    DIFF --> REV["6 · Ledger > WIP Review<br/>Passed · Check · Blocked · Held<br/>the owner approves"]:::tool
    REV --> APPLY["7 · --apply-review: apply_decisions<br/>the tab wins unless approved · up-only · hold_untrusted"]:::tool
    APPLY --> WRITER["8 · wip_writer.write_test_cp - THE one door<br/>trust gate runs AGAIN here, then wip_excel_guard + xlsx_verify"]:::tool
    READ -.->|"a direct command-line run skips 5-7, never 8"| WRITER
    WRITER --> TABS[("WIP - MASTER.xlsx<br/>Test - CP · Test - RP · Test-Master")]:::out
    READ -.->|"mfd_wip_test: held fields left as they are"| MFDT[("WIP - MFD tab<br/>the live MFD entry tab")]:::out
```

## What each result means

| Result | What happens to the number | How it clears |
|---|---|---|
| **Passed** | Goes up: written when approved | - |
| **Check** | Went down / blank (retainage), or a finding under $1,000: the owner's call; the note names the documents | Fix in QuickBooks when convenient |
| **Blocked** | Costs / billed went DOWN: never written, the tab keeps its number | Find the voided / deleted / re-coded line in QuickBooks |
| **Held** | QuickBooks not trusted for this job: never written, the tab keeps its number; a job not yet on the tab is not added | Fix the reason in QuickBooks (make the duplicate customer inactive, rename the typo, put the project on the line, move FW to the -FTW job), refresh, compute again |

The gate fails **closed**: if the QuickBooks copy cannot be read, or is older than 12 hours, every
job is held and nothing is overwritten. The ledger's **Company > Duplicate customers** page lists the
same duplicate / typo findings the gate holds on (one shared function, `qbo_trust.duplicate_groups`).
