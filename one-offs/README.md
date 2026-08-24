# one-offs/ — occasional & not-yet-developed tools

The explicit home for scripts that don't (yet) earn their own tool folder:
occasional audits, experiments, and tools whose driver isn't built yet.

**The rules:**

1. A script graduates OUT of here by earning its own folder — it never
   graduates to the repo root. The root holds no loose Python, ever.
2. Scripts here follow the same import pattern as everything else
   (`sys.path.insert` repo root → `from shared import …`). No tool-folder
   imports.
3. If a one-off starts being run on a schedule or by more than one person,
   that's the signal to promote it.

**Current residents:**

| Script | Status |
|---|---|
| `qbo_recode_review.py` | Audit-gated job-cost recoder (export → the user audits → apply). `get_auth()` is still an env-var stub — wire to `shared.qbo_vault` before real use. |
| `rp_wip_update.py` | **Updates the OWNER'S live RP WIP** in OneDrive in place (never regenerates). Refreshes SCHEDULE/billed/costs/GP% only; his colour marks (orange=verify, green=verified, red=changed) are never overwritten. Dry run → Downloads; `--commit` writes OneDrive. |
| `jobtread_schedule_writer.py` | **Writes to JobTread**: daily schedule → dated tasks using the estimator's own stage names. Upsert by (job, task); never deletes. Dry run by default. |
| `rp_wip_simple.py` | Builds the stripped-down RP WIP (`Test - RP` tab): contract/ETC/billed/costs/GP% + flag sections. ETC from cost-code ROWS (never the pier subtotal cell); FW never in a non-FTW ETC; no `file://` hyperlinks (they beachball Excel). |
| `jobtread_migration_setup.py` | Read-only. "What still needs adding to JobTread": active schedule jobs with no JobTread job + Notion Bid List RP bids still out (not won/lost) that aren't in JobTread. → `~/Downloads/JobTread Migration Setup.xlsx`. |
| `jobtread_bloat_report.py` | Read-only (QBO + JobTread + schedule). Open JobTread jobs that are finished in real life — paid & idle, or stale no-QBO shells — with the evidence (billed, AR balance, last invoice, days idle). → `~/Downloads/JobTread Bloat - Close Candidates.xlsx`. |
| `jobtread_close_jobs.py` | **Writes to JobTread**, audit-gated. `--export` → APPROVE workbook (`CLOSE? Y/N`), `--apply` dry-runs, `--apply --commit` closes by setting `closedOn`. `--reopen` is a true undo (clears `closedOn`); never deletes; MFD excluded by default; changes logged to `~/Library/Logs/Proficient/`. |
| `legacy_job_cost_pull.py` | Read-only QBO. Costs + billing for an OLDER job whose lines were never consistently project-coded. Line-item attribution in three ordered rules (project customer → line text → bill memo), with a hard guard: a memo naming more than one job is skipped, never split. Invoices are pulled from the project AND parent customer. `--expect` verifies a run against a known-good fixture kept outside the repo. |
| `cable_calculator.py` | Read-only. The RP takeoff's PT-cable engine (hidden `'0'` sheet) mapped cell-for-cell and validated on 64 takeoffs. Verification harness for the JobTread migration — **not** an estimator tool. |
| `loans_to_subs_audit.py` | Audit-gated reclass of the `Loans to Sub-Contractors` **parent** account into per-sub sub-accounts. Default run = read-only export (GL parent lines + suggested sub + `Confirm Sub-Account` dropdown + locked Entity/Txn Id/Line Id/SyncToken columns). `--apply` reads the confirmed workbook and moves each line's `AccountRef` to the sub; dry-run by default, writes only on `--commit`. Guards: exact sub name→id, stale SyncToken, closed-period, line-still-on-parent. Auths via `shared.qbo_api.load_credentials()` (one Touch ID). |
