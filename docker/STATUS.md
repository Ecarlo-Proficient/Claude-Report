# docker/ - STATUS

Progression record for the office server package. Update in the SAME commit as any change to this folder (repo rule).

## DONE / FINALIZED
- 2026-10-02 · First run on the Synology: the mirror failed "MIRROR_KEY is not in the Keychain library" - on Linux qbo_vault read only KNOWN_KEYS from the env. It now also reads MIRROR_KEY (`LINUX_EXTRA_KEYS`); tests/test_qbo_vault_linux.py pins it.
- 2026-10-02 · compose.yml APP_UID 1026 -> 1027: the office Synology's svc-automation user is uid 1027 (gid 100 users). The owned folders (data is 700) need the container to run as that uid.
- 2026-10-01 · First build + smoke run (test mode, dummy keys). Two fixes: the Dockerfile reuses GID 100 (Synology's "users" group already exists in Debian - `groupadd` failed the build); `qbo_vault.has_credentials()` on Linux checks only the four QBO keys (it required JT_GRANT_KEY too, so every QBO call on the server failed with "No QBO credentials in Keychain"). The image builds, all five jobs import, the scheduler starts, status.json + alerts work.
- 2026-09-30 · The runbook is plain setup instructions (owner: the developer is capable - no approval stops); the two hard rules stay (the server's own Intuit app; keys typed on the Synology only).
- 2026-09-30 · **The office server package** (replaces the retired invoice-sync-only container). `scheduler.py` (mirror
  every 3 min, AP then AR every 15 min, nightly count check, status.json + Teams alerts, test vs live, the writer file
  fails closed), `Dockerfile` (base pinned by digest, packages by sha256, a read-only allow-list), `.dockerignore`
  (only the allow-list gets in), `compose.yml` (Synology mounts, keys in an encrypted shared folder, no ports),
  `server.env.example` / `secrets.env.example`, `push_registers.sh` (Mac -> server rule files, one way), `CHECKUP.md` (step 0, look only), the runbook
  in `README.md`. `tests/test_office_server_image.py` fails the build if a QuickBooks writer is copied in (proven
  against pay_bills / reapply_check). Not built or run yet - there is no Docker on the owner's Mac; the developer
  builds on the Synology.

## IN PROGRESS
- 2026-10-05 · The test week is running on the office Synology (test mode). Every job ok since the first full mirror
  (mirror every 3 min, AP + AR + AR export every 15 min, reconcile counts QBO = mirror). First side-by-side vs the
  Mac's live trackers: Invoice Tracker - every open invoice and balance matches (layout differs only because the
  server runs the newer aging code); Bill Tracker - Inventory / Liens identical, Bills differ on payments (side note below).

## TO DO (before live)
- Notion: only update an invoice when a field changed (today every run rewrites every open invoice).
- DONE 10/06 - The Mac honors the writer file: sync-all stands down on AP/AR while it says "server" (push_registers.sh still run by hand after a rule change).
- Bill Tracker Inputs split: Lien / Notes / coding Status move to `Bill Tracker Inputs.xlsx`, the tracker read-only.
- The nightly cross-check between the server's mirror and the Mac's.
- The ledger shows the server's last good run (reads `Accounting/_automation/server-status.json`).
- Live runs `main` only - a release PR before the switch.

## OPEN ISSUES
- Side note (2026-10-05) - **the server's Bill Tracker can't be opened by anyone.** `excel_bill_sync.py` chmods the
  output 600; on the Mac that is the owner's own account, on the server it is svc-automation, so no person (or Excel
  over the share) can read `_server-test/Bill Tracker.xlsx` - a sudo copy was needed to compare it. Live would lock
  the real one the same way. Fix before live: readable by the share's users on the server (group-read, or skip the
  chmod when `ACB_SERVER_MODE` is set). The Invoice Tracker is not affected. FIXED 10/06: excel_bill_sync skips the
  chmod when ACB_SERVER_MODE is set (the share's permissions decide).
- Side note (2026-10-05) - **payments: 126 bill lines are paid on the Mac, Unpaid/Partial on the server.** Same bills,
  same lines; only Open Bal / Pay Status / Pay Ref-Date-Method / Lien differ, always paid -> unpaid. UPDATE 10/06: check
  #25510 is whole again on the Mac's mirror (115 bills, last changed in QBO 10/06 - the ledger's Clear copy re-apply);
  most likely QBO stripped it between the two 10/05 runs. Re-compare the two trackers once before going live. All of them trace
  to three check payments on the Mac's side - one of them (check #25510, April 2026) pays 109 bills by itself. Either
  the server's mirror is missing those BillPayments (suspect: very large payments) or QBO changed them between the
  two runs (Mac 08:09, server 11:50). Not yet settled: open check #25510 in QBO, or compare a same-hour pair of files.
  Probably the same cause: 6 extra bills only on the server (Unpaid) and 3 Bill List rows "approved" -> "check QBO".
