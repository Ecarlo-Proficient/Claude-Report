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
- The developer: the setup in README.md (the box, folders, QBO auth, keys, build, test mode).

## TO DO (before live)
- Notion: only update an invoice when a field changed (today every run rewrites every open invoice).
- The Mac honors the writer file: sync-all stands down on AP/AR while it says "server" (and runs push_registers.sh).
- Bill Tracker Inputs split: Lien / Notes / coding Status move to `Bill Tracker Inputs.xlsx`, the tracker read-only.
- The nightly cross-check between the server's mirror and the Mac's.
- The ledger shows the server's last good run (reads `Accounting/_automation/server-status.json`).
- Live runs `main` only - a release PR before the switch.

## OPEN ISSUES
- The image has been built and smoke-run off-box (10/01/2026), not yet on the Synology itself.
