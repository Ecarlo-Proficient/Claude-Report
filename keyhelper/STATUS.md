# keyhelper/ - STATUS

Progression record. Update in the SAME commit as any change to this tool (repo rule).

## DONE / FINALIZED
- 2026-09-29 · **The last keys move into the library.** The Notion token and the Teams paid-notice webhook
  (older `keyring` store `proficient-automation-worker`) become `NOTION_SECRET` / `TEAMS_WEBHOOK_MFD_PAID`
  (+ `TEAMS_WEBHOOK_ALERTS`, never stored). `shared/qbo_vault.get_secret(NAME)` is the one lookup (env var ->
  Key Helper -> the old item only until moved); `shared/notion_client`, invoice-sync `config.py` /
  `setup_keychain.py` and `one-offs/jobtread_migration_setup.py` use it. `migrate_keys.py` (dry run by
  default) copies, verifies through the helper, then removes the old item. `install.sh --register` updates
  the registration without a rebuild. Tests: lookup order + a guard that only the vault and the move script
  read the old store.
- 2026-09-29 · **Cutover done** with the owner at the Mac: `install.sh --adopt` - the Keychain item now trusts
  exactly one program (Key Helper, pinned to its build); no temp copy left; marker written. Proven after:
  `setup_qbo --test` AUTH OK through the helper, the ledger restarted (request gate still 403 for a foreign
  Host), a mirror refresh, the private workspace's two profiles by pass and this company read-only by GET -
  every one logged, no password prompt after the one unlock.
- 2026-09-29 · **Key Helper built** (security review: the key library was silently readable by any program
  run as the owner). Swift menu-bar app: Unix socket in a 0700 folder; answers only the owner's user, the
  pinned Python, and scripts inside a registered workspace; one Touch ID / password per session, keys in
  memory only, wiped on lock / sleep / logout / Lock now / 8 h; renews QuickBooks itself and hands out
  one-hour passes; read-only profiles get the GET run for them, never a pass; every request logged.
  Tools switched: `shared/qbo_api.get_pass()` is the ONE login (7 tools' own refresh-token copies retired:
  qbo-export, statement-reconciler, bill-tracker, health-dashboard, qbo_close_list, qbo_bulk_close,
  loan_sync - and invoice-sync's); `shared/qbo_vault` asks the helper once adopted; `setup_qbo --test`
  uses the shared login and no longer prints the company id. `tests/test_key_broker.py` (6, incl. a guard
  that fails any file other than qbo_api calling Intuit's token endpoint). Live-probed on this Mac without
  unlocking: status answered; unknown profile, master key, another workspace's key, a script outside the
  workspace, Apple's Python - all refused and logged. The private workspace registered: its two profiles
  full, this company read-only (a pass for it, its refresh token, changing it - all refused).

## IN PROGRESS

## TO DO
- The Synology move: the broker role relocates to the office server; this helper keeps the Mac's device
  key + the server unlock key (plan in the vault).

## OPEN ISSUES
- Ad-hoc signing pins the build: every rebuild asks for the macOS dialog once (self-heals). A stable
  signing identity would remove that.
