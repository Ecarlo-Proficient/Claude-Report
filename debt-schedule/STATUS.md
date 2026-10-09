# debt-schedule/ - STATUS

Progression record. Update in the SAME commit as any change to this tool (repo rule).

## DONE / FINALIZED
- 2026-09-29 · **QuickBooks login moved to the shared one (security review).** This tool's own copy of the refresh-token exchange is retired; its login is `shared/qbo_api` (`load_credentials` / `get_pass`), which asks Key Helper (`keyhelper/`) once the key library is adopted - the tool never holds the refresh token. The realm is no longer printed on login.
