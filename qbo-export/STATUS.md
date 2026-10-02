# qbo-export/ - STATUS

Progression record. Update in the SAME commit as any change to this tool (repo rule).

## DONE / FINALIZED
- 2026-09-29 · **Formula-injection guard (security review).** Outside text (vendor / QBO / Notion / email) this tool writes goes through `shared/xlsx_guard` - stored as text, never a live formula (every exported memo / description); every workbook save also passes the shared guard, and `xlsx_verify` fails a risky formula.
- 2026-09-29 · **QuickBooks login moved to the shared one (security review).** This tool's own copy of the refresh-token exchange is retired; its login is `shared/qbo_api` (`load_credentials` / `get_pass`), which asks Key Helper (`keyhelper/`) once the key library is adopted - the tool never holds the refresh token.
