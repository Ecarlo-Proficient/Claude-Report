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
| `qbo_recode_review.py` | Audit-gated job-cost recoder (export → Ted audits → apply). `get_auth()` is still an env-var stub — wire to `shared.qbo_vault` before real use. |
