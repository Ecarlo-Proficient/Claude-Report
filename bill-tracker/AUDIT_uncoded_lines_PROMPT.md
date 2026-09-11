# Task: add an "uncoded expense line" audit to the bill tracker

## Goal
Make sure **no expense line slips through without a Customer/Project**. Today the
QBO Audit sheet silently misses lines that have a **blank** Customer/Project,
which is exactly how job costs get left off a project. Add a new audit section
that catches every uncoded line and triages it so legitimate overhead doesn't
drown out real misses.

## File & functions
All changes are in **`bill-tracker/excel_bill_sync.py`**. Do **not** touch
`qbo_bill_tracker.py` or `bill_rows.py` — the row data already has everything
needed.

Relevant existing pieces (read them first):
- `build_audit_sheet(ws, rows)` — renders the "QBO Audit" sheet. Currently two
  sections: (1) stale NOT APPROVED bills, (2) "Data entry issues" from
  `_audit_row_checks`.
- `_audit_row_checks(r)` — the per-line checks. **This is where the gap is:**
  checks #1 and #2 are gated on `if expected_div ...` and #3 on
  `if project_num ...`, so a line with **no** project produces an empty issue
  list and is never flagged.
- `_normalize_class(class_name)` → returns `"RP"|"CP"|"MFD"` or `None`.
- `_audit_section_banner(ws, row_idx, text, fill, n_cols)`, `_qbo_link(bill_id)`,
  `_format_data_cell(c, kind)`, `AUDIT_HEADERS`, `HEADER_FONT/FILL`.
- The audit runs on `all_rows` (open **and** paid, line-level) — keep that
  source so paid bills with uncoded lines also surface.

Row dict fields available per line (from `bill_rows.build_rows`):
`bill_id, bill_doc, vendor, bill_date, bill_type ("COGS"|"Other"), account,
line_amount, line_desc, customer_name, project_num, division, class_name,
auto_status, approved`.

## What "uncoded" means
A line is **uncoded** when it has no project:
`not (r.get("project_num") or "").strip()`.

Distinguish two reasons in the message:
- `customer_name` blank → **"No Customer/Project"**
- `customer_name` set but no project parsed (coded to a parent GC only) →
  **"Customer set, no project # (parent only: <customer_name>)"**

## Two-tier triage
Add a new **Section 3** to `build_audit_sheet`, rendered after the existing two
sections, with two sub-banners (mirror the aging-bucket sub-banner pattern):

### Tier 1 — "Likely job cost — MISSING project"
An uncoded line lands here if **any** of these high-confidence signals is true:
1. `r.get("bill_type") == "COGS"` (Item/product line — carries a cost code), **or**
2. `_normalize_class(r.get("class_name"))` is not `None` (class is set to a
   division, so it's clearly job work), **or**
3. `r.get("line_desc")` contains a project code —
   `re.search(r"\b(MFD|CP|RP)\d+(?:-FTW)?\b", line_desc, re.IGNORECASE)`.

When signal #3 fires, append the detected code to the message, e.g.
`"desc says MFD281"`, so the clerk knows where it belongs.

### Tier 2 — "Review — uncoded, may be overhead"
Every other uncoded line (no job-cost signal). These are probably legitimate
overhead (utilities, office, software, Supplies bucket) but are shown so nothing
is left out.

## Implementation
1. Add a classifier:
   ```python
   def _uncoded_tier(r: dict) -> Optional[str]:
       """Return 'job_cost' | 'overhead' for an uncoded line, or None if the
       line already has a project."""
       if (r.get("project_num") or "").strip():
           return None
       is_item   = r.get("bill_type") == "COGS"
       class_div = _normalize_class(r.get("class_name") or "")
       desc      = r.get("line_desc") or ""
       has_code  = re.search(r"\b(MFD|CP|RP)\d+(?:-FTW)?\b", desc, re.IGNORECASE)
       return "job_cost" if (is_item or class_div or has_code) else "overhead"
   ```
2. Add a reason builder that returns the message string for a row (No
   Customer/Project vs parent-only; plus the desc-code hint when present).
3. In `build_audit_sheet`, after Section 2, collect uncoded rows
   (`tier = _uncoded_tier(r)`), split into the two tiers, sort vendor-first then
   by `bill_doc`, and render with the existing `AUDIT_HEADERS` columns:
   - `Project # (parsed)` → `"(none)"`
   - `Division (expected)` → division implied by class if class is set, else
     `"(none)"`
   - `Mismatch` column → the reason string
   - `Open` column → `_qbo_link(r["bill_id"])`
4. Use **distinct section fills** consistent with the file's palette: a strong
   color for the Tier-1 banner (these are the action items) and a lighter shade
   for Tier-2. Tint the Mismatch cell to match each tier (follow how Section 1
   tints amber and Section 2 tints pink).
5. Update the returned count and the summary `print(...)` at the end of `main()`
   so the run log reports uncoded counts, e.g.
   `Audit: N flagged (J job-cost misses, O overhead to review)`.
6. If a single uncoded line could match BOTH an existing `_audit_row_checks`
   issue and the new uncoded check — it can't, because the old checks require a
   project — so no de-dup is needed. Confirm this stays true.

## Constraints (don't break these)
- Keep the existing two sections and `_audit_row_checks` behavior exactly as-is.
- Keep `all_rows` (open + paid) as the audit source.
- Keep the sheet an Excel Table with AutoFilter and the `chmod 600` output
  behavior already in `main()`.
- One row per **line** (line-grain), so a bill with several uncoded lines shows
  each line — that's the point.
- Match existing styling helpers; don't introduce a new formatting style.

## Verification (required before declaring done)
1. `python3 -m py_compile excel_bill_sync.py`.
2. Unit-test `_uncoded_tier` and the reason builder against mock rows:
   - Item line, no project → `job_cost`, "No Customer/Project".
   - Account line, class "Multi Family", no project → `job_cost`.
   - Account line, desc "pour MFD281", no project → `job_cost`, "desc says MFD281".
   - Account line "Utilities", no class, no project → `overhead`.
   - Any line **with** a project → `None` (not flagged).
   - Customer "Greystar" set but no project parsed → flagged, parent-only message.
3. `python3 excel_bill_sync.py --dry-run` (and `--limit` for a smoke run) to
   confirm it builds without writing, and the printed counts look sane.
4. Open the resulting sheet and confirm Tier 1 / Tier 2 banners render and the
   ↗ Open links resolve to the right transactions.

## Tuning note for the user
If Tier 2 (overhead) is noisy because the sync window reaches far back, scope it
to open bills only or to a recent date window — but keep Tier 1 (job-cost
misses) covering everything. Flag this to the user rather than dropping coverage
silently.
