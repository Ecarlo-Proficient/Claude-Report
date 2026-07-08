# Power Automate — Tier Sync Flow Sketch

Goal: keep the trimmed lower-tier databases (Common Bid List, AR Project Summary, etc.) in sync with the source Bid List, using Project # as the match key. Power Automate handles what Notion automations cannot — **find-by-value and update** (upsert).

This doc sketches the flow shape and the exact field mapping per flow. Builders can use it as a spec.

---

## Why Power Automate, not Notion automations

| Need | Notion native | Power Automate |
|---|---|---|
| Create new row when source row is created | ✅ | ✅ |
| Update existing row when source row is edited (match by Project #) | ❌ — no find-by-value action | ✅ — HTTP query → update page |
| Backfill historical rows at setup | ❌ — automations only fire on future events | ⚠️ one-time CSV export + import still required on first setup, then flow handles forward |
| Run on your M365 stack | n/a | ✅ — already in Ted's stack |

Reference: `reference_notion_constraints.md`.

---

## Prerequisites

Set up once per flow:

1. **Notion internal integration** — create at notion.so/profile/integrations. Grab the Internal Integration Secret. Share both the source DB and the trimmed DB with this integration (`...` on DB → Connections → Add the integration).
2. **Project # is unique and stable on both sides.** Title of the lower-tier DB = Project # string. Title of the source Bid List = same Project # string. This is how the flow matches.
3. **Initial backfill done via CSV.** Export Bid List → open in Excel → trim columns to what the lower tier allows → import to the lower-tier DB. Do this before turning the flow on so the flow only handles deltas.
4. **Connection reference in Power Automate** — set up an HTTP action pattern or a custom connector so auth + `Notion-Version` + `Authorization: Bearer …` are reusable.

---

## Flow pattern — shared across all tier-sync flows

Every tier-sync flow has the same shape. Only the source DB, target DB, and field mapping change.

```
TRIGGER: Scheduled every 5 min (or: webhook if available)
  │
  ▼
STEP 1 — Query source DB (Bid List) for rows changed in last N minutes
  HTTP POST /v1/databases/{BID_LIST_DS_ID}/query
  body: { "filter": { "property": "Last edited time",
                      "last_edited_time": { "on_or_after": "<now-5min ISO>" }}}
  │
  ▼
STEP 2 — For each changed row in source:
  │
  ├─ STEP 2a — Query target DB for row with Title = source Project #
  │    HTTP POST /v1/databases/{TARGET_DS_ID}/query
  │    body: { "filter": { "property": "Project #",
  │                        "title": { "equals": "<source Project #>" }}}
  │
  ├─ STEP 2b — Condition: results.length > 0 ?
  │    │
  │    ├─ YES (row exists)
  │    │    HTTP PATCH /v1/pages/{target_page_id}
  │    │    body: { "properties": { ...mapped fields... }}
  │    │
  │    └─ NO (row doesn't exist)
  │         HTTP POST /v1/pages
  │         body: { "parent": { "database_id": "{TARGET_DS_ID}" },
  │                 "properties": { ...mapped fields... }}
  │
  ▼
STEP 3 — Log the run (success count, error count) to a log Excel or Teams channel
```

Notes:
- 5-min interval is a reasonable default. Tighten to 1 min if Ted wants near-real-time; loosen to 15 min to reduce API calls.
- Notion API rate limit ~3 req/sec. Each changed row costs 2-3 requests. At 5-min cadence, a shop like Proficient won't hit the ceiling.
- `Last edited time` trigger guarantees you don't miss an edit, but it also re-processes rows edited multiple times in the window. The upsert is idempotent, so re-processing is safe.
- If a relation (e.g. Customer) is edited in the source, the target can't hold the relation (tier boundary) — flatten to text.

---

## Flow #1 — Bid List → Common Bid List

**Purpose:** give the Common clerks (AP Clerk, PO Clerk, Subcontractor Clerk) a read-only, trimmed view of every job with only basic identifying info. No sales, no bid amount, no plans, no financials.

**Target DB:** Common Bid List (already exists — data source ID: TBD, Ted to provide).

**Field mapping (source Bid List → Common Bid List):**

| Source (Bid List) | Target (Common Bid List) | Type on target | Notes |
|---|---|---|---|
| Project # (title) | Project # (title) | TITLE | Match key. |
| Job Name | Job Name | RICH_TEXT | |
| Address | Address | RICH_TEXT | |
| City | City | RICH_TEXT | |
| Customer (relation → Customer List) | Customer | RICH_TEXT | **Flatten** — read the related page's Name and write as plain text on target. |
| Division | Division | SELECT | Residential / Commercial. |
| Active Status | Active Status | SELECT | |

**Fields NOT to map (tier gate):**
- Bid Amount (sales data)
- Lead Status (sales lifecycle)
- Estimator (internal role)
- Plans / Notes (may contain sensitive info)
- RP Stages / CP Stages relations (cross-DB access would leak tier)
- Superintendent (field role info not needed for common clerks)

**Trigger:** `Last edited time` within 5 min, OR on create.

**Failure behavior:** log to Teams channel `#power-automate-log`. Retry 3x on 429 rate limit.

---

## Flow #2 — Bid List → AR Project Summary

**Purpose:** give the AR team (Ana AP/AR, MRojas, ERivera Collections) a project roster they can attach Open Invoices against. They need enough to identify the job and the customer, plus the billed / collected state — but not the bid amount until after the sale.

**Target DB:** AR Project Summary (new — create it first using the schema below, then set up this flow).

**Target DB schema to create (if it doesn't exist yet):**

| Property | Type | Source of truth |
|---|---|---|
| Project # | TITLE | Bid List Project # |
| Job Name | RICH_TEXT | Bid List |
| Address | RICH_TEXT | Bid List |
| Customer | RICH_TEXT | Bid List (flattened from relation) |
| Division | SELECT | Bid List |
| Active Status | SELECT | Bid List |
| **Invoiced Amount** | NUMBER (dollar) | AR-side data / QBO sync |
| **Collected Amount** | NUMBER (dollar) | AR-side data / QBO sync |
| **Open Invoices** | RELATION (Open Invoices DB) | AR-side |
| **AR Owner** | PEOPLE | AR team sets |
| **Collection Status** | SELECT | Current / Past Due / Collections |

**Field mapping (source Bid List → AR Project Summary):**

| Source | Target | Type | Notes |
|---|---|---|---|
| Project # | Project # | TITLE | Match key. |
| Job Name | Job Name | RICH_TEXT | |
| Address | Address | RICH_TEXT | |
| Customer (relation) | Customer | RICH_TEXT | Flatten. |
| Division | Division | SELECT | |
| Active Status | Active Status | SELECT | Lets AR see when a project moves into completion and full billing is due. |

**Fields NOT to map:**
- Bid Amount — AR only needs invoiced amount, not estimate.
- Plans / Notes
- RP Stages / CP Stages
- Superintendent
- Lead Status (AR doesn't need pre-sale lifecycle)

**AR-side fields are written by QBO sync or manually by AR.** The Power Automate flow does NOT touch them — it only manages the fields copied from Bid List. That way AR's work is never overwritten by a sales-side edit.

**Trigger:** only when Active Status transitions to "Sold" or later (saves API calls on pre-sale churn). Use a filter on the query: `Active Status is NOT Lead AND Active Status is NOT Bidding`.

---

## Upsert pseudocode (reference for the builder)

```
# Get all rows edited in last 5 min
changed = notion.query(BID_LIST_DS_ID,
    filter={"timestamp": "last_edited_time",
            "last_edited_time": {"on_or_after": now - 5min}})

for row in changed.results:
    project_num = title_plain_text(row.properties["Project #"])

    # Build mapped payload (flatten relations, etc.)
    mapped = {
        "Project #":     {"title":      [{"text": {"content": project_num}}]},
        "Job Name":      {"rich_text":  [{"text": {"content": row.props["Job Name"]}}]},
        "Address":       {"rich_text":  [{"text": {"content": row.props["Address"]}}]},
        "City":          {"rich_text":  [{"text": {"content": row.props["City"]}}]},
        "Customer":      {"rich_text":  [{"text": {"content": flatten_relation(row.props["Customer"])}}]},
        "Division":      {"select":     {"name": row.props["Division"].name}},
        "Active Status": {"select":     {"name": row.props["Active Status"].name}},
    }

    # Find target row by Project # (title match)
    hits = notion.query(TARGET_DS_ID,
        filter={"property": "Project #",
                "title": {"equals": project_num}})

    if hits.results:
        # Update existing
        notion.pages.update(hits.results[0].id, properties=mapped)
    else:
        # Create new
        notion.pages.create(parent={"database_id": TARGET_DS_ID}, properties=mapped)
```

In Power Automate this is expressed as:
- **Scheduled recurrence** → initial trigger
- **HTTP** action → query source
- **Apply to each** on `body('HTTP')?['results']`
- **HTTP** action → query target
- **Condition** on `length(body('HTTP_target_query')?['results'])` > 0
- **HTTP** action (YES branch) → PATCH
- **HTTP** action (NO branch) → POST

---

## Rollout order

1. **Create AR Project Summary DB** with the schema above (manual, one-time).
2. **Populate Common Bid List and AR Project Summary via CSV backfill** from current Bid List. Trim sensitive columns before import.
3. **Build Flow #1 (Common Bid List sync)** in Power Automate. Test with a single test Bid List row: edit → wait 5 min → confirm Common Bid List row updated. Confirm sensitive columns are absent.
4. **Build Flow #2 (AR Project Summary sync)**. Same test pattern.
5. **Turn both on.** Monitor the log channel for the first week — expect 0 failures after auth is right.
6. **Grant lower-tier users access only to the trimmed DBs** (not the source Bid List). Revoke any legacy access to Bid List for those users.

---

## Future flows (same pattern)

Same flow shape will serve:
- Bid List → Payroll Project List (if payroll needs project identity without bid info)
- Bid List → Subcontractor Project List
- Field Log summary → a limited shared view (if needed)

Each one is a copy of the template flow with a different target DS ID and a different field mapping table.

---

## What can break & how to catch it

- **Source row's Title (Project #) edited.** Title is the match key. If a super or PM renames the title, the flow will stop matching and create a duplicate on the next edit. **Mitigation:** lock title edits via permission (Bid List should be editable only by admins + estimators), and/or add a second stable identifier and match on that instead.
- **Relation target renamed.** Flattening reads the related page's Name. If the Customer page is renamed, next sync will overwrite with the new name — usually fine, but note in change log.
- **Rate limits** (HTTP 429). Notion allows ~3 req/sec. If Ted ever bulk-edits 50+ rows at once, the flow might 429. **Mitigation:** retry 3x with exponential backoff on 429, already standard in Power Automate HTTP action settings.
- **Auth expires.** Internal integration secrets don't expire, but if the integration is removed from a DB, the flow will 401. **Mitigation:** log 401 separately and alert.

---

## Alternative if Power Automate feels heavy

A Python script on Ted's existing setup running on Windows Task Scheduler can do the same thing. Same pseudocode above, just runs on local cron. No external subscription needed beyond what's already paid. Power Automate is the preferred choice because Ted already has M365 and it gives a visual flow + retry + logging out of the box, but if Power Automate licensing is ever an issue, the script fallback is a one-day port.
