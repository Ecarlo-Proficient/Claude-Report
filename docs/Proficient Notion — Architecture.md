# Proficient Notion — Architecture Plan

> **⚠️ Partially superseded as of 2026-04-28.** This doc describes a 4-tier data model with **Common Bid List** as a Tier 1 lookup database for clerks. That database was deleted 2026-04-28 — the user decided clerks need direct Bid List access (they need bid amounts, lead status, etc. to do their jobs). The remaining tier separation is now just Bid List (everyone) + Field Log (super-restricted view, no financials). The Common Bid List sections below remain as historical context. The Field Log architecture and the Bid List → Field Log sync logic still apply.
>
> **⚠️ Further superseded as of 2026-07-13.** The **Field Log flow was dropped entirely** — the Bid List → Field Log / Project Plans sync code and its templates were removed from the automation repo by decision. Field Log sections below are historical context only.

## Design Journey — why we landed here

**Starting goal:** One system where projects originate in the Bid List and flow cleanly to field, AR, and other teams with no manual duplication and controlled visibility.

**Constraints that drove the architecture:**
- Project # is the single identifier across all systems
- Different teams need different levels of data (tiers)
- Supers must NOT see pricing/sales
- Data must stay in sync automatically
- No manual copying or maintenance

**What we tested and ruled out:**

**Attempt 1: One database + filtered views.**
- Views are not permissions.
- Users can still open pages and see everything.
- No way to hide specific properties per role.
- *Ruled out — does not meet tier/security requirement.*

**Attempt 2: Separate databases, no automation.**
- Correct for permissions.
- But no reliable way to keep them synced natively.
- Notion cannot match rows by Project #, update existing rows, or prevent duplicates on its own.
- *Correct structure, but missing automation layer.*

**Attempt 3: Native Notion automations.**
- Can create pages.
- Cannot "find and update existing row by key."
- No upsert logic.
- *Not sufficient for the system.*

**The real requirement identified:**
- A unique key (Project #)
- Upsert logic: if Project # exists → update, if not → create.
- That is database behavior, not Notion behavior.

**Final conclusion:**

Keep **Notion** for:
- UI
- Team workflows
- Data structure
- Permissions via separate databases

Add **Microsoft Power Automate** for:
- Real-time sync
- Matching by Project #
- Updating correct rows
- Creating new rows when needed

**Architecture that results:**

```
Bid List (source of truth)
    ↓ (Power Automate sync by Project #)
Common Bid List (Tier 1)
Field Logs (Tier 2 — via relation + rollups, not Power Automate)
AR Project Summary (Tier 3)
```

**What Power Automate solves:**
- Eliminates duplicates
- Keeps all tier databases in sync automatically
- Runs continuously (no manual scripts)
- Gives you true "update or create" behavior

**Bottom line:** Didn't outgrow Notion. Needed to pair it with an automation layer. Notion = structure + interface. Power Automate = logic + data integrity. Together they give the system we were trying to build.

**Note on mechanism boundaries:**
- **Button** creates downstream rows in same-teamspace databases via relations (Field Logs from Bid List)
- **Rollups** surface approved upstream data into a lower-tier DB via existing relation (Plans to Field Log)
- **Power Automate** crosses teamspace boundaries where relations would break visibility (Common Bid List, AR Project Summary)

Three mechanisms, each with a clear job, no overlap.

---

## Core principles

**Two things are independent and both matter:**

1. **Project # is the spine for OPERATIONS.** Projects are born on the Bid List and every operational database (field logs, invoices, sub assignments, etc.) relates back to Project #. Project # = Project #. No new codes, no format changes, whatever is assigned on Bid List is the identifier everywhere.

2. **Some teamspaces have NOTHING to do with projects.** Payroll is the clearest example — employee pay, timesheets, PTO. It exists parallel to the project spine, not attached to it. Future teamspaces like HR or internal ops may also be off-spine.

**And on top of that, there are info tiers within project data.** Membership in a teamspace alone doesn't determine what project data a role sees. We enforce tiers of detail:

### Info tiers (project data visibility)

| Tier | What's visible | Who |
|---|---|---|
| **1 — Basic lookup** | Job Name, Project #, Address, City, Customer | All employees, Common Clerks (AP, PO, Sub Clerks). No plans. No financial. No sales. |
| **2 — Field operational** | Tier 1 + phase status, dates, photos, plan file, super assignment, checklist progress | Superintendents, Field Clerk. **No bid amount, no customer pricing, no sales info.** |
| **3 — Sales & financial** | Tier 1 + Tier 2 + bid amount, margins, customer pricing, invoice status, AR aging | Estimators, PMs, AR/AP staff |
| **4 — Full** | Everything | Leadership (the user) |

Tiers are enforced at the **database level** in Notion, not the property level. Notion has no per-property permission. That means each tier's visible data must live in a database (or linked view) that only that tier's teamspaces can access.

### How tiers are enforced

Notion's permission model is **database-level**, not column-level. Even linked database views require access to the source database (and any related databases) to render values. That means "hide the sensitive column in the view" is NOT a security mechanism — any user with access to the source database can toggle the column back on or see it via a different view. **Enterprise plan does not change this** — per-property permissions don't exist at any Notion tier.

The only reliable enforcement mechanisms:

1. **Separate physical databases where the tier boundary exists.** The lower-tier database gets a synced copy of the columns that tier can see — all as text/plain fields, no relations to the source. Users get access to the lower-tier DB only, zero access to the source. Examples: Common Bid List (Tier 1 copy of Bid List basics), AR Project Summary (Tier 3 copy relevant to invoicing).
2. **Relations replaced with copied text values on the lower-tier DB.** If Customer is a relation on Bid List, it becomes a plain-text Customer Name on Common Bid List. No cross-DB link = no cross-DB access requirement.
3. **Rollups for cross-tier info that SHOULD surface.** Example: Plans is a property on Bid List (Tier 3). Supers need plans (Tier 2) — plans rollup from Bid List to Field Log phase rows. Supers see the rollup value only, not the rest of Bid List.
4. **Separate teamspaces per role group** — container for who can even see what.
5. **Per-user filters** (e.g. "only my jobs" for supers) for clutter, not security.

**Data sync between tiers — hard Notion constraint:**

Separate databases need to stay in sync with their source. **Notion's native automations cannot match rows by value (Project # or any property) and update the matching row in another database.** They can only:
- Create new pages (in any database)
- Update the page the automation chain is already operating on

That means: if a Bid List row is **edited** after the initial backfill, Notion automation alone cannot find the corresponding Common Bid List row by Project # and update it. It would either create a duplicate or do nothing.

This is a Notion-wide limitation — not fixed by Enterprise or any plan upgrade.

**Your two real options for each separate-database sync:**

**Option 1 — Notion-only (create-only, no edit propagation):**
- One-time backfill via CSV export/import for existing rows
- Notion automation fires on new Bid List rows going forward → creates matching row in target DB
- **Edits to existing Bid List rows DO NOT propagate.** Accept this, or periodically re-export/overwrite.
- Lowest complexity, no external tooling.

**Option 2 — External tool (Power Automate, Make, or script) with upsert logic:**
- Runs on schedule or on Notion webhook trigger
- Searches target DB for Project # → updates existing row OR creates if missing
- Handles both creates and edits cleanly
- Requires Power Automate setup (already in your Microsoft 365 stack) — recommended for the long term
- Same tool will power the Notion → company Excels sync, so one automation investment covers multiple use cases

**Recommendation:** Start with Option 1 per tier boundary to get the structure live, upgrade to Option 2 once Power Automate is configured for the Excel sync anyway.

---

## Teamspaces

### 1. Common Lookup (Tier 1 — all employees)

**Purpose:** Quick project lookup. Anyone in the company can find a job by name, know where it is and who the customer is. **No plans. No financials. No sales. No field operational detail. Basics only.**

**Users:** All employees — AP Clerk, PO Clerk, Subcontractor Clerk, plus anyone else who needs project lookup without seeing the rest

**Databases:**
- **Common Bid List** — a SEPARATE physical database (not a linked view). Holds ONLY:
  - Job Name (text, copied from Bid List)
  - Project # (text, copied)
  - Address (text, copied)
  - City (text, copied)
  - Customer (text, copied — **not** a relation)

**Why a separate database, not a linked view:** Notion enforces permissions at the database level, not the column level. A linked view of Bid List would require users to have access to Bid List AND any related databases (like Customer List) just to render the value — defeats the tier boundary. By making Common Bid List a standalone database with only text fields, Common Clerks need access to ONE database, zero upstream dependencies. Nothing to click through to.

**How data gets in:**

- **One-time backfill** — export the current Bid List to CSV, trim to the 5 columns (including converting the Customer relation to its text name), import into Common Bid List. Required because Notion automations will not backfill existing rows.
- **Ongoing sync — pick one:**
  - **Notion-only (create-only):** automation fires on new Bid List rows and creates matching Common Bid List row. **Edits to existing Bid List rows do NOT sync** — Notion can't match by Project # to update. Accept this or do a periodic re-export.
  - **Power Automate (upsert):** scheduled or webhook-triggered, searches Common Bid List by Project # and either updates the existing row or creates a new one. Handles both creates and edits. Recommended long-term.

**Permissions:** Full access to Common Bid List for all employees. Zero access to Bid List itself, Customer List, Field Logs, Invoices, or any other database.

---

### 2. AR / Invoicing Hub (Tier 3 — financial)

**Purpose:** Invoice issuance, collections, and tying Bid Amount to actual billed.

**Users:**
- The bill clerk — High-Tier Clerk
- The collections clerk
- The AP/AR specialist

**Databases:**
- **Open Invoices** (owned here)
- **AR Project Summary** — a SEPARATE physical database (not a linked view of Bid List). Synced from Bid List with: Job Name, Project #, Customer (text), Bid Amount, Active Status. Same pattern as Common Bid List — separate DB avoids the cross-database permission issue.

**Key relation:** Bid Amount on AR Project Summary = sum of QBO invoices per Project # (pulled in via the QBO P&L system, see Connections below).

**How data gets in:** Same pattern as Common Bid List —
- **One-time backfill** — export current Bid List, trim to AR columns, import into AR Project Summary.
- **Ongoing sync** — same Option 1 (Notion create-only, no edit propagation) or Option 2 (Power Automate upsert by Project #) tradeoff. See the tier-enforcement section above.

**Permissions:** Full edit on Open Invoices and AR Project Summary. Zero access to full Bid List.

---

### 3. Project Admins Hub (Tier 3 — full project)

**Purpose:** Full project lifecycle — bidding, awarding, managing handoff to field, tracking to close.

**Users:**
- Estimators
- Project Managers

**Databases:**
- **Bid List** (owned here — source of truth)
- **Field Logs** (RP, CP — shared view from Field Hub)

**Permissions:** Full edit on Bid List. Can see and edit field logs where needed.

---

### 4. Field Hub (Tier 2)

**Purpose:** Supers manage their assigned phases — update dates, check off tasks, upload photos, access plans. **No sales info, no bid amount, no customer pricing, no margins.**

**Users:**
- Superintendents
- Estimators (cross-access from Tier 3)
- Project Managers (cross-access from Tier 3)
- The project clerk — High-Tier Project Clerk

**Databases:**
- **Field Logs** (RP Field Log, CP Field Log — owned here)

**What supers see:**
- Field Log phase rows for their jobs
- Rolled-up Tier 2 info from Bid List: Job Name, Project #, Address, City, Customer, Super assignment, **Plan file**
- Phase: stage name, start/end dates, checklist, photos, Quick Note

**What supers DO NOT see (hard-blocked):**
- Bid amount, sales price, margin, estimator notes, customer contact details beyond customer name
- Anything else on Bid List not explicitly rolled up to Field Log

**Enforcement:** Supers have zero access to the Bid List database itself. Their Field Hub teamspace only contains Field Logs. Sensitive Bid List fields are simply never rolled up. No view toggle can expose them.

**View strategy:** Personalized "My Jobs" page per super — Notion's current-user filter on the Superintendent rollup, so each super only sees their own phase rows. Estimators/PMs/the user see everything by clearing the filter. This is clutter, not security (security is enforced by the teamspace boundary above).

---

### 5. Subcontractor Hub (mixed tier — project-connected)

**Purpose:** Manage sub compliance (insurance, W-9s, agreements) and pricing. Tied to projects through sub assignments per job.

**Users:**
- Subcontractor Clerk
- Payroll Manager
- Estimators
- Project Managers

**Databases:**
- **Subcontractor Compliance** — list of subs with doc expiry dates, insurance status, W-9 on file
- **Subcontractor Prices** — rate cards / unit prices by sub

**Relation to projects:** Each sub record can relate to the Bid List projects they're working on. Future enhancement: track sub payments per project (the "burn tracking" idea from memory).

---

### 6. Payroll Hub (OFF-SPINE — not project-connected)

**Purpose:** Employee payroll data. Lives parallel to the project spine. Has no relation to Bid List, no Project # tie, no overlap with field or sales data.

**Users:**
- Payroll Manager
- Leadership

**Databases (TBD — to define):**
- Employee roster
- Timesheets
- Pay rate history
- PTO / leave tracking

**Why off-spine:** Payroll is an employee-level concern, not a project-level one. Employees work across many projects; tying payroll to individual Project #s creates complexity with no payoff. Keep it isolated. Future off-spine teamspaces (HR, internal ops, admin infrastructure) follow the same pattern.

---

### 7. Leadership / Owner (the user)

**Purpose:** Full visibility across every other teamspace. Automatic owner access.

**Users:** the user, co-owners

**Access:** Everything.

---

## Connections

- **QBO API Script** → [[QBO Project P&L System — Build Plan]]
  - Exports QBO data (invoices, expenses, P&L per project) out of QuickBooks Enterprise
  - Feeds back into the Bid List as Bid Amount roll-forward and into Open Invoices as actual invoice list
  - Key tie: QBO Project # ↔ Notion Project # — must be the same identifier

- **Live Excels** → shared company Excels prepopulated from Notion via one entry point
  - Recommended bridge: **Power Automate** (already in your Microsoft 365 stack via Teams)
  - Trigger: Notion page create/edit → Power Automate → write to specific sheet on OneDrive/SharePoint
  - Outcome: supers/estimators enter data once in Notion; company shared Excels stay current

---

## Access matrix (quick reference)

Columns show **what database access the role has**. "—" means zero access. Tier column shows the info tier the role operates at.

| Role | Tier | Bid List | Field Logs | Open Invoices | Sub Compliance | Payroll |
|---|---|---|---|---|---|---|
| Superintendent | 2 | — | Own jobs only (Tier 2 rollups only — no bid amount, no sales) | — | — | — |
| Estimator | 3 | Full | Full | — | View | — |
| Project Manager | 3 | Full | Full | — | View | — |
| AP/AR (the bill clerk, the collections clerk, the AP/AR specialist) | 3 | — (via AR Project Summary — separate DB) | — | Full | — | — |
| Sub Clerk | 1 | — (via Common Bid List — separate DB) | — | — | Full | — |
| PO Clerk | 1 | — (via Common Bid List — separate DB) | — | — | — | — |
| Payroll Manager | off-spine | — | — | — | View | Full |
| The project clerk | 2 | — | Full | — | — | — |
| The user (Owner) | 4 | Full | Full | Full | Full | Full |

**Hard rules to preserve:**
- Supers never see Bid Amount or sales info. If a sales figure needs to surface to the field, it's approved explicitly and rolled up to Field Log — not exposed wholesale via Bid List access.
- Common Clerks (Tier 1) never see plans, financials, or field operational detail. They only see the 5-column linked lookup.
- Payroll stays off-spine. No Project # relation there, ever.

---

## Open items / decisions needed

1. **Payroll Hub contents** — what databases actually go in here? (Timesheet, Employee Roster, Pay Rates, etc.) Define when ready.
2. **PO workflow** — PO Clerk is in Common Group but no PO database is named. Does PO tracking live in a new "Purchasing" hub, or inside an existing teamspace?
3. **Subcontractor → Project link granularity** — one sub can work on many jobs. Does that relation live on Subcontractor Compliance (multi-select project relation), on Bid List (list of subs per project), or both ways (dual relation)? Both ways is cleanest.
4. **Bid List's "common" view ownership** — does the linked view live in a dedicated "All Employees" utility page, or embedded on a company-wide home page?
5. **Commercial phase list** — still needed to finish CP Field Log build (Residential done).

---

## What's already built vs what's next

**Built:**
- Bid List (existing, kept as-is)
- RP Field Log database (10 stages, dates, photos, formula-driven status, checklist templates planned per stage)
- Send-to-Field-Log button on Bid List (conditional by Division)
- QBO P&L export system (Phase 1 complete, Phases 2–5 deferred)

**Next, in order:**
1. Finish RP Field Log: delete Waste 10%, add Start/End Date, convert Status to formula, add checklist templates per stage, add Last Edited audit columns, update button to use templates
2. Set up teamspaces per roster above — move Bid List into Project Admins teamspace, move Field Logs into Field teamspace
3. Build **Common Bid List** as a separate database in an all-employees teamspace
   - Create DB with 5 text columns
   - **One-time: export current Bid List to CSV, trim columns, import into Common Bid List** (backfill)
   - Choose sync model: Notion-only (create-only — edits don't propagate) OR Power Automate upsert (handles both). See "Data sync between tiers" section for the tradeoff.
4. Build **AR Project Summary** as a separate database in AR teamspace — same backfill + sync pattern, same Option 1 vs Option 2 decision
5. Add Plan file property on Bid List (file type, rollup to Field Log phase rows)
6. Get CP phase list → build CP Field Log → extend button with Commercial branch
7. Build Subcontractor Compliance + Prices databases
8. Define Payroll Hub contents
9. Power Automate bridge from Notion → company Excels (same tool, extend to handle the Excel output)
