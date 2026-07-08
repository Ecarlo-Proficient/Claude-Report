# RP Field Log — Send to Field Log Setup

> **Note on filename:** the file is named `..._Button_Setup.md` for legacy reasons, but the trigger is now a **checkbox + database automation**, not a button. Content below reflects current model. Rename later if you want.

The mechanism that pushes a Bid List row into RP Field Log is a Notion **database automation** on Bid List, triggered by the `Send to Field Log` checkbox property changing to true. When the PM checks the box on a Sold Residential row, the automation creates 9 stage rows in RP Field Log using the matching page templates.

**Scope.** RP only. CP automation is deferred (commercial phases not defined yet — those rows will be created manually for now). MFD has no Field Log database, deferred entirely.

**Add Piers** is no longer button-driven either. If a project's `Pier Count` was 0 at the time the checkbox was checked but later turns out to need piers, the PM creates a Piers row in RP Field Log manually using the **Piers** template. No automation, no recovery button.

---

## One-time Notion setup

**1. Bid List** — two properties needed:
- `Send to Field Log` (checkbox, default unchecked) — already added.
- `Plans` (URL) — add this. PM enters the project's plans link (OneDrive/Synology/wherever) on each Bid List row when ready. Filled once per project, not per scope row.

**2. RP Field Log** — must have **10 page templates** before you build the automation. See [`field-log-templates/INDEX.md`](field-log-templates/INDEX.md) for the full walkthrough. Templates needed:

`Set Forms` · `Piers` · `Trench` · `Plumb Inspect` · `Cable Order Date` · `Grade` · `Back Out` · `Pour Slab` · `Wreck` · `Punch Work`

Each template has its `Stage` select preset to the matching value, and its body pasted from the corresponding `.md` file in `docs/field-log-templates/`. The automation fails at run-time if a template name doesn't match exactly (case-sensitive).

**3. RP Field Log schema** — confirm Job Name and Job Type columns are gone (cleaned up via API on 2026-04-28). Columns that should exist: Project #, Division, Builder, Superintendent, Square Foot, FTW Sq. Ft., Pier Count, Job Address, City, Stage Name (title), Stage, Active Status, Date Completed, Start Date, End Date, Photos, Quick Note, Lag (days) formula, Status formula.

**4. Project Plans database** — single source of truth for plan links, accessible to supers without giving them Bid List access. Schema: `Project #` (title), `Job Address` (rich_text), `Plans` (URL). Lives in the Field teamspace, private. Created via API on 2026-04-28; DS ID `80c4e252-4b4e-4345-b813-b9acd85b3ce6`. The automation writes one row per project at checkbox time. Sync keeps the Plans URL fresh if PM edits it on Bid List later.

---

## Automation setup (on Bid List)

Notion → Bid List → `...` (top right) → **Automations** → **+ New automation**.

### Trigger

`When` → **property changes** → **Send to Field Log**.

### Conditions (all three required)

The automation should only fire on the false → true transition for Sold Residential rows. Add three conditions:

- `Send to Field Log` **is** **checked**
- `Division` **is** **Residential**
- `Lead Status` **is** **Sold**

If any condition fails, the automation skips. (This blocks the case where someone unchecks then re-checks on a non-Sold or non-Residential row.)

### Actions

Add **10 actions** total: 9 stage rows in RP Field Log (8 unconditional + 1 conditional Piers) and 1 row in Project Plans. Build the first RP Field Log action carefully and verify, then duplicate it 8 times to save formula re-entry. The Project Plans action is a one-off, different shape (different target DB, fewer fields).

**Action 1 — Set Forms** (and the same shape for Trench, Plumb Inspect, Cable Order Date, Grade, Back Out, Pour Slab, Wreck — 8 of these unconditional)

| Field | Value |
|---|---|
| Template | `Set Forms` (pick from dropdown) |
| `Stage Name` | formula — see below |
| `Project #` | formula — see below |
| `Division` | direct: `Division` (Bid List property) |
| `Builder` | formula — see below |
| `Superintendent` | direct: `Superintendent` |
| `Square Foot` | direct: `Square Foot` |
| `FTW Sq. Ft.` | direct: `FTW Sq. Ft.` |
| `Pier Count` | direct: `Pier Count` |
| `Job Address` | direct: `Job Address` |
| `City` | formula: `prop("City").name` |

Leave **everything else blank** — Active Status, Date Completed, Start Date, End Date, Photos, Quick Note. Supers fill them in.

**Action 9 — Piers** (conditional on Pier Count > 0)

Same shape as the others, with **Template = Piers** and one extra condition on the action: `Pier Count > 0`. Notion lets you set per-action conditions in addition to the automation-level conditions.

**Action 10 — Project Plans** (one row per project)

Different target database. **Add page to → Project Plans**. No template needed (the DB only has 3 fields).

| Field | Value |
|---|---|
| `Project #` (title) | Project # formula (same as the RP Field Log actions) |
| `Job Address` | direct: `Job Address` |
| `Plans` | direct: `Plans` (the new URL property on Bid List) |

This creates one row per project in Project Plans, accessible to supers via the Field teamspace. They get the plan link without needing Bid List access.

### Formulas

**Project # — extract from title, untouched (any `-FTW` stays):**

```
slice(prop("Job Name"), 0, indexOf(prop("Job Name"), " - "))
```

This pulls the head of the Job Name title up to the first ` - ` (space-dash-space). It works for both `RPxxxx - <address>` and `RPxxxx-FTW - <address>` because the dash inside `RPxxxx-FTW` has no spaces around it and won't match the delimiter.

**Stage Name — compact label (`<Project#> — <Stage>`):**

```
slice(prop("Job Name"), 0, indexOf(prop("Job Name"), " - ")) + " — " + "<Stage>"
```

Replace `<Stage>` literal in each action with the matching stage name, e.g. `"Set Forms"`. Produces `"RP7482-FTW — Set Forms"`. Address shows up via the `Job Address` column in views — group views by Job Address for the cleanest UX (supers identify projects by address, the column is right there).

**Builder — relation flatten:**

```
prop("Builder").map(current.name).join(", ")
```

> If Notion's formula editor errors on the function-form `slice/indexOf`, swap to method-form: `prop("Job Name").slice(0, prop("Job Name").indexOf(" - "))`. **Build the FIRST action's formulas and verify they compile before duplicating to the rest** — saves re-editing eight broken formulas.

### What the automation does NOT do

- **Does not write Active Status anywhere.** Field Log's Active Status is supers-owned (per stage row); Bid List's is PM-owned. Both manual.
- **Does not flip Send to Field Log = true.** That's the trigger — the user already did it.
- **Does not handle dedupe.** If a user unchecks then re-checks Send to Field Log, the automation fires again and a second set of 9 stages gets created. Per Ted's directive (2026-04-28), duplicate prevention is handled by an agent inside Field Log, not the automation.

---

## Test workflow

1. Test row on Bid List: `RP9999 - TEST`, Division = Residential, Job Type = FOUNDATION, Pier Count = 5, Lead Status = Sold, Send to Field Log = unchecked.
2. **Check the box.** → 9 rows should appear in RP Field Log within seconds. Each row has its `Stage` preset, the checklist body from the template, and the copy-forward fields filled (Project # = `RP9999`, Division = Residential, Builder = whatever, etc.). Active Status on each stage row should be empty.
3. **Run sync:**
   ```bash
   cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/automation-worker"
   source venv/bin/activate
   python sync.py --dry-run
   ```
   Should show 1 tracked row, 9 stage rows that would be UPDATEd. Drop `--dry-run` to apply for real.
4. **Verify:** `python verify_field_log.py` → 0 issues for `RP9999`.
5. **Test FLATWORK variant:** `RP9998-FTW - TEST FTW`, Job Type = FLATWORK, Pier Count = 0, Sold. Check the box → 8 rows (no Piers since Pier Count = 0), Project # = `RP9998-FTW`.
6. **Delete the test rows** in both Bid List and Field Log when done.

---

## Common pitfalls

**Action runs but no rows appear in Field Log.**
- Integration `Automation Integrator` lost access to RP Field Log → re-add via Connections.
- Template name mismatch (case-sensitive) → check that all 10 template names match exactly.

**Rows appear but missing fields.**
- A formula didn't compile in one action — Notion silently leaves the field blank in that case. Re-open the action, re-enter the formula, verify the editor doesn't show a red squiggle.

**Duplicate rows after re-checking the box.**
- Expected. Field Log agent handles dedupe.

**Stage rows have weird Project # like `RP9999-` (trailing dash).**
- Bid List title has a malformed delimiter. Fix the title to use ` - ` (space-dash-space) between Project # and description.

**Re-running the dry-run shows new "Send=true" rows the sync doesn't recognize.**
- Sync is filtering on `Send to Field Log = true` AND looking for matching stage rows in Field Log. If the box is checked but the automation never fired (or fired and stages got deleted), sync will log "Send=true but no stage rows for Project # X". Either re-run the automation (uncheck + re-check) or manually create stage rows.
