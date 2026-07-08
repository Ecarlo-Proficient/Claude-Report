# RP Field Log — Per-Stage Template Source

One `.md` file per Notion page template in RP Field Log. Each file is the **body content** ready to paste directly into the matching Notion template.

This folder is the source of truth. When a checklist needs to change, edit the `.md` file here first, then update the Notion template to match.

## Stages and % Job Complete

6 contract milestones carry % toward Job Complete; 3 are excluded (Piers, Plumb Inspect, Cable Order Date — they aren't completion milestones or don't apply to every job). Punch Work is post-contract and excluded from the rollup entirely.

| Stage | File | % Job Complete |
|---|---|---|
| Set Forms | [set-forms.md](set-forms.md) | 10% |
| Piers | [piers.md](piers.md) | — |
| Trench | [trench.md](trench.md) | 10% |
| Plumb Inspect | [plumb-inspect.md](plumb-inspect.md) | — |
| Cable Order Date | [cable-order-date.md](cable-order-date.md) | — |
| Grade | [grade.md](grade.md) | 40% |
| Back Out | [back-out.md](back-out.md) | 10% |
| Pour Slab | [pour-slab.md](pour-slab.md) | 20% |
| Wreck | [wreck.md](wreck.md) | 10% |
| Punch Work | [punch-work.md](punch-work.md) | — (post-contract, manual only) |

**Completion is NOT a stage.** When Wreck closes, the PM flips the Bid List row's `Active Status` to `Completed`. No stage row is created for it.

## How to create the templates in Notion (one-time setup)

For each stage above:

1. Open **RP Field Log** in Notion.
2. Click the `▾` next to the blue **New** button (top right) → **+ New template**.
3. Name the template exactly as shown (e.g. `Set Forms` — case-sensitive).
4. On the template page, in the right-side property panel, set `Stage` = matching select option (e.g. `Set Forms`).
5. Open the matching `.md` file from this folder. Select All, Copy.
6. Paste into the template body. Notion auto-converts the markdown to checkbox / heading blocks.
7. Close the template — auto-saves.

Repeat for all 10. Order doesn't matter; what matters is that all 10 exist with the correct `Stage` preset before you build the automation.

## Conventions used in every template body

- `**Photos required**` — what photos the super attaches to the `Photos` property before the stage is considered complete.
- `**Checklist**` — to-do items the super checks off as work happens. Notion converts `- [ ]` to checkbox blocks automatically.
- `**Notes**` — free-text log (separate from `Quick Note` which is the column-level one-liner).

Supers fill `Start Date` when they begin and `End Date` when they finish. The `Status` formula auto-flips Not Started → In Progress → Done from those dates.

## Punch Work specifics

Punch Work is the only template applied **manually** — no button or automation creates it. When a builder calls back for post-completion work:

1. PM clicks the `▾` next to **New** in RP Field Log → picks the **Punch Work** template
2. Fills in `Project #` + the other copy-forward fields by hand (or copies from one of the project's existing stage rows)
3. Manually flips the Bid List row's `Active Status` from `Completed` → `In Progress`
4. When the punch closes, PM flips Active Status back to `Completed`

Multiple Punch Work rows are allowed for separate call-backs.

## Updating templates later

Two-step flow:

1. Edit the `.md` file here (this folder is the canonical version).
2. Open the corresponding Notion template, replace its body with the new content (Cmd+A on the Notion side, paste).

Existing stage rows already in Field Log don't update automatically — Notion templates only apply at row creation time. If a checklist change matters retroactively, you'd need to add the new items to existing rows by hand (or accept that older rows reflect the older checklist).
