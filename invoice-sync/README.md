# Invoice Sync — Folder Guide & Onboarding

> **Renamed from `automation-worker/` in the 2026-07-13 restructure.** This
> folder now holds exactly one tool: the QBO → Notion invoice sync (plus its
> own config/clients/verifiers). The WIP readers moved to `../wip/`; shared
> code (QBO vault, paths) moved to `../shared/`; the Field Log flow was
> removed entirely.

**Read this first.** This folder looks busy (20+ scripts), but most of them serve
**one** live job: the **AR Invoice Sync** (QuickBooks → Notion). The rest are
shared libraries it leans on, setup/diagnostic tools, read-only checkers, and a
couple of **dormant/retired** flows left in place. This guide is layered:

- **Part 1 — Plain English:** what runs, how to run it, how to tell it worked,
  what to do when it breaks. Read this to operate the system.
- **Part 2 — Technical map:** every file's role, the data flow, where things
  live, and the gotchas. Read this to maintain or change the code.

---

# PART 1 — Plain English

## The one thing that runs here: the Invoice Sync

It pulls **open invoices from QuickBooks Online**, mirrors them into **two Notion
databases** (one for **MFD / Multifamily**, one for **RP/CP — Residential +
Commercial**), marks invoices **Paid** when QBO says they're paid, **removes**
invoices that were deleted in QBO, writes a read-only **Excel** copy for the
collections clerk, and posts a card to **Teams** when an MFD invoice gets paid.

Notion is the team's working view of who owes what; QuickBooks is the source of
truth. This sync keeps Notion matching QuickBooks, every run.

## How to run it

Run it **visually** (shows phases, a progress bar, and colors):

```
sync-ar
```

That's a shell alias that runs `run_invoice_sync.sh`, which launches the viewer
when you're in a terminal. To preview without changing anything:

```
cd "/ABSOLUTE/PATH/TO/Automate Concrete Business/invoice-sync"
python3 sync_view.py --dry-run
```

For a real run without the visual wrapper: `python3 run_invoice_sync.py`.

## How to tell it worked (normal vs. broken)

- **Normal:** it ends with per-division summary lines and `… clean.` The viewer
  shows green check-marks down the phase list.
- **Errors:** the viewer shows a yellow `⚠ COMPLETED WITH ERRORS` or red
  `✗ FAILED` panel and writes a **crash report** to
  `~/Library/Logs/Proficient/automation-worker/crash-<timestamp>.log`. A warning
  also posts to the Teams **ops-alert** channel.

## If it breaks — first moves (in order)

```
cd "/ABSOLUTE/PATH/TO/Automate Concrete Business/invoice-sync"
python3 doctor.py
```

`doctor.py` checks the basics (config present, Notion secret in Keychain, Notion
auth round-trip, and each Notion database is reachable) and stops at the first
failure. Note it checks the **Notion** side; QBO auth surfaces when you run the
sync itself. Then:

1. Read the crash report named in the failure panel (full traceback + last output).
2. Run `python3 verify_invoices.py` — a read-only audit comparing QuickBooks to
   Notion; it tells you if anything actually drifted.
3. Fix the cause, re-run `sync-ar`.

Common ones: a Touch ID prompt after a reboot (run once to approve), or Notion/QBO
auth (see Troubleshooting in Part 2).

## What this folder does NOT do anymore

- **Field Log sync** — **removed entirely (2026-07-13).** The Bid List → Field
  Log / Project Plans flow is gone from the codebase by decision.
- **WIP → Notion** — retired. WIP now lives in Excel on SharePoint; the WIP
  readers live in `../wip/` (they moved there 2026-07-13, and the old
  `wip_sync.py` stub + its launchd plist were deleted).

---

# PART 2 — Technical Map

## Run cadence & versions

- **Today:** run manually via `sync-ar`. The 15-min launchd schedule is paused
  (a macOS update broke it; plist kept in `launchd/`).
- **Two version lines:** Docker = **v1.0.0** (the "true release", in testing on
  Synology); Mac = **mvN** (currently `mv1` — the `sync-ar` / viewer path). Every
  run logs which one it is: `Starting … [v1.0.0 (docker)]` or `[mv1 (mac)]`
  (`version.py`).

## The mental model: entry points vs. libraries

The single most useful thing to teach someone: **most files here are libraries
you never run directly.** Only a handful are meant to be executed.

**Entry points — you run these (`python3 <file>`):**

| File | What it does |
|---|---|
| `run_invoice_sync.py` | The invoice sync itself. The production job. |
| `sync_view.py` | Visual wrapper around `run_invoice_sync.py` (phases, bar, crash report). What `sync-ar` launches. |
| `verify_invoices.py` | Read-only audit: QBO open invoices vs. Notion. Safe anytime. |
| `verify_excel_export.py` | Read-only three-way audit: Excel ↔ Notion ↔ QBO. |
| `doctor.py` | Preflight diagnostics (config, Notion secret, Notion auth, Notion DB reachability). |
| `setup_keychain.py` | One-time: store the Notion secret + Teams webhook in Keychain. |
| `test_teams_webhook.py` | Fire a fake Teams card to test webhook wiring. |

**Libraries — imported by the above, never run directly:**

| File | Role |
|---|---|
| `invoice_sync.py` | The core invoice logic (parse, route, upsert, sweeps, CDC). The big one (~1,400 lines). |
| `qbo_client.py` | QuickBooks API wrapper (auth, queries, CDC, customer hierarchy). |
| `notion_client.py` | Notion API wrapper (query/create/update, retries). |
| `config.py` | Loads `.env` config + secrets from Keychain/env. One `Config` object. |
| `logger.py` | Rotating file log + colored stdout. |
| `state.py` | Tiny JSON store for "last run" watermarks (used by the CDC pass). |
| `version.py` | Version identity (Docker v1.0.0 / Mac mvN). |
| `teams_notify.py` | Builds + posts Teams cards: MFD payments **and** failure alerts. |
| `export_invoices_xlsx.py` | Writes the read-only `Open_Invoices.xlsx` mirror (2 tabs). |
| `aging_sheet.py` | Builds the **AR Aging** tab of that workbook — see below. |

## Invoice sync — end to end

`run_invoice_sync.py` is the conductor; `invoice_sync.py` does the work:

1. **Load config** (`config.py`) — DS IDs from `.env`, secrets from Keychain (Mac) or env (Docker).
2. **Auth to QBO** (`qbo_client`) — one Touch ID per process on Mac.
3. **Pull open invoices** (`Balance > 0`) + the customer hierarchy + term map.
4. **Upsert** each invoice into the right Notion DB — MFD prefix → MFD tracker, RP/CP → Res/Com tracker (routed by Project # prefix). Customer matched by a 3-pass matcher.
5. **Flip / delete sweep** — invoices no longer open in QBO get marked **Paid**; ones deleted while still open get **archived**.
6. **CDC deletion pass** — asks QBO what was deleted since last run and archives those Notion rows (catches invoices deleted *after* they were marked Paid). Watermark in `state/invoice_cdc_state.json`.
7. **Archive** paid invoices older than 12 months.
8. **Excel export** (`export_invoices_xlsx`) — unless `SKIP_EXCEL_EXPORT=1`.
9. **Teams** — MFD paid/short-pay cards during the run; a failure/error alert at the end if anything went wrong.

A full diagram of this flow lives at the bottom of
`docs/Invoice Tracker — System Reference.md` (the canonical system doc — read it
alongside this folder guide).

## Where things live

| Thing | Location |
|---|---|
| Logs | `~/Library/Logs/Proficient/automation-worker/sync.log` (+ `crash-*.log`) — **outside** this folder by design (privacy). |
| State / watermarks | `state/invoice_cdc_state.json` |
| Secrets | macOS Keychain (Mac) or env vars (Docker) — **never** in files. `setup_keychain.py` writes them. |
| Non-secret config | `.env` (copy from `.env.example`) — DS IDs + behavior flags. |
| Schedules | `launchd/*.plist` — currently **disabled**; run manually via `sync-ar`. |
| Docker package | one level up in `../docker/` (the v1 release). |

## First-time setup on a new Mac

```
cd "/ABSOLUTE/PATH/TO/Automate Concrete Business/invoice-sync"
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Fill `.env` with the real Notion data-source IDs. Then connect the Notion
integration (Settings → the `Automation Integrator` integration → add it to the
Invoice Tracker DBs + Customer lists under each DB's **Connections**), and store
the secrets:

```
python3 setup_keychain.py
python3 setup_keychain.py --teams
python3 doctor.py
python3 sync_view.py --dry-run
```

On first run macOS asks to allow Python to read the Keychain key — click
**Always Allow**.

## Verifiers (read-only — safe against production, anytime)

- `verify_invoices.py` — QBO vs. Notion, invoice by invoice → markdown report in `reports/`.
- `verify_excel_export.py` — confirms the OneDrive Excel mirror matches Notion + QBO.

## Gotchas worth knowing (the load-bearing ones)

- **QBO query language is AND-only** — OR conditions are split into multiple calls.
- **QBO class names are spelled out** (`Residential` / `Commercial` / `Multi Family`), not codes.
- **QBO `CustomerRef.name` on an invoice is the sub-customer** — the GC name comes from walking `ParentRef` (built into `qbo_client`).
- **CDC lookback caps at 30 days** — deletions older than that aren't reported; persist `state/` so the watermark survives restarts.
- **Notion has no row-level security** — MFD is isolated by being a *separate database*, not a filter. Keep the two trackers separate.
- **Excel mirror is one-way and overwrites** — never edit `Open_Invoices.xlsx` by hand; an open-file guard skips the write while it's open.
- **`Open_Invoices.xlsx` has two tabs.** `Open Invoices` is the flat mirror. `AR Aging` (added 2026-08-05) is the owner's collections view: Current / 1-30 / 31-60 / 61-90 / 90+ buckets aged **by due date**, RP + CP + MFD together with a `Division` filter, invoices grouped under the parent client and **collapsed by default**, the collections clerk's `Quick Status` as Notes, and **litigation invoices excluded**.
- **The aging tab's vendor columns come from `Bill Tracker.xlsx`, not from QBO, and they describe the PREVIOUS draw.** MFD/CP funding is a chain: the GC funds draw N, we pay draw N's vendor bills, those vendors issue unconditional waivers, and the GC needs the waivers before releasing draw N+1. So `Prev Draw Status` answers why *this* draw isn't funded — `PAY BILLS → unlock` means the previous draw was funded but our vendors on it are still owed (ours to fix); `Waiting GC on prev` means the previous draw isn't funded either. `Prev Draw` names the invoice used, so a wrong pick is visible. RP is a grey `n/a` block — it doesn't bill in draws.
- **Draw sequencing is parsed out of invoice memos (`draw_chain.py`) — treat it as the fragile part.** Memos are not uniform (`May Draw 2026`, `Draw #2`, `Draw #3 December 2024`, `March 2025 Draw`, `- Retainage - Draw #5`), and a new spelling shows up as one project splitting into phantom contracts. Projects with genuinely parallel contracts (MFD192, CP861) report `Multi-contract` rather than guessing, because bills carry a project #, not a contract.
- **RUN AP BEFORE AR.** That column is only as fresh as the last `sync-ap`, so `sync-all` (which runs **AP → AR**, ~5 min) is the daily command. Running `sync-ar` alone leaves the vendor columns reporting the previous AP run. This is a one-way dependency, not a cycle: the bill tracker pulls its bills *and* its invoices straight from QBO and never reads Notion or `Open_Invoices.xlsx`. If the tracker predates today the tab says so in red and the run logs a warning; a missing file shows `?`, never a false "Vendors Paid".
- **Logs/state live outside this folder** — don't add a `logs/` dir here.

## Troubleshooting

- **401 from Notion** → secret wrong/missing: re-run `setup_keychain.py`; `doctor.py` says which stage failed.
- **404 on a Notion DB** → the integration isn't connected to that DB (add it under **Connections**).
- **Touch ID prompt every run** → click "Always Allow" once after a reboot.
- **Sync "completed with errors"** → open the crash report named in the panel; run `verify_invoices.py` to see real drift.
- **A deleted invoice still shows in Notion** → check its Notion **Status**; if it's already **Paid**, the CDC pass handles it on the next run (it can't be seen by the open-invoice sweep). See the System Reference doc.
