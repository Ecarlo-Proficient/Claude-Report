---
purpose: Reconciles a vendor statement (PDF / image / Excel) against QBO open + recently-paid bills, produces an Excel report with the statement embedded and a per-bill QBO deep-link, and (in inbox mode) runs as a Synology folder automation.
inbox command: statement-reconcile --inbox
single command: statement-reconcile   (then drag a file, Enter)
code: /Users/sebas/Documents/Claude/Projects/Automate Concrete Business/statement-reconciler/statement_reconciler.py
Status: Active — v3.21 + inbox automation + embedded statement (2026-07-01)
---

## What it is

A single Python tool that takes a vendor statement, pulls that vendor's open and recently-paid bills from QuickBooks, matches them line by line, and writes an Excel report. It runs two ways: as a one-off on a file you hand it, or as an inbox sweep over a Synology folder. Auth is the shared QBO Keychain blob via `qbo_vault.py` (one Touch ID per run).

## Folder layout

Code (its own folder, split out of bill-tracker/ on 2026-07-01):

```
Automate Concrete Business/
├── qbo_vault.py                     shared Keychain auth (project root)
└── statement-reconciler/
    ├── statement_reconciler.py      the tool
    ├── vendor_aliases.json          vendor → QBO id cache
    └── .venv/                       its own virtualenv
```

Synology automation workflow (`/Volumes/Accounting/Automations/Vendor Statements/`):

```
Vendor Statements/
├── - Statement Inbox -/     clerk drops the statement here  (INPUT)
│   └── DONE/                source file archived here after it's reconciled
└── Reconciliations/         the finished Excel lands here    (clerk's WORK queue)
    └── Old-Done/            clerk moves the Excel here after fixing the bills
```

The Excel is self-contained: the statement is embedded inside it as a "Statement" tab, so moving files between these folders never breaks anything.

---

## Process 1 — Clerk daily reconciliation

1. Clerk drops the vendor statement (PDF, photo, or Excel) into `- Statement Inbox -`.
2. The inbox sweep runs (see Process 2) — it reconciles each statement, writes the Excel to `Reconciliations/`, and moves the source file into `Statement Inbox/DONE/`.
3. Clerk opens the Excel from `Reconciliations/`. It has the reconciliation on the Summary tab and the original statement on the Statement tab — everything in one file.
4. Clerk works the exceptions: clicks the ↗ in the leftmost column of any row to open that bill in QBO, and fixes what the report flags (enter missing bills, correct wrong amounts).
5. When done fixing the bills for that statement, clerk drags the Excel into `Reconciliations/Old-Done/`.

At any moment: files in the Inbox = not yet reconciled; files in DONE = reconciled; Excels in Reconciliations = waiting on the clerk; Excels in Old-Done = fully worked.

## Process 2 — Run the inbox sweep

1. Preview first (no QBO calls, nothing written or moved):

   ```
   statement-reconcile --inbox --dry-run
   ```

   It prints the Inbox / Out / Done paths and every file it would process. Confirm the paths and file list look right.

2. Run it for real:

   ```
   statement-reconcile --inbox
   ```

   Touch ID prompts once. For each statement it reconciles → writes the Excel (statement embedded) to `Reconciliations/` → moves the source to `DONE/`.

3. Read the INBOX SUMMARY at the end. "Reconciled + moved to DONE" is the success count. Anything under "Left in inbox (need a human)" stayed in the Inbox on purpose — see Process 4.

## Process 3 — One-off statement (ad-hoc, outside the inbox)

1. Run the bare command and drag the file into the terminal, then Enter:

   ```
   statement-reconcile
   ```

   Or pass the path directly: `statement-reconcile ~/Downloads/Statement.pdf`

2. Confirm the parse-sanity check (vendor, date, totals) when prompted.
3. Confirm the vendor match (skipped automatically if the vendor is already cached).
4. The Excel opens automatically. Ad-hoc runs land in the OneDrive statement-reconciles folder, not Synology. Add `--embed` if you want the statement embedded on a one-off run too.

## Process 4 — Onboard a new vendor

Inbox mode only auto-runs vendors it already knows (in the alias cache). A brand-new vendor is deliberately left in the Inbox so a wrong fuzzy match never happens unattended.

1. When the summary shows a statement "left in inbox — vendor not in alias cache", run that one file by hand:

   ```
   statement-reconcile "/Volumes/Accounting/Automations/Vendor Statements/- Statement Inbox -/<that file>"
   ```

2. Confirm the vendor match when prompted — this saves the alias.
3. From then on that vendor sweeps automatically. Move the file to DONE (or just re-run `--inbox`, which now handles it).

## Process 5 — Setup / reinstall (new machine or lost venv)

1. Install Tesseract once (for photo/scanned statements): `brew install tesseract`
2. Create the venv and install deps:

   ```
   cd "/Users/sebas/Documents/Claude/Projects/Automate Concrete Business/statement-reconciler"
   python3 -m venv .venv
   .venv/bin/python -m pip install --upgrade pip
   .venv/bin/python -m pip install pdfplumber openpyxl requests pytesseract pillow pillow-heif 'xlrd<2'
   ```

3. Point the `statement-reconcile` shell function in `~/.zshrc` at `statement-reconciler/.venv/bin/python3` and `statement-reconciler/statement_reconciler.py`, then `source ~/.zshrc`.
4. No new credentials — it reuses the QBO Keychain blob. If the bill tracker runs, this runs.

---

## The Excel report

- Summary tab: tie-out (statement total vs QBO open as-of, reconciling difference) and a collapsible section per category.
- Leftmost ↗ "Bill" column: opens that bill in QBO (`app/bill?txnId=…`). The Ref # columns are plain text on purpose — only the arrow is a link.
- Approved? column: flags bills whose QBO memo starts with "Not Approved" (needs PM signoff before payment).
- Statement tab: the source statement's pages embedded as images (self-contained; survives file moves). On by default in inbox mode; `--no-embed` to skip, `--embed` to force it on a one-off run.

## The 6 reconciliation categories

| Category | Trigger | Who acts |
|----------|---------|----------|
| MATCHED | Same Ref#, same amount (within $0.01) | nobody — clean |
| VENDOR_TAX_VIOLATION | Same Ref#, stmt = QBO × 1.0825 | Ted (vendor call) |
| CLERK_AMOUNT_MISMATCH | Same Ref#, any other amount diff | AP (fix QBO entry) |
| LIKELY_VENDOR_LAG | Stmt ref# matches a recently-paid QBO bill | AP (verify check cleared) |
| MISSING_IN_QBO | Stmt ref# not in any bill (open or paid) | AP (enter bill) |
| MISSING_ON_STATEMENT | Open QBO bill with no matching stmt ref# | Ted (likely vendor credit / payment in-flight) |

Vendor-tax rule: Proficient has no-sales-tax agreements with TX vendors (Estrada at minimum). A statement amount exactly 8.25% above QBO is the vendor billing tax against the agreement — a vendor issue, not an AP clerk error.

## Supported statement formats

Auto-detected by report signature (never by vendor name): QuickBooks Statement, QuickBooks Customer Open Balance, QuickBooks Open Invoices, Vendor Tabular (e.g. CMC), Vendor Columnar (e.g. Preferred / Sunrise), White Cap / Billtrust, and Excel (.xlsx/.xls). Photos/scans (.png/.jpg/.heic) go through Tesseract OCR then the same template chain.

## Vendor alias cache

`statement-reconciler/vendor_aliases.json` maps the parsed vendor string to the QBO vendor Id + DisplayName. Saved automatically after the first confirmed match; later runs skip the search and confirmation. If a cached id no longer resolves in QBO, it warns and re-searches — stale cache never causes a wrong match.

```
statement-reconcile --list-aliases
statement-reconcile --forget-vendor "Estrada Ready Mix"
statement-reconcile <file> --no-cache
```

## How failures behave (inbox mode)

- Unknown vendor, unreadable file, or a QBO error → the file stays in the Inbox (not moved to DONE) and is listed under "need a human." Nothing is silently lost.
- Only a clean success moves the source into DONE.
- If the Synology share isn't mounted, it stops immediately with a clear message rather than writing to the wrong place.
- Embedding failure never kills the run — the Excel still saves, just without the statement image, with a warning.

## Key CLI flags

| Flag | Purpose |
|------|---------|
| `--inbox` | Sweep the Synology Statement Inbox (reconcile → Reconciliations → DONE) |
| `--inbox --dry-run` | Preview the sweep: paths + file list, no QBO calls, no moves |
| `--embed` / `--no-embed` | Embed the statement as a tab / skip it (embed is default in inbox mode) |
| `--vendor "Name"` | Force the vendor (exact QBO DisplayName) |
| `--yes` | Skip confirmation prompts |
| `--list-aliases` / `--forget-vendor NAME` | Inspect / remove saved vendor aliases |

## Notes / limits

- On-demand, not always-on: the inbox is swept when you run `--inbox`, not by a background watcher. This is deliberate — launchd and the Synology mount have both been flaky, so a manual sweep is more reliable. A scheduled morning sweep can be added later if wanted.
- Embedded statement pages are images (not selectable text); each page adds roughly 100–300 KB. Cap is 20 pages.
- Ad-hoc runs output to OneDrive; inbox runs output to the Synology Reconciliations folder.

## History

- 2026-05-19/20 — v1–v3: built against Estrada; progress UI, alias cache, vendor-lag detection, clerk-performance CSV, batch mode.
- 2026-06 — v3.21: validated against 5–6 real vendor PDFs (Ready Cable, CMC, Preferred, Sunrise, White Cap, Post-Tension); added QB Open Balance / Open Invoices / columnar / White Cap templates, Excel + image (OCR) inputs.
- 2026-07-01 — split into its own `statement-reconciler/` folder; added the ↗ QBO bill-link column (Ref# stays non-clickable); added Synology inbox automation (`--inbox`); added embedded-statement tab (`--embed`) so the workbook survives file moves; hardened date-parse/grand-total/cache-key edge cases.
