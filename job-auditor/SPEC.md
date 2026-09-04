# Job Auditor — spec

An agent that watches the **live job folders on Synology**, detects what changed, and audits
the **numbers** against the rules the business runs on: priced scope in the proposal vs cost
bands in the takeoff. Flags carry a dollar figure. Proposes; never applies.

Supersedes the discarded `vault-steward/` spec. Status: design v3, calibrated against two
live runs over the active board (§2). See `STATUS.md`.

---

## 0. Why v1 was scrapped

v1 watched `~/Documents/Claude/Projects/` — the code repo — and reported stale `STATUS.md`
files and untracked scripts. Two failures:

1. **Wrong tree.** "My common folder current projects" meant
   `/Volumes/Common/CURRENT PROJECTS/` — the Synology share holding takeoffs, bid proposals,
   and plans. Not the code.
2. **Wrong altitude.** File hygiene is nearly worthless here. The errors that cost money are
   *inside* the workbooks and PDFs: an ETC missing a scope, a cost band with no matching sold
   line. A commit watcher cannot see any of it.

Only the runtime plumbing survived (launchd, watermarks, log location) plus the
Observed/Inferred discipline in §6.

---

## 1. What it watches

| Source | Path | Why |
|---|---|---|
| Job folders | `/Volumes/Common/CURRENT PROJECTS/Residential` (`RP.RP_ROOT`) | takeoffs (`.xlsm`), bid proposals (`.pdf`), plans |
| Schedule | `/Volumes/Common/OPERATIONS/SCHEDULE` | the active-job list — source of truth for what is live |
| General List | `RP.ALPHA_PATH` | cross-check only, lags the schedule (read-only, never written) |

QBO is **out of scope for the audit layer** — this compares *estimate to estimate* (proposal
vs takeoff). Costs-to-date belong to the existing WIP tools.

---

## 2. Measured baseline — 2026-08-11, live board

Run against the live board (`Schedule 8-11-26`). Hit rates only here; **exposures and the job
list live in the vault** (`tasks/2026-08-11_job-auditor.md`), per the repo scope filter.

**Pin the input version in every output.** An earlier pass read a two-week-old schedule off a
stale network mount, reported no error, and looked completely healthy while answering questions
about a board that no longer existed. Every run must record the schedule filename and mtime.

| Rule | Hit rate | Verdict |
|---|---:|---|
| **Reader mis-reads the cost-sheet variant → ETC wrong in both directions** | 32/37 classifiable | **VERIFIED. Highest value found.** 32 cumulative (reader double-counts the slab) vs 5 band-only. Defect is in OUR code, not the workbooks. 23 more unclassified. |
| **The variant FLIPS as workbooks are edited** | 9 jobs in 13 days | **Classification must be read every run and never cached.** One flip silently moves a job's ETC by a whole slab. |
| Hand-typed constants appended to a subtotal formula | 3/56 | Real; invisible to anyone reading the sheet. |
| Sold pier line vs PR cost band — cost exceeds revenue | 2/14 reconciled | **Was an artifact of the formula bug above**, not a margin problem. Rule 2 must run after rule 1. |
| **Cost band with no sold line — REAL cost, misfiled onto the Piers sheet** | 4/45 | **Verified: site lines (tie wire / scrape lot / reset forms) feed `PR3`/`PR6`.** Reattribute; never drop. |
| **PR cost carried, no proposal on file** | 7/45 | Real; overlaps the tract/no-proposal population. |
| **`#N/A` SLAB band → dropped from ETC** | 2/45 | **Real. Silent.** The main scope vanishes with no notice. |
| **`#N/A` PIERS band → dropped from ETC** | 19/45 | **Real.** Pier cost never reaching the ETC on ~42% of jobs. |
| FW cost present on a slab takeoff | 35/45 | Data noise — already excluded by both ETC paths. See §3. |
| Proposal fails the reconciliation gate | 13/34 | **Parser gap, not a finding.** See §5. |
| No `Cost Gral` sheet — unauditable | 3/48 | Structural; needs a human. |

---

## 3. Rulings

**Piers ARE sold as their own priced line — reversing the v2 ruling.** The proposal is a
line-item table (`DESCRIPTION · UNIT · PRICE/UNIT · TOTAL`) and piers carry their own
quantity, unit price, and total. The v2 conclusion that piers were "described but never
priced" was an artifact: the description **wraps across lines**, so the line containing
"Piers" holds no money and the line holding the money contains no "Piers". Any line-level
regex gets this wrong.

**Therefore scope-level comparison is possible** — each sold line has a total, each cost band
has a subtotal — and it is what surfaced the defect below. But it is a *detector*, not a
verdict: the jobs it flagged as "pier cost 2.4×–3.7× pier revenue" turned out to have a
corrupt cost input, not a margin problem.

**VERIFIED: the cost sheet has TWO scopes (foundation = slab+piers, and flatwork), and our
reader assumes three bands.** The piers-band foot is the FOUNDATION subtotal and legitimately
sums from the slab subtotal down. Our `ETC = SL_sub + PR_sub` therefore double-counts the slab
on every sheet built that way. **The spreadsheets are right; the reader is wrong.** Both
variants are in circulation (~60/40), so the sheet must be classified, not assumed.

**Generalised rule this produced — the expensive one:** an initial reading ("the formula is one
row too tall") was defended with four checks that all turned out to be consistent with the
opposite conclusion. **A test that cannot fail is not evidence.** Before treating a
cross-check as confirmation, state what result would refute the hypothesis; if nothing would,
it isn't a test. Ask the person who built the artifact before concluding it is broken.

**The `#N/A` band drop is a separate, silent defect.** Both live ETC paths correctly exclude
FW from a slab ETC — so flatwork on slab takeoffs is data noise, not budget error (exception:
a *side-scope* takeoff, filename matching the schedule description, deliberately sums
SL+PR+FW). What actually corrupts the ETC is a band whose subtotal cell errors: the band is
dropped whole and nothing says the budget is now missing its slab.

---

## 4. Rule catalog

Ordered by measured value:

1. **Cost-sheet VARIANT classification** — read the piers subtotal's *formula range*, decide
   which template variant the sheet is, and compute the ETC accordingly.

   **The design (from the estimator, confirmed):** the sheet holds **two scopes, not three
   bands**. **FOUNDATION = slab + piers**, whose SUB TOTAL is the piers-band foot and therefore
   sums *from the slab subtotal down*. No piers on the job ⇒ the pier rows read 0 and you read
   the slab subtotal instead. **FLATWORK** is its own scope, starting at its own first item.

   **The defect is in OUR reader, not the workbooks.** It computes `ETC = SL_sub + PR_sub`,
   but on the cumulative variant `PR_sub` already contains `SL_sub` — so the slab is counted
   twice. **Never edit the takeoffs to "fix" this.**

       CUMULATIVE  (range starts at the slab subtotal row)  → ETC = PR_sub alone
       BAND_ONLY   (range starts at the first pier row)     → ETC = SL_sub + PR_sub

   **Both variants are live and a job MOVES BETWEEN THEM.** The current board runs ~32:5 toward
   cumulative, but nine jobs flipped variant in 13 days — including one reversing from cumulative
   to band-only. **Never cache a classification and never carry one forward from a prior run.**
   Classification is the rule, and it is the prerequisite for every value-level rule, which
   otherwise reports a mis-read input as a business problem.

   Two follow-ons: an `#N/A` subtotal on a cumulative sheet makes the reader drop the band and
   *lose* the piers (understating), and hand-typed constants appended to a subtotal
   (`=SUM(...)+7975.25+4600`) are common and invisible to anyone reading the sheet.

2. **Sold scope vs cost band** — for each priced line, compare to the matching cost band.
   Flag when cost exceeds revenue, or margin falls outside the expected range. Requires §5,
   and requires rule 1 to pass first — otherwise it reports corrupt inputs as margin problems.
3. **`#N/A` band drop** — any of SL/PR/FW with an errored subtotal but non-zero items.
   Report the band, the item sum, and the cell.
4. **Missing slab** — SL contributes 0 to a slab ETC. Near-certainly wrong.
5. **Cost band with no sold line → REATTRIBUTE, don't drop.** A pier/flatwork cost band with
   no matching sold line is usually **real cost on the wrong sheet**, not phantom cost. The
   `Piers takeoff` sheet is not purely piers: its subtotal ranges sweep in site/general lines
   (tie wire, scrape lot, reset forms) that map to `PR3`/`PR6`. On a job with no piers every
   genuine pier line is zero and only those survive. **Verified: the slab sheet carries no
   scrape-lot or reset-forms line, so excluding the band loses real money.**

   Ruling: **count the money, flag the miscoding.** Report it as site cost sitting on the Piers
   sheet rather than silently dropping or silently accepting it. **Carve-out:** tie wire appears
   on both sheets and is sometimes identical on each — an identical-amount pair across two
   sheets is a possible duplicate and must be reported, not summed. Tract price-list jobs exempt
   (no proposal exists by design).

   **Beyond the WIP:** SL/PR/FW codes drive Budget-vs-Actual in the JobTread migration, so cost
   landing on a pier code for a job with no piers corrupts the comparison at code level, not
   just in total. Durable fix is a master-template change, not per-job edits.
6. **Sold line with no cost band** — the inverse.
7. **Implausible contract** — priced under $1,000 on a slab (wrong cell picked up).
8. **Unauditable input** — no `Cost Gral` sheet, or a proposal failing the §5 gate.
9. **Structural drift** — takeoff or proposal template changes shape. This is what makes
   every other rule quietly stop working.

---

## 5. Detector discipline (binding — earned, not assumed)

Four detectors produced confident, wrong numbers in a single afternoon:

| # | Bug | Effect |
|---|---|---|
| 1 | tested `PR subtotal > 0` | blind to the ~42% of jobs whose PR subtotal is `#N/A` |
| 2 | tested `\bpier\b` | **misses the plural "PIERS"** — invented 30 phantom jobs |
| 3 | line-level regex for "pier + a price" | descriptions **wrap**; concluded piers were never priced |
| 4 | `group(2)`/`group(3)` after adding named groups | shifted every unit and total one position — a full table of plausible, wrong money |
| 5 | stopped at the cost sheet instead of the source sheet | called real-but-misfiled site cost "phantom"; the answer was one sheet deeper |

None of the four was visible in the output. All four looked like findings.

**Binding rules:**

- **Reconciliation gate.** A proposal parse is trusted only when **Σ(sold line totals) equals
  the proposal's own `SUB TOTAL`** to the cent. Anything else is not evidence and cannot
  produce a flag — it goes on the parser-gap list. This is objective and self-checking.
- **Fixture first.** Every parser is validated against at least one hand-checked document
  before any portfolio number is reported.
- **Signal-quality line.** Every rule prints its match count and a sample. A rule matching
  ~0% or ~100% is presumed broken, not informative.

---

## 6. Output

Per run: `~/Library/Logs/Proficient/job-auditor/YYYY-MM-DD.md`, plus a plain xlsx when there
are flags (white/black, label + amount on the same row, one sheet per rule).

Two registers, never merged:
- **Observed** — "<job> sold <amount> of piers; PR cost band is <amount>."
- **Inferred** — "*Likely* the takeoff covers a scope the proposal doesn't." Always hedged.

**It never writes:** takeoffs, proposals, the General List, the WIP master, QBO, or the
vault's `log.md`. Vault updates are proposals only.

---

## 7. Change detection — the "what changed" layer

Watermarks at `~/Library/Logs/Proficient/job-auditor/watermarks.json`, keyed by job folder:
`(path, mtime, size, sha256)` per takeoff/proposal.

This turns the audit into a story: *"<job>'s takeoff was revised on the 20th; its pier band
doubled and no revised proposal was issued."* The flag list says what is wrong; the change
layer says what *just became* wrong — the part worth a weekly read.

Hash, not mtime — the share is network-mounted and mtimes move without content changing.

---

## 8. Runtime

- launchd weekly, Monday early; template with `/ABSOLUTE/PATH/TO/…` + sed install line,
  mirroring `invoice-sync/launchd/`.
- Ships `.disabled`. Run by hand until the flag list stops changing shape run-over-run.
- Requires the `Common` volume mounted; exits cleanly with a stated reason if not.
- No QBO ⇒ **no Touch ID, no token** — safe to run unattended.

---

## 9. Build order

**Phase 1 — rule 1 (subtotal formula range) alone.** Verified defect, mechanical, needs no
PDF parsing and no proposal at all: open the workbook, read the band-foot formulas, check the
ranges. It is the prerequisite for every value-level rule, because those report corrupt inputs
as business problems. Ship with the two known-bad and one known-good workbook as fixtures.

**Phase 1b — the reconciling proposal parser** (§5 gate) plus rules 2–4. The parser is the
enabling piece: it already reconciles on 14 of the reconcilable proposals and it is what makes
scope-level margin possible. Ship with the fixture test.

**Phase 2 — close the parser gap.** 13 proposals fail the gate today: 7 are a tract price list
with no `SUB TOTAL` (exempt by design — detect and skip), 6 are real proposals whose totals
don't reconcile and need the grammar extended.

**Phase 3 — rules 4–5** (scope presence both directions).

**Phase 4 — change detection** (§7), then schedule it. **Promoted from nicety to essential:**
because a job's variant flips as its workbook is edited, the interesting event is the change
itself — a silent flip moves that job's ETC by a whole slab with nothing in any report to say so.

---

## 10. Honest limits

- It compares the **proposal to the takeoff**. If both are wrong the same way, it sees nothing.
- It cannot read **plans** — pier counts and depths live in drawings, often as images. Scope
  that exists only on the plans is invisible.
- **Cost bands do not map 1:1 to sold lines.** A rate line can absorb costs that sit in another
  band, so a single-line margin is a flag to verify, never a verdict.
- Rule 8 (template drift) is unwritten and is the failure mode that would silently disable
  everything else.
