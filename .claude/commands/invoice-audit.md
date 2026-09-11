---
description: Find QuickBooks invoices billed to the wrong customer for the job their memo names (read-only), and report what to re-point
---

Run the invoice-to-project audit and report it. Read-only: this never writes to QuickBooks. Fixing an invoice is a manual QuickBooks edit (open the invoice, change the customer to the project sub-customer); after the owner fixes, re-run and confirm the rows are gone.

1. Run it from the repo root (pass through any arguments the owner gave, e.g. `--since 2024-01-01` or `--all`):

```bash
cd "<repo root>" && python3 one-offs/invoice_project_audit.py $ARGUMENTS
```

2. Read the stdout summary and open `~/Documents/CompanyHealth/Invoices Off Project.xlsx` with openpyxl (never Excel). The rule the script applies: an invoice is off its project when its memo names exactly one job number, that job exists as a customer in QuickBooks, and the invoice is billed to some other customer. Memos naming no job or two jobs are counted and left alone. Two refinements it already makes, do not undo them: invoices dated before the practice cut-off (`--practice-since`, default 2025-09-01, when residential invoices stopped being billed to the parent builder) sit on the "Before the change" sheet and are not the headline; and a base-job invoice on its -FTW sibling customer (or the reverse) is only flagged when the memo's scope contradicts the side, otherwise it is a memo missing its -FTW and the customer is right.

3. Report, in this order, with job numbers and invoice numbers (never people's names):
   - Projects with NOTHING billed under them while invoices naming them sit elsewhere - these are the "clearly missing" ones. Give invoice #, date, amount, who it is billed to (the project's parent, another project, or an unrelated customer) and the project it should be on.
   - The rest of the off-project invoices grouped by division (RP / CP / MFD) and by what they are billed to. For MFD and CP say plainly that parent-billing with the job in the memo is common practice there, so those rows are a judgement call for the owner, not an error by definition.
   - Invoices whose memo names a job with no customer in QuickBooks (a missing project, not a mis-pointed invoice).
   - The totals: how many invoices, how much, how many projects.

4. Say which reports were wrong because of these: the project P&L, the RP WIP billed figure and the ledger all read invoices by customer, so an off-project invoice under-states that job's billed until it is re-pointed (the RP P&L alone also sweeps the parent's invoices by memo since 2026-09-08).

5. Do not offer to fix them in QuickBooks. Do offer to re-run after the owner has re-pointed them.
