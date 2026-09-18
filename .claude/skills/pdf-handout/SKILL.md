---
name: pdf-handout
description: >-
  Build a one-page PDF handout / flow chart / how-to for the team (PMs, clerks,
  supers) from the owner's own screenshots - a process change, a "here is the
  new way" sheet, a guide to a QuickBooks or JobTread screen. Use whenever the
  owner asks for a flow chart, one-pager, handout, guide, cheat sheet or "make
  this simple for the PMs", or pastes screenshots and says "use the sc". Covers
  finding the pasted screenshots on disk, cropping them, the print layout, the
  mandatory read-back of the rendered PDF, and where the file lives. Never a
  hosted artifact.
---

# PDF handout (the owner's rules, learned 2026-09-17 on the bill-approval guide)

## Hard rules

1. **Never publish a hosted Artifact page.** The owner: shared claude.ai pages get Google-indexed.
   The deliverable is a local file sent with SendUserFile. Skip the Artifact quickstart entirely.
   **A PROCESS guide lives in the vault: `AI Brain_Vault/assets/processes/<PROCESS-ID>_<slug>.html`
   + `.pdf`** (owner 2026-09-18: "processes should be like this, not the weird flow charts ... add
   to assets/processes and include these in ledger"). The ID prefix is the link - the ledger's
   Systems tab shows a **Guide** pill under that row's ID and serves the file live
   (`/api/process-guide`), so naming it right IS publishing it. Add the row to that folder's
   README table. A handout that is not a registered process goes to `~/Documents/CompanyHealth/`.
2. **Use the owner's real screenshots. Never redraw a screen you were given.** A mock is only
   acceptable when no capture exists, and it must say so. Pasted images are NOT attached as files
   in chat, but the screenshot utility keeps them on disk - find them by time and size:

```bash
cd "$HOME/Library/Caches/com.vorssaint.utils/Copied Screenshots" && for f in Screenshot\ $(date +%Y-%m-%d)*.png; do printf '%s  ' "$f"; sips -g pixelWidth -g pixelHeight "$f" | awk '/pixel/{printf "%s ", $2}'; echo; done
```

   Then `Read` the candidates to confirm which is which (an unrelated capture from the same
   minute is normal - drop it), copy them into the scratchpad, and crop with Pillow
   (`from PIL import Image; Image.open(f).crop((x0,y0,x1,y1))`). The Read tool shows a scaled
   image and prints the scale factor - multiply displayed coordinates by it before cropping.
   Downscale to <= 1900 px wide and embed as `data:image/png;base64` so the HTML is one file.
   If a screenshot needs a line the capture lacks (e.g. the REVISED memo line), type it ONTO
   the crop with `ImageDraw` in a matching font and say so in the reply.
3. **ALWAYS read the rendered PDF back before sending, EVERY iteration** - count pages with
   a regex over the PDF bytes first (`/Type /Page` minus `/Pages`), then `Read` page 1. In the
   bill-approval round this loop ran ~15 times; every single render that skipped the read had a
   defect (blank page, clipped image, arrowhead on top of text). Read the rendered PDF back before sending - `Read` the `.pdf` with `pages`.
   Check the page count and look at every page. A tall image plus a keep-together rule pushed
   one screenshot to page 2 and left page 1 blank; nobody caught it because nobody looked.
   Shipping unreviewed output is the defect the owner called out ("can you not view pdf???").
4. **Straight to the point.** No preamble ("no more paper on your desk"), no rules list, no
   worked example when the screenshot already shows it, no closing offers. Title, effective
   date, the flow, where to click, the choices. One Letter page.
5. **No personal names in anything I write.** Names visible inside the owner's own screenshots
   are his call - leave the crop alone, mention it once.

## Layout that fits one page (Chrome headless, Letter portrait)

- `@page { size: Letter portrait; margin: 0.4in 0.5in }`, body 13px system sans, `.page` max
  7.5in. No web fonts needed for print.
- **"Where to click" is a numbered LIST, stacked** - caption, then its screenshot under it,
  next item. The owner rejected a side-by-side grid ("idk why you are putting it to the right").
  Keep the page by cropping, not by re-flowing: give each image a fixed print width
  (`width: 3.4in` for a bar, `2.3in` for a panel).
- **Crop tight to the wording.** Cut a screenshot down to the part that carries the message:
  the icon cluster with the arrow, the memo label plus its two lines, a panel trimmed to two
  cards plus its footer. Empty canvas around the words is what costs the page.
- **Simple choices side by side, the branching choice as its own full-width block underneath.**
  Approve and Reject are two half-width cards; Revise (the one with a decision in it) is a
  wide card below them laid out left to right: steps and question | the decision boxes |
  the screenshot. Never leave a tall card next to short ones - the empty space under the
  short cards is what the owner circles in red.
- **A decision is drawn as a flow, not two boxes in a row.** From the question, a real LINE
  with an arrowhead to each outcome, in the outcome's colour (orange to "Yes, total changed",
  green down to "No, same total"), and the line stops SHORT of the box - an arrowhead touching
  or overlapping a border reads as a defect. Draw it with a flex row: question text, then a
  `flex:1 0 26px` 2px bar with a CSS-triangle `::after`, `margin-right:10px` for the gap; the
  question span must be allowed to shrink (`flex:0 1 auto`) or the bar overflows into the
  next column and lands on top of the box text.
- **Everything that belongs to an outcome lives INSIDE that outcome's box.** The tip sits in
  the orange box beside the Yes text behind a vertical divider, and the memo screenshot sits in
  the same orange box under a horizontal divider - the box says "this is all one path".
- **The button screenshot sits bare, directly under step 3 of its own card** - no "Bottom of
  the bill" box around it, no label, no row at the foot of the page. The owner's words were
  "put the sc of the buttons below the 3" and I read "the 3" as the three cards instead of
  step 3; **when an instruction names a number, it is the numbered step, not a count of
  blocks.** Keep the simple cards short so the decision block gets the vertical room.
- **Yes and No boxes share one type size, a step above the page body** (headers 15px, text 13px)
  so both outcomes catch the eye equally; outcome text gets breathing room (`line-height:1.5`).
- **No dead gap inside a box.** A text column beside a divider is `max-content`, not a `fr`
  share, so the divider and the tip sit right after the text. The owner checks for gaps -
  look for them in the rendered PDF before sending.
- QBO memo new line is **Shift+Enter** (owner 2026-09-17), not Ctrl+Enter.
- QBO **Reject opens "Add comments to the rejection - this will be shared via email"**: the reject
  reason goes there (it emails the bill clerk), never in the memo. When a product behaviour is
  unconfirmed, say so in one line and ask the owner to click it - he had the answer in a minute;
  two web searches did not.
- Do not use `break-inside: avoid` on tall figures; it is what creates the blank page.
- **Fitting one page is done by trimming the where-to-click screenshots first** (bar 2.5in,
  panel 1.25in worked), never by shrinking the decision block - that block is the point.
- No "Effective <date>" line, no subtitle. Title, then the flow.
- Owner colour language: green Approve, red Reject, orange Revise (top border + pill); in the
  flow strip the AP boxes' labels are orange, QuickBooks green, You blue. Start the flow at AP,
  not the vendor.

## Build and render

```bash
cd "$HOME/Documents/Claude/AI Brain_Vault/assets/processes" && "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf="$PWD/<ID>_<slug>.pdf" "file://$PWD/<ID>_<slug>.html"
```

Then `Read` the PDF (pages 1-2), fix, re-render, read again. Send the PDF only (the HTML is
the source, mention it once).

## Companion record

A process handout means the process changed: update the vault registry
(`AI Brain_Vault/02_processes/`) and `log.md` in the same session, per the global upkeep rule.
**Keep the registry row a verb phrase** - the detail goes in the guide and the domain file's
chain section; a row that grows into a paragraph buries the Guide pill and the columns. And
**never put a `|` inside a registry table cell** (a `[[note|alias]]` wikilink did it once and
shifted every column) - re-parse with `python3 ledger/registry_view.py` after any row edit.
This one-page guide IS the standard for documenting a process; do not draw mermaid or ASCII
flow charts for a process again.
