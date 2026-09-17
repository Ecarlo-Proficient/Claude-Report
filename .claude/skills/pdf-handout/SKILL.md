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
   The deliverable is a local file - `~/Documents/CompanyHealth/<Name>.html` + `.pdf` - sent with
   SendUserFile. Skip the Artifact quickstart entirely.
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
3. **ALWAYS read the rendered PDF back before sending** - `Read` the `.pdf` with `pages`.
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
- Choice cards in a 3-column grid with `align-items:start`; the **snippet (button / memo
  crop) sits directly under that card's steps**, not pushed to the card bottom with
  `margin-top:auto`. A tip that belongs to a screenshot goes under the screenshot, not in the
  step text.
- Do not use `break-inside: avoid` on tall figures; it is what creates the blank page.
- Owner colour language on cards: green Approve, amber Revise, red Reject (top border + pill).

## Build and render

```bash
cd "$HOME/Documents/CompanyHealth" && "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" --headless=new --disable-gpu --no-pdf-header-footer --print-to-pdf="$HOME/Documents/CompanyHealth/<Name>.pdf" "file://$HOME/Documents/CompanyHealth/<Name>.html"
```

Then `Read` the PDF (pages 1-2), fix, re-render, read again. Send the PDF only (the HTML is
the source, mention it once).

## Companion record

A process handout means the process changed: update the vault registry
(`AI Brain_Vault/02_processes/`) and `log.md` in the same session, per the global upkeep rule.
