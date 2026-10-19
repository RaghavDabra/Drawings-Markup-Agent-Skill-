---
name: drawing-markup
description: Debuilt drawing colour-coding. Use whenever a construction drawing or floor plan PDF is attached or mentioned and the user wants anything coloured, color coded, highlighted, shaded, filled or marked up - apartments or units by sale status from a sales schedule Excel (sold, unsold, settled, reserved), a named room such as GYM, whatever a tag like ST2 points to, or balconies and other labelled areas by their m2. Also use for follow-ups on a drawing already being coloured (change colours, one aspect only, fill instead of highlight). Do not use for questions that need no colouring, for writing or editing drawing content, or for spreadsheets alone.
---

# Drawing markup

Colour-codes vector construction drawing PDFs by running the bundled script `scripts/drawing_markup.py`
(needs pymupdf, numpy, Pillow, openpyxl) in the code interpreter. Every edit comes from the script, which
refuses rather than guesses. Never write your own parsing, matching or drawing code, never estimate
coordinates, and never draw by hand to fix a result.

## 1. Sale-status colouring: ONE code run (the default)
For a drawing + sales Excel (no other words, "colour code", "go", "fill the apartments", colours, an aspect),
run this single block and nothing before it. Do not list or search for files first: the script finds the
attachments itself, including SEVERAL drawing PDFs (e.g. one file per floor) - never loop over files yourself.
```python
import glob, os, sys
hits = glob.glob("/home/oai/skills/**/drawing_markup.py", recursive=True) or glob.glob("/**/drawing_markup.py", recursive=True)
sys.path.insert(0, os.path.dirname(hits[0])); import drawing_markup as dm
r = dm.colour_by_status()          # style="rooms" to fill whole apartments; status_colours=..., aspect="north"
print(r)
```
- `r["ok"]` is False: reply with `r["message"]` only, and stop. `r["ok"]` True with a `r["message"]`: some PDFs
  were scans and were skipped; say so in one line.
- Otherwise report from `r` (section 4): `r["out"]` is the PDF, `r["log_csv"]` the log, `r["check_dir"]` the
  check images, `r["coloured"]` the counts per colour, `r["not_found"]` the units not coloured grouped by reason.
  Several PDFs: there is no `r["out"]`; offer `r["download_zip"]` (every coloured PDF, log and the checks) and
  list `r["files"]` (each PDF with its pages and coloured output).
- Pass `status_colours={"sold": dm.colour("green"), "unsold": dm.colour("red")}` ONLY when the user names
  colours; statuses left out of it stay uncoloured. Follow-ups ("make sold green", "fill instead") rerun the same
  block with the new arguments: it always starts from the ORIGINAL attached PDF.
- Give the code run the longest timeout allowed. If it still times out, do NOT list files or rerun the same
  thing: run the block once more with `dm.colour_by_status(checks=0)` (no check images, fastest). If that also
  times out, tell the user the drawing set is too large for one run and to send only the floor-plan sheets.
- All check images (one per unit) only when the user asks: `dm.colour_by_status(checks=None)`.
- If the import fails, show the exact error and stop. Never write replacement code, never ask the user to attach
  the script.

## 2. Anything else (no Excel, or a room / tag / area request)
Load the script with the first three lines above, then `info = dm.inspect(pdf)` (paths from
`pdf, xlsx = dm.find_inputs()`). `info["ok"]` False: reply with `dm.SCAN_MESSAGE` and stop. Only a drawing and
no request: ask ONE question, e.g. "What should I colour? e.g. 'GYM green', 'what ST2 points to in red',
'balconies >8 m2 blue, <8 orange', or attach the sales Excel to colour units by sale status."

## 3. Other operations
Always pass the ORIGINAL PDF. Never use extracted pages or an already coloured output (colours blend).

| User wants | Call |
|---|---|
| a room by its text ("GYM") | `log = dm.run_label(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| what a tag points to ("ST2") | `log = dm.run_leader(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| labelled areas by m2 | `log = dm.run_area_rule(pdf, out, page, "BALCONY", dm._parse_rules(">8:blue,<8:orange"), check_dir="checks")` |

Then always: `dm.write_log(log, "log.csv")`. (Sale-status colouring uses section 1 instead.)

- Unit labels the script reads: a number beside its area (`1.03 78.4 m2`), `UNIT 4`, `APARTMENT T1.0501`,
  `UNIT 2U.01` over `507`, and typical-floor tags listing one number per floor (`UNIT 2A.01` over `909 1009 1109`).
- `style="rooms"` fills each whole apartment through its doors and also highlights its label; corridors,
  lobbies, lifts and stairs are never filled. Units it cannot close are logged NOT coloured; offer the
  label highlight for those. `style="fill"` is only for plans that draw each unit as one coloured shape.
- Colours: red green blue yellow orange purple pink teal grey brown, or #RRGGBB. Pass
  `status_colours={"sold": dm.colour("green"), "unsold": dm.colour("red")}` ONLY when the user names colours;
  statuses left out of it stay uncoloured. One side only: `aspect="north"` (the Excel's Aspect column).
- No page given for label/leader/area: search every page for the text. One page, use it and say so; several,
  ask "all or just one?"; none, say so.
- Two operations in different colours: run the second on the first one's output.

## 4. Report (always, keep it short)
1. Offer as downloads: the coloured PDF (`r["out"]`, or `r["download_zip"]` for several PDFs), the check images zip (`r["checks_zip"]`, 3 images by default; offer all of them) and the log
   (`r["log_csv"]`). Show at most the 3 images in `r["preview"]`; never display every check image.
2. Counts per colour from `r["coloured"]`, e.g. "Coloured 60: 32 pink (unsold), 22 grey (settled), 6 blue (sold)".
3. Not coloured, one line per reason from `r["not_found"]`: "level not in this drawing set" (normal when the PDF
   has only some floors) and "no unit tag on its sheet". Never list every unit row in the chat; the log has them.
4. After a default run, offer: "Want me to FILL each whole apartment instead, or change colours (defaults: Sold
   blue, Unsold pink, Settled grey, Reserved orange)?" Files last only for this session: tell the user to
   download them.
