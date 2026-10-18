---
name: drawing-markup
description: Debuilt drawing colour-coding. Use whenever a construction drawing or floor plan PDF is attached or mentioned and the user wants anything coloured, color coded, highlighted, shaded, filled or marked up - apartments or units by sale status from a sales schedule Excel (sold, unsold, settled, reserved), a named room such as GYM, whatever a tag like ST2 points to, or balconies and other labelled areas by their m2. Also use for follow-ups on a drawing already being coloured (change colours, one aspect only, fill instead of highlight). Do not use for questions that need no colouring, for writing or editing drawing content, or for spreadsheets alone.
---

# Drawing markup

Colour-codes vector construction drawing PDFs by running the bundled script `scripts/drawing_markup.py`
(needs pymupdf, numpy, Pillow, openpyxl) in the code interpreter. Every edit comes from the script, which
refuses rather than guesses. Never write your own parsing, matching or drawing code, never estimate
coordinates, and never draw by hand to fix a result.

## 1. Load the script (every conversation)
```python
import glob, os, sys
roots = [os.getcwd(), "/mnt/data", "/mnt/user-data", "/workspace", "/home", "/tmp", "/"]
hits = []
for root in roots:
    hits = glob.glob(os.path.join(root, "**", "drawing_markup.py"), recursive=True)
    if hits:
        break
sys.path.insert(0, os.path.dirname(sorted(hits, key=lambda p: "drawing-markup" not in p)[0]))
import drawing_markup as dm
print(dm.__file__)
```
If nothing is found or the import fails, show the error and stop. Do not write replacement code, and never
ask the user to attach the script.

## 2. Setup (first message with a drawing)
```python
info = dm.inspect(pdf, xlsx)      # xlsx=None when no Excel is attached
print(info)
```
- `info["ok"]` is False (image, scan or flattened PDF): reply with `dm.SCAN_MESSAGE` and stop.
- Some pages in `info["scanned"]`: say those pages cannot be coloured.
- Use `info["statuses"]` for the counts. Never call `dm.status_of` on a whole row.
- `info["unit_labels"]` lists the sheets where unit labels were found. If it is empty with an Excel attached,
  say the unit numbers on the drawing do not match the Excel and show five unit numbers from each.

Then ask ONE question, e.g. "I found 31 sheets and 64 units (Unsold 35, Settled 23, Sold 6), with unit labels
on sheets 9-24. Do you want me to HIGHLIGHT each unit's label in its sale-status colour, or FILL each whole
apartment? Defaults: Sold blue, Unsold pink, Settled grey, Reserved orange. Other options: colour a room
('GYM green'), what a tag points to ('ST2 red'), areas by m2 ('balconies >8 blue, <8 orange'), one aspect
('north only')."

## 3. Run (when the user answers)
Always pass the ORIGINAL PDF. Never use extracted pages or an already coloured output (colours blend).

| User wants | Call |
|---|---|
| sale status, highlight labels (default, most reliable) | `log = dm.run_units(pdf, out, xlsx, check_dir="checks")` |
| sale status, fill whole apartments | `log = dm.run_units(pdf, out, xlsx, check_dir="checks", style="rooms")` |
| a room by its text ("GYM") | `log = dm.run_label(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| what a tag points to ("ST2") | `log = dm.run_leader(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| labelled areas by m2 | `log = dm.run_area_rule(pdf, out, page, "BALCONY", dm._parse_rules(">8:blue,<8:orange"), check_dir="checks")` |

Then always: `dm.write_log(log, "log.csv")`.

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

## 4. Report (always)
1. Every check PNG from `checks/`.
2. The edited PDF as a download.
3. A table: Page | Unit | Status | Result | Reason (from the log).
4. Counts per colour, and the NOT found / NOT coloured units with their reason. A unit whose label is not on
   any sheet in this PDF is "not on these pages", not an error. Commercial rows (Office, Retail) without a
   unit tag are expected.
5. Ask "Anything else to colour?". Files last only for this session: tell the user to download them.
