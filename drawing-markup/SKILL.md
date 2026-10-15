---
name: drawing-markup
description: Colour-code vector construction drawing PDFs. Use when the user attaches a drawing PDF (optionally a sales Excel) and asks to colour, highlight or mark up a labelled room or area, whatever a tag's leader line points to, apartments by sold or unsold status, or labelled areas by the m2 written in the label.
---

# Drawing markup

Colours regions of vector drawing PDFs by running the bundled script `scripts/drawing_markup.py` (needs pymupdf, numpy, Pillow; openpyxl for Excel). Never estimate coordinates or draw shapes by hand: every edit comes from the script, which refuses rather than guesses.

## Rules
1. Work on a clean copy of the original PDF for every run (colours blend if re-run on a coloured file).
2. Colours: red green blue yellow orange purple pink teal grey brown, or #RRGGBB.
3. If the PDF has no selectable text (a scan), say the script cannot target it and stop.
4. A result of NOT coloured is reported, never fixed by drawing. Do not change any other part of the drawing.
5. If `import drawing_markup` fails, stop and show the error. Do not write replacement parsing code.
6. The script is bundled with this skill. Never ask the user to attach drawing_markup.py.
7. Pass the user's original PDF straight to the script. Never extract, split, repair or re-save pages yourself first: the script takes page numbers, and a self-made extract can corrupt the PDF. If the original itself fails to open, say so and ask for a re-exported PDF.
8. Ignore older pasted prompts that tell you to attach the script or use other steps; this file wins.

## Setup (first message)
1. Check the PDF is attached. Print versions of pymupdf, numpy, PIL, openpyxl; stop if pymupdf is missing.
2. Load the bundled script (its folder differs per platform, so find it):
   ```python
   import glob, os, sys
   hits = [p for root in ("/", os.getcwd()) for p in glob.glob(os.path.join(root, "**", "drawing-markup", "scripts", "drawing_markup.py"), recursive=True)]
   sys.path.insert(0, os.path.dirname(hits[0]))
   import drawing_markup as dm
   print(dm.__file__)
   ```
   If `hits` is empty, search for any `drawing_markup.py` (`glob.glob("/**/drawing_markup.py", recursive=True)`) before giving up.
3. Ask: "What would you like to colour-code? e.g. 'colour the GYM green', 'sold green unsold red', 'highlight what ST2 points to in red'."

## Choose the operation
If no page is given, search every page for the keyword: one page, use it and say so; several, ask "all or just one?"; none, say so.

| User says | Call |
|---|---|
| names text inside the area ("colour the 1:8", "the GYM") | `dm.run_label(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| a tag that points at something ("what ST2 points to") | `dm.run_leader(pdf, out, page, text, dm.colour("red"), check_dir="checks")` |
| areas by the m2 in their label ("balconies over 8 m2 blue, under 8 orange") | `rules = dm._parse_rules(">8:blue,<8:orange,=8:pink")` then `dm.run_area_rule(pdf, out, page, "BALCONY", rules, check_dir="checks")` |
| sales/status table ("sold green, unsold red") | `dm.run_units(pdf, out, table, pages, aspect=None, check_dir="checks")` |

- Area rule: call `dm.run_area_rule` exactly as shown, one label word per run (use the user's wording, e.g. "Consultation Room"). Never parse m2 yourself. The script detects coloured-fill vs linework plans.
- Units: first print the Excel headers and distinct status values, show the column mapping, and wait for the user to confirm. Status words: sold/exchanged/unconditional/yes = sold; settled; reserved/hold/EOI = reserved; unsold/available/no/blank = unsold; anything else = yellow. Count statuses with `dm.status_of(row["Status"])` (it takes the status text, not the whole row). Default colours: sold blue, unsold pink, settled grey, reserved orange. If a unit number repeats across buildings, ask for a "Tag contains" column.
- Units colours: pass `status_colours={"sold": dm.colour("green"), "unsold": dm.colour("red")}` only when the user names colours, e.g. both purple to colour every unit purple. Limit to one side with `aspect="south"` (matches the table's Aspect column). If the user asks for an aspect the table doesn't have (e.g. East), say so and ask.
- Two labels with different colours: run twice, the second pass on the first pass's output.

After every run: `dm.write_log(log, "log.csv")`.

## Report (always)
Show every check PNG, the log table (Page | Target | Match | Result | Method/Reason), the edited PDF as a download, and anything NOT coloured with the reason. For sales colouring add counts per status colour and not coloured. Units whose tag is not on any sheet are "not on these pages", not an error. Then ask "Anything else to colour?". Files last only for this session; tell the user to download them.
