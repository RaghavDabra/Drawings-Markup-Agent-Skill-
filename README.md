# drawing-markup

An agent skill that colour-codes vector construction-drawing PDFs (Revit/ArchiCAD/CAD exports) by running a tested Python script instead of letting the model guess geometry.

## What it does
| Operation | Example |
|---|---|
| `label` | colour the closed room/region around a text label ("1:8", "GYM") |
| `leader` | colour the region a tag's leader line points to ("ST2") |
| `units` | colour apartments sold green / unsold red from an Excel or CSV |
| `area` | colour labelled areas by the m2 in the label (">8 blue, <8 orange") |

Every edit is drawn on its own PDF layer ("Debuilt markup"), comes with a zoomed check PNG and a log row, and anything ambiguous (open region, duplicate unit, area mismatch) is reported as NOT coloured instead of guessed.

## Layout
```
.
  drawing-markup/SKILL.md                 agent instructions (trigger + procedure)
  drawing-markup/scripts/drawing_markup.py  the tool (pymupdf, numpy, Pillow, openpyxl)
  Agent_Instructions.txt                  paste into the agent's instructions box
```

## Using it as an agent skill
1. **Get the skill zip.** Download `drawing-markup.zip` from the latest [Release](../../releases), or build it yourself:
   `cd drawing-markup && zip -r ../drawing-markup.zip SKILL.md scripts`
   (Don't upload GitHub's "Download ZIP" of the whole repo: the skill must have `SKILL.md` at the top.)
2. **Add it to your agent.** In your agent builder's skills section (e.g. Copilot Studio, Claude), upload `drawing-markup.zip`.
3. **Paste the agent instructions.** Copy `Agent_Instructions.txt` into the agent's instructions box. It tells the agent to always use the skill and never draw by hand.
4. **Check the runtime once.** Ask the agent to run `import pymupdf, numpy, PIL; print(pymupdf.__version__)`. If that fails, the platform can't run the script (see Limits).
5. **Use it.** Attach a drawing PDF (plus a sales Excel for sold/unsold) and say what to colour, e.g.:
   - "colour the GYM green on page 8"
   - "highlight what ST2 points to in red"
   - "balconies over 8 m2 blue, under 8 orange"
   - "sold green, unsold red" (with the Excel attached)

   The agent returns the edited PDF, a check image per edit and a log of anything it refused to colour.

Skill packaging differs between platforms. This layout follows the open Agent Skills format (`SKILL.md` with `name`/`description` frontmatter plus `scripts/`). If your platform rejects it, open an issue with the error.

## Command line
```
python drawing-markup/scripts/drawing_markup.py label  in.pdf out.pdf --page 8  --text "1:8" --colour red
python drawing-markup/scripts/drawing_markup.py leader in.pdf out.pdf --page 17 --text ST2 --colour red
python drawing-markup/scripts/drawing_markup.py units  in.pdf out.pdf --table sales.xlsx --pages 10-21
python drawing-markup/scripts/drawing_markup.py area   in.pdf out.pdf --page 9 --text "Consultation Room" --rules "==15:purple,==14:red"
```

## Limits
Vector PDFs only (scans are refused). The agent platform must be able to run Python with pymupdf, numpy and Pillow installed. Rooms open to each other cannot be isolated and are reported. Colours blend if you re-run on an already coloured PDF, so always start from the original.

## Install
`pip install -r requirements.txt`. To use as an agent skill, zip the `drawing-markup/` folder (SKILL.md + scripts/) and add it to your agent; paste `Agent_Instructions.txt` into the agent's instructions.

## Area rules
`--rules` takes comma-separated tests on the m2 written in each label; the first match wins.
Operators: `>`, `<`, `>=`, `<=`, `=`. Example: `">8:blue,<8:orange,=8:pink"`.
Labels with no m2, or rooms that are open to a neighbouring space, are reported as NOT coloured.

## Tests
```
pip install -r requirements.txt pytest pyyaml
pytest -q                 # everything
pytest -q -m "not slow"   # skip the end-to-end area test
```

## Licence
MIT, see LICENSE
