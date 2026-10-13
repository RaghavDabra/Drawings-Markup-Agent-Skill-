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

## Command line
```
python drawing-markup/scripts/drawing_markup.py label  in.pdf out.pdf --page 8  --text "1:8" --colour red
python drawing-markup/scripts/drawing_markup.py leader in.pdf out.pdf --page 17 --text ST2 --colour red
python drawing-markup/scripts/drawing_markup.py units  in.pdf out.pdf --table sales.xlsx --pages 10-21
python drawing-markup/scripts/drawing_markup.py area   in.pdf out.pdf --page 9 --text "Consultation Room" --rules "==15:purple,==14:red"
```

## Limits
Vector PDFs only (scans are refused). Rooms open to each other cannot be isolated and are reported. Colours blend if you re-run on an already coloured PDF, so always start from the original.

## Install
`pip install pymupdf numpy pillow openpyxl`. To use as an agent skill, zip the `drawing-markup/` folder (SKILL.md + scripts/) and add it to your agent; paste `Agent_Instructions.txt` into the agent's instructions.

## Licence
MIT, see LICENSE.
