# drawing-markup

An agent skill that colour-codes vector construction-drawing PDFs (Revit/ArchiCAD/CAD exports) by running a tested Python script instead of letting the model guess geometry.

## What it does
| Operation | Example |
|---|---|
| `label` | colour the closed room/region around a text label ("1:8", "GYM") |
| `leader` | colour the region a tag's leader line points to ("ST2") |
| `units` | colour apartments by sale status (sold blue, unsold pink, settled grey, reserved orange) from an Excel or CSV |
| `area` | colour labelled areas by the m2 in the label (">8 blue, <8 orange") |

Every edit is drawn on its own PDF layer ("Debuilt markup"), comes with a zoomed check PNG and a log row, and anything ambiguous (open region, duplicate unit, area mismatch) is reported as NOT coloured instead of guessed.

## Layout
```
.
  drawing-markup/SKILL.md                 agent instructions (trigger + procedure)
  drawing-markup/scripts/drawing_markup.py  the tool (pymupdf, numpy, Pillow, openpyxl)
  Agent_Instructions.txt                  paste into the agent's instructions box
```

## Using it in Microsoft 365 Copilot / Copilot Studio
Full step-by-step guide and test plan: [docs/M365_Copilot_Setup.md](docs/M365_Copilot_Setup.md).

1. **Build the skill zip** (or download `drawing-markup.zip` from the latest [Release](../../releases)):
   `cd drawing-markup && zip -r ../drawing-markup.zip SKILL.md scripts`
   `SKILL.md` must sit at the top of the zip. GitHub's "Download ZIP" of the whole repo will not work.
2. **Create the agent** with the GitHub Copilot harness (skills need it) and turn on the code interpreter.
3. **Upload the skill**: agent > Skills > Upload > `drawing-markup.zip`.
4. **Paste `Agent_Instructions.txt`** into the agent's instructions. The instructions are always in context,
   while a skill only loads when its description matches the request, so the instructions are what make the
   agent load the skill for every drawing, including follow-ups.
5. **Check the runtime once**: ask the agent to run `import pymupdf, numpy, PIL, openpyxl; print(pymupdf.__version__)`.
6. **Test**: attach a drawing PDF and its sales Excel and say "colour code the units". In the agent's
   reasoning view, check that `drawing-markup` was loaded and that `dm.inspect` ran first.

Example requests: "colour the GYM green on page 8", "highlight what ST2 points to in red",
"balconies over 8 m2 blue, under 8 orange", "fill the apartments, sold green unsold red".

## Before uploading a new version
Run the real projects you have already checked and compare with their known results:
```
cp tools/cases.example.json tools/cases.json      # point the paths at your local drawings (kept out of git)
python tools/regress.py tools/cases.json
```
Each line prints PASS/FAIL with the units coloured and the ones that were not. Raise `min_coloured` when a
project gets better, so a later change cannot quietly make it worse.

## Command line
```
python drawing-markup/scripts/drawing_markup.py inspect in.pdf --table sales.xlsx
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
