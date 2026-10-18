# Debuilt Drawing Markup agent in Microsoft 365 Copilot

Goal: in the Microsoft 365 Copilot desktop app, a user opens the agent, drops in the drawing PDF and the
sales Excel, and gets the coloured PDF back. The script ships inside the skill; users never attach it.

Two files from this repo do the work:

| File | Where it goes | Why |
|---|---|---|
| `drawing-markup.zip` (`SKILL.md` + `scripts/drawing_markup.py`) | the agent's **Skills** | the procedure and the tested script, loaded on demand |
| `Agent_Instructions.txt` | the agent's **Instructions** box | always in context: forces the skill to load on every drawing message |

Build the zip from the repo root, or download it from the latest Release:
```
cd drawing-markup && zip -r ../drawing-markup.zip SKILL.md scripts
```
`SKILL.md` must be at the top of the zip (not inside a folder).

## Before you start: which route?
Skills in the Microsoft 365 Copilot **Agent Builder** are in preview, and at the time of writing only
available to tenants in Microsoft's **Frontier** program. Open Microsoft 365 Copilot > Agents > New agent >
Configure and look for a **Skills** section.

- **You see Skills** → Route A (simplest, everything inside the desktop app).
- **No Skills section** → Route B: build the agent in Copilot Studio and publish it to Microsoft 365
  Copilot. It then appears in the desktop app the same way.

## Route A: Agent Builder in the Microsoft 365 Copilot app
1. Microsoft 365 Copilot app > **Agents** > **New agent** > switch to **Configure**.
2. **Name**: `Debuilt Drawing Markup`. **Description**: `Colour-codes construction drawing PDFs by sale status from the sales Excel.`
3. **Instructions**: paste all of `Agent_Instructions.txt`.
4. **Code interpreter**: nothing to switch on. Agent Builder now always includes it for every agent, so there is no "Create documents, charts, and code" toggle any more. The runtime test below confirms it works.
5. **Skills** > **Add** > upload `drawing-markup.zip`. Check it shows as `drawing-markup`.
6. **Work content**: turn off Cloud files, Outlook, Teams and Web search, and set the skill to **Limited use** (... menu on the skill). The agent only needs the files users attach.
7. **Conversation starters** (so users only drop files):
   - `Colour code these units by sale status` (attach PDF + Excel)
   - `Fill each whole apartment by sale status`
   - `Colour the GYM green`
8. Leave **Knowledge** empty: do not upload the script or the drawings there.
9. **Try it** (see the test plan below), then **Create** and share the agent with the team.

## Route B: Copilot Studio, published to Microsoft 365 Copilot
1. copilotstudio.microsoft.com > **Create** > **New agent**, using the **GitHub Copilot harness** (skills need it).
2. **Instructions**: paste `Agent_Instructions.txt`.
3. **Settings** > turn on the **code interpreter**.
4. **Skills** tab > **Upload** > `drawing-markup.zip`.
5. Test in the test pane (plan below). In the reasoning view, confirm `drawing-markup` was loaded and `dm.inspect` ran first.
6. **Channels** > **Microsoft 365 Copilot** (Teams and Microsoft 365) > **Publish**, then make the agent
   available to the team (or ask your admin to approve it in the Microsoft 365 admin centre).
7. Users find it in the Microsoft 365 Copilot desktop app under **Agents**.

## Test plan (run before sharing, and after every new zip)
| Test | Send | Expect |
|---|---|---|
| Runtime | `run: import pymupdf, numpy, PIL, openpyxl; print(pymupdf.__version__)` | a version number. If it fails, the platform cannot run the skill |
| Trigger, no words | Prospect St PDF + Excel, no text | no question asked; "31 sheets, 64 units ... colouring now"; 60 units coloured (32 pink, 22 grey, 6 blue); 1701, 1702, Office, Retail NOT found |
| Trigger, casual words | same files + "can u do this" | the same as above, the skill loaded |
| Fill | Franklin PDF + Excel + "fill the apartments" | whole apartments filled, corridors and lifts untouched |
| Scan | Alphington compressed PDF | "This file has no vector linework or text the script can read. Please send the PDF exported from CAD." and nothing else |
| Follow-up | "make sold green" after a run | the skill reloads; it reruns on the ORIGINAL PDF |
| Off-topic | "write me an email" | "this agent only colour-codes drawings" |

If the skill does not load in a test, look at the reasoning view: the agent's words should match the
skill description (colour, color code, highlight, fill, mark up, sale status). Add the user's wording to the
description in `SKILL.md`, rebuild the zip and upload it again.

Before uploading a changed script, run the real projects locally: `python tools/regress.py tools/cases.json`.

## Limits to know
- Vector PDFs exported from CAD/Revit only. Scans, "compressed" or flattened PDFs (iLovePDF etc.) are refused.
- Very large PDFs: platform upload limits apply. Send only the floor-plan sheets if a file is rejected.
- Output files last for the chat session only; users must download them.

## Troubleshooting
**"Upload and replace failed ... We couldn't replace this skill"**
The zip is valid (SKILL.md at the top, only .md/.py files, depth 2, under 100 KB). The usual causes:
1. **The agent also has uploaded files.** Microsoft's known preview issue: an agent cannot have both skills and
   embedded files. Configure > Knowledge: remove every uploaded file (old `drawing_markup.py`, PDFs, Excels).
2. **Replacing a skill with a different name.** Replace only swaps a skill of the same name. Delete the old
   skill (Configure > Skills > ... > Delete), save, then **Add** the new zip instead of replacing.
3. **A stale draft.** Save or discard pending changes, reopen the agent, then add the skill again. If it still
   fails, create a fresh agent (Route A steps) and add the zip there.

**The agent runs but says a module is missing.** Skill scripts run in a sandbox that cannot install packages.
`pymupdf`, `numpy`, `Pillow` and `openpyxl` must already be there: run the Runtime test above. If `pymupdf` is
missing in your tenant, the skill cannot work there; tell us the exact error.
