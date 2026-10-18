# Changelog

## 1.1.0
- Readable errors for unknown colour names and malformed hex colours.
- CLI reports missing arguments for each operation before running.
- Coloured-fill detection reuses the cached drawing list.
- Unit and end-to-end tests, run on GitHub Actions.

## 1.0.0
- Skill instructions, agent instructions and `drawing_markup.py` with four operations:
  label, leader, units (sold/unsold from Excel) and area (colour by labelled m2).

## 1.2.0
- Unit tags with a type code and no area (`UNIT 2U.01` over `507`) and stacked typical-floor tags (`909 1009 1109`).
- `inspect()` setup check: page count, scanned pages refused, status counts, sheets with unit labels.
- SKILL.md and Agent_Instructions.txt rewritten for Copilot Studio skills, so the skill loads on every drawing request.
- `tools/regress.py` to re-check real projects before each upload.
