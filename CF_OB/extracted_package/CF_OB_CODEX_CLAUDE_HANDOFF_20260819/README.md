# CF OB Codex/Claude Handoff Package — 2026-08-19

Start with **`markdown.md`**. It is the full architecture/business/refactor specification.

## Included

- `current/CURRENT_CF_OB_BASELINE_20260819.xlsm` — current workbook binary, SHA256 `9ebcc46fcdf25a10803792d2ed576c79f98a394eec838a783d5b0ee021905309`.
- `current_extracted/sheets/` — all 20 sheet snapshots as CSV.
- `current_extracted/evidence/latest_run_excerpt.csv` — latest run evidence.
- `current_extracted/vba_symbols_inventory.txt` and `vba_strings_full.txt` — VBA binary-readable clues.
- `support/` — older SAP automation notes.
- `legacy/` — old CF OB version manifest and lessons.
- `tests/` — offline/replay acceptance suite.
- `PROMPT_FOR_CODEX_CLAUDE.md` — ready-to-use implementation prompt.

## Important limitation

Older XLSM files found in ChatGPT File Library were references only and were not mounted as binary bytes in this runtime, so they are **not physically included**. Their exact names/file IDs and role are recorded in `legacy/legacy_reference_manifest.csv`; do not treat this ZIP as containing those old binaries.
