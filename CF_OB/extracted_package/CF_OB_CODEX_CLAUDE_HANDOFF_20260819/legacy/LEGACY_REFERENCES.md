# LEGACY REFERENCES — HOW TO USE THEM

The older workbooks listed in `legacy_reference_manifest.csv` were discoverable as ChatGPT File Library references, but their binary bytes were not mounted in the active runtime. Therefore the ZIP **does not falsely claim to contain those XLSM binaries**.

## Trust hierarchy

1. `current/CURRENT_CF_OB_BASELINE_20260819.xlsm` — only editable binary baseline included in this package.
2. Known live-pass behavior from the 2026-08-14 CF baseline — preserve proven CF confirmation behavior.
3. V0.4.x MON core — preserve exact Document/WT context rules; do not reintroduce loose row association.
4. 2026-08-17/18 builds — mine them for group/planner ideas, but do not copy their navigation/session bugs.
5. Legacy technical modules — reuse concepts, not wholesale code, unless verified against the current workbook.

## Critical historical lessons

- V0.4.x established exact MON context: OB maps to Document; WT rows are accepted only for the exact Document; WHO may only carry within the same Document.
- Safe baselines introduced fail-closed exact SAP IDs and configuration locks. Those safeguards were useful, but the current build accumulated too many overlapping runtime gates. Keep exact-target safety, remove duplicated/contradictory gates.
- The 2026-08-14 baseline proved real CF completion (`Status=C`). Preserve that path.
- 2026-08-17/18 logs show the move toward grouped Cancel and one PRDO Save, but also show `MON_SESSION_NOT_FOUND`, repeated document openings, and stale-state problems. Do not copy the navigation structure.

## If old binaries become available

Add them under `legacy/binaries/` and compare modules/sheets against current. Do not make an older file the new baseline merely because one action was previously live-proven. The current workbook has the richest history and should remain the integration baseline unless a deliberate rollback is requested.
