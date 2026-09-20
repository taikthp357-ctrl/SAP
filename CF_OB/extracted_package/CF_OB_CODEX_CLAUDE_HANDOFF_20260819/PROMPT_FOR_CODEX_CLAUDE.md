# PROMPT FOR CODEX / CLAUDE

You are refactoring an existing macro-enabled Excel SAP CF OB tool. **Do not rebuild it from scratch and do not run live SAP writes during offline work.**

## Required reading order
1. `markdown.md` — authoritative architecture/business contract.
2. `current/CURRENT_CF_OB_BASELINE_20260819.xlsm` — integration baseline.
3. `current_extracted/workbook_manifest.json`.
4. `current_extracted/sheets/CONFIG.csv`, `PLAN.csv`, `WH_ACTUAL.csv`, `RUN_LOG.csv`, `RUN_HISTORY.csv`, `REGRESSION_AUDIT.csv`.
5. `current_extracted/evidence/latest_run_excerpt.csv`.
6. `current_extracted/vba_symbols_inventory.txt`.
7. `legacy/LEGACY_REFERENCES.md`.
8. `tests/ACCEPTANCE_TESTS.md` and `tests/REPLAY_SCENARIOS.csv`.

## Mission
Refactor the current scheduler/executor into a deterministic **Group State Machine + RAM State Store + fixed MON/PRDO session roles + WHTask navigation**, while preserving proven SAP actions and workbook UI/shape/macro compatibility.

## Non-negotiable implementation requirements
- INPUT stays: OB list on the **left**, control buttons/status on the **right**.
- P/member loop is plan-only: READ/CLASSIFY/CALCULATE/PLAN; no SAP writes.
- Complete plan/group lock before first SAP write.
- Fixed phase order: Cancel → one WT-state update → Posting/Bin/Batch → grouped Qty → one PRDO Save/document → WHTask/OpenQty → Create WT allocations → WHTask verify → CF user handoff.
- Status A: blank/0 Actual = A_UNUSED; Actual>0 = recreate demand, retained qty 0.
- Status C: terminal/satisfied.
- Same WT number cannot transition A/C back to Active in the same run.
- Separate SAP roles: dedicated MON tab and dedicated PRDO/current-document tab. Never reuse PRDO as MON.
- Normal same-document `DOCUMENT_OPEN_COUNT=1`; after document is bound, use WHTask to return/read WH/WT.
- Qty target uses locked plan; after Save verify delta via WHTask Open Qty, not by repeated MON/OB reopen.
- Dynamic exact bin support; `APFL-CC01` must not be rejected merely because it is not P/ZF/HR.
- Reduce runtime hard stops to actual target/input/stock/write/postverify/fresh-state/session failures. Regression metadata is an offline build gate, not a live blocker.
- Implement replay mode and pass the supplied scenario suite before any live SAP test.
- Keep CF as user handoff by default.

## High-priority current functions to inspect/refactor
`CWR_SchedulerVerifyPRDOGroup`, `CWR_BatchSetPRDOQtyOneSave`, `CWR_BatchAllPRDOTargets`, `CWR_BatchOpenCreateWT`, `CWR_BatchAssertNoActiveOutboundWT`, `CWR_SchedulerUsePlan`, `CWR_ResetSchedulerExecutionContext`, `CWR_RequireRegressionGate`; add a single WHTask return helper such as `CWR_ReturnToWarehouseTask`.

## Current bug that your design MUST eliminate
The latest run proves old WT `5042889690` was cancelled and read back as `A`, then PRDO Qty 580→640 was saved/pass-verified, but a later Create step reused the PRDO tab as MON (`SAME_TAB=TRUE`), reopened/rebound context, and incorrectly blocked on the same old WT as active. This is a stale-state/navigation defect. Do not add another gate; eliminate the architecture that allows it.

## Work method
1. Make a copy of the current XLSM.
2. Export/read actual VBA source before editing. Preserve buttons/shape macro bindings.
3. Implement state store and group planner first.
4. Implement session-role/navigation invariants.
5. Move executor to fixed phases.
6. Simplify gates.
7. Add replay adapter and tests.
8. Only then clean UI/technical-sheet visibility.
9. Return updated XLSM + VBA diffs + changelog + offline test results.
10. Do not claim live pass until the staged company SAP tests are performed.
