# Outbound-WT execution learnings — 2026-08-17

## Confirmed outcomes

- A stale MON result grid can lack the locked technical column even while the map is valid. Startup must establish an idle session only; the loader must filter the requested OB and validate the fresh result grid.
- A verified PRDO Save can leave the executor in PRDO/S000 with no MON tab. Re-establishing a bounded, map-locked MON reader before the next readback is safe; it is not a retry of the preceding SAP write.
- The generic SAP session helper intentionally forbids creating a new tab in RUN mode. MON recovery must use a separate, hard-coded `/SCWM/MON` reader-only exception; it must never broaden RUN-mode tab creation for other TCodes.
- Do not create a new SAP tab for MON recovery in RUN. After a verified PRDO Save, move the one exact idle `/SCWM/PRDO` executor tab to `/SCWM/MON` for readback; later plan lines must reuse that MON tab. Missing or ambiguous sessions stop safely without creating blank tabs.
- A pre-Adopt/pre-Create stop with `SAP_WRITE=0` is a retryable blocked state, not `STATE_UNKNOWN`. Only an Adopt or Create+Save already sent with an unproven outcome is terminal/no-retry.
- `CF_ONLY` plus `READY_FOR_CF` means PLAN has no permitted SAP work left. Button #4 must return the explicit manual-CF handoff, not report a missing executable plan as an error.
- A missing VBA label is a compile blocker, not a runtime/SAP condition. Every staged executor import must run a local `GoTo`-to-label audit before the workbook is touched.
- A procedure named `CompileProbe` must be pure: it cannot read the mutable PLAN, SAP session, or current runtime context. State-sensitive diagnostics must have a separate name.
- In Create-WT's Warehouse Request grid, one request row activates the dependent `Stk Can Be Removed` panel with one selection; when the grid has multiple rows, double-click the exact matched upper row before reading or editing source quantity.
- The PLAN/WH lock hash must cover MON source fields and user-owned picking inputs only. `WH_ACTUAL` status/result columns (`O:S`, `Y:AD`) are executor-owned and may change after a completed OB; hashing them falsely blocked the next OB in the same WH snapshot with `WH_ACTUAL_CHANGED_AFTER_CHECK`.

## Repeatable operator path

1. After any verified SAP write or a stopped run, use `TẢI WH -> CHECK -> POSTING` to produce a fresh locked PLAN.
2. Run Button #4 only for `READY_TO_EXECUTE` rows. Rows at `READY_FOR_CF` are handed to the user for CF; the tool never confirms CF.
3. If a write outcome is unknown, do not repeat the write. Read MON and reconcile first.
4. After a completed OB, the next OB in the same loaded snapshot may continue under the same CHECK lock when source/user-input fields are unchanged. If quantity/bin/batch input or the MON source itself changed, run `TAI WH -> CHECK -> POSTING` again; the stale stop remains intentional.

## Regression expectations

- No stale-grid validation before the fresh MON OB filter.
- MON reader recovery is available to the executor after PRDO/S000 navigation.
- The MON recovery path can create only a `/SCWM/MON` reader tab in RUN mode; all other new-session requests remain blocked.
- MON post-PRDO recovery reuses the same tab and never creates a blank SAP session; execution compile probes are context-free.
- Multi-request activation is logged as `REQUEST_ROW_ACTIVATED` with `DOUBLE_CLICK=TRUE` only when `REQUEST_ROWS>1`; source-row matching must occur only after that activation.
- A no-write precondition failure returns `BLOCKED`, while sent Adopt/Create+Save with unknown result returns no-retry/unknown.
- All `READY_FOR_CF` rows yield `READY_FOR_CF_MANUAL_HANDOFF` with `CF_NOT_SENT`.
- The staged MON and executor modules pass the undefined-label audit before import.
