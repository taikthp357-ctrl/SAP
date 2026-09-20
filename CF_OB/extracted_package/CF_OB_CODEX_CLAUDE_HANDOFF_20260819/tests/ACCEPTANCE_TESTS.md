# OFFLINE / REPLAY ACCEPTANCE TESTS

The refactor is not ready for live SAP testing until all core scenarios pass in REPLAY mode.

## Global invariants checked in every scenario

- No SAP writes in member/P planning loop.
- Plan is fully built and locked before first write.
- Normal `DOCUMENT_OPEN_COUNT=1`.
- `MON_SESSION_ID != PRDO_SESSION_ID`.
- No `MON_SESSION_REUSE ... SAME_TAB=TRUE`.
- After document bound, WH return uses WHTask, not OB reopen.
- A/C terminal WT never becomes Active again for same WT number.
- One PRDO Save per document when Qty targets change.
- Rerun never duplicates a satisfied Create WT allocation.
- Exact bin/batch allocation is preserved.

## Failure policy

A scenario may STOP only for a real hard-stop class: incomplete/ambiguous exact input, insufficient stock, actual SAP/adapter write failure, postverify failure, fresh state contradiction, or lost session context. Historical regression metadata must not act as a runtime blocker.

## Evidence to capture

For each scenario emit:

`Scenario, Result, GroupCount, MemberCount, DocumentOpenCount, MonOpenCount, WHTaskClicks, WTMapRefreshes, PRDOSaves, Creates, HardStops, Warns, FinalWTStates, Notes`.
