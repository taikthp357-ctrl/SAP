# SAP CF OB VBA Tool - Architectural Analysis

**Analysis Date:** 2026-08-19  
**Analyzed By:** Exploration Agent  
**Source:** `current/CURRENT_CF_OB_BASELINE_20260819.xlsm` + extracted artifacts  
**Mission:** Deep architecture exploration of CURRENT implementation to understand the existing model before refactor

---

## EXECUTIVE SUMMARY

The current CF OB tool is a **complex, per-member navigation model** that has accumulated significant technical debt through repeated gate additions and navigation workarounds. The core bug is **stale state + session reuse**, not missing validation.

### Critical Finding: The Root Cause Bug

Latest run evidence (RunID `20260819_054040`) proves:
1. Old WT `5042889690` cancelled successfully → status `A` verified
2. PRDO Qty change 580→640 saved and post-verified PASS
3. Later, code **reused verified PRDO tab for MON** (`SAME_TAB=TRUE`)
4. Create phase then blocked because same old WT appeared "active" despite proven terminal state
5. Navigation counters: `DOCUMENT_OPEN_COUNT=3`, `WHTASK_OPEN_COUNT=0`

**This is architectural**: violated session separation + no WHTask navigation + stale cache precedence over RAM proof.

---

## 1. MODULE INVENTORY

Based on `vba_symbols_inventory.txt` and strings analysis, the VBA project contains these major areas:

### Core Execution Modules
- **`modCodexWTRework`** — Primary scheduler/executor module containing group execution logic
- VBA functions identified: ~100+ procedures with `CWR_*` prefix
- Integration with SAP GUI Scripting API

### Key Procedure Categories

**Scheduler/Planning:**
- `CWR_SchedulerUsePlan` — Loads execution plan and binds context
- `CWR_SchedulerVerifyPRDOGroup` — Validates PRDO document group before qty save
- `CWR_SchedulerVerifyQtySources` — Verifies source availability
- `CWR_SchedulerBuildPRDOBatch` — Builds PRDO batch for grouped qty changes
- `CWR_SchedulerCreateQty` — Calculates create quantities
- `CWR_SchedulerState` — Manages scheduler state transitions
- `CWR_ResetSchedulerExecutionContext` — Resets context (called too frequently)

**Group Execution:**
- `CWR_ExecuteLockedPlanSchedule` — Main locked plan executor
- `CWR_ExecutePlanToCFHandoffCore` — Core execution to CF handoff
- `CWR_ExecutePreflightGroups` — Pre-execution validation
- `CWR_ExecuteDocumentGroups` — Document-level group execution
- `CWR_ExecuteLocationGroups` — Location-based grouping
- `CWR_ExecuteQtyGroups` — Quantity change groups
- `CWR_ExecuteCreateGroups` — Create WT groups
- `CWR_BuildExecutionGroups` — Groups members by execution phase

**Batch Operations (Current Per-Member Pattern):**
- `CWR_BatchSetPRDOQtyOneSave` — Sets PRDO qty with one save per document
- `CWR_BatchAllPRDOTargets` — Verifies all PRDO targets
- `CWR_BatchOpenCreateWT` — Opens Create WT from PRDO context
- `CWR_BatchAssertNoActiveOutboundWT` — Checks for active WTs (where bug occurs)
- `CWR_BatchMapPRDORows` — Maps PRDO grid rows
- `CWR_BatchCollectActiveQtyRecords` — Collects active qty records
- `CWR_BatchPostVerifyAndRecord` — Post-verification recording
- `CWR_BatchNoSAPError` — SAP error checking

**PRDO (Process Delivery Order) Operations:**
- `CWR_OpenExpectedPRDO` — Opens PRDO document
- `CWR_FindCurrentPRDOGrid` — Finds PRDO grid in SAP
- `CWR_FindExactPRDOItem` — Locates exact PRDO line item
- `CWR_SelectExactPRDOItem` — Selects PRDO item for editing
- `CWR_ReadExactPRDOItem` — Reads PRDO item data
- `CWR_RecordPRDOQtyChange` — Records qty change to state
- `CWR_EnsurePRDOQtyAtTarget` — Ensures PRDO qty reaches target
- `CWR_EnsurePRDOQtyAtTargetOnce` — One-time PRDO qty enforcement
- `CWR_ExecutePRDOQtyToTarget` — Executes PRDO qty change
- `CWR_ResetPRDOWithoutChanges` — Resets PRDO without saving
- `CWR_GetVerifiedPRDOChangeAction` — Gets verified PRDO change action
- `CWR_GetVerifiedPRDOSaveAction` — Gets verified PRDO save action
- `CWR_AssertPRDOQtyVerifiedLocal` — Asserts PRDO qty local verification

**MON (Monitor) Operations:**
- `CWR_OpenMonitorForRead` — Opens MON for read-only access
- `CWR_GetMONSession` — Gets MON session reference
- `CWR_ReuseVerifiedPRDOTabForMON` — **THE BUG**: Reuses PRDO tab for MON
- `CWR_ReserveMONOpen` — Reserves MON session for opening
- `CWR_ReserveMONRefresh` — Reserves MON for refresh
- `CWR_ResetMONNavigationCounters` — Resets navigation counters
- `CWR_MONValueEquals` — Compares MON values
- `CWR_MONNavigationEvidence` — Records navigation evidence
- `CWR_LoadFilteredWTGrid` — Loads filtered WT grid from MON

**Create WT / WHTask Operations:**
- `CWR_BatchOpenCreateWT` — Opens Create WT dialog (NO WHTASK)
- `CWR_CreateOutboundWTFromExactCandidateFast` — Fast WT creation
- `CWR_FindCreateWTSession` — Finds Create WT session
- `CWR_FindExactCreateWTSession` — Finds exact Create WT session
- `CWR_WaitForCreateWTSession` — Waits for Create WT session ready
- `CWR_OpenExactCreateWTAfterQtyChange` — Opens Create WT after qty change
- `CWR_VerifyExactCreateWTRequest` — Verifies Create WT request
- `CWR_ProveCreateWTTargetCandidate` — Proves Create WT target validity
- `CWR_ReadExactNewOutboundWT` — Reads newly created WT

**Posting/Bin/Batch Operations:**
- `CWR_OpenPOSTSession` — Opens MIGO/posting session
- `CWR_QueryExpectedPOST` — Queries expected posting
- `CWR_FindCurrentPOSTGrid` — Finds current posting grid
- `CWR_SelectExactPOSTSource` — Selects exact source in posting
- `CWR_ProveExactPOSTSourceAndTarget` — Proves source/target validity
- `CWR_RecordPostingChange` — Records posting change
- `CWR_RecordPostingCreate` — Records posting creation
- `CWR_ResetPOSTWithoutSave` — Resets posting without save
- `CWR_AssertNoPendingPOSTDraft` — Asserts no pending drafts

**Cancel Operations:**
- `CWR_SendExactCancel` — Sends cancel for exact WT
- `CWR_PostVerifyCancelled` — Post-verifies cancellation
- `CWR_ReconcileOldWTCancelState` — Reconciles old WT cancel state
- `CWR_RecordCancelledSnapshot` — Records cancelled snapshot
- `CWR_SchedulerSendQueuedCancel` — Sends queued cancellations
- `CWR_SchedulerVerifyCancelledGroup` — Verifies cancelled group

**State Management:**
- `CWR_LoadExecutionContext` — Loads execution context (reopens too often)
- `CWR_RecordRemaining` — Records remaining work
- `CWR_RecordTargetStock` — Records target stock state
- `CWR_RecordOutboundWTCreated` — Records created WT
- `CWR_RecordCreatedPostingWT` — Records created posting WT
- `CWR_HistoryCheckpoint` — Creates history checkpoint
- `CWR_HistoryRunID` — Manages run IDs
- `CWR_HistorySheet` — History sheet operations

**Gates/Assertions (TOO MANY):**
- `CWR_AssertNoActiveOutboundWT` — No active WT assertion (buggy)
- `CWR_AssertCFHandoff` — CF handoff assertion
- `CWR_AssertReworkGate` — Rework gate assertion
- `CWR_AssertCancelledLocal` — Local cancellation assertion
- `CWR_AssertTargetStockLocal` — Target stock local assertion
- `CWR_AssertPostingVerifiedLocal` — Posting verified local assertion
- `CWR_AssertCurrentPlanExecutionContract` — Plan contract assertion
- `CWR_AssertFullRouteActionContract` — Full route contract assertion
- `CWR_AssertQueuePlanLock` — Queue plan lock assertion
- `CWR_RequireRegressionGate` — **Regression gate** (blocks live runs)
- `CWR_RequireReworkAction` — Rework action requirement

**Regression/Testing:**
- `CWR_RunRegressionAudit` — Runs regression audit
- `CWR_RecordSuccessfulSchedulerRegressionProofs` — Records proofs
- `CWR_RegressionResultCount` — Counts regression results
- `CWR_IsRequiredRegressionScenario` — Checks required scenarios
- `CWR_BuildRequiredRegressionScenarios` — Builds scenario list

---

## 2. KEY PROCEDURE CALL CHAINS

### Button #4 Execution Flow (Simplified)

```text
USER CLICKS "XỬ LÝ ĐẾN CF" (Button #4)
    ↓
[Entry Point - likely CWR_RunFastToCFCore or similar]
    ↓
CWR_ExecuteOBQueueToCFHandoffCore
    ↓
FOR EACH OB IN QUEUE:
    ├─ CWR_LoadExecutionContext (OPENS CONTEXT #1)
    ├─ CWR_SchedulerUsePlan (loads plan, may REOPEN)
    ├─ CWR_AssertReworkGate
    ├─ CWR_RequireRegressionGate (HARD STOP if scenarios not proven)
    │
    ├─ CANCEL PHASE:
    │   ├─ CWR_SchedulerSendQueuedCancelGroup
    │   ├─ CWR_SendExactCancel (per WT)
    │   ├─ CWR_PostVerifyCancelled
    │   └─ CWR_RecordCancelledSnapshot
    │
    ├─ LOCATION/POSTING PHASE:
    │   ├─ CWR_ExecuteLocationGroups
    │   ├─ CWR_OpenPOSTSession
    │   ├─ CWR_QueryExpectedPOST
    │   ├─ CWR_SelectExactPOSTSource
    │   └─ CWR_RecordPostingChange
    │
    ├─ QTY/PRDO PHASE:
    │   ├─ CWR_LoadExecutionContext (REOPENS CONTEXT #2)
    │   ├─ CWR_ExecuteQtyGroups
    │   ├─ CWR_ExecuteDocumentGroups
    │   ├─ CWR_SchedulerVerifyPRDOGroup (checks active member count)
    │   ├─ CWR_BatchSetPRDOQtyOneSave
    │   │   ├─ CWR_OpenExpectedPRDO
    │   │   ├─ CWR_FindExactPRDOItem (per item)
    │   │   ├─ CWR_SelectExactPRDOItem (per item)
    │   │   ├─ [Fill all rows top-to-bottom]
    │   │   └─ [ONE SAVE for document]
    │   └─ CWR_BatchAllPRDOTargets (post-verify)
    │
    ├─ CREATE WT PHASE:
    │   ├─ CWR_LoadExecutionContext (REOPENS CONTEXT #3) ← BUG SOURCE
    │   ├─ CWR_ReuseVerifiedPRDOTabForMON ← SESSION REUSE BUG
    │   ├─ CWR_BatchAssertNoActiveOutboundWT ← STALE CACHE BLOCKS HERE
    │   ├─ CWR_BatchOpenCreateWT (tries to open Create WT)
    │   │   ├─ [NO WHTASK NAVIGATION]
    │   │   └─ [Tries MON→OB→PRDO→CreateWT path]
    │   └─ CWR_CreateOutboundWTFromExactCandidateFast
    │
    └─ CF HANDOFF:
        └─ [User manual CF or automated if enabled]
```

### The Bug Sequence

```text
1. Cancel Phase:
   - CWR_SendExactCancel → WT 5042889690 cancelled
   - CWR_PostVerifyCancelled → Status A verified
   - RAM should mark: WTState[5042889690].Status = "A", Terminal = TRUE

2. PRDO Phase:
   - CWR_LoadExecutionContext → Reopens (unnecessary)
   - CWR_BatchSetPRDOQtyOneSave → PRDO 580→640
   - PRDO_POSTVERIFY → PASS
   - PRDO session now has correct Qty

3. Create Phase:
   - CWR_LoadExecutionContext → Reopens AGAIN (CONTEXT #3)
   - CWR_ReuseVerifiedPRDOTabForMON → **REUSES PRDO TAB FOR MON**
     - Log: "SAME_TAB=TRUE"
     - Sends "/SCWM/MON" to the PRDO session
     - Reads stale OB/WT grid
   - CWR_BatchAssertNoActiveOutboundWT:
     - Reads MON grid showing WT 5042889690
     - Grid may show old state before Cancel refresh
     - RAM says "A", but MON cache says "Active"
     - **STOPS**: "BATCH_CREATE_ACTIVE_OUTBOUND_WT_PRESENT|WT=5042889690"

4. Result:
   - DOCUMENT_OPEN_COUNT=3 (repeated reopening)
   - WHTASK_OPEN_COUNT=0 (never used WHTask navigation)
   - Old terminal WT blocks Create due to stale state
```

---

## 3. STATE MANAGEMENT APPROACH (CURRENT)

### Storage Mechanisms

**Worksheets as Runtime State:**
- `PLAN` sheet — execution plan with ~52 columns
- `WH_ACTUAL` — user intent + system state
- `POSTING` / `POSTING_CACHE` — posting state
- `OB_QUEUE` / `OB_WH_CACHE` — OB-level cache
- `RUN_LOG` — detailed execution log
- `RUN_HISTORY` — historical runs
- `REGRESSION_AUDIT` — regression tracking

**Global Variables (inferred from patterns):**
- Session IDs (MON_SESSION, PRDO_SESSION, CREATE_SESSION)
- Navigation counters (MON_OPEN_COUNT, DOCUMENT_OPEN_COUNT, etc.)
- Current context (CWR_EXPECTED_OB, CWR_EXPECTED_DOC, CWR_EXPECTED_WT, etc.)

**Problems:**
1. **Sheet writes are slow** — Every state change hits Excel cells
2. **No RAM-first model** — Verified terminal states not cached in RAM
3. **Context reopening** — `CWR_LoadExecutionContext` called 3x per OB
4. **Stale cache precedence** — Sheet/SAP cache overrides fresh proofs

### Current Context Variables (from symbols)

```text
CWR_EXPECTED_OB
CWR_EXPECTED_DOC
CWR_EXPECTED_WT
CWR_EXPECTED_ITEMS
CWR_EXPECTED_MATERIAL
CWR_EXPECTED_SOURCE_QTY
CWR_EXPECTED_SOURCE_BIN
CWR_EXPECTED_SOURCE_BATCH
CWR_EXPECTED_DEST_BIN
CWR_EXPECTED_PROCTY2
CWR_TARGET_QTY
CWR_CREATE_QTY
CWR_TARGET_BIN
CWR_TARGET_BATCH
CWR_TARGET_STORAGE_TYPE
CWR_TARGET_STORAGE_SECTION
CWR_TARGET_STOCK_TYPE
CWR_EXPECTED_WAREHOUSE
CWR_WH_BATCH
CWR_PLAN_ID
CWR_PLAN_DECISION
CWR_PLAN_QTY_MODE
CWR_PLAN_EXEC_MODE
CWR_PLAN_ROW
CWR_PLAN_WH_ROW
CWR_HISTORY_RUN_ID
CWR_HISTORY_SEQ
CWR_MON_OPEN_COUNT
CWR_MON_REFRESH_COUNT
CWR_DOCUMENT_OPEN_COUNT
CWR_DOCUMENT_RECOVERY_COUNT
CWR_WHTASK_OPEN_COUNT
CWR_WHTASK_REUSE_COUNT
CWR_WHTASK_RECOVERY_COUNT
CWR_NODE_KEY
CWR_UPPER_ID
CWR_WT_ID
CWR_ERP_IDD
CWR_DOC_ID
CWR_EXECUTE_ID
CWR_WT_ACTION
```

These are likely **module-level globals**, not a structured RunState object.

---

## 4. SAP SESSION HANDLING (CURRENT)

### Identified Sessions

**MON Session** — `/SCWM/MON` (Monitor Warehouse Tasks)
- Used for: Initial WT snapshots, global warehouse view
- Functions: `CWR_OpenMonitorForRead`, `CWR_GetMONSession`
- Counter: `CWR_MON_OPEN_COUNT`, `CWR_MON_REFRESH_COUNT`

**PRDO Session** — `/SCWM/PRDO` (Process Delivery Order)
- Used for: Document-specific operations, Qty changes
- Functions: `CWR_OpenExpectedPRDO`, `CWR_FindCurrentPRDOGrid`
- Counter: `CWR_DOCUMENT_OPEN_COUNT`, `CWR_DOCUMENT_RECOVERY_COUNT`

**POST Session** — `MIGO` or similar (Material Document)
- Used for: Bin/batch posting changes
- Functions: `CWR_OpenPOSTSession`, `CWR_QueryExpectedPOST`

**Create WT Session** — Opens from PRDO or MON
- Functions: `CWR_FindCreateWTSession`, `CWR_WaitForCreateWTSession`

### Session Reuse Bug

**Current Code:**
```vba
' CWR_ReuseVerifiedPRDOTabForMON
' Reuses the verified /SCWM/PRDO tab for /SCWM/MON readback
' SAME_TAB=TRUE; SAP_WRITE=0
```

**Why this is wrong:**
1. PRDO session is document-bound with filters/context
2. Sending `/SCWM/MON` into same session doesn't clear PRDO state
3. Grid may show OB-level data but with stale WT status
4. No proper session separation

**Evidence from strings:**
```text
"Reused the verified /SCWM/PRDO tab for /SCWM/MON readback; SAME_TAB=TRUE"
"STOP_MON_SESSION_REQUIRED: no existing idle /SCWM/MON or /SCWM/PRDO session"
"STOP_MON_SESSION_AMBIGUOUS: idle /SCWM/PRDO sessions="
```

### Navigation Pattern (Current — WRONG)

```text
Phase 1 (Cancel):
  MON session → Select OB → Cancel WT → Read status

Phase 2 (PRDO Qty):
  Open PRDO → Load Document → Change Qty → Save → Close

Phase 3 (Create WT — BUG):
  Reopen context → **Reuse PRDO tab as MON**
  → Read "MON" grid (actually stale PRDO-MON hybrid)
  → Block on old WT that was proven A
```

### WHTask Navigation (MISSING)

**Zero evidence of WHTask button use:**
- `WHTASK_OPEN_COUNT=0` in problem run
- No `CWR_ReturnToWarehouseTask` or equivalent
- No `CWR_PressWHTaskButton` or similar
- After PRDO Qty save, should press WHTask to see Open Qty
- Instead, reopens entire OB/Document context

---

## 5. WHERE GROUPING BREAKS DOWN

### Current Group Concepts (Partial)