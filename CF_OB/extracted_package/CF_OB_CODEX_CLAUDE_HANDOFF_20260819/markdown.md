# SAP CF OB — CODEX / CLAUDE REBUILD & REFACTOR SPECIFICATION

> **Package date:** 2026-08-19 (Asia/Bangkok / UTC+7)  
> **Current binary baseline:** `current/CURRENT_CF_OB_BASELINE_20260819.xlsm`  
> **Original uploaded filename:** `d2f0e204-a047-4d2c-814b-258d2d21e7b4.xlsm`  
> **SHA-256:** `9ebcc46fcdf25a10803792d2ed576c79f98a394eec838a783d5b0ee021905309`  
> **Purpose:** Give an implementation agent enough business, state, navigation, workbook, test, and regression context to rebuild/refactor the SAP CF OB VBA tool without falling back to the old per-P navigation model.

---

## 0. READ THIS FIRST — AUTHORITY AND NON-NEGOTIABLE RULES

This document is the architecture contract for the next refactor. When current VBA behavior conflicts with this document, treat the conflict as a **known legacy defect to refactor**, not as proof that the legacy behavior is desired.

**Never do these things while refactoring:**

1. Do not create a brand-new workbook and discard the current workbook structure, buttons, shapes, sheets, or VBA project. Work on a copy of the included current `.xlsm`.
2. Do not run live SAP writes during offline development/replay. All live writes must remain explicitly gated.
3. Do not execute SAP actions inside a `For Each P` member loop. P rows are planning members, not independent workflows.
4. Do not reuse the PRDO session/tab as MON. `MON_SESSION_ID <> PRDO_SESSION_ID` is a hard invariant.
5. Do not reopen the same OB/Document repeatedly. Normal target: `DOCUMENT_OPEN_COUNT = 1` per document per Button #4 execution.
6. After a Document is open, use the **WHTask** button/surface to return to WH/WT. Do not reopen the OB just to read WT again.
7. Do not allow a WT that has been proven `A` or `C` in the current run to become Active again because an older cache says so.
8. Do not use the old `AREA_ALLOWED=P,ZF,HR` model as a core restriction. Exact bins such as `APFL-CC01` must be supported when SAP/stock evidence validates them.
9. Do not make regression-history/live-proven gates block normal live execution. Regression is a development/offline quality gate; live safety gates must protect actual SAP target/write correctness only.
10. Keep automatic CF disabled/user-handoff unless explicitly enabled and live-proven for the target build.

---

## 1. PACKAGE CONTENTS AND TRUST LEVELS

### 1.1 Current binary (highest implementation authority)

- `current/CURRENT_CF_OB_BASELINE_20260819.xlsm` — exact current workbook binary included.
- `current_extracted/sheets/*.csv` — raw cached cell values of all current sheets, readable without Excel.
- `current_extracted/workbook_manifest.json` — sheet dimensions, package information and file hash.
- `current_extracted/formulas.json` — formula inventory from worksheet XML.
- `current_extracted/vbaProject.bin` — extracted VBA project binary for diagnostics only; the user-facing workbook remains the canonical carrier.
- `current_extracted/vba_symbols_inventory.txt` — strings/symbols found in VBA binary.
- `current_extracted/vba_strings_full.txt` — raw ANSI/UTF-16 strings from VBA binary; useful when no VBA source extractor is available.
- `current_extracted/evidence/latest_run_excerpt.csv` — latest RunID rows pulled from RUN_LOG.

### 1.2 Legacy references

See `legacy/legacy_reference_manifest.csv` and `legacy/LEGACY_REFERENCES.md`. The older binaries were visible as File Library references but were not mounted in this runtime, so their bytes are not fabricated into this ZIP.

### 1.3 Supporting historical patterns

The `support/` directory contains prior notes about stable SAP automation patterns: shared-lock recovery, MIGO session reuse/event polling, VL06O exact filtering/retry, and full SAP tab/control mapping. These are design references only; they are not CF OB business logic.

---

## 2. WHAT THE TOOL IS SUPPOSED TO DO

The CF OB workbook manages outbound warehouse task correction/confirmation against actual picking. User intent comes from `WH_ACTUAL`, while SAP provides current WH/WT, PRDO quantity, stock/bin/batch and WT status.

High-level user flow:

```text
INPUT left side: list of OBs
        ↓
1. TẢI WH
        ↓
User edits Actual Qty / Actual Bin / Actual Batch in WH_ACTUAL
        ↓
2/3. CHECK + POSTING (read/plan)
        ↓
4. XỬ LÝ ĐẾN CF
        ↓
Group Executor performs required Cancel / Posting / Qty / Create WT
        ↓
Stop at CF user handoff unless automatic CF is explicitly enabled
```

The outer queue may contain multiple OBs. **Each OB is processed sequentially; inside each OB, all P members are grouped and executed as one coordinated state machine.**

---

## 3. UI CONTRACT — DO NOT MISREAD INPUT

The `INPUT` sheet is not a single-OB form. Its layout contract is:

```text
┌──────────────────────────────┬──────────────────────────────┐
│ LEFT: OB LIST / QUEUE        │ RIGHT: CONTROL PANEL        │
│                              │                              │
│ OB 1                         │ 1. TẢI WH                   │
│ OB 2                         │ 2/3. CHECK + POSTING        │
│ OB 3                         │ 4. XỬ LÝ ĐẾN CF            │
│ ...                          │ status / progress           │
└──────────────────────────────┴──────────────────────────────┘
```

Do not redesign this as one OB input at the top. Preserve the list-on-left / buttons-on-right model. UI may be cleaned up, but semantics must remain.

The runtime hierarchy is:

```text
OB QUEUE
  ├─ OB #1 → Build Groups → Execute Group Pipeline
  ├─ OB #2 → Build Groups → Execute Group Pipeline
  └─ OB #N → Build Groups → Execute Group Pipeline
```

---

## 4. CURRENT WORKBOOK MAP

The current workbook contains 20 sheets. Snapshots are under `current_extracted/sheets/`.

| Sheet | Rows | Cols | Used dimension | Current state |
|---|---:|---:|---|---|
| `ORDER_QUERY` | 2 | 6 | `A1:F2` | `veryHidden` |
| `POSTING_RAW` | 71 | 17 | `A1:Q71` | `veryHidden` |
| `INPUT` | 506 | 14 | `A1:N506` | `visible` |
| `WH_ACTUAL` | 504 | 32 | `A1:AF504` | `visible` |
| `POSTING` | 500 | 31 | `A1:AE500` | `visible` |
| `RUN_LOG` | 1044 | 24 | `A1:X1044` | `visible` |
| `SAP_MAP` | 1747 | 16 | `A1:P1747` | `visible` |
| `CONFIG` | 261 | 6 | `A1:F261` | `visible` |
| `OB_QUEUE` | 2 | 24 | `A1:X2` | `visible` |
| `OB_WH_CACHE` | 3 | 26 | `A1:Z3` | `visible` |
| `POSTING_SUMMARY` | 2 | 20 | `A1:T2` | `veryHidden` |
| `POSTING_CACHE` | 23 | 19 | `A1:S23` | `veryHidden` |
| `PLAN_SOURCE_LEGS` | 2 | 7 | `A1:G2` | `veryHidden` |
| `PLAN_RELEASE` | 1 | 21 | `A1:U1` | `veryHidden` |
| `SAP_PROBE_LOG` | 1 | 17 | `A1:Q1` | `veryHidden` |
| `PLAN` | 5 | 52 | `A1:AZ5` | `visible` |
| `SAP_ACTION_MAP` | 28 | 10 | `A1:J28` | `visible` |
| `RUN_HISTORY` | 684 | 27 | `A1:AA684` | `visible` |
| `RUN_SUMMARY` | 38 | 14 | `A1:N38` | `visible` |
| `REGRESSION_AUDIT` | 1302 | 9 | `A1:I1302` | `visible` |

### 4.1 User-facing logical sheets

**INPUT**
- Left: OB queue/list.
- Right: buttons/control/status.
- Should show high-level progress, not raw technical STOP strings.

**WH_ACTUAL**
- Main operational table.
- System fields: OB, Document, Item, WH, WT, Material, Sys Qty/Bin/Batch, Status.
- User intent fields: Actual Qty, Actual Bin, Actual Batch.
- User should not need to understand internal plan fingerprints/gates.

**RUN_SUMMARY / history view**
- Appropriate place for concise run outcomes.

### 4.2 Engine/persistence sheets

`POSTING_RAW`, `POSTING`, `POSTING_CACHE`, `POSTING_SUMMARY`, `PLAN`, `PLAN_SOURCE_LEGS`, `PLAN_RELEASE`, `OB_QUEUE`, `OB_WH_CACHE` are engine/persistence/audit material. In the refactored system, they must **not be the live in-run state store**.

### 4.3 Developer/audit sheets

`RUN_LOG`, `RUN_HISTORY`, `REGRESSION_AUDIT`, `SAP_MAP`, `SAP_ACTION_MAP`, `SAP_PROBE_LOG`, `CONFIG` should be developer/audit surfaces and may be hidden/VeryHidden in normal UI after the core engine is stable.

---

## 5. BUSINESS STATE MODEL

### 5.1 WT status meaning

- blank/open/active: WT still active.
- `A`: WT cancelled / terminal.
- `C`: WT confirmed / terminal.

A status describes the **old WT**, not the user's current pick intent.

### 5.2 Required interpretation of Status A

| Old WT status | User Actual Qty | Meaning | Decision | Retained Qty | Need Create |
|---|---:|---|---|---:|---:|
| A | blank | User does not use this member | `A_UNUSED` | 0 | 0 |
| A | 0 | Explicit not-pick | `A_UNUSED` | 0 | 0 |
| A | >0 | User still wants this pick | `A_RECREATE_REQUIRED` | 0 | Actual Qty minus already reconciled matching WT |

After exact stock/bin/batch feasibility passes, transition `A_RECREATE_REQUIRED → A_RECREATE_READY`.

**Never demand that an A WT become Active again before continuing.** It is terminal. If the user wants quantity, create/reconcile a **new WT number**.

### 5.3 Status C

`C` means satisfied/terminal. On rerun, do not recreate simply because user inputs remain. Reconcile it as completed.

---

## 6. GROUP FIRST — NEVER EXECUTE PER P

P rows are member records for intent and allocation. They are **not independent SAP workflows**.

### 6.1 Member loop may only

```text
READ
NORMALIZE
CLASSIFY
CALCULATE
PLAN
```

No Cancel, Save, Posting, Create or CF inside the member loop.

### 6.2 Group executor owns SAP writes

```text
ExecuteCancelGroup
ExecutePostingGroup
ExecuteQtyGroup
ExecuteCreateWTGroup
ExecuteCFGroup / CF user handoff
```

### 6.3 Fixed phase order

```text
LOAD ALL MEMBERS
        ↓
NORMALIZE USER INTENT
        ↓
CLASSIFY A / ACTIVE / C
        ↓
BUILD ALLOCATION LEDGER
        ↓
RECONCILE EXISTING WT
        ↓
LOCK GROUP / PLAN
        ↓
CANCEL PHASE
        ↓
REFRESH / UPDATE WT STATE ONCE
        ↓
POSTING / BIN-BATCH PHASE
        ↓
CHANGE QTY PHASE
        ↓
ONE PRDO SAVE PER DOCUMENT
        ↓
WHTask
        ↓
VERIFY OPEN QTY
        ↓
CREATE REQUIRED ALLOCATIONS
        ↓
WHTask
        ↓
VERIFY NEW WT
        ↓
CF USER HANDOFF / CONFIRM
```

If the same document contains Cancel + Increase + Decrease + Change Batch, all of them are planned first and then executed in the same run in the fixed phase order. No P is allowed to start a private mini-workflow.

---

## 7. GROUP / MEMBER / ALLOCATION KEYS

Recommended core keys:

**Document group:**
```text
Document + Item (+ Material as identity proof)
```

**Allocation key:**
```text
Document + Item + Material + ExactActualBin + ActualBatch
```

The exact bin/batch allocation matters. If 50 EA goes to one bin and 20 EA goes to another, they are two Create WT allocations even if material/item are the same.

### 7.1 Suggested MemberState fields

```text
PId
OB
Document
Item
Material
OldWT
OldWTStatus
SysQty
SysBin
SysBatch
ActualQty
ActualBin
ActualBatch
DesiredQty
RetainedQty
NeedCreateQty
Decision
AllocationKey
```

### 7.2 GroupState fields

```text
GroupKey
Members[]
CancelMembers[]
KeepMembers[]
RecreateMembers[]
QtyIncreaseMembers[]
QtyDecreaseMembers[]
PostingMembers[]
CreateAllocations[]
PRDOTargetQty
RetainedQty
NeedCreateQty
State
```

Suggested state progression:

```text
PLANNED → CANCEL_DONE → POSTING_DONE → PRDO_DONE → CREATE_DONE → CF_READY → DONE
```

No backwards phase jumping in a normal execution.

---

## 8. QTY LOGIC — VERIFY DELTA AT WHTASK, NOT BY REOPENING EVERYTHING

The current tool over-checks Qty by reopening/refreshing surfaces. Refactor to use the locked plan and WHTask Open Qty.

Example:

```text
PRDO/target: 600
Valid active WT retained: 580
Expected Open Qty = 20
```

After setting PRDO Qty and saving once, navigate to WHTask and verify the Open Qty/delta. Do not reopen MON/OB just to re-prove the number.

General rule:

```text
ExpectedOpenQty = TargetQty - Sum(Qty of valid retained WT after prior actions)
```

For a cancelled A WT:

```text
RetainedQty(A) = 0
```

Therefore an A member with ActualQty 80 may require CreateQty 80, not `80 - old cancelled qty`.

### 8.1 Group Qty save

- Compute every target before writing.
- Apply all PRDO targets top-to-bottom for the document.
- Save once per document.
- Do not require active-member count to equal pre-cancel member count. Cancel naturally changes active membership.
- `GROUP_MEMBER_INCOMPLETE` is diagnostic/warning when exact PRDO target rows remain unambiguous.

---

## 9. CANCEL LOGIC

Cancel should be grouped. For each exact eligible WT:

1. Pre-action gate proves exact WT and that it is active.
2. Send Cancel.
3. Post-action proof verifies `Active → A`.
4. Update RAM `WTState` immediately.
5. Do **not** reopen OB or MON just because Cancel succeeded.

If multiple WTs in the same group need Cancel, perform them in one phase, then read/refresh the relevant WT surface once.

---

## 10. BIN / BATCH / POSTING — EXACT BIN IS FIRST CLASS

Current CONFIG historically restricts area to P/ZF/HR. That is too narrow. Bins such as `APFL-CC01` must work.

Core validation should be based on exact bin identity and SAP/stock evidence, not a hard-coded family list.

Recommended normalization:

```text
Raw Bin   = APFL-CC01
Warehouse = APFL
BinCode   = CC01
Family    = CC   (optional classification only)
```

Or:

```text
APFL-P16-01-20 → Warehouse=APFL, Family=P, ExactBin=APFL-P16-01-20
```

**Core acceptance:** an `APFL-*` exact bin that SAP/Posting cache proves exists and is eligible. Family classification must not reject an otherwise valid bin unless a specific business rule explicitly requires it.

Posting Change runs only when actual source/target bin/batch differs and the exact stock path requires it.

---

## 11. TWO FIXED SAP SESSION ROLES

This is a core navigation refactor.

### 11.1 MON_ROLE

- Dedicated SAP tab/session kept on `/SCWM/MON`.
- Use for initial global WH snapshot and holder checks outside current document.
- Do not turn it into PRDO.

### 11.2 PRDO_ROLE

- Dedicated SAP tab/session kept on `/SCWM/PRDO` / current outbound document.
- Open the OB/Document once.
- All document quantity/bin/batch/Create-WT interactions stay here.
- Do not send `/SCWM/MON` into this session.

### 11.3 Hard invariant

```text
MON_SESSION_ID <> PRDO_SESSION_ID
```

If roles collide, fail before SAP write with an internal/session-context error. Never silently reuse the PRDO tab for MON readback.

---

## 12. WHTASK IS THE RETURN PATH AFTER DOCUMENT OPEN

Once the Document is bound, WH/WT navigation should be:

```text
Current PRDO Document
   ↓
press WHTask
   ↓
read/reuse WT grid
```

Not:

```text
PRDO → MON → reopen OB → PRDO → WHTask
```

Implement one navigation helper, e.g.:

```vb
CWR_ReturnToWarehouseTask(...)
```

Contract:

1. Assert PRDO_ROLE still owns the expected Document.
2. If WT screen/grid is already available, reuse it.
3. Otherwise press WHTask.
4. If the proven UI requires a second WHTask press, allow a bounded second press.
5. Read grid / Open Qty.
6. **Never call OpenOB, MON navigation, or document search from this helper.**

Normal acceptance:

```text
DOCUMENT_OPEN_COUNT = 1
DOCUMENT_RECOVERY_COUNT = 0
WHTASK_OPEN_COUNT >= 1 when WT work is needed
```

---

## 13. RAM STATE STORE — WORKSHEETS ARE NOT THE IN-RUN DATABASE

The executor currently re-reads sheets/cache and SAP surfaces too often. Refactor to keep current execution state in RAM.

Recommended `RunState`:

```text
RunID
CurrentOB
CurrentDocument
MonSession
PrdoSession
DocumentBound
DocumentOpenCount
DocumentRecoveryCount
WTMapVersion
WTMapDirty
MemberStates
GroupStates
AllocationStates
PRDOTargets
CurrentPRDOQty
OpenQty
LastSAPAction
LastVerifiedAction
```

Recommended `WTState(WTNumber)`:

```text
WT
Status
Qty
Bin
Batch
Terminal
LastVerifiedAt
Source
MapVersion
```

### 13.1 Monotonic terminal transitions

```text
ACTIVE → A   (Cancel)
ACTIVE → C   (CF)
```

For the **same WT number** within a run:

```text
A → ACTIVE   forbidden
C → ACTIVE   forbidden
```

If an older snapshot says Active after RAM has verified A/C, mark that snapshot stale. Refresh once if needed. Only stop if a fresh SAP read truly contradicts the terminal state.

### 13.2 Narrow dirty invalidation

- Cancel: update WT immediately; mark WT map dirty if a fresh grid view is required. Document/PRDO context remains valid.
- Change Qty: update PRDO target state; `OpenQty=UNKNOWN` until WHTask read. No MON invalidation.
- Create WT: mark WT map dirty; WHTask read once then clean.
- CF: set exact WT to C; no document reopen.

Use `WTMapVersion` to know whether a consumer already has a sufficiently fresh snapshot.

---

## 14. SAFETY GATES — REDUCE TO THREE LAYERS

The current build accumulated overlapping safety gates that can PASS one moment and STOP on stale data the next. Simplify.

### 14.1 INPUT GATE — once before SAP write

Verify:

- exact OB / Document / Item / Material identity;
- required user Actual fields;
- if Actual Qty > 0, exact Actual Bin is present;
- batch present when required;
- no ambiguous target row.

Then set `PLAN_LOCKED=TRUE`. Do not repeatedly rerun the same gate at every phase.

### 14.2 ACTION GATE — immediately before a dangerous write

Examples:

- Cancel: exact WT and currently eligible/active.
- Qty: exact PRDO row and allowed target.
- Posting: exact source/target/qty/batch.
- Create WT: exact Open Qty / allocation and bin/batch target.

### 14.3 POST-ACTION VERIFY — verify only the action just sent

- Cancel → exact WT becomes A.
- Qty → Save success and later WHTask Open Qty matches target delta.
- Create → new WT appears with correct qty/bin/batch.
- CF → exact WT becomes C.

### 14.4 Keep hard STOP set small

Recommended categories:

```text
STOP_INPUT_INCOMPLETE
STOP_EXACT_TARGET_NOT_FOUND
STOP_AMBIGUOUS_TARGET
STOP_INSUFFICIENT_STOCK
STOP_SAP_WRITE_ERROR
STOP_POSTVERIFY_FAILED
STOP_STATE_CONTRADICTION_AFTER_REFRESH
STOP_SESSION_CONTEXT_LOST
```

Everything else should generally be WARN/diagnostic unless it genuinely makes the next SAP write ambiguous or unsafe.

Move regression/live-proven requirements out of normal live hard-stop path.

---

## 15. RERUN / IDEMPOTENCY

The system must resume safely after partial success.

Example: old WT is A, user wants 80 at P20/B1, and a previous run already created a new Active WT of 80 but stopped later. On rerun:

1. Reconcile exact Document+Item+Material+Bin+Batch.
2. If existing Active/C WT satisfies 80 → `A_RECREATE_SATISFIED`, `NeedCreate=0`.
3. If 60 is already satisfied → only `NeedCreate=20`.
4. Never recreate the whole 80 just because the old WT remains A.

Persist enough checkpoints to recover from:

- after Cancel,
- after PRDO Save,
- after Posting,
- after Create WT,
- before CF handoff.

But runtime logic should prefer fresh RAM within the same run.

---

## 16. CURRENT ROOT-CAUSE EVIDENCE — LATEST RUN

Use `current_extracted/evidence/latest_run_excerpt.csv` and full `RUN_LOG.csv`. The latest run proves the current architecture bug rather than a SAP business failure.

Observed sequence in the current workbook:

1. Group plan is built and locked.
2. Old WT `5042889690` is cancelled.
3. Cancel readback proves old WT status `A`.
4. Posting / bin-batch action passes.
5. PRDO quantity change from `580` to `640` is saved and post-verified PASS.
6. A later create step opens/rebinds execution context again.
7. Log explicitly reports reuse of verified `/SCWM/PRDO` tab for `/SCWM/MON` readback with `SAME_TAB=TRUE`.
8. Create then stops because the same old WT `5042889690` is considered active, contradicting its previously proven A state.
9. Navigation summary shows repeated document opening and no WHTask use (`DOCUMENT_OPEN_COUNT=3`, `WHTASK_OPEN_COUNT=0` in the problematic run).

**Interpretation:** this is a stale-state/session-navigation defect. Do not add another gate around it. Fix state ownership and navigation.

---

## 17. CURRENT CASE DATA TO KEEP AS A REPLAY FIXTURE

One high-value current fixture in the workbook:

```text
OB          1011832412
Document    1012087422
Item        10
WH          6026342483
Old WT      5042889690
Material    BIW32023716AA
Sys Qty     580
Target Qty  640
Sys Bin     APFL-P14-01-08
Target Bin  APFL-P09-01-06
Sys Batch   260716-02
Target Batch 260816-02
Decision    CHANGE_BIN_BATCH / Qty increase path
```

Expected refactored outcome:

- Cancel old WT once → terminal A in RAM.
- Do not reopen OB to re-prove A.
- Complete bin/batch Posting/change as planned.
- Set PRDO Qty target once and Save once.
- Use WHTask to read Open Qty.
- Create only required new WT allocation.
- Same old WT must never block Create as Active.
- Stop at CF handoff unless CF is explicitly enabled.

---

## 18. FUNCTIONS / AREAS TO REFACTOR FIRST

The VBA string inventory confirms these current procedures/symbols. Inspect actual VBA source in Excel before changing signatures.

### `CWR_SchedulerVerifyPRDOGroup`
- Remove active-member completeness as a hard Qty gate after Cancel.
- Verify exact target rows/identity, not unchanged member count.

### `CWR_BatchSetPRDOQtyOneSave`
- Apply all document/group targets once.
- One Save per document.
- Update RAM state on success.

### `CWR_BatchAllPRDOTargets`
- Use as diagnostic/postcheck as appropriate, not a reason to repeatedly reopen/recheck before Create.

### `CWR_BatchOpenCreateWT`
- Must use bound PRDO document + WHTask navigation.
- Must not reopen the same OB/document.
- Must consume RAM state and fresh WHTask/Open Qty.

### `CWR_BatchAssertNoActiveOutboundWT`
- Must respect RAM terminal proof.
- If cache disagrees with terminal state, classify it stale and refresh once.
- Do not stop on the old WT after verified A.

### `CWR_SchedulerUsePlan`
- If requested Document equals bound current Document, reuse.
- No second `EXECUTION_CONTEXT_OPEN` for the same Document in a normal phase transition.

### `CWR_ResetSchedulerExecutionContext`
- Reset at Button #4 / scheduler boundary or true session loss only.
- Do not clear/rebind between Cancel, Qty and Create phases.

### `CWR_RequireRegressionGate`
- Move regression blocking to development/offline build validation.
- Live executor should not stop because an historical scenario lacks a live-proven flag.

### Add a single WHTask navigation function

Recommended name:

```text
CWR_ReturnToWarehouseTask
```

See Section 12 contract.

---

## 19. PROPOSED MODULE BOUNDARIES

Keep changes understandable and avoid another monolithic module. A safe incremental split could be:

```text
modCWR_StateStore       RAM Run/WT/Member/Group state
modCWR_GroupPlanner     classification + group/allocation build, no SAP writes
modCWR_Navigation       MON_ROLE / PRDO_ROLE binding + WHTask navigation
modCWR_GroupExecutor    fixed phase state machine
modCWR_ReplayAdapter    offline SAP adapter / fixtures
```

Existing `modCodexWTRework` can remain the public entry point initially and delegate to these modules. This minimizes button/macro binding risk.

If introducing class modules would destabilize import/export, use `Scripting.Dictionary` and typed helper functions first. Architecture is more important than object-oriented style.

---

## 20. LIVE / REPLAY SAP ADAPTER

Create an adapter boundary so the same planner/executor can run at home without SAP.

```text
SAP Adapter
   ├── LIVE   → SAP GUI Scripting
   └── REPLAY → sheet/history fixtures
```

The planner/executor must not know whether data came from SAP or replay.

Replay must use **raw snapshots/events**, not the old `Decision` column as the expected answer. The new planner must independently derive decisions. This prevents pass-by-cheating.

Suggested adapter operations:

```text
ReadInitialWTs
ReadPostingStock
OpenDocument
CancelWT
ApplyPostingChange
SetPRDOTargetsAndSave
OpenWarehouseTask
ReadOpenQty
CreateWT
ReadWTs
ConfirmWT (disabled by default)
```

In REPLAY mode, write operations mutate an in-memory simulation state and append deterministic events.

---

## 21. SHEET / UI REFACTOR AFTER ENGINE STABILIZES

Do core State/Group/Navigation first, then UI cleanup.

Recommended normal visible sheets:

1. `INPUT` — left OB list + right control panel/progress.
2. `WH_ACTUAL` — system vs user actual input.
3. `RUN_SUMMARY` or a concise HISTORY surface.

Possible developer/VeryHidden sheets:

- POSTING / POSTING_RAW / caches
- PLAN and plan helper sheets
- OB_QUEUE / OB_WH_CACHE
- SAP_MAP / SAP_ACTION_MAP / SAP_PROBE_LOG
- RUN_LOG / RUN_HISTORY / REGRESSION_AUDIT
- CONFIG if users do not need to edit it routinely

**Important:** hiding a sheet is not permission to delete its data contract. First remove runtime dependencies, then hide/cleanup.

### User-friendly A status

Show operational meaning, e.g.:

- `A — KHÔNG LẤY`
- `A — CẦN TẠO WH MỚI`
- `C — HOÀN THÀNH`
- `THAY ĐỔI QTY/BIN/BATCH`

Keep raw technical decision/status available in hidden columns/log.

### Progress panel

Example:

```text
OB 1011832412
[✓] PLAN LOCKED
[✓] CANCEL
[✓] WT STATE
[✓] BIN/BATCH
[✓] QUANTITY
[>] CREATE WT
[ ] VERIFY
[ ] CF HANDOFF
```

If stopped, show user reason and exact item/material/bin/need/available. Raw technical code remains in RUN_LOG.

---

## 22. LOG RETENTION / PERFORMANCE

The workbook already contains large audit sheets. Prevent unbounded growth:

- keep concise RUN_SUMMARY for a reasonable recent window;
- cap/rotate RUN_LOG and RUN_HISTORY;
- do not append the entire regression scenario set on every live run;
- runtime state lives in RAM; cache sheets persist checkpoints, not every micro-read;
- avoid repeated SAP tree/control scans after IDs are mapped and validated.

---

## 23. OFFLINE ACCEPTANCE TEST MATRIX

See `tests/REPLAY_SCENARIOS.csv` and `tests/ACCEPTANCE_TESTS.md`. At minimum the build must cover:

1. Normal CF-only.
2. Increase Qty.
3. Decrease Qty.
4. Cancel only.
5. Cancel + Increase.
6. Cancel + Decrease.
7. Change Batch.
8. Change Batch + Increase.
9. Change Bin.
10. Bin `APFL-CC01`.
11. A + Actual blank.
12. A + Actual Qty > 0.
13. A + repick at new bin.
14. Multiple P same Document.
15. Increase + Decrease same group/run.
16. Multiple Create WT allocations at different bins.
17. Rerun after Cancel.
18. Rerun after PRDO Save.
19. Rerun after Create WT.
20. Existing WT already C.

Every refactor must run the full offline suite before live testing.

---

## 24. ACCEPTANCE METRICS / INVARIANTS

A normal document run should satisfy:

```text
DOCUMENT_OPEN_COUNT = 1
DOCUMENT_RECOVERY_COUNT = 0
MON_OPEN_COUNT <= 1
MON global read only at initial/exception points
PRDO_SAVE_COUNT <= 1 per document for grouped Qty update
No MON_SESSION_REUSE with SAME_TAB=TRUE
WHTask used for WT return/navigation after document is bound
No terminal WT (A/C) becomes Active for the same WT number
No SAP write inside member/P planning loop
No re-plan after write phase begins
No duplicate Create WT on rerun
```

Suggested run metrics:

```text
GROUP_COUNT
MEMBER_COUNT
DOCUMENT_OPEN_COUNT
DOCUMENT_RECOVERY_COUNT
MON_OPEN_COUNT
MON_READ_COUNT
WHTASK_CLICK_COUNT
WT_MAP_REFRESH_COUNT
CANCEL_ACTION_COUNT
PRDO_SAVE_COUNT
CREATE_WT_COUNT
RAM_HIT_COUNT
SAP_READ_COUNT
HARD_STOP_COUNT
WARN_COUNT
```

---

## 25. LIVE TEST LADDER AT COMPANY

Do not jump directly to a large OB queue. After offline suite passes:

### Live 1 — simple known-safe CF path
Correct Qty/Bin/Batch → reach CF/handoff; confirm proven navigation.

### Live 2 — Cancel
Active WT → Cancel → prove A → no repeated MON/OB reopen.

### Live 3 — Qty delta
`580 → 600` with valid retained 580 → Save → WHTask → Open Qty 20 → Create 20.

### Live 4 — Change Batch + Increase
The path currently most likely to expose stale-context bugs.

### Live 5 — mixed multi-P group
Cancel + Increase + Decrease + Batch/Bin + Create in the same document/run.

Only after these pass should multiple OBs in the INPUT queue be run.

---

## 26. IMPLEMENTATION PHASES FOR CODEX / CLAUDE

### Phase A — baseline preservation
- Copy current XLSM.
- Export VBA source using Excel/VBE tooling available on the developer machine.
- Record current module list, button macro bindings, shapes, sheet visibility and workbook hash.
- Do not touch live SAP.

### Phase B — state store
- Implement RunState/WTState/MemberState/GroupState.
- Add monotonic terminal transition rules.
- Add dirty/version mechanics.

### Phase C — group planner
- Move all SAP writes out of P/member loops.
- Fully classify every P before `PLAN_LOCKED`.
- Implement A semantics and exact allocation keys.

### Phase D — navigation roles
- Bind MON_ROLE and PRDO_ROLE separately.
- Ban same-session role collision.
- Enforce one document open.
- Add WHTask return helper.

### Phase E — executor phases
- CancelGroup → PostingGroup → QtyGroup → WHTask/OpenQty → CreateGroup → Verify → CF handoff.
- Update RAM after each verified action.

### Phase F — gate cleanup
- Consolidate to Input / Action / PostVerify layers.
- Convert non-critical legacy blockers to WARN.
- Remove regression suite from live hard-stop path.

### Phase G — dynamic bins
- Generalize exact bin validation.
- Keep area/family only as optional classification/business rule.

### Phase H — replay harness
- Implement LIVE/REPLAY adapter.
- Seed scenarios from raw current/legacy snapshots.
- Run full regression offline.

### Phase I — UI/sheet cleanup
- Preserve INPUT list-left/buttons-right.
- Improve WH_ACTUAL status/actions.
- Hide technical sheets only after engine no longer needs them as runtime state.

### Phase J — live proof
- Follow Section 25 ladder.
- Attach the new RUN_LOG/RUN_SUMMARY evidence to the next review.

---

## 27. DO-NOT-REGRESS CHECKLIST

- Exact MON Document/WT context from V0.4.x remains intact.
- Known real CF path that reached `Status=C` remains intact.
- Posting uses exact material/bin/batch/quantity targets.
- Current ability to group Cancel and perform one PRDO Save is preserved but rewritten around stable state/session ownership.
- Current user workflow and sheet field semantics remain recognizable.
- Automatic CF remains disabled/user handoff unless explicitly re-enabled.
- No destructive workbook rebuild that loses VBA, buttons, shapes, names or compatibility.

---

## 28. EXPECTED OUTPUT FROM AN IMPLEMENTATION AGENT

When Codex/Claude finishes an offline refactor, it should return:

1. Updated `.xlsm` based on the included current binary.
2. Exported/changed VBA modules or a patch diff.
3. `CHANGELOG.md` listing every behavioral change.
4. `OFFLINE_TEST_RESULTS.md` with all replay scenarios and invariants.
5. No claim of “100% live pass” before company SAP live tests.
6. Explicit list of remaining live-only proofs.
7. Confirmation that `DOCUMENT_OPEN_COUNT`, session role collision, stale terminal WT and per-P SAP-write regressions are covered by tests.

---

## 29. QUICK START FOR A NEW CODEX / CLAUDE SESSION

1. Read this file completely.
2. Read `PROMPT_FOR_CODEX_CLAUDE.md`.
3. Inspect `current/CURRENT_CF_OB_BASELINE_20260819.xlsm`.
4. Read `current_extracted/workbook_manifest.json`.
5. Read `current_extracted/sheets/CONFIG.csv`, `PLAN.csv`, `WH_ACTUAL.csv`, `RUN_LOG.csv`, `RUN_HISTORY.csv`, `REGRESSION_AUDIT.csv`.
6. Read `current_extracted/evidence/latest_run_excerpt.csv`.
7. Read `current_extracted/vba_symbols_inventory.txt`.
8. Read `legacy/LEGACY_REFERENCES.md`.
9. Build offline/replay first.
10. Do not write to live SAP unless the user explicitly moves to live test mode.

---

## 30. FINAL DESIGN PRINCIPLE

The goal is not “be safe by reopening SAP until everything is re-proven.” The goal is:

> **Read exact state → lock the complete group plan → execute deterministic phases → update trusted RAM state after each verified action → refresh only the state made dirty by that action → continue through WHTask without reopening the OB.**

That principle simultaneously fixes the four recurring problem classes: broken grouping, conflicting safety gates, inflexible bin assumptions, and repeated MON/OB navigation.
