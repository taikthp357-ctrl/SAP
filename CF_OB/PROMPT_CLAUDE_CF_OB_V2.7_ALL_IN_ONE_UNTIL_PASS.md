# MASTER ORCHESTRATOR PROMPT FOR CLAUDE
# CF_OB V2.7 — ALL-IN-ONE INGEST → REFACTOR → UI REDESIGN → TEST → FIX → REGRESSION → PACKAGE UNTIL PASS

> **Use this prompt as the top-level authority for the whole task.**
>
> This is not a request for analysis only. This is an implementation-and-verification task.
>
> You are expected to consume the entire supplied workspace/package, understand the history and current binary, implement V2.7, run every test that the environment can actually execute, fix failures, rerun, and continue iterating until all attainable acceptance gates pass.
>
> **Important clarification:** “ALL-IN-ONE” describes how YOU, Claude, must execute the engineering task end-to-end. It does **NOT** mean merging the workbook buttons. The final workbook UI must still have **4 separate buttons**:
>
> `TẢI WH → CHECK → XỬ LÝ → LÀM MỚI`
>
> `CHECK` remains READ-ONLY. `XỬ LÝ` remains the only normal business-write entry point.

---

# 0. PRIMARY GOAL

Upgrade the current CF_OB workbook lineage from:

`260819_CF_OB_V2.6_LICH_SU`

to a production-oriented:

`CF_OB V2.7`

that is:

- safer;
- easier to operate;
- easier to debug;
- materially lighter and cleaner at workbook/sheet level;
- group-first rather than P-row-first;
- correct for multi-WT quantity redistribution;
- deterministic;
- idempotent;
- recoverable after interruption;
- strict about MON/session ownership;
- able to plan release/cancel of other holder WT safely;
- explicit about what was sent to SAP and what was verified;
- designed around a single modern Excel Control Center;
- still stops at **USER CF HANDOFF**;
- never claims live success without evidence.

The intended architectural direction is the latest 4-button / 5-sheet V2.7 specification.

---

# 1. DO NOT STOP AT A PLAN

Do **not** respond with only:

- an audit;
- a proposal;
- pseudocode;
- a TODO list;
- a patch suggestion;
- “please test this manually” after the first edit;
- “I cannot continue without confirmation” for ordinary recoverable engineering decisions.

Perform the work.

Your normal loop is:

```text
INGEST
→ INVENTORY
→ AUTHORITY RESOLUTION
→ BASELINE SNAPSHOT
→ IMPLEMENT
→ COMPILE/STATIC CHECK
→ OFFLINE TEST
→ REGRESSION
→ UI/STRUCTURE TEST
→ INTEGRATION TEST
→ INSPECT FAILURE
→ PATCH
→ RERUN
→ REPEAT
→ PACKAGE
→ FINAL AUDIT
```

If a test fails, do not stop at reporting the failure.

Inspect the exact evidence, patch the underlying defect, and rerun the relevant suite plus regression.

Continue until the acceptance gates in this prompt are satisfied or until an **external boundary that the current environment genuinely cannot cross** is reached.

---

# 2. EXTERNAL-BLOCKER RULE — NEVER FAKE SUCCESS

If the environment lacks one of the following:

- Microsoft Excel/VBE compilation;
- Windows COM automation;
- SAP GUI Scripting;
- authenticated SAP session;
- an explicitly approved live test OB/data target;
- permission to perform a specific SAP write;

then:

1. finish **all work that does not depend on that missing capability**;
2. run all offline/static/replay/package tests;
3. create a fully testable output package;
4. identify the missing gate precisely;
5. mark it `BLOCKED_EXTERNAL`, not `PASS`;
6. do not invent SAP results;
7. do not claim “LIVE PASS” from static analysis.

If the missing capability becomes available inside the same working environment, continue automatically.

---

# 3. SOURCE MATERIAL — READ EVERYTHING PROVIDED

You may receive many files, including:

- current `.xlsm`;
- older `.xlsm`;
- `.md`;
- `.txt`;
- `.csv`;
- `.json`;
- screenshots/UI mockups;
- ZIP archives;
- nested ZIP archives;
- exported Codex/Claude conversations;
- extracted VBA source;
- extracted workbook XML;
- logs/history;
- regression reports;
- audit reports;
- SAP mapping documentation.

**Do not cherry-pick one markdown file and ignore the rest.**

Recursively inventory the entire supplied workspace.

---

# 4. ZIP / ARCHIVE INGESTION CONTRACT

For every `.zip` or archive supplied:

1. do not modify the original archive;
2. calculate hash if possible;
3. extract into an isolated read-only-style ingest directory, e.g.:

```text
/_claude_ingest/<archive_name>/
```

4. recursively inspect nested archives;
5. produce a manifest of extracted content;
6. detect duplicate files by hash;
7. identify:
   - current binary;
   - prior baselines;
   - MD specifications;
   - conversation exports;
   - VBA source;
   - logs;
   - screenshots;
   - reports;
   - test fixtures;
8. do not assume archive modification time alone establishes authority.

If a ZIP contains a packaged Codex conversation, treat it as **historical engineering context and evidence**, not automatically as the newest specification.

---

# 5. CODEX / CLAUDE CONVERSATION INGESTION

Conversation exports are important because they contain:

- user corrections;
- live test outcomes;
- discovered SAP behavior;
- failed approaches;
- exact stop codes;
- navigation discoveries;
- design decisions;
- known regressions;
- what was actually proven versus only proposed.

Parse them carefully.

For each important statement classify it as:

```text
USER_REQUIREMENT
USER_CORRECTION
IMPLEMENTATION_PROPOSAL
IMPLEMENTATION_CLAIM
LIVE_EVIDENCE
WORKBOOK_EVIDENCE
FAILED_ATTEMPT
SUPERSEDED_DECISION
OPEN_QUESTION
```

Do not treat an assistant's old claim such as “fixed successfully” as proof unless supported by:

- SAP readback;
- workbook history;
- regression evidence;
- a resulting binary;
- or later user confirmation.

User corrections override earlier assistant assumptions.

---

# 6. AUTHORITY HIERARCHY — CRITICAL

There are older documents that describe:

- 2 buttons;
- 3 buttons;
- `CHECK + POSTING`;
- `CHECK & XỬ LÝ`;
- many legacy sheets;
- older per-P workflow assumptions.

They are useful historically but are not the final V2.7 UI/architecture contract.

Resolve conflicts using this order:

## Authority Tier 0 — This master prompt

This file controls how the task is executed.

## Authority Tier 1 — Latest explicit user decisions

The latest required UI/architecture is:

```text
4 buttons:
1. TẢI WH
2. CHECK
3. XỬ LÝ
4. LÀM MỚI
```

and the target workbook architecture is:

```text
5 logical sheets:
INPUT
DATA_CORE
PLAN
SYSTEM
RUN_HISTORY
```

`POSTING` must no longer remain as a runtime sheet after migration is proven safe.

## Authority Tier 2 — Latest detailed V2.7 spec

Prefer the newest document with a name equivalent to:

`PROMPT_CODEX_CF_OB_V2.7_REDESIGN_4BUTTON_5SHEETS_HISTORY.md`

This is the detailed architectural contract for V2.7.

## Authority Tier 3 — Latest V2.6 audit and latest workbook/run evidence

Examples include files equivalent to:

`260820_CF_OB_V2.6_AUDIT_TO_V2.7.md`

and the latest `260819_CF_OB_V2.6_LICH_SU.xlsm`.

These establish current defects and proven behavior.

## Authority Tier 4 — Current binary reality

Actual code, sheet dependencies, control IDs, workbook history and SAP evidence in the current baseline are implementation facts.

When the newest design conflicts with legacy code, that conflict is normally a migration task, not a reason to silently keep the old design.

## Authority Tier 5 — Packaged Codex/Claude conversations

Use them to recover rationale, live-proof, failures, exact SAP flow and regression scenarios.

## Authority Tier 6 — Older MD/UI concepts

Older 2-button/3-button documents and UI mockups are **design references only** where they conflict with Tier 1/2.

Mine good ideas from them, but do not resurrect superseded behavior.

---

# 7. BUILD A CONFLICT LEDGER BEFORE EDITING

Create an internal file/report:

`V2.7_SOURCE_CONFLICT_LEDGER.md`

For every meaningful conflict record:

- topic;
- source A;
- source B;
- old behavior;
- latest required behavior;
- chosen decision;
- why;
- implementation impact.

Minimum expected conflicts to resolve:

- 3 buttons vs 4 buttons;
- CHECK+EXECUTE merged vs separated;
- 20 sheets vs 5 sheets;
- POSTING visible/runtime vs removed;
- old P-loop vs group scheduler;
- one/shared SAP session assumptions vs explicit session broker;
- automatic/old CF test code vs USER_CF_HANDOFF;
- holder cancel enabled flag vs actual mapping/readback readiness.

Do this once, then implement according to the resolved ledger.

---

# 8. BASELINE SAFETY

Never destroy the source baseline.

Create:

```text
/work/
  BASELINE/
  WORKING/
  OUTPUT/
  REPORTS/
  TEST_ARTIFACTS/
```

Before editing, capture:

- original filename;
- SHA-256;
- file size;
- workbook sheet list;
- visibility;
- used ranges;
- defined names;
- VBA module inventory;
- public macro inventory;
- button/shape macro bindings;
- relationships;
- workbook protection state;
- current CONFIG;
- SAP map/action map fingerprints if present.

Write:

`V2.7_BASELINE_AUDIT.md`

---

# 9. NEVER REBUILD THE WORKBOOK FROM SCRATCH

Work from a copy of the current workbook.

Do not create a blank workbook and copy data over unless there is a documented catastrophic binary corruption and no safer path exists.

Preserve:

- VBA project;
- workbook relationships;
- workbook names;
- controls/shapes;
- relevant styles;
- existing proven logic;
- latest run/history evidence during migration.

---

# 10. ALL-IN-ONE ENGINEERING DOES NOT MEAN “BIG BANG” EDITING

Execute the task in one continuous engineering run, but internally use controlled phases.

Suggested working checkpoints:

```text
CP0 BASELINE
CP1 CHECK/EXEC SPLIT
CP2 MEMBER/GROUP ALLOCATION
CP3 REPOSITORY MIGRATION
CP4 SINGLE MON BROKER
CP5 HOLDER RELEASE PLANNER
CP6 HISTORY V2
CP7 5-SHEET CONSOLIDATION
CP8 UI FINAL
CP9 FULL REGRESSION
CP10 PACKAGE
```

At each checkpoint:

- compile/static audit;
- run focused tests;
- preserve a rollback copy.

Do not wait for user confirmation between ordinary checkpoints.

---

# 11. FINAL UI — 4 BUTTONS, NOT 3

The final user-facing workflow is:

```text
TẢI WH
   ↓
USER INPUT ACTUAL
   ↓
CHECK
   ↓
PLAN REVIEW / READY
   ↓
XỬ LÝ
   ↓
USER CF HANDOFF
```

and:

`LÀM MỚI`

is a separate safe reset operation.

Buttons:

1. `TẢI WH`
2. `CHECK`
3. `XỬ LÝ`
4. `LÀM MỚI`

Do not merge #2 and #3.

---

# 12. CHECK IS READ-ONLY

This is non-negotiable.

`CHECK` may:

- validate user input;
- read MON;
- read current WT;
- read source stock;
- read holder candidates;
- build member plan;
- build group plan;
- calculate release needs;
- select release candidates;
- build source legs;
- build action plan;
- calculate fingerprints;
- display blockers;
- lock PlanID.

`CHECK` may not:

- Cancel;
- Save;
- change qty;
- change bin;
- change batch;
- POST;
- Create WT;
- Confirm WT;
- auto CF.

Implement an explicit execution context.

Example:

```text
ExecutionMode = CHECK | EXECUTE | RECOVERY_READ
```

All SAP business-write functions must assert:

```text
ExecutionMode = EXECUTE
```

before sending a write.

---

# 13. XỬ LÝ RUNS ONLY A LOCKED PLAN

`XỬ LÝ` must not silently rebuild CHECK.

Before any write compare:

- CheckID;
- PlanID;
- InputFingerprint;
- WHSnapshotFingerprint;
- SourceStockFingerprint;
- ReleasePlanFingerprint;
- PlanFingerprint;
- SapMapFingerprint;
- ConfigFingerprint.

If stale:

```text
STOP_PLAN_STALE_RECHECK_REQUIRED
```

Zero SAP write.

---

# 14. CRITICAL GROUP BUG — MUST PASS

The latest V2.6 evidence includes a net-zero total case similar to:

```text
WT A: 3500 → 0
WT B: 4500 → 8000

Current group total = 8000
Target group total  = 8000
```

Legacy logic can conclude:

```text
GROUP_CURRENT = GROUP_TARGET
→ PRDO_QTY_ALREADY_TARGET
→ NO_QTY_WRITE
```

That conclusion is only valid for aggregate PRDO quantity.

It must **not** suppress member reallocation.

Correct member ledger:

```text
A:
  CancelQty = 3500
  RetainQty = 0
  CreateQty = 0

B:
  CancelQty = 0
  RetainQty = 4500
  RequiredCreateQty = 3500
```

Correct group:

```text
GroupNetDelta = 0
GroupCancelTotal = 3500
GroupRetainTotal = 4500
GroupCreateTotal = 3500
PRDOAggregateWriteRequired = possibly FALSE
AllocationWriteRequired = TRUE
```

Expected execution:

```text
CANCEL 3500
→ VERIFY A
→ RETAIN 4500
→ CREATE 3500
→ VERIFY MEMBER FINAL 8000
→ VERIFY GROUP FINAL 8000
```

This case is a release-blocking regression test.

---

# 15. PRESERVE THE NET-ZERO REPLACE CASE THAT ALREADY WORKED

Do not over-fix the previous bug by applying create-delta blindly to every net-zero group.

There is historical evidence of a replace case equivalent to:

```text
5050 → 0
12950 → 18000 + change bin/batch
group total remains 18000
```

where the correct path was effectively:

- cancel old group;
- POST change;
- PRDO Save;
- create full required replacement allocation;
- verify new WT.

The planner must distinguish:

```text
KEEP_EXISTING_CREATE_DELTA
```

from:

```text
REPLACE_CANCEL_CREATE
```

Use the member ledger, exact retained quantity and action semantics.

---

# 16. GROUP FIRST, NEVER PRIVATE P WORKFLOWS

P rows/members are planning entities.

A P/member loop may:

```text
READ
NORMALIZE
CLASSIFY
CALCULATE
PLAN
```

It may not independently:

- open/reopen OB;
- Cancel;
- Save;
- POST;
- Create;
- CF.

Writes belong to a group/document executor.

---

# 17. SCHEDULER HIERARCHY

Use:

```text
RUN
  └─ OB
      └─ DOCUMENT
          ├─ CANCEL OWN GROUP
          ├─ MON READBACK
          ├─ RELEASE HOLDER GROUP
          ├─ MON READBACK
          ├─ POST CHANGE GROUP
          ├─ PRDO DOCUMENT GROUP
          ├─ CREATE GROUP
          └─ FINAL VERIFY
```

Normal target:

- one planned execution path per document;
- one PRDO save per document where possible;
- no reopen per member;
- no “one P passes, next P starts from the beginning”.

---

# 18. SINGLE MON BROKER

Implement one owner for MON navigation.

Recommended role:

`modSapSessionBroker`

or equivalent.

Contract:

- reuse a dedicated MON reader session/tab;
- all MON readers request access through the broker;
- holder reader does not open its own MON;
- final verify does not open its own MON;
- cancel verification reuses broker;
- no module sends `/SCWM/MON` directly except broker recovery.

Log navigation metrics:

- MON_OPEN_COUNT
- MON_REFRESH_COUNT
- MON_RECOVERY_COUNT
- DOCUMENT_OPEN_COUNT
- WHTASK_OPEN_COUNT
- WHTASK_REUSE_COUNT

For a normal single-document case target approximately:

```text
MON_OPEN_COUNT <= 1
MON_REFRESH_COUNT <= 4
DOCUMENT_OPEN_COUNT <= 1
```

unless exact SAP state requires bounded recovery.

---

# 19. SESSION ROLE COLLISION

If historical documents contain older advice that conflicts about session use, infer the latest safe architecture from evidence.

Prefer explicit role ownership.

At minimum distinguish:

```text
MON_ROLE
DOCUMENT/PRDO_ROLE
```

Do not casually transform the MON reader tab into PRDO and then rebuild MON repeatedly.

Reuse the existing correct context instead of opening tabs “just in case”.

---

# 20. EVENT / READBACK OVER FIXED SLEEP

Where possible use:

- session.Busy;
- status bar transition;
- popup presence/close;
- exact control state;
- row/state fingerprint;
- MON readback;
- PRDO readback;
- WT status.

Timeout is a bounded fallback.

Do not use long fixed sleeps as proof of success.

---

# 21. HOLDER WT / OTHER WH RELEASE — PLAN FIRST

The holder-cancel feature is not yet trustworthy enough to make it an unconditional live production action.

V2.7 must support:

```text
RELEASE_HOLDER_MODE =
  PLAN_ONLY
  PILOT
  LIVE_LOCKED
```

Final default:

`PLAN_ONLY`

unless the latest explicit user-supplied configuration says otherwise.

---

# 22. HOLDER RELEASE CALCULATION

Use:

```text
RequiredReleaseQty =
max(
  RequiredCreateQty
  - FreeAvailableQty
  - VerifiedOwnCancelCredit,
  0
)
```

Only query holder WT if the required release is positive.

---

# 23. HOLDER CANDIDATE FILTER

Exact identity:

- Material;
- Bin;
- Batch.

Candidate must have stable:

- HolderOB;
- HolderDocument;
- HolderWH;
- HolderWT;
- HolderQty;
- HolderStatus.

Exclude:

- A;
- C;
- current target WT;
- newly created WT;
- wrong material;
- wrong bin;
- wrong batch;
- unstable identity;
- unplanned same-OB holder unless explicitly allowed.

Default:

`RELEASE_ALLOW_SAME_OB_UNPLANNED = FALSE`

---

# 24. HOLDER SELECTOR

Deterministic selection.

Prefer:

1. exact source;
2. active status;
3. smallest whole-WT combination covering need;
4. fewer WTs;
5. stable WT-ID tie breaker.

Do not silently change candidates during EXECUTE.

If state changes after CHECK:

```text
STOP_RELEASE_PLAN_STALE_RECHECK_REQUIRED
```

---

# 25. HOLDER CANCEL SUCCESS REQUIRES READBACK

Never add released quantity to available stock just because Cancel was pressed.

Only add credit when MON proves exact holder WT:

```text
Status = A
```

If outcome is uncertain:

```text
ACTION_UNKNOWN
RECOVERY_REQUIRED = TRUE
NO_RETRY
NO_NEXT_DEPENDENT_WRITE
```

---

# 26. HOLDER LIVE TEST ROLLOUT

If SAP live access and an explicitly safe/approved test target are available:

## Step 1

Test `PLAN_ONLY`.

Expected:

- candidates found;
- selection deterministic;
- coverage calculated;
- no cancel sent.

## Step 2

Switch only the test scope to `PILOT`.

One release group.

Perform:

- exact fresh precheck;
- one cancel;
- one MON readback;
- verify A.

If not proven:

- stop;
- UNKNOWN;
- do not retry.

## Step 3

Only after repeated pilot proof may test `LIVE_LOCKED`.

Do not ship the final general configuration casually enabled merely because one pilot passed.

---

# 27. NO AUTOMATIC CF

Final V2.7 stops at:

`READY_FOR_USER_CF`

or equivalent.

Never automatically confirm CF in the final production path.

Any historical CF pilot code must remain disabled or isolated.

The purpose of regression is to prove the tool reaches a correct handoff, not to re-enable live CF.

---

# 28. TARGET 5-SHEET ARCHITECTURE

Final logical workbook:

1. `INPUT`
2. `DATA_CORE`
3. `PLAN`
4. `SYSTEM`
5. `RUN_HISTORY`

Normal user mode:

- only `INPUT` visible;
- the other four VeryHidden.

Debug/Admin mode may reveal technical sheets.

---

# 29. LEGACY 20-SHEET MIGRATION

The current workbook may include legacy sheets such as:

- ORDER_QUERY
- POSTING_RAW
- INPUT
- WH_ACTUAL
- POSTING
- RUN_LOG
- SAP_MAP
- CONFIG
- OB_QUEUE
- OB_WH_CACHE
- POSTING_SUMMARY
- POSTING_CACHE
- PLAN_SOURCE_LEGS
- PLAN_RELEASE
- SAP_PROBE_LOG
- PLAN
- SAP_ACTION_MAP
- RUN_HISTORY
- RUN_SUMMARY
- REGRESSION_AUDIT

Do not delete them first.

Migrate safely:

```text
A. create new repositories
B. dual-read / compare
C. switch engine
D. prove parity
E. search zero runtime references
F. remove legacy sheets
```

---

# 30. POSTING SHEET MUST BE REMOVED FROM RUNTIME

The old `POSTING` sheet is redundant if raw/normalized stock already exists elsewhere.

Target:

```text
SAP capture
→ normalize in memory
→ DATA_CORE.tblSourceStock
```

Optional raw evidence:

`DATA_CORE.tblSourceStockRaw`

No user-facing/runtime sheet named `POSTING` in final V2.7.

Before removing it:

1. export/search all VBA;
2. find every `Worksheets("POSTING")` and equivalent;
3. build dependency map;
4. migrate reader;
5. migrate writer;
6. compile;
7. run replay;
8. confirm zero runtime references;
9. only then delete/decommission.

No hidden fallback routine may recreate `POSTING`.

---

# 31. DATA_CORE

Consolidate data/state formerly spread over:

- ORDER_QUERY;
- POSTING_RAW;
- WH_ACTUAL backend;
- POSTING;
- OB_QUEUE;
- OB_WH_CACHE;
- POSTING_SUMMARY;
- POSTING_CACHE.

Logical tables should include at least:

```text
tblOBQueue
tblWHSnapshot
tblActualSnapshot
tblSourceStock
tblSourceStockRaw
tblCaptureSummary
tblRuntimeState
```

Do not use worksheet row number as business identity.

---

# 32. PLAN

Consolidate planning into tables equivalent to:

```text
tblMemberPlan
tblGroupPlan
tblSourceLeg
tblReleaseCandidate
tblActionPlan
```

Every member has a stable `SourceKey`.

Every write has a stable `ActionID`.

---

# 33. SYSTEM

Consolidate:

- CONFIG;
- SAP_MAP;
- SAP_ACTION_MAP;
- SAP_PROBE_LOG.

Logical tables:

```text
tblConfig
tblSapMap
tblSapActionMap
tblSapProbe
```

Prefer technical IDs and map fingerprints.

No production fallback via SendKeys or guessed captions.

---

# 34. RUN_HISTORY V2

Unify:

- RUN_LOG;
- RUN_HISTORY;
- RUN_SUMMARY;
- REGRESSION_AUDIT.

Create one durable event ledger.

At minimum each meaningful action can answer:

- what run?
- what check?
- what plan?
- what action?
- what OB/document/item/WT?
- before?
- target?
- after?
- sent?
- read back?
- verified?
- unknown?
- duration?
- SAP session/control?
- stop code?
- resume state?

---

# 35. HISTORY ID MODEL

Use:

```text
LoadRunID
CheckID
PlanID
ExecuteID
ActionID
ParentActionID
```

An action already VERIFIED must not be resent.

An action UNKNOWN must not be resent until recovery reconciles it.

---

# 36. ACTION STATES

Normalize:

```text
PLANNED
PRECHECKED
SENT
READBACK
VERIFIED
SKIPPED_VERIFIED
BLOCKED
UNKNOWN
```

Never allow an ambiguous “PASS” that cannot distinguish whether the write was actually sent.

---

# 37. HISTORY EVENT TYPES

Support event classes equivalent to:

```text
LOAD_START
LOAD_ROW
LOAD_COMPLETE

CHECK_START
INPUT_SNAPSHOT
MON_READ
SOURCE_STOCK_READ
HOLDER_READ
PLAN_MEMBER
PLAN_GROUP
PLAN_RELEASE
PLAN_ACTION
CHECK_PASS
CHECK_BLOCKED

EXEC_START
ACTION_PRECHECK
ACTION_SENT
ACTION_READBACK
ACTION_VERIFIED
ACTION_UNKNOWN
ACTION_SKIPPED_VERIFIED
EXEC_BLOCKED
EXEC_COMPLETE

USER_CF_HANDOFF

RECOVERY_REQUIRED
RECOVERY_RESOLVED

REGRESSION_PASS
REGRESSION_FAIL
```

Do not log every polling tick.

---

# 38. HISTORY PERFORMANCE

Do not make V2.7 slower/heavier by logging noise.

Log:

- state transition;
- decision;
- SAP write boundary;
- readback;
- timeout;
- proof;
- blocker;
- recovery.

Do not log:

- every DoEvents;
- every 100ms poll;
- the same unchanged status repeatedly.

Use buffered/batch worksheet writes where safe.

Flush immediately for:

- SENT;
- UNKNOWN;
- fatal error;
- recovery checkpoint.

---

# 39. RECOVERY

On workbook open, inspect latest unresolved history.

If a prior SAP business write is `SENT/UNKNOWN`:

- set `RECOVERY_REQUIRED`;
- do not run a new plan;
- perform read-only reconciliation.

Reconcile exact:

- WT;
- status;
- document/item;
- created WT;
- quantity/bin/batch where needed.

If proven successful:

`ACTION_VERIFIED_AFTER_RECOVERY`

If proven not sent/no mutation and retry is safe:

retry only through the normal action contract.

Never blindly resend.

---

# 40. NUMBER PARSER — MANDATORY REGRESSION

Test at least:

```text
"1.200" → 1200
"1.380" → 1380
"1.640" → 1640
"20,5"  → 20.5
"0"     → valid zero
blank   → blank/unset, not zero
```

Do not reintroduce:

`1200 → 1.2`

Quantities must not use Integer where overflow/scale risk exists.

---

# 41. BLANK VS ZERO

Semantics:

```text
ActualQty blank = user has not supplied intent
ActualQty 0     = explicit target/cancel intent
```

Never coerce blank to 0 during planning.

---

# 42. STATUS A / C

Preserve established semantics:

- `A` = old WT cancelled/terminal;
- `C` = confirmed/terminal.

A cancelled WT does not have to become active again.

If user still wants quantity, create/reconcile a new WT.

C rows:

- remain in backend/history;
- do not appear in the default actionable UI projection.

---

# 43. UI DESIGN — USE ALL UI REFERENCES, BUT IMPROVE THEM

Study:

- existing INPUT sheet;
- screenshots/mockups in the supplied workspace;
- older UI MD documents;
- user corrections in conversation packages.

You may improve beyond the mockups when it materially improves:

- clarity;
- speed;
- readability;
- error prevention;
- long-list usability;
- recovery visibility;
- plan review.

However, do not violate the final interaction contract.

---

# 44. UI NON-NEGOTIABLES

- one user-facing sheet: `INPUT`;
- 4 buttons;
- multi-OB queue;
- actionable worklist;
- status C excluded from default worklist;
- editable cells visually distinct;
- system fields protected;
- clear CHECK state;
- clear Plan state;
- clear Execute state;
- clear Recovery state;
- holder-release warning;
- no technical STOP dump as the primary user message;
- no MsgBox spam;
- freeze panes suitable for long lists;
- usable around 1366x768 desktop Excel;
- avoid excessive horizontal scrolling;
- do not rely on color alone for state.

---

# 45. UI SUGGESTED LAYOUT

Use the supplied dashboard images as inspiration, but adapt for 4 buttons.

Recommended high-level layout:

```text
┌─────────────────────────────────────────────────────────────────────────────┐
│ CF OB — CONTROL CENTER V2.7                           SAP ●   PLAN / RUN     │
├─────────────────┬───────────────────────────────────────────────────────────┤
│ DANH SÁCH OB    │ [1 TẢI WH] [2 CHECK] [3 XỬ LÝ] [4 LÀM MỚI]             │
│                 ├───────────────────────────────────────────────────────────┤
│ 101...          │ STATUS / SUMMARY CARDS                                   │
│ 101...          │ Loaded | Ready | Blocked | Release | Done C | Recovery  │
│ ...             ├───────────────────────────────────────────────────────────┤
│                 │ ACTIONABLE WORKLIST                                       │
│                 │ OB | Mat | WT | SysQty | ActualQty | Bin | Batch | ...  │
│                 │ Decision | ReleaseNeed | Plan | Exec | Message           │
└─────────────────┴───────────────────────────────────────────────────────────┘
```

Avoid decorative complexity that makes the VBA UI fragile.

Prefer stable Excel shapes/cells over unsupported custom dependencies.

---

# 46. UI BUTTON STATE MACHINE

Initial:

```text
TẢI WH  enabled
CHECK   disabled
XỬ LÝ   disabled
LÀM MỚI enabled
```

After Load PASS:

```text
CHECK enabled
XỬ LÝ disabled
```

After user edits:

```text
old plan = STALE
XỬ LÝ disabled
```

After CHECK PASS:

```text
XỬ LÝ enabled
```

After CHECK BLOCKED:

```text
XỬ LÝ disabled
```

During EXEC:

all action buttons locked against double-run.

After EXEC completes:

`XỬ LÝ` disabled until a new CHECK.

---

# 47. UI RELEASE WARNING

If a plan requires other-WT release, surface a visible warning:

```text
CÓ WH/WT KHÁC ĐANG GIỮ TỒN
CHECK CHỈ LẬP RELEASE PLAN
CHƯA CANCEL NẾU MODE = PLAN_ONLY
```

Show:

- RequiredReleaseQty;
- selected holder count;
- coverage;
- mode.

---

# 48. UI RESET

`LÀM MỚI` must not erase evidence required for recovery.

If unresolved SENT/UNKNOWN exists:

- preserve history;
- preserve recovery state;
- do not pretend the workbook is clean.

Allow safe visual/input reset only where appropriate.

---

# 49. SOURCEKEY AND ROW IDENTITY

Every visible worklist row must map via immutable `SourceKey`.

Do not sync by worksheet row position.

Detect:

- duplicate SourceKey;
- missing SourceKey;
- stale SourceKey.

Stop before write on ambiguity.

---

# 50. PERFORMANCE

Optimize both Excel and SAP interaction.

Excel:

- array/block reads;
- array/block writes;
- dictionaries;
- no preallocated hundreds of blank POSTING rows;
- do not render hidden backend data onto INPUT.

SAP:

- reuse session;
- reuse document;
- refresh by phase;
- bounded navigation;
- exact technical controls;
- event/readback over sleep.

---

# 51. CODE QUALITY

Use `Option Explicit`.

Avoid:

- global hidden state where possible;
- blanket `On Error Resume Next`;
- silent catch after SAP write;
- unbounded loops;
- duplicate implementations of the same business action;
- hard-coded worksheet row numbers as identity;
- SendKeys for production actions.

Centralize action permissions.

---

# 52. CENTRAL ACTION GATE

Create an authoritative function equivalent to:

```text
CanExecuteAction(ActionType, Plan, MapState, HistoryState, ExecutionMode)
```

Do not let multiple legacy flags contradict each other.

Absolute deny in final production path:

- AUTO_CF;
- unsafe coordinate/sendkeys fallback;
- unknown target identity;
- stale plan;
- unreconciled prior action.

---

# 53. ACTION ID EXAMPLES

Stable, deterministic forms such as:

```text
CANCEL|OB|DOC|ITEM|WT

RELEASE|TARGET_SOURCEKEY|HOLDER_WT

POST|OB|DOC|ITEM|MAT|SRCBIN|SRCBATCH|DSTBIN|DSTBATCH|QTY

PRDO_QTY|OB|DOC|ITEM|TARGET_QTY

CREATE|OB|DOC|ITEM|MAT|BIN|BATCH|QTY
```

Hash long IDs if needed but retain the canonical source string in evidence.

---

# 54. TEST HARNESS — BUILD/USE REPLAY MODE

If a replay harness already exists, improve it.

If not, build enough adapter separation to run planner/executor-state regression without touching live SAP.

Modes:

```text
LIVE
REPLAY
STATIC_TEST
```

Planner code should be testable without SAP writes.

---

# 55. OFFLINE TEST SUITE — MUST RUN UNTIL PASS

At minimum:

## Quantity parser

All cases in Section 40.

## Blank/zero

Blank never becomes explicit cancel.

## Group allocation

- 3500 + 4500 → 0 + 8000.
- 500 + 500 → 0 + 1000.
- 300 + 700 → 500 + 500.
- 1000 → 1200.
- 1000 → 800.
- 3-member redistribution.

## Replace

- same total, cancel+replace full allocation.
- change bin.
- change batch.
- change bin+batch.

## Multi-document

- 2 materials;
- multiple items;
- multiple documents;
- multiple OB queue.

## State

- A with blank;
- A with 0;
- A with >0 recreate;
- C hidden/completed.

## Idempotency

- verified Cancel not resent;
- verified POST not resent;
- verified Create not resent;
- UNKNOWN not resent.

## Recovery

- interrupted after Cancel SENT;
- interrupted after PRDO Save;
- interrupted after Create Save;
- interrupted after final verify.

---

# 56. HOLDER RELEASE TEST SUITE

Run:

1. free stock sufficient;
2. shortage 100, one holder 100;
3. shortage 100, holder 40 + 60;
4. holder A excluded;
5. holder C excluded;
6. wrong batch excluded;
7. wrong bin excluded;
8. wrong material excluded;
9. target WT excluded;
10. candidate changes after CHECK;
11. verified A adds credit;
12. UNKNOWN adds no credit;
13. insufficient coverage blocks CHECK;
14. PLAN_ONLY sends zero cancel;
15. PILOT restricted correctly;
16. LIVE_LOCKED gate cannot be entered without required proof.

---

# 57. SESSION / NAVIGATION TESTS

Test/mocks where possible:

- MON already open;
- MON wrong node;
- document already open;
- popup after write;
- busy session;
- session lost;
- multiple sessions;
- recovery once.

Assertions:

- no duplicate MON tab in normal flow;
- no MON reopen per P;
- broker owns MON;
- bounded retries.

---

# 58. CHECK READ-ONLY TEST

Instrument write functions.

Run CHECK scenarios.

Assert zero business writes:

```text
CancelCount = 0
PostSaveCount = 0
PRDOSaveCount = 0
CreateWTCount = 0
CFCount = 0
```

Navigation/filtering needed for reads is allowed.

---

# 59. STALE PLAN TEST

Sequence:

1. Load.
2. enter Actual.
3. CHECK PASS.
4. change Actual.
5. press XỬ LÝ.

Expected:

```text
STOP_PLAN_STALE_RECHECK_REQUIRED
SAP_WRITE_COUNT = 0
```

Repeat with changed holder status and map fingerprint.

---

# 60. FINAL VERIFY TEST

A business run cannot PASS solely because a Save call did not error.

Verify:

Member:

- effective qty;
- expected bin;
- expected batch;
- expected status/WT state.

Group:

```text
Sum(FinalVerifiedMemberQty) = GroupTargetTotal
```

Action ledger:

all write actions are:

```text
VERIFIED
or
SKIPPED_VERIFIED
```

No unresolved:

```text
SENT
UNKNOWN
```

---

# 61. WORKBOOK STRUCTURE TESTS

After sheet migration:

- exactly 5 logical final sheets;
- `INPUT` visible;
- other 4 VeryHidden in user mode;
- no `POSTING` runtime sheet;
- no legacy sheet recreated on open;
- no broken defined names;
- no broken button bindings;
- no broken VBA references;
- no broken XML relationships;
- no invalid merge ranges;
- no `#REF!`;
- no damaged vbaProject;
- package/ZIP CRC valid.

---

# 62. VBA COMPILE

If Excel/VBE is available:

- compile the VBA project;
- fix every compile error;
- rerun until clean.

Do not treat procedure-count balance or source parsing as equivalent to actual VBE compile when VBE is available.

If VBE compile is unavailable, mark it explicitly as an external gate and perform all static checks available.

---

# 63. UI TESTS

Verify:

1. workbook opens on INPUT;
2. only INPUT visible in user mode;
3. 4 buttons visible and correctly bound;
4. button order clear;
5. multi-OB queue usable;
6. worklist does not show C by default;
7. user-edit fields obvious;
8. protected system fields;
9. CHECK state visible;
10. Execute state visible;
11. Release warning visible;
12. Recovery warning persists;
13. freeze panes work;
14. 100+ active rows remain usable;
15. layout works at common desktop Excel size;
16. no old instruction text references merged CHECK+EXECUTE;
17. no obsolete POSTING button/sheet references in user-facing UI.

You may improve typography/layout beyond supplied mockups when these tests remain true.

---

# 64. LIVE TESTING — ONLY IF ENVIRONMENT SUPPORTS IT

If live SAP is available and the supplied materials contain approved/known test targets, follow a safety ladder.

Do not choose arbitrary warehouse documents.

Do not use production data merely because it appears in history.

Prefer exact test cases explicitly identified by the user/package as current test targets.

---

# 65. LIVE LADDER

## Live 1 — Load only

- TẢI WH;
- exact OB/WT read;
- zero write;
- verify one MON session.

## Live 2 — CHECK only

- source read;
- holder read if needed;
- build plan;
- zero business writes.

## Live 3 — simple execution without holder release

Use a known safe case.

Test:

- Cancel own if required;
- Qty/bin/batch path;
- Create as required;
- stop at USER_CF_HANDOFF.

## Live 4 — mixed member group

Include cancel + retain + create delta.

Prefer the net-zero allocation scenario when a safe test target exists.

## Live 5 — holder release PILOT

Only after prior live steps pass.

One exact release group.

## Live 6 — multi-OB

Only after individual group/document flows are proven.

---

# 66. LIVE FAILURE LOOP

When live test fails:

1. stop additional writes if state is uncertain;
2. persist history;
3. reconcile SAP read-only;
4. classify:
   - NOT_SENT;
   - SENT_SUCCESS_VERIFIED;
   - SENT_FAILED_VERIFIED;
   - UNKNOWN;
5. patch only after state is known;
6. rerun from safe state;
7. never resend UNKNOWN.

Continue until the live ladder gates available in the environment pass.

---

# 67. NO “TEST UNTIL SUCCESS” BY CHEATING

Do not make tests pass by:

- changing expected values to actual failures;
- disabling assertions;
- hiding blocker rows;
- changing PASS criteria;
- removing safety gates;
- marking UNKNOWN as PASS;
- reducing test coverage;
- skipping the problematic group;
- hardcoding one OB/material/bin;
- increasing sleeps until timing happens to work.

Fix root causes.

---

# 68. USE HISTORY AS A REGRESSION CORPUS

The supplied workbook/conversation packages contain many previous runs.

Extract real scenario fixtures from them.

Classify at least:

```text
LIVE_PROVEN
FAILED_KNOWN_DEFECT
RECOVERED
UNKNOWN
LEGACY_ONLY
```

Use LIVE_PROVEN cases as do-not-regress tests.

Use FAILED_KNOWN_DEFECT as expected-to-fail-on-V2.6 / expected-to-pass-on-V2.7 tests.

Do not replay live SAP writes merely to reproduce historical evidence.

---

# 69. BUILD A REGRESSION CATALOG

Create:

`V2.7_REGRESSION_CATALOG.md`

Columns/sections:

- TestCaseID;
- source file/run;
- scenario;
- old outcome;
- expected V2.7 outcome;
- offline/live;
- evidence;
- result.

Must include the net-zero multi-member quantity bug.

---

# 70. USE OLDER UI MOCKUPS INTELLIGENTLY

Older mockups may show a polished dark/navy header, left OB queue, status cards and worklist.

Keep good visual ideas such as:

- clear visual hierarchy;
- left OB queue;
- compact summary;
- status cards;
- modern professional palette;
- strong separation of input vs status.

But update them for:

- 4 buttons;
- separate CHECK and XỬ LÝ;
- plan/release/recovery states;
- simplified 5-sheet architecture.

Do not literally preserve outdated labels like:

`CHECK & XỬ LÝ`

or:

`CHECK + POSTING`

on the final user UI.

---

# 71. UI SHOULD BE BETTER THAN THE MD IF THERE IS A GOOD REASON

The MD is an architecture contract, not a pixel prison.

You may improve:

- spacing;
- column order;
- panel dimensions;
- summary cards;
- status badges;
- freeze position;
- button sizing;
- user instructions;
- visual grouping.

But if you diverge materially, document:

- what changed;
- why it is better;
- what acceptance constraint remains preserved.

Create:

`V2.7_UI_DECISIONS.md`

---

# 72. DO NOT OVER-ENGINEER UI

This is Excel/VBA, not a web app.

Prefer maintainable Excel-native UI:

- cells;
- tables;
- shapes;
- conditional formatting;
- frozen panes;
- protection;
- named ranges/tables.

Avoid fragile ActiveX/custom UI dependencies unless the current workbook already relies on them and they are necessary.

---

# 73. MIGRATION PARITY

Before deleting legacy sheets, compare old and new repositories for the same snapshot.

At minimum compare:

- OB list;
- WT count;
- Material;
- Sys Qty;
- Sys Bin;
- Sys Batch;
- Status;
- source stock rows;
- valid available quantities;
- plan member count;
- group totals.

Produce a parity report.

---

# 74. ZERO LEGACY DEPENDENCY AUDIT

Before final sheet deletion search VBA for:

```text
ORDER_QUERY
POSTING_RAW
WH_ACTUAL
POSTING
RUN_LOG
SAP_MAP
CONFIG
OB_QUEUE
OB_WH_CACHE
POSTING_SUMMARY
POSTING_CACHE
PLAN_SOURCE_LEGS
PLAN_RELEASE
SAP_PROBE_LOG
SAP_ACTION_MAP
RUN_SUMMARY
REGRESSION_AUDIT
```

Some names may legitimately remain as migration constants/comments, but final runtime must not depend on deleted sheets.

Document every remaining match.

---

# 75. PUBLIC ENTRY POINTS

Final buttons should call only high-level entry points equivalent to:

```text
CF_LoadWH
CF_Check
CF_Execute
CF_ResetUI
```

Buttons must never directly call low-level SAP routines.

---

# 76. ENTRY POINT REENTRANCY GUARD

Prevent double-click/double-run.

Use an application/workbook execution lock.

If already active:

- show concise status;
- do not start second scheduler;
- do not create second SAP navigation chain.

Always release safe local locks on controlled exit, but preserve recovery state for uncertain SAP writes.

---

# 77. CF_LoadWH CONTRACT

Expected:

```text
Acquire UI lock
Normalize OB queue
Acquire/reuse MON_ROLE
Read all requested OB snapshots
Build DATA_CORE snapshots
Render actionable worklist
Hide C projection
Set PlanState = NOT_CHECKED
Enable CHECK
Disable XỬ LÝ
Release lock
```

No business write.

---

# 78. CF_Check CONTRACT

Expected:

```text
Acquire lock
Validate load snapshot
Validate user input
Capture immutable ActualSnapshot
Read source stock
Read holder candidates
Build member ledger
Build group ledger
Build source legs
Build release plan
Build action plan
Run blocker gates
Compute fingerprints
Persist PlanID
Render preview
Set CHECK result
Release lock
```

Zero business write.

---

# 79. CF_Execute CONTRACT

Expected:

```text
Acquire exclusive execute lock
Load exact PlanID
Verify fingerprints
Verify recovery clear
Fresh precheck only the entities to be written

For action phase:
    validate exact ActionID
    persist precheck
    persist/send boundary
    send once
    read back
    mark VERIFIED or UNKNOWN
    stop on UNKNOWN

Final verify
Persist result
Render USER_CF_HANDOFF
Release lock
```

No automatic CF.

---

# 80. CF_ResetUI CONTRACT

Expected:

```text
If active execute:
    reject

If unresolved UNKNOWN:
    preserve recovery/history
    clear only safe UI areas if allowed

Else:
    clear user queue/worklist/current plan reference
    retain history/system
```

---

# 81. REPORT FAILURE EARLY INTERNALLY, NOT TO USER AS FINAL

During your own work, keep a failure ledger.

When a focused test discovers a defect:

- record it;
- patch it;
- rerun.

Do not produce a final answer while known attainable tests remain failing.

---

# 82. ITERATION LEDGER

Create:

`V2.7_ITERATION_LOG.md`

For each engineering iteration:

```text
Iteration
Change
Reason
Tests run
Failures
Root cause
Patch
Regression result
Artifact hash
```

This is especially important when many Codex conversations are part of the source corpus.

---

# 83. STOP CONDITIONS FOR ENGINEERING LOOP

Normal loop may stop only when:

## OFFLINE_FINAL_PASS

All environment-attainable gates pass:

- source ingest complete;
- conflict ledger complete;
- workbook migration complete;
- 4-button UI;
- 5-sheet final structure;
- compile if available;
- offline regression pass;
- UI pass;
- package integrity pass;
- no known reproducible static/replay defect.

and either:

### LIVE_FINAL_PASS

The available live ladder was executed and passed with evidence,

or:

### BLOCKED_EXTERNAL

A live-only gate cannot be executed because the environment lacks SAP/approved target/etc.

Do not use vague “mostly done”.

---

# 84. FINAL V2.7 FILE NAME

Prefer:

`260820_CF_OB_V2.7.xlsm`

If iterations require build numbers:

```text
260820_CF_OB_V2.7_build01.xlsm
...
```

but final successful candidate should use the clean V2.7 name once stable.

Never overwrite the only baseline copy.

---

# 85. REQUIRED FINAL DELIVERABLES

Return/package:

1. `260820_CF_OB_V2.7.xlsm`
2. `260820_CF_OB_V2.7_PACKAGE.zip`
3. `V2.7_CHANGELOG.md`
4. `V2.7_BASELINE_AUDIT.md`
5. `V2.7_SOURCE_CONFLICT_LEDGER.md`
6. `V2.7_MIGRATION_AUDIT.md`
7. `V2.7_REGRESSION_CATALOG.md`
8. `V2.7_REGRESSION_REPORT.md`
9. `V2.7_UI_DECISIONS.md`
10. `V2.7_ITERATION_LOG.md`
11. `V2.7_FINAL_AUDIT.md`

If LIVE SAP was available:

12. `V2.7_LIVE_TEST_REPORT.md`

---

# 86. FINAL AUDIT MUST INCLUDE

## File integrity

- SHA-256;
- ZIP/CRC;
- XML parse;
- relationship targets;
- vbaProject presence;
- workbook opens if Excel available.

## Workbook

- sheet count = 5;
- only INPUT visible user mode;
- defined names valid;
- shapes valid;
- macro bindings valid.

## VBA

- module list;
- compile PASS if available;
- public entry points;
- no duplicate old button entry paths;
- no runtime legacy sheet references.

## Functional

- CHECK read-only;
- stale-plan stop;
- member allocation;
- net-zero group case;
- replace case;
- holder PLAN_ONLY;
- idempotency;
- recovery;
- no auto CF.

## Performance

- navigation counters;
- no per-P reopen;
- worksheet footprint before/after;
- history event count/noise.

## UI

- screenshot or cell-map review if possible;
- 4 buttons;
- worklist;
- queue;
- status;
- recovery/release warnings.

---

# 87. REQUIRED FINAL STATUS LANGUAGE

Use one of:

```text
FINAL_PASS
```

only when all attainable tests pass and no unreported external gate exists.

Use:

```text
OFFLINE_PASS_LIVE_BLOCKED_EXTERNAL
```

when the engineering/package is complete but live SAP could not be executed.

Use:

```text
LIVE_PILOT_PASS_PRODUCTION_HOLDER_LOCKED
```

when safe live flows passed but holder release remains intentionally gated.

Do not say:

- “100% complete”;
- “live proven”;
- “production ready”;

without evidence matching those exact claims.

---

# 88. FINAL RESPONSE SHOULD BE CONCISE BUT EVIDENCE-BASED

Do not dump the whole analysis.

Report:

- final status;
- output workbook;
- major architecture changes;
- exact test counts;
- any blocked external gate;
- live test status;
- holder release mode;
- link/path to reports.

The detailed reasoning belongs in the generated audit files.

---

# 89. SPECIAL INSTRUCTION ABOUT OLD CODEX CONVERSATIONS

The conversation archive may contain many successive “fixes”.

Do not apply them chronologically as patches.

Instead:

1. extract user intent;
2. extract proven SAP facts;
3. identify repeated root causes;
4. build one coherent V2.7 model;
5. use old failures as regression tests.

The goal is to stop the cycle:

```text
new OB
→ discover new branch failure
→ patch one condition
→ next OB fails elsewhere
```

V2.7 must fix the model.

---

# 90. SPECIAL INSTRUCTION ABOUT “TEST UNTIL SUCCESS”

“Until success” means:

- keep fixing reproducible defects;
- rerun focused test;
- rerun broader regression;
- do not stop after the first failing run;
- do not require the user to repeatedly tell you “continue”.

It does **not** mean:

- repeat unsafe SAP writes until one happens to succeed;
- retry UNKNOWN actions;
- bypass safety;
- use arbitrary live OBs;
- auto CF.

A SAP write with unknown result must move to reconciliation, not brute-force retry.

---

# 91. FINAL NON-NEGOTIABLE CHECKLIST

Before declaring V2.7 final, answer YES to all attainable items:

## Sources

- [ ] All supplied archives recursively inventoried.
- [ ] Conversation packages parsed.
- [ ] Authority conflicts resolved.
- [ ] Latest 4-button spec treated as authoritative.

## Workbook

- [ ] Current V2.6 baseline preserved.
- [ ] Final target has 5 logical sheets.
- [ ] POSTING runtime sheet removed.
- [ ] Only INPUT visible in user mode.

## UI

- [ ] TẢI WH separate.
- [ ] CHECK separate.
- [ ] XỬ LÝ separate.
- [ ] LÀM MỚI separate.
- [ ] CHECK read-only.
- [ ] multi-OB.
- [ ] C hidden from actionable worklist.
- [ ] plan/release/recovery visible.

## Planner

- [ ] Member allocation ledger.
- [ ] Group ledger.
- [ ] Action ledger.
- [ ] net-zero reallocation bug passes.
- [ ] replace case does not regress.

## SAP navigation

- [ ] MON Broker.
- [ ] no per-P MON reopen.
- [ ] bounded recovery.
- [ ] exact Technical ID path.

## Holder release

- [ ] exact candidate reader.
- [ ] deterministic selector.
- [ ] stale-plan protection.
- [ ] readback A required.
- [ ] no credit on UNKNOWN.
- [ ] final default safely gated.

## History

- [ ] unified durable event ledger.
- [ ] SENT/VERIFIED/UNKNOWN explicit.
- [ ] before/target/after.
- [ ] fingerprints.
- [ ] resume state.
- [ ] no polling spam.

## Safety

- [ ] no auto CF.
- [ ] no resend UNKNOWN.
- [ ] no execute stale plan.
- [ ] blank != 0.
- [ ] quantity parser regression pass.

## Quality

- [ ] VBA compile pass if Excel/VBE available.
- [ ] offline regression pass.
- [ ] UI checks pass.
- [ ] package integrity pass.
- [ ] live gates either evidence-backed PASS or explicitly BLOCKED_EXTERNAL.

---

# 92. START NOW

Do not merely restate this prompt.

Begin by:

1. recursively inventorying all supplied files/ZIPs;
2. identifying the current V2.6 binary;
3. locating the latest 4-button/5-sheet V2.7 MD;
4. locating packaged Codex/Claude conversation histories;
5. generating the source conflict ledger;
6. taking a baseline snapshot;
7. then implementing V2.7 in controlled checkpoints;
8. testing after each checkpoint;
9. continuing the fix-test-regression loop until the defined final status is reached.

**The final outcome must be an implemented, tested artifact package — not only a design document.**
