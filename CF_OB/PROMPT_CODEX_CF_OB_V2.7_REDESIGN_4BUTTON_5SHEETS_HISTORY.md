# PROMPT CODEX — REDESIGN `260819_CF_OB_V2.6_LICH_SU` → `CF_OB V2.7`
## 4 NÚT TÁCH RỜI • 5 SHEET • KHÔNG CÒN SHEET POSTING • GROUP/MEMBER LEDGER • CANCEL HOLDER CÓ GATE • HISTORY V2 CHI TIẾT

> **File nguồn:** `260819_CF_OB_V2.6_LICH_SU.xlsm`
>
> **Mục tiêu cuối:** nâng cấp trực tiếp V2.6 thành V2.7 ổn định, dễ dùng, dễ debug, không tiếp tục vá lỗi theo từng OB.  
> V2.7 phải sửa đồng thời: kiến trúc CHECK/EXECUTE, group/member allocation, session MON, release/cancel WT khác giữ tồn, cấu trúc sheet, UI Control Center và lịch sử chạy.
>
> **Không rebuild từ workbook trắng. Không phá các luồng V2.6 đã LIVE proven. Không auto CF.**

---

# 0. THAY ĐỔI YÊU CẦU SO VỚI PROMPT V2.7 TRƯỚC

Prompt trước yêu cầu 3 nút:

`TẢI WH -> CHECK & XỬ LÝ -> LÀM MỚI`

**BỎ THIẾT KẾ NÀY.**

V2.7 mới bắt buộc tách CHECK và XỬ LÝ vì nhánh **Cancel WH/WT khác giữ tồn chưa đủ ổn định để cho phép một nút vừa phân tích vừa ghi SAP**.

## UI cuối cùng phải có đúng 4 nút vận hành

1. **TẢI WH**
2. **CHECK**
3. **XỬ LÝ**
4. **LÀM MỚI**

Flow:

`TẢI WH -> USER NHẬP ACTUAL -> CHECK READ-ONLY -> REVIEW PLAN -> XỬ LÝ LOCKED PLAN -> USER CF`

### Nguyên tắc tuyệt đối

- `CHECK` = **READ ONLY**.
- `CHECK` không Cancel WT.
- `CHECK` không Change Qty.
- `CHECK` không Change Bin.
- `CHECK` không Change Batch.
- `CHECK` không POST.
- `CHECK` không Create WT.
- `CHECK` không Save SAP document.
- `CHECK` không CF.
- `CHECK` được phép:
  - đọc MON;
  - đọc nguồn tồn;
  - đọc holder WT;
  - lập PLAN;
  - lập RELEASE PLAN;
  - tính group allocation;
  - đánh giá blocker;
  - khóa PlanFingerprint.

`XỬ LÝ` mới là entry point duy nhất được phép ghi SAP.

Nếu CHECK chưa PASS hoặc plan đã stale thì nút XỬ LÝ phải bị disable/hard-stop.

---

# 1. INVENTORY V2.6 HIỆN TẠI — KHÔNG ĐƯỢC BỎ QUA

Workbook V2.6 hiện có **20 sheet**.

| Sheet | Visibility | Used range hiện tại | Vai trò hiện tại |
|---|---:|---:|---|
| ORDER_QUERY | VeryHidden | A1:F5 | request Phase 2 |
| POSTING_RAW | VeryHidden | A1:Q253 | raw source-stock capture |
| INPUT | Visible | A1:N506 | UI |
| WH_ACTUAL | Visible | A1:AF504 | MON + actual + decision |
| POSTING | Visible | A1:AE500 | bảng source stock trung gian cũ |
| RUN_LOG | Visible | A1:X4590 | log người đọc |
| SAP_MAP | Visible | A1:P1747 | map SAP |
| CONFIG | Visible | A1:F261 | config |
| OB_QUEUE | Visible | A1:X2 | queue |
| OB_WH_CACHE | Visible | A1:Z7 | WH snapshot cache |
| POSTING_SUMMARY | VeryHidden | A1:T5 | capture summary |
| POSTING_CACHE | VeryHidden | A1:S23 | normalized source stock |
| PLAN_SOURCE_LEGS | VeryHidden | A1:G3 | source legs |
| PLAN_RELEASE | VeryHidden | A1:U1 | release candidates |
| SAP_PROBE_LOG | VeryHidden | A1:Q1 | probe |
| PLAN | Visible | A1:AZ7 | plan |
| SAP_ACTION_MAP | Visible | A1:J28 | action map |
| RUN_HISTORY | Visible | A1:AA3304 | durable history |
| RUN_SUMMARY | Visible | A1:N101 | run summary |
| REGRESSION_AUDIT | Visible | A1:I3244 | regression |

### Dấu hiệu cần tối ưu

- `POSTING` có khoảng **500 x 31 = 15.500 cell** nhưng dữ liệu gốc/normalized đã tồn tại ở `POSTING_RAW`, `POSTING_CACHE`, `POSTING_SUMMARY`.
- `RUN_LOG` khoảng 4.588 dòng.
- `RUN_HISTORY` khoảng 3.304 dòng.
- `REGRESSION_AUDIT` khoảng 3.244 dòng.
- History đang bị chia nhỏ thành nhiều nguồn, khó biết một lần chạy đã:
  - check gì;
  - gửi SAP write nào;
  - verify được gì;
  - dừng đúng action nào;
  - lần chạy sau phải resume từ đâu.

**V2.7 phải giảm từ 20 sheet xuống mục tiêu 5 sheet cuối cùng.**

---

# 2. KIẾN TRÚC 5 SHEET CUỐI CÙNG

Final V2.7 phải hướng tới **đúng 5 sheet logic**:

1. `INPUT`
2. `DATA_CORE`
3. `PLAN`
4. `SYSTEM`
5. `RUN_HISTORY`

## Visibility cuối

- `INPUT` = Visible.
- `DATA_CORE` = VeryHidden.
- `PLAN` = VeryHidden.
- `SYSTEM` = VeryHidden.
- `RUN_HISTORY` = VeryHidden trong User Mode.
- Debug/Admin mode mới được unhide sheet kỹ thuật.

Không để user bình thường nhìn thấy hàng loạt sheet engine.

---

# 3. MIGRATION 20 SHEET → 5 SHEET

## 3.1. `INPUT`

Giữ technical name `INPUT`.

Vai trò duy nhất:

- danh sách OB;
- vùng Actual input;
- trạng thái CHECK;
- trạng thái EXECUTE;
- preview plan;
- cảnh báo;
- 4 nút.

Không dùng INPUT làm cache engine dài hạn.

---

## 3.2. `DATA_CORE`

Tạo mới hoặc repurpose thành sheet dữ liệu backend duy nhất.

Gộp vai trò của:

- `ORDER_QUERY`
- `POSTING_RAW`
- `WH_ACTUAL` backend
- `POSTING`
- `OB_QUEUE`
- `OB_WH_CACHE`
- `POSTING_SUMMARY`
- `POSTING_CACHE`

### Sau migration

`POSTING` **không còn tồn tại như runtime sheet**.

Không được giữ logic:

`SAP -> POSTING_RAW -> POSTING -> POSTING_CACHE`

Thiết kế lại thành:

`SAP capture -> normalize in memory -> DATA_CORE.SourceStock`

Nếu cần giữ raw evidence:

`SAP capture -> DATA_CORE.SourceStockRaw`

nhưng không cần một sheet riêng tên POSTING.

### Các table logic trong DATA_CORE

#### A. `tblOBQueue`

Các cột:

- QueueID
- LoadRunID
- InputSeq
- OB
- RequestedAt
- ReturnedAt
- QueueState
- OBState
- DocumentCount
- WTCount
- ReadyCount
- BlockedCount
- Message

#### B. `tblWHSnapshot`

Một row = một WT snapshot.

Các cột tối thiểu:

- SnapshotID
- LoadRunID
- OB
- Document
- Item
- WH
- WT
- Material
- SysQty
- SysBin
- SysBatch
- WTStatus
- OBState
- RemainingQty
- SourceKey
- CapturedAt
- CaptureEvidence
- RowFingerprint

#### C. `tblActualSnapshot`

Snapshot của input user tại thời điểm CHECK.

Không dùng row number làm identity.

Các cột:

- CheckID
- SourceKey
- OB
- Document
- Item
- WT
- Material
- ActualQty
- Area
- ActualBin
- ActualBatch
- InputHash
- CapturedAt

#### D. `tblSourceStock`

Thay thế vai trò chính của `POSTING` và `POSTING_CACHE`.

Các cột:

- StockID
- CheckID
- Material
- SourceBin
- Batch
- StockType
- AvailableQty
- TotalStock
- GRDate
- GRTime
- ProcessType
- CaptureState
- IsUsable
- AllocatedQty
- RemainingAllocatableQty
- EvidenceID

#### E. `tblSourceStockRaw`

Chỉ giữ raw evidence cần thiết từ SAP.

Không format đẹp.
Không dùng user-facing.
Không clone 500 blank rows.

#### F. `tblCaptureSummary`

Thay `POSTING_SUMMARY`.

Các cột:

- CheckID
- Material
- ExpectedRows
- CapturedRawRows
- ValidRows
- InvalidRows
- DataReady
- CaptureFingerprint
- Message

#### G. `tblRuntimeState`

Key/value nhỏ:

- CurrentLoadRunID
- CurrentCheckID
- CurrentPlanID
- CurrentExecuteID
- CurrentPlanFingerprint
- CurrentInputFingerprint
- CurrentMapFingerprint
- LastCompletedPhase
- RecoveryRequired
- RecoveryReason

---

# 4. `PLAN` — MỘT SHEET NHƯNG NHIỀU TABLE CÓ HỢP ĐỒNG RÕ

Gộp:

- `PLAN`
- `PLAN_SOURCE_LEGS`
- `PLAN_RELEASE`

Không để ba sheet riêng.

## 4.1. `tblMemberPlan`

Một row = một member/WT cần quyết định.

Các cột bắt buộc:

- PlanID
- PlanSeq
- SourceKey
- OB
- Document
- Item
- WH
- WT
- Material
- QtyGroupKey
- ActionGroupKey
- SysQty
- TargetQty
- SysBin
- TargetBin
- SysBatch
- TargetBatch
- ExistingEffectiveQty
- CancelQty
- RetainQty
- ReplaceQty
- RequiredCreateQty
- RequiredReleaseQty
- SourceAvailableQty
- ReleaseCoveredQty
- ExecMode
- Decision
- PrecheckState
- BlockCode
- Message
- MemberFingerprint

### Không dùng duy nhất GroupNeedQty để quyết định member

V2.7 bắt buộc có:

`RequiredCreateQty = max(TargetQty - EffectiveQtyAfterPlannedCancelReplace, 0)`

---

## 4.2. `tblGroupPlan`

Một row = một group quantity/business group.

Các cột:

- PlanID
- QtyGroupKey
- OB
- Document
- Material
- MemberCount
- GroupCurrentTotal
- GroupTargetTotal
- GroupNetDelta
- GroupCancelTotal
- GroupRetainTotal
- GroupReplaceTotal
- GroupCreateTotal
- GroupReleaseNeed
- PRDOAggregateWriteRequired
- AllocationWriteRequired
- GroupState
- GroupFingerprint

### Quy tắc quan trọng nhất V2.7

`GroupNetDelta = 0`

**KHÔNG được suy ra**

`AllocationWriteRequired = FALSE`.

Ví dụ:

- member A: 3500 -> 0
- member B: 4500 -> 8000
- group current = 8000
- group target = 8000
- GroupNetDelta = 0

Nhưng:

- CancelTotal = 3500
- RetainTotal = 4500
- CreateTotal = 3500
- PRDO aggregate write có thể = FALSE
- Allocation write = TRUE

Expected:

`CANCEL 3500 -> RETAIN 4500 -> CREATE 3500 -> FINAL 8000`

---

## 4.3. `tblSourceLeg`

Thay `PLAN_SOURCE_LEGS`.

Các cột:

- PlanID
- SourceLegID
- SourceKey
- Material
- Bin
- Batch
- NeedQty
- AvailableQty
- ReservedQty
- ReleaseCredit
- FinalUsableQty
- SourceState
- EvidenceID

---

## 4.4. `tblReleaseCandidate`

Thay `PLAN_RELEASE`.

Các cột:

- PlanID
- ReleasePlanID
- TargetSourceKey
- Material
- Bin
- Batch
- RequiredReleaseQty
- HolderOB
- HolderDocument
- HolderWH
- HolderWT
- HolderQty
- HolderStatus
- CandidateRank
- CandidateEligible
- Selected
- SelectedQty
- CumulativeSelectedQty
- CoverageState
- RejectReason
- ReadEvidenceID
- CandidateFingerprint

---

## 4.5. `tblActionPlan`

Đây là bảng rất quan trọng, cần bổ sung.

Một row = một SAP action dự kiến.

Các cột:

- PlanID
- ActionSeq
- ActionID
- ParentGroupKey
- SourceKey
- ActionType
- TargetEntity
- BeforeState
- TargetState
- RequiresWrite
- RequiresReadback
- CanRetry
- StopOnUnknown
- Precondition
- ExpectedEvidence
- ActionState

ActionType chuẩn hóa:

- CANCEL_OWN_WT
- CANCEL_HOLDER_WT
- POST_CHANGE
- PRDO_QTY_WRITE
- PRDO_BIN_WRITE
- PRDO_BATCH_WRITE
- PRDO_SAVE
- CREATE_WT
- MON_READBACK
- FINAL_VERIFY
- USER_CF_HANDOFF

---

# 5. `SYSTEM` — GỘP CONFIG + SAP MAP + ACTION MAP + PROBE

Gộp:

- `CONFIG`
- `SAP_MAP`
- `SAP_ACTION_MAP`
- `SAP_PROBE_LOG`

## Các vùng/table

### `tblConfig`

Không dùng tọa độ cell hard-coded khắp code.
Dùng key lookup.

Các key mới:

- UI_MODE
- HISTORY_DETAIL_LEVEL
- RELEASE_HOLDER_MODE
- RELEASE_ALLOW_OTHER_OB
- RELEASE_ALLOW_SAME_OB_UNPLANNED
- RELEASE_EXCLUDE_STATUS
- RELEASE_REQUIRE_MON_READBACK
- RELEASE_REQUIRE_WHOLE_WT
- RELEASE_MAX_CANDIDATES
- RELEASE_MAX_TOTAL_WT
- SAP_MON_REUSE_REQUIRED
- SAP_WRITE_UNKNOWN_STOP
- PLAN_FINGERPRINT_REQUIRED
- INPUT_FINGERPRINT_REQUIRED
- CF_USER_HANDOFF_REQUIRED

### `RELEASE_HOLDER_MODE`

Chỉ cho phép:

- `PLAN_ONLY`
- `PILOT`
- `LIVE_LOCKED`

Default V2.7 khi xuất file:

`PLAN_ONLY`

#### PLAN_ONLY

CHECK:

- tìm candidate;
- select candidate;
- tính coverage;
- ghi plan.

XỬ LÝ:

- không Cancel holder;
- nếu action cần holder release:
  - dừng `WAIT_RELEASE_HOLDER_NOT_ARMED`;
  - user review.

#### PILOT

- chỉ một release group;
- chỉ một OB;
- explicit confirmation;
- cancel xong bắt buộc MON readback;
- nếu không chứng minh A -> UNKNOWN hard stop.

#### LIVE_LOCKED

Chỉ cho bật sau khi regression/live proof đạt điều kiện do V2.7 ghi trong history.

### `tblSapMap`

Thay SAP_MAP.

Bắt buộc có:

- MapKey
- Transaction
- ScreenRole
- TechnicalID
- ControlType
- ColumnKey
- ParentID
- VerifiedState
- MapVersion
- MapFingerprint
- VerifiedAt

### `tblSapActionMap`

Thay SAP_ACTION_MAP.

### `tblSapProbe`

Chỉ lưu probe quan trọng.
Không log polling loop hàng trăm lần.

---

# 6. `RUN_HISTORY` — THIẾT KẾ LẠI HOÀN TOÀN

Không tiếp tục dùng bốn sheet:

- RUN_LOG
- RUN_HISTORY
- RUN_SUMMARY
- REGRESSION_AUDIT

Tất cả gộp về **một durable event ledger**.

## 6.1. Mục tiêu

Một lần nhìn history phải trả lời được:

1. user đã Load WH lúc nào?
2. CHECK nào được tạo?
3. input lúc CHECK là gì?
4. plan nào được lock?
5. group nào có những member nào?
6. action nào dự kiến chạy?
7. XỬ LÝ đã thực sự gửi action nào?
8. trước write SAP đang ở trạng thái gì?
9. sau write đọc lại được gì?
10. action nào PASS?
11. action nào UNKNOWN?
12. action nào chưa từng gửi?
13. tool dừng tại phase nào?
14. có thể resume không?
15. resume từ action nào?
16. lần sau có resend action cũ hay không?
17. SAP session/tab nào đã được dùng?
18. mất bao nhiêu thời gian cho MON / POST / PRDO / Create?
19. version/map fingerprint/config nào được dùng?

---

# 7. HISTORY EVENT SCHEMA

Tạo `tblRunEvents`.

Một event = một transition có ý nghĩa.
Không log mỗi vòng polling.

Các cột:

1. EventSeq
2. Timestamp
3. WorkbookVersion
4. EventType
5. RunType
6. LoadRunID
7. CheckID
8. PlanID
9. ExecuteID
10. ActionID
11. ParentActionID
12. ActionSeq
13. Stage
14. Phase
15. OB
16. Document
17. Item
18. WH
19. WT
20. Material
21. QtyGroupKey
22. SourceKey
23. ActionType
24. ActionState
25. BeforeQty
26. TargetQty
27. AfterQty
28. BeforeBin
29. TargetBin
30. AfterBin
31. BeforeBatch
32. TargetBatch
33. AfterBatch
34. RequiredCreateQty
35. RequiredReleaseQty
36. AvailableBefore
37. ReleasedQty
38. AvailableAfter
39. SAPConnection
40. SAPSession
41. SAPTransaction
42. SAPWindowTitle
43. SAPControlKey
44. SAPTechnicalID
45. SAPStatusType
46. SAPStatusText
47. EvidenceType
48. EvidenceValue
49. EvidenceFingerprint
50. PlanFingerprint
51. InputFingerprint
52. MapFingerprint
53. ConfigFingerprint
54. DurationMs
55. Result
56. StopCode
57. RetryAllowed
58. ResumeState
59. ResumeFromActionID
60. ErrorNumber
61. ErrorSource
62. ErrorDescription
63. Message
64. UserName
65. MachineName

### EventType chuẩn

- LOAD_START
- LOAD_ROW
- LOAD_COMPLETE
- CHECK_START
- INPUT_SNAPSHOT
- MON_READ
- SOURCE_STOCK_READ
- HOLDER_READ
- PLAN_MEMBER
- PLAN_GROUP
- PLAN_RELEASE
- PLAN_ACTION
- CHECK_PASS
- CHECK_BLOCKED
- EXEC_START
- ACTION_PRECHECK
- ACTION_SENT
- ACTION_READBACK
- ACTION_VERIFIED
- ACTION_UNKNOWN
- ACTION_SKIPPED_VERIFIED
- EXEC_BLOCKED
- EXEC_COMPLETE
- USER_CF_HANDOFF
- RECOVERY_REQUIRED
- RECOVERY_RESOLVED
- REGRESSION_PASS
- REGRESSION_FAIL

---

# 8. RUN SUMMARY KHÔNG CẦN SHEET RIÊNG

Trên đầu `RUN_HISTORY` hoặc build bằng VBA memory summary:

mỗi CheckID / ExecuteID có:

- StartAt
- EndAt
- OBCount
- MemberCount
- GroupCount
- ActionCount
- SentCount
- VerifiedCount
- UnknownCount
- BlockedCount
- ReleaseCandidateCount
- ReleaseSelectedCount
- CancelOwnCount
- CancelHolderCount
- CreateCount
- TotalDurationMs
- SAPDurationMs
- ExcelDurationMs
- FinalResult
- FinalStopCode
- ResumeState

Không cần `RUN_SUMMARY` sheet.

---

# 9. REGRESSION AUDIT KHÔNG CẦN SHEET RIÊNG

Regression ghi vào cùng `tblRunEvents`:

- EventType = REGRESSION_PASS / REGRESSION_FAIL
- WorkbookVersion
- TestCaseID
- Expected
- Actual
- Result
- EvidenceFingerprint

Có thể tạo summary tạm thời trên INPUT khi Debug Mode.
Không cần sheet `REGRESSION_AUDIT`.

---

# 10. HISTORY PHẢI CHI TIẾT HƠN NHƯNG KHÔNG ĐƯỢC PHÌNH VÔ HẠN

Không ghi:

- mỗi DoEvents;
- mỗi Busy poll;
- mỗi 100 ms retry;
- cùng một status bar text lặp 20 lần.

Chỉ ghi:

- state transition;
- SAP write boundary;
- SAP readback;
- timeout;
- decision;
- proof;
- stop/resume.

## History compaction

Không tự xóa lịch sử nghiệp vụ.

Cho phép chức năng bảo trì riêng:

`ArchiveHistoryBefore(Date)`

nhưng:

- không chạy tự động;
- không gọi khi XỬ LÝ;
- không xóa run có UNKNOWN chưa resolved;
- không xóa evidence của regression hiện tại.

---

# 11. UI V2.7 — CONTROL CENTER MỚI

Sheet `INPUT` là sheet user duy nhất.

## 11.1. Header

A1: `CF OB — CONTROL CENTER V2.7`

Ngay dưới:

- Version
- SAP State
- MON State
- Current LoadRunID
- Current CheckID
- Current PlanID
- Plan State
- Execute State
- Recovery State

---

# 12. UI — 4 NÚT BẮT BUỘC

## Nút 1 — `TẢI WH`

Chức năng:

- đọc danh sách nhiều OB;
- reuse một MON session;
- load WH snapshot;
- đồng bộ UI.

Không đọc source POST nếu chưa cần.
Không lập final plan.
Không write SAP.

Sau TẢI WH:

- CHECK enabled.
- XỬ LÝ disabled.
- PlanState = `NOT_CHECKED`.

---

## Nút 2 — `CHECK`

Đây là nút mới tách riêng.

### CHECK phải làm

1. validate input.
2. freeze Actual input snapshot.
3. đọc source stock.
4. build member plan.
5. build group plan.
6. build source legs.
7. tìm holder WT nếu thiếu tồn.
8. build release candidate plan.
9. build action plan.
10. chạy blocker gates.
11. tính fingerprint.
12. render Preview trên INPUT.

### CHECK tuyệt đối không write SAP

Tạo helper guard:

`AssertCheckIsReadOnly()`

Mọi routine write SAP phải yêu cầu:

`ExecutionContext.Mode = EXECUTE`

Nếu Mode = CHECK:

hard stop trước `.Press`, `.ModifyCell`, `.Text=`, `.Key=`, Save, Cancel, Create.

Exception chỉ cho phép navigation/read-only filter cần thiết để đọc dữ liệu.

### CHECK result

Có ba trạng thái:

- `CHECK_PASS`
- `CHECK_PASS_WITH_RELEASE_PLAN`
- `CHECK_BLOCKED`

`CHECK_PASS_WITH_RELEASE_PLAN` không có nghĩa đã Cancel holder.

---

## Nút 3 — `XỬ LÝ`

Chỉ enabled khi:

- CheckID tồn tại;
- PlanID tồn tại;
- PlanState = LOCKED;
- InputFingerprint hiện tại = Check InputFingerprint;
- WH snapshot không stale;
- MapFingerprint không đổi;
- không RecoveryRequired unresolved.

### Khi bấm XỬ LÝ

Không rebuild plan âm thầm.

Không tự chạy CHECK lại.

Nếu khác fingerprint:

`STOP_PLAN_STALE_RECHECK_REQUIRED`

User phải bấm CHECK lại.

### XỬ LÝ chạy theo `tblActionPlan`

Scheduler order:

1. EXEC_PRECHECK
2. CANCEL_OWN_GROUP
3. MON_READBACK
4. RELEASE_HOLDER_GROUP
5. MON_READBACK
6. POST_CHANGE_GROUP
7. PRDO_DOCUMENT_GROUP
8. CREATE_GROUP
9. FINAL_MON_VERIFY
10. USER_CF_HANDOFF

---

## Nút 4 — `LÀM MỚI`

Chỉ reset UI/current session scope.

Không:

- xóa history;
- xóa UNKNOWN evidence;
- xóa action checkpoint;
- xóa recovery state.

Nếu có unresolved write:

hiện:

`RECOVERY_REQUIRED — không thể làm mới engine state trước khi reconcile.`

Cho phép clear vùng input khi an toàn.

---

# 13. UI LAYOUT

## Bên trái — OB QUEUE

Cột A:B hoặc A:C.

Hỗ trợ nhiều OB:

- paste list;
- không bắt user xử lý từng OB;
- loại blank;
- normalize text;
- phát hiện duplicate;
- không tự thêm OB mặc định.

Cột:

- OB
- Queue State
- Message

---

## Giữa — WORKLIST

Các cột đề xuất:

- OB
- Material
- WT
- Sys Qty
- Sys Bin
- Sys Batch
- Actual Qty
- Area
- Actual Bin
- Actual Batch
- Decision
- Release Need
- Plan State
- Execute State
- Message

### Editable

Chỉ:

- Actual Qty
- Area
- Actual Bin
- Actual Batch

Các cột system locked.

---

# 14. STATUS C — UI PROJECTION

User yêu cầu dòng C không cần xuất hiện trong danh sách cần xử lý.

Không xóa history.
Không xóa WT snapshot.

Tạo projection:

`UI Worklist = current rows where WTStatus <> C and ActionState not final`

Nếu C:

- backend vẫn có;
- history vẫn có;
- UI ẩn.

Không dùng `EntireRow.Hidden` bừa bãi vì có thể ẩn cả OB queue/control area.

---

# 15. UI — CHECK PREVIEW PANEL

Sau CHECK phải hiển thị:

- Total OB
- Total active members
- CF only
- Cancel own WT
- Increase Qty
- Decrease Qty
- Change Bin
- Change Batch
- Replace
- Release required
- Release candidates
- Release coverage
- Blocked
- Ready to Execute

### Nếu có release holder

Hiện warning nổi bật:

`CÓ WH/WT KHÁC GIỮ TỒN — CHECK CHỈ LẬP PLAN, CHƯA CANCEL.`

Nếu `RELEASE_HOLDER_MODE=PLAN_ONLY`:

nút XỬ LÝ vẫn có thể chạy các group không phụ thuộc release nếu scheduler bảo đảm isolation.

Group cần holder release phải dừng trước write phụ thuộc.

Không để một group release chưa armed làm tool gửi partial unsafe write rồi mới phát hiện thiếu tồn.

---

# 16. BUTTON STATE MACHINE

## Initial

- TẢI WH: enabled
- CHECK: disabled
- XỬ LÝ: disabled
- LÀM MỚI: enabled

## Sau TẢI WH PASS

- TẢI WH: enabled
- CHECK: enabled
- XỬ LÝ: disabled

## Sau user edit

- CHECK: enabled
- XỬ LÝ: disabled
- plan cũ = STALE

## Sau CHECK PASS

- CHECK: enabled
- XỬ LÝ: enabled

## Sau CHECK BLOCKED

- XỬ LÝ: disabled

## Trong XỬ LÝ

- TẢI WH: disabled
- CHECK: disabled
- XỬ LÝ: disabled
- LÀM MỚI: disabled

## Sau XỬ LÝ complete

- CHECK có thể enabled để refresh.
- XỬ LÝ disabled cho tới CHECK mới.

---

# 17. BUG GROUP QUANTITY BẮT BUỘC FIX

Case hiện tại:

`STD20000556`

- WT 5042936844: 3500 -> 0
- WT 5042936845: 4500 -> 8000

Current total = 8000.
Target total = 8000.

V2.6 đã suy ra:

`PRDO_QTY_ALREADY_TARGET -> NO_QTY_WRITE`

rồi bỏ mất create delta.

### V2.7 expected

Member ledger:

#### A

- SysQty = 3500
- Target = 0
- CancelQty = 3500
- Retain = 0
- Create = 0

#### B

- SysQty = 4500
- Target = 8000
- CancelQty = 0
- Retain = 4500
- RequiredCreateQty = 3500

Group ledger:

- CurrentTotal = 8000
- TargetTotal = 8000
- NetDelta = 0
- CancelTotal = 3500
- RetainTotal = 4500
- CreateTotal = 3500
- PRDOAggregateWriteRequired = FALSE hoặc tùy exact SAP contract
- AllocationWriteRequired = TRUE

### Hard regression assertion

Không được PASS group nếu:

`sum(FinalVerifiedMemberQty) <> GroupTargetTotal`

và không được coi member complete nếu:

`RequiredCreateQty > 0` nhưng chưa có verified created/reacquired quantity.

---

# 18. GROUP SCHEDULER — KHÔNG CHẠY TỪNG P000x RỜI RẠC

V2.7 phải xây plan hoàn chỉnh trước.

Không:

`P0001 -> open OB -> action -> close -> P0002 -> reopen OB`

Mà:

`OB -> Document -> Groups -> Action Ledger`

## Scheduler hierarchy

- Run
  - OB
    - Document
      - CancelGroup
      - ReleaseGroup
      - PostChangeGroup
      - PRDOGroup
      - CreateGroup
      - FinalVerifyGroup

Mỗi document chỉ Save khi cần và theo hợp đồng one-save/document group.

---

# 19. SINGLE MON BROKER

Hiện V2.6 vẫn reopen/refresh MON nhiều.

V2.7 tạo:

`modSapSessionBroker`
hoặc equivalent.

## Contract

- một owner duy nhất của MON session;
- reader khác không tự `/SCWM/MON`;
- release-holder reader dùng cùng broker;
- cancel readback dùng cùng broker;
- final verify dùng cùng broker.

### Không module nào được trực tiếp

- tạo tab MON mới;
- gửi `/N/SCWM/MON`;
- reopen outbound node;

trừ Broker recovery.

### Metrics phải log

- MonOpenCount
- MonRefreshCount
- MonRecoveryCount
- DocumentOpenCount
- WTActionOpenCount

Expected normal run một OB:

- MonOpenCount <= 1
- MonRefreshCount khoảng 2–4 tùy write phase
- recovery chỉ khi chứng minh context lost.

---

# 20. CANCEL OWN WT

Own planned WT cancellation đã có nhiều live evidence hơn holder cancel.

Vẫn cần:

- group cancel;
- một selection;
- một Cancel action;
- one group readback.

Không cancel từng member riêng nếu cùng cancel group.

### Proof

Sau Cancel:

MON phải chứng minh exact WT:

`Status = A`

Nếu không:

- `ACTION_UNKNOWN`
- no retry
- stop/resume required.

---

# 21. CANCEL HOLDER WT — THIẾT KẾ LẠI

Đây là phần chưa ổn định.

Không được nhảy thẳng từ "thiếu tồn" sang cancel WT khác.

## Formula

`RequiredReleaseQty = max(RequiredCreateQty - FreeAvailableQty - VerifiedOwnCancelCredit, 0)`

Chỉ chạy holder query nếu `RequiredReleaseQty > 0`.

---

# 22. HOLDER QUERY

Filter exact:

- Material
- Bin
- Batch

Candidate phải đọc được:

- OB
- Document
- WH
- WT
- Qty
- Status
- Source Bin
- Batch
- Material

### Exclude

- status A
- status C
- current target WT
- newly created WT
- candidate không exact Material/Bin/Batch
- candidate không có stable WT identity

Default:

`RELEASE_ALLOW_SAME_OB_UNPLANNED = FALSE`

---

# 23. RELEASE SELECTOR

Deterministic whole-WT selector.

Ưu tiên:

1. exact source
2. status active
3. smallest combination đủ RequiredReleaseQty
4. ít WT nhất
5. stable tie breaker theo WT ID

Không partial cancel nếu:

`RELEASE_REQUIRE_WHOLE_WT=TRUE`

Nếu không đủ:

`STOP_RELEASE_COVERAGE_INSUFFICIENT`

CHECK phải thấy blocker này trước XỬ LÝ.

---

# 24. RELEASE EXECUTION

Chỉ trong nút XỬ LÝ.

Trước Cancel:

- re-read candidates;
- verify exact WT;
- verify status vẫn active;
- verify qty không đổi;
- verify fingerprint.

Nếu đổi:

`STOP_RELEASE_PLAN_STALE_RECHECK_REQUIRED`

Không tự chọn candidate mới trong EXECUTE.

User phải CHECK lại.

### Sau Cancel holder

MON exact readback bắt buộc.

Chỉ khi exact holder WT = A:

`ReleasedQty` mới được cộng vào available credit.

Không cộng credit ngay sau `.Press Cancel`.

---

# 25. RELEASE ACTION UNKNOWN

Nếu SAP action đã gửi nhưng readback không chứng minh outcome:

- state = UNKNOWN
- no retry
- no alternate candidate
- no next write depending on released stock
- write history
- RecoveryRequired = TRUE

---

# 26. PRDO QUANTITY

Tách:

- aggregate PRDO total
- member allocation.

`PRDO_QTY_ALREADY_TARGET`

chỉ cho phép kết luận:

`No aggregate PRDO quantity write required.`

Không được kết luận:

`No member create/reallocation required.`

---

# 27. CREATE WT

CreateQty phải lấy từ member ledger.

### KEEP_EXISTING_CREATE_DELTA

`CreateQty = TargetQty - VerifiedExistingEffectiveQty`

Không lấy:

- GroupTarget
- PRDOTotal
- SysGroupTotal

### REPLACE_CANCEL_CREATE

Nếu old allocation đã cancel toàn phần:

`CreateQty = TargetQty`

Giữ logic đã pass ở `STD90000408AA`.

---

# 28. FINAL VERIFY

Không PASS chỉ vì SAP Save không lỗi.

Final verify phải kiểm:

## Member

- expected status;
- expected effective qty;
- expected bin;
- expected batch.

## Group

`sum(verified member effective qty) = GroupTargetTotal`

## Action ledger

mọi RequiresWrite action:

- VERIFIED
hoặc
- explicitly SKIPPED because already verified.

Không còn `SENT`.

Không còn `UNKNOWN`.

---

# 29. CHECK FINGERPRINT

CHECK tạo:

- InputFingerprint
- WHSnapshotFingerprint
- SourceStockFingerprint
- ReleasePlanFingerprint
- PlanFingerprint
- SapMapFingerprint
- ConfigFingerprint

XỬ LÝ phải compare lại.

Nếu user sửa Actual Qty sau CHECK:

XỬ LÝ hard stop.

Nếu SAP state materially change:

hard stop/recheck.

---

# 30. ACTION ID VÀ IDEMPOTENCY

ActionID không dùng row number.

Ví dụ:

`CANCEL|OB|DOC|ITEM|WT`

`RELEASE|TARGET_SOURCEKEY|HOLDER_WT`

`PRDO_QTY|OB|DOC|ITEM|TARGET_QTY`

`CREATE|OB|DOC|ITEM|MATERIAL|BIN|BATCH|QTY`

Một ActionID đã VERIFIED:

- không resend.

Một ActionID UNKNOWN:

- không resend.

---

# 31. RESUME ENGINE

History mới phải cho phép resume có chứng cứ.

## Khi mở workbook

Nếu latest run có:

`RecoveryRequired=TRUE`

Control Center hiển thị:

`RECOVERY REQUIRED`

Không cho user XỬ LÝ plan mới trước khi reconcile.

## Reconcile

Đọc SAP read-only:

- exact WT
- exact document/item
- exact created WT
- exact status

Nếu chứng minh action cũ succeeded:

`ACTION_VERIFIED_AFTER_RECOVERY`

Nếu chứng minh failed/no mutation:

chỉ cho phép retry khi action contract cho phép và chưa có SAP side effect.

---

# 32. BỎ SHEET POSTING

Đây là yêu cầu bắt buộc.

Không chỉ hide `POSTING`.

Phải audit toàn bộ VBA reference tới `"POSTING"`.

### Migration

Routine cũ:

`Worksheets("POSTING")`

phải đổi sang:

- in-memory SourceStockRepository
hoặc
- `DATA_CORE.tblSourceStock`.

### Trước khi delete POSTING

Codex phải:

1. search toàn source VBA.
2. lập dependency list.
3. thay mọi reader.
4. thay mọi writer.
5. compile.
6. replay regression.
7. xác nhận zero runtime reference.
8. mới delete sheet.

Không để catch-error tự recreate POSTING.

---

# 33. DEPRECATION MAP TỪ 20 SHEET

| Legacy | V2.7 destination |
|---|---|
| ORDER_QUERY | DATA_CORE.tblOBQueue / source request |
| POSTING_RAW | DATA_CORE.tblSourceStockRaw |
| INPUT | INPUT |
| WH_ACTUAL | DATA_CORE.tblWHSnapshot + tblActualSnapshot; UI projection ở INPUT |
| POSTING | DELETE sau migration |
| RUN_LOG | RUN_HISTORY.tblRunEvents |
| SAP_MAP | SYSTEM.tblSapMap |
| CONFIG | SYSTEM.tblConfig |
| OB_QUEUE | DATA_CORE.tblOBQueue |
| OB_WH_CACHE | DATA_CORE.tblWHSnapshot |
| POSTING_SUMMARY | DATA_CORE.tblCaptureSummary |
| POSTING_CACHE | DATA_CORE.tblSourceStock |
| PLAN_SOURCE_LEGS | PLAN.tblSourceLeg |
| PLAN_RELEASE | PLAN.tblReleaseCandidate |
| SAP_PROBE_LOG | SYSTEM.tblSapProbe |
| PLAN | PLAN.tblMemberPlan + tblGroupPlan + tblActionPlan |
| SAP_ACTION_MAP | SYSTEM.tblSapActionMap |
| RUN_HISTORY | RUN_HISTORY.tblRunEvents |
| RUN_SUMMARY | RUN_HISTORY summary/table |
| REGRESSION_AUDIT | RUN_HISTORY regression events |

---

# 34. MIGRATION STRATEGY — KHÔNG DELETE 15 SHEET NGAY LẬP TỨC

Trong development:

### Phase A

- clone V2.6.
- build adapters.
- tạo 5 target sheets.
- vẫn giữ legacy sheets.

### Phase B

- dual-read verify.
- output V2.7 repositories.
- compare legacy vs new.

### Phase C

- chuyển production engine sang new repositories.
- search zero legacy runtime reference.

### Phase D

- delete/decommission legacy sheets.
- final workbook chỉ còn 5 sheet.

Nếu Codex chưa chứng minh zero dependency:

**không được xóa sheet sớm.**

---

# 35. PERFORMANCE

Mục tiêu tối ưu không phải chỉ ít sheet.

## Không preallocate 500 blank rows

Không tạo table 500 dòng trống như POSTING hiện tại.

Chỉ ghi actual rows.

## In-memory

Trong một run:

- Load snapshot vào Dictionary/Collection once.
- Grouping xử lý trong memory.
- Chỉ persist snapshot/evidence cần thiết.

## Worksheet IO

- block read/write arrays.
- không cell-by-cell loop cho hàng nghìn log row.

---

# 36. HISTORY PERFORMANCE

Append event theo batch buffer.

Ví dụ:

- accumulate 20–100 events in memory;
- flush khi action boundary;
- flush ngay khi UNKNOWN/ERROR;
- flush cuối phase.

Không để mất log nếu Excel error.

---

# 37. UI PERFORMANCE

INPUT chỉ render active worklist.

Không copy toàn `DATA_CORE` ra UI.

Ví dụ backend có 500 WT nhưng 50 WT active:

UI chỉ render 50.

C status không render.

---

# 38. NUMBER PARSER — KHÔNG REGRESSION

V2.7 phải giữ canonical SAP quantity parser.

Regression bắt buộc:

- `1.200` -> `1200`
- `1.380` -> `1380`
- `1.640` -> `1640`
- `20,5` -> `20.5`
- `0` -> valid zero
- blank -> blank/invalid input, không tự thành zero

Không để `1200 -> 1.2`.

---

# 39. INPUT SEMANTICS

Blank khác 0.

## Actual Qty blank

= user chưa quyết định.

Không coi là cancel.

## Actual Qty = 0

= explicit cancel target.

Phải log InputSnapshot chứng minh user nhập 0.

---

# 40. MULTI-OB

UI nhập nhiều OB.

CHECK phải lập plan cho toàn batch nhưng execution isolation theo OB/document.

Một OB blocked không được làm hỏng history identity của OB khác.

Cho phép policy:

- `STOP_ALL_ON_BLOCKER`
hoặc
- `EXECUTE_READY_OB_ONLY`

Default safety:

`STOP_ALL_ON_BLOCKER=TRUE`

cho đến khi partial-batch execution được proven.

---

# 41. MESSAGE / ERROR DESIGN

Mỗi stop code unique.

Không chỉ:

`Scheduler halted`.

Phải có:

- StopCode
- OB
- Document
- Item
- WT
- ActionID
- Expected
- Actual
- RequiredNextStep

Ví dụ:

`STOP_RELEASE_PLAN_STALE_RECHECK_REQUIRED|WT=...|EXPECTED_STATUS=ACTIVE|ACTUAL=A`

---

# 42. USER-FACING STATUS

Không spam MsgBox.

INPUT có status bar:

- LOADING WH
- WH READY
- INPUT CHANGED
- CHECKING
- CHECK PASS
- CHECK BLOCKED
- RELEASE PLAN READY
- EXECUTING
- VERIFYING
- READY FOR USER CF
- RECOVERY REQUIRED

MsgBox chỉ dùng:

- safety confirmation pilot;
- hard recovery;
- fatal config/map issue.

---

# 43. PROGRESS

Hiển thị:

`OB 2/5 | Document 1/2 | Phase CREATE | Action 14/21`

Không hiển thị spinner vô nghĩa.

---

# 44. UI COLOR SEMANTICS

Không phụ thuộc màu làm logic.

Màu chỉ visual:

- neutral = chưa check
- warning = cần review/release
- blocked = blocker
- ready = plan pass
- executing = đang chạy
- complete = user handoff/verified

Logic phải dựa State code.

---

# 45. UI INPUT PROTECTION

Protect sheet nhưng unlock:

- OB input
- Actual Qty
- Area
- Actual Bin
- Actual Batch

Không để user sửa:

- Sys Qty
- Sys Bin
- Sys Batch
- WT
- Document
- SourceKey
- Plan state.

---

# 46. SHEET NAME COMPATIBILITY

Nếu nhiều macro public bên ngoài còn gọi legacy sheet:

tạo adapter VBA, không giữ 15 legacy sheet chỉ vì compatibility.

Ví dụ:

`GetSourceStockRows()` thay trực tiếp `Worksheets("POSTING")`.

---

# 47. CHECK CONTRACT TEST

Sau CHECK:

- không SAP business write.
- no status mutation.
- no WT cancellation.
- no PRDO save.
- no created WT.
- History chỉ có READ/PLAN events.

Bắt buộc static audit các method write.

---

# 48. EXECUTE CONTRACT TEST

XỬ LÝ không được tự rebuild plan.

History đầu EXEC phải ghi:

- CheckID
- PlanID
- all fingerprints
- plan action count.

Nếu mismatch:

0 SAP write.

---

# 49. RELEASE HOLDER TEST MATRIX

Tối thiểu:

1. free stock đủ -> no release.
2. thiếu 100, một holder 100 -> candidate selected.
3. thiếu 100, holder 40+60 -> two selected.
4. holder A -> excluded.
5. holder C -> excluded.
6. holder wrong batch -> excluded.
7. holder wrong bin -> excluded.
8. holder wrong material -> excluded.
9. holder same target WT -> excluded.
10. candidate changes status after CHECK -> EXEC stop stale.
11. cancel sent/readback A -> credit added.
12. cancel sent/readback unknown -> UNKNOWN/no credit.
13. insufficient coverage -> CHECK blocked.
14. PLAN_ONLY -> no cancel.
15. PILOT -> one group only.
16. LIVE_LOCKED -> only after proof gate.

---

# 50. GROUP TEST MATRIX

### Mandatory bug case

3500 + 4500 -> 0 + 8000.

Expected:

- net group delta 0;
- cancel 3500;
- retain 4500;
- create 3500;
- group final 8000.

### Other cases

- 500 + 500 -> 0 + 1000
- 300 + 700 -> 500 + 500
- 1000 -> 1200
- 1000 -> 800
- 1000 old bin -> 1000 new bin
- 1000 old batch -> 1000 new batch
- 3-member group reallocation
- 2 materials in one document
- multiple documents in one OB
- multiple OB.

---

# 51. SESSION TEST MATRIX

1. MON already open.
2. MON wrong node.
3. MON lost after Cancel.
4. POST screen open.
5. document already open.
6. multiple SAP sessions.
7. busy session.
8. popup after write.
9. session disappears.
10. Broker recovery once.

Assertions:

- no duplicate MON tabs in normal flow.
- no reopen per P000x.
- all readers use broker.

---

# 52. HISTORY TEST MATRIX

For one full test run, history must prove:

- LOAD_START/COMPLETE
- CHECK_START
- InputSnapshot
- PlanMember
- PlanGroup
- ReleasePlan if relevant
- CheckPass
- ExecStart
- ActionPrecheck
- ActionSent
- ActionReadback
- ActionVerified
- FinalVerify
- UserCFHandoff

If unknown:

- ActionUnknown
- RecoveryRequired

No ambiguity về whether action was sent.

---

# 53. HISTORY RESUME TEST

Simulate interruption:

## A. after Cancel sent

Restart/reopen.
No resend until readback.

## B. after POST Save

No second Save before reconcile.

## C. after Create Save

Find created WT before any retry.

## D. after final verify

Do not rerun business write.

---

# 54. 5-SHEET ACCEPTANCE

Final package fails nếu còn user/runtime sheet thừa.

Expected workbook:

1. INPUT
2. DATA_CORE
3. PLAN
4. SYSTEM
5. RUN_HISTORY

No:

- POSTING
- RUN_LOG
- RUN_SUMMARY
- REGRESSION_AUDIT
- separate SAP_MAP
- separate SAP_ACTION_MAP
- separate OB_QUEUE
- separate caches

unless Codex explicitly documents a technical blocker preventing migration; trong trường hợp đó **không được gọi build là final V2.7**.

---

# 55. UI ACCEPTANCE

1. Workbook open -> INPUT.
2. chỉ INPUT visible user mode.
3. 4 buttons present.
4. CHECK/XỬ LÝ separate.
5. XỬ LÝ disabled before CHECK.
6. editing Actual after CHECK invalidates plan.
7. C rows hidden from worklist.
8. OB queue supports multiple OB.
9. no default OB auto-filled.
10. no technical sheet visible.
11. status/progress clear.
12. no MsgBox spam.
13. release warning visible.
14. recovery warning persistent.
15. window/freeze works with 100+ active rows.

---

# 56. CODING RULES

- `Option Explicit`.
- Long for row index/counters where relevant.
- Double/Decimal-safe quantity parsing; không Integer cho qty lớn.
- no blanket `On Error Resume Next`.
- no silent catch after SAP write.
- no SendKeys in production RUN.
- no unbounded loop.
- no sleep as primary proof.
- no row-number business identity.
- no duplicate write path for same action type.

---

# 57. MODULE ARCHITECTURE GỢI Ý

Không bắt buộc đúng tên nhưng trách nhiệm phải tách.

- `modCFEntryPoints`
- `modCFUI`
- `modCFDataRepository`
- `modCFPlanRepository`
- `modCFHistory`
- `modCFPlanner`
- `modCFGroupAllocator`
- `modCFReleasePlanner`
- `modCFExecutor`
- `modCFRecovery`
- `modSapSessionBroker`
- `modSapMonReader`
- `modSapSourceStockReader`
- `modSapDocumentExecutor`
- `modSapWTExecutor`
- `modCFNumberParser`
- `modCFRegression`

Không để một module hàng nghìn dòng vừa UI vừa SAP vừa planner.

---

# 58. ENTRY POINTS CUỐI

Public entry points:

```text
CF_LoadWH
CF_Check
CF_Execute
CF_ResetUI
```

Không cho button gọi routine thấp hơn.

---

# 59. `CF_Check`

Pseudo-flow:

```text
Acquire UI lock
Assert no active EXEC
Validate OB input
Validate Actual input
Capture InputSnapshot
Acquire/reuse MON
Validate WH snapshot freshness
Read required source stock
Read holder candidates read-only
Build MemberPlan
Build GroupPlan
Build SourceLegs
Build ReleasePlan
Build ActionPlan
Run all gates
Compute fingerprints
Persist PLAN
Render preview
Mark CHECK PASS/BLOCKED
Release UI lock
```

No business write.

---

# 60. `CF_Execute`

Pseudo-flow:

```text
Acquire exclusive EXEC lock
Load locked PlanID
Verify PlanFingerprint
Verify InputFingerprint
Verify Map/Config fingerprint
Verify recovery clear
Fresh precheck of entities that will be written
For each scheduler phase:
    execute exact planned ActionID
    persist SENT boundary before/at write
    live readback
    mark VERIFIED or UNKNOWN
    on UNKNOWN stop
Final MON verify
Persist final result
Render UI
Release lock
```

---

# 61. `CF_ResetUI`

Pseudo-flow:

```text
If active EXEC -> reject
If unresolved UNKNOWN -> preserve recovery and only clear safe visual input
Else:
    clear OB queue input
    clear actual UI projection
    clear current Check/Plan references
    retain history
    retain SYSTEM
```

---

# 62. CANCEL HOLDER ROLLOUT PLAN

## V2.7.0

Default `PLAN_ONLY`.

Mục tiêu:

- candidate reader;
- selector;
- coverage;
- fingerprints;
- UI warning;
- history proof.

## V2.7.1 pilot

`PILOT`.

Test một release group.

## Sau live proof

mới cho `LIVE_LOCKED`.

Không bypass gate chỉ để demo end-to-end.

---

# 63. KHÔNG AUTO CF

Dù action khác pass:

final state:

`READY_FOR_USER_CF`

Không nhấn Confirm.

Không send CF popup.

Không mark completed C nếu MON chưa chứng minh user đã CF sau đó.

---

# 64. C STATUS REFRESH

Khi TẢI WH lần sau:

WT đã C:

- backend capture;
- update history if relevant;
- không render active worklist.

Nếu cả OB complete:

UI summary có thể:

`OB COMPLETE`

không đưa vào work queue.

---

# 65. DEBUG MODE

Cho một CONFIG:

`DEBUG_MODE=FALSE`

FALSE:

- chỉ INPUT visible.

TRUE:

- cho phép unhide 4 backend sheet;
- không đổi safety flags;
- không tự arm holder cancel.

---

# 66. VERSION / CHANGELOG

Final file:

`260820_CF_OB_V2.7.xlsm`

Version text:

`V2.7 — 4BUTTON / 5SHEET / ALLOCATION LEDGER / SINGLE MON / RELEASE GATED / HISTORY V2`

History đầu tiên sau migration ghi event:

`MIGRATION_V26_TO_V27`

kèm:

- old sheet count = 20
- new sheet count = 5
- posting removed = true
- history unified = true.

---

# 67. BEFORE/AFTER AUDIT BẮT BUỘC

Codex phải xuất report:

## Before

- sheet count;
- visible sheets;
- VBA modules;
- key public entry points;
- legacy sheet references.

## After

- 5 sheets;
- only INPUT visible;
- legacy sheet references = zero;
- compile result;
- regression result;
- file CRC/ZIP integrity;
- button mappings;
- no broken defined names;
- no broken shapes.

---

# 68. KHÔNG ĐƯỢC GỌI FINAL NẾU CHỈ ĐỔI UI

V2.7 không đạt nếu:

- chỉ đổi 3 nút thành 4;
- chỉ hide POSTING;
- chỉ merge history visually;
- vẫn dùng old group boolean;
- CHECK vẫn write SAP;
- XỬ LÝ tự rebuild plan;
- release holder không có readback;
- MON vẫn reopen theo member.

---

# 69. KHÔNG ĐƯỢC GỌI FINAL NẾU CHỈ REFACTOR SHEET

Sheet giảm còn 5 nhưng luồng nghiệp vụ regression thì cũng fail.

Cần cả:

- data parity;
- plan parity;
- live-proven branches preserved;
- bug mixed allocation fixed.

---

# 70. LIVE-PROVEN BRANCHES CẦN GIỮ

Không phá các hành vi đã có evidence:

- CF_ONLY -> user handoff.
- Cancel own WT readback A.
- replace/cancel/create full target branch.
- POST change exact source.
- Create WT then post verify.
- no automatic CF.
- unknown/no retry safety.
- number parser.

---

# 71. ACCEPTANCE GATE TỔNG

Final V2.7 chỉ PASS khi:

### Architecture

- [ ] exactly 5 final sheets.
- [ ] no POSTING runtime sheet.
- [ ] one visible INPUT.
- [ ] repositories normalized.

### UI

- [ ] 4 separate buttons.
- [ ] CHECK read-only.
- [ ] EXEC only after locked plan.
- [ ] worklist multi-OB.
- [ ] C projection hidden.
- [ ] release warning.

### Planner

- [ ] member ledger.
- [ ] group ledger.
- [ ] action plan.
- [ ] mixed net-zero allocation case pass.

### Holder cancel

- [ ] exact reader.
- [ ] deterministic selector.
- [ ] plan fingerprint.
- [ ] readback proof.
- [ ] PLAN_ONLY default.

### Session

- [ ] single MON broker.
- [ ] no per-member MON reopen.
- [ ] bounded recovery.

### History

- [ ] unified event ledger.
- [ ] SENT/VERIFIED/UNKNOWN explicit.
- [ ] before/target/after.
- [ ] action timing.
- [ ] session/control evidence.
- [ ] resume state.

### Safety

- [ ] no auto CF.
- [ ] no unknown retry.
- [ ] no stale plan execute.
- [ ] no row-index identity.

---

# 72. TEST CASE ƯU TIÊN ĐẦU TIÊN SAU KHI BUILD

Không test ngẫu nhiên trước.

## Offline replay 1

Case:

`3500 + 4500 -> 0 + 8000`

Phải chứng minh CreateQty=3500.

## Offline replay 2

`STD90000408AA`

Giữ replace branch pass.

## Offline replay 3

holder required nhưng mode PLAN_ONLY.

Expected:

- CHECK lập release plan.
- zero cancel.
- Execute không vượt safety gate.

## Live test 1

Case không cần holder cancel.

## Live test 2

mixed allocation.

## Live holder pilot

Chỉ sau khi hai test trên ổn.

---

# 73. OUTPUT CODEX PHẢI TRẢ VỀ

Không chỉ gửi file.

Phải trả:

1. file `260820_CF_OB_V2.7.xlsm`;
2. ZIP backup/package;
3. `V2.7_CHANGELOG.md`;
4. `V2.7_MIGRATION_AUDIT.md`;
5. `V2.7_REGRESSION_REPORT.md`.

Report phải ghi rõ:

- đã xóa/gộp sheet nào;
- POSTING reference đã loại chưa;
- 4 button macro mapping;
- net-zero allocation test;
- holder mode;
- MON navigation counts;
- history schema result;
- unresolved blockers.

---

# 74. ƯU TIÊN TRIỂN KHAI

Không làm tất cả một lần lộn xộn.

## P0 — Safety split

- tách CHECK / EXEC.
- button state.
- fingerprints.
- CHECK read-only assertion.

## P1 — Group allocation

- member/group ledger.
- fix net-zero group reallocation.

## P2 — Data repository

- remove POSTING dependency.
- consolidate DATA_CORE.

## P3 — Single MON

- session broker.

## P4 — Holder release planner

- reader/selector/readback/gate.

## P5 — History V2

- event ledger.
- resume evidence.

## P6 — Sheet consolidation

- 20 -> 5 after dependency parity.

## P7 — UI polish

- final Control Center.
- only INPUT visible.

---

# 75. FINAL DESIGN INTENT

V2.7 không phải "V2.6 thêm vài If".

Nó phải biến file từ:

`nhiều sheet + nhiều state phân tán + plan và execute dính nhau`

thành:

```text
INPUT
  |
  +-- TẢI WH
  |
  +-- CHECK ------------------------------+
  |    READ ONLY                          |
  |    snapshot                           |
  |    source stock                       |
  |    group/member allocation            |
  |    release candidates                 |
  |    locked action plan                 |
  |                                       v
  +-- XỬ LÝ -------------------------- LOCKED PLAN
       own cancel
       holder release (gated)
       POST change
       PRDO
       create
       verify
       |
       v
  READY FOR USER CF
```

Backend chỉ còn:

```text
INPUT
DATA_CORE
PLAN
SYSTEM
RUN_HISTORY
```

`POSTING` không còn.

`CHECK` không còn write.

`XỬ LÝ` không còn tự đoán plan.

`GroupNetDelta=0` không còn làm mất member create delta.

Holder cancel không còn được coi thành công chỉ vì nút Cancel đã được nhấn.

History phải đủ chi tiết để sau bất kỳ lần dừng nào cũng trả lời được:

> "Tool đã làm tới đâu, SAP đã nhận action nào, action nào đã verify, action nào chưa gửi, và lần chạy tiếp theo phải resume từ đâu?"

Đó là tiêu chuẩn để gọi bản này là **V2.7**.
