# PROMPT CODEX — NÂNG CẤP `260819_CF_OB_V2.6_LICH_SU` → `V2.7`

> **Mục tiêu:** nâng cấp trực tiếp workbook `260819_CF_OB_V2.6_LICH_SU.xlsm` thành một bản V2.7 ổn định hơn, xử lý đúng các nhóm nhiều WT/mã có tổng bằng nhau nhưng phân bổ thay đổi, siết chặt session/tab MON, hoàn thiện cancel WH/WT khác đang giữ tồn, gom luồng theo group/document đúng nghĩa, và thay UI bằng một Control Center duy nhất có 3 nút.
>
> Đây **không phải yêu cầu viết lại tool từ đầu**. Hãy giữ toàn bộ phần đã LIVE proven của V2.6, sửa kiến trúc group/scheduler/session/release-holder/UI tại các điểm gây lỗi, rồi chạy replay/regression từ lịch sử có sẵn trước khi xuất file V2.7.

---

## 0. FILE NGUỒN VÀ NGUYÊN TẮC LÀM VIỆC

### File chính

- `260819_CF_OB_V2.6_LICH_SU.xlsm`
- Workbook hiện có 20 sheet:
  - `ORDER_QUERY`
  - `POSTING_RAW`
  - `INPUT`
  - `WH_ACTUAL`
  - `POSTING`
  - `RUN_LOG`
  - `SAP_MAP`
  - `CONFIG`
  - `OB_QUEUE`
  - `OB_WH_CACHE`
  - `POSTING_SUMMARY`
  - `POSTING_CACHE`
  - `PLAN_SOURCE_LEGS`
  - `PLAN_RELEASE`
  - `SAP_PROBE_LOG`
  - `PLAN`
  - `SAP_ACTION_MAP`
  - `RUN_HISTORY`
  - `RUN_SUMMARY`
  - `REGRESSION_AUDIT`
- Defined name hiện có: `CFX_UseNewEngine=FALSE`.
- Không tự ý bật engine mới nếu chưa chứng minh parity với executor hiện tại.
- Không rebuild workbook trắng.
- Không tạo workbook mới từ đầu rồi copy dữ liệu sang.
- Làm việc trên một bản copy của V2.6, giữ nguyên VBA, XML relationships, styles, buttons, controls và lịch sử trừ các phần được yêu cầu sửa.

### Dữ kiện kỹ thuật trong folder cần kế thừa

1. **Shared-lock / reuse SAP context**
   - Ưu tiên tái sử dụng đúng SAP tab/session hiện có.
   - Không reopen/reload transaction nếu context hiện tại đã đúng và read-back đã chứng minh.
   - Không tạo tab mới chỉ để “chắc chắn”.

2. **MIGO fast/event principle**
   - Không dùng sleep cố định làm tín hiệu chính.
   - Ưu tiên `session.Busy`, status bar, popup close, read-back/state transition.
   - Timeout chỉ là fallback bounded.

3. **MAP SAP full technical ID principle**
   - RUN chỉ dùng Technical ID / technical column / verified command đã khóa.
   - Không đoán caption, index, tọa độ hoặc SendKeys trong RUN.
   - MAP mode và RUN mode phải tách rõ.

### Hard safety giữ nguyên

- Không auto CF ở V2.7.
- `CF_USER_HANDOFF_REQUIRED=TRUE` là chốt tuyệt đối.
- `CF_LIVE_ENABLED=FALSE` và `ALLOW_CF=FALSE` phải tiếp tục được tôn trọng.
- Không retry một SAP write khi outcome chưa rõ.
- Mọi write phải đi qua state `PLANNED -> SENT -> VERIFIED` hoặc `UNKNOWN`.
- Nếu `UNKNOWN` sau write: không resend; dừng hard-stop để user kiểm tra.
- Không dùng row index làm identity nghiệp vụ.
- Identity phải dùng tuple ổn định: OB + Document + Item/WH + Material + WT + SourceBin + SourceBatch + TargetBin + TargetBatch tùy action.
- Không kill Excel/SAP của user.
- Không xóa lịch sử cũ.
- Không sửa MIGO; MIGO vẫn locked/off trong project này.

---

# 1. EVIDENCE TỪ V2.6 — BẮT BUỘC HIỂU TRƯỚC KHI SỬA

## 1.1. Lịch sử chạy hiện tại

`RUN_SUMMARY` đang có 100 record tổng hợp:

- 55 lượt `CHECK`.
- 45 lượt `EXECUTE`.
- 9 lượt `PASS|LOCKED_PLAN_PHASES_COMPLETE`.
- 23 lượt có `SCHEDULER_HALTED_OR_BLOCKED`.
- 8 lượt dừng trực tiếp do `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE`.
- 7 lượt có `STATE_UNKNOWN`.
- 6 lượt bị `OB_REGRESSION_GATE_BLOCKED`.

Điều này chứng minh V2.6 đã có nhiều safety gate tốt nhưng group model vẫn gây false-stop/false-no-op trong các trường hợp nhiều member.

## 1.2. Case mới nhất — lỗi quan trọng nhất phải fix đầu tiên

Run:

- WH Run ID: `20260820_052113`
- CHECK: `CHECK|20260820_052144|1930477|20260820_052113`
- EXEC: `EXEC|20260820_052148|1930809|CHECK-20260820_052144-1930477-20260820_052113`
- OB: `1011846094`
- Document: `1012100366`

### Group `STD20000556`

Có 2 WT cùng QTY_GROUP:

#### P0002

- WT `5042936844`
- Sys Qty `3500`
- Target Qty `0`
- Bin `APFL-HR07-10-01`
- Batch `260518-17`
- ExecMode `CANCEL_PRDO_ZERO`
- Quyết định: cancel WT này về 0.

#### P0003

- WT `5042936845`
- Sys Qty `4500`
- Target Qty `8000`
- Bin `APFL-HR01-14-01`
- Batch `260602-13`
- ExecMode `KEEP_EXISTING_CREATE_DELTA`
- Nguồn Posting exact có `61000` available / `65500` total.
- Cần tăng thêm thực tế: `8000 - 4500 = 3500`.

### Tổng group

- `GROUP_CURRENT_TOTAL = 3500 + 4500 = 8000`
- `GROUP_TARGET_TOTAL = 0 + 8000 = 8000`
- `GROUP_DELTA_TOTAL = 0`
- `GROUP_MEMBER_COUNT = 2`
- `GROUP_NEED_QTY = FALSE` trong PLAN hiện tại.

### Lỗi V2.6

RUN_LOG ghi:

- `SCHED_QTY_GROUP_BASELINE ... GROUP_CURRENT=8000|GROUP_TARGET=8000`
- sau đó:
  - `PRDO_QTY_ALREADY_TARGET|LIVE_TOTAL=8000|TARGET=8000|SOURCE_VERIFY=SKIP|NO_QTY_WRITE`

Kết luận này **chỉ đúng cho tổng PRDO**, nhưng **sai nếu dùng để quyết định không còn action ở member**.

Sau cancel WT 3500, member P0003 vẫn đang giữ 4500 và phải tạo thêm 3500 để đạt 8000.

Đây là lỗi kiến trúc:

> **NET GROUP DELTA = 0 không đồng nghĩa ACTION DELTA = 0.**

Cần tách hai khái niệm:

1. **PRDO aggregate quantity contract** — tổng quantity của document/item/material.
2. **WT allocation contract** — mỗi member sau cancel/retain/replace phải còn/create bao nhiêu.

### False-stop tiếp theo

Sau khi P0003 bị coi là `QTY_RECONCILED_PASS`, scheduler xuất:

- `SCHED_PRDO_DOCUMENT_MEMBER_ISOLATION`
- `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED`

cho P0002, P0003 và P0005.

Tức document executor hiện chưa có một ledger đầy đủ để phân biệt:

- member không cần PRDO write,
- member đã cancel,
- member vẫn cần create delta,
- member chỉ cần CF,
- member bị blocked thật.

## 1.3. Case `STD90000408AA` — minh họa đúng cách create full target

Cùng run mới nhất:

### P0005

- WT `5042936841`
- Sys Qty `5050`
- Target `0`
- cancel.

### P0006

- WT `5042936842`
- Sys Qty `12950`
- Target `18000`
- đổi bin/batch:
  - từ `APFL-HR05-05-01`, batch `260626-09`
  - sang `APFL-HR01-13-01`, target batch `260802-01`
- ExecMode `REPLACE_CANCEL_CREATE`.

Group:

- current `18000`
- target `18000`
- net delta `0`

Nhưng V2.6 vẫn đã thực hiện đúng nhánh replace:

- cancel cả 2 old WT trong 1 cancel group,
- POST change,
- PRDO Save,
- Create WT `18000`,
- post verify ra WT mới `5043000572=18000`,
- `READY_FOR_CF`.

Điều này chứng minh V2.7 **không được phá** logic `REPLACE_CANCEL_CREATE` đang chạy được.

Phần cần sửa chủ yếu là mixed allocation/retained delta và document-member accounting.

## 1.4. Lịch sử lặp lại của bug group member incomplete

Có nhiều run bị:

`STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE`

Ví dụ:

- `1011839679`:
  - ACTIVE_MEMBERS=1
  - EXPECTED_MEMBERS=2
  - PLAN_BEFORE=20
  - GROUP_BEFORE=120
  - PLAN_TARGET=0
  - GROUP_TARGET=120

- `1011848120`:
  - ACTIVE_MEMBERS=1
  - EXPECTED_MEMBERS=3
  - GROUP_CURRENT=588
  - GROUP_TARGET=208

- nhiều OB khác có 2–3 member cùng QTY_GROUP cũng dừng tương tự.

Không vá riêng case `3500 + 4500`. Hãy sửa mô hình group tổng quát.

## 1.5. MON/session chưa đủ chặt

Run mới nhất kết thúc với:

`MON_OPEN_COUNT=0|MON_REFRESH_COUNT=9|DOCUMENT_OPEN_COUNT=5|DOCUMENT_RECOVERY_COUNT=1|WHTASK_OPEN_COUNT=1|WHTASK_REUSE_COUNT=0|LIMIT=20`

Với chỉ 1 OB, 6 WH rows và 1 Document:

- `MON_REFRESH_COUNT=9` là quá nhiều.
- `DOCUMENT_OPEN_COUNT=5` cho một document là chưa đúng mục tiêu group/reuse.
- `MON_SESSION_RECOVERY` xuất hiện trong giữa luồng create.

V2.7 phải chuyển từ **per-plan navigation** sang **phase/group navigation**.

## 1.6. Cancel WH/WT khác giữ tồn chưa hoàn thiện

CONFIG hiện có:

- `CANCEL_OTHER_WT = TRUE`
- `RELEASE_EXCLUDE_STATUS = A,C`
- `ALLOW_PARTIAL_WT_CANCEL = FALSE`
- `AFTER_CANCEL_ADD_QTY_TO_CACHE = TRUE`
- `REFRESH_POSTING_AFTER_CANCEL = FALSE`

Và đã có:

- module `modReleaseHolderReader`
- sheet `PLAN_RELEASE`
- logic strings cho:
  - `RELEASE_HOLDER_READER_PASS`
  - `RELEASE_SELECTOR_PASS`
  - `FINAL_RELEASE_PLAN_PASS`

Nhưng các key sau đang BLANK:

- `RELEASE_MON_MATERIAL_ID`
- `RELEASE_MON_BIN_ID`
- `RELEASE_MON_BATCH_ID`
- `RELEASE_MON_EXECUTE_ID`
- `RELEASE_MON_RESULT_GRID_ID`
- `RELEASE_MON_OB_COL`
- `RELEASE_MON_WH_COL`
- `RELEASE_MON_WT_COL`
- `RELEASE_MON_MATERIAL_COL`
- `RELEASE_MON_BATCH_COL`
- `RELEASE_MON_BIN_COL`
- `RELEASE_MON_QTY_COL`
- `RELEASE_MON_STATUS_COL`
- `RELEASE_MON_RESULT_COUNT_ID`
- `RELEASE_MON_RESULT_COUNT_PATTERN`

`PLAN_RELEASE` trong run mới nhất không có holder candidate.

Vậy `CANCEL_OTHER_WT=TRUE` hiện mới là **capability/config scaffold**, chưa phải một luồng production hoàn chỉnh.

---

# 2. MỤC TIÊU KIẾN TRÚC V2.7

V2.7 phải dùng mô hình 4 tầng:

1. **OB batch / queue**
2. **Document group**
3. **Qty group / business item group**
4. **Member allocation ledger**

Không được quyết định action chỉ từ `GROUP_CURRENT_TOTAL` và `GROUP_TARGET_TOTAL`.

Mỗi member phải có vai trò rõ ràng:

- `CF_ONLY`
- `RETAIN`
- `RETAIN_AND_CREATE_DELTA`
- `CANCEL_TO_ZERO`
- `REPLACE_CANCEL_CREATE`
- `CHANGE_BIN`
- `CHANGE_BATCH`
- `CHANGE_BIN_BATCH`
- `RECOVERY_READBACK_ONLY`
- `BLOCKED`
- `TERMINAL_C`

Mỗi group phải có hai contract độc lập:

### Contract A — PRDO total

- `PRDO_LIVE_TOTAL`
- `PRDO_TARGET_TOTAL`
- `PRDO_WRITE_REQUIRED`

### Contract B — member allocation

- member nào retain bao nhiêu,
- member nào cancel bao nhiêu,
- member nào cần create bao nhiêu,
- nguồn nào dùng để create,
- holder release cần bao nhiêu,
- cuối cùng member/tuple phải có active qty bao nhiêu.

Chỉ khi **cả hai contract** pass mới được coi group hoàn tất.

---

# 3. WORKSTREAM A — THIẾT KẾ LẠI QTY GROUP / MEMBER ALLOCATION

## 3.1. Không dùng `GROUP_DELTA_TOTAL=0` để kết luận no-op

Thay logic kiểu:

```text
IF GROUP_CURRENT_TOTAL = GROUP_TARGET_TOTAL THEN NO_QTY_ACTION
```

bằng:

```text
PRDO_WRITE_REQUIRED = (PRDO_LIVE_TOTAL <> PRDO_TARGET_TOTAL)

GROUP_ACTION_REQUIRED =
    ANY_MEMBER_CANCEL_REQUIRED
    OR ANY_MEMBER_LOCATION_CHANGE_REQUIRED
    OR ANY_MEMBER_CREATE_REQUIRED
    OR ANY_MEMBER_RECOVERY_REQUIRED
    OR PRDO_WRITE_REQUIRED
```

`GROUP_NEED_QTY` cũ không còn đủ ý nghĩa. Hoặc đổi tên, hoặc giữ để compatibility nhưng thêm các field mới.

## 3.2. Member effective quantity

Sau phase cancel, mỗi member phải tính:

```text
RetainedActiveQty = quantity của exact active WT được giữ lại sau cancel phase
CancelledQty      = quantity của old WT đã VERIFIED status A
RequiredCreateQty = max(TargetQty - RetainedActiveQty, 0)
```

### Theo ExecMode

#### `CF_ONLY`

```text
RetainedActiveQty = exact existing active WT qty
RequiredCreateQty = 0
```

#### `KEEP_EXISTING_CREATE_DELTA`

```text
RetainedActiveQty = exact old WT qty nếu WT vẫn active và tuple đúng
RequiredCreateQty = max(TargetQty - RetainedActiveQty, 0)
```

**Không dùng** `TargetQty - GroupLiveTotal`.

**Không dùng** `GroupTarget - GroupCurrent`.

Ví dụ P0003:

```text
TargetQty = 8000
RetainedActiveQty = 4500
RequiredCreateQty = 3500
```

mặc dù group total trước/target đều là 8000.

#### `REPLACE_CANCEL_CREATE`

Sau khi exact old WT đã cancel VERIFIED:

```text
RetainedActiveQty = 0
RequiredCreateQty = TargetQty
```

Ví dụ P0006:

```text
TargetQty = 18000
RequiredCreateQty = 18000
```

#### `CANCEL_PRDO_ZERO`

```text
RetainedActiveQty = 0
RequiredCreateQty = 0
```

#### Change bin/batch nhưng old WT không bị retain

Phải coi như replace nếu nguồn/tuple cũ không còn hợp lệ:

```text
cancel old -> verify -> posting change if needed -> create full target
```

## 3.3. Tạo member allocation ledger

Ưu tiên thêm các cột vào `PLAN` để giảm sheet mới. Nếu code quá khó đọc thì có thể thêm **một** sheet VeryHidden `GROUP_LEDGER`, nhưng không tạo thêm sheet visible.

Các field bắt buộc phải tồn tại ở PLAN hoặc ledger:

- `AllocationRole`
- `MemberOriginalQty`
- `MemberTargetQty`
- `RetainedActiveQty`
- `CancelRequiredQty`
- `CancelVerifiedQty`
- `RequiredCreateQty`
- `CreatedVerifiedQty`
- `EffectiveQtyAfterCancel`
- `FinalExpectedQty`
- `PRDOLiveTotal`
- `PRDOTargetTotal`
- `PRDOWriteRequired`
- `GroupActionRequired`
- `SourceAvailableBeforeRelease`
- `ReleaseNeedQty`
- `ReleaseSelectedQty`
- `SourceAvailableAfterRelease`
- `AllocationGate`
- `AllocationState`
- `CancelIdempotencyKey`
- `ReleaseCancelIdempotencyKey`
- `PostingIdempotencyKey`
- `PRDOIdempotencyKey`
- `CreateIdempotencyKey`
- `FinalVerifyKey`

Nếu thêm cột PLAN, append phía phải để không phá các index cũ; code mới phải resolve theo header name, không hard-code column number mới.

## 3.4. Group final verify phải kiểm tra cả tổng và phân bổ

Không chỉ:

```text
sum(active qty group) = group target
```

Mà còn phải check từng member/business allocation:

```text
For each plan member:
    VerifiedActiveQtyForMember = TargetQty
```

Trong trường hợp một target được tạo thành nhiều WT mới hợp lệ, được phép sum theo exact member create tuple.

Ví dụ net-zero redistribution:

```text
Old:
WT-A 3500 -> target 0
WT-B 4500 -> target 8000

After:
WT-A status A
WT-B active 4500
WT-new active 3500, cùng create tuple của B

Member B verified total = 8000
Group verified total = 8000
PASS
```

## 3.5. Sửa `PRDO_QTY_ALREADY_TARGET`

`PRDO_QTY_ALREADY_TARGET` chỉ được phép có nghĩa:

> Không cần write tổng PRDO.

Nó **không được** set member thành complete nếu member vẫn còn `RequiredCreateQty > 0`.

Sau log này phải tiếp tục:

```text
PRDO_WRITE_REQUIRED=FALSE
MEMBER_ALLOCATION_CONTINUES=TRUE
```

nếu có create/cancel/location action còn lại.

## 3.6. Sửa `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED`

Hiện document executor đang coi các member không tham gia một PRDO save cụ thể là “not included”.

V2.7 phải chuyển thành accounting ledger:

Mọi member trong document phải thuộc đúng một state:

- `ACCOUNTED_CF_ONLY`
- `ACCOUNTED_CANCEL_VERIFIED`
- `ACCOUNTED_PRDO_READBACK_ONLY`
- `ACCOUNTED_PRDO_WRITE_REQUIRED`
- `ACCOUNTED_CREATE_PENDING`
- `ACCOUNTED_CREATE_VERIFIED`
- `ACCOUNTED_BLOCKED`
- `ACCOUNTED_TERMINAL_C`

Chỉ STOP `MEMBER_UNACCOUNTED` nếu member thật sự không có state.

Không STOP chỉ vì member không cần PRDO write.

## 3.7. Document Save

Giữ nguyên mục tiêu tốt của V2.6:

- một document mở một lần,
- read tất cả item cần quantity,
- write tất cả item cần write,
- Save tối đa 1 lần cho document trong một phase.

Nếu `PRDO_WRITE_REQUIRED=FALSE` cho toàn document:

- không Save,
- nhưng vẫn mark `PRDO_DOCUMENT_READBACK_COMPLETE=TRUE`,
- rồi chuyển create phase.

---

# 4. WORKSTREAM B — SCHEDULER GROUP PHASE MỚI

Thứ tự bắt buộc cho mỗi OB:

```text
PHASE 0  PRECHECK / LOCK PLAN
PHASE 1  OWN-WT CANCEL GROUPS
PHASE 2  ONE MON READBACK FOR ALL OWN CANCELS
PHASE 3  HOLDER RELEASE PLANNING / CANCEL IF NEEDED
PHASE 4  ONE MON READBACK FOR HOLDER RELEASE
PHASE 5  LOCATION / POSTING CHANGE GROUPS
PHASE 6  PRDO DOCUMENT TOTAL READ/WRITE
PHASE 7  CREATE/REACQUIRE WT ALLOCATION GROUPS
PHASE 8  FINAL MON VERIFY BY GROUP + MEMBER
PHASE 9  USER CF HANDOFF
```

## 4.1. Phase 0 — freeze plan

Before first write:

- sync UI input -> WH_ACTUAL,
- validate all editable rows,
- build PLAN,
- compute group/member ledger,
- hash/fingerprint,
- verify POSTING cache freshness,
- verify MON snapshot run ID,
- verify all actions have exact Technical ID/command map,
- verify no state unknown from prior run.

Nếu blocker chỉ thuộc một independent group:

- mark group blocked,
- vẫn có thể chạy independent groups.

Nếu blocker là global/session/state-unknown:

- hard-stop whole run.

## 4.2. Phase 1 — own cancel groups

Group by QTY_GROUP / compatible multi-selection cancel context.

Một cancel group:

- select all exact WTs planned cancel,
- one cancel action if SAP supports multi-select,
- one idempotency key for group + individual WT keys,
- no per-member reopen.

Giữ behavior tốt hiện tại như run P0005/P0006:

`GROUP_ACTION_COUNT=1|GROUP_MEMBERS=2|SELECTED_ROWS=...`

## 4.3. Phase 2 — one MON readback

Sau khi gửi tất cả own-cancel groups của OB:

- refresh MON **một lần**,
- reconcile tất cả WTs đã `CANCEL_SENT`,
- status A -> VERIFIED,
- active unexpected -> UNKNOWN/blocked theo group,
- không resend cancel.

Không refresh MON riêng cho từng plan.

## 4.4. Phase 3–4 — holder release

Chi tiết ở Workstream D.

## 4.5. Phase 5 — location groups

Group theo exact source/target tuple.

Không reopen POST cho mỗi member cùng tuple.

Nếu nhiều member cùng exact POST change tuple:

- calculate total posting quantity cần move,
- write once per posting group,
- post verify once,
- distribute verified quantity back into member ledger.

Không gộp các tuple khác bin/batch chỉ vì cùng material.

## 4.6. Phase 6 — PRDO document group

Mỗi unique Document:

- open/reuse document once,
- map all relevant PRDO rows,
- compute exact live total per QTY_GROUP,
- write only groups where `PRDO_WRITE_REQUIRED=TRUE`,
- Save once maximum,
- readback all group totals.

## 4.7. Phase 7 — create allocation groups

Create logic phải dùng `RequiredCreateQty` từ member ledger.

Không dùng group net delta làm create qty.

Create group key nên bao gồm:

```text
OB
Document
Item/WH
Material
SourceBin
SourceBatch
TargetBin
TargetBatch
StockType/ProcessType nếu cần
```

### Retained delta

P0003 pattern:

```text
Existing retained active = 4500
RequiredCreateQty = 3500
Create only 3500
Final member active sum = 8000
```

### Replace

P0006 pattern:

```text
Old active = 0 after cancel
RequiredCreateQty = 18000
Create full 18000
```

## 4.8. Phase 8 — final verify

Một MON refresh cuối cho OB.

Verify:

- all planned cancels A,
- all new/reacquired WTs active,
- exact source tuple,
- quantity allocation per member,
- group final total,
- no unexpected duplicate active WT,
- no active old WT that should have been cancelled.

Nếu pass:

- `READY_FOR_CF`
- CF not sent.

## 4.9. Scope của blocker

### Soft/group blocker

Ví dụ:

- source shortage không giải phóng được,
- missing holder candidate,
- one QTY_GROUP mismatch nhưng không có write unknown.

=> block group, continue independent group/OB.

### Hard/global stop

Chỉ khi:

- `STATE_UNKNOWN_AFTER_*`
- SAP session identity lost sau write,
- unknown modal after write,
- duplicate/ambiguous exact control leading to unsafe write,
- save/create/cancel send outcome unprovable,
- workbook plan fingerprint changed during execution.

---

# 5. WORKSTREAM C — SIẾT CHẶT MON / SAP SESSION / TAB

## 5.1. Mục tiêu

Trong một EXEC run:

- đúng **1 dedicated MON reader tab/session**.
- không tạo MON tab theo member/group.
- một executor session dùng cho POST/PRDO/WT.
- holder-reader cũng reuse dedicated MON session, không mở MON thứ hai.

## 5.2. Session broker duy nhất

Tạo/chuẩn hóa một `SessionBroker` hoặc logic tương đương trong VBA.

Có role:

```text
ROLE_MON_READER
ROLE_EXECUTOR
```

Mỗi role lưu:

- SAP connection identity
- session object identity
- System
- Client
- User
- Session index/id
- current TCode
- current title
- Busy state
- bound generation/run ID
- last verified at

## 5.3. Binding MON

At run start:

1. scan logged-in sessions.
2. tìm existing session đang ở `/SCWM/MON` đúng System/Client/User.
3. nếu đúng một session phù hợp -> bind.
4. nếu không có -> được phép create **một** session/tab MON duy nhất.
5. nếu đã create rồi trong run -> tuyệt đối không create lần 2.
6. nếu busy -> bounded wait/readback; không duplicate tab.
7. nếu session chết trước write -> rebind existing MON only.
8. nếu session chết sau write và readback không thể chứng minh -> hard-stop.

## 5.4. Banned paths

Trong RUN, ngoài SessionBroker:

- không module nào được tự gọi `CreateSession`.
- không function nào tự `/N/SCWM/MON` trên session khác để “recovery”.
- không dùng new tab như default recovery.

Search toàn source và centralize các call path.

Các function/module cần audit đặc biệt:

- `CWR_GetMONSession`
- `CWR_ResetSchedulerExecutionContext`
- `CWR_LoadExecutionContext`
- `modCF_MONCore_STABLE`
- `modReleaseHolderReader`
- `modCodexWTRework`
- các helper có `CreateSession`, `/SCWM/MON`, `MON_SESSION_RECOVERY`.

## 5.5. Navigation budget V2.7

Normal case per OB:

```text
MON_TAB_CREATE_COUNT <= 1 per EXEC run
MON_OPEN_COUNT <= 1 per EXEC run
MON_REFRESH_COUNT <= 4 per OB
DOCUMENT_OPEN_COUNT <= number_of_unique_documents
DOCUMENT_RECOVERY_COUNT = 0 normal; <=1 only bounded recovery
WHTASK_OPEN_COUNT <= number_of_unique_documents_that_need_create
```

Với case mới nhất 1 document:

Target:

```text
MON_REFRESH_COUNT <= 4   (thay vì 9)
DOCUMENT_OPEN_COUNT <= 1 normal, <=2 nếu recovery
WHTASK_OPEN_COUNT <= 1
```

## 5.6. MON refresh theo phase

Không refresh per member.

Allowed refresh points:

1. initial snapshot / TẢI WH
2. post own-cancel batch
3. post external-holder release batch nếu có
4. final verify

Nếu không có cancel/release thì có thể ít hơn.

## 5.7. Holder-reader reuse MON context

Khi dùng cùng MON session để query holder:

- save logical MON context in memory,
- run holder query,
- cache result,
- restore OB context only khi thực sự cần final verify,
- không reopen tab.

Nếu có thể đọc holder bằng cùng lower-grid/filter context mà không đổi tree root, ưu tiên cách đó.

---

# 6. WORKSTREAM D — HOÀN THIỆN `CANCEL OTHER WT/WH HOLDING STOCK`

## 6.1. Mục tiêu nghiệp vụ

Khi source exact cần cho create/posting không đủ free available vì đang bị WT khác giữ:

- tìm các WT khác đang giữ cùng exact stock tuple,
- loại trừ WT không được phép cancel,
- chọn deterministic whole-WT candidates,
- cancel đủ để giải phóng tồn,
- verify một lần,
- cộng qty verified vào Available cache,
- tiếp tục plan hiện tại mà không reload Posting chỉ vì cancel.

## 6.2. Source key

Holder release phải query ít nhất theo:

```text
Material + Bin + Batch
```

Nếu SAP surface có Stock Type/Owner/Entitled/Process Type và chúng ảnh hưởng identity thì đưa vào key.

Không query chỉ Material.

## 6.3. Khi nào tạo NeedRelease

Sau khi tính member allocation:

```text
RequiredCreateQty = qty thật cần create của member/group
FreeAvailable = exact Posting free stock
OwnCancelCredit = quantity từ own WT cancel nếu exact cancel giải phóng đúng source stock và business rule cho phép credit

ReleaseNeedQty = max(RequiredCreateQty - FreeAvailable - OwnCancelCredit, 0)
```

Không dùng target total làm NeedRelease nếu member chỉ cần delta.

## 6.4. Map RELEASE_MON

Hoàn thiện MAP mode để người dùng có thể map/lock các `RELEASE_MON_*` từ chính MON layout hiện tại.

Ưu tiên reuse các Technical IDs đã proven của MON nếu cùng screen/control:

- material input
- bin input
- batch input
- execute
- result grid
- result count/status
- OB/WH/WT/Material/Batch/Bin/Qty/Status columns

Nhưng **không copy mù** từ MON main mapping nếu screen context khác. Phải verify read-back.

Sau khi map:

- update `SAP_MAP`
- update `SAP_ACTION_MAP` nếu cần
- lock fingerprint
- ghi evidence.

## 6.5. Candidate eligibility

Candidate holder chỉ eligible nếu:

- exact Material matches
- exact Bin matches
- exact Batch matches
- Status not in `A,C`
- WT is active
- WT is not one of current plan's own WTs
- WT is not a newly created WT of current run
- Holder OB != current OB mặc định
- nếu Holder OB = current OB nhưng WT không nằm PLAN -> do not auto cancel; mark `REVIEW_SAME_OB_UNPLANNED_HOLDER`
- quantity > 0
- identity complete and unique

## 6.6. Whole-WT selection

`ALLOW_PARTIAL_WT_CANCEL=FALSE` phải giữ.

Selection deterministic:

1. Nếu có một WT có Qty >= NeedRelease:
   - chọn WT nhỏ nhất vẫn đủ.
   - tie -> WT number ascending.
2. Nếu không có single WT đủ:
   - chọn các WT theo Qty descending để giảm số lần cancel,
   - dừng khi cumulative >= NeedRelease,
   - tie -> WT ascending.
3. Ghi rõ over-release.

Không dùng subset-search phức tạp nếu làm tăng rủi ro/độ chậm.

## 6.7. PLAN_RELEASE

Mỗi candidate ghi:

- PlanID
- Leg
- SourceKey
- Holder OB
- Holder WH
- Holder WT
- Material
- Bin
- Batch
- Held Qty
- WT Status
- Eligible
- Selected
- Select Order
- Need Release
- Projected Available
- Gate
- Reason
- Query Run ID
- Query Source
- Posting Available
- thêm nếu cần:
  - `CancelState`
  - `CancelIdempotencyKey`
  - `CancelVerifiedAt`
  - `ReleasedQtyVerified`

## 6.8. Cancel release group

Nếu nhiều holder WTs selected và cùng MON grid:

- multi-select exact rows,
- one cancel action,
- one group readback.

Không cancel từng WT rồi refresh từng lần.

Idempotency:

```text
RELEASE_CANCEL|SourceKey|HolderWT
```

Nếu history = SENT chưa VERIFIED:

- readback only,
- không resend.

## 6.9. Available cache update

Chỉ sau status A readback verified:

```text
AvailableCache(SourceKey) += ReleasedQtyVerified
```

Không cộng ngay khi mới send cancel.

Không refresh `/SCWM/POST` chỉ vì cancel nếu `REFRESH_POSTING_AFTER_CANCEL=FALSE`.

## 6.10. Nếu vẫn thiếu

Nếu projected available < RequiredCreateQty sau tất cả eligible holder:

- block source group:
  - `STOP_SOURCE_SHORTAGE_AFTER_RELEASE`
- không write group đó.
- independent groups tiếp tục nếu safe.

---

# 7. WORKSTREAM E — RESUME / IDEMPOTENCY / HISTORY

V2.6 đã có nền tảng tốt, phải giữ và làm rõ hơn.

## 7.1. Action state machine

Mọi destructive/write action:

```text
PLANNED
SENT
VERIFIED
UNKNOWN
SKIPPED_ALREADY_VERIFIED
```

Actions:

- OWN_CANCEL
- RELEASE_CANCEL
- POST_CHANGE
- PRDO_SAVE
- WT_ADOPT
- WT_CREATE
- WT_SAVE

## 7.2. On rerun

Trước action:

- search RUN_HISTORY by idempotency key,
- nếu VERIFIED -> read SAP current state và skip resend,
- nếu SENT -> readback only,
- nếu UNKNOWN -> hard-stop until user resolves,
- nếu no history -> normal action.

## 7.3. Status A

Nếu own WT đã A:

- không cancel lại.
- derive remaining allocation from current target and active WTs.
- nếu target > 0 -> create/reacquire remaining.
- nếu target = 0 -> member complete.

Không bắt user chạy lại full flow từ đầu.

## 7.4. Status C

- `C` = done.
- Không đưa vào actionable planner.
- Không hiển thị trong main UI worklist mặc định.
- Vẫn giữ trong WH_ACTUAL, RUN_HISTORY, RUN_LOG để audit.
- Nếu cả OB đều C -> `OB_DONE`, skip nhanh.

## 7.5. Unknown state

Không biến unknown thành generic blocked.

Log exact:

```text
STATE_UNKNOWN_AFTER_CANCEL
STATE_UNKNOWN_AFTER_RELEASE_CANCEL
STATE_UNKNOWN_AFTER_POST_CHANGE
STATE_UNKNOWN_AFTER_PRDO_SAVE
STATE_UNKNOWN_AFTER_CREATE
STATE_UNKNOWN_AFTER_SAVE
```

Unknown => hard-stop whole run, vì session có thể đang ở trạng thái dirty.

---

# 8. WORKSTREAM F — GIẢM XUNG ĐỘT SAFETY FLAGS

CONFIG hiện có legacy flags:

- `ALLOW_QTY_WRITE=FALSE`
- `ALLOW_POST=FALSE`
- `ALLOW_CANCEL=FALSE`
- `ALLOW_CREATE=FALSE`

nhưng fast production route hiện vẫn có các action đã LIVE proven theo:

- `REWORK_ACTION_WHITELIST`
- `FAST_FLOW_ENABLED`
- `CF_USER_HANDOFF_REQUIRED`
- action map locked.

V2.7 cần audit để không còn **hai hệ safety cùng nói khác nhau**.

Không đơn giản bật hết `ALLOW_*` thành TRUE.

Hãy làm:

1. Xác định flag nào thực sự còn được executor đọc.
2. Flag legacy không dùng nữa:
   - giữ để compatibility nhưng đánh dấu `LEGACY_NOT_AUTHORITATIVE`, hoặc
   - consolidate call path về một central permission function.
3. Central gate trả kết quả per action:

```text
CanExecuteAction(ActionType, Plan, MapState, HistoryState)
```

4. Absolute deny:
   - CF auto
   - MIGO
   - delete WT
   - fallback coordinate/sendkeys
5. Rework allowed action chỉ nằm trong:

```text
WT_CANCEL
RELEASE_WT_CANCEL
POST_CHANGE
PRDO_QTY
WT_ADOPT
WT_CREATE
WT_SAVE
```

6. Không để global safety của một scenario chặn independent safe scenario nếu không liên quan.

---

# 9. WORKSTREAM G — UI V2.7 MỚI: MỘT CONTROL CENTER, 3 NÚT

## 9.1. Nguyên tắc

Người dùng chỉ cần thấy **một sheet chính**.

Không cần xóa sheet engine; chỉ hide/VeryHidden.

Ưu tiên giữ tên sheet `INPUT` để tránh phá string reference. Có thể đổi title hiển thị trong sheet thành:

`CF OB — CONTROL CENTER V2.7`

Nếu muốn rename tab sang `CF CONTROL`, chỉ làm sau khi search/refactor toàn bộ `Worksheets("INPUT")` và regression pass. Không rename chỉ vì UI.

## 9.2. Ba nút duy nhất

### Nút 1 — `TẢI WH`

Chức năng:

- đọc list nhiều OB,
- bind/reuse đúng 1 MON reader session,
- load all OB snapshots,
- populate hidden WH_ACTUAL,
- build visible UI worklist,
- hide/filter status C,
- không write SAP.

### Nút 2 — `CHECK & XỬ LÝ`

Gộp button CHECK và button xử lý cũ.

Flow:

```text
SYNC UI INPUT
-> VALIDATE
-> CHECK + POSTING READ
-> BUILD/FREEZE PLAN
-> nếu blocker local: mark blocker
-> nếu không có hard blocker: EXECUTE GROUP SCHEDULER
-> stop at READY_FOR_CF / USER_CF_HANDOFF
```

Nếu check fail trước write:

- không action SAP.
- show exact reason in UI.

Nếu plan có independent groups pass + blocked:

- chạy nhóm pass,
- mark nhóm blocked,
- cuối run summary `PARTIAL_PASS_WITH_BLOCKERS`.

### Nút 3 — `LÀM MỚI`

Chỉ clear runtime/user working data:

- OB input list
- visible worklist
- temporary PLAN/QUEUE/CACHE của run hiện tại nếu safe
- UI status/progress

Không xóa:

- RUN_LOG
- RUN_HISTORY
- RUN_SUMMARY
- REGRESSION_AUDIT
- CONFIG
- SAP_MAP
- SAP_ACTION_MAP
- locked fingerprints

Nếu có state `SENT/UNKNOWN` chưa reconcile, `LÀM MỚI` không được xóa evidence; phải cảnh báo và giữ recovery state.

## 9.3. Layout đề xuất

### Top header

- A1:H2 merged title: `CF OB — CONTROL CENTER V2.7`
- subtitle nhỏ: `TẢI WH → NHẬP ACTUAL → CHECK & XỬ LÝ → USER CF`

### Left — OB list

Columns:

- OB
- Ghi chú
- Trạng thái

Cho nhập nhiều OB dọc xuống.

Không thiết kế kiểu nhập từng OB một.

### Right — action panel

3 nút lớn xếp dọc hoặc ngang:

1. TẢI WH
2. CHECK & XỬ LÝ
3. LÀM MỚI

Kèm status cards:

- OB: x/y
- Actionable rows
- Ready
- Blocked
- Waiting CF
- Done C hidden
- Last run time

### Main worklist phía dưới

Hiển thị chỉ các row cần user xử lý, ví dụ:

- OB
- Material
- WT
- Sys Qty
- Sys Bin
- Sys Batch
- Actual Qty **editable**
- Area **editable**
- Bin No. / Actual Bin **editable**
- Actual Batch **editable nếu cần**
- Decision
- State
- Message

Không phơi 30+ cột kỹ thuật.

Hidden engine WH_ACTUAL vẫn giữ full columns.

## 9.4. UI sync

Mỗi visible row phải có hidden `SourceKey` để sync chính xác về WH_ACTUAL.

Không sync theo row number.

Before button 2:

```text
UI SourceKey -> exact WH_ACTUAL SourceKey
```

Nếu duplicate/missing SourceKey -> STOP pre-write.

## 9.5. C rows

Status C:

- không hiện trong main worklist.
- summary card có thể ghi `Đã hoàn thành: N`.
- user không cần thao tác.

## 9.6. Visual style

Professional, gọn, hiện đại:

- dark/navy header,
- white/light work area,
- accent xanh cho ready/pass,
- amber cho waiting/review,
- red chỉ cho blocker/unknown,
- input cells dùng fill nhẹ,
- không quá nhiều màu RGB,
- font rõ, 10–11pt body,
- buttons 14–16pt bold,
- freeze pane trước worklist,
- row height 20–24,
- không để horizontal scroll quá rộng ở sheet chính.

---

# 10. WORKSTREAM H — TỐI ƯU MODULE / CALL PATH

Audit các module/procedure hiện có, đặc biệt:

- `modCodexWTRework`
- `CWR_BuildExecutionGroups`
- `CWR_ExecuteLockedPlanSchedule`
- `CWR_ExecutePreflightGroups`
- `CWR_ExecuteQtyGroups`
- `CWR_SchedulerVerifyPRDOGroup`
- `CWR_ExecuteDocumentGroups`
- `CWR_ExecuteLocationGroups`
- `CWR_ExecuteCreateGroups`
- `CWR_EnsureOldWTCancelledOnce`
- `CWR_EnsurePRDOQtyAtTarget`
- `CWR_EnsurePRDOQtyForDeltaAddOnce`
- `CWR_BatchSetPRDOQtyOneSave`
- `CWR_BatchAllocateCreateRemainder`
- `CWR_BatchAdoptCreateSave`
- `CWR_BatchPostVerifyAndRecord`
- `CWR_GetMONSession`
- `CWR_ResetSchedulerExecutionContext`
- `CWR_LoadExecutionContext`
- `CWR_RecordSuccessfulSchedulerRegressionProofs`
- `modReleaseHolderReader`
- `modCF_MONCore_STABLE`
- `modCF_Posting`
- `modCF_CheckPlanner`
- `modCF_UI`
- `modCF_InputLayout`
- `modSAPActionMap`
- `modSAPActionRegistry`
- `modSAPRunGuard`
- `modWTStateEngine`

### Không làm

- Không clone thêm executor thứ ba.
- Không để legacy single executor và group executor cùng write một action.
- Không tạo function mới chỉ để vá một OB cụ thể.
- Không thêm `On Error Resume Next` che lỗi nghiệp vụ.
- Không thêm loop/retry không bounded.

### Mục tiêu

Một call path production rõ:

```text
UI Entry
-> Precheck/Planner
-> Freeze Plan
-> SessionBroker
-> Group Scheduler
-> Action Registry
-> SAP Adapter
-> Post Verify
-> History
-> UI Summary
```

---

# 11. REPLAY / REGRESSION — BẮT BUỘC TRƯỚC KHI XUẤT V2.7

Không chỉ chạy static audit hiện tại.

Phải tạo replay từ dữ kiện `RUN_HISTORY`, `RUN_LOG`, `PLAN`, `WH_ACTUAL`, `POSTING_CACHE` để cover các lỗi live đã gặp.

## T01 — CF only

- qty/bin/batch unchanged.
- expected: no write, READY_FOR_CF.

## T02 — pure increase single WT

- 580 -> 600.
- source enough.
- create delta 20.

## T03 — pure decrease

- 580 -> 500.
- total final 500.

## T04 — cancel + increase, net positive

- old group 580.
- cancel 20.
- increase/create 60.
- final 620.
- action create must be 60, not net 40.

## T05 — cancel + increase, net negative

- old 580.
- cancel 420.
- increase/create 340.
- final 500.
- action create = 340, not net -80/no-op.

## T06 — **current critical net-zero redistribution**

Input:

```text
WT1 3500 -> target 0
WT2 4500 -> target 8000
Group current = 8000
Group target = 8000
```

Expected:

```text
Cancel WT1 = 3500
Retain WT2 = 4500
PRDO total write = 0 if live total already 8000
RequiredCreateQty WT2 = 3500
Create delta = 3500
Final WT2 allocation sum = 8000
Group final = 8000
PASS
```

Forbidden result:

```text
PRDO_QTY_ALREADY_TARGET -> skip create
STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED
```

## T07 — current replace net-zero

```text
5050 -> 0
12950 -> 18000 + change bin/batch
Group current=18000
Group target=18000
```

Expected:

- cancel both old WTs if replace plan requires,
- POST change exact target,
- create full 18000,
- no duplicate,
- final READY_FOR_CF.

## T08 — two retained increase members same QTY_GROUP

Both need positive delta.

Expected create qty = sum of **member deltas**, but each allocation remains attributable to correct member/source tuple.

## T09 — 3-member mixed group

- one cancel zero
- one retain no change
- one retain + delta

No member incomplete false-stop.

## T10 — one blocked member + independent group same document

- blocked group no write.
- independent group still executes if PRDO row identity isolated.

## T11 — multiple documents in one OB

- one document blocked,
- other document proceeds.

## T12 — multiple OB input list

- OB1 pass
- OB2 blocked data
- OB3 pass

Expected run continues to OB3 unless OB2 creates unknown SAP state.

## T13 — status A resume

Old cancel already verified before rerun.

No cancel resend.

Create remaining qty if required.

## T14 — status C

Hidden from UI, excluded planner, retained audit.

## T15 — change bin only

No unnecessary batch write.

## T16 — change batch only

No unnecessary bin change.

## T17 — change bin + batch

Keep P0006 proven path.

## T18 — source free stock enough

No holder reader/cancel.

## T19 — holder release single WT enough

- NeedRelease positive.
- select smallest single WT >= need.
- one cancel.
- one readback.
- cache update only after A.

## T20 — holder release multiple WT

No single enough.

Select deterministic cumulative candidates.

## T21 — holder A/C excluded

Never auto cancel A/C.

## T22 — same OB unplanned holder

Do not auto cancel; review/block exact source only.

## T23 — holder cancel unknown

No resend; hard-stop whole run.

## T24 — MON already open

No new MON tab.

## T25 — no MON open

Create only one MON tab.

## T26 — MON busy

Bounded wait, no duplicate session.

## T27 — one OB six rows (current case)

Acceptance counters:

```text
MON_REFRESH_COUNT <= 4
DOCUMENT_OPEN_COUNT <= 1 normal or <=2 with one recovery
WHTASK_OPEN_COUNT <= 1 for one document
```

## T28 — rerun after POST sent + verified

No duplicate posting change.

## T29 — rerun after WT Create+Save verified

No duplicate WT.

## T30 — unknown after Create

No create retry.

## T31 — numeric parser regression

Current SAP separators:

- thousands `.`
- decimal `,`

Must correctly parse:

```text
1.200 -> 1200
18.000 -> 18000
2.442 -> 2442
1,2 -> 1.2 if SAP returns decimal
```

Never convert 1200 into 1.2 due separator mistake.

## T32 — blank/zero input rules

- zero is intentional cancel-to-zero.
- blank required Actual Qty is user-input error.
- do not treat blank as zero.

## T33 — duplicate SourceKey

Pre-write STOP.

## T34 — plan fingerprint changed after CHECK

No execute stale plan.

## T35 — C rows hidden UI

Engine history still contains them.

## T36 — button 2 check fails before writes

SAP_WRITE=0.

## T37 — partial pass with blockers

Summary must distinguish:

- PASS
- PARTIAL_PASS_WITH_BLOCKERS
- HARD_STOP_UNKNOWN

Không dùng mọi thứ thành generic `STOP|SCHEDULER_HALTED_OR_BLOCKED|`.

---

# 12. REGRESSION_AUDIT V2.7 — BỔ SUNG CONTRACTS

Existing offline tests T1–T9 đang PASS nhưng chưa bắt được live bug.

Bổ sung scenario contracts ít nhất:

- `OFFLINE_T10_NET_ZERO_MEMBER_REALLOCATION`
- `OFFLINE_T11_PRDO_NO_WRITE_CREATE_DELTA_REQUIRED`
- `OFFLINE_T12_DOCUMENT_MEMBER_ACCOUNTING`
- `OFFLINE_T13_SINGLE_MON_SESSION_BUDGET`
- `OFFLINE_T14_EXTERNAL_HOLDER_RELEASE`
- `OFFLINE_T15_RELEASE_IDEMPOTENCY`
- `OFFLINE_T16_STATUS_C_HIDDEN_NOT_DELETED`
- `OFFLINE_T17_MULTI_OB_CONTINUE_AFTER_LOCAL_BLOCK`
- `OFFLINE_T18_NUMERIC_THOUSANDS_REGRESSION`

### Contract bắt buộc T10

```text
GROUP_CURRENT=8000
GROUP_TARGET=8000
CANCEL=3500
RETAINED=4500
TARGET_MEMBER=8000
PRDO_WRITE=0
CREATE_REQUIRED=3500
FINAL=8000
```

Nếu `CREATE_REQUIRED=0` thì test FAIL.

---

# 13. LOGGING V2.7

Giữ RUN_LOG/RUN_HISTORY nhưng bổ sung stage rõ:

```text
GROUP_LEDGER_BUILT
GROUP_MEMBER_ACCOUNTED
OWN_CANCEL_GROUP_SENT
OWN_CANCEL_GROUP_VERIFIED
RELEASE_QUERY_PASS
RELEASE_SELECTION_PASS
RELEASE_CANCEL_GROUP_SENT
RELEASE_CANCEL_GROUP_VERIFIED
PRDO_TOTAL_READBACK
PRDO_NO_WRITE_REQUIRED
PRDO_DOCUMENT_SAVE_SENT
PRDO_DOCUMENT_SAVE_VERIFIED
MEMBER_CREATE_REQUIRED
MEMBER_CREATE_SENT
MEMBER_CREATE_VERIFIED
GROUP_ALLOCATION_VERIFIED
MON_SESSION_BOUND
MON_SESSION_REUSED
MON_PHASE_REFRESH
DOCUMENT_CONTEXT_REUSED
PARTIAL_GROUP_BLOCKED
OB_PARTIAL_PASS
```

Mỗi log message nên chứa:

- OB
- Document
- QTY_GROUP
- PlanID/member
- WT old/new
- target qty
- retained qty
- cancel qty
- create qty
- source key
- idempotency key

Không log chung chung “STOP” nếu đã biết reason.

---

# 14. RUN SUMMARY V2.7

Bổ sung/chuẩn hóa Result:

### Full pass

`PASS|READY_FOR_CF_HANDOFF`

### Partial

`PARTIAL_PASS_WITH_BLOCKERS`

Detail:

```text
OB=...
GROUPS_PASS=n
GROUPS_BLOCKED=n
READY_FOR_CF=n
UNKNOWN=0
```

### Hard unknown

`HARD_STOP|STATE_UNKNOWN_AFTER_<ACTION>`

### Read-only input blocker

`CHECK_BLOCKED|<reason>`

Không để `SCHEDULER_HALTED_OR_BLOCKED|` rỗng reason như nhiều run hiện tại.

---

# 15. PERFORMANCE MỤC TIÊU

Không tối ưu bằng cách bỏ verify.

Tối ưu bằng:

- group actions,
- reuse session,
- reuse document,
- one MON refresh per phase,
- one PRDO save per document,
- one cancel action per compatible group,
- no POST reload after verified cancel,
- no full SAP map scan in RUN,
- no fixed long sleeps.

### Metrics cần ghi cuối run

- total elapsed
- SAP elapsed
- Excel elapsed
- MON refresh count
- MON tab create count
- document open/reuse count
- WHTASK open/reuse count
- cancel action count
- release cancel action count
- posting save count
- PRDO save count
- WT create count
- WT save count

---

# 16. UI SHEETS / VISIBILITY

Sau V2.7, user normal mode chỉ thấy `INPUT` Control Center.

Các sheet engine nên `VeryHidden` hoặc hidden theo mức phù hợp:

- ORDER_QUERY
- POSTING_RAW
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

Có thể giữ developer/debug unhide routine riêng nhưng **không thêm nút thứ 4 trên Control Center**.

Không xóa sheet vì user cần lịch sử và Codex cần replay.

---

# 17. MIGRATION / COMPATIBILITY

## 17.1. Không phá dữ liệu lịch sử

- giữ toàn bộ RUN_LOG cũ.
- giữ RUN_HISTORY cũ.
- giữ RUN_SUMMARY cũ.
- giữ REGRESSION_AUDIT cũ.

## 17.2. Plan contract version

Current có `CHECK_LOCKED_V2`.

V2.7 nên tăng contract:

`CHECK_LOCKED_V3_MEMBER_ALLOCATION`

Nếu plan cũ V2 được mở bằng V2.7:

- không execute stale locked plan.
- yêu cầu CHECK lại để rebuild ledger V3.

## 17.3. Version labels

Update:

- workbook UI version -> V2.7
- CONFIG version / fast flow version
- regression build name, ví dụ:
  - `V2.7-MEMBER-ALLOCATION-SINGLE-MON-RELEASE-HOLDER`

Không giả vờ LIVE proven cho scenario chưa live test.

---

# 18. AUDIT SAU PATCH

Trước khi xuất file:

## Workbook integrity

- ZIP CRC pass.
- relationships pass.
- sheet names intact.
- VBA project preserved.
- no missing shapes/buttons.
- no broken defined names.
- no merged-cell overlap.
- no `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#N/A` mới.

## VBA

- compile all VBA nếu Excel COM available.
- no duplicate procedure names.
- no unresolved references.
- no accidental `Option Explicit` regression.
- no unbounded loop.
- no hidden SendKeys/coordinate fallback introduced.

## Source diff

Báo rõ module nào thay đổi và lý do.

Không được chỉ nói “đã sửa”.

---

# 19. LIVE-SAFETY POLICY CHO FILE XUẤT RA

File V2.7 output **không tự chạy SAP live trong quá trình build/audit**.

User sẽ test tại công ty.

Trước live test:

- CF auto off.
- first scope = 1 OB.
- one current-bug case ưu tiên:
  - group 2 member net-zero redistribution.
- nếu pass mới tăng scope.

Nếu cần map `RELEASE_MON_*` mà không có SAP live trong môi trường Codex:

- implement mapper/read-only capability,
- để keys pending/unlocked,
- fail closed cho external-holder release,
- nhưng không được giả lập là chức năng đã LIVE ready.

Output note phải nói rõ:

- code complete hay chưa,
- mapping nào còn pending,
- scenario nào static/replay pass,
- scenario nào cần live proof.

---

# 20. DELIVERABLES BẮT BUỘC

Xuất:

1. `260819_CF_OB_V2.7_MEMBER_ALLOCATION_SINGLE_MON_RELEASE_UI.xlsm`
2. ZIP chứa XLSM trên.
3. `V2.7_CHANGELOG.md`
4. `V2.7_TEST_REPORT.md`

Trong changelog phải có:

- root cause V2.6
- modules changed
- group model before/after
- session model before/after
- release-holder status
- UI before/after
- known pending live items

Trong test report phải có:

- T01–T37 status
- PASS/FAIL/NOT_PROVEN
- exact evidence
- no fake PASS for LIVE-only actions.

---

# 21. FORBIDDEN IMPLEMENTATIONS

Tuyệt đối không:

- sửa riêng bằng `If OB = 1011846094 Then ...`
- hard-code WT numbers từ lịch sử.
- hard-code row 5/row 7 như business identity.
- đổi `GROUP_TARGET=GROUP_CURRENT` thành luôn create target.
- create full target cho mọi increase; retained delta phải chỉ create phần thiếu.
- cancel current retained WT chỉ để làm cho logic dễ hơn nếu không cần.
- refresh MON sau từng member.
- mở MON tab mới sau mỗi recovery.
- open PRDO 4–5 lần cho cùng document bình thường.
- bỏ `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE` mà không thay bằng member accounting proof.
- chuyển fail-closed thành silent skip.
- biến `STATE_UNKNOWN` thành PASS.
- auto CF.
- bật delete WT.
- bật MIGO.
- bật coordinate/SendKeys fallback.
- xóa history để “clean file”.
- rebuild workbook từ đầu.

---

# 22. DEFINITION OF DONE CHO V2.7

V2.7 chỉ được coi là hoàn tất development khi:

1. Current critical case `3500+4500 -> 0+8000` replay PASS:
   - cancel 3500,
   - retain 4500,
   - create 3500,
   - no PRDO total write nếu total đã 8000,
   - final 8000.
2. P0005/P0006 style replace path không regression.
3. Không còn false `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED` cho member đã accounted.
4. `GROUP_DELTA_TOTAL=0` không còn là no-op criterion.
5. External holder release có full code path + deterministic selector + idempotency; live mapping status được báo thật.
6. Exactly one MON reader session architecture.
7. Current 1-document case đạt navigation budget mục tiêu.
8. One visible Control Center, 3 buttons.
9. C rows hidden from user worklist nhưng không bị delete.
10. Multi-OB input list hoạt động.
11. Independent blocked group không làm chết toàn bộ queue trừ hard unknown.
12. All replay/static regressions pass hoặc được ghi NOT_PROVEN rõ ràng.
13. Workbook/VBA integrity pass.
14. Output XLSM + ZIP + changelog + test report đầy đủ.

---

# 23. THỨ TỰ THỰC HIỆN ĐỀ NGHỊ CHO CODEX

Làm đúng thứ tự này để tránh patch chồng patch:

### Stage 1 — Snapshot + audit

- copy V2.6.
- hash file.
- export/list VBA modules.
- snapshot sheet visibility, defined names, shapes/buttons.
- parse latest RUN_SUMMARY/RUN_LOG/RUN_HISTORY.

### Stage 2 — Reproduce bug offline

- dựng replay T06 từ run `20260820_052113`.
- chứng minh hiện tại logic cho `PRDO_QTY_ALREADY_TARGET` và member isolation sai.

### Stage 3 — Member allocation ledger

- implement fields/calculation.
- update planner and build groups.
- add T06/T09 tests.

### Stage 4 — PRDO document executor

- separate PRDO total write from member create allocation.
- remove false member-not-included blocker.
- preserve one-save-per-document.

### Stage 5 — Create executor

- consume `RequiredCreateQty`.
- verify retained delta vs replace full target.

### Stage 6 — Session broker

- centralize one MON reader.
- eliminate duplicate MON/session creation paths.
- add navigation counters/budget tests.

### Stage 7 — Release holder

- complete mapper + reader + selector + cancel group + verify + cache update.
- keep fail closed if live mapping unavailable.

### Stage 8 — Multi-OB scheduler behavior

- local blocker continue.
- hard unknown stop.

### Stage 9 — UI

- one visible Control Center.
- 3 buttons.
- SourceKey-based sync.
- hide C.

### Stage 10 — Regression + compile + package

- run T01–T37.
- compile.
- integrity audit.
- output V2.7 package.

---

# 24. BÁO CÁO CUỐI CỦA CODEX PHẢI TRẢ LỜI 10 CÂU NÀY

1. Root cause exact của net-zero group bug là gì?
2. `PRDO_QTY_ALREADY_TARGET` sau sửa còn làm gì và không còn được làm gì?
3. `RequiredCreateQty` được tính như nào cho `KEEP_EXISTING_CREATE_DELTA`?
4. `RequiredCreateQty` được tính như nào cho `REPLACE_CANCEL_CREATE`?
5. Case `3500+4500 -> 0+8000` replay ra action sequence nào?
6. Có còn `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED` false-positive không?
7. Một EXEC run tối đa tạo bao nhiêu MON tab? Current test có bao nhiêu refresh/open?
8. `CANCEL_OTHER_WT` hiện đã complete đến đâu? `RELEASE_MON_*` nào LIVE verified, nào pending?
9. UI có đúng 3 nút và chỉ 1 sheet visible không?
10. Scenario nào PASS static/replay, scenario nào vẫn cần LIVE proof?

Nếu không trả lời được 10 câu này bằng evidence thì chưa coi patch V2.7 hoàn tất.

