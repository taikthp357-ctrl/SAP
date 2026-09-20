# CLAUDE BOOTSTRAP — CF_OB V2.7
## Dùng prompt ngắn này thay cho master prompt dài

Bạn đang làm tiếp dự án CF_OB tại folder hiện tại. Không yêu cầu tôi xác nhận từng bước nhỏ.

## 1. Trước tiên hãy đọc các nguồn theo thứ tự ưu tiên

1. `260820_CF_OB_STATUS_DA_LAM_DANG_DO_CHUA_LAM.md`
   - Đây là checkpoint mới nhất về việc đã làm / đang dở / chưa làm.
   - Số liệu trong file này mới hơn các ước lượng cũ.
   - Dùng nó để biết bước tiếp theo.

2. `PROMPT_CODEX_CF_OB_V2.7_REDESIGN_4BUTTON_5SHEETS_HISTORY.md`
   - Đây là spec kiến trúc V2.7 mới nhất.
   - UI cuối bắt buộc 4 nút:
     `TẢI WH -> CHECK -> XỬ LÝ -> LÀM MỚI`
   - `CHECK` = READ ONLY.
   - `XỬ LÝ` mới được phép ghi SAP.
   - Mục tiêu cuối: 5 sheet logic:
     `INPUT, DATA_CORE, PLAN, SYSTEM, RUN_HISTORY`
   - `POSTING` phải được loại khỏi runtime sau khi migration parity đã được chứng minh.

3. `CLAUDE_OFFLINE_TEST_MATRIX_CF_OB_V2.7.md`
   - Đây là bộ regression offline.
   - Phải dùng mock/replay, không cần SAP thật.
   - Không được chuyển sang SAP LIVE nếu mandatory release gate chưa PASS.

4. Toàn bộ `.zip`, `.md`, `.txt`, workbook cũ và conversation export trong folder.
   - Dùng để lấy lịch sử, live evidence, failed attempts và regression fixtures.
   - Không được để tài liệu cũ ghi đè spec mới.

5. Workbook base:
   - V2.6 hiện tại là BASE.
   - V2.5 = superseded.
   - V2.7 = spec sẵn sàng, code phát triển tiếp từ COPY V2.6.

## 2. Thứ tự authority khi tài liệu xung đột

Ưu tiên:

1. Prompt bootstrap này.
2. `260820_CF_OB_STATUS_DA_LAM_DANG_DO_CHUA_LAM.md`
3. spec V2.7 4-button/5-sheet mới nhất.
4. workbook V2.6 hiện tại + lịch sử chạy thật.
5. conversation/ZIP Codex/Claude.
6. tài liệu UI/logic cũ.

Không phục hồi lại:
- 2 nút;
- 3 nút;
- `CHECK + POSTING`;
- `CHECK & XỬ LÝ`;
- workflow P000x tự chạy độc lập;
nếu chúng xung đột với spec mới.

## 3. Việc cần làm ngay

Tiếp tục đúng checkpoint hiện tại.

Bắt đầu từ **§12 Bước 1**:

### Fix Allocation Engine trên một COPY của V2.6

Phải xử lý triệt để ba contract:

```text
CONTRACT_A = Group/Document Total
CONTRACT_B = Member/WT Allocation
CONTRACT_C = User Intent
             blank = KEEP
             0     = CANCEL
             >0    = EXPLICIT TARGET
```

Case bắt buộc:

```text
3500 + 4500 -> 0 + 8000
```

Expected:

```text
member A:
  cancel 3500

member B:
  retain 4500
  create delta 3500

group:
  current total = 8000
  target total  = 8000
  net delta     = 0
  allocation write still required = TRUE
```

Không được dùng:
`GROUP_CURRENT = GROUP_TARGET`
để kết luận toàn group không còn action.

## 4. Blank khác 0

Bắt buộc test:

```text
3500 + 4500 -> blank + 8000
```

Expected:

```text
first member KEEP 3500
second member target 8000
group target = 11500
create delta second member = 3500
```

Blank không được:
- biến thành 0;
- bị bỏ khỏi group;
- bị coi là member không tồn tại.

## 5. Không làm UI/5-sheet migration trước khi Allocation Engine pass

Thứ tự phát triển:

1. Allocation Engine
2. Qty/group regression
3. CHECK / XỬ LÝ split
4. Group scheduler
5. Single MON broker
6. Holder Release PLAN_ONLY
7. History V2
8. Migration 20 -> 5 sheets
9. UI final
10. Full regression
11. Package

Không nhảy thẳng sang UI khi logic allocation chưa PASS.

## 6. Cách làm việc

Không dừng ở plan.

Loop:

```text
READ
-> IMPLEMENT
-> STATIC/COMPILE CHECK
-> RUN OFFLINE TEST
-> IF FAIL:
     inspect exact failure
     patch root cause
     rerun failed category
     rerun mandatory regression
-> continue
```

Không yêu cầu user nói "continue" sau mỗi lỗi thông thường.

## 7. Quy tắc test không SAP

Dùng `CLAUDE_OFFLINE_TEST_MATRIX_CF_OB_V2.7.md`.

Mỗi test phải kiểm:

- InputSnapshot
- MemberPlan
- GroupPlan
- ReleasePlan
- ActionPlan
- ExpectedHistory
- FinalState

Không PASS chỉ vì code không crash.

### Mandatory trước khi sang SAP LIVE

Tối thiểu các nhóm sau phải PASS:

- parser quantity
- blank / zero / KEEP
- A / C status
- net-zero multi-WT allocation
- replace mode
- cancel idempotency
- source stock
- holder PLAN_ONLY
- CHECK read-only
- stale plan
- action UNKNOWN
- recovery
- Create WT delta
- final member verify
- session broker
- history evidence

Nếu còn reproducible offline failure:
không test SAP LIVE.

## 8. Holder Release

Hiện chưa ổn định.

Default:

```text
RELEASE_HOLDER_MODE = PLAN_ONLY
```

CHECK được phép:
- đọc holder;
- chọn candidate;
- tính coverage;
- lập release plan.

CHECK không được Cancel.

Khi sau này test PILOT:
- exact Material + Bin + Batch;
- exact Holder WT;
- Cancel chỉ một lần;
- MON readback phải chứng minh Status=A;
- nếu UNKNOWN thì không retry.

## 9. CHECK / XỬ LÝ

Final UI vẫn phải tách:

```text
TẢI WH
CHECK
XỬ LÝ
LÀM MỚI
```

CHECK:
- read only;
- build locked PlanID;
- zero SAP business write.

XỬ LÝ:
- chạy đúng locked plan;
- không tự rebuild CHECK;
- nếu fingerprint stale:
  `STOP_PLAN_STALE_RECHECK_REQUIRED`

## 10. Resume / idempotency

Action states:

```text
PLANNED
SENT
VERIFIED
UNKNOWN
SKIPPED_VERIFIED
```

Rules:

- VERIFIED -> không resend.
- UNKNOWN -> không resend.
- UNKNOWN -> recovery/readback trước.
- Crash sau Create -> tìm created WT trước retry.
- Crash sau Save -> reconcile state trước Save lần nữa.

## 11. Single MON

Một broker duy nhất sở hữu MON.

Không module nào tự:
- `/SCWM/MON`
- reopen MON theo từng P/member.

Mục tiêu normal:
- MON open <= 1
- refresh theo phase
- không reopen per P000x.

## 12. History

History cuối phải cho biết chính xác:

- LoadRunID
- CheckID
- PlanID
- ExecuteID
- ActionID
- before
- target
- after
- SENT?
- VERIFIED?
- UNKNOWN?
- StopCode
- ResumeFrom
- duration
- SAP session/control evidence

Không log polling spam.

## 13. UI và sheet optimization

Sau khi logic pass:

Final user-facing UI:
- chỉ `INPUT` visible;
- multi-OB queue;
- actionable worklist;
- C không hiện trong worklist;
- editable cells rõ;
- warning release/recovery rõ;
- 4 buttons;
- không MsgBox spam.

Final logical sheets:

```text
INPUT
DATA_CORE
PLAN
SYSTEM
RUN_HISTORY
```

`POSTING`:
- không chỉ hide;
- phải search dependency;
- migrate reader/writer;
- regression;
- zero runtime reference;
- rồi mới delete.

## 14. Không auto CF

Final state:

```text
READY_FOR_USER_CF
```

Không tự confirm.

## 15. External boundary

Nếu môi trường không có:
- Excel/VBE compile;
- SAP GUI;
- approved live test target;

thì:
- hoàn thành tất cả offline/static/replay tests;
- không fake live PASS;
- ghi `BLOCKED_EXTERNAL`.

## 16. Cập nhật checkpoint sau mỗi milestone

Luôn cập nhật:

`260820_CF_OB_STATUS_DA_LAM_DANG_DO_CHUA_LAM.md`

Ghi:
- đã làm;
- đang dở;
- chưa làm;
- test PASS/FAIL;
- current artifact;
- next exact step.

Không tạo checkpoint trùng section hoặc append rác.

## 17. Output cuối

Khi V2.7 hoàn thành, trả:

- `260820_CF_OB_V2.7.xlsm`
- package ZIP
- changelog
- migration audit
- regression report
- UI decisions
- final audit
- live test report nếu có SAP LIVE

## 18. Bắt đầu

Đọc checkpoint trước.
Sau đó đọc spec V2.7.
Sau đó đọc offline test matrix.
Sau đó inspect COPY V2.6.

**Tiếp tục từ §12 Bước 1 — Allocation Engine.**
Không dán lại toàn bộ tài liệu vào context.
Đọc file khi cần, làm theo checkpoint và cập nhật tiến độ tại chỗ.
