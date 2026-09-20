# TRẠNG THÁI DỰ ÁN CF OB - CHỐT NGÀY 2026-08-20 (BASE = V2.6, NÂNG CẤP -> V2.7)

> File chốt (checkpoint) duy nhất để dừng lại và biết: việc ĐÃ LÀM, việc CHƯA LÀM, việc ĐANG DỞ.
> Nhãn trạng thái dùng trong file: **DONE** / **IN-PROGRESS** / **NOT-DONE** / **BLOCKER**.
> Nguyên tắc số liệu: chỉ dùng con số đã xác minh từ 4 báo cáo audit + lịch sử chạy thật; số nào chưa xác minh ghi rõ "chưa xác minh". Không bịa.

---

## 0. TL;DR - đang ở đâu, làm gì tiếp

1. **BASE HIỆN TẠI = V2.6** (`260819_CF_OB_V2.6_LICH_SU.xlsm`). Đây là bản đã chạy thật với SAP, có lịch sử EXECUTE, có các luồng proven live (P0005/P0006, fast-path tạo WT, DELTA_ADD, CANCEL_TO_ZERO). Mọi nâng cấp/tối ưu/test tiếp theo XUẤT PHÁT TỪ V2.6.
2. **V2.5 (nhánh refactor trong `build\`) ĐÃ BỎ / SUPERSEDED** - nội dung đã gộp vào V2.6. KHÔNG resume, KHÔNG chạy lại harness replay V2.5, KHÔNG dùng `CF_OB_V2.5_WORKING.xlsm` làm base nữa. Chỉ còn là di sản (xem Phụ lục §13).
3. **Đặc tả V2.7 đã SẴN SÀNG** (3 file PROMPT + 1 file AUDIT, ngày 8/20) nhưng **CODE V2.7 CHƯA bắt đầu trên base V2.6** - đây là toàn bộ phần ĐANG DỞ.
4. **Lỗi lõi phải diệt đầu tiên:** `3500+4500 -> 0+8000`. GROUP_DELTA=0 nhưng vẫn phải CREATE 3500. Nguyên tắc: **GROUP_DELTA=0 KHÁC ACTION_DELTA=0** (xem §5). Đây là regression contract số 1 của V2.7.
5. **Lỗ hổng lớn nhất còn lại:** Holder Release (giải phóng tồn do WT khác giữ) CHƯA chạy được lần nào - 15 map `RELEASE_MON_*` trống, `PLAN_RELEASE` rỗng, 22 kế hoạch kẹt `BLOCKED_RELEASE_REQUIRED`.
6. **Bước 1 khi tiếp tục:** tách `ACTION_DELTA` khỏi `GROUP_DELTA` (Member Allocation Engine) trên bản COPY của V2.6, chứng minh case `3500+4500->0+8000` PASS offline. KHÔNG có bước nào là resume V2.5.
7. **Lưu ý danh xưng:** "V2.6" là tên người dùng đặt cho file base. Build tag kỹ thuật thật bên trong artifact đọc ra là `V2.4-GROUP-EXECUTOR-WHTASK-REUSE` + `V2.3-CORE-SCHEDULER-CONTRACT` (chuỗi "V2.6"/"V2.5" không tồn tại trong artifact). Cần thống nhất cách gọi khi đối chiếu bằng chứng.

## 1. Bản đồ các nhánh

| Nhánh | File / Vị trí | Vai trò | Trạng thái |
|---|---|---|---|
| **V2.6** | `G:\SAP\CF_OB\260819_CF_OB_V2.6_LICH_SU.xlsm` | **BASE HIỆN TẠI** - bản chạy thật với SAP, có lịch sử EXECUTE, nơi lộ bug thiết kế group VÀ nơi chứa các luồng proven live | **DONE (đang dùng làm base)** |
| V2.4 gốc | `G:\SAP\CF_OB\260812_CF_OB_V2.4.xlsm` | Mốc lịch sử. KHÔNG được sửa | **FROZEN (mốc gốc)** |
| V2.5 refactor | `G:\SAP\CF_OB\build\` (7 module `modCFX_*`, harness, evidence, `CF_OB_V2.5_WORKING.xlsm`) | Deterministic Group State Machine (skeleton REPLAY-proven) | **SUPERSEDED / đã gộp vào V2.6 - KHÔNG dùng nữa** |
| **V2.7 redesign** | 3 file PROMPT + `260820_CF_OB_V2.6_AUDIT_TO_V2.7.md` | **ĐÍCH NÂNG CẤP TỪ V2.6**: Member Allocation Engine + Group Scheduler theo phase + Single MON Broker + Holder Release + Resume/Idempotency + History V2 + 5 sheet + 4 nút | **NOT-DONE (đặc tả xong, code chưa bắt đầu)** |

**Ghi chú danh xưng / baseline (BLOCKER nhẹ cần xác nhận):**
- File constraints ghi "production baseline duy nhất: `260819_CF_OB_V2.6.xlsm` (không dùng bản `_LICH_SU`/`_BACKUP` làm production)". Nhưng trên đĩa **chỉ có** `260819_CF_OB_V2.6_LICH_SU.xlsm`, KHÔNG thấy `260819_CF_OB_V2.6.xlsm`. Người dùng đã chỉ định dùng bản `_LICH_SU` làm base -> đi theo chỉ định người dùng, nhưng **cần xác nhận** đây đúng là bản mới nhất trước khi copy để nâng cấp.
- "V2.6" là nhãn người dùng; build tag kỹ thuật đọc ra từ log là `V2.4-GROUP-EXECUTOR-WHTASK-REUSE`.

Các file V2.7 (đặc tả đích, ngày 8/20):

| File | Vai trò | Mô hình nút |
|---|---|---|
| `G:\SAP\CF_OB\260820_CF_OB_V2.6_AUDIT_TO_V2.7.md` | Audit chuyển tiếp V2.6 -> V2.7 | đề xuất 3 nút |
| `G:\SAP\CF_OB\PROMPT_CODEX_CF_OB_V2.7_ULTRA_DETAIL.md` | Đặc tả chi tiết | 3 nút (CHECK & XỬ LÝ gộp) - **giữ 20 sheet VeryHidden** |
| `G:\SAP\CF_OB\PROMPT_CODEX_CF_OB_V2.7_REDESIGN_4BUTTON_5SHEETS_HISTORY.md` | **ĐẶC TẢ ĐÍCH (chuẩn)** | **4 nút tách rời + 5 sheet + xóa POSTING** |
| `G:\SAP\CF_OB\PROMPT_CLAUDE_CF_OB_V2.7_ALL_IN_ONE_UNTIL_PASS.md` | Prompt thực thi "until pass" | theo REDESIGN |
| `G:\SAP\CF_OB\CF_OB_CONVERSATION_HISTORY_20260820.md` | Hội thoại Codex đã xử lý | tài sản |
| `G:\SAP\CF_OB\SAP_CONVERSATIONS_STATES_MAPS_PASSED_20260820.md` | State + map SAP đã pass | tài sản |
| `G:\SAP\CF_OB\outbound-wt-execution-learnings-20260817.md` | Bài học tạo WT outbound | tài sản |

> **MÂU THUẪN 3 file (phải theo REDESIGN):** file REDESIGN (§0) EXPLICITLY viết "BỎ THIẾT KẾ 3 nút" và bắt buộc **ĐÚNG 4 NÚT tách rời** (TẢI WH / CHECK / XỬ LÝ / LÀM MỚI) + **đúng 5 sheet** + **xóa POSTING**. REDESIGN override 2 file kia. Đích V2.7 = mô hình REDESIGN.

## 2. ĐÃ LÀM (DONE) - kèm bằng chứng

Trọng tâm: V2.6 đã chạy được gì với SAP THẬT + tài sản đã tích lũy. (KHÔNG tính công việc V2.5 là thành tựu cần dùng tiếp.)

### 2.1. Pipeline CHECK -> EXECUTE chạy live thật, ghi SAP thật - **DONE**
- **Bằng chứng:** 41 RunID phân biệt trong 1 ngày 2026-08-19 = **27 CHECK + 14 EXECUTE** (không phải ~45 EXECUTE như trí nhớ ban đầu). 11 OB phân biệt được EXECUTE. Nguồn: `RUN_HISTORY.csv`, `RUN_SUMMARY.csv`.
- 35 dòng có `CANCEL_VERIFIED` / `OLD_WT_STATUS=A` / `CREATE_SAVE_SENT` / `POST_PICK_BIN_CREATE_SAVE_SENT`. Cancel readback chứng minh WT `5042889690` -> status A; `BATCH_PRDO_POSTVERIFY` OK sau 1 Save; `POST_QUAN=640`. Nguồn: `evidence/latest_run_key_findings.md`.

### 2.2. Group executor chạy live qua nhánh P0001-P0008 (gồm P0005/P0006) - **DONE**
- **P0005** (DELTA_ADD / KEEP_EXISTING_CREATE_DELTA): 23 READ_ONLY PASS + 5 WT_CREATE_SAVE PASS.
- **P0006** (CREATE+SAVE): 14 READ_ONLY PASS + 5 WT_CREATE_SAVE PASS.
- **P0007** (FINAL_TOTAL / CANCEL_PRDO_ZERO): 16 READ_ONLY + WT_CANCEL/WT_CREATE_SAVE.
- **P0001** (REPLACE_CANCEL_CREATE): đủ vòng đời WT_CANCEL / PRDO_QTY / POST_CHANGE / WT_CREATE_SAVE / WT_ADOPT_QUANTITY.
- Nguồn: `RUN_HISTORY.csv` (EXECUTE-channel).

### 2.3. Các luồng nghiệp vụ đã proven live (giữ nguyên, KHÔNG regression) - **DONE**
| Luồng | Bằng chứng | Trạng thái map |
|---|---|---|
| Fast-path tạo outbound WT | Tải WH đọc MON -> nhập Qty/Bin/Batch -> CHECK sinh PLAN (SAP_WRITE=0) -> tab "Stk Can Be Removed" match Product+Bin+Batch -> ghi `UI_VSOLA` readback -> chọn `SelectedRows` readback -> Adopt (`OK_OD_AVSTKDLV_TAKE_QTY`) -> Create (`OK_OI_DLV_CREATE_TO`) ROI Save (`OK_OI_DLV_SAVE`) -> MON post-verify -> READY_FOR_CF | LIVE_PASS |
| WT `5042711577` | Tạo outbound WT thành công, MON post-verify PASS, CF không gửi | LIVE_PASS |
| DELTA_ADD (V1.5/CF3) | Giữ WT cũ `5042769777` Qty 60, đổi PRDO -> 280, tạo WT mới `5042833176` Qty 220 (bin APFL-P22-02-01 batch 260809-10), post-verify PASS, CF_NOT_SENT | proven |
| CANCEL_TO_ZERO (CF3) | Cancel WT exact -> readback -> PRDO Qty=0 -> readback 0 -> không tạo WT mới, không CF | proven |
| Posting-change bin/batch | Material BIW32026215AA, Qty 128, bin APFL-ZF03-01-13, batch 260713-07 -> 260704-07, WR 1012034149 | proven (nghiệp vụ) |
| V2.6 RUN_SUMMARY PASS | OB `1011832369` và `1011839385` đạt `PASS\|LOCKED_PLAN_PHASES_COMPLETE` (C=0, executor xong checkpoint khóa, chưa CF) | 2 PASS cấp RUN |

### 2.4. Guard fail-closed hoạt động đúng - **DONE**
- `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE` fire live thật (run 053709). `BATCH_CREATE_ACTIVE_OUTBOUND_WT_PRESENT` fire ở run 054133 (WT=`5042889690`).
- Phân bố kết quả EXECUTE run: **2 PASS + 4 SCHEDULER_HALTED_OR_BLOCKED + 2 STATE_UNKNOWN + 2 OB_REGRESSION_GATE_BLOCKED** (+ 2 CHECK_FAIL). Nguồn: `RUN_SUMMARY.csv`.

### 2.5. Bộ map SAP đã xác nhận (tài sản dùng lại cho V2.7) - **DONE**
- Golden source: sheet `SAP_ACTION_MAP` trong `260819_CF_OB_V2.6.xlsm`. Status 27 action: **VERIFIED=13, PENDING=6, LIVE_PASS=5, DISCOVERED=3**.
- Map VERIFIED/LIVE_PASS: `MON_WT_GRID_ID`, `MON_WT_ACTION_ID=N000000096`, `WT_REQUEST_GRID`, `WT_TAB_STK_CAN_BE_REMOVED`, `WT_STOCK_GRID`, `WT_STOCK_SET_QTY=UI_VSOLA` (LIVE_PASS), `WT_STOCK_SELECT_ROW=SelectedRows`, `WT_ADOPT_QUANTITY=OK_OD_AVSTKDLV_TAKE_QTY` (LIVE_PASS), `WT_CREATE_COMMAND_ID=OK_OI_DLV_CREATE_TO` (LIVE_PASS), `WT_SAVE_COMMAND_ID=OK_OI_DLV_SAVE` (LIVE_PASS), `WT_CREATE_SAVE` (LIVE_PASS/LOCKED). Chi tiết ở §7.

### 2.6. Harness regression offline V2.4 đã chạy đủ - **DONE**
- `REGRESSION_AUDIT.csv`: **697 PASS / 376 NOT_PROVEN / 70 NOT_APPLICABLE / 9 FAIL**; Build `V2.4-GROUP-EXECUTOR-WHTASK-REUSE` (1139 dòng) + `V2.3-CORE-SCHEDULER-CONTRACT` (162). Offline T1-T9 đều PASS.
- Cảnh báo độ phủ: `STATIC_PROVEN=597` vs `LIVE_PROVEN` chỉ **4** -> phần lớn bằng chứng là static, live còn mỏng (xem §6).

### 2.7. Tài sản Codex / prompt / hội thoại proven live - **DONE (đã có, phải tận dụng)**
- Đặc tả V2.7 đầy đủ (3 PROMPT + AUDIT), hội thoại Codex đã xử lý với SAP thật, state + map đã pass. Đây là tài sản chính để dựng engine V2.7, không được bỏ. Vị trí file ở §1 và §13.

## 3. ĐANG DỞ (IN-PROGRESS) - chuyển V2.6 -> V2.7

> Phạm vi "đang dở" = quá trình chuyển tiếp. **Đặc tả/prompt đã sẵn sàng, nhưng CODE V2.7 CHƯA bắt đầu trên base V2.6.** (V2.5 KHÔNG thuộc phần này - đã bỏ.)

### 3.1. Đặc tả V2.7 - **IN-PROGRESS (đặc tả DONE, hiện thực chưa)**
- **Dừng ở đâu:** 3 file PROMPT + AUDIT đã viết xong ngày 8/20, mô hình đích đã chốt (4 nút / 5 sheet / Allocation Ledger / Single MON / Release gated / History V2). Mâu thuẫn giữa 3 file đã được giải quyết (lấy REDESIGN làm chuẩn).
- **Bước tiếp theo:** copy V2.6 -> bắt đầu Stage 1 (snapshot + audit workbook), rồi hiện thực Member Allocation Engine trước (vì nó chặn bug lõi §5).

### 3.2. Trích xuất module map từ V2.6 (staging) - **IN-PROGRESS / BLOCKER xác minh**
- **Dừng ở đâu:** Audit nhắc tới thư mục staging `_codex_patch\v26_current_extract_20260820\` chứa `modSAPActionMap.bas`, `modSAPActionRegistry.bas`, `modCF_SAP_MAP_MON.bas`, `modCFEvidenceV2.bas` (extract từ V2.6 để dùng làm nguồn map).
- **BLOCKER:** tại thời điểm chốt file này, **thư mục `_codex_patch\...` KHÔNG tìm thấy trên đĩa** (đã quét toàn `G:\SAP\CF_OB`). => Trạng thái "đã extract" **chưa xác minh**. Golden source vẫn là sheet `SAP_ACTION_MAP` trong workbook V2.6.
- **Bước tiếp theo:** hoặc tạo lại extract từ V2.6, hoặc xác nhận vị trí thật của staging trước khi coi map đã "import". Lưu ý: extract chỉ là staging, **chưa tự chứng minh đã import**.

### 3.3. Diệt bug lõi net-zero (§5) - **NOT-DONE nhưng là điểm khởi động của IN-PROGRESS**
- **Dừng ở đâu:** đã xác định đầy đủ root cause + nguyên tắc sửa + regression contract. Chưa viết code Member Allocation Ledger trên V2.6.
- **Bước tiếp theo:** xem §12 bước 1.

> **Không đưa V2.5 vào mục này.** Nhánh `build\` (harness replay S01-S22, `CF_OB_V2.5_WORKING.xlsm`) đã SUPERSEDED; các câu "đang dở" trong báo cáo build V2.5 (chứng minh REPLAY_ALL_PASS bằng chuỗi summary literal, đồng bộ workbook<->đĩa) **không còn là việc cần làm tiếp** - chỉ là bằng chứng cho biết V2.5 tới đâu để khẳng định nó đã gộp vào V2.6 (xem §13).

## 4. CHƯA LÀM (NOT-DONE) = phạm vi engine V2.7 chưa xây trên V2.6

Toàn bộ engine V2.7 chưa hiện thực trên base V2.6. Liệt kê theo hạng mục:

| # | Hạng mục | Trạng thái | Vì sao / bằng chứng |
|---|---|---|---|
| 4.1 | **Member Allocation Engine / Ledger** (cancel-X / retain-Y / create-Z trong 1 group net-zero) | **NOT-DONE / BLOCKER #1** | V2.6 mới freeze group current/target/delta; chưa có member ledger nội-group. Đây chính là gốc bug `3500+4500->0+8000` và STD20000556 8000->8000. |
| 4.2 | **Holder Release Engine** (CANCEL_OTHER_WT / NeedRelease / PLAN_RELEASE / holder query + selector + coverage + idempotency + readback) | **NOT-DONE / BLOCKER #2** | 15 map `RELEASE_MON_*` TRỐNG; `PLAN_RELEASE.csv` RỖNG (chỉ header); 22 kế hoạch kẹt `BLOCKED_RELEASE_REQUIRED`. Tính năng coi như chưa chạy được lần nào. `RELEASE_EXCLUDE_STATUS="A,C"` còn tự loại chính status A (holder cần cancel) -> mâu thuẫn. |
| 4.3 | **Single MON Session Broker** (đúng 1 MON reader/EXEC run, roles ROLE_MON_READER + ROLE_EXECUTOR, mọi reader reuse broker) | **NOT-DONE** | V2.6 còn reuse tab PRDO làm MON (`CWR_ReuseVerifiedPRDOTabForMON`) -> root cause thrash. Navigation đỉnh thật: DOC_OPEN=16, MON_REFRESH=16, DOC_RECOVERY=3. |
| 4.4 | **Group Scheduler theo phase cố định** (9 phase: PRECHECK/LOCK -> CANCEL_OWN -> MON_READBACK -> RELEASE_HOLDER -> MON_READBACK -> POST_CHANGE -> PRDO_DOCUMENT -> CREATE -> FINAL_VERIFY -> USER_CF) | **NOT-DONE** | V2.6 chạy từng P000x mở/đóng OB rời rạc; chưa build plan đầy đủ trước rồi chạy theo `tblActionPlan`. |
| 4.5 | **Action Ledger** (`PLAN.tblActionPlan`, 1 row = 1 SAP action dự kiến, ActionState PLANNED->SENT->VERIFIED->UNKNOWN->SKIPPED) | **NOT-DONE** | Chưa có bảng action plan làm nguồn duy nhất scheduler chạy theo. |
| 4.6 | **Resume / Idempotency / Recovery Engine** (ActionID = tuple ổn định; UNKNOWN -> RecoveryRequired + hard-stop; mở workbook thấy Recovery thì chặn XỬ LÝ) | **NOT-DONE** | Chưa có state machine action + ActionID tuple. 2 run V2.6 kết STATE_UNKNOWN chưa reconcile được. |
| 4.7 | **Document Executor + Member Accounting** (states ACCOUNTED_CF_ONLY/CANCEL_VERIFIED/PRDO_READBACK_ONLY/PRDO_WRITE_REQUIRED/CREATE_PENDING/CREATE_VERIFIED/BLOCKED/TERMINAL_C) | **NOT-DONE** | Cần thay `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED` bằng accounting đầy đủ, chỉ STOP khi MEMBER_UNACCOUNTED thật. |
| 4.8 | **Central Permission Gate** `CanExecuteAction(...)` (gộp 2 hệ safety đang mâu thuẫn: ALLOW_* vs REWORK_ACTION_WHITELIST/FAST_FLOW) | **NOT-DONE** | Hiện có 2 hệ cờ song song dễ vênh. |
| 4.9 | **Data / SourceStock Repository in-memory** + `DATA_CORE.tblSourceStock` (thay chuỗi SAP->POSTING_RAW->POSTING->POSTING_CACHE) | **NOT-DONE** | POSTING chưa gỡ; xem §10. |
| 4.10 | **History Event Ledger** `tblRunEvents` 65 cột + batch buffer 20-100 event | **NOT-DONE** | Hiện còn 4 sheet RUN_LOG/RUN_HISTORY/RUN_SUMMARY/REGRESSION_AUDIT riêng. |
| 4.11 | **Number Parser canonical** `modCFNumberParser` (Double/Decimal-safe) | **NOT-DONE (V2.7)** | V2.6 có contract parser MIGO số nguyên; V2.7 cần parser T31 riêng (1.200->1200, 18.000->18000, không bao giờ 1200->1.2). |
| 4.12 | **Mô hình 5 sheet + xóa POSTING** | **NOT-DONE** | V2.6 đang ~20 sheet; xem §10. |
| 4.13 | **UI 4 nút tách rời + 1 sheet visible + button state machine** | **NOT-DONE** | V2.6 chưa tách CHECK khỏi XỬ LÝ theo REDESIGN. |
| 4.14 | **LIVE wiring cho engine mới** | **NOT-DONE** | Toàn bộ engine V2.7 chưa tồn tại nên chưa nói tới nối dây live. |
| 4.15 | **Deliverables V2.7** (`260820_CF_OB_V2.7.xlsm` + ZIP + CHANGELOG + MIGRATION_AUDIT + REGRESSION_REPORT T01-T37 + before/after audit) | ✅ **DONE (offline)** | Sinh xong tại `work\OUTPUT\`: `260820_CF_OB_V2.7.xlsm` (SHA `3ba75a86…c590b0`) + `260820_CF_OB_V2.7_PACKAGE.zip` (10 entries) + 9 báo cáo `DELIVERABLES\V2.7_*.md`. FINAL_AUDIT xác minh THẬT trên file OUTPUT: compile PASS, ZIP/XML valid, 2 module mới present, lock cells intact. `v27_final_audit.log`. Status = `OFFLINE_PASS_LIVE_BLOCKED_EXTERNAL`. |

**Các khoảng trống live cụ thể của V2.6 (chưa chứng minh):**
- **NOT-DONE** Chưa có EXECUTE run nào hoàn tất group nhiều member CÓ kèm release (2 run PASS đều là plan 1 member, không cần release).
- **NOT-DONE** 2 run kết `STATE_UNKNOWN` sau CREATE_SAVE chưa giải quyết: run 120381 (`STATE_UNKNOWN_AFTER_CREATE_SAVE`, GROUP=70 BIW32023623AB) và run 839588.
- **NOT-DONE** Thrash navigation chưa dứt điểm: Create bị chặn trên chính WT `5042889690` vừa cancel-verified (run 054133); DOC_OPEN đỉnh 16.
- **NOT-DONE** Regression gate tự chặn 2 EXECUTE run (`STOP_REGRESSION_GATE_FAIL`): run 1133468 (CREATE_WT_DELTA) và run 1450070 (MIXED_OB) do PRIOR_LIVE=3 < REQUIRED 5-6.

## 5. Lỗi lõi `3500+4500 -> 0+8000` (regression contract SỐ 1 của V2.7)

### 5.1. Hiện tượng
Một group có 2 WT:
- **WT1:** 3500 -> 0 (user cancel về 0).
- **WT2:** 4500 -> 8000 (giữ 4500 + cần tạo thêm 3500).

Tổng group **trước = 8000**, **đích = 8000**.

### 5.2. Vì sao V2.6 sai
V2.6 suy ra `GROUP_DELTA = GroupTarget - GroupCurrent = 8000 - 8000 = 0`
-> kết luận `NO_QTY_WRITE`
-> log `PRDO_QTY_ALREADY_TARGET`
-> **BỎ MẤT CREATE 3500** (coi member đã complete).

### 5.3. Nguyên tắc bắt buộc phải sửa
> **GROUP_DELTA=0 KHÁC ACTION_DELTA=0.**

- Cấm suy ra `AllocationWriteRequired=FALSE` từ net delta 0.
- Mỗi group phải có **2 hợp đồng ĐỘC LẬP**, chỉ hoàn tất khi CẢ HAI pass:
  - **Contract A** = PRDO aggregate total (`PRDO_LIVE_TOTAL` / `PRDO_TARGET_TOTAL` / `PRDO_WRITE_REQUIRED`).
  - **Contract B** = WT member allocation (cancel-X / retain-Y / create-Z từng member).
- `PRDO_QTY_ALREADY_TARGET` CHỈ nghĩa "không cần write tổng PRDO"; **KHÔNG** được set member complete nếu `RequiredCreateQty>0`. Sau log này phải tiếp tục member allocation (`PRDO_WRITE_REQUIRED=FALSE` nhưng `MEMBER_ALLOCATION_CONTINUES=TRUE`).
- Công thức đúng: `RequiredCreateQty = max(TargetQty - EffectiveQtyAfterPlannedCancelReplace, 0)`.
  - `KEEP_EXISTING_CREATE_DELTA` -> create phần thiếu (retain 4500, create 3500).
  - `REPLACE_CANCEL_CREATE` (sau cancel VERIFIED) -> create full TargetQty (vd 18000).
  - **TUYỆT ĐỐI KHÔNG** dùng `GroupTarget - GroupCurrent` hay `TargetQty - GroupLiveTotal`; không create full target cho mọi increase.

### 5.4. Kết quả bắt buộc (đích)
```
CANCEL 3500  ->  RETAIN 4500  ->  PRDO_WRITE=0 (total đã 8000)  ->  CREATE 3500  ->  FINAL MEMBER=8000  ->  GROUP=8000  =>  PASS
```

### 5.5. Là regression contract số 1
- Contract cứng: **T06** = **OFFLINE_T10_NET_ZERO_MEMBER_REALLOCATION**:
  `GROUP_CURRENT=8000 / GROUP_TARGET=8000 / CANCEL=3500 / RETAINED=4500 / TARGET_MEMBER=8000 / PRDO_WRITE=0 / CREATE_REQUIRED=3500 / FINAL=8000`.
  **Nếu `CREATE_REQUIRED=0` -> test FAIL.**
- FORBIDDEN: `PRDO_QTY_ALREADY_TARGET -> skip create` hoặc `STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED`.

### 5.6. Bằng chứng lớp lỗi (số literal 3500+4500 KHÔNG có trong log - là mô tả LỚP lỗi)
Đã grep toàn bộ log: **không có** chuỗi `3500+4500->0+8000`. Bằng chứng thật cho LỚP lỗi "net-delta != tổng action member":
- `OFFLINE_T4_CANCEL420_INCREASE340`: NET_GROUP_DELTA=-80 nhưng ACTION_CREATE=340.
- `OFFLINE_T3`: NET_DELTA=40, CREATE=60.
- `OFFLINE_T9_MIXED_REPLACE_CREATE`: Cancel=126 / KeepCreate=63 / ReplaceCreate=126 / CreateTotal=189.
- Guard chống lỗi = `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE` **fire live 9 lần** (cờ `TERMINAL_COMPLETENESS=UNACCOUNTED_ZERO`, `MIXED_CREATE_EFFECTIVE_STATE=TRUE`).
- `PLAN.csv` có cột `GROUP_DELTA_TOTAL` / `GROUP_MEMBER_COUNT` / `GROUP_NEED_QTY` xác nhận thiết kế đã bắt đầu theo dõi member completeness.

> Kết luận: dùng `3500+4500->0+8000` như **minh họa lớp lỗi**, dẫn bằng chứng số thật là T3/T4/T9 + 9 lần `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE`.

## 6. Bằng chứng từ lịch sử V2.6 (số liệu thật đã xác minh)

> Toàn bộ log chỉ trong 1 ngày **2026-08-19** (Excel serial 46253.x) - mẫu nhỏ, không ngoại suy độ ổn định dài hạn. Nguồn: gói `CF_OB_CODEX_CLAUDE_HANDOFF_20260819` (`extracted_package\...\current_extracted\sheets\`).

### 6.1. Đếm cấp RUN (sửa số nhớ nhầm của người dùng)

| Chỉ số | Người dùng nhớ | **Số đã xác minh** | Nguồn |
|---|---|---|---|
| EXECUTE run | ~45 | **14 EXEC run** (10 có dòng RUN_SUMMARY) | `RUN_HISTORY.csv` (14 EXEC + 27 CHECK = 41 RunID) |
| PASS hoàn chỉnh cấp RUN | ~9 | **2** (`PASS\|LOCKED_PLAN_PHASES_COMPLETE`: OB 1011832369, OB 1011839385) | `RUN_SUMMARY.csv` |
| HALTED/BLOCKED | ~23 | **4 SCHEDULER_HALTED_OR_BLOCKED** | `RUN_SUMMARY.csv` |

**Vì sao lệch:** các số ~45/~23/~9 là **đếm EVENT ExecState**, không phải cấp RUN:
- tổng dòng `BLOCKED_*` = **45** (khớp "45"),
- `BLOCKED_RELEASE_REQUIRED` = **22** (khớp "~23 HALTED"),
- `GROUP_QTY_BASELINE_PROVEN` = **9** và `QTY_READY` = **9** (khớp "9 PASS").
=> Người dùng nhớ nhầm 3 cột ExecState-event thành cấp RUN.

### 6.2. Kết quả 10 EXECUTE run có summary

| Result | Số run |
|---|---|
| PASS \| LOCKED_PLAN_PHASES_COMPLETE | 2 |
| SCHEDULER_HALTED_OR_BLOCKED | 4 |
| STATE_UNKNOWN | 2 |
| OB_REGRESSION_GATE_BLOCKED | 2 |
| (thêm 2 CHECK_FAIL ở channel CHECK) | 2 |

Channel Result: EXECUTE PASS=488 / STOP=26 / SENT=20 / SKIP=19. CHECK: DONE=39 / CHANGE_QTY=28 / CANCEL_TO_ZERO=24 / NEED_RELEASE=22 / CHANGE_BIN_BATCH=10 / CF_ONLY=4 / FAIL=2.

### 6.3. Bộ đếm navigation (sửa số nhớ nhầm 9/5/1)

| Bộ đếm | Điển hình | **Đỉnh thật** | Run mới nhất |
|---|---|---|---|
| MON_REFRESH_COUNT | 4 | **16** | 4 |
| DOCUMENT_OPEN_COUNT | 1 | **16** | 3 |
| DOCUMENT_RECOVERY_COUNT | - | **3** | 1 |
| MON_OPEN_COUNT | - | {0,1,2,4} | 1 |
| WHTASK_OPEN_COUNT | - | {0,1} | - |

- Đỉnh 16/16/3 rơi vào run STATE_UNKNOWN 002003 (P0007 BIW32023623AB). Run mới nhất (`latest_run_key_findings.md`): MON_OPEN=1 / MON_REFRESH=4 / DOC_OPEN=3 / DOC_RECOVERY=1.
- Số nhớ 9/5/1 CÓ tồn tại (DOC_OPEN=5 có thật, MON_REFRESH gần nhất 10) nhưng KHÔNG phải đỉnh. Đỉnh thật xấu hơn nhiều.

### 6.4. Phân bố StopCode (tổng hợp mọi file)

| StopCode | Số | Ghi chú |
|---|---|---|
| STOP_TARGET_BIN_NOT_FOUND_AFTER_POST_SUCCESS | 29 | chủ yếu mô tả kịch bản trong REGRESSION_AUDIT, không phải run live |
| STOP_NO_RETRY | 29 | như trên |
| STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE | 9 | **bằng chứng live thật (bug lõi §5)** |
| STOP_ACTUAL_BIN_INPUT_INVALID | 8 | |
| STOP_REGRESSION_GATE_FAIL | 6 | regression gate chặn live (sai thiết kế) |
| STOP_STATE_UNKNOWN | 4 | |
| STOP_INCOMPLETE_EXECUTION | 4 | |
| STOP_DATA_READY_UNPROVEN | 4 | |
| STOP_CREATE_GROUP_MEMBER_NOT_DATA_READY | 4 | |
| STOP_SOURCE_PLAN_INCOMPLETE | 2 | |
| STOP_CREATE_GROUP_RECORDS_UNPROVEN | 2 | |

3 tuple live của `STOP_PRDO_QTY_GROUP_MEMBER_INCOMPLETE`:
- **STD90007016** (OB 1011819247, run 053709): ACTIVE_MEMBERS=2/EXPECTED=3, GROUP 1250->1710, PLAN 790->1250.
- **BIW32023687AA**: ACTIVE=1/EXP=3, GROUP 580->480.
- **BIW70004029**: ACTIVE=1/EXP=2, GROUP 500->400.

### 6.5. Holder release + độ phủ live

| Chỉ số | Giá trị | Nguồn |
|---|---|---|
| Map `RELEASE_MON_*` (15 khóa) | **TẤT CẢ TRỐNG** ("blank stops the reader") | `CONFIG.csv` d87-101 |
| `PLAN_RELEASE.csv` | **RỖNG** (chỉ header) | `PLAN_RELEASE.csv` |
| Kế hoạch kẹt `BLOCKED_RELEASE_REQUIRED` | 22 | `RUN_HISTORY.csv` |
| `CANCEL_OTHER_WT` | TRUE (d28) | `CONFIG.csv` |
| `RELEASE_EXCLUDE_STATUS` | "A,C" (d29) - **mâu thuẫn: loại cả A là status holder cần cancel** | `CONFIG.csv` |
| `ALLOW_PARTIAL_WT_CANCEL` | FALSE (d30) | `CONFIG.csv` |
| `AFTER_CANCEL_ADD_QTY_TO_CACHE` | TRUE (d31) | `CONFIG.csv` |
| LIVE_PROVEN vs STATIC_PROVEN | **4** vs 597 (NOT_LIVE_PROVEN=481) | `REGRESSION_AUDIT.csv` |
| `CONFIG_VERSION` | `CF_OB_SAFE_BASELINE_V2` | `CONFIG.csv` |

> Build tag thật trong artifact = `V2.4-GROUP-EXECUTOR-WHTASK-REUSE` (1157 lần) + `V2.3-CORE-SCHEDULER-CONTRACT` (163). Chuỗi "V2.6"/"V2.5" KHÔNG tồn tại trong artifact.

## 7. KHÔNG được regression (các luồng V2.6 đã proven live)

Khi nâng lên V2.7 phải giữ nguyên các luồng/ngữ nghĩa sau (mất là hỏng):

**Luồng proven live phải giữ:**
- **P0005/P0006** đã chạy live (DELTA_ADD / CREATE+SAVE) - giữ path `STD90000408AA`.
- Fast-path tạo WT (LIVE_PASS), WT `5042711577`, DELTA_ADD (WT cũ 5042769777 giữ 60, PRDO->280, WT mới 5042833176 +220), CANCEL_TO_ZERO.
- Path replace P0005/P0006 (`STD90000408AA`): cancel cả 2 old WT, POST change, PRDO Save, create full 18000, verify WT mới - **không regression**.

**Map SAP LIVE_PASS/VERIFIED phải giữ (nguồn control duy nhất = `SAP_ACTION_MAP`):**
- `MON_WT_GRID_ID`, `MON_WT_ACTION_ID=N000000096`, `WT_REQUEST_GRID`, `WT_TAB_STK_CAN_BE_REMOVED`, `WT_STOCK_GRID`, `WT_STOCK_SET_QTY=UI_VSOLA`, `WT_STOCK_SELECT_ROW=SelectedRows`, `WT_ADOPT_QUANTITY=OK_OD_AVSTKDLV_TAKE_QTY`, `WT_CREATE_COMMAND_ID=OK_OI_DLV_CREATE_TO`, `WT_SAVE_COMMAND_ID=OK_OI_DLV_SAVE`, `WT_CREATE_SAVE` (composite Create ROI Save, KHÔNG one-click).
- Cột MON/POST đã dùng: TANUM, DOCNO, MATNR, CHARG, VLPLA, NLPLA, VSOLM, TOSTAT, AVL_ALT_QTY, TO_CAT.

**Ngữ nghĩa phải giữ:**
- INPUT: danh sách OB bên **TRÁI**, nút điều khiển/status bên **PHẢI** (không gộp thành form 1 OB ở trên).
- WH_ACTUAL: cột hệ thống (OB/Document/Item/WH/WT/Material/Sys Qty/Bin/Batch/Status) + cột user intent (Actual Qty/Bin/Batch).
- Chỉ `READY_FOR_CF` (WT mới được MON post-verify chính xác) là terminal success. `CreatedCount>0` KHÔNG đồng nghĩa READY_FOR_CF.
- **CF luôn do user** - code không gọi Confirm WT; `MON_CONFIRM_BG_COMMAND_ID` khóa disabled trong RUN.
- Chỉ row `LIVE_PASS`/`VERIFIED` đúng context/fingerprint được gửi write. PENDING/DISCOVERED chỉ khảo sát.
- Sau `STATE_UNKNOWN` (vd OB 1011832350, 1011837627): chỉ reload/reconcile, cấm retry mù.
- Nguồn stock dùng Open/Delta Qty (không dùng tổng PRDO). Match Product+Bin+Batch + Available>=RequiredDelta.
- Contract parser MIGO số nguyên: `1.000`/`1,000`/`1.000,00`/`1,000.000` -> 1000; phần lẻ thực (1,5; 1.500,25) bị chặn trước SAP ("SAP chỉ nhận số nguyên").
- Status A: A + Actual blank/0 = A_UNUSED (không tạo mới, RetainedQty 0); A + Actual>0 = A_RECREATE_REQUIRED (RetainedQty(A)=0, tạo WT SỐ MỚI). Status C = terminal/satisfied. **Không đổi WT A/C về Active.**
- **Không phá workbook:** giữ VBA project, buttons, shapes, defined names, sheet visibility, compatibility; không rebuild from scratch.

**Fixture bug bắt buộc giữ để chứng minh V2.7 đã diệt bug:**
- `5042889690` (OB 1011832412 / Doc 1012087422 / Item 10 / WH 6026342483 / Material BIW32023716AA, CHANGE_BIN_BATCH + Qty increase). Kỳ vọng V2.7: Cancel A 1 lần (không reopen re-prove A), Posting bin/batch, 1 PRDO Save, WHTask đọc Open Qty, chỉ Create allocation còn thiếu, WT cũ không bao giờ chặn Create.

## 8. Ràng buộc bắt buộc (constraints)

- Làm trên **bản COPY** của .xlsm hiện hành (V2.6); không tạo workbook mới rồi vứt cấu trúc/buttons/shapes/sheets/VBA.
- Offline development/replay **tuyệt đối không ghi SAP live**; mọi live write phải gate tường minh.
- **Không** thực thi SAP action trong vòng lặp `For Each P` (member loop chỉ READ/NORMALIZE/CLASSIFY/CALCULATE/PLAN).
- `MON_SESSION_ID <> PRDO_SESSION_ID` (bất biến cứng); không bao giờ tái dùng tab PRDO làm MON; role trùng -> STOP trước write (session role collision), cấm auto-reuse.
- `DOCUMENT_OPEN_COUNT=1` mỗi document mỗi lần chạy nút XỬ LÝ (`DOCUMENT_RECOVERY_COUNT=0`); không reopen OB/Document; dùng WHTask (`CWR_ReturnToWarehouseTask`) để quay lại WH/WT.
- Không WT đã chứng minh A/C được trở lại Active (cùng số WT).
- Bin exact (vd APFL-CC01) phải được chấp nhận khi SAP/stock chứng minh hợp lệ; không dùng `AREA_ALLOWED=P,ZF,HR` làm ràng buộc lỗi.
- **1 PRDO Save/document** cho grouped Qty update (`PRDO_SAVE_COUNT<=1`); tính mọi target trước, ghi top-to-bottom rồi Save một lần.
- Khóa plan/group đầy đủ (PLAN_LOCKED) trước SAP write đầu tiên; không re-plan sau khi phase write bắt đầu; không duplicate Create WT khi rerun (idempotent, reconcile Document+Item+Material+Bin+Batch).
- Phải pass toàn bộ **22 scenario replay S01-S22 offline** trước bất kỳ live SAP test nào; replay dùng raw snapshot/event, không dùng cột Decision cũ làm đáp án (chống pass-by-cheating).
- Thứ tự phase cố định: Cancel -> cập nhật WT-state một lần -> Posting/Bin/Batch -> grouped Qty -> 1 PRDO Save/doc -> WHTask/OpenQty -> Create WT -> WHTask verify -> CF user handoff; không nhảy lùi phase.
- CF là user handoff mặc định; auto-CF chỉ bật khi được phép tường minh + đã live-proven; code không gọi Confirm WT.
- Chỉ **một executor** sở hữu SAP writes; CHECK/MAP/AUDIT chỉ đọc (`SAP_WRITE=0`).
- Không retry mù sau unknown write; chỉ read/reconcile. Unknown chỉ terminal khi Adopt/Create+Save đã gửi mà chưa chứng minh; stop trước Adopt/Create với `SAP_WRITE=0` là BLOCKED có thể retry.
- Không dùng screenshot/mouse/SendKeys/tọa độ/quét toàn bộ SAP tree-grid; chỉ dùng technical map.
- Live test theo ladder (Section 25): Live1 CF-only -> Live2 Cancel -> Live3 Qty delta (580->600) -> Live4 Change Batch+Increase (dễ lộ bug stale-context nhất) -> Live5 mixed multi-P; sau đó mới chạy nhiều OB.
- Production baseline: file V2.6 (không dùng V1.5/legacy hard-code hay flow POST/Cancel cũ nếu chưa có row `LIVE_PASS`/`VERIFIED` tương ứng).
- **Absolute deny:** không MIGO, không delete WT, không coordinate/SendKeys fallback trong RUN, không auto CF; không `On Error Resume Next` che lỗi nghiệp vụ; không loop/retry vô bound; không sleep làm proof chính; không hard-code WT number/row 5/row 7; không đổi `GROUP_TARGET=GROUP_CURRENT` thành luôn create target.

## 9. Kiến trúc đích V2.7 (engine + 5 sheet + UI + History V2)

### 9.1. Module engine (tách trách nhiệm)
4 public entry point: `CF_LoadWH` / `CF_Check` / `CF_Execute` / `CF_ResetUI`. Một call path production: UI -> Planner -> Freeze -> Broker -> Scheduler -> Registry -> Adapter -> PostVerify -> History -> UI. Không clone executor thứ ba, không 2 write path cho cùng ActionType.

| Module | Trách nhiệm |
|---|---|
| `modCFGroupAllocator` (+ `PLAN.tblMemberPlan`/`tblGroupPlan`) | **Member Allocation Engine** - tách Contract A (PRDO total) và Contract B (member alloc); `RequiredCreateQty = max(TargetQty - EffectiveQtyAfterPlannedCancelReplace, 0)` |
| `modCFExecutor` + Group Scheduler | Hierarchy Run > OB > Document > Cancel/Release/PostChange/PRDO/Create/FinalVerify Group; 9 phase cố định |
| `modSapSessionBroker` / `CWR_GetMONSession` | **Single MON Broker** - 1 MON reader/EXEC run; roles ROLE_MON_READER + ROLE_EXECUTOR; chỉ broker được CreateSession/gửi `/N/SCWM/MON` |
| `modCFReleasePlanner` / `modReleaseHolderReader` | **Holder Release** - query exact Material+Bin+Batch, loại A/C + WT plan + WT mới, deterministic WHOLE-WT selector, coverage/gate, cache += ReleasedQty CHỈ sau status A verified; điều khiển bởi `RELEASE_HOLDER_MODE` |
| `PLAN.tblActionPlan` | **Action Ledger** - 1 row = 1 SAP action; ActionType chuẩn hóa CANCEL_OWN_WT/CANCEL_HOLDER_WT/POST_CHANGE/PRDO_QTY_WRITE/PRDO_BIN_WRITE/PRDO_BATCH_WRITE/PRDO_SAVE/CREATE_WT/MON_READBACK/FINAL_VERIFY/USER_CF_HANDOFF |
| `modCFRecovery` | **Resume/Idempotency** - state machine PLANNED->SENT->VERIFIED->UNKNOWN->SKIPPED_ALREADY_VERIFIED; ActionID tuple ổn định; UNKNOWN -> RecoveryRequired + hard-stop |
| Document Executor + Member Accounting | states ACCOUNTED_CF_ONLY/CANCEL_VERIFIED/PRDO_READBACK_ONLY/PRDO_WRITE_REQUIRED/CREATE_PENDING/CREATE_VERIFIED/BLOCKED/TERMINAL_C |
| `CanExecuteAction(...)` | Central Permission Gate - gộp 2 hệ safety; flag legacy = LEGACY_NOT_AUTHORITATIVE |
| `SourceStockRepository` + `DATA_CORE.tblSourceStock` | SAP capture -> normalize in memory -> tblSourceStock (raw ở `tblSourceStockRaw`); POSTING loại bỏ |
| `modCFHistory` + `RUN_HISTORY.tblRunEvents` | History Event Ledger 65 cột, batch buffer 20-100 event |
| `modCFNumberParser` | Parser canonical (thousands '.', decimal ','), Double/Decimal-safe |

ActionID tuple (KHÔNG dùng row number): `CANCEL|OB|DOC|ITEM|WT` ; `RELEASE|TargetSourceKey|HolderWT` ; `PRDO_QTY|OB|DOC|ITEM|TargetQty` ; `CREATE|OB|DOC|ITEM|Material|Bin|Batch|Qty`.

### 9.2. Đúng 5 sheet logic

| Sheet | Visibility | Gộp từ | Tables chính |
|---|---|---|---|
| **INPUT** | Visible | (giữ nguyên) | UI Control Center: OB queue trái, Actual input, status, preview, 4 nút. KHÔNG làm cache engine dài hạn |
| **DATA_CORE** | VeryHidden | ORDER_QUERY + POSTING_RAW + WH_ACTUAL(backend) + POSTING + OB_QUEUE + OB_WH_CACHE + POSTING_SUMMARY + POSTING_CACHE | tblOBQueue, tblWHSnapshot, tblActualSnapshot, tblSourceStock, tblSourceStockRaw, tblCaptureSummary, tblRuntimeState |
| **PLAN** | VeryHidden | PLAN + PLAN_SOURCE_LEGS + PLAN_RELEASE | tblMemberPlan, tblGroupPlan, tblSourceLeg, tblReleaseCandidate, tblActionPlan |
| **SYSTEM** | VeryHidden | CONFIG + SAP_MAP + SAP_ACTION_MAP + SAP_PROBE_LOG | tblConfig, tblSapMap, tblSapActionMap, tblSapProbe |
| **RUN_HISTORY** | VeryHidden (Debug/Admin mới unhide) | RUN_LOG + RUN_HISTORY + RUN_SUMMARY + REGRESSION_AUDIT | tblRunEvents (65 cột); summary tính bằng VBA memory; regression = EventType REGRESSION_PASS/FAIL |
| ~~POSTING~~ | **XÓA HẲN** | - | thay bằng SourceStockRepository / tblSourceStock |

### 9.3. UI 4 nút (REDESIGN override 3-nút)
- **[1] TẢI WH** - đọc list nhiều OB, reuse 1 MON session, load WH snapshot, sync UI. KHÔNG đọc POST nếu chưa cần, KHÔNG lập final plan, KHÔNG write SAP. Sau đó: CHECK enabled, XỬ LÝ disabled, PlanState=NOT_CHECKED.
- **[2] CHECK** - **READ-ONLY TUYỆT ĐỐI**: validate input, freeze InputSnapshot, đọc source stock, build member/group/source-leg/release/action plan, blocker gates, tính fingerprints, render preview. Có `AssertCheckIsReadOnly()`; mọi routine write đòi `ExecutionContext.Mode=EXECUTE`, hard-stop trước `.Press`/`.ModifyCell`/`.Text=`/`.Key=`/Save/Cancel/Create nếu Mode=CHECK. Kết quả: CHECK_PASS / CHECK_PASS_WITH_RELEASE_PLAN / CHECK_BLOCKED.
- **[3] XỬ LÝ** - entry point DUY NHẤT ghi SAP. Chỉ enabled khi CheckID+PlanID tồn tại, PlanState=LOCKED, InputFingerprint khớp, WH snapshot không stale, MapFingerprint không đổi, không RecoveryRequired. KHÔNG rebuild plan âm thầm, KHÔNG tự CHECK lại; khác fingerprint -> `STOP_PLAN_STALE_RECHECK_REQUIRED`. Chạy theo `tblActionPlan` đúng thứ tự.
- **[4] LÀM MỚI** - chỉ clear UI/current-session scope. KHÔNG xóa RUN_HISTORY/UNKNOWN evidence/action checkpoint/recovery state/SYSTEM. Có unresolved write -> hiện "RECOVERY_REQUIRED - không thể làm mới engine state trước khi reconcile".

**Button state machine:** Initial (TẢI WH+LÀM MỚI enabled) -> sau TẢI WH PASS (CHECK enabled) -> user edit Actual (plan STALE, XỬ LÝ disabled) -> CHECK PASS (XỬ LÝ enabled) / CHECK BLOCKED (XỬ LÝ disabled) -> trong XỬ LÝ (cả 4 nút disabled) -> sau XỬ LÝ complete (CHECK có thể enabled, XỬ LÝ disabled tới CHECK mới).

**Worklist (REDESIGN §13):** OB / Material / WT / Sys Qty / Sys Bin / Sys Batch / Actual Qty / Area / Actual Bin / Actual Batch / Decision / Release Need / Plan State / Execute State / Message. Chỉ 4 cột editable: Actual Qty, Area, Actual Bin, Actual Batch. Mỗi visible row có hidden SourceKey sync theo SourceKey (KHÔNG theo row number); duplicate/missing -> STOP pre-write. C ẩn khỏi worklist nhưng backend/history vẫn giữ (không `EntireRow.Hidden` bừa).

### 9.4. History V2 (65 cột)
- MỘT durable event ledger `tblRunEvents`. Một event = một transition có nghĩa, KHÔNG log mỗi polling.
- Hệ ID phân cấp: `LoadRunID -> CheckID -> PlanID (locked) -> ExecuteID -> ActionID (+ParentActionID/ActionSeq)`.
- EventType chuẩn: LOAD_START/LOAD_ROW/LOAD_COMPLETE/CHECK_START/INPUT_SNAPSHOT/MON_READ/SOURCE_STOCK_READ/HOLDER_READ/PLAN_MEMBER/PLAN_GROUP/PLAN_RELEASE/PLAN_ACTION/CHECK_PASS/CHECK_BLOCKED/EXEC_START/ACTION_PRECHECK/ACTION_SENT/ACTION_READBACK/ACTION_VERIFIED/ACTION_UNKNOWN/ACTION_SKIPPED_VERIFIED/EXEC_BLOCKED/EXEC_COMPLETE/USER_CF_HANDOFF/RECOVERY_REQUIRED/RECOVERY_RESOLVED/REGRESSION_PASS/REGRESSION_FAIL.
- StopCode unique per stop (kèm OB/Document/Item/WT/ActionID/Expected/Actual/RequiredNextStep), KHÔNG generic "Scheduler halted".
- Append batch buffer 20-100 event, flush ở action boundary / ngay khi UNKNOWN-ERROR / cuối phase.
- 65 cột: EventSeq, Timestamp, WorkbookVersion, EventType, RunType, LoadRunID, CheckID, PlanID, ExecuteID, ActionID, ParentActionID, ActionSeq, Stage, Phase, OB, Document, Item, WH, WT, Material, QtyGroupKey, SourceKey, ActionType, ActionState, BeforeQty, TargetQty, AfterQty, BeforeBin, TargetBin, AfterBin, BeforeBatch, TargetBatch, AfterBatch, RequiredCreateQty, RequiredReleaseQty, AvailableBefore, ReleasedQty, AvailableAfter, SAPConnection, SAPSession, SAPTransaction, SAPWindowTitle, SAPControlKey, SAPTechnicalID, SAPStatusType, SAPStatusText, EvidenceType, EvidenceValue, EvidenceFingerprint, PlanFingerprint, InputFingerprint, MapFingerprint, ConfigFingerprint, DurationMs, Result, StopCode, RetryAllowed, ResumeState, ResumeFromActionID, ErrorNumber, ErrorSource, ErrorDescription, Message, UserName, MachineName.

## 10. Migration 20 -> 5 sheet + gỡ POSTING an toàn (áp trên workbook V2.6)

### 10.1. Stage 1 - Snapshot + audit (làm trên bản COPY, KHÔNG rebuild trắng)
Copy V2.6, hash file, export/list VBA modules, snapshot sheet visibility + defined names (`CFX_UseNewEngine=FALSE`) + shapes/buttons, parse RUN_SUMMARY/RUN_LOG/RUN_HISTORY mới nhất.

### 10.2. Migration theo Phase A/B/C/D (KHÔNG delete 15 sheet ngay)
- **Phase A** - Clone + build adapters + tạo 5 target sheet, vẫn giữ legacy.
- **Phase B** - Dual-read verify: output V2.7 repositories, compare legacy vs new để data/plan parity.
- **Phase C** - Chuyển production engine sang new repositories + search zero legacy runtime reference.
- **Phase D** - Mới delete/decommission legacy sheet, còn đúng 5 sheet.

### 10.3. Giao thức gỡ POSTING an toàn (REDESIGN §32 - đúng thứ tự, không chỉ hide)
1. Search toàn source VBA reference tới `"POSTING"` / `Worksheets("POSTING")`.
2. Lập dependency list.
3. Thay MỌI reader sang `SourceStockRepository` in-memory hoặc `DATA_CORE.tblSourceStock`.
4. Thay MỌI writer.
5. Compile toàn bộ VBA.
6. Replay regression.
7. Xác nhận **ZERO runtime reference**.
8. MỚI delete sheet POSTING - **cấm catch-error tự recreate POSTING**.

### 10.4. Deprecation map đầy đủ
`ORDER_QUERY`->tblOBQueue; `POSTING_RAW`->tblSourceStockRaw; `WH_ACTUAL`->tblWHSnapshot+tblActualSnapshot; `RUN_LOG`->tblRunEvents; `SAP_MAP`->tblSapMap; `CONFIG`->tblConfig; `OB_QUEUE`->tblOBQueue; `OB_WH_CACHE`->tblWHSnapshot; `POSTING_SUMMARY`->tblCaptureSummary; `POSTING_CACHE`->tblSourceStock; `PLAN_SOURCE_LEGS`->tblSourceLeg; `PLAN_RELEASE`->tblReleaseCandidate; `PLAN`->tblMemberPlan+tblGroupPlan+tblActionPlan; `SAP_ACTION_MAP`->tblSapActionMap; `SAP_PROBE_LOG`->tblSapProbe; `RUN_SUMMARY`->RUN_HISTORY summary; `REGRESSION_AUDIT`->RUN_HISTORY regression events.

### 10.5. Quy tắc kèm theo
- KHÔNG preallocate 500 blank row như POSTING; chỉ ghi actual rows; load snapshot vào Dictionary/Collection once, grouping in-memory, block read/write array.
- Compatibility: macro ngoài còn gọi legacy sheet -> tạo adapter VBA (vd `GetSourceStockRows()` thay `Worksheets("POSTING")`), KHÔNG giữ 15 legacy sheet.
- Plan contract: nâng `CHECK_LOCKED_V2` -> `CHECK_LOCKED_V3_MEMBER_ALLOCATION`; plan cũ V2 mở bằng V2.7 KHÔNG execute stale locked plan, phải CHECK lại.
- Không phá lịch sử: giữ toàn bộ RUN_LOG/RUN_HISTORY/RUN_SUMMARY/REGRESSION_AUDIT cũ khi gộp; ghi event `MIGRATION_V26_TO_V27` (old=20, new=5, posting removed=true, history unified=true). `ArchiveHistoryBefore(Date)` thủ công - không auto, không gọi khi XỬ LÝ, không xóa run UNKNOWN chưa resolved.
- Version labels: UI -> V2.7, output `260820_CF_OB_V2.7.xlsm`, version text "V2.7 - 4BUTTON/5SHEET/ALLOCATION LEDGER/SINGLE MON/RELEASE GATED/HISTORY V2".
- **BLOCKER:** nếu chưa chứng minh zero dependency thì KHÔNG được xóa sheet sớm và KHÔNG được gọi build là final V2.7 (technical blocker phải document rõ).

## 11. Regression contracts V2.7 (danh sách test bắt buộc)

### 11.1. T01-T37 (contract chính)
| ID | Nội dung | Kỳ vọng |
|---|---|---|
| T01 | CF only | qty/bin/batch unchanged -> no write, READY_FOR_CF |
| T02 | pure increase single WT | 580->600, source đủ -> create delta 20 |
| T03 | pure decrease | 580->500 -> total final 500 |
| T04 | cancel + increase net positive | old 580, cancel 20, create 60, final 620 - create PHẢI=60, không phải net 40 |
| T05 | cancel + increase net negative | old 580, cancel 420, create 340, final 500 - create=340 |
| **T06** | **CRITICAL net-zero redistribution** | WT1 3500->0, WT2 4500->8000, group 8000=8000 -> cancel 3500, retain 4500, PRDO write=0, RequiredCreateQty=3500, final 8000 PASS. FORBIDDEN: skip create / STOP_PRDO_DOCUMENT_MEMBER_NOT_INCLUDED |
| T07 | replace net-zero | 5050->0, 12950->18000 + change bin/batch, 18000->18000 -> cancel cả 2, POST change, create full 18000, READY_FOR_CF (giữ path STD90000408AA) |
| T08 | 2 retained increase cùng QTY_GROUP | create qty = tổng member deltas, mỗi allocation attributable đúng member/source |
| T09 | 3-member mixed | 1 cancel-zero + 1 retain no-change + 1 retain+delta -> không false-stop |
| T10 | 1 blocked member + independent group same doc | blocked no write, independent vẫn chạy nếu PRDO row identity isolated |
| T11 | multiple docs 1 OB | 1 doc blocked, doc khác vẫn proceed |
| T12 | multiple OB list | OB1 pass, OB2 blocked, OB3 pass -> continue tới OB3 trừ khi OB2 tạo unknown |
| T13 | status A resume | cancel đã VERIFIED -> không resend, create remaining nếu cần |
| T14 | status C | ẩn UI, excluded planner, retained audit |
| T15 | change bin only | không write batch thừa |
| T16 | change batch only | không change bin thừa |
| T17 | change bin + batch | giữ path P0006 proven |
| T18 | source free stock đủ | không holder reader/cancel |
| T19 | holder release single WT enough | chọn smallest single WT>=need, 1 cancel, 1 readback, cache chỉ sau A |
| T20 | holder release multiple WT | no single đủ -> deterministic cumulative candidates |
| T21 | holder A/C excluded | không bao giờ auto cancel WT A hoặc C |
| T22 | same OB unplanned holder | không auto cancel, mark REVIEW_SAME_OB_UNPLANNED_HOLDER |
| T23 | holder cancel unknown | no resend, hard-stop whole run |
| T24 | MON already open | không tạo MON tab mới |
| T25 | no MON open | tạo đúng 1 MON tab |
| T26 | MON busy | bounded wait, không duplicate session |
| T27 | 1 OB six rows (current case) | MON_REFRESH<=4, DOC_OPEN<=1 (<=2 nếu recovery), WHTASK_OPEN<=1 |
| T28 | rerun after POST verified | no duplicate posting change |
| T29 | rerun after WT Create+Save verified | no duplicate WT |
| T30 | unknown after Create | no create retry |
| T31 | numeric parser | 1.200->1200, 18.000->18000, 2.442->2442, 1,2->1.2; không bao giờ 1200->1.2 |
| T32 | blank/zero input | 0=intentional cancel-to-zero, blank required=user error, không coi blank là zero |
| T33 | duplicate SourceKey | pre-write STOP |
| T34 | plan fingerprint changed after CHECK | no execute stale plan |
| T35 | C rows hidden UI nhưng engine history vẫn chứa | |
| T36 | button check fails before writes | SAP_WRITE=0 |
| T37 | partial pass with blockers | phân biệt PASS / PARTIAL_PASS_WITH_BLOCKERS / HARD_STOP_UNKNOWN, không generic |

### 11.2. OFFLINE hard contracts (bổ sung T10-T18 vì T1-T9 offline cũ PASS nhưng chưa bắt được live bug)
- **OFFLINE_T10_NET_ZERO_MEMBER_REALLOCATION** - hard: GROUP_CURRENT=8000 / GROUP_TARGET=8000 / CANCEL=3500 / RETAINED=4500 / TARGET_MEMBER=8000 / PRDO_WRITE=0 / CREATE_REQUIRED=3500 / FINAL=8000. **CREATE_REQUIRED=0 -> FAIL.**
### 11.3. Ma trận REDESIGN
- **Group matrix:** mandatory `3500+4500->0+8000`; `500+500->0+1000`; `300+700->500+500`; `1000->1200`; `1000->800`; old bin->new bin; old batch->new batch; 3-member reallocation; 2 materials 1 doc; multiple docs 1 OB; multiple OB.
- **Holder release matrix (16 case):** free đủ->no release; thiếu 100 holder 100->1 candidate; thiếu 100 holder 40+60->2 selected; holder A->excluded; holder C->excluded; wrong batch/bin/material->excluded; same target WT->excluded; candidate đổi status sau CHECK->EXEC stop stale; cancel readback A->credit; cancel readback unknown->UNKNOWN/no credit; insufficient coverage->CHECK blocked; PLAN_ONLY->no cancel; PILOT->one group only; LIVE_LOCKED->chỉ sau proof gate.
- **Session matrix (10):** MON already open / MON wrong node / MON lost after Cancel / POST screen open / document already open / multiple SAP sessions / busy session / popup after write / session disappears / Broker recovery once. Assert: no duplicate MON tab, no reopen per P000x, all readers dùng broker.
- **History matrix:** một full run phải chứng minh LOAD_START/COMPLETE, CHECK_START, InputSnapshot, PlanMember, PlanGroup, ReleasePlan (nếu có), CheckPass, ExecStart, ActionPrecheck, ActionSent, ActionReadback, ActionVerified, FinalVerify, UserCFHandoff; nếu unknown có ActionUnknown+RecoveryRequired.
- **Resume test:** A) after Cancel sent -> no resend tới khi readback; B) after POST Save -> no second Save trước reconcile; C) after Create Save -> tìm created WT trước mọi retry; D) after final verify -> không rerun business write.

### 11.4. Hard regression assertion (REDESIGN §17)
- KHÔNG PASS group nếu `sum(FinalVerifiedMemberQty) <> GroupTargetTotal`.
- KHÔNG coi member complete nếu `RequiredCreateQty>0` mà chưa có verified created/reacquired quantity.

### 11.5. Deliverable gate (KHÔNG được gọi FINAL nếu chỉ...)
KHÔNG final nếu chỉ: đổi 3->4 nút, chỉ hide POSTING, chỉ merge history visually, vẫn dùng old group boolean, CHECK vẫn write SAP, XỬ LÝ tự rebuild plan, release holder không readback, MON vẫn reopen per member. Cũng KHÔNG final nếu chỉ refactor sheet mà regression nghiệp vụ fail (cần cả data parity + plan parity + live-proven branch giữ nguyên + bug mixed allocation fixed).
Deliverables đầy đủ: `260820_CF_OB_V2.7.xlsm` + ZIP + `V2.7_CHANGELOG.md` + `V2.7_MIGRATION_AUDIT.md` + `V2.7_REGRESSION_REPORT.md` (T01-T37 PASS/FAIL/NOT_PROVEN, no fake PASS cho LIVE-only action) + before/after audit. Báo cáo cuối phải trả lời 10 câu evidence + ledger trả lời 19 câu.

## 12. RESUME TỪ ĐÂY - checklist bước kế tiếp (xuất phát từ base V2.6)

> KHÔNG có bước nào là resume/chạy harness V2.5. Mọi bước áp trên **bản COPY của `260819_CF_OB_V2.6_LICH_SU.xlsm`**.

### 12.0. ROADMAP THỰC THI (thứ tự user chốt 2026-08-20) - trạng thái sống
| # | Bước | Trạng thái | Bằng chứng |
|---|---|---|---|
| 1 | Fix Allocation Engine (3500+4500→0+8000) | ✅ **DONE** | `cp2_alloc_regression.log` G01 totCreate=3500; CFAlloc_SelfTest=PASS |
| 2 | Regression Qty/group cũ (không phá replace/cancel/create) | ✅ **DONE** | 3 tầng PASS: allocator 11/11 + CF_FastModuleRegression PASS + CF_Phase3_RunUnitTests 11/11 SAP_Write=0 |
| 3 | Tách CHECK / XỬ LÝ | ✅ **DONE (offline)** | Seam 2 bịt: FAST `CWR_RunFastToCFCore` nay gọi `CWR_AssertQueuePlanLock` TRƯỚC create (lock@1300 < create@1318). `cp3_verify.log`: compile PASS, 3 tầng regression vẫn xanh, ASSERTQUEUEPLANLOCK_CALLS=7. `cp3_planlock_test.log`: plan khóa thật nhất quán + tamper phát hiện được. |
| 4 | Single MON Broker + Group Scheduler | ✅ **DONE (offline)** | Đã có sẵn trong V2.6 + hardening: `CWR_GetMONSession` (dedicated reader, ~18 call sites) là broker duy nhất; budget `CWR_ReserveMON*` STOP khi vượt; Group Scheduler phase-ordered đã tích hợp allocator CP2. Legacy `CWR_ReuseVerifiedPRDOTabForMON` (0 caller) nay fail-closed `STOP_LEGACY_MON_REUSE_DISABLED`. `cp4_harden_mon.log` + re-verify 3 tầng PASS. |
| 5 | Holder Release PLAN_ONLY | ✅ **DONE (offline) + LIVE blocked** | `modReleaseHolderReader` đầy đủ: query exact Material+Bin+Batch, whole-WT selector, coverage, idempotency, readback. Eligibility loại đúng A/C (gate `statusText=""`). READ-ONLY (chỉ `.Press` nút Execute filter, không write). Executor preflight `PLAN_ONLY=TRUE` mặc định. Compile PASS. **BLOCKER:** điền 15 map `RELEASE_MON_*` cần SAP GUI để discover ID → fail-closed `STOP_RELEASE_LAYOUT_UNPROVEN` (an toàn). |
| 6 | History V2 (65 cột + Action Ledger) | ✅ **DONE (core, offline)** | `CF_HistoryEnsureLedger`/`CF_HistoryAppend` (RUN_HISTORY 27 cột, 3304 dòng lịch sử thật). RunID phân cấp CHỨNG MINH: `EXECUTE|...\|CHECK-...` mang parent CheckID → chuỗi CheckID→ExecuteID→ActionID. Có đủ evidence lõi: before/requested/readback + executionState (SENT/VERIFIED/UNKNOWN) + actionKey + fingerprint. Mở rộng 27→65 cột = enhancement (lõi đã đủ). |
| 7 | Migration 20 → 5 sheet + gỡ POSTING | 🟡 **PROTOCOL SẴN SÀNG - phần phá hủy GATED** | 20 sheet đã liệt kê + map đích 5 sheet (§13 dưới). POSTING có ref THẬT ở 3 module runtime (ModuleSAP 2, modCF_Posting nhiều, modSAPReadOnlyProbe 2) + const/log label. **CỐ Ý CHƯA xóa POSTING**: `modCF_Posting` còn ghi (`wsRaw.Name="POSTING"`, MIGO 311) → xóa ngay sẽ vỡ luồng posting. Cần xây `SourceStockRepository`/`tblSourceStock` (phụ thuộc data POSTING live) TRƯỚC. Phần đảo ngược được (VeryHidden hoá) an toàn; phần merge+delete = BLOCKER cho tới khi repository xong + zero runtime ref. |
| 8 | UI final 4 nút | 🟡 **Logic DONE - shape vật lý còn lại** | Entry point mới `CF_ResetUI`/`CF_ResetUICore` (LÀM MỚI) built + `CFResetUI_SelfTest=PASS` (invalidate plan lock → ép re-CHECK, giữ lịch sử, SAP_WRITE=0). Map 4 nút: 1.TẢI WH=`CF_LoadWH`, 2.CHECK=`CF_CheckAndLoadPosting` (read-only), 3.XỬ LÝ=`CF_ExecutePlanToCFHandoff`, 4.LÀM MỚI=`CF_ResetUI` (mới). Ẩn 3 nút chẩn đoán (KIEM TRA SAP/MO RUN LOG/MAP POST). Việc còn lại = sắp xếp shape vật lý trên INPUT (cosmetic, làm cuối). |
| 9 | Package deliverables V2.7 | ✅ **DONE (offline)** | `work\OUTPUT\260820_CF_OB_V2.7.xlsm` (SHA `3ba75a86…c590b0`, 2,595,812 B) + `260820_CF_OB_V2.7_PACKAGE.zip` (SHA `6d88ce3c…6404c`, 10 entries) + 9 docs `DELIVERABLES\V2.7_{BASELINE_AUDIT,SOURCE_CONFLICT_LEDGER,CHANGELOG,MIGRATION_AUDIT,REGRESSION_CATALOG,REGRESSION_REPORT,UI_DECISIONS,ITERATION_LOG,FINAL_AUDIT}.md`. FINAL_AUDIT re-verify THẬT trên file OUTPUT (`v27_final_audit.log`): ZIP 55 entries + vbaProject.bin, XML_PARSE_OK, VBA=67, `CFAlloc_SelfTest=PASS` (compile proof), `CFResetUI_SelfTest=PASS`, PLAN AJ2=LOCKED/AT2=CHECK_LOCKED_V2. 10 câu evidence trả lời trong FINAL_AUDIT §7. Status = `OFFLINE_PASS_LIVE_BLOCKED_EXTERNAL`. (`LIVE_TEST_REPORT` bỏ đúng vì chưa có SAP GUI - disclosed §6.) |

> **Nguyên tắc thứ tự (user):** UI + tối ưu sheet làm GẦN CUỐI. Nếu đổi cấu trúc workbook trước khi Allocation Engine ổn thì sau này không phân biệt được lỗi do logic mới hay do migration sheet. COPY làm việc: `work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm`.

> **Phát hiện kiến trúc (2026-08-20, CP3):** V2.6 ĐÃ tách CHECK/XỬ LÝ phần lớn. `CF_Check` (modCF_CheckPlanner.bas:36) read-only tuyệt đối (grep: KHÔNG có `.Press`/`.Text=`/`FindById`) và ĐÃ sinh sẵn PlanID + fingerprint dòng + hash-lock toàn plan (`CP_PlanContractHash`→`PLAN!AN2`, `AJ2=LOCKED`, `AT2=CHECK_LOCKED_V2`). Luồng QUEUE `CWR_ExecuteOBQueueToCFHandoffCore` verify lock qua `CWR_AssertQueuePlanLock` (dòng 12497). **Lỗ hổng thật = luồng FAST** `CWR_RunFastToCFCore` ghi SAP mà KHÔNG verify lock. Fix CP3 = thêm ĐÚNG guard đó vào FAST (additive, tái dùng code đã có, KHÔNG dựng hạ tầng mới). Còn lại (nợ kỹ thuật, xử lý trước khi bật live): hợp nhất 2 cổng phép (WriteWhitelist vs AssertCFHandoff) + engine CFX chưa verify hash CHECK - ghi trong §12 Bước sau.

> Ghi chú: đánh số "Bước 1-8" ở các mục 12.x BÊN DƯỚI là bản kế hoạch CP cũ (theo checkpoint gốc), KHÁC thứ tự roadmap ở 12.0. Bảng 12.0 là thứ tự thực thi chính thức.

**Bước 0 (chuẩn bị) - BLOCKER xác nhận trước:**
- [ ] Xác nhận file base đúng (`_LICH_SU` là bản mới nhất? có tồn tại `260819_CF_OB_V2.6.xlsm` không?).
- [ ] Copy base -> workbook làm việc; hash + snapshot VBA modules + sheet visibility + defined names + shapes/buttons (Stage 1 §10.1).
- [ ] Thống nhất danh xưng: UI "V2.6" vs build tag `V2.4-GROUP-EXECUTOR`.

**Bước 1 (BLOCKER #1 - bug lõi §5) - ✅ DONE 2026-08-20 (đã xác minh offline):**
- [x] Xây Member Allocation Engine `modCFGroupAllocator` (COPY `work\WORKING\260820_CF_OB_V2.7_CP2_ALLOC.xlsm`): tách **Contract A (PRDO total)** khỏi **Contract B (member alloc)**. 9 hàm `CFAlloc_*` (ComputeRequiredCreate / IsContractAPass / IsContractBPass / IsGroupComplete / PRDOAggregateWriteRequired / AllocationWriteRequired / MemberStatusText / GroupStatusText / SelfTest).
- [x] Cài `RequiredCreateQty = max(TargetQty - EffectiveQtyAfterPlannedCancelReplace, 0)`; Effective theo ExecMode (CANCEL/REPLACE→0, KEEP→SysQty). Bỏ driver `GroupTarget-GroupCurrent`.
- [x] Patch `modCodexWTRework` (CWR_SchedulerVerifyPRDOGroup / CWR_SchedulerVerifyQtySources): gate `QtyReady` qua `CFAlloc_IsContractBPass`; đổi token sai `PRDO_QTY_ALREADY_TARGET|NO_QTY_WRITE` → `MEMBER_ALLOCATION_COMPLETE|REQUIRED_CREATE=0` / `CONTRACT_B_PASS`; thêm log `QTY_GROUP_ALLOCATION_STATUS`.
- [x] **PASS xác minh** (`work\REPORTS\cp2_alloc_regression.log`): case lõi `G01_NETZERO_3500_4500` → `CONTRACT_A=True | complete=False | allocWrite=True | totCreate=3500` (V2.6 cũ sai: DELTA=0→complete→mất CREATE 3500). Fixture live thật = STD20000556 (OB 1011846094, Doc 1012100366, WT 5042936844 3500→0 / WT 5042936845 4500→8000).
- [x] **Regression Qty không phá** (11/11 PASS): R01 no-op(0) / R02 delta+20 / R03 delta+400 / R04 decrease-replace(400) / R05 cancel-zero(0) / R06 bin-change-replace(580) / R07 replace-increase(640); G02 sau-create→complete/allocWrite=False / G03 multi-increase(120) / G04 mixed-realloc(700).
- [x] **Integrity COPY:** SHEETS=20 (bằng gốc), VBA_COMPONENTS=66 (65+allocator), `CFAlloc_SelfTest=PASS` (chứng minh full-project compile sạch, no reserved-word/syntax err). Harness đóng KHÔNG lưu regression module.
- [ ] **Còn lại (khi có SAP GUI):** T03/T04/T05/T08/T09 live-ladder + fixture 5042889690 stale-state - `BLOCKED_EXTERNAL` (offline gate đã đủ).

**Bước 2 (BLOCKER #2 - Holder Release §4.2):**
- [ ] Điền 15 map `RELEASE_MON_*` (đang trống) - reader mới đọc được holder.
- [ ] Sửa mâu thuẫn `RELEASE_EXCLUDE_STATUS` (đang loại cả A là status holder cần cancel).
- [ ] Xây `modCFReleasePlanner`/`modReleaseHolderReader`: query exact + deterministic WHOLE-WT selector + coverage + idempotency + readback; `RELEASE_HOLDER_MODE=PLAN_ONLY` mặc định (fail-closed).
- [ ] Chứng minh Holder release matrix 16 case + OFFLINE_T14/T15 offline.

**Bước 3 (Single MON Broker §4.3 - diệt thrash):**
- [ ] Xây `modSapSessionBroker`; bỏ `CWR_ReuseVerifiedPRDOTabForMON`.
- [ ] Chứng minh OFFLINE_T13 + Session matrix 10 case; T27 budget (MON_REFRESH<=4, DOC_OPEN<=1).

**Bước 4 (Scheduler + Action Ledger + Recovery §4.4-4.7):**
- [ ] `tblActionPlan` làm nguồn duy nhất; 9 phase cố định; ActionID tuple ổn định.
- [ ] `modCFRecovery`: PLANNED->SENT->VERIFIED->UNKNOWN; UNKNOWN -> RecoveryRequired + hard-stop.
- [ ] Chứng minh Resume test A/B/C/D + T13/T28/T29/T30.

**Bước 5 (Data repo + History V2 §4.9-4.10):**
- [ ] `SourceStockRepository` in-memory + `tblSourceStock`.
- [ ] `tblRunEvents` 65 cột + batch buffer.

**Bước 6 (Migration 20->5 + gỡ POSTING §10):**
- [ ] Phase A/B/C/D; giao thức gỡ POSTING 8 bước; xác nhận ZERO runtime reference.

**Bước 7 (UI 4 nút §9.3):**
- [ ] Tách CHECK khỏi XỬ LÝ; `AssertCheckIsReadOnly()`; button state machine; worklist 15 cột / 4 cột editable.

**Bước 8 (Regression toàn bộ + deliverables §11.5):**
- [ ] Pass 22 scenario S01-S22 offline TRƯỚC live.
- [ ] Chạy fixture bug `5042889690` chứng minh đã diệt stale-state/session-reuse.
- [ ] Live test ladder Live1->Live5.
- [ ] Sinh deliverables (xlsm + ZIP + 3 báo cáo + before/after).

**Nguyên tắc xuyên suốt:** offline không ghi SAP; giữ mọi luồng proven live §7; không phá workbook; không auto CF; không retry mù sau unknown; không xóa sheet trước khi zero-reference.

## 13. Phụ lục - bản đồ file quan trọng

### 13.1. File gốc (`G:\SAP\CF_OB\`)
| File | Vai trò |
|---|---|
| `260819_CF_OB_V2.6_LICH_SU.xlsm` | **BASE HIỆN TẠI** (copy để nâng cấp) |
| `260812_CF_OB_V2.4.xlsm` | Mốc gốc, không sửa |
| `260820_CF_OB_V2.6_AUDIT_TO_V2.7.md` | Audit chuyển tiếp |
| `PROMPT_CODEX_CF_OB_V2.7_REDESIGN_4BUTTON_5SHEETS_HISTORY.md` | **Đặc tả đích chuẩn (4 nút/5 sheet)** |
| `PROMPT_CODEX_CF_OB_V2.7_ULTRA_DETAIL.md` | Đặc tả chi tiết (3 nút - non-chuẩn) |
| `PROMPT_CLAUDE_CF_OB_V2.7_ALL_IN_ONE_UNTIL_PASS.md` | Prompt thực thi |
| `CF_OB_CONVERSATION_HISTORY_20260820.md` | Hội thoại Codex (tài sản) |
| `SAP_CONVERSATIONS_STATES_MAPS_PASSED_20260820.md` | State + map SAP đã pass (tài sản) |
| `outbound-wt-execution-learnings-20260817.md` | Bài học tạo WT (tài sản) |
| `CF_OB_CODEX_CLAUDE_HANDOFF_20260819.zip` | Gói handoff (nguồn evidence live) |

### 13.2. Evidence live (`extracted_package\CF_OB_CODEX_CLAUDE_HANDOFF_20260819\current_extracted\`)
- `sheets\RUN_HISTORY.csv` (568KB), `sheets\RUN_SUMMARY.csv`, `sheets\RUN_LOG.csv`, `sheets\REGRESSION_AUDIT.csv` (2.2MB), `sheets\CONFIG.csv`, `sheets\PLAN.csv`, `sheets\PLAN_RELEASE.csv`, `sheets\SAP_ACTION_MAP.csv`, `sheets\SAP_MAP.csv` (377KB).
- `evidence\latest_run_key_findings.md`, `evidence\latest_run_excerpt.csv`.
- Golden source map: sheet `SAP_ACTION_MAP` trong workbook V2.6.

### 13.3. Staging chưa xác minh
- `_codex_patch\v26_current_extract_20260820\` (modSAPActionMap.bas, modSAPActionRegistry.bas, modCF_SAP_MAP_MON.bas, modCFEvidenceV2.bas): **audit có nhắc nhưng KHÔNG tìm thấy trên đĩa** tại thời điểm chốt -> coi là chưa xác minh / cần tạo lại.

### 13.4. DI SẢN V2.5 - đã gộp vào V2.6, KHÔNG dùng nữa
Thư mục `G:\SAP\CF_OB\build\` (nhánh refactor deterministic Group State Machine) là **di sản, đã SUPERSEDED, KHÔNG phải việc đang dở:**
- `CF_OB_V2.5_WORKING.xlsm`, 7 module `modCFX_*` (State/SapAdapter/Planner/Gates/Executor/Replay/Entry) trong `build\vba_src\`, harness replay S01-S22, `evidence\REPLAY_EVIDENCE.csv`, 6 script `scripts\cfx_*.ps1`.
- Ý tưởng lõi của V2.5 (Group State Machine 9 phase, guard writer WT-state duy nhất chống terminal->Active, MON reader tách PRDO) **đã được kế thừa vào thiết kế V2.7** (§9). Bản thân code/harness/evidence V2.5 không resume, không chạy lại, không lưu tiếp.
- Nhánh này CHỈ chứng minh tới mức skeleton REPLAY-proven (Group Scheduler + MON Broker + Replay); Member Allocation mới một phần; Holder Release = 0%; LIVE wiring = 0% - tức đúng những phần V2.7 phải làm lại trên base V2.6.

---
*Hết. File này là checkpoint chốt ngày 2026-08-20. Base = V2.6. Đích = V2.7. V2.5 = di sản đã gộp.*












