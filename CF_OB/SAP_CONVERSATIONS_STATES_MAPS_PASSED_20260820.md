# SAP — hội thoại, trạng thái, map và artifact PASS

Ngày đóng gói: 2026-08-20 (Asia/Bangkok)  
Phạm vi: các task Codex có nội dung SAP/EWM/MIGO/GR tại `Custom Office Templates`, cùng workbook, log và source cục bộ còn truy xuất được.

> Đây là chỉ mục bằng chứng/handoff. Nó không cấp quyền thực hiện SAP write, Cancel, Posting, Save hay Confirm.

## 1. Cách đọc mức bằng chứng

| Mức | Ý nghĩa |
|---|---|
| `LIVE_PASS` | Đã có evidence chạy SAP thành công với post-verify chính xác. |
| `VERIFIED` | Đã kiểm chứng ID/control/read contract, nhưng chưa tự nó chứng minh một write thành công. |
| `DISCOVERED` | Đã nhìn thấy/capture; chưa được phép dùng cho RUN. |
| `PENDING` | Chưa đủ evidence; fail-closed. |
| `HISTORICAL_PASS` | Từng PASS ở file/thread cũ; chỉ để tham khảo, phải map/verify lại trước khi dùng. |

## 2. Hội thoại SAP theo thứ tự thời gian

| Thời điểm | Task ID | Phạm vi | Kết quả cần giữ |
|---|---|---|---|
| 2026-08-10 | `019feca7-aded-7711-80b6-eb9d2d38694c` | CF OB V0.5 / `/SCWM/MON` | Baseline đọc `OB → Document → WHO/WT`; module MON có retry bounded và log lỗi theo stage. Mục tiêu auto-CF ban đầu đã bị thay thế. |
| 2026-08-11 | `019fed1b-131a-7830-809d-74c566eaba3f` | CF OB → VBA | Dùng SAP session core, tải `POSTING`/`WH_ACTUAL`, ghi `RUN_LOG`; nền tảng cho tool Excel. |
| 2026-08-11 | `019ff0c9-844b-7cb2-8fab-3136d97f909c` | `/SCWM/PRDO` | **HISTORICAL_PASS:** màn `OIP_CHANGE`, cột `QTY_UI`, các cột `STATUS_PICKING/STATUS_GI/DOCNO_OD`, Save qua `sendVKey 11`. Cần remap lại trước RUN V2.6. |
| 2026-08-12 | `019ff60b-4ba3-7ff1-bb61-d425b97d4e18` | GR EWM | Rà hai luồng GR và lỗi cấu trúc VBA tại `GR_EWM_Combined`; xác định nơi áp `PROCTY=1086`. |
| 2026-08-12 | `019ff723-dfbe-71c2-a8e1-bb7f27e0ad8f` | GR EWM | **HISTORICAL_PASS:** `Post_GR_EWM` và `GR_EWM_Combined` được khóa chỉ `PROCTY=1086`; sửa `For without Next` và lỗi sau `End Sub`. |
| 2026-08-14 | `01a000fb-9b31-7db3-8c72-2e0691610519` | CF OB config | Fail-closed, khóa map/fingerprint và hạ rủi ro treo flow. |
| 2026-08-14 | `01a00239-8efa-7f40-a823-1b75d8452250` | Bin/batch | Chốt user input Qty/Bin/Batch, kiểm tồn exact, không FIFO. |
| 2026-08-14 → hiện tại | `01a000c9-8105-7631-8b5c-ab565114fd53` | CFOBV1 | Nguồn evidence CF_OB: preflight, batch/bin change, WT create, dừng CF user handoff. |
| 2026-08-15 | `01a0066f-ffb4-7a80-9d0d-53078779540d` | CF2 | Registry/action-map, CF khóa user-owned. |
| 2026-08-15 → 2026-08-17 | `01a00755-8138-7c62-8d32-00b6ab1861e9` | CF3 | `CANCEL_TO_ZERO`, `DELTA_ADD`, create delta WT rồi post-verify; `CF_NOT_SENT`. |
| 2026-08-16 | `01a00ab6-7c5a-7d80-9195-08fcf94e1706` | RA ĐƠN/MIGO | **HISTORICAL_PASS:** đọc posting theo trang, không copy grid chưa nạp; chuẩn hóa qty SAP nguyên. |
| 2026-08-17 | `01a00ff4-d8e1-70d0-9bc2-cac9cc86b1ae` | CF Button 4 | PLAN/executor tách riêng; session/stock/context không chứng minh thì STOP. |
| 2026-08-17 | `01a01119-00bb-78e2-a43c-e918fc5075fd` | Exact source row | Sửa false-stop: lượng nguồn phải dùng `OpenQty/Delta`, không dùng tổng PRDO. |
| 2026-08-17 → 2026-08-18 | `01a0111d-6265-71b3-b110-1826a5863c00` | Fast-path CF OB | Exact source → `UI_VSOLA` → selection readback → Adopt → Create/Save → MON post-verify. |
| 2026-08-19 | `01a01a22-0b12-76b3-a8f5-3ed37870ecb2` | V2.5 | Nhánh group-executor/outbound WT; không dùng task lỗi hệ thống làm bằng chứng write. |
| 2026-08-19 → 2026-08-20 | `01a01b39-1d37-7fd2-b4d5-3ab6d4dbf22a` | V2.5 continuation | Group current/target, delta qty, source selection, unknown-state recovery. |
| 2026-08-20 | `01a01b6e-d4e6-7450-bc7e-9822138ac848` | V2.6 MAP mode | Thêm nút map SAP; không tự mở rộng quyền RUN. |

## 3. Trạng thái/response đã ghi nhận

Các mã dưới đây đa phần là **tool status**, được ghi từ SAP readback/context; chúng không phải lúc nào cũng là nguyên văn message SAP.

| Mã / response | Ý nghĩa và hành động đúng |
|---|---|
| `PLAN_EXECUTABLE` | CHECK đã lập kế hoạch, không có SAP write (`SAP_WRITE=0`). Có thể vào executor sau preflight mới. |
| `PLAN_WITH_STOPS` | Có một hoặc nhiều dòng không đủ điều kiện; chỉ các dòng độc lập PASS mới có thể chạy. |
| `PASS|LOCKED_PLAN_PHASES_COMPLETE` | Một lượt executor kết thúc theo checkpoint khóa; `C=0` nghĩa là chưa CF. Vẫn cần MON post-verify chính xác trước `READY_FOR_CF`. |
| `WT_POST_VERIFY_PASS` / `READY_FOR_CF` | Terminal success duy nhất của outbound flow: WT mới được MON xác minh và chờ user CF. |
| `CF_NOT_SENT` / `CF=USER_HANDOFF` | Đúng yêu cầu: code không xác nhận WT. |
| `STOP|SCHEDULER_HALTED_OR_BLOCKED` | Queue có blocker; không thử ép chạy các dòng bị skip. |
| `STOP|STATE_UNKNOWN` hoặc `STATE_UNKNOWN_AFTER_CREATE_SAVE` | Outcome write chưa chứng minh. Bắt buộc reload/reconcile, **không retry** Set Qty/Adopt/Create/Save. |
| `STOP|OB_REGRESSION_GATE_BLOCKED` | Regression/contract chưa đạt; dừng trước SAP write. |
| `STOP_BATCH_SOURCE_ROW_NOT_FOUND_OR_INSUFFICIENT` | Lịch sử đã có false-stop nếu code dùng tổng PRDO thay Open/Delta Qty. Luồng hiện phải match Product+Bin+Batch và `Available >= RequiredDelta`. |
| `PRDO Qty = 0` | Trong `CANCEL_TO_ZERO`: Cancel WT exact → readback → PRDO Qty=0 → readback 0 → không tạo WT mới, không CF. |
| `Khong cuon duoc bang MIGO toi dong 12` | UI-grid/MIGO không đưa được dòng logic 12 vào vùng nhìn thấy; không phải posting thành công và không có write được phép coi là hoàn tất. |
| `SAP chi nhan so nguyen` | Contract parser MIGO: `1.000`, `1,000`, `1.000,00`, `1,000.000` → `1000`; phần lẻ thực (`1,5`, `1.500,25`) bị chặn trước SAP. |

Các evidence V2.6 đáng chú ý trong `RUN_SUMMARY`:

- OB `1011832369` và `1011839385`: có `PASS|LOCKED_PLAN_PHASES_COMPLETE`, `C=0`.
- OB `1011832350`: `STATE_UNKNOWN_AFTER_CREATE_SAVE`, không được retry write.
- OB `1011837627`: `STATE_UNKNOWN`, chỉ reconciliation/read-only.
- `CHECK` trong các bản ghi này đều có `SAP_WRITE=0`.

## 4. SAP action map hiện hành — V2.6

Golden source là sheet `SAP_ACTION_MAP` trong `260819_CF_OB_V2.6.xlsm`.

| Nhóm | Action/key đã đủ bằng chứng | Status | Ghi chú dùng thực tế |
|---|---|---:|---|
| MON | `MON_WT_GRID_ID`; `MON_WT_ACTION_ID=N000000096` | VERIFIED | Đọc exact `DOCNO/WHO/TANUM/ITEMNO/MATNR/CHARG/VLPLA/VSOLM/TOSTAT`. |
| Source stock | tab Stk Can Be Removed; stock grid; `SelectedRows` | VERIFIED | Chỉ lookup exact Product/Bin/Batch, không row index. |
| Qty/Adopt | `UI_VSOLA`; `OK_OD_AVSTKDLV_TAKE_QTY` | LIVE_PASS | Ghi qty rồi readback; chọn dòng rồi readback; không repeat sau unknown outcome. |
| Draft WT | tab/grid Warehouse Task | VERIFIED | Chỉ là surface kiểm tra sau Adopt, không thay Adopt. |
| Create/Save | `OK_OI_DLV_CREATE_TO`; `OK_OI_DLV_SAVE`; composite `Create > Save` | LIVE_PASS / LOCKED | Chuỗi outbound write được phép. |
| CF | `MON_CONFIRM_BG_COMMAND_ID` | VERIFIED kỹ thuật, **RUN disabled** | CF thuộc user, không được code gọi. |

Không được dùng ở RUN V2.6: `MON_CANCEL_COMMAND_ID`, `WT_CREATE_AND_SAVE_COMMAND_ID`, `WT_DELETE_COMMAND_ID`, `WT_SOURCE_BIN_ID`, `WT_QUANTITY_ID`, `OB_EDIT_COMMAND_ID`, `OB_QUANTITY_COL` — các mục này `PENDING` hoặc `DISCOVERED`.

## 5. File/artifact có PASS hoặc evidence rõ

| Artifact | Tình trạng | Cách dùng |
|---|---|---|
| `C:\Users\taipt8\Documents\Custom Office Templates\260819_CF_OB_V2.6.xlsm` | Production hiện hành; có `SAP_ACTION_MAP`, `RUN_LOG`, `RUN_HISTORY`, `RUN_SUMMARY` | Duy nhất dùng làm baseline RUN. |
| `C:\Users\taipt8\Documents\Custom Office Templates\260819_CF_OB_V2.6_LICH_SU.xlsm` | Bản lưu lịch sử | Tham khảo/đối chiếu, không mặc định là production. |
| `C:\Users\taipt8\Documents\Custom Office Templates\260819_CF_OB_V2.6_PREPATCH_BACKUP.xlsm` và `...before_map_sap_button_20260820.xlsm` | Backup recoverable | Rollback, không dùng chạy mới. |
| `_codex_patch\v26_current_extract_20260820\modSAPActionMap.bas`, `modSAPActionRegistry.bas`, `modCF_SAP_MAP_MON.bas`, `modCFEvidenceV2.bas` | Extract của V2.6 | Review/patch staging; không tự chứng minh đã import. |
| `_codex_patch\_archive\runtime\v13b-probe\v13b_probe.status` | `PROBED` | Bằng chứng probe cũ, không thay thế full regression V2.6. |
| `260810_V27.8.5_RA_DON_NVL_V8_5_SPEED_ENTER_FAST.xlsm` | HISTORICAL_PASS trong task RA ĐƠN; không còn trong danh sách workbook hiện tại | Không dùng làm production CF OB; chỉ tham khảo parser integer/posting loader. |
| `250528_FILE GR TOOL 3.xlsm` / `250528_FILE GR TOOL 31.xlsm` | HISTORICAL_PASS trong task GR; không còn trong danh sách workbook hiện tại | Chỉ tham khảo logic `PROCTY=1086`; cần locate/verify lại nếu khôi phục. |

## 6. Quy tắc handoff

1. Chỉ `V2.6` + sheet `SAP_ACTION_MAP` quyết định control/action cho outbound WT hiện tại.
2. Bất kỳ `PENDING`, `DISCOVERED`, `STATE_UNKNOWN`, `PLAN_WITH_STOPS` nào đều không được biến thành write bằng suy diễn.
3. Mỗi write route chỉ có một executor; CHECK/MAP/AUDIT độc lập chỉ đọc.
4. Sau write, đọc lại SAP và ghi ledger. Thành công chỉ là `READY_FOR_CF`; CF vẫn do user thao tác.
5. Không sử dụng screenshot/mouse/SendKeys/tọa độ thay cho technical map.

## 7. Liên kết hồ sơ

- CF-only timeline: `CF_OB_CONVERSATION_HISTORY_20260820.md`
- Current production map/log: `260819_CF_OB_V2.6.xlsm` → `SAP_ACTION_MAP`, `RUN_LOG`, `RUN_HISTORY`, `RUN_SUMMARY`
- Staging is not production: `C:\Users\taipt8\Documents\Custom Office Templates\_codex_patch\`
