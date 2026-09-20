# CF OB — lịch sử hội thoại, bằng chứng và SAP map

Ngày đóng gói: 2026-08-20 (Asia/Bangkok)  
Phạm vi: các task Codex liên quan trực tiếp đến workbook CF OB, từ V0.5 đến `260819_CF_OB_V2.6.xlsm`.

## 1. Kết luận vận hành hiện tại

Workbook sản xuất hiện hành là `C:\Users\taipt8\Documents\Custom Office Templates\260819_CF_OB_V2.6.xlsm`.

Mục tiêu đã chốt là xử lý OB đến `READY_FOR_CF`; CF luôn do người dùng thực hiện. Tool không được gọi Confirm WT (foreground/background), menu CF, hoặc tự xác nhận bất kỳ WT nào.

`SAP_ACTION_MAP` trong V2.6 là nguồn chuẩn duy nhất cho control/action SAP. Chỉ mục `LIVE_PASS` hoặc `VERIFIED` có đúng context/fingerprint mới được dùng; `DISCOVERED` và `PENDING` chỉ là bằng chứng khảo sát, không được suy diễn để gửi SAP write.

Luồng outbound-WT đã khóa:

```text
Tải WH (đọc MON) → user nhập Qty/Bin/Batch → CHECK tạo PLAN
→ Stk Can Be Removed: tìm đúng Product + Bin + Batch
→ UI_VSOLA = Qty, đọc lại → chọn đúng dòng, đọc lại
→ Adopt Quantity → kiểm tra đúng một WT draft
→ Create rồi Save (composite đã LIVE_PASS)
→ MON post-verify → READY_FOR_CF
```

Sau một write có outcome chưa xác định: chỉ đọc lại/reconcile, không retry mù `Set Qty`, `Adopt`, `Create`, hoặc `Save`.

## 2. SAP map đã xác thực (V2.6)

| Action key | Giá trị kỹ thuật / contract | Trạng thái | Dùng trong RUN |
|---|---|---:|---|
| `MON_WT_GRID_ID` | `wnd[0]/usr/shell/shellcont[1]/shell/shellcont[1]/shell` | VERIFIED | Đọc WT theo `DOCNO/WHO/TANUM/ITEMNO/MATNR/CHARG/VLPLA/VSOLM/TOSTAT` |
| `MON_WT_ACTION_ID` | `N000000096` | VERIFIED | Mở khu Warehouse Task |
| `WT_REQUEST_GRID` | ALV Warehouse Request đã map | VERIFIED | Xác định đúng request, không dùng row index |
| `WT_TAB_STK_CAN_BE_REMOVED` | tab `.../tabpOK_TAB_ODP_AVSTKDLV` | VERIFIED | Surface chọn tồn nguồn |
| `WT_STOCK_GRID` | ALV source-stock đã map | VERIFIED | Lookup Product/Bin/Batch chính xác |
| `WT_STOCK_SET_QTY` | `UI_VSOLA` | LIVE_PASS | Ghi qty nguồn rồi readback |
| `WT_STOCK_SELECT_ROW` | `SelectedRows` | VERIFIED | Chọn và đọc lại đúng dòng nguồn |
| `WT_ADOPT_QUANTITY` | `OK_OD_AVSTKDLV_TAKE_QTY` | LIVE_PASS | Adopt đúng một draft WT |
| `WT_TAB_WAREHOUSE_TASK` / `WT_WAREHOUSE_TASK_GRID` | tab/grid Warehouse Task draft đã map | VERIFIED | Chỉ kiểm tra sau Adopt, không thay Adopt |
| `WT_CREATE_COMMAND_ID` | `OK_OI_DLV_CREATE_TO` | LIVE_PASS | Tạo WT; không retry sau lệnh gửi đi |
| `WT_SAVE_COMMAND_ID` | `OK_OI_DLV_SAVE` | LIVE_PASS | Lưu WT; không retry sau lệnh gửi đi |
| `WT_CREATE_SAVE` | `WT_CREATE_COMMAND_ID > WT_SAVE_COMMAND_ID` | LIVE_PASS / LOCKED | Chuỗi write được phép |

Không dùng cho RUN hiện tại:

- `MON_CANCEL_COMMAND_ID`: PENDING, chưa được phép suy diễn/nhấn.
- `WT_CREATE_AND_SAVE_COMMAND_ID`: DISCOVERED; bằng chứng LIVE hiện là **Create rồi Save**, không phải một-click `Create + Save`.
- `WT_DELETE_COMMAND_ID`, `WT_SOURCE_BIN_ID`, `WT_QUANTITY_ID`, `OB_EDIT_COMMAND_ID`, `OB_QUANTITY_COL`: PENDING.
- `MON_CONFIRM_BG_COMMAND_ID` có bằng chứng kỹ thuật cũ, nhưng bị khóa vĩnh viễn bởi contract CF user-owned.

Các cột kiểm tra MON/POST lịch sử đã dùng: `TANUM`, `DOCNO`, `MATNR`, `CHARG`, `VLPLA`, `NLPLA`, `VSOLM`, `TOSTAT`, `AVL_ALT_QTY`, `TO_CAT`.

## 3. Bằng chứng SAP và ID nghiệp vụ đáng giữ

| Bằng chứng | ID / kết quả | Phân loại |
|---|---|---|
| Outbound WT tạo thành công | WT `5042711577`; Monitor post-verify PASS; CF không gửi | LIVE_PASS trong `SAP_ACTION_MAP` |
| Luồng delta đã hoàn tất | Giữ WT `5042769777` Qty 60; PRDO thành 280; tạo WT `5042833176` Qty 220 từ `APFL-P22-02-01`, batch `260809-10`; post-verify PASS, CF_NOT_SENT | Lịch sử V1.5, dùng làm evidence nghiệp vụ |
| Batch/bin posting-change đã được người dùng xác nhận | Material `BIW32026215AA`; Qty 128; bin `APFL-ZF03-01-13`; batch ở vị trí `260713-07` đổi về batch WH `260704-07`; Warehouse Request `1012034149` | Lịch sử V1.5, không phải map write V2.6 |
| Preflight chỉ-đọc V1.5 | OB `1011762630`; Document `1012020387`; WT `5042654963`; WH `6026216040`; Material `BIW70004156`; Qty 110; source `APFL-P22-01-12/260715-12`; target `APFL-P22-02-05/260717-01` | PASS read-only; không có Cancel/Post/Create/CF |
| V2.6 run summary PASS | OB `1011832369` và `1011839385`: `PASS|LOCKED_PLAN_PHASES_COMPLETE`, `C=0` | Có execution record; kiểm tra RUN_LOG trước khi resume |
| Unknown outcome cần dừng | OB `1011832350`: `STATE_UNKNOWN_AFTER_CREATE_SAVE`; OB `1011837627`: `STATE_UNKNOWN` | Chỉ reload/reconcile, cấm retry write mù |

## 4. Dòng thời gian các task liên quan

| Thời điểm | Task ID | Nội dung đã đóng góp |
|---|---|---|
| 2026-08-10 | `019feca7-aded-7711-80b6-eb9d2d38694c` — Nâng cấp CF OB V0.5 | Đặt baseline MON: OB → Document → Warehouse Task → WT/WH; sau đó tách `modCF_MONCore_STABLE.bas`, retry bounded và log `LOAD_MON_NAV`. Mục tiêu auto-CF ban đầu đã bị thay thế bởi CF user-owned. |
| 2026-08-11 | `019fed1b-131a-7830-809d-74c566eaba3f` — Chuyển CF OB sang VBA tool | Chuẩn hóa UI/WH_ACTUAL/POSTING/RUN_LOG và nguyên tắc không suy đoán control SAP. |
| 2026-08-11 | `019ff0c9-844b-7cb2-8fab-3136d97f909c` — Cập nhật CF OB UI và backup | Lấy `/SCWM/PRDO`, item grid và `QTY_UI` làm baseline cho đổi Qty có readback. |
| 2026-08-11 | `019ff0f4-0d6d-79d0-a3ae-8dd7d4b722f7` — Tiếp tục phân tích CF OB | Phân tích MAP/CONFIG và boundary giữa UI, MON, Posting, Log. |
| 2026-08-12 | `019ff5db-0508-7250-8049-69c889f6a318` — V1.1 multi-OB WH loader | Mở rộng tải WH đa-OB và dọn UI không dùng. |
| 2026-08-12 | `01a000fb-9b31-7db3-8c72-2e0691610519` — Nâng cấp config chống lỗi | Bổ sung fail-closed, khóa map/fingerprint và nguyên tắc không treo luồng. |
| 2026-08-14 | `01a00239-8efa-7f40-a823-1b75d8452250` — Vị trí và batch | Chốt ý nghĩa user nhập Qty/Bin/Batch, không FIFO, kiểm tra tồn theo đúng vị trí/batch. |
| 2026-08-14 → hiện tại | `01a000c9-8105-7631-8b5c-ab565114fd53` — CFOBV1 | Task nền: UI nút Tải WH → user input → CHECK → Xử lý đến CF; rework/preflight V1.5 và chuyển sang V2.6. |
| 2026-08-15 | `01a0066f-ffb4-7a80-9d0d-53078779540d` — CF2 | Đóng CF tự động, xây `SAP_ACTION_MAP`/registry, dùng đúng evidence CFOBV1. |
| 2026-08-15 → 2026-08-17 | `01a00755-8138-7c62-8d32-00b6ab1861e9` — CF3 | Hoàn thiện queue Button 4: `CANCEL_TO_ZERO`, `DELTA_ADD`, giữ WT cũ khi tăng qty; bằng chứng WT `5042833176`; CF_NOT_SENT. |
| 2026-08-17 | `01a00ff4-d8e1-70d0-9bc2-cac9cc86b1ae` — Nâng cấp Button #4 theo PLAN | Ràng buộc Button 4 theo PLAN, dừng an toàn khi session/stock/context không chứng minh. |
| 2026-08-17 | `01a01119-00bb-78e2-a43c-e918fc5075fd` — Fix source row selection | Khoanh false-stop ở source row; quy tắc qty nguồn là Open/Delta Qty, không phải tổng PRDO. |
| 2026-08-17 → 2026-08-18 | `01a0111d-6265-71b3-b110-1826a5863c00` — Kiểm tra plan nâng cấp | Chuẩn hóa fast-path: exact source → `UI_VSOLA` → select → Adopt → Create + Save → post-verify → CF handoff. |
| 2026-08-19 | `01a01a22-0b12-76b3-a8f5-3ed37870ecb2` — V2.5 | Chuyển dần sang group executor/outbound WT; task dừng lỗi hệ thống nhưng history còn hiệu lực tham khảo. |
| 2026-08-19 → 2026-08-20 | `01a01b39-1d37-7fd2-b4d5-3ab6d4dbf22a` — Tiếp tục V2.5 | Sửa group current/target, delta qty và exact source selection; các unknown state được khóa retry. |
| 2026-08-20 | `01a01b6e-d4e6-7450-bc7e-9822138ac848` — Add SAP mapping button | Bổ sung nút map SAP phục vụ MAP mode, không mở rộng quyền RUN. |

## 5. Điều không được suy diễn từ lịch sử

- Không tái sử dụng hard-code hoặc flow POST/Cancel cũ của V1.5 vào V2.6 nếu chưa có row `LIVE_PASS`/`VERIFIED` tương ứng trong `SAP_ACTION_MAP`.
- Không coi `CreatedCount > 0` là đã có `READY_FOR_CF`; luôn yêu cầu MON post-verify của WT chính xác.
- Không dùng `CF_PreLiveAudit` (audit CF lịch sử) làm điều kiện cho outbound path mới; outbound path dùng gate/registry V2.6.
- Không dùng điều khiển màn hình, `SendKeys`, tọa độ chuột, quét toàn bộ SAP tree/grid, hay retry write sau unknown outcome.
- Không dùng batch/bin/OB của bảng evidence như dữ liệu đầu vào cho OB mới.

## 6. Quy trình tham khảo cho lượt tiếp theo

1. Mở đúng `260819_CF_OB_V2.6.xlsm`, không dùng V1.5 làm production.
2. Bấm Tải WH; xác nhận MON snapshot thuộc đúng OB/Document/WT.
3. User nhập Qty/Bin/Batch; bấm CHECK để sinh PLAN. Không ghi SAP ở CHECK.
4. Chỉ chạy executor khi PLAN exact, stock exact và action registry hợp lệ. Một executor sở hữu SAP writes.
5. Sau create/save, đọc lại MON. Thành công chỉ khi `READY_FOR_CF`; CF do user thao tác riêng.
6. Nếu `STOP_*` hoặc `STATE_UNKNOWN`, giữ nguyên trạng thái, ghi evidence và chỉ resume sau read-only state check mới.

## 7. Nguồn đối chiếu cục bộ

- Production: `C:\Users\taipt8\Documents\Custom Office Templates\260819_CF_OB_V2.6.xlsm`
- Golden map: sheet `SAP_ACTION_MAP` trong workbook V2.6
- Runtime evidence: sheets `RUN_LOG`, `RUN_HISTORY`, `RUN_SUMMARY`
- Current exported modules: `C:\Users\taipt8\Documents\Custom Office Templates\_codex_patch\v26_current_extract_20260820\`
- Fast-path constraints: `\\vfs-file02-vp\Private\taipt8\My Documents\Custom Office Templates\AGENTS.md`

Tài liệu này là hồ sơ tham khảo và handoff; không tự cho phép một SAP write, Cancel, Posting Change hoặc CF.
