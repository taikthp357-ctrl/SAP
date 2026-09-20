# PROMPT NÂNG CẤP CF_OB V2.8 — TỐI ƯU GIAO DIỆN & GOM 5 SHEET
## File đích: `260820_CF_OB_V2.8.xlsm` (Phát triển tiếp từ bản V2.7 đã PASS 100%)

---

### 1. BỐI CẢNH VÀ MỤC TIÊU V2.8

Bản **V2.7** (`260820_CF_OB_V2.7.xlsm`) đã hoàn thành và PASS 100% bộ kiểm thử nghiệm thu lõi:
- Member Allocation Engine đã chuẩn (`CONTRACT_A` $\ne$ `CONTRACT_B`, case $3500+4500 \rightarrow 0+8000$, blank khác 0).
- Chế độ `CHECK` là Read-Only tuyệt đối (`SAP_WRITE = 0`).
- Đã bịt lỗ hổng Seam 2 (khóa plan trước khi Create/Save).
- Single MON Broker đã được cô lập chống xung đột session.
- Nút `LÀM MỚI` (`CF_ResetUI`) đã hoạt động chuẩn xác.

**Nhiệm vụ của V2.8:** Hoàn thiện nốt 2 mục tiêu lớn cuối cùng về **Giao diện người dùng (UI)** và **Kiến trúc dữ liệu (5-Sheet Consolidation & Gỡ bỏ sheet POSTING)** mà không làm hỏng bất kỳ logic nghiệp vụ nào của V2.7.

---

### 2. YÊU CẦU KIẾN TRÚC 5 SHEET CHUẨN (CANONICAL 5-SHEET MODEL)

Chuyển đổi toàn bộ 20 sheet cũ thành mô hình 5 sheet chuẩn:

| Tên Sheet đích | Trạng thái hiển thị | Gom từ các sheet cũ | Nhiệm vụ |
|---|:---:|---|---|
| **`INPUT`** | **Visible (Duy nhất)** | `INPUT` + Bảng nhập `WH_ACTUAL` | Giao diện điều khiển trung tâm duy nhất cho người dùng cuối. |
| **`DATA_CORE`** | `VeryHidden` | `WH_ACTUAL`, `OB_QUEUE`, `OB_WH_CACHE`, `POSTING_RAW`, `POSTING_SUMMARY`, `POSTING_CACHE` | Lưu trữ toàn bộ snapshot dữ liệu thô từ SAP và tồn kho nguồn. |
| **`PLAN`** | `VeryHidden` | `PLAN`, `PLAN_SOURCE_LEGS`, `PLAN_RELEASE` | Chứa kế hoạch phân bổ đã được khóa (Locked PlanID, Hash). |
| **`SYSTEM`** | `VeryHidden` | `CONFIG`, `SAP_MAP`, `SAP_ACTION_MAP`, `SAP_PROBE_LOG` | Lưu trữ cấu hình kỹ thuật và ánh xạ trường SAP GUI. |
| **`RUN_HISTORY`** | `VeryHidden` | `RUN_HISTORY`, `RUN_LOG`, `RUN_SUMMARY` | Sổ cái lịch sử chạy bền vững (mở rộng đầy đủ các trường đối soát). |

*(Lưu ý: Sheet `ORDER_QUERY` và `REGRESSION_AUDIT` chuyển vào kho lưu trữ nội bộ hoặc drop an toàn).*

---

### 3. QUY TRÌNH GỠ BỎ SHEET `POSTING` AN TOÀN (8-STEP PROTOCOL)

Sheet `POSTING` hiện đang có các tham chiếu runtime từ `modCF_Posting` (ghi MIGO 311) và `ModuleSAP`, `modSAPReadOnlyProbe` (đọc). Phải thực hiện đúng quy trình 8 bước sau:

1. **Xây dựng `SourceStockRepository`:** Tạo module in-memory repository + vùng lưu trữ `tblSourceStock` trên sheet `DATA_CORE`.
2. **Chuyển đổi các module ĐỌC (`Reader`):** Thay thế toàn bộ `Worksheets("POSTING")` trong `ModuleSAP` và `modSAPReadOnlyProbe` sang đọc từ `SourceStockRepository`.
3. **Chuyển đổi module GHI (`Writer`):** Cập nhật `modCF_Posting` để ghi kết quả vào `SourceStockRepository` / `DATA_CORE` thay vì gán cứng `wsRaw.Name = "POSTING"`.
4. **Compile & Self-test:** Đảm bảo toàn bộ dự án compile sạch sẽ không còn lỗi tham chiếu.
5. **Chạy hồi quy Offline:** Toàn bộ 11 Scenario của Allocation Engine và Fast Module Regression phải tiếp tục PASS 100%.
6. **Kiểm tra Zero-Reference:** Thực hiện `grep` tĩnh đảm bảo không còn bất kỳ dòng code nào chứa chuỗi `Worksheets("POSTING")` hoặc `Sheets("POSTING")`.
7. **Đặt `VeryHidden` bước đệm:** Chuyển sheet `POSTING` sang `xlSheetVeryHidden` để kiểm tra độ ổn định.
8. **Xóa hẳn sheet `POSTING`:** Xóa vĩnh viễn khỏi workbook sau khi các bước trên đã được chứng minh.

---

### 4. TỐI ƯU HÓA GIAO DIỆN NGƯỜI DÙNG TRÊN SHEET `INPUT`

Thiết kế lại sheet `INPUT` thành một Bảng điều khiển (Control Center) hiện đại, sạch sẽ và tiện dụng:

#### 4.1. Bảng Worklist 15 cột trực tiếp trên `INPUT`
- Di chuyển vùng nhập liệu từ sheet `WH_ACTUAL` cũ lên trực tiếp mặt sheet `INPUT`.
- **Cấu trúc 15 cột hiển thị:**
  1. `STT`
  2. `OB`
  3. `Chứng từ (Doc)`
  4. `Item`
  5. `Mã VT (Material)`
  6. `Tên VT / Quy cách`
  7. `WT hiện tại (Sys WT)`
  8. `Số lượng SAP (Sys Qty)`
  9. `Vị trí SAP (Sys Bin)`
  10. `Lô SAP (Sys Batch)`
  11. **`[SỬA] Số lượng đích (Actual Qty)`** *(Tô màu vàng/xanh nổi bật cho phép nhập tay)*
  12. **`[SỬA] Vị trí đích (Actual Bin)`** *(Cho phép sửa)*
  13. **`[SỬA] Lô đích (Actual Batch)`** *(Cho phép sửa)*
  14. **`Trạng thái (Status)`** *(Hiển thị kết quả CHECK / XỬ LÝ)*
  15. **`Ghi chú / Hành động dự kiến (Action Plan)`**
- **Quy tắc lọc Status = C:** Tự động ẩn (`Hidden/Filter`) các dòng có trạng thái `C` (WT đã hoàn tất) khỏi bảng nhìn của người dùng, không gây rối mắt nhưng vẫn lưu đầy đủ trong `DATA_CORE`.

#### 4.2. Khối 4 Nút Bấm Điều Khiển Chuẩn (Action Bar)
- Sắp xếp 4 nút bấm to, rõ ràng, phân biệt màu sắc trực quan tại phần đầu sheet `INPUT`:
  * **Nút 1: `[1] TẢI WH`** (Xanh dương) $\rightarrow$ Gán macro `CF_LoadWH`
  * **Nút 2: `[2] CHECK`** (Vàng cam) $\rightarrow$ Gán macro `CF_CheckAndLoadPosting` *(Read-Only 100%)*
  * **Nút 3: `[3] XỬ LÝ`** (Xanh lá) $\rightarrow$ Gán macro `CF_ExecutePlanToCFHandoff` *(Chỉ sáng khi Plan đã khóa)*
  * **Nút 4: `[4] LÀM MỚI`** (Xám / Trắng) $\rightarrow$ Gán macro `CF_ResetUI` *(Xóa khóa plan cũ, làm sạch bảng)*
- **Ẩn hoàn toàn** 3 nút chẩn đoán kỹ thuật cũ (`btnCF_CheckSAP`, `btnCF_OpenLog`, `btnCF_MapPOST`) khỏi tầm nhìn người dùng.

---

### 5. NGUYÊN TẮC BẤT BIẾN (INVARIANTS KHÔNG ĐƯỢC VI PHẠM)

1. **Bảo toàn 100% Allocation Engine V2.7:**
   - Không được thay đổi công thức tính toán `CFAlloc_ComputeRequiredCreate`.
   - Giữ nguyên việc phân tách `CONTRACT_A` $\ne$ `CONTRACT_B`.
   - Case $3500 + 4500 \rightarrow 0 + 8000$ bắt buộc sinh đúng `CreateQty = 3500`.
   - Quy tắc `blank` khác `0` phải giữ nguyên.
2. **Khóa An toàn Read-Only:** Nút `CHECK` tuyệt đối không được ghi SAP (`SAP_WRITE = 0`).
3. **Bảo toàn Dữ liệu Lịch sử:** Không được làm mất hay xáo trộn 3304 dòng lịch sử trên sheet `RUN_HISTORY`.
4. **Không Auto CF:** Luôn dừng lại ở trạng thái bàn giao `READY_FOR_USER_CF`.

---

### 6. TIÊU CHÍ NGHIỆM THU ĐẦU RA CHO V2.8 (DELIVERABLES CHECKLIST)

Sau khi hoàn thành, bắt buộc phải xuất đầy đủ:
1. File Excel: **`260820_CF_OB_V2.8.xlsm`** (Chỉ gồm đúng 5 sheet, chỉ `INPUT` là Visible).
2. Chạy kịch bản kiểm thử:
   - `CFAlloc_SelfTest` $\rightarrow$ **PASS**
   - `CFReg_RunAll` (11 Scenarios) $\rightarrow$ **11/11 PASS**
   - `CF_FastModuleRegression` $\rightarrow$ **PASS**
   - `CF_Phase3_RunUnitTests` $\rightarrow$ **11/11 PASS**
   - `CFResetUI_SelfTest` $\rightarrow$ **PASS**
   - Kiểm tra Zero Reference `Worksheets("POSTING")` $\rightarrow$ **0 hits**.
3. Báo cáo tài liệu: `V2.8_MIGRATION_AUDIT.md`, `V2.8_CHANGELOG.md`, `V2.8_FINAL_AUDIT.md`.
4. Đóng gói file nén `CF_OB_V2.8_RELEASE.zip` và sẵn sàng đẩy lên GitHub.
