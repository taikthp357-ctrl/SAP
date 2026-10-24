# Tổng kết Nghiệp vụ & Luồng vận hành: FG (Điền OB) và NVL (Posting Change & Ra đơn MIGO 311)

Ngày lập: 24/10/2026.
Tài liệu phân tích và tích hợp vào hệ thống **SAP Email Teams Bridge**.

---

## 1. Nguồn gốc các công cụ chuẩn từ xưởng

Hệ thống điều độ kho thực tế của xưởng được lưu trữ trong các gói công cụ chuẩn:
1. **`261010_OB_TOOL_V1_RC65.xlsm`**: Tool quét và điền đơn Outbound Delivery (FG Thành phẩm).
2. **`261008_Stock_FG_MA_BODY.xlsm`**: Tool quản lý và cảnh báo tồn kho FG MA BODY qua MB52.
3. **`260810_RA_DON_NVL_V31_ZONE.xlsm`**: Tool cơ cấu ra đơn NVL theo Zone Line, tải Posting Change và chia cụm OB_GROUPS.
4. **`MIGO311_RunPost_V24_ErrorGuard.bas`**: Module tự động hóa tạo phiếu chuyển kho MIGO 311 trên SAP GUI.
5. **`261002_MIGO_ATT_V3_4_3.xlsm`**: Đối soát BBBG & chứng từ MIGO chuyển kho.

---

## 2. Luồng nghiệp vụ 1: Thành phẩm (FG) — Điền đơn OB

### A. Chu trình vận hành:
```
[1. MB52 Kho FG 3093] ➔ [2. VL06O Quét đơn OB mở] ➔ [3. Cân đối & Điền OB] ➔ [4. Xuất kho]
```
1. **Kiểm tra tồn kho FG:**
   - Tra cứu qua MB52 tại Plant 3093, SLoc 3093 (Kho FG).
   - **Tuyệt đối loại bỏ SLoc 9999** (Kho QC / Phế phẩm / Blocked) để không tính nhầm hàng hỏng.
2. **Quét đơn OB mở (VL06O):**
   - Điền Shipping Point `3093` (để lọc đúng các đơn giao hàng `3098`).
   - Bấm `btn[18]` chuyển sang Item View để thấy chi tiết từng mã hàng và số lượng `LFIMG`.
   - Lọc các đơn có trạng thái Goods Movement Status `WBSTK = A` (Chưa xuất / Treo).
3. **Đánh giá điều độ công việc:**
   - $\text{Khả dụng xuất mới} = \text{Tồn FG 3093} - \text{Tổng SL các đơn OB đang chờ điền}$.
   - Nếu $\text{Khả dụng} \ge 0$: Đủ hàng để điền đơn OB hiện tại và còn dư để ra đơn mới.
   - Nếu $\text{Tồn FG} < \text{Tổng SL OB}$: Cảnh báo thiếu tồn kho, không đủ hàng để điền OB.

---

## 3. Luồng nghiệp vụ 2: Nguyên vật liệu (NVL) — Posting Change Kho 1005 & Ra đơn MIGO 311

### A. Chu trình vận hành:
```
[1. Quét /SCWM/POST (Posting Change)] ➔ [2. Sắp xếp FIFO theo Ngày/Giờ GR] ➔ [3. Đối chiếu MB52 SLoc 1005] ➔ [4. Sẵn sàng ra đơn MIGO 311]
```
1. **Phạm vi kiểm tra NVL:**
   - **Plant:** `3093`
   - **SLoc:** **CHỈ Ở SLOC 1005** (Kho NVL MA). Tuyệt đối không kiểm tra lan sang SLoc 3001 hay 2001.
2. **Kiểm tra Posting Change (`/SCWM/POST`):**
   - Mở màn hình `/SCWM/POST` với Owner `3093` (Kho 1005).
   - Trích xuất toàn bộ:
     * `Product`: Mã vật tư NVL.
     * `Batch`: Số lô hàng.
     * `Source Bin`: Vị trí Bin lưu trữ chính xác trong kho 1005.
     * `A.Qty AUoM` (`qty`): Số lượng tồn khả dụng thực tế trên Bin.
     * `TST` (`cat`): Trạng thái tồn (FT = Khả dụng, QT/QI = Chờ kiểm định QC, BT = Khóa).
     * `GR Date` & `GR Time`: Ngày và Giờ hoàn tất nhập kho GR.
3. **Sắp xếp thứ tự cấp phát FIFO (First In - First Out):**
   - Lô có Ngày và Giờ GR **cũ nhất** (nhập kho sớm nhất) được đánh số thứ tự ưu tiên **FIFO #1**, #2... để bốc hàng xuất trước cho đơn chuyển kho MIGO 311.
   - Các lô bị khóa hệ thống (Bin `APFL-BL01-01-01`) hoặc trạng thái QC (`QT`/`QI`) được đánh dấu cảnh báo rõ ràng, không đưa vào danh sách xuất.
4. **Đối chiếu bổ trợ qua MB52 (SLoc 1005):**
   - Trường hợp vật tư chưa hiển thị dòng trong `/SCWM/POST`, hệ thống tự động đối chiếu tồn qua MB52 tại Plant 3093 và duy nhất SLoc 1005:
     * Nếu MB52 > 0: Cảnh báo hàng có tồn nhưng chưa lên Bin trong Posting Change.
     * Nếu MB52 = 0: Xác nhận hết hàng tại Kho 1005 (Plant 3093).

---

## 4. Tích hợp vào Bot (SAP Email Teams Bridge)

Hệ thống Bot cung cấp 2 lệnh nghiệp vụ điều độ công việc:

| Lệnh Bot | Ý định tự nhiên | Mục tiêu công việc |
|---|---|---|
| **`CHECK_FG_OB`** | `check ob <mã>`, `điền ob <mã>`, `khả dụng ob <mã>` | Báo cáo Tồn FG 3093, Đơn OB đang chờ điền và Lượng hàng khả dụng thực tế để xuất kho |
| **`CHECK_NVL_MIGO`** | `check nvl <mã>`, `check postingchange <mã>`, `ra đơn migo <mã>`, `migo 311 <mã>` | Quét Posting Change (/SCWM/POST) Kho 1005 (Plant 3093), phân tích FIFO Lô/Bin và đề xuất bốc hàng ra đơn MIGO 311 |
