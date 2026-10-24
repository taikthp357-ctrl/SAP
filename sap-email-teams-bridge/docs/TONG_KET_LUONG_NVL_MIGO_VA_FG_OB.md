# Tổng kết Nghiệp vụ & Luồng vận hành: FG (Điền OB) và NVL (Ra đơn MIGO 311)

Ngày lập: 24/10/2026.
Tài liệu phân tích và tích hợp vào hệ thống **SAP Email Teams Bridge**.

---

## 1. Nguồn gốc các công cụ chuẩn từ xưởng

Hệ thống điều độ kho thực tế của xưởng được lưu trữ trong các gói công cụ chuẩn:
1. **`261010_OB_TOOL_V1_RC65.xlsm`**: Tool quét và điền đơn Outbound Delivery (FG Thành phẩm).
2. **`261008_Stock_FG_MA_BODY.xlsm`**: Tool quản lý và cảnh báo tồn kho FG MA BODY qua MB52.
3. **`260810_RA_DON_NVL_V31_ZONE.xlsm`**: Tool cơ cấu ra đơn NVL theo Zone Line, chia cụm OB_GROUPS.
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

## 3. Luồng nghiệp vụ 2: Nguyên vật liệu (NVL) — Ra đơn MIGO 311

### A. Chu trình vận hành:
```
[1. Đọc tồn Line 3001 vs Kho 1005] ➔ [2. Xác định thiếu hàng] ➔ [3. Gom OB_GROUPS] ➔ [4. Tạo phiếu MIGO 311]
```
1. **Cân đối tồn kho NVL:**
   - Tra cứu qua MB52 đồng thời 2 kho:
     * **SLoc 1005 (Kho NVL MA):** Kho nguồn cấp hàng.
     * **SLoc 3001 (Line Body) hoặc 2001 (Line XMĐ):** Line sản xuất nhận hàng.
2. **Nhận diện trạng thái cấp phát:**
   - **Line đói hàng (Line Stock $\le 10$ hoặc $= 0$):** Nếu Kho NVL 1005 đang còn tồn $\rightarrow$ **Cần ra đơn MIGO 311 khẩn cấp** để chuyển hàng từ 1005 sang Line 3001.
   - **Line đủ hàng (Line Stock an toàn):** Chưa cần cấp MIGO 311.
   - **Hết hàng toàn bộ:** Cả Kho 1005 và Line đều bằng 0 $\rightarrow$ Báo động hết hàng.
3. **Thao tác MIGO 311 trên SAP GUI (Tham chiếu `MIGO311_RunPost_V24_ErrorGuard.bas`):**
   - Transaction: `MIGO`.
   - Action: `Transfer Posting` (`A04`).
   - Reference: `Other` (`R10`).
   - Movement Type: `311`.
   - SLoc xuất: `1005` $\rightarrow$ SLoc nhập: `3001` (hoặc `2001`).
   - Nhập Material, Quantity và Lô xuất (FIFO từ `/SCWM/POST`).

---

## 4. Tích hợp vào Bot (SAP Email Teams Bridge)

Hệ thống Bot bổ sung 2 lệnh nghiệp vụ điều độ công việc:

| Lệnh Bot | Ý định tự nhiên | Mục tiêu công việc |
|---|---|---|
| **`CHECK_FG_OB`** | `check ob <mã>`, `điền ob <mã>`, `khả dụng ob <mã>` | Báo cáo Tồn FG 3093, Đơn OB đang chờ điền và Lượng hàng khả dụng thực tế để xuất kho |
| **`CHECK_NVL_MIGO`** | `check nvl <mã>`, `ra đơn migo <mã>`, `migo 311 <mã>`, `cấp nvl <mã>` | Báo cáo Tồn Kho 1005 vs Tồn Line 3001 và Khuyến nghị có cần ra phiếu MIGO 311 cấp hàng hay không |
