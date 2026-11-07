# ⚡ SAP ZONE v1.6 — TỰ ĐỘNG RA ĐƠN NVL, MIGO 311 & IN PHIẾU WHO

Ứng dụng độc lập Windows Executable (Standalone EXE) thay thế hoàn toàn file Excel VBA `260810_RA_DON_NVL_V31_ZONE(2).xlsm`.

---

## 🚀 Tính Năng & Cải Tiến Nổi Bật (Bản 1.6)

1. **Giao Diện Tối Giản 2 Tab Chuẩn Công Nghiệp**:
   - **Tab 1 — 🚀 Vận Hành Chính (1-Click Run)**:
     - Nhập liệu thông minh: Dán trực tiếp từ Excel (hỗ trợ phân tách Tab, khoảng trắng, CSV).
     - Nút bấm duy nhất: `[ ▶ BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU ]` kích hoạt toàn bộ chu trình ngầm.
     - Bảng kết quả 3 cột chuẩn: `Số OB | Mã hàng (Material) | Số đơn (WHO)` (gộp nhiều WHO bằng dấu `/`, ví dụ: `6028711569/6028711570`).
     - Nút `[ 📋 Sao Chép 3 Cột ]`: Đưa dữ liệu định dạng Tab vào Clipboard, chỉ cần mở file Excel theo dõi và nhấn `Ctrl + V`.
     - Xuất Excel và mở thư mục chứa file PDF với 1 click.
   - **Tab 2 — ⚙️ Cấu Hình Hệ Thống**:
     - Cô lập toàn bộ thiết lập SAP GUI, kho EWM (VF02, APFL, Owner 3093, Plant 3093, SLoc 3093).
     - Tùy chỉnh giới hạn gom nhóm (Max 5 mã, Max 20 dòng, FIFO/FEFO).
     - Thiết lập in 2 bản và thư mục lưu PDF.
     - Quản lý và nạp lại dữ liệu Zone Mapping (`zone_map.json`).

2. **Đọc Tồn Kho Siêu Tốc Từ /SCWM/MON**:
   - Loại bỏ giao diện `/SCWM/POST` gây đơ lag.
   - Truy vấn trực tiếp từ Node `N000000139` (Available Stock) lọc theo `APFL` và Owner `3093`.
   - Tính toán lượng khả dụng chuẩn (`Available Qty BUoM`), không bị trừ lặp `Outgoing`.

3. **Gom Nhóm Zone & Phân Bổ FIFO/FEFO**:
   - Dựa trên bảng ánh xạ `zone_map.json` (33 quy tắc).
   - Tự động chia nhóm tối đa 5 mã hàng, tối đa 20 dòng để tối ưu công suất nhặt hàng tại xưởng.

4. **Tạo Chứng Từ MIGO 311 Tự Động**:
   - Chạy Transfer Posting $\rightarrow$ Other $\rightarrow$ Movement Type 311.
   - Tự động điền dữ liệu dòng, check và post chứng từ, trích xuất số OB chính xác.

5. **Tra Cứu WHO Triệt Tiêu Lỗi Nhân Chéo (Cross-Join)**:
   - Tra cứu `/SCWM/MON` (Node `N000000010`) theo từng OB.
   - Gắn đúng số đơn WHO vào từng Material cụ thể.

6. **In Phiếu SAP F00112 (2 Bản) & Tự Động Bắt PDF**:
   - Gọi F00112 in 2 bản (`SSFPP-TDCOPIES = 2`).
   - Tiểu trình Win32 Watchdog chạy ngầm tự động nhận diện cửa sổ in Windows `Save Print Output As` và lưu file PDF mà không cần thao tác tay.

---

## 📦 Danh Sách Tệp Tải Về

| Tệp tải về | Dung lượng | Mô tả |
| :--- | :--- | :--- |
| **`SAP_ZONE.exe`** | ~26 MB | File chạy trực tiếp không cần cài đặt Python hay môi trường |
| **`SAP_ZONE_v1.6_PORTABLE.zip`** | ~25.7 MB | Gói đầy đủ gồm file EXE + `zone_map.json` + file Hướng dẫn sử dụng |
