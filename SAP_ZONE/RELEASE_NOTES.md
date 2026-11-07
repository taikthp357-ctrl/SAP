# ⚡ SAP ZONE v1.6 Pro — TỰ ĐỘNG RA ĐƠN NVL, MIGO 311 & IN PHIẾU WHO

Ứng dụng độc lập Windows Executable (Standalone EXE) hiện đại thay thế hoàn toàn file Excel VBA `260810_RA_DON_NVL_V31_ZONE(2).xlsm`.

---

## 🚀 Tính Năng & Sửa Lỗi Trọng Điểm (Bản v1.6 Pro)

1. **Khắc Phục Triệt Để Lỗi Kết Nối SAP GUI (Thread COM Initialization)**:
   - Áp dụng chuẩn kết nối đa luồng từ **`SAP_EMAIL_TEAMS_BRIDGE v1.1.12`**.
   - Khởi tạo căn hộ COM (`pythoncom.CoInitialize()`) tự động trong toàn bộ Worker Thread Windows và giải phóng an toàn khi hoàn tất.
   - Cơ chế quét phiên thông minh: Ràng buộc danh tính **System (PSV)** và **Client (908)**, tự động bỏ qua các tab SAP đang bận (`Busy == True`), ưu tiên các tab `SESSION_MANAGER` (Easy Access).
   - Cơ chế thử lại (Retry Loop) 5 lần khi lấy Scripting Engine từ COM.

2. **Nâng Cấp Giao Diện Web UI Hiện Đại (Modern Glassmorphism & Universal Dual-Bridge)**:
   - Giao diện Web App chuyên nghiệp xây dựng trên Edge Chromium WebView2 với CSS3 Glassmorphism, Stepper 6 bước trực quan và bảng dữ liệu phản hồi nhanh.
   - **Universal Dual-Bridge**: Tự động nhận diện môi trường, chạy mượt mà trên cả WebView2 nhúng và chế độ Edge App Mode (`msedge.exe --app=...`) thông qua HTTP REST API nội bộ.
   - Bảng kết quả 3 cột chuẩn: `Số OB | Mã Hàng | Số Đơn (WHO)`, nút `📋 Sao Chép 3 Cột` đưa dữ liệu phân tách Tab vào Clipboard để dán trực tiếp vào file Excel theo dõi (`Ctrl + V`).

3. **Tách Bạch Kho Xuất (1005) & Kho Nhận (3001) Cho MIGO 311**:
   - Khắc phục lỗi MIGO 311 cùng kho (`3093 == 3093`) khiến SAP từ chối chuyển kho.
   - Cấu hình chuẩn công nghiệp: Kho xuất SLoc `1005` (APFL NVL) $\rightarrow$ Kho nhận SLoc `3001` (Line sản xuất / Trạm).
   - Kiểm tra chặt chẽ giá trị ghi vào từng ô bảng MIGO (`_set_cell`), nếu SAP từ chối ghi ô sẽ dừng ngay và báo lỗi, không bao giờ post chứng từ rỗng.
   - Quy trình 2 pha (Phase 1 điền dữ liệu chung, Phase 2 điền Batch) chuẩn xác 100%.

4. **Nhật Ký SQLite Journal Bền Vững (`data/sap_zone_journal.db`)**:
   - Lưu trữ toàn bộ lịch sử các lần chạy (`runs`), nhóm đơn (`groups`) và kết quả ra đơn (`results`).
   - Tự động khôi phục bảng kết quả của lần chạy gần nhất ngay khi khởi động ứng dụng, không lo mất dữ liệu khi vô tình đóng app.

5. **Đọc Tồn MON Chuẩn FIFO & Không Trừ Lặp**:
   - Truy vấn `/SCWM/MON` (Node `N000000139`, Storage Type `APFL`, Owner `3093`).
   - Đọc cột Available Stock thực tế (`ABSQUAN` / `AVAILABLE_QTY`), không trừ thêm cột Outgoing đã được SAP tính sẵn.
   - Sắp xếp chuẩn FIFO theo ngày giờ nhập kho (`GR Date + Time`).

6. **Tra Cứu WHO & In Phiếu SAP F00112 (2 Bản) Tự Động Bắt PDF**:
   - Tra cứu `/SCWM/MON` (Node `N000000010`) theo từng OB, gắn chính xác số đơn WHO vào từng Material (triệt tiêu hoàn toàn lỗi nhân chéo cross-join).
   - In 2 bản tự động qua Win32 Watchdog, xác thực file PDF thực tế (> 500 bytes) trước khi báo thành công.

---

## 📦 Danh Sách Tệp Tải Về

| Tệp tải về | Dung lượng | Mô tả |
| :--- | :--- | :--- |
| **`SAP_ZONE.exe`** | ~29 MB | File chạy trực tiếp không cần cài đặt Python hay bất kỳ phần mềm phụ trợ nào |
| **`SAP_ZONE_v1.6_PORTABLE.zip`** | ~28.5 MB | Gói đầy đủ gồm file EXE + `zone_map.json` + `HUONG_DAN_SU_DUNG.md` |
