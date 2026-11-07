# ⚡ SAP ZONE v1.6 Pro — TỰ ĐỘNG RA ĐƠN NVL, MIGO 311 & IN PHIẾU WHO

Ứng dụng độc lập Windows Executable (Standalone EXE) hiện đại thay thế hoàn toàn file Excel VBA `260810_RA_DON_NVL_V31_ZONE(2).xlsm`.

---

## 🚀 Tính Năng & Nâng Cấp Trọng Điểm (Bản v1.6 Pro Update)

1. **Khắc Phục Triệt Để Lỗi Kết Nối SAP GUI (Thread COM Initialization & Modal Popup Dismiss)**:
   - Áp dụng chuẩn kết nối đa luồng từ **`SAP_EMAIL_TEAMS_BRIDGE v1.1.12`**.
   - Khởi tạo căn hộ COM (`pythoncom.CoInitialize()`) tự động trong toàn bộ Worker Thread Windows và giải phóng an toàn khi hoàn tất luồng.
   - Cơ chế quét phiên thông minh: Ràng buộc danh tính **System (PSV)** và **Client (908)**, tự động bỏ qua các tab SAP đang bận (`Busy == True`), ưu tiên các tab `SESSION_MANAGER` (Easy Access).
   - Vòng lặp thử lại (Retry Loop) 5 lần khi kết nối COM.
   - Tự động đóng các hộp thoại cảnh báo/popup đang treo trên `wnd[1]` bằng `dismiss_popups()` trước khi thực hiện thao tác.
   - Điều hướng an toàn cho `/SCWM/MON` (`open_mon_screen`) không phụ thuộc vào Favorite Node F00077, đảm bảo chạy ổn định trên mọi máy trạm.

2. **Cơ Chế Phân Giải 4 Tầng Ưu Tiên Zone Thông Minh**:
   - **Tầng 1 (Ưu tiên tuyệt đối - Input Explicit Zone)**: Nếu dòng dán từ Excel có chỉ định Zone (cột 4, cột 5 hoặc trích xuất từ cột Trạm như `Zone 3.1`, `Z2`, `1.2`), hệ thống dùng trực tiếp Zone này. Tuyệt đối không bị file tĩnh `zone_map.json` ghi đè!
   - **Tầng 2 (Tra cứu danh mục `zone_map.json`)**: Nếu dòng dán không có Zone, hệ thống tự động tra cứu mã hàng trong bảng mapping để lấy Zone và TL đã cấu hình.
   - **Tầng 3 (Gom nhóm động theo Trạm / Người nhận)**: Nếu mã hàng chưa có trong `zone_map.json` và dòng dán không có Zone (ví dụ: `BIW32211725AA 300 Hùng/Long`), hệ thống **không báo lỗi** mà tự động lấy Trạm (`Hùng/Long`) làm chìa khóa gom nhóm! Toàn bộ vật tư của `Hùng/Long` được gom chung vào các nhóm đơn OB của `Hùng/Long` (tối đa 5 mã hàng, 20 dòng mỗi OB).
   - **Tầng 4 (Mặc định)**: Nếu để trống cả Zone và Trạm, gom vào nhóm `"Chung"`.

3. **Bộ Phân Tích Dữ Liệu Dán Excel Linh Hoạt (`parse_paste_data`)**:
   - Nhận diện thông minh:
     - 2 cột: `[Mã hàng]  [Số lượng]`
     - 3 cột: `[Mã hàng]  [Số lượng]  [Trạm]` (hoặc `[Số lượng]  [Mã hàng]  [Trạm]`)
     - 3 cột có Zone: `[Mã hàng]  [Số lượng]  [Zone]`
     - 4 cột: `[Mã hàng]  [Số lượng]  [Trạm]  [Zone]`
     - 4 cột có STT: `[STT]  [Mã hàng]  [Số lượng]  [Trạm/Zone]`
     - 5 cột có STT: `[STT]  [Mã hàng]  [Số lượng]  [Trạm]  [Zone]`
   - Tự động lọc bỏ dòng tiêu đề và chuẩn hóa dấu thập phân số lượng.

4. **Tách Bạch Kho Xuất (1005) & Kho Nhận (3001) Cho MIGO 311**:
   - Khắc phục triệt để lỗi MIGO 311 cùng kho (`3093 == 3093`) khiến SAP từ chối chuyển kho.
   - Kho xuất SLoc `1005` (APFL NVL) $\rightarrow$ Kho nhận SLoc `3001` (Line sản xuất / Trạm).
   - Kiểm tra chặt chẽ từng ô bảng MIGO, ngăn chặn tuyệt đối tình trạng post chứng từ rỗng.

5. **Tra Cứu WHO & In Phiếu SAP F00112 (2 Bản) Tự Động Bắt PDF**:
   - Tra cứu `/SCWM/MON` (Node `N000000010`) theo từng OB, gắn chính xác số đơn WHO vào từng Material (triệt tiêu hoàn toàn lỗi nhân chéo cross-join).
   - In 2 bản tự động qua Win32 Watchdog, xác thực file PDF thực tế (> 500 bytes) trước khi báo thành công.

6. **Bảng Kết Quả 3 Cột Chuẩn & 1-Click Sao Chép**:
   - Hiển thị chuẩn 3 cột: `Số OB | Mã Hàng (Material) | Số Đơn (WHO)`.
   - Nút `📋 Sao Chép 3 Cột (Dán Vào File Theo Dõi)` đưa trực tiếp dữ liệu phân tách Tab vào Clipboard để dán ngay bằng `Ctrl + V`.

---

## 📦 Danh Sách Tệp Tải Về

| Tệp tải về | Dung lượng | Mô tả |
| :--- | :--- | :--- |
| **`SAP_ZONE.exe`** | ~29.5 MB | File chạy trực tiếp không cần cài đặt Python hay bất kỳ phần mềm phụ trợ nào |
| **`SAP_ZONE_v1.6_PORTABLE.zip`** | ~28.5 MB | Gói đầy đủ gồm file EXE + `zone_map.json` + `HUONG_DAN_SU_DUNG.md` + `RELEASE_NOTES.md` |
