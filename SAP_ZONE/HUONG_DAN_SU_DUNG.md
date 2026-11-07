# HƯỚNG DẪN SỬ DỤNG ỨNG DỤNG SAP ZONE (v1.6)
**Tự động hóa toàn diện quy trình Ra đơn NVL, Gom nhóm Zone, MIGO 311, Lấy số WHO và In phiếu SAP**

---

## 1. Giới thiệu tổng quan
Ứng dụng **SAP ZONE** (thay thế quy trình Excel VBA `260810_RA_DON_NVL_V31_ZONE(2).xlsm`) giúp loại bỏ hoàn toàn các thao tác thủ công, chống sai sót số liệu và tăng tốc độ xử lý ra đơn NVL lên nhiều lần.

### Điểm nổi bật:
- **Giao diện tối giản 2 Tab**: Tab Vận hành chính (1-Click Run) và Tab Cấu hình hệ thống riêng biệt.
- **Đọc tồn kho tức thì**: Truy vấn trực tiếp từ `/SCWM/MON` (Node `N000000139`, Storage Type `APFL`, Owner `3093`) thay vì giao diện `/SCWM/POST` gây đơ lag.
- **Phân bổ FIFO & Gom nhóm chuẩn**: Tự động gom tối đa 5 mã hàng, tối đa 20 dòng cho mỗi nhóm OB theo từng Zone (dựa trên bảng `zone_map.json`).
- **Tạo đơn MIGO 311 tự động**: Tự động điền mã hàng, số lượng, trạm, plant, sloc, batch và post chứng từ lấy số OB.
- **Liên kết chính xác 3 cột không bị trùng lặp**: Tra cứu `/SCWM/MON` (Node `N000000010`) theo OB, gắn đúng WHO vào từng Material (tránh bẫy cross-join).
- **In phiếu SAP F00112 (2 bản) tự động**: Tự động đặt 2 bản in (`SSFPP-TDCOPIES = 2`) và tiến trình Win32 Watchdog tự động bắt hộp thoại lưu file PDF vào thư mục chỉ định.
- **Xuất dữ liệu 3 cột nhanh**: Nút `[ 📋 Sao chép 3 cột ]` đưa thẳng dữ liệu chuẩn Tab vào Clipboard để dán trực tiếp (`Ctrl + V`) vào file Excel theo dõi.

---

## 2. Hướng dẫn sử dụng chi tiết

### TAB 1: 🚀 VẬN HÀNH CHÍNH (1-Click Dashboard)
1. **Dán dữ liệu yêu cầu**:
   - Sao chép các cột từ bảng Excel của bạn: `[Mã hàng]  [Số lượng]  [Trạm]`.
   - Bấm nút `[ 📋 Dán từ Clipboard ]` hoặc nhấn `Ctrl + V` vào khung nhập.
   - *(Tùy chọn: Có thể bấm `[ 📝 Nạp mẫu thử ]` để xem trước định dạng chuẩn).*
2. **Khởi chạy quy trình**:
   - Đảm bảo SAP GUI đang mở và đã đăng nhập (Hệ thống sẽ hiển thị biểu tượng `🟢 SAP Sẵn Sàng` ở góc trên bên phải).
   - *(Nếu muốn chạy thử nghiệm khi chưa mở SAP, tích chọn ô `[ ] Chạy mô phỏng (Test không cần mở SAP)`).*
   - Bấm nút to màu xanh: `[ ▶ BẮT ĐẦU CHẠY: RA ĐƠN & IN PHIẾU ]`.
   - Xem thanh tiến trình và thông báo trạng thái cập nhật từng bước: Đọc tồn MON $\rightarrow$ Phân bổ FIFO $\rightarrow$ Tạo MIGO 311 $\rightarrow$ Lấy số WHO $\rightarrow$ In phiếu PDF.
3. **Lấy kết quả theo dõi**:
   - Khi hoàn tất, bảng kết quả hiển thị chuẩn 3 cột:
     $$\textbf{Số OB} \quad | \quad \textbf{Mã hàng (Material)} \quad | \quad \textbf{Số đơn (WHO)}$$
     *(Nếu 1 mã hàng có nhiều WHO, các số sẽ được nối bằng dấu `/`, ví dụ: `6028711569/6028711570`).*
   - Bấm `[ 📋 Sao Chép 3 Cột ]` $\rightarrow$ Mở file Excel theo dõi của bạn $\rightarrow$ Nhấn `Ctrl + V` là xong!
   - Bấm `[ 📊 Xuất Excel ]` nếu cần lưu lại toàn bộ bảng kết quả ra file `.xlsx`.
   - Bấm `[ 📁 Mở Thư Mục PDF ]` để xem ngay các file phiếu in PDF vừa được tạo.

---

### TAB 2: ⚙️ CẤU HÌNH HỆ THỐNG
Mọi thông số kỹ thuật được tập trung tại Tab này:
- **1. Kết nối SAP GUI**:
  - `System`: Mặc định `PRD`
  - `Client`: Mặc định `300`
  - `Session Index`: Vị trí cửa sổ SAP (0 = Cửa sổ đầu tiên)
  - Nút `[ 🔌 Thử Kết Nối ]`: Kiểm tra xem SAP Scripting có phản hồi không.
- **2. Tham số EWM & Kho**:
  - `Mã kho (Warehouse)`: `VF02`
  - `Storage Type`: `APFL`
  - `Chủ hàng (Owner)`: `3093`
  - `Plant MIGO` & `SLoc MIGO`: `3093`
- **3. Quy tắc gom nhóm & Phân bổ**:
  - `Số mã tối đa / OB`: `5`
  - `Số dòng tối đa / OB`: `20`
  - `Chiến lược`: `FIFO` (Ưu tiên ngày nhập cũ) hoặc `FEFO` (Hạn dùng trước)
- **4. Cấu hình in ấn**:
  - `Số bản in SAP`: `2` bản
  - `Tự động gọi in và lưu PDF`: Đang bật
  - `Thư mục lưu PDF`: Mặc định `./OUTPUT_PDF` (có thể bấm `[ 📂 Chọn... ]` để đổi thư mục)
- **5. Dữ liệu Zone Mapping (`zone_map.json`)**:
  - Hiển thị số lượng quy tắc đã nạp từ file.
  - Bấm `[ 📝 Xem File JSON ]` để mở và thêm bớt trạm / zone / tổ trưởng nếu có thay đổi dây chuyền.
  - Bấm `[ 🔄 Tải Lại ]` để cập nhật dữ liệu mới mà không cần khởi động lại app.
- Bấm `[ 💾 LƯU CẤU HÌNH HỆ THỐNG ]` sau khi điều chỉnh.

---

## 3. Cấu trúc thư mục phát hành
```text
SAP_ZONE/
├── SAP_ZONE.exe              # File chạy chính độc lập (Không cần cài đặt Python)
├── zone_map.json             # File ánh xạ Trạm -> Zone -> Tổ trưởng
├── config.json               # File lưu cấu hình người dùng
├── HUONG_DAN_SU_DUNG.md      # Hướng dẫn chi tiết
└── OUTPUT_PDF/               # Thư mục chứa các phiếu in PDF được tự động tạo
```

---

## 4. Khắc phục sự cố thường gặp
1. **Lỗi "Không tìm thấy tiến trình SAP GUI"**:
   - Hãy mở SAP Logon và đăng nhập vào hệ thống trước khi bấm Chạy.
2. **Lỗi "Scripting Engine bị tắt"**:
   - Vào SAP GUI: Chọn biểu tượng `Options` (Alt + F12) $\rightarrow$ `Accessibility & Scripting` $\rightarrow$ `Scripting` $\rightarrow$ Tích chọn `Enable scripting` và bỏ tích `Notify when a script attaches to SAP GUI`.
3. **Dialog Lưu PDF không tự đóng**:
   - Đảm bảo máy in mặc định hoặc máy in chọn trên SAP là `Microsoft Print to PDF` để cơ chế Win32 Watchdog bắt đúng tiêu đề cửa sổ `Save Print Output As`.
