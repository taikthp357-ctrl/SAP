# ⚡ SAP ZONE v1.7 Pro (Build 20261110) — CHUẨN HÓA TOÀN DIỆN TRA CỨU MON, ĐÓNG SỐ INWH & AN TOÀN EWM

Ứng dụng độc lập Windows Executable (Standalone EXE) siêu nhẹ (< 14 MB), tốc độ mở tức thì (< 1s), vận hành an toàn tuyệt đối, thay thế hoàn toàn macro VBA và chuẩn hóa quy trình ra đơn kho EWM.

---

## 🚀 Tính Năng & Nâng Cấp Trọng Điểm trong v1.7 Pro

### 1. Chế Độ Tra Cứu Read-Only 100% — Tuyệt Đối Không Confirm Về "C"
- **Xóa bỏ hoàn toàn mã lệnh Confirm**: Loại bỏ triệt để việc gọi menu ngữ cảnh `METHODS -> @M00007` và nút `btn[0]`. Trước đây mã `@M00007` làm SAP EWM tự động xác nhận hoàn tất Picking trên hệ thống (chuyển trạng thái sang `C` - Confirmed).
- **An toàn dữ liệu kho**: Thao tác trên `/SCWM/MON` hiện tại chỉ thuần túy đọc và trích xuất dữ liệu từ các lưới ALV Grid, không bao giờ thay đổi trạng thái đơn trên SAP.

### 2. Giữ Nguyên Cửa Sổ MON — Chấm Dứt Đột Ngột Tắt Màn Hình
- **Loại bỏ lệnh `/n` cưỡng bức**: Trước đây hàm in phiếu tự động gửi `/n` ngay sau khi có WHO làm sập cửa sổ MON. Trong v1.7, màn hình MON được giữ nguyên vẹn để người dùng kiểm tra đối chiếu trực tiếp.
- **Tùy chọn In phiếu**: Cấu hình `auto_print_pdf` mặc định là `False`, tránh tự động chiếm quyền màn hình và loại bỏ hoàn toàn nguy cơ xung đột luồng watchdog Win32.

### 3. Tự Động Reload / F8 Lại Khi EWM Sinh Đơn Chậm
- Khi tra cứu bằng Số OB, nếu hệ thống EWM xử lý nền chưa kịp hiển thị Outbound Delivery Order (ODO), tool tự động chờ 2.5s và kích hoạt lại node `N000000010` kèm phím F8 từ 1 - 3 lần.
- Tự động chuyển view sang **Warehouse Order** (`N000000096`) và **Warehouse Task** (`N000000108`) để quét sạch mã WHO (dạng 10 chữ số bắt đầu bằng `60...`).

### 4. Tách Biệt Số Đơn (INWH) và Mã WHO (SAP) — Tương Thích Tool INWH v3.3
- **Bảng kết quả 4 cột chuẩn**: `Số OB` | `Mã Hàng (Material)` | `Số Đơn (INWH)` | `Mã WHO (SAP)`.
- **Cấu hình Số Đơn linh hoạt**: Cho phép người dùng tùy chọn tiền tố số đơn (ví dụ: `10-`, `06-`) và số bắt đầu (ví dụ: `1`). Có nút `✏️ Đổi Số Đơn Nhanh` để đánh lại toàn bộ số đơn trên bảng theo ý muốn.
- **Các nút Copy độc lập**:
  - `📋 Copy Cột OB`: Chỉ chép danh sách số OB.
  - `📋 Copy Số Đơn`: Chỉ chép cột số đơn (INWH).
  - `📋 Copy Mã WHO`: Chỉ chép cột mã WHO để dán vào công cụ đóng số PDF `INWH_v3.3`.
  - `📋 Sao Chép 3 Cột`: Copy 3 cột `Số OB \t Mã Hàng \t Số Đơn` dán file Excel theo dõi.
  - `📋 Copy 4 Cột`: Copy toàn bộ 4 cột.

### 5. Nút "Tiếp Tục Từ Số OB" Độc Lập
- Nút bấm `⚡ TIẾP TỤC TỪ SỐ OB (TRA MON & IN)` cho phép dán trực tiếp danh sách Số OB đã có từ trước (bỏ qua hoàn toàn bước MIGO 311, không lo trừ trùng tồn kho) để lấy mã WHO và gán số đơn tức thì.

### 6. Bảo Vệ Mã Hóa UTF-8 Toàn Diện Chống Crash Windows
- Cấu hình tự động `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` ngăn chặn triệt để lỗi `UnicodeEncodeError` khi in thông báo tiếng Việt trên môi trường Windows Codepage 1252 / 437.

---

# ⚡ SAP ZONE v1.6 Pro (Build 20261108) — TỰ ĐỘNG RA ĐƠN NVL, MIGO 311 & IN PHIẾU WHO

Ứng dụng độc lập Windows Executable (Standalone EXE) siêu nhẹ, tốc độ cao, thay thế hoàn toàn file Excel VBA `260810_RA_DON_NVL_V31_ZONE.xlsm`.

---

## 🚀 Tính Năng & Nâng Cấp Trọng Điểm

### 1. Khắc Phục Triệt Để Lỗi MIGO 311 Dừng/Treo (Chuẩn VBA V31 100%)
Dựa trên kiến trúc chuẩn của file VBA tham chiếu `260810_RA_DON_NVL_V31_ZONE.xlsm` (`MIGO311_RunPost_V24_ErrorGuard.bas`):
- **Tìm kiếm Control đệ quy**: Thay thế toàn bộ đường dẫn subscreen cứng bằng cơ chế quét đệ quy theo thuộc tính `Name` (`RUN_FindChildByNameExact`):
  - `GODYNPRO-ACTION`: Key `"A08"` (Transfer Posting)
  - `GODYNPRO-REFDOC`: Key `"R10"` (Other)
  - `GODEFAULT_TV-BWART`: Text `"311"` (Movement Type)
- **Tự động mở rộng bảng lưới**: Gửi lệnh thu gọn Header & Detail (`MIGO_OK_HEADER_CLOSE`, `MIGO_OK_DETAIL_CLOSE`) khi các panel này che khuất bảng dữ liệu `tblSAPLMIGOTV_GOITEM`.
- **Ánh xạ cột nhập liệu cố định chuẩn VBA (`RUN_GetPreBatchColumnMap`)**:
  - Cột 3: Mã hàng (`Material / Product`)
  - Cột 4: Số lượng (`Qty`)
  - Cột 5: Tên Trạm / Người nhận (`Text / Tram`)
  - Cột 6: Plant (`3093`)
  - Cột 7: Kho xuất (`1005`)
  - Cột 8: Kho nhận (`3001`)
- **Nhận diện cột Batch thật (`RUN_IsRealBatchTitle`)**:
  - Quét tiêu đề các cột, chỉ nhận diện cột Batch khi tiêu đề chứa `BATCH` và **không** chứa `TRFR`, `TRANSFER`, `PSTG` (tránh nhầm với cột Transfer Batch).
  - Ghi Batch ở Nhịp 2 và gửi Enter xác nhận.
- **Thứ tự nút Post & Lưu chứng từ (`RUN_PressPost`)**:
  - Ưu tiên 1: Nút Post trên Toolbar 1 (`wnd[0]/tbar[1]/btn[23]`).
  - Ưu tiên 2: Nút Save trên Toolbar 0 (`wnd[0]/tbar[0]/btn[11]`).
  - Ưu tiên 3: Lệnh Command nội bộ `session.SendCommand("MIGO_POST")`.
- **Vòng lặp xác nhận cảnh báo & Popup (`PostPollLoop`)**:
  - Tự động gửi phím Enter (`sendVKey 0`) để đi qua các cảnh báo loại vàng (`MessageType == 'W'`) hoặc hộp thoại thông báo `wnd[1]`.
  - Trích xuất chính xác số Material Document (10 chữ số bắt đầu bằng 5, 8, 49, 50) hoặc Outbound Delivery trước khi báo hoàn tất.

### 2. Tối Ưu Hóa Khởi Động Siêu Tốc & Giảm 54% Dung Lượng Binary
- Loại bỏ hoàn toàn các gói thư viện thừa (`webview`, `pythonnet`, `clr_loader`, `numpy`, `matplotlib`).
- Dung lượng file chạy `SAP_ZONE.exe` giảm mạnh từ **~30.3 MB xuống còn ~13.8 MB**.
- Tốc độ khởi động mở app: **Tức thì (< 1 giây)**, không còn độ trễ chờ giải nén DLL.
- Bọc an toàn phương thức ghép giao diện `paned.add` chống triệt để lỗi `_tkinter.TclError: unknown option "-weight"`.

### 3. Đồng Bộ Trạng Thái SAP & Quét Tồn /SCWM/MON
- Kết nối SAP chuẩn **`SAP_EMAIL_TEAMS_BRIDGE v1.1.12`** với khởi tạo căn hộ COM (`pythoncom.CoInitialize()`) trong mọi luồng Worker Thread.
- Điều hướng an toàn `/SCWM/MON` qua Favorite Node `F00077` (hoặc fallback `/n/SCWM/MON`), tự động điền kho `VF02`, Storage Type `APFL`, Owner `3093` và truy xuất tồn kho khả dụng FIFO/FEFO.

### 4. Bảng Kết Quả 3 Cột Chuẩn & 1-Click Dán File Theo Dõi
- Hiển thị và xuất kết quả đúng chuẩn: `Số OB | Mã Hàng (Material) | Số Đơn (WHO)`.
- Nút bấm `📋 Sao Chép 3 Cột (Dán Vào File Theo Dõi)` giúp dán ngay lập tức vào Excel bằng `Ctrl + V`.

---

## 📦 Danh Sách Tệp Tải Về

| Tệp tải về | Dung lượng | Mô tả |
| :--- | :--- | :--- |
| **`SAP_ZONE.exe`** | **~13.8 MB** | File chạy trực tiếp, siêu nhẹ, mở tức thì, không cần cài đặt Python |
| **`SAP_ZONE_v1.6_PORTABLE.zip`** | **~13.5 MB** | Gói Portable đầy đủ gồm file EXE + `zone_map.json` + `HUONG_DAN_SU_DUNG.md` + `RELEASE_NOTES.md` |
