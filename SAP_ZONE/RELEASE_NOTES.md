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
