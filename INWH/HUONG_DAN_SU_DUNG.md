# HƯỚNG DẪN VẬN HÀNH BỘ CÔNG CỤ IN PHIẾU KHO SAP (v1.0)

Tài liệu này dành cho nhân viên điều độ kho, quản lý đơn hàng và nhân sự vận hành in ấn phiếu Warehouse Order (WH).

---

### BƯỚC 1: CHUẨN BỊ THƯ MỤC LÀM VIỆC
1. Tải file nén `SAP_IN_WH_PORTABLE_v1.0.zip` về máy tính.
2. Giải nén vào một thư mục làm việc cố định trên ổ đĩa (ví dụ: `D:\IN_PHIEU_WH\`).
3. Đảm bảo 2 file **`IN_PHIEU_WH.xlsm`** và **`wh_engine.exe`** nằm **chung trong một thư mục**.

---

### BƯỚC 2: THAO TÁC TRÊN EXCEL
1. Mở file **`IN_PHIEU_WH.xlsm`**. Nếu phía trên bảng tính có dải màu vàng cảnh báo bảo mật, nhấn **Enable Content** (hoặc **Enable Macros**).
2. **Khai báo tiền tố ngày:**
   - Tại ô **C4**, tiền tố mặc định lấy theo ngày hiện tại kèm dấu gạch ngang (ví dụ: `06-`). Bạn có thể đổi sang ngày khác tùy ý (ví dụ: `07-`).
3. **Chọn file PDF gốc từ SAP:**
   - Nhấn nút **"Chọn File PDF..."** và duyệt tìm file PDF vừa xuất từ SAP (ví dụ: `3.pdf`).
4. **Nhập danh sách mã WH:**
   - Copy danh sách các số WH cần in từ file Excel kế hoạch/điều độ (ví dụ: `6028720086`, `6028720696`...).
   - Dán (Paste) vào **Cột C (Mã Warehouse Order)** bắt đầu từ dòng 10 trở đi.
5. **Gán số đơn cần in:**
   - Cách 1: Nhấn nút **"Tự điền số đơn nhanh"** $\rightarrow$ Hệ thống tự động điền `06-01`, `06-02`, ... `06-20` vào cột D theo đúng thứ tự bạn vừa dán.
   - Cách 2: Bạn có thể nhập tay hoặc sửa lại bất kỳ ô nào ở cột D theo quy ước riêng của ca làm việc (ví dụ: `06-01A`, `06-05_GAP`...).

---

### BƯỚC 3: XỬ LÝ VÀ ĐỐI SOÁT
1. Nhấn nút **"▶ 1. XỬ LÝ & ĐỐI SOÁT PDF"**.
2. Hệ thống sẽ xử lý ngầm trong 1 đến 3 giây và tự động cập nhật bảng kết quả:
   - **Trang gốc:** Số trang của mã WH đó có trong file PDF gốc.
   - **Trang giữ:** Số trang có mã vật tư được giữ lại để in.
   - **Trang bỏ:** Số trang thừa (chỉ có chữ ký/ngày giờ) đã được tự động loại bỏ.
   - **Trạng thái đối soát:**
     - 🟩 **Đã khớp & đóng số:** Đơn hàng bình thường, đã được đóng dấu số đơn.
     - 🟥 **❌ THIẾU:** Mã WH có trong danh sách Excel nhưng trong file PDF của SAP lại không tìm thấy $\rightarrow$ Cần kiểm tra lại trên SAP xem đơn này đã được xuất lệnh hay chưa.
     - 🟨 **⚠️ THỪA:** Có trong PDF nhưng chưa được khai báo số trong Excel $\rightarrow$ Đơn phát sinh thêm.

---

### BƯỚC 4: IN ẤN PHIẾU
1. Nhấn nút **"🖨 2. MỞ PDF KẾT QUẢ ĐỂ IN"**.
2. File PDF kết quả (đã được lọc trang rác, sắp xếp theo thứ tự danh sách Excel và đóng dấu số ở góc dưới bên phải) sẽ tự động mở lên.
3. Trong hộp thoại máy in (Print):
   - **Số lượng bản in (Copies): Chọn là `1`** (Vì mỗi WH trong file PDF của SAP đã có sẵn 2 bản in).
   - Tiến hành in ra máy in kho thông thường.
