# INWH — Đánh số và lọc trang in Warehouse Order

Công cụ xử lý PDF phiếu lệnh kho xuất từ SAP, phục vụ in hàng loạt.

## Chức năng

- Nhận diện Warehouse Order (WH) từ nội dung PDF.
- Đánh số phiếu theo thứ tự như `06-01`, `06-02` ...; hai bản của cùng WH có cùng số.
- Bỏ trang chỉ có ngày giờ/chữ ký mà **không có dòng vật tư**, nhưng giữ trang nối tiếp vẫn có vật tư.
- Giữ nguyên barcode và thông tin SAP; xuất file PDF để in cùng bảng đối chiếu CSV.

## Hướng dẫn

1. Xuất các phiếu WH thành PDF từ SAP.
2. Dùng công cụ `WH_Tool_Danh_So_PDF_v2_LOC_TRANG_TRONG.zip` (bản v2).
3. Chạy `CHAY_TOOL_WH.bat` trên máy có Python 3 và PyMuPDF.
4. Chọn PDF đầu vào, mã ngày và STT bắt đầu; lưu PDF kết quả.
5. PDF gốc đã có hai bản/WH thì chọn **1 copy** khi gửi máy in.

**Bảo mật:** Repository này hiện là public. Các PDF mẫu từ SAP có số chứng từ, mã vật tư và thông tin nghiệp vụ; cân nhắc chính sách công ty trước khi đưa tài liệu thực tế lên GitHub. Không đưa các chứng từ thật vào đây khi chưa được phép công khai.

## Tài liệu mẫu được xử lý trước đây

- `WH_06_DA_LOC_TRANG_TRONG.pdf`: 20 trang từ 24 trang, 9 WH, 2 bản mỗi WH.
- `WH_06_DA_LOC_TRANG_TRONG_DANH_SACH.csv`: bảng đối chiếu.

Tên file được mô tả để tra cứu; không có nghĩa tất cả file nhị phân đã được đưa lên repository.
