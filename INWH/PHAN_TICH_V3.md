# Phân tích và nâng cấp INWH v3.0

## Hai bản được so sánh

- INWH_TOOL_v2.0.zip: bảng hai cột WH – Số phiếu, thư viện tích hợp trong một HTML.
- SAP_IN_WH_HTML_v2.0.zip: lấy trực tiếp từ Release `v1.0-inwh` của `taikthp357-ctrl/SAP`.
  SHA-256 bản Release: `051224ada0065c5a41994cfb90d9610150dabcbe21003c0c9538a052dc9790f0`, khớp digest do GitHub trả về.

| Hạng mục | INWH_TOOL v2.0 | Bản Release HTML v2.0 | V3.0 |
| --- | --- | --- | --- |
| Cách sử dụng chính | Nhập/dán bảng | Tiền tố và số bắt đầu | Tự động từ PDF, có số bắt đầu |
| Chỉnh số riêng từng WH | Có bảng co giãn | Dán danh sách WH, tự sinh số | Giữ bảng chỉnh/dán như chế độ bổ sung |
| Giao diện | Bảng nhập lớn | Thẻ cấu hình, kéo thả, thống kê | Thẻ cấu hình, kéo thả, khoảng số, thống kê và bảng xem trước |
| Đóng gói | Một HTML | HTML và thư mục lib | Một HTML đã tích hợp thư viện |
| Trang tiếp nối | Đối chiếu Trang x/y và WH trước | Có thể chọn nhầm số bất kỳ trong trang | Đối chiếu Trang x/y, không lấy WT làm WH |
| Lọc trang | Giữ tất cả | Lọc tự động bằng nhận diện vật tư | Giữ tất cả mặc định; tùy chọn bỏ trang chỉ có số trang/ngày giờ |
| CSV | Chưa có | Một CSV đối soát | CSV theo WH và CSV theo từng trang |
| Thứ tự PDF | Giữ nguồn | Gom theo danh sách | Giữ nguồn mặc định; chế độ chỉnh bảng có tùy chọn theo bảng |

## Các lỗi xác định ở bản Release

1. `pageText` được nối bằng dấu cách, sau đó `hasMaterialItems()` lại xử lý theo dòng. Vì dòng chứa cụm `warehouse order` hoặc `người nhận`, cả dòng có vật tư có thể bị bỏ. Khi chạy chính hai hàm nhận diện của bản Release trên nội dung PDF 58 trang được trích bằng PDF.js, cả 58 trang đều bị đánh dấu không có vật tư.
2. `extractWHNumber()` cho phép lấy số 10 chữ số bất kỳ. Ở trang tiếp nối 19 và 23 của mẫu thử, nó nhận các số nhiệm vụ `5047496679` và `5047497302` thành WH.
3. Khi một WH không giữ được trang, mã dùng `st.pages.push(0)`. Số 0 là trang đầu toàn bộ PDF, không phải trang đầu của WH đó. Nhánh này có thể đưa cùng một trang đầu vào nhiều WH và đóng số khác nhau lên nội dung sai.
4. CSV có BOM UTF-8 nhưng nối chuỗi trực tiếp, chưa escape dấu nháy kép của dữ liệu. Nút tải cũng thu hồi Blob URL ngay sau thao tác click, thiếu liên kết tải dự phòng.
5. Gói Release còn phụ thuộc các file trong `lib/` và có nhánh tải thư viện từ CDN. Chỉ lấy riêng HTML chưa bảo đảm chạy ngoại tuyến. Tài liệu trong gói vẫn mô tả luồng Excel/EXE v1.0.

Đây là kết quả kiểm tra mã và chạy logic nhận diện, không phải kết luận dựa trên giao diện hay mô tả của bản Release.

## Thiết kế V3.0

Mặc định: chọn PDF → tự nhận diện WH theo thứ tự xuất hiện → lập số từ tiền tố, số bắt đầu và độ dài tối thiểu → xem khoảng số → tạo PDF → tải PDF/CSV. Thay tiền tố hoặc số bắt đầu sẽ lập lại số trong chế độ tự động. Sang Nhập / chỉnh bảng để đặt số riêng, dán hai cột hoặc thay thứ tự WH.

Giữ số phiếu giống nhau trên mọi bản và trang tiếp nối của cùng WH. Không tự tạo thêm bản. PDF đã có hai bản/WH thì khi in chọn một copy.

Tùy chọn bỏ trang mặc định tắt. Chỉ trang tiếp nối còn duy nhất số trang, ngày và giờ mới là ứng viên. Trang có hình ảnh hoặc thao tác vẽ được phát hiện vẫn được giữ. Không dùng tiền tố mã vật tư hoặc đơn vị tính để quyết định bỏ trang. Trang quét hoặc thiếu căn cứ nhận diện sẽ báo lỗi.

CSV chỉ được tạo từ bản ghi của lần tạo PDF thành công. Chỉnh bảng hoặc thiết lập sẽ vô hiệu hóa kết quả cũ, yêu cầu tạo PDF lại. CSV theo WH có số phiếu, số bản, số trang gốc/giữ/bỏ và số trang kết quả. CSV từng trang lưu đủ mọi trang gốc, kể cả trang bỏ. Dữ liệu được escape theo quy tắc CSV và có BOM UTF-8. Trong Excel, nhập các cột WH/Số phiếu dạng Text để giữ nguyên 06-01. Các giá trị mở đầu bằng ký tự công thức được thêm dấu nháy đơn khi xuất CSV.

## Giới hạn kiểm thử

Kết quả trên file mẫu:

| Trường hợp | WH | Trang đầu ra | Đối soát |
| --- | --- | --- | --- |
| Mặc định giữ toàn bộ | 23 | 58 | Đúng số phiếu trên mọi trang |
| Bật bỏ trang chỉ có số trang/ngày giờ | 23 | 50 | Bỏ đúng 8 trang, giữ 4 trang tiếp nối có vật tư |
| Nhập bảng đảo thứ tự và chọn xuất theo bảng | 23 | 58 | Trang đầu ra khớp thứ tự WH đã chọn |

Hai CSV đều được kiểm tra với số trang đầu ra thực tế. Văn bản nguồn giữ nguyên trên các trang được giữ. Ảnh kết xuất ngoài vùng chân trang trùng với trang gốc tương ứng, kể cả khi đổi thứ tự. Đã kiểm tra bắt đầu 98 sinh tiếp 99, 100; số bắt đầu không hợp lệ bị chặn; CSV xử lý được dấu nháy kép và dấu phẩy trong số phiếu.

Kiểm tra cú pháp trên mã đã đóng gói và chạy kiểm thử tích hợp bằng Node với DOM mô phỏng. Thử bằng PDF 58 trang/23 WH và bảng 10.000 dòng. Đối chiếu nội dung và ảnh kết xuất PDF đầu ra, đồng thời kiểm tra bảng kê CSV với số trang thực tế. Không đưa PDF nghiệp vụ vào gói phát hành.

Môi trường hiện tại không cho mở HTML cục bộ trong trình duyệt thử nghiệm. Vì vậy chưa xác nhận đầy đủ thao tác chuột, kéo thả và tải file trong Edge/Chrome thật. Không coi kiểm thử DOM mô phỏng là kiểm thử trình duyệt hoàn chỉnh.

## Nguồn

- https://github.com/taikthp357-ctrl/SAP/releases/download/v1.0-inwh/SAP_IN_WH_HTML_v2.0.zip
- https://github.com/taikthp357-ctrl/SAP/blob/main/INWH/INWH_v2.0.html
