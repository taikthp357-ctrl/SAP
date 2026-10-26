# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell) - v1.0.69

Công cụ cầu nối tự động giữa **Outlook** $\rightarrow$ **SAP GUI Scripting** $\rightarrow$ **Microsoft Teams (Power Automate)**.

## Điểm nổi bật
1. **Chốt 100% Nghiệp Vụ & Vị Trí Cột Theo File Excel V31 (`260810_RA_DON_NVL_V31_ZONE.xlsm`) (v1.0.69):**
   - **Khóa cố định vị trí các cột chuẩn theo layout của nhà máy (`RUN_GetPreBatchColumnMap`):**
     - `Material (Product)`: Cột 3
     - `Qty (Số lượng)`: Cột 4
     - `Trạm (Text)`: Cột 5
     - `Plant`: Cột 6
     - `Kho xuất (SLoc)`: Cột 7
     - `Kho nhận (SLocTo)`: Cột 8
   - **Dò tìm cột Lô (Batch) chuẩn `RUN_GetBatchColumnForRow`:** Duyệt từ cột 0, ngay khi gặp cột đầu tiên chứa từ khóa `BATCH` (loại trừ `TRFR`, `TRANSFER`, `PSTG`), dừng và thoát ngay lập tức (`Exit Function` / Early Exit), khóa chuẩn cột 9/10, triệt tiêu hoàn toàn lỗi cột Batch bị nhảy sang các cột kỹ thuật phía sau.
   - **Triệt tiêu hoàn toàn lỗi lệch cột `Material=61, Plant=62, KhoXuất=63`:** Khắc phục dứt điểm việc tìm kiếm động theo tiêu đề làm sai lệch sang các cột kỹ thuật ẩn, đảm bảo 100% dữ liệu điền trúng các ô hiển thị trên màn hình MIGO của nhà máy.

2. **Chuẩn Hóa Phân Giải Cột Đơn Lượt & Điền Đầy Đủ Tất Cả Cột MIGO 311 (v1.0.68):**
   - **Triệt tiêu lỗi vòng lặp tham lam (Greedy Loop) ghi đè cột MIGO:** Khắc phục triệt để hiện tượng chỉ điền được mỗi cột `Text` còn các cột `Material`, `Qty`, `Plant`, `SLoc`, `Dest SLoc`, `Batch` bị bỏ trống (do vòng lặp quét qua 84 cột của `tblSAPLMIGOTV_GOITEM` không có cơ chế khóa điểm dừng, khiến các cột phía sau như `PO Qty`, `Receiving Plant`, `Material Group`, `Batch Management - cột 76` ghi đè lên các cột thực tế).
   - **Cơ chế phân giải 2 lượt nghiêm ngặt kết hợp khóa First-Match-Wins:**
     - *Lượt 1 (ABAP Technical Name):* Khớp chính xác tên kỹ thuật ABAP `GOITEM-MATNR` (Cột 3), `GOITEM-ERFMG` (Cột 5), `GOITEM-SGTXT` (Cột 6), `GOITEM-WERKS` (Cột 7), `GOITEM-LGORT` (Cột 8), `GOITEM-UMLGO` (Cột 9), `GOITEM-CHARG` (Cột 10), `GOITEM-TAKE_IT` (Cột 14).
     - *Lượt 2 (Header Title Filter):* Chỉ quét tiêu đề cho các trường còn thiếu với bộ lọc phủ định nghiêm ngặt (loại trừ `TRFR`, `TRANSFER`, `SHORT`, `GROUP`, `PO`, `REC`, `DEST`, `EXT`). Khi đã tìm thấy cột thì khóa ngay (First-Match-Wins), tuyệt đối không để các cột kỹ thuật phía sau ghi đè.
     - *Lượt 3 (VBA Baseline Fallback):* Tự động điều chỉnh khoảng bù dự phòng (default_shift) nếu cột `Mat. Short Text` hiển thị.
   - **Tự động xác nhận dòng (`Item OK` / `cell.Selected = True`):** Tự động tick xác nhận checkbox OK của từng dòng trong bảng MIGO sau khi hoàn tất nhập liệu Nhịp 1 & Nhịp 2 theo đúng chuẩn giao dịch MIGO.
   - **Minh bạch vị trí từng cột trên Console:** Hiển thị rõ ràng vị trí phân giải của từng cột `Material`, `Qty`, `Trạm`, `Plant`, `KhoXuất`, `KhoNhận`, `Lô` trước khi điền, kèm log cảnh báo nếu có ô nhập thất bại.

2. **Khắc Phục Lỗi Cú Pháp T-Code `/n/MIGO` & Minh Bạch Tiến Độ Tạo Đơn MIGO 311 (v1.0.67):**
   - **Triệt tiêu lỗi cú pháp lệnh T-Code `/n/MIGO`:** Nâng cấp hàm `_open_transaction()` để phân biệt chuẩn giữa T-Code namespace (`/SCWM/POST`, `/SCWM/MON` ➔ `/n/SCWM/POST`) và T-Code chuẩn (`MIGO`, `CS03`, `MB52` ➔ `/nMIGO`, `/nCS03`). Khắc phục dứt điểm lỗi SAP báo *"Transaction /MIGO does not exist"* khiến MIGO bị đóng băng không thể mở màn hình.
   - **Minh bạch hóa toàn bộ tiến trình MIGO 311 trên Console:** Loại bỏ hiện tượng im lặng khi chạy MIGO; bổ sung log chi tiết từng nhịp: nạp chế độ `A08 (Transfer Posting)` - `R10 (Other)` - `311 (Movement)`, tìm bảng `tblSAPLMIGOTV_GOITEM`, điền Nhịp 1, điền Nhịp 2 (Batch), bấm Post, và trích xuất số chứng từ từ StatusBar.
   - **Bắt và hiển thị lỗi SAP trực tiếp:** Nếu SAP báo lỗi trên thanh StatusBar (`msg_type in ('E', 'A')`), thông tin lỗi được in ngay ra console và phản hồi minh bạch cho người dùng, không bao giờ nuốt lỗi âm thầm.

2. **Chuẩn Hóa Điền Ô MIGO 311 2 Nhịp & Điều Khiển Chuẩn Xác Theo File VBA V31 (`260810_RA_DON_NVL_V31_ZONE.xlsm`) (v1.0.66):**
   - **Điền ô chính xác bằng `GetCell()` thay vì `modifyCell()`:** Bảng nhập liệu MIGO là `GuiTableControl` nên không hỗ trợ method `modifyCell` (chỉ có trên ALV GridView). Chuyển sang cơ chế chuẩn `tbl.GetCell(row, col).Text = val` kết hợp fallback tìm đối tượng toạ độ `[col, row]` như VBA `RUN_SetGridCellMode`.
   - **Bắt chuẩn ô Movement Type `GODEFAULT_TV-BWART`:** Trực tiếp tìm kiếm control theo tên kỹ thuật `GODEFAULT_TV-BWART` (thay vì tìm nhầm `GODYNPRO-BWART`), đảm bảo MIGO kích hoạt chuẩn xác Movement Type `311`.
   - **Bổ sung cột Trạm Text (`SGTXT` / Cột 5):** Đưa thông tin Trạm (ví dụ `KV08-Thức/Vũ`) vào cột Text dòng chứng từ MIGO theo đúng thiết kế của `RUN_FillPreBatch_ByGrid`.
   - **Tự động đóng Header & Detail để giải phóng bảng Item Overview:** Khi người dùng mở MIGO với Header/Detail mở rộng khiến bảng item bị ẩn, Bot tự động gửi lệnh `MIGO_OK_HEADER_CLOSE` và `MIGO_OK_DETAIL_CLOSE` để bung lưới `tblSAPLMIGOTV_GOITEM`.
   - **Bắt số chứng từ Outbound Delivery bằng Regex toàn diện:** Tương thích cả định dạng `Material document 10xxxxxxxx posted` lẫn `Outbound Delivery 10xxxxxxxx created`.

2. **Triệt Tiêu Số OB Giả Lập Khi Lỗi MIGO & Tự Động Điều Hướng Subscreen MIGO 311 Linh Hoạt (v1.0.65):**
   - **Triệt tiêu hoàn toàn số chứng từ mẫu giả định (`1012924052`):** Loại bỏ triệt để hardcode fallback sinh số OB giả khi MIGO 311 thất bại hoặc bị hủy. Khi không tạo được chứng từ thật trên SAP, `delivery_no` giữ giá trị rỗng `""`, hiển thị `---`, trạng thái gắn cờ minh bạch `⚠️ Chưa tạo OB` và hiển thị huy hiệu `[⚠️ Chưa tạo OB]` trên Teams.
   - **Khắc phục lỗi không tìm thấy bảng nhập liệu `tblSAPLMIGOTV_GOITEM`:** Tự động dò tìm động các điều khiển `GODYNPRO-ACTION`, `GODYNPRO-REFDOC`, `GODYNPRO-BWART` qua `w0.findByName` và quét qua các Subscreen `:0001` đến `:0004` (thay vì cố định `:0003`). Bổ sung độ trễ thích ứng giữa các lần chuyển Key ComboBox để SAP vẽ lại giao diện.
   - **Tự động bung bảng Item Overview nếu bị đóng:** Gửi lệnh `MIGO_OK_ITEMS_OPEN` / `MIGO_OK_ITEMLIST` và quét đệ quy cây điều khiển `GuiTableControl` để tìm bảng nhập liệu ngay cả khi layout MIGO của người dùng đang ở chế độ thu gọn.
   - **Dò tìm cột động theo Metadata (`tbl.Columns`):** Tự động phát hiện vị trí các cột `MATNR`, `ERFMG`/`QTY`, `WERKS`, `LGORT`, `UMLGO`, `CHARG` qua thuộc tính tên cột kỹ thuật hoặc tiêu đề, không phụ thuộc vào thứ tự cột cố định của từng User.
   - **Bảo toàn kiểm tra số chứng từ nghiêm ngặt:** Chỉ công nhận tạo OB thành công khi có Regex số chứng từ 10 số (`10xxxxxxxx`, `49xxxxxxxx`, `500xxxxxxx`) trên StatusBar của SAP.
   - **Tối ưu câu chào Teams khi tạo đơn thất bại:** Khi bốc tồn thành công nhưng chưa tạo được đơn MIGO, Bot phản hồi trung thực: *"Sếp ơi, em đã bốc tồn FIFO nhưng chưa tạo được đơn OB MIGO 311 trên SAP ạ. Chi tiết em để bên dưới nhé."*

2. **Khắc Phục Lệch Dữ Liệu Tồn Kho /SCWM/POST & Tương Thích Đơn Vị Tính EA/PC Ra Đơn NVL (v1.0.64):**
   - **Kế thừa Plant & SLoc nguồn (`source_plant=3093`, `source_sloc=1005`):** Khắc phục lỗi màn hình ALV `/SCWM/POST` không hiển thị cột Plant/SLoc khiến `validated_stock_lot` loại bỏ toàn bộ dòng tồn. Cho phép dòng tồn tự động kế thừa Plant và Kho nguồn đã lọc theo `Owner`.
   - **Tương thích hai chiều giữa ĐVT EA và PC:** Đơn yêu cầu dán từ Teams không ghi ĐVT (mặc định `PC`) tự động khớp và phân bổ với các lô hàng mang đơn vị `EA` trên SAP. Giữ nguyên ĐVT gốc `EA` của lô khi tạo dòng kế hoạch.
   - **Bắt chuẩn cột loại tồn nguồn (`TST`):** Ưu tiên cột nguồn `TST` (Source Stock Type = `FT`) thay vì bị nhầm sang cột đích `TO_CAT`, đồng thời dự phòng `cat="FT"` khi layout ẩn cột.
   - **Giải phóng nghẽn phân bổ:** Các dòng tồn trích xuất từ `/SCWM/POST` (ví dụ 170 cái tại Bin `APFL-P23-01-12` và 297 cái tại Bin `APFL-P23-01-14`) được nạp và trừ lùi FIFO chính xác 100%, không còn hiện tượng báo sai "Thiếu tồn (0/100)".

2. **Tối Ưu Hiển Thị Bảng Teams & Xử Lý Đính Kèm MIGO Đa Tab Toàn Diện (v1.0.63):**
   - **Bảo toàn 100% dòng hiển thị không bị cắt ngắn (Full 42+ Rows):** Tối ưu hóa dung lượng Inline CSS của từng ô `<td>` và ưu tiên loại bỏ mảng JSON thô trước khi cắt giảm dòng. Đảm bảo toàn bộ 42-100 chứng từ hiển thị trọn vẹn trên Teams mà không bao giờ vượt trần 26 KB.
   - **Nhận diện số phiếu linh hoạt trong HeadText MIGO (`HEADTEXT_OK`):** So khớp thông minh dạng tập con/chuỗi con. Khi SAP đã có sẵn `Thành Công IN-6058` và người dùng gửi số phiếu `6058`, hệ thống tự động nhận diện chính xác là `HEADTEXT_OK` thay vì đánh nhãn `HEADTEXT_REVIEW`.
   - **Cách ly luồng tuyệt đối (`threading.local()`):** Khắc phục triệt để hiện tượng xung đột đa luồng khi chạy song song Multi-Tab (Tab 2 reset chứng từ làm mất trạng thái của Tab 1 dẫn đến `HEADTEXT_BLOCKED` / `DOCUMENT_NOT_VERIFIED`).
   - **Tự động phục hồi kết nối Outlook COM:** Ngăn ngừa ngắt tiến độ khi Outlook ngắt kết nối sau các lô chạy SAP kéo dài.
   - **Điều hướng Multi-Tab chuẩn xác:** Tự động sử dụng Tab 4 & Tab 5 (`index >= 3`) cho Batch Pool, bảo toàn tuyệt đối Tab 1..3 cho nhân sự kho thao tác.
   - **Chuẩn hóa thông điệp phản hồi:** Khi hoàn thành phần lớn chứng từ (ví dụ 41/42), bot gửi lời chào chúc mừng với trạng thái `PARTIAL`, chỉ báo lỗi khi toàn bộ lô thất bại.

2. **Phân Tách Ranh Giới Phiên User & Bot Thông Minh (Multi-Tab Session Isolation - v1.0.62):**
   - **Triệt tiêu dứt điểm lỗi "Tất cả các phiên SAP đều đang bận thao tác của người dùng":** Khắc phục lỗi khi nhân sự kho đang mở 3 tab đầu (Session 1, 2, 3 / `index 0, 1, 2`) để làm việc khiến Bot bị chặn không thể thực thi yêu cầu mới (như lệnh `SAP att mail` trong ảnh).
   - **Bảo lưu tuyệt đối 3 tab đầu cho User (`USER_RESERVED_SESSIONS = 3`):** Bot không bao giờ can thiệp hay cướp màn hình làm việc của nhân viên kho khi đang mở dở chứng từ.
   - **Chuyển toàn bộ tác vụ của Bot sang dải 3 tab tiếp theo (`index >= 3` / Tab 4, 5, 6):** Khai thác tối đa trần 6 phiên của SAP GUI (`MAX_SAP_SESSIONS = 6`). Khi người dùng mở 1, 2 hoặc 3 tab đầu, Bot tự động gọi `CreateSession()` mở Tab thứ 4 (`index = 3`) dành riêng cho Bot chạy nền.
   - **Tái sử dụng tab Bot thông minh:** Các lần gọi tiếp theo sẽ tái sử dụng ngay Tab thứ 4 mà không mở thêm tab thừa hay làm rung giật giao diện của User.

2. **Xử lý Bảng Đơn Hàng NVL Đa Dòng Dán Trực Tiếp Từ Excel / Teams & Tạo Đơn OB 10xxxxxxxx (v1.0.61):**
   - **Bóc tách bảng dữ liệu tự động (`parse_nvl_table_rows`):** Tự động nhận diện cấu trúc 3 cột dán trực tiếp từ Excel (`Mã NVL | Số lượng | Trạm`, ví dụ `BIW75111244AA | 100 | KV08-Thức/Vũ`) và cả biến thể 4 cột có `STT/OB` đứng đầu (`1 | BIW... | 100 | ...`).
   - **Khắc phục triệt để nhiễm bẩn mã hàng (Material Token Pollution):** Cô lập triệt để cột Trạm (`KV08-Thức/Vũ`), ngăn không cho regex bóc tách token tự nhiên biến chuỗi trạm thành mã vật tư giả (`KV08-TH`).
   - **Gom cụm OB & Cộng dồn Lô/Bin Tự Động (`group_into_ob_groups`):** Khi nhiều dòng đơn hàng cùng mã và trạm được phân bổ từ cùng một Lô (Batch) và Vị trí Bin, hệ thống tự động cộng dồn số lượng thành 1 dòng duy nhất (ví dụ: 3 dòng 100 cái từ Lô `240801` $\rightarrow$ 1 dòng 300 cái). Đánh số cụm OB chuẩn rút gọn 3 ký tự (`OB1`, `OB2`, ...).
   - **Chuẩn hóa số chứng từ Outbound Delivery `10xxxxxxxx`:** Bắt và gán số chứng từ Outbound thật từ SAP MIGO 311 (ví dụ: `1012924052`) vào toàn bộ các dòng của cụm OB.
   - **Huy hiệu KPI Badges & Bảng rút gọn trên Teams:** Hiển thị thẻ tóm tắt trực quan `[✅ Số OB: 1012924052]` `[🏭 Cụm: OB1 (2 dòng)]` `[📦 Tổng cấp: 400 PC]` kèm bảng dữ liệu với các cột rút gọn: `OB`, `Số OB`, `Mã NVL`, `SL`, `Lô`, `Bin`, `Trạm`, `TT`.
   - **Thực thi MIGO 311 2 Nhịp An Toàn (`post_migo_311_ob_group`):** Điền dữ liệu 2 nhịp (Nhịp 1: Mã, SL, Plant 3093, Kho xuất 1005, Kho nhận 3001; Nhịp 2: tìm động cột BATCH điền số lô; Post và bắt số chứng từ qua Regex trên StatusBar).

2. **Đính kèm Lũy Kế Đa Hóa Đơn (Multi-VAT & Incremental Upload) & Lọc Trùng Lặp Thông Minh (v1.0.60):**
   - **Xử lý trọn vẹn nghiệp vụ "Att thêm 1 VAT nữa":** Loại bỏ hoàn toàn cơ chế kiểm tra nhị phân cũ (`if sap_vat: skip`). Khi một chứng từ DOC đã có sẵn 1 file VAT trên SAP GOS (như `VAT_5122`), hệ thống áp dụng thuật toán *Claim-based Matching* để so khớp từng file local với từng file trên SAP GOS. Các file VAT mới hoặc hóa đơn bổ sung (như `VAT 5122(1).pdf`, `VAT 1213.pdf`, hoặc file VAT của mã hàng thứ hai trong DOC) sẽ **được tự động đính thêm vào ngay** mà không bị bỏ qua.
   - **Lọc trùng lặp file vật lý bằng Content Hash (SHA-256):** Tự động phát hiện và loại trừ các bản sao trùng lặp 100% nội dung (như người dùng vô tình copy hoặc đổi tên giữa `VAT_5122.pdf` và `VAT 5122.pdf`) trước khi tải lên, ngăn ngừa triệt để rác cơ sở dữ liệu SAP GOS. Đồng thời, giữ lại và tải lên đầy đủ khi hai file có nội dung hoặc dung lượng khác nhau.
   - **Nhận diện toàn diện các biến thể tên file theo số phiếu:** Bắt chuẩn xác mọi cách đặt tên file theo số phiếu HeadText của nhân sự kho: `VAT_5122.pdf`, `VAT 5122.pdf`, `VAT 5122(1).pdf`, `VAT_5122(1).pdf`, `VAT 5122_1.pdf`, `VAT_5122_1234.pdf`, v.v.
   - **Báo cáo chi tiết & minh bạch trên Teams:** Phản hồi rõ ràng số lượng file đã có và số lượng file được đính thêm (ví dụ: `DONE (Đính thêm 1, đã có 1)`).

2. **Triệt tiêu lỗi `ATTACHMENT_UNCERTAIN` & Khớp linh hoạt Tiêu đề SAP GOS (v1.0.59):**
   - **Khắc phục chênh lệch đuôi mở rộng giữa Local và SAP GOS:** Khi tải lên file `VAT 5319.pdf`, SAP GOS lưu tiêu đề vào cột Title dưới dạng `VAT 5319` (không kèm đuôi `.pdf`, loại tệp được lưu riêng ở cột `Typ`). Cơ chế kiểm tra sau đính kèm (Readback Verification) trong [`sap_attachment.py`](file:///G:/SAP/sap-email-teams-bridge/app/sap_attachment.py) được nâng cấp so sánh đa tầng (khớp tên file, khớp tên gốc `stem`, khớp chuỗi chuẩn hóa `normalize_text` và công nhận tín hiệu thành công từ StatusBar `msg_type == "S"`).
   - **Chấm dứt hoàn toàn báo lỗi sai `PARTIAL (Thiếu VAT)`:** Loại bỏ triệt để cảnh báo `ATTACHMENT_UNCERTAIN: Chưa đọc lại đúng tên file trên SAP`, đảm bảo kết quả phản hồi trên Teams hiển thị chính xác trạng thái `DONE` ngay sau khi đính kèm.

2. **Triệt tiêu Lỗi COM Dead Pointer `-2147417851` & Batch ZFIR266 Siêu Tốc (v1.0.58):**
   - **Batch Multiple Selection ZFIR266:** Khi tra cứu danh sách nhiều chứng từ (ví dụ 15 DOC), hệ thống đưa đồng loạt toàn bộ danh sách vào nút Multiple Selection (`btn%_S_MBELN_%_APP_%-VALU_PUSH`) và chỉ bấm F8 thực thi đúng 1 lượt duy nhất. Toàn bộ kết quả của 15 DOC được trích xuất từ ALV Grid trong ~1.5 giây (thay vì lặp 45 giây).
   - **Triệt tiêu lỗi con trỏ chết (Stale Object / Dead Pointer COM):** Khắc phục dứt điểm lỗi `(-2147417851, 'The server threw an exception.')` do dùng lại con trỏ COM cũ qua các lần đổi màn hình Dynpro. Toàn bộ các đối tượng điều khiển GUI được phân giải tươi mới (fresh resolution) trong từng vòng lặp và tự động phục hồi về `/nZFIR266` nếu có gián đoạn.
   - **Khóa chặt ranh giới MIGO Fallback:** MIGO Display chỉ được kích hoạt làm dự phòng khi giao dịch `ZFIR266` thực sự chưa được cài đặt trên hệ thống SAP (`not zfir_supported`). Khi ZFIR266 đã chạy thành công, các chứng từ không có dữ liệu sẽ kết luận ngay là chưa có file, tuyệt đối không nhảy sang MIGO Display làm giật màn hình và gây hiểu nhầm sang nghiệp vụ "att mail".
   - **Bảo vệ bóc tách Intake:** Loại bỏ từ đơn `att` khỏi Regex kích hoạt thư mục upload, tự động xóa sạch `folderPath = ""` khi nhận diện lệnh tra cứu `CHECK_ATT`.

2. **Tra cứu Mẫu đại diện (Wildcard `BIW*6100*`) & Ổn định ZFIR266 ALV (v1.0.57):**
   - **Tự động bung mẫu Wildcard từ Master Data:** Mở rộng bộ phân giải [`material_resolver.py`](file:///G:/SAP/sap-email-teams-bridge/app/material_resolver.py) với `fnmatch`. Khi người dùng gõ `BIW*6100*` hoặc `*6100*`, hệ thống tự động bung chính xác 5 mã SAP tương ứng (`BIW32016100AA`, `BIW32026100AB`, `BIW75126100AB`, `BIW75116100AA`, `BIW32216100AA`) từ `DATA_MASTER.xlsx` mà không cần chỉnh sửa bất kỳ dòng nào trong file Excel.
   - **Bảo toàn ký tự `*` và `?` trong Parser:** Tinh chỉnh Regex bóc tách token trong [`email_parser.py`](file:///G:/SAP/sap-email-teams-bridge/app/email_parser.py) (cả câu lệnh tự nhiên và cấu trúc JSON `[\w\-./*?]{2,40}`), giữ nguyên dấu `*` ở giữa hoặc hai đầu mã vật tư mà không bị cắt vụn thành các từ ngắn.
   - **Triệt tiêu lỗi `_return_to_main_menu` trên ZFIR266:** Bổ sung alias chuẩn hóa `_return_to_main_menu` $\rightarrow$ `_reset_to_session_manager` trên [`sap_engine.py`](file:///G:/SAP/sap-email-teams-bridge/app/sap_engine.py), khắc phục hoàn toàn sự cố văng lỗi ngoại lệ sau khi đọc báo cáo `ZFIR266`.
   - **Trích xuất thông minh ALV Grid ZFIR266:** Tự động đối chiếu cột tên file qua `ColumnOrder`, `Caption` và quét nội dung nhận dạng định dạng file (`.pdf`, `.xlsx`, v.v.), đảm bảo bắt trọn vẹn các file chứng từ như `VAT_6489.PDF`.

2. **Định dạng số chuẩn hóa EU/VN & Pipeline Hardening (v1.0.56):**
   - **Triệt tiêu lỗi `Ambiguous quantity`:** Tự động phân giải chính xác chuỗi số định dạng EU/VN có nhiều dấu chấm (ví dụ: `1.421.410`, `2.842.820`) và số nguyên hàng nghìn (`1.000`, `2.500`) mà không làm suy hao hoặc biến số lượng về `0.0`. Format hiển thị số lượng fallback an toàn sang chuẩn EU không cắt cụt chuỗi.
   - **Phân giải nhận diện Xưởng Sơn EDC (Plant 4007/4004):** Bổ sung nhận diện từ khóa `EDC`, `SƠN EDC`, `SON EDC`, `XUONG SON`, ánh xạ chuẩn xác bộ Plant `["4007", "3093"]`. Khóa chặt tìm kiếm gần đúng (Fuzzy match) với độ dài tối thiểu `>= 5` ký tự, tránh khớp nhầm các mã ngắn.
   - **Bảo vệ phiên & Fallback toàn diện cho `ZFIR266`:** Bảo vệ điều hướng phím `F3` (chỉ gửi Back khi SAP đã mở ALV Grid), ngăn chặn dứt điểm hiện tượng thoát nhầm về `SESSION_MANAGER` làm hỏng các chứng từ kế tiếp. Tích hợp cuộn trang `firstVisibleRow` quét trọn vẹn danh mục lớn và tự động fallback sang MIGO Display kiểm tra GOS Attachment cho từng chứng từ không trả về dòng nào trên ZFIR266.
   - **Tương thích hoàn hảo OB Service với dữ liệu tồn kho MB52:** Nới lỏng kiểm tra FIFO khi `require_fifo=False`, tự động gán `bin_no="N/A"`, `cat="UNRESTRICTED"` để xử lý mượt mà nguồn tồn kho từ MB52.
   - **Cảnh báo trực quan `REVIEW_REQUIRED`:** Định dạng rõ ràng giá trị `"N/A"` thay vì chuỗi trống hoặc số `0.00` trong file Excel, tự động tô màu nền đỏ cảnh báo trong cả file Excel và thẻ HTML gửi qua Teams.
   - **Loại trừ kích hoạt nhầm SIMULATED:** Tinh chỉnh Regex `dryRun` chỉ nhận diện các lệnh thử nghiệm tường minh (`xem truoc`, `dry run`, `chay thu`, `simulate`), loại bỏ từ đơn `test` để tránh kích hoạt nhầm do chữ ký email hoặc tên nhóm Teams.

2. **Nhận diện linh hoạt Plant & Sửa lỗi gõ thiếu "plan" (v1.0.55):**
   - **Chống nhận diện nhầm về Plant mặc định 3093:** Mở rộng Regex bóc tách Plant trong [`email_parser.py`](file:///G:/SAP/sap-email-teams-bridge/app/email_parser.py) để nhận diện các biến thể gõ nhanh của nhân viên kho (`plan`, `plat`, `pl`, `kho`, `nha may`) hoặc đứng độc lập một mã Plant 4 số thuộc whitelist (`3093`, `4007`, `4004`, `2001`, `3001`...).
   - **Xử lý sạch sẽ câu lệnh:** Khi người dùng gửi *"Sap tra tồn BIW32022638AB plan 4007"*, hệ thống bóc tách chính xác `plant = "4007"`, `explicitPlant = True` và loại trừ hoàn toàn các từ khóa khỏi mã hàng.

3. **Bảo toàn dòng Hàng In-Transit trên MB52 & Thẻ KPI phân tách (v1.0.55):**
   - **Nhận diện đầy đủ cột `Transit/Transf.`:** Bổ sung ánh xạ mã cột SAP `UMLMC`, `UMLMD`, `TRAME`, `TRANSIT/TRANSF.` vào [`sap_engine.py`](file:///G:/SAP/sap-email-teams-bridge/app/sap_engine.py).
   - **Không lọc bỏ dòng Transit:** Xử lý triệt để trường hợp hàng đang chuyển kho ở cấp độ Plant (chưa nhập vào SLoc cụ thể nên `SLoc == ""` và `Batch == ""`): Tự động tạo dòng đặc biệt `SLoc = "TRANSIT"`, `Batch = "Chờ xuất (từ 3093)"` (với Plant 4007/4004), forward-fill mã hàng từ dòng đầu.
   - **Thẻ KPI Teams tách bạch rõ ràng:** Header kết quả trên Teams hiển thị tách biệt Tồn khả dụng sẵn sàng xuất và Hàng đang trung chuyển:
     `✅ Tồn khả dụng: 140 EA` | `🚚 Chờ nhập (Transit từ 3093): 78 EA` | `📦 SLoc 4004: 4 EA` | `📦 SLoc 4007: 136 EA` | `🏷️ 2 Lô (Batch)`.

4. **Thiết kế bảo mật & an toàn (Zero Inbound Port Architecture):**
   - **0 Open Port:** Không mở cổng mạng lắng nghe (không socket server), chỉ gửi Outbound HTTPS POST duy nhất tới Power Automate Webhook.
   - **0 PowerShell:** Điều khiển SAP GUI và Outlook hoàn toàn qua Windows COM Object trong RAM, không gọi bất kỳ tiến trình PowerShell hay subprocess nhạy cảm nào.
   - **Bảo vệ phiên người dùng (Strict Session Ownership Guard):** Khi bật chế độ an toàn, hệ thống chỉ sử dụng phiên rảnh (Idle), tự động xin mở phiên mới nếu cho phép, tuyệt đối không can thiệp hoặc cướp màn hình làm việc của người dùng đang thao tác.
   - **Quét sạch bộ lọc cũ (Clean Selection Screen):** Tự động xóa sạch toàn bộ các ID/ô lọc còn sót lại từ phiên trước (MBLNR, BWART, SLoc, Batch...) trước khi chạy MB51, MB52, VL06O, tránh bị lọc sai khiến kết quả trả về rỗng.

5. **Cơ chế Keep-Alive & Auto-Reconnect COM (v1.0.54):**
   - Tự động làm mới phiên SAP mỗi 15 phút không thao tác, chống bị văng timeout do cài đặt của SAP NetWeaver.
   - Tự động kết nối lại nếu phiên bị rớt mạng hoặc người dùng vô tình đóng SAP.
   - **Tự phục hồi kết nối Outlook COM (-2147220995):** Tự động bắt lỗi RPC disconnect (`Object is not connected to server`) trong vòng lặp chính, re-dispatch lại Outlook MAPI namespace và lấy lại con trỏ thư mục mà không cần khởi động lại toàn bộ bridge.

6. **Chống Redraw Crash & Giữ MIGO CHANGE Xuyên Suốt Mẻ (v1.0.54):**
   - **Triệt tiêu lỗi COM `-2147417851` (`The server threw an exception`):** Thiết kế hàm `ensure_migo_mode` kiểm tra trạng thái ComboBox `GODYNPRO-ACTION` và `GODYNPRO-REFDOC`. Nếu đã ở `A12` hoặc `R02`, tuyệt đối không gán đè. Khi cần đổi Action, chèn khoảng nghỉ 0.5s để SAP GUI hoàn tất Screen Redraw trước khi re-fetch lại đối tượng RefDoc mới từ DOM, loại bỏ toàn bộ lỗi Dead Pointer COM.
   - **Loại bỏ Reload DISPLAY thừa:** Xóa bỏ lệnh nạp lại `mode="DISPLAY"` ở cuối hàm `handle_migo_headtext`. Giữ phiên MIGO luôn ở `A12 CHANGE` trong suốt 45 chứng từ của mẻ, nâng tốc độ xử lý lên 0.3s/chứng từ và loại trừ 100% việc lặp lại thao tác chuyển mode.

7. **Chống Lặp Vô Tận Recovery Loop & Tự Động Giải Phóng Thư Kẹt (v1.0.53):**
   - **Triệt tiêu vòng lặp vô tận `[RECOVERY] REF-...`:** Khi phát hiện thư cũ đã lưu kết quả hoặc lần chạy trước bị gián đoạn, hệ thống tự động đánh dấu đã đọc (`UnRead=False`) và chuyển thư ra khỏi hòm thư `SAP_BOT` sang thư mục con `Processed` (nếu đã có kết quả) hoặc `Failed` (nếu cần đối soát).
   - Ngăn chặn hoàn toàn hiện tượng thư cũ bị kẹt lại trong `SAP_BOT` khiến bot quét đi quét lại và in thông báo liên tục mỗi chu kỳ thăm dò.

8. **Chuyển chế độ MIGO Tự động & Khôi phục Multi-Tab Engine (v1.0.52):**
   - **Tự động chuyển từ DISPLAY sang CHANGE:** Khi thực hiện điền số phiếu HeadText hoặc đính kèm file, nếu phiên MIGO đang ở chế độ xem `A04 DISPLAY`, hệ thống tự động phát hiện ComboBox `GODYNPRO-ACTION != "A12"` và chuyển sang `A12 CHANGE`, đồng thời đảm bảo `GODYNPRO-REFDOC == "R02"`, giải quyết dứt điểm tình trạng bị khóa trường không cho lưu.
   - **Khôi phục cơ chế Chạy Đa Tab (Multi-Tab Engine):** Mở rộng khả năng xử lý song song với cấu hình `enable_multi_tab=True` và `max_batch_tabs=2`. Khi danh sách chứng từ lớn (> 5), hệ thống tự động mở đủ 3 Tabs: Tab 0 dành riêng cho Fast Lane (tra cứu tức thời), Tab 1 và Tab 2 dùng cho Batch Pool xử lý song song. Xác thực phiên được cải tiến dựa trên System, Client và User để các Tab con không bị chặn nhầm.

9. **Kích hoạt Chạy Thật ATT_MIGO & Nghiệp vụ CHECK_ATT (v1.0.51):**
   - **Chạy Thật MIGO (Live Execution Mode):** Triệt tiêu hoàn toàn lỗi chạy giả lập/mô phỏng (`SIMULATED`) trên các lệnh ATT MIGO thật; phân định rõ cờ `dryRun=False` (`WRITE_LIVE`) và `dryRun=True` (`WRITE_PREVIEW`); tự động nạp năm hiện tại (2026) khi người dùng không truyền năm, giải quyết dứt điểm lỗi `YEAR_REQUIRED`.
   - **Tra cứu Đính kèm theo DOC (CHECK_ATT / CHECK_ATT_DOC):** Dựa trên `PLAN BỔ SUNG CHECK ATT THEO DOC — BBBG VÀ VAT` (26/10/2026).
     - Ưu tiên gọi Transaction chuyên dụng `ZFIR266` (Company Code `P_BUKRS=2301`, Năm `P_GJAHR`, DOC `S_MBELN-LOW`, Max `P_MAX=500`).
     - Tự động Fallback sang MIGO A04 Display mode (`OK_REF_ATTA` / GOS Attachment list) nếu hệ thống chưa cài ZFIR266.
     - Phân loại rõ ràng 5 trạng thái đính kèm kèm huy hiệu trực quan:
       + Đã có cả BBBG và VAT (`FULL_ATT` - Huy hiệu Xanh lá)
       + Đã có BBBG, thiếu VAT (`MISSING_VAT` - Huy hiệu Vàng)
       + Đã có VAT, thiếu BBBG (`MISSING_BBBG` - Huy hiệu Vàng)
       + Chưa có cả BBBG lẫn VAT (`NO_ATT` - Huy hiệu Đỏ)
       + Có file nhưng chưa rõ loại (`REVIEW` - Huy hiệu Xanh dương)
   - **Ổn định tương tác Live SAP GUI (v1.0.50):** Prefill Selection Screen `VL06O`, `ME2N` (ALV + WE101 + Plant 3093), `VL06I` (Shift+F2 xóa ngày + Item View), chuẩn hóa nhãn `📦 SLoc {sloc}`, cơ chế Combo Check (VITRI $\rightarrow$ MB52 Fallback) và Easter Egg tác giả Phạm Tiến Tài (`WHO_IS_TAI`).

## Cấu trúc thư mục mã nguồn
```text
sap-email-teams-bridge/
├── app/
│   ├── adaptive_cache.py           # Bộ nhớ đệm UI thích nghi tự động
│   ├── bbbg_service.py             # Đối soát & tạo biên bản giao nhận FG
│   ├── config.py                   # Bộ tải cấu hình & tự sửa JSON
│   ├── email_parser.py             # Parser thông minh & Weighted Intent Matrix (13 intents)
│   ├── extended_queries.py         # Truy vấn mở rộng BOM, Queue EWM, ASN
│   ├── material_resolver.py        # Phân giải mã 8 số sang mã SAP chuẩn
│   ├── nvl_service.py              # Nghiệp vụ cấp phát NVL & mô phỏng MIGO 311
│   ├── ob_service.py               # Nghiệp vụ điều phối Outbound & phân bổ tồn kho
│   ├── runtime_control.py          # Kiểm soát an toàn phiên & ranh giới chạy
│   ├── sap_attachment.py           # Quản lý tải và đính kèm GOS / MIGO header text
│   ├── sap_engine.py               # Điều khiển SAP GUI COM, ZFIR266 & Live SAP GUI
│   ├── state_manager.py            # Quản lý hàng đợi Outbox & chống trùng bằng SQLite
│   └── teams_dispatcher.py         # Gửi kết quả Webhook kèm bảng HTML & KPI
├── data/
│   └── bridge_state.db             # Cơ sở dữ liệu SQLite lưu Outbox & Processed IDs
├── templates/
│   ├── power_automate_schema.json  # Schema mẫu dán vào Power Automate
│   └── email_samples.txt           # Mẫu nội dung email để test
├── docs/
│   ├── HUONG_DAN_SU_DUNG.md        # Hướng dẫn cài đặt & vận hành chi tiết
│   └── DANH_GIA_v1.0.48_20261025.md # Báo cáo đánh giá và khắc phục lỗi vận hành
└── tests/                          # Bộ unit tests hồi quy tự động (244 tests pass 100%)
```

## Chạy thử nghiệm
1. Đảm bảo đã mở SAP GUI và Outlook Classic (có thư mục `Inbox/SAP_BOT`).
2. Chạy file `START_BRIDGE.cmd`.
3. Gửi email mẫu:
   ```text
   LENH: VITRI | MA: <Mã_hàng> | BATCH: <Lô>
   hoặc: Check att 5007145690 5007145691
   hoặc: SAP att mail 5007145690
   ```
