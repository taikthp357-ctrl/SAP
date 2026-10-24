# Kế hoạch nâng cấp SAP Email Teams Bridge sau v1.0.15

Ngày lập: 24/10/2026. Trạng thái: kế hoạch đề xuất, chưa triển khai hoặc kích hoạt.

## 1. Mục tiêu và phạm vi

Ưu tiên kết quả đúng và không mất yêu cầu trước khi giảm dung lượng. Giữ một worker Python, master Excel, SQLite và log có giới hạn; tiếp tục dùng Outlook Classic, SAP GUI và Power Automate hiện có.

- Giữ các nghiệp vụ: vị trí, tồn tổng, tồn theo kho, tồn thấp đa plant, lịch sử, IB, STO, OB và lịch sử bin. Không bỏ lệnh đang dùng.
- Giữ các khóa payload Power Automate đang sử dụng: `requestId`, `messageId`, `command`, `status`, `senderName`, `summary`, `table_html`, `completedAt`, `totalRecords`, `displayedRecords`, `records`.
- Giữ cách xưng hô “Sếp ơi…” đã thống nhất. Phân biệt có dữ liệu, không có dữ liệu, dữ liệu chưa đầy đủ và truy vấn lỗi.
- Chưa chuyển master Excel sang CSV, chưa đổi ngôn ngữ, chưa thêm server hay dịch vụ.
- Thay đổi trigger Teams từ @mention sang tiền tố `SAP` là hạng mục riêng; chỉ sửa sau khi kiểm tra flow Teams → email. Không suy ra trigger đã đổi chỉ vì parser hiểu chữ SAP.

Tên phiên bản bên dưới là đề xuất để chia phạm vi, không phải cam kết đã phát hành. Có thể chia hotfix nhỏ hơn nếu cần.

## 2. Căn cứ và giới hạn của bằng chứng hiện tại

Đã kiểm tra patch và bản `SAP_EMAIL_TEAMS_BRIDGE_PORTABLE_v1.0.15 (1).zip` trong Downloads. Các file source đã đối chiếu khớp với bản dự án tại thời điểm kiểm tra. Bộ test trong bản review đạt 79/79; kết quả này chưa thay thế kiểm thử SAP/Outlook/Teams thực tế.

SQLite trong bản portable có 26 yêu cầu: 23 `SUCCEEDED`, một `FAILED`, hai `COACHING`; Outbox rỗng. DB không lưu kết quả đã gửi hoặc đầy đủ chi tiết lỗi. Vì vậy lịch sử chỉ chứng minh chương trình ghi nhận đã xử lý; chưa chứng minh số liệu đúng hoặc tin đã hiển thị trên Teams.

Các lỗi đã xác nhận để làm tiêu chí hồi quy:

| Mức | Lỗi | Bằng chứng/điểm xử lý |
|---|---|---|
| P1 | Lỗi MB52 bị biến thành tồn 0 trong báo cáo tồn thấp | `check_low_stock()` bắt exception rồi tổng hợp từ danh sách rỗng |
| P1 | Nhánh `TON_KHO` gọi biến chưa khai báo `cfg` | `main.py`; câu “SAP tồn kho 32013765” đi vào nhánh này |
| P1 | KPI sai khi đọc lại chuỗi đã định dạng theo US trong chế độ EU/VN | 1000 → `1,000` → 1 |
| P1 | Phục hồi Outbox còn khoảng trống khi tiến trình dừng giữa các bước | Lưu payload, gửi, xóa Outbox, đánh dấu thư là các thao tác riêng |
| P1 | Session đang tồn tại chưa được kiểm tra quyền sử dụng lại trước từng giao dịch | Chọn theo T-code không chứng minh phiên thuộc bot |
| P2 | Kết quả có thể bị giới hạn dòng nhưng chưa truyền bằng chứng độ đầy đủ | ALV có giới hạn; không được suy số 0 từ phần chưa đọc |
| P2 | Gói portable máy đã chạy chứa dữ liệu và file phát triển | DB lịch sử, config thật, cache, test và ảnh |

## 3. Chuẩn bị trước khi sửa

1. Lập manifest SHA-256 cho source và ba gói đầu vào; ghi phiên bản, kích thước và phụ thuộc runtime. Không coi tên ZIP là bằng chứng duy nhất về phiên bản.
2. Lưu bản sao source và diff hiện có; giữ nguyên config, master, log và DB vận hành. Khi cần sao lưu DB đang mở, dùng SQLite backup API hoặc dừng worker có kiểm soát.
3. Tạo thư mục build/review riêng. Chạy test với DB tạm, webhook giả và SAP/Outlook mock; không nạp lịch sử vận hành vào fixture dùng chung.
4. Thu bộ tình huống đại diện cho các nghiệp vụ hiện hành; dữ liệu kiểm thử phải được ẩn thông tin riêng và không chứa webhook thật.
5. Ghi số đo ban đầu: ZIP, dung lượng giải nén, RAM lúc rỗi/đỉnh, thời gian khởi động, thời gian tra cứu và gửi phản hồi. Chỉ đo live khi bước triển khai đó được yêu cầu.

## 4. Đợt 1 — v1.0.16: sửa tính đúng của kết quả

### 4.1. Luồng điều phối và lỗi nghiệp vụ

- Sửa `cfg` thành cấu hình hợp lệ; kiểm thử đi từ parse lệnh đến chọn hàm nghiệp vụ và dựng phản hồi bằng mock, không chỉ kiểm tra compile.
- Không bắt lỗi MB52 rồi biến thành danh sách rỗng trong tồn thấp. Timeout, mất kết nối, thiếu quyền, thiếu grid/cột bắt buộc hoặc không nhập được bộ lọc phải trả lỗi có nguyên nhân.
- Chỉ dùng “không có dữ liệu” khi truy vấn thực thi thành công, phạm vi đúng và có bằng chứng kết quả rỗng.
- Khi không tìm thấy master/model hoặc không xác định được danh sách mã cần đánh giá, yêu cầu làm rõ hoặc báo lỗi; không tự mở rộng thành quét toàn plant.

### 4.2. Số lượng và phạm vi tính tổng

- Tách định dạng nguồn SAP (`US`/`EU`/`VN`) khỏi định dạng hiển thị. Giữ số dạng numeric/Decimal từ lớp đọc dữ liệu đến phép tính, chỉ format ở lớp trình bày.
- Phân biệt ô Excel đã là số với chuỗi hiển thị ALV: không chuyển số Excel thành chuỗi rồi parse lại theo locale SAP.
- Dùng cùng quy tắc cho MB52, STO, tồn thấp, KPI và sắp xếp. Thay cả logic tự xóa dấu chấm/phẩy còn sót trong sort MB52.
- Không âm thầm đổi chuỗi số không hợp lệ thành 0 trong báo cáo tồn. Ghi lỗi hoặc đánh dấu giá trị chưa xác định.
- Không cộng gộp PC, KG và các đơn vị khác thành một tổng. Chỉ quy đổi khi có bảng quy đổi đã xác nhận; không tự suy đoán.
- Không gộp mã khác suffix/prefix chỉ vì trùng cụm số. Mapping mơ hồ phải được giữ tách hoặc xác nhận theo master.

### 4.3. Độ đầy đủ và tồn thấp

- Chuẩn hóa kết quả truy vấn thành dữ liệu kèm metadata: phạm vi vật tư/plant/SLoc, cột đã đọc, số dòng nguồn nếu có, số dòng đã đọc, phương thức đọc, thời điểm và `complete`/`partial`/`unknown`.
- Độ đầy đủ không chỉ dựa vào số dòng: phải kiểm tra nhập đủ danh sách multi-select, đúng plant/SLoc, không sót bộ lọc cũ và đọc đủ cột bắt buộc.
- Nếu ALV vượt trần, export không đủ hoặc chỉ đọc được một phần, không gọi báo cáo là “Top N tồn thấp toàn bộ”. Trả dữ liệu một phần với thông báo giới hạn; không suy tồn 0 cho mã vắng mặt.
- Chỉ kết luận “hết hàng trong phạm vi tra cứu” khi danh sách mã chuẩn, truy vấn đầy đủ và quy tắc tồn đã rõ. Không suy ra “hết hàng toàn hệ thống”.
- Chuyển plant/SLoc bị loại, ngưỡng tồn thấp và Top N thành quy tắc có cấu hình/ghi tài liệu. Giữ giá trị hiện hành cho đến khi nghiệp vụ xác nhận thay đổi.
- Lọc subtotal/total trước khi forward-fill; ngăn cùng một lượng bị tính cả dòng chi tiết lẫn tổng.

### 4.4. Payload tương thích Teams

- Thêm metadata theo cách tương thích ngược. Không đổi tùy tiện ý nghĩa `totalRecords` hoặc các trạng thái mà flow đang dùng.
- Flow hiện tại phải phân biệt lỗi và hướng dẫn trước nhánh “0 dòng”. Nếu thêm trạng thái `PARTIAL`, cập nhật và kiểm tra flow trả kết quả cùng đợt trước kích hoạt; thông báo ngay trong summary/bảng nếu vẫn dùng payload cũ.
- Giới hạn theo byte JSON UTF-8 trên toàn payload, gồm metadata; không chỉ cắt bảng. Escape nội dung văn bản đưa vào HTML; giữ HTML bảng do chương trình dựng.
- Kiểm thử kết quả rỗng, PING, HELP, lỗi, nhiều dòng, Unicode dài và dữ liệu một phần với lời dẫn đúng ngữ cảnh.

**Điều kiện hoàn tất đợt 1:** các tình huống đã xác nhận đều có test hồi quy; lỗi SAP không tạo cảnh báo hết hàng; phép tính US/EU/VN và dữ liệu Excel cùng cho số đúng; kết quả thiếu dữ liệu hiển thị rõ giới hạn. Chạy lại toàn bộ suite và đối chiếu các nghiệp vụ trên SAP thật trước khi đưa vào dùng chung.

## 5. Đợt 2 — v1.0.17: phục hồi và quản lý phiên

### 5.1. Trạng thái SQLite và Outbox

- Dùng UUID cho mã yêu cầu; giữ message ID từ nguồn để liên kết và chống trùng. Dùng khóa có phạm vi nguồn/hộp thư khi cần, không dựa vào giờ chạy và bộ đếm.
- Lưu kết quả tra cứu cùng trạng thái yêu cầu và Outbox trong một transaction. Không dùng `INSERT OR REPLACE` để âm thầm ghi đè mục cũ do trùng ID.
- Tách `query_status` (thành công/rỗng/một phần/lỗi) khỏi `delivery_status` (chờ gửi/đang gửi/webhook đã nhận/gửi lỗi/cần kiểm tra).
- HTTP 2xx chỉ xác nhận webhook tiếp nhận; chưa ghi “đã hiện trên Teams”. Nếu cần xác nhận đăng tin, bổ sung tín hiệu từ flow và message ID Teams.
- Khi timeout có thể đã được webhook nhận, ghi trạng thái chưa rõ. Thiết kế chống trùng theo `requestId` ở bên nhận nếu cần tránh tin lặp; không cam kết exactly-once chỉ bằng SQLite phía gửi.
- Retry có `next_attempt_at`, thời gian chờ tăng dần và số lượt tối đa mỗi vòng. Phân biệt lỗi tạm thời với cấu hình/quyền không hợp lệ.
- Khi hết retry, giữ payload và lỗi ở mục có thể xem/gửi lại; không xóa âm thầm. Gửi lại kết quả đã lưu không chạy lại SAP. Tra cứu mới là thao tác riêng và có thời điểm mới.
- Khởi tạo kết nối SAP khi có nhu cầu; Outbox vẫn xử lý được khi SAP chưa mở hoặc đã mất kết nối.
- Chỉ một worker sở hữu hàng đợi/session; lần khởi động thứ hai phải phát hiện và từ chối hoặc có cơ chế claim giao dịch rõ ràng.

### 5.2. Session SAP và vòng đời tiến trình

- Gắn bot vào phiên riêng theo system/client/user và session xác định. Không xem mọi phiên MB52/ME2N là của bot; không fallback sang phiên người dùng.
- Kiểm tra lại phiên trước truy vấn, keep-alive, đóng popup và reset `/n`. Nếu người dùng đã đổi phiên hoặc không xác định được quyền sở hữu, dừng thao tác SAP và giữ yêu cầu.
- Timeout và thiếu control bắt buộc phải dừng an toàn; không bấm Enter đồng loạt để vượt popup không nhận diện.
- Nút STOP dừng đúng worker bằng PID và đường dẫn đã xác minh; không terminate mọi Python có `main.py` trong command line. Cho phép hoàn tất giao dịch DB đang ghi rồi đóng COM/DB.
- Làm rõ `dryRun`: tách thử parser/payload, thử gửi và chạy SAP thật; không để cờ được parse nhưng bị bỏ qua. Test không được tạo trạng thái hoàn tất giả trong DB vận hành.
- Kiểm tra `allowed_senders` bằng địa chỉ SMTP chuẩn hoặc domain có ranh giới; không dùng đuôi chuỗi tùy ý và không cho qua người gửi trống khi bật allowlist. Không dùng email khai trong body để xác thực người gửi.

### 5.3. Migration và quan sát

- Migration SQLite có version, chạy lặp an toàn và giữ lịch sử hiện có. Không tự diễn giải các dòng `SUCCEEDED` cũ là đã đăng thành công trên Teams.
- Thêm index cho khóa chống trùng và lịch retry; quản lý connection đóng rõ ràng. Dọn lịch sử theo hạn nhưng giữ bản ghi còn tham chiếu bởi Outbox.
- Log xoay vòng có requestId, bước xử lý, thời gian, query/delivery status, completeness và mã lỗi. Không log URL có token, toàn body email hoặc dữ liệu kho không cần thiết.
- Lưu chứng cứ gọn: số dòng, phạm vi, thời gian, trạng thái; quy định rõ hạn giữ payload lỗi và dữ liệu chi tiết nếu cần phục vụ điều tra.

**Điều kiện hoàn tất đợt 2:** test dừng tiến trình trước/sau các bước lưu, gửi và cập nhật trạng thái; restart không làm mất mục chưa gửi; cùng message không tra lại SAP; timeout có trạng thái chưa rõ; mục FAILED gửi lại được; Outbox hoạt động khi SAP vắng; hai worker không chiếm cùng session.

## 6. Đợt 3 — v1.0.18: hiệu năng và gói gọn

### 6.1. Đo và tối ưu

- Đo cùng bộ dữ liệu nhỏ/trung bình/lớn và cùng môi trường trước/sau; báo median/p95 khi đủ mẫu, RAM đỉnh và số lần truy vấn SAP. Không đặt mục tiêu số MB hay tốc độ khi chưa có baseline tương đương.
- Đọc Outlook theo lô hữu hạn, chỉ đọc thuộc tính cần thiết; không nạp mọi thư/HTML vào RAM ngay mỗi vòng.
- Giới hạn lượt Outbox mỗi vòng để backlog lỗi không làm yêu cầu mới chờ quá lâu; SAP COM vẫn xử lý tuần tự trên phiên riêng.
- So sánh đọc ALV trực tiếp và export Excel theo kích thước bảng, dùng số đo để chọn ngưỡng. Mỗi phương án phải đạt kiểm tra phạm vi và độ đầy đủ như nhau.
- Đọc Excel export theo luồng khi có thể, tránh giữ cả `raw_rows` và bản records đầy đủ trong RAM; vẫn giữ bằng chứng subtotal/cột/phạm vi.
- Nạp lại master khi file thay đổi; tạo mapping mới rồi hoán đổi sau khi đọc/validate thành công. Không xóa cache tốt trước khi đọc xong file mới.
- Đóng workbook và file trong `finally`. Kiểm tra code hiện tại trước khi sửa: MB52 đã có cleanup export trên đường thành công; bổ sung trường hợp export thất bại/timeout để tránh thư mục mồ côi. Chỉ dọn thư mục tạm do bridge tạo và đã xác minh đường dẫn.
- Kiểm tra luồng dùng clipboard: sao lưu chỉ Unicode text chưa bảo toàn ảnh/file hoặc các định dạng khác. Ưu tiên nhập không qua clipboard; nếu vẫn cần clipboard, phải có chính sách bảo toàn được kiểm thử và không để clipboard bị khóa khi lỗi.
- Gom phần dựng kết quả/gửi chung trong main sau khi có regression, không refactor đồng loạt bộ điều khiển SAP đang chạy.

### 6.2. Quyết định Excel/CSV

Giữ `DATA_MASTER.xlsx` làm nguồn cập nhật trong các đợt đầu. File hiện khoảng 37 KB; module đọc SAP export vẫn cần `openpyxl`, nên chỉ đổi master sang CSV không loại được phụ thuộc này.

Chỉ thử CSV/JSON khi đo được lợi ích đáng kể hoặc có yêu cầu vận hành cụ thể. Nếu làm: mã là chuỗi, giữ số 0 đầu và suffix, UTF-8, có schema/version/hash nguồn, mapping nhiều mã rõ ràng và cập nhật nguyên tử. Tránh duy trì hai nguồn dữ liệu độc lập. Việc bỏ `openpyxl` chỉ xem xét sau khi toàn bộ luồng SAP export có phương án thay thế đã kiểm thử trên máy đích.

### 6.3. Phát hành sạch

- Dùng danh sách file được phép đưa vào gói thay cho quét toàn thư mục.
- Loại `.db`, WAL/SHM/journal, log, cache, ảnh xác nhận, file export tạm và test khỏi gói chạy. Lưu test và evidence trong gói phát triển riêng.
- Gói mới có `config.example.json`; patch không ghi đè `config.json` hiện có. Merge khóa mặc định có kiểm tra kiểu và báo các giá trị cần bổ sung; giữ webhook và cấu hình riêng tại máy đích.
- Master do người dùng cập nhật không được ghi đè khi apply patch. Phân biệt rõ cài mới và nâng cấp.
- Chỉ cắt runtime theo phụ thuộc đã xác minh; giữ SQLite, COM, HTTPS/TLS và thư viện đọc các định dạng export thực tế. Test bản giải nén bằng chính Python portable, không dựa vào Python cài sẵn trên máy phát triển.
- Build ra thư mục phiên bản mới, output directory rõ ràng, tạo manifest/checksum; không xóa mọi ZIP cũ trong dự án hoặc tự chép ra đường dẫn cá nhân cố định. Không đóng gói lại bản máy đã chạy làm release sạch.

**Điều kiện hoàn tất đợt 3:** bảng đo trước/sau có chứng cứ, không giảm độ đúng; gói chạy không chứa cấu hình thật/DB; cài mới và patch giữ được config/master/lịch sử; portable chạy đủ nghiệp vụ trên máy đích không cần Python hệ thống.

## 7. Ma trận kiểm thử và điều kiện phát hành

| Nhóm | Tình huống bắt buộc | Kết quả mong đợi |
|---|---|---|
| Điều phối | TON_KHO, TON, VITRI, LICHSU, IB, STO, TON_THAP, OB, lịch sử bin, HELP/PING | Đúng handler, đúng phạm vi, không NameError |
| Số liệu | US/EU/VN, ô số Excel, số âm đuôi, số lẻ, mã có 0 đầu, nhiều ĐVT | Cộng/sắp xếp/hiển thị thống nhất; không mất định danh |
| Tồn thấp | SAP lỗi, thiếu plant/cột, master rỗng, bộ lọc cũ, không nhập đủ multi-select | Báo lỗi/chưa đủ dữ liệu, không tạo số 0 giả |
| ALV/export | Vượt trần dòng, subtotal, mất control, file chưa xong, export sai định dạng | Đánh dấu partial/unknown hoặc lỗi rõ ràng |
| Outbox | 2xx, 4xx, lỗi tạm thời, timeout sau tiếp nhận, restart, gửi lại FAILED | Không mất payload; không tra lại SAP khi chỉ retry gửi |
| Session | Nhiều system/client/session, user chuyển T-code, popup lạ, hai worker | Không thao tác nhầm phiên; vẫn giữ yêu cầu |
| Teams | Rỗng, lỗi, HELP, PING, partial, bảng dài, Unicode, HTML | Lời dẫn đúng ngữ cảnh, payload trong budget và render đúng |
| Cập nhật | Config/master/DB có sẵn, migration lặp, portable trên máy đích | Giữ dữ liệu người dùng; phát hiện thiếu dependency |

Kiểm thử tự động dùng mock/temp DB trước. Đối chiếu SAP thật bằng tài khoản và session riêng, dữ liệu mẫu có kết quả biết trước; giữ nguyên quy tắc nghiệp vụ nếu chưa được xác nhận thay đổi. Kiểm thử webhook/Teams là bước tích hợp riêng, cần quan sát run flow và tin thực tế; không coi HTTP 2xx hoặc 79 test cũ là xác nhận end-to-end.

## 8. Triển khai, quay lui và bàn giao

1. Mỗi đợt có danh sách file thay đổi, test bổ sung, kết quả và giới hạn chưa giải quyết.
2. Triển khai thử trên một máy/session riêng trước; kiểm tra từng nghiệp vụ và lỗi mạng/restart, rồi mới dùng chung.
3. Trước nâng cấp, dừng đúng worker, sao lưu source/config/master và DB nhất quán; ghi manifest. Chọn thời điểm không có giao dịch SAP đang chạy.
4. Nếu lỗi, quay lui source/runtime bằng bản đã lưu; giữ DB và Outbox mới phát sinh. Migration phải có chiến lược tương thích ngược hoặc script hạ schema có kiểm chứng; nếu chưa có, dừng gửi và sửa tiến thay vì chạy source cũ với schema không tương thích.
5. Không phục hồi toàn DB cũ chỉ để quay lui source: việc đó có thể mất yêu cầu mới và làm gửi trùng.
6. Bàn giao changelog, hướng dẫn cấu hình, cách xem/gửi lại FAILED, hướng dẫn STOP/START, số đo và bằng chứng kiểm thử. Phân biệt trạng thái “đã kiểm thử tự động”, “đã đối chiếu SAP” và “đã thấy kết quả trên Teams”; không ghi nhận duyệt/QA khi chưa có.

Thứ tự thực hiện đề xuất: đợt 1 → đợt 2 → đợt 3. Chỉ thu gọn runtime và xem xét chuyển định dạng master sau khi kết quả và phục hồi đạt điều kiện kiểm thử.

## 9. Mở rộng chức năng từ MA Log GR Tool V7.1

Phân tích workbook gốc và VBA đã được ghi trong [Phân tích MA Log GR và mở rộng bridge](PHAN_TICH_MA_LOG_GR_V7_1_VA_MO_RONG.md). Đây là bổ sung phạm vi kế hoạch, chưa triển khai handler hoặc chạy macro.

Sau đợt sửa tính đúng và phục hồi/session:

1. **Mở rộng A:** TIM_GR_DOC (MB51 theo inbound/reference), CHECK_IB (IB/item/quantity/status), TIM_IB (PO/SA hoặc bill), CHECK_PO (PO/SA/item/release/quantity). Xác minh cột và quy tắc theo SAP đích, giữ entity PO/IB khác material.
2. **Mở rộng B:** QUEUE_IB (đọc queue lỗi), SHIPMENT_INFO (ZMME240 khi được hỗ trợ), sau đó PRECHECK_GR tổng hợp kết quả và các dữ kiện còn thiếu. Queue lỗi và đổi WO queue là hai chức năng khác nhau.
3. **Khảo sát riêng:** ZEWMR137 chỉ thấy macro export, chưa có định nghĩa cột/filter/nghiệp vụ. Chờ dữ liệu mẫu trước khi thiết kế KPI.
4. **Phạm vi tự động ghi riêng:** GR thật, tạo/sửa/xóa inbound/batch/bin/reservation, complete outbound và đổi queue chưa thuộc bot tra cứu. Không chuyển menu Yes/No của workbook thành intent chung có thể ghi SAP ngoài ý muốn.

Điều kiện phát hành lệnh mới: JSON/natural parity, định danh giữ số 0 đầu, không gộp nhiều item/GR sai, quantity numeric và đúng đơn vị, xử lý reversal, metadata độ đầy đủ và lỗi nguồn, queue một dòng/dòng cuối, regression lệnh cũ và bằng chứng đối chiếu SAP thật. Giữ payload Power Automate và cách xưng hô hiện hành; không lấy ngày GR stock ở /SCWM/POST làm nhật ký GR chứng từ.

Đợt tối ưu đóng gói có thể thực hiện độc lập sau nền tảng. Chưa ấn định số phiên bản hoặc kích hoạt chức năng mới. Không cần chuyển master sang CSV hoặc mang cả workbook vào portable để thực hiện các chức năng này.
