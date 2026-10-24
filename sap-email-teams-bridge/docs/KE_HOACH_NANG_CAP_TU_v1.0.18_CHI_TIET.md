# Kế hoạch nâng cấp chi tiết SAP Email Teams Bridge từ v1.0.18

Ngày lập: 24/10/2026 — múi giờ vận hành Asia/Saigon.

**Trạng thái:** tài liệu thiết kế và kế hoạch triển khai. Chưa sửa source, migrate DB, cập nhật Power Automate hoặc kích hoạt chức năng theo tài liệu này.

**Mốc áp dụng:** source tại `G:\SAP\sap-email-teams-bridge`, đang khai báo v1.0.18. Các mốc 1.0.16–1.0.18 trong kế hoạch cũ từng là tên dự kiến; không được dùng chúng để kết luận các hạng mục ổn định đã hoàn tất.

## Mục lục

1. Mục tiêu, nguyên tắc và phạm vi
2. Hiện trạng đã kiểm tra
3. Rủi ro và ưu tiên
4. Lộ trình và các cổng kiểm chứng
5. Kiến trúc và hợp đồng dữ liệu
6. Quy tắc số lượng, đơn vị và định danh
7. Truy vấn SAP và độ đầy đủ
8. Nghiệp vụ NVL, Posting Change và FIFO
9. Nghiệp vụ FG và OB
10. GR, inbound, PO và tồn thấp hiện có
11. Parser, command registry và chế độ thử
12. Session SAP, Outlook và vòng đời worker
13. SQLite, Outbox, phục hồi và chống trùng
14. Power Automate, Teams và lời phản hồi
15. Cấu hình và bảo toàn dữ liệu
16. Nhật ký, kiểm tra sức khỏe và xử lý sự cố
17. Hiệu năng, Excel/CSV và runtime
18. Đóng gói, triển khai, migration và quay lui
19. Ma trận kiểm thử chi tiết
20. Checklist nghiệm thu và bàn giao
21. Mở rộng sau khi nền tảng ổn định
22. Quyết định nghiệp vụ cần xác minh
23. Backlog theo file và thứ tự triển khai
24. Manifest baseline và chứng cứ

## 1. Mục tiêu, nguyên tắc và phạm vi

### 1.1. Kết quả cần đạt

- Trả kết quả phù hợp với dữ liệu SAP thực sự đã đọc, đúng vật tư/plant/kho/batch và đơn vị.
- Không biến mất kết nối, thiếu quyền, thiếu cột hoặc đọc thiếu dòng thành “hết hàng”, “không có OB”, “đã GR” hay “sẵn sàng xuất”.
- Không mất yêu cầu khi mất mạng, restart, thư không di chuyển được hoặc SQLite tạm thời bận.
- Không tác động vào session SAP người dùng đang làm việc.
- Phân biệt tra cứu/đề xuất với tạo/post chứng từ thật.
- Giữ hệ thống nhỏ: một worker Python, SQLite, master Excel và log giới hạn dung lượng. Không thêm server hoặc dashboard nếu chưa có nhu cầu riêng.
- Cập nhật có thể kiểm chứng, giữ config/master/lịch sử và có phương án quay lui không làm mất yêu cầu mới.

### 1.2. Các nguyên tắc bắt buộc

| Nguyên tắc | Hệ quả thiết kế |
|---|---|
| Thiếu dữ liệu khác số 0 | Giá trị chưa biết là null/unavailable, kèm lý do |
| Lỗi nguồn khác kết quả rỗng | Adapter phải trả trạng thái riêng, không nuốt exception thành [] |
| Số lượng luôn đi cùng đơn vị | Không cộng PC với KG; không lấy unit đầu tiên cho toàn bảng |
| Nhãn không thay thế bộ lọc | Gán `sloc=1005` vào output không chứng minh query đúng kho |
| Kết quả tổng hợp không vượt bằng chứng nguồn | Một nguồn lỗi làm kết luận liên nguồn chưa xác định |
| Query và gửi phản hồi là hai vòng đời | Retry webhook không chạy lại SAP khi kết quả đã lưu |
| HTTP 2xx không chứng minh tin đã hiện trên Teams | Chỉ ghi webhook đã tiếp nhận nếu chưa có chứng cứ giao cuối |
| T-code đang mở không chứng minh session thuộc bot | Xác minh system/client/user/session và quyền sở hữu trước thao tác |
| Duyệt code/test không đồng nghĩa duyệt nghiệp vụ | Ghi riêng kết quả unit test, đối chiếu SAP và kiểm tra Teams |

### 1.3. Giữ nguyên trong các đợt đầu

- Các câu lệnh tra cứu đang dùng và kiểu tích hợp Outlook → Python → Power Automate → Teams.
- Các field payload cũ và cách xưng hô “Sếp ơi…”, trừ việc sửa câu kết luận sai bằng chứng.
- Master `data/DATA_MASTER.xlsx`; không đổi sang CSV chỉ để giảm 37 KB dữ liệu.
- Luồng điều khiển SAP đang chạy chỉ sửa theo lỗi cụ thể, có hồi quy; không refactor toàn bộ cùng lúc.
- Không tự thay trigger @mention thành tiền tố SAP. Thay trigger là workstream riêng sau khi đọc flow Teams → email.
- Không thêm post GR/MIGO, tạo/xóa inbound, sửa batch/bin/reservation hoặc complete outbound trong kế hoạch ổn định bot tra cứu.

## 2. Hiện trạng đã kiểm tra

### 2.1. Baseline source và gói phát hành

- `package_release.py` khai báo v1.0.18.
- Kiểm tra lại trong lượt lập kế hoạch: **103/103 test hiện có pass** bằng Python máy phát triển. Chưa phải test SAP/Outlook/Teams end-to-end hoặc chứng minh runtime portable đủ dependency.
- Các source đã kiểm tra trong ba ZIP v1.0.18 khớp source làm baseline tại thời điểm lập tài liệu.
- Portable 13.876.367 bytes; Standard 225.851 bytes; Patch 88.421 bytes. Đây là số đo ZIP hiện có, không phải cam kết kích thước bản nâng cấp.
- Cả ba gói có `config.json`. Không thấy DB trong ba gói v1.0.18 này; build vẫn cần quy tắc loại trừ rõ để không mang DB vào các lần sau.

### 2.2. Chức năng đã có code

| Canonical/nhóm lệnh | Nguồn | Việc cần làm tiếp |
|---|---|---|
| VITRI | /SCWM/POST | Đúng cột qty/unit, scope, session, completeness |
| TON / TON_KHO | MB52 | Số liệu, bộ lọc thực tế, nhiều đơn vị, không giới hạn ngầm |
| LICHSU | MB51 | Phạm vi ngày, chứng từ/item/năm và lỗi nguồn |
| IB_TREO | VL06I | Giữ nghĩa danh sách chưa GR, không dùng như trạng thái IB đầy đủ |
| STO | ME2N | Định nghĩa quantity, return/reversal và đơn vị |
| TON_THAP | MB52 đa plant | Độ đầy đủ, master/model, mapping vật tư và định nghĩa hết hàng |
| CHECK_OB | VL06O | Item/split, trạng thái và quantity còn thực hiện |
| LICHSU_BIN | /SCWM/MON | Warehouse task không đồng nghĩa material-document GR |
| TIM_GR_DOC | MB51 | Đã có, cần năm/item, reference, movement/reversal và kiểm chứng nguồn |
| CHECK_IB | VL06I | Đã có, cần nghĩa expectedQty theo dòng/tổng và trạng thái chưa biết |
| TIM_IB | VL06I | Đã có, cần fail khi không đặt được filter và trả nhiều candidates |
| CHECK_PO | ME2N | Đã có, cần xác minh MGLIEF/release/deletion/completion theo SAP đích |
| CHECK_FG_OB | MB52 + VL06O | Đã có báo cáo, không thực sự điền OB; cần sửa kết luận khi nguồn lỗi |
| CHECK_NVL_MIGO | POST + MB52 | Đã có FIFO/gợi ý, không tạo/post MIGO; cần sửa scope, eligibility và số liệu |
| HELP / PING | Worker / payload | Phân biệt thử đường truyền với kiểm tra SAP thật |

### 2.3. Các điểm đã cải thiện so với 1.0.15

- Nhánh TON_KHO đã dùng biến `config` hợp lệ; lỗi `cfg` cũ không còn là việc cần sửa trong baseline này.
- `check_low_stock()` đã truyền lỗi MB52 lên ở trường hợp đã kiểm tra; không liệt kê lại lỗi cũ như thể chưa được sửa. Vẫn cần kiểm tra thiếu dòng, scope và mapping trước kết luận 0.
- JSON đã có obNo/bin và các entity IB/PO/item/ref/materialDocument/expectedQty.
- Có payload `greeting` theo từng chức năng. Chưa mặc định rằng flow live đang sử dụng trường này.
- Resolver đã đóng workbook bằng finally. Điểm cần cải thiện là giữ mapping tốt khi reload thất bại, không phải thêm lại cơ chế đóng workbook đã có.

## 3. Rủi ro và ưu tiên

Mức P1 dưới đây là chặn phát hành dùng để ra quyết định cấp phát/điều độ. “Đã tái hiện” là probe mock, không phải thao tác SAP thật.

| ID | Mức | Tình trạng | Vấn đề | Điều kiện đóng |
|---|---|---|---|---|
| R01 | P1 | Đã tái hiện | POST và MB52 cùng lỗi nhưng NVL báo hết hàng | ERROR, không có kết luận 0 |
| R02 | P1 | Đã tái hiện | QI-only vẫn được KPI ghi khả dụng do `tot_avail or tot_all` | Khả dụng 0; QC tách riêng |
| R03 | P1 | Đã tái hiện | Category thiếu/không rõ vẫn được xem là đủ điều kiện; cat rỗng thành FT | UNKNOWN, không FIFO-ready |
| R04 | P1 | Đã tái hiện | FIFO sort sai ngày ISO; ngày thiếu có thể lên đầu | Parse ngày/giờ typed, không rank sai |
| R05 | P1 | Đã tái hiện | VL06O lỗi nhưng FG báo không OB/sẵn sàng tạo mới | Nhu cầu OB chưa rõ, không suy số dư |
| R06 | P1 | Đã tái hiện | Format số US rồi parse EU: 1000 → 1 | Numeric xuyên suốt, format một lần |
| R07 | P1 | Đọc source | Owner POST chưa chứng minh SLoc; output gán 1005 | Scope có chứng cứ/mapping đã xác minh |
| R08 | P1 | Đọc source | Fallback MB52 NVL mất điều kiện batch | Giữ nguyên material/plant/sloc/batch |
| R09 | P1 | Đọc source | Mapping substring có thể gộp biến thể/record product rỗng | Exact/alias rõ và xử lý mơ hồ |
| R10 | P1 | Đọc source | Outbox/processed tách transaction, ID theo giờ, REPLACE | UUID, transaction, khóa chống trùng, failpoint pass |
| R11 | P1 | Đọc source | Chọn session theo T-code, chưa xác minh lại ownership | Session riêng và kiểm tra trước thao tác |
| R12 | P2 | Đọc source | Alias/validation JSON/main không đồng nhất; input thiếu có thể query rộng | Registry thống nhất và validation trước SAP |
| R13 | P2 | Đọc source | dryRun parse nhưng thiếu guard thực thi rõ | Test mode cách ly SAP/send/DB production |
| R14 | P2 | Đọc source | STOP match rộng mọi Python main.py | Dừng đúng PID/path/instance |
| R15 | P2 | Đọc source | Patch có config thật; build quét thư mục và xóa ZIP cũ | Allowlist, config mẫu, output version riêng |
| R16 | P2 | Cần kiểm chứng | Quantity/release/status của PO/IB/OB bị diễn giải theo giả định môi trường | Golden cases SAP và bảng mapping nguồn |

## 4. Lộ trình và các cổng kiểm chứng

Số phiên bản tiếp theo là đề xuất, phải kiểm tra lại lúc chuẩn bị release để tránh trùng bản người khác đã phát hành. Không gán ngày hoàn thành hoặc thời gian xử lý khi chưa đo.

| Mốc | Gói việc | Phụ thuộc | Kết quả bàn giao |
|---|---|---|---|
| M0 | Chốt baseline, golden cases, quyết định nghiệp vụ | Bắt đầu | Manifest, danh sách test/input, phạm vi đã xác minh |
| M1 / dự kiến 1.0.19 | Sửa tính đúng NVL/FG, số liệu, scope, điều phối và guard phiên tối thiểu | M0 | Hotfix có regression và payload tương thích; không migrate DB |
| M2 / dự kiến 1.0.20 | Request/result/Outbox bền vững, migration, session/worker lifecycle | M1; result contract ổn định | Phục hồi và replay rõ, compatibility fallback đã thử |
| M3 / dự kiến 1.0.21 | Hoàn thiện GR/IB/PO, phân bổ nhu cầu NVL ở chế độ đề xuất | M1/M2; mapping nghiệp vụ đã xác nhận | Đối chiếu chứng từ và đề xuất bốc hàng có phạm vi |
| M4 / dự kiến 1.0.22 | Đo/tối ưu runtime, phát hành sạch, công cụ vận hành | Correctness và recovery gates pass | Số đo trước/sau, portable đạt test máy đích |
| M5 / phạm vi riêng | Queue lỗi, shipment, PRECHECK_GR và báo cáo tùy biến | Adapter/nguồn có dữ liệu mẫu và quyền | Lệnh chỉ đọc mới có acceptance riêng |

Không cần chờ M4 để loại config thật/DB khỏi gói M1. Các sửa packaging bảo toàn dữ liệu phải áp dụng ngay khi phát hành hotfix. Tối ưu tốc độ/RAM sâu làm sau.

### 4.1. Cổng phát hành

- **G0 — Baseline:** source/gói khớp manifest, lịch sử thay đổi được lưu, biết dữ liệu nào không được ghi đè.
- **G1 — Correctness:** R01–R09 được xử lý hoặc chức năng liên quan bị vô hiệu hóa rõ; không còn khuyến nghị xuất khi thiếu bằng chứng.
- **G2 — Safety/compatibility:** session guard, input guard và payload tương thích đã kiểm tra; không có tác vụ ghi SAP ngoài phạm vi.
- **G3 — Recovery:** failpoint/migration/requeue/STOP/restart pass; chưa có receiver idempotency thì giữ quy trình xử lý delivery không rõ.
- **G4 — Business:** đối chiếu dữ liệu SAP từng nghiệp vụ/ĐVT/status trong phạm vi triển khai.
- **G5 — Delivery:** nhìn thấy đúng phản hồi trên Teams và trace flow trong môi trường thử; chưa có ACK thì không ghi delivered tự động.
- **G6 — Release:** patch giữ config/master/DB, runtime portable chạy được, có fallback rehearsal.

Mỗi gói chỉ được ghi nhận những gate thực sự đã thực hiện; một tổng số test pass không thay thế các gate này.

### 4.2. Áp dụng gate theo loại release

| Release | Gate bắt buộc | Giới hạn phải công bố |
|---|---|---|
| M1 không đổi schema | G0/G1/G2; G4 cho nghiệp vụ phát hành, G5 cho payload thay đổi, G6 cho patch/fallback cùng schema | Chưa giải quyết xong R10 thì không tuyên bố recovery/exactly-once đã hoàn chỉnh |
| M2 đổi schema/vòng đời | Toàn bộ gate liên quan, bắt buộc G3 và rehearsal migration/fallback | Chưa có receiver evidence thì chỉ xác nhận WEBHOOK_ACCEPTED |
| M3 nghiệp vụ/allocation | Correctness/recovery gates của nền tảng và G4/G5 cho rule mới | Allocation là đề xuất, không reservation/post SAP |
| M4 tối ưu/gói | Hồi quy phần bị ảnh hưởng, G6 và số đo so sánh đúng scope | Không gắn cải thiện dung lượng/tốc độ với độ đúng khi chưa đo |

Nếu cần hotfix khẩn trước khi hoàn thiện toàn bộ ResultEnvelope, có thể tách phần chặn kết luận sai và tắt khuyến nghị chưa đủ bằng chứng thành patch nhỏ có test. Không dùng việc chờ refactor/schema để tiếp tục trả khuyến nghị xanh sai; cũng không gọi patch nhỏ đó là đã hoàn tất toàn M1.

## 5. Kiến trúc và hợp đồng dữ liệu

### 5.1. Pipeline đề xuất

```mermaid
flowchart LR
    A[Outlook intake] --> B[Chuẩn hóa và kiểm tra lệnh]
    B --> C[Request SQLite]
    C --> D[SAP adapter trên phiên riêng]
    D --> E[SourceResult + dữ liệu numeric]
    E --> F[Quy tắc NVL FG GR IB PO]
    F --> G[Commit result + outbox]
    G --> H[Sender độc lập kết nối SAP]
    H --> I[Power Automate]
    I --> J[Teams]
```

Sender chạy được khi SAP/Outlook chưa sẵn sàng. Với một worker, “độc lập” trước hết là tách dependency và vòng xử lý; không bắt buộc thêm thread/process. Nếu cần thread sau khi đo, COM vẫn thuộc đúng thread đã CoInitialize, không chia sẻ session giữa thread tùy ý.

### 5.2. Tổ chức code tối thiểu

| Thành phần | Trách nhiệm | Không làm |
|---|---|---|
| `main.py` / worker orchestration | Khởi động, schedule, intake, gọi service, archive mail | Chứa lại toàn bộ công thức nghiệp vụ/HTML |
| `commands.py` dự kiến | Registry alias, required entities, validator, handler | Điều khiển SAP hoặc gửi webhook |
| `models.py` dự kiến | Request, source/result, quantity, scope và reason code | Format ngôn ngữ người dùng |
| `email_parser.py` | Tách ý định/entity, tạo yêu cầu chưa xác thực | Tự quyết định số tồn/cho phép post |
| `material_resolver.py` | Mapping đúng master và ghi provenance | Đoán suffix rồi xem là mapping đã xác minh |
| `sap_engine.py` + adapter tách dần | Query, kiểm control/filter/grid, trả dữ liệu nguồn | Suy thông điệp xanh/đỏ từ exception |
| `domain_nvl.py`, `domain_fg.py` dự kiến | Eligibility/FIFO/netting trên dữ liệu nguồn | COM, Outlook, HTTP |
| `state_manager.py` | Transaction, migrations, request/result/delivery | Dùng nội dung message để tính lại nghiệp vụ |
| `teams_dispatcher.py` | Render/escape/budget/send, trả transport result | Parse lại số đã format hoặc tính KPI qua chuỗi advice |

Chỉ tạo file khi triển khai gói tương ứng. Tách adapter theo hành vi và test, không di chuyển toàn bộ sap_engine trong hotfix.

### 5.3. Request chuẩn

| Field | Kiểu/ý nghĩa |
|---|---|
| requestId | UUID do worker cấp; khác sourceMessageId và clientRequestId |
| sourceNamespace | Nguồn/tenant/chat/mailbox phù hợp dữ liệu thực tế nhận được |
| sourceMessageId | ID ổn định của sự kiện nguồn; không tự tạo UUID làm khóa dedupe |
| sourceLocator | StoreID/EntryID và metadata tìm mail, chỉ phục vụ định vị |
| command | Canonical command từ registry |
| parameters | Entity đã validate, tách material/document/PO/IB/bin |
| identity | Transport sender đã chuẩn hóa; người yêu cầu trong body chỉ là metadata |
| receivedAt | Thời điểm nhận, lưu UTC timezone-aware |
| mode | LIVE_QUERY, PARSER_ONLY, OFFLINE_RESULT_TEST, SEND_TEST theo định nghĩa mục 11 |
| inputHash | Hash nội dung chuẩn để phát hiện cùng source key nhưng body khác |

### 5.4. SourceResult chuẩn

| Field | Quy tắc |
|---|---|
| source | POST, MB52, VL06O, MB51, VL06I, ME2N, MON... |
| executionStatus | OK / ERROR / NOT_RUN |
| completeness | COMPLETE / PARTIAL / UNKNOWN |
| requestedScope | System/client/warehouse/owner/plant/sloc/material/batch/date được yêu cầu |
| appliedScope | Filter thực tế đã áp dụng, cách xác minh; không chỉ echo input |
| scopeEvidence | Cột nguồn hoặc mapping có version/hash và môi trường áp dụng |
| records | Dữ liệu typed và giữ raw field cần truy vết |
| sourceRowCount | Số dòng nguồn nếu biết thật; không biết thì null |
| extractedRowCount | Số dòng đọc; tách droppedRows/rejectedRows/subtotalRows |
| requiredColumns | Cột bắt buộc, mapping đã dùng và cột thiếu |
| startedAt/fetchedAt | UTC; thêm timezone nguồn khi diễn giải GR timestamp |
| errorCode/reasonCodes | Mã ổn định, message gọn đã che secret |
| extractionMethod | ALV_DIRECT / XLSX_EXPORT / phương thức đã hỗ trợ |
| sourceFingerprint | SAP system/client/report/layout hoặc signature cần thiết |

Kết quả rỗng hợp lệ là OK + COMPLETE + 0 dòng + scope đã chứng minh. OK + UNKNOWN không được coi tương đương rỗng.

### 5.5. BusinessResult chuẩn

- `queryStatus`: SUCCEEDED, EMPTY, PARTIAL, FAILED hoặc NEEDS_INPUT.
- `records`: numeric typed; `sources`: SourceResult; `scope`; `asOf`; `warnings`; `businessRuleVersion`.
- `metrics`: tách tổng tồn, eligible, QC, blocked, unknown, shortage theo material và UoM.
- `assessment`: OBSERVED / ESTIMATED / UNKNOWN. “Observed” chỉ dùng với số nguồn quan sát; số dư liên nguồn là tính theo mô hình được công bố.
- `eligibility` và `reasonCode` là dữ liệu, không trích ngược từ câu văn “Sẵn sàng”.
- Các reason code tối thiểu: SOURCE_UNAVAILABLE, FILTER_NOT_APPLIED, MISSING_REQUIRED_COLUMN, TRUNCATED_ROWS, UNKNOWN_SCOPE, UNKNOWN_STOCK_CATEGORY, INVALID_QUANTITY, UNKNOWN_UOM, INVALID_GR_TIMESTAMP, AMBIGUOUS_MATERIAL, SNAPSHOT_SKEW.

### 5.6. Quy tắc serialize

- Python dùng Decimal cho lượng cần chính xác, không dùng float rồi chấp nhận NaN/Infinity.
- JSON nội bộ lưu decimal dưới dạng chuỗi invariant, ví dụ `"1000.25"`, có schema version/quantity encoding; không có dấu ngăn hàng nghìn.
- JSON không hỗ trợ Decimal trực tiếp: serializer phải chuyển rõ, không dùng `default=str` cho mọi kiểu tùy ý.
- `records` trong payload cũ có thể giữ chuỗi display để tương thích flow, nhưng không parse chúng trở lại để tính KPI. Mảng canonical nội bộ và mảng display phải được đặt tên/tầng sở hữu rõ.
- Thời điểm audit lưu UTC ISO-8601; hiển thị Asia/Saigon. Thời gian GR nguồn chỉ chuyển timezone khi biết nghĩa và timezone thật, không tự giả định.

## 6. Quy tắc số lượng, đơn vị và định danh

### 6.1. Quantity

1. Phân biệt input numeric từ Excel với text ALV. Numeric Excel giữ nguyên giá trị; không cast thành text rồi parse theo EU.
2. Chuỗi nguồn parse theo format đã cấu hình/kiểm chứng cho session. AUTO chỉ dùng khi không mơ hồ; `1,234` không được tự quyết khi chưa biết locale.
3. Invalid/missing quantity tạo trạng thái UNKNOWN hoặc lỗi cột; không mặc định 0 trong công thức nghiệp vụ.
4. Số âm đuôi, dấu trừ đầu, số lẻ, khoảng trắng và giá trị 0 có fixture riêng. Loại bool, NaN, Infinity và giá trị ngoài giới hạn hợp lý.
5. Precision/rounding theo UoM và quy tắc pack đã xác minh. Không dùng CInt hoặc round tổng sớm làm mất lượng.
6. Format chỉ ở UI/export, thống nhất theo cấu hình display; sửa mọi đoạn sort tự xóa dấu chấm/phẩy còn sót.

### 6.2. Đơn vị

- A.Qty AUoM phải ghép với đúng alternate UoM. Không tự ghép quantity alternate với MEINS base-unit vì tên cột gần giống.
- Nếu có base quantity và alternate quantity, giữ cả hai cùng factor/source, chọn một metric có nhãn đúng.
- Nhóm tổng theo material + UoM; chỉ chuyển đổi bằng hệ số có nguồn và version.
- Factor thiếu/0 hoặc hai nguồn UoM khác nhau: không tính netting/shortage số học chung, trả lý do.

### 6.3. Material và chứng từ

- Material exact hoặc alias từ master đã xác minh. Không dùng `m in product` hoặc `product in m`; product rỗng luôn không khớp.
- Giữ suffix AA/AB, số 0 đầu và SAP ALPHA conversion theo loại định danh, không xóa 0 cho mọi chuỗi.
- Mapping rút gọn ra nhiều mã: hiển thị candidates và tách kết quả, hoặc hỏi chọn theo command. Không gom tồn thành một mã đại diện.
- Suy đoán BIW...AA cho mã chưa có master chỉ được đánh dấu unverified; không đủ căn cứ cấp phát/tồn thấp tự động.
- Khóa chứng từ gồm system/client/document/year/item khi nguồn hỗ trợ; khóa IB/OB có item/split; khóa tồn giữ stock ID/quant nếu có.
- `sourceMessageId`, số PO, IB, material document và material là các kiểu khác nhau. Chỉ material được đi qua MaterialResolver.

## 7. Truy vấn SAP và độ đầy đủ

### 7.1. Trước execute

- Xác minh session đúng và màn hình đúng; không chỉ tìm control nào có cùng tên.
- Kiểm tra required filter controls tồn tại, changeable và ghi thành công. Đọc lại giá trị sau khi đặt nếu SAP cho phép.
- Xóa cả LOW/HIGH, include/exclude/multi-select và filter ALV ảnh hưởng phạm vi. Không xóa bừa control ngoài màn hình query.
- Nếu không nhập đủ nhiều mã/plant/batch thì dừng hoặc báo partial theo thiết kế; không âm thầm chỉ tra mã đầu.
- Lệnh thiếu input không được execute query toàn plant ngoài chủ đích. TON_THAP theo model phải chứng minh tập mã đầy đủ trước.

### 7.2. Sau execute

- Busy timeout là lỗi có nguyên nhân, không trả idle thành công khi COM exception.
- Kiểm status E/A và popup có nghĩa; nhận diện tình huống rỗng thật theo report/message/control đã xác minh.
- Validate grid bằng signature cột và context; “ALV lớn nhất màn hình” chỉ là candidate, không tự coi đúng bảng.
- Exact technical column match trước caption; fuzzy match không được chọn một cột cho hai metric khác nghĩa mà không phát hiện mơ hồ.
- Cột bắt buộc theo command: inventory cần material/quantity/unit và trường scope; FIFO thêm GR timestamp/category/bin; document thêm document/item/year nếu cần dùng làm khóa.

### 7.3. ALV và export

- Theo dõi RowCount thật, giới hạn đọc, số dòng lỗi và bước đọc. Trần 15.000 là giới hạn kỹ thuật, không chứng minh đủ dữ liệu.
- Subtotal/header/repeated display value phải được nhận diện trước forward-fill. Không cộng cả chi tiết và dòng total.
- Export cần file mới thuộc request, nội dung/format/cột hợp lệ và file ổn định trước đọc. Không lấy file cũ cùng tên làm dữ liệu mới.
- Nếu export lỗi rồi fallback ALV, giữ reason và kiểm độ đầy đủ lại; không tự bỏ cờ partial chỉ vì fallback trả vài dòng.
- Không tuyên bố không có material/không có tồn vì material không nằm trong phần đã đọc.
- Trả result immutable của một attempt; không mix dữ liệu của lần trước với lần sau mà không ghi provenance.

### 7.4. Snapshot nhiều nguồn

- Ghi thời điểm mỗi nguồn và thời gian chênh giữa nguồn đầu/cuối.
- Ngưỡng snapshot skew là cấu hình, chốt theo thời gian truy vấn thực tế; không áp một số cố định chưa có baseline.
- Nếu số dư thay đổi nhanh hoặc truy vấn kéo dài quá ngưỡng, hiển thị “dữ liệu đọc ở các thời điểm khác nhau” và hạ assessment.
- Không gọi báo cáo nhiều query là snapshot giao dịch nguyên tử hoặc ATP chính thức.

## 8. Nghiệp vụ NVL, Posting Change và FIFO

### 8.1. Scope NVL

- Release hiện mô tả Plant 3093/SLoc 1005. Trong hotfix, nếu chưa có mapping nhiều plant được xác nhận, từ chối plant khác 3093 trước SAP.
- Owner 3093 cần mapping đã kiểm theo system/client/warehouse/owner và trường stock thực tế mới được trình bày là kho 1005.
- Nếu nguồn có SLoc, đọc/lọc trực tiếp. Nếu phải dùng mapping, lưu version và phạm vi hiệu lực; mapping chưa đủ thì scope UNKNOWN.
- Nhãn summary/badge/row phải lấy scope đã chứng minh, không hardcode 3093/1005 khi request đang tra phạm vi khác.

### 8.2. Ma trận kết luận POST/MB52

| POST | MB52 đúng cùng scope/batch | Kết luận được phép |
|---|---|---|
| Complete, có stock | Không chạy hoặc complete | Có stock POST đã đọc; eligibility/FIFO theo điều kiện riêng |
| Complete, rỗng | Complete, dương | Có tồn MB52 nhưng chưa quan sát thấy bin POST trong phạm vi kiểm; kiểm tra mapping/chênh nguồn |
| Complete, rỗng | Complete, 0/rỗng hợp lệ | Không thấy tồn trong phạm vi đã kiểm, không nói hết toàn hệ thống |
| Error/partial/unknown | Complete, dương | Hiện số MB52; Bin/FIFO chưa xác định |
| Error/partial/unknown | Complete, 0 | MB52 trong scope ghi nhận 0; kết luận tổng hợp POST chưa đủ |
| Bất kỳ | Error/partial nếu MB52 cần để kết luận | Không suy hết hàng hoặc đủ cấp phát từ phần thiếu |
| Error | Error | FAILED; không tạo dòng số 0 giả |

MB52 fallback phải giữ material, plant, SLoc và batch. Nếu report không lọc được batch thì đọc/lọc batch từ nguồn đầy đủ; không có cột batch thì không trả kết quả như đã thỏa batch.

### 8.3. Eligibility có kiểu dữ liệu

| Trạng thái | Điều kiện tối thiểu | Có xếp ưu tiên cấp phát? |
|---|---|---|
| ELIGIBLE | Qty dương hợp lệ, unit/scope rõ, category thuộc allowlist, không bị khóa theo nguồn, dữ liệu cần thiết đủ | Có, nếu timestamp FIFO hợp lệ |
| QC | Category nguồn xác nhận chờ kiểm | Không |
| BLOCKED | Stock/bin lock xác nhận | Không |
| ZERO | Giá trị numeric đúng bằng 0 từ nguồn hợp lệ | Không |
| UNKNOWN | Category/quantity/UoM/scope/lock chưa xác định | Không |

TST FT/QT/QI/BT trong tài liệu nguồn chỉ là giá trị cần kiểm chứng trên môi trường đích. Không mặc định cat rỗng thành FT và không dùng blacklist đơn thuần để mọi giá trị khác trở thành ELIGIBLE.

Bin có tên chứa BLOCK là tín hiệu cảnh báo bổ sung, không thay thế bin-lock/status nguồn. Nếu nguồn không đủ để biết lock, ghi giới hạn thay vì khẳng định đã kiểm hết điều kiện xuất.

### 8.4. FIFO

1. Parse bằng danh sách format được công bố: dd.MM.yyyy, dd/MM/yyyy, yyyy-MM-dd và các format thực tế đã có fixture.
2. Giờ parse thành time; xử lý giờ không có leading zero. Ngày hai chữ số năm phải có quy tắc rõ hoặc từ chối mơ hồ.
3. Timestamp thiếu/sai không được lên FIFO #1; đưa vào nhóm chưa xếp được với lý do.
4. Chỉ rank trong cùng material/UoM/scope và tập đủ dữ liệu. Nếu chỉ có ngày, nêu FIFO theo ngày, không ngụ ý thứ tự giờ chính xác.
5. Cùng timestamp: tie-break ổn định bằng batch/bin/stock ID và đánh dấu đồng thời gian khi cần.
6. QC/blocked/unknown không nhận số ưu tiên cấp phát. Tách `timestampRank` và `eligibleFifoRank` nếu muốn hiển thị lịch sử tất cả stock.
7. Tên kết quả: “Gợi ý FIFO trong phạm vi tồn đã đọc”; chưa có demand/kho đích thì không nói đã đủ điều kiện tạo MIGO 311.

### 8.5. KPI NVL

- Tách totalObservedQty, eligibleQty, qcQty, blockedQty, unknownQty theo material/UoM.
- eligibleQty=0 phải giữ 0, không fallback sang totalObservedQty.
- Tính từ fields canonical; không dùng nội dung migoAdvice hoặc chữ “FIFO” để nhận diện đủ điều kiện.
- Nếu quantity của nhóm unknown chưa parse được, không thể đưa ra total chính xác; hiện số known và ghi phần chưa xác định.
- Hiển thị giới hạn nguồn, thời điểm và lý do không cấp phát cùng dòng, không chỉ badge xanh.

### 8.6. Phân bổ nhu cầu ở M3, chỉ là đề xuất

Input: material, requestedQty, requestedUom, plant/sourceSloc, destinationSloc, batch rule và pack increment nếu nghiệp vụ yêu cầu. Thiếu requestedQty thì trả tồn/FIFO, không tự coi toàn bộ tồn là lượng cần chuyển.

Thuật toán dự kiến:

1. Validate nhu cầu >0, UoM, kho đích và phạm vi. Chỉ quy đổi với hệ số đã xác minh.
2. Chọn stock đủ điều kiện và nguồn complete theo scope; sort eligible FIFO.
3. Phân bổ lần lượt `min(remainingDemand, availableStock)` theo bin/batch/stock ID; áp pack rule đã chốt.
4. Không vượt tồn của dòng, không vượt tổng nhu cầu, không lấy QC/blocked/unknown.
5. Trả proposedAllocations, totalAllocated, shortage, unallocatedRounding và asOf; giữ trace stock source.
6. Snapshot hết hạn hoặc dữ liệu nguồn đổi: báo cần tra lại; đề xuất không giữ chỗ SAP và không bảo đảm còn tồn lúc thực thi.

Chưa thêm MIGO post. Nếu về sau có executor, phải có thiết kế quyền/preview/xác nhận/revalidate/idempotency SAP riêng; không tái sử dụng sender retry làm cơ chế thực thi.

## 9. Nghiệp vụ FG và OB

### 9.1. Điều kiện netting

- MB52 và VL06O phải đúng cùng material/plant/FG scope/UoM, đủ dữ liệu và thời điểm đủ gần theo chính sách.
- Xác minh SLoc FG và exclusion QC theo môi trường; không để `slocs=[plant, "3093"]` vô tình mở rộng scope khi plant khác.
- Làm rõ quantity OB: total delivery, remaining GI hay remaining pick. LFIMG không tự đồng nghĩa lượng còn cần thực hiện cho mọi trạng thái.
- Xác minh tồn nguồn đã trừ reservation hay chưa để tránh trừ OB hai lần.
- Nếu thiếu bất kỳ bằng chứng nào, mô tả số nguồn và “chưa tính được số dư”, không coi đây là ATP/reservation chính thức.

### 9.2. Công thức sau khi thống nhất mô hình

```text
signedBalance = eligibleFgStock - remainingObDemand
surplus       = max(0, signedBalance)
shortage      = max(0, -signedBalance)
```

Giữ signedBalance và shortage, không chỉ lưu availableQty bị chặn tại 0. De-duplicate theo OB + item + split/batch phù hợp; số OB chỉ để đếm đơn, không phải khóa quantity.

### 9.3. Ma trận kết luận

| MB52 | VL06O | Kết luận |
|---|---|---|
| Complete | Complete | Tính số dư theo mô hình và scope đã công bố |
| Complete | Error/partial | Hiện tồn FG; nhu cầu OB chưa xác định, không báo không OB chờ |
| Error/partial | Complete | Hiện OB; tồn FG chưa xác định, không kết luận thiếu/đủ theo stock=0 |
| Error | Error | FAILED, không tạo số dư hoặc sẵn sàng giả |
| Complete, 0 | Complete, 0 | Không có stock/nhu cầu trong phạm vi kiểm, không nói toàn hệ thống |

CHECK_FG_OB hiện chỉ báo cáo và khuyến nghị. Các alias “điền OB/ra đơn” phải được giải thích trong HELP và phản hồi để không bị hiểu là đã sửa delivery.

## 10. GR, inbound, PO và tồn thấp hiện có

### 10.1. TIM_GR_DOC

- Giữ IB/reference nguyên bản; xác minh mapping IB → XBLNR thay vì tự thêm 0/cắt chuỗi.
- Thêm fiscalYear, documentItem, dateFrom/dateTo khi cần; nếu chỉ có material document và nhiều năm khớp, trả các candidates hoặc hỏi rõ.
- Trả tất cả dòng/chứng từ trong phạm vi, không lấy một GR đầu tiên làm lịch sử đầy đủ.
- Phân biệt material document tồn tại, GR đã post, GR còn hiệu lực và lượng nhận ròng. Net movement chỉ khi mapping movement/reversal/return đã chốt.
- Không dùng ngày GR của stock POST làm lịch sử chứng từ MB51.

### 10.2. CHECK_IB và TIM_IB

- Status trống/không map được là UNKNOWN; không tự gán chưa GR.
- expectedQty phải có `comparisonScope`: item, material trong IB hoặc toàn IB và UoM. Không lặp cùng một expectedQty tổng để so mỗi dòng.
- Không có item mà nhiều dòng cùng material: trình bày nhóm rõ và tách batch/split; không chọn dòng đầu.
- Quantity khác là mismatch, không phải bằng chứng sai số IB/material.
- TIM_IB có thể trả nhiều IB giao từng phần. Giữ PO/item và reference; không loại toàn bộ khi mỗi dòng không bằng tổng demand.
- Không nhập được PO/ref/material filter thì dừng, không tra rộng rồi gán nhãn đã lọc.

### 10.3. CHECK_PO và STO

- Xác minh nghĩa MGLIEF và các cột order/delivered/open theo report/layout/loại chứng từ.
- Release G/X/1, deletion L/S, completion và delivery complete phải có mapping môi trường; status lạ giữ raw và chưa xác định.
- Không dùng cùng selection parameter/layout STO để tuyên bố đã kiểm PO mua hàng/SA nếu chưa xác nhận.
- Tách từng PO item; xử lý delivery nhiều đợt, return/reversal và nhiều đơn vị.
- Nếu source không cung cấp đủ lịch sử, chỉ hiển thị trường report và giới hạn; không tính “còn được GR” bằng công thức chưa được xác nhận.

### 10.4. TON_THAP và inventory chung

- Giữ việc truyền lỗi MB52 đã được sửa; thêm metadata độ đầy đủ và kiểm tập master/model trước khi kết luận mã vắng=0.
- Ngưỡng <50 và plant/SLoc exclusion là quy tắc cần cấu hình/công bố, không tự đổi trong đợt sửa kỹ thuật.
- Top N chỉ có ý nghĩa toàn tập khi source complete và master scope đủ; dữ liệu partial thì ghi “trong phần đã đọc”.
- Không gộp AA/AB hoặc thêm material ngoài scope do fuzzy match; không tự đổi unit thiếu thành PC như sự thật nguồn.

### 10.5. Registry nghiệp vụ và hợp đồng trả kết quả

Mỗi command phải có required entities, required source columns, allowed scopes, trạng thái có thể trả và version của rule. Thay đổi định nghĩa một metric phải cập nhật test/golden case/HELP và ghi changelog; không chỉ sửa câu HTML.

## 11. Parser, command registry và chế độ thử

### 11.1. Một registry chung

Một định nghĩa command chứa: canonical name, alias, parser rules, required/optional entities, scope policy, handler, output schema và readOnly flag. Registry cấp danh sách hỗ trợ cho JSON, HELP, dispatch và test, tránh bốn allowlist khác nhau.

Những điểm cần đóng trong baseline:

- JSON hiện chấp nhận SCWM, /SCWM/POST nhưng dispatch chưa đồng nhất; alias VL06O/MON/WHT cũng cần kiểm đầy đủ hai phía.
- CHECK_OB/LICHSU_BIN/GR/IB/PO thiếu mọi định danh phải bị chặn trước SAP, trừ mode query rộng được thiết kế rõ.
- VITRI theo bin không được hứa hỗ trợ nếu main không truyền bin vào adapter. Hoặc triển khai đúng contract, hoặc hỏi material và giải thích giới hạn.
- Mọi alias phải gọi đúng một canonical handler. Không có nhánh fall-through trả SUCCEEDED với records/summary rỗng.

### 11.2. Entity và validation

| Entity | Kiểm tra tối thiểu |
|---|---|
| materials | List/string đúng schema, số phần tử giới hạn, không null/bool/object, không trộn số PO/IB |
| plant/sloc/warehouse/owner | Chuỗi định danh giữ zero, thuộc phạm vi command hỗ trợ |
| batch/bin | Giữ ký tự/độ dài hợp lệ; không xóa input khi chuyển nguồn |
| poNo/inboundNo/obNo/document/year/item | Tách loại, không qua resolver vật tư, validate tổ hợp required |
| requestedQty/expectedQty | Decimal hữu hạn, null khác 0, dấu/precision hợp lệ theo ngữ cảnh |
| unit | Có nghĩa cùng quantity; không tự mặc định khi thiếu dữ liệu nguồn |
| days/dateFrom/dateTo | Thứ tự ngày đúng, không vượt giới hạn đã cấu hình, không nhầm timestamp với ID |
| limit/plants/model/shop | Kiểu và bounds rõ; lỗi int/type không làm văng toàn vòng worker |
| boolean fields | Parse boolean thực; chuỗi `"false"` không trở thành True chỉ vì bool(string) |

Giới hạn lượng mã/ngày/Top N được cấu hình theo đo tải và nghiệp vụ. Nếu input vượt, trả yêu cầu thu hẹp; không silently truncate request rồi trình bày như đã xử lý hết.

### 11.3. JSON và tiếng Việt cùng nghĩa

- Cùng parameter set phải tạo cùng canonical request cho cả JSON/natural.
- Tách “vị trí/ngày GR của stock” với “chứng từ/lịch sử GR của inbound/PO”.
- Câu “ra đơn MIGO/điền OB” hiện map sang báo cáo đề xuất; phản hồi ghi rõ không phải chứng từ đã tạo.
- Nếu chỉ có một số dài không rõ loại, dùng context hoặc hỏi ngắn để chọn; không auto-resolve tất cả thành material.
- Giữ legacy syntax có fixture. Mọi thay đổi intent phải kiểm các câu đang dùng, kể cả có dấu/không dấu và footer email.
- Nội dung trong email, tên người gửi, note và log là dữ liệu; không cho phép chúng trở thành command/system instruction hoặc script thực thi.

### 11.4. Chế độ thử rõ ràng

| Mode dự kiến | SAP thật | Webhook thật | DB production | Mục đích |
|---|---|---|---|---|
| PARSER_ONLY | Không | Không | Không | Xem canonical command/entities |
| OFFLINE_RESULT_TEST | Không, dùng fixture | Không | Không | Quy tắc số liệu/FIFO/HTML |
| SEND_TEST | Không | Chỉ destination test đã cấu hình | Không | Kiểm tra render/flow bằng dữ liệu mẫu |
| LIVE_QUERY | Có, trên session riêng | Destination vận hành theo cấu hình | Có | Vận hành được yêu cầu |

Đây là thiết kế dự kiến, chưa có các mode này trong worker hiện tại. Legacy `dryRun` phải được ánh xạ có tài liệu hoặc bị từ chối khi mơ hồ; không im lặng bỏ qua cờ. PING đường truyền không tự khẳng định SAP đã đăng nhập/sẵn sàng; kiểm SAP health là một phép kiểm riêng có bằng chứng.

## 12. Session SAP, Outlook và vòng đời worker

### 12.1. Pin session

- Cấu hình chọn system/client/user và session mong muốn. Khi có nhiều candidate, chỉ chọn theo chính sách đã xác nhận; không chọn candidate đầu vì T-code nằm trong safe list.
- Dùng phiên dành cho bot; ghi session identity lúc attach và kiểm lại trước query, giữ phiên, popup handling và reset /n.
- Nếu người dùng đã chuyển/gắn phiên khác, đang bận hoặc context không rõ: không phát lệnh, giữ request chờ và ghi SAP_SESSION_NOT_OWNED/SAP_BUSY.
- Không tự mở connection đăng nhập bằng credential hardcode từ VBA. Dùng phiên hợp lệ hiện có; thông tin hỗ trợ cũ không đi vào code/log/gói.
- Bot không khóa hoặc thao tác session khác để cố hoàn tất yêu cầu.

### 12.2. Keep-alive và popup

- Chỉ giữ phiên thuộc bot theo chính sách được hệ thống đích chấp nhận; không coi đọc một property RAM là bằng chứng làm mới timeout phía SAP.
- Không dùng /n làm keep-alive nếu có thể ngắt query/dialog hoặc người dùng đã chiếm lại phiên.
- Popup handler có allowlist theo context/title/control đã kiểm chứng; popup không rõ phải dừng an toàn, ghi loại lỗi gọn.
- Không bấm Enter nhiều lần hoặc lưu tự động để vượt popup lạ. Những đoạn cũ cần được test và thay từng case.
- Timeout phải có giới hạn rõ, dùng monotonic timer để đo duration; thời gian audit vẫn là datetime aware.

### 12.3. Outlook intake

- Xác định mailbox/store/folder rõ, không chọn nhầm SAP_BOT ở tài khoản khác chỉ vì trùng tên.
- Đọc thư theo lô hữu hạn; chỉ lấy Body/HTMLBody khi cần. Một thư lỗi không chặn các thư sau.
- Normalize transport sender về SMTP, bao gồm Exchange sender cần tra property phù hợp. Nếu bật allowlist mà không xác định được sender, đưa vào review, không cho qua.
- So domain theo ranh giới `@domain`, không endswith tùy ý. SenderName/senderEmail khai trong body không phải xác thực danh tính.
- Rule từ Teams → email cần bảo vệ người gửi flow và namespace; tránh flow tự bắt phản hồi của chính mình thành yêu cầu mới.
- `SendAndReceive(False)` mỗi vòng là tác vụ cần đo và cấu hình; không ép sync thường xuyên chỉ để che lỗi intake.

### 12.4. Single instance và STOP/START

- Một instance trên một thư mục dữ liệu/session. Dùng cơ chế khóa OS có vòng đời theo process, ví dụ mutex/lock tương ứng canonical DB path; PID file chỉ hỗ trợ thông tin, không phải khóa duy nhất.
- START thứ hai báo đã có worker với đường dẫn/instance, không gắn thêm SAP và không gửi lại queue.
- STOP xác minh PID + executable + command path/instance token trước dừng; không match mọi `main.py`.
- Graceful stop: ngừng nhận thư mới, kết thúc hoặc đánh dấu attempt đang làm, commit/rollback transaction, giải phóng COM/DB/lock, ghi stopped reason.
- Khi buộc dừng, startup recovery xử lý các trạng thái dang dở; không đặt success chỉ để thoát.
- CoInitialize/CoUninitialize có vòng đời rõ. Gọi GC/CoFreeUnusedLibraries liên tục không thay thế việc đóng reference; chỉ giữ nếu đo được lợi ích và không làm lỗi handle.

## 13. SQLite, Outbox, phục hồi và chống trùng

### 13.1. Mục tiêu bảo đảm

- Không mất payload đã commit do crash/mất mạng.
- Không chạy lại query đã có kết quả commit chỉ vì send/move mail thất bại.
- Không claim bảo đảm exactly-once posting Teams chỉ bằng DB phía client.
- Phát hiện request conflict/duplicate và trạng thái gửi không rõ thay vì REPLACE im lặng.

### 13.2. Schema đề xuất, additive migration

Tên và field dưới đây là hợp đồng thiết kế; chưa tạo bảng trên DB hiện tại.

| Bảng | Field chính | Constraint/index |
|---|---|---|
| schema_migrations | version, checksum, appliedAt, from/to appVersion | Version UNIQUE; checksum không đổi khi đã apply |
| requests | requestId UUID, sourceNamespace, sourceMessageId, inputHash, command, paramsJson, receivedAt, queryStatus, archiveStatus, legacyId | UNIQUE namespace+sourceMessageId khi ID hợp lệ; index queryStatus/receivedAt |
| request_locators | requestId, storeId, entryId, internetMessageId nếu có, lastSeenAt | Index locator; EntryID có thể đổi khi move |
| query_results | resultId, requestId, version, canonicalJson, sourceMetadataJson, scopeJson, ruleVersion, queryStatus, startedAt/completedAt, resultHash | UNIQUE requestId+version; FK request |
| deliveries | deliveryId UUID, resultId, destinationKey, idempotencyKey, payloadJson, payloadHash, state, attempts, nextAttemptAt, leaseOwner/leaseUntil, acceptedAt, confirmedAt | UNIQUE resultId+destinationKey; index state+nextAttemptAt |
| delivery_attempts | attemptId, deliveryId, start/end, duration, transportClass, httpStatus, errorCode, flowRunId/teamsMessageId nếu có | Append-only; index deliveryId+time |
| legacy_import_map | sourceTable, legacyRowKey, newId, sourceHash | UNIQUE sourceTable+legacyRowKey để import lại không nhân bản |

Không lưu webhook có token vào log. Với queue mới, ưu tiên destinationKey tham chiếu cấu hình riêng; việc thay endpoint khi đang có pending phải có mapping/migration ghi nhận, không tự chuyển payload sang đích khác.

### 13.3. Định danh và dedupe

1. UUID định danh request nội bộ; `clientRequestId` và ID legacy giữ riêng để trace.
2. Khóa logic là sourceNamespace + sourceMessageId ổn định. Namespace dựa nguồn tin cậy, không tùy ý tin mọi field trong body.
3. Thiếu sourceMessageId: dùng transport ID ổn định có phạm vi mailbox/store hoặc fallback có tài liệu. Không dùng chỉ body hash để gộp hai yêu cầu giống nội dung được gửi ở thời điểm khác.
4. Cùng source key và cùng inputHash: reuse request/result, chỉ sửa locator/archive khi cần.
5. Cùng source key nhưng khác inputHash: SOURCE_CONFLICT, đưa review; không REPLACE kết quả cũ.
6. EntryID là locator vì move có thể đổi. Dedupe không phụ thuộc duy nhất vào EntryID hiện tại.

### 13.4. Trạng thái tách biệt

**Query:** RECEIVED → VALIDATED → RUNNING → SUCCEEDED/EMPTY/PARTIAL/FAILED/NEEDS_INPUT. Trạng thái điều phối RUNNING khác kết quả business/queryStatus ở envelope.

**Delivery:** PENDING → SENDING → WEBHOOK_ACCEPTED → TEAMS_DELIVERED nếu có bằng chứng. Nhánh lỗi: RETRY_WAIT, UNCERTAIN, NEEDS_REVIEW, CANCELLED có lý do.

**Archive mail:** NOT_ARCHIVED → PENDING → PROCESSED/FAILED/REVIEW hoặc ARCHIVE_ERROR theo chính sách. Thư ở Processed chỉ có nghĩa intake/query đã xử lý theo định nghĩa; UI không dùng folder này làm bằng chứng đã giao Teams.

Không xóa delivery ngay sau HTTP thành công. Giữ bản ghi tối thiểu và attempt evidence theo retention; payload chi tiết có thể được purge theo policy riêng sau khi không còn cần replay.

### 13.5. Transaction và side effect

| Bước | Trong SQLite transaction? | Quy tắc |
|---|---|---|
| Claim/insert request | Có, ngắn | Dedupe atomic trước gọi SAP |
| Query SAP | Không | Không giữ DB lock khi COM chậm |
| Save result + query terminal + enqueue delivery | **Cùng một transaction** | Commit tất cả hoặc không phần nào |
| Claim delivery | Có, ngắn | Conditional update/lease, không hai attempt cùng sở hữu |
| HTTP send | Không | Không giữ DB lock chờ mạng |
| Record attempt + delivery state | Có | Giữ HTTP/error/uncertain, không thay payload |
| Move/save Outlook mail | Không | Side effect riêng, retry archive không query/send lại |

Payload delivery sau commit là immutable. Muốn tra mới hoặc đổi thông điệp kết quả phải tạo result version/delivery mới có liên kết, không sửa payload của một idempotency key đã gửi.

### 13.6. Retry

- Timeout, lỗi DNS/kết nối, 429 và 5xx được phân loại riêng. 429 xử lý Retry-After hợp lệ trong giới hạn cấu hình.
- 400 payload/schema, 401/403 quyền và 404/410 endpoint cần review/cấu hình; không thử lại liên tục cùng dữ liệu.
- Phân loại retry còn phụ thuộc khả năng side effect đã xảy ra. Flow có thể đã đăng Teams rồi trả 5xx, reset kết nối hoặc client crash trước ghi ACK. Khi chưa có receiver idempotency/reconcile, không tự retry mọi 5xx/reset chỉ vì được coi là lỗi tạm thời; trường hợp có thể đã tiếp nhận phải vào UNCERTAIN.
- Retry giãn theo exponential backoff có jitter và trần; số attempt, tuổi queue, batch size và request budget được cấu hình, không hardcode 3 vòng rồi quên mục lỗi.
- Hết lượt/tuổi queue: NEEDS_REVIEW, giữ error/payload/hash. Có thao tác list/show/requeue có chủ đích.
- Query mới sau khi stock đã thay đổi là yêu cầu mới, không đổi nghĩa nút “gửi lại phản hồi”.
- Failed payload JSON hỏng/không khớp hash phải quarantine, không thay `{}` rồi gửi như kết quả hợp lệ.

### 13.7. Timeout không rõ và phía nhận

- Server có thể nhận request trước khi client timeout. Retry phải giữ cùng request/delivery/idempotency key và payload hash.
- Nếu chưa có cơ chế idempotency/reconcile phía nhận, timeout mơ hồ nên vào UNCERTAIN/NEEDS_REVIEW theo policy; không mặc định retry vô hạn không trùng.
- Flow nhận cần durable unique key và atomic claim; thao tác “tìm xem có chưa” rồi “tạo row” riêng rẽ chưa đủ khi nhận đồng thời.
- Sau Teams post nhưng trước ghi marker, receiver cũng có crash boundary. Cần message ID/run evidence/reconcile hoặc đánh dấu chưa rõ; unique key một mình không bảo đảm exactly-once.
- Không tạo inbound HTTP listener mới chỉ để nhận ACK. Nếu cần ACK, ưu tiên cơ chế đã kết nối/được phép hoặc kiểm tra run evidence riêng; thiết kế sau khi biết khả năng flow hiện tại.

### 13.8. Startup recovery và các dependency

- Sender khởi động từ SQLite/config mà không cần SAP hoặc Outlook đã mở.
- Expired delivery lease được phân loại phục hồi; trước đó đã gọi HTTP hay chưa có thể chưa rõ, không tự reset mọi SENDING thành chắc chắn chưa gửi.
- RUNNING read-only query bị crash có thể retry thành attempt/snapshot mới nếu chưa commit result; ghi lại timestamps, không tuyên bố snapshot cũ được giữ.
- Query đã commit result không chạy lại do failed mail move hoặc accepted state.
- SAP unavailable: giữ query chờ theo giới hạn, vẫn gửi Outbox đã có. Outlook unavailable: vẫn gửi kết quả đã lưu, intake resume sau.
- SQLite locked/disk full: không gửi kết quả chưa thể lưu bền vững rồi bỏ trace. Báo lỗi vận hành, giữ thư để phục hồi.

### 13.9. Migration DB v1.0.18

1. Sao lưu bằng SQLite backup API hoặc dừng có kiểm soát; ghi schema/version/hash và số dòng hai bảng hiện tại.
2. Chạy migration trên bản sao trước: integrity_check, foreign_key_check và kiểm số lượng/hash payload.
3. Tạo bảng mới additive, giữ bảng cũ nguyên trong ít nhất chu kỳ fallback. Không tạo lại DB trống đè file thật.
4. `processed_messages` cũ → LEGACY_PROCESSED/legacy_unverified, giữ command/status/timestamp/IDs. Không nâng thành TEAMS_DELIVERED và không tự phát lại lịch sử.
5. Outbox PENDING cũ → delivery imported; giữ payload, destination, attempts và request ID cũ làm idempotency key vì có thể đã tới receiver. Chỉ tự resume khi xác minh chưa gửi hoặc receiver có dedupe tương thích ID cũ; nếu lịch sử không đủ thì UNCERTAIN/NEEDS_REVIEW. attempts=0 không chứng minh chưa từng gửi, vì có thể crash sau send trước tăng attempts.
6. Outbox FAILED cũ → NEEDS_REVIEW; không tự gửi hàng loạt lúc nâng cấp.
7. Outbox không có processed row vẫn phải được bảo toàn, tạo imported request với scope unknown; không chạy lại SAP để bù metadata.
8. Collision req_id, body/hash/destination không khớp hoặc JSON hỏng → quarantine với mapping; không REPLACE.
9. Migration chạy lại idempotent qua legacy_import_map; validation fail thì rollback transaction migration, giữ backup và log gọn.
10. Chuẩn bị binary/source fallback hiểu schema mới trước khi activate migration; chi tiết ở mục 18.

## 14. Power Automate, Teams và lời phản hồi

### 14.1. Payload tương thích

Giữ field cũ: requestId, messageId, command, status, greeting, senderName/senderEmail, summary, table_html, completedAt, totalRecords, displayedRecords, records. Thêm schemaVersion/businessStatus/completeness/scope/asOf/reasonCodes theo kiểu additive nếu schema flow cho phép.

Tương thích bao gồm cả kiểu và nghĩa: lưu snapshot schema và expressions flow hiện tại, kiểm string/number/null/array/date format và cách consumer dùng từng field. Canonical numeric nội bộ phải qua adapter legacy khi dựng payload cũ; thay kiểu/ý nghĩa field cần schema version hoặc kiểm nhận consumer trước. Golden payload test gồm có dữ liệu, rỗng, partial, failed, HELP/PING và replay payload cũ.

Để tránh phá flow live đang kiểm FAILED/INFO/SUCCEEDED:

- Hotfix có thể giữ status legacy và đưa chi tiết FAILED/PARTIAL/EMPTY vào businessStatus/summary/greeting phù hợp.
- Không gửi PARTIAL như một status mới trước khi kiểm expression flow hiện tại. Source truth về incomplete vẫn phải xuất hiện trong nội dung ngay cả khi dùng envelope legacy.
- `totalRecords` tiếp tục là số records kết quả chuẩn mà app đang có; nếu cần số dòng nguồn thêm sourceRows/extractedRows riêng. Không biến nó thành tổng SAP giả.
- Bảng bị cắt chỉ ảnh hưởng hiển thị; metadata phải nói tổng phạm vi và số dòng đã đọc. Không để cắt bảng đồng nghĩa đủ dữ liệu nguồn.

### 14.2. Một nơi sở hữu lời dẫn

- Chọn Python `greeting` làm nguồn chính để cùng rules với BusinessResult. Flow dùng greeting nếu có, fallback tương thích cũ khi thiếu.
- Trước sửa flow, đọc đúng bản live và lưu nội dung action/schema hiện có; không giả định expression đã đồng bộ với Python.
- Tránh hai lớp chào “Sếp ơi” hoặc lời dẫn generic xanh che cảnh báo partial.

| Ngữ cảnh | Lời dẫn dự kiến |
|---|---|
| Query đủ, có dữ liệu | Sếp ơi, kết quả tra cứu trong phạm vi đã chọn đây ạ. |
| Rỗng hợp lệ | Sếp ơi, em chưa tìm thấy dữ liệu phù hợp trong phạm vi này ạ. |
| Một nguồn thiếu/lỗi | Sếp ơi, em mới kiểm tra được một phần ạ; phần chưa xác định em ghi rõ bên dưới. |
| Lỗi query | Sếp ơi, lần này em chưa tra cứu được SAP ạ. |
| Đề xuất FIFO | Sếp ơi, đây là gợi ý bốc hàng theo FIFO từ dữ liệu vừa đọc ạ; chưa tạo phiếu MIGO. |
| Thử đường truyền | Sếp ơi, em đã nhận yêu cầu thử đường truyền ạ; kết quả này chưa xác nhận phiên SAP. |

Các câu cụ thể có thể rút gọn khi render, nhưng ý nghĩa phải khớp trạng thái thật.

### 14.3. Bảng và KPI

- Render/escape text từ người gửi, note, SAP cell và error. Chỉ HTML do chương trình kiểm soát mới được đưa vào raw table_html.
- KPI lấy từ typed metrics, không regex advice. Tách các UoM, QC, blocked và unknown.
- Hiển thị scope/time và lý do giới hạn trước khuyến nghị quan trọng. Không đưa stack trace/secret vào nhóm chat.
- Ngưỡng 26.000 byte hiện là budget của app; đo cả JSON UTF-8, metadata, custom_html và nội dung flow thêm vào. Xác minh giới hạn action Teams thực tế, không ghi “không bao giờ lỗi 413” chỉ từ byte count client.
- Nếu cần nhiều phần, message part index/total và idempotency key từng phần; hoàn tất một phần không coi đã giao cả báo cáo.
- Giai đoạn đầu ưu tiên bảng tóm tắt có scope/filter rõ; chưa thêm attachment/upload chia sẻ nếu chưa có yêu cầu riêng.

### 14.4. Xác nhận giao tin

- HTTP response success → WEBHOOK_ACCEPTED.
- Flow run/action Teams success gắn đúng IDs, hoặc callback có xác thực → TEAMS_DELIVERED.
- Chưa có bằng chứng cuối → delivery confirmation unknown, không phát thông báo “đã hiện Teams” tự động.
- Receiver idempotency/delivery evidence và thay trigger Teams là thay đổi flow độc lập, thử trên bản test trước. Không thay quyền/chat đích chỉ để test.

## 15. Cấu hình và bảo toàn dữ liệu

### 15.1. Nhóm cấu hình dự kiến

| Nhóm | Khóa/ý nghĩa dự kiến | Quy tắc |
|---|---|---|
| App | timezone, dataDir, logDir, mode | Đường dẫn rõ, không chạy nhầm dataDir |
| Outlook | mailbox/store/folder, pollInterval, intakeBatchSize, sync policy | Giữ default hiện hành đến khi đo/chốt |
| SAP identity | system/client/user/session selector, strict ownership | Không fallback session user |
| Query | maxRows, maxMaterials, date bounds, timeouts, export policy | Validate type/bounds, partial có lý do |
| NVL | scope fixed/mapped, owner/warehouse mapping, eligible categories, blocked rules | Có nguồn/version và golden case |
| FG/OB | FG scope, remaining-demand definition, snapshot skew policy | Không suy từ tên biến |
| Numbers | source format, display format, UoM/rounding/conversion | Tách source và display |
| Delivery | destinationKey, HTTP timeout, retry/backoff/batch, uncertain policy | Không log token, không đổi pending đích ngầm |
| State | retention, migration policy, archive policy | Không purge pending/needs-review còn cần xử lý |
| Feature flags | allocation/queue/shipment/experimental rules | Default tắt cho chức năng chưa qua gate |

Tên khóa cụ thể chốt khi triển khai schema config; không thêm hàng loạt khóa không dùng. Chỉ tạo cấu hình cho rule/driver có người vận hành cần kiểm soát.

### 15.2. Load và nâng cấp config

- Validate JSON object, type/bounds và tổ hợp bắt buộc. Lỗi webhook/session/scope quan trọng phải fail rõ, không silently dùng placeholder rồi đánh dấu gửi thành công.
- Parser-only có thể dùng config test; production thiếu webhook không được coi dry-run success và hoàn tất thư.
- Auto-fix dấu quote/comma nếu giữ phải báo rõ nội dung đã hiệu chỉnh trong bộ nhớ; không ghi đè file người dùng ngoài tác vụ cập nhật được yêu cầu.
- Patch mang config.example.json/schema và logic merge default, không config.json thật. Key unknown giữ để tương thích nếu an toàn; key mâu thuẫn báo người vận hành.
- Master/state/log tách khỏi code package. Quyền truy cập theo tài khoản chạy; không đưa credential/token trong source hoặc release notes.
- Chỉ checksum/hash cấu hình vào evidence khi cần; evidence công khai chỉ dùng bản đã che secret, không copy URL token.

## 16. Nhật ký, kiểm tra sức khỏe và xử lý sự cố

### 16.1. Nhật ký cần có

requestId, deliveryId, source ID rút gọn khi phù hợp, command, app/rule/schema version, stage, duration, query/delivery/archive status, scope và reason codes. Trace từ request → sources → result → delivery attempts → mail archive.

Không log credential, webhook query token, toàn body email hoặc toàn bảng tồn mặc định. Dữ liệu chi tiết chỉ lưu theo policy điều tra có hạn và phạm vi cần thiết.

### 16.2. Retention

- Log xoay theo dung lượng/thời gian, giữ số file cấu hình; lỗi không làm log tăng vô hạn.
- Cleanup processed history không xóa khóa request còn liên quan pending/uncertain/needs-review.
- Payload chi tiết và metadata audit có retention khác nhau; purge ghi record đã dọn và giữ hash/IDs khi cần dedupe.
- Cleanup file export chỉ trong thư mục do bridge tạo, có manifest/marker và đường dẫn đã resolve. Không recursively delete thư mục lấy từ input người dùng.
- Failed export/timeout cần cleanup sau hạn; đường thành công hiện đã có cleanup phải được giữ và kiểm lại, không refactor chỉ để thêm duplicate cleanup.

### 16.3. Công cụ vận hành nhỏ gọn

Thiết kế CLI đọc state hoặc màn hình console/status file cục bộ, không bắt buộc web UI:

- Xem phiên bản/schema/config validation và instance đang chạy.
- Xem số request chờ/query lỗi, delivery pending/uncertain/needs-review và tuổi mục cũ nhất.
- Xem chi tiết một request/delivery với secret đã che.
- Gửi lại payload lỗi theo ID; thao tác yêu cầu mới/tra lại tách riêng.
- Kiểm integrity/backup/migration dry-run trên bản sao.
- Xem lý do SAP/Outlook không sẵn sàng và lần thành công gần nhất, không mặc định ready chỉ vì process còn sống.

Các lệnh CLI cụ thể chưa tồn tại; phải ghi đúng trong tài liệu bàn giao sau khi được triển khai/test. Không tạo automation gửi thông báo định kỳ chỉ vì có health metrics.

### 16.4. Runbook sự cố

| Sự cố | Hành động mong đợi |
|---|---|
| SAP chưa mở/mất phiên | Giữ query chờ theo policy; Outbox vẫn chạy; hướng dẫn mở đúng session |
| User đang thao tác session | Dừng SAP action, không /n; request giữ nguyên và báo ownership |
| Webhook lỗi cấu hình/quyền | NEEDS_REVIEW, sửa cấu hình có trace rồi requeue có chủ đích |
| Webhook timeout không rõ | Kiểm run/message evidence, tránh gửi lại mù |
| Đã có kết quả nhưng không thấy Teams | Theo request/delivery/flow run; không tra SAP lại mặc định |
| Mail move thất bại | Retry archive riêng; không query/send lại |
| SQLite locked/disk full | Dừng nhận việc có thể không lưu, giữ thư, báo lỗi; không tạo DB mới che sự cố |
| Master đang mở/lỗi cập nhật | Giữ mapping tốt trước đó có cờ stale nếu policy cho phép; command cần mapping mới thì báo thiếu |
| Sai số liệu phát hiện sau gửi | Dừng khuyến nghị liên quan, lưu evidence, tạo bản sửa có liên kết; không âm thầm thay lịch sử |

## 17. Hiệu năng, Excel/CSV và runtime

### 17.1. Baseline và phương pháp đo

Đo riêng parser, resolver, wait SAP, query, extraction, domain compute, render/serialize, HTTP và archive. Ghi tổng end-to-end cùng queue wait, RAM idle/peak, số COM calls nếu instrument được, số query SAP/request, ZIP và dung lượng giải nén.

- Mẫu nhỏ/trung bình/lớn phải cùng dữ liệu/scope và điều kiện SAP giữa trước/sau.
- Có warm/cold start; báo số mẫu và khoảng biến động. Median/p95 chỉ khi đủ mẫu, không bịa percentile từ một lần chạy.
- Chốt ngưỡng SLO sau baseline theo nghiệp vụ; không cam kết giảm X% hoặc dưới N giây trước khi đo.
- Acceptance tối ưu phải đi cùng correctness và completeness; nhanh hơn vì đọc thiếu dữ liệu không được tính là cải thiện.

### 17.2. Các tối ưu ưu tiên

1. Batch Outlook hữu hạn và retry delivery có quota mỗi vòng để không starvation.
2. Material resolver kiểm mtime/hash hợp lý, đọc một lần cho cả batch; build mapping mới rồi swap sau validate thành công.
3. COM query vẫn tuần tự; gộp multi-select đúng phạm vi để giảm số lượt SAP, không gộp trái filter/đơn vị.
4. Chọn ALV direct/export theo số đo; tránh hàng chục giây fallback nếu biết đường hiện tại không hỗ trợ export.
5. Read-only streaming workbook export, tránh cùng lúc giữ raw_rows, normalized rows, canonical records và HTML cho toàn bộ dữ liệu nếu không cần.
6. Renderer chỉ dựng phần hiển thị; metrics tính từ canonical records một lần, không rebuild lại toàn KPI sau mỗi lần cắt payload.
7. DB queries có index, lấy pending theo due time/batch, đóng connection rõ ràng. Không giữ transaction trong COM/HTTP.
8. Dọn reference COM theo vòng đời; không dùng GC như biện pháp duy nhất giải quyết tăng RAM.
9. Nếu cache được giới thiệu, key gồm scope/mapping/rule version, TTL và asOf rõ; không trả stock cũ như mới khi nguồn lỗi. Mặc định chưa cache tồn biến động qua request.

### 17.3. Excel/CSV

- Giữ XLSX master vì khoảng 37 KB và người dùng đang cập nhật bằng Excel.
- openpyxl còn dùng SAP export nên chuyển riêng master sang CSV chưa loại được dependency.
- Nếu về sau đo thấy lợi ích: một nguồn Excel authoritative, bản CSV/JSON sinh có source hash/version và cập nhật atomic; không để người dùng sửa hai nguồn song song.
- Giữ mã dạng chuỗi, zero đầu, Unicode, alias nhiều mã và schema. CSV không tự giải quyết số locale/UoM hoặc dữ liệu bị cắt.
- Không chuyển hay đóng gói workbook macro MA Log GR vào runtime; chuyển quy tắc đã xác minh sang adapter/domain nhỏ.

### 17.4. Clipboard

Luồng multi-select có dùng clipboard cần test bảo toàn các định dạng. Sao lưu Unicode text không bảo toàn ảnh/file clipboard. Ưu tiên phương án không clipboard nếu đủ tin cậy; nếu cần giữ, context manager/finally phải đóng clipboard khi lỗi và bảo toàn theo khả năng đã công bố. Không tuyên bố 0 Clipboard khi code thực tế còn dùng.

## 18. Đóng gói, triển khai, migration và quay lui

### 18.1. Gói phát hành

| Gói | Chứa | Không chứa |
|---|---|---|
| Patch | Source thay đổi, migration, manifest, config.example, changelog và hướng dẫn nâng cấp cần thiết | Config thật, master người dùng, DB/log, runtime không cần đổi |
| Standard | Source runtime cần, cấu hình mẫu, template cần cho vận hành | Test/dev evidence, cache, ZIP cũ, token, DB vận hành |
| Portable | Standard + runtime/phụ thuộc đã xác minh | Python test modules/dev packages không dùng, dữ liệu riêng từ máy đã chạy |
| Development/evidence riêng | Tests, fixtures đã ẩn dữ liệu, reports/checklist | Credential/webhook thật và dữ liệu kho không cần thiết |

Build theo allowlist; master demo/cài mới phải được ghi rõ và không overwrite master existing. Dữ liệu riêng sống ngoài phần code được update hoặc được updater bảo vệ rõ.

### 18.2. Manifest và dependency

- Manifest version/buildId/time, source SHA-256, dependency versions, Python architecture và các file đầu ra/hash.
- Lưu runtime origin/version, không lấy ngẫu nhiên một runtime ngoài project mà không kiểm hợp đồng.
- Giữ SQLite, COM, HTTPS/TLS và thư viện đọc các export thực tế. Không cắt DLL chỉ theo tên nhỏ/lớn.
- Xác minh Start/Silent/Stop đều được đưa vào đúng tên file và hoạt động trong đường dẫn có dấu/khoảng trắng.
- Build vào output directory rõ/version mới, không xóa toàn bộ ZIP trong BASE_DIR và không tự copy ra Downloads/đường dẫn cá nhân cố định.
- Thử bản giải nén bằng chính runtime đi kèm, trên môi trường không dựa vào Python hệ thống/site-packages của máy phát triển.

### 18.3. Quy trình hotfix không đổi schema

1. Chốt baseline/diff và gói M1; chạy G1/G2 cùng hồi quy.
2. Ghi danh sách file phải thay; không unpack đè cả thư mục data/config.
3. Dừng đúng worker, xác minh session không đang query, sao lưu source/config/master và DB nhất quán.
4. Apply patch theo manifest, validate config và dependency.
5. Start một instance, chạy smoke test có nguồn mẫu biết trước và quan sát Teams ở đích thử phù hợp.
6. Nếu lỗi, thay lại source tương thích cùng schema; giữ DB mới, kiểm queue trước gửi tiếp.

### 18.4. Quy trình release có schema mới

- Đã có migration dry-run, integration test và fallback tương thích trước triển khai.
- Dừng sender/intake, backup DB hiện tại bằng API, ghi pending/uncertain count.
- Apply source/migrate, kiểm integrity/count/hash và mapping legacy. Chưa vượt gate thì chưa gửi queue tự động.
- Chỉ PENDING đã được xác định an toàn được resume theo policy; PENDING legacy có thể đã gửi phải UNCERTAIN/review trừ khi có dedupe tương thích và bằng chứng đủ. FAILED legacy cũng giữ review, không tự phát hàng loạt.
- Smoke test request trước/sau migration và archive; không tự phát lại các processed cũ.

### 18.5. Quay lui không mất dữ liệu

Không mặc định copy nguyên source 1.0.18 rồi chạy với schema mới. Source cũ không hiểu bảng mới, có thể bỏ pending hoặc gửi trùng.

- Chuẩn bị compatibility fallback đã test, hiểu schema mới hoặc có compat layer đọc/ghi được thiết kế. Chỉ một dispatcher có quyền gửi.
- Nếu dual-write tạm thời, writes liên quan cùng transaction và có chủ sở hữu rõ; không chạy hai worker old/new cạnh nhau.
- Khi quay lui: stop instance, backup DB hiện tại, thay source fallback, giữ schema/requests/results/deliveries mới, verify rồi mới resume sender.
- Nếu không có fallback tương thích, dừng sender và sửa tiến; không restore DB cũ làm mất yêu cầu sau upgrade.
- Test rehearsal có request A trước nâng cấp, B sau nâng cấp và C sau fallback: cả ba còn trace/queue đúng; không mất B và không phát lại A mù quáng.

### 18.6. Triển khai theo phạm vi

Pilot trên một máy/session, một nhóm nghiệp vụ và đích Teams thử trước. Mở rộng sau khi đủ mẫu lỗi/happy path và nguồn thực đã đối chiếu. Feature flag có thể tắt riêng NVL/FG/adapter mới nếu phát hiện vấn đề, vẫn giữ HELP/Outbox và các query không liên quan khi an toàn.

Chỉ gọi bản phát hành là đã đạt một gate khi có evidence đi kèm. Không tự ghi manager review/QA approval hoặc trạng thái chứng từ SAP chưa quan sát.

## 19. Ma trận kiểm thử chi tiết

Các case dưới đây là yêu cầu kiểm thử dự kiến, không phải test mới đã được viết hoặc đã pass. Mỗi case phải ghi fixture, version/hash, expected result và evidence. Đầu ra business cần kiểm numeric fields và trạng thái, không chỉ tìm câu văn trong HTML.

### 19.1. NVL và scope — NVL01–NVL12

| ID | Tình huống | Kết quả bắt buộc | Lớp kiểm |
|---|---|---|---|
| NVL01 | POST và MB52 cùng timeout | FAILED; không stock=0/hết hàng | Adapter mock + domain + render |
| NVL02 | POST lỗi, MB52 1005 có 100 PC | Hiện MB52 100; bin/FIFO unknown, không khẳng định chưa có bin | Composite result |
| NVL03 | POST complete rỗng, MB52 complete cùng scope bằng 0 | Rỗng/0 đúng phạm vi, lý do rõ | Domain |
| NVL04 | Owner 3093 nhưng chưa có mapping SLoc | UNKNOWN_SCOPE; không in kho 1005 như fact | Scope validator |
| NVL05 | Request plant 3001 trong mode NVL chỉ hỗ trợ 3093 | Từ chối trước SAP; không gắn nhãn 3093 giả | Parse → dispatch |
| NVL06 | Hỏi batch B1, MB52 chỉ có B2 | Không cộng B2 cho B1; filter chưa đủ thì unknown | Scope/batch fallback |
| NVL07 | Chỉ QI 100 PC | Eligible = 0, QC = 100; không báo sẵn sàng | Domain + KPI |
| NVL08 | Category rỗng/UNKNOWN, quantity100 | UNKNOWN, không FT/eligible | Domain |
| NVL09 | Bin blocked với quantity dương | BLOCKED, không rank cấp phát | Domain |
| NVL10 | Product rỗng hoặc BIW...AA/AB gần giống | Không ghép bằng substring/empty | Resolver/domain |
| NVL11 | POST vượt trần/thiếu category hoặc timestamp | PARTIAL/UNKNOWN tương ứng; không FIFO toàn tập | Adapter + domain |
| NVL12 | Có cả 0, QC, blocked, eligible và quantity không hợp lệ | Metrics tách rõ; phần chưa xác định không bị giấu | End-to-end offline |

### 19.2. FIFO và đề xuất phân bổ — FIFO01–FIFO10

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| FIFO01 | 2025-12-31 và 2026-01-02 dạng ISO | Ngày 2025-12-31 trước |
| FIFO02 | dd.MM, dd/MM và ISO cùng ngày | Cùng timestamp chuẩn |
| FIFO03 | Thiếu ngày, ngày invalid, ngày hai năm mơ hồ | Không đẩy lên #1; reason rõ |
| FIFO04 | Giờ8:00 và09:00 | 08:00 trước, không sort text sai |
| FIFO05 | Cùng timestamp, nhiều bin/batch | Tie-break ổn định, không đổi thứ tự vô cớ |
| FIFO06 | Dòng QC cũ hơn dòng eligible | Không lấy QC vào allocation |
| FIFO07 | Demand bằng 0, âm, NaN hoặc thiếu UoM | Từ chối dữ liệu không hợp lệ; chỉ tra tồn nếu không cung cấp demand |
| FIFO08 | Demand đúng bằng tổng eligible | Tổng allocation = demand; shortage = 0 |
| FIFO09 | Demand lớn hơn tồn, nhiều batch | Không vượt từng dòng; shortage chính xác |
| FIFO10 | Pack rounding/batch rule/stock đổi sau snapshot | Phần dư/lý do rõ, đề xuất không được xem là đã giữ chỗ |

### 19.3. Số liệu và định danh — NUM01–NUM10

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| NUM01 | Numeric Excel 1000 và text US 1,000 | Canonical = 1000, không parse display lần hai |
| NUM02 | EU 1.000, US 1.000, EU 1,234 | Theo locale đã chốt; mơ hồ không tự chọn |
| NUM03 | 100.00-, -100.00, số lẻ và quantity0 | Dấu/precision/zero đúng |
| NUM04 | Chuỗi lỗi, None, bool, NaN, Infinity | Invalid/unknown theo contract; không0 giả |
| NUM05 | A.Qty alternate với alternate UoM khác base | Đúng cặp qty/unit; không dùng base unit sai |
| NUM06 | PC và KG trong cùng material/report | Tách metrics; không tổng chung |
| NUM07 | Conversion factor thiếu/0/chưa verified | Không quy đổi/netting |
| NUM08 | Material có suffix AA/AB, PO/item có zero đầu | Giữ định danh và không gộp |
| NUM09 | Short code ra hai full codes | Ambiguous/candidates tách rõ |
| NUM10 | Header đổi, cột nghĩa gần giống, qty và totalstock cùng có | Chọn đúng technical column hoặc báo mơ hồ |

### 19.4. FG/OB và tồn thấp — FG01–FG09

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| FG01 | MB52 dương, VL06O lỗi | Nhu cầu OB chưa biết; không báo “không OB/sẵn sàng” |
| FG02 | MB52 lỗi, VL06O có nhu cầu | Stock chưa biết; không báo thiếu hàng dựa vào stock = 0 |
| FG03 | Cả hai lỗi | FAILED, không số dư giả |
| FG04 | Stock = 100, remaining OB = 150 | signedBalance = -50; shortage = 50 |
| FG05 | OB partial, nhiều item/batch split | Dùng đúng remaining và không double-count |
| FG06 | Dòng ALV duplicate/subtotal/total | Không cộng trùng; rejectedRows trace được |
| FG07 | Hai nguồn khác UoM/scope hoặc timestamp quá xa | Không netting như snapshot chính xác |
| FG08 | TON_THAP master/model rỗng hoặc thiếu mã | Không query rộng/thêm số0 tự động |
| FG09 | TON_THAP có nguồn partial và yêu cầu Top N | Ghi giới hạn, không nói “Top N toàn bộ” |

### 19.5. GR/IB/PO — DOC01–DOC10

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| DOC01 | Reference có hơn 9 ký tự hoặc zero đầu | Giữ raw, không cắt hoặc thêm 0 tùy ý |
| DOC02 | Cùng material document khác năm/item | Trả/tách đúng khóa, hỏi rõ khi thiếu |
| DOC03 | Nhiều GR cùng IB và có reversal/return | Không lấy GR đầu làm kết luận nhận đủ |
| DOC04 | Posting date khác document date hoặc ngày GR của stock | Nhãn thời gian đúng nguồn |
| DOC05 | IB có nhiều item cùng material, expectedQty là tổng | So đúng comparisonScope, không so tổng lặp mỗi dòng |
| DOC06 | IB có status trống hoặc lạ | UNKNOWN, không gán chưa GR giả |
| DOC07 | Một PO giao nhiều IB, quantity lệch | Candidates đầy đủ, delta đúng; không báo sai IB chỉ vì quantity |
| DOC08 | PO có release code chưa có mapping | Hiện raw + unknown, không gán đã duyệt giả |
| DOC09 | MGLIEF/layout khác nghĩa; completed/blocked/deleted | Không suy openQty khi contract chưa đủ |
| DOC10 | Handler tra cứu bị gọi với input thiếu | Từ chối trước COM; không tự tạo/sửa/post |

### 19.6. Parser và điều phối — CMD01–CMD09

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| CMD01 | Toàn canonical commands hiện hành | Đúng handler/schema; không fall-through |
| CMD02 | Mọi alias JSON/natural có registry | Cùng canonical request và entities |
| CMD03 | PO/IB/document số dài trong câu có nhãn | Không đưa vào MaterialResolver |
| CMD04 | HELP/PING/thiếu material/thư không liên quan | Phản hồi đúng/ngừng theo policy, không query ngoài chủ đích |
| CMD05 | limit/days/plants sai kiểu hoặc bounds | Lỗi validation rõ; thư sau vẫn xử lý |
| CMD06 | dryRun dạng chuỗi false/true và mode ngoài danh sách | Parse rõ, không tự chuyển sang LIVE |
| CMD07 | PARSER_ONLY/OFFLINE_RESULT_TEST | Không SAP, không HTTP, không DB production |
| CMD08 | “ngày GR của stock” và “chứng từ GR của inbound” | VITRI và TIM_GR_DOC đúng intent |
| CMD09 | VITRI theo bin khi adapter chưa hỗ trợ | Không hứa đã lọc bin; từ chối hoặc route đúng contract |

### 19.7. Phục hồi và Outbox — REC01–REC15

| ID | Failpoint/tình huống | Kết quả bắt buộc sau resume |
|---|---|---|
| REC01 | Sau insert request, trước query | Request còn, không nhân bản |
| REC02 | Sau query, trước commit result | Retry query chỉ đọc có attempt/snapshot mới; không giả giữ snapshot cũ |
| REC03 | Giữa save result/status/enqueue | Transaction all-or-nothing |
| REC04 | Sau commit, trước HTTP | Chỉ gửi payload đã lưu, không query lại |
| REC05 | Receiver đã nhận, client timeout | UNCERTAIN/reconcile hoặc retry idempotent đã test |
| REC06 | Sau 2xx, trước ghi accepted | Không tự chắc chưa gửi; phục hồi có policy cho trạng thái chưa rõ |
| REC07 | Sau accepted, trước move mail | Retry archive, không query/send mới |
| REC08 | Move thành công, chưa update locator | Tìm theo stable key; không tạo duplicate vì EntryID đã đổi |
| REC09 | Restart nhiều lần và START hai lần | Một active instance, không claim hai lần |
| REC10 | SQLite locked/disk full | Không mất yêu cầu; không gửi payload chưa lưu bền vững |
| REC11 | HTTP 429/5xx/401/400 và Retry-After | Đúng classification/backoff/review |
| REC12 | JSON payload hỏng, hash khác hoặc destination đổi | Quarantine/review; không gửi {} hoặc đổi đích ngầm |
| REC13 | Hết retry, operator requeue | Payload/IDs giữ nguyên, có audit, không tra SAP |
| REC14 | Cùng source key khác inputHash | Conflict; không REPLACE |
| REC15 | SAP/Outlook đều chưa mở nhưng có Outbox | Sender hoạt động độc lập |

### 19.8. Session và môi trường — ENV01–ENV08

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| ENV01 | Nhiều system/client/user/session | Chọn đúng cấu hình hoặc từ chối |
| ENV02 | User đổi T-code sau attach | Không reset/ghi/query vào phiên bị chiếm |
| ENV03 | SAP Busy, COM exception hoặc popup lạ | Timeout rõ; không Enter/save mù |
| ENV04 | Thiếu control filter hoặc multi-select nhập thiếu | FILTER_NOT_APPLIED/PARTIAL; không trả scope giả |
| ENV05 | Hai mailbox có folder SAP_BOT | Đúng store được cấu hình |
| ENV06 | Exchange sender chưa resolve, trống hoặc domain giả | Không vượt allowlist |
| ENV07 | STOP khi Python khác cũng có main.py | Chỉ worker đúng bị dừng |
| ENV08 | Clipboard text/image/file và lỗi giữa thao tác | Chính sách bảo toàn/giới hạn rõ; clipboard không bị khóa |

### 19.9. Teams và phát hành — REL01–REL11

| ID | Tình huống | Kết quả bắt buộc |
|---|---|---|
| REL01 | Có dữ liệu/rỗng/partial/error/HELP/PING | Greeting đúng ngữ cảnh, không lặp hai lời chào |
| REL02 | Unicode/custom HTML/summary/metadata rất lớn | Đúng budget byte của toàn payload; thông báo cắt đúng |
| REL03 | Cell/error/sender chứa HTML | Escape đúng, không có markup ngoài chủ đích |
| REL04 | HTTP accepted nhưng action Teams trong flow lỗi | Không ghi TEAMS_DELIVERED |
| REL05 | Receiver nhận duplicate đồng thời hoặc crash sau Teams post | Atomic claim/reconcile; không cam kết exactly-once vô căn cứ |
| REL06 | Patch vào máy có config/master/DB | Giữ nguyên file người dùng, merge cấu hình có kiểm tra |
| REL07 | Migration lặp, legacy PENDING/FAILED/orphan/collision | Idempotent, không phát lại lịch sử, bảo toàn payload |
| REL08 | Quay lui source sau request A/B, nhận C | Không mất B; queue/correlation còn đúng |
| REL09 | Giải nén portable trên máy không có Python hệ thống | Import/fixture test và live smoke bằng bundled runtime |
| REL10 | Gói Standard/Patch/Portable | Không có secret/DB/cache/log/ảnh/test ngoài gói dev |
| REL11 | Đường dẫn có dấu/khoảng trắng, silent start/stop | Đúng cwd/runtime/PID, không có vòng lặp start ẩn |

### 19.10. Cấp độ bằng chứng

1. **Unit/domain:** data fixture nhỏ, xác minh phép tính/identity/status/decision boundaries.
2. **Adapter contract:** mock COM có signature/filter/cột/busy/popup và export mẫu đã ẩn dữ liệu; không chỉ stub trả records hợp lệ.
3. **Orchestration:** parse → registry → query → persist → dispatch → archive với mock ngoài hệ thống; kiểm cả thứ tự và lỗi giữa bước.
4. **Persistence integration:** SQLite thật tạm thời, crash subprocess/failpoint có kiểm soát, migration/restart nhiều lần.
5. **Bundled runtime:** chạy chính binary portable trên bản giải nén, không dùng test pass của Python hệ thống thay thế.
6. **SAP golden cases:** tài khoản/session được chỉ định, data biết trước, so từng dòng, scope, quantity/unit/status và giới hạn.
7. **Flow/Teams integration:** receiver test và nhóm đích thích hợp, quan sát action/run/message; không test live bằng số liệu giả gây hiểu nhầm vận hành.

Mỗi level ghi PASS/FAIL/NOT_RUN và lý do. Không đặt mục tiêu số lượng test để chứng minh ổn định; mục tiêu là đóng các case quan trọng và không phá hồi quy cũ.

## 20. Checklist nghiệm thu và bàn giao

### 20.1. Checklist cho mỗi thay đổi

- [ ] Có risk/task ID, phạm vi file và baseline/diff trước sửa.
- [ ] Có test tái hiện lỗi trước và kiểm đúng sau sửa, không chỉ mirror implementation.
- [ ] Quy tắc/cột/scope đổi có nguồn và quyết định nghiệp vụ ghi nhận.
- [ ] Không overwrite thay đổi đang có của người dùng/người khác.
- [ ] Không thêm dependency/runtime/package ngoài nhu cầu.
- [ ] Tác động payload/DB/config/backward compatibility được đánh giá.
- [ ] Kết quả test ghi đúng môi trường/version; hạn chế NOT_RUN rõ.
- [ ] Không chạy hoặc kích hoạt thao tác SAP ghi thật ngoài phạm vi.

### 20.2. Checklist đóng M1

- [ ] NVL không còn fake-zero, FT default hoặc QC-as-available.
- [ ] Scope/batch/material/UoM kiểm chứng được; mapping chưa đủ thì kết quả UNKNOWN.
- [ ] FIFO đúng ISO/formatnguồn, timestamp thiếu không được ưu tiên mù.
- [ ] FG không coi source lỗi là noOB/noStock; metrics thiếu/thừa rõ.
- [ ] Numeric chỉ format ở render; alias và required input của parser/handler đúng.
- [ ] Guard session tối thiểu và payload compatibility pass.
- [ ] Toàn suite cũ và case mới liên quan pass; golden cases trong phạm vi phát hành đã chạy.
- [ ] Patch sạch giữ config/master/DB; có rollback source cùng schema.

### 20.3. Checklist đóng M2

- [ ] UUID/stable source key/result immutable và atomic enqueue.
- [ ] Delivery/query/archive status tách biệt; timeout chưa rõ có policy.
- [ ] Receiver idempotency/reconcile đã test hoặc UNCERTAIN giữ review; không cam kết không trùng khi thiếu bằng chứng.
- [ ] Sender không phụ thuộc SAP/Outlook startup.
- [ ] Single instance/STOP/leases/requeue/crash cases pass.
- [ ] Migration idempotent bảo toàn pending/lịch sử và có compatibility fallback.
- [ ] Retention/log/health không làm mất dedupe hoặc token.

### 20.4. Checklist đóng M3/M4

- [ ] GR/IB/PO semantics được xác minh; không đổi unknown thành approved/chưa GR.
- [ ] Đề xuất allocation đúng từng dòng/pack/UoM và không giả giữ chỗ/post.
- [ ] Số đo trước/sau so cùng scope; không nhanh lên do đọc thiếu.
- [ ] Runtime portable và cách cập nhật config/master/state đã được kiểm thử.
- [ ] Runbook đủ để xem queue/requeue/backup/stop/start và phân biệt query mới/gửi lại.
- [ ] Changelog ghi hạn chế còn lại; chỉ ghi đã duyệt khi có người duyệt thực tế.

### 20.5. Hồ sơ bàn giao dự kiến

Manifest và checksums; changelog; config.example/schema; migration/recovery guide; command examples; bảng mapping cột/status/scope đã xác minh; test report theo level; số đo; runbook; danh sách NOT_RUN/open decisions; cách rollback/fallback.

Không đưa log nguyên văn chứa secret hoặc snapshot vận hành có dữ liệu riêng vào release sạch. Evidence chi tiết lưu theo quyền truy cập phù hợp.

## 21. Mở rộng sau khi nền tảng ổn định

### 21.1. Phân biệt đã có và chưa có

- TIM_GR_DOC/CHECK_IB/TIM_IB/CHECK_PO/CHECK_FG_OB/CHECK_NVL_MIGO đã có code; nâng cấp là hardening và đối chiếu.
- QUEUE_IB, SHIPMENT_INFO và PRECHECK_GR chưa được coi đã hoàn tất chỉ vì cùng dùng MON/ME2N hoặc có prototype VBA.
- R137 chỉ thấy macro export, chưa đủ mô tả KPI; cần sample/schema/scope trước.

### 21.2. Mở rộng chỉ đọc dự kiến

| Lệnh | Phạm vi | Gate trước kích hoạt |
|---|---|---|
| QUEUE_IB | Đọc queue errors theo IB/warehouse, không đổi WO queue | Node/warehouse/quyền đúng, 0/1/n dòng; không mặc định no queue khi lỗi |
| SHIPMENT_INFO | ZMME240 bill/material/PO/quantity/status nguồn | Report tồn tại, quyền/cột/ý nghĩa TRXID1 được xác minh |
| PRECHECK_GR | Tổng hợp IB/PO/queue/GR-doc, từng nguồn có thời điểm/trạng thái | Các adapter độc lập đạt gate; không tự post hoặc báo mọi điều kiện OK khi thiếu dữ liệu |
| Báo cáo R137 | Export/query theo thiết kế sau khảo sát | Có định nghĩa cột, filter, KPI và golden case |

### 21.3. Tự động ghi SAP là đề án riêng

Nếu tương lai muốn tạo MIGO 311, post GR, sửa/create/delete IB/batch/bin/reservation/complete OB:

- Tách readOnly query và executor bằng command/schema/quyền rõ, không alias mơ hồ.
- Preview đúng chứng từ/item/số lượng/kho/batch/trước-sau, xác nhận tác vụ cụ thể theo người có quyền.
- Revalidate stock/document status ngay trước ghi; snapshot báo cáo không là reservation.
- Idempotency/recovery dựa chứng từ SAP thực tế. Timeout sau post phải kiểm tra chứng từ trước bất kỳ retry nào.
- Outbox retry chỉ gửi thông báo, không gọi executor lại.
- Test và đối chiếu trên môi trường phù hợp trước dùng thật. Không bê menu Yes/No hoặc credential hardcode VBA sang bot.

Phần này không cấp quyền thực hiện post chỉ từ việc người dùng yêu cầu một báo cáo tra cứu.

## 22. Quyết định nghiệp vụ cần xác minh

Những mục này là đầu vào triển khai, không phải yêu cầu trả lời ngay để hoàn thành việc lập kế hoạch. Có thể chuẩn bị code/test mock không phụ thuộc câu trả lời; các rule chưa xác minh giữ UNKNOWN/feature-off khi chạy thật.

| ID | Cần chốt | Bằng chứng cần thu | Khi chưa chốt |
|---|---|---|---|
| D01 | Mapping system/client/warehouse/owner → NVL 1005 | Màn hình/cột stock/report đối chiếu với mẫu có kết quả biết trước | Scope UNKNOWN, không gán 1005 như fact |
| D02 | Allowlist stock category và nguồn bin lock | Giá trị thực và ý nghĩa nguồn theo môi trường | Category/lock chưa rõ → UNKNOWN |
| D03 | A.Qty AUoM, base/alternate unit và conversion | Tên cột, cặp qty-unit và hệ số trong master | Không cộng hoặc đổi unit |
| D04 | Format/timezone GR timestamp và quy tắc FIFO/pack | Mẫu ISO/dd.MM/giờ/đồng hạng | Không rank timestamp mơ hồ |
| D05 | Định nghĩa remaining OB và stock đã trừ reservation hay chưa | Trường nguồn, state A/B/C và đối chiếu OB nhận một phần | Không gọi số dư là ATP |
| D06 | Quantity PO/SA, MGLIEF và release/deletion/completion | Layout/loại chứng từ và golden cases | Hiện raw, không kết luận còn GR/đã duyệt |
| D07 | ExpectedQty so theo item/material/document | Cách người dùng nhập và ví dụ nhiều item | Hỏi rõ comparisonScope |
| D08 | IB → reference, movement/reversal/return | Chứng từ liên quan đủ năm/item và quy tắc nguồn | Chỉ trả references đã quan sát |
| D09 | Session riêng, mailbox/store và transport sender | Identity/mapping thật, không credential trong source | Không tự fallback vào phiên người dùng |
| D10 | Receiver idempotency/ACK có store/capability nào | Flow definition/schema, run evidence và quyền hiện có | UNCERTAIN/review, không cam kết exactly-once |
| D11 | Retention/SLO/batch/timeout/query bounds | Volume request, độ nhạy dữ liệu và baseline SAP | Default tương thích có giới hạn, công bố giới hạn |
| D12 | Kho đích/pack/rounding và nhu cầu allocation | Ví dụ thực tế đủ/thiếu/lẻ/nhiều đơn vị | Chỉ trả tồn/FIFO tham khảo |
| D13 | Trigger Teams và greeting flow hiện dùng | Đọc cả flow intake và flow reply | Không đổi trigger hoặc khẳng định đã đồng bộ |
| D14 | Phạm vi cho phép ZMME240/R137/queue | Report/quyền/cột và mẫu có kết quả biết trước | Feature off, không tự suy dữ liệu |

## 23. Backlog theo file và thứ tự triển khai

Các đường dẫn dưới đây là tương đối với project; file “dự kiến” chỉ tạo khi bắt đầu implement. Người phụ trách chưa được phân công; không ghi nhận người/nhóm đã QA từ tài liệu kế hoạch.

| Task | Mốc | File/chức năng | Đầu ra | Phụ thuộc |
|---|---|---|---|---|
| T00 | M0 | Source/gói hiện có | Manifest, diff, chiến lược backup; xác nhận baseline 103 tests | Bắt đầu |
| T01 | M0 | Fixture/golden case registry dự kiến | Cột/scope/UoM/status mẫu, decision log D01–D14 | T00 |
| T02 | M1 | models.py dự kiến | SourceResult/BusinessResult/Quantity/Scope và serializer | T01, phần contract |
| T03 | M1 | number_utils.py, adapter read/export | Numeric canonical, cặp quantity-unit, validation | T02 |
| T04 | M1 | sap_engine POST/MB52/VL06O | Filter bắt buộc, grid/cột, completeness, lỗi typed | T02/T03 |
| T05 | M1 | domain_nvl.py dự kiến + NVL handler | Ma trận fallback nguồn, batch/scope/eligibility/FIFO | T03/T04/D01–D04 |
| T06 | M1 | domain_fg.py dự kiến + FG handler | Ma trận lỗi nguồn và balance/shortage | T03/T04/D05 |
| T07 | M1 | teams_dispatcher.py | KPI typed, giữ zero đúng, partial/greeting, budget | T05/T06 |
| T08 | M1 | commands.py dự kiến/parser/main | Registry, validation, alias parity, guard không có handler | T02 |
| T09 | M1→M2 | SAP session manager/main | Ownership guard trước COM, lazy session, timeout | T01/D09 |
| T10 | M1 | package_release.py | Allowlist bảo vệ config, output directory, hotfix manifest | T00 |
| T11 | M1 | Tests domain/dispatch/fixtures dự kiến | NVL/FG/NUM/CMD cases, suite cũ pass | T03–T10 |
| T12 | M2 | state_manager.py/migrations dự kiến | Schema, atomic result/outbox, stable IDs | Result contract của M1 |
| T13 | M2 | Worker/sender dự kiến | Claim/send/attempt/retry/requeue/archive | T12 |
| T14 | M2 | Instance/lifecycle + START/STOP/SILENT | Một worker, đúng PID, recovery | T09/T12 |
| T15 | M2 | Power Automate receiver riêng | Contract idempotency/uncertain/ACK trong sandbox | T12/T13/D10 |
| T16 | M2 | Migration + compatibility fallback dự kiến | Import idempotent, diễn tập quay lui không restore DB | T12–T15 |
| T17 | M2 | Logging/health/queue CLI dự kiến | Xem state/requeue/runbook, che secret | T12/T13 |
| T18 | M2 | Tests recovery/session/runtime dự kiến | REC/ENV/REL migration cases pass | T12–T17 |
| T19 | M3 | MB51/VL06I/ME2N handlers | GR/year/item/reversal, expectedQty scope, PO semantics | T03/T04/D06–D08 |
| T20 | M3 | Allocation domain dự kiến | Demand → đề xuất quantity theo lot/bin và shortage | T05/D12 |
| T21 | M3 | Test và golden cases | DOC/FIFO allocation, regression và bằng chứng live | T19/T20 |
| T22 | M4 | Profiler/benchmark harness dự kiến | Baseline thời gian/RAM/COM/dung lượng và report | G1/G3 đã pass |
| T23 | M4 | Outlook/resolver/extractor/render | Batch/streaming/cache policy, tối ưu có số đo | T22 |
| T24 | M4 | Portable build/update/runbook | Runtime đúng, gói sạch, cài mới/upgrade/fallback | T10/T16/T23 |
| T25 | M5 | Queue/Z-report/PRECHECK_GR dự kiến | Chức năng chỉ đọc mới theo contract | D14 và gate adapter |

### 23.1. Thứ tự thực hiện đề xuất

1. T00–T02: đóng baseline và contract; có thể viết test mock trước khi xác minh SAP thật.
2. T03/T04: numeric và trạng thái nguồn trước mọi khuyến nghị.
3. T05/T06/T07: NVL/FG/KPI. T08/T09/T10 có thể là workstream riêng, nhưng cùng chờ G1/G2 trước pilot.
4. T11: unit/contract/dispatch và golden cases cho M1; không phát hành khuyến nghị stock nếu mapping chưa chốt.
5. T12–T18: schema/recovery/session/receiver có mốc tương thích và failpoint trước migration live.
6. T19–T21: hoàn thiện GR/IB/PO và allocation; tắt feature khi thiếu điều kiện nghiệp vụ.
7. T22–T24: tối ưu hệ thống/gói; kiểm lại các case bị ảnh hưởng và bundled runtime.
8. T25: mở rộng nguồn mới riêng; không trộn với sửa enum/scope hoặc gói phát hành trong cùng một đợt lớn.

Phân việc song song chỉ cho phần độc lập có giao diện đã chốt. Khi cùng sửa một file như sap_engine/main, chia commit/file ownership rõ để tránh ghi đè; reviewer kiểm trạng thái source mới nhất trước tích hợp.

### 23.2. Cách theo dõi thực hiện

Mỗi task có trạng thái NOT_STARTED/IN_PROGRESS/IMPLEMENTED/TESTED/BUSINESS_VERIFIED/RELEASED và evidence tương ứng. IMPLEMENTED không tự nâng lên TESTED; TESTED không tự thành BUSINESS_VERIFIED. Phiên bản ứng dụng không thay thế trạng thái task/gate.

Khi có user steering hoặc source mới, lập diff so baseline, cập nhật phần đã làm và tái đánh giá phạm vi; không tiếp tục checklist cũ như thể source không đổi.

## 24. Manifest baseline và chứng cứ

### 24.1. SHA-256 source tại lúc lập kế hoạch

| File | SHA-256 |
|---|---|
| main.py | `177ee93e2e487e5b89c589dfac59b0eb397fcd05b06a28c1208c6208e319b4b1` |
| app/sap_engine.py | `9456adb2a393b93cd157e970dbfd632097ada0aa5b61e8c5e5877df78f5f5647` |
| app/email_parser.py | `0052ce388688a41a96cebf973c1b0181a41bf43fb55c31968ad6d91f20bdfb52` |
| app/teams_dispatcher.py | `737cd95d32ba6ad80875e681fa839ec4767bf67b5b0bf72e23fa0941d95395a2` |
| app/state_manager.py | `34e6d63d2a0c7e1f1c3c73410803ecdd3e45fbcf0a6c40b1ae810303752f5663` |
| app/number_utils.py | `2eea1b86c4cbee835ea48629706d16ba5811e001fcf3846ef6a1d3c77efe3aea` |
| app/material_resolver.py | `e537b689f22274c11c487478f7588fa394c996e0dc19e9d29265cc21eecd33c0` |
| package_release.py | `593d0ddb17f8cab6c00fb32852a6788175cd7f6642d4cca850afb97e1cd9e053` |

Nếu source hash đã khác khi triển khai, phải kiểm diff và cập nhật baseline; không áp blind patch theo số dòng trong tài liệu.

### 24.2. Tài liệu liên quan

- `docs/KE_HOACH_NANG_CAP_SAU_v1.0.15.md`: lịch sử kế hoạch, không phải bằng chứng hoàn tất các mốc đã đặt tên.
- `docs/PHAN_TICH_MA_LOG_GR_V7_1_VA_MO_RONG.md`: nguồn VBA và phạm vi đọc/ghi; một số lệnh đề xuất trong đó hiện đã có code, xem inventory mục2 để không triển khai trùng.
- `docs/TONG_KET_LUONG_NVL_MIGO_VA_FG_OB.md`: mục tiêu nghiệp vụ NVL/FG hiện có; các lời “sẵn sàng/chuẩn xác” cần được đối chiếu với contract/gates trong kế hoạch này.
- Nội dung release note được người dùng gửi là đầu vào phân tích. Sự kiện commit/push/GitHub release trong đó chưa được tài liệu này kiểm chứng từ remote; không dựa vào chúng để suy QA/live acceptance.

### 24.3. Những gì đã và chưa làm trong lượt lập kế hoạch

Đã đọc source/tài liệu liên quan; kiểm lại 103 tests; so source trong ZIP v1.0.18; ghi baseline hash; tổng hợp các probe mock lỗi đã tái hiện và thiết kế nâng cấp.

Chưa bắt đầu các task triển khai code trong tài liệu, migrate/restore DB, đổi config, chạy worker/SAP/Outlook, sửa flow Power Automate hoặc gửi tin Teams trong lượt này. Không có migration/CLI/file code mới nào tồn tại chỉ vì tên của nó xuất hiện trong thiết kế.

Mục tiêu nghiệm thu là các bảo đảm cụ thể có bằng chứng về độ đúng, phục hồi và phạm vi; không tuyên bố hệ thống “hoàn hảo” chỉ từ số tính năng hoặc tổng test pass.
