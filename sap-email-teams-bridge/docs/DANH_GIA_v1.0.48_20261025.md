# Đánh giá ba gói SAP Email Teams Bridge v1.0.48

Ngày kiểm tra: 25/10/2026, Asia/Saigon. Phạm vi: ba ZIP trong Downloads, source trích xuất riêng và các phép kiểm cô lập. Chưa cập nhật bản đang chạy, chưa gọi SAP/Outlook hoặc gửi webhook thật.

## 1. Kết luận

v1.0.48 có thêm nhiều chức năng và một số sửa lỗi tốt so với v1.0.18. Ba gói dùng source khớp nhau. Tuy nhiên, chưa nên bật tự động các nhánh ghi SAP/điều phối hoặc dùng BBBG/ASN làm kết luận nghiệp vụ trước khi sửa các lỗi P1 dưới đây.

Điểm chặn quan trọng nhất là ranh giới đọc/ghi: JSON tra cứu ASN có thể bị parser nhận thành ATT_MIGO; ATT_MIGO thực sự có thể lưu Header Text và upload GOS. Đồng thời, một số nhánh mô phỏng lại xuất trạng thái như đã thực thi/readback thật.

188/188 test có sẵn pass trong môi trường cô lập. Các probe bổ sung vẫn tái hiện được nhiều lỗi ngoài bộ test; kết quả test hiện có không thay thế kiểm chứng nghiệp vụ/live.

## 2. Gói và kiểm tra runtime

| Gói | Bytes | Dung lượng nhị phân gần đúng | File | Nhận xét |
|---|---:|---:|---:|---|
| Portable | 13.993.589 | 13,35 MiB | 1.401 | Có 1.354 file runtime; giải nén tổng khoảng 40,38 MB |
| Standard | 343.073 | 335,03 KiB | 47 | Source, master, config, docs và test |
| Update Patch | 165.292 | 161,42 KiB | 32 | Không có config.json và DB; vẫn cần kiểm dependency khi áp cho bản cũ |

- Source chung trong ba gói khớp nhau và khớp source dự án ở thời điểm lấy snapshot.
- Không có DB/log vận hành trong ba ZIP được kiểm tra.
- Standard/Portable có config.json với URL webhook đã cấu hình, không phải placeholder; không chép URL/token vào báo cáo. `allowed_senders` trong config là danh sách rỗng, `prefer_dedicated_session` là false.
- Runtime portable khởi động bằng Python 3.12.10. python.exe có chữ ký hợp lệ của Python Software Foundation tại lúc kiểm tra.
- Smoke test bằng chính runtime portable import được sqlite3, ssl, win32com, pythoncom, openpyxl và các service chính. Không có kết nối SAP trong phép kiểm này.
- xlrd không có trong ZIP. Nhánh fallback đọc XLS bằng xlrd chưa được hỗ trợ đầy đủ bởi runtime này; đây là dependency tùy chọn cần xác minh theo export thực tế, không phải lỗi import core đã quan sát.
- Portable giúp không cần cài Python; vẫn cần SAP GUI/GUI Scripting, Outlook Classic và các quyền/môi trường phù hợp.

SHA-256:

```text
PORTABLE     7f8573c30671c52c2f2450e28d02a29723fc7ab1e19f845e33c32991833857be
RELEASE      bcd6011242927007c34bf1998dbd5251b0e0955f521f8036e859fced6274b774
UPDATE_PATCH a21603d4417656c7ca7c27e75fd7ec14966e877add55d7b912f0a3041483e500
```

## 3. Phạm vi chức năng thực tế

| Nhóm | Quan sát từ source |
|---|---|
| Tra cứu cũ và GR/IB/PO | Tiếp tục có handler; vẫn cần xác minh scope/cột/completeness và trạng thái nguồn |
| BOM/queue | Có adapter CS03/MON và lớp service; một số đường xử lý chưa khớp extractor hoặc chưa dùng service |
| BBBG | Đối soát MB51/VL06I, tạo XLSX local và HTML preview; chưa có luồng giao file/email hoàn chỉnh |
| NVL allocation | Có phân bổ FEFO/FIFO, plan và reservation SQLite cục bộ; không tương đương giữ chỗ SAP |
| Điều phối OB | Main gọi hàm simulated; chưa phải executor/readback SAP |
| ATT_MIGO | Có ghi thật: mở MIGO Change, lưu Header Text, upload GOS. Đây là thay đổi phạm vi quan trọng so với bot chỉ tra cứu |
| ASN | Tổng hợp từ VL06I; chưa đủ chứng cứ có nguồn container/customs/ZMME240 thật |
| Audit/UI cache | Có bảng audit và cache control ID; chưa bảo đảm immutable ledger, quyền sở hữu session hoặc freshness dữ liệu kho |

## 4. Lỗi chặn vận hành ưu tiên P1

### F01. JSON tra cứu có thể route sang thao tác ghi

**Đã tái hiện bằng parser thuần:**

```json
{"command":"TRA_CUU_ASN","inboundNo":"1800000001"}
```

Trả `valid=true`, `command=ATT_MIGO`, doc bằng số inbound và ticket bị nhận thành đoạn kết thúc JSON. Không có thao tác SAP nào được thực hiện trong probe.

Nguyên nhân: nhóm command không có materials trong `app/email_parser.py:358` chưa bao phủ các command mới; JSON rơi xuống parser tự nhiên và heuristic attachment ở khoảng 658–662. JSON ATT_MIGO cũng có thể bị đọc sai attachmentItems.

**Cần sửa:** JSON hợp lệ phải được dispatch/validate theo registry, không fallback sang suy đoán lệnh ghi. Command không hỗ trợ hoặc thiếu entity phải fail closed. Ranh giới tra cứu/ghi có guard độc lập sau parse.

### F02. Thiếu session vẫn báo ghi Header Text hoàn tất

`app/sap_attachment.py:1177`, 1240, 1303 gộp `dry_run` với `session is None` vào nhánh mô phỏng. Probe `process_single_migo_item(None, ..., dry_run=False)` trả `True`, `DONE`, `HEADTEXT_DONE`.

**Cần sửa:** thiếu session trong mode thật là lỗi, không mô phỏng thành công. Worker không kết nối được phải dừng batch/giữ yêu cầu; trạng thái simulated khác completed. Kiểm lại chứng từ+năm và đọc ngược Header Text/GOS sau ghi. Năm mặc định 2026 không phù hợp mọi chứng từ; một số success checks GOS còn có thể nhận chuỗi phủ định chứa CREATED thành thành công.

### F03. Mô phỏng OB bị ghi như đã thực thi/readback SAP

`main.py:733` gọi `ObService.dispatch_outbound_simulated()`, sau đó ghi audit DISPATCHED và đưa readback vào summary. `app/ob_service.py:150–173` tự gán C/CONFIRMED, không có SAP action/readback.

`NvlService.execute_migo_311_simulated()` còn sinh số chứng từ và ghi EXECUTED/CONSUMED trong DB. Chưa thấy main gọi hàm NVL này trong đường điều phối được kiểm tra; không được dùng các giá trị đó như chứng từ thật.

**Cần sửa:** cô lập simulation, dùng SIMULATED/PREVIEW và không sinh chứng từ/confirmation giả trong audit vận hành. Chỉ ghi thực thi khi có evidence SAP.

### F04. BBBG ghép nhầm và kết luận đã GR không đủ căn cứ

`app/bbbg_service.py:119–152` chọn dòng đầu cùng material+batch, bỏ qua IB/ref; fallback cùng IB không kiểm lại material/batch/unit. Quantity dùng abs nên reversal âm thành dương. Diff bằng 0 được gọi “Đã GR”.

Probe đã xác nhận: khác IB nhưng cùng material/batch vẫn MATCHED; cùng IB nhưng khác material/batch/unit vẫn có thể MATCHED; quantity -10 của reversal thành +10 rồi MATCHED. Main còn dùng danh sách IB treo chưa GR để đối soát.

**Cần sửa:** khóa document/year/item, IB/item, material/batch/UoM/reference đầy đủ; many-to-many và net movement/reversal có quy tắc. “Số lượng khớp” không tự đồng nghĩa “đã GR”.

### F05. ASN tự gán trạng thái thông quan và xử lý rỗng lỗi

`app/extended_query_service.py:161` mặc định customs_status thành “ĐÃ THÔNG QUAN” khi nguồn không có field. Probe đã xác nhận. Main 748 query VL06I theo materials, không truyền inboundNo làm filter; kết quả có thể không thuộc IB đang hỏi.

Khi không có dữ liệu, service trả found=false nhưng main 755 vẫn index container_no/bill_of_lading không tồn tại, dẫn đến KeyError.

**Cần sửa:** thiếu customs source → chưa xác định; query và verify đúng IB; dùng schema ổn định cho found=false; không tuyên bố container/customs từ VL06I nếu nguồn không cung cấp.

### F06. Phân bổ mới có thể cấp sai quantity/UoM hoặc vượt tồn

Các probe temp DB xác nhận:

- RA_DON_NVL JSON có qty=7, station=OP50 nhưng metadata bỏ cả hai; main 715 dùng mặc định 100/OP10.
- Stock QC 10 KG vẫn cấp demand 5 PC vì allocator không kiểm eligibility/UoM/scope đầy đủ.
- availQty=0 nhưng totalQty=10 vẫn phân bổ do dùng toán tử `or` bỏ mất zero.
- OB tồn một batch 10, dòng batch-specific cần 6 và dòng generic cần 6 vẫn gate-pass 12 vì hai bộ đếm không dùng chung pool.
- Demand âm bị abs thành dương; danh sách OB rỗng vẫn có thể gate-pass.

**Cần sửa:** giữ entity qua parser; quantity zero khác missing; validate demand; một stock ledger chung theo đủ material/batch/bin/plant/sloc/UoM/stock ID; chỉ phân bổ trên eligible stock và nguồn đầy đủ.

### F07. Reservation không idempotent, DB hardcode và crash có thể giữ tồn trùng

`app/nvl_service.py:33` mặc định DB ở G:/SAP/sap-email-teams-bridge/data/bridge_state.db, trong khi StateManager dùng data của bản giải nén. Main không truyền app-data path chung khi lập/lưu plan. Portable ở ổ khác có thể lỗi hoặc ghi vào DB của một bản cài khác.

Save cùng plan hai lần tạo reservation mới mỗi lần: probe qty10 → total20. Plan bị REPLACE nhưng reservation luôn INSERT; ID plan chỉ tới giây. Main 718 ghi reservation trước Outbox và processed marker, nên crash/replay có thể reserve lại.

**Cần sửa:** inject chung dataDir, stable request/intent key, atomic và idempotent plan/reservation/result, lifecycle expiry/release/version rõ. Local reservation không được mô tả như SAP đã giữ chỗ.

### F08. Lỗi NVL/số lượng còn sót và có regression

- Blank category đã thành UNKNOWN, nhưng category lạ có thể vẫn FIFO-ready; chưa có allowlist đầy đủ.
- Advice UNKNOWN có chữ “FIFO”; KPI dùng substring nên vẫn tính lượng đó là khả dụng dù rank “-”.
- POST rỗng hợp lệ + MB52 lỗi vẫn có đường trả hết hàng, chưa phân biệt nguồn cần thiết bị lỗi.
- Yêu cầu batch cụ thể nhưng record batch trống vẫn có thể được chấp nhận.
- AUTO parser hiện đọc `0.001` thành 1, `1.234` thành 1234 nhưng `1.000` thành 1; có thể sai nghiêm trọng với KG/BOM/số lẻ.
- Parse ngày đã xử lý ISO tốt hơn, nhưng chưa validate ngày thật/giờ thật; missing timestamp vẫn chưa được chặn eligibility đầy đủ.

**Cần sửa:** typed eligibility/metrics, SourceResult complete/partial/error, strict batch/scope, locale explicit và numeric xuyên suốt, datetime validation.

## 5. Những khoảng trống ổn định khác

1. **Outbox:** mã phiếu vẫn theo giờ+counter, REPLACE ghi đè, enqueue/send/delete/mark chưa cùng hợp đồng transaction; JSON hỏng có thể thành {}. Probe cùng req_id enqueue hai payload làm mất payload đầu. Chưa đóng các gate UUID, stable namespace, immutable result, accepted/delivered, requeue và recovery của kế hoạch trước.
2. **Session:** đã có Busy/CreateSession/prefer-dedicated, nhưng vẫn chọn theo index/T-code và quét nhiều connection; chưa xác minh system/client/user/ownership lại trước mọi query/reset. ATT multi-tab cũng cần ownership guard riêng.
3. **Worker:** SAP vẫn được khởi tạo trước xử lý Outbox; chưa thấy single-instance lock. STOP còn WMIC match rộng main.py/SAP_BOT. Mail archive đã được bắt lỗi riêng ở một số đường, là cải thiện cần giữ.
4. **Inbox:** mở rộng lấy 30 unread ở Inbox rồi match subject chứa SAP, có thể nhận nhầm ASAP hoặc bỏ request phía sau batch. Chỉ được move thư đã xác định thuộc intake bot.
5. **UI cache:** chỉ chia HIGH/LOW_SPEED và alias, không theo system/client/T-code/screen/type. Cached ID tồn tại được nhận ngay dù trỏ sai loại control. Save JSON không atomic, lỗi bị nuốt. Đây là cache UI ID, không bảo đảm độ mới dữ liệu kho.
6. **BOM/queue:** extractor chung bắt positional/document fields có thể loại hết row component/queue; CS03 chỉ đọc phần visible. “KHÔNG CÓ BOM” chứa “CÓ BOM” nên substring có thể tạo badge dương. Query queue lỗi không được biến thành queue sạch.
7. **Audit:** có table insert nhưng chưa bất biến theo nghĩa bảo đảm kỹ thuật, chưa atomic với mọi intent/side effect, identity trong body chưa là bằng chứng người thực hiện được xác thực.

## 6. Báo cáo và đóng gói

- BBBG export có template_path nhưng tạo Workbook mới; chưa giữ template như tham số gợi ý. Có hardcode thông tin phê duyệt/liên hệ trong report, không được đưa vào kết quả vận hành như đã xác nhận.
- Export không có dòng có nguy cơ công thức SUM tham chiếu chính ô tổng. Cần case rỗng và recalculation/visual QA trước dùng report chính thức.
- Main chỉ đưa tên file local vào summary; tên file trên máy worker không tự trở thành attachment/link tải được trên Teams. Preview xác nhận gửi chưa nối thành một workflow gửi batch hoàn chỉnh.
- Export BBBG cũng dùng đường dẫn G: cố định; cần app-data/config output directory chung.
- Patch đã bỏ config.json, là cải thiện rõ; nhưng không mang STOP/SILENT update, chỉ tham chiếu START_BACKGROUND có thể không tồn tại. Chưa có updater/manifest/dependency/migration gate đầy đủ.
- Standard/Portable còn config đã cấu hình, docs, ảnh, cache và test. Riêng ảnh docs/flow-reply-saved.png chiếm khoảng 80,6 KB sau nén.
- Build còn xóa ZIP cũ, quét cả cây và copy ra đường dẫn cá nhân; blacklist suffix chưa bao phủ WAL/SHM/journal. Chưa nên gọi packaging là đã sạch chỉ vì ZIP hiện không có DB.
- README vẫn ghi 54 tests và mô tả chức năng cũ, chưa khớp source/test suite hiện tại.

## 7. Những sửa tốt cần giữ

- Patch giữ config hiện có bằng cách không đóng gói config.json.
- FG/OB đã phân biệt nhiều trường hợp nguồn lỗi, giữ N/A thay vì giả đủ tồn như bản cũ.
- NVL từ chối plant không hỗ trợ ở đường đã kiểm tra, material match chặt hơn, có UNKNOWN category và parse ISO tốt hơn.
- Tách service NVL/OB/BBBG/extended, thêm test và bảng audit, session Busy/CreateSession, throttle sync/batch Inbox.
- Runtime Python/COM/SQLite/HTTPS/openpyxl import được trong bản portable đã kiểm tra.

Các cải thiện trên không khép các lỗi P1 của nhánh mới hoặc thay thế kiểm thử live.

## 8. Ưu tiên cho bản kế tiếp

1. Chặn route từ tra cứu sang ghi; JSON không fallback sang heuristic mutation. Cô lập ATT_MIGO bằng intent/quyền/preview/session guard rõ.
2. Chặn giả success khi session None; SIMULATED khác DONE/DISPATCHED/EXECUTED; chỉ báo chứng từ/readback khi có SAP evidence.
3. Sửa BBBG matching/reversal/UoM và ASN unknown; không dùng trạng thái hải quan/phê duyệt không có nguồn.
4. Sửa parser quantity/station, allocator eligibility, stock pool và reservation idempotency; inject chung dataDir.
5. Hoàn thiện numeric/source completeness, Outbox/recovery/session ownership, rồi mới tiếp tục mở rộng chức năng.
6. Sau correctness gates, cập nhật packaging/README/runbook và kiểm thử bundled runtime + golden cases trên SAP thật.

## 9. Phép kiểm đã thực hiện và giới hạn

- Trích xuất Standard vào thư mục tạm riêng, không unpack đè bản vận hành.
- So checksum source chung giữa ba ZIP và workspace tại lúc snapshot.
- Chạy 188/188 test hiện có bằng Python máy phát triển; chặn GetObject/Dispatch SAP/Outlook, urlopen và SQLite ngoài thư mục tạm. Kết quả suite này không bao gồm test SAP thật.
- Smoke import bằng Python portable của ZIP, không chạy main/run_bridge hoặc service ghi SAP.
- Probe parser/format/source services bằng dict/mock, và reservation/Outbox bằng SQLite tạm với db_path explicit. Không dùng DB G: mặc định trong các phép kiểm.
- Không tạo BBBG vận hành, không giả post SAP, không gửi file/email/Teams và không truy cập dữ liệu công ty từ mạng.
- Các lỗi đã mô tả không được sửa trong lượt phân tích này. Trước cập nhật đè cần regression cho từng lỗi, bản sao lưu source/config/master/DB và kiểm tra tương thích runtime/schema.
