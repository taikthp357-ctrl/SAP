# Kiểm tra độc lập SAP Email Teams Bridge v1.0.74

Ngày: 26/10/2026, Asia/Saigon. Source: commit `8296211`, các ZIP v1.0.74 tại dự án. Phạm vi lượt này là phân tích, không sửa mã ứng dụng, kích hoạt worker, ghi SAP hoặc gửi Teams.

## 1. Kết luận

**Chưa đủ điều kiện xác nhận ổn định cho các thao tác ghi SAP tự động.** Việc tách Router và 7 handler là cải thiện về tổ chức mã. Các lỗi ở ranh giới đọc/ghi, số lượng, bằng chứng thực thi và tích hợp vẫn còn; một số kiểm tra an toàn đã bị nới lỏng trong các bản sau v1.0.48.

Chạy lại toàn bộ suite trên bản Standard giải nén riêng: **304 passed in 69.91s**. Đã chặn COM SAP/Outlook, HTTP và SQLite ngoài thư mục tạm. Các probe bổ sung vẫn tái hiện lỗi trong bảng dưới. Test pass chưa chứng minh nghiệp vụ hoặc SAP GUI thực tế chạy đúng.

Không thể chấp nhận các phát biểu “Blast Radius = 0”, “độc lập 100%, không bao giờ ảnh hưởng” và “bảo toàn nghiệp vụ hoàn thiện 100%”. Các handler vẫn dùng chung parser, resolver, SAP session, cache, DB, payload và vòng worker. Router chưa có hợp đồng quyền ghi hoặc nguồn dữ liệu đầy đủ bắt buộc.

## 2. Các lỗi chặn ưu tiên

### P1-01. Lệnh kế hoạch NVL có thể ghi SAP dù cấu hình cấm ghi

- `app/handlers/nvl_handler.py:42` gộp `KE_HOACH_NVL` với `RA_DON_NVL`.
- `:81` mặc định `dryRun=False`; `:86` gọi `post_migo_311_ob_group(..., dry_run=False)` mà không kiểm `allow_sap_writes`, người gửi thực tế, xác nhận thực thi hoặc `is_full_allocation`.
- Probe: context `KE_HOACH_NVL`, `allow_sap_writes=False` vẫn gọi API executor ở chế độ thật. Executor trong probe là mock, không có tác động SAP.
- Probe yêu cầu 15, tồn 10: handler gửi 10 tới executor rồi trả `SUCCEEDED`. Kế hoạch bị BLOCKED cũng vẫn được gửi tới executor; với 0 dòng executor thật hiện có chặn rỗng, nên không suy diễn rằng ca rỗng đã ghi SAP.

**Cần sửa:** kế hoạch mặc định PREVIEW; guard ghi chung bắt buộc cho mọi handler, xác minh người gửi thực tế, chặn plan lỗi/thiếu tồn theo chính sách rõ ràng trước executor.

### P1-02. MIGO có thể báo thành công từ lỗi hoặc tiếp tục Post khi điền sai

- `app/sap_engine.py:3924` gộp `session is None` với mô phỏng và trả True dù `dry_run=False`.
- `:4335` trở đi bỏ qua kết quả False của `_set_cell`; lỗi sau Enter tại `:4353`, `:4375` chủ yếu được in log, vẫn tiến tới Post tại `:4392`.
- `:4462–4478` bắt số chứng từ và trả thành công trước khi kiểm MessageType E/A tại `:4480`.
- Mock GUI với mọi ô điền thất bại và thông báo E `Delivery 1000000001 does not exist`: nút Post vẫn được gọi và kết quả True kèm số đó. Đây là dữ liệu giả lập để tái hiện lỗi, không phải chứng từ SAP được tạo.

**Cần sửa:** fail closed khi không có session/field/layout/UoM; kiểm E/A và phủ định trước nhận số; xác minh document type, năm, vật tư, lô, số lượng, kho và đơn vị sau ghi. Không gọi mọi Material Document là OB.

### P1-03. Chuẩn hóa số còn đường làm sai 1.000 lần

- `app/number_utils.py:65`: AUTO `0,001` thành 1.
- `:88–94`: `parse_sap_float('1.234', fmt='AUTO')` thành 0; mất dấu hiệu lỗi/mơ hồ.
- `:101–103`: formatter lại đoán EU, `1.234` thành `1234`.
- Probe global EU: chuỗi SAP `0,001` → float 0.001 → chuỗi hiển thị `0.001` → allocator đọc theo EU thành 1.
- Đường chuyển này có trong adapter `sap_engine.py:1416–1419` và allocator `nvl_service.py:67`.

**Cần sửa:** giữ Decimal/numeric xuyên suốt, chỉ format ở lớp hiển thị; locale nguồn phải rõ; lỗi parse là UNKNOWN/ERROR, không thành tồn 0.

### P1-04. Tồn thiếu chứng cứ được tự gán kho 1005, PC và FT

- `app/nvl_service.py:50–54` lấy mặc định unit PC, plant/sloc từ tham số và category FT.
- `app/sap_engine.py:1420–1427` cũng tự bổ sung các trường đó; mapping đọc nguồn ở `:1388` không lấy plant/sloc.
- Handler gán `source_complete=True` tại `nvl_handler.py:77`.
- Probe dòng tồn thiếu plant/sloc/unit/cat vẫn phân bổ đủ 5, không error/exclusion; scope được dựng thành 3093/1005/PC.

Owner 3093 chưa phải bằng chứng SLoc 1005. Cần xác minh nguồn hoặc block khi không có mapping đã kiểm chứng, thay vì cho dữ liệu mặc định đi qua bộ kiểm.

### P1-05. Kiểm tra người gửi ghi ATT dùng dữ liệu trong body

- `app/handlers/att_handler.py:70–75` mặc định cho phép ghi, danh sách trusted rỗng không chặn; đối chiếu `meta.senderEmail` do nội dung thư cung cấp.
- `main.py:513` đã truyền `actualSenderEmail`, nhưng handler không dùng nó.
- Probe: người gửi thật ngoài danh sách, body khai địa chỉ trusted → runner vẫn được gọi.
- Standard/Portable hiện đóng gói config có `allow_sap_writes=true`, trusted rỗng. Patch giữ config cũ, nhưng mặc định trong `app/config.py` cũng bật ghi.

**Cần sửa:** dùng danh tính intake đã xác thực, cấu hình mặc định tắt ghi, allowlist bắt buộc cho mode ghi. Không dùng tên/email trong body như bằng chứng quyền thao tác.

### P1-06. Quyền sở hữu tab SAP chưa được bảo đảm

- `app/sap_attachment.py:905–907` chấp nhận expected identity rỗng; khi có identity chỉ kiểm system/client/user, bỏ session_id.
- Probe expected ses[5], thực tế ses[0] vẫn VERIFIED.
- `:2024–2027` fallback sang tab người dùng nếu không có tab trong dải bot. Probe 3 tab đều dành cho người dùng vẫn chọn ses[2].
- Worker ở `:2054` cố định connection 0, không theo connection đã chọn.
- `SapEngine._connect_sap` cũng coi tab theo chỉ số thuộc bot dù đang ở transaction khác; vị trí tab không chứng minh chủ sở hữu.

**Cần sửa:** lưu ownership/identity thực của phiên đã tạo hoặc đã được giao; không fallback sang tab người dùng; giữ đúng connection trên mọi worker; kiểm lại trước khi mở giao dịch/ghi.

### P1-07. GOS/HeadText còn đường xác nhận thành công giả

- `sap_attachment.py:1348` bỏ qua kết quả chờ popup đóng; `:1363–1366` nhận các từ CREATED/SAVED/ATTACHMENT để trả DONE mà không đọc lại.
- Probe popup không đóng, status cũ `Document created and saved`, attachment list rỗng → `DONE_SAP_SBAR`, không readback.
- `:1437–1444` URL attachment có thể trả DONE khi thiếu control TITLE/URL và status Warning.
- `:1121–1124` kiểm Header bằng substring: `OUT_P1_1234` được nhận đúng cho yêu cầu `OUT_P1_123`.
- Nhánh lưu Header ở `:1172` không reload/readback Header.
- Runner `:1896–1907` đặt success=True cho PARTIAL thiếu VAT/BBBG; handler `:116–117` nâng cả batch thành SUCCEEDED khi tất cả dòng đều “partial-success”.
- Năm thiếu được lấy năm hiện tại tại `att_handler.py:102`, có thể sai với chứng từ năm trước.

**Cần sửa:** ack liên quan đúng thao tác + document/year + filename/URL/readback thực; HeadText so token/chính xác; status mỗi bước là dữ liệu có kiểu; PARTIAL phải truyền nguyên nghĩa lên Teams.

### P1-08. Tra cứu ASN hiện lỗi ngay ở hợp đồng gọi hàm

`app/handlers/logistics_handler.py:104` truyền `inbound_no=...`, nhưng `SapEngine.find_inbound_vl06i` tại `app/sap_engine.py:3291` chỉ nhận po_no/ref_no/materials/plant. Probe dùng bound method thật trả `TypeError: unexpected keyword argument 'inbound_no'` trước khi gọi SAP.

Sau sửa signature còn phải lọc/đọc ngược đúng IB. Adapter dùng refPo nhưng service đọc poNo/ebeln nên có thể mất PO; handler bỏ PARTIAL/UNKNOWN của service và luôn trả SUCCEEDED. Nguồn ZMME240/container/customs thực chưa được chứng minh bởi đường handler này.

## 3. Các lỗi hoặc phần còn thiếu khác

| Vấn đề | Bằng chứng và tác động |
|---|---|
| BBBG chưa có nguồn đầy đủ | `bbbg_handler.py:20–22` dùng MB51 theo ngày và VL06I chỉ chưa GR, rồi gán source_complete=True. MB51 adapter chưa có khóa đầy đủ inbound/year/item. Matcher tốt hơn nhưng đầu vào chưa đủ để đối soát nghiệp vụ. |
| BBBG rỗng vẫn SUCCEEDED | Lệnh thiếu material có thể được chấp nhận, hai adapter trả [], handler báo 0/0 khớp; không được coi đó là xác nhận không chênh lệch. |
| BOM thiếu dòng bị che | Table có 2 dòng, dòng thứ hai lỗi: `sap_engine.py:4575` break, giữ dòng đầu CÓ BOM và handler SUCCEEDED; cần PARTIAL/ERROR. |
| Badge BOM sai | `teams_dispatcher.py:435` kiểm substring CÓ BOM nên KHÔNG CÓ BOM cũng thành badge dương; CHƯA XÁC ĐỊNH BOM lại hiển thị Không có BOM. Probe đã tái hiện. |
| Queue có thể bị mất | `sap_engine.py:4685` bỏ dòng có queue_name nhưng document/error rỗng; chẩn đoán service chưa được nối vào handler. |
| CHECK_RES chưa tra cứu | `production_handler.py:34` trả hướng dẫn cung cấp số phiếu; chưa có lời gọi SAP ngay cả khi đã cung cấp số. |
| Alias không đồng bộ | JSON SCWM và /SCWM/POST được parser chấp nhận nhưng router không có handler. Cú pháp tự nhiên SAP SCWM... hiện chuẩn hóa VITRI nên vẫn route được; không đánh đồng hai đường này. |
| PING khẳng định SAP chưa kiểm | SystemHandler trả Kết nối bình thường dù ctx.sap=None; probe thành công. PING chỉ chứng minh đường ứng dụng đã xử lý request. |
| FIFO ngày không hợp lệ | 31.02.2026, 25:70:00 được ráp chuỗi và allocator chấp nhận; cần validate lịch/giờ thật. |
| Reservation sau SAP | `nvl_handler.py:100` reserve sau post; lỗi DB có thể khiến kết quả báo FAILED dù SAP đã tạo chứng từ. Chưa liên kết lifecycle reservation với chứng từ thật. |
| Cache nhận control sai | Cache chỉ theo mode+alias; probe cached MATNR trỏ GuiButton tên WRONG_BUTTON vẫn được nhận. Cần namespace system/client/tcode/screen và kiểm Name/Type. |
| Dedup chưa namespace đồng nhất | Request UUID phân biệt Store, nhưng processed_messages kiểm messageId toàn cục. Probe cùng MSG1 ở Store B bị coi đã xử lý bởi Store A. |
| STOP vẫn WMIC rộng | STOP_BRIDGE.cmd vẫn terminate Python theo main.py/SAP_BOT; bypass cooperative stop và có thể bị watchdog khởi động lại. Patch không mang STOP/SILENT cập nhật. |

## 4. Điểm tốt đã giữ được

- Source trong Standard trích xuất khớp workspace lúc snapshot; commit refactor đúng `8296211`.
- Có RequestContext/CommandResult, Router và 7 handler; main chủ yếu xử lý Outlook, state và delivery. Đây là nền tốt để kiểm thử tích hợp.
- JSON ASN không còn rơi sang ATT; bảng CHECK_ATT + DOC/năm giữ route tra cứu trong probe.
- OB điều phối đã trả PREVIEW; không còn gán giả DISPATCHED/C trong handler.
- BBBG không suy ra Đã GR chỉ vì quantity bằng nhau; thông tin hải quan thiếu thành UNKNOWN; export portable/UUID và không hardcode phê duyệt.
- Zero available không bị allocator đổi thành total; QC rõ ràng, nhu cầu âm, sai đơn vị được chặn trong các case kiểm tra.
- Batch-specific và generic OB dùng chung pool; cùng plan lưu hai lần không tăng gấp đôi reservation.
- DB tương đối, đóng connection; request claim và outbox/result/processed lưu trước HTTP là cải thiện rõ.
- Accepted HTTP tách khỏi delivery Teams; timeout/5xx có trạng thái AMBIGUOUS, tránh tự gửi trùng khi chưa rõ.
- Worker lease và cooperative stop API tồn tại; Outlook archive/recovery được bắt lỗi; thread-local active_document cải thiện chia sẻ trạng thái.

## 5. Đánh giá các tuyên bố trong nội dung bàn giao

| Tuyên bố | Đánh giá |
|---|---|
| Có 7 handler/router | Xác nhận có trong mã nguồn. |
| 304 test pass | Xác nhận chạy lại 304 pass, 69.91s trong môi trường cô lập. |
| Git commit/tag/release v1.0.74 | Commit local 8296211 khớp; trang GitHub Release công khai cũng trỏ commit này. Không kiểm chứng trạng thái task Paperclip/manager approval. |
| Blast Radius=0, độc lập 100% | Không có cơ sở kỹ thuật; dependencies và execution state còn dùng chung. |
| Bảo toàn 100% logic hoàn thiện | Không đạt: ASN sai signature, quyền ghi/session/readback và numeric vẫn có lỗi tái hiện. |
| Tăng tốc 4.6–8.2x đã đo thực tế | Chưa được chứng minh. `scripts/benchmark_v1072_speedup.py:10` trở đi dùng hằng số latency và hệ số hiệu suất giả định để tính tỷ lệ. Stress tests chủ yếu mock. Đây là ước tính mô hình, không phải benchmark SAP thực tế. |
| Portable chạy ngay | Có runtime, nhưng vẫn cần SAP GUI/GUI Scripting, Outlook Classic và cấu hình đúng. Lượt này chưa đo chạy live trên máy triển khai. |

GitHub: https://github.com/taikthp357-ctrl/SAP/releases/tag/v1.0.74

## 6. Đóng gói

| Gói | Bytes | File |
|---|---:|---:|
| Portable | 14.103.905 | 1.422 |
| Standard | 453.389 | 68 |
| Patch | 260.482 | 50 |

- Patch có đủ Python handler/core mới, không có config/DB; chỉ có START_BRIDGE trong nhóm launcher.
- Standard/Portable có config đã cấu hình và sap_ui_cache.json từ máy build. Không in webhook/token vào báo cáo; nên phát hành config mẫu sạch và loại cache máy build.
- Builder vẫn quét toàn cây, blacklist chưa loại đầy đủ WAL/SHM/backup, xóa mọi ZIP cũ, copy ra đường dẫn cá nhân. Cần allowlist + manifest/checksum, giữ bản cũ để rollback và bảo toàn config/master/data.
- Không chạy package_release.py trong kiểm tra này vì script có tác động xóa/copy.

## 7. Thứ tự xử lý đề xuất

1. Khóa các đường ghi ngoài policy: kế hoạch PREVIEW, dùng sender thực, quyền ghi mặc định tắt; guard dùng chung ở router/executor.
2. Sửa MIGO: thiếu session/field/đơn vị/layout/lỗi SAP phải dừng trước Post; phân biệt UNKNOWN sau tác động, không suy thành FAILED an toàn để chạy lại.
3. Giữ numeric/Decimal xuyên suốt; xác minh plant/sloc/UoM/category từ nguồn; validate ngày thật.
4. Sửa ownership tab/connection, HeadText exact match, GOS evidence và PARTIAL propagation.
5. Sửa ASN signature/adapter schema; chuẩn hóa source_complete/PARTIAL/UNKNOWN cho BOM/Queue/BBBG.
6. Thêm test parser→router→handler→adapter với autospec, không chỉ MagicMock tự nhận mọi keyword; thêm fault injection ở từng bước trước/sau SAP và DB.
7. Sửa STOP/update/package; sau đó kiểm golden cases trên SAP test theo chứng từ đã xác minh, đối chiếu trước/sau và đo thời gian thực tế. Chỉ nghiệm thu live khi có evidence đó.

## 8. Bằng chứng và giới hạn

Kết quả lưu tại `docs/audits/v1.0.74-20261026/`: audit-tests.log, audit-root-probes.log, audit-att-results.json, audit-nvl-results.log, package-manifest.json.

Snapshot suite: `C:/Users/taikt/AppData/Local/Temp/sap-v1074-audit-s7yuopwo`. COM GetObject/Dispatch và urlopen bị chặn; SQLite chỉ cho thư mục tạm. Các probe executor dùng GUI giả/mocks, số chứng từ trong kết quả chỉ là dữ liệu test. Không có chứng từ SAP hoặc attachment công ty nào được tạo để kiểm này.

Đây là kết luận từ source và kiểm thử cô lập, không phải xác nhận QA/manager hoặc nghiệm thu live. Không sửa ứng dụng hay ghi đè các ZIP vận hành trong lượt phân tích.
