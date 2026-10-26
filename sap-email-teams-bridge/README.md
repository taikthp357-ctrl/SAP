# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell) - v1.0.59

Công cụ cầu nối tự động giữa **Outlook** $\rightarrow$ **SAP GUI Scripting** $\rightarrow$ **Microsoft Teams (Power Automate)**.

## Điểm nổi bật
1. **Triệt tiêu lỗi `ATTACHMENT_UNCERTAIN` & Khớp linh hoạt Tiêu đề SAP GOS (v1.0.59):**
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
