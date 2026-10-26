# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell) - v1.0.52

Công cụ cầu nối tự động giữa **Outlook** $\rightarrow$ **SAP GUI Scripting** $\rightarrow$ **Microsoft Teams (Power Automate)**.

## Điểm nổi bật
1. **Thiết kế bảo mật & an toàn (Zero Inbound Port Architecture):**
   - **0 Open Port:** Không mở cổng mạng lắng nghe (không socket server), chỉ gửi Outbound HTTPS POST duy nhất tới Power Automate Webhook.
   - **0 PowerShell:** Điều khiển SAP GUI và Outlook hoàn toàn qua Windows COM Object trong RAM, không gọi bất kỳ tiến trình PowerShell hay subprocess nhạy cảm nào.
   - **Bảo vệ phiên người dùng (Strict Session Ownership Guard):** Khi bật chế độ an toàn, hệ thống chỉ sử dụng phiên rảnh (Idle), tự động xin mở phiên mới nếu cho phép, tuyệt đối không can thiệp hoặc cướp màn hình làm việc của người dùng đang thao tác.
   - **Quét sạch bộ lọc cũ (Clean Selection Screen):** Tự động xóa sạch toàn bộ các ID/ô lọc còn sót lại từ phiên trước (MBLNR, BWART, SLoc, Batch...) trước khi chạy MB51, MB52, VL06O, tránh bị lọc sai khiến kết quả trả về rỗng.

2. **Cơ chế Keep-Alive & Auto-Reconnect:**
   - Tự động làm mới phiên SAP mỗi 15 phút không thao tác, chống bị văng timeout do cài đặt của SAP NetWeaver.
   - Tự động kết nối lại nếu phiên bị rớt mạng hoặc người dùng vô tình đóng SAP.

3. **Chuyển chế độ MIGO Tự động & Khôi phục Multi-Tab Engine (v1.0.52):**
   - **Tự động chuyển từ DISPLAY sang CHANGE:** Khi thực hiện điền số phiếu HeadText hoặc đính kèm file, nếu phiên MIGO đang ở chế độ xem `A04 DISPLAY`, hệ thống tự động phát hiện ComboBox `GODYNPRO-ACTION != "A12"` và chuyển sang `A12 CHANGE`, đồng thời đảm bảo `GODYNPRO-REFDOC == "R02"`, giải quyết dứt điểm tình trạng bị khóa trường không cho lưu.
   - **Khôi phục cơ chế Chạy Đa Tab (Multi-Tab Engine):** Mở rộng khả năng xử lý song song với cấu hình `enable_multi_tab=True` và `max_batch_tabs=2`. Khi danh sách chứng từ lớn (> 5), hệ thống tự động mở đủ 3 Tabs: Tab 0 dành riêng cho Fast Lane (tra cứu tức thời), Tab 1 và Tab 2 dùng cho Batch Pool xử lý song song. Xác thực phiên được cải tiến dựa trên System, Client và User để các Tab con không bị chặn nhầm.

4. **Kích hoạt Chạy Thật ATT_MIGO & Nghiệp vụ CHECK_ATT (v1.0.51):**
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
└── tests/                          # Bộ unit tests hồi quy tự động (230 tests pass 100%)
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
