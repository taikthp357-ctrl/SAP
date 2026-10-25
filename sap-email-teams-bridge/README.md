# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell) - v1.0.50

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

3. **Ổn định tương tác Live SAP GUI & Ma trận Ý định (v1.0.50 Live Stabilization):**
   - **VL06O (Check Outbound):** Điền thẳng mã vật tư vào `ctxtIT_MATNR-LOW` ngay tại Selection Screen trước khi F8; không ép cứng Shipping Point (`IF_VSTEL-LOW = '3093'`); xóa sạch giới hạn ngày giao hàng; bổ sung cơ chế fallback bấm nút Filter `wnd[0]/tbar[1]/btn[29]` trên Classic ALV Fullscreen Container.
   - **ME2N (Check STO/PO):** Cấu hình `LISTU = "ALV"`, tham số chọn `SELPA = "WE101"` (Open Goods Receipt); phân định chuẩn xác Nhà máy nhận hàng `S_WERKS-LOW` (Plant 3093) thay vì nhầm sang Supplying Plant.
   - **VL06I (Check Inbound):** Bổ sung phím tắt Shift+F2 (`sendVKey(14)`) tại ô `IT_LFDAT-LOW` để xóa sạch ngày mặc định; điền mã vật tư và chuyển sang Item View (`btn[18]`) trước khi trích xuất.
   - **Chuẩn hóa nhãn SLoc MB52:** Đổi nhãn tóm tắt từ `📦 Kho {sloc}` thành `📦 SLoc {sloc}` để phản ánh đúng thuật ngữ Storage Location trong SAP ERP.
   - **Bảng Ma trận Ý định theo Trọng số (Weighted Intent Matrix - 12 mức ưu tiên):** Tự động phân loại ngôn ngữ tự nhiên (viết tắt, tiếng lóng kho vận) với độ chính xác tuyệt đối:
     1. `WHO_IS_TAI`: Easter Egg card tôn vinh tác giả Phạm Tiến Tài (FVN-KSX-NMVFHP).
     2. `ATT_MIGO`: Đính kèm file GOS & điền số phiếu HeadText MIGO.
     3. `CHECK_BOM`: Định mức linh kiện con (CS03).
     4. `CHECK_QUEUE`: Lỗi kẹt Queue EWM (/SCWM/MON).
     5. `CHECK_OB`: Danh sách đơn xuất Outbound (VL06O).
     6. `CHECK_IB`: Danh sách đơn nhập Inbound (VL06I/ZMME240).
     7. `STO`: Đơn điều chuyển kho & PO mở (ME2N).
     8. `LICHSU`: Lịch sử xuất nhập 7–N ngày (MB51).
     9. `CHECK_RES`: Phiếu lĩnh / Reservation (MB23/MB25).
     10. `TON`: Tồn kho theo SLoc (MB52).
     11. `VITRI`: Vị trí Bin, Batch, Ngày/Giờ GR (/SCWM/POST).
     12. Mặc định thông minh: Tự động chạy `VITRI` khi chỉ gửi mã hàng.
   - **Kiểm tra kép Combo (VITRI $\rightarrow$ MB52 Fallback):** Khi tra cứu vị trí `/SCWM/POST` trả về 0 dòng, hệ thống tự động kiểm tra MB52 để phát hiện ngay hàng đang nằm ở các SLoc ngoài EWM (như 3000, 9999).

4. **An toàn nghiệp vụ & Kiểm soát ranh giới Đọc / Ghi (P1 Safety Hardening):**
   - **F01 (Phân tách truy vấn JSON & Đính kèm MIGO):** JSON tra cứu được định tuyến độc lập, không rơi vào nhánh ghi `ATT_MIGO`.
   - **F02 (Kiểm soát phiên đính kèm MIGO):** Chế độ dry-run và kiểm tra session an toàn; không tự tạo báo cáo giả mạo trạng thái `DONE`/`HEADTEXT_DONE`.
   - **F03 (Minh bạch mô phỏng Điều phối OB & MIGO 311):** Gắn nhãn rõ ràng `SIMULATED`/`PREVIEW`, không sinh số chứng từ giả mạo.
   - **F04 (Đối soát BBBG nghiêm ngặt):** Khóa đối soát chặt chẽ `(material, batch, doc_no, unit)`. Khấu trừ chuẩn xác chứng từ hủy theo dấu số lượng.
   - **F05 (Xuất xứ dữ liệu ASN rõ ràng):** Dữ liệu thiếu tờ khai thể hiện minh bạch `UNKNOWN`.
   - **F06 (Phân bổ NVL & OB chính xác):** Loại trừ tồn kho QC/Blocked; phân biệt rõ tồn bằng 0 và không có dữ liệu.
   - **F07 (Idempotency đặt trước & Tính linh hoạt đường dẫn CSDL):** Hỗ trợ đường dẫn tương đối, chống nhân bản đặt trước trùng lặp.
   - **F08 (Chuẩn hóa số lượng đa định dạng US/EU):** Phân tích chính xác số thập phân và phân cách hàng nghìn; badge KPI hiển thị nhất quán.

## Cấu trúc thư mục
```text
sap-email-teams-bridge/
├── config.json                     # File cấu hình URL Webhook & tham số vận hành
├── main.py                         # Điểm chạy chính của Worker
├── START_BRIDGE.cmd                # File khởi động 1-Click trên Windows
├── package_release.py              # Script đóng gói file ZIP phát hành (Patch, Standard, Portable)
├── app/
│   ├── bbbg_service.py             # Nghiệp vụ đối soát Biên bản bàn giao & MIGO
│   ├── config.py                   # Quản lý đọc cấu hình & auto-fix JSON
│   ├── email_parser.py             # Bộ bóc tách lệnh email, Weighted Intent Matrix V2.0
│   ├── extended_query_service.py   # Dịch vụ tra cứu mở rộng (ASN, STO, EWM Mon...)
│   ├── material_resolver.py        # Quy đổi mã nội bộ / mã khách hàng qua Excel
│   ├── number_utils.py             # Chuẩn hóa số lượng SAP đa định dạng (US/EU/AUTO)
│   ├── nvl_service.py              # Nghiệp vụ cấp phát NVL & mô phỏng MIGO 311
│   ├── ob_service.py               # Nghiệp vụ điều phối Outbound & phân bổ tồn kho
│   ├── runtime_control.py          # Kiểm soát an toàn phiên & ranh giới chạy
│   ├── sap_attachment.py           # Quản lý tải và đính kèm GOS / MIGO header text
│   ├── sap_engine.py               # Điều khiển SAP GUI COM, Live SAP GUI Stabilization
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
└── tests/                          # Bộ unit tests hồi quy tự động (224 tests pass 100%)
```

## Chạy thử nghiệm
1. Đảm bảo đã mở SAP GUI và Outlook Classic (có thư mục `Inbox/SAP_BOT`).
2. Chạy file `START_BRIDGE.cmd`.
3. Gửi email mẫu:
   ```text
   LENH: VITRI | MA: <Mã_hàng> | BATCH: <Lô>
   ```
