# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell) - v1.0.49

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

3. **An toàn nghiệp vụ & Kiểm soát ranh giới Đọc / Ghi (P1 Safety Hardening - v1.0.49):**
   - **F01 (Phân tách truy vấn JSON & Đính kèm MIGO):** JSON tra cứu (e.g. `TRA_CUU_ASN`, `MB52`, `VITRI`) được định tuyến độc lập, tuyệt đối không bị nhận nhầm sang lệnh ghi đính kèm `ATT_MIGO`.
   - **F02 (Kiểm soát phiên đính kèm MIGO):** Chế độ dry-run và kiểm tra session an toàn; không tự tạo báo cáo giả mạo trạng thái `DONE`/`HEADTEXT_DONE` khi không có phiên thực thi.
   - **F03 (Minh bạch mô phỏng Điều phối OB & MIGO 311):** Các lệnh xem trước/mô phỏng gắn nhãn rõ ràng `SIMULATED`/`PREVIEW`, tuyệt đối không sinh số chứng từ SAP giả mạo (`mat_doc` dạng 49...).
   - **F04 (Đối soát BBBG nghiêm ngặt):** Khóa đối soát chặt chẽ theo bộ 4 thuộc tính `(material, batch, doc_no, unit)`. Khấu trừ chuẩn xác chứng từ hủy (102/312) theo dấu số lượng.
   - **F05 (Xuất xứ dữ liệu ASN rõ ràng):** Dữ liệu thiếu tờ khai hải quan thể hiện minh bạch `UNKNOWN`, không mặc định giả định "ĐÃ THÔNG QUAN".
   - **F06 (Phân bổ NVL & OB chính xác):** Loại trừ tồn kho QC/Blocked; tách biệt giữa tồn bằng 0 và không có dữ liệu; thống nhất nguồn tồn kho giữa nhu cầu chỉ định lô và nhu cầu chung.
   - **F07 (Idempotency đặt trước & Tính linh hoạt đường dẫn CSDL):** Hỗ trợ đường dẫn CSDL tương đối/cấu hình động; tránh nhân bản đặt trước trùng lặp.
   - **F08 (Chuẩn hóa số lượng đa định dạng US/EU):** Phân tích chính xác số thập phân và phân cách hàng nghìn; badge KPI hiển thị nhất quán, trực quan.

4. **Hỗ trợ toàn diện 8 nhóm nghiệp vụ logistics & tra cứu thông minh:**
   - **/SCWM/POST:** Vị trí lưu kho (Bin), Số lượng khả dụng, Số Lô, Ngày GR, Giờ GR, Trạng thái kho EWM.
   - **MB52 (Tồn kho):** Tồn kho tổng hợp hoặc lọc riêng Kho 1005 & 3093 (Tồn, Lô, ĐVT).
   - **MB52 (Tồn thấp đa Plant):** Cảnh báo Top N mã tồn thấp theo Dòng xe (BODY: Plant 3093 & 3001; XMĐ: Plant 3093, 2001 & 3001).
   - **MB51:** Lịch sử xuất - nhập - chuyển kho trong N ngày gần nhất (Tự động xóa trắng Material Document & Movement Type cũ).
   - **VL06I & Tra cứu ASN:** Danh sách đơn Inbound (IB) treo chưa Goods Receipt (GM = A), tình trạng thông quan và chi tiết chứng từ.
   - **ME2N:** Danh sách đơn điều chuyển kho (STO) theo mã hàng (SL đặt, đã xuất, còn nợ).
   - **VL06O & Phân bổ OB:** Danh sách đơn Outbound (OB) đang mở chưa xuất hàng và thuật toán phân bổ tồn kho khả dụng.
   - **/SCWM/MON:** Lịch sử xuất nhập & di chuyển Bin-to-Bin (Warehouse Tasks: Confirmed & Cancelled).
   - **Đối soát BBBG & GR Document:** So khớp chéo biên bản bàn giao và chứng từ nhập kho vật tư/thành phẩm.

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
│   ├── email_parser.py             # Bộ bóc tách lệnh email, JSON & coaching
│   ├── extended_query_service.py   # Dịch vụ tra cứu mở rộng (ASN, STO, EWM Mon...)
│   ├── material_resolver.py        # Quy đổi mã nội bộ / mã khách hàng qua Excel
│   ├── number_utils.py             # Chuẩn hóa số lượng SAP đa định dạng (US/EU/AUTO)
│   ├── nvl_service.py              # Nghiệp vụ cấp phát NVL & mô phỏng MIGO 311
│   ├── ob_service.py               # Nghiệp vụ điều phối Outbound & phân bổ tồn kho
│   ├── runtime_control.py          # Kiểm soát an toàn phiên & ranh giới chạy
│   ├── sap_attachment.py           # Quản lý tải và đính kèm GOS / MIGO header text
│   ├── sap_engine.py               # Điều khiển SAP GUI COM, bảo vệ session & Keep-Alive
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
└── tests/                          # Bộ unit tests hồi quy tự động (221 tests pass 100%)
```

## Chạy thử nghiệm
1. Đảm bảo đã mở SAP GUI và Outlook Classic (có thư mục `Inbox/SAP_BOT`).
2. Chạy file `START_BRIDGE.cmd`.
3. Gửi email mẫu:
   ```text
   LENH: VITRI | MA: <Mã_hàng> | BATCH: <Lô>
   ```
