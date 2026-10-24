# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell)

Công cụ cầu nối tự động giữa **Outlook** $\rightarrow$ **SAP GUI Scripting** $\rightarrow$ **Microsoft Teams (Power Automate)**.

## Điểm nổi bật
1. **Thiết kế bảo mật & an toàn (Zero Inbound Port Architecture):**
   - **0 Open Port:** Không mở cổng mạng lắng nghe (không socket server), chỉ gửi Outbound HTTPS POST duy nhất tới Power Automate Webhook.
   - **0 PowerShell:** Điều khiển SAP GUI và Outlook hoàn toàn qua Windows COM Object trong RAM, không gọi bất kỳ tiến trình PowerShell hay subprocess nhạy cảm nào.
   - **Kiểm soát phiên SAP:** Tự động dò tìm session rỗi tại Main Menu / bot T-Code, tránh can thiệp nhầm vào phiên làm việc thực tế của người dùng.
2. **Cơ chế Keep-Alive & Auto-Reconnect:**
   - Tự động làm mới phiên SAP mỗi 15 phút không thao tác, chống bị văng timeout do cài đặt của SAP NetWeaver.
   - Tự động kết nối lại nếu phiên bị rớt mạng hoặc người dùng vô tình đóng SAP.
3. **Hỗ trợ 8 nghiệp vụ logistics & tra cứu thông minh:**
   - **/SCWM/POST:** Vị trí lưu kho (Bin), Số lượng khả dụng, Số Lô, Ngày GR, Giờ GR, Trạng thái kho EWM.
   - **MB52 (Tồn kho):** Tồn kho tổng hợp hoặc lọc riêng Kho 1005 & 3093 (Tồn, Lô, ĐVT).
   - **MB52 (Tồn thấp đa Plant):** Cảnh báo Top N mã tồn thấp theo Dòng xe (BODY: Plant 3093 & 3001; XMĐ: Plant 3093, 2001 & 3001).
   - **MB51:** Lịch sử xuất - nhập - chuyển kho trong N ngày gần nhất (Posting Date, Movement Type).
   - **VL06I:** Danh sách đơn Inbound (IB) treo chưa Goods Receipt (GM = A).
   - **ME2N:** Danh sách đơn điều chuyển kho (STO) theo mã hàng (SL đặt, đã xuất, còn nợ).
   - **VL06O:** Danh sách đơn Outbound (OB) đang mở chưa xuất hàng.
   - **/SCWM/MON:** Lịch sử xuất nhập & di chuyển Bin-to-Bin (Warehouse Tasks: Confirmed & Cancelled).

## Cấu trúc thư mục
```text
sap-email-teams-bridge/
├── config.json                     # File cấu hình URL Webhook & tham số
├── main.py                         # Điểm chạy chính của Worker
├── START_BRIDGE.cmd                # File khởi động 1-Click trên Windows
├── package_release.py              # Script đóng gói file ZIP phát hành
├── app/
│   ├── config.py                   # Quản lý đọc cấu hình & auto-fix JSON
│   ├── email_parser.py             # Bộ bóc tách lệnh email, JSON & coaching
│   ├── material_resolver.py        # Quy đổi mã nội bộ / mã khách hàng qua Excel
│   ├── number_utils.py             # Chuẩn hóa số lượng SAP đa định dạng (US/EU/AUTO)
│   ├── sap_engine.py               # Điều khiển SAP GUI COM, bảo vệ session & Keep-Alive
│   ├── state_manager.py            # Quản lý hàng đợi Outbox & chống trùng bằng SQLite
│   └── teams_dispatcher.py         # Gửi kết quả Webhook kèm bảng HTML & KPI
├── data/
│   └── bridge_state.db             # Cơ sở dữ liệu SQLite lưu Outbox & Processed IDs
├── templates/
│   ├── power_automate_schema.json  # Schema mẫu dán vào Power Automate
│   └── email_samples.txt           # Mẫu nội dung email để test
├── docs/
│   └── HUONG_DAN_SU_DUNG.md        # Hướng dẫn cài đặt & vận hành chi tiết
└── tests/                          # Bộ unit tests hồi quy tự động (54 tests pass 100%)
```

## Chạy thử nghiệm
1. Đảm bảo đã mở SAP GUI và Outlook Classic (có thư mục `Inbox/SAP_BOT`).
2. Chạy file `START_BRIDGE.cmd`.
3. Gửi email mẫu:
   ```text
   LENH: VITRI | MA: <Mã_hàng> | BATCH: <Lô>
   ```
