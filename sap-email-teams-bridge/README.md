# SAP - Outlook - Teams Bridge (0 Open Port / 0 PowerShell)

Công cụ cầu nối tự động giữa **Outlook** $\rightarrow$ **SAP GUI Scripting** $\rightarrow$ **Microsoft Teams (Power Automate)**.

## Điểm nổi bật
1. **Tuyệt đối an toàn (Enterprise Grade):**
   - **0 Open Port:** Không mở cổng mạng lắng nghe (không socket server), chỉ gửi Outbound HTTPS POST duy nhất tới Power Automate.
   - **0 PowerShell:** Điều khiển SAP GUI và Outlook hoàn toàn qua Windows COM Object trong RAM, không gọi bất kỳ lệnh PowerShell hay tạo tiến trình nhạy cảm nào.
2. **Cơ chế Keep-Alive & Auto-Reconnect:**
   - Tự động làm mới phiên SAP mỗi 15 phút không thao tác, chống bị văng timeout do cài đặt của SAP NetWeaver.
   - Tự động kết nối lại nếu phiên bị rớt mạng hoặc người dùng vô tình đóng SAP.
3. **Tra cứu vị trí chính xác:**
   - Tích hợp chuẩn giao diện **Posting Change (`/SCWM/POST`)** để trích xuất đầy đủ: Mã hàng, Số Batch, Vị trí Bin, Ngày và Giờ GR.
   - Hỗ trợ cả **MB52** (Tồn kho tự do) và **MB51** (Lịch sử hạch toán).

## Cấu trúc thư mục
```text
sap-email-teams-bridge/
├── config.json                     # File cấu hình URL Webhook & tham số
├── main.py                         # Điểm chạy chính của Worker
├── START_BRIDGE.cmd                # File khởi động 1-Click trên Windows
├── package_release.py              # Script đóng gói file ZIP phát hành
├── app/
│   ├── config.py                   # Quản lý đọc cấu hình
│   ├── email_parser.py             # Bộ bóc tách lệnh email
│   ├── sap_engine.py               # Điều khiển SAP GUI COM & Keep-Alive
│   └── teams_dispatcher.py         # Gửi kết quả Webhook kèm bảng HTML
├── templates/
│   ├── power_automate_schema.json  # Schema mẫu dán vào Power Automate
│   └── email_samples.txt           # Mẫu nội dung email để test
├── docs/
│   └── HUONG_DAN_SU_DUNG.md        # Hướng dẫn cài đặt & vận hành chi tiết
└── tests/
    ├── test_email_parser.py        # Unit tests bộ đọc email
    └── test_teams_dispatcher.py    # Unit tests gửi bảng HTML
```

## Chạy thử nghiệm
1. Đảm bảo đã mở SAP GUI và Outlook Classic (có thư mục `Inbox/SAP_BOT`).
2. Chạy file `START_BRIDGE.cmd`.
3. Gửi email mẫu:
   ```text
   LENH: VITRI | MA: <Mã_hàng> | BATCH: <Lô>
   ```
