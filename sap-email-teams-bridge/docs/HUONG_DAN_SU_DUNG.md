# Hướng dẫn Vận hành & Cấu hình SAP Email Teams Bridge

Hệ thống kết nối tự động **Outlook $\rightarrow$ SAP GUI $\rightarrow$ Power Automate Teams Webhook** đạt tiêu chuẩn an toàn:
- **0 Open Port:** Không mở cổng mạng lắng nghe (không lo bị tường lửa chặn hoặc vi phạm bảo mật).
- **0 PowerShell / 0 Subprocess:** Không kích hoạt tiến trình hệ thống nhạy cảm, chạy thuần túy Windows COM và thư viện chuẩn Python.
- **Tự động quy đổi mã thông minh (Smart Material Resolver):** Tự động đối chiếu mã 8 số rút gọn (vd: `32013765`) thành mã SAP đầy đủ (vd: `BIW32013765AA`) qua sheet `DATA` tổng hợp.
- **Tự động giữ phiên (Keep-Alive):** Tự động phát xung nhịp mỗi 15 phút để chống bị SAP NetWeaver logout do timeout nhàn rỗi.

---

## 1. Sheet DATA Tổng hợp Mã (`data/DATA_MASTER.xlsx`)

Hệ thống đi kèm sẵn tệp `data/DATA_MASTER.xlsx` chứa hơn 1.100 mã vật tư thực tế tại kho và xưởng:
- **Cột B (Mã đầy đủ):** `BIW32013765AA`, `BIW32024300AB`...
- **Cột C (Mã rút gọn):** `32013765`, `32024300`...
- **Cột D (Model / Xe):** `Limo`, `VF5`, `VF8`...

> [!TIP]
> Bạn có thể mở file `data/DATA_MASTER.xlsx` thêm bớt mã vật tư mới bất kỳ lúc nào. Python sẽ tự động nhận diện và cập nhật danh sách vào bộ nhớ RAM ngay khi phát hiện file được lưu lại.

---

## 2. Cách thức gửi yêu cầu qua Email

Hệ thống hỗ trợ cực kỳ linh hoạt cho người dùng kho/xưởng:

### Cách 1: Chỉ nhắn mã số (Siêu nhanh - Tiện dụng nhất)
Người dùng chỉ cần gửi email với nội dung là mã hàng (ở Tiêu đề hoặc Nội dung thư):
- `32013765` $\rightarrow$ Tự động hiểu là tra cứu Vị trí & Ngày giờ GR của mã `BIW32013765AA`.
- `32013765 B2610` $\rightarrow$ Tự động lọc vị trí theo đúng số Batch `B2610`.
- `32013765, 32024300` $\rightarrow$ Tự động tra cứu vị trí của nhiều mã cùng lúc.

### Cách 2: Cú pháp đầy đủ
| Mục đích tra cứu | Cú pháp nội dung email |
|---|---|
| **1. Vị trí Bin, Batch, Ngày & Giờ GR (/SCWM/POST)** | `LENH: VITRI \| MA: 32013765 \| BATCH: B2610` |
| **2. Tồn kho khả dụng (MB52)** | `LENH: TON \| MA: 32013765, 32024300 \| PLANT: 3093` |
| **3. Lịch sử giao dịch (MB51)** | `LENH: LICHSU \| MA: 32013765 \| NGAY: 7 \| PLANT: 3093` |
| **4. Đơn Inbound treo chưa GR (VL06I)** | `LENH: IB_TREO \| MA: 32013765 \| PLANT: 3093` hoặc `check ib treo 32013765` |
| **5. Đơn điều chuyển kho STO (ME2N)** | `LENH: STO \| MA: 32013765 \| PLANT: 3093` hoặc `check sto 32013765` |

---

## 3. Quy trình khởi động

1. Mở sẵn **Outlook Classic** (đã có thư mục `Inbox/SAP_BOT`).
2. Mở sẵn **SAP Logon** (đã bật Enable Scripting).
3. Cập nhật link Webhook trong `config.json`.
4. Nhấp đúp chuột vào `START_BRIDGE.cmd` để chạy.
