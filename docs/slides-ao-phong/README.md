# Slide Deck Kế Hoạch Khởi Nghiệp Áo Phông Trendy (PTIT Standard)

Bộ slide trình chiếu tương tác chuẩn phong cách Quadro / McKinsey được thiết kế sẵn bằng HTML5 + Tailwind CSS, sẵn sàng chạy trực tiếp trên trình duyệt hoặc xuất bản qua **GitHub Pages**.

## 1. Xem và Chạy Slide Trực Tiếp Trên Máy (Offline / Local)
- Mở file `index.html` trong thư mục `presentation/` bằng bất kỳ trình duyệt nào (Chrome, Edge, Brave, Cốc Cốc, Safari).
- Điều khiển:
  - Bấm phím **Mũi tên phải (`→`)** hoặc phím **Cách (Space)** để sang slide tiếp theo.
  - Bấm phím **Mũi tên trái (`←`)** để quay lại slide trước.

## 2. Xuất Bản Lên GitHub & Xem Trực Tuyến Qua GitHub Pages

Để có link web trình chiếu trực tiếp trên mạng (`https://<username>.github.io/<repo-name>/`), bạn chỉ cần làm 3 bước:

```bash
# 1. Khởi tạo và commit code
git init
git add presentation/
git commit -m "feat: add trendy t-shirt startup presentation slide deck"

# 2. Tạo repository trên GitHub và đẩy code lên
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO_NAME>.git
git branch -M main
git push -u origin main
```

### Bật GitHub Pages:
1. Vào **Settings** của repository trên GitHub.
2. Chọn menu **Pages** ở cột trái.
3. Tại phần **Build and deployment > Branch**, chọn nhánh `main` và thư mục `/presentation` (hoặc `/root`).
4. Nhấn **Save**. Sau 1–2 phút, bạn sẽ nhận được link trình chiếu trực tiếp dạng:
   `https://<YOUR_USERNAME>.github.io/<YOUR_REPO_NAME>/presentation/`

## 3. Cấu Trúc Các Slide
- **Slide 01**: Bối cảnh thị trường & Nhu cầu Gen Z
- **Slide 02**: Định vị phân khúc & Khách hàng mục tiêu (STP)
- **Slide 03**: Ma trận SWOT & Phương án xử lý phôi an toàn
- **Slide 04**: Mục tiêu thử nghiệm 3 tháng đầu (Doanh số - Thương hiệu - Vận hành)
- **Slide 05**: Chiến lược Marketing Mix 4P & Cấu trúc kinh tế (COGS > 60%)
- **Slide 06**: Hạ tầng kênh bán & Giải pháp công nghệ Omnichannel (Sapo API)
- **Slide 07**: Xúc tiến bán hàng (TikTok, Shopee, Affiliate KOC)
- **Slide 08**: Vận hành & Quy trình học nhanh từ dữ liệu đơn hàng
- **Slide 09**: Lộ trình 90 ngày (Dựng nền - Mở bán - Tăng tốc)
- **Slide 10**: Dự toán ngân sách đầu tư ban đầu (CAPEX, OPEX & Vốn lưu động)
- **Slide 11**: Hệ thống đo lường & Chỉ số kiểm soát sống còn (KPIs)
- **Slide 12**: Phân tích tài chính, sản lượng hòa vốn & tỷ suất sinh lời
- **Slide 13**: Kết luận mô hình tinh gọn & Cam kết thực thi
