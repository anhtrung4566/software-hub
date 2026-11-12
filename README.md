# Kho Phần Mềm Chuyên Dụng · VBTM Pro & VietNamOCR 3.2 Pro

Cổng thông tin và ứng dụng Web/Desktop với giao diện **iOS 18 Liquid Glassmorphism** phục vụ công tác văn thư, lưu trữ của **Đảng ủy và Ủy ban nhân dân phường**, bao gồm 2 phần mềm cốt lõi:

1. **VBTM Pro 2.2.4** (`/products/vbtm-pro.html`):
   - Chuẩn hóa thể thức văn bản hành chính của **Ủy ban nhân dân phường** theo **Nghị định 30/2020/NĐ-CP** và văn bản của **Đảng ủy phường** theo **Hướng dẫn số 05-HD/VPTW**.
   - Tạo và chèn **Mã QR động** xác thực văn bản trực tiếp.
   - **Soát lỗi chính tả & thẩm định thể thức AI ngoại tuyến 100%**.
   - Xem trước trang A4 thời gian thực, chuẩn hóa trực tiếp cửa sổ Word đang mở và **Xuất ra file Microsoft Word (`.doc`)**.

2. **VietNamOCR 3.2 Pro** (`/products/vietnam-ocr.html`):
   - Nhận diện chữ Tiếng Việt từ hình ảnh (`PNG`, `JPG`, `WEBP`, `BMP`, `TIF`) và tài liệu `PDF` scan phục vụ lưu trữ hồ sơ Đảng ủy & UBND phường.
   - Xoay trang, tăng tương phản làm rõ chữ scan mờ, tự động sửa lỗi dấu câu và khoảng trắng tiếng Việt.
   - **Bóc tách & Nối bảng biểu đa trang (Multi-page Table Merge)**.
   - Đồng bộ chỉnh sửa trực tiếp và **Xuất ra Microsoft Word (`.doc`) & Microsoft Excel (`.xls`)**.

## Cấu trúc thư mục

- `index.html` — Trang chủ Software Hub
- `products/vbtm-pro.html` — Trang chi tiết & Phần mềm VBTM Pro Studio
- `products/vietnam-ocr.html` — Trang chi tiết & Phần mềm VietNamOCR 3.2 Studio
- `css/style.css` — Hệ thống giao diện kính trong suốt iOS 18 (Dark/Light Mode)
- `js/main.js`, `js/vbtm-app.js`, `js/ocr-app.js` — Mã nguồn xử lý giao diện và 2 phần mềm
- `downloads/` — Bộ cài đặt Windows Native (`.exe`) và gói nén (`.zip`)
- `desktop-src/` — Mã nguồn C# Windows Forms & Word VBA Macro (`.bas`)
