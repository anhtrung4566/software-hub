# Kho Phần Mềm Chuyên Dụng · VBTM Pro & VietNamOCR 3.2 Pro

Cổng thông tin và ứng dụng Web/Desktop với giao diện **iOS 18 Liquid Glassmorphism**, bao gồm 2 phần mềm cốt lõi:

1. **VBTM Pro 2.2.4** (`/products/vbtm-pro.html`):
   - Chuẩn hóa thể thức văn bản hành chính theo **Nghị định 30/2020/NĐ-CP**, **Hướng dẫn 05-HD/VPTW** và **Quy chuẩn CAND (Thông tư 30/2023/TT-BCA)**.
   - Tạo và chèn **Mã QR động** xác thực văn bản trực tiếp.
   - **Soát lỗi chính tả & thẩm định thể thức AI ngoại tuyến 100%**.
   - Xem trước trang A4 thời gian thực và **Xuất trực tiếp ra file Microsoft Word (`.doc`)**.

2. **VietNamOCR 3.2 Pro** (`/products/vietnam-ocr.html`):
   - Nhận diện chữ Tiếng Việt từ hình ảnh (`PNG`, `JPG`, `WEBP`) và tài liệu `PDF` scan.
   - Xoay trang, tăng tương phản làm rõ chữ scan mờ, tự động sửa lỗi dấu câu và khoảng trắng tiếng Việt.
   - **Bóc tách & Nối bảng biểu đa trang (Multi-page Table Merge)**.
   - Đồng bộ chỉnh sửa trực tiếp và **Xuất ra Microsoft Word (`.doc`) & Microsoft Excel (`.xls`)**.

## Cấu trúc thư mục

- `index.html` — Trang chủ Software Hub
- `products/vbtm-pro.html` — Trang chi tiết & Phần mềm VBTM Pro Studio
- `products/vietnam-ocr.html` — Trang chi tiết & Phần mềm VietNamOCR 3.2 Studio
- `css/style.css` — Hệ thống giao diện kính trong suốt iOS 18 (Dark/Light Mode)
- `js/main.js`, `js/vbtm-app.js`, `js/ocr-app.js` — Mã nguồn xử lý giao diện và 2 phần mềm
- `downloads/` — Bộ cài đặt Desktop & Word Add-in Macro (`.zip`)
- `desktop-src/` — Mã nguồn Python GUI & Word VBA Macro (`.bas`)
