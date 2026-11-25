# Kiến Trúc & Hướng Dẫn Hệ Thống OCR Tích Hợp (VietNamOCR 3.5 Pro)

## 1. Đánh giá & Lựa chọn Công nghệ OCR Tiếng Việt

| Công nghệ | Độ chính xác dấu Tiếng Việt | Nhận diện bố cục / bảng | PDF 2 lớp (Searchable) | Chạy Offline / Triển khai trên Web + Windows |
| :--- | :--- | :--- | :--- | :--- |
| **PaddleOCR / RapidOCR** | Cao (PP-OCRv4/v5) | Rất tốt (DBNet) | Cần ghép `pdf-lib` / `pypdf` | Yêu cầu Python + ONNX/Paddle runtime (~300–600MB) |
| **VietOCR** | Rất cao (Transformer) | Không tự tách dòng (cần detector riêng) | Cần detector + PDF builder | Yêu cầu Python + PyTorch |
| **OCRmyPDF** | Trung bình (dùng Tesseract mặc định) | Tốt (auto-deskew) | Tốt | Yêu cầu Python + Ghostscript + Tesseract CLI |
| **Tesseract.js v5 (`tessdata_best`) + `pdf-lib` + `fontkit` + Ollama Local** | **Cao với `tessdata_best` + tiền xử lý ảnh & hậu kiểm AI Local** | **Tái tạo 2 cột NĐ 30 / 05-HD bằng tọa độ `bbox` thực** | **100% Searchable PDF 2 lớp chuẩn Unicode (khớp từng dòng)** | **Chạy trực tiếp 100% trên trình duyệt & Windows, không cần cài Python/server ngoài** |

### Phương án kiến trúc được lựa chọn và triển khai trực tiếp trong mã nguồn
1. **Nhân OCR (`js/ocr-engine-pro.js`)**:
   - Sử dụng **Tesseract.js v5 WebAssembly Worker** tái sử dụng bộ nhớ (persistent worker) với kho mô hình chính xác cao **`tessdata_best`** (`vie`, `vie+eng`, `eng`) và tự động fallback nếu mạng hạn chế.
   - **Tiền xử lý ảnh scan thực tế**: Tự động phát hiện và xoay thẳng ảnh nghiêng **Auto-Deskew (±6°)**, tẩy dấu mộc đỏ (giữ nguyên mực đen bên dưới), nâng tương phản mềm bảo toàn nét mảnh dấu tiếng Việt và lọc sắc nét **Unsharp Mask 3×3**.
   - **Đa định dạng & Hàng đợi đa tệp**: Hỗ trợ `JPG`, `JPEG`, `PNG`, `WEBP`, `BMP`, **`TIFF` (`.tif/.tiff` qua `UTIF.js`)** và **PDF nhiều trang**, hỗ trợ chọn/kéo thả nhiều tệp cùng lúc.
2. **Xuất PDF 2 Lớp có thể tìm kiếm (`js/ocr-2layer-pdf.js`)**:
   - **Lớp 1**: Ảnh scan gốc giữ nguyên 100% con dấu đỏ, chữ ký, bảng biểu.
   - **Lớp 2**: Lớp chữ ẩn (`opacity: 0`) nhúng phông TrueType Unicode tiếng Việt qua `pdf-lib` + `fontkit`, tự co giãn cỡ chữ theo bề rộng `bbox` thực (`widthOfTextAtSize`) để bôi đen, sao chép và `Ctrl+F` khớp chính xác trên **tất cả các trang**.
3. **AI Kiểm tra & Hiệu chỉnh OCR có Bảo vệ Thực thể (`js/ocr-ai-proofread.js`)**:
   - Kết hợp lớp kiểm tra **Offline tức thì** + **Ollama Local (`qwen2.5:7b` / `qwen2.5:3b`)** (không gửi tài liệu ra máy chủ bên ngoài).
   - **Bảo vệ tuyệt đối dữ liệu quan trọng**: Tự động gắn cờ `🛡 BẢO VỆ` và không bao giờ tự ý sửa **Họ tên, Địa chỉ, CCCD, Mã số thuế, Số hợp đồng, Số văn bản, Ngày tháng năm, Số tiền, Tỷ lệ %, Ký hiệu pháp lý**.
   - Cho phép đối chiếu song song **Bản OCR gốc** với từng đề xuất sửa và **Chấp nhận / Từ chối** riêng lẻ.

## 2. Hướng dẫn sử dụng nhanh
1. Mở trang **VietNamOCR 3.5 Pro** (`products/vietnam-ocr.html`).
2. Chọn ngôn ngữ nhận diện (`Tiếng Việt — tessdata_best`, `Song ngữ Việt + Anh`, hoặc `Tiếng Anh`) và bật **`Tự căn thẳng ảnh nghiêng`**.
3. Bấm **`📂 Chọn 1 hoặc Nhiều File (PDF / Ảnh / TIFF)`** hoặc kéo thả tệp vào khung.
4. Sau khi OCR xong:
   - Bấm **`🤖 AI Hiệu Chỉnh OCR`** để rà soát và duyệt từng đề xuất sửa lỗi OCR.
   - Bấm **`📑 Xuất PDF 2 Lớp (Searchable PDF)`** để tải về PDF tìm kiếm & sao chép tiếng Việt đầy đủ mọi trang.
   - Bấm **`Xuất Word (.doc)`**, **`📄 Xuất TXT`**, hoặc **`Xuất Bảng Excel (.xls)`** tùy nhu cầu.
