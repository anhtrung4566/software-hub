# Hệ Thống Xử Lý Tài Liệu Thông Minh (VietNamOCR 3.5 Pro — IDP Edition)

## 1. Kiến trúc & Các module đã tích hợp thực tế
- **`js/ocr-engine-pro.js`**:
  - Tự động phân loại tài liệu trước OCR (`classifyDocumentCanvas`) và chọn `tessedit_pageseg_mode` phù hợp (`3` / `6`).
  - Tự động căn thẳng trang bị nghiêng (`Auto-Deskew ±6°`), tẩy dấu mộc đỏ, Unsharp Mask 3×3 và **tự động chạy lại bằng bộ lọc Otsu khi độ tin cậy < 72%**.
  - Đánh dấu trực tiếp các vùng độ tin cậy thấp (`< 70%`) lên ảnh gốc (`drawConfidenceOverlayOnCanvas`) và phát hiện ký tự dễ nhầm lẫn (`O/0`, `I/l/1`, `S/5`).
  - Trích xuất bảng biểu nâng cao từ tọa độ `word bbox` (`extractTableFromWordBboxes`) và phát hiện ô có độ tin cậy thấp.
  - Hỗ trợ hàng đợi nhiều file có **Tạm dừng / Tiếp tục / Hủy**, kiểm tra an toàn đầu vào (tối đa 80MB) và nhật ký lỗi kỹ thuật không lưu nội dung nhạy cảm.
- **`js/ocr-2layer-pdf.js`**:
  - Tạo **Searchable PDF 2 lớp** cho **tất cả các trang**: Lớp 1 giữ nguyên ảnh scan & con dấu đỏ gốc; Lớp 2 nhúng font TrueType Unicode tiếng Việt (`opacity: 0`) với tự động co giãn theo bề rộng `bbox` (`widthOfTextAtSize`) để bôi đen, copy và `Ctrl+F` khớp từng dòng.
- **`js/ocr-ai-proofread.js`**:
  - Đối chiếu song song bản OCR gốc vs từng thẻ đề xuất sửa (**Chấp nhận / Từ chối** riêng lẻ).
  - **Bảo vệ thực thể quan trọng (`detectProtectedEntityReason`)**: Không tự động sửa Họ tên, Địa chỉ, CCCD, Mã số thuế, Số hợp đồng, Số văn bản, Ngày tháng năm, Số tiền, Tỷ lệ %, Ký hiệu pháp lý.
- **`js/ocr-idp-studio.js`**:
  - **AI Hiểu & Phân tích tài liệu có dẫn chiếu `[Đoạn #]`**: Trích xuất thực thể, Phân loại tài liệu, Tạo Mục lục tự động, Phát hiện đoạn trùng lặp, Tóm tắt & Hỏi đáp (Ollama Local + fallback Offline trích dẫn nguyên văn), So sánh 2 tài liệu.
  - **Xuất đa định dạng thực tế**: Searchable PDF 2 lớp, **Word `.docx` (OOXML)**, Word `.doc`, Excel `.xls`, **Markdown `.md`**, **JSON có cấu trúc `.json`**, và **TXT `.txt`**.
  - **Kho lưu trữ & Tìm kiếm thông minh cục bộ (`localStorage`)**: Tự động lưu sau khi OCR, tìm theo tên file / từ khóa / ngày tháng / số văn bản / họ tên, mở lại ngay, Sao lưu / Khôi phục `.json` và nút **Xóa sạch dữ liệu OCR & bộ nhớ tạm**.

## 2. Đánh giá trung thực: Chức năng đã hoàn thành & Giới hạn thực tế
- **Đã hoàn thành và hoạt động thực tế 100% cục bộ trên trình duyệt**: OCR đa ngôn ngữ (`vie` `tessdata_best`, `vie+eng`, `eng`), Auto-Deskew, nhận diện TIFF/PDF nhiều trang, xuất PDF 2 lớp đầy đủ mọi trang, DOCX/MD/JSON/TXT/Excel, AI hiệu chỉnh có bảo vệ thực thể, IDP Q&A dẫn chiếu đoạn, Kho lưu trữ tìm kiếm & Sao lưu/Xóa sạch.
- **Giới hạn kỹ thuật hiện tại (báo cáo minh bạch)**:
  1. Do chạy hoàn toàn bằng WebAssembly trong trình duyệt (không cần cài Python/CUDA), tài liệu scan quá lớn (> 50–100 trang độ phân giải 300 DPI) có thể tốn nhiều RAM trình duyệt; khuyến nghị tách lô 20–30 trang/lần.
  2. Chữ viết tay hoặc bản scan quá mờ (< 120 DPI) vẫn cần người dùng đối chiếu các khung màu vàng/đỏ đánh dấu trên ảnh gốc.
  3. Để dùng tính năng **AI Tóm tắt / Hỏi đáp ngữ nghĩa nâng cao**, cần bật **Ollama Local** trên máy (`qwen2.5:3b` hoặc `qwen2.5:7b`); nếu chưa bật Ollama, hệ thống tự chuyển sang chế độ tra cứu dẫn chiếu nguyên văn Offline (không gửi dữ liệu ra máy chủ ngoài).
