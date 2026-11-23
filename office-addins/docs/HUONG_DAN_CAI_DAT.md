# HƯỚNG DẪN CÀI ĐẶT VBTM PRO OFFICE INTEGRATION (WINDOWS PERSONAL EDITION)

## 1. Yêu cầu hệ thống
- **Hệ điều hành**: Windows 10 / Windows 11 (32-bit hoặc 64-bit).
- **Microsoft Office**: Office 2016, Office 2019, Office 2021, Office 2024 hoặc Microsoft 365 (Click-to-Run hoặc MSI).
- **Trình duyệt nhúng**: Microsoft Edge WebView2 Runtime (tự kiểm tra bởi `VBTM_Pro_Office_Setup.exe`).
- **Kết nối mạng**: HTTPS tới `https://anhtrung4566.github.io/software-hub/office-addins/` và `https://generativelanguage.googleapis.com` (hoặc máy chủ `VBTM.Backend` nội bộ).

## 2. Cài đặt bằng `VBTM_Pro_Office_Setup.exe` (Phạm vi cá nhân — không cần quyền Admin)
1. Tải và chạy `downloads/VBTM_Pro_Office_Setup.exe`.
2. Bấm **1. Kiểm tra Office & HTTPS** để kiểm tra tự động: Windows, phiên bản & build Office, 32/64-bit, Word/Excel/PowerPoint/Outlook (Classic vs New), WebView2 và kết nối HTTPS.
3. Bấm **2. Cài đặt / Cập nhật**: bộ cài sẽ sao lưu bản cũ (nếu có), **giữ nguyên API Key cá nhân**, và xuất 4 tệp XML Manifest chuẩn (`Add-in-only XML Manifest`) vào `%LOCALAPPDATA%\VBTMProOffice\manifests\`:
   - `vbtm-word-manifest.xml`
   - `vbtm-excel-manifest.xml`
   - `vbtm-powerpoint-manifest.xml`
   - `vbtm-outlook-manifest.xml`
   - **Cam kết bảo mật**: Bộ cài **không tự ý sửa Registry Trust Center** và **không sử dụng network share sideload** làm phương thức phát hành chính thức.

## 3. Nạp Manifest vào từng ứng dụng Office
### A. Microsoft 365 / Office on the Web
1. Mở Word / Excel / PowerPoint → Tab **Home** (hoặc **Insert**) → **Add-ins** → **More Add-ins** → **Upload My Add-in**.
2. Chọn tệp XML Manifest tương ứng trong `%LOCALAPPDATA%\VBTMProOffice\manifests\`.
3. Biểu tượng **VBTM PRO AI** xuất hiện trên thanh Ribbon (Tab Home).

### B. Classic Outlook & New Outlook
1. Mở Outlook → **All Apps** → **Get Add-ins** (hoặc truy cập `https://aka.ms/olksideload`) → **My add-ins** → **Add a custom add-in** → **Add from file...**.
2. Chọn `vbtm-outlook-manifest.xml`.

### C. Office 2016 / 2019 / 2021 / 2024 Desktop cá nhân
1. Mở Word / Excel / PowerPoint → **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**.
2. Chủ động thêm thư mục danh mục tin cậy chứa manifest (do người dùng tự xác nhận), tích **Show in Menu**, khởi động lại Office → **Insert** → **My Add-ins** → chọn **VBTM PRO AI**.

## 4. Cấu hình AI (2 chế độ)
- **Chế độ Cơ quan (Khuyên dùng)**: Chạy `office-addins/VBTM.Backend/server.js` với biến môi trường `GEMINI_API_KEY` (có thể bật HTTPS qua `VBTM_TLS_KEY` và `VBTM_TLS_CERT`), sau đó chọn chế độ *VBTM.Backend AI Gateway* trong tab **Cài đặt** của Task Pane.
- **Chế độ Cá nhân (BYOK)**: Chọn *Tự nhập Gemini API Key*, nhập Key từ Google AI Studio. Key được lưu cục bộ trong `localStorage` của Office WebView2 và gửi trực tiếp qua HTTPS tới Gemini API (không nhúng trong mã nguồn/manifest).

## 5. Khôi phục & Gỡ cài đặt sạch
- Mở `VBTM_Pro_Office_Setup.exe` → bấm **Khôi phục bản cũ** để quay lại bản sao lưu trước, hoặc **Gỡ cài đặt sạch** để xóa toàn bộ `%LOCALAPPDATA%\VBTMProOffice` mà không làm mất tài liệu Office hay API Key.
