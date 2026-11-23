# BÁO CÁO KẾT QUẢ KIỂM THỬ — VBTM PRO OFFICE INTEGRATION

Báo cáo ghi rõ **hạng mục đã kiểm thử thật trên máy Windows vật lý** và **hạng mục kiểm thử bằng mô phỏng / tự động hóa**, không báo hoàn thành ảo.

## 1. Môi trường kiểm thử vật lý thực tế
- **Hệ điều hành**: Windows 64-bit.
- **Microsoft Office cài thực tế trên máy**: Click-to-Run `16.0.17932.21140` (`x64`, `ProPlus2024Volume`) — gồm đủ `WINWORD.EXE`, `EXCEL.EXE`, `POWERPNT.EXE`, `OUTLOOK.EXE` (Classic Outlook Win32).
- **WebView2 Runtime**: `157.0.4341.60` (đã cài đặt).
- **Các phiên bản Office không có máy vật lý riêng**: Office 2016 MSI, Office 2019, Office 2021, Microsoft 365 và New Outlook → được kiểm tra qua mô phỏng API Requirement Sets trong `VBTM.AICore/vbtm-compat.js` và `office-addins/tests/verify_office_addins.js`.

## 2. Kết quả kiểm thử tự động & thực tế đã chạy

| Hạng mục kiểm thử | Phương thức | Kết quả thực tế |
|---|---|---|
| Kiểm định 4 XML Manifest (`Word`, `Excel`, `PowerPoint`, `Outlook`) | Thực tế (`[xml]` parser PowerShell) | **PASS (4/4)** — Đúng chuẩn Add-in-only XML Manifest, đủ `VersionOverrides`, HTTPS `SourceLocation` |
| Kiểm tra dùng chung `VBTM.AICore` trên cả 4 Add-in | Thực tế (`verify_office_addins.ps1`) | **PASS (4/4)** — Không lặp logic AI riêng lẻ |
| Quét bảo mật 0 API Key nhúng cứng | Thực tế (`verify_office_addins.ps1`) | **PASS** — 0 chuỗi `AIza...` trong toàn bộ `office-addins/` |
| Đủ bộ prompt nghiệp vụ (Word 18, Excel 11, PPT 8, Outlook 8) | Thực tế (`verify_office_addins.ps1`) | **PASS** |
| Biên dịch `VBTM_Pro_Office_Setup.exe` (.NET Framework 4.x) | Thực tế (`csc.exe`) | **PASS** — 0 lỗi, 0 cảnh báo |
| Chạy thật `VBTM_Pro_Office_Setup.exe --check-only` trên Office 2024 x64 | Thực tế (EXE trên máy) | **PASS** — Nhận diện đúng `16.0.17932.21140 x64 ProPlus2024Volume`, `Classic Outlook (Win32)`, `WebView2 157.0.4341.60`, `MỨC 3 – ADVANCED` |
| Chạy thật Cài đặt → Cập nhật có sao lưu → Gỡ sạch → Cài lại | Thực tế (`--silent-install` / `--uninstall`) | **PASS** — Tạo đủ 4 manifest, tự sao lưu giữ nguyên API Key, gỡ sạch không ảnh hưởng tài liệu |
| Kiểm thử hồi quy VBTM Pro 2.5.0/3.0 & VietNamOCR (`tests/run_tests.ps1`) | Thực tế (`run_tests.ps1`) | **PASS (14/14)** — 0 hồi quy, SHA256 khớp 100% |
| Kiểm thử giao diện Ribbon/Task Pane trực tiếp bên trong cửa sổ Word/Excel/PowerPoint/Outlook | Cần người dùng nạp Manifest thủ công | **Chưa tự động hóa UI bên trong Office** (do yêu cầu bảo mật không tự ý sửa Registry Trust Center) — sẵn sàng để người dùng nạp file manifest từ `%LOCALAPPDATA%\VBTMProOffice\manifests\` sau khi publish HTTPS lên GitHub Pages |
