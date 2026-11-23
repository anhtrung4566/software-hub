# BẢNG TƯƠNG THÍCH TỰ ĐỘNG — VBTM PRO OFFICE INTEGRATION

Hệ thống tự động xác định mức tương thích (`VBTM.AICore/vbtm-compat.js` và `VBTM.CompatibilityChecker/CompatibilityChecker.cs`) dựa trên **API Requirement Sets thực tế** và **số Build của Office**, không chỉ dựa vào tên phiên bản.

| Mức tương thích | Phiên bản Office điển hình | Word (`WordApi`) | Excel (`ExcelApi`) | PowerPoint (`PowerPointApi`) | Outlook (`Mailbox`) |
|---|---|---|---|---|---|
| **MỨC 1 – BASIC** | Office 2016 MSI / Build cũ (< 10325) | Common API: đọc/thay/chèn đoạn đang chọn an toàn | `ExcelApi 1.1`: đọc vùng chọn, phân tích, ghi ô có xác nhận | Common API: đọc/chèn khung chọn & sao chép dàn ý | `Mailbox 1.5`: đọc email, tóm tắt, soạn & chèn nội dung |
| **MỨC 2 – STANDARD** | Office 2019 / C2R (Build 10325 – 14325) | `WordApi 1.1–1.3`: đủ 18 chức năng + quét toàn văn bản theo ID đoạn + nhảy tới vị trí lỗi | `ExcelApi 1.4`: thêm tổng hợp nhiều Sheet & tạo Sheet báo cáo mới | `PowerPointApi 1.2`: thêm tạo slide mới không xóa slide cũ | `Mailbox 1.8`: hỗ trợ đầy đủ Classic Outlook (Read & Compose) |
| **MỨC 3 – ADVANCED** | Office 2021 / Office 2024 / Microsoft 365 (Build ≥ 14326) | `WordApi 1.4+`: đầy đủ toàn bộ tính năng nâng cao & bảo toàn định dạng OOXML | `ExcelApi 1.8+`: đầy đủ tính năng nâng cao nhiều Sheet & công thức | `PowerPointApi 1.4+`: đầy đủ tạo slide & biên tập nâng cao | `Mailbox 1.12+`: hỗ trợ cả Classic Outlook (WebView2) & New Outlook |

## Nguyên tắc hạ cấp an toàn (Graceful Degradation)
- Nếu một Requirement Set không được hỗ trợ trên phiên bản Office cũ, Add-in tự động ẩn hoặc chuyển sang phương án thay thế (Common API / Sao chép) và hiển thị thông báo rõ ràng trên thanh trạng thái — **không làm lỗi toàn bộ Add-in**.
- Sử dụng **Add-in-only XML Manifest** (`VersionOverridesV1_0`) để tương thích rộng từ Office 2016 vĩnh viễn tới Microsoft 365.
