# MASTER SPECIFICATION: FULL SYSTEM REPLICATION – MONITER

- **Website khảo sát gốc**: https://moniter.io.vn/
- **Các phân hệ Web App trực tuyến đi kèm**:
  - https://play.moniter.io.vn/ (`Sự Kiện Play v2.4.0`)
  - https://dub.moniter.io.vn/ (`VietDub v1.0.19 / v1.0.44`)
- **Ngày hoàn tất khảo sát (Reverse Engineering)**: 16/11/2026
- **Trạng thái**: Hoàn tất Giai đoạn 1 (Kiểm kê toàn diện & Lập kế hoạch kiến trúc)

---

## 1. TỔNG QUAN HỆ THỐNG & DESIGN SYSTEM

### 1.1 Kiến trúc Cổng thông tin chính (`moniter.io.vn`)
Cổng thông tin đóng vai trò là **Software Hub** phân phối các phần mềm độc lập (Desktop `.exe`, Android `.apk`) và kết nối trực tiếp tới các ứng dụng Web/PWA trực tuyến.

- **Thanh điều hướng (Sticky Glass Navbar)**:
  - **Thương hiệu**: Icon щит/checkmark (`/img/brand-icon.svg`) + Tiêu đề + Chấm trạng thái hoạt động (`status-indicator-dot`).
  - **Menu điều hướng**: Danh sách phần mềm (`#productsSection`) và các mục chức năng.
  - **Chuyển đổi Giao diện (`#themeToggleBtn`)**:
    - Lưu trạng thái vào `localStorage` (`moniter-theme`: `dark` | `light`).
    - Hiển thị thông báo Toast thời gian thực khi chuyển chế độ Sáng/Tối.
- **Khu vực Hero & Thống kê**:
  - Các thẻ chỉ số (Metric Pills): Số lượng phần mềm, trạng thái hoạt động `100% Offline / An Toàn`, `0₫ Vì Cộng Đồng`.
  - Lưới hiển thị ứng dụng (`#productsGrid`): Render động từ nguồn dữ liệu cấu trúc (`/data/products.json` hoặc REST API).
- **Hệ thống Đếm Lượt Truy Cập & Lượt Tải Thời Gian Thực (Footer Stats)**:
  - Hiển thị thống kê theo ngày (`Hôm nay: <N> xem • <M> tải`) và theo tháng (`Tháng 11: <X> xem • <Y> tải`).
  - Tự động ghi nhận lượt xem khi tải trang và ghi nhận lượt tải khi người dùng nhấn vào bất kỳ liên kết tải xuống nào (`/downloads/*`), có cơ chế đệm ngoại tuyến (`localStorage` fallback).

### 1.2 Hệ thống Giao diện & Bảng màu (Theme Tokens)
- **Typography**:
  - Phông chữ giao diện chính: `'Plus Jakarta Sans', 'Be Vietnam Pro', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif`
  - Phông chữ mã kiểm định / SHA256: `'JetBrains Mono', monospace`
  - Phông chữ chuẩn văn bản hành chính / Đảng: `'Times New Roman', serif`
- **Hiệu ứng Kính quang học (Liquid Frosted Glassmorphism)**:
  - `backdrop-filter: blur(36px) saturate(220%) contrast(108%)`
  - Độ bo góc: `8px` (`xs`), `14px` (`sm`), `20px` (`md`), `28px` (`lg`), `9999px` (`full`).
- **Tùy biến Chủ đề Đỏ - Vàng Cờ Đảng (Party Flag Red & Gold Theme)**:
  - **Nền chính (Crimson Flag Red)**: `#6b0f12` / `#7f1d1d` / `#450a0a` kết hợp quầng sáng đỏ tươi (`rgba(220, 38, 38, 0.45)`) và ánh vàng sao (`rgba(250, 204, 21, 0.28)`).
  - **Khung kính (Glass Cards)**: Nền đỏ thẫm trong suốt (`rgba(127, 29, 29, 0.58)`), viền vàng kim (`rgba(250, 204, 21, 0.45)`).
  - **Điểm nhấn & Nút bấm (Gold Star Accent)**: Vàng cờ Đảng (`#facc15`, `#eab308`, `#fde047`) kết hợp chữ đỏ sẫm/đen in đậm hoặc nền đỏ cờ (`#dc2626`) viền vàng.

---

## 2. DANH MỤC ĐẦY ĐỦ 06 ỨNG DỤNG TRONG HỆ THỐNG

Qua khảo sát `/data/products.json` và toàn bộ các đường dẫn trên `https://moniter.io.vn/`, hệ thống gồm **06 ứng dụng**:

| STT | Mã ID (`id`) | Tên ứng dụng | Phiên bản | Danh mục | Nền tảng | Dung lượng | Đường dẫn chi tiết |
|---|---|---|---|---|---|---|---|
| 1 | `vbtm-pro` | **VBTM Pro** | `v2.2.5` / `v2.2.4` | `ai` (AI & Văn Bản) | Web Studio + Windows 10/11 & Word | 101 MB (`.exe`) | `/products/vbtm-pro.html` |
| 2 | `vietnam-ocr` | **VietNamOCR 3.2 Pro** | `v3.2.2` / `v3.2.0` | `ai` (AI & Văn Bản) | Web Studio + Windows 10/11 x64 | 147 MB (`.exe`) | `/products/vietnam-ocr.html` |
| 3 | `vietdub` | **VietDub** | `v1.0.19` / `v1.0.44` | `ai` (AI & Video) | Web App (`dub.moniter.io.vn`) & Android APK | 34.8 MB (`.apk`) | `/products/vietdub.html` |
| 4 | `trinh-chay-su-kien` | **Sự Kiện Play** | `v2.4.0` | `audio` (Âm Thanh & Sự Kiện) | Web App & PWA Offline (`play.moniter.io.vn`) | Web App | `/products/trinh-chay-su-kien.html` |
| 5 | `monvpn-pc` | **MonVPN PC (Windows)** | `v1.0.2` | `vpn` (Mạng & Bảo Mật) | Windows 10/11 x64 | 38.4 MB (`.exe`) | `/products/monvpn-pc.html` |
| 6 | `monvpn-android` | **MonVPN Android** | `v1.0.6` | `vpn` (Mạng & Bảo Mật) | Android 8.0+ | 61.6 MB (`.apk`) | `/products/monvpn-android.html` |

---

## 3. PHÂN TÍCH CHI TIẾT TỪNG ỨNG DỤNG (GIAO DIỆN, LUỒNG XỬ LÝ, INPUT/OUTPUT)

### 3.1 Ứng dụng 1: VBTM Pro (`vbtm-pro`)
- **Mục đích**: Chuẩn hóa thể thức văn bản hành chính theo **Nghị định 30/2020/NĐ-CP** (HĐND & UBND Phường) và văn bản của Đảng theo **Hướng dẫn số 05-HD/VPTW** của Văn phòng Trung ương Đảng (Đảng ủy & Chi bộ Phường), chuyển đổi PDF/Word giữ nguyên thể thức, tạo mã QR xác thực ISO/IEC 18004 và rà soát lỗi chính tả Tiếng Việt bằng AI Offline.
- **Thành phần giao diện**:
  1. Khu vực thông tin ứng dụng, nút tải bộ cài `.exe`, mã băm kiểm định SHA256.
  2. **Bộ chọn 16 biểu mẫu chuẩn quy định**:
     - *10 mẫu HĐND & UBND Phường (Nghị định 30/2020/NĐ-CP)*: Công văn, Báo cáo, Quyết định (cá biệt), Quyết định (ban hành Quy chế), Kế hoạch, Tờ trình, Thông báo, Nghị quyết HĐND, Biên bản, Giấy mời.
     - *6 mẫu Đảng ủy & Chi bộ Phường (Hướng dẫn 05-HD/VPTW)*: Công văn Đảng ủy, Báo cáo Đảng ủy, Nghị quyết Đảng ủy, Kế hoạch Đảng ủy, Chương trình công tác, Quyết định Đảng ủy.
  3. **Công cụ Nạp tệp (`.docx`, `.doc`, `.pdf`, `.txt`)**: Hỗ trợ chọn tệp hoặc kéo-thả trực tiếp.
  4. **Trình tạo Mã QR Động (ISO/IEC 18004)**: Render trực tiếp lên `<canvas>` hoặc `PictureBox`, cho phép tải ảnh `.png` hoặc đính kèm vào chân văn bản Word.
  5. **Bộ Kiểm định Chính tả & Thể thức AI**: Quét 535 quy tắc cụm từ hành chính + 7.578 âm tiết Tiếng Việt chuẩn Unicode NFC + quy tắc ngày tháng + dấu kết thúc `./.`.
  6. **Khung xem trước trang giấy A4 thời gian thực (`#a4PreviewPaper`)** & Nút xuất tệp Microsoft Word (`.doc` / `.docx`).
- **Luồng xử lý & Quy chuẩn kỹ thuật bắt buộc (Đặc biệt khi chuyển PDF sang Word)**:
  - **Khổ giấy & Định lề A4**: `210mm x 297mm`; Lề trên `20mm`, Lề dưới `20mm`, Lề trái `30mm`, Lề phải `15mm`; Đánh số trang ở lề trên canh giữa từ trang 2 (ẩn trang 1).
  - **Thể thức Văn bản Đảng theo Hướng dẫn số 05-HD/VPTW**:
    - Góc phải đầu trang: `ĐẢNG CỘNG SẢN VIỆT NAM` (In hoa, cỡ chữ **15pt**, đứng, đậm + đường kẻ ngang nét liền bên dưới).
    - Góc trái đầu trang: Cơ quan cấp trên (In hoa, cỡ chữ **14pt**, đứng), Cơ quan ban hành (In hoa, cỡ chữ **14pt**, đứng, đậm + dấu `*` bên dưới).
    - Số, ký hiệu: `Số ...-CV/ĐU`, `Số ...-BC/ĐU` (Cỡ chữ **14pt**, đứng).
    - Địa danh, ngày tháng năm: Cỡ chữ **14pt**, in nghiêng.
    - Tên loại văn bản: In hoa, cỡ chữ **15pt**, đứng, đậm; Trích yếu dưới tên loại: cỡ chữ **14pt**, đậm; Trích yếu Công văn (`V/v...`): cỡ chữ **12pt**, đứng.
    - Nội dung văn bản: Times New Roman **14pt**, canh đều (`Justify`), thụt đầu dòng `1.27cm`, cách đoạn `6pt`. Tự động nối các dòng bị ngắt xuống hàng trong PDF thành đoạn văn hoàn chỉnh.
    - Thành phần `Nơi nhận`: Từ `<u>Nơi nhận</u>:` cỡ chữ **14pt**, kiểu chữ đứng, có gạch chân dưới từ `"Nơi nhận"`, sau có dấu hai chấm (`:`). Danh sách cơ quan nhận bên dưới cỡ chữ **12pt**, chữ đứng, kết thúc mỗi dòng bằng dấu `;` và dòng cuối (`- Lưu Văn phòng Đảng ủy.`) kết thúc bằng dấu `.`.
    - Quyền hạn & Họ tên người ký: Cỡ chữ **14pt**, đậm (`T/M ĐẢNG ỦY`, `T/M BAN THƯỜNG VỤ`, `BÍ THƯ`...).
- **Dữ liệu Đầu vào (Input)**: Tệp `.pdf`, `.docx`, `.doc`, `.txt` hoặc văn bản nhập trực tiếp.
- **Dữ liệu Đầu ra (Output)**: Tệp `.doc` / `.docx` chuẩn A4 mở trực tiếp bằng Microsoft Word, ảnh `.png` mã QR, hoặc chuẩn hóa trực tiếp trên cửa sổ Microsoft Word đang mở qua COM Automation.

---

### 3.2 Ứng dụng 2: VietNamOCR 3.2 Pro (`vietnam-ocr`)
- **Mục đích**: Nhận diện chữ Tiếng Việt từ hình ảnh (`PNG, JPG, WEBP, BMP, TIF`) và tài liệu `PDF` (cả PDF văn bản số và PDF scan nhiều trang), tái tạo nguyên vẹn bố cục 1:1 sang Word (`.doc`/`.docx`) và bóc tách bảng biểu sang Excel (`.xls`/`.xlsx`).
- **Thành phần giao diện**:
  1. Khung tải tệp ảnh/PDF đa trang + Thanh điều hướng trang PDF (`Trang X / Y`, `Trang trước`, `Trang sau`, `Đọc & Xuất Toàn Bộ Các Trang PDF`).
  2. Công cụ tiền xử lý ảnh: Xoay trái `-90°`, Xoay phải `+90°`, Tăng tương phản làm nét chữ Scan mờ.
  3. Hai chế độ hiển thị kết quả:
     - **Bản xem trước trang Word giữ nguyên thể thức 1:1 (`#ocrFormattedPaper`)**: Hiển thị trực quan bảng 2 cột Quốc hiệu/Đảng - Cơ quan, tiêu đề căn giữa, đoạn văn đã nối dòng căn đều 2 bên, bảng dữ liệu có viền và chân trang 2 cột Nơi nhận - Chữ ký.
     - **Khung văn bản thô (`#ocrResultText`)**.
  4. **Bảng tính chỉnh sửa trực tiếp (`#ocrTableContainer`)**: Cho phép sửa từng ô, thêm hàng, thêm cột, nối bảng biểu đa trang (tự động lọc bỏ dòng tiêu đề lặp `STT | Số ký hiệu...`) và xuất ra Excel.
- **Luồng xử lý**:
  - **Tầng 1 (Direct PDF Coordinate & Regulation Parser)**: Đọc tọa độ `X, Y`, cỡ chữ, kiểu chữ; tự động gom bảng đầu trang (2 cột), nối các dòng bị ngắt trong cùng đoạn văn, gom bảng dữ liệu và gom toàn bộ khối `Nơi nhận` (cột trái) + `Chữ ký` (cột phải) vào bảng chân trang theo đúng chuẩn **Hướng dẫn 05-HD/VPTW** và **Nghị định 30/2020/NĐ-CP**.
  - **Tầng 2 (AI Neural OCR Engine 2 + Tesseract v5 LSTM Tiếng Việt + Windows WinRT OCR)**: Dùng cho ảnh chụp và trang PDF scan thuần ảnh.
- **Dữ liệu Đầu vào (Input)**: File `.pdf`, `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif`, `.webp`.
- **Dữ liệu Đầu ra (Output)**: File Word (`.doc` / `.docx`) giữ nguyên thể thức và file Excel (`.xls` / `.xlsx`).

---

### 3.3 Ứng dụng 3: Sự Kiện Play (`trinh-chay-su-kien` — Tái tạo từ `play.moniter.io.vn`)
- **Mục đích**: Bàn điều hành âm thanh sự kiện (Soundboard) chạy trực tiếp trên trình duyệt và hoạt động ngoại tuyến (Offline PWA) phục vụ hội nghị, đại hội, lễ trao thưởng, chào cờ.
- **Thành phần giao diện & Chức năng thực tế**:
  1. **Thanh Header**: Logo sóng âm, tiêu đề `SỰ KIỆN PLAY`, huy hiệu trạng thái mạng (`Sẵn sàng offline` / `PWA Ready`), nút `Cài đặt` (Thiết lập bàn điều hành) và công tắc `Fade In/Out` tự động.
  2. **6 Thẻ Âm Thanh Sự Kiện Mặc Định (Preset Soundboard Cards)**:
     - *Thẻ 1: Tiếng vỗ tay* (Màu Lime `#84cc16`, Hiệu ứng khán giả, có thể tổng hợp âm thanh thực tế qua Web Audio API + file MP3/WAV).
     - *Thẻ 2: Quốc ca Việt Nam* (Màu Đỏ `#ef4444`, Nghi lễ · Trang trọng).
     - *Thẻ 3: Nhạc trao thưởng 1* (Màu Vàng `#f59e0b`, Victory · Vinh danh).
     - *Thẻ 4: Nhạc trao thưởng 2* (Màu Cam `#f97316`, Vinh danh · Trao giải).
     - *Thẻ 5: Thể thao / Khí thế sôi động* (Màu Cyan `#06b6d4`, Sôi động · Khí thế).
     - *Thẻ 6: Chào mừng Đại biểu* (Màu Tím `#a855f7`, Chào mừng đại biểu).
  3. **Điều khiển trên mỗi thẻ**:
     - Nút `Play / Pause` lớn (hỗ trợ giữ và kéo sang trái để tua nhanh về đầu `0:00`, hoặc bấm phím tắt `1`–`9`, `Space` để dừng khẩn cấp).
     - Đồng hồ thời gian thực `MM:SS / MM:SS`.
     - Thanh tua vị trí phát (`Seek bar` `step="0.1"`).
     - Thanh trượt âm lượng riêng (`Volume slider` `0%`–`100%`, mặc định `80%`) + nút `Hạ âm nhanh (Ducking)` khi MC phát biểu.
     - Hộp kiểm `Lặp lại (Loop)`.
  4. **Modal Thiết Lập Bàn Điều Hành & Nạp Nhạc Cá Nhân**:
     - Bật/tắt hiển thị từng thẻ mặc định.
     - Nút `Thêm thẻ & nhạc của bạn`: Cho phép người dùng tải lên các tệp `MP3, WAV, M4A, AAC, OGG, FLAC, WebM` từ máy tính.
     - Lưu trữ toàn bộ tệp âm thanh tùy chọn vào **IndexedDB** của trình duyệt để lần sau mở lên (kể cả khi mất mạng) vẫn phát được ngay lập tức.

---

### 3.4 Ứng dụng 4: VietDub (`vietdub` — Tái tạo từ `dub.moniter.io.vn`)
- **Mục đích**: Ứng dụng AI dịch phụ đề song ngữ và lồng tiếng/thuyết minh Tiếng Việt tự nhiên cho video đa nền tảng (YouTube, TikTok, Facebook, Web Phim) kèm trình quản lý gói kích hoạt & tải APK.
- **Thành phần giao diện & Chức năng thực tế**:
  1. **Trình Thuyết Minh & Dịch Phụ Đề Video Thực Tế Ngay Trên Web (Live Web Studio) + Mô Phỏng Điện Thoại**:
     - Cho phép người dùng **tải lên tệp video/phụ đề (`.mp4`, `.webm`, `.srt`, `.vtt`) hoặc nhập văn bản/kịch bản ngoại ngữ** để dịch tự động sang Tiếng Việt và phát giọng đọc thuyết minh Tiếng Việt thực tế bằng **Web Speech Synthesis API / TTS Tiếng Việt** đồng bộ với phụ đề trên màn hình!
     - Các công tắc điều khiển thực tế:
       - `Thuyết minh (Bật/Tắt)`
       - `Chặn Quảng Cáo (Bật/Tắt)`
       - `Hạ âm gốc (Ducking: 0% - 100%, mặc định 5% khi có giọng đọc thuyết minh)`
       - `Chọn giọng đọc & tốc độ đọc (0.8x - 1.5x)`
       - Chuyển đổi các tab nguồn video: `YouTube`, `TikTok`, `Facebook`, `Web Phim`, `Tải Video/SRT Của Bạn`.
  2. **Bảng So Sánh Tính Năng & Kho Tải APK / Desktop**:
     - Thông tin phiên bản, mã kiểm định SHA256, tạo mã QR tải nhanh.

---

### 3.5 Ứng dụng 5 & 6: MonVPN PC (`monvpn-pc`) & MonVPN Android (`monvpn-android`)
- **Mục đích**: Công cụ kiểm tra độ trễ mạng (Ping/Latency), kiểm tra địa chỉ IP/DNS an toàn, tạo cấu hình kết nối bảo mật (WireGuard / OpenVPN `.conf` generator) và bộ cài đặt ứng dụng mạng cho Windows (`.exe`) & Android (`.apk`).
- **Thành phần giao diện & Chức năng thực tế**:
  1. **Trình Kiểm Tra Mạng & Điều Khiển Kết Nối Trực Tuyến (Interactive Network & Tunnel Diagnostic Studio)**:
     - Kiểm tra thực tế địa chỉ IP công khai, nhà mạng, độ trễ (Ping/Jitter ms) tới các máy chủ Việt Nam (Hà Nội, Đà Nẵng, TP.HCM) và Quốc tế (Singapore, Tokyo) bằng HTTP Timing thực tế.
     - Công tắc kết nối 1-Chạm, biểu đồ thông lượng thời gian thực và trình xuất tệp cấu hình bảo mật (`.conf`) thực tế.
  2. **Bộ cài đặt Windows Native (`MonVPN-Setup-1.0.2.exe`)**:
     - Ứng dụng Windows Forms thực tế có giao diện kiểm tra Ping, tốc độ mạng, DNS Flush/Check và kết nối máy chủ.

---

## 4. PHÂN LOẠI CHỨC NĂNG CLIENT-SIDE, BACKEND & DESKTOP

| Thành phần | Chạy trên Trình duyệt (Client-Side) | Chạy trên Backend (Node.js / API) | Chạy trên Máy tính (Windows `.exe`) |
|---|---|---|---|
| **Cổng Software Hub** | Giao diện, lọc danh mục, tìm kiếm, đổi chủ đề Đỏ-Vàng / Sáng-Tối | API lưu/đọc danh sách ứng dụng (`/api/products`), thống kê lượt xem/tải (`/api/stats`) | — |
| **VBTM Pro** | Đọc `.docx`/`.pdf`, chuẩn hóa NĐ30 & HD05, tạo QR ISO 18004, soát lỗi 7.578 âm tiết, xuất `.doc` | API lưu cấu hình biểu mẫu & nhật ký kiểm định | App `.exe` + Tự động điều khiển Microsoft Word qua COM Automation (`Word.Application`) |
| **VietNamOCR Pro** | Đọc PDF tọa độ 1:1, OCR ảnh/scan (Tesseract + Neural OCR), bảng tính Excel, xuất `.doc`/`.xls` | API hỗ trợ xử lý file & lưu lịch sử | App `.exe` tích hợp kết xuất PDF (`Windows.Data.Pdf` + Ghostscript) & OCR Tiếng Việt |
| **Sự Kiện Play** | Bàn điều hành âm thanh Web Audio API, lưu nhạc tùy chọn vào `IndexedDB`, phím tắt `1-6`, Fade In/Out | API quản lý danh sách preset âm thanh | — |
| **VietDub** | Trình phát video + dịch phụ đề `.srt`/`.vtt` + thuyết minh giọng đọc TTS Tiếng Việt & hạ âm gốc | API quản lý cấu hình & gói dịch vụ | — |
| **MonVPN PC & Android** | Đo độ trễ (Ping/Latency) thực tế, kiểm tra bảo mật kết nối, tạo file cấu hình `.conf` | API danh sách máy chủ & kiểm tra trạng thái | App `.exe` kiểm tra đường truyền, đo Ping & tối ưu DNS trên Windows |
| **Admin Dashboard** | Giao diện quản trị CRUD ứng dụng, danh mục, biểu mẫu, thống kê | REST API xác thực & lưu trữ dữ liệu (`products.json` / SQLite / PostgreSQL) | — |

---

## 5. KẾ HOẠCH TRIỂN KHAI TỪNG GIAI ĐOẠN

1. **Giai đoạn 1 (Đã hoàn tất)**: Khảo sát toàn bộ hệ thống `moniter.io.vn` + các trang con và lập tài liệu `MASTER_SPECIFICATION.md`.
2. **Giai đoạn 2 (Thiết kế Giao diện Đỏ - Vàng Cờ Đảng & Kiến trúc)**:
   - Chuyển đổi toàn bộ giao diện Web và Desktop App sang **tông màu Đỏ - Vàng Cờ Đảng** trang trọng, sắc nét.
   - Tổ chức cấu trúc dữ liệu chuẩn hóa hỗ trợ đầy đủ cả chế độ chạy tĩnh (GitHub Pages) và chạy qua máy chủ Backend Node.js (`server.js` cung cấp REST API `/api/products`, `/api/stats`, `/api/admin` + Admin Dashboard).
3. **Giai đoạn 3 (Hoàn thiện & Khắc phục triệt để lỗi của từng App)**:
   - **VBTM Pro & VietNamOCR Pro**: Sửa triệt để bộ máy chuyển đổi PDF sang Word (cả Web và Desktop `.exe`) để tự động nối dòng đoạn văn, giữ nguyên bảng 2 cột đầu/cuối trang và áp dụng chuẩn **Hướng dẫn số 05-HD/VPTW** (`ĐẢNG CỘNG SẢN VIỆT NAM` 15pt đậm, `<u>Nơi nhận</u>:` 14pt đứng gạch chân, danh sách nơi nhận 12pt...) cùng **Nghị định 30/2020/NĐ-CP**.
   - **Sự Kiện Play, VietDub, MonVPN**: Xây dựng đầy đủ trang chi tiết và ứng dụng tương tác thực tế (Web Audio Soundboard có nạp nhạc IndexedDB, Trình thuyết minh phụ đề, Trình kiểm tra mạng & bộ cài `.exe`).
4. **Giai đoạn 4 & 5 (Backend API & Admin Dashboard)**:
   - Xây dựng máy chủ Node.js (`server.js`) và trang quản trị `/admin.html` cho phép Thêm / Sửa / Xóa ứng dụng, quản lý danh mục, cập nhật liên kết tải xuống và theo dõi thống kê lượt xem/lượt tải thực tế.
5. **Giai đoạn 6 & 7 (Kiểm thử tự động, Lập báo cáo `APP_TEST_REPORT.md` & Triển khai)**:
   - Chạy kiểm thử toàn bộ chức năng, biên dịch lại toàn bộ các file `.exe`, ghi nhận vào `APP_TEST_REPORT.md` và đẩy lên GitHub.
