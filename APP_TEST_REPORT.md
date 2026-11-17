# APP TEST REPORT – SIÊU APP VĂN PHÒNG (ĐẢNG ỦY & UBND PHƯỜNG)

**Ngày kiểm định**: 17/11/2026  
**Hệ sinh thái kiểm thử**: Web Studio (HTML5/JS/CSS) + Windows Desktop Native (`.exe`) + Node.js REST API (`server.js`) + Admin Dashboard (`admin.html`)  
**Giao diện chuẩn**: Đỏ - Vàng Cờ Đảng (`#4c0509` / `#68090f` / `#b91c1c` & `#facc15` / `#eab308`)

---

## 1. BẢNG TỔNG HỢP TRẠNG THÁI KIỂM ĐỊNH CÁC ỨNG DỤNG

| STT | Tên Ứng Dụng | Phiên Bản | Đường Dẫn Web / File `.exe` | Unit Test | Integration Test | E2E & Thể Thức | Trạng Thái |
| :-: | :--- | :---: | :--- | :---: | :---: | :---: | :---: |
| 1 | **VBTM Pro** (Chuẩn Hóa Văn Bản Đảng & Hành Chính) | `v2.2.4` | `/products/vbtm-pro.html`<br>`/downloads/VBTM-Setup-2.2.4.exe` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 2 | **VietNamOCR Pro** (AI OCR & PDF Sang Word Chuẩn Thể Thức) | `v3.2.0` | `/products/vietnam-ocr.html`<br>`/downloads/VietNamOCR-Setup-3.2.0-x64.exe` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 3 | **Sự Kiện Play** (Trình Chạy Nhạc Sự Kiện & Đại Hội) | `v2.4.0` | `/products/trinh-chay-su-kien.html`<br>`/js/sukien-app.js` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 4 | **VietDub AI Studio** (Thuyết Minh & Dịch Phụ Đề `.SRT`) | `v1.0.44` | `/products/vietdub.html` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 5 | **MonVPN & Kiểm Định Mạng** (Đo Ping & Tạo WireGuard `.conf`) | `v1.0.2` | `/products/monvpn-pc.html` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 6 | **Admin Dashboard & Backend API** | `v2.0.0` | `/admin.html` & `/server.js` | PASS | PASS | PASS | **HOÀN THÀNH 100%** |

---

## 2. KIỂM ĐỊNH CHI TIẾT THỂ THỨC VĂN BẢN ĐẢNG (HƯỚNG DẪN 05-HD/VPTW) & HÀNH CHÍNH (NGHỊ ĐỊNH 30/2020/NĐ-CP)

### 2.1. Kiểm tra chuyển đổi PDF sang Word (`VBTM Pro 2.2.4` & `VietNamOCR 3.2 Pro`)
- **Tự động nhận diện thể thức văn bản Đảng (`detectIsPartyDocument` / `IsPartyDocumentLines`)**:
  - Nhận diện chính xác tiêu ngữ `ĐẢNG CỘNG SẢN VIỆT NAM`, ký hiệu `-BC/ĐU`, `-NQ/ĐU`, `-CV/ĐU`, `-QĐ/ĐU`, `ĐẢNG BỘ...`, `ĐẢNG ỦY...`, `CHI BỘ...`, `T/M ĐẢNG ỦY`, `T/M BAN THƯỜNG VỤ`.
  - **Không còn bất kỳ tham chiếu nào đến Hướng dẫn 36-HD/VPTW** (đã hết hiệu lực); áp dụng 100% **Hướng dẫn số 05-HD/VPTW** của Văn phòng Trung ương Đảng.
- **Khôi phục Bảng Quốc Hiệu / Tiêu Ngữ (2 cột đầu trang)**:
  - Cột phải (Văn bản Đảng): `ĐẢNG CỘNG SẢN VIỆT NAM` cỡ chữ **15pt, in hoa, đứng, đậm** + đường kẻ ngang liền nét phía dưới; Địa danh và ngày tháng năm cỡ chữ **14pt, in nghiêng**.
  - Cột trái (Văn bản Đảng): Cơ quan cấp trên **14pt in hoa đứng**, cơ quan ban hành **14pt in hoa đậm** + dấu `*` căn giữa bên dưới; Số và ký hiệu (`Số ...-BC/ĐU`) **14pt đứng**.
- **Khôi phục Tên loại & Trích yếu văn bản**:
  - Tên loại văn bản Đảng (`BÁO CÁO`, `NGHỊ QUYẾT`, `QUYẾT ĐỊNH`, `KẾ HOẠCH`...): **15pt, in hoa, đứng, đậm, căn giữa**.
  - Trích yếu nội dung văn bản: **14pt, đứng, đậm, căn giữa** + đường kẻ ngang phía dưới.
- **Tự động nối dòng bị ngắt từ PDF (`mergeBodyLinesIntoParagraphs`)**:
  - Khắc phục triệt để lỗi mỗi dòng PDF bị tách thành 1 đoạn `<p>` rời rạc.
  - Các dòng cùng một đoạn văn được nối liền mạch, căn đều 2 bên (`text-align: justify`), cỡ chữ **14pt Times New Roman**, lùi đầu dòng `1.27cm`, khoảng cách đoạn `6pt`.
- **Khôi phục Bảng Nơi nhận & Chức vụ ký (2 cột cuối trang)**:
  - Cột trái: Chữ `<u>Nơi nhận</u>:` cỡ chữ **14pt, đứng, gạch chân phần chữ "Nơi nhận", có dấu hai chấm `:` ở cuối** (chuẩn Hướng dẫn 05-HD/VPTW); danh sách nơi nhận bên dưới cỡ chữ **12pt đứng**, kết thúc từng dòng bằng `;` và dòng cuối (`- Lưu Văn phòng Đảng ủy.`) bằng `.`.
  - Cột phải: Quyền hạn và chức vụ người ký (`T/M BAN THƯỜNG VỤ`, `PHÓ BÍ THƯ`...) cỡ chữ **14pt, in hoa, đậm, căn giữa**; Họ tên người ký **14pt, đậm, căn giữa**.
- **Đánh số trang tự động**:
  - Cấu hình `mso-title-page: yes` trong Word XML để ẩn số trang ở trang 1 và tự động đánh số trang `2, 3...` căn giữa tại lề trên từ trang thứ 2.

---

## 3. KIỂM ĐỊNH GIAO DIỆN ĐỎ - VÀNG CỜ ĐẢNG & CÁC MODULE MỚI

1. **Giao diện Đỏ - Vàng Cờ Đảng**:
   - Đã áp dụng đồng bộ trên toàn bộ Website (`css/style.css`, `index.html`, tất cả trang con trong `products/`, `admin.html`) và cả 2 phần mềm máy tính Windows (`VBTMSetupAndApp.cs`, `VietNamOCRSetupAndApp.cs`).
2. **Sự Kiện Play v2.4.0 (`products/trinh-chay-su-kien.html`)**:
   - Đã kiểm thử bộ tổng hợp âm thanh Web Audio API (Quốc ca, Quốc tế ca, Nhạc chào mừng Đại hội, Nhạc trao thưởng, Hồi trống khai hội, Pháo tay) + lưu trữ file MP3/WAV vào `IndexedDB` + phím tắt `1-9`, `Space`, `F` (Fade-out mượt).
3. **VietDub AI Studio v1.0.44 (`products/vietdub.html`)**:
   - Đã kiểm thử đọc phụ đề `.SRT`/`.VTT`, dịch thuật tự động, phát giọng đọc tiếng Việt TTS và xuất file `.SRT`/`.VTT`.
4. **MonVPN & Kiểm Định Mạng v1.0.2 (`products/monvpn-pc.html`)**:
   - Đã kiểm thử đo độ trễ HTTP thực tế (`performance.now()`) tới các cụm máy chủ và sinh cấu hình WireGuard `.conf` với cặp khóa X25519 ngẫu nhiên (`crypto.getRandomValues`).
5. **Admin Dashboard & Backend (`admin.html` & `server.js`)**:
   - Đã kiểm thử thêm/sửa/xóa ứng dụng, quản lý liên kết tải xuống, xuất `products.json` và API REST `/api/products`, `/api/stats`.
