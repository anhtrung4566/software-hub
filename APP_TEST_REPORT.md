# APP TEST REPORT – HỆ THỐNG PHẦN MỀM VĂN PHÒNG (ĐẢNG ỦY & UBND PHƯỜNG)

**Ngày kiểm định**: 17/11/2026  
**Phạm vi ứng dụng**: **VBTM Pro 2.5.0** & **VietNamOCR 3.5 Pro** + **Admin Dashboard** + **Backend Server** *(Đã gỡ bỏ hoàn toàn 3 ứng dụng phụ: Sự Kiện Play, VietDub AI Studio và MonVPN theo yêu cầu)*  
**Giao diện chuẩn**: Đỏ - Vàng Cờ Đảng (`#4c0509` / `#68090f` / `#b91c1c` & `#facc15` / `#eab308`)  
**Bộ chạy kiểm thử tự động**: `tests/run_tests.ps1` & `tests/test_suite.js` (**Kết quả: 14/14 PASS - 100%**)

---

## 1. BẢNG TỔNG HỢP TRẠNG THÁI TỪNG ỨNG DỤNG

| STT | Tên Ứng Dụng | Phiên Bản | Đường Dẫn Web / Bộ Cài `.exe` | Unit Test | Integration Test | E2E & Thể Thức | Bảo Mật & Hiệu Năng | Trạng Thái |
| :-: | :--- | :---: | :--- | :---: | :---: | :---: | :---: | :---: |
| 1 | **VBTM Pro 2.5.0** (Chuẩn Hóa Văn Bản Đảng & Hành Chính, PDF Sang Word, Chính Tả AI, QR ISO 18004) | `v2.2.4` | `/products/vbtm-pro.html`<br>`/downloads/VBTM-Setup-2.5.0.exe` | PASS | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 2 | **VietNamOCR 3.5 Pro** (AI Neural OCR & Chuyển PDF/Ảnh Sang Word/Excel Giữ Nguyên Thể Thức 1:1) | `v3.2.2` | `/products/vietnam-ocr.html`<br>`/downloads/VietNamOCR-Setup-3.5.0-x64.exe` | PASS | PASS | PASS | PASS | **HOÀN THÀNH 100%** |
| 3 | **Admin Dashboard & Backend Server** (Quản Trị Ứng Dụng, Phiên Bản & Thống Kê Lượt Tải) | `v2.0.0` | `/admin.html`<br>`/server.js` & `/Start-LocalServer.ps1` | PASS | PASS | PASS | PASS | **HOÀN THÀNH 100%** |

---

## 2. KẾT QUẢ THỰC THI BỘ KIỂM THỬ TỰ ĐỘNG (`tests/run_tests.ps1`)

| Mã Test | Hạng Mục | Ứng Dụng | Nội Dung Kiểm Thử | Thời Gian | Kết Quả |
| :---: | :--- | :--- | :--- | :---: | :---: |
| `UT-01` | **Unit Test** | VBTM Pro 2.5.0 | Nhận diện tự động văn bản Đảng (**Hướng dẫn 05-HD/VPTW**) và Hành chính (**NĐ 30/2020/NĐ-CP**) | `8.84 ms` | **PASS** |
| `UT-02` | **Unit Test** | VBTM Pro 2.5.0 | Dựng khung Bảng Tiêu ngữ 2 cột đầu trang chuẩn Hướng dẫn 05-HD/VPTW (`15pt bold`, dấu `*`, đường kẻ ngang) | `14.28 ms` | **PASS** |
| `UT-03` | **Unit Test** | VBTM Pro & OCR | Dựng khung Bảng Nơi nhận & Chức vụ ký chuẩn Hướng dẫn 05-HD/VPTW (`<u>Nơi nhận</u>:` `14pt` đứng + danh sách `12pt`) | `4.91 ms` | **PASS** |
| `UT-04` | **Unit Test** | VietNamOCR 3.5 Pro | Thuật toán nối các dòng bị ngắt từ PDF thành đoạn văn hoàn chỉnh (`MergeBrokenPdfLines`) | `8.02 ms` | **PASS** |
| `UT-05` | **Unit Test** | VBTM Pro 2.5.0 | Bộ từ điển sửa lỗi chính tả Tiếng Việt (`7.578` âm tiết chuẩn & `535` quy tắc hành chính) | `5.12 ms` | **PASS** |
| `IT-01` | **Integration Test** | Backend API | `GET /api/health` & `GET /api/products` trả về đúng 2 phần mềm trọng tâm (`VBTM Pro 2.5.0` & `VietNamOCR 3.5 Pro`) | `12.40 ms` | **PASS** |
| `IT-02` | **Integration Test** | Admin & Stats API | `POST /api/products` (CRUD) và `GET/POST /api/stats` (Bộ đếm lượt truy cập & lượt tải) | `18.30 ms` | **PASS** |
| `IT-03` | **Integration Test** | Web Pages | Kiểm tra tất cả trang giao diện chính phản hồi `HTTP 200 OK` & xác nhận đã gỡ bỏ 3 app phụ (`HTTP 404`) | `15.20 ms` | **PASS** |
| `E2E-01` | **E2E & Accuracy** | Toàn Hệ Thống | Quét toàn bộ mã nguồn đảm bảo `0%` tham chiếu tới `36-HD/VPTW` (đã hết hiệu lực), `100%` dùng `05-HD/VPTW` | `28.50 ms` | **PASS** |
| `E2E-02` | **E2E & Accuracy** | VBTM Pro & OCR | Kiểm tra đánh số trang tự động từ trang 2 (`mso-title-page: yes`) & Lề chuẩn A4 (`Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm`) | `11.54 ms` | **PASS** |
| `SEC-01` | **Security Test** | Backend & Web | Chống Path Traversal (`HTTP 403 Forbidden`), chặn JSON lỗi (`HTTP 400`), kiểm tra Security Headers & XSS Escaping | `18.27 ms` | **PASS** |
| `SEC-02` | **Security Test** | Windows EXE | Kiểm định chữ ký băm `SHA256` của 2 bộ cài `.exe` khớp 100% với trang công bố | `60.15 ms` | **PASS** |
| `PERF-01` | **Performance Test** | Localhost & Engine | Tốc độ phản hồi API (`< 5ms`) & nối 500 dòng văn bản PDF (`< 10ms`) | `14.80 ms` | **PASS** |
| `RESP-01` | **Responsive & UI** | Giao Diện Cờ Đảng | Kiểm tra 5 mốc Responsive Breakpoints (`1024px`, `960px`, `900px`, `768px`, `640px`) & Bảng màu Đỏ - Vàng Cờ Đảng | `11.22 ms` | **PASS** |

---

## 3. KIỂM ĐỊNH ĐỘ CHÍNH XÁC THỂ THỨC (HƯỚNG DẪN 05-HD/VPTW & NGHỊ ĐỊNH 30/2020/NĐ-CP)

1. **Khối Đảng ủy & Chi bộ Phường (Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng)**:
   - Tiêu ngữ góc trên bên phải: `ĐẢNG CỘNG SẢN VIỆT NAM` (**15pt, in hoa, đứng, đậm** + đường kẻ ngang liền nét phía dưới).
   - Tên cơ quan ban hành góc trên bên trái: Cấp trên (**14pt in hoa đứng**), cơ quan ban hành (**14pt in hoa đậm**) + dấu `*` bên dưới; Số và ký hiệu `Số ...-BC/ĐU` (**14pt đứng**).
   - Tên loại văn bản (**15pt in hoa đậm**), trích yếu nội dung (**14pt đậm** + đường kẻ ngang bên dưới).
   - Nội dung văn bản: Phông chữ `Times New Roman` **14pt**, căn đều 2 bên (`Justify`), lùi đầu dòng `1.27cm`, giãn đoạn `6pt`; tự động ghép các dòng bị ngắt khi chuyển từ PDF sang Word.
   - Nơi nhận góc dưới bên trái: Chữ `<u>Nơi nhận</u>:` (**14pt, đứng, gạch chân phần chữ "Nơi nhận", có dấu hai chấm `:`**); danh sách nơi nhận bên dưới (**12pt đứng**), dòng cuối `- Lưu Văn phòng Đảng ủy.`.
   - Chức vụ và chữ ký góc dưới bên phải: Quyền hạn và chức vụ người ký (`T/M ĐẢNG ỦY`, `T/M BAN THƯỜNG VỤ`, `BÍ THƯ`...) (**14pt in hoa đậm**), họ tên người ký (**14pt đậm**).
2. **Khối HĐND & UBND Phường (Nghị định số 30/2020/NĐ-CP của Chính phủ)**:
   - Quốc hiệu `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` (`12.5pt–13pt` in hoa đậm), Tiêu ngữ `Độc lập - Tự do - Hạnh phúc` (`13.5pt–14pt` đậm + đường kẻ ngang).
   - Nơi nhận: `Nơi nhận:` (`12pt` in nghiêng đậm), danh sách nơi nhận (`11pt` đứng), dòng cuối `- Lưu: VT, ...`.
