# HƯỚNG DẪN CHẠY LOCALHOST & TRIỂN KHAI LÊN HOSTING / VPS

Tài liệu hướng dẫn vận hành hệ thống **Phần Mềm Văn Phòng · Đảng Ủy & UBND Phường** (`VBTM Pro 2.2.4` & `VietNamOCR 3.2 Pro`) trên máy tính cá nhân (`localhost`), Shared Hosting, GitHub Pages hoặc máy chủ VPS (Ubuntu/Debian/Windows Server).

---

## 1. Chạy Hệ Thống Trên Localhost (Máy Tính Windows)

### Cách 1: Chạy trực tiếp bằng máy chủ .NET / PowerShell tích hợp sẵn (Không cần cài thêm phần mềm)
1. Nhấp đúp chuột vào tệp `chay_server_local.bat` trong thư mục dự án, hoặc mở PowerShell chạy lệnh:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Start-LocalServer.ps1 -Port 8080
   ```
2. Mở trình duyệt truy cập:
   - **Trang chủ Kho Phần Mềm**: `http://localhost:8080/`
   - **VBTM Pro 2.2.4 Studio**: `http://localhost:8080/products/vbtm-pro.html`
   - **VietNamOCR 3.2 Pro Studio**: `http://localhost:8080/products/vietnam-ocr.html`
   - **Trang Quản Trị (Admin Dashboard)**: `http://localhost:8080/admin.html`

### Cách 2: Chạy bằng Node.js Backend (`server.js`)
```bash
node server.js
```
Máy chủ lắng nghe tại cổng `8080` (hoặc biến môi trường `PORT`).

### Cách 3: Chạy bộ kiểm thử tự động (Unit / Integration / E2E / Security / Performance)
```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run_tests.ps1
```

---

## 2. Triển Khai Tự Động Lên GitHub Pages

Dự án đã tích hợp sẵn kịch bản đồng bộ tự động qua GitHub REST API:
1. Nhấp đúp chuột vào tệp `day_len_github.bat`.
2. Hệ thống tự động tải toàn bộ mã nguồn, giao diện và bộ cài `.exe` / `.zip` lên nhánh `main` và xuất bản tại:
   - `https://anhtrung4566.github.io/software-hub/`

---

## 3. Triển Khai Lên Máy Chủ VPS (Ubuntu / Debian + Nginx + PM2)

### Bước 1: Cài đặt Node.js, PM2 và Nginx trên VPS
```bash
sudo apt update && sudo apt install -y nodejs npm nginx git
sudo npm install -g pm2
```

### Bước 2: Tải mã nguồn về VPS và khởi chạy Backend với PM2
```bash
cd /var/www
sudo git clone https://github.com/anhtrung4566/software-hub.git
cd software-hub
PORT=8080 pm2 start server.js --name "sieu-app-van-phong"
pm2 save
pm2 startup
```

### Bước 3: Cấu hình Nginx Reverse Proxy & SSL (Let's Encrypt)
Tạo tệp `/etc/nginx/sites-available/sieu-app-van-phong`:
```nginx
server {
    listen 80;
    server_name vanphong.tenmien.vn;

    client_max_body_size 20M;

    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```
Kích hoạt cấu hình và cài đặt chứng chỉ HTTPS miễn phí:
```bash
sudo ln -s /etc/nginx/sites-available/sieu-app-van-phong /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d vanphong.tenmien.vn
```
