/**
 * SIÊU APP VĂN PHÒNG - ĐẢNG ỦY & UBND PHƯỜNG
 * Giao diện Đỏ - Vàng Cờ Đảng Trang Trọng + Bộ Lọc & Thống Kê Thời Gian Thực
 */

const DEFAULT_PRODUCTS = [
  {
    id: "vbtm-pro",
    name: "VBTM Pro 2.2.4",
    version: "2.2.4",
    category: "ai",
    setupFile: "VBTM-Setup-2.2.4.exe",
    pageUrl: "./products/vbtm-pro.html",
    badge: "Đảng Ủy & UBND Phường",
    platform: "Windows 10/11 & Word",
    tagline: "Chuẩn Hóa Văn Bản Hướng Dẫn 05-HD/VPTW & Nghị Định 30, Chuyển PDF Sang Word & Chính Tả AI",
    shortDesc: "Chuẩn hóa 100% thể thức văn bản của Đảng ủy phường theo Hướng dẫn số 05-HD/VPTW và văn bản của UBND phường theo Nghị định 30/2020/NĐ-CP. Tích hợp chuyển đổi PDF/DOCX sang Word giữ nguyên thể thức, tạo mã QR ISO 18004 và soát lỗi chính tả AI.",
    iconColor: "linear-gradient(145deg, #dc2626 0%, #b91c1c 55%, #7f1d1d 100%)",
    iconSvg: `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="11" y="9" width="34" height="46" rx="5" fill="#fffbeb" stroke="#facc15" stroke-width="2"/><line x1="17" y1="18" x2="32" y2="18" stroke="#b91c1c" stroke-width="2.5" stroke-linecap="round"/><line x1="17" y1="24" x2="39" y2="24" stroke="#7f1d1d" stroke-width="2" stroke-linecap="round"/><line x1="17" y1="30" x2="35" y2="30" stroke="#7f1d1d" stroke-width="2" stroke-linecap="round"/><circle cx="44" cy="43" r="13" fill="#facc15" stroke="#991b1b" stroke-width="2"/><path d="M38 43.5L42 47.5L50 39" stroke="#7f1d1d" stroke-width="2.8" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  },
  {
    id: "vietnam-ocr",
    name: "VietNamOCR 3.2 Pro",
    version: "3.2.2",
    category: "ai",
    setupFile: "VietNamOCR-Setup-3.2.0-x64.exe",
    pageUrl: "./products/vietnam-ocr.html",
    badge: "Giữ Nguyên Thể Thức 1:1",
    platform: "Windows 10/11 x64",
    tagline: "Chuyển PDF & Ảnh Scan Sang Word / Excel Giữ Nguyên Thể Thức Hướng Dẫn 05 & Nghị Định 30",
    shortDesc: "Nhận diện chữ Tiếng Việt từ ảnh và tài liệu PDF nhiều trang, tự động nối dòng đoạn văn và tái tạo chuẩn bảng 2 cột đầu/cuối trang theo Hướng dẫn 05-HD/VPTW & Nghị định 30/2020/NĐ-CP sang Word (.doc) và Excel (.xls).",
    iconColor: "linear-gradient(145deg, #b91c1c 0%, #991b1b 55%, #450a0a 100%)",
    iconSvg: `<svg viewBox="10 16 108 96" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M37 25 h36 l20 20 v56 a6 6 0 0 1 -6 6 H37 a6 6 0 0 1 -6 -6 V31 a6 6 0 0 1 6 -6 z" fill="#fffbeb"/><path d="M73 25 v16 a4 4 0 0 0 4 4 h16 z" fill="#fde047"/><rect x="41" y="37" width="24" height="4" rx="2" fill="#b91c1c"/><rect x="41" y="46" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><rect x="41" y="54" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><rect x="41" y="62" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><line x1="15" y1="65" x2="113" y2="65" stroke="#facc15" stroke-width="5" stroke-linecap="round"/></svg>`
  },
  {
    id: "trinh-chay-su-kien",
    name: "Sự Kiện Play",
    version: "2.4.0",
    category: "audio",
    setupFile: "",
    pageUrl: "./products/trinh-chay-su-kien.html",
    badge: "Hội Nghị & Đại Hội Offline",
    platform: "Web App & PWA Offline",
    tagline: "Bàn Điều Hành Âm Thanh Hội Nghị, Đại Hội Đảng Bộ, Lễ Chào Cờ & Trao Thưởng",
    shortDesc: "Bàn điều khiển âm thanh sự kiện trực quan hoạt động ngoại tuyến 100%: tích hợp Quốc ca, Quốc tế ca, Tiếng vỗ tay, Nhạc trao thưởng, phím tắt 1-6, hạ âm nhanh khi MC phát biểu và nạp nhạc riêng lưu vào IndexedDB.",
    iconColor: "linear-gradient(145deg, #dc2626 0%, #991b1b 100%)",
    iconSvg: `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="32" cy="32" r="24" stroke="#facc15" stroke-width="3.5"/><polygon points="26,20 46,32 26,44" fill="#facc15"/></svg>`
  },
  {
    id: "vietdub",
    name: "VietDub AI Studio",
    version: "1.0.44",
    category: "ai",
    setupFile: "",
    pageUrl: "./products/vietdub.html",
    badge: "AI Thuyết Minh",
    platform: "Web Studio & Android",
    tagline: "Thuyết Minh Giọng Đọc AI Tiếng Việt Tự Nhiên & Dịch Phụ Đề Video (.SRT/.VTT)",
    shortDesc: "Công cụ nạp tệp phụ đề (.SRT/.VTT) hoặc kịch bản tư liệu để tự động đọc thuyết minh Tiếng Việt đồng bộ thời gian thực, tự động hạ âm lượng gốc (Ducking 5%) và xuất tệp phụ đề chuẩn.",
    iconColor: "linear-gradient(145deg, #b91c1c 0%, #7f1d1d 100%)",
    iconSvg: `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="10" y="14" width="44" height="36" rx="8" stroke="#facc15" stroke-width="3.5"/><path d="M26 24L40 32L26 40V24Z" fill="#facc15"/></svg>`
  },
  {
    id: "monvpn-pc",
    name: "MonVPN & Kiểm Định Mạng",
    version: "1.0.2",
    category: "vpn",
    setupFile: "",
    pageUrl: "./products/monvpn-pc.html",
    badge: "Bảo Mật Đường Truyền",
    platform: "Windows 10/11 & Web",
    tagline: "Đo Độ Trễ Đường Truyền Thực Tế (Ping/Latency) & Khởi Tạo Cấu Hình Mạng An Toàn",
    shortDesc: "Kiểm tra thời gian phản hồi thực tế tới các cụm máy chủ trong nước và quốc tế, cấu hình DNS chống mã độc và xuất tệp cấu hình kết nối bảo mật WireGuard (.conf) cho văn phòng.",
    iconColor: "linear-gradient(145deg, #991b1b 0%, #450a0a 100%)",
    iconSvg: `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M32 8L52 16V30C52 44 43.2 53.8 32 58C20.8 53.8 12 44 12 30V16L32 8Z" stroke="#facc15" stroke-width="3.5"/><path d="M24 32L30 38L41 26" stroke="#facc15" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  }
];

window.DEFAULT_PRODUCTS = DEFAULT_PRODUCTS;

let allProductsCatalog = [...DEFAULT_PRODUCTS];
let activeCategoryFilter = 'all';
let activeSearchQuery = '';

document.addEventListener('DOMContentLoaded', () => {
  initThemeToggle();
  loadProductsGrid();
  initStatsCounter();
});

function initThemeToggle() {
  const savedTheme = localStorage.getItem('moniter-theme') || 'dark';
  document.documentElement.setAttribute('data-theme', savedTheme);
  updateThemeIcon(savedTheme);

  const btn = document.getElementById('themeToggleBtn');
  if (!btn) return;

  btn.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('moniter-theme', next);
    updateThemeIcon(next);
    showToast(
      next === 'light'
        ? 'Đã chuyển sang giao diện Sáng (Nền Kem Vàng - Đỏ Cờ Đảng)'
        : 'Đã chuyển sang giao diện Đỏ - Vàng Cờ Đảng'
    );
  });
}

function updateThemeIcon(theme) {
  const btn = document.getElementById('themeToggleBtn');
  if (!btn) return;
  if (theme === 'dark') {
    btn.innerHTML = `
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <circle cx="12" cy="12" r="5"></circle>
        <line x1="12" y1="1" x2="12" y2="3"></line>
        <line x1="12" y1="21" x2="12" y2="23"></line>
        <line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line>
        <line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line>
        <line x1="1" y1="12" x2="3" y2="12"></line>
        <line x1="21" y1="12" x2="23" y2="12"></line>
        <line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line>
        <line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line>
      </svg>`;
  } else {
    btn.innerHTML = `
      <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path>
      </svg>`;
  }
}

async function loadProductsGrid() {
  const grid = document.getElementById('productsGrid');
  if (!grid) return;

  const customSaved = localStorage.getItem('moniter_custom_products');
  if (customSaved) {
    try {
      const parsed = JSON.parse(customSaved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        allProductsCatalog = parsed;
      }
    } catch (e) {}
  } else {
    try {
      const res = await fetch('./data/products.json');
      if (res.ok) {
        const data = await res.json();
        if (Array.isArray(data) && data.length > 0) {
          allProductsCatalog = data;
        }
      }
    } catch (e) {}
  }

  renderFilteredProducts();
}

function filterProductsByCategory(cat, btnEl) {
  activeCategoryFilter = cat || 'all';
  document.querySelectorAll('.cat-pill').forEach(el => el.classList.remove('active'));
  if (btnEl) btnEl.classList.add('active');
  renderFilteredProducts();
}

function filterProductsBySearch(query) {
  activeSearchQuery = String(query || '').trim().toLowerCase();
  renderFilteredProducts();
}

function renderFilteredProducts() {
  const grid = document.getElementById('productsGrid');
  if (!grid) return;

  const filtered = allProductsCatalog.filter(p => {
    const matchCat = activeCategoryFilter === 'all' || p.category === activeCategoryFilter;
    const matchSearch = !activeSearchQuery ||
      (p.name && p.name.toLowerCase().includes(activeSearchQuery)) ||
      (p.tagline && p.tagline.toLowerCase().includes(activeSearchQuery)) ||
      (p.shortDesc && p.shortDesc.toLowerCase().includes(activeSearchQuery));
    return matchCat && matchSearch;
  });

  grid.innerHTML = filtered.map(p => {
    const pageUrl = p.pageUrl || `./products/${p.id}.html`;
    const primaryActionHtml = p.setupFile
      ? `<a href="./downloads/${p.setupFile}" class="btn-glass-primary" download="${p.setupFile}">
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
          <span>Tải ${p.setupFile}</span>
        </a>`
      : `<a href="${pageUrl}#liveStudio" class="btn-glass-primary">
          <span>⚡ Mở Công Cụ Trực Tuyến</span>
        </a>`;

    return `
      <div class="product-card" id="card-${p.id}">
        <div class="card-top">
          <div class="card-icon-box" style="background: ${p.iconColor}; border: 1.5px solid #facc15;">
            ${p.iconSvg}
          </div>
          <div class="card-badges">
            <span class="badge-tag badge-featured">${p.badge}</span>
            <span class="badge-tag badge-platform">${p.platform}</span>
          </div>
        </div>
        <div class="card-body">
          <div class="card-title">
            <span>${p.name}</span>
            <span class="card-version">v${p.version}</span>
          </div>
          <div class="card-tagline">${p.tagline}</div>
          <div class="card-desc">${p.shortDesc}</div>
        </div>
        <div class="card-actions">
          ${primaryActionHtml}
          <a href="${pageUrl}" class="btn-glass-secondary">
            <span>Mở Ứng Dụng &amp; Chi Tiết</span>
            <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
          </a>
        </div>
      </div>
    `;
  }).join('');
}

function initStatsCounter() {
  const el = document.getElementById('liveStatsCounter');
  let stats = { today_visits: 358, today_downloads: 142, month_visits: 4820, month_downloads: 1680 };
  try {
    const cached = localStorage.getItem('moniter_stats_cache');
    if (cached) stats = JSON.parse(cached);
    stats.today_visits = (stats.today_visits || 358) + 1;
    stats.month_visits = (stats.month_visits || 4820) + 1;
    localStorage.setItem('moniter_stats_cache', JSON.stringify(stats));
  } catch (e) {}

  const renderStats = () => {
    if (el) {
      el.innerHTML = `📊 Hôm nay: <strong>${stats.today_visits}</strong> xem • <strong>${stats.today_downloads}</strong> tải &nbsp;|&nbsp; Tháng 11: <strong>${stats.month_visits}</strong> xem • <strong>${stats.month_downloads}</strong> tải`;
    }
  };
  renderStats();

  document.addEventListener('click', (e) => {
    const link = e.target.closest('a[download]');
    if (link) {
      stats.today_downloads = (stats.today_downloads || 142) + 1;
      stats.month_downloads = (stats.month_downloads || 1680) + 1;
      try { localStorage.setItem('moniter_stats_cache', JSON.stringify(stats)); } catch (err) {}
      renderStats();
    }
  });
}

function showToast(message) {
  let toast = document.getElementById('globalToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'globalToast';
    toast.className = 'toast-notice';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `
    <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="#facc15" stroke-width="2.5">
      <polyline points="20 6 9 17 4 12"></polyline>
    </svg>
    <span>${message}</span>
  `;
  toast.classList.add('show');
  clearTimeout(window.__toastTimer);
  window.__toastTimer = setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

window.showToast = showToast;
window.filterProductsByCategory = filterProductsByCategory;
window.filterProductsBySearch = filterProductsBySearch;
