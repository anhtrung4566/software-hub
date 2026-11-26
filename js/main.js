/**
 * SIÊU APP VĂN PHÒNG - ĐẢNG ỦY & UBND PHƯỜNG
 * Giao diện Đỏ - Vàng Cờ Đảng Trang Trọng + Bộ Lọc & Thống Kê Thời Gian Thực
 */

const DEFAULT_PRODUCTS = [
  {
    id: "vbtm-pro",
    name: "VBTM Pro 3.6.0",
    version: "3.6.0",
    category: "ai",
    setupFile: "VBTM-Setup-3.6.0.zip",
    pageUrl: "./products/vbtm-pro.html",
    badge: "Đảng Ủy & UBND Phường",
    platform: "Windows 10/11 & Word",
    tagline: "Chuẩn Hóa Văn Bản 05-HD/VPTW & NĐ 30, Kho 84 Mẫu Chuẩn & Chính Tả AI Offline (ProtonX / Ollama)",
    shortDesc: "Chuẩn hóa 100% thể thức văn bản Đảng ủy phường theo Hướng dẫn 05-HD/VPTW, UBND phường theo Nghị định 30/2020/NĐ-CP và CAND. Tích hợp kiểm tra chính tả AI 3 tầng Offline (CTranslate2 ProtonX + 147.000 từ điển + Ollama), kho 84 mẫu văn bản chuẩn và tạo mã QR ISO 18004.",
    iconColor: "linear-gradient(145deg, #dc2626 0%, #b91c1c 55%, #7f1d1d 100%)",
    iconSvg: `<img src="./desktop-src/VBTM-Pro/VBTM-Pro-logo.png" alt="VBTM Pro 3.6.0" style="width:100%;height:100%;object-fit:cover;border-radius:14px;display:block;">`
  },
  {
    id: "vietnam-ocr",
    name: "VietNamOCR 3.5 Pro",
    version: "3.5.0",
    category: "ai",
    setupFile: "VietNamOCR-Setup-3.5.0-x64.exe",
    pageUrl: "./products/vietnam-ocr.html",
    badge: "Giữ Nguyên Thể Thức 1:1",
    platform: "Windows 10/11 x64",
    tagline: "Chuyển PDF & Ảnh Scan Sang Word / Excel Giữ Nguyên Thể Thức Hướng Dẫn 05 & Nghị Định 30",
    shortDesc: "Nhận diện chữ Tiếng Việt từ ảnh và tài liệu PDF nhiều trang, tự động nối dòng đoạn văn và tái tạo chuẩn bảng 2 cột đầu/cuối trang theo Hướng dẫn 05-HD/VPTW & Nghị định 30/2020/NĐ-CP sang Word (.doc) và Excel (.xls).",
    iconColor: "linear-gradient(145deg, #b91c1c 0%, #991b1b 55%, #450a0a 100%)",
    iconSvg: `<svg viewBox="10 16 108 96" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M37 25 h36 l20 20 v56 a6 6 0 0 1 -6 6 H37 a6 6 0 0 1 -6 -6 V31 a6 6 0 0 1 6 -6 z" fill="#fffbeb"/><path d="M73 25 v16 a4 4 0 0 0 4 4 h16 z" fill="#fde047"/><rect x="41" y="37" width="24" height="4" rx="2" fill="#b91c1c"/><rect x="41" y="46" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><rect x="41" y="54" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><rect x="41" y="62" width="45" height="3.5" rx="1.75" fill="#7f1d1d"/><line x1="15" y1="65" x2="113" y2="65" stroke="#facc15" stroke-width="5" stroke-linecap="round"/></svg>`
  }
];

const REMOVED_APP_IDS = new Set(['trinh-chay-su-kien', 'vietdub', 'monvpn-pc']);

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
      const parsed = JSON.parse(customSaved).filter(p => !REMOVED_APP_IDS.has(p.id));
      localStorage.setItem('moniter_custom_products', JSON.stringify(parsed));
      if (Array.isArray(parsed) && parsed.length > 0) {
        allProductsCatalog = parsed;
      }
    } catch (e) {}
  } else {
    try {
      const res = await fetch('./data/products.json');
      if (res.ok) {
        const data = (await res.json()).filter(p => !REMOVED_APP_IDS.has(p.id));
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
