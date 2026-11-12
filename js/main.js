/**
 * SOFTWARE HUB - Interactive Client Script
 * Apple iOS 18 Frosted Glassmorphism Portal
 */

const DEFAULT_PRODUCTS = [
  {
    id: "vbtm-pro",
    name: "VBTM Pro Ultimate",
    version: "4.0",
    setupFile: "VBTM-Setup-2.2.4.exe",
    badge: "Đảng Ủy & UBND Phường",
    platform: "Windows 10/11 & Word",
    tagline: "Chuẩn Hóa Thể Thức Văn Bản, Đọc PDF/Word Giữ Nguyên Thể Thức & Chính Tả AI",
    shortDesc: "Bản cao nhất 4.0 Ultimate: Chuẩn hóa thể thức văn bản của UBND phường (NĐ 30/2020/NĐ-CP) và Đảng ủy phường (Hướng dẫn số 05-HD/VPTW), chuyển đổi PDF/DOCX giữ nguyên 100% thể thức gốc, tạo mã QR và soát lỗi chính tả AI.",
    iconColor: "linear-gradient(145deg, #0ea5e9 0%, #2563eb 50%, #1e1b4b 100%)",
    iconSvg: `<svg viewBox="0 0 64 64" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="11" y="9" width="34" height="46" rx="5" fill="#ffffff"/><line x1="17" y1="18" x2="32" y2="18" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round"/><line x1="17" y1="24" x2="39" y2="24" stroke="#64748b" stroke-width="2" stroke-linecap="round"/><line x1="17" y1="30" x2="35" y2="30" stroke="#64748b" stroke-width="2" stroke-linecap="round"/><line x1="17" y1="36" x2="28" y2="36" stroke="#64748b" stroke-width="2" stroke-linecap="round"/><circle cx="44" cy="43" r="13" fill="#0ea5e9" stroke="#ffffff" stroke-width="2"/><path d="M38 43.5L42 47.5L50 39" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`
  },
  {
    id: "vietnam-ocr",
    name: "VietNamOCR 4.0 Pro",
    version: "4.0",
    setupFile: "VietNamOCR-Setup-3.2.0-x64.exe",
    badge: "Giữ Nguyên Thể Thức 1:1",
    platform: "Windows 10/11 x64",
    tagline: "Chuyển PDF Sang Word Giữ Nguyên Thể Thức 100% & AI Neural OCR Tiếng Việt",
    shortDesc: "Bản cao nhất 4.0 Pro Ultimate: Chuyển đổi PDF nhiều trang sang Word (.doc) giữ nguyên 100% thể thức và bố cục gốc không tự ý sửa đổi câu chữ; nhận diện chữ tiếng Việt từ ảnh & PDF scan xuất sang Word/Excel.",
    iconColor: "linear-gradient(145deg, #6443ff 0%, #4a29db 50%, #281195 100%)",
    iconSvg: `<svg viewBox="10 16 108 96" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M37 25 h36 l20 20 v56 a6 6 0 0 1 -6 6 H37 a6 6 0 0 1 -6 -6 V31 a6 6 0 0 1 6 -6 z" fill="#ffffff"/><path d="M73 25 v16 a4 4 0 0 0 4 4 h16 z" fill="#cbd5e1"/><rect x="41" y="37" width="24" height="4" rx="2" fill="#64748b"/><rect x="41" y="46" width="45" height="3.5" rx="1.75" fill="#94a3b8"/><rect x="41" y="54" width="45" height="3.5" rx="1.75" fill="#94a3b8"/><rect x="41" y="62" width="45" height="3.5" rx="1.75" fill="#94a3b8"/><rect x="41" y="70" width="45" height="3.5" rx="1.75" fill="#94a3b8"/><line x1="15" y1="65" x2="113" y2="65" stroke="#00f5d4" stroke-width="4" stroke-linecap="round"/></svg>`
  }
];

document.addEventListener('DOMContentLoaded', () => {
  initThemeToggle();
  loadProducts();
});

function initThemeToggle() {
  const themeBtn = document.getElementById('themeToggleBtn');
  const savedTheme = localStorage.getItem('moniter-theme') || 'dark';
  document.documentElement.setAttribute('data-theme', savedTheme);
  updateThemeIcon(savedTheme);

  if (themeBtn) {
    themeBtn.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const nextTheme = current === 'light' ? 'dark' : 'light';
      document.documentElement.setAttribute('data-theme', nextTheme);
      localStorage.setItem('moniter-theme', nextTheme);
      updateThemeIcon(nextTheme);
      showToast(`Đã chuyển sang giao diện ${nextTheme === 'light' ? 'Sáng (Titanium Tech)' : 'Tối (Deep Cyber Matrix)'}`);
    });
  }
}

function updateThemeIcon(theme) {
  const themeBtn = document.getElementById('themeToggleBtn');
  if (!themeBtn) return;
  if (theme === 'light') {
    themeBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>`;
    themeBtn.setAttribute('title', 'Chuyển sang giao diện tối');
  } else {
    themeBtn.innerHTML = `<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>`;
    themeBtn.setAttribute('title', 'Chuyển sang giao diện sáng');
  }
}

async function loadProducts() {
  const grid = document.getElementById('productsGrid');
  if (!grid) return;

  try {
    const res = await fetch('./data/products.json');
    if (res.ok) {
      const items = await res.json();
      renderProducts(items);
      return;
    }
  } catch (e) {}
  renderProducts(DEFAULT_PRODUCTS);
}

function renderProducts(items) {
  const grid = document.getElementById('productsGrid');
  if (!grid) return;

  grid.innerHTML = items.map(product => `
    <div class="product-card" id="card-${product.id}">
      <div class="card-top">
        <div class="card-icon-box" style="background: ${product.iconColor}">
          ${product.iconSvg}
        </div>
        <div class="card-badges">
          <span class="badge-tag badge-featured">${product.badge}</span>
          <span class="badge-tag badge-platform">${product.platform}</span>
        </div>
      </div>

      <div class="card-body">
        <div class="card-title">
          <span>${product.name}</span>
          <span class="card-version">v${product.version}</span>
        </div>
        <div class="card-tagline">${product.tagline}</div>
        <div class="card-desc">${product.shortDesc}</div>
      </div>

      <div class="card-actions">
        <a href="./downloads/${product.setupFile}" class="btn-glass-primary" download="${product.setupFile}">
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
          <span>Tải ${product.setupFile}</span>
        </a>
        <a href="./products/${product.id}.html" class="btn-glass-secondary">
          <span>Mở Ứng Dụng & Chi Tiết</span>
          <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
        </a>
      </div>
    </div>
  `).join('');
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
    <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="#38bdf8" stroke-width="2.4">
      <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path>
      <polyline points="22 4 12 14.01 9 11.01"></polyline>
    </svg>
    <span>${message}</span>
  `;

  toast.classList.add('show');
  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 3800);
}

window.showToast = showToast;
