/**
 * VBTM Ecosystem Website Synchronizer
 * Đọc nguồn cấu hình duy nhất data/product-config.json để đồng bộ phiên bản,
 * danh mục tích hợp Microsoft Office, AI đa nhà cung cấp và nhật ký phiên bản trên toàn Website Software Hub.
 */
(function () {
  'use strict';

  function esc(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  async function syncWebsiteFromConfig() {
    const isSubPage = window.location.pathname.includes('/products/');
    const prefix = isSubPage ? '../' : './';
    let cfg;
    try {
      const r = await fetch(prefix + 'data/product-config.json', { cache: 'no-cache' });
      if (!r.ok) return;
      cfg = await r.json();
    } catch (e) {
      return;
    }

    // 1. Đồng bộ nhãn phiên bản thống nhất
    document.querySelectorAll('[data-vbtm-version]').forEach(el => {
      el.textContent = 'v' + cfg.ecosystem.version;
    });
    document.querySelectorAll('[data-vbtm-release-date]').forEach(el => {
      el.textContent = cfg.ecosystem.releaseDate;
    });

    // 2. Đồng bộ khu vực VBTM PRO – MICROSOFT OFFICE INTEGRATION (nếu có container)
    const officeContainer = document.getElementById('vbtmUnifiedOfficeGrid');
    if (officeContainer && Array.isArray(cfg.officeIntegrations)) {
      const colors = { Word: '#2563eb', Excel: '#059669', PowerPoint: '#ea580c', Outlook: '#0284c7' };
      officeContainer.innerHTML = cfg.officeIntegrations.map(item => `
        <div class="feature-glass-card" style="border-left: 4px solid ${colors[item.host] || '#b91c1c'};">
          <div style="display:flex;justify-content:space-between;align-items:center;gap:6px;flex-wrap:wrap;">
            <h3 class="feature-title" style="margin:0;">${esc(item.title)}</h3>
            <span class="card-version" style="font-size:0.72rem;">${esc(item.host)}</span>
          </div>
          <p style="color: var(--accent-color); font-size: 0.80rem; margin: 4px 0 6px;">${esc(item.role)}</p>
          <ul style="margin: 0 0 10px 16px; padding: 0; color: var(--text-secondary); font-size: 0.82rem; line-height: 1.55;">
            ${(item.features || []).map(f => `<li>${esc(f)}</li>`).join('')}
          </ul>
          <div style="font-size: 0.76rem; color: var(--text-muted); margin-bottom: 8px;">
            Tương thích: ${esc(item.supportedOffice)} · Trạng thái: ${esc(item.status)}
          </div>
          <div style="display:flex;flex-wrap:wrap;gap:6px;">
            <a href="${prefix + esc(item.manifest)}" download="vbtm-${item.host.toLowerCase()}-manifest.xml" class="btn-tool" style="text-decoration:none;">⬇ Manifest ${esc(item.host)}</a>
            <a href="${prefix + esc(item.taskpane)}" target="_blank" class="btn-tool" style="text-decoration:none;">👁 Xem Task Pane</a>
          </div>
        </div>
      `).join('');
    }

    // 3. Đồng bộ Nhật ký phiên bản & Trung tâm AI đa nhà cung cấp (nếu có container trên vbtm-pro.html)
    const changelogEl = document.getElementById('vbtmUnifiedChangelog');
    if (changelogEl && Array.isArray(cfg.changelog)) {
      changelogEl.innerHTML = cfg.changelog.map(c => `
        <div style="margin-bottom: 10px;">
          <strong style="color: #fef08a;">Phiên bản ${esc(c.version)}</strong> <span style="color: var(--text-muted); font-size: 0.82rem;">(${esc(c.date)})</span>
          <ul style="margin: 4px 0 0 18px; color: var(--text-secondary); font-size: 0.86rem; line-height: 1.6;">
            ${(c.highlights || []).map(h => `<li>${esc(h)}</li>`).join('')}
          </ul>
        </div>
      `).join('');
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', syncWebsiteFromConfig);
  } else {
    syncWebsiteFromConfig();
  }
})();
