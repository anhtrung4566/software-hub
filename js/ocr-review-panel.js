/**
 * VietNamOCR Pro — Giao diện KIỂM TRA OCR ĐỐI CHIẾU ẢNH GỐC (ocr-review-panel.js)
 *  - Trái: ảnh trang đã OCR, tô khung vùng độ tin cậy thấp / nghi dính chữ / đã tách từ dựa ảnh. Bấm vào ảnh để chọn dòng.
 *  - Phải: từng dòng OCR (kèm ảnh cắt đúng vùng), sửa trực tiếp, chấp nhận/từ chối đề xuất, khôi phục nội dung OCR gốc.
 *  - Không tự sửa hàng loạt nội dung không chắc chắn: đề xuất loại A (OCR lại vùng ảnh) chỉ áp dụng khi người dùng đồng ý.
 */
(function (global) {
  'use strict';

  const pages = new Map(); // pageNo -> { pageNo, canvas, W, H, items[], layout }
  let curPage = 1;
  let selectedId = null;
  let onlyFlagged = true;
  let modal = null;
  let seq = 0;

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
  function nfc(s) { return String(s == null ? '' : s).normalize('NFC'); }

  function reset() { pages.clear(); curPage = 1; selectedId = null; }

  function registerPage(pageNo, res) {
    if (!res || !res.procCanvas) return;
    const layout = res.layout || null;
    const lines = (layout && layout.lines) ? layout.lines : (res.lines || []);
    const suggByLine = new Map();
    if (layout && Array.isArray(layout.suggestions)) {
      layout.suggestions.forEach(s => { if (s.lineRef) { if (!suggByLine.has(s.lineRef)) suggByLine.set(s.lineRef, []); suggByLine.get(s.lineRef).push(s); } });
    }
    const items = [];
    lines.forEach(ln => {
      const text = nfc(ln.text || '').trim();
      if (!text || !ln.bbox) return;
      const words = ln.words || [];
      const flags = new Set();
      if ((ln.minWordConf != null ? ln.minWordConf : 100) < 70) flags.add('lowconf');
      if (ln.spacingRepaired) flags.add('spacing');
      words.forEach(w => (w.flags || []).forEach(f => flags.add(f)));
      const confs = words.map(w => Number(w.confidence)).filter(v => !isNaN(v));
      items.push({
        id: 'L' + (++seq),
        bbox: ln.bbox,
        raw: nfc(ln.rawText != null ? ln.rawText : ln.text).trim(),
        applied: text, // nội dung hiện đang nằm trong kết quả
        cur: text,
        conf: confs.length ? Math.round(confs.reduce((a, b) => a + b, 0) / confs.length) : null,
        flags,
        cell: ln.cell || null,
        sugg: (suggByLine.get(ln) || []).map(s => ({ kind: s.kind, original: s.original, proposed: s.proposed, evidence: s.evidence, confGain: s.confGain, state: 'pending' })),
        words
      });
    });
    pages.set(pageNo, { pageNo, canvas: res.procCanvas, W: res.procCanvas.width, H: res.procCanvas.height, items, layout, repairs: layout ? layout.repairs : [] });
  }

  // ------------------------------------------------------------------ giao diện
  function ensureModal() {
    if (modal) return modal;
    const st = document.createElement('style');
    st.textContent = `
      #ocrReviewModal{position:fixed;inset:0;z-index:99999;background:rgba(2,6,23,.78);display:none;align-items:stretch;justify-content:center;padding:14px}
      #ocrReviewModal .rv-box{background:#0f172a;color:#e2e8f0;border:1.5px solid #facc15;border-radius:14px;width:min(1500px,100%);display:flex;flex-direction:column;overflow:hidden}
      .rv-head{display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:10px 14px;border-bottom:1px solid rgba(250,204,21,.35);background:#111c33}
      .rv-head h3{margin:0;font-size:1.02rem;color:#fef08a;flex:1;min-width:260px}
      .rv-btn{background:#1e293b;color:#e2e8f0;border:1px solid #475569;border-radius:8px;padding:6px 12px;font-size:.84rem;font-weight:600;cursor:pointer}
      .rv-btn.pri{background:#15803d;border-color:#22c55e;color:#fff}.rv-btn.warn{background:#b45309;border-color:#f59e0b;color:#fff}
      .rv-body{display:grid;grid-template-columns:minmax(320px,46%) 1fr;gap:0;flex:1;min-height:0}
      .rv-left{overflow:auto;background:#020617;padding:8px;border-right:1px solid #334155}
      .rv-left canvas{width:100%;height:auto;display:block;cursor:crosshair;background:#fff}
      .rv-right{overflow:auto;padding:8px 10px}
      .rv-legend{font-size:.76rem;color:#94a3b8;padding:4px 14px;border-bottom:1px solid #1e293b}
      .rv-row{border:1px solid #334155;border-radius:10px;padding:8px;margin-bottom:8px;background:#0b1224}
      .rv-row.sel{border-color:#facc15;box-shadow:0 0 0 2px rgba(250,204,21,.35)}
      .rv-meta{font-size:.76rem;color:#94a3b8;margin-bottom:5px;display:flex;flex-wrap:wrap;gap:5px;align-items:center}
      .rv-tag{border-radius:999px;padding:1px 8px;font-size:.70rem;font-weight:700}
      .rv-tag.lowconf{background:#7f1d1d;color:#fecaca}.rv-tag.spacing{background:#581c87;color:#e9d5ff}.rv-tag.stuck{background:#9a3412;color:#fed7aa}
      .rv-tag.date,.rv-tag.confusable{background:#713f12;color:#fde68a}.rv-tag.cell{background:#1e3a8a;color:#bfdbfe}
      .rv-img{max-width:100%;background:#fff;border-radius:6px;display:block;margin:4px 0}
      .rv-input{width:100%;box-sizing:border-box;background:#fff;color:#0f172a;border:1px solid #94a3b8;border-radius:6px;padding:6px 8px;font:15px 'Times New Roman',serif}
      .rv-orig{font-size:.76rem;color:#fbbf24;margin-top:4px;word-break:break-word}
      .rv-sugg{font-size:.80rem;margin-top:6px;padding:6px 8px;background:#052e16;border:1px solid #166534;border-radius:8px}
      .rv-sugg.done{opacity:.55}.rv-sugg b{color:#86efac}.rv-sugg .ev{color:#94a3b8;font-size:.72rem;display:block;margin:2px 0 4px}
    `;
    document.head.appendChild(st);
    modal = document.createElement('div');
    modal.id = 'ocrReviewModal';
    modal.innerHTML = `
      <div class="rv-box">
        <div class="rv-head">
          <h3>🔎 Kiểm tra kết quả OCR — đối chiếu ảnh gốc (trái) và văn bản OCR (phải)</h3>
          <select id="rvPageSel" class="rv-btn"></select>
          <label style="font-size:.82rem"><input type="checkbox" id="rvOnlyFlag" checked> Chỉ dòng cần kiểm tra</label>
          <button class="rv-btn warn" id="rvAcceptSure" title="Chỉ chấp nhận đề xuất loại A có độ tin cậy tăng ≥ 15%">✓ Chấp nhận đề xuất chắc chắn</button>
          <button class="rv-btn pri" id="rvApply">⬇ Áp dụng vào kết quả</button>
          <button class="rv-btn" id="rvClose">✕ Đóng</button>
        </div>
        <div class="rv-legend">Khung đỏ/cam: độ tin cậy thấp · Khung tím: đã tách từ dựa khoảng trắng thật trong ảnh · Khung cam đậm: nghi dính chữ (cần bạn xác nhận) · Khung xanh dương: dòng đang chọn.
          Đề xuất <b>A</b> = OCR lại đúng vùng ảnh · <b>B</b> = khôi phục khoảng trắng theo tọa độ/ảnh · Sửa chính tả (<b>C</b>) chỉ là hỗ trợ.</div>
        <div class="rv-body"><div class="rv-left"><canvas id="rvCanvas"></canvas></div><div class="rv-right" id="rvList"></div></div>
      </div>`;
    document.body.appendChild(modal);
    modal.querySelector('#rvClose').onclick = () => { modal.style.display = 'none'; };
    modal.querySelector('#rvApply').onclick = applyToResult;
    modal.querySelector('#rvAcceptSure').onclick = acceptSure;
    modal.querySelector('#rvOnlyFlag').onchange = (e) => { onlyFlagged = e.target.checked; renderList(); };
    modal.querySelector('#rvPageSel').onchange = (e) => { curPage = Number(e.target.value); selectedId = null; renderAll(); };
    modal.querySelector('#rvCanvas').addEventListener('click', onCanvasClick);
    return modal;
  }

  function pageObj() { return pages.get(curPage); }

  function drawCanvas() {
    const pg = pageObj();
    const cv = modal.querySelector('#rvCanvas');
    if (!pg) { cv.width = 10; cv.height = 10; return; }
    cv.width = pg.W; cv.height = pg.H;
    const c = cv.getContext('2d');
    c.drawImage(pg.canvas, 0, 0);
    c.lineWidth = Math.max(2, pg.W / 700);
    pg.items.forEach(it => {
      const f = it.flags;
      let color = null, fill = null;
      if (f.has('stuck')) { color = '#ea580c'; fill = 'rgba(234,88,12,.20)'; }
      else if (f.has('spacing')) { color = '#9333ea'; fill = 'rgba(147,51,234,.16)'; }
      else if (f.has('lowconf')) { color = it.conf != null && it.conf < 50 ? '#dc2626' : '#d97706'; fill = 'rgba(239,68,68,.16)'; }
      else if (f.has('date') || f.has('confusable')) { color = '#ca8a04'; fill = 'rgba(234,179,8,.18)'; }
      if (it.id === selectedId) { color = '#2563eb'; fill = 'rgba(37,99,235,.20)'; }
      if (!color) return;
      const b = it.bbox;
      c.fillStyle = fill; c.strokeStyle = color;
      c.fillRect(b.x0, b.y0, b.x1 - b.x0, b.y1 - b.y0);
      c.strokeRect(b.x0, b.y0, b.x1 - b.x0, b.y1 - b.y0);
    });
  }

  function cropUrl(pg, b) {
    const pad = 6;
    const x0 = Math.max(0, b.x0 - pad), y0 = Math.max(0, b.y0 - pad);
    const w = Math.min(pg.W - x0, b.x1 - b.x0 + pad * 2), h = Math.min(pg.H - y0, b.y1 - b.y0 + pad * 2);
    const c = document.createElement('canvas');
    c.width = Math.max(1, w); c.height = Math.max(1, h);
    c.getContext('2d').drawImage(pg.canvas, x0, y0, w, h, 0, 0, w, h);
    return c.toDataURL('image/png');
  }

  const TAGS = { lowconf: 'Độ tin cậy thấp', spacing: 'Đã tách từ theo ảnh', stuck: 'Nghi dính chữ — cần xác nhận', date: 'Ngày không hợp lệ', confusable: 'Dễ nhầm O/0 I/l/1 S/5' };

  function renderList() {
    const pg = pageObj();
    const host = modal.querySelector('#rvList');
    if (!pg) { host.innerHTML = '<p style="padding:20px;color:#94a3b8">Chưa có dữ liệu OCR để đối chiếu. Hãy tải ảnh/PDF scan và bấm “Bắt đầu OCR” (trang PDF có lớp chữ gốc không cần OCR).</p>'; return; }
    let items = pg.items;
    if (onlyFlagged) items = items.filter(it => it.flags.size || it.sugg.length || it.cur !== it.raw);
    if (!items.length) { host.innerHTML = '<p style="padding:20px;color:#86efac">✓ Không có dòng nào cần kiểm tra trên trang này.</p>'; return; }
    host.innerHTML = items.map((it, idx) => {
      const tags = Array.from(it.flags).map(f => `<span class="rv-tag ${f}">${TAGS[f]}</span>`).join('') + (it.cell ? `<span class="rv-tag cell">Ô hàng ${it.cell.r + 1} · cột ${it.cell.c + 1}</span>` : '');
      const sg = it.sugg.map((s, si) => `<div class="rv-sugg ${s.state !== 'pending' ? 'done' : ''}">
          <b>Đề xuất ${s.kind}</b>: “${esc(s.proposed)}” <span style="color:#86efac">(+${s.confGain}% tin cậy)</span>
          <span class="ev">${esc(s.evidence)}</span>
          ${s.state === 'pending' ? `<button class="rv-btn pri" data-act="acc" data-id="${it.id}" data-si="${si}">Chấp nhận</button> <button class="rv-btn" data-act="rej" data-id="${it.id}" data-si="${si}">Từ chối</button>` : `<i>${s.state === 'accepted' ? 'Đã chấp nhận' : 'Đã từ chối'}</i>`}
        </div>`).join('');
      return `<div class="rv-row ${it.id === selectedId ? 'sel' : ''}" data-id="${it.id}">
        <div class="rv-meta"><b>#${idx + 1}</b> · tin cậy ${it.conf == null ? '—' : it.conf + '%'} ${tags}</div>
        <img class="rv-img" src="${cropUrl(pg, it.bbox)}" alt="vùng ảnh gốc">
        <input class="rv-input" data-act="edit" data-id="${it.id}" value="${esc(it.cur)}">
        ${it.cur !== it.raw ? `<div class="rv-orig">OCR gốc (chưa sửa khoảng trắng/ký tự): “${esc(it.raw)}” <button class="rv-btn" data-act="restore" data-id="${it.id}">↺ Khôi phục OCR gốc</button></div>` : ''}
        ${sg}
      </div>`;
    }).join('');
    host.querySelectorAll('.rv-row').forEach(r => r.addEventListener('click', () => select(r.dataset.id, false)));
    host.querySelectorAll('[data-act]').forEach(el => {
      const act = el.dataset.act;
      if (act === 'edit') el.addEventListener('input', () => { const it = findItem(el.dataset.id); if (it) it.cur = nfc(el.value); });
      else el.addEventListener('click', (ev) => { ev.stopPropagation(); onAction(act, el.dataset.id, Number(el.dataset.si)); });
    });
  }

  function findItem(id) { const pg = pageObj(); return pg ? pg.items.find(i => i.id === id) : null; }

  function onAction(act, id, si) {
    const it = findItem(id);
    if (!it) return;
    if (act === 'restore') it.cur = it.raw;
    else if (act === 'acc') { const s = it.sugg[si]; if (s) { s.state = 'accepted'; it.cur = s.proposed; } }
    else if (act === 'rej') { const s = it.sugg[si]; if (s) s.state = 'rejected'; }
    renderList();
  }

  function select(id, scroll) {
    selectedId = id;
    drawCanvas();
    modal.querySelectorAll('.rv-row').forEach(r => r.classList.toggle('sel', r.dataset.id === id));
    if (scroll) { const row = modal.querySelector(`.rv-row[data-id="${id}"]`); if (row) row.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
  }

  function onCanvasClick(ev) {
    const pg = pageObj();
    if (!pg) return;
    const cv = ev.currentTarget, r = cv.getBoundingClientRect();
    const x = (ev.clientX - r.left) * (cv.width / r.width), y = (ev.clientY - r.top) * (cv.height / r.height);
    let hit = null;
    pg.items.forEach(it => { const b = it.bbox; if (x >= b.x0 - 4 && x <= b.x1 + 4 && y >= b.y0 - 4 && y <= b.y1 + 4) hit = hit || it; });
    if (!hit) return;
    if (onlyFlagged && !(hit.flags.size || hit.sugg.length || hit.cur !== hit.raw)) {
      onlyFlagged = false; modal.querySelector('#rvOnlyFlag').checked = false; renderList();
    }
    select(hit.id, true);
  }

  function acceptSure() {
    let n = 0;
    pages.forEach(pg => pg.items.forEach(it => it.sugg.forEach(s => {
      if (s.state === 'pending' && s.kind === 'A' && s.confGain >= 15) { s.state = 'accepted'; it.cur = s.proposed; n++; }
    })));
    renderList();
    if (global.showToast) global.showToast(`Đã chấp nhận ${n} đề xuất có độ tin cậy tăng ≥ 15% (các đề xuất khác vẫn chờ bạn quyết định).`);
  }

  // thay đúng đoạn văn bản cũ bằng đoạn mới trong bản xem trước Word (giữ nguyên cấu trúc bảng/ô)
  function replaceInPaper(oldText, newText) {
    const paper = document.getElementById('ocrFormattedPaper');
    if (!paper || !oldText) return false;
    const walker = document.createTreeWalker(paper, NodeFilter.SHOW_TEXT);
    let n;
    while ((n = walker.nextNode())) {
      const v = n.nodeValue.normalize('NFC');
      const i = v.indexOf(oldText);
      if (i >= 0) { n.nodeValue = v.slice(0, i) + newText + v.slice(i + oldText.length); return true; }
    }
    return false;
  }

  function applyToResult() {
    let changed = 0, notFound = 0;
    const ta = document.getElementById('ocrResultText');
    pages.forEach(pg => pg.items.forEach(it => {
      if (it.cur === it.applied) return;
      const okPaper = replaceInPaper(it.applied, it.cur);
      if (ta) {
        const v = ta.value.normalize('NFC');
        const i = v.indexOf(it.applied);
        if (i >= 0) ta.value = v.slice(0, i) + it.cur + v.slice(i + it.applied.length);
      }
      if (!okPaper) notFound++;
      it.applied = it.cur;
      changed++;
    }));
    if (global.showToast) global.showToast(changed ? `✓ Đã áp dụng ${changed} chỉnh sửa vào kết quả OCR${notFound ? ` (${notFound} dòng chỉ cập nhật ở văn bản thuần)` : ''}. Bấm “Xuất DOCX” để tải về.` : 'Không có chỉnh sửa nào cần áp dụng.');
  }

  function renderAll() {
    const sel = modal.querySelector('#rvPageSel');
    const keys = Array.from(pages.keys()).sort((a, b) => a - b);
    sel.innerHTML = keys.map(k => `<option value="${k}" ${k === curPage ? 'selected' : ''}>Trang ${k}</option>`).join('') || '<option>—</option>';
    drawCanvas();
    renderList();
  }

  function open() {
    ensureModal();
    const keys = Array.from(pages.keys()).sort((a, b) => a - b);
    if (keys.length && !pages.has(curPage)) curPage = keys[0];
    modal.style.display = 'flex';
    renderAll();
  }

  function getModel() {
    const out = [];
    pages.forEach(pg => out.push({
      page: pg.pageNo, width: pg.W, height: pg.H,
      regions: pg.layout ? pg.layout.regions : [], columns: pg.layout ? pg.layout.columns : 1,
      lines: pg.items.map(it => ({ bbox: it.bbox, ocrRaw: it.raw, text: it.cur, confidence: it.conf, flags: Array.from(it.flags), cell: it.cell, suggestions: it.sugg.map(s => ({ kind: s.kind, proposed: s.proposed, state: s.state })) })),
      spacingRepairs: pg.repairs || []
    }));
    return out;
  }

  global.VietNamOcrReview = { reset, registerPage, open, getModel };
})(window);
