/**
 * VietNamOCR 3.5 Pro — Bộ AI Kiểm Tra & Hiệu Chỉnh OCR Có Bảo Vệ Thực Thể (js/ocr-ai-proofread.js)
 *
 * Tính năng:
 * - Phân tích văn bản OCR gốc theo 2 lớp:
 *   1. Lớp Offline tức thì: Phát hiện lỗi dấu/vần/cụm từ OCR tiếng Việt phổ biến.
 *   2. Lớp AI Local (Ollama qwen2.5:7b / qwen2.5:3b tại http://127.0.0.1:11434): Phát hiện lỗi ngữ cảnh & đoạn thiếu/sai.
 * - Bảo vệ tuyệt đối thực thể quan trọng (Entity Protection Guard):
 *   Họ tên, Địa chỉ, CCCD (9/12 số), Mã số thuế, Số hợp đồng, Số văn bản, Ngày tháng năm, Số tiền, Tỷ lệ %, Điều/Khoản.
 * - Giao diện đối chiếu trực quan: Xem bản OCR gốc vs đề xuất sửa, Chấp nhận / Từ chối từng đề xuất riêng lẻ!
 */
(function (global) {
  'use strict';

  const OCR_OFFLINE_PAIRS = [
    ['CỘNG HỎA', 'CỘNG HÒA', 'Lỗi dấu OCR Quốc hiệu'],
    ['CÔNG HÒA', 'CỘNG HÒA', 'Lỗi dấu OCR Quốc hiệu'],
    ['XÃ HÔI', 'XÃ HỘI', 'Thiếu dấu nặng OCR'],
    ['ĐÔC LẬP', 'ĐỘC LẬP', 'Thiếu dấu nặng OCR'],
    ['Hạnh phục', 'Hạnh phúc', 'Sai dấu sắc/nặng OCR'],
    ['ĐẢNG CÔNG SẢN', 'ĐẢNG CỘNG SẢN', 'Thiếu dấu nặng OCR'],
    ['ỦY BẠN', 'ỦY BAN', 'Dư dấu nặng OCR'],
    ['QUYÊT ĐỊNH', 'QUYẾT ĐỊNH', 'Thiếu dấu sắc OCR'],
    ['NGHỊ QUYÊT', 'NGHỊ QUYẾT', 'Thiếu dấu sắc OCR'],
    ['KẾ HOACH', 'KẾ HOẠCH', 'Thiếu dấu nặng OCR'],
    ['BÁO CAO', 'BÁO CÁO', 'Thiếu dấu sắc OCR'],
    ['Nơi nhạn', 'Nơi nhận', 'Sai dấu OCR'],
    ['Kính gừi', 'Kính gửi', 'Sai dấu hỏi/huyền OCR'],
    ['Căn cự', 'Căn cứ', 'Sai dấu sắc/nặng OCR'],
    ['hướng dân', 'hướng dẫn', 'Thiếu dấu ngã OCR'],
    ['bổ xung', 'bổ sung', 'Lỗi chính tả s/x'],
    ['sát nhập', 'sáp nhập', 'Lỗi chính tả p/t'],
    ['sử lý', 'xử lý', 'Lỗi chính tả s/x'],
    ['đông chí', 'đồng chí', 'Thiếu dấu huyền OCR'],
    ['tô chức', 'tổ chức', 'Thiếu dấu hỏi OCR'],
    ['thưc hiện', 'thực hiện', 'Thiếu dấu nặng OCR'],
    ['thực hiên', 'thực hiện', 'Thiếu dấu nặng OCR'],
    ['nhiệm vu', 'nhiệm vụ', 'Thiếu dấu nặng OCR'],
    ['chỉ đao', 'chỉ đạo', 'Thiếu dấu nặng OCR'],
    ['lãnh đao', 'lãnh đạo', 'Thiếu dấu nặng OCR'],
    ['đưoc', 'được', 'Lỗi dấu OCR'],
    ['ngời', 'người', 'Thiếu chữ ư OCR'],
    ['cán bô', 'cán bộ', 'Thiếu dấu nặng OCR'],
    ['Đảng phải', 'Đảng phái', 'Lỗi ngữ cảnh / thanh điệu'],
    ['năm giữ', 'nắm giữ', 'Lỗi ngữ cảnh / thanh điệu'],
    ['chất lương', 'chất lượng', 'Thiếu dấu nặng OCR'],
    ['hiêu quả', 'hiệu quả', 'Thiếu dấu nặng OCR'],
    ['đề nghi', 'đề nghị', 'Thiếu dấu nặng OCR'],
    ['trách nhiêm', 'trách nhiệm', 'Thiếu dấu nặng OCR'],
    ['giải quyêt', 'giải quyết', 'Thiếu dấu sắc OCR'],
    ['phát triên', 'phát triển', 'Thiếu dấu hỏi OCR']
  ];

  // Bộ bảo vệ thực thể quan trọng — KHÔNG BAO GIỜ tự động sửa nếu không có người dùng xác nhận
  function detectProtectedEntityReason(originalText, suggestedText) {
    const s = String(originalText || '').trim();
    const g = String(suggestedText || '').trim();
    if (!s) return null;

    // 1. Số CCCD / CMND (9 hoặc 12 chữ số) hoặc Mã số thuế (10 chữ số)
    if (/\b\d{9,13}\b/.test(s) || /\b\d{10}-\d{3}\b/.test(s)) {
      return 'Bảo vệ CCCD / Mã số thuế';
    }
    // 2. Số văn bản / Số hợp đồng (VD: 12/QĐ-UBND, 45-BC/ĐU, HĐ-2026)
    if (/\b\d+[A-Za-zÀ-Ỹà-ỹĐđ\-]*\/[A-ZÀ-ỸĐ]{2,}/i.test(s) || /\b(HĐ|Hợp đồng|Số)\s*:?\s*\d+/i.test(s)) {
      return 'Bảo vệ Số văn bản / Số hợp đồng';
    }
    // 3. Ngày tháng năm
    if (/\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/.test(s) || /\bngày\s+\d{1,2}\s+tháng\s+\d{1,2}/i.test(s)) {
      return 'Bảo vệ Ngày tháng năm';
    }
    // 4. Số tiền / Tỷ lệ % / Số liệu bất kỳ bị thay đổi chữ số
    if (/\d+\s*(%|đồng|vnđ|vnd|triệu|tỷ|tỉ)/i.test(s)) {
      return 'Bảo vệ Số tiền / Tỷ lệ %';
    }
    const dOrig = s.replace(/[^\d]/g, '');
    const dSug = g.replace(/[^\d]/g, '');
    if (dOrig.length > 0 && dOrig !== dSug) {
      return 'Bảo vệ Con số / Số liệu';
    }
    // 5. Ký hiệu pháp lý (Điều, Khoản, Điểm, Nghị định, Luật, Thông tư)
    if (/\b(Điều|Khoản|Điểm|Luật|Nghị định|Thông tư|Hướng dẫn)\s+(số\s+)?\d+/i.test(s)) {
      return 'Bảo vệ Ký hiệu pháp lý';
    }
    // 6. Họ tên riêng (2 từ trở lên viết hoa chữ cái đầu liên tiếp, không phải đầu câu quốc hiệu)
    if (/^(?:[A-ZĐÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸ][a-zà-ỹđ]+\s+){1,4}[A-ZĐÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴÈÉẸẺẼÊỀẾỆỂỄÌÍỊỈĨÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠÙÚỤỦŨƯỪỨỰỬỮỲÝỴỶỸ][a-zà-ỹđ]+$/.test(s)) {
      return 'Bảo vệ Họ tên / Địa danh riêng';
    }
    return null;
  }

  function scanOfflineOcrSuggestions(text) {
    const issues = [];
    const seen = new Set();
    for (const [wrong, suggest, reason] of OCR_OFFLINE_PAIRS) {
      const re = new RegExp(`(?<![A-Za-zÀ-Ỹà-ỹĐđ])${wrong.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Za-zÀ-Ỹà-ỹĐđ])`, 'g');
      if (re.test(text) && !seen.has(wrong)) {
        seen.add(wrong);
        const prot = detectProtectedEntityReason(wrong, suggest);
        issues.push({
          id: 'off_' + issues.length,
          original: wrong,
          suggestion: suggest,
          reason,
          source: 'Offline',
          protectedReason: prot,
          status: 'pending' // pending | accepted | rejected
        });
      }
    }
    return issues;
  }

  async function scanOllamaOcrSuggestions(text) {
    const out = [];
    try {
      const ctrlTags = new AbortController();
      const tTags = setTimeout(() => ctrlTags.abort(), 1600);
      const tagRes = await fetch('http://127.0.0.1:11434/api/tags', { signal: ctrlTags.signal });
      clearTimeout(tTags);
      if (!tagRes.ok) return { items: out, model: null, note: 'Ollama Local chưa bật — đang dùng chế độ Offline 100%.' };
      const tagJson = await tagRes.json();
      const models = (tagJson.models || []).map(m => m.name);
      if (!models.length) return { items: out, model: null, note: 'Chưa có model Ollama nào trên máy.' };

      const picked = models.find(n => /qwen2\.5:7b/i.test(n)) || models.find(n => /qwen2\.5:3b/i.test(n)) || models[0];
      const sample = text.length > 2600 ? text.slice(0, 2600) : text;

      const ctrl = new AbortController();
      const tGen = setTimeout(() => ctrl.abort(), 22000);
      const genRes = await fetch('http://127.0.0.1:11434/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: ctrl.signal,
        body: JSON.stringify({
          model: picked,
          stream: false,
          options: { temperature: 0.05, num_predict: 900 },
          system: 'Bạn là chuyên gia kiểm tra lỗi OCR tiếng Việt. Chỉ trả về JSON array: [{"original":"cụm sai đúng nguyên văn trong bài","suggestion":"cụm đã sửa","reason":"lý do ngắn gọn"}]. TUYỆT ĐỐI KHÔNG sửa họ tên, địa chỉ, số CCCD, mã số thuế, số hợp đồng, số văn bản, ngày tháng năm, số tiền. Nếu không có lỗi trả về [].',
          prompt: sample
        })
      });
      clearTimeout(tGen);
      if (!genRes.ok) return { items: out, model: picked, note: 'AI Local tạm bận.' };
      const gJson = await genRes.json();
      const raw = String(gJson.response || '').trim();
      const sIdx = raw.indexOf('[');
      const eIdx = raw.lastIndexOf(']');
      if (sIdx >= 0 && eIdx > sIdx) {
        const parsed = JSON.parse(raw.slice(sIdx, eIdx + 1));
        if (Array.isArray(parsed)) {
          parsed.forEach((it, idx) => {
            const orig = String(it?.original || '').trim();
            const sug = String(it?.suggestion || '').trim();
            const why = String(it?.reason || 'AI phát hiện lỗi nhận diện OCR').trim();
            if (!orig || !sug || orig === sug || orig.length > 70) return;
            if (!text.includes(orig)) return;
            const prot = detectProtectedEntityReason(orig, sug);
            out.push({
              id: 'ai_' + idx,
              original: orig,
              suggestion: sug,
              reason: why,
              source: `AI (${picked})`,
              protectedReason: prot,
              status: 'pending'
            });
          });
        }
      }
      return { items: out, model: picked, note: `Đã kiểm tra xong bằng AI Local (${picked}).` };
    } catch {
      return { items: out, model: null, note: 'Ollama Local chưa chạy — vẫn dùng đầy đủ bộ kiểm tra Offline.' };
    }
  }

  let modalState = { originalText: '', suggestions: [] };

  function ensureModalDom() {
    let modal = document.getElementById('ocrAiProofreadModal');
    if (modal) return modal;
    modal = document.createElement('div');
    modal.id = 'ocrAiProofreadModal';
    modal.style.cssText = 'display:none;position:fixed;inset:0;z-index:10001;background:rgba(15,23,42,0.82);backdrop-filter:blur(4px);align-items:center;justify-content:center;padding:16px;';
    modal.innerHTML = `
      <div style="width:100%;max-width:980px;max-height:90vh;background:#ffffff;color:#0f172a;border-radius:16px;box-shadow:0 24px 64px rgba(0,0,0,0.45);display:flex;flex-direction:column;overflow:hidden;font-family:'Segoe UI',system-ui,sans-serif;">
        <div style="background:linear-gradient(135deg,#1d4ed8,#2563eb);padding:14px 20px;display:flex;justify-content:space-between;align-items:center;color:#fff;">
          <div>
            <div style="font-size:1.05rem;font-weight:800;">🤖 AI Kiểm Tra &amp; Hiệu Chỉnh Lỗi OCR (Bảo Vệ Dữ Liệu Quan Trọng)</div>
            <div id="ocrAiModalSub" style="font-size:0.82rem;color:#dbeafe;margin-top:2px;">Đang phân tích văn bản OCR gốc...</div>
          </div>
          <button type="button" onclick="VietNamOcrAiProofread.close()" style="background:rgba(255,255,255,0.18);border:none;color:#fff;width:32px;height:32px;border-radius:8px;cursor:pointer;font-weight:700;">✕</button>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1.15fr;gap:14px;padding:16px 20px;flex:1;overflow:hidden;background:#f8fafc;">
          <div style="display:flex;flex-direction:column;min-height:320px;">
            <div style="font-size:0.85rem;font-weight:700;color:#334155;margin-bottom:6px;">📄 Văn bản OCR gốc (Giữ nguyên để đối chiếu)</div>
            <textarea id="ocrAiOriginalBox" readonly style="flex:1;width:100%;resize:none;border:1px solid #cbd5e1;border-radius:10px;padding:12px;font-family:'Times New Roman',serif;font-size:14.5px;line-height:1.55;background:#fff;color:#1e293b;"></textarea>
          </div>
          <div style="display:flex;flex-direction:column;min-height:320px;overflow:hidden;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
              <span id="ocrAiFoundLabel" style="font-size:0.85rem;font-weight:700;color:#1d4ed8;">ĐỀ XUẤT HIỆU CHỈNH</span>
              <button type="button" onclick="VietNamOcrAiProofread.acceptAllSafe()" style="padding:5px 12px;border-radius:8px;border:none;background:#10b981;color:#fff;font-size:0.8rem;font-weight:700;cursor:pointer;">✓ Chấp nhận tất cả (trừ mục Bảo vệ)</button>
            </div>
            <div id="ocrAiSuggestionList" style="flex:1;overflow-y:auto;display:flex;flex-direction:column;gap:8px;padding-right:4px;"></div>
          </div>
        </div>
        <div style="padding:12px 20px;background:#fff;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;">
          <span style="font-size:0.82rem;color:#64748b;">🛡 Tự động bảo vệ: Họ tên, Địa chỉ, CCCD, MST, Số HĐ/VB, Ngày tháng, Số tiền, Ký hiệu pháp lý.</span>
          <div style="display:flex;gap:10px;">
            <button type="button" onclick="VietNamOcrAiProofread.close()" style="padding:8px 16px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">Đóng</button>
            <button type="button" onclick="VietNamOcrAiProofread.applyAcceptedToEditor()" style="padding:8px 18px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;">✅ Áp dụng các mục đã chấp nhận</button>
          </div>
        </div>
      </div>`;
    document.body.appendChild(modal);
    return modal;
  }

  function esc(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function renderSuggestionCards() {
    const host = document.getElementById('ocrAiSuggestionList');
    const lbl = document.getElementById('ocrAiFoundLabel');
    if (!host) return;
    const total = modalState.suggestions.length;
    const accepted = modalState.suggestions.filter(x => x.status === 'accepted').length;
    if (lbl) lbl.textContent = `PHÁT HIỆN ${total} ĐỀ XUẤT (${accepted} đã chấp nhận)`;

    if (total === 0) {
      host.innerHTML = '<div style="background:#ecfdf5;border:1px solid #a7f3d0;color:#065f46;padding:16px;border-radius:10px;font-size:0.9rem;">✓ Không phát hiện lỗi OCR hoặc lỗi chính tả nào cần sửa!</div>';
      return;
    }

    host.innerHTML = modalState.suggestions.map((it, idx) => {
      const bg = it.status === 'accepted' ? '#f0fdf4' : (it.status === 'rejected' ? '#f8fafc' : '#ffffff');
      const op = it.status === 'rejected' ? '0.55' : '1';
      const protBadge = it.protectedReason
        ? `<span style="background:#fef3c7;color:#92400e;border:1px solid #fcd34d;padding:2px 6px;border-radius:6px;font-size:0.74rem;font-weight:700;">🛡 ${esc(it.protectedReason)}</span>`
        : `<span style="background:#eff6ff;color:#1d4ed8;padding:2px 6px;border-radius:6px;font-size:0.74rem;font-weight:600;">${esc(it.source)}</span>`;
      return `
        <div style="background:${bg};opacity:${op};border:1px solid #e2e8f0;border-radius:10px;padding:10px 12px;">
          <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:6px;">
            <span style="background:#fef2f2;color:#b91c1c;border:1px solid #fecaca;padding:2px 8px;border-radius:6px;font-weight:700;font-size:0.88rem;">${esc(it.original)}</span>
            <span style="color:#64748b;font-weight:700;">→</span>
            <span style="background:#f0fdf4;color:#15803d;border:1px solid #bbf7d0;padding:2px 8px;border-radius:6px;font-weight:700;font-size:0.88rem;">${esc(it.suggestion)}</span>
            ${protBadge}
          </div>
          <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;">
            <span style="font-size:0.8rem;color:#64748b;">${esc(it.reason)}</span>
            <div style="display:flex;gap:6px;">
              <button type="button" onclick="VietNamOcrAiProofread.setStatus(${idx}, 'accepted')" style="padding:4px 10px;border-radius:6px;border:1px solid #10b981;background:${it.status === 'accepted' ? '#10b981' : '#fff'};color:${it.status === 'accepted' ? '#fff' : '#047857'};font-size:0.78rem;font-weight:700;cursor:pointer;">✓ Chấp nhận</button>
              <button type="button" onclick="VietNamOcrAiProofread.setStatus(${idx}, 'rejected')" style="padding:4px 10px;border-radius:6px;border:1px solid #cbd5e1;background:${it.status === 'rejected' ? '#64748b' : '#fff'};color:${it.status === 'rejected' ? '#fff' : '#475569'};font-size:0.78rem;font-weight:600;cursor:pointer;">✕ Từ chối</button>
            </div>
          </div>
        </div>`;
    }).join('');
  }

  async function openModal() {
    const textEl = document.getElementById('ocrResultText');
    const text = (textEl?.value || '').trim();
    if (!text) {
      if (global.showToast) global.showToast('Chưa có văn bản OCR để kiểm tra. Hãy tải tệp và chạy OCR trước!');
      return;
    }

    const modal = ensureModalDom();
    const origBox = document.getElementById('ocrAiOriginalBox');
    const subEl = document.getElementById('ocrAiModalSub');
    if (origBox) origBox.value = text;
    if (subEl) subEl.textContent = '✓ Đã quét lớp Offline tức thì · Đang kiểm tra bổ sung bằng AI Local (Ollama)...';

    modalState.originalText = text;
    modalState.suggestions = scanOfflineOcrSuggestions(text);
    renderSuggestionCards();
    modal.style.display = 'flex';

    const aiRes = await scanOllamaOcrSuggestions(text);
    if (aiRes.items.length > 0) {
      const existing = new Set(modalState.suggestions.map(x => x.original));
      aiRes.items.forEach(it => {
        if (!existing.has(it.original)) modalState.suggestions.push(it);
      });
      renderSuggestionCards();
    }
    if (subEl) subEl.textContent = aiRes.note;
  }

  function setStatus(idx, st) {
    if (modalState.suggestions[idx]) {
      modalState.suggestions[idx].status = st;
      renderSuggestionCards();
    }
  }

  function acceptAllSafe() {
    modalState.suggestions.forEach(it => {
      if (!it.protectedReason && it.status !== 'rejected') {
        it.status = 'accepted';
      }
    });
    renderSuggestionCards();
  }

  function applyAcceptedToEditor() {
    let updated = modalState.originalText;
    let count = 0;
    modalState.suggestions.forEach(it => {
      if (it.status === 'accepted') {
        const re = new RegExp(it.original.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        updated = updated.replace(re, it.suggestion);
        count++;
      }
    });
    const textEl = document.getElementById('ocrResultText');
    const paperEl = document.getElementById('ocrFormattedPaper');
    if (textEl) textEl.value = updated;
    if (paperEl && typeof global.reconstructExactLayoutFromPlainText === 'function') {
      paperEl.innerHTML = global.reconstructExactLayoutFromPlainText(updated);
    }
    close();
    if (global.showToast) {
      global.showToast(count > 0 ? `✓ Đã áp dụng ${count} hiệu chỉnh OCR vào văn bản!` : 'Không có mục nào được chọn áp dụng — giữ nguyên bản OCR gốc.');
    }
  }

  function close() {
    const modal = document.getElementById('ocrAiProofreadModal');
    if (modal) modal.style.display = 'none';
  }

  global.VietNamOcrAiProofread = {
    openModal,
    close,
    setStatus,
    acceptAllSafe,
    applyAcceptedToEditor,
    detectProtectedEntityReason,
    scanOfflineOcrSuggestions
  };
})(window);
