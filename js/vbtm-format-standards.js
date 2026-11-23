/**
 * VBTM Pro v3.5.0 — Bộ Kiểm Tra & Chuẩn Hóa Thể Thức CHỦ ĐỘNG (KHÔNG tự động chạy khi mở/dán/lưu)
 * Tách riêng 2 hệ quy chuẩn độc lập:
 *   1. Văn bản Đảng: Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng
 *   2. Văn bản hành chính: Nghị định số 30/2020/NĐ-CP của Chính phủ
 *
 * Quy trình 8 bước bắt buộc:
 *   Chọn loại -> Kiểm tra -> Danh sách lỗi -> Đề xuất -> Xem trước -> Xác nhận -> Áp dụng -> Hoàn tác (Undo)
 */
(function (global) {
  'use strict';

  const undoSnapshots = [];

  function readDocFields() {
    return {
      standard: document.getElementById('vbtmStandard')?.value || 'nd30',
      orgName: document.getElementById('vbtmOrgName')?.value || '',
      docNum: document.getElementById('vbtmDocNum')?.value || '',
      docDate: document.getElementById('vbtmDocDate')?.value || '',
      docSubject: document.getElementById('vbtmDocSubject')?.value || '',
      docBody: document.getElementById('vbtmDocBody')?.value || '',
      recipients: document.getElementById('vbtmRecipients')?.value || '',
      signerRole: document.getElementById('vbtmSignerRole')?.value || ''
    };
  }

  function writeDocFields(snap) {
    const map = [
      ['vbtmStandard', snap.standard],
      ['vbtmOrgName', snap.orgName],
      ['vbtmDocNum', snap.docNum],
      ['vbtmDocDate', snap.docDate],
      ['vbtmDocSubject', snap.docSubject],
      ['vbtmDocBody', snap.docBody],
      ['vbtmRecipients', snap.recipients],
      ['vbtmSignerRole', snap.signerRole]
    ];
    map.forEach(([id, val]) => {
      const el = document.getElementById(id);
      if (el && val !== undefined) el.value = val;
    });
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    if (typeof global.updateVbtmLiveStats === 'function') global.updateVbtmLiveStats();
  }

  // 1. KIỂM TRA THỂ THỨC (CHỈ ĐỌC — KHÔNG SỬA NỘI DUNG)
  function auditStandard(standardType, customFields) {
    const f = customFields || readDocFields();
    const isParty = standardType === 'hd05';
    const issues = [];

    if (isParty) {
      // QUY CHUẨN VĂN BẢN ĐẢNG — HƯỚNG DẪN SỐ 05-HD/VPTW
      if (/\b36-HD\/VPTW\b/i.test(f.docBody + ' ' + f.docSubject)) {
        issues.push({ level: 'Nghiêm trọng', field: 'Căn cứ', detail: 'Phát hiện trích dẫn Hướng dẫn 36-HD/VPTW đã hết hiệu lực.', suggestion: 'Thay bằng Hướng dẫn số 05-HD/VPTW.' });
      }
      if (!f.docNum.trim()) {
        issues.push({ level: 'Cần sửa', field: 'Số và ký hiệu', detail: 'Thiếu số, ký hiệu văn bản Đảng (mẫu: 12-CV/ĐU hoặc 05-BC/ĐU).', suggestion: 'Bổ sung số ký hiệu theo Hướng dẫn 05-HD/VPTW.' });
      }
      if (!f.docSubject.trim()) {
        issues.push({ level: 'Cần sửa', field: 'Trích yếu', detail: 'Thiếu trích yếu nội dung văn bản Đảng.', suggestion: 'Nhập trích yếu ngắn gọn, rõ nội dung.' });
      }
      if (f.recipients && !/Lưu Văn phòng Đảng ủy|Lưu VPĐU/i.test(f.recipients)) {
        issues.push({ level: 'Cần xem xét', field: 'Nơi nhận', detail: 'Mục Nơi nhận chưa có dòng lưu theo chuẩn văn bản Đảng.', suggestion: 'Bổ sung "- Lưu Văn phòng Đảng ủy." ở dòng cuối Nơi nhận.' });
      }
      if (f.docBody && /\s+([,.;:!?])/.test(f.docBody)) {
        issues.push({ level: 'Cần sửa', field: 'Dấu câu & trình bày', detail: 'Có khoảng trắng thừa trước dấu câu trong nội dung.', suggestion: 'Xóa khoảng trắng thừa trước dấu câu (giữ nguyên từ ngữ & ý nghĩa).' });
      }
      issues.push({ level: 'Thông số chuẩn', field: 'Khổ giấy & Phông chữ (05-HD/VPTW)', detail: 'Khổ A4 (210x297mm) · Times New Roman 14pt · Lề Trái 30–35mm, Phải 15–20mm, Trên/Dưới 20–25mm · Đánh số trang giữa lề trên từ trang 2.', suggestion: 'Áp dụng khi xuất DOCX/PDF.' });
    } else {
      // QUY CHUẨN VĂN BẢN HÀNH CHÍNH — NGHỊ ĐỊNH 30/2020/NĐ-CP
      if (!f.docNum.trim()) {
        issues.push({ level: 'Cần sửa', field: 'Số và ký hiệu', detail: 'Thiếu số, ký hiệu văn bản hành chính (mẫu: 15/UBND-VP).', suggestion: 'Bổ sung số/ký hiệu chuẩn NĐ 30/2020/NĐ-CP.' });
      }
      if (/\bngày\s+[1-9]\s+tháng\s+[12]\b/i.test(f.docBody + ' ' + f.docDate)) {
        issues.push({ level: 'Cần sửa', field: 'Ngày tháng (NĐ 30)', detail: 'Ngày dưới 10 hoặc tháng 1, tháng 2 cần ghi thêm số 0 ở trước (VD: ngày 05 tháng 02).', suggestion: 'Chuẩn hóa số 0 cho ngày < 10 và tháng 1–2.' });
      }
      if (f.recipients && !/Lưu:\s*VT/i.test(f.recipients)) {
        issues.push({ level: 'Cần xem xét', field: 'Nơi nhận', detail: 'Mục Nơi nhận chưa có dòng lưu chuẩn Nghị định 30.', suggestion: 'Bổ sung "- Lưu: VT, ..." ở dòng cuối Nơi nhận.' });
      }
      if (f.docBody.trim() && !f.docBody.trim().endsWith('./.')) {
        issues.push({ level: 'Gợi ý cải thiện', field: 'Dấu kết thúc', detail: 'Cuối phần nội dung văn bản hành chính thường có dấu kết thúc "./.".', suggestion: 'Thêm "./." vào cuối đoạn kết nếu là văn bản ban hành chính thức.' });
      }
      if (f.docBody && /\s+([,.;:!?])/.test(f.docBody)) {
        issues.push({ level: 'Cần sửa', field: 'Dấu câu & trình bày', detail: 'Có khoảng trắng thừa trước dấu câu.', suggestion: 'Xóa khoảng trắng thừa trước dấu câu.' });
      }
      issues.push({ level: 'Thông số chuẩn', field: 'Khổ giấy & Phông chữ (NĐ 30/2020)', detail: 'Khổ A4 (210x297mm) · Times New Roman 13–14pt · Lề Trái 30–35mm, Phải 15–20mm, Trên/Dưới 20–25mm · Đánh số trang giữa lề trên từ trang 2.', suggestion: 'Áp dụng khi xuất DOCX/PDF.' });
    }

    return {
      standard: isParty ? 'hd05' : 'nd30',
      standardName: isParty ? 'Hướng dẫn số 05-HD/VPTW (Văn bản Đảng)' : 'Nghị định 30/2020/NĐ-CP (Văn bản hành chính)',
      issues
    };
  }

  // 2. TẠO BẢN XEM TRƯỚC CHUẨN HÓA (CHƯA ÁP DỤNG — ĐỂ NGƯỜI DÙNG XEM & XÁC NHẬN)
  function buildProposedStandardization(standardType, customFields) {
    const orig = customFields || readDocFields();
    const next = Object.assign({}, orig);
    const isParty = standardType === 'hd05';
    next.standard = isParty ? 'hd05' : 'nd30';

    // Chỉ chỉnh đúng phạm vi thể thức được phép, KHÔNG thay đổi nội dung chuyên môn hay số liệu
    if (next.docNum) next.docNum = next.docNum.trim().replace(/^Số\s*:?\s*/i, '');

    if (next.docBody) {
      let body = next.docBody
        .replace(/[ \t]+/g, ' ')
        .replace(/\s+([,.;:!?])/g, '$1')
        .replace(/\n{3,}/g, '\n\n');
      if (isParty) {
        body = body.replace(/\b36-HD\/VPTW\b/g, '05-HD/VPTW');
      } else {
        body = body.replace(/\bngày\s+([1-9])\s+tháng\s+([12])\b/gi, 'ngày 0$1 tháng 0$2');
        body = body.replace(/\bngày\s+([1-9])\s+tháng\s+(\d{1,2})\b/gi, (m, d, mo) => `ngày 0${d} tháng ${mo.length === 1 && (mo === '1' || mo === '2') ? '0' + mo : mo}`);
        if (body.trim() && !body.trim().endsWith('./.')) body = body.trim().replace(/\.+$/, '') + './.';
      }
      next.docBody = body;
    }

    if (next.recipients) {
      let lines = next.recipients.split('\n').map(l => l.trim()).filter(Boolean).map(l => l.startsWith('-') ? l : '- ' + l);
      if (isParty && !lines.some(l => /Lưu/i.test(l))) lines.push('- Lưu Văn phòng Đảng ủy.');
      if (!isParty && !lines.some(l => /Lưu/i.test(l))) lines.push('- Lưu: VT, VP.');
      lines = lines.map((l, idx) => {
        const clean = l.replace(/[;.]+\s*$/, '');
        return idx < lines.length - 1 ? clean + ';' : clean + '.';
      });
      next.recipients = lines.join('\n');
    }

    return { before: orig, after: next, standardName: isParty ? 'Hướng dẫn 05-HD/VPTW' : 'Nghị định 30/2020/NĐ-CP' };
  }

  // 3. ÁP DỤNG CÓ XÁC NHẬN + LƯU BẢN HOÀN TÁC (UNDO)
  function applyStandardizationWithUndo(standardType) {
    const proposal = buildProposedStandardization(standardType);
    const audit = auditStandard(standardType, proposal.before);
    const summary = audit.issues.map((x, i) => `${i + 1}. [${x.level}] ${x.field}: ${x.detail}`).join('\n');

    const ok = typeof confirm === 'function'
      ? confirm(`XEM TRƯỚC CHUẨN HÓA (${proposal.standardName})\n\nCác mục kiểm tra:\n${summary}\n\nBạn có xác nhận áp dụng chuẩn hóa thể thức này không? (Có thể bấm "Hoàn tác" bất cứ lúc nào.)`)
      : true;
    if (!ok) return { applied: false, reason: 'cancelled' };

    undoSnapshots.push(proposal.before);
    writeDocFields(proposal.after);
    if (global.showToast) global.showToast(`✓ Đã áp dụng chuẩn hóa ${proposal.standardName} (có thể Hoàn tác).`);
    return { applied: true, standard: standardType, undoDepth: undoSnapshots.length };
  }

  function undoLastStandardization() {
    if (!undoSnapshots.length) {
      if (global.showToast) global.showToast('Không có thao tác chuẩn hóa nào để hoàn tác.');
      return false;
    }
    const prev = undoSnapshots.pop();
    writeDocFields(prev);
    if (global.showToast) global.showToast('↩ Đã hoàn tác về nguyên trạng trước khi chuẩn hóa.');
    return true;
  }

  function renderAuditPanel(standardType) {
    const box = document.getElementById('vbtmFormatAuditResult');
    const audit = auditStandard(standardType);
    if (!box) return audit;
    box.style.display = 'block';
    box.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:6px;">
        <strong style="color:#fef08a;">Kết quả kiểm tra: ${audit.standardName} (${audit.issues.length} mục)</strong>
        <div style="display:flex;gap:6px;">
          <button type="button" class="btn-tool primary" onclick="VbtmFormatStandards.applyStandardizationWithUndo('${audit.standard}')">✅ Xác nhận &amp; Chuẩn hóa ${audit.standard === 'hd05' ? '05-HD/VPTW' : 'NĐ 30/2020'}</button>
          <button type="button" class="btn-tool" onclick="VbtmFormatStandards.undoLastStandardization()">↩ Hoàn tác</button>
          <button type="button" class="btn-tool" onclick="document.getElementById('vbtmFormatAuditResult').style.display='none'">✕ Đóng</button>
        </div>
      </div>
      <ul style="margin:0 0 0 18px;padding:0;font-size:0.86rem;line-height:1.65;color:#e2e8f0;">
        ${audit.issues.map(it => `<li><strong>[${it.level}] ${it.field}:</strong> ${it.detail} <em style="color:#86efac;">→ ${it.suggestion}</em></li>`).join('')}
      </ul>
    `;
    return audit;
  }

  global.VbtmFormatStandards = {
    auditStandard,
    buildProposedStandardization,
    applyStandardizationWithUndo,
    undoLastStandardization,
    renderAuditPanel
  };
})(typeof window !== 'undefined' ? window : globalThis);
