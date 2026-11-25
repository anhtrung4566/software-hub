/**
 * VietNamOCR 3.5 Pro — Intelligent Document Processing Studio (js/ocr-idp-studio.js)
 *
 * Bao gồm 4 nhóm năng lực thực tế (100% xử lý cục bộ trên máy người dùng):
 * 1. AI Hiểu & Phân tích tài liệu có dẫn chiếu [Đoạn #]: Tóm tắt, Hỏi đáp Q&A, Trích xuất thực thể (Họ tên, Ngày, Số tiền, Số VB/HĐ, CCCD, MST), Phân loại, Mục lục tự động, Phát hiện trùng lặp, So sánh 2 tài liệu.
 * 2. Xuất tài liệu đa định dạng thực thụ: DOCX (.docx OOXML chỉnh sửa được), Excel (.xlsx/.xls), Markdown (.md), JSON có cấu trúc (.json), TXT (.txt).
 * 3. Kho lưu trữ & Tìm kiếm tài liệu thông minh (localStorage cục bộ): Tự động lưu kết quả OCR, tìm theo tên file/nội dung/từ khóa/ngày/số VB/họ tên, mở lại đúng tài liệu.
 * 4. Quản lý dữ liệu & Bảo mật: Sao lưu / Khôi phục JSON, Xóa sạch dữ liệu OCR & bộ nhớ tạm 1-click, xem nhật ký lỗi kỹ thuật (không lưu nội dung nhạy cảm).
 */
(function (global) {
  'use strict';

  const ARCHIVE_KEY = 'vbtm_ocr_idp_archive_v1';
  const MAX_ARCHIVE_ITEMS = 40;

  function escHtml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function escXml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  // =========================================================================
  // 1. TRÍCH XUẤT THỰC THỂ, MỤC LỤC, PHÂN LOẠI & PHÁT HIỆN TRÙNG LẶP (OFFLINE 100%)
  // =========================================================================
  function analyzeDocumentStructure(rawText) {
    const text = String(rawText || '').normalize('NFC').trim();
    const paras = text.split(/\n+/).map(p => p.trim()).filter(Boolean);

    // Phân loại tài liệu
    let docType = 'Văn bản tổng hợp';
    if (/ĐẢNG CỘNG SẢN VIỆT NAM|ĐẢNG ỦY|CHI BỘ|BAN THƯỜNG VỤ|05-HD\/VPTW|\/ĐU\b/i.test(text)) {
      docType = 'Văn bản Đảng (Hướng dẫn 05-HD/VPTW)';
    } else if (/CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|ỦY BAN NHÂN DÂN|NGHỊ ĐỊNH 30|QUYẾT ĐỊNH|CÔNG VĂN/i.test(text)) {
      docType = 'Văn bản Hành chính (NĐ 30/2020/NĐ-CP)';
    } else if (/HỢP ĐỒNG|BÊN A|BÊN B|GIÁ TRỊ HỢP ĐỒNG|MÃ SỐ THUẾ/i.test(text)) {
      docType = 'Hợp đồng / Tài chính';
    }

    const uniq = (arr) => Array.from(new Set(arr)).slice(0, 20);
    const entities = {
      docNumbers: [],
      dates: [],
      amounts: [],
      cccdMst: [],
      names: []
    };
    const toc = [];
    const duplicates = [];
    const seenNorm = new Map();

    paras.forEach((p, idx) => {
      const ref = `Đoạn ${idx + 1}`;
      // Số văn bản / Hợp đồng
      (p.match(/\b(?:Số\s*:?\s*)?\d+[A-Za-zÀ-Ỹà-ỹĐđ\-]*\/[A-ZÀ-ỸĐ0-9\-]{2,}/g) || []).forEach(m => entities.docNumbers.push(`${m.trim()} [${ref}]`));
      (p.match(/\bHĐ[\-\s]?\d[\w\-\/]+/gi) || []).forEach(m => entities.docNumbers.push(`${m.trim()} [${ref}]`));
      // Ngày tháng năm
      (p.match(/\bngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}\b/gi) || []).forEach(m => entities.dates.push(`${m.trim()} [${ref}]`));
      (p.match(/\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}\b/g) || []).forEach(m => entities.dates.push(`${m.trim()} [${ref}]`));
      // Số tiền
      (p.match(/\b\d{1,3}(?:[.,]\d{3})+(?:\s*(?:đồng|VNĐ|VND|triệu|tỷ|tỉ))?/gi) || []).forEach(m => {
        if (m.replace(/[^\d]/g, '').length >= 4) entities.amounts.push(`${m.trim()} [${ref}]`);
      });
      // CCCD / MST
      (p.match(/\b\d{9,12}\b|\b\d{10}-\d{3}\b/g) || []).forEach(m => entities.cccdMst.push(`${m.trim()} [${ref}]`));
      // Họ tên sau chức danh/từ khóa
      (p.match(/(?:Ông|Bà|Đồng chí|Đ\/c|Chủ tịch|Bí thư|Giám đốc|Trưởng phòng)\s*:?\s*([A-ZĐÀ-Ỹ][a-zà-ỹđ]+(?:\s+[A-ZĐÀ-Ỹ][a-zà-ỹđ]+){1,3})/g) || []).forEach(m => entities.names.push(`${m.trim()} [${ref}]`));

      // Mục lục tự động
      if (/^(PHẦN\s+[IVX0-9]+|CHƯƠNG\s+[IVX0-9]+|[IVXLCDM]+\.\s+|Điều\s+\d+\.|[1-9]\.\s+[A-ZĐÀ-Ỹ])/i.test(p) && p.length <= 140) {
        toc.push({ ref, title: p });
      }

      // Phát hiện đoạn trùng lặp (độ dài >= 30 ký tự)
      const norm = p.toLowerCase().replace(/\s+/g, ' ');
      if (norm.length >= 30) {
        if (seenNorm.has(norm)) {
          duplicates.push({ firstRef: seenNorm.get(norm), dupRef: ref, preview: p.slice(0, 90) });
        } else {
          seenNorm.set(norm, ref);
        }
      }
    });

    return {
      docType,
      paragraphCount: paras.length,
      wordCount: text ? text.split(/\s+/).length : 0,
      entities: {
        docNumbers: uniq(entities.docNumbers),
        dates: uniq(entities.dates),
        amounts: uniq(entities.amounts),
        cccdMst: uniq(entities.cccdMst),
        names: uniq(entities.names)
      },
      toc,
      duplicates
    };
  }

  // =========================================================================
  // 2. AI TÓM TẮT & HỎI ĐÁP DỰA TRÊN TÀI LIỆU THỰC TẾ (CÓ DẪN CHIẾU ĐOẠN)
  // =========================================================================
  async function askGroundedAi(question, rawText) {
    const paras = String(rawText || '').split(/\n+/).map(p => p.trim()).filter(Boolean);
    if (paras.length === 0) return 'Chưa có nội dung tài liệu OCR để phân tích.';

    const numberedContext = paras.map((p, i) => `[Đoạn ${i + 1}] ${p}`).join('\n');
    const clipped = numberedContext.length > 5500 ? numberedContext.slice(0, 5500) : numberedContext;

    try {
      const ctrlT = new AbortController();
      const tT = setTimeout(() => ctrlT.abort(), 1600);
      const tagRes = await fetch('http://127.0.0.1:11434/api/tags', { signal: ctrlT.signal });
      clearTimeout(tT);
      if (tagRes.ok) {
        const tagJson = await tagRes.json();
        const models = (tagJson.models || []).map(m => m.name);
        if (models.length > 0) {
          const picked = models.find(n => /qwen2\.5:7b/i.test(n)) || models.find(n => /qwen2\.5:3b/i.test(n)) || models[0];
          const ctrlG = new AbortController();
          const tG = setTimeout(() => ctrlG.abort(), 28000);
          const gRes = await fetch('http://127.0.0.1:11434/api/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            signal: ctrlG.signal,
            body: JSON.stringify({
              model: picked,
              stream: false,
              options: { temperature: 0.1, num_predict: 700 },
              system: 'Bạn là trợ lý phân tích tài liệu OCR tiếng Việt. CHỈ trả lời dựa trên nội dung tài liệu được cung cấp. LUÔN dẫn chiếu [Đoạn X] làm căn cứ. TUYỆT ĐỐI KHÔNG bịa thông tin không có trong tài liệu; nếu tài liệu không đề cập hãy nói rõ "Tài liệu không có thông tin này".',
              prompt: `NỘI DUNG TÀI LIỆU:\n${clipped}\n\nYÊU CẦU: ${question}`
            })
          });
          clearTimeout(tG);
          if (gRes.ok) {
            const gJson = await gRes.json();
            if (gJson && gJson.response) {
              return `[AI Local: ${picked}]\n${gJson.response.trim()}`;
            }
          }
        }
      }
    } catch {}

    // Fallback Offline thông minh: tìm các đoạn chứa từ khóa khớp nhất và trích dẫn nguyên văn
    const qWords = String(question || '').toLowerCase().split(/\s+/).filter(w => w.length >= 2);
    const scored = paras.map((p, idx) => {
      const lp = p.toLowerCase();
      let score = 0;
      qWords.forEach(w => { if (lp.includes(w)) score++; });
      return { idx: idx + 1, text: p, score };
    }).filter(x => x.score > 0).sort((a, b) => b.score - a.score).slice(0, 5);

    if (scored.length === 0) {
      return `[Chế độ Offline — Ollama chưa bật]\nTóm tắt nhanh: Tài liệu có ${paras.length} đoạn.\n• [Đoạn 1] ${paras[0]}\n${paras[1] ? '• [Đoạn 2] ' + paras[1] : ''}`;
    }
    return `[Tra cứu căn cứ trực tiếp trong tài liệu (Offline)]\n` +
      scored.map(s => `• [Đoạn ${s.idx}] ${s.text}`).join('\n');
  }

  // =========================================================================
  // 3. XUẤT ĐA ĐỊNH DẠNG: REAL DOCX (OOXML), MARKDOWN (.md), JSON (.json)
  // =========================================================================
  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // Xuất tệp Word .DOCX chuẩn OOXML qua bộ xuất thống nhất VietNamDocxExporter (bảo toàn 100% font Unicode & bố cục)
  async function exportOcrToDocx() {
    const paperEl = document.getElementById('ocrFormattedPaper');
    const plainWrap = document.getElementById('ocrPlainEditorWrap');
    const text = (document.getElementById('ocrResultText')?.value || '').trim();
    const hasFormattedHtml = paperEl && paperEl.innerHTML.trim() && (!plainWrap || plainWrap.style.display === 'none');
    if (!text && !hasFormattedHtml) {
      if (global.showToast) global.showToast('Chưa có nội dung OCR để xuất .DOCX!');
      return;
    }
    const baseName = (typeof global.uploadedFileBaseName === 'string' && global.uploadedFileBaseName) ? global.uploadedFileBaseName : 'VanBan_OCR';
    if (global.VietNamDocxExporter && global.JSZip) {
      await global.VietNamDocxExporter.exportUnifiedDocx({
        html: hasFormattedHtml ? paperEl.innerHTML : '',
        plainText: text,
        fileName: baseName
      });
      return;
    }
    if (typeof global.exportOcrToWord === 'function') global.exportOcrToWord();
  }

  function exportOcrToMarkdown() {
    const text = (document.getElementById('ocrResultText')?.value || '').trim();
    if (!text) return;
    const baseName = (typeof global.uploadedFileBaseName === 'string' && global.uploadedFileBaseName) ? global.uploadedFileBaseName : 'VanBan_OCR';
    const analysis = analyzeDocumentStructure(text);
    const mdLines = [
      `# ${baseName}`,
      '',
      `> **Phân loại:** ${analysis.docType} | **Số đoạn:** ${analysis.paragraphCount} | **Số từ:** ${analysis.wordCount}`,
      ''
    ];
    text.split(/\n+/).forEach(line => {
      const t = line.trim();
      if (!t) return;
      if (/^(PHẦN|CHƯƠNG|[IVXLCDM]+\.)\s+/i.test(t) && t.length < 120) mdLines.push(`## ${t}`);
      else if (/^(Điều\s+\d+\.)/i.test(t)) mdLines.push(`### ${t}`);
      else mdLines.push(t);
      mdLines.push('');
    });
    triggerDownload(new Blob([mdLines.join('\n')], { type: 'text/markdown;charset=utf-8' }), `${baseName}.md`);
    if (global.showToast) global.showToast(`✓ Đã xuất Markdown: ${baseName}.md`);
  }

  function exportOcrToStructuredJson() {
    const text = (document.getElementById('ocrResultText')?.value || '').trim();
    if (!text) return;
    const baseName = (typeof global.uploadedFileBaseName === 'string' && global.uploadedFileBaseName) ? global.uploadedFileBaseName : 'VanBan_OCR';
    const analysis = analyzeDocumentStructure(text);
    const metrics = global.VietNamOcrEnginePro ? global.VietNamOcrEnginePro.getLastMetrics() : null;
    const payload = {
      fileName: baseName,
      exportedAt: new Date().toISOString(),
      ocrQuality: metrics ? {
        docProfile: metrics.docProfile,
        confidence: metrics.confidence,
        lowConfWords: metrics.lowConfCount,
        confusableTokens: metrics.confusableTokens,
        deskewAngle: metrics.deskewAngle,
        elapsedMs: metrics.elapsedMs,
        engineSource: metrics.engineSource
      } : null,
      classification: analysis.docType,
      stats: { paragraphs: analysis.paragraphCount, words: analysis.wordCount },
      entities: analysis.entities,
      tableOfContents: analysis.toc,
      duplicates: analysis.duplicates,
      paragraphs: text.split(/\n+/).map((p, i) => ({ index: i + 1, text: p.trim() })).filter(x => x.text)
    };
    triggerDownload(new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' }), `${baseName}.json`);
    if (global.showToast) global.showToast(`✓ Đã xuất JSON có cấu trúc: ${baseName}.json`);
  }

  // =========================================================================
  // 4. KHO LƯU TRỮ TÀI LIỆU OCR CỤC BỘ & TÌM KIẾM THÔNG MINH (LOCALSTORAGE)
  // =========================================================================
  function loadArchive() {
    try {
      const raw = localStorage.getItem(ARCHIVE_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch { return []; }
  }

  function saveArchive(arr) {
    try {
      localStorage.setItem(ARCHIVE_KEY, JSON.stringify(arr.slice(0, MAX_ARCHIVE_ITEMS)));
    } catch (e) {
      if (global.VietNamOcrEnginePro) global.VietNamOcrEnginePro.logSafeError('ArchiveSave', 'Bộ nhớ lưu trữ cục bộ đã đầy');
    }
  }

  function autoSaveDocumentToArchive(fileName, text, confidence = 0) {
    if (!text || text.trim().length < 10) return;
    const analysis = analyzeDocumentStructure(text);
    const list = loadArchive();
    const id = 'doc_' + Date.now();
    const entry = {
      id,
      fileName: fileName || 'TaiLieu_OCR',
      savedAt: new Date().toLocaleString('vi-VN'),
      docType: analysis.docType,
      confidence: Number(confidence) || 0,
      docNumbers: analysis.entities.docNumbers.slice(0, 3),
      dates: analysis.entities.dates.slice(0, 3),
      names: analysis.entities.names.slice(0, 4),
      wordCount: analysis.wordCount,
      text: text.slice(0, 35000)
    };
    const existingIdx = list.findIndex(x => x.fileName === entry.fileName);
    if (existingIdx >= 0) list.splice(existingIdx, 1);
    list.unshift(entry);
    saveArchive(list);
  }

  function clearAllLocalOcrData() {
    try {
      localStorage.removeItem(ARCHIVE_KEY);
      if (global.VietNamOcr2LayerPdf && global.VietNamOcr2LayerPdf.clearMultiPageStore) {
        global.VietNamOcr2LayerPdf.clearMultiPageStore();
      }
      if (global.VietNamOcrEnginePro) {
        global.VietNamOcrEnginePro.clearErrorLogs();
        global.VietNamOcrEnginePro.releaseWorkerMemory();
      }
      if (global.showToast) global.showToast('✓ Đã xóa sạch toàn bộ lịch sử OCR, bộ nhớ tạm và giải phóng RAM!');
    } catch (e) {}
  }

  function exportBackupArchive() {
    const data = { version: '3.6.0', exportedAt: new Date().toISOString(), documents: loadArchive() };
    triggerDownload(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json;charset=utf-8' }), `VietNamOCR_Backup_${new Date().toISOString().slice(0, 10)}.json`);
  }

  function importBackupArchive(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(String(reader.result || '{}'));
        const docs = Array.isArray(parsed.documents) ? parsed.documents : [];
        if (docs.length > 0) {
          saveArchive(docs);
          renderArchiveList('');
          if (global.showToast) global.showToast(`✓ Đã khôi phục ${docs.length} tài liệu OCR từ bản sao lưu!`);
        }
      } catch {
        if (global.showToast) global.showToast('⚠ Tệp sao lưu JSON không hợp lệ.');
      }
    };
    reader.readAsText(file, 'utf-8');
  }

  // =========================================================================
  // 5. GIAO DIỆN MODAL IDP STUDIO (AI HIỂU TÀI LIỆU + KHO TÌM KIẾM + SO SÁNH + BẢO MẬT)
  // =========================================================================
  function ensureIdpModal() {
    let m = document.getElementById('ocrIdpStudioModal');
    if (m) return m;
    m = document.createElement('div');
    m.id = 'ocrIdpStudioModal';
    m.style.cssText = 'display:none;position:fixed;inset:0;z-index:10002;background:rgba(15,23,42,0.84);backdrop-filter:blur(4px);align-items:center;justify-content:center;padding:16px;';
    m.innerHTML = `
      <div style="width:100%;max-width:1020px;max-height:92vh;background:#ffffff;color:#0f172a;border-radius:16px;box-shadow:0 24px 64px rgba(0,0,0,0.5);display:flex;flex-direction:column;overflow:hidden;font-family:'Segoe UI',system-ui,sans-serif;">
        <div style="background:linear-gradient(135deg,#0f172a,#1d4ed8);padding:14px 20px;display:flex;justify-content:space-between;align-items:center;color:#fff;">
          <div>
            <div style="font-size:1.06rem;font-weight:800;">🧠 Trung Tâm Xử Lý Tài Liệu Thông Minh (IDP &amp; Kho Tìm Kiếm Cục Bộ)</div>
            <div style="font-size:0.82rem;color:#cbd5e1;margin-top:2px;">Phân tích cấu trúc, Trích xuất thực thể, Hỏi đáp dẫn chiếu đoạn, So sánh tài liệu &amp; Bảo mật 100% trên máy</div>
          </div>
          <button type="button" onclick="VietNamOcrIdpStudio.close()" style="background:rgba(255,255,255,0.16);border:none;color:#fff;width:32px;height:32px;border-radius:8px;cursor:pointer;font-weight:700;">✕</button>
        </div>
        <div style="display:flex;gap:8px;padding:10px 20px;background:#f1f5f9;border-bottom:1px solid #e2e8f0;flex-wrap:wrap;">
          <button type="button" id="idpTabBtn1" onclick="VietNamOcrIdpStudio.switchTab(1)" style="padding:7px 14px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;">📊 Phân tích &amp; Hỏi đáp AI</button>
          <button type="button" id="idpTabBtn2" onclick="VietNamOcrIdpStudio.switchTab(2)" style="padding:7px 14px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">🔍 Kho Lưu Trữ &amp; Tìm Kiếm</button>
          <button type="button" id="idpTabBtn3" onclick="VietNamOcrIdpStudio.switchTab(3)" style="padding:7px 14px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">⚖️ So Sánh 2 Tài Liệu</button>
          <button type="button" id="idpTabBtn4" onclick="VietNamOcrIdpStudio.switchTab(4)" style="padding:7px 14px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">🛡 Bảo Mật, Sao Lưu &amp; Nhật Ký</button>
        </div>
        <div style="flex:1;overflow-y:auto;padding:16px 20px;background:#f8fafc;">
          <div id="idpPane1"></div>
          <div id="idpPane2" style="display:none;"></div>
          <div id="idpPane3" style="display:none;"></div>
          <div id="idpPane4" style="display:none;"></div>
        </div>
      </div>`;
    document.body.appendChild(m);
    return m;
  }

  function switchTab(num) {
    for (let i = 1; i <= 4; i++) {
      const p = document.getElementById('idpPane' + i);
      const b = document.getElementById('idpTabBtn' + i);
      if (p) p.style.display = (i === num) ? 'block' : 'none';
      if (b) {
        b.style.background = (i === num) ? '#2563eb' : '#fff';
        b.style.color = (i === num) ? '#fff' : '#334155';
      }
    }
    if (num === 1) renderAnalysisPane();
    if (num === 2) renderArchivePane();
    if (num === 3) renderDiffPane();
    if (num === 4) renderSecurityPane();
  }

  function renderAnalysisPane() {
    const pane = document.getElementById('idpPane1');
    if (!pane) return;
    const text = (document.getElementById('ocrResultText')?.value || '').trim();
    const a = analyzeDocumentStructure(text);
    const metrics = global.VietNamOcrEnginePro ? global.VietNamOcrEnginePro.getLastMetrics() : null;

    const qualityBanner = metrics
      ? `<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:10px 14px;margin-bottom:12px;font-size:0.86rem;color:#1e3a8a;">
          <strong>Độ tin cậy OCR trang gần nhất:</strong> <b>${metrics.confidence}%</b> · Loại nhận diện: <b>${escHtml(metrics.docProfile)}</b> · Từ cần kiểm tra (&lt;70%): <b>${metrics.lowConfCount}</b> · Ký tự dễ nhầm (O/0, I/1, S/5): <b>${metrics.confusableCount}</b> · Góc xoay căn thẳng: <b>${metrics.deskewAngle}°</b> · Tốc độ: <b>${metrics.elapsedMs}ms</b>
        </div>` : '';

    pane.innerHTML = `
      ${qualityBanner}
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:14px;">
        <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;">
          <div style="font-weight:700;color:#1d4ed8;margin-bottom:6px;">📌 Phân loại &amp; Thực thể trích xuất</div>
          <div style="font-size:0.85rem;line-height:1.65;">
            <div><strong>Loại tài liệu:</strong> ${escHtml(a.docType)} (${a.paragraphCount} đoạn · ${a.wordCount} từ)</div>
            <div><strong>Số văn bản / HĐ:</strong> ${a.entities.docNumbers.length ? escHtml(a.entities.docNumbers.join('; ')) : '<em>Chưa phát hiện</em>'}</div>
            <div><strong>Ngày tháng:</strong> ${a.entities.dates.length ? escHtml(a.entities.dates.join('; ')) : '<em>Chưa phát hiện</em>'}</div>
            <div><strong>Số tiền / Số liệu:</strong> ${a.entities.amounts.length ? escHtml(a.entities.amounts.join('; ')) : '<em>Chưa phát hiện</em>'}</div>
            <div><strong>CCCD / MST:</strong> ${a.entities.cccdMst.length ? escHtml(a.entities.cccdMst.join('; ')) : '<em>Chưa phát hiện</em>'}</div>
            <div><strong>Họ tên / Chức danh:</strong> ${a.entities.names.length ? escHtml(a.entities.names.join('; ')) : '<em>Chưa phát hiện</em>'}</div>
          </div>
        </div>
        <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;">
          <div style="font-weight:700;color:#1d4ed8;margin-bottom:6px;">📑 Mục lục tự động &amp; Kiểm tra trùng lặp</div>
          <div style="font-size:0.84rem;max-height:145px;overflow-y:auto;line-height:1.6;">
            ${a.toc.length ? a.toc.map(t => `<div>• <b>[${t.ref}]</b> ${escHtml(t.title)}</div>`).join('') : '<em>Không có đề mục rõ ràng.</em>'}
            ${a.duplicates.length ? `<div style="margin-top:8px;color:#b91c1c;"><strong>⚠ Phát hiện ${a.duplicates.length} đoạn trùng lặp:</strong><br/>${a.duplicates.map(d => `• ${d.firstRef} trùng ${d.dupRef}: "${escHtml(d.preview)}..."`).join('<br/>')}</div>` : '<div style="margin-top:8px;color:#15803d;">✓ Không phát hiện đoạn nội dung bị trùng lặp.</div>'}
          </div>
        </div>
      </div>
      <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;">
        <div style="font-weight:700;color:#0f172a;margin-bottom:8px;">🤖 Hỏi đáp &amp; Tóm tắt tài liệu (Chỉ trả lời từ nội dung thực tế · Dẫn chiếu [Đoạn #])</div>
        <div style="display:flex;gap:8px;margin-bottom:8px;flex-wrap:wrap;">
          <input id="idpQuestionInput" type="text" placeholder="Nhập câu hỏi về tài liệu (VD: Văn bản quy định nội dung chính gì? Số tiền hoặc thời hạn là bao nhiêu?)" style="flex:1;min-width:260px;padding:8px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:0.9rem;">
          <button type="button" onclick="VietNamOcrIdpStudio.runQa()" style="padding:8px 16px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;">Hỏi AI</button>
          <button type="button" onclick="VietNamOcrIdpStudio.runQa('Tóm tắt ngắn gọn các ý chính quan trọng nhất của tài liệu kèm dẫn chiếu [Đoạn #].')" style="padding:8px 14px;border-radius:8px;border:none;background:#10b981;color:#fff;font-weight:700;cursor:pointer;">✨ Tóm tắt tài liệu</button>
        </div>
        <pre id="idpQaResult" style="white-space:pre-wrap;font-family:'Segoe UI',sans-serif;font-size:0.88rem;line-height:1.6;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px;min-height:110px;max-height:220px;overflow-y:auto;margin:0;color:#1e293b;">Nhập câu hỏi hoặc bấm "Tóm tắt tài liệu" để AI phân tích có dẫn chiếu đoạn.</pre>
      </div>`;
  }

  async function runQa(presetQuestion) {
    const inp = document.getElementById('idpQuestionInput');
    const out = document.getElementById('idpQaResult');
    const q = presetQuestion || (inp ? inp.value.trim() : '');
    if (!q || !out) return;
    out.textContent = '⏳ Đang phân tích tài liệu và trích dẫn căn cứ [Đoạn #]...';
    const text = (document.getElementById('ocrResultText')?.value || '').trim();
    out.textContent = await askGroundedAi(q, text);
  }

  function renderArchivePane() {
    const pane = document.getElementById('idpPane2');
    if (!pane) return;
    pane.innerHTML = `
      <div style="display:flex;gap:8px;margin-bottom:12px;">
        <input id="idpArchiveSearchInput" type="text" oninput="VietNamOcrIdpStudio.renderArchiveList(this.value)" placeholder="🔍 Tìm theo tên file, từ khóa nội dung, ngày tháng, số văn bản, họ tên..." style="flex:1;padding:9px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:0.9rem;">
      </div>
      <div id="idpArchiveListHost" style="display:flex;flex-direction:column;gap:8px;"></div>`;
    renderArchiveList('');
  }

  function renderArchiveList(query) {
    const host = document.getElementById('idpArchiveListHost');
    if (!host) return;
    const q = String(query || '').trim().toLowerCase();
    const list = loadArchive().filter(item => {
      if (!q) return true;
      return (item.fileName || '').toLowerCase().includes(q) ||
        (item.text || '').toLowerCase().includes(q) ||
        (item.docType || '').toLowerCase().includes(q) ||
        (item.docNumbers || []).join(' ').toLowerCase().includes(q) ||
        (item.dates || []).join(' ').toLowerCase().includes(q) ||
        (item.names || []).join(' ').toLowerCase().includes(q);
    });

    if (list.length === 0) {
      host.innerHTML = '<div style="padding:18px;text-align:center;color:#64748b;background:#fff;border:1px solid #e2e8f0;border-radius:10px;">Chưa có tài liệu nào khớp điều kiện tìm kiếm. Hãy chạy OCR để hệ thống tự động lưu cục bộ!</div>';
      return;
    }

    host.innerHTML = list.map(it => {
      let snippet = (it.text || '').slice(0, 150);
      if (q && (it.text || '').toLowerCase().includes(q)) {
        const pos = it.text.toLowerCase().indexOf(q);
        snippet = '...' + it.text.slice(Math.max(0, pos - 45), pos + 110) + '...';
      }
      return `
        <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 14px;display:flex;justify-content:space-between;align-items:center;gap:12px;">
          <div style="flex:1;">
            <div style="font-weight:700;color:#1d4ed8;font-size:0.92rem;">📄 ${escHtml(it.fileName)} <span style="font-weight:500;color:#64748b;font-size:0.78rem;">(${escHtml(it.savedAt)} · ${it.wordCount} từ · ${it.confidence || 0}%)</span></div>
            <div style="font-size:0.82rem;color:#475569;margin-top:3px;">${escHtml(snippet)}</div>
          </div>
          <button type="button" onclick="VietNamOcrIdpStudio.openArchivedDoc('${it.id}')" style="padding:7px 14px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;white-space:nowrap;">Mở lại</button>
        </div>`;
    }).join('');
  }

  function openArchivedDoc(id) {
    const found = loadArchive().find(x => x.id === id);
    if (!found) return;
    const textEl = document.getElementById('ocrResultText');
    const paperEl = document.getElementById('ocrFormattedPaper');
    if (textEl) textEl.value = found.text;
    if (paperEl && typeof global.reconstructExactLayoutFromPlainText === 'function') {
      paperEl.innerHTML = global.reconstructExactLayoutFromPlainText(found.text);
    }
    close();
    if (global.showToast) global.showToast(`✓ Đã mở lại tài liệu "${found.fileName}" từ kho lưu trữ cục bộ!`);
  }

  function renderDiffPane() {
    const pane = document.getElementById('idpPane3');
    if (!pane) return;
    const curText = (document.getElementById('ocrResultText')?.value || '').trim();
    pane.innerHTML = `
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:10px;">
        <div>
          <label style="font-size:0.84rem;font-weight:700;color:#334155;">Tài liệu A (Hiện tại):</label>
          <textarea id="idpDiffA" style="width:100%;height:160px;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:0.86rem;">${escHtml(curText)}</textarea>
        </div>
        <div>
          <label style="font-size:0.84rem;font-weight:700;color:#334155;">Tài liệu B (Dán nội dung cần đối chiếu):</label>
          <textarea id="idpDiffB" placeholder="Dán văn bản thứ 2 vào đây để so sánh từng đoạn..." style="width:100%;height:160px;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:0.86rem;"></textarea>
        </div>
      </div>
      <button type="button" onclick="VietNamOcrIdpStudio.runDiff()" style="padding:8px 18px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;margin-bottom:10px;">⚖️ So sánh 2 tài liệu ngay</button>
      <div id="idpDiffOut" style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;font-size:0.85rem;max-height:200px;overflow-y:auto;"></div>`;
  }

  function runDiff() {
    const a = (document.getElementById('idpDiffA')?.value || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
    const b = (document.getElementById('idpDiffB')?.value || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
    const out = document.getElementById('idpDiffOut');
    if (!out) return;
    const setA = new Set(a);
    const setB = new Set(b);
    const onlyA = a.filter(x => !setB.has(x));
    const onlyB = b.filter(x => !setA.has(x));
    out.innerHTML = `
      <div style="margin-bottom:8px;"><strong>Kết quả đối chiếu:</strong> Giống nhau <b>${a.length - onlyA.length}</b> đoạn · Chỉ có ở Tài liệu A: <b style="color:#b91c1c;">${onlyA.length}</b> · Chỉ có ở Tài liệu B: <b style="color:#15803d;">${onlyB.length}</b></div>
      ${onlyA.map(x => `<div style="background:#fef2f2;color:#991b1b;padding:4px 8px;border-radius:6px;margin-bottom:4px;">− [Chỉ ở A] ${escHtml(x)}</div>`).join('')}
      ${onlyB.map(x => `<div style="background:#f0fdf4;color:#166534;padding:4px 8px;border-radius:6px;margin-bottom:4px;">+ [Chỉ ở B] ${escHtml(x)}</div>`).join('')}`;
  }

  function renderSecurityPane() {
    const pane = document.getElementById('idpPane4');
    if (!pane) return;
    const logs = global.VietNamOcrEnginePro ? global.VietNamOcrEnginePro.getErrorLogs() : [];
    const count = loadArchive().length;
    pane.innerHTML = `
      <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-bottom:12px;">
        <div style="font-weight:700;color:#0f172a;margin-bottom:8px;">🛡 Bảo mật &amp; Quản lý dữ liệu cục bộ (${count} tài liệu đang lưu trên máy này)</div>
        <div style="font-size:0.86rem;color:#475569;margin-bottom:12px;">Mọi thao tác OCR, PDF 2 lớp và lưu trữ đều chạy trực tiếp trên máy tính của bạn — KHÔNG gửi tài liệu lên máy chủ bên ngoài.</div>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          <button type="button" onclick="VietNamOcrIdpStudio.exportBackupArchive()" style="padding:8px 14px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;">💾 Sao lưu kho OCR (.JSON)</button>
          <label style="padding:8px 14px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">📂 Khôi phục bản sao lưu<input type="file" accept=".json" onchange="VietNamOcrIdpStudio.importBackupArchive(this.files[0])" style="display:none;"></label>
          <button type="button" onclick="VietNamOcrIdpStudio.clearAllLocalOcrData(); VietNamOcrIdpStudio.renderSecurityPane();" style="padding:8px 14px;border-radius:8px;border:none;background:#dc2626;color:#fff;font-weight:700;cursor:pointer;">🗑 Xóa sạch dữ liệu OCR &amp; Bộ nhớ tạm</button>
        </div>
      </div>
      <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;">
        <div style="font-weight:700;color:#0f172a;margin-bottom:6px;">📋 Nhật ký lỗi kỹ thuật (Không lưu nội dung văn bản nhạy cảm)</div>
        <div style="font-size:0.82rem;color:#475569;max-height:150px;overflow-y:auto;">
          ${logs.length ? logs.map(l => `<div>[${escHtml(l.time)}] <b>${escHtml(l.stage)}</b>: ${escHtml(l.message)}</div>`).join('') : '<em>Chưa ghi nhận lỗi kỹ thuật nào.</em>'}
        </div>
      </div>`;
  }

  function open(tabNum = 1) {
    const m = ensureIdpModal();
    m.style.display = 'flex';
    switchTab(tabNum);
  }

  function close() {
    const m = document.getElementById('ocrIdpStudioModal');
    if (m) m.style.display = 'none';
  }

  global.VietNamOcrIdpStudio = {
    open,
    close,
    switchTab,
    runQa,
    runDiff,
    renderArchiveList,
    renderSecurityPane,
    openArchivedDoc,
    analyzeDocumentStructure,
    autoSaveDocumentToArchive,
    exportOcrToDocx,
    exportOcrToMarkdown,
    exportOcrToStructuredJson,
    exportBackupArchive,
    importBackupArchive,
    clearAllLocalOcrData
  };
})(window);
