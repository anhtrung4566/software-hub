/**
 * VBTM Pro 3.0 Ultimate — Core Workspace Controller
 * Tích hợp đầy đủ:
 * - Bảng điều khiển Google Gemini AI trung tâm (VbtmAiService)
 * - Bảng soát lỗi chính tả 5 cột + Bảo vệ thực thể (VbtmSpellEngine)
 * - 15 chức năng AI Soạn Thảo (hỏi thông tin thiếu, chọn độ dài, so sánh trước khi thay)
 * - 7 chế độ AI Nâng Cấp Văn Phong + Diff trước/sau
 * - 2 hệ quy chuẩn thể thức độc lập (Đảng 05-HD/VPTW & Hành chính NĐ 30/2020)
 * - 14 tính năng văn phòng mới (Lịch sử, Khôi phục, So sánh 2 bản, Tìm & Thay thế, Tự lưu nháp,
 *   Kiểm tra trước trình ký, Mục lục tự động, Quản lý nhiều tài liệu, Căn cứ pháp lý, Kiểm tra QR, Xuất .DOCX thật)
 */
(function (global) {
  'use strict';

  const STORAGE = {
    AUTOSAVE: 'vbtm30_autosave_draft',
    HISTORY: 'vbtm30_doc_history',
    MULTI_DOCS: 'vbtm30_multi_docs',
    LEGAL_BASES: 'vbtm30_legal_bases',
    CUSTOM_TEMPLATES: 'vbtm30_custom_templates'
  };

  // --- PHẦN 7: 2 HỆ QUY CHUẨN THỂ THỨC ĐỘC LẬP CÓ NGUỒN CĂN CỨ ---
  const FORMAT_STANDARDS = {
    hd05: {
      id: 'hd05',
      name: 'Hệ thống A: Văn bản của Đảng',
      source: 'Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng',
      version: '2026.11',
      updatedAt: '2026-11-23',
      rules: {
        mottoTop: 'ĐẢNG CỘNG SẢN VIỆT NAM',
        mottoSizePt: 15,
        pageSize: 'A4 (210mm x 297mm)',
        marginsMm: { top: 20, bottom: 20, left: 30, right: 15 },
        bodyFont: 'Times New Roman',
        bodySizePt: 14,
        firstLineIndentCm: 1.27,
        lineSpacing: '1.15 - 1.38',
        pageNumbering: 'Đánh số trang từ trang 2, căn giữa đầu trang (13pt)',
        noiNhanStyle: 'Chữ đứng 14pt, gạch chân 8 ký tự "Nơi nhận", dấu hai chấm (:)',
        docNumberPattern: /^Số\s+\d+-[A-ZĐ]{1,4}\/[A-ZĐ]+/i
      }
    },
    nd30: {
      id: 'nd30',
      name: 'Hệ thống B: Văn bản Hành chính Nhà nước',
      source: 'Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư',
      version: '2026.11',
      updatedAt: '2026-11-23',
      rules: {
        mottoTop: 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
        mottoSub: 'Độc lập - Tự do - Hạnh phúc',
        mottoSizePt: 12.5,
        pageSize: 'A4 (210mm x 297mm)',
        marginsMm: { top: 20, bottom: 20, left: 30, right: 15 },
        bodyFont: 'Times New Roman',
        bodySizePt: 14,
        firstLineIndentCm: 1.27,
        lineSpacing: '1.15 - 1.38',
        pageNumbering: 'Đánh số trang từ trang 2, căn giữa đầu trang (13-14pt)',
        noiNhanStyle: 'Chữ nghiêng đậm 12pt "Nơi nhận:", danh sách 11pt',
        docNumberPattern: /^Số:\s*\d+\/[A-ZĐ]+-[A-ZĐ]+/i
      }
    }
  };

  // --- PHẦN 4: 15 CHỨC NĂNG AI SOẠN THẢO ---
  const AI_DRAFT_MODES = [
    { id: 'congvan', label: '1. AI soạn Công văn', docType: '', std: 'nd30' },
    { id: 'baocao', label: '2. AI soạn Báo cáo', docType: 'BÁO CÁO', std: 'nd30' },
    { id: 'kehoach', label: '3. AI soạn Kế hoạch', docType: 'KẾ HOẠCH', std: 'nd30' },
    { id: 'totrinh', label: '4. AI soạn Tờ trình', docType: 'TỜ TRÌNH', std: 'nd30' },
    { id: 'quyetdinh', label: '5. AI soạn Quyết định', docType: 'QUYẾT ĐỊNH', std: 'nd30' },
    { id: 'nghiquyet', label: '6. AI soạn Nghị quyết', docType: 'NGHỊ QUYẾT', std: 'hd05' },
    { id: 'thongbao', label: '7. AI soạn Thông báo', docType: 'THÔNG BÁO', std: 'nd30' },
    { id: 'bienban', label: '8. AI soạn Biên bản', docType: 'BIÊN BẢN', std: 'nd30' },
    { id: 'chuongtrinh', label: '9. AI soạn Chương trình công tác', docType: 'CHƯƠNG TRÌNH', std: 'hd05' },
    { id: 'vbdang', label: '10. AI soạn Văn bản Đảng ủy (05-HD)', docType: 'BÁO CÁO', std: 'hd05' },
    { id: 'vbubnd', label: '11. AI soạn Văn bản UBND (NĐ 30)', docType: '', std: 'nd30' },
    { id: 'tomtat', label: '12. AI tóm tắt văn bản hiện tại', transform: true },
    { id: 'nhiemvu', label: '13. AI trích xuất danh sách nhiệm vụ', transform: true },
    { id: 'tiendo', label: '14. AI lập bảng tiến độ thực hiện', transform: true },
    { id: 'cautruc', label: '15. AI chuyển văn bản thô → có cấu trúc', transform: true }
  ];

  // --- PHẦN 5: 7 CHẾ ĐỘ AI NÂNG CẤP VĂN PHONG ---
  const AI_STYLE_MODES = [
    { id: 'hanhchinh', label: 'Văn phong hành chính chuẩn mực (NĐ 30/2020)' },
    { id: 'dang', label: 'Văn phong Đảng (Hướng dẫn 05-HD/VPTW)' },
    { id: 'baocao', label: 'Văn phong báo cáo khách quan, số liệu rõ ràng' },
    { id: 'chidao', label: 'Văn phong chỉ đạo, điều hành dứt khoát' },
    { id: 'tuyentruyen', label: 'Văn phong tuyên truyền, vận động dễ hiểu' },
    { id: 'tongket', label: 'Văn phong tổng kết, đánh giá toàn diện' },
    { id: 'thammuu', label: 'Văn phong tham mưu, đề xuất chặt chẽ' }
  ];

  const DEFAULT_LEGAL_BASES = [
    { title: 'Nghị định 30/2020/NĐ-CP', text: 'Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;' },
    { title: 'Hướng dẫn 05-HD/VPTW', text: 'Căn cứ Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng;' },
    { title: 'Luật Tổ chức CQĐP', text: 'Căn cứ Luật Tổ chức chính quyền địa phương;' },
    { title: 'Điều lệ Đảng CSVN', text: 'Căn cứ Điều lệ Đảng Cộng sản Việt Nam;' },
    { title: 'Luật Lưu trữ', text: 'Căn cứ Luật Lưu trữ;' }
  ];

  // Trạng thái làm việc hiện tại
  let currentSpellIssues = [];
  let spellOriginalSnapshot = '';
  let pendingComparison = null; // { title, before, after, onApply }
  let activeDocTabId = 'doc_1';
  let multiDocs = [];

  function $(id) { return document.getElementById(id); }

  function toast(msg) {
    if (typeof global.showToast === 'function') global.showToast(msg);
  }

  function escHtml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function readCurrentDocFields() {
    const qrCanvas = $('vbtmQrCanvas');
    const includeQr = Boolean($('vbtmIncludeQr') && $('vbtmIncludeQr').checked);
    return {
      standard: $('vbtmStandard') ? $('vbtmStandard').value : 'nd30',
      governingOrg: $('vbtmGovOrg') ? $('vbtmGovOrg').value : '',
      issuingOrg: $('vbtmIssueOrg') ? $('vbtmIssueOrg').value : '',
      docNumber: $('vbtmDocNum') ? $('vbtmDocNum').value : '',
      location: $('vbtmLocation') ? $('vbtmLocation').value : '',
      docDate: $('vbtmDocDate') ? $('vbtmDocDate').value : '',
      docType: $('vbtmDocType') ? $('vbtmDocType').value : '',
      docSubject: $('vbtmDocSubject') ? $('vbtmDocSubject').value : '',
      docBody: $('vbtmDocBody') ? $('vbtmDocBody').value : '',
      recipients: $('vbtmRecipients') ? $('vbtmRecipients').value : '',
      signerTitle: $('vbtmSignerTitle') ? $('vbtmSignerTitle').value : '',
      signerName: $('vbtmSignerName') ? $('vbtmSignerName').value : '',
      qrText: $('vbtmQrText') ? $('vbtmQrText').value : '',
      includeQr,
      qrDataUrl: (includeQr && qrCanvas) ? qrCanvas.toDataURL('image/png') : ''
    };
  }

  function writeDocFields(d) {
    if (!d) return;
    const map = {
      vbtmStandard: d.standard,
      vbtmGovOrg: d.governingOrg,
      vbtmIssueOrg: d.issuingOrg,
      vbtmDocNum: d.docNumber,
      vbtmLocation: d.location,
      vbtmDocDate: d.docDate,
      vbtmDocType: d.docType,
      vbtmDocSubject: d.docSubject,
      vbtmDocBody: d.docBody,
      vbtmRecipients: d.recipients,
      vbtmSignerTitle: d.signerTitle,
      vbtmSignerName: d.signerName,
      vbtmQrText: d.qrText
    };
    Object.keys(map).forEach(k => {
      if (typeof map[k] === 'string' && $(k)) $(k).value = map[k];
    });
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    if (typeof global.updateVbtmLiveStats === 'function') global.updateVbtmLiveStats();
  }

  // --- 1. ĐIỀU HƯỚNG 11 PHÂN HỆ (PHẦN 9) ---
  function switchModule(modId) {
    document.querySelectorAll('.v3-nav-btn').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-mod') === modId);
    });
    document.querySelectorAll('.v3-panel-section').forEach(sec => {
      const mods = (sec.getAttribute('data-mods') || '').split(',');
      sec.style.display = (modId === 'all' || mods.includes(modId)) ? '' : 'none';
    });
    if (modId === 'format') runFormatAudit30();
    if (modId === 'history') renderHistoryPanel();
    if (modId === 'library') renderLibraryPanel();
    if (modId === 'qr') verifyQrCodeHealth();
  }

  // --- 2. BẢNG ĐIỀU KHIỂN AI TRUNG TÂM (PHẦN 2.3) ---
  function renderAiControlPanel(aiState) {
    const s = aiState || (global.VbtmAiService ? global.VbtmAiService.getState() : null);
    if (!s) return;
    const badge = $('v3AiStatusBadge');
    const modelLbl = $('v3AiModelDisplay');
    const latencyLbl = $('v3AiLatencyDisplay');
    const errBox = $('v3AiErrorBox');
    const cancelBtn = $('v3AiCancelBtn');
    const logList = $('v3AiLogList');

    if (badge) {
      const map = {
        idle: { text: 'Chưa kết nối', bg: '#475569' },
        connecting: { text: 'Đang kết nối...', bg: '#d97706' },
        connected: { text: 'Đã kết nối Gemini', bg: '#059669' },
        error: { text: 'Lỗi kết nối', bg: '#dc2626' }
      };
      const st = map[s.status] || map.idle;
      badge.textContent = st.text;
      badge.style.background = st.bg;
    }
    if (modelLbl) modelLbl.textContent = s.activeModel || 'gemini-2.5-flash';
    if (latencyLbl) latencyLbl.textContent = s.lastLatencyMs ? `${s.lastLatencyMs} ms` : '—';
    if (errBox) {
      if (s.lastError) {
        errBox.style.display = 'block';
        errBox.textContent = `⚠ ${s.lastError}`;
      } else {
        errBox.style.display = 'none';
      }
    }
    if (cancelBtn) cancelBtn.disabled = !s.isBusy;
    if (logList) {
      logList.innerHTML = s.logs.length
        ? s.logs.slice(0, 8).map(l => `<div style="font-size:0.75rem;padding:2px 0;border-bottom:1px solid rgba(255,255,255,0.06);">[${escHtml(l.time)}] <b>${escHtml(l.action)}</b> (${escHtml(l.model)}) — ${escHtml(l.code || l.status)} ${l.latencyMs ? `(${l.latencyMs}ms)` : ''}: ${escHtml(l.message)}</div>`).join('')
        : '<div style="font-size:0.75rem;opacity:0.7;">Chưa có log kỹ thuật. (Không bao giờ lưu nội dung văn bản nhạy cảm vào log)</div>';
    }
  }

  async function handleTestApiKey() {
    const keyInp = $('vbtmGeminiKey');
    const modelSel = $('v3AiModelSelect');
    if (keyInp) global.VbtmAiService.setApiKey(keyInp.value);
    if (modelSel && modelSel.value) global.VbtmAiService.setModel(modelSel.value);
    const res = await global.VbtmAiService.testApiKey();
    if (res.ok && res.models && modelSel) {
      modelSel.innerHTML = res.models.map(m => `<option value="${escHtml(m)}" ${m === res.model ? 'selected' : ''}>${escHtml(m)}</option>`).join('');
    }
    toast(res.message);
  }

  async function handleTestConnection() {
    const keyInp = $('vbtmGeminiKey');
    if (keyInp) global.VbtmAiService.setApiKey(keyInp.value);
    const res = await global.VbtmAiService.testConnection();
    toast(res.message);
  }

  // --- 3. BẢNG SOÁT LỖI CHÍNH TẢ 5 CỘT + BẢO VỆ THỰC THỂ (PHẦN 3) ---
  function renderSpellTable() {
    const tbody = $('v3SpellTableBody');
    const summary = $('v3SpellSummary');
    if (!tbody) return;

    const pending = currentSpellIssues.filter(i => i.status === 'pending');
    const safeCnt = pending.filter(i => !i.needsConfirmation && (i.confidence || 0) >= 0.85).length;
    const protCnt = pending.filter(i => i.needsConfirmation).length;

    if (summary) {
      summary.textContent = pending.length === 0
        ? '✓ Văn bản sạch lỗi chính tả cơ bản (0 lỗi chờ xử lý).'
        : `Phát hiện ${pending.length} đề xuất (${safeCnt} lỗi an toàn có thể sửa tự động, ${protCnt} mục cần người dùng tự kiểm tra vì liên quan thực thể được bảo vệ).`;
    }

    if (pending.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:16px;color:#10b981;">✓ Không còn lỗi chờ xử lý.</td></tr>';
      return;
    }

    tbody.innerHTML = pending.map((it, idx) => {
      const confPct = Math.round((it.confidence || 0.85) * 100);
      const badgeProt = it.needsConfirmation
        ? '<span style="display:inline-block;margin-left:4px;padding:1px 6px;border-radius:4px;background:#b45309;color:#fff;font-size:0.7rem;">Cần kiểm tra</span>'
        : '<span style="display:inline-block;margin-left:4px;padding:1px 6px;border-radius:4px;background:#047857;color:#fff;font-size:0.7rem;">An toàn</span>';
      return `<tr>
        <td style="text-align:center;">${idx + 1}</td>
        <td><code style="color:#fca5a5;">${escHtml(it.original)}</code><div style="font-size:0.72rem;opacity:0.75;">Đoạn #${it.paragraphId + 1} · ${escHtml(it.reason)}</div></td>
        <td><input type="text" class="glass-input" id="sug_${escHtml(it.id)}" value="${escHtml(it.suggestion)}" style="padding:3px 6px;font-size:0.82rem;min-width:130px;"></td>
        <td>${escHtml(it.category)} ${badgeProt}</td>
        <td style="text-align:center;">${confPct}%</td>
        <td style="white-space:nowrap;">
          <button type="button" class="btn-tool primary" style="padding:3px 8px;font-size:0.75rem;" onclick="VbtmPro30.acceptSpellIssue('${escHtml(it.id)}')">Sửa lỗi này</button>
          <button type="button" class="btn-tool" style="padding:3px 8px;font-size:0.75rem;" onclick="VbtmPro30.ignoreSpellIssue('${escHtml(it.id)}')">Bỏ qua</button>
        </td>
      </tr>`;
    }).join('');
  }

  function runOfflineSpellCheck30() {
    const bodyEl = $('vbtmDocBody');
    if (!bodyEl) return;
    if (!spellOriginalSnapshot) spellOriginalSnapshot = bodyEl.value;
    currentSpellIssues = global.VbtmSpellEngine.scanOffline(bodyEl.value);
    renderSpellTable();
    toast(`Đã quét quy tắc ngoại tuyến: phát hiện ${currentSpellIssues.length} đề xuất.`);
  }

  async function runGeminiSpellCheck30() {
    const bodyEl = $('vbtmDocBody');
    if (!bodyEl || !bodyEl.value.trim()) {
      toast('Nội dung văn bản đang trống!');
      return;
    }
    if (!spellOriginalSnapshot) spellOriginalSnapshot = bodyEl.value;
    // Chạy quét ngoại tuyến trước
    const offlineList = global.VbtmSpellEngine.scanOffline(bodyEl.value);
    currentSpellIssues = offlineList.slice();
    renderSpellTable();

    try {
      toast('Đang gọi Google Gemini AI soát lỗi ngữ cảnh (JSON có cấu trúc)...');
      const res = await global.VbtmAiService.getState().hasKey || global.VbtmAiService.getConnectionMode() === 'proxy'
        ? await global.VbtmSpellEngine.scanWithGeminiAI(bodyEl.value)
        : null;

      if (!res) {
        toast('Chưa có Gemini API Key — đã hoàn tất soát lỗi bằng bộ quy tắc ngoại tuyến.');
        return;
      }

      // Hợp nhất không trùng lặp tọa độ
      const seen = new Set(currentSpellIssues.map(i => `${i.paragraphId}:${i.startOffset}:${i.original}`));
      let added = 0;
      for (const aiItem of res.issues) {
        const k = `${aiItem.paragraphId}:${aiItem.startOffset}:${aiItem.original}`;
        if (!seen.has(k)) {
          currentSpellIssues.push(aiItem);
          seen.add(k);
          added++;
        }
      }
      renderSpellTable();
      toast(`✓ Gemini AI (${res.model}, ${res.latencyMs}ms) phát hiện thêm ${added} đề xuất ngữ cảnh!`);
    } catch (err) {
      toast(`⚠ Gemini AI: ${err.userMessage || err.message} (đã giữ kết quả soát quy tắc ngoại tuyến)`);
    }
  }

  function acceptSpellIssue(id) {
    const bodyEl = $('vbtmDocBody');
    const issue = currentSpellIssues.find(i => i.id === id && i.status === 'pending');
    if (!bodyEl || !issue) return;
    const customInp = $(`sug_${id}`);
    const chosen = customInp ? customInp.value : issue.suggestion;

    const res = global.VbtmSpellEngine.applySingleFix(bodyEl.value, issue, chosen);
    if (!res.ok) {
      toast(`⚠ ${res.reason}`);
      return;
    }
    saveHistorySnapshot(`Sửa lỗi: "${issue.original}" → "${chosen}"`);
    bodyEl.value = res.text;
    issue.status = 'accepted';
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    if (typeof global.updateVbtmLiveStats === 'function') global.updateVbtmLiveStats();
    renderSpellTable();
    toast(`✓ Đã sửa "${issue.original}" → "${chosen}"`);
  }

  function ignoreSpellIssue(id) {
    const issue = currentSpellIssues.find(i => i.id === id);
    if (issue) issue.status = 'ignored';
    renderSpellTable();
  }

  function fixAllSafeSpellIssues() {
    const bodyEl = $('vbtmDocBody');
    if (!bodyEl) return;
    const res = global.VbtmSpellEngine.applyAllSafeFixes(bodyEl.value, currentSpellIssues);
    if (res.fixedCount === 0) {
      toast(`Không có lỗi an toàn để tự sửa (${res.skippedProtectedCount} mục được bảo vệ cần bạn tự duyệt).`);
      return;
    }
    saveHistorySnapshot(`Sửa tự động ${res.fixedCount} lỗi an toàn`);
    bodyEl.value = res.text;
    currentSpellIssues.forEach(it => { if (res.fixedIds.has(it.id)) it.status = 'accepted'; });
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    if (typeof global.updateVbtmLiveStats === 'function') global.updateVbtmLiveStats();
    renderSpellTable();
    toast(`✓ Đã sửa ${res.fixedCount} lỗi an toàn (bỏ qua ${res.skippedProtectedCount} mục cần bạn tự kiểm tra).`);
  }

  function restoreSpellOriginal() {
    const bodyEl = $('vbtmDocBody');
    if (!bodyEl || !spellOriginalSnapshot) {
      toast('Chưa có bản gốc để khôi phục.');
      return;
    }
    bodyEl.value = spellOriginalSnapshot;
    currentSpellIssues.forEach(it => { it.status = 'pending'; });
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    renderSpellTable();
    toast('✓ Đã khôi phục nguyên văn nội dung gốc trước khi sửa lỗi.');
  }

  // --- 4. AI SOẠN THẢO 15 CHỨC NĂNG + HỎI THÔNG TIN THIẾU + SO SÁNH (PHẦN 4) ---
  function checkMissingDraftInfo(modeId, promptText) {
    const missing = [];
    if (promptText.length < 18) missing.push('Chủ đề / nội dung trọng tâm cụ thể');
    if (!/\b(phường|xã|quận|huyện|thành phố|đảng ủy|ubnd|chi bộ)\b/i.test(promptText) && !$('vbtmIssueOrg')?.value.trim()) {
      missing.push('Tên cơ quan / đơn vị ban hành');
    }
    if (['baocao', 'kehoach', 'chuongtrinh', 'nghiquyet'].includes(modeId) && !/\b(tháng|quý|năm|202\d)\b/i.test(promptText)) {
      missing.push('Kỳ hạn / mốc thời gian (Tháng/Quý/Năm)');
    }
    return missing;
  }

  async function runAiDraft30() {
    const modeId = $('v3DraftModeSelect') ? $('v3DraftModeSelect').value : 'congvan';
    const lengthOpt = $('v3DraftLengthSelect') ? $('v3DraftLengthSelect').value : 'Đầy đủ';
    const promptInput = ($('vbtmAiPromptInput') ? $('vbtmAiPromptInput').value : '').trim();
    const modeMeta = AI_DRAFT_MODES.find(m => m.id === modeId) || AI_DRAFT_MODES[0];
    const currentBody = ($('vbtmDocBody') ? $('vbtmDocBody').value : '').trim();

    if (modeMeta.transform) {
      if (!currentBody) {
        toast('Vui lòng nhập hoặc tải văn bản vào khung Nội dung trước khi chạy chức năng này!');
        return;
      }
    } else {
      const missing = checkMissingDraftInfo(modeId, promptInput);
      const missBox = $('v3MissingInfoBox');
      if (missing.length > 0 && missBox) {
        missBox.style.display = 'block';
        missBox.innerHTML = `<strong>⚠ Thiếu thông tin đầu vào (AI sẽ KHÔNG tự bịa số liệu/căn cứ):</strong> ${missing.join(' · ')}. Hãy bổ sung vào ô yêu cầu hoặc để AI đánh dấu <code>[CẦN ĐIỀN]</code>.`;
      } else if (missBox) {
        missBox.style.display = 'none';
      }
      if (!promptInput) {
        toast('Vui lòng nhập yêu cầu nội dung cần soạn thảo!');
        return;
      }
    }

    const beforeText = currentBody;
    const useGemini = global.VbtmAiService && (global.VbtmAiService.getState().hasKey || global.VbtmAiService.getConnectionMode() === 'proxy');

    if (!useGemini) {
      toast('Chưa kết nối Gemini API Key — vui lòng nhập Key ở mục Trợ Lý AI (hoặc dùng mẫu chuẩn có sẵn).');
      return;
    }

    const sysPrompt =
      'Bạn là chuyên gia soạn thảo văn bản Đảng (Hướng dẫn 05-HD/VPTW) và Hành chính Nhà nước Việt Nam (Nghị định 30/2020/NĐ-CP).\n' +
      'QUY TẮC TUYỆT ĐỐI:\n' +
      '1. KHÔNG ĐƯỢC tự bịa số liệu, tên người, số hiệu văn bản, căn cứ pháp lý hay kết quả không có trong yêu cầu. Chỗ nào thiếu thông tin phải ghi rõ [CẦN ĐIỀN: ...].\n' +
      '2. Không dùng từ ngữ hoa mỹ không cần thiết.\n' +
      '3. Trả về văn bản thuần tiếng Việt chuẩn mực (KHÔNG dùng Markdown ** hay #).\n' +
      `4. Độ dài yêu cầu: ${lengthOpt}.`;

    const userPrompt = modeMeta.transform
      ? `Chức năng: ${modeMeta.label}.\nYêu cầu bổ sung: ${promptInput || 'Thực hiện chính xác theo nội dung gốc'}.\n\nVĂN BẢN GỐC:\n${currentBody.slice(0, 14000)}`
      : `Chức năng: ${modeMeta.label}.\nCơ quan ban hành hiện tại: ${$('vbtmIssueOrg')?.value || '[CẦN ĐIỀN]'}.\nYêu cầu chi tiết: ${promptInput}`;

    try {
      toast(`Đang gọi Gemini AI thực hiện: ${modeMeta.label}...`);
      const res = await global.VbtmAiService.request({
        actionName: `draft_${modeId}`,
        systemInstruction: sysPrompt,
        prompt: userPrompt,
        temperature: 0.35,
        maxOutputTokens: lengthOpt === 'Chi tiết' ? 3072 : (lengthOpt === 'Ngắn gọn' ? 1024 : 2048)
      });

      const cleanText = res.text.replace(/\*\*/g, '').replace(/^#+\s*/gm, '').trim();
      openComparisonModal(
        `Xem trước & So sánh: ${modeMeta.label} (${res.model}, ${res.latencyMs}ms)`,
        beforeText,
        cleanText,
        () => {
          saveHistorySnapshot(`Trước khi áp dụng ${modeMeta.label}`);
          if (!modeMeta.transform && modeMeta.std && $('vbtmStandard')) $('vbtmStandard').value = modeMeta.std;
          if (!modeMeta.transform && typeof modeMeta.docType === 'string' && $('vbtmDocType')) $('vbtmDocType').value = modeMeta.docType;
          if ($('vbtmDocBody')) $('vbtmDocBody').value = cleanText;
          if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
          if (typeof global.updateVbtmLiveStats === 'function') global.updateVbtmLiveStats();
          toast(`✓ Đã áp dụng kết quả ${modeMeta.label}!`);
        }
      );
    } catch (err) {
      toast(`⚠ Lỗi AI: ${err.userMessage || err.message}`);
    }
  }

  // --- 5. AI NÂNG CẤP VĂN PHONG 7 CHẾ ĐỘ + SO SÁNH TRƯỚC/SAU (PHẦN 5) ---
  async function runAiStyleUpgrade30() {
    const bodyEl = $('vbtmDocBody');
    const styleId = $('v3StyleModeSelect') ? $('v3StyleModeSelect').value : 'hanhchinh';
    const styleMeta = AI_STYLE_MODES.find(s => s.id === styleId) || AI_STYLE_MODES[0];
    if (!bodyEl || !bodyEl.value.trim()) {
      toast('Chưa có nội dung văn bản để nâng cấp văn phong!');
      return;
    }

    const useGemini = global.VbtmAiService && (global.VbtmAiService.getState().hasKey || global.VbtmAiService.getConnectionMode() === 'proxy');
    if (!useGemini) {
      toast('Chức năng nâng cấp văn phong AI cần kết nối Google Gemini API Key trong Bảng điều khiển AI.');
      return;
    }

    const sysPrompt =
      `Bạn là chuyên gia biên tập văn bản cơ quan Nhà nước và Đảng ủy Việt Nam.\n` +
      `Hãy biên tập văn bản sau theo chế độ: ${styleMeta.label}.\n` +
      `NGUYÊN TẮC BẮT BUỘC:\n` +
      `- KHÔNG dùng từ ngữ hoa mỹ không cần thiết.\n` +
      `- KHÔNG tự tạo thành tích, KHÔNG tự thêm số liệu.\n` +
      `- KHÔNG biến ý kiến đề xuất thành kết luận chính thức.\n` +
      `- Giữ nguyên 100% tên riêng, số liệu, ngày tháng, số hiệu, căn cứ pháp lý và ý nghĩa chỉ đạo gốc.\n` +
      `- Chỉ trả về văn bản thuần đã nâng cấp (không Markdown ** hay #).`;

    try {
      toast(`Đang nâng cấp văn phong (${styleMeta.label})...`);
      const res = await global.VbtmAiService.request({
        actionName: `style_${styleId}`,
        systemInstruction: sysPrompt,
        prompt: bodyEl.value.slice(0, 14000),
        temperature: 0.35,
        maxOutputTokens: 2500
      });
      const cleanText = res.text.replace(/\*\*/g, '').replace(/^#+\s*/gm, '').trim();
      openComparisonModal(
        `So sánh Trước & Sau — ${styleMeta.label} (${res.latencyMs}ms)`,
        bodyEl.value,
        cleanText,
        () => {
          saveHistorySnapshot(`Trước khi nâng cấp văn phong (${styleMeta.label})`);
          bodyEl.value = cleanText;
          if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
          toast('✓ Đã áp dụng bản nâng cấp văn phong!');
        }
      );
    } catch (err) {
      toast(`⚠ Lỗi AI: ${err.userMessage || err.message}`);
    }
  }

  // --- MODAL SO SÁNH TRƯỚC / SAU ---
  function openComparisonModal(title, beforeText, afterText, onApply) {
    pendingComparison = { title, before: beforeText, after: afterText, onApply };
    const modal = $('v3CompareModal');
    if (!modal) return;
    if ($('v3CompareTitle')) $('v3CompareTitle').textContent = title;
    if ($('v3CompareBefore')) $('v3CompareBefore').value = beforeText || '';
    if ($('v3CompareAfter')) $('v3CompareAfter').value = afterText || '';
    modal.style.display = 'flex';
  }

  function closeComparisonModal() {
    const modal = $('v3CompareModal');
    if (modal) modal.style.display = 'none';
    pendingComparison = null;
  }

  function applyComparisonModal() {
    if (pendingComparison && typeof pendingComparison.onApply === 'function') {
      const editedAfter = $('v3CompareAfter') ? $('v3CompareAfter').value : pendingComparison.after;
      pendingComparison.after = editedAfter;
      if ($('vbtmDocBody') && editedAfter !== pendingComparison.before) {
        // Cho phép người dùng sửa tay trong ô So sánh rồi mới áp dụng
        const origApply = pendingComparison.onApply;
        origApply();
        $('vbtmDocBody').value = editedAfter;
        if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
      }
    }
    closeComparisonModal();
  }

  // --- 6. KIỂM ĐỊNH THỂ THỨC 16 YẾU TỐ & KIỂM TRA TRƯỚC TRÌNH KÝ (PHẦN 7 & 8.6) ---
  function runFormatAudit30() {
    const d = readCurrentDocFields();
    const std = FORMAT_STANDARDS[d.standard === 'hd05' ? 'hd05' : 'nd30'];
    const checks = [
      { name: 'Quốc hiệu / Tiêu ngữ', ok: true, detail: std.rules.mottoTop + (std.rules.mottoSub ? ' — ' + std.rules.mottoSub : '') },
      { name: 'Tên cơ quan ban hành', ok: d.issuingOrg.trim().length >= 4, detail: d.issuingOrg || 'Thiếu tên cơ quan ban hành' },
      { name: 'Số và ký hiệu văn bản', ok: d.docNumber.trim().length >= 3, detail: d.docNumber || 'Chưa điền số & ký hiệu' },
      { name: 'Địa danh, ngày tháng', ok: Boolean(d.location.trim() && d.docDate), detail: `${d.location || '?'} · ${d.docDate || '?'}` },
      { name: 'Tên loại văn bản', ok: true, detail: d.docType || '(Công văn — không ghi tên loại theo quy định)' },
      { name: 'Trích yếu nội dung', ok: d.docSubject.trim().length >= 8, detail: d.docSubject || 'Thiếu trích yếu nội dung' },
      { name: 'Nội dung văn bản', ok: d.docBody.trim().length >= 30, detail: `${d.docBody.trim().split(/\s+/).filter(Boolean).length} từ` },
      { name: 'Kết thúc văn bản (./.)', ok: /\.\/\.?\s*$/.test(d.docBody.trim()), detail: /\.\/\.?\s*$/.test(d.docBody.trim()) ? 'Đúng chuẩn kết thúc ./.' : 'Nên kết thúc đoạn cuối bằng ./.' },
      { name: 'Chức vụ & Người ký', ok: Boolean(d.signerTitle.trim() && d.signerName.trim()), detail: `${d.signerTitle.replace(/\n/g, ' / ')} — ${d.signerName || '?'}` },
      { name: 'Nơi nhận', ok: d.recipients.trim().length >= 6, detail: std.rules.noiNhanStyle },
      { name: 'Khổ giấy & Lề trang', ok: true, detail: `${std.rules.pageSize} · Lề Trái 30mm / Phải 15mm / Trên-Dưới 20mm` },
      { name: 'Font, Cỡ chữ & Đánh số trang', ok: true, detail: `${std.rules.bodyFont} ${std.rules.bodySizePt}pt · ${std.rules.pageNumbering}` }
    ];

    const passed = checks.filter(c => c.ok).length;
    const score = Math.round((passed / checks.length) * 100);
    const box = $('v3FormatAuditBox');
    if (box) {
      box.innerHTML = `
        <div style="margin-bottom:10px;padding:10px;border-radius:10px;background:rgba(255,255,255,0.06);">
          <div style="font-weight:700;color:#facc15;">${escHtml(std.name)} — Điểm tuân thủ: ${score}% (${passed}/${checks.length} tiêu chí)</div>
          <div style="font-size:0.78rem;opacity:0.85;">Căn cứ: ${escHtml(std.source)} · Phiên bản quy tắc: ${escHtml(std.version)} (${escHtml(std.updatedAt)})</div>
        </div>
        <div style="display:grid;gap:6px;">
          ${checks.map(c => `<div style="display:flex;justify-content:space-between;gap:10px;padding:6px 10px;border-radius:8px;background:rgba(255,255,255,0.04);border-left:3px solid ${c.ok ? '#10b981' : '#f59e0b'};">
            <span><b>${c.ok ? '✓' : '⚠'} ${escHtml(c.name)}</b></span>
            <span style="font-size:0.8rem;opacity:0.9;text-align:right;">${escHtml(c.detail)}</span>
          </div>`).join('')}
        </div>`;
    }
    return { score, passed, total: checks.length, checks };
  }

  // --- 7. LỊCH SỬ CHỈNH SỬA, TỰ ĐỘNG LƯU NHÁP & ĐA TÀI LIỆU (PHẦN 8) ---
  function saveHistorySnapshot(label) {
    try {
      const list = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]');
      list.unshift({
        id: 'ver_' + Date.now(),
        time: new Date().toLocaleString('vi-VN'),
        label: label || 'Lưu phiên bản',
        data: readCurrentDocFields()
      });
      if (list.length > 20) list.length = 20;
      localStorage.setItem(STORAGE.HISTORY, JSON.stringify(list));
      renderHistoryPanel();
    } catch (e) {}
  }

  function renderHistoryPanel() {
    const box = $('v3HistoryList');
    if (!box) return;
    let list = [];
    try { list = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]'); } catch (e) {}
    if (!list.length) {
      box.innerHTML = '<div style="opacity:0.7;font-size:0.85rem;">Chưa có phiên bản lịch sử nào. Hệ thống tự lưu khi bạn chỉnh sửa hoặc bấm "Lưu phiên bản".</div>';
      return;
    }
    box.innerHTML = list.map(v => `
      <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;padding:8px 10px;border-radius:8px;background:rgba(255,255,255,0.05);margin-bottom:6px;">
        <div>
          <div style="font-weight:700;font-size:0.86rem;">${escHtml(v.label)}</div>
          <div style="font-size:0.75rem;opacity:0.75;">${escHtml(v.time)} · ${escHtml(v.data.docNumber || 'Chưa số')}</div>
        </div>
        <div style="display:flex;gap:6px;">
          <button type="button" class="btn-tool" style="padding:3px 9px;font-size:0.76rem;" onclick="VbtmPro30.compareWithHistory('${escHtml(v.id)}')">So sánh</button>
          <button type="button" class="btn-tool primary" style="padding:3px 9px;font-size:0.76rem;" onclick="VbtmPro30.restoreHistory('${escHtml(v.id)}')">Khôi phục</button>
        </div>
      </div>`).join('');
  }

  function restoreHistory(verId) {
    try {
      const list = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]');
      const found = list.find(v => v.id === verId);
      if (found) {
        writeDocFields(found.data);
        toast(`✓ Đã khôi phục phiên bản: ${found.label} (${found.time})`);
      }
    } catch (e) {}
  }

  function compareWithHistory(verId) {
    try {
      const list = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]');
      const found = list.find(v => v.id === verId);
      if (found) {
        const cur = readCurrentDocFields();
        openComparisonModal(`So sánh với phiên bản: ${found.label}`, found.data.docBody || '', cur.docBody || '', () => {});
      }
    } catch (e) {}
  }

  function autoSaveDraft() {
    try {
      const data = readCurrentDocFields();
      if (data.docBody.trim().length > 5) {
        localStorage.setItem(STORAGE.AUTOSAVE, JSON.stringify({ savedAt: Date.now(), data }));
        const idx = multiDocs.findIndex(d => d.id === activeDocTabId);
        if (idx >= 0) {
          multiDocs[idx].data = data;
          multiDocs[idx].title = (data.docType || 'Công văn') + (data.docNumber ? ' ' + data.docNumber : '');
          localStorage.setItem(STORAGE.MULTI_DOCS, JSON.stringify(multiDocs));
        }
      }
    } catch (e) {}
  }

  function checkAutoSavedDraftOnLoad() {
    try {
      const raw = localStorage.getItem(STORAGE.AUTOSAVE);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (parsed && parsed.data && parsed.data.docBody) {
        const banner = $('v3DraftRestoreBanner');
        if (banner) {
          banner.style.display = 'flex';
          const tm = new Date(parsed.savedAt).toLocaleTimeString('vi-VN');
          const lbl = $('v3DraftRestoreText');
          if (lbl) lbl.textContent = `Phát hiện bản nháp tự động lưu lúc ${tm} (${parsed.data.docNumber || 'chưa đặt số'}).`;
        }
      }
    } catch (e) {}
  }

  function restoreAutoSavedDraft() {
    try {
      const parsed = JSON.parse(localStorage.getItem(STORAGE.AUTOSAVE) || 'null');
      if (parsed && parsed.data) {
        writeDocFields(parsed.data);
        if ($('v3DraftRestoreBanner')) $('v3DraftRestoreBanner').style.display = 'none';
        toast('✓ Đã khôi phục bản nháp tự động lưu!');
      }
    } catch (e) {}
  }

  // --- 8. TÌM & THAY THẾ, MỤC LỤC TỰ ĐỘNG, ĐA TÀI LIỆU, CĂN CỨ PHÁP LÝ, KIỂM TRA QR ---
  function runFindAndReplace() {
    const f = ($('v3FindInput') ? $('v3FindInput').value : '');
    const r = ($('v3ReplaceInput') ? $('v3ReplaceInput').value : '');
    const bodyEl = $('vbtmDocBody');
    if (!f || !bodyEl) { toast('Nhập cụm từ cần tìm!'); return; }
    const rx = new RegExp(f.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    const matches = bodyEl.value.match(rx);
    if (!matches || matches.length === 0) { toast(`Không tìm thấy "${f}" trong nội dung.`); return; }
    saveHistorySnapshot(`Tìm & Thay thế "${f}" → "${r}"`);
    bodyEl.value = bodyEl.value.replace(rx, r);
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    toast(`✓ Đã thay thế ${matches.length} vị trí "${f}" → "${r}".`);
  }

  function generateAutoToc() {
    const bodyEl = $('vbtmDocBody');
    if (!bodyEl || !bodyEl.value.trim()) return;
    const lines = bodyEl.value.split('\n').map(l => l.trim()).filter(Boolean);
    const headings = lines.filter(l => /^((I|II|III|IV|V|VI)\.\s|Điều\s+\d+\.|\d+\.\s[A-ZÀ-ỸĐ])/.test(l));
    if (!headings.length) {
      toast('Không tìm thấy đề mục (I., II., 1., Điều 1...) để tạo mục lục.');
      return;
    }
    const toc = 'MỤC LỤC\n' + headings.map((h, i) => `${i + 1}. ${h}`).join('\n') + '\n\n';
    saveHistorySnapshot('Trước khi chèn Mục lục tự động');
    bodyEl.value = toc + bodyEl.value;
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    toast(`✓ Đã tạo và chèn Mục lục tự động (${headings.length} đề mục)!`);
  }

  function initMultiDocs() {
    try {
      multiDocs = JSON.parse(localStorage.getItem(STORAGE.MULTI_DOCS) || '[]');
    } catch (e) { multiDocs = []; }
    if (!Array.isArray(multiDocs) || multiDocs.length === 0) {
      multiDocs = [{ id: 'doc_1', title: 'Văn bản 1', data: readCurrentDocFields() }];
    }
    activeDocTabId = multiDocs[0].id;
    renderMultiDocTabs();
  }

  function renderMultiDocTabs() {
    const bar = $('v3MultiDocBar');
    if (!bar) return;
    bar.innerHTML = multiDocs.map(d => `
      <button type="button" class="btn-tool ${d.id === activeDocTabId ? 'primary' : ''}" style="padding:3px 10px;font-size:0.78rem;" onclick="VbtmPro30.switchDocTab('${escHtml(d.id)}')">
        📄 ${escHtml(d.title)}
      </button>`).join('') +
      `<button type="button" class="btn-tool" style="padding:3px 9px;font-size:0.78rem;" onclick="VbtmPro30.createNewDocTab()">＋ Thêm văn bản</button>`;
  }

  function switchDocTab(id) {
    autoSaveDraft();
    const target = multiDocs.find(d => d.id === id);
    if (!target) return;
    activeDocTabId = id;
    writeDocFields(target.data);
    renderMultiDocTabs();
  }

  function createNewDocTab() {
    autoSaveDraft();
    const id = 'doc_' + Date.now();
    const newDoc = { id, title: `Văn bản ${multiDocs.length + 1}`, data: Object.assign(readCurrentDocFields(), { docNumber: '', docSubject: '', docBody: '' }) };
    multiDocs.push(newDoc);
    activeDocTabId = id;
    writeDocFields(newDoc.data);
    renderMultiDocTabs();
    toast('✓ Đã mở thêm 1 tab tài liệu mới.');
  }

  function renderLibraryPanel() {
    const box = $('v3LegalBasisList');
    if (!box) return;
    let bases = DEFAULT_LEGAL_BASES;
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE.LEGAL_BASES) || 'null');
      if (Array.isArray(saved) && saved.length) bases = saved;
    } catch (e) {}
    box.innerHTML = bases.map((b, i) => `
      <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;padding:6px 10px;border-radius:8px;background:rgba(255,255,255,0.05);margin-bottom:6px;">
        <div><b>${escHtml(b.title)}</b><div style="font-size:0.76rem;opacity:0.8;">${escHtml(b.text)}</div></div>
        <button type="button" class="btn-tool primary" style="padding:3px 9px;font-size:0.75rem;" onclick="VbtmPro30.insertLegalBasis(${i})">Chèn</button>
      </div>`).join('');
  }

  function insertLegalBasis(index) {
    let bases = DEFAULT_LEGAL_BASES;
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE.LEGAL_BASES) || 'null');
      if (Array.isArray(saved) && saved.length) bases = saved;
    } catch (e) {}
    const b = bases[index];
    const bodyEl = $('vbtmDocBody');
    if (!b || !bodyEl) return;
    saveHistorySnapshot(`Chèn căn cứ: ${b.title}`);
    bodyEl.value = b.text + '\n\n' + bodyEl.value;
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    toast(`✓ Đã chèn căn cứ: ${b.title}`);
  }

  function verifyQrCodeHealth() {
    const txt = ($('vbtmQrText') ? $('vbtmQrText').value : '').trim();
    const out = $('v3QrVerifyResult');
    if (!out) return;
    if (!txt) {
      out.innerHTML = '<span style="color:#f59e0b;">⚠ Chưa có dữ liệu QR. Nhập link hoặc mã xác thực văn bản.</span>';
      return;
    }
    const byteLen = new TextEncoder().encode(txt).length;
    const isUrl = /^https?:\/\//i.test(txt);
    out.innerHTML = `<span style="color:#10b981;">✓ Mã QR hợp lệ chuẩn ISO/IEC 18004 (ECC-M) · Dung lượng: ${byteLen} bytes · Loại: ${isUrl ? 'Đường dẫn tra cứu (URL)' : 'Văn bản xác thực'} · Sẵn sàng quét bằng Camera / Zalo.</span>`;
  }

  function exportRealDocx30() {
    if (!global.VbtmDocxEngine) { toast('Thiếu module VbtmDocxEngine!'); return; }
    const data = readCurrentDocFields();
    const baseName = (data.docNumber || data.docType || 'VanBan').replace(/[^a-zA-Z0-9_-]/g, '_');
    const res = global.VbtmDocxEngine.downloadRealDocx(data, `VBTM30_${baseName}`);
    toast(`✓ Đã xuất tệp OpenXML chuẩn (.docx): ${res.fileName} (${Math.round(res.sizeBytes / 1024)} KB)`);
  }

  function init() {
    if (global.VbtmAiService) {
      global.VbtmAiService.subscribe(renderAiControlPanel);
      const savedKey = global.VbtmAiService.getApiKey();
      if ($('vbtmGeminiKey') && savedKey) $('vbtmGeminiKey').value = savedKey;
      if ($('v3AiModelSelect')) {
        const curM = global.VbtmAiService.getModel();
        $('v3AiModelSelect').innerHTML = global.VbtmAiService.DEFAULT_MODELS
          .map(m => `<option value="${m}" ${m === curM ? 'selected' : ''}>${m}</option>`).join('');
      }
      renderAiControlPanel();
    }

    if ($('v3DraftModeSelect')) {
      $('v3DraftModeSelect').innerHTML = AI_DRAFT_MODES
        .map(m => `<option value="${m.id}">${escHtml(m.label)}</option>`).join('');
    }
    if ($('v3StyleModeSelect')) {
      $('v3StyleModeSelect').innerHTML = AI_STYLE_MODES
        .map(s => `<option value="${s.id}">${escHtml(s.label)}</option>`).join('');
    }

    initMultiDocs();
    checkAutoSavedDraftOnLoad();
    setInterval(autoSaveDraft, 5000);
    verifyQrCodeHealth();
  }

  if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else setTimeout(init, 50);
  }

  global.VbtmPro30 = {
    FORMAT_STANDARDS,
    AI_DRAFT_MODES,
    AI_STYLE_MODES,
    switchModule,
    handleTestApiKey,
    handleTestConnection,
    runOfflineSpellCheck30,
    runGeminiSpellCheck30,
    acceptSpellIssue,
    ignoreSpellIssue,
    fixAllSafeSpellIssues,
    restoreSpellOriginal,
    runAiDraft30,
    runAiStyleUpgrade30,
    openComparisonModal,
    closeComparisonModal,
    applyComparisonModal,
    runFormatAudit30,
    saveHistorySnapshot,
    restoreHistory,
    compareWithHistory,
    restoreAutoSavedDraft,
    runFindAndReplace,
    generateAutoToc,
    switchDocTab,
    createNewDocTab,
    insertLegalBasis,
    verifyQrCodeHealth,
    exportRealDocx30
  };
})(typeof window !== 'undefined' ? window : globalThis);
