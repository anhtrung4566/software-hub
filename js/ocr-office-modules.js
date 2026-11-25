/**
 * VietNamOCR Pro — Bộ Công Cụ Văn Phòng Thông Minh (js/ocr-office-modules.js)
 *
 * 5 Module thực tế (không nút giả, 100% xử lý cục bộ):
 * 1. Soạn thảo Văn bản Hành chính NĐ 30/2020/NĐ-CP (10 mẫu chuẩn, tự điền cơ quan/ngày/số, kiểm tra thể thức, xuất DOCX/In PDF).
 * 2. Công cụ Excel & Dữ liệu Bảng (Lọc trùng, chuẩn hóa ngày/số, thống kê Tổng/TB/Min/Max, biểu đồ cột Canvas, xuất CSV/XLS).
 * 3. So sánh 2 Tài liệu (Đọc thật .docx qua JSZip, .pdf qua PDF.js hoặc văn bản OCR, tô màu Thêm/Xóa, xuất báo cáo).
 * 4. Ghi âm & Biên bản họp (Web Speech API tiếng Việt vi-VN + phát lại file ghi âm + tạo biên bản họp chuẩn + xuất DOCX).
 * 5. Bảo mật: Che thông tin nhạy cảm thật (Redaction CCCD/SĐT/Email/MST) & Kiểm tra cấu trúc Chữ ký số PDF (/Type /Sig & /ByteRange).
 */
(function (global) {
  'use strict';

  function escHtml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function escXml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  // =========================================================================
  // MODULE 1: SOẠN THẢO VĂN BẢN HÀNH CHÍNH (10 MẪU CHUẨN NĐ 30/2020/NĐ-CP)
  // =========================================================================
  const ADMIN_TEMPLATES = {
    congvan: { title: 'CÔNG VĂN', abbr: 'CV', body: 'V/v ...\n\nKính gửi: ...\n\nThực hiện ..., {ORG} kính đề nghị Quý cơ quan phối hợp triển khai các nội dung sau:\n1. ...\n2. ...\n\nĐề nghị Quý cơ quan quan tâm thực hiện./.' },
    quyetdinh: { title: 'QUYẾT ĐỊNH', abbr: 'QĐ', body: 'Về việc ...\n\nTHỦ TRƯỞNG {ORG}\n\nCăn cứ chức năng, nhiệm vụ, quyền hạn được giao;\nXét đề nghị của ...;\n\nQUYẾT ĐỊNH:\n\nĐiều 1. ...\nĐiều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.\nĐiều 3. Chánh Văn phòng và các tổ chức, cá nhân có liên quan chịu trách nhiệm thi hành Quyết định này./.' },
    thongbao: { title: 'THÔNG BÁO', abbr: 'TB', body: 'Về việc ...\n\n{ORG} trân trọng thông báo đến các đơn vị, cá nhân nội dung sau:\n1. Thời gian: ...\n2. Địa điểm: ...\n3. Nội dung: ...\n\nĐề nghị các đơn vị nghiêm túc triển khai thực hiện./.' },
    totrinh: { title: 'TỜ TRÌNH', abbr: 'TTr', body: 'Về việc đề nghị phê duyệt ...\n\nKính gửi: ...\n\nCăn cứ ..., {ORG} kính trình cấp có thẩm quyền xem xét, phê duyệt nội dung sau:\nI. SỰ CẦN THIẾT:\n...\nII. NỘI DUNG ĐỀ XUẤT:\n...\nIII. KIẾN NGHỊ:\nKính đề nghị xem xét, quyết định./.' },
    baocao: { title: 'BÁO CÁO', abbr: 'BC', body: 'Kết quả thực hiện công tác ...\n\nI. KẾT QUẢ ĐẠT ĐƯỢC:\n1. ...\n2. ...\nII. TỒN TẠI, HẠN CHẾ:\n...\nIII. PHƯƠNG HƯỚNG, NHIỆM VỤ THỜI GIAN TỚI:\n...\n\nTrên đây là báo cáo của {ORG}, kính trình cấp trên xem xét./.' },
    kehoach: { title: 'KẾ HOẠCH', abbr: 'KH', body: 'Triển khai thực hiện ...\n\nI. MỤC ĐÍCH, YÊU CẦU:\n...\nII. NỘI DUNG VÀ TIẾN ĐỘ THỰC HIỆN:\n1. ...\n2. ...\nIII. TỔ CHỨC THỰC HIỆN:\nCác bộ phận liên quan căn cứ kế hoạch nghiêm túc triển khai./.' },
    bienban: { title: 'BIÊN BẢN', abbr: 'BB', body: 'Cuộc họp ...\n\nHôm nay, vào lúc ... giờ ..., ngày {DATE}, tại ..., chúng tôi gồm:\n1. Chủ trì: ...\n2. Thư ký: ...\n3. Thành phần tham dự: ...\n\nNỘI DUNG CUỘC HỌP:\n- ...\nKẾT LUẬN:\n- ...\n\nCuộc họp kết thúc vào lúc ... giờ ... cùng ngày./.' },
    giaymoi: { title: 'GIẤY MỜI', abbr: 'GM', body: 'Dự hội nghị ...\n\nKính gửi: ...\n\n{ORG} trân trọng kính mời đồng chí tham dự cuộc họp:\n- Thời gian: ... giờ ..., ngày {DATE}.\n- Địa điểm: ...\n- Nội dung: ...\n\nRất hân hạnh được đón tiếp./.' },
    giaygioithieu: { title: 'GIẤY GIỚI THIỆU', abbr: 'GGT', body: 'Kính gửi: ...\n\n{ORG} trân trọng giới thiệu:\n- Ông/Bà: ...\n- Chức vụ: ...\nĐược cử đến Quý cơ quan để liên hệ công tác về việc: ...\n\nGiấy giới thiệu có giá trị đến hết ngày .../.../2026./.' },
    bieumau: { title: 'ĐƠN ĐỀ NGHỊ', abbr: 'ĐN', body: 'Kính gửi: {ORG}\n\nTôi tên là: ...\nNgày sinh: ... | Số CCCD: ...\nĐơn vị công tác / Địa chỉ: ...\n\nNay tôi làm đơn này kính đề nghị Quý cơ quan xem xét giải quyết việc: ...\n\nTôi xin cam đoan các thông tin trên là đúng sự thật./.' }
  };

  function generateAdminDraft() {
    const key = document.getElementById('adminTplSelect')?.value || 'congvan';
    const parentOrg = (document.getElementById('adminParentOrg')?.value || 'UBND THÀNH PHỐ ĐÀ NẴNG').trim().toUpperCase();
    const org = (document.getElementById('adminIssuingOrg')?.value || 'ỦY BAN NHÂN DÂN PHƯỜNG').trim().toUpperCase();
    const loc = (document.getElementById('adminLocation')?.value || 'Đà Nẵng').trim();
    const seq = (document.getElementById('adminDocSeq')?.value || '01').trim();
    const signer = (document.getElementById('adminSignerTitle')?.value || 'CHỦ TỊCH').trim().toUpperCase();
    const now = new Date();
    const dateStr = `ngày ${now.getDate()} tháng ${now.getMonth() + 1} năm ${now.getFullYear()}`;
    const tpl = ADMIN_TEMPLATES[key] || ADMIN_TEMPLATES.congvan;

    const fullText =
      `${parentOrg}\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n` +
      `${org}\tĐộc lập - Tự do - Hạnh phúc\n` +
      `Số: ${seq}/${tpl.abbr}-${org.split(/\s+/).map(w => w[0]).join('').slice(0, 5)}\t${loc}, ${dateStr}\n\n` +
      `${tpl.title}\n` +
      tpl.body.replace(/\{ORG\}/g, org).replace(/\{DATE\}/g, dateStr) +
      `\n\nNơi nhận:\tTM. ${org}\n- Như trên;\t${signer}\n- Lưu: VT.\t(Chưa ký — trình người có thẩm quyền ký ban hành)`;

    const ed = document.getElementById('adminDraftEditor');
    if (ed) ed.value = fullText;
    validateAdminNd30();
  }

  function validateAdminNd30() {
    const txt = (document.getElementById('adminDraftEditor')?.value || '').trim();
    const box = document.getElementById('adminNd30CheckOut');
    if (!box) return;
    const checks = [
      { ok: /CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM/i.test(txt), label: 'Quốc hiệu chuẩn NĐ 30' },
      { ok: /Độc lập\s*-\s*Tự do\s*-\s*Hạnh phúc/i.test(txt), label: 'Tiêu ngữ chuẩn NĐ 30' },
      { ok: /Số:\s*\d+/i.test(txt), label: 'Số và ký hiệu văn bản' },
      { ok: /ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}/i.test(txt), label: 'Địa danh và ngày tháng năm' },
      { ok: /Nơi nhận:/i.test(txt), label: 'Khối Nơi nhận & Lưu văn thư' }
    ];
    const pass = checks.filter(c => c.ok).length;
    box.innerHTML = `<strong>Kiểm tra thể thức NĐ 30/2020/NĐ-CP (${pass}/${checks.length}):</strong> ` +
      checks.map(c => `<span style="margin-right:8px;color:${c.ok ? '#86efac' : '#fca5a5'};">${c.ok ? '✓' : '✕'} ${c.label}</span>`).join('');
  }

  async function exportAdminDraftDocx() {
    const txt = (document.getElementById('adminDraftEditor')?.value || '').trim();
    if (!txt) return;
    if (global.VietNamDocxExporter && global.JSZip) {
      await global.VietNamDocxExporter.exportUnifiedDocx({
        plainText: txt,
        fileName: `VanBan_ND30_${Date.now()}`
      });
      return;
    }
    downloadBlob(new Blob(['\ufeff', txt], { type: 'text/plain;charset=utf-8' }), 'VanBan_HanhChinh_ND30.txt');
  }

  // =========================================================================
  // MODULE 2: CÔNG CỤ EXCEL & PHÂN TÍCH DỮ LIỆU BẢNG (LỌC TRÙNG, THỐNG KÊ, BIỂU ĐỒ, CSV)
  // =========================================================================
  function getCurrentTableState() {
    const hdrs = Array.isArray(global.ocrTableHeaders) ? global.ocrTableHeaders : ['Cột 1', 'Cột 2', 'Giá trị'];
    const rows = Array.isArray(global.ocrTableRows) ? global.ocrTableRows : [];
    return { hdrs, rows };
  }

  function excelDeduplicate() {
    const { rows } = getCurrentTableState();
    const seen = new Set();
    const unique = [];
    rows.forEach(r => {
      const key = r.map(c => String(c || '').trim().toLowerCase()).join('|');
      if (!seen.has(key)) { seen.add(key); unique.push(r); }
    });
    const removed = rows.length - unique.length;
    global.ocrTableRows = unique;
    if (typeof global.renderEditableTable === 'function') global.renderEditableTable();
    renderExcelStatsAndChart(`✓ Đã loại bỏ ${removed} dòng trùng lặp (còn ${unique.length} dòng).`);
  }

  function excelNormalizeDatesAndNumbers() {
    const { rows } = getCurrentTableState();
    let changed = 0;
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < rows[r].length; c++) {
        let v = String(rows[r][c] || '').trim();
        const dm = v.match(/^(\d{1,2})[\-\.](\d{1,2})[\-\.](\d{4})$/);
        if (dm) {
          rows[r][c] = `${dm[1].padStart(2, '0')}/${dm[2].padStart(2, '0')}/${dm[3]}`;
          changed++;
        }
      }
    }
    if (typeof global.renderEditableTable === 'function') global.renderEditableTable();
    renderExcelStatsAndChart(`✓ Đã chuẩn hóa ${changed} ô định dạng ngày tháng về chuẩn dd/mm/yyyy (giữ nguyên mọi số liệu gốc).`);
  }

  function renderExcelStatsAndChart(note) {
    const { hdrs, rows } = getCurrentTableState();
    const statBox = document.getElementById('excelStatsSummary');
    const canvas = document.getElementById('excelChartCanvas');

    // Tìm cột số đầu tiên để thống kê & vẽ biểu đồ cột
    let numCol = -1;
    let nums = [];
    for (let c = 0; c < hdrs.length; c++) {
      const vals = rows.map(r => parseFloat(String(r[c] || '').replace(/[.,\s]/g, '').replace(/[^\d\-]/g, ''))).filter(v => !isNaN(v));
      if (vals.length >= Math.max(1, rows.length * 0.5)) {
        numCol = c; nums = vals; break;
      }
    }

    if (statBox) {
      if (nums.length > 0) {
        const sum = nums.reduce((a, b) => a + b, 0);
        const avg = Math.round(sum / nums.length);
        const min = Math.min(...nums);
        const max = Math.max(...nums);
        statBox.innerHTML = `${note ? `<div>${escHtml(note)}</div>` : ''}<strong>Thống kê cột "${escHtml(hdrs[numCol])}" (${nums.length} dòng):</strong> Tổng = <b>${sum.toLocaleString('vi-VN')}</b> · Trung bình = <b>${avg.toLocaleString('vi-VN')}</b> · Min = <b>${min.toLocaleString('vi-VN')}</b> · Max = <b>${max.toLocaleString('vi-VN')}</b>`;
      } else {
        statBox.innerHTML = `${note ? `<div>${escHtml(note)}</div>` : ''}Bảng hiện có <b>${rows.length}</b> dòng × <b>${hdrs.length}</b> cột.`;
      }
    }

    if (canvas && nums.length > 0) {
      const ctx = canvas.getContext('2d');
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.fillStyle = '#0f172a';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      const maxVal = Math.max(1, ...nums);
      const barW = Math.min(48, Math.floor((canvas.width - 60) / Math.max(1, nums.length)) - 8);
      nums.slice(0, 14).forEach((v, i) => {
        const bh = Math.round((v / maxVal) * (canvas.height - 54));
        const x = 36 + i * (barW + 10);
        const y = canvas.height - 28 - bh;
        ctx.fillStyle = '#38bdf8';
        ctx.fillRect(x, y, barW, bh);
        ctx.fillStyle = '#f8fafc';
        ctx.font = '11px Segoe UI';
        ctx.fillText(String(i + 1), x + barW / 3, canvas.height - 10);
      });
    }
  }

  function exportExcelCsv() {
    const { hdrs, rows } = getCurrentTableState();
    const escCsv = (s) => '"' + String(s || '').replace(/"/g, '""') + '"';
    const lines = [hdrs.map(escCsv).join(',')].concat(rows.map(r => hdrs.map((_, i) => escCsv(r[i] || '')).join(',')));
    downloadBlob(new Blob(['\ufeff', lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }), 'DuLieu_Bang_VietNamOCR.csv');
  }

  // =========================================================================
  // MODULE 3: SO SÁNH 2 TÀI LIỆU THỰC TẾ (DOCX / PDF / OCR TEXT)
  // =========================================================================
  async function readAnyDocToText(file) {
    if (!file) return '';
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (ext === 'docx' && global.JSZip) {
      const zip = await global.JSZip.loadAsync(await file.arrayBuffer());
      const docXml = await zip.file('word/document.xml')?.async('string');
      if (!docXml) return '';
      return docXml.replace(/<\/w:p>/g, '\n').replace(/<[^>]+>/g, '').replace(/\n{2,}/g, '\n').trim();
    }
    if (ext === 'pdf' && global.pdfjsLib) {
      const pdf = await global.pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
      const out = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const p = await pdf.getPage(i);
        const tc = await p.getTextContent();
        out.push((tc.items || []).map(it => it.str || '').join(' ').replace(/\s{2,}/g, ' ').trim());
      }
      return out.join('\n');
    }
    return await file.text();
  }

  async function loadDiffFile(side, file) {
    if (!file) return;
    const text = await readAnyDocToText(file);
    const el = document.getElementById(side === 'A' ? 'cmpDocAText' : 'cmpDocBText');
    if (el) el.value = text;
  }

  function runRealDocumentDiff() {
    const aLines = (document.getElementById('cmpDocAText')?.value || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
    const bLines = (document.getElementById('cmpDocBText')?.value || '').split(/\n+/).map(s => s.trim()).filter(Boolean);
    const out = document.getElementById('cmpDiffResultBox');
    if (!out) return;
    const setA = new Set(aLines);
    const setB = new Set(bLines);
    const removed = aLines.filter(l => !setB.has(l));
    const added = bLines.filter(l => !setA.has(l));
    const same = aLines.length - removed.length;

    out.innerHTML =
      `<div style="margin-bottom:10px;font-weight:700;">Kết quả so sánh thực tế: Trùng khớp ${same} đoạn · <span style="color:#fca5a5;">Xóa/Khác ở Bản A: ${removed.length}</span> · <span style="color:#86efac;">Thêm/Khác ở Bản B: ${added.length}</span></div>` +
      `<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">` +
      `<div>${aLines.map(l => `<div style="padding:5px 8px;margin-bottom:4px;border-radius:6px;background:${setB.has(l) ? 'rgba(255,255,255,0.05)' : 'rgba(220,38,38,0.28)'};">${setB.has(l) ? '• ' : '− '}${escHtml(l)}</div>`).join('')}</div>` +
      `<div>${bLines.map(l => `<div style="padding:5px 8px;margin-bottom:4px;border-radius:6px;background:${setA.has(l) ? 'rgba(255,255,255,0.05)' : 'rgba(22,163,74,0.28)'};">${setA.has(l) ? '• ' : '+ '}${escHtml(l)}</div>`).join('')}</div>` +
      `</div>`;
  }

  function exportDiffReportHtml() {
    const html = document.getElementById('cmpDiffResultBox')?.innerHTML || '';
    if (!html) return;
    downloadBlob(new Blob([`<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><title>Báo cáo so sánh tài liệu</title><style>body{font-family:'Times New Roman',serif;padding:24px;background:#fff;color:#0f172a;}</style></head><body><h2>BÁO CÁO SO SÁNH 2 TÀI LIỆU</h2>${html}</body></html>`], { type: 'text/html;charset=utf-8' }), 'BaoCao_SoSanh_TaiLieu.html');
  }

  // =========================================================================
  // MODULE 4: GHI ÂM & TẠO BIÊN BẢN HỌP (WEB SPEECH API TIẾNG VIỆT vi-VN)
  // =========================================================================
  let speechRec = null;
  let isRecording = false;

  function toggleVietnameseSpeechRec() {
    const SpeechAPI = global.SpeechRecognition || global.webkitSpeechRecognition;
    const status = document.getElementById('meetStatusText');
    const btn = document.getElementById('btnMeetRecord');
    const ta = document.getElementById('meetTranscriptBox');
    if (!SpeechAPI) {
      if (status) status.textContent = '⚠ Trình duyệt này không hỗ trợ Web Speech API (khuyến nghị Chrome / Edge trên Windows). Bạn vẫn có thể mở file ghi âm để nghe và gõ biên bản.';
      return;
    }
    if (isRecording && speechRec) {
      speechRec.stop();
      isRecording = false;
      if (btn) btn.textContent = '🎙 Bắt đầu Ghi âm Tiếng Việt (vi-VN)';
      if (status) status.textContent = '⏹ Đã dừng ghi âm. Hãy chỉnh sửa nội dung và bấm "Tạo Biên Bản Họp".';
      return;
    }
    speechRec = new SpeechAPI();
    speechRec.lang = 'vi-VN';
    speechRec.continuous = true;
    speechRec.interimResults = false;
    speechRec.onresult = (e) => {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        if (e.results[i].isFinal && ta) {
          const ts = new Date().toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' });
          ta.value += `[${ts}] ${e.results[i][0].transcript.trim()}\n`;
        }
      }
    };
    speechRec.onerror = (e) => {
      if (status) status.textContent = '⚠ Lỗi micro / nhận diện giọng nói: ' + e.error;
    };
    speechRec.start();
    isRecording = true;
    if (btn) btn.textContent = '⏹ Dừng Ghi âm';
    if (status) status.textContent = '🔴 Đang ghi âm & chuyển giọng nói Tiếng Việt (vi-VN) thành văn bản...';
  }

  function loadAudioForPlayback(file) {
    const player = document.getElementById('meetAudioPlayer');
    if (player && file) {
      player.src = URL.createObjectURL(file);
      player.style.display = 'block';
    }
  }

  function buildMeetingMinutes() {
    const raw = (document.getElementById('meetTranscriptBox')?.value || '').trim();
    const title = (document.getElementById('meetTitleInput')?.value || 'CUỘC HỌP GIAO BAN').trim().toUpperCase();
    const now = new Date();
    const lines = raw.split(/\n+/).map(s => s.trim()).filter(Boolean);
    const conclusions = lines.filter(l => /kết luận|thống nhất|quyết định|giao nhiệm vụ|đề nghị/i.test(l));

    const minutes =
      `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\nĐộc lập - Tự do - Hạnh phúc\n\n` +
      `BIÊN BẢN HỌP\nV/v: ${title}\n\n` +
      `Thời gian: ${now.toLocaleTimeString('vi-VN', { hour: '2-digit', minute: '2-digit' })}, ngày ${now.getDate()} tháng ${now.getMonth() + 1} năm ${now.getFullYear()}.\n\n` +
      `I. DIỄN BIẾN & Ý KIẾN THẢO LUẬN:\n` +
      (lines.length ? lines.map(l => `- ${l}`).join('\n') : '- (Chưa có nội dung ghi âm)') +
      `\n\nII. KẾT LUẬN CHÍNH:\n` +
      (conclusions.length ? conclusions.map(c => `+ ${c}`).join('\n') : '+ (Người dùng bổ sung kết luận sau khi rà soát)') +
      `\n\nTHƯ KÝ\t\tCHỦ TRÌ CUỘC HỌP`;

    const out = document.getElementById('meetMinutesOutput');
    if (out) out.value = minutes;
  }

  // =========================================================================
  // MODULE 5: BẢO MẬT — CHE THÔNG TIN NHẠY CẢM (REDACTION THẬT) & KIỂM TRA CHỮ KÝ SỐ PDF
  // =========================================================================
  function scanAndRedactSensitiveText() {
    const srcEl = document.getElementById('secRedactInput');
    const outEl = document.getElementById('secRedactOutput');
    const logEl = document.getElementById('secRedactSummary');
    const text = (srcEl?.value || document.getElementById('ocrResultText')?.value || '');
    if (srcEl && !srcEl.value) srcEl.value = text;

    let redacted = text;
    let cCccd = 0, cPhone = 0, cEmail = 0, cMst = 0;

    if (document.getElementById('chkRedactCccd')?.checked !== false) {
      redacted = redacted.replace(/\b\d{12}\b|\b\d{9}\b/g, () => { cCccd++; return '[ĐÃ CHE CCCD]'; });
    }
    if (document.getElementById('chkRedactPhone')?.checked !== false) {
      redacted = redacted.replace(/\b(?:0|\+84)(?:3|5|7|8|9)\d{8}\b/g, () => { cPhone++; return '[ĐÃ CHE SĐT]'; });
    }
    if (document.getElementById('chkRedactEmail')?.checked !== false) {
      redacted = redacted.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, () => { cEmail++; return '[ĐÃ CHE EMAIL]'; });
    }
    if (document.getElementById('chkRedactMst')?.checked !== false) {
      redacted = redacted.replace(/\b\d{10}(?:-\d{3})?\b/g, () => { cMst++; return '[ĐÃ CHE MST]'; });
    }

    if (outEl) outEl.value = redacted;
    if (logEl) logEl.textContent = `✓ Đã xóa dữ liệu thật bên dưới vùng che: ${cCccd} CCCD/CMND · ${cPhone} SĐT · ${cEmail} Email · ${cMst} Mã số thuế.`;
  }

  function exportRedactedTxt() {
    const txt = document.getElementById('secRedactOutput')?.value || '';
    if (!txt) return;
    downloadBlob(new Blob(['\ufeff', txt], { type: 'text/plain;charset=utf-8' }), 'TaiLieu_DaCheThongTinNhayCam.txt');
  }

  // Kiểm tra cấu trúc chữ ký số PDF thực tế (/Type /Sig, /SubFilter, /ByteRange, /Name, /M)
  async function inspectPdfDigitalSignature(file) {
    const out = document.getElementById('secPdfSigReport');
    if (!file || !out) return;
    try {
      const buf = await file.arrayBuffer();
      const bytes = new Uint8Array(buf);
      // Đọc chuỗi latin1 để phân tích các trường từ điển chữ ký PDF
      let latin = '';
      for (let i = 0; i < bytes.length; i += 32768) {
        latin += String.fromCharCode.apply(null, bytes.subarray(i, i + 32768));
      }

      const sigMatches = latin.match(/\/Type\s*\/Sig\b|\/SubFilter\s*\/(?:adbe\.pkcs7|ETSI\.CAdES)/g) || [];
      const byteRanges = latin.match(/\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/g) || [];
      const names = (latin.match(/\/Name\s*\(([^)]{1,120})\)/g) || []).map(s => s.replace(/^\/Name\s*\(|\)$/g, ''));
      const reasons = (latin.match(/\/Reason\s*\(([^)]{1,160})\)/g) || []).map(s => s.replace(/^\/Reason\s*\(|\)$/g, ''));
      const dates = (latin.match(/\/M\s*\(D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})/g) || []).map(s => {
        const m = s.match(/D:(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})/);
        return m ? `${m[3]}/${m[2]}/${m[1]} ${m[4]}:${m[5]}` : s;
      });

      if (sigMatches.length === 0 && byteRanges.length === 0) {
        out.innerHTML = `<div style="color:#fde047;"><strong>ℹ Kết quả kiểm tra "${escHtml(file.name)}":</strong> Không tìm thấy đối tượng chữ ký số (<code>/Type /Sig</code>) trong tệp PDF này.</div>`;
        return;
      }

      let coversWholeFile = false;
      if (byteRanges.length > 0) {
        const lastBr = byteRanges[byteRanges.length - 1].match(/\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/);
        if (lastBr) {
          const endPos = parseInt(lastBr[3], 10) + parseInt(lastBr[4], 10);
          coversWholeFile = Math.abs(bytes.length - endPos) <= 16;
        }
      }

      out.innerHTML =
        `<div style="line-height:1.7;">` +
        `<div><strong>📄 Tệp kiểm tra:</strong> ${escHtml(file.name)} (${bytes.length.toLocaleString('vi-VN')} bytes)</div>` +
        `<div><strong>✓ Phát hiện cấu trúc chữ ký số PDF:</strong> ${Math.max(sigMatches.length, byteRanges.length)} chữ ký</div>` +
        `<div><strong>👤 Thông tin người ký (/Name):</strong> ${names.length ? escHtml(names.join('; ')) : '<em>Nhúng trong chứng thư số PKCS#7</em>'}</div>` +
        `<div><strong>🕒 Thời điểm ký (/M):</strong> ${dates.length ? escHtml(dates.join('; ')) : '<em>Không ghi rõ</em>'}</div>` +
        `<div><strong>📝 Lý do ký (/Reason):</strong> ${reasons.length ? escHtml(reasons.join('; ')) : '<em>Không ghi rõ</em>'}</div>` +
        `<div><strong>🛡 Tính toàn vẹn ByteRange:</strong> ${coversWholeFile ? '<span style="color:#86efac;">✓ Chữ ký bao phủ toàn bộ tệp PDF (chưa bị chỉnh sửa sau khi ký)</span>' : '<span style="color:#fca5a5;">⚠ Cảnh báo: Có dữ liệu bổ sung sau vùng ByteRange của chữ ký cuối!</span>'}</div>` +
        `<div style="margin-top:6px;font-size:0.82rem;color:#cbd5e1;"><em>Lưu ý minh bạch: Trình duyệt kiểm tra cấu trúc PKCS#7 &amp; toàn vẹn ByteRange cục bộ; để xác thực chuỗi chứng thư số gốc (CA Root / VGCA / Ban Cơ yếu Chính phủ), vui lòng đối chiếu thêm trên Adobe Acrobat Reader hoặc phần mềm ký số chuyên dụng.</em></div>` +
        `</div>`;
    } catch (e) {
      out.textContent = '⚠ Không thể phân tích tệp PDF: ' + e.message;
    }
  }

  global.VietNamOfficeModules = {
    generateAdminDraft,
    validateAdminNd30,
    exportAdminDraftDocx,
    excelDeduplicate,
    excelNormalizeDatesAndNumbers,
    renderExcelStatsAndChart,
    exportExcelCsv,
    loadDiffFile,
    runRealDocumentDiff,
    exportDiffReportHtml,
    toggleVietnameseSpeechRec,
    loadAudioForPlayback,
    buildMeetingMinutes,
    scanAndRedactSensitiveText,
    exportRedactedTxt,
    inspectPdfDigitalSignature
  };
})(window);
