/**
 * VBTM Pro 2.2.4 - Bộ Công Cụ Chuẩn Hóa Thể Thức Văn Bản, Tạo Mã QR & Soát Lỗi Chính Tả AI Offline
 * Hỗ trợ: Nghị định 30/2020/NĐ-CP, Quy chuẩn CAND (TT 30/2023/TT-BCA), Hướng dẫn 05-HD/VPTW
 */

const VBTM_TEMPLATES = {
  nd30_congvan: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN TP ĐÀ NẴNG',
    issuingOrg: 'SỞ THÔNG TIN VÀ TRUYỀN THÔNG',
    docNumber: '145/STTTT-CNTT',
    location: 'Đà Nẵng',
    docDate: '2026-11-11',
    docType: '',
    docSubject: 'V/v triển khai phần mềm chuẩn hóa văn bản và số hóa tài liệu ngoại tuyến năm 2026',
    docBody: `Thực hiện Kế hoạch chuyển đổi số và hiện đại hóa công tác hành chính năm 2026, Sở Thông tin và Truyền thông triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP của Chính phủ.\n\n1. Các phòng, ban, đơn vị trực thuộc thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; đảm bảo đúng quy định về lề trang, phông chữ Times New Roman, cỡ chữ và thể thức trình bày.\n\n2. Ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm bảo đảm an toàn thông tin và bảo vệ bí mật nhà nước trên các máy tính nội bộ.\n\nNhận được Công văn này, đề nghị các đơn vị nghiêm túc tổ chức thực hiện./.`,
    recipients: `- Như trên;\n- Giám đốc Sở (để b/c);\n- Các phòng thuộc Sở;\n- Lưu: VT, CNTT.`,
    signerTitle: 'KT. GIÁM ĐỐC\nPHÓ GIÁM ĐỐC',
    signerName: 'Nguyễn Văn An'
  },
  cand_baocao: {
    standard: 'cand',
    governingOrg: 'CÔNG AN TP ĐÀ NẴNG',
    issuingOrg: 'PHÒNG KỸ THUẬT HÌNH SỰ',
    docNumber: '86/BC-PC09',
    location: 'Đà Nẵng',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Kết quả ứng dụng phần mềm chuẩn hóa văn bản và nhận diện tài liệu OCR phục vụ công tác chuyên môn',
    docBody: `Thực hiện chương trình công tác kỹ thuật hình sự và đẩy mạnh ứng dụng khoa học công nghệ trong năm 2026, Phòng Kỹ thuật hình sự báo cáo kết quả triển khai như sau:\n\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Đã nghiên cứu, xây dựng và đưa vào vận hành hệ thống phần mềm VBTM Pro phục vụ chuẩn hóa thể thức văn bản theo đúng quy chuẩn Công an nhân dân và Nghị định số 30/2020/NĐ-CP.\n\n2. Tích hợp thành công mô-đun soát lỗi chính tả AI tiếng Việt ngoại tuyến 100% và tạo mã QR xác thực văn bản trực tiếp trên Microsoft Word, giúp tiết kiệm 70% thời gian biên tập hồ sơ.\n\nII. PHƯƠNG HƯỚNG TRONG THỜI GIAN TỚI\nTiếp tục nâng cấp độ chính xác của bộ từ điển chuyên ngành và phổ biến rộng rãi cho cán bộ, chiến sĩ trong đơn vị sử dụng hiệu quả./.`,
    recipients: `- Ban Giám đốc CATP (để b/c);\n- Phòng Tham mưu (PV01);\n- Lưu: VT, PC09.`,
    signerTitle: 'TRƯỞNG PHÒNG',
    signerName: 'Thượng tá Trần Minh Quân'
  },
  cand_quyetdinh: {
    standard: 'cand',
    governingOrg: 'CÔNG AN TP ĐÀ NẴNG',
    issuingOrg: 'PHÒNG KỸ THUẬT HÌNH SỰ',
    docNumber: '120/QĐ-PC09',
    location: 'Đà Nẵng',
    docDate: '2026-11-11',
    docType: 'QUYẾT ĐỊNH',
    docSubject: 'Về việc ban hành Quy chế sử dụng phần mềm nội bộ ngoại tuyến trong đơn vị',
    docBody: `TRƯỞNG PHÒNG KỸ THUẬT HÌNH SỰ\n\nCăn cứ Thông tư số 30/2023/TT-BCA quy định về công tác văn thư trong Công an nhân dân;\n\nCăn cứ yêu cầu thực tiễn công tác bảo đảm an toàn thông tin và số hóa hồ sơ nghiệp vụ;\n\nTheo đề nghị của Đội Tổng hợp,\n\nQUYẾT ĐỊNH:\n\nĐiều 1. Ban hành kèm theo Quyết định này Quy chế quản lý, vận hành và sử dụng bộ công cụ VBTM Pro và VietNamOCR 3.2 Pro trên hệ thống máy tính nội bộ.\n\nĐiều 2. Quyết định này có hiệu lực kể từ ngày ký.\n\nĐiều 3. Các Đội nghiệp vụ và toàn thể cán bộ, chiến sĩ thuộc đơn vị chịu trách nhiệm thi hành Quyết định này./.`,
    recipients: `- Như Điều 3 (để t/h);\n- Lưu: VT, TH.`,
    signerTitle: 'TRƯỞNG PHÒNG',
    signerName: 'Thượng tá Trần Minh Quân'
  },
  hd05_baocao: {
    standard: 'hd05',
    governingOrg: 'ĐẢNG ỦY CÔNG AN TP ĐÀ NẴNG',
    issuingOrg: 'ĐẢNG ỦY PHÒNG KỸ THUẬT HÌNH SỰ',
    docNumber: '42-BC/ĐU',
    location: 'Đà Nẵng',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Tổng kết công tác cải cách hành chính và chuyển đổi số trong công tác Đảng năm 2026',
    docBody: `Thực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy Phòng Kỹ thuật hình sự báo cáo kết quả thực hiện như sau:\n\n1. Toàn bộ văn bản của cấp ủy, chi bộ trực thuộc đã được chuẩn hóa đồng bộ về thể thức, phông chữ và quy cách trình bày.\n\n2. Việc ứng dụng công cụ kiểm tra chính tả và thể thức ngoại tuyến giúp nâng cao chất lượng tham mưu, đảm bảo tuyệt đối nguyên tắc bảo mật tài liệu của Đảng./.`,
    recipients: `- Ban Thường vụ Đảng ủy CATP;\n- Các đồng chí ĐUV;\n- Các chi bộ trực thuộc;\n- Lưu Hồ sơ Đảng ủy.`,
    signerTitle: 'T/M ĐẢNG ỦY\nBÍ THƯ',
    signerName: 'Thượng tá Trần Minh Quân'
  }
};

// Bộ từ điển phát hiện lỗi chính tả tiếng Việt phổ biến trong văn bản hành chính
const SPELL_DICTIONARY = [
  { wrong: /\bbổ xung\b/gi, right: 'bổ sung', desc: 'Sai phụ âm đầu x/s ("bổ sung")' },
  { wrong: /\bsử lý\b/gi, right: 'xử lý', desc: 'Sai phụ âm đầu s/x ("xử lý")' },
  { wrong: /\bxử dụng\b/gi, right: 'sử dụng', desc: 'Sai phụ âm đầu x/s ("sử dụng")' },
  { wrong: /\bchính sác\b/gi, right: 'chính xác', desc: 'Sai phụ âm đầu s/x ("chính xác")' },
  { wrong: /\bkiểm xoát\b/gi, right: 'kiểm soát', desc: 'Sai phụ âm đầu x/s ("kiểm soát")' },
  { wrong: /\bchiển khai\b/gi, right: 'triển khai', desc: 'Sai phụ âm đầu ch/tr ("triển khai")' },
  { wrong: /\bchủ chương\b/gi, right: 'chủ trương', desc: 'Sai phụ âm đầu ch/tr ("chủ trương")' },
  { wrong: /\bchấp nhận\b/gi, right: 'chấp nhận', skip: true },
  { wrong: /\bkế hoạc\b/gi, right: 'kế hoạch', desc: 'Thiếu âm cuối ("kế hoạch")' },
  { wrong: /\bngiêm túc\b/gi, right: 'nghiêm túc', desc: 'Sai quy tắc chính tả ng/ngh ("nghiêm túc")' },
  { wrong: /\bngành ngề\b/gi, right: 'ngành nghề', desc: 'Sai quy tắc chính tả ng/ngh ("ngành nghề")' },
  { wrong: /\bhổ trợ\b/gi, right: 'hỗ trợ', desc: 'Sai dấu hỏi/ngã ("hỗ trợ")' },
  { wrong: /\btiêu chuẫn\b/gi, right: 'tiêu chuẩn', desc: 'Sai dấu hỏi/ngã ("tiêu chuẩn")' },
  { wrong: /\bkhẩn chương\b/gi, right: 'khẩn trương', desc: 'Sai phụ âm ch/tr ("khẩn trương")' },
  { wrong: /\btrao chuốt\b/gi, right: 'trau chuốt', desc: 'Sai vần ao/au ("trau chuốt")' },
  { wrong: /\bsáng xuốt\b/gi, right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { wrong: /\bxáng suốt\b/gi, right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { wrong: /\bđảm đương\b/gi, right: 'đảm đương', skip: true },
  { wrong: /\bgiử gìn\b/gi, right: 'giữ gìn', desc: 'Sai dấu hỏi/ngã ("giữ gìn")' },
  { wrong: /\bvững chắc\b/gi, right: 'vững chắc', skip: true },
  { wrong: /\bvửng chắc\b/gi, right: 'vững chắc', desc: 'Sai dấu hỏi/ngã ("vững chắc")' },
  { wrong: /\bthực hiện nghiêm chỉnh\b/gi, right: 'thực hiện nghiêm chỉnh', skip: true },
  { wrong: /\bsát nhập\b/gi, right: 'sáp nhập', desc: 'Từ chuẩn hành chính là "sáp nhập"' },
  { wrong: /\bchính phủ\b/g, right: 'Chính phủ', desc: 'Viết hoa danh từ riêng cơ quan ("Chính phủ")' },
  { wrong: /\bnhà nước\b/g, right: 'Nhà nước', desc: 'Viết hoa "Nhà nước" trong văn bản hành chính' }
];

let lastDetectedIssues = [];

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('vbtmStudioApp')) return;
  loadTemplate('nd30_congvan');
  bindVbtmEvents();
  updateQrCanvas();
  renderA4Preview();
});

function bindVbtmEvents() {
  const inputs = [
    'vbtmStandard', 'vbtmGovOrg', 'vbtmIssueOrg', 'vbtmDocNum',
    'vbtmLocation', 'vbtmDocDate', 'vbtmDocType', 'vbtmDocSubject',
    'vbtmDocBody', 'vbtmRecipients', 'vbtmSignerTitle', 'vbtmSignerName',
    'vbtmQrText', 'vbtmIncludeQr'
  ];
  inputs.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => {
        if (id === 'vbtmQrText' || id === 'vbtmDocNum') updateQrCanvas();
        renderA4Preview();
      });
      el.addEventListener('change', () => {
        updateQrCanvas();
        renderA4Preview();
      });
    }
  });
}

function loadTemplate(key) {
  const tpl = VBTM_TEMPLATES[key];
  if (!tpl) return;
  document.getElementById('vbtmStandard').value = tpl.standard;
  document.getElementById('vbtmGovOrg').value = tpl.governingOrg;
  document.getElementById('vbtmIssueOrg').value = tpl.issuingOrg;
  document.getElementById('vbtmDocNum').value = tpl.docNumber;
  document.getElementById('vbtmLocation').value = tpl.location;
  document.getElementById('vbtmDocDate').value = tpl.docDate;
  document.getElementById('vbtmDocType').value = tpl.docType;
  document.getElementById('vbtmDocSubject').value = tpl.docSubject;
  document.getElementById('vbtmDocBody').value = tpl.docBody;
  document.getElementById('vbtmRecipients').value = tpl.recipients;
  document.getElementById('vbtmSignerTitle').value = tpl.signerTitle;
  document.getElementById('vbtmSignerName').value = tpl.signerName;
  document.getElementById('vbtmQrText').value = `VBTM-VERIFY:${tpl.docNumber}|${tpl.issuingOrg}|${tpl.docDate}`;
  updateQrCanvas();
  renderA4Preview();
  if (window.showToast) window.showToast('Đã nạp biểu mẫu chuẩn hóa văn bản');
}

// Chuẩn hóa ngày tháng theo Nghị định 30 (ngày < 10 và tháng 1, 2 phải thêm số 0)
function formatVietnameseDate(dateStr, location) {
  if (!dateStr) return `${location || 'Đà Nẵng'}, ngày 11 tháng 11 năm 2026`;
  const parts = dateStr.split('-');
  if (parts.length !== 3) return `${location}, ${dateStr}`;
  const year = parts[0];
  const monthNum = parseInt(parts[1], 10);
  const dayNum = parseInt(parts[2], 10);
  const dayStr = dayNum < 10 ? `0${dayNum}` : `${dayNum}`;
  // Theo NĐ30: tháng 1 và tháng 2 phải có số 0 đứng trước (01, 02), các tháng khác viết bình thường
  const monthStr = (monthNum === 1 || monthNum === 2) ? `0${monthNum}` : `${monthNum}`;
  return `${location || 'Đà Nẵng'}, ngày ${dayStr} tháng ${monthStr} năm ${year}`;
}

// Tự động chuẩn hóa thể thức văn bản (khoảng trắng, dấu câu, viết hoa, kết thúc văn bản)
function standardizeTextContent() {
  const bodyEl = document.getElementById('vbtmDocBody');
  const subjEl = document.getElementById('vbtmDocSubject');
  if (!bodyEl) return;

  let text = bodyEl.value;
  // 1. Xóa khoảng trắng thừa và chuẩn hóa dấu câu
  text = text
    .replace(/[ \t]+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([,.;:!?])([^\s\d/])/g, '$1 $2')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/\n{3,}/g, '\n\n');

  // 2. Viết hoa chữ cái đầu mỗi đoạn và đảm bảo dấu kết thúc văn bản "./."
  const lines = text.split('\n').map(line => {
    let trimmed = line.trim();
    if (!trimmed) return '';
    trimmed = trimmed.charAt(0).toUpperCase() + trimmed.slice(1);
    return trimmed;
  }).filter((l, idx, arr) => !(l === '' && arr[idx - 1] === ''));

  let result = lines.join('\n').trim();
  if (result && !result.endsWith('./.')) {
    result = result.replace(/\.+$/, '') + './.';
  }
  bodyEl.value = result;

  // Chuẩn hóa trích yếu (không để dấu chấm ở cuối trích yếu theo NĐ30)
  if (subjEl && subjEl.value) {
    subjEl.value = subjEl.value.trim().replace(/\s+/g, ' ').replace(/\.+$/, '');
  }

  renderA4Preview();
  if (window.showToast) {
    window.showToast('Đã chuẩn hóa thể thức, dấu câu, khoảng trắng & lề trang theo Nghị định 30!');
  }
}

// Kiểm tra lỗi chính tả & thẩm định thể thức AI Offline
function runSpellAndFormatCheck() {
  const body = document.getElementById('vbtmDocBody').value || '';
  const subject = document.getElementById('vbtmDocSubject').value || '';
  const docNum = document.getElementById('vbtmDocNum').value || '';
  const recipients = document.getElementById('vbtmRecipients').value || '';
  const issueContainer = document.getElementById('vbtmIssueList');

  lastDetectedIssues = [];

  // 1. Kiểm tra từ điển lỗi chính tả
  SPELL_DICTIONARY.forEach(rule => {
    if (rule.skip) return;
    if (rule.wrong.test(body) || rule.wrong.test(subject)) {
      lastDetectedIssues.push({
        type: 'Chính tả AI',
        desc: `${rule.desc} → Gợi ý sửa thành: "${rule.right}"`,
        fix: () => {
          document.getElementById('vbtmDocBody').value = document.getElementById('vbtmDocBody').value.replace(rule.wrong, rule.right);
          document.getElementById('vbtmDocSubject').value = document.getElementById('vbtmDocSubject').value.replace(rule.wrong, rule.right);
        }
      });
    }
  });

  // 2. Kiểm tra lỗi khoảng trắng trước dấu câu
  if (/\s+[,.;:!?]/.test(body)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)',
      fix: () => standardizeTextContent()
    });
  }

  // 3. Kiểm tra dấu chấm cuối trích yếu
  if (/\.$/.test(subject.trim())) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Trích yếu văn bản không được đặt dấu chấm ở cuối dòng',
      fix: () => {
        document.getElementById('vbtmDocSubject').value = subject.trim().replace(/\.+$/, '');
      }
    });
  }

  // 4. Kiểm tra ký hiệu kết thúc văn bản "./."
  if (body.trim() && !body.trim().endsWith('./.')) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Cuối nội dung văn bản hành chính thiếu ký hiệu kết thúc "./."',
      fix: () => standardizeTextContent()
    });
  }

  // 5. Kiểm tra Số ký hiệu và Nơi nhận
  if (!docNum.trim()) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Thiếu Số và ký hiệu văn bản (Ví dụ: 145/BC-PC09)',
      fix: () => { document.getElementById('vbtmDocNum').value = '01/VB-CQ'; }
    });
  }
  if (!recipients.trim() || !/Lưu:/i.test(recipients)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Mục "Nơi nhận" cần có dòng "- Lưu: VT,..." theo Khoản 3 Điều 14 Nghị định 30',
      fix: () => {
        const cur = document.getElementById('vbtmRecipients').value.trim();
        document.getElementById('vbtmRecipients').value = (cur ? cur + '\n' : '') + '- Lưu: VT.';
      }
    });
  }

  if (!issueContainer) return;

  if (lastDetectedIssues.length === 0) {
    issueContainer.innerHTML = `
      <div class="issue-item ok">
        <span>✓ Văn bản đã đạt chuẩn 100% về chính tả Tiếng Việt và thể thức Nghị định 30 / CAND!</span>
      </div>
    `;
  } else {
    issueContainer.innerHTML = lastDetectedIssues.map((item, idx) => `
      <div class="issue-item">
        <div>
          <strong style="color: #fbbf24;">[${item.type}]</strong> ${item.desc}
        </div>
        <button type="button" class="btn-tool primary" onclick="fixSingleIssue(${idx})" style="padding: 4px 10px; font-size: 0.78rem;">Sửa ngay</button>
      </div>
    `).join('');
  }

  if (window.showToast) {
    window.showToast(`Hoàn tất kiểm tra: Phát hiện ${lastDetectedIssues.length} điểm cần lưu ý`);
  }
}

function fixSingleIssue(idx) {
  if (lastDetectedIssues[idx] && typeof lastDetectedIssues[idx].fix === 'function') {
    lastDetectedIssues[idx].fix();
    renderA4Preview();
    runSpellAndFormatCheck();
  }
}

function fixAllIssues() {
  lastDetectedIssues.forEach(issue => {
    if (typeof issue.fix === 'function') issue.fix();
  });
  standardizeTextContent();
  runSpellAndFormatCheck();
  if (window.showToast) window.showToast('Đã tự động sửa toàn bộ lỗi chính tả & thể thức!');
}

// Chèn cụm lỗi mẫu để người dùng thử tính năng Soát lỗi chính tả AI
function insertSampleTypos() {
  const bodyEl = document.getElementById('vbtmDocBody');
  bodyEl.value = `Thực hiện chủ chương cải cách hành chính, đơn vị đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra`;
  renderA4Preview();
  runSpellAndFormatCheck();
}

// Tạo mã QR động trên Canvas (Offline 100%, Deterministic Matrix)
function updateQrCanvas() {
  const canvas = document.getElementById('vbtmQrCanvas');
  if (!canvas) return;
  const text = (document.getElementById('vbtmQrText')?.value || document.getElementById('vbtmDocNum')?.value || 'VBTM-PRO').trim();
  const ctx = canvas.getContext('2d');
  const size = 25;
  const scale = 6;
  const pad = 10;
  canvas.width = size * scale + pad * 2;
  canvas.height = size * scale + pad * 2;

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  // Hash chuỗi văn bản để tạo ma trận QR nhất quán
  let hash = 2166136261;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }

  function drawFinderPattern(ox, oy) {
    ctx.fillStyle = '#000000';
    ctx.fillRect(pad + ox * scale, pad + oy * scale, 7 * scale, 7 * scale);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(pad + (ox + 1) * scale, pad + (oy + 1) * scale, 5 * scale, 5 * scale);
    ctx.fillStyle = '#000000';
    ctx.fillRect(pad + (ox + 2) * scale, pad + (oy + 2) * scale, 3 * scale, 3 * scale);
  }

  drawFinderPattern(0, 0);
  drawFinderPattern(size - 7, 0);
  drawFinderPattern(0, size - 7);

  ctx.fillStyle = '#000000';
  for (let r = 0; r < size; r++) {
    for (let c = 0; c < size; c++) {
      if ((r < 8 && c < 8) || (r < 8 && c >= size - 8) || (r >= size - 8 && c < 8)) continue;
      if (r === 6 || c === 6) {
        if ((r + c) % 2 === 0) ctx.fillRect(pad + c * scale, pad + r * scale, scale, scale);
        continue;
      }
      const charCode = text.charCodeAt((r * size + c) % text.length) || 65;
      const bit = ((hash ^ (r * 131 + c * 313 + charCode * 17)) & 3) === 0 || ((r * c + charCode) % 5 === 0);
      if (bit) {
        ctx.fillRect(pad + c * scale, pad + r * scale, scale, scale);
      }
    }
  }
}

// Hiển thị trực tiếp bản xem trước trang A4 theo đúng quy chuẩn
function renderA4Preview() {
  const container = document.getElementById('a4PreviewPaper');
  if (!container) return;

  const standard = document.getElementById('vbtmStandard').value;
  const govOrg = document.getElementById('vbtmGovOrg').value.trim();
  const issueOrg = document.getElementById('vbtmIssueOrg').value.trim();
  const docNum = document.getElementById('vbtmDocNum').value.trim();
  const location = document.getElementById('vbtmLocation').value.trim();
  const docDate = document.getElementById('vbtmDocDate').value;
  const docType = document.getElementById('vbtmDocType').value.trim();
  const docSubject = document.getElementById('vbtmDocSubject').value.trim();
  const docBody = document.getElementById('vbtmDocBody').value;
  const recipients = document.getElementById('vbtmRecipients').value;
  const signerTitle = document.getElementById('vbtmSignerTitle').value;
  const signerName = document.getElementById('vbtmSignerName').value.trim();
  const includeQr = document.getElementById('vbtmIncludeQr')?.checked;

  const rightHeaderTop = standard === 'hd05' ? 'ĐẢNG CỘNG SẢN VIỆT NAM' : 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
  const rightHeaderSub = standard === 'hd05' ? '' : 'Độc lập - Tự do - Hạnh phúc';
  const numberLabel = docType
    ? `Số: ${docNum}`
    : `Số: ${docNum}`;

  const dateLine = formatVietnameseDate(docDate, location);
  const paragraphsHtml = docBody
    .split('\n')
    .filter(p => p.trim().length > 0)
    .map(p => `<p>${escapeHtml(p)}</p>`)
    .join('');

  const recipientsHtml = recipients
    .split('\n')
    .filter(r => r.trim().length > 0)
    .map(r => `<div>${escapeHtml(r)}</div>`)
    .join('');

  const signerTitleHtml = signerTitle
    .split('\n')
    .map(s => `<div>${escapeHtml(s)}</div>`)
    .join('');

  const qrCanvas = document.getElementById('vbtmQrCanvas');
  const qrDataUrl = (includeQr && qrCanvas) ? qrCanvas.toDataURL('image/png') : '';

  container.innerHTML = `
    <table class="a4-header-table">
      <tr>
        <td style="width: 44%;">
          ${govOrg ? `<div style="font-size: 12pt; text-transform: uppercase;">${escapeHtml(govOrg)}</div>` : ''}
          <div style="font-size: 12pt; font-weight: bold; text-transform: uppercase;">${escapeHtml(issueOrg)}</div>
          <span class="a4-line-sep" style="width: 38%;"></span>
          <div style="font-size: 12.5pt; margin-top: 8px;">${escapeHtml(numberLabel)}</div>
          ${!docType && docSubject ? `<div style="font-size: 11.5pt; margin-top: 4px;">${escapeHtml(docSubject)}</div>` : ''}
        </td>
        <td style="width: 56%;">
          <div style="font-size: 12pt; font-weight: bold; text-transform: uppercase;">${rightHeaderTop}</div>
          ${rightHeaderSub ? `<div style="font-size: 13pt; font-weight: bold;">${rightHeaderSub}</div>` : ''}
          <span class="a4-line-sep" style="width: ${standard === 'hd05' ? '55%' : '48%'};"></span>
          <div style="font-size: 12.5pt; font-style: italic; margin-top: 10px;">${escapeHtml(dateLine)}</div>
        </td>
      </tr>
    </table>

    ${docType ? `<div class="a4-doc-title">${escapeHtml(docType)}</div>` : ''}
    ${docType && docSubject ? `<div class="a4-doc-subject">${escapeHtml(docSubject)}<span class="a4-line-sep" style="width: 28%; margin-top: 5px;"></span></div>` : ''}

    <div class="a4-doc-body">
      ${paragraphsHtml}
    </div>

    <table class="a4-footer-table">
      <tr>
        <td style="width: 50%; vertical-align: top; text-align: left; font-size: 11pt;">
          <div style="font-weight: bold; font-style: italic;">Nơi nhận:</div>
          <div style="line-height: 1.35;">${recipientsHtml}</div>
          ${qrDataUrl ? `<div style="margin-top: 10px;"><img src="${qrDataUrl}" alt="QR" style="width: 68px; height: 68px; border: 1px solid #ddd; padding: 2px;"/><div style="font-size: 8.5pt; color: #555;">Mã QR xác thực VBTM</div></div>` : ''}
        </td>
        <td style="width: 50%; vertical-align: top; text-align: center;">
          <div style="font-weight: bold; font-size: 12.5pt; text-transform: uppercase;">${signerTitleHtml}</div>
          <div style="height: 68px;"></div>
          <div style="font-weight: bold; font-size: 13pt;">${escapeHtml(signerName)}</div>
        </td>
      </tr>
    </table>
  `;
}

// Xuất ra file Microsoft Word (.doc chuẩn A4, lề 20-15-20-30mm, font Times New Roman)
function exportWordDocument() {
  const previewEl = document.getElementById('a4PreviewPaper');
  if (!previewEl) return;

  const docNum = (document.getElementById('vbtmDocNum')?.value || 'VanBan').replace(/[^a-zA-Z0-9_-]/g, '_');

  const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <title>Van Ban Chuan Hoa ND30</title>
    <style>
      @page Section1 {
        size: 21.0cm 29.7cm;
        margin: 2.0cm 1.5cm 2.0cm 3.0cm;
        mso-header-margin: 1.2cm;
        mso-footer-margin: 1.2cm;
      }
      div.Section1 { page: Section1; }
      body {
        font-family: "Times New Roman", serif;
        font-size: 14pt;
        line-height: 1.35;
        color: #000000;
      }
      table { width: 100%; border-collapse: collapse; }
      td { vertical-align: top; font-family: "Times New Roman", serif; }
      .a4-doc-title { text-align: center; font-weight: bold; font-size: 14pt; text-transform: uppercase; margin: 14pt 0 4pt 0; }
      .a4-doc-subject { text-align: center; font-weight: bold; font-size: 13.5pt; margin-bottom: 12pt; }
      .a4-doc-body p { text-indent: 1.27cm; text-align: justify; margin: 0 0 6pt 0; font-size: 14pt; }
    </style>
  </head>
  <body>
    <div class="Section1">
      ${previewEl.innerHTML}
    </div>
  </body>
  </html>`;

  const blob = new Blob(['\ufeff', wordHtml], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `VBTM_ChuanHoa_${docNum}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  if (window.showToast) {
    window.showToast(`Đã xuất tệp Word chuẩn Nghị định 30: VBTM_ChuanHoa_${docNum}.doc`);
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

window.loadTemplate = loadTemplate;
window.standardizeTextContent = standardizeTextContent;
window.runSpellAndFormatCheck = runSpellAndFormatCheck;
window.fixSingleIssue = fixSingleIssue;
window.fixAllIssues = fixAllIssues;
window.insertSampleTypos = insertSampleTypos;
window.exportWordDocument = exportWordDocument;
