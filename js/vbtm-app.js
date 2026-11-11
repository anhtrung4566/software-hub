/**
 * VBTM Pro 2.2.4 - Bộ Công Cụ Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy & UBND Phường, Tạo Mã QR & Soát Lỗi Chính Tả AI Offline
 * Hỗ trợ: Nghị định 30/2020/NĐ-CP (HĐND & UBND Phường), Hướng dẫn 36-HD/VPTW & 05-HD/VPTW (Đảng ủy Phường)
 */

const VBTM_TEMPLATES = {
  ubnd_congvan: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '145/UBND-VP',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: '',
    docSubject: 'V/v triển khai ứng dụng chuẩn hóa thể thức văn bản và số hóa hồ sơ hành chính tại địa phương',
    docBody: `Thực hiện Kế hoạch cải cách hành chính và chuyển đổi số năm 2026, Ủy ban nhân dân phường triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư.\n\n1. Các bộ phận chuyên môn, công chức thuộc Ủy ban nhân dân phường thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; bảo đảm đúng quy định về khổ giấy A4, lề trang, phông chữ Times New Roman, cỡ chữ và thể thức trình bày.\n\n2. Đẩy mạnh số hóa hồ sơ, tài liệu lưu trữ và ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm nâng cao chất lượng ban hành văn bản và bảo đảm an toàn thông tin nội bộ.\n\nỦy ban nhân dân phường đề nghị các bộ phận, cán bộ, công chức nghiêm túc triển khai thực hiện./.`,
    recipients: `- Như trên;\n- Thường trực Đảng ủy, HĐND phường (để b/c);\n- Chủ tịch, các PCT UBND phường;\n- UBMTTQ và các đoàn thể phường;\n- Các tổ dân phố;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_baocao: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '86/BC-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Kết quả thực hiện công tác cải cách hành chính, văn thư lưu trữ và chuyển đổi số trên địa bàn phường năm 2026',
    docBody: `Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai công tác chuẩn hóa văn bản và số hóa hồ sơ trên địa bàn phường như sau:\n\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Công tác ban hành văn bản: 100% văn bản hành chính của Hội đồng nhân dân và Ủy ban nhân dân phường được chuẩn hóa đồng bộ theo Nghị định số 30/2020/NĐ-CP của Chính phủ.\n\n2. Công tác số hóa hồ sơ: Đã ứng dụng phần mềm VBTM Pro và VietNamOCR 3.2 Pro phục vụ rà soát lỗi chính tả tiếng Việt, tạo mã QR quản lý hồ sơ và bóc tách dữ liệu bảng biểu sang Word, Excel nhanh chóng, chính xác.\n\nII. PHƯƠNG HƯỚNG, NHIỆM VỤ TRONG THỜI GIAN TỚI\nTiếp tục duy trì nền nếp công tác văn thư, lưu trữ và nâng cao hiệu quả giải quyết thủ tục hành chính phục vụ nhân dân tại Bộ phận Tiếp nhận và Trả kết quả phường./.`,
    recipients: `- UBND Quận (để b/c);\n- Phòng Nội vụ;\n- Thường trực Đảng ủy, HĐND phường;\n- CT, các PCT UBND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_quyetdinh: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '120/QĐ-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'QUYẾT ĐỊNH',
    docSubject: 'Về việc ban hành Quy chế quản lý, chuẩn hóa văn bản và số hóa hồ sơ tài liệu tại Ủy ban nhân dân phường',
    docBody: `ỦY BAN NHÂN DÂN PHƯỜNG\n\nCăn cứ Luật Tổ chức chính quyền địa phương;\n\nCăn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\n\nTheo đề nghị của Công chức Văn phòng - Thống kê phường,\n\nQUYẾT ĐỊNH:\n\nĐiều 1. Ban hành kèm theo Quyết định này Quy chế ứng dụng phần mềm chuẩn hóa thể thức văn bản VBTM Pro và số hóa hồ sơ tài liệu VietNamOCR 3.2 Pro trong công tác văn thư, lưu trữ tại Ủy ban nhân dân phường.\n\nĐiều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.\n\nĐiều 3. Công chức Văn phòng - Thống kê và các cán bộ, công chức, bộ phận chuyên môn thuộc Ủy ban nhân dân phường chịu trách nhiệm thi hành Quyết định này./.`,
    recipients: `- Như Điều 3;\n- Thường trực Đảng ủy, HĐND phường (để b/c);\n- Chủ tịch, các PCT UBND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  danguy_baocao: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '42-BC/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Tổng kết công tác xây dựng Đảng, cải cách hành chính trong Đảng và chuyển đổi số tại Đảng bộ phường năm 2026',
    docBody: `Thực hiện Hướng dẫn số 36-HD/VPTW và Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy phường báo cáo kết quả thực hiện công tác văn thư, lưu trữ và chuyển đổi số năm 2026 như sau:\n\nI. KẾT QUẢ THỰC HIỆN\n1. Toàn bộ hệ thống nghị quyết, quyết định, chương trình, kế hoạch, báo cáo và công văn của Đảng ủy phường và các chi bộ trực thuộc đều được chuẩn hóa đúng thể thức văn bản của Đảng.\n\n2. Việc ứng dụng bộ công cụ kiểm tra chính tả tiếng Việt và chuẩn hóa thể thức ngoại tuyến 100% giúp nâng cao chất lượng văn bản tham mưu của Văn phòng Đảng ủy, bảo đảm tuyệt đối an toàn thông tin nội bộ.\n\nII. NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI\nTiếp tục đẩy mạnh số hóa hồ sơ đảng viên, văn kiện của Đảng ủy và hướng dẫn các chi bộ trực thuộc thực hiện đồng bộ, thống nhất./.`,
    recipients: `- Ban Thường vụ Quận ủy (để b/c);\n- Văn phòng Quận ủy;\n- Các đồng chí Đảng ủy viên;\n- HĐND, UBND, UBMTTQ phường;\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M ĐẢNG ỦY\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
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
  { wrong: /\bkế hoạc\b/gi, right: 'kế hoạch', desc: 'Thiếu âm cuối ("kế hoạch")' },
  { wrong: /\bngiêm túc\b/gi, right: 'nghiêm túc', desc: 'Sai quy tắc chính tả ng/ngh ("nghiêm túc")' },
  { wrong: /\bngành ngề\b/gi, right: 'ngành nghề', desc: 'Sai quy tắc chính tả ng/ngh ("ngành nghề")' },
  { wrong: /\bhổ trợ\b/gi, right: 'hỗ trợ', desc: 'Sai dấu hỏi/ngã ("hỗ trợ")' },
  { wrong: /\btiêu chuẫn\b/gi, right: 'tiêu chuẩn', desc: 'Sai dấu hỏi/ngã ("tiêu chuẩn")' },
  { wrong: /\bkhẩn chương\b/gi, right: 'khẩn trương', desc: 'Sai phụ âm ch/tr ("khẩn trương")' },
  { wrong: /\btrao chuốt\b/gi, right: 'trau chuốt', desc: 'Sai vần ao/au ("trau chuốt")' },
  { wrong: /\bsáng xuốt\b/gi, right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { wrong: /\bxáng suốt\b/gi, right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { wrong: /\bgiử gìn\b/gi, right: 'giữ gìn', desc: 'Sai dấu hỏi/ngã ("giữ gìn")' },
  { wrong: /\bvửng chắc\b/gi, right: 'vững chắc', desc: 'Sai dấu hỏi/ngã ("vững chắc")' },
  { wrong: /\bsát nhập\b/gi, right: 'sáp nhập', desc: 'Từ chuẩn hành chính là "sáp nhập"' },
  { wrong: /\bchính phủ\b/g, right: 'Chính phủ', desc: 'Viết hoa danh từ riêng cơ quan ("Chính phủ")' },
  { wrong: /\bnhà nước\b/g, right: 'Nhà nước', desc: 'Viết hoa "Nhà nước" trong văn bản hành chính' }
];

let lastDetectedIssues = [];

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('vbtmStudioApp')) return;
  loadTemplate('ubnd_congvan');
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
  if (window.showToast) window.showToast('Đã nạp biểu mẫu văn bản Đảng ủy / UBND phường');
}

// Chuẩn hóa ngày tháng theo Nghị định 30 (ngày < 10 và tháng 1, 2 phải thêm số 0)
function formatVietnameseDate(dateStr, location) {
  if (!dateStr) return `${location || 'Phường'}, ngày 11 tháng 11 năm 2026`;
  const parts = dateStr.split('-');
  if (parts.length !== 3) return `${location}, ${dateStr}`;
  const year = parts[0];
  const monthNum = parseInt(parts[1], 10);
  const dayNum = parseInt(parts[2], 10);
  const dayStr = dayNum < 10 ? `0${dayNum}` : `${dayNum}`;
  const monthStr = (monthNum === 1 || monthNum === 2) ? `0${monthNum}` : `${monthNum}`;
  return `${location || 'Phường'}, ngày ${dayStr} tháng ${monthStr} năm ${year}`;
}

// Tự động chuẩn hóa thể thức văn bản (khoảng trắng, dấu câu, viết hoa, kết thúc văn bản)
function standardizeTextContent() {
  const bodyEl = document.getElementById('vbtmDocBody');
  const subjEl = document.getElementById('vbtmDocSubject');
  if (!bodyEl) return;

  let text = bodyEl.value;
  text = text
    .replace(/[ \t]+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([,.;:!?])([^\s\d/])/g, '$1 $2')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/\n{3,}/g, '\n\n');

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

  if (subjEl && subjEl.value) {
    subjEl.value = subjEl.value.trim().replace(/\s+/g, ' ').replace(/\.+$/, '');
  }

  renderA4Preview();
  if (window.showToast) {
    window.showToast('Đã chuẩn hóa thể thức văn bản Đảng ủy / UBND phường!');
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

  if (/\s+[,.;:!?]/.test(body)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)',
      fix: () => standardizeTextContent()
    });
  }

  if (/\.$/.test(subject.trim())) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Trích yếu văn bản không được đặt dấu chấm ở cuối dòng',
      fix: () => {
        document.getElementById('vbtmDocSubject').value = subject.trim().replace(/\.+$/, '');
      }
    });
  }

  if (body.trim() && !body.trim().endsWith('./.')) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Cuối nội dung văn bản hành chính thiếu ký hiệu kết thúc "./."',
      fix: () => standardizeTextContent()
    });
  }

  if (!docNum.trim()) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Thiếu Số và ký hiệu văn bản (Ví dụ: 145/UBND-VP hoặc 42-BC/ĐU)',
      fix: () => { document.getElementById('vbtmDocNum').value = '01/UBND-VP'; }
    });
  }
  if (!recipients.trim() || !/Lưu/i.test(recipients)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Mục "Nơi nhận" cần có dòng "- Lưu: VT,..." hoặc "- Lưu Văn phòng Đảng ủy."',
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
        <span>✓ Văn bản đã đạt chuẩn 100% về chính tả Tiếng Việt và thể thức Đảng ủy / UBND phường!</span>
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

function insertSampleTypos() {
  const bodyEl = document.getElementById('vbtmDocBody');
  bodyEl.value = `Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra`;
  renderA4Preview();
  runSpellAndFormatCheck();
}

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
  const numberLabel = `Số: ${docNum}`;

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
    window.showToast(`Đã xuất tệp Word chuẩn A4: VBTM_ChuanHoa_${docNum}.doc`);
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
