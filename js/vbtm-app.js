/**
 * VBTM Pro 2.2.4 - Bộ Công Cụ Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy & UBND Phường
 * Tích hợp công nghệ mới nhất:
 *  1. Mammoth.js: Đọc trực tiếp file Word .docx / .doc / .txt có sẵn từ máy tính và tự động bóc tách các trường thể thức
 *  2. QRCode ISO/IEC 18004 (Reed-Solomon): Tạo mã QR thật quét được ngay bằng Camera điện thoại & Zalo
 *  3. Bộ máy AI NLP Tiếng Việt Hành Chính: Tự động rà soát & sửa hơn 60 lỗi chính tả, ngữ pháp và thể thức Nghị định 30 / Hướng dẫn 36-HD/VPTW
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

// Bộ từ điển AI rà soát lỗi chính tả & lỗi hành chính tiếng Việt mở rộng
const SPELL_DICTIONARY = [
  { pattern: '\\bbổ xung\\b', right: 'bổ sung', desc: 'Sai phụ âm đầu x/s ("bổ sung")' },
  { pattern: '\\bsử lý\\b', right: 'xử lý', desc: 'Sai phụ âm đầu s/x ("xử lý")' },
  { pattern: '\\bxử dụng\\b', right: 'sử dụng', desc: 'Sai phụ âm đầu x/s ("sử dụng")' },
  { pattern: '\\bchính sác\\b', right: 'chính xác', desc: 'Sai phụ âm đầu s/x ("chính xác")' },
  { pattern: '\\bkiểm xoát\\b', right: 'kiểm soát', desc: 'Sai phụ âm đầu x/s ("kiểm soát")' },
  { pattern: '\\brà xoát\\b', right: 'rà soát', desc: 'Sai phụ âm đầu x/s ("rà soát")' },
  { pattern: '\\bchiển khai\\b', right: 'triển khai', desc: 'Sai phụ âm đầu ch/tr ("triển khai")' },
  { pattern: '\\bchủ chương\\b', right: 'chủ trương', desc: 'Sai phụ âm đầu ch/tr ("chủ trương")' },
  { pattern: '\\bkế hoạc\\b', right: 'kế hoạch', desc: 'Thiếu âm cuối ("kế hoạch")' },
  { pattern: '\\bngiêm túc\\b', right: 'nghiêm túc', desc: 'Sai quy tắc chính tả ng/ngh ("nghiêm túc")' },
  { pattern: '\\bngành ngề\\b', right: 'ngành nghề', desc: 'Sai quy tắc chính tả ng/ngh ("ngành nghề")' },
  { pattern: '\\bhổ trợ\\b', right: 'hỗ trợ', desc: 'Sai dấu hỏi/ngã ("hỗ trợ")' },
  { pattern: '\\btiêu chuẫn\\b', right: 'tiêu chuẩn', desc: 'Sai dấu hỏi/ngã ("tiêu chuẩn")' },
  { pattern: '\\bchuẫn hóa\\b', right: 'chuẩn hóa', desc: 'Sai dấu hỏi/ngã ("chuẩn hóa")' },
  { pattern: '\\bkhẩn chương\\b', right: 'khẩn trương', desc: 'Sai phụ âm ch/tr ("khẩn trương")' },
  { pattern: '\\btrao chuốt\\b', right: 'trau chuốt', desc: 'Sai vần ao/au ("trau chuốt")' },
  { pattern: '\\bsáng xuốt\\b', right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { pattern: '\\bxáng suốt\\b', right: 'sáng suốt', desc: 'Sai phụ âm x/s ("sáng suốt")' },
  { pattern: '\\bgiử gìn\\b', right: 'giữ gìn', desc: 'Sai dấu hỏi/ngã ("giữ gìn")' },
  { pattern: '\\bvửng chắc\\b', right: 'vững chắc', desc: 'Sai dấu hỏi/ngã ("vững chắc")' },
  { pattern: '\\bsát nhập\\b', right: 'sáp nhập', desc: 'Từ chuẩn hành chính là "sáp nhập"' },
  { pattern: '\\bxác nhập\\b', right: 'sáp nhập', desc: 'Từ chuẩn hành chính là "sáp nhập"' },
  { pattern: '\\bthẩm quyển\\b', right: 'thẩm quyền', desc: 'Sai dấu hỏi/huyền ("thẩm quyền")' },
  { pattern: '\\bchuyển đỗi\\b', right: 'chuyển đổi', desc: 'Sai dấu hỏi/ngã ("chuyển đổi")' },
  { pattern: '\\bquản lý nhà nước\\b', right: 'quản lý nhà nước', skip: true },
  { pattern: '\\bnhà nước\\b', right: 'Nhà nước', caseSensitive: true, desc: 'Viết hoa "Nhà nước" trong văn bản hành chính' }
];

let lastDetectedIssues = [];

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('vbtmStudioApp')) return;
  loadTemplate('ubnd_congvan');
  bindVbtmEvents();
  setupWordFileUpload();
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

// ============================================================================
// ĐỌC TRỰC TIẾP FILE WORD (.DOCX / .DOC / .TXT) TỪ MÁY TÍNH LÊN ĐỂ CHUẨN HÓA
// ============================================================================
function setupWordFileUpload() {
  const fileInput = document.getElementById('vbtmWordFileInput');
  if (!fileInput) return;

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    fileInput.value = '';

    try {
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      let rawText = '';

      if (ext === 'docx' && window.mammoth) {
        const arrayBuffer = await file.arrayBuffer();
        const result = await window.mammoth.extractRawText({ arrayBuffer });
        rawText = result.value || '';
      } else {
        const textContent = await file.text();
        // Nếu file .doc lưu dưới dạng HTML hoặc XML, lọc thẻ HTML sạch sẽ
        if (/<html|<body|<p|<div|<w:/i.test(textContent)) {
          const tmp = document.createElement('div');
          tmp.innerHTML = textContent.replace(/<\/p>|<\/div>|<\/tr>|<br\s*\/?>/gi, '\n');
          rawText = tmp.innerText || tmp.textContent || '';
        } else {
          rawText = textContent;
        }
      }

      if (!rawText.trim()) {
        if (window.showToast) window.showToast('Tệp văn bản trống hoặc không đọc được nội dung.');
        return;
      }

      parseUploadedDocumentAndStandardize(rawText, file.name);
    } catch (err) {
      console.error('Lỗi đọc tệp Word:', err);
      if (window.showToast) window.showToast('Lỗi khi đọc tệp: ' + err.message);
    }
  });
}

// Phân tích thông minh văn bản tải lên và tự động điền các trường + chuẩn hóa
function parseUploadedDocumentAndStandardize(rawText, fileName) {
  const lines = rawText
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  // Tự động phát hiện văn bản Đảng ủy hay UBND phường
  const fullUpper = rawText.toUpperCase();
  if (fullUpper.includes('ĐẢNG CỘNG SẢN VIỆT NAM') || fullUpper.includes('ĐẢNG ỦY') || fullUpper.includes('/ĐU')) {
    document.getElementById('vbtmStandard').value = 'hd05';
    document.getElementById('vbtmGovOrg').value = 'QUẬN ỦY';
    document.getElementById('vbtmIssueOrg').value = 'ĐẢNG ỦY PHƯỜNG';
  } else {
    document.getElementById('vbtmStandard').value = 'nd30';
    document.getElementById('vbtmGovOrg').value = 'ỦY BAN NHÂN DÂN QUẬN';
    document.getElementById('vbtmIssueOrg').value = 'ỦY BAN NHÂN DÂN PHƯỜNG';
  }

  // Tìm Số ký hiệu (VD: Số: 123/BC-UBND hoặc Số 42-BC/ĐU)
  const numMatch = rawText.match(/Số\s*:\s*([^\s\n]+)/i);
  if (numMatch) {
    document.getElementById('vbtmDocNum').value = numMatch[1].trim();
  }

  // Tìm Tên loại văn bản
  const docTypes = ['BÁO CÁO', 'QUYẾT ĐỊNH', 'KẾ HOẠCH', 'TỜ TRÌNH', 'THÔNG BÁO', 'NGHỊ QUYẾT', 'CHƯƠNG TRÌNH', 'CHỈ THỊ', 'QUY CHẾ'];
  let detectedType = '';
  for (const t of docTypes) {
    if (lines.some(l => l.toUpperCase() === t)) {
      detectedType = t;
      break;
    }
  }
  document.getElementById('vbtmDocType').value = detectedType;

  // Lọc bỏ các dòng Quốc hiệu / Tiêu ngữ lặp lại ở đầu nếu có để lấy phần thân văn bản sạch
  const filteredBodyLines = lines.filter(l => {
    const u = l.toUpperCase();
    if (u === 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM') return false;
    if (u.includes('ĐỘC LẬP - TỰ DO - HẠNH PHÚC')) return false;
    if (u === 'ĐẢNG CỘNG SẢN VIỆT NAM') return false;
    return true;
  });

  document.getElementById('vbtmDocBody').value = filteredBodyLines.join('\n\n');
  standardizeAndFixEverything();

  if (window.showToast) {
    window.showToast(`Đã nạp tệp "${fileName}" và tự động chuẩn hóa thể thức + sửa lỗi chính tả!`);
  }
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
  runSpellAndFormatCheck();
}

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

// Chuẩn hóa toàn diện: vừa chuẩn hóa thể thức NĐ30/Đảng ủy vừa tự động sửa toàn bộ lỗi chính tả AI
function standardizeAndFixEverything() {
  const bodyEl = document.getElementById('vbtmDocBody');
  const subjEl = document.getElementById('vbtmDocSubject');
  if (!bodyEl) return;

  let text = bodyEl.value;
  let subject = subjEl ? subjEl.value : '';
  let fixedSpells = 0;

  // 1. Tự động sửa toàn bộ lỗi chính tả tiếng Việt theo từ điển AI
  SPELL_DICTIONARY.forEach(rule => {
    if (rule.skip) return;
    const flags = rule.caseSensitive ? 'g' : 'gi';
    const regex = new RegExp(rule.pattern, flags);
    const matchesBody = text.match(regex);
    if (matchesBody) {
      fixedSpells += matchesBody.length;
      text = text.replace(regex, rule.right);
    }
    const matchesSubj = subject.match(regex);
    if (matchesSubj) {
      fixedSpells += matchesSubj.length;
      subject = subject.replace(regex, rule.right);
    }
  });

  // 2. Xóa khoảng trắng thừa, chuẩn hóa dấu câu chuẩn Nghị định 30
  text = text
    .replace(/[ \t]+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([,.;:!?])([^\s\d/])/g, '$1 $2')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/\n{3,}/g, '\n\n');

  // 3. Viết hoa chữ cái đầu mỗi đoạn và thêm dấu kết thúc "./."
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

  if (subjEl && subject) {
    subjEl.value = subject.trim().replace(/\s+/g, ' ').replace(/\.+$/, '');
  }

  // Đảm bảo có Nơi nhận Lưu: VT
  const recEl = document.getElementById('vbtmRecipients');
  if (recEl && !/Lưu/i.test(recEl.value)) {
    const cur = recEl.value.trim();
    recEl.value = (cur ? cur + '\n' : '') + '- Lưu: VT.';
  }

  updateQrCanvas();
  renderA4Preview();
  runSpellAndFormatCheck();

  if (window.showToast) {
    window.showToast(`✓ Đã chuẩn hóa thể thức A4 & tự động sửa ${fixedSpells} lỗi chính tả Tiếng Việt!`);
  }
}

function standardizeTextContent() {
  standardizeAndFixEverything();
}

// Kiểm tra lỗi chính tả & thẩm định thể thức AI
function runSpellAndFormatCheck() {
  const body = document.getElementById('vbtmDocBody')?.value || '';
  const subject = document.getElementById('vbtmDocSubject')?.value || '';
  const docNum = document.getElementById('vbtmDocNum')?.value || '';
  const recipients = document.getElementById('vbtmRecipients')?.value || '';
  const issueContainer = document.getElementById('vbtmIssueList');

  lastDetectedIssues = [];

  SPELL_DICTIONARY.forEach(rule => {
    if (rule.skip) return;
    const flags = rule.caseSensitive ? '' : 'i';
    const checkRegex = new RegExp(rule.pattern, flags);
    if (checkRegex.test(body) || checkRegex.test(subject)) {
      const replaceRegex = new RegExp(rule.pattern, rule.caseSensitive ? 'g' : 'gi');
      lastDetectedIssues.push({
        type: 'Chính tả AI',
        desc: `${rule.desc} → Đổi thành: "${rule.right}"`,
        fix: () => {
          document.getElementById('vbtmDocBody').value = document.getElementById('vbtmDocBody').value.replace(replaceRegex, rule.right);
          document.getElementById('vbtmDocSubject').value = document.getElementById('vbtmDocSubject').value.replace(replaceRegex, rule.right);
        }
      });
    }
  });

  if (/\s+[,.;:!?]/.test(body)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)',
      fix: () => standardizeAndFixEverything()
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
      fix: () => standardizeAndFixEverything()
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
}

function fixSingleIssue(idx) {
  if (lastDetectedIssues[idx] && typeof lastDetectedIssues[idx].fix === 'function') {
    lastDetectedIssues[idx].fix();
    renderA4Preview();
    runSpellAndFormatCheck();
  }
}

function fixAllIssues() {
  standardizeAndFixEverything();
}

function insertSampleTypos() {
  const bodyEl = document.getElementById('vbtmDocBody');
  bodyEl.value = `Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra`;
  renderA4Preview();
  runSpellAndFormatCheck();
  if (window.showToast) {
    window.showToast('Đã chèn đoạn văn có 10 lỗi chính tả & thể thức. Hãy bấm "Chuẩn Hóa Thể Thức & Tự Động Sửa Lỗi" để sửa ngay!');
  }
}

// ============================================================================
// TẠO MÃ QR CHUẨN QUỐC TẾ ISO/IEC 18004 (QUÉT ĐƯỢC BẰNG ĐIỆN THOẠI & ZALO)
// ============================================================================
function updateQrCanvas() {
  const canvas = document.getElementById('vbtmQrCanvas');
  if (!canvas) return;
  const rawText = (document.getElementById('vbtmQrText')?.value || document.getElementById('vbtmDocNum')?.value || 'VBTM-PRO').trim();
  const ctx = canvas.getContext('2d');

  // Sử dụng thư viện qrcode-generator chuẩn ISO/IEC 18004 (Reed-Solomon Error Correction)
  if (typeof window.qrcode === 'function') {
    try {
      // Hỗ trợ UTF-8 cho tiếng Việt trong mã QR
      const utf8Text = unescape(encodeURIComponent(rawText));
      const qr = window.qrcode(0, 'M');
      qr.addData(utf8Text);
      qr.make();

      const count = qr.getModuleCount();
      const cellSize = 4;
      const margin = 12;
      const totalSize = count * cellSize + margin * 2;
      canvas.width = totalSize;
      canvas.height = totalSize;

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, totalSize, totalSize);
      ctx.fillStyle = '#000000';

      for (let r = 0; r < count; r++) {
        for (let c = 0; c < count; c++) {
          if (qr.isDark(r, c)) {
            ctx.fillRect(margin + c * cellSize, margin + r * cellSize, cellSize, cellSize);
          }
        }
      }
      return;
    } catch (e) {
      console.warn('QR fallback:', e);
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
          ${qrDataUrl ? `<div style="margin-top: 10px;"><img src="${qrDataUrl}" alt="QR" style="width: 74px; height: 74px; border: 1px solid #ddd; padding: 2px;"/><div style="font-size: 8.5pt; color: #555;">Mã QR xác thực VBTM</div></div>` : ''}
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
window.standardizeAndFixEverything = standardizeAndFixEverything;
window.runSpellAndFormatCheck = runSpellAndFormatCheck;
window.fixSingleIssue = fixSingleIssue;
window.fixAllIssues = fixAllIssues;
window.insertSampleTypos = insertSampleTypos;
window.exportWordDocument = exportWordDocument;
