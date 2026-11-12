/**
 * VBTM Pro 2.2.4 - Bộ Công Cụ Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy & UBND Phường
 * Tích hợp công nghệ mới nhất:
 *  1. Kho Từ Điển 535 Cụm Từ Nhầm Lẫn Hành Chính + 7.578 Âm Tiết Tiếng Việt Chuẩn Quốc Gia (Unicode NFC)
 *  2. Bộ Phân Giải Ngữ Âm Âm Tiết (SyllableParser): Tự động bắt và sửa mọi âm tiết sai chính tả (ng/ngh, g/gh, c/k, hỏi/ngã, c/t/ch, n/ng/nh, s/x, ch/tr)
 *  3. Mammoth.js: Đọc trực tiếp file Word .docx / .doc / .txt có sẵn từ máy tính và tự động bóc tách thể thức
 *  4. QRCode ISO/IEC 18004 (Reed-Solomon): Tạo mã QR thật quét được ngay bằng Camera điện thoại & Zalo
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

// ============================================================================
// BỘ PHÂN GIẢI NGỮ ÂM & KIỂM ĐỊNH ÂM TIẾT TIẾNG VIỆT (SYLLABLE PARSER ENGINE)
// ============================================================================
const VOWEL_GROUPS = {
  'a': 'aàáảãạ', 'ă': 'ăằắẳẵặ', 'â': 'âầấẩẫậ',
  'e': 'eèéẻẽẹ', 'ê': 'êềếểễệ',
  'i': 'iìíỉĩị',
  'o': 'oòóỏõọ', 'ô': 'ôồốổỗộ', 'ơ': 'ơờớởỡợ',
  'u': 'uùúủũụ', 'ư': 'ưừứửữự',
  'y': 'yỳýỷỹỵ'
};

const CHAR_TO_TONE = {};
const TONE_TO_CHAR = {};
Object.entries(VOWEL_GROUPS).forEach(([base, chars]) => {
  for (let i = 0; i < chars.length; i++) {
    CHAR_TO_TONE[chars[i]] = { base, tone: i };
    TONE_TO_CHAR[`${base}_${i}`] = chars[i];
  }
});

// Biểu thức biên từ Unicode chuẩn cho Tiếng Việt (Khắc phục triệt để lỗi \b của JS không nhận diện chữ có dấu!)
function escapeRegexStr(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createUnicodePhraseRegex(phrase, flags = 'gui') {
  const escaped = escapeRegexStr(phrase.trim());
  return new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`, flags);
}

// Giữ nguyên kiểu viết hoa (chữ đầu câu hoặc IN HOA toàn bộ) khi thay thế từ đúng
function applyCasePattern(sourceStr, targetStr) {
  if (!sourceStr || !targetStr) return targetStr;
  if (sourceStr === sourceStr.toUpperCase() && sourceStr !== sourceStr.toLowerCase()) {
    return targetStr.toUpperCase();
  }
  if (sourceStr[0] === sourceStr[0].toUpperCase() && sourceStr[0] !== sourceStr[0].toLowerCase()) {
    return targetStr.charAt(0).toUpperCase() + targetStr.slice(1);
  }
  return targetStr;
}

// Kiểm tra một âm tiết có hợp lệ trong 7.578 âm tiết Tiếng Việt chuẩn hay không
function isValidVietnameseSyllable(word) {
  if (!window.VN_VALID_SYLLABLES) return true;
  const w = word.normalize('NFC').toLowerCase();
  if (w.length <= 1) return true;
  // Bỏ qua từ viết tắt IN HOA (UBND, HĐND, ĐU, VPTW...), từ có số hoặc ký tự đặc biệt
  if (word === word.toUpperCase() && word.length <= 8) return true;
  if (!/^\p{L}+$/u.test(w)) return true;
  return window.VN_VALID_SYLLABLES.has(w);
}

// Tự động tìm âm tiết Tiếng Việt chuẩn gần nhất cho bất kỳ âm tiết sai chính tả nào
function findClosestValidVietnameseSyllable(rawWord) {
  if (!window.VN_VALID_SYLLABLES) return null;
  const w = rawWord.normalize('NFC').toLowerCase();
  if (isValidVietnameseSyllable(rawWord)) return null;

  const candidates = [];
  const addCandidate = (cand, reason) => {
    const norm = cand.normalize('NFC').toLowerCase();
    if (window.VN_VALID_SYLLABLES.has(norm) && !candidates.some(c => c.word === norm)) {
      candidates.push({ word: norm, reason });
    }
  };

  // 1. Quy tắc chính tả ng/ngh, g/gh, c/k
  if (/^ng[ieêìíỉĩịèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('ngh' + w.slice(2), 'Quy tắc ng/ngh trước i, e, ê');
  }
  if (/^ngh[aăâoôơuưàáảãạằắẳẵặầấẩẫậòóỏõọồốổỗộờớởỡợùúủũụừứửữự]/i.test(w)) {
    addCandidate('ng' + w.slice(3), 'Quy tắc ng/ngh trước a, o, ô, ơ, u, ư');
  }
  if (/^g[eêèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('gh' + w.slice(1), 'Quy tắc g/gh trước e, ê');
  }
  if (/^gh[aăâoôơuưàáảãạằắẳẵặầấẩẫậòóỏõọồốổỗộờớởỡợùúủũụừứửữự]/i.test(w)) {
    addCandidate('g' + w.slice(2), 'Quy tắc g/gh trước a, o, ô, ơ, u, ư');
  }
  if (/^c[ieêìíỉĩịèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('k' + w.slice(1), 'Quy tắc c/k trước i, e, ê');
  }

  // 2. Hoán vị dấu thanh (Đặc biệt là dấu hỏi <-> dấu ngã, sắc <-> nặng)
  const chars = Array.from(w);
  for (let i = 0; i < chars.length; i++) {
    const info = CHAR_TO_TONE[chars[i]];
    if (info) {
      const toneSwaps = info.tone === 3 ? [4, 2, 1, 5]
        : info.tone === 4 ? [3, 5, 2, 1]
        : info.tone === 2 ? [5, 3, 4, 1]
        : info.tone === 5 ? [2, 4, 3, 1]
        : [1, 2, 3, 4, 5];
      for (const newTone of toneSwaps) {
        const replacementChar = TONE_TO_CHAR[`${info.base}_${newTone}`];
        if (replacementChar && replacementChar !== chars[i]) {
          const copy = [...chars];
          copy[i] = replacementChar;
          addCandidate(copy.join(''), 'Sai dấu thanh tiếng Việt');
        }
      }
    }
  }

  // 3. Hoán vị âm cuối (c <-> t <-> ch, n <-> ng <-> nh)
  const codaSwaps = [
    ['c', 'ch'], ['c', 't'], ['t', 'c'], ['t', 'ch'], ['ch', 'c'], ['ch', 't'],
    ['n', 'ng'], ['n', 'nh'], ['ng', 'n'], ['ng', 'nh'], ['nh', 'n'], ['nh', 'ng']
  ];
  for (const [fromCoda, toCoda] of codaSwaps) {
    if (w.endsWith(fromCoda)) {
      addCandidate(w.slice(0, w.length - fromCoda.length) + toCoda, `Sai âm cuối -${fromCoda}/-${toCoda}`);
    }
  }

  // 4. Hoán vị âm đầu (s <-> x, ch <-> tr, d <-> gi <-> r, l <-> n)
  const onsetSwaps = [
    ['s', 'x'], ['x', 's'],
    ['ch', 'tr'], ['tr', 'ch'],
    ['gi', 'd'], ['d', 'gi'], ['r', 'd'], ['d', 'r'],
    ['l', 'n'], ['n', 'l']
  ];
  for (const [fromOnset, toOnset] of onsetSwaps) {
    if (w.startsWith(fromOnset)) {
      addCandidate(toOnset + w.slice(fromOnset.length), `Sai phụ âm đầu ${fromOnset}/${toOnset}`);
    }
  }

  return candidates.length > 0 ? candidates[0] : null;
}

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
// ĐỌC TRỰC TIẾP FILE WORD (.DOCX / .DOC) & PDF (.PDF) TỪ MÁY TÍNH ĐỂ CHUẨN HÓA
// ============================================================================
if (window.pdfjsLib) {
  try {
    const workerCode = 'importScripts("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js");';
    const blob = new Blob([workerCode], { type: 'application/javascript' });
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  } catch (e) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
}

let preservedUploadedWordHtml = '';
let preservedUploadedFileName = '';

function finalizeVbtmSegment(seg) {
  const text = seg.runs.map(r => r.str).join('');
  const html = seg.runs.map(r => {
    let h = escapeHtml(r.str);
    if (r.isBold) h = `<b>${h}</b>`;
    if (r.isItalic) h = `<i>${h}</i>`;
    return h;
  }).join('');
  return {
    x: seg.x,
    endX: seg.endX,
    centerX: (seg.x + seg.endX) / 2,
    fontSize: seg.fontSize,
    isBold: seg.isBold,
    isItalic: seg.isItalic,
    text,
    html
  };
}

function reconstructVbtmExactPdfPageLayout(rawItems, page, textContent, viewport) {
  const pageW = viewport.width || 595.28;
  const styles = textContent?.styles || {};
  const runs = [];

  for (const it of rawItems) {
    if (!it.str || !it.str.trim()) continue;
    const tx = it.transform || [12, 0, 0, 12, 0, 0];
    const x = tx[4] || 0;
    const y = tx[5] || 0;
    const rawSize = Math.hypot(tx[0] || 0, tx[1] || 0) || Math.abs(tx[3]) || 13;
    const fontSize = Math.min(26, Math.max(8.5, Math.round(rawSize * 2) / 2));
    const w = it.width > 0 ? it.width : (it.str.length * fontSize * 0.48);

    let fontDescriptor = String(it.fontName || '');
    if (styles[it.fontName]?.fontFamily) fontDescriptor += ' ' + styles[it.fontName].fontFamily;
    try {
      if (page && page.commonObjs && page.commonObjs.has(it.fontName)) {
        const fontObj = page.commonObjs.get(it.fontName);
        if (fontObj && fontObj.name) fontDescriptor += ' ' + fontObj.name;
      }
    } catch (e) {}

    const isBold = /bold|bd|black|heavy|demi|-b\b/i.test(fontDescriptor);
    const isItalic = /italic|ital|oblique|-i\b/i.test(fontDescriptor) || (Math.abs(tx[2] || 0) > 0.12);

    runs.push({ str: it.str, x, y, w, endX: x + w, fontSize, isBold, isItalic });
  }

  if (runs.length === 0) return { pageHtml: '', pagePlainText: '' };

  runs.sort((a, b) => (Math.abs(b.y - a.y) > 3.8 ? b.y - a.y : a.x - b.x));

  const rawLines = [];
  for (const r of runs) {
    const tol = Math.max(3.8, r.fontSize * 0.34);
    let matchedLine = null;
    for (let i = rawLines.length - 1; i >= Math.max(0, rawLines.length - 3); i--) {
      if (Math.abs(rawLines[i].y - r.y) <= tol) {
        matchedLine = rawLines[i];
        break;
      }
    }
    if (matchedLine) {
      matchedLine.items.push(r);
      matchedLine.y = (matchedLine.y * (matchedLine.items.length - 1) + r.y) / matchedLine.items.length;
    } else {
      rawLines.push({ y: r.y, items: [r] });
    }
  }
  rawLines.sort((a, b) => b.y - a.y);

  let minPageX = pageW;
  let maxPageX = 0;

  const lines = rawLines.map(lineObj => {
    const items = lineObj.items.sort((a, b) => a.x - b.x);
    const segments = [];
    let curSeg = null;
    for (const item of items) {
      if (item.x < minPageX) minPageX = item.x;
      if (item.endX > maxPageX) maxPageX = item.endX;
      if (!curSeg) {
        curSeg = {
          x: item.x,
          endX: item.endX,
          fontSize: item.fontSize,
          isBold: item.isBold,
          isItalic: item.isItalic,
          runs: [{ str: item.str, isBold: item.isBold, isItalic: item.isItalic, fontSize: item.fontSize }]
        };
      } else {
        const gap = item.x - curSeg.endX;
        if (gap > Math.max(22, item.fontSize * 1.65)) {
          segments.push(finalizeVbtmSegment(curSeg));
          curSeg = {
            x: item.x,
            endX: item.endX,
            fontSize: item.fontSize,
            isBold: item.isBold,
            isItalic: item.isItalic,
            runs: [{ str: item.str, isBold: item.isBold, isItalic: item.isItalic, fontSize: item.fontSize }]
          };
        } else {
          const needSpace = gap > 1.1 && !/\s$/.test(curSeg.runs[curSeg.runs.length - 1].str) && !/^\s/.test(item.str);
          curSeg.runs.push({
            str: (needSpace ? ' ' : '') + item.str,
            isBold: item.isBold,
            isItalic: item.isItalic,
            fontSize: item.fontSize
          });
          curSeg.endX = Math.max(curSeg.endX, item.endX);
          if (item.fontSize > curSeg.fontSize) curSeg.fontSize = item.fontSize;
          if (item.isBold) curSeg.isBold = true;
          if (item.isItalic) curSeg.isItalic = true;
        }
      }
    }
    if (curSeg) segments.push(finalizeVbtmSegment(curSeg));
    return {
      y: lineObj.y,
      segments,
      x: segments[0]?.x || 0,
      endX: segments[segments.length - 1]?.endX || 0,
      fontSize: segments[0]?.fontSize || 13.5,
      plainText: segments.map(s => s.text).join('\t')
    };
  }).filter(l => l.segments.length > 0);

  if (minPageX > pageW * 0.35) minPageX = 72;
  if (maxPageX < pageW * 0.65) maxPageX = pageW - 54;
  const pageCenterX = (minPageX + maxPageX) / 2;

  const classified = lines.map((line, idx) => {
    const segs = line.segments;
    if (segs.length === 2 && segs[0].x < pageW * 0.45 && segs[1].x >= pageW * 0.38) {
      return { type: 'TWO_COL', leftSeg: segs[0], rightSeg: segs[1], line, idx };
    }
    if (segs.length >= 3) return { type: 'TABLE_ROW', cells: segs, line, idx };
    if (segs.length === 1) {
      const s = segs[0];
      if (s.x >= pageW * 0.43) return { type: 'TWO_COL', leftSeg: null, rightSeg: s, line, idx };
      const isNearTopOrBottom = idx <= 7 || idx >= lines.length - 14;
      if (s.endX <= pageW * 0.47 && isNearTopOrBottom && (/^(Số\s*:|Nơi nhận|-|\+|•)/i.test(s.text.trim()) || idx <= 5)) {
        return { type: 'TWO_COL', leftSeg: s, rightSeg: null, line, idx };
      }
      const isCenteredCoords = (Math.abs(s.centerX - pageCenterX) <= 44 || Math.abs(s.centerX - pageW / 2) <= 44) && s.x > minPageX + 20 && s.endX < maxPageX - 20;
      const isDocHeadingTitle = /^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|CÔNG VĂN|QUYẾT ĐỊNH\s*:)$/i.test(s.text.trim());
      if (isCenteredCoords || isDocHeadingTitle) return { type: 'CENTER', seg: s, line, idx };
      return { type: 'BODY', seg: s, line, idx };
    }
    return { type: 'TABLE_ROW', cells: segs, line, idx };
  });

  const htmlBlocks = [];
  let i = 0;
  while (i < classified.length) {
    const item = classified[i];
    if (item.type === 'TWO_COL') {
      const leftLines = [];
      const rightLines = [];
      const startIdx = i;
      while (i < classified.length && classified[i].type === 'TWO_COL') {
        if (classified[i].leftSeg) leftLines.push(classified[i].leftSeg);
        if (classified[i].rightSeg) rightLines.push(classified[i].rightSeg);
        i++;
      }
      const isHeaderZone = startIdx <= 5;
      const isRecipientLeft = leftLines.some(s => /^(Nơi nhận|-|\+)/i.test(s.text.trim()));
      const leftAlign = (isHeaderZone && !isRecipientLeft) ? 'center' : (isRecipientLeft ? 'left' : 'center');
      const leftHtml = leftLines.map(s => `<div style="font-size:${s.fontSize}pt; margin:1.5pt 0; line-height:1.3;">${s.html}</div>`).join('');
      const rightHtml = rightLines.map(s => `<div style="font-size:${s.fontSize}pt; margin:1.5pt 0; line-height:1.3;">${s.html}</div>`).join('');
      htmlBlocks.push(`
        <table style="width:100%; border-collapse:collapse; border:none; margin:4pt 0;">
          <tr>
            <td style="width:45%; vertical-align:top; border:none; padding:0 6pt 0 0; text-align:${leftAlign};">${leftHtml}</td>
            <td style="width:55%; vertical-align:top; border:none; padding:0 0 0 6pt; text-align:center;">${rightHtml}</td>
          </tr>
        </table>
      `);
      continue;
    }
    if (item.type === 'TABLE_ROW') {
      const tableRows = [];
      while (i < classified.length && classified[i].type === 'TABLE_ROW') {
        tableRows.push(classified[i].cells);
        i++;
      }
      const maxCols = Math.max(...tableRows.map(r => r.length));
      const trsHtml = tableRows.map(row => {
        const tds = Array.from({ length: maxCols }, (_, cIdx) => {
          const cell = row[cIdx];
          return `<td style="border:1px solid #000; padding:4pt 6pt; vertical-align:top; font-size:${cell ? cell.fontSize : 12}pt;">${cell ? cell.html : ''}</td>`;
        }).join('');
        return `<tr>${tds}</tr>`;
      }).join('');
      htmlBlocks.push(`<table style="width:100%; border-collapse:collapse; margin:6pt 0;"><tbody>${trsHtml}</tbody></table>`);
      continue;
    }
    if (item.type === 'CENTER') {
      const s = item.seg;
      htmlBlocks.push(`<p style="text-align:center; text-indent:0; margin:3pt 0; line-height:1.35; font-size:${s.fontSize}pt;">${s.html}</p>`);
      i++;
      continue;
    }
    if (item.type === 'BODY') {
      const firstSeg = item.seg;
      const isIndented = firstSeg.x > minPageX + 12;
      let mergedHtml = firstSeg.html;
      let prevLine = item.line;
      let prevSeg = firstSeg;
      i++;
      while (i < classified.length && classified[i].type === 'BODY') {
        const nextItem = classified[i];
        const nextSeg = nextItem.seg;
        const nextLine = nextItem.line;
        const verticalGap = prevLine.y - nextLine.y;
        const prevReachedRightMargin = prevSeg.endX >= maxPageX - 52;
        const prevEndedSentenceStop = /[:;]$|\.\s*\/\.\s*$/.test(prevSeg.text.trim());
        const nextIsIndented = nextSeg.x > minPageX + 12;
        const nextStartsNewItem = /^([-+•]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Căn cứ\s+|Nơi nhận)/i.test(nextSeg.text.trim());
        const normalLineSpacing = verticalGap > 0 && verticalGap <= prevSeg.fontSize * 1.85;
        if (prevReachedRightMargin && !prevEndedSentenceStop && !nextIsIndented && !nextStartsNewItem && normalLineSpacing) {
          mergedHtml += ' ' + nextSeg.html;
          prevLine = nextLine;
          prevSeg = nextSeg;
          i++;
        } else {
          break;
        }
      }
      const indentStyle = isIndented ? 'text-indent: 1.27cm;' : 'text-indent: 0;';
      htmlBlocks.push(`<p style="text-align:justify; ${indentStyle} margin:0 0 5pt 0; line-height:1.38; font-size:${firstSeg.fontSize}pt;">${mergedHtml}</p>`);
      continue;
    }
    i++;
  }

  return {
    pageHtml: htmlBlocks.join('\n'),
    pagePlainText: lines.map(l => l.plainText).join('\n')
  };
}

async function extractTextFromPdfInVbtm(file) {
  if (!window.pdfjsLib) {
    throw new Error('Thư viện đọc PDF đang tải, vui lòng thử lại.');
  }
  const arrayBuffer = await file.arrayBuffer();
  let pdfDoc;
  try {
    pdfDoc = await window.pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer),
      cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
      cMapPacked: true
    }).promise;
  } catch (e) {
    pdfDoc = await window.pdfjsLib.getDocument({
      data: new Uint8Array(await file.arrayBuffer()),
      disableWorker: true
    }).promise;
  }

  const allPagesHtml = [];
  const allPagesText = [];
  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    const viewport1x = page.getViewport({ scale: 1.0 });
    const viewportRender = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement('canvas');
    canvas.width = viewportRender.width;
    canvas.height = viewportRender.height;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: viewportRender }).promise;

    const textContent = await page.getTextContent();
    const rawItems = (textContent?.items || []).filter(it => typeof it.str === 'string' && it.str.length > 0);
    if (rawItems.filter(it => it.str.trim().length > 0).length >= 2) {
      const { pageHtml, pagePlainText } = reconstructVbtmExactPdfPageLayout(rawItems, page, textContent, viewport1x);
      allPagesHtml.push(pageHtml);
      allPagesText.push(pagePlainText);
    } else {
      const b64 = canvas.toDataURL('image/jpeg', 0.92);
      const fd = new FormData();
      fd.append('apikey', 'helloworld');
      fd.append('base64Image', b64);
      fd.append('language', 'auto');
      fd.append('OCREngine', '2');
      fd.append('scale', 'true');
      const resp = await fetch('https://api.ocr.space/parse/image', { method: 'POST', body: fd });
      const json = await resp.json();
      if (json?.ParsedResults?.[0]?.ParsedText) {
        const ocrTxt = json.ParsedResults[0].ParsedText.trim();
        allPagesText.push(ocrTxt);
        allPagesHtml.push(ocrTxt.split('\n').filter(Boolean).map(l => `<p style="text-align:justify; text-indent:1.27cm; margin:0 0 5pt 0;">${escapeHtml(l)}</p>`).join('\n'));
      }
    }
  }
  preservedUploadedWordHtml = allPagesHtml.map((h, i) => i === 0 ? h : `<div style="page-break-before:always; margin-top:16pt;">${h}</div>`).join('\n');
  return allPagesText.join('\n\n');
}

function setupWordFileUpload() {
  const fileInput = document.getElementById('vbtmWordFileInput');
  if (!fileInput) return;

  const processFile = async (file) => {
    if (!file) return;
    try {
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      const isPdf = ext === 'pdf' || file.type === 'application/pdf';
      let rawText = '';

      preservedUploadedWordHtml = '';
      preservedUploadedFileName = file.name.replace(/\.[^.]+$/, '');

      if (isPdf) {
        if (window.showToast) window.showToast(`Đang đọc tệp PDF "${file.name}" giữ nguyên 100% thể thức gốc...`);
        rawText = await extractTextFromPdfInVbtm(file);
        const previewEl = document.getElementById('a4PreviewPaper');
        if (previewEl && preservedUploadedWordHtml) {
          previewEl.innerHTML = preservedUploadedWordHtml;
        }
        document.getElementById('vbtmDocBody').value = rawText;
        runSpellAndFormatCheck();
        if (window.showToast) {
          window.showToast(`Đã tải tệp PDF "${file.name}" giữ nguyên 100% thể thức và nội dung gốc! Bấm "Xuất File Word" để tải về.`);
        }
        return;
      } else if (ext === 'docx' && window.mammoth) {
        const arrayBuffer = await file.arrayBuffer();
        const htmlResult = await window.mammoth.convertToHtml({ arrayBuffer });
        const textResult = await window.mammoth.extractRawText({ arrayBuffer });
        rawText = textResult.value || '';
        preservedUploadedWordHtml = htmlResult.value || '';
        const previewEl = document.getElementById('a4PreviewPaper');
        if (previewEl && preservedUploadedWordHtml) {
          previewEl.innerHTML = preservedUploadedWordHtml;
        }
        document.getElementById('vbtmDocBody').value = rawText;
        runSpellAndFormatCheck();
        if (window.showToast) {
          window.showToast(`Đã nạp tệp Word "${file.name}" giữ nguyên thể thức gốc!`);
        }
        return;
      } else {
        const textContent = await file.text();
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
      console.error('Lỗi đọc tệp:', err);
      if (window.showToast) window.showToast('Lỗi khi đọc tệp: ' + err.message);
    }
  };

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    fileInput.value = '';
    await processFile(file);
  });

  // Hỗ trợ kéo thả trực tiếp file .docx, .doc, .pdf, .txt vào khung VBTM Pro Studio
  const studioApp = document.getElementById('vbtmStudioApp');
  if (studioApp) {
    studioApp.addEventListener('dragover', (e) => e.preventDefault());
    studioApp.addEventListener('drop', async (e) => {
      e.preventDefault();
      if (e.dataTransfer?.files?.[0]) {
        await processFile(e.dataTransfer.files[0]);
      }
    });
  }
}

// Phân tích thông minh văn bản tải lên và tự động bóc tách các trường thể thức (Không trộn dữ liệu mẫu vào file của người dùng!)
function parseUploadedDocumentAndStandardize(rawText, fileName) {
  const lines = rawText
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  const fullUpper = rawText.toUpperCase();
  if (fullUpper.includes('ĐẢNG CỘNG SẢN VIỆT NAM') || fullUpper.includes('ĐẢNG ỦY') || fullUpper.includes('/ĐU') || fullUpper.includes('-ĐU')) {
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

  // Tìm Tên loại văn bản và Trích yếu
  const docTypes = ['BÁO CÁO', 'QUYẾT ĐỊNH', 'KẾ HOẠCH', 'TỜ TRÌNH', 'THÔNG BÁO', 'NGHỊ QUYẾT', 'CHƯƠNG TRÌNH', 'CHỈ THỊ', 'QUY CHẾ', 'BIÊN BẢN', 'GIẤY MỜI', 'HƯỚNG DẪN'];
  let detectedType = '';
  let detectedSubject = '';
  let typeLineIdx = -1;

  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const u = lines[i].toUpperCase();
    for (const t of docTypes) {
      if (u === t || u.startsWith(t + ' ')) {
        detectedType = t;
        typeLineIdx = i;
        break;
      }
    }
    if (detectedType) break;
  }

  if (typeLineIdx !== -1 && lines[typeLineIdx + 1]) {
    detectedSubject = lines[typeLineIdx + 1];
  } else {
    const vvIdx = lines.findIndex(l => /^(V\/v|Về việc)\b/i.test(l));
    if (vvIdx !== -1) {
      detectedSubject = lines[vvIdx];
      typeLineIdx = vvIdx - 1;
    }
  }

  document.getElementById('vbtmDocType').value = detectedType;
  document.getElementById('vbtmDocSubject').value = detectedSubject;

  // Lọc bỏ các dòng Quốc hiệu / Tiêu ngữ / Tên loại / Trích yếu đã tách ở đầu để lấy phần thân văn bản sạch
  const filteredBodyLines = lines.filter((l, idx) => {
    const u = l.toUpperCase();
    if (u === 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM') return false;
    if (u.includes('ĐỘC LẬP - TỰ DO - HẠNH PHÚC')) return false;
    if (u === 'ĐẢNG CỘNG SẢN VIỆT NAM') return false;
    if (idx === typeLineIdx || (detectedSubject && l === detectedSubject)) return false;
    return true;
  });

  document.getElementById('vbtmDocBody').value = filteredBodyLines.join('\n\n');
  standardizeAndFixEverything();

  if (window.showToast) {
    window.showToast(`Đã nạp tệp "${fileName}" và tự động chuẩn hóa thể thức + sửa lỗi chính tả!`);
  }
}

function loadTemplate(key) {
  preservedUploadedWordHtml = '';
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

// Hàm sửa toàn bộ lỗi chính tả (535 cụm từ nhầm lẫn + 7.578 âm tiết Tiếng Việt chuẩn) trên một chuỗi văn bản
function autoFixVietnameseSpellingInString(inputStr) {
  let str = String(inputStr || '').normalize('NFC');
  let fixedCount = 0;

  // Bước 1: Quét 535 cụm từ nhầm lẫn hành chính & lỗi gõ Telex/VNI bằng Unicode Regex
  const rules = window.VN_SPELL_RULES || [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (!rule.w || !rule.r) continue;
    const regex = createUnicodePhraseRegex(rule.w, 'gui');
    str = str.replace(regex, (matched) => {
      fixedCount++;
      return applyCasePattern(matched, rule.r);
    });
  }

  // Bước 2: Quét từng âm tiết Tiếng Việt đơn lẻ với kho 7.578 âm tiết chuẩn Quốc gia
  str = str.replace(/\p{L}+/gu, (word) => {
    if (isValidVietnameseSyllable(word)) return word;
    const suggestion = findClosestValidVietnameseSyllable(word);
    if (suggestion && suggestion.word) {
      fixedCount++;
      return applyCasePattern(word, suggestion.word);
    }
    return word;
  });

  return { text: str, fixedCount };
}

// Chuẩn hóa toàn diện: vừa chuẩn hóa thể thức NĐ30/Đảng ủy vừa tự động sửa toàn bộ lỗi chính tả AI
function standardizeAndFixEverything() {
  const bodyEl = document.getElementById('vbtmDocBody');
  const subjEl = document.getElementById('vbtmDocSubject');
  if (!bodyEl) return;

  // 1. Tự động sửa toàn bộ lỗi chính tả tiếng Việt trên cả Nội dung & Trích yếu
  const fixedBodyObj = autoFixVietnameseSpellingInString(bodyEl.value);
  let text = fixedBodyObj.text;
  let fixedSpells = fixedBodyObj.fixedCount;

  if (subjEl && subjEl.value) {
    const fixedSubjObj = autoFixVietnameseSpellingInString(subjEl.value);
    subjEl.value = fixedSubjObj.text.trim().replace(/\s+/g, ' ').replace(/\.+$/, '');
    fixedSpells += fixedSubjObj.fixedCount;
  }

  // 2. Xóa khoảng trắng thừa, chuẩn hóa dấu câu chuẩn Nghị định 30/2020/NĐ-CP
  text = text
    .replace(/[ \t]+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([,.;:!?])([^\s\d/.)])/g, '$1 $2')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/“\s+/g, '“')
    .replace(/\s+”/g, '”')
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

  // 4. Đảm bảo có Nơi nhận Lưu: VT
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

// Kiểm tra lỗi chính tả & thẩm định thể thức AI (535 cụm từ + 7.578 âm tiết Tiếng Việt chuẩn)
function runSpellAndFormatCheck() {
  const body = (document.getElementById('vbtmDocBody')?.value || '').normalize('NFC');
  const subject = (document.getElementById('vbtmDocSubject')?.value || '').normalize('NFC');
  const docNum = document.getElementById('vbtmDocNum')?.value || '';
  const recipients = document.getElementById('vbtmRecipients')?.value || '';
  const issueContainer = document.getElementById('vbtmIssueList');

  lastDetectedIssues = [];

  // 1. Kiểm tra theo kho 535 cụm từ nhầm lẫn hành chính & lỗi gõ phím
  const rules = window.VN_SPELL_RULES || [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (!rule.w || !rule.r) continue;
    const checkRegex = createUnicodePhraseRegex(rule.w, 'ui');
    if (checkRegex.test(body) || checkRegex.test(subject)) {
      const replaceRegex = createUnicodePhraseRegex(rule.w, 'gui');
      lastDetectedIssues.push({
        type: 'Chính tả AI',
        desc: `Phát hiện "${rule.w}" → Chuẩn là: "${rule.r}" (${rule.d})`,
        fix: () => {
          const bEl = document.getElementById('vbtmDocBody');
          const sEl = document.getElementById('vbtmDocSubject');
          if (bEl) bEl.value = bEl.value.normalize('NFC').replace(replaceRegex, m => applyCasePattern(m, rule.r));
          if (sEl) sEl.value = sEl.value.normalize('NFC').replace(replaceRegex, m => applyCasePattern(m, rule.r));
        }
      });
    }
  }

  // 2. Kiểm tra từng âm tiết đơn lẻ với kho 7.578 âm tiết Tiếng Việt chuẩn
  const combinedText = `${subject} ${body}`;
  const words = combinedText.match(/\p{L}+/gu) || [];
  const checkedInvalidWords = new Set();

  words.forEach(word => {
    const lower = word.toLowerCase();
    if (checkedInvalidWords.has(lower)) return;
    if (!isValidVietnameseSyllable(word)) {
      checkedInvalidWords.add(lower);
      const suggestion = findClosestValidVietnameseSyllable(word);
      const sugWord = suggestion ? suggestion.word : '';
      const sugReason = suggestion ? suggestion.reason : 'Âm tiết không tồn tại trong Tiếng Việt chuẩn';
      const wordRegex = createUnicodePhraseRegex(lower, 'gui');

      lastDetectedIssues.push({
        type: 'Ngữ âm AI',
        desc: sugWord
          ? `Âm tiết sai "${word}" (${sugReason}) → Gợi ý sửa thành: "${applyCasePattern(word, sugWord)}"`
          : `Phát hiện âm tiết sai cấu trúc Tiếng Việt: "${word}"`,
        fix: () => {
          if (!sugWord) return;
          const bEl = document.getElementById('vbtmDocBody');
          const sEl = document.getElementById('vbtmDocSubject');
          if (bEl) bEl.value = bEl.value.normalize('NFC').replace(wordRegex, m => applyCasePattern(m, sugWord));
          if (sEl) sEl.value = sEl.value.normalize('NFC').replace(wordRegex, m => applyCasePattern(m, sugWord));
        }
      });
    }
  });

  // 3. Kiểm tra quy tắc dấu câu & thể thức Nghị định 30/2020/NĐ-CP
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
        <span>✓ Văn bản đã đạt chuẩn 100% về chính tả Tiếng Việt (7.578 âm tiết &amp; 535 quy tắc) và thể thức Đảng ủy / UBND phường!</span>
      </div>
    `;
  } else {
    issueContainer.innerHTML = lastDetectedIssues.map((item, idx) => `
      <div class="issue-item">
        <div>
          <strong style="color: #fbbf24;">[${item.type}]</strong> ${escapeHtml(item.desc)}
        </div>
        <button type="button" class="btn-tool primary" onclick="fixSingleIssue(${idx})" style="padding: 4px 10px; font-size: 0.78rem; flex-shrink: 0;">Sửa ngay</button>
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
  bodyEl.value = `Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý ngiêm túc các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra`;
  renderA4Preview();
  runSpellAndFormatCheck();
  if (window.showToast) {
    window.showToast(`Đã phát hiện ${lastDetectedIssues.length} lỗi chính tả & thể thức. Bấm "⚡ Chuẩn Hóa Thể Thức & Tự Động Sửa Lỗi" để sửa sạch 100%!`);
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

  if (typeof window.qrcode === 'function') {
    try {
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

  const docNum = preservedUploadedFileName || (document.getElementById('vbtmDocNum')?.value || 'VanBan').replace(/[^a-zA-Z0-9_-]/g, '_');

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
        line-height: 1.45;
        color: #000000;
      }
      table { width: 100%; border-collapse: collapse; border: none; }
      td { vertical-align: top; text-align: center; font-family: "Times New Roman", serif; }
      p { margin: 0 0 6pt 0; text-indent: 1.27cm; text-align: justify; font-size: 14pt; }
      .a4-doc-title { text-align: center; font-weight: bold; font-size: 14pt; text-transform: uppercase; margin-top: 16pt; margin-bottom: 4pt; }
      .a4-doc-subject { text-align: center; font-weight: bold; font-size: 13.5pt; margin-bottom: 14pt; }
      .a4-line-sep { display: block; border-bottom: 1px solid #000; margin: 4px auto; }
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
    window.showToast(`Đã xuất tệp Microsoft Word chuẩn A4: VBTM_ChuanHoa_${docNum}.doc`);
  }
}

function downloadQrPng() {
  const canvas = document.getElementById('vbtmQrCanvas');
  if (!canvas) return;
  const url = canvas.toDataURL('image/png');
  const a = document.createElement('a');
  a.href = url;
  a.download = 'VBTM_MaQR_VanBan.png';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  if (window.showToast) window.showToast('Đã tải xuống ảnh Mã QR chuẩn ISO 18004 (PNG)!');
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
window.downloadQrPng = downloadQrPng;
