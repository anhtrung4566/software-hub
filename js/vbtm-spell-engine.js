/**
 * VBTM Pro 3.0 Ultimate — Vietnamese Spell & Proofreading Engine + Entity Protection
 * - Phân loại đầy đủ 10 nhóm lỗi chính tả / văn bản
 * - Bảo vệ tuyệt đối: Tên người, Địa danh, Cơ quan, Số hiệu, Ngày tháng, Số liệu, Tiền, %, Điều/Khoản/Điểm, Căn cứ pháp lý, Chức vụ
 * - Thay thế chính xác theo (paragraphId, startOffset) — KHÔNG thay thế chuỗi toàn cục
 * - Hỗ trợ song song: Chế độ Quy tắc Cục bộ (tức thì) & Chế độ AI Gemini Structured JSON
 */
(function (global) {
  'use strict';

  const ERROR_CATEGORIES = [
    'Lỗi chính tả',
    'Lỗi đánh máy',
    'Lỗi dấu câu',
    'Lỗi viết hoa',
    'Lỗi ngữ pháp',
    'Lỗi diễn đạt',
    'Lỗi lặp từ',
    'Câu thiếu thành phần',
    'Lỗi thuật ngữ hành chính',
    'Lỗi thể thức văn bản'
  ];

  // Quy tắc rà soát ngoại tuyến (bổ sung cho từ điển 7.578 âm tiết)
  const SPELL_RULES = [
    { wrong: 'bổ xung', suggest: 'bổ sung', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'sử lý', suggest: 'xử lý', category: 'Lỗi chính tả', reason: 'Sai phụ âm s/x', confidence: 0.99 },
    { wrong: 'xử dụng', suggest: 'sử dụng', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'chính sác', suggest: 'chính xác', category: 'Lỗi chính tả', reason: 'Sai phụ âm s/x', confidence: 0.99 },
    { wrong: 'kiểm xoát', suggest: 'kiểm soát', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'rà xoát', suggest: 'rà soát', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'sâu xát', suggest: 'sâu sát', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.98 },
    { wrong: 'xắp xếp', suggest: 'sắp xếp', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'suất sắc', suggest: 'xuất sắc', category: 'Lỗi chính tả', reason: 'Sai phụ âm s/x', confidence: 0.99 },
    { wrong: 'sản suất', suggest: 'sản xuất', category: 'Lỗi chính tả', reason: 'Sai phụ âm s/x', confidence: 0.99 },
    { wrong: 'đề suất', suggest: 'đề xuất', category: 'Lỗi chính tả', reason: 'Sai phụ âm s/x', confidence: 0.99 },
    { wrong: 'sơ xuất', suggest: 'sơ suất', category: 'Lỗi chính tả', reason: 'Sai phụ âm x/s', confidence: 0.99 },
    { wrong: 'chiển khai', suggest: 'triển khai', category: 'Lỗi chính tả', reason: 'Sai phụ âm ch/tr', confidence: 0.99 },
    { wrong: 'chủ chương', suggest: 'chủ trương', category: 'Lỗi chính tả', reason: 'Sai phụ âm ch/tr', confidence: 0.99 },
    { wrong: 'khẩn chương', suggest: 'khẩn trương', category: 'Lỗi chính tả', reason: 'Sai phụ âm ch/tr', confidence: 0.99 },
    { wrong: 'tập chung', suggest: 'tập trung', category: 'Lỗi chính tả', reason: 'Sai phụ âm ch/tr', confidence: 0.98 },
    { wrong: 'ngiêm túc', suggest: 'nghiêm túc', category: 'Lỗi chính tả', reason: 'Sai quy tắc ng/ngh', confidence: 0.99 },
    { wrong: 'hổ trợ', suggest: 'hỗ trợ', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'chuẫn hóa', suggest: 'chuẩn hóa', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'chuẫn bị', suggest: 'chuẩn bị', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'tiêu chuẫn', suggest: 'tiêu chuẩn', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'kỷ thuật', suggest: 'kỹ thuật', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'lỉnh vực', suggest: 'lĩnh vực', category: 'Lỗi chính tả', reason: 'Sai dấu hỏi/ngã', confidence: 0.99 },
    { wrong: 'chặc chẽ', suggest: 'chặt chẽ', category: 'Lỗi chính tả', reason: 'Sai âm cuối c/t', confidence: 0.99 },
    { wrong: 'chắc chẽ', suggest: 'chặt chẽ', category: 'Lỗi chính tả', reason: 'Từ chuẩn: chặt chẽ', confidence: 0.98 },
    { wrong: 'sát nhập', suggest: 'sáp nhập', category: 'Lỗi thuật ngữ hành chính', reason: 'Thuật ngữ chuẩn hành chính: sáp nhập', confidence: 0.98 },
    { wrong: 'hành chánh', suggest: 'hành chính', category: 'Lỗi thuật ngữ hành chính', reason: 'Thuật ngữ chuẩn NĐ 30: hành chính', confidence: 0.98 },
    { wrong: 'kính gởi', suggest: 'kính gửi', category: 'Lỗi thuật ngữ hành chính', reason: 'Thuật ngữ chuẩn NĐ 30: kính gửi', confidence: 0.99 },
    { wrong: 'chuẩn đoán', suggest: 'chẩn đoán', category: 'Lỗi chính tả', reason: 'Từ chuẩn tiếng Việt: chẩn đoán', confidence: 0.98 },
    { wrong: 'thăm quan', suggest: 'tham quan', category: 'Lỗi chính tả', reason: 'Từ chuẩn tiếng Việt: tham quan', confidence: 0.98 },
    { wrong: 'vô hình chung', suggest: 'vô hình trung', category: 'Lỗi chính tả', reason: 'Từ chuẩn tiếng Việt: vô hình trung', confidence: 0.98 },
    { wrong: 'chỉnh chu', suggest: 'chỉn chu', category: 'Lỗi chính tả', reason: 'Từ chuẩn tiếng Việt: chỉn chu', confidence: 0.98 },
    { wrong: 'nồng cốt', suggest: 'nòng cốt', category: 'Lỗi chính tả', reason: 'Từ chuẩn tiếng Việt: nòng cốt', confidence: 0.98 },
    { wrong: 'kế hoạc', suggest: 'kế hoạch', category: 'Lỗi đánh máy', reason: 'Thiếu chữ h cuối', confidence: 0.99 },
    { wrong: 'đưọc', suggest: 'được', category: 'Lỗi đánh máy', reason: 'Sai dấu ươ', confidence: 0.99 },
    { wrong: 'ngưòi', suggest: 'người', category: 'Lỗi đánh máy', reason: 'Sai dấu ươ', confidence: 0.99 },
    { wrong: 'phưòng', suggest: 'phường', category: 'Lỗi đánh máy', reason: 'Sai dấu ươ', confidence: 0.99 },
    { wrong: 'trưóc', suggest: 'trước', category: 'Lỗi đánh máy', reason: 'Sai dấu ươ', confidence: 0.99 }
  ];

  const PROPER_NOUN_RULES = [
    { wrong: 'đảng cộng sản việt nam', suggest: 'Đảng Cộng sản Việt Nam', reason: 'Viết hoa chuẩn 05-HD/VPTW' },
    { wrong: 'ban chấp hành trung ương', suggest: 'Ban Chấp hành Trung ương', reason: 'Viết hoa chuẩn 05-HD/VPTW' },
    { wrong: 'ban thường vụ', suggest: 'Ban Thường vụ', reason: 'Viết hoa chuẩn 05-HD/VPTW' },
    { wrong: 'ban chấp hành', suggest: 'Ban Chấp hành', reason: 'Viết hoa chuẩn 05-HD/VPTW' },
    { wrong: 'ủy ban kiểm tra', suggest: 'Ủy ban Kiểm tra', reason: 'Viết hoa chuẩn 05-HD/VPTW' },
    { wrong: 'chính phủ', suggest: 'Chính phủ', reason: 'Viết hoa chuẩn NĐ 30/2020' },
    { wrong: 'quốc hội', suggest: 'Quốc hội', reason: 'Viết hoa chuẩn NĐ 30/2020' },
    { wrong: 'ủy ban nhân dân', suggest: 'Ủy ban nhân dân', reason: 'Viết hoa chuẩn NĐ 30/2020' },
    { wrong: 'hội đồng nhân dân', suggest: 'Hội đồng nhân dân', reason: 'Viết hoa chuẩn NĐ 30/2020' }
  ];

  /**
   * Kiểm tra xem chuỗi gốc có chạm tới thực thể được bảo vệ hay không
   * (số hiệu, số liệu, tiền, %, ngày tháng, Điều/Khoản/Điểm, căn cứ pháp lý, chức vụ, tên riêng)
   */
  function detectProtectedEntityReason(originalText, surroundingText) {
    const t = String(originalText || '').trim();
    const ctx = String(surroundingText || '');
    if (!t) return null;

    if (/\d+\s*(%|đồng|vnđ|vnd|triệu|tỷ|tỉ|nghìn|ngàn|km|m2|ha|kg|người|hộ|bản)/i.test(t)) {
      return 'Chứa số liệu / tiền tệ / tỷ lệ phần trăm';
    }
    if (/\b\d+[A-Za-zÀ-ỹĐđ\-]*\/[A-ZÀ-ỸĐ][A-ZÀ-ỸĐa-z0-9\-]+/.test(t)) {
      return 'Chứa số hiệu văn bản';
    }
    if (/\b(Điều|Khoản|Điểm)\s+\d+/i.test(t)) {
      return 'Chứa dẫn chiếu Điều / Khoản / Điểm';
    }
    if (/\b(Luật|Nghị định|Thông tư|Hướng dẫn|Quyết định|Nghị quyết|Chỉ thị)\s+(số\s+)?\d+/i.test(t) || /\bCăn cứ\b/i.test(t)) {
      return 'Chứa căn cứ pháp lý';
    }
    if (/\b(ngày|tháng|năm)\s+\d+/i.test(t) && !/^(ngày\s+[1-9]\s+tháng|tháng\s+[12]\s+năm)$/i.test(t)) {
      return 'Chứa ngày tháng cụ thể';
    }
    if (/\b(Chủ tịch|Phó Chủ tịch|Bí thư|Phó Bí thư|Giám đốc|Trưởng phòng|Bộ trưởng|Thủ tướng)\b/.test(t)) {
      return 'Chứa chức vụ lãnh đạo';
    }
    // Tên người hoặc địa danh viết hoa 2–4 từ liên tiếp (VD: Nguyễn Văn An, Tân Sơn Nhất)
    if (/^([A-ZÀ-ỸĐ][a-zà-ỹđ]+\s+){1,3}[A-ZÀ-ỸĐ][a-zà-ỹđ]+$/.test(t) && !/^(Ủy ban|Hội đồng|Ban Thường|Ban Chấp|Chính phủ|Quốc hội)/i.test(t)) {
      return 'Nghi ngờ tên riêng / địa danh';
    }
    if (/\d/.test(t) && !/^(ngày\s+[1-9]\s+tháng|tháng\s+[12]\s+năm)$/i.test(t)) {
      return 'Chứa chữ số / số liệu cần xác nhận';
    }
    if (ctx && /\b(kết luận|chỉ đạo|quyết nghị|quyết định)\b/i.test(ctx) && t.length > 40) {
      return 'Thuộc đoạn chỉ đạo / kết luận quan trọng';
    }
    return null;
  }

  function matchCase(orig, replacement) {
    if (!orig || !replacement) return replacement;
    if (orig === orig.toUpperCase()) return replacement.toUpperCase();
    if (orig[0] === orig[0].toUpperCase()) {
      return replacement.charAt(0).toUpperCase() + replacement.slice(1);
    }
    return replacement;
  }

  /**
   * Quét ngoại tuyến nhanh theo từng đoạn văn (trả về danh sách có paragraphId & startOffset chính xác)
   */
  function scanOffline(fullText) {
    const normalized = String(fullText || '').normalize('NFC');
    const paragraphs = normalized.split('\n');
    const issues = [];
    let seq = 1;

    paragraphs.forEach((para, pIdx) => {
      if (!para || !para.trim()) return;

      // 1. Quy tắc từ điển & thuật ngữ hành chính
      for (const rule of SPELL_RULES) {
        const rx = new RegExp(`(?<![A-Za-zÀ-Ỹà-ỹĐđ])${rule.wrong.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Za-zÀ-Ỹà-ỹĐđ])`, 'gi');
        let m;
        while ((m = rx.exec(para)) !== null) {
          const orig = m[0];
          const sug = matchCase(orig, rule.suggest);
          if (orig === sug) continue;
          const prot = detectProtectedEntityReason(orig, para);
          issues.push({
            id: `off_${seq++}`,
            paragraphId: pIdx,
            startOffset: m.index,
            length: orig.length,
            original: orig,
            suggestion: sug,
            category: rule.category,
            reason: prot ? `${rule.reason} (${prot})` : rule.reason,
            confidence: prot ? Math.min(rule.confidence, 0.72) : rule.confidence,
            needsConfirmation: Boolean(prot),
            source: 'offline',
            status: 'pending'
          });
        }
      }

      // 2. Từ gõ lặp ký tự (VD: HOẠCHH -> HOẠCH) & kẹt phím Telex
      const wordRx = /[A-Za-zÀ-Ỹà-ỹĐđ]{2,}/g;
      const whiteSet = new Set(['II', 'III', 'VII', 'VIII', 'XII', 'XIII', 'BOONG', 'SOONG', 'XOONG', 'GOOGLE', 'OFFLINE', 'OFFICE', 'EMAIL', 'ONLINE', 'PDF', 'DOCX', 'HTTP', 'HTTPS', 'WWW']);
      let wm;
      while ((wm = wordRx.exec(para)) !== null) {
        const tok = wm[0];
        if (whiteSet.has(tok.toUpperCase())) continue;

        if (/([A-Za-zÀ-Ỹà-ỹĐđ])\1+/i.test(tok)) {
          const dedup = tok.replace(/([A-Za-zÀ-Ỹà-ỹĐđ])\1+/gi, '$1');
          if (dedup !== tok) {
            const prot = detectProtectedEntityReason(tok, para);
            issues.push({
              id: `off_${seq++}`,
              paragraphId: pIdx,
              startOffset: wm.index,
              length: tok.length,
              original: tok,
              suggestion: dedup,
              category: 'Lỗi đánh máy',
              reason: 'Lặp ký tự bất thường khi gõ phím',
              confidence: prot ? 0.7 : 0.96,
              needsConfirmation: Boolean(prot),
              source: 'offline',
              status: 'pending'
            });
            continue;
          }
        }

        if (/[À-Ỹà-ỹĐđ]/.test(tok) && tok.length >= 3 && /[fsrxjwz]$/i.test(tok)) {
          const trimmed = tok.slice(0, -1);
          issues.push({
            id: `off_${seq++}`,
            paragraphId: pIdx,
            startOffset: wm.index,
            length: tok.length,
            original: tok,
            suggestion: trimmed,
            category: 'Lỗi đánh máy',
            reason: `Kẹt phím dấu Telex ('${tok.slice(-1)}') cuối từ tiếng Việt`,
            confidence: 0.92,
            needsConfirmation: false,
            source: 'offline',
            status: 'pending'
          });
        }
      }

      // 3. Lặp từ liền nhau (VD: "các các", "đã đã")
      const dupRx = /\b([A-Za-zÀ-Ỹà-ỹĐđ]{2,})\s+\1\b/gi;
      const allowDup = new Set(['người', 'ngày', 'năm', 'nhà', 'lần', 'bước', 'tầng']);
      let dm;
      while ((dm = dupRx.exec(para)) !== null) {
        if (allowDup.has(dm[1].toLowerCase())) continue;
        issues.push({
          id: `off_${seq++}`,
          paragraphId: pIdx,
          startOffset: dm.index,
          length: dm[0].length,
          original: dm[0],
          suggestion: dm[1],
          category: 'Lỗi lặp từ',
          reason: 'Từ bị lặp 2 lần liên tiếp',
          confidence: 0.95,
          needsConfirmation: false,
          source: 'offline',
          status: 'pending'
        });
      }

      // 4. Lỗi viết hoa tên cơ quan chuẩn NĐ 30 / 05-HD
      for (const pn of PROPER_NOUN_RULES) {
        const rx = new RegExp(`(?<![A-Za-zÀ-Ỹà-ỹĐđ])${pn.wrong.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![A-Za-zÀ-Ỹà-ỹĐđ])`, 'gi');
        let pm;
        while ((pm = rx.exec(para)) !== null) {
          const val = pm[0];
          if (val === pn.suggest || val === pn.suggest.toUpperCase()) continue;
          issues.push({
            id: `off_${seq++}`,
            paragraphId: pIdx,
            startOffset: pm.index,
            length: val.length,
            original: val,
            suggestion: pn.suggest,
            category: 'Lỗi viết hoa',
            reason: pn.reason,
            confidence: 0.94,
            needsConfirmation: false,
            source: 'offline',
            status: 'pending'
          });
        }
      }

      // 5. Lỗi dấu câu & khoảng trắng
      const punctRx = /\s+([,.;:!?])|([,;])([A-Za-zÀ-Ỹà-ỹĐđ])|\s{2,}/g;
      let pm;
      while ((pm = punctRx.exec(para)) !== null) {
        let sug = ' ';
        let reason = 'Nhiều khoảng trắng thừa liền nhau';
        if (pm[1]) { sug = pm[1]; reason = `Khoảng trắng thừa trước dấu "${pm[1]}"`; }
        else if (pm[2]) { sug = `${pm[2]} ${pm[3]}`; reason = `Thiếu khoảng trắng sau dấu "${pm[2]}"`; }
        issues.push({
          id: `off_${seq++}`,
          paragraphId: pIdx,
          startOffset: pm.index,
          length: pm[0].length,
          original: pm[0],
          suggestion: sug,
          category: 'Lỗi dấu câu',
          reason,
          confidence: 0.98,
          needsConfirmation: false,
          source: 'offline',
          status: 'pending'
        });
      }

      // 6. Thể thức ngày 1–9 & tháng 1–2 (NĐ 30 & 05-HD)
      const dateRx = /\bngày\s+([1-9])\s+tháng\b|\btháng\s+([12])\s+năm\b/gi;
      let dt;
      while ((dt = dateRx.exec(para)) !== null) {
        const rep = dt[0].replace(/\b([1-9])\b/, '0$1');
        issues.push({
          id: `off_${seq++}`,
          paragraphId: pIdx,
          startOffset: dt.index,
          length: dt[0].length,
          original: dt[0],
          suggestion: rep,
          category: 'Lỗi thể thức văn bản',
          reason: 'Ngày dưới 10 và tháng 1, 2 phải thêm số 0 phía trước (NĐ 30/2020)',
          confidence: 0.97,
          needsConfirmation: false,
          source: 'offline',
          status: 'pending'
        });
      }
    });

    return issues;
  }

  /**
   * Quét sâu bằng Google Gemini AI (Structured JSON Schema)
   * Tuyệt đối KHÔNG viết lại toàn bộ văn bản, chỉ trả về danh sách lỗi có tọa độ đoạn
   */
  async function scanWithGeminiAI(fullText) {
    if (!global.VbtmAiService) {
      throw { code: 'NO_SERVICE', userMessage: 'Thiếu module VbtmAiService.' };
    }
    const normalized = String(fullText || '').normalize('NFC');
    const paragraphs = normalized.split('\n');
    const indexedDoc = paragraphs
      .map((p, idx) => (p.trim() ? `[P${idx}] ${p}` : ''))
      .filter(Boolean)
      .join('\n')
      .slice(0, 14000);

    const sysInstruction =
      'Bạn là chuyên gia hiệu đính văn bản Đảng (Hướng dẫn 05-HD/VPTW) và Hành chính Nhà nước Việt Nam (Nghị định 30/2020/NĐ-CP).\n' +
      'NHIỆM VỤ: Chỉ tìm các lỗi thực sự tồn tại trong văn bản được đánh số đoạn [P0], [P1], ... KHÔNG ĐƯỢC viết lại toàn bộ văn bản.\n' +
      'NGUYÊN TẮC BẢO VỆ NỘI DUNG: Tuyệt đối KHÔNG tự ý thay đổi tên người, địa danh, tên cơ quan, số hiệu văn bản, ngày tháng, số liệu, số tiền, tỷ lệ %, Điều/Khoản/Điểm, căn cứ pháp lý, chức vụ, nội dung chỉ đạo hay kết luận. Nếu nghi ngờ, đặt needsConfirmation = true.\n' +
      'Trả về JSON array (tối đa 30 phần tử), mỗi phần tử có cấu trúc:\n' +
      '{"paragraphId": number, "original": "trích đúng nguyên văn cụm sai trong đoạn", "suggestion": "cụm sửa đề xuất", "category": "1 trong 10 loại: Lỗi chính tả | Lỗi đánh máy | Lỗi dấu câu | Lỗi viết hoa | Lỗi ngữ pháp | Lỗi diễn đạt | Lỗi lặp từ | Câu thiếu thành phần | Lỗi thuật ngữ hành chính | Lỗi thể thức văn bản", "reason": "giải thích ngắn", "confidence": 0.0-1.0, "needsConfirmation": boolean}\n' +
      'Nếu không có lỗi, trả về [].';

    const res = await global.VbtmAiService.request({
      actionName: 'spellCheckAI',
      systemInstruction: sysInstruction,
      prompt: indexedDoc,
      expectJson: true,
      temperature: 0.15,
      maxOutputTokens: 2048
    });

    const rawList = Array.isArray(res.json) ? res.json : (res.json && Array.isArray(res.json.issues) ? res.json.issues : []);
    const verified = [];
    let seq = 1;

    for (const item of rawList) {
      if (!item || typeof item.original !== 'string' || typeof item.suggestion !== 'string') continue;
      const orig = item.original.trim();
      const sug = item.suggestion.trim();
      if (!orig || !sug || orig === sug) continue;

      let pIdx = Number(item.paragraphId);
      if (!Number.isInteger(pIdx) || pIdx < 0 || pIdx >= paragraphs.length) {
        pIdx = paragraphs.findIndex(p => p.includes(orig));
      }
      if (pIdx < 0) continue;

      const offset = paragraphs[pIdx].indexOf(orig);
      if (offset < 0) continue; // Không tồn tại nguyên văn -> loại bỏ để tránh AI bịa

      const protReason = detectProtectedEntityReason(orig, paragraphs[pIdx]);
      const cat = ERROR_CATEGORIES.includes(item.category) ? item.category : 'Lỗi chính tả';
      const conf = Math.max(0.1, Math.min(1, Number(item.confidence) || 0.85));

      verified.push({
        id: `ai_${seq++}`,
        paragraphId: pIdx,
        startOffset: offset,
        length: orig.length,
        original: orig,
        suggestion: sug,
        category: cat,
        reason: protReason ? `${item.reason || 'Cần kiểm tra'} (Bảo vệ: ${protReason})` : (item.reason || 'AI phát hiện lỗi ngữ cảnh'),
        confidence: protReason ? Math.min(conf, 0.7) : conf,
        needsConfirmation: Boolean(item.needsConfirmation || protReason || cat === 'Lỗi diễn đạt' || cat === 'Câu thiếu thành phần'),
        source: 'gemini',
        status: 'pending'
      });
    }

    return { issues: verified, model: res.model, latencyMs: res.latencyMs, truncated: res.truncated };
  }

  /**
   * Áp dụng sửa MỘT lỗi theo đúng tọa độ (paragraphId, startOffset) — KHÔNG replace toàn cục!
   */
  function applySingleFix(fullText, issue, customSuggestion) {
    const normalized = String(fullText || '').normalize('NFC');
    const paragraphs = normalized.split('\n');
    const pIdx = issue.paragraphId;
    if (pIdx < 0 || pIdx >= paragraphs.length) return { ok: false, text: fullText, reason: 'Không tìm thấy đoạn văn.' };

    const para = paragraphs[pIdx];
    const replacement = typeof customSuggestion === 'string' ? customSuggestion : issue.suggestion;

    // Kiểm tra chính xác tại startOffset; nếu trước đó đoạn đã đổi độ dài, tìm vị trí gần startOffset nhất trong chính đoạn đó
    let pos = -1;
    if (para.slice(issue.startOffset, issue.startOffset + issue.original.length) === issue.original) {
      pos = issue.startOffset;
    } else {
      let bestDist = Infinity;
      let searchFrom = 0;
      while (true) {
        const found = para.indexOf(issue.original, searchFrom);
        if (found < 0) break;
        const dist = Math.abs(found - issue.startOffset);
        if (dist < bestDist) { bestDist = dist; pos = found; }
        searchFrom = found + 1;
      }
    }
    if (pos < 0) return { ok: false, text: fullText, reason: 'Nội dung gốc đã thay đổi hoặc đã được sửa.' };

    paragraphs[pIdx] = para.slice(0, pos) + replacement + para.slice(pos + issue.original.length);
    return { ok: true, text: paragraphs.join('\n') };
  }

  /**
   * Sửa tất cả lỗi AN TOÀN (needsConfirmation === false && confidence >= 0.85)
   * Sắp xếp ngược từ cuối đoạn về đầu đoạn để không lệch offset!
   */
  function applyAllSafeFixes(fullText, issueList) {
    const normalized = String(fullText || '').normalize('NFC');
    const paragraphs = normalized.split('\n');
    const safe = issueList
      .filter(it => it.status === 'pending' && !it.needsConfirmation && (it.confidence || 0) >= 0.85)
      .slice()
      .sort((a, b) => (b.paragraphId - a.paragraphId) || (b.startOffset - a.startOffset));

    let fixedCount = 0;
    const fixedIds = new Set();

    for (const it of safe) {
      const para = paragraphs[it.paragraphId];
      if (!para) continue;
      if (para.slice(it.startOffset, it.startOffset + it.original.length) === it.original) {
        paragraphs[it.paragraphId] = para.slice(0, it.startOffset) + it.suggestion + para.slice(it.startOffset + it.original.length);
        fixedCount++;
        fixedIds.add(it.id);
      }
    }

    return {
      text: paragraphs.join('\n'),
      fixedCount,
      fixedIds,
      skippedProtectedCount: issueList.filter(it => it.status === 'pending' && it.needsConfirmation).length
    };
  }

  global.VbtmSpellEngine = {
    ERROR_CATEGORIES,
    detectProtectedEntityReason,
    scanOffline,
    scanWithGeminiAI,
    applySingleFix,
    applyAllSafeFixes
  };
})(typeof window !== 'undefined' ? window : globalThis);
