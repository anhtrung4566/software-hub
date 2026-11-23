/**
 * VBTM.SpellUnified — Hệ thống kiểm tra & soát lỗi văn bản tiếng Việt 15 nhóm, 4 mức độ
 * - Phân biệt đủ 15 nhóm lỗi (KHÔNG gộp tất cả thành lỗi chính tả)
 * - 4 mức độ: Nghiêm trọng | Cần sửa | Cần xem xét | Gợi ý cải thiện
 * - Bảo vệ từ chuyên ngành, tên riêng, số liệu, căn cứ pháp lý & từ điển tùy chỉnh
 * - Xác minh nội dung gốc còn khớp trước khi áp dụng sửa đổi
 * - Kiểm tra chặt chẽ schema JSON khi gọi AI (không suy diễn lỗi nếu JSON hỏng)
 */
(function (global) {
  'use strict';

  const ERROR_CATEGORIES = {
    CHINH_TA: 'Lỗi chính tả',
    DAU_THANH: 'Lỗi dấu thanh',
    GO_NHAM: 'Lỗi gõ nhầm',
    THIEU_TU: 'Thiếu từ',
    THUA_TU: 'Thừa từ',
    LAP_TU: 'Lặp từ',
    DAU_CAU: 'Lỗi dấu câu',
    NGU_PHAP: 'Lỗi ngữ pháp',
    DIEN_DAT: 'Lỗi diễn đạt',
    DUNG_TU: 'Lỗi dùng từ',
    TRINH_BAY: 'Lỗi trình bày',
    THE_THUC: 'Lỗi thể thức',
    SO_LIEU: 'Lỗi số liệu',
    LOGIC: 'Lỗi logic',
    CAN_CU_PHAP_LY: 'Lỗi căn cứ pháp lý'
  };

  const SEVERITIES = ['Nghiêm trọng', 'Cần sửa', 'Cần xem xét', 'Gợi ý cải thiện'];

  const DETERMINISTIC_RULES = [
    { re: /\b36-HD\/VPTW\b/g, suggestion: '05-HD/VPTW', category: 'CAN_CU_PHAP_LY', severity: 'Nghiêm trọng', reason: 'Hướng dẫn 36-HD/VPTW đã hết hiệu lực, cần thay bằng Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng.', impact: 'Sai căn cứ pháp lý thể thức văn bản Đảng', reference: 'Hướng dẫn 05-HD/VPTW', confidence: 0.99 },
    { re: /\bđề\s+suất\b/gi, suggestion: 'đề xuất', category: 'CHINH_TA', severity: 'Cần sửa', reason: 'Nhầm phụ âm đầu s/x.', impact: 'Ảnh hưởng tính chuẩn mực văn bản', reference: 'Từ điển Tiếng Việt', confidence: 0.98 },
    { re: /\bbổ\s+xung\b/gi, suggestion: 'bổ sung', category: 'CHINH_TA', severity: 'Cần sửa', reason: 'Nhầm phụ âm đầu x/s.', impact: 'Ảnh hưởng tính chuẩn mực văn bản', reference: 'Từ điển Tiếng Việt', confidence: 0.98 },
    { re: /\bxử\s+dụng\b/gi, suggestion: 'sử dụng', category: 'CHINH_TA', severity: 'Cần sửa', reason: 'Nhầm phụ âm đầu x/s.', impact: 'Ảnh hưởng tính chuẩn mực văn bản', reference: 'Từ điển Tiếng Việt', confidence: 0.98 },
    { re: /\bngân\s+xách\b/gi, suggestion: 'ngân sách', category: 'CHINH_TA', severity: 'Cần sửa', reason: 'Nhầm phụ âm đầu x/s.', impact: 'Sai thuật ngữ tài chính công', reference: 'Từ điển Tiếng Việt', confidence: 0.98 },
    { re: /\bkế\s+hoạc\b/gi, suggestion: 'kế hoạch', category: 'GO_NHAM', severity: 'Cần sửa', reason: 'Thiếu chữ cái cuối "h" do gõ nhầm.', impact: 'Lỗi chính tả/gõ phím', reference: '', confidence: 0.98 },
    { re: /\bchính\s+xác\b/gi, suggestion: 'chính xác', category: 'CHINH_TA', severity: 'Cần sửa', skipExact: true },
    { re: /\bsát\s+nhập\b/gi, suggestion: 'sáp nhập', category: 'DUNG_TU', severity: 'Cần sửa', reason: 'Thuật ngữ hành chính chuẩn là "sáp nhập".', impact: 'Sai thuật ngữ hành chính', reference: 'Văn bản quy phạm pháp luật', confidence: 0.95 },
    { re: /\b(của|và|để|được|các|những|trong|theo)\s+\1\b/gi, suggestion: '$1', category: 'LAP_TU', severity: 'Cần sửa', reason: 'Lặp từ liền kề.', impact: 'Gây rườm rà câu văn', reference: '', confidence: 0.96 },
    { re: /\s+([,.;:!?])/g, suggestion: '$1', category: 'DAU_CAU', severity: 'Cần sửa', reason: 'Thừa khoảng trắng trước dấu câu.', impact: 'Lỗi trình bày dấu câu chuẩn NĐ 30', reference: 'Nghị định 30/2020/NĐ-CP', confidence: 0.97 }
  ];

  function isProtectedText(fragment, customProtected) {
    const f = String(fragment || '').trim();
    if (!f) return true;
    const list = (customProtected || []).concat(global.VBTM_SyncEngine ? global.VBTM_SyncEngine.getProtectedTerms() : []);
    for (const term of list) {
      if (term && f.toLowerCase() === String(term).toLowerCase()) return true;
    }
    // Bảo vệ từ viết tắt hợp lệ (UBND, HĐND, MTTQ, VPTW, NĐ-CP...)
    if (/^[A-ZĐÀ-Ỹ]{2,8}(-[A-ZĐÀ-Ỹ]{2,6})?$/.test(f)) return true;
    return false;
  }

  function runOfflineChecks(paragraphs, customProtected) {
    const issues = [];
    (paragraphs || []).forEach((p) => {
      const text = String(p.text || '');
      for (const rule of DETERMINISTIC_RULES) {
        if (rule.skipExact) continue;
        rule.re.lastIndex = 0;
        let m;
        while ((m = rule.re.exec(text)) !== null) {
          const orig = m[0];
          const sugg = orig.replace(new RegExp(rule.re.source, rule.re.flags.replace('g', '')), rule.suggestion);
          if (orig === sugg || isProtectedText(orig, customProtected)) continue;
          issues.push({
            paragraphId: p.id,
            position: `Đoạn #${p.id + 1} (vị trí ${m.index})`,
            original: orig,
            suggestion: sugg,
            category: rule.category,
            categoryLabel: ERROR_CATEGORIES[rule.category] || rule.category,
            severity: rule.severity,
            reason: rule.reason,
            confidence: rule.confidence,
            impact: rule.impact,
            reference: rule.reference || '',
            source: 'offline-rule'
          });
        }
      }
    });
    return issues;
  }

  function validateAiIssuesSchema(rawArray, paragraphs, customProtected) {
    if (!Array.isArray(rawArray)) return { valid: false, issues: [], error: 'Kết quả AI không phải JSON Array.' };
    const pMap = new Map((paragraphs || []).map(p => [Number(p.id), String(p.text || '')]));
    const clean = [];

    for (const item of rawArray) {
      if (!item || typeof item !== 'object') continue;
      const orig = String(item.original || '').trim();
      const sugg = String(item.suggestion || '').trim();
      if (!orig || !sugg || orig === sugg) continue;
      if (isProtectedText(orig, customProtected)) continue;

      const pid = Number(item.paragraphId || 0);
      const pText = pMap.get(pid);
      // Xác minh cụm từ gốc thực sự tồn tại trong đoạn (chống AI bịa lỗi)
      if (pText !== undefined && pText.indexOf(orig) === -1) continue;

      const catKey = ERROR_CATEGORIES[item.category] ? item.category : 'CHINH_TA';
      const sev = SEVERITIES.includes(item.severity) ? item.severity : 'Cần xem xét';
      const conf = Number(item.confidence);

      clean.push({
        paragraphId: pid,
        position: String(item.position || `Đoạn #${pid + 1}`),
        original: orig,
        suggestion: sugg,
        category: catKey,
        categoryLabel: ERROR_CATEGORIES[catKey],
        severity: sev,
        reason: String(item.reason || 'Cần rà soát theo chuẩn tiếng Việt công vụ.'),
        confidence: isNaN(conf) ? 0.8 : Math.max(0, Math.min(1, conf)),
        impact: String(item.impact || 'Ảnh hưởng độ chính xác văn bản'),
        reference: String(item.reference || ''),
        source: 'ai-verified'
      });
    }
    return { valid: true, issues: clean, error: null };
  }

  function verifyOriginalStillMatches(currentParagraphText, originalFragment) {
    return String(currentParagraphText || '').indexOf(String(originalFragment || '')) !== -1;
  }

  async function inspectParagraphs(paragraphs, opts) {
    opts = opts || {};
    const customProtected = opts.protectedTerms || [];
    const offlineIssues = runOfflineChecks(paragraphs, customProtected);

    if (opts.offlineOnly) {
      return {
        issues: offlineIssues,
        aiUsed: false,
        message: offlineIssues.length ? `Phát hiện ${offlineIssues.length} vấn đề (kiểm tra quy tắc cục bộ).` : 'Không phát hiện lỗi trong phạm vi kiểm tra quy tắc cục bộ.'
      };
    }

    const sysPrompt =
      'Bạn là chuyên gia thẩm định & soát lỗi văn bản hành chính/Đảng Việt Nam. ' +
      'Phân loại chính xác vào 1 trong 15 mã category: CHINH_TA, DAU_THANH, GO_NHAM, THIEU_TU, THUA_TU, LAP_TU, DAU_CAU, NGU_PHAP, DIEN_DAT, DUNG_TU, TRINH_BAY, THE_THUC, SO_LIEU, LOGIC, CAN_CU_PHAP_LY. ' +
      'Mức severity: "Nghiêm trọng" | "Cần sửa" | "Cần xem xét" | "Gợi ý cải thiện". ' +
      'TUYỆT ĐỐI KHÔNG sửa tên riêng, tên cơ quan, số hiệu văn bản hợp lệ, thuật ngữ Đảng/pháp lý. KHÔNG đánh dấu sai nếu chỉ là lựa chọn văn phong. ' +
      'Trả về JSON array: [{"paragraphId":0,"position":"Đoạn #1","original":"cụm gốc chính xác","suggestion":"đề xuất","category":"CHINH_TA","severity":"Cần sửa","reason":"lý do","confidence":0.95,"impact":"ảnh hưởng","reference":"nguồn nếu có"}]. Nếu không có lỗi trả về [].';

    const inputText = (paragraphs || []).map(p => `[Đoạn #${p.id}] ${p.text}`).join('\n');
    const router = global.VBTM_SmartRouter ? global.VBTM_SmartRouter.route : global.VBTM_AICore.generate;
    const res = await router(sysPrompt, inputText, {
      host: opts.host || 'Word',
      taskType: 'spell',
      expectJson: true,
      historyTitle: 'Soát lỗi 15 nhóm'
    });

    const parsed = global.VBTM_AICore.safeParseJson(res.text);
    if (!parsed.ok) {
      return {
        issues: offlineIssues,
        aiUsed: true,
        warning: 'Kết quả JSON từ AI không hợp lệ — chỉ hiển thị lỗi đã xác minh cục bộ, không tự suy diễn.',
        message: `Đã giữ ${offlineIssues.length} lỗi cục bộ xác minh.`
      };
    }

    const validated = validateAiIssuesSchema(parsed.data, paragraphs, customProtected);
    const merged = offlineIssues.slice();
    validated.issues.forEach(aiIt => {
      const dup = merged.some(x => x.paragraphId === aiIt.paragraphId && x.original === aiIt.original);
      if (!dup) merged.push(aiIt);
    });

    return {
      issues: merged,
      aiUsed: true,
      provider: res.provider || 'gemini',
      model: res.model,
      message: merged.length
        ? `Phát hiện ${merged.length} lỗi/gợi ý đã xác minh khớp nội dung gốc.`
        : 'Không phát hiện lỗi trong phạm vi đã kiểm tra.'
    };
  }

  global.VBTM_SpellUnified = {
    ERROR_CATEGORIES,
    SEVERITIES,
    runOfflineChecks,
    validateAiIssuesSchema,
    verifyOriginalStillMatches,
    inspectParagraphs
  };
})(typeof window !== 'undefined' ? window : globalThis);
