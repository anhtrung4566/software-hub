/**
 * VBTM Pro v3.5.0 — Công cụ Tìm Kiếm & Thay Từ Tiếng Việt Có Dấu Chuẩn Xác
 * Hỗ trợ:
 *  - Tìm tiếng Việt có dấu (NFC normalized)
 *  - Phân biệt chữ HOA / thường (matchCase)
 *  - Tìm từ nguyên vẹn tiếng Việt (wholeWord — không thay nhầm một phần của từ khác)
 *  - Đếm số kết quả, nhảy tới kết quả trước / sau
 *  - Thay từng vị trí, Xem trước trước khi Thay tất cả, Thay tất cả, Hoàn tác (Undo)
 */
(function (global) {
  'use strict';

  const VI_WORD_CHAR = '0-9A-Za-zÀ-Ỹà-ỹĐđ_';
  const replaceUndoStack = [];
  let lastMatches = [];
  let currentMatchIdx = -1;

  function escapeRegExp(s) {
    return String(s || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function buildSearchRegex(query, opts) {
    opts = opts || {};
    const q = String(query || '').normalize('NFC');
    if (!q) return null;
    const escaped = escapeRegExp(q);
    const flags = opts.matchCase ? 'g' : 'gi';
    if (opts.wholeWord) {
      // Dùng ranh giới ký tự tiếng Việt chuẩn để không thay nhầm 1 phần từ khác
      return new RegExp(`(^|[^${VI_WORD_CHAR}])(${escaped})(?![${VI_WORD_CHAR}])`, flags);
    }
    return new RegExp(`()(${escaped})`, flags);
  }

  function findMatches(text, query, opts) {
    const src = String(text || '').normalize('NFC');
    const re = buildSearchRegex(query, opts);
    if (!re) return [];
    const out = [];
    let m;
    while ((m = re.exec(src)) !== null) {
      const start = m.index + (m[1] ? m[1].length : 0);
      const val = m[2];
      out.push({ start, end: start + val.length, text: val });
      if (re.lastIndex === m.index) re.lastIndex++;
    }
    return out;
  }

  function replaceAllPure(text, query, replacement, opts) {
    const src = String(text || '').normalize('NFC');
    const matches = findMatches(src, query, opts);
    if (!matches.length) return { text: src, count: 0 };
    let out = '';
    let cursor = 0;
    for (const m of matches) {
      out += src.slice(cursor, m.start) + replacement;
      cursor = m.end;
    }
    out += src.slice(cursor);
    return { text: out, count: matches.length };
  }

  function replaceSinglePure(text, match, replacement) {
    const src = String(text || '').normalize('NFC');
    return src.slice(0, match.start) + replacement + src.slice(match.end);
  }

  function getEditor() { return document.getElementById('vbtmDocBody'); }

  function getUiOpts() {
    return {
      query: document.getElementById('vbtmFindInput')?.value || '',
      replacement: document.getElementById('vbtmReplaceInput')?.value || '',
      matchCase: !!document.getElementById('vbtmFindMatchCase')?.checked,
      wholeWord: !!document.getElementById('vbtmFindWholeWord')?.checked
    };
  }

  function updateCounterUi(msg) {
    const el = document.getElementById('vbtmFindStatus');
    if (el) el.textContent = msg;
  }

  function runFind(stepDir) {
    const ed = getEditor();
    if (!ed) return;
    const o = getUiOpts();
    lastMatches = findMatches(ed.value, o.query, o);
    if (!lastMatches.length) {
      currentMatchIdx = -1;
      updateCounterUi('Không tìm thấy kết quả phù hợp.');
      return;
    }
    if (stepDir === 1) currentMatchIdx = (currentMatchIdx + 1) % lastMatches.length;
    else if (stepDir === -1) currentMatchIdx = (currentMatchIdx - 1 + lastMatches.length) % lastMatches.length;
    else currentMatchIdx = 0;

    const hit = lastMatches[currentMatchIdx];
    ed.focus();
    ed.setSelectionRange(hit.start, hit.end);
    updateCounterUi(`Tìm thấy ${lastMatches.length} vị trí (đang ở ${currentMatchIdx + 1}/${lastMatches.length}).`);
  }

  function replaceCurrent() {
    const ed = getEditor();
    if (!ed) return;
    const o = getUiOpts();
    lastMatches = findMatches(ed.value, o.query, o);
    if (!lastMatches.length) { updateCounterUi('Không có vị trí để thay thế.'); return; }
    if (currentMatchIdx < 0 || currentMatchIdx >= lastMatches.length) currentMatchIdx = 0;

    replaceUndoStack.push(ed.value);
    ed.value = replaceSinglePure(ed.value, lastMatches[currentMatchIdx], o.replacement);
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    runFind(0);
    updateCounterUi(`Đã thay 1 vị trí. Còn ${lastMatches.length} vị trí.`);
  }

  function previewAndReplaceAll() {
    const ed = getEditor();
    if (!ed) return;
    const o = getUiOpts();
    const res = replaceAllPure(ed.value, o.query, o.replacement, o);
    if (!res.count) { updateCounterUi('Không có kết quả nào để thay thế.'); return; }
    const ok = confirm(`XEM TRƯỚC THAY HÀNG LOẠT:\n- Tìm: "${o.query}" (${o.wholeWord ? 'Từ nguyên vẹn' : 'Mọi chuỗi'}${o.matchCase ? ', phân biệt HOA/thường' : ''})\n- Thay bằng: "${o.replacement}"\n- Số vị trí sẽ thay: ${res.count}\n\nBạn có xác nhận thay tất cả? (Có thể bấm Hoàn tác.)`);
    if (!ok) return;
    replaceUndoStack.push(ed.value);
    ed.value = res.text;
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    updateCounterUi(`✓ Đã thay tất cả ${res.count} vị trí (có thể Hoàn tác).`);
  }

  function undoReplace() {
    const ed = getEditor();
    if (!ed || !replaceUndoStack.length) {
      updateCounterUi('Chưa có thao tác thay từ nào để hoàn tác.');
      return;
    }
    ed.value = replaceUndoStack.pop();
    if (typeof global.renderA4Preview === 'function') global.renderA4Preview();
    updateCounterUi('↩ Đã hoàn tác thao tác thay từ vừa rồi.');
  }

  function togglePanel() {
    const p = document.getElementById('vbtmFindReplacePanel');
    if (!p) return;
    p.style.display = (p.style.display === 'none' || !p.style.display) ? 'block' : 'none';
    if (p.style.display === 'block') document.getElementById('vbtmFindInput')?.focus();
  }

  global.VbtmFindReplace = {
    findMatches,
    replaceAllPure,
    replaceSinglePure,
    runFind,
    replaceCurrent,
    previewAndReplaceAll,
    undoReplace,
    togglePanel
  };
})(typeof window !== 'undefined' ? window : globalThis);
