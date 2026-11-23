/**
 * VBTM.WordAddin — Direct Microsoft Word Integration (Office.js)
 * Tương thích 3 mức:
 * - BASIC (Office 2016 Common API): đọc/chèn/thay thế đoạn chọn an toàn
 * - STANDARD / ADVANCED (WordApi 1.1+): đọc từng đoạn theo ID, điều hướng đến vị trí lỗi,
 *   thay thế chính xác từng cụm từ mà vẫn bảo toàn định dạng, bảng biểu, ảnh, Header/Footer.
 */
(function () {
  'use strict';

  let runtimeInfo = null;
  let lastOriginalSelection = '';
  let undoStack = []; // { mode: 'selection'|'paragraph', paragraphIndex, previousText, replacedText }
  let scannedParagraphs = []; // { id, text }

  function showAlert(type, msg) {
    const el = document.getElementById('wordAlert');
    if (!el) return;
    if (!msg) { el.className = 'hidden'; el.textContent = ''; return; }
    el.className = type === 'err' ? 'alert-err' : (type === 'ok' ? 'alert-ok' : 'alert-warn');
    el.textContent = msg;
  }

  async function readSelectedTextFromWord() {
    if (!runtimeInfo || !runtimeInfo.inOffice) {
      showAlert('warn', 'Đang mở ngoài Word (chế độ xem trước). Bạn có thể gõ nội dung mẫu trực tiếp vào ô bên trên để thử AI.');
      return document.getElementById('txtInput').value.trim();
    }
    if (window.Word && runtimeInfo.capabilities.paragraphScan) {
      return await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        return (sel.text || '').trim();
      });
    }
    // Fallback Common API (Office 2016 Basic)
    return await new Promise((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || '').trim());
        else reject(new Error(res.error ? res.error.message : 'Không đọc được đoạn đang chọn.'));
      });
    });
  }

  async function applyTextToWord(newText, mode) {
    if (!runtimeInfo || !runtimeInfo.inOffice) {
      showAlert('warn', 'Chế độ xem trước Web: đã sao chép nội dung đề xuất (khi chạy trong Word sẽ chèn/thay trực tiếp).');
      try { await navigator.clipboard.writeText(newText); } catch (e) {}
      return;
    }

    if (window.Word && runtimeInfo.capabilities.paragraphScan) {
      await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.load('text');
        await context.sync();
        const prev = sel.text || '';
        const loc = mode === 'replace' ? Word.InsertLocation.replace : Word.InsertLocation.after;
        sel.insertText(mode === 'replace' ? newText : ('\n' + newText), loc);
        await context.sync();
        undoStack.push({ mode: 'selection', previousText: prev, replacedText: newText });
      });
    } else {
      // Common API fallback
      const prev = await readSelectedTextFromWord();
      await new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(newText, { coercionType: Office.CoercionType.Text }, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error(res.error ? res.error.message : 'Không ghi được vào Word.'));
        });
      });
      undoStack.push({ mode: 'selection', previousText: prev, replacedText: newText });
    }
    document.getElementById('btnUndoWord').classList.remove('hidden');
  }

  async function undoLastWordEdit() {
    if (!undoStack.length) return;
    const last = undoStack.pop();
    if (!runtimeInfo || !runtimeInfo.inOffice) return;

    if (last.mode === 'paragraph' && window.Word && runtimeInfo.capabilities.paragraphScan) {
      await Word.run(async (context) => {
        const paras = context.document.body.paragraphs;
        paras.load('items');
        await context.sync();
        const p = paras.items[last.paragraphIndex];
        if (!p) throw new Error('Không tìm thấy đoạn văn để hoàn tác.');
        const found = p.search(last.replacedText, { matchCase: false });
        found.load('items');
        await context.sync();
        if (found.items.length > 0) {
          found.items[0].insertText(last.previousText, Word.InsertLocation.replace);
          await context.sync();
        }
      });
    } else if (window.Word && runtimeInfo.capabilities.paragraphScan) {
      await Word.run(async (context) => {
        const sel = context.document.getSelection();
        sel.insertText(last.previousText, Word.InsertLocation.replace);
        await context.sync();
      });
    } else {
      await new Promise((resolve, reject) => {
        Office.context.document.setSelectedDataAsync(last.previousText, { coercionType: Office.CoercionType.Text }, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
          else reject(new Error('Không thể hoàn tác qua Common API.'));
        });
      });
    }
    if (!undoStack.length) document.getElementById('btnUndoWord').classList.add('hidden');
    showAlert('ok', 'Đã hoàn tác thao tác sửa vừa rồi trong Word.');
  }

  async function scanFullDocumentParagraphs() {
    const progressEl = document.getElementById('scanProgress');
    const listEl = document.getElementById('scanIssuesList');
    listEl.innerHTML = '';

    if (!runtimeInfo || !runtimeInfo.inOffice || !runtimeInfo.capabilities.paragraphScan) {
      progressEl.textContent = 'Chức năng quét toàn văn bản yêu cầu chạy trong Microsoft Word (WordApi 1.1+). Với Office 2016 Basic, vui lòng bôi đen từng đoạn ở tab Trợ lý AI.';
      return;
    }

    progressEl.textContent = 'Đang đọc danh sách đoạn văn bản (giữ nguyên bảng biểu, ảnh, Header/Footer)...';
    scannedParagraphs = await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load('items/text');
      await context.sync();
      const out = [];
      paras.items.forEach((p, idx) => {
        const t = (p.text || '').trim();
        if (t.length >= 4) out.push({ id: idx, text: t });
      });
      return out;
    });

    if (!scannedParagraphs.length) {
      progressEl.textContent = 'Tài liệu hiện chưa có nội dung văn bản để quét.';
      return;
    }

    // Chia thành từng nhóm 8 đoạn để tránh quá tải token và giữ nguyên ID đoạn
    const chunkSize = 8;
    const allIssues = [];
    const totalChunks = Math.ceil(scannedParagraphs.length / chunkSize);

    for (let c = 0; c < totalChunks; c++) {
      const slice = scannedParagraphs.slice(c * chunkSize, (c + 1) * chunkSize);
      progressEl.textContent = `Đang rà soát nhóm đoạn ${c + 1}/${totalChunks} (${scannedParagraphs.length} đoạn · 15 nhóm lỗi)...`;
      try {
        const res = await window.VBTM_SpellUnified.inspectParagraphs(slice, { host: 'Word' });
        if (res && Array.isArray(res.issues)) {
          res.issues.forEach(item => allIssues.push(item));
        }
      } catch (err) {
        progressEl.textContent = `Lỗi tại nhóm ${c + 1}/${totalChunks}: ${err.message}`;
        break;
      }
    }

    progressEl.textContent = `Hoàn tất rà soát ${scannedParagraphs.length} đoạn: phát hiện ${allIssues.length} đề xuất chỉnh sửa.`;
    renderIssues(allIssues);
  }

  function renderIssues(issues) {
    const listEl = document.getElementById('scanIssuesList');
    const esc = window.VBTM_SettingsUI.escapeHtml;
    if (!issues.length) {
      listEl.innerHTML = '<div class="alert-ok">Không phát hiện lỗi chính tả hoặc ngữ pháp đáng kể!</div>';
      return;
    }
    listEl.innerHTML = issues.map((it, i) => `
      <div class="issue-item" data-idx="${i}">
        <div class="issue-head">[${esc(it.severity || 'Cần sửa')} · ${esc(it.categoryLabel || 'Chính tả')}] Đoạn #${it.paragraphId + 1}: "${esc(it.original)}"</div>
        <div class="issue-sugg">→ Đề xuất: "${esc(it.suggestion)}"</div>
        <div class="issue-reason">${esc(it.reason)}${it.reference ? ' · Nguồn: ' + esc(it.reference) : ''}</div>
        <div class="btn-row">
          <button class="btn btn-secondary btn-nav-issue" data-idx="${i}" type="button">Đi tới vị trí</button>
          <button class="btn btn-success btn-fix-issue" data-idx="${i}" type="button">Xác nhận sửa lỗi này</button>
        </div>
      </div>
    `).join('');

    listEl.querySelectorAll('.btn-nav-issue').forEach(btn => {
      btn.addEventListener('click', async () => {
        const it = issues[Number(btn.getAttribute('data-idx'))];
        await navigateToParagraphIssue(it);
      });
    });

    listEl.querySelectorAll('.btn-fix-issue').forEach(btn => {
      btn.addEventListener('click', async () => {
        const it = issues[Number(btn.getAttribute('data-idx'))];
        try {
          await replaceIssueInParagraph(it);
          btn.disabled = true;
          btn.textContent = 'Đã sửa';
        } catch (err) {
          showAlert('err', 'Không sửa được vị trí này: ' + err.message);
        }
      });
    });
  }

  async function navigateToParagraphIssue(issue) {
    if (!window.Word) return;
    await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const p = paras.items[issue.paragraphId];
      if (!p) return;
      const found = p.search(issue.original, { matchCase: false });
      found.load('items');
      await context.sync();
      if (found.items.length > 0) found.items[0].select();
      else p.select();
      await context.sync();
    });
  }

  async function replaceIssueInParagraph(issue) {
    if (!window.Word) return;
    await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load('items');
      await context.sync();
      const p = paras.items[issue.paragraphId];
      if (!p) throw new Error('Đoạn văn không còn tồn tại.');
      p.load('text');
      await context.sync();
      if (window.VBTM_SpellUnified && !window.VBTM_SpellUnified.verifyOriginalStillMatches(p.text, issue.original)) {
        throw new Error('Nội dung gốc tại đoạn này đã thay đổi — hủy thay thế để tránh sửa nhầm.');
      }
      const found = p.search(issue.original, { matchCase: false });
      found.load('items');
      await context.sync();
      if (!found.items.length) throw new Error('Không tìm thấy cụm từ gốc trong đoạn.');
      // Chỉ thay thế đúng cụm từ lỗi để bảo toàn định dạng xung quanh
      found.items[0].insertText(issue.suggestion, Word.InsertLocation.replace);
      await context.sync();
      undoStack.push({
        mode: 'paragraph',
        paragraphIndex: issue.paragraphId,
        previousText: issue.original,
        replacedText: issue.suggestion
      });
    });
    document.getElementById('btnUndoWord').classList.remove('hidden');
  }

  function initWordUi() {
    runtimeInfo = window.VBTM_SettingsUI.bindCommonUi('Word');

    const sel = document.getElementById('wordActionSelect');
    sel.innerHTML = window.VBTM_Prompts.WORD_ACTIONS.map(a => `<option value="${a.id}">${a.label}</option>`).join('');

    document.getElementById('btnReadSelection').addEventListener('click', async () => {
      try {
        const txt = await readSelectedTextFromWord();
        if (!txt) {
          showAlert('warn', 'Chưa có đoạn văn bản nào được bôi đen trong Word.');
          return;
        }
        lastOriginalSelection = txt;
        document.getElementById('txtInput').value = txt;
        showAlert('ok', `Đã đọc đoạn đang chọn (${txt.length} ký tự).`);
      } catch (err) {
        showAlert('err', err.message);
      }
    });

    const btnRun = document.getElementById('btnRunAi');
    const btnCancel = document.getElementById('btnCancelAi');
    btnRun.addEventListener('click', async () => {
      const inputEl = document.getElementById('txtInput');
      let text = inputEl.value.trim();
      if (!text && runtimeInfo && runtimeInfo.inOffice) {
        try {
          text = await readSelectedTextFromWord();
          if (text) inputEl.value = text;
        } catch (e) {}
      }
      if (!text) {
        showAlert('warn', 'Vui lòng bôi đen đoạn văn bản trong Word hoặc nhập yêu cầu trước khi chạy AI.');
        return;
      }
      lastOriginalSelection = text;
      const act = window.VBTM_Prompts.WORD_ACTIONS.find(a => a.id === sel.value) || window.VBTM_Prompts.WORD_ACTIONS[0];
      showAlert('', '');
      btnRun.disabled = true;
      btnCancel.classList.remove('hidden');

      try {
        const bridge = window.VBTM_OfficeBridge ? window.VBTM_OfficeBridge.createBridge('Word') : null;
        const res = bridge
          ? await bridge.runTask(act.id, act.prompt, text, { expectJson: false, historyTitle: act.label })
          : await window.VBTM_AICore.generate(act.prompt, text, { host: 'Word', expectJson: false, historyTitle: act.label });
        document.getElementById('origPreview').textContent = lastOriginalSelection;
        document.getElementById('txtProposal').value = res.text;
        document.getElementById('resultMeta').textContent = `${res.provider || 'gemini'} · ${res.model} · ${res.latencyMs}ms`;
        document.getElementById('resultCard').classList.remove('hidden');
      } catch (err) {
        showAlert('err', err.message || 'Lỗi gọi AI.');
      } finally {
        btnRun.disabled = false;
        btnCancel.classList.add('hidden');
      }
    });

    btnCancel.addEventListener('click', () => window.VBTM_AICore.cancelActive());

    document.getElementById('btnReplaceSel').addEventListener('click', async () => {
      const proposed = document.getElementById('txtProposal').value;
      if (!proposed.trim()) return;
      try {
        await applyTextToWord(proposed, 'replace');
        showAlert('ok', 'Đã thay thế đoạn đang chọn trong Word (có thể bấm Hoàn tác nếu cần).');
      } catch (err) {
        showAlert('err', err.message);
      }
    });

    document.getElementById('btnInsertCursor').addEventListener('click', async () => {
      const proposed = document.getElementById('txtProposal').value;
      if (!proposed.trim()) return;
      try {
        await applyTextToWord(proposed, 'insert');
        showAlert('ok', 'Đã chèn nội dung mới tại vị trí con trỏ trong Word.');
      } catch (err) {
        showAlert('err', err.message);
      }
    });

    document.getElementById('btnReject').addEventListener('click', () => {
      document.getElementById('resultCard').classList.add('hidden');
      showAlert('warn', 'Đã từ chối đề xuất — tài liệu Word không bị thay đổi.');
    });

    document.getElementById('btnUndoWord').addEventListener('click', async () => {
      try { await undoLastWordEdit(); } catch (err) { showAlert('err', err.message); }
    });

    document.getElementById('btnScanDoc').addEventListener('click', scanFullDocumentParagraphs);
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady(() => initWordUi());
  } else {
    window.addEventListener('DOMContentLoaded', initWordUi);
  }
})();
