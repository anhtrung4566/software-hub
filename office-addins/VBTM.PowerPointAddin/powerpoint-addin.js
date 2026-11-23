/**
 * VBTM.PowerPointAddin — Direct Microsoft PowerPoint Integration (Office.js)
 * Nguyên tắc an toàn: KHÔNG tự động xóa slide hiện tại.
 */
(function () {
  'use strict';

  let runtimeInfo = null;
  let lastSlidesJson = null;

  function showAlert(type, msg) {
    const el = document.getElementById('pptAlert');
    if (!el) return;
    if (!msg) { el.className = 'hidden'; el.textContent = ''; return; }
    el.className = type === 'err' ? 'alert-err' : (type === 'ok' ? 'alert-ok' : 'alert-warn');
    el.textContent = msg;
  }

  function formatSlidesOutline(slides) {
    return slides.map((s, idx) => {
      const bullets = Array.isArray(s.bullets) ? s.bullets.map(b => '  • ' + b).join('\n') : '';
      const notes = s.notes ? `\n  [Lời thuyết trình]: ${s.notes}` : '';
      return `SLIDE ${idx + 1}: ${s.title || ''}\n${bullets}${notes}`;
    }).join('\n\n');
  }

  async function readSelectedPptText() {
    if (!runtimeInfo || !runtimeInfo.inOffice) {
      showAlert('warn', 'Đang mở ngoài PowerPoint — bạn có thể dán trực tiếp nội dung báo cáo vào ô bên trên.');
      return;
    }
    const txt = await new Promise((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || '').trim());
        else reject(new Error(res.error ? res.error.message : 'Hãy bấm vào khung văn bản trên Slide trước.'));
      });
    });
    document.getElementById('txtPptInput').value = txt;
    showAlert('ok', `Đã đọc khung chọn (${txt.length} ký tự).`);
  }

  async function insertIntoSelectedBox(text) {
    if (!runtimeInfo || !runtimeInfo.inOffice) {
      showAlert('warn', 'Chế độ xem trước ngoài PowerPoint — hãy dùng nút Sao chép.');
      return;
    }
    await new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error ? res.error.message : 'Hãy chọn một khung văn bản trên Slide rồi thử lại.'));
      });
    });
    showAlert('ok', 'Đã chèn nội dung vào khung văn bản đang chọn trên Slide.');
  }

  async function addSlidesToPresentation() {
    if (!runtimeInfo || !runtimeInfo.inOffice || !window.PowerPoint || !runtimeInfo.capabilities.pptSlideCreate) {
      showAlert('warn', 'Phiên bản PowerPoint này dùng MỨC 1 – BASIC (chưa hỗ trợ tạo Slide tự động bằng API). Vui lòng dùng nút "Chèn vào khung đang chọn" hoặc "Sao chép".');
      return;
    }
    const count = Array.isArray(lastSlidesJson) && lastSlidesJson.length ? lastSlidesJson.length : 1;
    if (!confirm(`Xác nhận thêm ${count} slide mới vào cuối bài trình chiếu (không xóa bất kỳ slide hiện có nào)?`)) return;
    await PowerPoint.run(async (context) => {
      for (let i = 0; i < count; i++) {
        context.presentation.slides.add();
      }
      await context.sync();
    });
    showAlert('ok', `Đã thêm ${count} slide mới an toàn! Bạn có thể dán từng mục nội dung từ khung kết quả vào các slide mới.`);
  }

  function initPptUi() {
    runtimeInfo = window.VBTM_SettingsUI.bindCommonUi('PowerPoint');
    const sel = document.getElementById('pptActionSelect');
    sel.innerHTML = window.VBTM_Prompts.PPT_ACTIONS.map(a => `<option value="${a.id}">${a.label}</option>`).join('');

    document.getElementById('btnReadPptSel').addEventListener('click', () => readSelectedPptText().catch(e => showAlert('err', e.message)));

    const btnRun = document.getElementById('btnRunPptAi');
    const btnCancel = document.getElementById('btnCancelPptAi');
    btnRun.addEventListener('click', async () => {
      const text = document.getElementById('txtPptInput').value.trim();
      if (!text) { showAlert('warn', 'Vui lòng nhập hoặc dán nội dung báo cáo / số liệu.'); return; }
      const act = window.VBTM_Prompts.PPT_ACTIONS.find(a => a.id === sel.value) || window.VBTM_Prompts.PPT_ACTIONS[0];
      const wantsJson = ['from_word', 'from_excel', 'conference', 'year_end', 'summarize_long'].includes(act.id);

      showAlert('', '');
      lastSlidesJson = null;
      btnRun.disabled = true;
      btnCancel.classList.remove('hidden');
      try {
        const res = await window.VBTM_AICore.generate(act.prompt, text, {
          host: 'PowerPoint',
          expectJson: wantsJson,
          historyTitle: act.label
        });
        let displayText = res.text;
        if (wantsJson) {
          const parsed = window.VBTM_AICore.safeParseJson(res.text);
          if (parsed.ok && Array.isArray(parsed.data)) {
            lastSlidesJson = parsed.data;
            displayText = formatSlidesOutline(parsed.data);
          }
        }
        document.getElementById('txtPptOutput').value = displayText;
        document.getElementById('pptResultMeta').textContent = `${res.model} · ${res.latencyMs}ms`;
        document.getElementById('pptResultCard').classList.remove('hidden');
      } catch (err) {
        showAlert('err', err.message || 'Lỗi gọi AI.');
      } finally {
        btnRun.disabled = false;
        btnCancel.classList.add('hidden');
      }
    });

    btnCancel.addEventListener('click', () => window.VBTM_AICore.cancelActive());
    document.getElementById('btnAddSlidesPpt').addEventListener('click', () => addSlidesToPresentation().catch(e => showAlert('err', e.message)));
    document.getElementById('btnInsertPptBox').addEventListener('click', () => insertIntoSelectedBox(document.getElementById('txtPptOutput').value).catch(e => showAlert('err', e.message)));
    document.getElementById('btnCopyPptOut').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText(document.getElementById('txtPptOutput').value); showAlert('ok', 'Đã sao chép nội dung Slide.'); } catch (e) {}
    });
  }

  if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(() => initPptUi());
  else window.addEventListener('DOMContentLoaded', initPptUi);
})();
