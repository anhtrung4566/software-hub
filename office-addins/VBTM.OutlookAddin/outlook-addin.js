/**
 * VBTM.OutlookAddin — Direct Microsoft Outlook Integration (Classic Outlook & New Outlook)
 * Nguyên tắc an toàn: KHÔNG tự động gửi thư khi chưa có xác nhận của người dùng.
 */
(function () {
  'use strict';

  let runtimeInfo = null;

  function showAlert(type, msg) {
    const el = document.getElementById('mailAlert');
    if (!el) return;
    if (!msg) { el.className = 'hidden'; el.textContent = ''; return; }
    el.className = type === 'err' ? 'alert-err' : (type === 'ok' ? 'alert-ok' : 'alert-warn');
    el.textContent = msg;
  }

  async function readCurrentEmail() {
    if (!runtimeInfo || !runtimeInfo.inOffice || !Office.context.mailbox || !Office.context.mailbox.item) {
      showAlert('warn', 'Đang mở ngoài Outlook — bạn có thể dán nội dung email mẫu vào ô trên để thử AI.');
      return;
    }
    const item = Office.context.mailbox.item;
    let subject = '';
    if (typeof item.subject === 'string') {
      subject = item.subject;
    } else if (item.subject && item.subject.getAsync) {
      subject = await new Promise(resolve => item.subject.getAsync(r => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')));
    }
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve(String(res.value || '').trim());
        else reject(new Error(res.error ? res.error.message : 'Không đọc được nội dung email.'));
      });
    });
    document.getElementById('txtMailInput').value = (subject ? `Tiêu đề: ${subject}\n\n` : '') + body;
    showAlert('ok', `Đã đọc email (${body.length} ký tự · ${runtimeInfo.outlookClient || 'Outlook'}).`);
  }

  async function insertIntoComposeBody(text) {
    if (!runtimeInfo || !runtimeInfo.inOffice || !Office.context.mailbox || !Office.context.mailbox.item) {
      showAlert('warn', 'Chế độ xem trước ngoài Outlook — hãy dùng nút Sao chép.');
      return;
    }
    const item = Office.context.mailbox.item;
    if (!item.body || !item.body.setSelectedDataAsync) {
      showAlert('warn', 'Email hiện đang ở chế độ Đọc (Read). Hãy mở cửa sổ Trả lời / Soạn thư mới để chèn trực tiếp, hoặc bấm Sao chép.');
      return;
    }
    await new Promise((resolve, reject) => {
      item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(res.error ? res.error.message : 'Không chèn được vào nội dung thư.'));
      });
    });
    showAlert('ok', 'Đã chèn nội dung đề xuất vào cửa sổ soạn thư (thư CHƯA được gửi — hãy kiểm tra kỹ trước khi bấm Gửi).');
  }

  function initOutlookUi() {
    runtimeInfo = window.VBTM_SettingsUI.bindCommonUi('Outlook');
    const sel = document.getElementById('outlookActionSelect');
    sel.innerHTML = window.VBTM_Prompts.OUTLOOK_ACTIONS.map(a => `<option value="${a.id}">${a.label}</option>`).join('');

    document.getElementById('btnReadEmail').addEventListener('click', () => readCurrentEmail().catch(e => showAlert('err', e.message)));

    const btnRun = document.getElementById('btnRunMailAi');
    const btnCancel = document.getElementById('btnCancelMailAi');
    btnRun.addEventListener('click', async () => {
      const text = document.getElementById('txtMailInput').value.trim();
      if (!text) { showAlert('warn', 'Vui lòng bấm "Đọc Email hiện tại" hoặc nhập yêu cầu soạn thư.'); return; }
      const act = window.VBTM_Prompts.OUTLOOK_ACTIONS.find(a => a.id === sel.value) || window.VBTM_Prompts.OUTLOOK_ACTIONS[0];
      showAlert('', '');
      btnRun.disabled = true;
      btnCancel.classList.remove('hidden');
      try {
        const bridge = window.VBTM_OfficeBridge ? window.VBTM_OfficeBridge.createBridge('Outlook') : null;
        const res = bridge
          ? await bridge.runTask(act.id, act.prompt, text, { historyTitle: act.label })
          : await window.VBTM_AICore.generate(act.prompt, text, { host: 'Outlook', historyTitle: act.label });
        document.getElementById('txtMailOutput').value = res.text;
        document.getElementById('mailResultMeta').textContent = `${res.provider || 'gemini'} · ${res.model} · ${res.latencyMs}ms`;
        document.getElementById('mailResultCard').classList.remove('hidden');
      } catch (err) {
        showAlert('err', err.message || 'Lỗi gọi AI.');
      } finally {
        btnRun.disabled = false;
        btnCancel.classList.add('hidden');
      }
    });

    btnCancel.addEventListener('click', () => window.VBTM_AICore.cancelActive());
    document.getElementById('btnInsertMailBody').addEventListener('click', () => insertIntoComposeBody(document.getElementById('txtMailOutput').value).catch(e => showAlert('err', e.message)));
    document.getElementById('btnCopyMailOut').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText(document.getElementById('txtMailOutput').value); showAlert('ok', 'Đã sao chép nội dung email.'); } catch (e) {}
    });
  }

  if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(() => initOutlookUi());
  else window.addEventListener('DOMContentLoaded', initOutlookUi);
})();
