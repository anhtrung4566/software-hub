/**
 * VBTM.ExcelAddin — Direct Microsoft Excel Integration (Office.js)
 * Nguyên tắc an toàn: KHÔNG tự ý sửa dữ liệu gốc. Mọi thao tác ghi vào ô/sheet đều cần xác nhận.
 */
(function () {
  'use strict';

  let runtimeInfo = null;

  function showAlert(type, msg) {
    const el = document.getElementById('excelAlert');
    if (!el) return;
    if (!msg) { el.className = 'hidden'; el.textContent = ''; return; }
    el.className = type === 'err' ? 'alert-err' : (type === 'ok' ? 'alert-ok' : 'alert-warn');
    el.textContent = msg;
  }

  function valuesToTsv(values) {
    return (values || []).map(row => row.map(c => String(c == null ? '' : c).replace(/\t|\r|\n/g, ' ')).join('\t')).join('\n');
  }

  async function readSelectedRange() {
    if (!runtimeInfo || !runtimeInfo.inOffice || !window.Excel) {
      showAlert('warn', 'Đang ở chế độ xem trước ngoài Excel. Hãy nhập/dán bảng số liệu mẫu vào ô trên để thử AI.');
      return;
    }
    const result = await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values', 'formulas', 'rowCount', 'columnCount']);
      await context.sync();
      return {
        address: range.address,
        rows: range.rowCount,
        cols: range.columnCount,
        tsv: valuesToTsv(range.values)
      };
    });
    document.getElementById('txtExcelData').value = `[Vùng ${result.address} — ${result.rows} dòng x ${result.cols} cột]\n${result.tsv}`;
    showAlert('ok', `Đã đọc vùng ${result.address} (${result.rows} dòng x ${result.cols} cột).`);
  }

  async function readMultiSheetSummary() {
    if (!runtimeInfo || !runtimeInfo.inOffice || !window.Excel) {
      showAlert('warn', 'Cần mở trực tiếp trong Microsoft Excel (ExcelApi 1.4+) để đọc nhiều Sheet.');
      return;
    }
    const summary = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const blocks = [];
      for (let i = 0; i < Math.min(sheets.items.length, 6); i++) {
        const sh = sheets.items[i];
        const used = sh.getUsedRangeOrNullObject();
        used.load(['isNullObject', 'address', 'values', 'rowCount', 'columnCount']);
        await context.sync();
        if (!used.isNullObject) {
          const limited = (used.values || []).slice(0, 25).map(r => r.slice(0, 12));
          blocks.push(`=== SHEET: ${sh.name} (${used.rowCount}x${used.columnCount}) ===\n${valuesToTsv(limited)}`);
        }
      }
      return blocks.join('\n\n');
    });
    document.getElementById('txtExcelData').value = summary;
    showAlert('ok', 'Đã đọc tóm tắt dữ liệu từ các Sheet (tối đa 6 Sheet, 25 dòng đầu mỗi Sheet).');
  }

  async function writeResultToNewSheet(text) {
    if (!runtimeInfo || !runtimeInfo.inOffice || !window.Excel) {
      showAlert('warn', 'Chế độ xem trước ngoài Excel — vui lòng dùng nút Sao chép.');
      return;
    }
    if (!confirm('Xác nhận tạo Sheet báo cáo mới "VBTM_AI_BaoCao" để ghi kết quả (dữ liệu gốc hoàn toàn được giữ nguyên)?')) return;
    await Excel.run(async (context) => {
      const name = 'VBTM_AI_' + new Date().toTimeString().slice(0, 8).replace(/:/g, '');
      const sh = context.workbook.worksheets.add(name);
      const lines = text.split(/\r?\n/);
      const maxCols = Math.max(1, ...lines.map(l => l.split('\t').length));
      const matrix = lines.map(l => {
        const cols = l.split('\t');
        while (cols.length < maxCols) cols.push('');
        return cols;
      });
      const range = sh.getRangeByIndexes(0, 0, matrix.length, maxCols);
      range.values = matrix;
      sh.activate();
      await context.sync();
    });
    showAlert('ok', 'Đã ghi kết quả sang Sheet báo cáo mới an toàn!');
  }

  async function writeResultToActiveCell(text) {
    if (!runtimeInfo || !runtimeInfo.inOffice || !window.Excel) {
      showAlert('warn', 'Chế độ xem trước ngoài Excel.');
      return;
    }
    if (!confirm('Xác nhận ghi kết quả AI vào ô đang chọn trong Excel?')) return;
    await Excel.run(async (context) => {
      const cell = context.workbook.getActiveCell();
      cell.values = [[text]];
      await context.sync();
    });
    showAlert('ok', 'Đã ghi kết quả vào ô đang chọn.');
  }

  function initExcelUi() {
    runtimeInfo = window.VBTM_SettingsUI.bindCommonUi('Excel');
    const sel = document.getElementById('excelActionSelect');
    sel.innerHTML = window.VBTM_Prompts.EXCEL_ACTIONS.map(a => `<option value="${a.id}">${a.label}</option>`).join('');

    document.getElementById('btnReadRange').addEventListener('click', () => readSelectedRange().catch(e => showAlert('err', e.message)));
    document.getElementById('btnReadMultiSheet').addEventListener('click', () => readMultiSheetSummary().catch(e => showAlert('err', e.message)));

    const btnRun = document.getElementById('btnRunExcelAi');
    const btnCancel = document.getElementById('btnCancelExcelAi');
    btnRun.addEventListener('click', async () => {
      const dataText = document.getElementById('txtExcelData').value.trim();
      const extraQ = document.getElementById('txtExcelQuestion').value.trim();
      if (!dataText && !extraQ) {
        showAlert('warn', 'Vui lòng đọc vùng dữ liệu đang chọn hoặc nhập câu hỏi trước khi gọi AI.');
        return;
      }
      const act = window.VBTM_Prompts.EXCEL_ACTIONS.find(a => a.id === sel.value) || window.VBTM_Prompts.EXCEL_ACTIONS[0];
      const userPrompt = (extraQ ? `Yêu cầu bổ sung: ${extraQ}\n\n` : '') + dataText;
      showAlert('', '');
      btnRun.disabled = true;
      btnCancel.classList.remove('hidden');
      try {
        const bridge = window.VBTM_OfficeBridge ? window.VBTM_OfficeBridge.createBridge('Excel') : null;
        const res = bridge
          ? await bridge.runTask(act.id, act.prompt, userPrompt, { historyTitle: act.label })
          : await window.VBTM_AICore.generate(act.prompt, userPrompt, { host: 'Excel', historyTitle: act.label });
        document.getElementById('txtExcelOutput').value = res.text;
        document.getElementById('excelResultMeta').textContent = `${res.provider || 'gemini'} · ${res.model} · ${res.latencyMs}ms`;
        document.getElementById('excelResultCard').classList.remove('hidden');
      } catch (err) {
        showAlert('err', err.message || 'Lỗi gọi AI.');
      } finally {
        btnRun.disabled = false;
        btnCancel.classList.add('hidden');
      }
    });

    btnCancel.addEventListener('click', () => window.VBTM_AICore.cancelActive());
    document.getElementById('btnWriteNewSheet').addEventListener('click', () => writeResultToNewSheet(document.getElementById('txtExcelOutput').value).catch(e => showAlert('err', e.message)));
    document.getElementById('btnWriteActiveCell').addEventListener('click', () => writeResultToActiveCell(document.getElementById('txtExcelOutput').value).catch(e => showAlert('err', e.message)));
    document.getElementById('btnCopyExcelOut').addEventListener('click', async () => {
      try { await navigator.clipboard.writeText(document.getElementById('txtExcelOutput').value); showAlert('ok', 'Đã sao chép kết quả.'); } catch (e) {}
    });
  }

  if (typeof Office !== 'undefined' && Office.onReady) Office.onReady(() => initExcelUi());
  else window.addEventListener('DOMContentLoaded', initExcelUi);
})();
