/**
 * VBTM.AICore — Shared Task Pane Settings, Status Bar, History & Tech Log UI Controller
 * Dùng chung cho Word, Excel, PowerPoint, Outlook Add-ins.
 */
(function (global) {
  'use strict';

  function bindCommonUi(hostName) {
    const core = global.VBTM_AICore;
    const compat = global.VBTM_Compat;
    const runtime = compat.detectRuntime(hostName);

    // 1. Render Compatibility Tier badge & warnings
    const tierEl = document.getElementById('tierBadge');
    if (tierEl) tierEl.textContent = runtime.tierLabel;
    const compatBox = document.getElementById('compatWarnBox');
    if (compatBox) {
      if (runtime.warnings.length) {
        compatBox.className = 'alert-warn';
        compatBox.textContent = runtime.warnings.join(' ');
      } else {
        compatBox.className = 'hidden';
      }
    }

    // 2. Tabs
    document.querySelectorAll('.vbtm-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-tab');
        document.querySelectorAll('.vbtm-tab').forEach(b => b.classList.toggle('active', b === btn));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === target));
        if (target === 'tabHistory') renderHistory(hostName);
        if (target === 'tabSettings') renderLogs();
      });
    });

    // 3. Status bar subscription
    const statusBadge = document.getElementById('aiStatusBadge');
    const modelLabel = document.getElementById('aiModelLabel');
    function updateStatus(snap) {
      if (statusBadge) {
        statusBadge.className = 'badge badge-' + snap.status;
        const map = {
          idle: snap.mode === 'gateway' ? 'Gateway sẵn sàng' : (snap.hasKey ? 'AI sẵn sàng' : 'Chưa cấu hình AI'),
          connecting: 'Đang gọi AI...',
          connected: `Đã kết nối (${snap.lastLatencyMs}ms)`,
          error: 'Lỗi kết nối AI'
        };
        statusBadge.textContent = map[snap.status] || snap.status;
      }
      if (modelLabel) {
        modelLabel.textContent = `${snap.activeModel} · ${snap.mode === 'gateway' ? 'Gateway' : 'Key cá nhân'}`;
      }
    }
    core.onStateChange(updateStatus);
    updateStatus(core.getState());

    // 4. Settings controls
    const selMode = document.getElementById('cfgMode');
    const inpGw = document.getElementById('cfgGatewayUrl');
    const inpKey = document.getElementById('cfgApiKey');
    const selModel = document.getElementById('cfgModel');
    const keyNotice = document.getElementById('byokNotice');
    const gwGroup = document.getElementById('gatewayGroup');
    const keyGroup = document.getElementById('byokGroup');
    const cfgFeedback = document.getElementById('cfgFeedback');

    if (selModel) {
      selModel.innerHTML = core.PREFERRED_MODELS.map(m => `<option value="${m}">${m}</option>`).join('');
      selModel.value = core.getModel();
    }
    if (selMode) selMode.value = core.getMode();
    if (inpGw) inpGw.value = core.getGatewayUrl();
    if (inpKey) inpKey.placeholder = core.getApiKey() ? ('Đang lưu: ' + core.maskKey(core.getApiKey())) : 'AIza...';

    function toggleModeFields() {
      const m = selMode ? selMode.value : 'byok';
      if (gwGroup) gwGroup.classList.toggle('hidden', m !== 'gateway');
      if (keyGroup) keyGroup.classList.toggle('hidden', m !== 'byok');
      if (keyNotice) keyNotice.classList.toggle('hidden', m !== 'byok');
    }
    if (selMode) selMode.addEventListener('change', toggleModeFields);
    toggleModeFields();

    const btnSave = document.getElementById('btnSaveSettings');
    if (btnSave) {
      btnSave.addEventListener('click', () => {
        if (selMode) core.setMode(selMode.value);
        if (inpGw) core.setGatewayUrl(inpGw.value);
        if (inpKey && inpKey.value.trim()) {
          core.setApiKey(inpKey.value.trim());
          inpKey.value = '';
          inpKey.placeholder = 'Đang lưu: ' + core.maskKey(core.getApiKey());
        }
        if (selModel) core.setModel(selModel.value);
        if (cfgFeedback) {
          cfgFeedback.className = 'alert-ok';
          cfgFeedback.textContent = 'Đã lưu cấu hình AI (không hiển thị API Key đầy đủ trên giao diện).';
        }
      });
    }

    const btnClearKey = document.getElementById('btnClearKey');
    if (btnClearKey) {
      btnClearKey.addEventListener('click', () => {
        core.setApiKey('');
        if (inpKey) { inpKey.value = ''; inpKey.placeholder = 'AIza...'; }
        if (cfgFeedback) {
          cfgFeedback.className = 'alert-warn';
          cfgFeedback.textContent = 'Đã xóa API Key khỏi bộ nhớ cục bộ.';
        }
      });
    }

    const btnTest = document.getElementById('btnTestAi');
    if (btnTest) {
      btnTest.addEventListener('click', async () => {
        if (cfgFeedback) { cfgFeedback.className = 'alert-warn'; cfgFeedback.textContent = 'Đang kiểm tra kết nối Gemini AI...'; }
        btnTest.disabled = true;
        try {
          const r = await core.testConnection(hostName);
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-ok';
            cfgFeedback.textContent = `Kết nối thành công! Model: ${r.model} · Độ trễ: ${r.latencyMs}ms · Phản hồi: "${r.text}"`;
          }
        } catch (err) {
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-err';
            cfgFeedback.textContent = 'Thất bại: ' + (err.message || err);
          }
        } finally {
          btnTest.disabled = false;
          renderLogs();
        }
      });
    }

    renderHistory(hostName);
    renderLogs();
    return runtime;
  }

  function renderHistory(hostName) {
    const box = document.getElementById('historyList');
    if (!box) return;
    const list = global.VBTM_AICore.getHistory(hostName);
    if (!list.length) {
      box.innerHTML = '<div class="muted">Chưa có lịch sử thao tác AI trong phiên làm việc này.</div>';
      return;
    }
    box.innerHTML = list.map(h => `
      <div class="hist-item">
        <div><strong>${escapeHtml(h.action)}</strong> <span class="muted">(${escapeHtml(h.ts)})</span></div>
        <div class="muted">${escapeHtml(h.preview)}${h.outLen > 140 ? '…' : ''}</div>
      </div>
    `).join('');
  }

  function renderLogs() {
    const box = document.getElementById('techLogList');
    if (!box) return;
    const logs = global.VBTM_AICore.getTechLogs().slice(0, 12);
    if (!logs.length) {
      box.innerHTML = '<div class="muted">Chưa có log kỹ thuật.</div>';
      return;
    }
    box.innerHTML = logs.map(l => `
      <div class="log-item">
        <div><strong>[${escapeHtml(l.code)}]</strong> ${escapeHtml(l.model)} · HTTP ${l.status} · ${l.latencyMs}ms (retry=${l.retry})</div>
        <div class="muted">${escapeHtml(l.ts.slice(11, 19))} · ${escapeHtml(l.note)}</div>
      </div>
    `).join('');
  }

  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  global.VBTM_SettingsUI = { bindCommonUi, renderHistory, renderLogs, escapeHtml };
})(typeof window !== 'undefined' ? window : globalThis);
