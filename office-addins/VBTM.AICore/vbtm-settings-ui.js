/**
 * VBTM.AICore — Giao diện Quản lý AI Thống nhất (Mặc định: Ollama Local http://127.0.0.1:11434)
 * Dùng chung cho Word, Excel, PowerPoint, Outlook Add-ins.
 * - Chỉ hiển thị danh sách model thực tế lấy từ GET /api/tags (không đoán tên model)
 * - Nút Kiểm tra kết nối, Tải lại danh sách model, Chọn model, Thử AI bằng câu hỏi mẫu
 * - Gemini & OpenAI chỉ là tùy chọn bổ sung, KHÔNG bao giờ báo thiếu Gemini API Key khi đang dùng Ollama
 */
(function (global) {
  'use strict';

  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function injectAiCenterCard(hostName) {
    const tabSettings = document.getElementById('tabSettings');
    if (!tabSettings || document.getElementById('vbtmAiCenterCard')) return;

    const router = global.VBTM_SmartRouter;
    const card = document.createElement('div');
    card.id = 'vbtmAiCenterCard';
    card.className = 'card';
    card.innerHTML = `
      <div class="card-title">
        <span>QUẢN LÝ AI — OLLAMA LOCAL (MẶC ĐỊNH)</span>
        <span id="ollamaStatePill" class="badge badge-connecting">Đang kiểm tra...</span>
      </div>

      <label class="field-label" for="activeProviderSel">Nhà cung cấp AI đang chọn:</label>
      <select id="activeProviderSel">
        <option value="ollama">Ollama Local — http://127.0.0.1:11434 (Mặc định, không cần API Key)</option>
        <option value="gemini">Google Gemini API (Tùy chọn bổ sung — cần API Key)</option>
        <option value="openai">OpenAI API (Tùy chọn trả phí — cần API Key)</option>
      </select>

      <div id="ollamaDetailBox" class="alert-warn" style="margin-top:8px;margin-bottom:8px;">
        Đang kiểm tra kết nối tới Ollama Local tại <code>http://127.0.0.1:11434/api/tags</code>...
      </div>

      <label class="field-label" for="ollamaInstalledModelSel">Model thực tế đã cài trên máy (từ /api/tags):</label>
      <select id="ollamaInstalledModelSel">
        <option value="">-- Đang tải danh sách model thực tế --</option>
      </select>

      <div class="btn-row">
        <button id="btnCheckOllamaConn" class="btn btn-primary" type="button">Kiểm tra kết nối</button>
        <button id="btnReloadOllamaModels" class="btn btn-secondary" type="button">Tải lại danh sách model</button>
        <button id="btnSampleTestAi" class="btn btn-success" type="button">Thử AI câu mẫu</button>
      </div>
      <div class="btn-row" style="margin-top:6px;">
        <a href="https://ollama.com/download" target="_blank" class="btn btn-secondary" style="text-decoration:none;">Tải Ollama (Miễn phí)</a>
        <a href="https://aistudio.google.com/app/apikey" target="_blank" class="btn btn-secondary" style="text-decoration:none;">Lấy Key Gemini</a>
        <a href="https://platform.openai.com/api-keys" target="_blank" class="btn btn-secondary" style="text-decoration:none;">Lấy Key OpenAI</a>
      </div>
      <div id="sampleAiOutput" class="hidden" style="margin-top:8px;"></div>
    `;
    tabSettings.insertBefore(card, tabSettings.firstChild);

    const provSel = document.getElementById('activeProviderSel');
    const modelSel = document.getElementById('ollamaInstalledModelSel');
    if (router) provSel.value = router.getActiveProvider();

    provSel.addEventListener('change', () => {
      if (router) router.setActiveProvider(provSel.value);
      refreshOllamaUi(hostName);
    });

    modelSel.addEventListener('change', () => {
      if (router && modelSel.value) router.setOllamaModel(modelSel.value);
      refreshOllamaUi(hostName);
    });

    document.getElementById('btnCheckOllamaConn').addEventListener('click', async () => {
      if (router) router.resetCircuitBreaker();
      await refreshOllamaUi(hostName);
    });

    document.getElementById('btnReloadOllamaModels').addEventListener('click', async () => {
      await refreshOllamaUi(hostName);
    });

    document.getElementById('btnSampleTestAi').addEventListener('click', async () => {
      if (!router) return;
      const outEl = document.getElementById('sampleAiOutput');
      const btn = document.getElementById('btnSampleTestAi');
      btn.disabled = true;
      outEl.className = 'alert-warn';
      outEl.textContent = `Đang gửi câu hỏi mẫu tới ${router.getActiveProvider().toUpperCase()}...`;
      try {
        const res = await router.testActiveConnection(hostName);
        outEl.className = 'alert-ok';
        outEl.textContent = `Thành công! [${res.provider}] Model: ${res.model} (${res.latencyMs}ms) → "${res.text}"`;
      } catch (err) {
        outEl.className = 'alert-err';
        outEl.textContent = 'Lỗi thực tế: ' + (err.message || err);
      } finally {
        btn.disabled = false;
        renderLogs();
      }
    });
  }

  async function refreshOllamaUi(hostName) {
    const router = global.VBTM_SmartRouter;
    if (!router) return;
    const pill = document.getElementById('ollamaStatePill');
    const box = document.getElementById('ollamaDetailBox');
    const modelSel = document.getElementById('ollamaInstalledModelSel');
    const statusBadge = document.getElementById('aiStatusBadge');
    const modelLabel = document.getElementById('aiModelLabel');
    const activeProv = router.getActiveProvider();

    const st = await router.checkOllamaStatus();

    if (modelSel) {
      if (st.models && st.models.length > 0) {
        modelSel.innerHTML = st.models.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
        modelSel.value = st.selectedModel || st.models[0];
      } else {
        modelSel.innerHTML = '<option value="">(Chưa có model nào được cài trên máy)</option>';
      }
    }

    if (st.running && st.models.length > 0) {
      if (pill) { pill.className = 'badge badge-connected'; pill.textContent = 'Ollama: Đang hoạt động'; }
      if (box) {
        box.className = 'alert-ok';
        box.innerHTML = `<strong>Ollama Local đang hoạt động (${st.latencyMs}ms)!</strong><br/>Địa chỉ: <code>${escapeHtml(router.getOllamaUrl())}</code> · Model đang dùng: <strong>${escapeHtml(st.selectedModel)}</strong> (${st.models.length} model đã cài).`;
      }
      if (activeProv === 'ollama') {
        if (statusBadge) { statusBadge.className = 'badge badge-connected'; statusBadge.textContent = 'Ollama Local sẵn sàng'; }
        if (modelLabel) modelLabel.textContent = `${st.selectedModel} · Local`;
      }
    } else if (st.running && st.models.length === 0) {
      if (pill) { pill.className = 'badge badge-connecting'; pill.textContent = 'Ollama: Chưa cài model'; }
      if (box) {
        box.className = 'alert-warn';
        box.innerHTML = `<strong>Ollama đang chạy nhưng chưa có model nào.</strong><br/>Mở PowerShell/CMD và chạy lệnh: <code>ollama pull qwen2.5:3b</code> (hoặc <code>ollama pull qwen2.5:7b</code>), sau đó bấm <strong>Tải lại danh sách model</strong>.`;
      }
      if (activeProv === 'ollama' && statusBadge) {
        statusBadge.className = 'badge badge-connecting';
        statusBadge.textContent = 'Ollama: Chưa có model';
      }
    } else {
      if (pill) { pill.className = 'badge badge-error'; pill.textContent = 'Ollama: Không kết nối được'; }
      if (box) {
        box.className = 'alert-err';
        box.innerHTML = `<strong>${escapeHtml(st.error || 'Không kết nối được Ollama Local tại http://127.0.0.1:11434.')}</strong><br/>Hướng dẫn: Cài Ollama từ <a href="https://ollama.com/download/windows" target="_blank">ollama.com/download/windows</a>, khởi chạy Ollama, rồi chạy <code>ollama pull qwen2.5:3b</code> và bấm <strong>Kiểm tra kết nối</strong>.`;
      }
      if (activeProv === 'ollama' && statusBadge) {
        statusBadge.className = 'badge badge-error';
        statusBadge.textContent = 'Ollama chưa chạy';
      }
    }

    if (activeProv !== 'ollama' && statusBadge) {
      statusBadge.className = 'badge badge-idle';
      statusBadge.textContent = `Đang chọn: ${activeProv.toUpperCase()} (tùy chọn)`;
    }
  }

  function bindCommonUi(hostName) {
    const core = global.VBTM_AICore;
    const compat = global.VBTM_Compat;
    const runtime = compat.detectRuntime(hostName);

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

    injectAiCenterCard(hostName);

    document.querySelectorAll('.vbtm-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-tab');
        document.querySelectorAll('.vbtm-tab').forEach(b => b.classList.toggle('active', b === btn));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === target));
        if (target === 'tabHistory') renderHistory(hostName);
        if (target === 'tabSettings') { renderLogs(); refreshOllamaUi(hostName); }
      });
    });

    // Optional Cloud Settings (Gemini/OpenAI — chỉ dùng khi người dùng chủ động chọn)
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
    if (inpKey) inpKey.placeholder = core.getApiKey() ? ('Đang lưu: ' + core.maskKey(core.getApiKey())) : '(Tùy chọn bổ sung — không bắt buộc khi dùng Ollama)';

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
        if (global.VBTM_SmartRouter) global.VBTM_SmartRouter.resetCircuitBreaker('gemini');
        if (cfgFeedback) {
          cfgFeedback.className = 'alert-ok';
          cfgFeedback.textContent = 'Đã lưu cấu hình bổ sung và đặt lại trạng thái kết nối.';
        }
      });
    }

    const btnClearKey = document.getElementById('btnClearKey');
    if (btnClearKey) {
      btnClearKey.addEventListener('click', () => {
        core.setApiKey('');
        if (inpKey) { inpKey.value = ''; inpKey.placeholder = '(Tùy chọn bổ sung)'; }
        if (global.VBTM_SmartRouter) global.VBTM_SmartRouter.resetCircuitBreaker('gemini');
        if (cfgFeedback) {
          cfgFeedback.className = 'alert-warn';
          cfgFeedback.textContent = 'Đã xóa Gemini API Key — Ollama Local vẫn hoạt động bình thường.';
        }
      });
    }

    const btnTest = document.getElementById('btnTestAi');
    if (btnTest) {
      btnTest.addEventListener('click', async () => {
        if (cfgFeedback) { cfgFeedback.className = 'alert-warn'; cfgFeedback.textContent = 'Đang kiểm tra kết nối qua AI Gateway...'; }
        btnTest.disabled = true;
        try {
          const r = await core.testConnection(hostName);
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-ok';
            cfgFeedback.textContent = `Kết nối thành công! [${r.provider}] Model: ${r.model} (${r.latencyMs}ms) → "${r.text}"`;
          }
        } catch (err) {
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-err';
            cfgFeedback.textContent = 'Lỗi kết nối: ' + (err.message || err);
          }
        } finally {
          btnTest.disabled = false;
          renderLogs();
        }
      });
    }

    refreshOllamaUi(hostName);
    renderHistory(hostName);
    renderLogs();
    return runtime;
  }

  function renderHistory(hostName) {
    const box = document.getElementById('historyList');
    if (!box) return;
    const list = global.VBTM_SyncEngine ? global.VBTM_SyncEngine.getHistory(hostName) : global.VBTM_AICore.getHistory(hostName);
    if (!list.length) {
      box.innerHTML = '<div class="muted">Chưa có lịch sử thao tác AI.</div>';
      return;
    }
    box.innerHTML = list.map(h => `
      <div class="hist-item">
        <div><strong>${escapeHtml(h.action)}</strong> <span class="muted">(${escapeHtml((h.ts || '').slice(11, 19) || h.ts)})</span></div>
        <div class="muted">${escapeHtml(h.preview)}</div>
      </div>
    `).join('');
  }

  function renderLogs() {
    const box = document.getElementById('techLogList');
    if (!box) return;
    const audit = global.VBTM_SmartRouter ? global.VBTM_SmartRouter.getRequestAuditLog().slice(0, 12) : [];
    if (!audit.length) {
      box.innerHTML = '<div class="muted">Chưa có nhật ký gọi AI.</div>';
      return;
    }
    box.innerHTML = audit.map(l => `
      <div class="log-item">
        <div><strong>[${escapeHtml(l.provider.toUpperCase())} · ${escapeHtml(l.status)}]</strong> ${escapeHtml(l.model)} · ${l.latencyMs}ms</div>
        <div class="muted">${escapeHtml(l.ts.slice(11, 19))} · ${escapeHtml(l.note)}</div>
      </div>
    `).join('');
  }

  global.VBTM_SettingsUI = { bindCommonUi, refreshOllamaUi, renderHistory, renderLogs, escapeHtml };
})(typeof window !== 'undefined' ? window : globalThis);
