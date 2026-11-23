/**
 * VBTM.AICore — Unified AI Center (TRUNG TÂM AI) & Task Pane Controller
 * Dùng chung cho Word, Excel, PowerPoint, Outlook Add-ins.
 * - Tự động kiểm tra Ollama Local khi mở lần đầu (không đòi Gemini API Key)
 * - Hiển thị trạng thái rõ ràng: Đang hoạt động | Chưa tải model | Chưa cài đặt | Đang tải
 * - Cho phép chọn chính sách Local Only / Ollama First / Gemini bổ sung / OpenAI trả phí
 * - Tải model Ollama có xác nhận dung lượng & theo dõi tiến trình thực
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
    const rec = router ? router.recommendModelForHardware() : { name: 'qwen2.5:7b', sizeGb: 4.7, minRamGb: 10 };
    const catalogOpts = router
      ? router.MODEL_CATALOG.map(m => `<option value="${m.name}">${m.name} (${m.sizeGb}GB · RAM ≥ ${m.minRamGb}GB · ${m.viQuality})</option>`).join('')
      : '<option value="qwen2.5:7b">qwen2.5:7b (4.7GB)</option>';

    const card = document.createElement('div');
    card.id = 'vbtmAiCenterCard';
    card.className = 'card';
    card.innerHTML = `
      <div class="card-title">
        <span>TRUNG TÂM AI (Mặc định không cần API Key)</span>
        <span id="ollamaStatePill" class="badge badge-connecting">Đang kiểm tra Ollama...</span>
      </div>
      <div id="ollamaDetailBox" class="alert-warn" style="margin-bottom:8px;">
        Đang tự động kiểm tra dịch vụ Ollama Local tại <code>http://localhost:11434</code>...
      </div>

      <label class="field-label" for="routerPolicySel">Chính sách điều phối AI (Smart AI Router):</label>
      <select id="routerPolicySel">
        <option value="ollama-first">Mặc định — Ưu tiên Ollama Local (Không cần API Key)</option>
        <option value="local-only">Chỉ dùng AI cục bộ (Local Only — không gửi dữ liệu ra Cloud)</option>
        <option value="gemini-first">Ưu tiên Google Gemini (nếu đã nhập Key bổ sung), dự phòng Ollama</option>
        <option value="ask-before-paid">Hỏi trước khi dùng OpenAI trả phí</option>
      </select>

      <label class="field-label" for="ollamaModelSel">Model tiếng Việt chạy cục bộ (Đề xuất cho máy này: <strong>${escapeHtml(rec.name)}</strong>):</label>
      <select id="ollamaModelSel">${catalogOpts}</select>

      <div class="btn-row">
        <button id="btnRecheckOllama" class="btn btn-secondary" type="button">Kiểm tra Ollama</button>
        <button id="btnPullOllamaModel" class="btn btn-blue" type="button">Tải model đề xuất</button>
        <button id="btnTestGatewayAi" class="btn btn-success" type="button">Kiểm tra kết nối AI</button>
        <a href="https://ollama.com/download/windows" target="_blank" class="btn btn-secondary" style="text-decoration:none;">Cài đặt Ollama</a>
      </div>
      <div id="ollamaPullProgress" class="muted" style="margin-top:6px;"></div>
    `;
    tabSettings.insertBefore(card, tabSettings.firstChild);

    // Wire events
    const policySel = document.getElementById('routerPolicySel');
    const modelSel = document.getElementById('ollamaModelSel');
    if (router) {
      policySel.value = router.getPolicy();
      modelSel.value = router.getOllamaModel() || rec.name;
    }

    policySel.addEventListener('change', () => {
      if (router) router.setPolicy(policySel.value);
      refreshOllamaUi(hostName);
    });
    modelSel.addEventListener('change', () => {
      if (router) router.setOllamaModel(modelSel.value);
    });

    document.getElementById('btnRecheckOllama').addEventListener('click', () => refreshOllamaUi(hostName));

    document.getElementById('btnPullOllamaModel').addEventListener('click', async () => {
      if (!router) return;
      const chosenName = modelSel.value || rec.name;
      const info = router.MODEL_CATALOG.find(m => m.name === chosenName) || rec;
      if (!confirm(`Xác nhận tải model tiếng Việt "${chosenName}" (~${info.sizeGb}GB, yêu cầu RAM ≥ ${info.minRamGb}GB) về Ollama cục bộ trên máy này?`)) return;
      const progEl = document.getElementById('ollamaPullProgress');
      const btnPull = document.getElementById('btnPullOllamaModel');
      btnPull.disabled = true;
      progEl.textContent = `Đang bắt đầu tải ${chosenName}...`;
      try {
        const res = await router.pullOllamaModel(chosenName, (p) => {
          progEl.textContent = p.percent != null ? `${p.status} (${p.percent}%)` : p.status;
        });
        progEl.textContent = res.alreadyInstalled
          ? `Model "${chosenName}" đã có sẵn trên máy — không cần tải lại.`
          : `Đã tải xong model "${chosenName}"! Đang tự động kiểm tra sẵn sàng...`;
        await refreshOllamaUi(hostName);
      } catch (err) {
        progEl.textContent = 'Lỗi tải model: ' + (err.message || err);
      } finally {
        btnPull.disabled = false;
      }
    });

    document.getElementById('btnTestGatewayAi').addEventListener('click', async () => {
      const box = document.getElementById('ollamaDetailBox');
      const btn = document.getElementById('btnTestGatewayAi');
      btn.disabled = true;
      box.className = 'alert-warn';
      box.textContent = 'Đang gửi câu thử nghiệm thực tế qua VBTM AI Gateway...';
      try {
        const r = await router.testActiveConnection(hostName);
        box.className = 'alert-ok';
        box.textContent = `Kết nối thành công! AI: ${r.provider} (${r.connectionType || 'Local'}) · Model: ${r.model} · Độ trễ: ${r.latencyMs}ms · Phản hồi: "${r.text}"`;
      } catch (err) {
        box.className = 'alert-err';
        box.textContent = 'Chưa kết nối được AI: ' + (err.message || err);
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
    const statusBadge = document.getElementById('aiStatusBadge');
    const modelLabel = document.getElementById('aiModelLabel');

    const st = await router.checkOllamaStatus();
    const hasGem = !!(global.VBTM_AICore && global.VBTM_AICore.getApiKey());

    if (st.running && st.models.length > 0) {
      if (pill) { pill.className = 'badge badge-connected'; pill.textContent = 'Ollama: Đang hoạt động (Local)'; }
      if (box) {
        box.className = 'alert-ok';
        box.innerHTML = `<strong>Sẵn sàng sử dụng AI không cần API Key!</strong><br/>Ollama Local đang chạy tại <code>${escapeHtml(router.getOllamaUrl())}</code> · Model sẵn sàng: <strong>${escapeHtml(st.selectedModel)}</strong> (${st.models.length} model đã cài).`;
      }
      if (statusBadge) { statusBadge.className = 'badge badge-connected'; statusBadge.textContent = 'AI Local sẵn sàng'; }
      if (modelLabel) modelLabel.textContent = `${st.selectedModel} · Ollama Local (Không cần Key)`;
    } else if (st.running && st.models.length === 0) {
      const rec = st.recommended || router.recommendModelForHardware();
      if (pill) { pill.className = 'badge badge-connecting'; pill.textContent = 'Ollama: Chưa tải model'; }
      if (box) {
        box.className = 'alert-warn';
        box.innerHTML = `<strong>Ollama đã chạy nhưng chưa có model tiếng Việt.</strong> Bấm <em>"Tải model đề xuất"</em> để tải <strong>${escapeHtml(rec.name)}</strong> (~${rec.sizeGb}GB, RAM ≥ ${rec.minRamGb}GB) và dùng AI miễn phí không cần API Key.`;
      }
      if (statusBadge) { statusBadge.className = 'badge badge-connecting'; statusBadge.textContent = 'Chưa tải model Ollama'; }
    } else {
      if (pill) { pill.className = 'badge badge-idle'; pill.textContent = 'Ollama: Chưa cài / Chưa chạy'; }
      if (box) {
        box.className = hasGem ? 'alert-ok' : 'alert-warn';
        box.innerHTML = hasGem
          ? `Ollama Local chưa chạy, nhưng bạn đã cấu hình <strong>Google Gemini API Key bổ sung</strong> — hệ thống vẫn dùng được AI Cloud bình thường.`
          : `<strong>Chưa phát hiện Ollama tại <code>http://localhost:11434</code>.</strong> Để dùng AI miễn phí <strong>không cần API Key</strong>: bấm <em>"Cài đặt Ollama"</em> (hoặc chạy <code>VBTM-Pro-Unified-Setup.exe</code>) rồi khởi chạy Ollama. Hoặc nhập Gemini Key bổ sung bên dưới nếu muốn dùng Cloud.`;
      }
      if (statusBadge && !hasGem) { statusBadge.className = 'badge badge-idle'; statusBadge.textContent = 'Cần cài Ollama (không cần Key)'; }
    }
  }

  function bindCommonUi(hostName) {
    const core = global.VBTM_AICore;
    const compat = global.VBTM_Compat;
    const runtime = compat.detectRuntime(hostName);

    // 1. Compatibility Tier badge & warnings
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

    // 2. Inject Unified AI Center card
    injectAiCenterCard(hostName);

    // 3. Tabs
    document.querySelectorAll('.vbtm-tab').forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.getAttribute('data-tab');
        document.querySelectorAll('.vbtm-tab').forEach(b => b.classList.toggle('active', b === btn));
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === target));
        if (target === 'tabHistory') renderHistory(hostName);
        if (target === 'tabSettings') { renderLogs(); refreshOllamaUi(hostName); }
      });
    });

    // 4. Status bar subscription
    const statusBadge = document.getElementById('aiStatusBadge');
    const modelLabel = document.getElementById('aiModelLabel');
    function updateStatus(snap) {
      if (statusBadge && snap.status === 'connecting') {
        statusBadge.className = 'badge badge-connecting';
        statusBadge.textContent = 'Đang gọi AI...';
      } else if (statusBadge && snap.status === 'connected') {
        statusBadge.className = 'badge badge-connected';
        statusBadge.textContent = `Đã kết nối (${snap.lastLatencyMs}ms)`;
      } else if (statusBadge && snap.status === 'error') {
        statusBadge.className = 'badge badge-error';
        statusBadge.textContent = 'AI chưa sẵn sàng';
      }
      if (modelLabel && snap.status === 'connected') {
        modelLabel.textContent = `${snap.activeModel} · ${snap.mode === 'gateway' ? 'Gateway' : 'Cloud bổ sung'}`;
      }
    }
    core.onStateChange(updateStatus);

    // 5. Optional Cloud Settings controls (Gemini Key is OPTIONAL, never mandatory)
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
    if (inpKey) inpKey.placeholder = core.getApiKey() ? ('Đang lưu: ' + core.maskKey(core.getApiKey())) : '(Tùy chọn — không bắt buộc nếu dùng Ollama Local)';

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
          cfgFeedback.textContent = 'Đã lưu cấu hình AI.';
        }
        refreshOllamaUi(hostName);
      });
    }

    const btnClearKey = document.getElementById('btnClearKey');
    if (btnClearKey) {
      btnClearKey.addEventListener('click', () => {
        core.setApiKey('');
        if (inpKey) { inpKey.value = ''; inpKey.placeholder = '(Tùy chọn — không bắt buộc nếu dùng Ollama Local)'; }
        if (cfgFeedback) {
          cfgFeedback.className = 'alert-warn';
          cfgFeedback.textContent = 'Đã xóa Gemini API Key (vẫn dùng được AI qua Ollama Local).';
        }
        refreshOllamaUi(hostName);
      });
    }

    const btnTest = document.getElementById('btnTestAi');
    if (btnTest) {
      btnTest.addEventListener('click', async () => {
        if (cfgFeedback) { cfgFeedback.className = 'alert-warn'; cfgFeedback.textContent = 'Đang kiểm tra kết nối qua VBTM AI Gateway...'; }
        btnTest.disabled = true;
        try {
          const r = await core.testConnection(hostName);
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-ok';
            cfgFeedback.textContent = `Kết nối thành công! AI: ${r.provider || 'gateway'} · Model: ${r.model} · Độ trễ: ${r.latencyMs}ms · Phản hồi: "${r.text}"`;
          }
        } catch (err) {
          if (cfgFeedback) {
            cfgFeedback.className = 'alert-err';
            cfgFeedback.textContent = 'Chưa kết nối được AI: ' + (err.message || err);
          }
        } finally {
          btnTest.disabled = false;
          renderLogs();
        }
      });
    }

    // Tự động kiểm tra Ollama ngay khi mở Add-in lần đầu
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
      box.innerHTML = '<div class="muted">Chưa có lịch sử thao tác AI trong phiên làm việc này.</div>';
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

  global.VBTM_SettingsUI = { bindCommonUi, refreshOllamaUi, renderHistory, renderLogs, escapeHtml };
})(typeof window !== 'undefined' ? window : globalThis);
