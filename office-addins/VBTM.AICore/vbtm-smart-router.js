/**
 * VBTM.SmartRouter — VBTM AI Gateway & Trung tâm điều phối AI mặc định KHÔNG CẦN API KEY
 * Thứ tự ưu tiên mặc định:
 *   1. Ollama Local (http://localhost:11434) — KHÔNG CẦN API KEY, chạy trực tiếp trên Windows
 *   2. Google Gemini API (chỉ dùng nếu người dùng đã chủ động cấu hình Key và chính sách cho phép)
 *   3. OpenAI API chính thức (chỉ dùng nếu người dùng đã cấu hình Key và xác nhận cho phép trả phí)
 *
 * Bảo mật & Trung thực:
 * - Chế độ 'local-only': TUYỆT ĐỐI không gửi dữ liệu sang Cloud (Gemini/OpenAI).
 * - Không tự ý tải model dung lượng lớn khi chưa có xác nhận người dùng.
 * - Không giả lập phản hồi AI; không báo lỗi chặn "Chưa nhập Gemini API Key" cho toàn hệ thống.
 */
(function (global) {
  'use strict';

  const KEYS = {
    POLICY: 'vbtm_router_policy',           // 'ollama-first' (mặc định) | 'local-only' | 'gemini-first' | 'ask-before-paid' | 'manual'
    MANUAL_PROV: 'vbtm_router_manual_prov', // 'ollama' | 'gemini' | 'openai'
    OPENAI_KEY: 'vbtm_office_openai_key',
    OPENAI_MODEL: 'vbtm_office_openai_model',
    OLLAMA_URL: 'vbtm_office_ollama_url',
    OLLAMA_MODEL: 'vbtm_office_ollama_model',
    ALLOW_PAID: 'vbtm_router_allow_paid'
  };

  // Danh mục model tiếng Việt được đánh giá theo cấu hình máy & loại tác vụ
  const MODEL_CATALOG = [
    {
      name: 'qwen2.5:3b',
      sizeGb: 1.9,
      minRamGb: 6,
      tier: 'low',
      viQuality: 'Khá tốt (nhanh)',
      bestFor: ['spell', 'summary', 'extract'],
      note: 'Máy cấu hình thấp (RAM ≤ 8GB): soát lỗi chính tả nhanh, tóm tắt ngắn.'
    },
    {
      name: 'qwen2.5:7b',
      sizeGb: 4.7,
      minRamGb: 10,
      tier: 'medium',
      viQuality: 'Rất tốt (chuẩn tiếng Việt công vụ)',
      bestFor: ['draft', 'spell', 'summary', 'table', 'extract'],
      note: 'Máy cấu hình trung bình (RAM 12–16GB): khuyến nghị mặc định cho soạn thảo, soát lỗi ngữ cảnh & phân tích bảng.'
    },
    {
      name: 'gemma2:9b',
      sizeGb: 5.5,
      minRamGb: 12,
      tier: 'medium',
      viQuality: 'Tốt (văn phong mạch lạc)',
      bestFor: ['draft', 'summary', 'rewrite'],
      note: 'Máy RAM 16GB: văn phong mượt, tóm tắt rõ ý.'
    },
    {
      name: 'qwen2.5:14b',
      sizeGb: 9.0,
      minRamGb: 16,
      tier: 'high',
      viQuality: 'Xuất sắc (tài liệu dài & tham mưu)',
      bestFor: ['long_analysis', 'draft', 'table', 'spell'],
      note: 'Máy cấu hình cao (RAM ≥ 24GB / VRAM ≥ 12GB): phân tích tài liệu dài, tham mưu chuyên sâu.'
    }
  ];

  const circuitState = {
    ollama: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    gemini: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    openai: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' }
  };

  const runtimeDiscovery = {
    status: 'checking', // 'ready' | 'no_model' | 'not_running' | 'pulling' | 'checking'
    ollamaRunning: false,
    installedModels: [],
    selectedModel: '',
    estimatedRamGb: (typeof navigator !== 'undefined' && navigator.deviceMemory) ? navigator.deviceMemory : 8,
    lastCheckAt: 0,
    errorDetail: ''
  };

  function getStr(k, def) {
    try { return (localStorage.getItem(k) || def).trim(); } catch (e) { return def; }
  }
  function setStr(k, v) {
    try {
      if (v) localStorage.setItem(k, String(v).trim());
      else localStorage.removeItem(k);
    } catch (e) {}
  }

  function getPolicy() { return getStr(KEYS.POLICY, 'ollama-first'); }
  function setPolicy(p) { setStr(KEYS.POLICY, p); }
  function getManualProvider() { return getStr(KEYS.MANUAL_PROV, 'ollama'); }
  function setManualProvider(p) { setStr(KEYS.MANUAL_PROV, p); }
  function getOpenAiKey() { return getStr(KEYS.OPENAI_KEY, ''); }
  function setOpenAiKey(k) { setStr(KEYS.OPENAI_KEY, k); }
  function getOpenAiModel() { return getStr(KEYS.OPENAI_MODEL, 'gpt-4o-mini'); }
  function setOpenAiModel(m) { setStr(KEYS.OPENAI_MODEL, m); }
  function getOllamaUrl() { return getStr(KEYS.OLLAMA_URL, 'http://localhost:11434'); }
  function setOllamaUrl(u) { setStr(KEYS.OLLAMA_URL, u); }
  function getOllamaModel() { return getStr(KEYS.OLLAMA_MODEL, runtimeDiscovery.selectedModel || 'qwen2.5:7b'); }
  function setOllamaModel(m) { setStr(KEYS.OLLAMA_MODEL, m); runtimeDiscovery.selectedModel = m; }

  function recommendModelForHardware(ramGb) {
    const r = Number(ramGb) || runtimeDiscovery.estimatedRamGb || 8;
    if (r >= 24) return MODEL_CATALOG.find(m => m.name === 'qwen2.5:14b') || MODEL_CATALOG[1];
    if (r >= 10) return MODEL_CATALOG.find(m => m.name === 'qwen2.5:7b') || MODEL_CATALOG[1];
    return MODEL_CATALOG.find(m => m.name === 'qwen2.5:3b') || MODEL_CATALOG[0];
  }

  function pickBestInstalledModelForTask(installedModels, taskType) {
    if (!Array.isArray(installedModels) || !installedModels.length) return '';
    const saved = getStr(KEYS.OLLAMA_MODEL, '');
    if (saved && installedModels.includes(saved)) return saved;

    // Ưu tiên model tiếng Việt phù hợp loại tác vụ
    const task = String(taskType || 'general').toLowerCase();
    const preferList = (task === 'spell' || task === 'summary')
      ? ['qwen2.5:7b', 'qwen2.5:3b', 'gemma2:9b', 'qwen2.5:14b']
      : ['qwen2.5:14b', 'qwen2.5:7b', 'gemma2:9b', 'qwen2.5:3b'];

    for (const pref of preferList) {
      const hit = installedModels.find(m => m.toLowerCase().startsWith(pref.toLowerCase()));
      if (hit) return hit;
    }
    return installedModels[0];
  }

  function isCircuitOpen(prov) {
    const c = circuitState[prov];
    return !!(c && c.openUntil > Date.now());
  }

  function recordSuccess(prov, ms) {
    const c = circuitState[prov];
    if (!c) return;
    c.fails = 0; c.openUntil = 0; c.lastLatencyMs = ms || 0; c.lastError = '';
  }

  function recordFailure(prov, errMsg) {
    const c = circuitState[prov];
    if (!c) return;
    c.fails += 1;
    c.lastError = String(errMsg || '').slice(0, 180);
    if (c.fails >= 3) c.openUntil = Date.now() + 60000;
  }

  async function checkOllamaStatus() {
    const base = getOllamaUrl().replace(/\/+$/, '');
    const t0 = Date.now();
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 3500);
      const res = await fetch(base + '/api/tags', { method: 'GET', signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) {
        runtimeDiscovery.status = 'not_running';
        runtimeDiscovery.ollamaRunning = false;
        runtimeDiscovery.errorDetail = 'Ollama phản hồi HTTP ' + res.status;
        return { running: false, status: 'not_running', models: [], error: runtimeDiscovery.errorDetail };
      }
      const data = await res.json();
      const models = (data && Array.isArray(data.models)) ? data.models.map(m => m.name) : [];
      runtimeDiscovery.ollamaRunning = true;
      runtimeDiscovery.installedModels = models;
      runtimeDiscovery.lastCheckAt = Date.now();
      if (models.length > 0) {
        runtimeDiscovery.status = 'ready';
        runtimeDiscovery.selectedModel = pickBestInstalledModelForTask(models, 'general');
        if (!getStr(KEYS.OLLAMA_MODEL, '')) setOllamaModel(runtimeDiscovery.selectedModel);
        recordSuccess('ollama', Date.now() - t0);
      } else {
        runtimeDiscovery.status = 'no_model';
      }
      return {
        running: true,
        status: runtimeDiscovery.status,
        models,
        selectedModel: runtimeDiscovery.selectedModel,
        recommended: recommendModelForHardware(),
        latencyMs: Date.now() - t0
      };
    } catch (e) {
      runtimeDiscovery.ollamaRunning = false;
      runtimeDiscovery.status = 'not_running';
      runtimeDiscovery.installedModels = [];
      runtimeDiscovery.errorDetail = 'Chưa phát hiện dịch vụ Ollama tại ' + base + ' (chưa cài đặt hoặc dịch vụ chưa khởi chạy).';
      return {
        running: false,
        status: 'not_running',
        models: [],
        recommended: recommendModelForHardware(),
        error: runtimeDiscovery.errorDetail
      };
    }
  }

  // Tải model Ollama có xác nhận người dùng & theo dõi tiến trình thực
  async function pullOllamaModel(modelName, onProgress) {
    const m = String(modelName || recommendModelForHardware().name).trim();
    const base = getOllamaUrl().replace(/\/+$/, '');
    // Không tải lại nếu máy đã có model tương thích
    const st = await checkOllamaStatus();
    if (!st.running) throw new Error('Ollama chưa chạy tại ' + base + '. Hãy cài đặt và khởi động Ollama trước khi tải model.');
    if (st.models.some(x => x === m || x.startsWith(m + ':'))) {
      setOllamaModel(m);
      return { alreadyInstalled: true, model: m };
    }

    runtimeDiscovery.status = 'pulling';
    const res = await fetch(base + '/api/pull', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: m, stream: true })
    });
    if (!res.ok) throw new Error('Không thể tải model "' + m + '" (HTTP ' + res.status + ').');

    const reader = res.body && res.body.getReader ? res.body.getReader() : null;
    if (reader) {
      const dec = new TextDecoder();
      let buf = '';
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buf += dec.decode(value, { stream: true });
        const lines = buf.split('\n');
        buf = lines.pop() || '';
        for (const line of lines) {
          if (!line.trim()) continue;
          try {
            const j = JSON.parse(line);
            const pct = (j.total && j.completed) ? Math.round((j.completed / j.total) * 100) : null;
            if (typeof onProgress === 'function') onProgress({ status: j.status || 'Đang tải...', percent: pct });
          } catch (e) {}
        }
      }
    }
    setOllamaModel(m);
    await checkOllamaStatus();
    return { alreadyInstalled: false, model: m };
  }

  async function callOllamaLocal(systemPrompt, userText, opts) {
    const st = await checkOllamaStatus();
    if (!st.running) {
      throw new Error('Ollama Local chưa chạy tại ' + getOllamaUrl() + '. Mở Trung tâm AI để cài đặt hoặc khởi động Ollama (không cần API Key).');
    }
    if (!st.models.length) {
      const rec = recommendModelForHardware();
      throw new Error(`Ollama đã chạy nhưng chưa có model tiếng Việt. Mở Trung tâm AI để tải model đề xuất "${rec.name}" (~${rec.sizeGb}GB, RAM ≥ ${rec.minRamGb}GB).`);
    }

    const base = getOllamaUrl().replace(/\/+$/, '');
    const model = pickBestInstalledModelForTask(st.models, opts && opts.taskType);
    const t0 = Date.now();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeoutMs) || 60000);
    try {
      const res = await fetch(base + '/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model,
          prompt: (systemPrompt ? systemPrompt + '\n\n' : '') + userText,
          stream: false
        }),
        signal: ctrl.signal
      });
      if (!res.ok) {
        const errBody = await res.text().catch(() => '');
        if (/memory|ram|oom/i.test(errBody)) {
          throw new Error(`Model "${model}" vượt quá RAM khả dụng. Hãy chuyển sang model nhẹ hơn "qwen2.5:3b" trong Trung tâm AI.`);
        }
        throw new Error(`[Ollama HTTP ${res.status}] ${errBody.slice(0, 160)}`);
      }
      const data = await res.json();
      const out = String((data && data.response) || '').trim();
      if (!out) throw new Error('Ollama trả về nội dung rỗng.');
      const ms = Date.now() - t0;
      recordSuccess('ollama', ms);
      return { provider: 'ollama', connectionType: 'Local (Ngoại tuyến)', model, text: out, latencyMs: ms };
    } catch (err) {
      recordFailure('ollama', err.message);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  async function callOpenAiDirect(systemPrompt, userText, opts) {
    const key = getOpenAiKey();
    if (!key) throw new Error('OpenAI API chưa được cấu hình Key.');
    const model = getOpenAiModel();
    const t0 = Date.now();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeoutMs) || 28000);
    try {
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + key },
        body: JSON.stringify({
          model,
          temperature: opts && opts.expectJson ? 0.2 : 0.4,
          messages: [
            { role: 'system', content: systemPrompt || 'Bạn là trợ lý văn phòng VBTM Pro.' },
            { role: 'user', content: userText }
          ]
        }),
        signal: ctrl.signal
      });
      const txt = await res.text();
      let parsed = null;
      try { parsed = JSON.parse(txt); } catch (e) {}
      if (!res.ok) {
        const msg = (parsed && parsed.error && parsed.error.message) ? parsed.error.message : txt.slice(0, 180);
        throw new Error(`[OpenAI HTTP ${res.status}] ${msg}`);
      }
      const out = parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].message
        ? String(parsed.choices[0].message.content || '').trim() : '';
      if (!out) throw new Error('OpenAI trả về nội dung rỗng.');
      const ms = Date.now() - t0;
      recordSuccess('openai', ms);
      return { provider: 'openai', connectionType: 'Cloud (OpenAI)', model, text: out, latencyMs: ms };
    } catch (err) {
      recordFailure('openai', err.message);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  function hasGeminiConfigured() {
    return !!(global.VBTM_AICore && (global.VBTM_AICore.getMode() === 'gateway' || global.VBTM_AICore.getApiKey()));
  }

  function resolveProviderOrder(opts) {
    const policy = (opts && opts.policy) || getPolicy();
    if (policy === 'local-only' || (opts && opts.sensitiveDocument)) return ['ollama'];
    if (policy === 'manual') return [getManualProvider()];
    if (policy === 'gemini-first') {
      return hasGeminiConfigured() ? ['gemini', 'ollama'] : ['ollama'];
    }
    // Mặc định 'ollama-first': Ưu tiên Ollama Local không cần API Key -> Gemini (nếu đã cấu hình Key) -> OpenAI (chỉ khi ask-before-paid)
    const order = ['ollama'];
    if (hasGeminiConfigured()) order.push('gemini');
    if (policy === 'ask-before-paid' && getOpenAiKey()) order.push('openai');
    return order;
  }

  async function route(systemPrompt, userText, opts) {
    opts = opts || {};
    const policy = opts.policy || getPolicy();
    const order = resolveProviderOrder(opts);
    let lastErr = null;

    for (const prov of order) {
      if (isCircuitOpen(prov)) {
        lastErr = new Error(`Nhà cung cấp "${prov}" đang tạm ngắt mạch 60s do lỗi liên tiếp.`);
        continue;
      }

      if (prov === 'openai') {
        if (policy === 'local-only' || policy === 'ollama-first') continue;
        const allowPaid = getStr(KEYS.ALLOW_PAID, 'false') === 'true';
        if (policy === 'ask-before-paid' && !allowPaid) {
          const ok = typeof confirm === 'function'
            ? confirm('Ollama Local và Gemini chưa sẵn sàng. Bạn có đồng ý gửi dữ liệu lên OpenAI API (có tính phí) không?')
            : false;
          if (!ok) throw new Error('Người dùng từ chối chuyển sang OpenAI trả phí.');
        }
      }

      try {
        if (prov === 'ollama') {
          return await callOllamaLocal(systemPrompt, userText, opts);
        }
        if (prov === 'gemini') {
          if (!hasGeminiConfigured()) continue; // KHÔNG báo lỗi chặn hệ thống khi chưa nhập Gemini Key!
          const r = await global.VBTM_AICore.generate(systemPrompt, userText, opts);
          recordSuccess('gemini', r.latencyMs);
          return Object.assign({ provider: 'gemini', connectionType: 'Cloud (Google Gemini)' }, r);
        }
        if (prov === 'openai') {
          return await callOpenAiDirect(systemPrompt, userText, opts);
        }
      } catch (err) {
        lastErr = err;
        if (prov === 'gemini') recordFailure('gemini', err.message);
        if ((err && err.name === 'AbortError') || policy === 'manual' || policy === 'local-only') throw err;
      }
    }

    throw lastErr || new Error('Ollama Local chưa sẵn sàng và chưa có AI Cloud bổ sung. Hãy mở Trung tâm AI để cài đặt Ollama (miễn phí, không cần API Key) hoặc nhập Key bổ sung.');
  }

  async function testActiveConnection(host) {
    return await route('Chỉ trả lời đúng 1 từ: OK', 'Kiểm tra kết nối VBTM AI Gateway.', {
      host: host || 'office',
      taskType: 'spell',
      expectJson: false,
      timeoutMs: 25000
    });
  }

  function getProviderStatuses() {
    return [
      {
        id: 'ollama',
        name: 'Ollama Local (Mặc định — Không cần API Key)',
        connectionType: 'Local',
        status: runtimeDiscovery.status,
        running: runtimeDiscovery.ollamaRunning,
        models: runtimeDiscovery.installedModels,
        activeModel: getOllamaModel(),
        endpoint: getOllamaUrl(),
        recommended: recommendModelForHardware(),
        lastLatencyMs: circuitState.ollama.lastLatencyMs,
        lastError: circuitState.ollama.lastError || runtimeDiscovery.errorDetail
      },
      {
        id: 'gemini',
        name: 'Google Gemini API (Tùy chọn bổ sung)',
        connectionType: 'Cloud',
        configured: hasGeminiConfigured(),
        status: hasGeminiConfigured() ? 'configured' : 'optional_not_set',
        lastLatencyMs: circuitState.gemini.lastLatencyMs,
        lastError: circuitState.gemini.lastError
      },
      {
        id: 'openai',
        name: 'OpenAI API (Tùy chọn trả phí)',
        connectionType: 'Cloud',
        configured: !!getOpenAiKey(),
        status: getOpenAiKey() ? 'configured' : 'optional_not_set',
        lastLatencyMs: circuitState.openai.lastLatencyMs,
        lastError: circuitState.openai.lastError
      }
    ];
  }

  global.VBTM_SmartRouter = {
    MODEL_CATALOG,
    route,
    testActiveConnection,
    resolveProviderOrder,
    checkOllamaStatus,
    pullOllamaModel,
    recommendModelForHardware,
    pickBestInstalledModelForTask,
    getProviderStatuses,
    getPolicy, setPolicy,
    getManualProvider, setManualProvider,
    getOpenAiKey, setOpenAiKey,
    getOpenAiModel, setOpenAiModel,
    getOllamaUrl, setOllamaUrl,
    getOllamaModel, setOllamaModel
  };
})(typeof window !== 'undefined' ? window : globalThis);
