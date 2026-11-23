/**
 * VBTM.SmartRouter — VBTM AI Gateway duy nhất (Mặc định: Ollama Local http://127.0.0.1:11434)
 *
 * Nguyên tắc kiến trúc bắt buộc:
 * 1. Provider mặc định là 'ollama' (Ollama Local tại http://127.0.0.1:11434) — KHÔNG cần Gemini API Key.
 * 2. Khi đang chọn 'ollama': TUYỆT ĐỐI KHÔNG gọi Gemini/OpenAI, KHÔNG chuyển ngầm sang Gemini khi Ollama lỗi.
 * 3. Chỉ sử dụng model thực sự tồn tại trên máy (lấy trực tiếp từ GET /api/tags), không tự đoán tên model.
 * 4. Circuit Breaker quản lý ĐỘC LẬP cho từng provider; lỗi Gemini cũ KHÔNG BAO GIỜ khóa Ollama.
 * 5. Đổi provider, sửa cấu hình hoặc bấm Kiểm tra kết nối sẽ tự động reset Circuit Breaker của provider đó.
 * 6. Không giả lập câu trả lời AI; luôn trả lỗi thật, rõ ràng và giữ nguyên nội dung người dùng.
 */
(function (global) {
  'use strict';

  const KEYS = {
    PROVIDER: 'vbtm_active_provider_v4',     // 'ollama' (mặc định) | 'gemini' | 'openai'
    OLLAMA_URL: 'vbtm_office_ollama_url',    // mặc định http://127.0.0.1:11434
    OLLAMA_MODEL: 'vbtm_office_ollama_model',
    OPENAI_KEY: 'vbtm_office_openai_key',
    OPENAI_MODEL: 'vbtm_office_openai_model',
    LEGACY_POLICY: 'vbtm_router_policy'
  };

  // Dọn chính sách cũ có thể kéo về Gemini từ localStorage cũ
  try {
    const oldPol = localStorage.getItem(KEYS.LEGACY_POLICY);
    if (oldPol && !localStorage.getItem(KEYS.PROVIDER)) {
      localStorage.setItem(KEYS.PROVIDER, 'ollama');
    }
    if (oldPol) localStorage.removeItem(KEYS.LEGACY_POLICY);
  } catch (e) {}

  // Circuit Breaker riêng biệt hoàn toàn cho từng provider
  const circuitState = {
    ollama: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    gemini: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    openai: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' }
  };

  const ollamaCache = {
    running: false,
    models: [],
    lastError: '',
    checkedAt: 0
  };

  const requestAuditLog = []; // Ghi nhận provider nào thực sự được gọi (để kiểm chứng 0 request sang Gemini khi chọn Ollama)

  function getStr(k, def) {
    try { return (localStorage.getItem(k) || def).trim(); } catch (e) { return def; }
  }
  function setStr(k, v) {
    try {
      if (v) localStorage.setItem(k, String(v).trim());
      else localStorage.removeItem(k);
    } catch (e) {}
  }

  function getActiveProvider() {
    const p = getStr(KEYS.PROVIDER, 'ollama').toLowerCase();
    return (p === 'gemini' || p === 'openai') ? p : 'ollama';
  }

  function setActiveProvider(prov) {
    const clean = (prov === 'gemini' || prov === 'openai') ? prov : 'ollama';
    setStr(KEYS.PROVIDER, clean);
    // Reset Circuit Breaker của provider mới chọn để không dùng lỗi cũ trong cache
    resetCircuitBreaker(clean);
  }

  function getOllamaUrl() {
    const u = getStr(KEYS.OLLAMA_URL, 'http://127.0.0.1:11434');
    return u.replace(/\/+$/, '');
  }
  function setOllamaUrl(u) {
    setStr(KEYS.OLLAMA_URL, u || 'http://127.0.0.1:11434');
    resetCircuitBreaker('ollama');
  }

  function getOllamaModel() { return getStr(KEYS.OLLAMA_MODEL, ''); }
  function setOllamaModel(m) {
    setStr(KEYS.OLLAMA_MODEL, m || '');
    resetCircuitBreaker('ollama');
  }

  function getOpenAiKey() { return getStr(KEYS.OPENAI_KEY, ''); }
  function setOpenAiKey(k) { setStr(KEYS.OPENAI_KEY, k); resetCircuitBreaker('openai'); }
  function getOpenAiModel() { return getStr(KEYS.OPENAI_MODEL, 'gpt-4o-mini'); }
  function setOpenAiModel(m) { setStr(KEYS.OPENAI_MODEL, m || 'gpt-4o-mini'); resetCircuitBreaker('openai'); }

  function isCircuitOpen(prov) {
    const c = circuitState[prov];
    return !!(c && c.openUntil > Date.now());
  }

  function resetCircuitBreaker(prov) {
    if (prov && circuitState[prov]) {
      circuitState[prov].fails = 0;
      circuitState[prov].openUntil = 0;
      circuitState[prov].lastError = '';
      return;
    }
    Object.keys(circuitState).forEach(k => {
      circuitState[k].fails = 0;
      circuitState[k].openUntil = 0;
      circuitState[k].lastError = '';
    });
  }

  function tripCircuitForTest(prov, ms) {
    if (circuitState[prov]) {
      circuitState[prov].fails = 3;
      circuitState[prov].openUntil = Date.now() + (ms || 60000);
      circuitState[prov].lastError = 'Circuit tripped for isolation test';
    }
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

  function logAudit(prov, model, status, ms, note) {
    requestAuditLog.unshift({
      ts: new Date().toISOString(),
      provider: prov,
      model: model || '-',
      status: status || 'ok',
      latencyMs: ms || 0,
      note: String(note || '').slice(0, 160)
    });
    if (requestAuditLog.length > 50) requestAuditLog.length = 50;
  }

  function getRequestAuditLog() { return requestAuditLog.slice(); }

  // API chính thức của Ollama: GET /api/tags
  async function checkOllamaStatus() {
    const base = getOllamaUrl();
    const t0 = Date.now();
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 4000);
      const res = await fetch(base + '/api/tags', { method: 'GET', signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) {
        ollamaCache.running = false;
        ollamaCache.models = [];
        ollamaCache.lastError = `Ollama phản hồi lỗi HTTP ${res.status} tại ${base}/api/tags.`;
        return { running: false, status: 'error', models: [], selectedModel: '', error: ollamaCache.lastError };
      }
      const data = await res.json();
      const models = (data && Array.isArray(data.models)) ? data.models.map(m => String(m.name || '')).filter(Boolean) : [];
      ollamaCache.running = true;
      ollamaCache.models = models;
      ollamaCache.checkedAt = Date.now();
      ollamaCache.lastError = '';

      // Chỉ chọn model thực sự tồn tại trên máy, không tự đoán tên model
      let activeModel = getOllamaModel();
      if (!activeModel || !models.includes(activeModel)) {
        activeModel = models.length > 0 ? models[0] : '';
        if (activeModel) setStr(KEYS.OLLAMA_MODEL, activeModel);
      }

      if (models.length > 0) {
        resetCircuitBreaker('ollama');
        recordSuccess('ollama', Date.now() - t0);
        return { running: true, status: 'ready', models, selectedModel: activeModel, latencyMs: Date.now() - t0 };
      }
      return {
        running: true,
        status: 'no_model',
        models: [],
        selectedModel: '',
        latencyMs: Date.now() - t0,
        error: 'Ollama đang chạy nhưng chưa cài model nào. Hãy tải model bằng lệnh: ollama pull qwen2.5:3b (hoặc qwen2.5:7b).'
      };
    } catch (err) {
      ollamaCache.running = false;
      ollamaCache.models = [];
      ollamaCache.lastError = `Không kết nối được tới Ollama Local tại ${base}. Kiểm tra đã cài đặt và khởi chạy Ollama chưa.`;
      return { running: false, status: 'not_running', models: [], selectedModel: '', error: ollamaCache.lastError };
    }
  }

  // API chính thức của Ollama: POST /api/generate (hoặc /api/chat)
  async function callOllamaLocal(systemPrompt, userText, opts) {
    // Luôn kiểm tra trạng thái thực tế trước khi gửi (không dựa vào lỗi cũ trong cache)
    const st = await checkOllamaStatus();
    if (!st.running) {
      recordFailure('ollama', st.error);
      logAudit('ollama', '-', 'error', 0, st.error);
      throw new Error(st.error);
    }
    if (!st.models.length || !st.selectedModel) {
      const msg = st.error || 'Ollama chưa có model nào được cài đặt. Hãy chạy "ollama pull qwen2.5:3b" trên máy.';
      logAudit('ollama', '-', 'no_model', 0, msg);
      throw new Error(msg);
    }

    const base = getOllamaUrl();
    const model = st.selectedModel;
    const t0 = Date.now();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeoutMs) || 60000);
    try {
      const res = await fetch(base + '/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: model,
          system: systemPrompt || 'Bạn là trợ lý văn phòng VBTM Pro. Trả lời trực tiếp bằng tiếng Việt chuẩn mực.',
          prompt: userText,
          stream: false
        }),
        signal: ctrl.signal
      });
      if (!res.ok) {
        const errBody = await res.text().catch(() => '');
        const msg = /memory|ram|oom/i.test(errBody)
          ? `Model "${model}" vượt quá RAM khả dụng trên máy. Hãy chọn model nhẹ hơn (ví dụ qwen2.5:3b).`
          : `Ollama lỗi HTTP ${res.status} khi chạy model "${model}": ${errBody.slice(0, 160)}`;
        recordFailure('ollama', msg);
        logAudit('ollama', model, 'error', Date.now() - t0, msg);
        throw new Error(msg);
      }
      const data = await res.json();
      const text = String((data && data.response) || '').trim();
      if (!text) {
        const msg = `Model "${model}" trả về nội dung rỗng.`;
        recordFailure('ollama', msg);
        logAudit('ollama', model, 'empty', Date.now() - t0, msg);
        throw new Error(msg);
      }
      const ms = Date.now() - t0;
      recordSuccess('ollama', ms);
      logAudit('ollama', model, 'ok', ms, 'len=' + text.length);
      return { provider: 'ollama', connectionType: 'Local (127.0.0.1:11434)', model, text, latencyMs: ms };
    } catch (err) {
      if (err && err.name === 'AbortError') {
        const msg = `Quá thời gian chờ phản hồi từ model "${model}" trên Ollama Local.`;
        recordFailure('ollama', msg);
        logAudit('ollama', model, 'timeout', Date.now() - t0, msg);
        throw new Error(msg);
      }
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  async function callOpenAiDirect(systemPrompt, userText, opts) {
    if (isCircuitOpen('openai')) {
      throw new Error('OpenAI API đang tạm ngắt mạch 60s do lỗi liên tiếp. Bấm "Kiểm tra kết nối" để đặt lại.');
    }
    const key = getOpenAiKey();
    if (!key) throw new Error('Chưa cấu hình OpenAI API Key (tùy chọn trả phí). Hãy chuyển sang Ollama Local hoặc nhập Key OpenAI.');
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
      logAudit('openai', model, 'ok', ms, 'len=' + out.length);
      return { provider: 'openai', connectionType: 'Cloud (OpenAI)', model, text: out, latencyMs: ms };
    } catch (err) {
      recordFailure('openai', err.message);
      logAudit('openai', model, 'error', Date.now() - t0, err.message);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  // Điểm gọi AI DUY NHẤT cho toàn bộ hệ thống (KHÔNG chuyển ngầm sang Gemini khi chọn Ollama!)
  async function route(systemPrompt, userText, opts) {
    opts = opts || {};
    const prov = (opts.forceProvider || getActiveProvider()).toLowerCase();

    if (prov === 'ollama') {
      // Lỗi Circuit Breaker của Gemini tuyệt đối không ảnh hưởng tới Ollama
      return await callOllamaLocal(systemPrompt, userText, opts);
    }

    if (prov === 'gemini') {
      if (isCircuitOpen('gemini')) {
        throw new Error('Google Gemini đang tạm ngắt mạch 60s do lỗi liên tiếp. Bạn có thể chuyển sang Ollama Local (không cần Key) hoặc bấm "Kiểm tra kết nối" để đặt lại.');
      }
      if (!global.VBTM_AICore) throw new Error('VBTM_AICore chưa được nạp.');
      if (!global.VBTM_AICore.getApiKey() && global.VBTM_AICore.getMode() !== 'gateway') {
        throw new Error('Bạn đang chọn nhà cung cấp bổ sung Google Gemini nhưng chưa nhập Gemini API Key. Hãy nhập Key hoặc chuyển về mặc định Ollama Local.');
      }
      try {
        const r = await global.VBTM_AICore.generate(systemPrompt, userText, opts);
        recordSuccess('gemini', r.latencyMs);
        logAudit('gemini', r.model, 'ok', r.latencyMs, 'len=' + r.text.length);
        return Object.assign({ provider: 'gemini', connectionType: 'Cloud (Google Gemini)' }, r);
      } catch (err) {
        recordFailure('gemini', err.message);
        logAudit('gemini', '-', 'error', 0, err.message);
        throw err;
      }
    }

    if (prov === 'openai') {
      return await callOpenAiDirect(systemPrompt, userText, opts);
    }

    return await callOllamaLocal(systemPrompt, userText, opts);
  }

  async function testActiveConnection(host) {
    const prov = getActiveProvider();
    // Reset Circuit Breaker của provider đang chọn trước khi kiểm tra thủ công
    resetCircuitBreaker(prov);
    return await route('Chỉ trả lời đúng 1 từ: OK', 'Xin chào, hãy trả lời OK để xác nhận kết nối.', {
      host: host || 'office',
      expectJson: false,
      timeoutMs: 30000
    });
  }

  function getProviderStatuses() {
    const geminiKeySet = !!(global.VBTM_AICore && global.VBTM_AICore.getApiKey());
    return {
      activeProvider: getActiveProvider(),
      ollama: {
        running: ollamaCache.running,
        endpoint: getOllamaUrl(),
        models: ollamaCache.models,
        selectedModel: getOllamaModel(),
        circuitOpen: isCircuitOpen('ollama'),
        lastLatencyMs: circuitState.ollama.lastLatencyMs,
        lastError: circuitState.ollama.lastError || ollamaCache.lastError
      },
      gemini: {
        configured: geminiKeySet,
        circuitOpen: isCircuitOpen('gemini'),
        lastLatencyMs: circuitState.gemini.lastLatencyMs,
        lastError: circuitState.gemini.lastError
      },
      openai: {
        configured: !!getOpenAiKey(),
        circuitOpen: isCircuitOpen('openai'),
        lastLatencyMs: circuitState.openai.lastLatencyMs,
        lastError: circuitState.openai.lastError
      }
    };
  }

  // Giữ tương thích ngược cho các hàm getPolicy/setPolicy nếu module cũ gọi
  function getPolicy() { return getActiveProvider(); }
  function setPolicy(p) { setActiveProvider(p === 'gemini-first' ? 'gemini' : (p === 'openai' ? 'openai' : 'ollama')); }

  global.VBTM_SmartRouter = {
    route,
    testActiveConnection,
    checkOllamaStatus,
    getActiveProvider, setActiveProvider,
    getOllamaUrl, setOllamaUrl,
    getOllamaModel, setOllamaModel,
    getOpenAiKey, setOpenAiKey,
    getOpenAiModel, setOpenAiModel,
    resetCircuitBreaker,
    tripCircuitForTest,
    isCircuitOpen,
    getProviderStatuses,
    getRequestAuditLog,
    getPolicy, setPolicy
  };
})(typeof window !== 'undefined' ? window : globalThis);
