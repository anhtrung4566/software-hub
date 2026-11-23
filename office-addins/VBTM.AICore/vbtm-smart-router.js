/**
 * VBTM.SmartRouter — Trung tâm điều phối AI đa nhà cung cấp (VBTM AI Provider Center)
 * Hỗ trợ API chính thức:
 *  1. Google Gemini API (ưu tiên miễn phí/nhanh qua VBTM_AICore)
 *  2. OpenAI API chính thức (trả phí theo token — KHÔNG bao giờ tự chuyển sang trả phí nếu chưa được phép!)
 *  3. Ollama Local (http://127.0.0.1:11434 — chạy cục bộ hoàn toàn ngoại tuyến, không tự tải model lớn)
 *
 * Cơ chế an toàn:
 *  - Circuit Breaker (tự ngắt 60s nếu 1 nhà cung cấp lỗi liên tiếp 3 lần)
 *  - 5 chế độ chính sách: 'free-only' | 'free-first' | 'ask-before-paid' | 'local-only' | 'manual'
 *  - Không thu thập cookie/mật khẩu, không né hạn mức, minh bạch dữ liệu gửi ra ngoài.
 */
(function (global) {
  'use strict';

  const KEYS = {
    POLICY: 'vbtm_router_policy',          // 'free-only' | 'free-first' | 'ask-before-paid' | 'local-only' | 'manual'
    MANUAL_PROV: 'vbtm_router_manual_prov', // 'gemini' | 'openai' | 'ollama'
    OPENAI_KEY: 'vbtm_office_openai_key',
    OPENAI_MODEL: 'vbtm_office_openai_model',
    OLLAMA_URL: 'vbtm_office_ollama_url',
    OLLAMA_MODEL: 'vbtm_office_ollama_model',
    ALLOW_PAID: 'vbtm_router_allow_paid'
  };

  const circuitState = {
    gemini: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    openai: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' },
    ollama: { fails: 0, openUntil: 0, lastLatencyMs: 0, lastError: '' }
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

  function getPolicy() { return getStr(KEYS.POLICY, 'free-first'); }
  function setPolicy(p) { setStr(KEYS.POLICY, p); }
  function getManualProvider() { return getStr(KEYS.MANUAL_PROV, 'gemini'); }
  function setManualProvider(p) { setStr(KEYS.MANUAL_PROV, p); }
  function getOpenAiKey() { return getStr(KEYS.OPENAI_KEY, ''); }
  function setOpenAiKey(k) { setStr(KEYS.OPENAI_KEY, k); }
  function getOpenAiModel() { return getStr(KEYS.OPENAI_MODEL, 'gpt-4o-mini'); }
  function setOpenAiModel(m) { setStr(KEYS.OPENAI_MODEL, m); }
  function getOllamaUrl() { return getStr(KEYS.OLLAMA_URL, 'http://127.0.0.1:11434'); }
  function setOllamaUrl(u) { setStr(KEYS.OLLAMA_URL, u); }
  function getOllamaModel() { return getStr(KEYS.OLLAMA_MODEL, 'qwen2.5:7b'); }
  function setOllamaModel(m) { setStr(KEYS.OLLAMA_MODEL, m); }

  function isCircuitOpen(prov) {
    const c = circuitState[prov];
    return !!(c && c.openUntil > Date.now());
  }

  function recordSuccess(prov, ms) {
    const c = circuitState[prov];
    if (!c) return;
    c.fails = 0;
    c.openUntil = 0;
    c.lastLatencyMs = ms || 0;
    c.lastError = '';
  }

  function recordFailure(prov, errMsg) {
    const c = circuitState[prov];
    if (!c) return;
    c.fails += 1;
    c.lastError = String(errMsg || '').slice(0, 160);
    if (c.fails >= 3) {
      c.openUntil = Date.now() + 60000; // Ngắt mạch 60 giây tránh lặp vô hạn
    }
  }

  async function checkOllamaStatus() {
    const base = getOllamaUrl().replace(/\/+$/, '');
    const t0 = Date.now();
    try {
      const ctrl = new AbortController();
      const timer = setTimeout(() => ctrl.abort(), 3500);
      const res = await fetch(base + '/api/tags', { method: 'GET', signal: ctrl.signal });
      clearTimeout(timer);
      if (!res.ok) return { running: false, models: [], error: 'HTTP ' + res.status };
      const data = await res.json();
      const models = (data && Array.isArray(data.models)) ? data.models.map(m => m.name) : [];
      recordSuccess('ollama', Date.now() - t0);
      return { running: true, models, latencyMs: Date.now() - t0 };
    } catch (e) {
      return { running: false, models: [], error: 'Ollama chưa chạy hoặc chưa cài đặt tại ' + base };
    }
  }

  async function callOpenAiDirect(systemPrompt, userText, opts) {
    const key = getOpenAiKey();
    if (!key) throw new Error('Chưa nhập OpenAI API Key chính thức trong Trung tâm AI.');
    const model = getOpenAiModel();
    const t0 = Date.now();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeoutMs) || 28000);
    try {
      const res = await fetch('https://api.openai.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + key
        },
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
      return { provider: 'openai', model, text: out, latencyMs: ms };
    } catch (err) {
      recordFailure('openai', err.message);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  async function callOllamaLocal(systemPrompt, userText, opts) {
    const base = getOllamaUrl().replace(/\/+$/, '');
    const model = getOllamaModel();
    const t0 = Date.now();
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), (opts && opts.timeoutMs) || 45000);
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
      if (!res.ok) throw new Error(`[Ollama HTTP ${res.status}] Model "${model}" chưa sẵn sàng trên Ollama cục bộ.`);
      const data = await res.json();
      const out = String((data && data.response) || '').trim();
      if (!out) throw new Error('Ollama trả về nội dung rỗng.');
      const ms = Date.now() - t0;
      recordSuccess('ollama', ms);
      return { provider: 'ollama', model, text: out, latencyMs: ms };
    } catch (err) {
      recordFailure('ollama', err.message);
      throw err;
    } finally {
      clearTimeout(timer);
    }
  }

  function resolveProviderOrder(opts) {
    const policy = (opts && opts.policy) || getPolicy();
    const isSensitive = !!(opts && opts.sensitiveDocument);
    if (policy === 'local-only' || isSensitive) return ['ollama'];
    if (policy === 'manual') return [getManualProvider()];
    if (policy === 'free-only') return ['gemini', 'ollama'];
    if (policy === 'ask-before-paid') return ['gemini', 'ollama', 'openai'];
    // 'free-first' mặc định: Gemini -> Ollama (KHÔNG tự chuyển sang OpenAI trả phí!)
    return ['gemini', 'ollama'];
  }

  async function route(systemPrompt, userText, opts) {
    opts = opts || {};
    const policy = opts.policy || getPolicy();
    const order = resolveProviderOrder(opts);
    let lastErr = null;

    for (const prov of order) {
      if (isCircuitOpen(prov)) {
        lastErr = new Error(`Nhà cung cấp "${prov}" đang tạm ngắt mạch (Circuit Breaker) do lỗi liên tiếp.`);
        continue;
      }

      // Chính sách chống tự phát sinh chi phí: không tự gọi OpenAI trả phí nếu chưa xác nhận
      if (prov === 'openai') {
        const allowPaid = getStr(KEYS.ALLOW_PAID, 'false') === 'true';
        if (policy === 'free-only' || policy === 'free-first') continue;
        if (policy === 'ask-before-paid' && !allowPaid) {
          const ok = typeof confirm === 'function'
            ? confirm('AI miễn phí/cục bộ hiện không khả dụng. Bạn có đồng ý chuyển sang OpenAI API (có tính phí theo token) cho yêu cầu này không?')
            : false;
          if (!ok) throw new Error('Người dùng từ chối chuyển sang AI trả phí (OpenAI).');
        }
      }

      try {
        if (prov === 'gemini') {
          if (!global.VBTM_AICore) throw new Error('VBTM_AICore chưa được nạp.');
          const r = await global.VBTM_AICore.generate(systemPrompt, userText, opts);
          recordSuccess('gemini', r.latencyMs);
          return Object.assign({ provider: 'gemini' }, r);
        }
        if (prov === 'ollama') {
          return await callOllamaLocal(systemPrompt, userText, opts);
        }
        if (prov === 'openai') {
          return await callOpenAiDirect(systemPrompt, userText, opts);
        }
      } catch (err) {
        lastErr = err;
        if (prov === 'gemini') recordFailure('gemini', err.message);
        // Nếu người dùng hủy (AbortError) hoặc chọn thủ công -> dừng ngay
        if ((err && err.name === 'AbortError') || policy === 'manual') throw err;
      }
    }
    throw lastErr || new Error('Không có nhà cung cấp AI khả dụng theo chính sách hiện tại (' + policy + ').');
  }

  function getProviderStatuses() {
    const geminiReady = !!(global.VBTM_AICore && (global.VBTM_AICore.getMode() === 'gateway' || global.VBTM_AICore.getApiKey()));
    return [
      {
        id: 'gemini', name: 'Google Gemini API', paid: false,
        configured: geminiReady,
        circuitOpen: isCircuitOpen('gemini'),
        lastLatencyMs: circuitState.gemini.lastLatencyMs,
        lastError: circuitState.gemini.lastError
      },
      {
        id: 'openai', name: 'OpenAI API (Trả phí)', paid: true,
        configured: !!getOpenAiKey(),
        circuitOpen: isCircuitOpen('openai'),
        lastLatencyMs: circuitState.openai.lastLatencyMs,
        lastError: circuitState.openai.lastError
      },
      {
        id: 'ollama', name: 'Ollama Local (Ngoại tuyến)', paid: false,
        configured: true,
        endpoint: getOllamaUrl(),
        model: getOllamaModel(),
        circuitOpen: isCircuitOpen('ollama'),
        lastLatencyMs: circuitState.ollama.lastLatencyMs,
        lastError: circuitState.ollama.lastError
      }
    ];
  }

  global.VBTM_SmartRouter = {
    route,
    resolveProviderOrder,
    checkOllamaStatus,
    getProviderStatuses,
    getPolicy, setPolicy,
    getManualProvider, setManualProvider,
    getOpenAiKey, setOpenAiKey,
    getOpenAiModel, setOpenAiModel,
    getOllamaUrl, setOllamaUrl,
    getOllamaModel, setOllamaModel
  };
})(typeof window !== 'undefined' ? window : globalThis);
