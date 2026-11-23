/**
 * VBTM Pro 3.0 Ultimate — Central Google Gemini AI Service
 * Module trung tâm quản lý toàn bộ kết nối Google Gemini API:
 * - Kiểm tra API Key, kiểm tra kết nối, tự động dò model khả dụng
 * - Timeout, Retry tự động cho lỗi 429 / 500 / 503, Hủy yêu cầu (AbortController)
 * - Phân loại & xử lý chi tiết 12 nhóm lỗi
 * - Hỗ trợ Structured JSON Output
 * - Mask API Key & ghi log kỹ thuật KHÔNG lưu nội dung văn bản nhạy cảm
 */
(function (global) {
  'use strict';

  const STORAGE_KEYS = {
    API_KEY: 'vbtm30_gemini_api_key',
    MODEL: 'vbtm30_gemini_model',
    MODE: 'vbtm30_ai_mode', // 'direct' | 'proxy'
    PROXY_URL: 'vbtm30_ai_proxy_url'
  };

  const DEFAULT_MODELS = [
    'gemini-flash-lite-latest',
    'gemini-flash-latest',
    'gemini-3.5-flash-lite',
    'gemini-3.5-flash',
    'gemini-3.8-flash',
    'gemini-2.5-flash'
  ];

  const state = {
    status: 'idle', // 'idle' | 'connecting' | 'connected' | 'error'
    activeModel: localStorage.getItem(STORAGE_KEYS.MODEL) || 'gemini-2.5-flash',
    lastLatencyMs: 0,
    lastError: null,
    availableModels: [],
    abortController: null,
    logs: []
  };

  const listeners = new Set();

  function emitState() {
    const snap = getState();
    listeners.forEach(fn => {
      try { fn(snap); } catch (e) {}
    });
  }

  function maskKey(key) {
    const k = String(key || '').trim();
    if (!k) return '(trống)';
    if (k.length <= 10) return '****';
    return k.slice(0, 6) + '••••••' + k.slice(-4);
  }

  function pushLog(entry) {
    const rec = {
      time: new Date().toLocaleTimeString('vi-VN'),
      action: entry.action || 'request',
      model: entry.model || state.activeModel,
      status: entry.status || 'info',
      latencyMs: entry.latencyMs || 0,
      code: entry.code || '',
      message: entry.message || ''
    };
    state.logs.unshift(rec);
    if (state.logs.length > 40) state.logs.pop();
    emitState();
  }

  function getApiKey() {
    const domInput = typeof document !== 'undefined' ? document.getElementById('vbtmGeminiKey') : null;
    if (domInput && domInput.value.trim()) return domInput.value.trim();
    return (localStorage.getItem(STORAGE_KEYS.API_KEY) || '').trim();
  }

  function setApiKey(key) {
    const k = String(key || '').trim();
    if (k) localStorage.setItem(STORAGE_KEYS.API_KEY, k);
    else localStorage.removeItem(STORAGE_KEYS.API_KEY);
    const domInput = typeof document !== 'undefined' ? document.getElementById('vbtmGeminiKey') : null;
    if (domInput && domInput.value !== k) domInput.value = k;
    emitState();
  }

  function getModel() {
    return (localStorage.getItem(STORAGE_KEYS.MODEL) || state.activeModel || 'gemini-2.5-flash').trim();
  }

  function setModel(model) {
    const m = String(model || '').trim() || 'gemini-2.5-flash';
    state.activeModel = m;
    localStorage.setItem(STORAGE_KEYS.MODEL, m);
    emitState();
  }

  function getConnectionMode() {
    return localStorage.getItem(STORAGE_KEYS.MODE) || 'direct';
  }

  function setConnectionMode(mode, proxyUrl) {
    localStorage.setItem(STORAGE_KEYS.MODE, mode === 'proxy' ? 'proxy' : 'direct');
    if (typeof proxyUrl === 'string') {
      localStorage.setItem(STORAGE_KEYS.PROXY_URL, proxyUrl.trim());
    }
    emitState();
  }

  function getProxyUrl() {
    return localStorage.getItem(STORAGE_KEYS.PROXY_URL) || '/api/ai/gemini';
  }

  function classifyError(httpStatus, rawMessage, errObj) {
    const msg = String(rawMessage || (errObj && errObj.message) || '');
    if (errObj && errObj.name === 'AbortError') {
      return { code: 'ABORTED', retryable: false, userMessage: 'Đã hủy yêu cầu AI theo lệnh người dùng.' };
    }
    if (errObj && errObj.code === 'TIMEOUT') {
      return { code: 'TIMEOUT', retryable: true, userMessage: 'Hết thời gian chờ phản hồi từ Gemini API (Timeout). Vui lòng thử lại.' };
    }
    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      return { code: 'OFFLINE', retryable: false, userMessage: 'Máy tính đang mất kết nối Internet. Các tính năng chuẩn hóa & rà soát quy tắc ngoại tuyến vẫn dùng bình thường.' };
    }
    if (/Failed to fetch|NetworkError|ERR_INTERNET_DISCONNECTED|ENOTFOUND/i.test(msg)) {
      return { code: 'NETWORK_ERROR', retryable: true, userMessage: 'Không thể kết nối tới máy chủ Google Gemini (lỗi mạng hoặc bị tường lửa chặn).' };
    }
    if (httpStatus === 401 || httpStatus === 403 || /API_KEY_INVALID|API key not valid|PERMISSION_DENIED/i.test(msg)) {
      return { code: 'INVALID_API_KEY', retryable: false, userMessage: 'API Key Google Gemini không hợp lệ hoặc không có quyền truy cập. Kiểm tra lại khóa (bắt đầu bằng AIza...).' };
    }
    if (httpStatus === 404 || /NOT_FOUND|no longer available|is not found/i.test(msg)) {
      return { code: 'MODEL_NOT_FOUND', retryable: false, userMessage: 'Model AI không tồn tại hoặc không còn được hỗ trợ. Hệ thống sẽ tự dò model khả dụng.' };
    }
    if (httpStatus === 429 || /RESOURCE_EXHAUSTED|quota|rate limit/i.test(msg)) {
      return { code: 'RATE_LIMIT_429', retryable: true, userMessage: 'Gemini API đang quá tải hoặc vượt hạn mức lượt gọi (429). Đợi vài giây rồi thử lại.' };
    }
    if (httpStatus === 500 || httpStatus === 503 || /INTERNAL|UNAVAILABLE|overloaded/i.test(msg)) {
      return { code: 'SERVER_ERROR_5XX', retryable: true, userMessage: `Máy chủ Google Gemini tạm thời bận (${httpStatus || '5xx'}). Đang thử lại...` };
    }
    if (httpStatus === 400 || /INVALID_ARGUMENT/i.test(msg)) {
      return { code: 'INVALID_REQUEST', retryable: false, userMessage: 'Cấu trúc yêu cầu API không hợp lệ: ' + msg.slice(0, 180) };
    }
    return { code: 'API_ERROR', retryable: false, userMessage: msg.slice(0, 220) || 'Lỗi không xác định khi gọi Gemini API.' };
  }

  async function fetchWithTimeout(url, options, timeoutMs, externalSignal) {
    const controller = new AbortController();
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, timeoutMs || 25000);

    const onExternalAbort = () => controller.abort();
    if (externalSignal) {
      if (externalSignal.aborted) {
        clearTimeout(timer);
        const ab = new Error('Aborted');
        ab.name = 'AbortError';
        throw ab;
      }
      externalSignal.addEventListener('abort', onExternalAbort, { once: true });
    }

    try {
      const resp = await fetch(url, Object.assign({}, options, { signal: controller.signal }));
      return resp;
    } catch (e) {
      if (timedOut) {
        const te = new Error('Request timed out');
        te.code = 'TIMEOUT';
        throw te;
      }
      throw e;
    } finally {
      clearTimeout(timer);
      if (externalSignal) externalSignal.removeEventListener('abort', onExternalAbort);
    }
  }

  async function listAvailableModels(apiKey) {
    const key = String(apiKey || getApiKey()).trim();
    if (!key) throw { code: 'INVALID_API_KEY', userMessage: 'Chưa nhập Gemini API Key.' };

    const t0 = performance.now();
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${encodeURIComponent(key)}`;
    const resp = await fetchWithTimeout(url, { method: 'GET' }, 15000, null);
    let data = null;
    try { data = await resp.json(); } catch (e) {}
    if (!resp.ok) {
      const errMsg = data && data.error && data.error.message ? data.error.message : `HTTP ${resp.status}`;
      throw classifyError(resp.status, errMsg);
    }
    const models = (data && Array.isArray(data.models) ? data.models : [])
      .filter(m => Array.isArray(m.supportedGenerationMethods) && m.supportedGenerationMethods.includes('generateContent'))
      .map(m => String(m.name || '').replace(/^models\//, ''))
      .filter(Boolean);

    state.availableModels = models;
    const latency = Math.round(performance.now() - t0);
    pushLog({ action: 'listModels', status: 'ok', latencyMs: latency, message: `Tìm thấy ${models.length} model khả dụng` });
    return models;
  }

  async function pickBestFlashModel(apiKey) {
    const models = await listAvailableModels(apiKey);
    const preferred = [
      getModel(),
      'gemini-2.5-flash',
      'gemini-2.0-flash',
      'gemini-3.8-flash',
      'gemini-3.0-flash',
      'gemini-1.5-flash'
    ];
    for (const p of preferred) {
      if (models.includes(p)) return p;
    }
    const anyFlash = models.find(m => /flash/i.test(m) && !/embedding|tts|image|audio/i.test(m));
    return anyFlash || models[0] || 'gemini-2.5-flash';
  }

  async function testApiKey(customKey) {
    const key = String(customKey || getApiKey()).trim();
    if (!key) {
      state.status = 'error';
      state.lastError = 'Chưa nhập API Key Google Gemini.';
      emitState();
      return { ok: false, code: 'INVALID_API_KEY', message: state.lastError };
    }
    if (!key.startsWith('AIza') || key.length < 30) {
      state.status = 'error';
      state.lastError = `Định dạng API Key (${maskKey(key)}) không giống khóa Google Gemini (phải bắt đầu bằng AIza..., dài ~39 ký tự).`;
      emitState();
      return { ok: false, code: 'INVALID_API_KEY', message: state.lastError };
    }
    state.status = 'connecting';
    state.lastError = null;
    emitState();
    const t0 = performance.now();
    try {
      const bestModel = await pickBestFlashModel(key);
      setModel(bestModel);
      state.status = 'connected';
      state.lastLatencyMs = Math.round(performance.now() - t0);
      emitState();
      pushLog({ action: 'testApiKey', model: bestModel, status: 'ok', latencyMs: state.lastLatencyMs, message: `Key hợp lệ (${maskKey(key)})` });
      return { ok: true, model: bestModel, models: state.availableModels, latencyMs: state.lastLatencyMs, message: `API Key hợp lệ! Đã chọn model: ${bestModel}` };
    } catch (err) {
      const info = err && err.userMessage ? err : classifyError(0, err && err.message, err);
      state.status = 'error';
      state.lastError = info.userMessage;
      state.lastLatencyMs = Math.round(performance.now() - t0);
      pushLog({ action: 'testApiKey', status: 'error', code: info.code, latencyMs: state.lastLatencyMs, message: info.userMessage });
      return { ok: false, code: info.code, message: info.userMessage };
    }
  }

  async function callOllamaLocalFallback(promptText, systemInstruction, timeoutMs) {
    const base = (localStorage.getItem('vbtm_office_ollama_url') || 'http://localhost:11434').replace(/\/+$/, '');
    const t0 = performance.now();
    let models = [];
    try {
      const tr = await fetchWithTimeout(base + '/api/tags', { method: 'GET' }, 3500, null);
      if (tr.ok) {
        const td = await tr.json();
        models = (td && Array.isArray(td.models)) ? td.models.map(m => m.name) : [];
      }
    } catch (e) {
      throw {
        code: 'OLLAMA_NOT_RUNNING',
        retryable: false,
        userMessage: 'Chưa phát hiện Ollama Local tại http://localhost:11434. Hãy cài đặt/khởi chạy Ollama (miễn phí, không cần API Key) hoặc nhập Gemini API Key bổ sung.'
      };
    }
    if (!models.length) {
      throw {
        code: 'OLLAMA_NO_MODEL',
        retryable: false,
        userMessage: 'Ollama đã chạy nhưng chưa có model tiếng Việt. Hãy chạy lệnh "ollama pull qwen2.5:7b" (hoặc qwen2.5:3b) để dùng AI miễn phí không cần API Key.'
      };
    }
    const pref = ['qwen2.5:7b', 'qwen2.5:3b', 'gemma2:9b', 'qwen2.5:14b'];
    let chosen = models[0];
    for (const p of pref) {
      const hit = models.find(m => m.toLowerCase().startsWith(p));
      if (hit) { chosen = hit; break; }
    }
    const gr = await fetchWithTimeout(base + '/api/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: chosen,
        prompt: (systemInstruction ? systemInstruction + '\n\n' : '') + promptText,
        stream: false
      })
    }, timeoutMs || 60000, null);
    if (!gr.ok) {
      throw { code: 'OLLAMA_ERROR', retryable: false, userMessage: `Ollama Local lỗi HTTP ${gr.status} với model ${chosen}.` };
    }
    const gd = await gr.json();
    const text = String((gd && gd.response) || '').trim();
    const latencyMs = Math.round(performance.now() - t0);
    state.status = 'connected';
    state.activeModel = chosen + ' (Ollama Local)';
    state.lastLatencyMs = latencyMs;
    state.lastError = null;
    emitState();
    pushLog({ action: 'ollama', model: chosen, status: 'ok', latencyMs, message: 'Local AI OK (không cần API Key)' });
    return { text, model: chosen, provider: 'ollama', latencyMs };
  }

  async function testConnection() {
    const key = getApiKey();
    if (!key) {
      try {
        const res = await callOllamaLocalFallback('Trả lời đúng 2 chữ: SAN_SANG', 'Chỉ trả lời đúng 2 chữ SAN_SANG.', 20000);
        return { ok: true, model: res.model + ' (Ollama Local)', latencyMs: res.latencyMs, message: `Kết nối Ollama Local thành công (${res.model} · ${res.latencyMs}ms · không cần API Key)!` };
      } catch (err) {
        state.status = 'idle';
        state.lastError = err.userMessage || err.message;
        emitState();
        return { ok: false, code: err.code || 'OLLAMA_NOT_RUNNING', message: state.lastError };
      }
    }
    const keyCheck = await testApiKey();
    if (!keyCheck.ok) return keyCheck;
    try {
      const res = await request({
        prompt: 'Trả lời đúng 2 chữ: SAN_SANG',
        systemInstruction: 'Chỉ trả lời đúng 2 chữ SAN_SANG.',
        temperature: 0,
        maxOutputTokens: 32,
        timeoutMs: 15000,
        maxRetries: 1
      });
      return {
        ok: true,
        model: res.model,
        latencyMs: res.latencyMs,
        message: `Kết nối thành công tới ${res.model} (${res.latencyMs}ms)!`
      };
    } catch (err) {
      return { ok: false, code: err.code || 'API_ERROR', message: err.userMessage || err.message || 'Kiểm tra kết nối thất bại.' };
    }
  }

  function cancelActiveRequest() {
    if (state.abortController) {
      state.abortController.abort();
      state.abortController = null;
      state.status = 'idle';
      pushLog({ action: 'cancel', status: 'warn', code: 'ABORTED', message: 'Người dùng đã hủy yêu cầu AI.' });
    }
  }

  async function invokeSingleAttempt(model, payload, apiKey, timeoutMs, signal) {
    const mode = getConnectionMode();
    if (mode === 'proxy') {
      const resp = await fetchWithTimeout(getProxyUrl(), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ model, payload })
      }, timeoutMs, signal);
      let data = null;
      try { data = await resp.json(); } catch (e) {}
      if (!resp.ok) {
        const msg = (data && data.error && data.error.message) || (data && data.message) || `HTTP ${resp.status}`;
        throw classifyError(resp.status, msg);
      }
      return data;
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`;
    const resp = await fetchWithTimeout(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }, timeoutMs, signal);

    let data = null;
    try { data = await resp.json(); } catch (e) {
      if (!resp.ok) throw classifyError(resp.status, `HTTP ${resp.status} (Invalid JSON)`);
      throw { code: 'INVALID_JSON', retryable: false, userMessage: 'Dữ liệu JSON trả về từ Gemini API không hợp lệ.' };
    }

    if (!resp.ok) {
      const msg = (data && data.error && data.error.message) || `HTTP ${resp.status}`;
      throw classifyError(resp.status, msg);
    }
    return data;
  }

  async function request(opts) {
    const options = opts || {};
    const mode = getConnectionMode();
    const apiKey = String(options.apiKey || getApiKey()).trim();

    if (mode === 'direct' && !apiKey) {
      // Không chặn thiếu Gemini API Key — tự động chạy qua Ollama Local (http://localhost:11434)
      state.status = 'connecting';
      state.lastError = null;
      emitState();
      try {
        return await callOllamaLocalFallback(options.prompt || '', options.systemInstruction || '', options.timeoutMs || 60000);
      } catch (oErr) {
        state.status = 'error';
        state.lastError = oErr.userMessage || oErr.message;
        emitState();
        throw oErr;
      }
    }

    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      const err = classifyError(0, 'OFFLINE');
      state.status = 'error';
      state.lastError = err.userMessage;
      emitState();
      throw err;
    }

    cancelActiveRequest();
    const controller = new AbortController();
    state.abortController = controller;
    state.status = 'connecting';
    state.lastError = null;
    emitState();

    const timeoutMs = options.timeoutMs || 28000;
    const maxRetries = typeof options.maxRetries === 'number' ? options.maxRetries : 2;
    let model = String(options.model || getModel()).trim() || 'gemini-2.5-flash';

    const generationConfig = {
      temperature: typeof options.temperature === 'number' ? options.temperature : 0.3,
      maxOutputTokens: options.maxOutputTokens || 2048
    };
    if (options.expectJson) {
      generationConfig.responseMimeType = 'application/json';
    }

    const fullText = options.systemInstruction
      ? `${options.systemInstruction}\n\n${options.prompt || ''}`
      : String(options.prompt || '');

    const payload = {
      contents: [{ parts: [{ text: fullText }] }],
      generationConfig
    };

    const t0 = performance.now();
    let lastErr = null;
    let modelSwitched = false;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const data = await invokeSingleAttempt(model, payload, apiKey, timeoutMs, controller.signal);
        const cand = data && Array.isArray(data.candidates) ? data.candidates[0] : null;
        const finishReason = cand && cand.finishReason ? cand.finishReason : '';
        const text = cand && cand.content && Array.isArray(cand.content.parts)
          ? cand.content.parts.map(p => p.text || '').join('').trim()
          : '';

        if (!text) {
          throw { code: 'EMPTY_RESPONSE', retryable: false, userMessage: 'Gemini AI trả về kết quả rỗng (có thể do bộ lọc an toàn hoặc nội dung đầu vào).' };
        }

        const latencyMs = Math.round(performance.now() - t0);
        state.status = 'connected';
        state.activeModel = model;
        state.lastLatencyMs = latencyMs;
        state.abortController = null;
        pushLog({
          action: options.actionName || 'generateContent',
          model,
          status: finishReason === 'MAX_TOKENS' ? 'warn' : 'ok',
          latencyMs,
          code: finishReason === 'MAX_TOKENS' ? 'TRUNCATED' : '200',
          message: finishReason === 'MAX_TOKENS' ? 'Kết quả chạm giới hạn độ dài (MAX_TOKENS)' : 'Thành công'
        });

        let parsedJson = null;
        if (options.expectJson) {
          try {
            const cleanJson = text.replace(/^```json\s*/i, '').replace(/```$/i, '').trim();
            parsedJson = JSON.parse(cleanJson);
          } catch (je) {
            const m = text.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
            if (m) {
              try { parsedJson = JSON.parse(m[1]); } catch (je2) {}
            }
            if (!parsedJson) {
              throw { code: 'INVALID_JSON', retryable: false, userMessage: 'AI trả về định dạng JSON không hợp lệ. Vui lòng thử lại.' };
            }
          }
        }

        return {
          text,
          json: parsedJson,
          model,
          latencyMs,
          truncated: finishReason === 'MAX_TOKENS'
        };
      } catch (err) {
        const info = err && err.code && err.userMessage ? err : classifyError(0, err && err.message, err);
        lastErr = info;

        if (info.code === 'ABORTED') {
          state.status = 'idle';
          state.abortController = null;
          emitState();
          throw info;
        }

        if (info.code === 'MODEL_NOT_FOUND' && !modelSwitched && mode === 'direct') {
          modelSwitched = true;
          try {
            const discovered = await pickBestFlashModel(apiKey);
            if (discovered && discovered !== model) {
              model = discovered;
              setModel(model);
              pushLog({ action: 'fallbackModel', model, status: 'warn', code: '404_RECOVERED', message: `Tự động chuyển sang model khả dụng: ${model}` });
              continue;
            }
          } catch (e2) {}
        }

        if (info.retryable && attempt < maxRetries) {
          const backoffMs = 800 * Math.pow(2, attempt);
          pushLog({ action: 'retry', model, status: 'warn', code: info.code, message: `Thử lại lần ${attempt + 1}/${maxRetries} sau ${backoffMs}ms` });
          await new Promise(r => setTimeout(r, backoffMs));
          continue;
        }
        break;
      }
    }

    const latencyMs = Math.round(performance.now() - t0);
    state.status = 'error';
    state.lastError = lastErr ? lastErr.userMessage : 'Lỗi gọi AI';
    state.lastLatencyMs = latencyMs;
    state.abortController = null;
    pushLog({
      action: options.actionName || 'generateContent',
      model,
      status: 'error',
      code: lastErr ? lastErr.code : 'UNKNOWN',
      latencyMs,
      message: state.lastError
    });
    throw lastErr || { code: 'UNKNOWN', userMessage: 'Lỗi gọi Gemini API.' };
  }

  function getState() {
    return {
      status: state.status,
      activeModel: getModel(),
      hasKey: Boolean(getApiKey()),
      maskedKey: maskKey(getApiKey()),
      mode: getConnectionMode(),
      proxyUrl: getProxyUrl(),
      lastLatencyMs: state.lastLatencyMs,
      lastError: state.lastError,
      availableModels: state.availableModels.slice(),
      isBusy: Boolean(state.abortController),
      logs: state.logs.slice()
    };
  }

  function subscribe(fn) {
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  global.VbtmAiService = {
    DEFAULT_MODELS,
    getState,
    subscribe,
    getApiKey,
    setApiKey,
    maskKey,
    getModel,
    setModel,
    getConnectionMode,
    setConnectionMode,
    getProxyUrl,
    listAvailableModels,
    testApiKey,
    testConnection,
    request,
    cancelActiveRequest,
    classifyError
  };
})(typeof window !== 'undefined' ? window : globalThis);
