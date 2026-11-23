/**
 * VBTM.AICore — Shared AI Gateway Client for all Microsoft Office Add-ins
 * Dùng chung cho: VBTM.WordAddin, VBTM.ExcelAddin, VBTM.PowerPointAddin, VBTM.OutlookAddin
 *
 * Bảo mật:
 * - KHÔNG nhúng API Key bí mật vào mã nguồn, manifest hay bộ cài.
 * - Chế độ Gateway (khuyên dùng): gọi qua VBTM.Backend (/api/v1/ai/generate) bảo vệ API Key phía máy chủ.
 * - Chế độ Cá nhân (BYOK): người dùng tự nhập Gemini API Key lưu cục bộ trong localStorage của WebView2/Office,
 *   có cảnh báo rõ phương thức lưu trữ và rủi ro bảo mật.
 */
(function (global) {
  'use strict';

  const STORAGE = {
    MODE: 'vbtm_office_ai_mode',          // 'gateway' | 'byok'
    GATEWAY_URL: 'vbtm_office_gateway',   // e.g. https://localhost:8443/api/v1/ai/generate or http://127.0.0.1:8080/api/ai/gemini
    API_KEY: 'vbtm_office_gemini_key',
    MODEL: 'vbtm_office_gemini_model',
    HISTORY: 'vbtm_office_history_v1',
    LOGS: 'vbtm_office_tech_logs_v1'
  };

  const PREFERRED_MODELS = [
    'gemini-flash-lite-latest',
    'gemini-flash-latest',
    'gemini-3.5-flash-lite',
    'gemini-3.1-flash-lite',
    'gemini-3.5-flash',
    'gemini-3.6-flash',
    'gemini-3.7-flash',
    'gemini-3.8-flash',
    'gemini-2.5-flash'
  ];

  const state = {
    status: 'idle', // 'idle' | 'connecting' | 'connected' | 'error'
    activeModel: loadStr(STORAGE.MODEL, 'gemini-flash-lite-latest'),
    lastLatencyMs: 0,
    lastError: null,
    activeController: null,
    sessionId: 'sess_' + Date.now().toString(36)
  };

  const listeners = new Set();

  function loadStr(k, fallback) {
    try { return (localStorage.getItem(k) || fallback).trim(); } catch (e) { return fallback; }
  }
  function saveStr(k, v) {
    try {
      if (v) localStorage.setItem(k, String(v).trim());
      else localStorage.removeItem(k);
    } catch (e) {}
  }

  function emit() {
    const snap = getState();
    listeners.forEach(fn => { try { fn(snap); } catch (e) {} });
  }

  function maskKey(k) {
    const s = String(k || '').trim();
    if (!s) return '(chưa nhập)';
    if (s.length <= 10) return '****';
    return s.slice(0, 6) + '••••••' + s.slice(-4);
  }

  function writeTechLog(entry) {
    // Tuyệt đối KHÔNG ghi API Key hoặc nội dung văn bản công vụ vào log
    const rec = {
      ts: new Date().toISOString(),
      session: state.sessionId,
      host: entry.host || 'office',
      stage: entry.stage || 'generate',
      model: entry.model || state.activeModel,
      status: entry.status || 0,
      latencyMs: entry.latencyMs || 0,
      retry: entry.retry || 0,
      code: entry.code || 'OK',
      note: String(entry.note || '').replace(/AIza[0-9A-Za-z_-]{20,}/g, '[REDACTED]').slice(0, 180)
    };
    try {
      const arr = JSON.parse(localStorage.getItem(STORAGE.LOGS) || '[]');
      arr.unshift(rec);
      if (arr.length > 60) arr.length = 60;
      localStorage.setItem(STORAGE.LOGS, JSON.stringify(arr));
    } catch (e) {}
    return rec;
  }

  function getTechLogs() {
    try { return JSON.parse(localStorage.getItem(STORAGE.LOGS) || '[]'); } catch (e) { return []; }
  }

  function pushHistory(host, actionTitle, inputPreview, outputPreview) {
    try {
      const arr = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]');
      arr.unshift({
        id: 'h_' + Date.now(),
        ts: new Date().toLocaleTimeString('vi-VN'),
        host: host,
        action: actionTitle,
        inLen: (inputPreview || '').length,
        outLen: (outputPreview || '').length,
        preview: String(outputPreview || '').slice(0, 140)
      });
      if (arr.length > 25) arr.length = 25;
      localStorage.setItem(STORAGE.HISTORY, JSON.stringify(arr));
    } catch (e) {}
  }

  function getHistory(host) {
    try {
      const arr = JSON.parse(localStorage.getItem(STORAGE.HISTORY) || '[]');
      return host ? arr.filter(x => x.host === host) : arr;
    } catch (e) { return []; }
  }

  function classifyError(httpStatus, rawMsg, err) {
    const msg = String(rawMsg || (err && err.message) || '');
    if (err && err.name === 'AbortError') {
      return { code: 'ABORTED', retryable: false, failover: false, message: 'Đã hủy yêu cầu AI theo lệnh người dùng.' };
    }
    if (err && err.code === 'TIMEOUT') {
      return { code: 'TIMEOUT', retryable: true, failover: true, message: 'Quá thời gian chờ phản hồi từ máy chủ AI (Timeout).' };
    }
    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      return { code: 'OFFLINE', retryable: false, failover: false, message: 'Máy tính đang mất kết nối Internet.' };
    }
    if (httpStatus === 400) {
      if (/API key not valid|API_KEY_INVALID/i.test(msg)) {
        return { code: 'INVALID_KEY', retryable: false, failover: false, message: '[400] API Key Gemini không hợp lệ. Vui lòng kiểm tra lại trong Cài đặt.' };
      }
      return { code: 'BAD_REQUEST', retryable: false, failover: false, message: '[400] Yêu cầu không hợp lệ: ' + msg.slice(0, 160) };
    }
    if (httpStatus === 401 || httpStatus === 403) {
      return { code: 'AUTH_ERROR', retryable: false, failover: false, message: `[${httpStatus}] API Key không có quyền truy cập Gemini API.` };
    }
    if (httpStatus === 404) {
      return { code: 'MODEL_NOT_FOUND', retryable: false, failover: true, message: '[404] Model AI không khả dụng, hệ thống sẽ tự chuyển model dự phòng.' };
    }
    if (httpStatus === 429) {
      return { code: 'RATE_LIMIT', retryable: true, failover: true, message: '[429] Vượt giới hạn tần suất gọi API. Đang tự động chờ và thử lại.' };
    }
    if (httpStatus >= 500) {
      return { code: 'SERVER_ERROR', retryable: true, failover: true, message: `[${httpStatus}] Máy chủ AI đang quá tải hoặc bảo trì.` };
    }
    if (/Failed to fetch|NetworkError|Load failed|ERR_/i.test(msg)) {
      return { code: 'NETWORK_ERROR', retryable: true, failover: false, message: 'Lỗi kết nối mạng tới máy chủ AI / Gateway. Kiểm tra Internet hoặc HTTPS.' };
    }
    return { code: 'UNKNOWN', retryable: false, failover: false, message: msg.slice(0, 200) || 'Lỗi không xác định khi gọi AI.' };
  }

  function sleep(ms, signal) {
    return new Promise((resolve, reject) => {
      const t = setTimeout(resolve, ms);
      if (signal) {
        signal.addEventListener('abort', () => {
          clearTimeout(t);
          const e = new Error('Aborted');
          e.name = 'AbortError';
          reject(e);
        }, { once: true });
      }
    });
  }

  async function fetchWithTimeout(url, options, timeoutMs, outerSignal) {
    const ctrl = new AbortController();
    const timer = setTimeout(() => {
      const te = new Error('Timeout');
      te.code = 'TIMEOUT';
      ctrl.abort(te);
    }, timeoutMs);

    const onOuterAbort = () => ctrl.abort(outerSignal.reason || new DOMException('Aborted', 'AbortError'));
    if (outerSignal) {
      if (outerSignal.aborted) { clearTimeout(timer); throw new DOMException('Aborted', 'AbortError'); }
      outerSignal.addEventListener('abort', onOuterAbort, { once: true });
    }

    try {
      return await fetch(url, Object.assign({}, options, { signal: ctrl.signal }));
    } catch (err) {
      if (ctrl.signal.reason && ctrl.signal.reason.code === 'TIMEOUT') throw ctrl.signal.reason;
      throw err;
    } finally {
      clearTimeout(timer);
      if (outerSignal) outerSignal.removeEventListener('abort', onOuterAbort);
    }
  }

  function buildGeminiPayload(systemPrompt, userText, expectJson, includeThinkingZero) {
    const genConfig = {
      temperature: expectJson ? 0.2 : 0.4,
      maxOutputTokens: expectJson ? 2400 : 2048
    };
    if (expectJson) genConfig.responseMimeType = 'application/json';
    if (includeThinkingZero) genConfig.thinkingConfig = { thinkingBudget: 0 };
    return {
      contents: [{ parts: [{ text: (systemPrompt ? systemPrompt + '\n\n' : '') + userText }] }],
      generationConfig: genConfig
    };
  }

  function extractGeminiText(json) {
    const cands = json && Array.isArray(json.candidates) ? json.candidates : [];
    if (!cands.length) throw new Error('AI trả về phản hồi rỗng (không có candidates).');
    const parts = cands[0].content && Array.isArray(cands[0].content.parts) ? cands[0].content.parts : [];
    const txt = parts.map(p => p.text || '').join('').trim();
    if (!txt) throw new Error('AI trả về nội dung văn bản rỗng.');
    return txt;
  }

  function safeParseJson(rawText) {
    if (!rawText || typeof rawText !== 'string') return { ok: false, data: null, error: 'Phản hồi rỗng' };
    let cleaned = rawText.trim();
    cleaned = cleaned.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/i, '').trim();
    try {
      return { ok: true, data: JSON.parse(cleaned), error: null };
    } catch (e1) {
      const arrMatch = cleaned.match(/\[[\s\S]*\]/);
      if (arrMatch) {
        try { return { ok: true, data: JSON.parse(arrMatch[0]), error: null }; } catch (e2) {}
      }
      const objMatch = cleaned.match(/\{[\s\S]*\}/);
      if (objMatch) {
        try { return { ok: true, data: JSON.parse(objMatch[0]), error: null }; } catch (e3) {}
      }
      return { ok: false, data: null, error: 'JSON trả về không hợp lệ.' };
    }
  }

  async function callSingleModel(model, systemPrompt, userText, opts, outerSignal) {
    const mode = getMode();
    const timeoutMs = opts.timeoutMs || 28000;
    const maxRetries = opts.maxRetries !== undefined ? opts.maxRetries : 2;
    let useThinkingZero = true;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      const t0 = performance.now();
      try {
        const payload = buildGeminiPayload(systemPrompt, userText, !!opts.expectJson, useThinkingZero);
        let res;
        if (mode === 'gateway') {
          const gw = getGatewayUrl();
          if (!gw) throw new Error('Chưa cấu hình địa chỉ VBTM AI Gateway.');
          res = await fetchWithTimeout(gw, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ model: model, payload: payload, host: opts.host || 'office' })
          }, timeoutMs, outerSignal);
        } else {
          const key = getApiKey();
          if (!key) throw new Error('Chưa nhập Gemini API Key trong Cài đặt.');
          const url = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(key)}`;
          res = await fetchWithTimeout(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          }, timeoutMs, outerSignal);
        }

        const latency = Math.round(performance.now() - t0);
        const bodyText = await res.text();
        let parsed = null;
        try { parsed = JSON.parse(bodyText); } catch (e) {}

        if (!res.ok) {
          const errMsg = (parsed && parsed.error && parsed.error.message) ? parsed.error.message : bodyText.slice(0, 200);
          if (res.status === 400 && /thinking/i.test(errMsg) && useThinkingZero) {
            useThinkingZero = false;
            continue;
          }
          const cls = classifyError(res.status, errMsg, null);
          writeTechLog({ host: opts.host, stage: 'generate', model, status: res.status, latencyMs: latency, retry: attempt, code: cls.code, note: cls.message });
          if (cls.retryable && attempt < maxRetries) {
            const backoff = Math.min(4000, 600 * Math.pow(2, attempt)) + Math.floor(Math.random() * 250);
            await sleep(backoff, outerSignal);
            continue;
          }
          const err = new Error(cls.message);
          err.info = cls;
          throw err;
        }

        const text = extractGeminiText(parsed);
        writeTechLog({ host: opts.host, stage: 'generate', model, status: 200, latencyMs: latency, retry: attempt, code: 'OK', note: 'len=' + text.length });
        return { text, model, latencyMs: latency };
      } catch (err) {
        if (err && err.info) throw err;
        const latency = Math.round(performance.now() - t0);
        const cls = classifyError(0, err && err.message, err);
        writeTechLog({ host: opts.host, stage: 'generate', model, status: 0, latencyMs: latency, retry: attempt, code: cls.code, note: cls.message });
        if (cls.retryable && attempt < maxRetries) {
          const backoff = Math.min(4000, 600 * Math.pow(2, attempt)) + Math.floor(Math.random() * 250);
          await sleep(backoff, outerSignal);
          continue;
        }
        const wrapped = new Error(cls.message);
        wrapped.info = cls;
        throw wrapped;
      }
    }
  }

  async function generate(systemPrompt, userText, opts) {
    opts = opts || {};
    cancelActive();
    const ctrl = new AbortController();
    state.activeController = ctrl;
    state.status = 'connecting';
    state.lastError = null;
    emit();

    const primary = getModel();
    const candidates = [primary].concat(PREFERRED_MODELS.filter(m => m !== primary));
    let lastErr = null;

    try {
      for (let i = 0; i < candidates.length; i++) {
        const m = candidates[i];
        try {
          const res = await callSingleModel(m, systemPrompt, userText, opts, ctrl.signal);
          state.status = 'connected';
          state.activeModel = res.model;
          state.lastLatencyMs = res.latencyMs;
          state.lastError = null;
          emit();
          if (opts.historyTitle) pushHistory(opts.host || 'office', opts.historyTitle, userText, res.text);
          return res;
        } catch (err) {
          lastErr = err;
          const canFailover = err && err.info && err.info.failover;
          if (!canFailover) throw err;
        }
      }
      throw lastErr || new Error('Không gọi được model Gemini khả dụng.');
    } catch (err) {
      state.status = 'error';
      state.lastError = err.message || 'Lỗi gọi AI';
      emit();
      throw err;
    } finally {
      if (state.activeController === ctrl) state.activeController = null;
    }
  }

  async function testConnection(host) {
    const res = await generate('Chỉ trả lời đúng 1 từ: OK', 'Kiểm tra kết nối VBTM Pro Office AI.', {
      host: host || 'office',
      expectJson: false,
      maxRetries: 1,
      timeoutMs: 15000
    });
    return res;
  }

  function cancelActive() {
    if (state.activeController) {
      try { state.activeController.abort(); } catch (e) {}
      state.activeController = null;
    }
  }

  function getMode() { return loadStr(STORAGE.MODE, 'byok'); }
  function setMode(m) { saveStr(STORAGE.MODE, m === 'gateway' ? 'gateway' : 'byok'); emit(); }
  function getGatewayUrl() { return loadStr(STORAGE.GATEWAY_URL, 'https://localhost:8443/api/v1/ai/generate'); }
  function setGatewayUrl(u) { saveStr(STORAGE.GATEWAY_URL, u); emit(); }
  function getApiKey() { return loadStr(STORAGE.API_KEY, ''); }
  function setApiKey(k) { saveStr(STORAGE.API_KEY, k); emit(); }
  function getModel() { return loadStr(STORAGE.MODEL, 'gemini-flash-lite-latest'); }
  function setModel(m) { state.activeModel = m || 'gemini-flash-lite-latest'; saveStr(STORAGE.MODEL, state.activeModel); emit(); }

  function getState() {
    return {
      status: state.status,
      mode: getMode(),
      gatewayUrl: getGatewayUrl(),
      hasKey: !!getApiKey(),
      maskedKey: maskKey(getApiKey()),
      activeModel: state.activeModel,
      lastLatencyMs: state.lastLatencyMs,
      lastError: state.lastError,
      sessionId: state.sessionId
    };
  }

  function onStateChange(fn) {
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  global.VBTM_AICore = {
    PREFERRED_MODELS,
    generate,
    testConnection,
    cancelActive,
    safeParseJson,
    getState,
    onStateChange,
    getMode, setMode,
    getGatewayUrl, setGatewayUrl,
    getApiKey, setApiKey,
    getModel, setModel,
    maskKey,
    getHistory,
    getTechLogs
  };
})(typeof window !== 'undefined' ? window : globalThis);
