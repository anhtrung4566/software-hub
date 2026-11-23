/**
 * VBTM.Backend — Centralized VBTM AI Gateway for Microsoft Office Add-ins
 * Chức năng:
 * - Bảo vệ Gemini API Key phía máy chủ (đọc từ biến môi trường GEMINI_API_KEY, không lộ ra trình duyệt/Add-in)
 * - Tự động failover model khi 404 / 429 / 503
 * - Rate limit theo IP, timeout 25s, kiểm tra cấu trúc request/response
 * - Ghi log kỹ thuật (KHÔNG ghi API Key, KHÔNG ghi nội dung văn bản công vụ)
 * - Hỗ trợ cả HTTP (local dev) và HTTPS (khi cung cấp VBTM_TLS_KEY & VBTM_TLS_CERT)
 */
'use strict';

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = Number(process.env.PORT || 8443);
const MAX_BODY_BYTES = 512 * 1024;
const RATE_LIMIT_PER_MIN = Number(process.env.VBTM_RATE_LIMIT || 40);
const rateBuckets = new Map();

const FALLBACK_MODELS = [
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

function logTech(stage, model, status, latencyMs, note) {
  const safeNote = String(note || '').replace(/AIza[0-9A-Za-z_-]{20,}/g, '[REDACTED]');
  console.log(`[VBTM-Gateway] ${new Date().toISOString()} | stage=${stage} | model=${model} | http=${status} | ms=${latencyMs} | ${safeNote}`);
}

function sendJson(res, status, obj) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'X-Content-Type-Options': 'nosniff'
  });
  res.end(JSON.stringify(obj));
}

function callGeminiOnce(apiKey, model, payloadObj, timeoutMs) {
  return new Promise((resolve, reject) => {
    const bodyStr = JSON.stringify(payloadObj);
    const t0 = Date.now();
    const req = https.request({
      hostname: 'generativelanguage.googleapis.com',
      path: `/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(bodyStr)
      },
      timeout: timeoutMs || 25000
    }, (res) => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        resolve({ status: res.statusCode || 500, body: data, latencyMs: Date.now() - t0 });
      });
    });
    req.on('timeout', () => {
      req.destroy(new Error('TIMEOUT'));
    });
    req.on('error', err => reject(err));
    req.write(bodyStr);
    req.end();
  });
}

function handleRequest(req, res) {
  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type, Authorization'
    });
    res.end();
    return;
  }

  if (pathname === '/api/v1/health' && req.method === 'GET') {
    return sendJson(res, 200, {
      status: 'ok',
      service: 'VBTM.Backend AI Gateway',
      version: '3.0.0',
      geminiConfigured: Boolean((process.env.GEMINI_API_KEY || '').trim()),
      models: FALLBACK_MODELS,
      timestamp: new Date().toISOString()
    });
  }

  if (pathname === '/api/v1/ai/ollama/status' && req.method === 'GET') {
    const oReq = http.request({ hostname: '127.0.0.1', port: 11434, path: '/api/tags', method: 'GET', timeout: 3000 }, oRes => {
      let d = '';
      oRes.on('data', c => { d += c; });
      oRes.on('end', () => {
        try {
          const j = JSON.parse(d || '{}');
          const models = (j.models || []).map(m => m.name);
          return sendJson(res, 200, { running: true, models, endpoint: 'http://localhost:11434' });
        } catch (e) {
          return sendJson(res, 200, { running: false, models: [], error: 'Invalid Ollama response' });
        }
      });
    });
    oReq.on('error', () => sendJson(res, 200, { running: false, models: [], error: 'Ollama chưa chạy tại http://localhost:11434' }));
    oReq.on('timeout', () => { oReq.destroy(); sendJson(res, 200, { running: false, models: [], error: 'Ollama timeout' }); });
    oReq.end();
    return;
  }

  if ((pathname === '/api/v1/ai/generate' || pathname === '/api/ai/gemini') && req.method === 'POST') {
    const apiKey = (process.env.GEMINI_API_KEY || '').trim();
    if (!apiKey) {
      // Không chặn hệ thống nếu chưa cấu hình Gemini Key — thử chuyển sang Ollama Local (http://127.0.0.1:11434)
      let rawO = '';
      req.on('data', c => { rawO += c; });
      req.on('end', () => {
        try {
          const parsed = JSON.parse(rawO || '{}');
          const textPrompt = parsed.payload && parsed.payload.contents && parsed.payload.contents[0] && parsed.payload.contents[0].parts
            ? parsed.payload.contents[0].parts.map(p => p.text || '').join('\n')
            : String(parsed.prompt || '');
          const oBody = JSON.stringify({ model: parsed.ollamaModel || 'qwen2.5:7b', prompt: textPrompt, stream: false });
          const oReq = http.request({
            hostname: '127.0.0.1', port: 11434, path: '/api/generate', method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(oBody) },
            timeout: 45000
          }, oRes => {
            let od = '';
            oRes.on('data', ch => { od += ch; });
            oRes.on('end', () => {
              if (oRes.statusCode >= 200 && oRes.statusCode < 300) {
                try {
                  const oj = JSON.parse(od);
                  return sendJson(res, 200, {
                    provider: 'ollama',
                    candidates: [{ content: { parts: [{ text: oj.response || '' }] } }]
                  });
                } catch (e) {}
              }
              return sendJson(res, 503, { error: { code: 503, message: 'Ollama Local chưa sẵn sàng và chưa cấu hình GEMINI_API_KEY. Hãy cài Ollama (miễn phí, không cần API Key).' } });
            });
          });
          oReq.on('error', () => sendJson(res, 503, { error: { code: 503, message: 'Ollama Local chưa chạy tại http://localhost:11434 và chưa cấu hình GEMINI_API_KEY.' } }));
          oReq.write(oBody);
          oReq.end();
        } catch (e) {
          return sendJson(res, 400, { error: { code: 400, message: 'JSON request không hợp lệ.' } });
        }
      });
      return;
    }

    const ip = req.socket.remoteAddress || 'local';
    const now = Date.now();
    const recent = (rateBuckets.get(ip) || []).filter(t => now - t < 60000);
    if (recent.length >= RATE_LIMIT_PER_MIN) {
      logTech('rate_limit', '-', 429, 0, `ip=${ip}`);
      return sendJson(res, 429, { error: { code: 429, message: `Vượt giới hạn ${RATE_LIMIT_PER_MIN} lượt gọi/phút.` } });
    }
    recent.push(now);
    rateBuckets.set(ip, recent);

    let raw = '';
    let exceeded = false;
    req.on('data', chunk => {
      raw += chunk;
      if (Buffer.byteLength(raw) > MAX_BODY_BYTES) {
        exceeded = true;
        req.destroy();
      }
    });
    req.on('end', async () => {
      if (exceeded) return sendJson(res, 413, { error: { code: 413, message: 'Payload vượt giới hạn 512KB.' } });
      let parsed;
      try { parsed = JSON.parse(raw || '{}'); } catch (e) {
        return sendJson(res, 400, { error: { code: 400, message: 'JSON request không hợp lệ.' } });
      }
      if (!parsed.payload || typeof parsed.payload !== 'object') {
        return sendJson(res, 400, { error: { code: 400, message: 'Thiếu trường payload.' } });
      }
      const reqModel = String(parsed.model || FALLBACK_MODELS[0]).replace(/[^a-zA-Z0-9._-]/g, '');
      const candidates = [reqModel].concat(FALLBACK_MODELS.filter(m => m !== reqModel));

      let lastStatus = 502;
      let lastBody = JSON.stringify({ error: { code: 502, message: 'Không nhận được phản hồi từ Gemini.' } });

      for (const m of candidates) {
        try {
          const r = await callGeminiOnce(apiKey, m, parsed.payload, 25000);
          logTech('generate', m, r.status, r.latencyMs, `host=${parsed.host || 'office'}`);
          if (r.status >= 200 && r.status < 300) {
            res.writeHead(r.status, {
              'Content-Type': 'application/json; charset=utf-8',
              'Access-Control-Allow-Origin': '*',
              'X-VBTM-Model': m
            });
            res.end(r.body);
            return;
          }
          lastStatus = r.status;
          lastBody = r.body;
          if (r.status === 404 || r.status === 429 || r.status === 503) continue;
          break;
        } catch (err) {
          logTech('generate_err', m, 0, 0, err.message);
          if (err.message === 'TIMEOUT') { lastStatus = 408; continue; }
          return sendJson(res, 502, { error: { code: 502, message: 'Lỗi mạng từ Gateway tới Gemini: ' + err.message } });
        }
      }

      res.writeHead(lastStatus, {
        'Content-Type': 'application/json; charset=utf-8',
        'Access-Control-Allow-Origin': '*'
      });
      res.end(lastBody);
    });
    return;
  }

  return sendJson(res, 404, { error: { code: 404, message: 'Endpoint không tồn tại trên VBTM.Backend.' } });
}

function createServer() {
  const keyPath = process.env.VBTM_TLS_KEY;
  const certPath = process.env.VBTM_TLS_CERT;
  if (keyPath && certPath && fs.existsSync(keyPath) && fs.existsSync(certPath)) {
    return {
      proto: 'https',
      server: https.createServer({
        key: fs.readFileSync(keyPath),
        cert: fs.readFileSync(certPath)
      }, handleRequest)
    };
  }
  return { proto: 'http', server: http.createServer(handleRequest) };
}

if (require.main === module) {
  const { proto, server } = createServer();
  server.listen(PORT, '127.0.0.1', () => {
    console.log(`VBTM.Backend AI Gateway listening at ${proto}://127.0.0.1:${PORT}/api/v1/health`);
  });
}

module.exports = { createServer, handleRequest, FALLBACK_MODELS };
