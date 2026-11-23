/**
 * SIÊU APP VĂN PHÒNG (MONITER REPLICATION) - NODE.JS BACKEND & STATIC SERVER
 * Cung cấp REST API:
 *  - GET  /api/health
 *  - GET  /api/products
 *  - POST /api/products
 *  - GET  /api/stats?action=visit
 *  - POST /api/stats?action=download
 */

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 8080;
const aiRateLimitMap = new Map();
const ROOT_DIR = __dirname;
const PRODUCTS_FILE = path.join(ROOT_DIR, 'data', 'products.json');
const STATS_FILE = path.join(ROOT_DIR, 'data', 'stats.json');

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.exe': 'application/octet-stream',
  '.zip': 'application/zip',
  '.md': 'text/markdown; charset=utf-8'
};

function readStats() {
  try {
    if (fs.existsSync(STATS_FILE)) {
      return JSON.parse(fs.readFileSync(STATS_FILE, 'utf8'));
    }
  } catch (e) {}
  return {
    status: 'ok',
    date: new Date().toISOString().slice(0, 10),
    month: '2026-11',
    month_label: 'Tháng 11',
    today_visits: 358,
    today_downloads: 142,
    month_visits: 4820,
    month_downloads: 1680,
    total_visits: 6150,
    total_downloads: 2180
  };
}

function writeStats(stats) {
  try {
    fs.writeFileSync(STATS_FILE, JSON.stringify(stats, null, 2), 'utf8');
  } catch (e) {}
}

const server = http.createServer((req, res) => {
  const parsedUrl = url.parse(req.url, true);
  const pathname = decodeURIComponent(parsedUrl.pathname);

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (pathname === '/api/health') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ status: 'ok', service: 'SieuAppVanPhong-Backend', version: '3.0.0', aiProxyEnabled: Boolean(process.env.GEMINI_API_KEY), timestamp: new Date().toISOString() }));
    return;
  }

  // Backend Proxy bảo vệ Gemini API Key (chỉ bật khi quản trị viên cấu hình biến môi trường GEMINI_API_KEY)
  if (pathname === '/api/ai/gemini' && req.method === 'POST') {
    const serverKey = (process.env.GEMINI_API_KEY || '').trim();
    if (!serverKey) {
      res.writeHead(503, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ error: { code: 503, message: 'Backend chưa cấu hình biến môi trường GEMINI_API_KEY.' } }));
      return;
    }
    const clientIp = req.socket.remoteAddress || 'local';
    const now = Date.now();
    const bucket = (aiRateLimitMap.get(clientIp) || []).filter(ts => now - ts < 60000);
    if (bucket.length >= 30) {
      res.writeHead(429, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({ error: { code: 429, message: 'Vượt giới hạn 30 lượt gọi AI/phút từ địa chỉ này.' } }));
      return;
    }
    bucket.push(now);
    aiRateLimitMap.set(clientIp, bucket);

    let body = '';
    req.on('data', chunk => { if (body.length < 512 * 1024) body += chunk; });
    req.on('end', () => {
      try {
        const parsed = JSON.parse(body || '{}');
        const model = String(parsed.model || 'gemini-2.5-flash').replace(/[^a-zA-Z0-9._-]/g, '');
        const payloadStr = JSON.stringify(parsed.payload || {});
        const apiReq = https.request({
          hostname: 'generativelanguage.googleapis.com',
          path: `/v1beta/models/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(serverKey)}`,
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(payloadStr) },
          timeout: 25000
        }, apiRes => {
          let respBody = '';
          apiRes.on('data', d => { respBody += d; });
          apiRes.on('end', () => {
            res.writeHead(apiRes.statusCode || 500, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(respBody);
          });
        });
        apiReq.on('error', err => {
          res.writeHead(502, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: { code: 502, message: err.message } }));
        });
        apiReq.write(payloadStr);
        apiReq.end();
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: { code: 400, message: 'Invalid JSON request body.' } }));
      }
    });
    return;
  }

  if (pathname === '/api/stats' || pathname === '/api/stats.php') {
    const stats = readStats();
    const action = parsedUrl.query.action || 'visit';
    if (action === 'download' || req.method === 'POST') {
      stats.today_downloads += 1;
      stats.month_downloads += 1;
      stats.total_downloads += 1;
    } else {
      stats.today_visits += 1;
      stats.month_visits += 1;
      stats.total_visits += 1;
    }
    writeStats(stats);
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(stats));
    return;
  }

  if (pathname === '/api/products') {
    if (req.method === 'GET') {
      const data = fs.existsSync(PRODUCTS_FILE) ? fs.readFileSync(PRODUCTS_FILE, 'utf8') : '[]';
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(data);
      return;
    }
    if (req.method === 'POST') {
      let body = '';
      let tooLarge = false;
      const MAX_BODY_SIZE = 2 * 1024 * 1024;
      req.on('data', chunk => {
        body += chunk;
        if (body.length > MAX_BODY_SIZE) {
          tooLarge = true;
          req.destroy();
        }
      });
      req.on('end', () => {
        if (tooLarge) {
          res.writeHead(413, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ status: 'error', message: 'Payload too large' }));
          return;
        }
        try {
          const parsed = JSON.parse(body);
          if (Array.isArray(parsed)) {
            fs.writeFileSync(PRODUCTS_FILE, JSON.stringify(parsed, null, 2), 'utf8');
            res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
            res.end(JSON.stringify({ status: 'ok', count: parsed.length }));
            return;
          }
        } catch (err) {}
        res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ status: 'error', message: 'Invalid JSON payload' }));
      });
      return;
    }
  }

  const rawPathname = req.url.split('?')[0];
  if (rawPathname.includes('..') || pathname.includes('..') || pathname.includes('\0')) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('403 Forbidden');
    return;
  }

  let filePath = path.resolve(ROOT_DIR, '.' + (pathname === '/' ? '/index.html' : pathname));
  if (!filePath.startsWith(ROOT_DIR + path.sep) && filePath !== ROOT_DIR) {
    res.writeHead(403, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('403 Forbidden');
    return;
  }
  if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
    filePath = path.join(filePath, 'index.html');
  }
  if (!fs.existsSync(filePath)) {
    res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end('404 Not Found');
    return;
  }

  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';
  res.writeHead(200, { 'Content-Type': contentType });
  fs.createReadStream(filePath).pipe(res);
});

if (require.main === module) {
  server.listen(PORT, () => {
    console.log(`Siêu App Văn Phòng Server running at http://localhost:${PORT}/`);
  });
}

module.exports = { server };
