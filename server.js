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
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 8080;
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

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (pathname === '/api/health') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ status: 'ok', service: 'SieuAppVanPhong-Backend', timestamp: new Date().toISOString() }));
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
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
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

  let filePath = path.join(ROOT_DIR, pathname === '/' ? 'index.html' : pathname);
  if (!filePath.startsWith(ROOT_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
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

server.listen(PORT, () => {
  console.log(`Siêu App Văn Phòng Server running at http://localhost:${PORT}/`);
});
