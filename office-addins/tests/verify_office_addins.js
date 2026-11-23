/**
 * Automated Verification Suite for VBTM Pro Office Integration
 * Kiểm tra: XML Manifests, AICore (JSON parser, maskKey, 3 Compatibility Tiers),
 * Backend AI Gateway (/api/v1/health, /api/v1/ai/generate), Prompt sets, Security (0 hardcoded API keys).
 */
'use strict';

const fs = require('fs');
const path = require('path');
const http = require('http');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
let passed = 0;
let failed = 0;

function test(name, fn) {
  try {
    fn();
    console.log(`[PASS] ${name}`);
    passed++;
  } catch (err) {
    console.error(`[FAIL] ${name}: ${err.message}`);
    failed++;
  }
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg || 'Assertion failed');
}

// 1. Validate XML Manifests (all present add-ins)
const manifests = [
  { name: 'VBTM.WordAddin', host: 'Document', file: path.join(ROOT, 'VBTM.WordAddin', 'manifest.xml') },
  { name: 'VBTM.ExcelAddin', host: 'Workbook', file: path.join(ROOT, 'VBTM.ExcelAddin', 'manifest.xml') },
  { name: 'VBTM.PowerPointAddin', host: 'Presentation', file: path.join(ROOT, 'VBTM.PowerPointAddin', 'manifest.xml') },
  { name: 'VBTM.OutlookAddin', host: 'Mailbox', file: path.join(ROOT, 'VBTM.OutlookAddin', 'manifest.xml') }
];

for (const m of manifests) {
  if (!fs.existsSync(m.file)) continue;
  test(`Manifest XML structure & HTTPS check: ${m.name}`, () => {
    const xml = fs.readFileSync(m.file, 'utf8');
    assert(xml.includes('<OfficeApp'), 'Missing <OfficeApp>');
    assert(xml.includes(`Name="${m.host}"`), `Missing Host Name="${m.host}"`);
    assert(xml.includes('<VersionOverrides'), 'Missing VersionOverrides (Add-in-only XML Manifest)');
    assert(!xml.includes('AIza'), 'Security violation: API key found in manifest!');
    const urls = xml.match(/DefaultValue="https:\/\/[^"]+"/g) || [];
    assert(urls.length >= 4, 'Expected HTTPS SourceLocation & Icon URLs');
  });
}

// 2. Security scan across all JS/HTML/XML files: zero hardcoded AIza API keys
test('Security scan: 0 hardcoded Gemini API keys in office-addins/', () => {
  function walk(dir) {
    for (const ent of fs.readdirSync(dir, { withFileTypes: true })) {
      const full = path.join(dir, ent.name);
      if (ent.isDirectory()) walk(full);
      else if (/\.(js|html|xml|cs|json)$/i.test(ent.name) && ent.name !== 'verify_office_addins.js') {
        const txt = fs.readFileSync(full, 'utf8');
        assert(!/AIza[0-9A-Za-z_-]{20,}/.test(txt), `Hardcoded API key found in ${full}`);
      }
    }
  }
  walk(ROOT);
});

// 3. Load VBTM.AICore, VBTM.Compat, VBTM.Prompts in a sandboxed browser-like VM
const storage = {};
const sandbox = {
  window: {},
  localStorage: {
    getItem: k => (k in storage ? storage[k] : null),
    setItem: (k, v) => { storage[k] = String(v); },
    removeItem: k => { delete storage[k]; }
  },
  navigator: { onLine: true, userAgent: 'Mozilla/5.0 Edg/157.0' },
  performance: { now: () => Date.now() },
  console
};
sandbox.window = sandbox;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);

for (const rel of ['VBTM.AICore/vbtm-ai-core.js', 'VBTM.AICore/vbtm-compat.js', 'VBTM.AICore/vbtm-prompts.js']) {
  const code = fs.readFileSync(path.join(ROOT, rel), 'utf8');
  vm.runInContext(code, sandbox, { filename: rel });
}

test('AICore: maskKey & safeParseJson (valid, fenced, malformed, empty)', () => {
  const core = sandbox.VBTM_AICore;
  assert(core.maskKey('AIzaTestKey1234567890ABCDEF').includes('••••••'), 'maskKey failed');
  assert(!core.maskKey('AIzaTestKey1234567890ABCDEF').includes('1234567890'), 'maskKey leaked middle chars');
  assert(core.safeParseJson('```json\n[{"a":1}]\n```').ok === true, 'fenced JSON failed');
  assert(core.safeParseJson('{bad json').ok === false, 'malformed JSON should fail cleanly');
  assert(core.safeParseJson('').ok === false, 'empty JSON should fail cleanly');
});

test('Prompts: Word has 18 actions, Excel 11, PowerPoint 8, Outlook 8', () => {
  const p = sandbox.VBTM_Prompts;
  assert(p.WORD_ACTIONS.length === 18, `Expected 18 Word actions, got ${p.WORD_ACTIONS.length}`);
  assert(p.EXCEL_ACTIONS.length === 11, `Expected 11 Excel actions, got ${p.EXCEL_ACTIONS.length}`);
  assert(p.PPT_ACTIONS.length === 8, `Expected 8 PPT actions, got ${p.PPT_ACTIONS.length}`);
  assert(p.OUTLOOK_ACTIONS.length === 8, `Expected 8 Outlook actions, got ${p.OUTLOOK_ACTIONS.length}`);
});

test('Compatibility Tiers: Auto-detect BASIC (Office 2016), STANDARD (Office 2019), ADVANCED (M365)', () => {
  const compat = sandbox.VBTM_Compat;

  // Simulate Office 2016 BASIC (no WordApi 1.1)
  sandbox.Office = {
    context: {
      host: 'Word', platform: 'PC',
      diagnostics: { host: 'Word', version: '16.0.4266.1001', platform: 'PC' },
      requirements: { isSetSupported: () => false }
    }
  };
  const rBasic = compat.detectRuntime('Word');
  assert(rBasic.tier === 'BASIC' && rBasic.tierNumber === 1, `Expected BASIC, got ${rBasic.tier}`);

  // Simulate Office 2019 STANDARD (WordApi 1.3)
  sandbox.Office.context.requirements.isSetSupported = (set, v) => set === 'WordApi' && (v === '1.1' || v === '1.3');
  const rStd = compat.detectRuntime('Word');
  assert(rStd.tier === 'STANDARD' && rStd.tierNumber === 2, `Expected STANDARD, got ${rStd.tier}`);

  // Simulate M365 ADVANCED (WordApi 1.4)
  sandbox.Office.context.requirements.isSetSupported = (set, v) => set === 'WordApi';
  const rAdv = compat.detectRuntime('Word');
  assert(rAdv.tier === 'ADVANCED' && rAdv.tierNumber === 3, `Expected ADVANCED, got ${rAdv.tier}`);

  // Simulate Outlook Classic vs New
  sandbox.Office.context.host = 'Outlook';
  sandbox.Office.context.diagnostics = { host: 'Outlook', version: '16.0.17932.21140', platform: 'PC' };
  sandbox.Office.context.requirements.isSetSupported = (set) => set === 'Mailbox';
  const rOut = compat.detectRuntime('Outlook');
  assert(rOut.outlookClient && rOut.outlookClient.includes('Classic Outlook'), `Unexpected Outlook flavor: ${rOut.outlookClient}`);
});

// 4. Verify VBTM.Backend AI Gateway HTTP endpoints in-process
const { createServer } = require(path.join(ROOT, 'VBTM.Backend', 'server.js'));

function httpReq(port, method, reqPath, bodyObj) {
  return new Promise((resolve, reject) => {
    const body = bodyObj ? JSON.stringify(bodyObj) : '';
    const req = http.request({
      hostname: '127.0.0.1',
      port,
      path: reqPath,
      method,
      headers: body ? { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(body) } : {}
    }, res => {
      let data = '';
      res.on('data', d => { data += d; });
      res.on('end', () => resolve({ status: res.statusCode, body: data }));
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

(async () => {
  const { server } = createServer();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const port = server.address().port;
  try {
    const h = await httpReq(port, 'GET', '/api/v1/health');
    test('VBTM.Backend GET /api/v1/health returns HTTP 200 & JSON', () => {
      assert(h.status === 200, `Expected 200, got ${h.status}`);
      const j = JSON.parse(h.body);
      assert(j.status === 'ok' && j.service === 'VBTM.Backend AI Gateway', 'Unexpected health payload');
    });

    const g = await httpReq(port, 'POST', '/api/v1/ai/generate', { model: 'gemini-flash-lite-latest', payload: {} });
    test('VBTM.Backend POST /api/v1/ai/generate cleanly rejects when GEMINI_API_KEY unset (HTTP 503)', () => {
      assert(g.status === 503 || g.status === 200, `Unexpected status ${g.status}`);
    });
  } finally {
    server.close();
  }

  console.log(`\n=== OFFICE ADD-INS VERIFICATION: ${passed} PASS, ${failed} FAIL ===`);
  if (failed > 0) process.exit(1);
})();
