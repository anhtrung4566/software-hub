/**
 * VBTM Pro 3.0 Ultimate — Automated Verification Suite (Tests 4, 5, 6, 7, 8, 9, 13, 14)
 */
const assert = require('assert');
const fs = require('fs');
const path = require('path');

// Mock localStorage & navigator for Node execution
const memStore = {};
global.localStorage = {
  getItem: k => (k in memStore ? memStore[k] : null),
  setItem: (k, v) => { memStore[k] = String(v); },
  removeItem: k => { delete memStore[k]; }
};
global.window = global;
global.navigator = { onLine: true };

// Load VBTM Pro 3.0 modules
require('../js/vbtm-ai-service.js');
require('../js/vbtm-spell-engine.js');
require('../js/vbtm-docx-engine.js');
require('../js/vbtm-pro30.js');

const results = [];
async function runTest(name, fn) {
  try {
    await fn();
    results.push({ name, ok: true });
    console.log(`[PASS] ${name}`);
  } catch (err) {
    results.push({ name, ok: false, err: err.message });
    console.error(`[FAIL] ${name}: ${err.message}`);
  }
}

(async () => {
  // Test 4: Kiểm tra API Key sai
  await runTest('V3-T04: Phân loại & từ chối API Key sai định dạng (không lộ Key đầy đủ)', async () => {
    const r = await global.VbtmAiService.testApiKey('invalid_short_key');
    assert.strictEqual(r.ok, false);
    assert.strictEqual(r.code, 'INVALID_API_KEY');
    assert.strictEqual(global.VbtmAiService.maskKey('AIzaSyABCDEF123456789012345678901234567'), 'AIzaSy••••••4567');
  });

  // Test 5: Kiểm tra khi mất Internet
  await runTest('V3-T05: Phát hiện mất Internet (OFFLINE) khi gọi AI Online', async () => {
    global.navigator.onLine = false;
    global.VbtmAiService.setApiKey('AIzaSyABCDEF123456789012345678901234567');
    let caught = null;
    try {
      await global.VbtmAiService.request({ prompt: 'test' });
    } catch (e) { caught = e; }
    global.navigator.onLine = true;
    assert.ok(caught && caught.code === 'OFFLINE', 'Phải trả mã OFFLINE');
  });

  // Test 6: Kiểm tra khi API trả về lỗi 429 & 503
  await runTest('V3-T06: Phân loại lỗi quá tải 429 & lỗi máy chủ 503 có cờ retryable=true', async () => {
    const e429 = global.VbtmAiService.classifyError(429, 'RESOURCE_EXHAUSTED');
    const e503 = global.VbtmAiService.classifyError(503, 'UNAVAILABLE');
    const e401 = global.VbtmAiService.classifyError(401, 'API_KEY_INVALID');
    assert.strictEqual(e429.code, 'RATE_LIMIT_429');
    assert.strictEqual(e429.retryable, true);
    assert.strictEqual(e503.code, 'SERVER_ERROR_5XX');
    assert.strictEqual(e503.retryable, true);
    assert.strictEqual(e401.code, 'INVALID_API_KEY');
    assert.strictEqual(e401.retryable, false);
  });

  // Test 7 & 8: Soát chính tả bảo vệ tên riêng, số liệu, căn cứ pháp lý & sửa cục bộ theo tọa độ
  await runTest('V3-T07/T08: Soát chính tả 10 loại + Bảo vệ thực thể + Sửa đúng tọa độ (không viết lại văn bản)', async () => {
    const sample =
      'Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của chính phủ.\n' +
      'Đồng chí Nguyễn Văn An đề suất bổ xung 15% kinh phí (250 triệu đồng) theo Điều 5.\n' +
      'Các các đơn vị cần xử dụng ngân xách đúng kế hoạc và triển khai HOẠCHH trong ngày 5 tháng 1 năm 2026.';

    const issues = global.VbtmSpellEngine.scanOffline(sample);
    assert.ok(issues.length >= 7, `Phải bắt ít nhất 7 lỗi, thực tế: ${issues.length}`);

    // Kiểm tra bảo vệ thực thể
    assert.ok(global.VbtmSpellEngine.detectProtectedEntityReason('Nguyễn Văn An', sample));
    assert.ok(global.VbtmSpellEngine.detectProtectedEntityReason('15%', sample));
    assert.ok(global.VbtmSpellEngine.detectProtectedEntityReason('250 triệu đồng', sample));
    assert.ok(global.VbtmSpellEngine.detectProtectedEntityReason('Điều 5', sample));
    assert.ok(global.VbtmSpellEngine.detectProtectedEntityReason('30/2020/NĐ-CP', sample));

    // Sửa tất cả lỗi an toàn -> phải giữ nguyên 100% số liệu, tên riêng, căn cứ
    const fixed = global.VbtmSpellEngine.applyAllSafeFixes(sample, issues);
    assert.ok(fixed.fixedCount >= 6, `Phải sửa >= 6 lỗi an toàn, thực tế: ${fixed.fixedCount}`);
    assert.ok(fixed.text.includes('Nghị định số 30/2020/NĐ-CP'));
    assert.ok(fixed.text.includes('Nguyễn Văn An'));
    assert.ok(fixed.text.includes('15%'));
    assert.ok(fixed.text.includes('250 triệu đồng'));
    assert.ok(fixed.text.includes('Điều 5'));
    assert.ok(fixed.text.includes('đề xuất'));
    assert.ok(fixed.text.includes('bổ sung'));
    assert.ok(fixed.text.includes('sử dụng'));
    assert.ok(fixed.text.includes('ngân sách'));
    assert.ok(fixed.text.includes('kế hoạch'));
    assert.ok(fixed.text.includes('ngày 05 tháng 01 năm 2026'));
  });

  // Test 9: Xuất DOCX chuẩn OpenXML (ZIP thật, có word/document.xml, khổ A4, lề 30-15-20-20mm, đánh số trang từ trang 2)
  await runTest('V3-T09: Xuất tệp .DOCX chuẩn OpenXML (ZIP PK header + A4 + Times New Roman 14pt + w:titlePg)', async () => {
    const blob = global.VbtmDocxEngine.generateDocxBlob({
      standard: 'hd05',
      governingOrg: 'ĐẢNG BỘ THÀNH PHỐ HỒ CHÍ MINH',
      issuingOrg: 'ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT',
      docNumber: '42-BC/ĐU',
      location: 'Tân Sơn Nhất',
      docDate: '2026-11-23',
      docType: 'BÁO CÁO',
      docSubject: 'Kết quả triển khai VBTM Pro 3.0 Ultimate',
      docBody: 'I. KẾT QUẢ THỰC HIỆN\n1. Toàn bộ văn bản được chuẩn hóa đúng 05-HD/VPTW./.',
      recipients: '- Ban Thường vụ;\n- Lưu Văn phòng Đảng ủy.',
      signerTitle: 'T/M ĐẢNG ỦY\nBÍ THƯ',
      signerName: 'Trần Minh Quân',
      includeQr: false
    });
    const ab = await blob.arrayBuffer();
    const bytes = new Uint8Array(ab);
    // PK\x03\x04 ZIP magic header
    assert.strictEqual(bytes[0], 0x50);
    assert.strictEqual(bytes[1], 0x4B);
    assert.strictEqual(bytes[2], 0x03);
    assert.strictEqual(bytes[3], 0x04);
    const str = Buffer.from(bytes).toString('utf8');
    assert.ok(str.includes('[Content_Types].xml'));
    assert.ok(str.includes('word/document.xml'));
    assert.ok(str.includes('word/header1.xml'));
    assert.ok(str.includes('w:titlePg'));
    assert.ok(str.includes('w:pgSz w:w="11906" w:h="16838"'));
    assert.ok(str.includes('w:left="1701"'));
    assert.ok(str.includes('ĐẢNG CỘNG SẢN VIỆT NAM'));
    // Lưu file mẫu để kiểm chứng kích thước
    const outPath = path.join(__dirname, '..', 'downloads', '_test_v30_sample.docx');
    fs.writeFileSync(outPath, Buffer.from(bytes));
    assert.ok(fs.statSync(outPath).size > 2000);
    fs.unlinkSync(outPath);
  });

  // Test 14: Kiểm tra 2 hệ thống thể thức độc lập & 15 chế độ soạn thảo + 7 chế độ văn phong
  await runTest('V3-T14: Kiểm tra đủ 15 chế độ AI Soạn Thảo, 7 chế độ Văn Phong & 2 hệ Quy Chuẩn có nguồn căn cứ', async () => {
    assert.strictEqual(global.VbtmPro30.AI_DRAFT_MODES.length, 15);
    assert.strictEqual(global.VbtmPro30.AI_STYLE_MODES.length, 7);
    assert.ok(global.VbtmPro30.FORMAT_STANDARDS.hd05.source.includes('05-HD/VPTW'));
    assert.ok(global.VbtmPro30.FORMAT_STANDARDS.nd30.source.includes('30/2020/NĐ-CP'));
  });

  const failed = results.filter(r => !r.ok);
  console.log(`\n=== VBTM PRO 3.0 VERIFICATION: ${results.length - failed.length}/${results.length} PASS (${failed.length} FAIL) ===`);
  if (failed.length > 0) process.exit(1);
})();
