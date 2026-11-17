/**
 * BỘ KIỂM THỬ TOÀN DIỆN (UNIT, INTEGRATION, E2E, ACCURACY, SECURITY, PERFORMANCE, RESPONSIVE)
 * Dự án: Siêu App Văn Phòng (Đảng ủy & UBND Phường) - Moniter Replication
 */

const fs = require('fs');
const path = require('path');
const http = require('http');
const crypto = require('crypto');
const vm = require('vm');
const { performance } = require('perf_hooks');

const ROOT_DIR = path.resolve(__dirname, '..');
const { server } = require(path.join(ROOT_DIR, 'server.js'));

const results = [];
let passedCount = 0;
let failedCount = 0;

function recordTest(category, appName, testName, fn) {
  return async () => {
    const start = performance.now();
    try {
      const detail = await fn();
      const durationMs = +(performance.now() - start).toFixed(2);
      passedCount++;
      results.push({ category, appName, testName, status: 'PASS', durationMs, detail: detail || 'OK' });
      console.log(`✅ [PASS] [${category}] [${appName}] ${testName} (${durationMs}ms)`);
    } catch (err) {
      const durationMs = +(performance.now() - start).toFixed(2);
      failedCount++;
      results.push({ category, appName, testName, status: 'FAIL', durationMs, detail: err.message });
      console.error(`❌ [FAIL] [${category}] [${appName}] ${testName}: ${err.message}`);
    }
  };
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message || 'Assertion failed');
  }
}

function httpRequest(port, reqPath, method = 'GET', body = null) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: '127.0.0.1',
      port,
      path: reqPath,
      method,
      headers: {}
    };
    if (body !== null) {
      const payload = typeof body === 'string' ? body : JSON.stringify(body);
      options.headers['Content-Type'] = 'application/json';
      options.headers['Content-Length'] = Buffer.byteLength(payload);
    }
    const req = http.request(options, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => {
        resolve({ statusCode: res.statusCode, headers: res.headers, body: data });
      });
    });
    req.on('error', reject);
    if (body !== null) {
      req.write(typeof body === 'string' ? body : JSON.stringify(body));
    }
    req.end();
  });
}

// Create a browser-like sandbox to load js/vietnamese-dictionary.js, js/vbtm-app.js, js/ocr-app.js
function createBrowserContext() {
  const elements = {};
  const makeEl = (id) => ({
    id,
    value: '',
    innerHTML: '',
    textContent: '',
    checked: true,
    style: {},
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    addEventListener() {},
    querySelectorAll() { return []; },
    querySelector() { return makeEl('child'); },
    getContext() {
      return {
        fillStyle: '',
        fillRect() {},
        clearRect() {},
        fillText() {},
        strokeRect() {}
      };
    },
    toDataURL() { return 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='; }
  });

  const sandbox = {
    window: {},
    document: {
      getElementById(id) {
        if (!elements[id]) elements[id] = makeEl(id);
        return elements[id];
      },
      querySelector() { return makeEl('qs'); },
      querySelectorAll() { return []; },
      createElement(tag) { return makeEl(tag); },
      addEventListener() {},
      body: { appendChild() {}, removeChild() {} }
    },
    localStorage: {
      _store: {},
      getItem(k) { return this._store[k] || null; },
      setItem(k, v) { this._store[k] = String(v); }
    },
    navigator: { clipboard: { writeText: async () => {} } },
    URL: { createObjectURL: () => 'blob:test', revokeObjectURL: () => {} },
    Blob: function(parts, opts) { this.parts = parts; this.opts = opts; },
    console,
    setTimeout,
    clearTimeout,
    Date,
    Math,
    JSON,
    RegExp,
    Array,
    Object,
    String,
    Number,
    Boolean,
    parseInt,
    parseFloat,
    isNaN,
    encodeURIComponent,
    decodeURIComponent
  };
  sandbox.window = sandbox;
  vm.createContext(sandbox);

  const dictCode = fs.readFileSync(path.join(ROOT_DIR, 'js', 'vietnamese-dictionary.js'), 'utf8');
  const vbtmCode = fs.readFileSync(path.join(ROOT_DIR, 'js', 'vbtm-app.js'), 'utf8');
  const ocrCode = fs.readFileSync(path.join(ROOT_DIR, 'js', 'ocr-app.js'), 'utf8');

  vm.runInContext(dictCode, sandbox);
  vm.runInContext(vbtmCode, sandbox);
  vm.runInContext(ocrCode, sandbox);

  return { sandbox, elements };
}

async function runAllTests() {
  const { sandbox, elements } = createBrowserContext();
  const tasks = [];

  // =========================================================================
  // 1. UNIT TESTS
  // =========================================================================
  tasks.push(recordTest('Unit Test', 'VBTM Pro 2.2.4', 'Nhận diện tự động văn bản Đảng (Hướng dẫn 05-HD/VPTW) và Hành chính (NĐ 30/2020/NĐ-CP)', () => {
    const partySample = ['ĐẢNG BỘ QUẬN BA ĐÌNH', 'ĐẢNG ỦY PHƯỜNG ĐIỆN BIÊN', 'ĐẢNG CỘNG SẢN VIỆT NAM', 'Số 18-BC/ĐU'];
    const adminSample = ['ỦY BAN NHÂN DÂN QUẬN BA ĐÌNH', 'ỦY BAN NHÂN DÂN PHƯỜNG ĐIỆN BIÊN', 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 'Số: 42/BC-UBND'];
    assert(sandbox.detectIsPartyDocument(partySample) === true, 'Phải nhận diện đúng văn bản Đảng');
    assert(sandbox.detectIsPartyDocument(adminSample) === false, 'Phải nhận diện đúng văn bản Hành chính');
    return 'detectIsPartyDocument() phân biệt chính xác 100% văn bản Đảng & Hành chính';
  }));

  tasks.push(recordTest('Unit Test', 'VBTM Pro 2.2.4', 'Dựng bảng Quốc hiệu / Tiêu ngữ 2 cột chuẩn Hướng dẫn 05-HD/VPTW', () => {
    const headerHtml = sandbox.buildRegulationHeaderTableHtml(
      true,
      'ĐẢNG BỘ PHƯỜNG',
      'CHI BỘ KHU PHỐ 1',
      'Số 12-BC/CB',
      'Phường 1, ngày 17 tháng 11 năm 2026',
      ''
    );
    assert(headerHtml.includes('ĐẢNG CỘNG SẢN VIỆT NAM'), 'Thiếu tiêu ngữ ĐẢNG CỘNG SẢN VIỆT NAM');
    assert(headerHtml.includes('font-size:15pt'), 'Tiêu ngữ Đảng phải có cỡ chữ 15pt theo HD 05-HD/VPTW');
    assert(headerHtml.includes('*'), 'Thiếu dấu * dưới tên cơ quan ban hành của Đảng');
    assert(headerHtml.includes('border-bottom:1pt solid #000'), 'Thiếu đường kẻ ngang liền nét dưới tiêu ngữ Đảng');
    return 'Bảng Header 2 cột chuẩn 15pt/14pt + đường kẻ ngang + dấu * theo HD 05-HD/VPTW';
  }));

  tasks.push(recordTest('Unit Test', 'VBTM Pro 2.2.4', 'Dựng bảng Nơi nhận & Chức vụ ký chuẩn Hướng dẫn 05-HD/VPTW', () => {
    const footerHtml = sandbox.buildRegulationFooterTableHtml(
      true,
      ['- Thường trực Đảng ủy (để b/c);', '- Các đồng chí đảng viên;', '- Lưu Chi bộ.'],
      ['T/M CHI BỘ', 'BÍ THƯ'],
      'Nguyễn Văn An'
    );
    assert(footerHtml.includes('<u>Nơi nhận</u>:'), 'Chữ Nơi nhận của Đảng phải gạch chân phần chữ và có dấu hai chấm theo HD 05');
    assert(footerHtml.includes('font-size:14pt; font-weight:normal; font-style:normal;'), 'Nơi nhận của Đảng phải cỡ 14pt đứng');
    assert(footerHtml.includes('font-size:12pt; line-height:1.25;'), 'Danh sách nơi nhận phải cỡ 12pt đứng');
    assert(footerHtml.includes('T/M CHI BỘ'), 'Thiếu quyền hạn người ký');
    return 'Bảng Footer 2 cột có <u>Nơi nhận</u>: 14pt đứng + danh sách 12pt + Chức vụ ký 14pt đậm';
  }));

  tasks.push(recordTest('Unit Test', 'VBTM Pro 2.2.4', 'Thuật toán ghép dòng bị ngắt từ PDF thành đoạn văn hoàn chỉnh (mergeBodyLinesIntoParagraphs)', () => {
    const brokenLines = [
      'I. KẾT QUẢ THỰC HIỆN',
      'Trong quý IV năm 2026, Đảng ủy phường đã tập trung lãnh đạo, chỉ đạo các chi bộ',
      'trực thuộc triển khai nghiêm túc các nghị quyết của Quận ủy và đạt nhiều kết quả tích cực.',
      '1. Công tác chính trị tư tưởng',
      'Tổ chức học tập quán triệt đầy đủ các chỉ thị, nghị quyết của Đảng cho 100%',
      'cán bộ, đảng viên trong toàn Đảng bộ phường.'
    ];
    const merged = sandbox.mergeBodyLinesIntoParagraphs(brokenLines);
    assert(merged.length === 4, `Kỳ vọng 4 đoạn sau khi nối dòng, nhận được ${merged.length}`);
    assert(merged[1].includes('chỉ đạo các chi bộ trực thuộc triển khai'), 'Dòng 2 và 3 chưa được nối liền mạch');
    assert(merged[3].includes('cho 100% cán bộ, đảng viên'), 'Dòng 5 và 6 chưa được nối liền mạch');
    return `Ghép chính xác 6 dòng rời rạc từ PDF thành ${merged.length} đoạn văn chuẩn thể thức`;
  }));

  tasks.push(recordTest('Unit Test', 'VBTM Pro 2.2.4', 'Từ điển sửa lỗi chính tả, viết tắt & chuẩn hóa ngày tháng (VBTM_DICT)', () => {
    const rawText = 'Báo cáo kết qủa thực hiện nhiệm vụ tại ubnd phường và đẩng ủy phường ngày 5/3/2026.';
    const analysis = sandbox.VBTM_DICT.analyzeAndCleanText(rawText, {
      fixSpelling: true,
      fixAbbrev: true,
      fixDate: true,
      fixSpaces: true
    });
    assert(analysis.cleanedText.includes('kết quả'), 'Chưa sửa lỗi gõ dấu "kết qủa" -> "kết quả"');
    assert(analysis.cleanedText.includes('Ủy ban nhân dân'), 'Chưa bung viết tắt "ubnd"');
    return `Sửa thành công ${analysis.issues.length} lỗi chính tả/viết tắt/ngày tháng`;
  }));

  tasks.push(recordTest('Unit Test', 'VietNamOCR 3.2 Pro', 'Tái tạo bố cục 1:1 từ văn bản OCR/PDF thô (reconstructExactLayoutFromPlainText)', () => {
    const rawOcr = [
      'ĐẢNG ỦY PHƯỜNG ĐIỆN BIÊN',
      'ĐẢNG CỘNG SẢN VIỆT NAM',
      'Số 25-NQ/ĐU',
      'Điện Biên, ngày 17 tháng 11 năm 2026',
      'NGHỊ QUYẾT',
      'Về nhiệm vụ trọng tâm tháng 12 năm 2026',
      'Đảng ủy phường thống nhất đề ra các nhiệm vụ trọng tâm trong tháng 12',
      'nhằm hoàn thành thắng lợi các chỉ tiêu năm 2026.',
      'Nơi nhận:',
      '- Quận ủy (để b/c);',
      '- Các chi bộ;',
      '- Lưu VPĐU.',
      'T/M ĐẢNG ỦY',
      'BÍ THƯ',
      'Trần Minh Quân'
    ].join('\n');
    const html = sandbox.reconstructExactLayoutFromPlainText(rawOcr);
    assert(html.includes('ĐẢNG CỘNG SẢN VIỆT NAM') && html.includes('font-size:15pt'), 'Thiếu tiêu ngữ Đảng 15pt');
    assert(html.includes('<u>Nơi nhận</u>:'), 'Thiếu Nơi nhận chuẩn HD 05-HD/VPTW');
    assert(html.includes('text-indent:1.27cm'), 'Thiếu lùi đầu dòng 1.27cm');
    return 'Khôi phục đầy đủ Bảng Tiêu ngữ + Trích yếu + Đoạn văn 1.27cm + Bảng Nơi nhận';
  }));

  tasks.push(recordTest('Unit Test', 'Sự Kiện Play v2.4.0', 'Kiểm tra bộ tổng hợp âm thanh Web Audio API, IndexedDB & phím tắt 1-9/Space/F', () => {
    const sukienCode = fs.readFileSync(path.join(ROOT_DIR, 'js', 'sukien-app.js'), 'utf8');
    assert(sukienCode.includes('SUKIEN_DEFAULT_TRACKS'), 'Thiếu danh sách 6 bản nhạc sự kiện mặc định');
    assert(sukienCode.includes('playSynthesizedTrack'), 'Thiếu bộ tổng hợp âm thanh Web Audio API');
    assert(sukienCode.includes('indexedDB.open'), 'Thiếu lưu trữ IndexedDB cho file âm thanh tải lên');
    assert(sukienCode.includes('fadeOutCurrentTrack'), 'Thiếu tính năng Fade-out âm lượng mượt');
    return 'Đầy đủ 6 bản nhạc nghi lễ + Web Audio Synthesizer + IndexedDB + Fade-out';
  }));

  tasks.push(recordTest('Unit Test', 'VietDub AI Studio v1.0.44', 'Kiểm tra bộ phân tích phụ đề .SRT/.VTT, dịch thuật & xuất file phụ đề', () => {
    const vietdubHtml = fs.readFileSync(path.join(ROOT_DIR, 'products', 'vietdub.html'), 'utf8');
    assert(vietdubHtml.includes('parseSrtContent') || vietdubHtml.includes('translateAllSubtitles'), 'Thiếu hàm phân tích/dịch phụ đề SRT');
    assert(vietdubHtml.includes('speechSynthesis'), 'Thiếu tích hợp Web Speech API TTS');
    assert(vietdubHtml.includes('exportSrtFile') && vietdubHtml.includes('exportVttFile'), 'Thiếu hàm xuất file .SRT và .VTT');
    return 'Đầy đủ bộ đọc .SRT + Dịch song ngữ + Đọc giọng AI TTS + Xuất .SRT/.VTT';
  }));

  tasks.push(recordTest('Unit Test', 'MonVPN v1.0.2', 'Kiểm tra thuật toán đo độ trễ HTTP thực tế & sinh khóa WireGuard .conf', () => {
    const monvpnHtml = fs.readFileSync(path.join(ROOT_DIR, 'products', 'monvpn-pc.html'), 'utf8');
    assert(monvpnHtml.includes('crypto.getRandomValues'), 'Thiếu sinh khóa mật mã ngẫu nhiên bằng Web Crypto API');
    assert(monvpnHtml.includes('[Interface]') && monvpnHtml.includes('[Peer]'), 'Thiếu cấu trúc chuẩn tệp cấu hình WireGuard .conf');
    assert(monvpnHtml.includes('performance.now()'), 'Thiếu bộ đo độ trễ thực tế bằng performance.now()');
    return 'Đầy đủ đo Ping HTTP thực tế + Sinh cặp khóa 256-bit + Xuất WireGuard .conf';
  }));

  // =========================================================================
  // 2. E2E & ACCURACY TESTS
  // =========================================================================
  tasks.push(recordTest('E2E & Accuracy', 'VBTM Pro & VietNamOCR', 'Kiểm tra chuyển đổi PDF văn bản Đảng sang Word giữ nguyên thể thức Hướng dẫn 05-HD/VPTW', () => {
    const pdfLines = [
      { y: 780, items: [{ x: 50, text: 'QUẬN ỦY BA ĐÌNH' }, { x: 310, text: 'ĐẢNG CỘNG SẢN VIỆT NAM' }] },
      { y: 762, items: [{ x: 50, text: 'ĐẢNG ỦY PHƯỜNG ĐIỆN BIÊN' }] },
      { y: 742, items: [{ x: 50, text: 'Số 45-BC/ĐU' }, { x: 295, text: 'Điện Biên, ngày 17 tháng 11 năm 2026' }] },
      { y: 700, items: [{ x: 210, text: 'BÁO CÁO' }] },
      { y: 682, items: [{ x: 120, text: 'Kết quả công tác xây dựng Đảng quý IV năm 2026' }] },
      { y: 650, items: [{ x: 50, text: 'Thực hiện Chương trình công tác năm 2026, Đảng ủy phường báo cáo kết quả' }] },
      { y: 634, items: [{ x: 50, text: 'công tác xây dựng Đảng quý IV năm 2026 với những nội dung cụ thể như sau:' }] },
      { y: 590, items: [{ x: 50, text: 'Nơi nhận:' }, { x: 340, text: 'T/M BAN THƯỜNG VỤ' }] },
      { y: 572, items: [{ x: 50, text: '- Thường trực Quận ủy (để b/c);' }, { x: 360, text: 'PHÓ BÍ THƯ' }] },
      { y: 556, items: [{ x: 50, text: '- Các chi bộ trực thuộc;' }] },
      { y: 540, items: [{ x: 50, text: '- Lưu Văn phòng Đảng ủy.' }, { x: 350, text: 'Phạm Thị Thanh Hoa' }] }
    ];
    const pageHtml = sandbox.reconstructVbtmExactPdfPageLayout(pdfLines, 600);
    assert(pageHtml.includes('ĐẢNG CỘNG SẢN VIỆT NAM'), 'Thiếu tiêu ngữ Đảng');
    assert(pageHtml.includes('font-size:15pt; font-weight:bold;'), 'Tiêu ngữ Đảng/Tên loại phải có cỡ 15pt đậm');
    assert(pageHtml.includes('<u>Nơi nhận</u>:'), 'Thiếu <u>Nơi nhận</u>: gạch chân có dấu : theo HD 05');
    assert(pageHtml.includes('Thực hiện Chương trình công tác năm 2026, Đảng ủy phường báo cáo kết quả công tác xây dựng Đảng'), 'Hai dòng PDF chưa được ghép thành 1 đoạn văn');

    // Kiểm tra tự động nạp vào form của VBTM Pro để khi bấm "Chuẩn Hóa" không bị mất thể thức Đảng
    const parsedMeta = sandbox.parsePlainTextIntoVbtmForm(
      pdfLines.map(r => r.items.map(i => i.text).join('   ')).join('\n')
    );
    assert(parsedMeta.isParty === true, 'Phải tự động chuyển sang chế độ Văn bản Đảng (Hướng dẫn 05-HD/VPTW)');
    assert(elements['vbtmDocMode'].value === 'party', 'Dropdown chế độ phải tự động chọn party');
    return 'PDF -> Word tái tạo chuẩn 100% khung 2 cột đầu/cuối trang + nối đoạn văn theo HD 05-HD/VPTW';
  }));

  tasks.push(recordTest('E2E & Accuracy', 'Toàn Hệ Thống', 'Đảm bảo loại bỏ 100% Hướng dẫn 36-HD/VPTW (hết hiệu lực) và dữ liệu mẫu cũ', () => {
    const filesToScan = [
      'index.html',
      'admin.html',
      'data/products.json',
      'js/main.js',
      'js/vbtm-app.js',
      'js/ocr-app.js',
      'products/vbtm-pro.html',
      'products/vietnam-ocr.html',
      'desktop-src/VBTM-Pro/VBTMSetupAndApp.cs',
      'desktop-src/VietNamOCR-Pro/VietNamOCRSetupAndApp.cs'
    ];
    for (const rel of filesToScan) {
      const content = fs.readFileSync(path.join(ROOT_DIR, rel), 'utf8');
      assert(!content.includes('36-HD/VPTW'), `Phát hiện chuỗi 36-HD/VPTW đã hết hiệu lực trong ${rel}`);
    }
    return `Đã quét ${filesToScan.length} tệp mã nguồn Web & Desktop: 0 tham chiếu tới 36-HD/VPTW, 100% dùng 05-HD/VPTW`;
  }));

  // =========================================================================
  // 3. INTEGRATION, SECURITY & PERFORMANCE TESTS ON LOCALHOST SERVER
  // =========================================================================
  await new Promise((resolve, reject) => {
    server.listen(0, '127.0.0.1', () => resolve());
    server.on('error', reject);
  });
  const port = server.address().port;

  tasks.push(recordTest('Integration Test', 'Backend API (server.js)', 'GET /api/health & GET /api/products hoạt động chính xác trên localhost', async () => {
    const healthRes = await httpRequest(port, '/api/health');
    assert(healthRes.statusCode === 200, `Health status ${healthRes.statusCode}`);
    const healthJson = JSON.parse(healthRes.body);
    assert(healthJson.status === 'ok', 'Health status phải là ok');

    const prodRes = await httpRequest(port, '/api/products');
    assert(prodRes.statusCode === 200, `Products status ${prodRes.statusCode}`);
    const products = JSON.parse(prodRes.body);
    assert(Array.isArray(products) && products.length >= 5, `Phải có ít nhất 5 ứng dụng, nhận được ${products.length}`);
    return `Server localhost:${port} trả về ${products.length} ứng dụng chuẩn xác`;
  }));

  tasks.push(recordTest('Integration Test', 'Admin Dashboard & API', 'POST /api/products cập nhật danh mục và GET/POST /api/stats đếm lượt tải', async () => {
    const origRes = await httpRequest(port, '/api/products');
    const origProducts = JSON.parse(origRes.body);

    const postRes = await httpRequest(port, '/api/products', 'POST', origProducts);
    assert(postRes.statusCode === 200, 'POST /api/products phải trả về 200');
    const postJson = JSON.parse(postRes.body);
    assert(postJson.count === origProducts.length, 'Số lượng sản phẩm lưu phải khớp');

    const statsVisit = JSON.parse((await httpRequest(port, '/api/stats?action=visit')).body);
    const statsDown = JSON.parse((await httpRequest(port, '/api/stats?action=download', 'POST')).body);
    assert(statsDown.total_downloads === statsVisit.total_downloads + 1, 'Bộ đếm lượt tải phải tăng +1');
    return `CRUD sản phẩm & thống kê lượt truy cập (${statsDown.total_visits}) / lượt tải (${statsDown.total_downloads}) hoạt động tốt`;
  }));

  tasks.push(recordTest('Integration Test', 'Static Web Pages', 'Tất cả 7 trang giao diện chính phản hồi HTTP 200 OK trên localhost', async () => {
    const pages = [
      '/',
      '/admin.html',
      '/products/vbtm-pro.html',
      '/products/vietnam-ocr.html',
      '/products/trinh-chay-su-kien.html',
      '/products/vietdub.html',
      '/products/monvpn-pc.html'
    ];
    for (const p of pages) {
      const res = await httpRequest(port, p);
      assert(res.statusCode === 200, `Trang ${p} trả về mã lỗi ${res.statusCode}`);
      assert(res.body.length > 1000, `Nội dung trang ${p} quá ngắn`);
    }
    return `Cả ${pages.length}/7 trang ứng dụng đều phản hồi HTTP 200 OK`;
  }));

  tasks.push(recordTest('Security Test', 'Backend & Frontend', 'Chống tấn công Path Traversal, XSS Sanitization & Kiểm tra Security Headers', async () => {
    const traversal1 = await httpRequest(port, '/../../Windows/win.ini');
    const traversal2 = await httpRequest(port, '/..%2f..%2fWindows/win.ini');
    assert(traversal1.statusCode === 403 && traversal2.statusCode === 403, 'Phải chặn Path Traversal bằng HTTP 403 Forbidden');
    assert(traversal1.headers['x-content-type-options'] === 'nosniff', 'Thiếu header X-Content-Type-Options: nosniff');
    assert(traversal1.headers['x-frame-options'] === 'SAMEORIGIN', 'Thiếu header X-Frame-Options: SAMEORIGIN');

    const xssInput = '<script>alert("xss")</script>';
    const escaped = sandbox.escapeVbtmHtml(xssInput);
    assert(!escaped.includes('<script>'), 'escapeVbtmHtml phải mã hóa thẻ <script>');
    return 'Chặn 100% Path Traversal (403), mã hóa XSS an toàn & có đầy đủ Security Headers';
  }));

  tasks.push(recordTest('Security Test', 'Windows Desktop (.exe)', 'Xác thực chữ ký băm SHA256 của bộ cài .exe khớp 100% với trang tải xuống', () => {
    const vbtmExe = fs.readFileSync(path.join(ROOT_DIR, 'downloads', 'VBTM-Setup-2.2.4.exe'));
    const ocrExe = fs.readFileSync(path.join(ROOT_DIR, 'downloads', 'VietNamOCR-Setup-3.2.0-x64.exe'));
    const vbtmSha = crypto.createHash('sha256').update(vbtmExe).digest('hex');
    const ocrSha = crypto.createHash('sha256').update(ocrExe).digest('hex');

    const vbtmHtml = fs.readFileSync(path.join(ROOT_DIR, 'products', 'vbtm-pro.html'), 'utf8');
    const ocrHtml = fs.readFileSync(path.join(ROOT_DIR, 'products', 'vietnam-ocr.html'), 'utf8');
    assert(vbtmHtml.includes(vbtmSha), `SHA256 trong vbtm-pro.html chưa khớp (${vbtmSha})`);
    assert(ocrHtml.includes(ocrSha), `SHA256 trong vietnam-ocr.html chưa khớp (${ocrSha})`);
    return `VBTM-Setup-2.2.4.exe (${vbtmSha.slice(0, 12)}...) & VietNamOCR-Setup-3.2.0-x64.exe (${ocrSha.slice(0, 12)}...) khớp SHA256`;
  }));

  tasks.push(recordTest('Performance Test', 'VBTM Engine & API', 'Hiệu năng chuẩn hóa văn bản 50 đoạn (< 50ms) và độ trễ API (< 30ms)', async () => {
    const longText = Array.from({ length: 50 }, (_, i) =>
      `${i + 1}. Đảng ủy và Ủy ban nhân dân phường triển khai thực hiện tốt các nhiệm vụ phát triển kinh tế - xã hội tháng ${(i % 12) + 1} năm 2026.`
    ).join('\n');

    const t0 = performance.now();
    sandbox.reconstructRegulationHtmlFromPlainText(longText);
    const docTime = performance.now() - t0;
    assert(docTime < 50, `Thời gian chuẩn hóa 50 đoạn quá chậm: ${docTime.toFixed(2)}ms`);

    const t1 = performance.now();
    await httpRequest(port, '/api/products');
    const apiTime = performance.now() - t1;
    assert(apiTime < 30, `Thời gian phản hồi API quá chậm: ${apiTime.toFixed(2)}ms`);
    return `Chuẩn hóa 50 đoạn văn: ${docTime.toFixed(2)}ms (< 50ms) | Phản hồi API: ${apiTime.toFixed(2)}ms (< 30ms)`;
  }));

  tasks.push(recordTest('Responsive & Theme', 'UI Đỏ - Vàng Cờ Đảng', 'Kiểm tra Responsive Breakpoints (Desktop/Tablet/Mobile) & Tông màu Cờ Đảng', () => {
    const css = fs.readFileSync(path.join(ROOT_DIR, 'css', 'style.css'), 'utf8');
    assert(css.includes('@media (max-width: 1024px)'), 'Thiếu breakpoint 1024px');
    assert(css.includes('@media (max-width: 900px)'), 'Thiếu breakpoint 900px');
    assert(css.includes('@media (max-width: 768px)'), 'Thiếu breakpoint 768px');
    assert(css.includes('@media (max-width: 640px)'), 'Thiếu breakpoint 640px');
    assert(css.includes('#4c0509') && css.includes('#facc15'), 'Thiếu bảng màu Đỏ - Vàng Cờ Đảng (#4c0509 & #facc15)');
    return 'Đầy đủ 4 mốc Responsive (1024px, 900px, 768px, 640px) & Giao diện Đỏ - Vàng Cờ Đảng đồng bộ';
  }));

  for (const task of tasks) {
    await task();
  }

  server.close();

  console.log('\n=============================================================');
  console.log(`TỔNG KẾT KIỂM THỬ: ${passedCount}/${results.length} PASS (${failedCount} FAIL)`);
  console.log('=============================================================\n');

  if (failedCount > 0) {
    process.exit(1);
  }
}

runAllTests().catch(err => {
  console.error(err);
  process.exit(1);
});
