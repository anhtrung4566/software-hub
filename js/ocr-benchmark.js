/**
 * VietNamOCR Pro — Bộ kiểm thử chất lượng OCR (benchmark) với tài liệu sinh có đáp án chuẩn (ground truth)
 * Mỗi tài liệu được OCR 2 lần bằng CHÍNH engine của phần mềm:
 *   TRƯỚC = Tesseract thô (rawBaseline: không tiền xử lý thích ứng, không OCR từng ô, không sửa khoảng trắng theo ảnh)
 *   SAU   = pipeline đầy đủ (adaptivePreprocess + OCR từng ô + tách từ dính theo ảnh + phân vùng bố cục)
 * Chỉ số: CER (ký tự, không tính dấu cách), WER, độ chính xác khoảng trắng (P/R/F1 theo căn chỉnh ký tự),
 * phân đoạn dòng, cấu trúc bảng (hàng/cột/ô gộp), nội dung từng ô + lẫn ô, bố cục, số liệu, và kiểm tra DOCX thật.
 */
(function () {
  'use strict';
  const FONT = '"Times New Roman", Times, serif';
  const $ = (id) => document.getElementById(id);
  function setStatus(s) { const el = $('benchStatus'); if (el) el.textContent = s; window.OCR_BENCH_STATUS = s; }

  // ------------------------------------------------------------ chỉ số
  const nfc = (s) => String(s == null ? '' : s).normalize('NFC');
  const normText = (s) => nfc(s).replace(/[|]/g, ' ').replace(/\s+/g, ' ').trim();

  function lev(a, b) {
    const n = a.length, m = b.length;
    if (!n) return m; if (!m) return n;
    let prev = new Uint32Array(m + 1), cur = new Uint32Array(m + 1);
    for (let j = 0; j <= m; j++) prev[j] = j;
    for (let i = 1; i <= n; i++) {
      cur[0] = i;
      for (let j = 1; j <= m; j++) {
        const cost = a[i - 1] === b[j - 1] ? 0 : 1;
        cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
      }
      const t = prev; prev = cur; cur = t;
    }
    return prev[m];
  }
  const sim = (a, b) => 1 - lev(a, b) / Math.max(a.length, b.length, 1);

  function boundarySet(text) {
    const t = normText(text);
    const chars = []; const b = new Set();
    for (let i = 0; i < t.length; i++) {
      if (t[i] === ' ') { if (chars.length) b.add(chars.length); } // b = độ dài tiền tố (số ký tự) đứng trước dấu cách
      else chars.push(t[i]);
    }
    return { ns: chars.join(''), b };
  }

  // căn chỉnh ký tự (bỏ dấu cách) rồi so sánh vị trí dấu cách -> P/R
  function spacingCounts(gtLine, ocLine) {
    const G = boundarySet(gtLine), O = boundarySet(ocLine);
    const n = G.ns.length, m = O.ns.length;
    const D = new Uint16Array((n + 1) * (m + 1));
    const W = m + 1;
    for (let i = 0; i <= n; i++) D[i * W] = i;
    for (let j = 0; j <= m; j++) D[j] = j;
    for (let i = 1; i <= n; i++) for (let j = 1; j <= m; j++) {
      const c = G.ns[i - 1] === O.ns[j - 1] ? 0 : 1;
      D[i * W + j] = Math.min(D[(i - 1) * W + j] + 1, D[i * W + j - 1] + 1, D[(i - 1) * W + j - 1] + c);
    }
    const mp = new Int32Array(m + 1); // mp[k] = số ký tự GT đã tiêu thụ sau khi xử lý k ký tự OCR
    let i = n, j = m;
    while (i > 0 || j > 0) {
      if (i > 0 && j > 0 && D[i * W + j] === D[(i - 1) * W + j - 1] + (G.ns[i - 1] === O.ns[j - 1] ? 0 : 1)) { mp[j] = i; i--; j--; }
      else if (j > 0 && D[i * W + j] === D[i * W + j - 1] + 1) { mp[j] = i; j--; }
      else i--;
    }
    mp[0] = 0;
    const pred = new Set();
    O.b.forEach(k => pred.add(mp[k]));
    let tp = 0; pred.forEach(v => { if (G.b.has(v)) tp++; });
    return { tp, pred: pred.size, gt: G.b.size };
  }

  function evalText(gtLines, ocLines) {
    const g = gtLines.map(normText).filter(Boolean);
    const o = ocLines.map(normText).filter(Boolean);
    const used = new Array(o.length).fill(false);
    const claimed = new Set();
    let charErr = 0, charTot = 0, wordErr = 0, wordTot = 0, tp = 0, pred = 0, gtB = 0, matched = 0;
    g.forEach(gl => {
      const gn = gl.replace(/ /g, '');
      let best = -1, bestSim = -1, bestD = 0;
      o.forEach((ol, j) => {
        const on = ol.replace(/ /g, '');
        const d = lev(gn, on);
        const s = 1 - d / Math.max(gn.length, on.length, 1);
        if (s > bestSim) { bestSim = s; best = j; bestD = d; }
      });
      const gw = gl.split(' ');
      charTot += gn.length; wordTot += gw.length;
      if (best < 0) { charErr += gn.length; wordErr += gw.length; gtB += gw.length - 1; return; }
      used[best] = true;
      charErr += bestD;
      wordErr += lev(gw, o[best].split(' '));
      if (bestSim >= 0.5) { const sc = spacingCounts(gl, o[best]); tp += sc.tp; pred += sc.pred; gtB += sc.gt; }
      else gtB += gw.length - 1;
      if (bestSim >= 0.8 && !claimed.has(best)) { matched++; claimed.add(best); }
    });
    let extraC = 0, extraW = 0;
    o.forEach((ol, j) => { if (!used[j]) { extraC += ol.replace(/ /g, '').length; extraW += ol.split(' ').length; } });
    const P = pred ? tp / pred : (gtB ? 0 : 1), R = gtB ? tp / gtB : 1;
    const gnum = (g.join(' ').match(/\d+(?:[\/.:,-]\d+)*/g) || []).map(x => x.replace(/[.,:-]+$/, ''));
    const onum = (o.join(' ').match(/\d+(?:[\/.:,-]\d+)*/g) || []).map(x => x.replace(/[.,:-]+$/, ''));
    const bag = new Map(); onum.forEach(x => bag.set(x, (bag.get(x) || 0) + 1));
    let numOk = 0; gnum.forEach(x => { if (bag.get(x) > 0) { numOk++; bag.set(x, bag.get(x) - 1); } });
    return {
      cer: (charErr + extraC) / Math.max(1, charTot),
      wer: (wordErr + extraW) / Math.max(1, wordTot),
      spaceP: P, spaceR: R, spaceF1: (P + R) ? 2 * P * R / (P + R) : 0,
      seg: matched / Math.max(1, g.length),
      lineCount: { gt: g.length, ocr: o.length },
      numeric: gnum.length ? numOk / gnum.length : 1,
      numericTotal: gnum.length
    };
  }

  function evalTable(gt, found, grid) {
    const out = { rowsOk: false, colsOk: false, spanAcc: 0, cellExact: 0, cellFuzzy: 0, contamination: 0, cells: gt.cells.length };
    if (!grid) return out;
    out.rowsOk = grid.numRows === gt.rows; out.colsOk = grid.numCols === gt.cols;
    const keys = new Set((grid.cells || []).map(c => `${c.r},${c.c},${c.rowspan},${c.colspan}`));
    let spanOk = 0;
    gt.cells.forEach(c => { if (keys.has(`${c.r},${c.c},${c.rs || 1},${c.cs || 1}`)) spanOk++; });
    out.spanAcc = spanOk / gt.cells.length;
    const map = new Map(); found.forEach(f => map.set(`${f.r},${f.c}`, normText(f.text)));
    const gtNorm = gt.cells.map(c => ({ key: `${c.r},${c.c}`, text: normText(c.lines.join(' ')) }));
    let ex = 0, fz = 0;
    gtNorm.forEach(c => {
      const f = map.get(c.key) || '';
      if (f === c.text) ex++;
      if (sim(f, c.text) >= 0.9) fz++;
      gtNorm.forEach(o2 => {
        if (o2.key !== c.key && o2.text.length >= 10 && f.includes(o2.text.slice(0, 10)) && !c.text.includes(o2.text.slice(0, 10))) out.contamination++;
      });
    });
    out.cellExact = ex / gtNorm.length; out.cellFuzzy = fz / gtNorm.length;
    return out;
  }

  // phương pháp CŨ: gán từ mức trang vào ô theo tâm, gom dòng theo y, nối bằng 1 dấu cách
  function oldCellAssignment(grid, words) {
    const res = grid.cells.map(c => ({ r: c.r, c: c.c, words: [], x0: c.x0, y0: c.y0, x1: c.x1, y1: c.y1 }));
    (words || []).forEach(w => {
      if (!w.bbox || !String(w.text || '').trim()) return;
      const cx = (w.bbox.x0 + w.bbox.x1) / 2, cy = (w.bbox.y0 + w.bbox.y1) / 2;
      for (const c of res) if (cx >= c.x0 - 3 && cx <= c.x1 + 3 && cy >= c.y0 - 3 && cy <= c.y1 + 3) { c.words.push(w); break; }
    });
    return res.map(c => {
      const sorted = c.words.slice().sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
      const lines = [];
      sorted.forEach(w => {
        const cy = (w.bbox.y0 + w.bbox.y1) / 2, h = Math.max(8, w.bbox.y1 - w.bbox.y0);
        let row = lines.find(l => Math.abs(l.cy - cy) <= h * 0.65);
        if (!row) { row = { cy, words: [] }; lines.push(row); }
        row.words.push(w);
      });
      const text = lines.map(l => l.words.sort((a, b) => a.bbox.x0 - b.bbox.x0).map(w => String(w.text).trim()).join(' ')).join(' ');
      return { r: c.r, c: c.c, text };
    });
  }

  // ------------------------------------------------------------ vẽ tài liệu thử
  function mk(w, h) {
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const x = c.getContext('2d', { willReadFrequently: true });
    x.fillStyle = '#fff'; x.fillRect(0, 0, w, h); x.fillStyle = '#000';
    return { c, x, gt: [], table: null, expect: { types: [], columns: null }, landscape: w > h };
  }
  function T(d, text, px, py, size, o = {}) {
    d.x.font = `${o.italic ? 'italic ' : ''}${o.bold ? 'bold ' : ''}${size}px ${FONT}`;
    d.x.textAlign = o.align || 'left';
    d.x.fillStyle = '#000';
    d.x.fillText(text, px, py);
    if (!o.noGt) d.gt.push(text);
  }
  function drawCells(d, xs, ys, cells, size, bold) {
    d.x.strokeStyle = '#000'; d.x.lineWidth = 2;
    cells.forEach(c => {
      const x0 = xs[c.c], y0 = ys[c.r], x1 = xs[c.c + (c.cs || 1)], y1 = ys[c.r + (c.rs || 1)];
      d.x.strokeRect(x0, y0, x1 - x0, y1 - y0);
      const lines = c.lines || [];
      const lh = size * 1.45;
      const isHead = !!c.head;
      const blockH = lines.length * lh;
      let y = y0 + ((y1 - y0) - blockH) / 2 + size * 1.05;
      lines.forEach((ln, i) => {
        if (isHead || c.center) T(d, ln, (x0 + x1) / 2, y + i * lh, size, { bold: isHead || (c.bold && i === 0), align: 'center' });
        else T(d, ln, x0 + 18, y + i * lh, size, { bold: c.bold && i === 0 });
      });
    });
  }

  function docLich() {
    const d = mk(2100, 1500);
    T(d, 'ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT', 110, 90, 30, { bold: true });
    T(d, 'VĂN PHÒNG', 110, 135, 30, { bold: true });
    T(d, 'ĐẢNG CỘNG SẢN VIỆT NAM', 1990, 90, 30, { bold: true, align: 'right' });
    T(d, 'Tân Sơn Nhất, ngày 05 tháng 9 năm 2026', 1990, 135, 28, { italic: true, align: 'right' });
    T(d, 'LỊCH CÔNG TÁC TUẦN', 1050, 270, 52, { bold: true, align: 'center' });
    T(d, 'PHƯỜNG TÂN SƠN NHẤT', 1050, 335, 42, { bold: true, align: 'center' });
    T(d, 'Tuần 36 năm 2026', 1050, 395, 36, { align: 'center' });
    T(d, 'từ ngày 07/9/2026 đến ngày 13/9/2026', 1050, 445, 34, { italic: true, align: 'center' });
    const xs = [100, 480, 1290, 2000], ys = [500, 570, 700, 830, 960, 1090, 1220, 1350];
    const rows = [
      [['Thứ Hai', '07/9/2026'], ['Họp giao ban Đảng ủy phường', 'Tại phòng họp tầng 2'], ['Kiểm tra công tác phòng cháy chữa cháy']],
      [['Thứ Ba', '08/9/2026'], ['Tiếp công dân định kỳ'], ['Làm việc với Ban Chấp hành Mặt trận Tổ quốc']],
      [['Thứ Tư', '09/9/2026'], ['Dự hội nghị tổng kết quý III năm 2026'], []],
      [['Thứ Năm', '10/9/2026'], ['Kiểm tra vệ sinh môi trường các tổ dân phố'], ['Họp Hội đồng nhân dân phường']],
      [['Thứ Sáu', '11/9/2026'], ['Giải quyết thủ tục hành chính cho nhân dân'], ['Sinh hoạt chi bộ định kỳ tháng 9']],
      [['Thứ Bảy', '12/9/2026'], ['Trực chỉ huy phòng chống thiên tai'], ['Trực cơ quan']]
    ];
    const cells = [
      { r: 0, c: 0, lines: ['NGÀY'], head: true }, { r: 0, c: 1, lines: ['SÁNG'], head: true }, { r: 0, c: 2, lines: ['CHIỀU'], head: true }
    ];
    rows.forEach((row, i) => row.forEach((ls, c) => cells.push({ r: i + 1, c, lines: ls, bold: c === 0, center: c === 0 })));
    drawCells(d, xs, ys, cells, 30);
    d.table = { rows: 7, cols: 3, cells };
    d.expect = { types: ['header', 'title', 'table'], columns: null };
    d.docxExpect = { tbl: 1, tr: 7, tc: 21 };
    return d;
  }

  function docCongVan(variant) {
    const d = mk(1500, 2120);
    T(d, 'ỦY BAN NHÂN DÂN', 100, 100, 30, { bold: true });
    T(d, 'PHƯỜNG TÂN SƠN NHẤT', 100, 145, 30, { bold: true });
    T(d, 'Số: 125/UBND-VP', 100, 240, 28);
    T(d, 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 1400, 100, 28, { bold: true, align: 'right' });
    T(d, 'Độc lập - Tự do - Hạnh phúc', 1400, 145, 30, { bold: true, align: 'right' });
    T(d, 'Tân Sơn Nhất, ngày 05 tháng 9 năm 2026', 1400, 200, 28, { italic: true, align: 'right' });
    T(d, 'V/v tổ chức hội nghị tổng kết công tác quý III năm 2026', 750, 340, 32, { italic: true, align: 'center' });
    T(d, 'Kính gửi: Các tổ dân phố trên địa bàn phường', 200, 420, 34);
    const body = variant === 'b' ? [
      ['Căn cứ Nghị quyết số 12/NQ-HĐND ngày 20/7/2026 của Hội đồng nhân dân', 'phường về nhiệm vụ phát triển kinh tế - xã hội năm 2026, Ủy ban nhân dân', 'phường yêu cầu các tổ dân phố thực hiện các nội dung sau đây:'],
      ['1. Rà soát toàn bộ hộ nghèo, hộ cận nghèo trước ngày 25/9/2026.', '2. Báo cáo kết quả về Văn phòng Ủy ban nhân dân phường.', '3. Bảo đảm an ninh trật tự trong dịp Quốc khánh 02/9.'],
      ['Đề nghị các đồng chí tổ trưởng nghiêm túc triển khai thực hiện;', 'mọi vướng mắc phản ánh về Văn phòng để được hướng dẫn kịp thời.']
    ] : [
      ['Thực hiện Kế hoạch số 36/KH-UBND ngày 28/8/2026 của Ủy ban nhân dân', 'phường về công tác tổng kết quý III năm 2026, Ủy ban nhân dân phường', 'tổ chức hội nghị tổng kết với các nội dung chính như sau:'],
      ['1. Thời gian: 08 giờ 00 phút, ngày 15/9/2026 (thứ Ba).', '2. Địa điểm: Hội trường Ủy ban nhân dân phường, số 45 đường Trường Sơn.', '3. Thành phần: Lãnh đạo phường, trưởng các tổ dân phố và cán bộ, công chức.'],
      ['Đề nghị các đồng chí chuẩn bị báo cáo, bảo đảm đúng thời gian quy định;', 'tuyệt đối không vắng mặt khi chưa được sự đồng ý của Chủ tịch.']
    ];
    let y = 510;
    body.forEach(par => { par.forEach((ln, i) => { T(d, ln, i === 0 && !/^\d\./.test(ln) ? 200 : 100, y, 34); y += 58; }); y += 40; });
    T(d, 'Nơi nhận:', 100, 1750, 28, { bold: true, italic: true });
    T(d, '- Như trên;', 100, 1795, 28);
    T(d, '- Đảng ủy phường (để báo cáo);', 100, 1840, 28);
    T(d, '- Lưu: VT, VP.', 100, 1885, 28);
    T(d, 'TM. ỦY BAN NHÂN DÂN', 1100, 1750, 30, { bold: true, align: 'center' });
    T(d, 'CHỦ TỊCH', 1100, 1795, 30, { bold: true, align: 'center' });
    T(d, 'Nguyễn Văn An', 1100, 1990, 32, { bold: true, align: 'center' });
    d.expect = { types: ['header', 'list', 'footer'], columns: null };
    return d;
  }

  function docBangGop() {
    const d = mk(2000, 1500);
    T(d, 'BẢNG THỐNG KÊ KẾT QUẢ THỰC HIỆN NHIỆM VỤ QUÝ III', 1000, 130, 44, { bold: true, align: 'center' });
    const xs = [100, 380, 1100, 1450, 1900], ys = [250, 330, 410, 510, 610, 710, 810, 910];
    const cells = [
      { r: 0, c: 0, rs: 2, lines: ['STT'], head: true }, { r: 0, c: 1, rs: 2, lines: ['NỘI DUNG CÔNG VIỆC'], head: true },
      { r: 0, c: 2, cs: 2, lines: ['KẾT QUẢ'], head: true },
      { r: 1, c: 2, lines: ['Số lượng'], head: true }, { r: 1, c: 3, lines: ['Tỷ lệ'], head: true }
    ];
    const data = [['1', 'Tuyên truyền pháp luật', '125', '100%'], ['2', 'Kiểm tra an toàn thực phẩm', '48', '96,5%'], ['3', 'Giải quyết hồ sơ hành chính', '1.250', '98,2%'], ['4', 'Hòa giải cơ sở', '36', '92%'], ['5', 'Vận động quỹ vì người nghèo', '85.000.000', '105,3%']];
    data.forEach((row, i) => row.forEach((t, c) => cells.push({ r: i + 2, c, lines: [t], center: c !== 1 })));
    drawCells(d, xs, ys, cells, 32);
    d.table = { rows: 7, cols: 4, cells };
    d.expect = { types: ['title', 'table'], columns: null };
    d.docxExpect = { tbl: 1, tr: 7, tc: 27 };
    return d;
  }

  function docHaiCot() {
    const d = mk(2000, 1400);
    const L = ['Công tác chuyển đổi số trên địa bàn', 'phường đã đạt nhiều kết quả quan trọng.', 'Hệ thống một cửa điện tử tiếp nhận', 'hơn 1.200 hồ sơ trong quý III năm', '2026, tỷ lệ giải quyết đúng hạn đạt', '98,5%. Cán bộ, công chức được tập', 'huấn kỹ năng số định kỳ hằng tháng.', 'Ủy ban nhân dân phường tiếp tục đẩy', 'mạnh ứng dụng công nghệ thông tin', 'trong quản lý, điều hành và phục vụ', 'nhân dân, bảo đảm minh bạch, hiệu quả.'];
    const R = ['Về công tác an ninh trật tự, lực lượng', 'công an phường phối hợp với các tổ dân', 'phố tuần tra, kiểm soát thường xuyên.', 'Trong quý đã giải quyết 36 vụ việc, hòa', 'giải thành 31 vụ, đạt tỷ lệ 86,1%. Công', 'tác phòng cháy chữa cháy được kiểm tra', 'tại 125 cơ sở, kịp thời nhắc nhở các', 'trường hợp chưa bảo đảm an toàn theo', 'quy định hiện hành của pháp luật.'];
    L.forEach((t, i) => T(d, t, 100, 140 + i * 62, 34));
    R.forEach((t, i) => T(d, t, 1060, 140 + i * 62, 34));
    d.expect = { types: ['paragraph'], columns: 2 };
    return d;
  }

  function docChuNho() {
    const d = mk(1500, 2120);
    for (let k = 1; k <= 24; k++) {
      T(d, `Điều ${k}. Ủy ban nhân dân phường tổ chức thực hiện quy định tại Nghị định số ${30 + k}/2020/NĐ-CP ngày ${k + 2}/3/2026.`, 90, 110 + k * 36, 22);
    }
    d.expect = { types: ['paragraph'], columns: null };
    return d;
  }

  function docTrang(n) {
    const d = mk(1500, 2120);
    const base = n === 1
      ? ['Báo cáo kết quả thực hiện nhiệm vụ quý III năm 2026', 'Trong quý III năm 2026, Đảng ủy phường đã lãnh đạo, chỉ đạo', 'triển khai đồng bộ các nhiệm vụ chính trị, bảo đảm quốc phòng,', 'an ninh và phát triển kinh tế - xã hội trên địa bàn.', 'Tổng thu ngân sách đạt 12.450 triệu đồng, bằng 104,2% kế hoạch.', 'Đã kết nạp 03 đảng viên mới; hoàn thành 100% chỉ tiêu giao.']
      : ['Phương hướng nhiệm vụ quý IV năm 2026', 'Tập trung lãnh đạo hoàn thành các chỉ tiêu kế hoạch năm 2026,', 'tăng cường công tác kiểm tra, giám sát và kỷ luật của Đảng.', 'Chuẩn bị tốt các điều kiện tổ chức đại hội chi bộ nhiệm kỳ mới,', 'bảo đảm đoàn kết, thống nhất và nâng cao chất lượng sinh hoạt.', 'Phấn đấu đến ngày 31/12/2026 hoàn thành 100% chương trình.'];
    base.forEach((t, i) => T(d, t, 120, 200 + i * 70, 36, { bold: i === 0 }));
    T(d, `- ${n} -`, 750, 2040, 30, { align: 'center' });
    d.expect = { types: ['page-number'], columns: null };
    return d;
  }

  // làm xấu: nghiêng, mờ, nền xám loang, nhiễu hạt
  function degrade(src) {
    const w = src.c.width, h = src.c.height;
    const out = mk(w, h);
    let seed = 20260926;
    const rnd = () => { seed |= 0; seed = seed + 0x6D2B79F5 | 0; let t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; };
    const g = out.x.createLinearGradient(0, 0, w, h);
    g.addColorStop(0, '#e2e0d6'); g.addColorStop(1, '#b9b7ad');
    out.x.fillStyle = g; out.x.fillRect(0, 0, w, h);
    out.x.save();
    out.x.translate(w / 2, h / 2); out.x.rotate(1.2 * Math.PI / 180);
    out.x.filter = 'blur(1.1px)'; out.x.globalCompositeOperation = 'multiply';
    out.x.drawImage(src.c, -w / 2, -h / 2);
    out.x.restore();
    const id = out.x.getImageData(0, 0, w, h), p = id.data;
    for (let i = 0; i < p.length; i += 4) {
      const n = (rnd() - 0.5) * 36;
      p[i] = Math.max(0, Math.min(255, p[i] + n)); p[i + 1] = Math.max(0, Math.min(255, p[i + 1] + n)); p[i + 2] = Math.max(0, Math.min(255, p[i + 2] + n));
      if (rnd() < 0.0012) { p[i] = p[i + 1] = p[i + 2] = 50; }
    }
    out.x.putImageData(id, 0, 0);
    out.gt = src.gt.slice(); out.expect = src.expect; out.landscape = src.landscape;
    return out;
  }

  function buildDocs() {
    const scan = docCongVan('a');
    return [
      { id: 'D1', name: 'Lịch công tác tuần (A4 ngang, bảng NGÀY|SÁNG|CHIỀU, có ô trống)', doc: docLich() },
      { id: 'D2', name: 'Công văn hành chính (header 2 cột, danh sách, nơi nhận, chữ ký)', doc: docCongVan('a') },
      { id: 'D3', name: 'Bảng nhiều hàng-cột có ô gộp ngang + dọc', doc: docBangGop() },
      { id: 'D4', name: 'Văn bản 2 cột', doc: docHaiCot() },
      { id: 'D5', name: 'Công văn scan mờ + nghiêng 1,2° + nền xám + nhiễu hạt', doc: degrade(scan) },
      { id: 'D6', name: 'Chữ nhỏ dày đặc (22px)', doc: docChuNho() },
      { id: 'D7a', name: 'Nhiều trang — trang 1', doc: docTrang(1) },
      { id: 'D7b', name: 'Nhiều trang — trang 2', doc: docTrang(2) }
    ];
  }

  // ------------------------------------------------------------ DOCX
  async function checkDocx(res, doc) {
    const out = { ok: false };
    try {
      const TG = window.VietNamOcrTableGrid, EX = window.VietNamDocxExporter;
      const html = TG.reconstructPageWithRealTables(res.procCanvas, res.words, res.lines, 'auto');
      if (!html || !html.html) { out.error = 'không dựng được bảng HTML'; return out; }
      const blob = await EX.buildDocxBlob({ html: html.html, landscape: doc.landscape });
      const zip = await window.JSZip.loadAsync(blob);
      const xml = await zip.file('word/document.xml').async('string');
      const dom = new DOMParser().parseFromString(xml, 'application/xml');
      const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const cnt = (n) => dom.getElementsByTagNameNS(W, n).length;
      out.zip = true; out.parseError = !!dom.getElementsByTagName('parsererror').length;
      out.tbl = cnt('tbl'); out.tr = cnt('tr'); out.tc = cnt('tc'); out.gridSpan = cnt('gridSpan'); out.vMerge = cnt('vMerge');
      out.cellText = Array.from(dom.getElementsByTagNameNS(W, 'tc')).map(tc => Array.from(tc.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join(' '));
      const ex = doc.docxExpect || {};
      out.structureOk = out.tbl === ex.tbl && out.tr === ex.tr && out.tc === ex.tc;
      // so khớp nội dung từng <w:tc> với đáp án (các ô có nội dung)
      const gtCells = doc.table.cells.map(c => normText(c.lines.join(' '))).filter(Boolean);
      const docCells = out.cellText.map(normText);
      let hit = 0; const pool = docCells.slice();
      gtCells.forEach(g => { const i = pool.findIndex(x => x === g); if (i >= 0) { hit++; pool.splice(i, 1); } });
      out.cellContentExact = gtCells.length ? hit / gtCells.length : 1;
      out.ok = out.zip && !out.parseError && out.structureOk;
      delete out.cellText;
    } catch (e) { out.error = String(e && e.message || e); }
    return out;
  }

  // ------------------------------------------------------------ chạy
  function typesOf(layout) { return layout && layout.regions ? Array.from(new Set(layout.regions.map(r => r.type))) : []; }

  async function runDoc(item) {
    const EP = window.VietNamOcrEnginePro, TG = window.VietNamOcrTableGrid;
    const d = item.doc;
    const r = { id: item.id, name: item.name, w: d.c.width, h: d.c.height };
    setStatus(`${item.id}: OCR thô (TRƯỚC)…`);
    const before = await EP.recognizeCanvasWithBestEngine(d.c, { lang: 'vie', deskew: true, rawBaseline: true });
    setStatus(`${item.id}: OCR pipeline đầy đủ (SAU)…`);
    const t0 = performance.now();
    const after = await EP.recognizeCanvasWithBestEngine(d.c, { lang: 'vie', deskew: true });
    r.afterMs = Math.round(performance.now() - t0);
    const bLines = (before.lines || []).map(l => l.text || '');
    const aLines = ((after.layout && after.layout.lines) || after.lines || []).map(l => l.text || '');
    r.before = evalText(d.gt, bLines);
    r.after = evalText(d.gt, aLines);
    r.beforeConf = before.confidence; r.afterConf = after.confidence;
    r.preprocess = (after.preprocess && after.preprocess.applied) || [];
    r.deskewDeg = after.deskewAngle;
    if (after.layout) {
      r.repairs = after.layout.repairs.map(x => ({ from: x.original, to: x.proposed }));
      r.suggestions = after.layout.suggestions.map(s => ({ kind: s.kind, from: s.original, to: s.proposed, gain: s.confGain }));
      r.regions = after.layout.regions.map(x => x.type);
      r.columns = after.layout.columns;
      r.flagged = after.layout.flags.length;
      const exp = d.expect.types || [];
      const found = typesOf(after.layout);
      const rec = exp.length ? exp.filter(t => found.includes(t)).length / exp.length : 1;
      const colOk = d.expect.columns ? (after.layout.columns === d.expect.columns ? 1 : 0) : null;
      r.layoutScore = colOk == null ? rec : (rec + colOk) / 2;
    }
    if (d.table) {
      const gb = TG.detectTableGridFromCanvas(before.procCanvas);
      r.tableBefore = evalTable(d.table, gb ? oldCellAssignment(gb, before.words) : [], gb);
      const ga = after.layout && after.layout.grid;
      const found = after.layout && after.layout.cellRecs ? after.layout.cellRecs.map(c => ({ r: c.r, c: c.c, text: c.lines.map(l => l.text).join(' ') })) : [];
      r.tableAfter = evalTable(d.table, found, ga);
      if (d.docxExpect) r.docx = await checkDocx(after, d);
    }
    return r;
  }

  function avg(arr) { const v = arr.filter(x => typeof x === 'number' && !isNaN(x)); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; }
  const pct = (x) => (x == null ? '—' : (x * 100).toFixed(1) + '%');

  function summarize(results) {
    const S = (k, side) => avg(results.map(r => r[side] && r[side][k]));
    const sum = { before: {}, after: {} };
    ['cer', 'wer', 'spaceP', 'spaceR', 'spaceF1', 'seg', 'numeric'].forEach(k => { sum.before[k] = S(k, 'before'); sum.after[k] = S(k, 'after'); });
    const T = (k, side) => avg(results.filter(r => r[side]).map(r => { const t = r[side]; return typeof t[k] === 'boolean' ? (t[k] ? 1 : 0) : t[k]; }));
    ['rowsOk', 'colsOk', 'spanAcc', 'cellExact', 'cellFuzzy', 'contamination'].forEach(k => { sum.before['t_' + k] = T(k, 'tableBefore'); sum.after['t_' + k] = T(k, 'tableAfter'); });
    sum.after.layout = avg(results.map(r => r.layoutScore));
    const dx = results.filter(r => r.docx);
    sum.after.docxOk = dx.length ? dx.filter(r => r.docx.ok).length / dx.length : null;
    sum.after.docxCellContent = avg(dx.map(r => r.docx.cellContentExact));
    return sum;
  }

  function render(results, sum) {
    const row = (label, b, a, lowerBetter, isCount) => {
      const f = (x) => (x == null ? '—' : (isCount ? x.toFixed(1) : pct(x)));
      let cls = '';
      if (b != null && a != null) cls = (lowerBetter ? a < b : a > b) ? 'good' : (a === b ? '' : 'bad');
      return `<tr><td>${label}</td><td>${f(b)}</td><td class="${cls}">${f(a)}</td></tr>`;
    };
    let h = '<h2>Tổng hợp trước / sau (trung bình trên các tài liệu thử)</h2><table><tr><th>Chỉ số</th><th>TRƯỚC (Tesseract thô)</th><th>SAU (pipeline mới)</th></tr>';
    h += row('Sai số ký tự CER (thấp hơn = tốt)', sum.before.cer, sum.after.cer, true);
    h += row('Sai số từ WER (thấp hơn = tốt)', sum.before.wer, sum.after.wer, true);
    h += row('Khoảng trắng — Precision', sum.before.spaceP, sum.after.spaceP);
    h += row('Khoảng trắng — Recall', sum.before.spaceR, sum.after.spaceR);
    h += row('Khoảng trắng — F1', sum.before.spaceF1, sum.after.spaceF1);
    h += row('Phân đoạn dòng (dòng đúng)', sum.before.seg, sum.after.seg);
    h += row('Số liệu / ngày tháng đúng', sum.before.numeric, sum.after.numeric);
    h += row('Bảng: đúng số hàng', sum.before.t_rowsOk, sum.after.t_rowsOk);
    h += row('Bảng: đúng số cột', sum.before.t_colsOk, sum.after.t_colsOk);
    h += row('Bảng: đúng ô gộp (rowspan/colspan)', sum.before.t_spanAcc, sum.after.t_spanAcc);
    h += row('Bảng: nội dung ô khớp 100%', sum.before.t_cellExact, sum.after.t_cellExact);
    h += row('Bảng: nội dung ô khớp ≥ 90%', sum.before.t_cellFuzzy, sum.after.t_cellFuzzy);
    h += row('Bảng: số lần lẫn nội dung giữa ô (thấp hơn = tốt)', sum.before.t_contamination, sum.after.t_contamination, true, true);
    h += row('Phân vùng bố cục (trước: không có)', null, sum.after.layout);
    h += row('DOCX thật: cấu trúc bảng đúng', null, sum.after.docxOk);
    h += row('DOCX thật: nội dung từng ô khớp', null, sum.after.docxCellContent);
    h += '</table><h2>Chi tiết từng tài liệu</h2>';
    results.forEach(r => {
      h += `<div class="doc"><h3>${r.id} — ${r.name} <small>(${r.w}×${r.h})</small></h3>`;
      h += `<table><tr><th></th><th>CER</th><th>WER</th><th>Space F1</th><th>Phân đoạn</th><th>Số liệu</th><th>Tin cậy engine</th></tr>`;
      h += `<tr><td>Trước</td><td>${pct(r.before.cer)}</td><td>${pct(r.before.wer)}</td><td>${pct(r.before.spaceF1)}</td><td>${pct(r.before.seg)}</td><td>${pct(r.before.numeric)}</td><td>${r.beforeConf}%</td></tr>`;
      h += `<tr><td>Sau</td><td>${pct(r.after.cer)}</td><td>${pct(r.after.wer)}</td><td>${pct(r.after.spaceF1)}</td><td>${pct(r.after.seg)}</td><td>${pct(r.after.numeric)}</td><td>${r.afterConf}%</td></tr></table>`;
      h += `<p>Tiền xử lý: ${r.preprocess.length ? r.preprocess.join('; ') : 'không cần'} · deskew ${r.deskewDeg}° · Phân vùng: ${(r.regions || []).join(', ') || '—'}${r.columns > 1 ? ' · ' + r.columns + ' cột' : ''} · Thời gian pipeline ${r.afterMs}ms</p>`;
      if (r.tableAfter) h += `<p>Bảng — trước: hàng ${r.tableBefore.rowsOk ? '✓' : '✗'} cột ${r.tableBefore.colsOk ? '✓' : '✗'} ô khớp ${pct(r.tableBefore.cellExact)} lẫn ô ${r.tableBefore.contamination}; sau: hàng ${r.tableAfter.rowsOk ? '✓' : '✗'} cột ${r.tableAfter.colsOk ? '✓' : '✗'} ô gộp ${pct(r.tableAfter.spanAcc)} ô khớp ${pct(r.tableAfter.cellExact)} lẫn ô ${r.tableAfter.contamination}</p>`;
      if (r.docx) h += `<p>DOCX: ${r.docx.error ? 'LỖI ' + r.docx.error : `zip ✓ · tbl=${r.docx.tbl} tr=${r.docx.tr} tc=${r.docx.tc} gridSpan=${r.docx.gridSpan} vMerge=${r.docx.vMerge} · cấu trúc ${r.docx.structureOk ? '✓' : '✗'} · nội dung ô ${pct(r.docx.cellContentExact)}`}</p>`;
      if (r.repairs && r.repairs.length) h += `<p>Tách từ dính theo ảnh (${r.repairs.length}): ${r.repairs.slice(0, 6).map(x => `“${x.from}” → “${x.to}”`).join('; ')}</p>`;
      if (r.suggestions && r.suggestions.length) h += `<p>Đề xuất đối chiếu ảnh (${r.suggestions.length}, chưa tự áp dụng): ${r.suggestions.slice(0, 4).map(x => `“${x.from}” → “${x.to}” (+${x.gain}%)`).join('; ')}</p>`;
      h += '</div>';
    });
    $('benchResult').innerHTML = h;
  }

  async function runAll() {
    $('benchRun').disabled = true;
    const docs = buildDocs();
    const prev = $('benchPreview'); prev.innerHTML = '';
    docs.forEach(it => { const c = it.doc.c; c.style.width = '220px'; c.style.border = '1px solid #999'; c.title = it.name; prev.appendChild(c); });
    const results = [];
    try {
      for (const it of docs) results.push(await runDoc(it));
      const sum = summarize(results);
      render(results, sum);
      window.__OCR_BENCH__ = { done: true, summary: sum, docs: results };
      $('benchJson').textContent = JSON.stringify(window.__OCR_BENCH__, null, 1);
      setStatus('HOÀN TẤT — ' + results.length + ' tài liệu');
      document.title = 'OCR-BENCH-DONE';
    } catch (e) {
      window.__OCR_BENCH__ = { done: false, error: String(e && e.stack || e), partial: results };
      $('benchJson').textContent = JSON.stringify(window.__OCR_BENCH__, null, 1);
      setStatus('LỖI: ' + (e && e.message || e));
      document.title = 'OCR-BENCH-ERROR';
    }
    $('benchRun').disabled = false;
  }

  window.addEventListener('DOMContentLoaded', () => {
    $('benchRun').addEventListener('click', runAll);
    if (/[?&]run=1/.test(location.search)) setTimeout(runAll, 300);
  });
})();
