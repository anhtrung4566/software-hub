/**
 * VietNamOCR Pro — Lõi bố cục & hiệu chỉnh OCR dựa trên HÌNH ẢNH GỐC (ocr-layout.js)
 *
 * Quy trình (mỗi trang):
 *   0. adaptivePreprocess : đo nền/bóng, nhiễu hạt, chiều cao chữ -> chọn phép xử lý riêng cho từng trang
 *   1. buildInkMap        : bản đồ điểm mực (Otsu) làm "bằng chứng ảnh" cho mọi bước sửa
 *   2. phát hiện lưới bảng -> OCR RIÊNG TỪNG Ô (không bao giờ trộn nội dung 2 ô)
 *   3. repairWordSpacing  : (loại B) tách từ dính dựa trên bbox từng KÝ TỰ + khoảng trắng thật trong ảnh
 *   4. reOcrSuspectLines  : (loại A) cắt đúng vùng ảnh dòng nghi ngờ, phóng to, OCR lại, so sánh -> chỉ ĐỀ XUẤT
 *   5. analyzeRegions     : phân vùng tiêu đề / header / bảng / danh sách / ghi chú / footer / số trang / cột
 *   6. composePageText    : dựng văn bản theo thứ tự đọc, bảng giữ cấu trúc hàng-cột
 *
 * Nguyên tắc: KHÔNG thay chuỗi cố định, KHÔNG chèn dấu cách theo cảm tính, KHÔNG viết lại câu.
 * Mọi thay đổi đều có bằng chứng ảnh + lưu lại nội dung OCR gốc để người dùng xem lại / hoàn tác.
 */
(function (global) {
  'use strict';

  const MAX_CELL_OCR = 120;
  const MAX_REOCR_LINES = 24;

  // ---------------------------------------------------------------- tiện ích
  function nfc(s) { return String(s == null ? '' : s).normalize('NFC'); }
  function median(arr) {
    if (!arr || !arr.length) return 0;
    const a = arr.slice().sort((x, y) => x - y);
    const m = a.length >> 1;
    return a.length % 2 ? a[m] : (a[m - 1] + a[m]) / 2;
  }
  function mean(arr) { return arr.length ? arr.reduce((s, v) => s + v, 0) / arr.length : 0; }
  function ctx2d(canvas) { return canvas.getContext('2d', { willReadFrequently: true }); }
  function newCanvas(w, h) {
    const c = document.createElement('canvas');
    c.width = Math.max(1, Math.round(w));
    c.height = Math.max(1, Math.round(h));
    return c;
  }
  function unionBBox(items) {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    for (const it of items) {
      const b = it.bbox || it;
      if (!b) continue;
      x0 = Math.min(x0, b.x0); y0 = Math.min(y0, b.y0);
      x1 = Math.max(x1, b.x1); y1 = Math.max(y1, b.y1);
    }
    return isFinite(x0) ? { x0, y0, x1, y1 } : { x0: 0, y0: 0, x1: 0, y1: 0 };
  }
  function offsetNode(node, dx, dy) {
    if (!node) return;
    if (node.bbox) {
      node.bbox = { x0: node.bbox.x0 + dx, y0: node.bbox.y0 + dy, x1: node.bbox.x1 + dx, y1: node.bbox.y1 + dy };
    }
    if (Array.isArray(node.symbols)) node.symbols.forEach(s => offsetNode(s, dx, dy));
  }

  // ---------------------------------------------------------------- bản đồ mực
  function buildInkMap(canvas) {
    const w = canvas.width, h = canvas.height;
    const d = ctx2d(canvas).getImageData(0, 0, w, h).data;
    const n = w * h;
    const gray = new Uint8Array(n);
    const hist = new Int32Array(256);
    for (let i = 0, p = 0; p < n; i += 4, p++) {
      const g = (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]) | 0;
      gray[p] = g; hist[g]++;
    }
    let sum = 0;
    for (let t = 0; t < 256; t++) sum += t * hist[t];
    let sumB = 0, wB = 0, best = 0, thr = 160;
    for (let t = 0; t < 256; t++) {
      wB += hist[t];
      if (!wB) continue;
      const wF = n - wB;
      if (!wF) break;
      sumB += t * hist[t];
      const mB = sumB / wB, mF = (sum - sumB) / wF;
      const v = wB * wF * (mB - mF) * (mB - mF);
      if (v > best) { best = v; thr = t; }
    }
    const cut = Math.min(200, Math.max(110, thr + 8));
    const dark = new Uint8Array(n);
    for (let p = 0; p < n; p++) dark[p] = gray[p] < cut ? 1 : 0;
    return { w, h, dark, cut };
  }

  // tỉ lệ cột hoàn toàn trắng trong vùng [x0,x1) x [y0,y1)
  function blankColumnFraction(ink, x0, x1, y0, y1) {
    x0 = Math.max(0, Math.floor(x0)); x1 = Math.min(ink.w, Math.ceil(x1));
    y0 = Math.max(0, Math.floor(y0)); y1 = Math.min(ink.h, Math.ceil(y1));
    if (x1 <= x0 || y1 <= y0) return 0;
    let blank = 0;
    for (let x = x0; x < x1; x++) {
      let hit = false;
      for (let y = y0; y < y1; y++) { if (ink.dark[y * ink.w + x]) { hit = true; break; } }
      if (!hit) blank++;
    }
    return blank / (x1 - x0);
  }
  function darkRatio(ink, x0, x1, y0, y1) {
    x0 = Math.max(0, Math.floor(x0)); x1 = Math.min(ink.w, Math.ceil(x1));
    y0 = Math.max(0, Math.floor(y0)); y1 = Math.min(ink.h, Math.ceil(y1));
    if (x1 <= x0 || y1 <= y0) return 0;
    let c = 0;
    for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) c += ink.dark[y * ink.w + x];
    return c / ((x1 - x0) * (y1 - y0));
  }

  // các dải trắng liên tiếp bên trong 1 vùng (dùng để nghi ngờ chữ dính khi không có bbox ký tự)
  function findBlankRuns(ink, x0, x1, y0, y1, minWidth) {
    x0 = Math.max(0, Math.floor(x0)); x1 = Math.min(ink.w, Math.ceil(x1));
    const runs = [];
    let start = -1;
    for (let x = x0; x <= x1; x++) {
      let blank = true;
      if (x < x1) {
        for (let y = Math.max(0, Math.floor(y0)); y < Math.min(ink.h, Math.ceil(y1)); y++) {
          if (ink.dark[y * ink.w + x]) { blank = false; break; }
        }
      }
      if (blank && x < x1) { if (start < 0) start = x; }
      else if (start >= 0) {
        if (x - start >= minWidth && start > x0 + 1 && x < x1) runs.push({ x0: start, x1: x });
        start = -1;
      }
    }
    return runs;
  }

  // ước lượng chiều cao dải chữ (px) để biết trang có chữ nhỏ cần phóng to
  function estimateTextHeight(ink) {
    const rowCount = new Int32Array(ink.h);
    for (let y = 0; y < ink.h; y++) {
      let c = 0;
      const off = y * ink.w;
      for (let x = 0; x < ink.w; x++) c += ink.dark[off + x];
      rowCount[y] = c;
    }
    const bands = [];
    let s = -1, gap = 0;
    for (let y = 0; y < ink.h; y++) {
      const has = rowCount[y] >= 2 && rowCount[y] < ink.w * 0.45; // loại đường kẻ dài
      if (has) { if (s < 0) s = y; gap = 0; }
      else if (s >= 0) {
        gap++;
        if (gap > 2) { const hgt = y - gap - s; if (hgt >= 5 && hgt < ink.h * 0.12) bands.push(hgt); s = -1; gap = 0; }
      }
    }
    return median(bands);
  }

  // ---------------------------------------------------------------- tiền xử lý thích ứng
  function measureBackground(canvas) {
    const bw = 24, bh = 24;
    const small = newCanvas(bw, bh);
    const sctx = small.getContext('2d', { willReadFrequently: true });
    sctx.imageSmoothingEnabled = true;
    sctx.drawImage(canvas, 0, 0, bw, bh);
    const d = sctx.getImageData(0, 0, bw, bh).data;
    const vals = [];
    for (let i = 0; i < d.length; i += 4) vals.push(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
    const m = mean(vals);
    const sd = Math.sqrt(mean(vals.map(v => (v - m) * (v - m))));
    // nền = phần sáng nhất của từng ô; ước lượng độ không đều của nền
    const sorted = vals.slice().sort((a, b) => b - a);
    const top = sorted.slice(0, Math.max(6, Math.round(sorted.length * 0.5)));
    const topMean = mean(top);
    const topSd = Math.sqrt(mean(top.map(v => (v - topMean) * (v - topMean))));
    return { meanGray: m, sd, bgMean: topMean, unevenness: topSd };
  }

  function flattenBackground(canvas) {
    const w = canvas.width, h = canvas.height;
    const sw = Math.max(8, Math.round(w / 8)), sh = Math.max(8, Math.round(h / 8));
    const small = newCanvas(sw, sh);
    const sctx = small.getContext('2d', { willReadFrequently: true });
    sctx.imageSmoothingEnabled = true;
    sctx.drawImage(canvas, 0, 0, sw, sh);
    const sd = sctx.getImageData(0, 0, sw, sh);
    const g = new Float32Array(sw * sh);
    for (let i = 0, p = 0; p < g.length; i += 4, p++) g[p] = 0.299 * sd.data[i] + 0.587 * sd.data[i + 1] + 0.114 * sd.data[i + 2];
    // giãn nở (max filter) để xóa nét chữ khỏi ước lượng nền, rồi làm mờ hộp
    const R = 4;
    const dil = new Float32Array(g.length);
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      let mx = 0;
      for (let yy = Math.max(0, y - R); yy <= Math.min(sh - 1, y + R); yy++)
        for (let xx = Math.max(0, x - R); xx <= Math.min(sw - 1, x + R); xx++) mx = Math.max(mx, g[yy * sw + xx]);
      dil[y * sw + x] = mx;
    }
    const blur = new Float32Array(g.length);
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      let s = 0, c = 0;
      for (let yy = Math.max(0, y - 3); yy <= Math.min(sh - 1, y + 3); yy++)
        for (let xx = Math.max(0, x - 3); xx <= Math.min(sw - 1, x + 3); xx++) { s += dil[yy * sw + xx]; c++; }
      blur[y * sw + x] = s / c;
    }
    const bgSmall = newCanvas(sw, sh);
    const bctx = bgSmall.getContext('2d');
    const bd = bctx.createImageData(sw, sh);
    for (let p = 0, i = 0; p < blur.length; p++, i += 4) { bd.data[i] = bd.data[i + 1] = bd.data[i + 2] = Math.min(255, blur[p]); bd.data[i + 3] = 255; }
    bctx.putImageData(bd, 0, 0);
    const bgFull = newCanvas(w, h);
    const fctx = bgFull.getContext('2d', { willReadFrequently: true });
    fctx.imageSmoothingEnabled = true;
    fctx.imageSmoothingQuality = 'high';
    fctx.drawImage(bgSmall, 0, 0, w, h);
    const bgData = fctx.getImageData(0, 0, w, h).data;

    const out = newCanvas(w, h);
    const octx = out.getContext('2d', { willReadFrequently: true });
    octx.drawImage(canvas, 0, 0);
    const od = octx.getImageData(0, 0, w, h);
    const o = od.data;
    for (let i = 0; i < o.length; i += 4) {
      const gray = 0.299 * o[i] + 0.587 * o[i + 1] + 0.114 * o[i + 2];
      const bg = Math.max(70, bgData[i]);
      let v = (gray / bg) * 255;
      v = v > 238 ? 255 : Math.max(0, Math.min(255, v));
      o[i] = o[i + 1] = o[i + 2] = v;
    }
    octx.putImageData(od, 0, 0);
    return out;
  }

  // xóa điểm nhiễu hạt cô lập (không đụng tới nét chữ/dấu vì nét luôn liền ≥ 2 điểm lân cận)
  function despeckle(canvas, ink) {
    const w = canvas.width, h = canvas.height;
    const c = ctx2d(canvas);
    const id = c.getImageData(0, 0, w, h);
    const d = id.data;
    let removed = 0;
    for (let y = 1; y < h - 1; y++) {
      for (let x = 1; x < w - 1; x++) {
        const p = y * w + x;
        if (!ink.dark[p]) continue;
        const nb = ink.dark[p - 1] + ink.dark[p + 1] + ink.dark[p - w] + ink.dark[p + w] +
          ink.dark[p - w - 1] + ink.dark[p - w + 1] + ink.dark[p + w - 1] + ink.dark[p + w + 1];
        if (nb === 0) { const i = p * 4; d[i] = d[i + 1] = d[i + 2] = 255; removed++; }
      }
    }
    c.putImageData(id, 0, 0);
    return removed;
  }

  function speckleRatio(ink) {
    let dark = 0, iso = 0;
    const w = ink.w, h = ink.h;
    for (let y = 1; y < h - 1; y += 2) for (let x = 1; x < w - 1; x += 2) {
      const p = y * w + x;
      if (!ink.dark[p]) continue;
      dark++;
      if (ink.dark[p - 1] + ink.dark[p + 1] + ink.dark[p - w] + ink.dark[p + w] === 0) iso++;
    }
    return dark ? iso / dark : 0;
  }

  /** Chọn phép tiền xử lý riêng cho từng trang. Trả về canvas mới (hệ tọa độ MỚI) + nhật ký. */
  function adaptivePreprocess(canvas) {
    const applied = [];
    let work = canvas;
    let scale = 1;
    try {
      const bg = measureBackground(work);
      if (bg.unevenness > 11 || bg.bgMean < 222) {
        work = flattenBackground(work);
        applied.push(`Làm phẳng nền/bóng (độ lệch nền ${bg.unevenness.toFixed(1)}, nền ${Math.round(bg.bgMean)})`);
      }
      let ink = buildInkMap(work);
      const sp = speckleRatio(ink);
      if (sp > 0.12) {
        const n = despeckle(work, ink);
        applied.push(`Khử nhiễu hạt (${n} điểm cô lập, tỉ lệ ${(sp * 100).toFixed(0)}%)`);
        ink = buildInkMap(work);
      }
      const th = estimateTextHeight(ink);
      if (th > 0 && th < 26 && work.width < 3200) {
        scale = Math.min(2.2, Math.max(1.25, 34 / th));
        scale = Math.min(scale, 3400 / work.width);
        if (scale > 1.1) {
          const up = newCanvas(work.width * scale, work.height * scale);
          const uctx = up.getContext('2d');
          uctx.fillStyle = '#fff'; uctx.fillRect(0, 0, up.width, up.height);
          uctx.imageSmoothingEnabled = true; uctx.imageSmoothingQuality = 'high';
          uctx.drawImage(work, 0, 0, up.width, up.height);
          work = up;
          applied.push(`Phóng chữ nhỏ ×${scale.toFixed(2)} (cao chữ ~${Math.round(th)}px)`);
        } else scale = 1;
      }
      return { canvas: work, applied, scale, textHeight: th };
    } catch (e) {
      return { canvas, applied: ['Bỏ qua tiền xử lý thích ứng: ' + e.message], scale: 1, textHeight: 0 };
    }
  }

  // ---------------------------------------------------------------- sửa khoảng trắng dựa trên ảnh
  function lineHeightOf(words) {
    return Math.max(8, median(words.map(w => w.bbox.y1 - w.bbox.y0)));
  }

  /**
   * (Loại B) Tách từ bị dính. Chỉ tách khi đồng thời:
   *  - khoảng cách giữa 2 KÝ TỰ liên tiếp (bbox ký tự do engine trả về) lớn hơn rõ rệt so với khoảng cách
   *    giữa các ký tự trong cùng dòng, và
   *  - dải ảnh giữa 2 ký tự đó thực sự TRẮNG (≥ 70% cột không có điểm mực).
   * Không đủ bằng chứng -> giữ nguyên và đánh dấu "nghi ngờ dính chữ" để người dùng xác nhận.
   */
  function repairLineSpacing(line, ink, log) {
    const words = (line.words || []).filter(w => w && w.bbox && String(w.text || '').trim());
    if (!words.length) return [];
    const H = lineHeightOf(words);
    const inner = [];
    for (const w of words) {
      const sy = Array.isArray(w.symbols) ? w.symbols.filter(s => s && s.bbox) : [];
      for (let i = 0; i + 1 < sy.length; i++) {
        const g = sy[i + 1].bbox.x0 - sy[i].bbox.x1;
        if (g < 0.15 * H) inner.push(Math.max(0, g));
      }
    }
    const m = median(inner);
    const thr = Math.max(3, 0.15 * H, 2.4 * m + 1.5);
    const out = [];
    for (const w of words) {
      const sy = Array.isArray(w.symbols) ? w.symbols.filter(s => s && s.bbox && String(s.text || '') !== '') : [];
      const joined = nfc(sy.map(s => s.text).join(''));
      const hasSym = sy.length >= 2 && joined === nfc(String(w.text || '').trim());
      if (!hasSym) {
        const wText = String(w.text || '').trim();
        const copy = Object.assign({}, w);
        if (wText.length >= 8) {
          const runs = findBlankRuns(ink, w.bbox.x0, w.bbox.x1, w.bbox.y0, w.bbox.y1, Math.max(3, 0.2 * H));
          if (runs.length) { copy.suspectStuck = true; log.suspects.push({ text: wText, bbox: w.bbox, reason: 'Có khoảng trắng trong ảnh nhưng không có bbox ký tự để tách an toàn' }); }
        }
        out.push(copy);
        continue;
      }
      const cuts = [];
      for (let i = 0; i + 1 < sy.length; i++) {
        const g = sy[i + 1].bbox.x0 - sy[i].bbox.x1;
        if (g < thr) continue;
        const blank = blankColumnFraction(ink, sy[i].bbox.x1, sy[i + 1].bbox.x0, w.bbox.y0, w.bbox.y1);
        if (blank >= 0.7) cuts.push({ at: i + 1, gap: g, blank });
      }
      if (!cuts.length) { out.push(w); continue; }
      const bounds = [0].concat(cuts.map(c => c.at), [sy.length]);
      const parts = [];
      for (let k = 0; k + 1 < bounds.length; k++) {
        const seg = sy.slice(bounds[k], bounds[k + 1]);
        const text = seg.map(s => s.text).join('');
        if (!text.trim()) continue;
        parts.push({
          text: nfc(text),
          confidence: mean(seg.map(s => Number(s.confidence) || Number(w.confidence) || 0)),
          bbox: unionBBox(seg),
          symbols: seg,
          splitFromSpacing: true,
          origText: String(w.text || '').trim()
        });
      }
      if (parts.length >= 2) {
        log.repairs.push({
          kind: 'B',
          original: String(w.text || '').trim(),
          proposed: parts.map(p => p.text).join(' '),
          bbox: w.bbox,
          evidence: `Khoảng trống ảnh ${cuts.map(c => Math.round(c.gap) + 'px/' + Math.round(c.blank * 100) + '% trắng').join(', ')} (chuẩn ký tự trong dòng ≈ ${Math.round(m)}px, cao dòng ${Math.round(H)}px)`
        });
        parts.forEach(p => out.push(p));
      } else out.push(w);
    }
    return out;
  }

  function joinWordsText(words) {
    return nfc(words.map(w => String(w.text || '').trim()).filter(Boolean).join(' '));
  }

  function applyRepairToLine(line, ink, log) {
    const before = line.words ? line.words.length : 0;
    const raw = nfc(line.text || '');
    line.rawText = raw;
    const fixed = repairLineSpacing(line, ink, log);
    if (fixed.length) {
      line.words = fixed;
      line.text = joinWordsText(fixed);
    }
    line.spacingRepaired = fixed.length !== before;
    const confs = (line.words || []).map(w => Number(w.confidence)).filter(v => !isNaN(v));
    line.minWordConf = confs.length ? Math.min.apply(null, confs) : 100;
    line.meanWordConf = mean(confs);
    return line;
  }

  // ---------------------------------------------------------------- OCR từng ô bảng
  async function ocrTableCells(worker, canvas, grid, ink, onProgress) {
    if (!grid || !grid.cells || !grid.cells.length || grid.cells.length > MAX_CELL_OCR) return null;
    const inset = Math.max(4, Math.round(canvas.width * 0.0025));
    const pad = 18;
    const results = [];
    await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
    let k = 0;
    for (const cell of grid.cells) {
      k++;
      if (onProgress) onProgress(84 + Math.round((k / grid.cells.length) * 12), `OCR riêng từng ô bảng ${k}/${grid.cells.length} (không trộn nội dung giữa các ô)...`);
      const x0 = cell.x0 + inset, y0 = cell.y0 + inset, x1 = cell.x1 - inset, y1 = cell.y1 - inset;
      const cw = x1 - x0, ch = y1 - y0;
      const rec = { cell, r: cell.r, c: cell.c, lines: [], words: [], empty: false };
      if (cw < 14 || ch < 10 || darkRatio(ink, x0, x1, y0, y1) < 0.0012) { rec.empty = true; results.push(rec); continue; }
      const crop = newCanvas(cw + pad * 2, ch + pad * 2);
      const cc = crop.getContext('2d');
      cc.fillStyle = '#fff'; cc.fillRect(0, 0, crop.width, crop.height);
      cc.drawImage(canvas, x0, y0, cw, ch, pad, pad, cw, ch);
      let data = null;
      try { data = (await worker.recognize(crop)).data; } catch (e) { data = null; }
      if (data && Array.isArray(data.lines)) {
        const dx = x0 - pad, dy = y0 - pad;
        for (const ln of data.lines) {
          offsetNode(ln, dx, dy);
          (ln.words || []).forEach(w => { offsetNode(w, dx, dy); w.cell = { r: cell.r, c: cell.c }; });
          ln.cell = { r: cell.r, c: cell.c };
          if (String(ln.text || '').trim()) { rec.lines.push(ln); (ln.words || []).forEach(w => rec.words.push(w)); }
        }
      }
      results.push(rec);
    }
    return results;
  }

  // ---------------------------------------------------------------- OCR lại vùng ảnh nghi ngờ (loại A)
  async function reOcrSuspectLines(worker, canvas, lines, onProgress) {
    const cand = lines
      .filter(l => l.bbox && String(l.text || '').trim().length >= 3 && (l.minWordConf < 70 || l.meanWordConf < 82 || l.words.some(w => w.suspectStuck)))
      .sort((a, b) => a.meanWordConf - b.meanWordConf)
      .slice(0, MAX_REOCR_LINES);
    const suggestions = [];
    if (!cand.length) return suggestions;
    await worker.setParameters({ tessedit_pageseg_mode: '7', preserve_interword_spaces: '1' });
    let k = 0;
    for (const ln of cand) {
      k++;
      if (onProgress) onProgress(96, `Đối chiếu ảnh gốc: OCR lại vùng nghi ngờ ${k}/${cand.length}...`);
      const b = ln.bbox;
      const hh = Math.max(10, b.y1 - b.y0);
      const px = 10, py = Math.round(hh * 0.25) + 3;
      const sx0 = Math.max(0, b.x0 - px), sy0 = Math.max(0, b.y0 - py);
      const sx1 = Math.min(canvas.width, b.x1 + px), sy1 = Math.min(canvas.height, b.y1 + py);
      const sw = sx1 - sx0, sh = sy1 - sy0;
      if (sw < 12 || sh < 8) continue;
      const scale = hh < 40 ? 2 : 1.4;
      const crop = newCanvas(sw * scale + 40, sh * scale + 40);
      const cc = crop.getContext('2d');
      cc.fillStyle = '#fff'; cc.fillRect(0, 0, crop.width, crop.height);
      cc.imageSmoothingEnabled = true; cc.imageSmoothingQuality = 'high';
      cc.drawImage(canvas, sx0, sy0, sw, sh, 20, 20, sw * scale, sh * scale);
      let d2 = null;
      try { d2 = (await worker.recognize(crop)).data; } catch (e) { continue; }
      const t2 = nfc(String(d2.text || '').replace(/\s+/g, ' ').trim());
      const t1 = nfc(String(ln.text || '').replace(/\s+/g, ' ').trim());
      if (!t2 || t2 === t1) continue;
      const c2 = Number(d2.confidence) || 0;
      const gain = c2 - ln.meanWordConf;
      if (gain >= 6) {
        suggestions.push({
          kind: 'A',
          original: t1,
          proposed: t2,
          bbox: b,
          lineRef: ln,
          confGain: Math.round(gain),
          evidence: `OCR lại đúng vùng ảnh dòng (phóng ×${scale}, PSM 7): tin cậy ${Math.round(ln.meanWordConf)}% → ${Math.round(c2)}%`
        });
      }
    }
    return suggestions;
  }

  // ---------------------------------------------------------------- phân vùng bố cục
  function analyzeRegions(lines, grid, W, H) {
    const regions = [];
    const textLines = lines.filter(l => l.bbox && String(l.text || '').trim());
    const lh = median(textLines.map(l => l.bbox.y1 - l.bbox.y0)) || 20;
    if (grid) {
      regions.push({ type: 'table', bbox: { x0: grid.tableLeft, y0: grid.tableTop, x1: grid.tableRight, y1: grid.tableBottom }, rows: grid.numRows, cols: grid.numCols });
    }
    const outside = textLines.filter(l => {
      if (!grid) return true;
      const cy = (l.bbox.y0 + l.bbox.y1) / 2, cx = (l.bbox.x0 + l.bbox.x1) / 2;
      return !(cx >= grid.tableLeft - 4 && cx <= grid.tableRight + 4 && cy >= grid.tableTop - 4 && cy <= grid.tableBottom + 4);
    }).sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);

    // phát hiện bố cục 2 cột (không phải bảng): có nhiều dòng nằm trọn nửa trái và nhiều dòng nằm trọn nửa phải
    let columns = 1;
    if (!grid && outside.length >= 8) {
      const mid = W / 2;
      const left = outside.filter(l => l.bbox.x1 < mid + W * 0.03).length;
      const right = outside.filter(l => l.bbox.x0 > mid - W * 0.03).length;
      const wide = outside.filter(l => (l.bbox.x1 - l.bbox.x0) > W * 0.6).length;
      if (left >= outside.length * 0.3 && right >= outside.length * 0.3 && wide <= outside.length * 0.2) columns = 2;
    }

    const typed = outside.map(l => {
      const t = nfc(l.text).trim();
      const letters = t.replace(/[^A-Za-zÀ-ỹĐđ]/g, '');
      const upper = letters ? (t.replace(/[^A-ZÀ-ỸĐ]/g, '').length / letters.length) : 0;
      const cx = (l.bbox.x0 + l.bbox.x1) / 2;
      const hLine = l.bbox.y1 - l.bbox.y0;
      let type = 'paragraph';
      if (/^[-–—]?\s*\d{1,3}\s*[-–—]?$/.test(t) && (l.bbox.y0 > H * 0.9 || l.bbox.y1 < H * 0.07)) type = 'page-number';
      else if (/^(Ghi chú|Chú thích|\*)/i.test(t)) type = 'note';
      else if (/^(\d+[.)]|[a-zđ][.)]|[-–•+*])\s/.test(t)) type = 'list';
      else if (l.bbox.y0 > H * 0.62 && /^(Nơi nhận|Lưu:|T\/M|TM\.|KT\.|TL\.)/i.test(t)) type = 'footer';
      else if (letters.length >= 4 && upper > 0.75 && Math.abs(cx - W / 2) < 0.12 * W && hLine >= lh * 1.05) type = 'title';
      else if (l.bbox.y1 < H * 0.24 && letters.length >= 3 && (upper > 0.6 || /^(Số|Ngày|Hà Nội|TP\.|Thành phố)/i.test(t))) type = 'header';
      return { l, type, col: columns === 2 ? (cx < W / 2 ? 0 : 1) : 0 };
    });
    // footer kéo dài tới hết trang
    let seenFooter = false;
    typed.forEach(o => { if (o.type === 'footer') seenFooter = true; else if (seenFooter && o.type === 'paragraph') o.type = 'footer'; });

    let cur = null;
    const flush = () => { if (cur) { cur.bbox = unionBBox(cur.lines.map(x => x.bbox ? x : { bbox: x })); regions.push(cur); cur = null; } };
    typed.forEach(o => {
      if (cur && cur.type === o.type && cur.col === o.col && o.l.bbox.y0 - cur.lastY1 < lh * 1.9) {
        cur.lines.push(o.l); cur.lastY1 = o.l.bbox.y1;
      } else {
        flush();
        cur = { type: o.type, col: o.col, lines: [o.l], lastY1: o.l.bbox.y1 };
      }
    });
    flush();
    regions.forEach(r => { if (r.lines) { r.lineCount = r.lines.length; r.bbox = unionBBox(r.lines); delete r.lines; delete r.lastY1; } });
    regions.sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
    return { regions, columns };
  }

  // ---------------------------------------------------------------- dựng văn bản theo thứ tự đọc
  function cellMatrixText(cells) {
    const byRow = new Map();
    cells.forEach(c => { if (!byRow.has(c.r)) byRow.set(c.r, []); byRow.get(c.r).push(c); });
    return Array.from(byRow.keys()).sort((a, b) => a - b).map(r => {
      return byRow.get(r).sort((a, b) => a.c - b.c).map(c => c.lines.map(l => nfc(l.text).trim()).filter(Boolean).join(' ')).join(' | ');
    }).join('\n');
  }

  function composePageText(pageLines, cellRecs, grid, columns, W) {
    const lh = median(pageLines.map(l => l.bbox ? l.bbox.y1 - l.bbox.y0 : 0).filter(Boolean)) || 20;
    const items = pageLines.filter(l => String(l.text || '').trim()).map(l => ({
      y: l.bbox.y0, y1: l.bbox.y1, col: columns === 2 ? (((l.bbox.x0 + l.bbox.x1) / 2) < W / 2 ? 0 : 1) : 0, text: nfc(l.text).trim(), x: l.bbox.x0
    }));
    if (grid && cellRecs) items.push({ y: grid.tableTop, y1: grid.tableBottom, col: 0, text: cellMatrixText(cellRecs), table: true, x: grid.tableLeft });
    items.sort((a, b) => (a.col - b.col) || (a.y - b.y) || (a.x - b.x));
    const out = [];
    let prev = null;
    for (const it of items) {
      if (prev && prev.col === it.col && (it.y - prev.y1) > lh * 1.1) out.push('');
      else if (prev && prev.col !== it.col) out.push('');
      out.push(it.text);
      prev = it;
    }
    return out.join('\n');
  }

  // ---------------------------------------------------------------- gắn cờ cần kiểm tra
  function annotateWords(words, confusable) {
    const flags = [];
    for (const w of words) {
      const t = String(w.text || '').trim();
      if (!t || !w.bbox) continue;
      const f = [];
      if (Number(w.confidence) < 70) f.push('lowconf');
      if (w.splitFromSpacing) f.push('spacing');
      if (w.suspectStuck) f.push('stuck');
      const dm = t.match(/^(\d{1,2})[\/.-](\d{1,2})(?:[\/.-](\d{2,4}))?$/);
      if (dm && (Number(dm[1]) > 31 || Number(dm[1]) < 1 || Number(dm[2]) > 12 || Number(dm[2]) < 1)) f.push('date');
      if (/\d/.test(t) && /[OIlS]/.test(t) && !/^(QĐ|NĐ|HD|CV|BC|TT|TB|NQ|KH|UBND|HĐND|VPTW)/i.test(t)) f.push('confusable');
      w.flags = f;
      if (f.length) flags.push({ text: t, flags: f, bbox: w.bbox, confidence: Math.round(Number(w.confidence) || 0) });
    }
    return flags;
  }

  // ---------------------------------------------------------------- điều phối 1 trang
  async function enhancePage(opts) {
    const canvas = opts.canvas, data = opts.data, worker = opts.worker, onProgress = opts.onProgress || null;
    const t0 = performance.now();
    const ink = buildInkMap(canvas);
    const log = { repairs: [], suspects: [] };
    const W = canvas.width, H = canvas.height;

    // sao chép nông các dòng để giữ nguyên dữ liệu OCR gốc (raw) phục vụ "khôi phục nội dung ban đầu"
    let lines = (data.lines || []).map(l => Object.assign({}, l, { words: (l.words || []).map(w => Object.assign({}, w)) }));

    let grid = null;
    try {
      const TG = global.VietNamOcrTableGrid;
      if (TG && TG.detectTableGridFromCanvas) grid = TG.detectTableGridFromCanvas(canvas);
    } catch (e) { grid = null; }

    let cellRecs = null;
    if (grid && opts.cellOcr !== false) {
      try { cellRecs = await ocrTableCells(worker, canvas, grid, ink, onProgress); } catch (e) { cellRecs = null; }
    }

    const inTable = (b) => {
      if (!grid || !b) return false;
      const cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2;
      return cx >= grid.tableLeft - 4 && cx <= grid.tableRight + 4 && cy >= grid.tableTop - 4 && cy <= grid.tableBottom + 4;
    };

    let pageLines = lines;
    if (grid && cellRecs) pageLines = lines.filter(l => !inTable(l.bbox));
    pageLines.forEach(l => applyRepairToLine(l, ink, log));
    if (cellRecs) cellRecs.forEach(rec => rec.lines.forEach(l => applyRepairToLine(l, ink, log)));

    // danh sách dòng + từ phẳng sau hiệu chỉnh
    const allLines = pageLines.slice();
    if (cellRecs) cellRecs.forEach(rec => rec.lines.forEach(l => allLines.push(l)));
    let words = [];
    allLines.forEach(l => (l.words || []).forEach(w => { if (String(w.text || '').trim()) words.push(w); }));
    // nếu có bảng nhưng không OCR được từng ô: giữ từ ở mức trang
    if (grid && !cellRecs) { words = []; pageLines.forEach(l => (l.words || []).forEach(w => words.push(w))); }

    const confusable = annotateWords(words);

    // đối chiếu lại vùng ảnh nghi ngờ (chỉ đề xuất, không tự áp dụng)
    let suggestions = [];
    if (opts.reOcr !== false) {
      try { suggestions = await reOcrSuspectLines(worker, canvas, allLines, onProgress); } catch (e) { suggestions = []; }
    }

    const regionInfo = analyzeRegions(pageLines, grid, W, H);
    const text = composePageText(pageLines, cellRecs, grid, regionInfo.columns, W);

    return {
      lines: allLines,
      pageLines,
      words,
      text,
      grid,
      cellRecs,
      regions: regionInfo.regions,
      columns: regionInfo.columns,
      repairs: log.repairs,
      suspects: log.suspects,
      flags: confusable,
      suggestions,
      elapsedMs: Math.round(performance.now() - t0)
    };
  }

  global.VietNamOcrLayout = {
    buildInkMap, blankColumnFraction, darkRatio, findBlankRuns, estimateTextHeight,
    adaptivePreprocess, measureBackground, flattenBackground, despeckle,
    repairLineSpacing, ocrTableCells, reOcrSuspectLines, analyzeRegions, composePageText,
    cellMatrixText, annotateWords, enhancePage
  };
})(window);
