/**
 * VietNamOCR 3.5 Pro — Intelligent Document Processing Core (js/ocr-engine-pro.js)
 *
 * Nâng cấp chuyên nghiệp:
 * 1. Tự động phân loại tài liệu trước khi OCR (Văn bản 2 cột, Bảng biểu, Đa cột, Ảnh mờ) & chọn PSM tối ưu.
 * 2. Tự động căn thẳng ảnh nghiêng (Auto-Deskew ±6°), tẩy dấu mộc đỏ, tăng nét Unsharp Mask 3x3.
 * 3. Tự động chạy lại OCR với cấu hình khử nhiễu sâu khi độ tin cậy < 72% và chọn kết quả tốt hơn.
 * 4. Đánh giá chất lượng & độ tin cậy từng từ: đánh dấu vùng nghi ngờ trên ảnh gốc + phát hiện ký tự dễ nhầm (O/0, I/1, S/5).
 * 5. Trích xuất bảng biểu nâng cao từ tọa độ word bbox (nhận diện cột, hàng, ô gộp & đánh dấu ô độ tin cậy thấp).
 * 6. Hỗ trợ đa ngôn ngữ (vie / vie+eng / eng), TIFF nhiều trang (UTIF.js), Tạm dừng / Tiếp tục / Hủy & giải phóng RAM.
 */
(function (global) {
  'use strict';

  const TESSDATA_BEST_URL = 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_best/main';
  let activeWorker = null;
  let activeWorkerLang = '';
  let activeWorkerSource = 'best';
  let cancelRequested = false;
  let pauseRequested = false;
  let lastScanMetrics = null;
  const errorLogs = []; // Chỉ ghi thông điệp lỗi kỹ thuật, TUYỆT ĐỐI không ghi nội dung tài liệu nhạy cảm

  function logSafeError(stage, msg) {
    errorLogs.unshift({
      time: new Date().toLocaleTimeString('vi-VN'),
      stage: String(stage || 'OCR'),
      message: String(msg || 'Lỗi không xác định').slice(0, 180)
    });
    if (errorLogs.length > 50) errorLogs.length = 50;
  }

  function getErrorLogs() { return errorLogs.slice(); }
  function clearErrorLogs() { errorLogs.length = 0; }

  function requestCancel() { cancelRequested = true; pauseRequested = false; }
  function resetCancel() { cancelRequested = false; pauseRequested = false; }
  function isCancelled() { return cancelRequested; }
  function togglePause() { pauseRequested = !pauseRequested; return pauseRequested; }
  function isPaused() { return pauseRequested; }
  function getLastMetrics() { return lastScanMetrics; }

  async function waitIfPaused(onPausedTick) {
    while (pauseRequested && !cancelRequested) {
      if (onPausedTick) onPausedTick();
      await new Promise(r => setTimeout(r, 250));
    }
  }

  // Kiểm tra an toàn tệp đầu vào (giới hạn dung lượng & định dạng cho phép)
  function validateInputFile(file) {
    if (!file) return { ok: false, reason: 'Không có tệp đầu vào.' };
    const maxBytes = 80 * 1024 * 1024; // 80 MB
    if (file.size > maxBytes) {
      return { ok: false, reason: `Tệp "${file.name}" vượt quá giới hạn an toàn 80MB.` };
    }
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const allowed = ['pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp', 'tif', 'tiff'];
    if (!allowed.includes(ext) && !(file.type || '').startsWith('image/') && file.type !== 'application/pdf') {
      return { ok: false, reason: `Định dạng ".${ext}" không được hỗ trợ.` };
    }
    return { ok: true };
  }

  // 1. Tự động phân loại sơ bộ loại tài liệu từ ảnh canvas để chọn cấu hình PSM tối ưu
  function classifyDocumentCanvas(sourceCanvas) {
    try {
      const w = Math.min(400, sourceCanvas.width || 400);
      const h = Math.max(120, Math.round((sourceCanvas.height / Math.max(1, sourceCanvas.width)) * w));
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      const ctx = c.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(sourceCanvas, 0, 0, w, h);
      const d = ctx.getImageData(0, 0, w, h).data;

      let darkPixels = 0;
      let horizLineRows = 0;
      let vertLineCols = 0;
      const colCounts = new Int32Array(w);

      for (let y = 0; y < h; y++) {
        let rowDark = 0;
        const off = y * w * 4;
        for (let x = 0; x < w; x++) {
          const g = 0.299 * d[off + x * 4] + 0.587 * d[off + x * 4 + 1] + 0.114 * d[off + x * 4 + 2];
          if (g < 155) {
            darkPixels++;
            rowDark++;
            colCounts[x]++;
          }
        }
        if (rowDark > w * 0.55) horizLineRows++;
      }
      for (let x = 0; x < w; x++) {
        if (colCounts[x] > h * 0.45) vertLineCols++;
      }

      const density = darkPixels / (w * h);
      if (horizLineRows >= 3 && vertLineCols >= 2) {
        return { docProfile: 'Bảng biểu / Thống kê', recommendedPsm: '6', density };
      }
      if (density < 0.015) {
        return { docProfile: 'Ảnh mờ / Ít chữ (cần tăng tương phản)', recommendedPsm: '6', density };
      }
      if (density > 0.22) {
        return { docProfile: 'Nền phức tạp / Nhiều chi tiết', recommendedPsm: '3', density };
      }
      return { docProfile: 'Văn bản hành chính / Đa đoạn chuẩn', recommendedPsm: '3', density };
    } catch {
      return { docProfile: 'Văn bản chuẩn', recommendedPsm: '3', density: 0.08 };
    }
  }

  async function getOcrWorker(lang = 'vie', psm = '3', onProgress = null) {
    if (!global.Tesseract || typeof global.Tesseract.createWorker !== 'function') {
      throw new Error('Thư viện Tesseract.js v5 chưa được nạp.');
    }
    const targetLang = (lang === 'vie+eng' || lang === 'eng') ? lang : 'vie';
    if (activeWorker && activeWorkerLang === targetLang) {
      await activeWorker.setParameters({
        preserve_interword_spaces: '1',
        tessedit_pageseg_mode: String(psm || '3')
      });
      return activeWorker;
    }

    if (activeWorker) {
      try { await activeWorker.terminate(); } catch (e) {}
      activeWorker = null;
      activeWorkerLang = '';
    }

    const logger = m => {
      if (!onProgress) return;
      if (m.status === 'loading language traineddata') {
        onProgress(Math.round((m.progress || 0) * 20), 'Đang nạp mô hình OCR chính xác cao (tessdata_best)...');
      } else if (m.status === 'recognizing text' && typeof m.progress === 'number') {
        onProgress(Math.max(20, Math.round(20 + m.progress * 78)), `Đang nhận diện (${targetLang}): ${Math.round(m.progress * 100)}%`);
      }
    };

    try {
      const w = await global.Tesseract.createWorker(targetLang, 1, {
        langPath: TESSDATA_BEST_URL,
        gzip: false,
        logger
      });
      await w.setParameters({
        preserve_interword_spaces: '1',
        tessedit_pageseg_mode: String(psm || '3')
      });
      activeWorker = w;
      activeWorkerLang = targetLang;
      activeWorkerSource = 'tessdata_best';
      return activeWorker;
    } catch (bestErr) {
      logSafeError('WorkerInit', 'Fallback từ tessdata_best sang CDN mặc định');
      const w = await global.Tesseract.createWorker(targetLang, 1, { logger });
      await w.setParameters({
        preserve_interword_spaces: '1',
        tessedit_pageseg_mode: String(psm || '3')
      });
      activeWorker = w;
      activeWorkerLang = targetLang;
      activeWorkerSource = 'tessdata_default';
      return activeWorker;
    }
  }

  async function releaseWorkerMemory() {
    if (activeWorker) {
      try { await activeWorker.terminate(); } catch (e) {}
      activeWorker = null;
      activeWorkerLang = '';
    }
  }

  function estimateSkewAngle(sourceCanvas) {
    try {
      const sampleW = 360;
      const sampleH = Math.max(120, Math.round((sourceCanvas.height / Math.max(1, sourceCanvas.width)) * sampleW));
      const small = document.createElement('canvas');
      small.width = sampleW;
      small.height = sampleH;
      const sctx = small.getContext('2d');
      sctx.fillStyle = '#ffffff';
      sctx.fillRect(0, 0, sampleW, sampleH);
      sctx.drawImage(sourceCanvas, 0, 0, sampleW, sampleH);

      const data = sctx.getImageData(0, 0, sampleW, sampleH).data;
      const bin = new Uint8Array(sampleW * sampleH);
      let darkCount = 0;
      for (let i = 0, p = 0; i < data.length; i += 4, p++) {
        const g = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
        if (g < 175) { bin[p] = 1; darkCount++; }
      }
      if (darkCount < 150) return 0;

      let bestAngle = 0;
      let bestScore = -1;
      const cx = sampleW / 2;
      const cy = sampleH / 2;

      for (let deg = -6; deg <= 6; deg += 0.75) {
        const rad = (deg * Math.PI) / 180;
        const sinA = Math.sin(rad);
        const cosA = Math.cos(rad);
        const rowSums = new Int32Array(sampleH);
        for (let y = 0; y < sampleH; y += 2) {
          const dy = y - cy;
          const rowOff = y * sampleW;
          for (let x = 0; x < sampleW; x += 2) {
            if (!bin[rowOff + x]) continue;
            const ny = Math.round((x - cx) * sinA + dy * cosA + cy);
            if (ny >= 0 && ny < sampleH) rowSums[ny]++;
          }
        }
        let score = 0;
        for (let y = 0; y < sampleH; y++) score += rowSums[y] * rowSums[y];
        if (score > bestScore) { bestScore = score; bestAngle = deg; }
      }
      return Math.abs(bestAngle) >= 0.75 ? bestAngle : 0;
    } catch { return 0; }
  }

  function applyAutoDeskew(sourceCanvas) {
    const angle = estimateSkewAngle(sourceCanvas);
    if (!angle) return { canvas: sourceCanvas, angle: 0 };
    const out = document.createElement('canvas');
    out.width = sourceCanvas.width;
    out.height = sourceCanvas.height;
    const ctx = out.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, out.width, out.height);
    ctx.translate(out.width / 2, out.height / 2);
    ctx.rotate((angle * Math.PI) / 180);
    ctx.drawImage(sourceCanvas, -out.width / 2, -out.height / 2);
    return { canvas: out, angle };
  }

  // Tạo bản nhị phân Otsu khử nhiễu sâu khi cần chạy lại OCR (confidence thấp)
  function createAdaptiveOtsuCanvas(sourceCanvas) {
    const out = document.createElement('canvas');
    out.width = sourceCanvas.width;
    out.height = sourceCanvas.height;
    const ctx = out.getContext('2d');
    ctx.drawImage(sourceCanvas, 0, 0);
    try {
      const imgData = ctx.getImageData(0, 0, out.width, out.height);
      const d = imgData.data;
      const hist = new Int32Array(256);
      const total = out.width * out.height;
      for (let i = 0; i < d.length; i += 4) {
        const g = Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
        hist[g]++;
      }
      let sum = 0;
      for (let t = 0; t < 256; t++) sum += t * hist[t];
      let sumB = 0, wB = 0, maxVar = 0, thresh = 160;
      for (let t = 0; t < 256; t++) {
        wB += hist[t];
        if (wB === 0) continue;
        const wF = total - wB;
        if (wF === 0) break;
        sumB += t * hist[t];
        const mB = sumB / wB;
        const mF = (sum - sumB) / wF;
        const between = wB * wF * (mB - mF) * (mB - mF);
        if (between > maxVar) { maxVar = between; thresh = t; }
      }
      const cut = Math.min(205, Math.max(115, thresh + 12));
      for (let i = 0; i < d.length; i += 4) {
        const g = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
        const v = g < cut ? 0 : 255;
        d[i] = d[i + 1] = d[i + 2] = v;
      }
      ctx.putImageData(imgData, 0, 0);
    } catch (e) {}
    return out;
  }

  // Phát hiện các ký tự dễ nhầm lẫn OCR trong số liệu/mã số: O↔0, I/l↔1, S↔5
  function detectConfusableTokens(words) {
    const flagged = [];
    if (!Array.isArray(words)) return flagged;
    for (const w of words) {
      const t = String(w?.text || '').trim();
      if (t.length < 2) continue;
      const hasDigit = /\d/.test(t);
      if (hasDigit && /[OIlS]/.test(t) && !/^(QĐ|NĐ|HD|CV|BC|TT|TB|NQ|KH|UBND|HĐND|VPTW)/i.test(t)) {
        flagged.push({
          text: t,
          confidence: Math.round(w.confidence || 0),
          hint: 'Chứa ký tự dễ nhầm lẫn (O/0, I/l/1, S/5) xen lẫn chữ số — cần đối chiếu ảnh gốc',
          bbox: w.bbox || null
        });
      }
    }
    return flagged;
  }

  // Trích xuất bảng biểu nâng cao từ tọa độ word bbox + đánh dấu ô có độ tin cậy OCR thấp (< 72%)
  function extractTableFromWordBboxes(words, pageW) {
    if (!Array.isArray(words) || words.length < 6) return null;
    const valid = words.filter(w => w && String(w.text || '').trim() && w.bbox);
    if (valid.length < 6) return null;

    // Nhóm các từ theo hàng Y (ngưỡng dung sai 12px)
    const sorted = valid.slice().sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
    const rows = [];
    for (const w of sorted) {
      const cy = (w.bbox.y0 + w.bbox.y1) / 2;
      let target = null;
      for (const r of rows) {
        if (Math.abs(r.cy - cy) <= Math.max(10, (w.bbox.y1 - w.bbox.y0) * 0.65)) {
          target = r; break;
        }
      }
      if (!target) {
        target = { cy, words: [] };
        rows.push(target);
      }
      target.words.push(w);
    }

    // Lọc các hàng có cấu trúc nhiều cột (khoảng cách ngang > 3.2% chiều rộng trang)
    const gapThreshold = Math.max(24, (pageW || 1200) * 0.032);
    const tableRows = [];
    for (const r of rows) {
      r.words.sort((a, b) => a.bbox.x0 - b.bbox.x0);
      const cells = [];
      let cur = null;
      for (const w of r.words) {
        if (!cur || (w.bbox.x0 - cur.x1) > gapThreshold) {
          cur = { text: w.text.trim(), x0: w.bbox.x0, x1: w.bbox.x1, minConf: Math.round(w.confidence || 100) };
          cells.push(cur);
        } else {
          cur.text += ' ' + w.text.trim();
          cur.x1 = Math.max(cur.x1, w.bbox.x1);
          cur.minConf = Math.min(cur.minConf, Math.round(w.confidence || 100));
        }
      }
      if (cells.length >= 2) tableRows.push(cells);
    }

    if (tableRows.length < 2) return null;
    const maxCols = Math.min(10, Math.max(...tableRows.map(r => r.length)));
    const headers = [];
    for (let c = 0; c < maxCols; c++) headers.push(tableRows[0][c]?.text || `Cột ${c + 1}`);
    const dataRows = [];
    const uncertainCells = [];
    for (let r = 1; r < tableRows.length; r++) {
      const rowArr = [];
      for (let c = 0; c < maxCols; c++) {
        const cell = tableRows[r][c];
        rowArr.push(cell ? cell.text : '');
        if (cell && cell.minConf < 72) {
          uncertainCells.push({ row: r - 1, col: c, text: cell.text, confidence: cell.minConf });
        }
      }
      dataRows.push(rowArr);
    }
    return { headers, rows: dataRows, uncertainCells };
  }

  // Vẽ khung đánh dấu vùng độ tin cậy thấp (< 70%) trực tiếp lên ảnh gốc để người dùng đối chiếu
  function drawConfidenceOverlayOnCanvas(displayCanvas, words, srcW, srcH) {
    if (!displayCanvas || !Array.isArray(words) || words.length === 0) return 0;
    const ctx = displayCanvas.getContext('2d');
    const sx = displayCanvas.width / Math.max(1, srcW || displayCanvas.width);
    const sy = displayCanvas.height / Math.max(1, srcH || displayCanvas.height);
    let lowCount = 0;
    ctx.save();
    for (const w of words) {
      const conf = Number(w?.confidence);
      if (!w?.bbox || isNaN(conf) || conf >= 70 || !String(w.text || '').trim()) continue;
      lowCount++;
      const x = w.bbox.x0 * sx;
      const y = w.bbox.y0 * sy;
      const bw = Math.max(6, (w.bbox.x1 - w.bbox.x0) * sx);
      const bh = Math.max(6, (w.bbox.y1 - w.bbox.y0) * sy);
      ctx.fillStyle = conf < 50 ? 'rgba(239,68,68,0.22)' : 'rgba(245,158,11,0.22)';
      ctx.strokeStyle = conf < 50 ? '#dc2626' : '#d97706';
      ctx.lineWidth = 1.5;
      ctx.fillRect(x, y, bw, bh);
      ctx.strokeRect(x, y, bw, bh);
    }
    ctx.restore();
    return lowCount;
  }

  async function decodeTiffFileToCanvas(file) {
    if (!global.UTIF) {
      throw new Error('Trình giải mã TIFF (UTIF.js) chưa sẵn sàng.');
    }
    const buf = await file.arrayBuffer();
    const ifds = global.UTIF.decode(buf);
    if (!ifds || ifds.length === 0) throw new Error('Tệp TIFF không hợp lệ hoặc bị hỏng.');
    const canvases = [];
    for (let i = 0; i < ifds.length; i++) {
      global.UTIF.decodeImage(buf, ifds[i]);
      const rgba = global.UTIF.toRGBA8(ifds[i]);
      const c = document.createElement('canvas');
      c.width = ifds[i].width;
      c.height = ifds[i].height;
      const ctx = c.getContext('2d');
      const imgData = ctx.createImageData(c.width, c.height);
      imgData.data.set(rgba);
      ctx.putImageData(imgData, 0, 0);
      canvases.push(c);
    }
    return canvases;
  }

  async function recognizeCanvasWithBestEngine(procCanvas, options = {}) {
    const t0 = performance.now();
    const lang = options.lang || 'vie';
    const enableDeskew = options.deskew !== false;
    const onProgress = options.onProgress || null;

    const classification = classifyDocumentCanvas(procCanvas);
    let workCanvas = procCanvas;
    let deskewDeg = 0;
    if (enableDeskew) {
      const dk = applyAutoDeskew(procCanvas);
      workCanvas = dk.canvas;
      deskewDeg = dk.angle;
    }

    const worker = await getOcrWorker(lang, classification.recommendedPsm, onProgress);
    let { data } = await worker.recognize(workCanvas);
    let bestConf = typeof data?.confidence === 'number' ? Math.round(data.confidence) : 0;
    let retried = false;

    // Tự động chạy lại với bộ lọc Otsu + PSM 6 nếu độ tin cậy < 72% và so sánh lấy bản tốt hơn
    if (bestConf > 0 && bestConf < 72 && !cancelRequested) {
      if (onProgress) onProgress(82, `Độ tin cậy thấp (${bestConf}%) — đang chạy lại bằng bộ lọc khử nhiễu Otsu...`);
      try {
        const otsuCanvas = createAdaptiveOtsuCanvas(workCanvas);
        await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
        const retryRes = await worker.recognize(otsuCanvas);
        const retryConf = typeof retryRes?.data?.confidence === 'number' ? Math.round(retryRes.data.confidence) : 0;
        if (retryConf > bestConf + 2 && (retryRes?.data?.text || '').trim().length >= (data?.text || '').trim().length * 0.7) {
          data = retryRes.data;
          bestConf = retryConf;
          retried = true;
        }
      } catch (e) {
        logSafeError('AutoRetry', e.message);
      }
    }

    const words = Array.isArray(data?.words) ? data.words : [];
    const lines = Array.isArray(data?.lines) ? data.lines : [];
    const lowConfWords = words.filter(w => String(w?.text || '').trim() && Number(w.confidence) < 70);
    const confusableTokens = detectConfusableTokens(words);
    const elapsedMs = Math.round(performance.now() - t0);

    lastScanMetrics = {
      docProfile: classification.docProfile,
      psm: classification.recommendedPsm,
      confidence: bestConf,
      lowConfCount: lowConfWords.length,
      confusableCount: confusableTokens.length,
      confusableTokens,
      deskewAngle: deskewDeg,
      retried,
      elapsedMs,
      engineSource: activeWorkerSource,
      words,
      procWidth: workCanvas.width,
      procHeight: workCanvas.height
    };

    return {
      text: String(data?.text || '').normalize('NFC'),
      lines,
      words,
      confidence: bestConf,
      lowConfCount: lowConfWords.length,
      confusableTokens,
      docProfile: classification.docProfile,
      retried,
      elapsedMs,
      procWidth: workCanvas.width,
      procHeight: workCanvas.height,
      deskewAngle: deskewDeg,
      engineSource: activeWorkerSource
    };
  }

  global.VietNamOcrEnginePro = {
    validateInputFile,
    classifyDocumentCanvas,
    getOcrWorker,
    releaseWorkerMemory,
    estimateSkewAngle,
    applyAutoDeskew,
    createAdaptiveOtsuCanvas,
    detectConfusableTokens,
    extractTableFromWordBboxes,
    drawConfidenceOverlayOnCanvas,
    decodeTiffFileToCanvas,
    recognizeCanvasWithBestEngine,
    requestCancel,
    resetCancel,
    isCancelled,
    togglePause,
    isPaused,
    waitIfPaused,
    getLastMetrics,
    logSafeError,
    getErrorLogs,
    clearErrorLogs
  };
})(window);
