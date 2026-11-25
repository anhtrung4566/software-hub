/**
 * VietNamOCR Pro — Bộ Nhận Diện Lưới Bảng Biểu & Ô Gộp Từ Ảnh Scan/PDF (js/ocr-table-grid.js)
 *
 * Quy trình thực tế (KHÔNG chèn ảnh giả lập bảng, KHÔNG gộp thành văn bản phẳng):
 * 1. Quét pixel đường kẻ ngang (horizontal rulings) và đường kẻ dọc (vertical rulings) trên ảnh scan.
 * 2. Gom cụm tọa độ đường kẻ thành lưới cột xs[] và lưới hàng ys[].
 * 3. Kiểm tra từng vách ngăn dọc/ngang giữa các ô liền kề để phát hiện chính xác:
 *    - Ô gộp ngang (colspan)
 *    - Ô gộp dọc (rowspan)
 *    - Ô trống giữ nguyên bố cục
 * 4. Ánh xạ từng từ OCR (words bbox) vào đúng ô (r, c) theo tọa độ tâm (cx, cy), gom dòng trong từng ô riêng biệt.
 * 5. Dựng cấu trúc HTML <table> thật (border, width%, colspan, rowspan, nhiều đoạn <p> trong ô) để hiển thị & xuất thẳng sang <w:tbl> của DOCX!
 */
(function (global) {
  'use strict';

  function escHtml(s) {
    return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function clusterCoords(coords, tolerance) {
    if (!coords || coords.length === 0) return [];
    const sorted = coords.slice().sort((a, b) => a - b);
    const groups = [[sorted[0]]];
    for (let i = 1; i < sorted.length; i++) {
      const last = groups[groups.length - 1];
      if (sorted[i] - last[last.length - 1] <= tolerance) {
        last.push(sorted[i]);
      } else {
        groups.push([sorted[i]]);
      }
    }
    return groups.map(g => Math.round(g.reduce((a, b) => a + b, 0) / g.length));
  }

  // Phát hiện lưới đường kẻ ngang & dọc thực tế trên canvas
  function detectTableGridFromCanvas(sourceCanvas) {
    if (!sourceCanvas || sourceCanvas.width < 120 || sourceCanvas.height < 120) return null;
    try {
      const w = sourceCanvas.width;
      const h = sourceCanvas.height;
      const ctx = sourceCanvas.getContext('2d');
      const data = ctx.getImageData(0, 0, w, h).data;

      // Nhị phân hóa nhanh để dò đường kẻ (độ xám < 175)
      const isDark = (x, y) => {
        const idx = (y * w + x) * 4;
        return (0.299 * data[idx] + 0.587 * data[idx + 1] + 0.114 * data[idx + 2]) < 175;
      };

      // 1. Dò các hàng có đoạn kẻ ngang liên tục dài >= 28% chiều rộng trang
      const minHRun = Math.round(w * 0.28);
      const rawYs = [];
      for (let y = 4; y < h - 4; y += 2) {
        let run = 0, maxRun = 0;
        for (let x = 2; x < w - 2; x += 2) {
          if (isDark(x, y)) { run += 2; if (run > maxRun) maxRun = run; }
          else run = 0;
        }
        if (maxRun >= minHRun) rawYs.push(y);
      }
      const ys = clusterCoords(rawYs, Math.max(8, Math.round(h * 0.008)));
      if (ys.length < 3) return null;

      const tableTop = ys[0];
      const tableBottom = ys[ys.length - 1];
      const tableH = tableBottom - tableTop;
      if (tableH < h * 0.10) return null;

      // 2. Dò các cột có đoạn kẻ dọc liên tục bên trong [tableTop, tableBottom] dài >= 22% chiều cao bảng
      const minVRun = Math.max(40, Math.round(tableH * 0.22));
      const rawXs = [];
      for (let x = 4; x < w - 4; x += 2) {
        let run = 0, maxRun = 0;
        for (let y = tableTop; y <= tableBottom; y += 2) {
          if (isDark(x, y)) { run += 2; if (run > maxRun) maxRun = run; }
          else run = 0;
        }
        if (maxRun >= minVRun) rawXs.push(x);
      }
      const xs = clusterCoords(rawXs, Math.max(8, Math.round(w * 0.008)));
      if (xs.length < 3) return null;

      const tableLeft = xs[0];
      const tableRight = xs[xs.length - 1];
      if ((tableRight - tableLeft) < w * 0.30) return null;

      // 3. Kiểm tra vách ngăn giữa các ô để phát hiện chính xác ô gộp ngang (colspan) & ô gộp dọc (rowspan)
      const numRows = ys.length - 1;
      const numCols = xs.length - 1;

      const hasVertBorder = (colLineIdx, r) => {
        const x = xs[colLineIdx];
        const y0 = ys[r] + 4;
        const y1 = ys[r + 1] - 4;
        if (y1 <= y0) return true;
        let dark = 0, total = 0;
        for (let y = y0; y <= y1; y += 2) {
          total++;
          if (isDark(x, y) || (x > 1 && isDark(x - 1, y)) || (x + 1 < w && isDark(x + 1, y))) dark++;
        }
        return total > 0 && (dark / total) >= 0.42;
      };

      const hasHorizBorder = (rowLineIdx, c) => {
        const y = ys[rowLineIdx];
        const x0 = xs[c] + 4;
        const x1 = xs[c + 1] - 4;
        if (x1 <= x0) return true;
        let dark = 0, total = 0;
        for (let x = x0; x <= x1; x += 2) {
          total++;
          if (isDark(x, y) || (y > 1 && isDark(x, y - 1)) || (y + 1 < h && isDark(x, y + 1))) dark++;
        }
        return total > 0 && (dark / total) >= 0.42;
      };

      // Xây ma trận ô có hỗ trợ colspan & rowspan
      const visited = Array.from({ length: numRows }, () => new Array(numCols).fill(false));
      const cells = [];

      for (let r = 0; r < numRows; r++) {
        for (let c = 0; c < numCols; c++) {
          if (visited[r][c]) continue;
          let colspan = 1;
          while (c + colspan < numCols && !visited[r][c + colspan] && !hasVertBorder(c + colspan, r)) {
            colspan++;
          }
          let rowspan = 1;
          while (r + rowspan < numRows) {
            let openAll = true;
            for (let cc = c; cc < c + colspan; cc++) {
              if (visited[r + rowspan][cc] || hasHorizBorder(r + rowspan, cc)) {
                openAll = false; break;
              }
            }
            if (!openAll) break;
            rowspan++;
          }
          for (let rr = r; rr < r + rowspan; rr++) {
            for (let cc = c; cc < c + colspan; cc++) visited[rr][cc] = true;
          }
          cells.push({
            r, c, colspan, rowspan,
            x0: xs[c], y0: ys[r], x1: xs[c + colspan], y1: ys[r + rowspan],
            words: []
          });
        }
      }

      return { xs, ys, numRows, numCols, tableLeft, tableRight, tableTop, tableBottom, cells };
    } catch {
      return null;
    }
  }

  // Nhóm các từ trong 1 ô thành nhiều dòng riêng biệt theo trục Y (giữ nguyên xuống dòng trong ô)
  function groupCellWordsToLines(words) {
    if (!words || words.length === 0) return [];
    const sorted = words.slice().sort((a, b) => a.bbox.y0 - b.bbox.y0 || a.bbox.x0 - b.bbox.x0);
    const lines = [];
    for (const w of sorted) {
      const cy = (w.bbox.y0 + w.bbox.y1) / 2;
      const h = Math.max(8, w.bbox.y1 - w.bbox.y0);
      let row = null;
      for (const ln of lines) {
        if (Math.abs(ln.cy - cy) <= h * 0.65) { row = ln; break; }
      }
      if (!row) { row = { cy, words: [] }; lines.push(row); }
      row.words.push(w);
    }
    return lines.map(ln => {
      ln.words.sort((a, b) => a.bbox.x0 - b.bbox.x0);
      const text = ln.words.map(w => String(w.text || '').trim()).filter(Boolean).join(' ');
      return (typeof global.cleanAndFixVietnameseOcrText === 'function')
        ? global.cleanAndFixVietnameseOcrText(text)
        : text;
    }).filter(Boolean);
  }

  // Tái tạo toàn bộ trang: Phần trên bảng + BẢNG WORD THẬT (giữ đủ hàng, cột, viền, ô gộp, ô trống) + Phần dưới bảng
  function reconstructPageWithRealTables(procCanvas, words, lines, forceMode = 'auto') {
    if (!procCanvas || !Array.isArray(words) || words.length === 0) return null;
    const grid = detectTableGridFromCanvas(procCanvas);
    if (!grid) return null;

    const aboveLines = [];
    const belowLines = [];

    // Phân loại các dòng ngoài bảng (trên và dưới bảng)
    if (Array.isArray(lines)) {
      for (const ln of lines) {
        if (!ln || !ln.bbox || !String(ln.text || '').trim()) continue;
        const cy = (ln.bbox.y0 + ln.bbox.y1) / 2;
        if (cy < grid.tableTop - 4) aboveLines.push(ln);
        else if (cy > grid.tableBottom + 4) belowLines.push(ln);
      }
    }

    // Ánh xạ từng từ vào đúng ô (r, c) theo tọa độ tâm — tuyệt đối không trộn 2 ô khác nhau
    let mappedWordsCount = 0;
    for (const w of words) {
      if (!w || !w.bbox || !String(w.text || '').trim()) continue;
      const cx = (w.bbox.x0 + w.bbox.x1) / 2;
      const cy = (w.bbox.y0 + w.bbox.y1) / 2;
      if (cx < grid.tableLeft - 4 || cx > grid.tableRight + 4 || cy < grid.tableTop - 4 || cy > grid.tableBottom + 4) {
        continue;
      }
      for (const cell of grid.cells) {
        if (cx >= cell.x0 - 3 && cx <= cell.x1 + 3 && cy >= cell.y0 - 3 && cy <= cell.y1 + 3) {
          cell.words.push(w);
          mappedWordsCount++;
          break;
        }
      }
    }

    if (mappedWordsCount < 4) return null;

    const totalTableW = Math.max(1, grid.tableRight - grid.tableLeft);
    const rowsByIdx = Array.from({ length: grid.numRows }, () => []);
    grid.cells.forEach(cell => rowsByIdx[cell.r].push(cell));
    rowsByIdx.forEach(arr => arr.sort((a, b) => a.c - b.c));

    // Dựng bảng HTML chuẩn (tương thích 100% với VietNamDocxExporter để xuất <w:tbl>, <w:gridSpan>, <w:vMerge>)
    let tableHtml = `<table width="100%" cellpadding="4" cellspacing="0" border="1" style="width:100%; border-collapse:collapse; border:1px solid #000000; margin:8pt 0;">`;
    const excelHeaders = [];
    const excelRows = [];

    for (let r = 0; r < grid.numRows; r++) {
      const isHeaderRow = (r === 0);
      tableHtml += '<tr>';
      const flatRow = [];
      for (const cell of rowsByIdx[r]) {
        const wPct = Math.max(4, Math.round(((cell.x1 - cell.x0) / totalTableW) * 1000) / 10);
        const cellLines = groupCellWordsToLines(cell.words);
        flatRow.push(cellLines.join(' | '));
        const tag = isHeaderRow ? 'th' : 'td';
        const bgStyle = isHeaderRow ? 'background:#f1f5f9; font-weight:bold; text-align:center;' : 'text-align:left;';
        const csAttr = cell.colspan > 1 ? ` colspan="${cell.colspan}"` : '';
        const rsAttr = cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : '';
        const innerHtml = cellLines.length > 0
          ? cellLines.map(ln => `<p style="margin:0 0 3pt 0; text-indent:0; font-size:13pt; ${isHeaderRow ? 'font-weight:bold; text-align:center;' : 'text-align:left;'}">${escHtml(ln)}</p>`).join('')
          : `<p style="margin:0; text-indent:0; font-size:13pt;">&nbsp;</p>`;
        tableHtml += `<${tag}${csAttr}${rsAttr} width="${wPct}%" style="width:${wPct}%; border:1px solid #000000; padding:4pt 6pt; vertical-align:top; ${bgStyle}">${innerHtml}</${tag}>`;
      }
      tableHtml += '</tr>';
      if (isHeaderRow) flatRow.forEach(h => excelHeaders.push(h || 'Cột'));
      else excelRows.push(flatRow);
    }
    tableHtml += '</table>';

    const parts = [];
    if (aboveLines.length > 0 && typeof global.reconstructLayoutFromTesseractBboxes === 'function') {
      parts.push(global.reconstructLayoutFromTesseractBboxes(aboveLines, procCanvas.width, procCanvas.height, forceMode));
    }
    parts.push(tableHtml);
    if (belowLines.length > 0 && typeof global.reconstructExactLayoutFromPlainText === 'function') {
      const belowTxt = belowLines.map(l => l.text || '').join('\n');
      parts.push(global.reconstructExactLayoutFromPlainText(belowTxt, forceMode));
    }

    return {
      html: parts.join('\n'),
      numRows: grid.numRows,
      numCols: grid.numCols,
      cellCount: grid.cells.length,
      excelHeaders,
      excelRows
    };
  }

  global.VietNamOcrTableGrid = {
    detectTableGridFromCanvas,
    reconstructPageWithRealTables
  };
})(window);
