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

    // Lưu mô hình bảng trung gian (kèm ảnh cắt từng ô & bbox) để người dùng soi & sửa từng ô trước khi xuất DOCX
    lastStructuredModel = {
      source: 'scanned_ocr',
      numRows: grid.numRows,
      numCols: grid.numCols,
      cells: grid.cells.map(cell => {
        const cLines = groupCellWordsToLines(cell.words);
        return {
          r: cell.r, c: cell.c, colspan: cell.colspan, rowspan: cell.rowspan,
          bbox: { x0: cell.x0, y0: cell.y0, x1: cell.x1, y1: cell.y1 },
          lines: cLines, text: cLines.join('\n'),
          confidence: cell.words.length ? Math.round(cell.words.reduce((a, w) => a + (Number(w.confidence) || 85), 0) / cell.words.length) : 100,
          cropDataUrl: cropCellToDataUrl(procCanvas, cell)
        };
      })
    };

    return {
      html: parts.join('\n'),
      numRows: grid.numRows,
      numCols: grid.numCols,
      cellCount: grid.cells.length,
      excelHeaders,
      excelRows
    };
  }

  // Cấu trúc dữ liệu trung gian lưu toàn bộ bảng & từng ô (ảnh gốc, bbox, nội dung, độ tin cậy) để kiểm tra & chỉnh sửa trước khi xuất DOCX
  let lastStructuredModel = null;

  function getLastStructuredTableModel() {
    return lastStructuredModel;
  }

  function joinCellPdfRuns(runs) {
    if (!runs || runs.length === 0) return '';
    runs.sort((a, b) => a.x0 - b.x0);
    let out = runs[0].str;
    let prevX1 = runs[0].x1;
    for (let i = 1; i < runs.length; i++) {
      const r = runs[i];
      const gap = r.x0 - prevX1;
      const minGap = Math.max(0.8, (r.fontSize || 12) * 0.12);
      const needSpace = (r.precededBySpace || gap > minGap) && !/\s$/.test(out) && !/^\s/.test(r.str);
      out += (needSpace ? ' ' : '') + r.str;
      prevX1 = Math.max(prevX1, r.x1);
    }
    return out.replace(/[ \t]{2,}/g, ' ').trim();
  }

  function groupPdfRunsToLines(runs) {
    if (!runs || runs.length === 0) return [];
    const sorted = runs.slice().sort((a, b) => a.cy - b.cy || a.x0 - b.x0);
    const rows = [];
    for (const r of sorted) {
      let target = null;
      for (const ln of rows) {
        if (Math.abs(ln.cy - r.cy) <= Math.max(6, (r.h || 14) * 0.55)) { target = ln; break; }
      }
      if (!target) { target = { cy: r.cy, items: [] }; rows.push(target); }
      target.items.push(r);
    }
    return rows.map(ln => joinCellPdfRuns(ln.items)).filter(Boolean);
  }

  function cropCellToDataUrl(sourceCanvas, cell) {
    try {
      const x = Math.max(0, Math.round(cell.x0));
      const y = Math.max(0, Math.round(cell.y0));
      const w = Math.max(8, Math.min(sourceCanvas.width - x, Math.round(cell.x1 - cell.x0)));
      const h = Math.max(8, Math.min(sourceCanvas.height - y, Math.round(cell.y1 - cell.y0)));
      const c = document.createElement('canvas');
      c.width = w; c.height = h;
      c.getContext('2d').drawImage(sourceCanvas, x, y, w, h, 0, 0, w, h);
      return c.toDataURL('image/png');
    } catch { return ''; }
  }

  // Tái tạo trang PDF gốc (có lớp văn bản) thành BẢNG THẬT giữ nguyên khoảng trắng, hàng, cột, ô gộp & ô trống
  function reconstructNativePdfPageWithRealTables(offCanvas, rawItems, viewportRender, viewport1x) {
    if (!offCanvas || !Array.isArray(rawItems) || rawItems.length === 0 || !global.pdfjsLib) return null;
    const grid = detectTableGridFromCanvas(offCanvas);
    if (!grid) return null;

    const mappedRuns = [];
    let pendingSpace = false;
    for (const it of rawItems) {
      if (typeof it.str !== 'string' || it.str.length === 0) continue;
      if (it.str.trim().length === 0) {
        pendingSpace = true;
        continue;
      }
      if (!it.transform) continue;
      const tx = global.pdfjsLib.Util.transform(viewportRender.transform, it.transform);
      const x0 = tx[4];
      const yBottom = tx[5];
      const h = Math.max(10, Math.hypot(tx[2], tx[3]) || 14);
      const w = Math.max(2, (it.width || 0) * viewportRender.scale);
      mappedRuns.push({
        str: it.str.normalize('NFC'),
        x0, x1: x0 + w,
        y0: yBottom - h, y1: yBottom,
        cx: x0 + w / 2, cy: yBottom - h / 2,
        h, fontSize: h,
        precededBySpace: pendingSpace
      });
      pendingSpace = false;
    }

    const aboveRuns = [];
    const belowRuns = [];
    let insideCount = 0;

    for (const r of mappedRuns) {
      if (r.cy < grid.tableTop - 4) { aboveRuns.push(r); continue; }
      if (r.cy > grid.tableBottom + 4) { belowRuns.push(r); continue; }
      let placed = false;
      for (const cell of grid.cells) {
        if (r.cx >= cell.x0 - 3 && r.cx <= cell.x1 + 3 && r.cy >= cell.y0 - 3 && r.cy <= cell.y1 + 3) {
          cell.words.push(r);
          insideCount++;
          placed = true;
          break;
        }
      }
      if (!placed) {
        if (r.cy < (grid.tableTop + grid.tableBottom) / 2) aboveRuns.push(r);
        else belowRuns.push(r);
      }
    }

    if (insideCount < 4) return null;

    const totalTableW = Math.max(1, grid.tableRight - grid.tableLeft);
    const rowsByIdx = Array.from({ length: grid.numRows }, () => []);
    grid.cells.forEach(cell => rowsByIdx[cell.r].push(cell));
    rowsByIdx.forEach(arr => arr.sort((a, b) => a.c - b.c));

    const structuredCells = [];
    const plainLinesOut = [];
    const aboveLines = groupPdfRunsToLines(aboveRuns);
    const belowLines = groupPdfRunsToLines(belowRuns);

    let htmlParts = [];
    aboveLines.forEach(ln => {
      plainLinesOut.push(ln);
      const isCenter = /^(LỊCH|BẢNG|DANH SÁCH|PHỤ LỤC|CỘNG HÒA|Độc lập|Tuần\s+\d+|Từ ngày)/i.test(ln) || ln === ln.toUpperCase();
      htmlParts.push(`<p align="${isCenter ? 'center' : 'left'}" style="text-align:${isCenter ? 'center' : 'left'}; text-indent:0; margin:0 0 4pt 0; font-size:14pt; ${ln === ln.toUpperCase() ? 'font-weight:bold;' : ''}">${escHtml(ln)}</p>`);
    });

    let tableHtml = `<table width="100%" cellpadding="4" cellspacing="0" border="1" style="width:100%; border-collapse:collapse; border:1px solid #000000; margin:8pt 0;">`;
    for (let r = 0; r < grid.numRows; r++) {
      const isHeaderRow = (r === 0);
      tableHtml += '<tr>';
      const rowPlainCells = [];
      for (const cell of rowsByIdx[r]) {
        const wPct = Math.max(4, Math.round(((cell.x1 - cell.x0) / totalTableW) * 1000) / 10);
        const cellLines = groupPdfRunsToLines(cell.words);
        const cellText = cellLines.join('\n');
        rowPlainCells.push(cellLines.join(' '));
        structuredCells.push({
          r: cell.r, c: cell.c, colspan: cell.colspan, rowspan: cell.rowspan,
          bbox: { x0: cell.x0, y0: cell.y0, x1: cell.x1, y1: cell.y1 },
          lines: cellLines, text: cellText, confidence: 100,
          cropDataUrl: cropCellToDataUrl(offCanvas, cell)
        });
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
      plainLinesOut.push(rowPlainCells.join(' | '));
    }
    tableHtml += '</table>';
    htmlParts.push(tableHtml);

    belowLines.forEach(ln => {
      plainLinesOut.push(ln);
      htmlParts.push(`<p style="margin:4pt 0; text-indent:0; font-size:13.5pt;">${escHtml(ln)}</p>`);
    });

    lastStructuredModel = {
      source: 'native_pdf',
      numRows: grid.numRows,
      numCols: grid.numCols,
      aboveLines,
      belowLines,
      cells: structuredCells
    };

    return {
      html: htmlParts.join('\n'),
      plainText: plainLinesOut.join('\n'),
      numRows: grid.numRows,
      numCols: grid.numCols
    };
  }

  // Cửa sổ kiểm tra & chỉnh sửa trực tiếp từng ô bảng (hiển thị ảnh cắt ô gốc + bbox + ô nhập sửa)
  function openCellInspectorModal() {
    if (!lastStructuredModel || !lastStructuredModel.cells || lastStructuredModel.cells.length === 0) {
      if (global.showToast) global.showToast('Chưa phát hiện cấu trúc bảng trên tài liệu hiện tại. Hãy tải tệp có bảng (như Lịch công tác) trước!');
      return;
    }
    let modal = document.getElementById('ocrCellInspectorModal');
    if (!modal) {
      modal = document.createElement('div');
      modal.id = 'ocrCellInspectorModal';
      modal.style.cssText = 'display:none;position:fixed;inset:0;z-index:10003;background:rgba(15,23,42,0.85);backdrop-filter:blur(4px);align-items:center;justify-content:center;padding:16px;';
      document.body.appendChild(modal);
    }

    const rowsHtml = lastStructuredModel.cells.map((cell, idx) => `
      <div style="display:grid;grid-template-columns:220px 1fr;gap:12px;background:#fff;border:1px solid #cbd5e1;border-radius:10px;padding:10px;align-items:start;">
        <div>
          <div style="font-size:0.78rem;font-weight:700;color:#1d4ed8;margin-bottom:4px;">Ô [Hàng ${cell.r + 1}, Cột ${cell.c + 1}] ${cell.colspan > 1 ? `· Gộp ngang ${cell.colspan}` : ''} ${cell.rowspan > 1 ? `· Gộp dọc ${cell.rowspan}` : ''}</div>
          ${cell.cropDataUrl ? `<img src="${cell.cropDataUrl}" style="max-width:100%;max-height:110px;border:1px solid #94a3b8;border-radius:6px;background:#f8fafc;">` : ''}
          <div style="font-size:0.72rem;color:#64748b;margin-top:3px;">bbox: (${Math.round(cell.bbox.x0)}, ${Math.round(cell.bbox.y0)}) → (${Math.round(cell.bbox.x1)}, ${Math.round(cell.bbox.y1)})</div>
        </div>
        <textarea data-cell-idx="${idx}" rows="3" style="width:100%;padding:8px;border:1px solid #cbd5e1;border-radius:8px;font-family:'Times New Roman',serif;font-size:14px;color:#0f172a;">${escHtml(cell.text)}</textarea>
      </div>`).join('');

    modal.innerHTML = `
      <div style="width:100%;max-width:960px;max-height:90vh;background:#f8fafc;color:#0f172a;border-radius:16px;display:flex;flex-direction:column;overflow:hidden;font-family:'Segoe UI',sans-serif;">
        <div style="background:linear-gradient(135deg,#1d4ed8,#2563eb);padding:14px 20px;display:flex;justify-content:space-between;align-items:center;color:#fff;">
          <div>
            <div style="font-weight:800;font-size:1.05rem;">🔬 Kiểm Tra &amp; Chỉnh Sửa Từng Ô Bảng (${lastStructuredModel.numRows} hàng × ${lastStructuredModel.numCols} cột)</div>
            <div style="font-size:0.82rem;color:#dbeafe;">Đối chiếu ảnh cắt gốc từng ô với kết quả nhận diện trước khi xuất Word (.docx)</div>
          </div>
          <button type="button" onclick="document.getElementById('ocrCellInspectorModal').style.display='none'" style="background:rgba(255,255,255,0.2);border:none;color:#fff;width:32px;height:32px;border-radius:8px;cursor:pointer;font-weight:700;">✕</button>
        </div>
        <div style="flex:1;overflow-y:auto;padding:14px 20px;display:flex;flex-direction:column;gap:10px;">${rowsHtml}</div>
        <div style="padding:12px 20px;background:#fff;border-top:1px solid #e2e8f0;display:flex;justify-content:flex-end;gap:10px;">
          <button type="button" onclick="document.getElementById('ocrCellInspectorModal').style.display='none'" style="padding:8px 16px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;color:#334155;font-weight:600;cursor:pointer;">Đóng</button>
          <button type="button" onclick="VietNamOcrTableGrid.applyEditedCellsToPreview()" style="padding:8px 18px;border-radius:8px;border:none;background:#2563eb;color:#fff;font-weight:700;cursor:pointer;">✅ Áp dụng chỉnh sửa vào Bản Word</button>
        </div>
      </div>`;
    modal.style.display = 'flex';
  }

  function applyEditedCellsToPreview() {
    const modal = document.getElementById('ocrCellInspectorModal');
    if (!modal || !lastStructuredModel) return;
    modal.querySelectorAll('textarea[data-cell-idx]').forEach(ta => {
      const idx = parseInt(ta.getAttribute('data-cell-idx'), 10);
      if (lastStructuredModel.cells[idx]) {
        lastStructuredModel.cells[idx].text = ta.value;
        lastStructuredModel.cells[idx].lines = ta.value.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      }
    });
    // Cập nhật lại bảng HTML trong #ocrFormattedPaper
    const paperEl = document.getElementById('ocrFormattedPaper');
    if (paperEl) {
      const tds = paperEl.querySelectorAll('table[border="1"] th, table[border="1"] td');
      lastStructuredModel.cells.forEach((cell, i) => {
        if (tds[i]) {
          tds[i].innerHTML = cell.lines.length > 0
            ? cell.lines.map(ln => `<p style="margin:0 0 3pt 0; text-indent:0; font-size:13pt;">${escHtml(ln)}</p>`).join('')
            : `<p style="margin:0; text-indent:0; font-size:13pt;">&nbsp;</p>`;
        }
      });
    }
    modal.style.display = 'none';
    if (global.showToast) global.showToast('✓ Đã cập nhật toàn bộ nội dung từng ô bảng vào bản xem trước Word!');
  }

  global.VietNamOcrTableGrid = {
    detectTableGridFromCanvas,
    reconstructPageWithRealTables,
    reconstructNativePdfPageWithRealTables,
    getLastStructuredTableModel,
    openCellInspectorModal,
    applyEditedCellsToPreview
  };
})(window);
