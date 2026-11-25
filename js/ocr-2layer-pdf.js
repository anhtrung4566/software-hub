/**
 * VietNamOCR 3.5 Pro — Bộ Xuất PDF 2 Lớp (Searchable PDF chuẩn Unicode Tiếng Việt)
 *
 * Cấu trúc 2 lớp thực thụ trong tệp PDF xuất ra:
 *  - Lớp 1 (Hiển thị): Ảnh trang scan gốc (JPEG chất lượng cao), giữ nguyên 100% con dấu đỏ, chữ ký, bảng biểu, bố cục.
 *  - Lớp 2 (Chữ ẩn tiếng Việt đúng tọa độ): Nhúng phông TrueType Unicode hỗ trợ đầy đủ dấu tiếng Việt,
 *    vẽ từng dòng chữ tại đúng tọa độ (x, y) từ OCR bbox với opacity = 0 để bôi đen, copy và Ctrl+F tìm kiếm tiếng Việt chuẩn xác.
 */
(function (global) {
  'use strict';

  const UNICODE_FONT_URLS = [
    'https://cdn.jsdelivr.net/gh/notofonts/latin-greek-cyrillic@main/fonts/NotoSerif/hinted/ttf/NotoSerif-Regular.ttf',
    'https://cdn.jsdelivr.net/gh/googlefonts/roboto@main/src/hinted/Roboto-Regular.ttf'
  ];

  let cachedFontBytes = null;

  // Lưu cache kết quả OCR (ảnh canvas + tọa độ bbox dòng) cho ảnh đơn lẻ hoặc TẤT CẢ các trang PDF
  const pageStore = {
    singlePage: null, // { imageDataUrl, width, height, text, lines: [{ text, bbox }] }
    multiPages: []    // Array of { pageNum, imageDataUrl, width, height, text, lines }
  };

  function clearMultiPageStore() {
    pageStore.multiPages = [];
  }

  function buildScaledLines(tesseractLines, dstW, dstH, bboxSrcW, bboxSrcH) {
    const sx = (bboxSrcW && bboxSrcW > 0) ? (dstW / bboxSrcW) : 1;
    const sy = (bboxSrcH && bboxSrcH > 0) ? (dstH / bboxSrcH) : 1;
    const cleanLines = [];
    if (Array.isArray(tesseractLines) && tesseractLines.length > 0) {
      tesseractLines.forEach(l => {
        const t = String(l && l.text || '').replace(/\r?\n/g, ' ').trim();
        if (t && l.bbox) {
          cleanLines.push({
            text: t.normalize('NFC'),
            bbox: {
              x0: (Number(l.bbox.x0) || 0) * sx,
              y0: (Number(l.bbox.y0) || 0) * sy,
              x1: (Number(l.bbox.x1) || dstW) * sx,
              y1: (Number(l.bbox.y1) || 20) * sy
            }
          });
        }
      });
    }
    return cleanLines;
  }

  function recordSinglePageOcr(canvas, text, tesseractLines, bboxSrcW, bboxSrcH) {
    if (!canvas) return;
    try {
      const dstW = canvas.width || 1240;
      const dstH = canvas.height || 1754;
      pageStore.singlePage = {
        imageDataUrl: canvas.toDataURL('image/jpeg', 0.90),
        width: dstW,
        height: dstH,
        text: String(text || '').normalize('NFC'),
        lines: buildScaledLines(tesseractLines, dstW, dstH, bboxSrcW, bboxSrcH)
      };
    } catch (e) {}
  }

  function recordPdfPageOcr(pageNum, canvas, text, lines, bboxSrcW, bboxSrcH) {
    if (!canvas) return;
    try {
      const dstW = canvas.width || 1240;
      const dstH = canvas.height || 1754;
      const entry = {
        pageNum: Number(pageNum) || (pageStore.multiPages.length + 1),
        imageDataUrl: canvas.toDataURL('image/jpeg', 0.88),
        width: dstW,
        height: dstH,
        text: String(text || '').normalize('NFC'),
        lines: buildScaledLines(lines, dstW, dstH, bboxSrcW, bboxSrcH)
      };
      const idx = pageStore.multiPages.findIndex(p => p.pageNum === entry.pageNum);
      if (idx >= 0) pageStore.multiPages[idx] = entry;
      else pageStore.multiPages.push(entry);
      pageStore.multiPages.sort((a, b) => a.pageNum - b.pageNum);
    } catch (e) {}
  }

  async function loadVietnameseTtfBytes() {
    if (cachedFontBytes) return cachedFontBytes;
    for (const url of UNICODE_FONT_URLS) {
      try {
        const res = await fetch(url);
        if (res.ok) {
          const buf = await res.arrayBuffer();
          if (buf && buf.byteLength > 20000) {
            cachedFontBytes = new Uint8Array(buf);
            return cachedFontBytes;
          }
        }
      } catch (e) {}
    }
    return null;
  }

  function dataUrlToUint8Array(dataUrl) {
    const base64 = dataUrl.split(',')[1] || '';
    const bin = atob(base64);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  function stripVietnameseFallback(s) {
    return String(s || '')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .replace(/[^\x20-\x7E]/g, ' ');
  }

  function buildFallbackLines(text, widthPx, heightPx) {
    const rawLines = String(text || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    if (!rawLines.length) return [];
    const marginX = Math.round(widthPx * 0.08);
    const topY = Math.round(heightPx * 0.06);
    const usableH = heightPx * 0.88;
    const stepY = Math.max(18, Math.min(34, usableH / Math.max(1, rawLines.length)));
    return rawLines.map((t, idx) => ({
      text: t.normalize('NFC'),
      bbox: {
        x0: marginX,
        y0: topY + idx * stepY,
        x1: widthPx - marginX,
        y1: topY + (idx + 0.82) * stepY
      }
    }));
  }

  async function addTwoLayerPage(pdfDoc, embeddedFont, hasUnicodeFont, pageData) {
    // Khổ trang A4 tỉ lệ theo ảnh scan gốc (chuẩn chiều rộng A4 = 595.28 pt)
    const imgW = pageData.width || 1240;
    const imgH = pageData.height || 1754;
    const pageW = 595.28;
    const pageH = Math.max(200, (imgH / imgW) * pageW);
    const scaleX = pageW / imgW;
    const scaleY = pageH / imgH;

    const page = pdfDoc.addPage([pageW, pageH]);

    // LỚP 1: Ảnh trang scan gốc (hiển thị nguyên vẹn con dấu đỏ, chữ ký, bảng biểu)
    const jpgBytes = dataUrlToUint8Array(pageData.imageDataUrl);
    const jpgImg = await pdfDoc.embedJpg(jpgBytes);
    page.drawImage(jpgImg, { x: 0, y: 0, width: pageW, height: pageH });

    // LỚP 2: Lớp chữ ẩn tiếng Việt đặt đúng tọa độ bbox (opacity: 0 -> bôi đen & Ctrl+F được)
    const lines = (Array.isArray(pageData.lines) && pageData.lines.length > 0)
      ? pageData.lines
      : buildFallbackLines(pageData.text, imgW, imgH);

    for (const ln of lines) {
      const rawTxt = String(ln.text || '').trim();
      if (!rawTxt) continue;
      const drawTxt = hasUnicodeFont ? rawTxt : stripVietnameseFallback(rawTxt);
      const box = ln.bbox || { x0: imgW * 0.08, y0: 40, x1: imgW * 0.92, y1: 60 };
      const boxW = Math.max(8, (box.x1 - box.x0) * scaleX);
      const boxH = Math.max(8, (box.y1 - box.y0) * scaleY);
      let fontSize = Math.max(6, Math.min(28, boxH * 0.84));
      // Tự co giãn fontSize theo bề rộng bbox thực để bôi đen & Ctrl+F khớp đúng vị trí từng chữ trên ảnh scan
      try {
        const measuredW = embeddedFont.widthOfTextAtSize(drawTxt, fontSize);
        if (measuredW > boxW * 1.06 && measuredW > 0) {
          fontSize = Math.max(5.5, fontSize * (boxW / measuredW));
        }
      } catch (e) {}
      const pdfX = Math.max(4, Math.min(pageW - 20, box.x0 * scaleX));
      // Hệ tọa độ PDF gốc ở góc dưới trái, ngược trục Y với Canvas
      const pdfY = Math.max(4, Math.min(pageH - fontSize, pageH - (box.y1 * scaleY) + (boxH * 0.12)));

      try {
        page.drawText(drawTxt, {
          x: pdfX,
          y: pdfY,
          size: fontSize,
          font: embeddedFont,
          color: global.PDFLib.rgb(0, 0, 0),
          opacity: 0 // Chữ ẩn hoàn toàn trên ảnh scan gốc nhưng bôi đen, copy & Ctrl+F bình thường
        });
      } catch (e) {
        // Bỏ qua ký tự điều khiển lạ không có glyph
      }
    }
  }

  async function exportSearchableTwoLayerPdf() {
    const statusEl = document.getElementById('ocrStatusText');
    const progressFill = document.getElementById('ocrProgressFill');
    const canvas = document.getElementById('ocrPreviewCanvas');
    const textOutput = document.getElementById('ocrResultText');

    if (!global.PDFLib) {
      alert('Thư viện pdf-lib chưa sẵn sàng (cần kết nối Internet để tải thư viện tạo PDF 2 lớp).');
      return;
    }

    const setStatus = (msg, pct) => {
      if (statusEl) statusEl.textContent = msg;
      if (progressFill && typeof pct === 'number') progressFill.style.width = `${pct}%`;
    };

    try {
      setStatus('Đang khởi tạo bộ xuất PDF 2 lớp (nhúng phông Unicode Tiếng Việt)...', 10);
      const pdfDoc = await global.PDFLib.PDFDocument.create();
      pdfDoc.setTitle('Tài liệu PDF 2 lớp — VietNamOCR 3.5 Pro');
      pdfDoc.setProducer('VietNamOCR 3.5 Pro (2-Layer Searchable PDF)');

      let embeddedFont = null;
      let hasUnicodeFont = false;
      if (global.fontkit) {
        pdfDoc.registerFontkit(global.fontkit);
        const ttfBytes = await loadVietnameseTtfBytes();
        if (ttfBytes) {
          embeddedFont = await pdfDoc.embedFont(ttfBytes, { subset: true });
          hasUnicodeFont = true;
        }
      }
      if (!embeddedFont) {
        embeddedFont = await pdfDoc.embedFont(global.PDFLib.StandardFonts.Helvetica);
      }

      const pagesToBuild = [];

      // Trường hợp 1: Đã có cache đầy đủ TẤT CẢ các trang PDF từ lúc tải tệp -> xuất tức thì toàn bộ 100% số trang!
      if (pageStore.multiPages.length > 0 && (!global.loadedPdfDoc || pageStore.multiPages.length >= global.loadedPdfDoc.numPages)) {
        for (let i = 0; i < pageStore.multiPages.length; i++) {
          setStatus(`Đang đóng gói PDF 2 lớp — trang ${i + 1}/${pageStore.multiPages.length}...`, Math.round(20 + ((i + 1) / pageStore.multiPages.length) * 65));
          pagesToBuild.push(pageStore.multiPages[i]);
        }
      } else if (typeof global.loadedPdfDoc !== 'undefined' && global.loadedPdfDoc && global.loadedPdfDoc.numPages > 0) {
        // Trường hợp 2: Đang mở tệp PDF nhiều trang bằng PDF.js -> duyệt từng trang 1..numPages
        const total = global.loadedPdfDoc.numPages;
        for (let p = 1; p <= total; p++) {
          const cached = pageStore.multiPages.find(x => x.pageNum === p);
          if (cached) {
            pagesToBuild.push(cached);
            continue;
          }
          setStatus(`Đang xử lý PDF 2 lớp — trang ${p}/${total}...`, Math.round(15 + (p / total) * 70));
          const pdfPage = await global.loadedPdfDoc.getPage(p);
          const vp = pdfPage.getViewport({ scale: 2.0 });
          const off = document.createElement('canvas');
          off.width = Math.round(vp.width);
          off.height = Math.round(vp.height);
          const ctx = off.getContext('2d');
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, off.width, off.height);
          await pdfPage.render({ canvasContext: ctx, viewport: vp }).promise;

          let pageLines = [];
          let pageText = '';
          const tc = await pdfPage.getTextContent();
          if (tc && Array.isArray(tc.items) && tc.items.length > 0) {
            tc.items.forEach(it => {
              const s = String(it.str || '').trim();
              if (!s || !it.transform) return;
              const tx = global.pdfjsLib.Util.transform(vp.transform, it.transform);
              const x0 = tx[4];
              const yBottom = tx[5];
              const h = Math.max(10, Math.hypot(tx[2], tx[3]) || 14);
              pageLines.push({
                text: s.normalize('NFC'),
                bbox: { x0, y0: yBottom - h, x1: x0 + (it.width * vp.scale), y1: yBottom }
              });
              pageText += s + ' ';
            });
          }

          const badChars = (pageText.match(/\uFFFD/g) || []).length;
          if ((pageLines.length < 2 || badChars > 3) && global.Tesseract) {
            setStatus(`Trang ${p}/${total} là ảnh scan — đang OCR tiếng Việt tạo lớp chữ ẩn...`, Math.round(15 + (p / total) * 70));
            const ocrRes = await global.Tesseract.recognize(off, 'vie', { preserve_interword_spaces: '1' });
            pageText = String(ocrRes?.data?.text || '').normalize('NFC');
            const tLines = ocrRes?.data?.lines || [];
            pageLines = tLines.filter(l => l && l.text && l.text.trim() && l.bbox).map(l => ({
              text: l.text.replace(/\r?\n/g, ' ').trim().normalize('NFC'),
              bbox: l.bbox
            }));
          }

          pagesToBuild.push({
            imageDataUrl: off.toDataURL('image/jpeg', 0.88),
            width: off.width,
            height: off.height,
            text: pageText,
            lines: pageLines
          });
        }
      } else if (pageStore.singlePage) {
        // Trường hợp 2: Ảnh scan đã nhận diện OCR (có sẵn tọa độ bbox)
        const sp = Object.assign({}, pageStore.singlePage);
        if (textOutput && textOutput.value.trim()) sp.text = textOutput.value.trim().normalize('NFC');
        pagesToBuild.push(sp);
      } else if (canvas && canvas.width > 0 && canvas.height > 0) {
        // Trường hợp 3: Ảnh đang hiển thị trên canvas nhưng chưa có bbox -> chạy Tesseract nhanh để lấy tọa độ từng dòng
        let tLines = [];
        let txt = textOutput ? textOutput.value.trim() : '';
        if (global.Tesseract) {
          setStatus('Đang quét tọa độ từng dòng chữ để dựng lớp chữ ẩn...', 45);
          const ocrRes = await global.Tesseract.recognize(canvas, 'vie');
          if (!txt) txt = String(ocrRes?.data?.text || '').normalize('NFC');
          tLines = (ocrRes?.data?.lines || []).filter(l => l && l.text && l.text.trim() && l.bbox).map(l => ({
            text: l.text.replace(/\r?\n/g, ' ').trim().normalize('NFC'),
            bbox: l.bbox
          }));
        }
        pagesToBuild.push({
          imageDataUrl: canvas.toDataURL('image/jpeg', 0.90),
          width: canvas.width,
          height: canvas.height,
          text: txt,
          lines: tLines
        });
      }

      if (!pagesToBuild.length) {
        setStatus('⚠ Vui lòng tải file PDF hoặc ảnh scan lên trước khi xuất PDF 2 lớp.', 0);
        return;
      }

      for (let i = 0; i < pagesToBuild.length; i++) {
        setStatus(`Đang đóng gói PDF 2 lớp — trang ${i + 1}/${pagesToBuild.length}...`, Math.round(85 + ((i + 1) / pagesToBuild.length) * 12));
        await addTwoLayerPage(pdfDoc, embeddedFont, hasUnicodeFont, pagesToBuild[i]);
      }

      const pdfBytes = await pdfDoc.save();
      const blob = new Blob([pdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `VietNamOCR_PDF_2Lop_${new Date().toISOString().slice(0, 10)}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);

      setStatus(`✓ Đã xuất thành công PDF 2 lớp (${pagesToBuild.length} trang · ${hasUnicodeFont ? 'Unicode Tiếng Việt đầy đủ' : 'ASCII dự phòng'})! Mở bằng Chrome/Edge/Adobe để Ctrl+F và copy chữ trực tiếp trên ảnh scan gốc.`, 100);
    } catch (err) {
      setStatus('⚠ Lỗi xuất PDF 2 lớp: ' + (err && err.message ? err.message : err), 0);
    }
  }

  global.VietNamOcr2LayerPdf = {
    recordSinglePageOcr,
    recordPdfPageOcr,
    clearMultiPageStore,
    exportSearchableTwoLayerPdf
  };
  global.exportOcrToTwoLayerPdf = exportSearchableTwoLayerPdf;
})(typeof window !== 'undefined' ? window : globalThis);
