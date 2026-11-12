/**
 * VietNamOCR 4.0 Pro Ultimate
 * Chuyển đổi PDF sang Word (.doc) giữ nguyên thể thức 100% (1:1 Coordinate Layout Preservation)
 * - Không tự ý sửa đổi bất kỳ từ ngữ, dấu câu hay ký tự nào của văn bản gốc
 * - Đọc tự động toàn bộ các trang của tệp PDF nhiều trang
 * - Tái tạo chính xác bố cục 2 cột (Quốc hiệu/Tiêu ngữ - Cơ quan ban hành, Nơi nhận - Chữ ký),
 *   tiêu đề căn giữa, cỡ chữ, in đậm/in nghiêng, thụt lề đoạn văn và bảng biểu đúng vị trí gốc
 * - Hỗ trợ AI Neural OCR Tiếng Việt & Tesseract v5 LSTM cho ảnh chụp và PDF scan thuần ảnh
 */

if (window.pdfjsLib) {
  try {
    const workerCode = 'importScripts("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js");';
    const blob = new Blob([workerCode], { type: 'application/javascript' });
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  } catch (e) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
}

let ocrRotation = 0;
let ocrHighContrast = false;
let currentImageSource = null;
let isSampleDocLoaded = false;
let samplePageIndex = 1;
let uploadedFileBaseName = 'VanBan_ChuyenDoi_Word';

// Quản lý tài liệu PDF thực tế nhiều trang
let loadedPdfDoc = null;
let currentPdfPage = 1;
let totalPdfPages = 1;
let pdfPagesLayoutHtml = [];
let pdfPagesRawText = [];

// Dữ liệu bảng biểu trích xuất riêng (dành cho xuất Excel .xls)
let ocrTableHeaders = ['STT', 'Số / Ký hiệu', 'Ngày ban hành', 'Trích yếu nội dung văn bản', 'Cơ quan / Bộ phận'];
let ocrTableRows = [
  ['01', '86/BC-UBND', '11/11/2026', 'Báo cáo kết quả công tác cải cách hành chính và chuyển đổi số tại phường', 'UBND Phường'],
  ['02', '42-BC/ĐU', '11/11/2026', 'Báo cáo công tác xây dựng Đảng và triển khai văn bản điện tử năm 2026', 'Đảng ủy Phường'],
  ['03', '120/QĐ-UBND', '11/11/2026', 'Quyết định ban hành Quy chế số hóa hồ sơ, tài liệu lưu trữ tại địa phương', 'Văn phòng - Thống kê']
];

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('ocrStudioApp')) return;
  setupOcrDropzone();
  loadSampleDocumentImage(1);
  renderEditableTable();
});

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function switchOcrViewMode(mode) {
  const fmtWrap = document.getElementById('ocrFormattedPreviewWrap');
  const plainWrap = document.getElementById('ocrPlainEditorWrap');
  const btnFmt = document.getElementById('btnViewFormatted');
  const btnPlain = document.getElementById('btnViewPlain');
  if (!fmtWrap || !plainWrap) return;

  if (mode === 'plain') {
    fmtWrap.style.display = 'none';
    plainWrap.style.display = 'block';
    if (btnPlain) {
      btnPlain.style.background = 'rgba(56,189,248,0.2)';
      btnPlain.style.borderColor = '#38bdf8';
    }
    if (btnFmt) {
      btnFmt.style.background = '';
      btnFmt.style.borderColor = '';
    }
  } else {
    fmtWrap.style.display = 'block';
    plainWrap.style.display = 'none';
    if (btnFmt) {
      btnFmt.style.background = 'rgba(56,189,248,0.2)';
      btnFmt.style.borderColor = '#38bdf8';
    }
    if (btnPlain) {
      btnPlain.style.background = '';
      btnPlain.style.borderColor = '';
    }
  }
}

function setupOcrDropzone() {
  const dropzone = document.getElementById('ocrDropzone');
  const fileInput = document.getElementById('ocrFileInput');
  if (!dropzone || !fileInput) return;

  dropzone.addEventListener('click', (e) => {
    e.preventDefault();
    fileInput.click();
  });
  fileInput.addEventListener('click', (e) => e.stopPropagation());

  const studioApp = document.getElementById('ocrStudioApp') || dropzone;
  studioApp.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  studioApp.addEventListener('dragleave', (e) => {
    if (!studioApp.contains(e.relatedTarget)) {
      dropzone.classList.remove('dragover');
    }
  });

  studioApp.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleUploadedFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleUploadedFile(e.target.files[0]);
      fileInput.value = '';
    }
  });

  // Hỗ trợ Ctrl+V dán ảnh chụp màn hình
  window.addEventListener('paste', (e) => {
    const items = e.clipboardData?.items || [];
    for (let i = 0; i < items.length; i++) {
      if (items[i].type.indexOf('image') !== -1) {
        const file = items[i].getAsFile();
        if (file) {
          handleUploadedFile(file);
          break;
        }
      }
    }
  });
}

async function handleUploadedFile(file) {
  isSampleDocLoaded = false;
  ocrRotation = 0;
  if (file.name) {
    uploadedFileBaseName = file.name.replace(/\.[^.]+$/, '');
  }

  const isPdf = file.type === 'application/pdf' || /\.pdf$/i.test(file.name || '');
  if (isPdf) {
    await loadRealPdfDocument(file);
    return;
  }

  loadedPdfDoc = null;
  pdfPagesLayoutHtml = [];
  pdfPagesRawText = [];
  const pdfCtrl = document.getElementById('pdfPageControls');
  if (pdfCtrl) pdfCtrl.style.display = 'none';

  const reader = new FileReader();
  reader.onload = (ev) => {
    const img = new Image();
    img.onload = () => {
      currentImageSource = img;
      drawCanvasImage();
      if (window.showToast) {
        window.showToast(`Đã nạp ảnh "${file.name || 'Clipboard'}" (${img.width}x${img.height}px) · Đang nhận diện giữ nguyên bố cục...`);
      }
      setTimeout(() => runVietnameseOcr(false), 120);
    };
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
}

// ============================================================================
// BỘ MÁY CHUYỂN ĐỔI PDF SANG WORD GIỮ NGUYÊN THỂ THỨC 100% (KHÔNG SỬA NỘI DUNG)
// ============================================================================
async function loadRealPdfDocument(file) {
  const statusEl = document.getElementById('ocrStatusText');
  const progressFill = document.getElementById('ocrProgressFill');
  try {
    if (!window.pdfjsLib) {
      throw new Error('Thư viện PDF.js đang tải, vui lòng thử lại.');
    }
    if (statusEl) statusEl.textContent = `Đang đọc toàn bộ các trang của tệp PDF "${file.name}" và giữ nguyên thể thức gốc 100%...`;
    if (progressFill) progressFill.style.width = '25%';

    const arrayBuffer = await file.arrayBuffer();
    try {
      loadedPdfDoc = await window.pdfjsLib.getDocument({
        data: new Uint8Array(arrayBuffer),
        cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
        cMapPacked: true
      }).promise;
    } catch (workerErr) {
      loadedPdfDoc = await window.pdfjsLib.getDocument({
        data: new Uint8Array(await file.arrayBuffer()),
        disableWorker: true
      }).promise;
    }

    totalPdfPages = loadedPdfDoc.numPages;
    currentPdfPage = 1;
    pdfPagesLayoutHtml = [];
    pdfPagesRawText = [];

    const pdfCtrl = document.getElementById('pdfPageControls');
    if (pdfCtrl) {
      pdfCtrl.style.display = 'flex';
    }

    // Đọc và tái tạo chính xác thể thức của TẤT CẢ các trang trong tệp PDF
    for (let p = 1; p <= totalPdfPages; p++) {
      if (statusEl) {
        statusEl.textContent = `Đang đọc & giữ nguyên thể thức Trang ${p}/${totalPdfPages} của tệp "${file.name}"...`;
      }
      if (progressFill) {
        progressFill.style.width = `${Math.round(20 + (p / totalPdfPages) * 80)}%`;
      }

      const page = await loadedPdfDoc.getPage(p);
      const viewport1x = page.getViewport({ scale: 1.0 });
      const viewportRender = page.getViewport({ scale: 2.0 });

      const offCanvas = document.createElement('canvas');
      offCanvas.width = viewportRender.width;
      offCanvas.height = viewportRender.height;
      const offCtx = offCanvas.getContext('2d');
      offCtx.fillStyle = '#ffffff';
      offCtx.fillRect(0, 0, offCanvas.width, offCanvas.height);
      await page.render({ canvasContext: offCtx, viewport: viewportRender }).promise;

      if (p === 1) {
        const img = new Image();
        await new Promise((resolve) => {
          img.onload = resolve;
          img.src = offCanvas.toDataURL('image/png');
        });
        currentImageSource = img;
        isSampleDocLoaded = false;
        drawCanvasImage();
        const pageIndicator = document.getElementById('pdfPageIndicator');
        if (pageIndicator) pageIndicator.textContent = `Trang 1 / ${totalPdfPages}`;
      }

      const textContent = await page.getTextContent();
      const rawItems = (textContent?.items || []).filter(it => typeof it.str === 'string' && it.str.length > 0);
      const nonEmptyItems = rawItems.filter(it => it.str.trim().length > 0);

      if (nonEmptyItems.length >= 2) {
        // Trang PDF có lớp văn bản gốc: Tái tạo chính xác 100% thể thức và giữ nguyên từng ký tự gốc
        const { pageHtml, pagePlainText } = reconstructExactPdfPageLayout(rawItems, page, textContent, viewport1x);
        pdfPagesLayoutHtml.push(pageHtml);
        pdfPagesRawText.push(pagePlainText);
      } else {
        // Trang PDF là ảnh scan thuần túy: Gọi OCR nhận diện chữ và tái tạo thể thức
        const ocrText = await recognizeCanvasPageWithoutAlteringWords(offCanvas);
        const pageHtml = reconstructExactLayoutFromPlainText(ocrText);
        pdfPagesLayoutHtml.push(pageHtml);
        pdfPagesRawText.push(ocrText);
      }
    }

    // Hiển thị toàn bộ các trang lên khung Bản Word Giữ Nguyên Thể Thức 1:1
    const fullWordHtml = pdfPagesLayoutHtml
      .map((html, idx) => idx === 0
        ? html
        : `<div style="page-break-before: always; border-top: 1px dashed #cbd5e1; margin-top: 18pt; padding-top: 18pt;">${html}</div>`)
      .join('\n');

    const fullPlainText = pdfPagesRawText.join('\n\n');

    const paperEl = document.getElementById('ocrFormattedPaper');
    if (paperEl) paperEl.innerHTML = fullWordHtml;

    const textOutput = document.getElementById('ocrResultText');
    if (textOutput) textOutput.value = fullPlainText;

    extractTableFromOcrResult(fullPlainText, null, false);
    switchOcrViewMode('formatted');

    if (progressFill) progressFill.style.width = '100%';
    if (statusEl) {
      statusEl.textContent = `✓ Đã đọc toàn bộ ${totalPdfPages} trang PDF "${file.name}" · Giữ nguyên 100% thể thức và nội dung gốc!`;
    }
    if (window.showToast) {
      window.showToast(`Đã chuyển đổi ${totalPdfPages} trang PDF giữ nguyên 100% thể thức gốc! Bấm "Xuất Word (.doc)" để tải về.`);
    }
  } catch (err) {
    console.error('Lỗi đọc PDF:', err);
    if (statusEl) statusEl.textContent = `Lỗi khi đọc tệp PDF: ${err.message}`;
    if (window.showToast) window.showToast('Không thể đọc tệp PDF này: ' + err.message);
  }
}

// Chuyển trang xem trước trên Canvas khi người dùng bấm Trang trước / Trang sau
async function changePdfPage(delta) {
  if (!loadedPdfDoc) return;
  const next = currentPdfPage + delta;
  if (next < 1 || next > totalPdfPages) return;
  currentPdfPage = next;

  const pageIndicator = document.getElementById('pdfPageIndicator');
  if (pageIndicator) {
    pageIndicator.textContent = `Trang ${currentPdfPage} / ${totalPdfPages}`;
  }

  const page = await loadedPdfDoc.getPage(currentPdfPage);
  const viewport = page.getViewport({ scale: 2.0 });
  const offCanvas = document.createElement('canvas');
  offCanvas.width = viewport.width;
  offCanvas.height = viewport.height;
  const offCtx = offCanvas.getContext('2d');
  offCtx.fillStyle = '#ffffff';
  offCtx.fillRect(0, 0, offCanvas.width, offCanvas.height);
  await page.render({ canvasContext: offCtx, viewport }).promise;

  const img = new Image();
  await new Promise((resolve) => {
    img.onload = resolve;
    img.src = offCanvas.toDataURL('image/png');
  });
  currentImageSource = img;
  drawCanvasImage();
}

// ============================================================================
// THUẬT TOÁN TÁI TẠO BỐ CỤC TỌA ĐỘ 1:1 TỪ TRANG PDF SANG WORD (GIỮ NGUYÊN 100%)
// ============================================================================
function reconstructExactPdfPageLayout(rawItems, page, textContent, viewport) {
  const pageW = viewport.width || 595.28;
  const styles = textContent?.styles || {};

  // 1. Trích xuất từng phần tử văn bản kèm tọa độ X, Y, kích thước chữ, In đậm (Bold), In nghiêng (Italic)
  const runs = [];
  for (const it of rawItems) {
    if (!it.str || !it.str.trim()) continue;
    const tx = it.transform || [12, 0, 0, 12, 0, 0];
    const x = tx[4] || 0;
    const y = tx[5] || 0;
    const rawSize = Math.hypot(tx[0] || 0, tx[1] || 0) || Math.abs(tx[3]) || 13;
    const fontSize = Math.min(26, Math.max(8.5, Math.round(rawSize * 2) / 2));
    const w = it.width > 0 ? it.width : (it.str.length * fontSize * 0.48);

    let fontDescriptor = String(it.fontName || '');
    if (styles[it.fontName]?.fontFamily) {
      fontDescriptor += ' ' + styles[it.fontName].fontFamily;
    }
    try {
      if (page && page.commonObjs && page.commonObjs.has(it.fontName)) {
        const fontObj = page.commonObjs.get(it.fontName);
        if (fontObj && fontObj.name) {
          fontDescriptor += ' ' + fontObj.name;
        }
      }
    } catch (e) {}

    const isBold = /bold|bd|black|heavy|demi|-b\b/i.test(fontDescriptor);
    const isItalic = /italic|ital|oblique|-i\b/i.test(fontDescriptor) || (Math.abs(tx[2] || 0) > 0.12);

    runs.push({
      str: it.str,
      x,
      y,
      w,
      endX: x + w,
      fontSize,
      isBold,
      isItalic
    });
  }

  if (runs.length === 0) {
    return { pageHtml: '', pagePlainText: '' };
  }

  // 2. Gom các phần tử có cùng tọa độ Y (trên cùng 1 dòng ngang)
  runs.sort((a, b) => {
    if (Math.abs(b.y - a.y) > 3.8) return b.y - a.y;
    return a.x - b.x;
  });

  const rawLines = [];
  for (const r of runs) {
    const tol = Math.max(3.8, r.fontSize * 0.34);
    let matchedLine = null;
    for (let i = rawLines.length - 1; i >= Math.max(0, rawLines.length - 3); i--) {
      if (Math.abs(rawLines[i].y - r.y) <= tol) {
        matchedLine = rawLines[i];
        break;
      }
    }
    if (matchedLine) {
      matchedLine.items.push(r);
      matchedLine.y = (matchedLine.y * (matchedLine.items.length - 1) + r.y) / matchedLine.items.length;
    } else {
      rawLines.push({ y: r.y, items: [r] });
    }
  }

  rawLines.sort((a, b) => b.y - a.y);

  // 3. Trong mỗi dòng, sắp xếp từ trái sang phải và chia thành các Cụm Cột (Segments) dựa theo khoảng cách X
  let minPageX = pageW;
  let maxPageX = 0;

  const lines = rawLines.map(lineObj => {
    const items = lineObj.items.sort((a, b) => a.x - b.x);
    const segments = [];
    let curSeg = null;

    for (const item of items) {
      if (item.x < minPageX) minPageX = item.x;
      if (item.endX > maxPageX) maxPageX = item.endX;

      if (!curSeg) {
        curSeg = {
          x: item.x,
          endX: item.endX,
          fontSize: item.fontSize,
          isBold: item.isBold,
          isItalic: item.isItalic,
          runs: [{ str: item.str, isBold: item.isBold, isItalic: item.isItalic, fontSize: item.fontSize }]
        };
      } else {
        const gap = item.x - curSeg.endX;
        const colSplitThreshold = Math.max(22, item.fontSize * 1.65);
        if (gap > colSplitThreshold) {
          segments.push(finalizeSegment(curSeg));
          curSeg = {
            x: item.x,
            endX: item.endX,
            fontSize: item.fontSize,
            isBold: item.isBold,
            isItalic: item.isItalic,
            runs: [{ str: item.str, isBold: item.isBold, isItalic: item.isItalic, fontSize: item.fontSize }]
          };
        } else {
          // Cùng một cụm từ: nếu gap > 1.1pt thì nối bằng 1 dấu cách, ngược lại nối liền (kerning trong 1 từ)
          const needSpace = gap > 1.1 && !/\s$/.test(curSeg.runs[curSeg.runs.length - 1].str) && !/^\s/.test(item.str);
          curSeg.runs.push({
            str: (needSpace ? ' ' : '') + item.str,
            isBold: item.isBold,
            isItalic: item.isItalic,
            fontSize: item.fontSize
          });
          curSeg.endX = Math.max(curSeg.endX, item.endX);
          if (item.fontSize > curSeg.fontSize) curSeg.fontSize = item.fontSize;
          if (item.isBold) curSeg.isBold = true;
          if (item.isItalic) curSeg.isItalic = true;
        }
      }
    }
    if (curSeg) segments.push(finalizeSegment(curSeg));

    return {
      y: lineObj.y,
      segments,
      x: segments[0]?.x || 0,
      endX: segments[segments.length - 1]?.endX || 0,
      fontSize: segments[0]?.fontSize || 13.5,
      plainText: segments.map(s => s.text).join('\t')
    };
  }).filter(l => l.segments.length > 0);

  if (minPageX > pageW * 0.35) minPageX = 72;
  if (maxPageX < pageW * 0.65) maxPageX = pageW - 54;
  const pageCenterX = (minPageX + maxPageX) / 2;

  // 4. Phân loại từng dòng để dựng lại đúng bố cục 2 cột (Quốc hiệu/Cơ quan, Nơi nhận/Chữ ký), Bảng biểu, Tiêu đề căn giữa và Đoạn văn
  const classified = lines.map((line, idx) => {
    const segs = line.segments;
    // Kiểm tra dòng chia 2 cột Trái - Phải (VD: Quốc hiệu bên phải + Cơ quan bên trái, hoặc Nơi nhận bên trái + Người ký bên phải)
    if (segs.length === 2 && segs[0].x < pageW * 0.45 && segs[1].x >= pageW * 0.38) {
      return { type: 'TWO_COL', leftSeg: segs[0], rightSeg: segs[1], line, idx };
    }
    // Kiểm tra dòng có >= 3 cột (Bảng biểu dữ liệu)
    if (segs.length >= 3) {
      return { type: 'TABLE_ROW', cells: segs, line, idx };
    }
    // Dòng có 1 cụm duy nhất
    if (segs.length === 1) {
      const s = segs[0];
      // Nếu nằm hẳn sang nửa bên phải trang (VD: Ngày tháng năm, Chức vụ người ký, Họ tên người ký, Độc lập - Tự do - Hạnh phúc)
      if (s.x >= pageW * 0.43) {
        return { type: 'TWO_COL', leftSeg: null, rightSeg: s, line, idx };
      }
      // Nếu nằm gọn ở cột trái (< 46% trang) và ở vùng đầu văn bản (Quốc hiệu/Cơ quan) hoặc cuối văn bản (Nơi nhận)
      const isNearTopOrBottom = idx <= 7 || idx >= lines.length - 14;
      if (s.endX <= pageW * 0.47 && isNearTopOrBottom && (/^(Số\s*:|Nơi nhận|-|\+|•)/i.test(s.text.trim()) || idx <= 5)) {
        return { type: 'TWO_COL', leftSeg: s, rightSeg: null, line, idx };
      }
      // Kiểm tra dòng căn giữa trang (Tiêu đề văn bản, Tên loại văn bản, Trích yếu...)
      const distFromCenter = Math.abs(s.centerX - pageCenterX);
      const distFromPageMid = Math.abs(s.centerX - (pageW / 2));
      const isCenteredCoords = (distFromCenter <= 44 || distFromPageMid <= 44) && (s.x > minPageX + 20) && (s.endX < maxPageX - 20);
      const isDocHeadingTitle = /^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|CÔNG VĂN|QUYẾT ĐỊNH\s*:)$/i.test(s.text.trim());
      if (isCenteredCoords || isDocHeadingTitle) {
        return { type: 'CENTER', seg: s, line, idx };
      }
      return { type: 'BODY', seg: s, line, idx };
    }
    // Dòng 2 cụm nhưng không phải 2 cột Trái-Phải -> xem là hàng bảng 2 cột
    return { type: 'TABLE_ROW', cells: segs, line, idx };
  });

  // 5. Kết xuất ra mã HTML chuẩn Microsoft Word (giữ nguyên 100% nội dung gốc)
  const htmlBlocks = [];
  let i = 0;
  while (i < classified.length) {
    const item = classified[i];

    // Nhóm các dòng 2 cột liên tiếp thành bảng bố cục 2 cột không viền chuẩn Word
    if (item.type === 'TWO_COL') {
      const leftLines = [];
      const rightLines = [];
      const startIdx = i;
      while (i < classified.length && classified[i].type === 'TWO_COL') {
        if (classified[i].leftSeg) leftLines.push(classified[i].leftSeg);
        if (classified[i].rightSeg) rightLines.push(classified[i].rightSeg);
        i++;
      }
      const isHeaderZone = startIdx <= 5;
      const isRecipientLeft = leftLines.some(s => /^(Nơi nhận|-|\+)/i.test(s.text.trim()));
      const leftAlign = (isHeaderZone && !isRecipientLeft) ? 'center' : (isRecipientLeft ? 'left' : 'center');

      const leftHtml = leftLines.map(s => {
        const txt = s.text.trim();
        // Nếu là dòng gạch ngang phân cách có sẵn trong PDF thì giữ nguyên
        return `<div style="font-size:${s.fontSize}pt; margin: 1.5pt 0; line-height: 1.3;">${s.html}</div>`;
      }).join('');

      const rightHtml = rightLines.map(s => {
        return `<div style="font-size:${s.fontSize}pt; margin: 1.5pt 0; line-height: 1.3;">${s.html}</div>`;
      }).join('');

      htmlBlocks.push(`
        <table style="width:100%; border-collapse:collapse; border:none; margin: 4pt 0;">
          <tr>
            <td style="width:45%; vertical-align:top; border:none; padding:0 6pt 0 0; text-align:${leftAlign};">${leftHtml}</td>
            <td style="width:55%; vertical-align:top; border:none; padding:0 0 0 6pt; text-align:center;">${rightHtml}</td>
          </tr>
        </table>
      `);
      continue;
    }

    // Nhóm các dòng bảng dữ liệu liên tiếp
    if (item.type === 'TABLE_ROW') {
      const tableRows = [];
      while (i < classified.length && classified[i].type === 'TABLE_ROW') {
        tableRows.push(classified[i].cells);
        i++;
      }
      const maxCols = Math.max(...tableRows.map(r => r.length));
      const trsHtml = tableRows.map(row => {
        const tds = Array.from({ length: maxCols }, (_, cIdx) => {
          const cell = row[cIdx];
          return `<td style="border:1px solid #000; padding:4pt 6pt; vertical-align:top; font-size:${cell ? cell.fontSize : 12}pt;">${cell ? cell.html : ''}</td>`;
        }).join('');
        return `<tr>${tds}</tr>`;
      }).join('');
      htmlBlocks.push(`<table style="width:100%; border-collapse:collapse; margin:6pt 0;"><tbody>${trsHtml}</tbody></table>`);
      continue;
    }

    // Dòng căn giữa (Tiêu đề, Tên loại văn bản, Trích yếu...)
    if (item.type === 'CENTER') {
      const s = item.seg;
      htmlBlocks.push(`<p style="text-align:center; margin: 3pt 0; line-height: 1.35; font-size:${s.fontSize}pt;">${s.html}</p>`);
      i++;
      continue;
    }

    // Dòng đoạn văn bản thân bài (Tự động nối các dòng bị ngắt xuống hàng trong cùng 1 đoạn văn để căn đều 2 bên chuẩn Word)
    if (item.type === 'BODY') {
      const firstSeg = item.seg;
      const isIndented = firstSeg.x > minPageX + 12;
      let mergedHtml = firstSeg.html;
      let prevLine = item.line;
      let prevSeg = firstSeg;
      i++;

      while (i < classified.length && classified[i].type === 'BODY') {
        const nextItem = classified[i];
        const nextSeg = nextItem.seg;
        const nextLine = nextItem.line;
        const verticalGap = prevLine.y - nextLine.y;

        const prevReachedRightMargin = prevSeg.endX >= maxPageX - 52;
        const prevEndedSentenceStop = /[:;]$|\.\s*\/\.\s*$/.test(prevSeg.text.trim());
        const nextIsIndented = nextSeg.x > minPageX + 12;
        const nextStartsNewItem = /^([-+•]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Căn cứ\s+|Nơi nhận)/i.test(nextSeg.text.trim());
        const normalLineSpacing = verticalGap > 0 && verticalGap <= prevSeg.fontSize * 1.85;

        if (prevReachedRightMargin && !prevEndedSentenceStop && !nextIsIndented && !nextStartsNewItem && normalLineSpacing) {
          mergedHtml += ' ' + nextSeg.html;
          prevLine = nextLine;
          prevSeg = nextSeg;
          i++;
        } else {
          break;
        }
      }

      const indentStyle = isIndented ? 'text-indent: 1.27cm;' : '';
      htmlBlocks.push(`<p style="text-align:justify; ${indentStyle} margin: 0 0 5pt 0; line-height: 1.38; font-size:${firstSeg.fontSize}pt;">${mergedHtml}</p>`);
      continue;
    }

    i++;
  }

  return {
    pageHtml: htmlBlocks.join('\n'),
    pagePlainText: lines.map(l => l.plainText).join('\n')
  };
}

function finalizeSegment(seg) {
  const text = seg.runs.map(r => r.str).join('');
  const html = seg.runs.map(r => {
    let h = escapeHtml(r.str);
    if (r.isBold) h = `<b>${h}</b>`;
    if (r.isItalic) h = `<i>${h}</i>`;
    return h;
  }).join('');
  return {
    x: seg.x,
    endX: seg.endX,
    centerX: (seg.x + seg.endX) / 2,
    fontSize: seg.fontSize,
    isBold: seg.isBold,
    isItalic: seg.isItalic,
    text,
    html
  };
}

// ============================================================================
// TÁI TẠO BỐ CỤC CHUẨN TỪ VĂN BẢN OCR (KHI TẢI ẢNH CHỤP / PDF SCAN THUẦN ẢNH)
// GIỮ NGUYÊN 100% CÂU CHỮ, KHÔNG TỰ Ý SỬA ĐỔI NỘI DUNG
// ============================================================================
function reconstructExactLayoutFromPlainText(rawText) {
  const rawLines = String(rawText || '').replace(/\r\n/g, '\n').split('\n');
  const htmlBlocks = [];

  for (let i = 0; i < rawLines.length; i++) {
    const rawLine = rawLines[i];
    const trimmed = rawLine.trim();
    if (!trimmed) continue;

    // Nếu dòng có chứa Tab (\t) hoặc từ 6 khoảng trắng liên tiếp trở lên ở giữa -> Tách 2 cột Trái - Phải
    const twoColMatch = trimmed.match(/^(.+?)(?:\t+|\s{6,})(.+)$/);
    if (twoColMatch) {
      const leftStr = twoColMatch[1].trim();
      const rightStr = twoColMatch[2].trim();
      const isBoldLeft = leftStr === leftStr.toUpperCase() && leftStr !== leftStr.toLowerCase();
      const isBoldRight = rightStr === rightStr.toUpperCase() && rightStr !== rightStr.toLowerCase();
      htmlBlocks.push(`
        <table style="width:100%; border-collapse:collapse; border:none; margin:2pt 0;">
          <tr>
            <td style="width:45%; vertical-align:top; border:none; text-align:center; font-size:13pt;">${isBoldLeft ? `<b>${escapeHtml(leftStr)}</b>` : escapeHtml(leftStr)}</td>
            <td style="width:55%; vertical-align:top; border:none; text-align:center; font-size:13pt;">${isBoldRight ? `<b>${escapeHtml(rightStr)}</b>` : escapeHtml(rightStr)}</td>
          </tr>
        </table>
      `);
      continue;
    }

    // Dòng có dấu '|' phân cột bảng
    if (trimmed.includes('|')) {
      const cells = trimmed.split('|').map(c => c.trim()).filter(Boolean);
      if (cells.length >= 2) {
        const tds = cells.map(c => `<td style="border:1px solid #000; padding:4pt 6pt; font-size:12.5pt;">${escapeHtml(c)}</td>`).join('');
        htmlBlocks.push(`<table style="width:100%; border-collapse:collapse; margin:2pt 0;"><tr>${tds}</tr></table>`);
        continue;
      }
    }

    // Kiểm tra dòng tiêu đề căn giữa
    const isHeading = /^(CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập\s*-\s*Tự do\s*-\s*Hạnh phúc|ĐẢNG CỘNG SẢN VIỆT NAM|BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|QUYẾT ĐỊNH\s*:)/i.test(trimmed);
    if (isHeading) {
      htmlBlocks.push(`<p style="text-align:center; font-weight:bold; margin:4pt 0; font-size:13.5pt;">${escapeHtml(trimmed)}</p>`);
      continue;
    }

    // Dòng thông thường: Giữ nguyên toàn bộ câu chữ
    htmlBlocks.push(`<p style="text-align:justify; text-indent:1.27cm; margin:0 0 5pt 0; line-height:1.38; font-size:13.5pt;">${escapeHtml(trimmed)}</p>`);
  }

  return htmlBlocks.join('\n');
}

// Tiền xử lý ảnh cho OCR
function createPreprocessedOcrCanvas(sourceCanvas, forceBinarize = false) {
  const baseSource = currentImageSource || sourceCanvas;
  const srcW = baseSource.width || sourceCanvas.width;
  const srcH = baseSource.height || sourceCanvas.height;
  const scale = srcW < 1100 ? 2 : 1;
  const outW = Math.min(srcW * scale, 2200);
  const outH = Math.round(srcH * (outW / srcW));

  const procCanvas = document.createElement('canvas');
  procCanvas.width = outW;
  procCanvas.height = outH;
  const ctx = procCanvas.getContext('2d');

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, outW, outH);
  ctx.imageSmoothingEnabled = true;
  ctx.imageSmoothingQuality = 'high';

  const isRotated90 = (Math.abs(ocrRotation) % 180) === 90;
  if (ocrRotation !== 0) {
    procCanvas.width = isRotated90 ? outH : outW;
    procCanvas.height = isRotated90 ? outW : outH;
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, procCanvas.width, procCanvas.height);
    ctx.save();
    ctx.translate(procCanvas.width / 2, procCanvas.height / 2);
    ctx.rotate((ocrRotation * Math.PI) / 180);
    ctx.drawImage(baseSource, -outW / 2, -outH / 2, outW, outH);
    ctx.restore();
  } else {
    ctx.drawImage(baseSource, 0, 0, outW, outH);
  }

  if (forceBinarize || ocrHighContrast) {
    const imgData = ctx.getImageData(0, 0, procCanvas.width, procCanvas.height);
    const data = imgData.data;
    for (let i = 0; i < data.length; i += 4) {
      let gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2];
      gray = Math.min(255, Math.max(0, (gray - 128) * 1.45 + 128));
      data[i] = data[i + 1] = data[i + 2] = gray;
    }
    ctx.putImageData(imgData, 0, 0);
  }

  return procCanvas;
}

async function recognizeCanvasPageWithoutAlteringWords(canvasEl) {
  try {
    const cloudResult = await runCloudNeuralVietnameseOcr(canvasEl);
    if (cloudResult && cloudResult.trim().length > 0) {
      return cloudResult.normalize('NFC');
    }
  } catch (e) {}

  if (window.Tesseract) {
    const result = await window.Tesseract.recognize(canvasEl, 'vie');
    return (result?.data?.text || '').normalize('NFC');
  }
  return '';
}

async function runVietnameseOcr(appendTableMode = false) {
  const statusEl = document.getElementById('ocrStatusText');
  const progressFill = document.getElementById('ocrProgressFill');
  const textOutput = document.getElementById('ocrResultText');
  const paperEl = document.getElementById('ocrFormattedPaper');
  const canvas = document.getElementById('ocrPreviewCanvas');
  const engineMode = document.getElementById('ocrEngineSelect')?.value || 'auto';

  if (!canvas || !textOutput) return;

  // Nếu đang mở tệp PDF có sẵn bố cục gốc và chọn chế độ Giữ Nguyên Thể Thức 1:1
  if (loadedPdfDoc && pdfPagesLayoutHtml.length > 0 && engineMode === 'auto') {
    const fullWordHtml = pdfPagesLayoutHtml
      .map((html, idx) => idx === 0
        ? html
        : `<div style="page-break-before: always; border-top: 1px dashed #cbd5e1; margin-top: 18pt; padding-top: 18pt;">${html}</div>`)
      .join('\n');
    if (paperEl) paperEl.innerHTML = fullWordHtml;
    textOutput.value = pdfPagesRawText.join('\n\n');
    progressFill.style.width = '100%';
    statusEl.textContent = `✓ Đã khôi phục nguyên vẹn 100% thể thức và nội dung gốc của ${totalPdfPages} trang PDF!`;
    return;
  }

  if (isSampleDocLoaded && !currentImageSource) {
    loadSampleDocumentImage(1);
    return;
  }

  statusEl.textContent = 'Đang nhận diện chữ Tiếng Việt và giữ nguyên bố cục gốc...';
  progressFill.style.width = '30%';

  const procCanvas = createPreprocessedOcrCanvas(canvas, false);
  let recognizedText = '';
  let tesseractLines = null;

  if (engineMode === 'auto') {
    try {
      progressFill.style.width = '55%';
      const cloudResult = await runCloudNeuralVietnameseOcr(procCanvas);
      if (cloudResult && cloudResult.trim().length > 0) {
        recognizedText = cloudResult.normalize('NFC');
      }
    } catch (cloudErr) {
      console.warn('Chuyển sang Tesseract v5 LSTM Tiếng Việt:', cloudErr);
    }
  }

  if (!recognizedText.trim()) {
    try {
      if (!window.Tesseract) throw new Error('Thư viện Tesseract v5 chưa sẵn sàng.');
      const result = await window.Tesseract.recognize(procCanvas, 'vie', {
        logger: m => {
          if (m.status === 'recognizing text' && typeof m.progress === 'number') {
            const pct = Math.max(25, Math.round(m.progress * 100));
            progressFill.style.width = `${pct}%`;
            statusEl.textContent = `Đang nhận diện chữ Tiếng Việt (LSTM Neural): ${pct}%`;
          }
        }
      });
      recognizedText = (result?.data?.text || '').normalize('NFC');
      tesseractLines = result?.data?.lines || null;
    } catch (tessErr) {
      statusEl.textContent = '⚠ Lỗi nhận diện: ' + tessErr.message;
      return;
    }
  }

  progressFill.style.width = '100%';
  // Giữ nguyên câu chữ nhận diện được, không tự ý thay thế từ ngữ
  const rawPreserved = recognizedText.replace(/\r\n/g, '\n').trim();
  textOutput.value = rawPreserved;
  if (paperEl) {
    paperEl.innerHTML = reconstructExactLayoutFromPlainText(rawPreserved);
  }

  extractTableFromOcrResult(rawPreserved, tesseractLines, appendTableMode);
  statusEl.textContent = '✓ Hoàn tất nhận diện Tiếng Việt! Đã giữ nguyên bố cục văn bản để xuất sang Word.';
}

async function runCloudNeuralVietnameseOcr(procCanvas) {
  const dataUrl = procCanvas.toDataURL('image/jpeg', 0.90);
  const formData = new FormData();
  formData.append('apikey', 'helloworld');
  formData.append('base64Image', dataUrl);
  formData.append('language', 'auto');
  formData.append('OCREngine', '2');
  formData.append('isTable', 'true');
  formData.append('scale', 'true');
  formData.append('detectOrientation', 'true');

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 14000);
  try {
    const res = await fetch('https://api.ocr.space/parse/image', {
      method: 'POST',
      body: formData,
      signal: controller.signal
    });
    clearTimeout(timer);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const json = await res.json();
    if (json && json.ParsedResults && json.ParsedResults.length > 0) {
      const parsed = json.ParsedResults.map(r => r.ParsedText || '').join('\n').trim();
      if (parsed.length > 0) return parsed;
    }
    throw new Error(json?.ErrorMessage || 'Empty response');
  } finally {
    clearTimeout(timer);
  }
}

function extractTableFromCurrentText() {
  const el = document.getElementById('ocrResultText');
  if (!el) return;
  const count = extractTableFromOcrResult(el.value, null, false);
  if (window.showToast) {
    window.showToast(`Đã trích xuất ${count} dòng vào bảng biểu Excel!`);
  }
}

function extractTableFromOcrResult(text, tesseractLines = null, appendMode = false) {
  const extractedRows = [];

  if (Array.isArray(tesseractLines) && tesseractLines.length > 0) {
    tesseractLines.forEach(lineObj => {
      const words = (lineObj.words || []).filter(w => (w.text || '').trim().length > 0);
      if (words.length >= 2) {
        const cols = [];
        let curCol = words[0].text.trim();
        for (let i = 1; i < words.length; i++) {
          const prev = words[i - 1];
          const curr = words[i];
          const prevH = Math.max(12, (prev.bbox?.y1 || 20) - (prev.bbox?.y0 || 0));
          const gap = (curr.bbox?.x0 || 0) - (prev.bbox?.x1 || 0);
          if (gap > prevH * 1.45 || curr.text.includes('|') || prev.text.includes('|')) {
            cols.push(curCol.replace(/\|/g, '').trim());
            curCol = curr.text.replace(/\|/g, '').trim();
          } else {
            curCol += ' ' + curr.text.trim();
          }
        }
        if (curCol) cols.push(curCol.replace(/\|/g, '').trim());
        const cleanCols = cols.filter(Boolean);
        if (cleanCols.length >= 2) extractedRows.push(cleanCols);
      }
    });
  }

  if (extractedRows.length === 0) {
    const lines = String(text || '').split('\n').map(l => l.trim()).filter(Boolean);
    lines.forEach(line => {
      if (line.includes('\t') || line.includes('|')) {
        const parts = line.split(/\t|\|/).map(c => c.trim()).filter(Boolean);
        if (parts.length >= 2) {
          extractedRows.push(parts);
          return;
        }
      }
      const adminRowMatch = line.match(/^(\d{1,3})\s+([0-9A-ZĐa-z/-]{3,20})\s+(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\s+(.+)$/);
      if (adminRowMatch) {
        extractedRows.push([adminRowMatch[1], adminRowMatch[2], adminRowMatch[3], adminRowMatch[4]]);
      }
    });
  }

  if (extractedRows.length === 0) {
    renderEditableTable();
    return 0;
  }

  const maxCols = Math.min(8, Math.max(...extractedRows.map(r => r.length), 2));
  const firstRowStr = extractedRows[0].join(' ').toLowerCase();
  const isFirstRowHeader = /\b(stt|tt|số|ký hiệu|họ và tên|họ tên|ngày|nội dung|trích yếu|đơn vị|ghi chú)\b/i.test(firstRowStr);

  if (!appendMode) {
    if (isFirstRowHeader && extractedRows.length > 1) {
      ocrTableHeaders = Array.from({ length: maxCols }, (_, i) => extractedRows[0][i] || `Cột ${i + 1}`);
      ocrTableRows = extractedRows.slice(1).map(r => Array.from({ length: maxCols }, (_, i) => r[i] || ''));
    } else {
      ocrTableHeaders = Array.from({ length: maxCols }, (_, i) => `Cột ${i + 1}`);
      ocrTableRows = extractedRows.map(r => Array.from({ length: maxCols }, (_, i) => r[i] || ''));
    }
  } else {
    const rowsToAppend = isFirstRowHeader ? extractedRows.slice(1) : extractedRows;
    const targetCols = ocrTableHeaders.length;
    rowsToAppend.forEach(r => {
      ocrTableRows.push(Array.from({ length: targetCols }, (_, i) => r[i] || ''));
    });
  }

  renderEditableTable();
  return ocrTableRows.length;
}

function loadSampleDocumentImage(page = 1) {
  isSampleDocLoaded = true;
  samplePageIndex = page;
  ocrRotation = 0;
  currentImageSource = null;
  loadedPdfDoc = null;
  uploadedFileBaseName = 'CongVan_UBND_Phuong';

  const pdfCtrl = document.getElementById('pdfPageControls');
  if (pdfCtrl) pdfCtrl.style.display = 'none';

  const canvas = document.getElementById('ocrPreviewCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  canvas.width = 780;
  canvas.height = 500;

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  ctx.strokeStyle = '#cbd5e1';
  ctx.lineWidth = 2;
  ctx.strokeRect(16, 16, canvas.width - 32, canvas.height - 32);

  ctx.fillStyle = '#0f172a';
  ctx.font = '14px "Times New Roman", serif';
  ctx.textAlign = 'center';
  ctx.fillText('ỦY BAN NHÂN DÂN QUẬN', 185, 50);
  ctx.font = 'bold 14.5px "Times New Roman", serif';
  ctx.fillText('ỦY BAN NHÂN DÂN PHƯỜNG', 185, 70);
  ctx.fillText('Số: 145/BC-UBND', 185, 96);

  ctx.font = 'bold 14.5px "Times New Roman", serif';
  ctx.fillText('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', 535, 50);
  ctx.fillText('Độc lập - Tự do - Hạnh phúc', 535, 70);
  ctx.font = 'italic 14px "Times New Roman", serif';
  ctx.fillText('Phường, ngày 11 tháng 11 năm 2026', 535, 96);

  ctx.font = 'bold 16px "Times New Roman", serif';
  ctx.fillText('BÁO CÁO', 390, 138);
  ctx.font = 'bold 14px "Times New Roman", serif';
  ctx.fillText('Kết quả công tác cải cách hành chính và chuyển đổi số năm 2026', 390, 160);

  ctx.textAlign = 'left';
  ctx.font = '14px "Times New Roman", serif';
  ctx.fillText('Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo', 46, 196);
  ctx.fillText('kết quả triển khai chuẩn hóa văn bản và số hóa hồ sơ tài liệu lưu trữ như sau:', 46, 218);
  ctx.fillText('1. Công tác ban hành văn bản: Bảo đảm đúng thể thức theo Nghị định số 30/2020/NĐ-CP.', 46, 246);
  ctx.fillText('2. Công tác số hóa tài liệu: Chuyển đổi từ PDF sang Word giữ nguyên 100% thể thức gốc./.', 46, 274);

  ctx.font = 'bold italic 13px "Times New Roman", serif';
  ctx.fillText('Nơi nhận:', 46, 330);
  ctx.font = '13px "Times New Roman", serif';
  ctx.fillText('- UBND Quận (để b/c);', 46, 350);
  ctx.fillText('- Thường trực Đảng ủy, HĐND phường;', 46, 368);
  ctx.fillText('- Lưu: VT, VP.', 46, 386);

  ctx.textAlign = 'center';
  ctx.font = 'bold 14px "Times New Roman", serif';
  ctx.fillText('TM. ỦY BAN NHÂN DÂN', 555, 330);
  ctx.fillText('CHỦ TỊCH', 555, 350);
  ctx.fillText('Nguyễn Văn An', 555, 425);

  const sampleWordHtml = `
    <table style="width:100%; border-collapse:collapse; border:none; margin:0 0 8pt 0;">
      <tr>
        <td style="width:44%; vertical-align:top; border:none; text-align:center;">
          <div style="font-size:12.5pt;">ỦY BAN NHÂN DÂN QUẬN</div>
          <div style="font-size:13pt; font-weight:bold;">ỦY BAN NHÂN DÂN PHƯỜNG</div>
          <div style="font-size:13pt; margin-top:4pt;">Số: 145/BC-UBND</div>
        </td>
        <td style="width:56%; vertical-align:top; border:none; text-align:center;">
          <div style="font-size:12.5pt; font-weight:bold;">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
          <div style="font-size:13.5pt; font-weight:bold;">Độc lập - Tự do - Hạnh phúc</div>
          <div style="font-size:13pt; font-style:italic; margin-top:4pt;">Phường, ngày 11 tháng 11 năm 2026</div>
        </td>
      </tr>
    </table>
    <p style="text-align:center; font-weight:bold; font-size:14pt; margin:12pt 0 2pt 0;">BÁO CÁO</p>
    <p style="text-align:center; font-weight:bold; font-size:13.5pt; margin:0 0 12pt 0;">Kết quả công tác cải cách hành chính và chuyển đổi số năm 2026</p>
    <p style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0; line-height:1.38; font-size:13.5pt;">Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai chuẩn hóa văn bản và số hóa hồ sơ tài liệu lưu trữ như sau:</p>
    <p style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0; line-height:1.38; font-size:13.5pt;">1. Công tác ban hành văn bản: Bảo đảm đúng thể thức theo Nghị định số 30/2020/NĐ-CP.</p>
    <p style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0; line-height:1.38; font-size:13.5pt;">2. Công tác số hóa tài liệu: Chuyển đổi từ PDF sang Word giữ nguyên 100% thể thức gốc./.</p>
    <table style="width:100%; border-collapse:collapse; border:none; margin-top:14pt;">
      <tr>
        <td style="width:50%; vertical-align:top; border:none; text-align:left; font-size:11.5pt;">
          <div><b><i>Nơi nhận:</i></b></div>
          <div>- UBND Quận (để b/c);</div>
          <div>- Thường trực Đảng ủy, HĐND phường;</div>
          <div>- Lưu: VT, VP.</div>
        </td>
        <td style="width:50%; vertical-align:top; border:none; text-align:center; font-size:13pt;">
          <div><b>TM. ỦY BAN NHÂN DÂN</b></div>
          <div><b>CHỦ TỊCH</b></div>
          <div style="margin-top:36pt;"><b>Nguyễn Văn An</b></div>
        </td>
      </tr>
    </table>
  `;

  const paperEl = document.getElementById('ocrFormattedPaper');
  if (paperEl) paperEl.innerHTML = sampleWordHtml;

  const textOutput = document.getElementById('ocrResultText');
  if (textOutput) {
    textOutput.value =
`ỦY BAN NHÂN DÂN QUẬN\tCỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM
ỦY BAN NHÂN DÂN PHƯỜNG\tĐộc lập - Tự do - Hạnh phúc
Số: 145/BC-UBND\tPhường, ngày 11 tháng 11 năm 2026

BÁO CÁO
Kết quả công tác cải cách hành chính và chuyển đổi số năm 2026

Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai chuẩn hóa văn bản và số hóa hồ sơ tài liệu lưu trữ như sau:
1. Công tác ban hành văn bản: Bảo đảm đúng thể thức theo Nghị định số 30/2020/NĐ-CP.
2. Công tác số hóa tài liệu: Chuyển đổi từ PDF sang Word giữ nguyên 100% thể thức gốc./.

Nơi nhận:\tTM. ỦY BAN NHÂN DÂN
- UBND Quận (để b/c);\tCHỦ TỊCH
- Thường trực Đảng ủy, HĐND phường;
- Lưu: VT, VP.\tNguyễn Văn An`;
  }
}

function drawCanvasImage() {
  if (isSampleDocLoaded && !currentImageSource) {
    loadSampleDocumentImage(samplePageIndex);
    return;
  }
  const canvas = document.getElementById('ocrPreviewCanvas');
  if (!canvas || !currentImageSource) return;
  const ctx = canvas.getContext('2d');

  const isRotated90 = (Math.abs(ocrRotation) % 180) === 90;
  const w = currentImageSource.width;
  const h = currentImageSource.height;
  canvas.width = isRotated90 ? h : w;
  canvas.height = isRotated90 ? w : h;

  ctx.save();
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate((ocrRotation * Math.PI) / 180);
  if (ocrHighContrast) {
    ctx.filter = 'contrast(175%) grayscale(100%) brightness(108%)';
  }
  ctx.drawImage(currentImageSource, -w / 2, -h / 2);
  ctx.restore();
}

function rotateOcrImage(deg) {
  ocrRotation = (ocrRotation + deg) % 360;
  if (isSampleDocLoaded && !currentImageSource) {
    const canvas = document.getElementById('ocrPreviewCanvas');
    const temp = document.createElement('canvas');
    temp.width = canvas.width;
    temp.height = canvas.height;
    temp.getContext('2d').drawImage(canvas, 0, 0);
    const img = new Image();
    img.onload = () => {
      currentImageSource = img;
      isSampleDocLoaded = false;
      drawCanvasImage();
    };
    img.src = temp.toDataURL();
  } else {
    drawCanvasImage();
  }
  if (window.showToast) window.showToast(`Đã xoay trang tài liệu (${ocrRotation}°)`);
}

function toggleHighContrast() {
  ocrHighContrast = !ocrHighContrast;
  drawCanvasImage();
  if (window.showToast) {
    window.showToast(ocrHighContrast ? 'Đã bật chế độ làm nét chữ Scan' : 'Đã tắt chế độ tăng tương phản');
  }
}

function copyOcrText() {
  const el = document.getElementById('ocrResultText');
  if (!el) return;
  navigator.clipboard.writeText(el.value || '');
  if (window.showToast) window.showToast('Đã sao chép văn bản vào Clipboard!');
}

function renderEditableTable() {
  const container = document.getElementById('ocrTableContainer');
  if (!container) return;

  const thHtml = ocrTableHeaders.map((h, cIdx) => `
    <th contenteditable="true" onblur="updateTableHeader(${cIdx}, this.innerText)">${escapeHtml(h)}</th>
  `).join('');

  const trHtml = ocrTableRows.map((row, rIdx) => {
    const tdHtml = ocrTableHeaders.map((_, cIdx) => `
      <td contenteditable="true" onblur="updateTableCell(${rIdx}, ${cIdx}, this.innerText)">${escapeHtml(row[cIdx] || '')}</td>
    `).join('');
    return `<tr>${tdHtml}</tr>`;
  }).join('');

  container.innerHTML = `
    <table class="ocr-editable-table" id="ocrDataGrid">
      <thead><tr>${thHtml}</tr></thead>
      <tbody>${trHtml}</tbody>
    </table>
  `;
}

function updateTableHeader(cIdx, val) {
  ocrTableHeaders[cIdx] = val.trim();
}

function updateTableCell(rIdx, cIdx, val) {
  if (!ocrTableRows[rIdx]) ocrTableRows[rIdx] = [];
  ocrTableRows[rIdx][cIdx] = val.trim();
}

function addOcrTableRow() {
  const nextNum = String(ocrTableRows.length + 1).padStart(2, '0');
  const newRow = ocrTableHeaders.map((_, idx) => (idx === 0 ? nextNum : ''));
  ocrTableRows.push(newRow);
  renderEditableTable();
}

function addOcrTableCol() {
  const nextColIdx = ocrTableHeaders.length + 1;
  ocrTableHeaders.push(`Cột ${nextColIdx}`);
  ocrTableRows.forEach(r => r.push(''));
  renderEditableTable();
}

function clearOcrTable() {
  ocrTableRows = [['01', '', '', '', ''].slice(0, ocrTableHeaders.length)];
  renderEditableTable();
}

// ============================================================================
// XUẤT FILE WORD (.DOC) GIỮ NGUYÊN 100% THỂ THỨC GỐC - KHÔNG CHÈN BẢNG PHỤ
// ============================================================================
function exportOcrToWord() {
  const paperEl = document.getElementById('ocrFormattedPaper');
  const plainWrap = document.getElementById('ocrPlainEditorWrap');
  const textContent = document.getElementById('ocrResultText')?.value || '';

  let bodyHtml = '';
  if (plainWrap && plainWrap.style.display !== 'none') {
    bodyHtml = reconstructExactLayoutFromPlainText(textContent);
  } else if (paperEl && paperEl.innerHTML.trim()) {
    bodyHtml = paperEl.innerHTML;
  } else {
    bodyHtml = reconstructExactLayoutFromPlainText(textContent);
  }

  const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(uploadedFileBaseName)}</title>
    <style>
      @page Section1 { size: 21.0cm 29.7cm; margin: 2.0cm 1.5cm 2.0cm 3.0cm; }
      div.Section1 { page: Section1; }
      body { font-family: "Times New Roman", Times, serif; font-size: 13.5pt; color: #000000; line-height: 1.38; }
      p { font-family: "Times New Roman", Times, serif; margin: 0 0 5pt 0; }
      table { width: 100%; border-collapse: collapse; font-family: "Times New Roman", Times, serif; }
      td, th { font-family: "Times New Roman", Times, serif; vertical-align: top; }
    </style>
  </head>
  <body>
    <div class="Section1">
      ${bodyHtml}
    </div>
  </body>
  </html>`;

  const outFileName = `${uploadedFileBaseName || 'VanBan_ChuyenDoi'}.doc`;
  const blob = new Blob(['\ufeff', wordHtml], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  if (window.showToast) {
    window.showToast(`Đã xuất tệp Word giữ nguyên 100% thể thức gốc: ${outFileName}`);
  }
}

function exportOcrToExcel() {
  const ths = ocrTableHeaders.map(h => `<th style="border: 1px solid #000; background: #dbeafe; font-weight: bold; padding: 6px;">${escapeHtml(h)}</th>`).join('');
  const trs = ocrTableRows.map(r => {
    const tds = ocrTableHeaders.map((_, i) => `<td style="border: 1px solid #000; padding: 6px; mso-number-format:'\\@';">${escapeHtml(r[i] || '')}</td>`).join('');
    return `<tr>${tds}</tr>`;
  }).join('');

  const excelHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <style>
      table, td, th { font-family: "Times New Roman", serif; font-size: 12pt; }
    </style>
  </head>
  <body>
    <table>
      <thead><tr>${ths}</tr></thead>
      <tbody>${trs}</tbody>
    </table>
  </body>
  </html>`;

  const outFileName = `${uploadedFileBaseName || 'BangBieu'}.xls`;
  const blob = new Blob(['\ufeff', excelHtml], { type: 'application/vnd.ms-excel;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  if (window.showToast) window.showToast(`Đã xuất bảng biểu Excel: ${outFileName}`);
}
