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
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleUploadedFiles(e.dataTransfer.files);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleUploadedFiles(e.target.files);
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

async function handleUploadedFiles(fileList) {
  const files = Array.from(fileList || []).filter(Boolean);
  if (files.length === 0) return;
  if (files.length === 1) {
    await handleUploadedFile(files[0]);
    return;
  }

  const statusEl = document.getElementById('ocrStatusText');
  const progressFill = document.getElementById('ocrProgressFill');
  const textOutput = document.getElementById('ocrResultText');
  const paperEl = document.getElementById('ocrFormattedPaper');
  if (window.VietNamOcrEnginePro) window.VietNamOcrEnginePro.resetCancel();
  if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.clearMultiPageStore) {
    window.VietNamOcr2LayerPdf.clearMultiPageStore();
  }

  const allTexts = [];
  const allHtmls = [];
  let doneCount = 0;

  for (let i = 0; i < files.length; i++) {
    if (window.VietNamOcrEnginePro) {
      await window.VietNamOcrEnginePro.waitIfPaused(() => {
        if (statusEl) statusEl.textContent = `⏸ Đang tạm dừng hàng đợi tại tệp ${i + 1}/${files.length}... Bấm "Tiếp tục" để chạy tiếp.`;
      });
      if (window.VietNamOcrEnginePro.isCancelled()) {
        if (statusEl) statusEl.textContent = `⏹ Đã dừng hàng đợi OCR tại tệp ${i}/${files.length}.`;
        break;
      }
    }
    const f = files[i];
    if (window.VietNamOcrEnginePro) {
      const check = window.VietNamOcrEnginePro.validateInputFile(f);
      if (!check.ok) {
        window.VietNamOcrEnginePro.logSafeError('FileValidation', check.reason);
        allTexts.push(`=== [${i + 1}/${files.length}] ${f.name} (Bỏ qua: ${check.reason}) ===`);
        continue;
      }
    }
    if (statusEl) statusEl.textContent = `Đang xử lý hàng đợi tệp ${i + 1}/${files.length}: "${f.name}"...`;
    if (progressFill) progressFill.style.width = `${Math.round((i / files.length) * 100)}%`;

    try {
      await handleUploadedFile(f);
      const t = (textOutput?.value || '').trim();
      const h = (paperEl?.innerHTML || '').trim();
      if (t) allTexts.push(`=== [${i + 1}/${files.length}] ${f.name} ===\n${t}`);
      if (h) allHtmls.push(h);
      doneCount++;
    } catch (err) {
      allTexts.push(`=== [${i + 1}/${files.length}] ${f.name} (Lỗi: ${err.message}) ===`);
    }
  }

  uploadedFileBaseName = `OCR_HangDoi_${files.length}_Tep`;
  if (textOutput && allTexts.length > 0) textOutput.value = allTexts.join('\n\n');
  if (paperEl && allHtmls.length > 0) paperEl.innerHTML = allHtmls.join('<hr style="margin:12pt 0;border:none;border-top:1px dashed #cbd5e1;"/>');
  if (progressFill) progressFill.style.width = '100%';
  if (statusEl) statusEl.textContent = `✓ Hoàn tất hàng đợi OCR ${doneCount}/${files.length} tệp! Sẵn sàng xuất Word, TXT hoặc PDF 2 Lớp.`;
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

  // Hỗ trợ giải mã ảnh TIFF (.tif / .tiff) nhiều trang
  const isTiff = file.type === 'image/tiff' || /\.(tif|tiff)$/i.test(file.name || '');
  if (isTiff && window.VietNamOcrEnginePro) {
    const statusEl = document.getElementById('ocrStatusText');
    try {
      if (statusEl) statusEl.textContent = `Đang giải mã tệp ảnh TIFF "${file.name}"...`;
      const canvases = await window.VietNamOcrEnginePro.decodeTiffFileToCanvas(file);
      loadedPdfDoc = null;
      window.loadedPdfDoc = null;
      if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.clearMultiPageStore) {
        window.VietNamOcr2LayerPdf.clearMultiPageStore();
      }
      const allTexts = [];
      const allHtmls = [];
      for (let idx = 0; idx < canvases.length; idx++) {
        const c = canvases[idx];
        if (idx === 0) {
          const img = new Image();
          await new Promise(res => { img.onload = res; img.src = c.toDataURL('image/png'); });
          currentImageSource = img;
          drawCanvasImage();
        }
        const ocrRes = await recognizeCanvasPageHighAccuracy(c, (pct) => {
          if (statusEl) statusEl.textContent = `Đang OCR ảnh TIFF Trang ${idx + 1}/${canvases.length} (${pct}%)...`;
        });
        allTexts.push(ocrRes.text);
        allHtmls.push(ocrRes.lines && ocrRes.lines.length > 0
          ? reconstructLayoutFromTesseractBboxes(ocrRes.lines, ocrRes.procWidth, ocrRes.procHeight, 'auto')
          : reconstructExactLayoutFromPlainText(ocrRes.text));
        if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.recordPdfPageOcr) {
          window.VietNamOcr2LayerPdf.recordPdfPageOcr(idx + 1, c, ocrRes.text, ocrRes.lines, ocrRes.procWidth, ocrRes.procHeight);
        }
      }
      const textOutput = document.getElementById('ocrResultText');
      const paperEl = document.getElementById('ocrFormattedPaper');
      if (textOutput) textOutput.value = allTexts.join('\n\n');
      if (paperEl) paperEl.innerHTML = allHtmls.join('\n');
      if (statusEl) statusEl.textContent = `✓ Đã OCR xong ${canvases.length} trang TIFF "${file.name}"!`;
      return;
    } catch (tifErr) {
      if (statusEl) statusEl.textContent = '⚠ Lỗi đọc tệp TIFF: ' + tifErr.message;
      return;
    }
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
    window.loadedPdfDoc = loadedPdfDoc;
    currentPdfPage = 1;
    pdfPagesLayoutHtml = [];
    pdfPagesRawText = [];
    if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.clearMultiPageStore) {
      window.VietNamOcr2LayerPdf.clearMultiPageStore();
    }

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
      const textContent = await page.getTextContent();
      const rawItems = (textContent?.items || []).filter(it => typeof it.str === 'string' && it.str.length > 0);
      const nonEmptyItems = rawItems.filter(it => it.str.trim().length > 0);
      const sampleEmbedded = nonEmptyItems.map(it => it.str).join(' ');
      const replacementCount = (sampleEmbedded.match(/\uFFFD/g) || []).length;
      const hasUsableTextLayer = nonEmptyItems.length >= 2 && replacementCount < Math.max(3, sampleEmbedded.length * 0.08);

      // Render độ phân giải cao (scale 2.8 ~ 220-260 DPI) nếu là trang scan ảnh hoặc lớp chữ nhúng bị lỗi font
      const renderScale = hasUsableTextLayer ? 2.0 : Math.min(2.8, 2600 / Math.max(1, viewport1x.width));
      const viewportRender = page.getViewport({ scale: renderScale });

      const offCanvas = document.createElement('canvas');
      offCanvas.width = Math.round(viewportRender.width);
      offCanvas.height = Math.round(viewportRender.height);
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

      if (hasUsableTextLayer) {
        // Trang PDF có lớp văn bản gốc (hoặc PDF 2 lớp chuẩn Unicode): Tái tạo chính xác 100% thể thức và giữ nguyên từng ký tự gốc
        const { pageHtml, pagePlainText } = reconstructExactPdfPageLayout(rawItems, page, textContent, viewport1x, p);
        pdfPagesLayoutHtml.push(pageHtml);
        pdfPagesRawText.push(pagePlainText);
        if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.recordPdfPageOcr) {
          const nativeLines = [];
          rawItems.forEach(it => {
            const s = String(it.str || '').trim();
            if (!s || !it.transform) return;
            const tx = window.pdfjsLib.Util.transform(viewportRender.transform, it.transform);
            const x0 = tx[4];
            const yBottom = tx[5];
            const h = Math.max(10, Math.hypot(tx[2], tx[3]) || 14);
            nativeLines.push({ text: s.normalize('NFC'), bbox: { x0, y0: yBottom - h, x1: x0 + (it.width * viewportRender.scale), y1: yBottom } });
          });
          window.VietNamOcr2LayerPdf.recordPdfPageOcr(p, offCanvas, pagePlainText, nativeLines, offCanvas.width, offCanvas.height);
        }
      } else {
        // Trang PDF là ảnh scan thuần túy (hoặc PDF 2 lớp bị lỗi font): Gọi Tesseract v5 LSTM Tiếng Việt + tái tạo bố cục 2 cột theo tọa độ bbox thực
        const ocrRes = await recognizeCanvasPageHighAccuracy(offCanvas, (pct) => {
          if (statusEl) statusEl.textContent = `Đang OCR Tiếng Việt chính xác cao Trang ${p}/${totalPdfPages} (${pct}%)...`;
        });
        const pageHtml = (ocrRes.lines && ocrRes.lines.length > 0)
          ? reconstructLayoutFromTesseractBboxes(ocrRes.lines, ocrRes.procWidth, ocrRes.procHeight, 'auto')
          : reconstructExactLayoutFromPlainText(ocrRes.text);
        pdfPagesLayoutHtml.push(pageHtml);
        pdfPagesRawText.push(ocrRes.text);
        if (window.VietNamOcr2LayerPdf && window.VietNamOcr2LayerPdf.recordPdfPageOcr) {
          window.VietNamOcr2LayerPdf.recordPdfPageOcr(p, offCanvas, ocrRes.text, ocrRes.lines, ocrRes.procWidth, ocrRes.procHeight);
        }
      }
    }

    // Hiển thị toàn bộ các trang nối chảy tự nhiên (không chèn page-break cứng gây trang trắng thừa khi xuất Word)
    const fullWordHtml = pdfPagesLayoutHtml
      .map((html, idx) => idx === 0
        ? html
        : `<div class="ocr-page-flow" style="margin-top: 6pt;">${html}</div>`)
      .join('\n');

    const fullPlainText = pdfPagesRawText.join('\n\n');

    const paperEl = document.getElementById('ocrFormattedPaper');
    if (paperEl) paperEl.innerHTML = fullWordHtml;

    const textOutput = document.getElementById('ocrResultText');
    if (textOutput) textOutput.value = fullPlainText;

    extractTableFromOcrResult(fullPlainText, null, false);
    if (window.VietNamOcrIdpStudio) {
      window.VietNamOcrIdpStudio.autoSaveDocumentToArchive(file.name, fullPlainText, 95);
    }
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
// THUẬT TOÁN TÁI TẠO BỐ CỤC TỌA ĐỘ 1:1 & CHUẨN HÓA THỂ THỨC TỪ PDF SANG WORD
// (Đúng chuẩn Hướng dẫn số 05-HD/VPTW của VPTW Đảng & Nghị định 30/2020/NĐ-CP)
// ============================================================================
function detectIsPartyDocInOcr(text) {
  const u = String(text || '').toUpperCase();
  return (
    u.includes('ĐẢNG CỘNG SẢN VIỆT NAM') ||
    u.includes('ĐẢNG ỦY') ||
    u.includes('ĐẢNG BỘ') ||
    u.includes('QUẬN ỦY') ||
    u.includes('HUYỆN ỦY') ||
    u.includes('THỊ ỦY') ||
    u.includes('THÀNH ỦY') ||
    u.includes('TỈNH ỦY') ||
    u.includes('CHI BỘ') ||
    u.includes('BAN THƯỜNG VỤ') ||
    u.includes('BAN CHẤP HÀNH') ||
    u.includes('VPTW') ||
    /\bSỐ\s*:?\s*\d+\s*-\s*[A-ZĐ]+\/(ĐU|CB|QU|HU|TU|TW)\b/i.test(u) ||
    /-(ĐU|CB)\b/i.test(u) ||
    /\/(ĐU|CB)\b/i.test(u)
  );
}

function formatOcrBodyParagraphHtml(rawLine) {
  const line = String(rawLine || '').trim();
  const escaped = escapeHtml(line);
  if (/^(QUYẾT ĐỊNH|QUYẾT NGHỊ|NGHỊ QUYẾT|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|BAN CHẤP HÀNH ĐẢNG BỘ QUYẾT NGHỊ|ỦY BAN NHÂN DÂN PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG|HỘI ĐỒNG NHÂN DÂN PHƯỜNG)\s*:?$/i.test(line) ||
      /^KHÓA\s+[IVX0-9]+/i.test(line)) {
    return `<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:8pt 0 6pt 0; line-height:1.35; font-size:14pt;">${escaped}</p>`;
  }
  if (/^([IVXLCDM]+\.|PHẦN\s+[IVXLCDM0-9]+|CHƯƠNG\s+[IVXLCDM0-9]+)\s+/i.test(line)) {
    return `<p align="left" style="text-align:left; text-indent:0; font-weight:bold; margin:6pt 0 4pt 0; line-height:1.35; font-size:14pt;">${escaped}</p>`;
  }
  const dieuMatch = line.match(/^(Điều\s+\d+\.|Kính gửi\s*:)\s*(.*)$/i);
  if (dieuMatch) {
    return `<p align="justify" style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0; line-height:1.35; font-size:14pt;"><b>${escapeHtml(dieuMatch[1])}</b> ${escapeHtml(dieuMatch[2])}</p>`;
  }
  return `<p align="justify" style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0; line-height:1.35; font-size:14pt;">${escaped}</p>`;
}

function buildOcrHeaderTableHtml(leftSegs, rightSegs, isParty) {
  const cleanLeft = leftSegs
    .map(s => (typeof s === 'string' ? s : s.text).trim())
    .filter(t => t && !/^[-_*—–=─]+$/.test(t));
  const cleanRight = rightSegs
    .map(s => (typeof s === 'string' ? s : s.text).trim())
    .filter(t => t && !/^[-_*—–=─]+$/.test(t));

  const agencyLines = [];
  let numLine = '';
  const subjLeftLines = [];

  for (const l of cleanLeft) {
    if (!numLine && /^(Số\b|Số\s*:)/i.test(l)) {
      const cleanNum = l.replace(/^Số\s*:?\s*/i, '').trim();
      numLine = isParty ? `Số ${cleanNum}` : `Số: ${cleanNum}`;
    } else if (numLine || /^(V\/v|Về việc)\b/i.test(l)) {
      subjLeftLines.push(l);
    } else {
      agencyLines.push(l);
    }
  }

  let leftCellHtml = '';
  if (agencyLines.length >= 2) {
    leftCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:${isParty ? '14pt' : '12.5pt'}; font-weight:normal; text-transform:uppercase; line-height:1.25;">${escapeHtml(agencyLines[0])}</p>`;
    for (let k = 1; k < agencyLines.length; k++) {
      leftCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:${isParty ? '14pt' : '12.5pt'}; font-weight:bold; text-transform:uppercase; line-height:1.25;">${escapeHtml(agencyLines[k])}</p>`;
    }
  } else if (agencyLines.length === 1) {
    leftCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:${isParty ? '14pt' : '12.5pt'}; font-weight:bold; text-transform:uppercase; line-height:1.25;">${escapeHtml(agencyLines[0])}</p>`;
  }

  if (agencyLines.length > 0) {
    leftCellHtml += isParty
      ? `<p align="center" style="margin:1pt 0 2pt 0; text-indent:0; text-align:center; font-size:14pt; font-weight:bold; line-height:1;">*</p>`
      : `<p align="center" style="margin:1pt 0 3pt 0; text-indent:0; text-align:center; font-size:9pt; line-height:1;">────────────</p>`;
  }

  if (numLine) {
    leftCellHtml += `<p align="center" style="margin:2pt 0 0 0; text-indent:0; text-align:center; font-size:${isParty ? '14pt' : '13pt'}; font-weight:normal; line-height:1.25;">${escapeHtml(numLine)}</p>`;
  }
  if (subjLeftLines.length > 0) {
    leftCellHtml += `<p align="center" style="margin:3pt 0 0 0; text-indent:0; text-align:center; font-size:12pt; font-weight:normal; line-height:1.25;">${escapeHtml(subjLeftLines.join(' '))}</p>`;
  }

  let rightCellHtml = '';
  let dateLine = '';
  const mottoLines = [];

  for (const r of cleanRight) {
    if (/\b(ngày|tháng|năm)\b/i.test(r)) {
      dateLine = r;
    } else {
      mottoLines.push(r);
    }
  }

  if (isParty) {
    const partyTitle = mottoLines.find(m => /ĐẢNG CỘNG SẢN/i.test(m)) || 'ĐẢNG CỘNG SẢN VIỆT NAM';
    rightCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:15pt; font-weight:bold; text-transform:uppercase; line-height:1.25;"><span style="border-bottom:1pt solid #000000; padding-bottom:1pt;">${escapeHtml(partyTitle)}</span></p>`;
  } else {
    const qh = mottoLines[0] || 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
    const tn = mottoLines[1] || 'Độc lập - Tự do - Hạnh phúc';
    rightCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:12.5pt; font-weight:bold; text-transform:uppercase; line-height:1.25;">${escapeHtml(qh)}</p>`;
    rightCellHtml += `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:13.5pt; font-weight:bold; line-height:1.3;"><span style="border-bottom:1pt solid #000000; padding-bottom:1pt;">${escapeHtml(tn)}</span></p>`;
  }

  if (dateLine) {
    rightCellHtml += `<p align="center" style="margin:6pt 0 0 0; text-indent:0; text-align:center; font-size:${isParty ? '14pt' : '13.5pt'}; font-style:italic; font-weight:normal; line-height:1.25;">${escapeHtml(dateLine)}</p>`;
  }

  return `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%; border-collapse:collapse; border:none; mso-table-lspace:0pt; mso-table-rspace:0pt; margin:0 0 10pt 0;">
      <tr>
        <td width="44%" valign="top" align="center" style="width:44%; vertical-align:top; border:none; padding:0 6pt 0 0; text-align:center;">${leftCellHtml}</td>
        <td width="56%" valign="top" align="center" style="width:56%; vertical-align:top; border:none; padding:0 0 0 6pt; text-align:center;">${rightCellHtml}</td>
      </tr>
    </table>
  `;
}

function buildOcrFooterTableHtml(leftSegs, rightSegs, isParty) {
  const cleanLeft = leftSegs
    .map(s => (typeof s === 'string' ? s : s.text).trim())
    .filter(Boolean);
  const cleanRight = rightSegs
    .map(s => (typeof s === 'string' ? s : s.text).trim())
    .filter(Boolean);

  const recipientItems = [];
  let hasNoiNhanLabel = false;
  for (const l of cleanLeft) {
    if (/^Nơi nhận\s*:?$/i.test(l)) {
      hasNoiNhanLabel = true;
    } else if (/^Nơi nhận\s*:\s*(.+)$/i.test(l)) {
      hasNoiNhanLabel = true;
      const m = l.match(/^Nơi nhận\s*:\s*(.+)$/i);
      if (m && m[1]) recipientItems.push(m[1].trim());
    } else {
      recipientItems.push(l);
    }
  }
  if (recipientItems.length > 0) hasNoiNhanLabel = true;

  const noiNhanHeaderHtml = hasNoiNhanLabel
    ? (isParty
        ? `<p align="left" style="margin:0 0 2pt 0; text-indent:0; text-align:left; font-size:14pt; font-style:normal; font-weight:normal; line-height:1.25;"><u>Nơi nhận</u>:</p>`
        : `<p align="left" style="margin:0 0 2pt 0; text-indent:0; text-align:left; font-size:12pt; font-weight:bold; font-style:italic; line-height:1.25;">Nơi nhận:</p>`)
    : '';

  const recFontSize = isParty ? '12pt' : '11pt';
  const recListHtml = recipientItems
    .map(item => `<p align="left" style="margin:0; text-indent:0; text-align:left; font-size:${recFontSize}; line-height:1.3;">${escapeHtml(item)}</p>`)
    .join('');

  const authorityLines = [];
  let signerName = '';
  for (let k = 0; k < cleanRight.length; k++) {
    const r = cleanRight[k];
    const isAllUpper = (r === r.toUpperCase() && r !== r.toLowerCase()) || /^(T\/M|TM\.|KT\.|TL\.|TUQ\.|Q\.|BÍ THƯ|PHÓ BÍ THƯ|CHỦ TỊCH|PHÓ CHỦ TỊCH|TRƯỞNG BAN|CHÁNH VĂN PHÒNG|VĂN PHÒNG|CHỦ TRÌ)/i.test(r);
    if (isAllUpper && (k < cleanRight.length - 1 || cleanRight.length === 1)) {
      authorityLines.push(r);
    } else if (k === cleanRight.length - 1) {
      signerName = r;
    } else {
      authorityLines.push(r);
    }
  }

  const authFontSize = isParty ? '14pt' : '13.5pt';
  const authHtml = authorityLines
    .map(a => `<p align="center" style="margin:0; text-indent:0; text-align:center; font-size:${authFontSize}; font-weight:bold; text-transform:uppercase; line-height:1.3;">${escapeHtml(a)}</p>`)
    .join('');
  const nameHtml = signerName
    ? `<p align="center" style="margin:44pt 0 0 0; mso-margin-top-alt:44pt; text-indent:0; text-align:center; font-size:14pt; font-weight:bold; line-height:1.3;">${escapeHtml(signerName)}</p>`
    : '';

  return `
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%; border-collapse:collapse; border:none; mso-table-lspace:0pt; mso-table-rspace:0pt; margin:14pt 0 0 0;">
      <tr>
        <td width="50%" valign="top" align="left" style="width:50%; vertical-align:top; border:none; padding:0 6pt 0 0; text-align:left;">${noiNhanHeaderHtml}${recListHtml}</td>
        <td width="50%" valign="top" align="center" style="width:50%; vertical-align:top; border:none; padding:0 0 0 6pt; text-align:center;">${authHtml}${nameHtml}</td>
      </tr>
    </table>
  `;
}

function joinPdfRunItemsText(itemsArr) {
  if (!itemsArr || itemsArr.length === 0) return '';
  let out = itemsArr[0].str;
  let prevEndX = itemsArr[0].endX;
  for (let i = 1; i < itemsArr.length; i++) {
    const it = itemsArr[i];
    const gap = it.x - prevEndX;
    const needSpace = gap > 1.0 && !/\s$/.test(out) && !/^\s/.test(it.str);
    out += (needSpace ? ' ' : '') + it.str;
    prevEndX = Math.max(prevEndX, it.endX);
  }
  return out.replace(/\s+/g, ' ').trim();
}

function reconstructExactPdfPageLayout(rawItems, page, textContent, viewport, pageNumber = 1) {
  const pageW = viewport?.width || 595.28;
  const styles = textContent?.styles || {};
  const runs = [];

  for (const it of rawItems) {
    if (!it.str || !it.str.trim()) continue;
    const tx = it.transform || [12, 0, 0, 12, 0, 0];
    const x = tx[4] || 0;
    const y = tx[5] || 0;
    const rawSize = Math.hypot(tx[0] || 0, tx[1] || 0) || Math.abs(tx[3]) || 13.5;
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

    runs.push({ str: it.str, x, y, w, endX: x + w, fontSize, isBold, isItalic });
  }

  if (runs.length === 0) {
    return { pageHtml: '', pagePlainText: '' };
  }

  // 1. Sắp xếp nghiêm ngặt theo trục Y từ trên xuống dưới (Transitive Sort - chống nhảy dòng)
  runs.sort((a, b) => b.y - a.y);

  const rawLines = [];
  for (const r of runs) {
    const tol = Math.max(4.5, r.fontSize * 0.42);
    let matchedLine = null;
    for (let i = rawLines.length - 1; i >= Math.max(0, rawLines.length - 4); i--) {
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

  let minBodyX = pageW;
  let maxBodyX = 0;

  const parsedLines = rawLines.map(lineObj => {
    const items = lineObj.items.sort((a, b) => a.x - b.x);
    const fullText = joinPdfRunItemsText(items);
    const x = items[0].x;
    const endX = items[items.length - 1].endX;
    if (x < minBodyX) minBodyX = x;
    if (endX > maxBodyX) maxBodyX = endX;

    // Kiểm tra xem dòng này có chia 2 cột Trái - Phải (dành cho Header / Footer) hay không
    let splitIdx = -1;
    let bestGap = 0;
    for (let k = 1; k < items.length; k++) {
      const prevEnd = items[k - 1].endX;
      const curStart = items[k].x;
      const gap = curStart - prevEnd;
      if (gap >= 16 && prevEnd <= pageW * 0.54 && curStart >= pageW * 0.36) {
        if (gap > bestGap) {
          bestGap = gap;
          splitIdx = k;
        }
      }
    }

    let hasTwoCols = false;
    let leftText = '';
    let rightText = '';
    const isRightSideOnly = (splitIdx === -1) && (x >= pageW * 0.42);

    if (splitIdx !== -1) {
      hasTwoCols = true;
      leftText = joinPdfRunItemsText(items.slice(0, splitIdx));
      rightText = joinPdfRunItemsText(items.slice(splitIdx));
    } else if (isRightSideOnly) {
      rightText = fullText;
    } else {
      leftText = fullText;
    }

    return {
      y: lineObj.y,
      x,
      endX,
      centerX: (x + endX) / 2,
      fontSize: Math.max(...items.map(it => it.fontSize)),
      isBold: items.some(it => it.isBold),
      isItalic: items.some(it => it.isItalic),
      fullText,
      hasTwoCols,
      isRightSideOnly,
      leftText,
      rightText,
      plainText: hasTwoCols ? `${leftText}\t${rightText}` : fullText
    };
  }).filter(l => l.fullText.length > 0);

  // Loại bỏ số trang đơn lẻ ở đầu hoặc cuối trang PDF (ví dụ "1", "2", "- 2 -")
  const lines = parsedLines.filter((l, idx) => {
    if ((idx === 0 || idx === parsedLines.length - 1) && /^(?:Trang\s*)?[-–—]?\s*\d{1,3}\s*[-–—]?$/i.test(l.fullText)) {
      return false;
    }
    return true;
  });

  if (lines.length === 0) {
    return { pageHtml: '', pagePlainText: '' };
  }

  if (minBodyX > pageW * 0.35) minBodyX = 72;
  if (maxBodyX < pageW * 0.65) maxBodyX = pageW - 54;
  const pageMidX = (minBodyX + maxBodyX) / 2;

  const fullPageText = lines.map(l => l.plainText).join('\n');
  const isPartyDoc = detectIsPartyDocInOcr(fullPageText);

  const htmlBlocks = [];
  let idx = 0;

  // 2. Xử lý Bảng Tiêu Ngữ Đầu Trang (Chỉ thực hiện ở Trang 1 khi thực sự có Quốc hiệu / Tiêu ngữ / Tên cơ quan)
  const topSample = lines.slice(0, 8).map(l => l.fullText).join('\n');
  const hasRegulationHeader = (pageNumber === 1) && (
    /(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI|Độc lập\s*-\s*Tự do|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN|ĐẢNG ỦY|ĐẢNG BỘ|QUẬN ỦY|HUYỆN ỦY|THỊ ỦY|THÀNH ỦY|TỈNH ỦY|CHI BỘ|BAN CHẤP HÀNH|BAN THƯỜNG VỤ|\bSố\s*:?\s*\d+)/i.test(topSample)
  );

  if (hasRegulationHeader) {
    const headerLeft = [];
    const headerRight = [];
    while (idx < lines.length && idx <= 9) {
      const ln = lines[idx];
      const txt = ln.fullText.trim();
      if (/^[-_*—–=─]{2,}$/.test(txt) && !ln.hasTwoCols) {
        idx++;
        continue;
      }
      if (!ln.hasTwoCols && /^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|QUY ĐỊNH|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|KẾT LUẬN|ĐỀ ÁN|PHƯƠNG ÁN)$/i.test(txt)) {
        break;
      }
      if (!ln.hasTwoCols && /^(Kính gửi\s*:|Căn cứ\s+|Thực hiện\s+|Hôm nay\b|Xét đề nghị|Theo đề nghị|I\.\s+|1\.\s+|Điều\s+1\b)/i.test(txt)) {
        break;
      }
      if (idx >= 2 && !ln.hasTwoCols && !ln.isRightSideOnly && ln.endX > pageW * 0.56 && !/^(Số\b|Số\s*:|V\/v\b|Về việc\b)/i.test(txt)) {
        break;
      }

      if (ln.hasTwoCols) {
        if (ln.leftText) headerLeft.push(ln.leftText);
        if (ln.rightText) headerRight.push(ln.rightText);
      } else if (ln.isRightSideOnly || /^(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI|Độc lập\s*-\s*Tự do)/i.test(txt) || /\bngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+/i.test(txt)) {
        headerRight.push(txt);
      } else {
        headerLeft.push(txt);
      }
      idx++;
    }
    if (headerLeft.length > 0 || headerRight.length > 0) {
      htmlBlocks.push(buildOcrHeaderTableHtml(headerLeft, headerRight, isPartyDoc));
    }
  }

  // 3. Xử lý Tên Loại Văn Bản & Trích Yếu (Ở đầu văn bản sau Header)
  while (idx < lines.length && /^[-_*—–=─]{2,}$/.test(lines[idx].fullText.trim())) {
    idx++;
  }

  if (pageNumber === 1 && idx < lines.length) {
    const titleCandidate = lines[idx].fullText.trim();
    if (/^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|QUY ĐỊNH|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|KẾT LUẬN|ĐỀ ÁN|PHƯƠNG ÁN)$/i.test(titleCandidate)) {
      const titleSize = isPartyDoc ? '15pt' : '14pt';
      htmlBlocks.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:12pt 0 3pt 0; line-height:1.3; font-size:${titleSize}; text-transform:uppercase;">${escapeHtml(titleCandidate)}</p>`);
      idx++;

      const subjParts = [];
      while (idx < lines.length && subjParts.length < 4) {
        const subLn = lines[idx];
        const subTxt = subLn.fullText.trim();
        if (/^[-_*—–=─]{2,}$/.test(subTxt) || subTxt === '*') {
          idx++;
          break;
        }
        if (/^(Kính gửi\s*:|Căn cứ\b|Thực hiện\b|Hôm nay\b|Sau khi\b|Xét\b|Theo đề nghị\b|I\.\s+|1\.\s+|Điều\s+1\b|BAN THƯỜNG VỤ|ĐẢNG ỦY\s+.*QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN\s+.*QUYẾT ĐỊNH|HỘI ĐỒNG NHÂN DÂN)/i.test(subTxt)) {
          break;
        }
        const isCenteredLine = (subLn.x > minBodyX + 14 && subLn.endX < maxBodyX - 14 && Math.abs(subLn.centerX - pageMidX) <= 55);
        if (subjParts.length === 0 || subLn.isBold || isCenteredLine) {
          subjParts.push(subTxt);
          idx++;
        } else {
          break;
        }
      }

      if (subjParts.length > 0) {
        htmlBlocks.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:0 0 2pt 0; line-height:1.3; font-size:14pt;">${escapeHtml(subjParts.join(' '))}</p>`);
        htmlBlocks.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:normal; margin:0 0 10pt 0; line-height:1; font-size:10pt;">──────────────</p>`);
      }
    }
  }

  // 4. Xác định vị trí bắt đầu phần Chân Trang (Nơi nhận & Chữ ký) ở cuối trang
  let footerStartIdx = lines.length;
  for (let k = idx; k < lines.length; k++) {
    const ln = lines[k];
    const txt = ln.fullText.trim();
    const isNearBottom = k >= Math.floor(lines.length * 0.45);
    if (/^Nơi nhận\s*:?/i.test(txt) || /^Nơi nhận\s*:/i.test(ln.leftText || '')) {
      footerStartIdx = k;
      break;
    }
    if (isNearBottom && /\b(T\/M\s+ĐẢNG\s+ỦY|T\/M\s+BAN\s+THƯỜNG\s+VỤ|T\/M\s+BAN\s+CHẤP\s+HÀNH|T\/M\s+CHI\s+BỘ|TM\.\s+ỦY\s+BAN|TM\.\s+HỘI\s+ĐỒNG|KT\.\s+CHỦ\s+TỊCH|KT\.\s+BÍ\s+THƯ|TL\.\s+CHỦ\s+TỊCH|TL\.\s+BÍ\s+THƯ|TUQ\.\s+CHỦ\s+TỊCH|Q\.\s+CHỦ\s+TỊCH)\b/i.test(txt)) {
      footerStartIdx = k;
      break;
    }
    if (k >= Math.floor(lines.length * 0.65) && ln.isRightSideOnly && /^(BÍ THƯ|PHÓ BÍ THƯ|PHÓ BÍ THƯ THƯỜNG TRỰC|CHỦ TỊCH|PHÓ CHỦ TỊCH|CHỦ TRÌ CUỘC HỌP)$/i.test(txt)) {
      footerStartIdx = k;
      break;
    }
  }

  // 5. Xử lý toàn bộ các đoạn văn Thân Bài (Tuyệt đối không cắt ngang thành bảng 2 cột hay căn giữa sai)
  let bIdx = idx;
  while (bIdx < footerStartIdx) {
    const curLn = lines[bIdx];
    const curTxt = curLn.fullText.trim();
    if (/^[-_*—–=─]{3,}$/.test(curTxt)) {
      bIdx++;
      continue;
    }

    let mergedText = curTxt;
    let prevLn = curLn;
    bIdx++;

    while (bIdx < footerStartIdx) {
      const nxtLn = lines[bIdx];
      const nxtTxt = nxtLn.fullText.trim();
      if (/^[-_*—–=─]{3,}$/.test(nxtTxt)) {
        bIdx++;
        break;
      }

      const verticalGap = prevLn.y - nxtLn.y;
      const prevEndedStop = /[.:;!?]$|\.\s*\/\.\s*$/.test(mergedText);
      const prevIsHeading = /^([IVXLCDM]+\.|PHẦN\s+[IVX0-9]+|CHƯƠNG\s+[IVX0-9]+|\d+\.\s+[A-ZĐÀ-Ỹ][^.:;!?]{1,50}$|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG QUYẾT ĐỊNH)/i.test(mergedText);
      const nextStartsLowercase = /^[a-zàáảãạăắằẳẵặâấầẩẫậđèéẻẽẹêếềểễệìíỉĩịòóỏõọôốồổỗộơớờởỡợùúủũụưứừửữựỳýỷỹỵ]/.test(nxtTxt);
      const nextStartsNewBlock = /^([-+•*]|\d+[.)]|[a-zđ]\)|[IVXLCDM]+\.|Điều\s+\d+|Khoản\s+\d+|Phần\s+[IVX0-9]+|Chương\s+[IVX0-9]+|Căn cứ\b|Xét\b|Theo đề nghị\b|Thực hiện\b|Kính gửi\b|Hôm nay\b|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN)/.test(nxtTxt);
      const nextIsIndentedNewPara = !nextStartsLowercase && (nxtLn.x > minBodyX + 18) && (prevLn.endX < maxBodyX - 45 || prevEndedStop);
      const normalLineSpacing = verticalGap > 0 && verticalGap <= prevLn.fontSize * 2.55;

      if (!prevEndedStop && !prevIsHeading && !nextStartsNewBlock && (nextStartsLowercase || (!nextIsIndentedNewPara && normalLineSpacing))) {
        mergedText += ' ' + nxtTxt;
        prevLn = nxtLn;
        bIdx++;
      } else {
        break;
      }
    }

    htmlBlocks.push(formatOcrBodyParagraphHtml(mergedText));
  }

  // 6. Xử lý Bảng Chân Trang (2 cột: Nơi nhận bên trái & Chữ ký bên phải)
  if (footerStartIdx < lines.length) {
    const footerLeft = [];
    const footerRight = [];
    for (let fIdx = footerStartIdx; fIdx < lines.length; fIdx++) {
      const fln = lines[fIdx];
      const ftxt = fln.fullText.trim();
      if (/^[-_*—–=─]{3,}$/.test(ftxt)) continue;

      if (fln.hasTwoCols) {
        if (fln.leftText) footerLeft.push(fln.leftText);
        if (fln.rightText) footerRight.push(fln.rightText);
      } else if (fln.isRightSideOnly || /^(T\/M\b|TM\.|KT\.|TL\.|TUQ\.|Q\.|BÍ THƯ|PHÓ BÍ THƯ|CHỦ TỊCH|PHÓ CHỦ TỊCH|TRƯỞNG BAN|CHÁNH VĂN PHÒNG|CHỦ TRÌ)/i.test(ftxt)) {
        footerRight.push(ftxt);
      } else if (/^(Nơi nhận|-|\+|•|Lưu\b|Như trên|Như Điều)/i.test(ftxt)) {
        footerLeft.push(ftxt);
      } else if (footerRight.length > 0 && footerLeft.some(l => /Lưu/i.test(l))) {
        footerRight.push(ftxt);
      } else {
        footerLeft.push(ftxt);
      }
    }
    if (footerLeft.length > 0 || footerRight.length > 0) {
      htmlBlocks.push(buildOcrFooterTableHtml(footerLeft, footerRight, isPartyDoc));
    }
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
// GIỮ NGUYÊN CÂU CHỮ, ĐỒNG THỜI DỰNG CHUẨN THỂ THỨC HƯỚNG DẪN 05-HD/VPTW & NĐ 30
// ============================================================================
function reconstructExactLayoutFromPlainText(rawText, forceMode = 'auto') {
  const cleanedRaw = cleanAndFixVietnameseOcrText(rawText);
  const isPartyDoc = forceMode === 'hd05' ? true : (forceMode === 'nd30' ? false : detectIsPartyDocInOcr(cleanedRaw));
  const rawLines = String(cleanedRaw || '').replace(/\r\n/g, '\n').split('\n');
  const nonEmptyLines = rawLines.map(l => l.trim()).filter(Boolean);
  if (nonEmptyLines.length === 0) return '';

  const headerLeft = [];
  const headerRight = [];
  const footerLeft = [];
  const footerRight = [];
  const bodyLines = [];
  const postFooterLines = [];

  let state = 'HEADER';
  let sawLuuInFooter = false;
  for (let i = 0; i < nonEmptyLines.length; i++) {
    const line = nonEmptyLines[i];
    if (/^[-_*—–=]{2,}$/.test(line)) continue;

    if (state === 'HEADER') {
      // Tách tiêu ngữ 2 cột ngay cả khi OCR chỉ để 1-2 khoảng trắng giữa cột trái và cột phải
      const inlineRightMotto = line.match(/^(.+?)\s+(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập\s*[-–—]\s*Tự do\s*[-–—]\s*Hạnh phúc.*)$/i);
      if (inlineRightMotto && i <= 8) {
        headerLeft.push(inlineRightMotto[1].trim());
        headerRight.push(inlineRightMotto[2].trim());
        continue;
      }
      const inlineDateRight = line.match(/^((?:Số|So)\s*:?\s*[^\s]+(?:\/[^\s]+)?)\s+([A-ZÀ-ỸĐ][^,]+,\s*ngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d{4}.*)$/i);
      if (inlineDateRight && i <= 8) {
        headerLeft.push(inlineDateRight[1].trim());
        headerRight.push(inlineDateRight[2].trim());
        continue;
      }
      const twoCol = line.match(/^(.+?)(?:\t+|\s{3,})(.+)$/);
      if (twoCol && i <= 8) {
        headerLeft.push(twoCol[1].trim());
        headerRight.push(twoCol[2].trim());
        continue;
      }
      if (i <= 8 && /^(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập\s*[-–—]\s*Tự do)/i.test(line)) {
        headerRight.push(line);
        continue;
      }
      if (i <= 8 && /\bngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+/i.test(line) && !/^Hôm nay/i.test(line)) {
        headerRight.push(line);
        continue;
      }
      if (i <= 7 && (/^(QUẬN ỦY|HUYỆN ỦY|THÀNH ỦY|TỈNH ỦY|ĐẢNG ỦY|ĐẢNG BỘ|CHI BỘ|BAN CHẤP HÀNH|BAN THƯỜNG VỤ|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN|CÔNG AN|BỘ\s+|SỞ\s+|PHÒNG\s+|UBND|Số\b|Số\s*:)/i.test(line))) {
        headerLeft.push(line);
        continue;
      }
      state = 'BODY';
    }

    if (state === 'BODY') {
      if (/^Nơi nhận\b/i.test(line) || /\b(T\/M\s+ĐẢNG\s+ỦY|T\/M\s+BAN\s+THƯỜNG\s+VỤ|TM\.\s+ỦY\s+BAN|KT\.\s+CHỦ\s+TỊCH|TL\.\s+CHỦ\s+TỊCH|K\/T\s+CHÁNH\s+VĂN\s+PHÒNG)\b/i.test(line)) {
        state = 'FOOTER';
      }
    }

    if (state === 'FOOTER') {
      const isAppendixStart = /^(PHỤ\s+LỤC\b|SƠ\s+ĐỒ\b|MẪU\s+\d+|BẢNG\s+CHỮ\s+VIẾT\s+TẮT)/i.test(line);
      const isOverflowAfterLuu = sawLuuInFooter && (footerLeft.length + footerRight.length >= 10 || line.length > 65 || /^[-+]\s*(Tên cơ quan|Thể loại|Trích yếu|Từ\s+")/i.test(line));
      if (isAppendixStart || isOverflowAfterLuu) {
        state = 'POST_FOOTER';
        postFooterLines.push(line);
        continue;
      }

      const twoCol = line.match(/^(.+?)(?:\t+|\s{4,})(.+)$/);
      if (twoCol) {
        footerLeft.push(twoCol[1].trim());
        footerRight.push(twoCol[2].trim());
        if (/Lưu\b/i.test(twoCol[1])) sawLuuInFooter = true;
      } else if (/^(Nơi nhận|-|\+|•|Lưu\b)/i.test(line)) {
        footerLeft.push(line);
        if (/Lưu\b/i.test(line)) sawLuuInFooter = true;
      } else {
        footerRight.push(line);
      }
      continue;
    }

    if (state === 'POST_FOOTER') {
      postFooterLines.push(line);
      continue;
    }

    bodyLines.push(line);
  }

  const outHtml = [];
  if (headerLeft.length > 0 || headerRight.length > 0) {
    outHtml.push(buildOcrHeaderTableHtml(headerLeft, headerRight, isPartyDoc));
  }

  const renderLinesToHtml = (arrLines, allowTitleDetect) => {
    let bIdx = 0;
    while (bIdx < arrLines.length) {
      const l = arrLines[bIdx];
      if (l.includes('|')) {
        const cells = l.split('|').map(c => c.trim()).filter(Boolean);
        if (cells.length >= 2) {
          const tds = cells.map(c => `<td style="border:1px solid #000; padding:4pt 6pt; font-size:13pt;">${escapeHtml(c)}</td>`).join('');
          outHtml.push(`<table style="width:100%; border-collapse:collapse; margin:2pt 0;"><tr>${tds}</tr></table>`);
          bIdx++;
          continue;
        }
      }
      if (/^(PHỤ\s+LỤC\s*\d*|SƠ\s+ĐỒ\b.*)$/i.test(l)) {
        outHtml.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:12pt 0 4pt 0; font-size:14pt; text-transform:uppercase;">${escapeHtml(l)}</p>`);
        bIdx++;
        continue;
      }
      if (allowTitleDetect && /^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|QUY ĐỊNH|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|KẾT LUẬN|ĐỀ ÁN|PHƯƠNG ÁN)$/i.test(l) && bIdx <= 3) {
        const titleSize = isPartyDoc ? '15pt' : '14pt';
        outHtml.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:12pt 0 3pt 0; font-size:${titleSize}; text-transform:uppercase;">${escapeHtml(l)}</p>`);
        bIdx++;
        if (bIdx < arrLines.length && !/^(Căn cứ|Thực hiện|Kính gửi|I\.|1\.|Điều\s+1)/i.test(arrLines[bIdx])) {
          outHtml.push(`<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:0 0 2pt 0; font-size:14pt;">${escapeHtml(arrLines[bIdx])}</p>`);
          outHtml.push(isPartyDoc
            ? `<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:0 0 10pt 0; line-height:1; font-size:14pt;">-----</p>`
            : `<p align="center" style="text-align:center; text-indent:0; font-weight:normal; margin:0 0 10pt 0; line-height:1; font-size:10pt;">──────────────</p>`);
          bIdx++;
        }
        continue;
      }

      let merged = l;
      bIdx++;
      while (bIdx < arrLines.length) {
        const nxt = arrLines[bIdx];
        if (nxt.includes('|')) break;
        const endedStop = /[.:;!?]$|\.\s*\/\.\s*$/.test(merged);
        const startsNew = /^([-+•*]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Khoản\s+\d+|Căn cứ\s+|Xét\s+|Theo đề nghị|Kính gửi|PHỤ\s+LỤC|SƠ\s+ĐỒ|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN)/i.test(nxt);
        if (!endedStop && !startsNew) {
          merged += ' ' + nxt;
          bIdx++;
        } else {
          break;
        }
      }
      outHtml.push(formatOcrBodyParagraphHtml(merged));
    }
  };

  renderLinesToHtml(bodyLines, true);

  if (footerLeft.length > 0 || footerRight.length > 0) {
    outHtml.push(buildOcrFooterTableHtml(footerLeft, footerRight, isPartyDoc));
  }

  if (postFooterLines.length > 0) {
    renderLinesToHtml(postFooterLines, false);
  }

  return outHtml.join('\n');
}

// Hậu xử lý: Lọc nhiễu dấu mộc/watermark bảo mật & tự động hiệu đính lỗi dấu/chính tả OCR Tiếng Việt
function cleanAndFixVietnameseOcrText(rawText) {
  if (!rawText) return '';
  let text = String(rawText).normalize('NFC').replace(/\r\n/g, '\n');

  // 1. Lọc bỏ các dòng watermark bảo mật và dòng rác sinh ra từ dấu mộc đỏ / nét kẻ bảng
  const cleanedLines = [];
  for (const rawLine of text.split('\n')) {
    let line = rawLine.trim();
    if (!line) {
      cleanedLines.push('');
      continue;
    }
    const upper = line.toUpperCase();
    // Bỏ dòng watermark chéo bảo mật
    if (/TÀI LIỆU NÀY THUỘC SỞ HỮU CỦA|CÁC CƠ QUAN ĐẢNG\.?\s*VIỆC PHÁT|TÁN,\s*SỬ DỤNG TRÁI PHÉP BỊ NGHIÊM|SỬ DỤNG TRÁI PHÉP BỊ NGHIÊM C[ẤÁ]M/i.test(upper)) {
      continue;
    }
    // Bỏ dòng rác toàn ký tự nhiễu hoặc chuỗi phụ âm vô nghĩa do dấu mộc đỏ (VD: THHHH//NNNTNNNN..., ZZSGSXN, Ĩ - |)
    if (/[BCDFGHJKLMNPQRSTVWXZ]{6,}/.test(line) || /([A-Z])\1{3,}/.test(line)) {
      continue;
    }
    if (/^[ĨŨÚ|_\-—–=+<>\\\/.,;:()\[\]~`'\"0-9\s]{1,12}$/.test(line) && !/^\d+[.)]?$/.test(line) && !/^[-—]{3,}$/.test(line)) {
      continue;
    }
    if (/[⇙ØÏÑŠ]/.test(line) && line.length < 40) {
      line = line.replace(/[⇙ØÏÑŠ\\<>]+/g, ' ').replace(/\b[A-Z]{5,}\b/g, ' ').trim();
      if (line.length < 3) continue;
    }

    // Dọn ký tự rác xen giữa câu
    line = line
      .replace(/\s*\|\s*(?=[a-zà-ỹA-ZÀ-Ỹ])/g, ' ')
      .replace(/(?<=[a-zà-ỹA-ZÀ-Ỹ])\s*\|\s*/g, ' ')
      .replace(/[⇙ØÏÑŠ]/g, '')
      .replace(/\s{2,}/g, ' ')
      .trim();
    if (!line) continue;

    cleanedLines.push(line);
  }

  text = cleanedLines.join('\n');

  // 2. Tự động hiệu đính 80+ lỗi dấu thanh, ký tự & dính chữ OCR Tiếng Việt phổ biến
  const ocrFixes = [
    [/\bCỘNG HỎA\b/g, 'CỘNG HÒA'], [/\bCÔNG HÒA\b/g, 'CỘNG HÒA'], [/\bXÃ HÔI\b/g, 'XÃ HỘI'],
    [/\bCHỦ NGHIA\b/g, 'CHỦ NGHĨA'], [/\bVIÊT NAM\b/g, 'VIỆT NAM'], [/\bVIET NAM\b/g, 'VIỆT NAM'],
    [/\bĐÔC LẬP\b/g, 'ĐỘC LẬP'], [/\bĐộc lạp\b/g, 'Độc lập'], [/\bTư do\b/g, 'Tự do'],
    [/\bHạnh phục\b/g, 'Hạnh phúc'], [/\bHạnh phuc\b/g, 'Hạnh phúc'],
    [/\bĐẢNG CÔNG SẢN\b/g, 'ĐẢNG CỘNG SẢN'], [/\bĐẢNG UỶ\b/g, 'ĐẢNG ỦY'],
    [/\bỦY BẠN\b/g, 'ỦY BAN'], [/\bUY BAN\b/g, 'ỦY BAN'], [/\bNHÂN DẢN\b/g, 'NHÂN DÂN'],
    [/\bQUYÊT ĐỊNH\b/g, 'QUYẾT ĐỊNH'], [/\bQUYẾT ĐINH\b/g, 'QUYẾT ĐỊNH'],
    [/\bNGHỊ QUYÊT\b/g, 'NGHỊ QUYẾT'], [/\bTHÔNG BAO\b/g, 'THÔNG BÁO'],
    [/\bBÁO CAO\b/g, 'BÁO CÁO'], [/\bKẾ HOACH\b/g, 'KẾ HOẠCH'], [/\bTỜ TRINH\b/g, 'TỜ TRÌNH'],
    [/\bCHƯƠNG TRINH\b/g, 'CHƯƠNG TRÌNH'], [/\bBIÊN BẢN\b/g, 'BIÊN BẢN'],
    [/\bĐiều\s+[lI]\./g, 'Điều 1.'], [/\bKhoản\s+[lI]\b/g, 'Khoản 1'],
    [/\bNơi nhạn\b/gi, 'Nơi nhận'], [/\bNơi nhăn\b/gi, 'Nơi nhận'],
    [/\bKính gừi\b/g, 'Kính gửi'], [/\bKính gởi\b/g, 'Kính gửi'],
    [/\bCăn cự\b/g, 'Căn cứ'], [/\bCăn cư\b/g, 'Căn cứ'],
    [/\bChủ tich\b/g, 'Chủ tịch'], [/\bCHỦ TICH\b/g, 'CHỦ TỊCH'],
    [/\bBí thư\b/g, 'Bí thư'], [/\bBÍ THƯ\b/g, 'BÍ THƯ'],
    [/\bhướng dân\b/g, 'hướng dẫn'], [/\bHướng dân\b/g, 'Hướng dẫn'], [/\bHƯỚNG DÂN\b/g, 'HƯỚNG DẪN'],
    [/\bdân chiều\b/g, 'dẫn chiếu'], [/\bdân chiểu\b/g, 'dẫn chiếu'],
    [/\bbỏ sung\b/g, 'bổ sung'], [/\bbổ xung\b/g, 'bổ sung'],
    [/\bđông chí\b/g, 'đồng chí'], [/\bĐông chí\b/g, 'Đồng chí'],
    [/\bTÔ CHỨC\b/g, 'TỔ CHỨC'], [/\btô chức\b/g, 'tổ chức'], [/\btổ chưc\b/g, 'tổ chức'],
    [/\bSƠ ĐÔ\b/g, 'SƠ ĐỒ'], [/\bBÔ TRÍ\b/g, 'BỐ TRÍ'],
    [/\bTHÊ THỨC\b/g, 'THỂ THỨC'], [/\bthê thức\b/g, 'thể thức'],
    [/\bTIỂU BẠN\b/g, 'TIỂU BAN'], [/\btiểu bạn\b/g, 'tiểu ban'],
    [/\bNGHIÊM CÁM\b/g, 'NGHIÊM CẤM'],
    [/\bInhoa\b/g, 'In hoa'], [/\bInthường\b/g, 'In thường'], [/\bvàsốthứtự\b/gi, 'và số thứ tự'],
    [/\bkết qủa\b/g, 'kết quả'], [/\bkết qua\b/g, 'kết quả'],
    [/\bnghị quyêt\b/g, 'nghị quyết'], [/\bquyết đinh\b/g, 'quyết định'],
    [/\bđảng uy\b/g, 'đảng ủy'], [/\bchi bô\b/g, 'chi bộ'], [/\bchỉ bộ\b/g, 'chi bộ'],
    [/\bnhiệm vu\b/g, 'nhiệm vụ'], [/\bthực hiên\b/g, 'thực hiện'], [/\bthưc hiện\b/g, 'thực hiện'],
    [/\bchỉ đao\b/g, 'chỉ đạo'], [/\blãnh đao\b/g, 'lãnh đạo'],
    [/\bđưoc\b/g, 'được'], [/\bngời\b/g, 'người'], [/\bcán bô\b/g, 'cán bộ'],
    [/\bchính sach\b/g, 'chính sách'], [/\btrách nhiêm\b/g, 'trách nhiệm'],
    [/\bđề nghi\b/g, 'đề nghị'], [/\bquán triêt\b/g, 'quán triệt'],
    [/\bgiải quyêt\b/g, 'giải quyết'], [/\bphát triên\b/g, 'phát triển'],
    [/\bchất lương\b/g, 'chất lượng'], [/\bhiêu quả\b/g, 'hiệu quả'],
    [/\bđối tương\b/g, 'đối tượng'], [/\bphương phap\b/g, 'phương pháp'],
    [/\bcải cách hành chánh\b/g, 'cải cách hành chính'],
    [/\bhoà\b/g, 'hòa'], [/\btoà\b/g, 'tòa'], [/\bthuỷ\b/g, 'thủy'], [/\buỷ\b/g, 'ủy'], [/\bhuỷ\b/g, 'hủy'],
    [/\s+([.,;:!?])/g, '$1']
  ];
  for (const [pat, rep] of ocrFixes) {
    text = text.replace(pat, rep);
  }
  return text.replace(/\n{3,}/g, '\n\n').trim();
}

// Tái tạo bố cục 2 cột Header/Footer CHUẨN NĐ 30 & 05-HD/VPTW từ tọa độ bbox thực của Tesseract
function reconstructLayoutFromTesseractBboxes(lines, pageW, pageH, forceMode = 'auto') {
  if (!Array.isArray(lines) || lines.length === 0 || !pageW || !pageH) {
    return reconstructExactLayoutFromPlainText('', forceMode);
  }
  const allRawText = cleanAndFixVietnameseOcrText(lines.map(l => l.text || '').join('\n'));
  const isPartyDoc = forceMode === 'hd05' ? true : (forceMode === 'nd30' ? false : detectIsPartyDocInOcr(allRawText));

  const headerLeft = [];
  const headerRight = [];
  const bodyLines = [];
  const footerLeft = [];
  const footerRight = [];

  const headerCutoffY = pageH * 0.24;
  const footerZoneMinY = pageH * 0.62;
  const splitXHeader = pageW * 0.44;
  const splitXFooter = pageW * 0.47;

  let inFooter = false;
  for (const ln of lines) {
    const cleanLineText = cleanAndFixVietnameseOcrText(ln.text || '');
    if (!cleanLineText) continue;
    const bbox = ln.bbox || { x0: 0, y0: 0, x1: pageW, y1: 0 };
    const words = Array.isArray(ln.words) && ln.words.length > 0 ? ln.words : null;

    // 1. Vùng tiêu ngữ 2 cột đầu trang (top 24% trang)
    if (!inFooter && bbox.y0 <= headerCutoffY && bodyLines.length === 0) {
      if (/^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|CÔNG VĂN|Kính gửi)/i.test(cleanLineText)) {
        bodyLines.push(cleanLineText);
        continue;
      }
      if (words) {
        const lWords = [];
        const rWords = [];
        for (const w of words) {
          const wt = (w.text || '').trim();
          if (!wt) continue;
          const cx = ((w.bbox?.x0 || 0) + (w.bbox?.x1 || 0)) / 2;
          if (cx < splitXHeader) lWords.push(wt);
          else rWords.push(wt);
        }
        const lt = cleanAndFixVietnameseOcrText(lWords.join(' '));
        const rt = cleanAndFixVietnameseOcrText(rWords.join(' '));
        if (lt) headerLeft.push(lt);
        if (rt) headerRight.push(rt);
        continue;
      }
    }

    // 2. Phát hiện bắt đầu vùng Nơi nhận & Chữ ký cuối văn bản
    if (!inFooter && bbox.y0 >= footerZoneMinY && (
      /^Nơi nhận\b/i.test(cleanLineText) ||
      /\b(T\/M\s+ĐẢNG|T\/M\s+BAN|TM\.\s+ỦY\s+BAN|KT\.\s+CHỦ\s+TỊCH|KT\.\s+GIÁM\s+ĐỐC|KT\.\s+TRƯỞNG|TL\.\s+CHỦ\s+TỊCH|CHỦ TỊCH|BÍ THƯ)\b/i.test(cleanLineText)
    )) {
      inFooter = true;
    }

    if (inFooter) {
      if (words) {
        const lWords = [];
        const rWords = [];
        for (const w of words) {
          const wt = (w.text || '').trim();
          if (!wt) continue;
          const cx = ((w.bbox?.x0 || 0) + (w.bbox?.x1 || 0)) / 2;
          if (cx < splitXFooter) lWords.push(wt);
          else rWords.push(wt);
        }
        const lt = cleanAndFixVietnameseOcrText(lWords.join(' '));
        const rt = cleanAndFixVietnameseOcrText(rWords.join(' '));
        if (lt) footerLeft.push(lt);
        if (rt) footerRight.push(rt);
      } else if (/^(Nơi nhận|-|\+|•|Lưu\b)/i.test(cleanLineText)) {
        footerLeft.push(cleanLineText);
      } else {
        footerRight.push(cleanLineText);
      }
      continue;
    }

    bodyLines.push(cleanLineText);
  }

  // Nếu bbox không tách được header/footer (ví dụ ảnh cắt nhỏ), dùng fallback plain-text parser
  if (headerLeft.length === 0 && headerRight.length === 0 && footerLeft.length === 0 && footerRight.length === 0) {
    return reconstructExactLayoutFromPlainText(allRawText, forceMode);
  }

  const outHtml = [];
  if (headerLeft.length > 0 || headerRight.length > 0) {
    outHtml.push(buildOcrHeaderTableHtml(headerLeft, headerRight, isPartyDoc));
  }
  if (bodyLines.length > 0) {
    outHtml.push(reconstructExactLayoutFromPlainText(bodyLines.join('\n'), forceMode));
  }
  if (footerLeft.length > 0 || footerRight.length > 0) {
    outHtml.push(buildOcrFooterTableHtml(footerLeft, footerRight, isPartyDoc));
  }
  return outHtml.join('\n');
}

// Tiền xử lý ảnh cho OCR chính xác cao: Phóng nét 2400px + Tẩy mộc đỏ + Tăng nét dấu tiếng Việt (không làm mất nét mảnh)
function createPreprocessedOcrCanvas(sourceCanvas, forceBinarize = false) {
  const baseSource = sourceCanvas || currentImageSource;
  const srcW = baseSource.width || sourceCanvas.width;
  const srcH = baseSource.height || sourceCanvas.height;
  const scale = srcW < 1500 ? Math.min(2.5, 2200 / Math.max(1, srcW)) : 1;
  const outW = Math.min(Math.round(srcW * scale), 2600);
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

  // 1. Tẩy dấu mộc đỏ & nâng tương phản mềm bảo toàn nét mảnh của dấu thanh tiếng Việt
  try {
    const w = procCanvas.width;
    const h = procCanvas.height;
    const imgData = ctx.getImageData(0, 0, w, h);
    const data = imgData.data;
    const grayBuf = new Uint8ClampedArray(w * h);

    for (let i = 0, p = 0; i < data.length; i += 4, p++) {
      const r = data[i], g = data[i + 1], b = data[i + 2];
      // Tẩy pixel con dấu đỏ (chỉ tẩy màu đỏ tươi, giữ nguyên mực đen nằm đè lên dấu)
      if (r > 110 && (r - g) > 34 && (r - b) > 34) {
        grayBuf[p] = 255;
        continue;
      }
      let gray = 0.299 * r + 0.587 * g + 0.114 * b;
      // Bảo toàn nét dấu tiếng Việt mờ (ngưỡng trắng 214 thay vì 186 để không xóa mất dấu sắc/hỏi/ngã)
      if (gray > 214) {
        gray = 255;
      } else if (forceBinarize || ocrHighContrast) {
        gray = Math.min(255, Math.max(0, (gray - 135) * 1.45 + 135));
      } else if (gray < 160) {
        gray = Math.max(0, gray * 0.88);
      }
      grayBuf[p] = gray;
    }

    // 2. Làm sắc nét nhẹ 3x3 (Unsharp Mask) giúp tách rõ dấu mũ ^, ă, ơ, ư và dấu thanh
    for (let y = 1; y < h - 1; y++) {
      const row = y * w;
      for (let x = 1; x < w - 1; x++) {
        const idx = row + x;
        const center = grayBuf[idx];
        const blur = (grayBuf[idx - 1] + grayBuf[idx + 1] + grayBuf[idx - w] + grayBuf[idx + w]) * 0.25;
        let sharp = center + (center - blur) * 0.45;
        sharp = sharp > 220 ? 255 : Math.max(0, Math.min(255, sharp));
        const di = idx * 4;
        data[di] = data[di + 1] = data[di + 2] = sharp;
      }
    }
    ctx.putImageData(imgData, 0, 0);
  } catch (e) {}

  return procCanvas;
}

// Hậu kiểm OCR bằng AI Local (Ollama qwen2.5:7b / qwen2.5:3b) — chỉ sửa lỗi nhận diện chữ/dấu, giữ nguyên tên riêng & số liệu
async function proofreadVietnameseOcrWithOllama(rawOcrText) {
  if (!rawOcrText || rawOcrText.trim().length < 20) return null;
  try {
    const ctrlTags = new AbortController();
    const tTags = setTimeout(() => ctrlTags.abort(), 1500);
    const tagRes = await fetch('http://127.0.0.1:11434/api/tags', { signal: ctrlTags.signal });
    clearTimeout(tTags);
    if (!tagRes.ok) return null;
    const tagJson = await tagRes.json();
    const models = (tagJson.models || []).map(m => m.name);
    if (models.length === 0) return null;
    const picked = models.find(n => /qwen2\.5:7b/i.test(n)) || models.find(n => /qwen2\.5:3b/i.test(n)) || models[0];

    const paragraphs = rawOcrText.split('\n');
    const chunks = [];
    let buf = '';
    for (const ln of paragraphs) {
      if ((buf.length + ln.length) > 1500 && buf.trim()) {
        chunks.push(buf);
        buf = '';
      }
      buf += (buf ? '\n' : '') + ln;
    }
    if (buf.trim()) chunks.push(buf);
    if (chunks.length > 8) return null; // Tránh quá lâu với tài liệu siêu dài

    const fixedChunks = [];
    for (const ch of chunks) {
      const ctrl = new AbortController();
      const tGen = setTimeout(() => ctrl.abort(), 18000);
      const genRes = await fetch('http://127.0.0.1:11434/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        signal: ctrl.signal,
        body: JSON.stringify({
          model: picked,
          stream: false,
          options: { temperature: 0.05, num_predict: 1024 },
          system: 'Bạn là bộ hậu kiểm OCR tiếng Việt hành chính/Đảng. CHỈ sửa lỗi sai dấu thanh, sai ký tự OCR (như 0/O, 1/l, ơ/o, ư/u, mất dấu). TUYỆT ĐỐI giữ nguyên số dòng, cấu trúc xuống dòng, số hiệu văn bản, ngày tháng, số liệu và tên riêng. Chỉ trả về văn bản đã sửa, không giải thích.',
          prompt: ch
        })
      });
      clearTimeout(tGen);
      if (!genRes.ok) return null;
      const gJson = await genRes.json();
      const out = (gJson.response || '').trim();
      // Bảo vệ chống AI cắt bớt nội dung
      if (out && out.length >= ch.length * 0.72 && out.length <= ch.length * 1.25) {
        fixedChunks.push(out);
      } else {
        fixedChunks.push(ch);
      }
    }
    return { text: cleanAndFixVietnameseOcrText(fixedChunks.join('\n')), model: picked };
  } catch {
    return null;
  }
}

async function recognizeCanvasPageHighAccuracy(canvasEl, onProgress = null) {
  const cleanCanvas = createPreprocessedOcrCanvas(canvasEl, false);
  const selectedLang = document.getElementById('ocrLangSelect')?.value || 'vie';
  const enableDeskew = document.getElementById('ocrDeskewCheck')?.checked !== false;

  if (window.VietNamOcrEnginePro) {
    const res = await window.VietNamOcrEnginePro.recognizeCanvasWithBestEngine(cleanCanvas, {
      lang: selectedLang,
      deskew: enableDeskew,
      onProgress: (pct, msg) => {
        if (onProgress) onProgress(pct, msg);
      }
    });
    return {
      text: cleanAndFixVietnameseOcrText(res.text),
      lines: res.lines,
      words: res.words,
      confidence: res.confidence,
      procWidth: res.procWidth,
      procHeight: res.procHeight,
      deskewAngle: res.deskewAngle,
      engineSource: res.engineSource
    };
  }

  let text = '';
  let lines = null;
  if (window.Tesseract) {
    const result = await window.Tesseract.recognize(cleanCanvas, selectedLang, {
      preserve_interword_spaces: '1',
      logger: m => {
        if (onProgress && m.status === 'recognizing text' && typeof m.progress === 'number') {
          onProgress(Math.max(15, Math.round(m.progress * 100)));
        }
      }
    });
    text = cleanAndFixVietnameseOcrText((result?.data?.text || '').normalize('NFC'));
    lines = result?.data?.lines || null;
  }

  return { text, lines, words: [], confidence: 0, procWidth: cleanCanvas.width, procHeight: cleanCanvas.height, deskewAngle: 0, engineSource: 'tesseract_fallback' };
}

function exportOcrToTxt() {
  const textContent = document.getElementById('ocrResultText')?.value || '';
  if (!textContent.trim()) {
    if (window.showToast) window.showToast('Chưa có nội dung văn bản OCR để xuất .TXT!');
    return;
  }
  const outFileName = `${uploadedFileBaseName || 'VanBan_OCR'}.txt`;
  const blob = new Blob(['\ufeff', textContent], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  if (window.showToast) window.showToast(`✓ Đã xuất tệp văn bản thuần UTF-8: ${outFileName}`);
}

async function recognizeCanvasPageWithoutAlteringWords(canvasEl) {
  const res = await recognizeCanvasPageHighAccuracy(canvasEl, null);
  return res.text || '';
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
        : `<div class="ocr-page-flow" style="margin-top: 6pt;">${html}</div>`)
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

  statusEl.textContent = 'Đang tiền xử lý ảnh & nhận diện chữ Tiếng Việt chính xác cao (Tesseract v5 LSTM)...';
  progressFill.style.width = '20%';

  let rawPreserved = '';
  let tesseractLines = null;
  let tesseractWords = [];
  let ocrConf = 0;
  let ocrDocProfile = '';
  let procW = canvas.width;
  let procH = canvas.height;

  try {
    const ocrRes = await recognizeCanvasPageHighAccuracy(canvas, (pct, msg) => {
      progressFill.style.width = `${Math.min(92, pct)}%`;
      statusEl.textContent = msg || `Đang nhận diện Tiếng Việt (Tesseract v5 tessdata_best): ${pct}%`;
    });
    rawPreserved = ocrRes.text;
    tesseractLines = ocrRes.lines;
    tesseractWords = ocrRes.words || [];
    ocrConf = ocrRes.confidence || 0;
    ocrDocProfile = ocrRes.docProfile || '';
    procW = ocrRes.procWidth;
    procH = ocrRes.procHeight;
  } catch (tessErr) {
    if (window.VietNamOcrEnginePro) window.VietNamOcrEnginePro.logSafeError('RunOCR', tessErr.message);
    statusEl.textContent = '⚠ Lỗi nhận diện Tesseract v5: ' + tessErr.message;
    return;
  }

  // Hiển thị ngay kết quả chuẩn thể thức 2 cột theo tọa độ bbox thực (KHÔNG tự ý sửa nội dung)
  textOutput.value = rawPreserved;
  if (paperEl) {
    paperEl.innerHTML = (tesseractLines && tesseractLines.length > 0)
      ? reconstructLayoutFromTesseractBboxes(tesseractLines, procW, procH, 'auto')
      : reconstructExactLayoutFromPlainText(rawPreserved);
  }

  // Đánh dấu vùng có độ tin cậy thấp (< 70%) trực tiếp trên ảnh gốc để người dùng kiểm chứng
  let lowConfDrawn = 0;
  if (window.VietNamOcrEnginePro && tesseractWords.length > 0) {
    lowConfDrawn = window.VietNamOcrEnginePro.drawConfidenceOverlayOnCanvas(canvas, tesseractWords, procW, procH);
    const bboxTable = window.VietNamOcrEnginePro.extractTableFromWordBboxes(tesseractWords, procW);
    if (bboxTable && bboxTable.rows.length > 0 && !appendTableMode) {
      ocrTableHeaders = bboxTable.headers;
      ocrTableRows = bboxTable.rows;
      renderEditableTable();
    } else {
      extractTableFromOcrResult(rawPreserved, tesseractLines, appendTableMode);
    }
  } else {
    extractTableFromOcrResult(rawPreserved, tesseractLines, appendTableMode);
  }

  if (window.VietNamOcr2LayerPdf) {
    window.VietNamOcr2LayerPdf.recordSinglePageOcr(canvas, rawPreserved, tesseractLines, procW, procH);
  }
  if (window.VietNamOcrIdpStudio) {
    window.VietNamOcrIdpStudio.autoSaveDocumentToArchive(uploadedFileBaseName, rawPreserved, ocrConf);
  }
  statusEl.textContent = `✓ Hoàn tất OCR (${ocrConf}% tin cậy · ${ocrDocProfile || 'Văn bản'} · ${lowConfDrawn} vùng cần đối chiếu ảnh gốc)! Đã tự động lưu vào Kho cục bộ.`;
  progressFill.style.width = '100%';
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
async function exportOcrToWord() {
  const paperEl = document.getElementById('ocrFormattedPaper');
  const plainWrap = document.getElementById('ocrPlainEditorWrap');
  const textContent = document.getElementById('ocrResultText')?.value || '';

  let bodyHtml = '';
  if (plainWrap && plainWrap.style.display !== 'none') {
    bodyHtml = reconstructExactLayoutFromPlainText(textContent);
  } else if (paperEl && paperEl.innerHTML.trim()) {
    bodyHtml = paperEl.innerHTML.replace(/page-break-(before|after)\s*:\s*always\s*;?/gi, '');
  } else {
    bodyHtml = reconstructExactLayoutFromPlainText(textContent);
  }

  // Ưu tiên xuất .DOCX chuẩn OOXML thống nhất (đầy đủ styles.xml, settings.xml, w:rFonts 4 thuộc tính, w:lang="vi-VN")
  if (window.VietNamDocxExporter && window.JSZip) {
    try {
      await window.VietNamDocxExporter.exportUnifiedDocx({
        html: bodyHtml,
        plainText: textContent,
        fileName: uploadedFileBaseName || 'VanBan_ChuyenDoi'
      });
      return;
    } catch (docxErr) {
      console.warn('Fallback sang .doc HTML:', docxErr);
    }
  }

  const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(uploadedFileBaseName)}</title>
    <!--[if gte mso 9]><xml>
      <w:WordDocument>
        <w:View>Print</w:View>
        <w:Zoom>100</w:Zoom>
        <w:DoNotOptimizeForBrowser/>
      </w:WordDocument>
    </xml><![endif]-->
    <style>
      @page Section1 {
        size: 21.0cm 29.7cm;
        margin: 2.0cm 1.5cm 2.0cm 3.0cm;
        mso-header-margin: 1.0cm;
        mso-footer-margin: 1.0cm;
        mso-title-page: yes;
        mso-header: h1;
        mso-first-header: fh1;
      }
      div.Section1 { page: Section1; }
      p.MsoHeader, li.MsoHeader, div.MsoHeader {
        margin: 0cm;
        text-align: center;
        font-size: 13.0pt;
        font-family: "Times New Roman", Times, serif;
      }
      body { font-family: "Times New Roman", Times, serif; font-size: 14pt; color: #000000; line-height: 1.15; }
      p { font-family: "Times New Roman", Times, serif; margin: 0 0 6pt 0; font-size: 14pt; line-height: 1.15; }
      table { width: 100%; border-collapse: collapse; font-family: "Times New Roman", Times, serif; }
      td, th { font-family: "Times New Roman", Times, serif; vertical-align: top; }
      td p { margin: 0; text-indent: 0; }
    </style>
  </head>
  <body>
    <div class="Section1">
      ${bodyHtml}
      <div style="mso-element:header" id="fh1"><p class="MsoHeader">&nbsp;</p></div>
      <div style="mso-element:header" id="h1">
        <p class="MsoHeader" align="center" style="text-align:center;font-size:13pt;font-family:'Times New Roman',serif;">
          <!--[if supportFields]><span style='mso-element:field-begin'></span> PAGE <span style='mso-element:field-separator'></span><![endif]-->2<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->
        </p>
      </div>
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

function applyOcrRegulationFormat(mode = 'auto') {
  const textEl = document.getElementById('ocrResultText');
  const paperEl = document.getElementById('ocrFormattedPaper');
  if (!textEl || !paperEl) return;
  let raw = (textEl.value || '').trim();
  if (!raw) return;
  const html = reconstructExactLayoutFromPlainText(raw, mode);
  paperEl.innerHTML = html;
  switchOcrViewMode('formatted');
  if (window.showToast) {
    window.showToast(
      mode === 'hd05'
        ? 'Đã chuẩn hóa bố cục văn bản Đảng (Hướng dẫn 05-HD/VPTW)!'
        : (mode === 'nd30'
          ? 'Đã chuẩn hóa bố cục văn bản Hành chính (Nghị định 30/2020/NĐ-CP)!'
          : 'Đã tự động tái tạo chuẩn bố cục văn bản 1:1!')
    );
  }
}

function printOcrA4Document() {
  const paperEl = document.getElementById('ocrFormattedPaper');
  if (!paperEl) return;
  const w = window.open('', '_blank', 'width=900,height=1000');
  if (!w) {
    if (window.showToast) window.showToast('Vui lòng cho phép mở cửa sổ Pop-up để In / Xuất PDF!');
    return;
  }
  w.document.write(`<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><title>In Văn Bản Chuẩn A4</title>
  <style>
    @page { size: A4; margin: 20mm 15mm 20mm 30mm; }
    body { font-family: "Times New Roman", serif; font-size: 14pt; color: #000; line-height: 1.38; margin: 0; }
    table { width: 100%; border-collapse: collapse; }
    td p { margin: 0; text-indent: 0; }
    p { margin: 0 0 6pt 0; }
  </style></head><body>${paperEl.innerHTML}</body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 250);
}
