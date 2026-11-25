/**
 * VietNamOCR Pro — Bộ Công Cụ Xử Lý PDF Chuyên Nghiệp (js/ocr-pdf-toolkit.js)
 *
 * Chạy 100% cục bộ trên trình duyệt bằng pdf-lib + PDF.js (không tải file lên máy chủ ngoài, không ghi đè file gốc):
 * - Ghép nhiều file PDF (Merge)
 * - Tách PDF theo dải trang (Split: VD 1-3, 5)
 * - Xóa trang / Xoay trang (90°, 180°, 270°)
 * - Đánh số trang tự động (từ trang 2 hoặc mọi trang)
 * - Đóng dấu Watermark / "BẢN SAO" / "SAO Y BẢN CHÍNH"
 * - Trích xuất toàn bộ văn bản & Trích xuất hình ảnh từng trang PDF
 */
(function (global) {
  'use strict';

  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function setPdfStatus(msg) {
    const el = document.getElementById('pdfToolStatus');
    if (el) el.textContent = msg;
  }

  function parsePageRanges(expr, totalPages) {
    const out = [];
    const parts = String(expr || '').split(',').map(s => s.trim()).filter(Boolean);
    for (const p of parts) {
      const m = p.match(/^(\d+)\s*-\s*(\d+)$/);
      if (m) {
        const a = Math.max(1, parseInt(m[1], 10));
        const b = Math.min(totalPages, parseInt(m[2], 10));
        for (let i = a; i <= b; i++) if (!out.includes(i - 1)) out.push(i - 1);
      } else if (/^\d+$/.test(p)) {
        const n = parseInt(p, 10);
        if (n >= 1 && n <= totalPages && !out.includes(n - 1)) out.push(n - 1);
      }
    }
    return out;
  }

  async function mergePdfFiles() {
    const inp = document.getElementById('pdfMergeInput');
    const files = Array.from(inp?.files || []);
    if (files.length < 2) {
      setPdfStatus('⚠ Vui lòng chọn ít nhất 2 tệp PDF để ghép.');
      return;
    }
    try {
      setPdfStatus(`⏳ Đang ghép ${files.length} tệp PDF...`);
      const merged = await global.PDFLib.PDFDocument.create();
      let totalPages = 0;
      for (const f of files) {
        const bytes = await f.arrayBuffer();
        const src = await global.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
        const copied = await merged.copyPages(src, src.getPageIndices());
        copied.forEach(p => merged.addPage(p));
        totalPages += copied.length;
      }
      const outBytes = await merged.save();
      downloadBlob(new Blob([outBytes], { type: 'application/pdf' }), `Ghep_${files.length}_File_${totalPages}Trang.pdf`);
      setPdfStatus(`✓ Đã ghép thành công ${files.length} tệp PDF (tổng ${totalPages} trang) — không ghi đè tệp gốc!`);
    } catch (e) {
      setPdfStatus('⚠ Lỗi ghép PDF: ' + e.message);
    }
  }

  async function splitOrDeletePdf(mode) {
    const inp = document.getElementById('pdfSingleInput');
    const file = inp?.files?.[0];
    const rangeExpr = document.getElementById('pdfPageRangeInput')?.value || '1';
    if (!file) { setPdfStatus('⚠ Vui lòng chọn 1 tệp PDF trước.'); return; }
    try {
      const bytes = await file.arrayBuffer();
      const src = await global.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
      const total = src.getPageCount();
      const selected = parsePageRanges(rangeExpr, total);
      if (selected.length === 0) { setPdfStatus(`⚠ Dải trang không hợp lệ (tệp có ${total} trang, VD: 1-3, 5).`); return; }

      const keepIndices = (mode === 'delete')
        ? src.getPageIndices().filter(i => !selected.includes(i))
        : selected;
      if (keepIndices.length === 0) { setPdfStatus('⚠ Không thể xóa toàn bộ trang của tệp PDF.'); return; }

      const outDoc = await global.PDFLib.PDFDocument.create();
      const pages = await outDoc.copyPages(src, keepIndices);
      pages.forEach(p => outDoc.addPage(p));
      const outBytes = await outDoc.save();
      const base = file.name.replace(/\.pdf$/i, '');
      const suffix = mode === 'delete' ? 'DaXoaTrang' : 'DaTach';
      downloadBlob(new Blob([outBytes], { type: 'application/pdf' }), `${base}_${suffix}.pdf`);
      setPdfStatus(`✓ Đã ${mode === 'delete' ? 'xóa trang' : 'tách PDF'} thành công (còn ${keepIndices.length}/${total} trang)!`);
    } catch (e) {
      setPdfStatus('⚠ Lỗi xử lý trang PDF: ' + e.message);
    }
  }

  async function rotatePdfPages() {
    const file = document.getElementById('pdfSingleInput')?.files?.[0];
    const deg = parseInt(document.getElementById('pdfRotateDegSelect')?.value || '90', 10);
    if (!file) { setPdfStatus('⚠ Vui lòng chọn 1 tệp PDF trước.'); return; }
    try {
      const bytes = await file.arrayBuffer();
      const doc = await global.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
      doc.getPages().forEach(p => {
        const cur = p.getRotation().angle || 0;
        p.setRotation(global.PDFLib.degrees((cur + deg) % 360));
      });
      const outBytes = await doc.save();
      const base = file.name.replace(/\.pdf$/i, '');
      downloadBlob(new Blob([outBytes], { type: 'application/pdf' }), `${base}_Xoay${deg}.pdf`);
      setPdfStatus(`✓ Đã xoay toàn bộ ${doc.getPageCount()} trang thêm ${deg}°!`);
    } catch (e) {
      setPdfStatus('⚠ Lỗi xoay PDF: ' + e.message);
    }
  }

  async function stampWatermarkOrPageNum(mode) {
    const file = document.getElementById('pdfSingleInput')?.files?.[0];
    if (!file) { setPdfStatus('⚠ Vui lòng chọn 1 tệp PDF trước.'); return; }
    try {
      const bytes = await file.arrayBuffer();
      const doc = await global.PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
      const font = await doc.embedFont(global.PDFLib.StandardFonts.HelveticaBold);
      const pages = doc.getPages();
      const wmText = (document.getElementById('pdfWatermarkInput')?.value || 'BAN SAO - LUU HANH NOI BO')
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D');

      pages.forEach((p, idx) => {
        const { width, height } = p.getSize();
        if (mode === 'pagenum') {
          if (idx === 0) return; // Chuẩn NĐ 30: đánh số từ trang 2 ở giữa lề trên
          p.drawText(String(idx + 1), {
            x: width / 2 - 4,
            y: height - 28,
            size: 12,
            font,
            color: global.PDFLib.rgb(0, 0, 0)
          });
        } else if (mode === 'copy_stamp') {
          p.drawRectangle({ x: width - 170, y: height - 56, width: 145, height: 30, borderColor: global.PDFLib.rgb(0.8, 0.1, 0.1), borderWidth: 1.5 });
          p.drawText('SAO Y BAN CHINH', {
            x: width - 158,
            y: height - 46,
            size: 11,
            font,
            color: global.PDFLib.rgb(0.8, 0.1, 0.1)
          });
        } else {
          p.drawText(wmText, {
            x: width * 0.16,
            y: height * 0.45,
            size: 32,
            font,
            color: global.PDFLib.rgb(0.75, 0.15, 0.15),
            opacity: 0.22,
            rotate: global.PDFLib.degrees(35)
          });
        }
      });
      const outBytes = await doc.save();
      const base = file.name.replace(/\.pdf$/i, '');
      const tag = mode === 'pagenum' ? 'DanhSoTrang' : (mode === 'copy_stamp' ? 'DauBanSao' : 'Watermark');
      downloadBlob(new Blob([outBytes], { type: 'application/pdf' }), `${base}_${tag}.pdf`);
      setPdfStatus(`✓ Đã xử lý "${tag}" cho toàn bộ ${pages.length} trang PDF!`);
    } catch (e) {
      setPdfStatus('⚠ Lỗi đóng dấu / đánh số trang PDF: ' + e.message);
    }
  }

  async function extractPdfTextOrImages(mode) {
    const file = document.getElementById('pdfSingleInput')?.files?.[0];
    if (!file || !global.pdfjsLib) { setPdfStatus('⚠ Vui lòng chọn 1 tệp PDF.'); return; }
    try {
      setPdfStatus(`⏳ Đang trích xuất ${mode === 'images' ? 'hình ảnh' : 'văn bản'} từ "${file.name}"...`);
      const buf = await file.arrayBuffer();
      const pdf = await global.pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
      const base = file.name.replace(/\.pdf$/i, '');

      if (mode === 'text') {
        const chunks = [];
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const tc = await page.getTextContent();
          const line = (tc.items || []).map(it => it.str || '').join(' ').replace(/\s{2,}/g, ' ').trim();
          chunks.push(`--- TRANG ${i} ---\n${line}`);
        }
        downloadBlob(new Blob(['\ufeff', chunks.join('\n\n')], { type: 'text/plain;charset=utf-8' }), `${base}_TrichXuatVanBan.txt`);
        setPdfStatus(`✓ Đã trích xuất văn bản của ${pdf.numPages} trang PDF thành tệp .TXT!`);
      } else {
        if (!global.JSZip) { setPdfStatus('⚠ Thư viện JSZip chưa sẵn sàng để nén ảnh.'); return; }
        const zip = new global.JSZip();
        for (let i = 1; i <= Math.min(pdf.numPages, 30); i++) {
          const page = await pdf.getPage(i);
          const vp = page.getViewport({ scale: 1.8 });
          const c = document.createElement('canvas');
          c.width = Math.round(vp.width); c.height = Math.round(vp.height);
          const ctx = c.getContext('2d');
          ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, c.width, c.height);
          await page.render({ canvasContext: ctx, viewport: vp }).promise;
          const b64 = c.toDataURL('image/jpeg', 0.92).split(',')[1];
          zip.file(`Trang_${String(i).padStart(2, '0')}.jpg`, b64, { base64: true });
        }
        const zBlob = await zip.generateAsync({ type: 'blob' });
        downloadBlob(zBlob, `${base}_HinhAnhTrang.zip`);
        setPdfStatus(`✓ Đã trích xuất ảnh từng trang PDF thành gói ZIP!`);
      }
    } catch (e) {
      setPdfStatus('⚠ Lỗi trích xuất PDF: ' + e.message);
    }
  }

  global.VietNamPdfToolkit = {
    mergePdfFiles,
    splitOrDeletePdf,
    rotatePdfPages,
    stampWatermarkOrPageNum,
    extractPdfTextOrImages
  };
})(window);
