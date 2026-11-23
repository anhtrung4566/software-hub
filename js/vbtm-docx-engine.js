/**
 * VBTM Pro 3.0 Ultimate — Native OpenXML (.DOCX) Generator & Reader
 * Tạo tệp .docx chuẩn ISO/IEC 29500 (OOXML ZIP) trực tiếp bằng JavaScript thuần:
 * - Không đổi đuôi HTML thành .doc
 * - Khổ giấy A4 (11906 x 16838 twips), lề Trái 30mm / Phải 15mm / Trên 20mm / Dưới 20mm
 * - Font Times New Roman 14pt, đánh số trang tự động từ trang 2 (w:titlePg)
 * - Bảng Tiêu ngữ 2 cột & Bảng Nơi nhận / Chữ ký chuẩn 05-HD/VPTW & NĐ 30/2020/NĐ-CP
 * - Hỗ trợ nhúng ảnh Mã QR PNG thật (word/media/qr.png)
 */
(function (global) {
  'use strict';

  // Bảng CRC32 chuẩn cho định dạng ZIP OOXML
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) {
        c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      }
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) {
      c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function strToUtf8Bytes(str) {
    return new TextEncoder().encode(String(str || ''));
  }

  function dataUrlToBytes(dataUrl) {
    if (!dataUrl || !dataUrl.includes(',')) return null;
    const b64 = dataUrl.split(',')[1];
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  /**
   * Đóng gói mảng [{ name, data: Uint8Array }] thành tệp ZIP chuẩn (STORE method 0)
   */
  function buildZipBlob(entries) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = strToUtf8Bytes(entry.name);
      const dataBytes = entry.data instanceof Uint8Array ? entry.data : strToUtf8Bytes(entry.data);
      const crc = crc32(dataBytes);
      const size = dataBytes.length;

      // Local file header (30 bytes + name)
      const lh = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true);
      lv.setUint16(4, 20, true);
      lv.setUint16(6, 0x0800, true); // UTF-8 flag
      lv.setUint16(8, 0, true);      // STORE
      lv.setUint16(10, 0, true);
      lv.setUint16(12, 0, true);
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);

      localParts.push(lh, dataBytes);

      // Central directory header (46 bytes + name)
      const ch = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      centralParts.push(ch);

      offset += lh.length + size;
    }

    const centralSize = centralParts.reduce((s, b) => s + b.length, 0);
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, offset, true);
    ev.setUint16(20, 0, true);

    return new Blob([...localParts, ...centralParts, eocd], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  function xmlEsc(s) {
    return String(s || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function makeRun(text, opts) {
    const o = opts || {};
    const sz = Math.round((o.sizePt || 14) * 2); // half-points
    const rPr = [
      `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`,
      o.bold ? '<w:b/>' : '',
      o.italic ? '<w:i/>' : '',
      o.underline ? '<w:u w:val="single"/>' : '',
      `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`,
      '<w:w w:val="100"/>'
    ].filter(Boolean).join('');
    return `<w:r><w:rPr>${rPr}</w:rPr><w:t xml:space="preserve">${xmlEsc(text)}</w:t></w:r>`;
  }

  function makePara(runsXml, opts) {
    const o = opts || {};
    const jc = o.align ? `<w:jc w:val="${o.align}"/>` : '<w:jc w:val="both"/>';
    const spacing = `<w:spacing w:before="${o.before != null ? o.before : 60}" w:after="${o.after != null ? o.after : 60}" w:line="${o.line || 276}" w:lineRule="auto"/>`;
    const ind = o.firstLine ? `<w:ind w:firstLine="${o.firstLine}"/>` : '<w:ind w:firstLine="0"/>';
    const bdr = o.bottomBorder ? '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="000000"/></w:pBdr>' : '';
    return `<w:p><w:pPr>${jc}${spacing}${ind}${bdr}</w:pPr>${runsXml}</w:p>`;
  }

  function makeQrDrawingXml(rId, sizePx) {
    const emu = Math.round((sizePx || 82) * 9525);
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${emu}" cy="${emu}"/><wp:docPr id="1" name="VBTM_QR"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="1" name="qr.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu}" cy="${emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }

  function formatDateLine(location, dateStr) {
    let d = new Date();
    if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      const [y, m, day] = dateStr.split('-').map(Number);
      d = new Date(y, m - 1, day);
    }
    const dd = d.getDate() < 10 ? '0' + d.getDate() : String(d.getDate());
    const mm = (d.getMonth() + 1) <= 2 ? '0' + (d.getMonth() + 1) : String(d.getMonth() + 1);
    const loc = String(location || 'Phường').trim();
    return `${loc}, ngày ${dd} tháng ${mm} năm ${d.getFullYear()}`;
  }

  /**
   * Tạo tệp .docx chuẩn OOXML từ dữ liệu văn bản hiện tại
   */
  function generateDocxBlob(docData) {
    const isParty = docData.standard === 'hd05';
    const govOrg = String(docData.governingOrg || '').trim().toUpperCase();
    const issueOrg = String(docData.issuingOrg || (isParty ? 'ĐẢNG ỦY PHƯỜNG' : 'ỦY BAN NHÂN DÂN PHƯỜNG')).trim().toUpperCase();
    const docNum = String(docData.docNumber || '').trim();
    const numberLabel = docNum ? (docNum.toLowerCase().startsWith('số') ? docNum : (isParty ? `Số ${docNum}` : `Số: ${docNum}`)) : (isParty ? 'Số ...-CV/ĐU' : 'Số: .../UBND-VP');
    const dateLine = formatDateLine(docData.location, docData.docDate);
    const docType = String(docData.docType || '').trim().toUpperCase();
    const docSubject = String(docData.docSubject || '').trim();
    const bodyLines = String(docData.docBody || '').normalize('NFC').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const recipientLines = String(docData.recipients || '').normalize('NFC').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const signerTitleLines = String(docData.signerTitle || '').normalize('NFC').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    const signerName = String(docData.signerName || '').trim();
    const qrBytes = docData.includeQr ? dataUrlToBytes(docData.qrDataUrl) : null;

    const noBorders = '<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>';

    // Bảng Tiêu ngữ 2 cột (tổng rộng A4 khả dụng = 11906 - 1701 - 850 = 9355 twips)
    const leftW = isParty ? 4677 : 3930;
    const rightW = 9355 - leftW;

    const leftCellParas = [
      govOrg ? makePara(makeRun(govOrg, { sizePt: isParty ? 13 : 12.5 }), { align: 'center', before: 0, after: 20 }) : '',
      makePara(makeRun(issueOrg, { sizePt: isParty ? 13 : 12.5, bold: true }), { align: 'center', before: 0, after: 20 }),
      makePara(makeRun(isParty ? '*' : '──────────', { sizePt: isParty ? 13 : 10, bold: isParty }), { align: 'center', before: 0, after: 40 }),
      makePara(makeRun(numberLabel, { sizePt: isParty ? 14 : 13 }), { align: 'center', before: 20, after: 20 }),
      (!docType && docSubject) ? makePara(makeRun(docSubject, { sizePt: 12 }), { align: 'center', before: 20, after: 20 }) : ''
    ].filter(Boolean).join('');

    const rightCellParas = isParty
      ? [
          makePara(makeRun('ĐẢNG CỘNG SẢN VIỆT NAM', { sizePt: 15, bold: true }), { align: 'center', before: 0, after: 40, bottomBorder: true }),
          makePara(makeRun(dateLine, { sizePt: 14, italic: true }), { align: 'center', before: 80, after: 20 })
        ].join('')
      : [
          makePara(makeRun('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM', { sizePt: 12.5, bold: true }), { align: 'center', before: 0, after: 20 }),
          makePara(makeRun('Độc lập - Tự do - Hạnh phúc', { sizePt: 13.5, bold: true }), { align: 'center', before: 0, after: 40, bottomBorder: true }),
          makePara(makeRun(dateLine, { sizePt: 13.5, italic: true }), { align: 'center', before: 80, after: 20 })
        ].join('');

    const headerTableXml =
      `<w:tbl><w:tblPr><w:tblW w:w="9355" w:type="dxa"/>${noBorders}</w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="${leftW}"/><w:gridCol w:w="${rightW}"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="${leftW}" w:type="dxa"/></w:tcPr>${leftCellParas}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="${rightW}" w:type="dxa"/></w:tcPr>${rightCellParas}</w:tc></w:tr></w:tbl>`;

    const titleParas = [
      docType ? makePara(makeRun(docType, { sizePt: isParty ? 15 : 14, bold: true }), { align: 'center', before: 200, after: 40 }) : '',
      (docType && docSubject) ? makePara(makeRun(docSubject, { sizePt: 14, bold: true }), { align: 'center', before: 0, after: 60 }) : '',
      (docType && docSubject) ? makePara(makeRun('-----', { sizePt: 14, bold: true }), { align: 'center', before: 0, after: 160 }) : ''
    ].filter(Boolean).join('');

    const bodyParas = bodyLines.map(line => {
      const isHeading = /^(I|II|III|IV|V|VI)\.\s/.test(line) || /^(QUYẾT ĐỊNH|QUYẾT NGHỊ):?$/i.test(line);
      const isCenteredHeading = /^(QUYẾT ĐỊNH|QUYẾT NGHỊ):?$/i.test(line);
      const isKinhGui = /^Kính gửi:/i.test(line);
      return makePara(
        makeRun(line, { sizePt: 14, bold: isHeading || isKinhGui }),
        {
          align: isCenteredHeading ? 'center' : 'both',
          before: isHeading ? 120 : 60,
          after: 60,
          firstLine: isCenteredHeading ? 0 : 720 // 1.27cm = 720 twips
        }
      );
    }).join('');

    // Bảng Nơi nhận & Chữ ký
    const noiNhanTitleRun = isParty
      ? makeRun('Nơi nhận', { sizePt: 14, underline: true }) + makeRun(':', { sizePt: 14 })
      : makeRun('Nơi nhận:', { sizePt: 12, bold: true, italic: true });

    const leftFooterParas = [
      makePara(noiNhanTitleRun, { align: 'left', before: 0, after: 20 }),
      ...recipientLines.map(r => makePara(makeRun(r, { sizePt: isParty ? 12 : 11 }), { align: 'left', before: 0, after: 10, line: 240 })),
      qrBytes ? makePara(makeQrDrawingXml('rIdQr1', 84), { align: 'left', before: 80, after: 0 }) : ''
    ].filter(Boolean).join('');

    const rightFooterParas = [
      ...signerTitleLines.map((s, idx) => makePara(makeRun(s.toUpperCase(), { sizePt: isParty ? 14 : 13.5, bold: !(isParty && idx === signerTitleLines.length - 1 && /^BÍ THƯ$/i.test(s)) }), { align: 'center', before: 0, after: 20 })),
      makePara(makeRun('', { sizePt: 14 }), { align: 'center', before: 520, after: 240 }),
      makePara(makeRun(signerName, { sizePt: 14, bold: true }), { align: 'center', before: 0, after: 0 })
    ].join('');

    const footerTableXml =
      `<w:tbl><w:tblPr><w:tblW w:w="9355" w:type="dxa"/>${noBorders}</w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="4677"/><w:gridCol w:w="4678"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="4677" w:type="dxa"/></w:tcPr>${leftFooterParas}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="4678" w:type="dxa"/></w:tcPr>${rightFooterParas}</w:tc></w:tr></w:tbl>`;

    const documentXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:body>${headerTableXml}${titleParas}${bodyParas}${makePara(makeRun('', { sizePt: 10 }), { before: 120, after: 0 })}${footerTableXml}` +
      `<w:sectPr>` +
      `<w:headerReference w:type="default" r:id="rIdHdr1"/>` +
      `<w:headerReference w:type="first" r:id="rIdHdr2"/>` +
      `<w:titlePg/>` +
      `<w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="567" w:footer="567" w:gutter="0"/>` +
      `</w:sectPr></w:body></w:document>`;

    const headerDefaultXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="26"/></w:rPr><w:t>2</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:hdr>`;

    const headerFirstXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>`;

    const stylesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;

    const docRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdHdr2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>` +
      (qrBytes ? `<Relationship Id="rIdQr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/qr.png"/>` : '') +
      `</Relationships>`;

    const rootRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`;

    const contentTypesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      (qrBytes ? `<Default Extension="png" ContentType="image/png"/>` : '') +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `</Types>`;

    const zipEntries = [
      { name: '[Content_Types].xml', data: contentTypesXml },
      { name: '_rels/.rels', data: rootRelsXml },
      { name: 'word/_rels/document.xml.rels', data: docRelsXml },
      { name: 'word/document.xml', data: documentXml },
      { name: 'word/styles.xml', data: stylesXml },
      { name: 'word/header1.xml', data: headerDefaultXml },
      { name: 'word/header2.xml', data: headerFirstXml }
    ];
    if (qrBytes) {
      zipEntries.push({ name: 'word/media/qr.png', data: qrBytes });
    }

    return buildZipBlob(zipEntries);
  }

  function downloadRealDocx(docData, fileNameBase) {
    const blob = generateDocxBlob(docData);
    const safeName = String(fileNameBase || 'VanBan_ChuanHoa').replace(/[^a-zA-Z0-9_\-\u00C0-\u1EF9]/g, '_');
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${safeName}.docx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    return { ok: true, fileName: `${safeName}.docx`, sizeBytes: blob.size };
  }

  global.VbtmDocxEngine = {
    generateDocxBlob,
    downloadRealDocx
  };
})(typeof window !== 'undefined' ? window : globalThis);
