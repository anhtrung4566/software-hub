/**
 * VietNamOCR Pro — Bộ Xuất Word (.DOCX) Chuẩn OOXML & Unicode Tiếng Việt Thống Nhất (js/ocr-docx-exporter.js)
 *
 * Khắc phục triệt để lỗi font tiếng Việt khi xuất Word:
 * 1. Chuẩn hóa chuỗi sang Unicode NFC xuyên suốt (không đổi sang TCVN3/VNI, không mất dấu, không thay đổi số liệu/tên riêng).
 * 2. Lọc ký tự điều khiển bất hợp lệ theo chuẩn XML 1.0 (ngăn Microsoft Word báo lỗi hỏng tệp / Repair).
 * 3. Kiểm tra trước khi xuất: phát hiện \uFFFD / mojibake, kiểm tra font trên máy tính và cảnh báo minh bạch.
 * 4. Tạo gói OOXML .docx đầy đủ chuẩn Microsoft Word:
 *    - [Content_Types].xml, _rels/.rels, word/_rels/document.xml.rels
 *    - word/styles.xml (w:docDefaults + Normal/Heading1/Heading2/Header/Footer/TableGrid gán đầy đủ w:ascii, w:hAnsi, w:eastAsia, w:cs và w:lang="vi-VN")
 *    - word/settings.xml (themeFontLang vi-VN)
 *    - word/header1.xml (trang 1 trống) & word/header2.xml (số trang canh giữa từ trang 2 đúng font)
 *    - word/document.xml (chuyển đổi trực tiếp cây DOM bố cục 2 cột, tiêu đề, bảng biểu, in đậm/nghiêng/gạch chân, căn lề và gán w:rFonts cho TỪNG RUN)
 */
(function (global) {
  'use strict';

  const SUPPORTED_FONTS = ['Times New Roman', 'Arial', 'Calibri', 'Aptos'];
  const DEFAULT_FONT = 'Times New Roman';
  const DEFAULT_SIZE_PT = 14;

  function normalizeVietnameseUnicode(str) {
    if (str == null) return '';
    // Chuẩn hóa về Unicode NFC và loại bỏ duy nhất các byte điều khiển bất hợp lệ của XML 1.0 (giữ nguyên \t, \n, \r)
    return String(str)
      .normalize('NFC')
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '');
  }

  function escXml(str) {
    return normalizeVietnameseUnicode(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  // Kiểm tra font có tồn tại trên máy tính người dùng (bằng kỹ thuật đo độ rộng chữ trên Canvas)
  function isFontAvailableOnSystem(fontName) {
    try {
      const c = document.createElement('canvas');
      const ctx = c.getContext('2d');
      const sample = 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM 1234567890 Độc lập - Tự do - Hạnh phúc';
      ctx.font = '32px monospace';
      const baseW = ctx.measureText(sample).width;
      ctx.font = `32px "${fontName}", monospace`;
      const testW = ctx.measureText(sample).width;
      return Math.abs(testW - baseW) > 0.5;
    } catch {
      return true;
    }
  }

  function getSelectedFontConfig() {
    const fontEl = document.getElementById('ocrWordFontSelect');
    const sizeEl = document.getElementById('ocrWordSizeSelect');
    const fontName = (fontEl && fontEl.value) ? fontEl.value : DEFAULT_FONT;
    const sizePt = (sizeEl && parseFloat(sizeEl.value)) ? parseFloat(sizeEl.value) : DEFAULT_SIZE_PT;
    return { fontName, sizePt };
  }

  // Kiểm tra Unicode, ký tự lỗi mã hóa & font trước khi xuất DOCX
  function validateBeforeDocxExport(rawText, fontName) {
    const text = normalizeVietnameseUnicode(rawText);
    const warnings = [];
    const replacementCount = (text.match(/\uFFFD/g) || []).length;
    if (replacementCount > 0) {
      warnings.push(`Phát hiện ${replacementCount} ký tự lỗi mã hóa (\uFFFD) trong văn bản gốc (được giữ nguyên để bạn tự kiểm tra).`);
    }
    const mojibakeMatches = text.match(/(?:Ã[\x80-\xBF]|Ä[\x80-\xBF]|áº[\x80-\xBF]|á»[\x80-\xBF]|Æ°)/g) || [];
    if (mojibakeMatches.length >= 2) {
      warnings.push(`Phát hiện ${mojibakeMatches.length} cụm ký tự có dấu hiệu lỗi mã hóa UTF-8 (mojibake).`);
    }
    if (!SUPPORTED_FONTS.includes(fontName)) {
      warnings.push(`Font "${fontName}" không nằm trong danh sách font Unicode tiêu chuẩn (Times New Roman, Arial, Calibri, Aptos).`);
    } else if (!isFontAvailableOnSystem(fontName)) {
      warnings.push(`Máy tính hiện tại chưa cài font "${fontName}" — Word có thể thay thế sang font Unicode dự phòng (Times New Roman / Arial).`);
    }
    return {
      ok: warnings.length === 0,
      warnings,
      normalizedText: text
    };
  }

  function buildRunPrXml(fontName, halfPt, bold, italic, underline) {
    const f = escXml(fontName);
    const sz = Math.round(halfPt || 28);
    return `<w:rPr>` +
      `<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>` +
      (bold ? '<w:b/><w:bCs/>' : '') +
      (italic ? '<w:i/><w:iCs/>' : '') +
      (underline ? '<w:u w:val="single"/>' : '') +
      `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` +
      `<w:lang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="vi-VN"/>` +
      `</w:rPr>`;
  }

  function buildTextRunXml(text, fontName, halfPt, bold, italic, underline) {
    const clean = normalizeVietnameseUnicode(text);
    if (!clean) return '';
    const rPr = buildRunPrXml(fontName, halfPt, bold, italic, underline);
    // Tách dòng \n thành <w:br/> bên trong run
    const parts = clean.split('\n');
    const inner = parts.map((pt, idx) =>
      (idx > 0 ? '<w:br/>' : '') + `<w:t xml:space="preserve">${escXml(pt)}</w:t>`
    ).join('');
    return `<w:r>${rPr}${inner}</w:r>`;
  }

  function parsePtFromStyle(styleStr, fallbackPt) {
    if (!styleStr) return fallbackPt;
    const m = String(styleStr).match(/font-size\s*:\s*(\d+(?:\.\d+)?)\s*pt/i);
    return m ? parseFloat(m[1]) : fallbackPt;
  }

  function parseAlignFromNode(el) {
    const alignAttr = (el.getAttribute && el.getAttribute('align')) || '';
    const style = (el.getAttribute && el.getAttribute('style')) || '';
    const m = style.match(/text-align\s*:\s*(left|right|center|justify)/i);
    const val = (m ? m[1] : alignAttr).toLowerCase();
    if (val === 'center') return 'center';
    if (val === 'right') return 'right';
    if (val === 'justify') return 'both';
    return 'left';
  }

  function collectRunsFromInlineNodes(node, fontName, curPt, bold, italic, underline) {
    if (!node) return '';
    if (node.nodeType === 3) { // TEXT_NODE
      const txt = node.nodeValue || '';
      return txt ? buildTextRunXml(txt, fontName, curPt * 2, bold, italic, underline) : '';
    }
    if (node.nodeType !== 1) return '';
    const tag = node.tagName.toUpperCase();
    if (tag === 'BR') return `<w:r>${buildRunPrXml(fontName, curPt * 2, bold, italic, underline)}<w:br/></w:r>`;

    const st = node.getAttribute('style') || '';
    const isB = bold || tag === 'B' || tag === 'STRONG' || /font-weight\s*:\s*(bold|[6-9]00)/i.test(st);
    const isI = italic || tag === 'I' || tag === 'EM' || /font-style\s*:\s*italic/i.test(st);
    const isU = underline || tag === 'U' || /text-decoration[^;]*underline|border-bottom\s*:\s*1pt/i.test(st);
    const nextPt = parsePtFromStyle(st, curPt);

    let out = '';
    for (let i = 0; i < node.childNodes.length; i++) {
      out += collectRunsFromInlineNodes(node.childNodes[i], fontName, nextPt, isB, isI, isU);
    }
    return out;
  }

  function buildParagraphXmlFromElement(pEl, fontName, defaultPt, inTable) {
    const st = (pEl.getAttribute && pEl.getAttribute('style')) || '';
    const jc = parseAlignFromNode(pEl);
    const pt = parsePtFromStyle(st, defaultPt);
    const isB = /font-weight\s*:\s*(bold|[6-9]00)/i.test(st);
    const isI = /font-style\s*:\s*italic/i.test(st);
    const hasIndent = !inTable && (/text-indent\s*:\s*1\.27cm/i.test(st) || (jc === 'both' && !/text-indent\s*:\s*0/i.test(st)));
    const indXml = hasIndent ? '<w:ind w:firstLine="720"/>' : '';
    const afterSpace = inTable ? '40' : '120';

    const runs = collectRunsFromInlineNodes(pEl, fontName, pt, isB, isI, false);
    const rPrP = buildRunPrXml(fontName, pt * 2, isB, isI, false);
    return `<w:p><w:pPr><w:jc w:val="${jc}"/>${indXml}<w:spacing w:after="${afterSpace}" w:line="276" w:lineRule="auto"/>${rPrP}</w:pPr>${runs}</w:p>`;
  }

  function buildTableXmlFromElement(tblEl, fontName, defaultPt) {
    const st = (tblEl.getAttribute && tblEl.getAttribute('style')) || '';
    const borderAttr = tblEl.getAttribute('border');
    const isBorderless = borderAttr === '0' || /border\s*:\s*none/i.test(st);
    const bordersXml = isBorderless
      ? `<w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders>`
      : `<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>`;

    const trNodes = tblEl.querySelectorAll('tr');
    let rowsXml = '';
    trNodes.forEach(tr => {
      const cells = tr.querySelectorAll('td, th');
      let cellsXml = '';
      const cellCount = Math.max(1, cells.length);
      cells.forEach(cell => {
        const cst = cell.getAttribute('style') || '';
        const wAttr = cell.getAttribute('width') || '';
        const pctMatch = (cst + ' ' + wAttr).match(/(\d+(?:\.\d+)?)%/);
        const pct = pctMatch ? parseFloat(pctMatch[1]) : (100 / cellCount);
        const dxa = Math.max(600, Math.round((pct / 100) * 9355)); // A4 usable width = 11906 - 1701 - 850 = 9355 dxa
        const cellPt = isBorderless ? defaultPt : Math.min(defaultPt, 13);

        let cellInner = '';
        const childPs = cell.querySelectorAll('p');
        if (childPs.length > 0) {
          childPs.forEach(p => { cellInner += buildParagraphXmlFromElement(p, fontName, cellPt, true); });
        } else {
          const isTh = cell.tagName.toUpperCase() === 'TH';
          const runs = collectRunsFromInlineNodes(cell, fontName, cellPt, isTh, false, false);
          const jc = isTh ? 'center' : parseAlignFromNode(cell);
          cellInner = `<w:p><w:pPr><w:jc w:val="${jc}"/><w:spacing w:after="40" w:line="240" w:lineRule="auto"/>${buildRunPrXml(fontName, cellPt * 2, isTh, false, false)}</w:pPr>${runs}</w:p>`;
        }
        if (!cellInner) cellInner = `<w:p><w:pPr>${buildRunPrXml(fontName, cellPt * 2, false, false, false)}</w:pPr></w:p>`;
        cellsXml += `<w:tc><w:tcPr><w:tcW w:w="${dxa}" w:type="dxa"/><w:vAlign w:val="top"/></w:tcPr>${cellInner}</w:tc>`;
      });
      rowsXml += `<w:tr>${cellsXml}</w:tr>`;
    });

    return `<w:tbl><w:tblPr><w:tblW w:w="9355" w:type="dxa"/>${bordersXml}<w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr>${rowsXml}</w:tbl>`;
  }

  function convertHtmlToWordBodyXml(htmlString, fontName, defaultPt) {
    const host = document.createElement('div');
    host.innerHTML = normalizeVietnameseUnicode(htmlString || '');
    let out = '';

    function walk(container) {
      for (let i = 0; i < container.childNodes.length; i++) {
        const n = container.childNodes[i];
        if (n.nodeType === 3) {
          const t = (n.nodeValue || '').trim();
          if (t) {
            out += `<w:p><w:pPr><w:jc w:val="both"/><w:spacing w:after="120" w:line="276" w:lineRule="auto"/>${buildRunPrXml(fontName, defaultPt * 2, false, false, false)}</w:pPr>${buildTextRunXml(t, fontName, defaultPt * 2, false, false, false)}</w:p>`;
          }
        } else if (n.nodeType === 1) {
          const tag = n.tagName.toUpperCase();
          if (tag === 'TABLE') {
            out += buildTableXmlFromElement(n, fontName, defaultPt);
          } else if (tag === 'P' || tag === 'H1' || tag === 'H2' || tag === 'H3' || tag === 'LI') {
            out += buildParagraphXmlFromElement(n, fontName, defaultPt, false);
          } else if (tag === 'DIV' || tag === 'SECTION' || tag === 'UL' || tag === 'OL') {
            walk(n);
          } else {
            const t = (n.textContent || '').trim();
            if (t) out += buildParagraphXmlFromElement(n, fontName, defaultPt, false);
          }
        }
      }
    }
    walk(host);
    return out;
  }

  function convertPlainTextToWordBodyXml(plainText, fontName, defaultPt) {
    const clean = normalizeVietnameseUnicode(plainText);
    const lines = clean.split(/\r?\n/);
    let out = '';
    for (const rawLine of lines) {
      const line = rawLine.trim();
      if (!line) continue;
      // Hỗ trợ dòng 2 cột tách bằng \t (từ module Soạn thảo NĐ 30)
      if (rawLine.includes('\t')) {
        const parts = rawLine.split('\t').map(s => s.trim());
        const left = parts[0] || '';
        const right = parts.slice(1).join(' ').trim();
        const rL = buildTextRunXml(left, fontName, defaultPt * 2, left === left.toUpperCase() && left.length > 3, false, false);
        const rR = buildTextRunXml(right, fontName, defaultPt * 2, right === right.toUpperCase() && right.length > 3, /ngày\s+\d+/i.test(right), false);
        out += `<w:tbl><w:tblPr><w:tblW w:w="9355" w:type="dxa"/><w:tblBorders><w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/><w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/></w:tblBorders></w:tblPr><w:tr>` +
          `<w:tc><w:tcPr><w:tcW w:w="4116" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr>${rL}</w:p></w:tc>` +
          `<w:tc><w:tcPr><w:tcW w:w="5239" w:type="dxa"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="40"/></w:pPr>${rR}</w:p></w:tc>` +
          `</w:tr></w:tbl>`;
        continue;
      }

      const isHeading = /^(CÔNG VĂN|QUYẾT ĐỊNH|THÔNG BÁO|TỜ TRÌNH|BÁO CÁO|KẾ HOẠCH|BIÊN BẢN|GIẤY MỜI|GIẤY GIỚI THIỆU|NGHỊ QUYẾT|CHƯƠNG TRÌNH|PHỤ LỤC|ĐƠN ĐỀ NGHỊ)$/i.test(line);
      const isSection = /^(PHẦN\s+[IVX0-9]+|CHƯƠNG\s+[IVX0-9]+|[IVXLCDM]+\.\s+)/i.test(line);
      const jc = isHeading ? 'center' : (isSection ? 'left' : 'both');
      const bold = isHeading || isSection;
      const ind = (!isHeading && !isSection) ? '<w:ind w:firstLine="720"/>' : '';
      const rPr = buildRunPrXml(fontName, defaultPt * 2, bold, false, false);
      out += `<w:p><w:pPr><w:jc w:val="${jc}"/>${ind}<w:spacing w:after="120" w:line="276" w:lineRule="auto"/>${rPr}</w:pPr>${buildTextRunXml(line, fontName, defaultPt * 2, bold, false, false)}</w:p>`;
    }
    return out;
  }

  // Tạo gói OOXML .DOCX hoàn chỉnh (tương thích 100% Microsoft Word & LibreOffice Writer)
  async function buildDocxBlob(options) {
    if (!global.JSZip) {
      throw new Error('Thư viện JSZip chưa được nạp để đóng gói tệp .DOCX.');
    }
    const cfg = getSelectedFontConfig();
    const fontName = options.fontName || cfg.fontName || DEFAULT_FONT;
    const sizePt = options.sizePt || cfg.sizePt || DEFAULT_SIZE_PT;
    const halfPt = Math.round(sizePt * 2);
    const fXml = escXml(fontName);

    let bodyInnerXml = '';
    if (options.html && String(options.html).trim()) {
      bodyInnerXml = convertHtmlToWordBodyXml(options.html, fontName, sizePt);
    } else {
      bodyInnerXml = convertPlainTextToWordBodyXml(options.plainText || '', fontName, sizePt);
    }
    if (!bodyInnerXml) {
      bodyInnerXml = `<w:p><w:pPr>${buildRunPrXml(fontName, halfPt, false, false, false)}</w:pPr></w:p>`;
    }

    const contentTypesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>` +
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
      `</Types>`;

    const rootRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`;

    const docRelsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>` +
      `<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
      `<Relationship Id="rIdHdr2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>` +
      `</Relationships>`;

    const stylesXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:docDefaults>` +
      `<w:rPrDefault><w:rPr>` +
      `<w:rFonts w:ascii="${fXml}" w:hAnsi="${fXml}" w:eastAsia="${fXml}" w:cs="${fXml}"/>` +
      `<w:sz w:val="${halfPt}"/><w:szCs w:val="${halfPt}"/>` +
      `<w:lang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="vi-VN"/>` +
      `</w:rPr></w:rPrDefault>` +
      `<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>` +
      `</w:docDefaults>` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>` +
      `<w:rPr><w:rFonts w:ascii="${fXml}" w:hAnsi="${fXml}" w:eastAsia="${fXml}" w:cs="${fXml}"/><w:sz w:val="${halfPt}"/><w:szCs w:val="${halfPt}"/><w:lang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="vi-VN"/></w:rPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>` +
      `<w:rPr><w:rFonts w:ascii="${fXml}" w:hAnsi="${fXml}" w:eastAsia="${fXml}" w:cs="${fXml}"/><w:sz w:val="26"/><w:szCs w:val="26"/><w:lang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="vi-VN"/></w:rPr></w:style>` +
      `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>` +
      `</w:styles>`;

    const settingsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:defaultTabStop w:val="720"/>` +
      `<w:themeFontLang w:val="vi-VN" w:eastAsia="vi-VN" w:bidi="vi-VN"/>` +
      `</w:settings>`;

    // Trang 1 không đánh số trang (chuẩn NĐ 30 & 05-HD/VPTW)
    const header1Xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr></w:p>` +
      `</w:hdr>`;

    // Trang 2 trở đi: số trang canh giữa lề trên, đúng font đã chọn
    const header2Xml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      `<w:p><w:pPr><w:pStyle w:val="Header"/><w:jc w:val="center"/></w:pPr>` +
      `<w:fldSimple w:instr=" PAGE "><w:r>${buildRunPrXml(fontName, 26, false, false, false)}<w:t>2</w:t></w:r></w:fldSimple>` +
      `</w:p></w:hdr>`;

    const documentXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:body>` +
      bodyInnerXml +
      `<w:sectPr>` +
      `<w:headerReference w:type="first" r:id="rIdHdr1"/>` +
      `<w:headerReference w:type="default" r:id="rIdHdr2"/>` +
      `<w:titlePg/>` +
      `<w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="567" w:footer="567" w:gutter="0"/>` +
      `</w:sectPr>` +
      `</w:body></w:document>`;

    const zip = new global.JSZip();
    zip.file('[Content_Types].xml', contentTypesXml);
    zip.file('_rels/.rels', rootRelsXml);
    zip.file('word/_rels/document.xml.rels', docRelsXml);
    zip.file('word/styles.xml', stylesXml);
    zip.file('word/settings.xml', settingsXml);
    zip.file('word/header1.xml', header1Xml);
    zip.file('word/header2.xml', header2Xml);
    zip.file('word/document.xml', documentXml);

    return await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
  }

  async function exportUnifiedDocx(options = {}) {
    const cfg = getSelectedFontConfig();
    const fontName = options.fontName || cfg.fontName;
    const sizePt = options.sizePt || cfg.sizePt;
    const checkText = options.plainText || (options.html ? options.html.replace(/<[^>]+>/g, ' ') : '');
    const validation = validateBeforeDocxExport(checkText, fontName);

    if (validation.warnings.length > 0 && global.showToast) {
      global.showToast('⚠ Lưu ý trước khi xuất DOCX: ' + validation.warnings.join(' | '));
    }

    const blob = await buildDocxBlob({
      html: options.html || '',
      plainText: options.plainText || '',
      fontName,
      sizePt
    });

    const fileName = (options.fileName || 'VanBan_VietNamOCR').replace(/\.(docx|doc)$/i, '') + '.docx';
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    if (global.showToast) {
      global.showToast(`✓ Đã xuất tệp Word chuẩn Unicode NFC (.docx): ${fileName} (Font: ${fontName} ${sizePt}pt)`);
    }
    return { fileName, fontName, sizePt, warnings: validation.warnings };
  }

  global.VietNamDocxExporter = {
    SUPPORTED_FONTS,
    normalizeVietnameseUnicode,
    isFontAvailableOnSystem,
    getSelectedFontConfig,
    validateBeforeDocxExport,
    buildDocxBlob,
    exportUnifiedDocx
  };
})(window);
