/**
 * VBTM.AICore — Runtime Office Compatibility & Requirement Set Detector
 * Tự động xác định MỨC 1 (BASIC), MỨC 2 (STANDARD), MỨC 3 (ADVANCED)
 * dựa trên API Requirement Sets thực tế của Office (không chỉ dựa vào tên phiên bản).
 */
(function (global) {
  'use strict';

  function hasReq(setName, version) {
    try {
      return !!(global.Office && Office.context && Office.context.requirements &&
                Office.context.requirements.isSetSupported(setName, version));
    } catch (e) {
      return false;
    }
  }

  function detectRuntime(expectedHost) {
    const info = {
      inOffice: false,
      host: expectedHost || 'Unknown',
      platform: 'WebPreview',
      officeVersion: 'Chưa kết nối Office (Chế độ xem trước trình duyệt)',
      tier: 'BASIC',
      tierNumber: 1,
      tierLabel: 'MỨC 1 – BASIC',
      outlookClient: null, // 'Classic Outlook' | 'New Outlook' | 'Outlook on the Web'
      capabilities: {
        selectionReadWrite: false,
        paragraphScan: false,
        paragraphNavigate: false,
        ooXmlPreserve: false,
        excelRanges: false,
        excelMultiSheet: false,
        excelCharts: false,
        pptSlideCreate: false,
        pptSelectedText: false,
        outlookCompose: false,
        outlookRead: false
      },
      warnings: []
    };

    if (!global.Office || !Office.context) {
      info.warnings.push('Đang mở ngoài Microsoft Office (chế độ xem trước/kiểm thử Web). Khi chạy trong Office, các nút chèn trực tiếp sẽ hoạt động.');
      return info;
    }

    info.inOffice = true;
    try {
      info.host = String(Office.context.host || expectedHost || 'Unknown');
      info.platform = String(Office.context.platform || 'PC');
      const diag = Office.context.diagnostics;
      if (diag) {
        info.officeVersion = `${diag.host || info.host} ${diag.version || ''} (${diag.platform || info.platform})`.trim();
      }
    } catch (e) {}

    const hostLower = info.host.toLowerCase();

    if (hostLower.includes('word')) {
      const w11 = hasReq('WordApi', '1.1');
      const w13 = hasReq('WordApi', '1.3');
      const w14 = hasReq('WordApi', '1.4');
      info.capabilities.selectionReadWrite = true; // Common API hoặc WordApi 1.1 đều hỗ trợ
      info.capabilities.paragraphScan = w11;
      info.capabilities.paragraphNavigate = w11;
      info.capabilities.ooXmlPreserve = w11;

      if (w14) {
        info.tier = 'ADVANCED'; info.tierNumber = 3; info.tierLabel = 'MỨC 3 – ADVANCED (WordApi 1.4+)';
      } else if (w13) {
        info.tier = 'STANDARD'; info.tierNumber = 2; info.tierLabel = 'MỨC 2 – STANDARD (WordApi 1.3)';
      } else if (w11) {
        info.tier = 'STANDARD'; info.tierNumber = 2; info.tierLabel = 'MỨC 2 – STANDARD (WordApi 1.1)';
      } else {
        info.tier = 'BASIC'; info.tierNumber = 1; info.tierLabel = 'MỨC 1 – BASIC (Common API / Office 2016)';
        info.warnings.push('Office hiện tại chỉ hỗ trợ Common API: chức năng quét toàn bộ từng đoạn nâng cao được chuyển sang chế độ đoạn chọn.');
      }
    } else if (hostLower.includes('excel')) {
      const e11 = hasReq('ExcelApi', '1.1');
      const e14 = hasReq('ExcelApi', '1.4');
      const e18 = hasReq('ExcelApi', '1.8');
      info.capabilities.excelRanges = e11;
      info.capabilities.excelMultiSheet = e14;
      info.capabilities.excelCharts = e14;

      if (e18) {
        info.tier = 'ADVANCED'; info.tierNumber = 3; info.tierLabel = 'MỨC 3 – ADVANCED (ExcelApi 1.8+)';
      } else if (e14) {
        info.tier = 'STANDARD'; info.tierNumber = 2; info.tierLabel = 'MỨC 2 – STANDARD (ExcelApi 1.4)';
      } else {
        info.tier = 'BASIC'; info.tierNumber = 1; info.tierLabel = 'MỨC 1 – BASIC (ExcelApi 1.1 / Common API)';
        if (!e14) info.warnings.push('Phiên bản Excel này chưa hỗ trợ tổng hợp nhiều Sheet tự động; vẫn hỗ trợ đầy đủ vùng dữ liệu đang chọn.');
      }
    } else if (hostLower.includes('powerpoint')) {
      const p12 = hasReq('PowerPointApi', '1.2');
      const p14 = hasReq('PowerPointApi', '1.4');
      info.capabilities.pptSelectedText = true;
      info.capabilities.pptSlideCreate = p12;

      if (p14) {
        info.tier = 'ADVANCED'; info.tierNumber = 3; info.tierLabel = 'MỨC 3 – ADVANCED (PowerPointApi 1.4+)';
      } else if (p12) {
        info.tier = 'STANDARD'; info.tierNumber = 2; info.tierLabel = 'MỨC 2 – STANDARD (PowerPointApi 1.2)';
      } else {
        info.tier = 'BASIC'; info.tierNumber = 1; info.tierLabel = 'MỨC 1 – BASIC (Common API / Office 2016-2019)';
        info.warnings.push('PowerPoint hiện tại dùng Common API: hỗ trợ chèn nội dung vào khung chọn & sao chép dàn ý slide chuẩn.');
      }
    } else if (hostLower.includes('outlook')) {
      const m15 = hasReq('Mailbox', '1.5');
      const m18 = hasReq('Mailbox', '1.8');
      const m112 = hasReq('Mailbox', '1.12');
      info.capabilities.outlookRead = true;
      info.capabilities.outlookCompose = m15;

      // Phân biệt Classic Outlook và New Outlook (Monarch / WebView2)
      const ua = (typeof navigator !== 'undefined' && navigator.userAgent) ? navigator.userAgent : '';
      const plat = info.platform.toLowerCase();
      if (plat.includes('officeonline')) {
        info.outlookClient = 'Outlook on the Web / New Outlook (Web Shell)';
      } else if (m112 && /Edg\//.test(ua) && info.officeVersion.includes('16.0')) {
        info.outlookClient = 'Classic Outlook (Win32 + WebView2)';
      } else if (/Edg\//.test(ua) && !info.officeVersion.includes('16.0')) {
        info.outlookClient = 'New Outlook for Windows';
      } else {
        info.outlookClient = 'Classic Outlook for Windows';
      }

      if (m112) {
        info.tier = 'ADVANCED'; info.tierNumber = 3; info.tierLabel = `MỨC 3 – ADVANCED (Mailbox 1.12+ · ${info.outlookClient})`;
      } else if (m18) {
        info.tier = 'STANDARD'; info.tierNumber = 2; info.tierLabel = `MỨC 2 – STANDARD (Mailbox 1.8 · ${info.outlookClient})`;
      } else {
        info.tier = 'BASIC'; info.tierNumber = 1; info.tierLabel = `MỨC 1 – BASIC (Mailbox 1.5 · ${info.outlookClient})`;
      }
    }

    return info;
  }

  global.VBTM_Compat = {
    hasReq,
    detectRuntime
  };
})(typeof window !== 'undefined' ? window : globalThis);
