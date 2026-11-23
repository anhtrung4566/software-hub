/**
 * VBTM.OfficeBridge — Lớp cầu nối thống nhất giữa Microsoft Office (Word, Excel, PowerPoint, Outlook)
 * và VBTM Core / VBTM AI Core / VBTM Sync Engine.
 * Đảm bảo 4 ứng dụng Office chỉ là các điểm tích hợp của 1 sản phẩm VBTM Pro duy nhất.
 */
(function (global) {
  'use strict';

  function createBridge(hostName) {
    const compat = global.VBTM_Compat ? global.VBTM_Compat.detectRuntime(hostName) : { host: hostName, tier: 'BASIC', inOffice: false };

    return {
      host: hostName,
      runtime: compat,

      async runTask(taskType, systemPrompt, userText, opts) {
        opts = Object.assign({ host: hostName, taskType: taskType || 'general' }, opts || {});
        // Ưu tiên gọi qua Smart AI Router thống nhất (nếu đã nạp), fallback về VBTM_AICore
        let res;
        if (global.VBTM_SmartRouter && global.VBTM_SmartRouter.route) {
          res = await global.VBTM_SmartRouter.route(systemPrompt, userText, opts);
        } else if (global.VBTM_AICore && global.VBTM_AICore.generate) {
          res = await global.VBTM_AICore.generate(systemPrompt, userText, opts);
        } else {
          throw new Error('VBTM AI Core chưa được khởi tạo.');
        }

        if (global.VBTM_SyncEngine && opts.historyTitle) {
          global.VBTM_SyncEngine.addHistory({
            host: hostName,
            action: opts.historyTitle,
            provider: res.provider || 'gemini',
            model: res.model || '-',
            preview: res.text || ''
          });
        }
        return res;
      },

      getSharedProtectedTerms() {
        return global.VBTM_SyncEngine ? global.VBTM_SyncEngine.getProtectedTerms() : [];
      },

      getSharedHistory() {
        return global.VBTM_SyncEngine ? global.VBTM_SyncEngine.getHistory(hostName) : [];
      }
    };
  }

  global.VBTM_OfficeBridge = { createBridge };
})(typeof window !== 'undefined' ? window : globalThis);
