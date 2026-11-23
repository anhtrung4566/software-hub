/**
 * VBTM.SyncEngine — Hệ thống dữ liệu & đồng bộ thống nhất cho toàn hệ sinh thái VBTM Pro
 * Quản lý chung: Lịch sử xử lý, Nhiệm vụ, Từ điển bảo vệ, Kho tri thức RAG, Mẫu văn bản, Cấu hình (KHÔNG đồng bộ API Key thuần).
 */
(function (global) {
  'use strict';

  const STORE_KEY = 'vbtm_unified_store_v35';

  const DEFAULT_STORE = {
    schemaVersion: '3.5.0',
    updatedAt: new Date().toISOString(),
    preferences: {
      aiPolicy: 'free-first', // 'free-only' | 'free-first' | 'ask-before-paid' | 'local-only' | 'manual'
      warnExternalSend: true,
      defaultStandard: 'auto'
    },
    protectedTerms: [
      'Hướng dẫn số 05-HD/VPTW',
      'Nghị định 30/2020/NĐ-CP',
      'Đảng ủy phường',
      'UBND phường',
      'HĐND phường',
      'Ủy ban Mặt trận Tổ quốc',
      'Ban Thường vụ',
      'Chi bộ',
      'Văn phòng Trung ương Đảng'
    ],
    history: [],
    tasks: [],
    templates: [],
    knowledgeDocs: []
  };

  function loadStore() {
    try {
      const raw = localStorage.getItem(STORE_KEY);
      if (!raw) return JSON.parse(JSON.stringify(DEFAULT_STORE));
      const parsed = JSON.parse(raw);
      return Object.assign(JSON.parse(JSON.stringify(DEFAULT_STORE)), parsed);
    } catch (e) {
      return JSON.parse(JSON.stringify(DEFAULT_STORE));
    }
  }

  function saveStore(store) {
    store.updatedAt = new Date().toISOString();
    // Tuyệt đối loại bỏ mọi trường có thể chứa API Key khỏi snapshot đồng bộ
    delete store.apiKey;
    delete store.geminiKey;
    delete store.openaiKey;
    try { localStorage.setItem(STORE_KEY, JSON.stringify(store)); } catch (e) {}
    return store;
  }

  function addHistory(entry) {
    const st = loadStore();
    st.history.unshift({
      id: 'his_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
      ts: new Date().toISOString(),
      host: entry.host || 'Web',
      action: entry.action || 'AI Task',
      provider: entry.provider || '-',
      model: entry.model || '-',
      preview: String(entry.preview || '').slice(0, 180)
    });
    if (st.history.length > 60) st.history.length = 60;
    saveStore(st);
  }

  function getHistory(hostFilter) {
    const st = loadStore();
    return hostFilter ? st.history.filter(h => h.host === hostFilter) : st.history;
  }

  function addTask(task) {
    const st = loadStore();
    st.tasks.unshift({
      id: 'tsk_' + Date.now(),
      title: String(task.title || '').trim(),
      owner: String(task.owner || '').trim(),
      dueDate: String(task.dueDate || '').trim(),
      source: String(task.source || 'VBTM Pro'),
      status: 'Đang thực hiện',
      createdAt: new Date().toISOString()
    });
    saveStore(st);
    return st.tasks;
  }

  function getTasks() { return loadStore().tasks; }

  function addProtectedTerm(term) {
    const t = String(term || '').trim();
    if (!t) return loadStore().protectedTerms;
    const st = loadStore();
    if (!st.protectedTerms.some(x => x.toLowerCase() === t.toLowerCase())) {
      st.protectedTerms.push(t);
      saveStore(st);
    }
    return st.protectedTerms;
  }

  function getProtectedTerms() { return loadStore().protectedTerms; }

  function exportSyncSnapshot() {
    const st = loadStore();
    const safe = JSON.parse(JSON.stringify(st));
    safe.exportedAt = new Date().toISOString();
    safe.securityNotice = 'Snapshot đồng bộ VBTM Pro — KHÔNG chứa API Key bí mật.';
    return JSON.stringify(safe, null, 2);
  }

  function importSyncSnapshot(jsonStr, strategy) {
    // strategy: 'merge' (mặc định, không ghi đè âm thầm) | 'replace'
    strategy = strategy || 'merge';
    const incoming = JSON.parse(jsonStr);
    if (!incoming || typeof incoming !== 'object' || !incoming.schemaVersion) {
      throw new Error('Tệp đồng bộ không đúng định dạng VBTM Pro.');
    }
    const local = loadStore();
    if (strategy === 'replace') {
      saveStore(Object.assign(JSON.parse(JSON.stringify(DEFAULT_STORE)), incoming));
      return { status: 'replaced', conflicts: 0 };
    }
    // Merge an toàn, phát hiện xung đột dựa trên id / nội dung
    let conflicts = 0;
    const mergeById = (arrA, arrB) => {
      const map = new Map();
      arrA.forEach(x => map.set(x.id, x));
      (arrB || []).forEach(x => {
        if (map.has(x.id)) conflicts++;
        else map.set(x.id, x);
      });
      return Array.from(map.values());
    };
    local.history = mergeById(local.history, incoming.history).slice(0, 60);
    local.tasks = mergeById(local.tasks, incoming.tasks);
    local.knowledgeDocs = mergeById(local.knowledgeDocs, incoming.knowledgeDocs);
    (incoming.protectedTerms || []).forEach(t => {
      if (!local.protectedTerms.includes(t)) local.protectedTerms.push(t);
    });
    saveStore(local);
    return { status: 'merged', conflicts };
  }

  function clearAllUserData() {
    try { localStorage.removeItem(STORE_KEY); } catch (e) {}
  }

  global.VBTM_SyncEngine = {
    loadStore,
    saveStore,
    addHistory,
    getHistory,
    addTask,
    getTasks,
    addProtectedTerm,
    getProtectedTerms,
    exportSyncSnapshot,
    importSyncSnapshot,
    clearAllUserData
  };
})(typeof window !== 'undefined' ? window : globalThis);
