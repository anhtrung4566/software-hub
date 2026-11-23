/**
 * VBTM.KnowledgeRAG — Kho tri thức & tra cứu có trích dẫn nguồn xác minh
 * - Nạp tài liệu, phân loại, tìm kiếm ngữ nghĩa theo đoạn, hiển thị đúng mã văn bản & vị trí đoạn
 * - Theo dõi trạng thái hiệu lực (cảnh báo ngay nếu văn bản đã hết hiệu lực, ví dụ 36-HD/VPTW)
 * - Phân biệt rõ: Thông tin đã xác minh từ Kho tri thức vs Nội dung AI đề xuất (chống bịa căn cứ pháp lý)
 */
(function (global) {
  'use strict';

  const BUILTIN_DOCS = [
    {
      id: 'kb_05hd_vptw',
      code: '05-HD/VPTW',
      title: 'Hướng dẫn số 05-HD/VPTW về thể thức và kỹ thuật trình bày văn bản của Đảng',
      category: 'Văn bản Đảng',
      effective: true,
      version: 'Hiện hành',
      chunks: [
        'Khổ giấy A4 (210mm x 297mm), phông chữ Times New Roman Unicode, cỡ chữ nội dung 14pt. Lề trên 20-25mm, lề dưới 20-25mm, lề trái 30-35mm, lề phải 15-20mm.',
        'Đánh số trang đặt chính giữa ở lề trên (header), từ trang 2 trở đi (không hiển thị số trang ở trang 1).',
        'Thay thế Hướng dẫn số 36-HD/VPTW ngày 03/4/2018 của Văn phòng Trung ương Đảng.'
      ]
    },
    {
      id: 'kb_nd30_2020',
      code: '30/2020/NĐ-CP',
      title: 'Nghị định số 30/2020/NĐ-CP ngày 05/3/2020 của Chính phủ về công tác văn thư',
      category: 'Văn bản hành chính',
      effective: true,
      version: 'Hiện hành',
      chunks: [
        'Áp dụng cho cơ quan, tổ chức nhà nước (HĐND, UBND phường/xã...). Thể thức gồm Quốc hiệu và Tiêu ngữ, Tên cơ quan ban hành, Số ký hiệu, Địa danh ngày tháng năm, Tên loại và trích yếu, Nội dung, Chức vụ ký, Dấu và Nơi nhận.',
        'Khổ giấy A4, phông Times New Roman cỡ 13-14pt. Lề trên 20-25mm, lề dưới 20-25mm, lề trái 30-35mm, lề phải 15-20mm. Số trang căn giữa lề trên, bắt đầu từ trang 2.'
      ]
    },
    {
      id: 'kb_36hd_expired',
      code: '36-HD/VPTW',
      title: 'Hướng dẫn số 36-HD/VPTW ngày 03/4/2018 (ĐÃ HẾT HIỆU LỰC)',
      category: 'Văn bản Đảng',
      effective: false,
      version: 'Hết hiệu lực — Đã thay thế bởi 05-HD/VPTW',
      chunks: [
        'CẢNH BÁO HIỆU LỰC: Hướng dẫn 36-HD/VPTW đã hết hiệu lực. Không sử dụng làm căn cứ ban hành văn bản Đảng mới; sử dụng Hướng dẫn số 05-HD/VPTW.'
      ]
    }
  ];

  function getAllDocs() {
    const custom = global.VBTM_SyncEngine ? (global.VBTM_SyncEngine.loadStore().knowledgeDocs || []) : [];
    return BUILTIN_DOCS.concat(custom);
  }

  function addDocument(doc) {
    if (!global.VBTM_SyncEngine) throw new Error('VBTM_SyncEngine chưa sẵn sàng.');
    const st = global.VBTM_SyncEngine.loadStore();
    const chunks = String(doc.content || '')
      .split(/\n{2,}|\r?\n/)
      .map(s => s.trim())
      .filter(s => s.length >= 10);
    const rec = {
      id: 'kb_' + Date.now(),
      code: String(doc.code || 'TÀI LIỆU NỘI BỘ').trim(),
      title: String(doc.title || 'Tài liệu chưa đặt tên').trim(),
      category: String(doc.category || 'Tài liệu cơ quan').trim(),
      effective: doc.effective !== false,
      version: String(doc.version || 'Hiện hành'),
      chunks: chunks.length ? chunks : [String(doc.content || '').trim()]
    };
    st.knowledgeDocs.unshift(rec);
    global.VBTM_SyncEngine.saveStore(st);
    return rec;
  }

  function search(query, topK) {
    topK = topK || 5;
    const tokens = String(query || '')
      .toLowerCase()
      .split(/[^a-z0-9àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ/-]+/i)
      .filter(t => t.length >= 2);
    if (!tokens.length) return [];

    const hits = [];
    for (const doc of getAllDocs()) {
      (doc.chunks || []).forEach((chunk, idx) => {
        const hay = (doc.code + ' ' + doc.title + ' ' + chunk).toLowerCase();
        let score = 0;
        for (const t of tokens) {
          if (hay.indexOf(t) !== -1) score += (doc.code.toLowerCase().indexOf(t) !== -1 ? 3 : 1);
        }
        if (score > 0) {
          hits.push({
            docId: doc.id,
            code: doc.code,
            title: doc.title,
            category: doc.category,
            effective: doc.effective,
            version: doc.version,
            paragraphIndex: idx + 1,
            excerpt: chunk,
            score,
            verifiedSource: true
          });
        }
      });
    }
    hits.sort((a, b) => b.score - a.score);
    return hits.slice(0, topK);
  }

  global.VBTM_KnowledgeRAG = {
    getAllDocs,
    addDocument,
    search
  };
})(typeof window !== 'undefined' ? window : globalThis);
