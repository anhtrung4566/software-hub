/**
 * VBTM Pro 2.2.4 - Bộ Công Cụ Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy & UBND Phường
 * Tích hợp công nghệ mới nhất:
 *  1. Kho Từ Điển 535 Cụm Từ Nhầm Lẫn Hành Chính + 7.578 Âm Tiết Tiếng Việt Chuẩn Quốc Gia (Unicode NFC)
 *  2. Bộ Phân Giải Ngữ Âm Âm Tiết (SyllableParser): Tự động bắt và sửa mọi âm tiết sai chính tả (ng/ngh, g/gh, c/k, hỏi/ngã, c/t/ch, n/ng/nh, s/x, ch/tr)
 *  3. Mammoth.js: Đọc trực tiếp file Word .docx / .doc / .txt có sẵn từ máy tính và tự động bóc tách thể thức
 *  4. QRCode ISO/IEC 18004 (Reed-Solomon): Tạo mã QR thật quét được ngay bằng Camera điện thoại & Zalo
 */

const VBTM_TEMPLATES = {
  ubnd_congvan: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '145/UBND-VP',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: '',
    docSubject: 'V/v triển khai ứng dụng chuẩn hóa thể thức văn bản và số hóa hồ sơ hành chính tại địa phương',
    docBody: `Kính gửi: Các bộ phận chuyên môn, cán bộ, công chức thuộc Ủy ban nhân dân phường.\n\nThực hiện Kế hoạch cải cách hành chính và chuyển đổi số năm 2026, Ủy ban nhân dân phường triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư.\n\n1. Các bộ phận chuyên môn, công chức thuộc Ủy ban nhân dân phường thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; bảo đảm đúng quy định về khổ giấy A4, lề trang, phông chữ Times New Roman, cỡ chữ và thể thức trình bày.\n\n2. Đẩy mạnh số hóa hồ sơ, tài liệu lưu trữ và ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm nâng cao chất lượng ban hành văn bản và bảo đảm an toàn thông tin nội bộ.\n\nỦy ban nhân dân phường đề nghị các bộ phận, cán bộ, công chức nghiêm túc triển khai thực hiện./.`,
    recipients: `- Như trên;\n- Thường trực Đảng ủy, HĐND phường (để b/c);\n- Chủ tịch, các PCT UBND phường;\n- UBMTTQ và các đoàn thể phường;\n- Các tổ dân phố;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_baocao: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '86/BC-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Kết quả thực hiện công tác cải cách hành chính, văn thư lưu trữ và chuyển đổi số trên địa bàn phường năm 2026',
    docBody: `Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai công tác chuẩn hóa văn bản và số hóa hồ sơ trên địa bàn phường như sau:\n\nI. KẾT QUẢ ĐẠT ĐƯỢC\n1. Công tác ban hành văn bản: 100% văn bản hành chính của Hội đồng nhân dân và Ủy ban nhân dân phường được chuẩn hóa đồng bộ theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ.\n\n2. Công tác số hóa hồ sơ: Đã ứng dụng phần mềm VBTM Pro 2.2.4 và VietNamOCR 3.2 Pro phục vụ rà soát lỗi chính tả tiếng Việt, tạo mã QR quản lý hồ sơ và bóc tách dữ liệu bảng biểu sang Word, Excel nhanh chóng, chính xác.\n\nII. PHƯƠNG HƯỚNG, NHIỆM VỤ TRONG THỜI GIAN TỚI\nTiếp tục duy trì nền nếp công tác văn thư, lưu trữ và nâng cao hiệu quả giải quyết thủ tục hành chính phục vụ nhân dân tại Bộ phận Tiếp nhận và Trả kết quả phường./.`,
    recipients: `- UBND Quận (để b/c);\n- Phòng Nội vụ;\n- Thường trực Đảng ủy, HĐND phường;\n- CT, các PCT UBND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_quyetdinh: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '120/QĐ-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'QUYẾT ĐỊNH',
    docSubject: 'Về việc thành lập Tổ công tác kiểm tra, rà soát công tác văn thư, lưu trữ và chuyển đổi số tại Ủy ban nhân dân phường',
    docBody: `ỦY BAN NHÂN DÂN PHƯỜNG\n\nCăn cứ Luật Tổ chức chính quyền địa phương;\n\nCăn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\n\nTheo đề nghị của Công chức Văn phòng - Thống kê phường,\n\nQUYẾT ĐỊNH:\n\nĐiều 1. Thành lập Tổ công tác kiểm tra, rà soát công tác chuẩn hóa thể thức văn bản hành chính và lập hồ sơ điện tử tại Ủy ban nhân dân phường.\n\nĐiều 2. Tổ công tác có nhiệm vụ hướng dẫn, kiểm tra việc thực hiện quy chuẩn văn bản theo Nghị định số 30/2020/NĐ-CP và báo cáo Chủ tịch Ủy ban nhân dân phường kết quả thực hiện.\n\nĐiều 3. Quyết định này có hiệu lực thi hành kể từ ngày ký. Công chức Văn phòng - Thống kê và các ông, bà có tên tại Điều 1 chịu trách nhiệm thi hành Quyết định này./.`,
    recipients: `- Như Điều 3;\n- Thường trực Đảng ủy, HĐND phường (để b/c);\n- Chủ tịch, các PCT UBND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_quyetdinh_quyche: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '125/QĐ-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'QUYẾT ĐỊNH',
    docSubject: 'Ban hành Quy chế công tác văn thư, lưu trữ và chuẩn hóa văn bản điện tử tại Ủy ban nhân dân phường',
    docBody: `ỦY BAN NHÂN DÂN PHƯỜNG\n\nCăn cứ Luật Tổ chức chính quyền địa phương;\n\nCăn cứ Luật Lưu trữ và Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\n\nTheo đề nghị của Công chức Văn phòng - Thống kê phường,\n\nQUYẾT ĐỊNH:\n\nĐiều 1. Ban hành kèm theo Quyết định này Quy chế công tác văn thư, lưu trữ và chuẩn hóa thể thức văn bản hành chính tại Ủy ban nhân dân phường.\n\nĐiều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.\n\nĐiều 3. Công chức Văn phòng - Thống kê và các bộ phận chuyên môn, cán bộ, công chức thuộc Ủy ban nhân dân phường chịu trách nhiệm thi hành Quyết định này./.`,
    recipients: `- Như Điều 3;\n- UBND Quận (để b/c);\n- Thường trực Đảng ủy, HĐND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_kehoach: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '92/KH-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'KẾ HOẠCH',
    docSubject: 'Triển khai công tác cải cách hành chính, chuyển đổi số và chuẩn hóa văn bản trên địa bàn phường năm 2026',
    docBody: `Thực hiện Kế hoạch của Ủy ban nhân dân quận về công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường ban hành Kế hoạch triển khai thực hiện cụ thể như sau:\n\nI. MỤC ĐÍCH, YÊU CẦU\n1. Mục đích: Bảo đảm 100% văn bản ban hành của Ủy ban nhân dân phường tuân thủ đúng quy định tại Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ.\n\n2. Yêu cầu: Gắn công tác chuẩn hóa văn bản với việc số hóa hồ sơ, giải quyết thủ tục hành chính nhanh gọn, chính xác.\n\nII. NỘI DUNG THỰC HIỆN\n1. Rà soát thể thức, kỹ thuật trình bày và chính tả tiếng Việt đối với toàn bộ dự thảo văn bản trước khi trình ký.\n\n2. Ứng dụng mã QR định danh văn bản và số hóa tài liệu lưu trữ tại Văn phòng Ủy ban nhân dân phường.\n\nIII. TỔ CHỨC THỰC HIỆN\nGiao Công chức Văn phòng - Thống kê chủ trì, phối hợp với các công chức chuyên môn triển khai thực hiện nghiêm túc Kế hoạch này./.`,
    recipients: `- UBND Quận (để b/c);\n- Thường trực Đảng ủy, HĐND phường;\n- Chủ tịch, các PCT UBND phường;\n- Các bộ phận chuyên môn thuộc phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_totrinh: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '64/TTr-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'TỜ TRÌNH',
    docSubject: 'Về việc đề nghị phê duyệt phương án nâng cấp hạ tầng công nghệ thông tin và số hóa hồ sơ lưu trữ tại phường',
    docBody: `Kính gửi: Ủy ban nhân dân quận.\n\nCăn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\n\nNhằm đáp ứng yêu cầu hiện đại hóa nền hành chính và số hóa hồ sơ tài liệu tại địa phương, Ủy ban nhân dân phường kính trình Ủy ban nhân dân quận xem xét, phê duyệt phương án triển khai phần mềm chuẩn hóa thể thức văn bản và số hóa tài liệu lưu trữ năm 2026.\n\nỦy ban nhân dân phường kính trình Ủy ban nhân dân quận xem xét, quyết định./.`,
    recipients: `- Như trên;\n- Phòng Nội vụ, Phòng Tài chính - Kế hoạch;\n- Thường trực Đảng ủy, HĐND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_thongbao: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '168/TB-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'THÔNG BÁO',
    docSubject: 'Kết luận của Chủ tịch Ủy ban nhân dân phường tại cuộc họp giao ban công tác cải cách hành chính tháng 11 năm 2026',
    docBody: `Ngày 10 tháng 11 năm 2026, Chủ tịch Ủy ban nhân dân phường đã chủ trì cuộc họp giao ban đánh giá kết quả công tác cải cách hành chính và văn thư lưu trữ. Sau khi nghe báo cáo và ý kiến phát biểu của các bộ phận chuyên môn, Chủ tịch Ủy ban nhân dân phường kết luận như sau:\n\n1. Biểu dương bộ phận Văn phòng - Thống kê và các công chức chuyên môn đã thực hiện tốt việc chuẩn hóa thể thức văn bản theo Nghị định số 30/2020/NĐ-CP.\n\n2. Yêu cầu các bộ phận tiếp tục kiểm soát chặt chẽ chất lượng dự thảo văn bản, sử dụng tính năng soát lỗi chính tả và đính kèm mã QR xác thực trên toàn bộ văn bản ban hành.\n\nThừa lệnh Chủ tịch Ủy ban nhân dân phường, Văn phòng thông báo để các bộ phận, cá nhân biết và thực hiện./.`,
    recipients: `- Thường trực Đảng ủy, HĐND phường (để b/c);\n- Chủ tịch, các PCT UBND phường;\n- Các công chức chuyên môn;\n- Các tổ dân phố;\n- Lưu: VT, VP.`,
    signerTitle: 'TL. CHỦ TỊCH\nVĂN PHÒNG - THỐNG KÊ',
    signerName: 'Lê Hoàng Nam'
  },
  hdnd_nghiquyet: {
    standard: 'nd30',
    governingOrg: '',
    issuingOrg: 'HỘI ĐỒNG NHÂN DÂN PHƯỜNG',
    docNumber: '18/NQ-HĐND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'NGHỊ QUYẾT',
    docSubject: 'Về nhiệm vụ phát triển kinh tế - xã hội, quốc phòng - an ninh và đẩy mạnh cải cách hành chính năm 2026',
    docBody: `HỘI ĐỒNG NHÂN DÂN PHƯỜNG\nKHÓA XII, KỲ HỌP THỨ 10\n\nCăn cứ Luật Tổ chức chính quyền địa phương;\n\nXét Tờ trình của Ủy ban nhân dân phường và Báo cáo thẩm tra của các Ban Hội đồng nhân dân phường,\n\nQUYẾT NGHỊ:\n\nĐiều 1. Tán thành Báo cáo kết quả thực hiện nhiệm vụ phát triển kinh tế - xã hội, cải cách hành chính và chuyển đổi số trên địa bàn phường.\n\nĐiều 2. Giao Ủy ban nhân dân phường tổ chức triển khai thực hiện Nghị quyết này; đẩy mạnh chuẩn hóa văn bản theo Nghị định số 30/2020/NĐ-CP và nâng cao chất lượng phục vụ nhân dân.\n\nĐiều 3. Thường trực Hội đồng nhân dân, các Ban của Hội đồng nhân dân và đại biểu Hội đồng nhân dân phường giám sát việc thực hiện Nghị quyết này./.`,
    recipients: `- Thường trực HĐND, UBND Quận (để b/c);\n- Thường trực Đảng ủy phường;\n- Thường trực HĐND, UBND, UBMTTQ phường;\n- Các đại biểu HĐND phường;\n- Lưu: VT.`,
    signerTitle: 'CHỦ TỊCH',
    signerName: 'Trần Minh Quân'
  },
  ubnd_bienban: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '09/BB-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'BIÊN BẢN',
    docSubject: 'Họp rà soát, nghiệm thu kết quả chuẩn hóa văn bản và số hóa hồ sơ hành chính năm 2026',
    docBody: `Hôm nay, vào lúc 08 giờ 00 phút ngày 11 tháng 11 năm 2026, tại Phòng họp Ủy ban nhân dân phường, chúng tôi gồm:\n\nI. THÀNH PHẦN THAM DỰ\n1. Ông Nguyễn Văn An - Chủ tịch Ủy ban nhân dân phường, Chủ trì.\n2. Ông Lê Hoàng Nam - Công chức Văn phòng - Thống kê, Thư ký.\n3. Cùng các công chức chuyên môn thuộc Ủy ban nhân dân phường.\n\nII. NỘI DUNG CUỘC HỌP\n1. Công chức Văn phòng - Thống kê báo cáo kết quả chuẩn hóa 100% văn bản theo Nghị định số 30/2020/NĐ-CP và số hóa tài liệu lưu trữ.\n2. Các thành viên dự họp thảo luận và thống nhất nghiệm thu kết quả thực hiện.\n\nCuộc họp kết thúc vào lúc 10 giờ 30 phút cùng ngày. Biên bản này được lập thành 02 bản có giá trị như nhau và đã được đọc lại cho mọi người cùng nghe./.`,
    recipients: `- Thành phần dự họp;\n- Lưu: VT, VP.`,
    signerTitle: 'CHỦ TRÌ CUỘC HỌP\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  ubnd_giaymoi: {
    standard: 'nd30',
    governingOrg: 'ỦY BAN NHÂN DÂN QUẬN',
    issuingOrg: 'ỦY BAN NHÂN DÂN PHƯỜNG',
    docNumber: '210/GM-UBND',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'GIẤY MỜI',
    docSubject: 'Họp triển khai công tác chuẩn hóa thể thức văn bản và tập huấn số hóa hồ sơ lưu trữ',
    docBody: `Kính gửi: Các đồng chí cán bộ, công chức chuyên môn và Tổ trưởng các tổ dân phố thuộc phường.\n\nỦy ban nhân dân phường tổ chức hội nghị triển khai công tác chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP và hướng dẫn sử dụng phần mềm VBTM Pro 2.2.4.\n\n1. Thành phần: Lãnh đạo Ủy ban nhân dân phường, toàn thể cán bộ, công chức chuyên môn và Tổ trưởng các tổ dân phố.\n\n2. Thời gian: 08 giờ 00 phút, ngày 15 tháng 11 năm 2026.\n\n3. Địa điểm: Hội trường Ủy ban nhân dân phường.\n\nĐề nghị các đồng chí sắp xếp công việc tham dự đầy đủ, đúng giờ./.`,
    recipients: `- Như trên;\n- Chủ tịch, các PCT UBND phường;\n- Lưu: VT, VP.`,
    signerTitle: 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH',
    signerName: 'Nguyễn Văn An'
  },
  danguy_congvan: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '58-CV/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: '',
    docSubject: 'V/v quán triệt thực hiện thể thức và kỹ thuật trình bày văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW',
    docBody: `Kính gửi: Các chi bộ trực thuộc Đảng ủy phường.\n\nThực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Ban Thường vụ Đảng ủy phường yêu cầu các chi bộ trực thuộc thực hiện nghiêm túc các nội dung sau:\n\n1. Thực hiện thống nhất phông chữ Times New Roman (bộ mã ký tự Unicode), cỡ chữ 14pt, định lề trang A4 (lề trên 20mm, lề dưới 20mm, lề trái 30mm, lề phải 15mm) đối với toàn bộ nghị quyết, báo cáo, chương trình, kế hoạch, biên bản sinh hoạt chi bộ.\n\n2. Tăng cường ứng dụng phần mềm chuẩn hóa thể thức và rà soát chính tả tiếng Việt ngoại tuyến để bảo đảm tính trang trọng, chính xác của văn kiện Đảng.\n\nĐề nghị cấp ủy các chi bộ trực thuộc nghiêm túc quán triệt và tổ chức thực hiện./.`,
    recipients: `- Như trên;\n- Văn phòng Quận ủy (để b/c);\n- Các đồng chí Đảng ủy viên;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M BAN THƯỜNG VỤ\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
  },
  danguy_baocao: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '42-BC/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'BÁO CÁO',
    docSubject: 'Tổng kết công tác xây dựng Đảng, cải cách hành chính trong Đảng và chuyển đổi số tại Đảng bộ phường năm 2026',
    docBody: `Thực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy phường báo cáo kết quả thực hiện công tác văn thư, lưu trữ và chuyển đổi số năm 2026 như sau:\n\nI. KẾT QUẢ THỰC HIỆN\n1. Toàn bộ hệ thống nghị quyết, quyết định, chương trình, kế hoạch, báo cáo và công văn của Đảng ủy phường và các chi bộ trực thuộc đều được chuẩn hóa đúng thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW.\n\n2. Việc ứng dụng bộ công cụ kiểm tra chính tả tiếng Việt và chuẩn hóa thể thức ngoại tuyến 100% giúp nâng cao chất lượng văn bản tham mưu của Văn phòng Đảng ủy, bảo đảm tuyệt đối an toàn thông tin nội bộ.\n\nII. NHIỆM VỤ TRỌNG TÂM THỜI GIAN TỚI\nTiếp tục đẩy mạnh số hóa hồ sơ đảng viên, văn kiện của Đảng ủy và hướng dẫn các chi bộ trực thuộc thực hiện đồng bộ, thống nhất./.`,
    recipients: `- Ban Thường vụ Quận ủy (để b/c);\n- Văn phòng Quận ủy;\n- Các đồng chí Đảng ủy viên;\n- HĐND, UBND, UBMTTQ phường;\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M ĐẢNG ỦY\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
  },
  danguy_nghiquyet: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '15-NQ/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'NGHỊ QUYẾT',
    docSubject: 'Về tăng cường sự lãnh đạo của Đảng đối với công tác cải cách hành chính và chuyển đổi số trên địa bàn phường',
    docBody: `Thực hiện Chương trình công tác toàn khóa của Ban Chấp hành Đảng bộ phường, sau khi thảo luận đề án của Ban Thường vụ Đảng ủy, Ban Chấp hành Đảng bộ phường quyết nghị:\n\nI. TÌNH HÌNH VÀ NGUYÊN NHÂN\nCông tác cải cách hành chính và chuyển đổi số tại Đảng bộ và chính quyền phường thời gian qua đạt nhiều kết quả quan trọng; chất lượng ban hành văn kiện, nghị quyết được chuẩn hóa đồng bộ.\n\nII. MỤC TIÊU, NHIỆM VỤ VÀ GIẢI PHÁP\n1. Phấn đấu 100% văn bản của Đảng ủy và các chi bộ trực thuộc được soạn thảo, thẩm định đúng quy định tại Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng.\n\n2. Giao Văn phòng Đảng ủy chủ trì, phối hợp với Ủy ban nhân dân phường hướng dẫn, kiểm tra việc thực hiện Nghị quyết này và định kỳ báo cáo Ban Thường vụ Đảng ủy./.`,
    recipients: `- Ban Thường vụ Quận ủy (để b/c);\n- Các đồng chí Đảng ủy viên;\n- HĐND, UBND, UBMTTQ phường;\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M ĐẢNG ỦY\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
  },
  danguy_kehoach: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '24-KH/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'KẾ HOẠCH',
    docSubject: 'Tập huấn nghiệp vụ công tác văn thư, lưu trữ và chuẩn hóa thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW',
    docBody: `Căn cứ Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Ban Thường vụ Đảng ủy phường ban hành Kế hoạch tập huấn nghiệp vụ công tác văn thư cho cấp ủy các chi bộ trực thuộc như sau:\n\nI. MỤC ĐÍCH, YÊU CẦU\nNâng cao kỹ năng soạn thảo, chuẩn hóa thể thức nghị quyết, biên bản, báo cáo của các chi bộ trực thuộc Đảng ủy phường.\n\nII. NỘI DUNG VÀ THỜI GIAN\n1. Nội dung: Hướng dẫn quy chuẩn trình bày văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW và thực hành phần mềm chuẩn hóa văn bản VBTM Pro 2.2.4.\n\n2. Thời gian: Hoàn thành trong tháng 11 năm 2026.\n\nIII. TỔ CHỨC THỰC HIỆN\nVăn phòng Đảng ủy chuẩn bị tài liệu, phần mềm và đôn đốc các chi bộ tham gia đầy đủ./.`,
    recipients: `- Văn phòng Quận ủy (để b/c);\n- Các đồng chí Đảng ủy viên;\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M BAN THƯỜNG VỤ\nPHÓ BÍ THƯ THƯỜNG TRỰC',
    signerName: 'Phạm Thị Thanh Hoa'
  },
  danguy_chuongtrinh: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '10-CTr/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'CHƯƠNG TRÌNH',
    docSubject: 'Công tác kiểm tra, giám sát của Ban Thường vụ Đảng ủy phường năm 2026',
    docBody: `Căn cứ Điều lệ Đảng Cộng sản Việt Nam và Chương trình kiểm tra, giám sát của Quận ủy, Ban Thường vụ Đảng ủy phường xây dựng Chương trình công tác kiểm tra, giám sát năm 2026 như sau:\n\nI. MỤC ĐÍCH, YÊU CẦU\nChủ động phòng ngừa vi phạm, nâng cao năng lực lãnh đạo và sức chiến đấu của các tổ chức đảng trực thuộc.\n\nII. NỘI DUNG KIỂM TRA, GIÁM SÁT\n1. Giám sát việc thực hiện Quy chế làm việc, nền nếp sinh hoạt chi bộ và việc chấp hành thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW.\n\n2. Kiểm tra việc thực hiện nhiệm vụ cải cách hành chính, chuyển đổi số gắn với trách nhiệm nêu gương của cán bộ, đảng viên.\n\nIII. TỔ CHỨC THỰC HIỆN\nỦy ban Kiểm tra Đảng ủy chủ trì tham mưu giúp Ban Thường vụ Đảng ủy tổ chức thực hiện Chương trình này./.`,
    recipients: `- Ủy ban Kiểm tra Quận ủy (để b/c);\n- Các đồng chí Đảng ủy viên;\n- UBKT Đảng ủy phường;\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M BAN THƯỜNG VỤ\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
  },
  danguy_quyetdinh: {
    standard: 'hd05',
    governingOrg: 'QUẬN ỦY',
    issuingOrg: 'ĐẢNG ỦY PHƯỜNG',
    docNumber: '35-QĐ/ĐU',
    location: 'Phường',
    docDate: '2026-11-11',
    docType: 'QUYẾT ĐỊNH',
    docSubject: 'Ban hành Quy chế quản lý, khai thác văn bản điện tử và lưu trữ hồ sơ tại Văn phòng Đảng ủy phường',
    docBody: `Căn cứ Điều lệ Đảng Cộng sản Việt Nam;\n\nCăn cứ Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng;\n\nXét đề nghị của Văn phòng Đảng ủy phường,\n\nBAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH\n\nĐiều 1. Ban hành kèm theo Quyết định này Quy chế quản lý, chuẩn hóa thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW và lưu trữ hồ sơ điện tử tại Đảng bộ phường.\n\nĐiều 2. Văn phòng Đảng ủy, các ban xây dựng Đảng và các chi bộ trực thuộc chịu trách nhiệm thi hành Quyết định này.\n\nĐiều 3. Quyết định này có hiệu lực kể từ ngày ký./.`,
    recipients: `- Văn phòng Quận ủy (để b/c);\n- Như Điều 2;\n- Các đồng chí Đảng ủy viên;\n- Lưu Văn phòng Đảng ủy.`,
    signerTitle: 'T/M BAN THƯỜNG VỤ\nBÍ THƯ',
    signerName: 'Trần Minh Quân'
  }
};

// ============================================================================
// BỘ PHÂN GIẢI NGỮ ÂM & KIỂM ĐỊNH ÂM TIẾT TIẾNG VIỆT (SYLLABLE PARSER ENGINE)
// ============================================================================
const VOWEL_GROUPS = {
  'a': 'aàáảãạ', 'ă': 'ăằắẳẵặ', 'â': 'âầấẩẫậ',
  'e': 'eèéẻẽẹ', 'ê': 'êềếểễệ',
  'i': 'iìíỉĩị',
  'o': 'oòóỏõọ', 'ô': 'ôồốổỗộ', 'ơ': 'ơờớởỡợ',
  'u': 'uùúủũụ', 'ư': 'ưừứửữự',
  'y': 'yỳýỷỹỵ'
};

const CHAR_TO_TONE = {};
const TONE_TO_CHAR = {};
Object.entries(VOWEL_GROUPS).forEach(([base, chars]) => {
  for (let i = 0; i < chars.length; i++) {
    CHAR_TO_TONE[chars[i]] = { base, tone: i };
    TONE_TO_CHAR[`${base}_${i}`] = chars[i];
  }
});

// Biểu thức biên từ Unicode chuẩn cho Tiếng Việt (Khắc phục triệt để lỗi \b của JS không nhận diện chữ có dấu!)
function escapeRegexStr(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function createUnicodePhraseRegex(phrase, flags = 'gui') {
  const escaped = escapeRegexStr(phrase.trim());
  return new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`, flags);
}

// Giữ nguyên kiểu viết hoa (chữ đầu câu hoặc IN HOA toàn bộ) khi thay thế từ đúng
function applyCasePattern(sourceStr, targetStr) {
  if (!sourceStr || !targetStr) return targetStr;
  if (sourceStr === sourceStr.toUpperCase() && sourceStr !== sourceStr.toLowerCase()) {
    return targetStr.toUpperCase();
  }
  if (sourceStr[0] === sourceStr[0].toUpperCase() && sourceStr[0] !== sourceStr[0].toLowerCase()) {
    return targetStr.charAt(0).toUpperCase() + targetStr.slice(1);
  }
  return targetStr;
}

// Kiểm tra một âm tiết có hợp lệ trong 7.578 âm tiết Tiếng Việt chuẩn hay không
const TECH_LOANWORDS = new Set([
  'times', 'new', 'roman', 'arial', 'unicode', 'offline', 'online', 'word', 'excel',
  'pdf', 'docx', 'doc', 'qr', 'ai', 'iso', 'email', 'internet', 'website', 'web',
  'app', 'software', 'pro', 'windows', 'wifi', 'usb', 'cloud', 'server', 'data',
  'file', 'scan', 'ocr', 'digital', 'smart', 'admin', 'system', 'macro', 'addin'
]);

function isValidVietnameseSyllable(word) {
  if (!window.VN_VALID_SYLLABLES) return true;
  const w = word.normalize('NFC').toLowerCase();
  if (w.length <= 1) return true;
  if (TECH_LOANWORDS.has(w)) return true;
  // Bỏ qua từ viết tắt IN HOA (UBND, HĐND, ĐU, VPTW...), từ có số hoặc ký tự đặc biệt
  if (word === word.toUpperCase() && word.length <= 8) return true;
  if (!/^\p{L}+$/u.test(w)) return true;
  return window.VN_VALID_SYLLABLES.has(w);
}

// Tự động tìm âm tiết Tiếng Việt chuẩn gần nhất cho bất kỳ âm tiết sai chính tả nào
function findClosestValidVietnameseSyllable(rawWord) {
  if (!window.VN_VALID_SYLLABLES) return null;
  const w = rawWord.normalize('NFC').toLowerCase();
  if (isValidVietnameseSyllable(rawWord)) return null;

  const candidates = [];
  const addCandidate = (cand, reason) => {
    const norm = cand.normalize('NFC').toLowerCase();
    if (window.VN_VALID_SYLLABLES.has(norm) && !candidates.some(c => c.word === norm)) {
      candidates.push({ word: norm, reason });
    }
  };

  // 1. Quy tắc chính tả ng/ngh, g/gh, c/k
  if (/^ng[ieêìíỉĩịèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('ngh' + w.slice(2), 'Quy tắc ng/ngh trước i, e, ê');
  }
  if (/^ngh[aăâoôơuưàáảãạằắẳẵặầấẩẫậòóỏõọồốổỗộờớởỡợùúủũụừứửữự]/i.test(w)) {
    addCandidate('ng' + w.slice(3), 'Quy tắc ng/ngh trước a, o, ô, ơ, u, ư');
  }
  if (/^g[eêèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('gh' + w.slice(1), 'Quy tắc g/gh trước e, ê');
  }
  if (/^gh[aăâoôơuưàáảãạằắẳẵặầấẩẫậòóỏõọồốổỗộờớởỡợùúủũụừứửữự]/i.test(w)) {
    addCandidate('g' + w.slice(2), 'Quy tắc g/gh trước a, o, ô, ơ, u, ư');
  }
  if (/^c[ieêìíỉĩịèéẻẽẹềếểễệ]/i.test(w)) {
    addCandidate('k' + w.slice(1), 'Quy tắc c/k trước i, e, ê');
  }

  // 2. Hoán vị dấu thanh (Đặc biệt là dấu hỏi <-> dấu ngã, sắc <-> nặng)
  const chars = Array.from(w);
  for (let i = 0; i < chars.length; i++) {
    const info = CHAR_TO_TONE[chars[i]];
    if (info) {
      const toneSwaps = info.tone === 3 ? [4, 2, 1, 5]
        : info.tone === 4 ? [3, 5, 2, 1]
        : info.tone === 2 ? [5, 3, 4, 1]
        : info.tone === 5 ? [2, 4, 3, 1]
        : [1, 2, 3, 4, 5];
      for (const newTone of toneSwaps) {
        const replacementChar = TONE_TO_CHAR[`${info.base}_${newTone}`];
        if (replacementChar && replacementChar !== chars[i]) {
          const copy = [...chars];
          copy[i] = replacementChar;
          addCandidate(copy.join(''), 'Sai dấu thanh tiếng Việt');
        }
      }
    }
  }

  // 3. Hoán vị âm cuối (c <-> t <-> ch, n <-> ng <-> nh)
  const codaSwaps = [
    ['c', 'ch'], ['c', 't'], ['t', 'c'], ['t', 'ch'], ['ch', 'c'], ['ch', 't'],
    ['n', 'ng'], ['n', 'nh'], ['ng', 'n'], ['ng', 'nh'], ['nh', 'n'], ['nh', 'ng']
  ];
  for (const [fromCoda, toCoda] of codaSwaps) {
    if (w.endsWith(fromCoda)) {
      addCandidate(w.slice(0, w.length - fromCoda.length) + toCoda, `Sai âm cuối -${fromCoda}/-${toCoda}`);
    }
  }

  // 4. Hoán vị âm đầu (s <-> x, ch <-> tr, d <-> gi <-> r, l <-> n)
  const onsetSwaps = [
    ['s', 'x'], ['x', 's'],
    ['ch', 'tr'], ['tr', 'ch'],
    ['gi', 'd'], ['d', 'gi'], ['r', 'd'], ['d', 'r'],
    ['l', 'n'], ['n', 'l']
  ];
  for (const [fromOnset, toOnset] of onsetSwaps) {
    if (w.startsWith(fromOnset)) {
      addCandidate(toOnset + w.slice(fromOnset.length), `Sai phụ âm đầu ${fromOnset}/${toOnset}`);
    }
  }

  return candidates.length > 0 ? candidates[0] : null;
}

let lastDetectedIssues = [];

document.addEventListener('DOMContentLoaded', () => {
  if (!document.getElementById('vbtmStudioApp')) return;
  loadTemplate('ubnd_congvan');
  bindVbtmEvents();
  setupWordFileUpload();
  updateQrCanvas();
  renderA4Preview();
  updateVbtmLiveStats();
});

function bindVbtmEvents() {
  const inputs = [
    'vbtmStandard', 'vbtmGovOrg', 'vbtmIssueOrg', 'vbtmDocNum',
    'vbtmLocation', 'vbtmDocDate', 'vbtmDocType', 'vbtmDocSubject',
    'vbtmDocBody', 'vbtmRecipients', 'vbtmSignerTitle', 'vbtmSignerName',
    'vbtmQrText', 'vbtmIncludeQr'
  ];
  inputs.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => {
        if (id === 'vbtmQrText' || id === 'vbtmDocNum') updateQrCanvas();
        renderA4Preview();
      });
      el.addEventListener('change', () => {
        updateQrCanvas();
        renderA4Preview();
      });
    }
  });
}

// ============================================================================
// ĐỌC TRỰC TIẾP FILE WORD (.DOCX / .DOC) & PDF (.PDF) TỪ MÁY TÍNH ĐỂ CHUẨN HÓA
// ============================================================================
if (window.pdfjsLib) {
  try {
    const workerCode = 'importScripts("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js");';
    const blob = new Blob([workerCode], { type: 'application/javascript' });
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  } catch (e) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  }
}

let preservedUploadedWordHtml = '';
let preservedUploadedFileName = '';

function finalizeVbtmSegment(seg) {
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

// Kiểm tra văn bản thuộc khối Đảng (Hướng dẫn số 05-HD/VPTW) hay Hành chính (Nghị định 30/2020/NĐ-CP)
function detectIsPartyDocument(text) {
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

// Chuẩn hóa bảng đầu trang (2 cột: Cơ quan ban hành & Tiêu đề Đảng / Quốc hiệu) đúng chuẩn Hướng dẫn 05-HD/VPTW và Nghị định 30/2020/NĐ-CP
function buildRegulationHeaderTableHtml(leftSegs, rightSegs, isParty) {
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
      dateLine = normalizeVietnameseDateExpressions(r);
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

// Chuẩn hóa bảng chân trang (2 cột: Nơi nhận bên trái & Chữ ký bên phải) đúng chuẩn Hướng dẫn 05-HD/VPTW và Nghị định 30/2020/NĐ-CP
function buildRegulationFooterTableHtml(leftSegs, rightSegs, isParty) {
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

  // Theo Hướng dẫn 05-HD/VPTW: Từ "Nơi nhận:" cỡ chữ 14pt, kiểu chữ đứng, có gạch chân dưới từ "Nơi nhận", sau có dấu hai chấm (:)
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

function joinVbtmPdfRunsText(itemsArr) {
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

function reconstructVbtmExactPdfPageLayout(rawItems, page, textContent, viewport, pageNumber = 1) {
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
    if (styles[it.fontName]?.fontFamily) fontDescriptor += ' ' + styles[it.fontName].fontFamily;
    try {
      if (page && page.commonObjs && page.commonObjs.has(it.fontName)) {
        const fontObj = page.commonObjs.get(it.fontName);
        if (fontObj && fontObj.name) fontDescriptor += ' ' + fontObj.name;
      }
    } catch (e) {}

    const isBold = /bold|bd|black|heavy|demi|-b\b/i.test(fontDescriptor);
    const isItalic = /italic|ital|oblique|-i\b/i.test(fontDescriptor) || (Math.abs(tx[2] || 0) > 0.12);

    runs.push({ str: it.str, x, y, w, endX: x + w, fontSize, isBold, isItalic });
  }

  if (runs.length === 0) return { pageHtml: '', pagePlainText: '' };

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
    const fullText = joinVbtmPdfRunsText(items);
    const x = items[0].x;
    const endX = items[items.length - 1].endX;
    if (x < minBodyX) minBodyX = x;
    if (endX > maxBodyX) maxBodyX = endX;

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
      leftText = joinVbtmPdfRunsText(items.slice(0, splitIdx));
      rightText = joinVbtmPdfRunsText(items.slice(splitIdx));
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

  const lines = parsedLines.filter((l, idx) => {
    if ((idx === 0 || idx === parsedLines.length - 1) && /^(?:Trang\s*)?[-–—]?\s*\d{1,3}\s*[-–—]?$/i.test(l.fullText)) {
      return false;
    }
    return true;
  });

  if (lines.length === 0) return { pageHtml: '', pagePlainText: '' };

  if (minBodyX > pageW * 0.35) minBodyX = 72;
  if (maxBodyX < pageW * 0.65) maxBodyX = pageW - 54;
  const pageMidX = (minBodyX + maxBodyX) / 2;

  const fullPageText = lines.map(l => l.plainText).join('\n');
  const isPartyDoc = detectIsPartyDocument(fullPageText);

  const htmlBlocks = [];
  let idx = 0;

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
      htmlBlocks.push(buildRegulationHeaderTableHtml(headerLeft, headerRight, isPartyDoc));
    }
  }

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
        htmlBlocks.push(isParty
          ? `<p align="center" style="text-align:center; text-indent:0; font-weight:bold; margin:0 0 10pt 0; line-height:1; font-size:14pt;">-----</p>`
          : `<p align="center" style="text-align:center; text-indent:0; font-weight:normal; margin:0 0 10pt 0; line-height:1; font-size:10pt;">──────────────</p>`);
      }
    }
  }

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
      const prevIsHeading = /^([IVXLCDM]+\.|PHẦN\s+[IVX0-9]+|CHƯƠNG\s+[IVX0-9]+|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG QUYẾT ĐỊNH)/i.test(mergedText);
      const nextStartsNewBlock = /^([-+•*]|\d+[.)]|[a-zđ]\)|[IVXLCDM]+\.|Điều\s+\d+|Khoản\s+\d+|Phần\s+[IVX0-9]+|Chương\s+[IVX0-9]+|Căn cứ\b|Xét\b|Theo đề nghị\b|Thực hiện\b|Kính gửi\b|Hôm nay\b|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN)/i.test(nxtTxt);
      const nextIsIndentedNewPara = (nxtLn.x > minBodyX + 18) && (prevLn.endX < maxBodyX - 45 || prevEndedStop);
      const normalLineSpacing = verticalGap > 0 && verticalGap <= prevLn.fontSize * 2.05;

      if (!prevEndedStop && !prevIsHeading && !nextStartsNewBlock && !nextIsIndentedNewPara && normalLineSpacing) {
        mergedText += ' ' + nxtTxt;
        prevLn = nxtLn;
        bIdx++;
      } else {
        break;
      }
    }

    htmlBlocks.push(formatBodyParagraphHtml(mergedText));
  }

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
      htmlBlocks.push(buildRegulationFooterTableHtml(footerLeft, footerRight, isPartyDoc));
    }
  }

  return {
    pageHtml: htmlBlocks.join('\n'),
    pagePlainText: lines.map(l => l.plainText).join('\n')
  };
}

// Tái tạo bố cục văn bản chuẩn Hướng dẫn 05-HD/VPTW & Nghị định 30/2020/NĐ-CP từ văn bản thuần (OCR / Text / DOCX)
function reconstructRegulationHtmlFromPlainText(rawText) {
  const isPartyDoc = detectIsPartyDocument(rawText);
  const rawLines = String(rawText || '').replace(/\r\n/g, '\n').split('\n');
  const nonEmptyLines = rawLines.map(l => l.trim()).filter(Boolean);
  if (nonEmptyLines.length === 0) return '';

  const headerLeft = [];
  const headerRight = [];
  const footerLeft = [];
  const footerRight = [];
  const bodyLines = [];

  let state = 'HEADER';
  for (let i = 0; i < nonEmptyLines.length; i++) {
    const line = nonEmptyLines[i];
    if (/^[-_*—–=]{2,}$/.test(line)) continue;

    if (state === 'HEADER') {
      const twoCol = line.match(/^(.+?)(?:\t+|\s{4,})(.+)$/);
      if (twoCol && i <= 8) {
        headerLeft.push(twoCol[1].trim());
        headerRight.push(twoCol[2].trim());
        continue;
      }
      if (i <= 8 && /^(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập\s*-\s*Tự do)/i.test(line)) {
        headerRight.push(line);
        continue;
      }
      if (i <= 8 && /\bngày\s+\d+\s+tháng\s+\d+\s+năm\s+\d+/i.test(line) && !/^Hôm nay/i.test(line)) {
        headerRight.push(line);
        continue;
      }
      if (i <= 7 && (/^(QUẬN ỦY|HUYỆN ỦY|THÀNH ỦY|TỈNH ỦY|ĐẢNG ỦY|ĐẢNG BỘ|CHI BỘ|BAN CHẤP HÀNH|BAN THƯỜNG VỤ|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN|Số\b|Số\s*:)/i.test(line))) {
        headerLeft.push(line);
        continue;
      }
      state = 'BODY';
    }

    if (state === 'BODY') {
      if (/^Nơi nhận\b/i.test(line) || /\b(T\/M\s+ĐẢNG\s+ỦY|T\/M\s+BAN\s+THƯỜNG\s+VỤ|TM\.\s+ỦY\s+BAN|KT\.\s+CHỦ\s+TỊCH|TL\.\s+CHỦ\s+TỊCH)\b/i.test(line)) {
        state = 'FOOTER';
      }
    }

    if (state === 'FOOTER') {
      const twoCol = line.match(/^(.+?)(?:\t+|\s{4,})(.+)$/);
      if (twoCol) {
        footerLeft.push(twoCol[1].trim());
        footerRight.push(twoCol[2].trim());
      } else if (/^(Nơi nhận|-|\+|•|Lưu\b)/i.test(line)) {
        footerLeft.push(line);
      } else {
        footerRight.push(line);
      }
      continue;
    }

    bodyLines.push(line);
  }

  const outHtml = [];
  if (headerLeft.length > 0 || headerRight.length > 0) {
    outHtml.push(buildRegulationHeaderTableHtml(headerLeft, headerRight, isPartyDoc));
  }

  let bIdx = 0;
  while (bIdx < bodyLines.length) {
    const l = bodyLines[bIdx];
    if (/^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|QUY ĐỊNH|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|KẾT LUẬN)$/i.test(l) && bIdx <= 3) {
      const titleSize = isPartyDoc ? '15pt' : '14pt';
      outHtml.push(`<p style="text-align:center; text-indent:0; font-weight:bold; margin:12pt 0 4pt 0; font-size:${titleSize}; text-transform:uppercase;">${escapeHtml(l)}</p>`);
      bIdx++;
      if (bIdx < bodyLines.length && !/^(Căn cứ|Thực hiện|Kính gửi|I\.|1\.|Điều\s+1)/i.test(bodyLines[bIdx])) {
        outHtml.push(`<p style="text-align:center; text-indent:0; font-weight:bold; margin:0 0 4pt 0; font-size:14pt;">${escapeHtml(bodyLines[bIdx])}</p>`);
        outHtml.push(`<hr style="width:32%; border:0; border-top:1px solid #000; margin:2pt auto 12pt auto;"/>`);
        bIdx++;
      }
      continue;
    }

    let merged = l;
    bIdx++;
    while (bIdx < bodyLines.length) {
      const nxt = bodyLines[bIdx];
      const endedStop = /[.:;!?]$|\.\s*\/\.\s*$/.test(merged);
      const startsNew = /^([-+•*]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Khoản\s+\d+|Căn cứ\s+|Xét\s+|Theo đề nghị|Kính gửi|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN)/i.test(nxt);
      if (!endedStop && !startsNew) {
        merged += ' ' + nxt;
        bIdx++;
      } else {
        break;
      }
    }
    outHtml.push(formatBodyParagraphHtml(merged));
  }

  if (footerLeft.length > 0 || footerRight.length > 0) {
    outHtml.push(buildRegulationFooterTableHtml(footerLeft, footerRight, isPartyDoc));
  }

  return outHtml.join('\n');
}

async function extractTextFromPdfInVbtm(file) {
  if (!window.pdfjsLib) {
    throw new Error('Thư viện đọc PDF đang tải, vui lòng thử lại.');
  }
  const arrayBuffer = await file.arrayBuffer();
  let pdfDoc;
  try {
    pdfDoc = await window.pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer),
      cMapUrl: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/cmaps/',
      cMapPacked: true
    }).promise;
  } catch (e) {
    pdfDoc = await window.pdfjsLib.getDocument({
      data: new Uint8Array(await file.arrayBuffer()),
      disableWorker: true
    }).promise;
  }

  const allPagesHtml = [];
  const allPagesText = [];
  for (let p = 1; p <= pdfDoc.numPages; p++) {
    const page = await pdfDoc.getPage(p);
    const viewport1x = page.getViewport({ scale: 1.0 });
    const viewportRender = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement('canvas');
    canvas.width = viewportRender.width;
    canvas.height = viewportRender.height;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: viewportRender }).promise;

    const textContent = await page.getTextContent();
    const rawItems = (textContent?.items || []).filter(it => typeof it.str === 'string' && it.str.length > 0);
    if (rawItems.filter(it => it.str.trim().length > 0).length >= 2) {
      const { pageHtml, pagePlainText } = reconstructVbtmExactPdfPageLayout(rawItems, page, textContent, viewport1x, p);
      allPagesHtml.push(pageHtml);
      allPagesText.push(pagePlainText);
    } else {
      const b64 = canvas.toDataURL('image/jpeg', 0.92);
      const fd = new FormData();
      fd.append('apikey', 'helloworld');
      fd.append('base64Image', b64);
      fd.append('language', 'auto');
      fd.append('OCREngine', '2');
      fd.append('scale', 'true');
      const resp = await fetch('https://api.ocr.space/parse/image', { method: 'POST', body: fd });
      const json = await resp.json();
      if (json?.ParsedResults?.[0]?.ParsedText) {
        const ocrTxt = json.ParsedResults[0].ParsedText.trim();
        allPagesText.push(ocrTxt);
        allPagesHtml.push(reconstructRegulationHtmlFromPlainText(ocrTxt));
      }
    }
  }
  preservedUploadedWordHtml = allPagesHtml.map((h, i) => i === 0 ? h : `<div style="page-break-before:always; margin-top:16pt;">${h}</div>`).join('\n');
  return allPagesText.join('\n\n');
}

function setupWordFileUpload() {
  const fileInput = document.getElementById('vbtmWordFileInput');
  if (!fileInput) return;

  const processFile = async (file) => {
    if (!file) return;
    try {
      const ext = (file.name.split('.').pop() || '').toLowerCase();
      const isPdf = ext === 'pdf' || file.type === 'application/pdf';
      let rawText = '';

      preservedUploadedWordHtml = '';
      preservedUploadedFileName = file.name.replace(/\.[^.]+$/, '');

      if (isPdf) {
        if (window.showToast) window.showToast(`Đang đọc tệp PDF "${file.name}" và chuẩn hóa thể thức (Hướng dẫn 05-HD/VPTW & Nghị định 30)...`);
        rawText = await extractTextFromPdfInVbtm(file);
        parseUploadedDocumentAndStandardize(rawText, file.name, true);
        const previewEl = document.getElementById('a4PreviewPaper');
        if (previewEl && preservedUploadedWordHtml) {
          previewEl.innerHTML = preservedUploadedWordHtml;
        }
        if (window.showToast) {
          window.showToast(`Đã tải tệp PDF "${file.name}" chuẩn thể thức Hướng dẫn 05-HD/VPTW & Nghị định 30! Bấm "Xuất File Word" để tải về.`);
        }
        return;
      } else if (ext === 'docx' && window.mammoth) {
        const arrayBuffer = await file.arrayBuffer();
        const textResult = await window.mammoth.extractRawText({ arrayBuffer });
        rawText = textResult.value || '';
      } else {
        const textContent = await file.text();
        if (/<html|<body|<p|<div|<w:/i.test(textContent)) {
          const tmp = document.createElement('div');
          tmp.innerHTML = textContent.replace(/<\/p>|<\/div>|<\/tr>|<br\s*\/?>/gi, '\n');
          rawText = tmp.innerText || tmp.textContent || '';
        } else {
          rawText = textContent;
        }
      }

      if (!rawText.trim()) {
        if (window.showToast) window.showToast('Tệp văn bản trống hoặc không đọc được nội dung.');
        return;
      }

      parseUploadedDocumentAndStandardize(rawText, file.name, false);
    } catch (err) {
      console.error('Lỗi đọc tệp:', err);
      if (window.showToast) window.showToast('Lỗi khi đọc tệp: ' + err.message);
    }
  };

  fileInput.addEventListener('change', async (e) => {
    const file = e.target.files?.[0];
    fileInput.value = '';
    await processFile(file);
  });

  // Hỗ trợ kéo thả trực tiếp file .docx, .doc, .pdf, .txt vào khung VBTM Pro Studio
  const studioApp = document.getElementById('vbtmStudioApp');
  if (studioApp) {
    studioApp.addEventListener('dragover', (e) => e.preventDefault());
    studioApp.addEventListener('drop', async (e) => {
      e.preventDefault();
      if (e.dataTransfer?.files?.[0]) {
        await processFile(e.dataTransfer.files[0]);
      }
    });
  }
}

// Phân tích thông minh văn bản tải lên và tự động bóc tách đầy đủ các trường thể thức (Hỗ trợ cả Hướng dẫn 05-HD/VPTW và Nghị định 30/2020/NĐ-CP)
function parseUploadedDocumentAndStandardize(rawText, fileName, keepPreservedHtml = false) {
  const isParty = detectIsPartyDocument(rawText);
  document.getElementById('vbtmStandard').value = isParty ? 'hd05' : 'nd30';

  const lines = rawText
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  if (isParty) {
    document.getElementById('vbtmGovOrg').value = 'QUẬN ỦY';
    document.getElementById('vbtmIssueOrg').value = 'ĐẢNG ỦY PHƯỜNG';
  } else {
    document.getElementById('vbtmGovOrg').value = 'ỦY BAN NHÂN DÂN QUẬN';
    document.getElementById('vbtmIssueOrg').value = 'ỦY BAN NHÂN DÂN PHƯỜNG';
  }

  // Tìm Số ký hiệu (Hỗ trợ cả "Số: 145/UBND-VP" và "Số 42-BC/ĐU" không có dấu hai chấm của Đảng)
  const numMatch = rawText.match(/\bSố\s*:?\s*([0-9A-ZĐa-z/-]+(?:\/[A-ZĐa-z0-9-]+)?)/i);
  if (numMatch && numMatch[1]) {
    document.getElementById('vbtmDocNum').value = numMatch[1].trim();
  }

  // Tìm Tên loại văn bản và Trích yếu
  const docTypes = ['BÁO CÁO', 'QUYẾT ĐỊNH', 'KẾ HOẠCH', 'TỜ TRÌNH', 'THÔNG BÁO', 'NGHỊ QUYẾT', 'CHƯƠNG TRÌNH', 'CHỈ THỊ', 'QUY CHẾ', 'QUY ĐỊNH', 'BIÊN BẢN', 'GIẤY MỜI', 'HƯỚNG DẪN', 'KẾT LUẬN'];
  let detectedType = '';
  let detectedSubject = '';
  let typeLineIdx = -1;

  for (let i = 0; i < Math.min(lines.length, 15); i++) {
    const u = lines[i].toUpperCase();
    for (const t of docTypes) {
      if (u === t || u.startsWith(t + ' ')) {
        detectedType = t;
        typeLineIdx = i;
        break;
      }
    }
    if (detectedType) break;
  }

  if (typeLineIdx !== -1 && lines[typeLineIdx + 1]) {
    detectedSubject = lines[typeLineIdx + 1];
  } else {
    const vvIdx = lines.findIndex(l => /^(V\/v|Về việc)\b/i.test(l));
    if (vvIdx !== -1) {
      detectedSubject = lines[vvIdx];
      typeLineIdx = vvIdx - 1;
    }
  }

  document.getElementById('vbtmDocType').value = detectedType;
  document.getElementById('vbtmDocSubject').value = detectedSubject;

  // Bóc tách Nơi nhận & Người ký ở cuối văn bản để không bị lặp trong phần Thân bài
  let footerStartIdx = lines.findIndex((l, idx) => idx > 3 && (/^Nơi nhận\b/i.test(l) || /\b(T\/M\s+ĐẢNG\s+ỦY|T\/M\s+BAN\s+THƯỜNG\s+VỤ|TM\.\s+ỦY\s+BAN|KT\.\s+CHỦ\s+TỊCH|TL\.\s+CHỦ\s+TỊCH)\b/i.test(l)));
  if (footerStartIdx !== -1) {
    const footerLines = lines.slice(footerStartIdx);
    const recLines = [];
    const signLines = [];
    footerLines.forEach(fl => {
      const parts = fl.split(/\t+|\s{4,}/);
      parts.forEach(p => {
        const pt = p.trim();
        if (!pt || /^Nơi nhận\s*:?$/i.test(pt)) return;
        if (/^(-|\+|•|Như\s+|Lưu\b)/i.test(pt)) {
          recLines.push(pt.startsWith('-') ? pt : `- ${pt}`);
        } else {
          signLines.push(pt);
        }
      });
    });
    if (recLines.length > 0) {
      document.getElementById('vbtmRecipients').value = recLines.join('\n');
    }
    if (signLines.length >= 2) {
      document.getElementById('vbtmSignerName').value = signLines[signLines.length - 1];
      document.getElementById('vbtmSignerTitle').value = signLines.slice(0, signLines.length - 1).join('\n');
    }
  }

  const bodyEndIdx = footerStartIdx !== -1 ? footerStartIdx : lines.length;
  const filteredBodyLines = lines.slice(0, bodyEndIdx).filter((l, idx) => {
    const u = l.toUpperCase();
    if (/^[-_*—–=]+$/.test(l)) return false;
    if (u.includes('CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM')) return false;
    if (u.includes('ĐỘC LẬP - TỰ DO - HẠNH PHÚC')) return false;
    if (u.includes('ĐẢNG CỘNG SẢN VIỆT NAM')) return false;
    if (idx <= 5 && (/^(QUẬN ỦY|ĐẢNG ỦY|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN|Số\b|Số\s*:)/i.test(l) || /\bngày\s+\d+\s+tháng/i.test(l))) return false;
    if (idx === typeLineIdx || (detectedSubject && l === detectedSubject)) return false;
    return true;
  });

  // Nối các dòng bị ngắt trong cùng 1 đoạn văn
  const mergedParagraphs = [];
  let pIdx = 0;
  while (pIdx < filteredBodyLines.length) {
    let curP = filteredBodyLines[pIdx];
    pIdx++;
    while (pIdx < filteredBodyLines.length) {
      const nxt = filteredBodyLines[pIdx];
      const endedStop = /[.:;!?]$|\.\s*\/\.\s*$/.test(curP);
      const startsNew = /^([-+•*]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Khoản\s+\d+|Căn cứ\s+|Xét\s+|Theo đề nghị|Kính gửi|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ)/i.test(nxt);
      if (!endedStop && !startsNew) {
        curP += ' ' + nxt;
        pIdx++;
      } else {
        break;
      }
    }
    mergedParagraphs.push(curP);
  }

  document.getElementById('vbtmDocBody').value = mergedParagraphs.join('\n\n');
  if (!keepPreservedHtml) {
    standardizeAndFixEverything();
  } else {
    runSpellAndFormatCheck();
  }

  if (window.showToast) {
    window.showToast(`Đã nạp tệp "${fileName}" và chuẩn hóa thể thức theo ${isParty ? 'Hướng dẫn số 05-HD/VPTW' : 'Nghị định 30/2020/NĐ-CP'}!`);
  }
}

function loadTemplate(key) {
  preservedUploadedWordHtml = '';
  const tpl = VBTM_TEMPLATES[key];
  if (!tpl) return;
  const sel = document.getElementById('vbtmTemplateSelect');
  if (sel && sel.value !== key) sel.value = key;
  document.getElementById('vbtmStandard').value = tpl.standard;
  document.getElementById('vbtmGovOrg').value = tpl.governingOrg;
  document.getElementById('vbtmIssueOrg').value = tpl.issuingOrg;
  document.getElementById('vbtmDocNum').value = tpl.docNumber;
  document.getElementById('vbtmLocation').value = tpl.location;
  document.getElementById('vbtmDocDate').value = tpl.docDate;
  document.getElementById('vbtmDocType').value = tpl.docType;
  document.getElementById('vbtmDocSubject').value = tpl.docSubject;
  document.getElementById('vbtmDocBody').value = tpl.docBody;
  document.getElementById('vbtmRecipients').value = tpl.recipients;
  document.getElementById('vbtmSignerTitle').value = tpl.signerTitle;
  document.getElementById('vbtmSignerName').value = tpl.signerName;
  document.getElementById('vbtmQrText').value = `VBTM-VERIFY:${tpl.docNumber}|${tpl.issuingOrg}|${tpl.docDate}`;
  updateQrCanvas();
  renderA4Preview();
  runSpellAndFormatCheck();
  updateVbtmLiveStats();
}

function formatVietnameseDate(dateStr, location) {
  if (!dateStr) return `${location || 'Phường'}, ngày 11 tháng 11 năm 2026`;
  const parts = dateStr.split('-');
  if (parts.length !== 3) return `${location}, ${dateStr}`;
  const year = parts[0];
  const monthNum = parseInt(parts[1], 10);
  const dayNum = parseInt(parts[2], 10);
  const dayStr = dayNum < 10 ? `0${dayNum}` : `${dayNum}`;
  const monthStr = (monthNum === 1 || monthNum === 2) ? `0${monthNum}` : `${monthNum}`;
  return `${location || 'Phường'}, ngày ${dayStr} tháng ${monthStr} năm ${year}`;
}

// Chuẩn hóa cách ghi ngày tháng trong văn bản theo Phụ lục I Nghị định 30/2020/NĐ-CP:
// Ngày nhỏ hơn 10 và tháng 1, 2 phải thêm số 0 phía trước; tháng 3 đến tháng 12 không có số 0
function normalizeVietnameseDateExpressions(text) {
  return String(text || '')
    .replace(/\bngày\s+([1-9])(?=\s+tháng\b)/gi, 'ngày 0$1')
    .replace(/\btháng\s+([12])(?=\s+năm\b)/gi, 'tháng 0$1')
    .replace(/\btháng\s+0([3-9])(?=\s+năm\b)/gi, 'tháng $1');
}

// Hàm sửa toàn bộ lỗi chính tả (535 cụm từ nhầm lẫn + 7.578 âm tiết Tiếng Việt chuẩn) trên một chuỗi văn bản
function autoFixVietnameseSpellingInString(inputStr) {
  let str = String(inputStr || '').normalize('NFC');
  let fixedCount = 0;

  // Bước 1: Quét 535 cụm từ nhầm lẫn hành chính & lỗi gõ Telex/VNI bằng Unicode Regex
  const rules = window.VN_SPELL_RULES || [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (!rule.w || !rule.r) continue;
    const regex = createUnicodePhraseRegex(rule.w, 'gui');
    str = str.replace(regex, (matched) => {
      fixedCount++;
      return applyCasePattern(matched, rule.r);
    });
  }

  // Bước 2: Quét từng âm tiết Tiếng Việt đơn lẻ với kho 7.578 âm tiết chuẩn Quốc gia
  str = str.replace(/\p{L}+/gu, (word) => {
    if (isValidVietnameseSyllable(word)) return word;
    const suggestion = findClosestValidVietnameseSyllable(word);
    if (suggestion && suggestion.word) {
      fixedCount++;
      return applyCasePattern(word, suggestion.word);
    }
    return word;
  });

  return { text: str, fixedCount };
}

// Chuẩn hóa toàn diện: vừa chuẩn hóa thể thức NĐ30/Đảng ủy vừa tự động sửa toàn bộ lỗi chính tả AI
function standardizeAndFixEverything() {
  const bodyEl = document.getElementById('vbtmDocBody');
  const subjEl = document.getElementById('vbtmDocSubject');
  if (!bodyEl) return;

  // 1. Tự động sửa toàn bộ lỗi chính tả tiếng Việt trên cả Nội dung & Trích yếu
  const fixedBodyObj = autoFixVietnameseSpellingInString(bodyEl.value);
  let text = fixedBodyObj.text;
  let fixedSpells = fixedBodyObj.fixedCount;

  if (subjEl && subjEl.value) {
    const fixedSubjObj = autoFixVietnameseSpellingInString(subjEl.value);
    subjEl.value = fixedSubjObj.text.trim().replace(/\s+/g, ' ').replace(/\.+$/, '');
    fixedSpells += fixedSubjObj.fixedCount;
  }

  // 2. Xóa khoảng trắng thừa, chuẩn hóa dấu câu & cách viết ngày tháng chuẩn Nghị định 30/2020/NĐ-CP
  text = normalizeVietnameseDateExpressions(text)
    .replace(/[ \t]+/g, ' ')
    .replace(/\s+([,.;:!?])/g, '$1')
    .replace(/([,.;:!?])([^\s\d/.)])/g, '$1 $2')
    .replace(/\(\s+/g, '(')
    .replace(/\s+\)/g, ')')
    .replace(/“\s+/g, '“')
    .replace(/\s+”/g, '”')
    .replace(/\n{3,}/g, '\n\n');

  // 3. Viết hoa chữ cái đầu mỗi đoạn và thêm dấu kết thúc "./."
  const lines = text.split('\n').map(line => {
    let trimmed = line.trim();
    if (!trimmed) return '';
    trimmed = trimmed.charAt(0).toUpperCase() + trimmed.slice(1);
    return trimmed;
  }).filter((l, idx, arr) => !(l === '' && arr[idx - 1] === ''));

  let result = lines.join('\n').trim();
  if (result && !result.endsWith('./.')) {
    result = result.replace(/\.+$/, '') + './.';
  }
  bodyEl.value = result;

  // 4. Chuẩn hóa tự động theo Hướng dẫn 05-HD/VPTW (Khối Đảng) hoặc Nghị định 30/2020/NĐ-CP (Khối Hành chính)
  const isPartyStd = document.getElementById('vbtmStandard')?.value === 'hd05';
  const docNumEl = document.getElementById('vbtmDocNum');
  if (docNumEl && docNumEl.value) {
    docNumEl.value = docNumEl.value.trim().replace(/^Số\s*:?\s*/i, '');
  }

  const recEl = document.getElementById('vbtmRecipients');
  if (recEl) {
    let recLines = recEl.value
      .split('\n')
      .map(l => l.trim())
      .filter(Boolean)
      .map(l => (l.startsWith('-') ? l : `- ${l}`));
    if (!recLines.some(l => /Lưu/i.test(l))) {
      recLines.push(isPartyStd ? '- Lưu Văn phòng Đảng ủy.' : '- Lưu: VT, VP.');
    }
    if (recLines.length > 0) {
      recLines = recLines.map((l, idx) => {
        const clean = l.replace(/[;.]+\s*$/, '');
        return idx < recLines.length - 1 ? `${clean};` : `${clean}.`;
      });
      recEl.value = recLines.join('\n');
    }
  }

  updateQrCanvas();
  renderA4Preview();
  runSpellAndFormatCheck();
  updateVbtmLiveStats();

  if (window.showToast) {
    window.showToast(`✓ Đã chuẩn hóa tự động ${isPartyStd ? 'Hướng dẫn 05-HD/VPTW' : 'Nghị định 30/2020/NĐ-CP'} & sửa ${fixedSpells} lỗi chính tả!`);
  }
}

function standardizeTextContent() {
  standardizeAndFixEverything();
}

// Kiểm tra lỗi chính tả & thẩm định thể thức AI (535 cụm từ + 7.578 âm tiết Tiếng Việt chuẩn)
function runSpellAndFormatCheck() {
  const body = (document.getElementById('vbtmDocBody')?.value || '').normalize('NFC');
  const subject = (document.getElementById('vbtmDocSubject')?.value || '').normalize('NFC');
  const docNum = document.getElementById('vbtmDocNum')?.value || '';
  const recipients = document.getElementById('vbtmRecipients')?.value || '';
  const issueContainer = document.getElementById('vbtmIssueList');

  lastDetectedIssues = [];

  // 1. Kiểm tra theo kho 535 cụm từ nhầm lẫn hành chính & lỗi gõ phím
  const rules = window.VN_SPELL_RULES || [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (!rule.w || !rule.r) continue;
    const checkRegex = createUnicodePhraseRegex(rule.w, 'ui');
    if (checkRegex.test(body) || checkRegex.test(subject)) {
      const replaceRegex = createUnicodePhraseRegex(rule.w, 'gui');
      lastDetectedIssues.push({
        type: 'Chính tả AI',
        desc: `Phát hiện "${rule.w}" → Chuẩn là: "${rule.r}" (${rule.d})`,
        fix: () => {
          const bEl = document.getElementById('vbtmDocBody');
          const sEl = document.getElementById('vbtmDocSubject');
          if (bEl) bEl.value = bEl.value.normalize('NFC').replace(replaceRegex, m => applyCasePattern(m, rule.r));
          if (sEl) sEl.value = sEl.value.normalize('NFC').replace(replaceRegex, m => applyCasePattern(m, rule.r));
        }
      });
    }
  }

  // 2. Kiểm tra từng âm tiết đơn lẻ với kho 7.578 âm tiết Tiếng Việt chuẩn
  const combinedText = `${subject} ${body}`;
  const words = combinedText.match(/\p{L}+/gu) || [];
  const checkedInvalidWords = new Set();

  words.forEach(word => {
    const lower = word.toLowerCase();
    if (checkedInvalidWords.has(lower)) return;
    if (!isValidVietnameseSyllable(word)) {
      checkedInvalidWords.add(lower);
      const suggestion = findClosestValidVietnameseSyllable(word);
      const sugWord = suggestion ? suggestion.word : '';
      const sugReason = suggestion ? suggestion.reason : 'Âm tiết không tồn tại trong Tiếng Việt chuẩn';
      const wordRegex = createUnicodePhraseRegex(lower, 'gui');

      lastDetectedIssues.push({
        type: 'Ngữ âm AI',
        desc: sugWord
          ? `Âm tiết sai "${word}" (${sugReason}) → Gợi ý sửa thành: "${applyCasePattern(word, sugWord)}"`
          : `Phát hiện âm tiết sai cấu trúc Tiếng Việt: "${word}"`,
        fix: () => {
          if (!sugWord) return;
          const bEl = document.getElementById('vbtmDocBody');
          const sEl = document.getElementById('vbtmDocSubject');
          if (bEl) bEl.value = bEl.value.normalize('NFC').replace(wordRegex, m => applyCasePattern(m, sugWord));
          if (sEl) sEl.value = sEl.value.normalize('NFC').replace(wordRegex, m => applyCasePattern(m, sugWord));
        }
      });
    }
  });

  // 3. Kiểm tra quy tắc dấu câu & thể thức Nghị định 30/2020/NĐ-CP
  if (/\s+[,.;:!?]/.test(body)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)',
      fix: () => standardizeAndFixEverything()
    });
  }

  if (/\.$/.test(subject.trim())) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Trích yếu văn bản không được đặt dấu chấm ở cuối dòng',
      fix: () => {
        document.getElementById('vbtmDocSubject').value = subject.trim().replace(/\.+$/, '');
      }
    });
  }

  if (body.trim() && !body.trim().endsWith('./.')) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Cuối nội dung văn bản hành chính thiếu ký hiệu kết thúc "./."',
      fix: () => standardizeAndFixEverything()
    });
  }

  if (!docNum.trim()) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Thiếu Số và ký hiệu văn bản (Ví dụ: 145/UBND-VP hoặc 42-BC/ĐU)',
      fix: () => { document.getElementById('vbtmDocNum').value = '01/UBND-VP'; }
    });
  }
  if (!recipients.trim() || !/Lưu/i.test(recipients)) {
    lastDetectedIssues.push({
      type: 'Thể thức NĐ30',
      desc: 'Mục "Nơi nhận" cần có dòng "- Lưu: VT,..." hoặc "- Lưu Văn phòng Đảng ủy."',
      fix: () => {
        const cur = document.getElementById('vbtmRecipients').value.trim();
        document.getElementById('vbtmRecipients').value = (cur ? cur + '\n' : '') + '- Lưu: VT.';
      }
    });
  }

  if (!issueContainer) return;

  if (lastDetectedIssues.length === 0) {
    issueContainer.innerHTML = `
      <div class="issue-item ok">
        <span>✓ Văn bản đã đạt chuẩn 100% về chính tả Tiếng Việt (7.578 âm tiết &amp; 535 quy tắc) và thể thức Đảng ủy / UBND phường!</span>
      </div>
    `;
  } else {
    issueContainer.innerHTML = lastDetectedIssues.map((item, idx) => `
      <div class="issue-item">
        <div>
          <strong style="color: #fbbf24;">[${item.type}]</strong> ${escapeHtml(item.desc)}
        </div>
        <button type="button" class="btn-tool primary" onclick="fixSingleIssue(${idx})" style="padding: 4px 10px; font-size: 0.78rem; flex-shrink: 0;">Sửa ngay</button>
      </div>
    `).join('');
  }
}

function fixSingleIssue(idx) {
  if (lastDetectedIssues[idx] && typeof lastDetectedIssues[idx].fix === 'function') {
    lastDetectedIssues[idx].fix();
    renderA4Preview();
    runSpellAndFormatCheck();
  }
}

function fixAllIssues() {
  standardizeAndFixEverything();
}

function insertSampleTypos() {
  const bodyEl = document.getElementById('vbtmDocBody');
  bodyEl.value = `Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý ngiêm túc các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra`;
  renderA4Preview();
  runSpellAndFormatCheck();
  if (window.showToast) {
    window.showToast(`Đã phát hiện ${lastDetectedIssues.length} lỗi chính tả & thể thức. Bấm "⚡ Chuẩn Hóa Thể Thức & Tự Động Sửa Lỗi" để sửa sạch 100%!`);
  }
}

// ============================================================================
// TẠO MÃ QR CHUẨN QUỐC TẾ ISO/IEC 18004 (QUÉT ĐƯỢC BẰNG ĐIỆN THOẠI & ZALO)
// ============================================================================
function updateQrCanvas() {
  const canvas = document.getElementById('vbtmQrCanvas');
  if (!canvas) return;
  const rawText = (document.getElementById('vbtmQrText')?.value || document.getElementById('vbtmDocNum')?.value || 'VBTM-PRO').trim();
  const ctx = canvas.getContext('2d');

  if (typeof window.qrcode === 'function') {
    try {
      const utf8Text = unescape(encodeURIComponent(rawText));
      const qr = window.qrcode(0, 'M');
      qr.addData(utf8Text);
      qr.make();

      const count = qr.getModuleCount();
      const cellSize = 4;
      const margin = 12;
      const totalSize = count * cellSize + margin * 2;
      canvas.width = totalSize;
      canvas.height = totalSize;

      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, totalSize, totalSize);
      ctx.fillStyle = '#000000';

      for (let r = 0; r < count; r++) {
        for (let c = 0; c < count; c++) {
          if (qr.isDark(r, c)) {
            ctx.fillRect(margin + c * cellSize, margin + r * cellSize, cellSize, cellSize);
          }
        }
      }
      return;
    } catch (e) {
      console.warn('QR fallback:', e);
    }
  }
}

function formatBodyParagraphHtml(rawLine) {
  const line = rawLine.trim();
  const escaped = escapeHtml(line);
  if (/^(QUYẾT ĐỊNH|QUYẾT NGHỊ|NGHỊ QUYẾT|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG|HỘI ĐỒNG NHÂN DÂN PHƯỜNG)\s*:?$/i.test(line) ||
      /^KHÓA\s+[IVX0-9]+/i.test(line)) {
    return `<p style="text-align:center; text-indent:0; font-weight:bold; margin:6pt 0;">${escaped}</p>`;
  }
  if (/^[IVX]+\.\s+/i.test(line)) {
    return `<p style="text-align:left; text-indent:0; font-weight:bold; margin:6pt 0 4pt 0;">${escaped}</p>`;
  }
  const dieuMatch = line.match(/^(Điều\s+\d+\.|Kính gửi\s*:)\s*(.*)$/i);
  if (dieuMatch) {
    return `<p style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0;"><b>${escapeHtml(dieuMatch[1])}</b> ${escapeHtml(dieuMatch[2])}</p>`;
  }
  return `<p style="text-align:justify; text-indent:1.27cm; margin:0 0 6pt 0;">${escaped}</p>`;
}

function renderA4Preview() {
  const container = document.getElementById('a4PreviewPaper');
  if (!container) return;

  const standard = document.getElementById('vbtmStandard').value;
  const isParty = standard === 'hd05';
  const govOrg = document.getElementById('vbtmGovOrg').value.trim();
  const issueOrg = document.getElementById('vbtmIssueOrg').value.trim();
  const docNum = document.getElementById('vbtmDocNum').value.trim();
  const location = document.getElementById('vbtmLocation').value.trim();
  const docDate = document.getElementById('vbtmDocDate').value;
  const docType = document.getElementById('vbtmDocType').value.trim();
  const docSubject = document.getElementById('vbtmDocSubject').value.trim();
  const docBody = document.getElementById('vbtmDocBody').value;
  const recipients = document.getElementById('vbtmRecipients').value;
  const signerTitle = document.getElementById('vbtmSignerTitle').value;
  const signerName = document.getElementById('vbtmSignerName').value.trim();
  const includeQr = document.getElementById('vbtmIncludeQr')?.checked;

  const rightHeaderTop = isParty ? 'ĐẢNG CỘNG SẢN VIỆT NAM' : 'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM';
  const rightHeaderSub = isParty ? '' : 'Độc lập - Tự do - Hạnh phúc';
  const cleanNum = docNum.replace(/^Số\s*:?\s*/i, '');
  const numberLabel = isParty ? `Số ${cleanNum}` : `Số: ${cleanNum}`;

  const dateLine = formatVietnameseDate(docDate, location);
  const paragraphsHtml = docBody
    .split('\n')
    .filter(p => p.trim().length > 0)
    .map(p => formatBodyParagraphHtml(p))
    .join('');

  const recipientsHtml = recipients
    .split('\n')
    .filter(r => r.trim().length > 0)
    .map(r => `<div>${escapeHtml(r)}</div>`)
    .join('');

  const signerTitleHtml = signerTitle
    .split(/\n|\s+-\s+/)
    .filter(s => s.trim().length > 0)
    .map(s => `<div>${escapeHtml(s.trim())}</div>`)
    .join('');

  const qrCanvas = document.getElementById('vbtmQrCanvas');
  const qrDataUrl = (includeQr && qrCanvas) ? qrCanvas.toDataURL('image/png') : '';

  const leftSeparatorHtml = isParty
    ? `<div style="font-size: 13pt; font-weight: bold; line-height: 1; margin: 2px 0;">*</div>`
    : `<span class="a4-line-sep" style="width: 38%;"></span>`;

  const congVanSubjectHtml = (!docType && docSubject)
    ? `<div style="font-size: 12pt; margin-top: 4px;">${escapeHtml(docSubject)}</div>`
    : '';

  // Theo Hướng dẫn 05-HD/VPTW: Từ "Nơi nhận:" cỡ chữ 14pt, kiểu chữ đứng, có gạch chân dưới từ "Nơi nhận", sau có dấu hai chấm (:)
  // Theo Nghị định 30/2020/NĐ-CP: Từ "Nơi nhận:" cỡ chữ 12pt, in nghiêng, đậm
  const noiNhanHeaderHtml = isParty
    ? `<div style="font-size: 14pt; font-style: normal; font-weight: normal;"><u>Nơi nhận</u>:</div>`
    : `<div style="font-size: 12pt; font-weight: bold; font-style: italic;">Nơi nhận:</div>`;

  const noiNhanListFontSize = isParty ? '12pt' : '11pt';

  container.innerHTML = `
    <table class="a4-header-table">
      <tr>
        <td style="width: 44%; text-align: center; vertical-align: top;">
          ${govOrg ? `<div style="font-size: ${isParty ? '14pt' : '12.5pt'}; text-transform: uppercase;">${escapeHtml(govOrg)}</div>` : ''}
          <div style="font-size: ${isParty ? '14pt' : '12.5pt'}; font-weight: bold; text-transform: uppercase;">${escapeHtml(issueOrg)}</div>
          ${leftSeparatorHtml}
          <div style="font-size: ${isParty ? '14pt' : '13pt'}; margin-top: 6px;">${escapeHtml(numberLabel)}</div>
          ${congVanSubjectHtml}
        </td>
        <td style="width: 56%; text-align: center; vertical-align: top;">
          <div style="font-size: ${isParty ? '15pt' : '12.5pt'}; font-weight: bold; text-transform: uppercase;">${rightHeaderTop}</div>
          ${rightHeaderSub ? `<div style="font-size: 13.5pt; font-weight: bold;">${rightHeaderSub}</div>` : ''}
          <span class="a4-line-sep" style="width: ${isParty ? '58%' : '52%'};"></span>
          <div style="font-size: ${isParty ? '14pt' : '13.5pt'}; font-style: italic; margin-top: 8px;">${escapeHtml(dateLine)}</div>
        </td>
      </tr>
    </table>

    ${docType ? `<div class="a4-doc-title" style="font-size: ${isParty ? '15pt' : '14pt'}; font-weight: bold; text-align: center; margin-top: 14pt; margin-bottom: 4pt;">${escapeHtml(docType)}</div>` : ''}
    ${docType && docSubject ? `<div class="a4-doc-subject" style="font-size: 14pt; font-weight: bold; text-align: center; margin-bottom: 14pt;">${escapeHtml(docSubject)}<span class="a4-line-sep" style="width: 32%; margin-top: 5px;"></span></div>` : ''}

    <div class="a4-doc-body" style="margin-top: ${docType ? '4pt' : '12pt'};">
      ${paragraphsHtml}
    </div>

    <table class="a4-footer-table" style="margin-top: 16pt; width: 100%;">
      <tr>
        <td style="width: 50%; vertical-align: top; text-align: left; font-size: ${noiNhanListFontSize};">
          ${noiNhanHeaderHtml}
          <div style="line-height: 1.35; font-size: ${noiNhanListFontSize};">${recipientsHtml}</div>
          ${qrDataUrl ? `<div style="margin-top: 10px;"><img src="${qrDataUrl}" alt="QR" style="width: 74px; height: 74px; border: 1px solid #ddd; padding: 2px;"/><div style="font-size: 8.5pt; color: #555;">Mã QR xác thực VBTM</div></div>` : ''}
        </td>
        <td style="width: 50%; vertical-align: top; text-align: center;">
          <div style="font-weight: bold; font-size: ${isParty ? '14pt' : '13.5pt'}; text-transform: uppercase; line-height: 1.35;">${signerTitleHtml}</div>
          <div style="height: 68px;"></div>
          <div style="font-weight: bold; font-size: 14pt;">${escapeHtml(signerName)}</div>
        </td>
      </tr>
    </table>
  `;
}

function exportWordDocument() {
  const previewEl = document.getElementById('a4PreviewPaper');
  if (!previewEl) return;

  const docNum = preservedUploadedFileName || (document.getElementById('vbtmDocNum')?.value || 'VanBan').replace(/[^a-zA-Z0-9_-]/g, '_');

  const wordHtml = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns="http://www.w3.org/TR/REC-html40">
  <head>
    <meta charset="utf-8">
    <title>Van Ban Chuan Hoa ND30 & HD05</title>
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
        font-family: "Times New Roman", serif;
      }
      body {
        font-family: "Times New Roman", serif;
        font-size: 14pt;
        line-height: 1.38;
        color: #000000;
      }
      table { width: 100%; border-collapse: collapse; border: none; }
      td { vertical-align: top; font-family: "Times New Roman", serif; }
      td p { margin: 0; text-indent: 0; }
      p { margin: 0 0 6pt 0; font-size: 14pt; }
      .a4-doc-title { text-align: center; font-weight: bold; font-size: 14pt; text-transform: uppercase; margin-top: 16pt; margin-bottom: 4pt; }
      .a4-doc-subject { text-align: center; font-weight: bold; font-size: 14pt; margin-bottom: 14pt; }
      .a4-line-sep { display: block; border-bottom: 1px solid #000; margin: 4px auto; }
    </style>
  </head>
  <body>
    <div class="Section1">
      ${previewEl.innerHTML}
      <div style="mso-element:header" id="fh1"><p class="MsoHeader">&nbsp;</p></div>
      <div style="mso-element:header" id="h1">
        <p class="MsoHeader" align="center" style="text-align:center;font-size:13pt;font-family:'Times New Roman',serif;">
          <!--[if supportFields]><span style='mso-element:field-begin'></span> PAGE <span style='mso-element:field-separator'></span><![endif]-->2<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->
        </p>
      </div>
    </div>
  </body>
  </html>`;

  const blob = new Blob(['\ufeff', wordHtml], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `VBTM_ChuanHoa_${docNum}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  if (window.showToast) {
    window.showToast(`Đã xuất tệp Microsoft Word chuẩn A4: VBTM_ChuanHoa_${docNum}.doc`);
  }
}

function downloadQrPng() {
  const canvas = document.getElementById('vbtmQrCanvas');
  if (!canvas) return;
  const url = canvas.toDataURL('image/png');
  const a = document.createElement('a');
  a.href = url;
  a.download = 'VBTM_MaQR_VanBan.png';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  if (window.showToast) window.showToast('Đã tải xuống ảnh Mã QR chuẩn ISO 18004 (PNG)!');
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function updateVbtmLiveStats() {
  const el = document.getElementById('vbtmLiveStatsBadge');
  if (!el) return;
  const body = (document.getElementById('vbtmDocBody')?.value || '').trim();
  const words = body ? body.split(/\s+/).length : 0;
  const paras = body ? body.split(/\n+/).filter(l => l.trim()).length : 0;
  const std = document.getElementById('vbtmStandard')?.value === 'hd05' ? 'Chuẩn Đảng (05-HD/VPTW)' : 'Chuẩn Hành chính (NĐ 30)';
  el.textContent = `${std} · ${words} từ · ${paras} đoạn`;
}

function printVbtmA4Document() {
  const previewEl = document.getElementById('a4PreviewPaper');
  if (!previewEl) return;
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
  </style></head><body>${previewEl.innerHTML}</body></html>`);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 250);
}

function fillAiPromptAndRun(promptText) {
  const inp = document.getElementById('vbtmAiPromptInput');
  if (inp) inp.value = promptText;
  aiDraftDocumentFromPrompt();
}

function extractTopicFromPrompt(prompt) {
  let topic = String(prompt || '').trim();
  topic = topic.replace(/^(hãy\s+|vui lòng\s+|giúp tôi\s+|ai\s+)?(soạn thảo|soạn|viết|lập|tạo)\s+(giúp\s+|một\s+|bản\s+)?(báo cáo|nghị quyết|quyết định|kế hoạch|tờ trình|thông báo|công văn|giấy mời|chương trình|biên bản)?\s*(của\s+(đảng ủy|ubnd|ủy ban nhân dân|chi bộ)\s*(phường)?)?\s*(về việc|về|v\/v)?\s*/i, '');
  topic = topic.replace(/[.?!]+$/, '').trim();
  if (!topic) topic = 'triển khai nhiệm vụ trọng tâm và công tác chuyển đổi số tại địa phương';
  return topic.charAt(0).toUpperCase() + topic.slice(1);
}

async function aiDraftDocumentFromPrompt() {
  const promptText = (document.getElementById('vbtmAiPromptInput')?.value || '').trim();
  if (!promptText) {
    if (window.showToast) window.showToast('Vui lòng nhập yêu cầu nội dung cần AI soạn thảo!');
    return;
  }

  const apiKey = (document.getElementById('vbtmGeminiKey')?.value || '').trim();
  const isParty = /\b(đảng|đảng ủy|chi bộ|đảng bộ|ban thường vụ|bí thư|05-hd)\b/i.test(promptText);
  const stdVal = isParty ? 'hd05' : 'nd30';

  if (apiKey) {
    try {
      if (window.showToast) window.showToast('Đang kết nối Gemini AI để soạn thảo văn bản...');
      const sysPrompt = `Bạn là chuyên gia văn thư Việt Nam. Hãy soạn thảo văn bản từ yêu cầu: "${promptText}". Trả về JSON thuần (không markdown) với các khóa: standard ("hd05" hoặc "nd30"), docType (in hoa, để rỗng nếu là công văn), docSubject, docBody (phân tách đoạn bằng \\n\\n, kết thúc bằng ./.), recipients (mỗi dòng 1 nơi nhận), signerTitle, signerName.`;
      const resp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(apiKey)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts: [{ text: sysPrompt }] }] })
      });
      if (resp.ok) {
        const data = await resp.json();
        const rawTxt = data?.candidates?.[0]?.content?.parts?.[0]?.text || '';
        const jsonMatch = rawTxt.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
          const parsed = JSON.parse(jsonMatch[0]);
          document.getElementById('vbtmStandard').value = parsed.standard === 'hd05' ? 'hd05' : 'nd30';
          document.getElementById('vbtmGovOrg').value = parsed.standard === 'hd05' ? 'QUẬN ỦY' : 'ỦY BAN NHÂN DÂN QUẬN';
          document.getElementById('vbtmIssueOrg').value = parsed.standard === 'hd05' ? 'ĐẢNG ỦY PHƯỜNG' : 'ỦY BAN NHÂN DÂN PHƯỜNG';
          document.getElementById('vbtmDocType').value = parsed.docType || '';
          document.getElementById('vbtmDocSubject').value = parsed.docSubject || '';
          document.getElementById('vbtmDocBody').value = parsed.docBody || '';
          if (parsed.recipients) document.getElementById('vbtmRecipients').value = parsed.recipients;
          if (parsed.signerTitle) document.getElementById('vbtmSignerTitle').value = parsed.signerTitle;
          if (parsed.signerName) document.getElementById('vbtmSignerName').value = parsed.signerName;
          standardizeAndFixEverything();
          if (window.showToast) window.showToast('✓ Gemini AI đã soạn thảo và chuẩn hóa xong văn bản!');
          return;
        }
      }
    } catch (e) {
      console.warn('Gemini fallback to offline AI engine:', e);
    }
  }

  // Bộ máy AI Soạn Thảo Ngoại Tuyến 100%
  const lower = promptText.toLowerCase();
  let docType = '';
  let numCode = isParty ? '68-CV/ĐU' : '188/UBND-VP';
  if (lower.includes('báo cáo')) { docType = 'BÁO CÁO'; numCode = isParty ? '52-BC/ĐU' : '96/BC-UBND'; }
  else if (lower.includes('nghị quyết')) { docType = 'NGHỊ QUYẾT'; numCode = isParty ? '18-NQ/ĐU' : '12/NQ-HĐND'; }
  else if (lower.includes('quyết định')) { docType = 'QUYẾT ĐỊNH'; numCode = isParty ? '40-QĐ/ĐU' : '155/QĐ-UBND'; }
  else if (lower.includes('kế hoạch')) { docType = 'KẾ HOẠCH'; numCode = isParty ? '28-KH/ĐU' : '84/KH-UBND'; }
  else if (lower.includes('tờ trình')) { docType = 'TỜ TRÌNH'; numCode = isParty ? '14-TTr/ĐU' : '62/TTr-UBND'; }
  else if (lower.includes('thông báo')) { docType = 'THÔNG BÁO'; numCode = isParty ? '33-TB/ĐU' : '118/TB-UBND'; }
  else if (lower.includes('chương trình')) { docType = 'CHƯƠNG TRÌNH'; numCode = isParty ? '11-CTr/ĐU' : '08/CTr-UBND'; }

  const topic = extractTopicFromPrompt(promptText);
  const agency = isParty ? 'Đảng ủy phường' : 'Ủy ban nhân dân phường';
  const regRef = isParty ? 'Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng' : 'Nghị định số 30/2020/NĐ-CP của Chính phủ';

  let body = '';
  if (docType === 'QUYẾT ĐỊNH') {
    body = `Căn cứ chức năng, nhiệm vụ, quyền hạn của ${agency};\n\nCăn cứ ${regRef} và yêu cầu thực tiễn công tác tại địa phương;\n\nXét đề nghị của Văn phòng ${agency},\n\n${isParty ? 'BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH' : 'QUYẾT ĐỊNH:'}\n\nĐiều 1. Ban hành kèm theo Quyết định này các nội dung triển khai ${topic.charAt(0).toLowerCase() + topic.slice(1)}.\n\nĐiều 2. Giao các bộ phận chuyên môn, các ${isParty ? 'chi bộ trực thuộc' : 'tổ dân phố và công chức liên quan'} nghiêm túc tổ chức thực hiện bảo đảm tiến độ và chất lượng.\n\nĐiều 3. Quyết định này có hiệu lực kể từ ngày ký./.`;
  } else if (docType === 'BÁO CÁO') {
    body = `Thực hiện chương trình công tác và chỉ đạo của cấp trên, ${agency} báo cáo kết quả thực hiện ${topic.charAt(0).toLowerCase() + topic.slice(1)} như sau:\n\nI. KẾT QUẢ THỰC HIỆN\n1. ${agency} đã chủ động ban hành văn bản chỉ đạo, phân công nhiệm vụ cụ thể cho từng bộ phận phụ trách và tổ chức triển khai đồng bộ trên toàn địa bàn.\n\n2. Công tác tuyên truyền, kiểm tra, đôn đốc được thực hiện thường xuyên; chất lượng hồ sơ, văn bản được chuẩn hóa đúng quy định hiện hành.\n\nII. PHƯƠNG HƯỚNG, NHIỆM VỤ THỜI GIAN TỚI\nTiếp tục phát huy kết quả đạt được, kịp thời tháo gỡ khó khăn, vướng mắc và định kỳ tổng hợp báo cáo cấp trên theo quy định./.`;
  } else if (docType === 'KẾ HOẠCH' || docType === 'CHƯƠNG TRÌNH') {
    body = `Thực hiện chỉ đạo của cấp trên về ${topic.charAt(0).toLowerCase() + topic.slice(1)}, ${agency} xây dựng ${docType.toLowerCase()} triển khai thực hiện với các nội dung cụ thể như sau:\n\nI. MỤC ĐÍCH, YÊU CẦU\nNâng cao hiệu lực, hiệu quả quản lý, điều hành; bảo đảm sự phối hợp chặt chẽ, thống nhất giữa các bộ phận chuyên môn.\n\nII. NỘI DUNG VÀ TIẾN ĐỘ THỰC HIỆN\n1. Rà soát, hoàn thiện các quy trình nghiệp vụ và phân công rõ trách nhiệm cá nhân phụ trách.\n\n2. Tổ chức kiểm tra, đánh giá định kỳ hàng tháng và báo cáo kết quả về Văn phòng ${agency}.\n\nIII. TỔ CHỨC THỰC HIỆN\nĐề nghị các đơn vị, cá nhân liên quan nghiêm túc quán triệt và tổ chức thực hiện hiệu quả ${docType.toLowerCase()} này./.`;
  } else if (docType === 'NGHỊ QUYẾT') {
    body = `Thực hiện Quy chế làm việc của ${agency}, sau khi thảo luận và xem xét tình hình thực tế tại địa phương, ${agency} quyết nghị:\n\nI. ĐÁNH GIÁ TÌNH HÌNH\nCông tác lãnh đạo, chỉ đạo thực hiện ${topic.charAt(0).toLowerCase() + topic.slice(1)} thời gian qua đã đạt được nhiều chuyển biến tích cực, tạo sự đồng thuận cao trong cán bộ, đảng viên và nhân dân.\n\nII. NHIỆM VỤ VÀ GIẢI PHÁP CHỦ YẾU\n1. Tập trung nâng cao trách nhiệm người đứng đầu, tăng cường kỷ luật, kỷ cương hành chính và ứng dụng công nghệ thông tin.\n\n2. Giao Văn phòng ${agency} thường xuyên theo dõi, đôn đốc và tổng hợp báo cáo kết quả thực hiện Nghị quyết này./.`;
  } else {
    const subjPrefix = docType ? '' : 'V/v ';
    document.getElementById('vbtmDocSubject').value = subjPrefix + topic;
    body = `Kính gửi: ${isParty ? 'Các chi bộ trực thuộc Đảng ủy phường.' : 'Các bộ phận chuyên môn và Tổ trưởng các tổ dân phố thuộc phường.'}\n\nThực hiện chỉ đạo của cấp trên và yêu cầu nhiệm vụ công tác tại địa phương, ${agency} đề nghị các đơn vị triển khai thực hiện nghiêm túc nội dung ${topic.charAt(0).toLowerCase() + topic.slice(1)}:\n\n1. Chủ động rà soát các nhiệm vụ được giao, tổ chức thực hiện bảo đảm đúng tiến độ, chất lượng và quy trình quy định.\n\n2. Kịp thời báo cáo những khó khăn, vướng mắc (nếu có) về Văn phòng ${agency} để tổng hợp, tham mưu giải quyết.\n\nĐề nghị các đồng chí nghiêm túc triển khai thực hiện./.`;
  }

  document.getElementById('vbtmStandard').value = stdVal;
  document.getElementById('vbtmGovOrg').value = isParty ? 'QUẬN ỦY' : 'ỦY BAN NHÂN DÂN QUẬN';
  document.getElementById('vbtmIssueOrg').value = isParty ? 'ĐẢNG ỦY PHƯỜNG' : 'ỦY BAN NHÂN DÂN PHƯỜNG';
  document.getElementById('vbtmDocNum').value = numCode;
  document.getElementById('vbtmLocation').value = 'Phường';
  document.getElementById('vbtmDocType').value = docType;
  if (docType) document.getElementById('vbtmDocSubject').value = topic;
  document.getElementById('vbtmDocBody').value = body;
  document.getElementById('vbtmRecipients').value = isParty
    ? `- Ban Thường vụ Quận ủy (để b/c);\n- Các chi bộ trực thuộc;\n- Lưu Văn phòng Đảng ủy.`
    : `- Như trên;\n- Chủ tịch, các PCT UBND phường;\n- Lưu: VT, VP.`;
  document.getElementById('vbtmSignerTitle').value = isParty ? 'T/M ĐẢNG ỦY\nBÍ THƯ' : 'TM. ỦY BAN NHÂN DÂN\nCHỦ TỊCH';
  document.getElementById('vbtmSignerName').value = isParty ? 'Trần Minh Quân' : 'Nguyễn Văn An';

  standardizeAndFixEverything();
  if (window.showToast) {
    window.showToast(`✓ AI đã soạn thảo xong ${docType || 'CÔNG VĂN'} chuẩn ${isParty ? 'Hướng dẫn 05-HD/VPTW' : 'Nghị định 30/2020/NĐ-CP'}!`);
  }
}

function aiUpgradeWritingStyle() {
  const bodyEl = document.getElementById('vbtmDocBody');
  if (!bodyEl || !bodyEl.value.trim()) return;
  const styleReplacements = [
    [/\blàm gấp\b/gi, 'khẩn trương triển khai thực hiện'],
    [/\bbáo lại\b/gi, 'báo cáo kết quả về'],
    [/\bmọi người\b/gi, 'các đồng chí và cán bộ, công chức'],
    [/\bcác bạn\b/gi, 'các đồng chí'],
    [/\bcần phải\b/gi, 'đề nghị tập trung'],
    [/\blàm xong\b/gi, 'hoàn thành'],
    [/\bgửi gấp\b/gi, 'khẩn trương gửi'],
    [/\bkiểm tra kỹ\b/gi, 'tăng cường công tác kiểm tra, giám sát chặt chẽ'],
    [/\bxem lại\b/gi, 'rà soát, đánh giá lại'],
    [/\bcông việc\b/gi, 'nhiệm vụ được giao']
  ];
  let txt = bodyEl.value;
  let changed = 0;
  for (const [rx, rep] of styleReplacements) {
    if (rx.test(txt)) {
      txt = txt.replace(rx, rep);
      changed++;
    }
  }
  bodyEl.value = txt;
  standardizeAndFixEverything();
  if (window.showToast) {
    window.showToast(`✨ AI đã nâng cấp văn phong hành chính/Đảng (${changed} cụm từ trang trọng) & chuẩn hóa 100%!`);
  }
}

window.loadTemplate = loadTemplate;
window.fillAiPromptAndRun = fillAiPromptAndRun;
window.aiDraftDocumentFromPrompt = aiDraftDocumentFromPrompt;
window.aiUpgradeWritingStyle = aiUpgradeWritingStyle;
window.standardizeTextContent = standardizeTextContent;
window.standardizeAndFixEverything = standardizeAndFixEverything;
window.runSpellAndFormatCheck = runSpellAndFormatCheck;
window.fixSingleIssue = fixSingleIssue;
window.fixAllIssues = fixAllIssues;
window.insertSampleTypos = insertSampleTypos;
window.exportWordDocument = exportWordDocument;
window.printVbtmA4Document = printVbtmA4Document;
window.updateVbtmLiveStats = updateVbtmLiveStats;
window.downloadQrPng = downloadQrPng;
