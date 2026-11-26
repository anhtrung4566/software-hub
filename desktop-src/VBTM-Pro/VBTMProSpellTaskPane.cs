using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace VBTMProSuite
{
    // =========================================================================
    // HIGH-PRECISION OFFLINE VIETNAMESE SPELL & TELEX TYPO ENGINE (< 50ms)
    // =========================================================================
    public class SpellCardIssue
    {
        public string Wrong { get; set; }
        public string Suggest { get; set; }
        public string Badge { get; set; }      // "Chính tả" or "Nghi ngờ"
        public string Reason { get; set; }
        public int OccurrenceIndex { get; set; }
        public bool IsSuspicious { get; set; }
    }

    public static class OfflineVietnameseSpellEngine
    {
        private static readonly HashSet<string> ValidInitials = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "", "b", "c", "ch", "d", "đ", "g", "gh", "gi", "h", "k", "kh", "l", "m", "n",
            "ng", "ngh", "nh", "p", "ph", "q", "qu", "r", "s", "t", "th", "tr", "v", "x"
        };

        private static readonly HashSet<string> ValidCodas = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "", "c", "ch", "m", "n", "ng", "nh", "p", "t"
        };

        private static readonly HashSet<string> WhitelistWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "XIII", "XIV", "XV",
            "UBND", "HĐND", "MTTQ", "VPTW", "BCH", "BTV", "CAND", "QĐND", "TTXVN", "BHXH", "BHYT",
            "VNĐ", "VND", "USD", "PDF", "DOCX", "DOC", "XLSX", "PPTX", "HTML", "HTTP", "HTTPS",
            "WWW", "COM", "GOV", "VN", "ORG", "NET", "AI", "QR", "ISO", "OCR", "VBTM", "WIFI",
            "WORD", "EXCEL", "OFFICE", "WINDOWS", "MICROSOFT", "GOOGLE", "CHROME", "OLLAMA",
            "EMAIL", "ONLINE", "OFFLINE", "WEBSITE", "INTERNET", "ZALO", "FACEBOOK", "YOUTUBE",
            "BOONG", "SOONG", "XOONG", "GÒONG", "GOÒNG", "PC09", "PA05", "CSGT", "PCCC",
            "FILE", "FOLDER", "FORMAT", "STYLE", "TABLE", "CHART", "DATA", "CODE", "LINK",
            "WEB", "APP", "PRO", "PLUS", "SMART", "AUTO", "TEST", "DEMO", "ADMIN", "USER",
            "LOGIN", "TOKEN", "MODEL", "PROMPT", "CHAT", "CLOUD", "DRIVE", "SHEET", "SLIDE",
            "MAIL", "ZOOM", "MEET", "TEAMS", "SKYPE", "TIKTOK", "ADOBE", "PYTHON", "JAVA",
            "LINUX", "MACOS", "ANDROID", "IPHONE", "IPAD", "SAMSUNG", "VIETTEL", "VNPT", "FPT",
            "COVID", "UNESCO", "ASEAN", "APEC", "WTO", "WHO", "UNDP", "UNICEF", "GDP", "CPI",
            "VAT", "KPI", "OKR", "TCVN", "QCVN", "CMND", "CCCD", "BHTN", "THPT", "THCS",
            "CSVC", "TNHH", "CTCP", "KCNC", "SERVER", "CLIENT", "SYSTEM", "UPDATE", "VERSION",
            "PRINT", "START", "STOP", "ENTER", "SHIFT", "CTRL", "ALT", "SAVE", "OPEN", "CLOSE",
            "COPY", "PASTE", "EDIT", "VIEW", "HELP", "TOOL", "TOOLS", "MENU", "ICON", "FONT",
            "SIZE", "COLOR", "TEXT", "PAGE", "LINE", "LIST", "ITEM", "NOTE", "INFO", "DATE",
            "TIME", "NAME", "TYPE", "PATH", "PORT", "HOST", "LOCAL", "CHECK", "SCAN", "SEARCH"
        };

        private static readonly HashSet<string> CommonVietnameseSyllables = new HashSet<string>(
            (
                "a à á ả ã ạ ai ăn ân ấn ấp âu ẩu ấy ba bà bá bả bã bạ bác bạc bài bãi ban bàn bán bản bạn bang bảng bao bảo báo bát bay bắc bằng bắt bầu bây bấy bé bên bền bến bệnh bị bí bì bia biển biết biểu bình bỏ bó bò bồ bổ bộ bởi bốn bông bóng bột bới bước bướu bưởi " +
                "ca cà cá cả cã cạ các cách cái cam cảm cạn càng cảnh cao cáo cấp cắt câu cầu cấu cây cắm căn cắt cần cẩn cận cấp cất cân câu chà chả chán chánh chào cháu chạy chắc chắn chăm chăng chặt chất châu chấy che chè chế chia chỉ chi chiến chiếc chiều chịu chính cho chờ chở chợ chơi chọn chồng chống chỗ chủ chú chu chưa chức chứng chương chuyện chuyển chuyên chung chúng chuỗi chuông chuẩn chuỗi chúi chúc có cỏ cò cõ cọ coi con còn cọn công cộng cổng cống cơ cờ cỡ cớ cơm cứu cử cứ cửa cực cứng cười cường cứu " +
                "da dà đá dả dã dạ dài dám dán dân dẫn dần dấu dây dày dặn dắt đặt đầu đấu đấy dễ để đế đề đêm đến đều điểm điện điều diện diệt diễn diệu dinh dịp do đó dò dõ dọ dọc dòng dùng đúng đứng dư dự dựa đưa đường được dưới dương dưỡng duy duyệt " +
                "đa đà đá đả đã đạ đại đan đàn đản đạn đang đảng đạo đạt đau đắc đằng đặng đặt đầu đấu đây đầy đẩy đấy đen đèn đem đến đều đi địa điểm điền điều điện điệp đinh đỉnh định đo đỏ đó đò đọa đoạn đọc đòi đón đóng đông đồng động đỗ độ đổi đối đôn đốn đơn đợt đức đứng đưa đường được dưới đương đương đứt " +
                "e è é ẻ ẽ ẹ em ép êm ếch " +
                "ga gà gá gả gã gạ gái gan gắn gặp gần gấp gắt gật gây gãy ghế ghi gia già giá giả giã giạ giai giam giảm gian giản giang giao giáo giáp giàu gay gãy gay gắt gầy gây giấy ghép ghé ghi gì gia giảm gián giao giáp giàu giây giấy giữ giữa gọi gói gọn góp gỗ gồ gồm gốc gối giới gửi gương gươm gương " +
                "ha hà há hả hã hạ hai hại hàm hàn hạn hàng hằng hào hảo háo hát hay hãy hặc hằn hằng hẳn hẳn hắt hầu hầu hậu hệ hết hi hiển hiện hiệp hiệu hình hò hỏi hóa hòa họa hoài hoàn hoãn hoạt học hỏi hôn hỗn hồng hổ hỗ hộ hồi hội hối hỗn hợp hơi hơn hứa hướng hưởng hư hử hứ hữu huy huyền huyện huyết hủy " +
                "i ì í ỉ ĩ ị ích in ít " +
                "ke kẻ kế kệ kem kém kèm ken kén kéo kép kết kêu kể kê kênh kết kha khả khá khác khai khái kham khảm khám khan khản khang kháng khảo khát khắc khăn khăng khẳng khắp khắt khâu khẩu khu khủ khúc khuê khuôn khuya khuyên khuyến khích khi khí khiêm khiến khiếu kho khó khỏ khoa khóa khoan khoản khoáng khoát khóc khỏi khói khóm không khổ khố khô khối khôn khủng khởi khu khuất khúc khuông khung khuynh khứa " +
                "la là lá lả lã lạ lại làm lan lạn làng lãnh lao lão lát lay lắp lắt lặng lặp lâm lầm lẫn lập lâu lầu lấy lẻ lề lễ lên lệnh lệ li lịch liên liền liệm liều liệu linh lĩnh lính lo lò ló lỏ lọ loa loại loan loàn loãng lọc lợi lộn lồng lối lỗi lộc lớn lời lợi lùi lúc luật luận lùng luôn lượt lương lượng lưới lướt lưu lựu lý " +
                "ma mà má mả mã mạ mai mái mại mang màng mạng mạnh mao màu máy mặc mặn mắng mắt mặt mâm mầm mẫn mất mẫu mây mấy mẹ mẻ mè mê mềm mệnh mi miền miễn mình mịn mọi mỏi món mong mỏng mổ mồ mỗ mộ môi mỗi mối mồi môn mông mộng một mới mời mở mỡ mớ mua mùa múa mục mùi mũi mụn muôn muốn muộn mưa mừa mứa mức mực mừng mười mượn mứt mỹ " +
                "na nà ná nả nã nạ nai nài nại nam năm nắm năn nắn năng nắng nặng nâng nấp nâu nấu này nảy nãy nay nấy nề nẻ né nêm nếm nên nền nếp nếu nêu ngà ngả ngã ngạ ngại ngài ngàn ngạn ngang ngàng ngàng ngày ngắn ngắt ngẫu ngây nghe nghè nghé nghề nghệ nghi nghỉ nghĩ nghị nghiêm nghiên nghiện nghèo nghẽn nghẹt ngó ngỏ ngõ ngọ ngoa ngoái ngoài ngoan ngoãn ngoặt ngọc ngòi ngọn ngon ngọt ngôi ngô ngộ ngôn ngông ngơ ngỡ ngớ ngủ ngũ ngu nguồn nguyên nguyễn nguyệt nguy hiểm ngữ ngư ngự ngựa ngược ngươi người ngưỡng ngưng ngừng nha nhà nhá nhả nhã nhạ nhai nhầm nhân nhận nhẫn nhắn nhặt nhăn nhằng nhất nhạt nhau nhảy nhạy nhắc nhắm nhắn nhặt nhầm nhẫn nhập nhất nhì nhi nhiêm nhiệm nhiên nhiệt nhiều nhiễu nhịn nhìn nhịp nho nhỏ nhõ nhọ nhọc nhóm nhọn nhổ nhô nhộn nhớ nhỡ nhờ nhơn nhơ nhu nhủ nhũ nhu cầu nhuận nhục nhún nhung nhường nhược như nhữ nhứ nhựa nhưng những " +
                "o ò ó ỏ õ ọ oa oai oan oán oanh óc ôi ối ôm ốm ôn ổn ồn ông ổng ống ộp ở ỡ ớ ơi ới ơn ớt " +
                "pha phà phá phả phã phạ phai phải phàm phàn phán phản phạm phan phanh phao phát phay phần phân phấn phận phất phẩu phê phế phên phi phí phì phỉ phia phiên phiện phiếu phím phìn phình pho phó phòng phỏng phố phổi phồn phổ phối phong phóng phụ phủ phú phù phu phúc phục phun phụng phút phương phượng phượt " +
                "qua quà quá quả quã quạ quai quái quan quản quán quạn quang quàng quảng quãng quát quay quắc quặn quăng quắt quấy quê quề quế quên quét qui quí quì quỉ quĩ quị quy quý quỳ quỷ quỹ quỵ quyền quyết quốc quở quơ quỡ quớ quờ quợ " +
                "ra rà rá rả rã rạ rải ranh rành rãnh rao ráo rào rạo rát ray rắc răn rắn rằng rặng rắp rắt rậm rần rấn rập rất râu rầu rầy rẫy rẻ rẽ rệ rèm rèn rẽo rét rêu rì rỉ rĩ rị ria riêng riệt rinh rình rính ro rõ rọ rọc roi rọi ròn rón rong ròng rộng rổ rồ rộ rỗi rồi rối rỗng rốt rơi rời rợi rơm rớm rợn ru rủ rũ rụ rủa rục rui rủi rùi run rủn rung rùng rụng ruồi ruộng rút rứa rữa rửa rực rừng rưng rửng rước rưới rướn rượt rượu " +
                "sa sà sá sả sã sạ sai sài sải sạm san sàn sản sạn sang sàng sáng sảng sành sao sào sáo sảo sáp sát sau sáu say sắc sắm săn sẵn sáng sắp sắt sâm sầm sấm sẩm sậm sân sần sấn sập sâu sầu sấu sẽ sẻ sé sê sến sếp si sĩ sỉ sỹ sỉa siêng siết sim sin sinh sình sính so sò só sỏ sõ sọ soa soái soạn soạt soi sói sỏi sòi son sonh song sòng sóng sỏng sót số sổ sỗ sộ sôi sồi sối sổi sống sốt sơ sở sờ sỡ sớ sợ sớm sơn sờn sớn sợi su sù sú sủ sụ sua suối suốt suy suýt sư sử sứ sự sữa sửa sức sưng sừng sững sước sườn sướng sượt sứt " +
                "ta tà tá tả tã tạ tai tài tái tải tại tam tàm tám tạm tan tàn tán tản tạn tang tàng táng tảng tạng tanh tạnh tao tào táo tảo tạo táp tát tau tàu táu tay tày táy tắc tăm tằm tắm tắn tăng tằng tặng tắt tâm tầm tấm tẩm tậm tân tần tấn tẩn tận tập tất tây tầy tấy tẩu te tè té tẻ tẽ tẹ tem tên tế tề tệ tết thà thả thã thạ thai thải thái tham thảm thám than thản thán thang thàng tháng thẳng thanh thảnh thánh thành thạo tháo thảo thác thách thăm thẳm thắm thăn thắn thăng thẳng thắng thắp thắt thâm thầm thấm thẩm thậm thân thần thấn thận thấp thất thâu thấu thầu thấy thầy thẩy thể thế thề thệ thêm thềm thèm thẹn theo thép thét thi thì thí thỉ thị thia thích thiên thiền thiển thiện thiệp thiết thiếu thiệu thinh thỉnh thịnh thít thịt tho thỏ thó thò thọ thoa thỏa thoái thoải thoảng thoát thọc thoi thói thỏi thòi thom thòm thôn thổ thố thồ thôi thối thổi thông thống thổng thốt thơ thở thớ thờ thợ thời thơm thu thù thú thủ thụ thua thuấn thuận thuật thuê thuế thuở thuôn thuổng thuốc thuộm thuôn thuồng thưa thừa thửa thửa thức thực thừng thưởng thước thướt thượt thư thử thứ tự " +
                "ti tỉ tí tì tị tia tích tiền tiên tiến tiễn tiện tiếp tiết tiêu tiểu tiếu tiệu tim tìm tím tin tín tỉnh tình tính tít thịt to tò tó tỏ tõ tọ toa tỏa tọa toàn toán toan toang toáng toát tóc toi tỏi tòi tóm tọn tòng tông tổng tống tốt tổ tố tồ tội tối tồi tôn tồn tốn tổn tông tống tơ tờ tớ tở tợ tới tơi tời tớp tu tù tú tủ tụ tua tuần tuân tuấn tuất tuế tục tuổi túi tủi tùi tùm tụm tung tùng túng tụng tuôn tuột tuy tùy túy tủy tụy tuyên tuyển tuyến tuyệt tư từ tứ tử tự tựa tức tực từng tứng tưng tương tường tưởng tượng tươi tười tưới tưởi tước tướt tựu " +
                "u ù ú ủ ũ ụ ủa ủi ủn ủng uốn uống uất ư ừ ứ ử ữ ự ưa ứa ửa ức ưng ừng ứng ựng ước ướt ưu uy ủy úy uý uyên uyển " +
                "va và vá vả vã vạ vai vài vái vải vại vam vạm van vàn ván vản vạn vang vàng váng vảng vạng vành vao vào váo vảo váp vát vay vày váy vảy vạy vắc văm vằm vắn văng vằng vắng vẳng vắt vặt vân vần vấn vẩn vận vấp vất vây vầy vấy vẩy vậy ve vè vé vẻ vẽ vẹ ven vén vèn vẹn veo vèo véo vẹo vét vêu về vẻ vé vế vệ vên vệnh vi vì ví vỉ vĩ vị vỉa việc viên viền viễn viện viết việt vinh vịnh vít vịt vo vò vó vỏ võ vọ vọc voi vòi vói vỏi vọi vòm von vón vòng võng vọng vót vọt vô vồ vố vổ vỗ vộ vôi vội vồn vốn vồng vơ vờ vớ vở vỡ vợ vơi vời với vợi vu vù vú vũ vụ vua vủa vui vùi vung vùng vũng vụng vuốt vuột vưa vừa vứa vửa vựa vực vừng vững vưng vước vương vướng vượng vượt vứt " +
                "xa xà xá xả xã xạ xài xám xan xán xang xàng xáng xảnh xao xào xáo xảo xáp xát xay xáy xắc xăm xắm xăn xắn xăng xằng xắp xắt xâm xầm xấm xẩm xậm xân xấp xất xấu xâu xầu xây xầy xấy xe xè xé xẻ xẽ xẹ xem xen xén xèn xẻn xẹo xép xét xê xế xề xệ xếp xi xì xí xỉ xĩ xị xía xích xiêm xiển xiết xin xỉn xinh xình xính xít xịt xo xò xó xỏ xõ xọ xoa xóa xỏa xoài xoan xoàn xoán xoang xoàng xoảng xoay xoáy xoắc xoăn xoắn xoắt xóc xoi xói xỏi xòi xom xóm xỏm xòm xon xón xong xòng xóng xót xọt xô xồ xố xổ xỗ xộ xôi xối xổi xồi xồm xôm xông xống xốp xốt xơ xờ xớ xở xỡ xợ xơi xới xởi xời xới xơm xớm xu xù xú xủ xũ xụ xua xủa xuẩn xuân xuấn xuất xúc xui xúi xủi xùi xum xùm xúm xủm xun xùn xún xung xùng xúng xụng xuôi xuồng xuổng xúp xút xuy xuýt xưa xửa xứa xức xưng xừng xứng xựng xước xướng xướt xử xứ sự " +
                "y ỳ ý ỷ ỹ ỵ yểm yếm yên yến yểm yết yêu yếu yểu"
            ).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        private static readonly HashSet<string> CommonVietnameseBigrams = new HashSet<string>(
            new string[]
            {
                "tập trung", "vận động", "cơ cấu", "đảng viên", "thành lập", "xung đột", "nước nhà", "nước và", "nhà nước",
                "phát triển", "thực hiện", "tổ chức", "cán bộ", "quản lý", "chính sách", "kinh tế", "xã hội", "an ninh",
                "quốc phòng", "xây dựng", "bảo vệ", "kiểm tra", "giám sát", "đoàn kết", "nhiệm vụ", "mục tiêu", "giải pháp",
                "kế hoạch", "nghị quyết", "chủ trương", "đường lối", "chất lượng", "hiệu quả", "trách nhiệm", "quyền lực",
                "chính quyền", "nhân dân", "mặt trận", "đoàn thể", "cơ quan", "đơn vị", "địa phương", "trung ương", "cấp ủy",
                "chi bộ", "đảng bộ", "thường vụ", "chấp hành", "tư tưởng", "đạo đức", "lối sống", "tuyên truyền", "giáo dục",
                "đào tạo", "bồi dưỡng", "khoa học", "công nghệ", "chuyển đổi", "hành chính", "cải cách", "thủ tục", "dịch vụ",
                "ngân sách", "tài chính", "đầu tư", "hạ tầng", "môi trường", "tài nguyên", "y tế", "văn hóa", "thể thao",
                "du lịch", "an sinh", "giảm nghèo", "lao động", "việc làm", "pháp luật", "tư pháp", "thanh tra", "khiếu nại",
                "tố cáo", "phòng chống", "tham nhũng", "tiêu cực", "lãng phí", "kiểm điểm", "phê bình", "sâu sắc", "nghiêm túc",
                "nghiêm minh", "kỷ luật", "kỷ cương", "nắm giữ", "vai trò", "nòng cốt", "đúng người", "đứng người", "đảng phái",
                "cục bộ", "lợi ích", "tự phê", "chỉnh đốn", "hệ thống", "chính trị", "toàn diện", "đồng bộ", "kịp thời",
                "chủ động", "tích cực", "sáng tạo", "đổi mới", "nâng cao", "tăng cường", "bảo đảm", "đảm bảo", "hoàn thiện",
                "triển khai", "ban hành", "quy định", "hướng dẫn", "đánh giá", "tổng kết", "sơ kết", "báo cáo", "thông báo",
                "quyết định", "công tác", "nội dung", "phương pháp", "đối tượng", "thời gian", "không gian", "điều kiện",
                "hoạt động", "phong trào", "thi đua", "khen thưởng", "xử lý", "giải quyết", "tháo gỡ", "khó khăn", "vướng mắc",
                "tồn tại", "hạn chế", "nguyên nhân", "bài học", "kinh nghiệm", "phương hướng", "trọng tâm", "đột phá",
                "chiến lược", "quy hoạch", "dự án", "chương trình", "đề án", "công trình", "sản xuất", "kinh doanh",
                "thương mại", "xuất khẩu", "nhập khẩu", "thị trường", "doanh nghiệp", "hợp tác", "liên kết", "hội nhập",
                "quốc tế", "đối ngoại", "chủ quyền", "biên giới", "hải đảo", "trật tự", "an toàn", "giao thông", "đô thị",
                "nông thôn", "nông nghiệp", "công nghiệp", "xây lắp", "giao lưu", "văn bản", "hồ sơ", "lưu trữ", "văn thư",
                "lãnh đạo", "chỉ đạo", "điều hành", "phối hợp", "hỗ trợ", "tạo điều", "thuận lợi", "công khai", "minh bạch",
                "dân chủ", "công bằng", "văn minh", "hiện đại", "bền vững", "ổn định", "phồn vinh", "hạnh phúc", "độc lập",
                "tự do", "thống nhất", "toàn vẹn", "lãnh thổ", "tổ quốc", "đất nước", "dân tộc", "giai cấp", "công nhân",
                "nông dân", "trí thức", "thanh niên", "phụ nữ", "cựu chiến", "người cao", "trẻ em", "gia đình", "cộng đồng",
                "cơ sở", "phường xã", "quận huyện", "tỉnh thành", "thành phố", "thủ đô", "tổ dân", "khu phố", "thôn xóm",
                "bản làng", "ấp khóm", "bầu cử", "ứng cử", "đại biểu", "đại hội", "hội nghị", "hội thảo", "tọa đàm",
                "gặp mặt", "tiếp xúc", "cử tri", "lắng nghe", "ý kiến", "kiến nghị", "phản ánh", "nguyện vọng", "chính đáng",
                "quyền lợi", "nghĩa vụ", "công dân", "nhà máy", "xí nghiệp", "trường học", "bệnh viện", "cơ sở", "trụ sở",
                "học tập", "làm theo", "tấm gương", "phong cách", "hồ chí", "chủ nghĩa", "mác lê", "cách mạng", "truyền thống",
                "lịch sử", "anh hùng", "liệt sĩ", "thương binh", "có công", "chính sách", "đền ơn", "đáp nghĩa", "nhân đạo",
                "từ thiện", "xóa đói", "nhà ở", "đất đai", "giải phóng", "mặt bằng", "tái định", "bồi thường", "cấp giấy",
                "chứng nhận", "đăng ký", "cấp phép", "thuế phí", "lệ phí", "kiểm toán", "quyết toán", "dự toán", "phân bổ",
                "sử dụng", "tiết kiệm", "hiệu lực", "thi hành", "áp dụng", "bổ sung", "sửa đổi", "thay thế", "bãi bỏ",
                "tổng hợp", "thống kê", "phân tích", "dự báo", "cảnh báo", "ứng phó", "thiên tai", "dịch bệnh", "biến đổi",
                "khí hậu", "năng lượng", "tái tạo", "số hóa", "dữ liệu", "thông tin", "truyền thông", "báo chí", "xuất bản",
                "mạng xã", "không gian", "an ninh", "trực tuyến", "tự động", "thông minh", "kết nối", "chia sẻ", "đồng thuận"
            },
            StringComparer.OrdinalIgnoreCase);

        private static readonly string[,] ContextAndTypoPairs = new string[,]
        {
            // Cụm từ hay nhầm thanh điệu / ngữ cảnh (Nghi ngờ)
            { "Đảng phải", "Đảng phái", "Nghi ngờ", "Nhầm thanh điệu 'phải' -> 'phái'" },
            { "đảng phải", "đảng phái", "Nghi ngờ", "Nhầm thanh điệu 'phải' -> 'phái'" },
            { "năm giữ", "nắm giữ", "Nghi ngờ", "Nhầm dấu 'năm giữ' -> 'nắm giữ'" },
            { "Năm giữ", "Nắm giữ", "Nghi ngờ", "Nhầm dấu 'Năm giữ' -> 'Nắm giữ'" },
            { "đúng người", "đứng người", "Nghi ngờ", "Kiểm tra ngữ cảnh 'đúng người' / 'đứng người'" },
            { "sát nhập", "sáp nhập", "Chính tả", "Chuẩn chính tả: 'sáp nhập'" },
            { "bổ xung", "bổ sung", "Chính tả", "Chuẩn chính tả: 'bổ sung'" },
            { "xuất ăn", "suất ăn", "Chính tả", "Chuẩn chính tả: 'suất ăn'" },
            { "nghành", "ngành", "Chính tả", "Chuẩn chính tả: 'ngành'" },
            { "chuẩn đoán", "chẩn đoán", "Chính tả", "Chuẩn chính tả: 'chẩn đoán'" },
            { "thăm quan", "tham quan", "Nghi ngờ", "Chuẩn hành chính: 'tham quan'" },
            { "chín mùi", "chín muồi", "Chính tả", "Chuẩn chính tả: 'chín muồi'" },
            { "xác xuất", "xác suất", "Chính tả", "Chuẩn chính tả: 'xác suất'" },
            { "sử lý", "xử lý", "Chính tả", "Chuẩn chính tả: 'xử lý'" },
            { "sữ lý", "xử lý", "Chính tả", "Chuẩn chính tả: 'xử lý'" },
            { "xữ lý", "xử lý", "Chính tả", "Chuẩn chính tả: 'xử lý'" },
            { "đề xuất ý kiên", "đề xuất ý kiến", "Chính tả", "Thiếu dấu sắc 'ý kiến'" },
            { "chỉ đao", "chỉ đạo", "Chính tả", "Thiếu dấu nặng 'chỉ đạo'" },
            { "lãnh đao", "lãnh đạo", "Chính tả", "Thiếu dấu nặng 'lãnh đạo'" },
            { "thưc hiện", "thực hiện", "Chính tả", "Thiếu dấu nặng 'thực hiện'" },
            { "tổ chưc", "tổ chức", "Chính tả", "Thiếu dấu sắc 'tổ chức'" },
            { "nghị quyêt", "nghị quyết", "Chính tả", "Thiếu dấu sắc 'nghị quyết'" },
            { "kế hoach", "kế hoạch", "Chính tả", "Thiếu dấu nặng 'kế hoạch'" },
            { "nhiệm vu", "nhiệm vụ", "Chính tả", "Thiếu dấu nặng 'nhiệm vụ'" },
            { "cán bô", "cán bộ", "Chính tả", "Thiếu dấu nặng 'cán bộ'" },
            { "chính sach", "chính sách", "Chính tả", "Thiếu dấu sắc 'chính sách'" },
            { "kiểm điêm", "kiểm điểm", "Chính tả", "Thiếu dấu hỏi 'kiểm điểm'" },
            { "phê binh", "phê bình", "Chính tả", "Thiếu dấu huyền 'phê bình'" },
            { "nghiêm tuc", "nghiêm túc", "Chính tả", "Thiếu dấu sắc 'nghiêm túc'" },
            { "triển khai thưc", "triển khai thực", "Chính tả", "Thiếu dấu nặng 'thực'" },
            { "giải quyêt", "giải quyết", "Chính tả", "Thiếu dấu sắc 'giải quyết'" },
            { "phát triên", "phát triển", "Chính tả", "Thiếu dấu hỏi 'phát triển'" },
            { "đảm bảo", "bảo đảm", "Nghi ngờ", "Gợi ý văn bản hành chính: 'bảo đảm'" },
            { "chấp hành nghiêm chỉnh", "chấp hành nghiêm chỉnh", "Nghi ngờ", "" },
            { "độc lập tự do", "Độc lập - Tự do", "Nghi ngờ", "Chuẩn quốc hiệu" },
            { "công tác tuyên truyên", "công tác tuyên truyền", "Chính tả", "Thiếu dấu huyền 'tuyên truyền'" },
            { "trách nhiêm", "trách nhiệm", "Chính tả", "Thiếu dấu nặng 'trách nhiệm'" },
            { "chủ trương đương lối", "chủ trương đường lối", "Chính tả", "Thiếu dấu huyền 'đường lối'" },
            { "quán triêt", "quán triệt", "Chính tả", "Thiếu dấu nặng 'quán triệt'" },
            { "đề nghi", "đề nghị", "Chính tả", "Thiếu dấu nặng 'đề nghị'" },
            { "báo cao", "báo cáo", "Chính tả", "Thiếu dấu sắc 'báo cáo'" },
            { "thông bao", "thông báo", "Chính tả", "Thiếu dấu sắc 'thông báo'" },
            { "quyêt định", "quyết định", "Chính tả", "Thiếu dấu sắc 'quyết định'" },
            { "quyết đinh", "quyết định", "Chính tả", "Thiếu dấu nặng 'quyết định'" },
            { "chỉnh đôn", "chỉnh đốn", "Chính tả", "Thiếu dấu sắc 'chỉnh đốn'" },
            { "đoàn kêt", "đoàn kết", "Chính tả", "Thiếu dấu sắc 'đoàn kết'" },
            { "nội dung công tac", "nội dung công tác", "Chính tả", "Thiếu dấu sắc 'công tác'" },
            { "hiêu quả", "hiệu quả", "Chính tả", "Thiếu dấu nặng 'hiệu quả'" },
            { "chất lương", "chất lượng", "Chính tả", "Thiếu dấu nặng 'chất lượng'" },
            { "đối tương", "đối tượng", "Chính tả", "Thiếu dấu nặng 'đối tượng'" },
            { "phương phap", "phương pháp", "Chính tả", "Thiếu dấu sắc 'phương pháp'" },
            { "thời gian quy đinh", "thời gian quy định", "Chính tả", "Thiếu dấu nặng 'quy định'" },
            { "giám sat", "giám sát", "Chính tả", "Thiếu dấu sắc 'giám sát'" },
            { "thanh tra kiểm tra", "thanh tra, kiểm tra", "Nghi ngờ", "Cân nhắc dấu phẩy phân tách" },
            { "sâu săc", "sâu sắc", "Chính tả", "Thiếu dấu sắc 'sâu sắc'" },
            { "sâu sac", "sâu sắc", "Chính tả", "Thiếu dấu 'sâu sắc'" },
            { "sư dụng", "sử dụng", "Chính tả", "Thiếu dấu hỏi 'sử dụng'" },
            { "xử dụng", "sử dụng", "Chính tả", "Chuẩn chính tả: 'sử dụng'" },
            { "ngiêm", "nghiêm", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "ngĩ", "nghĩ", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "ngĩa", "nghĩa", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "ngị", "nghị", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "ngề", "nghề", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "ngệ", "nghệ", "Chính tả", "Thiếu chữ 'h' sau 'ng' trước 'i/e/ê'" },
            { "kỹ luât", "kỷ luật", "Chính tả", "Chuẩn chính tả: 'kỷ luật'" },
            { "kỹ cương", "kỷ cương", "Chính tả", "Chuẩn chính tả: 'kỷ cương'" },
            { "kỹ niệm", "kỷ niệm", "Chính tả", "Chuẩn chính tả: 'kỷ niệm'" }
        };

        private static readonly HashSet<string> ValidBaseNuclei = new HashSet<string>(StringComparer.Ordinal)
        {
            "a", "ă", "â", "e", "ê", "i", "o", "ô", "ơ", "u", "ư", "y",
            "ai", "ao", "au", "âu", "ay", "ây", "eo", "êu", "ia", "iê", "iu",
            "oa", "oă", "oe", "oi", "ôi", "ơi", "oo", "ua", "uâ", "uê", "ui",
            "uô", "uơ", "uy", "ưa", "ươ", "ưu", "yê",
            "iêu", "oai", "oao", "oay", "oeo", "uai", "uây", "uôi", "uya", "uyê", "uyu", "ươi", "ươu", "yêu"
        };

        private static readonly Dictionary<string, string> DirectWordRepairs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "ngời", "người" }, { "ngưoi", "người" }, { "ngươi", "người" },
            { "đưoc", "được" }, { "đuợc", "được" }, { "đươc", "được" },
            { "thưc", "thực" }, { "quyêt", "quyết" }, { "hoach", "hoạch" },
            { "chưc", "chức" }, { "nhưg", "nhưng" }, { "khôg", "không" },
            { "cũg", "cũng" }, { "trườg", "trường" }, { "tưởg", "tưởng" },
            { "chươg", "chương" }, { "chưong", "chương" }, { "đồg", "đồng" },
            { "côg", "công" }, { "hòan", "hoàn" }, { "tòan", "toàn" },
            { "quôc", "quốc" }, { "luât", "luật" }, { "phap", "pháp" },
            { "nhât", "nhất" }, { "môt", "một" }, { "tôt", "tốt" },
            { "đôc", "độc" }, { "lâp", "lập" }, { "hơp", "hợp" },
            { "tâp", "tập" }, { "câp", "cấp" }, { "kêt", "kết" },
            { "viêt", "viết" }, { "biêt", "biết" }, { "xuât", "xuất" },
            { "phat", "phát" }, { "đat", "đạt" }, { "mât", "mất" },
            { "hêt", "hết" }, { "thiêt", "thiết" }, { "triêt", "triệt" },
            { "nghiêt", "nghiệt" }, { "nhiêt", "nhiệt" }, { "săc", "sắc" },
            { "tuc", "túc" }, { "đăc", "đặc" }, { "măc", "mặc" },
            { "băc", "bắc" }, { "mưc", "mức" }, { "sưc", "sức" },
            { "lưc", "lực" }, { "tưc", "tức" }, { "kiêm", "kiểm" },
            { "điêm", "điểm" }, { "nghiêm", "nghiêm" }, { "ngiêm", "nghiêm" },
            { "ngĩ", "nghĩ" }, { "ngĩa", "nghĩa" }, { "ngị", "nghị" },
            { "ngề", "nghề" }, { "ngệ", "nghệ" }, { "nghành", "ngành" }
        };

        private static bool IsVietnameseVowel(char c)
        {
            const string vowels = "aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵ";
            return vowels.IndexOf(char.ToLower(c)) >= 0;
        }

        private static int GetToneType(char c)
        {
            // 0 = ngang, 1 = sac, 2 = huyen, 3 = hoi, 4 = nga, 5 = nang
            char lc = char.ToLower(c);
            if ("áắấéếíóốớúứý".IndexOf(lc) >= 0) return 1;
            if ("àằầèềìòồờùừỳ".IndexOf(lc) >= 0) return 2;
            if ("ảẳẩẻểỉỏổởủửỷ".IndexOf(lc) >= 0) return 3;
            if ("ãẵẫẽễĩõỗỡũữỹ".IndexOf(lc) >= 0) return 4;
            if ("ạặậẹệịọộợụựỵ".IndexOf(lc) >= 0) return 5;
            return 0;
        }

        private static char StripToneOnly(char c)
        {
            char lc = char.ToLower(c);
            if ("aàáảãạ".IndexOf(lc) >= 0) return 'a';
            if ("ăằắẳẵặ".IndexOf(lc) >= 0) return 'ă';
            if ("âầấẩẫậ".IndexOf(lc) >= 0) return 'â';
            if ("eèéẻẽẹ".IndexOf(lc) >= 0) return 'e';
            if ("êềếểễệ".IndexOf(lc) >= 0) return 'ê';
            if ("iìíỉĩị".IndexOf(lc) >= 0) return 'i';
            if ("oòóỏõọ".IndexOf(lc) >= 0) return 'o';
            if ("ôồốổỗộ".IndexOf(lc) >= 0) return 'ô';
            if ("ơờớởỡợ".IndexOf(lc) >= 0) return 'ơ';
            if ("uùúủũụ".IndexOf(lc) >= 0) return 'u';
            if ("ưừứửữự".IndexOf(lc) >= 0) return 'ư';
            if ("yỳýỷỹỵ".IndexOf(lc) >= 0) return 'y';
            return lc;
        }

        private static bool HasVietnameseDiacritic(string s)
        {
            foreach (char c in s)
            {
                if (c > 127) return true;
            }
            return false;
        }

        public static bool IsValidVietnameseSyllable(string word)
        {
            if (string.IsNullOrWhiteSpace(word) || word.Length > 7) return false;
            string w = word.ToLowerInvariant().Normalize(NormalizationForm.FormC);

            int firstVowel = -1;
            int lastVowel = -1;
            for (int i = 0; i < w.Length; i++)
            {
                if (IsVietnameseVowel(w[i]))
                {
                    if (firstVowel < 0) firstVowel = i;
                    lastVowel = i;
                }
            }
            if (firstVowel < 0) return false;

            string initial = w.Substring(0, firstVowel);
            string nucleus = w.Substring(firstVowel, lastVowel - firstVowel + 1);
            string coda = w.Substring(lastVowel + 1);

            if (initial == "q" && nucleus.Length >= 2 && nucleus[0] == 'u')
            {
                initial = "qu";
                nucleus = nucleus.Substring(1);
            }
            else if (initial == "g" && nucleus.Length >= 2 && nucleus[0] == 'i')
            {
                initial = "gi";
                nucleus = nucleus.Substring(1);
            }

            if (!ValidInitials.Contains(initial)) return false;
            if (!ValidCodas.Contains(coda)) return false;
            if (nucleus.Length == 0 || nucleus.Length > 3) return false;

            StringBuilder baseNuc = new StringBuilder();
            int toneCount = 0;
            int toneVal = 0;
            foreach (char c in nucleus)
            {
                if (!IsVietnameseVowel(c)) return false;
                int t = GetToneType(c);
                if (t > 0) { toneCount++; toneVal = t; }
                baseNuc.Append(StripToneOnly(c));
            }
            if (toneCount > 1) return false;
            string bn = baseNuc.ToString();
            if (!ValidBaseNuclei.Contains(bn)) return false;

            // Strict Vietnamese stop-coda tone rule: -c, -ch, -p, -t MUST take sắc (1) or nặng (5)!
            if (coda == "c" || coda == "ch" || coda == "p" || coda == "t")
            {
                if (toneVal != 1 && toneVal != 5) return false;
            }

            // Rules for ng/ngh, g/gh, c/k before i/e/ê
            char nv = bn[0];
            bool isFrontVowel = (nv == 'e' || nv == 'ê' || nv == 'i');
            if (isFrontVowel && (initial == "ng" || initial == "c")) return false;
            if (!isFrontVowel && (initial == "ngh" || initial == "gh")) return false;

            return true;
        }

        private static string PreserveCase(string original, string replacement)
        {
            if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(replacement)) return replacement;
            if (original == original.ToUpper()) return replacement.ToUpper();
            if (char.IsUpper(original[0])) return char.ToUpper(replacement[0]) + replacement.Substring(1);
            return replacement;
        }

        private static int ScoreCandidate(string original, string cand, int deletedIndex, string prevWord, string nextWord)
        {
            if (string.IsNullOrEmpty(cand) || !IsValidVietnameseSyllable(cand)) return -1;
            string lc = cand.ToLowerInvariant();
            int score = 5;

            if (CommonVietnameseSyllables.Contains(lc) || VbtmKnowledgeHub.DictionaryWords.Contains(lc)) score += 45;

            int leftExt = VbtmKnowledgeHub.GetBigramScore(prevWord, lc);
            if (leftExt > 0) score += leftExt;
            else if (!string.IsNullOrEmpty(prevWord) && CommonVietnameseBigrams.Contains(prevWord.ToLowerInvariant() + " " + lc))
                score += 120;

            int rightExt = VbtmKnowledgeHub.GetBigramScore(lc, nextWord);
            if (rightExt > 0) score += rightExt;
            else if (!string.IsNullOrEmpty(nextWord) && CommonVietnameseBigrams.Contains(lc + " " + nextWord.ToLowerInvariant()))
                score += 120;

            if (deletedIndex >= 0 && deletedIndex < original.Length)
            {
                char removed = original[deletedIndex];
                // Ưu tiên giữ nguyên nguyên âm có dấu tiếng Việt
                if (removed <= 127) score += 15;
                char rl = char.ToLowerInvariant(removed);
                if ("sfrxjzw".IndexOf(rl) >= 0) score += 8;
                if ((deletedIndex > 0 && char.ToLowerInvariant(original[deletedIndex - 1]) == rl) ||
                    (deletedIndex + 1 < original.Length && char.ToLowerInvariant(original[deletedIndex + 1]) == rl))
                {
                    score += 10;
                }
            }

            return score;
        }

        public static string TryRepairTelexOrExtraChar(string word)
        {
            return TryRepairTelexOrExtraChar(word, null, null);
        }

        /// <summary>
        /// Universal high-precision Vietnamese typo & Telex repair engine:
        /// Catches extra consonant/vowel anywhere (0..N-1), stray Telex keys, transpositions, missing n/h, missing tone on -c/-ch/-p/-t, and 1-char substitutions.
        /// </summary>
        public static string TryRepairTelexOrExtraChar(string word, string prevWord, string nextWord)
        {
            if (string.IsNullOrWhiteSpace(word) || word.Length < 2 || word.Length > 16) return null;
            VbtmKnowledgeHub.EnsureLoaded();
            if (WhitelistWords.Contains(word) || VbtmKnowledgeHub.ProtectedWhitelist.Contains(word)) return null;

            string w = word.Normalize(NormalizationForm.FormC);
            // Bỏ qua từ viết tắt toàn chữ HOA ngắn không dấu (VD: UBND, HĐND, TW, QR)
            if (w.Length <= 6 && w == w.ToUpperInvariant() && !HasVietnameseDiacritic(w)) return null;

            string direct;
            if (DirectWordRepairs.TryGetValue(w, out direct) && !string.Equals(w, direct, StringComparison.OrdinalIgnoreCase))
            {
                return PreserveCase(w, direct);
            }
            if (VbtmKnowledgeHub.ExternalTypoMap.TryGetValue(w, out direct) && !string.Equals(w, direct, StringComparison.OrdinalIgnoreCase))
            {
                return PreserveCase(w, direct);
            }
            if (IsValidVietnameseSyllable(w)) return null;

            bool hasVn = HasVietnameseDiacritic(w);
            bool neighborHasVn = (!string.IsNullOrEmpty(prevWord) && HasVietnameseDiacritic(prevWord)) ||
                                 (!string.IsNullOrEmpty(nextWord) && HasVietnameseDiacritic(nextWord));

            // 0. Rút gọn các ký tự bị dính phím lặp liên tiếp 2+ lần (VD: "hoạchhh" -> "hoạch", "HOẠCHHH" -> "HOẠCH", "cônggg" -> "công", "quyếttt" -> "quyết")
            StringBuilder collapsedSb = new StringBuilder(w.Length);
            for (int i = 0; i < w.Length; i++)
            {
                if (i == 0 || char.ToLowerInvariant(w[i]) != char.ToLowerInvariant(w[i - 1]))
                {
                    collapsedSb.Append(w[i]);
                }
            }
            string collapsed = collapsedSb.ToString();
            if (collapsed.Length < w.Length)
            {
                int colScore = ScoreCandidate(w, collapsed, -1, prevWord, nextWord);
                if (colScore >= 50 || (hasVn && colScore > 0))
                {
                    return PreserveCase(w, collapsed);
                }
            }

            // Nếu là từ thuần ASCII không dấu và cả 2 từ kề đều không có dấu tiếng Việt -> chỉ bắt lỗi gõ lặp chữ cuối (vd: trongg, tronggg)
            if (!hasVn && !neighborHasVn)
            {
                return null;
            }

            string bestCand = null;
            int bestScore = -1;

            // 1. Thử xóa BẤT KỲ 1 ký tự thừa nào tại mọi vị trí 0..w.Length-1 (bắt: tập treung, cqác, vận độeng, tquyền, cơ scấu, đảng sviên, thàdnh, xunga, evà, điểsm, sắcs...)
            if (w.Length >= 3)
            {
                for (int i = 0; i < w.Length; i++)
                {
                    string cand = w.Remove(i, 1);
                    int sc = ScoreCandidate(w, cand, i, prevWord, nextWord);
                    // Với từ thuần ASCII (không dấu), yêu cầu ứng viên phải là âm tiết phổ biến hoặc khớp cụm từ tiếng Việt để tránh báo nhầm từ tiếng Anh
                    if (!hasVn && sc < 50) continue;
                    if (sc > bestScore)
                    {
                        bestScore = sc;
                        bestCand = cand;
                    }
                }
            }

            // 1b. Nếu từ có dấu tiếng Việt và vẫn chưa khớp -> thử xóa 2 ký tự thừa bất kỳ (VD: "hoạchhs" -> "hoạch", "thànhdf" -> "thành", "đượcjk" -> "được")
            if (bestCand == null && hasVn && w.Length >= 4 && w.Length <= 11)
            {
                for (int i = 0; i < w.Length; i++)
                {
                    if (w[i] > 127) continue;
                    string sub1 = w.Remove(i, 1);
                    for (int j = i; j < sub1.Length; j++)
                    {
                        if (sub1[j] > 127) continue;
                        string cand2 = sub1.Remove(j, 1);
                        int sc = ScoreCandidate(w, cand2, -1, prevWord, nextWord);
                        if (sc >= 50 && sc > bestScore)
                        {
                            bestScore = sc;
                            bestCand = cand2;
                        }
                    }
                }
            }

            // 2. Thử đảo vị trí 2 ký tự liền kề (lỗi gõ nhanh: thnhà -> thành, trugn -> trung, côgn -> công)
            if (w.Length >= 3)
            {
                char[] arr = w.ToCharArray();
                for (int i = 0; i < arr.Length - 1; i++)
                {
                    if (arr[i] == arr[i + 1]) continue;
                    char tmp = arr[i]; arr[i] = arr[i + 1]; arr[i + 1] = tmp;
                    string cand = new string(arr);
                    arr[i + 1] = arr[i]; arr[i] = tmp;

                    int sc = ScoreCandidate(w, cand, -1, prevWord, nextWord);
                    if (!hasVn && sc < 50) continue;
                    if (sc > bestScore)
                    {
                        bestScore = sc;
                        bestCand = cand;
                    }
                }
            }

            // 3. Thử chèn thiếu 'n' trước 'g' cuối (nhưg -> nhưng, trườg -> trường, côg -> công)
            if (w.EndsWith("g", StringComparison.OrdinalIgnoreCase) && w.Length >= 3)
            {
                string candNg = w.Substring(0, w.Length - 1) + "ng";
                int sc = ScoreCandidate(w, candNg, -1, prevWord, nextWord);
                if (sc > bestScore)
                {
                    bestScore = sc;
                    bestCand = candNg;
                }
            }

            // 4. Thiếu dấu sắc/nặng trên vần kết thúc bằng -c, -ch, -p, -t (thưc -> thực, hoach -> hoạch, quyêt -> quyết)
            string lw = w.ToLowerInvariant();
            if (lw.EndsWith("c") || lw.EndsWith("ch") || lw.EndsWith("p") || lw.EndsWith("t"))
            {
                string[,] toneAdd = new string[,] {
                    { "a", "á" }, { "a", "ạ" }, { "ă", "ắ" }, { "ă", "ặ" }, { "â", "ấ" }, { "â", "ậ" },
                    { "e", "é" }, { "e", "ẹ" }, { "ê", "ế" }, { "ê", "ệ" },
                    { "i", "í" }, { "i", "ị" }, { "o", "ó" }, { "o", "ọ" }, { "ô", "ố" }, { "ô", "ộ" },
                    { "ơ", "ớ" }, { "ơ", "ợ" }, { "u", "ú" }, { "u", "ụ" }, { "ư", "ứ" }, { "ư", "ự" }
                };
                for (int i = w.Length - 1; i >= 0; i--)
                {
                    if (IsVietnameseVowel(w[i]))
                    {
                        string ch = w[i].ToString().ToLowerInvariant();
                        for (int k = 0; k < toneAdd.GetLength(0); k++)
                        {
                            if (toneAdd[k, 0] == ch)
                            {
                                string cand = w.Substring(0, i) + toneAdd[k, 1] + w.Substring(i + 1);
                                int sc = ScoreCandidate(w, cand, -1, prevWord, nextWord);
                                if (!hasVn && sc < 50) continue;
                                if (sc > bestScore)
                                {
                                    bestScore = sc;
                                    bestCand = cand;
                                }
                            }
                        }
                        break;
                    }
                }
            }

            // 5. Nếu chưa tìm thấy bằng xóa/đảo và từ có dấu tiếng Việt -> thử thay 1 ký tự cuối/giữa (vd: cônh -> công, viêb -> viên)
            if (bestCand == null && hasVn && w.Length >= 3)
            {
                const string subChars = "acdeghiklmnopqrstuvxy";
                for (int i = w.Length - 1; i >= 0; i--)
                {
                    if (w[i] > 127) continue; // giữ nguyên nguyên âm có dấu
                    for (int k = 0; k < subChars.Length; k++)
                    {
                        if (char.ToLowerInvariant(w[i]) == subChars[k]) continue;
                        string cand = w.Substring(0, i) + subChars[k] + w.Substring(i + 1);
                        int sc = ScoreCandidate(w, cand, -1, prevWord, nextWord);
                        if (sc >= 50 && sc > bestScore)
                        {
                            bestScore = sc;
                            bestCand = cand;
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(bestCand))
            {
                return PreserveCase(w, bestCand);
            }

            return null;
        }

        public static List<SpellCardIssue> ScanDocumentFast(string docText)
        {
            List<SpellCardIssue> results = new List<SpellCardIssue>();
            if (string.IsNullOrWhiteSpace(docText)) return results;

            string text = docText.Normalize(NormalizationForm.FormC);
            HashSet<string> goodWords = new HashSet<string>(VbtmCustomDictStore.LoadGoodWords(), StringComparer.OrdinalIgnoreCase);

            // Track (position, wrong) to allow multiple occurrences like 2x "năm giữ" just like the user's screenshot!
            List<KeyValuePair<int, SpellCardIssue>> ordered = new List<KeyValuePair<int, SpellCardIssue>>();
            HashSet<int> coveredPositions = new HashSet<int>();

            // 0. User-defined fixed bad words (Từ sai cố định)
            foreach (KeyValuePair<string, string> kv in VbtmCustomDictStore.LoadBadWords())
            {
                if (goodWords.Contains(kv.Key)) continue;
                foreach (Match m in Regex.Matches(text, @"(?<![A-Za-zÀ-Ỹà-ỹĐđ])" + Regex.Escape(kv.Key) + @"(?![A-Za-zÀ-Ỹà-ỹĐđ])", RegexOptions.IgnoreCase))
                {
                    ordered.Add(new KeyValuePair<int, SpellCardIssue>(m.Index, new SpellCardIssue
                    {
                        Wrong = m.Value,
                        Suggest = PreserveCase(m.Value, kv.Value),
                        Badge = "Chính tả",
                        Reason = "Từ sai cố định trong từ điển đơn vị",
                        IsSuspicious = false
                    }));
                    for (int p = m.Index; p < m.Index + m.Length; p++) coveredPositions.Add(p);
                }
            }

            // 1. Scan Telex / extra-character / transposition typos with smart 2-way context (left & right neighbor)
            MatchCollection wordMatches = Regex.Matches(text, @"[A-Za-zÀ-Ỹà-ỹĐđ]+");
            for (int i = 0; i < wordMatches.Count; i++)
            {
                Match wm = wordMatches[i];
                if (coveredPositions.Contains(wm.Index)) continue;
                string tok = wm.Value;
                if (goodWords.Contains(tok) || WhitelistWords.Contains(tok)) continue;

                string prevWord = null;
                bool hasLeftSpace = false;
                if (i > 0)
                {
                    Match prev = wordMatches[i - 1];
                    int gapStart = prev.Index + prev.Length;
                    if (wm.Index - gapStart == 1 && text[gapStart] == ' ')
                    {
                        prevWord = prev.Value;
                        hasLeftSpace = true;
                    }
                }

                string nextWord = null;
                bool hasRightSpace = false;
                if (i + 1 < wordMatches.Count)
                {
                    Match next = wordMatches[i + 1];
                    int gapStart = wm.Index + wm.Length;
                    if (next.Index - gapStart == 1 && text[gapStart] == ' ')
                    {
                        nextWord = next.Value;
                        hasRightSpace = true;
                    }
                }

                string repaired = TryRepairTelexOrExtraChar(tok, prevWord, nextWord);
                if (!string.IsNullOrEmpty(repaired) && !string.Equals(tok, repaired, StringComparison.Ordinal))
                {
                    string displayWrong = tok;
                    string displaySug = repaired;
                    int spanStart = wm.Index;
                    int spanLen = wm.Length;

                    bool leftBigram = hasLeftSpace && (CommonVietnameseBigrams.Contains(prevWord.ToLowerInvariant() + " " + repaired.ToLowerInvariant()) || VbtmKnowledgeHub.GetBigramScore(prevWord, repaired) > 0);
                    bool rightBigram = hasRightSpace && (CommonVietnameseBigrams.Contains(repaired.ToLowerInvariant() + " " + nextWord.ToLowerInvariant()) || VbtmKnowledgeHub.GetBigramScore(repaired, nextWord) > 0);

                    if (leftBigram)
                    {
                        Match prev = wordMatches[i - 1];
                        displayWrong = prev.Value + " " + tok;
                        displaySug = prev.Value + " " + repaired;
                        spanStart = prev.Index;
                        spanLen = (wm.Index + wm.Length) - prev.Index;
                    }
                    else if (rightBigram)
                    {
                        Match next = wordMatches[i + 1];
                        displayWrong = tok + " " + next.Value;
                        displaySug = repaired + " " + next.Value;
                        spanStart = wm.Index;
                        spanLen = (next.Index + next.Length) - wm.Index;
                    }
                    else if (hasLeftSpace && repaired.Length <= 2)
                    {
                        // Cụm ngắn như "nước evà -> nước và" giữ thêm từ trước để dễ xác định vị trí trong văn bản
                        Match prev = wordMatches[i - 1];
                        displayWrong = prev.Value + " " + tok;
                        displaySug = prev.Value + " " + repaired;
                        spanStart = prev.Index;
                        spanLen = (wm.Index + wm.Length) - prev.Index;
                    }

                    if (!goodWords.Contains(displayWrong))
                    {
                        ordered.Add(new KeyValuePair<int, SpellCardIssue>(spanStart, new SpellCardIssue
                        {
                            Wrong = displayWrong,
                            Suggest = displaySug,
                            Badge = "Chính tả",
                            Reason = "Lỗi gõ dư phím / sai âm tiết tiếng Việt",
                            IsSuspicious = false
                        }));
                        for (int p = spanStart; p < spanStart + spanLen; p++) coveredPositions.Add(p);
                    }
                }
            }

            // 2. Scan Context & Common Confusable Phrases (Nghi ngờ / Chính tả)
            for (int r = 0; r < ContextAndTypoPairs.GetLength(0); r++)
            {
                string w = ContextAndTypoPairs[r, 0];
                string s = ContextAndTypoPairs[r, 1];
                string badge = ContextAndTypoPairs[r, 2];
                string reason = ContextAndTypoPairs[r, 3];
                if (string.IsNullOrEmpty(reason) || string.Equals(w, s, StringComparison.OrdinalIgnoreCase)) continue;
                if (goodWords.Contains(w)) continue;

                foreach (Match m in Regex.Matches(text, @"(?<![A-Za-zÀ-Ỹà-ỹĐđ])" + Regex.Escape(w) + @"(?![A-Za-zÀ-Ỹà-ỹĐđ])", RegexOptions.IgnoreCase))
                {
                    if (coveredPositions.Contains(m.Index)) continue;
                    if (goodWords.Contains(m.Value)) continue;
                    string prop = PreserveCase(m.Value, s);
                    if (string.Equals(m.Value, prop, StringComparison.Ordinal)) continue;

                    ordered.Add(new KeyValuePair<int, SpellCardIssue>(m.Index, new SpellCardIssue
                    {
                        Wrong = m.Value,
                        Suggest = prop,
                        Badge = badge,
                        Reason = reason,
                        IsSuspicious = string.Equals(badge, "Nghi ngờ", StringComparison.OrdinalIgnoreCase)
                    }));
                    for (int p = m.Index; p < m.Index + m.Length; p++) coveredPositions.Add(p);
                }
            }

            // 2b. Scan ALL 150+ Administrative & Party rules from WordSpellInspectorForm.InspectorRules
            for (int r = 0; r < WordSpellInspectorForm.InspectorRules.GetLength(0); r++)
            {
                string w = WordSpellInspectorForm.InspectorRules[r, 0];
                string s = WordSpellInspectorForm.InspectorRules[r, 1];
                string reason = WordSpellInspectorForm.InspectorRules[r, 2];
                if (string.IsNullOrEmpty(reason) || string.Equals(w, s, StringComparison.OrdinalIgnoreCase)) continue;
                if (goodWords.Contains(w)) continue;

                foreach (Match m in Regex.Matches(text, @"(?<![A-Za-zÀ-Ỹà-ỹĐđ])" + Regex.Escape(w) + @"(?![A-Za-zÀ-Ỹà-ỹĐđ])", RegexOptions.IgnoreCase))
                {
                    if (coveredPositions.Contains(m.Index)) continue;
                    if (goodWords.Contains(m.Value)) continue;
                    string prop = PreserveCase(m.Value, s);
                    if (string.Equals(m.Value, prop, StringComparison.Ordinal)) continue;

                    ordered.Add(new KeyValuePair<int, SpellCardIssue>(m.Index, new SpellCardIssue
                    {
                        Wrong = m.Value,
                        Suggest = prop,
                        Badge = "Chính tả",
                        Reason = reason,
                        IsSuspicious = false
                    }));
                    for (int p = m.Index; p < m.Index + m.Length; p++) coveredPositions.Add(p);
                }
            }

            // 2c. Scan 350+ Confusable Vietnamese Phrases from SpellCheck\Data\confusable_phrases.json
            VbtmKnowledgeHub.EnsureLoaded();
            string lowerText = text.ToLowerInvariant();
            foreach (ConfusablePhraseRule cr in VbtmKnowledgeHub.ConfusableRules)
            {
                if (goodWords.Contains(cr.Wrong)) continue;
                // Fast pre-filter before regex so scanning 350+ phrases takes < 5ms!
                if (lowerText.IndexOf(cr.Wrong.ToLowerInvariant(), StringComparison.Ordinal) < 0) continue;

                foreach (Match m in Regex.Matches(text, @"(?<![A-Za-zÀ-Ỹà-ỹĐđ])" + Regex.Escape(cr.Wrong) + @"(?![A-Za-zÀ-Ỹà-ỹĐđ])", RegexOptions.IgnoreCase))
                {
                    if (coveredPositions.Contains(m.Index)) continue;
                    if (goodWords.Contains(m.Value)) continue;
                    string prop = PreserveCase(m.Value, cr.Replacement);
                    if (string.Equals(m.Value, prop, StringComparison.Ordinal)) continue;

                    ordered.Add(new KeyValuePair<int, SpellCardIssue>(m.Index, new SpellCardIssue
                    {
                        Wrong = m.Value,
                        Suggest = prop,
                        Badge = cr.IsSuspicious ? "Nghi ngờ" : "Chính tả",
                        Reason = cr.Description,
                        IsSuspicious = cr.IsSuspicious
                    }));
                    for (int p = m.Index; p < m.Index + m.Length; p++) coveredPositions.Add(p);
                }
            }

            ordered.Sort((a, b) => a.Key.CompareTo(b.Key));
            Dictionary<string, int> occCounter = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var kv in ordered)
            {
                SpellCardIssue iss = kv.Value;
                int occ = occCounter.ContainsKey(iss.Wrong) ? occCounter[iss.Wrong] + 1 : 1;
                occCounter[iss.Wrong] = occ;
                iss.OccurrenceIndex = occ;
                results.Add(iss);
            }
            return results;
        }
    }

    // =========================================================================
    // MODERN RIGHT-DOCKED TASKPANE SPELL INSPECTOR (MATCHES USER SCREENSHOT)
    // =========================================================================
    public class WordSpellTaskPaneForm : Form
    {
        private Panel pnlTabBar, pnlSpellView, pnlFormatView, pnlCardsHost;
        private Button btnTabFormat, btnTabSpell, btnApplyAll, btnRescan, btnAiDeep;
        private Label lblFoundCount, lblStatusBottom;
        private List<SpellCardIssue> activeIssues = new List<SpellCardIssue>();
        private string currentDocText = "";

        public WordSpellTaskPaneForm()
        {
            this.Text = "VBTM Pro 3.6 — Kiểm tra Chính tả AI";
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int w = 456;
            this.Size = new Size(w, wa.Height - 12);
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(wa.Right - w, wa.Top + 6);
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.2f);

            // 1. Top Header Bar
            Panel hdr = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.White };
            hdr.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawLine(p, 0, hdr.Height - 1, hdr.Width, hdr.Height - 1); };
            hdr.Controls.Add(new Label { Text = "▌ Chuẩn hoá VB", Location = new Point(12, 13), AutoSize = true, ForeColor = Color.FromArgb(15, 23, 42), Font = new Font("Segoe UI", 11f, FontStyle.Bold) });

            Button btnGuide = MakePillBtn("? Hướng dẫn", 158, 10, 102, 28, Color.FromArgb(239, 246, 255), Color.FromArgb(29, 78, 216), Color.FromArgb(191, 219, 254));
            btnGuide.Click += (s, e) => { using (var f = new WordShortcutsForm()) f.ShowDialog(); };
            Button btnInfo = MakePillBtn("● Thông tin PM", 268, 10, 152, 28, Color.FromArgb(240, 253, 244), Color.FromArgb(21, 128, 61), Color.FromArgb(187, 247, 208));
            btnInfo.Click += (s, e) => WordQuickUtilities.ShowAboutDialog();
            hdr.Controls.AddRange(new Control[] { btnGuide, btnInfo });
            this.Controls.Add(hdr);

            // 2. Segmented Tab Bar: [ Chuẩn Hóa & Mẫu ] | [ Chính Tả AI ]
            pnlTabBar = new Panel { Dock = DockStyle.Top, Height = 50, BackColor = Color.White, Padding = new Padding(14, 7, 14, 7) };
            Panel segBg = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(241, 245, 249), Padding = new Padding(3) };
            btnTabFormat = new Button { Text = "Chuẩn Hóa & Mẫu", Dock = DockStyle.Left, Width = 200, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnTabFormat.FlatAppearance.BorderSize = 0;
            btnTabSpell = new Button { Text = "Chính Tả AI", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnTabSpell.FlatAppearance.BorderSize = 0;
            btnTabFormat.Click += (s, e) => SwitchTab(false);
            btnTabSpell.Click += (s, e) => SwitchTab(true);
            segBg.Controls.Add(btnTabSpell);
            segBg.Controls.Add(btnTabFormat);
            pnlTabBar.Controls.Add(segBg);
            this.Controls.Add(pnlTabBar);

            // 3. Bottom Action Bar
            Panel ftr = new Panel { Dock = DockStyle.Bottom, Height = 64, BackColor = Color.White };
            ftr.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawLine(p, 0, 0, ftr.Width, 0); };
            lblStatusBottom = new Label { Text = "Sẵn sàng rà soát chính tả Offline 100%", Location = new Point(14, 5), Size = new Size(416, 18), ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.2f) };
            btnRescan = new Button { Text = "🔍 Quét lại văn bản", Location = new Point(14, 26), Size = new Size(198, 32), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnRescan.FlatAppearance.BorderSize = 0;
            btnRescan.Click += (s, e) => ScanActiveWordInstant();

            btnAiDeep = new Button { Text = "⚡ Quét AI (CT2 / Ollama)", Location = new Point(222, 26), Size = new Size(206, 32), BackColor = Color.FromArgb(241, 245, 249), ForeColor = Color.FromArgb(30, 41, 59), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnAiDeep.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnAiDeep.Click += (s, e) => RunDeepOllamaSupplementAsync();
            ftr.Controls.AddRange(new Control[] { lblStatusBottom, btnRescan, btnAiDeep });
            this.Controls.Add(ftr);

            // 4. Spell Check Panel
            pnlSpellView = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 250, 252) };

            Panel pnlSpellTop = new Panel { Dock = DockStyle.Top, Height = 128, BackColor = Color.White };
            pnlSpellTop.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawLine(p, 0, pnlSpellTop.Height - 1, pnlSpellTop.Width, pnlSpellTop.Height - 1); };
            pnlSpellTop.Controls.Add(new Label { Text = "▌ KIỂM TRA CHÍNH TẢ AI", Location = new Point(14, 10), AutoSize = true, ForeColor = Color.FromArgb(15, 23, 42), Font = new Font("Segoe UI", 10f, FontStyle.Bold) });
            Label lblOfflineBadge = new Label { Text = "● Offline 100%", Location = new Point(324, 9), Size = new Size(102, 22), BackColor = Color.FromArgb(240, 253, 244), ForeColor = Color.FromArgb(21, 128, 61), TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 8.2f, FontStyle.Bold) };
            pnlSpellTop.Controls.Add(lblOfflineBadge);

            Button btnClearMark = MakePillBtn("✓ Xóa dấu xanh", 14, 40, 128, 30, Color.White, Color.FromArgb(29, 78, 216), Color.FromArgb(203, 213, 225));
            btnClearMark.Click += (s, e) => ClearWordHighlights();
            Button btnAddGood = MakePillBtn("+ Từ đúng", 150, 40, 130, 30, Color.White, Color.FromArgb(51, 65, 85), Color.FromArgb(203, 213, 225));
            btnAddGood.Click += (s, e) => { using (var d = new SpellDictionaryManagerForm("add-good")) d.ShowDialog(); ScanActiveWordInstant(); };
            Button btnDict = MakePillBtn("Từ điển...", 288, 40, 138, 30, Color.White, Color.FromArgb(51, 65, 85), Color.FromArgb(203, 213, 225));
            btnDict.Click += (s, e) => { using (var d = new SpellDictionaryManagerForm("manage")) d.ShowDialog(); ScanActiveWordInstant(); };
            pnlSpellTop.Controls.AddRange(new Control[] { btnClearMark, btnAddGood, btnDict });

            lblFoundCount = new Label { Text = "PHÁT HIỆN 0 TỪ CẦN SỬA", Location = new Point(14, 88), AutoSize = true, ForeColor = Color.FromArgb(15, 23, 42), Font = new Font("Segoe UI", 9.8f, FontStyle.Bold) };
            btnApplyAll = new Button { Text = "✓ Áp dụng tất cả", Location = new Point(282, 82), Size = new Size(144, 34), BackColor = Color.FromArgb(16, 185, 129), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnApplyAll.FlatAppearance.BorderSize = 0;
            btnApplyAll.Click += (s, e) => ApplyAllFixesInWord();
            pnlSpellTop.Controls.AddRange(new Control[] { lblFoundCount, btnApplyAll });

            pnlCardsHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(248, 250, 252), Padding = new Padding(12, 8, 12, 12) };
            pnlSpellView.Controls.Add(pnlCardsHost);
            pnlSpellView.Controls.Add(pnlSpellTop);

            // 5. Quick Format & Templates Panel (Tab 1)
            pnlFormatView = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(248, 250, 252), Padding = new Padding(14) };
            BuildQuickFormatTab();

            this.Controls.Add(pnlSpellView);
            this.Controls.Add(pnlFormatView);
            pnlSpellView.BringToFront();
            hdr.SendToBack();
            pnlTabBar.SendToBack();
            ftr.SendToBack();
            pnlSpellView.BringToFront();

            SwitchTab(true);
            this.Shown += (s, e) =>
            {
                ScanActiveWordInstant();
                // Tự động chạy quét sâu AI Local toàn bộ văn bản ở nền (không chặn giao diện)
                if (!string.IsNullOrWhiteSpace(currentDocText))
                {
                    RunDeepOllamaSupplementAsync();
                }
            };
        }

        private Button MakePillBtn(string text, int x, int y, int w, int h, Color bg, Color fg, Color border)
        {
            Button b = new Button { Text = text, Location = new Point(x, y), Size = new Size(w, h), BackColor = bg, ForeColor = fg, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderColor = border;
            return b;
        }

        private void SwitchTab(bool showSpell)
        {
            btnTabSpell.BackColor = showSpell ? Color.FromArgb(29, 78, 216) : Color.Transparent;
            btnTabSpell.ForeColor = showSpell ? Color.White : Color.FromArgb(71, 85, 105);
            btnTabFormat.BackColor = !showSpell ? Color.FromArgb(29, 78, 216) : Color.Transparent;
            btnTabFormat.ForeColor = !showSpell ? Color.White : Color.FromArgb(71, 85, 105);
            pnlSpellView.Visible = showSpell;
            pnlFormatView.Visible = !showSpell;
        }

        private void BuildQuickFormatTab()
        {
            VbtmKnowledgeHub.EnsureLoaded();
            int y = 12;
            int tplCount = VbtmKnowledgeHub.OfficialTemplates.Count;
            AddQuickActionCard(pnlFormatView,
                "📚 Kho " + (tplCount > 0 ? tplCount.ToString() : "84") + " Mẫu Văn Bản Chuẩn (Đảng · NĐ 30 · CAND)",
                "Mở nhanh mẫu Nghị quyết, Quyết định, Công văn, Báo cáo, CAND, Cá nhân...",
                y,
                () => { using (var f = new OfficialTemplateBrowserForm()) f.ShowDialog(); });
            y += 66;

            // Thanh chèn nhanh Dấu Mật / Khẩn theo NĐ 30 & CAND
            Panel pnlMarks = new Panel { Location = new Point(14, y), Size = new Size(410, 56), BackColor = Color.White };
            pnlMarks.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(p, 0, 0, pnlMarks.Width - 1, pnlMarks.Height - 1); };
            pnlMarks.Controls.Add(new Label { Text = "Chèn nhanh dấu Độ khẩn / Độ mật:", Location = new Point(10, 5), AutoSize = true, Font = new Font("Segoe UI", 8.2f, FontStyle.Bold), ForeColor = Color.FromArgb(71, 85, 105) });
            string[] marks = new string[] { "HỎA TỐC", "MẬT", "TỐI MẬT", "TUYỆT MẬT" };
            int mx = 10;
            foreach (string mk in marks)
            {
                string capturedMk = mk;
                bool isRed = (capturedMk == "HỎA TỐC");
                Button bm = MakePillBtn(capturedMk, mx, 25, 92, 24,
                    isRed ? Color.FromArgb(254, 242, 242) : Color.FromArgb(241, 245, 249),
                    isRed ? Color.FromArgb(185, 28, 28) : Color.FromArgb(30, 41, 59),
                    isRed ? Color.FromArgb(254, 202, 202) : Color.FromArgb(203, 213, 225));
                bm.Click += (s, e) => VbtmKnowledgeHub.InsertSpecialMarkIntoWord(capturedMk, isRed);
                pnlMarks.Controls.Add(bm);
                mx += 98;
            }
            pnlFormatView.Controls.Add(pnlMarks);
            y += 64;

            AddQuickActionCard(pnlFormatView, "📐 Thiết lập Căn lề & Khổ giấy (Trực quan)", "Mở giao diện căn lề A4 + xem trước trang in (NĐ 30 / HD 05)", y, () => { using (var f = new PageSetupVisualForm()) f.ShowDialog(); }); y += 66;
            AddQuickActionCard(pnlFormatView, "⚙ Cài đặt Hệ thống & Mẫu Thể thức", "Quản lý quy chuẩn NĐ 30, 05-HD/VPTW, CAND, Quốc hiệu, Chữ ký", y, () => { using (var f = new SystemAndTemplateSettingsForm(0)) f.ShowDialog(); }); y += 66;
            AddQuickActionCard(pnlFormatView, "💰 Đọc tiền sang chữ trong Word", "Đổi số tiền đang bôi đen thành chữ tiếng Việt chuẩn hành chính", y, () => WordQuickUtilities.ConvertSelectedMoneyToVietnameseWords()); y += 66;
            AddQuickActionCard(pnlFormatView, "💬 Xóa toàn bộ nhận xét (Comments)", "Dọn sạch toàn bộ ghi chú nhận xét trước khi phát hành văn bản", y, () => WordQuickUtilities.DeleteWordComments(true)); y += 66;
            AddQuickActionCard(pnlFormatView, "📱 Tạo Mã QR Code văn bản", "Chèn mã QR chuẩn ISO 18004 trực tiếp vào tài liệu Word", y, () => { using (var f = new WordQrInsertForm()) f.ShowDialog(); }); y += 66;
            AddQuickActionCard(pnlFormatView, "⌨️ Các phím tắt trong Word", "Tra cứu nhanh 32 phím tắt Word có highlight trực tiếp", y, () => { using (var f = new WordShortcutsForm()) f.ShowDialog(); });
        }

        private void AddQuickActionCard(Panel parent, string title, string sub, int y, Action onClick)
        {
            Panel c = new Panel { Location = new Point(14, y), Size = new Size(410, 58), BackColor = Color.White, Cursor = Cursors.Hand };
            c.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(p, 0, 0, c.Width - 1, c.Height - 1); };
            Label l1 = new Label { Text = title, Location = new Point(12, 8), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(29, 78, 216), Cursor = Cursors.Hand };
            Label l2 = new Label { Text = sub, Location = new Point(12, 31), AutoSize = true, Font = new Font("Segoe UI", 8.2f), ForeColor = Color.FromArgb(100, 116, 139), Cursor = Cursors.Hand };
            EventHandler click = (s, e) => onClick();
            c.Click += click; l1.Click += click; l2.Click += click;
            c.Controls.AddRange(new Control[] { l1, l2 });
            parent.Controls.Add(c);
        }

        private string ReadActiveWordText()
        {
            string docText = "";
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                docText = Convert.ToString(content.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, content, null)) ?? "";
            }
            catch { }
            if (string.IsNullOrWhiteSpace(docText))
            {
                try
                {
                    string snapPath = Path.Combine(Path.GetTempPath(), "vbtm_word_snapshot.txt");
                    if (File.Exists(snapPath)) docText = File.ReadAllText(snapPath, Encoding.Unicode);
                }
                catch { }
            }
            return docText;
        }

        private void ScanActiveWordInstant()
        {
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            currentDocText = ReadActiveWordText();
            activeIssues = OfflineVietnameseSpellEngine.ScanDocumentFast(currentDocText);
            sw.Stop();
            RenderCards();
            lblStatusBottom.Text = string.Format("✓ Quét Offline hoàn tất trong {0}ms • Phát hiện {1} từ cần sửa", sw.ElapsedMilliseconds, activeIssues.Count);
        }

        private void RenderCards()
        {
            pnlCardsHost.SuspendLayout();
            pnlCardsHost.Controls.Clear();
            lblFoundCount.Text = "PHÁT HIỆN " + activeIssues.Count + " TỪ CẦN SỬA";

            if (activeIssues.Count == 0)
            {
                Label ok = new Label
                {
                    Text = "✓ Không phát hiện lỗi chính tả nào trong tài liệu đang mở.\r\n\r\nBạn có thể bấm '⚡ Quét sâu AI (Ollama)' ở bên dưới để AI rà soát ngữ cảnh chuyên sâu.",
                    Location = new Point(16, 24),
                    Size = new Size(394, 90),
                    ForeColor = Color.FromArgb(21, 128, 61),
                    Font = new Font("Segoe UI", 9.5f)
                };
                pnlCardsHost.Controls.Add(ok);
                pnlCardsHost.ResumeLayout();
                return;
            }

            int y = 8;
            foreach (SpellCardIssue issue in activeIssues)
            {
                SpellCardIssue captured = issue;
                Panel card = new Panel { Location = new Point(10, y), Size = new Size(412, 52), BackColor = Color.White };
                card.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(p, 0, 0, card.Width - 1, card.Height - 1); };

                // Red wrong pill
                Label lblWrong = new Label
                {
                    Text = captured.Wrong,
                    Location = new Point(10, 12),
                    AutoSize = true,
                    Padding = new Padding(6, 3, 6, 3),
                    BackColor = Color.FromArgb(254, 242, 242),
                    ForeColor = Color.FromArgb(185, 28, 28),
                    Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                lblWrong.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(252, 165, 165))) e.Graphics.DrawRectangle(p, 0, 0, lblWrong.Width - 1, lblWrong.Height - 1); };
                lblWrong.Click += (s, e) => JumpToWordInDocument(captured.Wrong, captured.OccurrenceIndex);

                card.Controls.Add(lblWrong);
                int curX = lblWrong.Right + 6;

                Label lblArrow = new Label { Text = "→", Location = new Point(curX, 15), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
                card.Controls.Add(lblArrow);
                curX = lblArrow.Right + 6;

                // Green suggestion pill (click to apply fix!)
                Label lblSug = new Label
                {
                    Text = captured.Suggest,
                    Location = new Point(curX, 12),
                    AutoSize = true,
                    Padding = new Padding(6, 3, 6, 3),
                    BackColor = Color.FromArgb(240, 253, 244),
                    ForeColor = Color.FromArgb(21, 128, 61),
                    Font = new Font("Segoe UI", 9.2f, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                lblSug.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(134, 239, 172))) e.Graphics.DrawRectangle(p, 0, 0, lblSug.Width - 1, lblSug.Height - 1); };
                lblSug.Click += (s, e) => ApplySingleFixInWord(captured);
                card.Controls.Add(lblSug);

                // Category badge ("Chính tả" / "Nghi ngờ")
                Label lblBadge = new Label
                {
                    Text = captured.Badge,
                    Location = new Point(Math.Min(lblSug.Right + 6, 255), 16),
                    AutoSize = true,
                    Padding = new Padding(4, 1, 4, 1),
                    BackColor = captured.IsSuspicious ? Color.FromArgb(254, 249, 195) : Color.FromArgb(254, 226, 226),
                    ForeColor = captured.IsSuspicious ? Color.FromArgb(146, 64, 14) : Color.FromArgb(185, 28, 28),
                    Font = new Font("Segoe UI", 7.6f, FontStyle.Bold)
                };
                card.Controls.Add(lblBadge);

                // Right links: "+ Đúng" · "Bỏ qua"
                LinkLabel lnkGood = new LinkLabel { Text = "+ Đúng", Location = new Point(312, 16), AutoSize = true, LinkColor = Color.FromArgb(37, 99, 235), LinkBehavior = LinkBehavior.HoverUnderline, Font = new Font("Segoe UI", 8.2f, FontStyle.Bold) };
                lnkGood.Click += (s, e) =>
                {
                    VbtmCustomDictStore.AddGoodWord(captured.Wrong);
                    activeIssues.RemoveAll(x => string.Equals(x.Wrong, captured.Wrong, StringComparison.OrdinalIgnoreCase));
                    RenderCards();
                    lblStatusBottom.Text = "✓ Đã thêm \"" + captured.Wrong + "\" vào từ điển từ đúng.";
                };

                LinkLabel lnkSkip = new LinkLabel { Text = "Bỏ qua", Location = new Point(362, 16), AutoSize = true, LinkColor = Color.FromArgb(100, 116, 139), LinkBehavior = LinkBehavior.HoverUnderline, Font = new Font("Segoe UI", 8.2f) };
                lnkSkip.Click += (s, e) =>
                {
                    activeIssues.Remove(captured);
                    RenderCards();
                };

                card.Controls.Add(lnkGood);
                card.Controls.Add(lnkSkip);
                pnlCardsHost.Controls.Add(card);
                y += 60;
            }
            pnlCardsHost.ResumeLayout();
        }

        private void JumpToWordInDocument(string target, int occurrence)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                sel.GetType().InvokeMember("HomeKey", BindingFlags.InvokeMethod, null, sel, new object[] { 6 }); // wdStory
                object find = sel.GetType().InvokeMember("Find", BindingFlags.GetProperty, null, sel, null);
                find.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, find, null);
                find.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, find, new object[] { target });
                find.GetType().InvokeMember("MatchCase", BindingFlags.SetProperty, null, find, new object[] { false });
                find.GetType().InvokeMember("Forward", BindingFlags.SetProperty, null, find, new object[] { true });
                find.GetType().InvokeMember("Wrap", BindingFlags.SetProperty, null, find, new object[] { 1 }); // wdFindContinue
                int hits = Math.Max(1, occurrence);
                for (int i = 0; i < hits; i++)
                {
                    bool found = (bool)find.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, find, null);
                    if (!found) break;
                }
                wordApp.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, wordApp, null);
            }
            catch { }
        }

        private bool ReplaceOneInWord(string wrong, string suggest)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                object find = content.GetType().InvokeMember("Find", BindingFlags.GetProperty, null, content, null);
                find.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, find, null);
                object repl = find.GetType().InvokeMember("Replacement", BindingFlags.GetProperty, null, find, null);
                repl.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, repl, null);

                // Replace 1 occurrence (wdReplaceOne = 1)
                object[] args = new object[] {
                    wrong, true, false, false, false, false, true, 1, false, suggest, 1
                };
                return (bool)find.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, find, args);
            }
            catch { return false; }
        }

        private void ApplySingleFixInWord(SpellCardIssue issue)
        {
            bool ok = ReplaceOneInWord(issue.Wrong, issue.Suggest);
            activeIssues.Remove(issue);
            RenderCards();
            lblStatusBottom.Text = ok
                ? ("✓ Đã sửa \"" + issue.Wrong + "\" → \"" + issue.Suggest + "\" trong Word.")
                : ("Không tìm thấy \"" + issue.Wrong + "\" trong Word đang mở.");
        }

        private void ApplyAllFixesInWord()
        {
            if (activeIssues.Count == 0) return;
            int fixedCount = 0;
            List<SpellCardIssue> copy = new List<SpellCardIssue>(activeIssues);
            foreach (var iss in copy)
            {
                if (ReplaceOneInWord(iss.Wrong, iss.Suggest)) fixedCount++;
            }
            activeIssues.Clear();
            RenderCards();
            lblStatusBottom.Text = "✓ Đã áp dụng sửa tự động " + fixedCount + " lỗi trong tài liệu Word!";
        }

        private void ClearWordHighlights()
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                content.GetType().InvokeMember("HighlightColorIndex", BindingFlags.SetProperty, null, content, new object[] { 0 }); // wdNoHighlight
                lblStatusBottom.Text = "✓ Đã xóa toàn bộ dấu đánh dấu trong tài liệu Word.";
            }
            catch (Exception ex)
            {
                lblStatusBottom.Text = "Không thể xóa dấu: " + ex.Message;
            }
        }

        private void RunDeepOllamaSupplementAsync()
        {
            if (string.IsNullOrWhiteSpace(currentDocText))
            {
                ScanActiveWordInstant();
                if (string.IsNullOrWhiteSpace(currentDocText)) return;
            }
            btnAiDeep.Enabled = false;
            VbtmKnowledgeHub.EnsureLoaded();
            lblStatusBottom.Text = VbtmKnowledgeHub.HasCTranslate2
                ? "⏳ Đang nạp mô hình AI CTranslate2 Offline & soát ngữ cảnh..."
                : "⏳ AI Local đang tự động quét sâu toàn bộ văn bản...";
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    int totalAdded = 0;

                    // 1. Ưu tiên chạy mô hình AI CTranslate2 Offline nếu có sẵn trên máy
                    if (VbtmKnowledgeHub.HasCTranslate2)
                    {
                        List<SpellCardIssue> ct2Found = VbtmKnowledgeHub.RunCTranslate2DeepScan(currentDocText);
                        int ct2Added = 0;
                        foreach (SpellCardIssue ci in ct2Found)
                        {
                            if (VbtmCustomDictStore.IsGoodWord(ci.Wrong)) continue;
                            bool exists = false;
                            foreach (var ex in activeIssues)
                            {
                                if (string.Equals(ex.Wrong, ci.Wrong, StringComparison.OrdinalIgnoreCase) ||
                                    ex.Wrong.IndexOf(ci.Wrong, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    exists = true;
                                    break;
                                }
                            }
                            if (!exists)
                            {
                                ci.OccurrenceIndex = 1;
                                activeIssues.Add(ci);
                                ct2Added++;
                                totalAdded++;
                            }
                        }
                        if (ct2Added > 0)
                        {
                            try { this.BeginInvoke((Action)(() => RenderCards())); } catch { }
                        }
                    }

                    // 2. Bổ sung quét sâu ngữ cảnh bằng Ollama Local (nếu đang bật)
                    List<string> chunks = new List<string>();
                    string[] paras = currentDocText.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    StringBuilder cur = new StringBuilder();
                    foreach (string p in paras)
                    {
                        string t = p.Trim();
                        if (t.Length < 4) continue;
                        if (cur.Length + t.Length > 1800 && cur.Length > 0)
                        {
                            chunks.Add(cur.ToString());
                            cur.Length = 0;
                        }
                        cur.AppendLine(t);
                    }
                    if (cur.Length > 0) chunks.Add(cur.ToString());
                    if (chunks.Count > 15) chunks = chunks.GetRange(0, 15);

                    for (int ci = 0; ci < chunks.Count; ci++)
                    {
                        int curIdx = ci + 1;
                        int totalCh = chunks.Count;
                        try
                        {
                            this.BeginInvoke((Action)(() =>
                            {
                                lblStatusBottom.Text = string.Format("⏳ AI Local đang quét sâu đoạn {0}/{1}...", curIdx, totalCh);
                            }));
                        }
                        catch { }

                        string prompt = "Chỉ liệt kê lỗi chính tả, sai dấu, sai thanh điệu hoặc nhầm từ ngữ cảnh tiếng Việt thật sự sai trong đoạn văn sau dưới dạng JSON array: [{\"original\":\"cụm sai\",\"suggestion\":\"cụm đúng\",\"reason\":\"ngắn gọn\"}]. Không sửa tên riêng/số liệu. Nếu không có lỗi trả về [].\n\n" + chunks[ci];
                        long ttftMs, totalMs; bool fromCache; string usedModel;
                        string resp = DesktopAiGateway.GenerateStreaming(
                            "Bạn là chuyên gia hiệu đính văn bản tiếng Việt chuẩn NĐ 30 và Hướng dẫn 05-HD/VPTW.",
                            prompt,
                            null,
                            out ttftMs, out totalMs, out fromCache, out usedModel);
                        List<string[]> aiFound = WordSpellInspectorForm.ParseAiSpellResponseSafe(resp, chunks[ci]);
                        int chunkAdded = 0;
                        foreach (string[] row in aiFound)
                        {
                            if (VbtmCustomDictStore.IsGoodWord(row[0])) continue;
                            bool exists = false;
                            foreach (var ex in activeIssues) if (string.Equals(ex.Wrong, row[0], StringComparison.OrdinalIgnoreCase)) { exists = true; break; }
                            if (!exists)
                            {
                                activeIssues.Add(new SpellCardIssue { Wrong = row[0], Suggest = row[1], Badge = "Nghi ngờ", Reason = row[2], OccurrenceIndex = 1, IsSuspicious = true });
                                chunkAdded++;
                                totalAdded++;
                            }
                        }
                        if (chunkAdded > 0)
                        {
                            try { this.BeginInvoke((Action)(() => RenderCards())); } catch { }
                        }
                    }

                    this.BeginInvoke((Action)(() =>
                    {
                        btnAiDeep.Enabled = true;
                        RenderCards();
                        lblStatusBottom.Text = string.Format("✓ Hoàn tất ({0}) • Tổng {1} từ cần sửa (AI +{2})", VbtmKnowledgeHub.GetKnowledgeSummary(), activeIssues.Count, totalAdded);
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        btnAiDeep.Enabled = true;
                        lblStatusBottom.Text = "✓ " + VbtmKnowledgeHub.GetKnowledgeSummary() + " (" + activeIssues.Count + " lỗi)";
                    }));
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                }
            });
        }
    }
}
