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
            "BOONG", "SOONG", "XOONG", "GÒONG", "GOÒNG", "PC09", "PA05", "CSGT", "PCCC"
        };

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

        private static bool IsVietnameseVowel(char c)
        {
            const string vowels = "aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵ";
            return vowels.IndexOf(char.ToLower(c)) >= 0;
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

            // Handle "gi" + vowel (e.g., "gì" -> initial="g", nucleus="ì", or "giữ" -> initial="g", nucleus="iữ")
            if (initial == "q" && nucleus.Length >= 2 && nucleus[0] == 'u')
            {
                initial = "qu";
                nucleus = nucleus.Substring(1);
            }

            if (!ValidInitials.Contains(initial)) return false;
            if (!ValidCodas.Contains(coda)) return false;
            if (nucleus.Length > 3) return false;

            // Every char inside nucleus must be a vowel
            foreach (char c in nucleus) if (!IsVietnameseVowel(c)) return false;

            // Rules for ng/ngh, g/gh, c/k before i/e/ê
            char nv = nucleus[0];
            bool isFrontVowel = ("eèéẻẽẹêềếểễệiìíỉĩị".IndexOf(nv) >= 0);
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

        /// <summary>
        /// Repairs stray Telex keys / doubled chars inside a Vietnamese word:
        /// e.g. "điểsm" -> "điểm", "sắcs" -> "sắc", "hoạchh" -> "hoạch", "thựcj" -> "thực", "côngf" -> "công"
        /// </summary>
        public static string TryRepairTelexOrExtraChar(string word)
        {
            if (string.IsNullOrWhiteSpace(word) || word.Length < 3 || word.Length > 9) return null;
            if (WhitelistWords.Contains(word)) return null;
            if (IsValidVietnameseSyllable(word)) return null;

            string w = word.Normalize(NormalizationForm.FormC);
            if (!HasVietnameseDiacritic(w))
            {
                // Also catch obvious doubled trailing consonant on pure ASCII Vietnamese words like "trongg", "cungg"
                if (w.Length >= 4 && char.ToLower(w[w.Length - 1]) == char.ToLower(w[w.Length - 2]))
                {
                    string trimmed = w.Substring(0, w.Length - 1);
                    if (IsValidVietnameseSyllable(trimmed)) return trimmed;
                }
                return null;
            }

            // 1. Try deleting 1 stray Telex key ('s','f','r','x','j','w','z') or adjacent duplicate letter anywhere after the first char
            const string telexKeys = "sfrxjzw";
            for (int i = w.Length - 1; i >= 1; i--)
            {
                char ch = char.ToLower(w[i]);
                bool isTelex = telexKeys.IndexOf(ch) >= 0;
                bool isDup = (i > 0 && char.ToLower(w[i]) == char.ToLower(w[i - 1]));
                if (isTelex || isDup)
                {
                    string cand = w.Remove(i, 1);
                    if (IsValidVietnameseSyllable(cand))
                    {
                        return cand;
                    }
                }
            }

            // 2. Try deleting any single extra consonant inside the coda (e.g. "điểkm" -> "điểm")
            for (int i = w.Length - 1; i >= 2; i--)
            {
                if (!IsVietnameseVowel(w[i]))
                {
                    string cand = w.Remove(i, 1);
                    if (IsValidVietnameseSyllable(cand))
                    {
                        return cand;
                    }
                }
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

            // 1. Scan Telex / extra-character typos with 1-word left context (e.g., "kiểm điểsm" -> "kiểm điểm", "sâu sắcs" -> "sâu sắc")
            MatchCollection wordMatches = Regex.Matches(text, @"[A-Za-zÀ-Ỹà-ỹĐđ]+");
            for (int i = 0; i < wordMatches.Count; i++)
            {
                Match wm = wordMatches[i];
                if (coveredPositions.Contains(wm.Index)) continue;
                string tok = wm.Value;
                if (goodWords.Contains(tok) || WhitelistWords.Contains(tok)) continue;

                string repaired = TryRepairTelexOrExtraChar(tok);
                if (!string.IsNullOrEmpty(repaired) && !string.Equals(tok, repaired, StringComparison.Ordinal))
                {
                    string displayWrong = tok;
                    string displaySug = repaired;
                    int spanStart = wm.Index;
                    int spanLen = wm.Length;

                    // Include immediately preceding word if separated by a single space so "kiểm điểsm -> kiểm điểm" & "sâu sắcs -> sâu sắc" read naturally!
                    if (i > 0)
                    {
                        Match prev = wordMatches[i - 1];
                        int gapStart = prev.Index + prev.Length;
                        if (wm.Index - gapStart == 1 && text[gapStart] == ' ')
                        {
                            displayWrong = prev.Value + " " + tok;
                            displaySug = prev.Value + " " + repaired;
                            spanStart = prev.Index;
                            spanLen = (wm.Index + wm.Length) - prev.Index;
                        }
                    }

                    if (!goodWords.Contains(displayWrong))
                    {
                        ordered.Add(new KeyValuePair<int, SpellCardIssue>(spanStart, new SpellCardIssue
                        {
                            Wrong = displayWrong,
                            Suggest = displaySug,
                            Badge = "Chính tả",
                            Reason = "Lỗi gõ dư phím Telex / sai âm tiết tiếng Việt",
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

            btnAiDeep = new Button { Text = "⚡ Quét sâu AI (Ollama)", Location = new Point(222, 26), Size = new Size(206, 32), BackColor = Color.FromArgb(241, 245, 249), ForeColor = Color.FromArgb(30, 41, 59), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold), Cursor = Cursors.Hand };
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
            this.Shown += (s, e) => ScanActiveWordInstant();
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
            int y = 12;
            AddQuickActionCard(pnlFormatView, "📐 Thiết lập Căn lề & Khổ giấy (Trực quan)", "Mở giao diện căn lề A4 + xem trước trang in (NĐ 30 / HD 05)", y, () => { using (var f = new PageSetupVisualForm()) f.ShowDialog(); }); y += 68;
            AddQuickActionCard(pnlFormatView, "⚙ Cài đặt Hệ thống & Mẫu Thể thức", "Quản lý quy chuẩn NĐ 30, 05-HD/VPTW, CAND, Quốc hiệu, Chữ ký", y, () => { using (var f = new SystemAndTemplateSettingsForm(0)) f.ShowDialog(); }); y += 68;
            AddQuickActionCard(pnlFormatView, "💰 Đọc tiền sang chữ trong Word", "Đổi số tiền đang bôi đen thành chữ tiếng Việt chuẩn hành chính", y, () => WordQuickUtilities.ConvertSelectedMoneyToVietnameseWords()); y += 68;
            AddQuickActionCard(pnlFormatView, "💬 Xóa toàn bộ nhận xét (Comments)", "Dọn sạch toàn bộ ghi chú nhận xét trước khi phát hành văn bản", y, () => WordQuickUtilities.DeleteWordComments(true)); y += 68;
            AddQuickActionCard(pnlFormatView, "📱 Tạo Mã QR Code văn bản", "Chèn mã QR chuẩn ISO 18004 trực tiếp vào tài liệu Word", y, () => { using (var f = new WordQrInsertForm()) f.ShowDialog(); }); y += 68;
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
            lblStatusBottom.Text = "⏳ Đang gọi AI Local (Ollama) kiểm tra ngữ cảnh chuyên sâu...";
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    string sample = currentDocText.Length > 2200 ? currentDocText.Substring(0, 2200) : currentDocText;
                    string prompt = "Chỉ liệt kê lỗi chính tả/ngữ cảnh tiếng Việt thật sự sai trong đoạn văn sau dưới dạng JSON array: [{\"original\":\"cụm sai\",\"suggestion\":\"cụm đúng\",\"reason\":\"ngắn gọn\"}]. Không bịa lỗi, không sửa tên riêng/số liệu. Nếu không có lỗi trả về [].\n\n" + sample;
                    long ttftMs, totalMs; bool fromCache; string usedModel;
                    string resp = DesktopAiGateway.GenerateStreaming(
                        "Bạn là chuyên gia hiệu đính văn bản tiếng Việt chuẩn NĐ 30 và Hướng dẫn 05-HD/VPTW.",
                        prompt,
                        null,
                        out ttftMs, out totalMs, out fromCache, out usedModel);
                    List<string[]> aiFound = WordSpellInspectorForm.ParseAiSpellResponseSafe(resp, currentDocText);
                    int added = 0;
                    foreach (string[] row in aiFound)
                    {
                        if (VbtmCustomDictStore.IsGoodWord(row[0])) continue;
                        bool exists = false;
                        foreach (var ex in activeIssues) if (string.Equals(ex.Wrong, row[0], StringComparison.OrdinalIgnoreCase)) { exists = true; break; }
                        if (!exists)
                        {
                            activeIssues.Add(new SpellCardIssue { Wrong = row[0], Suggest = row[1], Badge = "Nghi ngờ", Reason = row[2], OccurrenceIndex = 1, IsSuspicious = true });
                            added++;
                        }
                    }
                    this.BeginInvoke((Action)(() =>
                    {
                        btnAiDeep.Enabled = true;
                        RenderCards();
                        lblStatusBottom.Text = "✓ AI Local hoàn tất — bổ sung thêm " + added + " gợi ý ngữ cảnh.";
                    }));
                }
                catch (Exception ex)
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        btnAiDeep.Enabled = true;
                        lblStatusBottom.Text = "AI Local chưa sẵn sàng (" + ex.Message + ") — vẫn dùng tốt chế độ Offline 100%.";
                    }));
                }
            });
        }
    }
}
