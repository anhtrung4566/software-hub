using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBTMProSuite
{
    // =========================================================================
    // 1. PERSISTENT CUSTOM SPELL DICTIONARY (%APPDATA%\VBTM\)
    // =========================================================================
    public static class VbtmCustomDictStore
    {
        public static string GetStoreDir()
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
            try { Directory.CreateDirectory(dir); } catch { }
            return dir;
        }

        public static string GoodPath { get { return Path.Combine(GetStoreDir(), "custom_dict_good.txt"); } }
        public static string BadPath { get { return Path.Combine(GetStoreDir(), "custom_dict_bad.txt"); } }

        public static List<string> LoadGoodWords()
        {
            List<string> list = new List<string>();
            try
            {
                if (File.Exists(GoodPath))
                {
                    foreach (string ln in File.ReadAllLines(GoodPath, Encoding.UTF8))
                    {
                        string w = ln.Trim();
                        if (!string.IsNullOrEmpty(w) && !list.Contains(w)) list.Add(w);
                    }
                }
            }
            catch { }
            return list;
        }

        public static void SaveGoodWords(List<string> words)
        {
            try
            {
                List<string> clean = new List<string>();
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string w in words)
                {
                    string t = (w ?? "").Trim();
                    if (!string.IsNullOrEmpty(t) && seen.Add(t)) clean.Add(t);
                }
                File.WriteAllLines(GoodPath, clean.ToArray(), new UTF8Encoding(false));
            }
            catch { }
        }

        public static void AddGoodWord(string word)
        {
            string t = (word ?? "").Trim();
            if (string.IsNullOrEmpty(t)) return;
            List<string> cur = LoadGoodWords();
            bool exists = false;
            foreach (string w in cur) if (string.Equals(w, t, StringComparison.OrdinalIgnoreCase)) { exists = true; break; }
            if (!exists)
            {
                cur.Add(t);
                SaveGoodWords(cur);
            }
        }

        public static bool IsGoodWord(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return false;
            string t = word.Trim();
            foreach (string w in LoadGoodWords())
            {
                if (string.Equals(w, t, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        public static Dictionary<string, string> LoadBadWords()
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (File.Exists(BadPath))
                {
                    foreach (string ln in File.ReadAllLines(BadPath, Encoding.UTF8))
                    {
                        int eq = ln.IndexOf('=');
                        if (eq > 0)
                        {
                            string k = ln.Substring(0, eq).Trim();
                            string v = ln.Substring(eq + 1).Trim();
                            if (!string.IsNullOrEmpty(k) && !string.IsNullOrEmpty(v)) map[k] = v;
                        }
                    }
                }
            }
            catch { }
            return map;
        }

        public static void SaveBadWords(Dictionary<string, string> map)
        {
            try
            {
                List<string> lines = new List<string>();
                foreach (KeyValuePair<string, string> kv in map)
                {
                    if (!string.IsNullOrWhiteSpace(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                        lines.Add(kv.Key.Trim() + "=" + kv.Value.Trim());
                }
                File.WriteAllLines(BadPath, lines.ToArray(), new UTF8Encoding(false));
            }
            catch { }
        }

        public static void AddBadWord(string wrong, string suggest)
        {
            string k = (wrong ?? "").Trim();
            string v = (suggest ?? "").Trim();
            if (string.IsNullOrEmpty(k) || string.IsNullOrEmpty(v)) return;
            Dictionary<string, string> map = LoadBadWords();
            map[k] = v;
            SaveBadWords(map);
        }
    }

    // =========================================================================
    // 2. SYSTEM & TEMPLATE PRESETS STORE (%APPDATA%\VBTM\vbtm_system_settings.json)
    // =========================================================================
    public class StandardRulePreset
    {
        public string Name { get; set; }
        public string LegalBasis { get; set; }
        public string Category { get; set; }
        public string PaperSize { get; set; }
        public bool Landscape { get; set; }
        public double TopMm { get; set; }
        public double BottomMm { get; set; }
        public double LeftMm { get; set; }
        public double RightMm { get; set; }
        public double GutterMm { get; set; }
        public bool MirrorMargins { get; set; }
        public string FontName { get; set; }
        public double FontSizePt { get; set; }
    }

    public class AgencyBlockPreset
    {
        public string Label { get; set; }
        public string StandardType { get; set; }
        public string ParentOrg { get; set; }
        public string IssuingOrg { get; set; }
        public string Location { get; set; }
        public string CodeAbbr { get; set; }
        public bool IsDefault { get; set; }
    }

    public class DocTypePreset
    {
        public string Label { get; set; }
        public string StandardType { get; set; }
        public string DocTypeName { get; set; }
        public string SummarySample { get; set; }
    }

    public class SignaturePreset
    {
        public string Label { get; set; }
        public string StandardType { get; set; }
        public string Authority { get; set; }
        public string Position { get; set; }
        public string MilitaryRank { get; set; }
        public string SignerName { get; set; }
        public string Recipients { get; set; }
        public bool IsDefault { get; set; }
    }

    public class SecurityMarkPreset
    {
        public string Label { get; set; }
        public string Category { get; set; }
        public string MarkText { get; set; }
        public string Position { get; set; }
    }

    public class RestReminderConfig
    {
        public bool Enabled { get; set; }
        public int IntervalMinutes { get; set; }
        public int SnoozeMinutes { get; set; }
    }

    public class VbtmSystemConfig
    {
        public List<StandardRulePreset> Standards { get; set; }
        public List<AgencyBlockPreset> Agencies { get; set; }
        public List<DocTypePreset> DocTypes { get; set; }
        public List<SignaturePreset> Signatures { get; set; }
        public List<SecurityMarkPreset> SecurityMarks { get; set; }
        public RestReminderConfig RestReminder { get; set; }
    }

    public static class VbtmSystemPresetStore
    {
        public static string ConfigPath
        {
            get { return Path.Combine(VbtmCustomDictStore.GetStoreDir(), "vbtm_system_settings.json"); }
        }

        public static VbtmSystemConfig CreateDefaults()
        {
            VbtmSystemConfig cfg = new VbtmSystemConfig();
            cfg.Standards = new List<StandardRulePreset>
            {
                new StandardRulePreset { Name = "Văn bản Hành chính", LegalBasis = "Nghị định 30/2020/NĐ-CP về công tác văn thư", Category = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", PaperSize = "A4", Landscape = false, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15, GutterMm = 0, MirrorMargins = false, FontName = "Times New Roman", FontSizePt = 14 },
                new StandardRulePreset { Name = "Văn bản Đảng", LegalBasis = "Hướng dẫn 05-HD/VPTW của Ban Chấp hành Trung ương", Category = "Văn bản Đảng (HD 05-HD/VPTW)", PaperSize = "A4", Landscape = false, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15, GutterMm = 0, MirrorMargins = false, FontName = "Times New Roman", FontSizePt = 14 },
                new StandardRulePreset { Name = "Văn bản Công an nhân dân (CAND)", LegalBasis = "Quy chuẩn mẫu văn bản Công an nhân dân (Bộ Công an)", Category = "Văn bản CAND", PaperSize = "A4", Landscape = false, TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15, GutterMm = 0, MirrorMargins = false, FontName = "Times New Roman", FontSizePt = 14 }
            };

            cfg.Agencies = new List<AgencyBlockPreset>
            {
                new AgencyBlockPreset { Label = "UBND Phường / Cơ quan Hành chính", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", ParentOrg = "UBND THÀNH PHỐ / TỈNH", IssuingOrg = "ỦY BAN NHÂN DÂN PHƯỜNG", Location = "Đà Nẵng", CodeAbbr = "UBND", IsDefault = true },
                new AgencyBlockPreset { Label = "Đảng ủy Phường / Chi bộ", StandardType = "Văn bản Đảng (HD 05-HD/VPTW)", ParentOrg = "ĐẢNG ỦY PHƯỜNG", IssuingOrg = "BAN CHẤP HÀNH ĐẢNG BỘ", Location = "Đà Nẵng", CodeAbbr = "ĐU", IsDefault = false },
                new AgencyBlockPreset { Label = "Đơn vị Công an nhân dân (CAND)", StandardType = "Văn bản CAND", ParentOrg = "CÔNG AN TỈNH / THÀNH PHỐ", IssuingOrg = "PHÒNG KỸ THUẬT HÌNH SỰ", Location = "Đà Nẵng", CodeAbbr = "PC09", IsDefault = false }
            };

            cfg.DocTypes = new List<DocTypePreset>
            {
                new DocTypePreset { Label = "1. Công văn hành chính / chỉ đạo", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "CÔNG VĂN", SummarySample = "V/v ... / Kính gửi: ..." },
                new DocTypePreset { Label = "2. Tiêu đề QUYẾT ĐỊNH cá biệt", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "QUYẾT ĐỊNH", SummarySample = "Về việc ..." },
                new DocTypePreset { Label = "3. Tiêu đề QUYẾT ĐỊNH kèm Quy chế", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "QUYẾT ĐỊNH", SummarySample = "Ban hành Quy chế ..." },
                new DocTypePreset { Label = "4. Tiêu đề NGHỊ QUYẾT", StandardType = "Văn bản Đảng (HD 05-HD/VPTW)", DocTypeName = "NGHỊ QUYẾT", SummarySample = "Về nhiệm vụ trọng tâm ..." },
                new DocTypePreset { Label = "5. Tiêu đề TỜ TRÌNH", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "TỜ TRÌNH", SummarySample = "Về việc đề nghị phê duyệt ..." },
                new DocTypePreset { Label = "6. Tiêu đề KẾ HOẠCH công tác", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "KẾ HOẠCH", SummarySample = "Triển khai thực hiện ..." },
                new DocTypePreset { Label = "7. Tiêu đề BÁO CÁO định kỳ", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "BÁO CÁO", SummarySample = "Kết quả thực hiện công tác ..." },
                new DocTypePreset { Label = "8. Tiêu đề THÔNG BÁO kết luận", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "THÔNG BÁO", SummarySample = "Kết luận của ... tại cuộc họp ..." },
                new DocTypePreset { Label = "9. Tiêu đề CHƯƠNG TRÌNH công tác", StandardType = "Văn bản Đảng (HD 05-HD/VPTW)", DocTypeName = "CHƯƠNG TRÌNH", SummarySample = "Công tác trọng tâm năm ..." },
                new DocTypePreset { Label = "10. Tiêu đề BIÊN BẢN cuộc họp", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "BIÊN BẢN", SummarySample = "Cuộc họp giao ban ..." },
                new DocTypePreset { Label = "11. Tiêu đề GIẤY MỜI họp", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "GIẤY MỜI", SummarySample = "Dự hội nghị triển khai ..." },
                new DocTypePreset { Label = "12. Tiêu đề CHỈ THỊ / CÔNG ĐIỆN", StandardType = "Văn bản CAND", DocTypeName = "CHỈ THỊ", SummarySample = "Về tăng cường công tác ..." },
                new DocTypePreset { Label = "13. Tiêu đề HƯỚNG DẪN chuyên môn", StandardType = "Văn bản Đảng (HD 05-HD/VPTW)", DocTypeName = "HƯỚNG DẪN", SummarySample = "Thực hiện quy trình ..." }
            };

            cfg.Signatures = new List<SignaturePreset>
            {
                new SignaturePreset { Label = "Thủ trưởng đơn vị (CAND / Hành chính)", StandardType = "Văn bản CAND", Authority = "KT. TRƯỞNG PHÒNG", Position = "PHÓ TRƯỞNG PHÒNG", MilitaryRank = "Thượng tá", SignerName = "Nguyễn Văn A", Recipients = "- Như trên;\r\n- Lưu: VT, TH.", IsDefault = true },
                new SignaturePreset { Label = "Bí thư Đảng ủy (Khối Đảng)", StandardType = "Văn bản Đảng (HD 05-HD/VPTW)", Authority = "T/M ĐẢNG ỦY", Position = "BÍ THƯ", MilitaryRank = "", SignerName = "Nguyễn Văn A", Recipients = "- Như Điều ...;\r\n- Lưu: VPĐU.", IsDefault = false }
            };

            cfg.SecurityMarks = new List<SecurityMarkPreset>
            {
                new SecurityMarkPreset { Label = "KHẨN", Category = "Độ khẩn", MarkText = "KHẨN", Position = "Dưới Số ký hiệu" },
                new SecurityMarkPreset { Label = "THƯỢNG KHẨN", Category = "Độ khẩn", MarkText = "THƯỢNG KHẨN", Position = "Dưới Số ký hiệu" },
                new SecurityMarkPreset { Label = "HỎA TỐC", Category = "Độ khẩn", MarkText = "HỎA TỐC", Position = "Dưới Số ký hiệu" },
                new SecurityMarkPreset { Label = "MẬT", Category = "Độ mật", MarkText = "MẬT", Position = "Góc trên trái trang" },
                new SecurityMarkPreset { Label = "TỐI MẬT", Category = "Độ mật", MarkText = "TỐI MẬT", Position = "Góc trên trái trang" }
            };

            cfg.RestReminder = new RestReminderConfig { Enabled = true, IntervalMinutes = 60, SnoozeMinutes = 30 };
            return cfg;
        }

        public static VbtmSystemConfig Load()
        {
            try
            {
                if (File.Exists(ConfigPath))
                {
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    VbtmSystemConfig cfg = js.Deserialize<VbtmSystemConfig>(File.ReadAllText(ConfigPath, Encoding.UTF8));
                    if (cfg != null)
                    {
                        VbtmSystemConfig def = CreateDefaults();
                        if (cfg.Standards == null || cfg.Standards.Count == 0) cfg.Standards = def.Standards;
                        if (cfg.Agencies == null || cfg.Agencies.Count == 0) cfg.Agencies = def.Agencies;
                        if (cfg.DocTypes == null || cfg.DocTypes.Count == 0) cfg.DocTypes = def.DocTypes;
                        if (cfg.Signatures == null || cfg.Signatures.Count == 0) cfg.Signatures = def.Signatures;
                        if (cfg.SecurityMarks == null || cfg.SecurityMarks.Count == 0) cfg.SecurityMarks = def.SecurityMarks;
                        if (cfg.RestReminder == null) cfg.RestReminder = def.RestReminder;
                        return cfg;
                    }
                }
            }
            catch { }
            return CreateDefaults();
        }

        public static void Save(VbtmSystemConfig cfg)
        {
            try
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                File.WriteAllText(ConfigPath, js.Serialize(cfg), new UTF8Encoding(false));
            }
            catch { }
        }
    }

    // =========================================================================
    // 3. WORD COM QUICK UTILITIES: ĐỌC TIỀN SANG CHỮ, XÓA NHẬN XÉT, THÔNG TIN
    // =========================================================================
    public static class WordQuickUtilities
    {
        private static readonly string[] Digits = { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };
        private static readonly string[] ScaleUnits = { "", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ" };

        private static string ReadThreeDigits(int n, bool fullHundreds)
        {
            int h = n / 100;
            int t = (n % 100) / 10;
            int u = n % 10;
            List<string> parts = new List<string>();

            if (h > 0 || fullHundreds)
            {
                parts.Add(Digits[h] + " trăm");
            }

            if (t == 0)
            {
                if (u > 0)
                {
                    if (h > 0 || fullHundreds) parts.Add("linh");
                    parts.Add(Digits[u]);
                }
            }
            else if (t == 1)
            {
                parts.Add("mười");
                if (u == 5) parts.Add("lăm");
                else if (u > 0) parts.Add(Digits[u]);
            }
            else
            {
                parts.Add(Digits[t] + " mươi");
                if (u == 1) parts.Add("mốt");
                else if (u == 4) parts.Add("tư");
                else if (u == 5) parts.Add("lăm");
                else if (u > 0) parts.Add(Digits[u]);
            }
            return string.Join(" ", parts.ToArray());
        }

        public static string NumberToVietnameseMoney(string rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return "";
            string s = rawInput.Trim();
            s = Regex.Replace(s, @"(vnđ|vnd|đồng|đ|\s+)", "", RegexOptions.IgnoreCase);
            bool negative = false;
            if (s.StartsWith("-")) { negative = true; s = s.Substring(1); }

            // Strip thousand separators (. or ,) — support Vietnamese format e.g. 15.500.000 or 15,500,000
            string digitsOnly = Regex.Replace(s, @"[.,\s]", "");
            if (!Regex.IsMatch(digitsOnly, @"^\d+$")) return "";

            digitsOnly = digitsOnly.TrimStart('0');
            if (string.IsNullOrEmpty(digitsOnly)) return "Không đồng";

            List<int> groups = new List<int>();
            for (int i = digitsOnly.Length; i > 0; i -= 3)
            {
                int start = Math.Max(0, i - 3);
                int len = i - start;
                groups.Add(int.Parse(digitsOnly.Substring(start, len)));
            }

            List<string> parts = new List<string>();
            for (int i = groups.Count - 1; i >= 0; i--)
            {
                int val = groups[i];
                if (val == 0) continue;
                bool fullHundreds = (i < groups.Count - 1);
                string text3 = ReadThreeDigits(val, fullHundreds);
                string unit = (i < ScaleUnits.Length) ? ScaleUnits[i] : "";
                if (!string.IsNullOrEmpty(unit)) parts.Add(text3 + " " + unit);
                else parts.Add(text3);
            }

            string res = string.Join(" ", parts.ToArray()).Trim();
            if (negative) res = "âm " + res;
            if (res.Length > 0) res = char.ToUpper(res[0]) + res.Substring(1);
            return res + " đồng";
        }

        public static string GetWordSelectedText()
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                string txt = Convert.ToString(sel.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, sel, null)) ?? "";
                return txt.Trim('\r', '\n', '\a', ' ', '\t');
            }
            catch { return ""; }
        }

        public static void ConvertSelectedMoneyToVietnameseWords()
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                string txt = Convert.ToString(sel.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, sel, null)) ?? "";
                string clean = txt.Trim('\r', '\n', '\a', ' ', '\t');
                string words = NumberToVietnameseMoney(clean);
                if (!string.IsNullOrEmpty(words))
                {
                    sel.GetType().InvokeMember("Collapse", BindingFlags.InvokeMethod, null, sel, new object[] { 0 }); // wdCollapseEnd
                    sel.GetType().InvokeMember("TypeText", BindingFlags.InvokeMethod, null, sel, new object[] { " (Bằng chữ: " + words + ")" });
                    return;
                }
            }
            catch { }

            // Fallback: interactive converter dialog if nothing valid was selected
            using (Form f = new Form())
            {
                f.Text = "Đọc tiền sang chữ — VBTM Pro 3.6";
                f.Size = new Size(520, 260);
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.BackColor = Color.FromArgb(248, 250, 252);
                f.Font = new Font("Segoe UI", 9.5f);

                Label lbl1 = new Label { Text = "Nhập số tiền cần đọc (VD: 15.500.000 hoặc 2500000):", Location = new Point(20, 18), AutoSize = true, ForeColor = Color.FromArgb(30, 41, 59), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
                TextBox txtNum = new TextBox { Location = new Point(20, 44), Width = 460, Font = new Font("Segoe UI", 11f) };
                Label lbl2 = new Label { Text = "Kết quả bằng chữ:", Location = new Point(20, 82), AutoSize = true, ForeColor = Color.FromArgb(37, 99, 235), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
                TextBox txtOut = new TextBox { Location = new Point(20, 106), Size = new Size(460, 56), Multiline = true, ReadOnly = true, BackColor = Color.White, Font = new Font("Segoe UI", 10f) };

                txtNum.TextChanged += (s, e) => { txtOut.Text = NumberToVietnameseMoney(txtNum.Text); };

                Button btnInsert = new Button { Text = "✓ Chèn vào Word", Location = new Point(195, 176), Size = new Size(145, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
                btnInsert.FlatAppearance.BorderSize = 0;
                btnInsert.Click += (s, e) =>
                {
                    if (string.IsNullOrEmpty(txtOut.Text)) return;
                    try
                    {
                        object wordApp = Marshal.GetActiveObject("Word.Application");
                        object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                        sel.GetType().InvokeMember("TypeText", BindingFlags.InvokeMethod, null, sel, new object[] { txtNum.Text.Trim() + " đồng (Bằng chữ: " + txtOut.Text + ")" });
                        f.Close();
                    }
                    catch
                    {
                        Clipboard.SetText(txtOut.Text);
                        MessageBox.Show("Không tìm thấy tài liệu Word đang mở. Đã sao chép kết quả vào Clipboard!", "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                };

                Button btnCopy = new Button { Text = "Sao chép", Location = new Point(350, 176), Size = new Size(130, 36), BackColor = Color.White, ForeColor = Color.FromArgb(30, 41, 59), FlatStyle = FlatStyle.Flat };
                btnCopy.Click += (s, e) => { if (!string.IsNullOrEmpty(txtOut.Text)) { Clipboard.SetText(txtOut.Text); MessageBox.Show("Đã sao chép: " + txtOut.Text, "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information); } };

                f.Controls.AddRange(new Control[] { lbl1, txtNum, lbl2, txtOut, btnInsert, btnCopy });
                f.ShowDialog();
            }
        }

        public static void DeleteWordComments(bool deleteAll)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                int deleted = 0;
                if (deleteAll)
                {
                    object comments = doc.GetType().InvokeMember("Comments", BindingFlags.GetProperty, null, doc, null);
                    int count = (int)comments.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, comments, null);
                    for (int i = count; i >= 1; i--)
                    {
                        object c = comments.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, comments, new object[] { i });
                        c.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, c, null);
                        deleted++;
                    }
                    MessageBox.Show("Đã xóa toàn bộ " + deleted + " nhận xét (Comments) trong tài liệu Word.", "VBTM Pro 3.6 — Xóa nhận xét", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                    object rng = sel.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, sel, null);
                    object comments = rng.GetType().InvokeMember("Comments", BindingFlags.GetProperty, null, rng, null);
                    int count = (int)comments.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, comments, null);
                    for (int i = count; i >= 1; i--)
                    {
                        object c = comments.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, comments, new object[] { i });
                        c.GetType().InvokeMember("Delete", BindingFlags.InvokeMethod, null, c, null);
                        deleted++;
                    }
                    MessageBox.Show(deleted > 0 ? ("Đã xóa " + deleted + " nhận xét trong vùng đang chọn.") : "Không có nhận xét nào trong vùng đang chọn.", "VBTM Pro 3.6 — Xóa nhận xét", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Chưa mở tài liệu Word hoặc không thể xóa nhận xét: " + ex.Message, "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public static void ShowAboutDialog()
        {
            using (Form f = new Form())
            {
                f.Text = "Thông tin Phần mềm — VBTM Pro 3.6 (Offline Edition)";
                f.Size = new Size(560, 390);
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.BackColor = Color.White;
                f.Font = new Font("Segoe UI", 9.5f);

                Panel hdr = new Panel { Dock = DockStyle.Top, Height = 82, BackColor = Color.FromArgb(37, 99, 235) };
                hdr.Controls.Add(new Label { Text = "VBTM PRO 3.6 — OFFLINE EDITION", Location = new Point(22, 16), AutoSize = true, ForeColor = Color.White, Font = new Font("Segoe UI", 14f, FontStyle.Bold) });
                hdr.Controls.Add(new Label { Text = "Chuẩn hóa Thể thức Văn bản & Trợ lý Chính tả AI Local cho Microsoft Word", Location = new Point(22, 48), AutoSize = true, ForeColor = Color.FromArgb(219, 234, 254), Font = new Font("Segoe UI", 9.5f) });
                f.Controls.Add(hdr);

                Label body = new Label
                {
                    Location = new Point(24, 98),
                    Size = new Size(500, 210),
                    ForeColor = Color.FromArgb(30, 41, 59),
                    Text =
                        "• Quy chuẩn tích hợp:\r\n" +
                        "  - Nghị định 30/2020/NĐ-CP (Văn bản hành chính Nhà nước & Doanh nghiệp)\r\n" +
                        "  - Hướng dẫn 05-HD/VPTW (Văn bản Đảng & Tổ chức chính trị - xã hội)\r\n" +
                        "  - Quy chuẩn mẫu văn bản Công an nhân dân (CAND)\r\n\r\n" +
                        "• Công nghệ AI & Bảo mật (100% Offline):\r\n" +
                        "  - Hỗ trợ AI Local (Ollama qwen2.5:3b / qwen2.5:7b) không gửi dữ liệu ra Internet.\r\n" +
                        "  - Tương thích Microsoft Word 32-bit & 64-bit trên Windows 10/11.\r\n" +
                        "  - Không yêu cầu quyền Quản trị viên (Administrator) khi cài đặt.\r\n\r\n" +
                        "• Trang chủ: https://anhtrung4566.github.io/software-hub/products/vbtm-pro.html"
                };
                f.Controls.Add(body);

                Button btnClose = new Button { Text = "Đóng", Location = new Point(428, 312), Size = new Size(96, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.Click += (s, e) => f.Close();
                f.Controls.Add(btnClose);
                f.ShowDialog();
            }
        }
    }

    // =========================================================================
    // 4. DIALOG 1: THIẾT LẬP CĂN LỀ & KHỔ GIẤY (LIVE VISUAL PREVIEW)
    // =========================================================================
    public class PageSetupVisualForm : Form
    {
        private ComboBox cboPaper, cboGutterPos;
        private RadioButton rbPortrait, rbLandscape;
        private NumericUpDown numTop, numBottom, numLeft, numRight, numGutter, numHeader, numFooter;
        private Panel pnlPreviewCanvas;
        private Label lblPreviewSummary;

        public PageSetupVisualForm()
        {
            this.Text = "THIẾT LẬP CĂN LỀ & KHỔ GIẤY — VBTM Pro 3.6";
            this.Size = new Size(760, 595);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(241, 245, 249);
            this.Font = new Font("Segoe UI", 9.5f);

            // Header bar
            Panel hdr = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Color.White };
            Panel iconBox = new Panel { Location = new Point(18, 12), Size = new Size(40, 40), BackColor = Color.FromArgb(37, 99, 235) };
            iconBox.Controls.Add(new Label { Text = "📄", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.White, Font = new Font("Segoe UI", 14f) });
            hdr.Controls.Add(iconBox);
            hdr.Controls.Add(new Label { Text = "THIẾT LẬP CĂN LỀ & KHỔ GIẤY", Location = new Point(70, 12), AutoSize = true, ForeColor = Color.FromArgb(29, 78, 216), Font = new Font("Segoe UI", 12f, FontStyle.Bold) });
            hdr.Controls.Add(new Label { Text = "VBTM Pro • Quy chuẩn thể thức Quốc gia, Đảng & CAND", Location = new Point(70, 36), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.8f) });
            this.Controls.Add(hdr);

            // Left Card 1: Khổ giấy & Hướng in
            Panel cardPaper = MakeCard(18, 80, 415, 108, "KHỔ GIẤY & HƯỚNG IN");
            cardPaper.Controls.Add(new Label { Text = "Khổ giấy:", Location = new Point(16, 36), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            cboPaper = new ComboBox { Location = new Point(16, 58), Width = 215, DropDownStyle = ComboBoxStyle.DropDownList };
            cboPaper.Items.AddRange(new object[] { "A4 (210 × 297 mm) - Tiêu chuẩn", "A5 (148 × 210 mm)", "Letter (215.9 × 279.4 mm)" });
            cboPaper.SelectedIndex = 0;
            cardPaper.Controls.Add(cboPaper);

            cardPaper.Controls.Add(new Label { Text = "Hướng in:", Location = new Point(248, 36), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            rbPortrait = new RadioButton { Text = "Đứng", Location = new Point(248, 60), AutoSize = true, Checked = true };
            rbLandscape = new RadioButton { Text = "Ngang", Location = new Point(325, 60), AutoSize = true };
            cardPaper.Controls.Add(rbPortrait);
            cardPaper.Controls.Add(rbLandscape);
            this.Controls.Add(cardPaper);

            // Left Card 2: 4 thông số căn lề chính
            Panel cardMargins = MakeCard(18, 200, 415, 152, "4 THÔNG SỐ CĂN LỀ CHÍNH (MM)");
            numTop = AddNumField(cardMargins, "Lề trên (Top mm):", 16, 36, 20);
            numBottom = AddNumField(cardMargins, "Lề dưới (Bottom mm):", 214, 36, 20);
            numLeft = AddNumField(cardMargins, "Lề trái (Left mm - Gáy):", 16, 92, 30);
            numRight = AddNumField(cardMargins, "Lề phải (Right mm):", 214, 92, 15);
            this.Controls.Add(cardMargins);

            // Left Card 3: Lề gáy & Header/Footer
            Panel cardGutter = MakeCard(18, 364, 415, 134, "LỀ GÁY (GUTTER) & KHOẢNG CÁCH MÉP");
            numGutter = AddNumField(cardGutter, "Lề gáy (Gutter mm):", 16, 32, 0);
            cardGutter.Controls.Add(new Label { Text = "Vị trí lề gáy:", Location = new Point(214, 32), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            cboGutterPos = new ComboBox { Location = new Point(214, 54), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            cboGutterPos.Items.AddRange(new object[] { "Lề trái (Left) - Đóng gáy", "Lề trên (Top)" });
            cboGutterPos.SelectedIndex = 0;
            cardGutter.Controls.Add(cboGutterPos);
            numHeader = AddNumField(cardGutter, "Đầu trang (Header mm):", 16, 82, 10);
            numFooter = AddNumField(cardGutter, "Chân trang (Footer mm):", 214, 82, 10);
            this.Controls.Add(cardGutter);

            // Right Card 1: Trực quan hóa trang in
            Panel cardPreview = MakeCard(450, 80, 276, 240, "📄 TRỰC QUAN HÓA TRANG IN");
            pnlPreviewCanvas = new Panel { Location = new Point(28, 36), Size = new Size(220, 170), BackColor = Color.FromArgb(241, 245, 249) };
            pnlPreviewCanvas.Paint += DrawPagePreview;
            cardPreview.Controls.Add(pnlPreviewCanvas);
            lblPreviewSummary = new Label { Location = new Point(12, 212), Size = new Size(252, 20), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.5f) };
            cardPreview.Controls.Add(lblPreviewSummary);
            this.Controls.Add(cardPreview);

            // Right Card 2: Mức chuẩn quy định (4 Quick Presets)
            Panel cardPresets = MakeCard(450, 332, 276, 166, "MỨC CHUẨN QUY ĐỊNH");
            AddPresetBtn(cardPresets, "⚡ NĐ 30 Chuẩn (20-20-30-15)", 36, 20, 20, 30, 15);
            AddPresetBtn(cardPresets, "⚡ NĐ 30 Đều (20-20-30-20)", 68, 20, 20, 30, 20);
            AddPresetBtn(cardPresets, "⚡ NĐ 30 Rộng (25-25-35-20)", 100, 25, 25, 35, 20);
            AddPresetBtn(cardPresets, "⚡ Đảng HD 05 (20-20-30-15)", 132, 20, 20, 30, 15);
            this.Controls.Add(cardPresets);

            // Bottom footer bar
            Panel ftr = new Panel { Dock = DockStyle.Bottom, Height = 56, BackColor = Color.White };
            Button btnReset = new Button { Text = "↩ Mặc định QC", Location = new Point(18, 10), Size = new Size(130, 36), BackColor = Color.White, ForeColor = Color.FromArgb(51, 65, 85), FlatStyle = FlatStyle.Flat };
            btnReset.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnReset.Click += (s, e) => ApplyValues(20, 20, 30, 15, 0, 10, 10, false);

            Button btnApply = new Button { Text = "✓ Áp Dụng Căn Lề", Location = new Point(468, 10), Size = new Size(160, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            btnApply.FlatAppearance.BorderSize = 0;
            btnApply.Click += (s, e) => ApplyToActiveWord();

            Button btnClose = new Button { Text = "Đóng", Location = new Point(638, 10), Size = new Size(88, 36), BackColor = Color.White, ForeColor = Color.FromArgb(51, 65, 85), FlatStyle = FlatStyle.Flat };
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnClose.Click += (s, e) => this.Close();

            ftr.Controls.AddRange(new Control[] { btnReset, btnApply, btnClose });
            this.Controls.Add(ftr);

            // Live redraw events
            EventHandler redraw = (s, e) => RefreshPreview();
            numTop.ValueChanged += redraw;
            numBottom.ValueChanged += redraw;
            numLeft.ValueChanged += redraw;
            numRight.ValueChanged += redraw;
            numGutter.ValueChanged += redraw;
            rbPortrait.CheckedChanged += redraw;
            rbLandscape.CheckedChanged += redraw;
            cboPaper.SelectedIndexChanged += redraw;

            ReadCurrentWordMargins();
            RefreshPreview();
        }

        private Panel MakeCard(int x, int y, int w, int h, string title)
        {
            Panel p = new Panel { Location = new Point(x, y), Size = new Size(w, h), BackColor = Color.White };
            p.Paint += (s, e) => { using (Pen pen = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1); };
            p.Controls.Add(new Label { Text = title, Location = new Point(16, 10), AutoSize = true, ForeColor = Color.FromArgb(37, 99, 235), Font = new Font("Segoe UI", 9f, FontStyle.Bold) });
            return p;
        }

        private NumericUpDown AddNumField(Panel parent, string label, int x, int y, decimal val)
        {
            parent.Controls.Add(new Label { Text = label, Location = new Point(x, y), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            NumericUpDown n = new NumericUpDown { Location = new Point(x, y + 20), Width = 180, Minimum = 0, Maximum = 100, DecimalPlaces = 0, Value = val };
            parent.Controls.Add(n);
            return n;
        }

        private void AddPresetBtn(Panel parent, string text, int y, decimal t, decimal b, decimal l, decimal r)
        {
            LinkLabel lnk = new LinkLabel { Text = text, Location = new Point(18, y), Size = new Size(240, 24), TextAlign = ContentAlignment.MiddleCenter, LinkColor = Color.FromArgb(37, 99, 235), ActiveLinkColor = Color.FromArgb(29, 78, 216), LinkBehavior = LinkBehavior.HoverUnderline, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            lnk.Click += (s, e) => ApplyValues(t, b, l, r, 0, 10, 10, false);
            parent.Controls.Add(lnk);
        }

        private void ApplyValues(decimal t, decimal b, decimal l, decimal r, decimal g, decimal h, decimal f, bool landscape)
        {
            numTop.Value = t; numBottom.Value = b; numLeft.Value = l; numRight.Value = r;
            numGutter.Value = g; numHeader.Value = h; numFooter.Value = f;
            rbLandscape.Checked = landscape; rbPortrait.Checked = !landscape;
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            string orient = rbLandscape.Checked ? "Ngang" : "Đứng";
            string pName = cboPaper.SelectedIndex == 1 ? "A5" : (cboPaper.SelectedIndex == 2 ? "Letter" : "A4");
            lblPreviewSummary.Text = string.Format("{0} {1} • {2}-{3}-{4}-{5} mm", pName, orient, numTop.Value, numBottom.Value, numLeft.Value, numRight.Value);
            pnlPreviewCanvas.Invalidate();
        }

        private void DrawPagePreview(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int cw = pnlPreviewCanvas.Width;
            int ch = pnlPreviewCanvas.Height;

            bool land = rbLandscape.Checked;
            int pw = land ? 168 : 112;
            int ph = land ? 116 : 152;
            int px = (cw - pw) / 2;
            int py = (ch - ph) / 2;

            using (SolidBrush brBg = new SolidBrush(Color.FromArgb(239, 246, 255)))
                g.FillRectangle(brBg, px, py, pw, ph);
            using (Pen borderPen = new Pen(Color.FromArgb(148, 163, 184), 1.2f))
                g.DrawRectangle(borderPen, px, py, pw, ph);

            float pageWmm = land ? 297f : 210f;
            float pageHmm = land ? 210f : 297f;
            float sx = pw / pageWmm;
            float sy = ph / pageHmm;

            int ml = (int)Math.Round(((float)numLeft.Value + (cboGutterPos.SelectedIndex == 0 ? (float)numGutter.Value : 0f)) * sx);
            int mr = (int)Math.Round((float)numRight.Value * sx);
            int mt = (int)Math.Round(((float)numTop.Value + (cboGutterPos.SelectedIndex == 1 ? (float)numGutter.Value : 0f)) * sy);
            int mb = (int)Math.Round((float)numBottom.Value * sy);

            int cx = px + ml;
            int cy = py + mt;
            int contentW = Math.Max(16, pw - ml - mr);
            int contentH = Math.Max(16, ph - mt - mb);

            g.FillRectangle(Brushes.White, cx, cy, contentW, contentH);
            using (Pen innerPen = new Pen(Color.FromArgb(59, 130, 246), 1.2f))
                g.DrawRectangle(innerPen, cx, cy, contentW, contentH);

            // Simulated document lines
            using (SolidBrush darkBar = new SolidBrush(Color.FromArgb(148, 163, 184)))
            using (SolidBrush lightBar = new SolidBrush(Color.FromArgb(203, 213, 225)))
            {
                if (contentH > 24) g.FillRectangle(darkBar, cx + 6, cy + 8, Math.Max(10, (int)(contentW * 0.65)), 3);
                if (contentH > 34) g.FillRectangle(lightBar, cx + 6, cy + 16, Math.Max(10, contentW - 12), 2);
                if (contentH > 42) g.FillRectangle(lightBar, cx + 6, cy + 23, Math.Max(10, contentW - 12), 2);
                if (contentH > 50) g.FillRectangle(lightBar, cx + 6, cy + 30, Math.Max(10, (int)(contentW * 0.8)), 2);
            }
        }

        private void ReadCurrentWordMargins()
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object ps = doc.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, doc, null);
                double pt2mm = 25.4 / 72.0;
                double t = Convert.ToDouble(ps.GetType().InvokeMember("TopMargin", BindingFlags.GetProperty, null, ps, null)) * pt2mm;
                double b = Convert.ToDouble(ps.GetType().InvokeMember("BottomMargin", BindingFlags.GetProperty, null, ps, null)) * pt2mm;
                double l = Convert.ToDouble(ps.GetType().InvokeMember("LeftMargin", BindingFlags.GetProperty, null, ps, null)) * pt2mm;
                double r = Convert.ToDouble(ps.GetType().InvokeMember("RightMargin", BindingFlags.GetProperty, null, ps, null)) * pt2mm;
                int orient = Convert.ToInt32(ps.GetType().InvokeMember("Orientation", BindingFlags.GetProperty, null, ps, null));
                numTop.Value = Math.Max(0, Math.Min(100, (decimal)Math.Round(t)));
                numBottom.Value = Math.Max(0, Math.Min(100, (decimal)Math.Round(b)));
                numLeft.Value = Math.Max(0, Math.Min(100, (decimal)Math.Round(l)));
                numRight.Value = Math.Max(0, Math.Min(100, (decimal)Math.Round(r)));
                rbLandscape.Checked = (orient == 1);
                rbPortrait.Checked = (orient != 1);
            }
            catch { }
        }

        private void ApplyToActiveWord()
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object ps = doc.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, doc, null);
                float mm2pt = 72f / 25.4f;
                int paperCode = cboPaper.SelectedIndex == 1 ? 8 : (cboPaper.SelectedIndex == 2 ? 2 : 7); // 7=A4, 8=A5, 2=Letter
                ps.GetType().InvokeMember("PaperSize", BindingFlags.SetProperty, null, ps, new object[] { paperCode });
                ps.GetType().InvokeMember("Orientation", BindingFlags.SetProperty, null, ps, new object[] { rbLandscape.Checked ? 1 : 0 });
                ps.GetType().InvokeMember("TopMargin", BindingFlags.SetProperty, null, ps, new object[] { (float)numTop.Value * mm2pt });
                ps.GetType().InvokeMember("BottomMargin", BindingFlags.SetProperty, null, ps, new object[] { (float)numBottom.Value * mm2pt });
                ps.GetType().InvokeMember("LeftMargin", BindingFlags.SetProperty, null, ps, new object[] { (float)numLeft.Value * mm2pt });
                ps.GetType().InvokeMember("RightMargin", BindingFlags.SetProperty, null, ps, new object[] { (float)numRight.Value * mm2pt });
                ps.GetType().InvokeMember("Gutter", BindingFlags.SetProperty, null, ps, new object[] { (float)numGutter.Value * mm2pt });
                ps.GetType().InvokeMember("GutterPos", BindingFlags.SetProperty, null, ps, new object[] { cboGutterPos.SelectedIndex });
                ps.GetType().InvokeMember("HeaderDistance", BindingFlags.SetProperty, null, ps, new object[] { (float)numHeader.Value * mm2pt });
                ps.GetType().InvokeMember("FooterDistance", BindingFlags.SetProperty, null, ps, new object[] { (float)numFooter.Value * mm2pt });

                MessageBox.Show("Đã áp dụng căn lề & khổ giấy (" + lblPreviewSummary.Text + ") vào tài liệu Word!", "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không tìm thấy tài liệu Word đang mở: " + ex.Message, "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    // =========================================================================
    // 5. DIALOG 2: CÀI ĐẶT HỆ THỐNG & MẪU THỂ THỨC (6 SIDEBAR SECTIONS + CAND + REST REMINDER)
    // =========================================================================
    public class RestReminderToastForm : Form
    {
        protected override bool ShowWithoutActivation { get { return true; } }

        public RestReminderToastForm(int workMins, int snoozeMins)
        {
            this.FormBorderStyle = FormBorderStyle.None;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Size = new Size(370, 146);
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(wa.Right - this.Width - 18, wa.Bottom - this.Height - 18);

            this.Paint += (s, e) =>
            {
                using (Pen p = new Pen(Color.FromArgb(37, 99, 235), 2))
                    e.Graphics.DrawRectangle(p, 1, 1, this.Width - 3, this.Height - 3);
            };

            Panel topBar = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = Color.FromArgb(239, 246, 255) };
            topBar.Controls.Add(new Label { Text = "☕ VBTM Pro — Nhắc nhở nghỉ ngơi", Location = new Point(12, 7), AutoSize = true, ForeColor = Color.FromArgb(29, 78, 216), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) });
            this.Controls.Add(topBar);

            Label msg = new Label
            {
                Location = new Point(14, 42),
                Size = new Size(342, 48),
                ForeColor = Color.FromArgb(51, 65, 85),
                Text = "Bạn đã soạn thảo liên tục " + workMins + " phút. Hãy thư giãn mắt 20 giây, uống một ngụm nước và đứng dậy vận động nhẹ nhé!"
            };
            this.Controls.Add(msg);

            Button btnSnooze = new Button { Text = "Báo lại (" + snoozeMins + "p)", Location = new Point(14, 100), Size = new Size(125, 32), BackColor = Color.FromArgb(241, 245, 249), ForeColor = Color.FromArgb(30, 41, 59), FlatStyle = FlatStyle.Flat };
            btnSnooze.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnSnooze.Click += (s, e) => this.Close();

            Button btnOk = new Button { Text = "✓ Đã hiểu", Location = new Point(256, 100), Size = new Size(98, 32), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.Click += (s, e) => this.Close();

            this.Controls.Add(btnSnooze);
            this.Controls.Add(btnOk);

            Timer autoHide = new Timer { Interval = 12000 };
            autoHide.Tick += (s, e) => { autoHide.Stop(); try { this.Close(); } catch { } };
            autoHide.Start();
        }
    }

    public class SystemAndTemplateSettingsForm : Form
    {
        private VbtmSystemConfig cfg;
        private int activeTab = 0;
        private Panel[] navButtons = new Panel[6];
        private Label[] navSubLabels = new Label[6];
        private Panel pnlContentHost;
        private Label lblFooterStatus;

        public SystemAndTemplateSettingsForm(int initialTab = 0)
        {
            cfg = VbtmSystemPresetStore.Load();
            activeTab = Math.Max(0, Math.Min(5, initialTab));

            this.Text = "CÀI ĐẶT HỆ THỐNG & MẪU THỂ THỨC — VBTM Pro 3.6";
            this.Size = new Size(960, 640);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.2f);

            // Header
            Panel hdr = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Color.White };
            hdr.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawLine(p, 0, hdr.Height - 1, hdr.Width, hdr.Height - 1); };
            Panel iconBox = new Panel { Location = new Point(18, 12), Size = new Size(38, 38), BackColor = Color.FromArgb(37, 99, 235) };
            iconBox.Controls.Add(new Label { Text = "⚙", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.White, Font = new Font("Segoe UI", 15f) });
            hdr.Controls.Add(iconBox);
            hdr.Controls.Add(new Label { Text = "CÀI ĐẶT HỆ THỐNG & MẪU THỂ THỨC", Location = new Point(68, 12), AutoSize = true, ForeColor = Color.FromArgb(15, 23, 42), Font = new Font("Segoe UI", 11.5f, FontStyle.Bold) });
            hdr.Controls.Add(new Label { Text = "Cấu hình quy chuẩn văn bản (NĐ 30, HD 05-HD/VPTW, CAND), thiết lập mặc định và quản lý danh mục mẫu thể thức", Location = new Point(68, 36), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.6f) });
            this.Controls.Add(hdr);

            // Footer
            Panel ftr = new Panel { Dock = DockStyle.Bottom, Height = 56, BackColor = Color.White };
            ftr.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawLine(p, 0, 0, ftr.Width, 0); };
            Button btnReset = new Button { Text = "Khôi phục mặc định", Location = new Point(18, 11), Size = new Size(148, 34), BackColor = Color.White, ForeColor = Color.FromArgb(51, 65, 85), FlatStyle = FlatStyle.Flat };
            btnReset.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnReset.Click += (s, e) =>
            {
                if (MessageBox.Show("Khôi phục toàn bộ quy chuẩn & mẫu thể thức về mặc định ban đầu?", "VBTM Pro 3.6", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    cfg = VbtmSystemPresetStore.CreateDefaults();
                    VbtmSystemPresetStore.Save(cfg);
                    UpdateNavCounts();
                    RenderActiveTab();
                    lblFooterStatus.Text = "Đã khôi phục cấu hình mặc định.";
                }
            };
            lblFooterStatus = new Label { Text = "Sẵn sàng quản lý và tùy biến mẫu thể thức (100% Offline).", Location = new Point(178, 18), AutoSize = true, ForeColor = Color.FromArgb(37, 99, 235) };

            Button btnSave = new Button { Text = "Lưu cài đặt", Location = new Point(735, 11), Size = new Size(110, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) =>
            {
                VbtmSystemPresetStore.Save(cfg);
                lblFooterStatus.Text = "✓ Đã lưu toàn bộ cài đặt & mẫu thể thức thành công!";
                MessageBox.Show("Đã lưu cài đặt hệ thống & mẫu thể thức thành công!", "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            Button btnClose = new Button { Text = "Đóng", Location = new Point(854, 11), Size = new Size(80, 34), BackColor = Color.White, ForeColor = Color.FromArgb(51, 65, 85), FlatStyle = FlatStyle.Flat };
            btnClose.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnClose.Click += (s, e) => this.Close();
            ftr.Controls.AddRange(new Control[] { btnReset, lblFooterStatus, btnSave, btnClose });
            this.Controls.Add(ftr);

            // Left Sidebar
            Panel sidebar = new Panel { Dock = DockStyle.Left, Width = 230, BackColor = Color.FromArgb(241, 245, 249) };
            sidebar.Controls.Add(new Label { Text = "CẤU HÌNH THỂ THỨC", Location = new Point(16, 16), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8f, FontStyle.Bold) });
            AddNavItem(sidebar, 0, "Cấu hình quy chuẩn", 38);

            sidebar.Controls.Add(new Label { Text = "QUẢN LÝ MẪU THỂ THỨC", Location = new Point(16, 108), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8f, FontStyle.Bold) });
            AddNavItem(sidebar, 1, "Quốc hiệu & Cơ quan", 130);
            AddNavItem(sidebar, 2, "Tên loại & Trích yếu", 190);
            AddNavItem(sidebar, 3, "Chữ ký & Nơi nhận", 250);
            AddNavItem(sidebar, 4, "Ký hiệu Mật & Khẩn cấp", 310);

            sidebar.Controls.Add(new Label { Text = "TIỆN ÍCH HỆ THỐNG", Location = new Point(16, 384), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8f, FontStyle.Bold) });
            AddNavItem(sidebar, 5, "Nhắc nhở nghỉ ngơi", 406);
            this.Controls.Add(sidebar);

            // Right Content Area
            pnlContentHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 250, 252), AutoScroll = true, Padding = new Padding(16) };
            this.Controls.Add(pnlContentHost);
            pnlContentHost.BringToFront();

            UpdateNavCounts();
            SelectTab(activeTab);
        }

        private void AddNavItem(Panel sidebar, int idx, string title, int y)
        {
            Panel p = new Panel { Location = new Point(10, y), Size = new Size(210, 52), Cursor = Cursors.Hand };
            Label lTitle = new Label { Text = title, Location = new Point(14, 7), AutoSize = true, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold), ForeColor = Color.FromArgb(15, 23, 42) };
            Label lSub = new Label { Text = "", Location = new Point(14, 28), AutoSize = true, Font = new Font("Segoe UI", 8.2f), ForeColor = Color.FromArgb(37, 99, 235) };
            EventHandler onClick = (s, e) => SelectTab(idx);
            p.Click += onClick; lTitle.Click += onClick; lSub.Click += onClick;
            p.Controls.Add(lTitle);
            p.Controls.Add(lSub);
            navButtons[idx] = p;
            navSubLabels[idx] = lSub;
            sidebar.Controls.Add(p);
        }

        private void UpdateNavCounts()
        {
            navSubLabels[0].Text = cfg.Standards.Count + " quy chuẩn";
            navSubLabels[1].Text = cfg.Agencies.Count + " mẫu";
            navSubLabels[2].Text = cfg.DocTypes.Count + " mẫu";
            navSubLabels[3].Text = cfg.Signatures.Count + " mẫu";
            navSubLabels[4].Text = cfg.SecurityMarks.Count + " mẫu";
            navSubLabels[5].Text = cfg.RestReminder.Enabled ? (cfg.RestReminder.IntervalMinutes + " phút/lần") : "Đang tắt";
        }

        private void SelectTab(int idx)
        {
            activeTab = idx;
            for (int i = 0; i < 6; i++)
            {
                navButtons[i].BackColor = (i == idx) ? Color.White : Color.Transparent;
            }
            RenderActiveTab();
        }

        private void RenderActiveTab()
        {
            pnlContentHost.Controls.Clear();
            if (activeTab == 0) RenderStandardsTab();
            else if (activeTab == 1) RenderAgenciesTab();
            else if (activeTab == 2) RenderDocTypesTab();
            else if (activeTab == 3) RenderSignaturesTab();
            else if (activeTab == 4) RenderSecurityMarksTab();
            else RenderRestReminderTab();
        }

        private Panel CreateWhiteCard(int x, int y, int w, int h, string heading)
        {
            Panel c = new Panel { Location = new Point(x, y), Size = new Size(w, h), BackColor = Color.White };
            c.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(p, 0, 0, c.Width - 1, c.Height - 1); };
            if (!string.IsNullOrEmpty(heading))
                c.Controls.Add(new Label { Text = "▌ " + heading, Location = new Point(14, 12), AutoSize = true, ForeColor = Color.FromArgb(37, 99, 235), Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) });
            return c;
        }

        // TAB 0: Cấu hình quy chuẩn (NĐ 30, HD 05-HD/VPTW, CAND)
        private void RenderStandardsTab()
        {
            Button btnAdd = new Button { Text = "+ Thêm quy chuẩn", Location = new Point(16, 16), Size = new Size(185, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnAdd.FlatAppearance.BorderSize = 0;
            Button btnDel = new Button { Text = "🗑", Location = new Point(208, 16), Size = new Size(36, 36), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            ListBox lb = new ListBox { Location = new Point(16, 62), Size = new Size(228, 420), Font = new Font("Segoe UI", 9.5f), BorderStyle = BorderStyle.FixedSingle };
            Panel detail = CreateWhiteCard(260, 16, 430, 466, "THÔNG TIN QUY CHUẨN THỂ THỨC");

            TextBox txtName = AddLabeledText(detail, "Tên quy chuẩn:", 16, 42, 190);
            TextBox txtCat = AddLabeledText(detail, "Loại hình thể thức:", 218, 42, 194);
            TextBox txtBasis = AddLabeledText(detail, "Mô tả / Căn cứ pháp lý:", 16, 102, 396);
            NumericUpDown nT = AddSmallNum(detail, "Lề trên (mm):", 16, 174, 20);
            NumericUpDown nB = AddSmallNum(detail, "Lề dưới (mm):", 118, 174, 20);
            NumericUpDown nL = AddSmallNum(detail, "Lề trái (mm):", 220, 174, 30);
            NumericUpDown nR = AddSmallNum(detail, "Lề phải (mm):", 322, 174, 15);
            NumericUpDown nG = AddSmallNum(detail, "Lề gáy (Gutter mm):", 16, 236, 0);
            CheckBox chkMirror = new CheckBox { Text = "In hai mặt đối xứng (Mirror Margins)", Location = new Point(160, 256), AutoSize = true };
            detail.Controls.Add(chkMirror);
            TextBox txtFont = AddLabeledText(detail, "Phông chữ mặc định:", 16, 304, 220);
            NumericUpDown nSize = AddSmallNum(detail, "Cỡ chữ (pt):", 252, 304, 14);

            bool loading = false;
            Action bindList = () =>
            {
                lb.Items.Clear();
                foreach (var st in cfg.Standards) lb.Items.Add(st.Name);
                if (lb.Items.Count > 0) lb.SelectedIndex = 0;
            };

            lb.SelectedIndexChanged += (s, e) =>
            {
                if (lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Standards.Count) return;
                loading = true;
                var st = cfg.Standards[lb.SelectedIndex];
                txtName.Text = st.Name ?? "";
                txtCat.Text = st.Category ?? "";
                txtBasis.Text = st.LegalBasis ?? "";
                nT.Value = (decimal)Math.Max(0, Math.Min(100, st.TopMm));
                nB.Value = (decimal)Math.Max(0, Math.Min(100, st.BottomMm));
                nL.Value = (decimal)Math.Max(0, Math.Min(100, st.LeftMm));
                nR.Value = (decimal)Math.Max(0, Math.Min(100, st.RightMm));
                nG.Value = (decimal)Math.Max(0, Math.Min(100, st.GutterMm));
                chkMirror.Checked = st.MirrorMargins;
                txtFont.Text = string.IsNullOrEmpty(st.FontName) ? "Times New Roman" : st.FontName;
                nSize.Value = (decimal)Math.Max(8, Math.Min(24, st.FontSizePt > 0 ? st.FontSizePt : 14));
                loading = false;
            };

            EventHandler sync = (s, e) =>
            {
                if (loading || lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Standards.Count) return;
                var st = cfg.Standards[lb.SelectedIndex];
                st.Name = txtName.Text; st.Category = txtCat.Text; st.LegalBasis = txtBasis.Text;
                st.TopMm = (double)nT.Value; st.BottomMm = (double)nB.Value; st.LeftMm = (double)nL.Value; st.RightMm = (double)nR.Value;
                st.GutterMm = (double)nG.Value; st.MirrorMargins = chkMirror.Checked;
                st.FontName = txtFont.Text; st.FontSizePt = (double)nSize.Value;
            };
            txtName.TextChanged += sync; txtCat.TextChanged += sync; txtBasis.TextChanged += sync;
            nT.ValueChanged += sync; nB.ValueChanged += sync; nL.ValueChanged += sync; nR.ValueChanged += sync;
            nG.ValueChanged += sync; chkMirror.CheckedChanged += sync; txtFont.TextChanged += sync; nSize.ValueChanged += sync;

            btnAdd.Click += (s, e) =>
            {
                cfg.Standards.Add(new StandardRulePreset { Name = "Quy chuẩn mới", Category = "Tùy chỉnh", LegalBasis = "Văn bản nội bộ", PaperSize = "A4", TopMm = 20, BottomMm = 20, LeftMm = 30, RightMm = 15, FontName = "Times New Roman", FontSizePt = 14 });
                UpdateNavCounts();
                bindList();
                lb.SelectedIndex = lb.Items.Count - 1;
            };
            btnDel.Click += (s, e) =>
            {
                if (lb.SelectedIndex >= 0 && cfg.Standards.Count > 1)
                {
                    cfg.Standards.RemoveAt(lb.SelectedIndex);
                    UpdateNavCounts();
                    bindList();
                }
            };

            pnlContentHost.Controls.AddRange(new Control[] { btnAdd, btnDel, lb, detail });
            bindList();
        }

        // TAB 1: Quốc hiệu & Cơ quan
        private void RenderAgenciesTab()
        {
            Button btnAdd = new Button { Text = "+ Thêm khối mới", Location = new Point(16, 16), Size = new Size(185, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnAdd.FlatAppearance.BorderSize = 0;
            Button btnDel = new Button { Text = "🗑", Location = new Point(208, 16), Size = new Size(36, 36), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            ListBox lb = new ListBox { Location = new Point(16, 62), Size = new Size(228, 420), Font = new Font("Segoe UI", 9.5f), BorderStyle = BorderStyle.FixedSingle };
            Panel detail = CreateWhiteCard(260, 16, 430, 466, "CHI TIẾT KHỐI QUỐC HIỆU & CƠ QUAN");

            TextBox txtLabel = AddLabeledText(detail, "Tên khối (gợi nhớ):", 16, 42, 190);
            TextBox txtStd = AddLabeledText(detail, "Quy chuẩn thể thức áp dụng:", 218, 42, 194);
            TextBox txtParent = AddLabeledText(detail, "Cơ quan / Đơn vị cấp trên (VD: UBND TỈNH..., CÔNG AN...):", 16, 102, 396);
            TextBox txtIssuing = AddLabeledText(detail, "Cơ quan / Đơn vị ban hành (VD: PHÒNG KỸ THUẬT HÌNH SỰ):", 16, 162, 396);
            TextBox txtLoc = AddLabeledText(detail, "Địa danh hành chính (VD: Hà Nội, Đà Nẵng):", 16, 222, 190);
            TextBox txtAbbr = AddLabeledText(detail, "Ký hiệu số văn bản viết tắt:", 218, 222, 194);
            CheckBox chkDef = new CheckBox { Text = "Dùng khối cơ quan này làm mặc định khi chèn tiêu đề tự động", Location = new Point(16, 290), Size = new Size(396, 36), Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(37, 99, 235) };
            detail.Controls.Add(chkDef);

            bool loading = false;
            Action bindList = () =>
            {
                lb.Items.Clear();
                foreach (var a in cfg.Agencies) lb.Items.Add(a.Label);
                if (lb.Items.Count > 0) lb.SelectedIndex = 0;
            };
            lb.SelectedIndexChanged += (s, e) =>
            {
                if (lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Agencies.Count) return;
                loading = true;
                var a = cfg.Agencies[lb.SelectedIndex];
                txtLabel.Text = a.Label ?? ""; txtStd.Text = a.StandardType ?? "";
                txtParent.Text = a.ParentOrg ?? ""; txtIssuing.Text = a.IssuingOrg ?? "";
                txtLoc.Text = a.Location ?? ""; txtAbbr.Text = a.CodeAbbr ?? "";
                chkDef.Checked = a.IsDefault;
                loading = false;
            };
            EventHandler sync = (s, e) =>
            {
                if (loading || lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Agencies.Count) return;
                var a = cfg.Agencies[lb.SelectedIndex];
                a.Label = txtLabel.Text; a.StandardType = txtStd.Text; a.ParentOrg = txtParent.Text;
                a.IssuingOrg = txtIssuing.Text; a.Location = txtLoc.Text; a.CodeAbbr = txtAbbr.Text;
                if (chkDef.Checked) { foreach (var x in cfg.Agencies) x.IsDefault = false; a.IsDefault = true; }
                else a.IsDefault = false;
            };
            txtLabel.TextChanged += sync; txtStd.TextChanged += sync; txtParent.TextChanged += sync;
            txtIssuing.TextChanged += sync; txtLoc.TextChanged += sync; txtAbbr.TextChanged += sync; chkDef.CheckedChanged += sync;

            btnAdd.Click += (s, e) =>
            {
                cfg.Agencies.Add(new AgencyBlockPreset { Label = "Khối cơ quan mới", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", ParentOrg = "CƠ QUAN CẤP TRÊN", IssuingOrg = "ĐƠN VỊ BAN HÀNH", Location = "Đà Nẵng", CodeAbbr = "VP" });
                UpdateNavCounts(); bindList(); lb.SelectedIndex = lb.Items.Count - 1;
            };
            btnDel.Click += (s, e) =>
            {
                if (lb.SelectedIndex >= 0 && cfg.Agencies.Count > 1) { cfg.Agencies.RemoveAt(lb.SelectedIndex); UpdateNavCounts(); bindList(); }
            };

            pnlContentHost.Controls.AddRange(new Control[] { btnAdd, btnDel, lb, detail });
            bindList();
        }

        // TAB 2: Tên loại & Trích yếu
        private void RenderDocTypesTab()
        {
            Button btnAdd = new Button { Text = "+ Thêm khối mới", Location = new Point(16, 16), Size = new Size(185, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnAdd.FlatAppearance.BorderSize = 0;
            Button btnDel = new Button { Text = "🗑", Location = new Point(208, 16), Size = new Size(36, 36), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            ListBox lb = new ListBox { Location = new Point(16, 62), Size = new Size(228, 420), Font = new Font("Segoe UI", 9.2f), BorderStyle = BorderStyle.FixedSingle };
            Panel detail = CreateWhiteCard(260, 16, 430, 466, "CHI TIẾT KHỐI TÊN LOẠI & TRÍCH YẾU");

            TextBox txtLabel = AddLabeledText(detail, "Tên khối (gợi nhớ):", 16, 42, 190);
            TextBox txtStd = AddLabeledText(detail, "Quy chuẩn thể thức áp dụng:", 218, 42, 194);
            TextBox txtType = AddLabeledText(detail, "Tên loại văn bản (Viết HOA đứng đậm, e.g. QUYẾT ĐỊNH):", 16, 102, 396);
            detail.Controls.Add(new Label { Text = "Nội dung trích yếu mẫu (VD: Về việc ban hành..., Kính gửi: ...):", Location = new Point(16, 162), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            TextBox txtSample = new TextBox { Location = new Point(16, 184), Size = new Size(396, 110), Multiline = true, ScrollBars = ScrollBars.Vertical };
            detail.Controls.Add(txtSample);

            bool loading = false;
            Action bindList = () => { lb.Items.Clear(); foreach (var d in cfg.DocTypes) lb.Items.Add(d.Label); if (lb.Items.Count > 0) lb.SelectedIndex = 0; };
            lb.SelectedIndexChanged += (s, e) =>
            {
                if (lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.DocTypes.Count) return;
                loading = true;
                var d = cfg.DocTypes[lb.SelectedIndex];
                txtLabel.Text = d.Label ?? ""; txtStd.Text = d.StandardType ?? "";
                txtType.Text = d.DocTypeName ?? ""; txtSample.Text = d.SummarySample ?? "";
                loading = false;
            };
            EventHandler sync = (s, e) =>
            {
                if (loading || lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.DocTypes.Count) return;
                var d = cfg.DocTypes[lb.SelectedIndex];
                d.Label = txtLabel.Text; d.StandardType = txtStd.Text; d.DocTypeName = txtType.Text; d.SummarySample = txtSample.Text;
            };
            txtLabel.TextChanged += sync; txtStd.TextChanged += sync; txtType.TextChanged += sync; txtSample.TextChanged += sync;

            btnAdd.Click += (s, e) =>
            {
                cfg.DocTypes.Add(new DocTypePreset { Label = (cfg.DocTypes.Count + 1) + ". Loại văn bản mới", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", DocTypeName = "THÔNG BÁO", SummarySample = "Về việc ..." });
                UpdateNavCounts(); bindList(); lb.SelectedIndex = lb.Items.Count - 1;
            };
            btnDel.Click += (s, e) => { if (lb.SelectedIndex >= 0 && cfg.DocTypes.Count > 1) { cfg.DocTypes.RemoveAt(lb.SelectedIndex); UpdateNavCounts(); bindList(); } };

            pnlContentHost.Controls.AddRange(new Control[] { btnAdd, btnDel, lb, detail });
            bindList();
        }

        // TAB 3: Chữ ký & Nơi nhận (có Cấp bậc hàm CAND/QĐ)
        private void RenderSignaturesTab()
        {
            Button btnAdd = new Button { Text = "+ Thêm khối mới", Location = new Point(16, 16), Size = new Size(185, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnAdd.FlatAppearance.BorderSize = 0;
            Button btnDel = new Button { Text = "🗑", Location = new Point(208, 16), Size = new Size(36, 36), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            ListBox lb = new ListBox { Location = new Point(16, 62), Size = new Size(228, 420), Font = new Font("Segoe UI", 9.5f), BorderStyle = BorderStyle.FixedSingle };
            Panel detail = CreateWhiteCard(260, 16, 430, 466, "CHI TIẾT MẪU CHỮ KÝ & NƠI NHẬN");

            TextBox txtLabel = AddLabeledText(detail, "Tên mẫu gợi nhớ:", 16, 40, 190);
            TextBox txtStd = AddLabeledText(detail, "Quy chuẩn áp dụng:", 218, 40, 194);
            TextBox txtAuth = AddLabeledText(detail, "Quyền hạn ký (VD: TM. UBND, KT. GIÁM ĐỐC, T/M ĐẢNG ỦY):", 16, 96, 396);
            TextBox txtPos = AddLabeledText(detail, "Chức vụ người ký (CHỦ TỊCH, BÍ THƯ):", 16, 152, 190);
            TextBox txtRank = AddLabeledText(detail, "Cấp bậc hàm CAND/QĐ (nếu có):", 218, 152, 194);
            TextBox txtName = AddLabeledText(detail, "Họ và tên người ký (Đứng đậm 14pt, VD: Nguyễn Văn A):", 16, 208, 396);
            detail.Controls.Add(new Label { Text = "Danh mục nơi nhận đi kèm (Cột bên trái bảng kết thúc):", Location = new Point(16, 264), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            TextBox txtRecip = new TextBox { Location = new Point(16, 286), Size = new Size(396, 86), Multiline = true, ScrollBars = ScrollBars.Vertical };
            detail.Controls.Add(txtRecip);
            CheckBox chkDef = new CheckBox { Text = "Dùng làm mẫu Chữ ký & Nơi nhận mặc định khi chèn", Location = new Point(16, 384), Size = new Size(396, 30), Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Color.FromArgb(37, 99, 235) };
            detail.Controls.Add(chkDef);

            bool loading = false;
            Action bindList = () => { lb.Items.Clear(); foreach (var sg in cfg.Signatures) lb.Items.Add(sg.Label); if (lb.Items.Count > 0) lb.SelectedIndex = 0; };
            lb.SelectedIndexChanged += (s, e) =>
            {
                if (lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Signatures.Count) return;
                loading = true;
                var sg = cfg.Signatures[lb.SelectedIndex];
                txtLabel.Text = sg.Label ?? ""; txtStd.Text = sg.StandardType ?? "";
                txtAuth.Text = sg.Authority ?? ""; txtPos.Text = sg.Position ?? "";
                txtRank.Text = sg.MilitaryRank ?? ""; txtName.Text = sg.SignerName ?? "";
                txtRecip.Text = sg.Recipients ?? ""; chkDef.Checked = sg.IsDefault;
                loading = false;
            };
            EventHandler sync = (s, e) =>
            {
                if (loading || lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.Signatures.Count) return;
                var sg = cfg.Signatures[lb.SelectedIndex];
                sg.Label = txtLabel.Text; sg.StandardType = txtStd.Text; sg.Authority = txtAuth.Text;
                sg.Position = txtPos.Text; sg.MilitaryRank = txtRank.Text; sg.SignerName = txtName.Text;
                sg.Recipients = txtRecip.Text;
                if (chkDef.Checked) { foreach (var x in cfg.Signatures) x.IsDefault = false; sg.IsDefault = true; }
                else sg.IsDefault = false;
            };
            txtLabel.TextChanged += sync; txtStd.TextChanged += sync; txtAuth.TextChanged += sync;
            txtPos.TextChanged += sync; txtRank.TextChanged += sync; txtName.TextChanged += sync;
            txtRecip.TextChanged += sync; chkDef.CheckedChanged += sync;

            btnAdd.Click += (s, e) =>
            {
                cfg.Signatures.Add(new SignaturePreset { Label = "Mẫu chữ ký mới", StandardType = "Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", Authority = "TM. ỦY BAN NHÂN DÂN", Position = "CHỦ TỊCH", MilitaryRank = "", SignerName = "Nguyễn Văn B", Recipients = "- Như trên;\r\n- Lưu: VT." });
                UpdateNavCounts(); bindList(); lb.SelectedIndex = lb.Items.Count - 1;
            };
            btnDel.Click += (s, e) => { if (lb.SelectedIndex >= 0 && cfg.Signatures.Count > 1) { cfg.Signatures.RemoveAt(lb.SelectedIndex); UpdateNavCounts(); bindList(); } };

            pnlContentHost.Controls.AddRange(new Control[] { btnAdd, btnDel, lb, detail });
            bindList();
        }

        // TAB 4: Ký hiệu Mật & Khẩn cấp
        private void RenderSecurityMarksTab()
        {
            Button btnAdd = new Button { Text = "+ Thêm dấu mới", Location = new Point(16, 16), Size = new Size(185, 36), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
            btnAdd.FlatAppearance.BorderSize = 0;
            Button btnDel = new Button { Text = "🗑", Location = new Point(208, 16), Size = new Size(36, 36), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            ListBox lb = new ListBox { Location = new Point(16, 62), Size = new Size(228, 420), Font = new Font("Segoe UI", 9.5f), BorderStyle = BorderStyle.FixedSingle };
            Panel detail = CreateWhiteCard(260, 16, 430, 320, "CHI TIẾT KÝ HIỆU MẬT & KHẨN CẤP");

            TextBox txtLabel = AddLabeledText(detail, "Tên hiển thị:", 16, 44, 190);
            TextBox txtCat = AddLabeledText(detail, "Nhóm (Độ khẩn / Độ mật):", 218, 44, 194);
            TextBox txtText = AddLabeledText(detail, "Nội dung đóng dấu (VD: HỎA TỐC, TỐI MẬT):", 16, 108, 396);
            TextBox txtPos = AddLabeledText(detail, "Vị trí đặt dấu mặc định:", 16, 172, 396);

            bool loading = false;
            Action bindList = () => { lb.Items.Clear(); foreach (var m in cfg.SecurityMarks) lb.Items.Add(m.Label); if (lb.Items.Count > 0) lb.SelectedIndex = 0; };
            lb.SelectedIndexChanged += (s, e) =>
            {
                if (lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.SecurityMarks.Count) return;
                loading = true;
                var m = cfg.SecurityMarks[lb.SelectedIndex];
                txtLabel.Text = m.Label ?? ""; txtCat.Text = m.Category ?? ""; txtText.Text = m.MarkText ?? ""; txtPos.Text = m.Position ?? "";
                loading = false;
            };
            EventHandler sync = (s, e) =>
            {
                if (loading || lb.SelectedIndex < 0 || lb.SelectedIndex >= cfg.SecurityMarks.Count) return;
                var m = cfg.SecurityMarks[lb.SelectedIndex];
                m.Label = txtLabel.Text; m.Category = txtCat.Text; m.MarkText = txtText.Text; m.Position = txtPos.Text;
            };
            txtLabel.TextChanged += sync; txtCat.TextChanged += sync; txtText.TextChanged += sync; txtPos.TextChanged += sync;

            btnAdd.Click += (s, e) =>
            {
                cfg.SecurityMarks.Add(new SecurityMarkPreset { Label = "TUYỆT MẬT", Category = "Độ mật", MarkText = "TUYỆT MẬT", Position = "Góc trên trái trang" });
                UpdateNavCounts(); bindList(); lb.SelectedIndex = lb.Items.Count - 1;
            };
            btnDel.Click += (s, e) => { if (lb.SelectedIndex >= 0 && cfg.SecurityMarks.Count > 1) { cfg.SecurityMarks.RemoveAt(lb.SelectedIndex); UpdateNavCounts(); bindList(); } };

            pnlContentHost.Controls.AddRange(new Control[] { btnAdd, btnDel, lb, detail });
            bindList();
        }

        // TAB 5: Nhắc nhở nghỉ ngơi & Bảo vệ Sức khỏe (Image 7)
        private void RenderRestReminderTab()
        {
            Panel cTop = CreateWhiteCard(16, 16, 674, 86, null);
            cTop.Controls.Add(new Label { Text = "💙", Location = new Point(18, 22), AutoSize = true, Font = new Font("Segoe UI", 20f), ForeColor = Color.FromArgb(37, 99, 235) });
            cTop.Controls.Add(new Label { Text = "Chế độ Nhắc nhở Nghỉ ngơi & Bảo vệ Sức khỏe", Location = new Point(72, 16), AutoSize = true, Font = new Font("Segoe UI", 10.5f, FontStyle.Bold), ForeColor = Color.FromArgb(15, 23, 42) });
            cTop.Controls.Add(new Label { Text = "Tự động nhắc nhở vận động, uống nước và thư giãn mắt sau mỗi khoảng thời gian soạn thảo liên tục", Location = new Point(72, 42), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139) });
            CheckBox chkEn = new CheckBox { Text = "Bật chức năng", Location = new Point(542, 30), AutoSize = true, Checked = cfg.RestReminder.Enabled, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Color.FromArgb(37, 99, 235) };
            chkEn.CheckedChanged += (s, e) => { cfg.RestReminder.Enabled = chkEn.Checked; UpdateNavCounts(); };
            cTop.Controls.Add(chkEn);

            Panel cSet = CreateWhiteCard(16, 116, 674, 198, "THIẾT LẬP THỜI GIAN NHẮC NHỞ");
            cSet.Controls.Add(new Label { Text = "Thời gian làm việc liên tục:", Location = new Point(18, 46), AutoSize = true, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) });
            cSet.Controls.Add(new Label { Text = "Nhắc sau khi gõ văn bản liên tục", Location = new Point(18, 66), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.4f) });
            ComboBox cboWork = new ComboBox { Location = new Point(250, 48), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            cboWork.Items.AddRange(new object[] { "30 phút", "45 phút", "60 phút (Khuyến nghị)", "90 phút" });
            cboWork.SelectedIndex = cfg.RestReminder.IntervalMinutes == 30 ? 0 : (cfg.RestReminder.IntervalMinutes == 45 ? 1 : (cfg.RestReminder.IntervalMinutes == 90 ? 3 : 2));
            cboWork.SelectedIndexChanged += (s, e) =>
            {
                int[] vals = { 30, 45, 60, 90 };
                cfg.RestReminder.IntervalMinutes = vals[cboWork.SelectedIndex];
                UpdateNavCounts();
            };
            cSet.Controls.Add(cboWork);

            cSet.Controls.Add(new Label { Text = "Thời gian báo lại (Snooze):", Location = new Point(18, 98), AutoSize = true, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) });
            cSet.Controls.Add(new Label { Text = "Khoảng cách khi bấm 'Báo lại'", Location = new Point(18, 118), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.4f) });
            ComboBox cboSnooze = new ComboBox { Location = new Point(250, 100), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            cboSnooze.Items.AddRange(new object[] { "10 phút", "15 phút", "30 phút (Khuyến nghị)" });
            cboSnooze.SelectedIndex = cfg.RestReminder.SnoozeMinutes == 10 ? 0 : (cfg.RestReminder.SnoozeMinutes == 15 ? 1 : 2);
            cboSnooze.SelectedIndexChanged += (s, e) => { int[] vals = { 10, 15, 30 }; cfg.RestReminder.SnoozeMinutes = vals[cboSnooze.SelectedIndex]; };
            cSet.Controls.Add(cboSnooze);

            cSet.Controls.Add(new Label { Text = "Kiểm tra giao diện thông báo", Location = new Point(18, 150), AutoSize = true, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) });
            cSet.Controls.Add(new Label { Text = "Bật cửa sổ thông báo xem thử ở góc phải màn hình", Location = new Point(18, 170), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139), Font = new Font("Segoe UI", 8.4f) });
            Button btnPreview = new Button { Text = "☕ Xem thử thông báo ngay", Location = new Point(458, 150), Size = new Size(196, 34), BackColor = Color.White, ForeColor = Color.FromArgb(30, 41, 59), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            btnPreview.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnPreview.Click += (s, e) =>
            {
                RestReminderToastForm toast = new RestReminderToastForm(cfg.RestReminder.IntervalMinutes, cfg.RestReminder.SnoozeMinutes);
                toast.Show();
            };
            cSet.Controls.Add(btnPreview);

            Panel cInfo = CreateWhiteCard(16, 328, 674, 150, "Cơ chế hoạt động thông minh (100% Offline & An toàn):");
            cInfo.BackColor = Color.FromArgb(248, 250, 252);
            cInfo.Controls.Add(new Label
            {
                Location = new Point(18, 42),
                Size = new Size(638, 96),
                ForeColor = Color.FromArgb(51, 65, 85),
                Text =
                    "• Không cướp quyền bàn phím (No Focus-Stealing): Cửa sổ bật lên êm ái ở góc phải màn hình, bạn vẫn tiếp tục gõ mà không mất con trỏ soạn thảo.\r\n" +
                    "• Tự động phát hiện rời bàn làm việc: Nếu rời máy quá 3 phút (đi họp, ăn trưa...), bộ đếm tự tạm dừng hoặc làm mới.\r\n" +
                    "• Tự động giới hạn ca làm việc: Tối đa 2 lần nhắc trong một ca làm việc liên tục để tránh làm phiền khi xử lý văn bản khẩn."
            });

            pnlContentHost.Controls.AddRange(new Control[] { cTop, cSet, cInfo });
        }

        private TextBox AddLabeledText(Panel parent, string label, int x, int y, int width)
        {
            parent.Controls.Add(new Label { Text = label, Location = new Point(x, y), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            TextBox tb = new TextBox { Location = new Point(x, y + 20), Width = width };
            parent.Controls.Add(tb);
            return tb;
        }

        private NumericUpDown AddSmallNum(Panel parent, string label, int x, int y, decimal val)
        {
            parent.Controls.Add(new Label { Text = label, Location = new Point(x, y), AutoSize = true, ForeColor = Color.FromArgb(71, 85, 105) });
            NumericUpDown n = new NumericUpDown { Location = new Point(x, y + 20), Width = 92, Minimum = 0, Maximum = 100, Value = val };
            parent.Controls.Add(n);
            return n;
        }
    }

    // =========================================================================
    // 6. DIALOG 3: CÁC PHÍM TẮT TRONG WORD (IMAGE 8 + LIVE KEYBOARD HIGHLIGHT)
    // =========================================================================
    public class WordShortcutsForm : Form
    {
        private class ShortcutRow
        {
            public string Keys;
            public string Desc;
            public string MatchKey;
            public Panel RowPanel;
        }

        private List<ShortcutRow> rows = new List<ShortcutRow>();

        public WordShortcutsForm()
        {
            this.Text = "Các phím tắt trong Word — VBTM Pro 3.6";
            this.Size = new Size(920, 620);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9.2f);
            this.KeyPreview = true;

            Label lblTitle = new Label
            {
                Text = "Các phím tắt trong Word",
                Dock = DockStyle.Top,
                Height = 52,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(29, 78, 216),
                Font = new Font("Segoe UI", 14.5f, FontStyle.Bold)
            };
            this.Controls.Add(lblTitle);

            // Left blue Word badge
            Panel badge = new Panel { Location = new Point(18, 240), Size = new Size(76, 82), BackColor = Color.FromArgb(37, 99, 235) };
            badge.Controls.Add(new Label { Text = "W", Location = new Point(0, 10), Size = new Size(76, 42), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.White, Font = new Font("Segoe UI", 24f, FontStyle.Bold) });
            badge.Controls.Add(new Label { Text = "Word", Location = new Point(0, 52), Size = new Size(76, 22), TextAlign = ContentAlignment.MiddleCenter, ForeColor = Color.FromArgb(219, 234, 254), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) });
            this.Controls.Add(badge);

            string[,] leftItems = new string[,] {
                { "Ctrl + O", "Mở tài liệu mới", "O" },
                { "Ctrl + A", "Chọn tất cả", "A" },
                { "Ctrl + B", "In đậm văn bản", "B" },
                { "F12", "Lưu thành (Save As)", "F12" },
                { "Ctrl + P", "In tài liệu", "P" },
                { "Ctrl + W", "Đóng tài liệu", "W" },
                { "Ctrl + Z", "Hoàn tác", "Z" },
                { "Ctrl + Y", "Làm lại", "Y" },
                { "Ctrl + C", "Sao chép", "C" },
                { "Ctrl + X", "Cắt", "X" },
                { "Ctrl + V", "Dán", "V" },
                { "Ctrl + L", "Căn lề trái", "L" },
                { "Ctrl + U", "Gạch chân", "U" },
                { "Ctrl + T", "Thụt lề treo (Hanging indent)", "T" },
                { "Ctrl + N", "Tạo tài liệu mới", "N" },
                { "Ctrl + I", "In nghiêng", "I" }
            };

            string[,] rightItems = new string[,] {
                { "Ctrl + S", "Lưu tài liệu", "S" },
                { "Ctrl + K", "Chèn siêu liên kết (Hyperlink)", "K" },
                { "Ctrl + Q", "Căn lề trái và xóa thụt đầu dòng", "Q" },
                { "Ctrl + E", "Căn giữa", "E" },
                { "Ctrl + R", "Căn lề phải", "R" },
                { "Ctrl + J", "Căn đều hai bên", "J" },
                { "Ctrl + 1", "Giãn dòng đơn (1.0)", "D1" },
                { "Ctrl + 2", "Giãn dòng kép (2.0)", "D2" },
                { "Ctrl + 5", "Giãn dòng 1.5", "D5" },
                { "Ctrl + M", "Tăng thụt đầu dòng", "M" },
                { "Ctrl + Shift + M", "Giảm thụt đầu dòng", "M" },
                { "Ctrl + D", "Thay đổi cỡ chữ (Font size)", "D" },
                { "Ctrl + Shift + >", "Tăng cỡ chữ", "OemPeriod" },
                { "Ctrl + Shift + <", "Giảm cỡ chữ", "Oemcomma" },
                { "Ctrl + Home", "Về đầu tài liệu", "Home" },
                { "Ctrl + End", "Về cuối tài liệu", "End" }
            };

            BuildColumnCard(110, 62, 382, leftItems);
            BuildColumnCard(506, 62, 382, rightItems);

            // Footer
            Panel ftr = new Panel { Dock = DockStyle.Bottom, Height = 52, BackColor = Color.FromArgb(248, 250, 252) };
            ftr.Controls.Add(new Label { Text = "Mẹo: Bấm thử một phím bất kỳ trên bàn phím để tra cứu nhanh vị trí phím tắt tương ứng!", Location = new Point(18, 16), AutoSize = true, ForeColor = Color.FromArgb(100, 116, 139) });
            Button btnClose = new Button { Text = "Đóng", Location = new Point(792, 9), Size = new Size(96, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => this.Close();
            ftr.Controls.Add(btnClose);
            this.Controls.Add(ftr);

            this.KeyDown += (s, e) =>
            {
                string k = e.KeyCode.ToString();
                foreach (var r in rows)
                {
                    bool hit = string.Equals(r.MatchKey, k, StringComparison.OrdinalIgnoreCase);
                    r.RowPanel.BackColor = hit ? Color.FromArgb(219, 234, 254) : Color.Transparent;
                }
            };
        }

        private void BuildColumnCard(int x, int y, int w, string[,] items)
        {
            Panel card = new Panel { Location = new Point(x, y), Size = new Size(w, 468), BackColor = Color.FromArgb(250, 251, 253) };
            card.Paint += (s, e) => { using (Pen p = new Pen(Color.FromArgb(226, 232, 240))) e.Graphics.DrawRectangle(p, 0, 0, card.Width - 1, card.Height - 1); };
            int rowH = 28;
            for (int i = 0; i < items.GetLength(0); i++)
            {
                Panel rp = new Panel { Location = new Point(8, 8 + i * rowH), Size = new Size(w - 16, rowH - 2), BackColor = Color.Transparent };
                Label lblKey = new Label
                {
                    Text = items[i, 0],
                    Location = new Point(6, 3),
                    Size = new Size(124, 20),
                    BackColor = Color.White,
                    BorderStyle = BorderStyle.FixedSingle,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                    ForeColor = Color.FromArgb(30, 41, 59)
                };
                Label lblDesc = new Label
                {
                    Text = ": " + items[i, 1],
                    Location = new Point(138, 4),
                    Size = new Size(w - 160, 20),
                    ForeColor = Color.FromArgb(51, 65, 85),
                    Font = new Font("Segoe UI", 9f)
                };
                rp.Controls.Add(lblKey);
                rp.Controls.Add(lblDesc);
                card.Controls.Add(rp);
                rows.Add(new ShortcutRow { Keys = items[i, 0], Desc = items[i, 1], MatchKey = items[i, 2], RowPanel = rp });
            }
            this.Controls.Add(card);
        }
    }

    // =========================================================================
    // 7. DIALOG 4: QUẢN LÝ TỪ ĐIỂN CHÍNH TẢ (IMAGE 9 DROPDOWN)
    // =========================================================================
    public class SpellDictionaryManagerForm : Form
    {
        public SpellDictionaryManagerForm(string mode = "manage")
        {
            string selectedInWord = WordQuickUtilities.GetWordSelectedText();

            if (mode == "add-good")
            {
                PromptAddGoodWord(selectedInWord);
                return;
            }
            if (mode == "add-bad")
            {
                PromptAddBadWord(selectedInWord);
                return;
            }

            this.Text = "Quản lý Từ điển Chính tả — VBTM Pro 3.6";
            this.Size = new Size(640, 490);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.5f);

            TabControl tabs = new TabControl { Location = new Point(14, 14), Size = new Size(596, 386) };

            // Tab 1: Từ đúng
            TabPage tpGood = new TabPage("✓ Từ đúng (Bỏ qua khi soát lỗi)");
            tpGood.BackColor = Color.White;
            TextBox txtNewGood = new TextBox { Location = new Point(14, 16), Width = 390, Text = selectedInWord };
            Button btnAddGood = new Button { Text = "+ Thêm từ đúng", Location = new Point(414, 14), Size = new Size(154, 28), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            btnAddGood.FlatAppearance.BorderSize = 0;
            ListBox lbGood = new ListBox { Location = new Point(14, 54), Size = new Size(554, 246), BorderStyle = BorderStyle.FixedSingle };
            Button btnDelGood = new Button { Text = "🗑 Xóa từ đang chọn", Location = new Point(14, 310), Size = new Size(165, 32), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            Action reloadGood = () => { lbGood.Items.Clear(); foreach (string w in VbtmCustomDictStore.LoadGoodWords()) lbGood.Items.Add(w); };
            btnAddGood.Click += (s, e) => { if (!string.IsNullOrWhiteSpace(txtNewGood.Text)) { VbtmCustomDictStore.AddGoodWord(txtNewGood.Text); txtNewGood.Clear(); reloadGood(); } };
            btnDelGood.Click += (s, e) =>
            {
                if (lbGood.SelectedIndex >= 0)
                {
                    var list = VbtmCustomDictStore.LoadGoodWords();
                    list.Remove(Convert.ToString(lbGood.SelectedItem));
                    VbtmCustomDictStore.SaveGoodWords(list);
                    reloadGood();
                }
            };
            tpGood.Controls.AddRange(new Control[] { txtNewGood, btnAddGood, lbGood, btnDelGood });
            reloadGood();

            // Tab 2: Từ sai cố định
            TabPage tpBad = new TabPage("✕ Từ sai cố định (Tự động bắt lỗi)");
            tpBad.BackColor = Color.White;
            tpBad.Controls.Add(new Label { Text = "Từ sai:", Location = new Point(14, 18), AutoSize = true });
            TextBox txtBadW = new TextBox { Location = new Point(70, 15), Width = 165, Text = selectedInWord };
            tpBad.Controls.Add(new Label { Text = "Sửa thành:", Location = new Point(244, 18), AutoSize = true });
            TextBox txtBadS = new TextBox { Location = new Point(318, 15), Width = 150 };
            Button btnAddBad = new Button { Text = "+ Lưu", Location = new Point(476, 13), Size = new Size(92, 28), BackColor = Color.FromArgb(220, 38, 38), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9f, FontStyle.Bold) };
            btnAddBad.FlatAppearance.BorderSize = 0;

            ListBox lbBad = new ListBox { Location = new Point(14, 54), Size = new Size(554, 246), BorderStyle = BorderStyle.FixedSingle };
            Button btnDelBad = new Button { Text = "🗑 Xóa cặp đang chọn", Location = new Point(14, 310), Size = new Size(175, 32), BackColor = Color.White, ForeColor = Color.FromArgb(220, 38, 38), FlatStyle = FlatStyle.Flat };

            Action reloadBad = () =>
            {
                lbBad.Items.Clear();
                foreach (var kv in VbtmCustomDictStore.LoadBadWords()) lbBad.Items.Add(kv.Key + "  →  " + kv.Value);
            };
            btnAddBad.Click += (s, e) => { if (!string.IsNullOrWhiteSpace(txtBadW.Text) && !string.IsNullOrWhiteSpace(txtBadS.Text)) { VbtmCustomDictStore.AddBadWord(txtBadW.Text, txtBadS.Text); txtBadW.Clear(); txtBadS.Clear(); reloadBad(); } };
            btnDelBad.Click += (s, e) =>
            {
                if (lbBad.SelectedIndex >= 0)
                {
                    string item = Convert.ToString(lbBad.SelectedItem);
                    int idx = item.IndexOf("  →  ");
                    if (idx > 0)
                    {
                        var map = VbtmCustomDictStore.LoadBadWords();
                        map.Remove(item.Substring(0, idx));
                        VbtmCustomDictStore.SaveBadWords(map);
                        reloadBad();
                    }
                }
            };
            tpBad.Controls.AddRange(new Control[] { txtBadW, txtBadS, btnAddBad, lbBad, btnDelBad });
            reloadBad();

            tabs.TabPages.Add(tpGood);
            tabs.TabPages.Add(tpBad);
            this.Controls.Add(tabs);

            Button btnClose = new Button { Text = "Đóng", Location = new Point(514, 408), Size = new Size(96, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private void PromptAddGoodWord(string prefill)
        {
            using (Form f = new Form { Text = "Thêm từ đúng vào từ điển — VBTM Pro 3.6", Size = new Size(440, 180), StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, BackColor = Color.White, Font = new Font("Segoe UI", 9.5f) })
            {
                f.Controls.Add(new Label { Text = "Nhập từ / tên riêng / từ viết tắt đúng cần bỏ qua khi soát lỗi:", Location = new Point(18, 18), AutoSize = true });
                TextBox tb = new TextBox { Location = new Point(18, 46), Width = 386, Text = prefill };
                Button btnOk = new Button { Text = "✓ Thêm vào từ điển", Location = new Point(254, 92), Size = new Size(150, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
                btnOk.FlatAppearance.BorderSize = 0;
                btnOk.Click += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(tb.Text))
                    {
                        VbtmCustomDictStore.AddGoodWord(tb.Text);
                        MessageBox.Show("Đã thêm \"" + tb.Text.Trim() + "\" vào từ điển từ đúng!", "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        f.Close();
                    }
                };
                f.Controls.AddRange(new Control[] { tb, btnOk });
                f.ShowDialog();
            }
            this.Load += (s, e) => this.Close();
        }

        private void PromptAddBadWord(string prefill)
        {
            using (Form f = new Form { Text = "Đánh dấu từ sai cố định — VBTM Pro 3.6", Size = new Size(440, 230), StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, BackColor = Color.White, Font = new Font("Segoe UI", 9.5f) })
            {
                f.Controls.Add(new Label { Text = "Từ sai cố định cần luôn cảnh báo:", Location = new Point(18, 16), AutoSize = true });
                TextBox tbW = new TextBox { Location = new Point(18, 40), Width = 386, Text = prefill };
                f.Controls.Add(new Label { Text = "Gợi ý sửa thành từ đúng:", Location = new Point(18, 78), AutoSize = true });
                TextBox tbS = new TextBox { Location = new Point(18, 102), Width = 386 };
                Button btnOk = new Button { Text = "✕ Lưu từ sai cố định", Location = new Point(244, 146), Size = new Size(160, 34), BackColor = Color.FromArgb(220, 38, 38), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold) };
                btnOk.FlatAppearance.BorderSize = 0;
                btnOk.Click += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(tbW.Text) && !string.IsNullOrWhiteSpace(tbS.Text))
                    {
                        VbtmCustomDictStore.AddBadWord(tbW.Text, tbS.Text);
                        MessageBox.Show("Đã lưu quy tắc bắt lỗi cố định: \"" + tbW.Text.Trim() + "\" → \"" + tbS.Text.Trim() + "\"!", "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        f.Close();
                    }
                };
                f.Controls.AddRange(new Control[] { tbW, tbS, btnOk });
                f.ShowDialog();
            }
            this.Load += (s, e) => this.Close();
        }
    }
}
