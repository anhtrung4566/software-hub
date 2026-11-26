using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VBTMProSuite
{
    public class ConfusablePhraseRule
    {
        public string Wrong { get; set; }
        public string Replacement { get; set; }
        public string Description { get; set; }
        public bool IsSuspicious { get; set; }
    }

    public class OfficialTemplateItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public string FullPath { get; set; }
        public string Profile { get; set; }
        public string DocumentType { get; set; }
        public string Description { get; set; }
    }

    /// <summary>
    /// Nạp động toàn bộ kho tri thức ngôn ngữ (79.000+ từ điển, 20.600+ Bigram, 350+ cụm nhầm lẫn,
    /// Typo Map, 84 mẫu văn bản chính thức Đảng/Hành chính/CAND/Cá nhân & CTranslate2 AI Offline).
    /// Hoạt động an toàn 100% kể cả khi chạy độc lập trên máy chưa cài gói dữ liệu mở rộng.
    /// </summary>
    public static class VbtmKnowledgeHub
    {
        private static readonly object InitLock = new object();
        private static bool initialized = false;

        public static string VbtmRootDir { get; private set; }
        public static string SpellDataDir { get; private set; }
        public static string ChinhtaAiDir { get; private set; }
        public static string TemplatesDir { get; private set; }
        public static bool HasCTranslate2 { get; private set; }

        public static readonly HashSet<string> DictionaryWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public static readonly HashSet<string> ProtectedWhitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public static readonly Dictionary<string, int> BigramFrequencies = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public static readonly Dictionary<string, string> ExternalTypoMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public static readonly List<ConfusablePhraseRule> ConfusableRules = new List<ConfusablePhraseRule>();
        public static readonly List<OfficialTemplateItem> OfficialTemplates = new List<OfficialTemplateItem>();

        public static readonly Dictionary<string, string> OldToNewToneMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "hoà", "hòa" }, { "toà", "tòa" }, { "loà", "lòa" }, { "thoà", "thòa" }, { "xoà", "xòa" },
            { "hoá", "hóa" }, { "toá", "tóa" }, { "loá", "lóa" }, { "thoá", "thóa" }, { "xoá", "xóa" }, { "khoá", "khóa" },
            { "hoả", "hỏa" }, { "toả", "tỏa" }, { "loả", "lỏa" }, { "thoả", "thỏa" }, { "khoả", "khỏa" },
            { "hoã", "hõa" }, { "toã", "tõa" }, { "loã", "lõa" }, { "thoã", "thõa" },
            { "hoạ", "họa" }, { "toạ", "tọa" }, { "loạ", "lọa" }, { "thoạ", "thọa" }, { "khoạ", "khọa" },
            { "khoẻ", "khỏe" }, { "loè", "lòe" }, { "toè", "tòe" }, { "hoè", "hòe" },
            { "khoé", "khóe" }, { "loé", "lóe" }, { "toé", "tóe" }, { "hoé", "hóe" },
            { "thuỷ", "thủy" }, { "huỷ", "hủy" }, { "quỷ", "quỷ" }, { "tuỷ", "tủy" },
            { "thuý", "thúy" }, { "huý", "húy" }, { "quý", "quý" }, { "tuý", "túy" },
            { "thuỳ", "thùy" }, { "huỳ", "hùy" }, { "quỳ", "quỳ" }, { "tuỳ", "tùy" },
            { "thuỵ", "thụy" }, { "huỵ", "hụy" }, { "quỵ", "quỵ" }, { "tuỵ", "tụy" }
        };

        public static void EnsureLoaded()
        {
            if (initialized) return;
            lock (InitLock)
            {
                if (initialized) return;
                try
                {
                    DiscoverPaths();
                    LoadDictionaries();
                    LoadBigramFrequencies();
                    LoadTypoMap();
                    LoadConfusablePhrases();
                    LoadOfficialTemplates();
                }
                catch { }
                initialized = true;
            }
        }

        private static void DiscoverPaths()
        {
            string localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string[] candidateRoots = new string[]
            {
                Path.Combine(localApp, "Programs", "VBTM"),
                Path.Combine(localApp, "Programs", "VBTM-Pro"),
                Path.Combine(localApp, "VBTMPro"),
                AppDomain.CurrentDomain.BaseDirectory
            };

            foreach (string r in candidateRoots)
            {
                if (string.IsNullOrEmpty(r) || !Directory.Exists(r)) continue;
                if (string.IsNullOrEmpty(SpellDataDir) && Directory.Exists(Path.Combine(r, "SpellCheck", "Data")))
                {
                    SpellDataDir = Path.Combine(r, "SpellCheck", "Data");
                    if (string.IsNullOrEmpty(VbtmRootDir)) VbtmRootDir = r;
                }
                if (string.IsNullOrEmpty(ChinhtaAiDir) && Directory.Exists(Path.Combine(r, "ChinhtaAI")))
                {
                    ChinhtaAiDir = Path.Combine(r, "ChinhtaAI");
                    if (string.IsNullOrEmpty(VbtmRootDir)) VbtmRootDir = r;
                }
                if (string.IsNullOrEmpty(TemplatesDir) && Directory.Exists(Path.Combine(r, "templates")))
                {
                    TemplatesDir = Path.Combine(r, "templates");
                    if (string.IsNullOrEmpty(VbtmRootDir)) VbtmRootDir = r;
                }
            }

            if (!string.IsNullOrEmpty(ChinhtaAiDir))
            {
                string pyExe = Path.Combine(ChinhtaAiDir, "python_embed", "python.exe");
                string modelDir = Path.Combine(ChinhtaAiDir, "model_ct2");
                HasCTranslate2 = File.Exists(pyExe) && Directory.Exists(modelDir);
            }
        }

        private static void LoadWordListFile(string path, HashSet<string> target, bool splitSingleSyllables)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
            try
            {
                foreach (string raw in File.ReadLines(path, Encoding.UTF8))
                {
                    string w = raw.Trim();
                    if (w.Length == 0 || w.StartsWith("#")) continue;
                    string norm = w.Replace('_', ' ').Normalize(NormalizationForm.FormC);
                    target.Add(norm);
                    if (splitSingleSyllables && norm.IndexOf(' ') > 0)
                    {
                        foreach (string part in norm.Split(' '))
                        {
                            if (part.Length > 0) target.Add(part);
                        }
                    }
                }
            }
            catch { }
        }

        private static void LoadDictionaries()
        {
            if (!string.IsNullOrEmpty(SpellDataDir))
            {
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_core.dic"), DictionaryWords, true);
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_vbtm.dic"), DictionaryWords, true);
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_admin_legal.dic"), DictionaryWords, true);
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_proper_names.dic"), ProtectedWhitelist, false);
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_abbreviations.dic"), ProtectedWhitelist, false);
                LoadWordListFile(Path.Combine(SpellDataDir, "vi_foreign_terms.dic"), ProtectedWhitelist, false);
            }
            if (!string.IsNullOrEmpty(ChinhtaAiDir))
            {
                LoadWordListFile(Path.Combine(ChinhtaAiDir, "tudien.txt"), DictionaryWords, true);
            }
        }

        private static void LoadBigramFrequencies()
        {
            if (string.IsNullOrEmpty(SpellDataDir)) return;
            string path = Path.Combine(SpellDataDir, "bigram_frequency.json");
            if (!File.Exists(path)) return;
            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                JavaScriptSerializer ser = new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 };
                Dictionary<string, object> map = ser.Deserialize<Dictionary<string, object>>(json);
                if (map == null) return;
                foreach (KeyValuePair<string, object> kv in map)
                {
                    int freq = 1;
                    try { freq = Convert.ToInt32(kv.Value); } catch { }
                    BigramFrequencies[kv.Key.Normalize(NormalizationForm.FormC)] = freq;
                }
            }
            catch { }
        }

        private static void LoadTypoMap()
        {
            if (string.IsNullOrEmpty(SpellDataDir)) return;
            string path = Path.Combine(SpellDataDir, "typo_map.json");
            if (!File.Exists(path)) return;
            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                JavaScriptSerializer ser = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
                Dictionary<string, object> map = ser.Deserialize<Dictionary<string, object>>(json);
                if (map == null) return;
                foreach (KeyValuePair<string, object> kv in map)
                {
                    Dictionary<string, object> obj = kv.Value as Dictionary<string, object>;
                    if (obj == null || !obj.ContainsKey("suggestions")) continue;
                    object[] sugs = obj["suggestions"] as object[];
                    if (sugs != null && sugs.Length > 0)
                    {
                        string sug = Convert.ToString(sugs[0]);
                        if (!string.IsNullOrWhiteSpace(sug) && !string.Equals(kv.Key, sug, StringComparison.OrdinalIgnoreCase))
                        {
                            ExternalTypoMap[kv.Key.Normalize(NormalizationForm.FormC)] = sug.Normalize(NormalizationForm.FormC);
                        }
                    }
                }
            }
            catch { }
        }

        private static void LoadConfusablePhrases()
        {
            if (string.IsNullOrEmpty(SpellDataDir)) return;
            string path = Path.Combine(SpellDataDir, "confusable_phrases.json");
            if (!File.Exists(path)) return;
            try
            {
                string json = File.ReadAllText(path, Encoding.UTF8);
                JavaScriptSerializer ser = new JavaScriptSerializer { MaxJsonLength = 20 * 1024 * 1024 };
                Dictionary<string, object> map = ser.Deserialize<Dictionary<string, object>>(json);
                if (map == null) return;
                foreach (KeyValuePair<string, object> kv in map)
                {
                    Dictionary<string, object> obj = kv.Value as Dictionary<string, object>;
                    if (obj == null) continue;
                    string rep = obj.ContainsKey("replacement") ? Convert.ToString(obj["replacement"]) : "";
                    if (string.IsNullOrWhiteSpace(rep) || string.Equals(kv.Key, rep, StringComparison.OrdinalIgnoreCase)) continue;
                    string desc = obj.ContainsKey("description") ? Convert.ToString(obj["description"]) : "";
                    string sev = obj.ContainsKey("severity") ? Convert.ToString(obj["severity"]) : "error";
                    ConfusableRules.Add(new ConfusablePhraseRule
                    {
                        Wrong = kv.Key.Normalize(NormalizationForm.FormC),
                        Replacement = rep.Normalize(NormalizationForm.FormC),
                        Description = string.IsNullOrWhiteSpace(desc) ? ("Chuẩn chính tả: " + rep) : desc,
                        IsSuspicious = !string.Equals(sev, "error", StringComparison.OrdinalIgnoreCase)
                    });
                }
            }
            catch { }
        }

        private static void LoadOfficialTemplates()
        {
            if (string.IsNullOrEmpty(TemplatesDir)) return;
            string metaPath = Path.Combine(TemplatesDir, "templates_metadata.json");
            if (!File.Exists(metaPath)) return;
            try
            {
                string json = File.ReadAllText(metaPath, Encoding.UTF8);
                JavaScriptSerializer ser = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
                object[] items = ser.Deserialize<object[]>(json);
                if (items == null) return;
                foreach (object raw in items)
                {
                    Dictionary<string, object> d = raw as Dictionary<string, object>;
                    if (d == null) continue;
                    string file = d.ContainsKey("file") ? Convert.ToString(d["file"]) : "";
                    if (string.IsNullOrWhiteSpace(file)) continue;
                    string full = Path.Combine(TemplatesDir, file.Replace('/', Path.DirectorySeparatorChar));
                    if (!File.Exists(full))
                    {
                        string candPath = Path.Combine(TemplatesDir, "CAND", Path.GetFileName(file));
                        if (File.Exists(candPath)) full = candPath;
                    }
                    if (!File.Exists(full)) continue;
                    OfficialTemplates.Add(new OfficialTemplateItem
                    {
                        Id = d.ContainsKey("id") ? Convert.ToString(d["id"]) : file,
                        Name = d.ContainsKey("name") ? Convert.ToString(d["name"]) : Path.GetFileNameWithoutExtension(file),
                        FileName = file,
                        FullPath = full,
                        Profile = d.ContainsKey("profile") ? Convert.ToString(d["profile"]) : "PROFILE_HANH_CHINH",
                        DocumentType = d.ContainsKey("documentType") ? Convert.ToString(d["documentType"]) : "",
                        Description = d.ContainsKey("description") ? Convert.ToString(d["description"]) : ""
                    });
                }
            }
            catch { }
        }

        public static int GetBigramScore(string w1, string w2)
        {
            if (string.IsNullOrEmpty(w1) || string.IsNullOrEmpty(w2)) return 0;
            int freq;
            if (BigramFrequencies.TryGetValue(w1.ToLowerInvariant() + " " + w2.ToLowerInvariant(), out freq))
            {
                if (freq >= 5000) return 150;
                if (freq >= 500) return 130;
                if (freq >= 50) return 115;
                return 95;
            }
            return 0;
        }

        public static string GetKnowledgeSummary()
        {
            EnsureLoaded();
            if (DictionaryWords.Count == 0 && OfficialTemplates.Count == 0)
            {
                return "Offline tích hợp sẵn • 100% bảo mật";
            }
            return string.Format("Từ điển {0:N0} mục • {1:N0} Bigram • {2} mẫu VB{3}",
                DictionaryWords.Count, BigramFrequencies.Count, OfficialTemplates.Count,
                HasCTranslate2 ? " • AI CTranslate2" : "");
        }

        /// <summary>
        /// Gọi động cơ AI CTranslate2 Offline cục bộ (nếu máy có sẵn gói ChinhtaAI) để soát lỗi ngữ cảnh sâu.
        /// </summary>
        public static List<SpellCardIssue> RunCTranslate2DeepScan(string docText)
        {
            EnsureLoaded();
            List<SpellCardIssue> extra = new List<SpellCardIssue>();
            if (!HasCTranslate2 || string.IsNullOrWhiteSpace(docText)) return extra;

            string pyExe = Path.Combine(ChinhtaAiDir, "python_embed", "python.exe");
            string tmpIn = Path.Combine(Path.GetTempPath(), "vbtm_ct2_in_" + Guid.NewGuid().ToString("N") + ".txt");
            string tmpOut = Path.Combine(Path.GetTempPath(), "vbtm_ct2_out_" + Guid.NewGuid().ToString("N") + ".json");
            string tmpPy = Path.Combine(Path.GetTempPath(), "vbtm_ct2_runner_" + Guid.NewGuid().ToString("N") + ".py");

            try
            {
                File.WriteAllText(tmpIn, docText, new UTF8Encoding(false));
                string pyCode =
                    "import sys, os, json\n" +
                    "sys.path.insert(0, r'" + ChinhtaAiDir.Replace("'", "\\'") + "')\n" +
                    "os.chdir(r'" + ChinhtaAiDir.Replace("'", "\\'") + "')\n" +
                    "from core.speller import HybridSpeller\n" +
                    "sp = HybridSpeller()\n" +
                    "txt = open(sys.argv[1], 'r', encoding='utf-8').read()\n" +
                    "res = sp.check_text(txt) if hasattr(sp, 'check_text') else []\n" +
                    "out = []\n" +
                    "for it in (res or []):\n" +
                    "    if isinstance(it, dict):\n" +
                    "        w = str(it.get('wrong') or it.get('original') or it.get('word') or '').strip()\n" +
                    "        s = str(it.get('suggest') or it.get('suggestion') or it.get('corrected') or '').strip()\n" +
                    "        if w and s and w != s:\n" +
                    "            out.append({'wrong': w, 'suggest': s, 'type': str(it.get('type') or 'Chính tả'), 'reason': str(it.get('reason') or 'AI CTranslate2 ngữ cảnh')})\n" +
                    "open(sys.argv[2], 'w', encoding='utf-8').write(json.dumps(out, ensure_ascii=False))\n";
                File.WriteAllText(tmpPy, pyCode, new UTF8Encoding(false));

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = pyExe,
                    Arguments = "\"" + tmpPy + "\" \"" + tmpIn + "\" \"" + tmpOut + "\"",
                    WorkingDirectory = ChinhtaAiDir,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using (Process proc = Process.Start(psi))
                {
                    if (proc != null && proc.WaitForExit(35000))
                    {
                        if (File.Exists(tmpOut))
                        {
                            string json = File.ReadAllText(tmpOut, Encoding.UTF8);
                            JavaScriptSerializer ser = new JavaScriptSerializer { MaxJsonLength = 10 * 1024 * 1024 };
                            object[] arr = ser.Deserialize<object[]>(json);
                            if (arr != null)
                            {
                                foreach (object o in arr)
                                {
                                    Dictionary<string, object> d = o as Dictionary<string, object>;
                                    if (d == null) continue;
                                    string w = d.ContainsKey("wrong") ? Convert.ToString(d["wrong"]) : "";
                                    string s = d.ContainsKey("suggest") ? Convert.ToString(d["suggest"]) : "";
                                    if (string.IsNullOrWhiteSpace(w) || string.IsNullOrWhiteSpace(s) || w == s) continue;
                                    string tp = d.ContainsKey("type") ? Convert.ToString(d["type"]) : "Chính tả";
                                    string rs = d.ContainsKey("reason") ? Convert.ToString(d["reason"]) : "AI CTranslate2 ngữ cảnh";
                                    extra.Add(new SpellCardIssue
                                    {
                                        Wrong = w,
                                        Suggest = s,
                                        Badge = tp.IndexOf("Nghi", StringComparison.OrdinalIgnoreCase) >= 0 ? "Nghi ngờ" : "AI CT2",
                                        Reason = rs,
                                        IsSuspicious = tp.IndexOf("Nghi", StringComparison.OrdinalIgnoreCase) >= 0
                                    });
                                }
                            }
                        }
                    }
                    else if (proc != null && !proc.HasExited)
                    {
                        try { proc.Kill(); } catch { }
                    }
                }
            }
            catch { }
            finally
            {
                try { if (File.Exists(tmpIn)) File.Delete(tmpIn); } catch { }
                try { if (File.Exists(tmpOut)) File.Delete(tmpOut); } catch { }
                try { if (File.Exists(tmpPy)) File.Delete(tmpPy); } catch { }
            }
            return extra;
        }

        public static void OpenOfficialTemplateInWord(OfficialTemplateItem tpl)
        {
            if (tpl == null || !File.Exists(tpl.FullPath)) return;
            try
            {
                object wordApp = null;
                try { wordApp = Marshal.GetActiveObject("Word.Application"); } catch { }
                if (wordApp == null)
                {
                    Type wt = Type.GetTypeFromProgID("Word.Application");
                    if (wt != null)
                    {
                        wordApp = Activator.CreateInstance(wt);
                        wt.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { true });
                    }
                }
                if (wordApp != null)
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    // Documents.Add(TemplatePath) mở bản sao mới sạch sẽ, không làm hỏng tệp mẫu gốc!
                    docs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, docs, new object[] { tpl.FullPath });
                    try { wordApp.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, wordApp, null); } catch { }
                    return;
                }
            }
            catch { }
            try { Process.Start(tpl.FullPath); } catch { }
        }

        public static void InsertSpecialMarkIntoWord(string markText, bool isRed)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                object font = sel.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, sel, null);
                font.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, font, new object[] { 1 });
                font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { 13f });
                // wdColorRed = 255, wdColorBlack = 0
                font.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, font, new object[] { isRed ? 255 : 0 });
                sel.GetType().InvokeMember("TypeText", BindingFlags.InvokeMethod, null, sel, new object[] { "[" + markText + "] " });
                font.GetType().InvokeMember("Color", BindingFlags.SetProperty, null, font, new object[] { 0 });
                font.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, font, new object[] { 0 });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Vui lòng mở sẵn tài liệu Microsoft Word để chèn dấu '" + markText + "'.\r\n" + ex.Message,
                    "VBTM Pro 3.6", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>
    /// Hộp thoại Kho 84 Mẫu Văn Bản Chuẩn (Đảng 05-HD/VPTW · Hành chính NĐ 30 · CAND · Cá nhân)
    /// </summary>
    public class OfficialTemplateBrowserForm : Form
    {
        private ComboBox cboProfile;
        private TextBox txtSearch;
        private ListView lvTemplates;
        private Label lblCount;

        public OfficialTemplateBrowserForm()
        {
            VbtmKnowledgeHub.EnsureLoaded();
            this.Text = "VBTM Pro 3.6 — Kho 84 Mẫu Văn Bản Chuẩn (05-HD · NĐ 30 · CAND · Cá nhân)";
            this.Size = new Size(760, 560);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.3f);

            Panel top = new Panel { Dock = DockStyle.Top, Height = 58, BackColor = Color.White, Padding = new Padding(12) };
            cboProfile = new ComboBox { Location = new Point(12, 15), Size = new Size(240, 28), DropDownStyle = ComboBoxStyle.DropDownList };
            cboProfile.Items.AddRange(new object[]
            {
                "Tất cả nhóm mẫu (" + VbtmKnowledgeHub.OfficialTemplates.Count + ")",
                "Văn bản Đảng (05-HD/VPTW)",
                "Văn bản Hành chính (NĐ 30/2020)",
                "Văn bản CAND / LLVT",
                "Đơn từ & Cá nhân"
            });
            cboProfile.SelectedIndex = 0;
            cboProfile.SelectedIndexChanged += (s, e) => RefreshList();

            txtSearch = new TextBox { Location = new Point(264, 15), Size = new Size(340, 28) };
            txtSearch.TextChanged += (s, e) => RefreshList();

            lblCount = new Label { Location = new Point(614, 19), Size = new Size(124, 22), ForeColor = Color.FromArgb(37, 99, 235), Font = new Font("Segoe UI", 9f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleRight };
            top.Controls.AddRange(new Control[] { cboProfile, txtSearch, lblCount });

            lvTemplates = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                BackColor = Color.White
            };
            lvTemplates.Columns.Add("Tên mẫu văn bản", 260);
            lvTemplates.Columns.Add("Loại", 110);
            lvTemplates.Columns.Add("Nhóm quy chuẩn", 130);
            lvTemplates.Columns.Add("Mô tả chi tiết", 220);
            lvTemplates.DoubleClick += (s, e) => OpenSelected();

            Panel bot = new Panel { Dock = DockStyle.Bottom, Height = 54, BackColor = Color.White };
            Button btnOpen = new Button { Text = "📄 Mở mẫu mới trong Word", Location = new Point(450, 10), Size = new Size(190, 34), BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.3f, FontStyle.Bold), Cursor = Cursors.Hand };
            btnOpen.FlatAppearance.BorderSize = 0;
            btnOpen.Click += (s, e) => OpenSelected();
            Button btnClose = new Button { Text = "Đóng", Location = new Point(650, 10), Size = new Size(86, 34), BackColor = Color.FromArgb(241, 245, 249), ForeColor = Color.FromArgb(51, 65, 85), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnClose.Click += (s, e) => this.Close();
            bot.Controls.AddRange(new Control[] { btnOpen, btnClose });

            this.Controls.Add(lvTemplates);
            this.Controls.Add(top);
            this.Controls.Add(bot);
            RefreshList();
        }

        private string ProfileLabel(string profile)
        {
            if (profile == "PROFILE_DANG") return "Đảng (05-HD)";
            if (profile == "PROFILE_TT68_BQP") return "CAND / LLVT";
            if (profile == "PROFILE_CA_NHAN") return "Cá nhân";
            return "Hành chính (NĐ 30)";
        }

        private void RefreshList()
        {
            lvTemplates.BeginUpdate();
            lvTemplates.Items.Clear();
            string q = (txtSearch.Text ?? "").Trim().ToLowerInvariant();
            int sel = cboProfile.SelectedIndex;
            foreach (OfficialTemplateItem t in VbtmKnowledgeHub.OfficialTemplates)
            {
                if (sel == 1 && t.Profile != "PROFILE_DANG") continue;
                if (sel == 2 && t.Profile != "PROFILE_HANH_CHINH") continue;
                if (sel == 3 && t.Profile != "PROFILE_TT68_BQP") continue;
                if (sel == 4 && t.Profile != "PROFILE_CA_NHAN") continue;
                if (q.Length > 0 &&
                    t.Name.ToLowerInvariant().IndexOf(q) < 0 &&
                    t.Description.ToLowerInvariant().IndexOf(q) < 0 &&
                    t.DocumentType.ToLowerInvariant().IndexOf(q) < 0)
                {
                    continue;
                }
                ListViewItem it = new ListViewItem(t.Name);
                it.SubItems.Add(t.DocumentType);
                it.SubItems.Add(ProfileLabel(t.Profile));
                it.SubItems.Add(t.Description);
                it.Tag = t;
                lvTemplates.Items.Add(it);
            }
            lblCount.Text = lvTemplates.Items.Count + " mẫu";
            lvTemplates.EndUpdate();
        }

        private void OpenSelected()
        {
            if (lvTemplates.SelectedItems.Count == 0) return;
            OfficialTemplateItem tpl = lvTemplates.SelectedItems[0].Tag as OfficialTemplateItem;
            if (tpl != null)
            {
                VbtmKnowledgeHub.OpenOfficialTemplateInWord(tpl);
                this.Close();
            }
        }
    }
}
