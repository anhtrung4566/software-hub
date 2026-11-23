using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography;
using Microsoft.Win32;

namespace VBTM.OfficeIntegration
{
    public static class DpapiSecretStore
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("VBTM-Pro-Unified-DPAPI-v3.5");

        public static string SecretFilePath
        {
            get { return Path.Combine(PersonalInstallerEngine.InstallRoot, "secrets.dpapi"); }
        }

        public static void SaveProtectedSecret(string secretText)
        {
            Directory.CreateDirectory(PersonalInstallerEngine.InstallRoot);
            byte[] plain = Encoding.UTF8.GetBytes(secretText ?? "");
            byte[] cipher = ProtectedData.Protect(plain, Entropy, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(SecretFilePath, cipher);
        }

        public static string LoadProtectedSecret()
        {
            if (!File.Exists(SecretFilePath)) return "";
            byte[] cipher = File.ReadAllBytes(SecretFilePath);
            byte[] plain = ProtectedData.Unprotect(cipher, Entropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plain);
        }
    }

    static class Program
    {
        public const string APP_VERSION = "3.5.0";
        public const string BASE_HTTPS_URL = "https://anhtrung4566.github.io/software-hub/office-addins/";

        [STAThread]
        static int Main(string[] args)
        {
            if (args != null && args.Length > 0)
            {
                string flag = args[0].Trim().ToLowerInvariant();
                if (flag == "--check-only")
                {
                    bool checkNet = !(args.Length > 1 && args[1].ToLowerInvariant() == "--offline");
                    var rep = CompatibilityChecker.Inspect(checkNet);
                    string outPath = Path.Combine(Path.GetTempPath(), "vbtm_office_compat_report.txt");
                    File.WriteAllText(outPath, rep.ToSummaryText(), Encoding.UTF8);
                    return 0;
                }
                if (flag == "--silent-install")
                {
                    string msg;
                    bool ok = PersonalInstallerEngine.InstallOrUpdate(out msg);
                    File.WriteAllText(Path.Combine(Path.GetTempPath(), "vbtm_office_install_result.txt"), msg, Encoding.UTF8);
                    return ok ? 0 : 1;
                }
                if (flag == "--register-office")
                {
                    string msg1, msg2;
                    PersonalInstallerEngine.InstallOrUpdate(out msg1);
                    bool ok = PersonalInstallerEngine.RegisterOfficeAddinsCurrentUser(out msg2);
                    File.WriteAllText(Path.Combine(Path.GetTempPath(), "vbtm_office_register_result.txt"), msg1 + Environment.NewLine + msg2, Encoding.UTF8);
                    return ok ? 0 : 1;
                }
                if (flag == "--test-dpapi")
                {
                    string sample = "vbtm-dpapi-roundtrip-ok";
                    DpapiSecretStore.SaveProtectedSecret(sample);
                    byte[] rawCipher = File.ReadAllBytes(DpapiSecretStore.SecretFilePath);
                    string rawAscii = Encoding.UTF8.GetString(rawCipher);
                    string dec = DpapiSecretStore.LoadProtectedSecret();
                    bool ok = (dec == sample) && (rawAscii.IndexOf(sample, StringComparison.Ordinal) < 0);
                    File.WriteAllText(Path.Combine(Path.GetTempPath(), "vbtm_dpapi_test.txt"),
                        "DPAPI_OK=" + ok + " | CipherBytes=" + rawCipher.Length + " | PlaintextHidden=" + (rawAscii.IndexOf(sample, StringComparison.Ordinal) < 0),
                        Encoding.UTF8);
                    return ok ? 0 : 1;
                }
                if (flag == "--uninstall")
                {
                    string msg;
                    bool ok = PersonalInstallerEngine.CleanUninstall(out msg);
                    File.WriteAllText(Path.Combine(Path.GetTempPath(), "vbtm_office_uninstall_result.txt"), msg, Encoding.UTF8);
                    return ok ? 0 : 1;
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new OfficeSetupForm());
            return 0;
        }
    }

    public static class PersonalInstallerEngine
    {
        public static string InstallRoot
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBTMProOffice"); }
        }

        public static string ManifestsDir { get { return Path.Combine(InstallRoot, "manifests"); } }
        public static string BackupsDir { get { return Path.Combine(InstallRoot, "backups"); } }
        public static string ConfigFile { get { return Path.Combine(InstallRoot, "install_state.txt"); } }

        private static readonly string[] ManifestMap = new string[] {
            "VBTM.WordAddin|vbtm-word-manifest.xml",
            "VBTM.ExcelAddin|vbtm-excel-manifest.xml",
            "VBTM.PowerPointAddin|vbtm-powerpoint-manifest.xml",
            "VBTM.OutlookAddin|vbtm-outlook-manifest.xml"
        };

        public static bool InstallOrUpdate(out string report)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                Directory.CreateDirectory(InstallRoot);
                Directory.CreateDirectory(ManifestsDir);
                Directory.CreateDirectory(BackupsDir);

                // 1. Sao lưu phiên bản manifest cũ (nếu có) trước khi cập nhật, giữ nguyên API Key người dùng
                if (File.Exists(ConfigFile))
                {
                    string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string bkSub = Path.Combine(BackupsDir, "backup_" + stamp);
                    Directory.CreateDirectory(bkSub);
                    foreach (string f in Directory.GetFiles(ManifestsDir, "*.xml"))
                        File.Copy(f, Path.Combine(bkSub, Path.GetFileName(f)), true);
                    File.Copy(ConfigFile, Path.Combine(bkSub, "install_state.txt"), true);
                    sb.AppendLine("[OK] Đã sao lưu cấu hình cũ vào: " + bkSub + " (API Key cá nhân được giữ nguyên).");
                }

                // 2. Nạp 4 XML Manifest (ưu tiên bản cục bộ đi kèm hoặc tải qua HTTPS chính thức)
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                int okCount = 0;
                foreach (string entry in ManifestMap)
                {
                    string[] parts = entry.Split('|');
                    string proj = parts[0];
                    string outName = parts[1];
                    string targetPath = Path.Combine(ManifestsDir, outName);

                    string localCandidate1 = Path.Combine(exeDir, "..", "office-addins", proj, "manifest.xml");
                    string localCandidate2 = Path.Combine(exeDir, "office-addins", proj, "manifest.xml");
                    string content = null;

                    if (File.Exists(localCandidate1))
                        content = File.ReadAllText(localCandidate1, Encoding.UTF8);
                    else if (File.Exists(localCandidate2))
                        content = File.ReadAllText(localCandidate2, Encoding.UTF8);
                    else
                        content = TryDownloadManifest(Program.BASE_HTTPS_URL + proj + "/manifest.xml");

                    if (string.IsNullOrEmpty(content))
                        content = BuildEmbeddedFallbackManifest(proj);

                    File.WriteAllText(targetPath, content, Encoding.UTF8);
                    okCount++;
                    sb.AppendLine("[OK] Đã chuẩn bị Manifest: " + outName);
                }

                // 3. Ghi trạng thái cài đặt & tạo hướng dẫn cá nhân (KHÔNG tự ý sửa Registry hay Trust Center Office)
                File.WriteAllText(ConfigFile,
                    "Version=" + Program.APP_VERSION + Environment.NewLine +
                    "InstalledAt=" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine +
                    "Scope=CurrentUser" + Environment.NewLine +
                    "RegistryModified=False" + Environment.NewLine,
                    Encoding.UTF8);

                WritePersonalInstallGuide();
                sb.AppendLine("[OK] Hoàn tất cài đặt phạm vi người dùng hiện tại tại: " + InstallRoot);
                sb.AppendLine("[LƯU Ý BẢO MẬT] Bộ cài KHÔNG tự ý sửa Registry Office, KHÔNG dùng network share sideload và KHÔNG làm mất tài liệu/API Key.");
                report = sb.ToString();
                return okCount == 4;
            }
            catch (Exception ex)
            {
                sb.AppendLine("[LỖI] " + ex.Message);
                report = sb.ToString();
                return false;
            }
        }

        public static bool RollbackLastBackup(out string report)
        {
            try
            {
                if (!Directory.Exists(BackupsDir))
                {
                    report = "Chưa có bản sao lưu nào để khôi phục.";
                    return false;
                }
                string[] dirs = Directory.GetDirectories(BackupsDir, "backup_*");
                if (dirs.Length == 0)
                {
                    report = "Không tìm thấy thư mục sao lưu trước đó.";
                    return false;
                }
                Array.Sort(dirs);
                string latest = dirs[dirs.Length - 1];
                Directory.CreateDirectory(ManifestsDir);
                foreach (string f in Directory.GetFiles(latest, "*.xml"))
                    File.Copy(f, Path.Combine(ManifestsDir, Path.GetFileName(f)), true);
                report = "Đã khôi phục thành công cấu hình & manifest từ: " + latest;
                return true;
            }
            catch (Exception ex)
            {
                report = "Lỗi khôi phục: " + ex.Message;
                return false;
            }
        }

        private static readonly string[] RegisteredGuidMap = new string[] {
            "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3001|vbtm-word-manifest.xml|Microsoft Word",
            "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3002|vbtm-excel-manifest.xml|Microsoft Excel",
            "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3003|vbtm-powerpoint-manifest.xml|Microsoft PowerPoint"
        };

        public static bool RegisterOfficeAddinsCurrentUser(out string report)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                if (!Directory.Exists(ManifestsDir))
                {
                    string tmp;
                    InstallOrUpdate(out tmp);
                }
                using (RegistryKey devKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\16.0\WEF\Developer"))
                {
                    if (devKey == null) throw new Exception("Không mở được khóa HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer");
                    foreach (string row in RegisteredGuidMap)
                    {
                        string[] p = row.Split('|');
                        string guid = p[0];
                        string filePath = Path.Combine(ManifestsDir, p[1]);
                        string appName = p[2];
                        if (!File.Exists(filePath)) throw new FileNotFoundException("Thiếu tệp manifest: " + filePath);
                        devKey.SetValue(guid, filePath, RegistryValueKind.String);
                        sb.AppendLine("[OK] Đã đăng ký " + appName + " -> " + filePath);
                    }
                }
                sb.AppendLine("[LƯU Ý] Khởi động lại Word / Excel / PowerPoint -> vào tab Insert (Chèn) -> My Add-ins -> chọn VBTM PRO AI (sau lần đầu nút sẽ nằm ngay trên tab Home).");
                sb.AppendLine("[OUTLOOK] Mở Outlook -> Get Add-ins (hoặc https://aka.ms/olksideload) -> My Add-ins -> Add from file -> chọn vbtm-outlook-manifest.xml.");
                report = sb.ToString();
                return true;
            }
            catch (Exception ex)
            {
                sb.AppendLine("[LỖI ĐĂNG KÝ] " + ex.Message);
                report = sb.ToString();
                return false;
            }
        }

        public static bool CleanUninstall(out string report)
        {
            try
            {
                try
                {
                    using (RegistryKey devKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\WEF\Developer", true))
                    {
                        if (devKey != null)
                        {
                            foreach (string row in RegisteredGuidMap)
                            {
                                string guid = row.Split('|')[0];
                                if (devKey.GetValue(guid) != null) devKey.DeleteValue(guid, false);
                            }
                        }
                    }
                }
                catch { }
                if (Directory.Exists(InstallRoot))
                {
                    Directory.Delete(InstallRoot, true);
                    report = "Đã gỡ cài đặt sạch thư mục VBTMProOffice và hủy đăng ký Add-in khỏi HKCU Office (không ảnh hưởng tài liệu và không xóa API Key của VBTM Desktop).";
                }
                else
                {
                    report = "Không phát hiện dữ liệu cài đặt VBTMProOffice để gỡ.";
                }
                return true;
            }
            catch (Exception ex)
            {
                report = "Lỗi gỡ cài đặt: " + ex.Message;
                return false;
            }
        }

        private static string TryDownloadManifest(string url)
        {
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
                using (WebClient wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    return wc.DownloadString(url);
                }
            }
            catch { return null; }
        }

        private static string BuildEmbeddedFallbackManifest(string proj)
        {
            string host = "Document";
            string id = "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3001";
            string title = "VBTM PRO AI FOR WORD";
            if (proj == "VBTM.ExcelAddin") { host = "Workbook"; id = "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3002"; title = "VBTM PRO AI FOR EXCEL"; }
            else if (proj == "VBTM.PowerPointAddin") { host = "Presentation"; id = "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3003"; title = "VBTM PRO AI FOR POWERPOINT"; }
            else if (proj == "VBTM.OutlookAddin") { host = "Mailbox"; id = "9c1f4a20-3b7e-4e82-9a11-7d5c2b8f3004"; title = "VBTM PRO AI FOR OUTLOOK"; }
            string url = Program.BASE_HTTPS_URL + proj + "/taskpane.html";
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
                   "<OfficeApp xmlns=\"http://schemas.microsoft.com/office/appforoffice/1.1\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:type=\"TaskPaneApp\">\n" +
                   "  <Id>" + id + "</Id>\n" +
                   "  <Version>3.0.0.0</Version>\n" +
                   "  <ProviderName>VBTM Pro</ProviderName>\n" +
                   "  <DefaultLocale>vi-VN</DefaultLocale>\n" +
                   "  <DisplayName DefaultValue=\"" + title + "\"/>\n" +
                   "  <Description DefaultValue=\"VBTM Pro Office AI Integration\"/>\n" +
                   "  <Hosts><Host Name=\"" + host + "\"/></Hosts>\n" +
                   "  <DefaultSettings><SourceLocation DefaultValue=\"" + url + "\"/></DefaultSettings>\n" +
                   "  <Permissions>ReadWriteDocument</Permissions>\n" +
                   "</OfficeApp>";
        }

        private static void WritePersonalInstallGuide()
        {
            string guide =
                "HƯỚNG DẪN NẠP ADD-IN VBTM PRO AI CÁ NHÂN VÀO MICROSOFT OFFICE\r\n" +
                "==============================================================\r\n" +
                "Thư mục chứa 4 file XML Manifest: " + ManifestsDir + "\r\n\r\n" +
                "1. ĐỐI VỚI MICROSOFT 365 / OFFICE ON THE WEB (Khuyên dùng - không sửa Registry):\r\n" +
                "   - Mở Word / Excel / PowerPoint -> Tab Home (hoặc Insert) -> Add-ins -> More Add-ins -> Upload My Add-in.\r\n" +
                "   - Chọn file vbtm-word-manifest.xml (hoặc excel / powerpoint tương ứng).\r\n\r\n" +
                "2. ĐỐI VỚI CLASSIC OUTLOOK & NEW OUTLOOK:\r\n" +
                "   - Mở Outlook -> All Apps -> Get Add-ins (hoặc truy cập https://aka.ms/olksideload) -> My add-ins -> Add a custom add-in -> Add from file...\r\n" +
                "   - Chọn file vbtm-outlook-manifest.xml.\r\n\r\n" +
                "3. ĐỐI VỚI OFFICE 2016 / 2019 / 2021 / 2024 DESKTOP CÀI TRÊN MÁY CÁ NHÂN:\r\n" +
                "   - Mở Word/Excel/PowerPoint -> File -> Options -> Trust Center -> Trust Center Settings -> Trusted Add-in Catalogs.\r\n" +
                "   - Tự thêm thư mục danh mục tin cậy do bạn chủ động xác nhận -> tích Show in Menu -> khởi động lại Office -> Insert -> My Add-ins -> chọn VBTM PRO AI.\r\n" +
                "   (Bộ cài VBTM Pro tôn trọng bảo mật Windows, KHÔNG tự ý thay đổi Registry Trust Center khi chưa có sự đồng ý của bạn.)\r\n";
            File.WriteAllText(Path.Combine(InstallRoot, "HUONG_DAN_NAP_ADDIN.txt"), guide, Encoding.UTF8);
        }
    }

    public class OfficeSetupForm : Form
    {
        private TextBox txtLog;
        private Button btnCheck;
        private Button btnInstall;
        private Button btnOpenDir;
        private Button btnRollback;
        private Button btnUninstall;
        private Label lblStatus;

        public OfficeSetupForm()
        {
            this.Text = "VBTM Pro Office Setup v" + Program.APP_VERSION + " — Windows Personal Edition";
            this.Size = new Size(760, 560);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(248, 250, 252);
            this.Font = new Font("Segoe UI", 9.5f);

            Panel header = new Panel { Dock = DockStyle.Top, Height = 64, BackColor = Color.FromArgb(127, 29, 29) };
            Label title = new Label
            {
                Text = "VBTM PRO OFFICE INTEGRATION — WINDOWS PERSONAL EDITION",
                ForeColor = Color.FromArgb(254, 240, 138),
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                Location = new Point(16, 10),
                AutoSize = true
            };
            Label sub = new Label
            {
                Text = "Tích hợp trực tiếp Word · Excel · PowerPoint · Outlook (Office 2016 / 2019 / 2021 / 2024 / M365) · Không cần quyền Admin",
                ForeColor = Color.FromArgb(254, 226, 226),
                Font = new Font("Segoe UI", 8.5f),
                Location = new Point(16, 36),
                AutoSize = true
            };
            header.Controls.Add(title);
            header.Controls.Add(sub);
            this.Controls.Add(header);

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Location = new Point(14, 74),
                Size = new Size(718, 42),
                FlowDirection = FlowDirection.LeftToRight
            };

            btnCheck = MakeBtn("1. Kiểm tra Office", Color.FromArgb(29, 78, 216), Color.White);
            btnInstall = MakeBtn("2. Cài đặt / Cập nhật", Color.FromArgb(153, 27, 27), Color.White);
            Button btnRegister = MakeBtn("3. Đăng ký hiện trên Office", Color.FromArgb(217, 119, 6), Color.White);
            btnOpenDir = MakeBtn("4. Mở thư mục Manifest", Color.FromArgb(5, 150, 105), Color.White);
            btnRollback = MakeBtn("Khôi phục", Color.FromArgb(241, 245, 249), Color.FromArgb(15, 23, 42));
            btnUninstall = MakeBtn("Gỡ sạch", Color.FromArgb(254, 226, 226), Color.FromArgb(153, 27, 27));

            btnCheck.Click += (s, e) => RunCheck();
            btnInstall.Click += (s, e) => RunInstall();
            btnRegister.Click += (s, e) =>
            {
                if (MessageBox.Show("Bạn có đồng ý đăng ký 3 Add-in VBTM PRO AI (Word, Excel, PowerPoint) vào phạm vi tài khoản Windows hiện tại (HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer)?\nKhông cần quyền Admin và sẽ tự động gỡ sạch khi bấm 'Gỡ sạch'.",
                    "Xác nhận đăng ký Office Add-in", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                string rmsg;
                bool rok = PersonalInstallerEngine.RegisterOfficeAddinsCurrentUser(out rmsg);
                txtLog.AppendText(Environment.NewLine + Environment.NewLine + "=== ĐĂNG KÝ OFFICE ADD-IN ===" + Environment.NewLine + rmsg);
                lblStatus.Text = rok ? "Đã đăng ký xong! Hãy khởi động lại Word / Excel / PowerPoint." : "Đăng ký gặp lỗi — xem chi tiết.";
            };
            btnOpenDir.Click += (s, e) => OpenManifestFolder();
            btnRollback.Click += (s, e) => RunRollback();
            btnUninstall.Click += (s, e) => RunUninstall();

            toolbar.Controls.AddRange(new Control[] { btnCheck, btnInstall, btnRegister, btnOpenDir, btnRollback, btnUninstall });
            this.Controls.Add(toolbar);

            txtLog = new TextBox
            {
                Location = new Point(14, 122),
                Size = new Size(716, 360),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9.5f),
                BackColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtLog);

            lblStatus = new Label
            {
                Location = new Point(14, 490),
                Size = new Size(716, 24),
                Text = "Sẵn sàng kiểm tra môi trường Microsoft Office...",
                ForeColor = Color.FromArgb(71, 85, 105),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStatus);

            this.Shown += (s, e) => RunCheck();
        }

        private Button MakeBtn(string text, Color bg, Color fg)
        {
            Button b = new Button
            {
                Text = text,
                AutoSize = true,
                Height = 32,
                Padding = new Padding(8, 0, 8, 0),
                BackColor = bg,
                ForeColor = fg,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, 6, 0)
            };
            b.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            return b;
        }

        private void RunCheck()
        {
            lblStatus.Text = "Đang kiểm tra Windows, Office 32/64-bit, Build, WebView2 và kết nối HTTPS...";
            Application.DoEvents();
            var rep = CompatibilityChecker.Inspect(true);
            txtLog.Text = rep.ToSummaryText();
            lblStatus.Text = "Hoàn tất kiểm tra: " + rep.OverallTier;
        }

        private void RunInstall()
        {
            string msg;
            bool ok = PersonalInstallerEngine.InstallOrUpdate(out msg);
            txtLog.AppendText(Environment.NewLine + Environment.NewLine + "=== KẾT QUẢ CÀI ĐẶT ===" + Environment.NewLine + msg);
            lblStatus.Text = ok ? "Cài đặt thành công (phạm vi người dùng hiện tại)." : "Cài đặt gặp cảnh báo — xem chi tiết bên trên.";
        }

        private void OpenManifestFolder()
        {
            if (!Directory.Exists(PersonalInstallerEngine.ManifestsDir))
            {
                MessageBox.Show("Vui lòng bấm '2. Cài đặt / Cập nhật' trước để tạo thư mục Manifest.", "VBTM Pro Office", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Process.Start("explorer.exe", PersonalInstallerEngine.ManifestsDir);
        }

        private void RunRollback()
        {
            string msg;
            PersonalInstallerEngine.RollbackLastBackup(out msg);
            txtLog.AppendText(Environment.NewLine + "[KHÔI PHỤC] " + msg);
            lblStatus.Text = msg;
        }

        private void RunUninstall()
        {
            if (MessageBox.Show("Gỡ cài đặt sạch VBTM Pro Office khỏi %LOCALAPPDATA%\\VBTMProOffice?\n(Không làm mất tài liệu Office và không xóa API Key của bạn.)",
                "Xác nhận gỡ cài đặt", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            string msg;
            PersonalInstallerEngine.CleanUninstall(out msg);
            txtLog.AppendText(Environment.NewLine + "[GỠ CÀI ĐẶT] " + msg);
            lblStatus.Text = msg;
        }
    }
}
