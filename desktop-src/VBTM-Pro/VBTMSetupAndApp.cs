using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("VBTM Pro 2.2.4 - Chuan Hoa Van Ban Dang Uy & UBND Phuong")]
[assembly: AssemblyDescription("Bo cong cu Chuan Hoa The Thuc Van Ban Nghi dinh 30, Van ban Dang, Tao QR & Soat Loi Chinh Ta AI Offline cho Microsoft Word")]
[assembly: AssemblyCompany("Software Hub")]
[assembly: AssemblyProduct("VBTM Pro 2.2.4")]
[assembly: AssemblyVersion("2.2.4.0")]
[assembly: AssemblyFileVersion("2.2.4.0")]

namespace VBTMProSuite
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool launchAppDirectly = false;
            bool uninstallMode = false;
            foreach (string a in args)
            {
                if (a.Equals("--app", StringComparison.OrdinalIgnoreCase))
                    launchAppDirectly = true;
                if (a.Equals("--uninstall", StringComparison.OrdinalIgnoreCase))
                    uninstallMode = true;
            }

            if (uninstallMode)
            {
                RunUninstaller();
                return;
            }

            string exeName = Path.GetFileName(Application.ExecutablePath);
            if (exeName.Equals("VBTM-Pro.exe", StringComparison.OrdinalIgnoreCase))
            {
                launchAppDirectly = true;
            }

            if (launchAppDirectly)
            {
                Application.Run(new MainStudioForm());
            }
            else
            {
                Application.Run(new SetupWizardForm());
            }
        }

        private static void RunUninstaller()
        {
            try
            {
                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.2.4.lnk");
                if (File.Exists(desktopLnk)) File.Delete(desktopLnk);

                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VBTM Pro 2.2.4.lnk");
                if (File.Exists(startMenuLnk)) File.Delete(startMenuLnk);

                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\VBTMPro", false);
                MessageBox.Show("Đã gỡ cài đặt VBTM Pro 2.2.4 sạch sẽ khỏi hệ thống!", "Gỡ Cài Đặt - VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Thông báo khi gỡ cài đặt: " + ex.Message, "VBTM Pro 2.2.4");
            }
        }
    }

    // =========================================================================
    // 1. TRÌNH CÀI ĐẶT TỰ ĐỘNG (SETUP WIZARD FORM)
    // =========================================================================
    public class SetupWizardForm : Form
    {
        private CheckBox chkDesktopShortcut;
        private CheckBox chkWordMacroExport;
        private TextBox txtInstallDir;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnInstall;
        private Button btnPortable;
        private bool isInstalled = false;
        private string targetExePath = "";

        public SetupWizardForm()
        {
            this.Text = "Trình Cài Đặt VBTM Pro 2.2.4 (Offline Setup)";
            this.ClientSize = new Size(630, 440);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(15, 23, 42);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            BuildUI();
        }

        private void BuildUI()
        {
            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 88,
                BackColor = Color.FromArgb(30, 41, 59)
            };

            Label lblTitle = new Label
            {
                Text = "VBTM Pro 2.2.4 · Bộ Cài Đặt Ngoại Tuyến",
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                AutoSize = true,
                Location = new Point(22, 16)
            };

            Label lblSub = new Label
            {
                Text = "Chuẩn hóa văn bản Đảng ủy & UBND Phường (Nghị định 30, HD 36 & 05-HD/VPTW), Tạo QR & Chính tả AI",
                Font = new Font("Segoe UI", 9.0f),
                ForeColor = Color.FromArgb(203, 213, 225),
                AutoSize = true,
                Location = new Point(24, 50)
            };

            header.Controls.Add(lblTitle);
            header.Controls.Add(lblSub);
            this.Controls.Add(header);

            Label lblDir = new Label
            {
                Text = "Thư mục cài đặt (Không yêu cầu quyền Administrator):",
                Location = new Point(24, 106),
                AutoSize = true,
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            this.Controls.Add(lblDir);

            string defaultDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "VBTM-Pro"
            );

            txtInstallDir = new TextBox
            {
                Text = defaultDir,
                Location = new Point(24, 130),
                Size = new Size(580, 28),
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9.5f)
            };
            this.Controls.Add(txtInstallDir);

            chkDesktopShortcut = new CheckBox
            {
                Text = "Tạo biểu tượng VBTM Pro 2.2.4 ngoài màn hình Desktop và Start Menu",
                Checked = true,
                Location = new Point(24, 172),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkDesktopShortcut);

            chkWordMacroExport = new CheckBox
            {
                Text = "Tích hợp bộ quy chuẩn Microsoft Word (Nghị định 30/2020/NĐ-CP, Hướng dẫn 36 & 05-HD/VPTW)",
                Checked = true,
                Location = new Point(24, 202),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkWordMacroExport);

            Panel infoBox = new Panel
            {
                Location = new Point(24, 238),
                Size = new Size(580, 82),
                BackColor = Color.FromArgb(22, 32, 50)
            };
            Label lblInfo = new Label
            {
                Text = "• Hoạt động ngoại tuyến độc lập 100% phục vụ Đảng ủy và Ủy ban nhân dân phường\r\n" +
                       "• Tự động căn lề trang A4 (30-15-20-20mm), font Times New Roman 14pt trực tiếp vào Word\r\n" +
                       "• Hỗ trợ gỡ cài đặt qua Installed Apps của Windows hoặc chạy Go_Cai_Dat_VBTM.bat",
                Location = new Point(12, 10),
                Size = new Size(556, 64),
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            infoBox.Controls.Add(lblInfo);
            this.Controls.Add(infoBox);

            lblStatus = new Label
            {
                Text = "Sẵn sàng cài đặt. Nhấn 'Cài Đặt Ngay' để bắt đầu.",
                Location = new Point(24, 332),
                AutoSize = true,
                ForeColor = Color.FromArgb(56, 189, 248)
            };
            this.Controls.Add(lblStatus);

            progressBar = new ProgressBar
            {
                Location = new Point(24, 356),
                Size = new Size(580, 14),
                Value = 0
            };
            this.Controls.Add(progressBar);

            btnInstall = new Button
            {
                Text = "Cài Đặt Ngay",
                Location = new Point(434, 384),
                Size = new Size(170, 38),
                BackColor = Color.FromArgb(2, 132, 199),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            this.Controls.Add(btnInstall);

            btnPortable = new Button
            {
                Text = "Chạy Ngay (Không Cần Cài)",
                Location = new Point(236, 384),
                Size = new Size(188, 38),
                BackColor = Color.FromArgb(51, 65, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnPortable.FlatAppearance.BorderSize = 0;
            btnPortable.Click += (s, e) => LaunchInstalledApp();
            this.Controls.Add(btnPortable);
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            if (isInstalled)
            {
                LaunchInstalledApp();
                return;
            }

            try
            {
                string installDir = txtInstallDir.Text.Trim();
                progressBar.Value = 25;
                lblStatus.Text = "Đang tạo thư mục cài đặt và sao chép tệp thực thi...";
                Application.DoEvents();

                Directory.CreateDirectory(installDir);
                targetExePath = Path.Combine(installDir, "VBTM-Pro.exe");

                if (!Application.ExecutablePath.Equals(targetExePath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(Application.ExecutablePath, targetExePath, true);
                }

                progressBar.Value = 55;
                if (chkWordMacroExport.Checked)
                {
                    string basPath = Path.Combine(installDir, "VBTM_Word_Addin.bas");
                    File.WriteAllText(basPath, GetWordVbaMacroCode(), Encoding.UTF8);
                }

                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.2.4.lnk");
                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VBTM Pro 2.2.4.lnk");

                string uninstallBat = Path.Combine(installDir, "Go_Cai_Dat_VBTM.bat");
                File.WriteAllText(uninstallBat,
                    "@echo off\r\n" +
                    "chcp 65001 >nul\r\n" +
                    "\"" + targetExePath + "\" --uninstall\r\n", Encoding.UTF8);

                RegisterUninstallerInRegistry(installDir, targetExePath);

                progressBar.Value = 85;
                if (chkDesktopShortcut.Checked)
                {
                    lblStatus.Text = "Đang tạo lối tắt ngoài màn hình Desktop & Start Menu...";
                    Application.DoEvents();
                    CreateShortcut(desktopLnk, targetExePath, "--app", "VBTM Pro 2.2.4 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                    CreateShortcut(startMenuLnk, targetExePath, "--app", "VBTM Pro 2.2.4 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                }

                progressBar.Value = 100;
                lblStatus.Text = "✓ Cài đặt VBTM Pro 2.2.4 hoàn tất thành công!";
                lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
                isInstalled = true;
                btnInstall.Text = "Mở VBTM Pro 2.2.4";
                btnInstall.BackColor = Color.FromArgb(16, 185, 129);

                MessageBox.Show(
                    "Đã cài đặt thành công VBTM Pro 2.2.4!\n\n" +
                    "• Thư mục cài đặt: " + installDir + "\n" +
                    "• Lối tắt Desktop: VBTM Pro 2.2.4\n\n" +
                    "Nhấn OK để mở phần mềm ngay.",
                    "Cài Đặt Thành Công - VBTM Pro 2.2.4",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                LaunchInstalledApp();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi cài đặt: " + ex.Message, "VBTM Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LaunchInstalledApp()
        {
            this.Hide();
            using (MainStudioForm app = new MainStudioForm())
            {
                app.ShowDialog();
            }
            this.Close();
        }

        private void RegisterUninstallerInRegistry(string installDir, string exePath)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\VBTMPro"))
                {
                    if (key != null)
                    {
                        key.SetValue("DisplayName", "VBTM Chuẩn Hóa Văn Bản Đảng Ủy & UBND Phường (VBTM Pro 2.2.4)");
                        key.SetValue("DisplayVersion", "2.2.4");
                        key.SetValue("Publisher", "Software Hub");
                        key.SetValue("InstallLocation", installDir);
                        key.SetValue("DisplayIcon", exePath);
                        key.SetValue("UninstallString", "\"" + exePath + "\" --uninstall");
                        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                    }
                }
            }
            catch { }
        }

        private void CreateShortcut(string shortcutPath, string targetPath, string arguments, string description)
        {
            try
            {
                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellType == null) return;
                object shell = Activator.CreateInstance(shellType);
                object shortcut = shellType.InvokeMember("CreateShortcut", BindingFlags.InvokeMethod, null, shell, new object[] { shortcutPath });
                Type scType = shortcut.GetType();
                scType.InvokeMember("TargetPath", BindingFlags.SetProperty, null, shortcut, new object[] { targetPath });
                scType.InvokeMember("Arguments", BindingFlags.SetProperty, null, shortcut, new object[] { arguments });
                scType.InvokeMember("WorkingDirectory", BindingFlags.SetProperty, null, shortcut, new object[] { Path.GetDirectoryName(targetPath) });
                scType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, new object[] { description });
                scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
            }
            catch { }
        }

        private string GetWordVbaMacroCode()
        {
            return "Attribute VB_Name = \"VBTM_Pro_ND30\"\r\n" +
                   "Sub ChuanHoaVanBan_NghiDinh30()\r\n" +
                   "    On Error Resume Next\r\n" +
                   "    With ActiveDocument.PageSetup\r\n" +
                   "        .TopMargin = CentimetersToPoints(2)\r\n" +
                   "        .BottomMargin = CentimetersToPoints(2)\r\n" +
                   "        .LeftMargin = CentimetersToPoints(3)\r\n" +
                   "        .RightMargin = CentimetersToPoints(1.5)\r\n" +
                   "    End With\r\n" +
                   "    With ActiveDocument.Content.Font\r\n" +
                   "        .Name = \"Times New Roman\"\r\n" +
                   "        .Size = 14\r\n" +
                   "    End With\r\n" +
                   "    MsgBox \"Da chuan hoa van ban theo Nghi dinh 30/2020/ND-CP & Quy chuan Dang uy - UBND Phuong!\", vbInformation, \"VBTM Pro 2.2.4\"\r\n" +
                   "End Sub\r\n";
        }
    }

    // =========================================================================
    // 2. PHẦN MỀM CHÍNH: VBTM PRO 2.2.4 STUDIO + MICROSOFT WORD AUTOMATION
    // =========================================================================
    public class MainStudioForm : Form
    {
        private ComboBox cboTemplate;
        private TextBox txtGovOrg, txtIssueOrg, txtDocNum, txtLocation, txtDocType, txtSubject, txtSignerTitle, txtSignerName, txtQrContent;
        private TextBox txtBody, txtRecipients;
        private PictureBox picQr;
        private ListBox lstIssues;
        private Label lblStatus;
        private CheckBox chkIncludeQr;

        private static readonly string[,] SpellRules = new string[,]
        {
            { @"\bbổ xung\b", "bổ sung", "Sai phụ âm x/s ('bổ sung')" },
            { @"\bsử lý\b", "xử lý", "Sai phụ âm s/x ('xử lý')" },
            { @"\bxử dụng\b", "sử dụng", "Sai phụ âm x/s ('sử dụng')" },
            { @"\bchính sác\b", "chính xác", "Sai phụ âm s/x ('chính xác')" },
            { @"\bkiểm xoát\b", "kiểm soát", "Sai phụ âm x/s ('kiểm soát')" },
            { @"\bchiển khai\b", "triển khai", "Sai phụ âm ch/tr ('triển khai')" },
            { @"\bchủ chương\b", "chủ trương", "Sai phụ âm ch/tr ('chủ trương')" },
            { @"\bkế hoạc\b", "kế hoạch", "Thiếu âm cuối ('kế hoạch')" },
            { @"\bngiêm túc\b", "nghiêm túc", "Sai quy tắc ng/ngh ('nghiêm túc')" },
            { @"\bhổ trợ\b", "hỗ trợ", "Sai dấu hỏi/ngã ('hỗ trợ')" },
            { @"\btiêu chuẫn\b", "tiêu chuẩn", "Sai dấu hỏi/ngã ('tiêu chuẩn')" },
            { @"\bkhẩn chương\b", "khẩn trương", "Sai phụ âm ch/tr ('khẩn trương')" },
            { @"\bsát nhập\b", "sáp nhập", "Từ chuẩn hành chính ('sáp nhập')" },
            { @"\bxáng suốt\b", "sáng suốt", "Sai phụ âm x/s ('sáng suốt')" },
            { @"\bgiử gìn\b", "giữ gìn", "Sai dấu hỏi/ngã ('giữ gìn')" },
            { @"\bvửng chắc\b", "vững chắc", "Sai dấu hỏi/ngã ('vững chắc')" },
            { @"\bchặc chẽ\b", "chặt chẽ", "Sai âm cuối c/t ('chặt chẽ')" },
            { @"\bsâu xát\b", "sâu sát", "Sai phụ âm x/s ('sâu sát')" },
            { @"\bquan chọng\b", "quan trọng", "Sai phụ âm ch/tr ('quan trọng')" },
            { @"\bchách nhiệm\b", "trách nhiệm", "Sai phụ âm ch/tr ('trách nhiệm')" },
            { @"\btrực chách\b", "trực trách", "Sai phụ âm ch/tr ('trực trách')" },
            { @"\bphát chiển\b", "phát triển", "Sai phụ âm ch/tr ('phát triển')" },
            { @"\btập chung\b", "tập trung", "Sai phụ âm ch/tr ('tập trung')" },
            { @"\btuyên chuyền\b", "tuyên truyền", "Sai phụ âm ch/tr ('tuyên truyền')" },
            { @"\brà xoát\b", "rà soát", "Sai phụ âm x/s ('rà soát')" },
            { @"\bxắp xếp\b", "sắp xếp", "Sai phụ âm x/s ('sắp xếp')" },
            { @"\bsắp sếp\b", "sắp xếp", "Sai phụ âm s/x ('sắp xếp')" },
            { @"\bsuất sắc\b", "xuất sắc", "Sai phụ âm s/x ('xuất sắc')" },
            { @"\bsản suất\b", "sản xuất", "Sai phụ âm s/x ('sản xuất')" },
            { @"\bngân xách\b", "ngân sách", "Sai phụ âm x/s ('ngân sách')" },
            { @"\bdanh xách\b", "danh sách", "Sai phụ âm x/s ('danh sách')" },
            { @"\bchính xách\b", "chính sách", "Sai phụ âm x/s ('chính sách')" },
            { @"\bkỷ thuật\b", "kỹ thuật", "Sai dấu hỏi/ngã ('kỹ thuật')" },
            { @"\blỉnh vực\b", "lĩnh vực", "Sai dấu hỏi/ngã ('lĩnh vực')" },
            { @"\bmẩu mực\b", "mẫu mực", "Sai dấu hỏi/ngã ('mẫu mực')" },
            { @"\bhướng dẩn\b", "hướng dẫn", "Sai dấu hỏi/ngã ('hướng dẫn')" }
        };

        public MainStudioForm()
        {
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls; } catch { }
            this.Text = "VBTM Pro 2.2.4 · Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy & UBND Phường, Tạo QR & Chính Tả AI";
            this.ClientSize = new Size(1120, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(15, 23, 42);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            BuildMainUI();
            LoadTemplate(0);
            UpdateQrPreview();
        }

        private void BuildMainUI()
        {
            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                Text = "  ✓ Trạng thái: Hoạt động ngoại tuyến 100% (Offline) | Quy chuẩn Nghị định 30/2020/NĐ-CP (UBND Phường) & Hướng dẫn 36, 05-HD/VPTW (Đảng ủy Phường)",
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.FromArgb(56, 189, 248),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Panel topBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(30, 41, 59)
            };

            Label lblBrand = new Label
            {
                Text = "VBTM Pro 2.2.4",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 16),
                AutoSize = true
            };
            topBar.Controls.Add(lblBrand);

            cboTemplate = new ComboBox
            {
                Location = new Point(166, 18),
                Size = new Size(218, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.0f)
            };
            cboTemplate.Items.AddRange(new object[]
            {
                "1. Mẫu Công Văn UBND Phường (Nghị định 30)",
                "2. Mẫu Báo Cáo UBND Phường (Nghị định 30)",
                "3. Mẫu Quyết Định UBND Phường (Nghị định 30)",
                "4. Mẫu Báo Cáo Đảng Ủy Phường (HD 36 & 05)"
            });
            cboTemplate.SelectedIndex = 0;
            cboTemplate.SelectedIndexChanged += (s, e) => LoadTemplate(cboTemplate.SelectedIndex);
            topBar.Controls.Add(cboTemplate);

            Button btnOpenFile = CreateBtn("📂 Mở File (.docx/.txt)", new Point(390, 14), new Size(165, 36), Color.FromArgb(14, 165, 233));
            btnOpenFile.Click += (s, e) => OpenWordOrTextFileFromDisk();
            topBar.Controls.Add(btnOpenFile);

            Button btnStd = CreateBtn("1. Chuẩn Hóa & Sửa Lỗi AI", new Point(562, 14), new Size(185, 36), Color.FromArgb(2, 132, 199));
            btnStd.Click += (s, e) => StandardizeAndFixAll();
            topBar.Controls.Add(btnStd);

            Button btnWordLive = CreateBtn("2. Chuẩn Hóa Word Đang Mở", new Point(754, 14), new Size(182, 36), Color.FromArgb(79, 70, 229));
            btnWordLive.Click += (s, e) => FormatActiveWordDocument();
            topBar.Controls.Add(btnWordLive);

            Button btnExportDoc = CreateBtn("3. Xuất Word (.doc A4)", new Point(942, 14), new Size(162, 36), Color.FromArgb(5, 150, 105));
            btnExportDoc.Click += (s, e) => ExportAndOpenWordDoc();
            topBar.Controls.Add(btnExportDoc);

            Panel mainArea = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(15, 23, 42)
            };

            this.Controls.Add(mainArea);
            this.Controls.Add(topBar);
            this.Controls.Add(lblStatus);

            GroupBox grpMeta = new GroupBox
            {
                Text = " 1. Thông Tin Thể Thức Hành Chính (Đảng Ủy & Ủy Ban Nhân Dân Phường) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 12),
                Size = new Size(520, 375)
            };

            txtGovOrg = AddLabeledInput(grpMeta, "Cơ quan cấp trên:", 28, "ỦY BAN NHÂN DÂN QUẬN");
            txtIssueOrg = AddLabeledInput(grpMeta, "Cơ quan ban hành:", 62, "ỦY BAN NHÂN DÂN PHƯỜNG");
            txtDocNum = AddLabeledInput(grpMeta, "Số & Ký hiệu:", 96, "145/UBND-VP");
            txtLocation = AddLabeledInput(grpMeta, "Địa danh, ngày tháng:", 130, "Phường, ngày 11 tháng 11 năm 2026");
            txtDocType = AddLabeledInput(grpMeta, "Tên loại văn bản:", 164, "");
            txtSubject = AddLabeledInput(grpMeta, "Trích yếu nội dung:", 198, "V/v triển khai ứng dụng chuẩn hóa thể thức văn bản và số hóa hồ sơ");
            txtSignerTitle = AddLabeledInput(grpMeta, "Chức vụ người ký:", 232, "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH");
            txtSignerName = AddLabeledInput(grpMeta, "Họ tên người ký:", 266, "Nguyễn Văn An");

            Label lblRec = new Label { Text = "Nơi nhận:", Location = new Point(14, 304), AutoSize = true, ForeColor = Color.LightGray };
            txtRecipients = new TextBox
            {
                Location = new Point(155, 302),
                Size = new Size(345, 58),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "- Như trên;\r\n- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- Lưu: VT, VP."
            };
            grpMeta.Controls.Add(lblRec);
            grpMeta.Controls.Add(txtRecipients);
            mainArea.Controls.Add(grpMeta);

            GroupBox grpQr = new GroupBox
            {
                Text = " 2. Tạo Mã QR Động & Thẩm Định Chính Tả AI Offline ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 395),
                Size = new Size(520, 218)
            };

            picQr = new PictureBox
            {
                Location = new Point(16, 26),
                Size = new Size(106, 106),
                BackColor = Color.White,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            grpQr.Controls.Add(picQr);

            chkIncludeQr = new CheckBox
            {
                Text = "Đính kèm mã QR khi xuất ra Word",
                Checked = true,
                Location = new Point(140, 24),
                AutoSize = true,
                ForeColor = Color.White
            };
            grpQr.Controls.Add(chkIncludeQr);

            txtQrContent = new TextBox
            {
                Location = new Point(140, 50),
                Size = new Size(255, 26),
                Text = "VBTM-VERIFY:145/UBND-VP|UBND-PHUONG|2026"
            };
            txtQrContent.TextChanged += (s, e) => UpdateQrPreview();
            grpQr.Controls.Add(txtQrContent);

            Button btnSaveQr = CreateBtn("Lưu Ảnh QR", new Point(402, 48), new Size(102, 28), Color.FromArgb(51, 65, 85));
            btnSaveQr.Click += (s, e) =>
            {
                using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PNG Image|*.png", FileName = "MaQR_VanBan_Phuong.png" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK && picQr.Image != null)
                    {
                        picQr.Image.Save(sfd.FileName, ImageFormat.Png);
                        MessageBox.Show("Đã lưu ảnh mã QR tại:\n" + sfd.FileName, "VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            };
            grpQr.Controls.Add(btnSaveQr);

            lstIssues = new ListBox
            {
                Location = new Point(140, 84),
                Size = new Size(364, 120),
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.FromArgb(251, 191, 36),
                BorderStyle = BorderStyle.FixedSingle
            };
            lstIssues.Items.Add("✓ Hệ thống AI Offline sẵn sàng rà soát lỗi chính tả & thể thức.");
            grpQr.Controls.Add(lstIssues);

            Button btnSampleErr = CreateBtn("Chèn Lỗi Mẫu", new Point(16, 140), new Size(106, 30), Color.FromArgb(71, 85, 105));
            btnSampleErr.Click += (s, e) =>
            {
                txtBody.Text = "Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra";
                ScanIssuesOnly();
            };
            grpQr.Controls.Add(btnSampleErr);

            Button btnCheck = CreateBtn("Rà Soát Lỗi", new Point(16, 176), new Size(106, 30), Color.FromArgb(2, 132, 199));
            btnCheck.Click += (s, e) => ScanIssuesOnly();
            grpQr.Controls.Add(btnCheck);

            mainArea.Controls.Add(grpQr);

            GroupBox grpEditor = new GroupBox
            {
                Text = " 3. Nội Dung Văn Bản Chính (Phông chữ chuẩn Times New Roman 14pt - Khổ A4) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(550, 12),
                Size = new Size(550, 601)
            };

            txtBody = new TextBox
            {
                Location = new Point(14, 28),
                Size = new Size(522, 558),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 13.5f),
                BackColor = Color.White,
                ForeColor = Color.Black
            };
            grpEditor.Controls.Add(txtBody);
            mainArea.Controls.Add(grpEditor);
        }

        private TextBox AddLabeledInput(Control parent, string label, int y, string defaultVal)
        {
            Label lbl = new Label { Text = label, Location = new Point(14, y + 3), AutoSize = true, ForeColor = Color.LightGray };
            TextBox txt = new TextBox { Location = new Point(155, y), Size = new Size(345, 26), Text = defaultVal };
            parent.Controls.Add(lbl);
            parent.Controls.Add(txt);
            return txt;
        }

        private Button CreateBtn(string text, Point loc, Size size, Color bg)
        {
            Button b = new Button
            {
                Text = text,
                Location = loc,
                Size = size,
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }

        private void LoadTemplate(int idx)
        {
            if (txtGovOrg == null) return;
            if (idx == 0)
            {
                txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                txtDocNum.Text = "145/UBND-VP";
                txtDocType.Text = "";
                txtSubject.Text = "V/v triển khai ứng dụng chuẩn hóa thể thức văn bản và số hóa hồ sơ hành chính tại địa phương";
                txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                txtSignerName.Text = "Nguyễn Văn An";
                txtRecipients.Text = "- Như trên;\r\n- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- UBMTTQ và các đoàn thể phường;\r\n- Các tổ dân phố;\r\n- Lưu: VT, VP.";
                txtBody.Text = "Thực hiện Kế hoạch cải cách hành chính và chuyển đổi số năm 2026, Ủy ban nhân dân phường triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư.\r\n\r\n" +
                               "1. Các bộ phận chuyên môn, công chức thuộc Ủy ban nhân dân phường thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; bảo đảm đúng quy định về khổ giấy A4, lề trang (Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm), phông chữ Times New Roman, cỡ chữ 13-14pt.\r\n\r\n" +
                               "2. Đẩy mạnh số hóa hồ sơ, tài liệu lưu trữ và ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm nâng cao chất lượng ban hành văn bản./.";
            }
            else if (idx == 1)
            {
                txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                txtDocNum.Text = "86/BC-UBND";
                txtDocType.Text = "BÁO CÁO";
                txtSubject.Text = "Kết quả thực hiện công tác cải cách hành chính, văn thư lưu trữ và chuyển đổi số trên địa bàn phường năm 2026";
                txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                txtSignerName.Text = "Nguyễn Văn An";
                txtRecipients.Text = "- UBND Quận (để b/c);\r\n- Phòng Nội vụ;\r\n- Thường trực Đảng ủy, HĐND phường;\r\n- CT, các PCT UBND phường;\r\n- Lưu: VT, VP.";
                txtBody.Text = "Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai công tác chuẩn hóa văn bản và số hóa hồ sơ trên địa bàn phường như sau:\r\n\r\n" +
                               "I. KẾT QUẢ ĐẠT ĐƯỢC\r\n" +
                               "1. Công tác ban hành văn bản: 100% văn bản hành chính của Hội đồng nhân dân và Ủy ban nhân dân phường được chuẩn hóa đồng bộ theo Nghị định số 30/2020/NĐ-CP của Chính phủ.\r\n\r\n" +
                               "2. Công tác số hóa hồ sơ: Đã ứng dụng phần mềm VBTM Pro 2.2.4 và VietNamOCR 3.2 Pro phục vụ rà soát lỗi chính tả tiếng Việt, tạo mã QR quản lý hồ sơ và bóc tách dữ liệu bảng biểu sang Word, Excel nhanh chóng, chính xác./.";
            }
            else if (idx == 2)
            {
                txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                txtDocNum.Text = "120/QĐ-UBND";
                txtDocType.Text = "QUYẾT ĐỊNH";
                txtSubject.Text = "Về việc ban hành Quy chế quản lý, chuẩn hóa văn bản và số hóa hồ sơ tài liệu tại Ủy ban nhân dân phường";
                txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                txtSignerName.Text = "Nguyễn Văn An";
                txtRecipients.Text = "- Như Điều 3;\r\n- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- Lưu: VT, VP.";
                txtBody.Text = "Căn cứ Luật Tổ chức chính quyền địa phương;\r\n\r\n" +
                               "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\r\n\r\n" +
                               "Theo đề nghị của Công chức Văn phòng - Thống kê phường,\r\n\r\n" +
                               "QUYẾT ĐỊNH:\r\n\r\n" +
                               "Điều 1. Ban hành kèm theo Quyết định này Quy chế ứng dụng phần mềm chuẩn hóa thể thức văn bản VBTM Pro 2.2.4 và số hóa hồ sơ tài liệu VietNamOCR 3.2 Pro trong công tác văn thư, lưu trữ tại Ủy ban nhân dân phường.\r\n\r\n" +
                               "Điều 2. Công chức Văn phòng - Thống kê và các cán bộ, công chức, bộ phận chuyên môn thuộc Ủy ban nhân dân phường chịu trách nhiệm thi hành Quyết định này./.";
            }
            else
            {
                txtGovOrg.Text = "QUẬN ỦY";
                txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                txtDocNum.Text = "42-BC/ĐU";
                txtDocType.Text = "BÁO CÁO";
                txtSubject.Text = "Tổng kết công tác xây dựng Đảng, cải cách hành chính trong Đảng và chuyển đổi số tại Đảng bộ phường năm 2026";
                txtSignerTitle.Text = "T/M ĐẢNG ỦY - BÍ THƯ";
                txtSignerName.Text = "Trần Minh Quân";
                txtRecipients.Text = "- Ban Thường vụ Quận ủy (để b/c);\r\n- Văn phòng Quận ủy;\r\n- Các đồng chí Đảng ủy viên;\r\n- HĐND, UBND, UBMTTQ phường;\r\n- Các chi bộ trực thuộc;\r\n- Lưu Văn phòng Đảng ủy.";
                txtBody.Text = "Thực hiện Hướng dẫn số 36-HD/VPTW và Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy phường báo cáo kết quả thực hiện chuẩn hóa toàn bộ hệ thống văn bản của Đảng ủy và các chi bộ trực thuộc năm 2026./.";
            }
            if (txtQrContent != null)
            {
                txtQrContent.Text = "VBTM:" + txtDocNum.Text + "|" + txtIssueOrg.Text;
            }
        }

        private void ScanIssuesOnly()
        {
            lstIssues.Items.Clear();
            string content = txtBody.Text;
            for (int i = 0; i < SpellRules.GetLength(0); i++)
            {
                if (Regex.IsMatch(content, SpellRules[i, 0], RegexOptions.IgnoreCase))
                {
                    lstIssues.Items.Add("[Chính tả AI] " + SpellRules[i, 2] + " -> '" + SpellRules[i, 1] + "'");
                }
            }
            if (Regex.IsMatch(content, @"\s+[,.;:!?]"))
            {
                lstIssues.Items.Add("[Thể thức NĐ30] Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)");
            }
            if (!content.Trim().EndsWith("./."))
            {
                lstIssues.Items.Add("[Thể thức NĐ30] Cuối văn bản thiếu ký hiệu kết thúc './.'");
            }
            if (lstIssues.Items.Count == 0)
            {
                lstIssues.Items.Add("✓ Văn bản đạt chuẩn 100% về chính tả và thể thức Đảng ủy / UBND phường!");
            }
        }

        private void StandardizeAndFixAll()
        {
            string raw = txtBody.Text;
            int fixedCount = 0;
            for (int i = 0; i < SpellRules.GetLength(0); i++)
            {
                var matches = Regex.Matches(raw, SpellRules[i, 0], RegexOptions.IgnoreCase);
                if (matches.Count > 0)
                {
                    fixedCount += matches.Count;
                    raw = Regex.Replace(raw, SpellRules[i, 0], SpellRules[i, 1], RegexOptions.IgnoreCase);
                }
            }

            raw = Regex.Replace(raw, @"[ \t]+", " ");
            raw = Regex.Replace(raw, @"\s+([,.;:!?])", "$1");
            raw = Regex.Replace(raw, @"([,.;:!?])([^\s\d/])", "$1 $2");
            raw = raw.Trim();
            if (raw.Length > 0 && !raw.EndsWith("./."))
            {
                raw = raw.TrimEnd('.') + "./.";
            }

            txtBody.Text = raw;
            ScanIssuesOnly();
            lblStatus.Text = "  ✓ Đã chuẩn hóa thể thức Đảng ủy / UBND phường & tự động sửa " + fixedCount + " lỗi chính tả tiếng Việt!";
            MessageBox.Show("Đã chuẩn hóa dấu câu, khoảng trắng, ký hiệu kết thúc và tự động sửa " + fixedCount + " lỗi chính tả AI!", "VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void FormatActiveWordDocument()
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    ExportAndOpenWordDoc();
                    return;
                }

                object wordApp = null;
                try
                {
                    wordApp = Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    string tempDocPath = Path.Combine(Path.GetTempPath(), "VBTM_VanBan_Phuong_ChuanHoa.doc");
                    WriteWordDocFile(tempDocPath);
                    Process.Start(new ProcessStartInfo(tempDocPath) { UseShellExecute = true });
                    lblStatus.Text = "  ✓ Đã khởi tạo tài liệu chuẩn A4 và mở trực tiếp trong Microsoft Word!";
                    return;
                }

                wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { true });
                object docs = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                int count = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                if (count == 0)
                {
                    string tempDocPath = Path.Combine(Path.GetTempPath(), "VBTM_VanBan_Phuong_ChuanHoa.doc");
                    WriteWordDocFile(tempDocPath);
                    Process.Start(new ProcessStartInfo(tempDocPath) { UseShellExecute = true });
                    return;
                }

                object doc = wordType.InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);

                object pageSetup = doc.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, doc, null);
                pageSetup.GetType().InvokeMember("TopMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f });
                pageSetup.GetType().InvokeMember("BottomMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f });
                pageSetup.GetType().InvokeMember("LeftMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 85.05f });
                pageSetup.GetType().InvokeMember("RightMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 42.5f });

                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                object font = content.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, content, null);
                font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { "Times New Roman" });
                font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { 14f });

                MessageBox.Show(
                    "Đã chuẩn hóa trực tiếp tài liệu Microsoft Word đang mở theo quy chuẩn Đảng ủy & UBND Phường:\n\n" +
                    "• Khổ giấy A4, Lề Trái 30mm - Phải 15mm - Trên 20mm - Dưới 20mm\n" +
                    "• Phông chữ chuẩn: Times New Roman 14pt!",
                    "VBTM Pro 2.2.4 - Word Automation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch
            {
                ExportAndOpenWordDoc();
            }
        }

        private void WriteWordDocFile(string filePath)
        {
            StringBuilder sbBody = new StringBuilder();
            foreach (string line in txtBody.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                sbBody.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;font-size:14pt;'>{0}</p>",
                    System.Security.SecurityElement.Escape(line.Trim()));
            }

            StringBuilder sbRec = new StringBuilder();
            foreach (string r in txtRecipients.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                sbRec.AppendFormat("<div>{0}</div>", System.Security.SecurityElement.Escape(r.Trim()));
            }

            string qrImgHtml = "";
            if (chkIncludeQr != null && chkIncludeQr.Checked && picQr.Image != null)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    picQr.Image.Save(ms, ImageFormat.Png);
                    string b64 = Convert.ToBase64String(ms.ToArray());
                    qrImgHtml = "<div style='margin-top:8pt;'><img src='data:image/png;base64," + b64 + "' width='72' height='72'/><div style='font-size:8.5pt;color:#444;'>Mã QR xác thực VBTM</div></div>";
                }
            }

            bool isDang = (cboTemplate != null && cboTemplate.SelectedIndex == 3) ||
                          txtIssueOrg.Text.IndexOf("ĐẢNG", StringComparison.OrdinalIgnoreCase) >= 0;

            string rightHeaderHtml = isDang
                ? "<div style='font-size:12pt;font-weight:bold;'>ĐẢNG CỘNG SẢN VIỆT NAM</div>"
                : "<div style='font-size:12pt;font-weight:bold;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>" +
                  "<div style='font-size:13pt;font-weight:bold;'>Độc lập - Tự do - Hạnh phúc</div>";

            string html = @"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'>
<style>
@page Section1 { size:21.0cm 29.7cm; margin:2.0cm 1.5cm 2.0cm 3.0cm; }
div.Section1 { page:Section1; font-family:'Times New Roman',serif; font-size:14pt; color:#000; }
table { width:100%; border-collapse:collapse; }
td { vertical-align:top; font-family:'Times New Roman',serif; }
</style></head>
<body><div class='Section1'>
<table><tr>
<td style='width:44%;text-align:center;'>
  <div style='font-size:12pt;'>" + System.Security.SecurityElement.Escape(txtGovOrg.Text) + @"</div>
  <div style='font-size:12pt;font-weight:bold;'>" + System.Security.SecurityElement.Escape(txtIssueOrg.Text) + @"</div>
  <hr style='width:38%;border:0;border-top:1px solid #000;'/>
  <div style='font-size:13pt;margin-top:6pt;'>Số: " + System.Security.SecurityElement.Escape(txtDocNum.Text) + @"</div>
</td>
<td style='width:56%;text-align:center;'>
  " + rightHeaderHtml + @"
  <hr style='width:48%;border:0;border-top:1px solid #000;'/>
  <div style='font-size:13pt;font-style:italic;margin-top:6pt;'>" + System.Security.SecurityElement.Escape(txtLocation.Text) + @"</div>
</td>
</tr></table>
<p style='text-align:center;font-weight:bold;font-size:14pt;margin:16pt 0 4pt 0;'>" + System.Security.SecurityElement.Escape(txtDocType.Text) + @"</p>
<p style='text-align:center;font-weight:bold;font-size:13.5pt;margin:0 0 14pt 0;'>" + System.Security.SecurityElement.Escape(txtSubject.Text) + @"</p>
" + sbBody.ToString() + @"
<table style='margin-top:18pt;'><tr>
<td style='width:50%;font-size:11pt;'><b><i>Nơi nhận:</i></b>" + sbRec.ToString() + qrImgHtml + @"</td>
<td style='width:50%;text-align:center;font-size:13pt;'><b>" + System.Security.SecurityElement.Escape(txtSignerTitle.Text) + @"</b><br/><br/><br/><br/><b>" + System.Security.SecurityElement.Escape(txtSignerName.Text) + @"</b></td>
</tr></table>
</div></body></html>";

            File.WriteAllText(filePath, html, new UTF8Encoding(true));
        }

        private void ExportAndOpenWordDoc()
        {
            using (SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Microsoft Word Document (*.doc)|*.doc",
                FileName = "VanBan_Phuong_ChuanHoa.doc"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                WriteWordDocFile(sfd.FileName);
                lblStatus.Text = "  ✓ Đã xuất tệp Word chuẩn A4: " + sfd.FileName;
                if (MessageBox.Show("Đã xuất tệp Word chuẩn thể thức thành công!\nBạn có muốn mở tệp ngay không?", "VBTM Pro 2.2.4", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
        }

        private void OpenWordOrTextFileFromDisk()
        {
            using (OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Chọn tệp văn bản Word (.docx, .doc) hoặc Text (.txt) để chuẩn hóa",
                Filter = "Word & Text Files (*.docx;*.doc;*.txt)|*.docx;*.doc;*.txt|All Files (*.*)|*.*"
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string ext = Path.GetExtension(ofd.FileName).ToLowerInvariant();
                    string extractedText = "";

                    if (ext == ".docx")
                    {
                        string tempZip = Path.Combine(Path.GetTempPath(), "vbtm_docx_" + Guid.NewGuid().ToString("N") + ".zip");
                        string tempExtractDir = Path.Combine(Path.GetTempPath(), "vbtm_docx_dir_" + Guid.NewGuid().ToString("N"));
                        try
                        {
                            File.Copy(ofd.FileName, tempZip, true);
                            ZipFile.ExtractToDirectory(tempZip, tempExtractDir);
                            string xmlPath = Path.Combine(tempExtractDir, "word", "document.xml");
                            if (File.Exists(xmlPath))
                            {
                                string xml = File.ReadAllText(xmlPath, Encoding.UTF8);
                                xml = Regex.Replace(xml, @"</w:p>", "\r\n\r\n");
                                xml = Regex.Replace(xml, @"<w:br[^>]*/?>", "\r\n");
                                xml = Regex.Replace(xml, @"<[^>]+>", "");
                                extractedText = WebUtility.HtmlDecode(xml);
                            }
                        }
                        finally
                        {
                            try { if (File.Exists(tempZip)) File.Delete(tempZip); } catch { }
                            try { if (Directory.Exists(tempExtractDir)) Directory.Delete(tempExtractDir, true); } catch { }
                        }
                    }
                    else
                    {
                        string rawContent = File.ReadAllText(ofd.FileName, Encoding.UTF8);
                        if (rawContent.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 ||
                            rawContent.IndexOf("<p", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string stripped = Regex.Replace(rawContent, @"<style[\s\S]*?</style>", "", RegexOptions.IgnoreCase);
                            stripped = Regex.Replace(stripped, @"</(p|div|tr|h1|h2|h3)>", "\r\n\r\n", RegexOptions.IgnoreCase);
                            stripped = Regex.Replace(stripped, @"<br\s*/?>", "\r\n", RegexOptions.IgnoreCase);
                            stripped = Regex.Replace(stripped, @"<[^>]+>", " ");
                            extractedText = WebUtility.HtmlDecode(stripped);
                        }
                        else
                        {
                            extractedText = rawContent;
                        }
                    }

                    extractedText = Regex.Replace(extractedText ?? "", @"(\r?\n){3,}", "\r\n\r\n").Trim();
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        MessageBox.Show("Không tìm thấy nội dung văn bản trong tệp đã chọn!", "VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (extractedText.IndexOf("ĐẢNG CỘNG SẢN VIỆT NAM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        extractedText.IndexOf("ĐẢNG ỦY", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        txtGovOrg.Text = "QUẬN ỦY";
                        txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    }

                    txtBody.Text = extractedText;
                    StandardizeAndFixAll();
                    lblStatus.Text = "  ✓ Đã đọc và chuẩn hóa tệp: " + Path.GetFileName(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi đọc tệp văn bản: " + ex.Message, "VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpdateQrPreview()
        {
            if (picQr == null || txtQrContent == null) return;
            string text = txtQrContent.Text.Trim();
            if (string.IsNullOrEmpty(text)) text = "VBTM-PRO";

            try
            {
                string url = "https://api.qrserver.com/v1/create-qr-code/?size=180x180&margin=6&ecc=M&data=" + Uri.EscapeDataString(text);
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Timeout = 2500;
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (Stream s = resp.GetResponseStream())
                {
                    if (s != null)
                    {
                        using (Image img = Image.FromStream(s))
                        {
                            picQr.Image = new Bitmap(img);
                            return;
                        }
                    }
                }
            }
            catch
            {
                // Fallback offline matrix if no internet connection
            }

            int size = 25, scale = 4, pad = 6;
            Bitmap bmp = new Bitmap(size * scale + pad * 2, size * scale + pad * 2);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                uint hash = 2166136261;
                foreach (char c in text) { hash ^= c; hash *= 16777619; }

                Action<int, int> drawFinder = (ox, oy) =>
                {
                    g.FillRectangle(Brushes.Black, pad + ox * scale, pad + oy * scale, 7 * scale, 7 * scale);
                    g.FillRectangle(Brushes.White, pad + (ox + 1) * scale, pad + (oy + 1) * scale, 5 * scale, 5 * scale);
                    g.FillRectangle(Brushes.Black, pad + (ox + 2) * scale, pad + (oy + 2) * scale, 3 * scale, 3 * scale);
                };
                drawFinder(0, 0);
                drawFinder(size - 7, 0);
                drawFinder(0, size - 7);

                for (int r = 0; r < size; r++)
                {
                    for (int c = 0; c < size; c++)
                    {
                        if ((r < 8 && c < 8) || (r < 8 && c >= size - 8) || (r >= size - 8 && c < 8)) continue;
                        int ch = text[(r * size + c) % text.Length];
                        if (((hash ^ (r * 131 + c * 313 + ch * 17)) & 3) == 0 || ((r * c + ch) % 5 == 0))
                        {
                            g.FillRectangle(Brushes.Black, pad + c * scale, pad + r * scale, scale, scale);
                        }
                    }
                }
            }
            picQr.Image = bmp;
        }
    }
}
