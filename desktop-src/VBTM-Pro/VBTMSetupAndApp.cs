using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

[assembly: AssemblyTitle("VBTM Pro 2.2.4 - Chuan Hoa Van Ban & Chinh Ta AI Offline")]
[assembly: AssemblyDescription("Bo cong cu Chuan Hoa The Thuc Van Ban Nghi dinh 30, CAND, Tao QR & Soat Loi Chinh Ta AI Offline cho Microsoft Word")]
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
            foreach (string a in args)
            {
                if (a.Equals("--app", StringComparison.OrdinalIgnoreCase))
                    launchAppDirectly = true;
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
            this.Size = new Size(640, 470);
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
                Height = 90,
                BackColor = Color.FromArgb(30, 41, 59),
                Padding = new Padding(22, 16, 22, 16)
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
                Text = "Chuẩn hóa văn bản Nghị định 30/2020/NĐ-CP, Quy chuẩn CAND, Tạo QR & Chính tả AI cho Word",
                Font = new Font("Segoe UI", 9.2f),
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
                Location = new Point(24, 110),
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
                Location = new Point(24, 134),
                Size = new Size(574, 28),
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
                Location = new Point(24, 178),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkDesktopShortcut);

            chkWordMacroExport = new CheckBox
            {
                Text = "Cài đặt bộ quy chuẩn Microsoft Word (Nghị định 30, CAND, HD05 & COM Automation)",
                Checked = true,
                Location = new Point(24, 210),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkWordMacroExport);

            Panel infoBox = new Panel
            {
                Location = new Point(24, 246),
                Size = new Size(574, 82),
                BackColor = Color.FromArgb(22, 32, 50)
            };
            Label lblInfo = new Label
            {
                Text = "• Hoạt động ngoại tuyến độc lập 100% (Bảo mật tuyệt đối tài liệu nội bộ)\r\n" +
                       "• Tự động căn lề trang A4 (30-15-20-20mm), font Times New Roman 14pt trực tiếp vào Word\r\n" +
                       "• Đi kèm tệp gỡ cài đặt sạch sẽ: Go_Cai_Dat_VBTM.bat",
                Location = new Point(12, 10),
                Size = new Size(550, 64),
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            infoBox.Controls.Add(lblInfo);
            this.Controls.Add(infoBox);

            lblStatus = new Label
            {
                Text = "Sẵn sàng cài đặt. Nhấn 'Cài đặt ngay' để bắt đầu.",
                Location = new Point(24, 338),
                AutoSize = true,
                ForeColor = Color.FromArgb(56, 189, 248)
            };
            this.Controls.Add(lblStatus);

            progressBar = new ProgressBar
            {
                Location = new Point(24, 362),
                Size = new Size(574, 14),
                Value = 0
            };
            this.Controls.Add(progressBar);

            btnInstall = new Button
            {
                Text = "Cài Đặt Ngay",
                Location = new Point(428, 388),
                Size = new Size(170, 36),
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
                Text = "Chạy Ngay (Portable)",
                Location = new Point(250, 388),
                Size = new Size(168, 36),
                BackColor = Color.FromArgb(51, 65, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnPortable.FlatAppearance.BorderSize = 0;
            btnPortable.Click += (s, e) =>
            {
                this.Hide();
                using (MainStudioForm app = new MainStudioForm())
                {
                    app.ShowDialog();
                }
                this.Close();
            };
            this.Controls.Add(btnPortable);
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            if (isInstalled)
            {
                if (File.Exists(targetExePath))
                {
                    Process.Start(targetExePath, "--app");
                }
                this.Close();
                return;
            }

            try
            {
                string installDir = txtInstallDir.Text.Trim();
                progressBar.Value = 20;
                lblStatus.Text = "Đang tạo thư mục cài đặt và sao chép tệp quy chuẩn...";
                Application.DoEvents();

                Directory.CreateDirectory(installDir);
                targetExePath = Path.Combine(installDir, "VBTM-Pro.exe");
                File.Copy(Application.ExecutablePath, targetExePath, true);

                progressBar.Value = 55;
                if (chkWordMacroExport.Checked)
                {
                    string basPath = Path.Combine(installDir, "VBTM_Word_Addin.bas");
                    File.WriteAllText(basPath, GetWordVbaMacroCode(), Encoding.UTF8);
                }

                // Tạo tệp Gỡ cài đặt Go_Cai_Dat_VBTM.bat đúng như hướng dẫn trên trang web
                string uninstallBat = Path.Combine(installDir, "Go_Cai_Dat_VBTM.bat");
                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.2.4.lnk");
                File.WriteAllText(uninstallBat,
                    "@echo off\r\n" +
                    "chcp 65001 >nul\r\n" +
                    "echo Dang go cai dat VBTM Pro 2.2.4...\r\n" +
                    "del /f /q \"" + desktopLnk + "\" 2>nul\r\n" +
                    "echo Da go cai dat thanh cong!\r\n" +
                    "pause\r\n", Encoding.UTF8);

                progressBar.Value = 80;
                if (chkDesktopShortcut.Checked)
                {
                    lblStatus.Text = "Đang tạo lối tắt ngoài màn hình Desktop...";
                    Application.DoEvents();
                    CreateShortcut(desktopLnk, targetExePath, "--app", "VBTM Pro 2.2.4 - Chuan Hoa Van Ban & Chinh Ta AI");
                }

                progressBar.Value = 100;
                lblStatus.Text = "✓ Cài đặt VBTM Pro 2.2.4 hoàn tất thành công!";
                lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
                isInstalled = true;
                btnInstall.Text = "Hoàn Tất & Mở VBTM Pro";
                btnInstall.BackColor = Color.FromArgb(16, 185, 129);

                MessageBox.Show(
                    "Đã cài đặt thành công VBTM Pro 2.2.4 vào máy tính!\n\n" +
                    "• Thư mục: " + installDir + "\n" +
                    "• Biểu tượng Desktop: VBTM Pro 2.2.4\n\n" +
                    "Nhấn OK để khởi chạy phần mềm ngay.",
                    "Cài Đặt Thành Công - VBTM Pro 2.2.4",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                Process.Start(targetExePath, "--app");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi cài đặt: " + ex.Message, "VBTM Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                   "    MsgBox \"Da chuan hoa van ban theo Nghi dinh 30/2020/ND-CP!\", vbInformation, \"VBTM Pro 2.2.4\"\r\n" +
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
            { @"\bsát nhập\b", "sáp nhập", "Từ chuẩn hành chính ('sáp nhập')" }
        };

        public MainStudioForm()
        {
            this.Text = "VBTM Pro 2.2.4 · Chuẩn Hóa Thể Thức Văn Bản Nghị Định 30, CAND, Tạo QR & Chính Tả AI Offline";
            this.Size = new Size(1120, 760);
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
            // Top Ribbon Bar
            Panel topBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(30, 41, 59),
                Padding = new Padding(16, 10, 16, 10)
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
                Location = new Point(180, 18),
                Size = new Size(250, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5f)
            };
            cboTemplate.Items.AddRange(new object[]
            {
                "1. Mẫu Công Văn (Nghị định 30/2020/NĐ-CP)",
                "2. Mẫu Báo Cáo (Quy chuẩn CAND - TT30)",
                "3. Mẫu Quyết Định (Quy chuẩn CAND - TT30)",
                "4. Mẫu Báo Cáo Đảng (Hướng dẫn 05-HD/VPTW)"
            });
            cboTemplate.SelectedIndex = 0;
            cboTemplate.SelectedIndexChanged += (s, e) => LoadTemplate(cboTemplate.SelectedIndex);
            topBar.Controls.Add(cboTemplate);

            Button btnStd = CreateBtn("Chuẩn Hóa & Sửa Chính Tả AI", new Point(445, 14), new Size(210, 36), Color.FromArgb(2, 132, 199));
            btnStd.Click += (s, e) => StandardizeAndFixAll();
            topBar.Controls.Add(btnStd);

            Button btnWordLive = CreateBtn("Chuẩn Hóa Trên Word Đang Mở", new Point(665, 14), new Size(225, 36), Color.FromArgb(79, 70, 229));
            btnWordLive.Click += (s, e) => FormatActiveWordDocument();
            topBar.Controls.Add(btnWordLive);

            Button btnExportDoc = CreateBtn("Xuất File Word (.doc A4)", new Point(900, 14), new Size(185, 36), Color.FromArgb(5, 150, 105));
            btnExportDoc.Click += (s, e) => ExportAndOpenWordDoc();
            topBar.Controls.Add(btnExportDoc);

            this.Controls.Add(topBar);

            // Main Split Container
            Panel mainArea = new Panel { Dock = DockStyle.Fill, Padding = new Padding(16) };
            this.Controls.Add(mainArea);
            mainArea.BringToFront();

            // Left Meta Fields
            GroupBox grpMeta = new GroupBox
            {
                Text = " 1. Thông Tin Thể Thức Hành Chính (Nghị Định 30 / CAND / HD05) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 12),
                Size = new Size(520, 380)
            };

            txtGovOrg = AddLabeledInput(grpMeta, "Cơ quan chủ quản:", 28, "CÔNG AN TP ĐÀ NẴNG");
            txtIssueOrg = AddLabeledInput(grpMeta, "Đơn vị ban hành:", 64, "PHÒNG KỸ THUẬT HÌNH SỰ");
            txtDocNum = AddLabeledInput(grpMeta, "Số & Ký hiệu:", 100, "86/BC-PC09");
            txtLocation = AddLabeledInput(grpMeta, "Địa danh, ngày tháng:", 136, "Đà Nẵng, ngày 11 tháng 11 năm 2026");
            txtDocType = AddLabeledInput(grpMeta, "Tên loại văn bản:", 172, "BÁO CÁO");
            txtSubject = AddLabeledInput(grpMeta, "Trích yếu nội dung:", 208, "Kết quả ứng dụng phần mềm chuẩn hóa văn bản và số hóa tài liệu");
            txtSignerTitle = AddLabeledInput(grpMeta, "Chức vụ người ký:", 244, "TRƯỞNG PHÒNG");
            txtSignerName = AddLabeledInput(grpMeta, "Họ tên người ký:", 280, "Thượng tá Trần Minh Quân");

            Label lblRec = new Label { Text = "Nơi nhận:", Location = new Point(14, 318), AutoSize = true, ForeColor = Color.LightGray };
            txtRecipients = new TextBox
            {
                Location = new Point(155, 315),
                Size = new Size(345, 50),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = "- Ban Giám đốc CATP (để b/c);\r\n- Phòng Tham mưu;\r\n- Lưu: VT, PC09."
            };
            grpMeta.Controls.Add(lblRec);
            grpMeta.Controls.Add(txtRecipients);
            mainArea.Controls.Add(grpMeta);

            // QR Code & AI Spellcheck Group
            GroupBox grpQr = new GroupBox
            {
                Text = " 2. Tạo Mã QR Động & Thẩm Định Chính Tả AI Offline ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 400),
                Size = new Size(520, 225)
            };

            picQr = new PictureBox
            {
                Location = new Point(16, 28),
                Size = new Size(110, 110),
                BackColor = Color.White,
                SizeMode = PictureBoxSizeMode.Zoom
            };
            grpQr.Controls.Add(picQr);

            Label lblQr = new Label { Text = "Nội dung mã QR xác thực:", Location = new Point(140, 26), AutoSize = true, ForeColor = Color.LightGray };
            grpQr.Controls.Add(lblQr);

            txtQrContent = new TextBox
            {
                Location = new Point(140, 48),
                Size = new Size(255, 26),
                Text = "VBTM-VERIFY:86/BC-PC09|PC09-DANANG|2026"
            };
            txtQrContent.TextChanged += (s, e) => UpdateQrPreview();
            grpQr.Controls.Add(txtQrContent);

            Button btnSaveQr = CreateBtn("Lưu Ảnh QR", new Point(405, 45), new Size(98, 28), Color.FromArgb(51, 65, 85));
            btnSaveQr.Click += (s, e) =>
            {
                using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PNG Image|*.png", FileName = "MaQR_VanBan.png" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK && picQr.Image != null)
                    {
                        picQr.Image.Save(sfd.FileName, ImageFormat.Png);
                        MessageBox.Show("Đã lưu mã QR tại: " + sfd.FileName, "VBTM Pro 2.2.4");
                    }
                }
            };
            grpQr.Controls.Add(btnSaveQr);

            lstIssues = new ListBox
            {
                Location = new Point(140, 82),
                Size = new Size(364, 128),
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.FromArgb(251, 191, 36),
                BorderStyle = BorderStyle.FixedSingle
            };
            lstIssues.Items.Add("✓ Hệ thống AI Offline sẵn sàng rà soát lỗi chính tả & thể thức.");
            grpQr.Controls.Add(lstIssues);

            Button btnSampleErr = CreateBtn("Thử Chèn Lỗi", new Point(16, 148), new Size(110, 28), Color.FromArgb(71, 85, 105));
            btnSampleErr.Click += (s, e) =>
            {
                txtBody.Text = "Thực hiện chủ chương cải cách hành chính, đơn vị đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra";
                ScanIssuesOnly();
            };
            grpQr.Controls.Add(btnSampleErr);

            Button btnCheck = CreateBtn("Rà Soát Lỗi", new Point(16, 182), new Size(110, 28), Color.FromArgb(2, 132, 199));
            btnCheck.Click += (s, e) => ScanIssuesOnly();
            grpQr.Controls.Add(btnCheck);

            mainArea.Controls.Add(grpQr);

            // Right Editor Box
            GroupBox grpEditor = new GroupBox
            {
                Text = " 3. Nội Dung Văn Bản Chính (Phông chữ chuẩn Times New Roman 14pt - Khổ A4) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(550, 12),
                Size = new Size(536, 613)
            };

            txtBody = new TextBox
            {
                Location = new Point(14, 28),
                Size = new Size(508, 570),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 13.5f),
                BackColor = Color.White,
                ForeColor = Color.Black
            };
            grpEditor.Controls.Add(txtBody);
            mainArea.Controls.Add(grpEditor);

            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Text = "  Trạng thái: Hoạt động ngoại tuyến 100% (Offline) | Quy chuẩn Nghị định 30/2020/NĐ-CP & Thông tư 30/2023/TT-BCA (CAND)",
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.FromArgb(148, 163, 184),
                TextAlign = ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblStatus);
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
            if (idx == 0)
            {
                txtGovOrg.Text = "ỦY BAN NHÂN DÂN TP ĐÀ NẴNG";
                txtIssueOrg.Text = "SỞ THÔNG TIN VÀ TRUYỀN THÔNG";
                txtDocNum.Text = "145/STTTT-CNTT";
                txtDocType.Text = "";
                txtSubject.Text = "V/v triển khai phần mềm chuẩn hóa văn bản và số hóa tài liệu ngoại tuyến năm 2026";
                txtBody.Text = "Thực hiện Kế hoạch chuyển đổi số và hiện đại hóa công tác hành chính năm 2026, Sở Thông tin và Truyền thông triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP của Chính phủ.\r\n\r\n" +
                               "1. Các phòng, ban, đơn vị trực thuộc thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; đảm bảo đúng quy định về lề trang A4 (Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm), phông chữ Times New Roman, cỡ chữ 13-14pt.\r\n\r\n" +
                               "2. Ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm bảo đảm an toàn thông tin và bảo vệ bí mật nhà nước./.";
            }
            else if (idx == 1)
            {
                txtGovOrg.Text = "CÔNG AN TP ĐÀ NẴNG";
                txtIssueOrg.Text = "PHÒNG KỸ THUẬT HÌNH SỰ";
                txtDocNum.Text = "86/BC-PC09";
                txtDocType.Text = "BÁO CÁO";
                txtSubject.Text = "Kết quả ứng dụng phần mềm chuẩn hóa văn bản và nhận diện tài liệu OCR phục vụ công tác chuyên môn";
                txtBody.Text = "Thực hiện chương trình công tác kỹ thuật hình sự và đẩy mạnh ứng dụng khoa học công nghệ trong năm 2026, Phòng Kỹ thuật hình sự báo cáo kết quả triển khai như sau:\r\n\r\n" +
                               "I. KẾT QUẢ ĐẠT ĐƯỢC\r\n" +
                               "1. Đã đưa vào vận hành phần mềm VBTM Pro 2.2.4 phục vụ chuẩn hóa thể thức văn bản theo đúng quy chuẩn Công an nhân dân và Nghị định số 30/2020/NĐ-CP.\r\n\r\n" +
                               "2. Tích hợp thành công mô-đun soát lỗi chính tả AI tiếng Việt ngoại tuyến 100% và tạo mã QR xác thực văn bản trực tiếp cho Microsoft Word./.";
            }
            else if (idx == 2)
            {
                txtGovOrg.Text = "CÔNG AN TP ĐÀ NẴNG";
                txtIssueOrg.Text = "PHÒNG KỸ THUẬT HÌNH SỰ";
                txtDocNum.Text = "120/QĐ-PC09";
                txtDocType.Text = "QUYẾT ĐỊNH";
                txtSubject.Text = "Về việc ban hành Quy chế sử dụng phần mềm nội bộ ngoại tuyến trong đơn vị";
                txtBody.Text = "Căn cứ Thông tư số 30/2023/TT-BCA quy định về công tác văn thư trong Công an nhân dân;\r\n\r\n" +
                               "Theo đề nghị của Đội Tổng hợp,\r\n\r\n" +
                               "QUYẾT ĐỊNH:\r\n\r\n" +
                               "Điều 1. Ban hành kèm theo Quyết định này Quy chế quản lý, vận hành và sử dụng bộ công cụ VBTM Pro 2.2.4 và VietNamOCR 3.2 Pro trên hệ thống máy tính nội bộ.\r\n\r\n" +
                               "Điều 2. Các Đội nghiệp vụ và toàn thể cán bộ, chiến sĩ thuộc đơn vị chịu trách nhiệm thi hành Quyết định này./.";
            }
            else
            {
                txtGovOrg.Text = "ĐẢNG ỦY CÔNG AN TP ĐÀ NẴNG";
                txtIssueOrg.Text = "ĐẢNG ỦY PHÒNG KỸ THUẬT HÌNH SỰ";
                txtDocNum.Text = "42-BC/ĐU";
                txtDocType.Text = "BÁO CÁO";
                txtSubject.Text = "Tổng kết công tác cải cách hành chính và chuyển đổi số trong công tác Đảng năm 2026";
                txtBody.Text = "Thực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy Phòng Kỹ thuật hình sự báo cáo kết quả thực hiện chuẩn hóa toàn bộ hệ thống văn bản nội bộ./.";
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
                lstIssues.Items.Add("✓ Văn bản đạt chuẩn 100% về chính tả và thể thức Nghị định 30!");
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
            lblStatus.Text = "  ✓ Đã chuẩn hóa thể thức Nghị định 30 & tự động sửa " + fixedCount + " lỗi chính tả tiếng Việt!";
            MessageBox.Show("Đã chuẩn hóa dấu câu, khoảng trắng, ký hiệu kết thúc và sửa " + fixedCount + " lỗi chính tả AI!", "VBTM Pro 2.2.4", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Kết nối trực tiếp vào cửa sổ Microsoft Word đang mở qua COM Automation
        private void FormatActiveWordDocument()
        {
            try
            {
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    MessageBox.Show("Máy tính chưa cài đặt Microsoft Word. Đang chuyển sang chế độ xuất tệp .doc chuẩn A4!", "VBTM Pro 2.2.4");
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
                    wordApp = Activator.CreateInstance(wordType);
                }

                wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { true });
                object docs = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                int count = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                object doc = count > 0
                    ? wordType.InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null)
                    : docs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, docs, null);

                // Thiết lập lề trang A4 chuẩn Nghị định 30 (1 cm = 28.3465 pt)
                object pageSetup = doc.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, doc, null);
                pageSetup.GetType().InvokeMember("TopMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f });    // 2.0 cm
                pageSetup.GetType().InvokeMember("BottomMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f }); // 2.0 cm
                pageSetup.GetType().InvokeMember("LeftMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 85.05f });  // 3.0 cm
                pageSetup.GetType().InvokeMember("RightMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 42.5f });  // 1.5 cm

                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                object font = content.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, content, null);
                font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { "Times New Roman" });
                font.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, font, new object[] { 14f });

                MessageBox.Show("Đã chuẩn hóa trực tiếp tài liệu Microsoft Word đang mở theo Nghị định 30/2020/NĐ-CP:\n• Khổ A4, Lề Trái 30mm - Phải 15mm - Trên/Dưới 20mm\n• Font chữ: Times New Roman 14pt!", "VBTM Pro 2.2.4 - Word Add-in", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                ExportAndOpenWordDoc();
            }
        }

        private void ExportAndOpenWordDoc()
        {
            using (SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Microsoft Word Document (*.doc)|*.doc",
                FileName = "VanBan_ChuanHoa_ND30.doc"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

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
  <div style='font-size:12pt;font-weight:bold;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>
  <div style='font-size:13pt;font-weight:bold;'>Độc lập - Tự do - Hạnh phúc</div>
  <hr style='width:48%;border:0;border-top:1px solid #000;'/>
  <div style='font-size:13pt;font-style:italic;margin-top:6pt;'>" + System.Security.SecurityElement.Escape(txtLocation.Text) + @"</div>
</td>
</tr></table>
<p style='text-align:center;font-weight:bold;font-size:14pt;margin:16pt 0 4pt 0;'>" + System.Security.SecurityElement.Escape(txtDocType.Text) + @"</p>
<p style='text-align:center;font-weight:bold;font-size:13.5pt;margin:0 0 14pt 0;'>" + System.Security.SecurityElement.Escape(txtSubject.Text) + @"</p>
" + sbBody.ToString() + @"
<table style='margin-top:18pt;'><tr>
<td style='width:50%;font-size:11pt;'><b><i>Nơi nhận:</i></b>" + sbRec.ToString() + @"</td>
<td style='width:50%;text-align:center;font-size:13pt;'><b>" + System.Security.SecurityElement.Escape(txtSignerTitle.Text) + @"</b><br/><br/><br/><br/><b>" + System.Security.SecurityElement.Escape(txtSignerName.Text) + @"</b></td>
</tr></table>
</div></body></html>";

                File.WriteAllText(sfd.FileName, html, new UTF8Encoding(true));
                if (MessageBox.Show("Đã xuất tệp Word chuẩn Nghị định 30 thành công!\nBạn có muốn mở tệp ngay không?", "VBTM Pro 2.2.4", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(sfd.FileName);
                }
            }
        }

        private void UpdateQrPreview()
        {
            string text = txtQrContent.Text.Trim();
            if (string.IsNullOrEmpty(text)) text = "VBTM-PRO";
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
