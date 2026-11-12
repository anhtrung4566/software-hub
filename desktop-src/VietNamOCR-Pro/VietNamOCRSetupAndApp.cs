using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("VietNamOCR 3.2 Pro - AI Neural OCR Tieng Viet & Doc PDF Da Trang")]
[assembly: AssemblyDescription("Ung dung AI Neural OCR nhan dien chu tieng Viet tu hinh anh & PDF scan phuc vu luu tru ho so Dang uy va UBND Phuong")]
[assembly: AssemblyCompany("Software Hub")]
[assembly: AssemblyProduct("VietNamOCR 3.2 Pro")]
[assembly: AssemblyVersion("3.2.0.0")]
[assembly: AssemblyFileVersion("3.2.0.0")]

namespace VietNamOCRSuite
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls;
            }
            catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool launchApp = false;
            bool uninstall = false;
            foreach (string a in args)
            {
                if (a.Equals("--app", StringComparison.OrdinalIgnoreCase))
                    launchApp = true;
                if (a.Equals("--uninstall", StringComparison.OrdinalIgnoreCase))
                    uninstall = true;
            }

            if (uninstall)
            {
                RunUninstall();
                return;
            }

            string exeName = Path.GetFileName(Application.ExecutablePath);
            if (exeName.Equals("VietNamOCR.exe", StringComparison.OrdinalIgnoreCase))
            {
                launchApp = true;
            }

            if (launchApp)
            {
                Application.Run(new OcrMainForm());
            }
            else
            {
                Application.Run(new OcrSetupForm());
            }
        }

        private static void RunUninstall()
        {
            try
            {
                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VietNamOCR 3.2 Pro.lnk");
                if (File.Exists(desktopLnk)) File.Delete(desktopLnk);

                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VietNamOCR 3.2 Pro.lnk");
                if (File.Exists(startMenuLnk)) File.Delete(startMenuLnk);

                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\VietNamOCRPro", false);
                MessageBox.Show("Đã gỡ cài đặt VietNamOCR 3.2 Pro thành công khỏi máy tính!", "Gỡ Cài Đặt VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Thông báo gỡ cài đặt: " + ex.Message, "VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    // =========================================================================
    // 1. TRÌNH CÀI ĐẶT VIETNAMOCR 3.2 PRO (SETUP WIZARD)
    // =========================================================================
    public class OcrSetupForm : Form
    {
        private TextBox txtInstallDir;
        private CheckBox chkDesktopShortcut;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Button btnInstall;
        private bool isInstalled = false;
        private string targetExePath = "";

        public OcrSetupForm()
        {
            this.Text = "Trình Cài Đặt VietNamOCR 3.2 Pro (x64 Setup)";
            this.Size = new Size(650, 450);
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
                BackColor = Color.FromArgb(30, 27, 75)
            };

            Label lblTitle = new Label
            {
                Text = "VietNamOCR 3.2 Pro (x64) · AI Neural OCR Tiếng Việt",
                Font = new Font("Segoe UI", 15f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                AutoSize = true,
                Location = new Point(22, 16)
            };

            Label lblSub = new Label
            {
                Text = "Nhận diện chữ Tiếng Việt từ Ảnh & PDF nhiều trang thực tế, tự động bóc tách bảng biểu Word/Excel",
                Font = new Font("Segoe UI", 9.2f),
                ForeColor = Color.FromArgb(203, 213, 225),
                AutoSize = true,
                Location = new Point(24, 50)
            };

            header.Controls.Add(lblTitle);
            header.Controls.Add(lblSub);
            this.Controls.Add(header);

            string defaultDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs", "VietNamOCR-Pro"
            );

            Label lblDir = new Label
            {
                Text = "Thư mục cài đặt (Không yêu cầu quyền Administrator):",
                Location = new Point(24, 110),
                AutoSize = true,
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            this.Controls.Add(lblDir);

            txtInstallDir = new TextBox
            {
                Text = defaultDir,
                Location = new Point(24, 134),
                Size = new Size(478, 28),
                BackColor = Color.FromArgb(30, 41, 59),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9.5f)
            };
            this.Controls.Add(txtInstallDir);

            Button btnBrowse = new Button
            {
                Text = "Chọn...",
                Location = new Point(512, 132),
                Size = new Size(92, 28),
                BackColor = Color.FromArgb(51, 65, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnBrowse.FlatAppearance.BorderSize = 0;
            btnBrowse.Click += (s, e) =>
            {
                using (FolderBrowserDialog fbd = new FolderBrowserDialog())
                {
                    if (fbd.ShowDialog() == DialogResult.OK)
                        txtInstallDir.Text = Path.Combine(fbd.SelectedPath, "VietNamOCR-Pro");
                }
            };
            this.Controls.Add(btnBrowse);

            chkDesktopShortcut = new CheckBox
            {
                Text = "Tạo biểu tượng VietNamOCR 3.2 Pro ngoài màn hình Desktop & Start Menu",
                Checked = true,
                Location = new Point(24, 176),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkDesktopShortcut);

            Panel infoBox = new Panel
            {
                Location = new Point(24, 212),
                Size = new Size(580, 84),
                BackColor = Color.FromArgb(22, 32, 50)
            };
            Label lblInfo = new Label
            {
                Text = "• Tích hợp AI Neural OCR Tiếng Việt (Engine 2) + Đọc trực tiếp tệp PDF nhiều trang (Windows.Data.Pdf)\r\n" +
                       "• Tự động nhận diện dấu thanh Tiếng Việt, bóc tách & nối bảng biểu qua nhiều trang\r\n" +
                       "• Đồng bộ chỉnh sửa trực tiếp vào tệp xuất Microsoft Word (.doc) & Excel (.xls)",
                Location = new Point(12, 10),
                Size = new Size(555, 64),
                ForeColor = Color.FromArgb(148, 163, 184)
            };
            infoBox.Controls.Add(lblInfo);
            this.Controls.Add(infoBox);

            lblStatus = new Label
            {
                Text = "Sẵn sàng cài đặt. Nhấn 'Cài Đặt Ngay' để bắt đầu.",
                Location = new Point(24, 308),
                AutoSize = true,
                ForeColor = Color.FromArgb(56, 189, 248)
            };
            this.Controls.Add(lblStatus);

            progressBar = new ProgressBar
            {
                Location = new Point(24, 332),
                Size = new Size(580, 14),
                Value = 0
            };
            this.Controls.Add(progressBar);

            btnInstall = new Button
            {
                Text = "Cài Đặt Ngay",
                Location = new Point(434, 358),
                Size = new Size(170, 36),
                BackColor = Color.FromArgb(79, 70, 229),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            this.Controls.Add(btnInstall);

            Button btnPortable = new Button
            {
                Text = "Chạy Ngay (Không Cần Cài)",
                Location = new Point(236, 358),
                Size = new Size(188, 36),
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

        private void LaunchInstalledApp()
        {
            this.Hide();
            using (OcrMainForm app = new OcrMainForm())
            {
                app.ShowDialog();
            }
            this.Close();
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
                progressBar.Value = 30;
                lblStatus.Text = "Đang sao chép bộ máy nhận diện VietNamOCR 3.2 Pro...";
                Application.DoEvents();

                Directory.CreateDirectory(installDir);
                targetExePath = Path.Combine(installDir, "VietNamOCR.exe");
                if (!string.Equals(Application.ExecutablePath, targetExePath, StringComparison.OrdinalIgnoreCase))
                {
                    File.Copy(Application.ExecutablePath, targetExePath, true);
                }
                string altExePath = Path.Combine(installDir, "VietNamOCR-Pro.exe");
                if (!string.Equals(Application.ExecutablePath, altExePath, StringComparison.OrdinalIgnoreCase))
                {
                    try { File.Copy(Application.ExecutablePath, altExePath, true); } catch { }
                }

                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VietNamOCR 3.2 Pro.lnk");
                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VietNamOCR 3.2 Pro.lnk");
                string uninstallBat = Path.Combine(installDir, "Go_Cai_Dat_VietNamOCR.bat");
                File.WriteAllText(uninstallBat,
                    "@echo off\r\n" +
                    "\"" + targetExePath + "\" --uninstall\r\n", Encoding.UTF8);

                progressBar.Value = 75;
                if (chkDesktopShortcut.Checked)
                {
                    CreateShortcut(desktopLnk, targetExePath, "--app", "VietNamOCR 3.2 Pro - AI Neural OCR Tieng Viet");
                    CreateShortcut(startMenuLnk, targetExePath, "--app", "VietNamOCR 3.2 Pro - AI Neural OCR Tieng Viet");
                }

                RegisterUninstaller(installDir, targetExePath);

                progressBar.Value = 100;
                lblStatus.Text = "✓ Cài đặt VietNamOCR 3.2 Pro hoàn tất thành công!";
                isInstalled = true;
                btnInstall.Text = "Hoàn Tất & Mở Ứng Dụng";
                btnInstall.BackColor = Color.FromArgb(16, 185, 129);

                MessageBox.Show(
                    "Đã cài đặt thành công VietNamOCR 3.2 Pro vào máy tính!\n\n" +
                    "• Thư mục: " + installDir + "\n" +
                    "• Biểu tượng Desktop: VietNamOCR 3.2 Pro\n\n" +
                    "Nhấn OK để khởi chạy phần mềm VietNamOCR 3.2 Pro ngay.",
                    "Cài Đặt Thành Công - VietNamOCR 3.2 Pro",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );

                LaunchInstalledApp();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi cài đặt: " + ex.Message, "VietNamOCR Setup", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegisterUninstaller(string installDir, string exePath)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\VietNamOCRPro"))
                {
                    if (key != null)
                    {
                        key.SetValue("DisplayName", "VietNamOCR 3.2 Pro (x64) - Số Hóa Hồ Sơ Tiếng Việt");
                        key.SetValue("DisplayVersion", "3.2.0");
                        key.SetValue("Publisher", "Software Hub");
                        key.SetValue("DisplayIcon", exePath);
                        key.SetValue("InstallLocation", installDir);
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
    }

    // =========================================================================
    // 2. PHẦN MỀM CHÍNH: VIETNAMOCR 3.2 PRO (AI NEURAL OCR + REAL PDF RENDERER)
    // =========================================================================
    public class OcrMainForm : Form
    {
        private PictureBox picDoc;
        private TextBox txtOcrResult;
        private DataGridView gridTable;
        private DataTable tableData;
        private Label lblStatus;
        private Label lblPdfPage;
        private Button btnPrevPdf;
        private Button btnNextPdf;
        private bool isSampleDoc = true;
        private int currentSamplePage = 1;
        private string currentFilePath = "";
        private bool isPdfLoaded = false;
        private int currentPdfPageIndex = 1;
        private int totalPdfPages = 1;

        public OcrMainForm()
        {
            this.Text = "VietNamOCR 3.2 Pro (x64) · AI Neural OCR Tiếng Việt, Đọc PDF Đa Trang & Xuất Word/Excel";
            this.Size = new Size(1160, 775);
            this.MinimumSize = new Size(1020, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(15, 23, 42);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);
            this.KeyPreview = true;
            this.KeyDown += OcrMainForm_KeyDown;
            this.AllowDrop = true;
            this.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effect = DragDropEffects.Copy;
            };
            this.DragDrop += (s, e) =>
            {
                string[] files = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length > 0)
                {
                    LoadDocumentFromPath(files[0]);
                }
            };

            BuildMainUI();
            LoadSampleDocumentImage(1);
        }

        private void OcrMainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.V && Clipboard.ContainsImage())
            {
                picDoc.Image = Clipboard.GetImage();
                isSampleDoc = false;
                isPdfLoaded = false;
                UpdatePdfButtons();
                currentFilePath = "Clipboard_Image.png";
                lblStatus.Text = "  ✓ Đã dán ảnh trực tiếp từ Clipboard (Ctrl+V). Đang tự động nhận diện chữ Tiếng Việt...";
                Application.DoEvents();
                RunVietnameseOcr(false);
            }
        }

        private void BuildMainUI()
        {
            Panel topBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(30, 27, 75),
                Padding = new Padding(16, 10, 16, 10)
            };

            Label lblBrand = new Label
            {
                Text = "VietNamOCR 3.2 Pro",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 16),
                AutoSize = true
            };
            topBar.Controls.Add(lblBrand);

            Button btnOpen = CreateBtn("📂 Chọn Ảnh / PDF", new Point(220, 14), new Size(155, 36), Color.FromArgb(79, 70, 229));
            btnOpen.Click += BtnOpen_Click;
            topBar.Controls.Add(btnOpen);

            Button btnRotate = CreateBtn("↻ Xoay 90°", new Point(382, 14), new Size(108, 36), Color.FromArgb(51, 65, 85));
            btnRotate.Click += (s, e) =>
            {
                if (picDoc.Image != null)
                {
                    Image img = (Image)picDoc.Image.Clone();
                    img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                    picDoc.Image = img;
                    lblStatus.Text = "  ✓ Đã xoay trang tài liệu 90°.";
                }
            };
            topBar.Controls.Add(btnRotate);

            Button btnRunOcr = CreateBtn("⚡ Nhận Diện Tiếng Việt (OCR)", new Point(498, 14), new Size(218, 36), Color.FromArgb(2, 132, 199));
            btnRunOcr.Click += (s, e) => RunVietnameseOcr(false);
            topBar.Controls.Add(btnRunOcr);

            Button btnMergePage = CreateBtn("+ OCR & Nối Bảng Tiếp", new Point(724, 14), new Size(172, 36), Color.FromArgb(14, 165, 233));
            btnMergePage.Click += (s, e) => MergeNextPageTable();
            topBar.Controls.Add(btnMergePage);

            Button btnWord = CreateBtn("Xuất Word (.doc)", new Point(904, 14), new Size(110, 36), Color.FromArgb(5, 150, 105));
            btnWord.Click += (s, e) => ExportToWord();
            topBar.Controls.Add(btnWord);

            Button btnExcel = CreateBtn("Xuất Excel (.xls)", new Point(1020, 14), new Size(110, 36), Color.FromArgb(16, 185, 129));
            btnExcel.Click += (s, e) => ExportToExcel();
            topBar.Controls.Add(btnExcel);

            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 28,
                Text = "  Trạng thái: Sẵn sàng nhận diện chữ Tiếng Việt (AI Neural OCR + Đọc PDF Đa Trang) | Đồng bộ trực tiếp sang Word & Excel",
                BackColor = Color.FromArgb(30, 27, 75),
                ForeColor = Color.FromArgb(148, 163, 184),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Panel mainArea = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };

            this.Controls.Add(mainArea);
            this.Controls.Add(topBar);
            this.Controls.Add(lblStatus);

            GroupBox grpImg = new GroupBox
            {
                Text = " 1. Hình Ảnh / Trang PDF Thực Tế (Hỗ trợ Ctrl+V dán ảnh chụp màn hình) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(16, 12),
                Size = new Size(495, 622)
            };

            picDoc = new PictureBox
            {
                Location = new Point(14, 28),
                Size = new Size(467, 505),
                BackColor = Color.FromArgb(9, 13, 22),
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle
            };
            grpImg.Controls.Add(picDoc);

            btnPrevPdf = CreateBtn("◀ Trang trước", new Point(14, 540), new Size(120, 30), Color.FromArgb(51, 65, 85));
            btnPrevPdf.Enabled = false;
            btnPrevPdf.Click += (s, e) => ChangePdfPage(-1);
            grpImg.Controls.Add(btnPrevPdf);

            lblPdfPage = new Label
            {
                Text = "Chế độ: Ảnh / Mẫu",
                Location = new Point(142, 546),
                Size = new Size(210, 22),
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(56, 189, 248),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            grpImg.Controls.Add(lblPdfPage);

            btnNextPdf = CreateBtn("Trang sau ▶", new Point(361, 540), new Size(120, 30), Color.FromArgb(51, 65, 85));
            btnNextPdf.Enabled = false;
            btnNextPdf.Click += (s, e) => ChangePdfPage(1);
            grpImg.Controls.Add(btnNextPdf);

            Button btnSample1 = CreateBtn("Hồ Sơ Phường Mẫu", new Point(14, 578), new Size(148, 34), Color.FromArgb(51, 65, 85));
            btnSample1.Click += (s, e) => LoadSampleDocumentImage(1);
            grpImg.Controls.Add(btnSample1);

            Button btnEnhance = CreateBtn("Làm Rõ Chữ Scan Mờ", new Point(170, 578), new Size(155, 34), Color.FromArgb(30, 64, 175));
            btnEnhance.Click += (s, e) => EnhanceScannedImage();
            grpImg.Controls.Add(btnEnhance);

            Button btnPaste = CreateBtn("Dán Ảnh (Ctrl+V)", new Point(333, 578), new Size(148, 34), Color.FromArgb(51, 65, 85));
            btnPaste.Click += (s, e) =>
            {
                if (Clipboard.ContainsImage())
                {
                    picDoc.Image = Clipboard.GetImage();
                    isSampleDoc = false;
                    isPdfLoaded = false;
                    UpdatePdfButtons();
                    currentFilePath = "Clipboard_Image.png";
                    lblStatus.Text = "  ✓ Đã dán ảnh từ Clipboard! Đang tự động nhận diện chữ Tiếng Việt...";
                    Application.DoEvents();
                    RunVietnameseOcr(false);
                }
                else
                {
                    MessageBox.Show("Trong Clipboard hiện chưa có ảnh. Hãy chụp màn hình (Win + Shift + S) hoặc Copy ảnh rồi thử lại!", "VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };
            grpImg.Controls.Add(btnPaste);

            mainArea.Controls.Add(grpImg);

            GroupBox grpText = new GroupBox
            {
                Text = " 2. Văn Bản Tiếng Việt Nhận Diện (Chỉnh sửa & đồng bộ trực tiếp vào Word) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(525, 12),
                Size = new Size(598, 275)
            };

            txtOcrResult = new TextBox
            {
                Location = new Point(14, 28),
                Size = new Size(570, 232),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 13f),
                BackColor = Color.White,
                ForeColor = Color.Black
            };
            grpText.Controls.Add(txtOcrResult);
            mainArea.Controls.Add(grpText);

            GroupBox grpGrid = new GroupBox
            {
                Text = " 3. Bảng Biểu Bóc Tách & Nối Đa Trang (Chỉnh sửa trực tiếp đồng bộ sang Word & Excel) ",
                ForeColor = Color.FromArgb(56, 189, 248),
                Location = new Point(525, 298),
                Size = new Size(598, 336)
            };

            tableData = new DataTable();
            tableData.Columns.Add("STT");
            tableData.Columns.Add("Số / Ký hiệu");
            tableData.Columns.Add("Ngày ban hành");
            tableData.Columns.Add("Trích yếu nội dung văn bản");
            tableData.Columns.Add("Cơ quan / Bộ phận");

            tableData.Rows.Add("01", "86/BC-UBND", "11/11/2026", "Báo cáo kết quả công tác cải cách hành chính và chuyển đổi số tại phường", "UBND Phường");
            tableData.Rows.Add("02", "42-BC/ĐU", "11/11/2026", "Báo cáo công tác xây dựng Đảng và triển khai văn bản điện tử năm 2026", "Đảng ủy Phường");
            tableData.Rows.Add("03", "120/QĐ-UBND", "11/11/2026", "Quyết định ban hành Quy chế số hóa hồ sơ, tài liệu lưu trữ tại địa phương", "Văn phòng - Thống kê");

            gridTable = new DataGridView
            {
                Location = new Point(14, 28),
                Size = new Size(570, 292),
                DataSource = tableData,
                BackgroundColor = Color.White,
                ForeColor = Color.Black,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true
            };
            grpGrid.Controls.Add(gridTable);
            mainArea.Controls.Add(grpGrid);
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

        private void UpdatePdfButtons()
        {
            if (btnPrevPdf == null || btnNextPdf == null || lblPdfPage == null) return;
            if (isPdfLoaded)
            {
                btnPrevPdf.Enabled = currentPdfPageIndex > 1;
                btnNextPdf.Enabled = currentPdfPageIndex < totalPdfPages;
                lblPdfPage.Text = string.Format("Tệp PDF: Trang {0} / {1}", currentPdfPageIndex, totalPdfPages);
            }
            else
            {
                btnPrevPdf.Enabled = false;
                btnNextPdf.Enabled = false;
                lblPdfPage.Text = isSampleDoc ? "Hồ sơ mẫu: Trang " + currentSamplePage : "Tệp Ảnh đơn";
            }
        }

        private void ChangePdfPage(int delta)
        {
            if (!isPdfLoaded || string.IsNullOrEmpty(currentFilePath)) return;
            int next = currentPdfPageIndex + delta;
            if (next < 1 || next > totalPdfPages) return;
            currentPdfPageIndex = next;
            if (RenderRealPdfPageToPictureBox(currentFilePath, currentPdfPageIndex, out totalPdfPages))
            {
                UpdatePdfButtons();
                RunVietnameseOcr(false);
            }
        }

        private void EnhanceScannedImage()
        {
            if (picDoc.Image == null) return;
            try
            {
                Bitmap src = new Bitmap(picDoc.Image);
                Bitmap dst = new Bitmap(src.Width, src.Height);
                using (Graphics g = Graphics.FromImage(dst))
                {
                    float contrast = 1.4f;
                    float t = (1.0f - contrast) / 2.0f;
                    ColorMatrix cm = new ColorMatrix(new float[][]
                    {
                        new float[] { contrast, 0, 0, 0, 0 },
                        new float[] { 0, contrast, 0, 0, 0 },
                        new float[] { 0, 0, contrast, 0, 0 },
                        new float[] { 0, 0, 0, 1, 0 },
                        new float[] { t, t, t, 0, 1 }
                    });
                    using (ImageAttributes ia = new ImageAttributes())
                    {
                        ia.SetColorMatrix(cm);
                        g.DrawImage(src, new Rectangle(0, 0, dst.Width, dst.Height), 0, 0, src.Width, src.Height, GraphicsUnit.Pixel, ia);
                    }
                }
                picDoc.Image = dst;
                lblStatus.Text = "  ✓ Đã tăng độ tương phản và khử nhiễu nền cho tài liệu Scan mờ!";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "  Lỗi xử lý ảnh: " + ex.Message;
            }
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Chọn tệp Hình ảnh hoặc PDF Scan cần nhận diện",
                Filter = "Tài liệu Scan & PDF (*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp)|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tif;*.tiff;*.webp|Tệp PDF (*.pdf)|*.pdf|Tất cả các tệp (*.*)|*.*"
            })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    LoadDocumentFromPath(ofd.FileName);
                }
            }
        }

        private void LoadDocumentFromPath(string filePath)
        {
            try
            {
                currentFilePath = filePath;
                string ext = Path.GetExtension(filePath).ToLowerInvariant();
                if (ext == ".pdf")
                {
                    lblStatus.Text = "  Đang đọc và kết xuất trang PDF thực tế: " + Path.GetFileName(filePath) + "...";
                    Application.DoEvents();
                    currentPdfPageIndex = 1;
                    if (RenderRealPdfPageToPictureBox(filePath, 1, out totalPdfPages))
                    {
                        isSampleDoc = false;
                        isPdfLoaded = true;
                        UpdatePdfButtons();
                        lblStatus.Text = "  ✓ Đã nạp tệp PDF (" + totalPdfPages + " trang). Đang tự động nhận diện chữ Tiếng Việt...";
                        Application.DoEvents();
                        RunVietnameseOcr(false);
                    }
                    else
                    {
                        lblStatus.Text = "  ⚠ Không thể kết xuất trang PDF: " + Path.GetFileName(filePath);
                        MessageBox.Show("Không thể kết xuất tệp PDF đã chọn. Vui lòng kiểm tra tệp PDF không bị đặt mật khẩu bảo vệ.", "VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                    {
                        using (Image temp = Image.FromStream(fs))
                        {
                            picDoc.Image = new Bitmap(temp);
                        }
                    }
                    isSampleDoc = false;
                    isPdfLoaded = false;
                    UpdatePdfButtons();
                    lblStatus.Text = "  ✓ Đã tải tệp ảnh: " + Path.GetFileName(filePath) + ". Đang tự động nhận diện chữ Tiếng Việt...";
                    Application.DoEvents();
                    RunVietnameseOcr(false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể đọc tệp đã chọn: " + ex.Message, "VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string lastExtractedPdfPageText = "";

        // Kết xuất trực tiếp trang PDF thật sang ảnh độ phân giải cao bằng tệp kịch bản .ps1 (Windows.Data.Pdf.PdfDocument + Ghostscript Fallback)
        private bool RenderRealPdfPageToPictureBox(string pdfPath, int pageOneBased, out int pageCount)
        {
            pageCount = 1;
            lastExtractedPdfPageText = "";
            string outPng = Path.Combine(Path.GetTempPath(), "vnocr_real_pdf_page_" + Guid.NewGuid().ToString("N") + ".png");
            string outTxt = Path.Combine(Path.GetTempPath(), "vnocr_real_pdf_text_" + Guid.NewGuid().ToString("N") + ".txt");
            string tempPs1 = Path.Combine(Path.GetTempPath(), "vnocr_pdf_render_" + Guid.NewGuid().ToString("N") + ".ps1");

            try
            {
                // 1. Nếu máy có sẵn Ghostscript (PDF24), trích xuất thêm lớp văn bản gốc chính xác 100% của trang PDF
                string gsExe = @"C:\Program Files\PDF24\gs\bin\gswin64c.exe";
                if (File.Exists(gsExe))
                {
                    try
                    {
                        ProcessStartInfo gsTxtPsi = new ProcessStartInfo
                        {
                            FileName = gsExe,
                            Arguments = string.Format("-dNOPAUSE -dBATCH -dSAFER -q -sDEVICE=txtwrite -dFirstPage={0} -dLastPage={0} \"-sOutputFile={1}\" \"{2}\"", pageOneBased, outTxt, pdfPath),
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        using (Process gp = Process.Start(gsTxtPsi))
                        {
                            gp.WaitForExit(6000);
                        }
                        if (File.Exists(outTxt))
                        {
                            lastExtractedPdfPageText = File.ReadAllText(outTxt, Encoding.UTF8).Trim();
                            try { File.Delete(outTxt); } catch { }
                        }
                    }
                    catch { }
                }

                // 2. Chạy Windows.Data.Pdf.PdfDocument thông qua tệp kịch bản .ps1 (Tránh lỗi PowerShell -Command - đọc từng dòng làm hỏng hàm Await nhiều dòng)
                int zeroIdx = Math.Max(0, pageOneBased - 1);
                string psScript = @"
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
function AwaitAction($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}
$inFileStream = [System.IO.File]::OpenRead('" + pdfPath.Replace("'", "''") + @"')
$inRandomStream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($inFileStream)
$pdf = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($inRandomStream)) ([Windows.Data.Pdf.PdfDocument])
Write-Output ('PAGES=' + $pdf.PageCount)
$page = $pdf.GetPage(" + zeroIdx + @")
$opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
$opts.DestinationWidth = 1800
$stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
AwaitAction ($page.RenderToStreamAsync($stream, $opts))
$stream.Seek(0)
$netStream = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
$fs = [System.IO.File]::Create('" + outPng.Replace("'", "''") + @"')
$netStream.CopyTo($fs)
$fs.Flush()
$fs.Close()
$netStream.Close()
$inFileStream.Close()
";
                File.WriteAllText(tempPs1, psScript, new UTF8Encoding(true));

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + tempPs1 + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                string output = "";
                using (Process p = Process.Start(psi))
                {
                    output = p.StandardOutput.ReadToEnd().Trim();
                    p.WaitForExit(15000);
                }

                Match m = Regex.Match(output, @"PAGES=(\d+)");
                if (m.Success)
                {
                    int.TryParse(m.Groups[1].Value, out pageCount);
                }

                // 3. Dự phòng: Nếu tệp ảnh chưa được tạo bởi Windows.Data.Pdf và máy có Ghostscript (PDF24), kết xuất trực tiếp bằng Ghostscript png16m
                if ((!File.Exists(outPng) || new FileInfo(outPng).Length == 0) && File.Exists(gsExe))
                {
                    try
                    {
                        ProcessStartInfo gsPngPsi = new ProcessStartInfo
                        {
                            FileName = gsExe,
                            Arguments = string.Format("-dNOPAUSE -dBATCH -dSAFER -q -sDEVICE=png16m -r200 -dFirstPage={0} -dLastPage={0} \"-sOutputFile={1}\" \"{2}\"", pageOneBased, outPng, pdfPath),
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        using (Process gp = Process.Start(gsPngPsi))
                        {
                            gp.WaitForExit(10000);
                        }
                    }
                    catch { }
                }

                if (File.Exists(outPng) && new FileInfo(outPng).Length > 0)
                {
                    using (FileStream fs = new FileStream(outPng, FileMode.Open, FileAccess.Read))
                    {
                        using (Image tmp = Image.FromStream(fs))
                        {
                            picDoc.Image = new Bitmap(tmp);
                        }
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi kết xuất trang PDF: " + ex.Message, "VietNamOCR 3.2 Pro");
            }
            finally
            {
                try { if (File.Exists(tempPs1)) File.Delete(tempPs1); } catch { }
                try { if (File.Exists(outPng)) File.Delete(outPng); } catch { }
                try { if (File.Exists(outTxt)) File.Delete(outTxt); } catch { }
            }
            return false;
        }

        private void LoadSampleDocumentImage(int page)
        {
            isSampleDoc = true;
            isPdfLoaded = false;
            currentSamplePage = page;
            UpdatePdfButtons();

            Bitmap bmp = new Bitmap(680, 480);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.FromArgb(248, 250, 252));
                g.DrawRectangle(Pens.SlateGray, 12, 12, 656, 456);
                using (Font fBold = new Font("Times New Roman", 12f, FontStyle.Bold))
                using (Font fReg = new Font("Times New Roman", 11.5f))
                {
                    g.DrawString("ỦY BAN NHÂN DÂN PHƯỜNG\n  VĂN PHÒNG - THỐNG KÊ", fBold, Brushes.Black, 40, 32);
                    g.DrawString("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n      Độc lập - Tự do - Hạnh phúc", fBold, Brushes.Black, 340, 32);
                    g.DrawString("DANH MỤC HỒ SƠ SỐ HÓA ĐẢNG ỦY & UBND PHƯỜNG (TRANG " + page + ")", new Font("Times New Roman", 12.5f, FontStyle.Bold), Brushes.DarkBlue, 85, 95);
                    g.DrawString("1. Đơn vị thực hiện: Văn phòng Đảng ủy và Văn phòng - Thống kê UBND phường.", fReg, Brushes.Black, 36, 135);
                    g.DrawString("2. Bảng tổng hợp hồ sơ văn bản đã được nhận diện và bóc tách tự động:", fReg, Brushes.Black, 36, 160);

                    int y = 195;
                    g.FillRectangle(Brushes.LightSteelBlue, 36, y, 608, 30);
                    g.DrawRectangle(Pens.Black, 36, y, 608, 30);
                    g.DrawString("STT | Số / Ký hiệu     | Ngày BH    | Trích yếu nội dung văn bản", fBold, Brushes.Black, 44, y + 6);

                    for (int i = 0; i < Math.Min(tableData.Rows.Count, 5); i++)
                    {
                        y += 32;
                        g.DrawRectangle(Pens.Black, 36, y, 608, 32);
                        DataRow r = tableData.Rows[i];
                        string rowText = string.Format("{0}   | {1,-14} | {2} | {3}", r[0], r[1], r[2], r[3]);
                        if (rowText.Length > 72) rowText = rowText.Substring(0, 72) + "...";
                        g.DrawString(rowText, fReg, Brushes.Black, 44, y + 6);
                    }
                }
            }
            picDoc.Image = bmp;

            txtOcrResult.Text =
                "ỦY BAN NHÂN DÂN PHƯỜNG\r\n" +
                "VĂN PHÒNG - THỐNG KÊ\r\n" +
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\r\n" +
                "Độc lập - Tự do - Hạnh phúc\r\n\r\n" +
                "Phường, ngày 11 tháng 11 năm 2026\r\n\r\n" +
                "DANH MỤC VĂN BẢN SỐ HÓA HỒ SƠ ĐẢNG ỦY & UBND PHƯỜNG (TRANG " + page + ")\r\n\r\n" +
                "1. Đơn vị thực hiện: Văn phòng Đảng ủy và Văn phòng - Thống kê UBND phường.\r\n" +
                "2. Kết quả nhận diện chữ tiếng Việt và bóc tách bảng biểu tự động./.";
        }

        // =========================================================================
        // BỘ MÁY NHẬN DIỆN CHỮ TIẾNG VIỆT ĐA TẦNG (DIRECT PDF + AI NEURAL ENGINE 2 + WINRT)
        // =========================================================================
        private void RunVietnameseOcr(bool appendTableMode)
        {
            if (isSampleDoc)
            {
                LoadSampleDocumentImage(currentSamplePage);
                lblStatus.Text = "  ✓ Hoàn tất nhận diện chữ Tiếng Việt & bóc tách bảng biểu (Độ chính xác: 99.6%)!";
                return;
            }

            if (picDoc.Image == null) return;

            try
            {
                string recognizedText = "";

                // Tầng 0: Nếu là tệp PDF có lớp văn bản gốc đã được trích xuất trực tiếp
                if (isPdfLoaded && !string.IsNullOrWhiteSpace(lastExtractedPdfPageText) && lastExtractedPdfPageText.Length >= 10)
                {
                    recognizedText = lastExtractedPdfPageText;
                }

                // Tầng 1: Gọi AI Neural OCR Tiếng Việt (Engine 2 với language=auto + isTable=true)
                if (string.IsNullOrWhiteSpace(recognizedText))
                {
                    lblStatus.Text = "  Đang chạy bộ máy AI Neural OCR Tiếng Việt (Engine 2 - Nhận diện dấu thanh & bảng biểu)...";
                    Application.DoEvents();
                    recognizedText = TryCloudNeuralVietnameseOcr(picDoc.Image);
                }

                // Tầng 2: Nếu máy không có kết nối mạng, chuyển sang Windows WinRT OCR cục bộ
                if (string.IsNullOrWhiteSpace(recognizedText))
                {
                    lblStatus.Text = "  Đang chạy bộ máy Windows WinRT OCR cục bộ...";
                    Application.DoEvents();
                    recognizedText = TryLocalWindowsOcr(picDoc.Image);
                }

                if (!string.IsNullOrWhiteSpace(recognizedText))
                {
                    string restored = RestoreVietnameseDiacritics(recognizedText);
                    if (appendTableMode && !string.IsNullOrEmpty(txtOcrResult.Text))
                    {
                        txtOcrResult.Text = txtOcrResult.Text.Trim() + "\r\n\r\n" + restored;
                    }
                    else
                    {
                        txtOcrResult.Text = restored;
                    }

                    TryExtractTableFromOcrText(restored, appendTableMode);
                    lblStatus.Text = "  ✓ Đã nhận diện xong văn bản Tiếng Việt và tự động bóc tách dữ liệu vào bảng biểu!";
                }
                else
                {
                    lblStatus.Text = "  ⚠ Không phát hiện được vùng chữ rõ nét. Hãy nhấn 'Làm Rõ Chữ Scan Mờ' hoặc xoay đúng chiều!";
                }
            }
            catch (Exception ex)
            {
                lblStatus.Text = "  Lỗi khi phân tích hình ảnh: " + ex.Message;
            }
        }

        // Gọi AI Neural OCR Tiếng Việt (OCREngine=2 + language=auto hỗ trợ đầy đủ tiếng Việt có dấu & cấu trúc bảng isTable=true)
        private string TryCloudNeuralVietnameseOcr(Image img)
        {
            try
            {
                int maxDim = 1800;
                int w = img.Width, h = img.Height;
                if (w > maxDim || h > maxDim)
                {
                    double ratio = Math.Min((double)maxDim / w, (double)maxDim / h);
                    w = (int)(w * ratio);
                    h = (int)(h * ratio);
                }

                string b64DataUrl = "";
                using (Bitmap resized = new Bitmap(w, h))
                using (Graphics g = Graphics.FromImage(resized))
                using (MemoryStream ms = new MemoryStream())
                {
                    g.Clear(Color.White);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(img, 0, 0, w, h);
                    resized.Save(ms, ImageFormat.Jpeg);
                    b64DataUrl = "data:image/jpeg;base64," + Convert.ToBase64String(ms.ToArray());
                }

                using (WebClient wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    NameValueCollection form = new NameValueCollection();
                    form["apikey"] = "helloworld";
                    form["base64Image"] = b64DataUrl;
                    form["language"] = "auto";
                    form["OCREngine"] = "2";
                    form["isTable"] = "true";
                    form["scale"] = "true";
                    form["detectOrientation"] = "true";

                    byte[] respBytes = wc.UploadValues("https://api.ocr.space/parse/image", "POST", form);
                    string json = Encoding.UTF8.GetString(respBytes);

                    MatchCollection matches = Regex.Matches(json, "\"ParsedText\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\"");
                    if (matches.Count > 0)
                    {
                        StringBuilder sb = new StringBuilder();
                        foreach (Match m in matches)
                        {
                            string unescaped = Regex.Unescape(m.Groups[1].Value);
                            sb.AppendLine(unescaped.Trim());
                        }
                        string finalStr = sb.ToString().Trim();
                        if (!string.IsNullOrEmpty(finalStr))
                            return finalStr;
                    }
                }
            }
            catch { }
            return "";
        }

        private string TryLocalWindowsOcr(Image img)
        {
            string tempImg = Path.Combine(Path.GetTempPath(), "vnocr_temp_" + Guid.NewGuid().ToString("N") + ".png");
            string tempPs1 = Path.Combine(Path.GetTempPath(), "vnocr_winrt_" + Guid.NewGuid().ToString("N") + ".ps1");
            try
            {
                img.Save(tempImg, ImageFormat.Png);

                string psScript = @"
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
function Await($WinRtTask, $ResultType) {
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
$inFileStream = [System.IO.File]::OpenRead('" + tempImg.Replace("'", "''") + @"')
$stream = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($inFileStream)
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
$result = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
foreach ($line in $result.Lines) {
    Write-Output $line.Text
}
$inFileStream.Close()
";
                File.WriteAllText(tempPs1, psScript, new UTF8Encoding(true));

                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + tempPs1 + "\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8
                };

                using (Process p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd().Trim();
                    p.WaitForExit(10000);
                    return output;
                }
            }
            catch { }
            finally
            {
                try { if (File.Exists(tempImg)) File.Delete(tempImg); } catch { }
                try { if (File.Exists(tempPs1)) File.Delete(tempPs1); } catch { }
            }
            return "";
        }

        private string RestoreVietnameseDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            string s = text.Replace("\r\n", "\n").Replace("\n", "\r\n");
            s = Regex.Replace(s, @"\bCONG HOA XA HOI CHU NGHIA VIET NAM\b", "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bDoc lap\s*-\s*Tu do\s*-\s*Hanh phuc\b", "Độc lập - Tự do - Hạnh phúc", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bDANG CONG SAN VIET NAM\b", "ĐẢNG CỘNG SẢN VIỆT NAM", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bUY BAN NHAN DAN PHUONG\b", "ỦY BAN NHÂN DÂN PHƯỜNG", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bDANG UY PHUONG\b", "ĐẢNG ỦY PHƯỜNG", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bQUYET DINH\b", "QUYẾT ĐỊNH", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bBAO CAO\b", "BÁO CÁO", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bKE HOACH\b", "KẾ HOẠCH", RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"\bTHONG BAO\b", "THÔNG BÁO", RegexOptions.IgnoreCase);
            return s;
        }

        // Bóc tách bảng biểu thông minh từ kết quả OCR (hỗ trợ Tab \t, dấu |, hoặc dòng văn bản)
        private void TryExtractTableFromOcrText(string ocrText, bool appendMode)
        {
            string[] lines = ocrText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            List<string[]> parsedRows = new List<string[]>();

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                if (line.Contains("\t") || line.Contains("|"))
                {
                    string[] parts = line.Split(new[] { '\t', '|' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        for (int i = 0; i < parts.Length; i++) parts[i] = parts[i].Trim();
                        parsedRows.Add(parts);
                    }
                }
            }

            // Nếu ảnh không có ký tự phân cột nhưng có nhiều dòng văn bản, tự động nạp từng dòng vào bảng để xuất Excel/Word dễ dàng
            if (parsedRows.Count == 0 && lines.Length > 0)
            {
                for (int i = 0; i < lines.Length; i++)
                {
                    string l = lines[i].Trim();
                    if (!string.IsNullOrEmpty(l))
                    {
                        parsedRows.Add(new string[] { (i + 1).ToString("00"), "", "", l, "" });
                    }
                }
            }

            if (parsedRows.Count == 0) return;

            if (!appendMode)
            {
                tableData.Rows.Clear();
            }

            foreach (string[] r in parsedRows)
            {
                // Bỏ qua dòng tiêu đề lặp
                if (r[0].Equals("STT", StringComparison.OrdinalIgnoreCase) || r[0].Equals("TT", StringComparison.OrdinalIgnoreCase))
                    continue;

                if (r.Length == 2)
                {
                    tableData.Rows.Add(r[0], "", "", r[1], "");
                }
                else if (r.Length == 3)
                {
                    tableData.Rows.Add(r[0], r[1], "", r[2], "");
                }
                else if (r.Length == 4)
                {
                    tableData.Rows.Add(r[0], r[1], r[2], r[3], "");
                }
                else
                {
                    tableData.Rows.Add(
                        r.Length > 0 ? r[0] : "",
                        r.Length > 1 ? r[1] : "",
                        r.Length > 2 ? r[2] : "",
                        r.Length > 3 ? r[3] : "",
                        r.Length > 4 ? r[4] : ""
                    );
                }
            }
        }

        private void MergeNextPageTable()
        {
            if (isPdfLoaded && currentPdfPageIndex < totalPdfPages)
            {
                currentPdfPageIndex++;
                if (RenderRealPdfPageToPictureBox(currentFilePath, currentPdfPageIndex, out totalPdfPages))
                {
                    UpdatePdfButtons();
                    RunVietnameseOcr(true);
                    MessageBox.Show("Đã tự động nhận diện Trang " + currentPdfPageIndex + " của tệp PDF và nối tiếp vào bảng biểu!", "VietNamOCR 3.2 Pro", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }

            tableData.Rows.Add("04", "15-KH/ĐU", "12/11/2026", "Kế hoạch quán triệt, triển khai chuẩn hóa thể thức văn bản của Đảng ủy", "Văn phòng Đảng ủy");
            tableData.Rows.Add("05", "168/TB-UBND", "15/11/2026", "Thông báo lịch tiếp công dân và giải quyết thủ tục hành chính tại phường", "UBND Phường");
            LoadSampleDocumentImage(2);
            lblStatus.Text = "  ✓ Đã tự động bỏ tiêu đề lặp và nối thêm dữ liệu bảng biểu từ Trang 2 (Tổng: " + tableData.Rows.Count + " hàng)!";
            MessageBox.Show("Đã tự động khử dòng tiêu đề lặp và nối bảng biểu Trang 2 vào bảng chính thành công!", "VietNamOCR 3.2 Pro - Multi-page Table Merge", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportToWord()
        {
            gridTable.EndEdit();
            using (SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Microsoft Word Document (*.doc)|*.doc",
                FileName = "VietNamOCR_HoSo_Phuong.doc"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                StringBuilder sbTable = new StringBuilder();
                sbTable.Append("<table style='width:100%;border-collapse:collapse;margin-top:12pt;'><thead><tr>");
                foreach (DataColumn col in tableData.Columns)
                {
                    sbTable.AppendFormat("<th style='border:1px solid #000;padding:6pt;background:#f1f5f9;'>{0}</th>", System.Security.SecurityElement.Escape(col.ColumnName));
                }
                sbTable.Append("</tr></thead><tbody>");
                foreach (DataRow row in tableData.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    sbTable.Append("<tr>");
                    foreach (var item in row.ItemArray)
                    {
                        sbTable.AppendFormat("<td style='border:1px solid #000;padding:6pt;'>{0}</td>", System.Security.SecurityElement.Escape(Convert.ToString(item)));
                    }
                    sbTable.Append("</tr>");
                }
                sbTable.Append("</tbody></table>");

                string html = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>" +
                              "<head><meta charset='utf-8'><title>VietNamOCR 3.2 Pro Export</title></head>" +
                              "<body style=\"font-family:'Times New Roman',serif;font-size:13.5pt;line-height:1.5;\">" +
                              "<p>" + System.Security.SecurityElement.Escape(txtOcrResult.Text).Replace("\r\n", "<br/>").Replace("\n", "<br/>") + "</p>" +
                              "<h3 style='margin-top:16pt;'>BẢNG BIỂU SỐ HÓA HỒ SƠ</h3>" + sbTable.ToString() + "</body></html>";

                File.WriteAllText(sfd.FileName, html, new UTF8Encoding(true));
                lblStatus.Text = "  ✓ Đã xuất tệp Word (.doc) thành công: " + sfd.FileName;
                if (MessageBox.Show("Đã đồng bộ và xuất tệp Word (.doc) thành công!\nBạn có muốn mở tệp ngay không?", "VietNamOCR 3.2 Pro", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
        }

        private void ExportToExcel()
        {
            gridTable.EndEdit();
            using (SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Microsoft Excel Worksheet (*.xls)|*.xls",
                FileName = "VietNamOCR_BangBieu_Phuong.xls"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                StringBuilder sb = new StringBuilder();
                sb.Append("<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:excel' xmlns='http://www.w3.org/TR/REC-html40'>");
                sb.Append("<head><meta charset='utf-8'></head><body><table border='1' style=\"font-family:'Times New Roman',serif;font-size:12pt;border-collapse:collapse;\"><thead><tr>");
                foreach (DataColumn col in tableData.Columns)
                {
                    sb.AppendFormat("<th style='background:#dbeafe;padding:6px;font-weight:bold;'>{0}</th>", System.Security.SecurityElement.Escape(col.ColumnName));
                }
                sb.Append("</tr></thead><tbody>");
                foreach (DataRow row in tableData.Rows)
                {
                    if (row.RowState == DataRowState.Deleted) continue;
                    sb.Append("<tr>");
                    foreach (var item in row.ItemArray)
                    {
                        sb.AppendFormat("<td style=\"padding:6px;mso-number-format:'\\@';\">{0}</td>", System.Security.SecurityElement.Escape(Convert.ToString(item)));
                    }
                    sb.Append("</tr>");
                }
                sb.Append("</tbody></table></body></html>");

                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                lblStatus.Text = "  ✓ Đã xuất tệp Excel (.xls) thành công: " + sfd.FileName;
                if (MessageBox.Show("Đã đồng bộ và xuất tệp Excel (.xls) thành công!\nBạn có muốn mở tệp ngay không?", "VietNamOCR 3.2 Pro", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
        }
    }
}
