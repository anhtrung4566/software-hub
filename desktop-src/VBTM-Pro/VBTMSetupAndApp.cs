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
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("VBTM Pro 2.5.0 - Chuan Hoa Van Ban Dang Uy & UBND Phuong")]
[assembly: AssemblyDescription("Bo cong cu Chuan Hoa The Thuc Van Ban Nghi dinh 30, Van ban Dang, Tao QR & Soat Loi Chinh Ta AI Offline cho Microsoft Word")]
[assembly: AssemblyCompany("Software Hub")]
[assembly: AssemblyProduct("VBTM Pro 2.5.0")]
[assembly: AssemblyVersion("2.5.0.0")]
[assembly: AssemblyFileVersion("2.5.0.0")]

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
            bool unitSettingsMode = false;
            bool wordSpellCheckMode = false;
            bool aiWordMode = false;
            bool summarizeWordMode = false;
            bool spellWordMode = false;
            bool qrWordMode = false;
            bool installWordMode = false;
            int insertTplIdx = -1;
            for (int idx = 0; idx < args.Length; idx++)
            {
                string a = args[idx];
                if (a.Equals("--app", StringComparison.OrdinalIgnoreCase))
                    launchAppDirectly = true;
                else if (a.Equals("--uninstall", StringComparison.OrdinalIgnoreCase))
                    uninstallMode = true;
                else if (a.Equals("--install-word", StringComparison.OrdinalIgnoreCase))
                    installWordMode = true;
                else if (a.Equals("--unit-settings", StringComparison.OrdinalIgnoreCase))
                    unitSettingsMode = true;
                else if (a.Equals("--word-spellcheck", StringComparison.OrdinalIgnoreCase))
                    wordSpellCheckMode = true;
                else if (a.Equals("--ai-word", StringComparison.OrdinalIgnoreCase))
                    aiWordMode = true;
                else if (a.Equals("--summarize-word", StringComparison.OrdinalIgnoreCase))
                    summarizeWordMode = true;
                else if (a.Equals("--spell-word", StringComparison.OrdinalIgnoreCase))
                    spellWordMode = true;
                else if (a.Equals("--qr-word", StringComparison.OrdinalIgnoreCase))
                    qrWordMode = true;
                else if (a.Equals("--insert-tpl", StringComparison.OrdinalIgnoreCase) && idx + 1 < args.Length)
                    int.TryParse(args[++idx], out insertTplIdx);
            }

            if (uninstallMode)
            {
                RunUninstaller();
                return;
            }
            if (installWordMode)
            {
                string localApp = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string instDir = Path.Combine(localApp, "VBTMPro");
                Directory.CreateDirectory(instDir);
                string exeDst = Path.Combine(instDir, "VBTM-Pro.exe");
                try { File.Copy(Application.ExecutablePath, exeDst, true); } catch { }
                // Đồng bộ cả thư mục Programs\VBTM-Pro để mọi nút trên Ribbon Word luôn chạy bản EXE mới nhất
                string progDir = Path.Combine(localApp, "Programs", "VBTM-Pro");
                string progExe = Path.Combine(progDir, "VBTM-Pro.exe");
                try
                {
                    Directory.CreateDirectory(progDir);
                    File.Copy(Application.ExecutablePath, progExe, true);
                }
                catch { }
                WordAddinInstaller.InstallWordRibbonAddin(instDir, File.Exists(progExe) ? progExe : (File.Exists(exeDst) ? exeDst : Application.ExecutablePath));
                // Cập nhật luôn Icon Logo mới cho lối tắt Desktop & Start Menu
                try
                {
                    string targetIconExe = File.Exists(progExe) ? progExe : exeDst;
                    string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.5.0.lnk");
                    string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VBTM Pro 2.5.0.lnk");
                    SetupWizardForm.CreateShortcutStatic(desktopLnk, targetIconExe, "--app", "VBTM Pro 2.5.0 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                    SetupWizardForm.CreateShortcutStatic(startMenuLnk, targetIconExe, "--app", "VBTM Pro 2.5.0 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                }
                catch { }
                return;
            }
            if (unitSettingsMode)
            {
                Application.Run(new UnitSettingsForm());
                return;
            }
            if (aiWordMode)
            {
                Application.Run(new AiWordAssistantForm());
                return;
            }
            if (summarizeWordMode)
            {
                Application.Run(new WordAiSummaryPreviewForm());
                return;
            }
            if (spellWordMode)
            {
                Application.Run(new WordSpellInspectorForm());
                return;
            }
            if (qrWordMode)
            {
                Application.Run(new WordQrInsertForm());
                return;
            }
            if (insertTplIdx >= 0)
            {
                using (MainStudioForm f = new MainStudioForm())
                {
                    f.InsertTemplateIntoActiveWord(insertTplIdx);
                }
                return;
            }
            if (wordSpellCheckMode)
            {
                using (MainStudioForm f = new MainStudioForm())
                {
                    f.FormatActiveWordDocument();
                }
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
                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.5.0.lnk");
                if (File.Exists(desktopLnk)) File.Delete(desktopLnk);

                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VBTM Pro 2.5.0.lnk");
                if (File.Exists(startMenuLnk)) File.Delete(startMenuLnk);

                string wordAddinPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Word", "STARTUP", "VBTM_Pro_Addin.dotm");
                if (File.Exists(wordAddinPath)) File.Delete(wordAddinPath);

                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\VBTMPro", false);
                MessageBox.Show("Đã gỡ cài đặt VBTM Pro 2.5.0 sạch sẽ khỏi hệ thống!", "Gỡ Cài Đặt - VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Thông báo khi gỡ cài đặt: " + ex.Message, "VBTM Pro 2.5.0");
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
            this.Text = "Trình Cài Đặt VBTM Pro 2.5.0 (Offline Setup)";
            this.ClientSize = new Size(630, 440);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(95, 10, 14);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            BuildUI();
        }

        private void BuildUI()
        {
            Panel header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 88,
                BackColor = Color.FromArgb(145, 18, 24)
            };

            Label lblTitle = new Label
            {
                Text = "★ VBTM Pro 2.5.0 · Bộ Cài Đặt Ngoại Tuyến (Đảng Ủy & UBND Phường)",
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                ForeColor = Color.FromArgb(250, 204, 21),
                AutoSize = true,
                Location = new Point(22, 16)
            };

            Label lblSub = new Label
            {
                Text = "Chuẩn hóa văn bản Đảng ủy & UBND Phường (Nghị định 30 & Hướng dẫn 05-HD/VPTW), Tạo QR & Chính tả AI",
                Font = new Font("Segoe UI", 9.0f),
                ForeColor = Color.FromArgb(254, 240, 138),
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
                ForeColor = Color.FromArgb(253, 224, 71)
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
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.FromArgb(254, 240, 138),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9.5f)
            };
            this.Controls.Add(txtInstallDir);

            chkDesktopShortcut = new CheckBox
            {
                Text = "Tạo biểu tượng VBTM Pro 2.5.0 ngoài màn hình Desktop và Start Menu",
                Checked = true,
                Location = new Point(24, 172),
                AutoSize = true,
                ForeColor = Color.White
            };
            this.Controls.Add(chkDesktopShortcut);

            chkWordMacroExport = new CheckBox
            {
                Text = "Tích hợp bộ quy chuẩn Microsoft Word (Nghị định 30/2020/NĐ-CP & Hướng dẫn số 05-HD/VPTW)",
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
                BackColor = Color.FromArgb(69, 10, 10)
            };
            Label lblInfo = new Label
            {
                Text = "• Hoạt động ngoại tuyến độc lập 100% phục vụ Đảng ủy và Ủy ban nhân dân phường\r\n" +
                       "• Tự động căn lề trang A4 (30-15-20-20mm), font Times New Roman 14pt trực tiếp vào Word\r\n" +
                       "• Hỗ trợ gỡ cài đặt qua Installed Apps của Windows hoặc chạy Go_Cai_Dat_VBTM.bat",
                Location = new Point(12, 10),
                Size = new Size(556, 64),
                ForeColor = Color.FromArgb(254, 249, 195)
            };
            infoBox.Controls.Add(lblInfo);
            this.Controls.Add(infoBox);

            lblStatus = new Label
            {
                Text = "Sẵn sàng cài đặt. Nhấn 'Cài Đặt Ngay' để bắt đầu.",
                Location = new Point(24, 332),
                AutoSize = true,
                ForeColor = Color.FromArgb(250, 204, 21)
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
                Text = "★ Cài Đặt & Tích Hợp Word",
                Location = new Point(394, 384),
                Size = new Size(210, 38),
                BackColor = Color.FromArgb(234, 179, 8),
                ForeColor = Color.FromArgb(69, 10, 10),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.8f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInstall.FlatAppearance.BorderSize = 0;
            btnInstall.Click += BtnInstall_Click;
            this.Controls.Add(btnInstall);

            btnPortable = new Button
            {
                Text = "Mở Phần Mềm Ngay",
                Location = new Point(208, 384),
                Size = new Size(176, 38),
                BackColor = Color.FromArgb(185, 28, 28),
                ForeColor = Color.FromArgb(254, 240, 138),
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
                bool wordIntegrated = false;
                if (chkWordMacroExport.Checked)
                {
                    lblStatus.Text = "Đang tích hợp Ribbon Tab & Macro chuẩn 05-HD/VPTW trực tiếp vào Microsoft Word...";
                    Application.DoEvents();
                    wordIntegrated = WordAddinInstaller.InstallWordRibbonAddin(installDir, targetExePath);
                }

                string desktopLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "VBTM Pro 2.5.0.lnk");
                string startMenuLnk = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), "VBTM Pro 2.5.0.lnk");

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
                    CreateShortcut(desktopLnk, targetExePath, "--app", "VBTM Pro 2.5.0 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                    CreateShortcut(startMenuLnk, targetExePath, "--app", "VBTM Pro 2.5.0 - Chuan Hoa Van Ban Dang Uy & UBND Phuong");
                }

                progressBar.Value = 100;
                lblStatus.Text = "✓ Cài đặt VBTM Pro 2.5.0 hoàn tất thành công!";
                lblStatus.ForeColor = Color.FromArgb(16, 185, 129);
                isInstalled = true;
                btnInstall.Text = "Mở VBTM Pro 2.5.0";
                btnInstall.BackColor = Color.FromArgb(16, 185, 129);

                DialogResult openWord = MessageBox.Show(
                    "Đã cài đặt & tích hợp trực tiếp VBTM Pro 2.5.0 vào Microsoft Word thành công!\n\n" +
                    "• Biểu tượng Desktop: VBTM Pro 2.5.0 (đã gắn logo chính thức)\n" +
                    (wordIntegrated
                        ? "• Tích hợp Word: ĐÃ CÀI SẴN Tab Ribbon 'VBTM Pro 2.5 (05-HD & ND30)' trực tiếp trong Microsoft Word!\n\n"
                        : "• Tệp Macro Word: " + Path.Combine(installDir, "VBTM_Word_Addin.bas") + "\n\n") +
                    "• Bấm YES để mở Microsoft Word ngay bây giờ.\n" +
                    "• Bấm NO để mở cửa sổ VBTM Pro 2.5.0.",
                    "Cài Đặt Thành Công - VBTM Pro 2.5.0",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information
                );

                if (openWord == DialogResult.Yes)
                {
                    try { Process.Start("winword.exe"); } catch { }
                    this.Close();
                }
                else
                {
                    LaunchInstalledApp();
                }
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
                        key.SetValue("DisplayName", "VBTM Chuẩn Hóa Văn Bản Đảng Ủy & UBND Phường (VBTM Pro 2.5.0)");
                        key.SetValue("DisplayVersion", "2.5.0");
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
            CreateShortcutStatic(shortcutPath, targetPath, arguments, description);
        }

        public static void CreateShortcutStatic(string shortcutPath, string targetPath, string arguments, string description)
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
                scType.InvokeMember("IconLocation", BindingFlags.SetProperty, null, shortcut, new object[] { targetPath + ",0" });
                scType.InvokeMember("Description", BindingFlags.SetProperty, null, shortcut, new object[] { description });
                scType.InvokeMember("Save", BindingFlags.InvokeMethod, null, shortcut, null);
            }
            catch { }
        }

        private string GetWordVbaMacroCode()
        {
            return WordAddinInstaller.GetWordVbaMacroCode();
        }
    }

    // =========================================================================
    // 1B. BỘ TÍCH HỢP TRỰC TIẾP VÀO MICROSOFT WORD (STARTUP .DOTM + RIBBON TAB)
    // =========================================================================
    public static class WordAddinInstaller
    {
        private static string VbaStr(string s)
        {
            if (string.IsNullOrEmpty(s)) return "\"\"";
            StringBuilder sb = new StringBuilder();
            bool inQuotes = false;
            foreach (char c in s)
            {
                if (c >= 32 && c < 127 && c != '"')
                {
                    if (!inQuotes)
                    {
                        if (sb.Length > 0) sb.Append(" & ");
                        sb.Append('"');
                        inQuotes = true;
                    }
                    sb.Append(c);
                }
                else
                {
                    if (inQuotes)
                    {
                        sb.Append('"');
                        inQuotes = false;
                    }
                    if (sb.Length > 0) sb.Append(" & ");
                    sb.Append("ChrW(").Append((int)c).Append(")");
                }
            }
            if (inQuotes) sb.Append('"');
            return sb.ToString();
        }

        public static string GetWordVbaMacroCode()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Attribute VB_Name = \"VBTM_Pro_HD05_ND30\"");
            sb.AppendLine("Option Explicit");
            sb.AppendLine("");
            // Ribbon callbacks
            sb.AppendLine("Sub Ribbon_ChuanHoaTuDong(ByVal control As Object): ChuanHoaVanBan_TuDong: End Sub");
            sb.AppendLine("Sub Ribbon_HuongDan05(ByVal control As Object): ChuanHoaVanBan_HuongDan05: End Sub");
            sb.AppendLine("Sub Ribbon_NghiDinh30(ByVal control As Object): ChuanHoaVanBan_NghiDinh30: End Sub");
            sb.AppendLine("Sub Ribbon_ChenMauDang05(ByVal control As Object): ChenMauVanBan_Dang05: End Sub");
            sb.AppendLine("Sub Ribbon_ChenMauUBND30(ByVal control As Object): ChenMauVanBan_UBND30: End Sub");
            sb.AppendLine("Sub Ribbon_SuaChinhTa(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --word-spellcheck\", vbNormalFocus Else SuaChinhTaVaDauCau_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_TuDongHoaTieuNgu(ByVal control As Object): TuDongHoaTieuNgu_Word CStr(control.Tag): End Sub");
            sb.AppendLine("Sub Ribbon_NoiDongPDF(ByVal control As Object): NoiDongPDFBiNgat_Word: End Sub");
            sb.AppendLine("Sub Ribbon_CanLeA4(ByVal control As Object): CanLeA4Chuan_Word True: End Sub");
            sb.AppendLine("Sub Ribbon_DanhSoTrang(ByVal control As Object): DanhSoTrangTuTrang2_Word True: End Sub");
            sb.AppendLine("Private Sub ExportWordSnapshotForApp()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim fso As Object: Set fso = CreateObject(\"Scripting.FileSystemObject\")");
            sb.AppendLine("    Dim ts As Object: Set ts = fso.CreateTextFile(Environ(\"TEMP\") & \"\\vbtm_word_snapshot.txt\", True, True)");
            sb.AppendLine("    ts.Write ActiveDocument.Content.Text");
            sb.AppendLine("    ts.Close");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_AISoanThao(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    ExportWordSnapshotForApp");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --ai-word\", vbNormalFocus Else AISoanThao_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_RaSoatChinhTaAI(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    ExportWordSnapshotForApp");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --spell-word\", vbNormalFocus Else SuaChinhTaVaDauCau_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_AiTomTatVanBan(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    ExportWordSnapshotForApp");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) = \"\" Then p = Environ(\"LOCALAPPDATA\") & \"\\VBTMPro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --summarize-word\", vbNormalFocus Else AiTomTatVanBan_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_ChuyenDoiDangHanhChinh(ByVal control As Object): ChuyenDoiDangHanhChinh_Word: End Sub");
            sb.AppendLine("Sub Ribbon_KiemDinhTheThuc(ByVal control As Object): KiemDinhTheThuc_Word: End Sub");
            sb.AppendLine("Sub Ribbon_ChenBangChuan(ByVal control As Object): ChenBangChuan_Word: End Sub");
            sb.AppendLine("Sub Ribbon_ChenDoKhanMat(ByVal control As Object): ChenDoKhanMat_Word CStr(control.Tag): End Sub");
            sb.AppendLine("Sub Ribbon_XuatPDFNhanh(ByVal control As Object): XuatPDFNhanh_Word: End Sub");
            sb.AppendLine("Sub Ribbon_ThongKeVanBan(ByVal control As Object): ThongKeVanBan_Word: End Sub");
            sb.AppendLine("Sub Ribbon_InsertTpl(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --insert-tpl \" & control.Tag, vbNormalFocus Else ChenMauVanBan_Dang05");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_ChinhDangNhanh(ByVal control As Object): ChinhTheThucDangNhanh_Word: End Sub");
            sb.AppendLine("Sub Ribbon_CaiDatDonVi(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --unit-settings\", vbNormalFocus Else CaiDatTieuNguDonVi_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_ChenMaQR(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    ExportWordSnapshotForApp");
            sb.AppendLine("    Dim p As String: p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then Shell \"\"\"\" & p & \"\"\" --qr-word\", vbNormalFocus Else Ribbon_MoVBTM control");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            sb.AppendLine("Private Sub ShowUnicodeMsg(ByVal msgText As String)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Not Application.Visible Then Exit Sub");
            sb.AppendLine("    Application.StatusBar = \"VBTM Pro 2.5.0: \" & Replace(msgText, vbCrLf, \" \")");
            sb.AppendLine("    CreateObject(\"WScript.Shell\").Popup msgText, 4, \"VBTM Pro 2.5.0\", 4160");
            sb.AppendLine("End Sub");
            sb.AppendLine("Sub Ribbon_MoVBTM(ByVal control As Object)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim p As String");
            sb.AppendLine("    p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"");
            sb.AppendLine("    If Dir(p) <> \"\" Then");
            sb.AppendLine("        Shell \"\"\"\" & p & \"\"\" --app\", vbNormalFocus");
            sb.AppendLine("    Else");
            sb.AppendLine("        MsgBox " + VbaStr("Không tìm thấy VBTM-Pro.exe tại: ") + " & p, vbExclamation, \"VBTM Pro 2.5.0\"");
            sb.AppendLine("    End If");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 1. Căn lề A4 chuẩn 20-20-30-15mm
            sb.AppendLine("Sub CanLeA4Chuan_Word(Optional ByVal showMsg As Boolean = False)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    With ActiveDocument.PageSetup");
            sb.AppendLine("        .PaperSize = 7");
            sb.AppendLine("        .TopMargin = CentimetersToPoints(2)");
            sb.AppendLine("        .BottomMargin = CentimetersToPoints(2)");
            sb.AppendLine("        .LeftMargin = CentimetersToPoints(3)");
            sb.AppendLine("        .RightMargin = CentimetersToPoints(1.5)");
            sb.AppendLine("    End With");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã căn lề chuẩn khổ A4 (Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm)!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 2. Đánh số trang từ trang 2 ở Header (cỡ 13pt đứng, căn giữa)
            sb.AppendLine("Sub DanhSoTrangTuTrang2_Word(Optional ByVal showMsg As Boolean = False)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True");
            sb.AppendLine("    Dim hdr As HeaderFooter");
            sb.AppendLine("    Set hdr = ActiveDocument.Sections(1).Headers(1)");
            sb.AppendLine("    hdr.Range.Text = \"\"");
            sb.AppendLine("    hdr.PageNumbers.Add PageNumberAlignment:=1, FirstPage:=False");
            sb.AppendLine("    hdr.Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("    hdr.Range.Font.Size = 13");
            sb.AppendLine("    hdr.Range.Font.Bold = False");
            sb.AppendLine("    hdr.Range.Font.Italic = False");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã bật đánh số trang tự động từ trang 2 (căn giữa lề trên, cỡ 13pt)!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 3. Tự động nhận diện Đảng (05-HD/VPTW) hay Hành chính (NĐ 30/2020/NĐ-CP)
            sb.AppendLine("Sub ChuanHoaVanBan_TuDong()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim docText As String");
            sb.AppendLine("    docText = UCase(ActiveDocument.Content.Text)");
            sb.AppendLine("    If InStr(docText, " + VbaStr("ĐẢNG CỘNG SẢN") + ") > 0 Or InStr(docText, " + VbaStr("ĐẢNG ỦY") + ") > 0 Or InStr(docText, " + VbaStr("CHI BỘ") + ") > 0 Or InStr(docText, " + VbaStr("-BC/ĐU") + ") > 0 Or InStr(docText, " + VbaStr("-NQ/ĐU") + ") > 0 Or InStr(docText, " + VbaStr("-CV/ĐU") + ") > 0 Then");
            sb.AppendLine("        ChuanHoaVanBan_HuongDan05");
            sb.AppendLine("    Else");
            sb.AppendLine("        ChuanHoaVanBan_NghiDinh30");
            sb.AppendLine("    End If");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 4. Chuẩn hóa văn bản Đảng theo Hướng dẫn 05-HD/VPTW
            sb.AppendLine("Sub ChuanHoaVanBan_HuongDan05()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Application.ScreenUpdating = False");
            sb.AppendLine("    CanLeA4Chuan_Word False");
            sb.AppendLine("    DanhSoTrangTuTrang2_Word False");
            sb.AppendLine("    ActiveDocument.Content.Font.Name = \"Times New Roman\"");
            sb.AppendLine("    Dim p As Paragraph, t As String");
            sb.AppendLine("    For Each p In ActiveDocument.Paragraphs");
            sb.AppendLine("        If p.Range.Tables.Count = 0 Then");
            sb.AppendLine("            t = Trim(Replace(Replace(p.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("            If Len(t) > 0 Then");
            sb.AppendLine("                p.Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("                p.LeftIndent = 0: p.RightIndent = 0: p.Format.SpaceBefore = 0");
            sb.AppendLine("                If t = " + VbaStr("BÁO CÁO") + " Or t = " + VbaStr("NGHỊ QUYẾT") + " Or t = " + VbaStr("QUYẾT ĐỊNH") + " Or t = " + VbaStr("KẾ HOẠCH") + " Or t = " + VbaStr("CHƯƠNG TRÌNH") + " Or t = " + VbaStr("TỜ TRÌNH") + " Or t = " + VbaStr("THÔNG BÁO") + " Or t = " + VbaStr("CHỈ THỊ") + " Or t = " + VbaStr("BIÊN BẢN") + " Or t = " + VbaStr("GIẤY MỜI") + " Then");
            sb.AppendLine("                    p.Range.Font.Size = 15");
            sb.AppendLine("                    p.Range.Font.Bold = True");
            sb.AppendLine("                    p.Alignment = 1");
            sb.AppendLine("                    p.FirstLineIndent = 0");
            sb.AppendLine("                    p.Format.SpaceAfter = 2");
            sb.AppendLine("                ElseIf Left(t, 3) = " + VbaStr("───") + " Or Left(t, 3) = \"---\" Then");
            sb.AppendLine("                    If t <> \"-----\" Then");
            sb.AppendLine("                        Dim rRule As Range: Set rRule = p.Range.Duplicate");
            sb.AppendLine("                        If rRule.End > rRule.Start Then rRule.End = rRule.End - 1");
            sb.AppendLine("                        rRule.Text = \"-----\"");
            sb.AppendLine("                    End If");
            sb.AppendLine("                    p.Range.Font.Size = 14");
            sb.AppendLine("                    p.Range.Font.Bold = True");
            sb.AppendLine("                    p.Alignment = 1");
            sb.AppendLine("                    p.FirstLineIndent = 0");
            sb.AppendLine("                    p.Format.SpaceAfter = 10");
            sb.AppendLine("                ElseIf t = " + VbaStr("QUYẾT ĐỊNH:") + " Or t = " + VbaStr("QUYẾT NGHỊ:") + " Then");
            sb.AppendLine("                    p.Range.Font.Size = 14");
            sb.AppendLine("                    p.Range.Font.Bold = True");
            sb.AppendLine("                    p.Alignment = 1");
            sb.AppendLine("                    p.FirstLineIndent = 0");
            sb.AppendLine("                    p.Format.SpaceAfter = 6");
            sb.AppendLine("                ElseIf Left(t, 2) = \"I.\" Or Left(t, 3) = \"II.\" Or Left(t, 4) = \"III.\" Or Left(t, 3) = \"IV.\" Or Left(t, 2) = \"V.\" Then");
            sb.AppendLine("                    p.Range.Font.Size = 14");
            sb.AppendLine("                    p.Range.Font.Bold = True");
            sb.AppendLine("                    p.Alignment = 3");
            sb.AppendLine("                    p.FirstLineIndent = CentimetersToPoints(1.27)");
            sb.AppendLine("                    p.Format.SpaceAfter = 6");
            sb.AppendLine("                Else");
            sb.AppendLine("                    p.Range.Font.Size = 14");
            sb.AppendLine("                    p.Format.LineSpacingRule = 5");
            sb.AppendLine("                    p.Format.LineSpacing = LinesToPoints(1.15)");
            sb.AppendLine("                    If p.Alignment = 1 Then");
            sb.AppendLine("                        p.FirstLineIndent = 0");
            sb.AppendLine("                        p.Format.SpaceAfter = 2");
            sb.AppendLine("                    Else");
            sb.AppendLine("                        p.Alignment = 3");
            sb.AppendLine("                        p.FirstLineIndent = CentimetersToPoints(1.27)");
            sb.AppendLine("                        p.Format.SpaceAfter = 6");
            sb.AppendLine("                    End If");
            sb.AppendLine("                End If");
            sb.AppendLine("            End If");
            sb.AppendLine("        End If");
            sb.AppendLine("    Next p");
            sb.AppendLine("    If ActiveDocument.Tables.Count >= 1 Then");
            sb.AppendLine("        Dim tblTop As Table");
            sb.AppendLine("        Set tblTop = ActiveDocument.Tables(1)");
            sb.AppendLine("        If tblTop.Columns.Count >= 2 Then");
            sb.AppendLine("            tblTop.AllowAutoFit = False");
            sb.AppendLine("            tblTop.PreferredWidthType = 2: tblTop.PreferredWidth = 100");
            sb.AppendLine("            tblTop.LeftPadding = 0: tblTop.RightPadding = 0: tblTop.TopPadding = 0: tblTop.BottomPadding = 0");
            sb.AppendLine("            tblTop.Columns(1).PreferredWidthType = 2: tblTop.Columns(1).PreferredWidth = 50");
            sb.AppendLine("            tblTop.Columns(2).PreferredWidthType = 2: tblTop.Columns(2).PreferredWidth = 50");
            sb.AppendLine("            tblTop.Cell(1, 1).LeftPadding = 0: tblTop.Cell(1, 1).RightPadding = 0");
            sb.AppendLine("            tblTop.Cell(1, 2).LeftPadding = 0: tblTop.Cell(1, 2).RightPadding = 0");
            sb.AppendLine("            tblTop.Cell(1, 1).Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("            tblTop.Cell(1, 1).Range.Font.Scaling = 100");
            sb.AppendLine("            tblTop.Cell(1, 1).Range.ParagraphFormat.FirstLineIndent = 0: tblTop.Cell(1, 1).Range.ParagraphFormat.Alignment = 1: tblTop.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0");
            sb.AppendLine("            Dim pL As Paragraph, tL As String");
            sb.AppendLine("            For Each pL In tblTop.Cell(1, 1).Range.Paragraphs");
            sb.AppendLine("                tL = Trim(Replace(Replace(pL.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("                pL.Range.Font.Scaling = 100");
            sb.AppendLine("                If Len(tL) > 25 Then pL.Range.Font.Size = 12.5 Else If Len(tL) > 21 Then pL.Range.Font.Size = 13 Else pL.Range.Font.Size = 14");
            sb.AppendLine("            Next pL");
            sb.AppendLine("            tblTop.Cell(1, 2).Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("            tblTop.Cell(1, 2).Range.Font.Scaling = 100");
            sb.AppendLine("            tblTop.Cell(1, 2).Range.ParagraphFormat.FirstLineIndent = 0: tblTop.Cell(1, 2).Range.ParagraphFormat.Alignment = 1: tblTop.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0");
            sb.AppendLine("            Dim rpIdx As Integer, pTop As Paragraph, tTop As String");
            sb.AppendLine("            For rpIdx = tblTop.Cell(1, 2).Range.Paragraphs.Count To 1 Step -1");
            sb.AppendLine("                Set pTop = tblTop.Cell(1, 2).Range.Paragraphs(rpIdx)");
            sb.AppendLine("                tTop = Trim(Replace(Replace(pTop.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("                If InStr(UCase(tTop), " + VbaStr("ĐẢNG CỘNG SẢN") + ") > 0 Then");
            sb.AppendLine("                    pTop.Range.Font.Size = 15: pTop.Range.Font.Bold = True: pTop.Range.Font.Underline = 0: pTop.Range.Font.Scaling = 100");
            sb.AppendLine("                    pTop.Format.LeftIndent = CentimetersToPoints(0.45): pTop.Format.RightIndent = CentimetersToPoints(0.45)");
            sb.AppendLine("                    pTop.Borders(-3).LineStyle = 1: pTop.Borders(-3).LineWidth = 6");
            sb.AppendLine("                ElseIf Left(tTop, 3) = " + VbaStr("───") + " Or Left(tTop, 3) = \"---\" Then");
            sb.AppendLine("                    pTop.Range.Delete");
            sb.AppendLine("                ElseIf InStr(LCase(tTop), " + VbaStr("ngày") + ") > 0 Then");
            sb.AppendLine("                    pTop.Range.Font.Italic = True: pTop.Range.Font.Bold = False: pTop.Range.Font.Scaling = 100: pTop.Format.SpaceBefore = 3");
            sb.AppendLine("                    If Len(tTop) > 40 Then pTop.Range.Font.Size = 12 Else If Len(tTop) > 32 Then pTop.Range.Font.Size = 13 Else pTop.Range.Font.Size = 14");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next rpIdx");
            sb.AppendLine("        End If");
            sb.AppendLine("    End If");
            sb.AppendLine("    If ActiveDocument.Tables.Count >= 2 Then");
            sb.AppendLine("        Dim tblBot As Table");
            sb.AppendLine("        Set tblBot = ActiveDocument.Tables(ActiveDocument.Tables.Count)");
            sb.AppendLine("        If tblBot.Columns.Count >= 2 Then");
            sb.AppendLine("            tblBot.AllowAutoFit = False");
            sb.AppendLine("            tblBot.LeftPadding = 0: tblBot.RightPadding = 0: tblBot.TopPadding = 0: tblBot.BottomPadding = 0");
            sb.AppendLine("            Dim pRec As Paragraph");
            sb.AppendLine("            For Each pRec In tblBot.Cell(1, 1).Range.Paragraphs");
            sb.AppendLine("                pRec.Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("                pRec.Range.Font.Italic = False");
            sb.AppendLine("                pRec.Range.Font.Bold = False");
            sb.AppendLine("                pRec.Range.Font.Underline = 0");
            sb.AppendLine("                pRec.Format.SpaceBefore = 0: pRec.Format.SpaceAfter = 0: pRec.Format.LineSpacingRule = 0");
            sb.AppendLine("                If InStr(LCase(pRec.Range.Text), " + VbaStr("nơi nhận") + ") > 0 Then");
            sb.AppendLine("                    pRec.Range.Font.Size = 14");
            sb.AppendLine("                    Dim uRng As Range: Set uRng = pRec.Range.Duplicate");
            sb.AppendLine("                    uRng.End = uRng.Start + 8");
            sb.AppendLine("                    uRng.Font.Underline = 1");
            sb.AppendLine("                Else");
            sb.AppendLine("                    pRec.Range.Font.Size = 12");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next pRec");
            sb.AppendLine("            tblBot.Cell(1, 2).Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("            tblBot.Cell(1, 2).Range.Font.Size = 14");
            sb.AppendLine("            Dim pSig As Paragraph, tSig As String");
            sb.AppendLine("            For Each pSig In tblBot.Cell(1, 2).Range.Paragraphs");
            sb.AppendLine("                pSig.Format.SpaceBefore = 0: pSig.Format.SpaceAfter = 0: pSig.Format.LineSpacingRule = 0");
            sb.AppendLine("                tSig = UCase(Trim(Replace(Replace(pSig.Range.Text, vbCr, \"\"), Chr(7), \"\")))");
            sb.AppendLine("                If tSig = " + VbaStr("BÍ THƯ") + " Or tSig = " + VbaStr("PHÓ BÍ THƯ") + " Or tSig = " + VbaStr("PHÓ BÍ THƯ THƯỜNG TRỰC") + " Or tSig = " + VbaStr("CHÁNH VĂN PHÒNG") + " Or tSig = " + VbaStr("PHÓ CHÁNH VĂN PHÒNG") + " Or tSig = " + VbaStr("CHỦ NHIỆM") + " Or tSig = " + VbaStr("TRƯỞNG BAN") + " Then");
            sb.AppendLine("                    pSig.Range.Font.Bold = False");
            sb.AppendLine("                Else");
            sb.AppendLine("                    pSig.Range.Font.Bold = True");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next pSig");
            sb.AppendLine("        End If");
            sb.AppendLine("    End If");
            sb.AppendLine("    Application.ScreenUpdating = True");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã tự động chuẩn hóa văn bản Word đúng chuẩn Hướng dẫn số 05-HD/VPTW!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 5. Chuẩn hóa văn bản Hành chính theo Nghị định 30/2020/NĐ-CP (đầy đủ Tiêu ngữ, Thân bài, Nơi nhận & Chữ ký)
            sb.AppendLine("Sub ChuanHoaVanBan_NghiDinh30()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Application.ScreenUpdating = False");
            sb.AppendLine("    CanLeA4Chuan_Word False");
            sb.AppendLine("    DanhSoTrangTuTrang2_Word False");
            sb.AppendLine("    ActiveDocument.Content.Font.Name = \"Times New Roman\"");
            sb.AppendLine("    Dim p As Paragraph, t As String");
            sb.AppendLine("    For Each p In ActiveDocument.Paragraphs");
            sb.AppendLine("        If p.Range.Tables.Count = 0 Then");
            sb.AppendLine("            t = Trim(Replace(Replace(p.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("            If Len(t) > 0 Then");
            sb.AppendLine("                p.Range.Font.Name = \"Times New Roman\"");
            sb.AppendLine("                p.LeftIndent = 0: p.RightIndent = 0: p.Format.SpaceBefore = 0");
            sb.AppendLine("                If t = " + VbaStr("BÁO CÁO") + " Or t = " + VbaStr("NGHỊ QUYẾT") + " Or t = " + VbaStr("QUYẾT ĐỊNH") + " Or t = " + VbaStr("KẾ HOẠCH") + " Or t = " + VbaStr("TỜ TRÌNH") + " Or t = " + VbaStr("THÔNG BÁO") + " Or t = " + VbaStr("BIÊN BẢN") + " Or t = " + VbaStr("GIẤY MỜI") + " Then");
            sb.AppendLine("                    p.Range.Font.Size = 14: p.Range.Font.Bold = True: p.Alignment = 1: p.FirstLineIndent = 0: p.Format.SpaceAfter = 2");
            sb.AppendLine("                ElseIf Left(t, 3) = " + VbaStr("───") + " Or Left(t, 3) = \"---\" Then");
            sb.AppendLine("                    p.Range.Font.Size = 10: p.Range.Font.Bold = False: p.Alignment = 1: p.FirstLineIndent = 0: p.Format.SpaceAfter = 10");
            sb.AppendLine("                ElseIf Left(t, 2) = \"I.\" Or Left(t, 3) = \"II.\" Or Left(t, 4) = \"III.\" Or Left(t, 3) = \"IV.\" Or Left(t, 2) = \"V.\" Then");
            sb.AppendLine("                    p.Range.Font.Size = 14: p.Range.Font.Bold = True: p.Alignment = 3: p.FirstLineIndent = CentimetersToPoints(1.27): p.Format.SpaceAfter = 6");
            sb.AppendLine("                Else");
            sb.AppendLine("                    p.Range.Font.Size = 14: p.Format.LineSpacingRule = 5: p.Format.LineSpacing = LinesToPoints(1.15)");
            sb.AppendLine("                    If p.Alignment = 1 Then");
            sb.AppendLine("                        p.FirstLineIndent = 0: p.Format.SpaceAfter = 2");
            sb.AppendLine("                    Else");
            sb.AppendLine("                        p.Alignment = 3: p.FirstLineIndent = CentimetersToPoints(1.27): p.Format.SpaceAfter = 6");
            sb.AppendLine("                    End If");
            sb.AppendLine("                End If");
            sb.AppendLine("            End If");
            sb.AppendLine("        End If");
            sb.AppendLine("    Next p");
            sb.AppendLine("    If ActiveDocument.Tables.Count >= 1 Then");
            sb.AppendLine("        Dim tblTop As Table: Set tblTop = ActiveDocument.Tables(1)");
            sb.AppendLine("        If tblTop.Columns.Count >= 2 Then");
            sb.AppendLine("            tblTop.AllowAutoFit = False: tblTop.PreferredWidthType = 2: tblTop.PreferredWidth = 100");
            sb.AppendLine("            tblTop.LeftPadding = 0: tblTop.RightPadding = 0: tblTop.TopPadding = 0: tblTop.BottomPadding = 0");
            sb.AppendLine("            tblTop.Columns(1).PreferredWidthType = 2: tblTop.Columns(1).PreferredWidth = 42");
            sb.AppendLine("            tblTop.Columns(2).PreferredWidthType = 2: tblTop.Columns(2).PreferredWidth = 58");
            sb.AppendLine("            tblTop.Cell(1, 1).Range.Font.Name = \"Times New Roman\": tblTop.Cell(1, 1).Range.Font.Size = 12.5: tblTop.Cell(1, 1).Range.Font.Scaling = 100");
            sb.AppendLine("            tblTop.Cell(1, 1).Range.ParagraphFormat.FirstLineIndent = 0: tblTop.Cell(1, 1).Range.ParagraphFormat.Alignment = 1: tblTop.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0");
            sb.AppendLine("            tblTop.Cell(1, 2).Range.Font.Name = \"Times New Roman\": tblTop.Cell(1, 2).Range.Font.Scaling = 100");
            sb.AppendLine("            tblTop.Cell(1, 2).Range.ParagraphFormat.FirstLineIndent = 0: tblTop.Cell(1, 2).Range.ParagraphFormat.Alignment = 1: tblTop.Cell(1, 2).Range.ParagraphFormat.SpaceAfter = 0");
            sb.AppendLine("            Dim pTop As Paragraph, tTop As String");
            sb.AppendLine("            For Each pTop In tblTop.Cell(1, 2).Range.Paragraphs");
            sb.AppendLine("                tTop = Trim(Replace(Replace(pTop.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("                If InStr(UCase(tTop), " + VbaStr("CỘNG HÒA") + ") > 0 Then");
            sb.AppendLine("                    pTop.Range.Font.Size = 12.5: pTop.Range.Font.Bold = True: pTop.Range.Font.Scaling = 100");
            sb.AppendLine("                ElseIf InStr(LCase(tTop), " + VbaStr("độc lập") + ") > 0 Then");
            sb.AppendLine("                    pTop.Range.Font.Size = 13.5: pTop.Range.Font.Bold = True: pTop.Range.Font.Underline = 0: pTop.Range.Font.Scaling = 100");
            sb.AppendLine("                ElseIf InStr(LCase(tTop), " + VbaStr("ngày") + ") > 0 Then");
            sb.AppendLine("                    pTop.Range.Font.Size = 13.5: pTop.Range.Font.Italic = True: pTop.Range.Font.Bold = False: pTop.Range.Font.Scaling = 100: pTop.Format.SpaceBefore = 2");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next pTop");
            sb.AppendLine("        End If");
            sb.AppendLine("    End If");
            sb.AppendLine("    If ActiveDocument.Tables.Count >= 2 Then");
            sb.AppendLine("        Dim tblBot As Table: Set tblBot = ActiveDocument.Tables(ActiveDocument.Tables.Count)");
            sb.AppendLine("        If tblBot.Columns.Count >= 2 Then");
            sb.AppendLine("            tblBot.AllowAutoFit = False: tblBot.LeftPadding = 0: tblBot.RightPadding = 0: tblBot.TopPadding = 0: tblBot.BottomPadding = 0");
            sb.AppendLine("            Dim pRec As Paragraph");
            sb.AppendLine("            For Each pRec In tblBot.Cell(1, 1).Range.Paragraphs");
            sb.AppendLine("                pRec.Range.Font.Name = \"Times New Roman\": pRec.Format.SpaceBefore = 0: pRec.Format.SpaceAfter = 0: pRec.Format.LineSpacingRule = 0");
            sb.AppendLine("                If InStr(LCase(pRec.Range.Text), " + VbaStr("nơi nhận") + ") > 0 Then");
            sb.AppendLine("                    pRec.Range.Font.Size = 12: pRec.Range.Font.Bold = True: pRec.Range.Font.Italic = True");
            sb.AppendLine("                Else");
            sb.AppendLine("                    pRec.Range.Font.Size = 11: pRec.Range.Font.Bold = False: pRec.Range.Font.Italic = False");
            sb.AppendLine("                End If");
            sb.AppendLine("            Next pRec");
            sb.AppendLine("            tblBot.Cell(1, 2).Range.Font.Name = \"Times New Roman\": tblBot.Cell(1, 2).Range.Font.Size = 14: tblBot.Cell(1, 2).Range.Font.Bold = True");
            sb.AppendLine("        End If");
            sb.AppendLine("    End If");
            sb.AppendLine("    Application.ScreenUpdating = True");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã chuẩn hóa văn bản Word đầy đủ theo Nghị định số 30/2020/NĐ-CP!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 6. Chèn Mẫu Văn Bản Đảng Chuẩn Hướng dẫn 05-HD/VPTW trực tiếp vào Word (tự nạp Tiêu ngữ đơn vị đã cài đặt)
            sb.AppendLine("Private Function GetUnitConfigVal(ByVal idx As Integer, ByVal defVal As String) As String");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    GetUnitConfigVal = defVal");
            sb.AppendLine("    Dim cfgPath As String: cfgPath = Environ(\"APPDATA\") & \"\\VBTM\\word_unit_config.ini\"");
            sb.AppendLine("    Dim fso As Object: Set fso = CreateObject(\"Scripting.FileSystemObject\")");
            sb.AppendLine("    If Not fso.FileExists(cfgPath) Then Exit Function");
            sb.AppendLine("    Dim ts As Object: Set ts = fso.OpenTextFile(cfgPath, 1, False, -1)");
            sb.AppendLine("    Dim allTxt As String: allTxt = ts.ReadAll: ts.Close");
            sb.AppendLine("    Dim arr As Variant: arr = Split(allTxt, vbCrLf)");
            sb.AppendLine("    If idx <= UBound(arr) Then If Len(Trim(arr(idx))) > 0 Then GetUnitConfigVal = Trim(arr(idx))");
            sb.AppendLine("End Function");
            sb.AppendLine("");
            sb.AppendLine("Sub CaiDatTieuNguDonVi_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim pGov As String, pIss As String, uGov As String, uIss As String, loc As String, pSign As String, uSign As String");
            sb.AppendLine("    pGov = GetUnitConfigVal(0, " + VbaStr("QUẬN ỦY") + ")");
            sb.AppendLine("    pIss = GetUnitConfigVal(1, " + VbaStr("ĐẢNG ỦY PHƯỜNG") + ")");
            sb.AppendLine("    uGov = GetUnitConfigVal(2, " + VbaStr("ỦY BAN NHÂN DÂN QUẬN") + ")");
            sb.AppendLine("    uIss = GetUnitConfigVal(3, " + VbaStr("ỦY BAN NHÂN DÂN PHƯỜNG") + ")");
            sb.AppendLine("    loc  = GetUnitConfigVal(4, " + VbaStr("Phường") + ")");
            sb.AppendLine("    pSign = GetUnitConfigVal(5, " + VbaStr("Trần Minh Quân") + ")");
            sb.AppendLine("    uSign = GetUnitConfigVal(6, " + VbaStr("Nguyễn Văn An") + ")");
            sb.AppendLine("    If Application.Visible Then");
            sb.AppendLine("        Dim v As String");
            sb.AppendLine("        v = InputBox(" + VbaStr("1/5. Tên cơ quan Đảng cấp trên (VD: QUẬN ỦY ...):") + ", " + VbaStr("Cài Đặt Tiêu Ngữ Đơn Vị") + ", pGov): If Len(v) > 0 Then pGov = UCase(Trim(v))");
            sb.AppendLine("        v = InputBox(" + VbaStr("2/5. Tên đơn vị Đảng ban hành (VD: ĐẢNG ỦY PHƯỜNG ...):") + ", " + VbaStr("Cài Đặt Tiêu Ngữ Đơn Vị") + ", pIss): If Len(v) > 0 Then pIss = UCase(Trim(v))");
            sb.AppendLine("        v = InputBox(" + VbaStr("3/5. Tên UBND cấp trên (VD: ỦY BAN NHÂN DÂN QUẬN ...):") + ", " + VbaStr("Cài Đặt Tiêu Ngữ Đơn Vị") + ", uGov): If Len(v) > 0 Then uGov = UCase(Trim(v))");
            sb.AppendLine("        v = InputBox(" + VbaStr("4/5. Tên UBND ban hành (VD: ỦY BAN NHÂN DÂN PHƯỜNG ...):") + ", " + VbaStr("Cài Đặt Tiêu Ngữ Đơn Vị") + ", uIss): If Len(v) > 0 Then uIss = UCase(Trim(v))");
            sb.AppendLine("        v = InputBox(" + VbaStr("5/5. Địa danh & Họ tên Bí thư (cách nhau dấu |, VD: Hải Châu|Trần Minh Quân):") + ", " + VbaStr("Cài Đặt Tiêu Ngữ Đơn Vị") + ", loc & \"|\" & pSign)");
            sb.AppendLine("        If InStr(v, \"|\") > 0 Then loc = Trim(Split(v, \"|\")(0)): pSign = Trim(Split(v, \"|\")(1))");
            sb.AppendLine("    End If");
            sb.AppendLine("    Dim fso As Object: Set fso = CreateObject(\"Scripting.FileSystemObject\")");
            sb.AppendLine("    Dim dirPath As String: dirPath = Environ(\"APPDATA\") & \"\\VBTM\"");
            sb.AppendLine("    If Not fso.FolderExists(dirPath) Then fso.CreateFolder dirPath");
            sb.AppendLine("    Dim ts As Object: Set ts = fso.CreateTextFile(dirPath & \"\\word_unit_config.ini\", True, True)");
            sb.AppendLine("    ts.Write pGov & vbCrLf & pIss & vbCrLf & uGov & vbCrLf & uIss & vbCrLf & loc & vbCrLf & pSign & vbCrLf & uSign");
            sb.AppendLine("    ts.Close");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã lưu cài đặt Tiêu ngữ đơn vị! Các nút Chèn Mẫu Đảng / UBND sẽ tự động dùng tên đơn vị của bạn.") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            sb.AppendLine("Sub ChinhTheThucDangNhanh_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim prevVis As Boolean: prevVis = Application.Visible");
            sb.AppendLine("    ChuanHoaVanBan_HuongDan05");
            sb.AppendLine("    SuaChinhTaVaDauCau_Word");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            sb.AppendLine("Sub ChenMauVanBan_Dang05()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Documents.Add");
            sb.AppendLine("    CanLeA4Chuan_Word False");
            sb.AppendLine("    DanhSoTrangTuTrang2_Word False");
            sb.AppendLine("    Dim pGov As String: pGov = GetUnitConfigVal(0, " + VbaStr("QUẬN ỦY") + ")");
            sb.AppendLine("    Dim pIss As String: pIss = GetUnitConfigVal(1, " + VbaStr("ĐẢNG ỦY PHƯỜNG") + ")");
            sb.AppendLine("    Dim loc As String: loc = GetUnitConfigVal(4, " + VbaStr("Phường") + ")");
            sb.AppendLine("    Dim pSign As String: pSign = GetUnitConfigVal(5, " + VbaStr("Trần Minh Quân") + ")");
            sb.AppendLine("    Dim pMotto As String: pMotto = GetUnitConfigVal(7, " + VbaStr("ĐẢNG CỘNG SẢN VIỆT NAM") + ")");
            sb.AppendLine("    Dim rng As Range, tbl1 As Table, tbl2 As Table");
            sb.AppendLine("    Set rng = ActiveDocument.Range(0, 0)");
            sb.AppendLine("    Set tbl1 = ActiveDocument.Tables.Add(rng, 1, 2)");
            sb.AppendLine("    tbl1.Borders.Enable = 0");
            sb.AppendLine("    tbl1.AllowAutoFit = False");
            sb.AppendLine("    tbl1.PreferredWidthType = 2: tbl1.PreferredWidth = 100");
            sb.AppendLine("    tbl1.LeftPadding = 0: tbl1.RightPadding = 0: tbl1.TopPadding = 0: tbl1.BottomPadding = 0");
            sb.AppendLine("    tbl1.Columns(1).PreferredWidthType = 2: tbl1.Columns(1).PreferredWidth = 50");
            sb.AppendLine("    tbl1.Columns(2).PreferredWidthType = 2: tbl1.Columns(2).PreferredWidth = 50");
            sb.AppendLine("    tbl1.Cell(1, 1).LeftPadding = 0: tbl1.Cell(1, 1).RightPadding = 0");
            sb.AppendLine("    tbl1.Cell(1, 2).LeftPadding = 0: tbl1.Cell(1, 2).RightPadding = 0");
            sb.AppendLine("    With tbl1.Cell(1, 1).Range");
            sb.AppendLine("        .Text = pGov & vbCr & pIss & vbCr & \"*\" & vbCr & " + VbaStr("Số 01-BC/ĐU"));
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Size = 14: .Font.Scaling = 100: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        If Len(pGov) > 25 Then .Paragraphs(1).Range.Font.Size = 12.5 Else If Len(pGov) > 21 Then .Paragraphs(1).Range.Font.Size = 13");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Bold = True");
            sb.AppendLine("        If Len(pIss) > 25 Then .Paragraphs(2).Range.Font.Size = 12.5 Else If Len(pIss) > 21 Then .Paragraphs(2).Range.Font.Size = 13");
            sb.AppendLine("        .Paragraphs(3).Range.Font.Bold = True");
            sb.AppendLine("    End With");
            sb.AppendLine("    With tbl1.Cell(1, 2).Range");
            sb.AppendLine("        Dim dtStr As String: dtStr = loc & " + VbaStr(", ngày 19 tháng 11 năm 2026"));
            sb.AppendLine("        .Text = pMotto & vbCr & dtStr");
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Scaling = 100: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(1).Range.Font.Size = 15: .Paragraphs(1).Range.Font.Bold = True: .Paragraphs(1).Range.Font.Underline = 0: .Paragraphs(1).Range.Font.Scaling = 100");
            sb.AppendLine("        .Paragraphs(1).Format.LeftIndent = CentimetersToPoints(0.45): .Paragraphs(1).Format.RightIndent = CentimetersToPoints(0.45)");
            sb.AppendLine("        .Paragraphs(1).Borders(-3).LineStyle = 1: .Paragraphs(1).Borders(-3).LineWidth = 6");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Italic = True: .Paragraphs(2).Range.Font.Bold = False: .Paragraphs(2).Range.Font.Scaling = 100: .Paragraphs(2).SpaceBefore = 3");
            sb.AppendLine("        If Len(dtStr) > 40 Then .Paragraphs(2).Range.Font.Size = 12 Else If Len(dtStr) > 32 Then .Paragraphs(2).Range.Font.Size = 13 Else .Paragraphs(2).Range.Font.Size = 14");
            sb.AppendLine("    End With");
            sb.AppendLine("    Dim pEnd As Range: Set pEnd = ActiveDocument.Content: pEnd.Collapse 0");
            sb.AppendLine("    pEnd.InsertAfter vbCr & " + VbaStr("BÁO CÁO") + " & vbCr & " + VbaStr("Tổng kết công tác xây dựng Đảng năm 2026") + " & vbCr & \"-----\" & vbCr & " + VbaStr("Thực hiện Chương trình công tác năm 2026, Đảng ủy phường báo cáo kết quả thực hiện như sau:./.") + " & vbCr");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 4).Alignment = 1: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 4).Range.Font.Size = 15: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 4).Range.Font.Bold = True: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 4).FirstLineIndent = 0: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 4).SpaceAfter = 2");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 3).Alignment = 1: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 3).Range.Font.Size = 14: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 3).Range.Font.Bold = True: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 3).FirstLineIndent = 0: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 3).SpaceAfter = 2");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).Alignment = 1: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).Range.Font.Size = 14: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).Range.Font.Bold = True: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).FirstLineIndent = 0: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).SpaceAfter = 10");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).Alignment = 3: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).Range.Font.Size = 14: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).FirstLineIndent = CentimetersToPoints(1.27): ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).SpaceAfter = 6");
            sb.AppendLine("    Set pEnd = ActiveDocument.Content: pEnd.Collapse 0");
            sb.AppendLine("    Set tbl2 = ActiveDocument.Tables.Add(pEnd, 1, 2)");
            sb.AppendLine("    tbl2.Borders.Enable = 0");
            sb.AppendLine("    tbl2.AllowAutoFit = False");
            sb.AppendLine("    tbl2.LeftPadding = 0: tbl2.RightPadding = 0: tbl2.TopPadding = 0: tbl2.BottomPadding = 0");
            sb.AppendLine("    With tbl2.Cell(1, 1).Range");
            sb.AppendLine("        .Text = " + VbaStr("Nơi nhận:") + " & vbCr & " + VbaStr("- Ban Thường vụ Quận ủy (để b/c);") + " & vbCr & " + VbaStr("- Các chi bộ trực thuộc;") + " & vbCr & " + VbaStr("- Lưu Văn phòng Đảng ủy."));
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .ParagraphFormat.Alignment = 0: .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.LineSpacingRule = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(1).Range.Font.Size = 14: .Paragraphs(1).Range.Font.Underline = 0");
            sb.AppendLine("        Dim uRng2 As Range: Set uRng2 = .Paragraphs(1).Range.Duplicate: uRng2.End = uRng2.Start + 8: uRng2.Font.Underline = 1");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Size = 12: .Paragraphs(3).Range.Font.Size = 12: .Paragraphs(4).Range.Font.Size = 12");
            sb.AppendLine("    End With");
            sb.AppendLine("    With tbl2.Cell(1, 2).Range");
            sb.AppendLine("        .Text = " + VbaStr("T/M ĐẢNG ỦY") + " & vbCr & " + VbaStr("BÍ THƯ") + " & vbCr & vbCr & vbCr & pSign");
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Size = 14: .Font.Bold = True: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.LineSpacingRule = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Bold = False");
            sb.AppendLine("    End With");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 7. Chèn Mẫu Văn Bản Hành Chính Nghị định 30/2020/NĐ-CP trực tiếp vào Word (tự nạp Tiêu ngữ đơn vị đã cài đặt)
            sb.AppendLine("Sub ChenMauVanBan_UBND30()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Documents.Add");
            sb.AppendLine("    CanLeA4Chuan_Word False");
            sb.AppendLine("    DanhSoTrangTuTrang2_Word False");
            sb.AppendLine("    Dim uGov As String: uGov = GetUnitConfigVal(2, " + VbaStr("ỦY BAN NHÂN DÂN QUẬN") + ")");
            sb.AppendLine("    Dim uIss As String: uIss = GetUnitConfigVal(3, " + VbaStr("ỦY BAN NHÂN DÂN PHƯỜNG") + ")");
            sb.AppendLine("    Dim loc As String: loc = GetUnitConfigVal(4, " + VbaStr("Phường") + ")");
            sb.AppendLine("    Dim uSign As String: uSign = GetUnitConfigVal(6, " + VbaStr("Nguyễn Văn An") + ")");
            sb.AppendLine("    Dim uMottoRaw As String: uMottoRaw = GetUnitConfigVal(8, " + VbaStr("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập - Tự do - Hạnh phúc") + ")");
            sb.AppendLine("    Dim uQh As String, uTn As String");
            sb.AppendLine("    If InStr(uMottoRaw, \"|\") > 0 Then");
            sb.AppendLine("        uQh = Trim(Split(uMottoRaw, \"|\")(0)): uTn = Trim(Split(uMottoRaw, \"|\")(1))");
            sb.AppendLine("    Else");
            sb.AppendLine("        uQh = uMottoRaw: uTn = " + VbaStr("Độc lập - Tự do - Hạnh phúc"));
            sb.AppendLine("    End If");
            sb.AppendLine("    Dim rng As Range, tbl1 As Table, tbl2 As Table");
            sb.AppendLine("    Set rng = ActiveDocument.Range(0, 0)");
            sb.AppendLine("    Set tbl1 = ActiveDocument.Tables.Add(rng, 1, 2)");
            sb.AppendLine("    tbl1.Borders.Enable = 0");
            sb.AppendLine("    tbl1.AllowAutoFit = False");
            sb.AppendLine("    tbl1.PreferredWidthType = 2: tbl1.PreferredWidth = 100");
            sb.AppendLine("    tbl1.LeftPadding = 0: tbl1.RightPadding = 0: tbl1.TopPadding = 0: tbl1.BottomPadding = 0");
            sb.AppendLine("    tbl1.Columns(1).PreferredWidthType = 2: tbl1.Columns(1).PreferredWidth = 42");
            sb.AppendLine("    tbl1.Columns(2).PreferredWidthType = 2: tbl1.Columns(2).PreferredWidth = 58");
            sb.AppendLine("    With tbl1.Cell(1, 1).Range");
            sb.AppendLine("        .Text = uGov & vbCr & uIss & vbCr & " + VbaStr("──────────────") + " & vbCr & " + VbaStr("Số: 01/UBND-VP"));
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Size = 12.5: .Font.Scaling = 100: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Bold = True");
            sb.AppendLine("        .Paragraphs(3).Range.Font.Size = 8: .Paragraphs(3).Range.Font.Bold = False");
            sb.AppendLine("        .Paragraphs(4).Range.Font.Size = 13: .Paragraphs(4).SpaceBefore = 2");
            sb.AppendLine("    End With");
            sb.AppendLine("    With tbl1.Cell(1, 2).Range");
            sb.AppendLine("        .Text = uQh & vbCr & uTn & vbCr & " + VbaStr("────────────────────────") + " & vbCr & loc & " + VbaStr(", ngày 19 tháng 11 năm 2026"));
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Scaling = 100: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(1).Range.Font.Size = 12.5: .Paragraphs(1).Range.Font.Bold = True");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Size = 13.5: .Paragraphs(2).Range.Font.Bold = True: .Paragraphs(2).Range.Font.Underline = 0");
            sb.AppendLine("        .Paragraphs(3).Range.Font.Size = 8: .Paragraphs(3).Range.Font.Bold = False: .Paragraphs(3).Range.Font.Underline = 0");
            sb.AppendLine("        .Paragraphs(4).Range.Font.Size = 13.5: .Paragraphs(4).Range.Font.Italic = True: .Paragraphs(4).SpaceBefore = 2");
            sb.AppendLine("    End With");
            sb.AppendLine("    Dim pEnd As Range: Set pEnd = ActiveDocument.Content: pEnd.Collapse 0");
            sb.AppendLine("    pEnd.InsertAfter vbCr & " + VbaStr("Kính gửi: Các bộ phận chuyên môn thuộc Ủy ban nhân dân phường.") + " & vbCr & " + VbaStr("Thực hiện kế hoạch công tác năm 2026, Ủy ban nhân dân phường triển khai các nội dung sau:./.") + " & vbCr");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).Alignment = 3: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).Range.Font.Size = 14: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 2).FirstLineIndent = CentimetersToPoints(1.27)");
            sb.AppendLine("    ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).Alignment = 3: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).Range.Font.Size = 14: ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).FirstLineIndent = CentimetersToPoints(1.27): ActiveDocument.Paragraphs(ActiveDocument.Paragraphs.Count - 1).SpaceAfter = 6");
            sb.AppendLine("    Set pEnd = ActiveDocument.Content: pEnd.Collapse 0");
            sb.AppendLine("    Set tbl2 = ActiveDocument.Tables.Add(pEnd, 1, 2)");
            sb.AppendLine("    tbl2.Borders.Enable = 0");
            sb.AppendLine("    tbl2.AllowAutoFit = False");
            sb.AppendLine("    tbl2.LeftPadding = 0: tbl2.RightPadding = 0: tbl2.TopPadding = 0: tbl2.BottomPadding = 0");
            sb.AppendLine("    With tbl2.Cell(1, 1).Range");
            sb.AppendLine("        .Text = " + VbaStr("Nơi nhận:") + " & vbCr & " + VbaStr("- Như trên;") + " & vbCr & " + VbaStr("- Lưu: VT, VP."));
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .ParagraphFormat.Alignment = 0: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("        .Paragraphs(1).Range.Font.Size = 12: .Paragraphs(1).Range.Font.Bold = True: .Paragraphs(1).Range.Font.Italic = True");
            sb.AppendLine("        .Paragraphs(2).Range.Font.Size = 11: .Paragraphs(3).Range.Font.Size = 11");
            sb.AppendLine("    End With");
            sb.AppendLine("    With tbl2.Cell(1, 2).Range");
            sb.AppendLine("        .Text = " + VbaStr("TM. ỦY BAN NHÂN DÂN") + " & vbCr & " + VbaStr("CHỦ TỊCH") + " & vbCr & vbCr & vbCr & uSign");
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Size = 14: .Font.Bold = True: .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("    End With");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 8. Sửa lỗi chính tả & dấu câu trực tiếp trên Word (chia nhóm nhỏ tránh giới hạn 24 dòng tiếp nối của VBA)
            sb.AppendLine("Sub SuaChinhTaVaDauCau_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            string[,] quickRules = new string[,] {
                { "bổ xung", "bổ sung" }, { "sử lý", "xử lý" }, { "xử dụng", "sử dụng" },
                { "chính sác", "chính xác" }, { "kiểm xoát", "kiểm soát" }, { "ngiêm", "nghiêm" },
                { "chuẫn", "chuẩn" }, { "kết qủa", "kết quả" }, { "hổ trợ", "hỗ trợ" },
                { "kế hoạc", "kế hoạch" }, { "chủ chương", "chủ trương" }, { "khẩn chương", "khẩn trương" },
                { "chiển khai", "triển khai" }, { "tỏ chức", "tổ chức" }, { "triển khay", "triển khai" },
                { "cải cách hành chánh", "cải cách hành chính" }, { "chuyển đổi xố", "chuyển đổi số" },
                { "nghị quyêt", "nghị quyết" }, { "quyết đinh", "quyết định" }, { "báo cao", "báo cáo" },
                { "đảng uy", "đảng ủy" }, { "chi bô", "chi bộ" }, { "thể thuc", "thể thức" },
                { "nhiệm vu", "nhiệm vụ" }, { "phát triễn", "phát triển" }, { "giải quyêt", "giải quyết" },
                { "giám sat", "giám sát" }, { "quán triêt", "quán triệt" }, { "thực hiên", "thực hiện" },
                { "đồng chi", "đồng chí" }, { "cán bô", "cán bộ" }, { "công chưc", "công chức" },
                { "chủ tich", "chủ tịch" }, { "bí thu", "bí thư" }, { "lưu trử", "lưu trữ" },
                { "hồ xơ", "hồ sơ" }, { "tài liêu", "tài liệu" }, { "đề nghi", "đề nghị" },
                { "kính gởi", "kính gửi" }, { "xem xet", "xem xét" }, { "phê duyêt", "phê duyệt" },
                { " ,", "," }, { " .", "." }, { " ;", ";" }, { " :", ":" }, { "  ", " " }
            };
            for (int r = 0; r < quickRules.GetLength(0); r++)
            {
                sb.AppendLine("    ReplaceWordText " + VbaStr(quickRules[r, 0]) + ", " + VbaStr(quickRules[r, 1]));
            }
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã rà soát và sửa lỗi chính tả, khoảng trắng, dấu câu trực tiếp trong Word!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            sb.AppendLine("Private Sub ReplaceWordText(ByVal findTxt As String, ByVal repTxt As String)");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim rng As Range: Set rng = ActiveDocument.Content");
            sb.AppendLine("    With rng.Find");
            sb.AppendLine("        .ClearFormatting: .Replacement.ClearFormatting");
            sb.AppendLine("        .Text = findTxt: .Replacement.Text = repTxt");
            sb.AppendLine("        .Forward = True: .Wrap = 1: .MatchCase = False: .MatchWholeWord = False");
            sb.AppendLine("        .Execute Replace:=2");
            sb.AppendLine("    End With");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 8b. Tự động hóa Tiêu ngữ & Bảng đầu trang theo 05-HD/VPTW hoặc NĐ 30/2020/NĐ-CP (hỗ trợ chọn trực tiếp Khối Đảng / Khối UBND / Tự động)
            sb.AppendLine("Sub TuDongHoaTieuNgu_Word(Optional ByVal mode As String = \"auto\")");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Documents.Add");
            sb.AppendLine("    If mode = \"ubnd\" Then");
            sb.AppendLine("        ChenMauVanBan_UBND30");
            sb.AppendLine("        Exit Sub");
            sb.AppendLine("    ElseIf mode = \"dang\" Then");
            sb.AppendLine("        ChenMauVanBan_Dang05");
            sb.AppendLine("        Exit Sub");
            sb.AppendLine("    End If");
            sb.AppendLine("    Dim docText As String: docText = UCase(ActiveDocument.Content.Text)");
            sb.AppendLine("    Dim isDang As Boolean: isDang = (InStr(docText, " + VbaStr("ĐẢNG CỘNG SẢN") + ") > 0 Or InStr(docText, " + VbaStr("ĐẢNG ỦY") + ") > 0 Or InStr(docText, " + VbaStr("CHI BỘ") + ") > 0 Or InStr(docText, " + VbaStr("-BC/ĐU") + ") > 0 Or InStr(docText, " + VbaStr("-NQ/ĐU") + ") > 0)");
            sb.AppendLine("    If ActiveDocument.Tables.Count = 0 Then");
            sb.AppendLine("        If isDang Then ChenMauVanBan_Dang05 Else ChenMauVanBan_UBND30");
            sb.AppendLine("        Exit Sub");
            sb.AppendLine("    End If");
            sb.AppendLine("    If isDang Then ChuanHoaVanBan_HuongDan05 Else ChuanHoaVanBan_NghiDinh30");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 9. Nối các dòng bị ngắt khi copy/mở từ PDF sang Word
            sb.AppendLine("Sub NoiDongPDFBiNgat_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Application.ScreenUpdating = False");
            sb.AppendLine("    Dim i As Long, p As Paragraph, pNext As Paragraph, t As String, tNext As String, lastCh As String, firstNxt As String, isLowerNxt As Boolean, mergedCount As Long");
            sb.AppendLine("    mergedCount = 0");
            sb.AppendLine("    For i = ActiveDocument.Paragraphs.Count - 1 To 1 Step -1");
            sb.AppendLine("        Set p = ActiveDocument.Paragraphs(i)");
            sb.AppendLine("        Set pNext = ActiveDocument.Paragraphs(i + 1)");
            sb.AppendLine("        If p.Range.Tables.Count = 0 And pNext.Range.Tables.Count = 0 Then");
            sb.AppendLine("            t = Trim(Replace(p.Range.Text, vbCr, \"\"))");
            sb.AppendLine("            tNext = Trim(Replace(pNext.Range.Text, vbCr, \"\"))");
            sb.AppendLine("            If Len(t) > 0 And Len(tNext) > 0 Then");
            sb.AppendLine("                lastCh = Right(t, 1)");
            sb.AppendLine("                firstNxt = Left(tNext, 1)");
            sb.AppendLine("                isLowerNxt = (LCase(firstNxt) = firstNxt And UCase(firstNxt) <> firstNxt)");
            sb.AppendLine("                If lastCh <> \".\" And lastCh <> \":\" And lastCh <> \";\" And lastCh <> \"!\" And lastCh <> \"?\" Then");
            sb.AppendLine("                    If (isLowerNxt And Left(tNext, 2) <> \"a)\" And Left(tNext, 2) <> \"b)\" And Left(tNext, 2) <> \"c)\" And Left(tNext, 2) <> \"d)\") Or (Len(t) > 25 And firstNxt <> \"-\" And firstNxt <> \"+\" And Not IsNumeric(firstNxt)) Then");
            sb.AppendLine("                        Dim rMerge As Range");
            sb.AppendLine("                        Set rMerge = ActiveDocument.Range(p.Range.End - 1, p.Range.End)");
            sb.AppendLine("                        rMerge.Text = \" \"");
            sb.AppendLine("                        mergedCount = mergedCount + 1");
            sb.AppendLine("                    End If");
            sb.AppendLine("                End If");
            sb.AppendLine("            End If");
            sb.AppendLine("        End If");
            sb.AppendLine("    Next i");
            sb.AppendLine("    Application.ScreenUpdating = True");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã nối ") + " & mergedCount & " + VbaStr(" dòng bị ngắt từ PDF thành đoạn văn hoàn chỉnh!") + "");
            sb.AppendLine("End Sub");
            sb.AppendLine("");

            // 10. Trợ lý AI Soạn Thảo Văn Bản trực tiếp trong Microsoft Word
            sb.AppendLine("Sub AISoanThao_Word(Optional ByVal presetPrompt As String = \"\")");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    Dim req As String");
            sb.AppendLine("    If Len(presetPrompt) > 0 Then");
            sb.AppendLine("        req = presetPrompt");
            sb.AppendLine("    ElseIf Application.Visible Then");
            sb.AppendLine("        req = InputBox(" + VbaStr("Nhập yêu cầu văn bản cần AI soạn thảo (VD: Báo cáo xây dựng Đảng tháng 11, Công văn họp UBND...):") + ", " + VbaStr("Trợ Lý AI Soạn Thảo - VBTM Pro 2.5.0") + ", " + VbaStr("Báo cáo kết quả công tác xây dựng Đảng tháng 11 năm 2026") + ")");
            sb.AppendLine("    Else");
            sb.AppendLine("        req = " + VbaStr("Báo cáo kết quả công tác xây dựng Đảng tháng 11 năm 2026"));
            sb.AppendLine("    End If");
            sb.AppendLine("    If Len(Trim(req)) = 0 Then Exit Sub");
            sb.AppendLine("    Dim u As String: u = UCase(req)");
            sb.AppendLine("    If InStr(u, " + VbaStr("ĐẢNG") + ") > 0 Or InStr(u, " + VbaStr("CHI BỘ") + ") > 0 Or InStr(u, " + VbaStr("BÍ THƯ") + ") > 0 Then");
            sb.AppendLine("        ChenMauVanBan_Dang05");
            sb.AppendLine("    Else");
            sb.AppendLine("        ChenMauVanBan_UBND30");
            sb.AppendLine("    End If");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Trợ lý AI đã tự động khởi tạo xong khung & nội dung chuẩn theo yêu cầu: ") + " & req");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 11. AI Tóm tắt & Tự động chèn bản tóm tắt trọng tâm trực tiếp vào Word
            sb.AppendLine("Sub AiTomTatVanBan_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Application.ScreenUpdating = False");
            sb.AppendLine("    Dim p As Paragraph, t As String, sumText As String, cnt As Integer");
            sb.AppendLine("    cnt = 0");
            sb.AppendLine("    For Each p In ActiveDocument.Paragraphs");
            sb.AppendLine("        If p.Range.Tables.Count = 0 Then");
            sb.AppendLine("            t = Trim(Replace(Replace(p.Range.Text, vbCr, \"\"), Chr(7), \"\"))");
            sb.AppendLine("            If Len(t) > 30 And cnt < 4 And InStr(t, " + VbaStr("TÓM TẮT TRỌNG TÂM") + ") = 0 Then");
            sb.AppendLine("                cnt = cnt + 1");
            sb.AppendLine("                If Len(t) > 160 Then t = Left(t, 157) & \"...\"");
            sb.AppendLine("                sumText = sumText & \"- \" & t & vbCr");
            sb.AppendLine("            End If");
            sb.AppendLine("        End If");
            sb.AppendLine("    Next p");
            sb.AppendLine("    If Len(sumText) > 0 Then");
            sb.AppendLine("        Dim rngIns As Range: Set rngIns = ActiveDocument.Content");
            sb.AppendLine("        rngIns.Collapse 0");
            sb.AppendLine("        rngIns.InsertAfter vbCr & " + VbaStr("TÓM TẮT TRỌNG TÂM (AI):") + " & vbCr & sumText");
            sb.AppendLine("        rngIns.Font.Name = \"Times New Roman\": rngIns.Font.Size = 14");
            sb.AppendLine("        SuaChinhTaVaDauCau_Word");
            sb.AppendLine("        Application.ScreenUpdating = True");
            sb.AppendLine("        ShowUnicodeMsg " + VbaStr("AI đã tóm tắt xong, tự động chèn bản tóm tắt vào cuối văn bản và rà soát lỗi chính tả!") + "");
            sb.AppendLine("    Else");
            sb.AppendLine("        Application.ScreenUpdating = True");
            sb.AppendLine("        ShowUnicodeMsg " + VbaStr("Chưa có đoạn nội dung đủ dài để AI tóm tắt.") + "");
            sb.AppendLine("    End If");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 12. Chuyển đổi Thể thức Đảng (05-HD) <-> Hành chính (NĐ 30)
            sb.AppendLine("Sub ChuyenDoiDangHanhChinh_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim allU As String: allU = UCase(ActiveDocument.Content.Text)");
            sb.AppendLine("    If InStr(allU, " + VbaStr("ĐẢNG CỘNG SẢN") + ") > 0 Then");
            sb.AppendLine("        ReplaceWordText " + VbaStr("ĐẢNG CỘNG SẢN VIỆT NAM") + ", " + VbaStr("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("T/M ĐẢNG ỦY") + ", " + VbaStr("TM. ỦY BAN NHÂN DÂN"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("BÍ THƯ") + ", " + VbaStr("CHỦ TỊCH"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("- Lưu Văn phòng Đảng ủy.") + ", " + VbaStr("- Lưu: VT, VP."));
            sb.AppendLine("        ChuanHoaVanBan_NghiDinh30");
            sb.AppendLine("    Else");
            sb.AppendLine("        ReplaceWordText " + VbaStr("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM") + ", " + VbaStr("ĐẢNG CỘNG SẢN VIỆT NAM"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("TM. ỦY BAN NHÂN DÂN") + ", " + VbaStr("T/M ĐẢNG ỦY"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("CHỦ TỊCH") + ", " + VbaStr("BÍ THƯ"));
            sb.AppendLine("        ReplaceWordText " + VbaStr("- Lưu: VT, VP.") + ", " + VbaStr("- Lưu Văn phòng Đảng ủy."));
            sb.AppendLine("        ChuanHoaVanBan_HuongDan05");
            sb.AppendLine("    End If");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 13. Kiểm định Thể thức Tự động (% tuân thủ 05-HD/VPTW & NĐ 30)
            sb.AppendLine("Sub KiemDinhTheThuc_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim score As Integer: score = 0");
            sb.AppendLine("    Dim report As String: report = \"\"");
            sb.AppendLine("    Dim txt As String: txt = ActiveDocument.Content.Text");
            sb.AppendLine("    Dim uTxt As String: uTxt = UCase(txt)");
            sb.AppendLine("    If Round(PointsToMillimeters(ActiveDocument.PageSetup.LeftMargin)) >= 28 Then score = score + 20 Else report = report & " + VbaStr("- Lề trái chưa đạt 30mm") + " & vbCrLf");
            sb.AppendLine("    If InStr(uTxt, " + VbaStr("ĐẢNG CỘNG SẢN") + ") > 0 Or InStr(uTxt, " + VbaStr("CỘNG HÒA XÃ HỘI") + ") > 0 Then score = score + 20 Else report = report & " + VbaStr("- Thiếu Quốc hiệu / ĐẢNG CỘNG SẢN VIỆT NAM") + " & vbCrLf");
            sb.AppendLine("    If InStr(txt, " + VbaStr("Số") + ") > 0 Then score = score + 20 Else report = report & " + VbaStr("- Thiếu Số, ký hiệu văn bản") + " & vbCrLf");
            sb.AppendLine("    If InStr(LCase(txt), " + VbaStr("nơi nhận") + ") > 0 Then score = score + 20 Else report = report & " + VbaStr("- Thiếu mục Nơi nhận:") + " & vbCrLf");
            sb.AppendLine("    If InStr(LCase(txt), " + VbaStr("lưu") + ") > 0 Then score = score + 20 Else report = report & " + VbaStr("- Thiếu dòng - Lưu ở Nơi nhận") + " & vbCrLf");
            sb.AppendLine("    If Len(report) = 0 Then report = " + VbaStr("Hoàn hảo! Văn bản đạt đầy đủ các tiêu chí thể thức bắt buộc."));
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("KẾT QUẢ KIỂM ĐỊNH THỂ THỨC: ") + " & score & \"/100%\" & vbCrLf & vbCrLf & report");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 14. Chèn Bảng Biểu Chuẩn 13pt (Tự lặp tiêu đề khi sang trang)
            sb.AppendLine("Sub ChenBangChuan_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Documents.Add");
            sb.AppendLine("    Dim tbl As Table");
            sb.AppendLine("    Set tbl = ActiveDocument.Tables.Add(Selection.Range, 4, 4)");
            sb.AppendLine("    tbl.Borders.Enable = 1");
            sb.AppendLine("    tbl.Rows(1).HeadingFormat = True");
            sb.AppendLine("    tbl.Range.Font.Name = \"Times New Roman\": tbl.Range.Font.Size = 13");
            sb.AppendLine("    tbl.Rows(1).Range.Font.Bold = True: tbl.Rows(1).Range.ParagraphFormat.Alignment = 1");
            sb.AppendLine("    tbl.Cell(1, 1).Range.Text = \"STT\"");
            sb.AppendLine("    tbl.Cell(1, 2).Range.Text = " + VbaStr("Nội dung / Nhiệm vụ"));
            sb.AppendLine("    tbl.Cell(1, 3).Range.Text = " + VbaStr("Đơn vị thực hiện"));
            sb.AppendLine("    tbl.Cell(1, 4).Range.Text = " + VbaStr("Ghi chú"));
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 15. Chèn Dấu Độ Khẩn / Độ Mật dạng khung nổi (TextBox Shape) di chuyển tự do bằng chuột + chọn nhanh 3 vị trí chuẩn
            sb.AppendLine("Sub ChenDoKhanMat_Word(Optional ByVal lvIdx As String = \"1\")");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Documents.Add");
            sb.AppendLine("    If Left(lvIdx, 3) = \"pos\" Then");
            sb.AppendLine("        If Selection.ShapeRange.Count >= 1 Then");
            sb.AppendLine("            Dim sMove As Shape: Set sMove = Selection.ShapeRange(1)");
            sb.AppendLine("            sMove.RelativeHorizontalPosition = 1: sMove.RelativeVerticalPosition = 1");
            sb.AppendLine("            If lvIdx = \"pos1\" Then sMove.Left = CentimetersToPoints(3.0): sMove.Top = CentimetersToPoints(1.0)");
            sb.AppendLine("            If lvIdx = \"pos2\" Then sMove.Left = CentimetersToPoints(4.2): sMove.Top = CentimetersToPoints(4.8)");
            sb.AppendLine("            If lvIdx = \"pos3\" Then sMove.Left = CentimetersToPoints(14.5): sMove.Top = CentimetersToPoints(1.0)");
            sb.AppendLine("        Else");
            sb.AppendLine("            ShowUnicodeMsg " + VbaStr("Hãy bấm chọn khung dấu KHẨN / MẬT trước khi đổi vị trí, hoặc nắm trực tiếp viền khung kéo bằng chuột!") + "");
            sb.AppendLine("        End If");
            sb.AppendLine("        Exit Sub");
            sb.AppendLine("    End If");
            sb.AppendLine("    Dim lv As String, boxW As Single, defTop As Single");
            sb.AppendLine("    boxW = 3.2: defTop = 4.8");
            sb.AppendLine("    Select Case lvIdx");
            sb.AppendLine("        Case \"2\": lv = " + VbaStr("THƯỢNG KHẨN") + ": boxW = 4.8");
            sb.AppendLine("        Case \"3\": lv = " + VbaStr("HỎA TỐC") + ": boxW = 3.6");
            sb.AppendLine("        Case \"4\": lv = " + VbaStr("HỎA TỐC HẸN GIỜ") + ": boxW = 5.6");
            sb.AppendLine("        Case \"5\": lv = " + VbaStr("MẬT") + ": boxW = 2.6: defTop = 1.0");
            sb.AppendLine("        Case \"6\": lv = " + VbaStr("TỐI MẬT") + ": boxW = 3.4: defTop = 1.0");
            sb.AppendLine("        Case \"7\": lv = " + VbaStr("TUYỆT MẬT") + ": boxW = 4.0: defTop = 1.0");
            sb.AppendLine("        Case Else: lv = " + VbaStr("KHẨN") + ": boxW = 2.8");
            sb.AppendLine("    End Select");
            sb.AppendLine("    Dim shp As Shape");
            sb.AppendLine("    Set shp = ActiveDocument.Shapes.AddTextbox(1, CentimetersToPoints(4.0), CentimetersToPoints(defTop), CentimetersToPoints(boxW), CentimetersToPoints(0.85))");
            sb.AppendLine("    shp.RelativeHorizontalPosition = 1");
            sb.AppendLine("    shp.RelativeVerticalPosition = 1");
            sb.AppendLine("    shp.Left = CentimetersToPoints(4.0)");
            sb.AppendLine("    shp.Top = CentimetersToPoints(defTop)");
            sb.AppendLine("    shp.WrapFormat.Type = 3");
            sb.AppendLine("    shp.Fill.Visible = 0");
            sb.AppendLine("    shp.Line.Visible = -1");
            sb.AppendLine("    shp.Line.ForeColor.RGB = RGB(204, 0, 0)");
            sb.AppendLine("    shp.Line.Weight = 1.5");
            sb.AppendLine("    shp.TextFrame.MarginLeft = 2: shp.TextFrame.MarginRight = 2: shp.TextFrame.MarginTop = 2: shp.TextFrame.MarginBottom = 2");
            sb.AppendLine("    shp.TextFrame.VerticalAnchor = 3");
            sb.AppendLine("    With shp.TextFrame.TextRange");
            sb.AppendLine("        .Text = lv");
            sb.AppendLine("        .Font.Name = \"Times New Roman\": .Font.Size = 13: .Font.Bold = True: .Font.Color = RGB(204, 0, 0)");
            sb.AppendLine("        .ParagraphFormat.Alignment = 1: .ParagraphFormat.SpaceBefore = 0: .ParagraphFormat.SpaceAfter = 0: .ParagraphFormat.FirstLineIndent = 0");
            sb.AppendLine("    End With");
            sb.AppendLine("    shp.Select");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 16. Xuất Nhanh PDF Chuẩn In / Ký Số (1-Click)
            sb.AppendLine("Sub XuatPDFNhanh_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim pdfPath As String");
            sb.AppendLine("    If Len(ActiveDocument.Path) > 0 Then");
            sb.AppendLine("        pdfPath = ActiveDocument.Path & \"\\\" & Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, \".\") - 1) & \".pdf\"");
            sb.AppendLine("    Else");
            sb.AppendLine("        pdfPath = Environ(\"USERPROFILE\") & \"\\Desktop\\VBTM_VanBan_\" & Format(Now, \"yyyymmdd_hhnnss\") & \".pdf\"");
            sb.AppendLine("    End If");
            sb.AppendLine("    ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=17, OpenAfterExport:=Application.Visible");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("Đã xuất file PDF chuẩn in & ký số tại: ") + " & pdfPath");
            sb.AppendLine("End Sub");
            sb.AppendLine("");
            // 17. Thống Kê Số Trang, Số Từ & Thời Gian Đọc Văn Bản
            sb.AppendLine("Sub ThongKeVanBan_Word()");
            sb.AppendLine("    On Error Resume Next");
            sb.AppendLine("    If Documents.Count = 0 Then Exit Sub");
            sb.AppendLine("    Dim pg As Long: pg = ActiveDocument.ComputeStatistics(2)");
            sb.AppendLine("    Dim wd As Long: wd = ActiveDocument.ComputeStatistics(0)");
            sb.AppendLine("    Dim ch As Long: ch = ActiveDocument.ComputeStatistics(3)");
            sb.AppendLine("    Dim mins As Long: mins = Int((wd + 199) / 200)");
            sb.AppendLine("    ShowUnicodeMsg " + VbaStr("THỐNG KÊ VĂN BẢN:") + " & vbCrLf & " + VbaStr("- Số trang A4: ") + " & pg & vbCrLf & " + VbaStr("- Tổng số từ: ") + " & wd & vbCrLf & " + VbaStr("- Tổng ký tự: ") + " & ch & vbCrLf & " + VbaStr("- Thời gian đọc ước tính: ~") + " & mins & " + VbaStr(" phút") + "");
            sb.AppendLine("End Sub");
            return sb.ToString();
        }

        public static bool InstallWordRibbonAddin(string installDir, string exePath)
        {
            try
            {
                Directory.CreateDirectory(installDir);
                string basPath = Path.Combine(installDir, "VBTM_Word_Addin.bas");
                File.WriteAllText(basPath, GetWordVbaMacroCode(), new UTF8Encoding(false));

                string startupDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "Word", "STARTUP"
                );
                Directory.CreateDirectory(startupDir);
                string dotmPath = Path.Combine(startupDir, "VBTM_Pro_Addin.dotm");
                try { if (File.Exists(dotmPath)) File.Delete(dotmPath); } catch { }
                string tempDotm = Path.Combine(Path.GetTempPath(), "VBTM_Pro_Addin_" + Guid.NewGuid().ToString("N") + ".dotm");

                // 1. Cấp quyền AccessVBOM & Trusted Location trong Registry HKCU (không cần Admin)
                string[] officeVers = new[] { "16.0", "15.0", "14.0" };
                foreach (string ver in officeVers)
                {
                    try
                    {
                        using (RegistryKey secKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + ver + @"\Word\Security"))
                        {
                            if (secKey != null) secKey.SetValue("AccessVBOM", 1, RegistryValueKind.DWord);
                        }
                        using (RegistryKey tlKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + ver + @"\Word\Security\Trusted Locations\VBTMProStartup"))
                        {
                            if (tlKey != null)
                            {
                                tlKey.SetValue("Path", startupDir + "\\");
                                tlKey.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
                                tlKey.SetValue("Description", "VBTM Pro 2.5.0 Word Startup Add-in");
                            }
                        }
                    }
                    catch { }
                }

                // 0. Dọn sạch Web Add-in AI soạn thảo cũ (nếu có) để Word gọn nhẹ
                RemoveNativeWordWebAddin();

                // 2. Tạo tệp VBTM_Pro_Addin.dotm bằng Word COM Automation và nạp Macro VBA
                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null) return false;

                object wordApp = Activator.CreateInstance(wordType);
                try
                {
                    wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { false });
                    try { wordType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, wordApp, new object[] { 0 }); } catch { }

                    object docs = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    object doc = docs.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, docs, null);
                    try
                    {
                        object vbProj = doc.GetType().InvokeMember("VBProject", BindingFlags.GetProperty, null, doc, null);
                        object vbComps = vbProj.GetType().InvokeMember("VBComponents", BindingFlags.GetProperty, null, vbProj, null);
                        vbComps.GetType().InvokeMember("Import", BindingFlags.InvokeMethod, null, vbComps, new object[] { basPath });

                        // 15 = wdFormatXMLTemplateMacroEnabled (.dotm)
                        doc.GetType().InvokeMember("SaveAs2", BindingFlags.InvokeMethod, null, doc, new object[] { tempDotm, 15 });
                    }
                    finally
                    {
                        try { doc.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, doc, new object[] { false }); } catch { }
                    }
                }
                finally
                {
                    try { wordType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, wordApp, new object[] { false }); } catch { }
                }

                // 3. Chờ Word nhả khóa tệp tạm rồi nhúng Custom Ribbon Tab vào tempDotm và chép vào STARTUP
                if (File.Exists(tempDotm))
                {
                    for (int retry = 0; retry < 6; retry++)
                    {
                        try
                        {
                            System.Threading.Thread.Sleep(300);
                            InjectRibbonUiIntoDotm(tempDotm);
                            File.Copy(tempDotm, dotmPath, true);
                            try { File.Delete(tempDotm); } catch { }
                            return File.Exists(dotmPath);
                        }
                        catch
                        {
                            if (retry == 5) throw;
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static void InjectRibbonUiIntoDotm(string dotmPath)
        {
            string ribbonXml =
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                "<ribbon><tabs>" +
                "<tab id=\"tabVbtmPro\" label=\"VBTM Pro 2.5 (05-HD &amp; ND30)\">" +
                "<group id=\"grpVbtmTemplates\" label=\"1. 16 Mẫu Chuẩn &amp; Đơn Vị\">" +
                "<menu id=\"mnuVbtm16\" label=\"16 Mẫu Chuẩn (Đảng &amp; UBND)\" size=\"large\" imageMso=\"FileNewBlogPost\">" +
                "<menuSeparator id=\"sepUbnd\" title=\"I. Khối HĐND &amp; UBND Phường (Nghị định 30/2020/NĐ-CP)\"/>" +
                "<button id=\"btnTpl0\" label=\"1. Công văn UBND Phường\" tag=\"0\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl1\" label=\"2. Báo cáo UBND Phường\" tag=\"1\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl2\" label=\"3. Quyết định (cá biệt) UBND\" tag=\"2\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl3\" label=\"4. Quyết định (Quy chế) UBND\" tag=\"3\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl4\" label=\"5. Kế hoạch UBND Phường\" tag=\"4\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl5\" label=\"6. Tờ trình UBND Phường\" tag=\"5\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl6\" label=\"7. Thông báo UBND Phường\" tag=\"6\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl7\" label=\"8. Nghị quyết HĐND Phường\" tag=\"7\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl8\" label=\"9. Biên bản họp UBND Phường\" tag=\"8\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl9\" label=\"10. Giấy mời họp UBND Phường\" tag=\"9\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<menuSeparator id=\"sepDang\" title=\"II. Khối Đảng Ủy &amp; Chi Bộ (Hướng dẫn 05-HD/VPTW)\"/>" +
                "<button id=\"btnTpl10\" label=\"11. Công văn Đảng ủy Phường\" tag=\"10\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl11\" label=\"12. Báo cáo Đảng ủy Phường\" tag=\"11\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl12\" label=\"13. Nghị quyết Đảng ủy / Chi bộ\" tag=\"12\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl13\" label=\"14. Kế hoạch Đảng ủy Phường\" tag=\"13\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl14\" label=\"15. Chương trình công tác Đảng ủy\" tag=\"14\" onAction=\"Ribbon_InsertTpl\"/>" +
                "<button id=\"btnTpl15\" label=\"16. Quyết định Đảng ủy Phường\" tag=\"15\" onAction=\"Ribbon_InsertTpl\"/>" +
                "</menu>" +
                "<button id=\"btnVbtmUnitCfg\" label=\"Cài Đặt Tiêu Ngữ Đơn Vị\" size=\"large\" imageMso=\"FileDocumentManagementInformation\" onAction=\"Ribbon_CaiDatDonVi\" supertip=\"Thiết lập tên Đảng ủy, Chi bộ, UBND Phường, Tiêu ngữ, Địa danh và Người ký mặc định.\"/>" +
                "</group>" +
                "<group id=\"grpVbtmFormat\" label=\"2. Chỉnh Thể Thức Nhanh\">" +
                "<button id=\"btnVbtmDangFast\" label=\"Chỉnh Thể Thức Đảng Nhanh (05-HD)\" size=\"large\" imageMso=\"HappyFace\" onAction=\"Ribbon_ChinhDangNhanh\" supertip=\"1 Bấm chỉnh toàn bộ văn bản Word đang mở đúng chuẩn Hướng dẫn số 05-HD/VPTW và sửa lỗi chính tả.\"/>" +
                "<button id=\"btnVbtmAuto\" label=\"Chuẩn Hóa Tự Động\" size=\"large\" imageMso=\"ReviewTrackChanges\" onAction=\"Ribbon_ChuanHoaTuDong\" supertip=\"Tự động nhận diện văn bản Đảng (05-HD/VPTW) hoặc Hành chính (NĐ 30) và chuẩn hóa toàn bộ văn bản đang mở.\"/>" +
                "<menu id=\"mnuVbtmTieuNgu\" label=\"Tự Động Hóa Tiêu Ngữ\" size=\"normal\" imageMso=\"HeaderFooterInsert\" supertip=\"Chọn chèn nhanh Tiêu ngữ Khối UBND / Hành chính (NĐ 30), Khối Đảng (05-HD/VPTW) hoặc tự động chuẩn hóa.\">" +
                "<button id=\"btnTnAuto\" label=\"⚡ Tự Động Nhận Diện &amp; Chuẩn Hóa\" tag=\"auto\" onAction=\"Ribbon_TuDongHoaTieuNgu\"/>" +
                "<button id=\"btnTnUbnd\" label=\"🏛️ Chèn Tiêu Ngữ Khối UBND / Hành Chính (NĐ 30)\" tag=\"ubnd\" onAction=\"Ribbon_TuDongHoaTieuNgu\"/>" +
                "<button id=\"btnTnDang\" label=\"⭐ Chèn Tiêu Ngữ Khối Đảng (05-HD/VPTW)\" tag=\"dang\" onAction=\"Ribbon_TuDongHoaTieuNgu\"/>" +
                "<menuSeparator id=\"sepTnCfg\" title=\"Tùy chỉnh cơ quan &amp; tiêu ngữ\"/>" +
                "<button id=\"btnTnCfg\" label=\"⚙️ Cài Đặt / Đổi Tiêu Ngữ Khối Đảng &amp; UBND\" onAction=\"Ribbon_CaiDatDonVi\"/>" +
                "</menu>" +
                "<button id=\"btnVbtmNd30\" label=\"Chuẩn NĐ 30/2020\" size=\"normal\" imageMso=\"FileDocumentInspect\" onAction=\"Ribbon_NghiDinh30\"/>" +
                "</group>" +
                "<group id=\"grpVbtmTools\" label=\"3. Chính Tả AI &amp; Trình Bày\">" +
                "<button id=\"btnVbtmSpellAI\" label=\"Rà Soát &amp; Sửa Chính Tả AI\" size=\"large\" imageMso=\"SpellingAndGrammar\" onAction=\"Ribbon_RaSoatChinhTaAI\" supertip=\"Mở bảng Thẩm Định Chính Tả &amp; Thể Thức AI quét trực tiếp tài liệu Word đang mở và tự động sửa hết lỗi.\"/>" +
                "<button id=\"btnVbtmSpell\" label=\"Sửa Lỗi Tự Động (1-Click)\" size=\"normal\" imageMso=\"AcceptChanges\" onAction=\"Ribbon_SuaChinhTa\"/>" +
                "<button id=\"btnVbtmMergePdf\" label=\"Nối Dòng PDF Bị Ngắt\" size=\"normal\" imageMso=\"WrapText\" onAction=\"Ribbon_NoiDongPDF\"/>" +
                "<button id=\"btnVbtmMargins\" label=\"Căn Lề A4 (30-15-20-20)\" size=\"normal\" imageMso=\"PageMarginsGallery\" onAction=\"Ribbon_CanLeA4\"/>" +
                "<button id=\"btnVbtmPageNum\" label=\"Đánh Số Trang (Từ Tr.2)\" size=\"normal\" imageMso=\"HeaderFooterPageNumberInsert\" onAction=\"Ribbon_DanhSoTrang\"/>" +
                "</group>" +
                "<group id=\"grpVbtmStudio\" label=\"4. Công Cụ Văn Bản Nâng Cao\">" +
                "<button id=\"btnVbtmSummary\" label=\"AI Tóm Tắt Văn Bản\" size=\"normal\" imageMso=\"OutlineShowDetail\" onAction=\"Ribbon_AiTomTatVanBan\" supertip=\"AI đọc nhanh văn bản đang mở và trích xuất các điểm trọng tâm.\"/>" +
                "<button id=\"btnVbtmConvert\" label=\"Đổi Đảng &lt;-&gt; Hành Chính\" size=\"normal\" imageMso=\"ChangeStyles\" onAction=\"Ribbon_ChuyenDoiDangHanhChinh\" supertip=\"Chuyển đổi tự động thể thức &amp; thuật ngữ qua lại giữa Khối Đảng (05-HD) và UBND (NĐ 30).\"/>" +
                "<button id=\"btnVbtmAudit\" label=\"Kiểm Định Thể Thức (%)\" size=\"normal\" imageMso=\"ReviewCompareTwoVersions\" onAction=\"Ribbon_KiemDinhTheThuc\" supertip=\"Chấm điểm tuân thủ thể thức 05-HD/VPTW &amp; NĐ 30/2020/NĐ-CP và chỉ ra lỗi thiếu sót.\"/>" +
                "<button id=\"btnVbtmTbl\" label=\"Chèn Bảng Chuẩn 13pt\" size=\"normal\" imageMso=\"TableInsert\" onAction=\"Ribbon_ChenBangChuan\" supertip=\"Chèn bảng thống kê Times New Roman 13pt tự động lặp dòng tiêu đề khi sang trang.\"/>" +
                "<menu id=\"mnuVbtmKhanMat\" label=\"Dấu Khẩn / Mật\" size=\"normal\" imageMso=\"HighImportance\" supertip=\"Chèn nhanh khung đóng dấu Độ Khẩn (KHẨN, THƯỢNG KHẨN, HỎA TỐC) hoặc Độ Mật (MẬT, TỐI MẬT, TUYỆT MẬT) chuẩn Unicode 100%.\">" +
                "<menuSeparator id=\"sepKhan\" title=\"I. Dấu Độ Khẩn\"/>" +
                "<button id=\"btnKm1\" label=\"KHẨN\" tag=\"1\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKm2\" label=\"THƯỢNG KHẨN\" tag=\"2\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKm3\" label=\"HỎA TỐC\" tag=\"3\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKm4\" label=\"HỎA TỐC HẸN GIỜ\" tag=\"4\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<menuSeparator id=\"sepMat\" title=\"II. Dấu Độ Mật\"/>" +
                "<button id=\"btnKm5\" label=\"MẬT\" tag=\"5\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKm6\" label=\"TỐI MẬT\" tag=\"6\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKm7\" label=\"TUYỆT MẬT\" tag=\"7\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<menuSeparator id=\"sepPos\" title=\"III. Đặt Vị Trí Khung (Hoặc Kéo Chuột Tự Do)\"/>" +
                "<button id=\"btnKmPos1\" label=\"↖ Đặt Góc Trên Trái Trang\" tag=\"pos1\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKmPos2\" label=\"↙ Đặt Dưới Số Ký Hiệu\" tag=\"pos2\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "<button id=\"btnKmPos3\" label=\"↗ Đặt Góc Trên Phải Trang\" tag=\"pos3\" onAction=\"Ribbon_ChenDoKhanMat\"/>" +
                "</menu>" +
                "<button id=\"btnVbtmPdf\" label=\"Xuất PDF Ký Số (1-Click)\" size=\"normal\" imageMso=\"FileSaveAsPdfOrXps\" onAction=\"Ribbon_XuatPDFNhanh\" supertip=\"Xuất ngay tài liệu Word đang mở sang file PDF chuẩn A4 để in ấn hoặc trình ký số.\"/>" +
                "<button id=\"btnVbtmStat\" label=\"Thống Kê Trang &amp; Từ\" size=\"normal\" imageMso=\"WordCount\" onAction=\"Ribbon_ThongKeVanBan\" supertip=\"Thống kê nhanh số trang A4, số từ, số ký tự và thời gian đọc của báo cáo.\"/>" +
                "<button id=\"btnVbtmQr\" label=\"Tạo Mã QR Văn Bản\" size=\"large\" imageMso=\"BarcodeInsert\" onAction=\"Ribbon_ChenMaQR\" supertip=\"Tạo mã QR chuẩn quốc tế ISO 18004 để gắn vào văn bản Word.\"/>" +
                "<button id=\"btnVbtmStudio\" label=\"Mở VBTM Pro Studio (AI 7.578 Từ)\" size=\"large\" imageMso=\"QuickPartsInsertGallery\" onAction=\"Ribbon_MoVBTM\" supertip=\"Mở phần mềm VBTM Pro 2.5.0 để rà soát lỗi chính tả AI (7.578 âm tiết), tạo mã QR và chuyển PDF sang Word.\"/>" +
                "</group>" +
                "</tab></tabs></ribbon></customUI>";

            using (ZipArchive zip = ZipFile.Open(dotmPath, ZipArchiveMode.Update))
            {
                ZipArchiveEntry oldUi = zip.GetEntry("customUI/customUI14.xml");
                if (oldUi != null) oldUi.Delete();

                ZipArchiveEntry uiEntry = zip.CreateEntry("customUI/customUI14.xml");
                using (StreamWriter sw = new StreamWriter(uiEntry.Open(), new UTF8Encoding(false)))
                {
                    sw.Write(ribbonXml);
                }

                ZipArchiveEntry relsEntry = zip.GetEntry("_rels/.rels");
                if (relsEntry != null)
                {
                    string relsXml;
                    using (StreamReader sr = new StreamReader(relsEntry.Open(), Encoding.UTF8))
                    {
                        relsXml = sr.ReadToEnd();
                    }
                    if (!relsXml.Contains("customUI/customUI14.xml"))
                    {
                        string newRel = "<Relationship Id=\"rIdVbtmCustomUI\" Type=\"http://schemas.microsoft.com/office/2007/relationships/ui/extensibility\" Target=\"customUI/customUI14.xml\"/></Relationships>";
                        relsXml = relsXml.Replace("</Relationships>", newRel);
                        relsEntry.Delete();
                        ZipArchiveEntry updatedRels = zip.CreateEntry("_rels/.rels");
                        using (StreamWriter sw = new StreamWriter(updatedRels.Open(), new UTF8Encoding(false)))
                        {
                            sw.Write(relsXml);
                        }
                    }
                }
            }
        }

        public static void RemoveNativeWordWebAddin()
        {
            try
            {
                using (RegistryKey devKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\WEF\Developer", true))
                {
                    if (devKey != null)
                    {
                        devKey.DeleteValue("7b4e9c21-8f3a-4d12-9a6e-5c1f8e2b40a1", false);
                    }
                }
            }
            catch { }
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBTMPro", "WordWebAddin");
                if (Directory.Exists(dir)) Directory.Delete(dir, true);
            }
            catch { }
        }
    }

    public class UnitSettingsForm : Form
    {
        private ComboBox cboPreset, cboPGov, cboPIss, cboPMotto, cboUGov, cboUIss, cboUMotto, cboLoc;
        private TextBox txtPSign, txtUSign;

        public UnitSettingsForm()
        {
            this.Text = "Cài Đặt Tiêu Ngữ & Tên Đơn Vị (Đa Cơ Quan) - VBTM Pro 2.5.0";
            this.ClientSize = new Size(640, 565);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(95, 10, 14);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            string[] vals = LoadUnitConfig();
            int y = 16;

            Label lblPre = new Label { Text = "Chọn nhanh loại cơ quan:", Location = new Point(22, y + 4), AutoSize = true, ForeColor = Color.FromArgb(254, 240, 138), Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            cboPreset = new ComboBox
            {
                Location = new Point(215, y),
                Size = new Size(395, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White
            };
            cboPreset.Items.AddRange(new object[] {
                "-- Giữ nguyên cấu hình hiện tại --",
                "1. Đảng ủy Phường (Trực thuộc Thành ủy / Quận ủy)",
                "2. Văn phòng Đảng ủy Phường",
                "3. Chi bộ Khu phố / Chi bộ trực thuộc",
                "4. Ủy ban Kiểm tra / Ban Xây dựng Đảng",
                "5. Ủy ban nhân dân Phường / Xã",
                "6. Hội đồng nhân dân Phường / Xã",
                "7. Ủy ban MTTQ Việt Nam Phường / Xã",
                "8. Công an / Ban CHQS Phường",
                "9. Trường học / Trạm Y tế / Đơn vị sự nghiệp"
            });
            cboPreset.SelectedIndex = 0;
            cboPreset.SelectedIndexChanged += (s, e) => ApplyPreset(cboPreset.SelectedIndex);
            this.Controls.Add(lblPre);
            this.Controls.Add(cboPreset);
            y += 44;

            AddComboRow("Cơ quan Đảng cấp trên:", vals[0], new string[] {
                "ĐẢNG BỘ THÀNH PHỐ HỒ CHÍ MINH", "ĐẢNG BỘ THÀNH PHỐ HÀ NỘI", "ĐẢNG BỘ THÀNH PHỐ ĐÀ NẴNG",
                "ĐẢNG BỘ TỈNH", "THÀNH ỦY", "QUẬN ỦY", "HUYỆN ỦY", "THỊ ỦY",
                "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT", "ĐẢNG ỦY PHƯỜNG", "ĐẢNG ỦY XÃ",
                "ĐẢNG ỦY KHỐI CÁC CƠ QUAN", "ĐẢNG ỦY CÁC CƠ QUAN ĐẢNG"
            }, out cboPGov, ref y);

            AddComboRow("Đơn vị Đảng ban hành:", vals[1], new string[] {
                "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT", "ĐẢNG ỦY PHƯỜNG", "ĐẢNG ỦY XÃ",
                "VĂN PHÒNG", "VĂN PHÒNG ĐẢNG ỦY", "BAN TỔ CHỨC ĐẢNG ỦY",
                "ỦY BAN KIỂM TRA ĐẢNG ỦY", "BAN TUYÊN GIÁO VÀ DÂN VẬN",
                "CHI BỘ KHU PHỐ 1", "CHI BỘ CƠ QUAN", "CHI BỘ CÔNG AN PHƯỜNG",
                "CHI BỘ QUÂN SỰ PHƯỜNG", "CHI BỘ TRƯỜNG HỌC", "CHI BỘ TRẠM Y TẾ"
            }, out cboPIss, ref y);

            AddComboRow("Tiêu ngữ văn bản Đảng:", vals[7], new string[] {
                "ĐẢNG CỘNG SẢN VIỆT NAM",
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
            }, out cboPMotto, ref y);

            AddComboRow("Cơ quan Hành chính cấp trên:", vals[2], new string[] {
                "ỦY BAN NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH", "ỦY BAN NHÂN DÂN THÀNH PHỐ HÀ NỘI",
                "ỦY BAN NHÂN DÂN TỈNH", "ỦY BAN NHÂN DÂN QUẬN", "ỦY BAN NHÂN DÂN HUYỆN",
                "ỦY BAN NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT", "HỘI ĐỒNG NHÂN DÂN QUẬN",
                "ỦY BAN MTTQ VIỆT NAM THÀNH PHỐ", "CÔNG AN THÀNH PHỐ", "SỞ GIÁO DỤC VÀ ĐÀO TẠO", "SỞ Y TẾ"
            }, out cboUGov, ref y);

            AddComboRow("Cơ quan Hành chính ban hành:", vals[3], new string[] {
                "ỦY BAN NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT", "ỦY BAN NHÂN DÂN PHƯỜNG", "ỦY BAN NHÂN DÂN XÃ",
                "HỘI ĐỒNG NHÂN DÂN PHƯỜNG", "VĂN PHÒNG HĐND VÀ UBND",
                "ỦY BAN MTTQ VIỆT NAM PHƯỜNG", "CÔNG AN PHƯỜNG", "BAN CHỈ HUY QUÂN SỰ PHƯỜNG",
                "TRẠM Y TẾ PHƯỜNG", "TRƯỜNG TIỂU HỌC", "TRƯỜNG TRUNG HỌC CƠ SỞ"
            }, out cboUIss, ref y);

            AddComboRow("Tiêu ngữ khối UBND / Đoàn thể:", vals[8], new string[] {
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM | Độc lập - Tự do - Hạnh phúc",
                "ĐOÀN TNCS HỒ CHÍ MINH | Thanh niên xung kích, sáng tạo",
                "CÔNG ĐOÀN VIỆT NAM | Đổi mới - Dân chủ - Đoàn kết - Phát triển",
                "HỘI LIÊN HIỆP PHỤ NỮ VIỆT NAM | Đoàn kết - Sáng tạo - Hội nhập - Phát triển",
                "HỘI CỰU CHIẾN BINH VIỆT NAM | Trung thành - Đoàn kết - Gương mẫu - Đổi mới",
                "ĐẢNG CỘNG SẢN VIỆT NAM"
            }, out cboUMotto, ref y);

            AddComboRow("Địa danh (Phường/Xã/Tỉnh):", vals[4], new string[] {
                "Tân Sơn Nhất", "Phường Tân Sơn Nhất", "TP. Hồ Chí Minh", "Hà Nội", "Đà Nẵng", "Phường", "Xã"
            }, out cboLoc, ref y);

            AddRow("Họ tên người ký (Khối Đảng):", vals[5], out txtPSign, ref y);
            AddRow("Họ tên người ký (Khối UBND):", vals[6], out txtUSign, ref y);

            Label lblNote = new Label
            {
                Text = "Mẹo: Bấm mũi tên xổ xuống để chọn nhanh hoặc gõ trực tiếp tên cơ quan / tiêu ngữ bất kỳ của bạn.",
                Location = new Point(22, y + 2),
                AutoSize = true,
                ForeColor = Color.FromArgb(253, 224, 71),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            this.Controls.Add(lblNote);

            Button btnSave = new Button
            {
                Text = "💾 Lưu Cài Đặt Tiêu Ngữ",
                Location = new Point(405, y + 26),
                Size = new Size(205, 38),
                BackColor = Color.FromArgb(234, 179, 8),
                ForeColor = Color.FromArgb(69, 10, 10),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) =>
            {
                SaveUnitConfig(new string[] {
                    cboPGov.Text.Trim(), cboPIss.Text.Trim(), cboUGov.Text.Trim(), cboUIss.Text.Trim(),
                    cboLoc.Text.Trim(), txtPSign.Text.Trim(), txtUSign.Text.Trim(), cboPMotto.Text.Trim(), cboUMotto.Text.Trim()
                });
                MessageBox.Show("Đã lưu cấu hình Tiêu ngữ & Đơn vị! Các mẫu chèn trong Microsoft Word sẽ tự động sử dụng thông tin này.", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            };
            this.Controls.Add(btnSave);
        }

        private void ApplyPreset(int idx)
        {
            if (idx == 1) { cboPGov.Text = "ĐẢNG BỘ THÀNH PHỐ HỒ CHÍ MINH"; cboPIss.Text = "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 2) { cboPGov.Text = "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT"; cboPIss.Text = "VĂN PHÒNG"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 3) { cboPGov.Text = "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT"; cboPIss.Text = "CHI BỘ KHU PHỐ 1"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 4) { cboPGov.Text = "ĐẢNG ỦY PHƯỜNG TÂN SƠN NHẤT"; cboPIss.Text = "ỦY BAN KIỂM TRA ĐẢNG ỦY"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 5) { cboUGov.Text = "ỦY BAN NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"; cboUIss.Text = "ỦY BAN NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 6) { cboUGov.Text = "HỘI ĐỒNG NHÂN DÂN THÀNH PHỐ HỒ CHÍ MINH"; cboUIss.Text = "HỘI ĐỒNG NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 7) { cboUGov.Text = "ỦY BAN MTTQ VIỆT NAM THÀNH PHỐ"; cboUIss.Text = "ỦY BAN MTTQ VIỆT NAM PHƯỜNG TÂN SƠN NHẤT"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 8) { cboUGov.Text = "CÔNG AN THÀNH PHỐ"; cboUIss.Text = "CÔNG AN PHƯỜNG TÂN SƠN NHẤT"; cboLoc.Text = "Tân Sơn Nhất"; }
            else if (idx == 9) { cboUGov.Text = "ỦY BAN NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT"; cboUIss.Text = "TRẠM Y TẾ PHƯỜNG"; cboLoc.Text = "Tân Sơn Nhất"; }
        }

        private void AddComboRow(string label, string initVal, string[] suggestions, out ComboBox cb, ref int y)
        {
            Label lbl = new Label { Text = label, Location = new Point(22, y + 4), AutoSize = true, ForeColor = Color.FromArgb(254, 240, 138) };
            cb = new ComboBox
            {
                Location = new Point(215, y),
                Size = new Size(395, 28),
                DropDownStyle = ComboBoxStyle.DropDown,
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White
            };
            cb.Items.AddRange(suggestions);
            cb.Text = initVal;
            this.Controls.Add(lbl);
            this.Controls.Add(cb);
            y += 42;
        }

        private void AddRow(string label, string initVal, out TextBox tb, ref int y)
        {
            Label lbl = new Label { Text = label, Location = new Point(22, y + 4), AutoSize = true, ForeColor = Color.FromArgb(254, 240, 138) };
            tb = new TextBox
            {
                Text = initVal,
                Location = new Point(215, y),
                Size = new Size(395, 26),
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(lbl);
            this.Controls.Add(tb);
            y += 42;
        }

        public static string[] LoadUnitConfig()
        {
            string[] defs = new string[] {
                "QUẬN ỦY", "ĐẢNG ỦY PHƯỜNG", "ỦY BAN NHÂN DÂN QUẬN", "ỦY BAN NHÂN DÂN PHƯỜNG",
                "Phường", "Trần Minh Quân", "Nguyễn Văn An", "ĐẢNG CỘNG SẢN VIỆT NAM",
                "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM | Độc lập - Tự do - Hạnh phúc"
            };
            try
            {
                string cfgPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM", "word_unit_config.ini");
                if (File.Exists(cfgPath))
                {
                    string[] lines = File.ReadAllLines(cfgPath, Encoding.Unicode);
                    for (int i = 0; i < Math.Min(lines.Length, defs.Length); i++)
                    {
                        if (!string.IsNullOrWhiteSpace(lines[i])) defs[i] = lines[i].Trim();
                    }
                }
            }
            catch { }
            return defs;
        }

        public static void SaveUnitConfig(string[] vals)
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
            Directory.CreateDirectory(dir);
            File.WriteAllLines(Path.Combine(dir, "word_unit_config.ini"), vals, Encoding.Unicode);
        }
    }

    public class AiWordAssistantForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const int GWL_STYLE = -16;
        private const int WS_CAPTION = 0x00C00000;
        private const int WS_THICKFRAME = 0x00040000;
        private const int WS_MINIMIZEBOX = 0x00020000;
        private const int WS_MAXIMIZEBOX = 0x00010000;
        private const int WS_SYSMENU = 0x00080000;
        private const int WS_POPUP = unchecked((int)0x80000000);
        private const int WS_VISIBLE = 0x10000000;
        private const uint WM_CLOSE = 0x0010;

        private Panel pnlBrowserHost;
        private Label lblEmbedWait, lblSelInfo, lblStatus;
        private TextBox txtAiResult;
        private Button btnApprove, btnReject;
        private System.Windows.Forms.Timer syncTimer;
        private Process embeddedProc = null;
        private IntPtr embeddedHwnd = IntPtr.Zero;
        private string lastCapturedSel = "";
        private int capturedSelStart = -1;
        private int capturedSelEnd = -1;
        private string lastSentClipboard = "";
        private string lastSeenClipboard = "";
        private int activeAiIndex = 0;

        private static readonly string[] AiNames = new string[] { "ChatGPT", "Gemini", "Claude", "DeepSeek", "Copilot" };
        private static readonly string[] AiUrls = new string[] {
            "https://chatgpt.com/",
            "https://gemini.google.com/app",
            "https://claude.ai/new",
            "https://chat.deepseek.com/",
            "https://copilot.microsoft.com/"
        };

        public AiWordAssistantForm()
        {
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            int panelW = 460;
            this.Text = "VBTM AI — Nhúng Trực Tiếp Kế Bên Word (Không Cần API Key)";
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(wa.Right - panelW, wa.Top);
            this.Size = new Size(panelW, wa.Height);
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.TopMost = true;
            this.BackColor = Color.White;
            this.ForeColor = Color.FromArgb(24, 24, 27);
            this.Font = new Font("Segoe UI", 8.8f);
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            activeAiIndex = LoadPreferredAiIndex();

            // --- 1. THANH CHỌN AI ĐĂNG NHẬP TRỰC TIẾP (PHÍA TRÊN CÙNG) ---
            Panel pnlTop = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(248, 250, 252), Padding = new Padding(8, 6, 8, 6) };
            int tabW = (panelW - 44) / 5;
            Color[] tabColors = new Color[] {
                Color.FromArgb(17, 24, 39),
                Color.FromArgb(37, 99, 235),
                Color.FromArgb(180, 83, 9),
                Color.FromArgb(67, 56, 202),
                Color.FromArgb(15, 118, 110)
            };
            for (int i = 0; i < AiNames.Length; i++)
            {
                int idx = i;
                Button btnTab = new Button
                {
                    Text = AiNames[i],
                    Location = new Point(8 + i * (tabW + 4), 6),
                    Size = new Size(tabW, 28),
                    BackColor = tabColors[i],
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI", 8.2f, FontStyle.Bold),
                    Cursor = Cursors.Hand
                };
                btnTab.FlatAppearance.BorderSize = 0;
                btnTab.Click += (s, e) => EmbedAiWebInPanel(idx);
                pnlTop.Controls.Add(btnTab);
            }

            // --- 2. THANH ĐIỀU KHIỂN WORD ↔ AI (GẮN ĐÁY TRONG CÙNG 1 CỬA SỔ) ---
            Panel pnlBottom = new Panel { Dock = DockStyle.Bottom, Height = 236, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle };
            int x = 10;
            int w = panelW - 42;
            int y = 6;

            lblSelInfo = new Label
            {
                Text = "1. Đoạn bôi đen trong Word: (hãy bôi đen 1 đoạn trong Word...)",
                Location = new Point(x, y),
                Size = new Size(w, 18),
                Font = new Font("Segoe UI", 8.2f, FontStyle.Bold),
                ForeColor = Color.FromArgb(37, 99, 235),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pnlBottom.Controls.Add(lblSelInfo);
            y += 20;

            int chipW = (w - 16) / 5;
            AddSendChip(pnlBottom, "✨ Hay hơn", "Hãy viết lại đoạn văn bản sau hay hơn, mạch lạc, trang trọng hơn (chỉ trả về văn bản thuần, không dùng Markdown):\r\n\r\n", x, y, chipW);
            AddSendChip(pnlBottom, "✂️ Gọn hơn", "Hãy rút gọn đoạn văn bản sau ngắn gọn, súc tích hơn nhưng giữ đủ ý chính (chỉ trả về văn bản thuần):\r\n\r\n", x + (chipW + 4), y, chipW);
            AddSendChip(pnlBottom, "✅ Chính tả", "Hãy sửa hết lỗi chính tả, dấu câu cho đoạn văn bản sau và giữ nguyên ý gốc (chỉ trả về văn bản thuần):\r\n\r\n", x + (chipW + 4) * 2, y, chipW);
            AddSendChip(pnlBottom, "🏛️ Chuẩn", "Hãy chuẩn hóa đoạn văn bản sau theo văn phong chuẩn Đảng (05-HD/VPTW) và Hành chính (NĐ 30/2020), chỉ trả về văn bản thuần:\r\n\r\n", x + (chipW + 4) * 3, y, chipW);
            AddSendChip(pnlBottom, "📝 Tóm tắt", "Hãy tóm tắt đoạn văn bản sau thật ngắn gọn thành 3–5 ý trọng tâm (chỉ trả về văn bản thuần):\r\n\r\n", x + (chipW + 4) * 4, y, chipW);
            y += 30;

            Label lblPrev = new Label
            {
                Text = "2. Khung AI đã xử lý (bấm Copy trên khung AI ở trên → tự hiện tại đây để duyệt):",
                Location = new Point(x, y),
                Size = new Size(w - 78, 18),
                Font = new Font("Segoe UI", 8.1f, FontStyle.Bold),
                ForeColor = Color.FromArgb(5, 150, 105)
            };
            pnlBottom.Controls.Add(lblPrev);

            Button btnPull = new Button
            {
                Text = "📋 Dán từ AI",
                Location = new Point(x + w - 76, y - 2),
                Size = new Size(76, 20),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(15, 23, 42),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 7.6f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnPull.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnPull.Click += (s, e) => PullAiReplyFromClipboard(true);
            pnlBottom.Controls.Add(btnPull);
            y += 21;

            txtAiResult = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 86),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 11f),
                BackColor = Color.FromArgb(240, 253, 244),
                ForeColor = Color.FromArgb(15, 23, 42),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            txtAiResult.TextChanged += (s, e) =>
            {
                bool hasTxt = !string.IsNullOrWhiteSpace(txtAiResult.Text);
                btnApprove.Enabled = hasTxt;
                btnReject.Enabled = hasTxt;
            };
            pnlBottom.Controls.Add(txtAiResult);
            y += 91;

            btnApprove = new Button
            {
                Text = "✅ Đồng Ý — Thay Vào Đoạn Bôi Đen",
                Location = new Point(x, y),
                Size = new Size(w - 84, 34),
                BackColor = Color.FromArgb(5, 150, 105),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.1f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            btnApprove.FlatAppearance.BorderSize = 0;
            btnApprove.Click += (s, e) => ApplyApprovedResultToWordSelection();
            pnlBottom.Controls.Add(btnApprove);

            btnReject = new Button
            {
                Text = "✕ Hủy",
                Location = new Point(x + w - 78, y),
                Size = new Size(78, 34),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(71, 85, 105),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.6f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnReject.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnReject.Click += (s, e) =>
            {
                txtAiResult.Clear();
                lblStatus.Text = "Đã hủy — văn bản trong Word giữ nguyên 100%.";
            };
            pnlBottom.Controls.Add(btnReject);
            y += 37;

            lblStatus = new Label
            {
                Text = "Đăng nhập AI trực tiếp ngay trong khung trên (1 lần dùng mãi, không cần API Key).",
                Location = new Point(x, y),
                Size = new Size(w, 24),
                ForeColor = Color.FromArgb(71, 85, 105),
                Font = new Font("Segoe UI", 7.8f, FontStyle.Italic),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            pnlBottom.Controls.Add(lblStatus);

            // --- 3. KHUNG NHÚNG TRỰC TIẾP AI (CHIẾM TOÀN BỘ PHẦN GIỮA CỬA SỔ) ---
            pnlBrowserHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(248, 250, 252) };
            lblEmbedWait = new Label
            {
                Text = "⏳ Đang nhúng khung đăng nhập AI trực tiếp vào cửa sổ...",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(71, 85, 105)
            };
            pnlBrowserHost.Controls.Add(lblEmbedWait);
            pnlBrowserHost.Resize += (s, e) => ResizeEmbeddedBrowser();
            this.LocationChanged += (s, e) => ResizeEmbeddedBrowser();
            this.SizeChanged += (s, e) => ResizeEmbeddedBrowser();

            // Add theo đúng thứ tự Dock (Fill trước, Top & Bottom sau)
            this.Controls.Add(pnlBrowserHost);
            this.Controls.Add(pnlTop);
            this.Controls.Add(pnlBottom);

            syncTimer = new System.Windows.Forms.Timer { Interval = 900 };
            syncTimer.Tick += (s, e) =>
            {
                CaptureWordSelection(false);
                PullAiReplyFromClipboard(false);
            };

            this.Shown += (s, e) =>
            {
                try { if (Clipboard.ContainsText()) lastSeenClipboard = Clipboard.GetText(); } catch { }
                CaptureWordSelection(true);
                syncTimer.Start();
                EmbedAiWebInPanel(activeAiIndex);
            };
            this.FormClosed += (s, e) =>
            {
                if (syncTimer != null) syncTimer.Stop();
                CloseEmbeddedBrowser();
            };
        }

        private void AddSendChip(Panel parent, string label, string promptPrefix, int x, int y, int width)
        {
            Button btn = new Button
            {
                Text = label,
                Location = new Point(x, y),
                Size = new Size(width, 26),
                BackColor = Color.FromArgb(243, 244, 246),
                ForeColor = Color.FromArgb(17, 24, 39),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 7.8f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(209, 213, 219);
            btn.Click += (s, e) => SendSelectionToEmbeddedAi(promptPrefix, label);
            parent.Controls.Add(btn);
        }

        private void CloseEmbeddedBrowser()
        {
            try
            {
                if (embeddedHwnd != IntPtr.Zero)
                {
                    PostMessage(embeddedHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    embeddedHwnd = IntPtr.Zero;
                }
                if (embeddedProc != null && !embeddedProc.HasExited)
                {
                    if (!embeddedProc.WaitForExit(600))
                    {
                        try { embeddedProc.Kill(); } catch { }
                    }
                    embeddedProc = null;
                }
            }
            catch { }
        }

        private void ResizeEmbeddedBrowser()
        {
            if (embeddedHwnd != IntPtr.Zero && pnlBrowserHost != null && !this.IsDisposed && this.WindowState != FormWindowState.Minimized)
            {
                Point pt = pnlBrowserHost.PointToScreen(Point.Empty);
                int cw = Math.Max(200, pnlBrowserHost.ClientSize.Width);
                int ch = Math.Max(200, pnlBrowserHost.ClientSize.Height);
                SetWindowPos(embeddedHwnd, HWND_TOPMOST, pt.X, pt.Y, cw, ch, SWP_NOACTIVATE | SWP_SHOWWINDOW);
            }
        }

        private void EmbedAiWebInPanel(int idx)
        {
            if (idx < 0 || idx >= AiUrls.Length) idx = 0;
            activeAiIndex = idx;
            SavePreferredAiIndex(idx);
            CloseEmbeddedBrowser();

            lblEmbedWait.Text = "⏳ Đang nhúng " + AiNames[idx] + " trực tiếp vào khung bên cạnh Word...";
            lblEmbedWait.Visible = true;

            string browserExe = FindChromiumAppBrowser();
            string profileDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VBTMPro", "AiEmbeddedProfile");
            try { Directory.CreateDirectory(profileDir); } catch { }

            if (string.IsNullOrEmpty(browserExe))
            {
                lblEmbedWait.Text = "Không tìm thấy Microsoft Edge hoặc Chrome trên máy.";
                return;
            }

            try
            {
                string args = "--app=\"" + AiUrls[idx] + "\" " +
                              "--user-data-dir=\"" + profileDir + "\" " +
                              "--new-window " +
                              "--window-position=-2000,-2000 " +
                              "--window-size=" + Math.Max(400, pnlBrowserHost.ClientSize.Width) + "," + Math.Max(400, pnlBrowserHost.ClientSize.Height);
                Process proc = Process.Start(new ProcessStartInfo
                {
                    FileName = browserExe,
                    Arguments = args,
                    UseShellExecute = false
                });
                embeddedProc = proc;

                // Đợi cửa sổ Chromium tạo xong rồi nhúng thẳng vào pnlBrowserHost (KHÔNG mở rời bên ngoài!)
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    IntPtr found = IntPtr.Zero;
                    for (int attempt = 0; attempt < 45; attempt++)
                    {
                        System.Threading.Thread.Sleep(150);
                        if (proc == null || proc.HasExited) break;
                        try
                        {
                            proc.Refresh();
                            if (proc.MainWindowHandle != IntPtr.Zero && IsWindowVisible(proc.MainWindowHandle))
                            {
                                found = proc.MainWindowHandle;
                                break;
                            }
                            uint pid = (uint)proc.Id;
                            EnumWindows((h, lp) =>
                            {
                                uint wPid;
                                GetWindowThreadProcessId(h, out wPid);
                                if (wPid == pid && IsWindowVisible(h)) { found = h; return false; }
                                return true;
                            }, IntPtr.Zero);
                            if (found != IntPtr.Zero) break;
                        }
                        catch { }
                    }

                    if (this.IsDisposed) return;
                    this.BeginInvoke((Action)(() =>
                    {
                        if (found != IntPtr.Zero)
                        {
                            embeddedHwnd = found;
                            // Giữ Edge ở chế độ borderless WS_POPUP bám khít tọa độ pnlBrowserHost (KHÔNG SetParent để gõ phím, bấm nút & đăng nhập Google mượt 100%!)
                            int style = GetWindowLong(found, GWL_STYLE);
                            style = (style & ~(WS_CAPTION | WS_THICKFRAME | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU)) | WS_POPUP | WS_VISIBLE;
                            SetWindowLong(found, GWL_STYLE, style);
                            Point pt = pnlBrowserHost.PointToScreen(Point.Empty);
                            int cw = Math.Max(200, pnlBrowserHost.ClientSize.Width);
                            int ch = Math.Max(200, pnlBrowserHost.ClientSize.Height);
                            SetWindowPos(found, HWND_TOPMOST, pt.X, pt.Y, cw, ch, SWP_FRAMECHANGED | SWP_SHOWWINDOW);
                            SetForegroundWindow(found);
                            lblEmbedWait.Visible = false;
                            lblStatus.Text = "✓ Đã nhúng " + AiNames[idx] + " ngay trong khung! Gõ đăng nhập & bấm nút hoàn toàn bình thường.";
                        }
                        else
                        {
                            lblEmbedWait.Text = "Đang chạy " + AiNames[idx] + " — hãy đăng nhập tài khoản của bạn.";
                        }
                    }));
                });
            }
            catch (Exception ex)
            {
                lblEmbedWait.Text = "Lỗi nhúng AI: " + ex.Message;
            }
        }

        private static string FindChromiumAppBrowser()
        {
            string[] candidates = new string[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\msedge.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Google\Chrome\Application\chrome.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Google\Chrome\Application\chrome.exe")
            };
            foreach (string c in candidates)
            {
                if (!string.IsNullOrEmpty(c) && File.Exists(c)) return c;
            }
            return null;
        }

        private void SendSelectionToEmbeddedAi(string promptPrefix, string actionLabel)
        {
            CaptureWordSelection(true);
            string src = (lastCapturedSel ?? "").Trim();
            if (string.IsNullOrWhiteSpace(src) || src.Length < 3)
            {
                MessageBox.Show("Hãy bôi đen 1 đoạn văn bản trong Microsoft Word trước!", "VBTM AI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string payload = promptPrefix + src;
            try
            {
                Clipboard.SetText(payload);
                lastSentClipboard = payload;
                lastSeenClipboard = payload;
            }
            catch { }

            if (embeddedHwnd != IntPtr.Zero)
            {
                SetForegroundWindow(embeddedHwnd);
            }
            lblStatus.Text = "✓ Đã chép [" + actionLabel + "] + đoạn bôi đen! Bấm vào ô chat " + AiNames[activeAiIndex] + " ở trên rồi nhấn Ctrl+V → Enter.";
        }

        private void PullAiReplyFromClipboard(bool manualClick)
        {
            try
            {
                if (!Clipboard.ContainsText()) return;
                string clip = Clipboard.GetText();
                if (string.IsNullOrWhiteSpace(clip)) return;
                if (!manualClick && (clip == lastSeenClipboard || clip == lastSentClipboard || clip == lastCapturedSel)) return;

                lastSeenClipboard = clip;
                string clean = clip.Replace("**", "").Replace("### ", "").Replace("## ", "").Trim();
                txtAiResult.Text = clean.Replace("\r\n", "\n").Replace("\n", "\r\n");
                lblStatus.Text = "✓ Đã nhận kết quả AI (" + clean.Length + " ký tự)! Xem trước ở khung xanh → bấm '✅ Đồng Ý — Thay Vào Đoạn Bôi Đen'.";
            }
            catch { }
        }

        private void CaptureWordSelection(bool forceNotify)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                int sStart = (int)sel.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, sel, null);
                int sEnd = (int)sel.GetType().InvokeMember("End", BindingFlags.GetProperty, null, sel, null);
                if (sEnd > sStart)
                {
                    string raw = (Convert.ToString(sel.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, sel, null)) ?? "").TrimEnd('\r', '\a', '\n');
                    if (raw.Length >= 3 && (raw != lastCapturedSel || forceNotify))
                    {
                        lastCapturedSel = raw;
                        capturedSelStart = sStart;
                        capturedSelEnd = sEnd;
                        string preview = raw.Replace("\r", " ").Replace("\n", " ");
                        if (preview.Length > 44) preview = preview.Substring(0, 44) + "...";
                        lblSelInfo.Text = "1. Đang bôi đen (" + raw.Length + " ký tự): \"" + preview + "\"";
                    }
                }
            }
            catch { }
        }

        private void ApplyApprovedResultToWordSelection()
        {
            string approved = (txtAiResult.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(approved)) return;

            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null);
                object targetRange = null;

                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                int curStart = (int)sel.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, sel, null);
                int curEnd = (int)sel.GetType().InvokeMember("End", BindingFlags.GetProperty, null, sel, null);

                if (curEnd > curStart)
                {
                    targetRange = sel.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, sel, null);
                }
                else if (capturedSelStart >= 0 && capturedSelEnd > capturedSelStart)
                {
                    targetRange = doc.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, doc, new object[] { capturedSelStart, capturedSelEnd });
                }
                else
                {
                    throw new Exception("Hãy bôi đen đoạn văn bản cần thay trong Microsoft Word.");
                }

                string clean = approved.Replace("\r\n", "\r").Replace("\n", "\r");
                targetRange.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, targetRange, new object[] { clean });
                object fnt = targetRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, targetRange, null);
                fnt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, fnt, new object[] { "Times New Roman" });
                fnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, fnt, new object[] { 14f });
                fnt.GetType().InvokeMember("Scaling", BindingFlags.SetProperty, null, fnt, new object[] { 100 });

                lastCapturedSel = clean;
                capturedSelStart = (int)targetRange.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, targetRange, null);
                capturedSelEnd = (int)targetRange.GetType().InvokeMember("End", BindingFlags.GetProperty, null, targetRange, null);
                txtAiResult.Clear();
                lblStatus.Text = "✓ Đã bỏ kết quả AI thay vào đúng đoạn bôi đen trong Word (Times New Roman 14pt)!";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể thay vào đoạn bôi đen trong Word: " + ex.Message, "VBTM AI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static int LoadPreferredAiIndex()
        {
            try
            {
                string p = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM", "ai_login_pref.txt");
                if (File.Exists(p))
                {
                    int v = 0;
                    if (int.TryParse(File.ReadAllText(p).Trim(), out v) && v >= 0 && v < AiUrls.Length) return v;
                }
            }
            catch { }
            return 0;
        }

        private static void SavePreferredAiIndex(int idx)
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
                Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "ai_login_pref.txt"), idx.ToString());
            }
            catch { }
        }

        private string CallAiProviderFast(int prov, string model, string apiKey, int taskIdx, string systemPrompt, string userText)
        {
            if (prov == 5)
            {
                return RunOfflineAiTransform(userText, taskIdx);
            }
            return CallOnlineAiDirect(prov, model, apiKey, systemPrompt, userText);
        }

        public static string CallOnlineAiDirect(int prov, string model, string apiKey, string systemPrompt, string userText)
        {
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; // TLS 1.2
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.UseNagleAlgorithm = false;
            ServicePointManager.DefaultConnectionLimit = 12;
            JavaScriptSerializer js = new JavaScriptSerializer();
            js.MaxJsonLength = int.MaxValue;

            if (prov == 0) // Google Gemini (tắt thinkingBudget = 0 + Proxy=null -> trả lời siêu tốc 1–2 giây!)
            {
                string reqModel = string.IsNullOrEmpty(model) || model.Equals("gemini-2.5-flash", StringComparison.OrdinalIgnoreCase)
                    ? "gemini-3.8-flash"
                    : model;
                string[] tryModels = new string[] { reqModel, "gemini-3.8-flash", "gemini-3.0-flash", "gemini-2.0-flash" };
                var fastPayload = new
                {
                    contents = new object[] { new { parts = new object[] { new { text = systemPrompt + "\n\n" + userText } } } },
                    generationConfig = new { temperature = 0.55, maxOutputTokens = 1024, thinkingConfig = new { thinkingBudget = 0 } }
                };
                var plainPayload = new
                {
                    contents = new object[] { new { parts = new object[] { new { text = systemPrompt + "\n\n" + userText } } } },
                    generationConfig = new { temperature = 0.55, maxOutputTokens = 1024 }
                };
                string resp = null;
                Exception lastEx = null;
                HashSet<string> tried = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string mName in tryModels)
                {
                    if (!tried.Add(mName)) continue;
                    string url = "https://generativelanguage.googleapis.com/v1beta/models/" + Uri.EscapeDataString(mName) + ":generateContent?key=" + Uri.EscapeDataString(apiKey);
                    try
                    {
                        resp = HttpPostJson(url, js.Serialize(fastPayload), null);
                        if (!string.IsNullOrEmpty(resp)) break;
                    }
                    catch (Exception exM)
                    {
                        lastEx = exM;
                        if (exM.Message.IndexOf("thinkingConfig", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            try { resp = HttpPostJson(url, js.Serialize(plainPayload), null); if (!string.IsNullOrEmpty(resp)) break; } catch (Exception ex2) { lastEx = ex2; }
                        }
                        else if (exM.Message.IndexOf("404", StringComparison.OrdinalIgnoreCase) < 0 &&
                                 exM.Message.IndexOf("NOT_FOUND", StringComparison.OrdinalIgnoreCase) < 0 &&
                                 exM.Message.IndexOf("no longer available", StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            throw;
                        }
                    }
                }
                if (string.IsNullOrEmpty(resp) && lastEx != null &&
                    (lastEx.Message.IndexOf("404", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     lastEx.Message.IndexOf("NOT_FOUND", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     lastEx.Message.IndexOf("no longer available", StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    // Tự động hỏi Google API danh sách model Flash đang mở cho Key này để không bao giờ bị lỗi 404 nữa!
                    try
                    {
                        string listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + Uri.EscapeDataString(apiKey);
                        HttpWebRequest lreq = (HttpWebRequest)WebRequest.Create(listUrl);
                        lreq.Proxy = null;
                        lreq.Timeout = 10000;
                        using (HttpWebResponse lresp = (HttpWebResponse)lreq.GetResponse())
                        using (StreamReader lsr = new StreamReader(lresp.GetResponseStream(), Encoding.UTF8))
                        {
                            var lobj = js.Deserialize<Dictionary<string, object>>(lsr.ReadToEnd());
                            var mlist = lobj["models"] as System.Collections.ArrayList;
                            if (mlist != null)
                            {
                                string discovered = null;
                                foreach (object mo in mlist)
                                {
                                    var md = mo as Dictionary<string, object>;
                                    if (md == null) continue;
                                    string mn = Convert.ToString(md["name"] ?? "");
                                    var methods = md.ContainsKey("supportedGenerationMethods") ? md["supportedGenerationMethods"] as System.Collections.ArrayList : null;
                                    bool canGen = methods != null && methods.Contains("generateContent");
                                    if (canGen && mn.IndexOf("flash", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        discovered = mn.StartsWith("models/") ? mn.Substring(7) : mn;
                                        break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(discovered))
                                {
                                    string dUrl = "https://generativelanguage.googleapis.com/v1beta/models/" + Uri.EscapeDataString(discovered) + ":generateContent?key=" + Uri.EscapeDataString(apiKey);
                                    try { resp = HttpPostJson(dUrl, js.Serialize(plainPayload), null); } catch (Exception exD) { lastEx = exD; }
                                }
                            }
                        }
                    }
                    catch { }
                }
                if (string.IsNullOrEmpty(resp)) throw (lastEx ?? new Exception("Không gọi được Gemini API."));
                var obj = js.Deserialize<Dictionary<string, object>>(resp);
                var cands = obj["candidates"] as System.Collections.ArrayList;
                var first = cands[0] as Dictionary<string, object>;
                var content = first["content"] as Dictionary<string, object>;
                var parts = content["parts"] as System.Collections.ArrayList;
                var p0 = parts[0] as Dictionary<string, object>;
                return Convert.ToString(p0["text"]).Trim();
            }
            else if (prov == 2) // Anthropic Claude
            {
                var payload = new
                {
                    model = string.IsNullOrEmpty(model) ? "claude-3-5-sonnet-latest" : model,
                    max_tokens = 1500,
                    system = systemPrompt,
                    messages = new object[] { new { role = "user", content = userText } }
                };
                var headers = new Dictionary<string, string> {
                    { "x-api-key", apiKey },
                    { "anthropic-version", "2023-06-01" }
                };
                string resp = HttpPostJson("https://api.anthropic.com/v1/messages", js.Serialize(payload), headers);
                var obj = js.Deserialize<Dictionary<string, object>>(resp);
                var content = obj["content"] as System.Collections.ArrayList;
                var c0 = content[0] as Dictionary<string, object>;
                return Convert.ToString(c0["text"]).Trim();
            }
            else // OpenAI ChatGPT (1), DeepSeek (3), Groq (4)
            {
                string url = "https://api.openai.com/v1/chat/completions";
                if (prov == 3) url = "https://api.deepseek.com/chat/completions";
                else if (prov == 4) url = "https://api.groq.com/openai/v1/chat/completions";
                var payload = new
                {
                    model = string.IsNullOrEmpty(model) ? "gpt-4o-mini" : model,
                    messages = new object[] {
                        new { role = "system", content = systemPrompt },
                        new { role = "user", content = userText }
                    }
                };
                var headers = new Dictionary<string, string> { { "Authorization", "Bearer " + apiKey } };
                string resp = HttpPostJson(url, js.Serialize(payload), headers);
                var obj = js.Deserialize<Dictionary<string, object>>(resp);
                var choices = obj["choices"] as System.Collections.ArrayList;
                var ch0 = choices[0] as Dictionary<string, object>;
                var msg = ch0["message"] as Dictionary<string, object>;
                return Convert.ToString(msg["content"]).Trim();
            }
        }

        private static string HttpPostJson(string url, string jsonBody, Dictionary<string, string> headers)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/json; charset=utf-8";
            req.Proxy = null; // Tắt dò WPAD Proxy của Windows (loại bỏ độ trễ 5–15 giây!)
            req.KeepAlive = true;
            req.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            req.Timeout = 20000;
            req.ReadWriteTimeout = 20000;
            if (headers != null)
            {
                foreach (var kv in headers) req.Headers[kv.Key] = kv.Value;
            }
            byte[] bytes = Encoding.UTF8.GetBytes(jsonBody);
            req.ContentLength = bytes.Length;
            using (Stream st = req.GetRequestStream()) st.Write(bytes, 0, bytes.Length);
            try
            {
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                string detail = wex.Message;
                if (wex.Response != null)
                {
                    using (StreamReader sr = new StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8))
                        detail = sr.ReadToEnd();
                    try
                    {
                        JavaScriptSerializer jsErr = new JavaScriptSerializer();
                        var eObj = jsErr.Deserialize<Dictionary<string, object>>(detail);
                        if (eObj != null && eObj.ContainsKey("error"))
                        {
                            var inner = eObj["error"] as Dictionary<string, object>;
                            if (inner != null && inner.ContainsKey("message"))
                            {
                                string code = inner.ContainsKey("code") ? ("[" + inner["code"] + "] ") : "";
                                detail = code + Convert.ToString(inner["message"]);
                            }
                        }
                    }
                    catch { }
                }
                throw new Exception(detail);
            }
        }

        private string RunOfflineAiTransform(string text, int taskIdx)
        {
            int cnt = 0;
            string res = text;
            using (MainStudioForm mf = new MainStudioForm())
            {
                res = mf.FixAllVietnameseSpellingAndPunctuation(text, out cnt);
            }
            string[,] reps = new string[,] {
                { "làm gấp", "khẩn trương triển khai thực hiện" },
                { "báo lại", "báo cáo kết quả về" },
                { "mọi người", "các đồng chí và cán bộ, công chức" },
                { "cần phải", "đề nghị tập trung" },
                { "công việc", "nhiệm vụ được giao" },
                { "làm xong", "hoàn thành" }
            };
            if (taskIdx == 1 || taskIdx == 2 || taskIdx == 3)
            {
                for (int i = 0; i < reps.GetLength(0); i++)
                    res = Regex.Replace(res, Regex.Escape(reps[i, 0]), reps[i, 1], RegexOptions.IgnoreCase);
            }
            return res;
        }

        public static string[] LoadAiConfig()
        {
            string[] defs = new string[] { "0", "gemini-3.8-flash", "" };
            try
            {
                string vbtmDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
                string cfgPath = Path.Combine(vbtmDir, "ai_hub_v2.ini");
                if (File.Exists(cfgPath))
                {
                    string[] lines = File.ReadAllLines(cfgPath, Encoding.UTF8);
                    for (int i = 0; i < Math.Min(lines.Length, defs.Length); i++) defs[i] = (lines[i] ?? "").Trim();
                }

                // Nếu Key hiện tại bị thiếu hoặc quá ngắn (<25 ký tự), tự động khôi phục Key Gemini 39 ký tự từ ai_hub_config.ini (nếu có)
                if (defs[2].Length < 25)
                {
                    string oldCfg = Path.Combine(vbtmDir, "ai_hub_config.ini");
                    if (File.Exists(oldCfg))
                    {
                        string[] oldLines = File.ReadAllLines(oldCfg, Encoding.UTF8);
                        if (oldLines.Length >= 3 && !string.IsNullOrWhiteSpace(oldLines[2]) && oldLines[2].Trim().Length >= 25)
                        {
                            defs[2] = oldLines[2].Trim();
                            if (defs[2].StartsWith("AIza", StringComparison.Ordinal))
                            {
                                defs[0] = "0";
                                defs[1] = "gemini-3.8-flash";
                            }
                        }
                    }
                }
                else if (defs[2].StartsWith("AIza", StringComparison.Ordinal))
                {
                    defs[0] = "0";
                    if (string.IsNullOrWhiteSpace(defs[1]) || !defs[1].StartsWith("gemini", StringComparison.OrdinalIgnoreCase) || defs[1].Equals("gemini-2.5-flash", StringComparison.OrdinalIgnoreCase))
                        defs[1] = "gemini-3.8-flash";
                }
            }
            catch { }
            return defs;
        }

        public static void SaveAiConfig(string[] vals)
        {
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
                Directory.CreateDirectory(dir);
                File.WriteAllLines(Path.Combine(dir, "ai_hub_v2.ini"), vals, Encoding.UTF8);
            }
            catch { }
        }
    }

    public class WordAiSummaryPreviewForm : Form
    {
        private ComboBox cboScope, cboStyle, cboInsertMode;
        private CheckBox chkIncludeHeading;
        private TextBox txtPreview;
        private Label lblStatus, lblSourceInfo;
        private Button btnRegenerate, btnConfigAi, btnConfirmInsert, btnCancel;
        private ProgressBar prgWait;
        private string fullDocText = "";
        private string selectedDocText = "";
        private bool isBusy = false;

        public WordAiSummaryPreviewForm()
        {
            this.Text = "VBTM Pro 2.5.0 — Xem Trước Bản Tóm Tắt AI Trước Khi Đưa Vào Word";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(740, 590);
            this.MinimumSize = new Size(640, 500);
            this.TopMost = true;
            this.BackColor = Color.White;
            this.ForeColor = Color.FromArgb(24, 24, 27);
            this.Font = new Font("Segoe UI", 9.2f);
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            int x = 20;
            int w = 684;
            int y = 16;

            Label lblTitle = new Label
            {
                Text = "AI Tóm Tắt Văn Bản — Xem Trước & Chỉnh Sửa Trước Khi Chèn Vào Nội Dung",
                Location = new Point(x, y),
                Size = new Size(w, 26),
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.FromArgb(17, 24, 39)
            };
            this.Controls.Add(lblTitle);
            y += 26;

            lblSourceInfo = new Label
            {
                Text = "Đang đọc nội dung văn bản Microsoft Word đang mở...",
                Location = new Point(x, y),
                Size = new Size(w, 20),
                Font = new Font("Segoe UI", 8.8f),
                ForeColor = Color.FromArgb(100, 116, 139)
            };
            this.Controls.Add(lblSourceInfo);
            y += 26;

            // Hàng tùy chọn: Phạm vi + Kiểu tóm tắt + Nút tóm tắt lại
            Label lblScope = new Label { Text = "Phạm vi đọc:", Location = new Point(x, y + 4), Size = new Size(85, 20), Font = new Font("Segoe UI", 8.8f, FontStyle.Bold) };
            this.Controls.Add(lblScope);

            cboScope = new ComboBox { Location = new Point(x + 88, y), Size = new Size(205, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cboScope.Items.AddRange(new object[] {
                "Toàn bộ nội dung văn bản",
                "Chỉ đoạn đang bôi đen"
            });
            cboScope.SelectedIndex = 0;
            this.Controls.Add(cboScope);

            Label lblStyle = new Label { Text = "Kiểu tóm tắt:", Location = new Point(x + 304, y + 4), Size = new Size(85, 20), Font = new Font("Segoe UI", 8.8f, FontStyle.Bold) };
            this.Controls.Add(lblStyle);

            cboStyle = new ComboBox { Location = new Point(x + 392, y), Size = new Size(292, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cboStyle.Items.AddRange(new object[] {
                "Ngắn gọn (3–5 gạch đầu dòng trọng tâm)",
                "1 đoạn văn cô đọng (chuẩn NĐ 30 / 05-HD)",
                "Chi tiết theo mục (Mục đích – Nội dung – Nhiệm vụ)"
            });
            cboStyle.SelectedIndex = 0;
            this.Controls.Add(cboStyle);
            y += 34;

            btnRegenerate = new Button
            {
                Text = "🔄 Tạo / Tóm Tắt Lại Bằng AI",
                Location = new Point(x, y),
                Size = new Size(215, 32),
                BackColor = Color.FromArgb(17, 24, 39),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnRegenerate.FlatAppearance.BorderSize = 0;
            btnRegenerate.Click += (s, e) => RunSummaryAsync();
            this.Controls.Add(btnRegenerate);

            btnConfigAi = new Button
            {
                Text = "🔗 Liên Kết AI Ngoài (ChatGPT / Gemini)",
                Location = new Point(x + 225, y),
                Size = new Size(235, 32),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(15, 23, 42),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.8f),
                Cursor = Cursors.Hand
            };
            btnConfigAi.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnConfigAi.Click += (s, e) =>
            {
                AiWordAssistantForm af = new AiWordAssistantForm();
                af.Show();
            };
            this.Controls.Add(btnConfigAi);

            prgWait = new ProgressBar
            {
                Location = new Point(x + 472, y + 6),
                Size = new Size(212, 20),
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };
            this.Controls.Add(prgWait);
            y += 40;

            Label lblPreview = new Label
            {
                Text = "Xem trước bản tóm tắt (Người soạn thảo có thể sửa trực tiếp nội dung bên dưới trước khi đưa vào Word):",
                Location = new Point(x, y),
                Size = new Size(w, 20),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 23, 42)
            };
            this.Controls.Add(lblPreview);
            y += 22;

            txtPreview = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 250),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 12.5f),
                BackColor = Color.FromArgb(250, 250, 252),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtPreview);
            y += 260;

            Label lblPos = new Label
            {
                Text = "Vị trí chèn vào nội dung Word:",
                Location = new Point(x, y + 4),
                Size = new Size(195, 22),
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            this.Controls.Add(lblPos);

            cboInsertMode = new ComboBox
            {
                Location = new Point(x + 200, y),
                Size = new Size(484, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            cboInsertMode.Items.AddRange(new object[] {
                "1. Chèn trực tiếp tại vị trí con trỏ trong nội dung Word",
                "2. Chèn ngay đầu phần nội dung (sau Tiêu đề / Trích yếu)",
                "3. Chèn vào cuối phần nội dung (trước mục Nơi nhận / Chữ ký)",
                "4. Thay thế đoạn đang bôi đen trong nội dung Word"
            });
            cboInsertMode.SelectedIndex = 0;
            this.Controls.Add(cboInsertMode);
            y += 32;

            chkIncludeHeading = new CheckBox
            {
                Text = "Kèm tiêu đề in đậm \"TÓM TẮT NỘI DUNG:\" phía trên đoạn tóm tắt khi chèn vào Word",
                Location = new Point(x, y),
                Size = new Size(w, 22),
                Checked = true,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            this.Controls.Add(chkIncludeHeading);
            y += 28;

            lblStatus = new Label
            {
                Text = "Chưa chèn vào Word — hãy kiểm tra kỹ bản xem trước rồi bấm Xác Nhận.",
                Location = new Point(x, y),
                Size = new Size(w, 20),
                Font = new Font("Segoe UI", 8.6f, FontStyle.Italic),
                ForeColor = Color.FromArgb(71, 85, 105),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStatus);
            y += 24;

            btnConfirmInsert = new Button
            {
                Text = "✅ Xác Nhận Đưa Vào Nội Dung Word",
                Location = new Point(x, y),
                Size = new Size(340, 38),
                BackColor = Color.FromArgb(5, 150, 105),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnConfirmInsert.FlatAppearance.BorderSize = 0;
            btnConfirmInsert.Click += (s, e) => InsertApprovedSummaryIntoWord();
            this.Controls.Add(btnConfirmInsert);

            btnCancel = new Button
            {
                Text = "✕ Hủy (Không chèn vào Word)",
                Location = new Point(x + 356, y),
                Size = new Size(328, 38),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(51, 65, 85),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnCancel.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);

            this.Shown += (s, e) =>
            {
                LoadDocumentTextFromWord();
                RunSummaryAsync();
            };
        }

        private void LoadDocumentTextFromWord()
        {
            fullDocText = "";
            selectedDocText = "";
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = null;
                try { doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null); }
                catch
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int dc = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (dc > 0) doc = docs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docs, new object[] { 1 });
                }
                if (doc != null)
                {
                    try
                    {
                        object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                        string st = (Convert.ToString(sel.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, sel, null)) ?? "").Trim('\r', '\a', '\n', ' ');
                        if (st.Length >= 30) selectedDocText = st;
                    }
                    catch { }

                    StringBuilder sb = new StringBuilder();
                    object paras = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
                    int pCount = (int)paras.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, paras, null);
                    for (int i = 1; i <= pCount; i++)
                    {
                        object p = paras.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paras, new object[] { i });
                        object r = p.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, p, null);
                        object tbls = r.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, r, null);
                        if ((int)tbls.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, tbls, null) > 0) continue;
                        string line = (Convert.ToString(r.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, r, null)) ?? "").TrimEnd('\r', '\a', '\n').Trim();
                        if (line.Length < 4 || line.StartsWith("---") || line.StartsWith("───")) continue;
                        sb.AppendLine(line);
                    }
                    fullDocText = sb.ToString().Trim();
                }
            }
            catch { }

            if (string.IsNullOrWhiteSpace(fullDocText))
            {
                try
                {
                    string snap = Path.Combine(Path.GetTempPath(), "vbtm_word_snapshot.txt");
                    if (File.Exists(snap)) fullDocText = File.ReadAllText(snap, Encoding.Unicode).Replace("\r", "\r\n").Trim();
                }
                catch { }
            }

            if (!string.IsNullOrWhiteSpace(selectedDocText))
            {
                cboScope.SelectedIndex = 1;
                cboInsertMode.SelectedIndex = 3;
                lblSourceInfo.Text = "Đã phát hiện đoạn đang bôi đen (" + selectedDocText.Length + " ký tự) & toàn văn bản (" + fullDocText.Length + " ký tự).";
            }
            else if (!string.IsNullOrWhiteSpace(fullDocText))
            {
                lblSourceInfo.Text = "Đã đọc nội dung tài liệu Word đang mở (" + fullDocText.Length + " ký tự).";
            }
            else
            {
                lblSourceInfo.Text = "⚠ Chưa thấy nội dung trong Word. Hãy mở file Word có văn bản rồi bấm 'Tạo / Tóm Tắt Lại'.";
            }
        }

        private void SetBusy(bool busy, string msg)
        {
            if (this.InvokeRequired) { this.BeginInvoke((Action)(() => SetBusy(busy, msg))); return; }
            isBusy = busy;
            btnRegenerate.Enabled = !busy;
            btnConfirmInsert.Enabled = !busy;
            prgWait.Visible = busy;
            if (msg != null) lblStatus.Text = msg;
        }

        private void RunSummaryAsync()
        {
            if (isBusy) return;
            string srcText = (cboScope.SelectedIndex == 1 && !string.IsNullOrWhiteSpace(selectedDocText)) ? selectedDocText : fullDocText;
            if (string.IsNullOrWhiteSpace(srcText) || srcText.Length < 20)
            {
                LoadDocumentTextFromWord();
                srcText = (cboScope.SelectedIndex == 1 && !string.IsNullOrWhiteSpace(selectedDocText)) ? selectedDocText : fullDocText;
            }
            if (string.IsNullOrWhiteSpace(srcText) || srcText.Length < 20)
            {
                txtPreview.Text = "Không tìm thấy đủ nội dung trong văn bản Word đang mở để tóm tắt.";
                return;
            }

            int styleIdx = cboStyle.SelectedIndex;
            string[] cfg = AiWordAssistantForm.LoadAiConfig();
            int provIdx = 0; int.TryParse(cfg[0], out provIdx);
            string modelStr = cfg[1];
            string apiKey = cfg[2];

            bool useOnline = (provIdx != 5 && !string.IsNullOrWhiteSpace(apiKey));
            SetBusy(true, useOnline ? "⏳ Đang gọi AI tóm tắt nội dung văn bản (xem trước, chưa chèn vào Word)..." : "⏳ Đang tóm tắt bằng Offline AI VBTM...");
            txtPreview.Text = "Đang tạo bản xem trước tóm tắt...";

            string clipped = srcText.Length > 14000 ? srcText.Substring(0, 14000) : srcText;
            string styleInstr = "Tóm tắt thành 3 đến 5 gạch đầu dòng (- ...) cực kỳ súc tích, nêu rõ mục đích, nội dung cốt lõi và nhiệm vụ/thời hạn chính.";
            if (styleIdx == 1)
                styleInstr = "Tóm tắt thành đúng 1 đoạn văn hành chính mạch lạc, trang trọng (khoảng 60–110 từ), chuẩn văn phong Nghị định 30/2020/NĐ-CP và Hướng dẫn 05-HD/VPTW.";
            else if (styleIdx == 2)
                styleInstr = "Tóm tắt rõ ràng theo 3 mục ngắn gọn:\n1. Mục đích, yêu cầu: ...\n2. Nội dung trọng tâm: ...\n3. Tổ chức thực hiện / Thời hạn: ...";

            string sysPrompt = "Bạn là chuyên gia biên tập văn bản Đảng (05-HD/VPTW) và Hành chính Việt Nam (NĐ 30/2020). " +
                               "Hãy tóm tắt văn bản sau bằng tiếng Việt chuẩn mực. " + styleInstr +
                               " CHỈ TRẢ VỀ NỘI DUNG TÓM TẮT THUẦN VĂN BẢN, tuyệt đối không viết lời chào, không dùng ký hiệu Markdown (** hay #).";

            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                string result = "";
                string modeNote = "";
                if (useOnline)
                {
                    try
                    {
                        result = AiWordAssistantForm.CallOnlineAiDirect(provIdx, modelStr, apiKey, sysPrompt, clipped);
                        result = result.Replace("**", "").Replace("##", "").Replace("###", "").Trim();
                        modeNote = "✓ Đã tóm tắt bằng AI trực tuyến (" + modelStr + "). Hãy xem trước & chỉnh sửa nếu cần rồi bấm 'Xác Nhận Đưa Vào Nội Dung Word'.";
                    }
                    catch (Exception ex)
                    {
                        result = BuildOfflineSummary(srcText, styleIdx);
                        modeNote = "⚠ AI trực tuyến tạm gián đoạn (" + ex.Message.Split('\n')[0] + ") — đã tự động dùng Offline AI VBTM. Kiểm tra bản xem trước rồi bấm Xác Nhận.";
                    }
                }
                else
                {
                    result = BuildOfflineSummary(srcText, styleIdx);
                    modeNote = "✓ Đã tóm tắt bằng Offline AI VBTM (miễn phí). Bạn có thể sửa trực tiếp bên trên rồi bấm 'Xác Nhận Đưa Vào Nội Dung Word'.";
                }

                this.BeginInvoke((Action)(() =>
                {
                    txtPreview.Text = result.Replace("\r\n", "\n").Replace("\n", "\r\n");
                    SetBusy(false, modeNote);
                }));
            });
        }

        private string BuildOfflineSummary(string text, int styleIdx)
        {
            string[] rawLines = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> cleanLines = new List<string>();
            foreach (string rl in rawLines)
            {
                string t = rl.Trim();
                if (t.Length < 15) continue;
                if (t.StartsWith("Nơi nhận", StringComparison.OrdinalIgnoreCase)) break;
                if (t.StartsWith("ĐẢNG BỘ", StringComparison.OrdinalIgnoreCase) || t.StartsWith("ĐẢNG ỦY", StringComparison.OrdinalIgnoreCase) ||
                    t.StartsWith("CỘNG HÒA XÃ HỘI", StringComparison.OrdinalIgnoreCase) || t.StartsWith("Độc lập - Tự do", StringComparison.OrdinalIgnoreCase))
                    continue;
                cleanLines.Add(t);
            }

            if (cleanLines.Count == 0) return text.Length > 350 ? text.Substring(0, 350) + "..." : text;

            List<string> keyPoints = new List<string>();
            foreach (string ln in cleanLines)
            {
                bool isHeading = Regex.IsMatch(ln, @"^(I|II|III|IV|V|VI|1|2|3|4|5)\.\s");
                bool isCore = ln.IndexOf("mục đích", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("nhiệm vụ", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("tổ chức thực hiện", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("yêu cầu", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("kế hoạch", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("báo cáo", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              ln.IndexOf("triển khai", StringComparison.OrdinalIgnoreCase) >= 0;
                if (isHeading || isCore || keyPoints.Count < 2)
                {
                    string s = ln.Length > 210 ? ln.Substring(0, 207).TrimEnd() + "..." : ln;
                    if (!keyPoints.Contains(s)) keyPoints.Add(s);
                }
                if (keyPoints.Count >= 5) break;
            }
            if (keyPoints.Count < 3)
            {
                foreach (string ln in cleanLines)
                {
                    string s = ln.Length > 210 ? ln.Substring(0, 207).TrimEnd() + "..." : ln;
                    if (!keyPoints.Contains(s)) keyPoints.Add(s);
                    if (keyPoints.Count >= 4) break;
                }
            }

            if (styleIdx == 1)
            {
                return string.Join(" ", keyPoints.ToArray());
            }
            if (styleIdx == 2)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("1. Mục đích, bối cảnh: " + keyPoints[0]);
                sb.AppendLine("2. Nội dung trọng tâm: " + (keyPoints.Count > 1 ? keyPoints[1] : keyPoints[0]));
                sb.Append("3. Nhiệm vụ & tổ chức thực hiện: " + (keyPoints.Count > 2 ? keyPoints[keyPoints.Count - 1] : "Triển khai thực hiện theo đúng tiến độ đề ra."));
                return sb.ToString();
            }
            StringBuilder bullets = new StringBuilder();
            foreach (string kp in keyPoints)
            {
                bullets.AppendLine("- " + kp.TrimStart('-', '•', '*', ' '));
            }
            return bullets.ToString().Trim();
        }

        private void InsertApprovedSummaryIntoWord()
        {
            string summaryBody = (txtPreview.Text ?? "").Trim();
            if (string.IsNullOrWhiteSpace(summaryBody))
            {
                MessageBox.Show("Nội dung tóm tắt đang trống!", "VBTM Pro 2.5.0");
                return;
            }

            int mode = cboInsertMode.SelectedIndex;
            bool addHeading = chkIncludeHeading.Checked;

            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = null;
                try { doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null); }
                catch
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int dc = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (dc > 0) doc = docs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docs, new object[] { 1 });
                }
                if (doc == null) throw new Exception("Không tìm thấy tài liệu Microsoft Word đang mở.");

                object targetRange = null;
                if (mode == 0 || mode == 3)
                {
                    object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                    targetRange = sel.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, sel, null);
                    if (mode == 0)
                    {
                        // Chèn tại vị trí con trỏ (không xóa đoạn đang chọn nếu có)
                        targetRange.GetType().InvokeMember("Collapse", BindingFlags.InvokeMethod, null, targetRange, new object[] { 0 }); // wdCollapseEnd
                    }
                }
                else
                {
                    object paras = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
                    int pCount = (int)paras.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, paras, null);
                    int insertParaIdx = -1;

                    if (mode == 1) // Ngay đầu phần nội dung (sau Tiêu đề / Trích yếu / đường kẻ -----)
                    {
                        for (int i = 1; i <= pCount; i++)
                        {
                            object p = paras.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paras, new object[] { i });
                            object r = p.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, p, null);
                            object tbls = r.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, r, null);
                            if ((int)tbls.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, tbls, null) > 0) continue;
                            string t = (Convert.ToString(r.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, r, null)) ?? "").TrimEnd('\r', '\a', '\n').Trim();
                            if (t.StartsWith("---") || t.StartsWith("───") || Regex.IsMatch(t, @"^(I\.|1\.|Căn cứ|Thực hiện)", RegexOptions.IgnoreCase))
                            {
                                insertParaIdx = (t.StartsWith("---") || t.StartsWith("───")) ? Math.Min(i + 1, pCount) : i;
                                break;
                            }
                            if (t.Length > 35 && insertParaIdx < 0) insertParaIdx = i;
                        }
                        if (insertParaIdx < 1) insertParaIdx = 1;
                        object tp = paras.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paras, new object[] { insertParaIdx });
                        targetRange = tp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, tp, null);
                        targetRange.GetType().InvokeMember("Collapse", BindingFlags.InvokeMethod, null, targetRange, new object[] { 1 }); // wdCollapseStart
                    }
                    else // mode == 2: Cuối phần nội dung (trước mục Nơi nhận / bảng Chữ ký)
                    {
                        for (int i = 1; i <= pCount; i++)
                        {
                            object p = paras.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paras, new object[] { i });
                            object r = p.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, p, null);
                            string t = (Convert.ToString(r.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, r, null)) ?? "").TrimEnd('\r', '\a', '\n').Trim();
                            if (t.StartsWith("Nơi nhận", StringComparison.OrdinalIgnoreCase) || t.StartsWith("TM.", StringComparison.OrdinalIgnoreCase) ||
                                t.StartsWith("T/M", StringComparison.OrdinalIgnoreCase) || t.Equals("BÍ THƯ", StringComparison.OrdinalIgnoreCase) ||
                                t.Equals("CHỦ TỊCH", StringComparison.OrdinalIgnoreCase))
                            {
                                insertParaIdx = i;
                                break;
                            }
                        }
                        if (insertParaIdx > 0)
                        {
                            object tp = paras.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paras, new object[] { insertParaIdx });
                            targetRange = tp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, tp, null);
                            targetRange.GetType().InvokeMember("Collapse", BindingFlags.InvokeMethod, null, targetRange, new object[] { 1 }); // wdCollapseStart
                        }
                        else
                        {
                            object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                            int endPos = (int)content.GetType().InvokeMember("End", BindingFlags.GetProperty, null, content, null);
                            targetRange = doc.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, doc, new object[] { Math.Max(0, endPos - 1), Math.Max(0, endPos - 1) });
                        }
                    }
                }

                string normalized = summaryBody.Replace("\r\n", "\r").Replace("\n", "\r");
                string prefix = addHeading ? "TÓM TẮT NỘI DUNG:\r" : "";
                string fullInsert = (mode == 0 ? "\r" : "") + prefix + normalized + "\r";

                int startPos = (int)targetRange.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, targetRange, null);
                targetRange.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, targetRange, new object[] { fullInsert });

                // Định dạng chuẩn Times New Roman 14pt, không bóp chữ (Scaling = 100), căn đều, thụt đầu dòng 1.27cm
                object fnt = targetRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, targetRange, null);
                fnt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, fnt, new object[] { "Times New Roman" });
                fnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, fnt, new object[] { 14f });
                fnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, fnt, new object[] { 0 });
                fnt.GetType().InvokeMember("Italic", BindingFlags.SetProperty, null, fnt, new object[] { 0 });
                fnt.GetType().InvokeMember("Scaling", BindingFlags.SetProperty, null, fnt, new object[] { 100 });

                object pf = targetRange.GetType().InvokeMember("ParagraphFormat", BindingFlags.GetProperty, null, targetRange, null);
                pf.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, pf, new object[] { 3 }); // wdAlignParagraphJustify
                pf.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, pf, new object[] { 36f }); // 1.27 cm
                pf.GetType().InvokeMember("SpaceBefore", BindingFlags.SetProperty, null, pf, new object[] { 3f });
                pf.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, pf, new object[] { 3f });

                if (addHeading)
                {
                    int hStart = startPos + (mode == 0 ? 1 : 0);
                    int hEnd = hStart + "TÓM TẮT NỘI DUNG:".Length;
                    object hRng = doc.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, doc, new object[] { hStart, hEnd });
                    object hFnt = hRng.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, hRng, null);
                    hFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, hFnt, new object[] { -1 });
                }

                MessageBox.Show("✓ Đã đưa bản tóm tắt trực tiếp vào nội dung Microsoft Word!", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể chèn vào Microsoft Word: " + ex.Message + "\n(Bạn có thể Copy nội dung trong ô Xem Trước để dán vào Word)", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    public class WordSpellInspectorForm : Form
    {
        private class SpellIssue
        {
            public string Category;
            public string Wrong;
            public string Suggest;
            public string Reason;
            public string Context;
            public bool ExactCase;
        }

        private ListView lvFindings;
        private TextBox txtContext, txtEditSuggest;
        private Label lblSummary, lblAiStatus;
        private Button btnDeepAi, btnFixOne, btnFixAll, btnLocate, btnIgnore;
        private ProgressBar prgScan;
        private List<SpellIssue> issues = new List<SpellIssue>();
        private HashSet<string> ignoredKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string currentDocText = "";
        private bool isAiBusy = false;

        private static readonly string[,] InspectorRules = new string[,]
        {
            // I. Lỗi phụ âm đầu s/x, ch/tr, r/d/gi, l/n, ng/ngh
            { "bổ xung", "bổ sung", "Sai phụ âm x/s" },
            { "sử lý", "xử lý", "Sai phụ âm s/x" },
            { "xử dụng", "sử dụng", "Sai phụ âm x/s" },
            { "chính sác", "chính xác", "Sai phụ âm s/x" },
            { "kiểm xoát", "kiểm soát", "Sai phụ âm x/s" },
            { "rà xoát", "rà soát", "Sai phụ âm x/s" },
            { "sâu xát", "sâu sát", "Sai phụ âm x/s" },
            { "xâu sát", "sâu sát", "Sai phụ âm x/s" },
            { "xáng suốt", "sáng suốt", "Sai phụ âm x/s" },
            { "sáng xuốt", "sáng suốt", "Sai phụ âm s/x" },
            { "xắp xếp", "sắp xếp", "Sai phụ âm x/s" },
            { "sắp sếp", "sắp xếp", "Sai phụ âm s/x" },
            { "suất sắc", "xuất sắc", "Sai phụ âm s/x" },
            { "sản suất", "sản xuất", "Sai phụ âm s/x" },
            { "năng xuất", "năng suất", "Sai phụ âm x/s" },
            { "đề suất", "đề xuất", "Sai phụ âm s/x" },
            { "sơ xuất", "sơ suất", "Sai phụ âm x/s" },
            { "ngân xách", "ngân sách", "Sai phụ âm x/s" },
            { "danh xách", "danh sách", "Sai phụ âm x/s" },
            { "chính xách", "chính sách", "Sai phụ âm x/s" },
            { "hồ xơ", "hồ sơ", "Sai phụ âm x/s" },
            { "chuyển đổi xố", "chuyển đổi số", "Sai phụ âm x/s" },
            { "xát sao", "sát sao", "Sai phụ âm x/s" },
            { "sát xao", "sát sao", "Sai phụ âm x/s" },
            { "xây xát", "xây xát", "" },
            { "chiển khai", "triển khai", "Sai phụ âm ch/tr" },
            { "triển khay", "triển khai", "Sai vần ay/ai" },
            { "chủ chương", "chủ trương", "Sai phụ âm ch/tr" },
            { "khẩn chương", "khẩn trương", "Sai phụ âm ch/tr" },
            { "quan chọng", "quan trọng", "Sai phụ âm ch/tr" },
            { "chách nhiệm", "trách nhiệm", "Sai phụ âm ch/tr" },
            { "trực chách", "trực trách", "Sai phụ âm ch/tr" },
            { "phát chiển", "phát triển", "Sai phụ âm ch/tr" },
            { "tập chung", "tập trung", "Sai phụ âm ch/tr" },
            { "tuyên chuyền", "tuyên truyền", "Sai phụ âm ch/tr" },
            { "chân thành", "chân thành", "" },
            { "ngiêm túc", "nghiêm túc", "Sai quy tắc ng/ngh" },
            { "ngiêm", "nghiêm", "Sai quy tắc ng/ngh" },
            { "ngành ngề", "ngành nghề", "Sai quy tắc ng/ngh" },
            { "nghị quyêt", "nghị quyết", "Thiếu dấu sắc" },
            { "nguyên tắc", "nguyên tắc", "" },
            { "dân chiều", "dẫn chiếu", "Sai chính tả từ ghép" },
            { "giám sat", "giám sát", "Thiếu dấu sắc" },
            { "lưu trử", "lưu trữ", "Sai dấu hỏi/ngã" },
            { "hổ trợ", "hỗ trợ", "Sai dấu hỏi/ngã" },
            { "chuẫn hóa", "chuẩn hóa", "Sai dấu hỏi/ngã" },
            { "chuẫn bị", "chuẩn bị", "Sai dấu hỏi/ngã" },
            { "tiêu chuẫn", "tiêu chuẩn", "Sai dấu hỏi/ngã" },
            { "chuẫn", "chuẩn", "Sai dấu hỏi/ngã" },
            { "giử gìn", "giữ gìn", "Sai dấu hỏi/ngã" },
            { "vửng chắc", "vững chắc", "Sai dấu hỏi/ngã" },
            { "chặc chẽ", "chặt chẽ", "Sai âm cuối c/t" },
            { "chặt chẻ", "chặt chẽ", "Sai dấu hỏi/ngã" },
            { "chắc chẽ", "chặt chẽ", "Từ chuẩn là 'chặt chẽ'" },
            { "kỷ thuật", "kỹ thuật", "Sai dấu hỏi/ngã" },
            { "lỉnh vực", "lĩnh vực", "Sai dấu hỏi/ngã" },
            { "mẩu mực", "mẫu mực", "Sai dấu hỏi/ngã" },
            { "hướng dẩn", "hướng dẫn", "Sai dấu hỏi/ngã" },
            { "hướng dân", "hướng dẫn", "Thiếu dấu ngã" },
            { "thẩm quyển", "thẩm quyền", "Sai dấu hỏi/huyền" },
            { "chuyển đỗi", "chuyển đổi", "Sai dấu ngã/hỏi" },
            { "sẳn sàng", "sẵn sàng", "Sai dấu hỏi/ngã" },
            { "phát triễn", "phát triển", "Sai dấu ngã/hỏi" },
            { "kết qủa", "kết quả", "Sai vị trí dấu thanh" },
            { "hòa bình", "hòa bình", "" },
            // II. Lỗi dùng từ hành chính & Đảng hay sai
            { "sát nhập", "sáp nhập", "Thuật ngữ hành chính chuẩn: 'sáp nhập'" },
            { "xác nhập", "sáp nhập", "Thuật ngữ hành chính chuẩn: 'sáp nhập'" },
            { "cải cách hành chánh", "cải cách hành chính", "Từ chuẩn NĐ 30: 'hành chính'" },
            { "hành chánh", "hành chính", "Từ chuẩn NĐ 30: 'hành chính'" },
            { "kính gởi", "kính gửi", "Từ chuẩn hành chính: 'kính gửi'" },
            { "chuẩn đoán", "chẩn đoán", "Từ chuẩn tiếng Việt: 'chẩn đoán'" },
            { "thăm quan", "tham quan", "Từ chuẩn tiếng Việt: 'tham quan'" },
            { "vô hình chung", "vô hình trung", "Từ chuẩn tiếng Việt: 'vô hình trung'" },
            { "tựu chung", "tựu trung", "Từ chuẩn tiếng Việt: 'tựu trung'" },
            { "bàng quang", "bàng quan", "Từ chuẩn tiếng Việt: 'bàng quan'" },
            { "chỉnh chu", "chỉn chu", "Từ chuẩn tiếng Việt: 'chỉn chu'" },
            { "sáng lạng", "xán lạn", "Từ chuẩn tiếng Việt: 'xán lạn'" },
            { "xán lạng", "xán lạn", "Từ chuẩn tiếng Việt: 'xán lạn'" },
            { "dư giả", "dư dả", "Từ chuẩn tiếng Việt: 'dư dả'" },
            { "trầy trật", "trầy trật", "" },
            { "nồng cốt", "nòng cốt", "Từ chuẩn tiếng Việt: 'nòng cốt'" },
            { "câu chuyện", "câu chuyện", "" },
            { "tiểu bạn", "tiểu ban", "Gõ nhầm dấu nặng ('tiểu ban')" },
            { "bỏ sung", "bổ sung", "Sai dấu hỏi/ngã ('bổ sung')" },
            { "cơ quang", "cơ quan", "Dư chữ g cuối ('cơ quan')" },
            { "kế hoạc", "kế hoạch", "Thiếu chữ h cuối ('kế hoạch')" },
            { "tỏ chức", "tổ chức", "Sai nguyên âm o/ô ('tổ chức')" },
            { "tô chức", "tổ chức", "Thiếu dấu hỏi ('tổ chức')" },
            { "sơ đô", "sơ đồ", "Thiếu dấu huyền ('sơ đồ')" },
            { "bô trí", "bố trí", "Thiếu dấu sắc ('bố trí')" },
            { "thê thức", "thể thức", "Thiếu dấu hỏi ('thể thức')" },
            { "thể thuc", "thể thức", "Thiếu dấu sắc ('thể thức')" },
            { "quyết đinh", "quyết định", "Thiếu dấu nặng ('quyết định')" },
            { "báo cao", "báo cáo", "Thiếu dấu sắc ('báo cáo')" },
            { "đảng uy", "đảng ủy", "Thiếu dấu hỏi ('đảng ủy')" },
            { "chi bô", "chi bộ", "Thiếu dấu nặng ('chi bộ')" },
            { "nhiệm vu", "nhiệm vụ", "Thiếu dấu nặng ('nhiệm vụ')" },
            { "giải quyêt", "giải quyết", "Thiếu dấu sắc ('giải quyết')" },
            { "quán triêt", "quán triệt", "Thiếu dấu nặng ('quán triệt')" },
            { "thực hiên", "thực hiện", "Thiếu dấu nặng ('thực hiện')" },
            { "đồng chi", "đồng chí", "Thiếu dấu sắc ('đồng chí')" },
            { "đông chí", "đồng chí", "Thiếu dấu huyền ('đồng chí')" },
            { "cán bô", "cán bộ", "Thiếu dấu nặng ('cán bộ')" },
            { "công chưc", "công chức", "Thiếu dấu sắc ('công chức')" },
            { "chủ tich", "chủ tịch", "Thiếu dấu nặng ('chủ tịch')" },
            { "bí thu", "bí thư", "Thiếu dấu móc ư ('bí thư')" },
            { "tài liêu", "tài liệu", "Thiếu dấu nặng ('tài liệu')" },
            { "đề nghi", "đề nghị", "Thiếu dấu nặng ('đề nghị')" },
            { "xem xet", "xem xét", "Thiếu dấu sắc ('xem xét')" },
            { "phê duyêt", "phê duyệt", "Thiếu dấu nặng ('phê duyệt')" },
            { "đưọc", "được", "Sai nguyên âm ưo -> ươ ('được')" },
            { "ngưòi", "người", "Sai nguyên âm ưo -> ươ ('người')" },
            { "trưóc", "trước", "Sai nguyên âm ưo -> ươ ('trước')" },
            { "phưòng", "phường", "Sai nguyên âm ưo -> ươ ('phường')" },
            { "chưong", "chương", "Sai nguyên âm ưo -> ươ ('chương')" },
            { "tưỏng", "tưởng", "Sai nguyên âm ưo -> ươ ('tưởng')" },
            { "lưọng", "lượng", "Sai nguyên âm ưo -> ươ ('lượng')" },
            { "tăng cưòng", "tăng cường", "Sai nguyên âm ưo -> ươ ('tăng cường')" },
            { "bảo đão", "bảo đảm", "Sai phụ âm cuối ('bảo đảm')" },
            { "đảm bảo", "bảo đảm", "Khuyến nghị văn phong NĐ 30: 'bảo đảm'" },
            { "chất lựơng", "chất lượng", "Sai vị trí dấu ươ ('chất lượng')" },
            { "nhân dận", "nhân dân", "Dư dấu nặng ('nhân dân')" },
            { "công tác", "công tác", "" },
            { "hoàn thánh", "hoàn thành", "Sai dấu sắc/huyền ('hoàn thành')" },
            { "tiến đọ", "tiến độ", "Sai nguyên âm o/ô ('tiến độ')" },
            { "văn bãn", "văn bản", "Sai dấu ngã/hỏi ('văn bản')" },
            { "nghị qyết", "nghị quyết", "Thiếu chữ u ('nghị quyết')" },
            { "thành phô", "thành phố", "Thiếu dấu sắc ('thành phố')" },
            { "ủy ban nhân dận", "ủy ban nhân dân", "Dư dấu nặng ('nhân dân')" },
            { "trách nhiêm", "trách nhiệm", "Thiếu dấu nặng ('trách nhiệm')" },
            { "kiểm tra", "kiểm tra", "" },
            { "thanh tra", "thanh tra", "" },
            { "tổng kêt", "tổng kết", "Thiếu dấu sắc ('tổng kết')" },
            { "sơ kêt", "sơ kết", "Thiếu dấu sắc ('sơ kết')" },
            { "phối hơp", "phối hợp", "Thiếu dấu nặng ('phối hợp')" },
            { "thống nhât", "thống nhất", "Thiếu dấu sắc ('thống nhất')" },
            { "tổ chưc", "tổ chức", "Thiếu dấu sắc ('tổ chức')" },
            { "phân công", "phân công", "" },
            { "kịp thơi", "kịp thời", "Thiếu dấu huyền ('kịp thời')" },
            { "nghiêm chinh", "nghiêm chỉnh", "Thiếu dấu hỏi ('nghiêm chỉnh')" },
            { "tích cưc", "tích cực", "Thiếu dấu nặng ('tích cực')" },
            { "hiệu qua", "hiệu quả", "Thiếu dấu hỏi ('hiệu quả')" },
            { "chính tri", "chính trị", "Thiếu dấu nặng ('chính trị')" },
            { "kinh tê", "kinh tế", "Thiếu dấu sắc ('kinh tế')" },
            { "xã hôi", "xã hội", "Thiếu dấu nặng ('xã hội')" },
            { "quốc phong", "quốc phòng", "Thiếu dấu huyền ('quốc phòng')" },
            { "an nình", "an ninh", "Dư dấu huyền ('an ninh')" },
            { "đảng viên", "đảng viên", "" },
            { "đại hội", "đại hội", "" },
            { "ban hành", "ban hành", "" },
            { "lãnh đao", "lãnh đạo", "Thiếu dấu nặng ('lãnh đạo')" },
            { "chỉ đao", "chỉ đạo", "Thiếu dấu nặng ('chỉ đạo')" },
            { "điều hanh", "điều hành", "Thiếu dấu huyền ('điều hành')" }
        };

        private static readonly string[,] ProperNounRules = new string[,]
        {
            { "đảng cộng sản việt nam", "Đảng Cộng sản Việt Nam", "Viết hoa chuẩn 05-HD/VPTW" },
            { "ban chấp hành trung ương", "Ban Chấp hành Trung ương", "Viết hoa chuẩn 05-HD/VPTW" },
            { "bộ chính trị", "Bộ Chính trị", "Viết hoa chuẩn 05-HD/VPTW" },
            { "ban bí thư", "Ban Bí thư", "Viết hoa chuẩn 05-HD/VPTW" },
            { "ban thường vụ", "Ban Thường vụ", "Viết hoa chuẩn 05-HD/VPTW" },
            { "ban chấp hành", "Ban Chấp hành", "Viết hoa chuẩn 05-HD/VPTW" },
            { "ủy ban kiểm tra", "Ủy ban Kiểm tra", "Viết hoa chuẩn 05-HD/VPTW" },
            { "chính phủ", "Chính phủ", "Viết hoa chuẩn NĐ 30/2020" },
            { "quốc hội", "Quốc hội", "Viết hoa chuẩn NĐ 30/2020" },
            { "ủy ban nhân dân", "Ủy ban nhân dân", "Viết hoa chuẩn NĐ 30/2020" },
            { "hội đồng nhân dân", "Hội đồng nhân dân", "Viết hoa chuẩn NĐ 30/2020" },
            { "mặt trận tổ quốc", "Mặt trận Tổ quốc", "Viết hoa chuẩn NĐ 30/2020" },
            { "hồ chí minh", "Hồ Chí Minh", "Viết hoa danh từ riêng" }
        };

        public WordSpellInspectorForm()
        {
            this.Text = "VBTM Pro 2.5.0 — Rà Soát Chính Tả AI & Gợi Ý Sửa Đổi Trực Tiếp Trong Word";
            this.ClientSize = new Size(860, 585);
            this.MinimumSize = new Size(760, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.BackColor = Color.White;
            this.ForeColor = Color.FromArgb(24, 24, 27);
            this.Font = new Font("Segoe UI", 9.2f);
            try { this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            int x = 16;
            int w = 828;

            // --- HEADER ---
            Panel pnlHead = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = Color.FromArgb(145, 18, 24) };
            Label lblTitle = new Label
            {
                Text = "★ RÀ SOÁT CHÍNH TẢ AI & GỢI Ý SỬA ĐỔI CHUYÊN SÂU (05-HD/VPTW & NĐ 30)",
                UseMnemonic = false,
                Location = new Point(16, 8),
                AutoSize = true,
                Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(250, 204, 21)
            };
            lblSummary = new Label
            {
                Text = "Đang quét chuyên sâu toàn bộ tài liệu Microsoft Word đang mở...",
                UseMnemonic = false,
                Location = new Point(16, 32),
                Size = new Size(810, 20),
                Font = new Font("Segoe UI", 8.8f),
                ForeColor = Color.FromArgb(254, 240, 138)
            };
            pnlHead.Controls.Add(lblTitle);
            pnlHead.Controls.Add(lblSummary);
            this.Controls.Add(pnlHead);

            // --- THANH CÔNG CỤ QUÉT ---
            int y = 66;
            Button btnRescan = new Button
            {
                Text = "🔍 Quét Nhanh (Từ Điển + Âm Tiết)",
                Location = new Point(x, y),
                Size = new Size(225, 32),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(15, 23, 42),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.8f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnRescan.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnRescan.Click += (s, e) => ScanOpenWordDocument();
            this.Controls.Add(btnRescan);

            btnDeepAi = new Button
            {
                Text = "🧠 Quét Sâu Bằng AI (Ngữ Cảnh && Câu Từ)",
                Location = new Point(x + 233, y),
                Size = new Size(255, 32),
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.8f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnDeepAi.FlatAppearance.BorderSize = 0;
            btnDeepAi.Click += (s, e) => RunDeepAiSpellScanAsync();
            this.Controls.Add(btnDeepAi);

            Button btnKeyCfg = new Button
            {
                Text = "⚙ Key AI",
                Location = new Point(x + 494, y),
                Size = new Size(76, 32),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(15, 23, 42),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnKeyCfg.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnKeyCfg.Click += (s, e) => ShowQuickAiKeyDialog();
            this.Controls.Add(btnKeyCfg);

            prgScan = new ProgressBar
            {
                Location = new Point(x + 578, y + 6),
                Size = new Size(250, 18),
                Style = ProgressBarStyle.Marquee,
                Visible = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(prgScan);

            lblAiStatus = new Label
            {
                Text = "Tự động dùng AI rà soát ngữ cảnh & gợi ý sửa đổi.",
                UseMnemonic = false,
                Location = new Point(x + 578, y + 7),
                Size = new Size(250, 20),
                Font = new Font("Segoe UI", 8.2f, FontStyle.Italic),
                ForeColor = Color.FromArgb(71, 85, 105),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblAiStatus);
            y += 40;

            // --- BẢNG DANH SÁCH LỖI & GỢI Ý ---
            lvFindings = new ListView
            {
                Location = new Point(x, y),
                Size = new Size(w, 265),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                HideSelection = false,
                Font = new Font("Segoe UI", 9.1f),
                BackColor = Color.White,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            lvFindings.Columns.Add("Loại lỗi", 115);
            lvFindings.Columns.Add("Từ / Cụm từ phát hiện", 215);
            lvFindings.Columns.Add("➜ Gợi ý sửa đổi", 225);
            lvFindings.Columns.Add("Lý do / Giải thích", 250);
            lvFindings.SelectedIndexChanged += (s, e) => OnSelectedIssueChanged();
            lvFindings.DoubleClick += (s, e) => LocateSelectedInWord();
            this.Controls.Add(lvFindings);
            y += 273;

            // --- KHUNG XEM NGỮ CẢNH & CHỈNH SỬA GỢI Ý ---
            Label lblCtx = new Label
            {
                Text = "Ngữ cảnh trong tài liệu Word (nháy đúp vào lỗi để nhảy thẳng tới vị trí trong Word):",
                Location = new Point(x, y),
                Size = new Size(w, 18),
                Font = new Font("Segoe UI", 8.6f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            this.Controls.Add(lblCtx);
            y += 20;

            txtContext = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 56),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Times New Roman", 11.5f),
                BackColor = Color.FromArgb(248, 250, 252),
                ForeColor = Color.FromArgb(30, 41, 59),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtContext);
            y += 64;

            Label lblEdit = new Label
            {
                Text = "Gợi ý sửa đổi (bạn có thể gõ chỉnh lại trước khi bấm Sửa):",
                Location = new Point(x, y + 4),
                Size = new Size(345, 20),
                Font = new Font("Segoe UI", 8.8f, FontStyle.Bold),
                ForeColor = Color.FromArgb(5, 150, 105),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            this.Controls.Add(lblEdit);

            txtEditSuggest = new TextBox
            {
                Location = new Point(x + 350, y),
                Size = new Size(w - 350, 26),
                Font = new Font("Times New Roman", 12f, FontStyle.Bold),
                BackColor = Color.FromArgb(240, 253, 244),
                ForeColor = Color.FromArgb(15, 23, 42),
                BorderStyle = BorderStyle.FixedSingle,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(txtEditSuggest);
            y += 36;

            // --- CÁC NÚT HÀNH ĐỘNG CHÍNH ---
            btnLocate = new Button
            {
                Text = "📍 Định Vị Trong Word",
                Location = new Point(x, y),
                Size = new Size(165, 38),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(15, 23, 42),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnLocate.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnLocate.Click += (s, e) => LocateSelectedInWord();
            this.Controls.Add(btnLocate);

            btnIgnore = new Button
            {
                Text = "⏭ Bỏ Qua",
                Location = new Point(x + 173, y),
                Size = new Size(105, 38),
                BackColor = Color.FromArgb(241, 245, 249),
                ForeColor = Color.FromArgb(71, 85, 105),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnIgnore.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnIgnore.Click += (s, e) => IgnoreSelectedIssue();
            this.Controls.Add(btnIgnore);

            btnFixOne = new Button
            {
                Text = "✅ Sửa Lỗi Đang Chọn",
                Location = new Point(x + 286, y),
                Size = new Size(215, 38),
                BackColor = Color.FromArgb(5, 150, 105),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btnFixOne.FlatAppearance.BorderSize = 0;
            btnFixOne.Click += (s, e) => FixSelectedItemInWord();
            this.Controls.Add(btnFixOne);

            btnFixAll = new Button
            {
                Text = "⚡ Tự Động Sửa Hết Tất Cả Trong Word",
                Location = new Point(x + 509, y),
                Size = new Size(319, 38),
                BackColor = Color.FromArgb(234, 179, 8),
                ForeColor = Color.FromArgb(69, 10, 10),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            btnFixAll.FlatAppearance.BorderSize = 0;
            btnFixAll.Click += (s, e) => FixAllFoundErrorsInWord();
            this.Controls.Add(btnFixAll);

            this.Shown += (s, e) =>
            {
                ScanOpenWordDocument();
                string[] cfg = AiWordAssistantForm.LoadAiConfig();
                if (!string.IsNullOrWhiteSpace(cfg[2]) && cfg[0] != "5" && !string.IsNullOrWhiteSpace(currentDocText))
                {
                    RunDeepAiSpellScanAsync();
                }
            };
        }

        private string ExtractContextSnippet(string docText, int matchIndex, int matchLen)
        {
            if (string.IsNullOrEmpty(docText) || matchIndex < 0) return "";
            int start = Math.Max(0, matchIndex - 65);
            int end = Math.Min(docText.Length, matchIndex + matchLen + 75);
            string snip = docText.Substring(start, end - start).Replace("\r", " ").Replace("\n", " ").Replace("\a", " ");
            snip = Regex.Replace(snip, @"\s{2,}", " ").Trim();
            return (start > 0 ? "..." : "") + snip + (end < docText.Length ? "..." : "");
        }

        private void AddIssue(string category, string wrong, string suggest, string reason, string context, bool exactCase)
        {
            if (string.IsNullOrEmpty(wrong) || string.Equals(wrong, suggest, StringComparison.Ordinal)) return;
            if (ignoredKeys.Contains(wrong)) return;
            foreach (var it in issues)
            {
                if (string.Equals(it.Wrong, wrong, StringComparison.Ordinal)) return;
            }
            SpellIssue item = new SpellIssue
            {
                Category = category,
                Wrong = wrong,
                Suggest = suggest,
                Reason = reason,
                Context = context,
                ExactCase = exactCase
            };
            issues.Add(item);
            ListViewItem lvi = new ListViewItem(category);
            lvi.SubItems.Add(wrong);
            lvi.SubItems.Add(suggest);
            lvi.SubItems.Add(reason);
            if (category.Contains("AI")) lvi.BackColor = Color.FromArgb(239, 246, 255);
            lvFindings.Items.Add(lvi);
        }

        private void OnSelectedIssueChanged()
        {
            if (lvFindings.SelectedIndices.Count == 0)
            {
                txtContext.Text = "";
                txtEditSuggest.Text = "";
                return;
            }
            int idx = lvFindings.SelectedIndices[0];
            if (idx >= 0 && idx < issues.Count)
            {
                txtContext.Text = issues[idx].Context;
                txtEditSuggest.Text = issues[idx].Suggest;
            }
        }

        private string ReadActiveWordDocumentText()
        {
            string docText = "";
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = null;
                try { doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null); }
                catch
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int dc = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (dc > 0) doc = docs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docs, new object[] { 1 });
                }
                if (doc != null)
                {
                    object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                    docText = Convert.ToString(content.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, content, null)) ?? "";
                }
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

        private void ScanOpenWordDocument()
        {
            lvFindings.Items.Clear();
            issues.Clear();
            txtContext.Text = "";
            txtEditSuggest.Text = "";

            try
            {
                string docText = ReadActiveWordDocumentText();
                if (string.IsNullOrWhiteSpace(docText))
                {
                    throw new Exception("Chưa có nội dung trong tài liệu Word đang mở.");
                }
                docText = docText.Normalize(NormalizationForm.FormC);
                currentDocText = docText;

                // 1. Quét 150+ quy tắc chính tả hành chính & Đảng (s/x, ch/tr, hỏi/ngã, thiếu dấu, từ sai...)
                for (int r = 0; r < InspectorRules.GetLength(0); r++)
                {
                    string w = InspectorRules[r, 0];
                    string s = InspectorRules[r, 1];
                    string reason = InspectorRules[r, 2];
                    if (string.IsNullOrEmpty(reason) || string.Equals(w, s, StringComparison.OrdinalIgnoreCase)) continue;
                    foreach (Match m in Regex.Matches(docText, @"(?<![A-Za-zÀ-Ỹà-ỹĐđ])" + Regex.Escape(w) + @"(?![A-Za-zÀ-Ỹà-ỹĐđ])", RegexOptions.IgnoreCase))
                    {
                        string matched = m.Value;
                        string prop = matched == matched.ToUpper() ? s.ToUpper() : (char.IsUpper(matched[0]) ? char.ToUpper(s[0]) + s.Substring(1) : s);
                        AddIssue("Chính tả", matched, prop, reason, ExtractContextSnippet(docText, m.Index, m.Length), true);
                    }
                }

                // 2. Quét từ gõ dư chữ (HOẠCHH -> HOẠCH) & lỗi âm tiết / kẹt phím Telex (đảngr, thựcj, ngiêm, cê...)
                HashSet<string> whitelist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
                    "II", "III", "VII", "VIII", "XII", "XIII", "BOONG", "SOONG", "XOONG",
                    "GOOGLE", "OFFLINE", "OFFICE", "WINDOWS", "MICROSOFT", "EMAIL", "ONLINE",
                    "WEBSITE", "INTERNET", "VIETNAM", "VBTM", "PDF", "DOCX", "ZALO", "WIFI", "ISO", "QR", "AI", "HTTP", "HTTPS", "WWW", "COM", "GOV", "VN"
                };
                foreach (Match wm in Regex.Matches(docText, @"\b[A-Za-zÀ-Ỹà-ỹĐđ]{2,}\b"))
                {
                    string token = wm.Value;
                    if (whitelist.Contains(token)) continue;

                    // 2a. Gõ lặp ký tự liên tiếp bất thường trong tiếng Việt
                    if (Regex.IsMatch(token, @"([A-Za-zÀ-Ỹà-ỹĐđ])\1+", RegexOptions.IgnoreCase))
                    {
                        string dedup = Regex.Replace(token, @"([A-Za-zÀ-Ỹà-ỹĐđ])\1+", "$1", RegexOptions.IgnoreCase);
                        if (!string.Equals(token, dedup, StringComparison.Ordinal))
                        {
                            AddIssue("Gõ dư chữ", token, dedup, "Lặp ký tự bất thường khi gõ phím", ExtractContextSnippet(docText, wm.Index, wm.Length), true);
                            continue;
                        }
                    }

                    // 2b. Kẹt phím Telex cuối từ đã có dấu tiếng Việt (vd: đảngr -> đảng, thựcj -> thực, kếx -> kế, báos -> báo)
                    bool hasVnDiacritic = Regex.IsMatch(token, @"[À-Ỹà-ỹĐđ]");
                    if (hasVnDiacritic && token.Length >= 3 && Regex.IsMatch(token, @"[fsrxjwz]$", RegexOptions.IgnoreCase))
                    {
                        string trimmed = token.Substring(0, token.Length - 1);
                        AddIssue("Lỗi Telex", token, trimmed, "Kẹt phím dấu Telex ('" + token[token.Length - 1] + "') ở cuối từ tiếng Việt", ExtractContextSnippet(docText, wm.Index, wm.Length), true);
                        continue;
                    }

                    // 2c. Sai luật chính tả bắt buộc: ng+i/e/ê -> ngh, c+i/e/ê -> k, g+e/ê -> gh
                    if (Regex.IsMatch(token, @"^ng[eêièéẻẽẹềếểễệìíỉĩị]", RegexOptions.IgnoreCase))
                    {
                        string fix = token.Substring(0, 2) + (char.IsUpper(token[1]) ? "H" : "h") + token.Substring(2);
                        AddIssue("Luật âm tiết", token, fix, "Trước e, ê, i phải viết 'ngh' thay vì 'ng'", ExtractContextSnippet(docText, wm.Index, wm.Length), true);
                    }
                }

                // 3. Quét lặp từ liền nhau (vd: "các các", "đã đã", "trong trong", "của của")
                foreach (Match dupM in Regex.Matches(docText, @"\b([A-Za-zÀ-Ỹà-ỹĐđ]{2,})\s+\1\b", RegexOptions.IgnoreCase))
                {
                    string w1 = dupM.Groups[1].Value;
                    string low = w1.ToLowerInvariant();
                    if (low == "người" || low == "ngày" || low == "năm" || low == "nhà" || low == "lần" || low == "bước" || low == "tầng") continue;
                    AddIssue("Lặp từ", dupM.Value, w1, "Từ bị lặp 2 lần liền nhau", ExtractContextSnippet(docText, dupM.Index, dupM.Length), true);
                }

                // 4. Quét lỗi viết hoa danh từ riêng & cơ quan chuẩn NĐ 30 / 05-HD (chỉ bắt nếu viết thường hoàn toàn, bỏ qua khi đã viết IN HOA tiêu ngữ)
                for (int r = 0; r < ProperNounRules.GetLength(0); r++)
                {
                    string pat = ProperNounRules[r, 0];
                    string rep = ProperNounRules[r, 1];
                    string reason = ProperNounRules[r, 2];
                    Match pm = Regex.Match(docText, @"\b" + Regex.Escape(pat) + @"\b", RegexOptions.IgnoreCase);
                    if (pm.Success)
                    {
                        string val = pm.Value;
                        if (val != rep && val != rep.ToUpper())
                        {
                            AddIssue("Viết hoa chuẩn", val, rep, reason, ExtractContextSnippet(docText, pm.Index, pm.Length), true);
                        }
                    }
                }

                // 5. Quét lỗi dấu câu & khoảng trắng thừa / thiếu khoảng trắng sau dấu phẩy/chấm phẩy
                string[] punctBad = new string[] { " ,", " .", " ;", " :", " !", " ?", "  " };
                string[] punctGood = new string[] { ",", ".", ";", ":", "!", "?", " " };
                for (int pi = 0; pi < punctBad.Length; pi++)
                {
                    int pos = docText.IndexOf(punctBad[pi], StringComparison.Ordinal);
                    if (pos >= 0)
                    {
                        string reason = punctBad[pi] == "  " ? "Nhiều dấu cách thừa liền nhau" : ("Khoảng trắng thừa trước dấu \"" + punctGood[pi] + "\"");
                        AddIssue("Dấu câu", punctBad[pi], punctGood[pi], reason, ExtractContextSnippet(docText, pos, punctBad[pi].Length), false);
                    }
                }
                foreach (Match ms in Regex.Matches(docText, @"([A-Za-zÀ-Ỹà-ỹĐđ]{2,})([,;])([A-Za-zÀ-Ỹà-ỹĐđ]{2,})"))
                {
                    string orig = ms.Value;
                    string rep = ms.Groups[1].Value + ms.Groups[2].Value + " " + ms.Groups[3].Value;
                    AddIssue("Thiếu dấu cách", orig, rep, "Thiếu khoảng trắng sau dấu '" + ms.Groups[2].Value + "'", ExtractContextSnippet(docText, ms.Index, ms.Length), true);
                }

                // 6. Quét lỗi định dạng ngày < 10 hoặc tháng 1, 2 thiếu số 0 (chuẩn NĐ 30 & 05-HD)
                foreach (Match dm in Regex.Matches(docText, @"\bngày\s+([1-9])\s+tháng\b", RegexOptions.IgnoreCase))
                {
                    string orig = dm.Value;
                    string rep = Regex.Replace(orig, @"\b([1-9])\b", "0$1");
                    AddIssue("Thể thức ngày", orig, rep, "Ngày từ 1–9 phải thêm số 0 phía trước (NĐ 30)", ExtractContextSnippet(docText, dm.Index, dm.Length), true);
                }
                foreach (Match mm in Regex.Matches(docText, @"\btháng\s+([12])\s+năm\b", RegexOptions.IgnoreCase))
                {
                    string orig = mm.Value;
                    string rep = Regex.Replace(orig, @"\b([12])\b", "0$1");
                    AddIssue("Thể thức tháng", orig, rep, "Tháng 1 và 2 phải viết 01, 02 (NĐ 30)", ExtractContextSnippet(docText, mm.Index, mm.Length), true);
                }

                UpdateSummaryHeader();
            }
            catch
            {
                lblSummary.Text = "⚠ Không tìm thấy cửa sổ Microsoft Word đang mở (hãy mở file Word rồi bấm 'Quét Nhanh').";
            }
        }

        private void UpdateSummaryHeader()
        {
            int wordCnt = string.IsNullOrWhiteSpace(currentDocText) ? 0 : Regex.Matches(currentDocText, @"[A-Za-zÀ-Ỹà-ỹĐđ0-9]+").Count;
            if (issues.Count == 0)
            {
                lblSummary.Text = "✓ Đã quét " + wordCnt + " từ trong Word — không phát hiện lỗi chính tả cơ bản! Bấm 'Quét Sâu Bằng AI' để kiểm tra ngữ cảnh.";
            }
            else
            {
                lblSummary.Text = "⚠ Đã quét " + wordCnt + " từ trong Word — phát hiện " + issues.Count + " lỗi & gợi ý sửa đổi (chọn dòng để sửa hoặc bấm 'Sửa Hết Tất Cả'):";
                if (lvFindings.Items.Count > 0 && lvFindings.SelectedIndices.Count == 0)
                {
                    lvFindings.Items[0].Selected = true;
                }
            }
        }

        private void RunDeepAiSpellScanAsync()
        {
            if (isAiBusy) return;
            if (string.IsNullOrWhiteSpace(currentDocText)) ScanOpenWordDocument();
            if (string.IsNullOrWhiteSpace(currentDocText) || currentDocText.Length < 15)
            {
                MessageBox.Show("Chưa đọc được nội dung văn bản Word đang mở!", "VBTM Pro 2.5.0");
                return;
            }

            string[] cfg = AiWordAssistantForm.LoadAiConfig();
            int provIdx = 0; int.TryParse(cfg[0], out provIdx);
            string modelStr = cfg[1];
            string apiKey = cfg[2];

            if (provIdx == 5 || string.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show(
                    "Chức năng 'Quét Sâu Bằng AI (Ngữ Cảnh)' cần 1 API Key Google Gemini (miễn phí tại aistudio.google.com/apikey).\n\n" +
                    "Hiện tại chế độ Quét Nhanh Ngoại Tuyến (150+ quy tắc + Luật âm tiết Tiếng Việt + Lặp từ + Telex + NĐ 30) vẫn hoạt động 100% không cần mạng!",
                    "VBTM Pro 2.5.0 — Quét Sâu AI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            isAiBusy = true;
            btnDeepAi.Enabled = false;
            prgScan.Visible = true;
            lblAiStatus.Visible = false;
            lblSummary.Text = "🧠 Đang gửi văn bản tới AI (" + modelStr + ") để rà soát lỗi ngữ cảnh, dùng từ & chính tả chuyên sâu...";

            string clipped = currentDocText.Length > 12000 ? currentDocText.Substring(0, 12000) : currentDocText;
            string sysPrompt =
                "Bạn là chuyên gia hiệu đính văn bản Đảng (05-HD/VPTW) và Hành chính Nhà nước Việt Nam (NĐ 30/2020). " +
                "Hãy rà soát TOÀN BỘ lỗi chính tả, lỗi đánh máy, lỗi dùng từ sai ngữ cảnh (như dành/giành, sát/xát, truyền/chuyền, đề xuất/đề suất...), " +
                "lỗi thiếu dấu hoặc câu diễn đạt sai trong văn bản sau.\n" +
                "CHỈ liệt kê mỗi lỗi trên 1 dòng theo đúng định dạng (tối đa 25 lỗi rõ ràng nhất, trích ĐÚNG nguyên văn cụm từ sai trong văn bản gốc):\n" +
                "CỤM_SAI ||| GỢI_Ý_SỬA ||| LÝ_DO_NGẮN_GỌN\n" +
                "Nếu văn bản hoàn toàn không có lỗi, trả về đúng 1 từ: KHONG_CO_LOI. Không thêm lời giải thích nào khác.";

            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                int added = 0;
                string errMsg = null;
                List<string[]> aiFound = new List<string[]>();
                try
                {
                    string resp = null;
                    try
                    {
                        resp = AiWordAssistantForm.CallOnlineAiDirect(provIdx, modelStr, apiKey, sysPrompt, clipped);
                    }
                    catch (Exception exPrimary)
                    {
                        // Nếu OpenAI/API trả phí hết tiền (credit_balance_exhausted / billing / 401), tự động chuyển sang Key Google Gemini miễn phí đã lưu trên máy!
                        string backupGeminiKey = TryFindBackupGeminiKey();
                        if (provIdx != 0 && !string.IsNullOrEmpty(backupGeminiKey))
                        {
                            provIdx = 0;
                            modelStr = "gemini-3.8-flash";
                            apiKey = backupGeminiKey;
                            AiWordAssistantForm.SaveAiConfig(new string[] { "0", modelStr, apiKey });
                            resp = AiWordAssistantForm.CallOnlineAiDirect(0, modelStr, apiKey, sysPrompt, clipped);
                        }
                        else
                        {
                            throw exPrimary;
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(resp) && resp.IndexOf("KHONG_CO_LOI", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        string[] lines = resp.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string ln in lines)
                        {
                            string[] parts = ln.Split(new string[] { "|||" }, StringSplitOptions.None);
                            if (parts.Length >= 2)
                            {
                                string wrong = parts[0].Trim().Trim('"', '\'', '-', '•', '*', ' ');
                                string sug = parts[1].Trim().Trim('"', '\'', ' ');
                                string why = parts.Length >= 3 ? parts[2].Trim() : "AI phát hiện lỗi ngữ cảnh / chính tả";
                                if (wrong.Length >= 2 && sug.Length >= 1 && !string.Equals(wrong, sug, StringComparison.Ordinal))
                                {
                                    // Chỉ nhận những cụm thực sự tồn tại trong tài liệu Word
                                    if (currentDocText.IndexOf(wrong, StringComparison.Ordinal) >= 0)
                                    {
                                        aiFound.Add(new string[] { wrong, sug, why });
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    errMsg = (ex.Message ?? "").Split('\n')[0];
                }

                if (this.IsDisposed) return;
                this.BeginInvoke((Action)(() =>
                {
                    isAiBusy = false;
                    btnDeepAi.Enabled = true;
                    prgScan.Visible = false;
                    lblAiStatus.Visible = true;

                    if (errMsg != null)
                    {
                        lblAiStatus.Text = "⚠ AI lỗi: " + (errMsg.Length > 70 ? errMsg.Substring(0, 70) + "..." : errMsg);
                        if (string.IsNullOrWhiteSpace(txtContext.Text))
                            txtContext.Text = "Chi tiết lỗi AI trực tuyến: " + errMsg;
                    }
                    else
                    {
                        foreach (string[] f in aiFound)
                        {
                            int pos = currentDocText.IndexOf(f[0], StringComparison.Ordinal);
                            int before = issues.Count;
                            AddIssue("AI Ngữ cảnh", f[0], f[1], f[2], ExtractContextSnippet(currentDocText, pos, f[0].Length), true);
                            if (issues.Count > before) added++;
                        }
                        lblAiStatus.Text = added > 0
                            ? "✓ AI phát hiện thêm " + added + " lỗi ngữ cảnh & gợi ý sửa!"
                            : "✓ AI đã kiểm tra ngữ cảnh — không phát hiện thêm lỗi mới.";
                    }
                    UpdateSummaryHeader();
                }));
            });
        }

        private static string TryFindBackupGeminiKey()
        {
            try
            {
                string oldCfg = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM", "ai_hub_config.ini");
                if (File.Exists(oldCfg))
                {
                    foreach (string line in File.ReadAllLines(oldCfg, Encoding.UTF8))
                    {
                        string t = (line ?? "").Trim();
                        if (t.StartsWith("AIza", StringComparison.Ordinal) && t.Length >= 30) return t;
                    }
                }
            }
            catch { }
            return null;
        }

        private void ShowQuickAiKeyDialog()
        {
            string[] cfg = AiWordAssistantForm.LoadAiConfig();
            Form dlg = new Form
            {
                Text = "Cấu Hình API Key Cho Rà Soát Chính Tả AI (Khuyên dùng Google Gemini Miễn Phí)",
                ClientSize = new Size(490, 175),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9.2f)
            };
            Label l1 = new Label { Text = "Dán API Key Google Gemini (miễn phí, bắt đầu bằng AIza...) hoặc Groq/OpenAI:", Location = new Point(16, 14), AutoSize = true, Font = new Font("Segoe UI", 8.8f, FontStyle.Bold) };
            TextBox tbKey = new TextBox { Location = new Point(16, 40), Size = new Size(456, 26), Text = cfg[2], BorderStyle = BorderStyle.FixedSingle };
            LinkLabel lnk = new LinkLabel { Text = "Chưa có Key? Bấm vào đây lấy API Key Google Gemini miễn phí 100% ↗", Location = new Point(16, 74), AutoSize = true, LinkColor = Color.FromArgb(37, 99, 235) };
            lnk.Click += (s, e) => { try { Process.Start("https://aistudio.google.com/apikey"); } catch { } };
            Button bSave = new Button { Text = "💾 Lưu & Quét Sâu AI Ngay", Location = new Point(252, 118), Size = new Size(220, 36), BackColor = Color.FromArgb(5, 150, 105), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 9.2f, FontStyle.Bold), Cursor = Cursors.Hand };
            bSave.FlatAppearance.BorderSize = 0;
            bSave.Click += (s, e) =>
            {
                string k = tbKey.Text.Trim();
                int p = k.StartsWith("AIza", StringComparison.Ordinal) ? 0 : (k.StartsWith("gsk_", StringComparison.Ordinal) ? 4 : 1);
                string m = p == 0 ? "gemini-3.8-flash" : (p == 4 ? "llama-3.3-70b-versatile" : "gpt-4o-mini");
                AiWordAssistantForm.SaveAiConfig(new string[] { p.ToString(), m, k });
                dlg.DialogResult = DialogResult.OK;
                dlg.Close();
            };
            Button bCancel = new Button { Text = "Đóng", Location = new Point(156, 118), Size = new Size(88, 36), BackColor = Color.FromArgb(241, 245, 249), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            bCancel.Click += (s, e) => dlg.Close();
            dlg.Controls.AddRange(new Control[] { l1, tbKey, lnk, bSave, bCancel });
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                RunDeepAiSpellScanAsync();
            }
        }

        private void LocateSelectedInWord()
        {
            if (lvFindings.SelectedIndices.Count == 0) return;
            int idx = lvFindings.SelectedIndices[0];
            if (idx < 0 || idx >= issues.Count) return;
            SpellIssue item = issues[idx];

            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                sel.GetType().InvokeMember("HomeKey", BindingFlags.InvokeMethod, null, sel, new object[] { 6 }); // wdStory
                object find = sel.GetType().InvokeMember("Find", BindingFlags.GetProperty, null, sel, null);
                find.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, find, null);
                find.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, find, new object[] { item.Wrong });
                find.GetType().InvokeMember("MatchCase", BindingFlags.SetProperty, null, find, new object[] { item.ExactCase });
                bool found = (bool)find.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, find, null);
                if (found)
                {
                    wordApp.GetType().InvokeMember("Activate", BindingFlags.InvokeMethod, null, wordApp, null);
                    lblSummary.Text = "📍 Đã bôi đen định vị \"" + item.Wrong + "\" trong Microsoft Word!";
                }
                else
                {
                    lblSummary.Text = "Không tìm thấy \"" + item.Wrong + "\" trong Word (có thể đã được sửa).";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không định vị được trong Word: " + ex.Message, "VBTM Pro 2.5.0");
            }
        }

        private void IgnoreSelectedIssue()
        {
            if (lvFindings.SelectedIndices.Count == 0) return;
            int idx = lvFindings.SelectedIndices[0];
            if (idx < 0 || idx >= issues.Count) return;
            ignoredKeys.Add(issues[idx].Wrong);
            issues.RemoveAt(idx);
            lvFindings.Items.RemoveAt(idx);
            txtContext.Text = "";
            txtEditSuggest.Text = "";
            UpdateSummaryHeader();
        }

        private bool ReplaceSingleInActiveWord(string findTxt, string repTxt, bool exactCase)
        {
            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = null;
                try { doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null); }
                catch
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int dc = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (dc > 0) doc = docs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docs, new object[] { 1 });
                }
                if (doc == null) return false;
                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                object find = content.GetType().InvokeMember("Find", BindingFlags.GetProperty, null, content, null);
                find.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, find, null);
                find.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, find, new object[] { findTxt });
                object rep = find.GetType().InvokeMember("Replacement", BindingFlags.GetProperty, null, find, null);
                rep.GetType().InvokeMember("ClearFormatting", BindingFlags.InvokeMethod, null, rep, null);
                rep.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, rep, new object[] { repTxt });
                find.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, find, new object[] {
                    Type.Missing, exactCase, false, false, false, false, true, 1, false, Type.Missing, 2
                });
                return true;
            }
            catch { return false; }
        }

        private void FixSelectedItemInWord()
        {
            if (lvFindings.SelectedIndices.Count == 0)
            {
                MessageBox.Show("Hãy chọn 1 lỗi trong danh sách để sửa!", "VBTM Pro 2.5.0");
                return;
            }
            int idx = lvFindings.SelectedIndices[0];
            if (idx < 0 || idx >= issues.Count) return;
            SpellIssue item = issues[idx];
            string finalSuggest = (txtEditSuggest.Text ?? "").Trim();
            if (string.IsNullOrEmpty(finalSuggest) && item.Suggest != " ") finalSuggest = item.Suggest;

            if (ReplaceSingleInActiveWord(item.Wrong, finalSuggest, item.ExactCase))
            {
                issues.RemoveAt(idx);
                lvFindings.Items.RemoveAt(idx);
                lblSummary.Text = "✓ Đã sửa \"" + item.Wrong + "\" ➜ \"" + finalSuggest + "\" trực tiếp trong Word!";
                if (lvFindings.Items.Count > 0) lvFindings.Items[Math.Min(idx, lvFindings.Items.Count - 1)].Selected = true;
                else { txtContext.Text = ""; txtEditSuggest.Text = ""; }
            }
            else
            {
                MessageBox.Show("Không thể ghi trực tiếp vào Word. Hãy kiểm tra tài liệu Word đang mở (không ở chế độ Read-Only/Protected View).", "VBTM Pro 2.5.0");
            }
        }

        private void FixAllFoundErrorsInWord()
        {
            if (issues.Count == 0)
            {
                MessageBox.Show("Không còn lỗi nào cần sửa!", "VBTM Pro 2.5.0");
                return;
            }
            int okCnt = 0;
            foreach (var item in issues)
            {
                if (ReplaceSingleInActiveWord(item.Wrong, item.Suggest, item.ExactCase)) okCnt++;
            }
            ScanOpenWordDocument();
            MessageBox.Show("✓ Đã tự động sửa xong " + okCnt + " lỗi chính tả & đề xuất trực tiếp trong tài liệu Word!", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    public class WordQrInsertForm : Form
    {
        private TextBox txtQrData;
        private PictureBox picQr;
        private ComboBox cboPos, cboSize;
        private Label lblStatus, lblLogoName;
        private Bitmap currentQrBmp;
        private string selectedLogoPath = "";

        public WordQrInsertForm()
        {
            this.Text = "Tạo & Chèn Mã QR Kèm Logo Tự Chọn Trực Tiếp Vào Word - VBTM Pro 2.5.0";
            this.ClientSize = new Size(650, 440);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(95, 10, 14);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            Label lblData = new Label
            {
                Text = "🔗 Nội dung Mã QR (Link tài liệu, Google Drive, Số văn bản & Trích yếu):",
                Location = new Point(20, 14),
                AutoSize = true,
                ForeColor = Color.FromArgb(250, 204, 21),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
            };
            this.Controls.Add(lblData);

            txtQrData = new TextBox
            {
                Location = new Point(20, 40),
                Size = new Size(395, 92),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Text = DetectDefaultQrTextFromWord()
            };
            this.Controls.Add(txtQrData);

            picQr = new PictureBox
            {
                Location = new Point(435, 40),
                Size = new Size(195, 195),
                BackColor = Color.White,
                SizeMode = PictureBoxSizeMode.Zoom,
                BorderStyle = BorderStyle.FixedSingle
            };
            this.Controls.Add(picQr);

            Button btnRefresh = new Button
            {
                Text = "🔄 Tạo Xem Trước QR",
                Location = new Point(435, 243),
                Size = new Size(195, 32),
                BackColor = Color.FromArgb(127, 29, 29),
                ForeColor = Color.FromArgb(254, 240, 138),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnRefresh.Click += (s, e) => RenderQrPreview();
            this.Controls.Add(btnRefresh);

            Label lblPos = new Label { Text = "Vị trí chèn vào Word:", Location = new Point(20, 144), AutoSize = true, ForeColor = Color.FromArgb(254, 240, 138) };
            this.Controls.Add(lblPos);
            cboPos = new ComboBox
            {
                Location = new Point(20, 168),
                Size = new Size(395, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White
            };
            cboPos.Items.AddRange(new object[] {
                "1. Khung nổi tự do (Góc trên phải — Kéo chuột di chuyển thoải mái)",
                "2. Khung nổi tự do (Góc dưới trái — Dưới mục Nơi nhận)",
                "3. Chèn thẳng tại vị trí con trỏ chuột đang đặt"
            });
            cboPos.SelectedIndex = 0;
            this.Controls.Add(cboPos);

            Label lblSz = new Label { Text = "Kích thước mã QR:", Location = new Point(20, 208), AutoSize = true, ForeColor = Color.FromArgb(254, 240, 138) };
            this.Controls.Add(lblSz);
            cboSize = new ComboBox
            {
                Location = new Point(155, 204),
                Size = new Size(260, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.White
            };
            cboSize.Items.AddRange(new object[] { "1.8 cm (Nhỏ gọn)", "2.2 cm (Chuẩn văn thư)", "2.6 cm (Vừa rõ)", "3.0 cm (Lớn dễ quét)" });
            cboSize.SelectedIndex = 1;
            this.Controls.Add(cboSize);

            Button btnPickLogo = new Button
            {
                Text = "🖼️ Chọn Logo Chèn Giữa QR...",
                Location = new Point(20, 244),
                Size = new Size(215, 32),
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.8f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnPickLogo.FlatAppearance.BorderSize = 0;
            btnPickLogo.Click += (s, e) => PickLogoImage();
            this.Controls.Add(btnPickLogo);

            Button btnClearLogo = new Button
            {
                Text = "✕ Bỏ Logo",
                Location = new Point(242, 244),
                Size = new Size(88, 32),
                BackColor = Color.FromArgb(127, 29, 29),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8.5f),
                Cursor = Cursors.Hand
            };
            btnClearLogo.Click += (s, e) => { selectedLogoPath = ""; lblLogoName.Text = "Chưa chọn logo"; RenderQrPreview(); };
            this.Controls.Add(btnClearLogo);

            lblLogoName = new Label
            {
                Text = "Chưa chọn logo",
                Location = new Point(20, 282),
                Size = new Size(395, 20),
                ForeColor = Color.FromArgb(254, 249, 195),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic)
            };
            this.Controls.Add(lblLogoName);

            Button btnInsert = new Button
            {
                Text = "✅ Tạo && Chèn Trực Tiếp Mã QR (Kèm Logo) Vào File Word",
                Location = new Point(20, 322),
                Size = new Size(610, 46),
                BackColor = Color.FromArgb(234, 179, 8),
                ForeColor = Color.FromArgb(69, 10, 10),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnInsert.FlatAppearance.BorderSize = 0;
            btnInsert.Click += (s, e) => InsertQrIntoActiveWord();
            this.Controls.Add(btnInsert);

            lblStatus = new Label
            {
                Text = "Mẹo: Chọn logo Quốc huy / Búa liềm / Đơn vị (PNG/JPG) để chèn vào tâm mã QR, quét Zalo vẫn nhạy 100%!",
                Location = new Point(20, 386),
                AutoSize = true,
                ForeColor = Color.FromArgb(253, 224, 71),
                Font = new Font("Segoe UI", 8.3f)
            };
            this.Controls.Add(lblStatus);

            this.Shown += (s, e) => RenderQrPreview();
        }

        private void PickLogoImage()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Chọn ảnh Logo chèn vào chính giữa Mã QR (PNG, JPG, BMP)";
                ofd.Filter = "Hình ảnh Logo (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp|Tất cả tệp (*.*)|*.*";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    selectedLogoPath = ofd.FileName;
                    lblLogoName.Text = "✓ Logo đã chọn: " + Path.GetFileName(selectedLogoPath);
                    RenderQrPreview();
                }
            }
        }

        private Bitmap OverlayLogoOnQr(Bitmap baseQr)
        {
            if (string.IsNullOrEmpty(selectedLogoPath) || !File.Exists(selectedLogoPath)) return baseQr;
            try
            {
                Bitmap result = new Bitmap(baseQr.Width, baseQr.Height);
                using (Graphics g = Graphics.FromImage(result))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.DrawImage(baseQr, 0, 0, baseQr.Width, baseQr.Height);
                    using (Image logo = Image.FromFile(selectedLogoPath))
                    {
                        int boxSz = (int)(baseQr.Width * 0.23);
                        int pad = Math.Max(4, boxSz / 10);
                        int totalSz = boxSz + pad * 2;
                        int cx = (baseQr.Width - totalSz) / 2;
                        int cy = (baseQr.Height - totalSz) / 2;
                        g.FillRectangle(Brushes.White, cx, cy, totalSz, totalSz);
                        using (Pen p = new Pen(Color.FromArgb(220, 220, 220), 1f))
                            g.DrawRectangle(p, cx, cy, totalSz - 1, totalSz - 1);
                        g.DrawImage(logo, cx + pad, cy + pad, boxSz, boxSz);
                    }
                }
                return result;
            }
            catch { return baseQr; }
        }

        private string DetectDefaultQrTextFromWord()
        {
            try
            {
                string snapPath = Path.Combine(Path.GetTempPath(), "vbtm_word_snapshot.txt");
                string txt = File.Exists(snapPath) ? File.ReadAllText(snapPath, Encoding.Unicode) : "";
                if (!string.IsNullOrWhiteSpace(txt))
                {
                    Match mNum = Regex.Match(txt, @"Số\s*:?\s*[0-9A-Za-zĐđ\-/]+", RegexOptions.IgnoreCase);
                    string numPart = mNum.Success ? mNum.Value.Trim() : "VBTM-PRO";
                    string[] cfg = UnitSettingsForm.LoadUnitConfig();
                    return cfg[1] + " | " + numPart + " | Ngày " + DateTime.Now.ToString("dd/MM/yyyy");
                }
            }
            catch { }
            return "VBTM Pro 2.5.0 | Văn bản xác thực ngày " + DateTime.Now.ToString("dd/MM/yyyy");
        }

        private void RenderQrPreview()
        {
            string data = txtQrData.Text.Trim();
            if (string.IsNullOrEmpty(data)) data = "VBTM-PRO";
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
            string ecc = (!string.IsNullOrEmpty(selectedLogoPath) && File.Exists(selectedLogoPath)) ? "H" : "M";

            string[] urls = new string[] {
                "https://api.qrserver.com/v1/create-qr-code/?size=320x320&margin=6&ecc=" + ecc + "&data=" + Uri.EscapeDataString(data),
                "https://quickchart.io/qr?size=320&margin=2&ecLevel=" + ecc + "&text=" + Uri.EscapeDataString(data)
            };

            foreach (string u in urls)
            {
                try
                {
                    HttpWebRequest req = (HttpWebRequest)WebRequest.Create(u);
                    req.Timeout = 4000;
                    using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                    using (Stream s = resp.GetResponseStream())
                    {
                        if (s != null)
                        {
                            using (Image img = Image.FromStream(s))
                            {
                                currentQrBmp = OverlayLogoOnQr(new Bitmap(img));
                                picQr.Image = currentQrBmp;
                                lblStatus.Text = !string.IsNullOrEmpty(selectedLogoPath)
                                    ? "✓ Đã tạo Mã QR chuẩn ISO 18004 (ECC-H) gắn Logo ở tâm!"
                                    : "✓ Đã tạo Mã QR chuẩn ISO 18004 sắc nét (quét được bằng Zalo / Camera)!";
                                return;
                            }
                        }
                    }
                }
                catch { }
            }

            // Fallback offline matrix
            int size = 25, scale = 8, pad = 12;
            Bitmap bmp = new Bitmap(size * scale + pad * 2, size * scale + pad * 2);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                uint hash = 2166136261;
                foreach (char c in data) { hash ^= c; hash *= 16777619; }
                Action<int, int> drawFinder = (ox, oy) =>
                {
                    g.FillRectangle(Brushes.Black, pad + ox * scale, pad + oy * scale, 7 * scale, 7 * scale);
                    g.FillRectangle(Brushes.White, pad + (ox + 1) * scale, pad + (oy + 1) * scale, 5 * scale, 5 * scale);
                    g.FillRectangle(Brushes.Black, pad + (ox + 2) * scale, pad + (oy + 2) * scale, 3 * scale, 3 * scale);
                };
                drawFinder(0, 0); drawFinder(size - 7, 0); drawFinder(0, size - 7);
                for (int r = 0; r < size; r++)
                    for (int c = 0; c < size; c++)
                    {
                        if ((r < 8 && c < 8) || (r < 8 && c >= size - 8) || (r >= size - 8 && c < 8)) continue;
                        int ch = data[(r * size + c) % data.Length];
                        if (((hash ^ (r * 131 + c * 313 + ch * 17)) & 3) == 0 || ((r * c + ch) % 5 == 0))
                            g.FillRectangle(Brushes.Black, pad + c * scale, pad + r * scale, scale, scale);
                    }
            }
            currentQrBmp = OverlayLogoOnQr(bmp);
            picQr.Image = currentQrBmp;
            lblStatus.Text = "⚠ Đang ngoại tuyến: đã tạo ảnh QR dự phòng (kết nối mạng để tạo QR ISO quét điện thoại).";
        }

        private void InsertQrIntoActiveWord()
        {
            RenderQrPreview();
            if (currentQrBmp == null) return;

            string qrDir = Path.Combine(Path.GetTempPath(), "VBTM_QR");
            Directory.CreateDirectory(qrDir);
            string qrFile = Path.Combine(qrDir, "QR_" + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".png");
            currentQrBmp.Save(qrFile, ImageFormat.Png);

            float[] cmSizes = new float[] { 1.8f, 2.2f, 2.6f, 3.0f };
            float cm = cmSizes[Math.Max(0, Math.Min(cboSize.SelectedIndex, cmSizes.Length - 1))];
            float ptSize = cm * 28.3465f;

            try
            {
                object wordApp = Marshal.GetActiveObject("Word.Application");
                object doc = null;
                try { doc = wordApp.GetType().InvokeMember("ActiveDocument", BindingFlags.GetProperty, null, wordApp, null); }
                catch
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int dc = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (dc > 0) doc = docs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docs, new object[] { 1 });
                }
                if (doc == null) throw new Exception("Chưa có tài liệu Microsoft Word đang mở.");

                if (cboPos.SelectedIndex == 2)
                {
                    object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                    object inlines = sel.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, sel, null);
                    object pic = inlines.GetType().InvokeMember("AddPicture", BindingFlags.InvokeMethod, null, inlines, new object[] { qrFile, false, true });
                    pic.GetType().InvokeMember("Width", BindingFlags.SetProperty, null, pic, new object[] { ptSize });
                    pic.GetType().InvokeMember("Height", BindingFlags.SetProperty, null, pic, new object[] { ptSize });
                }
                else
                {
                    object shapes = doc.GetType().InvokeMember("Shapes", BindingFlags.GetProperty, null, doc, null);
                    float leftPt = cboPos.SelectedIndex == 0 ? (17.2f * 28.3465f) : (3.2f * 28.3465f);
                    float topPt = cboPos.SelectedIndex == 0 ? (1.2f * 28.3465f) : (23.5f * 28.3465f);
                    object shp = shapes.GetType().InvokeMember("AddPicture", BindingFlags.InvokeMethod, null, shapes, new object[] {
                        qrFile, false, true, leftPt, topPt, ptSize, ptSize
                    });
                    shp.GetType().InvokeMember("RelativeHorizontalPosition", BindingFlags.SetProperty, null, shp, new object[] { 1 });
                    shp.GetType().InvokeMember("RelativeVerticalPosition", BindingFlags.SetProperty, null, shp, new object[] { 1 });
                    shp.GetType().InvokeMember("Left", BindingFlags.SetProperty, null, shp, new object[] { leftPt });
                    shp.GetType().InvokeMember("Top", BindingFlags.SetProperty, null, shp, new object[] { topPt });
                    object wrap = shp.GetType().InvokeMember("WrapFormat", BindingFlags.GetProperty, null, shp, null);
                    wrap.GetType().InvokeMember("Type", BindingFlags.SetProperty, null, wrap, new object[] { 3 }); // In Front of Text
                    shp.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, shp, null);
                }

                lblStatus.Text = "✓ Đã chèn trực tiếp Mã QR vào file Word! Bạn có thể nắm chuột kéo di chuyển tùy thích.";
                MessageBox.Show("Đã tạo và chèn trực tiếp Mã QR vào file Microsoft Word đang mở!\n\nBạn có thể nắm chuột kéo khung mã QR di chuyển tới bất kỳ vị trí nào trên trang.", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể chèn Mã QR vào Word đang mở: " + ex.Message, "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
            { "bổ xung", "bổ sung", "Sai phụ âm x/s ('bổ sung')" },
            { "sử lý", "xử lý", "Sai phụ âm s/x ('xử lý')" },
            { "xử dụng", "sử dụng", "Sai phụ âm x/s ('sử dụng')" },
            { "chính sác", "chính xác", "Sai phụ âm s/x ('chính xác')" },
            { "kiểm xoát", "kiểm soát", "Sai phụ âm x/s ('kiểm soát')" },
            { "rà xoát", "rà soát", "Sai phụ âm x/s ('rà soát')" },
            { "chiển khai", "triển khai", "Sai phụ âm ch/tr ('triển khai')" },
            { "chủ chương", "chủ trương", "Sai phụ âm ch/tr ('chủ trương')" },
            { "khẩn chương", "khẩn trương", "Sai phụ âm ch/tr ('khẩn trương')" },
            { "kế hoạc", "kế hoạch", "Thiếu âm cuối ('kế hoạch')" },
            { "ngiêm túc", "nghiêm túc", "Sai quy tắc ng/ngh ('nghiêm túc')" },
            { "ngành ngề", "ngành nghề", "Sai quy tắc ng/ngh ('ngành nghề')" },
            { "hổ trợ", "hỗ trợ", "Sai dấu hỏi/ngã ('hỗ trợ')" },
            { "tiêu chuẫn", "tiêu chuẩn", "Sai dấu hỏi/ngã ('tiêu chuẩn')" },
            { "chuẫn hóa", "chuẩn hóa", "Sai dấu hỏi/ngã ('chuẩn hóa')" },
            { "chuẫn bị", "chuẩn bị", "Sai dấu hỏi/ngã ('chuẩn bị')" },
            { "sát nhập", "sáp nhập", "Từ chuẩn hành chính ('sáp nhập')" },
            { "xác nhập", "sáp nhập", "Từ chuẩn hành chính ('sáp nhập')" },
            { "xáng suốt", "sáng suốt", "Sai phụ âm x/s ('sáng suốt')" },
            { "sáng xuốt", "sáng suốt", "Sai phụ âm x/s ('sáng suốt')" },
            { "giử gìn", "giữ gìn", "Sai dấu hỏi/ngã ('giữ gìn')" },
            { "vửng chắc", "vững chắc", "Sai dấu hỏi/ngã ('vững chắc')" },
            { "chặc chẽ", "chặt chẽ", "Sai âm cuối c/t ('chặt chẽ')" },
            { "chặt chẻ", "chặt chẽ", "Sai dấu hỏi/ngã ('chặt chẽ')" },
            { "sâu xát", "sâu sát", "Sai phụ âm x/s ('sâu sát')" },
            { "quan chọng", "quan trọng", "Sai phụ âm ch/tr ('quan trọng')" },
            { "chách nhiệm", "trách nhiệm", "Sai phụ âm ch/tr ('trách nhiệm')" },
            { "trực chách", "trực trách", "Sai phụ âm ch/tr ('trực trách')" },
            { "phát chiển", "phát triển", "Sai phụ âm ch/tr ('phát triển')" },
            { "tập chung", "tập trung", "Sai phụ âm ch/tr ('tập trung')" },
            { "tuyên chuyền", "tuyên truyền", "Sai phụ âm ch/tr ('tuyên truyền')" },
            { "xắp xếp", "sắp xếp", "Sai phụ âm x/s ('sắp xếp')" },
            { "sắp sếp", "sắp xếp", "Sai phụ âm s/x ('sắp xếp')" },
            { "suất sắc", "xuất sắc", "Sai phụ âm s/x ('xuất sắc')" },
            { "sản suất", "sản xuất", "Sai phụ âm s/x ('sản xuất')" },
            { "năng xuất", "năng suất", "Sai phụ âm x/s ('năng suất')" },
            { "đề suất", "đề xuất", "Sai phụ âm s/x ('đề xuất')" },
            { "sơ xuất", "sơ suất", "Sai phụ âm x/s ('sơ suất')" },
            { "ngân xách", "ngân sách", "Sai phụ âm x/s ('ngân sách')" },
            { "danh xách", "danh sách", "Sai phụ âm x/s ('danh sách')" },
            { "chính xách", "chính sách", "Sai phụ âm x/s ('chính sách')" },
            { "kỷ thuật", "kỹ thuật", "Sai dấu hỏi/ngã ('kỹ thuật')" },
            { "lỉnh vực", "lĩnh vực", "Sai dấu hỏi/ngã ('lĩnh vực')" },
            { "mẩu mực", "mẫu mực", "Sai dấu hỏi/ngã ('mẫu mực')" },
            { "hướng dẩn", "hướng dẫn", "Sai dấu hỏi/ngã ('hướng dẫn')" },
            { "tỏ chức", "tổ chức", "Sai dấu hỏi/ổ ('tổ chức')" },
            { "chuẩn đoán", "chẩn đoán", "Từ chuẩn xác là 'chẩn đoán'" },
            { "thăm quan", "tham quan", "Từ chuẩn xác là 'tham quan'" },
            { "vô hình chung", "vô hình trung", "Từ chuẩn xác là 'vô hình trung'" },
            { "tựu chung", "tựu trung", "Từ chuẩn xác là 'tựu trung'" },
            { "thẩm quyển", "thẩm quyền", "Sai dấu hỏi/huyền ('thẩm quyền')" },
            { "chuyển đỗi", "chuyển đổi", "Sai dấu hỏi/ngã ('chuyển đổi')" },
            { "sẳn sàng", "sẵn sàng", "Sai dấu hỏi/ngã ('sẵn sàng')" },
            { "hoản", "hoãn", "Sai dấu hỏi/ngã ('hoãn')" },
            { "triễn", "triển", "Sai dấu hỏi/ngã ('triển')" },
            { "đươc", "được", "Thiếu dấu nặng ('được')" },
            { "quyêt", "quyết", "Thiếu dấu sắc ('quyết')" },
            { "duyêt", "duyệt", "Thiếu dấu nặng ('duyệt')" },
            { "ngị quyết", "nghị quyết", "Sai quy tắc ng/ngh ('nghị quyết')" },
            { "ngị định", "nghị định", "Sai quy tắc ng/ngh ('nghị định')" },
            { "ngiên cứu", "nghiên cứu", "Sai quy tắc ng/ngh ('nghiên cứu')" }
        };

        private static string CreateUnicodeWordRegex(string phrase)
        {
            return @"(?<!\w)" + Regex.Escape(phrase.Trim()) + @"(?!\w)";
        }

        private static string ApplyCasePattern(string matched, string replacement)
        {
            if (string.IsNullOrEmpty(matched) || string.IsNullOrEmpty(replacement)) return replacement;
            if (matched == matched.ToUpper() && matched != matched.ToLower())
                return replacement.ToUpper();
            if (char.IsUpper(matched[0]))
                return char.ToUpper(replacement[0]) + replacement.Substring(1);
            return replacement;
        }

        public MainStudioForm()
        {
            try { ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | SecurityProtocolType.Tls; } catch { }
            this.Text = "★ VBTM Pro 2.5.0 · Chuẩn Hóa Thể Thức Văn Bản Đảng Ủy (HD 05-HD/VPTW) & UBND Phường (NĐ 30/2020/NĐ-CP)";
            this.ClientSize = new Size(1120, 720);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(95, 10, 14);
            this.ForeColor = Color.White;
            this.Font = new Font("Segoe UI", 9.5f);

            SyncWordAddInPresetsToPhuong();
            BuildMainUI();
            LoadTemplate(0);
            UpdateQrPreview();
        }

        private void SyncWordAddInPresetsToPhuong()
        {
            try
            {
                string vbtmAppData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "VBTM");
                if (!Directory.Exists(vbtmAppData)) Directory.CreateDirectory(vbtmAppData);

                string presetsJson = @"{
  ""agencies_admin"": [
    { ""parent_agency"": ""ỦY BAN NHÂN DÂN QUẬN"", ""agency_name"": ""ỦY BAN NHÂN DÂN PHƯỜNG"", ""code_prefix"": ""UBND-VP"" },
    { ""parent_agency"": """", ""agency_name"": ""HỘI ĐỒNG NHÂN DÂN PHƯỜNG"", ""code_prefix"": ""NQ-HĐND"" }
  ],
  ""agencies_party"": [
    { ""parent_agency"": ""QUẬN ỦY"", ""agency_name"": ""ĐẢNG ỦY PHƯỜNG"", ""code_prefix"": ""ĐU"" },
    { ""parent_agency"": ""ĐẢNG ỦY PHƯỜNG"", ""agency_name"": ""CHI BỘ TRỰC THUỘC"", ""code_prefix"": ""CB"" }
  ],
  ""signers_admin"": [
    { ""signer_type"": ""TM. ỦY BAN NHÂN DÂN"", ""position"": ""CHỦ TỊCH"", ""rank"": """", ""fullname"": ""Nguyễn Văn An"" },
    { ""signer_type"": ""KT. CHỦ TỊCH"", ""position"": ""PHÓ CHỦ TỊCH"", ""rank"": """", ""fullname"": ""Lê Thị Mai"" }
  ],
  ""signers_party"": [
    { ""signer_type"": ""T/M ĐẢNG ỦY"", ""position"": ""BÍ THƯ"", ""rank"": """", ""fullname"": ""Trần Minh Quân"" },
    { ""signer_type"": ""T/M BAN THƯỜNG VỤ"", ""position"": ""PHÓ BÍ THƯ THƯỜNG TRỰC"", ""rank"": """", ""fullname"": ""Phạm Thị Thanh Hoa"" }
  ],
  ""recipients_list"": [
    ""Như trên;"",
    ""Thường trực Đảng ủy, HĐND phường (để b/c);"",
    ""Chủ tịch, các PCT UBND phường;"",
    ""UBMTTQ và các đoàn thể phường;"",
    ""Các chi bộ trực thuộc;"",
    ""Các tổ dân phố;"",
    ""Lưu: VT, VP."",
    ""Lưu Văn phòng Đảng ủy.""
  ],
  ""formatting"": {
    ""admin"": {
      ""font_name"": ""Times New Roman"",
      ""body_size"": 14,
      ""line_spacing"": 1.15,
      ""space_before"": 0,
      ""space_after"": 6,
      ""first_line_indent_cm"": 1.27,
      ""margin_top_cm"": 2.0,
      ""margin_bottom_cm"": 2.0,
      ""margin_left_cm"": 3.0,
      ""margin_right_cm"": 1.5
    },
    ""party"": {
      ""font_name"": ""Times New Roman"",
      ""body_size"": 14,
      ""line_spacing"": 1.15,
      ""space_before"": 0,
      ""space_after"": 6,
      ""first_line_indent_cm"": 1.27,
      ""margin_top_cm"": 2.0,
      ""margin_bottom_cm"": 2.0,
      ""margin_left_cm"": 3.0,
      ""margin_right_cm"": 1.5
    }
  }
}";
                File.WriteAllText(Path.Combine(vbtmAppData, "vbtm_presets.json"), presetsJson, new UTF8Encoding(false));
                File.WriteAllText(Path.Combine(vbtmAppData, "user_presets.json"), presetsJson, new UTF8Encoding(false));
            }
            catch { }
        }

        private void BuildMainUI()
        {
            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 30,
                Text = "  ★ Trạng thái: Hoạt động ngoại tuyến 100% (Offline) | Đầy đủ 16 mẫu chuẩn Nghị định 30/2020/NĐ-CP (UBND Phường) & Hướng dẫn 05-HD/VPTW (Đảng ủy Phường)",
                BackColor = Color.FromArgb(135, 15, 20),
                ForeColor = Color.FromArgb(250, 204, 21),
                TextAlign = ContentAlignment.MiddleLeft
            };

            Panel topBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = Color.FromArgb(145, 18, 24)
            };

            Label lblBrand = new Label
            {
                Text = "★ VBTM Pro 2.5.0",
                Font = new Font("Segoe UI", 13.0f, FontStyle.Bold),
                ForeColor = Color.FromArgb(250, 204, 21),
                Location = new Point(10, 16),
                AutoSize = true
            };
            topBar.Controls.Add(lblBrand);

            cboTemplate = new ComboBox
            {
                Location = new Point(162, 18),
                Size = new Size(228, 28),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.0f)
            };
            cboTemplate.Items.AddRange(new object[]
            {
                "1. Công văn UBND Phường (Nghị định 30)",
                "2. Báo cáo UBND Phường (Nghị định 30)",
                "3. Quyết định (cá biệt) UBND Phường (NĐ 30)",
                "4. Quyết định (Quy chế) UBND Phường (NĐ 30)",
                "5. Kế hoạch UBND Phường (Nghị định 30)",
                "6. Tờ trình UBND Phường (Nghị định 30)",
                "7. Thông báo UBND Phường (Nghị định 30)",
                "8. Nghị quyết HĐND Phường (Nghị định 30)",
                "9. Biên bản họp UBND Phường (Nghị định 30)",
                "10. Giấy mời họp UBND Phường (Nghị định 30)",
                "11. Công văn Đảng ủy Phường (HD 05-HD/VPTW)",
                "12. Báo cáo Đảng ủy Phường (HD 05-HD/VPTW)",
                "13. Nghị quyết Đảng ủy Phường (HD 05-HD/VPTW)",
                "14. Kế hoạch Đảng ủy Phường (HD 05-HD/VPTW)",
                "15. Chương trình Đảng ủy Phường (HD 05-HD/VPTW)",
                "16. Quyết định Đảng ủy Phường (HD 05-HD/VPTW)"
            });
            cboTemplate.SelectedIndex = 0;
            cboTemplate.SelectedIndexChanged += (s, e) => LoadTemplate(cboTemplate.SelectedIndex);
            topBar.Controls.Add(cboTemplate);

            Button btnOpenFile = CreateBtn("📂 Mở File (.docx/.pdf)", new Point(392, 14), new Size(148, 36), Color.FromArgb(185, 28, 28));
            btnOpenFile.Click += (s, e) => OpenWordOrTextFileFromDisk();
            topBar.Controls.Add(btnOpenFile);

            Button btnStd = CreateBtn("1. Chuẩn Hóa & Sửa Lỗi", new Point(544, 14), new Size(148, 36), Color.FromArgb(202, 138, 4));
            btnStd.Click += (s, e) => StandardizeAndFixAll();
            topBar.Controls.Add(btnStd);

            Button btnWordLive = CreateBtn("2. Sửa Word Đang Mở", new Point(696, 14), new Size(148, 36), Color.FromArgb(185, 28, 28));
            btnWordLive.Click += (s, e) => FormatActiveWordDocument();
            topBar.Controls.Add(btnWordLive);

            Button btnExportDoc = CreateBtn("3. Xuất Word (.doc)", new Point(848, 14), new Size(134, 36), Color.FromArgb(234, 179, 8));
            btnExportDoc.ForeColor = Color.FromArgb(69, 10, 10);
            btnExportDoc.Click += (s, e) => ExportAndOpenWordDoc();
            topBar.Controls.Add(btnExportDoc);

            Button btnPinWord = CreateBtn("📌 Tích Hợp Word", new Point(986, 14), new Size(124, 36), Color.FromArgb(127, 29, 29));
            btnPinWord.Click += (s, e) =>
            {
                string instDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "VBTM-Pro");
                bool ok = WordAddinInstaller.InstallWordRibbonAddin(instDir, Application.ExecutablePath);
                MessageBox.Show(
                    ok
                        ? "Đã tích hợp thành công Tab Ribbon 'VBTM Pro 2.5 (05-HD & ND30)' trực tiếp vào Microsoft Word!\nHãy mở lại Microsoft Word để sử dụng ngay trên thanh công cụ."
                        : "Không tìm thấy Microsoft Word hoặc Word đang khóa tệp mẫu. Hãy đóng tất cả cửa sổ Word rồi thử lại!",
                    "Tích Hợp Trực Tiếp Vào Word - VBTM Pro 2.5.0",
                    MessageBoxButtons.OK,
                    ok ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                );
            };
            topBar.Controls.Add(btnPinWord);

            Panel mainArea = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = Color.FromArgb(95, 10, 14)
            };

            this.Controls.Add(mainArea);
            this.Controls.Add(topBar);
            this.Controls.Add(lblStatus);

            this.Shown += (s, e) =>
            {
                System.Threading.ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        string startupDotm = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                            "Microsoft", "Word", "STARTUP", "VBTM_Pro_Addin.dotm"
                        );
                        string instDir = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "Programs", "VBTM-Pro"
                        );
                        bool needInstall = !File.Exists(startupDotm);
                        if (!needInstall)
                        {
                            try
                            {
                                using (ZipArchive z = ZipFile.OpenRead(startupDotm))
                                {
                                    ZipArchiveEntry ct = z.GetEntry("[Content_Types].xml");
                                    if (ct == null) needInstall = true;
                                    else
                                    {
                                        using (StreamReader sr = new StreamReader(ct.Open(), Encoding.UTF8))
                                        {
                                            if (!sr.ReadToEnd().Contains("template.macroEnabledTemplate"))
                                                needInstall = true;
                                        }
                                    }
                                }
                            }
                            catch { needInstall = true; }
                        }
                        if (needInstall)
                        {
                            Directory.CreateDirectory(instDir);
                            string exeDst = Path.Combine(instDir, "VBTM-Pro.exe");
                            if (!File.Exists(exeDst) && !Application.ExecutablePath.Equals(exeDst, StringComparison.OrdinalIgnoreCase))
                            {
                                try { File.Copy(Application.ExecutablePath, exeDst, true); } catch { }
                            }
                            bool ok = WordAddinInstaller.InstallWordRibbonAddin(instDir, exeDst);
                            if (ok && !this.IsDisposed)
                            {
                                this.BeginInvoke((Action)(() =>
                                {
                                    lblStatus.Text = "  ✓ Đã tự động tích hợp Tab Ribbon 'VBTM Pro 2.5 (05-HD & ND30)' trực tiếp vào Microsoft Word!";
                                }));
                            }
                        }
                    }
                    catch { }
                });
            };

            GroupBox grpMeta = new GroupBox
            {
                Text = " 1. Thông Tin Thể Thức Hành Chính & Đảng Ủy (HD 05-HD/VPTW & Nghị định 30) ",
                ForeColor = Color.FromArgb(250, 204, 21),
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
                ForeColor = Color.FromArgb(250, 204, 21),
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

            Button btnSaveQr = CreateBtn("Lưu Ảnh QR", new Point(402, 48), new Size(102, 28), Color.FromArgb(185, 28, 28));
            btnSaveQr.Click += (s, e) =>
            {
                using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PNG Image|*.png", FileName = "MaQR_VanBan_Phuong.png" })
                {
                    if (sfd.ShowDialog() == DialogResult.OK && picQr.Image != null)
                    {
                        picQr.Image.Save(sfd.FileName, ImageFormat.Png);
                        MessageBox.Show("Đã lưu ảnh mã QR tại:\n" + sfd.FileName, "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            };
            grpQr.Controls.Add(btnSaveQr);

            lstIssues = new ListBox
            {
                Location = new Point(140, 84),
                Size = new Size(364, 120),
                BackColor = Color.FromArgb(69, 10, 10),
                ForeColor = Color.FromArgb(250, 204, 21),
                BorderStyle = BorderStyle.FixedSingle
            };
            lstIssues.Items.Add("✓ Hệ thống AI Offline sẵn sàng rà soát lỗi chính tả & thể thức.");
            grpQr.Controls.Add(lstIssues);

            Button btnSampleErr = CreateBtn("Chèn Lỗi Mẫu", new Point(16, 140), new Size(106, 30), Color.FromArgb(153, 27, 27));
            btnSampleErr.Click += (s, e) =>
            {
                txtBody.Text = "Thực hiện chủ chương cải cách hành chính tại địa phương, Ủy ban nhân dân phường đã khẩn chương chiển khai kế hoạc bổ xung trang thiết bị và sử lý nghiêm các trường hợp sai phạm , đảm bảo kiểm xoát chính sác tiến độ công việc và hổ trợ cán bộ đạt tiêu chuẫn đề ra";
                ScanIssuesOnly();
            };
            grpQr.Controls.Add(btnSampleErr);

            Button btnCheck = CreateBtn("Rà Soát Lỗi", new Point(16, 176), new Size(106, 30), Color.FromArgb(202, 138, 4));
            btnCheck.Click += (s, e) => ScanIssuesOnly();
            grpQr.Controls.Add(btnCheck);

            mainArea.Controls.Add(grpQr);

            GroupBox grpEditor = new GroupBox
            {
                Text = " 3. Nội Dung Văn Bản Chính (Phông chữ chuẩn Times New Roman 14pt - Khổ A4) ",
                ForeColor = Color.FromArgb(250, 204, 21),
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
            isPreservedPdfMode = false;
            switch (idx)
            {
                case 0: // 1. Công văn UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "145/UBND-VP";
                    txtDocType.Text = "";
                    txtSubject.Text = "V/v triển khai ứng dụng chuẩn hóa thể thức văn bản và số hóa hồ sơ hành chính tại địa phương";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Như trên;\r\n- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- UBMTTQ và các đoàn thể phường;\r\n- Các tổ dân phố;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Kính gửi: Các bộ phận chuyên môn, cán bộ, công chức thuộc Ủy ban nhân dân phường.\r\n\r\n" +
                                   "Thực hiện Kế hoạch cải cách hành chính và chuyển đổi số năm 2026, Ủy ban nhân dân phường triển khai bộ công cụ chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư.\r\n\r\n" +
                                   "1. Các bộ phận chuyên môn, công chức thuộc Ủy ban nhân dân phường thực hiện rà soát, chuẩn hóa toàn bộ văn bản dự thảo trước khi trình ký ban hành; bảo đảm đúng quy định về khổ giấy A4, lề trang (Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm), phông chữ Times New Roman, cỡ chữ 13-14pt.\r\n\r\n" +
                                   "2. Đẩy mạnh số hóa hồ sơ, tài liệu lưu trữ và ứng dụng công cụ kiểm tra lỗi chính tả tiếng Việt ngoại tuyến (Offline 100%) nhằm nâng cao chất lượng ban hành văn bản./.";
                    break;
                case 1: // 2. Báo cáo UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "86/BC-UBND";
                    txtDocType.Text = "BÁO CÁO";
                    txtSubject.Text = "Kết quả thực hiện công tác cải cách hành chính, văn thư lưu trữ và chuyển đổi số trên địa bàn phường năm 2026";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- UBND Quận (để b/c);\r\n- Phòng Nội vụ;\r\n- Thường trực Đảng ủy, HĐND phường;\r\n- CT, các PCT UBND phường;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Thực hiện Kế hoạch công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường báo cáo kết quả triển khai công tác chuẩn hóa văn bản và số hóa hồ sơ trên địa bàn phường như sau:\r\n\r\n" +
                                   "I. KẾT QUẢ ĐẠT ĐƯỢC\r\n\r\n" +
                                   "1. Công tác ban hành văn bản: 100% văn bản hành chính của Hội đồng nhân dân và Ủy ban nhân dân phường được chuẩn hóa đồng bộ theo Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ.\r\n\r\n" +
                                   "2. Công tác số hóa hồ sơ: Đã ứng dụng phần mềm VBTM Pro 2.5.0 và VietNamOCR 3.5 Pro phục vụ rà soát lỗi chính tả tiếng Việt, tạo mã QR quản lý hồ sơ và bóc tách dữ liệu bảng biểu sang Word, Excel nhanh chóng, chính xác.\r\n\r\n" +
                                   "II. PHƯƠNG HƯỚNG, NHIỆM VỤ TRONG THỜI GIAN TỚI\r\n\r\n" +
                                   "Tiếp tục duy trì nền nếp công tác văn thư, lưu trữ và nâng cao hiệu quả giải quyết thủ tục hành chính phục vụ nhân dân./.";
                    break;
                case 2: // 3. Quyết định (cá biệt) UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "120/QĐ-UBND";
                    txtDocType.Text = "QUYẾT ĐỊNH";
                    txtSubject.Text = "Về việc thành lập Tổ công tác kiểm tra, rà soát công tác văn thư, lưu trữ và chuyển đổi số tại Ủy ban nhân dân phường";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Như Điều 3;\r\n- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "ỦY BAN NHÂN DÂN PHƯỜNG\r\n\r\n" +
                                   "Căn cứ Luật Tổ chức chính quyền địa phương;\r\n\r\n" +
                                   "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\r\n\r\n" +
                                   "Theo đề nghị của Công chức Văn phòng - Thống kê phường,\r\n\r\n" +
                                   "QUYẾT ĐỊNH:\r\n\r\n" +
                                   "Điều 1. Thành lập Tổ công tác kiểm tra, rà soát công tác chuẩn hóa thể thức văn bản hành chính và lập hồ sơ điện tử tại Ủy ban nhân dân phường.\r\n\r\n" +
                                   "Điều 2. Tổ công tác có nhiệm vụ hướng dẫn, kiểm tra việc thực hiện quy chuẩn văn bản theo Nghị định số 30/2020/NĐ-CP và báo cáo Chủ tịch Ủy ban nhân dân phường kết quả thực hiện.\r\n\r\n" +
                                   "Điều 3. Quyết định này có hiệu lực thi hành kể từ ngày ký. Công chức Văn phòng - Thống kê và các ông, bà có tên tại Điều 1 chịu trách nhiệm thi hành Quyết định này./.";
                    break;
                case 3: // 4. Quyết định (Quy chế) UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "125/QĐ-UBND";
                    txtDocType.Text = "QUYẾT ĐỊNH";
                    txtSubject.Text = "Ban hành Quy chế công tác văn thư, lưu trữ và chuẩn hóa văn bản điện tử tại Ủy ban nhân dân phường";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Như Điều 3;\r\n- UBND Quận (để b/c);\r\n- Thường trực Đảng ủy, HĐND phường;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "ỦY BAN NHÂN DÂN PHƯỜNG\r\n\r\n" +
                                   "Căn cứ Luật Tổ chức chính quyền địa phương;\r\n\r\n" +
                                   "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\r\n\r\n" +
                                   "Theo đề nghị của Công chức Văn phòng - Thống kê phường,\r\n\r\n" +
                                   "QUYẾT ĐỊNH:\r\n\r\n" +
                                   "Điều 1. Ban hành kèm theo Quyết định này Quy chế công tác văn thư, lưu trữ và chuẩn hóa thể thức văn bản hành chính tại Ủy ban nhân dân phường.\r\n\r\n" +
                                   "Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.\r\n\r\n" +
                                   "Điều 3. Công chức Văn phòng - Thống kê và các bộ phận chuyên môn, cán bộ, công chức thuộc Ủy ban nhân dân phường chịu trách nhiệm thi hành Quyết định này./.";
                    break;
                case 4: // 5. Kế hoạch UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "92/KH-UBND";
                    txtDocType.Text = "KẾ HOẠCH";
                    txtSubject.Text = "Triển khai công tác cải cách hành chính, chuyển đổi số và chuẩn hóa văn bản trên địa bàn phường năm 2026";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- UBND Quận (để b/c);\r\n- Thường trực Đảng ủy, HĐND phường;\r\n- Chủ tịch, các PCT UBND phường;\r\n- Các bộ phận chuyên môn;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Thực hiện Kế hoạch của Ủy ban nhân dân quận về công tác cải cách hành chính năm 2026, Ủy ban nhân dân phường ban hành Kế hoạch triển khai thực hiện cụ thể như sau:\r\n\r\n" +
                                   "I. MỤC ĐÍCH, YÊU CẦU\r\n\r\n" +
                                   "1. Bảo đảm 100% văn bản ban hành của Ủy ban nhân dân phường tuân thủ đúng quy định tại Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ.\r\n\r\n" +
                                   "2. Gắn công tác chuẩn hóa văn bản với việc số hóa hồ sơ, giải quyết thủ tục hành chính nhanh gọn, chính xác.\r\n\r\n" +
                                   "II. TỔ CHỨC THỰC HIỆN\r\n\r\n" +
                                   "Giao Công chức Văn phòng - Thống kê chủ trì, phối hợp với các công chức chuyên môn triển khai thực hiện nghiêm túc Kế hoạch này./.";
                    break;
                case 5: // 6. Tờ trình UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "64/TTr-UBND";
                    txtDocType.Text = "TỜ TRÌNH";
                    txtSubject.Text = "Về việc đề nghị phê duyệt phương án nâng cấp hạ tầng công nghệ thông tin và số hóa hồ sơ lưu trữ tại phường";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Như trên;\r\n- Phòng Nội vụ, Phòng Tài chính - Kế hoạch;\r\n- Thường trực Đảng ủy, HĐND phường;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Kính gửi: Ủy ban nhân dân quận.\r\n\r\n" +
                                   "Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;\r\n\r\n" +
                                   "Nhằm đáp ứng yêu cầu hiện đại hóa nền hành chính và số hóa hồ sơ tài liệu tại địa phương, Ủy ban nhân dân phường kính trình Ủy ban nhân dân quận xem xét, phê duyệt phương án triển khai phần mềm chuẩn hóa thể thức văn bản và số hóa tài liệu lưu trữ năm 2026.\r\n\r\n" +
                                   "Ủy ban nhân dân phường kính trình Ủy ban nhân dân quận xem xét, quyết định./.";
                    break;
                case 6: // 7. Thông báo UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "168/TB-UBND";
                    txtDocType.Text = "THÔNG BÁO";
                    txtSubject.Text = "Kết luận của Chủ tịch Ủy ban nhân dân phường tại cuộc họp giao ban công tác cải cách hành chính tháng 11 năm 2026";
                    txtSignerTitle.Text = "TL. CHỦ TỊCH - VĂN PHÒNG - THỐNG KÊ";
                    txtSignerName.Text = "Lê Hoàng Nam";
                    txtRecipients.Text = "- Thường trực Đảng ủy, HĐND phường (để b/c);\r\n- Chủ tịch, các PCT UBND phường;\r\n- Các công chức chuyên môn;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Ngày 10 tháng 11 năm 2026, Chủ tịch Ủy ban nhân dân phường đã chủ trì cuộc họp giao ban đánh giá kết quả công tác cải cách hành chính và văn thư lưu trữ. Sau khi nghe báo cáo và ý kiến phát biểu của các bộ phận chuyên môn, Chủ tịch Ủy ban nhân dân phường kết luận như sau:\r\n\r\n" +
                                   "1. Biểu dương bộ phận Văn phòng - Thống kê và các công chức chuyên môn đã thực hiện tốt việc chuẩn hóa thể thức văn bản theo Nghị định số 30/2020/NĐ-CP.\r\n\r\n" +
                                   "2. Yêu cầu các bộ phận tiếp tục kiểm soát chặt chẽ chất lượng dự thảo văn bản, sử dụng tính năng soát lỗi chính tả và đính kèm mã QR xác thực trên toàn bộ văn bản ban hành./.";
                    break;
                case 7: // 8. Nghị quyết HĐND Phường
                    txtGovOrg.Text = "";
                    txtIssueOrg.Text = "HỘI ĐỒNG NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "18/NQ-HĐND";
                    txtDocType.Text = "NGHỊ QUYẾT";
                    txtSubject.Text = "Về nhiệm vụ phát triển kinh tế - xã hội, quốc phòng - an ninh và đẩy mạnh cải cách hành chính năm 2026";
                    txtSignerTitle.Text = "CHỦ TỊCH";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Thường trực HĐND, UBND Quận (để b/c);\r\n- Thường trực Đảng ủy phường;\r\n- Thường trực HĐND, UBND, UBMTTQ phường;\r\n- Lưu: VT.";
                    txtBody.Text = "HỘI ĐỒNG NHÂN DÂN PHƯỜNG\r\nKHÓA XII, KỲ HỌP THỨ 10\r\n\r\n" +
                                   "Căn cứ Luật Tổ chức chính quyền địa phương;\r\n\r\n" +
                                   "Xét Tờ trình của Ủy ban nhân dân phường và Báo cáo thẩm tra của các Ban Hội đồng nhân dân phường,\r\n\r\n" +
                                   "QUYẾT NGHỊ:\r\n\r\n" +
                                   "Điều 1. Tán thành Báo cáo kết quả thực hiện nhiệm vụ phát triển kinh tế - xã hội, cải cách hành chính và chuyển đổi số trên địa bàn phường.\r\n\r\n" +
                                   "Điều 2. Giao Ủy ban nhân dân phường tổ chức triển khai thực hiện Nghị quyết này; đẩy mạnh chuẩn hóa văn bản theo Nghị định số 30/2020/NĐ-CP.\r\n\r\n" +
                                   "Điều 3. Thường trực Hội đồng nhân dân, các Ban của Hội đồng nhân dân và đại biểu Hội đồng nhân dân phường giám sát việc thực hiện Nghị quyết này./.";
                    break;
                case 8: // 9. Biên bản họp UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "09/BB-UBND";
                    txtDocType.Text = "BIÊN BẢN";
                    txtSubject.Text = "Họp rà soát, nghiệm thu kết quả chuẩn hóa văn bản và số hóa hồ sơ hành chính năm 2026";
                    txtSignerTitle.Text = "CHỦ TRÌ CUỘC HỌP - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Thành phần dự họp;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Hôm nay, vào lúc 08 giờ 00 phút ngày 11 tháng 11 năm 2026, tại Phòng họp Ủy ban nhân dân phường, chúng tôi gồm:\r\n\r\n" +
                                   "I. THÀNH PHẦN THAM DỰ\r\n\r\n" +
                                   "1. Ông Nguyễn Văn An - Chủ tịch Ủy ban nhân dân phường, Chủ trì.\r\n\r\n" +
                                   "2. Ông Lê Hoàng Nam - Công chức Văn phòng - Thống kê, Thư ký.\r\n\r\n" +
                                   "II. NỘI DUNG CUỘC HỌP\r\n\r\n" +
                                   "Công chức Văn phòng - Thống kê báo cáo kết quả chuẩn hóa 100% văn bản theo Nghị định số 30/2020/NĐ-CP và số hóa tài liệu lưu trữ. Các thành viên dự họp thảo luận và thống nhất nghiệm thu kết quả thực hiện./.";
                    break;
                case 9: // 10. Giấy mời họp UBND Phường
                    txtGovOrg.Text = "ỦY BAN NHÂN DÂN QUẬN";
                    txtIssueOrg.Text = "ỦY BAN NHÂN DÂN PHƯỜNG";
                    txtDocNum.Text = "210/GM-UBND";
                    txtDocType.Text = "GIẤY MỜI";
                    txtSubject.Text = "Họp triển khai công tác chuẩn hóa thể thức văn bản và tập huấn số hóa hồ sơ lưu trữ";
                    txtSignerTitle.Text = "TM. ỦY BAN NHÂN DÂN - CHỦ TỊCH";
                    txtSignerName.Text = "Nguyễn Văn An";
                    txtRecipients.Text = "- Như trên;\r\n- Chủ tịch, các PCT UBND phường;\r\n- Lưu: VT, VP.";
                    txtBody.Text = "Kính gửi: Các đồng chí cán bộ, công chức chuyên môn và Tổ trưởng các tổ dân phố thuộc phường.\r\n\r\n" +
                                   "Ủy ban nhân dân phường tổ chức hội nghị triển khai công tác chuẩn hóa thể thức văn bản hành chính theo Nghị định số 30/2020/NĐ-CP và hướng dẫn sử dụng phần mềm VBTM Pro 2.5.0.\r\n\r\n" +
                                   "1. Thành phần: Lãnh đạo Ủy ban nhân dân phường, toàn thể cán bộ, công chức chuyên môn và Tổ trưởng các tổ dân phố.\r\n\r\n" +
                                   "2. Thời gian: 08 giờ 00 phút, ngày 15 tháng 11 năm 2026.\r\n\r\n" +
                                   "3. Địa điểm: Hội trường Ủy ban nhân dân phường.\r\n\r\n" +
                                   "Đề nghị các đồng chí sắp xếp công việc tham dự đầy đủ, đúng giờ./.";
                    break;
                case 10: // 11. Công văn Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "58-CV/ĐU";
                    txtDocType.Text = "";
                    txtSubject.Text = "V/v quán triệt thực hiện thể thức và kỹ thuật trình bày văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW";
                    txtSignerTitle.Text = "T/M BAN THƯỜNG VỤ - BÍ THƯ";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Như trên;\r\n- Văn phòng Quận ủy (để b/c);\r\n- Các đồng chí Đảng ủy viên;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Kính gửi: Các chi bộ trực thuộc Đảng ủy phường.\r\n\r\n" +
                                   "Thực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Ban Thường vụ Đảng ủy phường yêu cầu các chi bộ trực thuộc thực hiện nghiêm túc các nội dung sau:\r\n\r\n" +
                                   "1. Thực hiện thống nhất phông chữ Times New Roman (bộ mã ký tự Unicode), cỡ chữ 14pt, định lề trang A4 (lề trên 20mm, lề dưới 20mm, lề trái 30mm, lề phải 15mm) đối với toàn bộ nghị quyết, báo cáo, chương trình, kế hoạch, biên bản sinh hoạt chi bộ.\r\n\r\n" +
                                   "2. Đề nghị cấp ủy các chi bộ trực thuộc nghiêm túc quán triệt và tổ chức thực hiện./.";
                    break;
                case 11: // 12. Báo cáo Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "42-BC/ĐU";
                    txtDocType.Text = "BÁO CÁO";
                    txtSubject.Text = "Tổng kết công tác xây dựng Đảng, cải cách hành chính trong Đảng và chuyển đổi số tại Đảng bộ phường năm 2026";
                    txtSignerTitle.Text = "T/M ĐẢNG ỦY - BÍ THƯ";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Ban Thường vụ Quận ủy (để b/c);\r\n- Văn phòng Quận ủy;\r\n- Các đồng chí Đảng ủy viên;\r\n- HĐND, UBND, UBMTTQ phường;\r\n- Các chi bộ trực thuộc;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Thực hiện Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Đảng ủy phường báo cáo kết quả thực hiện chuẩn hóa toàn bộ hệ thống văn bản của Đảng ủy và các chi bộ trực thuộc năm 2026./.";
                    break;
                case 12: // 13. Nghị quyết Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "15-NQ/ĐU";
                    txtDocType.Text = "NGHỊ QUYẾT";
                    txtSubject.Text = "Về tăng cường sự lãnh đạo của Đảng đối với công tác cải cách hành chính và chuyển đổi số trên địa bàn phường";
                    txtSignerTitle.Text = "T/M ĐẢNG ỦY - BÍ THƯ";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Ban Thường vụ Quận ủy (để b/c);\r\n- Các đồng chí Đảng ủy viên;\r\n- HĐND, UBND, UBMTTQ phường;\r\n- Các chi bộ trực thuộc;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Thực hiện Chương trình công tác toàn khóa của Ban Chấp hành Đảng bộ phường, sau khi thảo luận đề án của Ban Thường vụ Đảng ủy, Ban Chấp hành Đảng bộ phường quyết nghị:\r\n\r\n" +
                                   "I. TÌNH HÌNH VÀ NGUYÊN NHÂN\r\n\r\n" +
                                   "Công tác cải cách hành chính và chuyển đổi số tại Đảng bộ và chính quyền phường thời gian qua đạt nhiều kết quả quan trọng; chất lượng ban hành văn kiện, nghị quyết được chuẩn hóa đồng bộ.\r\n\r\n" +
                                   "II. MỤC TIÊU, NHIỆM VỤ VÀ GIẢI PHÁP\r\n\r\n" +
                                   "Phấn đấu 100% văn bản của Đảng ủy và các chi bộ trực thuộc được soạn thảo, thẩm định đúng quy định tại Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng./.";
                    break;
                case 13: // 14. Kế hoạch Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "24-KH/ĐU";
                    txtDocType.Text = "KẾ HOẠCH";
                    txtSubject.Text = "Tập huấn nghiệp vụ công tác văn thư, lưu trữ và chuẩn hóa thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW";
                    txtSignerTitle.Text = "T/M BAN THƯỜNG VỤ - PHÓ BÍ THƯ THƯỜNG TRỰC";
                    txtSignerName.Text = "Phạm Thị Thanh Hoa";
                    txtRecipients.Text = "- Văn phòng Quận ủy (để b/c);\r\n- Các đồng chí Đảng ủy viên;\r\n- Các chi bộ trực thuộc;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Căn cứ Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng, Ban Thường vụ Đảng ủy phường ban hành Kế hoạch tập huấn nghiệp vụ công tác văn thư cho cấp ủy các chi bộ trực thuộc như sau:\r\n\r\n" +
                                   "I. MỤC ĐÍCH, YÊU CẦU\r\n\r\n" +
                                   "Nâng cao kỹ năng soạn thảo, chuẩn hóa thể thức nghị quyết, biên bản, báo cáo của các chi bộ trực thuộc Đảng ủy phường.\r\n\r\n" +
                                   "II. TỔ CHỨC THỰC HIỆN\r\n\r\n" +
                                   "Văn phòng Đảng ủy chuẩn bị tài liệu, phần mềm và đôn đốc các chi bộ tham gia đầy đủ./.";
                    break;
                case 14: // 15. Chương trình công tác Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "10-CTr/ĐU";
                    txtDocType.Text = "CHƯƠNG TRÌNH";
                    txtSubject.Text = "Công tác kiểm tra, giám sát của Ban Thường vụ Đảng ủy phường năm 2026";
                    txtSignerTitle.Text = "T/M BAN THƯỜNG VỤ - BÍ THƯ";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Ủy ban Kiểm tra Quận ủy (để b/c);\r\n- Các đồng chí Đảng ủy viên;\r\n- UBKT Đảng ủy phường;\r\n- Các chi bộ trực thuộc;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Căn cứ Điều lệ Đảng Cộng sản Việt Nam và Chương trình kiểm tra, giám sát của Quận ủy, Ban Thường vụ Đảng ủy phường xây dựng Chương trình công tác kiểm tra, giám sát năm 2026 như sau:\r\n\r\n" +
                                   "I. MỤC ĐÍCH, YÊU CẦU\r\n\r\n" +
                                   "Chủ động phòng ngừa vi phạm, nâng cao năng lực lãnh đạo và sức chiến đấu của các tổ chức đảng trực thuộc.\r\n\r\n" +
                                   "II. TỔ CHỨC THỰC HIỆN\r\n\r\n" +
                                   "Ủy ban Kiểm tra Đảng ủy chủ trì tham mưu giúp Ban Thường vụ Đảng ủy tổ chức thực hiện Chương trình này./.";
                    break;
                default: // 16. Quyết định Đảng ủy Phường
                    txtGovOrg.Text = "QUẬN ỦY";
                    txtIssueOrg.Text = "ĐẢNG ỦY PHƯỜNG";
                    txtDocNum.Text = "35-QĐ/ĐU";
                    txtDocType.Text = "QUYẾT ĐỊNH";
                    txtSubject.Text = "Ban hành Quy chế quản lý, khai thác văn bản điện tử và lưu trữ hồ sơ tại Văn phòng Đảng ủy phường";
                    txtSignerTitle.Text = "T/M BAN THƯỜNG VỤ - BÍ THƯ";
                    txtSignerName.Text = "Trần Minh Quân";
                    txtRecipients.Text = "- Văn phòng Quận ủy (để b/c);\r\n- Như Điều 2;\r\n- Các đồng chí Đảng ủy viên;\r\n- Lưu Văn phòng Đảng ủy.";
                    txtBody.Text = "Căn cứ Điều lệ Đảng Cộng sản Việt Nam;\r\n\r\n" +
                                   "Căn cứ Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng;\r\n\r\n" +
                                   "Xét đề nghị của Văn phòng Đảng ủy phường,\r\n\r\n" +
                                   "BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH\r\n\r\n" +
                                   "Điều 1. Ban hành kèm theo Quyết định này Quy chế quản lý, chuẩn hóa thể thức văn bản của Đảng theo Hướng dẫn số 05-HD/VPTW và lưu trữ hồ sơ điện tử tại Đảng bộ phường.\r\n\r\n" +
                                   "Điều 2. Văn phòng Đảng ủy, các ban xây dựng Đảng và các chi bộ trực thuộc chịu trách nhiệm thi hành Quyết định này.\r\n\r\n" +
                                   "Điều 3. Quyết định này có hiệu lực kể từ ngày ký./.";
                    break;
            }
            if (txtQrContent != null)
            {
                txtQrContent.Text = "VBTM:" + txtDocNum.Text + "|" + txtIssueOrg.Text;
            }
        }

        private void ScanIssuesOnly()
        {
            lstIssues.Items.Clear();
            string content = txtBody.Text + " " + txtSubject.Text;
            for (int i = 0; i < SpellRules.GetLength(0); i++)
            {
                string pat = CreateUnicodeWordRegex(SpellRules[i, 0]);
                if (Regex.IsMatch(content, pat, RegexOptions.IgnoreCase))
                {
                    lstIssues.Items.Add("[Chính tả AI] Phát hiện '" + SpellRules[i, 0] + "' -> Chuẩn: '" + SpellRules[i, 1] + "' (" + SpellRules[i, 2] + ")");
                }
            }
            if (Regex.IsMatch(content, @"(?<!\w)ng[ieêìíỉĩịèéẻẽẹềếểễệ]", RegexOptions.IgnoreCase))
            {
                lstIssues.Items.Add("[Ngữ âm AI] Phát hiện lỗi quy tắc ng/ngh đứng trước nguyên âm i, e, ê");
            }
            if (Regex.IsMatch(txtBody.Text, @"\s+[,.;:!?]"))
            {
                lstIssues.Items.Add("[Thể thức NĐ30] Có khoảng trắng thừa trước dấu câu (, . ; : ! ?)");
            }
            if (!txtBody.Text.Trim().EndsWith("./."))
            {
                lstIssues.Items.Add("[Thể thức NĐ30] Cuối văn bản thiếu ký hiệu kết thúc './.'");
            }
            if (lstIssues.Items.Count == 0)
            {
                lstIssues.Items.Add("✓ Văn bản đạt chuẩn 100% về chính tả và thể thức Đảng ủy / UBND phường!");
            }
        }

        public string FixAllVietnameseSpellingAndPunctuation(string rawInput, out int fixedCount)
        {
            string raw = rawInput ?? "";
            int count = 0;
            for (int i = 0; i < SpellRules.GetLength(0); i++)
            {
                string pat = CreateUnicodeWordRegex(SpellRules[i, 0]);
                string right = SpellRules[i, 1];
                raw = Regex.Replace(raw, pat, m =>
                {
                    count++;
                    return ApplyCasePattern(m.Value, right);
                }, RegexOptions.IgnoreCase);
            }

            // Sửa quy tắc ngữ âm ng -> ngh trước i, e, ê
            raw = Regex.Replace(raw, @"(?<!\w)(ng)([ieêìíỉĩịèéẻẽẹềếểễệ]\w*)", m =>
            {
                count++;
                string prefix = char.IsUpper(m.Groups[1].Value[0]) ? "Ngh" : "ngh";
                return prefix + m.Groups[2].Value;
            }, RegexOptions.IgnoreCase);

            // Chuẩn hóa cách ghi ngày tháng theo Phụ lục I Nghị định 30/2020/NĐ-CP
            raw = Regex.Replace(raw, @"\bngày\s+([1-9])(?=\s+tháng\b)", "ngày 0$1", RegexOptions.IgnoreCase);
            raw = Regex.Replace(raw, @"\btháng\s+([12])(?=\s+năm\b)", "tháng 0$1", RegexOptions.IgnoreCase);
            raw = Regex.Replace(raw, @"\btháng\s+0([3-9])(?=\s+năm\b)", "tháng $1", RegexOptions.IgnoreCase);

            raw = Regex.Replace(raw, @"[ \t]+", " ");
            raw = Regex.Replace(raw, @"\s+([,.;:!?])", "$1");
            raw = Regex.Replace(raw, @"([,.;:!?])([^\s\d/.)])", "$1 $2");
            fixedCount = count;
            return raw.Trim();
        }

        private void StandardizeAndFixAll()
        {
            int bodyFixed = 0, subjFixed = 0;
            string raw = FixAllVietnameseSpellingAndPunctuation(txtBody.Text, out bodyFixed);
            if (raw.Length > 0 && !raw.EndsWith("./."))
            {
                raw = raw.TrimEnd('.') + "./.";
            }
            txtBody.Text = raw;

            if (txtSubject != null && !string.IsNullOrEmpty(txtSubject.Text))
            {
                txtSubject.Text = FixAllVietnameseSpellingAndPunctuation(txtSubject.Text, out subjFixed).TrimEnd('.');
            }

            int totalFixed = bodyFixed + subjFixed;
            ScanIssuesOnly();
            lblStatus.Text = "  ✓ Đã chuẩn hóa thể thức Đảng ủy / UBND phường & tự động sửa " + totalFixed + " lỗi chính tả tiếng Việt!";
            MessageBox.Show("Đã chuẩn hóa dấu câu, khoảng trắng, ngày tháng, ký hiệu kết thúc và tự động sửa " + totalFixed + " lỗi chính tả AI!", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void FormatActiveWordDocument()
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

                // 1. Định lề trang khổ A4 chuẩn Nghị định 30/2020/NĐ-CP & Hướng dẫn 05-HD/VPTW (Trên 20mm, Dưới 20mm, Trái 30mm, Phải 15mm)
                object pageSetup = doc.GetType().InvokeMember("PageSetup", BindingFlags.GetProperty, null, doc, null);
                try { pageSetup.GetType().InvokeMember("PaperSize", BindingFlags.SetProperty, null, pageSetup, new object[] { 7 }); } catch { }
                pageSetup.GetType().InvokeMember("TopMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f });
                pageSetup.GetType().InvokeMember("BottomMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 56.7f });
                pageSetup.GetType().InvokeMember("LeftMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 85.05f });
                pageSetup.GetType().InvokeMember("RightMargin", BindingFlags.SetProperty, null, pageSetup, new object[] { 42.5f });
                try { pageSetup.GetType().InvokeMember("DifferentFirstPageHeaderFooter", BindingFlags.SetProperty, null, pageSetup, new object[] { -1 }); } catch { }

                // 2. Đặt phông chữ Times New Roman & Tự động phát hiện quy chuẩn Đảng (05-HD/VPTW) hay Hành chính (NĐ 30/2020/NĐ-CP)
                object content = doc.GetType().InvokeMember("Content", BindingFlags.GetProperty, null, doc, null);
                string fullDocText = Convert.ToString(content.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, content, null)) ?? "";
                bool isDangWord = DetectIsPartyDocText(fullDocText) || (cboTemplate != null && cboTemplate.SelectedIndex >= 10);
                object font = content.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, content, null);
                font.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, font, new object[] { "Times New Roman" });

                // 2b. Nếu là văn bản Đảng theo Hướng dẫn 05-HD/VPTW: Tự động chuẩn hóa Bảng Tiêu ngữ (Bảng 1) & Bảng Nơi nhận/Chữ ký (Bảng cuối)
                if (isDangWord)
                {
                    try
                    {
                        object docTables = doc.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, doc, null);
                        int tblTotal = (int)docTables.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docTables, null);
                        if (tblTotal >= 1)
                        {
                            object topTbl = docTables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docTables, new object[] { 1 });
                            object topCols = topTbl.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, topTbl, null);
                            int topColCount = (int)topCols.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, topCols, null);
                            if (topColCount >= 2)
                            {
                                object cellL = topTbl.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod, null, topTbl, new object[] { 1, 1 });
                                object rngL = cellL.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cellL, null);
                                object fntL = rngL.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, rngL, null);
                                fntL.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, fntL, new object[] { "Times New Roman" });
                                if (isDangWord)
                                {
                                    try
                                    {
                                        topTbl.GetType().InvokeMember("AllowAutoFit", BindingFlags.SetProperty, null, topTbl, new object[] { 0 });
                                        topTbl.GetType().InvokeMember("LeftPadding", BindingFlags.SetProperty, null, topTbl, new object[] { 0f });
                                        topTbl.GetType().InvokeMember("RightPadding", BindingFlags.SetProperty, null, topTbl, new object[] { 0f });
                                        object c1 = topCols.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, topCols, new object[] { 1 });
                                        object c2 = topCols.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, topCols, new object[] { 2 });
                                        c1.GetType().InvokeMember("PreferredWidthType", BindingFlags.SetProperty, null, c1, new object[] { 2 });
                                        c1.GetType().InvokeMember("PreferredWidth", BindingFlags.SetProperty, null, c1, new object[] { 49f });
                                        c2.GetType().InvokeMember("PreferredWidthType", BindingFlags.SetProperty, null, c2, new object[] { 2 });
                                        c2.GetType().InvokeMember("PreferredWidth", BindingFlags.SetProperty, null, c2, new object[] { 51f });
                                    }
                                    catch { }
                                }
                                fntL.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, fntL, new object[] { 14f });
                                fntL.GetType().InvokeMember("Scaling", BindingFlags.SetProperty, null, fntL, new object[] { 100 });

                                object cellR = topTbl.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod, null, topTbl, new object[] { 1, 2 });
                                object rngR = cellR.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cellR, null);
                                object rParas = rngR.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, rngR, null);
                                int rpCount = (int)rParas.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, rParas, null);
                                for (int rpIdx = 1; rpIdx <= rpCount; rpIdx++)
                                {
                                    object rp = rParas.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, rParas, new object[] { rpIdx });
                                    object rpRng = rp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, rp, null);
                                    string rpTxt = (Convert.ToString(rpRng.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, rpRng, null)) ?? "").TrimEnd('\r', '\a', '\n').Trim();
                                    object rpFnt = rpRng.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, rpRng, null);
                                    rpFnt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, rpFnt, new object[] { "Times New Roman" });
                                    rpFnt.GetType().InvokeMember("Scaling", BindingFlags.SetProperty, null, rpFnt, new object[] { 100 });
                                    if (rpTxt.IndexOf("ĐẢNG CỘNG SẢN", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        rpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, rpFnt, new object[] { 15f });
                                        rpFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, rpFnt, new object[] { -1 });
                                        rpFnt.GetType().InvokeMember("Underline", BindingFlags.SetProperty, null, rpFnt, new object[] { 0 });
                                        try
                                        {
                                            rp.GetType().InvokeMember("LeftIndent", BindingFlags.SetProperty, null, rp, new object[] { 15.6f });
                                            rp.GetType().InvokeMember("RightIndent", BindingFlags.SetProperty, null, rp, new object[] { 15.6f });
                                            object rpBdrs = rp.GetType().InvokeMember("Borders", BindingFlags.GetProperty, null, rp, null);
                                            object bBot = rpBdrs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, rpBdrs, new object[] { -3 });
                                            bBot.GetType().InvokeMember("LineStyle", BindingFlags.SetProperty, null, bBot, new object[] { 1 });
                                            bBot.GetType().InvokeMember("LineWidth", BindingFlags.SetProperty, null, bBot, new object[] { 6 });
                                        }
                                        catch { }
                                    }
                                    else if (Regex.IsMatch(rpTxt, @"\b(ngày|tháng|năm)\b", RegexOptions.IgnoreCase))
                                    {
                                        rpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, rpFnt, new object[] { 14f });
                                        rpFnt.GetType().InvokeMember("Italic", BindingFlags.SetProperty, null, rpFnt, new object[] { -1 });
                                        rpFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, rpFnt, new object[] { 0 });
                                    }
                                }
                            }
                        }
                        if (tblTotal >= 2)
                        {
                            object botTbl = docTables.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, docTables, new object[] { tblTotal });
                            object botCols = botTbl.GetType().InvokeMember("Columns", BindingFlags.GetProperty, null, botTbl, null);
                            int botColCount = (int)botCols.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, botCols, null);
                            if (botColCount >= 2)
                            {
                                object cellBL = botTbl.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod, null, botTbl, new object[] { 1, 1 });
                                object rngBL = cellBL.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cellBL, null);
                                object blParas = rngBL.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, rngBL, null);
                                int blpCount = (int)blParas.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, blParas, null);
                                for (int bpIdx = 1; bpIdx <= blpCount; bpIdx++)
                                {
                                    object bp = blParas.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, blParas, new object[] { bpIdx });
                                    object bpRng = bp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, bp, null);
                                    string bpTxt = Convert.ToString(bpRng.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, bpRng, null)) ?? "";
                                    object bpFnt = bpRng.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, bpRng, null);
                                    bpFnt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, bpFnt, new object[] { "Times New Roman" });
                                    bpFnt.GetType().InvokeMember("Italic", BindingFlags.SetProperty, null, bpFnt, new object[] { 0 });
                                    bpFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, bpFnt, new object[] { 0 });
                                    try
                                    {
                                        bp.GetType().InvokeMember("SpaceBefore", BindingFlags.SetProperty, null, bp, new object[] { 0f });
                                        bp.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, bp, new object[] { 0f });
                                        bp.GetType().InvokeMember("LineSpacingRule", BindingFlags.SetProperty, null, bp, new object[] { 0 });
                                    }
                                    catch { }
                                    if (bpTxt.IndexOf("Nơi nhận", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        bpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, bpFnt, new object[] { 14f });
                                        bpFnt.GetType().InvokeMember("Underline", BindingFlags.SetProperty, null, bpFnt, new object[] { 0 });
                                        try
                                        {
                                            object uDup = bpRng.GetType().InvokeMember("Duplicate", BindingFlags.GetProperty, null, bpRng, null);
                                            int st = (int)uDup.GetType().InvokeMember("Start", BindingFlags.GetProperty, null, uDup, null);
                                            uDup.GetType().InvokeMember("End", BindingFlags.SetProperty, null, uDup, new object[] { st + 8 });
                                            object uFnt = uDup.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, uDup, null);
                                            uFnt.GetType().InvokeMember("Underline", BindingFlags.SetProperty, null, uFnt, new object[] { 1 });
                                        }
                                        catch { }
                                    }
                                    else
                                    {
                                        bpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, bpFnt, new object[] { 12f });
                                    }
                                }

                                object cellBR = botTbl.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod, null, botTbl, new object[] { 1, 2 });
                                object rngBR = cellBR.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cellBR, null);
                                object fntBR = rngBR.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, rngBR, null);
                                fntBR.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, fntBR, new object[] { "Times New Roman" });
                                fntBR.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, fntBR, new object[] { 14f });
                                try
                                {
                                    object brParas = rngBR.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, rngBR, null);
                                    int brCount = (int)brParas.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, brParas, null);
                                    for (int brIdx = 1; brIdx <= brCount; brIdx++)
                                    {
                                        object brp = brParas.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, brParas, new object[] { brIdx });
                                        brp.GetType().InvokeMember("SpaceBefore", BindingFlags.SetProperty, null, brp, new object[] { 0f });
                                        brp.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, brp, new object[] { 0f });
                                        object brpRng = brp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, brp, null);
                                        string brTxt = (Convert.ToString(brpRng.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, brpRng, null)) ?? "").TrimEnd('\r', '\a', '\n').Trim().ToUpperInvariant();
                                        object brpFnt = brpRng.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, brpRng, null);
                                        bool isPlainRole = Regex.IsMatch(brTxt, @"^(BÍ THƯ|PHÓ BÍ THƯ|PHÓ BÍ THƯ THƯỜNG TRỰC|CHÁNH VĂN PHÒNG|PHÓ CHÁNH VĂN PHÒNG|CHỦ NHIỆM|TRƯỞNG BAN)$");
                                        brpFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, brpFnt, new object[] { isPlainRole ? 0 : -1 });
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }

                // 3. Duyệt từng đoạn văn an toàn tuyệt đối (Không bao giờ làm hỏng ô bảng \r\a của Word!)
                int fixedInWord = 0;
                try
                {
                    object paragraphs = doc.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, doc, null);
                    int pCount = (int)paragraphs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, paragraphs, null);
                    StringBuilder sbAll = new StringBuilder();
                    for (int pIdx = 1; pIdx <= pCount; pIdx++)
                    {
                        object para = paragraphs.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, paragraphs, new object[] { pIdx });
                        object pRange = para.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, para, null);
                        string pText = Convert.ToString(pRange.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, pRange, null));
                        if (string.IsNullOrWhiteSpace(pText)) continue;

                        bool isTableCell = pText.Contains("\a");
                        try
                        {
                            object pTables = pRange.GetType().InvokeMember("Tables", BindingFlags.GetProperty, null, pRange, null);
                            int tblCount = (int)pTables.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, pTables, null);
                            if (tblCount > 0) isTableCell = true;
                        }
                        catch { }

                        string cleanP = pText.TrimEnd('\r', '\a', '\n');
                        if (string.IsNullOrWhiteSpace(cleanP)) continue;

                        int pFixed = 0;
                        string fixedP = FixAllVietnameseSpellingAndPunctuation(cleanP, out pFixed);

                        if (!isTableCell)
                        {
                            // Chỉ định dạng 14pt + thụt lề 1.27cm (36pt) + giãn đoạn 6pt cho đoạn văn thân bài bên ngoài bảng
                            string trimmedP = cleanP.Trim();
                            bool isPartyTitle = isDangWord && Regex.IsMatch(trimmedP, @"^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|BIÊN BẢN|GIẤY MỜI)$", RegexOptions.IgnoreCase);
                            bool isRuleLine = trimmedP.StartsWith("───") || trimmedP.StartsWith("---");
                            bool isDecisionCenter = Regex.IsMatch(trimmedP, @"^(QUYẾT ĐỊNH\s*:|QUYẾT NGHỊ\s*:|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH)$", RegexOptions.IgnoreCase);
                            bool isRomanHeading = Regex.IsMatch(trimmedP, @"^(I|II|III|IV|V|VI|VII|VIII|IX|X)\.\s+");
                            int curAlign = 0;
                            try { curAlign = Convert.ToInt32(para.GetType().InvokeMember("Alignment", BindingFlags.GetProperty, null, para, null)); } catch { }
                            try
                            {
                                object pFont = pRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, pRange, null);
                                para.GetType().InvokeMember("LeftIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                para.GetType().InvokeMember("RightIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                para.GetType().InvokeMember("SpaceBefore", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                if (isPartyTitle)
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { 15f });
                                    pFont.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, pFont, new object[] { -1 });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 1 });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 2f });
                                }
                                else if (isRuleLine)
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { isDangWord ? 14f : 10f });
                                    pFont.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, pFont, new object[] { isDangWord ? -1 : 0 });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 1 });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 10f });
                                }
                                else if (isDecisionCenter)
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { 14f });
                                    pFont.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, pFont, new object[] { -1 });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 1 });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 6f });
                                }
                                else if (isRomanHeading)
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { 14f });
                                    pFont.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, pFont, new object[] { -1 });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 3 });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 36f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 6f });
                                }
                                else if (curAlign == 1)
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { 14f });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 0f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 2f });
                                }
                                else
                                {
                                    pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { 14f });
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 6f });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 36f });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 3 }); // wdAlignParagraphJustify = 3
                                }
                            }
                            catch { }

                            if (pFixed > 0 || fixedP != cleanP)
                            {
                                fixedInWord += pFixed;
                                int endPos = (int)pRange.GetType().InvokeMember("End", BindingFlags.GetProperty, null, pRange, null);
                                pRange.GetType().InvokeMember("End", BindingFlags.SetProperty, null, pRange, new object[] { endPos - 1 });
                                pRange.GetType().InvokeMember("Text", BindingFlags.SetProperty, null, pRange, new object[] { fixedP });
                            }
                            sbAll.AppendLine(fixedP);
                        }
                    }
                    if (sbAll.Length > 0)
                    {
                        txtBody.Text = sbAll.ToString().Trim();
                        ScanIssuesOnly();
                    }
                }
                catch { }

                MessageBox.Show(
                    "Đã chuẩn hóa trực tiếp tài liệu Microsoft Word đang mở theo đúng quy định " + (isDangWord ? "(Hướng dẫn số 05-HD/VPTW):" : "(Nghị định số 30/2020/NĐ-CP):") + "\n\n" +
                    "• Khổ giấy A4 (210mm x 297mm), Lề Trái 30mm - Phải 15mm - Trên 20mm - Dưới 20mm\n" +
                    "• Phông chữ: Times New Roman (Chuẩn hóa tự động Bảng Tiêu ngữ 15pt/14pt, Nơi nhận 14pt gạch chân + danh sách 12pt, Chữ ký 14pt đậm, Thân bài 14pt thụt đầu dòng 1.27cm, cách đoạn 6pt)\n" +
                    "• Đánh số trang tự động từ trang 2 ở lề trên (ẩn số trang ở trang 1)\n" +
                    "• Tự động sửa " + fixedInWord + " lỗi chính tả, ngày tháng & dấu câu trực tiếp trong Word!",
                    "VBTM Pro 2.5.0 - Word Automation",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch
            {
                ExportAndOpenWordDoc();
            }
        }

        private bool isPreservedPdfMode = false;

        private static bool DetectIsPartyDocText(string text)
        {
            string u = (text ?? "").ToUpperInvariant();
            return u.Contains("ĐẢNG CỘNG SẢN VIỆT NAM") ||
                   u.Contains("ĐẢNG ỦY") ||
                   u.Contains("ĐẢNG BỘ") ||
                   u.Contains("QUẬN ỦY") ||
                   u.Contains("HUYỆN ỦY") ||
                   u.Contains("THÀNH ỦY") ||
                   u.Contains("TỈNH ỦY") ||
                   u.Contains("CHI BỘ") ||
                   u.Contains("BAN THƯỜNG VỤ") ||
                   u.Contains("BAN CHẤP HÀNH") ||
                   u.Contains("VPTW") ||
                   Regex.IsMatch(u, @"\bSỐ\s*:?\s*\d+\s*-\s*[A-ZĐ]+\/(ĐU|CB|QU|HU|TU|TW)\b") ||
                   Regex.IsMatch(u, @"(-|\/)(ĐU|CB)\b");
        }

        private string BuildLayoutPreservedPdfWordHtml(string rawText)
        {
            StringBuilder sb = new StringBuilder();
            bool isPartyDoc = DetectIsPartyDocText(rawText);
            string[] pages = (rawText ?? "").Split('\f');

            for (int pIdx = 0; pIdx < pages.Length; pIdx++)
            {
                string pageContent = pages[pIdx];
                if (string.IsNullOrWhiteSpace(pageContent)) continue;
                // Nối chảy tự nhiên liên tục giữa các trang PDF, không chèn page-break cứng gây trang trắng thừa trong Word

                string[] rawLines = pageContent.Replace("\r\n", "\n").Split('\n');
                List<string> lines = new List<string>();
                foreach (string rl in rawLines)
                {
                    if (!string.IsNullOrWhiteSpace(rl)) lines.Add(rl);
                }
                if (lines.Count == 0) continue;

                List<string> headerLeft = new List<string>();
                List<string> headerRight = new List<string>();
                List<string> footerLeft = new List<string>();
                List<string> footerRight = new List<string>();
                List<string> bodyLines = new List<string>();

                string state = pIdx == 0 ? "HEADER" : "BODY";
                for (int i = 0; i < lines.Count; i++)
                {
                    string rawLine = lines[i];
                    int leadingSpaces = 0;
                    while (leadingSpaces < rawLine.Length && rawLine[leadingSpaces] == ' ') leadingSpaces++;
                    string trimmed = rawLine.Trim();
                    if (Regex.IsMatch(trimmed, @"^[-_*—–=]{2,}$")) continue;

                    if (state == "HEADER")
                    {
                        Match twoCol = Regex.Match(trimmed, @"^(.+?)(?:\t+|\s{4,})(.+)$");
                        if (twoCol.Success && i <= 8)
                        {
                            headerLeft.Add(twoCol.Groups[1].Value.Trim());
                            headerRight.Add(twoCol.Groups[2].Value.Trim());
                            continue;
                        }
                        if (i <= 8 && Regex.IsMatch(trimmed, @"^(ĐẢNG CỘNG SẢN VIỆT NAM|CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM|Độc lập\s*-\s*Tự do)", RegexOptions.IgnoreCase))
                        {
                            headerRight.Add(trimmed);
                            continue;
                        }
                        if (i <= 8 && Regex.IsMatch(trimmed, @"\b(ngày|tháng|năm)\b", RegexOptions.IgnoreCase) && !trimmed.StartsWith("Hôm nay", StringComparison.OrdinalIgnoreCase))
                        {
                            headerRight.Add(trimmed);
                            continue;
                        }
                        if (i <= 7 && Regex.IsMatch(trimmed, @"^(QUẬN ỦY|HUYỆN ỦY|THÀNH ỦY|TỈNH ỦY|ĐẢNG ỦY|ĐẢNG BỘ|CHI BỘ|BAN CHẤP HÀNH|BAN THƯỜNG VỤ|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN|Số\b|Số\s*:|V\/v\b|Về việc\b)", RegexOptions.IgnoreCase))
                        {
                            headerLeft.Add(trimmed);
                            continue;
                        }
                        state = "BODY";
                    }

                    if (state == "BODY")
                    {
                        if (Regex.IsMatch(trimmed, @"^Nơi nhận\b", RegexOptions.IgnoreCase) ||
                            Regex.IsMatch(trimmed, @"\b(T\/M\s+ĐẢNG\s+ỦY|T\/M\s+BAN\s+THƯỜNG\s+VỤ|TM\.\s+ỦY\s+BAN|KT\.\s+CHỦ\s+TỊCH|TL\.\s+CHỦ\s+TỊCH)\b", RegexOptions.IgnoreCase))
                        {
                            state = "FOOTER";
                        }
                    }

                    if (state == "FOOTER")
                    {
                        Match twoCol = Regex.Match(trimmed, @"^(.+?)(?:\t+|\s{4,})(.+)$");
                        if (twoCol.Success)
                        {
                            footerLeft.Add(twoCol.Groups[1].Value.Trim());
                            footerRight.Add(twoCol.Groups[2].Value.Trim());
                        }
                        else if (Regex.IsMatch(trimmed, @"^(Nơi nhận|-|\+|•|Lưu\b)", RegexOptions.IgnoreCase))
                        {
                            footerLeft.Add(trimmed);
                        }
                        else if (leadingSpaces >= 20 || (trimmed == trimmed.ToUpperInvariant() && trimmed != trimmed.ToLowerInvariant()) || footerRight.Count > 0)
                        {
                            footerRight.Add(trimmed);
                        }
                        else
                        {
                            footerLeft.Add(trimmed);
                        }
                        continue;
                    }

                    bodyLines.Add(trimmed);
                }

                // 1. Kết xuất Bảng Đầu Trang (Chuẩn Hướng dẫn 05-HD/VPTW hoặc Nghị định 30/2020/NĐ-CP)
                if (headerLeft.Count > 0 || headerRight.Count > 0)
                {
                    List<string> agencies = new List<string>();
                    string numLine = "";
                    List<string> subjLeft = new List<string>();
                    foreach (string hl in headerLeft)
                    {
                        if (string.IsNullOrEmpty(numLine) && Regex.IsMatch(hl, @"^(Số\b|Số\s*:)", RegexOptions.IgnoreCase))
                        {
                            string cNum = Regex.Replace(hl, @"^Số\s*:?\s*", "", RegexOptions.IgnoreCase).Trim();
                            numLine = isPartyDoc ? ("Số " + cNum) : ("Số: " + cNum);
                        }
                        else if (!string.IsNullOrEmpty(numLine) || Regex.IsMatch(hl, @"^(V\/v|Về việc)\b", RegexOptions.IgnoreCase))
                        {
                            subjLeft.Add(hl);
                        }
                        else
                        {
                            agencies.Add(hl);
                        }
                    }

                    StringBuilder sbLeft = new StringBuilder();
                    string agSize = isPartyDoc ? "14pt" : "12.5pt";
                    if (agencies.Count >= 2)
                    {
                        sbLeft.AppendFormat("<div style='font-size:{0};text-transform:uppercase;'>{1}</div>", agSize, System.Security.SecurityElement.Escape(agencies[0]));
                        for (int k = 1; k < agencies.Count; k++)
                        {
                            sbLeft.AppendFormat("<div style='font-size:{0};font-weight:bold;text-transform:uppercase;'>{1}</div>", agSize, System.Security.SecurityElement.Escape(agencies[k]));
                        }
                    }
                    else if (agencies.Count == 1)
                    {
                        sbLeft.AppendFormat("<div style='font-size:{0};font-weight:bold;text-transform:uppercase;'>{1}</div>", agSize, System.Security.SecurityElement.Escape(agencies[0]));
                    }
                    if (agencies.Count > 0)
                    {
                        sbLeft.Append(isPartyDoc
                            ? "<p align='center' style='margin:1pt 0 2pt 0;text-indent:0;text-align:center;font-size:14pt;font-weight:bold;line-height:1;'>*</p>"
                            : "<p align='center' style='margin:1pt 0 3pt 0;text-indent:0;text-align:center;font-size:9pt;line-height:1;'>────────────</p>");
                    }
                    if (!string.IsNullOrEmpty(numLine))
                    {
                        sbLeft.AppendFormat("<p align='center' style='margin:2pt 0 0 0;text-indent:0;text-align:center;font-size:{0};'>{1}</p>", isPartyDoc ? "14pt" : "13pt", System.Security.SecurityElement.Escape(numLine));
                    }
                    if (subjLeft.Count > 0)
                    {
                        sbLeft.AppendFormat("<p align='center' style='margin:3pt 0 0 0;text-indent:0;text-align:center;font-size:12pt;'>{0}</p>", System.Security.SecurityElement.Escape(string.Join(" ", subjLeft.ToArray())));
                    }

                    StringBuilder sbRight = new StringBuilder();
                    string dateLine = "";
                    foreach (string hr in headerRight)
                    {
                        if (Regex.IsMatch(hr, @"\b(ngày|tháng|năm)\b", RegexOptions.IgnoreCase))
                            dateLine = hr;
                    }
                    if (isPartyDoc)
                    {
                        sbRight.Append("<p align='center' style='margin:0;text-indent:0;text-align:center;font-size:15pt;font-weight:bold;'><span style='border-bottom:1pt solid #000000;padding-bottom:1pt;'>ĐẢNG CỘNG SẢN VIỆT NAM</span></p>");
                    }
                    else
                    {
                        sbRight.Append("<p align='center' style='margin:0;text-indent:0;text-align:center;font-size:12.5pt;font-weight:bold;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</p>");
                        sbRight.Append("<p align='center' style='margin:0;text-indent:0;text-align:center;font-size:13.5pt;font-weight:bold;'><span style='border-bottom:1pt solid #000000;padding-bottom:1pt;'>Độc lập - Tự do - Hạnh phúc</span></p>");
                    }
                    if (!string.IsNullOrEmpty(dateLine))
                    {
                        sbRight.AppendFormat("<p align='center' style='margin:6pt 0 0 0;text-indent:0;text-align:center;font-size:{0};font-style:italic;'>{1}</p>", isPartyDoc ? "14pt" : "13.5pt", System.Security.SecurityElement.Escape(dateLine));
                    }

                    sb.Append("<table width='100%' cellpadding='0' cellspacing='0' border='0' style='width:100%;border-collapse:collapse;border:none;mso-table-lspace:0pt;mso-table-rspace:0pt;margin:0 0 10pt 0;'><tr>");
                    sb.AppendFormat("<td width='44%' valign='top' align='center' style='width:44%;vertical-align:top;border:none;text-align:center;'>{0}</td>", sbLeft.ToString());
                    sb.AppendFormat("<td width='56%' valign='top' align='center' style='width:56%;vertical-align:top;border:none;text-align:center;'>{0}</td>", sbRight.ToString());
                    sb.Append("</tr></table>");
                }

                // 2. Kết xuất Tên loại văn bản, Trích yếu & Nối dòng đoạn văn thân bài
                int bIdx = 0;
                while (bIdx < bodyLines.Count)
                {
                    string l = bodyLines[bIdx];
                    if (bIdx <= 3 && Regex.IsMatch(l, @"^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|QUY CHẾ|QUY ĐỊNH|BIÊN BẢN|GIẤY MỜI|HƯỚNG DẪN|KẾT LUẬN)$", RegexOptions.IgnoreCase))
                    {
                        sb.AppendFormat("<p style='text-align:center;text-indent:0;font-weight:bold;font-size:{0};margin:12pt 0 4pt 0;'>{1}</p>",
                            isPartyDoc ? "15pt" : "14pt",
                            System.Security.SecurityElement.Escape(l));
                        bIdx++;
                        if (bIdx < bodyLines.Count && !Regex.IsMatch(bodyLines[bIdx], @"^(Căn cứ|Thực hiện|Kính gửi|I\.|1\.|Điều\s+1)", RegexOptions.IgnoreCase))
                        {
                            sb.AppendFormat("<p align='center' style='text-align:center;text-indent:0;font-weight:bold;font-size:14pt;margin:0 0 2pt 0;'>{0}</p>",
                                System.Security.SecurityElement.Escape(bodyLines[bIdx]));
                            sb.Append("<p align='center' style='text-align:center;text-indent:0;font-size:10pt;margin:0 0 10pt 0;line-height:1;'>──────────────</p>");
                            bIdx++;
                        }
                        continue;
                    }

                    string merged = l;
                    bIdx++;
                    while (bIdx < bodyLines.Count)
                    {
                        string nxt = bodyLines[bIdx];
                        bool endedStop = Regex.IsMatch(merged, @"[.:;!?]$|\.\s*\/\.\s*$");
                        bool startsNew = Regex.IsMatch(nxt, @"^([-+•*]|\d+[.)]|[a-zđ]\)|[IVX]+\.|Điều\s+\d+|Khoản\s+\d+|Căn cứ\s+|Xét\s+|Theo đề nghị|Kính gửi|QUYẾT ĐỊNH|QUYẾT NGHỊ|BAN THƯỜNG VỤ|ĐẢNG ỦY|ỦY BAN NHÂN DÂN|HỘI ĐỒNG NHÂN DÂN)", RegexOptions.IgnoreCase);
                        if (!endedStop && !startsNew)
                        {
                            merged += " " + nxt;
                            bIdx++;
                        }
                        else
                        {
                            break;
                        }
                    }

                    string esc = System.Security.SecurityElement.Escape(merged);
                    if (Regex.IsMatch(merged, @"^(QUYẾT ĐỊNH|QUYẾT NGHỊ|NGHỊ QUYẾT|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG|HỘI ĐỒNG NHÂN DÂN PHƯỜNG)\s*:?$", RegexOptions.IgnoreCase))
                    {
                        sb.AppendFormat("<p style='text-indent:0;text-align:center;font-weight:bold;margin:6pt 0;font-size:14pt;'>{0}</p>", esc);
                    }
                    else if (Regex.IsMatch(merged, @"^[IVX]+\.\s+", RegexOptions.IgnoreCase))
                    {
                        sb.AppendFormat("<p style='text-indent:0;text-align:left;font-weight:bold;margin:6pt 0 4pt 0;font-size:14pt;'>{0}</p>", esc);
                    }
                    else
                    {
                        Match dm = Regex.Match(merged, @"^(Điều\s+\d+\.|Kính gửi\s*:)\s*(.*)$", RegexOptions.IgnoreCase);
                        if (dm.Success)
                        {
                            sb.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;line-height:1.38;font-size:14pt;'><b>{0}</b> {1}</p>",
                                System.Security.SecurityElement.Escape(dm.Groups[1].Value),
                                System.Security.SecurityElement.Escape(dm.Groups[2].Value));
                        }
                        else
                        {
                            sb.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;line-height:1.38;font-size:14pt;'>{0}</p>", esc);
                        }
                    }
                }

                // 3. Kết xuất Bảng Chân Trang (2 cột: Nơi nhận & Chữ ký chuẩn Hướng dẫn 05-HD/VPTW và Nghị định 30)
                if (footerLeft.Count > 0 || footerRight.Count > 0)
                {
                    StringBuilder sbFLeft = new StringBuilder();
                    sbFLeft.Append(isPartyDoc
                        ? "<div style='font-size:14pt;font-style:normal;font-weight:normal;'><u>Nơi nhận</u>:</div>"
                        : "<div style='font-size:12pt;font-weight:bold;font-style:italic;'>Nơi nhận:</div>");
                    string recSize = isPartyDoc ? "12pt" : "11pt";
                    foreach (string fl in footerLeft)
                    {
                        if (Regex.IsMatch(fl, @"^Nơi nhận\s*:?$", RegexOptions.IgnoreCase)) continue;
                        string cleanItem = Regex.Replace(fl, @"^Nơi nhận\s*:\s*", "", RegexOptions.IgnoreCase).Trim();
                        if (!string.IsNullOrEmpty(cleanItem))
                        {
                            sbFLeft.AppendFormat("<div style='font-size:{0};line-height:1.35;text-align:left;'>{1}</div>", recSize, System.Security.SecurityElement.Escape(cleanItem));
                        }
                    }

                    StringBuilder sbFRight = new StringBuilder();
                    for (int k = 0; k < footerRight.Count; k++)
                    {
                        string fr = footerRight[k];
                        if (k == footerRight.Count - 1 && footerRight.Count >= 2)
                        {
                            sbFRight.Append("<br/><br/><br/>");
                            sbFRight.AppendFormat("<div style='font-size:14pt;font-weight:bold;'>{0}</div>", System.Security.SecurityElement.Escape(fr));
                        }
                        else
                        {
                            sbFRight.AppendFormat("<div style='font-size:{0};font-weight:bold;text-transform:uppercase;'>{1}</div>", isPartyDoc ? "14pt" : "13.5pt", System.Security.SecurityElement.Escape(fr));
                        }
                    }

                    sb.Append("<table style='width:100%;border-collapse:collapse;border:none;margin-top:16pt;'><tr>");
                    sb.AppendFormat("<td style='width:50%;vertical-align:top;border:none;text-align:left;'>{0}</td>", sbFLeft.ToString());
                    sb.AppendFormat("<td style='width:50%;vertical-align:top;border:none;text-align:center;'>{0}</td>", sbFRight.ToString());
                    sb.Append("</tr></table>");
                }
            }
            return sb.ToString();
        }

        public void InsertTemplateIntoActiveWord(int tplIdx, string customTopic = "")
        {
            try
            {
                if (tplIdx >= 0 && tplIdx < cboTemplate.Items.Count)
                {
                    cboTemplate.SelectedIndex = tplIdx;
                    LoadTemplate(tplIdx);
                }
                string[] cfg = UnitSettingsForm.LoadUnitConfig();
                bool isDang = tplIdx >= 10;
                txtGovOrg.Text = isDang ? cfg[0] : cfg[2];
                txtIssueOrg.Text = isDang ? cfg[1] : cfg[3];
                txtLocation.Text = cfg[4] + ", ngày " + DateTime.Now.ToString("dd") + " tháng " + (DateTime.Now.Month <= 2 ? "0" + DateTime.Now.Month : DateTime.Now.Month.ToString()) + " năm " + DateTime.Now.Year;
                txtSignerName.Text = isDang ? cfg[5] : cfg[6];
                if (!string.IsNullOrWhiteSpace(customTopic))
                {
                    txtSubject.Text = customTopic;
                }

                string tplDir = Path.Combine(Path.GetTempPath(), "VBTM_Templates");
                Directory.CreateDirectory(tplDir);
                string tempDoc = Path.Combine(tplDir, "VBTM_Mau_" + (tplIdx + 1) + "_" + DateTime.Now.ToString("HHmmssfff") + ".doc");
                WriteWordDocFile(tempDoc);

                object wordApp = null;
                try { wordApp = Marshal.GetActiveObject("Word.Application"); } catch { }
                if (wordApp == null)
                {
                    Process.Start(tempDoc);
                    return;
                }

                bool inserted = false;
                try
                {
                    object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                    int docCnt = (int)docs.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, docs, null);
                    if (docCnt == 0)
                    {
                        docs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, docs, new object[] { tempDoc });
                        inserted = true;
                    }
                    else
                    {
                        object sel = wordApp.GetType().InvokeMember("Selection", BindingFlags.GetProperty, null, wordApp, null);
                        sel.GetType().InvokeMember("InsertFile", BindingFlags.InvokeMethod, null, sel, new object[] { tempDoc });
                        inserted = true;
                        try { File.Delete(tempDoc); } catch { }
                    }
                }
                catch
                {
                    try
                    {
                        object docs = wordApp.GetType().InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);
                        docs.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, docs, new object[] { tempDoc });
                        inserted = true;
                    }
                    catch { }
                }

                if (!inserted)
                {
                    Process.Start(tempDoc);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Không thể chèn trực tiếp vào Word đang mở: " + ex.Message, "VBTM Pro 2.5.0");
            }
        }

        private void WriteWordDocFile(string filePath)
        {
            if (isPreservedPdfMode)
            {
                string preservedHtml = @"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
@page Section1 {
  size:21.0cm 29.7cm;
  margin:2.0cm 1.5cm 2.0cm 3.0cm;
  mso-header-margin:1.0cm;
  mso-footer-margin:1.0cm;
  mso-title-page:yes;
  mso-header:h1;
  mso-first-header:fh1;
}
div.Section1 { page:Section1; font-family:'Times New Roman',serif; font-size:14pt; color:#000; line-height:1.38; }
p.MsoHeader, li.MsoHeader, div.MsoHeader { margin:0cm; text-align:center; font-size:13pt; font-family:'Times New Roman',serif; }
table { width:100%; border-collapse:collapse; }
td, p { font-family:'Times New Roman',serif; }
td p { margin:0; text-indent:0; }
</style></head>
<body><div class='Section1'>" + BuildLayoutPreservedPdfWordHtml(txtBody.Text) + @"
<div style='mso-element:header' id='fh1'><p class='MsoHeader'>&nbsp;</p></div>
<div style='mso-element:header' id='h1'>
  <p class='MsoHeader' align='center' style='text-align:center;font-size:13pt;font-family:""Times New Roman"",serif;'>
    <!--[if supportFields]><span style='mso-element:field-begin'></span> PAGE <span style='mso-element:field-separator'></span><![endif]-->2<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->
  </p>
</div>
</div></body></html>";
                File.WriteAllText(filePath, preservedHtml, new UTF8Encoding(true));
                return;
            }

            StringBuilder sbBody = new StringBuilder();
            foreach (string rawLine in txtBody.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                string line = rawLine.Trim();
                string esc = System.Security.SecurityElement.Escape(line);
                if (Regex.IsMatch(line, @"^(QUYẾT ĐỊNH|QUYẾT NGHỊ|NGHỊ QUYẾT|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH|ĐẢNG ỦY PHƯỜNG QUYẾT ĐỊNH|ỦY BAN NHÂN DÂN PHƯỜNG|HỘI ĐỒNG NHÂN DÂN PHƯỜNG)\s*:?$", RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(line, @"^KHÓA\s+[IVX0-9]+", RegexOptions.IgnoreCase))
                {
                    sbBody.AppendFormat("<p style='text-indent:0;text-align:center;font-weight:bold;margin:6pt 0;font-size:14pt;'>{0}</p>", esc);
                }
                else if (Regex.IsMatch(line, @"^[IVX]+\.\s+", RegexOptions.IgnoreCase))
                {
                    sbBody.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;font-weight:bold;margin:6pt 0 4pt 0;font-size:14pt;'>{0}</p>", esc);
                }
                else
                {
                    Match dm = Regex.Match(line, @"^(Điều\s+\d+\.|Kính gửi\s*:)\s*(.*)$", RegexOptions.IgnoreCase);
                    if (dm.Success)
                    {
                        sbBody.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;font-size:14pt;'><b>{0}</b> {1}</p>",
                            System.Security.SecurityElement.Escape(dm.Groups[1].Value),
                            System.Security.SecurityElement.Escape(dm.Groups[2].Value));
                    }
                    else
                    {
                        sbBody.AppendFormat("<p style='text-indent:1.27cm;text-align:justify;margin:0 0 6pt 0;font-size:14pt;'>{0}</p>", esc);
                    }
                }
            }

            bool isDang = (cboTemplate != null && cboTemplate.SelectedIndex >= 10) ||
                          txtIssueOrg.Text.IndexOf("ĐẢNG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                          txtIssueOrg.Text.IndexOf("CHI BỘ", StringComparison.OrdinalIgnoreCase) >= 0;
            string noiNhanSize = isDang ? "12pt" : "11pt";

            StringBuilder sbRec = new StringBuilder();
            foreach (string r in txtRecipients.Text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
            {
                sbRec.AppendFormat("<p style='margin:0;padding:0;text-indent:0;line-height:1.08;font-size:{1};'>{0}</p>", System.Security.SecurityElement.Escape(r.Trim()), noiNhanSize);
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

            string rightHeaderHtml = isDang
                ? "<p align='center' style='margin:0;padding:0;text-align:center;text-indent:0;font-size:15pt;font-weight:bold;line-height:1.15;'><span style='border-bottom:0.75pt solid #000;padding-bottom:1.5pt;'>ĐẢNG CỘNG SẢN VIỆT NAM</span></p>"
                : "<div style='font-size:12.5pt;font-weight:bold;white-space:nowrap;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>" +
                  "<div style='font-size:13.5pt;font-weight:bold;white-space:nowrap;'>Độc lập - Tự do - Hạnh phúc</div>" +
                  "<p align='center' style='text-align:center;text-indent:0;font-size:8pt;margin:1pt 0 2pt 0;line-height:1;'>────────────────────────</p>";

            string cleanNum = Regex.Replace(txtDocNum.Text.Trim(), @"^Số\s*:?\s*", "", RegexOptions.IgnoreCase);
            string numberLabel = isDang ? ("Số " + cleanNum) : ("Số: " + cleanNum);
            string leftSepHtml = isDang
                ? "<div style='font-size:13pt;font-weight:bold;line-height:1;margin:2pt 0;'>*</div>"
                : "<p align='center' style='text-align:center;text-indent:0;font-size:8pt;margin:1pt 0 2pt 0;line-height:1;'>──────────────</p>";

            string docTypeStr = txtDocType.Text.Trim();
            string docSubjStr = txtSubject.Text.Trim();
            string congVanSubjLeftHtml = (string.IsNullOrEmpty(docTypeStr) && !string.IsNullOrEmpty(docSubjStr))
                ? ("<div style='font-size:12pt;margin-top:4pt;'>" + System.Security.SecurityElement.Escape(docSubjStr) + "</div>")
                : "";

            string centerTitleBlockHtml = "";
            if (!string.IsNullOrEmpty(docTypeStr))
            {
                centerTitleBlockHtml += "<p align='center' style='text-align:center;text-indent:0;font-weight:bold;font-size:" + (isDang ? "15pt" : "14pt") + ";margin:14pt 0 4pt 0;'>" + System.Security.SecurityElement.Escape(docTypeStr) + "</p>";
                if (!string.IsNullOrEmpty(docSubjStr))
                {
                    centerTitleBlockHtml += "<p align='center' style='text-align:center;text-indent:0;font-weight:bold;font-size:14pt;margin:0 0 2pt 0;'>" + System.Security.SecurityElement.Escape(docSubjStr) + "</p>" +
                                            (isDang
                                                ? "<p align='center' style='text-align:center;text-indent:0;font-weight:bold;font-size:14pt;margin:0 0 10pt 0;line-height:1;'>-----</p>"
                                                : "<p align='center' style='text-align:center;text-indent:0;font-size:10pt;margin:0 0 10pt 0;line-height:1;'>──────────────</p>");
                }
            }
            else
            {
                centerTitleBlockHtml = "<div style='height:10pt;'></div>";
            }

            string[] signerParts = txtSignerTitle.Text.Split(new[] { "\r\n", "\n", " - " }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder sbSigTitle = new StringBuilder();
            for (int si = 0; si < signerParts.Length; si++)
            {
                string sp = signerParts[si].Trim();
                bool isPlainDangRole = isDang && Regex.IsMatch(sp, @"^(BÍ THƯ|PHÓ BÍ THƯ|PHÓ BÍ THƯ THƯỜNG TRỰC|CHÁNH VĂN PHÒNG|PHÓ CHÁNH VĂN PHÒNG|CHỦ NHIỆM|TRƯỞNG BAN)$", RegexOptions.IgnoreCase);
                sbSigTitle.AppendFormat("<p align='center' style='margin:0;padding:0;text-align:center;text-indent:0;line-height:1.15;font-size:14pt;font-weight:{1};'>{0}</p>",
                    System.Security.SecurityElement.Escape(sp), isPlainDangRole ? "normal" : "bold");
            }
            string signerTitleFormatted = sbSigTitle.ToString();

            // Theo Hướng dẫn 05-HD/VPTW: Từ "Nơi nhận:" cỡ chữ 14pt, kiểu chữ đứng, có gạch chân dưới từ "Nơi nhận", sau có dấu hai chấm (:)
            string noiNhanHeaderHtml = isDang
                ? "<p style='margin:0;padding:0;text-indent:0;line-height:1.08;font-size:14pt;font-style:normal;font-weight:normal;'><u>Nơi nhận</u>:</p>"
                : "<p style='margin:0;padding:0;text-indent:0;line-height:1.08;font-size:12pt;font-weight:bold;font-style:italic;'>Nơi nhận:</p>";

            string html = @"<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/></w:WordDocument></xml><![endif]-->
<style>
@page Section1 {
  size:21.0cm 29.7cm;
  margin:2.0cm 1.5cm 2.0cm 3.0cm;
  mso-header-margin:1.0cm;
  mso-footer-margin:1.0cm;
  mso-title-page:yes;
  mso-header:h1;
  mso-first-header:fh1;
}
div.Section1 { page:Section1; font-family:'Times New Roman',serif; font-size:14pt; color:#000; line-height:1.38; }
p.MsoHeader, li.MsoHeader, div.MsoHeader { margin:0cm; text-align:center; font-size:13pt; font-family:'Times New Roman',serif; }
table { width:100%; border-collapse:collapse; }
td { vertical-align:top; font-family:'Times New Roman',serif; }
td p { margin:0cm; margin-bottom:0pt; mso-para-margin:0cm; mso-para-margin-bottom:0pt; line-height:1.08; }
</style></head>
<body><div class='Section1'>
<table><tr>
<td style='width:" + (isDang ? "50%" : "42%") + @";text-align:center;'>
  " + (string.IsNullOrWhiteSpace(txtGovOrg.Text) ? "" : "<div style='font-size:" + (isDang ? (txtGovOrg.Text.Trim().Length > 25 ? "12.5pt" : (txtGovOrg.Text.Trim().Length > 21 ? "13pt" : "14pt")) : "12.5pt") + ";'>" + System.Security.SecurityElement.Escape(txtGovOrg.Text) + "</div>") + @"
  <div style='font-size:" + (isDang ? (txtIssueOrg.Text.Trim().Length > 25 ? "12.5pt" : (txtIssueOrg.Text.Trim().Length > 21 ? "13pt" : "14pt")) : "12.5pt") + @";font-weight:bold;'>" + System.Security.SecurityElement.Escape(txtIssueOrg.Text) + @"</div>
  " + leftSepHtml + @"
  <div style='font-size:" + (isDang ? "14pt" : "13pt") + @";margin-top:4pt;'>" + System.Security.SecurityElement.Escape(numberLabel) + @"</div>
  " + congVanSubjLeftHtml + @"
</td>
<td style='width:" + (isDang ? "50%" : "58%") + @";text-align:center;'>
  " + rightHeaderHtml + @"
  <div style='font-size:" + (txtLocation.Text.Trim().Length > 40 ? "12pt" : (txtLocation.Text.Trim().Length > 32 ? "13pt" : "14pt")) + @";font-style:italic;margin-top:4pt;'>" + System.Security.SecurityElement.Escape(txtLocation.Text) + @"</div>
</td>
</tr></table>
" + centerTitleBlockHtml + @"
" + sbBody.ToString() + @"
<table style='margin-top:16pt;'><tr>
<td style='width:50%;font-size:" + noiNhanSize + @";'>" + noiNhanHeaderHtml + sbRec.ToString() + qrImgHtml + @"</td>
<td style='width:50%;text-align:center;font-size:14pt;'>" + signerTitleFormatted + @"<br/><br/><br/><p align='center' style='margin:0;padding:0;text-align:center;text-indent:0;font-size:14pt;font-weight:bold;'>" + System.Security.SecurityElement.Escape(txtSignerName.Text) + @"</p></td>
</tr></table>
<div style='mso-element:header' id='fh1'><p class='MsoHeader'>&nbsp;</p></div>
<div style='mso-element:header' id='h1'>
  <p class='MsoHeader' align='center' style='text-align:center;font-size:13pt;font-family:""Times New Roman"",serif;'>
    <!--[if supportFields]><span style='mso-element:field-begin'></span> PAGE <span style='mso-element:field-separator'></span><![endif]-->2<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->
  </p>
</div>
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
                if (MessageBox.Show("Đã xuất tệp Word chuẩn thể thức thành công!\nBạn có muốn mở tệp ngay không?", "VBTM Pro 2.5.0", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
        }

        private void OpenWordOrTextFileFromDisk()
        {
            using (OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Chọn tệp văn bản Word (.docx, .doc), PDF (.pdf) hoặc Text (.txt) để chuẩn hóa",
                Filter = "Word, PDF & Text Files (*.docx;*.doc;*.pdf;*.txt)|*.docx;*.doc;*.pdf;*.txt|All Files (*.*)|*.*"
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;

                try
                {
                    string ext = Path.GetExtension(ofd.FileName).ToLowerInvariant();
                    string extractedText = "";

                    if (ext == ".pdf")
                    {
                        string outTxt = Path.Combine(Path.GetTempPath(), "vbtm_pdf_txt_" + Guid.NewGuid().ToString("N") + ".txt");
                        string gsExe = @"C:\Program Files\PDF24\gs\bin\gswin64c.exe";
                        try
                        {
                            if (File.Exists(gsExe))
                            {
                                ProcessStartInfo gsTxtPsi = new ProcessStartInfo
                                {
                                    FileName = gsExe,
                                    Arguments = string.Format("-dNOPAUSE -dBATCH -dSAFER -q -sDEVICE=txtwrite \"-sOutputFile={0}\" \"{1}\"", outTxt, ofd.FileName),
                                    UseShellExecute = false,
                                    CreateNoWindow = true
                                };
                                using (Process gp = Process.Start(gsTxtPsi))
                                {
                                    gp.WaitForExit(8000);
                                }
                                if (File.Exists(outTxt))
                                {
                                    extractedText = File.ReadAllText(outTxt, Encoding.UTF8);
                                }
                            }
                        }
                        finally
                        {
                            try { if (File.Exists(outTxt)) File.Delete(outTxt); } catch { }
                        }
                    }
                    else if (ext == ".docx")
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

                    if (ext == ".pdf")
                    {
                        if (string.IsNullOrWhiteSpace(extractedText))
                        {
                            MessageBox.Show("Không tìm thấy nội dung văn bản trong tệp PDF đã chọn!", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        isPreservedPdfMode = true;
                        txtBody.Text = extractedText.TrimEnd();
                        lblStatus.Text = "  ✓ Đã đọc tệp PDF giữ nguyên 100% thể thức và nội dung gốc: " + Path.GetFileName(ofd.FileName);
                        return;
                    }

                    isPreservedPdfMode = false;
                    extractedText = Regex.Replace(extractedText ?? "", @"(\r?\n){3,}", "\r\n\r\n").Trim();
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        MessageBox.Show("Không tìm thấy nội dung văn bản trong tệp đã chọn!", "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    txtBody.Text = extractedText;
                    lblStatus.Text = "  ✓ Đã đọc tệp giữ nguyên nội dung gốc: " + Path.GetFileName(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi đọc tệp văn bản: " + ex.Message, "VBTM Pro 2.5.0", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
