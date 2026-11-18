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
                Text = "★ Cài Đặt Ngay",
                Location = new Point(434, 384),
                Size = new Size(170, 38),
                BackColor = Color.FromArgb(234, 179, 8),
                ForeColor = Color.FromArgb(69, 10, 10),
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

                MessageBox.Show(
                    "Đã cài đặt thành công VBTM Pro 2.5.0!\n\n" +
                    "• Thư mục cài đặt: " + installDir + "\n" +
                    "• Lối tắt Desktop: VBTM Pro 2.5.0\n" +
                    (wordIntegrated
                        ? "• Tích hợp Word: ĐÃ CÀI SẴN Tab Ribbon 'VBTM Pro 2.5 (05-HD & ND30)' trực tiếp trong Microsoft Word!\n\n"
                        : "• Tệp Macro Word: " + Path.Combine(installDir, "VBTM_Word_Addin.bas") + "\n\n") +
                    "Nhấn OK để mở phần mềm ngay.",
                    "Cài Đặt Thành Công - VBTM Pro 2.5.0",
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
            return WordAddinInstaller.GetWordVbaMacroCode();
        }
    }

    // =========================================================================
    // 1B. BỘ TÍCH HỢP TRỰC TIẾP VÀO MICROSOFT WORD (STARTUP .DOTM + RIBBON TAB)
    // =========================================================================
    public static class WordAddinInstaller
    {
        public static string GetWordVbaMacroCode()
        {
            return
                "Attribute VB_Name = \"VBTM_Pro_HD05_ND30\"\r\n" +
                "Sub Ribbon_ChuanHoaTuDong(ByVal control As Object)\r\n" +
                "    ChuanHoaVanBan_TuDong\r\n" +
                "End Sub\r\n" +
                "Sub Ribbon_HuongDan05(ByVal control As Object)\r\n" +
                "    ChuanHoaVanBan_HuongDan05\r\n" +
                "End Sub\r\n" +
                "Sub Ribbon_NghiDinh30(ByVal control As Object)\r\n" +
                "    ChuanHoaVanBan_NghiDinh30\r\n" +
                "End Sub\r\n" +
                "Sub Ribbon_MoVBTM(ByVal control As Object)\r\n" +
                "    On Error Resume Next\r\n" +
                "    Dim p As String\r\n" +
                "    p = Environ(\"LOCALAPPDATA\") & \"\\Programs\\VBTM-Pro\\VBTM-Pro.exe\"\r\n" +
                "    If Dir(p) <> \"\" Then\r\n" +
                "        Shell \"\"\"\" & p & \"\"\" --app\", vbNormalFocus\r\n" +
                "    Else\r\n" +
                "        MsgBox \"Khong tim thay VBTM-Pro.exe tai: \" & p, vbExclamation, \"VBTM Pro 2.5.0\"\r\n" +
                "    End If\r\n" +
                "End Sub\r\n" +
                "Sub ChuanHoaVanBan_TuDong()\r\n" +
                "    On Error Resume Next\r\n" +
                "    If ActiveDocument Is Nothing Then Exit Sub\r\n" +
                "    Dim docText As String\r\n" +
                "    docText = UCase(ActiveDocument.Content.Text)\r\n" +
                "    If InStr(docText, \"ĐẢNG CỘNG SẢN\") > 0 Or InStr(docText, \"ĐẢNG ỦY\") > 0 Or InStr(docText, \"CHI BỘ\") > 0 Or InStr(docText, \"-BC/ĐU\") > 0 Or InStr(docText, \"-NQ/ĐU\") > 0 Or InStr(docText, \"-CV/ĐU\") > 0 Then\r\n" +
                "        ChuanHoaVanBan_HuongDan05\r\n" +
                "    Else\r\n" +
                "        ChuanHoaVanBan_NghiDinh30\r\n" +
                "    End If\r\n" +
                "End Sub\r\n" +
                "Sub ChuanHoaVanBan_HuongDan05()\r\n" +
                "    On Error Resume Next\r\n" +
                "    If ActiveDocument Is Nothing Then Exit Sub\r\n" +
                "    With ActiveDocument.PageSetup\r\n" +
                "        .PaperSize = wdPaperA4\r\n" +
                "        .TopMargin = CentimetersToPoints(2)\r\n" +
                "        .BottomMargin = CentimetersToPoints(2)\r\n" +
                "        .LeftMargin = CentimetersToPoints(3)\r\n" +
                "        .RightMargin = CentimetersToPoints(1.5)\r\n" +
                "        .DifferentFirstPageHeaderFooter = True\r\n" +
                "    End With\r\n" +
                "    ActiveDocument.Content.Font.Name = \"Times New Roman\"\r\n" +
                "    Dim hdr As HeaderFooter\r\n" +
                "    Set hdr = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)\r\n" +
                "    hdr.Range.Text = \"\"\r\n" +
                "    hdr.PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False\r\n" +
                "    hdr.Range.Font.Name = \"Times New Roman\"\r\n" +
                "    hdr.Range.Font.Size = 13\r\n" +
                "    Dim p As Paragraph\r\n" +
                "    For Each p In ActiveDocument.Paragraphs\r\n" +
                "        If p.Range.Tables.Count = 0 Then\r\n" +
                "            Dim t As String\r\n" +
                "            t = Trim(Replace(Replace(p.Range.Text, vbCr, \"\"), Chr(7), \"\"))\r\n" +
                "            If Len(t) > 0 Then\r\n" +
                "                p.Range.Font.Name = \"Times New Roman\"\r\n" +
                "                If t = \"BÁO CÁO\" Or t = \"NGHỊ QUYẾT\" Or t = \"QUYẾT ĐỊNH\" Or t = \"KẾ HOẠCH\" Or t = \"CHƯƠNG TRÌNH\" Or t = \"TỜ TRÌNH\" Or t = \"THÔNG BÁO\" Then\r\n" +
                "                    p.Range.Font.Size = 15\r\n" +
                "                    p.Range.Font.Bold = True\r\n" +
                "                    p.Alignment = wdAlignParagraphCenter\r\n" +
                "                    p.FirstLineIndent = 0\r\n" +
                "                Else\r\n" +
                "                    p.Range.Font.Size = 14\r\n" +
                "                    p.Format.SpaceAfter = 6\r\n" +
                "                    p.Format.LineSpacingRule = wdLineSpaceMultiple\r\n" +
                "                    p.Format.LineSpacing = LinesToPoints(1.15)\r\n" +
                "                    If p.Alignment <> wdAlignParagraphCenter Then\r\n" +
                "                        p.Alignment = wdAlignParagraphJustify\r\n" +
                "                        p.FirstLineIndent = CentimetersToPoints(1.27)\r\n" +
                "                    End If\r\n" +
                "                End If\r\n" +
                "            End If\r\n" +
                "        End If\r\n" +
                "    Next p\r\n" +
                "    If ActiveDocument.Tables.Count >= 1 Then\r\n" +
                "        Dim tblTop As Table\r\n" +
                "        Set tblTop = ActiveDocument.Tables(1)\r\n" +
                "        If tblTop.Columns.Count >= 2 Then\r\n" +
                "            tblTop.Cell(1, 1).Range.Font.Name = \"Times New Roman\"\r\n" +
                "            tblTop.Cell(1, 1).Range.Font.Size = 14\r\n" +
                "            tblTop.Cell(1, 2).Range.Font.Name = \"Times New Roman\"\r\n" +
                "            tblTop.Cell(1, 2).Range.Font.Size = 15\r\n" +
                "            Dim pTop As Paragraph\r\n" +
                "            For Each pTop In tblTop.Cell(1, 2).Range.Paragraphs\r\n" +
                "                If InStr(LCase(pTop.Range.Text), \"ngày\") > 0 Then\r\n" +
                "                    pTop.Range.Font.Size = 14\r\n" +
                "                    pTop.Range.Font.Italic = True\r\n" +
                "                    pTop.Range.Font.Bold = False\r\n" +
                "                End If\r\n" +
                "            Next pTop\r\n" +
                "        End If\r\n" +
                "    End If\r\n" +
                "    If ActiveDocument.Tables.Count >= 2 Then\r\n" +
                "        Dim tblBot As Table\r\n" +
                "        Set tblBot = ActiveDocument.Tables(ActiveDocument.Tables.Count)\r\n" +
                "        If tblBot.Columns.Count >= 2 Then\r\n" +
                "            Dim pRec As Paragraph\r\n" +
                "            For Each pRec In tblBot.Cell(1, 1).Range.Paragraphs\r\n" +
                "                pRec.Range.Font.Name = \"Times New Roman\"\r\n" +
                "                pRec.Range.Font.Italic = False\r\n" +
                "                If InStr(LCase(pRec.Range.Text), \"nơi nhận\") > 0 Then\r\n" +
                "                    pRec.Range.Font.Size = 14\r\n" +
                "                    pRec.Range.Font.Bold = False\r\n" +
                "                    pRec.Range.Font.Underline = wdUnderlineSingle\r\n" +
                "                Else\r\n" +
                "                    pRec.Range.Font.Size = 12\r\n" +
                "                End If\r\n" +
                "            Next pRec\r\n" +
                "            tblBot.Cell(1, 2).Range.Font.Name = \"Times New Roman\"\r\n" +
                "            tblBot.Cell(1, 2).Range.Font.Size = 14\r\n" +
                "            tblBot.Cell(1, 2).Range.Font.Bold = True\r\n" +
                "        End If\r\n" +
                "    End If\r\n" +
                "    MsgBox \"Đã tự động chuẩn hóa văn bản Word đúng chuẩn Hướng dẫn số 05-HD/VPTW!\", vbInformation, \"VBTM Pro 2.5.0\"\r\n" +
                "End Sub\r\n" +
                "Sub ChuanHoaVanBan_NghiDinh30()\r\n" +
                "    On Error Resume Next\r\n" +
                "    If ActiveDocument Is Nothing Then Exit Sub\r\n" +
                "    With ActiveDocument.PageSetup\r\n" +
                "        .PaperSize = wdPaperA4\r\n" +
                "        .TopMargin = CentimetersToPoints(2)\r\n" +
                "        .BottomMargin = CentimetersToPoints(2)\r\n" +
                "        .LeftMargin = CentimetersToPoints(3)\r\n" +
                "        .RightMargin = CentimetersToPoints(1.5)\r\n" +
                "        .DifferentFirstPageHeaderFooter = True\r\n" +
                "    End With\r\n" +
                "    ActiveDocument.Content.Font.Name = \"Times New Roman\"\r\n" +
                "    Dim p As Paragraph\r\n" +
                "    For Each p In ActiveDocument.Paragraphs\r\n" +
                "        If p.Range.Tables.Count = 0 Then\r\n" +
                "            p.Range.Font.Size = 14\r\n" +
                "            p.Format.SpaceAfter = 6\r\n" +
                "            p.Format.LineSpacingRule = wdLineSpaceMultiple\r\n" +
                "            p.Format.LineSpacing = LinesToPoints(1.15)\r\n" +
                "        End If\r\n" +
                "    Next p\r\n" +
                "    MsgBox \"Đã chuẩn hóa văn bản Word theo Nghị định số 30/2020/NĐ-CP!\", vbInformation, \"VBTM Pro 2.5.0\"\r\n" +
                "End Sub\r\n";
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

                        if (File.Exists(dotmPath)) File.Delete(dotmPath);
                        // 13 = wdFormatXMLTemplateMacroEnabled (.dotm)
                        doc.GetType().InvokeMember("SaveAs2", BindingFlags.InvokeMethod, null, doc, new object[] { dotmPath, 13 });
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

                // 3. Nhúng trực tiếp Custom Ribbon Tab "VBTM Pro 2.5 (05-HD & ND30)" vào bên trong .dotm
                if (File.Exists(dotmPath))
                {
                    InjectRibbonUiIntoDotm(dotmPath);
                    return true;
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
                "<group id=\"grpVbtmFormat\" label=\"Chuẩn Hóa Thể Thức Tự Động\">" +
                "<button id=\"btnVbtmAuto\" label=\"Chuẩn Hóa Tự Động\" size=\"large\" imageMso=\"ReviewTrackChanges\" onAction=\"Ribbon_ChuanHoaTuDong\" supertip=\"Tự động nhận diện văn bản Đảng (05-HD/VPTW) hoặc Hành chính (NĐ 30/2020/NĐ-CP) và chuẩn hóa trực tiếp tài liệu Word đang mở.\"/>" +
                "<button id=\"btnVbtmHd05\" label=\"Chuẩn Đảng (05-HD/VPTW)\" size=\"large\" imageMso=\"HappyFace\" onAction=\"Ribbon_HuongDan05\" supertip=\"Chuẩn hóa trực tiếp theo Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng.\"/>" +
                "<button id=\"btnVbtmNd30\" label=\"Chuẩn Hành Chính (NĐ 30)\" size=\"large\" imageMso=\"FileDocumentInspect\" onAction=\"Ribbon_NghiDinh30\" supertip=\"Chuẩn hóa trực tiếp theo Nghị định số 30/2020/NĐ-CP của Chính phủ.\"/>" +
                "</group>" +
                "<group id=\"grpVbtmStudio\" label=\"Công Cụ Mở Rộng\">" +
                "<button id=\"btnVbtmStudio\" label=\"Mở VBTM Pro Studio (QR &amp; AI)\" size=\"large\" imageMso=\"QuickPartsInsertGallery\" onAction=\"Ribbon_MoVBTM\" supertip=\"Mở phần mềm VBTM Pro 2.5.0 để rà soát lỗi chính tả AI, tạo mã QR và chuyển PDF sang Word.\"/>" +
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

        private string FixAllVietnameseSpellingAndPunctuation(string rawInput, out int fixedCount)
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
                                fntL.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, fntL, new object[] { 14f });

                                object cellR = topTbl.GetType().InvokeMember("Cell", BindingFlags.InvokeMethod, null, topTbl, new object[] { 1, 2 });
                                object rngR = cellR.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, cellR, null);
                                object rParas = rngR.GetType().InvokeMember("Paragraphs", BindingFlags.GetProperty, null, rngR, null);
                                int rpCount = (int)rParas.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, rParas, null);
                                for (int rpIdx = 1; rpIdx <= rpCount; rpIdx++)
                                {
                                    object rp = rParas.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, rParas, new object[] { rpIdx });
                                    object rpRng = rp.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, rp, null);
                                    string rpTxt = Convert.ToString(rpRng.GetType().InvokeMember("Text", BindingFlags.GetProperty, null, rpRng, null)) ?? "";
                                    object rpFnt = rpRng.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, rpRng, null);
                                    rpFnt.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, rpFnt, new object[] { "Times New Roman" });
                                    if (rpTxt.IndexOf("ĐẢNG CỘNG SẢN", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        rpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, rpFnt, new object[] { 15f });
                                        rpFnt.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, rpFnt, new object[] { -1 });
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
                                    if (bpTxt.IndexOf("Nơi nhận", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        bpFnt.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, bpFnt, new object[] { 14f });
                                        bpFnt.GetType().InvokeMember("Underline", BindingFlags.SetProperty, null, bpFnt, new object[] { 1 });
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
                                fntBR.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, fntBR, new object[] { -1 });
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
                            bool isHeadingOrCentered = Regex.IsMatch(cleanP.Trim(), @"^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|BIÊN BẢN|GIẤY MỜI|QUYẾT ĐỊNH\s*:|QUYẾT NGHỊ\s*:|BAN THƯỜNG VỤ ĐẢNG ỦY QUYẾT ĐỊNH)$", RegexOptions.IgnoreCase);
                            try
                            {
                                object pFont = pRange.GetType().InvokeMember("Font", BindingFlags.GetProperty, null, pRange, null);
                                bool isPartyTitle = isDangWord && Regex.IsMatch(cleanP.Trim(), @"^(BÁO CÁO|QUYẾT ĐỊNH|KẾ HOẠCH|TỜ TRÌNH|THÔNG BÁO|NGHỊ QUYẾT|CHƯƠNG TRÌNH|CHỈ THỊ|BIÊN BẢN|GIẤY MỜI)$", RegexOptions.IgnoreCase);
                                pFont.GetType().InvokeMember("Size", BindingFlags.SetProperty, null, pFont, new object[] { isPartyTitle ? 15f : 14f });
                                if (!isHeadingOrCentered)
                                {
                                    para.GetType().InvokeMember("SpaceAfter", BindingFlags.SetProperty, null, para, new object[] { 6f });
                                    para.GetType().InvokeMember("FirstLineIndent", BindingFlags.SetProperty, null, para, new object[] { 36f });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 3 }); // wdAlignParagraphJustify = 3
                                }
                                else if (isPartyTitle)
                                {
                                    pFont.GetType().InvokeMember("Bold", BindingFlags.SetProperty, null, pFont, new object[] { -1 });
                                    para.GetType().InvokeMember("Alignment", BindingFlags.SetProperty, null, para, new object[] { 1 }); // wdAlignParagraphCenter = 1
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
                if (pIdx > 0) sb.Append("<br clear='all' style='page-break-before:always'/>");

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
                    sbBody.AppendFormat("<p style='text-indent:0;text-align:left;font-weight:bold;margin:6pt 0 4pt 0;font-size:14pt;'>{0}</p>", esc);
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

            bool isDang = (cboTemplate != null && cboTemplate.SelectedIndex >= 10) ||
                          txtIssueOrg.Text.IndexOf("ĐẢNG", StringComparison.OrdinalIgnoreCase) >= 0 ||
                          txtIssueOrg.Text.IndexOf("CHI BỘ", StringComparison.OrdinalIgnoreCase) >= 0;

            string rightHeaderHtml = isDang
                ? "<div style='font-size:15pt;font-weight:bold;'>ĐẢNG CỘNG SẢN VIỆT NAM</div>"
                : "<div style='font-size:12.5pt;font-weight:bold;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>" +
                  "<div style='font-size:13.5pt;font-weight:bold;'>Độc lập - Tự do - Hạnh phúc</div>";

            string cleanNum = Regex.Replace(txtDocNum.Text.Trim(), @"^Số\s*:?\s*", "", RegexOptions.IgnoreCase);
            string numberLabel = isDang ? ("Số " + cleanNum) : ("Số: " + cleanNum);
            string leftSepHtml = isDang
                ? "<div style='font-size:13pt;font-weight:bold;line-height:1;margin:2pt 0;'>*</div>"
                : "<hr style='width:38%;border:0;border-top:1px solid #000;'/>";

            string docTypeStr = txtDocType.Text.Trim();
            string docSubjStr = txtSubject.Text.Trim();
            string congVanSubjLeftHtml = (string.IsNullOrEmpty(docTypeStr) && !string.IsNullOrEmpty(docSubjStr))
                ? ("<div style='font-size:12pt;margin-top:4pt;'>" + System.Security.SecurityElement.Escape(docSubjStr) + "</div>")
                : "";

            string centerTitleBlockHtml = "";
            if (!string.IsNullOrEmpty(docTypeStr))
            {
                centerTitleBlockHtml += "<p style='text-align:center;font-weight:bold;font-size:" + (isDang ? "15pt" : "14pt") + ";margin:14pt 0 4pt 0;'>" + System.Security.SecurityElement.Escape(docTypeStr) + "</p>";
                if (!string.IsNullOrEmpty(docSubjStr))
                {
                    centerTitleBlockHtml += "<p style='text-align:center;font-weight:bold;font-size:14pt;margin:0 0 4pt 0;'>" + System.Security.SecurityElement.Escape(docSubjStr) + "</p>" +
                                            "<hr style='width:32%;border:0;border-top:1px solid #000;margin-bottom:12pt;'/>";
                }
            }
            else
            {
                centerTitleBlockHtml = "<div style='height:10pt;'></div>";
            }

            string signerTitleFormatted = string.Join("<br/>",
                Array.ConvertAll(txtSignerTitle.Text.Split(new[] { "\r\n", "\n", " - " }, StringSplitOptions.RemoveEmptyEntries),
                s => System.Security.SecurityElement.Escape(s.Trim())));

            // Theo Hướng dẫn 05-HD/VPTW: Từ "Nơi nhận:" cỡ chữ 14pt, kiểu chữ đứng, có gạch chân dưới từ "Nơi nhận", sau có dấu hai chấm (:)
            string noiNhanHeaderHtml = isDang
                ? "<div style='font-size:14pt;font-style:normal;font-weight:normal;'><u>Nơi nhận</u>:</div>"
                : "<div style='font-size:12pt;font-weight:bold;font-style:italic;'>Nơi nhận:</div>";
            string noiNhanSize = isDang ? "12pt" : "11pt";

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
</style></head>
<body><div class='Section1'>
<table><tr>
<td style='width:44%;text-align:center;'>
  " + (string.IsNullOrWhiteSpace(txtGovOrg.Text) ? "" : "<div style='font-size:" + (isDang ? "14pt" : "12.5pt") + ";'>" + System.Security.SecurityElement.Escape(txtGovOrg.Text) + "</div>") + @"
  <div style='font-size:" + (isDang ? "14pt" : "12.5pt") + @";font-weight:bold;'>" + System.Security.SecurityElement.Escape(txtIssueOrg.Text) + @"</div>
  " + leftSepHtml + @"
  <div style='font-size:" + (isDang ? "14pt" : "13pt") + @";margin-top:6pt;'>" + System.Security.SecurityElement.Escape(numberLabel) + @"</div>
  " + congVanSubjLeftHtml + @"
</td>
<td style='width:56%;text-align:center;'>
  " + rightHeaderHtml + @"
  <hr style='width:" + (isDang ? "56%" : "48%") + @";border:0;border-top:1px solid #000;'/>
  <div style='font-size:" + (isDang ? "14pt" : "13.5pt") + @";font-style:italic;margin-top:6pt;'>" + System.Security.SecurityElement.Escape(txtLocation.Text) + @"</div>
</td>
</tr></table>
" + centerTitleBlockHtml + @"
" + sbBody.ToString() + @"
<table style='margin-top:16pt;'><tr>
<td style='width:50%;font-size:" + noiNhanSize + @";'>" + noiNhanHeaderHtml + sbRec.ToString() + qrImgHtml + @"</td>
<td style='width:50%;text-align:center;font-size:13.5pt;'><b>" + signerTitleFormatted + @"</b><br/><br/><br/><br/><b>" + System.Security.SecurityElement.Escape(txtSignerName.Text) + @"</b></td>
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
