using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using Microsoft.Win32;

namespace VBTM.OfficeIntegration
{
    public class OfficeAppStatus
    {
        public string Name;
        public string ExeName;
        public bool Installed;
        public string ExePath;
        public string FileVersion;
        public string SupportedTier; // BASIC / STANDARD / ADVANCED / UNSUPPORTED
        public string Notes;
    }

    public class OfficeEnvironmentReport
    {
        public string WindowsVersion;
        public bool Is64BitOs;
        public string OfficeChannel;     // Click-to-Run / MSI / Not Detected
        public string OfficeVersion;     // e.g. 16.0.17932.21140
        public string OfficeBuild;       // e.g. 17932.21140
        public string OfficeBitness;     // x64 / x86 / Unknown
        public string ProductReleaseIds; // e.g. ProPlus2024Volume / O365ProPlusRetail
        public string OutlookFlavor;     // Classic Outlook / New Outlook / Both / None
        public bool WebView2Installed;
        public string WebView2Version;
        public bool HttpsReachable;
        public string HttpsCheckDetail;
        public string OverallTier;       // MỨC 1 – BASIC / MỨC 2 – STANDARD / MỨC 3 – ADVANCED
        public List<OfficeAppStatus> Apps = new List<OfficeAppStatus>();

        public string ToSummaryText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("=== BÁO CÁO TƯƠNG THÍCH VBTM PRO OFFICE (WINDOWS PERSONAL EDITION) ===");
            sb.AppendLine("Hệ điều hành : " + WindowsVersion + " (" + (Is64BitOs ? "64-bit" : "32-bit") + ")");
            sb.AppendLine("Bộ Office    : " + OfficeChannel + " | Phiên bản: " + OfficeVersion + " | Kiến trúc: " + OfficeBitness);
            if (!string.IsNullOrEmpty(ProductReleaseIds))
                sb.AppendLine("Gói sản phẩm : " + ProductReleaseIds);
            sb.AppendLine("Outlook      : " + OutlookFlavor);
            sb.AppendLine("WebView2     : " + (WebView2Installed ? ("Đã cài (" + WebView2Version + ")") : "Chưa phát hiện (Office 2016/2019 cũ có thể cần cài Edge WebView2 Runtime)"));
            sb.AppendLine("Kết nối HTTPS: " + (HttpsReachable ? "OK — " : "Cảnh báo — ") + HttpsCheckDetail);
            sb.AppendLine("Mức tương thích tự động: " + OverallTier);
            sb.AppendLine();
            sb.AppendLine("Chi tiết từng ứng dụng:");
            foreach (var a in Apps)
            {
                sb.AppendLine(string.Format("  - {0,-11}: {1,-10} | Ver: {2,-18} | Mức: {3} | {4}",
                    a.Name, a.Installed ? "Đã cài" : "Chưa có",
                    string.IsNullOrEmpty(a.FileVersion) ? "-" : a.FileVersion,
                    a.SupportedTier, a.Notes));
            }
            return sb.ToString();
        }
    }

    public static class CompatibilityChecker
    {
        public static OfficeEnvironmentReport Inspect(bool checkHttps = true)
        {
            OfficeEnvironmentReport r = new OfficeEnvironmentReport();
            r.WindowsVersion = Environment.OSVersion.VersionString;
            r.Is64BitOs = Environment.Is64BitOperatingSystem;
            r.OfficeChannel = "Chưa phát hiện";
            r.OfficeVersion = "-";
            r.OfficeBuild = "-";
            r.OfficeBitness = "Unknown";
            r.ProductReleaseIds = "";
            r.OutlookFlavor = "Chưa cài";

            // 1. Kiểm tra Click-to-Run (Office 2016/2019/2021/2024/M365)
            try
            {
                using (RegistryKey c2r = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                {
                    if (c2r != null)
                    {
                        r.OfficeChannel = "Click-to-Run";
                        r.OfficeVersion = Convert.ToString(c2r.GetValue("VersionToReport") ?? "-");
                        r.OfficeBitness = Convert.ToString(c2r.GetValue("Platform") ?? "Unknown");
                        r.ProductReleaseIds = Convert.ToString(c2r.GetValue("ProductReleaseIds") ?? "");
                        string[] parts = r.OfficeVersion.Split('.');
                        if (parts.Length >= 4) r.OfficeBuild = parts[2] + "." + parts[3];
                    }
                }
            }
            catch { }

            // 2. Fallback kiểm tra Office MSI (Office 2016 MSI)
            if (r.OfficeChannel == "Chưa phát hiện")
            {
                foreach (string ver in new[] { "16.0", "15.0" })
                {
                    try
                    {
                        using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\" + ver + @"\Common\InstallRoot"))
                        {
                            if (k != null && k.GetValue("Path") != null)
                            {
                                r.OfficeChannel = "MSI";
                                r.OfficeVersion = ver + ".0.0 (MSI)";
                                r.OfficeBitness = "x64";
                                break;
                            }
                        }
                        using (RegistryKey k32 = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Office\" + ver + @"\Common\InstallRoot"))
                        {
                            if (k32 != null && k32.GetValue("Path") != null)
                            {
                                r.OfficeChannel = "MSI";
                                r.OfficeVersion = ver + ".0.0 (MSI)";
                                r.OfficeBitness = "x86";
                                break;
                            }
                        }
                    }
                    catch { }
                }
            }

            // 3. Kiểm tra WebView2 Runtime
            r.WebView2Version = DetectWebView2();
            r.WebView2Installed = !string.IsNullOrEmpty(r.WebView2Version);

            // 4. Kiểm tra từng ứng dụng Word, Excel, PowerPoint, Outlook
            int buildNum = ParseBuildNumber(r.OfficeVersion);
            r.OverallTier = MapBuildToTier(buildNum, r.ProductReleaseIds, r.WebView2Installed);

            r.Apps.Add(InspectApp("Word", "WINWORD.EXE", buildNum, r.WebView2Installed));
            r.Apps.Add(InspectApp("Excel", "EXCEL.EXE", buildNum, r.WebView2Installed));
            r.Apps.Add(InspectApp("PowerPoint", "POWERPNT.EXE", buildNum, r.WebView2Installed));
            var outlookApp = InspectApp("Outlook", "OUTLOOK.EXE", buildNum, r.WebView2Installed);
            r.Apps.Add(outlookApp);

            // 5. Phân biệt Classic Outlook & New Outlook (Olk.exe / Appx)
            bool hasClassic = outlookApp.Installed;
            bool hasNewOutlook = DetectNewOutlook();
            if (hasClassic && hasNewOutlook) r.OutlookFlavor = "Cả Classic Outlook & New Outlook";
            else if (hasClassic) r.OutlookFlavor = "Classic Outlook (Win32)";
            else if (hasNewOutlook) r.OutlookFlavor = "New Outlook for Windows";
            else r.OutlookFlavor = "Chưa cài Outlook";

            // 6. Kiểm tra kết nối HTTPS tới máy chủ lưu Add-in (GitHub Pages)
            if (checkHttps)
            {
                CheckHttpsEndpoint("https://anhtrung4566.github.io/software-hub/office-addins/VBTM.WordAddin/taskpane.html", out r.HttpsReachable, out r.HttpsCheckDetail);
            }
            else
            {
                r.HttpsReachable = false;
                r.HttpsCheckDetail = "Bỏ qua kiểm tra mạng (chế độ offline).";
            }

            return r;
        }

        private static OfficeAppStatus InspectApp(string name, string exeName, int buildNum, bool hasWebView2)
        {
            OfficeAppStatus st = new OfficeAppStatus { Name = name, ExeName = exeName, Installed = false, SupportedTier = "N/A", Notes = "Chưa cài đặt" };
            string path = FindAppPath(exeName);
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                st.Installed = true;
                st.ExePath = path;
                try
                {
                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(path);
                    st.FileVersion = fvi.ProductVersion;
                    int appBuild = ParseBuildNumber(st.FileVersion);
                    if (appBuild > buildNum) buildNum = appBuild;
                }
                catch { }

                if (buildNum >= 14326)
                {
                    st.SupportedTier = "ADVANCED";
                    st.Notes = "Hỗ trợ đầy đủ Office.js Add-in (Mức 3)";
                }
                else if (buildNum >= 10325)
                {
                    st.SupportedTier = "STANDARD";
                    st.Notes = "Hỗ trợ tốt Office.js Add-in (Mức 2)";
                }
                else if (buildNum > 0)
                {
                    st.SupportedTier = "BASIC";
                    st.Notes = hasWebView2 ? "Hỗ trợ cơ bản (Mức 1)" : "Cần Edge WebView2 Runtime để hiển thị Task Pane ổn định";
                }
                else
                {
                    st.SupportedTier = "STANDARD";
                    st.Notes = "Đã phát hiện tệp thực thi Office";
                }
            }
            return st;
        }

        private static string FindAppPath(string exeName)
        {
            string[] roots = new[] {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + exeName,
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\" + exeName
            };
            foreach (string sub in roots)
            {
                try
                {
                    using (RegistryKey k = Registry.LocalMachine.OpenSubKey(sub))
                    {
                        if (k != null)
                        {
                            string p = Convert.ToString(k.GetValue("") ?? "").Trim('"');
                            if (!string.IsNullOrEmpty(p) && File.Exists(p)) return p;
                        }
                    }
                }
                catch { }
            }
            return null;
        }

        private static string DetectWebView2()
        {
            string guid = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";
            try
            {
                using (RegistryKey k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\" + guid))
                    if (k != null && k.GetValue("pv") != null) return Convert.ToString(k.GetValue("pv"));
            }
            catch { }
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\EdgeUpdate\Clients\" + guid))
                    if (k != null && k.GetValue("pv") != null) return Convert.ToString(k.GetValue("pv"));
            }
            catch { }
            return null;
        }

        private static bool DetectNewOutlook()
        {
            try
            {
                string localApps = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "olk.exe");
                if (File.Exists(localApps)) return true;
            }
            catch { }
            return false;
        }

        private static int ParseBuildNumber(string ver)
        {
            if (string.IsNullOrEmpty(ver)) return 0;
            string[] p = ver.Split('.');
            int b = 0;
            if (p.Length >= 3) int.TryParse(p[2], out b);
            return b;
        }

        private static string MapBuildToTier(int buildNum, string releaseIds, bool hasWv2)
        {
            if (buildNum >= 14326 || (releaseIds != null && releaseIds.IndexOf("365", StringComparison.OrdinalIgnoreCase) >= 0))
                return "MỨC 3 – ADVANCED (Office 2021/2024/M365 · Build " + buildNum + ")";
            if (buildNum >= 10325)
                return "MỨC 2 – STANDARD (Office 2019/C2R · Build " + buildNum + ")";
            if (buildNum > 0)
                return "MỨC 1 – BASIC (Office 2016 · Build " + buildNum + (hasWv2 ? "" : " · Khuyến nghị cài WebView2") + ")";
            return "MỨC 1 – BASIC (Chế độ an toàn)";
        }

        private static void CheckHttpsEndpoint(string url, out bool ok, out string detail)
        {
            ok = false;
            try
            {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072; // TLS 1.2
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
                req.Method = "HEAD";
                req.Timeout = 8000;
                req.KeepAlive = false;
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                {
                    ok = ((int)resp.StatusCode >= 200 && (int)resp.StatusCode < 400);
                    detail = "HTTP " + (int)resp.StatusCode + " (" + resp.StatusDescription + ")";
                }
            }
            catch (WebException wex)
            {
                var r = wex.Response as HttpWebResponse;
                if (r != null)
                    detail = "HTTP " + (int)r.StatusCode + " (có thể chưa publish lên GitHub Pages)";
                else
                    detail = "Không kết nối được HTTPS (" + wex.Status + ")";
            }
            catch (Exception ex)
            {
                detail = ex.Message;
            }
        }
    }
}
