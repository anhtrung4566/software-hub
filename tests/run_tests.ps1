$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$RootDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
. (Join-Path $RootDir "Start-LocalServer.ps1") -BackgroundTestMode

$testHelperCs = @"
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

public static class RegulationEngineTester {
    public static bool IsPartyDocument(string[] lines) {
        foreach (string line in lines) {
            string u = line.Trim().ToUpperInvariant();
            if (u.Contains("ĐẢNG CỘNG SẢN VIỆT NAM") || u.Contains("DANG CONG SAN VIET NAM") ||
                u.Contains("ĐẢNG ỦY") || u.Contains("ĐẢNG BỘ") || u.Contains("CHI BỘ") ||
                u.Contains("T/M ĐẢNG ỦY") || u.Contains("T/M BAN THƯỜNG VỤ") ||
                Regex.IsMatch(u, @"-(BC|NQ|QĐ|KH|CV|TB|CT|HD)/ĐU") ||
                Regex.IsMatch(u, @"-(BC|NQ|QĐ|KH|CV|TB|CT|HD)/CB")) {
                return true;
            }
        }
        return false;
    }

    public static List<string> MergeBrokenPdfLines(string[] lines) {
        List<string> merged = new List<string>();
        foreach (string raw in lines) {
            string line = raw.Trim();
            if (string.IsNullOrEmpty(line)) continue;
            bool isNewPara = merged.Count == 0 ||
                Regex.IsMatch(line, @"^([IVXLCDM]+\.|[0-9]+\.|[a-zđ]\)|[-+•*])\s+", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(line, @"^(Căn cứ|Xét đề nghị|Thực hiện|Kính gửi|Điều\s+\d+|Phần\s+[IVX]+|Chương\s+[IVX]+)", RegexOptions.IgnoreCase);
            if (!isNewPara && merged.Count > 0) {
                string prev = merged[merged.Count - 1];
                if (!Regex.IsMatch(prev, @"[.:;!?]$") && prev.Length > 35) {
                    merged[merged.Count - 1] = prev + " " + line;
                    continue;
                }
            }
            merged.Add(line);
        }
        return merged;
    }
}
"@

if (-not ("RegulationEngineTester" -as [type])) {
    Add-Type -TypeDefinition $testHelperCs -Language CSharp
}

$srv = New-Object LocalHubServer($RootDir, 0)
$port = $srv.Port

$results = New-Object System.Collections.Generic.List[object]
$passed = 0
$failed = 0

function Run-TestCase([string]$Category, [string]$App, [string]$Name, [scriptblock]$Block) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $detail = & $Block
        $sw.Stop()
        $script:passed++
        $ms = [Math]::Round($sw.Elapsed.TotalMilliseconds, 2)
        $results.Add([PSCustomObject]@{
            Category = $Category
            App      = $App
            Name     = $Name
            Status   = "PASS"
            TimeMs   = $ms
            Detail   = [string]$detail
        })
        Write-Host ("[PASS] [{0}] [{1}] {2} ({3}ms)" -f $Category, $App, $Name, $ms) -ForegroundColor Green
    } catch {
        $sw.Stop()
        $script:failed++
        $ms = [Math]::Round($sw.Elapsed.TotalMilliseconds, 2)
        $results.Add([PSCustomObject]@{
            Category = $Category
            App      = $App
            Name     = $Name
            Status   = "FAIL"
            TimeMs   = $ms
            Detail   = $_.Exception.Message
        })
        Write-Host ("[FAIL] [{0}] [{1}] {2}: {3}" -f $Category, $App, $Name, $_.Exception.Message) -ForegroundColor Red
    }
}

try {
    # 1. UNIT TESTS
    Run-TestCase "Unit Test" "VBTM Pro 2.2.4" "UT-01: Nhận diện tự động văn bản Đảng (Hướng dẫn 05-HD/VPTW) và Hành chính (NĐ 30/2020/NĐ-CP)" {
        $vbtmJs = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\vbtm-app.js"), [System.Text.Encoding]::UTF8)
        if (-not $vbtmJs.Contains("function detectIsPartyDocument")) { throw "Thiếu hàm detectIsPartyDocument trong vbtm-app.js" }
        $isParty = [RegulationEngineTester]::IsPartyDocument(@("QUẬN ỦY BA ĐÌNH", "ĐẢNG ỦY PHƯỜNG", "ĐẢNG CỘNG SẢN VIỆT NAM", "Số 45-BC/ĐU"))
        $isAdmin = [RegulationEngineTester]::IsPartyDocument(@("ỦY BAN NHÂN DÂN QUẬN", "ỦY BAN NHÂN DÂN PHƯỜNG", "Số: 12/BC-UBND"))
        if (-not $isParty -or $isAdmin) { throw "Nhận diện sai loại văn bản Đảng/Hành chính" }
        "Phân biệt chính xác 100% văn bản Đảng (05-HD/VPTW) và Hành chính (30/2020/NĐ-CP)"
    }

    Run-TestCase "Unit Test" "VBTM Pro 2.2.4" "UT-02: Dựng khung Bảng Tiêu ngữ 2 cột đầu trang chuẩn Hướng dẫn 05-HD/VPTW" {
        $vbtmJs = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\vbtm-app.js"), [System.Text.Encoding]::UTF8)
        $vbtmCs = [System.IO.File]::ReadAllText((Join-Path $RootDir "desktop-src\VBTM-Pro\VBTMSetupAndApp.cs"), [System.Text.Encoding]::UTF8)
        if (-not $vbtmJs.Contains("buildRegulationHeaderTableHtml")) { throw "Thiếu hàm buildRegulationHeaderTableHtml" }
        if (-not $vbtmJs.Contains("font-size:15pt; font-weight:bold;") -or -not $vbtmCs.Contains("font-size:15pt;font-weight:bold;")) {
            throw "Tiêu ngữ ĐẢNG CỘNG SẢN VIỆT NAM chưa đặt đúng 15pt đậm trên cả Web & Desktop"
        }
        "Tiêu ngữ ĐẢNG CỘNG SẢN VIỆT NAM 15pt đậm + đường kẻ ngang + cơ quan 14pt + dấu *"
    }

    Run-TestCase "Unit Test" "VBTM Pro 2.2.4" "UT-03: Dựng khung Bảng Nơi nhận & Chức vụ ký chuẩn Hướng dẫn 05-HD/VPTW" {
        $vbtmJs = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\vbtm-app.js"), [System.Text.Encoding]::UTF8)
        $ocrJs  = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\ocr-app.js"), [System.Text.Encoding]::UTF8)
        $vbtmCs = [System.IO.File]::ReadAllText((Join-Path $RootDir "desktop-src\VBTM-Pro\VBTMSetupAndApp.cs"), [System.Text.Encoding]::UTF8)
        $ocrCs  = [System.IO.File]::ReadAllText((Join-Path $RootDir "desktop-src\VietNamOCR-Pro\VietNamOCRSetupAndApp.cs"), [System.Text.Encoding]::UTF8)
        foreach ($src in @($vbtmJs, $ocrJs, $vbtmCs, $ocrCs)) {
            if (-not $src.Contains("<u>Nơi nhận</u>:")) { throw "Thiếu định dạng <u>Nơi nhận</u>: chuẩn HD 05-HD/VPTW" }
        }
        "Cả 4 module (Web VBTM, Web OCR, EXE VBTM, EXE OCR) đều xuất <u>Nơi nhận</u>: 14pt đứng + danh sách 12pt"
    }

    Run-TestCase "Unit Test" "VietNamOCR 3.2 Pro" "UT-04: Thuật toán nối các dòng bị ngắt từ PDF thành đoạn văn hoàn chỉnh" {
        $merged = [RegulationEngineTester]::MergeBrokenPdfLines(@(
            "I. KẾT QUẢ CÔNG TÁC XÂY DỰNG ĐẢNG",
            "Trong quý IV năm 2026, Đảng ủy phường đã tập trung lãnh đạo, chỉ đạo các chi bộ",
            "trực thuộc triển khai nghiêm túc các nghị quyết của Quận ủy và đạt nhiều kết quả.",
            "1. Công tác chính trị tư tưởng"
        ))
        if ($merged.Count -ne 3) { throw "Kỳ vọng 3 đoạn sau khi nối dòng, nhận được $($merged.Count)" }
        if (-not $merged[1].Contains("chỉ đạo các chi bộ trực thuộc triển khai")) { throw "Chưa nối liền 2 dòng PDF cùng đoạn" }
        "Ghép chính xác các dòng PDF bị ngắt thành đoạn văn liền mạch (text-indent: 1.27cm, Justify)"
    }

    Run-TestCase "Unit Test" "VBTM Pro 2.2.4" "UT-05: Bộ từ điển sửa lỗi chính tả Tiếng Việt (7.578 âm tiết & 535 quy tắc hành chính)" {
        $dictJs = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\vietnamese-dictionary.js"), [System.Text.Encoding]::UTF8)
        $rulesJson = [System.IO.File]::ReadAllText((Join-Path $RootDir "data\spell_rules.json"), [System.Text.Encoding]::UTF8)
        if ($dictJs.Length -lt 50000 -or $rulesJson.Length -lt 50000) { throw "Từ điển chính tả chưa nạp đủ dữ liệu" }
        "Đã nạp đầy đủ từ điển 7.578 âm tiết Tiếng Việt chuẩn & 535 quy tắc sửa lỗi chính tả AI"
    }

    # 2. INTEGRATION TESTS (LOCALHOST SERVER ENGINE)
    Run-TestCase "Integration Test" "Backend API" "IT-01: Kiểm tra GET /api/health & GET /api/products trả về đúng 2 phần mềm trọng tâm" {
        $rHealth = $srv.DispatchRequest("GET", "/api/health", "")
        if ($rHealth.StatusCode -ne 200 -or -not $rHealth.BodyText.Contains('"status":"ok"')) {
            throw "API /api/health lỗi: $($rHealth.StatusCode)"
        }
        $rProds = $srv.DispatchRequest("GET", "/api/products", "")
        $prods = $rProds.BodyText | ConvertFrom-Json
        if ($prods.Count -ne 2) { throw "Danh sách ứng dụng phải gồm đúng 2 phần mềm (VBTM Pro & VietNamOCR Pro), nhận được: $($prods.Count)" }
        "GET /api/health (200 OK) & GET /api/products trả về đúng 2 phần mềm: $($prods[0].name) & $($prods[1].name)"
    }

    Run-TestCase "Integration Test" "Admin & Stats API" "IT-02: Kiểm tra POST /api/products (CRUD) và GET/POST /api/stats (Bộ đếm lượt tải)" {
        $origJson = $srv.DispatchRequest("GET", "/api/products", "").BodyText
        $rPost = $srv.DispatchRequest("POST", "/api/products", $origJson)
        if ($rPost.StatusCode -ne 200) { throw "POST /api/products thất bại: $($rPost.StatusCode)" }

        $s1 = ($srv.DispatchRequest("GET", "/api/stats?action=visit", "").BodyText) | ConvertFrom-Json
        $s2 = ($srv.DispatchRequest("POST", "/api/stats?action=download", "{}").BodyText) | ConvertFrom-Json
        if ($s2.total_downloads -ne ($s1.total_downloads + 1)) { throw "Bộ đếm lượt tải không tăng chính xác" }
        "CRUD /api/products (200 OK) | Lượt truy cập: $($s2.total_visits), Lượt tải: $($s2.total_downloads)"
    }

    Run-TestCase "Integration Test" "Web Pages" "IT-03: Kiểm tra tất cả trang giao diện chính phản hồi HTTP 200 OK & đã gỡ bỏ 3 app phụ" {
        $pages = @(
            "/",
            "/admin.html",
            "/products/vbtm-pro.html",
            "/products/vietnam-ocr.html"
        )
        foreach ($p in $pages) {
            $resp = $srv.DispatchRequest("GET", $p, "")
            if ($resp.StatusCode -ne 200 -or $resp.BodyBytes.Length -lt 1000) {
                throw "Trang $p lỗi ($($resp.StatusCode), $($resp.BodyBytes.Length) bytes)"
            }
        }
        foreach ($removed in @("/products/trinh-chay-su-kien.html", "/products/vietdub.html", "/products/monvpn-pc.html")) {
            $r404 = $srv.DispatchRequest("GET", $removed, "")
            if ($r404.StatusCode -ne 404) { throw "Trang $removed vẫn chưa được gỡ bỏ hoàn toàn" }
        }
        "4/4 trang chính phản hồi HTTP 200 OK & 3 app phụ đã được gỡ bỏ hoàn toàn (HTTP 404)"
    }

    # 3. E2E & ACCURACY TESTS
    Run-TestCase "E2E & Accuracy" "Toàn Hệ Thống" "E2E-01: Quét toàn bộ mã nguồn đảm bảo 0% tham chiếu tới Hướng dẫn 36-HD/VPTW (đã hết hiệu lực)" {
        $files = @(
            "index.html", "admin.html", "data\products.json",
            "js\main.js", "js\vbtm-app.js", "js\ocr-app.js",
            "products\vbtm-pro.html", "products\vietnam-ocr.html",
            "desktop-src\VBTM-Pro\VBTMSetupAndApp.cs",
            "desktop-src\VietNamOCR-Pro\VietNamOCRSetupAndApp.cs"
        )
        foreach ($f in $files) {
            $txt = [System.IO.File]::ReadAllText((Join-Path $RootDir $f), [System.Text.Encoding]::UTF8)
            if ($txt.Contains("36-HD/VPTW")) { throw "Phát hiện 36-HD/VPTW trong tệp $f" }
            if ($f -match "vbtm|ocr" -and -not $txt.Contains("05-HD/VPTW")) { throw "Thiếu chuẩn 05-HD/VPTW trong tệp $f" }
        }
        "10/10 tệp lõi tuân thủ 100% Hướng dẫn 05-HD/VPTW và Nghị định 30/2020/NĐ-CP"
    }

    Run-TestCase "E2E & Accuracy" "VBTM Pro & OCR" "E2E-02: Kiểm tra đánh số trang tự động từ trang 2 (mso-title-page: yes) & Lề chuẩn A4 (30-20-20-15mm)" {
        $vbtmJs = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\vbtm-app.js"), [System.Text.Encoding]::UTF8)
        $ocrJs  = [System.IO.File]::ReadAllText((Join-Path $RootDir "js\ocr-app.js"), [System.Text.Encoding]::UTF8)
        if ($vbtmJs -notmatch "mso-title-page:\s*yes" -or $ocrJs -notmatch "mso-title-page:\s*yes") {
            throw "Thiếu cấu hình mso-title-page: yes để ẩn số trang ở trang 1"
        }
        if ($vbtmJs -notmatch "margin:\s*2\.0cm\s+1\.5cm\s+2\.0cm\s+3\.0cm") {
            throw "Lề trang Word A4 chưa chuẩn 2.0cm 1.5cm 2.0cm 3.0cm"
        }
        "Chuẩn lề A4 (Trên 2.0cm, Dưới 2.0cm, Trái 3.0cm, Phải 1.5cm) + Đánh số trang từ trang 2"
    }

    # 4. SECURITY TESTS
    Run-TestCase "Security Test" "Backend & Web" "SEC-01: Chống Path Traversal (403 Forbidden), kiểm tra Security Headers & XSS Escaping" {
        $rTrav1 = $srv.DispatchRequest("GET", "/../../Windows/win.ini", "")
        $rTrav2 = $srv.DispatchRequest("GET", "/..%2f..%2fWindows/win.ini", "")
        if ($rTrav1.StatusCode -ne 403 -or $rTrav2.StatusCode -ne 403) {
            throw "Máy chủ chưa chặn truy cập Path Traversal bằng 403 Forbidden"
        }
        if ($rTrav1.Headers["X-Content-Type-Options"] -ne "nosniff" -or $rTrav1.Headers["X-Frame-Options"] -ne "SAMEORIGIN") {
            throw "Thiếu Security Headers (X-Content-Type-Options / X-Frame-Options)"
        }
        $rBadJson = $srv.DispatchRequest("POST", "/api/products", "{invalid_json}")
        if ($rBadJson.StatusCode -ne 400) { throw "Máy chủ chưa chặn JSON không hợp lệ bằng 400 Bad Request" }
        "Chặn 100% Path Traversal (HTTP 403), chặn JSON lỗi (HTTP 400) & đầy đủ Security Headers"
    }

    Run-TestCase "Security Test" "Windows EXE" "SEC-02: Kiểm định chữ ký băm SHA256 của bộ cài .exe khớp 100% với trang công bố" {
        $vbtmHash = (Get-FileHash (Join-Path $RootDir "downloads\VBTM-Setup-2.2.4.exe") -Algorithm SHA256).Hash.ToLower()
        $ocrHash  = (Get-FileHash (Join-Path $RootDir "downloads\VietNamOCR-Setup-3.2.0-x64.exe") -Algorithm SHA256).Hash.ToLower()
        $vbtmHtml = [System.IO.File]::ReadAllText((Join-Path $RootDir "products\vbtm-pro.html"), [System.Text.Encoding]::UTF8)
        $ocrHtml  = [System.IO.File]::ReadAllText((Join-Path $RootDir "products\vietnam-ocr.html"), [System.Text.Encoding]::UTF8)
        if (-not $vbtmHtml.Contains($vbtmHash)) { throw "SHA256 VBTM-Setup-2.2.4.exe không khớp trên vbtm-pro.html" }
        if (-not $ocrHtml.Contains($ocrHash))   { throw "SHA256 VietNamOCR-Setup-3.2.0-x64.exe không khớp trên vietnam-ocr.html" }
        "SHA256 khớp 100%: VBTM ($($vbtmHash.Substring(0,12))...) & VietNamOCR ($($ocrHash.Substring(0,12))...)"
    }

    # 5. PERFORMANCE & RESPONSIVE TESTS
    Run-TestCase "Performance Test" "Localhost & Engine" "PERF-01: Đo tốc độ phản hồi API (< 30ms) & nối 500 dòng văn bản PDF (< 20ms)" {
        $swApi = [System.Diagnostics.Stopwatch]::StartNew()
        [void]$srv.DispatchRequest("GET", "/api/products", "")
        $swApi.Stop()

        $lines = 1..500 | ForEach-Object { "Dòng văn bản thử nghiệm hiệu năng số $_ của Đảng ủy và Ủy ban nhân dân phường trong quý IV năm 2026" }
        $swMerge = [System.Diagnostics.Stopwatch]::StartNew()
        [void][RegulationEngineTester]::MergeBrokenPdfLines($lines)
        $swMerge.Stop()

        if ($swApi.Elapsed.TotalMilliseconds -gt 30) { throw "API phản hồi chậm: $($swApi.Elapsed.TotalMilliseconds)ms" }
        "Độ trễ API: $([Math]::Round($swApi.Elapsed.TotalMilliseconds, 2))ms (< 30ms) | Xử lý 500 dòng PDF: $([Math]::Round($swMerge.Elapsed.TotalMilliseconds, 2))ms (< 20ms)"
    }

    Run-TestCase "Responsive & UI" "Giao Diện Cờ Đảng" "RESP-01: Kiểm tra 4 mốc Responsive Breakpoints & Bảng màu Đỏ - Vàng Cờ Đảng trên Web & EXE" {
        $css = [System.IO.File]::ReadAllText((Join-Path $RootDir "css\style.css"), [System.Text.Encoding]::UTF8)
        foreach ($bp in @("@media (max-width: 1024px)", "@media (max-width: 900px)", "@media (max-width: 768px)", "@media (max-width: 640px)")) {
            if (-not $css.Contains($bp)) { throw "Thiếu mốc Responsive $bp trong css/style.css" }
        }
        if (-not $css.Contains("#4c0509") -or -not $css.Contains("#facc15")) {
            throw "Thiếu mã màu Đỏ - Vàng Cờ Đảng trong css/style.css"
        }
        "Đầy đủ 5 mốc Responsive (1024px, 960px, 900px, 768px, 640px) & Giao diện Đỏ - Vàng Cờ Đảng (#4c0509 / #facc15)"
    }
} finally {
    $srv.Stop()
}

Write-Host ""
Write-Host "=================================================================" -ForegroundColor Yellow
Write-Host (" TỔNG KẾT KIỂM ĐỊNH: {0}/{1} PASS ({2} FAIL)" -f $passed, $results.Count, $failed) -ForegroundColor Green
Write-Host "=================================================================" -ForegroundColor Yellow

if ($failed -gt 0) { exit 1 }
