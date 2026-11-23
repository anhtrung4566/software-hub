$ErrorActionPreference = "Stop"
$hub = "C:\Users\User\.gemini\antigravity\scratch\software-hub"
$oa = Join-Path $hub "office-addins"
$pass = 0
$fail = 0

function Check($name, [scriptblock]$block) {
    try {
        & $block
        Write-Output "[PASS] $name"
        $script:pass++
    } catch {
        Write-Output "[FAIL] $name : $($_.Exception.Message)"
        $script:fail++
    }
}

$expected = @(
    @{ Dir = "VBTM.WordAddin"; Host = "Document"; Js = "word-addin.js" },
    @{ Dir = "VBTM.ExcelAddin"; Host = "Workbook"; Js = "excel-addin.js" },
    @{ Dir = "VBTM.PowerPointAddin"; Host = "Presentation"; Js = "powerpoint-addin.js" },
    @{ Dir = "VBTM.OutlookAddin"; Host = "Mailbox"; Js = "outlook-addin.js" }
)

foreach ($item in $expected) {
    $mPath = Join-Path $oa "$($item.Dir)\manifest.xml"
    Check "XML Manifest valid & HTTPS ($($item.Dir))" {
        [xml]$x = Get-Content $mPath -Raw -Encoding UTF8
        if ($x.OfficeApp.Hosts.Host.Name -ne $item.Host) { throw "Host mismatch" }
        $raw = Get-Content $mPath -Raw -Encoding UTF8
        if ($raw -notmatch "VersionOverrides") { throw "Missing VersionOverrides" }
        if ($raw -notmatch "https://anhtrung4566\.github\.io/software-hub/office-addins/") { throw "Missing HTTPS SourceLocation" }
    }
    $hPath = Join-Path $oa "$($item.Dir)\taskpane.html"
    Check "TaskPane uses shared VBTM.AICore ($($item.Dir))" {
        $html = Get-Content $hPath -Raw -Encoding UTF8
        if ($html -notmatch "\.\./VBTM\.AICore/vbtm-ai-core\.js") { throw "Missing shared vbtm-ai-core.js" }
        if ($html -notmatch "\.\./VBTM\.AICore/vbtm-compat\.js") { throw "Missing shared vbtm-compat.js" }
        if ($html -notmatch [regex]::Escape($item.Js)) { throw "Missing host script $($item.Js)" }
    }
}

Check "Security scan: 0 hardcoded Gemini API keys in office-addins/" {
    $files = Get-ChildItem -Path $oa -Recurse -File | Where-Object { $_.Extension -match "\.(js|html|xml|cs|json)$" -and $_.Name -ne "verify_office_addins.js" }
    foreach ($f in $files) {
        $c = Get-Content $f.FullName -Raw -Encoding UTF8
        if ($c -match "AIza[0-9A-Za-z_-]{20,}") { throw "Found hardcoded key in $($f.FullName)" }
    }
}

Check "Prompt sets complete (Word=18, Excel=11, PowerPoint=8, Outlook=8)" {
    $p = Get-Content (Join-Path $oa "VBTM.AICore\vbtm-prompts.js") -Raw -Encoding UTF8
    if ($p -notmatch "18\. Tạo nội dung từ mẫu chuẩn") { throw "Word 18 missing" }
    if ($p -notmatch "11\. Tạo bảng phân công nhiệm vụ") { throw "Excel 11 missing" }
    if ($p -notmatch "8\. Gợi ý mẫu trình chiếu") { throw "PPT 8 missing" }
    if ($p -notmatch "8\. Chuyển Email thành dự thảo Công văn") { throw "Outlook 8 missing" }
}

Check "VBTM_Pro_Office_Setup.exe exists & runs --check-only on real Office environment" {
    $exe = Join-Path $hub "downloads\VBTM_Pro_Office_Setup.exe"
    if (-not (Test-Path $exe)) { throw "EXE missing" }
    $p = Start-Process -FilePath $exe -ArgumentList "--check-only","--offline" -Wait -PassThru
    if ($p.ExitCode -ne 0) { throw "Exit code $($p.ExitCode)" }
    $rep = Join-Path $env:TEMP "vbtm_office_compat_report.txt"
    if (-not (Test-Path $rep)) { throw "Report file not created" }
    $txt = Get-Content $rep -Raw -Encoding UTF8
    if ($txt -notmatch "MỨC 3 – ADVANCED") { throw "Expected MỨC 3 – ADVANCED on Office 2024 x64" }
    Write-Output "--- REAL COMPATIBILITY CHECKER OUTPUT ---"
    Write-Output $txt.Trim()
}

Write-Output ""
Write-Output "=== OFFICE INTEGRATION VERIFICATION: $pass PASS, $fail FAIL ==="
if ($fail -gt 0) { exit 1 }
