$ErrorActionPreference = "Stop"
$baseDir = "C:\Users\User\.gemini\antigravity\scratch\software-hub"
$dlDir = Join-Path $baseDir "downloads"

# 0. Cập nhật gói ZIP
Add-Type -AssemblyName System.IO.Compression.FileSystem
$vbtmZip = "$dlDir\VBTM-Setup-2.2.4.zip"
$ocrZip = "$dlDir\VietNamOCR-Setup-3.2.0-x64.zip"
if (Test-Path $vbtmZip) { Remove-Item $vbtmZip -Force }
if (Test-Path $ocrZip) { Remove-Item $ocrZip -Force }

$tmpVbtm = Join-Path $env:TEMP ("vbtm_pkg_" + [Guid]::NewGuid().ToString("N"))
$tmpOcr = Join-Path $env:TEMP ("ocr_pkg_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmpVbtm | Out-Null
New-Item -ItemType Directory -Path $tmpOcr | Out-Null
Copy-Item "$dlDir\VBTM-Setup-2.2.4.exe" "$tmpVbtm\VBTM-Setup-2.2.4.exe"
Copy-Item "$dlDir\VietNamOCR-Setup-3.2.0-x64.exe" "$tmpOcr\VietNamOCR-Setup-3.2.0-x64.exe"
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmpVbtm, $vbtmZip)
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmpOcr, $ocrZip)
Remove-Item $tmpVbtm -Recurse -Force
Remove-Item $tmpOcr -Recurse -Force

# 1. Cập nhật trực tiếp bản đã cài đặt trên máy tính của người dùng
try {
    Get-Process -Name "VietNamOCR*", "VBTM*" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 300
    $ocrInstallDir = Join-Path $env:LOCALAPPDATA "Programs\VietNamOCR-Pro"
    if (!(Test-Path $ocrInstallDir)) { New-Item -ItemType Directory -Path $ocrInstallDir -Force | Out-Null }
    Copy-Item "$dlDir\VietNamOCR-Setup-3.2.0-x64.exe" "$ocrInstallDir\VietNamOCR.exe" -Force -ErrorAction SilentlyContinue
    Copy-Item "$dlDir\VietNamOCR-Setup-3.2.0-x64.exe" "$ocrInstallDir\VietNamOCR-Pro.exe" -Force -ErrorAction SilentlyContinue

    $vbtmInstallDir = Join-Path $env:LOCALAPPDATA "Programs\VBTM-Pro"
    if (!(Test-Path $vbtmInstallDir)) { New-Item -ItemType Directory -Path $vbtmInstallDir -Force | Out-Null }
    Copy-Item "$dlDir\VBTM-Setup-2.2.4.exe" "$vbtmInstallDir\VBTM-Pro.exe" -Force -ErrorAction SilentlyContinue
    Write-Host "1. Đã cập nhật trực tiếp bản 4.0 Ultimate trên máy tính (VietNamOCR.exe & VBTM-Pro.exe)."
} catch {
    Write-Host "Lưu ý khi cập nhật bản cài trên máy: $($_.Exception.Message)"
}

$gh = "C:\Program Files\GitHub CLI\gh.exe"
$token = (& $gh auth token).Trim()
$owner = "anhtrung4566"
$repo = "software-hub"

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"    = "Antigravity-GitHub-Uploader"
}

$targetFiles = @(
    "index.html",
    "js/main.js",
    "data/products.json",
    "products/vietnam-ocr.html",
    "products/vbtm-pro.html",
    "js/ocr-app.js",
    "js/vbtm-app.js",
    "desktop-src/VietNamOCR-Pro/VietNamOCRSetupAndApp.cs",
    "desktop-src/VBTM-Pro/VBTMSetupAndApp.cs",
    "downloads/VietNamOCR-Setup-3.2.0-x64.exe",
    "downloads/VietNamOCR-Setup-3.2.0-x64.zip",
    "downloads/VBTM-Setup-2.2.4.exe",
    "downloads/VBTM-Setup-2.2.4.zip"
)

foreach ($relPath in $targetFiles) {
    $fullPath = Join-Path $baseDir ($relPath.Replace("/", "\"))
    if (!(Test-Path $fullPath)) { continue }
    $bytes = [System.IO.File]::ReadAllBytes($fullPath)
    $b64 = [Convert]::ToBase64String($bytes)

    $existingSha = $null
    try {
        $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$relPath" -Headers $headers -Method Get
        $existingSha = $existing.sha
    } catch {}

    $putObj = @{
        message = "Upgrade to 4.0 Ultimate with 1:1 layout-preserving PDF to Word in $relPath"
        content = $b64
        branch  = "main"
    }
    if ($existingSha) {
        $putObj["sha"] = $existingSha
    }

    $putBody = $putObj | ConvertTo-Json -Depth 5
    Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$relPath" -Headers $headers -Method Put -Body $putBody -ContentType "application/json; charset=utf-8" | Out-Null
    Write-Host "Đã đẩy lên GitHub: $relPath ($($bytes.Length) bytes)"
}

Write-Host "HOÀN TẤT! Kho lưu trữ: https://github.com/$owner/$repo"
Write-Host "Trang web trực tuyến: https://$owner.github.io/$repo/"
