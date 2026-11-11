$ErrorActionPreference = "Stop"
$gh = "C:\Program Files\GitHub CLI\gh.exe"
$token = (& $gh auth token).Trim()
$owner = "anhtrung4566"
$repo = "software-hub"
$baseDir = "C:\Users\User\.gemini\antigravity\scratch\software-hub"

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"    = "Antigravity-GitHub-Uploader"
}

Write-Host "1. Đang kiểm tra kho lưu trữ GitHub $owner/$repo..."
try {
    $repoInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo" -Headers $headers -Method Get
    Write-Host "Kho lưu trữ: $($repoInfo.html_url)"
} catch {
    $createBody = @{
        name        = $repo
        description = "Kho Phần Mềm Chuyên Dụng Đảng Ủy & UBND Phường · VBTM Pro 2.2.4 & VietNamOCR 3.2 Pro"
        private     = $false
        auto_init   = $true
    } | ConvertTo-Json
    $repoInfo = Invoke-RestMethod -Uri "https://api.github.com/user/repos" -Headers $headers -Method Post -Body $createBody -ContentType "application/json; charset=utf-8"
    Write-Host "Đã tạo kho lưu trữ: $($repoInfo.html_url)"
    Start-Sleep -Seconds 2
}

$files = Get-ChildItem -Path $baseDir -Recurse -File | Where-Object {
    $_.FullName -notmatch "\\\.github\\"
}

foreach ($file in $files) {
    $relPath = $file.FullName.Substring($baseDir.Length + 1).Replace("\", "/")
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $b64 = [Convert]::ToBase64String($bytes)

    $existingSha = $null
    try {
        $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$relPath" -Headers $headers -Method Get
        $existingSha = $existing.sha
    } catch {}

    $putObj = @{
        message = "Update $relPath"
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

Write-Host "2. Đang kích hoạt GitHub Pages..."
try {
    $pagesBody = @{
        source = @{
            branch = "main"
            path   = "/"
        }
    } | ConvertTo-Json -Depth 5
    $pagesInfo = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/pages" -Headers $headers -Method Post -Body $pagesBody -ContentType "application/json; charset=utf-8"
    Write-Host "GitHub Pages: $($pagesInfo.html_url)"
} catch {
    Write-Host "GitHub Pages đã hoạt động trên nhánh main."
}

Write-Host "HOÀN TẤT! Kho lưu trữ: https://github.com/$owner/$repo"
Write-Host "Trang web trực tuyến: https://$owner.github.io/$repo/"
