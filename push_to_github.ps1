$ErrorActionPreference = "Stop"
$gh = "C:\Program Files\GitHub CLI\gh.exe"
$token = (& $gh auth token).Trim()
$owner = "anhtrung4566"
$repo = "software-hub"
$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$headers = @{
    "Authorization" = "Bearer $token"
    "Accept"        = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent"    = "Antigravity-GitHub-Uploader"
}

$removedPaths = @(
    "products/trinh-chay-su-kien.html",
    "products/vietdub.html",
    "products/monvpn-pc.html",
    "js/sukien-app.js"
)
foreach ($delRel in $removedPaths) {
    try {
        $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$delRel" -Headers $headers -Method Get
        if ($existing -and $existing.sha) {
            $delBody = @{
                message = "Remove $delRel"
                sha     = $existing.sha
                branch  = "main"
            } | ConvertTo-Json
            Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$delRel" -Headers $headers -Method Delete -Body $delBody -ContentType "application/json; charset=utf-8" | Out-Null
            Write-Host "Deleted from GitHub: $delRel"
        }
    } catch {}
}

$sha1 = [System.Security.Cryptography.SHA1]::Create()
$files = Get-ChildItem -Path $baseDir -Recurse -File | Where-Object {
    $_.FullName -notmatch "\\\.github\\"
}

foreach ($file in $files) {
    $relPath = $file.FullName.Substring($baseDir.Length + 1).Replace("\", "/")
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

    $prefixBytes = [System.Text.Encoding]::ASCII.GetBytes("blob $($bytes.Length)`0")
    $combined = New-Object byte[] ($prefixBytes.Length + $bytes.Length)
    [Array]::Copy($prefixBytes, 0, $combined, 0, $prefixBytes.Length)
    [Array]::Copy($bytes, 0, $combined, $prefixBytes.Length, $bytes.Length)
    $localGitSha = ([BitConverter]::ToString($sha1.ComputeHash($combined))).Replace("-", "").ToLowerInvariant()

    $existingSha = $null
    try {
        $existing = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/contents/$relPath" -Headers $headers -Method Get
        $existingSha = $existing.sha
    } catch {}

    if ($existingSha -eq $localGitSha) {
        Write-Host "Skipped (unchanged): $relPath"
        continue
    }

    $b64 = [Convert]::ToBase64String($bytes)
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
    Write-Host "Uploaded: $relPath ($($bytes.Length) bytes)"
}

Write-Host "DONE! Live URL: https://$owner.github.io/$repo/"
