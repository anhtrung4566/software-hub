$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$hub = "C:\Users\User\.gemini\antigravity\scratch\software-hub"
$exe = Join-Path $hub "downloads\VBTM-Setup-3.6.0.exe"
$testDocx = Join-Path $env:TEMP ("vbtm_spell_benchmark_" + [DateTime]::Now.ToString("HHmmss") + ".docx")
$outTxt = Join-Path $env:TEMP "vbtm_test_spell_docx.txt"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 1. Tạo tệp Word .docx thật có đủ: Header, Footer, Đoạn văn, Bảng biểu 4x3, Chú thích (Footnote) và Tên riêng/Mã hiệu chuẩn
$contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>'

$headerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>ỦY BẠN NHÂN DÂN PHƯỜNG TÂN SƠN NHẤT — Sổ: 15/KH-UBND</w:t></w:r></w:p></w:hdr>'

$footerXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Trang 1 — Tài liệu báo cao định kỳ của UBND Phường Tân Sơn Nhất</w:t></w:r></w:p></w:ftr>'

$footnoteXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>Ghi chú: Quán triệt chủ chương mới theo Hướng dẫn 05-HD/VPTW.</w:t></w:r></w:p></w:footnote></w:footnotes>'

$docXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>KẾ HOẠCHHH</w:t></w:r></w:p>
    <w:p><w:r><w:t>Đồng chí Nguyễn Văn An, Phó ban phụ trách đơn vị PC09-Đ3 tại TP. Đà Nẵng đã tố chức hội nghị phù họp với Nghị định 30/2020/NĐ-CP ngày 07/9/2026 và kinh phí 15.500.000 VNĐ.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Các đơn vị thực hiện sát nhập tổ bộ máy, bổ xung nhân sự, đề suất chỉn sửa quy trình, giữ vững kỹ luật và đánh giá xác xuất rủi ro theo đúng quy định , không để sai sót.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Trong Tuần36 cần thựchiện tốt côngtác tuyên truyền, tập treung lãnh đạo cqác bộ phận, không để các các tồn tại kéo dài theo quy định hiện hành. triển khai đồng bộ ngay từ đầu tháng.</w:t></w:r></w:p>
    <w:p><w:r><w:t>Tập thể công nhân và cán bộ HĐND luôn luôn đoàn kết, ngày ngày phấn đấu, bộ gửi hồ sơ đầy đủ đúng hạn.</w:t></w:r></w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>STT</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Nội dung nhiệm vụ</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Đơn vị thực hiện</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Phối họp giải phong mặt bằng tuyến đường chính</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>UBND Phường Tân Sơn Nhất</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Thàdnh lập tổ công tác phòng chống xunga đột cơ sở</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Công an phường</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Cán bộ năm giữ vai trò nòng cốt, ghi rõ nơi nhân</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Văn phòng Đảng ủy</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
'@

if (Test-Path $testDocx) { Remove-Item $testDocx -Force }
$fs = [System.IO.File]::Open($testDocx, [System.IO.FileMode]::Create)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
function Add-ZipEntry($z, $name, $text) {
    $e = $z.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write($text)
    $sw.Dispose()
}
Add-ZipEntry $zip "[Content_Types].xml" $contentTypes
Add-ZipEntry $zip "word/document.xml" $docXml
Add-ZipEntry $zip "word/header1.xml" $headerXml
Add-ZipEntry $zip "word/footer1.xml" $footerXml
Add-ZipEntry $zip "word/footnotes.xml" $footnoteXml
$zip.Dispose()
$fs.Dispose()

# 2. Chạy VBTM Pro 3.6 trực tiếp trên tệp .docx vừa tạo
if (Test-Path $outTxt) { Remove-Item $outTxt -Force }
& $exe --test-spell-docx $testDocx
Start-Sleep -Milliseconds 900

if (-not (Test-Path $outTxt)) { throw "Khong tao duoc ket qua $outTxt" }
$lines = Get-Content $outTxt -Encoding UTF8
$headerLine = $lines[0]
$issueLines = $lines | Select-Object -Skip 1

# 3. Đáp án chuẩn: 27 lỗi đã cài trên đủ Header, Footer, Footnote, Đoạn văn & Bảng biểu
$expectedErrors = @(
    "ỦY BẠN", "Sổ", "KẾ HOẠCHHH", "tố chức", "phù họp",
    "sát nhập", "bổ xung", "đề suất", "chỉn sửa", "kỹ luật", "xác xuất",
    "định ,", "Tuần36", "thựchiện", "côngtác", "tập treung", "cqác",
    "các các", "hành. triển", "Phối họp", "giải phong", "Thàdnh lập",
    "xunga đột", "năm giữ", "nơi nhân", "báo cao", "chủ chương"
)

# Các thực thể/từ đúng không được báo nhầm (15 mục)
$protectedClean = @(
    "Nguyễn Văn An", "Tân Sơn Nhất", "UBND", "HĐND", "PC09-Đ3",
    "30/2020/NĐ-CP", "05-HD/VPTW", "07/9/2026", "15.500.000", "VNĐ",
    "luôn luôn", "ngày ngày", "bộ gửi", "Phó ban", "công nhân"
)

$tp = 0
$missed = @()
foreach ($exp in $expectedErrors) {
    $hit = $issueLines | Where-Object { $_ -like "*$exp ->*" }
    if ($hit) { $tp++ } else { $missed += $exp }
}

$fpList = @()
foreach ($ln in $issueLines) {
    $isExpected = $false
    foreach ($exp in $expectedErrors) {
        if ($ln -like "*$exp ->*") { $isExpected = $true; break }
    }
    if (-not $isExpected) { $fpList += $ln }
}

$protectedHit = @()
foreach ($pr in $protectedClean) {
    $hit = $issueLines | Where-Object { $_ -like "*$pr ->*" }
    if ($hit) { $protectedHit += $pr }
}

$totalExp = $expectedErrors.Count
$fn = $missed.Count
$fp = $fpList.Count
$recall = [Math]::Round(($tp / $totalExp) * 100, 1)
$missRate = [Math]::Round(($fn / $totalExp) * 100, 1)
$fpRate = if ($issueLines.Count -gt 0) { [Math]::Round(($fp / $issueLines.Count) * 100, 1) } else { 0 }

Write-Host "================================================================="
Write-Host " KIỂM THỬ CHÍNH TẢ WORD (.DOCX) THỰC TẾ — VBTM PRO 3.6"
Write-Host "================================================================="
Write-Host $headerLine
Write-Host "-----------------------------------------------------------------"
foreach ($ln in $issueLines) { Write-Host "  $ln" }
Write-Host "-----------------------------------------------------------------"
Write-Host "TỔNG LỖI ĐÃ CÀI (Ground Truth) : $totalExp (Đoạn văn + Bảng 4x3 + Header + Footer + Footnote)"
Write-Host "PHÁT HIỆN ĐÚNG (TP)            : $tp / $totalExp  -> Tỷ lệ phát hiện (Recall) : $recall%"
Write-Host "BỎ SÓT (FN)                    : $fn / $totalExp  -> Tỷ lệ bỏ sót (Miss Rate) : $missRate%"
Write-Host "CẢNH BÁO SAI (FP)              : $fp / $($issueLines.Count) -> Tỷ lệ cảnh báo sai (FPR) : $fpRate%"
Write-Host "BÁO NHẦM TÊN RIÊNG / MÃ HIỆU   : $($protectedHit.Count) / $($protectedClean.Count)"
if ($missed.Count -gt 0) { Write-Host "DANH SÁCH BỎ SÓT: $($missed -join ', ')" }
if ($fpList.Count -gt 0) { Write-Host "DANH SÁCH CẢNH BÁO SAI: $($fpList -join ' | ')" }
Write-Host "================================================================="

Remove-Item $testDocx -Force -ErrorAction SilentlyContinue
if ($fn -gt 0 -or $fp -gt 0 -or $protectedHit.Count -gt 0) { exit 1 }
exit 0
