Attribute VB_Name = "VBTM_Pro_ND30"
' =============================================================================
' VBTM Pro 2.2.4 - Module Chuẩn Hóa Thể Thức Văn Bản Hành Chính Cho Microsoft Word
' Quy chuẩn: Nghị định 30/2020/NĐ-CP & Thông tư 30/2023/TT-BCA (CAND)
' Cách dùng trong Word: Nhấn Alt + F11 -> File -> Import File -> Chọn tệp này
' =============================================================================

Sub ChuanHoaVanBan_NghiDinh30()
    On Error Resume Next
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' 1. Chuẩn hóa khổ giấy A4 (21cm x 29.7cm) và lề trang (Trái 3cm, Phải 1.5cm, Trên 2cm, Dưới 2cm)
    With doc.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2#)
        .BottomMargin = CentimetersToPoints(2#)
        .LeftMargin = CentimetersToPoints(3#)
        .RightMargin = CentimetersToPoints(1.5)
    End With
    
    ' 2. Chuẩn hóa Phông chữ Times New Roman & Cỡ chữ 14pt, giãn đoạn 6pt
    With doc.Content.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    
    With doc.Content.ParagraphFormat
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.15)
        .SpaceBefore = 0
        .SpaceAfter = 6
        .FirstLineIndent = CentimetersToPoints(1.27)
        .Alignment = wdAlignParagraphJustify
    End With
    
    ' 3. Xóa khoảng trắng kép và khoảng trắng thừa trước dấu câu
    Call XoaKhoangTrangThua("  ", " ")
    Call XoaKhoangTrangThua(" ,", ",")
    Call XoaKhoangTrangThua(" .", ".")
    Call XoaKhoangTrangThua(" ;", ";")
    Call XoaKhoangTrangThua(" :", ":")
    Call XoaKhoangTrangThua(" !", "!")
    Call XoaKhoangTrangThua(" ?", "?")
    
    MsgBox "VBTM Pro 2.2.4: Da chuan hoa le trang A4 (30-15-20-20mm), font Times New Roman 14pt va dau cau theo Nghi dinh 30/2020/ND-CP!", vbInformation, "VBTM Pro 2.2.4"
End Sub

Private Sub XoaKhoangTrangThua(findStr As String, replaceStr As String)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findStr
        .Replacement.Text = replaceStr
        .Forward = True
        .Wrap = wdFindContinue
        .Execute Replace:=wdReplaceAll
    End With
End Sub
