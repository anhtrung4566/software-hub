Attribute VB_Name = "VBTM_Pro_ND30"
' =============================================================================
' VBTM Pro 2.2.4 - Add-in Macro Chuan Hoa Van Ban Dang Uy & UBND Phuong
' Quy chuan: Nghi dinh 30/2020/ND-CP & Huong dan 36, 05-HD/VPTW
' =============================================================================

Sub ChuanHoaVanBan_NghiDinh30()
    On Error Resume Next
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1.5)
    End With
    With ActiveDocument.Content.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    With ActiveDocument.Content.ParagraphFormat
        .SpaceBefore = 6
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceMultiple
        .LineSpacing = LinesToPoints(1.2)
    End With
    MsgBox "Da chuan hoa the thuc van ban theo Nghi dinh 30/2020/ND-CP & Quy chuan Dang uy - UBND Phuong!", vbInformation, "VBTM Pro 2.2.4"
End Sub
