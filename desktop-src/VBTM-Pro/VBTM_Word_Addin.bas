Attribute VB_Name = "VBTM_Pro_HD05_ND30"
' =============================================================================
' VBTM Pro 2.5.0 - Add-in Macro Tự Động Hóa Thể Thức Văn Bản Word
' Quy chuẩn:
'   - Khối Đảng: Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng
'   - Khối Hành chính: Nghị định số 30/2020/NĐ-CP của Chính phủ
' =============================================================================

Sub ChuanHoaVanBan_TuDong()
    On Error Resume Next
    Dim docText As String
    docText = UCase(ActiveDocument.Content.Text)
    If InStr(docText, "ĐẢNG CỘNG SẢN") > 0 Or InStr(docText, "ĐẢNG ỦY") > 0 Or InStr(docText, "CHI BỘ") > 0 Or InStr(docText, "-BC/ĐU") > 0 Or InStr(docText, "-NQ/ĐU") > 0 Or InStr(docText, "-CV/ĐU") > 0 Then
        ChuanHoaVanBan_HuongDan05
    Else
        ChuanHoaVanBan_NghiDinh30
    End If
End Sub

Sub ChuanHoaVanBan_HuongDan05()
    On Error Resume Next
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1.5)
        .DifferentFirstPageHeaderFooter = True
    End With

    ' Đặt font Times New Roman toàn văn bản (không ghi đè cỡ chữ bảng Quốc hiệu / Nơi nhận)
    ActiveDocument.Content.Font.Name = "Times New Roman"

    ' Đánh số trang tự động từ trang 2 ở Header (cỡ 13pt đứng, căn giữa)
    Dim hdr As HeaderFooter
    Set hdr = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
    hdr.Range.Text = ""
    hdr.PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False
    hdr.Range.Font.Name = "Times New Roman"
    hdr.Range.Font.Size = 13

    ' Định dạng các đoạn thân bài bên ngoài bảng theo Hướng dẫn 05-HD/VPTW (14pt, Justify, thụt 1.27cm, giãn đoạn 6pt)
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Tables.Count = 0 Then
            Dim t As String
            t = Trim(Replace(Replace(p.Range.Text, vbCr, ""), Chr(7), ""))
            If Len(t) > 0 Then
                p.Range.Font.Name = "Times New Roman"
                If t = "BÁO CÁO" Or t = "NGHỊ QUYẾT" Or t = "QUYẾT ĐỊNH" Or t = "KẾ HOẠCH" Or t = "CHƯƠNG TRÌNH" Or t = "TỜ TRÌNH" Or t = "THÔNG BÁO" Then
                    p.Range.Font.Size = 15
                    p.Range.Font.Bold = True
                    p.Alignment = wdAlignParagraphCenter
                    p.FirstLineIndent = 0
                Else
                    p.Range.Font.Size = 14
                    p.Format.SpaceAfter = 6
                    p.Format.LineSpacingRule = wdLineSpaceMultiple
                    p.Format.LineSpacing = LinesToPoints(1.15)
                    If p.Alignment <> wdAlignParagraphCenter Then
                        p.Alignment = wdAlignParagraphJustify
                        p.FirstLineIndent = CentimetersToPoints(1.27)
                    End If
                End If
            End If
        End If
    Next p

    ' Định dạng Bảng Tiêu ngữ (Bảng 1) & Bảng Nơi nhận / Chữ ký (Bảng cuối) theo Hướng dẫn 05-HD/VPTW
    If ActiveDocument.Tables.Count >= 1 Then
        Dim tblTop As Table
        Set tblTop = ActiveDocument.Tables(1)
        If tblTop.Columns.Count >= 2 Then
            tblTop.Cell(1, 1).Range.Font.Name = "Times New Roman"
            tblTop.Cell(1, 1).Range.Font.Size = 14
            tblTop.Cell(1, 2).Range.Font.Name = "Times New Roman"
            tblTop.Cell(1, 2).Range.Font.Size = 15
            Dim pTop As Paragraph
            For Each pTop In tblTop.Cell(1, 2).Range.Paragraphs
                If InStr(LCase(pTop.Range.Text), "ngày") > 0 Then
                    pTop.Range.Font.Size = 14
                    pTop.Range.Font.Italic = True
                    pTop.Range.Font.Bold = False
                End If
            Next pTop
        End If
    End If

    If ActiveDocument.Tables.Count >= 2 Then
        Dim tblBot As Table
        Set tblBot = ActiveDocument.Tables(ActiveDocument.Tables.Count)
        If tblBot.Columns.Count >= 2 Then
            Dim pRec As Paragraph
            For Each pRec In tblBot.Cell(1, 1).Range.Paragraphs
                pRec.Range.Font.Name = "Times New Roman"
                pRec.Range.Font.Italic = False
                If InStr(LCase(pRec.Range.Text), "nơi nhận") > 0 Then
                    pRec.Range.Font.Size = 14
                    pRec.Range.Font.Bold = False
                    pRec.Range.Font.Underline = wdUnderlineSingle
                Else
                    pRec.Range.Font.Size = 12
                End If
            Next pRec
            tblBot.Cell(1, 2).Range.Font.Name = "Times New Roman"
            tblBot.Cell(1, 2).Range.Font.Size = 14
            tblBot.Cell(1, 2).Range.Font.Bold = True
        End If
    End If

    MsgBox "Đã tự động chuẩn hóa văn bản Word đúng chuẩn Hướng dẫn số 05-HD/VPTW của Văn phòng Trung ương Đảng!", vbInformation, "VBTM Pro 2.5.0"
End Sub

Sub ChuanHoaVanBan_NghiDinh30()
    On Error Resume Next
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1.5)
        .DifferentFirstPageHeaderFooter = True
    End With
    With ActiveDocument.Content.Font
        .Name = "Times New Roman"
    End With
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Tables.Count = 0 Then
            p.Range.Font.Size = 14
            p.Format.SpaceAfter = 6
            p.Format.LineSpacingRule = wdLineSpaceMultiple
            p.Format.LineSpacing = LinesToPoints(1.15)
        End If
    Next p
    MsgBox "Đã chuẩn hóa văn bản Word theo Nghị định số 30/2020/NĐ-CP!", vbInformation, "VBTM Pro 2.5.0"
End Sub
