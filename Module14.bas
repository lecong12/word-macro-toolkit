Attribute VB_Name = "Module14"
Sub HighlightParagraphsWithStyle()
    Dim styleName As String
    Dim Para As Paragraph
    
    ' Nh?p tên style b?n mu?n tìm
    styleName = InputBox("Nh?p tên style b?n mu?n tìm:", "Tìm do?n van b?n theo style")
    
    ' Duy?t qua t?t c? các do?n van b?n trong tài li?u
    For Each Para In ActiveDocument.Paragraphs
        ' N?u do?n van b?n có style kh?p v?i tên style dã nh?p
        If Para.style = styleName Then
            ' Ðánh d?u do?n van b?n
            Para.Range.HighlightColorIndex = wdYellow
        End If
    Next Para
End Sub

