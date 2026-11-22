Attribute VB_Name = "Module15"
Sub RemoveHighlight()
    Dim Para As Paragraph
    
    ' Duy?t qua t?t c? các do?n van b?n trong tài li?u
    For Each Para In ActiveDocument.Paragraphs
        ' N?u do?n van b?n có màu dánh d?u là màu vàng
        If Para.Range.HighlightColorIndex = wdYellow Then
            ' Xóa màu dánh d?u
            Para.Range.HighlightColorIndex = wdNoHighlight
        End If
    Next Para
End Sub

