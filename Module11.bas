Attribute VB_Name = "Module11"
Sub DeleteTextWithNormalStyle()
    Dim Para As Paragraph
    Dim doc As Document
    Set doc = ActiveDocument

    ' L?p qua t?t c? các do?n van b?n trong tài li?u
    For Each Para In doc.Paragraphs
        ' Ki?m tra xem style c?a do?n van có ph?i là Normal hay không (c? ti?ng Anh và ti?ng Vi?t)
        If Para.style = doc.Styles(wdStyleNormal) Then
            Para.Range.Delete
        End If
    Next Para
End Sub

