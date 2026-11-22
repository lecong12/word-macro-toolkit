Attribute VB_Name = "Module17"
Sub XoaDongTrong()
    Dim Para As Paragraph
    Dim i As Integer
    
    ' L?p qua t?t c? các do?n van trong tài li?u
    For i = ActiveDocument.Paragraphs.count To 1 Step -1
        Set Para = ActiveDocument.Paragraphs(i)
        
        ' Ki?m tra n?u do?n van ch? ch?a ký t? xu?ng dòng
        If Len(Para.Range.Text) = 1 Then
            Para.Range.Delete
        End If
    Next i
End Sub

