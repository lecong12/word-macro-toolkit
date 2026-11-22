Attribute VB_Name = "Module19"
Sub SplitDocumentByPageBreak()
    Dim doc As Document
    Dim newDoc As Document
    Dim splitRange As Range
    Dim splitCount As Integer
    Dim pageBreakPos As Long
    Dim savePath As String
    Dim nextPageBreak As Range
    
    ' Gán tài li?u hi?n t?i vào bi?n doc
    Set doc = ActiveDocument
    
    ' L?y du?ng d?n c?a file g?c
    savePath = doc.Path & Application.PathSeparator
    
    If savePath = "" Then
        MsgBox "Hãy luu file g?c tru?c khi ch?y macro!", vbExclamation
        Exit Sub
    End If
    
    splitCount = 0
    pageBreakPos = 0
    
    ' L?p qua t?ng ph?n gi?a các ng?t trang
    Do
        splitCount = splitCount + 1
        
        ' Tìm ng?t trang ti?p theo
        Set nextPageBreak = doc.Range(pageBreakPos, doc.Content.End).Duplicate
        If Not nextPageBreak.Find.Execute(findText:="^m") Then
            ' X? lý do?n cu?i cùng sau ng?t trang cu?i cùng
            Set splitRange = doc.Range(pageBreakPos, doc.Content.End)
            Set newDoc = Documents.Add
            newDoc.Content.FormattedText = splitRange.FormattedText
            newDoc.SaveAs2 fileName:=savePath & "Part" & splitCount & ".docx"
            newDoc.Close SaveChanges:=False
            Exit Do  ' D?ng vòng l?p khi không còn ng?t trang
        End If
        
        ' Chia do?n van b?n t? v? trí hi?n t?i d?n ng?t trang ti?p theo
        Set splitRange = doc.Range(pageBreakPos, nextPageBreak.Start)
        
        ' T?o tài li?u m?i
        Set newDoc = Documents.Add
        newDoc.Content.FormattedText = splitRange.FormattedText  ' Gi? nguyên d?nh d?ng
        
        ' Luu file m?i vào thu m?c cùng v?i file g?c
        newDoc.SaveAs2 fileName:=savePath & "Part" & splitCount & ".docx"
        newDoc.Close SaveChanges:=False
        
        ' C?p nh?t v? trí d? tìm ng?t trang ti?p theo
        pageBreakPos = nextPageBreak.End + 1
        
    Loop
    
    MsgBox "Ðã tách tài li?u thành công! Các file du?c luu t?i: " & savePath, vbInformation
End Sub

