Attribute VB_Name = "Module9"
Sub QuoteToTextSelct()
    ' Ki?m tra n?u có van b?n du?c ch?n
    If Selection.Type = wdSelectionNormal Then
        ' Thêm d?u ngo?c kép m? “
        Selection.InsertBefore ChrW(&H201C)
        ' Thêm d?u ngo?c kép dóng ”
        Selection.InsertAfter ChrW(&H201D)
        ' In nghiêng do?n van b?n du?c ch?n (bao g?m c? d?u ngo?c kép)
        Selection.Font.Italic = False
    Else
        MsgBox "Vui lòng ch?n van b?n tru?c khi ch?y macro.", vbExclamation
    End If
End Sub
