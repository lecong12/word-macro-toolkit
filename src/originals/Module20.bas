Attribute VB_Name = "Module20"
Sub MergeWordFiles()
    Dim mainDoc As Document
    Dim fileDialog As fileDialog
    Dim i As Long
    Dim tempDoc As Document
    
    ' T?o tài li?u m?i d? ghép
    Set mainDoc = Documents.Add
    
    ' H?p tho?i ch?n file
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .AllowMultiSelect = True
        .Title = "Ch?n các file Word c?n ghép"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx; *.docm; *.doc"
        
        ' Ki?m tra n?u ngu?i dùng ch?n file
        If .Show = -1 Then
            For i = 1 To .SelectedItems.count
                Set tempDoc = Documents.Open(.SelectedItems(i))
                tempDoc.Content.Copy
                mainDoc.Range(mainDoc.Content.End - 1).Paste
                tempDoc.Close SaveChanges:=False
                
                ' Thêm ng?t trang (tr? file cu?i)
                If i < .SelectedItems.count Then
                    mainDoc.Range(mainDoc.Content.End - 1).InsertBreak wdPageBreak
                End If
            Next i
        Else
            MsgBox "Không có file nào du?c ch?n.", vbExclamation
        End If
    End With
    
    MsgBox "Ðã ghép xong t?t c? các file!", vbInformation
End Sub

