Attribute VB_Name = "Module5"
Sub ImportAutoCorrectUTF8()
    Dim filePath As String
    Dim fileContent As String
    Dim entryParts() As String
    Dim fd As fileDialog
    Dim stream As Object
    Dim lines() As String
    Dim line As Variant
    
    On Error GoTo ErrorHandler
    
    ' T?o d?i tu?ng FileDialog d? ch?n file
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    
    ' C?u hình h?p tho?i
    fd.Title = "Ch?n file AutoCorrect d? nh?p"
    fd.Filters.Clear
    fd.Filters.Add "Text Files", "*.txt"
    
    ' Hi?n th? h?p tho?i và ki?m tra n?u ngu?i dùng ch?n file
    If fd.Show = -1 Then
        filePath = fd.SelectedItems(1)
    Else
        MsgBox "Không có file nào du?c ch?n.", vbExclamation
        Exit Sub
    End If
    
    ' S? d?ng ADODB.Stream d? d?c file UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Ð?t ki?u d? li?u là van b?n (text)
    stream.Mode = 3 ' Ð?t ch? d? là read/write
    stream.Charset = "UTF-8" ' Ð?t ki?u mã hóa là UTF-8
    stream.Open
    stream.LoadFromFile filePath
    
    ' Ð?c toàn b? n?i dung file và chuy?n thành chu?i Unicode
    fileContent = stream.ReadText
    stream.Close
    Set stream = Nothing
    
    ' Tách d? li?u thành các dòng
    lines = Split(fileContent, vbCrLf)
    
    ' Thêm các m?c AutoCorrect t? d? li?u
    For Each line In lines
        entryParts = Split(line, vbTab)
        If UBound(entryParts) = 1 Then
            Application.AutoCorrect.Entries.Add Name:=entryParts(0), Value:=entryParts(1)
        End If
    Next line
    
    MsgBox "Nh?p danh sách AutoCorrect thành công t? file: " & filePath

    Exit Sub

ErrorHandler:
    MsgBox "Ðã x?y ra l?i: " & Err.Description, vbCritical
    If Not stream Is Nothing Then stream.Close
End Sub
