Attribute VB_Name = "Module4"
Sub ExportAutoCorrectUTF8()
    Dim filePath As String
    Dim acEntry As Object
    Dim defaultExportPath As String
    Dim stream As Object
    
    On Error GoTo ErrorHandler
    
    ' Ð?t du?ng d?n m?c d?nh cho file xu?t
    defaultExportPath = "C:\Users\" & Environ("USERNAME") & "\Documents\AutoCorrectList_UTF8.txt"
    
    ' T?o d?i tu?ng ADODB.Stream d? ghi file v?i mã hóa UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Ch? d? Text
    stream.Charset = "UTF-8" ' Mã hóa UTF-8
    stream.Open
    
    ' Ki?m tra xem có m?c AutoCorrect nào không
    If Application.AutoCorrect.Entries.count = 0 Then
        MsgBox "Không có m?c AutoCorrect nào d? xu?t."
        stream.Close
        Exit Sub
    End If
    
    ' Ghi t?ng m?c AutoCorrect vào file
    For Each acEntry In Application.AutoCorrect.Entries
        On Error Resume Next  ' B?t l?i n?u có m?c nào không h?p l?
        stream.WriteText acEntry.Name & vbTab & acEntry.Value & vbCrLf
        If Err.Number <> 0 Then
            MsgBox "L?i khi x? lý m?c: " & acEntry.Name & " - " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler  ' Quay l?i ch? d? b?t l?i chính
    Next acEntry
    
    ' Luu và dóng file
    stream.SaveToFile defaultExportPath, 2 ' 2 nghia là t?o m?i ho?c ghi dè
    stream.Close
    MsgBox "Danh sách AutoCorrect dã du?c xu?t ra file: " & defaultExportPath

    Exit Sub

ErrorHandler:
    MsgBox "Ðã x?y ra l?i: " & Err.Description, vbCritical
    If Not stream Is Nothing Then stream.Close
End Sub

