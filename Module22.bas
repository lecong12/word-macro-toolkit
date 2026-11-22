Attribute VB_Name = "Module22"
Sub ConvertDocToDocx()
    Dim objWord As Object
    Dim objDoc As Object
    Dim fileName As String
    Dim folderPath As String
    Dim file As String
    Dim dialog As Object
    Dim foundDoc As Boolean

    ' M? h?p tho?i d? ch?n thu m?c
    Set dialog = CreateObject("Shell.Application").BrowseForFolder(0, "Ch?n thu m?c ch?a các t?p .doc", 0, 0)
    
    If dialog Is Nothing Then
        MsgBox "Không có thu m?c nào du?c ch?n!"
        Exit Sub
    End If

    folderPath = dialog.Items().Item().Path & "\"

    ' Ð?m b?o r?ng folderPath có d?u \ cu?i
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' Kh?i t?o d?i tu?ng Word
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False

    ' Duy?t t?t c? các t?p trong thu m?c
    foundDoc = False
    file = Dir(folderPath & "*.*")  ' Duy?t t?t c? các t?p trong thu m?c

    ' L?p qua t?t c? các t?p trong thu m?c
    Do While file <> ""
        ' Ki?m tra n?u là t?p .doc
        If LCase(Right(file, 4)) = ".doc" Then
            foundDoc = True
            ' M? t?p .doc
            On Error Resume Next ' X? lý l?i khi không m? du?c t?p
            Set objDoc = objWord.Documents.Open(folderPath & file)
            On Error GoTo 0 ' T?t ch? d? b?t l?i

            ' Ki?m tra n?u t?p m? thành công
            If Not objDoc Is Nothing Then
                ' Luu t?p du?i d?ng .docx
                fileName = folderPath & Left(file, Len(file) - 4) & ".docx"  ' Thay ph?n m? r?ng thành .docx
                objDoc.SaveAs fileName, 16  ' 16 là mã d?nh d?ng .docx

                ' Ðóng t?p .doc
                objDoc.Close
            End If
        End If

        ' Chuy?n sang t?p ti?p theo
        file = Dir
    Loop

    If Not foundDoc Then
        MsgBox "Thu m?c không có t?p .doc nào!"
    Else
        MsgBox "Chuy?n d?i hoàn t?t!"
    End If

    ' Thoát kh?i Word
    objWord.Quit
End Sub

