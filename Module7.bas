Attribute VB_Name = "Module7"
Sub OpenContainingFolder()
    Dim folderPath As String
    
    ' L?y du?ng d?n c?a tài li?u hi?n t?i
    folderPath = ThisDocument.Path
    
    ' M? thu m?c ch?a tài li?u trong Windows Explorer
    If folderPath <> "" Then
        Shell "explorer.exe /e," & folderPath, vbNormalFocus
    Else
        MsgBox "Tài li?u chua du?c luu. Vui lòng luu tài li?u tru?c.", vbExclamation
    End If
End Sub

