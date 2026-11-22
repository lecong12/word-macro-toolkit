Attribute VB_Name = "Module8"
Sub ClearAutoCorrectEntries()
    Dim acEntry As AutoCorrectEntry
    Dim acList As AutoCorrectEntries
    Dim i As Integer
    
    ' L?y danh sách các m?c AutoCorrect
    Set acList = Application.AutoCorrect.Entries
    
    ' Xóa t?ng m?c trong danh sách
    For i = acList.count To 1 Step -1
        Set acEntry = acList(i)
        acEntry.Delete
    Next i
    
    MsgBox "T?t c? các m?c AutoCorrect dã du?c xóa.", vbInformation
End Sub


