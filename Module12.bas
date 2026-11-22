Attribute VB_Name = "Module12"
Sub DeleteUnusedStyles()
    Dim sty As style
    Dim doc As Document
    Dim usedStyles As Collection
    Dim Para As Paragraph
    Dim i As Integer
    Dim found As Boolean
    Dim deletedCount As Integer
    
    Set doc = ActiveDocument
    Set usedStyles = New Collection
    deletedCount = 0
    
    ' Luu t?t c? các Style dang du?c s? d?ng vào Collection
    For Each Para In doc.Paragraphs
        On Error Resume Next
        usedStyles.Add Para.style, CStr(Para.style)
        Debug.Print "Added style: " & Para.style.NameLocal
        On Error GoTo 0
    Next Para
    
    ' Xóa các Style không du?c s? d?ng, ngo?i tr? các Style h? th?ng
    For i = doc.Styles.count To 1 Step -1
        Set sty = doc.Styles(i)
        If Not sty.BuiltIn Then
            found = False
            On Error Resume Next
            Dim styleName As String
            styleName = CStr(sty)
            Dim J As Integer
            For J = 1 To usedStyles.count
                If usedStyles(J) = styleName Then
                    found = True
                    Exit For
                End If
            Next J
            On Error GoTo 0
            If Not found Then
                Debug.Print "Deleting style: " & sty.NameLocal ' In ra tên c?a Style dang xóa
                sty.Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next i
    
    ' Thông báo s? lu?ng Style dã b? xóa
    MsgBox deletedCount & " styles have been deleted.", vbInformation
End Sub

