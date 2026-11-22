Attribute VB_Name = "Module21"
Sub ReplaceSquareWithCheckboxInTable()
    Dim doc As Document
    Dim tbl As Table
    Dim cell As cell
    Dim rng As Range
    Dim findText As String
    Dim checkbox As ContentControl
    
    ' Ký t? ô vuông nh? c?n thay th? (có th? là ký t? ? ho?c các ký t? ô vuông khác)
    findText = "?"  ' Thay ký t? ô vuông tr?ng n?u c?n
    
    Set doc = ActiveDocument
    
    ' L?p qua t?t c? các b?ng trong tài li?u
    For Each tbl In doc.Tables
        ' L?p qua t?ng ô trong b?ng
        For Each cell In tbl.Range.Cells
            Set rng = cell.Range
            rng.End = rng.End - 1 ' Lo?i b? ký t? k?t thúc dòng trong ô b?ng
            
            ' B?t d?u tìm ki?m ô vuông trong ô b?ng
            With rng.Find
                .Text = findText
                .Replacement.Text = ""
                .Forward = True
                .Wrap = wdFindStop
                
                ' L?p qua các ô vuông nh? trong ô b?ng và thay th? b?ng checkbox
                Do While .Execute(Replace:=wdReplaceNone) = True
                    ' Xóa n?i dung ô vuông hi?n t?i
                    rng.Text = ""
                    
                    ' Thêm Content Control Checkbox vào tài li?u
                    Set checkbox = doc.ContentControls.Add(wdContentControlCheckBox)
                    
                    ' Ð?t checkbox vào dúng v? trí trong ô b?ng
                    checkbox.Range.SetRange rng.Start, rng.Start
                    ' Ch?n checkbox sau khi thêm
                    checkbox.Range.Select
                    ' Di chuy?n con tr? d? có th? ti?p t?c thay th?
                    rng.Collapse wdCollapseEnd
                Loop
            End With
        Next cell
    Next tbl
    
    MsgBox "Ðã thay th? xong các ô vuông nh? trong b?ng b?ng checkbox.", vbInformation
End Sub

