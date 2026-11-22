Attribute VB_Name = "Module13"
Sub FormatFirstTwoWords()
    Dim Para As Paragraph
    Dim count As Integer
    Dim newText As String
    Dim paraRange As Range
    Dim rangeNewText As Range
    Dim rangeRemainingText As Range
    Dim paraText As String
    Dim paraCount As Integer
    
    ' Ð?t bi?n d?m b?t d?u t? 1
    count = 1
    
    ' L?p qua t?t c? các do?n van trong vùng ch?n
    For Each Para In Selection.Paragraphs
        ' Xác d?nh van b?n m?i c?n chèn vào d?u do?n
        Select Case count
            Case 1
                newText = "Th? nh?t, "
            Case 2
                newText = "Th? hai, "
            Case 3
                newText = "Th? ba, "
            Case 4
                newText = "Th? tu, "
            Case 5
                newText = "Th? nam, "
            Case 6
                newText = "Th? sáu, "
            Case 7
                newText = "Th? b?y, "
            Case 8
                newText = "Th? tám, "
            Case 9
                newText = "Th? chín, "
            Case Else
                Exit For ' D?ng l?i khi dã x? lý d?n "Th? chín"
        End Select
        
        ' L?y van b?n hi?n t?i c?a do?n
        Set paraRange = Para.Range
        paraText = paraRange.Text
        
        ' Chèn van b?n m?i vào d?u do?n
        paraRange.InsertBefore newText
        
        ' Xác d?nh ph?m vi c?a van b?n m?i chèn vào
        Set rangeNewText = Para.Range.Duplicate
        rangeNewText.Start = Para.Range.Start
        rangeNewText.End = Para.Range.Start + Len(newText)
        
        ' Ð?nh d?ng van b?n m?i chèn vào
        With rangeNewText.Font
            .Name = "Times New Roman" ' Ð?m b?o phông ch? h? tr? ti?ng Vi?t
            .Italic = True
            .Size = 14
        End With
        
        ' Ð?nh d?ng ph?n van b?n còn l?i trong do?n
        Set rangeRemainingText = Para.Range.Duplicate
        rangeRemainingText.Start = Para.Range.Start + Len(newText)
        rangeRemainingText.End = Para.Range.End
        
        With rangeRemainingText.Font
            .Name = "Times New Roman" ' Ð?m b?o phông ch? h? tr? ti?ng Vi?t
            .Italic = False
            .Size = 14
        End With
        
        ' Tang bi?n d?m lên
        count = count + 1
        
        ' N?u dã d?t d?n s? lu?ng t?i da, d?ng l?i
        If count > 9 Then Exit For
    Next Para
End Sub

