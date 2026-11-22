Attribute VB_Name = "Module2"
Sub InsertPageNumbers()
    ' Chèn s? trang vào d?u trang v?i ki?u Plain Number 2
    With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
        .PageNumbers.Add PageNumberAlignment:=wdAlignParagraphCenter, FirstPage:=True
        .Range.Fields.Add Range:=.Range, Type:=wdFieldPage
    End With
    
    ' Ð?nh d?ng ki?u Plain Number 2 (canh gi?a, s? n?m trên cùng)
    With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Name = "Times New Roman"
        .Font.Size = 12
    End With
End Sub

