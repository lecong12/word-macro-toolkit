Attribute VB_Name = "Module18"
Sub highlightPartialMatches()
    Dim i, J As Long
    Dim xRngFind, xRng As Range
    Dim startPos As Long, matchPos As Long
    Options.DefaultHighlightColorIndex = wdYellow
    Application.ScreenUpdating = False
    With ActiveDocument
        For i = 1 To .Paragraphs.count - 1
            Set xRngFind = .Paragraphs(i).Range
            If xRngFind.HighlightColorIndex <> wdYellow Then
                For J = i + 1 To .Paragraphs.count
                    Set xRng = .Paragraphs(J).Range
                    ' Tìm v? trí ph?n tuong d?ng gi?a hai do?n van
                    matchPos = InStr(1, xRng.Text, xRngFind.Text, vbTextCompare)
                    If matchPos > 0 Then
                        ' Tô màu cho ph?n tuong d?ng trong do?n J
                        With xRng.Duplicate
                            .Start = xRng.Start + matchPos - 1
                            .End = .Start + Len(xRngFind.Text)
                            .HighlightColorIndex = wdYellow
                        End With
                        
                        ' Tô màu cho ph?n tuong d?ng trong do?n I
                        matchPos = InStr(1, xRngFind.Text, xRng.Text, vbTextCompare)
                        With xRngFind.Duplicate
                            .Start = xRngFind.Start + matchPos - 1
                            .End = .Start + Len(xRng.Text)
                            .HighlightColorIndex = wdBrightGreen
                        End With
                    End If
                Next J
            End If
        Next i
    End With
    Application.ScreenUpdating = True
End Sub

