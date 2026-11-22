Attribute VB_Name = "Module16"
Sub ConvertFootnotesToText()
    Dim footnote As footnote
    Dim rng As Range

    For Each footnote In ActiveDocument.Footnotes
        Set rng = footnote.Reference
        rng.InsertAfter " (" & footnote.Range.Text & ")"
        footnote.Delete
    Next footnote
End Sub

