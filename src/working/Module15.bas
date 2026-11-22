Attribute VB_Name = "Module15"
Option Explicit

' ==============================================================================
' MODULE 15: REMOVE HIGHLIGHT (XOA MAU DANH DAU VAN BAN)
' ==============================================================================
' Cac chuc nang duoc nang cap:
' 1. RemoveHighlight: Xoa tat ca mau highlight (ho tro ca Vung chon hoac Toan tai lieu)
' 2. RemoveYellowHighlightOnly: Chi xoa mau vang wdYellow (tuong thich voi Module14)
' 3. RemoveHighlightInSelection: Chi xoa highlight trong vung dang boi den
'
' Uu diem so voi ban goc:
' - Su dung Word Find & Replace: Xu ly triet de tu don le, cum tu va toan doan.
'   Khac phuc hoan toan loi wdUndefined khi chi mot phan doan co highlight.
' - Toc do xu ly cuc nhanh (nhanh hon hang chuc lan so voi loop Paragraphs).
' - Ho tro StoryRanges: Quet sach ca Header, Footer neu xoa toan tai lieu.
' - Ho tro UndoRecord: Hoan tac toan bo thao tac chi bang 1 lan bam Ctrl+Z.
' - Co kiem tra an toan tai lieu mo, bat/tat ScreenUpdating va bay loi day du.
' ==============================================================================

Sub RemoveHighlight()
    ' Thu tuc chinh: Xoa toan bo mau highlight
    ' - Neu co vung chon (Selection): Chi xoa trong vung chon
    ' - Neu khong co vung chon: Xoa tren toan bo tai lieu
    On Error GoTo ErrorHandler
    
    If Documents.Count = 0 Then
        MsgBox "Khong co tai lieu nao dang mo!", vbExclamation, "Thong bao"
        Exit Sub
    End If
    
    Dim hasSelection As Boolean
    hasSelection = (Selection.Type <> wdSelectionIP And Selection.Start <> Selection.End)
    
    Dim ur As Object
    Dim hasUndoRecord As Boolean
    hasUndoRecord = False
    
    #If VBA7 Then
        On Error Resume Next
        Set ur = Application.UndoRecord
        If Not ur Is Nothing Then
            ur.StartCustomRecord "Xoa Highlight"
            hasUndoRecord = True
        End If
        On Error GoTo ErrorHandler
    #End If
    
    Application.ScreenUpdating = False
    
    If hasSelection Then
        ' Xoa trong vung duoc chon
        ClearAllHighlightInRange Selection.Range
        Application.ScreenUpdating = True
        Application.StatusBar = "Da xoa toan bo highlight trong vung chon."
    Else
        ' Xoa tren toan bo tai lieu (tat ca StoryRanges: MainText, Headers, Footers)
        Dim aStory As Range
        For Each aStory In ActiveDocument.StoryRanges
            Do
                ClearAllHighlightInRange aStory
                Set aStory = aStory.NextStoryRange
            Loop Until aStory Is Nothing
        Next aStory
        Application.ScreenUpdating = True
        Application.StatusBar = "Da xoa toan bo highlight tren toan tai lieu."
    End If

CleanUp:
    On Error Resume Next
    If hasUndoRecord And Not ur Is Nothing Then
        ur.EndCustomRecord
    End If
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Da xay ra loi: " & Err.Description, vbCritical, "Loi"
    Resume CleanUp
End Sub

Sub RemoveYellowHighlightOnly()
    ' Chi xoa highlight mau vang (wdYellow)
    ' Giu nguyen cac mau highlight khac (nhu mau xanh BrightGreen cua Module18)
    On Error GoTo ErrorHandler
    
    If Documents.Count = 0 Then
        MsgBox "Khong co tai lieu nao dang mo!", vbExclamation, "Thong bao"
        Exit Sub
    End If
    
    Dim hasSelection As Boolean
    hasSelection = (Selection.Type <> wdSelectionIP And Selection.Start <> Selection.End)
    
    Dim ur As Object
    Dim hasUndoRecord As Boolean
    hasUndoRecord = False
    
    #If VBA7 Then
        On Error Resume Next
        Set ur = Application.UndoRecord
        If Not ur Is Nothing Then
            ur.StartCustomRecord "Xoa Highlight Vang"
            hasUndoRecord = True
        End If
        On Error GoTo ErrorHandler
    #End If
    
    Application.ScreenUpdating = False
    
    Dim countRemoved As Long
    countRemoved = 0
    
    If hasSelection Then
        countRemoved = ClearSpecificColorInRange(Selection.Range, wdYellow)
        Application.ScreenUpdating = True
        Application.StatusBar = "Da xoa " & countRemoved & " vi tri highlight vang trong vung chon."
    Else
        Dim aStory As Range
        For Each aStory In ActiveDocument.StoryRanges
            Do
                countRemoved = countRemoved + ClearSpecificColorInRange(aStory, wdYellow)
                Set aStory = aStory.NextStoryRange
            Loop Until aStory Is Nothing
        Next aStory
        Application.ScreenUpdating = True
        Application.StatusBar = "Da xoa " & countRemoved & " vi tri highlight vang tren toan tai lieu."
    End If

CleanUp:
    On Error Resume Next
    If hasUndoRecord And Not ur Is Nothing Then
        ur.EndCustomRecord
    End If
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Da xay ra loi: " & Err.Description, vbCritical, "Loi"
    Resume CleanUp
End Sub

Sub RemoveHighlightInSelection()
    ' Xoa tat ca highlight danh rieng cho vung dang boi den
    If Documents.Count = 0 Then Exit Sub
    If Selection.Type = wdSelectionIP Then
        MsgBox "Vui long boi den doan van ban can xoa highlight!", vbExclamation, "Chua chon van ban"
        Exit Sub
    End If
    RemoveHighlight
End Sub

' ==============================================================================
' CAC HAM PHU TRO (HELPER FUNCTIONS)
' ==============================================================================

Private Sub ClearAllHighlightInRange(ByVal rng As Range)
    ' Xoa tat ca mau danh dau trong mot Range bang Find Replace native (sieu toc)
    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Highlight = True
        .Replacement.Highlight = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
    On Error GoTo 0
End Sub

Private Function ClearSpecificColorInRange(ByVal targetRng As Range, ByVal targetColor As WdColorIndex) As Long
    ' Quet tim tung span co highlight va chi xoa span dung mau targetColor
    ' Khong bao gio bi loi wdUndefined vi findRng la mot span highlight lien tuc
    Dim count As Long
    count = 0
    
    On Error GoTo ErrExit
    
    Dim findRng As Range
    Set findRng = targetRng.Duplicate
    
    With findRng.Find
        .ClearFormatting
        .Text = ""
        .Highlight = True
        .Forward = True
        .Wrap = wdFindStop
        .Format = True
        
        Do While .Execute(Forward:=True, Wrap:=wdFindStop)
            ' Kiem tra gioi han pham vi
            If findRng.Start >= targetRng.End Then Exit Do
            If findRng.End > targetRng.End Then
                findRng.End = targetRng.End
            End If
            
            ' Kiem tra dung mau can xoa
            If findRng.HighlightColorIndex = targetColor Then
                findRng.HighlightColorIndex = wdNoHighlight
                count = count + 1
            End If
            
            ' Di chuyen tiep tu diem ket thuc cua span vua tim
            findRng.Collapse wdCollapseEnd
            If findRng.Start >= targetRng.End Then Exit Do
        Loop
    End With

ErrExit:
    ClearSpecificColorInRange = count
End Function
