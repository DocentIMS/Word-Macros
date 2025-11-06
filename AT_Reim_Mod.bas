Attribute VB_Name = "AT_Reim_Mod"
Option Explicit

Function IsValidReimbursement(Optional Doc As Document) As Boolean
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Dim CC As ContentControl, CN As CustomXMLNode
    Dim GroupsColl As New Collection, i As Long
    Dim oCustomPart As Office.CustomXMLPart, ThisGroup As String
    IsValidReimbursement = True
    IsValidReimbursement = IsValidReimbursement And ValidateCCTb("Description", Doc)
    IsValidReimbursement = IsValidReimbursement And ValidateCCTb("AmountRequested", Doc)
    IsValidReimbursement = IsValidReimbursement And ValidateCCTb("PayeeName", Doc)
    IsValidReimbursement = IsValidReimbursement And ValidateCCTb("PayeeAddress1", Doc)
    IsValidReimbursement = IsValidReimbursement And ValidateCCTb("Approver_1", Doc)
    'OtherExpenses
'    If Len(GetContentControl("Description", Doc)) = 0 Then Reimbursement = False: GoTo inv
'    If Len(GetContentControl("PayeeName", Doc)) = 0 Then Reimbursement = False: GoTo inv
'    If Len(GetContentControl("PayeeAddress1", Doc)) = 0 Then Reimbursement = False: GoTo inv
'    If Len(GetContentControl("Approver_1", Doc)) = 0 Then Reimbursement = False: GoTo inv
    On Error Resume Next
    For Each CC In Doc.ContentControls
        If CC.Type = wdContentControlCheckBox Then
            GroupsColl.Add CStr(Split(CC.Tag, "_")(0)), CStr(Split(CC.Tag, "_")(0))
            
'            If Not CCCKGroupSelected(CStr(Split(CC.Tag, "_")(0)), Doc) Then Reimbursement = False: GoTo inv
        End If
    Next
    For i = 1 To GroupsColl.Count
        IsValidReimbursement = IsValidReimbursement And CCCKGroupSelected(GroupsColl(i), Doc)
    Next
End Function
Private Function ValidateCCTb(CCName As String, Doc As Document) As Boolean
    ValidateCCTb = Len(GetContentControl(CCName, Doc)) > 0
    ColorCC FindCCs(CCName, Doc)(1), IIf(ValidateCCTb, wdNoHighlight, wdRed)
'    On Error Resume Next
'    With FindCCs(CCName, Doc)(1)
'        .Range.HighlightColorIndex = IIf(ValidateCCTb, wdNoHighlight, wdRed)
'        If Err.Number Then
'            If .Range.Information(wdWithInTable) Then
'                Dim i As Long
'                With .Range.Cells(1).Range
'                    For i = 1 To .Paragraphs.Count
'                        .Paragraphs(i).Range.HighlightColorIndex = IIf(ValidateCCTb, wdNoHighlight, wdRed)
'                    Next
'                End With
'            End If
'        End If
'    End With
End Function
Private Function CCCKGroupSelected(GroupName As String, Optional Doc As Document) As Boolean
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Dim CC As ContentControl, Rng As Range, ErrFlag As Boolean
    For Each CC In Doc.ContentControls
        If CC.Type = wdContentControlCheckBox Then
            If CC.Tag Like GroupName & "_*" Then
                If CC.Checked Then
                    CCCKGroupSelected = CC.Tag <> "BudgetCat_5"
                    If Not CCCKGroupSelected Then CCCKGroupSelected = ValidateCCTb("OtherExpenses", Doc)
                    ErrFlag = Not CCCKGroupSelected And CC.Tag <> "BudgetCat_5"
                    If CCCKGroupSelected Then GoTo ex
                End If
            End If
        End If
    Next
'    Exit Function
ex:
'    If CCCKGroupSelected And GroupName = "BudgetCat" Then FindCCs("AmountRequested", Doc)(1).Range.HighlightColorIndex = wdNoHighlight
    For Each CC In Doc.ContentControls
        If CC.Type = wdContentControlCheckBox Then
            If CC.Tag Like GroupName & "_*" Then
                If CC.Tag = "BudgetCat_5" Then ColorCC FindCCs("OtherExpenses", Doc)(1), IIf(CCCKGroupSelected, wdNoHighlight, wdRed)
                ColorCC CC, IIf(ErrFlag, wdRed, wdNoHighlight)
'                If CC.Range.Information(wdWithInTable) Then
'                    Set Rng = CC.Range.Cells(1).Range.Paragraphs(1).Range
'                    Rng.MoveEnd 1, -1
'                Else
'                    Set Rng = CC.Range.Paragraphs(1).Range
'                End If
'                Rng.MoveStart 1, 3
'                Rng.HighlightColorIndex = IIf(ErrFlag, wdRed, wdNoHighlight)
            End If
        End If
    Next
End Function
Private Sub ColorCC(CC As ContentControl, clr As WdColorIndex)
    Dim Rng As Range, i As Long
    On Error GoTo ex
    With CC
        If .Type = wdContentControlCheckBox Then
            If .Range.Information(wdWithInTable) Then
                Set Rng = .Range.Cells(1).Range.Paragraphs(1).Range
                Rng.MoveEnd 1, -1
            Else
                Set Rng = .Range.Paragraphs(1).Range
            End If
            Rng.MoveStart 1, 3
        ElseIf .Type = wdContentControlComboBox Then
            If .Range.Information(wdWithInTable) Then
                Set Rng = .Range.Cells(1).Range
                Rng.MoveEnd 1, -1
            End If
        Else
            Set Rng = .Range           '.HighlightColorIndex = IIf(ValidateCCTb, wdNoHighlight, wdRed)
            Rng.MoveStart 1, -1
            Rng.MoveEnd 1, 1
        End If
    End With
    On Error Resume Next
    For i = 1 To Rng.Paragraphs.Count
        Rng.Paragraphs(i).Range.HighlightColorIndex = clr 'IIf(ValidateCCTb, wdNoHighlight, wdRed)
    Next
    
'    If CC.Tag = "AmountRequested" Then Stop
    Err.Clear
    Rng.HighlightColorIndex = clr
'    Rng.HighlightColorIndex = Clr 'IIf(ErrFlag, wdRed, wdNoHighlight)
    Exit Sub
ex:
'    Stop
'    Resume
End Sub
