Attribute VB_Name = "AD_AddHeader_mod"
Option Explicit
Private Const CurrentMod = "AddHeader_mod"
'Private Function RoleApplies(Ch As Range, Mode As Long) As Boolean
'    'Mode: 1=Bold, 2=Black
'    Select Case Mode
'    Case 1
'        RoleApplies = Ch.Bold = True
'    Case 2
'        Select Case Ch.Font.ColorIndex
'        Case wdAuto, wdBlack
'        Case Else: RoleApplies = True
'        End Select
'    End Select
'End Function
Sub RemoveHeader()
    WriteLog 1, CurrentMod, "RemoveHeader"
    Application.UndoRecord.StartCustomRecord "Remove Header"
    If Not ValidDocument("Scope") Then Exit Sub
    AllSOWs.RemoveByPosition Selection.start
    Application.UndoRecord.EndCustomRecord
End Sub

Sub InsertHeader(DocType As String)
    WriteLog 1, CurrentMod, "InsertHeader"
    Dim i As Long, Rng As Range, s As String, j As Long, WrdStart As Range
    If Not ValidDocument(DocType) Then Exit Sub
    Unprotect SDoc
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Insert Header"
    Application.DisplayAlerts = wdAlertsNone
    SDoc.Styles.Add "Heading Docent"
    Set WrdStart = Selection.Range
    With WrdStart
        If .Characters.Last = " " Or .Characters.Last = Chr(13) Then .MoveEnd 1, -1
        .MoveStartUntil " ." & Chr(9) & Chr(10) & Chr(13), wdBackward
        .MoveEndUntil " ." & Chr(9) & Chr(10) & Chr(13)
        .Select
    End With
    Set Rng = Selection.Paragraphs(1).Range
    Do
        i = i + 1
        Select Case Asc(Rng.Characters(i))
        Case 45 To 57, 32, 9
        Case Else
            If Not Rng.Characters(i).Bold Then
                If Rng.start <> WrdStart.start Then
                    WriteLog 3, CurrentMod, "InsertHeader", "Heading can only be in the beginning of a paragraph"
                    MsgBox "Heading can only be in the beginning of a paragraph", vbCritical, vbNullString
                    Exit Sub
                ElseIf frmMsgBox.Display(Array("The selection is not Bold.", "Are you sure this should be a header?"), _
                        Array("Yes", "No"), Exclamation, "Docent IMS") = "Yes" Then
                    WriteLog 2, CurrentMod, "InsertHeader", "The selection did not seem like a header, but the user insits"
                    WrdStart.FormattedText.Bold = True
                    Exit Do
                Else
                    Exit Sub
                End If
            End If
        End Select
    Loop Until Rng.Characters(i).Bold
    j = i
    If j > 1 Then
        Rng.Collapse
        Rng.Move 1, j - 1
        Rng.text = vbNewLine
        Rng.Move 1, 1
        Set Rng = Rng.Paragraphs(1).Range
    End If
    j = 1
    If Rng.Characters(j).Bold Then
        Do
            j = j + 1
            If j > Rng.Characters.Count Then Exit Do
            If Asc(Rng.Characters(j)) <> 32 Then If Not Rng.Characters(j).Bold Then Exit Do
        Loop
        If Len(Rng) <> j - 1 Then
            Rng.Collapse
            Rng.Move 1, j - 1
            Rng.MoveEnd 1, 1
            If Asc(Rng.text) = 13 Then
                Rng.MoveEnd 1, -1
            Else
                Rng.MoveEnd 1, -1
                Rng.text = vbNewLine
            End If
        End If
        Rng.Paragraphs(1).Range.Select
        On Error GoTo 0
        AllSOWs.Add(Rng.Paragraphs(1).Range, "Heading Docent").ColorHeader
    End If
    Application.DisplayAlerts = wdAlertsAll
    Application.UndoRecord.EndCustomRecord
End Sub
    


