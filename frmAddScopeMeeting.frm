VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddScopeMeeting 
   Caption         =   "Add a Task Related Meeting"
   ClientHeight    =   8730.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   OleObjectBlob   =   "frmAddScopeMeeting.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddScopeMeeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const CurrentMod = "frmAddScopeMeeting"
Private Evs As New CtrlEvents
Private OldTypes As Dictionary
Private StepsAdded As Long
Private Doc As Document
Private MTypes As Dictionary

Private Sub btnAddAtt_Click()
    Dim i As Long
    i = IsInList
    If i = -1 Then
        liAtt.AddItem cbAttendeeRole
        i = liAtt.ListCount - 1
    End If
    liAtt.List(i, 1) = tbnMtgCount
    Update_btnAddAtt i
End Sub
Private Sub btnDone_Click(): Unload Me: End Sub
Private Function Update_btnAddAtt(Optional i As Long = -2)
    If cbAttendeeRole Like "*Choose Attendee Role*" Then
        btnAddAtt.Caption = "Add"
        btnAddAtt.Enabled = False
    Else
        btnAddAtt.Enabled = Len(tbnMtgCount) > 0
        If i = -2 Then i = IsInList
        If i = -1 Then
            btnAddAtt.Caption = "Add"
        Else
            btnAddAtt.Caption = "Update"
            If btnAddAtt.Enabled Then btnAddAtt.Enabled = liAtt.List(i, 1) <> tbnMtgCount
        End If
    End If
End Function
Private Sub btnOk_Click()
    Dim Tbl As Table
    Set Tbl = Selection.Tables(1)
    If Tbl Is Nothing Then GoTo InvalidDoc
    Dim r As Long, i As Long
    Application.UndoRecord.StartCustomRecord btnOk.Caption & " " & cbMeetingType.value & " Meeting"
    On Error Resume Next
    If OldTypes.Exists(cbMeetingType.value) Then
        For i = Tbl.Rows.Count To 1 Step -1
            If CellText(Tbl.Cell(i, 1)) = cbMeetingType.value Then
                If Err.Number Then
                    Err.Clear
                Else
                    Tbl.Cell(i, 7).Delete wdDeleteCellsEntireRow
                    Do While CellText(Tbl.Cell(i, 1)) = ""
                        Tbl.Cell(i, 7).Delete wdDeleteCellsEntireRow
                        If i > Tbl.Rows.Count Then Exit Do
                    Loop
                    'Exit For
                End If
            End If
        Next
    End If
    OldTypes.Add cbMeetingType.value, cbMeetingType.value
    On Error GoTo 0
    With Tbl
        .Rows.Add
        r = .Rows.Count
        'With .Rows(r)
            .Cell(r, 1).Range.text = cbMeetingType.value
            .Cell(r, 2).Range.text = cbFrequency.value  'tbnDuration.Value
            .Cell(r, 3).Range.text = tbnTotalMeetings.value 'tbnDuration.Value
            .Cell(r, 4).Range.text = tbnLength.value
            .Cell(r, 5).Range.text = tbnPrepTime.value
            For i = 0 To liAtt.ListCount - 1
                If i Then Tbl.Rows.Add
                Tbl.Cell(r + i, 6).Range.text = liAtt.List(i, 0)
                Tbl.Cell(r + i, 7).Range.text = liAtt.List(i, 1)
            Next
        'End With
        If r < .Rows.Count Then
            For i = 1 To 5
                .Cell(r, i).Merge .Cell(.Rows.Count, i)
            Next
        End If
    End With
    cbMeetingType_Change
    StepsAdded = StepsAdded + 1
    Application.UndoRecord.EndCustomRecord
    Exit Sub
InvalidDoc:
    frmMsgBox.Display "Please select the meetings table of the task to be added to.", , Critical, "Invalid Selection"
End Sub
Private Sub cbAttendeeRole_Change(): Update_btnAddAtt: End Sub
Private Sub cbMeetingType_Change()
    On Error Resume Next
    btnOk.Caption = IIf(OldTypes.Exists(cbMeetingType.value), "Update", "Insert")
    cbFrequency.value = MTypes(cbMeetingType.ListIndex)("meeting_frequency")
'    cbFrequency.Value = MTypes(cbMeetingType.ListIndex)("meeting_frequency")
End Sub
Private Sub tbnMtgCount_Change(): Update_btnAddAtt: End Sub
Private Function IsInList() As Long
    Dim i As Long
    For i = 0 To liAtt.ListCount - 1
        If liAtt.List(i, 0) = cbAttendeeRole.value Then Exit For
    Next
    If i < liAtt.ListCount Then 'And liAtt.ListCount > 0 Then
        IsInList = i
    Else
        IsInList = -1
    End If
End Function
Sub Display()
    Dim Tbl As Table, i As Long, s As String
    Set Tbl = CorrectTableSelected
    If Tbl Is Nothing Then Exit Sub
    Set Doc = Tbl.Range.Document
    Set OldTypes = New Dictionary
    On Error Resume Next
    For i = 1 To Tbl.Rows.Count
        s = CellText(Tbl.Cell(i, 1))
        OldTypes.Add s, s
    Next
    Update_btnAddAtt
    Unprotect Doc
    Me.Show
End Sub
Private Sub UserForm_Initialize()
    On Error Resume Next
    Set Evs.Parent = Me
    Evs.AddOkButton btnOk
    Evs.MakeAllRequired "cb,tb"
    Dim OptionsStr As String, i As Long, ss() As String
    Dim RolesColl As Collection, FreqColl As Collection
'    OptionsStr = "General Meeting,Prj Manager Team,Community Stakeholder Group,Technical Advisory Group," & _
'                "Outreach Coordination,City Council Briefing,Community Open House,Environmental Team,Public Engagement," & _
'                "Workshops and Tabling,Underserved Communities"
'    ss = Split(OptionsStr, ",")
'    For i = 0 To UBound(ss)
    Set MTypes = GetMeetingTypes
    For i = 1 To MTypes.Count
        cbMeetingType.AddItem MTypes.KeyName(i)
    Next
    Set RolesColl = ProjectInfo("member_roles")
'    OptionsStr = "Project Manager,Quality Manager,Electrical Lead,Community Involvement"
'    ss = Split(OptionsStr, ",")
'    For i = 0 To UBound(ss)
    For i = 1 To RolesColl.Count
        cbAttendeeRole.AddItem RolesColl(i)("vocabulary_entry") 'Trim(ss(i))
    Next
    Set FreqColl = ProjectInfo("meetings_frequencies")
'    OptionsStr = "Weekly, Semi Weekly, Monthly, Quarterly, Other"
'    ss = Split(OptionsStr, ",")
'    For i = 0 To UBound(ss)
    For i = 1 To FreqColl.Count
        cbFrequency.AddItem FreqColl(i)
    Next
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = IIf(FullColor(ProjectColorStr).TooDark, vbWhite, vbBlack)
    lbPrjHeader.BackColor = ProjectColorStr
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel Button Clicked"
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Doc.Undo StepsAdded
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    Protect Doc
End Sub
