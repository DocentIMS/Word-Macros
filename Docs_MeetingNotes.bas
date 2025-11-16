Attribute VB_Name = "Docs_MeetingNotes"
Option Explicit
Private Const CurrentMod As String = "Docs_MeetingNotes"

'Sub FillMeetingNotesCC()
'
'    SetContentControl "PropManagerAttended", True
'    SetContentControl "ActualTime", GetContentControl("PlannedTime")
'    SetContentControl "ActualDate", GetContentControl("PlannedDate")
'    SetContentControl "PropManagerAttended", GetContentControl("PropManagerAttending")
'
'End Sub
'Private Sub FillPlannedGroups(Coll As Collection, Optional Doc As Document)
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, r As Long, i As Long, j As Long
'    Set Tbl = GetTableByTitle("Meeting Details", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    r = 2
'    Do: r = r + 1: Loop Until CellText(Tbl.Rows(r - 1).Cells(1).Range.Text) = "Groups"
'    If CellText(Tbl.Rows(r).Cells(1).Range.Text) <> "No Groups" Then Exit Sub
'
'
'    If Coll.Count > 1 Then
'    End If
'    For i = Coll.Count To 1 Step -1
''        r = r + 1
'        If i < Coll.Count Then
'            For j = 1 To Tbl.Rows(r).Range.ContentControls.Count
'                Tbl.Rows(r).Range.ContentControls(j).LockContentControl = False
'                Tbl.Rows(r).Range.ContentControls(j).LockContents = False
'            Next
'            Tbl.Rows.Add Tbl.Rows(r)
'        End If
'        FillPlannedGroupRow Tbl.Rows(r), CStr(Coll(i)("token")) '("title"))
'    Next
'    If Coll.Count > 1 Then
'        For r = r To r + Coll.Count - 1
'            For i = 1 To Tbl.Rows(r).Range.ContentControls.Count
'                Tbl.Rows(r).Range.ContentControls(i).LockContentControl = True
'            Next
'            Tbl.Rows(r).Cells(4).Range.ContentControls(1).LockContents = True
'        Next
'    End If
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
Function IsValidMeetingNotes(Doc As Document) As Boolean
    Dim Tbl As Table, r As Long, i As Long, j As Long
    Set Tbl = GetTableByTitle("Meeting Details", Doc)
    If Tbl Is Nothing Then Exit Function
    IsValidMeetingNotes = True
    r = 2
    Do: r = r + 1: Loop Until cellText(Tbl.Rows(r - 1).Cells(1).Range.text) = "Groups"
    If cellText(Tbl.Rows(r).Cells(1).Range.text) = "No Groups" Then Exit Function
    Do
        If Tbl.Rows(r).Cells.Count < 3 Then Exit Do
        If cellText(Tbl.Rows(r).Cells(3).Range.text) = "Select Attendees" Then
            IsValidMeetingNotes = False
            Tbl.Rows(r).Cells(3).Range.Font.Bold = True
            Tbl.Rows(r).Cells(3).Range.Font.Color = vbRed
        ElseIf Tbl.Rows(r).Cells(3).Range.Font.Color = vbRed Then
            Tbl.Rows(r).Cells(3).Range.Font.Bold = False
            Tbl.Rows(r).Cells(3).Range.Font.Color = Tbl.Rows(r).Cells(1).Range.Font.Color
        End If
        r = r + 1
    Loop Until cellText(Tbl.Rows(r - 1).Cells(1).Range.text) = "Individuals"
End Function
Sub ShowTimePicker(CCName As String)
    With frmTimePicker
        .CCName = CCName
        .Load
    End With
End Sub
Sub ShowAttendees(ByVal CC As ContentControl)
    On Error Resume Next
    Dim GroupID As String
    GroupID = Replace(CC.Title, " Attendees", "")
    If ProjectGroupsDict Is Nothing Then RefreshRibbon
'    On Error GoTo 0
    With frmAttendees
        .GroupID = GroupID
        Select Case GroupID
        Case "Individual"
            Set .ItemsDict = GetMembersNames
            .Caption = GroupID & " Attendees"
        Case "External"
            Set .ItemsDict = ArrToDict(Split(cellText(CC.Range.Cells(1).Range.Previous(wdCell)), Chr(13)))
            .Caption = GroupID & " Attendees"
        Case Else
            .Caption = ProjectGroupsDict(GroupID)("title") & " Attendees"
        End Select
'        If GroupID = "Individual" Then GroupID = "PrjTeam"
        
'        .OldItems = GetContentControl(GroupID & " Attendees")
'        Set .ItemsColl = GetMembersOf(GroupId)
        .Show
    End With
End Sub
Sub FillNotesFromZoom(Optional ID As String)
    Dim MtgDict As Dictionary
    On Error Resume Next
    If Len(ID) Then
        Set MtgDict = GetPastMeeting(Replace(ID, " ", ""))
    Else
        Set MtgDict = GetPastMeeting(GetProperty(pOnlineMeetingUID, ActiveDocument))
    End If
'    Set MtgDict = GetZoomMtg(GetProperty(pOnlineMeetingUID, ActiveDocument))
    If MtgDict Is Nothing Then
        frmMsgBox.Display Array("The meeting info is not ready yet.", "Usually, it is ready a few minutes after the meeting ending."), , Exclamation, "Docent IMS"
    Else
        Dim sTime As Date, eTime As Date, Participants As Collection
        Dim Members As New Dictionary, Groups As Dictionary, Tbl As Table, InGroups As New Dictionary, AllMembers As Dictionary
        Dim Dict As Dictionary
        Dim i As Long, j As Long, k As Long
        Set Tbl = GetTableByTitle("Meeting Details", ActiveDocument)
        sTime = TimeFromTFormat(MtgDict("start_time"))
        eTime = TimeFromTFormat(MtgDict("end_time"))
        Set Participants = GetZoomAttList(MtgDict("uuid"))
        For i = 2 To Tbl.Rows.Count
            If cellText(Tbl.Cell(i, 1)) = "Groups" Then Exit For
        Next
        If i < Tbl.Rows.Count Then
            i = i + 1
            Do Until cellText(Tbl.Cell(i, 1)) = "Individuals"
                InGroups.Add cellText(Tbl.Cell(i, 1)), i
                i = i + 1
            Loop
            Do Until cellText(Tbl.Cell(i, 1)) = "Individual Attendees": i = i + 1: Loop
            InGroups.Add cellText(Tbl.Cell(i, 1)), i
            Do Until cellText(Tbl.Cell(i, 1)) = "External Attendees": i = i + 1: Loop
            InGroups.Add cellText(Tbl.Cell(i, 1)), i
        End If
        Set Groups = GetAllGroups("can_*")
        For i = 1 To Groups.Count
            If InGroups.Exists(Groups(i)("title")) Then
                Members.Add InGroups(Groups(i)("title")), Groups(i)("groupMembers")
            End If
        Next
        Members.Add InGroups("Individual Attendees"), Nothing
        Members.Add InGroups("External Attendees"), Nothing
        Set AllMembers = GetAllMembers
        On Error Resume Next
'        Set InGroups = New Dictionary
        For i = 1 To Participants.Count
            With FindRow(Members, AllMembers, CStr(Participants(i)("name")), CStr(Participants(i)("user_email")))
                AddToCell Tbl, .Item("row"), .Item("name")
            End With
        Next
'         Msgs() As String
'        ReDim Msgs(0 To (Participants.Count * 3) + 3)
'        Msgs(0) = "Meeting Start: " & sTime
'        Msgs(1) = "Meeting End: " & eTime
'        Msgs(2) = "Attendees:"
'        For i = 1 To Participants.Count
'            If Len(Participants(i)("user_email")) Then
'                Msgs((i - 1) * 3 + 3) = " " & Participants(i)("name") & ": " & Participants(i)("user_email")
'            Else
'                Msgs((i - 1) * 3 + 3) = " " & Participants(i)("name")
'            End If
'            Msgs((i - 1) * 3 + 4) = "  Joined: " & TimeFromTFormat(CStr(Participants(i)("join_time")))
'            Msgs((i - 1) * 3 + 5) = "  Left: " & TimeFromTFormat(CStr(Participants(i)("leave_time")))
'        Next
'        frmMsgBox.Display Msgs, Clrs:=0
    End If
End Sub
Private Sub AddToCell(Tbl As Table, r As Long, ByVal s As String)
    Dim so As String
    so = Tbl.Cell(r, 3).Range.ContentControls(1).Range.text
    If Not so Like "Select *" And Not so Like "N/A" Then
        Tbl.Cell(r, 3).Range.ContentControls(1).Range.text = so & vbNewLine & s
    Else
        Tbl.Cell(r, 3).Range.ContentControls(1).Range.text = s
    End If
End Sub
Private Function FindRow(Members As Dictionary, AllMembers As Dictionary, UserName As String, UserEmail As String) As Dictionary
    Dim FindBy As String, FindWhat As String, j As Long, k As Long, Dict As Dictionary
    If Len(UserEmail) Then
        FindBy = "email"
        FindWhat = UserEmail
    Else
        FindBy = "fullname"
        FindWhat = UserName
    End If
    For j = 1 To Members.Count - 2 'Each InGroup
        For k = 1 To Members(j).Count 'Each Member
            If Members(j)(k)(FindBy) = FindWhat Then
                Set Dict = New Dictionary
                Dict.Add "row", Members.KeyName(j)
                Dict.Add "name", Members(j)(k)("fullname")
                Set FindRow = Dict
                Exit Function
            End If
        Next
    Next
    For k = 1 To AllMembers.Count 'Each Member
        If AllMembers(k)(FindBy) = FindWhat Then
'            FindRow = Members.KeyName("Individual Attendees")
            Set Dict = New Dictionary
            Dict.Add "row", Members.KeyName(Members.Count - 1)
            Dict.Add "name", AllMembers(k)("fullname")
            Set FindRow = Dict
            Exit Function
        End If
    Next
    If Len(UserEmail) Then Set FindRow = FindRow(Members, AllMembers, UserName, "")
    If FindRow Is Nothing Then
        Set Dict = New Dictionary
        Dict.Add "row", Members.KeyName(Members.Count)
        Dict.Add "name", UserName
        Set FindRow = Dict
        Exit Function
    End If
End Function
