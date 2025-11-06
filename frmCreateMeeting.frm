VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateMeeting 
   Caption         =   "Add Meeting"
   ClientHeight    =   10410
   ClientLeft      =   150
   ClientTop       =   570
   ClientWidth     =   10005
   OleObjectBlob   =   "frmCreateMeeting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateMeeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Const CurrentMod = "frmCreateMeeting"
Private MTypes As Dictionary, Groups As Dictionary, Members As Dictionary
Private DatePickers As New Collection, TimeSPPickers As New Collection
Private mURL As String, oldState As String, mLastLoc As String
Private OFiles As New Dictionary, OFilesRem As New Dictionary
Private mDocsURL(1 To 3) As String
Private mManualTitle As Boolean, mOldTitle As String
Private Evs As New CtrlEvents ', RequiredFields As New Collection
Sub OpenMeeting(URL As String)
    Dim i As Long
    Me.Caption = "Edit Meeting"
    mManualTitle = True
    mURL = URL
    On Error Resume Next
    With GetAPIContent(URL).Data
'Set OFilesRem = GetAPIContent(URL).Data
        tbURL.value = .Item("event_url")
        cbMeetingType.value = .Item("meeting_type")("title")
        SelectListItem li3Who ' li3Who
        SelectListItem li2Who 'li2Who
        tbText.value = Replace(.Item("text")("data"), Chr(13) & Chr(13), "")
        For i = 1 To .Item("attendees").Count
            SelectListItem li3Who, .Item("attendees")(i)("title")
        Next
        For i = 1 To .Item("attendees_group").Count
            SelectListItem li2Who, .Item("attendees_group")(i)("token") '("title")
        Next
        If TypeName(.Item("contact_name")) = "String" Then
            cbContact.value = GetMemberName(.Item("contact_name")) '("title")
        Else
            cbContact.value = .Item("contact_name")("title")
        End If
        SetDateTime "Starts", ToServerTime(CStr(.Item("start"))) 'TimeFromTFormat(.Item("start"))
        SetDateTime "Ends", ToServerTime(CStr(.Item("end"))) ' TimeFromTFormat(.Item("end"))
        tbTitle.value = .Item("title")
        If Not IsNull(.Item("event_url")) Then tbURL.value = .Item("event_url")
        cbLocation.value = .Item("location")("title")
        oldState = GetStateName(.Item("review_state"), "meeting")
        cbState = oldState
'        If .Item("will_property_manager_attend") Then
'            op1PropMgr = True
'        Else
'            op2PropMgr = True
'        End If
'        btnPublish.Visible = oldState = "Private"
        On Error Resume Next
        For i = 1 To 3
            mDocsURL(i) = .Item("meeting_document_" & i)("download")
'            With Controls("tbDoc" & i)
            Controls("tbDoc" & i).value = .Item("meeting_document_" & i)("filename")
            Controls("tbDoc" & i).Font.Underline = True
            Controls("tbDoc" & i).ForeColor = vbBlue
'            Controls("tbDoc" & i).
'            End With
        Next
    End With
    Set OFiles = CollToDict(GetAPIFolder(mURL, "File"), "title")
    For i = 1 To OFiles.Count
        liFiles.AddItem OFiles(i)("title")
    Next
    btn_OK.Caption = "Update"
    btn_OK.Enabled = False
    Me.Show
End Sub

Private Sub btnAddExAtt_Click()
    liExAtt.AddItem tbExAtt.value
    tbExAtt.value = ""
End Sub
Private Sub btnRemExAtt_Click()
    Dim i As Long
    For i = liExAtt.ListCount To 1 Step -1
        If liExAtt.Selected(i - 1) Then liExAtt.RemoveItem i - 1
    Next
End Sub

Private Sub btnFilesRem_Click()
    Dim i As Long
    For i = liFiles.ListCount To 1 Step -1
        If liFiles.Selected(i - 1) Then
            If OFiles.Exists(liFiles.List(i - 1)) Then OFilesRem.Add liFiles.List(i - 1), OFiles(liFiles.List(i - 1))("@id")
            liFiles.RemoveItem i - 1
        End If
    Next
End Sub
Private Sub btnFilesAdd_Click()
    Dim Files As Object, i As Long
    Set Files = GetFile("Browse to file(s) to attach", mLastLoc, True, AllFiles)
    If Not Files Is Nothing Then
        For i = 1 To Files.Count
            liFiles.AddItem Files(i)
        Next
    End If
End Sub
Private Sub AddAttachment(i As Long)
    WriteLog 1, CurrentMod, "AddAttachment", "Attachment Button " & i & " Clicked"
    On Error Resume Next
    If Len(mLastLoc) = 0 Then mLastLoc = Environ("Userprofile") & "\desktop"
    With Controls("tbDoc" & i)
        .value = GetFile("Browse to text file", mLastLoc, False, AllFiles)(1)
        If Err.Number = 0 Then
            .Font.Underline = False
            .ForeColor = vbBlack
            mDocsURL(i) = ""
        End If
    End With
    mLastLoc = GetParentDir(Controls("tbDoc" & i).value)
End Sub
'Private Sub btnPublish_Click()
'    DoAction "Publish"
'End Sub
Private Sub btn_OK_Click()
    DoAction IIf(Len(mURL), "updated", "created")
End Sub
Private Sub DoAction(action As String)
    Dim Mtg As New Dictionary, msgs(1 To 5) As String
    If Not ValidDates Then
        Evs.MarkInvalid "Ends"
        frmMsgBox.Display "The meeting end must be after the meeting start", , Critical, "Docent IMS"
        Exit Sub
    End If
    Dim Coll As New Collection, Resp, GroupsColl As New Collection, IndColl As New Collection ' As String
    Coll.Add cbMeetingType.value
    If ck1Who Then
        Set IndColl = Nothing
        GroupsColl.Add "PrjTeam"
    Else
        Set GroupsColl = GetItemsOf(li2Who)
        Set IndColl = GetItemsOf(li3Who)
    End If
    Mtg.Add "text", FakeRitchText(tbText)
    Mtg.Add "attendees", IndColl
    Mtg.Add "attendees_group", GroupsColl
    Mtg.Add "contact_name", GetMemberID(cbContact.value)
    Mtg.Add "meeting_type", cbMeetingType.value
    Mtg.Add "start", AlreadyServerTime(Evs.GetDateNTime("Starts"))
    Mtg.Add "end", AlreadyServerTime(Evs.GetDateNTime("Ends"))
    Mtg.Add "title", tbTitle.value
    Mtg.Add "type_title", cbMeetingType.value
    Mtg.Add "subjects", Coll
    Mtg.Add "location", cbLocation.value
    Mtg.Add "entire_team", ck1Who.value
    If oldState = GetStateID("Published", "meeting") Then
        UpdateAPIFileWorkflow mURL, GetTransitionID("Retract", "meeting")
        oldState = GetStateID("Private", "meeting")
    End If
    If cbState.value = "Published" And action = "created" Then Mtg.Add "transition_target", "publish"
    If action = "created" Then
        Set Resp = CreateAPIContent("meeting", DefaultMeetingsFolder, Mtg.Keys, Mtg.Items)
        If IsGoodResponse(Resp) Then mURL = Resp.Data("@id")
    Else 'updated
        Set Resp = UpdateAPIContent(mURL, Mtg.Keys, Mtg.Items)
    End If
    msgs(1) = cbMeetingType.value & " was " & action & "." & vbNewLine & vbNewLine
    'msgs(2) = "View Online" & vbNewLine & vbNewLine
    msgs(3) = CollectInvitees(1).Count & " attendees notified" & vbNewLine & vbNewLine
    msgs(4) = "Draft meeting agenda, notes and minutes created." & vbNewLine & vbNewLine
    
    
    If IsGoodResponse(Resp) Then
        If oldState <> cbState.value Then UpdateAPIFileWorkflow mURL, GetTransitionIdByStates(oldState, cbState.value, "meeting")
        Dim i As Long
        For i = 1 To liFiles.ListCount
            If Not OFiles.Exists(liFiles.List(i - 1)) Then
                UploadAPIFile liFiles.List(i - 1), mURL
            End If
        Next
        For i = 1 To OFilesRem.Count
            DeleteAPIContent OFilesRem(i)
        Next
        If InStr(1, cbLocation.value, "Zoom", vbTextCompare) Then
            Set Resp = CreateZoomMeeting(tbTitle.value, Evs.GetDateNTime("Starts"), _
                        DateDiff("n", Evs.GetDateNTime("Starts"), Evs.GetDateNTime("Ends")), CollectInvitees)
            With Resp
                Set Resp = UpdateAPIContent(mURL, Array("online_meeting_link", "event_url"), Array(.Data("uuid"), .Data("join_url")))
                        
            End With
            Me.Hide
            msgs(5) = "The Zoom meeting was also " & action
            frmMsgBox.Display msgs, _
                        Array(), _
                        Images:=Array(Success, , , , ZoomIcon), _
                        Clrs:=Array(0, 0, 0, 0, 0), _
                        Title:="DocentIMS", _
                        Links:=Array(mURL), _
                        AutoCloseTimer:=3
        Else
            frmMsgBox.Display msgs, , Success, "DocentIMS", , , Array(mURL)
        End If
'        frmMsgBox.Display "A " & ProjectInfo("very_short_name") & " meeting was " & Action
        Unload Me
    Else
        frmMsgBox.Display "Meeting not " & action, Array(), Critical, "Docent IMS"
    End If
End Sub
Private Function CollectInvitees() As Dictionary
    Dim Coll As New Collection, Dict As Dictionary, i As Long, j As Long, EmailStr As String
    Dim IndColl As Collection, GroupsColl As Collection, GroupMembers As Dictionary
    On Error Resume Next
    For i = 1 To liExAtt.ListCount
        Set Dict = New Dictionary
        EmailStr = liExAtt.List(i - 1)
        Dict.Add "email", EmailStr
        Coll.Add Dict, EmailStr
    Next
    If Len(tbExAtt.value) Then
        Set Dict = New Dictionary
        EmailStr = tbExAtt.value
        Dict.Add "email", EmailStr
        Coll.Add Dict, EmailStr
    End If
    If ck1Who Then
        Set IndColl = Nothing
        GroupsColl.Add "PrjTeam"
    Else
        Set GroupsColl = GetItemsOf(li2Who)
        Set IndColl = GetItemsOf(li3Who)
    End If
    For i = 1 To IndColl.Count
        Set Dict = New Dictionary
        EmailStr = GetUserInfo(IndColl(i), "id", "email")
        Dict.Add "email", EmailStr
        Coll.Add Dict, EmailStr
    Next
    For i = 1 To GroupsColl.Count
        Set GroupMembers = GetMembersOf(GroupsColl(i), "email")
        For j = 1 To GroupMembers.Count
            Set Dict = New Dictionary
            EmailStr = GroupMembers(j)
            Dict.Add "email", EmailStr
            Coll.Add Dict, EmailStr
        Next
    Next
    Set Dict = New Dictionary
    Dict.Add "meeting_invitees", Coll
    Set CollectInvitees = Dict
End Function
Private Sub SelectListItem(li As MSForms.ListBox, Optional Item)
    Dim i As Long
    If IsMissing(Item) Then
        For i = 1 To li.ListCount
            li.Selected(i - 1) = False
        Next
    Else
        For i = 1 To li.ListCount
            If li.List(i - 1) = Item Then
                Controls("ck" & Right$(li.Name, Len(li.Name) - 2)).value = True
                li.Selected(i - 1) = True
                Exit Sub
            End If
        Next
    End If
End Sub
Private Function ValidDates() As Boolean
    On Error Resume Next
    ValidDates = Evs.GetDateNTime("Ends") > Evs.GetDateNTime("Starts")
End Function
Private Function FakeRitchText(Ctrl As TextBox) As Dictionary
    Set FakeRitchText = New Dictionary
    FakeRitchText.Add "content-type", "text/html"
    FakeRitchText.Add "data", Ctrl.value
    FakeRitchText.Add "encoding", "utf-8"
End Function
Private Function GetItemsOf(Ctrl As ListBox) As Collection
    Dim i As Long, userID As String
    On Error Resume Next
    Set GetItemsOf = New Collection
    For i = 0 To UBound(Ctrl.List)
        If Ctrl.Selected(i) Then
            userID = ""
            userID = Members(Ctrl.List(i))("id")
            If Len(userID) = 0 Then userID = Groups(Ctrl.List(i))("id")
            If Len(userID) = 0 Then userID = Ctrl.List(i)
            GetItemsOf.Add userID 'Ctrl.List(i)
        End If
    Next
End Function
Private Sub UpdateTitle()
    On Error Resume Next
    If mManualTitle Then Exit Sub
    If Len(tbdStarts.value) > 0 Then 'meeting_title
        tbTitle.value = MTypes(cbMeetingType.value)("meeting_type") & " - " & Format(tbdStarts.value, LongDateFormat)
    Else
        tbTitle.value = MTypes(cbMeetingType.value)("meeting_type")
    End If
End Sub

Private Sub cbMeetingType_Change()
    Dim Coll As Collection, i As Long, MembersByID As Dictionary, GroupsByID As Dictionary
    On Error Resume Next
    If Not MTypes.Exists(cbMeetingType.value) Then Exit Sub
    If IsEmpty(MTypes(cbMeetingType.value)) Then Exit Sub
    Set MembersByID = GetAllMembers("id")
    Set GroupsByID = GetAllGroups("can_*", True)
    UpdateTitle 'tbTitle.Value = MTypes(cbMeetingType.Value)("meeting_title")
    If IsNull(MTypes(cbMeetingType.value)("meeting_contact")) Then
        cbContact.value = ""
    Else
        cbContact.value = MembersByID(MTypes(cbMeetingType.value)("meeting_contact"))("fullname")
    End If
    Set Coll = MTypes(cbMeetingType.value)("meeting_attendees")
    SelectListItem li2Who
    SelectListItem li3Who
    For i = 1 To Coll.Count
        SelectListItem li2Who, GroupsByID(Coll(i))("title")
'        SelectListItem li3Who, MembersByID(Coll(i))("fullname")
'        If Not SelectGroup(Coll(i)) Then SelectMember (Coll(i))
    Next
End Sub
Private Function SelectMember(memberID) As Boolean
    Dim i As Long
    For i = 0 To UBound(li3Who.List)
        li3Who.Selected(i) = GetMemberID(li3Who.List(i)) = memberID
        SelectMember = SelectMember Or li3Who.Selected(i)
    Next
End Function
Private Function SelectGroup(GroupID) As Boolean
    Dim i As Long
    For i = 0 To UBound(li2Who.List)
        li2Who.Selected(i) = li2Who.List(i) = GroupID
        SelectGroup = SelectGroup Or li2Who.Selected(i)
    Next
End Function
Private Sub SetDateTime(FldName As String, Dt As Variant)
    Dim h As Long
    h = Format(TimeValue(Dt), "h")
    If h = 0 Then
        Controls("tbh" & FldName).value = 12
        Controls("tbtt" & FldName).value = "AM"
    Else
        Controls("tbh" & FldName).value = IIf(h > 12, h - 12, h)
        Controls("tbtt" & FldName).value = IIf(h >= 12, "PM", "AM")
    End If
    Controls("tbm" & FldName).value = Format(TimeValue(Dt), "nn")
    Controls("tbd" & FldName).value = Format(DateValue(Dt), DateFormat)
End Sub
Private Sub tbdStarts_Change()
    UpdateTitle
    tbdEnds.value = tbdStarts.value
End Sub
Private Sub tbhStarts_Change()
    On Error Resume Next
    If tbhStarts.value = 12 Then
        sphEnds.value = 1
    Else
        sphEnds.value = tbhStarts.value + 1
    End If
    tbhEnds.value = sphEnds.value
    tbttStarts_Change
End Sub
Private Sub tbmStarts_Change()
    On Error Resume Next
    spmEnds.value = tbmStarts.value
End Sub
Private Sub tbttStarts_Change()
    On Error Resume Next
    If Not IsEmpty(Evs.GetDateNTime("Starts")) Then SetDateTime "Ends", Evs.GetDateNTime("Starts") + (1 / 24)
End Sub
Private Sub CheckValidForm()
    Evs.CheckOk
    If btn_OK.Enabled Then btn_OK.Enabled = ck1Who Or Evs.IsValid("2Who") Or Evs.IsValid("3Who")
End Sub
Private Sub ck3Who_Change()
    Dim EnCustom As Boolean
    EnCustom = Not ck1Who Or (ck2Who.value Or ck3Who.value)
    GrayoutWho 2, EnCustom
    GrayoutWho 3, EnCustom
    EnCustom = Not (ck2Who.value Or ck3Who.value)
    Evs.GrayOut "1Who", EnCustom, False
    ck1Who.Enabled = True
    CheckValidForm
End Sub
Private Sub ck2Who_Change()
    Dim EnCustom As Boolean
    EnCustom = Not ck1Who Or (ck2Who.value Or ck3Who.value)
    GrayoutWho 2, EnCustom
    GrayoutWho 3, EnCustom
    EnCustom = Not (ck2Who.value Or ck3Who.value)
    Evs.GrayOut "1Who", EnCustom, False
    ck1Who.Enabled = True
    CheckValidForm
End Sub
Private Sub ck1Who_Change()
    Dim EnCustom As Boolean
    EnCustom = ck1Who.value
    Evs.GrayOut "1Who", ck1Who.value Or (Not (ck2Who.value Or ck3Who.value)), True
    EnCustom = Not EnCustom
    GrayoutWho 2, EnCustom
    GrayoutWho 3, EnCustom
    ck2Who.Enabled = EnCustom
    ck3Who.Enabled = EnCustom
    CheckValidForm
End Sub
Private Sub GrayoutWho(n As Long, EnCustom As Boolean)
    Evs.GrayOut n & "Who", EnCustom, True
    If Not ck1Who.value Then Evs("li" & n & "Who").GrayOut Controls("ck" & n & "Who").value, Controls("ck" & n & "Who").value
End Sub
Private Sub li2Who_Change(): CheckValidForm: End Sub
Private Sub li3Who_Change(): CheckValidForm: End Sub

'Private Sub hbEnds_Change(): ValidDates: End Sub
'Private Sub ampmStarts_Change(): ValidDates: End Sub
'Private Sub ampmEnds_Change(): ValidDates: End Sub
Private Sub tbTitle_Enter(): mOldTitle = tbTitle.value: End Sub
Private Sub tbTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean): mManualTitle = mOldTitle <> tbTitle.value: End Sub
Private Sub tbURL_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(tbURL.value) Then If Left$(tbURL.value, 4) <> "http" Then tbURL.value = "https://" & tbURL.value
End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    
    Set MTypes = GetMeetingTypes
'    cbMeetingType.AddItem "-- Choose Meeting Type --"
    For i = 1 To MTypes.Count
        If Len(MTypes.KeyName(i)) > 0 Then cbMeetingType.AddItem MTypes.KeyName(i)
    Next
'    cbLocation.AddItem "-- Choose Location --"
'    For i = 1 To ProjectInfo("meeting_locations").Count
'        If Not IsNull(ProjectInfo("meeting_locations")(i)(1)) Then cbLocation.AddItem ProjectInfo("meeting_locations")(i)(1)
'    Next
    cbLocation.AddItem "Teams"
    cbLocation.AddItem "Zoom"
    cbLocation.AddItem "Client Office"
    cbLocation.AddItem "Client Office and Teams"
    cbLocation.AddItem "Client Office and Zoom"
    
    Set Groups = GetAllGroups("can_*")
    For i = 1 To Groups.Count
        li2Who.AddItem Groups(i)("title")
    Next
    Set Members = GetAllMembers 'GetMembersOf()
    For i = 1 To Members.Count
        If Not IsNull(Members(i)("fullname")) Then
            li3Who.AddItem Members(i)("fullname")
            cbContact.AddItem Members(i)("fullname")
        End If
    Next
    cbState.AddItem "Private"
    cbState.AddItem "Published"
    cbState.ListIndex = 0
    oldState = GetStateID("Private", "meeting")
'    cbMeetingType.ListIndex = 0
'    cbLocation.ListIndex = 0
'    ValidDates
    Set Evs.Parent = Me
'    Evs.CollectAllControls
    Evs.AddOkButton btn_OK
'    Evs.AddOkButton btnPublish
    Evs.MakeRequired "cbMeetingType,cbLocation,tbTitle,tbdStarts,tbdEnds,PropMgr,cbContact,tbText,cbState", , ErrorColor
    Evs("li2Who").GrayOut ck2Who.value, ck2Who.value
    Evs("li3Who").GrayOut ck3Who.value, ck3Who.value
    lbPrjHeader.Caption = ProjectNameStr
'    lbPrjHeader2.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = IIf(FullColor(ProjectColorStr).TooDark, vbWhite, vbBlack)
    lbPrjHeader.BackColor = ProjectColorStr
'    lbPrjHeader.BackStyle = fmBackStyleOpaque

'    Evs.MakeRequired "Starts,Ends", "*", ErrorColor
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel Button Clicked"
    Unload Me
End Sub


