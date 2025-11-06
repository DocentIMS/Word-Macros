VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateNotification 
   Caption         =   "Create Notification"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "frmCreateNotification.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateNotification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentMod = "frmNotification"
Private Groups As Dictionary, Members As Dictionary
Private Evs As New CtrlEvents

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim Resp As Object 'String
    Dim SelectedMembers As Collection, Groups As Collection, i As Long
'    If ck1Who.Value Then
'        Set SelectedMembers = New Collection
''        For i = 1 To Members.Count
''            SelectedMembers.Add Members(i)("id")
''        Next
'        Set Groups = New Collection
'    Else
        Set SelectedMembers = GetItemsOf(li3Who)
        Set Groups = GetItemsOf(li2Who)
'    End If
    
    Set Resp = CreateAPIContent("Notification", "notifications", _
            Array("notification_type", "message", _
                "user_filter", "notify_users", "notify_groups", _
                "time_filter", "relative_time", "effective_date"), _
            Array(LCase(cbNotifType.value), FakeRitchText(tbText), _
                IIf(ck1Who.value, "true", "false"), SelectedMembers, Groups, _
                IIf(op1Notif.value, "true", "false"), IIf(op1Notif.value, "", Format(Evs.GetDateNTime("Notif"), "HH:mm:ss")), IIf(op1Notif.value, "", Format(tbd2Notif.value, "yyyy-mm-dd"))))
    If IsGoodResponse(Resp) Then
        RefreshNotificationsGroup
        frmMsgBox.Display Array("A new Notification was created on " & ProjectNameStr & " site.", " ", , "View Online"), , Success, "DocentIMS", , , Array(, , , Resp.Data("@id"))
        Unload Me
    Else
        MsgBox "Notification not created", vbCritical, "Docent IMS"
    End If
End Sub
'Private Function GetTime(FldName As String) As String
'    If Len(Controls("tbh" & FldName)) = 0 Then Exit Function
'    GetTime = TimeSerial( _
'            Controls("tb" & FldName & "Hours").Value + _
'            IIf(Controls("cb" & FldName & "AMPM").Value = "PM" And Controls("tb" & FldName & "Hours").Value <> 12, 12, 0), _
'            Controls("tb" & FldName & "Minutes").Value, 0)
'End Function
'Private Function GetDateNTime(FldName As String) As Date
'    Dim ss() As String
'    On Error Resume Next
'    ss = Split(Controls("tb" & FldName), "/")
'    GetDateNTime = DateSerial(ss(2), ss(0), ss(1)) + TimeSerial( _
'            Controls("tb" & FldName & "Hours").Value + _
'            IIf(Controls("cb" & FldName & "AMPM").Value = "PM" And Controls("tb" & FldName & "Hours").Value <> 12, 12, 0), _
'            Controls("tb" & FldName & "Minutes").Value, 0)
'End Function
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
Private Function FakeRitchText(Ctrl As TextBox) As Dictionary
    Set FakeRitchText = New Dictionary
    FakeRitchText.Add "content-type", "text/html"
    FakeRitchText.Add "data", Ctrl.value
    FakeRitchText.Add "encoding", "utf-8"
End Function
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

Private Sub op2Notif_Change()
    Dim EnCustom As Boolean
    EnCustom = Not op2Notif
    Evs.GrayOut "1Notif", EnCustom, False
    op1Notif.Enabled = True
    
    Evs.GrayOut "2Notif", op2Notif, op2Notif
    op2Notif.Enabled = True
    CheckValidForm
End Sub
Private Sub op1Notif_Change()
    Dim EnCustom As Boolean
    EnCustom = Not op1Notif
    Evs.GrayOut "2Notif", EnCustom, False
    op2Notif.Enabled = True
    CheckValidForm
End Sub
Private Sub CheckValidForm()
    Dim En As Boolean
'    En = True
    En = op1Notif Or Evs.IsValid("2Notif")
    If En Then En = ck1Who Or Evs.IsValid("2Who") Or Evs.IsValid("3Who")
    btnOk.Enabled = En
End Sub

Private Sub cbtt2Notif_Change(): CheckValidForm: End Sub
Private Sub tbh2Notif_Change(): CheckValidForm: End Sub
Private Sub tbd2Notif_Change(): CheckValidForm: End Sub
Private Sub tbm2Notif_Change(): CheckValidForm: End Sub
Private Sub tbText_Change(): CheckValidForm: End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Set Evs.Parent = Me
    On Error Resume Next
'    Evs.CollectAllControls
'    Evs.AddOkButton btnOk
    Evs.MakeRequired "Notif,cbNotifType,tbText,frRecipients,frWhen", , ErrorColor ',frWhen,frRecipients
    
    For i = 1 To ProjectInfo("notification_types").Count
        cbNotifType.AddItem ProjectInfo("notification_types")(i)
    Next
'    cbNotifType.AddItem "Info"
'    cbNotifType.AddItem "Warning"
'    cbNotifType.AddItem "Error"
'    cbNotifType.AddItem "Basic"
    Set Groups = GetAllGroups("can_*")
    For i = 1 To Groups.Count
        li2Who.AddItem Groups(i)("title")
    Next
    Set Members = GetAllMembers
    For i = 1 To Members.Count
        If Not IsNull(Members(i)("fullname")) Then
            li3Who.AddItem Members(i)("fullname")
        End If
    Next
    Evs("li2Who").GrayOut ck2Who.value, ck2Who.value
    Evs("li3Who").GrayOut ck3Who.value, ck3Who.value
    Evs.GrayOut "2Notif", op2Notif, op2Notif
    Evs("op2Notif").GrayOut True, True
    op2Notif.Enabled = True
    op1Notif.value = True
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
    lbPrjHeader.BackColor = ProjectColorStr
    CheckValidForm
End Sub
Sub SelectThese(Users As String)
    Dim ss() As String, i As Long
    ss = Split(Users, ",")
    For i = 0 To UBound(ss)
        SelectListItem li2Who, ss(i)
        SelectListItem li3Who, ss(i)
    Next
    Me.Show
End Sub
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
