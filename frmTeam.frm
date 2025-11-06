VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTeam 
   ClientHeight    =   7890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   OleObjectBlob   =   "frmTeam.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private SelectedItem As Long, Evs As CtrlEvents

Private Sub btnNotify_Click()
    frmCreateNotification.SelectThese CheckedUsers(False)
End Sub

Private Sub btnSend_Click()
    GoToEmail CheckedUsers(True)
End Sub

'Private Sub ListView1_DblClick(): SendEmail: End Sub

'Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem): ForceSelection: End Sub
'
'Private Sub ListView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS): ForceSelection: End Sub
'Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS): ForceSelection: End Sub
Private Sub UserForm_Initialize()
    Dim i As Long, j As Long
    Dim Members As Dictionary ', Transitions As Collection
    Dim CurrentState As String, DocType As String
    CurrentState = GetProperty(pDocState)
    DocType = GetProperty(pDocType)
    Me.Caption = ProjectInfo("very_short_name") & " Team Members"
    Set Members = GetAllMembers
    
    With ListView1
        .View = lvwReport
        .LabelWrap = True
        .Checkboxes = True
        .MultiSelect = False
        .FullRowSelect = True
        .LabelEdit = lvwManual
        
'        .HotTracking = False
        .Gridlines = True
        With .ColumnHeaders
            .Clear
            With .Add()
                .text = "Full Name"
                .Width = 120 '(ListView1.Width / 2) - 1
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Email"
                .Width = 160 '(ListView1.Width / 2) - 1
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Role"
                .Width = 100
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Cellphone"
                .Width = 100
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Office No"
                .Width = 100
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Company"
                .Width = ListView1.Width - 553
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
        End With
        .ListItems.Clear
        For i = 1 To Members.Count
            If Not IsNull(Members(i)("fullname")) Then
                With .ListItems.Add
                    .text = Members(i)("fullname")
                    .ListSubItems.Add , , GetUserInfo(.text, "fullname", "email")
                    .ListSubItems.Add , , GetUserInfo(.text, "fullname", "your_team_role") ', , GetTransitionDescription(Transitions(j))
                    .ListSubItems.Add , , GetUserInfo(.text, "fullname", "cellphone_number") ', , GetTransitionDescription(Transitions(j))
                    .ListSubItems.Add , , GetUserInfo(.text, "fullname", "office_phone_number") ', , GetTransitionDescription(Transitions(j))
                    .ListSubItems.Add , , GetUserInfo(.text, "fullname", "company") ', , GetTransitionDescription(Transitions(j))
                End With
            End If
'            If Members(i) = CurrentState Then SelectedItem = .ListItems.Count
'            Set Transitions = GetStateTransitions(Members(i), DocType)
'            For j = 1 To Transitions.Count
'                With .ListItems.Add
'                End With
'            Next
'            .ListItems.Add
        Next
    End With
    ForceSelection
    Set Evs = New CtrlEvents
    Set Evs.Parent = Me
    Evs.AddOkButton btnNotify
    Evs.AddOkButton btnSend
    Evs.MakeRequired ListView1
End Sub
'Private Sub SendEmail()
'    On Error Resume Next
'    GoToEmail CheckedUsers 'ListView1.SelectedItem.SubItems(1)
'End Sub
Private Sub ForceSelection()
    On Error Resume Next
    If SelectedItem Then
        ListView1.ListItems(SelectedItem).Selected = True
    Else
        ListView1.SelectedItem.Selected = False
    End If
End Sub
Private Function CheckedUsers(Optional EmailsToo As Boolean) As String
    Dim i As Long
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            CheckedUsers = CheckedUsers & "," & ListView1.ListItems(i)
            If EmailsToo Then CheckedUsers = CheckedUsers & "<" & ListView1.ListItems(i).SubItems(1) & ">"
        End If
    Next
    If Len(CheckedUsers) Then CheckedUsers = Right$(CheckedUsers, Len(CheckedUsers) - 1)
End Function
