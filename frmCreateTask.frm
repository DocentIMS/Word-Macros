VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateTask 
   Caption         =   "Create a Task"
   ClientHeight    =   6660
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   9735.001
   OleObjectBlob   =   "frmCreateTask.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PArr As Variant, PColl As New Collection, BColl As New Collection ', BoardGroupID As Long
Private mUploadMode As Boolean ', DatePicker As New DateTimePickerEvents
Private Evs As New CtrlEvents ', RequiredFields As New Collection
Private Sub btn_OK_Click()
    Dim Resp As WebResponse
'    Dim MtngType As String
'    MtngType = GetProperty(pMeetingType)
    If Len(tbDetails.value) = 0 Then tbDetails.value = " " 'To bypass the server-side validation
    If mUploadMode Then
        Set Resp = CreateAPITask(tbTitle.value, tbDetails.value, _
                                tbdDueDate.value, PColl(cbPriority.value), _
                                GetMemberID(cbWho.value), tbNotes.value, tbPrivateNotes.value, , , GetItemsOf(liOthers))
        If Not IsGoodResponse(Resp) Then
            MsgBox "Task could not be created", vbCritical, "DocentIMS"
            Exit Sub
        End If
        If cbState.value <> "Private" Then
            UpdateAPIFileWorkflow Resp.Data("@id"), GetTransitionIdByStates("Private", cbState.value, "action_items")
            RefreshTasksGroup
        End If
        frmMsgBox.Display Array("A new Task was created on " & ProjectNameStr & " site.", " ", , "View Online"), , Success, "DocentIMS", , , Array(, , , Resp.Data("@id"))
        Unload Me
    Else
        SaveToMetadata
        AddToTaskTable
        Unload Me
'        Dim URL As String, FName As String
'        URL = GetProperty(pDocURL)
'        FName = GetFileName(URL)
'        FName = "minute" & Right$(FName, Len(FName) - 4)
'        URL = GetParentDir(URL) & FName
        'UpdateAPIContent GetProperty(pDocURL), Array("proposed_action_items"), Array(GetProperty(pProposedTasks))
    End If
End Sub
Private Sub SaveToMetadata()
    Dim s As String
    On Error Resume Next
    s = GetProperty(pProposedTasks)
    s = s & ";," & tbTitle.value & "," & _
                GetMemberID(cbWho.value) & "," & _
                PColl(cbPriority.value) & "," & _
                tbdDueDate.value & "," & _
                tbDetails.value & "," & _
                tbNotes.value & "," & tbPrivateNotes.value '& "," & _
                Join(CollToArr(GetItemsOf(liOthers)), ";")
    SetProperty pProposedTasks, s
End Sub
Private Function GetItemsOf(Ctrl As ListBox) As Collection
    Dim i As Long, UserID As String
    On Error Resume Next
    Set GetItemsOf = New Collection
    For i = 0 To UBound(Ctrl.List)
        If Ctrl.Selected(i) Then
            UserID = GetMemberID(Ctrl.List(i))
'            UserID = Members(Ctrl.List(i))("id")
            If Len(UserID) = 0 Then UserID = Ctrl.List(i)
            GetItemsOf.Add UserID 'Ctrl.List(i)
        End If
    Next
End Function
Private Sub AddToTaskTable()
    Dim i As Long
    With ActiveDocument
        Unprotect
        For i = 1 To .Tables.Count
            If .Tables(i).Title = "Proposed Tasks" Then Exit For
        Next
        With .Tables(i)
            .Rows.Add
            i = .Rows.Count
            .Rows(i).Cells(1).Range.text = tbTitle.value
            .Rows(i).Cells(2).Range.text = cbWho.value
            .Rows(i).Cells(3).Range.text = cbPriority.value
            .Rows(i).Cells(4).Range.text = Format(tbdDueDate.value, DateFormat)
        End With
        Protect 'wdAllowOnlyReading, False, "", False, True
    End With
End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub
Sub Display(Optional UploadMode As Boolean): mUploadMode = UploadMode: Me.Show: End Sub

'Private Sub tbDetails_Change(): CheckOk: End Sub
'Private Sub tbTitle_Change(): CheckOk: End Sub
'Sub CheckOk()
'    Dim i As Long, OkFlag As Boolean
'    OkFlag = True
'    For i = 1 To RequiredFields.Count
'        OkFlag = OkFlag And Len(RequiredFields(i).Value) > 0
'    Next
'    btn_OK.Enabled = OkFlag
'End Sub
'Private Sub RequiredField(CtrlName As String)
'    Dim Ctrl As Control
'    Set Ctrl = Me.Controls("lb" & CtrlName)
'    With Controls.Add("Forms.Label.1", "req" & CtrlName, True)
'        .Font.Size = 10
'        .Caption = "*"
'        .ForeColor = vbRed
'        .AutoSize = True
'        .Top = Ctrl.Top
'        .Left = Ctrl.Left
'        Ctrl.Left = Ctrl.Left + .Width
'    End With
'    On Error Resume Next
'    Set Ctrl = Nothing
'    Set Ctrl = Me.Controls("tb" & CtrlName)
'    If Ctrl Is Nothing Then Set Ctrl = Me.Controls("cb" & CtrlName)
'    If Ctrl Is Nothing Then Set Ctrl = Me.Controls("sb" & CtrlName)
'    RequiredFields.Add Ctrl, CtrlName
'End Sub
'Private Sub AddRequiredField(CtrlName As String)
'    Dim RQ As New RequiredFieldsEvents
'    With RQ
'        Set .Parent = Me
'        .AddOkBtn btn_OK
'        Set .RequiredFields = RequiredFields
'        .Add CtrlName
'    End With
'    RequiredFields.Add RQ
'End Sub
Private Sub UserForm_Initialize()
    Dim i As Long, BoardMembers As Dictionary
    Dim DocStates As Dictionary
    Set DocStates = GetStatesOfDoc("action_items")
    For i = 1 To DocStates.Count
        cbState.AddItem DocStates(i)
    Next
    
    Set Evs.Parent = Me
'    Evs.CollectAllControls
    Evs.AddOkButton btn_OK
'    Evs.AddOkButton btnPublish
    Evs.MakeRequired "Title,Details,Priority,Who,DueDate,cbState", , ErrorColor
'    Evs.MakeRequired "tbStarts,tbEnds", "*", ErrorColor
'    Me.Caption = "Add a new " & ProjectInfo("very_short_name") & " Meeting"

'    AddRequiredField "Title"
'    AddRequiredField "Details"
'    AddRequiredField "Priority"
'    AddRequiredField "Who"
'    Me.Caption = "Add a new " & ProjectInfo("very_short_name") & " Task"
    
'    With DatePicker
'        Set .tbDate = tbDueDate
'        Set .Parent = Me
'    End With
    On Error GoTo ex
    PArr = GetTaskPriorities
    If IsGoodResponse(PArr(1, 1)) Then
        For i = LBound(PArr, 2) To UBound(PArr, 2)
            cbPriority.AddItem PArr(1, i)
            PColl.Add PArr(2, i), PArr(1, i)
        Next
    End If
    PColl.Add "", ""
    Set BoardMembers = GetMembersOf
    For i = 1 To BoardMembers.Count
        cbWho.AddItem BoardMembers(i) '("fullname")
        liOthers.AddItem BoardMembers(i)
    Next
'    liOthers.List = cbWho.List
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
    lbPrjHeader.BackColor = ProjectColorStr
ex:
End Sub
