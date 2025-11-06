Attribute VB_Name = "AF_Tasks_mod"
Option Explicit
Sub AddTask()
    frmCreateTask.Show
End Sub
'Sub UpdateTasks(TasksStr As String)
'
'End Sub
Sub UploadTasks(Optional Planned As Boolean = True)
    Dim s As String, ss() As String, sx() As String, i As Long, j As Long, AllGood As Boolean
    Dim MtngType As String, Resp As WebResponse
    MtngType = GetProperty(pMeetingType)
    s = GetProperty(IIf(Planned, pPlannedTasks, pProposedTasks))
    ss = Split(s, ";")
    AllGood = True
    For i = 1 To UBound(ss)
        sx = Split(ss(i), ",")
        Set Resp = CreateAPITask(sx(0), sx(1), sx(2), sx(3), sx(4), sx(5), MtngType)
        AllGood = AllGood And IsGoodResponse(Resp)
    Next
    If AllGood Then
        frmMsgBox.Display Array("A new Task was created on " & ProjectNameStr & " site.", " ", , "View Online"), , Success, "DocentIMS", , , Array(, , , Resp.Data("@id"))
    Else
        frmMsgBox.Display "Tasks could not be created", , Critical
    End If
    SetProperty IIf(Planned, "Planned", "Proposed") & "Tasks", ""
End Sub
Sub Tasks()
    Dim TaskColl As Collection, i As Long
    Dim Coll As New Dictionary
    Set TaskColl = GetAPIFolder("action-items", "action_items", Array("is_this_item_closed", _
                                                                "priority", _
                                                                "revised_due_date", _
                                                                "duedate", _
                                                                "assigned_to", _
                                                                "id"), _
                                                                Array("review_state"), Array(GetStateID("Published", "action_items")))
    'In future, if we have too many tasks, we can limit the search by assigned_id, or assigned_to
    'assigned_id: use the user id
    'assigned_to: use the user name
    'We can alo limit the search by priority., example>2
    'Or priorityString "two" [Don't use]
    If IsGoodResponse(TaskColl) Then
        For i = 1 To TaskColl.Count
            Coll.Add , Array(GetID(CStr(TaskColl(i)("id"))), _
                            NoNull(TaskColl(i)("title")), _
                            NoNull(TaskColl(i)("assigned_to")), _
                            NoNull(TaskColl(i)("priority")), _
                            GetDueDate(TaskColl(i)("duedate"), TaskColl(i)("revised_due_date")), _
                            NoNull(TaskColl(i)("is_this_item_closed")), _
                            TaskColl(i)("@id"))
        Next
    End If
    Set frmListTasks.ItemsDict = Coll
    'frmListTasks.PopList
    frmListTasks.Show
End Sub
Private Function NoNull(s As Variant) As String: On Error Resume Next: NoNull = s: End Function
Private Function GetID(URL As String) As String
    Dim i As Long
    For i = Len(URL) To 1 Step -1
        If Not IsNumeric(Mid(URL, i, 1)) Then
            i = i + 1
            Exit For
        End If
    Next
    GetID = IIf(i = 0 Or i > Len(URL), "0", Mid(URL, i))
End Function
Private Function GetDueDate(DueDate, RevisedDueDate) As String
    DueDate = NoNull(DueDate)
    RevisedDueDate = NoNull(RevisedDueDate)
    GetDueDate = Format(IIf(Len(RevisedDueDate), RevisedDueDate, DueDate), DateFormat)
End Function

