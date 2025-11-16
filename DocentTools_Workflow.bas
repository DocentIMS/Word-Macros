Attribute VB_Name = "DocentTools_Workflow"
Option Explicit
Private Const CurrentMod As String = "DocentTools_Workflow"
Function GetInitalState(Optional ByVal DocType As String) As String
    Dim TypeNum As Long
    On Error Resume Next
    If DocType = "" Then DocType = documentName(DocNum)
    TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    GetInitalState = GetStateName(CStr(WorkflowInfo(1)(TypeNum)("initial_state")), DocType, TypeNum)
End Function
Private Function GetTransitionNum(ByVal TransitionNameOrID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    For GetTransitionNum = 1 To WorkflowInfo(1)(TypeNum)("workflow_transitions").Count
        If WorkflowInfo(1)(TypeNum)("workflow_transitions")(GetTransitionNum)("id") = TransitionNameOrID Then Exit For
        If WorkflowInfo(1)(TypeNum)("workflow_transitions")(GetTransitionNum)("title") = TransitionNameOrID Then Exit For
    Next
    If GetTransitionNum > WorkflowInfo(1)(TypeNum)("workflow_transitions").Count Then GetTransitionNum = 0
End Function
Private Function GetStateNum(ByVal StateNameOrID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    For GetStateNum = 1 To WorkflowInfo(1)(TypeNum)("workflow_states").Count
        If WorkflowInfo(1)(TypeNum)("workflow_states")(GetStateNum)("id") = StateNameOrID Then Exit For
        If WorkflowInfo(1)(TypeNum)("workflow_states")(GetStateNum)("title") = StateNameOrID Then Exit For
    Next
End Function
Private Function GetTypeNum(Optional ByRef DocType As String) As Long
    If Len(DocType) = 0 Then DocType = documentName(DocNum)
    Dim mDocType As String
    mDocType = Replace(LCase(DocType), " ", "_")
    If WorkflowInfo.Count = 0 Then LoadProjectInfoReg
    For GetTypeNum = 1 To WorkflowInfo(1).Count
        If WorkflowInfo(1)(GetTypeNum)("content_type") = mDocType Then Exit For
    Next
    If GetTypeNum > WorkflowInfo(1).Count Then
        If DocType = "docent_misc_document" Then
            GetTypeNum = GetTypeNum("docent_misc_document")
        Else
            GetTypeNum = 0
        End If
    End If
End Function
Function GetStatesOfDoc(Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As Dictionary
    Dim StateNum As Long
    If Len(DocType) = 0 Then DocType = documentName(DocNum)
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    Set GetStatesOfDoc = New Dictionary
    For StateNum = 1 To WorkflowInfo(1)(TypeNum)("workflow_states").Count
        GetStatesOfDoc.Add , WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("title")
    Next
End Function
Function GetStateTransitions(Optional ByVal StateNameOrID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As Dictionary
    Dim TransitionNum As Long, StateNum As Long
    If Len(DocType) = 0 Then DocType = documentName(DocNum)
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    StateNum = GetStateNum(StateNameOrID, DocType, TypeNum)
    Set GetStateTransitions = New Dictionary
    For TransitionNum = 1 To WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("transitions").Count
        GetStateTransitions.Add , WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("transitions")(TransitionNum)
    Next
End Function
Function GetTransitionDestination(ByVal transitionID As String, ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim TransitionNum As Long, NewStateID As String
    If InStr(transitionID, "@workflow") Then transitionID = GetFileName(transitionID)
    'If Len(DocType) = 0 Then DocType = DocumentName(DocNum)
    TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    For TransitionNum = 1 To WorkflowInfo(1)(TypeNum)("workflow_transitions").Count
        If WorkflowInfo(1)(TypeNum)("workflow_transitions")(TransitionNum)("id") = transitionID Then
            GetTransitionDestination = WorkflowInfo(1)(TypeNum)("workflow_transitions")(TransitionNum)("new_state_id")
            GetTransitionDestination = GetStateName(GetTransitionDestination, DocType, TypeNum)
            Exit For
        End If
    Next
End Function
Function GetTransitionIdByStates(ByVal ThisState As String, ByVal TargetState As String, ByVal DocType As String) As String
    Dim TypeNum As Long, i As Long, TxColl As Dictionary
    TypeNum = GetTypeNum(DocType)
    Set TxColl = GetStateTransitions(ThisState, DocType, TypeNum)
    For i = 1 To TxColl.Count
        If GetTransitionDestination(TxColl(i), DocType, TypeNum) = TargetState Then
            GetTransitionIdByStates = TxColl(i)
            Exit Function
        End If
    Next
End Function
Function GetStateName(ByVal StateID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim StateNum As Long
    If Len(StateID) = 0 Then Exit Function 'Stay on state
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    StateNum = GetStateNum(StateID, DocType, TypeNum)
    GetStateName = WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("title")
End Function
Function GetStateID(ByVal StateName As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim StateNum As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    StateNum = GetStateNum(StateName, DocType, TypeNum)
    GetStateID = WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("id")
End Function
Function GetTransitionName(ByVal transitionID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim TransitionNum As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    TransitionNum = GetTransitionNum(transitionID, DocType, TypeNum)
    GetTransitionName = WorkflowInfo(1)(TypeNum)("workflow_transitions")(TransitionNum)("title")
End Function
Function GetTransitionID(ByVal TransitionName As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim TransitionNum As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    TransitionNum = GetTransitionNum(TransitionName, DocType, TypeNum)
'    If TransitionNum = 0 Then
'
'    End If
    GetTransitionID = WorkflowInfo(1)(TypeNum)("workflow_transitions")(TransitionNum)("id")
End Function
Function GetStateDescription(ByVal StateNameOrID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim StateNum As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    StateNum = GetStateNum(StateNameOrID, DocType, TypeNum)
    GetStateDescription = WorkflowInfo(1)(TypeNum)("workflow_states")(StateNum)("description")
End Function
Function GetTransitionDescription(ByVal TransitionNameOrID As String, Optional ByVal DocType As String, Optional ByVal TypeNum As Long) As String
    Dim TransitionNum As Long
    If TypeNum = 0 Then TypeNum = GetTypeNum(DocType)
    If TypeNum = 0 Then Exit Function
    TransitionNum = GetTransitionNum(TransitionNameOrID, DocType, TypeNum)
    GetTransitionDescription = WorkflowInfo(1)(TypeNum)("workflow_transitions")(TransitionNum)("description")
End Function
Function GetInitialTransitions(ByVal DocType As String) As Collection
    Dim InitState As String, TransitionInfo As Dictionary, i As Long
    Dim NextTransitions As Dictionary
    InitState = GetInitalState(DocType)
    Set GetInitialTransitions = New Collection
    Set NextTransitions = GetStateTransitions(InitState, DocType)
    For i = 1 To NextTransitions.Count
        Set TransitionInfo = New Dictionary
        TransitionInfo.Add "@id", "FILEURL/@workflow/" & NextTransitions(i)
        TransitionInfo.Add "title", GetTransitionName(NextTransitions(i), DocType)
        GetInitialTransitions.Add TransitionInfo
    Next
End Function
