Attribute VB_Name = "DocentTools_Properties"
Option Explicit
Private Const CurrentMod = "DocentTools_Properties"
'Documents Properties
Public Enum DocProperty
    pDocLastSave = -2 ' "Last Save Time"
    pAuthor = -1 ' "Last Author"
    pIsDocument = 0 '"isDocentDocument" [3] True
    pPName = 1 ' "docentProject"
    pPURL = 2 '"docentProjectURL"
    pDocType = 3 ' "docentDocType"
    pDocVer = 4 ' "docentVersion"
    pDocDate = 5 ' "docentDocDate"
    pDocState = 6 '"docentDocState"
    pDocURL = 7 '"docentDocURL" [2] Just the folder URL, not the document itsef as it is not there yet
    pDocCreateDate = 8 '"docentCreationDate"
    pPublishDate = 9 ' "docentPublishingDate"
    pProposedTasks = 10 '"ProposedTasks
    pPlannedTasks = 11 '"PlannedTasks
    pContractNo = 12 '"docentContractNo
    pMeetingType = 13 '"docentMeetingType"
    pActuals = 14 '"docentMeetingActuals"
    pMeetingUID = 15 '"docentMeetingUID" [1]
    pOnlineMeetingUID = 16 '"onlineMeetingUID"
    pIsFinalRev = 17 '"docentIsFinalRev"
    pIsTemplate = 18 ' "isDocentTemplate" [4] False
    pTemplateVer = 19 ' "docentTemplateVersion"
    pTemplateDate = 20 ' "docentTemplateDate"
End Enum
Private Function GetPropertyStr(PropertyNo As DocProperty) As String
    Select Case PropertyNo
    Case pIsDocument: GetPropertyStr = "isDocentDocument"
    Case pPName: GetPropertyStr = "docentProject"
    Case pPURL: GetPropertyStr = "docentProjectURL"
    Case pDocType: GetPropertyStr = "docentDocType"
    Case pDocVer: GetPropertyStr = "docentVersion"
    Case pDocDate: GetPropertyStr = "docentDocDate"
    Case pDocState: GetPropertyStr = "docentDocState"
    Case pDocURL: GetPropertyStr = "docentDocURL"
    Case pDocLastSave: GetPropertyStr = "Last Save Time"
    Case pIsTemplate: GetPropertyStr = "isDocentTemplate"
    Case pTemplateVer: GetPropertyStr = "docentTemplateVersion"
    Case pTemplateDate: GetPropertyStr = "docentTemplateDate"
    Case pAuthor: GetPropertyStr = "Last Author"
    Case pPublishDate: GetPropertyStr = "docentPublishingDate"
    Case pProposedTasks: GetPropertyStr = "ProposedTasks"
    Case pPlannedTasks: GetPropertyStr = "PlannedTasks"
    Case pDocCreateDate: GetPropertyStr = "docentCreationDate"
    Case pContractNo: GetPropertyStr = "docentContractNo"
    Case pMeetingType: GetPropertyStr = "docentMeetingType"
    Case pActuals: GetPropertyStr = "docentMeetingActuals"
    Case pMeetingUID: GetPropertyStr = "docentMeetingUID"
    Case pOnlineMeetingUID: GetPropertyStr = "onlineMeetingUID"
    Case pIsFinalRev: GetPropertyStr = "docentIsFinalRev"
    End Select
End Function
Function GetProperty(PropertyNo As DocProperty, Optional Doc As Document)
    On Error Resume Next
    If Documents.Count = 0 Then Exit Function
    If Doc Is Nothing Then Set Doc = ActiveDocument
    If Doc Is Nothing Then Exit Function
    Dim i As Long
    i = DocsInfo.Count
    If Err.Number Then
        Set DocsInfo = New Collection
        Err.Clear
    End If
    i = DocsInfo(Doc.Name).Count
    If Err.Number Then
        DocsInfo.Remove Doc.Name
        DocsInfo.Add LoadDocInfo(Doc), Doc.Name
    End If
    GetProperty = DocsInfo(Doc.Name)(GetPropertyStr(PropertyNo))
'    GetProperty = GetDocProperty(GetPropertyStr(PropertyNo), Doc)
End Function
Sub SetProperty(PropertyNo As DocProperty, ByVal PPValue As Variant, Optional Doc As Document, Optional ByVal PPType As Office.MsoDocProperties = 4)
    SetDocProperty GetPropertyStr(PropertyNo), PPValue, Doc, PPType
End Sub
Sub DelProperty(PropertyNo As DocProperty, Optional Doc As Document)
    DeleteDocProperty GetPropertyStr(PropertyNo), Doc
End Sub
Function LoadDocInfo(Doc As Document) As Collection
    WriteLog 1, CurrentMod, "LoadDocInfo", Doc.Name
    Dim i As Long, Coll As New Collection 'IsFlag As Boolean,
    Set LoadDocInfo = New Collection
    On Error Resume Next
    Coll.Add Doc.Name, "Name"
    For i = -2 To 0
        Coll.Add GetCustomDocProperty(GetPropertyStr(i), Doc), GetPropertyStr(i)
    Next
'    LoadDocInfo.Add Doc.BuiltInDocumentProperties(GetPropertyStr(pDocLastSave)).Value, "Last Save Time"
'    LoadDocInfo.Add Doc.BuiltInDocumentProperties(GetPropertyStr(pAuthor)).Value, "Last Author"
'    IsFlag = False
'    IsFlag = Doc.CustomDocumentProperties(GetPropertyStr(pIsDocument)).Value = True
'    LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pPName)).Value, "docentProject"
'    LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocType)).Value, "docentDocType"
'    If IsFlag Then
    If Coll("isDocentDocument") Then
        For i = 1 To 17
            Coll.Add GetCustomDocProperty(GetPropertyStr(i), Doc), GetPropertyStr(i)
        Next
'        LoadDocInfo.Add True, "isDocentDocument"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocVer)).Value, "docentVersion"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocDate)).Value, "docentDocDate"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocState)).Value, "docentDocState"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocURL)).Value, "docentDocURL"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pDocCreateDate)).Value, "docentCreationDate"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pPublishDate)).Value, "docentPublishingDate"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pProposedTasks)).Value, "ProposedTasks"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pPlannedTasks)).Value, "PlannedTasks"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pContractNo)).Value, "docentContractNo"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pMeetingType)).Value, "docentMeetingType"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pActuals)).Value, "docentMeetingActuals"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pMeetingUID)).Value, "docentMeetingUID"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pIsFinalRev)).Value, "docentIsFinalRev"
    End If
    For i = 18 To 20
        Coll.Add GetCustomDocProperty(GetPropertyStr(i), Doc), GetPropertyStr(i)
    Next
'    IsFlag = False
'    IsFlag = Doc.CustomDocumentProperties(GetPropertyStr(pIsTemplate)).Value = True
''    If IsFlag Then
'        LoadDocInfo.Add IsFlag, "isDocentTemplate"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pTemplateVer)).Value, "docentTemplateVersion"
'        LoadDocInfo.Add Doc.CustomDocumentProperties(GetPropertyStr(pTemplateDate)).Value, "docentTemplateDate"
''    End If
    On Error Resume Next
    Set LoadDocInfo = Coll
    DocsInfo.Remove Doc.Name
    DocsInfo.Add LoadDocInfo, Doc.Name
End Function

'=======================================================
' HELPER FUNCTIONS FOR PROPERTY ACCESS
'=======================================================

'=======================================================
' Function: GetPropertySafe
' Purpose: Safely retrieve document property as string
'
' Parameters:
'   propertyName - Name of property to retrieve
'
' Returns:
'   String - Property value or "Unknown" on error
'=======================================================
Function GetPropertySafe(ByVal propertyName As DocProperty, _
                                Optional ByVal defaultValue As String = "Unknown") As String
    On Error Resume Next
    GetPropertySafe = CStr(GetProperty(propertyName))
    If Err.Number <> 0 Or Len(GetPropertySafe) = 0 Then
        GetPropertySafe = defaultValue
    End If
    On Error GoTo 0
End Function

'=======================================================
' Function: GetPropertySafeBool
' Purpose: Safely retrieve document property as boolean
'
' Parameters:
'   propertyName - Name of property to retrieve
'   defaultValue - Default value if property not found
'
' Returns:
'   Boolean - Property value or default value on error
'=======================================================
Function GetPropertySafeBool(ByVal propertyName As DocProperty, _
                                     Optional ByVal defaultValue As Boolean) As Boolean
    On Error Resume Next
    GetPropertySafeBool = CBool(GetProperty(propertyName))
    If Err.Number <> 0 Then
        GetPropertySafeBool = defaultValue
    End If
    On Error GoTo 0
End Function

'=======================================================
' Sub: AppendPropertyIfExists
' Purpose: Append property to message if it exists
'
' Parameters:
'   Msg - Message string to append to (ByRef)
'   label - Display label for property
'   propertyName - Name of property to retrieve
'=======================================================
Sub AppendPropertyIfExists(ByRef msg As String, _
                                   ByVal label As String, _
                                   ByVal propertyName As DocProperty)
    Dim value As String
    
    On Error Resume Next
    value = CStr(GetProperty(propertyName))
    If Err.Number = 0 And Len(value) > 0 Then
        msg = msg & "   " & label & ": " & value & vbLf
    End If
    On Error GoTo 0
End Sub

'=======================================================
' Sub: AppendPropertyDateIfExists
' Purpose: Append formatted date property to message if it exists
'
' Parameters:
'   Msg - Message string to append to (ByRef)
'   label - Display label for property
'   propertyName - Name of property to retrieve
'=======================================================
Sub AppendPropertyDateIfExists(ByRef msg As String, _
                                       ByVal label As String, _
                                       ByVal propertyName As DocProperty)
    Dim value As String
    
    On Error Resume Next
    value = CStr(GetProperty(propertyName))
    If Err.Number = 0 And Len(value) > 0 Then
        On Error Resume Next
        msg = msg & "   " & label & ": " & Format$(value, DateTimeFormat) & vbLf
        On Error GoTo 0
    End If
    On Error GoTo 0
End Sub

