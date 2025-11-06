Attribute VB_Name = "Ribbon_Functions_Mod"
Option Explicit
Option Compare Text
Option Private Module
Private Const CurrentMod = "AC_Ribbon_Functions_Mod" '"" '

Private ErrShown As Boolean
Private cProjectName As Variant
Private cProjectURL As Variant
Private cDocumentName As Variant
Private cTemplateName As Variant
Private RefreshFlag As Boolean

'===========
'Functions
'===========
Function GetSelectedProjectColor()
'    WriteLog 1, CurrentMod, "GetSelectedProjectColor"
'    On Error Resume Next
'    ProjectSelected ActiveDocument, PNum
'    On Error GoTo 0
'    Set GetSelectedProjectColor = LoadPictureGDI(ProjectColorImageStr)
End Function
Function GetButtonVisible(Mode As Long) As Boolean
    'Mode:
    '0 = Project Selected,
    '1 = Default,
    '2 = Documents,
    '3 = Scope,
    '4 = Tempalte,
    '5 = Planning
    '6 = MSP
    '7 = RFP
    '8 = Unpublished Scope
    GetButtonVisible = False
    On Error GoTo ex
'    WriteLog 1, CurrentMod, "GetButtonVisible", Mode
    Select Case Mode
    Case 0 'Default Project Selected
        If Documents.Count = 0 Then
            GetButtonVisible = IsProjectSelected(True)
        ElseIf GetProperty(pIsDocument) Then
            GetButtonVisible = False
        ElseIf GetProperty(pIsTemplate) Then
            GetButtonVisible = False
        Else
            GetButtonVisible = IsProjectSelected(True) 'GetProperty(pIsTemplate)
        End If
    Case 1 'Default
        If Documents.Count = 0 Then
            GetButtonVisible = True
        ElseIf GetProperty(pIsDocument) Then
            GetButtonVisible = False
        ElseIf GetProperty(pIsTemplate) Then
            GetButtonVisible = False
        Else
            GetButtonVisible = True 'GetProperty(pIsTemplate)
        End If
    Case 2 'Documents
'        If Documents.Count = 0 Then
'            GetButtonVisible = True
'        Else
        If GetProperty(pIsDocument) Then
            GetButtonVisible = Not GetProperty(pIsTemplate) And Not GetProperty(pDocType) = "MS Project"
        End If
    Case 3 'Scope
'        If Documents.Count = 0 Then
'            GetButtonVisible = True
'        Else
        Select Case GetProperty(pDocType)
        Case "Scope", "Scope Document"
            GetButtonVisible = Not GetProperty(pIsTemplate)  'GetProperty(pDocType) = "Scope"
        End Select
'        End If
    Case 4 'Tempalte
'        If Documents.Count = 0 Then
'            GetButtonVisible = True
'        Else
            GetButtonVisible = GetProperty(pIsTemplate)
'        End If
    Case 5 'Planning
        GetButtonVisible = Not GetProperty(pIsTemplate) And GetProperty(pDocType) = "Planning Document"
    Case 6 'MSP
        GetButtonVisible = Not GetProperty(pIsTemplate) And GetProperty(pDocType) = "MS Project"
    Case 7 'RFP
        GetButtonVisible = Not GetProperty(pIsTemplate) And GetProperty(pDocType) = "RFP"
    Case 8
        Select Case GetProperty(pDocType)
        Case "Scope", "Scope Document"
            GetButtonVisible = Not GetProperty(pIsTemplate) And Not GetProperty(pIsFinalRev) 'GetProperty(pDocType) = "Scope"
        End Select
    End Select
ex:
'    ''Debug.Print "GetButtonVisible " & Mode & " : " & GetButtonVisible
End Function
Private Function GetGroupVisibility(IsPrjMgr As Boolean, Tgl As Boolean, IsMgrGroup As Boolean, Optional CanDoGroup As String) As Boolean
    Dim CanDo As Boolean
    If Len(ProjectURLStr) = 0 Then Exit Function
    If IsMgrGroup Then If Not GetVisibleGroup("IdToggleButtonMgrMode") Then Exit Function
    CanDo = UserGroupsDict.Exists(CanDoGroup)
    If IsPrjMgr Then
        If IsMgrGroup Then
            GetGroupVisibility = Tgl Or GetProperty(pIsTemplate)
        Else
            GetGroupVisibility = Not Tgl Or (GetProperty(pIsDocument) And Not GetProperty(pIsTemplate))
        End If
    Else
        If IsMgrGroup Then
            GetGroupVisibility = CanDo And Not Tgl
        'ElseIf IsPrjMgr = 2 Then
            
        Else
            GetGroupVisibility = Tgl 'True
        End If
    End If
End Function
Function GetVisibleGroup(GroupID As String) As Boolean
    Dim i As Long, IsPrjMgr As Boolean
    GetVisibleGroup = False
'    If Len(ProjectColorStr) = 0 Then Exit Function
    On Error GoTo ex
    If UserGroupsDict Is Nothing Then GetMyGroupsDict
    IsPrjMgr = UserGroupsDict.Exists("PrjMgr")


'    WriteLog 1, CurrentMod, "getVisibleGroup", GroupId
    Dim Mode As Long
    If MainInfo Is Nothing Then
        RefreshRibbonGroups 'Invalidate "IdGroupParseScope"
        Exit Function
    End If
'    For i = 1 To GroupsDict.Count
'        Debug.Print GroupsDict(i)
'    Next
    Select Case GroupID
'    Case "IdGroupParseScope"
'        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_parse")
''        GetVisibleGroup = GroupsDict.Exists("can_parse") Xor PrjMgr
'        Mode = 1
    Case "IdGroupLinks"
'        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, False)
'        If GetVisibleGroup Then
        GetVisibleGroup = GetButtonVisible(0)
    Case "IdGroupScope"
'        If GetVisibleGroup("IdToggleButtonMgrMode") Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_project_scope")
'        ElseIf IsPrjMgr Then
        If Not GetVisibleGroup And IsPrjMgr Then
            GetVisibleGroup = GetButtonVisible(3)
        End If
    Case "IdGroupRFP"
'        If GetVisibleGroup("IdToggleButtonMgrMode") Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_rfp")
'        Else
        If Not GetVisibleGroup And IsPrjMgr Then
            GetVisibleGroup = GetButtonVisible(2)
        End If
'        GetVisibleGroup = GroupsDict.Exists("can_add_project_scope") Xor PrjMgr
'        Mode = 3
    Case "IdGroupPMP"
'        If Not GetProperty(pIsTemplate) Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_pmp")
    '        GetVisibleGroup = GroupsDict.Exists("can_parse") Xor PrjMgr
            Mode = 1
'        End If
    Case "IdGroupMSP"
'        If Not GetProperty(pIsTemplate) Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_msp")
            'If Not GetVisibleGroup Then
            If Not GetVisibleGroup Then GetVisibleGroup = GetButtonVisible(6)
'        End If
'        GetVisibleGroup = GroupsDict.Exists("can_parse") Xor PrjMgr
        'Mode = 1
    Case "IdGroupRFP"
        If Not GetProperty(pIsTemplate) Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_rfp")
            If Not GetVisibleGroup Then GetVisibleGroup = GetButtonVisible(7)
        End If
'        GetVisibleGroup = GroupsDict.Exists("can_parse") Xor PrjMgr
        'Mode = 1
    Case "IdGroupCreate", "IdGroupTasks", "IdGroupTeam"
        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, False)
'        GetVisibleGroup = GroupsDict.Count > 0 And Not PrjMgr
        Mode = 1
    Case "IdGroupNotifications"
'        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True)
'        GetVisibleGroup = GroupsDict.Count > 0 And Not PrjMgr
        GetVisibleGroup = Len(ProjectNameStr) > 0
        Mode = 1
'    Case "IdButtonCreateNotification"
'        GetVisibleGroup = PrjMgr
    Case "IdGroupDocument"
        GetVisibleGroup = GetVisibleGroup("IdGroupScope") And Not GetVisibleGroup("IdToggleButtonMgrMode")
        If Not GetVisibleGroup Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, False) ', "can_add_documents")
    '        GetVisibleGroup = GroupsDict.Exists("can_add_documents") And Not PrjMgr
            If GetVisibleGroup Then
                If GetProperty(pIsDocument) Then
                    GetVisibleGroup = Not GetProperty(pDocType) Like "Meeting *"
                    If GetVisibleGroup Then GetVisibleGroup = Not GetProperty(pIsTemplate)
                    If GetVisibleGroup Then
                        Select Case GetProperty(pDocType)
                        Case "Planning Document", "RFP", "Scope", "Scope Document": GetVisibleGroup = False
                        End Select
                    End If
                End If
            End If
        End If
    Case "IdGroupMeetingDoc"
        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, False) ', "can_add_documents")
'        GetVisibleGroup = GroupsDict.Exists("can_add_documents") And Not PrjMgr
        If GetVisibleGroup Then
            If GetProperty(pIsDocument) Then
                GetVisibleGroup = GetProperty(pDocType) Like "Meeting *"
                If GetVisibleGroup Then GetVisibleGroup = Not GetProperty(pIsTemplate)
                If GetVisibleGroup Then GetVisibleGroup = GetProperty(pDocType) <> "Planning Document"
                If GetVisibleGroup Then GetVisibleGroup = GetProperty(pDocType) <> "Scope Document"
                If GetVisibleGroup Then GetVisibleGroup = GetProperty(pDocType) <> "Scope"
            End If
        End If
    Case "IdGroupTemplate"
        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_modify_templates")
        GetVisibleGroup = GetVisibleGroup Or GetProperty(pIsTemplate)
''        GetVisibleGroup = Not (GroupsDict.Exists("can_modify_templates") Xor PrjMgr)
        If GetVisibleGroup Then
            If GetProperty(pIsDocument) Then GetVisibleGroup = GetProperty(pIsTemplate)
        End If
    Case "IdGroupCommandStatements"
        GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_command_statements")
'        GetVisibleGroup = GroupsDict.Exists("can_add_command_statements") Xor PrjMgr
        Mode = 1
    Case "IdGroupPlanning"
'        If GetVisibleGroup("IdToggleButtonMgrMode") Then
            GetVisibleGroup = GetGroupVisibility(IsPrjMgr, PrjMgr, True, "can_add_planning_document")
            If Not GetVisibleGroup Then GetVisibleGroup = GetButtonVisible(5)
'        End If
'        GetVisibleGroup = Not (GroupsDict.Exists("can_add_planning_document") Xor PrjMgr) 'And _
                            GetProperty(pDocType) = "Planning Document"
        'Mode = 5
    Case "IdCheckBoxPlanning"
        GetVisibleGroup = UBound(PlanningProjectName) > 0
'        For i = LBound(PlanningProjectName) To UBound(PlanningProjectName) 'And the user can add planning documents to at least one of the planning projects
'            GetVisibleGroup = ArrToDict(Split(UserGroups(GetProjectIndexByName(PlanningProjectName(i), ProjectName)), Seperator)).Exists("can_add_planning_document")
'            If GetVisibleGroup Then Exit For
'        Next
        If Not GetVisibleGroup Then PlanningOnly = False
    Case "IdToggleButtonMgrMode"
        GetVisibleGroup = IsPrjMgr
        If Not GetVisibleGroup Then
            i = i + 1
            Do Until IsPrjMgr Or i > UserGroupsDict.Count
                GetVisibleGroup = UserGroupsDict.KeyName(i) Like "can_*"
                i = i + 1
            Loop
        End If
        Mode = 1
    End Select
    If GetVisibleGroup And Mode > 0 Then GetVisibleGroup = GetButtonVisible(Mode)
'    GetVisibleGroup = True
ex:
End Function
Function GetProjectsCount() As Long
'    GetProjectsCount = 0
    WriteLog 1, CurrentMod, "GetProjectsCount"
    On Error GoTo ex
    If GetProperty(pIsDocument) Then
        ReDim cProjectName(0 To 0)
        ReDim cProjectURL(0 To 0)
        cProjectName(0) = GetProperty(pPName) 'DocumentName(DocNum)
        cProjectURL(0) = GetProperty(pPURL) 'DocumentName(DocNum)
        GetProjectsCount = 1
    Else
ex:
    On Error GoTo -1
        On Error Resume Next
        cProjectName = IIf(PlanningOnly, PlanningProjectName, NoPlanningProjectName)
        cProjectURL = IIf(PlanningOnly, PlanningProjectURL, NoPlanningProjectURL)
        If Not IsArrayAllocated(cProjectName) Then
            LoadProjects
            cProjectName = IIf(PlanningOnly, PlanningProjectName, NoPlanningProjectName)
            cProjectURL = IIf(PlanningOnly, PlanningProjectURL, NoPlanningProjectURL)
        End If
        GetProjectsCount = CheckReg(cProjectName, "Project") 'IIf(PlanningOnly, "PlanningProjectName", "ProjectName"))

    End If
    ShowHelp
'    GetProjectsCount = CheckReg(cProjectName, "Project")
'    GetProjectsCount = IIf(getproperty(pIsDocument), 1, CheckReg(ProjectName, "projects"))
End Function
Function CheckReg(Arr, ID As String) As Long 'ArrName As String) As Long
    Dim TriedOnce As Boolean
    On Error Resume Next
'    Dim Arr
    Do
'        If TriedOnce Then TriedOnce = True
        On Error Resume Next
    '    Select Case ArrName
    '    Case "PlanningProjectName": Arr = PlanningProjectName
    '    Case "ProjectName": Arr = ProjectName
    '    Case "DocumentName": Arr = DocumentName
    '    End Select
        CheckReg = UBound(Arr) + 1
        If Err.Number And Not TriedOnce Then
            LoadProjects
            TriedOnce = True
        Else
            Exit Do
        End If
    Loop Until CheckReg > 1
'    If (CheckReg = 1 Or Err.Number <> 0) And Not PlanningOnly Then BuildDocumentsArrays        'UpdateLocalProjectInfo
        
    If (CheckReg = 1 Or Err.Number <> 0) And Not PlanningOnly And InStr(Arr(0), "Select") = 0 Then
        'Stop
'        If Not ErrShown Then
'            If InStr(id, "project") Then 'If InStr(ArrName, "project") Then
'                WriteLog 3, CurrentMod, "CheckReg", "Registery info about Projects is missing" ' & ArrName & " is missing"
'                If frmMsgBox.Display("There are no Projects configured in Word." & vbNewLine & vbNewLine & _
'                        "Press ""Configure""  to setup projects, or, ""Cancel""" & vbNewLine & vbNewLine & _
'                        "Add projects anytime - click ""Project Configuration"" under ""Help""" _
'                        , Array("Configure", "Cancel"), Information, "Welcome to Docent IMS") = "Configure" Then
'                    frmProjectsList.Show
'                End If
'                ErrShown = True
'            ElseIf InStr(id, "document") Then 'ElseIf InStr(ArrName, "document") Then
'                'frmMsgBox.Display "There are no Document Types configured in your Docent account." & vbNewLine & vbNewLine & _
'                        "Please contact the Project Manager" _
'                        , Array("Cancel"), Exclamation, "No Document Types Found"
'            End If
'        End If
    End If
End Function
Private Function GetDocArr(Optional DocumentsMode As Long) As Variant
    Select Case DocumentsMode
    Case 0: GetDocArr = documentName
    Case 1: GetDocArr = MeetingDocName
    End Select
End Function
Function GetDocumentsCount(Optional DocumentsMode As Long) As Long
    cDocumentName = GetDocArr(DocumentsMode)
    If Not IsArrayAllocated(cDocumentName) Then
        LoadProjects
        cDocumentName = GetDocArr(DocumentsMode)
    End If
    GetDocumentsCount = UBound(cDocumentName) + 1
End Function
Function GetTemplatesCount() As Long
    cTemplateName = templateName
    If Not IsArrayAllocated(cTemplateName) Then
        LoadProjects
        cTemplateName = templateName
    End If
    GetTemplatesCount = UBound(cTemplateName) + 1 'CheckReg(cTemplateName, "Template") 'UBound(DocumentName) + 1
End Function
Function GetSelectedTemplateIndex() As Long
    Dim Index As Long, FName As String, DocName As String
    WriteLog 1, CurrentMod, "GetSelectedTemplateIndex"
    On Error Resume Next
    FName = GetActiveFName(ActiveDocument)
    If FName <> "" And RefreshFlag Then
        RefreshFlag = False
        Index = TemplateNum
'        DocName = GetProperty(pDocType)
'        If Len(DocName) Then
'            For Index = 1 To UBound(TemplateName)
'                If DocName = TemplateName(Index) Then Exit For
'            Next
'            If Index > UBound(TemplateName) Then Index = 0
'        End If
    End If
    On Error GoTo ex
ex:
    TemplateNum = Index
    GetSelectedTemplateIndex = Index
End Function
Sub SetSelectedTemplateIndex(Index As Integer)
    WriteLog 1, CurrentMod, "SetSelectedTemplateIndex"
    Dim FName As String
    On Error Resume Next
'    FName = GetFileName(GetActiveFName)
'    If Len(FName) > 0 Then
'        SetRegLastSelectedP FName & "_docentDocType", DocumentName(Index)
'        SetProperty "docentDocType", DocumentName(Index)
'    End If
    TemplateNum = Index
    PrintView
    RefreshFlag = True
    RefreshTemplatesGroup
'    RefreshDocumentGroup
'    On Error GoTo 0
'    DoEvents
End Sub
Function GetSelectedProjectIndex() As Long
    'stop'test
    Dim Index As Long, FName As String, PURL As String, Doc As Document
    WriteLog 1, CurrentMod, "GetSelectedProjectIndex"
    If NotScopeAsked Is Nothing Then Set NotScopeAsked = New Collection
'        Debug.Print SelectedProject
    On Error Resume Next
    AskForMissingPw = False
    Set Doc = ActiveDocument
    FName = GetActiveFName(Doc)
    On Error GoTo ex
'    If FName = "Nothing" Then
'        GoTo ex
'    Else
    If GetProperty(pIsDocument) Then
        'If FName = "Blank" Then FName = GenDocName("")
'        RefreshFlag = False
        PURL = GetProperty(pPURL)
        PNum = GetProjectIndexByURL(PURL, projectURL)
        SetRegSelection FName, PURL, selectedProject
        ProjectSelected Doc, PNum, True
        Index = 0
    Else
ex:
'        Debug.Print FName
        On Error GoTo -1
        On Error Resume Next
'        Select Case FName
'        Case "Blank", "Nothing"
''        Case "Blank" ', "Nothing"
''            SetRegLastSelectedP FName & "_ProjectName", ProjectName(0)
''            RefreshFlag = False
''        Case "Nothing"
''            RefreshFlag = False
'        Case Else
            RefreshFlag = Len(FName) > 0 'FName <> "Blank" '
'        End Select
'        Debug.Print FName
        'Stop
            Index = GetProjectIndexByURL(GetRegSelection(FName, selectedProject))
        If RefreshFlag Then  '
            RefreshFlag = False
            NotScopeAsked.Remove FName
            NotScopeAsked.Add False, FName
'            Invalidate "IdDDProject"
'            DoEvents
        End If
        ProjectSelected Doc, GetProjectIndexByIndex(Index) ', True
'        Invalidate "IdDDProject"
'        If Len(FName) > 0 Then
'    On Error Resume Next
'        Select Case FName
'        Case "Blank", "Nothing": SetRegSelection FName, ProjectName(0), SelectedProject
'        End Select
    End If
'    RefreshRibbon
'    RefreshProject 'X
'    RefreshRibbonGroups
    GetSelectedProjectIndex = Index
'    Sleep 200
    'stop'test
End Function
Function GetProjectNameByIndex(Optional Index) As String
    On Error Resume Next
    GetProjectNameByIndex = projectName(GetProjectIndexByIndex(Index))
    If GetProjectNameByIndex Like "Select *" Then GetProjectNameByIndex = ""
End Function
Function GetProjectURLByIndex(Optional Index) As String
    On Error Resume Next
    GetProjectURLByIndex = cProjectURL(Index)
    If GetProjectURLByIndex Like "Select *" Then GetProjectURLByIndex = ""
End Function
Function GetProjectIndexByIndex(Optional Index) As Long
    Dim PName As String
    PName = GetProjectURLByIndex(Index)
    On Error Resume Next
    For GetProjectIndexByIndex = 1 To UBound(projectURL)
        If PName = projectURL(GetProjectIndexByIndex) Then Exit For
    Next
    If GetProjectIndexByIndex > UBound(projectURL) Then GetProjectIndexByIndex = 0
End Function
Function GetProjectIndexByName(Optional PName As String, Optional mProjectName) As Long
    On Error GoTo ex
    If IsMissing(mProjectName) Then mProjectName = cProjectName
    If Len(PName) = 0 Then PName = GetProperty(pPName)
    If Len(PName) = 0 Then Exit Function
    For GetProjectIndexByName = 0 To UBound(mProjectName)
        If PName = mProjectName(GetProjectIndexByName) Then Exit For
    Next
    If GetProjectIndexByName > UBound(mProjectName) Then GetProjectIndexByName = 0
ex:
End Function
Function GetProjectIndexByURL(Optional PURL As String, Optional mProjectURL) As Long
    On Error GoTo ex
    If IsMissing(mProjectURL) Then mProjectURL = cProjectURL
    If Len(PURL) = 0 Then PURL = GetProperty(pPURL)
    If Len(PURL) = 0 Then Exit Function
    For GetProjectIndexByURL = 0 To UBound(mProjectURL)
        If PURL = mProjectURL(GetProjectIndexByURL) Then Exit For
    Next
    If GetProjectIndexByURL > UBound(mProjectURL) Then GetProjectIndexByURL = 0
ex:
End Function
Sub SetSelectedProjectIndex(Index)
    'stop'test
    WriteLog 1, CurrentMod, "SetSelectedProjectIndex"
    Dim FName As String, Doc As Document, PURL As String ', PName As String
    On Error Resume Next
    Set Doc = ActiveDocument
    AskForMissingPw = True
    FName = GetFileName(GetActiveFName(Doc))
    'PName = cProjectName(Index)
    PURL = cProjectURL(Index)
    If Len(FName) > 0 Then
'        SetProperty "docentProject", cProjectName(index)
        SetRegSelection FName, PURL, selectedProject
    End If
    PrintView
'    On Error GoTo 0
    RefreshFlag = True
    NewPNum = Index
'    ProjectSelected Doc, GetProjectIndexByName(PName) 'CLng(Index)
    RefreshProject  'Ribbon ' 'Invalidate "IdDDProject"
    DoEvents
    'stop'test
End Sub
Sub PrintView()
    WriteLog 1, CurrentMod, "PrintView"
    If ActiveWindow.View.SplitSpecial = wdPaneNone Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.View.Type = wdPrintView
    End If
End Sub
Function IsProjectSelected(Optional SilentMode As Boolean) As Boolean
    WriteLog 1, CurrentMod, "IsProjectSelected"
    On Error GoTo ex
    IsProjectSelected = Not (Len(ProjectNameStr) = 0 Or ProjectNameStr Like "Select *") 'Len(ProjectColorStr) <> 0
'    If Not IsProjectSelected Then 'Len(ProjectNameStr) = 0 Or ProjectNameStr Like "Select *" Then 'And UBound(URL) > 0 Then
'        RefreshRibbon
''        RedefineRibbon
'        ProjectSelected ActiveDocument, GetProjectIndexByName(GetProperty(pPName), ProjectName), True 'GetProperty(pPName) '
''        ProjectSelected ActiveDocument, GetProjectIndexByName(GetRegLastSelectedP(GetFileName(GetActiveFName(ActiveDocument)) & "_ProjectName"), ProjectName), True 'GetProperty(pPName) '
''        RefreshDocumentsGroups
'    End If
'    IsProjectSelected = Not (Len(ProjectNameStr) = 0 Or ProjectNameStr Like "Select *") 'Len(ProjectColorStr) <> 0
    If Not IsProjectSelected Then
        WriteLog 2, CurrentMod, "IsProjectSelected", "No project is selected"
        If Not SilentMode Then MsgBox "Please select a project first", vbExclamation, "No project is selected"
    End If
ex:
End Function
Function OpenSelectedTemplate()
    OpenTemplate CStr(cTemplateName(TemplateNum)), True
End Function
'-----MeetingDoc---------
Function GetSelectedMeetingDocIndex() As Long
    Dim Index As Long, FName As String, DocName As String
    WriteLog 1, CurrentMod, "GetSelectedMeetingDocIndex"
    On Error Resume Next
    FName = GetActiveFName(ActiveDocument)
    If FName <> "" And RefreshFlag Then
        RefreshFlag = False
        DocName = GetProperty(pDocType)
        If Len(DocName) Then
            For Index = 1 To UBound(MeetingDocName)
                If DocName = MeetingDocName(Index) Then Exit For
            Next
            If Index > UBound(MeetingDocName) Then Index = 0
        End If
        If Index = 0 Then Index = GetProjectIndexByURL(GetRegSelection(FName, SelectedMeetingDocument))
    End If
    'On Error GoTo ex
ex:
    MeetingDocumentSelected CLng(Index)
'    RefreshMeetingDocButtons
    GetSelectedMeetingDocIndex = Index
End Function
Sub SetSelectedMeetingDocIndex(Index As Integer)
    WriteLog 1, CurrentMod, "SetSelectedMeetingDocIndex"
    Dim FName As String, MDocName As String
    On Error Resume Next
    FName = GetFileName(GetActiveFName(ActiveDocument))
    MDocName = MeetingDocName(Index)
    If Len(FName) > 0 Then
        SetRegSelection FName, MDocName, SelectedMeetingDocument
'        SetProperty "docentDocType", MeetingDocName(Index)
    End If
    MeetingDocNum = Index
    PrintView
    RefreshFlag = True
    RefreshMeetingDocButtons
'    On Error GoTo 0
'    DoEvents
End Sub
'------Documents--------
Function GetSelectedDocumentIndex() As Long
    Dim Index As Long, FName As String, DocName As String
    WriteLog 1, CurrentMod, "GetSelectedDocumentIndex"
    On Error Resume Next
    FName = GetActiveFName(ActiveDocument)
    If FName <> "" And RefreshFlag Then
        RefreshFlag = False
        DocName = GetProperty(pDocType)
        If Len(DocName) Then
            For Index = 1 To UBound(documentName)
                If DocName = documentName(Index) Then Exit For
            Next
            If Index > UBound(documentName) Then Index = 0
        End If
        If Index = 0 Then Index = GetProjectIndexByURL(GetRegSelection(FName, SelectedDocument))
    End If
    'On Error GoTo ex
ex:
    DocumentSelected CLng(Index)
'    RefreshDocumentGroup'X
    GetSelectedDocumentIndex = Index
End Function
Sub SetSelectedDocumentIndex(Index As Integer)
    WriteLog 1, CurrentMod, "SetSelectedDocumentIndex"
    Dim FName As String, DocName As String
    On Error Resume Next
    FName = GetFileName(GetActiveFName(ActiveDocument))
    DocName = documentName(Index)
    If Len(FName) > 0 Then
        SetRegSelection FName, DocName, SelectedDocument
'        SetProperty "docentDocType", DocumentName(Index)
    End If
    DocNum = Index
    PrintView
    RefreshFlag = True
    RefreshDocumentGroup
'    On Error GoTo 0
'    DoEvents
End Sub
Function GetImage(ByVal ImageName As String) As IPicture
    If Right$(ImageName, Len(ImagesExtension)) <> ImagesExtension Then ImageName = ImageName & ImagesExtension
    If Left$(ImageName, Len(ImagesPath)) <> ImagesPath Then ImageName = ImagesPath & ImageName
    If Len(Dir(ImageName)) = 0 Then
        frmMsgBox.Display "Please wait...", Array(), None, "", ShowModal:=vbModeless
        frmMsgBox.Repaint
        UnzipAFile ThisDocument.FullName, ImagesPath, , "*.png", ImagesExtension
        Unload frmMsgBox
    End If
    Set GetImage = LoadPictureGDI(ImageName)
End Function
Private Function GetNumImage(ByVal ColorStr As String, n As Long) As IPicture
    Select Case ColorStr
    Case "Red": ColorStr = ColorStr & "W"
'    Case "Green": ColorStr = ColorStr & "W"
    End Select
    Set GetNumImage = GetImage(ColorStr & IIf(n > 9, "9PLUS", n))
'    ColorStr = ColorToCriticality(ColorStr)
'    ColorStr = ColorStr & IIf(n > 9, "9PLUS", n) & ImagesExtension
'
'    If Len(Dir(ImagesPath & ColorStr)) = 0 Then
''        Dim WtMsg As New frmMsgBox
'        frmMsgBox.Display "Please wait...", Array(), None, "", ShowModal:=vbModeless
'        frmMsgBox.Repaint
'        UnzipAFile ThisDocument.FullName, ImagesPath, , "*.png", ImagesExtension
'        Unload frmMsgBox
'    End If
'    Set GetImage = MLoadPictureGDI.LoadPictureGDI(ImagesPath & ColorStr)
End Function
Private Function ColorToCriticality(ByVal ColorStr As String) As String
    ColorStr = Replace(ColorStr, "Green", "Information")
    ColorStr = Replace(ColorStr, "Yellow", "Important")
    ColorStr = Replace(ColorStr, "Red", "Critical")
'    ColorStr = ColorStr & IIf(n > 9, "9PLUS", n) & ImagesExtension
    ColorToCriticality = ColorStr
End Function
Function GetTaskImage(ByVal ColorStr As String) As IPicture
    If Not GetVisibleGroup("IdGroupTasks") Then Exit Function
    Set GetTaskImage = GetNumImage(ColorStr, GetTaskCount(ColorStr))
End Function
Function GetTaskCount(ColorStr As String) As Long
    If Not GetVisibleGroup("IdGroupTasks") Then Exit Function
'    If Not IsProjectSelected(True) Then Exit Function
    On Error Resume Next
    If Len(ProjectNameStr) = 0 Then LoadProjects
    GetTraffic
'    If TasksDict Is Nothing Then
'        GetAPITasksCounts
'    ElseIf TasksDict.Count = 0 Then
'        GetAPITasksCounts
'    ElseIf TasksDict(1) Is Nothing Then
'        GetAPITasksCounts
'    End If
    GetTaskCount = TasksDict(ColorStr) '.Count
End Function
Sub GotoTaskCollection(ColorStr As String)
    If Not GetVisibleGroup("IdGroupTasks") Then Exit Sub
    On Error Resume Next
    If Len(ProjectNameStr) = 0 Then LoadProjects
    GetTraffic
'    If TasksDict Is Nothing Then
'        GetAPITasksCounts
'    ElseIf TasksDict.Count = 0 Then
'        GetAPITasksCounts
'    End If
    GoToLink ProjectURLStr & TasksDict(ColorStr & "URL") 'ProjectURLStr & "/site-collections/tasks-" & LCase(ColorStr)
End Sub

Function GetNotificationImage(ColorStr As String) As IPicture
'    if not GetVisibleGroup("IdGroupNotifications") then exit function
    Set GetNotificationImage = GetNumImage(ColorStr, GetNotificationCount(ColorStr))
End Function

Function GetNotificationCount(ColorStr As String) As Long
'    if not GetVisibleGroup("IdGroupNotifications") then exit function
'    If Not IsProjectSelected(True) Then Exit Function
    On Error Resume Next
    'If Len(ProjectNameStr) = 0 Then LoadProjects
    GetTraffic
'    If NotifsDict Is Nothing Then
'        GetAPINotificationsCounts
'    ElseIf NotifsDict.Count = 0 Then
'        GetAPINotificationsCounts
'    End If
    GetNotificationCount = NotifsDict(ColorStr) '.Count
End Function
Function GetTasksTrafficTooltip(clr As String, TName As String) As String
    Dim WDays As Long, TCount As Long
    If Not GetVisibleGroup("IdGroupTasks") Then Exit Function
    TCount = GetTaskCount(clr)
    WDays = ProjectInfo(LCase(clr))
'    GetTasksTrafficTooltip = "The color of the circle identifies the urgency of the Task." & vbNewLine & vbNewLine &
    GetTasksTrafficTooltip = clr & " Tasks are called """ & TName & " Tasks""" & " and are tasks due "
    Select Case clr
    Case "Green"
'        WDays = ProjectInfo("yellow") + 1
        GetTasksTrafficTooltip = GetTasksTrafficTooltip & "between " & ProjectInfo("yellow") + 1 & " and " & WDays & " working day" & IIf(WDays = 1, "", "s") & "."
    Case "Yellow"
        GetTasksTrafficTooltip = GetTasksTrafficTooltip & "between " & ProjectInfo("red") + 1 & " and " & WDays & " working day" & IIf(WDays = 1, "", "s") & "."
    Case "Red"
        GetTasksTrafficTooltip = GetTasksTrafficTooltip & "within " & WDays & " working day" & IIf(WDays = 1, "", "s or less.")
    End Select
    GetTasksTrafficTooltip = GetTasksTrafficTooltip & vbNewLine & vbNewLine & "You have " & TCount & " " & TName & " Task" & IIf(TCount = 1, "", "s") & "."
    GetTasksTrafficTooltip = GetTasksTrafficTooltip & vbNewLine & vbNewLine & "Clicking on the colored circle will bring you to the website list of the respective Tasks."
    
End Function
Sub GotoNotificationCollection(ColorStr As String)
'    if not GetVisibleGroup("IdGroupNotifications") then exit function
    On Error Resume Next
    If Len(ProjectNameStr) = 0 Then LoadProjects
    GetTraffic
'    If NotifsDict Is Nothing Then
'        GetAPINotificationsCounts
'    ElseIf NotifsDict.Count = 0 Then
'        GetAPINotificationsCounts
'    End If
    GoToLink ProjectURLStr & NotifsDict(ColorStr & "URL") 'ProjectURLStr & "/site-collections/tasks-" & LCase(ColorStr)
'    If IsProjectSelected(True) Then
'    GoToLink ProjectURLStr & "/site-collections/notifications-" & LCase(ColorToCriticality(ColorStr))
End Sub
Sub ProjectSelected(Doc As Document, SelectedItem As Long, Optional Force As Boolean)
    On Error Resume Next
    Dim PName As String
    Static LastPName As String
    'NewPNum = SelectedItem
    If CodeIsRunning And Not Force Then Exit Sub
    WriteLog 1, CurrentMod, "ProjectSelected", "SelectedItem: " & SelectedItem
    CreateBorder Doc, ProjectColorStr
    PName = GetProjectNameByIndex(GetProjectIndexByIndex(SelectedItem))
'    If PName = "Select Project" Then PName = ""
'    Debug.Print PName
    If Not CodeIsRunning Then ProjectNameStr = PName
'    IsAuthorized = IsValidUser = "OK"
    AskForMissingPw = True
    LoadProjectInfoReg PName
    If CodeIsRunning Then Doc.Saved = True
    RefreshProject
    If Len(ProjectNameStr) And LastPName <> ProjectNameStr Then
        If Mid(GetNeverHelpAgain, 4, 1) <> 1 Then
            LastPName = ProjectNameStr
            If frmMsgBox.Display(ProjectNameStr & " project was selected.", Array("Ok", "Don't Show Again")) = "Don't Show Again" Then  ', AutoCloseTimer:=3
                SetNeverHelpAgain 4
            End If
        End If
    End If
End Sub
Sub CancelEditingDoc()
    On Error Resume Next
    CodeIsRunning = True
    ActiveDocument.Close False
    CodeIsRunning = False
    RefreshRibbon True
End Sub
Function GetStateIcon(StateStr As String) As String
    On Error GoTo ex
    If InStr(StateStr, "None") Then GetStateIcon = "SlideNew"
    If InStr(StateStr, "Private") Then GetStateIcon = "ProtectDocument"
    If InStr(StateStr, "Draft") Then GetStateIcon = "AdpDiagramNewTable"
    If InStr(StateStr, "Review") Then GetStateIcon = "AdvertisePublishAs" '"SetLanguage"
    If InStr(StateStr, "Published") Then GetStateIcon = "SetLanguage" '"BlogPublish"
    If InStr(StateStr, "Closed") Then GetStateIcon = "FilePermissionRestrictMenu"
    If InStr(StateStr, "Submitted") Then GetStateIcon = "BlogPublish"
    If InStr(StateStr, "Archived") Then GetStateIcon = "WatermarkGallery" 'A for Archived!
    If InStr(StateStr, "Pending") Then GetStateIcon = "SetLanguage" '"BlogPublish"
    
    
'Image="ReviewNewComment" (For postItNote)
'Image="HighImportance"
ex:
End Function
Function ApplyTransitionNo(i As Long, Doc As Document)
    If Doc.Name = "" Then Exit Function
    If IsProjectSelected Then UploadDoc ActiveDocument, Replace(NextTransitions(i)("@id"), "FILEURl", GetProperty(pDocURL, Doc)) 'GetStateFromTrn(NextTransitions(3))
End Function


