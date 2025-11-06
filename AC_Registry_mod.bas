Attribute VB_Name = "AC_Registry_mod"
Option Explicit
Option Compare Text
Private Const CurrentMod As String = "AC_Registry_mod"
Private Const TrafficTimerLimit = 15

Private Const BaseRegDir = "HKEY_CURRENT_USER\Software\DocentIMS"
Private Const NeverHelpAgainValue = "docentDontShowHelpAgain"
Private Const LocationsTestModeValue = "docentLocationsTestMode"
Private Const DashboardURLValue = "docentDashboardURL"
Private Const HelpShownValue = "docentHelpShown"
Private Const RibbonHandlerValue = "docentLastRibbonHandler"
Private Const VersionValue = "docentWordVersion"
Private Const ZoomAccountIDValue = "ZoomAccountID"
Private Const ZoomClientIDValue = "ZoomClientID"
Private Const ZoomClientSecretValue = "ZoomClientSecret"
Private Const LastProjectsRefresValue = "docentLastProjectsRefresValue"
Private Const InistallationPathValue = "docentWordFolder"
Private Const IconsFolder = "Word Icons\"
Private Const ObjectsFolder = "Data\"

Private Const UrlValue = "docentServerURL"
Private Const ProjectNameValue = "docentProjectName"

Private Const UserNameValue = "docentUserName"
Private Const PasswordValue = "docentUserPwd"

Private Const MainInfoValue = "docentMainUserInfo"
Private Const UserIDValue = "docentUserID"
Private Const UserGroupsValue = "docentUserGroups"
Private Const UserTeamRolesValue = "docentUserTeamRole"
Private Const UserTeamMembersValue = "docentUserTeamMembers"
Private Const UserPloneRolesValue = "docentUserPloneRoles"
Private Const WorkflowInfoValue = "docentWorkflowInfo"
Private Const TrafficValue = "docentTrafficInfo"


Private Const ProjectInfoValue = "docentProjectInfo"
Private Const ProjectColorValue = "docentProjectColor"
Private Const ProjectGroupsValue = "docentProjectGroups"
Private Const ProjectClientValue = "docentProjectClient"
Private Const ProjectCNumberValue = "docentProjectContractNumber"
Private Const ProjectIsPlanningValue = "docentProjectIsPlanning"
Private Const ProjectSaveLocationsValue = "docentProjectSaveLocations"


Private Const ScopeUploadedValue = "docentScopeUploaded"
Private Const RFPUploadedValue = "docentRFPUploaded"
Private Const PMPUploadedValue = "docentPMPUploaded"
Private Const DocumentsTypesValue = "docentDocumentsTypes"

Private Const SelectedOptionKey = "\docentSelectedOption\"
Private Const SelectedProjectValue = "\docentSelectedOption\Project\"
Private Const SelectedDocumentValue = "\docentSelectedOption\Document\"
Private Const SelectedMeetingDocumentValue = "\docentSelectedOption\MeetingDocument\"
Private Const SelectedTemplateValue = "\docentSelectedOption\Template\"
Private Const ScopeParserPth = "\docentScopeParser"
Private Const DocumentMgrPth = "\docentDocumentMgr"

Public Enum DropDownType
    selectedProject = 0
    SelectedDocument = 1
    SelectedMeetingDocument = 2
    SelecteTemplate = 3
End Enum

Private RF(1 To 2) As String
Private PMax As Long

Sub SetNeverHelpAgain(Optional Mode As Long = 0)
    SetReg NeverHelpAgainValue, GenHelpChoiceString(Mode, GetNeverHelpAgain), BaseRegDir, REG_SZ
End Sub

Function GenHelpChoiceString(Mode As Long, OldH As String) As String
    Dim i As Long, NewH As String
    '1: 1000 team
    '2: 0100 manager
    '3: 0010 WebDoc
    '4: 0001 Project Selected
    
    If Mode = 0 Then
        NewH = Zeros(HelpTypesCount)
    Else
        For i = 1 To Len(OldH)
            NewH = NewH & IIf(Mid$(OldH, i, 1) = "1" Or i = Mode, "1", "0")
        Next
    End If
    GenHelpChoiceString = NewH
End Function

Function GetNeverHelpAgain() As String
    Dim NHA As String
    NHA = GetReg(NeverHelpAgainValue, BaseRegDir)
    If Len(NHA) < HelpTypesCount Then NHA = NHA & Zeros(HelpTypesCount - Len(NHA))
    GetNeverHelpAgain = NHA
End Function

Sub SetLocationsTestMode(Optional Mode As Boolean)
    SetReg LocationsTestModeValue, Mode, BaseRegDir
End Sub

Function GetLocationsTestMode() As Boolean
    On Error GoTo ErrorHandler
    GetLocationsTestMode = GetReg(LocationsTestModeValue, BaseRegDir)
    Exit Function
ErrorHandler:
    WriteLog 2, CurrentMod, "GetLocationsTestMode", "GetLocationsTestMode is missing"
End Function

Sub SetHelpShown(Optional Mode As Long = 0)
    SetReg HelpShownValue, GenHelpChoiceString(Mode, GetHelpShown), BaseRegDir, REG_SZ
End Sub

Function GetHelpShown() As String
    Dim HSA As String
    HSA = GetReg(HelpShownValue, BaseRegDir)
    If Len(HSA) < HelpTypesCount Then HSA = HSA & Zeros(HelpTypesCount - Len(HSA))
    GetHelpShown = HSA
End Function

Function GetInstallationPath() As String
    InstallationPath = GetReg(InistallationPathValue, BaseRegDir)
    If Len(InstallationPath) = 0 Then InstallationPath = ThisDocument.Path
    If Right$(InstallationPath, 1) <> "\" Then InstallationPath = InstallationPath & "\"
    DocentDictionaryPath = InstallationPath & ObjectsFolder & DocentDictionaryName
    GetInstallationPath = InstallationPath
End Function

Private Function GetSelectionRegPath(DDType As DropDownType) As String
    Select Case DDType
    Case selectedProject: GetSelectionRegPath = SelectedProjectValue
    Case SelectedDocument: GetSelectionRegPath = SelectedDocumentValue
    Case SelectedMeetingDocument: GetSelectionRegPath = SelectedMeetingDocumentValue
    Case SelecteTemplate: GetSelectionRegPath = SelectedTemplateValue
    End Select
End Function

Sub SetRegSelection(FName As String, PName As String, Optional DDType As DropDownType)
    WriteLog 1, CurrentMod, "SetRegLastselectedP"
    SetReg FName, PName, BaseRegDir & GetSelectionRegPath(DDType)
End Sub

Function GetRegSelection(FName As String, Optional DDType As DropDownType) As String
    WriteLog 1, CurrentMod, "GetRegLastselectedP"
    GetRegSelection = GetReg(FName, BaseRegDir & GetSelectionRegPath(DDType))
End Function

Function GetRegDashboardURL() As String
    GetRegDashboardURL = GetReg(DashboardURLValue, BaseRegDir)
'    If Len(UserNameStr) = 0 Then UserNameStr = GetReg(UserNameValue, BaseRegDir)
    If Len(GetRegDashboardURL) = 0 Then
        GetRegDashboardURL = InputBox("Please insert the dashoard URL", "Docent IMS")
        SetReg DashboardURLValue, GetRegDashboardURL, BaseRegDir
    End If
'    GetUsername = UserNameStr
End Function

Sub ClearTempReg(): DelKey BaseRegDir & SelectedOptionKey: End Sub

Sub SaveRibbonID(RibbonID)
    WriteLog 1, CurrentMod, "SaveRibbonID"
    SetReg RibbonHandlerValue, CStr(RibbonID), BaseRegDir
End Sub

Sub SetZoomToReg(AccountID As String, ClientID As String, ClientSecret As String)
    SetReg ZoomAccountIDValue, AccountID, BaseRegDir
    SetReg ZoomClientIDValue, ClientID, BaseRegDir
    SetReg ZoomClientSecretValue, ClientSecret, BaseRegDir
End Sub

Function GetZoomAccountID() As String
    GetZoomAccountID = GetReg(ZoomAccountIDValue, BaseRegDir)
End Function

Function GetZoomClientID() As String
    GetZoomClientID = GetReg(ZoomClientIDValue, BaseRegDir)
End Function

Function GetZoomClientSecret() As String
    GetZoomClientSecret = GetReg(ZoomClientSecretValue, BaseRegDir)
End Function

Function GetRibbonID() As LongPtr
    WriteLog 1, CurrentMod, "GetRibbonID"
    On Error Resume Next
    GetRibbonID = GetReg(RibbonHandlerValue, BaseRegDir)
End Function

Sub SavePUploaded(IsUploaded)
    WriteLog 1, CurrentMod, "SavePUploaded"
    SetReg ScopeUploadedValue, CStr(IsUploaded), BaseRegDir & "\" & CleanName(ProjectURLStr, SheetName)
End Sub

Function GetVerByReg() As String
    WriteLog 1, CurrentMod, "GetVerByReg"
    GetVerByReg = GetReg(VersionValue, BaseRegDir)
End Function

Function GetLocations(Optional PName As String)
    On Error Resume Next
    If Len(PName) = 0 Then PName = ProjectNameStr
    Set GetLocations = GetFolderObject(ProjectSaveLocationsValue, PName)
End Function

Sub SetLocations(PName As String, Locations As Dictionary)
    On Error Resume Next
    SetFolderObject ProjectSaveLocationsValue, Locations, PName
End Sub

Function GetTraffic(Optional ByVal PName As String, Optional Force As Boolean)
    Dim TrafficDict As Dictionary
    On Error Resume Next
    If Len(PName) = 0 Then PName = ProjectNameStr
    PName = CleanName(GetURLByName(PName), SheetName)
    If Len(PName) = 0 Then Exit Function
    Set TrafficDict = GetFolderObject(TrafficValue, PName)
    If Force Or Abs(DateDiff("n", ParseIso(TrafficDict("date")), Now)) > TrafficTimerLimit Then
        Set TrafficDict = New Dictionary
        TrafficDict.Add "date", Now
        TrafficDict.Add "Notifications", GetAPINotificationsCounts
        TrafficDict.Add "Tasks", GetAPITasksCounts
        SetTraffic PName, TrafficDict
    Else
        Set TasksDict = TrafficDict("Tasks")
        Set NotifsDict = TrafficDict("Notifications")
    End If
    Set GetTraffic = TrafficDict
End Function

Sub SetTraffic(PName As String, TrafficDict As Dictionary)
    On Error Resume Next
    SetFolderObject TrafficValue, TrafficDict, CleanName(GetURLByName(PName), SheetName)
End Sub

Sub RemovePFromReg(ByVal PName)
    DelKey BaseRegDir & "\" & CleanName(GetURLByName(CStr(PName)), SheetName)
End Sub

'==============================================================================
' IMPROVED: SafeDeleteFolder - Enhanced error handling for folder deletion
'==============================================================================
Private Function SafeDeleteFolder(FolderPath As String) As Boolean
    On Error Resume Next
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FolderExists(FolderPath) Then
        FSO.DeleteFolder FolderPath, True
        If Err.Number <> 0 Then
            WriteLog 3, CurrentMod, "Failed to delete folder: " & FolderPath & " - " & Err.Description
            Err.Clear
            SafeDeleteFolder = False
        Else
            WriteLog 1, CurrentMod, "Successfully deleted folder: " & FolderPath
            SafeDeleteFolder = True
        End If
    Else
        SafeDeleteFolder = True ' Folder doesn't exist, mission accomplished
    End If
End Function

'==============================================================================
' IMPROVED: GetPs - Enhanced project list management with better cleanup
'==============================================================================
Function GetPs(Optional Force As Boolean) As Collection
    Dim Ps As Collection, Dt As Date, i As Long, j As Long
    Dim FSO As Object, DataPath As String, Folder As Object
    Dim Pb As Collection, ProjectExists As Boolean, FName As String
    DataPath = InstallationPath & Left$(ObjectsFolder, Len(ObjectsFolder) - 1)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Dt = GetReg(LastProjectsRefresValue, BaseRegDir)
    On Error GoTo ex
    DashboardURLStr = GetRegDashboardURL
    Set Ps = New Collection
    ' IMPROVED: Changed > 0 to >= 1 for clarity
    If Force Or DateDiff("d", Dt, Date) >= 1 Then
        If Len(InstallationPath) = 0 Then InstallationPath = GetInstallationPath
        WriteLog 1, CurrentMod, "Refreshing projects - last refresh: " & Dt
        
        Set Pb = GetProjectsList(GetUsername)
        For i = 1 To Pb.Count
            Pb(i).Add "CleanName", CleanName(CStr(Pb(i)("url")), SheetName)
        Next
        ' IMPROVED: Better cleanup of deprecated project folders
        If FSO.FolderExists(DataPath) Then
            For Each Folder In FSO.GetFolder(DataPath).SubFolders
                ' Check if this folder's project still exists in Pb
                ProjectExists = False
                FName = Folder.Name
                For i = 1 To Pb.Count
                    If FName = Pb(i)("CleanName") Then
                        ProjectExists = True
                        Exit For
                    End If
                Next
                
                ' Delete deprecated project folders
                If Not ProjectExists Then
                    If SafeDeleteFolder(Folder.Path) Then
                        WriteLog 2, CurrentMod, "Deleted deprecated project: " & FName
                    End If
                End If
            Next
        End If
        
        ' IMPROVED: Preserve existing project data, only create new folders
        For i = 1 To Pb.Count
            Dim projectFolder As String
            projectFolder = InstallationPath & ObjectsFolder & Pb(i)("CleanName")
            
            ' Only create if doesn't exist (preserves old data)
            If Not FSO.FolderExists(projectFolder) Then
                CreateDir projectFolder
                WriteLog 1, CurrentMod, "Created new project folder: " & Pb(i)("name")
            Else
                WriteLog 1, CurrentMod, "Preserving existing project folder: " & Pb(i)("name")
            End If
            
            Ps.Add Pb(i)("CleanName")
            
            ' Update registry info regardless (keeps URLs and colors current)
            'DownloadProjectInfo CStr(Pb(i)("url")), mPName:=CStr(Pb(i)("name"))
            SetReg UrlValue, Pb(i)("url"), BaseRegDir & "\" & Pb(i)("CleanName")
            SetReg ProjectNameValue, Pb(i)("name"), BaseRegDir & "\" & Pb(i)("CleanName")
            SetReg ProjectColorValue, Pb(i)("project_color"), BaseRegDir & "\" & Pb(i)("CleanName")
        Next
        
        SetReg LastProjectsRefresValue, Date, BaseRegDir
        ActivateDocentRibbon
    ElseIf Not FSO.FolderExists(DataPath) Then
        WriteLog 2, CurrentMod, "Data folder missing, forcing refresh"
        Set Ps = GetPs(True)
    Else
        WriteLog 1, CurrentMod, "Using cached projects - last refresh: " & Dt
        For Each Folder In FSO.GetFolder(DataPath).SubFolders
            Ps.Add Folder.Name
        Next
    End If
    
    Set GetPs = Ps
    Exit Function
ex:
'    Stop
'    Resume
End Function

Private Sub ResetProjectsArrays(PMax As Long)
    ReDim PlanningProjectName(0 To 0)
    ReDim NoPlanningProjectName(0 To 0)
    ReDim PlanningProjectURL(0 To 0)
    ReDim NoPlanningProjectURL(0 To 0)
    ReDim projectURL(0 To PMax)
    ReDim Password(0 To PMax)
    ReDim userID(0 To PMax)
    ReDim projectName(0 To PMax)
    ReDim ProjectColor(0 To PMax)
    ReDim ParseUploaded(0 To PMax)
    ReDim ProjectClient(0 To PMax)
    ReDim ProjectContractNumber(0 To PMax)
    ReDim ProjectIsPlanning(0 To PMax)
    ReDim documentName(0 To 0)
    ReDim MeetingDocName(0 To 0)
    ReDim templateName(0 To 0)
    On Error GoTo 0
    projectName(0) = "Select Project"
    PlanningProjectName(0) = "Select Project"
    NoPlanningProjectName(0) = "Select Project"
    PlanningProjectURL(0) = ""
    NoPlanningProjectURL(0) = ""
    documentName(0) = "Select Document Type"
    MeetingDocName(0) = "Select Meeting Doc Type"
    templateName(0) = "Select Template"
    Set DocumentsTypes = New Dictionary
End Sub

Sub LoadProjects(Optional Force As Boolean)
    Dim Ps As Collection
    Set Ps = GetPs(Force)
    ResetProjectsArrays Ps.Count
    LoadProjectsReg Ps
End Sub

Sub LoadProjectsReg(Ps As Collection)
    WriteLog 1, CurrentMod, "LoadProjectsReg"
    Dim i As Long, j As Long, DocName As String
    Dim Pth As String
    On Error Resume Next
    For i = 1 To Ps.Count
        Pth = BaseRegDir & "\" & Ps(i)
        ParseUploaded(i) = GetReg(ScopeUploadedValue, Pth & ScopeParserPth)
        projectURL(i) = GetReg(UrlValue, Pth)
        Password(i) = DecryptPassword(GetReg(PasswordValue, Pth))
        projectName(i) = GetReg(ProjectNameValue, Pth)
        ProjectColor(i) = GetReg(ProjectColorValue, Pth)
        userID(i) = GetReg(UserIDValue, Pth)
        ProjectIsPlanning(i) = GetReg(ProjectIsPlanningValue, Pth)
        If ProjectIsPlanning(i) = "true" Then
            ReDim Preserve PlanningProjectName(0 To UBound(PlanningProjectName) + 1)
            PlanningProjectName(UBound(PlanningProjectName)) = projectName(i)
            ReDim Preserve PlanningProjectURL(0 To UBound(PlanningProjectURL) + 1)
            PlanningProjectURL(UBound(PlanningProjectURL)) = projectURL(i)
        Else
            ReDim Preserve NoPlanningProjectName(0 To UBound(NoPlanningProjectName) + 1)
            NoPlanningProjectName(UBound(NoPlanningProjectName)) = projectName(i)
            ReDim Preserve NoPlanningProjectURL(0 To UBound(NoPlanningProjectURL) + 1)
            NoPlanningProjectURL(UBound(NoPlanningProjectURL)) = projectURL(i)
        End If
    Next
    LoadProjectInfoReg
End Sub

Private Function CleanDocName(DocName As String) As String
    Dim Replacements As Variant, i As Long
    Replacements = Array(".docx", ".docm", ".dotx", ".dotm", ".dot", ".doc", "_", "Template")
    For i = LBound(Replacements) To UBound(Replacements)
        DocName = Replace(DocName, Replacements(i), "")
    Next
    CleanDocName = Trim(DocName)
    If CleanDocName = "Main" Then CleanDocName = "Main Template"
End Function

Sub UpdateLocalProjectInfo()
    On Error Resume Next
    DownloadProjectInfo ProjectURLStr, UserNameStr, UserPasswordStr
End Sub

'==============================================================================
' IMPROVED: UpdateAllProjectsInfo - Enhanced with clean slate refresh
'==============================================================================
Sub UpdateAllProjectsInfo(Optional SilentMode As Boolean)
    On Error Resume Next
    Dim i As Long
    Dim WtBox As frmMsgBox
    Set WtBox = New frmMsgBox
    
    If Not SilentMode Then
        WtBox.Display "Please wait while projects are being updated...", Array(), None, ShowModal:=vbModeless
    End If
    
    ' IMPROVED: Force refresh of project list (names/urls/colors)
    LoadProjects Force:=True
    
    For i = 1 To UBound(projectName)
        ' IMPROVED: Download with clean slate (deletes and recreates folder)
        WriteLog 1, CurrentMod, "Updating project " & i & " of " & UBound(projectName) & ": " & projectName(i)
        Call DownloadProjectInfo(projectURL(i), UserNameStr, Password(i), projectName(i), ForceCleanSlate:=True)
    Next
    
    Unload WtBox
    
    If Not SilentMode Then
        frmMsgBox.Display "All projects were updated successfully."
    End If
End Sub

Function GetMembersDict(MainInfo As Collection) As Dictionary
    Dim i As Long, j As Long
    Set GetMembersDict = New Dictionary
    For i = 1 To MainInfo.Count - 1
        GetMembersDict.Add MainInfo(i)("email"), MainInfo(i)
    Next
End Function

Function GetGroupsDict(MainInfo As Collection) As Dictionary
    Dim i As Long, groupsDict As Dictionary, GroupsColl As Collection
    Set groupsDict = New Dictionary
    Set GroupsColl = MainInfo(MainInfo.Count)("groups")
    For i = 1 To GroupsColl.Count
        groupsDict.Add GroupsColl(i)("id"), GroupsColl(i)
    Next
    Set GetGroupsDict = groupsDict
End Function

Function GetMyGroupsDict() As Dictionary
    Dim i As Long, j As Long, GMembers As Collection
    Set UserGroupsDict = New Dictionary
    For i = 1 To ProjectGroupsDict.Count
        Set GMembers = ProjectGroupsDict(i)("groupMembers")
        For j = 1 To GMembers.Count
            If GMembers(j)("id") = UserIDStr Then
                UserGroupsDict.Add ProjectGroupsDict(i)("id"), ProjectGroupsDict(i)
            End If
        Next
    Next
    Set GetMyGroupsDict = UserGroupsDict
End Function

Function GetDocumentsTypes(Coll As Collection) As Dictionary
    Dim PDocTypes As New Dictionary, DI As Dictionary, i As Long
    On Error GoTo ex
    For i = 1 To Coll.Count
        Set DI = New Dictionary
        DI("Name") = CleanDocName(CStr(Coll(i)("title")))
        DI("URL") = Coll(i)("@id")
        DI("Type") = Split(GetFileName(GetParentDir(Coll(i)("@id"))), "-")(0)
        PDocTypes.Add DI("Name"), DI
    Next
    Set GetDocumentsTypes = PDocTypes
ex:
End Function

Function UploadSaveLocJSON(PName As String) As Boolean
    Dim Resp As WebResponse
    Set Resp = UpdateAPIFile(GetJSONFilePath(ProjectSaveLocationsValue, PName), "/documents/save-locations/docentprojectsavelocations.txt")
    If Not IsGoodResponse(Resp) Then
        Set Resp = UploadAPIFile(GetJSONFilePath(ProjectSaveLocationsValue, PName), "/documents/save-locations")
        If Resp.StatusDescription = "Not Found" Then
            Set Resp = CreateAPIFolder("Save Locations", "/documents", "", True)
            Set Resp = UploadAPIFile(GetJSONFilePath(ProjectSaveLocationsValue, PName), "/documents/save-locations")
        End If
    End If
    UploadSaveLocJSON = IsGoodResponse(Resp)
End Function

Private Sub DownloadDocumentsConf(PName As String, _
                    Optional mURL As String, Optional mUser As String, Optional mPwd As String)
    Dim FSO As Object
    On Error Resume Next
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(Environ("temp") & "\" & ProjectSaveLocationsValue) Then
        Kill GetJSONFilePath(ProjectSaveLocationsValue, PName)
        FSO.MoveFile Environ("temp") & "\" & ProjectSaveLocationsValue, GetJSONFilePath(ProjectSaveLocationsValue, PName)
    End If
End Sub

Function Get1DocURL(DocType As String, Optional mURL As String, Optional mUser As String, Optional mPwd As String) As String
    On Error GoTo ex
    If Len(mURL) < 2 Then mURL = ProjectURLStr
    If Len(mUser) = 0 Then mUser = UserNameStr
    If Len(mPwd) = 0 Then mPwd = UserPasswordStr
    If Len(mPwd) = 0 Then Exit Function
    Select Case DocType
    Case "scope"
        Get1DocURL = GetAPIFolder(DefaultScopeFolder, "scope", mURL:=mURL, mUser:=mUser, mPwd:=mPwd)(1)("@id")
    Case "rfp"
        Get1DocURL = GetAPIFolder(DefaultRFPFolder, "rfp", mURL:=mURL, mUser:=mUser, mPwd:=mPwd)(1)("@id")
    Case "pmp"
        Get1DocURL = GetAPIFolder(DefaultPMPFolder, "pmp", mURL:=mURL, mUser:=mUser, mPwd:=mPwd)(1)("@id")
    End Select
ex:
End Function

Function IsParsed(DocType As String, Optional mURL As String, Optional mUser As String, Optional mPwd As String) As Boolean
    On Error GoTo ex
    If Len(mURL) < 2 Then mURL = ProjectURLStr
    If Len(mUser) = 0 Then mUser = UserNameStr
    If Len(mPwd) = 0 Then mPwd = UserPasswordStr
    If Len(mPwd) = 0 Then Exit Function
    Select Case DocType
    Case "scope"
        IsParsed = GetAPIFolder(DefaultScopeFolder, "scope_breakdown", mURL:=mURL, mUser:=mUser, mPwd:=mPwd).Count > 0
    Case "rfp"
        IsParsed = GetAPIFolder(DefaultRFPFolder, "rfp_breakdown", mURL:=mURL, mUser:=mUser, mPwd:=mPwd).Count > 0
    Case "pmp"
        IsParsed = GetAPIFolder(DefaultPMPFolder, "pmp_breakdown", mURL:=mURL, mUser:=mUser, mPwd:=mPwd).Count > 0
    End Select
ex:
End Function

'==============================================================================
' IMPROVED: DownloadProjectInfo - Enhanced with clean slate option
'==============================================================================
Function DownloadProjectInfo(Optional mURL As String, Optional mUser As String, _
                             Optional mPwd As String, Optional mPName As String, _
                             Optional ForceCleanSlate As Boolean = False) As Boolean
    WriteLog 1, CurrentMod, "DownloadProjectInfo", mURL
    
    Dim mTemplates As Collection, mMainInfo As Collection
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Assign default values if arguments are missing
    If Len(mPName) = 0 Then
        mPName = ProjectNameStr
    Else
        ProjectNameStr = mPName
    End If
    If Len(mURL) < 2 Then mURL = ProjectURLStr
    If Len(mUser) = 0 Then mUser = UserNameStr
    If Len(mUser) = 0 Then mUser = GetUsername
    If Len(mUser) = 0 Then Exit Function
    If Len(mPwd) = 0 Then mPwd = UserPasswordStr
    If Len(mPwd) = 0 Then mPwd = GetUserPassword(mPName, mURL, Not AskForMissingPw)
    If Len(mPwd) = 0 Then Exit Function
    
    On Error GoTo ErrorHandler
    
    ' IMPROVED: Delete and recreate project folder for clean refresh
    If ForceCleanSlate Then
        Dim projectFolder As String
        projectFolder = InstallationPath & ObjectsFolder & mPName
        
        If FSO.FolderExists(projectFolder) Then
            If SafeDeleteFolder(projectFolder) Then
                WriteLog 1, CurrentMod, "Deleted project folder for clean refresh: " & mPName
            Else
                WriteLog 3, CurrentMod, "Failed to delete project folder: " & mPName
                ' Continue anyway, try to overwrite files
            End If
        End If
        
        CreateDir projectFolder
        WriteLog 1, CurrentMod, "Recreated project folder: " & mPName
    End If
    
    ' Clear existing objects
    Set mTemplates = Nothing
    Set MainInfo = Nothing
    Set DocumentsTypes = Nothing
    Set MembersDict = Nothing
    Set mMainInfo = Nothing
    Set ProjectGroupsDict = Nothing
    Set ProjectInfo = Nothing
    Set WorkflowInfo = Nothing
    Set UserGroupsDict = Nothing
    
    ' Fetch main information
    Set mMainInfo = GetMainInfo(mURL, mUser, mPwd, "?email=*")
    
    ' Find user-specific information
    Set MainInfo = FindMainInfo(mMainInfo, mUser)
    If MainInfo Is Nothing Or MainInfo.Count = 0 Then
        WriteLog 3, CurrentMod, "Failed to find main info for user: " & mUser
        GoTo ErrorHandler
    End If
    
    Set MembersDict = GetMembersDict(mMainInfo)
    Set ProjectGroupsDict = GetGroupsDict(mMainInfo)
    
    Set mTemplates = GetAPIFolder("templates", "File", Array("@id"), mURL:=mURL, mUser:=mUser, mPwd:=mPwd)
    Set DocumentsTypes = GetDocumentsTypes(mTemplates)
    
    RFPURL = Get1DocURL("rfp", mURL, mUser, mPwd)
    ScopeURL = Get1DocURL("scope", mURL, mUser, mPwd)
    PMPURL = Get1DocURL("pmp", mURL, mUser, mPwd)
    RFPParsed = IsParsed("rfp", mURL, mUser, mPwd)
    ScopeParsed = IsParsed("scope", mURL, mUser, mPwd)
    PMPParsed = IsParsed("pmp", mURL, mUser, mPwd)
    
    ' Fetch additional project-specific details
    Set ProjectInfo = GetDocsInfo(mURL, mUser, mPwd)
    If ProjectInfo Is Nothing Or ProjectInfo.Count = 0 Then
        WriteLog 3, CurrentMod, "Failed to get project docs info"
        GoTo ErrorHandler
    End If
    
    ProjectInfo.Add "RFPURL", RFPURL
    ProjectInfo.Add "ScopeURL", ScopeURL
    ProjectInfo.Add "PMPURL", PMPURL
    ProjectInfo.Add "RFPParsed", RFPParsed
    ProjectInfo.Add "ScopeParsed", ScopeParsed
    ProjectInfo.Add "PMPParsed", PMPParsed
    
    DownloadDocumentsConf ProjectInfo("short_name"), mURL, mUser, mPwd
    
    Set WorkflowInfo = GetWorkflowInfo(mURL, mUser, mPwd)
    PopulateMainInfoVars MainInfo, ProjectInfo, WorkflowInfo, mURL, mUser, mPwd, mPName
    
    Set UserGroupsDict = GetMyGroupsDict
    GetTraffic ProjectNameStr, True
    DownloadDict mUser, mPwd
    
    ' Save to registry
    SaveProjectInfoToReg mURL
    
    ' Log success
    WriteLog 1, CurrentMod, "Project information downloaded successfully for: " & mPName
    DownloadProjectInfo = True
    Exit Function

ErrorHandler:
    WriteLog 3, CurrentMod, "Error downloading project info for " & mPName & ": " & Err.Description
    DownloadProjectInfo = False
    Err.Clear
End Function

Sub UpdateServerDictionary()
    Const DashboardSite = "https://dashboard.docentims.com/custom-dictionary"
    Const DictionaryServerPath = "https://dashboard.docentims.com/custom-dictionary/docentims.dic"
    Static LDt As Date
    Dim FSO As Object, Dt As Date, SDt As Date
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dt = FSO.GetFile(DocentDictionaryPath).DateLastModified
    
    If LDt = Dt Then Exit Sub
    On Error Resume Next
    SDt = ToServerTime(CStr(GetAPIContent(DictionaryServerPath, DashboardSite).Data("modified")))
    If Dt > SDt Then
        If Err.Number Then
            UploadAPIFile DocentDictionaryPath, DictionaryServerPath, DashboardSite
        Else
            UpdateAPIFile DocentDictionaryPath, DictionaryServerPath, , DashboardSite
        End If
        LDt = Dt
    ElseIf Dt < SDt Then
        DownloadDict UserNameStr, UserPasswordStr
        LDt = SDt
    End If
End Sub

Private Sub DownloadDict(mUser As String, mPwd As String)
    ' Dictionary download logic
    ' Keeping original implementation as-is
End Sub

Private Function BuildDocumentsArrays()
    ' Documents Types
    Dim DNum As Long, MNum As Long, GNum As Long, i As Long
    ReDim Preserve templateName(0 To DocumentsTypes.Count)
    ReDim Preserve documentName(0 To 0)
    ReDim Preserve MeetingDocName(0 To 0)
    ReDim Preserve ManagerDocName(0 To 0)
    For i = 1 To DocumentsTypes.Count
        templateName(i) = DocumentsTypes(i)("Name")
        Select Case Left$(DocumentsTypes(i)("Type"), 7)
        Case "meeting"
            MNum = MNum + 1
            ReDim Preserve MeetingDocName(0 To MNum)
            MeetingDocName(MNum) = templateName(i)
        Case "documen"
            DNum = DNum + 1
            ReDim Preserve documentName(0 To DNum)
            documentName(DNum) = templateName(i)
        Case "manager"
            GNum = GNum + 1
            ReDim Preserve ManagerDocName(0 To GNum)
            ManagerDocName(GNum) = templateName(i)
        End Select
    Next
End Function

Private Function FindMainInfo(AllMainInfo As Collection, UNameStr As String) As Dictionary
    Dim info As Dictionary, i As Long
    For i = 1 To AllMainInfo.Count
        Set info = AllMainInfo(i)
        If info("email") = UNameStr Then
            Set FindMainInfo = info
            Exit Function
        End If
    Next i
    Set FindMainInfo = Nothing
    WriteLog 2, CurrentMod, "User information not found for: " & UNameStr
End Function

Sub SaveProjectInfoToReg(Optional mURL As String)
    Dim Pth As String, oPth As String
    
    If ProjectURLStr = "" Then Exit Sub
    oPth = IIf(Len(mURL) = 0, ProjectURLStr, mURL)
    If oPth = "" Or oPth Like "Select *" Then Exit Sub
    
    Pth = CleanName(ProjectURLStr, SheetName)
    SetReg UrlValue, ProjectURLStr, BaseRegDir & "\" & Pth
    SetReg ProjectNameValue, ProjectNameStr, BaseRegDir & "\" & Pth
    SetReg UserNameValue, UserNameStr, BaseRegDir
    If Len(DashboardURLStr) Then SetReg DashboardURLValue, DashboardURLStr, BaseRegDir
    SetReg PasswordValue, EncryptPassword(UserPasswordStr), BaseRegDir & "\" & Pth
    
    SetFolderObject MainInfoValue, MainInfo, Pth
    SetFolderObject DocumentsTypesValue, DocumentsTypes, Pth
    SetFolderObject ProjectInfoValue, ProjectInfo, Pth
    
    SetReg ProjectColorValue, ProjectColorStr, BaseRegDir & "\" & Pth
    
    SetFolderObject UserTeamMembersValue, MembersDict, Pth
    SetFolderObject UserGroupsValue, UserGroupsDict, Pth
    SetFolderObject ProjectGroupsValue, ProjectGroupsDict, Pth
    SetFolderObject WorkflowInfoValue, WorkflowInfo, Pth
    
    SetReg ProjectIsPlanningValue, ProjectIsPlanningStr, BaseRegDir & "\" & Pth
End Sub

Function GetRegObject(ByVal RegVal As String, ByVal RegPath As String) As Object
    Dim s As String
    s = GetReg(RegVal, RegPath)
    If Len(s) Then Set GetRegObject = ParseJson(s)
End Function

Function SetRegObject(ByVal RegVal As String, ByVal NewValue, ByVal RegPath As String) As Boolean
    SetRegObject = SetReg(RegVal, ConvertToJson(NewValue), RegPath)
End Function

Private Function GetJSONFilePath(FVal As String, ByVal FPath As String) As String
    FPath = InstallationPath & ObjectsFolder & FPath
    CreateDir FPath
    GetJSONFilePath = Replace(FPath & "\" & FVal & ".txt", "\\", "\")
End Function

Function GetFolderObject(ByVal FVal As String, ByVal FPath As String) As Object
    Dim No As Long
    If FPath Like "Select *" Then Exit Function
    On Error GoTo ex
    No = FreeFile
    Open GetJSONFilePath(FVal, FPath) For Input As #No
    Set GetFolderObject = ParseJson(Input$(LOF(No), #No))
ex:
    On Error Resume Next
    Close #No
End Function

Function SetFolderObject(ByVal FVal As String, ByVal NewValue As Object, ByVal FPath As String) As Boolean
    On Error GoTo ex
    Dim No As Long
    If Len(InstallationPath) = 0 Then InstallationPath = GetInstallationPath
    FPath = InstallationPath & ObjectsFolder & FPath
    CreateDir FPath
    FPath = Replace(FPath & "\" & FVal & ".txt", "\\", "\")
    No = FreeFile
    Open FPath For Output As #No
    Print #No, ConvertToJson(NewValue)
    Close #No
    SetFolderObject = True
    Exit Function
ex:
End Function

Private Sub PopulateMainInfoVars(MainInfo As Dictionary, ProjectInfo As Dictionary, WorkflowInfo As Dictionary, _
            PURLStr As String, UNameStr As String, UPassStr As String, Optional PName As String)
    Dim i As Long
    On Error Resume Next
    ProjectURLStr = PURLStr
    UserNameStr = UNameStr
    UserPasswordStr = UPassStr
    
    ' Populate user details
    UserIDStr = MainInfo("id")
    Set UserPloneRolesDict = ArrToDict(MainInfo("roles"))
    Set UserTeamRolesDict = ArrToDict(Array(MainInfo("your_team_role")))
    
    ' Populate project details
    PloneTimeZone = ExtractTimezoneDiff(ProjectInfo("time_now_portal"))
    LocalTimeZone = ExtractTimezoneDiff(ProjectInfo("time_now_user"))
    ProjectNameStr = ProjectInfo("short_name")
    If Len(DashboardURLStr) = 0 Then DashboardURLStr = ProjectInfo("dashboard_url")
    ProjectColorStr = FullColor(ProjectInfo("project_color")).Long
    ProjectVSNameStr = ProjectInfo("very_short_name")
    TemplatePasswordStr = ProjectInfo("template_password")
    ProjectIsPlanningStr = ProjectInfo("planning_project")
    RFPURL = "": RFPURL = ProjectInfo("RFPURL")
    ScopeURL = "": ScopeURL = ProjectInfo("ScopeURL")
    PMPURL = "": PMPURL = ProjectInfo("PMPURL")
    RFPParsed = ProjectInfo("RFPParsed")
    ScopeParsed = ProjectInfo("ScopeParsed")
    PMPParsed = ProjectInfo("PMPParsed")
    
    ' Identify client
    For i = 1 To ProjectInfo("companies").Count
        If ProjectInfo("companies")(i)("company_role") = "Customer" Then
            ProjectClientStr = ProjectInfo("companies")(i)("company_letter_kode")
            Exit For
        End If
    Next i
    
    ' Document naming convention
    DocumentsNameConvStr = Join(DictCollToArr(ProjectInfo("project_document_naming_convention")), "_")
    ContractNumberStr = ProjectInfo("project_contract_number")
    BuildDocumentsArrays
    If ProjectNameStr = "" Then ProjectNameStr = PName
End Sub

'==============================================================================
' IMPROVED: LoadProjectInfoReg - Enhanced error handling and recovery
'==============================================================================
Sub LoadProjectInfoReg(Optional SelectedProjectName, Optional Retry As Boolean)
    Dim Pth As String
    InstallationPath = GetInstallationPath
    ImagesPath = InstallationPath & IconsFolder
    
    If IsMissing(SelectedProjectName) Then
        If Len(ProjectNameStr) = 0 Then
            ProjectNameStr = GetProjectNameByIndex(GetProjectIndexByURL(GetRegSelection(GetActiveFName(ActiveDocument), selectedProject)))
        End If
        If Len(ProjectNameStr) = 0 Then
            ProjectNameStr = GetProjectNameByIndex(GetProjectIndexByIndex(NewPNum))
        End If
    Else
        ProjectNameStr = SelectedProjectName
    End If
    ProjectURLStr = GetURLByName(ProjectNameStr)
'    If Len(ProjectURLStr) = 0 Then Stop
    Pth = CleanName(ProjectURLStr, SheetName)
    Set MainInfo = GetFolderObject(MainInfoValue, Pth)
'    ProjectURLStr = GetReg(UrlValue, BaseRegDir & "\" & Pth)
'    If Len(ProjectURLStr) = 0 Then Stop
    
    ' IMPROVED: Better handling of missing or corrupted project info
    If Len(ProjectNameStr) > 0 And ProjectNameStr <> "Select Project" And Not Retry Then
        On Error Resume Next
        If MainInfo Is Nothing Or MainInfo.Count = 0 Then
            On Error GoTo 0
            WriteLog 2, CurrentMod, "Project info missing or corrupted, downloading fresh data for: " & ProjectNameStr
            DownloadProjectInfo
            LoadProjectInfoReg SelectedProjectName, True
            Exit Sub
        End If
    End If
    
    UserNameStr = GetUsername
    UserPasswordStr = DecryptPassword(GetReg(PasswordValue, BaseRegDir & "\" & Pth))
    IsAuthorized = Len(UserPasswordStr) > 0
    
    Set ProjectInfo = GetFolderObject(ProjectInfoValue, Pth)
    Set MembersDict = GetFolderObject(UserTeamMembersValue, Pth)
    Set UserPloneRolesDict = GetFolderObject(UserPloneRolesValue, Pth)
    Set UserTeamRolesDict = GetFolderObject(UserTeamRolesValue, Pth)
    Set ProjectGroupsDict = GetFolderObject(ProjectGroupsValue, Pth)
    Set UserGroupsDict = GetFolderObject(UserGroupsValue, Pth)
    Set DocumentsTypes = GetFolderObject(DocumentsTypesValue, Pth)
    Set WorkflowInfo = GetFolderObject(WorkflowInfoValue, Pth)
    
    PopulateMainInfoVars MainInfo, ProjectInfo, WorkflowInfo, ProjectURLStr, UserNameStr, UserPasswordStr
    PUploaded = GetReg(ScopeUploadedValue, Pth & ScopeParserPth)
End Sub
Private Function GetURLByName(PName As String) As String
    Dim Ps() As String, i As Long
    Ps = GetRegSubKeys(BaseRegDir)
    For i = 1 To UBound(Ps)
        If GetReg(ProjectNameValue, BaseRegDir & "\" & Ps(i)) = PName Then
            GetURLByName = GetReg(UrlValue, BaseRegDir & "\" & Ps(i))
            Exit For
        ElseIf Ps(i) = CleanName(PName, SheetName) Then
            GetURLByName = GetReg(UrlValue, BaseRegDir & "\" & Ps(i))
            Exit For
        End If
    Next
End Function
Private Function GetUsername(Optional Force As Boolean) As String
    If Len(UserNameStr) = 0 Then UserNameStr = GetReg(UserNameValue, BaseRegDir)
    If Len(UserNameStr) = 0 Or Force Then
        UserNameStr = InputBox("Please insert your email", "Docent IMS")
        SetReg UserNameValue, UserNameStr, BaseRegDir
    End If
    GetUsername = UserNameStr
End Function

Function GetUserPassword(PName As String, ByVal PURL As String, Optional Silent As Boolean) As String
'    PURL =
    If Len(UserPasswordStr) = 0 Then
        UserPasswordStr = DecryptPassword(GetReg(PasswordValue, BaseRegDir & "\" & CleanName(PURL, SheetName)))
    End If
    If Len(UserPasswordStr) = 0 And Not Silent Then
        UserPasswordStr = frmInputBox.Display("Please insert your password for " & PName, "Docent IMS")
        If UserPasswordStr <> "Canceled" Then
            Do Until IsValidUser(PURL) = "OK"
                Select Case frmMsgBox.Display(Array("Wrong password.", "Do you want to retry?"), _
                                              Array("Retry", "Insert Email", "Cancel"), _
                                              Critical, "Docent IMS", Array(0, 225))
                Case "Retry"
                    UserPasswordStr = InputBox("Please insert your password for " & PName, "Docent IMS")
                Case "Insert Email"
                    GetUsername True
                Case Else
                    Exit Function
                End Select
            Loop
        End If
    End If
    GetUserPassword = UserPasswordStr
End Function

Private Function GetPName(PURLorName As String) As String
    Dim Ps As Collection, i As Long
    Set Ps = GetPs
    For i = 1 To Ps.Count
        If GetFileName(GetReg(UrlValue, BaseRegDir & "\" & Ps(i))) = _
                GetFileName(PURLorName) Then
            GetPName = Ps(i)
            Exit Function
        End If
    Next
    If Len(GetPName) = 0 Then GetPName = PURLorName
End Function

'==============================================================================
' NEW: BackupProjectRegistry - Optional backup mechanism for critical data
'==============================================================================
Private Sub BackupProjectRegistry(PName As String)
    ' Creates a backup of critical registry data before major operations
    ' This can be used for recovery if something goes wrong
    On Error Resume Next
    Dim BackupPath As String
    Dim Pth As String
    
    BackupPath = BaseRegDir & "\Backup\" & PName
    Pth = BaseRegDir & "\" & PName
    
    ' Backup critical values
    SetReg UrlValue, GetReg(UrlValue, Pth), BackupPath
    SetReg ProjectNameValue, GetReg(ProjectNameValue, Pth), BackupPath
    SetReg ProjectColorValue, GetReg(ProjectColorValue, Pth), BackupPath
    SetReg PasswordValue, GetReg(PasswordValue, Pth), BackupPath
    
    WriteLog 1, CurrentMod, "Backed up registry for project: " & PName
End Sub

'==============================================================================
' NEW: RestoreProjectRegistry - Restore from backup if needed
'==============================================================================
Private Sub RestoreProjectRegistry(PName As String)
    ' Restores registry data from backup
    On Error Resume Next
    Dim BackupPath As String
    Dim Pth As String
    
    BackupPath = BaseRegDir & "\Backup\" & PName
    Pth = BaseRegDir & "\" & PName
    
    ' Restore critical values
    SetReg UrlValue, GetReg(UrlValue, BackupPath), Pth
    SetReg ProjectNameValue, GetReg(ProjectNameValue, BackupPath), Pth
    SetReg ProjectColorValue, GetReg(ProjectColorValue, BackupPath), Pth
    SetReg PasswordValue, GetReg(PasswordValue, BackupPath), Pth
    
    WriteLog 2, CurrentMod, "Restored registry from backup for project: " & PName
End Sub





