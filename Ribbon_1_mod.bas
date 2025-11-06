Attribute VB_Name = "Ribbon_1_mod"
Option Explicit
Option Compare Text
Private Const CurrentMod = "Ribbon_mod" '"" '

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            destination As Any, Source As Any, ByVal length As LongPtr)

    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
    
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
            destination As Any, source As Any, ByVal length As Long)
#End If

Private mRibbonUI As IRibbonUI
Private mEvents As New WrdEvents

'Ribbon Handlers
Sub rDocentIMS_OnLoad(ByVal Ribbon As IRibbonUI)
'    Dim isDocentDocument As Boolean
    On Error Resume Next
    WriteLog 1, CurrentMod, "RibbonDocentIMS_OnLoad"
'    isDocentDocument = getproperty(pIsDocument)
    'If GetRibbonID = 0 Then Ribbon.ActivateTab "IdTabDocentIMS"
    SaveRibbonID ObjPtr(Ribbon)
    Set mRibbonUI = Ribbon
    Set SOWsColl = New Collection
    'RefreshRibbon
'    ResetRibbonGroups
    Set mEvents.App = Word.Application
    GetPs
End Sub
'Sub ActivateDocentRibbon()
'    On Error Resume Next
'    If mRibbonUI Is Nothing Then RedefineRibbon 'RedefineRibbon False
'    mRibbonUI.ActivateTab "IdTabDocentIMS"
''    Sleep 1000
'    DoEvents
''    Debug.Print "here"
'    mRibbonUI.ActivateTabMso "TabHome"
''    Sleep 1000
'    DoEvents
''    Debug.Print "here2"
'End Sub
Sub ActivateDocentRibbon()
'    DoEvents
    mRibbonUI.ActivateTab "IdTabDocentIMS"
'    DoEvents
End Sub
Sub RefreshRibbon(Optional Manually As Boolean)
    If Not Application.Visible Then Exit Sub
    WriteLog 1, CurrentMod, "RefreshRibbon"
   ' Stop
'    BusyRibbon = True
    BusyRibbon = False
    On Error Resume Next
    Dim IsDocentDocument As Boolean
    If Manually Then LoadProjects True
    If mRibbonUI Is Nothing Then RedefineRibbon
    IsDocentDocument = GetProperty(pIsDocument)
    'Stop
    'DarkMode = IsDarkModeSelected
    If Not Manually And Documents.Count > 0 And IsDocentDocument Then
        ActivateDocentRibbon ': DoEvents
        Do While Err.Number = 5
            ActiveDocument.Windows(1).Activate
            Sleep 100
            DoEvents: Sleep 50
            On Error Resume Next
            Err.Clear
            ActivateDocentRibbon ': DoEvents
        Loop
        DoEvents
    ElseIf Not Manually Then
''        If Not isDocentDocument Then
'            mRibbonUI.ActivateTabMso "TabHome"
''            DoEvents
''        End If
    End If
    LoadProjectInfoReg
    RefreshProject
    RefreshRibbonGroups
End Sub

'==============================================================================
' OLD ShowHelp - Kept for reference
'==============================================================================
'Sub ShowHelp()
'    Dim i As Long
'    On Error Resume Next
'    i = 1
'    If PrjMgr Then i = 2
''    If FirstPage Then i = 3
''    If i = 1 Then
'        If Mid(GetNeverHelpAgain, 3, 1) <> 1 Then
'            If Mid(GetHelpShown, 3, 1) <> 1 Then
'                Dim FName As String, Doc As Document
'                BusyRibbon = True
'                SetHelpShown 3
'                DashboardURLStr = GetRegDashboardURL
'                'FName = DownloadAPIFile("/docent-help/draftfirstpage.docx", mURL:=DashboardURLStr)
'                If Len(FName) Then
'                    Set Doc = Documents.Open(FName)
'                    If Not Doc Is Nothing Then Doc.Saved = True
'                End If
'                BusyRibbon = False
'            End If
'        End If
''    End If
'    If Mid$(GetNeverHelpAgain, i, 1) <> 1 Then
'        If Mid$(GetHelpShown, i, 1) <> 1 Then
'            If Not IsLoaded("frmHelpPics") Then frmHelpPics.Display i
'        End If
'    End If
''    If i = 1 Then OpenFirstPage
'
'
''    If PrjMgr Then
''        If Mid(GetNeverHelpAgain, 2, 1) <> 1 Then
''            If Mid(GetHelpShown, 2, 1) <> 1 Then frmHelpPics.Display "010"
''        End Select
'''        If Not (GetNeverHelpAgain Or GetHelpShown) Then frmHelpPics.Show
''    Else
''        Select Case GetNeverHelpAgain
''        Case 3, 1
''        Case Else
''            Select Case GetHelpShown
''            Case 3, 1
''            Case Else
''                frmHelpPics.Display 1
''            End Select
''        End Select
''    End If
''    Stop
'End Sub

'==============================================================================
' IMPROVED: ShowHelp - Cleaner logic with helper function
'==============================================================================
Sub ShowHelp()
    On Error Resume Next
    
    Dim helpType As Long
    helpType = IIf(PrjMgr, 2, 1)
    
    ' Check if web document help is needed (position 3)
    If ShouldShowHelp(3) Then
        'ShowWebDocHelp
    End If
    
    ' Check if role-specific help is needed
    If ShouldShowHelp(helpType) Then
        If Not IsLoaded("frmHelpPics") Then
            frmHelpPics.Display helpType
        End If
    End If
End Sub

'==============================================================================
' Helper function to check if help should be shown for a specific position
'==============================================================================
Private Function ShouldShowHelp(Position As Long) As Boolean
    ShouldShowHelp = (Mid$(GetNeverHelpAgain, Position, 1) <> "1") And _
                     (Mid$(GetHelpShown, Position, 1) <> "1")
End Function

'==============================================================================
' Helper function to show web document help
'==============================================================================
Private Sub ShowWebDocHelp()
    Dim FName As String, Doc As Document
    
    On Error Resume Next
    BusyRibbon = True
    SetHelpShown 3
    DashboardURLStr = GetRegDashboardURL
    
    FName = DownloadAPIFile("/docent-help/draftfirstpage.docx", mURL:=DashboardURLStr)
    
    If Len(FName) > 0 Then
        Set Doc = Documents.Open(FName)
        If Not Doc Is Nothing Then
            Doc.Saved = True
        End If
    End If
    
    BusyRibbon = False
End Sub

Sub RedefineRibbon(Optional UpdateSelections As Boolean = True)
    On Error Resume Next
    #If VBA7 Then
        Dim RibbonPtr As LongPtr
    #Else
        Dim RibbonPtr As Long
    #End If
    RibbonPtr = CLngPtr(GetRibbonID)

    If RibbonPtr = 0 Then
        WriteLog 3, CurrentMod, "Invalid Ribbon ID - cannot redefine"
        Exit Sub
    End If

    CopyMemory ByVal mRibbonUI, RibbonPtr, LenB(RibbonPtr)
    On Error Resume Next
    If UpdateSelections Then
        Set mEvents = New WrdEvents
        Set mEvents.App = Word.Application
        mEvents.IsBusy = True
        RefreshRibbon
    End If
End Sub

''Callback for TabSave getVisible
'Sub IsSaveTabVisible(control As IRibbonControl, ByRef returnedVal)
'    Stop
'    returnedVal = False
'End Sub
'
''Callback for FileSave getVisible
'Sub IsSaveButtonVisible(control As IRibbonControl, ByRef returnedVal)
'    Stop
'    returnedVal = False
'End Sub
'
''Callback for FileSaveAs getVisible
'Sub IsSaveAsButtonVisible(control As IRibbonControl, ByRef returnedVal)
'    Stop
'    returnedVal = False
'End Sub

'Refreshers
Sub Invalidate(Optional ID As String, Optional IsMSO As Boolean)
    'WriteLog 1, CurrentMod, "Invalidate", Id
    'On Error Resume Next
    If mRibbonUI Is Nothing Then RedefineRibbon
    If Len(ID) = 0 Then
        mRibbonUI.Invalidate
    ElseIf IsMSO Then
        mRibbonUI.InvalidateControlMso ID
    Else
        mRibbonUI.InvalidateControl ID
    End If
    DoEvents ': Sleep 5
End Sub
Sub RefreshRibbonGroups()
    Invalidate "IdToggleButtonMgrMode"
    RefreshTasksGroup
    RefreshNotificationsGroup
    RefreshDocumentsGroups
    RefreshTemplatesGroup
    RefreshPlanningGroup
    RefreshScopeGroup
    RefreshRFPGroup
'    RefreshTaskTrafficGroup
    Invalidate "IdGroupCommandStatements"
    Invalidate "IdGroupCreate"
    Invalidate "IdGroupTeam"
    Invalidate "IdGroupPMP"
    Invalidate "IdGroupMSP"
    Invalidate "IdButtonSaveLoc"
    Invalidate "IdMenuHelp"
    Invalidate "IdSplitButtonHelp"
'    DoEvents
End Sub
Sub RefreshScopeGroup()
    Invalidate "IdButtonCreateScope"
    Invalidate "IdSplitButtonParseScope"
    Invalidate "IdButtonScopeAddTop"
    Invalidate "IdButtonScopeAddSame"
    Invalidate "IdButtonScopeAddSub"
    Invalidate "IdButtonScopeRevesion"
    Invalidate "IdButtonScopeUnlock"
    Invalidate "IdButtonScopeCancel"
    Invalidate "IdGroupScope"
'    DoEvents
End Sub
Sub RefreshRFPGroup()
    Invalidate "IdSplitButtonRFPUpload"
    Invalidate "IdButtonOpenRFP"
    Invalidate "IdButtonRFPUpload"
    Invalidate "IdButtonRFPUpload0"
    Invalidate "IdButtonRFPUpload1"
'    Invalidate "IdButtonRFPBrowse"
    Invalidate "IdGroupRFP"
'    DoEvents
End Sub
Sub RefreshTasksGroup()
'    Set TasksDict = New Dictionary
    Invalidate "IdGroupTasks"
    If GetVisibleGroup("IdGroupTasks") Then
        RefreshTrafficGroup "Tasks"
        Invalidate "IdButtonCreateTask"
    End If
'    DoEvents
End Sub
Sub RefreshNotificationsGroup()
    Set NotifsDict = New Dictionary
    Invalidate "IdGroupNotifications"
    If GetVisibleGroup("IdGroupNotifications") Then
        RefreshTrafficGroup "Notifications"
        Invalidate "IdButtonCreateNotification"
    End If
'    DoEvents
End Sub
Sub RefreshTrafficGroup(ItemName As String)
'    Invalidate "IdGroup" & ItemName & "Traffic"
    Invalidate "IdButton" & ItemName & "Green"
    Invalidate "IdButton" & ItemName & "Yellow"
    Invalidate "IdButton" & ItemName & "Red"
'    DoEvents
End Sub
Sub RefreshRibbonPColor()
    Invalidate "IdButtonPColor"
End Sub
Sub RefreshProject()
    Invalidate ""
'    Invalidate "IdButtonNotificationsGetVisible"
'    RefreshTaskTrafficGroup
'    Invalidate "IdCheckBoxPlanning"
'    Invalidate "IdDDProject"
'    Invalidate "IdGroupProject"
'    Invalidate "IdDDDocument"
'    Invalidate "IdDDMeetingDoc"
'    Invalidate "IdDDTemplate"
'    Invalidate "IdButtonHelp"
'    DoEvents
End Sub
Sub RefreshDocumentsGroups()
    RefreshDocumentGroup
    RefreshMeetingDocButtons
    Invalidate "IdGroupScope"
'    DoEvents
End Sub
Sub RefreshPlanningGroup()
    Invalidate "IdButtonPlanningOpen"
    Invalidate "IdButtonPlanningUpload"
    Invalidate "ReviewNewComment"
    Invalidate "TextHighlightColorPicker"
    Invalidate "Spelling"
    Invalidate "IdButtonPlanningCancel"
    Invalidate "IdGroupPlanning"
'    DoEvents
End Sub
Sub RefreshDocumentGroup()
    Invalidate "IdButtonDocumentCreate"
    Invalidate "IdButtonDocumentOpen"
    Invalidate "IdButtonDocumentSave"
    Invalidate "IdSplitButtonDocumentSave"
    Invalidate "IdButtonDocumentSaveAs1"
    Invalidate "IdButtonDocumentSaveAs2"
    Invalidate "IdButtonDocumentSaveAs3"
    Invalidate "IdButtonDocumentCancel"
    Invalidate "IdButtonDocumentState"
    Invalidate "IdToggleButtonDocumentHide"
    Invalidate "IdGroupDocument"
'    DoEvents
End Sub
Sub RefreshMeetingDocButtons()
    Invalidate "IdButtonMeetingDocCreate"
    Invalidate "IdButtonMeetingDocOpen"
    Invalidate "IdButtonMeetingDocSave"
    Invalidate "IdSplitButtonMeetingDocSave"
    Invalidate "IdButtonMeetingDocSaveAs1"
    Invalidate "IdButtonMeetingDocSaveAs2"
    Invalidate "IdButtonMeetingDocSaveAs3"
    Invalidate "IdButtonMeetingDocCancel"
    Invalidate "IdButtonMeetingDocState"
    Invalidate "IdToggleButtonMeetingDocHide"
    Invalidate "IdButtonCreateMeeting"
    Invalidate "IdGroupMeetingDoc"
'    DoEvents
End Sub
Sub RefreshTemplatesGroup()
    Invalidate "IdButtonTemplateModify"
    Invalidate "IdButtonTemplateOpen"
    Invalidate "IdButtonTemplateCancel"
    Invalidate "IdGroupTemplate"
'    DoEvents
End Sub

Function BasicRibbonChecks() As Boolean
'    If Not IsProjectSelected(True) Then RefreshRibbon
    BasicRibbonChecks = Not BusyRibbon And Application.Visible 'Not CodeIsRunning And
'    If Not BasicRibbonChecks Then mRibbonUI.ActivateTab "IdTabDocentIMS"
    'If IsProjectSelected
'    If CodeIsRunning Then Exit Function
End Function

'==============================================================================
' OLD Callback Handlers - Kept for reference
'==============================================================================
'Sub rGetVisible(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetVisible", control.ID)
'End Sub
'Sub rGetLabel(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetLabel")
'End Sub
'Sub rGetEnabled(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetEnabled")
'End Sub
'Sub rGetSupertip(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetSupertip")
'End Sub
'Sub rGetScreentip(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetScreentip")
'End Sub
'Sub rGetPressed(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetPressed")
'End Sub
'Sub rGetImage(control As IRibbonControl, ByRef returnedVal)
'    On Error Resume Next
'    If Not BasicRibbonChecks Then Exit Sub
'    Set returnedVal = Application.Run(control.ID & "GetImage")
'End Sub
'Sub rGetImageName(control As IRibbonControl, ByRef returnedVal)
'    On Error Resume Next
'    If Not BasicRibbonChecks Then Exit Sub
'    returnedVal = Application.Run(control.ID & "GetImage")
'    If Len(returnedVal) = 0 Then
'        Set returnedVal = Application.Run(control.ID & "GetImage")
'        If Not TypeName(returnedVal) = "Picture" Then
'            If IsEmpty(returnedVal) = 0 Then Set returnedVal = Nothing
'            If Len(returnedVal) = 0 Then Set returnedVal = Nothing
'        End If
'    End If
'End Sub
'Sub rDDGetItemCount(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetItemCount")
'End Sub
'Sub rDDGetImage(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    Set returnedVal = Application.Run(control.ID & "GetImage")
'End Sub
'Sub rDDGetVisible(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetVisible")
'End Sub
'Sub rDDGetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetSelectedItemIndex")
'End Sub
'Sub rDDGetItemLabel(control As IRibbonControl, Index As Integer, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    returnedVal = Application.Run(control.ID & "GetItemLabel", Index)
'End Sub
'Sub rDDGetItemImage(control As IRibbonControl, Index As Integer, ByRef returnedVal)
'    If Not BasicRibbonChecks Then Exit Sub
'    On Error Resume Next
'    Set returnedVal = Application.Run(control.ID & "GetItemImage", Index)
'End Sub

'==============================================================================
' IMPROVED Callback Handlers - Enhanced error handling and logging
'==============================================================================
Sub rGetVisible(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetVisible", control.ID)
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get visible for " & control.ID & ": " & Err.Description
    returnedVal = False  ' Safe default
End Sub

Sub rGetLabel(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetLabel")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get label for " & control.ID & ": " & Err.Description
    returnedVal = control.ID  ' Fallback to control ID as label
End Sub

Sub rGetEnabled(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetEnabled")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get enabled for " & control.ID & ": " & Err.Description
    returnedVal = False  ' Safe default - disable control
End Sub

Sub rGetSupertip(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetSupertip")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get supertip for " & control.ID & ": " & Err.Description
    returnedVal = ""  ' Empty supertip on error
End Sub

Sub rGetScreentip(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetScreentip")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get screentip for " & control.ID & ": " & Err.Description
    returnedVal = ""  ' Empty screentip on error
End Sub

Sub rGetPressed(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetPressed")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get pressed state for " & control.ID & ": " & Err.Description
    returnedVal = False  ' Safe default - not pressed
End Sub

Sub rGetImage(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo ErrorHandler
    If Not BasicRibbonChecks Then Exit Sub
    Set returnedVal = Application.Run(control.ID & "GetImage")
    Exit Sub
    
ErrorHandler:
    'Stop
    WriteLog 3, CurrentMod, "Failed to get image for " & control.ID & ": " & Err.Description
    Set returnedVal = Nothing  ' No image on error
End Sub

Sub rGetImageName(control As IRibbonControl, ByRef returnedVal)
    On Error GoTo ErrorHandler
    If Not BasicRibbonChecks Then Exit Sub
    returnedVal = Application.Run(control.ID & "GetImage")
    If TypeName(returnedVal) = "Picture" Then
        rGetImage control, returnedVal
'        Set returnedVal = Application.Run(control.ID & "GetImage")
'        If Not TypeName(returnedVal) = "Picture" Then
'            If IsEmpty(returnedVal) = 0 Then Set returnedVal = Nothing
'            If Len(returnedVal) = 0 Then Set returnedVal = Nothing
'        End If
    ElseIf Len(returnedVal) = 0 Then
        rGetImage control, returnedVal
    End If
    Exit Sub
    
ErrorHandler:
'    Stop
'    Resume
    WriteLog 3, CurrentMod, "Failed to get image name for " & control.ID & ": " & Err.Description
    Set returnedVal = Nothing
'    returnedVal = ""  ' Empty string on error
End Sub

Sub rDDGetItemCount(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetItemCount")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get item count for " & control.ID & ": " & Err.Description
    returnedVal = 0  ' Empty dropdown on error
End Sub

Sub rDDGetImage(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    Set returnedVal = Application.Run(control.ID & "GetImage")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get dropdown image for " & control.ID & ": " & Err.Description
    Set returnedVal = Nothing
End Sub

Sub rDDGetVisible(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetVisible")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get dropdown visible for " & control.ID & ": " & Err.Description
    returnedVal = False
End Sub

Sub rDDGetSelectedItemIndex(control As IRibbonControl, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetSelectedItemIndex")
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get selected index for " & control.ID & ": " & Err.Description
    returnedVal = 0  ' Default to first item
End Sub

Sub rDDGetItemLabel(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    returnedVal = Application.Run(control.ID & "GetItemLabel", Index)
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get item label for " & control.ID & " at index " & Index & ": " & Err.Description
    returnedVal = "Item " & Index  ' Generic fallback label
End Sub

Sub rDDGetItemImage(control As IRibbonControl, Index As Integer, ByRef returnedVal)
    If Not BasicRibbonChecks Then Exit Sub
    On Error GoTo ErrorHandler
    Set returnedVal = Application.Run(control.ID & "GetItemImage", Index)
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "Failed to get item image for " & control.ID & " at index " & Index & ": " & Err.Description
    Set returnedVal = Nothing
End Sub

'Controls Handlers
Function FileSaveGetVisible(ID)
    FileSaveGetVisible = False
End Function
Function TabSaveGetVisible(ID)
    TabSaveGetVisible = False
End Function
'Sub backstageOnShow(contextObject As Object)
''    Debug.Print "FileSave"
''    Invalidate "FileSave" 'contextObject.Document
''    Invalidate "TabSave" 'contextObject.Document
'''    Debug.Print TypeName(contextObject)
'End Sub
'Sub backstageOnHide(contextObject As Object)
''    Invalidate contextObject.Document
''    Debug.Print TypeName(contextObject)
'End Sub
Sub rCheckBoxOnAction(control As IRibbonControl, Pressed As Boolean)
    If Not BasicRibbonChecks Then Exit Sub
    On Error Resume Next
    BusyRibbon = True
    CodeIsRunning = True
    Application.Run control.ID & "OnAction", Pressed
    CodeIsRunning = False
    BusyRibbon = False
End Sub
Sub rToggleButtonOnAction(control As IRibbonControl, Pressed As Boolean)
    If Not BasicRibbonChecks Then Exit Sub
    On Error Resume Next
    CodeIsRunning = True
    BusyRibbon = True
    Application.Run control.ID & "OnAction", Pressed
    CodeIsRunning = False
    BusyRibbon = False
End Sub
Sub rButtonOnAction(control As IRibbonControl)
    If Not BasicRibbonChecks Then Exit Sub
    On Error Resume Next
    CodeIsRunning = True
    BusyRibbon = True
    Application.Run control.ID & "OnAction"
    CodeIsRunning = False
    BusyRibbon = False
End Sub
Sub rDDOnAction(control As IRibbonControl, ID As String, Index As Integer)
'    If Not BasicRibbonChecks Then Exit Sub
'    CodeIsRunning = True
    On Error Resume Next
    Application.Run control.ID & "OnAction", Index ', id
'    CodeIsRunning = False
End Sub
'DocentIMS
Private Function IdButtonDocentIMSGetLabel(): IdButtonDocentIMSGetLabel = "Docent IMS": End Function
Private Sub IdButtonDocentIMSOnAction()
    If Application.UserName = "Abdallah Ali" Then
        On Error Resume Next
        RefreshRibbon True
    Else
        frmAbout.Show
    End If
End Sub
'Projects
Private Function IdButtonPColorGetImage()
    On Error Resume Next
'    Set IdButtonPColorGetImage = Generate1ColorBMP(FullColor(ProjectColorStr).Long)
    'If ProjectName(NewPNum) <> "Select Project" Then
'    Set IdButtonPColorGetImage = Generate1ColorBMP(FullColor(ProjectColor(GetProjectIndexByName(ProjectNameStr, ProjectName))).Long)
'    ProjectNameStr = GetProjectNameByIndex(NewPNum)
'    If ProjectNameStr Like "Select *" Then Exit Function
    Set IdButtonPColorGetImage = Generate1ColorBMP(FullColor(ProjectColor(NewPNum)).Long)
End Function
'Private Function IdButtonPColorGetImage(): On Error Resume Next: Set IdButtonPColorGetImage = LoadPictureGDI(PColorJPG(GetProjectIndexByIndex(NewPNum))): End Function
Function IdButtonPColorGetVisible(ID As String)
    On Error Resume Next 'If NewPNum = 0 Then Exit Function
    ProjectNameStr = GetProjectNameByIndex(GetProjectIndexByIndex(NewPNum)) ' GetProjectNameByIndex(NewPNum)
    If ProjectNameStr Like "Select *" Then ProjectNameStr = ""
    IdButtonPColorGetVisible = Len(ProjectNameStr)
End Function
Private Function IdButtonPColorOnAction()
    On Error Resume Next
    GoToLink ProjectURLStr
End Function

Private Function IdDDProjectGetItemCount(): IdDDProjectGetItemCount = GetProjectsCount: End Function
Private Function IdDDProjectGetItemLabel(Index As Integer)
    WriteLog 1, CurrentMod, "IdDDProjectGetItemLabel"
    On Error GoTo ex
    If GetProperty(pIsDocument) Then
         IdDDProjectGetItemLabel = GetProperty(pPName)
    Else
ex:
        If PlanningOnly Then IdDDProjectGetItemLabel = PlanningProjectName(Index) Else IdDDProjectGetItemLabel = NoPlanningProjectName(Index)
    End If
End Function
Private Function IdDDProjectGetSelectedItemIndex(): IdDDProjectGetSelectedItemIndex = GetSelectedProjectIndex: End Function
Private Function IdDDProjectGetImage()
    On Error Resume Next
''    Set IdDDProjectGetImage = LoadPictureGDI("D:\Ongoing\23-06-15 - Wayne Glover (Word)\test.emf")
'    Dim FName As String
'    FName = DownloadAPIFile("/images/word-box.jpg/@@download", False)
'    Set IdDDProjectGetImage = LoadPictureGDI(PColorJPG(PNum)) 'LoadPictureGDI(FName) '
End Function
Private Function IdDDProjectGetItemImage(Index As Integer): IdDDProjectGetItemImage = ProjectColor(GetProjectIndexByIndex(Index)): End Function
Private Sub IdDDProjectOnAction(Index As Integer): SetSelectedProjectIndex Index: End Sub
'Planning Only
Private Sub IdCheckBoxPlanningOnAction(Pressed As Boolean)
    PlanningOnly = Pressed
    IdButtonRefreshOnAction
'    SetCursor LoadCursorW(0&, IDC_WAIT)
'    UpdateAllProjectsInfo
'    RefreshRibbon True
'    Invalidate "IdDDProject"
'    SetCursor LoadCursorW(0&, IDC_ARROW)
'    frmMsgBox.Display "All projects were updated."
End Sub
Private Function IdCheckBoxPlanningGetPressed(): IdCheckBoxPlanningGetPressed = PlanningOnly: End Function
Private Function IdCheckBoxPlanningGetVisible(ID As String): IdCheckBoxPlanningGetVisible = GetVisibleGroup(ID): End Function
'Create Group
Function IdGroupCreateGetVisible(ID As String): IdGroupCreateGetVisible = GetVisibleGroup(ID): End Function
'PostItNote
Sub IdButtonCreatePostItNoteOnAction(): On Error Resume Next: frmCreatePostItNote.Show: End Sub
'Team
Function IdGroupTeamGetVisible(ID As String): IdGroupTeamGetVisible = GetVisibleGroup(ID): End Function
Sub IdButtonTeamMembersOnAction(): frmTeam.Show: End Sub
'PMP
Function IdGroupPMPGetVisible(ID As String): IdGroupPMPGetVisible = GetVisibleGroup(ID): End Function
Sub IdButtonPMPOnAction(): MsgBox "Under construction...", vbExclamation, "DocentIMS": End Sub


'Help
'Private Function IdButtonRefreshGetImage()
'    Dim DarkMode As Boolean
'    DarkMode = IsDarkModeSelected
'    Set IdButtonRefreshGetImage = MLoadPictureGDI.LoadPictureGDI("D:\Ongoing\23-06-15 - Wayne Glover (Word)\Old\Icons\Ribbon Icons\Refresh" & IIf(DarkMode, "B", "W") & ImagesExtension)
'End Function
Private Sub IdButtonRefreshOnAction()
    On Error Resume Next
    frmMsgBox.Display "Please wait while projects are being updated...", Array(), None, ShowModal:=vbModeless
    RefreshRibbon True
    'UpdateAllProjectsInfo True
    SetRegSelection GetFileName(GetActiveFName(ActiveDocument)), projectName(0), selectedProject
    NewPNum = 0
    ProjectSelected ActiveDocument, 0, True ',pnum, True
    RefreshRibbon 'True
    Unload frmMsgBox
    frmMsgBox.Display "All projects were updated."
End Sub
Private Function IdButtonHelpGetEnabled(): IdButtonHelpGetEnabled = NewPNum > 0 And IsAuthorized: End Function
Private Sub IdButtonHelpOnAction()
    WriteLog 1, CurrentMod, "IdButtonHelpOnAction", "Help button Clicked"
    On Error Resume Next
    If IsProjectSelected Then GoToLink HelpURL   'ProjectURLStr & "/help-files/word-help/word-help"
End Sub
Private Sub IdButtonHelp0OnAction(): IdButtonHelpOnAction: End Sub
Private Sub IdButtonConfigOnAction()
    On Error GoTo ex
    frmProjectsList.Show
    mRibbonUI.Invalidate
ex:
End Sub
Private Sub IdButtonSaveLocOnAction(): frmSaveLoc.Show: End Sub
Function IdButtonSaveLocGetEnabled()
    On Error Resume Next
    If UserGroupsDict Is Nothing Then GetMyGroupsDict
    IdButtonSaveLocGetEnabled = UserGroupsDict.Exists("PrjMgr")
End Function
Private Sub IdButtonFeedbackOnAction(): On Error Resume Next: frmFeedback.Show: End Sub
Private Function IdButtonFeedbackGetVisible(ID As String): IdButtonFeedbackGetVisible = Not PrjMgr: End Function
Private Function IdSplitButtonFeedbackGetVisible(ID As String): IdSplitButtonFeedbackGetVisible = PrjMgr: End Function
Private Sub IdButtonMgrFeedbackOnAction(): On Error Resume Next: frmFeedback.Show: End Sub
Private Sub IdButtonMgrFeedback0OnAction(): On Error Resume Next: frmFeedback.Show: End Sub
Private Sub IdButtonReviewFeedbackOnAction(): On Error Resume Next: GoToLink ProjectURLStr & "/feedback/": End Sub
Private Function IdButtonHelp0GetEnabled(): IdButtonHelp0GetEnabled = IdButtonHelpGetEnabled: End Function
Private Sub IdButtonFeedback2OnAction(): On Error Resume Next: frmFeedback.Show: End Sub
Private Function IdButtonFeedbackGetSupertip(): IdButtonFeedbackGetSupertip = IIf(NewPNum = 0, "A project must be selected to send feedback.", "Help us improve Docent."): End Function
Private Function IdButtonFeedback2GetSupertip(): IdButtonFeedback2GetSupertip = IdButtonFeedbackGetSupertip: End Function
Private Function IdButtonFeedbackGetEnabled(): IdButtonFeedbackGetEnabled = NewPNum > 0 And IsAuthorized: End Function
Private Function IdButtonFeedback2GetEnabled(): IdButtonFeedback2GetEnabled = IdButtonFeedbackGetEnabled: End Function
Private Sub IdButtonDocInfoOnAction(): ShowDocumentInfo: End Sub
Private Sub IdButtonAboutOnAction(): frmAbout.Show: End Sub
'==========
'Planning Documents
'==========
Private Function IdGroupPlanningGetVisible(ID As String): IdGroupPlanningGetVisible = GetVisibleGroup(ID): End Function
'Open
Function IdSplitButtonPlanningGetVisible(ID As String): IdSplitButtonPlanningGetVisible = GetButtonVisible(1): End Function
Private Sub IdButtonPlanningOpenOnAction()
    If Documents.Count = 0 Then Exit Sub
    UploadPlanningDocument ActiveDocument
End Sub
Private Sub IdButtonPlanningOpen0OnAction()
    If Documents.Count = 0 Then Exit Sub
    UploadPlanningDocument ActiveDocument
End Sub
Private Sub IdButtonPlanningBrowseOnAction(): OpenAsDocentDocument "Planning Document": End Sub
'Upload
Function IdButtonPlanningUploadGetVisible(ID As String): IdButtonPlanningUploadGetVisible = GetButtonVisible(5): End Function
Private Sub IdButtonPlanningUploadOnAction()
    If Documents.Count = 0 Then Exit Sub
    UploadPlanningDocument ActiveDocument
End Sub
'Private Function IdButtonPlanningUploadGetImage()
'    Dim DarkMode As Boolean
'    DarkMode = IsDarkModeSelected
'    Set IdButtonPlanningUploadGetImage = MLoadPictureGDI.LoadPictureGDI("D:\Ongoing\23-06-15 - Wayne Glover (Word)\Old\Icons\Ribbon Icons\Upload" & IIf(DarkMode, "B", "W") & ImagesExtension)
'End Function
'Highlight
Function TextHighlightColorPickerGetVisible(ID As String): TextHighlightColorPickerGetVisible = GetButtonVisible(5): End Function
'Comment
Function ReviewNewCommentGetVisible(ID As String): ReviewNewCommentGetVisible = GetButtonVisible(5): End Function
'Spelling
Function SpellingGetVisible(ID As String): SpellingGetVisible = GetButtonVisible(5): End Function
'Cancel
Function IdButtonPlanningCancelGetVisible(ID As String): IdButtonPlanningCancelGetVisible = GetButtonVisible(5): End Function
Private Sub IdButtonPlanningCancelOnAction(): CancelEditingDoc: End Sub
Private Sub IdButtonWebmailOnAction(): GoToLink "webmail." & Right$(ProjectURLStr, Len(ProjectURLStr) - InStr(ProjectURLStr, ".")): End Sub
Private Sub IdButtonCalendarOnAction(): GoToLink "/calendar/calendar": End Sub
Private Function IdGroupLinksGetVisible(ID As String): IdGroupLinksGetVisible = GetVisibleGroup(ID): End Function
