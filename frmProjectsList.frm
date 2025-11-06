VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectsList 
   Caption         =   "Projects Configuration"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   OleObjectBlob   =   "frmProjectsList.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjectsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=======================================================
' UserForm: frmProjectsList
' Purpose: Project configuration and management interface
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Provides interface for managing project configurations
'   including add, edit, remove, and update operations.
'   Displays list of configured projects and allows
'   direct editing via double-click.
'
' Dependencies:
'   - AC_Registry_mod (LoadProjects, RemovePFromReg, UpdateAllProjectsInfo)
'   - AC_Registry_mod (DownloadProjectInfo)
'   - AB_GlobalVars (projectName, projectURL, UserName, Password)
'   - frmLogin (project credentials form)
'   - frmMsgBox
'
' Public Interface:
'   None - Form is displayed via .Show
'
' Private Methods:
'   - btnAdd_Click - Add new project
'   - btnEdit_Click - Edit selected project
'   - btnRemove_Click - Remove selected project
'   - btnUpdate_Click - Update all projects
'   - lstProjects_Change - Handle list selection change
'   - lstProjects_DblClick - Handle list double-click
'   - EditProject - Edit specific project
'   - UpdateProject - Update specific project
'   - RefreshList - Refresh projects list
'   - SelectedPName - Get currently selected project name
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added detailed logging
'       * Added function documentation
'       * Removed commented dead code
'   v1.0 - Original version
'=======================================================

Option Explicit

Private Const CurrentMod As String = "frmProjectsList"

'=======================================================
' Event: btnAdd_Click
' Purpose: Handle Add button click - add new project
'
' Description:
'   Opens login form in "Add" mode to configure a new project.
'   Refreshes list after completion.
'
' Error Handling:
'   - Logs operation
'   - Handles form display errors
'   - Ensures list refresh
'=======================================================
Private Sub btnAdd_Click()
    Const PROC_NAME As String = "btnAdd_Click"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Add button clicked"
    
    frmLogin.Display "Add"
    Call RefreshList
    
    WriteLog 1, CurrentMod, PROC_NAME, "Add project completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to add project: " & Err.Description, vbExclamation, "Error"
    Call RefreshList
End Sub

'=======================================================
' Function: SelectedPName
' Purpose: Get currently selected project name from list
'
' Returns:
'   String - Selected project name, or empty string if none selected
'
' Error Handling:
'   - Returns empty string on error
'   - Logs errors
'=======================================================
Private Function SelectedPName() As String
    Const PROC_NAME As String = "SelectedPName"
    
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    For i = 0 To lstProjects.ListCount - 1
        If lstProjects.Selected(i) Then
            SelectedPName = lstProjects.List(i)
            WriteLog 1, CurrentMod, PROC_NAME, "Selected project: " & SelectedPName
            Exit Function
        End If
    Next i
    
    SelectedPName = ""
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    SelectedPName = ""
End Function

'=======================================================
' Event: btnEdit_Click
' Purpose: Handle Edit button click - edit selected project
'
' Description:
'   Opens edit dialog for currently selected project.
'
' Error Handling:
'   - Validates selection exists
'   - Logs operation
'   - Handles edit errors
'=======================================================
Private Sub btnEdit_Click()
    Const PROC_NAME As String = "btnEdit_Click"
    
    Dim selectedProject As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Edit button clicked"
    
    selectedProject = SelectedPName()
    
    If Len(selectedProject) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "No project selected"
        MsgBox "Please select a project to edit.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    Call EditProject(selectedProject)
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to edit project: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Event: btnRemove_Click
' Purpose: Handle Remove button click - remove selected project
'
' Description:
'   Confirms removal with user, then removes project from
'   registry and refreshes list.
'
' Error Handling:
'   - Validates selection exists
'   - Confirms with user before removal
'   - Logs operation
'   - Handles removal errors
'=======================================================
Private Sub btnRemove_Click()
    Const PROC_NAME As String = "btnRemove_Click"
    
    Dim selectedProject As String
    Dim userResponse As String
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Remove button clicked"
    
    selectedProject = SelectedPName()
    
    If Len(selectedProject) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "No project selected"
        MsgBox "Please select a project to remove.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Confirm removal
    userResponse = frmMsgBox.Display("Remove Project?", _
                                     Array("Remove", "Cancel"), _
                                     Exclamation, _
                                     "Docent IMS")
    
    If userResponse = "Remove" Then
        WriteLog 1, CurrentMod, PROC_NAME, "Removing project: " & selectedProject
        
        Call RemovePFromReg(selectedProject)
        Call RefreshList
        
        WriteLog 1, CurrentMod, PROC_NAME, "Project removed successfully"
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "Removal cancelled by user"
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Failed to remove project: " & errorMsg, vbCritical, "Error"
    Call RefreshList
End Sub

'=======================================================
' Event: btnUpdate_Click
' Purpose: Handle Update button click - update all projects
'
' Description:
'   Updates information for all configured projects by
'   downloading fresh data from their APIs.
'
' Error Handling:
'   - Logs operation
'   - Handles update errors
'   - Provides user feedback
'=======================================================
Private Sub btnUpdate_Click()
    Const PROC_NAME As String = "btnUpdate_Click"
    
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Update all button clicked"
    
    If lstProjects.ListCount = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "No projects to update"
        MsgBox "No projects are configured.", vbInformation, "No Projects"
        Exit Sub
    End If
    
    ' Update all projects
    Call UpdateAllProjectsInfo
    
    WriteLog 1, CurrentMod, PROC_NAME, "All projects updated"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Error updating projects: " & errorMsg, vbExclamation, "Error"
End Sub

'=======================================================
' Event: lstProjects_Change
' Purpose: Handle list selection change
'
' Description:
'   Updates button enabled states based on whether
'   a project is selected.
'
' Error Handling:
'   - Handles errors gracefully
'   - Ensures buttons have valid state
'=======================================================
Private Sub lstProjects_Change()
    Const PROC_NAME As String = "lstProjects_Change"
    
    Dim i As Long
    Dim hasSelection As Boolean
    
    On Error GoTo ErrorHandler
    
    hasSelection = False
    
    For i = 1 To lstProjects.ListCount
        If lstProjects.Selected(i - 1) Then
            hasSelection = True
            Exit For
        End If
    Next i
    
    ' Update button states
    btnRemove.Enabled = hasSelection
    btnEdit.Enabled = hasSelection
    btnUpdate.Enabled = (lstProjects.ListCount > 0)
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    ' Set safe defaults
    On Error Resume Next
    btnRemove.Enabled = False
    btnEdit.Enabled = False
    btnUpdate.Enabled = False
End Sub

'=======================================================
' Event: lstProjects_DblClick
' Purpose: Handle list double-click - edit project
'
' Description:
'   Provides quick access to edit functionality via
'   double-click on project in list.
'
' Error Handling:
'   - Validates selection
'   - Logs operation
'   - Handles edit errors
'=======================================================
Private Sub lstProjects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Const PROC_NAME As String = "lstProjects_DblClick"
    
    Dim selectedProject As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "List double-clicked"
    
    selectedProject = SelectedPName()
    
    If Len(selectedProject) > 0 Then
        Call EditProject(selectedProject)
    Else
        WriteLog 2, CurrentMod, PROC_NAME, "No project selected on double-click"
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to edit project: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Sub: EditProject
' Purpose: Edit specified project
'
' Parameters:
'   mPName - Project name to edit
'
' Description:
'   Finds project in array, opens login form with existing
'   credentials, and refreshes list after edit.
'
' Error Handling:
'   - Validates project exists in array
'   - Handles form display errors
'   - Ensures list refresh
'   - Validates array bounds
'=======================================================
Private Sub EditProject(ByVal mPName As String)
    Const PROC_NAME As String = "EditProject"
    
    Dim i As Long
    Dim foundProject As Boolean
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Editing project: " & mPName
    
    ' Validate arrays
    If Not IsArray(projectName) Then
        errorMsg = "Project array not initialized"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox "Project configuration error. Please restart.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Find project in array
    foundProject = False
    For i = 1 To UBound(projectName)
        If projectName(i) = mPName Then
            foundProject = True
            
            ' Open login form with existing credentials
            With frmLogin
                .Caption = mPName
                .tbPassword = Password(i)
                .tbURL = projectURL(i)
                .tbUser = UserName(i)
                .btnAdd.Caption = "Update"
                .Show
            End With
            
            Call RefreshList
            Me.Show
            
            WriteLog 1, CurrentMod, PROC_NAME, "Project edited successfully"
            Exit For
        End If
    Next i
    
    If Not foundProject Then
        errorMsg = "Project not found: " & mPName
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbExclamation, "Not Found"
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Failed to edit project: " & errorMsg, vbExclamation, "Error"
    Call RefreshList
End Sub

'=======================================================
' Sub: UpdateProject
' Purpose: Update specified project from API
'
' Parameters:
'   mPName - Project name to update
'
' Description:
'   Downloads fresh project information from API and
'   refreshes the list.
'
' Error Handling:
'   - Validates project exists
'   - Handles download errors
'   - Ensures list refresh
'=======================================================
Private Sub UpdateProject(ByVal mPName As String)
    Const PROC_NAME As String = "UpdateProject"
    
    Dim i As Long
    Dim foundProject As Boolean
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Updating project: " & mPName
    
    ' Validate arrays
    If Not IsArray(projectName) Then
        errorMsg = "Project array not initialized"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        Exit Sub
    End If
    
    ' Find and update project
    foundProject = False
    For i = 1 To UBound(projectName)
        If projectName(i) = mPName Then
            foundProject = True
            
            ' Download fresh project info
            Call DownloadProjectInfo(projectURL(i), UserName(i), Password(i))
            Call RefreshList
            
            WriteLog 1, CurrentMod, PROC_NAME, "Project updated successfully"
            Exit For
        End If
    Next i
    
    If Not foundProject Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project not found: " & mPName
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    ' Continue - non-critical
End Sub

'=======================================================
' Event: UserForm_Initialize
' Purpose: Initialize form on load
'
' Description:
'   Centers form, loads project list, and repaints.
'
' Error Handling:
'   - Logs initialization
'   - Handles errors gracefully
'   - Ensures form is functional
'=======================================================
Private Sub UserForm_Initialize()
    Const PROC_NAME As String = "UserForm_Initialize"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Initializing projects list form"
    
    ' Center form
    Call CenterUserform(Me)
    
    ' Load projects list
    Call RefreshList
    
    ' Repaint form
    Me.Repaint
    
    WriteLog 1, CurrentMod, PROC_NAME, "Form initialized successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Initialization error: " & Err.Description
    MsgBox "Form initialization error: " & Err.Description & vbCrLf & vbCrLf & _
           "Some features may not work correctly.", vbExclamation, "Warning"
End Sub

'=======================================================
' Sub: RefreshList
' Purpose: Refresh projects list from registry
'
' Description:
'   Loads project configuration from registry and
'   populates list control. Updates button states.
'
' Error Handling:
'   - Handles missing projects gracefully
'   - Validates array bounds
'   - Logs errors
'   - Ensures button states are valid
'=======================================================
Private Sub RefreshList()
    Const PROC_NAME As String = "RefreshList"
    
    Dim i As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Refreshing projects list"
    
    ' Clear existing list
    lstProjects.Clear
    
    ' Load projects from registry
    Call LoadProjects
    
    ' Validate arrays
    If Not IsArray(projectName) Then
        WriteLog 2, CurrentMod, PROC_NAME, "No projects configured"
        Call lstProjects_Change
        Exit Sub
    End If
    
    ' Populate list
    For i = 1 To UBound(projectName)
        If Len(projectName(i)) > 0 Then
            lstProjects.AddItem projectName(i)
        End If
    Next i
    
    ' Update button states
    Call lstProjects_Change
    
    WriteLog 1, CurrentMod, PROC_NAME, "List refreshed with " & lstProjects.ListCount & " projects"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    
    ' Ensure button states are valid
    On Error Resume Next
    Call lstProjects_Change
    On Error GoTo 0
End Sub

