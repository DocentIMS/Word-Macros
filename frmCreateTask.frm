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
'=======================================================
' UserForm: frmCreateTask
' Purpose: Task creation and editing interface
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Provides comprehensive task creation interface with
'   support for both upload mode (direct API) and
'   metadata mode (save to document properties).
'
' Dependencies:
'   - AC_API_Mod (CreateAPITask, UpdateAPIFileWorkflow, GetTransitionIdByStates)
'   - AC_API_Mod (GetStatesOfDoc, GetTaskPriorities, GetMembersOf, GetMemberID)
'   - AC_Properties (GetProperty, SetProperty)
'   - AB_GlobalConstants (ErrorColor, DateFormat, Success, Critical)
'   - AB_GlobalVars (ProjectNameStr, ProjectColorStr)
'   - CtrlEvents (field validation class)
'   - frmMsgBox
'
' Public Interface:
'   - Display(UploadMode) - Show form in upload or metadata mode
'
' Private Methods:
'   - btn_OK_Click - Handle OK button
'   - SaveToMetadata - Save task to document properties
'   - GetItemsOf - Get selected items from listbox
'   - AddToTaskTable - Add task to document table
'   - UserForm_Initialize - Initialize form controls
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added detailed logging
'       * Added function documentation
'       * Improved resource cleanup
'       * Removed commented dead code
'   v1.0 - Original version
'=======================================================

Option Explicit

Private Const CurrentMod As String = "frmCreateTask"

' Form-level collections and variables
Private PArr As Variant
Private PColl As New Collection
Private BColl As New Collection
Private mUploadMode As Boolean
Private Evs As New CtrlEvents

'=======================================================
' Event: btn_OK_Click
' Purpose: Handle OK button click - create or save task
'
' Description:
'   Creates task via API (upload mode) or saves to document
'   properties/table (metadata mode). Validates inputs and
'   provides user feedback.
'
' Error Handling:
'   - Validates all required fields
'   - Handles API failures gracefully
'   - Logs all operations
'   - Displays user-friendly error messages
'=======================================================
Private Sub btn_OK_Click()
    Const PROC_NAME As String = "btn_OK_Click"
    
    Dim response As WebResponse
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "OK button clicked (UploadMode=" & mUploadMode & ")"
    
    ' Validate details field (server requires non-empty)
    If Len(Trim$(tbDetails.value)) = 0 Then
        tbDetails.value = " "
        WriteLog 2, CurrentMod, PROC_NAME, "Empty details field - using space to bypass server validation"
    End If
    
    If mUploadMode Then
        ' Upload mode - create task via API
        Call CreateTaskViaAPI(response)
        
        If response Is Nothing Then
            WriteLog 3, CurrentMod, PROC_NAME, "API response is Nothing"
            Exit Sub
        End If
        
        ' Handle workflow transition if needed
        If cbState.value <> "Private" Then
            Call TransitionTaskState(response.Data("@id"))
        End If
        
        ' Display success message
        frmMsgBox.Display Array("A new Task was created on " & ProjectNameStr & " site.", " ", , "View Online"), _
                         , Success, "DocentIMS", , , Array(, , , response.Data("@id"))
        
        WriteLog 1, CurrentMod, PROC_NAME, "Task created successfully via API"
        Unload Me
        
    Else
        ' Metadata mode - save to document
        Call SaveToMetadata
        Call AddToTaskTable
        
        WriteLog 1, CurrentMod, PROC_NAME, "Task saved to document metadata"
        Unload Me
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Failed to create task: " & errorMsg, vbCritical, "Error"
End Sub

'=======================================================
' Sub: CreateTaskViaAPI
' Purpose: Create task via API call
'
' Parameters:
'   response - Output WebResponse object
'
' Error Handling:
'   - Validates priority value exists
'   - Validates member ID exists
'   - Handles API call failures
'   - Displays user-friendly error messages
'=======================================================
Private Sub CreateTaskViaAPI(ByRef response As WebResponse)
    Const PROC_NAME As String = "CreateTaskViaAPI"
    
    Dim priorityValue As String
    Dim memberID As String
    Dim errorMsg As String
    Dim otherMembers As Collection
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Creating task via API"
    
    ' Validate and get priority value
    On Error Resume Next
    priorityValue = PColl(cbPriority.value)
    If Err.Number <> 0 Or Len(priorityValue) = 0 Then
        errorMsg = "Invalid priority selection"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbExclamation, "Validation Error"
        Set response = Nothing
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Get member ID
    memberID = GetMemberID(cbWho.value)
    If Len(memberID) = 0 Then
        errorMsg = "Could not find member ID for: " & cbWho.value
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox "Invalid team member selection", vbExclamation, "Validation Error"
        Set response = Nothing
        Exit Sub
    End If
    
    ' Get other members list
    Set otherMembers = GetItemsOf(liOthers)
    
    ' Create task via API
    Set response = CreateAPITask(tbTitle.value, tbDetails.value, _
                                 tbdDueDate.value, priorityValue, _
                                 memberID, tbNotes.value, tbPrivateNotes.value, , , otherMembers)
    
    If Not IsGoodResponse(response) Then
        errorMsg = "Task creation failed"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox "Task could not be created on the server.", vbCritical, "DocentIMS"
        Set response = Nothing
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Task created successfully"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Failed to create task: " & errorMsg, vbCritical, "Error"
    Set response = Nothing
End Sub

'=======================================================
' Sub: TransitionTaskState
' Purpose: Transition task to selected state
'
' Parameters:
'   taskURL - Task API URL
'
' Error Handling:
'   - Logs transition attempts
'   - Handles transition failures
'   - Continues on error (non-critical)
'=======================================================
Private Sub TransitionTaskState(ByVal taskURL As String)
    Const PROC_NAME As String = "TransitionTaskState"
    
    Dim transitionID As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Transitioning task from Private to " & cbState.value
    
    transitionID = GetTransitionIdByStates("Private", cbState.value, "action_items")
    
    If Len(transitionID) > 0 Then
        UpdateAPIFileWorkflow taskURL, transitionID
        RefreshTasksGroup
        WriteLog 1, CurrentMod, PROC_NAME, "Task transitioned successfully"
    Else
        WriteLog 2, CurrentMod, PROC_NAME, "No transition ID found for state: " & cbState.value
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Transition error: " & Err.Description
    ' Continue - non-critical
End Sub

'=======================================================
' Sub: SaveToMetadata
' Purpose: Save task to document properties
'
' Description:
'   Appends task information to document's proposed tasks
'   property as semicolon/comma delimited string.
'
' Error Handling:
'   - Handles missing property gracefully
'   - Logs errors
'   - Continues execution
'=======================================================
Private Sub SaveToMetadata()
    Const PROC_NAME As String = "SaveToMetadata"
    
    Dim existingTasks As String
    Dim newTaskString As String
    Dim memberID As String
    Dim priorityValue As String
    Dim otherMembersStr As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Saving task to document metadata"
    
    ' Get existing tasks
    On Error Resume Next
    existingTasks = GetProperty(pProposedTasks)
    If Err.Number <> 0 Then existingTasks = ""
    On Error GoTo ErrorHandler
    
    ' Get member ID and priority
    memberID = GetMemberID(cbWho.value)
    If Len(memberID) = 0 Then memberID = cbWho.value
    
    On Error Resume Next
    priorityValue = PColl(cbPriority.value)
    If Err.Number <> 0 Then priorityValue = cbPriority.value
    On Error GoTo ErrorHandler
    
    ' Build other members string
    otherMembersStr = Join(CollToArr(GetItemsOf(liOthers)), ";")
    
    ' Build new task string
    newTaskString = existingTasks & ";," & _
                   tbTitle.value & "," & _
                   memberID & "," & _
                   priorityValue & "," & _
                   tbdDueDate.value & "," & _
                   tbDetails.value & "," & _
                   tbNotes.value & "," & _
                   tbPrivateNotes.value
    
    ' Save to property
    SetProperty pProposedTasks, newTaskString
    
    WriteLog 1, CurrentMod, PROC_NAME, "Task saved to metadata successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error saving to metadata: " & Err.Description
    ' Continue - let calling function handle
End Sub

'=======================================================
' Function: GetItemsOf
' Purpose: Get collection of selected items from listbox
'
' Parameters:
'   Ctrl - ListBox control
'
' Returns:
'   Collection - Selected member IDs
'
' Error Handling:
'   - Returns empty collection on error
'   - Handles invalid member lookups
'   - Logs errors
'=======================================================
Private Function GetItemsOf(ByVal Ctrl As MSForms.ListBox) As Collection
    Const PROC_NAME As String = "GetItemsOf"
    
    Dim i As Long
    Dim userID As String
    Dim result As New Collection
    
    On Error GoTo ErrorHandler
    
    For i = 0 To Ctrl.ListCount - 1
        If Ctrl.Selected(i) Then
            On Error Resume Next
            userID = GetMemberID(Ctrl.List(i))
            
            If Err.Number <> 0 Or Len(userID) = 0 Then
                userID = Ctrl.List(i)
            End If
            
            result.Add userID
            
            On Error GoTo ErrorHandler
        End If
    Next i
    
    Set GetItemsOf = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set GetItemsOf = New Collection
End Function

'=======================================================
' Sub: AddToTaskTable
' Purpose: Add task to document's Proposed Tasks table
'
' Description:
'   Finds the "Proposed Tasks" table in active document
'   and adds a new row with task information.
'
' Error Handling:
'   - Validates active document exists
'   - Validates table exists
'   - Handles table access errors
'   - Proper document protection handling
'=======================================================
Private Sub AddToTaskTable()
    Const PROC_NAME As String = "AddToTaskTable"
    
    Dim i As Long
    Dim tableIndex As Long
    Dim newRowIndex As Long
    Dim foundTable As Boolean
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding task to document table"
    
    ' Validate active document
    If ActiveDocument Is Nothing Then
        errorMsg = "No active document"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Unprotect document
    Call Unprotect(ActiveDocument)
    
    ' Find "Proposed Tasks" table
    foundTable = False
    For i = 1 To ActiveDocument.Tables.Count
        If ActiveDocument.Tables(i).Title = "Proposed Tasks" Then
            tableIndex = i
            foundTable = True
            Exit For
        End If
    Next i
    
    If Not foundTable Then
        errorMsg = "Proposed Tasks table not found"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbExclamation, "Error"
        Call Protect(ActiveDocument)
        Exit Sub
    End If
    
    ' Add row and populate
    With ActiveDocument.Tables(tableIndex)
        On Error Resume Next
        .Rows.Add
        
        If Err.Number <> 0 Then
            WriteLog 3, CurrentMod, PROC_NAME, "Failed to add table row: " & Err.Description
            Call Protect(ActiveDocument)
            MsgBox "Failed to add task to table", vbExclamation, "Error"
            Exit Sub
        End If
        
        On Error GoTo ErrorHandler
        
        newRowIndex = .Rows.Count
        
        .Rows(newRowIndex).Cells(1).Range.text = tbTitle.value
        .Rows(newRowIndex).Cells(2).Range.text = cbWho.value
        .Rows(newRowIndex).Cells(3).Range.text = cbPriority.value
        .Rows(newRowIndex).Cells(4).Range.text = Format$(tbdDueDate.value, DateFormat)
    End With
    
    ' Protect document
    Call Protect(ActiveDocument)
    
    WriteLog 1, CurrentMod, PROC_NAME, "Task added to table successfully"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    
    ' Ensure document is protected on error
    On Error Resume Next
    Call Protect(ActiveDocument)
    On Error GoTo 0
    
    MsgBox "Failed to add task to table: " & errorMsg, vbExclamation, "Error"
End Sub

'=======================================================
' Event: btnCancel_Click
' Purpose: Handle Cancel button click
'=======================================================
Private Sub btnCancel_Click()
    Const PROC_NAME As String = "btnCancel_Click"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Cancel button clicked"
    Unload Me
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Unload Me
End Sub

'=======================================================
' Sub: Display
' Purpose: Show form in specified mode
'
' Parameters:
'   UploadMode - True for API upload, False for metadata mode
'
' Error Handling:
'   - Logs mode selection
'   - Handles form display errors
'=======================================================
Sub Display(Optional ByVal UploadMode As Boolean = False)
    Const PROC_NAME As String = "Display"
    
    On Error GoTo ErrorHandler
    
    mUploadMode = UploadMode
    WriteLog 1, CurrentMod, PROC_NAME, "Displaying form (UploadMode=" & UploadMode & ")"
    
    Me.Show
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display form: " & Err.Description, vbCritical, "Error"
End Sub

'=======================================================
' Event: UserForm_Initialize
' Purpose: Initialize form controls and load data
'
' Description:
'   Sets up event handlers, loads dropdown data
'   (states, priorities, team members), and applies
'   project theme.
'
' Error Handling:
'   - Continues on non-critical errors
'   - Logs all steps
'   - Provides fallbacks for missing data
'=======================================================
Private Sub UserForm_Initialize()
    Const PROC_NAME As String = "UserForm_Initialize"
    
    Dim i As Long
    Dim boardMembers As Dictionary
    Dim docStates As Dictionary
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Initializing task form"
    
    ' Initialize event handlers
    Call InitializeEventHandlers
    
    ' Load states
    Call LoadStates(docStates)
    
    ' Load priorities
    Call LoadPriorities
    
    ' Load team members
    Call LoadTeamMembers(boardMembers)
    
    ' Apply project theme
    Call ApplyProjectTheme
    
    WriteLog 1, CurrentMod, PROC_NAME, "Form initialized successfully"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Form initialization error: " & errorMsg & vbCrLf & vbCrLf & _
           "Some features may not work correctly.", vbExclamation, "Warning"
End Sub

'=======================================================
' INITIALIZATION HELPER FUNCTIONS
'=======================================================

'=======================================================
' Sub: InitializeEventHandlers
' Purpose: Set up event handler class
'=======================================================
Private Sub InitializeEventHandlers()
    Const PROC_NAME As String = "InitializeEventHandlers"
    
    On Error GoTo ErrorHandler
    
    Set Evs.Parent = Me
    Evs.AddOkButton btn_OK
    Evs.MakeRequired "Title,Details,Priority,Who,DueDate,cbState", , ErrorColor
    
    WriteLog 1, CurrentMod, PROC_NAME, "Event handlers initialized"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
End Sub

'=======================================================
' Sub: LoadStates
' Purpose: Load task states into dropdown
'=======================================================
Private Sub LoadStates(ByRef docStates As Dictionary)
    Const PROC_NAME As String = "LoadStates"
    
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set docStates = GetStatesOfDoc("action_items")
    
    If docStates Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "States dictionary is Nothing"
        Exit Sub
    End If
    
    cbState.Clear
    For i = 1 To docStates.Count
        cbState.AddItem docStates(i)
    Next i
    
    WriteLog 1, CurrentMod, PROC_NAME, "Loaded " & docStates.Count & " states"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
End Sub

'=======================================================
' Sub: LoadPriorities
' Purpose: Load task priorities into dropdown
'=======================================================
Private Sub LoadPriorities()
    Const PROC_NAME As String = "LoadPriorities"
    
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    PArr = GetTaskPriorities()
    
    If Not IsGoodResponse(PArr(1, 1)) Then
        WriteLog 2, CurrentMod, PROC_NAME, "Failed to load priorities from API"
        Call LoadDefaultPriorities
        Exit Sub
    End If
    
    cbPriority.Clear
    For i = LBound(PArr, 2) To UBound(PArr, 2)
        cbPriority.AddItem PArr(1, i)
        PColl.Add PArr(2, i), PArr(1, i)
    Next i
    
    PColl.Add "", ""
    
    WriteLog 1, CurrentMod, PROC_NAME, "Loaded " & (UBound(PArr, 2) - LBound(PArr, 2) + 1) & " priorities"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Call LoadDefaultPriorities
End Sub

'=======================================================
' Sub: LoadDefaultPriorities
' Purpose: Load fallback priorities if API fails
'=======================================================
Private Sub LoadDefaultPriorities()
    Const PROC_NAME As String = "LoadDefaultPriorities"
    
    On Error Resume Next
    
    WriteLog 2, CurrentMod, PROC_NAME, "Loading default priorities"
    
    cbPriority.Clear
    cbPriority.AddItem "High"
    cbPriority.AddItem "Medium"
    cbPriority.AddItem "Low"
    
    PColl.Add "1", "High"
    PColl.Add "2", "Medium"
    PColl.Add "3", "Low"
    PColl.Add "", ""
End Sub

'=======================================================
' Sub: LoadTeamMembers
' Purpose: Load team members into dropdowns and lists
'=======================================================
Private Sub LoadTeamMembers(ByRef boardMembers As Dictionary)
    Const PROC_NAME As String = "LoadTeamMembers"
    
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set boardMembers = GetMembersOf()
    
    If boardMembers Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "Board members dictionary is Nothing"
        Exit Sub
    End If
    
    cbWho.Clear
    liOthers.Clear
    
    For i = 1 To boardMembers.Count
        cbWho.AddItem boardMembers(i)
        liOthers.AddItem boardMembers(i)
    Next i
    
    WriteLog 1, CurrentMod, PROC_NAME, "Loaded " & boardMembers.Count & " team members"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
End Sub

'=======================================================
' Sub: ApplyProjectTheme
' Purpose: Apply project colors and branding to form
'=======================================================
Private Sub ApplyProjectTheme()
    Const PROC_NAME As String = "ApplyProjectTheme"
    
    On Error GoTo ErrorHandler
    
    lbPrjHeader.Caption = ProjectNameStr
    
    If ProjectColorStr <> 0 Then
        lbPrjHeader.BackColor = ProjectColorStr
        
        On Error Resume Next
        lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
        On Error GoTo ErrorHandler
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Project theme applied"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    ' Continue - use default colors
End Sub
