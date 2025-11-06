Attribute VB_Name = "AF_Tasks_mod"
Option Explicit
Option Private Module

'=======================================================
' Module: AF_Tasks_mod
' Purpose: Task management operations
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Handles task creation, upload, and display operations
'   for the Docent IMS system. Provides interface for
'   managing project tasks including creation, viewing,
'   and uploading planned/proposed tasks.
'
' Dependencies:
'   - AC_API_Mod (CreateAPITask, GetAPIFolder, GetStateID)
'   - AC_Properties (GetProperty, SetProperty)
'   - AB_GlobalConstants (DateFormat, Success, Critical)
'   - AB_GlobalVars (ProjectNameStr)
'   - frmCreateTask, frmListTasks, frmMsgBox
'
' Public Interface:
'   - AddTask() - Display task creation form
'   - UploadTasks(Planned) - Upload planned/proposed tasks to server
'   - Tasks() - Display task list form
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added detailed logging
'       * Added function documentation
'       * Improved user feedback
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "AF_Tasks_mod"

' Task field indices
Private Const FIELD_TITLE As Long = 0
Private Const FIELD_DETAILS As Long = 1
Private Const FIELD_PRIORITY As Long = 2
Private Const FIELD_ASSIGNED_TO As Long = 3
Private Const FIELD_DUE_DATE As Long = 4
Private Const FIELD_STATE As Long = 5
Private Const REQUIRED_FIELD_COUNT As Long = 6

'=======================================================
' Sub: AddTask
' Purpose: Display task creation form
'
' Description:
'   Shows the task creation form to the user.
'   No parameters required.
'
' Error Handling:
'   - Logs any errors that occur
'   - Displays user-friendly error message
'=======================================================
Sub AddTask()
    Const PROC_NAME As String = "AddTask"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Opening task creation form"
    frmCreateTask.Show
    
    WriteLog 1, CurrentMod, PROC_NAME, "Form closed"
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Failed to open task form: " & errorMsg, vbCritical, "Error"
End Sub

'=======================================================
' Sub: UploadTasks
' Purpose: Upload meeting tasks to the server
'
' Parameters:
'   Planned - True for planned tasks, False for proposed tasks (default: True)
'
' Description:
'   Parses task string from document properties and uploads
'   individual tasks to the API. Tasks are semicolon-delimited,
'   with comma-separated fields within each task.
'   Displays success/failure message to user with detailed results.
'
' Error Handling:
'   - Validates meeting type exists
'   - Validates task string format
'   - Handles API failures gracefully
'   - Provides detailed error logging
'   - Continues processing valid tasks if some fail
'=======================================================
Sub UploadTasks(Optional ByVal Planned As Boolean = True)
    Const PROC_NAME As String = "UploadTasks"
    
    Dim taskString As String
    Dim taskArray() As String
    Dim taskFields() As String
    Dim i As Long
    Dim successCount As Long
    Dim failCount As Long
    Dim meetingType As String
    Dim response As WebResponse
    Dim errorMsg As String
    Dim taskType As String
    
    On Error GoTo ErrorHandler
    
    taskType = IIf(Planned, "Planned", "Proposed")
    WriteLog 1, CurrentMod, PROC_NAME, "Starting " & taskType & " task upload"
    
    ' Validate meeting type
    meetingType = GetProperty(pMeetingType)
    If Len(Trim$(meetingType)) = 0 Then
        errorMsg = "Meeting type is not set"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox "Cannot upload tasks: " & errorMsg, vbExclamation, "Missing Information"
        Exit Sub
    End If
    
    ' Get task string
    taskString = GetProperty(IIf(Planned, pPlannedTasks, pProposedTasks))
    If Len(Trim$(taskString)) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "No " & taskType & " tasks to upload"
        MsgBox "No " & taskType & " tasks found.", vbInformation, "No Tasks"
        Exit Sub
    End If
    
    ' Parse tasks
    On Error Resume Next
    taskArray = Split(taskString, ";")
    If Err.Number <> 0 Then
        errorMsg = "Failed to parse task string: " & Err.Description
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox "Failed to parse tasks: " & errorMsg, vbCritical, "Parse Error"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Validate array
    If UBound(taskArray) < LBound(taskArray) Then
        WriteLog 2, CurrentMod, PROC_NAME, "Task array is empty"
        MsgBox "No valid tasks found.", vbInformation, "No Tasks"
        Exit Sub
    End If
    
    ' Upload each task
    successCount = 0
    failCount = 0
    
    For i = LBound(taskArray) To UBound(taskArray)
        ' Skip empty entries
        If Len(Trim$(taskArray(i))) = 0 Then
            WriteLog 2, CurrentMod, PROC_NAME, "Skipping empty task entry at index " & i
            GoTo NextTask
        End If
        
        ' Parse task fields
        On Error Resume Next
        taskFields = Split(taskArray(i), ",")
        If Err.Number <> 0 Then
            errorMsg = "Failed to parse task fields at index " & i & ": " & Err.Description
            WriteLog 3, CurrentMod, PROC_NAME, errorMsg
            failCount = failCount + 1
            GoTo NextTask
        End If
        On Error GoTo ErrorHandler
        
        ' Validate field count
        If UBound(taskFields) < REQUIRED_FIELD_COUNT - 1 Then
            errorMsg = "Insufficient fields at index " & i & " (expected " & REQUIRED_FIELD_COUNT & _
                      ", got " & (UBound(taskFields) + 1) & ")"
            WriteLog 3, CurrentMod, PROC_NAME, errorMsg
            failCount = failCount + 1
            GoTo NextTask
        End If
        
        ' Validate required fields
        If Len(Trim$(taskFields(FIELD_TITLE))) = 0 Then
            WriteLog 3, CurrentMod, PROC_NAME, "Task title is empty at index " & i
            failCount = failCount + 1
            GoTo NextTask
        End If
        
        ' Upload task
        On Error Resume Next
        Set response = CreateAPITask( _
            taskFields(FIELD_TITLE), _
            taskFields(FIELD_DETAILS), _
            taskFields(FIELD_PRIORITY), _
            taskFields(FIELD_ASSIGNED_TO), _
            taskFields(FIELD_DUE_DATE), _
            taskFields(FIELD_STATE), _
            meetingType)
        
        If Err.Number <> 0 Then
            errorMsg = "API call exception for task '" & taskFields(FIELD_TITLE) & "': " & Err.Description
            WriteLog 3, CurrentMod, PROC_NAME, errorMsg
            failCount = failCount + 1
            GoTo NextTask
        End If
        On Error GoTo ErrorHandler
        
        ' Check response
        If IsGoodResponse(response) Then
            WriteLog 1, CurrentMod, PROC_NAME, "Task created successfully: " & taskFields(FIELD_TITLE)
            successCount = successCount + 1
        Else
            WriteLog 3, CurrentMod, PROC_NAME, "API call failed for task: " & taskFields(FIELD_TITLE)
            failCount = failCount + 1
        End If
        
NextTask:
    Next i
    
    ' Display results
    Call DisplayUploadResults(successCount, failCount, taskType, response)
    
    ' Clear tasks if all successful
    If failCount = 0 And successCount > 0 Then
        SetProperty taskType & "Tasks", ""
        WriteLog 1, CurrentMod, PROC_NAME, "Cleared " & taskType & " tasks property"
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Task upload completed. Success: " & successCount & ", Failed: " & failCount
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Task upload failed: " & errorMsg, vbCritical, "Error"
End Sub

'=======================================================
' Sub: DisplayUploadResults
' Purpose: Display upload results to user
'
' Parameters:
'   successCount - Number of successful uploads
'   failCount - Number of failed uploads
'   taskType - "Planned" or "Proposed"
'   lastResponse - Last API response for link (can be Nothing)
'
' Description:
'   Displays appropriate message based on success/failure counts.
'   Includes link to last created task if available.
'=======================================================
Private Sub DisplayUploadResults(ByVal successCount As Long, _
                                 ByVal failCount As Long, _
                                 ByVal taskType As String, _
                                 ByVal lastResponse As WebResponse)
    Const PROC_NAME As String = "DisplayUploadResults"
    
    Dim msg As String
    Dim icon As Long
    
    On Error GoTo ErrorHandler
    
    If failCount = 0 And successCount > 0 Then
        ' Complete success
        msg = successCount & " " & taskType & " task" & _
              IIf(successCount = 1, " was", "s were") & _
              " created on " & ProjectNameStr & " site."
        icon = Success
        
        If Not lastResponse Is Nothing Then
            On Error Resume Next
            frmMsgBox.Display Array(msg, " ", , "View Online"), _
                             , icon, "DocentIMS", , , _
                             Array(, , , lastResponse.Data("@id"))
            If Err.Number <> 0 Then
                ' Fallback without link
                frmMsgBox.Display msg, , icon, "DocentIMS"
            End If
            On Error GoTo ErrorHandler
        Else
            frmMsgBox.Display msg, , icon, "DocentIMS"
        End If
        
    ElseIf successCount = 0 Then
        ' Complete failure
        msg = "All " & taskType & " tasks failed to upload. Check the log for details."
        frmMsgBox.Display msg, , Critical, "Upload Failed"
        
    Else
        ' Partial success
        msg = successCount & " task" & IIf(successCount = 1, "", "s") & " uploaded successfully." & vbCrLf & _
              failCount & " task" & IIf(failCount = 1, "", "s") & " failed." & vbCrLf & vbCrLf & _
              "Check the log for details on failed tasks."
        frmMsgBox.Display msg, , Exclamation, "Partial Success"
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error displaying results: " & Err.Description
    ' Continue - non-critical
End Sub

'=======================================================
' Sub: Tasks
' Purpose: Display task list form
'
' Description:
'   Retrieves published action items from the API and displays
'   them in the task list form. Filters for published state only.
'
' Error Handling:
'   - Validates API response
'   - Handles missing or null fields
'   - Logs all errors
'   - Shows form even if API call fails (empty list)
'=======================================================
Sub Tasks()
    Const PROC_NAME As String = "Tasks"
    
    Dim taskColl As Collection
    Dim resultDict As Dictionary
    Dim i As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Retrieving task list"
    
    Set resultDict = New Dictionary
    
    ' Get tasks from API
    On Error Resume Next
    Set taskColl = GetAPIFolder("action-items", "action_items", _
                                Array("is_this_item_closed", _
                                      "priority", _
                                      "revised_due_date", _
                                      "duedate", _
                                      "assigned_to", _
                                      "id"), _
                                Array("review_state"), _
                                Array(GetStateID("Published", "action_items")))
    
    If Err.Number <> 0 Then
        errorMsg = "API call failed: " & Err.Description
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        ' Continue with empty collection
        Set taskColl = New Collection
    End If
    On Error GoTo ErrorHandler
    
    ' Process tasks
    If IsGoodResponse(taskColl) Then
        WriteLog 1, CurrentMod, PROC_NAME, "Processing " & taskColl.Count & " tasks"
        
        For i = 1 To taskColl.Count
            On Error Resume Next
            
            resultDict.Add , Array( _
                GetID(CStr(taskColl(i)("id"))), _
                NoNull(taskColl(i)("title")), _
                NoNull(taskColl(i)("assigned_to")), _
                NoNull(taskColl(i)("priority")), _
                GetDueDate(taskColl(i)("duedate"), taskColl(i)("revised_due_date")), _
                NoNull(taskColl(i)("is_this_item_closed")), _
                taskColl(i)("@id"))
            
            If Err.Number <> 0 Then
                WriteLog 3, CurrentMod, PROC_NAME, "Error processing task " & i & ": " & Err.Description
                Err.Clear
            End If
            
            On Error GoTo ErrorHandler
        Next i
    Else
        WriteLog 2, CurrentMod, PROC_NAME, "No tasks retrieved or bad response"
    End If
    
    ' Display form
    Set frmListTasks.ItemsDict = resultDict
    frmListTasks.Show
    
    WriteLog 1, CurrentMod, PROC_NAME, "Task form displayed with " & resultDict.Count & " tasks"
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "Error loading tasks: " & errorMsg & vbCrLf & vbCrLf & _
           "The task list may be incomplete.", vbExclamation, "Error"
    
    ' Show form anyway with whatever data we have
    On Error Resume Next
    If resultDict Is Nothing Then Set resultDict = New Dictionary
    Set frmListTasks.ItemsDict = resultDict
    frmListTasks.Show
End Sub

'=======================================================
' HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: NoNull
' Purpose: Convert null/empty variants to empty string
'
' Parameters:
'   s - Variant to convert
'
' Returns:
'   String value, or empty string if null/error
'
' Description:
'   Safe conversion function that handles null, empty,
'   and error conditions gracefully.
'=======================================================
Private Function NoNull(ByVal s As Variant) As String
    On Error Resume Next
    NoNull = CStr(s)
    If Err.Number <> 0 Then NoNull = ""
End Function

'=======================================================
' Function: GetID
' Purpose: Extract numeric ID from end of URL string
'
' Parameters:
'   URL - URL string containing numeric ID
'
' Returns:
'   Numeric ID as string, or "0" if not found
'
' Description:
'   Scans URL from right to left to extract trailing
'   numeric characters. Useful for extracting task IDs
'   from API URLs.
'
' Example:
'   GetID("http://site/tasks/action-item-123") returns "123"
'=======================================================
Private Function GetID(ByVal URL As String) As String
    Const PROC_NAME As String = "GetID"
    
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(Trim$(URL)) = 0 Then
        GetID = "0"
        Exit Function
    End If
    
    ' Scan from right to find non-numeric character
    For i = Len(URL) To 1 Step -1
        If Not IsNumeric(Mid$(URL, i, 1)) Then
            i = i + 1
            Exit For
        End If
    Next i
    
    ' Extract ID
    If i = 0 Or i > Len(URL) Then
        GetID = "0"
    Else
        GetID = Mid$(URL, i)
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error extracting ID: " & Err.Description
    GetID = "0"
End Function

'=======================================================
' Function: GetDueDate
' Purpose: Get formatted due date, preferring revised date
'
' Parameters:
'   DueDate - Original due date (can be null)
'   RevisedDueDate - Revised due date (can be null)
'
' Returns:
'   Formatted date string using DateFormat constant
'
' Description:
'   Returns revised due date if present, otherwise
'   original due date. Handles null values gracefully.
'=======================================================
Private Function GetDueDate(ByVal DueDate As Variant, _
                           ByVal RevisedDueDate As Variant) As String
    Const PROC_NAME As String = "GetDueDate"
    
    Dim dateToUse As String
    
    On Error GoTo ErrorHandler
    
    DueDate = NoNull(DueDate)
    RevisedDueDate = NoNull(RevisedDueDate)
    
    ' Use revised date if available, otherwise original
    If Len(RevisedDueDate) > 0 Then
        dateToUse = RevisedDueDate
    Else
        dateToUse = DueDate
    End If
    
    ' Format date
    If Len(dateToUse) > 0 Then
        On Error Resume Next
        GetDueDate = Format$(dateToUse, DateFormat)
        If Err.Number <> 0 Then
            GetDueDate = dateToUse  ' Return unformatted if format fails
        End If
    Else
        GetDueDate = ""
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error formatting date: " & Err.Description
    GetDueDate = ""
End Function
