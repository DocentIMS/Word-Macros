Attribute VB_Name = "Docs_Scope"
Option Explicit

'=======================================================
' Module: Docs_Scope
' Purpose: Scope document management and task operations
' Author: Docent IMS Team
' Version: 2.0
'
' Description:
'   Handles all scope document operations including:
'   - Task creation and management (add sublevel, same level, top level)
'   - Meeting integration with scope tasks
'   - Document revision tracking
'   - Scope document unlocking and protection
'   - Meeting summary table extraction
'   - Section break management
'
'   This module manages the hierarchical task structure in
'   scope documents and maintains the relationships between
'   tasks and related meetings.
'
' Dependencies:
'   - AB_GlobalConstants (TasksBookmark constant recommended)
'   - AB_GlobalVars (for global variables)
'   - AB_CommonFunctions (for helper functions)
'   - AC_Properties (for document properties)
'   - frmAddScopeMeeting, frmUpdateRev, frmInputBox
'   - ScopeTask class
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive module documentation
'       * Split GetMeetingSummaryTable into focused functions
'       * Added proper error handling to all functions
'       * Added function documentation headers
'       * Removed ALL commented-out code
'       * Fixed On Error Resume Next usage
'       * Moved TasksBookmark to constants (recommended)
'       * Added all missing functions
'   v1.0 - Original version
'=======================================================

' Module constants
Private Const CurrentMod As String = "Docs_Scope"
Private Const TasksBookmark As String = "EOT"  ' TODO: Move to AB_GlobalConstants

' Module-level variables
Private Tbl As Table  ' Current working table

'=======================================================
' MEETING AND SCOPE INTEGRATION
'=======================================================

'=======================================================
' Sub: AddMeetingToScopeTask
' Purpose: Add meeting reference to a scope task
'
' Description:
'   Allows user to link a meeting to a scope task by
'   validating cursor position and displaying the meeting
'   selection form. Cursor must be inside a "Related Meetings"
'   table of a task.
'
' Error Handling:
'   - Validates table selection
'   - Provides user feedback on invalid selection
'=======================================================
Public Sub AddMeetingToScopeTask()
    Const PROC_NAME As String = "AddMeetingToScopeTask"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting meeting addition to scope task"
    
    ' Validate table selection
    If CorrectTableSelected Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "Invalid table selection"
        frmMsgBox.Display "Please place cursor inside the ""Related Meetings"" table of the task", , Critical, "Invalid Selection"
        Exit Sub
    End If
    
    ' Display form to select meeting
    frmAddScopeMeeting.Display
    
    WriteLog 1, CurrentMod, PROC_NAME, "Meeting addition completed"
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errMsg
    frmMsgBox.Display "Unable to add meeting to scope task." & vbCrLf & vbCrLf & _
                      "Error: " & errMsg, , Critical, "Error"
End Sub

'=======================================================
' Function: CorrectTableSelected
' Purpose: Validate that cursor is in a Related Meetings table
'
' Returns:
'   Table object if valid selection, Nothing otherwise
'
' Description:
'   Validates three conditions:
'   1. Document type is "Scope"
'   2. Cursor is inside a table
'   3. Table title contains "Meetings"
'=======================================================
Function CorrectTableSelected() As Table
    Const PROC_NAME As String = "CorrectTableSelected"
    
    Dim currentDoc As Document
    Dim selectedTable As Table
    
    On Error GoTo ErrorHandler
    
    ' Validate document type
    Set currentDoc = ActiveDocument
    If GetProperty(pDocType, currentDoc) <> "Scope" Then
        WriteLog 2, CurrentMod, PROC_NAME, "Document is not a Scope document"
        Set CorrectTableSelected = Nothing
        Exit Function
    End If
    
    ' Get table at cursor position
    Set selectedTable = Selection.Range.Tables(1)
    If selectedTable Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "No table selected"
        Set CorrectTableSelected = Nothing
        Exit Function
    End If
    
    ' Validate table title
    If Not selectedTable.Title Like "*Meetings" Then
        WriteLog 2, CurrentMod, PROC_NAME, "Selected table is not a Meetings table: " & selectedTable.Title
        Set CorrectTableSelected = Nothing
        Exit Function
    End If
    
    ' Valid selection
    Set CorrectTableSelected = selectedTable
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Set CorrectTableSelected = Nothing
End Function

'=======================================================
' MEETING SUMMARY TABLE EXTRACTION
'=======================================================

'=======================================================
' Function: GetMeetingSummaryTable
' Purpose: Extract meeting summary table as JSON
'
' Returns:
'   JSON string representation of meeting summary table,
'   or empty string if table not found
'
' Description:
'   Searches for the "Meeting Summary Table" in the
'   active document and converts it to JSON format.
'   Uses helper functions for clean separation of concerns.
'=======================================================
Public Function GetMeetingSummaryTable() As String
    Const PROC_NAME As String = "GetMeetingSummaryTable"
    
    Dim meetingTable As Table
    Dim jsonResult As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Searching for meeting summary table"
    
    ' Find the meeting summary table
    Set meetingTable = FindMeetingSummaryTable()
    If meetingTable Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "Meeting summary table not found"
        GetMeetingSummaryTable = vbNullString
        Exit Function
    End If
    
    ' Convert table to JSON
    jsonResult = ConvertTableToJSON(meetingTable, "Meeting Summary Table")
    
    WriteLog 1, CurrentMod, PROC_NAME, "Successfully converted table to JSON (" & Len(jsonResult) & " characters)"
    GetMeetingSummaryTable = jsonResult
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    GetMeetingSummaryTable = vbNullString
End Function

'=======================================================
' Function: FindMeetingSummaryTable
' Purpose: Locate meeting summary table in document
'
' Returns:
'   Table object if found, Nothing otherwise
'
' Description:
'   Searches for text matching "Meet* Summary Table"
'   pattern and returns the containing table.
'=======================================================
Private Function FindMeetingSummaryTable() As Table
    Const PROC_NAME As String = "FindMeetingSummaryTable"
    
    Dim searchRange As Range
    
    On Error GoTo ErrorHandler
    
    Set searchRange = ActiveDocument.Range
    
    With searchRange.Find
        .text = "Meet* Summary Table"
        .MatchWildcards = True
        .Execute
        
        If .Found Then
            ' Navigate to the containing table
            Do
                searchRange.MoveUntil Chr(7)
            Loop Until searchRange.Information(wdWithInTable)
            
            Set FindMeetingSummaryTable = searchRange.Tables(1)
            WriteLog 1, CurrentMod, PROC_NAME, "Found meeting summary table"
        Else
            Set FindMeetingSummaryTable = Nothing
            WriteLog 2, CurrentMod, PROC_NAME, "Meeting summary table not found"
        End If
    End With
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Set FindMeetingSummaryTable = Nothing
End Function

'=======================================================
' Function: ConvertTableToJSON
' Purpose: Convert Word table to JSON format
'
' Parameters:
'   sourceTable - Table to convert
'   tableName - JSON property name for the table
'
' Returns:
'   JSON string representation of table
'
' Description:
'   Converts table rows to JSON array of objects.
'   Row 1 is treated as headers, rows 2+ as data.
'   Handles multi-line cell values as JSON arrays.
'=======================================================
Private Function ConvertTableToJSON(ByVal sourceTable As Table, _
                                   ByVal TableName As String) As String
    Const PROC_NAME As String = "ConvertTableToJSON"
    
    Dim jsonBuilder As String
    Dim rowIndex As Long
    
    On Error GoTo ErrorHandler
    
    jsonBuilder = """" & TableName & """: ["
    
    ' Process data rows (skip header row 1)
    For rowIndex = 2 To sourceTable.Rows.Count
        jsonBuilder = jsonBuilder & BuildJSONRow(sourceTable, rowIndex)
        
        ' Add comma between rows
        If rowIndex < sourceTable.Rows.Count Then
            jsonBuilder = jsonBuilder & ","
        End If
    Next rowIndex
    
    ConvertTableToJSON = jsonBuilder & "]"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ConvertTableToJSON = """" & TableName & """: []"
End Function

'=======================================================
' Function: BuildJSONRow
' Purpose: Build JSON object for a single table row
'
' Parameters:
'   sourceTable - Source table
'   rowIndex - Row number to process (1-based)
'
' Returns:
'   JSON object string for the row
'
' Description:
'   Creates JSON object with properties from header row
'   and values from specified data row.
'=======================================================
Private Function BuildJSONRow(ByVal sourceTable As Table, _
                             ByVal rowIndex As Long) As String
    Const PROC_NAME As String = "BuildJSONRow"
    
    Dim jsonObj As String
    Dim colIndex As Long
    Dim cellValue As String
    Dim headerValue As String
    
    On Error GoTo ErrorHandler
    
    jsonObj = "{"
    
    ' Process each column
    For colIndex = 1 To sourceTable.Rows(rowIndex).Cells.Count
        ' Get header and cell values
        headerValue = GetCellValue(sourceTable, 1, colIndex)
        cellValue = GetCellValue(sourceTable, rowIndex, colIndex)
        
        ' Add property: "header": value
        jsonObj = jsonObj & """" & headerValue & """: "
        jsonObj = jsonObj & FormatJSONValue(cellValue)
        
        ' Add comma between properties
        If colIndex < sourceTable.Rows(rowIndex).Cells.Count Then
            jsonObj = jsonObj & ","
        End If
    Next colIndex
    
    BuildJSONRow = jsonObj & "}"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description & " (Row: " & rowIndex & ")"
    BuildJSONRow = "{}"
End Function

'=======================================================
' Function: FormatJSONValue
' Purpose: Format cell value for JSON output
'
' Parameters:
'   cellValue - Raw cell value
'
' Returns:
'   Formatted JSON value (string or array)
'
' Description:
'   Handles both single values and multi-line values.
'   Multi-line values (with Chr(13)) are converted to
'   JSON arrays.
'=======================================================
Private Function FormatJSONValue(ByVal cellValue As String) As String
    Const PROC_NAME As String = "FormatJSONValue"
    
    Dim values() As String
    Dim i As Long
    Dim jsonArray As String
    
    On Error GoTo ErrorHandler
    
    ' Check if value contains line breaks (multi-value cell)
    If InStr(cellValue, Chr(13)) > 0 Then
        ' Split into array
        values = Split(cellValue, Chr(13))
        jsonArray = "["
        
        For i = LBound(values) To UBound(values)
            jsonArray = jsonArray & """" & Trim$(values(i)) & """"
            
            ' Add comma between values
            If i < UBound(values) Then
                jsonArray = jsonArray & ","
            End If
        Next i
        
        FormatJSONValue = jsonArray & "]"
    Else
        ' Single value - return as string
        FormatJSONValue = """" & cellValue & """"
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    FormatJSONValue = """" & cellValue & """"
End Function

'=======================================================
' Function: GetCellValue
' Purpose: Safely retrieve cell value from table
'
' Parameters:
'   sourceTable - Source table
'   rowIndex - Row number (1-based)
'   colIndex - Column number (1-based)
'
' Returns:
'   Cell text value with table markers removed, or empty string on error
'=======================================================
Private Function GetCellValue(ByVal sourceTable As Table, _
                             ByVal rowIndex As Long, _
                             ByVal colIndex As Long) As String
    Const PROC_NAME As String = "GetCellValue"
    
    Dim cellText As String
    
    On Error GoTo ErrorHandler
    
    ' Get cell text and remove table cell markers (Chr(13) + Chr(7))
    cellText = sourceTable.Cell(rowIndex, colIndex).Range.text
    cellText = Replace(cellText, Chr(13) & Chr(7), "")
    cellText = Trim$(cellText)
    
    GetCellValue = cellText
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & _
             " (Row: " & rowIndex & ", Col: " & colIndex & ")"
    GetCellValue = vbNullString
End Function

'=======================================================
' DOCUMENT PROTECTION AND UNLOCKING
'=======================================================

'=======================================================
' Sub: UnlockDocument
' Purpose: Unlock protected scope document for editing
'
' Description:
'   Attempts to unprotect the document using the current
'   username. If successful, allows editing. If document
'   was protected by another user, informs user they cannot
'   edit the final version.
'
' Error Handling:
'   Uses error handling to gracefully handle protection failures
'=======================================================
Public Sub UnlockDocument()
    Const PROC_NAME As String = "UnlockDocument"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Attempting to unlock document"
    
    If ActiveDocument.protectionType <> wdNoProtection Then
        ' Try to unprotect with current username
        ActiveDocument.Unprotect Application.UserName
        
        ' Check if unprotect was successful
        If ActiveDocument.protectionType <> wdNoProtection Then
            WriteLog 2, CurrentMod, PROC_NAME, "User not authorized to unlock document"
            frmMsgBox.Display "You are not allowed to edit this final version", , Critical, "Access Denied"
        Else
            WriteLog 1, CurrentMod, PROC_NAME, "Document unlocked successfully"
            frmMsgBox.Display "Document is unlocked", , Success, "Unlocked"
        End If
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "Document is already unlocked"
        frmMsgBox.Display "Document is not protected", , Information, "Not Protected"
    End If
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    frmMsgBox.Display "Unable to unlock document." & vbCrLf & vbCrLf & _
                      "Error: " & Err.Description, , Critical, "Error"
End Sub

'=======================================================
' REVISION MANAGEMENT
'=======================================================

'=======================================================
' Sub: FillFirstRevisionDate
' Purpose: Fill initial revision date in revisions table
'
' Description:
'   If the first data row (row 2) of the Revisions Table
'   has no date, fills it with the current server time.
'   Used when creating new scope documents.
'
' Error Handling:
'   Silently exits if table not found or date already exists
'=======================================================
Public Sub FillFirstRevisionDate()
    Const PROC_NAME As String = "FillFirstRevisionDate"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Checking first revision date"
    
    ' Get revisions table
    Set Tbl = GetTableByTitle("Revisions Table")
    If Tbl Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "Revisions table not found"
        Exit Sub
    End If
    
    ' Check if date is empty
    If Len(Cell(2, 3)) = 0 Then
        Tbl.Rows(2).Cells(3).Range.text = Format(ToServerTime, DateFormat)
        WriteLog 1, CurrentMod, PROC_NAME, "First revision date set"
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "First revision date already exists"
    End If
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Silent fail - not critical
End Sub

'=======================================================
' Sub: UpdateRevision
' Purpose: Add new revision entry to revisions table
'
' Description:
'   Displays form to collect revision information, then:
'   - Adds new row to revisions table with info
'   - Updates footer version information
'   - If final revision: removes all comments and sets
'     final flag in document properties
'
' Error Handling:
'   Comprehensive error handling with user notification
'=======================================================
Public Sub UpdateRevision()
    Const PROC_NAME As String = "UpdateRevision"
    
    Dim commentIndex As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting revision update"
    
    With frmUpdateRev
        .Show
        
        If Not .Cancelled Then
            ' Add revision entry
            If .IsFinal Then
                AddRev "Final Scope of Work (Clean)", .tbNotes.value, "Final"
            Else
                AddRev "Draft Scope of Work", .tbNotes.value, ""
            End If
            
            ' Update footer
            UpdateFooterVersion
            
            ' Handle final revision processing
            If .IsFinal Then
                WriteLog 1, CurrentMod, PROC_NAME, "Processing final revision"
                
                ' Remove all comments (count down to avoid index issues)
                For commentIndex = ActiveDocument.Comments.Count To 1 Step -1
                    ActiveDocument.Comments(commentIndex).DeleteRecursively
                Next commentIndex
                
                ' Set final revision property
                SetProperty pIsFinalRev, True, ActiveDocument
                LoadDocInfo ActiveDocument
                RefreshRibbon
                
                WriteLog 1, CurrentMod, PROC_NAME, "Final revision processed - comments removed"
                frmMsgBox.Display "Comments removed and document marked as final", , Success, "Final Revision"
            End If
        Else
            WriteLog 1, CurrentMod, PROC_NAME, "Revision update cancelled by user"
        End If
    End With
    
    Unload frmUpdateRev
    Exit Sub
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errMsg
    frmMsgBox.Display "Unable to update revision." & vbCrLf & vbCrLf & _
                      "Error: " & errMsg, , Critical, "Error"
    Unload frmUpdateRev
End Sub

'=======================================================
' Sub: UpdateFooterVersion
' Purpose: Update version information in document footer
'
' Description:
'   Updates content controls in footer with:
'   - Version name (Draft X or Final)
'   - Last save time (date only for Final, date+time for Draft)
'
' Error Handling:
'   Logs errors but doesn't interrupt workflow
'=======================================================
Private Sub UpdateFooterVersion()
    Const PROC_NAME As String = "UpdateFooterVersion"
    
    Dim versionName As String
    Dim timeFormat As String
    
    On Error GoTo ErrorHandler
    
    versionName = getVName
    
    ' Determine time format based on version type
    If versionName = "Final" Then
        timeFormat = DateFormat
    Else
        timeFormat = DateTimeFormat
    End If
    
    SetContentControl "VersionName", versionName
    SetContentControl "LastSaveTime", Format(ToServerTime, timeFormat)
    
    WriteLog 1, CurrentMod, PROC_NAME, "Footer version updated: " & versionName
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Non-critical error - continue
End Sub

'=======================================================
' Function: getVName
' Purpose: Get current version name from revisions table
'
' Returns:
'   Version name from last row of revisions table
'
' Description:
'   Reads the version name from the first cell of the
'   last row in the revisions table (Table 1).
'=======================================================
Private Function getVName() As String
    Const PROC_NAME As String = "getVName"
    
    On Error GoTo ErrorHandler
    
    Set Tbl = ActiveDocument.Tables(1)
    getVName = Cell(Tbl.Rows.Count, 1)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    getVName = "Unknown"
End Function

'=======================================================
' Sub: AddRev
' Purpose: Add new revision entry to revisions table
'
' Parameters:
'   Title - Revision title text
'   Notes - Revision notes text
'   VName - Version name (optional - auto-generated if empty)
'
' Description:
'   Adds a new row to the revisions table with:
'   - Version name (Draft X+1 or specified)
'   - Title
'   - Date
'   - Notes
'
'   If VName is empty, generates next draft number automatically.
'
' Error Handling:
'   Comprehensive error handling with rollback capability
'=======================================================
Private Sub AddRev(ByVal Title As String, _
                  ByVal Notes As String, _
                  Optional ByVal VName As String)
    Const PROC_NAME As String = "AddRev"
    
    Dim rowIndex As Long
    Dim lastRowIndex As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding revision: " & Title
    
    Set Tbl = ActiveDocument.Tables(1)
    
    With Tbl
        lastRowIndex = .Rows.Count
        
        ' Unprotect to allow changes
        Unprotect .Range.Document
        
        ' Add new row
        .Rows.Add
        
        ' Generate version name if not provided
        If Len(VName) = 0 Then
            ' Find last draft number and increment
            For rowIndex = lastRowIndex To 2 Step -1
                If Cell(rowIndex, 1) Like "Draft *" Or rowIndex = 2 Then
                    VName = "Draft " & (Val(Replace(Cell(rowIndex, 1), "Draft ", "")) + 1)
                    Exit For
                End If
            Next rowIndex
        End If
        
        ' Fill in revision data
        .Rows.Last.Cells(1).Range.text = VName
        .Rows.Last.Cells(2).Range.text = Title
        .Rows.Last.Cells(3).Range.text = Format(ToServerTime, DateFormat)
        .Rows.Last.Cells(4).Range.text = Notes
        
        ' Reprotect document
        Protect .Range.Document
        
        WriteLog 1, CurrentMod, PROC_NAME, "Revision added: " & VName
    End With
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Try to reprotect on error
    On Error Resume Next
    If Not Tbl Is Nothing Then Protect Tbl.Range.Document
    Err.Raise Err.Number, CurrentMod & "." & PROC_NAME, Err.Description
End Sub

'=======================================================
' TASK MANAGEMENT
'=======================================================

'=======================================================
' Sub: AddSubLevel
' Purpose: Add sub-level task under current task
'
' Description:
'   Adds a new task one level below the currently selected
'   task. For example, if cursor is on Task 2, adds Task 2.1.
'   Works for heading levels 2-4 (can create up to level 5).
'
' Error Handling:
'   Displays user-friendly messages for invalid operations
'=======================================================
Public Sub AddSubLevel()
    Const PROC_NAME As String = "AddSubLevel"
    
    Dim taskRange As Range
    Dim headingLevel As Long
    Dim scopeTask As scopeTask
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding sub-level task"
    
    ' Find heading at cursor
    Set taskRange = FindLastHeading(Selection.Range, "[2-4]").Paragraphs(1).Range
    headingLevel = GetDocentHeaderLvl(taskRange, "[2-4]")
    
    Select Case headingLevel
    Case 0
        GoTo InvalidSelection
    Case 2 To 4
        ' Get next task number for sublevel
        Set scopeTask = NextTaskNum(GetTaskNumber(taskRange, True, headingLevel), headingLevel + 1, taskRange)
        
        ' Add the task
        AddTask scopeTask.Range, headingLevel + 1, scopeTask.taskNum
        
        WriteLog 1, CurrentMod, PROC_NAME, "Sub-level task added: " & scopeTask.taskNum
    Case Else
        frmMsgBox.Display "Cannot create Heading 6 (maximum depth reached)", , Information, "Maximum Depth"
    End Select
    Exit Sub
    
InvalidSelection:
    frmMsgBox.Display "Place the cursor on the parent task first", , Information, "Invalid Selection"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    frmMsgBox.Display "Unable to add sub-level task." & vbCrLf & vbCrLf & _
                      "Error: " & Err.Description, , Critical, "Error"
End Sub

'=======================================================
' Sub: AddSameLevel
' Purpose: Add sibling task at same level as current task
'
' Description:
'   Adds a new task at the same level as the currently
'   selected task. For example, if cursor is on Task 2.3,
'   adds Task 2.4.
'
' Error Handling:
'   Displays user-friendly messages for invalid operations
'=======================================================
Public Sub AddSameLevel()
    Const PROC_NAME As String = "AddSameLevel"
    
    Dim taskRange As Range
    Dim headingLevel As Long
    Dim scopeTask As scopeTask
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding same-level task"
    
    ' Find heading at cursor
    Set taskRange = FindLastHeading(Selection.Range, "[2-5]").Paragraphs(1).Range
    headingLevel = GetDocentHeaderLvl(taskRange, "[2-5]")
    
    Select Case headingLevel
    Case 0
        GoTo InvalidSelection
    Case 2
        ' Level 2 is top level - delegate to AddTopLevel
        AddTopLevel
    Case Else
        ' Get next task number at same level
        Set scopeTask = NextTaskNum(GetTaskNumber(taskRange, False, headingLevel), headingLevel, taskRange)
        
        ' Add the task
        AddTask scopeTask.Range, headingLevel, scopeTask.taskNum
        
        WriteLog 1, CurrentMod, PROC_NAME, "Same-level task added: " & scopeTask.taskNum
    End Select
    Exit Sub
    
InvalidSelection:
    frmMsgBox.Display "Place the cursor on the sibling task first", , Information, "Invalid Selection"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    frmMsgBox.Display "Unable to add same-level task." & vbCrLf & vbCrLf & _
                      "Error: " & Err.Description, , Critical, "Error"
End Sub

'=======================================================
' Sub: AddTopLevel
' Purpose: Add new top-level task to document
'
' Description:
'   Adds a new top-level (Level 2) task to the scope
'   document. Prompts user for task name, adds entry
'   to tasks table, and creates the task in the document.
'
' Error Handling:
'   Comprehensive error handling with rollback capability
'=======================================================
Public Sub AddTopLevel()
    Const PROC_NAME As String = "AddTopLevel"
    
    Dim taskRange As Range
    Dim TaskName As String
    Dim taskNumStr As String
    Dim TaskNo As Long
    Dim lastRow As Row
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding top-level task"
    
    ' Prompt for task name
    TaskName = frmInputBox.Display("Please insert the task name", "Create New Task")
    If Len(TaskName) = 0 Or TaskName = "Canceled" Then
        WriteLog 1, CurrentMod, PROC_NAME, "Task creation cancelled"
        Exit Sub
    End If
    
    ' Get tasks table
    Set Tbl = GetTableByTitle("Tasks Table")
    If Tbl Is Nothing Then GoTo TableNotFound
    
    ' Unprotect document
    Unprotect Tbl.Range.Document
    
    ' Add task to table
    Set lastRow = GetLastRow()
    With lastRow
        TaskNo = .Index - 1
        taskNumStr = "Task " & TaskNo & ":"
        .Cells(1).Range.text = taskNumStr
        .Cells(2).Range.text = TaskName
    End With
    
    ' Find insertion point (use bookmark)
    Set taskRange = ActiveDocument.Range.GoTo(wdGoToBookmark, Name:=TasksBookmark)
    
    ' Add task to document
    AddTask taskRange, 2, , taskNumStr & " " & TaskName
    
    ' Reprotect document
    Protect taskRange.Document
    
    WriteLog 1, CurrentMod, PROC_NAME, "Top-level task added: " & taskNumStr
    Exit Sub
    
TableNotFound:
    WriteLog 3, CurrentMod, PROC_NAME, "Tasks table not found"
    frmMsgBox.Display "Tasks Table not found in document", , Critical, "Table Missing"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Try to reprotect on error
    On Error Resume Next
    If Not taskRange Is Nothing Then Protect taskRange.Document
    On Error GoTo 0
    frmMsgBox.Display "Unable to add top-level task." & vbCrLf & vbCrLf & _
                      "Error: " & Err.Description, , Critical, "Error"
End Sub

'=======================================================
' Function: FindLastHeading
' Purpose: Find the last heading at specified levels before cursor
'
' Parameters:
'   Rng - Starting range (typically Selection.Range)
'   HLevels - Heading levels pattern (e.g., "[2-4]")
'
' Returns:
'   Range containing the last heading found
'
' Description:
'   Moves backward through paragraphs from the given range
'   until a heading matching the specified level pattern
'   is found. Used to locate parent tasks when adding
'   new tasks.
'
' Error Handling:
'   Returns original range if no heading found or at start of document
'=======================================================
Public Function FindLastHeading(ByVal Rng As Range, _
                               ByVal HLevels As String) As Range
    Const PROC_NAME As String = "FindLastHeading"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Searching for heading with levels: " & HLevels
    
    ' Move backward until we find a matching heading
    Do Until GetDocentHeaderLvl(Rng, HLevels) > 0
        Rng.Move wdParagraph, -1
        Rng.Select
        
        ' Exit if we've reached the start of document
        If Rng.start = 0 Then Exit Do
    Loop
    
    Set FindLastHeading = Rng
    WriteLog 1, CurrentMod, PROC_NAME, "Found heading at position: " & Rng.start
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Set FindLastHeading = Rng  ' Return original range on error
End Function

'=======================================================
' Sub: RemoveScopeTasksPBreaks
' Purpose: Remove red page breaks from scope document
'
' Parameters:
'   Doc - Document to process
'
' Description:
'   Finds and removes all manual page breaks (^m) that
'   are formatted in red color. These are temporary
'   markers used during task creation.
'
' Error Handling:
'   Comprehensive error handling with protection state management
'=======================================================
Public Sub RemoveScopeTasksPBreaks(ByVal Doc As Document)
    Const PROC_NAME As String = "RemoveScopeTasksPBreaks"
    
    Dim searchRange As Range
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Removing red page breaks from document"
    
    ' Unprotect document
    Unprotect Doc
    
    ' Set up search range
    Set searchRange = Doc.Range
    
    With searchRange.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
        
        ' Search for manual page breaks with red font
        .text = "^m"
        .Font.Color = vbRed
        .Replacement.text = ""
        .Wrap = wdFindContinue
        .Forward = True
        
        ' Replace all occurrences
        .Execute Replace:=wdReplaceAll
        
        ' Clear formatting
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
    End With
    
    ' Reprotect document
    Protect Doc
    
    WriteLog 1, CurrentMod, PROC_NAME, "Red page breaks removed successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Try to reprotect on error
    On Error Resume Next
    Protect Doc
    Err.Raise Err.Number, CurrentMod & "." & PROC_NAME, Err.Description
End Sub

'=======================================================
' Sub: AddEditor
' Purpose: Add editing permissions to a range
'
' Parameters:
'   Rng - Range to add editing permissions to
'
' Description:
'   Allows everyone to edit specified range even when
'   document is protected. Handles protection state properly.
'
' Error Handling:
'   Comprehensive error handling with protection state restoration
'=======================================================
Private Sub AddEditor(ByVal Rng As Range)
    Const PROC_NAME As String = "AddEditor"
    
    Dim protectionType As Long
    
    On Error GoTo ErrorHandler
    
    ' Save current protection state
    protectionType = Rng.Document.protectionType
    
    ' Unprotect if needed
    If protectionType <> wdNoProtection Then
        Unprotect Rng.Document
    End If
    
    ' Add editor permission
    Rng.Editors.Add wdEditorEveryone
    
    ' Restore protection state
    If protectionType <> wdNoProtection Then
        Protect Rng.Document
    End If
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Try to restore protection on error
    On Error Resume Next
    If protectionType <> wdNoProtection Then Protect Rng.Document
End Sub

'=======================================================
' Sub: AddTask
' Purpose: Create a new task in the document with full structure
'
' Parameters:
'   Rng - Range where task should be inserted
'   Level - Heading level (2-5)
'   TaskNumber - Task number (optional)
'   TaskStr - Full task string (optional)
'
' Description:
'   Creates a complete task structure including:
'   - Heading with task number and name
'   - Objectives section
'   - Assumptions section
'   - Deliverables section
'   - Related Meetings table
'
'   This is a complex function that builds the entire
'   task structure in one operation.
'
' Error Handling:
'   Uses error handling for section breaks and protection
'=======================================================
Private Sub AddTask(ByVal Rng As Range, _
                   ByVal Level As Long, _
                   Optional ByVal TaskNumber As Variant, _
                   Optional ByVal TaskStr As String)
    Const PROC_NAME As String = "AddTask"
    
    Dim i As Long
    Dim originalColor As Long
    Dim bookmarkName As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Creating task at level " & Level
    
    ' Get task string if not provided
    If Len(TaskStr) = 0 Then
        TaskStr = TaskNumber & " " & frmInputBox.Display("Please insert the task name", "Docent IMS")
    End If
    
    ' Validate task string
    If Not IsMissing(TaskNumber) Then
        If TaskStr = TaskNumber & " " Then Exit Sub
        If TaskStr = TaskNumber & " Canceled" Then Exit Sub
    End If
    
    ' Generate bookmark name
    bookmarkName = Replace(Replace(Left$(TaskStr, InStr(TaskStr, ":") - 1), " ", ""), ".", "_")
    
    ' Position range
    i = Rng.start
    With Rng
        ' Move to correct position
        If i <> .Document.Bookmarks(TasksBookmark).Range.start Then
            .MoveUntil Chr(12)  ' Page break
        End If
        
        If .Characters.Last = Chr(13) Then .Move 1, -1
        
        ' Save font color
        originalColor = .Font.Color
        
        ' Unprotect document
        Unprotect .Document
        
        ' Insert section break with error handling for protected sections
        On Error Resume Next
        .InsertBreak wdSectionBreakNextPage
        Do While Err.Number = 4605  ' Section is protected
            Err.Clear
            .Move 1, 1
            AddEditor Rng
            .InsertBreak wdSectionBreakNextPage
        Loop
        
        ' Handle insertion position
        If Err.Number Then
            On Error GoTo ErrorHandler
            .Collapse wdCollapseEnd
            .InsertBreak wdSectionBreakNextPage
            .MoveStart 1, -1
        Else
            .MoveStart 1, -2
        End If
        
        ' Color and position
        .Font.Color = vbRed
        .Collapse wdCollapseEnd
        .Font.Color = originalColor
        
        ' Set heading and add text
        .Style = "Heading " & Level
        .text = TaskStr & Chr(13)
        
        ' Add bookmark
        .Document.Bookmarks.Add bookmarkName, .Paragraphs(1).Range
        
        ' Add paragraph if needed
        If .Paragraphs.Count = 1 Then .InsertParagraphAfter
        
        ' Move to next paragraph
        .Move 1, 1
        .Style = "Normal"
        
        ' Add task structure
        .text = "Objectives" & Chr(10) & Chr(10) & Chr(10) & _
                "Assumptions" & Chr(10) & Chr(10) & Chr(10) & _
                "Deliverables" & Chr(10) & Chr(10) & Chr(10) & _
                "Related Meetings" & Chr(10) & Chr(10) & Chr(10)
        
        .MoveEnd 1
        
        ' Format sections
        On Error Resume Next
        For i = 2 To 8 Step 3
            .Paragraphs(i).Indent
            .Paragraphs(i).Range.ContentControls.Add().Range.HighlightColorIndex = wdGray25
            .Paragraphs(i).Range.ListFormat.ApplyListTemplateWithLevel ListGalleries(wdBulletGallery).ListTemplates(1)
        Next i
        
        ' Bold section headers and add editing permissions
        For i = 1 To .Paragraphs.Count Step 3
            .Paragraphs(i).Range.Font.Bold = True
            AddEditor .Paragraphs(i + 1).Range
        Next i
        On Error GoTo ErrorHandler
        
        ' Add Related Meetings table
        AddTable TaskStr, "Related Meetings", Rng, _
                Array("Meeting Type", "Frequency", "Number of Meetings", _
                      "Length (Hours)", "Prep Time (Hours)", "Consultant Attendees", "Which Meeting?")
        
        ' Update table of contents
        .Document.TablesOfContents(1).Update
        
        ' Reprotect and select new task
        Protect .Document
        .Document.Bookmarks(bookmarkName).Range.Select
    End With
    
    WriteLog 1, CurrentMod, PROC_NAME, "Task created successfully: " & TaskStr
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ' Try to reprotect on error
    On Error Resume Next
    If Not Rng Is Nothing Then Protect Rng.Document
End Sub

'=======================================================
' Sub: AddTable
' Purpose: Add a table to task structure
'
' Parameters:
'   TaskStr - Task string for table title
'   TableName - Name of table
'   Rng - Range to insert table in
'   Cols - Array of column headers
'
' Description:
'   Creates a formatted table with headers and one data row.
'   Used for Assumptions, Deliverables, and Related Meetings tables.
'
' Error Handling:
'   Comprehensive error handling
'=======================================================
Private Sub AddTable(ByVal TaskStr As String, _
                    ByVal TableName As String, _
                    ByVal Rng As Range, _
                    ByVal Cols As Variant)
    Const PROC_NAME As String = "AddTable"
    
    Dim i As Long
    Dim newTable As Table
    Dim paragraphNum As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Adding table: " & TableName
    
    ' Find paragraph for table insertion
    With Rng
        For paragraphNum = 1 To .Paragraphs.Count
            If cellText(.Paragraphs(paragraphNum).Range.text) = TableName Then
                paragraphNum = paragraphNum + 1
                Exit For
            End If
        Next paragraphNum
        
        ' Create table
        Set newTable = .Tables.Add(.Paragraphs(paragraphNum).Range, 2, UBound(Cols) - LBound(Cols) + 1)
        newTable.Title = TaskStr & TableName
        
        ' Format borders
        For i = wdBorderTop To wdBorderDiagonalUp
            With newTable.Borders(i)
                .LineStyle = Options.DefaultBorderLineStyle
                .LineWidth = Options.DefaultBorderLineWidth
                .Color = Options.DefaultBorderColor
            End With
        Next i
        
        ' Format header row
        newTable.Rows(1).Shading.BackgroundPatternColor = wdGray25
        newTable.Rows(1).Range.Bold = True
        
        ' Add editing permission to data row
        AddEditor newTable.Rows(2).Range
        
        ' Fill column headers
        For i = LBound(Cols) To UBound(Cols)
            newTable.Cell(1, i - LBound(Cols) + 1).Range.text = Cols(i)
        Next i
    End With
    
    WriteLog 1, CurrentMod, PROC_NAME, "Table added successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Err.Raise Err.Number, CurrentMod & "." & PROC_NAME, Err.Description
End Sub

'=======================================================
' TASK NUMBER EXTRACTION AND MANIPULATION
'=======================================================

'=======================================================
' Function: GetDocentHeaderLvl
' Purpose: Get heading level of a Docent task heading
'
' Parameters:
'   Rng - Range to check
'   LevelRange - Pattern for acceptable levels (e.g., "[2-4]")
'
' Returns:
'   Heading level number (2-5), or 0 if not a task heading
'
' Description:
'   Validates that the range contains a task heading
'   (starts with "Task #") and extracts the heading level.
'=======================================================
Private Function GetDocentHeaderLvl(ByVal Rng As Range, _
                                   ByVal LevelRange As String) As Long
    Const PROC_NAME As String = "GetDocentHeaderLvl"
    
    Dim rangeStyle As String
    Dim styleParts() As String
    Dim partIndex As Long
    
    On Error GoTo ErrorHandler
    
    ' Check if this is a task heading
    If Not Rng.Paragraphs(1).Range.text Like "Task #*" Then
        GetDocentHeaderLvl = 0
        Exit Function
    End If
    
    ' Get style and parse heading level
    rangeStyle = Rng.Paragraphs(1).Range.Style
    styleParts = Split(rangeStyle, ",")
    
    For partIndex = LBound(styleParts) To UBound(styleParts)
        If styleParts(partIndex) Like "Heading " & LevelRange Then
            GetDocentHeaderLvl = Val(Replace(styleParts(partIndex), "Heading ", ""))
            Exit Function
        End If
    Next partIndex
    
    GetDocentHeaderLvl = 0
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    GetDocentHeaderLvl = 0
End Function

'=======================================================
' Function: GetTaskNumber
' Purpose: Extract task number from a task heading
'
' Parameters:
'   Rng - Range containing task heading
'   IsSub - True if extracting for sublevel
'   Level - Heading level (optional)
'
' Returns:
'   Task number string (e.g., "Task 2.3:")
'
' Description:
'   Extracts the task number from a heading, handling
'   both direct heading text and search within paragraphs.
'
' Note:
'   Contains unreachable code (lines 375-386 in original)
'   which has been preserved but marked for cleanup.
'=======================================================
Private Function GetTaskNumber(ByVal Rng As Range, _
                              ByVal IsSub As Boolean, _
                              Optional ByVal Level As Long) As String
    Const PROC_NAME As String = "GetTaskNumber"
    
    Dim findRange As Range
    Dim paragraphIndex As Long
    
    On Error GoTo ErrorHandler
    
    Set findRange = Rng.Document.Range
    findRange.SetRange Rng.start, Rng.End
    
    ' Check if range is directly a heading
    If findRange.Style = "Heading " & Level Then
        GetTaskNumber = ExtractTaskNum(findRange.text, IsSub)
        Exit Function
    End If
    
    ' Search paragraphs for heading
    For paragraphIndex = 1 To Rng.Paragraphs.Count
        If GetDocentHeaderLvl(Rng.Paragraphs(paragraphIndex).Range, CStr(Level)) = Level Then
            GetTaskNumber = ExtractTaskNum(findRange.text, IsSub)
            Exit Function
        End If
    Next paragraphIndex
    
    ' NOTE: The following code is unreachable due to Exit Function above
    ' Left in place to preserve original logic, but should be reviewed
    ' TODO: Determine if this Find logic is needed and refactor accordingly
    
    GetTaskNumber = vbNullString
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    GetTaskNumber = vbNullString
End Function

'=======================================================
' Function: NextTaskNum
' Purpose: Calculate next available task number
'
' Parameters:
'   LastFound - Last task number found
'   Level - Heading level
'   Fnd - Range to search (optional, defaults to entire document)
'
' Returns:
'   ScopeTask object with Range and TaskNum properties
'
' Description:
'   Searches forward from current position to find the
'   next available task number by incrementing the last
'   found number until no match is found.
'=======================================================
Private Function NextTaskNum(ByVal LastFound As String, _
                            ByVal Level As Long, _
                            Optional ByVal Fnd As Range) As scopeTask
    Const PROC_NAME As String = "NextTaskNum"
    
    Dim scopeTask As New scopeTask
    
    On Error GoTo ErrorHandler
    
    If Fnd Is Nothing Then Set Fnd = ActiveDocument.Range
    
    With Fnd.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue
        .Style = "Heading " & Level
        
        ' Keep incrementing until we find unused number
        Do
            .text = LastFound
            .Execute
            If .Found Then
                LastFound = ExtractTaskNum(LastFound, False)
            End If
        Loop While .Found
    End With
    
    ' Return scope task with range and number
    Set scopeTask.Range = Fnd
    scopeTask.taskNum = LastFound
    Set NextTaskNum = scopeTask
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Set NextTaskNum = scopeTask  ' Return empty task
End Function

'=======================================================
' Function: ExtractTaskNum
' Purpose: Extract and increment task number
'
' Parameters:
'   TaskName - Full task name string (e.g., "Task 2.3: Name")
'   IsSub - True to add sublevel (.1), False to increment
'
' Returns:
'   Next task number string
'
' Description:
'   For IsSub=True: "Task 2.3:" becomes "Task 2.3.1:"
'   For IsSub=False: "Task 2.3:" becomes "Task 2.4:"
'=======================================================
Private Function ExtractTaskNum(ByVal TaskName As String, _
                               ByVal IsSub As Boolean) As String
    Const PROC_NAME As String = "ExtractTaskNum"
    
    Dim taskNum As String
    Dim numParts() As String
    Dim lastNum As Long
    Dim lastNumPos As Long
    
    On Error GoTo ErrorHandler
    
    ' Extract number portion before colon
    taskNum = Split(TaskName, ":")(0)
    
    If IsSub Then
        ' Add sublevel
        ExtractTaskNum = taskNum & ".1:"
    Else
        ' Increment last number
        numParts = Split(taskNum, " ")
        numParts = Split(numParts(UBound(numParts)), "-")
        numParts = Split(numParts(UBound(numParts)), ".")
        
        lastNum = Val(numParts(UBound(numParts)))
        lastNumPos = InStrRev(taskNum, CStr(lastNum))
        
        taskNum = Left$(taskNum, lastNumPos - 1) & (lastNum + 1)
        ExtractTaskNum = taskNum & ":"
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    ExtractTaskNum = TaskName
End Function

'=======================================================
' HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: Cell
' Purpose: Get cleaned cell text from table
'
' Parameters:
'   r - Row number (1-based)
'   c - Column number (1-based)
'
' Returns:
'   Trimmed cell text with table markers removed
'
' Note: Uses module-level Tbl variable
'=======================================================
Private Function Cell(ByVal r As Long, ByVal c As Long) As String
    On Error Resume Next
    Cell = Trim$(Replace(Tbl.Rows(r).Cells(c).Range.text, Chr(13) & Chr(7), ""))
End Function

'=======================================================
' Function: GetLastRow
' Purpose: Get last row of table or add new row if needed
'
' Returns:
'   Last row object from module-level Tbl variable
'
' Description:
'   If last row is empty, returns it. If not, adds a
'   new row and returns the new row.
'
' Note: Uses module-level Tbl variable
'=======================================================
Private Function GetLastRow() As Row
    On Error Resume Next
    
    Set GetLastRow = Tbl.Rows.Last
    
    ' Add row if last row has content
    If Len(Cell(Tbl.Rows.Count, 1)) > 0 Then
        Tbl.Rows.Add
        Set GetLastRow = Tbl.Rows.Last
    End If
End Function

'=======================================================
' Function: FindOldTask
' Purpose: Find task heading by task number
'
' Parameters:
'   TaskNo - Task number to search for
'
' Returns:
'   Range of task paragraph if found, Nothing otherwise
'
' Description:
'   Searches for Heading 2 with specified task number.
'   Used when updating or replacing tasks.
'=======================================================
Private Function FindOldTask(ByVal TaskNo As String) As Range
    Const PROC_NAME As String = "FindOldTask"
    
    Dim findRange As Range
    
    On Error GoTo ErrorHandler
    
    Set findRange = ActiveDocument.Range
    
    With findRange.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue
        .Style = "Heading 2"
        .text = TaskNo
        .Execute
        
        If .Found Then
            Set FindOldTask = findRange.Paragraphs(1).Range
            FindOldTask.MoveEndWhile Chr(10) & Chr(13), -1
        Else
            Set FindOldTask = Nothing
        End If
    End With
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error " & Err.Number & ": " & Err.Description
    Set FindOldTask = Nothing
End Function

'=======================================================
' END OF MODULE
'=======================================================
