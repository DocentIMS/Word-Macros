Attribute VB_Name = "Ribbon_Tasks_Mod"
Option Explicit
Option Private Module

'=======================================================
' Module: Ribbon_Tasks_Mod
' Purpose: Ribbon callbacks for Tasks group
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Handles all ribbon callbacks for the Tasks group,
'   including visibility, images, tooltips, and actions
'   for traffic light task indicators (Green/Yellow/Red)
'   and task creation button.
'
' Dependencies:
'   - Ribbon_Functions_Mod (GetVisibleGroup, GetTaskImage, GetTaskCount, GetTasksTrafficTooltip, GotoTaskCollection)
'   - AB_GlobalVars (ProjectColorStr)
'   - frmCreateTask
'
' Ribbon Callbacks:
'   - IdGroupTasksGetVisible - Group visibility
'   - IdButtonTasks[Color]GetImage - Traffic light images
'   - IdButtonTasks[Color]GetEnabled - Traffic light enabled state
'   - IdButtonTasks[Color]OnAction - Traffic light actions
'   - IdButtonTasks[Color]GetSupertip - Traffic light tooltips
'   - IdButtonCreateTaskOnAction - Create task action
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added logging
'       * Added function documentation
'       * Improved null/error checks
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Ribbon_Tasks_Mod"

'=======================================================
' RIBBON GROUP VISIBILITY
'=======================================================

'=======================================================
' Function: IdGroupTasksGetVisible
' Purpose: Determine if Tasks group should be visible
'
' Parameters:
'   ID - Ribbon group ID
'
' Returns:
'   Boolean - True if group should be visible
'
' Error Handling:
'   - Returns False on any error
'   - Logs errors
'=======================================================
Function IdGroupTasksGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdGroupTasksGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdGroupTasksGetVisible = GetVisibleGroup(ID)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdGroupTasksGetVisible = False
End Function

'=======================================================
' GREEN TRAFFIC LIGHT CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonTasksGreenGetImage
' Purpose: Get image for green traffic light button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonTasksGreenGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonTasksGreenGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonTasksGreenGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonTasksGreenGetImage = GetTaskImage("Green")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonTasksGreenGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonTasksGreenGetEnabled
' Purpose: Determine if green traffic light button is enabled
'
' Returns:
'   Boolean - True if any green tasks exist
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksGreenGetEnabled() As Boolean
    Const PROC_NAME As String = "IdButtonTasksGreenGetEnabled"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksGreenGetEnabled = (GetTaskCount("Green") > 0)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksGreenGetEnabled = False
End Function

'=======================================================
' Sub: IdButtonTasksGreenOnAction
' Purpose: Handle green traffic light button click
'
' Description:
'   Navigates to collection of green (future) tasks
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonTasksGreenOnAction()
    Const PROC_NAME As String = "IdButtonTasksGreenOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Green tasks button clicked"
    GotoTaskCollection "Green"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display green tasks: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonTasksGreenGetSupertip
' Purpose: Get tooltip for green traffic light button
'
' Returns:
'   String - Tooltip text
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksGreenGetSupertip() As String
    Const PROC_NAME As String = "IdButtonTasksGreenGetSupertip"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksGreenGetSupertip = GetTasksTrafficTooltip("Green", "Future")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksGreenGetSupertip = "Future tasks"
End Function

'=======================================================
' YELLOW TRAFFIC LIGHT CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonTasksYellowGetImage
' Purpose: Get image for yellow traffic light button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonTasksYellowGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonTasksYellowGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonTasksYellowGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonTasksYellowGetImage = GetTaskImage("Yellow")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonTasksYellowGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonTasksYellowGetEnabled
' Purpose: Determine if yellow traffic light button is enabled
'
' Returns:
'   Boolean - True if any yellow tasks exist
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksYellowGetEnabled() As Boolean
    Const PROC_NAME As String = "IdButtonTasksYellowGetEnabled"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksYellowGetEnabled = (GetTaskCount("Yellow") > 0)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksYellowGetEnabled = False
End Function

'=======================================================
' Sub: IdButtonTasksYellowOnAction
' Purpose: Handle yellow traffic light button click
'
' Description:
'   Navigates to collection of yellow (soon) tasks
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonTasksYellowOnAction()
    Const PROC_NAME As String = "IdButtonTasksYellowOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Yellow tasks button clicked"
    GotoTaskCollection "Yellow"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display yellow tasks: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonTasksYellowGetSupertip
' Purpose: Get tooltip for yellow traffic light button
'
' Returns:
'   String - Tooltip text
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksYellowGetSupertip() As String
    Const PROC_NAME As String = "IdButtonTasksYellowGetSupertip"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksYellowGetSupertip = GetTasksTrafficTooltip("Yellow", "Soon")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksYellowGetSupertip = "Tasks due soon"
End Function

'=======================================================
' RED TRAFFIC LIGHT CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonTasksRedGetImage
' Purpose: Get image for red traffic light button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonTasksRedGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonTasksRedGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonTasksRedGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonTasksRedGetImage = GetTaskImage("Red")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonTasksRedGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonTasksRedGetEnabled
' Purpose: Determine if red traffic light button is enabled
'
' Returns:
'   Boolean - True if any red tasks exist
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksRedGetEnabled() As Boolean
    Const PROC_NAME As String = "IdButtonTasksRedGetEnabled"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksRedGetEnabled = (GetTaskCount("Red") > 0)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksRedGetEnabled = False
End Function

'=======================================================
' Sub: IdButtonTasksRedOnAction
' Purpose: Handle red traffic light button click
'
' Description:
'   Navigates to collection of red (urgent) tasks
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonTasksRedOnAction()
    Const PROC_NAME As String = "IdButtonTasksRedOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Red tasks button clicked"
    GotoTaskCollection "Red"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display red tasks: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonTasksRedGetSupertip
' Purpose: Get tooltip for red traffic light button
'
' Returns:
'   String - Tooltip text
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonTasksRedGetSupertip() As String
    Const PROC_NAME As String = "IdButtonTasksRedGetSupertip"
    
    On Error GoTo ErrorHandler
    
    IdButtonTasksRedGetSupertip = GetTasksTrafficTooltip("Red", "Urgent")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTasksRedGetSupertip = "Urgent tasks"
End Function

'=======================================================
' CREATE TASK BUTTON
'=======================================================

'=======================================================
' Sub: IdButtonCreateTaskOnAction
' Purpose: Handle create task button click
'
' Description:
'   Opens the task creation form
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Sub IdButtonCreateTaskOnAction()
    Const PROC_NAME As String = "IdButtonCreateTaskOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Create task button clicked"
    
    ' Display task creation form
    frmCreateTask.Display True
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to open task creation form: " & Err.Description, vbExclamation, "Error"
End Sub
