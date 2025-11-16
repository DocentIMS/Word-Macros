Attribute VB_Name = "Tools_WordBoost"
Option Explicit

'=======================================================
' Module: Tools_WordBoost
' Purpose: Application performance optimization utilities
' Author: Docent IMS Team
' Version: 2.0
'
' Description:
'   Provides functions to temporarily disable screen updating
'   and alerts to improve performance during bulk operations.
'   Uses reference counting to safely handle nested boost calls.
'
'   The module also includes timing utilities for performance
'   profiling during development.
'
' Usage:
'   Boost True          ' Start performance boost
'   ' ... perform operations ...
'   Boost False         ' End performance boost
'
'   ' For nested operations:
'   Boost True          ' BCount = 1
'     Boost True        ' BCount = 2
'     Boost False       ' BCount = 1 (still boosted)
'   Boost False         ' BCount = 0 (unboosted)
'
'   ' Force immediate restore:
'   Boost False, True   ' Resets counter and restores immediately
'
' Module Variables:
'   BCount - Reference counter for nested boost calls
'   LastT - Last timer value for PrintTimer function
'   DebugMode - Enable/disable debug timing output
'
' Dependencies:
'   - AB_GlobalVars (for WriteLog if available)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive module documentation
'       * Replaced On Error Resume Next with proper error handling
'       * Added error recovery that always restores UI state
'       * Added function documentation headers
'       * Removed unused variable (Fnd As Range)
'       * Added logging support
'       * Fixed typo in comments ("Falst" -> "False")
'   v1.0 - Original version (23-06-16 23:59)
'=======================================================

' Module constants
Private Const CurrentMod As String = "Tools_WordBoost"

' Module-level variables
Private BCount As Long          ' Boost reference counter
Private LastT As Single         ' Last timer value for performance tracking
Private DebugMode As Boolean    ' Enable debug output

'=======================================================
' Sub: Boost
' Purpose: Optimize or restore application performance settings
'
' Description:
'   Temporarily disables screen updating and alerts to improve
'   performance during bulk operations. Uses reference counting
'   to safely handle nested boost calls, ensuring the UI is only
'   restored when all operations complete.
'
'   On the first boost call (BCount=0->1), screen updating and
'   alerts are disabled. On the final unboost call (BCount=1->0),
'   they are re-enabled.
'
' Parameters:
'   Flag - True to boost performance, False to restore (default: True)
'   ForceUnboost - Force immediate restore regardless of call depth (default: False)
'
' Usage:
'   ' Simple usage
'   Boost True          ' Disable updates
'   ' ... do work ...
'   Boost False         ' Re-enable updates
'
'   ' Nested calls
'   Boost True          ' BCount = 1, updates disabled
'     Boost True        ' BCount = 2, still disabled
'     Boost False       ' BCount = 1, still disabled
'   Boost False         ' BCount = 0, updates re-enabled
'
'   ' Emergency restore
'   Boost False, True   ' Force restore regardless of counter
'
' Error Handling:
'   If an error occurs, always restores UI state to prevent lockup
'
' Note:
'   This module was originally designed for Excel but is used
'   with Word. Application references work for both.
'=======================================================
Public Sub Boost(Optional ByVal Flag As Boolean = True, _
                 Optional ByVal ForceUnboost As Boolean = False)
    
    Const PROC_NAME As String = "Boost"
    
    On Error GoTo ErrorHandler
    
    ' Handle force unboost
    If ForceUnboost Then
        BCount = 0
        Flag = False
        WriteLog 1, CurrentMod, PROC_NAME, "Force unboost requested"
    End If
    
    ' Update reference counter
    If Flag Then
        BCount = BCount + 1
        WriteLog 1, CurrentMod, PROC_NAME, "Boost count increased to " & BCount
    Else
        BCount = BCount - 1
        WriteLog 1, CurrentMod, PROC_NAME, "Boost count decreased to " & BCount
    End If
    
    ' Prevent negative counter
    If BCount < 0 Then BCount = 0
    
    ' Apply boost on first call (0->1)
    If Flag And BCount = 1 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        WriteLog 1, CurrentMod, PROC_NAME, "Performance boost enabled (screen updating OFF)"
        
    ' Restore on final unboost (1->0)
    ElseIf BCount = 0 Then
        Application.StatusBar = vbNullString
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        WriteLog 1, CurrentMod, PROC_NAME, "Performance boost disabled (screen updating ON)"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' CRITICAL: Always restore UI state on error to prevent lockup
    On Error Resume Next  ' Prevent cascading errors during cleanup
    
    Application.StatusBar = vbNullString
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    BCount = 0
    
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & _
             " - UI state restored"
End Sub

'=======================================================
' Sub: EndAll
' Purpose: Force immediate unboost and terminate execution
'
' Description:
'   Emergency stop function that forces UI restoration
'   and terminates code execution. Use only for critical
'   error situations where normal flow cannot continue.
'
' Warning:
'   This sub calls End statement which terminates VBA
'   execution immediately. Use with caution.
'
' Usage:
'   If criticalError Then
'       EndAll
'   End If
'=======================================================
Public Sub EndAll()
    Const PROC_NAME As String = "EndAll"
    
    On Error Resume Next  ' Ensure we can clean up even on error
    
    WriteLog 2, CurrentMod, PROC_NAME, "Emergency stop requested - forcing unboost"
    
    ' Force restore UI state
    Boost False, True
    
    ' Terminate execution
    End
End Sub

'=======================================================
' Sub: PrintTimer
' Purpose: Print elapsed time since last call (debug utility)
'
' Description:
'   Development/debugging utility that prints the elapsed
'   time since the previous call. Only active when DebugMode
'   is True. Useful for performance profiling.
'
' Parameters:
'   status - Description text to print with timing
'
' Usage:
'   PrintTimer "Started process"
'   ' ... do work ...
'   PrintTimer "Completed step 1"  ' Prints: "2.34: Completed step 1"
'   ' ... do more work ...
'   PrintTimer "Completed step 2"  ' Prints: "1.56: Completed step 2"
'
' Note:
'   To enable timing output, set DebugMode = True in the
'   Immediate window or add initialization code.
'
' Output Format:
'   "X.XX: <status text>"
'   where X.XX is elapsed seconds since last call
'=======================================================
Public Sub PrintTimer(ByVal status As String)
    Const PROC_NAME As String = "PrintTimer"
    
    On Error GoTo ErrorHandler
    
    ' Exit if debug mode not enabled
    If Not DebugMode Then Exit Sub
    
    ' Print elapsed time if we have a previous timestamp
    If LastT > 0 Then
        If Len(status) > 0 Then
            Debug.Print Format$(Timer - LastT, "0.00") & ": " & status
        End If
    End If
    
    ' Store current time for next call
    LastT = Timer
    Exit Sub
    
ErrorHandler:
    ' Don't log timer errors - would create infinite loop
    ' Just reset and continue
    LastT = Timer
End Sub

'=======================================================
' Sub: EnableDebugMode
' Purpose: Enable debug timing output
'
' Description:
'   Public helper to enable PrintTimer output.
'   Call this at the start of your debugging session.
'
' Usage:
'   EnableDebugMode
'   PrintTimer "Starting operations"
'=======================================================
Public Sub EnableDebugMode()
    DebugMode = True
    LastT = Timer
    Debug.Print "Debug mode enabled - timing started"
End Sub

'=======================================================
' Sub: DisableDebugMode
' Purpose: Disable debug timing output
'
' Description:
'   Public helper to disable PrintTimer output.
'
' Usage:
'   DisableDebugMode
'=======================================================
Public Sub DisableDebugMode()
    DebugMode = False
    Debug.Print "Debug mode disabled"
End Sub

'=======================================================
' END OF MODULE
'=======================================================
