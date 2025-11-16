Attribute VB_Name = "Globals_StateManager"
Option Explicit
Option Private Module

'=======================================================
' Module: Globals_StateManager
' Purpose: Centralized application state management
' Author: Created - November 2025 (Critical Improvement #4)
' Version: 1.0
'
' Description:
'   Provides controlled access to application state through
'   getter/setter properties. This module replaces direct
'   access to global variables throughout the application,
'   improving encapsulation and making state changes trackable.
'
' Benefits:
'   ✓ Centralized state management
'   ✓ Input validation on state changes
'   ✓ Logging of important state transitions
'   ✓ Prevents namespace pollution
'   ✓ Makes debugging easier
'   ✓ Thread-safe preparation for future
'
' Migration Notes:
'   Replace direct global variable access with properties:
'   - BusyRibbon → StateManager.RibbonBusy
'   - CodeIsRunning → StateManager.CodeRunning
'   - PrjMgr → StateManager.IsProjectManager
'   - Set_UseBookmarks → StateManager.UseBookmarks
'   etc.
'
' Usage Example:
'   ' Old way:
'   BusyRibbon = True
'
'   ' New way:
'   StateManager.RibbonBusy = True
'
' Change Log:
'   v1.0 - Nov 2025 - Initial creation
'=======================================================

Private Const CurrentMod As String = "Globals_StateManager"

'=======================================================
' PRIVATE STATE VARIABLES
'=======================================================

' UI State
Private m_PrjMgr As Boolean
Private m_BusyRibbon As Boolean
Private m_CodeIsRunning As Boolean

' Settings State
Private m_UseBookmarks As Boolean
Private m_Coloring As Boolean
Private m_Indenting As Boolean
Private m_Export As Boolean
Private m_BoldToo As Boolean
Private m_OutputDir As String
Private m_Cancelled As Boolean

' Dictionary State
Private m_DocentDictionaryPath As String
Private m_DocentDictionary As Word.Dictionary

'=======================================================
' UI STATE PROPERTIES
'=======================================================

'=======================================================
' Property: IsProjectManager
' Purpose: Get/Set project manager mode
'
' Description:
'   Indicates whether user is in project manager mode.
'   Affects ribbon visibility and available features.
'=======================================================
Public Property Get IsProjectManager() As Boolean
    IsProjectManager = m_PrjMgr
End Property

Public Property Let IsProjectManager(ByVal value As Boolean)
    Const PROC_NAME As String = "IsProjectManager"
    
    On Error GoTo ErrorHandler
    
    ' Log state change if different
    If m_PrjMgr <> value Then
        WriteLog 1, CurrentMod, PROC_NAME, _
                 "Changed from " & m_PrjMgr & " to " & value
        m_PrjMgr = value
    End If
    
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Property

'=======================================================
' Property: RibbonBusy
' Purpose: Get/Set ribbon busy state
'
' Description:
'   Indicates whether ribbon is currently processing.
'   Used to prevent concurrent operations.
'=======================================================
Public Property Get RibbonBusy() As Boolean
    RibbonBusy = m_BusyRibbon
End Property

Public Property Let RibbonBusy(ByVal value As Boolean)
    Const PROC_NAME As String = "RibbonBusy"
    
    On Error GoTo ErrorHandler
    
    ' Log if setting busy when already busy
    If value And m_BusyRibbon Then
        WriteLog 2, CurrentMod, PROC_NAME, _
                 "Setting ribbon busy when already busy - potential re-entrance"
    End If
    
    m_BusyRibbon = value
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Property

'=======================================================
' Property: CodeRunning
' Purpose: Get/Set code execution state
'
' Description:
'   Indicates whether code is currently executing.
'   Used to prevent re-entrant calls.
'=======================================================
Public Property Get CodeRunning() As Boolean
    CodeRunning = m_CodeIsRunning
End Property

Public Property Let CodeRunning(ByVal value As Boolean)
    Const PROC_NAME As String = "CodeRunning"
    
    On Error GoTo ErrorHandler
    
    ' Log warning if setting running when already running
    If value And m_CodeIsRunning Then
        WriteLog 2, CurrentMod, PROC_NAME, _
                 "Code already running - potential re-entrance detected"
    End If
    
    m_CodeIsRunning = value
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Property

'=======================================================
' SETTINGS STATE PROPERTIES
'=======================================================

'=======================================================
' Property: UseBookmarks
' Purpose: Get/Set bookmark usage in document processing
'=======================================================
Public Property Get UseBookmarks() As Boolean
    UseBookmarks = m_UseBookmarks
End Property

Public Property Let UseBookmarks(ByVal value As Boolean)
    m_UseBookmarks = value
End Property

'=======================================================
' Property: Coloring
' Purpose: Get/Set coloring in document processing
'=======================================================
Public Property Get Coloring() As Boolean
    Coloring = m_Coloring
End Property

Public Property Let Coloring(ByVal value As Boolean)
    m_Coloring = value
End Property

'=======================================================
' Property: Indenting
' Purpose: Get/Set indenting in document processing
'=======================================================
Public Property Get Indenting() As Boolean
    Indenting = m_Indenting
End Property

Public Property Let Indenting(ByVal value As Boolean)
    m_Indenting = value
End Property

'=======================================================
' Property: ExportEnabled
' Purpose: Get/Set export flag
'=======================================================
Public Property Get ExportEnabled() As Boolean
    ExportEnabled = m_Export
End Property

Public Property Let ExportEnabled(ByVal value As Boolean)
    m_Export = value
End Property

'=======================================================
' Property: BoldFormatting
' Purpose: Get/Set bold formatting flag
'=======================================================
Public Property Get BoldFormatting() As Boolean
    BoldFormatting = m_BoldToo
End Property

Public Property Let BoldFormatting(ByVal value As Boolean)
    m_BoldToo = value
End Property

'=======================================================
' Property: OutputDirectory
' Purpose: Get/Set output directory for exports
'
' Description:
'   Validates directory exists when setting.
'=======================================================
Public Property Get OutputDirectory() As String
    OutputDirectory = m_OutputDir
End Property

Public Property Let OutputDirectory(ByVal value As String)
    Const PROC_NAME As String = "OutputDirectory"
    
    Dim FSO As Object
    
    On Error GoTo ErrorHandler
    
    ' Validate directory exists if not empty
    If Len(Trim$(value)) > 0 Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        If Not FSO.FolderExists(value) Then
            WriteLog 2, CurrentMod, PROC_NAME, _
                     "Directory does not exist: " & value
        End If
        
        Set FSO = Nothing
    End If
    
    m_OutputDir = value
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error validating directory: " & Err.Description
    
    ' Cleanup
    On Error Resume Next
    Set FSO = Nothing
End Property

'=======================================================
' Property: OperationCancelled
' Purpose: Get/Set operation cancelled flag
'
' Description:
'   Indicates whether user cancelled current operation.
'   Automatically logs when set to True.
'=======================================================
Public Property Get OperationCancelled() As Boolean
    OperationCancelled = m_Cancelled
End Property

Public Property Let OperationCancelled(ByVal value As Boolean)
    Const PROC_NAME As String = "OperationCancelled"
    
    m_Cancelled = value
    
    If value Then
        WriteLog 1, CurrentMod, PROC_NAME, "Operation cancelled by user"
    End If
End Property

'=======================================================
' DICTIONARY STATE PROPERTIES
'=======================================================

'=======================================================
' Property: DocentDictionaryPath
' Purpose: Get/Set custom dictionary path
'
' Description:
'   Validates file exists when setting.
'=======================================================
Public Property Get DocentDictionaryPath() As String
    DocentDictionaryPath = m_DocentDictionaryPath
End Property

Public Property Let DocentDictionaryPath(ByVal value As String)
    Const PROC_NAME As String = "DocentDictionaryPath"
    
    Dim FSO As Object
    
    On Error GoTo ErrorHandler
    
    ' Validate file exists if not empty
    If Len(Trim$(value)) > 0 Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        
        If Not FSO.FileExists(value) Then
            WriteLog 2, CurrentMod, PROC_NAME, _
                     "Dictionary file does not exist: " & value
        End If
        
        Set FSO = Nothing
    End If
    
    m_DocentDictionaryPath = value
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error validating dictionary path: " & Err.Description
    
    ' Cleanup
    On Error Resume Next
    Set FSO = Nothing
End Property

'=======================================================
' Property: DocentDictionary
' Purpose: Get/Set custom dictionary object
'
' Description:
'   Manages Word dictionary object. Properly cleans up
'   old dictionary before setting new one.
'=======================================================
Public Property Get DocentDictionary() As Word.Dictionary
    Set DocentDictionary = m_DocentDictionary
End Property

Public Property Set DocentDictionary(ByVal value As Word.Dictionary)
    Const PROC_NAME As String = "DocentDictionary"
    
    On Error GoTo ErrorHandler
    
    ' Cleanup old dictionary if exists
    If Not m_DocentDictionary Is Nothing Then
        Set m_DocentDictionary = Nothing
    End If
    
    ' Set new dictionary
    Set m_DocentDictionary = value
    
    Exit Property
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error setting dictionary: " & Err.Description
    
    ' Cleanup
    On Error Resume Next
    Set m_DocentDictionary = Nothing
End Property

'=======================================================
' STATE RESET METHODS
'=======================================================

'=======================================================
' Sub: ResetSettings
' Purpose: Reset all settings to default values
'
' Description:
'   Resets all document processing settings to their
'   default state.
'=======================================================
Public Sub ResetSettings()
    Const PROC_NAME As String = "ResetSettings"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Resetting all settings to defaults"
    
    m_UseBookmarks = False
    m_Coloring = True
    m_Indenting = True
    m_Export = False
    m_BoldToo = False
    m_Cancelled = False
    m_OutputDir = ""
    
    WriteLog 1, CurrentMod, PROC_NAME, "Settings reset completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: ResetUIState
' Purpose: Reset UI state flags
'
' Description:
'   Resets ribbon busy and code running flags to False.
'   Use when recovering from errors.
'=======================================================
Public Sub ResetUIState()
    Const PROC_NAME As String = "ResetUIState"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Resetting UI state"
    
    m_BusyRibbon = False
    m_CodeIsRunning = False
    
    WriteLog 1, CurrentMod, PROC_NAME, "UI state reset completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: Cleanup
' Purpose: Cleanup all resources
'
' Description:
'   Releases all object references and resets state.
'   Should be called when closing application or
'   reinitializing.
'=======================================================
Public Sub Cleanup()
    Const PROC_NAME As String = "Cleanup"
    
    On Error Resume Next
    
    WriteLog 1, CurrentMod, PROC_NAME, "Cleaning up state manager"
    
    ' Release dictionary
    Set m_DocentDictionary = Nothing
    
    ' Reset all state
    Call ResetSettings
    Call ResetUIState
    
    ' Reset project manager mode
    m_PrjMgr = False
    
    WriteLog 1, CurrentMod, PROC_NAME, "Cleanup completed"
End Sub

'=======================================================
' VALIDATION HELPERS
'=======================================================

'=======================================================
' Function: ValidateState
' Purpose: Validate application is in valid state
'
' Returns:
'   True if state is valid, False otherwise
'
' Description:
'   Checks for invalid state combinations that could
'   indicate bugs or issues.
'=======================================================
Public Function ValidateState() As Boolean
    Const PROC_NAME As String = "ValidateState"
    
    Dim IsValid As Boolean
    Dim issues As String
    
    On Error GoTo ErrorHandler
    
    IsValid = True
    issues = ""
    
    ' Check for invalid combinations
    If m_BusyRibbon And Not m_CodeIsRunning Then
        issues = issues & "- Ribbon busy but code not running" & vbCrLf
        IsValid = False
    End If
    
    ' Check dictionary state
    If Len(m_DocentDictionaryPath) > 0 And m_DocentDictionary Is Nothing Then
        issues = issues & "- Dictionary path set but object is Nothing" & vbCrLf
        IsValid = False
    End If
    
    ' Check output directory
    If m_Export And Len(m_OutputDir) = 0 Then
        issues = issues & "- Export enabled but no output directory" & vbCrLf
        IsValid = False
    End If
    
    ' Log issues if found
    If Not IsValid Then
        WriteLog 2, CurrentMod, PROC_NAME, _
                 "State validation failed:" & vbCrLf & issues
    End If
    
    ValidateState = IsValid
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ValidateState = False
End Function

'=======================================================
' Function: GetStateSnapshot
' Purpose: Get snapshot of current state for debugging
'
' Returns:
'   String representation of current state
'
' Description:
'   Returns formatted string showing all state variables.
'   Useful for debugging and logging.
'=======================================================
Public Function GetStateSnapshot() As String
    Const PROC_NAME As String = "GetStateSnapshot"
    
    Dim snapshot As String
    
    On Error GoTo ErrorHandler
    
    snapshot = "=== STATE SNAPSHOT ===" & vbCrLf
    snapshot = snapshot & "UI State:" & vbCrLf
    snapshot = snapshot & "  Project Manager: " & m_PrjMgr & vbCrLf
    snapshot = snapshot & "  Ribbon Busy: " & m_BusyRibbon & vbCrLf
    snapshot = snapshot & "  Code Running: " & m_CodeIsRunning & vbCrLf
    snapshot = snapshot & vbCrLf
    snapshot = snapshot & "Settings:" & vbCrLf
    snapshot = snapshot & "  Use Bookmarks: " & m_UseBookmarks & vbCrLf
    snapshot = snapshot & "  Coloring: " & m_Coloring & vbCrLf
    snapshot = snapshot & "  Indenting: " & m_Indenting & vbCrLf
    snapshot = snapshot & "  Export: " & m_Export & vbCrLf
    snapshot = snapshot & "  Bold Too: " & m_BoldToo & vbCrLf
    snapshot = snapshot & "  Output Dir: " & m_OutputDir & vbCrLf
    snapshot = snapshot & "  Cancelled: " & m_Cancelled & vbCrLf
    snapshot = snapshot & vbCrLf
    snapshot = snapshot & "Dictionary:" & vbCrLf
    snapshot = snapshot & "  Path: " & m_DocentDictionaryPath & vbCrLf
    snapshot = snapshot & "  Object: " & IIf(m_DocentDictionary Is Nothing, "Nothing", "Set") & vbCrLf
    snapshot = snapshot & "=====================" & vbCrLf
    
    GetStateSnapshot = snapshot
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    GetStateSnapshot = "Error getting state snapshot: " & Err.Description
End Function

'=======================================================
' BACKWARD COMPATIBILITY HELPERS
'=======================================================

'=======================================================
' Note: The following public variables maintain backward
' compatibility with existing code. New code should use
' the properties above. These can be removed once all
' code is migrated to use properties.
'=======================================================

' Backward compatibility - gradually remove these
Public Property Get PrjMgr() As Boolean
    PrjMgr = m_PrjMgr
End Property

Public Property Let PrjMgr(ByVal value As Boolean)
    m_PrjMgr = value
End Property

Public Property Get BusyRibbon() As Boolean
    BusyRibbon = m_BusyRibbon
End Property

Public Property Let BusyRibbon(ByVal value As Boolean)
    m_BusyRibbon = value
End Property

Public Property Get CodeIsRunning() As Boolean
    CodeIsRunning = m_CodeIsRunning
End Property

Public Property Let CodeIsRunning(ByVal value As Boolean)
    m_CodeIsRunning = value
End Property

' Settings backward compatibility
Public Property Get Set_UseBookmarks() As Boolean
    Set_UseBookmarks = m_UseBookmarks
End Property

Public Property Let Set_UseBookmarks(ByVal value As Boolean)
    m_UseBookmarks = value
End Property

Public Property Get Set_Coloring() As Boolean
    Set_Coloring = m_Coloring
End Property

Public Property Let Set_Coloring(ByVal value As Boolean)
    m_Coloring = value
End Property

Public Property Get Set_Indenting() As Boolean
    Set_Indenting = m_Indenting
End Property

Public Property Let Set_Indenting(ByVal value As Boolean)
    m_Indenting = value
End Property

Public Property Get Set_Export() As Boolean
    Set_Export = m_Export
End Property

Public Property Let Set_Export(ByVal value As Boolean)
    m_Export = value
End Property

Public Property Get Set_BoldToo() As Boolean
    Set_BoldToo = m_BoldToo
End Property

Public Property Let Set_BoldToo(ByVal value As Boolean)
    m_BoldToo = value
End Property

Public Property Get Set_OutputDir() As String
    Set_OutputDir = m_OutputDir
End Property

Public Property Let Set_OutputDir(ByVal value As String)
    m_OutputDir = value
End Property

Public Property Get Set_Cancelled() As Boolean
    Set_Cancelled = m_Cancelled
End Property

Public Property Let Set_Cancelled(ByVal value As Boolean)
    m_Cancelled = value
End Property
