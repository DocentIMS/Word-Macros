Attribute VB_Name = "Ribbon_Manager"
Option Explicit

'=======================================================
' Module: Ribbon_Manager
' Purpose: Manager/Team mode toggle button ribbon handlers
' Author: Docent IMS Team
' Version: 2.0
'
' Description:
'   Implements ribbon callbacks for the Project Manager/Team
'   mode toggle button. This button switches the interface
'   between tools for project managers and tools for team members.
'
'   The mode is stored in the global PrjMgr variable and affects
'   which ribbon groups and controls are visible throughout the
'   application.
'
' Ribbon Control ID:
'   IdToggleButtonMgrMode - Main toggle button control
'
' Global Variables Used:
'   PrjMgr (Boolean) - Current mode state
'       True = Project Manager mode
'       False = Team mode
'
' Dependencies:
'   - AB_GlobalVars (PrjMgr variable)
'   - Ribbon_Functions_Mod (GetVisibleGroup, RefreshRibbon)
'   - Image resources ("PrjMgr", "MeetingsWorkspace")
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive module documentation
'       * Added proper error handling to all functions
'       * Added function documentation headers
'       * Improved code formatting
'       * Added logging support
'   v1.0 - Original version
'=======================================================

' Module constants
Private Const CurrentMod As String = "Ribbon_Manager"

'=======================================================
' Function: IdToggleButtonMgrModeGetVisible
' Purpose: Determine if manager mode toggle button should be visible
'
' Parameters:
'   ID - Control identifier string
'
' Returns:
'   Boolean - True if button should be visible
'
' Description:
'   Delegates to GetVisibleGroup to determine visibility
'   based on current document context and permissions.
'=======================================================
Public Function IdToggleButtonMgrModeGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdToggleButtonMgrModeGetVisible = GetVisibleGroup(ID)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetVisible = False
End Function

'=======================================================
' Sub: IdToggleButtonMgrModeOnAction
' Purpose: Handle toggle button click event
'
' Parameters:
'   Pressed - New pressed state of toggle button
'
' Description:
'   Toggles between Project Manager and Team modes by
'   flipping the PrjMgr global variable and refreshing
'   the ribbon to show/hide appropriate controls.
'=======================================================
Public Sub IdToggleButtonMgrModeOnAction(ByVal Pressed As Boolean)
    Const PROC_NAME As String = "IdToggleButtonMgrModeOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Mode toggle requested. Current mode: " & IIf(PrjMgr, "Manager", "Team")
    
    ' Toggle the mode
    PrjMgr = Not PrjMgr
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Mode changed to: " & IIf(PrjMgr, "Manager", "Team")
    
    ' Refresh ribbon to update control visibility
    RefreshRibbon
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    frmMsgBox.Display "Unable to switch mode." & vbCrLf & vbCrLf & _
                      "Error: " & Err.Description, , Critical, "Mode Switch Error"
End Sub

'=======================================================
' Function: IdToggleButtonMgrModeGetPressed
' Purpose: Get current pressed state of toggle button
'
' Returns:
'   Boolean - Always returns True (button maintains state)
'
' Description:
'   Returns the pressed state for the toggle button.
'   Currently always returns True as the actual state
'   is managed through the PrjMgr variable.
'=======================================================
Public Function IdToggleButtonMgrModeGetPressed() As Boolean
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetPressed"
    
    On Error GoTo ErrorHandler
    
    IdToggleButtonMgrModeGetPressed = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetPressed = False
End Function

'=======================================================
' Function: IdToggleButtonMgrModeGetEnabled
' Purpose: Determine if toggle button should be enabled
'
' Returns:
'   Boolean - True if button should be enabled
'
' Description:
'   Controls whether the user can click the mode toggle.
'   Currently always returns True to allow mode switching
'   at any time.
'=======================================================
Public Function IdToggleButtonMgrModeGetEnabled() As Boolean
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetEnabled"
    
    On Error GoTo ErrorHandler
    
    IdToggleButtonMgrModeGetEnabled = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetEnabled = True  ' Default to enabled on error
End Function

'=======================================================
' Function: IdToggleButtonMgrModeGetImage
' Purpose: Get icon image for current mode
'
' Returns:
'   Variant - Image object or image name string
'
' Description:
'   Returns appropriate icon based on current mode:
'   - Project Manager mode: Custom "PrjMgr" image
'   - Team mode: "MeetingsWorkspace" icon
'=======================================================
Private Function IdToggleButtonMgrModeGetImage() As Variant
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetImage"
    
    On Error GoTo ErrorHandler
    
    If PrjMgr Then
        ' Project Manager mode - return custom image object
        Set IdToggleButtonMgrModeGetImage = GetImage("PrjMgr")
    Else
        ' Team mode - return built-in Office icon name
        IdToggleButtonMgrModeGetImage = "MeetingsWorkspace"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ' Return default icon on error
    IdToggleButtonMgrModeGetImage = "MeetingsWorkspace"
End Function

'=======================================================
' Function: IdToggleButtonMgrModeGetLabel
' Purpose: Get button label text for current mode
'
' Returns:
'   String - Label text ("Prj Mgr Mode" or "Team Mode")
'
' Description:
'   Returns compact label text based on current mode.
'   Uses abbreviated "Prj Mgr" to save ribbon space.
'=======================================================
Private Function IdToggleButtonMgrModeGetLabel() As String
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetLabel"
    
    On Error GoTo ErrorHandler
    
    If PrjMgr Then
        IdToggleButtonMgrModeGetLabel = "Prj Mgr Mode"
    Else
        IdToggleButtonMgrModeGetLabel = "Team Mode"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetLabel = "Mode"
End Function

'=======================================================
' Function: IdToggleButtonMgrModeGetScreenTip
' Purpose: Get tooltip text for button
'
' Returns:
'   String - Short tooltip text
'
' Description:
'   Returns brief tooltip describing the current mode.
'   Appears when user hovers over the button.
'=======================================================
Private Function IdToggleButtonMgrModeGetScreenTip() As String
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetScreenTip"
    
    On Error GoTo ErrorHandler
    
    If PrjMgr Then
        IdToggleButtonMgrModeGetScreenTip = "Project Manager Mode"
    Else
        IdToggleButtonMgrModeGetScreenTip = "Team Mode"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetScreenTip = "Manager/Team Mode"
End Function

'=======================================================
' Function: IdToggleButtonMgrModeGetSuperTip
' Purpose: Get extended tooltip text for button
'
' Returns:
'   String - Detailed tooltip text with instructions
'
' Description:
'   Returns detailed tooltip explaining the current mode
'   and how to switch. Appears in enhanced tooltip popup.
'=======================================================
Private Function IdToggleButtonMgrModeGetSuperTip() As String
    Const PROC_NAME As String = "IdToggleButtonMgrModeGetSuperTip"
    
    On Error GoTo ErrorHandler
    
    If PrjMgr Then
        ' Currently in Project Manager mode
        IdToggleButtonMgrModeGetSuperTip = _
            "Tools used by the project manager at the beginning of each project." & _
            vbLf & vbLf & _
            "Click to switch to Team Mode"
    Else
        ' Currently in Team mode
        IdToggleButtonMgrModeGetSuperTip = _
            "Tools used by all team members throughout the project" & _
            vbLf & vbLf & _
            "Click to switch to Project Manager Mode"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    IdToggleButtonMgrModeGetSuperTip = "Toggle between Manager and Team modes"
End Function

'=======================================================
' END OF MODULE
'=======================================================
