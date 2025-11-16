Attribute VB_Name = "Ribbon_Notifications"
Option Explicit
Option Private Module

'=======================================================
' Module: Ribbon_Notifications
' Purpose: Ribbon callbacks for Notifications group
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Handles all ribbon callbacks for the Notifications group,
'   including visibility, images, tooltips, and actions for
'   notification traffic lights (Aqua/Yellow/Red) and
'   create notification button.
'
' Dependencies:
'   - Ribbon_Functions_Mod (GetVisibleGroup, GetNotificationImage, GetNotificationCount, GotoNotificationCollection)
'   - AB_GlobalVars (ProjectColorStr, PrjMgr)
'   - frmCreateNotification
'
' Ribbon Callbacks:
'   - IdGroupNotificationsGetVisible - Group visibility
'   - IdButtonNotifications[Color]GetImage - Traffic light images
'   - IdButtonNotifications[Color]GetVisible - Traffic light visibility
'   - IdButtonNotifications[Color]OnAction - Traffic light actions
'   - IdButtonNotifications[Color]GetSupertip - Traffic light tooltips
'   - IdButtonCreateNotificationOnAction - Create/view notifications action
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added logging
'       * Added function documentation
'       * Removed commented dead code
'       * Improved null/error checks
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Ribbon_Notifications"

'=======================================================
' RIBBON GROUP VISIBILITY
'=======================================================

'=======================================================
' Function: IdGroupNotificationsGetVisible
' Purpose: Determine if Notifications group should be visible
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
Private Function IdGroupNotificationsGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdGroupNotificationsGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdGroupNotificationsGetVisible = GetVisibleGroup(ID)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdGroupNotificationsGetVisible = False
End Function

'=======================================================
' AQUA (INFORMATIONAL) NOTIFICATION CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonNotificationsGreenGetImage
' Purpose: Get image for informational (aqua) notification button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsGreenGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonNotificationsGreenGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonNotificationsGreenGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonNotificationsGreenGetImage = GetNotificationImage("Aqua")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonNotificationsGreenGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonNotificationsGreenGetVisible
' Purpose: Determine if aqua notification button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if visible (not in project manager mode)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsGreenGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonNotificationsGreenGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonNotificationsGreenGetVisible = Not PrjMgr
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsGreenGetVisible = False
End Function

'=======================================================
' Sub: IdButtonNotificationsGreenOnAction
' Purpose: Handle aqua notification button click
'
' Description:
'   Navigates to collection of informational notifications
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonNotificationsGreenOnAction()
    Const PROC_NAME As String = "IdButtonNotificationsGreenOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Aqua notifications button clicked"
    GotoNotificationCollection "Aqua"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display informational notifications: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonNotificationsGreenGetSupertip
' Purpose: Get tooltip for aqua notification button
'
' Returns:
'   String - Tooltip showing count of informational notifications
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsGreenGetSupertip() As String
    Const PROC_NAME As String = "IdButtonNotificationsGreenGetSupertip"
    
    Dim notificationCount As Long
    
    On Error GoTo ErrorHandler
    
    notificationCount = GetNotificationCount("Aqua")
    IdButtonNotificationsGreenGetSupertip = notificationCount & " Informational Notification" & _
                                           IIf(notificationCount <> 1, "s", "")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsGreenGetSupertip = "Informational notifications"
End Function

'=======================================================
' YELLOW (IMPORTANT) NOTIFICATION CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonNotificationsYellowGetImage
' Purpose: Get image for important (yellow) notification button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsYellowGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonNotificationsYellowGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonNotificationsYellowGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonNotificationsYellowGetImage = GetNotificationImage("Yellow")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonNotificationsYellowGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonNotificationsYellowGetVisible
' Purpose: Determine if yellow notification button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if visible (not in project manager mode)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsYellowGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonNotificationsYellowGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonNotificationsYellowGetVisible = Not PrjMgr
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsYellowGetVisible = False
End Function

'=======================================================
' Sub: IdButtonNotificationsYellowOnAction
' Purpose: Handle yellow notification button click
'
' Description:
'   Navigates to collection of important notifications
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonNotificationsYellowOnAction()
    Const PROC_NAME As String = "IdButtonNotificationsYellowOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Yellow notifications button clicked"
    GotoNotificationCollection "Yellow"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display important notifications: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonNotificationsYellowGetSupertip
' Purpose: Get tooltip for yellow notification button
'
' Returns:
'   String - Tooltip showing count of important notifications
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsYellowGetSupertip() As String
    Const PROC_NAME As String = "IdButtonNotificationsYellowGetSupertip"
    
    Dim notificationCount As Long
    
    On Error GoTo ErrorHandler
    
    notificationCount = GetNotificationCount("Yellow")
    IdButtonNotificationsYellowGetSupertip = notificationCount & " Important Notification" & _
                                             IIf(notificationCount <> 1, "s", "")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsYellowGetSupertip = "Important notifications"
End Function

'=======================================================
' RED (CRITICAL) NOTIFICATION CALLBACKS
'=======================================================

'=======================================================
' Function: IdButtonNotificationsRedGetImage
' Purpose: Get image for critical (red) notification button
'
' Returns:
'   IPictureDisp - Image object, or Nothing on error
'
' Error Handling:
'   - Returns Nothing if project color not set
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsRedGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdButtonNotificationsRedGetImage"
    
    On Error GoTo ErrorHandler
    
    If Len(ProjectColorStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Project color not set"
        Set IdButtonNotificationsRedGetImage = Nothing
        Exit Function
    End If
    
    Set IdButtonNotificationsRedGetImage = GetNotificationImage("Red")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdButtonNotificationsRedGetImage = Nothing
End Function

'=======================================================
' Function: IdButtonNotificationsRedGetVisible
' Purpose: Determine if red notification button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if visible (not in project manager mode)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsRedGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonNotificationsRedGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonNotificationsRedGetVisible = Not PrjMgr
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsRedGetVisible = False
End Function

'=======================================================
' Sub: IdButtonNotificationsRedOnAction
' Purpose: Handle red notification button click
'
' Description:
'   Navigates to collection of critical notifications
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonNotificationsRedOnAction()
    Const PROC_NAME As String = "IdButtonNotificationsRedOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Red notifications button clicked"
    GotoNotificationCollection "Red"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to display critical notifications: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonNotificationsRedGetSupertip
' Purpose: Get tooltip for red notification button
'
' Returns:
'   String - Tooltip showing count of critical notifications
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonNotificationsRedGetSupertip() As String
    Const PROC_NAME As String = "IdButtonNotificationsRedGetSupertip"
    
    Dim notificationCount As Long
    
    On Error GoTo ErrorHandler
    
    notificationCount = GetNotificationCount("Red")
    IdButtonNotificationsRedGetSupertip = notificationCount & " Critical Notification" & _
                                          IIf(notificationCount <> 1, "s", "")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonNotificationsRedGetSupertip = "Critical notifications"
End Function

'=======================================================
' CREATE/VIEW NOTIFICATIONS BUTTON
'=======================================================

'=======================================================
' Sub: IdButtonCreateNotificationOnAction
' Purpose: Handle create/view notifications button click
'
' Description:
'   If in project manager mode, opens notification creation form.
'   If in team member mode, navigates to all notifications.
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Private Sub IdButtonCreateNotificationOnAction()
    Const PROC_NAME As String = "IdButtonCreateNotificationOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Create/view notifications button clicked (PrjMgr=" & PrjMgr & ")"
    
    If PrjMgr Then
        ' Project manager - show creation form
        frmCreateNotification.Show
    Else
        ' Team member - view all notifications
        GotoNotificationCollection "All"
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    
    Dim action As String
    action = IIf(PrjMgr, "create notification form", "notifications")
    MsgBox "Failed to open " & action & ": " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonCreateNotificationGetVisible
' Purpose: Determine if create/view notification button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - Always True (button always visible)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Private Function IdButtonCreateNotificationGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonCreateNotificationGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonCreateNotificationGetVisible = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonCreateNotificationGetVisible = False
End Function

'=======================================================
' Function: IdButtonCreateNotificationGetLabel
' Purpose: Get label for create/view notification button
'
' Returns:
'   String - "Create Notif." for PM, "Notifications" for team member
'
' Error Handling:
'   - Returns generic label on error
'   - Logs errors
'=======================================================
Private Function IdButtonCreateNotificationGetLabel() As String
    Const PROC_NAME As String = "IdButtonCreateNotificationGetLabel"
    
    On Error GoTo ErrorHandler
    
    IdButtonCreateNotificationGetLabel = IIf(PrjMgr, "Create Notif.", "Notifications")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonCreateNotificationGetLabel = "Notifications"
End Function

'=======================================================
' Function: IdButtonCreateNotificationGetSupertip
' Purpose: Get detailed tooltip for create/view notification button
'
' Returns:
'   String - Detailed tooltip based on user role
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonCreateNotificationGetSupertip() As String
    Const PROC_NAME As String = "IdButtonCreateNotificationGetSupertip"
    
    Dim tooltip As String
    
    On Error GoTo ErrorHandler
    
    If PrjMgr Then
        ' Project manager tooltip
        tooltip = "Avoid email and send a notification to any/all team members." & vbNewLine & vbNewLine & _
                 "Notifications are displayed as a horizontal bar on the project website"
    Else
        ' Team member tooltip
        tooltip = "Click to go to your notifications on Plone"
    End If
    
    IdButtonCreateNotificationGetSupertip = tooltip
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonCreateNotificationGetSupertip = "Notifications"
End Function

'=======================================================
' Function: IdButtonCreateNotificationGetScreentip
' Purpose: Get short tooltip for create/view notification button
'
' Returns:
'   String - Short tooltip based on user role
'
' Error Handling:
'   - Returns generic tooltip on error
'   - Logs errors
'=======================================================
Private Function IdButtonCreateNotificationGetScreentip() As String
    Const PROC_NAME As String = "IdButtonCreateNotificationGetScreentip"
    
    On Error GoTo ErrorHandler
    
    IdButtonCreateNotificationGetScreentip = IIf(PrjMgr, "Create Notification", "Notifications")
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonCreateNotificationGetScreentip = "Notifications"
End Function
