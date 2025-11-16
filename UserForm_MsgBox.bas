Attribute VB_Name = "UserForm_MsgBox"
Option Explicit

'=======================================================
' Module: UserForm_MsgBox
' Purpose: Custom message box style definitions
' Author: Docent IMS Team
' Version: 2.0
'
' Description:
'   Defines custom message box styles used throughout
'   the application. These styles control the appearance
'   of custom message box dialogs via frmMsgBox.
'
'   The styles use bit flag values to allow for potential
'   future combinations of multiple styles.
'
' Usage:
'   frmMsgBox.Display "Operation completed", , Success, "Success"
'   frmMsgBox.Display "Warning message", , Exclamation, "Warning"
'   frmMsgBox.Display "Critical error", , Critical, "Error"
'
' Dependencies:
'   - frmMsgBox form
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive module documentation
'       * Added detailed enum documentation
'       * Clarified bit flag usage
'   v1.0 - Original version
'=======================================================

'=======================================================
' Enum: NewMsgBoxStyle
' Purpose: Define visual styles for custom message boxes
'
' Description:
'   Enumeration of icon styles available for custom
'   message boxes. Each value represents a different
'   visual indicator to communicate message type to users.
'
' Values:
'   None        - Default style with no special icon
'   Success     - Green checkmark icon indicating successful operation
'   Exclamation - Yellow warning triangle for attention-needed messages
'   Critical    - Red X icon for critical errors or failures
'   Information - Blue information icon for general notifications
'   Question    - Question mark icon for user decision prompts
'   ZoomIcon    - Zoom meeting icon for Zoom-related operations
'
' Implementation Note:
'   Values use powers of 2 (bit flags) to enable future
'   functionality for combining multiple styles if needed.
'   Current implementation uses single values only.
'
' Examples:
'   ' Success message
'   frmMsgBox.Display "Document saved successfully", , Success, "Save Complete"
'
'   ' Warning message
'   frmMsgBox.Display "Some changes may be lost", , Exclamation, "Warning"
'
'   ' Error message
'   frmMsgBox.Display "Failed to connect to server", , Critical, "Connection Error"
'
'   ' Information message
'   frmMsgBox.Display "Process will take a few minutes", , Information, "Please Wait"
'=======================================================
Public Enum NewMsgBoxStyle
    None = 0            ' Default - no special icon displayed
    Success = 1         ' 2^0 - Green checkmark for success
    Exclamation = 2     ' 2^1 - Yellow warning triangle
    Critical = 4        ' 2^2 - Red X for critical errors
    Information = 8     ' 2^3 - Blue info icon
    Question = 16       ' 2^4 - Question mark for prompts
    ZoomIcon = 32       ' 2^5 - Zoom integration icon
End Enum

'=======================================================
' END OF MODULE
'=======================================================
