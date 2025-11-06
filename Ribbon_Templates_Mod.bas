Attribute VB_Name = "Ribbon_Templates_Mod"
Option Explicit
Option Private Module

'=======================================================
' Module: Ribbon_Templates_Mod
' Purpose: Ribbon callbacks for Templates group
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Handles all ribbon callbacks for the Templates group,
'   including template selection dropdown, open, modify,
'   show/hide hidden text, and cancel buttons.
'
' Dependencies:
'   - Ribbon_Functions_Mod (GetButtonVisible, GetVisibleGroup)
'   - Ribbon_Functions_Mod (GetTemplatesCount, templateName, GetSelectedTemplateIndex, SetSelectedTemplateIndex)
'   - Ribbon_Functions_Mod (IsProjectSelected, OpenSelectedTemplate, UploadDoc, GetInitalState, CancelEditingDoc)
'   - AB_GlobalVars (TemplateNum)
'   - frmTemplatesManager
'
' Ribbon Callbacks:
'   - Template manager button
'   - Template dropdown (selection, display)
'   - Open template button
'   - Modify template button
'   - Show/hide hidden text toggle
'   - Cancel editing button
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added logging
'       * Added function documentation
'       * Improved null/error checks
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Ribbon_Templates_Mod"

' Module-level state
Private ShowHidden As Boolean

'=======================================================
' TEMPLATE MANAGER BUTTON
'=======================================================

'=======================================================
' Sub: IdButtonTemplateOnAction
' Purpose: Handle template manager button click
'
' Description:
'   Opens the templates manager form
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Sub IdButtonTemplateOnAction()
    Const PROC_NAME As String = "IdButtonTemplateOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Template manager button clicked"
    frmTemplatesManager.Show
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to open templates manager: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonTemplateGetVisible
' Purpose: Determine if template manager button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if button should be visible
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdButtonTemplateGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonTemplateGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonTemplateGetVisible = GetButtonVisible(1)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTemplateGetVisible = False
End Function

'=======================================================
' TEMPLATES GROUP VISIBILITY
'=======================================================

'=======================================================
' Function: IdGroupTemplateGetVisible
' Purpose: Determine if templates group should be visible
'
' Parameters:
'   ID - Ribbon group ID
'
' Returns:
'   Boolean - True if group should be visible
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdGroupTemplateGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdGroupTemplateGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdGroupTemplateGetVisible = GetVisibleGroup(ID)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdGroupTemplateGetVisible = False
End Function

'=======================================================
' TEMPLATE DROPDOWN CALLBACKS
'=======================================================

'=======================================================
' Function: IdDDTemplateGetVisible
' Purpose: Determine if template dropdown is visible
'
' Parameters:
'   ID - Ribbon dropdown ID
'
' Returns:
'   Boolean - True if dropdown should be visible
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdDDTemplateGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdDDTemplateGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdDDTemplateGetVisible = GetButtonVisible(1)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdDDTemplateGetVisible = False
End Function

'=======================================================
' Function: IdDDTemplateGetImage
' Purpose: Get image for template dropdown
'
' Returns:
'   IPictureDisp - Image object (currently returns Nothing)
'
' Error Handling:
'   - Returns Nothing on error
'   - Logs errors
'=======================================================
Function IdDDTemplateGetImage() As IPictureDisp
    Const PROC_NAME As String = "IdDDTemplateGetImage"
    
    On Error GoTo ErrorHandler
    
    Set IdDDTemplateGetImage = Nothing
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    Set IdDDTemplateGetImage = Nothing
End Function

'=======================================================
' Function: IdDDTemplateGetItemCount
' Purpose: Get number of templates in dropdown
'
' Returns:
'   Long - Number of template items
'
' Error Handling:
'   - Returns 0 on error
'   - Logs errors
'=======================================================
Function IdDDTemplateGetItemCount() As Long
    Const PROC_NAME As String = "IdDDTemplateGetItemCount"
    
    On Error GoTo ErrorHandler
    
    IdDDTemplateGetItemCount = GetTemplatesCount()
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdDDTemplateGetItemCount = 0
End Function

'=======================================================
' Function: IdDDTemplateGetItemLabel
' Purpose: Get label for template dropdown item
'
' Parameters:
'   Index - Item index (1-based)
'
' Returns:
'   String - Template name
'
' Error Handling:
'   - Returns empty string on error
'   - Logs errors
'=======================================================
Function IdDDTemplateGetItemLabel(ByVal Index As Integer) As String
    Const PROC_NAME As String = "IdDDTemplateGetItemLabel"
    
    On Error GoTo ErrorHandler
    
    IdDDTemplateGetItemLabel = templateName(Index)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error for index " & Index & ": " & Err.Description
    IdDDTemplateGetItemLabel = ""
End Function

'=======================================================
' Function: IdDDTemplateGetSelectedItemIndex
' Purpose: Get currently selected template index
'
' Returns:
'   Integer - Selected template index (1-based)
'
' Error Handling:
'   - Returns 0 on error
'   - Logs errors
'=======================================================
Function IdDDTemplateGetSelectedItemIndex() As Integer
    Const PROC_NAME As String = "IdDDTemplateGetSelectedItemIndex"
    
    On Error GoTo ErrorHandler
    
    IdDDTemplateGetSelectedItemIndex = GetSelectedTemplateIndex()
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdDDTemplateGetSelectedItemIndex = 0
End Function

'=======================================================
' Sub: IdDDTemplateGetItemImage
' Purpose: Get image for template dropdown item
'
' Parameters:
'   Index - Item index (1-based)
'
' Description:
'   Currently not implemented (no images for template items)
'
' Error Handling:
'   - Logs errors if any occur
'=======================================================
Sub IdDDTemplateGetItemImage(ByVal Index As Integer)
    Const PROC_NAME As String = "IdDDTemplateGetItemImage"
    
    On Error GoTo ErrorHandler
    
    ' Not implemented - no images for template items
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error for index " & Index & ": " & Err.Description
End Sub

'=======================================================
' Sub: IdDDTemplateOnAction
' Purpose: Handle template dropdown selection change
'
' Parameters:
'   Index - Selected item index (1-based)
'
' Description:
'   Updates the selected template when user chooses from dropdown
'
' Error Handling:
'   - Logs errors
'   - Continues execution (non-critical)
'=======================================================
Sub IdDDTemplateOnAction(ByVal Index As Integer)
    Const PROC_NAME As String = "IdDDTemplateOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Template selected: " & Index
    SetSelectedTemplateIndex Index
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error setting template " & Index & ": " & Err.Description
End Sub

'=======================================================
' OPEN TEMPLATE BUTTON
'=======================================================

'=======================================================
' Function: IdButtonTemplateOpenGetVisible
' Purpose: Determine if open template button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if button should be visible
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdButtonTemplateOpenGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonTemplateOpenGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonTemplateOpenGetVisible = GetButtonVisible(1)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTemplateOpenGetVisible = False
End Function

'=======================================================
' Sub: IdButtonTemplateOpenOnAction
' Purpose: Handle open template button click
'
' Description:
'   Opens the currently selected template
'
' Error Handling:
'   - Validates project is selected
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Sub IdButtonTemplateOpenOnAction()
    Const PROC_NAME As String = "IdButtonTemplateOpenOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Open template button clicked"
    
    If Not IsProjectSelected() Then
        WriteLog 2, CurrentMod, PROC_NAME, "No project selected"
        MsgBox "Please select a project first.", vbExclamation, "No Project Selected"
        Exit Sub
    End If
    
    Call OpenSelectedTemplate
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to open template: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' Function: IdButtonTemplateOpenGetEnabled
' Purpose: Determine if open template button is enabled
'
' Returns:
'   Boolean - True if a template is selected
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdButtonTemplateOpenGetEnabled() As Boolean
    Const PROC_NAME As String = "IdButtonTemplateOpenGetEnabled"
    
    On Error GoTo ErrorHandler
    
    IdButtonTemplateOpenGetEnabled = (TemplateNum > 0)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTemplateOpenGetEnabled = False
End Function

'=======================================================
' MODIFY TEMPLATE BUTTON
'=======================================================

'=======================================================
' Function: IdButtonTemplateModifyGetVisible
' Purpose: Determine if modify template button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if button should be visible (template context)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdButtonTemplateModifyGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonTemplateModifyGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonTemplateModifyGetVisible = GetButtonVisible(4)  ' 4 = Template mode
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTemplateModifyGetVisible = False
End Function

'=======================================================
' Sub: IdButtonTemplateModifyOnAction
' Purpose: Handle modify template button click
'
' Description:
'   Uploads modified template to server
'
' Error Handling:
'   - Validates active document exists
'   - Validates project is selected
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Sub IdButtonTemplateModifyOnAction()
    Const PROC_NAME As String = "IdButtonTemplateModifyOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Modify template button clicked"
    
    ' Validate active document
    If ActiveDocument Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "No active document"
        MsgBox "No document is currently open.", vbExclamation, "No Document"
        Exit Sub
    End If
    
    If ActiveDocument.Name = "" Then
        WriteLog 2, CurrentMod, PROC_NAME, "Active document has no name"
        Exit Sub
    End If
    
    ' Validate project selection
    If Not IsProjectSelected() Then
        WriteLog 2, CurrentMod, PROC_NAME, "No project selected"
        MsgBox "Please select a project first.", vbExclamation, "No Project Selected"
        Exit Sub
    End If
    
    ' Upload document
    Call UploadDoc(ActiveDocument, GetInitalState(), True)
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to modify template: " & Err.Description, vbExclamation, "Error"
End Sub

'=======================================================
' SHOW/HIDE HIDDEN TEXT TOGGLE
'=======================================================

'=======================================================
' Function: IdToggleButtonTemplateHideGetVisible
' Purpose: Determine if show/hide toggle is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if button should be visible (template context)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdToggleButtonTemplateHideGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdToggleButtonTemplateHideGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdToggleButtonTemplateHideGetVisible = GetButtonVisible(4)  ' 4 = Template mode
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdToggleButtonTemplateHideGetVisible = False
End Function

'=======================================================
' Sub: IdToggleButtonTemplateHideOnAction
' Purpose: Handle show/hide hidden text toggle
'
' Parameters:
'   Pressed - True if button is pressed (show hidden text)
'
' Description:
'   Toggles visibility of hidden text in active document
'
' Error Handling:
'   - Validates active window exists
'   - Logs errors
'   - Continues execution (non-critical)
'=======================================================
Sub IdToggleButtonTemplateHideOnAction(ByVal Pressed As Boolean)
    Const PROC_NAME As String = "IdToggleButtonTemplateHideOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Show hidden text: " & Pressed
    
    If ActiveWindow Is Nothing Then
        WriteLog 2, CurrentMod, PROC_NAME, "No active window"
        Exit Sub
    End If
    
    ActiveWindow.View.ShowHiddenText = Pressed
    ShowHidden = Pressed
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
End Sub

'=======================================================
' Function: IdToggleButtonTemplateHideGetPressed
' Purpose: Get pressed state of show/hide toggle
'
' Returns:
'   Boolean - True if hidden text is currently shown
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdToggleButtonTemplateHideGetPressed() As Boolean
    Const PROC_NAME As String = "IdToggleButtonTemplateHideGetPressed"
    
    On Error GoTo ErrorHandler
    
    IdToggleButtonTemplateHideGetPressed = ShowHidden
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdToggleButtonTemplateHideGetPressed = False
End Function

'=======================================================
' Function: IdToggleButtonTemplateHideGetEnabled
' Purpose: Determine if show/hide toggle is enabled
'
' Returns:
'   Boolean - True if document contains hidden text
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdToggleButtonTemplateHideGetEnabled() As Boolean
    Const PROC_NAME As String = "IdToggleButtonTemplateHideGetEnabled"
    
    On Error Resume Next
    
    ' Check if active document has hidden text
    If ActiveDocument Is Nothing Then
        IdToggleButtonTemplateHideGetEnabled = False
        Exit Function
    End If
    
    IdToggleButtonTemplateHideGetEnabled = (ActiveDocument.Range.Font.Hidden <> 0)
    
    If Err.Number <> 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Error checking hidden text: " & Err.Description
        IdToggleButtonTemplateHideGetEnabled = False
    End If
End Function

'=======================================================
' CANCEL EDITING BUTTON
'=======================================================

'=======================================================
' Function: IdButtonTemplateCancelGetVisible
' Purpose: Determine if cancel button is visible
'
' Parameters:
'   ID - Ribbon button ID
'
' Returns:
'   Boolean - True if button should be visible (template context)
'
' Error Handling:
'   - Returns False on error
'   - Logs errors
'=======================================================
Function IdButtonTemplateCancelGetVisible(ByVal ID As String) As Boolean
    Const PROC_NAME As String = "IdButtonTemplateCancelGetVisible"
    
    On Error GoTo ErrorHandler
    
    IdButtonTemplateCancelGetVisible = GetButtonVisible(4)  ' 4 = Template mode
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    IdButtonTemplateCancelGetVisible = False
End Function

'=======================================================
' Sub: IdButtonTemplateCancelOnAction
' Purpose: Handle cancel editing button click
'
' Description:
'   Cancels template editing and closes document without saving
'
' Error Handling:
'   - Logs errors
'   - Displays user-friendly error message
'=======================================================
Sub IdButtonTemplateCancelOnAction()
    Const PROC_NAME As String = "IdButtonTemplateCancelOnAction"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Cancel editing button clicked"
    Call CancelEditingDoc
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, "Error: " & Err.Description
    MsgBox "Failed to cancel editing: " & Err.Description, vbExclamation, "Error"
End Sub
