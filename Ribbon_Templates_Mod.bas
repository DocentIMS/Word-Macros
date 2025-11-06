Attribute VB_Name = "Ribbon_Templates_Mod"
Option Explicit
Private ShowHidden As Boolean

Sub IdButtonTemplateOnAction(): frmTemplatesManager.Show: End Sub
Function IdButtonTemplateGetVisible(ID As String): IdButtonTemplateGetVisible = GetButtonVisible(1): End Function

'Templates Manager
Function IdGroupTemplateGetVisible(ID As String): IdGroupTemplateGetVisible = GetVisibleGroup(ID): End Function
Function IdDDTemplateGetVisible(ID As String): IdDDTemplateGetVisible = GetButtonVisible(1): End Function
Function IdDDTemplateGetImage(): Set IdDDTemplateGetImage = Nothing: End Function
Function IdDDTemplateGetItemCount(): IdDDTemplateGetItemCount = GetTemplatesCount: End Function
Function IdDDTemplateGetItemLabel(Index As Integer): IdDDTemplateGetItemLabel = templateName(Index): End Function
Function IdDDTemplateGetSelectedItemIndex(): IdDDTemplateGetSelectedItemIndex = GetSelectedTemplateIndex: End Function
Sub IdDDTemplateGetItemImage(Index As Integer): End Sub
Sub IdDDTemplateOnAction(Index As Integer): SetSelectedTemplateIndex Index: End Sub
'Open Template
Function IdButtonTemplateOpenGetVisible(ID As String): IdButtonTemplateOpenGetVisible = GetButtonVisible(1): End Function
Sub IdButtonTemplateOpenOnAction()
    On Error Resume Next
    If IsProjectSelected Then OpenSelectedTemplate
End Sub
Function IdButtonTemplateOpenGetEnabled(): IdButtonTemplateOpenGetEnabled = TemplateNum > 0: End Function
'Modify
Function IdButtonTemplateModifyGetVisible(ID As String): IdButtonTemplateModifyGetVisible = GetButtonVisible(4): End Function
Sub IdButtonTemplateModifyOnAction()
    On Error Resume Next
    If ActiveDocument.Name = "" Then Exit Sub
    If IsProjectSelected Then UploadDoc ActiveDocument, GetInitalState, True
End Sub
'Show/Hide Button
Function IdToggleButtonTemplateHideGetVisible(ID As String): IdToggleButtonTemplateHideGetVisible = GetButtonVisible(4): End Function
Sub IdToggleButtonTemplateHideOnAction(Pressed As Boolean)
    ActiveWindow.View.ShowHiddenText = Pressed
    ShowHidden = Pressed
End Sub
Function IdToggleButtonTemplateHideGetPressed(): IdToggleButtonTemplateHideGetPressed = ShowHidden: End Function
Function IdToggleButtonTemplateHideGetEnabled(): On Error Resume Next: IdToggleButtonTemplateHideGetEnabled = ActiveDocument.Range.Font.Hidden <> 0: End Function
'Close Button
Function IdButtonTemplateCancelGetVisible(ID As String): IdButtonTemplateCancelGetVisible = GetButtonVisible(4): End Function
Sub IdButtonTemplateCancelOnAction(): CancelEditingDoc: End Sub


