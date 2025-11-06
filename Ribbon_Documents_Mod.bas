Attribute VB_Name = "Ribbon_Documents_Mod"
Option Explicit
Option Private Module
Const CurrentMod = "AC_Ribbon_Documents_Mod"
Private ShowHidden As Boolean
Function IdButtonDocumentIconHolderGetVisible(ID As String): IdButtonDocumentIconHolderGetVisible = GetButtonVisible(1): End Function
'Documents Group
Function IdGroupDocumentGetVisible(ID As String): IdGroupDocumentGetVisible = GetVisibleGroup(ID): End Function
'Document Type Drop Down
Function IdDDDocumentGetVisible(ID As String): IdDDDocumentGetVisible = GetButtonVisible(1): End Function
Function IdDDDocumentGetItemCount(): IdDDDocumentGetItemCount = GetDocumentsCount: End Function
Function IdDDDocumentGetItemLabel(Index As Integer): IdDDDocumentGetItemLabel = documentName(Index): End Function
Function IdDDDocumentGetSelectedItemIndex(): IdDDDocumentGetSelectedItemIndex = GetSelectedDocumentIndex: End Function
Sub IdDDDocumentGetItemImage(Index As Integer): End Sub
Function IdDDDocumentGetImage(): Set IdDDDocumentGetImage = Nothing: End Function
Sub IdDDDocumentOnAction(Index As Integer): SetSelectedDocumentIndex Index: End Sub
'Create Document Button
Function IdButtonDocumentCreateGetVisible(ID As String): IdButtonDocumentCreateGetVisible = GetButtonVisible(1): End Function
Sub IdButtonDocumentCreateOnAction()
    On Error Resume Next
    If IsProjectSelected Then OpenTemplate documentName(DocNum) 'OpenDocuments DocumentName(DocNum), True '
End Sub
Function IdButtonDocumentCreateGetEnabled(): IdButtonDocumentCreateGetEnabled = DocNum > 0: End Function
Function IdButtonDocumentCreateGetSupertip()
    If DocNum > 0 Then
        IdButtonDocumentCreateGetSupertip = "Create a new template-based instance of " & documentName(DocNum) '"Create Document From Template."
    Else
        IdButtonDocumentCreateGetSupertip = "A Document Type must be selected."
    End If
End Function
'Open Document
Function IdButtonDocumentOpenGetVisible(ID As String): IdButtonDocumentOpenGetVisible = GetButtonVisible(1): End Function
Sub IdButtonDocumentOpenOnAction()
    On Error Resume Next
    If IsProjectSelected Then OpenDocuments documentName(DocNum), False 'OpenTemplate DocumentName(DocNum) 'OpenDocuments DocumentName(DocNum), False
    Invalidate "IdButtonDocumentState"
End Sub
Function IdButtonDocumentOpenGetEnabled(): IdButtonDocumentOpenGetEnabled = DocNum > 0: End Function
Function IdButtonDocumentOpenGetSupertip()
    If DocNum > 0 Then
        IdButtonDocumentOpenGetSupertip = "Open an existing " & documentName(DocNum) & " for viewing or editing"
    Else
        IdButtonDocumentOpenGetSupertip = "A Document Type must be selected."
    End If
End Function
'Save Document Button
Function IdButtonDocumentSaveGetVisible(ID As String)
    IdButtonDocumentSaveGetVisible = GetButtonVisible(2)
    If IdButtonDocumentSaveGetVisible Then
        IdButtonDocumentSaveGetVisible = False
        On Error Resume Next
        IdButtonDocumentSaveGetVisible = NextTransitions.Count > 0 'UBound(NextTransitions) > 0
        IdButtonDocumentSaveGetVisible = Not IdButtonDocumentSaveGetVisible
    End If
End Function
Sub IdButtonDocumentSaveOnAction()
    On Error Resume Next
    If ActiveDocument.Name = "" Then Exit Sub
    If IsProjectSelected Then UploadDoc ActiveDocument
    Invalidate "IdButtonDocumentState"
End Sub
Sub IdButtonDocumentSave0OnAction(): IdButtonDocumentSaveOnAction: End Sub
'Save Document Split-Button
Function IdSplitButtonDocumentSaveGetVisible(ID As String)
    On Error Resume Next
    IdSplitButtonDocumentSaveGetVisible = False
    IdSplitButtonDocumentSaveGetVisible = GetButtonVisible(2) And NextTransitions.Count > 0 'UBound(NextTransitions) > 0 And
End Function
Sub IdButtonDocumentSaveDefaultOnAction()
    ''Debug.Print "IdButtonDocumentSaveDefaultOnAction"
    WriteLog 1, CurrentMod, "IdButtonDocumentSaveDefaultOnAction", "Save Document Button Clicked"
    On Error Resume Next
    If ActiveDocument.Name = "" Then Exit Sub
    If IsProjectSelected Then UploadDoc ActiveDocument
    Invalidate "IdButtonDocumentState"
End Sub
'Save Document Menu-Button1
Function IdButtonDocumentSaveAs1GetVisible(ID As String): IdButtonDocumentSaveAs1GetVisible = GetVisibleButtonDocumentSaveAs(1): End Function
Function IdButtonDocumentSaveAs1GetLabel(): IdButtonDocumentSaveAs1GetLabel = GetLabelButtonDocumentSaveAs(1): End Function
Sub IdButtonDocumentSaveAs1OnAction(): OnActionButtonDocumentSaveAs 1: End Sub
Function IdButtonDocumentSaveAs1GetImage(): IdButtonDocumentSaveAs1GetImage = GetImageButtonDocumentSaveAs(1): End Function
Sub IdButtonDocumentSaveAs1GetScreentip(): End Sub
Sub IdButtonDocumentSaveAs1GetSupertip(): End Sub
'Save Document Menu-Button2
Function IdButtonDocumentSaveAs2GetVisible(ID As String): IdButtonDocumentSaveAs2GetVisible = GetVisibleButtonDocumentSaveAs(2): End Function
Function IdButtonDocumentSaveAs2GetLabel(): IdButtonDocumentSaveAs2GetLabel = GetLabelButtonDocumentSaveAs(2): End Function
Sub IdButtonDocumentSaveAs2OnAction(): OnActionButtonDocumentSaveAs 2: End Sub
Function IdButtonDocumentSaveAs2GetImage(): IdButtonDocumentSaveAs2GetImage = GetImageButtonDocumentSaveAs(2): End Function
Sub IdButtonDocumentSaveAs2GetScreentip(): End Sub
Sub IdButtonDocumentSaveAs2GetSupertip(): End Sub
'Save Document Menu-Button3
Function IdButtonDocumentSaveAs3GetVisible(ID As String): IdButtonDocumentSaveAs3GetVisible = GetVisibleButtonDocumentSaveAs(3): End Function
Function IdButtonDocumentSaveAs3GetLabel(): IdButtonDocumentSaveAs3GetLabel = GetLabelButtonDocumentSaveAs(3): End Function
Sub IdButtonDocumentSaveAs3OnAction(): OnActionButtonDocumentSaveAs 3: End Sub
Function IdButtonDocumentSaveAs3GetImage(): IdButtonDocumentSaveAs3GetImage = GetImageButtonDocumentSaveAs(3): End Function
Sub IdButtonDocumentSaveAs3GetScreentip(): End Sub
Sub IdButtonDocumentSaveAs3GetSupertip(): End Sub
'Save Buttons
Function GetVisibleButtonDocumentSaveAs(i As Long) As Boolean: On Error Resume Next: GetVisibleButtonDocumentSaveAs = NextTransitions.Count >= i: End Function
Function GetLabelButtonDocumentSaveAs(i As Long) As String: On Error Resume Next: GetLabelButtonDocumentSaveAs = NextTransitions(i)("title"): End Function
Sub OnActionButtonDocumentSaveAs(i As Long)
    On Error Resume Next
    WriteLog 1, CurrentMod, "IdButtonDocumentSaveAs" & i & "OnAction", NextTransitions(i)("title") & " Button Clicked"
    ApplyTransitionNo i, ActiveDocument
    Invalidate "IdButtonDocumentState"
End Sub
Function GetImageButtonDocumentSaveAs(i As Long) ' As String
    On Error GoTo ex
    If InStr(NextTransitions(i)("title"), "Publish") Then GetImageButtonDocumentSaveAs = "BlogPublish"
    If InStr(NextTransitions(i)("title"), "Draft") Then GetImageButtonDocumentSaveAs = "BlogPublishDraft"
    If InStr(NextTransitions(i)("title"), "Retract") Then GetImageButtonDocumentSaveAs = "AdpDiagramRecalculatePageBreaks"
    If InStr(NextTransitions(i)("title"), "Review") Then GetImageButtonDocumentSaveAs = "Magnifier"
    If InStr(NextTransitions(i)("title"), "Archive") Then GetImageButtonDocumentSaveAs = "SheetProtect"
ex:
End Function
'Close Button
Function IdButtonDocumentCancelGetVisible(ID As String): IdButtonDocumentCancelGetVisible = GetButtonVisible(2): End Function
Sub IdButtonDocumentCancelOnAction(): CancelEditingDoc: End Sub
'State Button
Function IdButtonDocumentStateGetEnabled(): IdButtonDocumentStateGetEnabled = IdButtonDocumentStateGetVisible(""): End Function
Function IdButtonDocumentStateGetVisible(ID As String): IdButtonDocumentStateGetVisible = GetButtonVisible(2): End Function
Function IdButtonDocumentStateGetLabel()
    On Error Resume Next
    IdButtonDocumentStateGetLabel = GetProperty(pDocState) & " Document State"
End Function
Sub IdButtonDocumentStateOnAction(): frmStatesList.Show: End Sub
Function IdButtonDocumentStateGetImage()
    IdButtonDocumentStateGetImage = GetStateIcon(GetProperty(pDocState, ActiveDocument))
End Function
Function IdButtonDocumentStateGetSupertip()
    On Error Resume Next
    IdButtonDocumentStateGetSupertip = GetProperty(pDocState)
End Function
'Show/Hide Button
Function IdToggleButtonDocumentHideGetVisible(ID As String): IdToggleButtonDocumentHideGetVisible = GetButtonVisible(2): End Function
Sub IdToggleButtonDocumentHideOnAction(Pressed As Boolean)
    ActiveWindow.View.ShowHiddenText = Pressed
    ShowHidden = Pressed
End Sub
Function IdToggleButtonDocumentHideGetPressed(): IdToggleButtonDocumentHideGetPressed = ShowHidden: End Function
Function IdToggleButtonDocumentHideGetEnabled(): On Error Resume Next: IdToggleButtonDocumentHideGetEnabled = ActiveDocument.Range.Font.Hidden <> 0: End Function
'Command Words
Function IdGroupCommandStatementsGetVisible(ID As String)
    IdGroupCommandStatementsGetVisible = GetVisibleGroup(ID)
End Function
Sub IdButtonCommandStatementsOnAction()
    On Error Resume Next
    If IsProjectSelected Then frmSettings.ShallWillMode
End Sub
Function IdButtonCommandStatementsGetEnabled(): IdButtonCommandStatementsGetEnabled = Documents.Count > 0: End Function
Function IdButtonCommandStatementsGetSupertip()
    If Documents.Count > 0 Then
        IdButtonCommandStatementsGetSupertip = "Collects Command statements into an excel sheet."
    Else
        IdButtonCommandStatementsGetSupertip = "A Document must be open."
    End If
End Function


