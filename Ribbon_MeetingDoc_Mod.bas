Attribute VB_Name = "Ribbon_MeetingDoc_Mod"
Option Explicit

Option Private Module
Const CurrentMod = "AC_Ribbon_MeetingDoc_Mod"
Private ShowHidden As Boolean
'MeetingDocs Group
Function IdGroupMeetingDocGetVisible(ID As String): IdGroupMeetingDocGetVisible = GetVisibleGroup(ID): End Function
'MeetingDoc Type Drop Down
Function IdDDMeetingDocGetVisible(ID As String): IdDDMeetingDocGetVisible = GetButtonVisible(1): End Function
Function IdDDMeetingDocGetItemCount(): IdDDMeetingDocGetItemCount = GetDocumentsCount(1): End Function
Function IdDDMeetingDocGetItemLabel(Index As Integer): IdDDMeetingDocGetItemLabel = MeetingDocName(Index): End Function
Function IdDDMeetingDocGetSelectedItemIndex(): IdDDMeetingDocGetSelectedItemIndex = GetSelectedMeetingDocIndex: End Function
Sub IdDDMeetingDocGetItemImage(Index As Integer): End Sub
Function IdDDMeetingDocGetImage(): Set IdDDMeetingDocGetImage = Nothing: End Function
Sub IdDDMeetingDocOnAction(Index As Integer): SetSelectedMeetingDocIndex Index: End Sub
'Create MeetingDoc Button
Function IdButtonMeetingDocCreateGetVisible(ID As String): IdButtonMeetingDocCreateGetVisible = GetButtonVisible(1): End Function
Sub IdButtonMeetingDocCreateOnAction()
    On Error Resume Next
    If IsProjectSelected Then OpenDocuments MeetingDocName(MeetingDocNum), True 'OpenTemplate MeetingDocName(MeetingDocNum) '
End Sub
Function IdButtonMeetingDocCreateGetEnabled(): IdButtonMeetingDocCreateGetEnabled = MeetingDocNum > 0: End Function
Function IdButtonMeetingDocCreateGetSupertip()
    If MeetingDocNum > 0 Then
        IdButtonMeetingDocCreateGetSupertip = "Create a new template-based instance of " & MeetingDocName(MeetingDocNum) '"Create Document From Template."
    Else
        IdButtonMeetingDocCreateGetSupertip = "A Document Type must be selected."
    End If
End Function
'Open MeetingDoc
Function IdButtonMeetingDocOpenGetVisible(ID As String): IdButtonMeetingDocOpenGetVisible = GetButtonVisible(1): End Function
Sub IdButtonMeetingDocOpenOnAction()
    On Error Resume Next
    If IsProjectSelected Then OpenDocuments MeetingDocName(MeetingDocNum), False 'OpenTemplate MeetingDocName(MeetingDocNum) '
    Invalidate "IdButtonMeetingDocState"
End Sub
Function IdButtonMeetingDocOpenGetEnabled(): IdButtonMeetingDocOpenGetEnabled = MeetingDocNum > 0: End Function
Function IdButtonMeetingDocOpenGetSupertip()
    If MeetingDocNum > 0 Then
        IdButtonMeetingDocOpenGetSupertip = "Open an existing " & MeetingDocName(MeetingDocNum) & " for viewing or editing"
    Else
        IdButtonMeetingDocOpenGetSupertip = "A Document Type must be selected."
    End If
End Function
'Save MeetingDoc Button
Function IdButtonMeetingDocSaveGetVisible(ID As String)
    IdButtonMeetingDocSaveGetVisible = GetButtonVisible(2)
    If IdButtonMeetingDocSaveGetVisible Then
        IdButtonMeetingDocSaveGetVisible = False
        On Error Resume Next
        IdButtonMeetingDocSaveGetVisible = NextTransitions.Count > 0 'UBound(NextTransitions) > 0
        IdButtonMeetingDocSaveGetVisible = Not IdButtonMeetingDocSaveGetVisible
    End If
End Function
Sub IdButtonMeetingDocSaveOnAction()
    On Error Resume Next
    If ActiveDocument.Name = "" Then Exit Sub
    If IsProjectSelected Then UploadDoc ActiveDocument
    Invalidate "IdButtonMeetingDocState"
End Sub
'Save MeetingDoc Split-Button
Function IdSplitButtonMeetingDocSaveGetVisible(ID As String)
    On Error Resume Next
    IdSplitButtonMeetingDocSaveGetVisible = False
    IdSplitButtonMeetingDocSaveGetVisible = GetButtonVisible(2) And NextTransitions.Count > 0 'UBound(NextTransitions) > 0 And
End Function
Sub IdButtonMeetingDocSaveDefaultOnAction()
    ''Debug.Print "IdButtonMeetingDocSaveDefaultOnAction"
    WriteLog 1, CurrentMod, "IdButtonMeetingDocSaveDefaultOnAction", "Save Document Button Clicked"
    On Error Resume Next
    If ActiveDocument.Name = "" Then Exit Sub
    If IsProjectSelected Then UploadDoc ActiveDocument
    Invalidate "IdButtonMeetingDocState"
End Sub
Sub IdButtonMeetingDocSaveDefault0OnAction(): IdButtonMeetingDocSaveDefaultOnAction: End Sub
'Save MeetingDoc Menu-Button1
Function IdButtonMeetingDocSaveAs1GetVisible(ID As String): IdButtonMeetingDocSaveAs1GetVisible = GetVisibleButtonMeetingDocSaveAs(1): End Function
Function IdButtonMeetingDocSaveAs1GetLabel(): IdButtonMeetingDocSaveAs1GetLabel = GetLabelButtonMeetingDocSaveAs(1): End Function
Sub IdButtonMeetingDocSaveAs1OnAction(): OnActionButtonMeetingDocSaveAs 1: End Sub
Function IdButtonMeetingDocSaveAs1GetImage(): IdButtonMeetingDocSaveAs1GetImage = GetImageButtonMeetingDocSaveAs(1): End Function
Sub IdButtonMeetingDocSaveAs1GetScreentip(): End Sub
Sub IdButtonMeetingDocSaveAs1GetSupertip(): End Sub
'Save MeetingDoc Menu-Button2
Function IdButtonMeetingDocSaveAs2GetVisible(ID As String): IdButtonMeetingDocSaveAs2GetVisible = GetVisibleButtonMeetingDocSaveAs(2): End Function
Function IdButtonMeetingDocSaveAs2GetLabel(): IdButtonMeetingDocSaveAs2GetLabel = GetLabelButtonMeetingDocSaveAs(2): End Function
Sub IdButtonMeetingDocSaveAs2OnAction(): OnActionButtonMeetingDocSaveAs 2: End Sub
Function IdButtonMeetingDocSaveAs2GetImage(): IdButtonMeetingDocSaveAs2GetImage = GetImageButtonMeetingDocSaveAs(2): End Function
Sub IdButtonMeetingDocSaveAs2GetScreentip(): End Sub
Sub IdButtonMeetingDocSaveAs2GetSupertip(): End Sub
'Save MeetingDoc Menu-Button3
Function IdButtonMeetingDocSaveAs3GetVisible(ID As String): IdButtonMeetingDocSaveAs3GetVisible = GetVisibleButtonMeetingDocSaveAs(3): End Function
Function IdButtonMeetingDocSaveAs3GetLabel(): IdButtonMeetingDocSaveAs3GetLabel = GetLabelButtonMeetingDocSaveAs(3): End Function
Sub IdButtonMeetingDocSaveAs3OnAction(): OnActionButtonMeetingDocSaveAs 3: End Sub
Function IdButtonMeetingDocSaveAs3GetImage(): IdButtonMeetingDocSaveAs3GetImage = GetImageButtonMeetingDocSaveAs(3): End Function
Sub IdButtonMeetingDocSaveAs3GetScreentip(): End Sub
Sub IdButtonMeetingDocSaveAs3GetSupertip(): End Sub
'Save Buttons
Function GetVisibleButtonMeetingDocSaveAs(i As Long) As Boolean: On Error Resume Next: GetVisibleButtonMeetingDocSaveAs = NextTransitions.Count >= i: End Function
Function GetLabelButtonMeetingDocSaveAs(i As Long) As String: On Error Resume Next: GetLabelButtonMeetingDocSaveAs = NextTransitions(i)("title"): End Function
Sub OnActionButtonMeetingDocSaveAs(i As Long)
    On Error Resume Next
    WriteLog 1, CurrentMod, "IdButtonMeetingDocSaveAs" & i & "OnAction", NextTransitions(i)("title") & " Button Clicked"
    ApplyTransitionNo i, ActiveDocument
'    If ActiveDocument.Name = "" Then Exit Sub
'    If IsProjectSelected Then UploadDoc ActiveDocument, NextTransitions(i)("@id") 'GetStateFromTrn(NextTransitions(3))
    Invalidate "IdButtonMeetingDocState"
End Sub
Function GetImageButtonMeetingDocSaveAs(i As Long) ' As String
    On Error GoTo ex
    If InStr(NextTransitions(i)("title"), "Publish") Then GetImageButtonMeetingDocSaveAs = "BlogPublish"
    If InStr(NextTransitions(i)("title"), "Draft") Then GetImageButtonMeetingDocSaveAs = "BlogPublishDraft"
    If InStr(NextTransitions(i)("title"), "Retract") Then GetImageButtonMeetingDocSaveAs = "AdpDiagramRecalculatePageBreaks"
    If InStr(NextTransitions(i)("title"), "Review") Then GetImageButtonMeetingDocSaveAs = "Magnifier"
    If InStr(NextTransitions(i)("title"), "Archive") Then GetImageButtonMeetingDocSaveAs = "SheetProtect"
    If InStr(NextTransitions(i)("title"), "Close") Then GetImageButtonMeetingDocSaveAs = "FilePermissionRestrictMenu"
ex:
End Function
'Get Zoom Info
Function IdButtonMeetingDocGetZoomInfoGetVisible(ID As String): IdButtonMeetingDocGetZoomInfoGetVisible = GetButtonVisible(2) And GetProperty(pDocType) = "Meeting Notes": End Function
Sub IdButtonMeetingDocGetZoomInfoOnAction(): FillNotesFromZoom: End Sub
'Close Button
Function IdButtonMeetingDocCancelGetVisible(ID As String): IdButtonMeetingDocCancelGetVisible = GetButtonVisible(2): End Function
Sub IdButtonMeetingDocCancelOnAction(): On Error Resume Next: LockAPIFile GetProperty(pDocURL), True: ActiveDocument.Close False: End Sub
'State Button
Function IdButtonMeetingDocStateGetEnabled(): IdButtonMeetingDocStateGetEnabled = IdButtonMeetingDocStateGetVisible(""): End Function
Function IdButtonMeetingDocStateGetVisible(ID As String): IdButtonMeetingDocStateGetVisible = GetButtonVisible(2): End Function
Function IdButtonMeetingDocStateGetLabel()
    On Error Resume Next
    IdButtonMeetingDocStateGetLabel = GetProperty(pDocState) & " Document State"
End Function
Sub IdButtonMeetingDocStateOnAction(): frmStatesList.Show: End Sub
Function IdButtonMeetingDocStateGetImage()
    IdButtonMeetingDocStateGetImage = GetStateIcon(GetProperty(pDocState, ActiveDocument))
End Function
Function IdButtonMeetingDocStateGetSupertip()
    On Error Resume Next
    IdButtonMeetingDocStateGetSupertip = GetProperty(pDocState)
End Function
'Show/Hide Button
Function IdToggleButtonMeetingDocHideGetVisible(ID As String): IdToggleButtonMeetingDocHideGetVisible = GetButtonVisible(2): End Function
Sub IdToggleButtonMeetingDocHideOnAction(Pressed As Boolean)
    ActiveWindow.View.ShowHiddenText = Pressed
    ShowHidden = Pressed
End Sub
Function IdToggleButtonMeetingDocHideGetPressed(): IdToggleButtonMeetingDocHideGetPressed = ShowHidden: End Function
Function IdToggleButtonMeetingDocHideGetEnabled(): On Error Resume Next: IdToggleButtonMeetingDocHideGetEnabled = ActiveDocument.Range.Font.Hidden <> 0: End Function
'Command Words
Function IdGroupCommandStatementsGetVisible(ID As String)
    IdGroupCommandStatementsGetVisible = GetVisibleGroup(ID)
End Function
Sub IdButtonCommandStatementsOnAction()
    If IsProjectSelected Then frmSettings.ShallWillMode
End Sub
Sub IdButtonCommandStatements0OnAction()
    If IsProjectSelected Then frmSettings.ShallWillMode
End Sub
Sub IdButtonCommandStatementsBrowseOnAction()
'    CodeIsRunning = True
    If Not OpenAsDocentDocument("Command Satements") Is Nothing Then
        If IsProjectSelected Then frmSettings.ShallWillMode
    End If
'    CodeIsRunning = False
End Sub
Function IdButtonCommandStatementsGetEnabled(): IdButtonCommandStatementsGetEnabled = Documents.Count > 0: End Function
Function IdButtonCommandStatementsGetSupertip()
    If Documents.Count > 0 Then
        IdButtonCommandStatementsGetSupertip = "Collects Command statements into an excel sheet."
    Else
        IdButtonCommandStatementsGetSupertip = "A Document must be open."
    End If
End Function
'Create Meeting
Function IdSplitButtonCreateMeetingGetVisible(ID As String): IdSplitButtonCreateMeetingGetVisible = GetButtonVisible(1): End Function
Private Sub IdButtonCreateMeetingOnAction()
    If IsProjectSelected Then frmCreateMeeting.Show
End Sub
Private Sub IdButtonCreateMeeting0OnAction(): IdButtonCreateMeetingOnAction: End Sub
Private Sub IdButtonOpenMeetingOnAction()
    OpenDocuments "meeting", False
End Sub

