Attribute VB_Name = "Ribbon_Scope"
Option Explicit
Private Const CurrentMod As String = "Ribbon_Scope"
Private Function HasNoScope() As Boolean
    On Error Resume Next
    ScopeURL = Get1DocURL("scope")
    HasNoScope = Len(ScopeURL) = 0
    If Not HasNoScope Then
        frmMsgBox.Display "There is already a Scope document uploaded. Openning it instead."
        DownloadProjectInfo
'        RefreshScopeGroup
        IdButtonOpenScopeOnAction
    End If
End Function
Private Sub IdButtonScopeUploadOnAction()
    If Not HasNoScope Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .IsDocument = True
        .PName = ProjectNameStr
        .PURL = ProjectURLStr
        .DocType = "Scope"
    End With
    SetMetaData ActiveDocument
    
'    ActiveDocument.Save
    frmMsgBox.Display "Uploading... Please wait...", Array(), None, "", ShowModal:=vbModeless
    frmMsgBox.Repaint
    UploadDoc ActiveDocument, NoSpelling:=True
    On Error Resume Next
    Unload frmMsgBox
End Sub
Private Sub IdButtonScopeCreate0OnAction()
    If HasNoScope Then OpenTemplate "Scope"
End Sub
Private Sub IdButtonScopeCreateOnAction()
    If HasNoScope Then OpenTemplate "Scope"
End Sub
Private Sub IdButtonScopeBrowseOnAction()
    If Not HasNoScope Then Exit Sub
    Dim Doc As Document
    Set Doc = OpenAsDocentDocument("Scope")
    If Doc Is Nothing Then Exit Sub
    frmMsgBox.Display "Uploading... Please wait...", Array(), None, "", ShowModal:=vbModeless
    frmMsgBox.Repaint
    UploadDoc Doc, NoSpelling:=True
    On Error Resume Next
    Unload frmMsgBox
End Sub
Private Function IdSplitButtonScopeCreateGetVisible(ID As String)
    IdSplitButtonScopeCreateGetVisible = Len(ScopeURL) = 0 And Not GetButtonVisible(3)
End Function
Private Function IdButtonOpenScopeGetVisible(ID As String)
    IdButtonOpenScopeGetVisible = Len(ScopeURL) > 0 And Not GetButtonVisible(3)
End Function

'==========
'Scope
'==========
'Private Function IdButtonScopeCreateGetVisible(id As String)
'    IdButtonScopeCreateGetVisible = Not ScopeUploaded And Not GetButtonVisible(3)
'End Function
'Private Function IdButtonScopeUploadGetVisible(id As String)
'    IdButtonScopeUploadGetVisible = Not ScopeUploaded And Not GetButtonVisible(3)
'End Function

'Private Function IdButtonScopeCreateGetEnabled()
'    IdButtonScopeCreateGetEnabled = Not ScopeUploaded
'End Function
'Private Function IdButtonScopeCreateGetLabel()
'    On Error Resume Next
'    If Not ScopeUploaded Then ScopeUploaded = GetAPIFolder(DefaultScopeFolder, "scope").Count > 0
'    IdButtonScopeCreateGetLabel = IIf(ScopeUploaded, "Open Scope", "Create Scope")
'End Function
Private Sub IdButtonOpenScopeOnAction()
    If Len(ScopeURL) = 0 Then ScopeURL = Get1DocURL("scope") ''GetAPIFolder(DefaultPMPFolder, "PMP")(1)("@id")
    Set NextTransitions = GetAPIFileWorkflowTransitions(ScopeURL)
    OpenDocumentAt ScopeURL, GetAPIContent(ScopeURL).Data("review_state")
'    Dim URL As String
'    URL = GetAPIFolder(DefaultScopeFolder, "scope")(1)("@id")
'    Set NextTransitions = GetAPIFileWorkflowTransitions(URL)
'    OpenDocumentAt URL
End Sub
Private Function IdGroupScopeGetVisible(ID As String)
    IdGroupScopeGetVisible = GetVisibleGroup(ID)
'    If IdGroupScopeGetVisible Then FillScopeFields
End Function
Private Sub IdButtonScopeAddSameOnAction(): AddSameLevel: End Sub
Private Sub IdButtonScopeAddSubOnAction(): AddSubLevel: End Sub
Private Sub IdButtonScopeRevesionOnAction(): UpdateRevision: End Sub
Private Sub IdButtonScopeUnlockOnAction(): UnlockDocument: End Sub
Private Sub IdButtonScopeAddTopOnAction(): AddTopLevel: End Sub
'Private Sub IdButtonScopeCancelOnAction(): CancelEditingDoc: End Sub
Private Sub IdButtonScopeAddMeetingOnAction(): AddMeetingToScopeTask: End Sub

Private Function IdButtonScopeAddSameGetVisible(ID As String): IdButtonScopeAddSameGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonScopeAddSubGetVisible(ID As String): IdButtonScopeAddSubGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonScopeRevesionGetVisible(ID As String): IdButtonScopeRevesionGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonScopeUnlockGetVisible(ID As String): IdButtonScopeUnlockGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonScopeAddTopGetVisible(ID As String): IdButtonScopeAddTopGetVisible = GetButtonVisible(3): End Function
'Private Function IdButtonScopeCancelGetVisible(id As String): IdButtonScopeCancelGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonScopeAddMeetingGetVisible(ID As String): IdButtonScopeAddMeetingGetVisible = GetButtonVisible(3): End Function

Private Function IdButtonScopeAddSameGetEnabled(): IdButtonScopeAddSameGetEnabled = GetButtonVisible(8): End Function
Private Function IdButtonScopeAddSubGetEnabled(): IdButtonScopeAddSubGetEnabled = GetButtonVisible(8): End Function
Private Function IdButtonScopeRevesionGetEnabled(): IdButtonScopeRevesionGetEnabled = GetButtonVisible(8): End Function
Private Function IdButtonScopeUnlockGetEnabled(): IdButtonScopeUnlockGetEnabled = GetButtonVisible(8): End Function
Private Function IdButtonScopeAddTopGetEnabled(): IdButtonScopeAddTopGetEnabled = GetButtonVisible(8): End Function
'Private Function IdButtonScopeCancelGetEnabled(): IdButtonScopeCancelGetEnabled = GetButtonVisible(8): End Function
Private Function IdButtonScopeAddMeetingGetEnabled(): IdButtonScopeAddMeetingGetEnabled = GetButtonVisible(8): End Function

'Parsing Group
Private Function IdGroupParseScopeGetVisible(ID As String): IdGroupParseScopeGetVisible = Not ScopeParsed And GetVisibleGroup(ID): End Function
Private Function IdButtonParseScopeGetVisible(ID As String): IdButtonParseScopeGetVisible = GetButtonVisible(3): End Function
Private Sub IdButtonParseScopeOnAction()
    On Error Resume Next
    If Not ValidDocument("Scope") Then Exit Sub
    If IsProjectSelected Then
        frmSettings.ParsingMode "Scope"
    End If
End Sub
'Private Sub IdButtonParseScope0OnAction()
'    On Error Resume Next
'    If Not ValidDocument Then Exit Sub
'    If IsProjectSelected Then frmSettings.ParsingMode
'End Sub
'Private Sub IdButtonParseScopeBrowseOnAction()
'    On Error Resume Next
'    If Not OpenAsDocentDocument("Scope") Is Nothing Then
'        If Not ValidDocument Then Exit Sub
'        If IsProjectSelected Then frmSettings.ParsingMode
'    End If
'End Sub
Private Function IdButtonParseScopeGetEnabled(): IdButtonParseScopeGetEnabled = Documents.Count > 0: End Function
Private Function IdButtonParseScopeGetSupertip()
    If Documents.Count > 0 Then
        IdButtonParseScopeGetSupertip = "The document is parsed based on assigned section breaks."
    Else
        IdButtonParseScopeGetSupertip = "A Scope Document must be open."
    End If
End Function

