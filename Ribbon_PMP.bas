Attribute VB_Name = "Ribbon_PMP"
Private Const CurrentMod As String = "Ribbon_PMP"
Option Explicit
Private Function HasNoPMP() As Boolean
    On Error Resume Next
'    ProjectInfo.Remove "PMPURL"
    PMPURL = Get1DocURL("pmp")
'    ProjectInfo.Add "PMPURL", PMPURL
    HasNoPMP = Len(PMPURL) = 0
    If Not HasNoPMP Then
        frmMsgBox.Display "There is already a PMP document uploaded. Openning it instead."
        DownloadProjectInfo
'        RefreshPmpGroup
        IdButtonOpenPMPOnAction
    End If
End Function
Private Sub IdButtonPMPUploadOnAction()
    If Not HasNoPMP Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .IsDocument = True
        .PName = ProjectNameStr
        .PURL = ProjectURLStr
        .DocType = "PMP"
    End With
    SetMetaData ActiveDocument
    frmMsgBox.Display "Uploading... Please wait...", Array(), None, "", ShowModal:=vbModeless
    frmMsgBox.Repaint
    UploadDoc ActiveDocument, NoSpelling:=True
    On Error Resume Next
    Unload frmMsgBox
End Sub
Private Sub IdButtonPMPCreate0OnAction()
    If HasNoPMP Then OpenTemplate "PMP"
End Sub
Private Sub IdButtonPMPCreateOnAction()
    If HasNoPMP Then OpenTemplate "PMP"
End Sub
Private Sub IdButtonPMPBrowseOnAction()
    If Not HasNoPMP Then Exit Sub
    Dim Doc As Document
    Set Doc = OpenAsDocentDocument("PMP")
    If Doc Is Nothing Then Exit Sub
    frmMsgBox.Display "Uploading... Please wait...", Array(), None, "", ShowModal:=vbModeless
    frmMsgBox.Repaint
    UploadDoc Doc, NoSpelling:=True
    On Error Resume Next
    Unload frmMsgBox
End Sub
Private Function IdSplitButtonPMPCreateGetVisible(ID As String)
    IdSplitButtonPMPCreateGetVisible = Len(PMPURL) = 0 And Not GetButtonVisible(3)
End Function
Private Function IdButtonOpenPMPGetVisible(ID As String)
    IdButtonOpenPMPGetVisible = Len(PMPURL) > 0 And Not GetButtonVisible(3)
End Function

'==========
'PMP
'==========
'Private Function IdButtonPMPCreateGetVisible(id As String)
'    IdButtonPMPCreateGetVisible = Not PMPUploaded And Not GetButtonVisible(3)
'End Function
'Private Function IdButtonPMPUploadGetVisible(id As String)
'    IdButtonPMPUploadGetVisible = Not PMPUploaded And Not GetButtonVisible(3)
'End Function

'Private Function IdButtonPMPCreateGetEnabled()
'    IdButtonPMPCreateGetEnabled = Not PMPUploaded
'End Function
'Private Function IdButtonPMPCreateGetLabel()
'    On Error Resume Next
'    If Not PMPUploaded Then PMPUploaded = GetAPIFolder(DefaultPMPFolder, "PMP").Count > 0
'    IdButtonPMPCreateGetLabel = IIf(PMPUploaded, "Open PMP", "Create PMP")
'End Function
Private Sub IdButtonOpenPMPOnAction()
'    Dim URL As String
    If Len(PMPURL) = 0 Then PMPURL = Get1DocURL("pmp") ''GetAPIFolder(DefaultPMPFolder, "PMP")(1)("@id")
    Set NextTransitions = GetAPIFileWorkflowTransitions(PMPURL)
    OpenDocumentAt PMPURL, GetAPIContent(PMPURL).Data("review_state")
End Sub
Private Function IdGroupPMPGetVisible(ID As String)
    IdGroupPMPGetVisible = GetVisibleGroup(ID)
'    If IdGroupPMPGetVisible Then FillPMPFields
End Function
Private Sub IdButtonPMPAddSameOnAction(): AddSameLevel: End Sub
Private Sub IdButtonPMPAddSubOnAction(): AddSubLevel: End Sub
Private Sub IdButtonPMPRevesionOnAction(): UpdateRevision: End Sub
Private Sub IdButtonPMPUnlockOnAction(): UnlockDocument: End Sub
Private Sub IdButtonPMPAddTopOnAction(): AddTopLevel: End Sub
Private Sub IdButtonPMPCancelOnAction(): CancelEditingDoc: End Sub

Private Function IdButtonPMPAddSameGetVisible(ID As String): IdButtonPMPAddSameGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonPMPAddSubGetVisible(ID As String): IdButtonPMPAddSubGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonPMPRevesionGetVisible(ID As String): IdButtonPMPRevesionGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonPMPUnlockGetVisible(ID As String): IdButtonPMPUnlockGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonPMPAddTopGetVisible(ID As String): IdButtonPMPAddTopGetVisible = GetButtonVisible(3): End Function
Private Function IdButtonPMPCancelGetVisible(ID As String): IdButtonPMPCancelGetVisible = GetButtonVisible(3): End Function

'Parsing Group
Private Function IdGroupParsePMPGetVisible(ID As String): IdGroupParsePMPGetVisible = GetVisibleGroup(ID): End Function
Private Function IdButtonParsePMPGetVisible(ID As String): IdButtonParsePMPGetVisible = Not PMPParsed And GetButtonVisible(3): End Function
Private Sub IdButtonParsePMPOnAction()
    On Error Resume Next
    If Not ValidDocument("PMP") Then Exit Sub
    If IsProjectSelected Then frmSettings.ParsingMode "PMP"
End Sub
'Private Sub IdButtonParsePMP0OnAction()
'    On Error Resume Next
'    If Not ValidDocument Then Exit Sub
'    If IsProjectSelected Then frmSettings.ParsingMode
'End Sub
'Private Sub IdButtonParsePMPBrowseOnAction()
'    On Error Resume Next
'    If Not OpenAsDocentDocument("PMP") Is Nothing Then
'        If Not ValidDocument Then Exit Sub
'        If IsProjectSelected Then frmSettings.ParsingMode
'    End If
'End Sub
Private Function IdButtonParsePMPGetEnabled(): IdButtonParsePMPGetEnabled = Documents.Count > 0: End Function
Private Function IdButtonParsePMPGetSupertip()
    If Documents.Count > 0 Then
        IdButtonParsePMPGetSupertip = "The document is parsed based on assigned section breaks."
    Else
        IdButtonParsePMPGetSupertip = "A PMP Document must be open."
    End If
End Function


