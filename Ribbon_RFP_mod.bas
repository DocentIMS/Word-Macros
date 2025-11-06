Attribute VB_Name = "Ribbon_RFP_mod"
Option Explicit
Private Function HasNoRFP() As Boolean
    On Error Resume Next
'    ProjectInfo.Remove "RFPURL"
    RFPURL = Get1DocURL("rfp")
'    ProjectInfo.Add "RFPURL", RFPURL
    HasNoRFP = Len(RFPURL) = 0
    If Not HasNoRFP Then
        frmMsgBox.Display "There is already an RFP document uploaded. Openning it instead."
        DownloadProjectInfo
'        RefreshRFPGroup
        OpenRFP
    End If
End Function

'==========
'RFP
'==========
Private Function IdButtonOpenRFPGetVisible(ID As String)
    IdButtonOpenRFPGetVisible = Len(RFPURL) > 0 And Not GetButtonVisible(7) And Not GetButtonVisible(2)
End Function
Private Function IdButtonRFPUploadMeGetVisible(ID As String)
    IdButtonRFPUploadMeGetVisible = Len(RFPURL) = 0 And GetButtonVisible(7)
End Function
Private Function IdButtonRFPUpload1GetVisible(ID As String)
    IdButtonRFPUpload1GetVisible = Len(RFPURL) > 0 And Not GetButtonVisible(2) And Not GetButtonVisible(7)
End Function
Private Function IdSplitButtonRFPUploadGetVisible(ID As String)
    IdSplitButtonRFPUploadGetVisible = Len(RFPURL) = 0 And Not GetButtonVisible(2) And Not GetButtonVisible(7)
End Function
Private Function IdButtonRFPUploadGetLabel(): IdButtonRFPUploadGetLabel = GetSplitRFPLabel: End Function
Private Function IdButtonRFPUpload0GetLabel(): IdButtonRFPUpload0GetLabel = GetSplitRFPLabel: End Function
Private Function GetSplitRFPLabel(): GetSplitRFPLabel = IIf(Len(RFPURL) > 0, "Open Current RFP", "Upload This RFP"): End Function
Private Function IdButtonRFPUploadGetSupertip()
    IdButtonRFPUploadGetSupertip = "If the open document is an RFP, press to begin analysis."
End Function
Private Function IdButtonRFPUploadGetScreentip()
    IdButtonRFPUploadGetScreentip = "RFP Analyzer"
End Function
'Private Function IdButtonRFPBrowseGetLabel(): IdButtonRFPBrowseGetLabel = IIf(RFPUploaded, "Browse to New RFP", "Browse to RFP"): End Function
Private Sub RFPUploadAction()
    If Len(RFPURL) > 0 Then
        OpenRFP
    ElseIf HasNoRFP Then
        If Documents.Count = 0 Then Exit Sub
        UploadRFP ActiveDocument
    End If
End Sub
Private Sub OpenRFP()
    If Len(RFPURL) = 0 Then RFPURL = Get1DocURL("RFP") ''GetAPIFolder(DefaultRFPFolder, "RFP")(1)("@id")
    Set NextTransitions = GetAPIFileWorkflowTransitions(RFPURL)
    OpenDocumentAt RFPURL, GetAPIContent(RFPURL).Data("review_state")
'    Dim Coll As Collection
'    Set Coll = GetAPIFolder(DefaultRFPFolder, "rfp")
'    If Not Coll Is Nothing Then If Coll.Count Then OpenDocumentAt Coll(1)("@id"): Exit Sub
'    Beep
End Sub
Private Sub IdButtonRFPUploadOnAction()
    If HasNoRFP Then RFPUploadAction
End Sub
Private Sub IdButtonRFPUpload0OnAction()
    If HasNoRFP Then RFPUploadAction
End Sub
Private Sub IdButtonRFPUpload1OnAction()
    If Documents.Count = 0 Then Exit Sub
    If HasNoRFP Then UploadRFP ActiveDocument
End Sub
Private Sub IdButtonRFPUploadMeOnAction()
    If Documents.Count = 0 Then Exit Sub
    If HasNoRFP Then UploadRFP ActiveDocument
End Sub
Private Sub IdButtonRFPBrowseOnAction()
    If HasNoRFP Then OpenAsDocentDocument "RFP"
End Sub
Private Sub IdButtonOpenRFPOnAction()
    OpenRFP
End Sub
Sub UploadRFP(Doc As Document)
    Dim FName As String
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .IsDocument = True
        .PName = ProjectNameStr
        .PURL = ProjectURLStr
        .DocType = "RFP"
    End With
    SetMetaData Doc
    FName = SaveForUpload("RFP")
    CreateAPIContent "rfp", DefaultRFPFolder, Array("file"), Array(FName)
    DownloadProjectInfo
    RefreshRibbon
End Sub
'RFP Group
Private Function IdGroupRFPGetVisible(ID As String)
    IdGroupRFPGetVisible = GetVisibleGroup(ID)
'    If IdGroupRFPGetVisible Then FillRFPFields
End Function
'Private Function IdGroupParseRFPGetVisible(id As String): IdGroupParseRFPGetVisible = GetVisibleGroup(id): End Function
Private Function IdButtonRFPParseGetVisible(ID As String): IdButtonRFPParseGetVisible = Not RFPParsed And GetButtonVisible(7): End Function
Private Sub IdButtonRFPParseOnAction()
    On Error Resume Next
    If Not ValidDocument("Request For Proposal") Then Exit Sub
    If IsProjectSelected Then frmSettings.ParsingMode "RFP"
End Sub

'Close Button
Function IdButtonRFPCancelGetVisible(ID As String): IdButtonRFPCancelGetVisible = GetButtonVisible(7): End Function
Sub IdButtonRFPCancelOnAction(): CancelEditingDoc: End Sub
'Private Sub IdButtonRFPParse0OnAction()
'    On Error Resume Next
'    If Not ValidDocument Then Exit Sub
'    If IsProjectSelected Then frmSettings.ParsingMode
'End Sub
'Private Sub IdButtonRFPParseBrowseOnAction()
'    On Error Resume Next
'    If Not OpenAsDocentDocument("RFP") Is Nothing Then
'        If Not ValidDocument Then Exit Sub
'        If IsProjectSelected Then frmSettings.ParsingMode
'    End If
'End Sub
'Private Function IdButtonRFPParseGetEnabled(): IdButtonRFPParseGetEnabled = Documents.Count > 0: End Function
Private Function IdButtonRFPParseGetSupertip()
    If Documents.Count > 0 Then
        IdButtonRFPParseGetSupertip = "The document is parsed based on assigned section breaks."
    Else
        IdButtonRFPParseGetSupertip = "An RFP Document must be open."
    End If
End Function

