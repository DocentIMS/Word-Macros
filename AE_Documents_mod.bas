Attribute VB_Name = "AE_Documents_mod"
'Option Explicit
'Option Compare Text
'Const CurrentMod = "Documents_mod"
'Private Const NotSavedLocal = "The document could not be saved to the following locations:"
'Private Const NotSavedTemplates = "The following templates were not updated:"
'Private Const NotSavedWeb = "     Document Upload Failed"
'Private Const RemainsOpenMsg = "This document remains open after upload because it could not " & _
'                            "be saved to one or more locations shown above.  Please check on the reasons and fix."
'Private Const CCDDPlaceholder = "Choose..."
'
'
'Sub ShowDocumentInfo()
'    Dim isDoc As Boolean, isTemp As Boolean
'    Dim Msg As String, DocType As String
'    On Error Resume Next
'    isDoc = GetProperty(pIsDocument)
'    DocType = "Document "
'    If Not isDoc Then
'        Msg = "Document Type: Unknown"
'    Else
'        isTemp = GetProperty(pIsTemplate)
'        If isTemp Then
'            DocType = "Template "
'            Msg = "Project: " & GetProperty(pPName) & Chr(10) & Chr(10)
'            Msg = Msg & "Template Info:" & Chr(10)
'            If Len(GetProperty(pDocType)) Then Msg = Msg & "   Type: " & GetProperty(pDocType) & Chr(10)
'            If Len(GetProperty(pTemplateVer)) Then Msg = Msg & "   Version: " & GetProperty(pTemplateVer) & Chr(10) & Chr(10)
'            If Len(GetProperty(pTemplateDate)) Then Msg = Msg & "   Date: " & GetProperty(pTemplateDate) & Chr(10)
'            If Len(GetProperty(pAuthor)) Then Msg = Msg & "   Author: " & GetProperty(pAuthor)
'        Else
'            Msg = "Project: " & GetProperty(pPName) & Chr(10) & Chr(10)
'            Msg = Msg & "Template Info:" & Chr(10)
'            If Len(GetProperty(pTemplateVer)) Then Msg = Msg & "   Version: " & GetProperty(pTemplateVer) & Chr(10)
'            If Len(GetProperty(pTemplateDate)) Then Msg = Msg & "   Revision Date: " & Format(GetProperty(pTemplateDate), DateTimeFormat) & Chr(10) & Chr(10)
'            Msg = Msg & "Document Info:" & Chr(10)
'            If Len(GetProperty(pDocType)) Then Msg = Msg & "   Type: " & GetProperty(pDocType) & Chr(10)
'            If Len(GetProperty(pDocVer)) Then Msg = Msg & "   Version: " & GetProperty(pDocVer) & Chr(10)
'            If Len(GetProperty(pDocState)) Then Msg = Msg & "   State: " & GetProperty(pDocState) & Chr(10)
'            If Len(GetProperty(pDocCreateDate)) Then Msg = Msg & "   Creation Date: " & Format(GetProperty(pDocCreateDate), DateTimeFormat) & Chr(10)
'            If Len(GetProperty(pAuthor)) Then Msg = Msg & "   Author: " & GetProperty(pAuthor)
'        End If
'    End If
'    frmMsgBox.Width = frmMsgBox.Width + 20
'    frmMsgBox.Display Msg, , Information, DocType & "Information"
'End Sub
'Sub DocumentSelected(SelectedItem As Long)
'    On Error Resume Next
'    If CodeIsRunning Then Exit Sub
'    WriteLog 1, CurrentMod, "DocumentSelected"
'    DocNum = SelectedItem
'End Sub
'Sub MeetingDocumentSelected(SelectedItem As Long)
'    On Error Resume Next
'    If CodeIsRunning Then Exit Sub
'    WriteLog 1, CurrentMod, "MeetingDocumentSelected"
'    MeetingDocNum = SelectedItem
'End Sub
'Function OpenDocumentAt(ByVal URL As String, Optional ByVal State As String) As Document
'    Dim Doc As Document, i As Long
'    URL = Replace(URL, ProjectURLStr, "")
'    If Len(State) Then
'        Set OpeningDocInfo = New DocInfo
'        OpeningDocInfo.UpdateMode = True
'        OpeningDocInfo.DocState = State
'    End If
'    Set Doc = Documents.Open(DownloadAPIFile(URL, False))
'    If Doc Is Nothing Then Exit Function
'    LockAPIFile URL
'    Unprotect Doc
''    SetProperty pDocURL, URL, Doc, msoPropertyTypeString
''    If Len(State) Then SetProperty pDocState, State, Doc, msoPropertyTypeString
'    If State = "Published" Then SetContentControl "DocDate", "Published Date: " & GetProperty(pPublishDate, Doc), Doc
'    On Error Resume Next
'    Doc.Bookmarks("DocentTemplateStart").Select
'    Protect Doc
'    Set OpenDocumentAt = Doc
'End Function
'Sub OpenDocuments(ByVal DocType As String, CreateMode As Boolean)
'    Dim DocsResp As Collection 'Arr
'    'DocType = LCase(Replace(DocumentName(DocNum), " ", "_"))
''    DocType = LCase(Replace(DocType, " ", "_"))
'    If InStr(DocType, "meeting") Then
'        Set DocsResp = GetMeetingDocOfType(LCase(Replace(DocType, " ", "_")), Not CreateMode, "")
'    Else
'        Set DocsResp = GetDocumentsOfType(DocType, DefaultDocumentsFolder)
'    End If
'    If IsGoodResponse(DocsResp) Then frmOpenDoc.ShowDocuments DocsResp, DocType, CreateMode
'End Sub
'Sub SetMetaData(Doc As Document)
'    On Error Resume Next
'    If OpeningDocInfo Is Nothing Then Exit Sub
'    With OpeningDocInfo
'        If Not .UpdateMode Or Len(.IsTemplate) Then SetProperty pIsTemplate, .IsTemplate, , msoPropertyTypeBoolean
'        If Not .UpdateMode Or Len(.DocURL) Then SetProperty pDocURL, .DocURL, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.DocState) Then SetProperty pDocState, .DocState, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.IsDocument) Then SetProperty pIsDocument, .IsDocument, , msoPropertyTypeBoolean
'        If Not .UpdateMode Or Len(.DocType) Then SetProperty pDocType, .DocType, , msoPropertyTypeString  ' msoPropertyTypeNumber
'        If Not .UpdateMode Or Len(.PName) Then SetProperty pPName, .PName, , msoPropertyTypeString  'msoPropertyTypeNumber
'        If Not .UpdateMode Or Len(.PURL) Then SetProperty pPURL, .PURL, , msoPropertyTypeString  'msoPropertyTypeNumber
'        If Not .UpdateMode Or Len(.DocVer) Then SetProperty pDocVer, .DocVer, , msoPropertyTypeNumber
'        If Not .UpdateMode Or Len(.ContractNo) Then SetProperty pContractNo, .ContractNo, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.DocCreateDate) Then SetProperty pDocCreateDate, .DocCreateDate, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.publishDate) Then SetProperty pPublishDate, .publishDate, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.MeetingType) Then SetProperty pMeetingType, .MeetingType, , msoPropertyTypeString
'        If Not .UpdateMode Or Len(.IsFinalRev) Then SetProperty pIsFinalRev, .IsFinalRev, , msoPropertyTypeBoolean
'    End With
'    RefreshDictionary
'    LoadDocInfo Doc
'ex:
'    Set OpeningDocInfo = Nothing
'End Sub
'Sub RefreshDictionary()
'    Dim DDict As Word.Dictionary
'    On Error Resume Next
'    If GetProperty(pIsDocument) Then
'        Application.CustomDictionaries.ClearAll
''        If GetDictNo(DocentDictionaryName) = 0 Then
''            Set DDict = Application.CustomDictionaries.Add(DocentDictionaryPath)
''        Else
''            Set DDict = CustomDictionaries(GetDictNo(DocentDictionaryName))
''        End If
''        Application.CustomDictionaries.ActiveCustomDictionary = DDict
'
''        'UploadAPIFile DocentDictionaryPath, DictionaryServerPath, DashboardSite
''        CustomDictionaries(GetDictNo(DocentDictionaryName)).Delete
''        DoEvents: Sleep 100
''        Set DDict = CustomDictionaries.Add(DocentDictionaryPath)
''        DoEvents: Sleep 100
''        CustomDictionaries.ActiveCustomDictionary = DDict
''        DoEvents: Sleep 100
''        'ActiveDocument.SpellingChecked = False
''        'ActiveDocument.CheckSpelling DDict ', , False
'    Else
'        CustomDictionaries(GetDictNo(DocentDictionaryName)).Delete
'        DoEvents: Sleep 100
''        ActiveDocument.CheckSpelling
'    End If
''    ActiveDocument.Range.LanguageID = wdFrenchHaiti
''    ActiveDocument.Range.LanguageID = wdEnglishUS
''    Application.DisplayAlerts = False
''    ActiveDocument.SpellingChecked = False
''    ActiveDocument.CheckSpelling DDict ', , False
''    Application.DisplayAlerts = wdAlertsAll
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Function GetDictNo(DictName As String) As Long
''    Dim i As Long
'    For GetDictNo = CustomDictionaries.Count To 1 Step -1
'        If CustomDictionaries(GetDictNo).Name = DictName Then
'            Exit Function
'        End If
'    Next
'End Function
'Sub CopyAgendaItems(Doc As Document, ADoc As Document)
'    Dim cRng As Range, pRng As Range
'    Unprotect Doc
'    Set cRng = ADoc.Range.GoToEditableRange(wdEditorCurrent)
'    Doc.Range.Bookmarks("DocentMNStart").Range.FormattedText = cRng.FormattedText
'    Set pRng = Doc.Range.Bookmarks("DocentMNStart").Range
'    pRng.MoveEnd 1, cRng.Characters.Count
'    pRng.Font.ColorIndex = wdGray50
'    ADoc.Close False
'    Protect Doc
'    Doc.Activate
'End Sub
'Sub ApplyLocationHyperlink(URL, Doc As Document)
'    If IsNull(URL) Then Exit Sub
'    Dim Rng As Range
'    Set Rng = FindCC("PlannedLocation", Doc).Range
'    HyperlinkWord Rng, "Zoom", CStr(URL)
'    HyperlinkWord Rng, "Teams", CStr(URL)
'End Sub
'Sub UpdateMeetingFile(Doc As Document, xColl As Dictionary)
'
'    Dim MeetingDocURL As String, MeetingTitle As String, MeetingLocation As String
'    Dim PlannedTasks As String, ProposedTasks As String
'    Dim MeetingStart As String, MeetingEnd As String
'    Dim ActualValue As String ', ActualMeetingEnd As String
'    Dim PrpMgr As String, ActualPrpMgr As String
'    Dim ActualsDict As Dictionary
'    Dim PlannedGroups As Collection, ActualGroups As Collection
'    Dim PlannedIndividuals As Collection, ActualIndividuals As Collection
'                        ', MNAtt As String, MNExAtt As String, _
'
'
'    Unprotect Doc
'    On Error Resume Next
'    Set ActualsDict = ParseJson(xColl("ParentDoc")("actuals"))
'    SetProperty pMeetingUID, xColl("ParentMeeting")("UID"), , msoPropertyTypeString
'
'    MeetingDocURL = xColl("ParentMeeting")("online_meeting_link")
'    SetProperty pOnlineMeetingUID, MeetingDocURL, Doc, msoPropertyTypeString
'
''    ApplyMeetingHyperlink xColl, Doc
'
'    MeetingDocURL = Replace(xColl("@id"), ProjectURLStr, "")
'    SetProperty pDocURL, MeetingDocURL, Doc, msoPropertyTypeString
'
'    MeetingStart = xColl("ParentMeeting")("MeetingDateTime")
'    SetContentControl "PlannedDate", Format(MeetingStart, LongDateFormat), Doc
'    SetContentControl "PlannedTime", Format(MeetingStart, TimeFormat), Doc
'    SetContentControl "MeetingMonth", Format(MeetingStart, "mmmm"), Doc
'
'    MeetingEnd = ToServerTime(CStr(xColl("ParentMeeting")("end")))
'    SetContentControl "PlannedEndTime", Format(MeetingEnd, TimeFormat), Doc
'
'    ActualValue = ""
'    ActualValue = ActualsDict("Motion to End") 'xColl("ParentDoc")("actual_meeting_date_time")
'    If Len(ActualValue) Then SetContentControl "MeetingEnder", ActualValue, Doc   'MeetingEnder
'    PopulateMembers "MeetingEnder", , Doc
'
'    ActualValue = ""
'    ActualValue = ActualsDict("2nd. Motion to End") 'xColl("ParentDoc")("actual_meeting_date_time")
'    If Len(ActualValue) Then SetContentControl "MeetingEnder2", ActualValue, Doc   'MeetingEnder2
'    PopulateMembers "MeetingEnder2", , Doc
'
'    ActualValue = ""
'    ActualValue = ActualsDict("Start Time") 'xColl("ParentDoc")("actual_meeting_date_time")
'    If Len(ActualValue) Then
'        SetContentControl "ActualTime", Format(ActualValue, TimeFormat), Doc
'    Else
'        FillEmptyActual "Time", Doc
'    End If
'
'    ActualValue = ""
'    ActualValue = ActualsDict("Meeting Location") 'xColl("ParentDoc")("actual_meeting_date_time")
'    MeetingLocation = xColl("ParentMeeting")("location")("title")
'    SetContentControl "PlannedLocation", MeetingLocation, Doc
'    ApplyLocationHyperlink xColl("ParentMeeting")("event_url"), Doc
'
''    PopulateLocations MeetingLocation, Doc
'    If Len(ActualValue) Then
''        PopulateLocations ActualLocation, Doc 'SetContentControl "ActualLocation", ActualValue, Doc
'        SetContentControl "ActualLocation", ActualValue, Doc
'    Else
'        PopulateLocations MeetingLocation, Doc
'    End If
'
'    ActualValue = ""
'    ActualValue = ActualsDict("Meeting Date") 'xColl("ParentDoc")("actual_meeting_date_time")
'    If Len(ActualValue) Then
'        SetContentControl "ActualDate", Format(ActualValue, LongDateFormat), Doc
'    Else
'        FillEmptyActual "Date", Doc
'    End If
'    ActualValue = ""
'    ActualValue = ActualsDict("End Time")
'    If Len(ActualValue) Then
'        SetContentControl "ActualEndTime", Format(ActualValue, TimeFormat), Doc
'    Else
'        FillEmptyActual "EndTime", Doc
'    End If
'
'    MeetingTitle = xColl("ParentMeeting")("title")
'    SetContentControl "MeetingTitle", MeetingTitle, Doc
'    PlannedTasks = xColl("ParentDoc")("planned_action_items")
'    If Len(PlannedTasks) Then
'        SetProperty pPlannedTasks, PlannedTasks, Doc
'        FillTable PlannedTasks, False, Doc '"Planned Tasks"
'    End If
'    ProposedTasks = xColl("ParentDoc")("proposed_action_items")
'    If Len(ProposedTasks) Then
'        SetProperty pProposedTasks, ProposedTasks, Doc
'        FillTable ProposedTasks, True, Doc '"Proposed Tasks"
'    End If
'    Set PlannedGroups = xColl("ParentMeeting")("attendees_group")
'    If Not PlannedGroups Is Nothing Then FillPlannedGroups PlannedGroups, Doc
''    Set ActualGroups = ActualsDict("Groups").Collection
''    If Not ActualsDict("Groups") Is Nothing Then FillAttendees ActualsDict("Groups"), Doc
''    If Not ActualGroups Is Nothing Then FillPlannedGroups PlannedGroups, Doc
'
'    Set PlannedIndividuals = xColl("ParentMeeting")("attendees")
'    If Not PlannedIndividuals Is Nothing Then FillPlannedAttendees "Individual Attendees", PlannedIndividuals, Doc
''    Set ActualIndividuals = ActualsDict("Individuals").Collection
''    If Not ActualsDict("Individuals") Is Nothing Then FillAttendees ActualsDict("Individuals"), Doc
''    If Not ActualsDict Is Nothing Then FillAttendees ActualsDict, Doc
'
''    PrpMgr = xColl("ParentMeeting")("will_property_manager_attend")
''    SetContentControl "PlannedPMgrAttending", PrpMgr = "True", Doc
'
''    ActualPrpMgr = ActualsDict("Prop Mgr Attended?")
''    SetContentControl "ActualPMgrAttending", IIf(Len(ActualPrpMgr), ActualPrpMgr = "True", PrpMgr = "True"), Doc
'
'    FillAttendees ActualsDict, Doc
'    LoadDocInfo Doc
'    RefreshRibbon
'    Protect Doc
'    Doc.Saved = True
'End Sub
'Private Sub FillAttendees(ActualsDict As Dictionary, Optional Doc As Document)
'    Dim Tbl As Table, i As Long, r As Long, j As Long
'    Dim GroupName As String, IndDict As Dictionary, AllMembers As Dictionary
'    Dim x As Long 'Arr(1 To 2) As String,
'    If ActualsDict Is Nothing Then Exit Sub
'    Set Tbl = GetTableByTitle("Attendees", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
''    Arr(1) = "Groups"
''    Arr(2) = "Individuals"
'    r = 2
'    Set AllMembers = GetAllMembers
'    For x = 1 To ActualsDict.Count
'        If TypeName(ActualsDict(x)) = "Dictionary" Then
'            For i = 1 To ActualsDict(x).Count
'                For j = 1 To ActualsDict(x)(i).Count
''                Stop
'                    If ActualsDict(x)(i)(j) <> "Select Attendees" And ActualsDict(x)(i)(j) <> "N/A" And ActualsDict(x)(i)(j) <> "False" And ActualsDict(x)(i)(j) <> "True" Then
'                        GroupName = ActualsDict(x).KeyName(i)
'                        Set IndDict = AllMembers(ActualsDict(x)(i)(j))
'                        Tbl.Rows.Add
'                        r = r + 1
'                        Tbl.Rows(r).Cells(1).Range.text = GroupName
'                        Tbl.Rows(r).Cells(2).Range.text = IndDict("fullname")
'                        Tbl.Rows(r).Cells(3).Range.text = IndDict("company")
'                        Tbl.Rows(r).Cells(4).Range.text = IndDict("email")
''                        Set IndDict = getMemberDict(ActualsDict(x)(i)(j))
'                    End If
'                Next
'            Next
''        ElseIf TypeName(ActualsDict(x)) = "Collection" Then
''            For i = 1 To ActualsDict(x).Count
'''                For j = 1 To ActualsDict(x)(i).Count
'''                Stop
''                    If ActualsDict(x)(i) <> "Select Attendees" And ActualsDict(x)(i) <> "N/A" And ActualsDict(x)(i) <> "False" And ActualsDict(x)(i)(j) <> "True" Then
''                        'GroupName = ActualsDict(x).KeyName(i)
'''                        Set IndDict = GetUserInfo(CStr(ActualsDict(x)(i)), "fullname", "")
''                        Tbl.Rows.Add
''                        r = r + 1
''                        Tbl.Rows(r).Cells(1).Range.Text = GroupName
''                        Tbl.Rows(r).Cells(2).Range.Text = CStr(ActualsDict(x)(i)) 'IndDict("fullname")
''                        Tbl.Rows(r).Cells(3).Range.Text = GetUserInfo(CStr(ActualsDict(x)(i)), "fullname", "company") 'IndDict("company")
''                        Tbl.Rows(r).Cells(4).Range.Text = GetUserInfo(CStr(ActualsDict(x)(i)), "fullname", "email") 'IndDict("email")
'''                        Set IndDict = getMemberDict(ActualsDict(x)(i)(j))
''                    End If
'''                Next
''            Next
'        End If
'    Next
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Sub FillPlannedAttendees(ReferenceStr As String, Coll As Collection, Optional Doc As Document)
'    If Coll.Count = 0 Then Exit Sub
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, i As Long, r As Long, Atts As String
'    Set Tbl = GetTableByTitle("Meeting Details", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    r = 1
'    Do: r = r + 1: Loop Until CellText(Tbl.Rows(r).Cells(1).Range.text) = ReferenceStr
'    If CellText(Tbl.Rows(r).Cells(1).Range.text) <> ReferenceStr Then Exit Sub
''    If Coll.Count Then
'    Atts = Join(CollToArr(Coll), Chr(10))
''    Else
''        Atts = "N/A"
''    End If
'    Tbl.Rows(r).Cells(2).Range.text = Atts
'    SetContentControl Tbl.Rows(r).Cells(3).Range.ContentControls(1).Title, Atts
'    If ReferenceStr <> "Individual Attendees" Then CopyEditCC Tbl.Rows(r)
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Sub FillPlannedGroups(Coll As Collection, Optional Doc As Document)
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, r As Long, i As Long, j As Long
'    Set Tbl = GetTableByTitle("Meeting Details", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    r = 1
'    Do: r = r + 1: Loop Until CellText(Tbl.Rows(r - 1).Cells(1).Range.text) = "Groups"
'    If CellText(Tbl.Rows(r).Cells(1).Range.text) <> "No Groups" Then Exit Sub
''    If Coll.Count > 1 Then
''    End If
'    For i = Coll.Count To 1 Step -1
''        r = r + 1
'        If i < Coll.Count Then
'            For j = 1 To Tbl.Rows(r).Range.ContentControls.Count
'                Tbl.Rows(r).Range.ContentControls(j).LockContentControl = False
'                Tbl.Rows(r).Range.ContentControls(j).LockContents = False
'            Next
'            Tbl.Rows.Add Tbl.Rows(r)
'        End If
'        FillPlannedGroupRow Tbl.Rows(r), CStr(Coll(i)("token")) '("title"))
'    Next
'    If Coll.Count > 1 Then
'        For r = r To r + Coll.Count - 1
'            For i = 1 To Tbl.Rows(r).Range.ContentControls.Count
'                Tbl.Rows(r).Range.ContentControls(i).LockContentControl = True
'            Next
'            Tbl.Rows(r).Cells(4).Range.ContentControls(1).LockContents = True
'        Next
'    End If
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Sub FillPlannedGroupRow(Rw As Row, GroupID As String)
'    On Error GoTo ex
'    Dim GroupName As String
'    GroupName = ProjectGroupsDict(GroupID)("title")
'    With Rw.Cells(1).Range
'        .Italic = False
'        .text = GroupName 'Coll(i)("title")
'    End With
'    Rw.Cells(2).Range.text = GroupName & " Members" 'Coll(i)("title")
'    Rw.Cells(3).Range.text = ""
'    With Rw.Cells(3).Range.ContentControls.Add '.Text = Coll(i)("title")
'        .Tag = GroupID & " Attendees"
'        .Title = GroupID & " Attendees"
'        .SetPlaceholderText text:="Select Attendees"
''        .Range.Text = GroupName & " Members"
'        .LockContentControl = True
'    End With
'    Rw.Cells(3).Range.Editors.Add wdEditorEveryone
'    CopyEditCC Rw
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Sub CopyEditCC(Rw As Row)
'    Dim j As Long
'    Rw.Parent.Rows(2).Cells(4).Range.ContentControls(1).Copy
'    On Error Resume Next
'    For j = 1 To Rw.Range.ContentControls.Count
'        Rw.Range.ContentControls(j).LockContentControl = False
'        Rw.Range.ContentControls(j).LockContents = False
'    Next
'    Rw.Cells(4).Range.Paste
'    Do While Err.Number
'        Err.Clear
'        DoEvents
'        Sleep 100
'        Rw.Cells(4).Range.Paste
'    Loop
'    With Rw.Cells(4).Range.ContentControls(1)
'        .LockContentControl = True
'        .LockContents = True
'    End With
'End Sub
'Private Sub PopulateUsers(CCName As String, Members As Dictionary, Optional Doc As Document)
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim CC As ContentControl, i As Long
'    Set CC = FindCC(CCName, Doc)
'    If CC Is Nothing Then Exit Sub
''    Set Members = GetMembersOf
'    With CC
'        .DropdownListEntries.Clear
'        For i = 1 To Members.Count
'            .DropdownListEntries.Add Members(i)
'        Next
'    End With
'End Sub
'Private Sub PopulateLocations(SelectedLocation As String, Optional Doc As Document)
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim CC As ContentControl, i As Long, Locs As Collection
'    Set CC = FindCC("ActualLocation", Doc)
'    If CC Is Nothing Then Exit Sub
'    Set Locs = ProjectInfo("meeting_locations")
'    With CC
'        .DropdownListEntries.Clear
'        For i = 1 To Locs.Count
'            .DropdownListEntries.Add Locs(i)("location_name"), Locs(i)("location_name")
'            If Locs(i)("location_name") = SelectedLocation Then .DropdownListEntries(i).Select
'        Next
'    End With
'    Exit Sub
'ex:
''    Stop
'End Sub
'Private Sub PopulateMembers(CCName As String, Optional SelectedMember As String, Optional Doc As Document)
'    On Error GoTo ex
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim CC As ContentControl, i As Long, Members As Dictionary
'    Set CC = FindCC(CCName, Doc)
'    If CC Is Nothing Then Exit Sub
'    If CC.Type <> wdContentControlDropdownList Then Exit Sub
'    If Len(SelectedMember) = 0 Then SelectedMember = CC.Range.text
'    Set Members = GetMembersNames("PrjTeam")
''    Set Locs = ProjectInfo("meeting_locations")
'    With CC
'        .DropdownListEntries.Clear
'        For i = 1 To Members.Count
'            .DropdownListEntries.Add Members(i), Members(i)
'            If Members(i) = SelectedMember Then .DropdownListEntries(i).Select
'        Next
'    End With
'    Exit Sub
'ex:
''    Stop
''    Resume
'End Sub
'Private Function TasksToMetadata(Optional IsProposed As Boolean, Optional Doc As Document) As String
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim s As String
'    Dim Tbl As Table, i As Long, c As Long
'    Set Tbl = GetTableByTitle(IIf(IsProposed, "Proposed", "Planned") & " Tasks", Doc)
'    For i = 3 To Tbl.Rows.Count
'        s = s & ";"
'        For c = 1 To Tbl.Rows(i).Cells.Count
'            s = s & "," & CellText(Tbl.Rows(i).Cells(c).Range.text)
'        Next
'    Next
'    TasksToMetadata = s
'    SetProperty IIf(IsProposed, pProposedTasks, pPlannedTasks), s, Doc
''    SetProperty KWrd & "Tasks", s
''    On Error Resume Next
''    s = GetDocProperty(KWrd & "Tasks")
''    s = s & ";" & tbTitle.Value & "," & _
''                tbDetails.Value & "," & _
''                tbDueDate.Value & "," & _
''                PColl(cbPriority.Value) & "," & _
''                GetMemberId(cbWho.Value) & "," & _
''                tbNotes.Value
''    SetProperty "PlannedTasks", s
'End Function
'Private Sub UpdateTasks(Optional IsProposed As Boolean, Optional Doc As Document)
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, i As Long, c As Long
'    Dim Ps() As String, Cs() As String
'
''    i =
''    Ps = Split(GetProperty(IIf(IsProposed, pProposedTasks, pPlannedTasks), Doc), ";")
'    Ps = Split(TasksToMetadata(IsProposed), ";")
'    Set Tbl = GetTableByTitle(IIf(IsProposed, "Proposed", "Planned") & " Tasks", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    c = Tbl.Rows(1).Cells.Count
'    If Not CellText(Tbl.Rows(1).Cells(c).Range.text) Like "Closed*" Then Exit Sub
'    For i = 3 To Tbl.Rows.Count
''        If Tbl.Rows(i).Cells(c).Range.ContentControls(1).Range.Text = "Yes" Then
'            Cs = Split(Ps(i - 2), ",")
'            UpdateAPIContent "/action-items/action_items" & IIf(Cs(1) = 0, "", "-" & Cs(1)), _
'                Array("is_this_item_closed"), Array(IIf(Cs(6) = "Yes", "True", IIf(Cs(6) = "No", "False", Cs(6))))
''            CreateAPITask Cs(1), Cs(5), Cs(4), Cs(3), Cs(2), Cs(6)
''        End If
'    Next
'End Sub
'Private Sub CreateTasks(Optional Doc As Document)
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, i As Long, c As Long
'    Dim Ps() As String, Cs() As String, MtngType As String, MMURL As String
'    MtngType = GetProperty(pMeetingType)
'    MMURL = GetProperty(pDocURL)
'    Ps = Split(GetProperty(pProposedTasks, Doc), ";")
'    Set Tbl = GetTableByTitle("Proposed Tasks", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    c = Tbl.Rows(1).Cells.Count
'    If CellText(Tbl.Rows(1).Cells(c).Range.text) <> "Approved?" Then Exit Sub
'    For i = 3 To Tbl.Rows.Count
'        If Tbl.Rows(i).Cells(c).Range.ContentControls(1).Range.text = "Yes" Then
'            Cs = Split(Ps(i - 2), ",")
'            CreateAPITask Cs(1), Cs(5), Cs(4), Cs(3), Cs(2), Cs(6), Cs(7), MtngType, MMURL, , GetProperty(pMeetingUID)
'        End If
'    Next
'End Sub
'Private Function ProposedTaskDecided(Optional Doc As Document) As Boolean
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, i As Long, c As Long
'    ProposedTaskDecided = True
'    Set Tbl = GetTableByTitle("Proposed Tasks", Doc)
'    If Tbl Is Nothing Then Exit Function
'    c = Tbl.Rows(1).Cells.Count
'    If CellText(Tbl.Rows(1).Cells(c).Range.text) <> "Approved?" Then Exit Function
'    For i = 3 To Tbl.Rows.Count
'        Select Case Tbl.Rows(i).Cells(c).Range.ContentControls(1).Range.text
'        Case "Yes", "No"
'        Case Else: ProposedTaskDecided = False: Exit Function
'        End Select
'    Next
'End Function
'Private Sub FillTable(PItems As String, IsProposed As Boolean, Optional Doc As Document)
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Tbl As Table, Ps() As String, Cs() As String
'    Dim i As Long, j As Long
'    Set Tbl = GetTableByTitle(IIf(IsProposed, "Proposed", "Planned") & " Tasks", Doc)
'    If Tbl Is Nothing Then Exit Sub
'    Ps = Split(PItems, ";")
'    For i = Tbl.Rows.Count To 3 Step -1
'        Tbl.Rows(i).Delete
'    Next
'    For i = LBound(Ps) To UBound(Ps)
'        If Len(Ps(i)) Then
'        Cs = Split(Ps(i), ",")
'        With Tbl.Rows.Add
'            For j = 1 To .Cells.Count 'LBound(Cs) + 1 To UBound(Cs)
'                Select Case CellText(Tbl.Rows(1).Cells(j).Range.text)
'                Case "Approved?", "Closed?"
'                    .Cells(j).Range.Editors.Add wdEditorEveryone
'                    With .Cells(j).Range.ContentControls.Add(wdContentControlDropdownList)
'                        .DropdownListEntries.Clear
'                        .DropdownListEntries.Add "Yes", "Yes"
'                        .DropdownListEntries.Add "No", "No"
'                        .SetPlaceholderText text:=CCDDPlaceholder
'                    End With
'                Case Else
'                    .Cells(j).Range.text = Cs(j)
'                End Select
'            Next
'        End With
'        End If
'    Next
'    SetProperty IIf(IsProposed, pProposedTasks, pPlannedTasks), PItems
'End Sub
'Function GetActuals(Doc As Document) As String
'    Dim Tbl As Table, r As Long, c As Long
'    Dim ActualsDict As New Dictionary
'    Dim RowInfo As Dictionary, RowTxt As String, RowVal As String
'    Dim PartInfo As Dictionary, PartName As String
'    Set Tbl = GetTableByTitle("Meeting Details", Doc)
'    If Tbl Is Nothing Then Exit Function
'    For r = 2 To Tbl.Rows.Count
''    If r = Tbl.Rows.Count Then Stop
'        RowTxt = Trim(CellText(Tbl.Rows(r).Cells(1).Range.text))
'        If Tbl.Rows(r).Cells.Count = 1 Then
'            If Len(PartName) Then ActualsDict.Add PartName, PartInfo
'            PartName = RowTxt
'            Set PartInfo = New Dictionary
'        Else
''            If r = Tbl.Rows.Count Then
'''                RowVal = GetContentControl("ActualPMgrAttending")
''            Else
'            If Tbl.Rows(r).Cells.Count > 2 Then
'                RowVal = CellText(Tbl.Rows(r).Cells(3).Range.text)
'            End If
'            If Len(PartName) Then
'                PartInfo.Add RowTxt, ArrToColl(Split(RowVal, Chr(13))) 'CellTextOrColl(RowVal)
'            Else
'                ActualsDict.Add RowTxt, RowVal 'CellTextOrColl(RowVal)
'            End If
'        End If
'    Next
'     If Len(PartName) Then ActualsDict.Add PartName, PartInfo
'     GetActuals = ConvertToJson(ActualsDict)
'End Function
'Private Function CellTextOrColl(CellTxt As String)
'    If InStr(CellTxt, Chr(13)) Then
'        Set CellTextOrColl = ArrToColl(Split(CellTxt, Chr(13)))
'    Else
'       CellTextOrColl = CellTxt
'    End If
'End Function
'Private Sub FillEmptyActual(CCName As String, Doc As Document)
'    Select Case GetContentControl("Actual" & CCName)
'    Case " ", ""
'        SetContentControl "Actual" & CCName, GetContentControl("Planned" & CCName), Doc
'    End Select
'End Sub
'Function OpenTemplate(Optional DocName As String, Optional AsTemplate As Boolean, _
'                Optional SilentMode As Boolean, Optional MeetingMode As Boolean, _
'                Optional MtngType As String, Optional Unprotected As Boolean = False) As Document
'    Dim Doc As Document, DocURL As String, FName As String, Members As Dictionary, PWD As String
'    If AsTemplate And Not SilentMode Then
'        PWD = frmInputBox.Display("Please insert the password", "Unlock Template for Editing", "Password")
'        If PWD = "Canceled" Then
'            GoTo ex
'        ElseIf PWD <> TemplatePasswordStr Then
'            frmMsgBox.Display "Worng Password", , Critical, ""
'            GoTo ex
'        End If
'    End If
'    If InStr(DocName, "Reimbursement") Then Set Members = GetMembersOf
''    OpeningTemplate = AsTemplate
'    On Error Resume Next
'    'If Len(DocName) = 0 Then Stop 'DocName = DocumentName(DocNum)
'    If Len(DocURL) = 0 Then DocURL = DocumentsTypes(DocName)("URL")
'    FName = DownloadAPIFile(DocURL, AsTemplate, DocName & GetFileExtension(DocURL))
'    Set OpeningDocInfo = New DocInfo
'    With OpeningDocInfo
'        .ContractNo = ContractNumberStr
'        .DocState = GetInitalState(DocName)
'        .PURL = ProjectURLStr
'        .DocType = DocName
'        .IsDocument = True
'        .IsTemplate = AsTemplate
'        .Name = ""
'        .PName = ProjectNameStr
'        .DocCreateDate = Format(ToServerTime, DateFormat)
'        .publishDate = ""
'        '.TemplateVersion = ProjectNameStr
'        .DocURL = IIf(AsTemplate, DocURL, "") 'Or MeetingMode
'        .DocVer = 1
'        .MeetingType = MtngType
'    End With
'    If AsTemplate Then
'         Set NextTransitions = Nothing
'        Set Doc = Documents.Open(FName)
'    Else
'        If Not MeetingMode Then Set NextTransitions = GetInitialTransitions(DocName)
'        Set Doc = Documents.Add(FName)
'    End If
''    Set OpeningDocInfo = Nothing
'    If Doc Is Nothing Then GoTo ex
'    On Error Resume Next
'    Unprotect Doc
''    On Error GoTo 0
''        Set NextTransitions = Nothing
''    Else
''        Set NextTransitions = GetAPIFileWorkflowTransitions(DocURL)
''    End If
''    On Error Resume Next
'    SetContentControl "ProjectNumber", ContractNumberStr, Doc
'    SetContentControl "Contract Number", ContractNumberStr, Doc
'    SetContentControl "Project Name", ProjectNameStr, Doc
'    SetContentControl "Client Name", ProjectClientStr, Doc
'    SetContentControl "Customer Name", ProjectClientStr, Doc
'    SetContentControl "UserName", MainInfo("fullname"), Doc
'    PopulateUsers "Approver_1", Members, Doc
''    PopulateUsers "Approver_2", Members, Doc
'    If AsTemplate Then
'        If Len(GetProperty(pTemplateVer)) = 0 Then SetProperty pTemplateVer, 1
'        If Len(GetProperty(pTemplateDate)) = 0 Then SetProperty pTemplateDate, GetProperty(pDocLastSave)
''        SetContentControl "Version & Date", "Template Version: " & GetProperty(pDocVer) & " (" & Format(ToServerTime, DateTimeFormat) & ")"
''        SetContentControl "Version & Date", ""
'    Else
'        DocsInfo.Remove Doc.Name
'        DocsInfo.Add LoadDocInfo(Doc), Doc.Name
'        SetProperty pTemplateVer, GetProperty(pTemplateVer, Doc)
'        SetProperty pTemplateDate, GetProperty(pDocLastSave, Doc)
'
'        SetContentControl "DocDate", Format(GetProperty(pDocCreateDate), DateFormat), Doc
'        If Len(GetProperty(pDocDate)) = 0 Then SetProperty pDocDate, GetProperty(pDocLastSave)
'        If DocName = "Scope" Then
'            SetContentControl "CreationDate", Format(ToServerTime, DateFormat), Doc
'            SetContentControl "Version & Date", "Document State: " & GetProperty(pDocState) & _
'                     " | Document Version: Original" & " (" & Format(ToServerTime, DateTimeFormat) & ")"
'        Else
'            SetContentControl "CreationDate", IIf(InStr(DocName, "Reimbursement"), "", "Creation Date: ") & Format(ToServerTime, DateFormat), Doc
'            SetContentControl "Version & Date", "Document State: " & GetProperty(pDocState) & _
'                     " | Document Version: " & GetProperty(pDocVer) & " (" & Format(ToServerTime, DateTimeFormat) & ")"
'        End If
'        FillFirstRevisionDate
'    End If
'    SetContentControl "LastSaveTime", Format(ToServerTime, DateTimeFormat), Doc
''    FillFirstRevisionDate
''    Doc.Bookmarks("DocentTemplateStart").Select
'    Doc.Saved = True
'ex:
'    Set OpenTemplate = Doc
'    If Not AsTemplate And Not Unprotected Then Protect Doc
''    OpeningTemplate = False
'End Function
'Sub SaveDocAsDraft(Doc As Document)
''    MsgBox "Under Construction"
'    SetProperty pDocState, "Private", , msoPropertyTypeString
'End Sub
'Private Function UploadTemplate(Doc As Document, ByVal FName As String, ByVal DocURL As String) As WebResponse
'    SetProperty pTemplateDate, Date + Time
'    SetProperty pTemplateVer, GetProperty(pTemplateVer) + 1, , msoPropertyTypeNumber
'    FName = SaveForUpload(GetFileName(FName, False))
'    If InStr(DocURL, "@@download") Then DocURL = Left$(DocURL, InStr(DocURL, "@@download") - 2)
'    Set UploadTemplate = UpdateAPIFile(FName, DocURL)
'End Function
'Function SaveForUpload(Optional FName As String, Optional Doc As Document) As String
'    If Doc Is Nothing Then Set Doc = ActiveDocument
'    Dim Extension As String
'    If Len(FName) = 0 Then FName = Doc.Name
'    Extension = GetFileExtension(Doc.Name)
'    If Len(Extension) = 0 Then Extension = Replace(GetFileExtension(Doc.AttachedTemplate), "t", "c")
'    If Left$(Doc.Name, 1) <> "_" Then
'        Doc.Saved = False
'        Doc.SaveAs2 Environ("Temp") & "\" & FName & Extension, GetFileFormat(Extension)
'        Doc.SaveAs2 Environ("Temp") & "\_" & FName & Extension, GetFileFormat(Extension)
'    Else
'        FName = Right$(Doc.Name, Len(Doc.Name) - 1)
'    End If
'    SaveForUpload = Environ("Temp") & "\" & FName & IIf(Right$(FName, Len(Extension)) = Extension, "", Extension)
'End Function
''Function UploadTemplate(Doc As Document)
''    Dim DocType As String
''    DocType = GetProperty(pDocType, Doc)
''
''End Function
'Private Function RemoveRedPageBreaks(Doc As Document)
'    Dim Rng As Range
'    Set Rng = Doc.Range
'    With Rng.Find
'        .text = Chr(12)
'        .Font.Color = 255
'        .Wrap = wdFindContinue
'        Do While .Execute
'            Rng.Delete
'        Loop
'    End With
'End Function
'Private Function UpdateDocInfo(Doc As Document, AsTemplate As Boolean, oldState As String, TransitionID As String, DocType As String) As String
'    Dim VNum As Long, NewState As String
'    If AsTemplate Then
'        'TransitionID = ""
'        NewState = "" 'GetInitalState
'        If Not Doc.Saved Then
'            SetContentControl "Version & Date", "Template Version: " & GetProperty(pDocVer) & " (" & Format(ToServerTime, DateTimeFormat) & ")"
'            SetContentControl "DocDate", Format(ToServerTime, DateFormat)
'            VNum = GetProperty(pDocVer) + 1
'            SetProperty pDocVer, VNum, , msoPropertyTypeNumber
'        End If
'    ElseIf Not IsValidDoc(DocType, Doc) Then
'        NewState = "Error"
'    Else
'        If Len(TransitionID) = 0 Then
'            NewState = GetProperty(pDocState)
'            If Len(NewState) = 0 Then NewState = GetInitalState(DocType)
'        Else
'            NewState = GetTransitionDestination(TransitionID, DocType)
'        End If
'        Select Case NewState
'        Case "Published"
''            SetContentControl "DocDate", Format(ToServerTime, DateFormat), Doc
'            If oldState <> NewState Then
'                RemoveRedPageBreaks Doc
'                If DocType = "Meeting Notes" Then
'                    RemoveSpecificCommandButton "Add Proposed Tasks", Doc
'                    'activedocument.InlineShapes(activedocument.InlineShapes.Count).AlternativeText = "Add Proposed Tasks"
'                    UpdateTasks
'                ElseIf DocType = "Meeting Minutes" Then
'                    If ProposedTaskDecided Then
'                        CreateTasks
'                        UpdateTasks
'                    Else
'                        NewState = "Please review all of the proposed Tasks."
'                        GoTo ex
'                    End If
'                End If
'                SetProperty pPublishDate, Format(ToServerTime, DateTimeFormat), Doc
'            End If
'            SetContentControl "Version & Date", " ", Doc
'        Case Else
'            SetContentControl "Version & Date", "Document State: " & NewState & " | Document Version: " & VNum & " (" & Format(ToServerTime, DateTimeFormat) & ")"
'        End Select
'    End If
'ex:
'    UpdateDocInfo = NewState
'End Function
'Function UploadDoc(Doc As Document, Optional ByVal TransitionID As String, _
'        Optional AsTemplate As Boolean, Optional SilentMode As Boolean, Optional NoSpelling As Boolean, _
'        Optional CloseToo As Boolean = True) As Boolean
'    Dim FName As String, DocURL As String, oldState As String, Resp As WebResponse
'    Dim ErMsg As String, ScMsg As String, TempMsg As String, TempLnk As String, FailedCount As Long
'    Dim DocType As String, mDoc As Document, WebLocation As String
'    ', LocalLocation As String, CustomerLocation As String
''    Dim Links(1 To 6) As String, ss() As String, i As Long, Msgs(1 To 6) As String ', Things(1 To 3) As String
'    Dim i As Long, Svd As Boolean, NewState As String
'    Dim ActualDateTime As Date
'    Dim DoUpload As Boolean
'    ReDim ScMsgs(0 To 0) As String
'    ReDim ErMsgs(0 To 0) As String
'    ReDim Links(0 To 0) As String
'    DocType = GetProperty(pDocType, Doc)
'    Svd = Doc.Saved
'    On Error Resume Next
'    If Not SilentMode And Not NoSpelling Then Doc.CheckSpelling
'    Unprotect Doc
'    Unprotect Doc, Application.UserName
'    DoUpload = True
'    On Error GoTo 0
'    oldState = GetProperty(pDocState, Doc)
'
'    DocURL = GetProperty(pDocURL, Doc)
'    If Len(DocURL) = 0 Then
'        Select Case DocType
'        Case "Scope", "RFP", "PMP"
'            DocURL = Get1DocURL(DocType)
'            SetProperty pDocURL, DocURL, Doc
'        End Select
'    End If
'
'    Application.StatusBar = "Uploading..."
'    NewState = UpdateDocInfo(Doc, AsTemplate, oldState, TransitionID, DocType)
'    'NextTransitions = Split(GetDocProperty("docentDocNextState"), Seperator)
'    'Doc.Save
'    If AsTemplate Then
'        NewState = ""
'        Protect Doc
'        FName = Doc.FullName
'        If GetFileName(FName, False) = "Main Template" Then
'            If Application.UserName = "Abdallah Ali" Then DoUpload = MsgBox("Upload?", vbYesNo, "") = vbYes
'            ProgressBar.Reset
'            ProgressBar.HideApplication = True
'            ProgressBar.BarsColor CLng(ProjectColorStr)
'            ProgressBar.Dom(1) = UBound(templateName)
'            If ProgressBar.Progress(0, "Updating Main Template") Or Set_Cancelled Then GoTo ex
'            ProgressBar.Caption = "Templates Updating Progress"
'            ProgressBar.Show
'            If DoUpload Then
'                Set Resp = UploadTemplate(Doc, FName, DocURL)
'                If Not IsGoodResponse(Resp) Then
'                    ErMsg = ErMsg & Chr(10) & "Main Template"
'                Else
'                    Doc.Close False
'                    For i = 1 To UBound(templateName)
'                        If templateName(i) <> "Main Template" Then
'                            If ProgressBar.Progress(, "Updating " & templateName(i) & " Template") Or Set_Cancelled Then GoTo ex
'                            Set mDoc = OpenTemplate(templateName(i), True, True)
'                            If DoUpload Then UpdateTemplate mDoc, FName
'                            If Not UploadDoc(mDoc, "", True, True) Then 'GetInitalState(TemplateName(i))
'                                FailedCount = FailedCount + 1
'                                ErMsg = ErMsg & Chr(10) & templateName(i)
'                            End If
'                        End If
'                    Next
'                    ProgressBar.Progress
'                End If
'            End If
''            Set Doc = Documents.Open(FName)
'        Else
'            Set Resp = UploadTemplate(Doc, FName, DocURL)
'        End If
'    Else
'        If Not IsValidDoc(DocType, Doc) Then
'            ReDim Preserve ErMsgs(0 To UBound(ErMsgs) + 1) As String
'            ErMsgs(0) = "Invalid document"
'            ErMsgs(UBound(ErMsgs)) = "     The required fields are not filled"
'            GoTo ex
'        End If
'
'        FName = GenDocName(NewState)
'        If Len(FName) = 0 Then
'            ErMsg = Chr(10) & "     The naming convention is missing"
'            GoTo ex
'        End If
'
'        Protect Doc
'        FName = SaveForUpload(FName)
'
'        TempLnk = SaveTo("Local", Doc, DocType, NewState, True)
'        If Len(TempLnk) Then TempMsg = SaveTo("Local", Doc, DocType, NewState)
'        If InStr(TempMsg, " (") > 0 Then
'            ReDim Preserve ErMsgs(0 To UBound(ErMsgs) + 1) As String
'            ErMsgs(UBound(ErMsgs)) = TempMsg
'            ErMsg = ErMsg & Chr(10) & TempMsg
'        Else
'            ReDim Preserve ScMsgs(0 To UBound(ScMsgs) + 1) As String
'            ScMsgs(UBound(ScMsgs)) = TempMsg
'            ReDim Preserve Links(0 To UBound(Links) + 1) As String
'            Links(UBound(ScMsgs)) = TempLnk
'            ScMsg = ScMsg & Chr(10) & TempMsg
'        End If
'
'        TempLnk = SaveTo("Customer", Doc, DocType, NewState, True)
'        If Len(TempLnk) Then TempMsg = SaveTo("Customer", Doc, DocType, NewState)
'        If InStr(TempMsg, " (") > 0 Then
'            ReDim Preserve ErMsgs(0 To UBound(ErMsgs) + 1) As String
'            ErMsgs(UBound(ErMsgs)) = TempMsg
'            ErMsg = ErMsg & Chr(10) & TempMsg
'        Else
'            ReDim Preserve ScMsgs(0 To UBound(ScMsgs) + 1) As String
'            ScMsgs(UBound(ScMsgs)) = TempMsg
'            ReDim Preserve Links(0 To UBound(Links) + 1) As String
'            Links(UBound(ScMsgs)) = TempLnk
'            ScMsg = ScMsg & Chr(10) & TempMsg
'        End If
'        Select Case DocType
'        Case "Scope"
'            If NewState = "" Then
'                If MsgBox("Do you want to remove page breaks between tasks?", vbQuestion + vbYesNo, "") = vbYes Then
'                    RemoveScopeTasksPBreaks Doc
'                End If
'            End If
'        Case "Tasks"
'            UploadTasks
'        End Select
'        On Error Resume Next
'        ActualDateTime = AlreadyServerTime(CStr(DateValue(GetContentControl("ActualDate")) + _
'                                                TimeValue(GetContentControl("ActualTime"))))
'        On Error GoTo ex
'        If Len(DocURL) Then
'            WebLocation = DocURL
'            Select Case DocType
''            Case "Meeting Agenda"
''                set resp = UpdateAPIMeetingAgenda(FName, _
''                        GetContentControl("MeetingLocation"), _
''                        GetContentControl("MeetingSubject"), _
''                        GetContentControl("Meeting Date"), _
''                        GetContentControl("MeetingTime"), _
''                        DocURL)
'            Case "Scope"
'                Set Resp = UpdateAPIContent(DocURL, Array("file", "table"), Array(FName, GetMeetingSummaryTable))
''                ScopeURL = Get1DocURL("scope")
''                RefreshRibbon
'            Case Else
'                Set Resp = UpdateAPIContent(DocURL, Array("file", "planned_action_items", "proposed_action_items", _
'                            "actuals"), _
'                            Array(FName, GetProperty(pPlannedTasks), GetProperty(pProposedTasks), _
'                            GetActuals(Doc)))
'            End Select
'            LockAPIFile DocURL, True
'        Else
'            Select Case DocType
''            Case "Meeting Agenda"
''                set resp = CreateAPIMeetingAgenda(FName, _
''                        GetContentControl("MeetingLocation"), _
''                        GetContentControl("MeetingSubject"), _
''                        GetContentControl("Meeting Date"), _
''                        GetContentControl("MeetingTime"), _
''                        DefaultDocumentsFolder)
'            Case "Scope"
'                WebLocation = DefaultScopeFolder
'                Set Resp = CreateAPIContent(DocType, DefaultScopeFolder, _
'                        Array("@type", "file", "table"), _
'                        Array(DocType, FName, GetMeetingSummaryTable))
'                DocURL = Resp.Data("@id") 'Get1DocURL("scope") '
'                ScopeURL = DocURL
'                ProjectInfo.Remove "ScopeURL"
'                ProjectInfo.Add "ScopeURL", ScopeURL
'                SaveProjectInfoToReg
'                RefreshRibbon
''                Set Resp = CreateAPIScope(FName, _
'                        GetMeetingSummaryTable, _
'                        DefaultDocumentsFolder)
'            Case Else
'                If DocType Like "meeting*" Then
'                    WebLocation = DefaultDocumentsFolder
'                    Set Resp = CreateAPIContent(DocType, DefaultDocumentsFolder, _
'                                Array("@type", "file", "planned_action_items", "proposed_action_items", _
'                                "actuals"), _
'                                Array(DocType, FName, GetProperty(pPlannedTasks), GetProperty(pProposedTasks), _
'                                GetActuals(Doc)))
'                Else
'                    WebLocation = DefaultDocumentsFolder
'                    Set Resp = CreateAPIContent("docent_misc_document", DefaultDocumentsFolder, Array("document_type", "file"), Array(DocType, FName))
'                End If
'                DocURL = Resp.Data("@id")
'            End Select
'        End If
'        If Len(DocURL) And Not AsTemplate And oldState <> NewState Then
'            UpdateAPIFileWorkflow DocURL, TransitionID 'NewState
'            SetProperty pDocState, StrConv(NewState, vbProperCase), , msoPropertyTypeString
'        End If
'    End If
'    'If Len(ErMsg) > 0 Then ErMsg = NotSavedWeb & Chr(10) & ErMsg
'ex:
'    If IsGoodResponse(Resp) Then
'        ScMsg = ScMsg & Chr(10) & "     Project Website"
'        ReDim Preserve ScMsgs(0 To UBound(ScMsgs) + 1) As String
'        ScMsgs(UBound(ScMsgs)) = "     Project Website"
'        ReDim Preserve Links(0 To UBound(Links) + 1) As String
'        Links(UBound(ScMsgs)) = WebLocation
'    Else
'        On Error Resume Next
'        TempMsg = NotSavedWeb & " (" & Resp.StatusDescription & ")"
'        If Err.Number Then TempMsg = NotSavedWeb
'        ErMsg = TempMsg & Chr(10) & ErMsg
'        ReDim Preserve ErMsgs(0 To UBound(ErMsgs) + 1) As String
'        ErMsgs(UBound(ErMsgs)) = TempMsg
'    End If
'    ErMsg = CleanErMsg(ErMsg)
'    ScMsg = CleanErMsg(ScMsg)
'
''    if len(ErMsg) then If instr(ErMsg,NotSavedWeb)  = 0 then
''    Stop
''    Resume
'    On Error GoTo -1
'    On Error Resume Next
'    Application.StatusBar = ""
'    Protect Doc
'    UploadDoc = Len(ErMsg) = 0  'And Len(Resp) > 0'IsGoodResponse(Resp) And
'    If SilentMode Then
'        If CloseToo Then Doc.Close False
'    Else
'        Unload ProgressBar
'        Dim Msgs(1 To 8) As String, Lnks(1 To 6) As String, Clrs(1 To 6)
'        If UploadDoc Then 'IsGoodResponse(Resp) And Len(ErMsg) = 0 And Len(Resp) > 0 Then
'            If DocType = "Scope" Then ScopeURL = Resp.Data("@id"): DownloadProjectInfo
'            If DocType = "RFP" Then RFPURL = Resp.Data("@id"): DownloadProjectInfo
'            If DocType = "PMP" Then PMPURL = Resp.Data("@id"): DownloadProjectInfo
'            Msgs(1) = DocType & " Saving Summary" & Chr(10) & Chr(10) & "  Uploaded Successfully to these locations:"
'            Clrs(1) = 0
'            Lnks(1) = ""
'            For i = 1 To UBound(ScMsgs)
'                Msgs(1 + i) = ScMsgs(i)
'                Clrs(1 + i) = 4496968
'                Lnks(1 + i) = Links(i)
'            Next
'            'Msgs(UBound(Msgs)) = RemainsOpenMsg
'            Application.Visible = True
'            frmMsgBox.Display Msgs, , , "Docent IMS", Clrs, , Lnks
'                'Array(DocType & " Saving Summary" & Chr(10) & Chr(10) & "Uploaded Successfully to these locations:", _
'                    ScMsg), _
'                    , , "Docent IMS", Clrs:=Array(0, 4496968)
'            Doc.Close False
'        ElseIf FailedCount Then
'            frmMsgBox.Display NotSavedTemplates & Chr(10) & ErMsg, , Critical, "Docent IMS"
''        ElseIf Len(ErMsg) = 0 Then
''            frmMsgBox.Display NotSavedWeb & Chr(10) & Resp.StatusDescription, , Critical, ""
'        Else
'            LoadDocInfo Doc
'            If Len(DocURL) Then Set NextTransitions = GetAPIFileWorkflowTransitions(DocURL)
'            RefreshRibbon
''            If InStr(ErMsg, NotSavedLocal) Then
''                If ErMsg Like NotSavedLocal & "*" Then
''                    frmMsgBox.Display ErMsg, , Critical, ""
''                Else
''
''                End If
''            Else
'            If Len(ScMsg) Then
'                Msgs(1) = DocType & " Saving Summary" & Chr(10) & Chr(10) & "  Uploaded Successfully to these locations:"
'                Clrs(1) = 0
'                Lnks(1) = ""
'                For i = 1 To UBound(ScMsgs)
'                    Msgs(1 + i) = ScMsgs(i)
'                    Clrs(1 + i) = 4496968
'                    Lnks(1 + i) = Links(i)
'                Next
'                Dim x As Long
'                x = i + 1
'                Msgs(x) = Chr(10) & Chr(10) & "  " & IIf(Len(ErMsgs(0)) = 0, "Not saved to these locations:", ErMsgs(0))
'                Clrs(x) = 0
'                Lnks(x) = ""
'                For i = 1 To UBound(ErMsgs)
'                    Msgs(x + i) = ErMsgs(i)
'                    Clrs(x + i) = vbRed
'                    Lnks(x + i) = ""
'                Next
'
'                Msgs(UBound(Msgs)) = RemainsOpenMsg
'                frmMsgBox.Display Msgs, , Exclamation, "Docent IMS", Clrs, , Lnks
''                frmMsgBox.Display Array(DocType & " Saving Summary" & Chr(10) & Chr(10) & "Uploaded Successfully to these locations:", _
''                            ScMsg, _
''                            Chr(10) & "Not saved to these locations:", ErMsg), Array("OK", "Retry"), Exclamation, "Docent IMS", Array(0, 4496968, 0, vbRed)
'            Else
'                Msgs(1) = DocType & " Saving Summary" & Chr(10) & Chr(10) & "  " & IIf(Len(ErMsgs(0)) = 0, "Not saved to these locations:", ErMsgs(0))
'                Clrs(1) = 0
'                Lnks(1) = ""
'                For i = 1 To UBound(ErMsgs)
'                    Msgs(1 + i) = ErMsgs(i)
'                    Clrs(1 + i) = vbRed
'                    Lnks(1 + i) = ""
'                Next
'                Msgs(UBound(Msgs)) = RemainsOpenMsg
'                frmMsgBox.Display Msgs, , Critical, "Docent IMS", Clrs, , Lnks
''                frmMsgBox.Display Array(DocType & " Saving Summary" & Chr(10) & Chr(10) & "Not saved to these locations:", ErMsg), _
''                        , Critical, "Docent IMS", Array(0, vbRed)
'            End If
''                frmMsgBox.Display ErMsg, , Critical, ""
''            End If
'        End If
'    End If
'End Function
'Private Function CleanErMsg(ByVal ErMsg As String) As String
'    If Len(Replace(ErMsg, Chr(10), "")) = 0 Then
'        ErMsg = ""
'    ElseIf Len(ErMsg) > 0 Then
'        Do While Left$(ErMsg, 1) = Chr(10)
'            ErMsg = Right$(ErMsg, Len(ErMsg) - 1)
'            If Len(ErMsg) = 0 Then Exit Do
'        Loop
'    End If
'    CleanErMsg = ErMsg
'End Function
'Private Function SaveTo(LocationName As String, Doc As Document, DocType As String, DocState As String, Optional CheckMode As Boolean) As String
'    Dim Locs As Dictionary, FPath As String, errMsg As String
'    'If Application.UserName = "Abdallah Ali" Then Exit Function
'    Set Locs = GetLocations
'    On Error Resume Next
'    FPath = Locs(DocType)("States")(DocState)(LocationName)
'    If Locs Is Nothing Then
'        errMsg = "Missing Settings"
'    ElseIf Locs("Testing") = True Then
'        errMsg = "Testing Mode"
'    ElseIf FPath = "Missing" Then
'        errMsg = "Not Set"
'    ElseIf FPath = "False" Then
'        errMsg = "Skipped for this state"
'    Else
'        FPath = GenPathName(CStr(Locs(DocType)(MainInfo("company"))(LocationName)), DocState, DocType)
'        If Len(FPath) = 0 Then
'            errMsg = "Location Not Set"
'        Else
'            'CreateDir FPath
'            If Not IsWritable(FPath) Then
'                errMsg = IIf(CheckMode, "", "Inaccessable")
'            ElseIf CheckMode Then
'                'ErrMsg = FPath
'            Else
'                Err.Clear
'                Doc.SaveAs2 FPath & Right$(Doc.Name, Len(Doc.Name) - 1)
'                If Err.Number Then Err.Clear: errMsg = "Inaccessable"
'            End If
'        End If
'    End If
'    If CheckMode Then
'        If Len(errMsg) Then
'            WriteLog 3, CurrentMod, "SaveTo (Check Mode)", errMsg
'            SaveTo = errMsg
'        Else
'            WriteLog 1, CurrentMod, "SaveTo (Check Mode)"
'            SaveTo = FPath
'        End If
'    Else
'        If Len(errMsg) Then
'            errMsg = " (" & errMsg & ")"
'            WriteLog 3, CurrentMod, "SaveTo", errMsg
'        Else
'            WriteLog 1, CurrentMod, "SaveTo"
'        End If
'        Select Case LocationName
'        Case "Local": errMsg = "     Local Network" & errMsg
'        Case "Customer": errMsg = "     Customer Drive" & errMsg
'        Case "Web": errMsg = "     Project Website" & errMsg
'        End Select
'        SaveTo = errMsg
'    End If
'End Function
''Private Function SaveToOtherLocations(Doc As Document, DocType As String, DocState As String) As String
''    Dim Locs As Dictionary, FPath As String
''    Set Locs = GetLocations
''    If Application.UserName = "Abdallah Ali" Then Exit Function
''    If Locs("Testing") = True Then Exit Function
''    On Error Resume Next
''    FPath = Locs(DocType)("States")(DocState)("Local")
''    If Len(FPath) Then
''        FPath = GenPathName(CStr(Locs(DocType)(MainInfo("company"))("Local")), DocState, DocType)
''        CreateDir FPath
''        'FPath = Replace(FPath, "\\", "\")
''        Doc.SaveAs2 FPath & Right$(Doc.Name, Len(Doc.Name) - 1)
''        If Err.Number Then
''            Err.Clear
''            SaveToOtherLocations = "Local Network (Inaccessable)"
''            WriteLog 3, CurrentMod, "SaveToOtherLocations", "Local Network (Inaccessable)"
'''            MsgBox "Could not save to Local Network", vbExclamation, "Docent IMS"
''        End If
''    Else
''        SaveToOtherLocations = "Local Network (Not Set)"
''        WriteLog 3, CurrentMod, "SaveToOtherLocations", "Local Network (Not Set)"
''    End If
''    FPath = ""
''    FPath = Locs(DocType)("States")(DocState)("Customer")
''    If Len(FPath) Then
''        FPath = GenPathName(CStr(Locs(DocType)(MainInfo("company"))("Customer")), DocState, DocType)
''        CreateDir FPath
''        Doc.SaveAs2 FPath & Doc.Name
''        If Err.Number Then
''            SaveToOtherLocations = SaveToOtherLocations & IIf(Len(SaveToOtherLocations), Chr(10), "") & "Customer Drive (Inaccessable)"
''            WriteLog 3, CurrentMod, "SaveToOtherLocations", "Customer Drive (Inaccessable)"
'''            MsgBox "Could not save to Customer Drive", vbExclamation, "Docent IMS"
''        End If
''    Else
''        SaveToOtherLocations = SaveToOtherLocations & IIf(Len(SaveToOtherLocations), Chr(10), "") & "Customer Drive (Not Set)"
''        WriteLog 3, CurrentMod, "SaveToOtherLocations", "Customer Drive (Not Set)"
''    End If
''    If Len(SaveToOtherLocations) Then
''        SaveToOtherLocations = NotSavedLocal & Chr(10) & Chr(10) & SaveToOtherLocations
''    Else
''        WriteLog 1, CurrentMod, "SaveToOtherLocations"
''    End If
''End Function
'Private Function GenPathName(Pth As String, NewState As String, DocType As String) As String
'    GenPathName = Pth
'    GenPathName = Replace(GenPathName, "%Document State%", NewState)
'    GenPathName = Replace(GenPathName, "%Contract Number%", ContractNumberStr)
'    GenPathName = Replace(GenPathName, "%Date%", Replace(Format(ToServerTime, DateFormat), "/", "-"))
'    GenPathName = Replace(GenPathName, "%Project Name%", ProjectNameStr)
'    GenPathName = Replace(GenPathName, "%Documents Type%", DocType)
'    GenPathName = Replace(GenPathName, "%User Name%", Application.UserName)
'    If Right$(GenPathName, 1) <> "\" Then GenPathName = GenPathName & "\"
'End Function
'Private Function GenDocName(NewState As String) As String
'    GenDocName = DocumentsNameConvStr
'    GenDocName = Replace(GenDocName, "ContractNumber", ContractNumberStr)
'    GenDocName = Replace(GenDocName, "DocState", NewState)
'    GenDocName = Replace(GenDocName, "Doctype", GetProperty(pDocType))
'    GenDocName = Replace(GenDocName, "DocDate", Replace(Format(ToServerTime, DateFormat), "/", "-"))
'    GenDocName = Replace(GenDocName, "DocTime", Replace(Format(ToServerTime, TimeFormat), ":", "-"))
'    GenDocName = Replace(GenDocName, "PrjName", GetProperty(pPName))
'End Function
''Private Function GetTemplatePath() As String
''    Dim DocName As String
''    DocName =
''    GetTemplatePath =
''
''End Function
'
'Function IsValidDoc(DocType As String, Doc As Document) As Boolean
'    IsValidDoc = ValidateAll(Doc)
'    If IsValidDoc Then
'        Select Case DocType
'        Case "Meeting Notes": IsValidDoc = IsValidMeetingNotes(Doc)
'        Case "Reimbursement": IsValidDoc = IsValidReimbursement(Doc)
'        End Select
'    End If
'End Function
'Private Function ValidateAll(Doc As Document) As Boolean
'    Dim Tbl As Table, r As Long, n As Long, s As String, clr As Long
'    ValidateAll = True
'    For Each Tbl In Doc.Tables
'        For r = 1 To Tbl.Rows.Count
'            On Error Resume Next
'            clr = Tbl.Cell(r, 1).Range.Font.Color
'            For n = 2 To Tbl.Columns.Count
'                On Error Resume Next
'                s = CellText(Tbl.Cell(r, n))
'                If Err.Number Then
'                    Err.Clear
'                ElseIf s = CCDDPlaceholder Then
'                    ValidateAll = False
'                    Tbl.Cell(r, n).Range.Font.Bold = True
'                    Tbl.Cell(r, n).Range.Font.Color = vbRed
'                ElseIf Tbl.Cell(r, n).Range.Font.Color = vbRed Then
'                    Tbl.Cell(r, n).Range.Font.Bold = False
'                    Tbl.Cell(r, n).Range.Font.Color = clr
'                End If
'            Next
'        Next
'    Next
'    Exit Function
'ex:
''    Stop
''    Resume
'End Function
