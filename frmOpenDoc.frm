VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOpenDoc 
   Caption         =   "Open Document"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "frmOpenDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOpenDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private mColl As Collection, mCreateMode As Boolean, mIsMeeting As Boolean, mDocType As String, mMeetingMode As Boolean

Sub ShowDocuments(DocsColl As Collection, DocType As String, CreateMode As Boolean)
'    If IsEmpty(Arr) Then Exit Sub
    mDocType = DocType '= DocumentName(DocNum)
    On Error Resume Next
    If DocsColl.Count = 0 Then
        If CreateMode Then
            frmMsgBox.Display "If a Meeting Document was selected and " & _
                            "you see this message, then there are no published Meetings " & _
                            "without existing " & DocType & "." & Chr(10) & Chr(10) & _
                            "To Edit an existing " & DocType & _
                            ", close this window and in the ""Document Manager"", " & _
                            "use the ""Open Document"" option and select the document to edit.", _
                             Array(), Exclamation, "DocentIMS"
        Else
            frmMsgBox.Display "The ""Open"" action opens all existing " & _
                        DocType & " for editing.  However, there are no " & _
                        DocType & " available for editing.", Array(), Exclamation, "DocentIMS"
        End If
'        MsgBox "No """ & DocType & """ documents to open.", vbCritical, "DocentIMS"
        Exit Sub
    End If
    Set mColl = DocsColl
'    mArr = Arr
    If CreateMode Then
        Caption = "Create " & DocType
        lbState.Visible = False
        sbStates.ColumnCount = 1
        lbTitle.Width = sbStates.Width - 1
        btnAdd.Caption = "Create"
    Else
        Caption = "Open " & DocType
    End If
    mCreateMode = CreateMode
    mMeetingMode = DocType = "meeting"
    mIsMeeting = InStr(DocType, "meeting") > 0
    If mIsMeeting Then lbTitle.Caption = "Meeting Title"
    sbStates.List = GetDocsArr
    Me.Show
End Sub
Private Function GetDocsArr()
    On Error Resume Next
    Dim i As Long
    ReDim Arr(1 To mColl.Count, 1 To 3)
    For i = 1 To mColl.Count
        Arr(i, 1) = mColl(i)("ParentMeeting")("MeetingShortName")
        If Len(Arr(i, 1)) = 0 Then Arr(i, 1) = mColl(i)("title")
        Arr(i, 2) = mColl(i)("State")
        If Len(Arr(i, 2)) = 0 Then Arr(i, 2) = StrConv(mColl(i)("review_state"), vbProperCase)
        Arr(i, 3) = mColl(i)("@id")
    Next
    GetDocsArr = Arr
End Function
Private Sub btnAdd_Click()
    Dim xColl As Dictionary
    If sbStates.ListIndex = -1 Then Exit Sub
    Set xColl = mColl(sbStates.ListIndex + 1)
    Dim URL As String, DocURL As String, DocState As String, Doc As Document, ParentDoc As Document ', MeetingTitle As String
    Dim MtngType As String
'    Dim PlannedTask As String, ProposedTask As String
'    Dim MeetingStart As String ', MeetingLocation As String
'    Dim ActualMeetingStart As String, PrpMgr As String, ActualPrpMgr As String
'    Dim MNURL As String, MNState As String, MNAtt As String, MNExAtt As String
    If sbStates.ListIndex < 0 Then Exit Sub
    Application.ScreenUpdating = False
    Me.Hide
    On Error Resume Next
    DocURL = xColl("@id")
    URL = Replace(DocURL, ProjectURLStr, "")
'    PlannedTask = xColl("planned_action_items")
'    ProposedTask = xColl("proposed_action_items")
    xColl.Remove "ParentMeeting"
    xColl.Add "ParentMeeting", GetParentMeetingObject(GetParentDir(DocURL))
    '"start", "end", "location", "attendees", "attendees_group", "will_property_manager_attend"
'    MeetingTitle = xColl("ParentMeeting")("title")
    MtngType = xColl("ParentMeeting")("meeting_type")("title")
'    MeetingStart = xColl("ParentMeeting")("MeetingDateTime")
'    MeetingLocation = xColl("ParentMeeting")("location")("title")
'    PrpMgr = xColl("ParentMeeting")("will_property_manager_attend")
    
'    Set MNActuals = ParseJson(xColl("ParentDoc")("actuals"))
    
'    ActualPrpMgr = xColl("ParentDoc")("did_property_manager_attend")
'    ActualMeetingStart = xColl("ParentDoc")("actual_meeting_date_time")
'    MNAtt = xColl("ParentDoc")("which_board_members_attended")
'    MNExAtt = xColl("ParentDoc")("external_attendees")
    Set NextTransitions = GetAPIFileWorkflowTransitions(URL)
'    Set OpeningDocInfo = New DocInfo
'    With OpeningDocInfo
'        '.
'    End With
    If mCreateMode Then
        DocState = GetInitalState(mDocType)
        Select Case mDocType
        Case "Meeting Notes"
            Set Doc = OpenTemplate(mDocType, , , mIsMeeting, MtngType)  ', , , StrConv(CStr(xcoll( 2)), vbProperCase), URL
            Set ParentDoc = OpenDocumentAt(xColl("ParentDoc")("@id"), xColl("ParentDoc")("State"))
            CopyAgendaItems Doc, ParentDoc
            Set ParentDoc = Nothing
        Case "Meeting Minutes"
            Set Doc = OpenTemplate(mDocType, , , mIsMeeting, MtngType, Unprotected:=True) ', , , StrConv(CStr(xcoll( 2)), vbProperCase), URL
            Set ParentDoc = OpenDocumentAt(xColl("ParentDoc")("@id"), xColl("ParentDoc")("State"))
        Case Else
            Set Doc = OpenTemplate(mDocType, , , mIsMeeting, MtngType)  ', , , StrConv(CStr(xcoll( 2)), vbProperCase), URL
        End Select
    ElseIf mMeetingMode Then
'        Me.Hide
        frmCreateMeeting.OpenMeeting DocURL
    Else
'        DocState = xcoll("State") 'CStr(xcoll(2))
        Set Doc = OpenDocumentAt(DocURL, xColl("State"))
    End If
    If mIsMeeting Then
        UpdateMeetingFile Doc, xColl ' DocURL, DocState, MeetingTitle, _
                PlannedTask, ProposedTask, MeetingStart, ActualMeetingStart, _
                PrpMgr, ActualPrpMgr, MNAtt, MNExAtt, _
                MeetingLocation
    End If
    If Not ParentDoc Is Nothing Then
'        ParentDoc.Windows(1).Activate
        ParentDoc.Windows.CompareSideBySideWith Doc
        ParentDoc.Windows.SyncScrollingSideBySide = False
        ParentDoc.Windows.ResetPositionsSideBySide
    End If
    Application.ScreenUpdating = True
    Unload Me
End Sub
Private Sub sbStates_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnAdd_Click
End Sub
Private Sub UserForm_Initialize()
    CenterUserform Me
End Sub
