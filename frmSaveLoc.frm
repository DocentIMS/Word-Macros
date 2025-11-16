VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaveLoc 
   Caption         =   "Docent File Naming & Saving Setup"
   ClientHeight    =   11895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "frmSaveLoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaveLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private OldLocations As Dictionary
Private NewLocations As Dictionary
Private LastDocType As String
Private LastPName As String
'Private SaveLocTestingMode As Boolean
Private DisableEvents As Boolean
Private ss() As String
Private Paths As Dictionary
Private Evs As New CtrlEvents
Private ChangesMade As Boolean
Private OriginalTestMode As Boolean
Private Const PHBrowse = "Set on the project's site" '"Click to select the location"
Private Const DDPlaceHolder = "-- Choose"
Private Const Configured = "Location Set"
Private Sub btnCancel_Click(): Unload Me: End Sub
Private Function IsProjectSelected() As Boolean
    If LastPName Like DDPlaceHolder & "*" Then Exit Function
    If LastPName Like "" Then Exit Function
    IsProjectSelected = True
End Function
Private Function IsDocumentSelected() As Boolean
    If LastDocType Like DDPlaceHolder & "*" Then Exit Function
    If LastDocType Like "" Then Exit Function
    IsDocumentSelected = True
End Function
Private Sub btnReset_Click()
    ClearAll 2
End Sub

Private Sub btnSaveDoc_Click()
    SaveLastChoices
    If SaveConfig Then
'        ClearAll 1
        cbProject_Change
    End If
End Sub
Private Sub lbHelpLink_Click(): GoToLink lbHelpLink.Tag: End Sub
Private Sub SaveLastChoices()
    If Not IsProjectSelected Then Exit Sub
    If Not IsDocumentSelected Then Exit Sub
    
    If DisableEvents Then Exit Sub
    Set NewLocations = OldLocations
    Dim DocLocs As New Dictionary, docStates As New Dictionary, StatesChoices As Dictionary
    Dim DocCompLoc As Dictionary
    Dim i As Long, j As Long, IsMissed() As Boolean
    'Choices
    ReDim IsMissed(LBound(ss) To UBound(ss)) As Boolean
    For i = LBound(ss) To UBound(ss)
        IsMissed(i) = True
        For j = 1 To lvStates.ListItems.Count
'            IsMissed(i) = IsMissed(i) Or Controls("lv" & ss(i)).ListItems(j).Checked
            If Controls("lv" & ss(i)).ListItems(j).Checked Then
                IsMissed(i) = False
                Exit For
            End If
        Next
    Next
    For i = 1 To lvStates.ListItems.Count
        Set StatesChoices = New Dictionary
        For j = LBound(ss) To UBound(ss)
            If IsMissed(j) Then
                StatesChoices.Add ss(j), "Missing"
            Else
                StatesChoices.Add ss(j), Controls("lv" & ss(j)).ListItems(i).Checked
            End If
        Next
        docStates.Add lvStates.ListItems(i).text, StatesChoices
    Next
    DocLocs.Add "States", docStates
    'Paths
'    Set DocCompLoc = New Dictionary
'    For j = LBound(ss) To UBound(ss)
'        DocCompLoc.Add ss(j), Paths(ss(i)) 'Controls("tb" & ss(j)).Value
'    Next
    DocLocs.Add MainInfo("company"), Paths ' DocCompLoc
    
    
    On Error Resume Next
    If NewLocations Is Nothing Then
        Set NewLocations = New Dictionary
    Else
        For i = 1 To NewLocations(LastDocType).Count
            If Not DocLocs.Exists(NewLocations(LastDocType).KeyName(i)) Then
                DocLocs.Add NewLocations(LastDocType).KeyName(i), NewLocations(LastDocType)(i)
            Else
                
            End If
        Next
    End If
    NewLocations.Remove LastDocType
    NewLocations.Add LastDocType, DocLocs
'    NewLocations.Remove "Testing"
'    NewLocations.Add "Testing", GetLocationsTestMode
    SetLocations LastPName, NewLocations
End Sub
Private Sub ClearAll(Optional Level As Long = 0)
    Dim i As Long
    If Level >= 2 Then
        cbProject.Enabled = True
        Do While cbProject.ListCount > 1: cbProject.RemoveItem 1: Loop
'        cbProject.Clear
'        cbProject.AddItem "-- Choose Project --"
        On Error Resume Next
        i = UBound(ProjectName)
        On Error GoTo 0
        If i = 0 Then RefreshRibbon
        For i = 1 To UBound(ProjectName)
            cbProject.AddItem ProjectName(i)
        Next
        cbProject.ListIndex = 0
        Exit Sub
    End If
    If Level >= 1 Then
        cbDocumentType.Enabled = True
        cbDocumentType.Clear
        cbDocumentType.AddItem "-- Choose Document Type --"
        cbDocumentType.ListIndex = 0
        Exit Sub
    End If
    For i = LBound(ss) To UBound(ss)
        Controls("lv" & ss(i)).ListItems.Clear
'        Controls("tb" & ss(i)).Value = ""
    Next
    lvStates.ListItems.Clear
    tbNamingConv.value = ""
    Set Paths = New Dictionary
'    ckTest.Value = False
End Sub
Private Function GetWebLoc(DocType As String)
    If InStr(DocType, "Meeting") Then
        GetWebLoc = ProjectURLStr & "/meetings/{MeetingID}/" '{DocumentNamingConvention}"
    Else
        GetWebLoc = ProjectURLStr & "/documents/" '{DocumentNamingConvention}"
    End If
End Function
Private Sub LoadChoices()
    Dim i As Long, j As Long
    Dim docStates As Dictionary, DocLocs As Dictionary
'    If DisableEvents Then Exit Sub
'    If cbProject.Value Like DDPlaceHolder & "*" Then Exit Sub
'    If cbProject.Value Like "" Then Exit Sub
'    If cbDocumentType.Value Like DDPlaceHolder & "*" Then Exit Sub
'    If cbDocumentType.Value Like "" Then Exit Sub
    On Error Resume Next
    If NewLocations Is Nothing Then
        Set OldLocations = GetLocations(cbProject.value)
    Else
        Set OldLocations = NewLocations    'GetLocations(cbProject.Value)
    End If
'    tbWeb.Value = GetWebLoc(cbDocumentType.Value)
    If (Not cbProject.value Like DDPlaceHolder & "*") And (cbProject.value <> "") Then tbNamingConv.value = DocumentsNameConvStr
    'States
    Set docStates = GetStatesOfDoc(cbDocumentType.value)
    For i = 1 To docStates.Count
        lvStates.ListItems.Add , , docStates(i)
        For j = LBound(ss) To UBound(ss)
            Controls("lv" & ss(j)).ListItems.Add
            If docStates(i) = "Private" And j > LBound(ss) Then
'                Me.lvCustomer.ListItems(1).ListSubItems(0)..SubItems '= True '.Ghosted = True
                Controls("lv" & ss(j)).ListItems(i).text = "  Not Allowed"
                Controls("lv" & ss(j)).ListItems(i).ForeColor = vbRed
'                Controls("lv" & ss(j)).ListItems(i).Ghosted = True '.ForeColor = vbBlack '
            End If
        Next
        Controls("lv" & ss(LBound(ss))).ListItems(i).Checked = True
    Next
    'Testing
'    SaveLocTestingMode = SaveLocTestingMode Or (OldLocations("Testing") = True)
    ckTest.value = GetLocationsTestMode 'SaveLocTestingMode
    btnSaveExit.Enabled = GetLocationsTestMode 'SaveLocTestingMode
    'Locations
    Set DocLocs = OldLocations(cbDocumentType.value)
    Set Paths = New Dictionary
    For i = LBound(ss) + 1 To UBound(ss)
        Paths.Add ss(i), DocLocs(MainInfo("company"))(ss(i))
        CheckPath ss(i)
'        Controls("tb" & ss(i)).Value = DocLocs(MainInfo("company"))(ss(i))
    Next
    'States Choices
    'Set DocLocs = Nothing
    If Not docStates Is Nothing Then
        Set DocLocs = DocLocs("States")
        For i = 1 To docStates.Count
            If DocLocs.Exists(docStates(i)) Then
                For j = LBound(ss) + 1 To UBound(ss)
                    Controls("lv" & ss(j)).ListItems(i).Checked = DocLocs(i)(ss(j))
                Next
            End If
        Next
    End If
'    UpdateEnabledButtons
End Sub
Private Function SaveConfig() As Boolean
    Dim i As Long, MsgOkStr As String, MsgBadStr As String, MsgType As NewMsgBoxStyle
    If GetLocationsTestMode Then 'SaveLocTestingMode Then
        MsgOkStr = "Documents Configurations:" & Chr(10) & "Testing Mode"
        MsgType = Success
        SaveConfig = True
    ElseIf IsFullyConfigured Then
ConsiderComplete:
        If UploadSaveLocJSON(LastPName) Then
            frmMsgBox.Display cbDocumentType.value & " Configurations were uploaded successfully", , Success, "DocentIMS"
            SaveConfig = True
        Else
            frmMsgBox.Display cbDocumentType.value & " Configurations uploading failed", , Critical, "DocentIMS"
        End If
    Else 'If IsProjectSelected Then
        If frmMsgBox.Display(Array(cbDocumentType.value & " will not be available to team members to use.", "Are you sure you want to save?"), _
                    Array("Save", "Cancel"), Exclamation, "Incomplete Configuration", Array(0, 255)) = "Save" Then
            GoTo ConsiderComplete
        End If
'    Else
'        For i = 1 To UBound(ProjectName)
'            If UploadSaveLocJSON(ProjectName(i)) Then
'                If Len(MsgOkStr) = 0 Then MsgOkStr = "Documents Configurations Uploaded for:"
'                MsgOkStr = MsgOkStr & Chr(10) & ProjectName(i)
'            Else
'                If Len(MsgBadStr) = 0 Then MsgBadStr = "Documents Configurations could not be uploaded for:"
'                MsgBadStr = MsgBadStr & Chr(10) & ProjectName(i)
'            End If
'        Next
'        If Len(MsgOkStr) > 0 Then
'            If Len(MsgBadStr) > 0 Then
'                MsgBadStr = Chr(10) & Chr(10) & MsgBadStr
'                MsgType = Exclamation
'            Else
'                MsgType = Success
'                SaveConfig = True
'            End If
'        Else
'            MsgType = Critical
'        End If
    End If
'    frmMsgBox.Display MsgOkStr & MsgBadStr, , MsgType, "DocentIMS"
End Function
Private Sub btnSaveNew_Click()
    SaveLastChoices
    If SaveConfig Then ClearAll 2
End Sub
Private Sub btnSaveExit_Click()
    SaveLastChoices
    If SaveConfig Then Unload Me
End Sub
Private Sub cbDocumentType_Change()
    On Error Resume Next
    SaveLastChoices
    ChangesMade = False
    ClearAll
    LoadChoices
    LastDocType = cbDocumentType.value
    UpdateEnabledButtons
End Sub
Private Sub cbProject_Change()
    Dim i As Long
    SaveLastChoices
    ClearAll 1
    LastDocType = ""
    LastPName = cbProject.value
    ChangesMade = False
'    LoadProjectInfoReg LastPName
    If LastPName = "-- Choose Project --" Then
        lbPrjHeader.BackStyle = fmBackStyleTransparent
        lbPrjHeader.Caption = ""
    Else
        LoadProjectInfoReg LastPName
        For i = 1 To UBound(documentName)
            cbDocumentType.AddItem documentName(i)
        Next
        For i = 1 To UBound(MeetingDocName)
            cbDocumentType.AddItem MeetingDocName(i)
        Next
        For i = 1 To UBound(ManagerDocName)
            cbDocumentType.AddItem ManagerDocName(i)
        Next
        Set OldLocations = GetLocations(cbProject.value)
        lbPrjHeader.Caption = IIf(LastPName Like DDPlaceHolder & "*", "", LastPName)
        lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
        lbPrjHeader.BackColor = ProjectColorStr
        lbPrjHeader.BackStyle = fmBackStyleOpaque
    End If
    UpdateEnabledButtons
End Sub
Private Sub UpdateEnabledButtons()
    Dim SvExt As Boolean, SvNew As Boolean, DocSelected As Boolean, ProjectSelected As Boolean
    Dim TestModeUpdated As Boolean
    TestModeUpdated = OriginalTestMode <> GetLocationsTestMode
    ProjectSelected = (Not cbProject.value Like DDPlaceHolder & "*") And (cbProject.value <> "")
    DocSelected = (Not cbDocumentType.value Like DDPlaceHolder & "*") And (cbDocumentType.value <> "")
    DocSelected = ProjectSelected And DocSelected
    SvExt = DocSelected Or GetLocationsTestMode ' SaveLocTestingMode
    SvNew = DocSelected And Not GetLocationsTestMode ' SaveLocTestingMode
    btnSaveExit.Enabled = ChangesMade Or TestModeUpdated '(SvExt And ChangesMade) Or ChangesMade
    btnSaveNew.Enabled = ChangesMade Or TestModeUpdated '(SvNew And ChangesMade) Or ChangesMade
    btnSaveDoc.Enabled = ChangesMade
    cbProject.Enabled = Not DocSelected
    cbDocumentType.Enabled = ProjectSelected And Not DocSelected
End Sub
Private Sub ckTest_Change()
    If IsNull(ckTest.value) Then
'        SaveLocTestingMode = False
        SetLocationsTestMode False
    Else
'        SaveLocTestingMode = ckTest.Value
        SetLocationsTestMode ckTest.value 'False
    End If
    Dim Ctrl As control
    For Each Ctrl In Me.Controls
        Select Case Ctrl.Name
        Case "ckTest", "btnCancel", "btnSaveNew", "btnSaveExit"
'        Case "btnSaveExit"
'            Ctrl.Enabled = SaveLocTestingMode Or ((Not LastPName Like DDPlaceHolder & "*") And (Not LastDocType Like DDPlaceHolder & "*"))
        Case Else: Ctrl.Enabled = Not GetLocationsTestMode 'Not SaveLocTestingMode
        End Select
    Next
'    ChangesMade = True
    UpdateEnabledButtons
    SaveLastChoices
End Sub

Private Sub btnCustomer_Click(): ShowBrowser btnCustomer: End Sub
'    Dim OldPath As String
''    tbCustomer.Value = frmBrowse.Display("Customer Repository", tbCustomer.Value)
'End Sub
Private Sub btnLocal_Click(): ShowBrowser btnLocal: End Sub
'    Dim OldPath As String
'    On Error Resume Next
'    OldPath = Paths("Local")
'    Paths.Remove "Local"
'    Paths.Add "Local", frmBrowse.Display("My Company", OldPath)
'    If Paths.Exists("Local") Then
'
'    End If
'End Sub
Private Sub ShowBrowser(btn As CommandButton)
    Dim BName As String, OldPath As String
    BName = Right$(btn.Name, Len(btn.Name) - 3)
    On Error Resume Next
    OldPath = Paths(BName)
    Paths.Remove BName
    Paths.Add BName, frmBrowse.Display("Browse to " & IIf(BName = "Local", "My Company", BName) & " Location", OldPath)
    ChangesMade = ChangesMade Or (OldPath <> Paths(BName) And CheckPath(BName))
    UpdateEnabledButtons
End Sub
Private Function CheckPath(BName As String) As Boolean
    On Error GoTo ex
    With Controls("btn" & BName)
        If Len(Paths(BName)) Then
            .ControlTipText = Paths(BName)
            .BackColor = &HC0FFC0
            .Caption = Configured
            CheckPath = True
        Else
            .ControlTipText = ""
            .BackColor = &HC0C0FF
            .Caption = "Configure"
        End If
    End With
ex:
End Function
Private Sub lvCustomer_ItemCheck(ByVal Item As MSComctlLib.ListItem): NotPrivate Item: End Sub
Private Sub lvLocal_ItemCheck(ByVal Item As MSComctlLib.ListItem): NotPrivate Item: End Sub
Private Sub NotPrivate(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ex
    If Item.Checked And lvStates.ListItems(Item.Index).text = "Private" Then
        Item.Checked = False
    Else
        ChangesMade = True
    End If
    btnSaveDoc.Enabled = ChangesMade
ex:
End Sub

'Private Sub tbCustomer_Enter()
'    If tbCustomer.Value = PHBrowse Then
'        tbCustomer.Value = frmBrowse.Display("Customer Repository", tbCustomer.Value)
'    End If
'End Sub
'Private Sub tbCustomer_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    tbCustomer.Value = frmBrowse.Display("Customer Repository", tbCustomer.Value)
'End Sub
'Private Sub tbLocal_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    tbLocal.Value = frmBrowse.Display("Local Repository", tbLocal.Value)
'End Sub
'Private Sub tbLocal_Enter()
'    If tbLocal.Value = PHBrowse Then
'        tbLocal.Value = frmBrowse.Display("Local Repository", tbLocal.Value)
'    End If
'End Sub
'Private Sub tbLocal_Change(): UpdateTip tbLocal: End Sub
'Private Sub tbWeb_Change(): UpdateTip tbWeb: End Sub
'Private Sub tbCustomer_Change(): UpdateTip tbCustomer: End Sub
Private Sub tbNamingConv_Change(): UpdateTip tbNamingConv: End Sub

Private Sub UpdateTip(tb As TextBox)
    If tb.value = "" Then tb.value = PHBrowse
    If tb.value = PHBrowse Then
        tb.ForeColor = &H80000000
    Else
        'tb.ForeColor = 0
        tb.ControlTipText = tb.value & "    Set on the project's site"
    End If
End Sub
Private Sub ResetMe()
    
End Sub
Private Sub UserForm_Initialize()
    Dim i As Long
    DisableEvents = True
    ClearAll 2
    With lvStates
        .View = lvwReport
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        With .ColumnHeaders.Add()
            .text = "Document State"
            .Width = lvStates.Width - 4.5
        End With
    End With
    ss = Split("Web,Local,Customer", ",")
    For i = LBound(ss) To UBound(ss)
        With Controls("lv" & ss(i))
            .View = lvwReport
            .FullRowSelect = True
            .LabelEdit = lvwManual
            .ColumnHeaders.Clear
        End With
    Next
    lbHelpLink.Tag = HelpURL
    lvWeb.ColumnHeaders.Add , , "Project Website", lvWeb.Width - 4.5
    lvCustomer.ColumnHeaders.Add , , "Customer", lvCustomer.Width - 4.5
    lvLocal.ColumnHeaders.Add , , "My Company", lvLocal.Width - 4.5
    lvWeb.Enabled = False
'    lvWeb.Locked = True
    lbPrjHeader.BackStyle = fmBackStyleTransparent
    lbPrjHeader.Caption = ""
    tbNamingConv.value = ""
    CenterUserform Me
'    tbCustomer.Value = PHBrowse
'    tbLocal.Value = PHBrowse
    Set Evs.Parent = Me
''    Evs.CollectAllControls
'    Evs.AddOkButton btnSaveExit
'    Evs.AddOkButton btnSaveNew
''    Evs.AddOkButton btnPublish
    Evs.MakeRequired "cbProject,cbDocumentType,tbNamingConv,lvCustomer,lvLocal", , ErrorColor
    LoadChoices
    UpdateEnabledButtons
    DisableEvents = False
    OriginalTestMode = GetLocationsTestMode
'    LoadChoices
End Sub
Private Sub UserForm_Terminate(): RefreshRibbon: End Sub
Private Function IsFullyConfigured() As Boolean
    Dim i As Long, j As Long, iFlag As Boolean, jFlag As Boolean
    IsFullyConfigured = True
    For j = LBound(ss) + 1 To UBound(ss)
        jFlag = Controls("btn" & ss(j)).Caption = Configured
        For i = 1 To Controls("lv" & ss(j)).ListItems.Count
            iFlag = Controls("lv" & ss(j)).ListItems(i).Checked
            If iFlag Then Exit For
        Next
        Controls("lv" & ss(j)).BackColor = IIf(iFlag, -2147483643, vbRed)  'Controls("lv" & ss(j))
        IsFullyConfigured = IsFullyConfigured And iFlag And jFlag
    Next
End Function
