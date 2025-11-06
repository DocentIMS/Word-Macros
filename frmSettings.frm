VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Will/Shall Settings"
   ClientHeight    =   10980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6270
   OleObjectBlob   =   "frmSettings.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DefaultTestPLimit = 25
Private Const CurrentMod = "frmSettings"
Private mHeadingCount As Long, mBookmarksCount As Long, mBoldsCount As Long
Private mDocType As String
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Private Sub btnEditSections_Click()
    WriteLog 1, CurrentMod, "btnEditSections_Click", "Edit Sections button clicked"
    Me.Hide
    frmEditSections.Display mDocType
End Sub
Public Function Refresh() As Boolean
    WriteLog 1, CurrentMod, "Refresh", "Refreshing form"
    Me.Repaint
    Sleep NapDuration
    DoEvents
End Function

Private Sub BtnRun_Click()
    WriteLog 1, CurrentMod, "BtnRun_Click", "Application.run button clicked"
    Me.Hide
    Application.Visible = False
    ProgressBar.Reset
    ProgressBar.Show
    ProgressBar.Spin
    On Error Resume Next
    Set_Coloring = cbColoring.value
    Set_Export = cbExport.value
    Set_Odir = tbODir.value
    Set_TestLimit = IIf(CkTestRun.value, TbTestPLimit.value, 0)
    Set_UseBookmarks = ckUseBookmarks.value
    Set_SearchMode = IIf(Set_UseBookmarks, 0, CoSearchMethod.ListIndex + 2)
    Set_BoldToo = Set_SearchMode = 2 'Or Set_SearchMode = 3
    Set_Indenting = IIf(Set_UseBookmarks, False, cbIndenting.value)
    Set Wrds = New SearchWords
    If CkShall.value And CkShall.Enabled And CkShall.Visible Then Wrds.Add "Shall", "Shall"
    If CKWill.value And CKWill.Enabled And CKWill.Visible Then Wrds.Add "Will", "Will"
    If CkMust.value And CkMust.Enabled And CkMust.Visible Then Wrds.Add "Must", "Must"
    If CkOthers.value And CkOthers.Enabled And CkOthers.Visible And Len(TbOthers.value) > 0 Then
        Dim ss() As String, i As Long, x As Long
        ss = Split(TbOthers.value, ",") ', " ")
        i = LBound(ss)
        Do Until i > UBound(ss)
            If InStr(ss(i), """") > 0 Then
                x = i
                ss(i) = Replace(ss(i), """", "")
                Do
                    i = i + 1
                    ss(x) = ss(x) & " " & Replace(ss(i), """", "")
                Loop Until InStr(ss(i), """") > 0
                Wrds.Add Trim$(ss(x)), "Others"
            Else
                Wrds.Add Trim$(ss(i)), "Others"
            End If
            i = i + 1
        Loop
    End If
    FindSearchRange
    If mParse Then ProgressBar.Caption = mDocType & " Parsing Progress"
    If mWillShall Then ProgressBar.Caption = "Command Statements Parsing Progress"
    ProgressBar.Spin
    On Error GoTo ex
'    If Set_UseBookmarks Then
'        Set FilteredSOWs = AllSOWs.GetRanges(Bookmarks, mParse)
'    Else
    Select Case Set_SearchMode
    Case 0: Set FilteredSOWs = AllSOWs.GetRanges(Bookmarks, mParse)
    Case 2: Set FilteredSOWs = AllSOWs.GetRanges(Headings, mParse)
    Case 1: Set FilteredSOWs = AllSOWs.GetRanges(Bolds, mParse)
    Case Else: Set FilteredSOWs = AllSOWs.GetRanges(HeadingsAndBolds, mParse)
    End Select
'    End If
    If FilteredSOWs Is Nothing Then
        WriteLog 3, CurrentMod, "BtnRun_Click", "No " & mDocType & " found"
        GoTo ex
    End If
    If mWillShall Then CollectWords
    If mParse Then ExportActiveDocument mDocType
ex:
    ProgressBar.HideApplication = False
    Application.Visible = True
    Unload Me
    'Resume
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel button clicked"
    Set_SearchMode = 0
    Set_Cancelled = True
    Unload Me
End Sub
Private Sub btnODir_Click()
    WriteLog 1, CurrentMod, "btnODir_Click", "Browse to output folder button clicked"
    tbODir.value = SelectFolder("Browse to output folder", tbODir.value)
End Sub
Private Sub CbHighlight_Change()
    WriteLog 1, CurrentMod, "CbHighlight_Change", "Highlighting Option set to " & CbHighlight.value
    CKWill.Enabled = CbHighlight.value
    CkShall.Enabled = CbHighlight.value
    CkMust.Enabled = CbHighlight.value
    CkOthers.Enabled = CbHighlight.value
    TbOthers.Enabled = CbHighlight.value
End Sub

Private Sub CkTestRun_Change()
    WriteLog 1, CurrentMod, "CkTestRun_Change", "Test Application.run Option set to " & CkTestRun.value
    SbTestPLimit.Enabled = CkTestRun.value
    TbTestPLimit.Enabled = CkTestRun.value
End Sub
Private Sub CoSearchMethod_Change()
    WriteLog 1, CurrentMod, "CoSearchMethod_Change", "Search Method Option set to " & CoSearchMethod.value
    UpdateSelectedCount
End Sub
Private Sub ckFindHeaders_Change()
    WriteLog 1, CurrentMod, "ObFindHeaders_Change", "Use Headers Option Selected"
    ToggleFrame fraSearchOpts
    UpdateSelectedCount
    btnRun.Enabled = True
End Sub
Private Sub ckUseBookmarks_Change()
    WriteLog 1, CurrentMod, "ObUseBookmarks_Change", "Use Bookmarks Option Selected"
    UpdateSelectedCount
    btnRun.Enabled = True
End Sub
Private Sub SbTestPLimit_Change()
    TbTestPLimit.value = SbTestPLimit.value
    CkTestRun.value = True
End Sub
Private Sub TbOthers_Change()
    CkOthers.value = Len(TbOthers.value) > 0
End Sub
Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "Settings Form Initializing"
    Dim c As Long
    ProgressBar.BarsColor CLng(ProjectColorStr)
'    CoSearchMethod.AddItem "Nearest bold font"
    CoSearchMethod.AddItem "Word Headings Only"
    CoSearchMethod.AddItem "Word Headings OR Bold"
    CoSearchMethod.ListIndex = 1
    tbODir.value = Environ("Userprofile") & "\Desktop"
    lbBookmarksCount.Caption = "Counting..."
'    LbStylesCount2.Caption = "Counting..."
    
    Dim Rng As Range
    On Error Resume Next
    Set Rng = ActiveDocument.Range
    Rng.Move wdStory
    SbTestPLimit.Max = Rng.Information(wdActiveEndPageNumber)
    If SbTestPLimit.Max > DefaultTestPLimit Then SbTestPLimit.value = DefaultTestPLimit
    TbTestPLimit.value = SbTestPLimit.value
    ToggleFrame fraSearchOpts
    CkTestRun.value = False
    ckAdv.value = False
    RemoveCloseButton Me
    CenterUserform Me
End Sub
Sub ParsingMode(DocType As String)
    mDocType = DocType
    WriteLog 1, CurrentMod, "ParsingMode", "Parsing Mode Initializing"
    If IsValidUser = "Ok" Then
        fraOthers.Caption = "Options"
        Me.Caption = DocType & " Parsing Configuration"
        HideControl Array(cbColoring)
        'HideControl Array(fraWords)
        'HideControl Array(fraSearch)
        HideControl Array(cbExport)
        HideControl Array(btnODir, LblODir, tbODir)
        CKWill.Left = CKWill.Left + 12
        CkShall.Left = CkShall.Left + 12
        CkMust.Left = CkMust.Left + 12
        CkOthers.Left = CkOthers.Left + 12
        CbHighlight.value = False
        'Set AllSOWs = Nothing
        ResetSetGlobals
        mParse = True
        Me.Show
        ProgressBar.Show
        Unprotect SDoc
        UpdateCounts
        AllSOWs.ColorAll
        ProgressBar.Hide
        btnRun.Enabled = True
        Protect SDoc
    Else
        MsgBox "Incorrect login info", vbCritical, ""
        Unload Me
    End If
End Sub
Sub ShallWillMode()
    WriteLog 1, CurrentMod, "ShallWillMode", "Shall Will Mode Initializing"
    Dim i As Long, c As Long
    Me.Caption = "Command Statements Configuration"
    Label5.Caption = "Select the Command Words to be highlighted"
    btnRun.Caption = "Run"
    fraSearch.Caption = "Headings Criteria"
    btnEditSections.Visible = False
    HideControl Array(CbHighlight)
    ResetSetGlobals
    mWillShall = True
    Me.Show
    UpdateCounts
    btnRun.Enabled = True
End Sub
Sub UpdateCounts()
    Dim t As Single, i As Long, IsDocentScope As Boolean
    WriteLog 1, CurrentMod, "UpdateCounts", "Updating Counts"
    t = Timer
    'Dim IsCounted As Boolean
    ProgressBar.BarsCount = 2
    ProgressBar.Dom(1) = 6
    ProgressBar.Caption = mDocType & " Analysis"
    ProgressBar.Show
    On Error Resume Next
    i = ActiveDocument.Bookmarks("EOT").Range.start
    IsDocentScope = mDocType = "Scope" And i > 0
    If IsDocentScope Then
        lbHeadingsCount.Visible = False
        lbBoldsCount.Visible = False
        btnAdv.Enabled = False
        ckUseBookmarks.value = True
    Else
        ckFindHeaders.value = True
    End If
    With AllSOWs
        If lbBookmarksCount.Visible Then ' Or LbBookmarksCount2.Visible Then
            mBookmarksCount = .bookmarksCount
            ckUseBookmarks.Enabled = mBookmarksCount > 0
            lbBookmarksCount.Caption = mBookmarksCount
            Refresh
        End If
        If lbHeadingsCount.Visible Then ' Or LbStylesCount2.Visible Then
            mHeadingCount = .headingCount
            lbHeadingsCount.Caption = mHeadingCount
            Refresh
        End If
        If lbBoldsCount.Visible Then ' Or LbStylesCount2.Visible Then
            mBoldsCount = .boldsCount
            lbBoldsCount.Caption = mBoldsCount
        End If
    End With
    ProgressBar.Hide
    t = Timer - t
    If t < 5 Then
        lbTime.Caption = IIf(IsDocentScope, "This is a Docent-created Scope." & Chr(10) & "Thus, advanced options are not needed.", "")
    Else
        If t \ 60 > 0 Then
            If t \ 60 = 1 Then
                lbTime.Caption = "Time taken: 1 minute and " & (t Mod 60) \ 1 & " seconds."
            Else
                lbTime.Caption = "Time taken: " & t \ 60 & " minutes and " & (t Mod 60) \ 1 & " seconds."
            End If
        Else
            lbTime.Caption = "Time taken: " & (t Mod 60) \ 1 & " seconds."
        End If
    End If
    UpdateSelectedCount
End Sub
Private Sub UpdateSelectedCount()
    Dim n As Long
    If ckUseBookmarks.value Then n = n + mBookmarksCount
    If ckFindHeaders.value Then
        Select Case CoSearchMethod.ListIndex + 2
        '    Case 0: lbSelectedCount.Caption = mBookmarksCount
        Case 2: n = n + mHeadingCount
        Case 1: n = n + mBoldsCount
        Case Else: n = n + mHeadingCount + mBoldsCount
        End Select
    End If
    lbSelectedCount.Caption = n
End Sub
Private Sub ckAdv_Change()
    HideControl Array(fraSearch, fraOthers), Not ckAdv.value
   ' HideControl Array(fraOthers), Not ckAdv.Value
End Sub
Private Sub btnAdv_Click()
    HideControl Array(fraSearch, fraOthers), fraSearch.Visible
    'HideControl Array(fraOthers), fraOthers.Visible
End Sub
Private Sub ToggleFrame(Fra As MSForms.Frame)
    Dim Ctrl As MSForms.control
    On Error Resume Next
    Fra.Enabled = Not Fra.Enabled
    For Each Ctrl In Fra.Controls
        Ctrl.Enabled = Fra.Enabled
    Next
End Sub
Private Sub HideControl(Ctrls, Optional Hide As Boolean = True)
    Dim t As Single, h As Single, tMin As Single, hSpacing As Single
    Dim i As Long, Ctrl As control
    Dim Parents
    tMin = Me.Height + Me.Top
    Const frmHSpace = 6
    Const OtherHSpacing = 0.5
    For i = LBound(Ctrls) To UBound(Ctrls)
        Set Ctrl = Ctrls(i)
        Ctrl.Visible = Not Hide
        t = GetTop(Ctrl)
        If tMin > t Then tMin = t
        If h < Ctrl.Height Then h = Ctrl.Height
        hSpacing = IIf(Ctrl.Name Like "fra*", frmHSpace, OtherHSpacing)
    Next
    For Each Ctrl In Me.Controls
        If GetTop(Ctrl) >= tMin And Ctrl.Name <> Ctrls(0).Name Then
            If Ctrl.Parent.Name = Me.Name Then
                Ctrl.Top = IIf(Hide, Ctrl.Top - h, Ctrl.Top + h)
                Ctrl.Top = IIf(Hide, Ctrl.Top - hSpacing, Ctrl.Top + hSpacing)
            ElseIf Ctrl.Parent.Name = Ctrls(0).Parent.Name Then
                Ctrl.Top = IIf(Hide, Ctrl.Top - h, Ctrl.Top + h)
            End If
        End If
    Next
    Me.Height = IIf(Hide, Me.Height - h, Me.Height + h)
    Parents = GetParents(Ctrls(0))
    For i = 0 To UBound(Parents)
        Parents(i).Height = IIf(Hide, Parents(i).Height - h, Parents(i).Height + h)
    Next
End Sub
Private Function GetTop(ByVal Ctrl) As Single
    Dim Parents, i As Long
    Parents = GetParents(Ctrl)
    GetTop = Ctrl.Top
    For i = 0 To UBound(Parents)
        GetTop = GetTop + Parents(i).Top
    Next
End Function
Private Function GetParents(ByVal Ctrl) As Variant
    Dim Arr() As control, i As Long
    Do While Ctrl.Parent.Name <> Me.Name
        ReDim Preserve Arr(0 To i)
        Set Arr(i) = Ctrl.Parent
        Set Ctrl = Ctrl.Parent
        i = i + 1
    Loop
    GetParents = Arr
End Function
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Set_SearchMode = 0
        Set_Cancelled = True
        Unload Me
    End If
End Sub
Sub IdentifyMode()
    WriteLog 1, CurrentMod, "IdentifyMode", "Identify Mode Initializing"
    fraWords.Caption = "Search And Highligh"
    HideControl Array(fraSearch)
    HideControl Array(cbColoring)
    HideControl Array(cbExport)
    HideControl Array(btnODir, LblODir, tbODir)
    CKWill.Left = CKWill.Left + 12
    CkShall.Left = CkShall.Left + 12
    CkMust.Left = CkMust.Left + 12
    CkOthers.Left = CkOthers.Left + 12
    CbHighlight.value = False
    Me.Caption = "Identify Sections"
    btnRun.Enabled = True
    Me.Show
End Sub

