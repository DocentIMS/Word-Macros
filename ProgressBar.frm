VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar 
   Caption         =   "Progress"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   OleObjectBlob   =   "ProgressBar.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Private Const CurrentMod = "ProgressBar"
Const SpinnersCount = 3
Const SpinnerSteps = 20
Const SpinnerSpacing = 0
Const Freq = 5
Const BetweenLevels = 53
Const BorderOffset = 0.5

Private mBarsCount
Private vNom() As Long, vDom() As Long
Private mMaxW As Long
Private mBoxLeft As Long
'Private mFrmRatio1Top As Long
'Private mLbT1Top As Long
'Private mTxtRatio1Top As Long
Private mMeHeight As Long
Private mDivider As Long
Private mKilled As Boolean
Private mIsFinished As Boolean
Private mKeepHidden As Boolean

Private SpinP() As Long, t As Single
Public Property Get Canceled() As Boolean
    Set_Cancelled = mKilled
End Property
Public Property Let Dom(Optional ByVal n As Long, ByVal DomVal As Long)
    If n = 0 Then n = mBarsCount
    vDom(n) = DomVal
End Property
Public Property Let RefreshRate(ByVal n As Long)
    mDivider = n
End Property
Public Property Let HideApplication(ByVal Hide As Boolean): mKeepHidden = Hide: End Property
Sub BarsColor(clr As Long)
    Dim Ctrl As control
    For Each Ctrl In Controls
        If Ctrl.Name Like "bar*" Or Ctrl.Name Like "Spin*" Then Ctrl.BackColor = clr
    Next
End Sub
Public Property Get BarsCount() As Long
    BarsCount = mBarsCount
End Property
Private Function GetControl(CtrlName) As control
    On Error Resume Next
    Set GetControl = Controls(CtrlName)
End Function
Private Sub AddProgressBar(n As Long)
    Dim Ctrl As control, i As Long
    Set Ctrl = GetControl("frmRatio_" & n)
    If Ctrl Is Nothing Then Set Ctrl = Controls.Add("Forms.Label.1", "frmRatio_" & n, True)
    Ctrl.Left = frmRatio0.Left
    Ctrl.Width = frmRatio0.Width
    Ctrl.Top = frmRatio0.Top + (BetweenLevels * (mBarsCount - n))
    Ctrl.BorderStyle = frmRatio0.BorderStyle
    Ctrl.BorderColor = frmRatio0.BorderColor
    Ctrl.BackStyle = frmRatio0.BackStyle
    Ctrl.Enabled = False
    Ctrl.Height = frmRatio0.Height
    Ctrl.Visible = True
    Set Ctrl = GetControl("bar_" & n)
    If Ctrl Is Nothing Then Set Ctrl = Controls.Add("Forms.Label.1", "bar_" & n, True)
    Ctrl.Left = bar0.Left
    Ctrl.Top = bar0.Top + (BetweenLevels * (mBarsCount - n))
    Ctrl.Height = bar0.Height
    Ctrl.BackColor = bar0.BackColor
    Ctrl.Width = 0
    Ctrl.Visible = True
    Set Ctrl = GetControl("LbT_" & n)
    If Ctrl Is Nothing Then Set Ctrl = Controls.Add("Forms.Label.1", "LbT_" & n, True)
    Ctrl.Left = LbT0.Left
    Ctrl.Top = LbT0.Top + (BetweenLevels * (mBarsCount - n))
    Ctrl.Height = LbT0.Height
    Ctrl.Font.Name = LbT0.Font.Name
    Ctrl.Font.Size = LbT0.Font.Size
    Ctrl.Caption = IIf(n = 1, "Overall Progress", "Current Progress")
'    Ctrl.Caption = Replace(LbT0.Caption, "Overall", "Current")
    Ctrl.Width = LbT0.Width
    Ctrl.Visible = True
    Set Ctrl = GetControl("txtRatio_" & n)
    If Ctrl Is Nothing Then Set Ctrl = Controls.Add("Forms.Label.1", "txtRatio_" & n, True)
    Ctrl.Left = txtRatio0.Left
    Ctrl.Top = txtRatio0.Top + (BetweenLevels * (mBarsCount - n))
    Ctrl.Height = txtRatio0.Height
    Ctrl.Width = txtRatio0.Width
    Ctrl.Font.Name = txtRatio0.Font.Name
    Ctrl.Font.Size = txtRatio0.Font.Size
    Ctrl.Caption = txtRatio0.Caption
    Ctrl.Visible = True
    For i = 1 To SpinnersCount
        Set Ctrl = GetControl("Spin_" & i & "_" & n)
        If Ctrl Is Nothing Then Set Ctrl = Me.Controls.Add("Forms.Label.1", "Spin" & i & "_" & n, True)
        Ctrl.Visible = False
        Ctrl.Top = Controls("Spin1_0").Top + (BetweenLevels * (mBarsCount - n))
        Ctrl.Width = Controls("Spin1_0").Width
        Ctrl.Height = Controls("Spin1_0").Height
        Ctrl.BackColor = Controls("Spin1_0").BackColor
        Ctrl.Left = Controls("Spin1_0").Left
    Next
    RefreshPositions
End Sub
Private Sub RefreshPositions()
    Dim n As Long, Ctrl As control
    Me.Height = mMeHeight + (BetweenLevels * (mBarsCount))
    'For n = 1 To mBarsCount
        For Each Ctrl In Me.Controls
            setCtrlTop Ctrl
        Next
   ' Next
End Sub
Private Function setCtrlTop(Ctrl As control) As Single
    Dim i As Long, n As Long, CtrlName As String
    CtrlName = Ctrl.Name
    i = InStrRev(CtrlName, "_")
    If i > 0 Then
        If IsNumeric(Mid(CtrlName, i + 1)) Then
            n = Mid(CtrlName, i + 1)
            If n Then Ctrl.Top = Controls(Left(CtrlName, i) & "0").Top + (BetweenLevels * (mBarsCount + 1 - n))
        End If
    End If
End Function
Public Property Let BarsCount(ByVal x As Long)
    Dim n As Long, i As Long, Ctrl As control ', Lb As Label
'    If x > mBarsCount Then
'        For n = x + 1 To mBarsCount
'            Controls.Remove "frmRatio_" & n
'            Controls.Remove "bar_" & n
'            Controls.Remove "LbT" & n
'            Controls.Remove "txtRatio_" & n
'            For i = 1 To SpinnersCount
'                Controls.Remove "Spin" & i & "_" & n
'            Next
'        Next
'    Else
    mBarsCount = x
    ReDim vNom(1 To mBarsCount) As Long
    ReDim vDom(1 To mBarsCount) As Long
    ReDim SpinP(1 To mBarsCount) As Long
    Dim mBoxLeft As Single
    On Error Resume Next
    For n = 1 To mBarsCount
        vNom(n) = 0
        vDom(n) = 1
        SpinP(n) = 0
        AddProgressBar n
    Next
End Property
Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "Progress Bar Form Initialized"
    Dim i As Long, Ctrl As control
    mKeepHidden = False
    mMeHeight = Me.Height - BetweenLevels
    mMaxW = Fix(frmRatio0.Width / SpinnerSteps) * SpinnerSteps
    mBoxLeft = frmRatio0.Left + ((frmRatio0.Width - mMaxW) / 2)
    LbT00.Left = mBoxLeft + BorderOffset
    LbT0.Left = mBoxLeft + BorderOffset
    txtStatus00.Left = mBoxLeft + LbT00.Width
    frmRatio0.Left = mBoxLeft
    frmRatio0.Width = mMaxW
    bar0.Left = mBoxLeft + BorderOffset
    bar0.Top = frmRatio0.Top + BorderOffset
    For i = 1 To SpinnersCount
        Set Ctrl = Me.Controls.Add("Forms.Label.1", "Spin" & i & "_0", True)
        Ctrl.Visible = False
        Ctrl.Top = bar0.Top 'Spin1_1.Top + (BetweenLevels * (mBarsCount - n))
        Ctrl.Width = (mMaxW - 1) / SpinnerSteps
        Ctrl.Width = Ctrl.Width - SpinnerSpacing
        Ctrl.Height = frmRatio0.Height - (2 * BorderOffset) ' Spin1_1.Height
        Ctrl.Left = mBoxLeft - Ctrl.Width + BorderOffset
        Ctrl.BackColor = bar0.BackColor
    Next
    bar0.Width = 0
    mDivider = 1
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
    lbPrjHeader.BackColor = ProjectColorStr
    CenterUserform Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        If Not mIsFinished Then
            WriteLog 1, CurrentMod, "UserForm_QueryClose", "User Cancelled"
            Cancel = True
            Progress 1, "Cancelling...", 0
            mKilled = True
            Set_Cancelled = True
        End If
    End If
    Application.Visible = True
End Sub
Public Sub Finished(Optional ByVal status As String = "Done")
    Dim n As Long
    For n = 1 To mBarsCount
        vNom(n) = vDom(n)
        If n = 1 Then txtStatus00.Caption = status
        ProgressN n
    Next
    mIsFinished = True
    Me.MousePointer = fmMousePointerDefault
    mKeepHidden = False
    Application.Visible = True
End Sub
Public Function Progress(Optional ByVal n As Long, _
        Optional ByVal status As String, _
        Optional ByVal Step As Long = 1) As Boolean
    If mKeepHidden Then If Application.Visible Then Application.Visible = False
    If mBarsCount = 0 Then BarsCount = 1
    If n = 0 Then n = mBarsCount
    vNom(n) = vNom(n) + Step
    If vNom(n) >= vDom(n) Then vNom(n) = vDom(n)
    If vNom(n) = vDom(n) Then 'Or n <> mBarsCount
        ProgressN n, status
    Else 'If n = mBarsCount Then
        If vNom(n) Mod mDivider = 0 Then ProgressN n, status
    End If
    If n = 1 And vNom(n) = vDom(n) Then Finished Else Reset n + 1
    Progress = mKilled
End Function
Public Sub Reset(Optional ByVal n As Long = -1)
    If mBarsCount = 0 Then BarsCount = 1
    If n = -1 Then n = 1
    For n = n To mBarsCount
        vNom(n) = 0
        Controls("bar_" & n).Width = 0
        Controls("txtRatio_" & n).Caption = "0 %"
        If n = 1 Then txtStatus00.Caption = ""
    Next
    Refresh
End Sub
Public Function Refresh() As Boolean
    Me.Repaint
    Sleep NapDuration
    DoEvents
    Refresh = mKilled
End Function
Private Sub ProgressN(ByVal n As Long, Optional ByVal status As String)
    Dim Rstr As String, Ratio As Single, i As Long
    If vDom(n) = 0 Then vNom(n) = 1: vDom(n) = 1
    Ratio = vNom(n) / vDom(n)
    Rstr = vNom(n) & " / " & vDom(n) & " (" & Format(Ratio, "0 %") & ")"
    Ratio = (mMaxW - (2 * BorderOffset)) * Ratio
    Controls("bar_" & n).Width = Ratio
    'Status = "(" & vNom(n) & " / " & vDom(n) & ") " & Status
'    If Len(Controls("txtRatio_" & n).Caption) > 0 Then
'    If Controls("Spin1_" & n).Visible Then
        For i = 1 To SpinnersCount
            If Controls("Spin" & i & "_" & n).Visible Then
                Controls("Spin" & i & "_" & n).Visible = False
            End If
        Next
'    End If
    Controls("txtRatio_" & n).Caption = Rstr
    If status <> "" And n = 1 Then txtStatus00.Caption = status
    Refresh
End Sub
Public Function Spin(Optional ByVal n As Long, _
        Optional ByVal status As String, _
        Optional mDivider As Long) As Boolean
    If mKeepHidden Then If Application.Visible Then Application.Visible = False
    If mBarsCount = 0 Then BarsCount = 1
    If n = 0 Then n = mBarsCount
    SpinP(n) = SpinP(n) + 1
    'Status = "(" & vNom(n) & " / " & vDom(n) & ") " & Status
    If mDivider = 0 Then
        If Abs(Timer - t) > (1 / Freq) Then SpinN n, status: t = Timer
    ElseIf SpinP(n) Mod mDivider = 0 Then
        SpinN n, status
    End If
    If status <> "" Then txtStatus00.Caption = status
    Spin = mKilled
End Function
Private Sub SpinN(ByVal n As Long, Optional ByVal status As String)
    Dim Rstr As String, Ratio As Single, i As Long, Ctrl As control
    'Ratio = vNom(n) / vDom(n)
    If Len(Controls("txtRatio_" & n).Caption) Then Controls("txtRatio_" & n).Caption = ""
    If Controls("bar_" & n).Width Then Controls("bar_" & n).Width = 0
    If status <> "" And n = 1 Then txtStatus00.Caption = status
    For i = 1 To SpinnersCount
        Set Ctrl = Controls("Spin" & i & "_" & n)
        If i = 1 Then
            Ratio = Ctrl.Left
        Else
            Ratio = Controls("Spin" & i - 1 & "_" & n).Left + SpinnerSpacing
        End If
        Ctrl.Left = Ratio + Ctrl.Width
        If Ctrl.Left + Ctrl.Width <= _
                Controls("frmRatio_" & n).Left + _
                Controls("frmRatio_" & n).Width + BorderOffset Then
            Ctrl.Visible = True
        Else
            Ctrl.Visible = False
            Ctrl.Left = Controls("frmRatio_" & n).Left - Ctrl.Width + BorderOffset
        End If
    Next
    Refresh
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
End Sub


