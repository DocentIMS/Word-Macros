VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsgBox 
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5025
   OleObjectBlob   =   "frmMsgBox.frx":0000
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Const CurrentMod = "frmMsgBox"
Private mClickedButton As String
Private ButtonsArray() As New ButtonsEvents
Private LabelsArray() As New ButtonsEvents
Private mBtnWidth As Single
Private mButtonsTop As Single
Private mBtnHeight As Single
Private mAutocloseTimer As Single
Private iMax As Long
Private Const MinSpace = 6
Private Const WidthDiff = 10

'#If Win64 Then
'    Private Declare PtrSafe Function DrawText _
'        Lib "user32" Alias "DrawTextA" ( _
'        ByVal hdc As Long, _
'        ByVal lpStr As String, _
'        ByVal nCount As Long, _
'        lpRect As RECT, _
'        ByVal wFormat As Long) As Long
'#Else
'    Private Declare Function DrawText _
'        Lib "user32" Alias "DrawTextA" ( _
'        ByVal hdc As Long, _
'        ByVal lpStr As String, _
'        ByVal nCount As Long, _
'        lpRect As RECT, _
'        ByVal wFormat As Long) As Long

'#End If
'
'
'Private Const DT_CALCRECT = &H400&
'Private Const DT_WORDBREAK = &H10&
'
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Private Sub SetLabelCaption(lbl As Label, ByVal sCaption As String)
'    Dim r As RECT
'    Dim nOldScaleMode As Long
'    Dim nBorder As Long
'    nOldScaleMode = Me.ScaleMode
'    'Change the scalemode to Pixels to simplify the calculations
'    Me.ScaleMode = vbPixels
'    If lbl.BorderStyle <> vbBSNone Then
'        If lbl.Appearance = 1 Then
'            '3D border
'            nBorder = 4
'        Else
'            nBorder = 2
'        End If
'    End If
'    r.Right = lbl.Width - nBorder
'    DrawText Me.hdc, sCaption, -1, r, DT_WORDBREAK + DT_CALCRECT
'    lbl.Height = r.Bottom + nBorder
'    lbl.Caption = sCaption
'    'Restore the ScaleMode
'    Me.ScaleMode = nOldScaleMode
'End Sub
Private Sub UserForm_Initialize()
    Reset
    CenterUserform Me
End Sub
'Private Sub Lb4_Click()
'    On Error Resume Next
'    GoToLink Lb4.Caption
'    Unload Me
'End Sub
Sub ClickedButton(ButtonCaption As String)
    mClickedButton = ButtonCaption
    Me.Hide
End Sub
Sub ClickedLabel(LabelTag As String)
    If Len(LabelTag) = 0 Then Exit Sub
    GoToLink LabelTag
   ' Unload Me
End Sub
Private Sub AddButton(ButtonCaption, ByVal n As Long)
    If mBtnWidth = 0 Then mBtnWidth = 80
    If mBtnHeight = 0 Then mBtnHeight = 40
    Dim Ctrl As CommandButton
    ReDim Preserve ButtonsArray(1 To n)
    Set Ctrl = Controls.Add("Forms.CommandButton.1", "btn" & n, True)
    Ctrl.Top = 120
    Ctrl.Height = mBtnHeight
    Ctrl.Width = mBtnWidth
    Ctrl.Font.Size = 12
    Ctrl.WordWrap = True
    Ctrl.AutoSize = True
    Ctrl.Caption = ButtonCaption
'    Ctrl.AutoSize = False
'    Ctrl.AutoSize = True
    Ctrl.AutoSize = False
    If Ctrl.Width > mBtnWidth Then mBtnWidth = Ctrl.Width
'    If Ctrl.Height > mBtnHeight Then mBtnHeight = Ctrl.Height
    Set ButtonsArray(n).btn = Ctrl
    Set ButtonsArray(n).Parent = Me
End Sub
Private Sub ArrangeButtons(n As Long)
    Dim TotWidth As Single, Space As Single
    TotWidth = Me.Width ' - WidthDiff
    Space = (TotWidth - (mBtnWidth * n)) / (n + 1)
    If Space < MinSpace Then Space = MinSpace
    For n = 1 To n
        With Controls("btn" & n)
            .Width = mBtnWidth
            .Height = mBtnHeight
            .Left = Space + ((Space + mBtnWidth) * (n - 1))
            .Top = mButtonsTop
        End With
    Next
    If mAutocloseTimer <> -1 Then
        Dim Ctrl As label
        'ReDim Preserve ButtonsArray(1 To n)
        Set Ctrl = Controls.Add("Forms.Label.1", "lbAutoClose", True)
        Ctrl.WordWrap = False
        Ctrl.Top = mButtonsTop + mBtnHeight + 5 'Me.Height - 20
        Ctrl.Width = Space
        Ctrl.Left = 10
        Ctrl.Caption = "Auto-Close in " & mAutocloseTimer & " seconds ..."
        Ctrl.Height = 20
        Ctrl.Font.Size = 12
        Ctrl.AutoSize = True
    End If
    Me.Height = mButtonsTop + IIf(mBtnHeight = 0, 10, mBtnHeight) + IIf(mAutocloseTimer = -1, 0, 15) + 40
    If Me.Height < 65 Then Me.Height = 65
End Sub
    
Private Sub UpdateLablesWidth(ButtonsCount As Long, Imgs)
    Dim Space As Single, i As Long
'    Space = (Me.Width - WidthDiff - (mBtnWidth * ButtonsCount)) / (ButtonsCount + 1)
'    If Space < MinSpace Then Space = MinSpace
'    Space = Space + ((Space + mBtnWidth) * (ButtonsCount)) + WidthDiff
'    If Space > Me.Width Then Me.Width = Space
'    Space = Me.Width - WidthDiff - MinSpace - MinSpace
'    If UBound(Imgs) <> -1 Then Space = Space - imgSuccess.Width - imgSuccess.Left
'    For i = 1 To iMax + 1
'        Controls("lb" & i).Width = Space
'    Next
    HideEmptyLbs
    If UBound(Imgs) <> -1 Then
        Dim Img As MSForms.image
        For i = LBound(Imgs) To UBound(Imgs)
            Set Img = GetImg(Imgs(i))
            If Not Img Is Nothing Then
                Img.Visible = True
                Img.Top = Controls("lb" & i + 1).Top
            End If
        Next
    End If
End Sub
'Private Sub PositionLables()
'    Dim i As Single, n As Single
'    For i = n + 1 To 4
'        If Controls("lb" & i).Visible Then
'            Controls("lb" & i).Top = Controls("lb" & n).Top + Controls("lb" & n).Height + 2
'            n = i
'        End If
'    Next
'End Sub
Private Sub HideEmptyLbs()
    Dim Ctrl As control, tCtrl As Single, i As Long, n As Single
    Do Until n = iMax + 1
        n = n + 1
        Set Ctrl = Controls("lb" & n)
        If Len(Ctrl.Caption) = 0 Then
            HideControl Array(Ctrl)
        Else
            Ctrl.AutoSize = False
            Ctrl.AutoSize = True
            tCtrl = Ctrl.Top + Ctrl.Height + 6
            mButtonsTop = IIf(mButtonsTop > tCtrl, mButtonsTop, tCtrl)
            Exit Do
        End If
    Loop
    For i = n + 1 To iMax + 1
        Set Ctrl = Controls("lb" & i)
        Ctrl.Top = Controls("lb" & n).Top + Controls("lb" & n).Height + 2
        If Len(Ctrl.Caption) = 0 Then
            HideControl Array(Ctrl)
        Else
            n = i
            Ctrl.AutoSize = False
            Ctrl.AutoSize = True
            tCtrl = Ctrl.Top + Ctrl.Height + 6
            mButtonsTop = IIf(mButtonsTop > tCtrl, mButtonsTop, tCtrl)
        End If
    Next
'    If Len(Lb1.Caption) = 0 Then HideControl Array(Lb1)
'    If Len(Lb2.Caption) = 0 Then HideControl Array(Lb2)
'    If Len(Lb3.Caption) = 0 Then HideControl Array(Lb3)
'    If Len(Lb4.Caption) = 0 Then HideControl Array(Lb4)
End Sub
'Private Sub HideEmptyLb(Ctrl As Label)
'
'End Sub
Private Sub Reset()
    Dim Ctrl As control
    For Each Ctrl In Me.Controls ' i = Me.Controls.Count To 1 Step -1
'        ctrl.name = Me.ctrl.Name
        If Ctrl.Name Like "lb#" And Ctrl.Name <> "lb1" Then
            Ctrl.Remove
        ElseIf Ctrl.Name Like "btn#" Then
            Ctrl.Remove
        ElseIf Ctrl.Name Like "lb##" Then
            Ctrl.Remove
        ElseIf Ctrl.Name Like "btn##" Then
            Ctrl.Remove
        End If
    Next
    mClickedButton = -1
    ReDim ButtonsArray(1 To 1)
    ReDim LabelsArray(1 To 1)
    mBtnWidth = 0
    mAutocloseTimer = 0
    iMax = 0
End Sub
Function Display(Msgs, Optional Buttons = "OK", Optional MsgType As NewMsgBoxStyle = Success, _
                Optional Title As String, Optional Clrs, Optional BackClrs, _
                Optional Links, Optional ShowModal As FormShowConstants = vbModal, _
                Optional Images, Optional AutoCloseTimer As Single = -1) As String
    WriteLog 1, CurrentMod, "Display", "Displaying MessageBox"
    Dim i As Long, btnsCount As Long
    Reset
    mAutocloseTimer = AutoCloseTimer
    Select Case TypeName(Msgs)
    Case "Variant()", "String()"
        iMax = UBound(Msgs) - LBound(Msgs)
    End Select
    For i = 1 To iMax
        With Controls.Add("Forms.Label.1", "lb" & i + 1, True)
            .Left = Lb1.Left
            .Font.Name = Lb1.Font.Name
            .Font.Size = Lb1.Font.Size
            .WordWrap = True
            If i = 1 Then
                .Top = Lb1.Top + 5
            Else
                .Top = Controls("lb" & i).Top + 5
            End If
        End With
    Next
    ReDim mMsgs(0 To iMax) As String
    ReDim mLinks(0 To iMax) As String
    ReDim mClrs(0 To iMax) As String
    ReDim mBackClrs(0 To iMax) As String
    ReDim mImages(-1 To -1) As String
    
    Select Case TypeName(Msgs)
    Case "Variant()", "String()"
        For i = LBound(Msgs) To UBound(Msgs)
            If IsMissing(Msgs(i)) Then Msgs(i) = ""
            mMsgs(i - LBound(Msgs)) = Msgs(i)
        Next
    Case "String"
        mMsgs(0) = Msgs
    End Select
    
    
    Select Case TypeName(Images)
    Case "Variant()", "String()"
        ReDim mImages(0 To iMax) As String
        For i = LBound(Images) To UBound(Images)
            If IsMissing(Images(i)) Then Images(i) = ""
            mImages(i - LBound(Images)) = Images(i)
        Next
    Case "String"
        ReDim mImages(0 To 0) As String
        mImages(0) = Images
    End Select
    
    i = 0
    Select Case TypeName(Links)
    Case "Variant()", "String()"
        For i = LBound(Links) To UBound(Links)
            If IsMissing(Links(i)) Then Links(i) = ""
            mLinks(i - LBound(Links)) = Links(i)
        Next
    Case "String"
        mLinks(0) = Links
    End Select
    i = 0
    Select Case TypeName(Buttons)
    Case "Variant()", "String()"
        For i = LBound(Buttons) To UBound(Buttons)
            AddButton Buttons(i), i - LBound(Buttons) + 1
        Next
        btnsCount = i - LBound(Buttons)
    Case "String"
        i = i + 1
        AddButton Buttons, i
        btnsCount = i
    End Select
    If IsMissing(BackClrs) Then
        For i = LBound(mBackClrs) To UBound(mBackClrs)
            mBackClrs(i) = &H8000000F
        Next
    Else
        Select Case TypeName(BackClrs)
        Case "Variant()", "String()", "Long()"
            For i = LBound(BackClrs) To UBound(BackClrs)
                mBackClrs(i - LBound(BackClrs)) = BackClrs(i)
            Next
            For i = i To iMax
                mBackClrs(i) = &H8000000F
            Next
        Case "String"
            For i = 0 To iMax
                mBackClrs(i) = BackClrs
            Next
        End Select
    End If
    If IsMissing(Clrs) Then
        For i = LBound(mClrs) To UBound(mClrs)
            mClrs(i) = IIf(FullColor(mBackClrs(i)).TooDark, 16777215, IIf(i = 1, vbRed, IIf(i = 3, 8388608, 0)))
        Next
    Else
        Select Case TypeName(Clrs)
        Case "Variant()", "String()", "Long()"
            For i = LBound(Clrs) To UBound(Clrs)
                mClrs(i - LBound(Clrs)) = Clrs(i)
            Next
            For i = i To iMax
                mClrs(i) = IIf(FullColor(mBackClrs(i)).TooDark, 16777215, IIf(i = 1, vbRed, IIf(i = 3, 8388608, 0)))
            Next
        Case "String", "Long"
            For i = 0 To iMax
                mClrs(i) = Clrs
            Next
        End Select
    End If
    Me.Caption = Title
    On Error Resume Next
    GetImg(MsgType).Visible = True
    If MsgType = None And UBound(mImages) = -1 Then
        For i = 1 To iMax + 1
            Controls("lb" & i).Left = 6
        Next
    End If
    On Error Resume Next
    ReDim LabelsArray(1 To iMax + 1)
    For i = 1 To iMax + 1
        Do While Left$(mMsgs(i - 1), 1) = " "
            Controls("lb" & i).Left = Controls("lb" & i).Left + 4
            mMsgs(i - 1) = Right$(mMsgs(i - 1), Len(mMsgs(i - 1)) - 1)
        Loop
        Controls("lb" & i).Caption = mMsgs(i - 1)
        Controls("lb" & i).AutoSize = False
        Controls("lb" & i).Width = Me.Width - Controls("lb" & i).Left - 5
        Controls("lb" & i).AutoSize = True
        If Len(mLinks(i - 1)) > 0 Then
            Controls("lb" & i).Tag = mLinks(i - 1)
            Controls("lb" & i).Font.Underline = True
            Controls("lb" & i).MousePointer = fmMousePointerCustom
'            Controls("lb" & i).MouseIcon = LoadPicture("C:\Path\To\Your\CustomCursor.ico")
        End If
        Controls("lb" & i).BackColor = mBackClrs(i - 1)
        Controls("lb" & i).ForeColor = mClrs(i - 1)
        Set LabelsArray(i).lb = Controls("lb" & i)
        Set LabelsArray(i).Parent = Me
    Next
'    Lb2.Caption = mMsgs(1)
'    Lb2.BackColor = mBackClrs(1)
'    Lb2.ForeColor = mClrs(1)
'    Lb3.Caption = mMsgs(2)
'    Lb3.ForeColor = mClrs(2)
'    Lb4.Caption = mMsgs(3)
'    Lb4.ForeColor = mClrs(3)
'    PositionLables
    UpdateLablesWidth btnsCount, mImages ' MsgType
    ArrangeButtons btnsCount
    CenterUserform Me
'    PositionLables
'    If ShowModal = vbModal Then Application.Visible = True
'    Me.Visible = True
'    CenterUserform Me
    'Me.Repaint
'    me.Visible
'End
    Dim x As Single
    x = Timer
    i = mAutocloseTimer
    If i <> -1 Then ShowModal = vbModeless
    Me.Show ShowModal
    Repaint
    Do While x > 0 And i <> -1
        If Timer - x > 1 Then
            x = Timer
            i = i - 1
            Controls("lbAutoClose").Caption = "Auto-Close in " & i & " seconds ..."
            Repaint
            AutoCloseTimer = AutoCloseTimer - 1
            If AutoCloseTimer <= 0 Then
                WriteLog 1, CurrentMod, "Display", "Timed Out after " & mAutocloseTimer & " seconds"
                Unload Me
                Exit Function
            End If
        End If
        Sleep 10
        DoEvents
'        If AutoCloseTimer <= 0 Then
'            WriteLog 1, CurrentMod, "Display", "Timed Out after " & mAutocloseTimer & " seconds"
'            Unload Me
'            Exit Function
'        Else
        If Len(mClickedButton) > 0 And mClickedButton <> -1 Then
            WriteLog 1, CurrentMod, "Display", "User Clicked " & mClickedButton
            Display = mClickedButton
            Unload Me
            Exit Function
        End If
    Loop
'    Me.Show msoModeModeless
    If ShowModal = vbModal Then
        Display = mClickedButton
        WriteLog 1, CurrentMod, "Display", "User Clicked " & mClickedButton
        Unload Me
    End If
End Function
Private Function GetImg(MsgType) As MSForms.image
    Select Case MsgType
    Case NewMsgBoxStyle.Critical: Set GetImg = imgCritical: Beep
    Case NewMsgBoxStyle.Success: Set GetImg = imgSuccess
    Case NewMsgBoxStyle.Exclamation: Set GetImg = imgExclamation
    Case NewMsgBoxStyle.Question: Set GetImg = ImgQuestion
    Case NewMsgBoxStyle.Information: Set GetImg = imgInformation
    Case NewMsgBoxStyle.ZoomIcon: Set GetImg = imgZoom
    End Select
End Function
Private Sub HideControl(Ctrls)
    Dim t As Single, h As Single, tMin As Single, hSpacing As Single
    Dim i As Long, Ctrl As control
    Dim Parents
    tMin = Me.Height + Me.Top
    Const frmHSpace = 6
    Const OtherHSpacing = 0.5
    For i = LBound(Ctrls) To UBound(Ctrls)
        Set Ctrl = Ctrls(i)
        Ctrl.Visible = False
        t = GetTop(Ctrl)
        If tMin > t Then tMin = t
        If h < Ctrl.Height Then h = Ctrl.Height
        If Ctrl.Name Like "fra*" Then
            hSpacing = frmHSpace
        Else
            hSpacing = OtherHSpacing
        End If
    Next
    For Each Ctrl In Me.Controls
        If GetTop(Ctrl) > tMin Then
            If Ctrl.Parent.Name = Me.Name Then
                Ctrl.Top = Ctrl.Top - h
                Ctrl.Top = Ctrl.Top - hSpacing
            ElseIf Ctrl.Parent.Name = Ctrls(0).Parent.Name Then
                Ctrl.Top = Ctrl.Top - h
            End If
        End If
    Next
    Me.Height = Me.Height - h
    Parents = GetParents(Ctrls(0))
    For i = 0 To UBound(Parents)
        Parents(i).Height = Parents(i).Height - h
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


