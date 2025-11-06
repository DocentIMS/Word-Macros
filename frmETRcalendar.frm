VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmETRcalendar 
   Caption         =   "frmETRcalendar"
   ClientHeight    =   5895
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "frmETRcalendar.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmETRcalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'24-12-11-6-46
Option Explicit
Public SelectedDate As String
Private DateSet As Boolean
Private curDate As Date
Private i As Long
Private thisDay As Integer, thisMonth As Integer, thisYear As Integer
Private ButtonsArray() As New CalendarButtons
Private NewXpos As Single
Private NewYpos As Single
Private Cal_theme As CalendarThemes
Private LdtFormat As String, SdtFormat As String
Public Property Let SetDate(mDateStr As String)
    Me.LongDateFormat = "dddd, mmm dd, yyyy"
    Me.ShortDateFormat = DateFormat 'GetDateFormat
    If Len(mDateStr) Then
        curDate = ParseDate(mDateStr, Me.ShortDateFormat)
    Else
        curDate = Date
    End If
    thisDay = Day(curDate): thisMonth = Month(curDate): thisYear = Year(curDate)
    CurYear = Year(curDate): CurMonth = Month(curDate)
    PopulateCalendar curDate
    DateSet = True
End Property
Public Property Let LongDateFormat(s As String)
    LdtFormat = s
    lblTitleCurDt.Caption = Format(curDate, LdtFormat)
End Property
Public Property Get LongDateFormat() As String
    LongDateFormat = LdtFormat
End Property
Public Property Let ShortDateFormat(s As String)
    SdtFormat = s
End Property
Public Property Get ShortDateFormat() As String
    ShortDateFormat = SdtFormat
End Property
Public Property Let Caltheme(Theme As CalendarThemes)
    Cal_theme = Theme
    '--> Set the color of controls
    Select Case Cal_theme
        Case CalendarThemes.Venom
            MyBackColor = RGB(69, 69, 69)
            MyForeColor = RGB(252, 248, 248)
            CurDateColor = RGB(246, 127, 8)
            CurDateForeColor = RGB(0, 0, 0)
            NotCurDateColor = RGB(90, 90, 90)
        Case CalendarThemes.MartianRed
            MyBackColor = RGB(87, 0, 0)
            MyForeColor = RGB(203, 146, 146)
            CurDateColor = RGB(122, 185, 247)
            CurDateForeColor = RGB(0, 0, 0)
            NotCurDateColor = RGB(116, 0, 0)
        Case CalendarThemes.ArcticBlue
            MyBackColor = RGB(42, 48, 92)
            MyForeColor = RGB(179, 179, 179)
            CurDateColor = RGB(122, 185, 247)
            CurDateForeColor = RGB(0, 0, 0)
            NotCurDateColor = RGB(66, 71, 118)
        Case CalendarThemes.Greyscale
            MyBackColor = RGB(240, 240, 240)
            MyForeColor = RGB(0, 0, 0)
            CurDateColor = RGB(246, 127, 8)
            CurDateForeColor = RGB(0, 0, 0)
            NotCurDateColor = RGB(225, 225, 225)
    End Select
    Me.BackColor = MyBackColor
    FrameDay.BackColor = MyBackColor
    FrameMonth.BackColor = MyBackColor
    FrameYr.BackColor = MyBackColor
        
    lblTitleCurDt.ForeColor = CurDateColor
        
    lblTitleCurMY.ForeColor = MyForeColor
    lblTitleCurMY.BorderColor = MyForeColor
        
    lblTitleClock.ForeColor = MyForeColor
    lblTitleAMPM.ForeColor = MyForeColor
    lblUnload.ForeColor = MyForeColor
    lblThemes.ForeColor = MyForeColor
        
    lblUP.ForeColor = MyForeColor
    lblDOWN.ForeColor = MyForeColor
    '--> Days
    For i = 1 To 42
        With Me.Controls("D" & i)
            .ForeColor = MyForeColor
            .BorderColor = MyForeColor
        End With
    Next i
    '--> Weekdays
    For i = 1 To 7
        With Me.Controls("WD" & i)
            .ForeColor = MyForeColor
        End With
    Next i
    '--> Month
    For i = 1 To 12
        With Me.Controls("M" & i)
            .ForeColor = MyForeColor
            .BorderColor = MyForeColor
        End With
    Next i
    '--> Year
    For i = 1 To 12
        With Me.Controls("Y" & i)
            .ForeColor = MyForeColor
            .BorderColor = MyForeColor
        End With
    Next i
    '--> Populate this months calendar
    PopulateCalendar curDate
End Property
Public Property Get Caltheme() As CalendarThemes
    Caltheme = Cal_theme
End Property
'--> allow user to cycle thru avialable themes
Private Sub lblThemes_Click()
    Dim t As Byte
    t = GetSetting("PontalBrazilFinances", "Calendar", "Theme", 0) 'ShForCode.Cells(8, 2)
    t = t + 1
    If t > 3 Then t = 0
    Me.Caltheme = t
    Me.Repaint
    SaveSetting "PontalBrazilFinances", "Calendar", "Theme", t
End Sub
'--> Unload form
Private Sub lblUnload_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    '--> remove borders from day labels. i keep them in place for the dev environment.
    Dim lblCtrl As control
    i = 0
    For Each lblCtrl In Me.Controls
        If TypeOf lblCtrl Is MSForms.label Then
            lblCtrl.BorderStyle = fmBorderStyleNone
        End If
    Next
    '--> Hide the Title Bar
    HideTitleBar Me
    CenterUserform Me
    '--> Create a command button control array so that
    '--> when we press escape, we can unload the userform
    Dim CBCtl As control
    i = 0
    For Each CBCtl In Me.Controls
        If TypeOf CBCtl Is MSForms.label Then
            i = i + 1
            ReDim Preserve ButtonsArray(1 To i)
            Set ButtonsArray(i).control = CBCtl
            Set ButtonsArray(i).Parent = Me
        End If
    Next CBCtl
    Set CBCtl = Nothing
    '~~> Set the Time
    frmETRcalendar.lblTitleClock.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
    frmETRcalendar.lblTitleAMPM.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
    Me.ShortDateFormat = DateFormat 'GetDateFormat
    SetDate = Format(Date, Me.ShortDateFormat)
    Me.LongDateFormat = "dddd, mmm dd, yyyy"
    frmETRcalendar.Caltheme = GetSetting("PontalBrazilFinances", "Calendar", "Theme", 0)
    'StartTimer
    'PopulateCalendar curDate
End Sub
'--> The below 4 procedures will assist in moving the borderless userform
Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = 1 Then
        NewXpos = x
        NewYpos = Y
    End If
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button And 1 Then
        Me.Left = Me.Left + (x - NewXpos)
        Me.Top = Me.Top + (Y - NewYpos)
    End If
    lblDOWN.ForeColor = MyForeColor
    lblUP.ForeColor = MyForeColor
    lblUnload.ForeColor = MyForeColor
    lblThemes.ForeColor = MyForeColor
End Sub
Private Sub Frame1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    If Button = 1 Then
        NewXpos = x
        NewYpos = Y
    End If
End Sub
'--> Stop timer in the terminate event
Private Sub UserForm_Terminate()
    EndTimer
End Sub
'--> UP Button
Private Sub lblUP_Click()
    Select Case Label5.Caption
        Case 1 '~~> When user presses the up button when the dates are displayed
            curDate = DateSerial(CurYear, CurMonth, 0)
            '~~> Check if date is >= 1/1/1919
            If curDate >= DateSerial(1919, 1, 1) Then
                '~~> Populate prev months calendar
                PopulateCalendar curDate
            End If
        Case 2 '<~~ Do nothing
        Case 3 '~~> When user presses the up button when the Year Range is displayed
            If frmYr > 1919 Then
                Dim NewToYr As Integer
                ToYr = frmYr - 1
                NewToYr = frmYr - 1
                For i = 1 To 12
                    Me.Controls("Y" & i).Caption = ""
                Next i
                For i = 12 To 1 Step -1
                    If Not NewToYr < 1919 Then
                        With Me.Controls("Y" & i)
                            .Caption = NewToYr
                            If NewToYr = thisYear Then
                                .BackStyle = fmBackStyleOpaque
                                .BackColor = CurDateColor
                            Else
                                .BackStyle = fmBackStyleTransparent
                            End If
                            NewToYr = NewToYr - 1
                        End With
                    End If
                Next i
                frmYr = NewToYr + 1
                lblTitleCurMY.Caption = (NewToYr + 1) & " - " & ToYr
            End If
    End Select
End Sub
'--> Down Button
Private Sub lblDOWN_Click()
    Select Case Label5.Caption
        Case 1 '~~> When user presses the down button when the dates are displayed
            curDate = DateAdd("m", 1, DateSerial(CurYear, CurMonth, 1))
            '~~> Check if date is <= 31/12/2119
            If curDate <= DateSerial(2119, 12, 31) Then
                '~~> Populate prev months calendar
                PopulateCalendar curDate
            End If
        Case 2 '<~~ Do nothing
        Case 3 '~~> When user presses the down button when the Year Range is displayed
            frmYr = Val(Split(lblTitleCurMY.Caption, "-")(0))
            ToYr = Val(Split(lblTitleCurMY.Caption, "-")(1))
            If ToYr < 2119 Then
                Dim NewFrmYr As Integer
                frmYr = ToYr + 1
                NewFrmYr = ToYr + 1
                For i = 1 To 12
                    Me.Controls("Y" & i).Caption = ""
                Next i
                For i = 1 To 12
                    If NewFrmYr < 2119 Then
                        With Me.Controls("Y" & i)
                            .Caption = NewFrmYr
                            If NewFrmYr = thisYear Then
                                .BackStyle = fmBackStyleOpaque
                                .BackColor = CurDateColor
                            Else
                                .BackStyle = fmBackStyleTransparent
                            End If
                            NewFrmYr = NewFrmYr + 1
                        End With
                    ElseIf NewFrmYr = 2119 Then
                        With Me.Controls("Y" & i)
                            .Caption = NewFrmYr
                            NewFrmYr = NewFrmYr + 1
                        End With
                    End If
                Next i
                If NewFrmYr = 2119 Then ToYr = NewFrmYr Else ToYr = NewFrmYr - 1
                lblTitleCurMY.Caption = frmYr & " - " & ToYr
            End If
    End Select
End Sub
'--> Populate the calendar for a specific month
Sub PopulateCalendar(D As Date)
    Dim m As Integer, Y As Integer
    Dim i As Integer, j As Integer
    Dim LastDay As Integer, NextCounter As Integer, PrevCounter As Integer
    Dim dtOne As Date, dtLast As Date, dtNext As Date
    CurYear = Year(D)
    CurMonth = Month(D)
    m = Month(D): Y = Year(D)
    '--> 1st day of the current month
    dtOne = DateSerial(Y, m, 1)
    '--> last day of the previous month
    dtLast = dtOne - 1
'    dtLast = DateSerial(Year(dtOne), Month(dtOne), 0)
    '--> 1st day of the next month
    dtNext = DateSerial(Y, m + 1, 1)
'    dtNext = DateAdd("m", 1, DateSerial(Year(dtOne), Month(dtOne), 1))
    '--> Get the last day of the current month
    LastDay = Day(dtNext - 1) 'Val(Format(DateSerial(Year(dtOne), Month(dtOne) + 1, 0), "dd"))
    
    '--> Set the 1st day of the month to its proper weekday
'    NextCounter = Weekday(dtOne, 0)
'    PrevCounter = NextCounter - 1
    Select Case Weekday(dtOne, 0)
        Case 1: NextCounter = 1: PrevCounter = 0
        Case 2: NextCounter = 2: PrevCounter = 1
        Case 3: NextCounter = 3: PrevCounter = 2
        Case 4: NextCounter = 4: PrevCounter = 3
        Case 5: NextCounter = 5: PrevCounter = 4
        Case 6: NextCounter = 6: PrevCounter = 5
        Case 7: NextCounter = 7: PrevCounter = 6
    End Select
    '--> Populate all days for the current month
    For i = 1 To LastDay
        Me.Controls("D" & NextCounter).Caption = i
        Me.Controls("D" & NextCounter).Tag = Format(DateSerial(Year(D), Month(D), i), frmETRcalendar.ShortDateFormat)
        '--> Highlight the current day
        If i = thisDay And Month(D) = thisMonth And Year(D) = thisYear Then
            With Me.Controls("D" & NextCounter)
                .BackStyle = fmBackStyleOpaque
                .BackColor = CurDateColor
                .ForeColor = CurDateForeColor
            End With
        Else '--> no highlight
            With Me.Controls("D" & NextCounter)
                .BackStyle = fmBackStyleTransparent
                .BackColor = MyBackColor
                .ForeColor = MyForeColor
            End With
            '*** KEEP JUST IN CASE
            '                Select Case Cal_theme
            '                    Case CalendarThemes.ArcticBlue
            '                        Me.Controls("D" & NextCounter).BackColor = CurDateColor
            '                        Me.Controls("D" & NextCounter).ForeColor = RGB(0, 0, 0)
            '                    Case Else
            '                        Me.Controls("CB" & NextCounter).ForeColor = RGB(0, 0, 0)
            '                End Select
                
            '********
        End If
        NextCounter = NextCounter + 1
    Next i
    '--> Populate days for the next month
    j = 1
    If NextCounter < 43 Then
        For i = NextCounter To 42
            With Me.Controls("D" & i)
                .Caption = j
                .Tag = Format(DateSerial(Year(dtNext), Month(dtNext), j), frmETRcalendar.ShortDateFormat)
                .ForeColor = NotCurDateColor
            End With
            j = j + 1
        Next i
    End If
    'Populate days of previous month
    LastDay = Val(Format(dtLast, "dd"))
    If PrevCounter > 1 Then
        For i = PrevCounter To 1 Step -1
            With Me.Controls("D" & i)
                .Caption = LastDay
                .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmETRcalendar.ShortDateFormat)
                .ForeColor = NotCurDateColor
            End With
            LastDay = LastDay - 1
        Next i
    ElseIf PrevCounter = 1 Then
        With Me.Controls("D1")
            .Caption = LastDay
            .Tag = Format(DateSerial(Year(dtLast), Month(dtLast), LastDay), frmETRcalendar.ShortDateFormat)
            .ForeColor = NotCurDateColor
        End With
    End If
    lblTitleCurMY.Caption = Format(D, "mmmm yyyy")
End Sub
'--> Show the months when user clicks on the date label
Sub HiglightCurMonthControl()
    For i = 1 To 12
            
        If i = thisMonth Then
            With Me.Controls("M" & i)
                .BackStyle = fmBackStyleOpaque
                .BackColor = CurDateColor
                .ForeColor = CurDateForeColor
            End With
        End If
    Next i
End Sub
'--> Show the details for the selected month
Sub ShowSpecificMonth()
    lblTitleCurMY.Caption = Format(DateSerial(CurYear, CurMonth, 1), "mmm yyyy")
    MPmainDisplay.value = 0 'switch multipage back to 'Day' page
    PopulateCalendar DateSerial(CurYear, CurMonth, 1)
    Label5.Caption = 1
    lblUP.Visible = True
    lblDOWN.Visible = True
End Sub
'--> Handles the month to year multipage display
Private Sub lblTitleCurMY_Click()
    Select Case Label5.Caption
        Case 1
            lblTitleCurMY.Caption = Split(lblTitleCurMY.Caption)(1)
            Label5.Caption = 2
            Me.MPmainDisplay.value = 1 '--> Switch active multipage
            HiglightCurMonthControl
            lblDOWN.Visible = False
            lblUP.Visible = False
        Case 2 '--> Prep & show year buttons
            lblDOWN.Visible = True
            lblUP.Visible = True
            Me.MPmainDisplay.value = 2 '--> Switch active multipage
            ToYr = Val(lblTitleCurMY.Caption)
            frmYr = ToYr - 11
            If frmYr < 1919 Then frmYr = 1919
            lblTitleCurMY.Caption = frmYr & " - " & ToYr
            Label5.Caption = 3
            For i = 1 To 12
                Me.Controls("Y" & i).Caption = ""
            Next i
            For i = 12 To 1 Step -1
                If Not ToYr < 1919 Then
                    With Me.Controls("Y" & i)
                        .Caption = ToYr
                        .Visible = True
                            
                        If ToYr = thisYear Then
                            With Me.Controls("Y" & i)
                                .BackStyle = fmBackStyleOpaque
                                .BackColor = CurDateColor
                                .ForeColor = CurDateForeColor
                            End With
                        End If
                            
                        ToYr = ToYr - 1
                    End With
                End If
            Next i
            Label5.Caption = 3
        Case 3 'Do Nothing
    End Select
End Sub
