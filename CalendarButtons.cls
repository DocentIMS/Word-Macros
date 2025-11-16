VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'~~> This section is used for handling Commandbutton Control Array
'
Public WithEvents control As MSForms.label
Attribute control.VB_VarHelpID = -1
Public Parent As frmETRcalendar
'~~> Unload the form when the user presses Escape
Private Sub Control_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then Unload Parent
End Sub
'~~> This section delas with showing/displaying controls
'~~> and updating different labels
Private Sub Control_Click()
    Parent.SelectedDate = ""
    If control.Tag <> "" Then
        Parent.SelectedDate = control.Tag
        Parent.Hide
    End If
    Select Case Left(control.Name, 1)
    Case "Y"
        If Len(Trim(control.Caption)) <> 0 Then
            CurYear = Val(control.Caption)
            With frmETRcalendar
                .MPmainDisplay.value = 1
                .lblTitleCurMY.Caption = CurYear
                .Label5.Caption = 2
                .lblUP.Visible = False
                .lblDOWN.Visible = False
            End With
        End If
    Case "M"
        Select Case control.Caption
        Case "JAN": CurMonth = 1
        Case "FEB": CurMonth = 2
        Case "MAR": CurMonth = 3
        Case "APR": CurMonth = 4
        Case "MAY": CurMonth = 5
        Case "JUN": CurMonth = 6
        Case "JUL": CurMonth = 7
        Case "AUG": CurMonth = 8
        Case "SEP": CurMonth = 9
        Case "OCT": CurMonth = 10
        Case "NOV": CurMonth = 11
        Case "DEC": CurMonth = 12
        End Select
        frmETRcalendar.ShowSpecificMonth
    End Select
End Sub
Private Sub Control_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    Dim i As Byte, Prefix As String, xMax As Long, xMin As Long, yMax As Long, yMin As Long
    Select Case control.Name
    Case "lblTitleCurMY", "lblUP", "lblDOWN", "lblUnload", "lblThemes"
        xMax = 68: xMin = 2: yMax = 10: yMin = 2
        If x >= xMax Or x <= xMin Or Y <= yMin Or Y >= yMax Then
            control.ForeColor = MyForeColor
        Else
            control.ForeColor = RGB(73, 255, 60)
        End If
    Case Else
        Prefix = Left$(control.Name, 1)
        Select Case Prefix
        Case "W"
            i = 0
            xMax = Day_Xmax: xMin = Day_Xmin: yMax = Day_Ymax: yMin = Day_Ymin
            Prefix = "WD"
        Case "D"
            i = Right$(control.Name, Len(control.Name) - 1)
            xMax = Day_Xmax: xMin = Day_Xmin: yMax = Day_Ymax: yMin = Day_Ymin
        Case "M", "Y"
            i = Right$(control.Name, Len(control.Name) - 1)
            xMax = Mo_Xmax: xMin = Mo_Xmin: yMax = Mo_Ymax: yMin = Mo_Ymin
        End Select
        If x >= xMax Or x <= xMin Or Y <= yMin Or Y >= yMax Then
            control.BorderStyle = fmBorderStyleNone
        Else
            control.BorderStyle = fmBorderStyleSingle
            NoBorder i, Prefix
        End If
    End Select
End Sub
Sub NoBorder(SkipLabel As Byte, Prefix As String)
    Dim D As Byte, ObjCnt As Byte
    Select Case Prefix
    Case "D", "W": ObjCnt = 42
    Case "M", "Y": ObjCnt = 12
    End Select
    For D = 1 To ObjCnt
        If D <> SkipLabel Then Parent.Controls(Prefix & D).BorderStyle = fmBorderStyleNone
    Next
End Sub
