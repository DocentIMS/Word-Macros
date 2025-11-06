Attribute VB_Name = "AZ_Calendar_Mod"
'Option Explicit
'Option Private Module
'
''Public Const GWL_STYLE = -16
''Public Const WS_CAPTION = &HC00000
'
'#If VBA7 Then
'    #If Win64 Then
'        Public Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias _
'        "GetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
'
'        Public Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias _
'        "SetWindowLongPtrA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
'        ByVal dwNewLong As LongPtr) As LongPtr
'    #Else
'        Public Declare Function GetWindowLongPtr Lib "user32" Alias _
'        "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr
'
'        Private Declare Function SetWindowLongPtr Lib "user32" Alias _
'        "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, _
'        ByVal dwNewLong As LongPtr) As LongPtr
'    #End If
'
'    Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
'    (ByVal hwnd As LongPtr) As LongPtr
'
'    Private Declare PtrSafe Function FindWindow Lib "user32" Alias _
'    "FindWindowA" (ByVal lpClassName As String, _
'    ByVal lpWindowName As String) As LongPtr
'
'    Private Declare PtrSafe Function SetTimer Lib "user32" _
'    (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
'    ByVal uElapse As LongPtr, ByVal lpTimerFunc As LongPtr) As LongPtr
'
'    Public Declare PtrSafe Function KillTimer Lib "user32" _
'    (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr
'
'    Public TimerID As LongPtr
'
'    Dim lngWindow As LongPtr, lFrmHdl As LongPtr
'#Else
'
'    Public Declare Function GetWindowLong _
'    Lib "user32" Alias "GetWindowLongA" ( _
'    ByVal hwnd As Long, ByVal nIndex As Long) As Long
'
'    Public Declare Function SetWindowLong _
'    Lib "user32" Alias "SetWindowLongA" ( _
'    ByVal hwnd As Long, ByVal nIndex As Long, _
'    ByVal dwNewLong As Long) As Long
'
'    Public Declare Function DrawMenuBar _
'    Lib "user32" (ByVal hwnd As Long) As Long
'
'    Public Declare Function FindWindowA _
'    Lib "user32" (ByVal lpClassName As String, _
'    ByVal lpWindowName As String) As Long
'
'    Public Declare Function SetTimer Lib "user32" ( _
'    ByVal hwnd As Long, ByVal nIDEvent As Long, _
'    ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'
'    Public Declare Function KillTimer Lib "user32" ( _
'    ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'
'    Public TimerID As Long
'    Dim lngWindow As Long, lFrmHdl As Long
'#End If

Public TimerSeconds As Single, tim As Boolean
Public CurMonth As Integer, CurYear As Integer
Public frmYr As Integer, ToYr As Integer

'Public F As frmETRcalendar

Enum CalendarThemes
    Venom = 0
    MartianRed = 1
    ArcticBlue = 2
    Greyscale = 3
End Enum

Public Const Day_Xmax = 20
Public Const Day_Xmin = 2
Public Const Day_Ymax = 15.75
Public Const Day_Ymin = 2

Public Const Mo_Xmax = 38
Public Const Mo_Xmin = 2
Public Const Mo_Ymax = 15.75
Public Const Mo_Ymin = 2

Public MyBackColor As Long, MyForeColor As Long, CurDateColor As Long, CurDateForeColor As Long, NotCurDateColor As Long

'Sub Launch() '(control As IRibbonControl)
'
'    With frmETRcalendar
'        .Caltheme = ShForCode.Cells(8, 2)
'        .LongDateFormat = "dddd, mmmm dd" '"mmmm dddd dd, yyyy" '"dddd dd. mmmm yyyy" ' etc
'        .ShortDateFormat = GetDateFormat ' "mm/dd/yyyy" 'or "d/m/y" etc
'        .Show
'    End With
'
'End Sub

'~~> Hide the title bar of the userform

Sub StartTimer()
    '~~ Set the timer for 1 second
    TimerSeconds = 1
    TimerID = SetTimer(0&, 0&, TimerSeconds * 1000&, AddressOf TimerProc)
    'Debug.Print TimerID
End Sub

Sub EndTimer()
    On Error Resume Next
    KillTimer 0&, TimerID
End Sub
    
'~~> Update Time
#If VBA7 And Win64 Then ' 64 bit Excel under 64-bit windows  ' Use LongLong and LongPtr
    Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As LongLong, _
    ByVal nIDEvent As LongPtr, ByVal dwTimer As LongLong)
        TimerProcA
    End Sub
#ElseIf VBA7 Then ' 64 bit Excel in all environments
    Public Sub TimerProc(ByVal hwnd As LongPtr, ByVal uMsg As Long, _
    ByVal nIDEvent As LongPtr, ByVal dwTimer As Long)
        TimerProcA
    End Sub
#End If
Private Sub TimerProcA()
    If IsLoaded("frmETRcalendar") Then
        frmETRcalendar.lblTitleClock.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(0)
        frmETRcalendar.lblTitleAMPM.Caption = Split(Format(Time, "h:mm:ss AM/PM"))(1)
    Else
        EndTimer
    End If
End Sub
Function GetDateFormat() As String
'    GetDateFormat = DateFormat
    GetDateFormat = Application.International(17)
    Select Case Application.International(32)
        Case 0: GetDateFormat = "m" & GetDateFormat & "d" & GetDateFormat & "yy"
        Case 1: GetDateFormat = "d" & GetDateFormat & "m" & GetDateFormat & "yy"
        Case 2: GetDateFormat = "yy" & GetDateFormat & "m" & GetDateFormat & "d"
    End Select
    If Application.International(43) Then GetDateFormat = Replace(GetDateFormat, "yy", "yyyy")
    If Application.International(41) Then GetDateFormat = Replace(GetDateFormat, "m", "mm")
    If Application.International(42) Then GetDateFormat = Replace(GetDateFormat, "d", "dd")
End Function
Public Function IsLoaded(formName As String) As Boolean
    Dim Frm As Object
    For Each Frm In VBA.UserForms
        If Frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next Frm
    IsLoaded = False
End Function
