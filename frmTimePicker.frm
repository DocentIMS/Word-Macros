VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimePicker 
   Caption         =   "Time Picker"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2865
   OleObjectBlob   =   "frmTimePicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCCName As String
Private Evs As CtrlEvents
Property Let CCName(ContentControlName As String)
    mCCName = ContentControlName
End Property
Private Sub btnSave_Click()
    SetContentControl mCCName, Format(TimeSerial(tbhTime + IIf(cbttTime.value = "AM", 0, 12), tbmTime, 0), timeFormat)
    Unload Me
End Sub
'Private Sub sbHours_Change()
'    tbHours.Value = sbHours.Value
'End Sub
'Private Sub sbminutes_Change()
'    tbMinutes.Value = sbMinutes.Value
'End Sub
'Private Sub UserForm_Activate()
'    cbAMPM.AddItem "AM"
'    cbAMPM.AddItem "PM"
'End Sub
Sub Load(Optional OldTime As String)
    Dim ss() As String
    If Len(OldTime) = 0 Then OldTime = Now
    ss = Split(Format(OldTime, "h:m:AM/PM"), ":")
    tbhTime.value = ss(0)
    sbhTime.value = ss(0)
    tbmTime.value = ss(1)
    sbmTime.value = ss(1)
    cbttTime.value = ss(2)
    Me.Show
End Sub
Private Sub UserForm_Initialize()
    Set Evs = New CtrlEvents
    Set Evs.Parent = Me
    Evs.AddOkButton btnSave
    Evs.MakeRequired "Time", , ErrorColor
End Sub
