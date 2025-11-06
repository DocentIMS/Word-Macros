VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RequiredFieldsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tb As MSForms.TextBox
Attribute tb.VB_VarHelpID = -1
Private WithEvents sb As MSForms.ListBox
Attribute sb.VB_VarHelpID = -1
Private WithEvents cb As MSForms.ComboBox
Attribute cb.VB_VarHelpID = -1
Private WithEvents kb As MSForms.CheckBox
Attribute kb.VB_VarHelpID = -1
Private WithEvents op As MSForms.OptionButton
Attribute op.VB_VarHelpID = -1
Public mCtrl As MSForms.control
Public mOkBtnColl As New Collection ' MSForms.CommandButton
Public RequiredFields As Collection
Public Parent As MSForms.UserForm
Sub AddOkBtn(btn As CommandButton)
    mOkBtnColl.Add btn
End Sub
'Property Set OkBtn(Btn As CommandButton): Set mOkBtn = Btn: End Property
Property Get value(): On Error Resume Next: value = mCtrl.value: End Property
Private Sub sb_Change(): CheckOk: End Sub
Private Sub cb_Change(): CheckOk: End Sub
Private Sub kb_Change(): CheckOk: End Sub
Private Sub tb_Change(): CheckOk: End Sub
Sub Add(CtrlName As String)
    AddMark CtrlName
    Dim Ctrl As control
    With Parent
        For Each Ctrl In Parent.Controls
            If Ctrl.Name Like "*" & CtrlName Then
                Select Case TypeName(Ctrl)
                Case "ListBox": Set mCtrl = Ctrl: Set sb = Ctrl
                Case "ComboBox": Set mCtrl = Ctrl: Set cb = Ctrl
                Case "CheckBox": Set mCtrl = Ctrl: Set kb = Ctrl
                Case "TextBox": Set mCtrl = Ctrl: Set tb = Ctrl
                Case "OptionButton": Set mCtrl = Ctrl: Set op = Ctrl
                End Select
            End If
        Next
    End With
End Sub
Private Sub AddMark(CtrlName As String)
    Dim Ctrl As control
'    With Parent
        Set Ctrl = Parent.Controls("lb" & CtrlName)
        With Ctrl.Parent.Controls.Add("Forms.Label.1", "req" & CtrlName, True)
            .Font.Size = 10
            .Caption = "*"
            .ForeColor = vbRed
            .AutoSize = True
            .Top = Ctrl.Top
            .Left = Ctrl.Left
            Ctrl.Left = Ctrl.Left + .Width
        End With
'    End With
End Sub
Sub CheckOk()
    Dim i As Long, EnableOK As Boolean
    Dim IsValid As Boolean
'    If kb Is Nothing Then
'        If mCtrl.Value Like "-- Choose *" Then
'            mCtrl.ForeColor = xlGray
'        Else
'            mCtrl.ForeColor = vbBlack
'        End If
'    End If
    EnableOK = True
    For i = 1 To RequiredFields.Count
        IsValid = True
        Select Case TypeName(RequiredFields(i))
        Case "CheckBox", "OptionButton"
        Case Else
            IsValid = Len(RequiredFields(i).value) > 0 And Not RequiredFields(i).value Like "-- Choose *"
            RequiredFields(i).mCtrl.BackColor = IIf(IsValid, &H80000005, ErrorColor)
        End Select
        EnableOK = EnableOK And IsValid
    Next
'    EnableOK = i > RequiredFields.Count
    For i = 1 To mOkBtnColl.Count
        mOkBtnColl(i).Enabled = EnableOK
    Next
End Sub

