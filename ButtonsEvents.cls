VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ButtonsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'~~> This section is used for handling Commandbutton Control Array
'
Public WithEvents btn As MSForms.CommandButton
Attribute btn.VB_VarHelpID = -1
Public WithEvents lb As MSForms.label
Attribute lb.VB_VarHelpID = -1
Public Parent As Object

Private Sub btn_Click(): Parent.ClickedButton btn.Caption: End Sub
Private Sub lb_Click(): Parent.ClickedLabel lb.Tag: End Sub

