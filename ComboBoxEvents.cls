VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboBoxEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'~~> This section is used for handling ComboBox Control Array
'
Public WithEvents control As MSForms.ComboBox
Attribute control.VB_VarHelpID = -1
Public Parent As Object

Private Sub Control_Change()
    Parent.ReFilter
End Sub
