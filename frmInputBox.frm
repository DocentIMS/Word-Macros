VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInputBox 
   ClientHeight    =   1710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5250
   OleObjectBlob   =   "frmInputBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Inpt As String
Private Sub btnCancel_Click()
    Inpt = "Canceled"
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Inpt = tbResponse.value
    Me.Hide
End Sub

Function Display(ByVal Prompt As String, Optional ByVal Title As String = "Microsoft Word", Optional ByVal Default As String, Optional PasswordMode) As String
    If IsMissing(PasswordMode) Then PasswordMode = InStr(Prompt, "password") + InStr(Title, "password") > 0
    tbResponse.PasswordChar = IIf(PasswordMode, "*", "")
    lbMsg.Caption = Prompt
    Me.Caption = Title
    tbResponse.value = Default
    Me.Show
    Display = Inpt
    Unload Me
End Function

Private Sub tbResponse_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case keyCode
    Case vbKeyReturn: btnOk_Click
    Case vbKeyEscape: btnCancel_Click
    End Select
End Sub
Private Sub UserForm_Initialize()
    CenterUserform Me
End Sub
