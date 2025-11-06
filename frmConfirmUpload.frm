VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfirmUpload 
   Caption         =   "Uploading process"
   ClientHeight    =   1995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5985
   OleObjectBlob   =   "frmConfirmUpload.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConfirmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CurrentMod = "frmConfirmUpload"
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel button Clicked"
    Me.Hide
    TextBox1.value = ""
End Sub
Private Sub btnOk_Click()
    WriteLog 1, CurrentMod, "btnOk_Click", "Ok button Clicked"
    Me.Hide
End Sub
Private Sub TextBox1_Change()
    btnOk.Enabled = LCase(TextBox1.value) = "yes"
End Sub
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If LCase(TextBox1.value) = "yes" Then Me.Hide
End Sub
Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "Confirm Upload form initialized"
    RemoveCloseButton Me
    CenterUserform Me
End Sub
