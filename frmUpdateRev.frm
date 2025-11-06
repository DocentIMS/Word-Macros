VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUpdateRev 
   Caption         =   "Update Revision Level"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmUpdateRev.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUpdateRev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public Cancelled As Boolean, IsFinal As Boolean

Private Sub btnAddRev_Click()
    Cancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub btnFinal_Click()
    Cancelled = False
    IsFinal = True
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    IsFinal = False
    Cancelled = False
    CenterUserform Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True
    btnCancel_Click
End Sub
