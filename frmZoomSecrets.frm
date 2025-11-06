VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZoomSecrets 
   Caption         =   "UserForm3"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11055
   OleObjectBlob   =   "frmZoomSecrets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZoomSecrets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOk_Click()
    SetZoomToReg tbAccountID.value, tbClientID.value, tbClientSecret.value
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    tbAccountID.value = GetZoomAccountID
    If Len(tbAccountID.value) = 0 Then tbAccountID.value = "SVoPCc5NSqCyL86gXc7P8w"
    tbClientID.value = GetZoomClientID
    If Len(tbClientID.value) = 0 Then tbClientID.value = "BupIUn1TaGXWGR935HAw"
    tbClientSecret.value = GetZoomClientSecret
    If Len(tbClientSecret.value) = 0 Then tbClientSecret.value = "1MQ4JKaOt3icCHOku9P4cfzI7KxAS6mQ"
End Sub
