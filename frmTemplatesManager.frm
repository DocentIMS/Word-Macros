VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTemplatesManager 
   Caption         =   "Select a tempate"
   ClientHeight    =   930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmTemplatesManager.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTemplatesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnTemplate_Click()
    On Error Resume Next
    If IsProjectSelected(False) Then
        Me.Hide
        OpenTemplate cbTemplate.value, True
        Unload Me
    End If
End Sub

Private Sub cbTemplate_Change()
    TemplateNum = cbTemplate.ListIndex
    btnTemplate.Enabled = TemplateNum <> 0
End Sub
Private Sub UserForm_Initialize()
    On Error Resume Next
    CenterUserform Me
    cbTemplate.List = templateName
    cbTemplate.ListIndex = 0
'    Dim i As Long
'    For i = LBound(ctemplates) To UBound(ctemplates)
'
'    Next
End Sub
