VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditSections 
   Caption         =   "Edit Sections"
   ClientHeight    =   690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "frmEditSections.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditSections"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CurrentMod = "frmEditSections"
Private mDocType As String
Private Sub btnAddSection_Click()
    WriteLog 1, CurrentMod, "btnAddSection_Click", "Insert Header Button Clicked"
    InsertHeader mDocType
End Sub
Sub Display(DocType As String)
    mDocType = DocType
    Me.Show
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel Button Clicked"
    Unload frmSettings
    Unload Me
End Sub
Private Sub btnRemSection_Click()
    WriteLog 1, CurrentMod, "btnRemSection_Click", "Remove Header Clicked"
    RemoveHeader
End Sub
Private Sub BtnRun_Click()
    WriteLog 1, CurrentMod, "BtnRun_Click", "Application.run Button Clicked"
    Me.Hide
    frmSettings.Show
    frmSettings.UpdateCounts
End Sub
Private Sub btnUndo_Click()
    ActiveDocument.Undo
End Sub

Private Sub UserForm_Initialize()
    CenterUserform Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    WriteLog 1, CurrentMod, "UserForm_QueryClose", "Cancel Button Clicked"
    Unload frmSettings
End Sub
