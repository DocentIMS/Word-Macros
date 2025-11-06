VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFeedback 
   Caption         =   "Feedback"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6555
   OleObjectBlob   =   "frmFeedback.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CurrentMod = "frmFeedback"

Private Sub btnAttach_Click()
    WriteLog 1, CurrentMod, "btnAttach_Click", "Attach Button Clicked"
    On Error Resume Next
    TbFilePath.value = GetFile("Browse to text file", Environ("Userprofile") & "\desktop", False, CSVFiles + TxtFiles + AllFiles)(1)
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel Button Clicked"
    Unload Me
End Sub
Private Sub btnOk_Click()
    WriteLog 1, CurrentMod, "btnOk_Click", "Ok Button Clicked"
    Dim Title As String
    Title = IIf(tbTitle.value = "(Optional)", "", tbTitle.value)
    SendFeedback tbMain.text, CkResp.value, tbName.value, TbFilePath.value, ckLogToo.value, cbToolName.value, Title
    Unload Me
End Sub

Private Sub cbToolName_Change()
    btnOk.Enabled = cbToolName.ListIndex <> 0
End Sub

Private Sub tbTitle_Change()
    If tbTitle.value <> "(Optional)" Then tbTitle.ForeColor = 0 Else tbTitle.ForeColor = -2147483645
End Sub
Private Sub tbTitle_Enter()
    If tbTitle.value = "(Optional)" Then tbTitle.value = ""
End Sub

Private Sub tbTitle_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tbTitle.value = "" Then tbTitle.value = "(Optional)"
End Sub

'Private Sub UserForm_Activate()
'    If Len(tbName.Text) = 0 Then Unload Me
'End Sub
Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "Feedback Form Initialize"
    cbToolName.AddItem "(Select Tool)"
    cbToolName.AddItem "Scope Parser"
    cbToolName.AddItem "Documents Manager"
    cbToolName.AddItem "Command Statements"
    cbToolName.ListIndex = 0
    If IsProjectSelected Then tbName.text = Application.UserName Else Unload Me
    CenterUserform Me
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
    lbPrjHeader.BackColor = ProjectColorStr
    tbMain.SetFocus
End Sub
