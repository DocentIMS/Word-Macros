VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "New Project"
   ClientHeight    =   2340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PNum As Long

Private Sub btnAdd_Click()
    Dim ProjectURLStr As String, RespMsg As String
    ProjectURLStr = tbURL.value
    If Not ProjectURLStr Like "http*" Then ProjectURLStr = "https://" & ProjectURLStr
    RespMsg = IsValidUser(ProjectURLStr, tbUser.value, tbPassword.value)
    If RespMsg = "Ok" Then
        If DownloadProjectInfo(ProjectURLStr, tbUser.value, tbPassword.value) Then
            MsgBox "Project " & btnAdd.Caption & IIf(Right(btnAdd.Caption, 1) = "e", "d", "ed"), vbInformation, ""
            Unload Me
        Else
            MsgBox "The server setup is not complete. Please contact the project manager.", vbExclamation, ""
        End If
    Else
        MsgBox RespMsg, vbExclamation, ""
    End If
    RefreshRibbon
End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub
Private Sub tbUser_Change(): OkEnabled: End Sub
Private Sub tbPassword_Change(): OkEnabled: End Sub
Private Sub tbURL_Change(): OkEnabled: End Sub
Private Function OkEnabled() As Boolean
    Dim i As Long
    On Error GoTo ex
    OkEnabled = True
    For i = 1 To Controls.Count
        If Controls(i).Name Like "tb*" Then
            If Len(Controls(i).value) = 0 Then OkEnabled = False: Exit For
        End If
    Next
ex:
    btnAdd.Enabled = OkEnabled
End Function

Private Sub UserForm_Initialize()
    CenterUserform Me
End Sub
Sub Display(OkBtnCaption As String)
    btnAdd.Caption = OkBtnCaption
    Me.Show
End Sub
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    RefreshRibbon
'End Sub
