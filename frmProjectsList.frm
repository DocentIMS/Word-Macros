VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProjectsList 
   Caption         =   "Projects Configuration"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   OleObjectBlob   =   "frmProjectsList.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProjectsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()
    frmLogin.Display "Add"
    RefreshList
End Sub
Private Function SelectedPName() As String
    Dim i As Long
    For i = 0 To lstProjects.ListCount - 1
        If lstProjects.Selected(i) Then
            SelectedPName = lstProjects.List(i)
            Exit For
        End If
    Next
End Function
Private Sub btnEdit_Click()
    EditProject SelectedPName
End Sub
Private Sub btnRemove_Click()
'    Dim i As Long,
    Dim mPName As String
    mPName = SelectedPName
    If Len(mPName) Then
        If frmMsgBox.Display("Remove Project?", Array("Remove", "Cancel"), Exclamation, "Docent IMS") = "Remove" Then
            RemovePFromReg mPName
            RefreshList
        End If
    End If
'    For i = 1 To UBound(ProjectName)
'        If ProjectName(i) = mPName Then
'            mPName = URL(i)
'            RemovePFromReg mPName
'            RefreshList
'            Exit For
'        End If
'    Next
End Sub
Private Sub btnUpdate_Click()
'    SetCursor LoadCursorW(0&, IDC_WAIT)
'    DoEvents
    UpdateAllProjectsInfo
'    Dim i As Long, mPName As String
'    For i = 0 To lstProjects.ListCount - 1
'        UpdateProject lstProjects.List(i)
'    Next
'    SetCursor LoadCursorW(0&, IDC_ARROW)
'    frmMsgBox.Display "All projects were updated."
End Sub

Private Sub lstProjects_Change()
    Dim i As Long
    For i = 1 To lstProjects.ListCount
        If lstProjects.Selected(i - 1) Then Exit For
    Next
    If i > lstProjects.ListCount Then
        btnRemove.Enabled = False
        btnEdit.Enabled = False
    Else
        btnRemove.Enabled = True
        btnEdit.Enabled = True
    End If
    btnUpdate.Enabled = lstProjects.ListCount > 0
End Sub

Private Sub lstProjects_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EditProject SelectedPName
End Sub
Private Sub EditProject(mPName As String)
    Dim i As Long
    For i = 1 To UBound(projectName)
        If projectName(i) = mPName Then
            'Me.Hide
            With frmLogin
                .Caption = mPName
                .tbPassword = Password(i)
                .tbURL = projectURL(i)
                .tbUser = UserName(i)
                .btnAdd.Caption = "Update"
                .Show
            End With
            RefreshList
            Me.Show
            Exit For
        End If
    Next
End Sub
Private Sub UpdateProject(mPName As String)
    Dim i As Long
    For i = 1 To UBound(projectName)
        If projectName(i) = mPName Then
'            If IsValidUser(URL(i), UserName(i), Password(i)) = "Ok" Then
                DownloadProjectInfo projectURL(i), UserName(i), Password(i)
                RefreshList
'            End If
            Exit For
        End If
    Next
End Sub
Private Sub UserForm_Initialize()
    CenterUserform Me
    RefreshList
    Me.Repaint
End Sub
Private Sub RefreshList()
    Dim i As Long
    lstProjects.Clear
'    LoadProjectInfoReg
    LoadProjects
    For i = 1 To UBound(projectName)
        lstProjects.AddItem projectName(i)
    Next
    lstProjects_Change
End Sub
