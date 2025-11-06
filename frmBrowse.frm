VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBrowse 
   Caption         =   "Browse to a folder"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10005
   OleObjectBlob   =   "frmBrowse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddVariable_Click()
    If Right$(tbLoc.value, 1) <> "\" Then tbLoc.value = tbLoc.value & "\"
    tbLoc.value = tbLoc.value & "%" & cbVariables.value & "%\"
    VarsBtnsEnabled
End Sub
Private Sub btnRemVariable_Click()
    tbLoc.value = Replace(tbLoc.value, "%" & cbVariables.value & "%\", "")
    If Right$(tbLoc.value, 1) <> "\" Then tbLoc.value = tbLoc.value & "\"
    VarsBtnsEnabled
End Sub
Private Sub btnBrowse_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = Me.Caption
        .Show
        If .SelectedItems.Count Then
            tbLoc.value = MappedToServer(.SelectedItems(1))
'            btnAddVariable.Enabled = True
'            cbVariables.Enabled = True
        End If
    End With
End Sub
Private Function MappedToServer(Pth) As String
    Dim FSO, DriveLetter As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    DriveLetter = FSO.GetDriveName(FSO.GetAbsolutePathName(Pth))
    If Len(FSO.GetDrive(DriveLetter).ShareName) Then
        MappedToServer = FSO.GetDrive(DriveLetter).ShareName & Right$(Pth, Len(Pth) - Len(DriveLetter))
    Else
        MappedToServer = Pth
    End If
End Function
Private Sub btnOk_Click()
    Unload Me
'    Me.Hide
End Sub
Private Sub VarsBtnsEnabled()
    btnRemVariable.Enabled = InStr(tbLoc.value, cbVariables.value)
    btnAddVariable.Enabled = cbVariables.ListIndex > 0 And Not btnRemVariable.Enabled
End Sub
Private Sub cbVariables_Change(): VarsBtnsEnabled: End Sub

Private Sub tbLoc_Change()
    cbVariables.Enabled = Len(tbLoc.value)
End Sub

Private Sub UserForm_Initialize()
    cbVariables.AddItem "-- Choose Variable --"
    cbVariables.AddItem "Project Name"
    cbVariables.AddItem "Documents Type"
    cbVariables.AddItem "Document State"
    cbVariables.AddItem "Contract Number"
    cbVariables.AddItem "Date"
    cbVariables.AddItem "User Name"
    cbVariables.ListIndex = 0
'    btnAddVariable.Enabled = False
    cbVariables.Enabled = False
    CenterUserform Me
End Sub
Public Function Display(Optional Title As String, Optional InitalLocation As String) As String
    If Len(Title) Then Me.Caption = Title
    tbLoc.value = InitalLocation
'    btnAddVariable.Enabled = Len(InitalLocation)
'    cbVariables.Enabled = Len(InitalLocation)
    Me.Show
    Display = tbLoc.value
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
