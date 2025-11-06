VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHelpPics 
   Caption         =   "Help Pictures"
   ClientHeight    =   11925
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18255
   OleObjectBlob   =   "frmHelpPics.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHelpPics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private i As Long, iMax As Long, Pth As String, mMode As Long
Private Sub btnNext_Click()
    'Dim Pth As String
    i = i + 1
    If i > iMax Then
        Unload Me
        Exit Sub
    End If
    LoadImageNo i
End Sub
Private Function LoadImageNo(i As Long) As Boolean
    imgPic.Picture = LoadPicture(Pth & i & ".jpg")
    lbPageNo.Caption = "Page " & i & " of " & iMax
    If i = iMax Then btnNext.Caption = "Finish"
End Function
Private Sub UserForm_Initialize()
'    CenterUserform Me
'    RemoveCloseButton Me
    HideTitleBar Me
End Sub
Sub Display(Optional Mode As Long = 1)
    mMode = Mode
    CreateDir ThisDocument.Path & "\Word Help Images\Team Member\"
    CreateDir ThisDocument.Path & "\Word Help Images\Project Manager\"
    Select Case Mode
    Case 1 'Team Member
        Pth = ThisDocument.Path & "\Word Help Images\Team Member\"
    Case 2 'Project Manager
        Pth = ThisDocument.Path & "\Word Help Images\Project Manager\"
    Case 3
        
    Case 4
    End Select
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    iMax = FSO.GetFolder(Pth).Files.Count
    If iMax Then
        i = 1
        LoadImageNo i
        'btnNext_Click
        Me.Show
    Else
        Unload Me
    End If
End Sub
    
'Private Sub UserForm_Click()
'    CommandButton1_Click
'End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If ckNeverAgain.value Then SetNeverHelpAgain mMode
    SetHelpShown mMode
End Sub
