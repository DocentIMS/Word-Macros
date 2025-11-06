VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreatePostItNote 
   Caption         =   "Create PostIt Note"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8085
   OleObjectBlob   =   "frmCreatePostItNote.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreatePostItNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CurrentMod = "frmCreatePostItNote"
Private Evs As New CtrlEvents

Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel Button Clicked"
    Unload Me
End Sub
Private Sub btnCreate_Click()
    Dim Resp As WebResponse
    Set Resp = CreateAPIContent("postit_note", "notes", _
                            Array("color", "description", "file"), _
                            Array(cbColor.value, tbNoteText.value, TbFilePath.value))
    If IsGoodResponse(Resp) Then
        frmMsgBox.Display Array("A new Posted It Note was created on " & ProjectNameStr & " site.", " ", , "View Online"), , Success, "DocentIMS", , , Array(, , , Resp.Data("@id"))
        Unload Me
    Else
        frmMsgBox.Display "Could not post it", , Critical, "DocentIMS"
    End If
End Sub
Private Sub btnAttach_Click()
    WriteLog 1, CurrentMod, "btnAttach_Click", "Attach Button Clicked"
    On Error Resume Next
    TbFilePath.value = GetFile("Browse to text file", Environ("Userprofile") & "\desktop", False, CSVFiles + TxtFiles + AllFiles)(1)
End Sub

'Private Sub tbNoteText_Change()
'
'End Sub
'
'Private Sub tbNoteText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyReturn: btnOK_Click
'End Sub

Private Sub UserForm_Activate()
    ResizableForm Me
End Sub

Private Sub UserForm_Initialize()
    Dim Coll As Collection, i As Long
    Set Coll = GetAPIPostItNoteColors()
    For i = 1 To Coll.Count
        cbColor.AddItem Coll(i)(1)
    Next
    cbColor.value = "yellow"
    Set Evs.Parent = Me
    Evs.AddOkButton btnCreate
    Evs.MakeRequired "tbNoteText", , ErrorColor
'    CenterUserform Me
    lbPrjHeader.Caption = ProjectNameStr
    lbPrjHeader.ForeColor = FullColor(ProjectColorStr).Inverse
    lbPrjHeader.BackColor = ProjectColorStr
End Sub

Private Sub UserForm_Resize()
    btnCreate.Left = (Me.Width - btnCreate.Width) / 2
    lbColor.Top = Me.Height - 114
    cbColor.Top = Me.Height - 117
    btnCreate.Top = Me.Height - 57
    btnAttach.Top = Me.Height - 81
    lbAttach.Top = Me.Height - 78
    TbFilePath.Top = Me.Height - 81
    
    cbColor.Width = Me.Width - 92.25
    tbNoteText.Width = Me.Width - 26.25
    TbFilePath.Width = Me.Width - 133.75
    
    tbNoteText.Height = Me.Height - 135
    
    btnAttach.Left = Me.Width - 62.23
End Sub
