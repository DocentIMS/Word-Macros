VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServerFilesList 
   Caption         =   "Duplicate files"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   OleObjectBlob   =   "frmServerFilesList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmServerFilesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mClickedButton As String
Private Const CurrentMod = "frmServerFilesList"
Private ListDup As New Collection
Private ListNew As New Collection
Private mMsg As String
Function ShowList(msg As String, Dups As Collection, News As Collection) As String
    WriteLog 1, CurrentMod, "DefineLists"
    Set ListDup = Dups
    Set ListNew = News
    mMsg = msg
    If ListNew.Count = 0 Then
        chListNews.Visible = False
        chListDuplicates.value = True
    ElseIf ListDup.Count > ListNew.Count Then
        chListNews.value = True
    Else
        chListDuplicates.value = True
    End If
    Me.Show
    ShowList = mClickedButton
    Unload Me
End Function
Private Sub UpdateList(Coll As Collection)
    WriteLog 1, CurrentMod, "ShowList", "Poplulating duplicates List"
    Dim i As Long
    LbMain.Caption = mMsg
    ListBox1.Clear
    For i = 1 To Coll.Count
        ListBox1.AddItem Coll(i)
    Next
End Sub

Private Sub btnReupload_Click()
    WriteLog 1, CurrentMod, "btnReupload_Click", "ReUpload Everything button clicked"
    mClickedButton = "ReUpload"
    Me.Hide
End Sub

Private Sub chListDuplicates_Change()
    If chListDuplicates.value Then
        WriteLog 1, CurrentMod, "chListDuplicates_Change"
        UpdateList ListDup
    End If
End Sub
Private Sub chListNews_Change()
    If chListNews.value Then
        WriteLog 1, CurrentMod, "chListNews_Change"
        UpdateList ListNew
    End If
End Sub
'Private Sub ckListDuplicates_Change()
'    WriteLog 1, CurrentMod, "ckListDuplicates_Change", "Upload Duplicates clicked"
'    HideControl Array(ListBox1), Not ckListDuplicates.Value
'End Sub
Private Sub HideControl(Ctrls, Optional Hide As Boolean = True)
    Dim t As Single, h As Single, tMin As Single, hSpacing As Single
    Dim i As Long, Ctrl As control
    Dim Parents
    tMin = Me.Height + Me.Top
    Const frmHSpace = 6
    Const OtherHSpacing = 0.5
    For i = LBound(Ctrls) To UBound(Ctrls)
        Set Ctrl = Ctrls(i)
        Ctrl.Visible = Hide = False
        t = GetTop(Ctrl)
        If tMin > t Then tMin = t
        If h < Ctrl.Height Then h = Ctrl.Height
        If Ctrl.Name Like "fra*" Then
            hSpacing = frmHSpace
        Else
            hSpacing = OtherHSpacing
        End If
    Next
    For Each Ctrl In Me.Controls
        If GetTop(Ctrl) > tMin Then
            If Ctrl.Parent.Name = Me.Name Then
                Ctrl.Top = IIf(Hide, Ctrl.Top - h, Ctrl.Top + h)
                Ctrl.Top = Ctrl.Top - hSpacing
            ElseIf Ctrl.Parent.Name = Ctrls(0).Parent.Name Then
                Ctrl.Top = IIf(Hide, Ctrl.Top - h, Ctrl.Top + h)
            End If
        End If
    Next
    Me.Height = IIf(Hide, Me.Height - h, Me.Height + h)
    Parents = GetParents(Ctrls(0))
    For i = 0 To UBound(Parents)
        Parents(i).Height = IIf(Hide, Parents(i).Height - h, Parents(i).Height + h)
    Next
End Sub
Private Function GetTop(ByVal Ctrl) As Single
    Dim Parents, i As Long
    Parents = GetParents(Ctrl)
    GetTop = Ctrl.Top
    For i = 0 To UBound(Parents)
        GetTop = GetTop + Parents(i).Top
    Next
End Function
Private Function GetParents(ByVal Ctrl) As Variant
    Dim Arr() As control, i As Long
    Do While Ctrl.Parent.Name <> Me.Name
        ReDim Preserve Arr(0 To i)
        Set Arr(i) = Ctrl.Parent
        Set Ctrl = Ctrl.Parent
        i = i + 1
    Loop
    GetParents = Arr
End Function

Private Sub btnDuplicates_Click()
    WriteLog 1, CurrentMod, "btnDuplicates_Click", "Upload Duplicates button clicked"
    mClickedButton = "Upload Duplicates"
    Me.Hide
End Sub
Private Sub btnNewFiles_Click()
    WriteLog 1, CurrentMod, "btnNewFiles_Click", "Upload New Files button clicked"
    mClickedButton = "Upload New Files Only"
    Me.Hide
End Sub
Private Sub btnCancel_Click()
    WriteLog 1, CurrentMod, "btnCancel_Click", "Cancel button clicked"
    mClickedButton = "Cancel"
    Me.Hide
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "List Form Initialize"
    CenterUserform Me
'    HideControl Array(ListBox1), True
End Sub
