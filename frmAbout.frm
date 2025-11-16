VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "About"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4845
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CurrentMod = "frmAbout"
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal fileName As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus _
        ) As Long
#Else
    Private Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal Filename As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus _
        ) As Long
#End If

Private Sub LbMail_Click()
    WriteLog 1, CurrentMod, "LbMail_Click", "Email Address Clicked"
    On Error Resume Next
    Dim s As String
    s = LbMail.Caption
    If Not s Like "Mailto:*" Then s = "Mailto:" & s
    ShellExecute 0, "open", s
End Sub
Private Sub UserForm_Initialize()
    WriteLog 1, CurrentMod, "UserForm_Initialize", "About Form Initialized"
    Dim VNum As String
    VNum = GetVerByReg
    If Len(VNum) = 0 Then VNum = Replace(Replace(ThisDocument.Name, "DocentIMS_", ""), ".dotm", "")
    LbVersion.Caption = Replace(LbVersion.Caption, "XX", VNum) 'GetVer(ThisDocument.Name))
    CenterUserform Me
End Sub
Private Function GetVer(FName As String) As Long
    WriteLog 1, CurrentMod, "GetVer", "Getting the version number"
    GetVer = Val(Right$(FName, Len(FName) - InStr(FName, "_")))
End Function
