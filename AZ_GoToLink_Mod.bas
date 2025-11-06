Attribute VB_Name = "AZ_GoToLink_Mod"
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, _
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
Sub GoToLocation(ByVal Link)
    If Len(Link) = 0 Then Exit Sub
'    If Not Link Like "http*" Then Link = "http://" & Link
    ShellExecute 1, "open", Link, "", "", 1
End Sub

Sub GoToLink(ByVal Link)
    If Len(Link) = 0 Then Exit Sub
    If Link Like "?:\*" Then GoToLocation Link: Exit Sub
    If InStr(Link, ".") = 0 Then Link = ProjectURLStr & Link
    If Not Link Like "http*" Then Link = "http://" & Link
    ShellExecute 0, "open", Link
End Sub
Sub GoToEmail(ByVal Link)
    If Len(Link) = 0 Then Exit Sub
    If Not Link Like "*@*.*" Then Exit Sub
    ShellExecute 0, "open", "mailto:" & Link
End Sub
Sub OpenFileIn(ByVal Link, ByVal FName As String)
    If Len(Link) = 0 Then Exit Sub
    If Not Link Like "http*" Then Link = "http://" & Link
    FName = DownloadAPIFile(Link, False, FName)
    ShellExecute 0, "open", FName
End Sub
