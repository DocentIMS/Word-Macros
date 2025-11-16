Attribute VB_Name = "DocentActions_Feedback"
Option Explicit
Private Const CurrentMod = "DocentActions_Feedback"
Sub SendFeedback(MainFeedback As String, NeedsResponse As String, _
        UserName As String, Attachment As String, LogToo As Boolean, SelectedTool As String, Optional Title As String)
    Dim Resp As WebResponse, LogFName As String
    WriteLog 1, CurrentMod, "btnOk_Click", "SendFeedback"
    AttachZip = ""
    LogZip = ""
    If LogToo Then ZipLog
    If Len(Attachment) Then
        CreateZipFile "Att"
        ZipFiles AttachZip, Attachment
    End If
    If Len(Title) = 0 Then Title = "Feedback " & Format(ToServerTime, "yy-mm-dd")
    Set Resp = SendAPIFeedback(Title, MainFeedback, _
                SelectedTool, NeedsResponse, UserName, AttachZip, LogZip, "/feedback/")
    If IsGoodResponse(Resp) Then
        frmMsgBox.Display "Thank you for the feedback."
    Else
        MsgBox "We could not send feedback. Please email us on wglover@docentims.com", vbCritical, ""
    End If
    DeleteZip
End Sub


