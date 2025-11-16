Attribute VB_Name = "DocentBreakdown_Upload"
Option Explicit
Private FSO As Object
Private ServerFiles
Private FailedFolder As String
Private DoneFolder As String
Private APIFoName As String
Private ServerFolderContents As Collection
Private DupColl As Collection
Private NewColl As Collection
Private Const CurrentMod = "DocentBreakdown_Upload"
Private Sub ResetFailed()
    WriteLog 1, CurrentMod, "ResetFailed"
    Dim FName As String
    FName = Dir(FailedFolder)
    Do While Len(FName)
        FSO.MoveFile FailedFolder & FName, HTMLPath & FName
        FName = Dir(FailedFolder)
    Loop
End Sub
Private Sub GetFilesLists(DocType As String, HTMLPath As String)
    WriteLog 1, CurrentMod, "GetServerFolderContents"
    Dim Description As String, fs As String, FName As String, Resp As WebResponse
    Dim i As Long
    Description = "These files were parsed from the project " & DocType & "." & _
            " Now, each section is analyzed and Tasks created." & _
            " Select any section listed and edit to assign Tasks, owner, priority, etc."
    On Error GoTo EmptyServer
    Set ServerFolderContents = GetAPIFolder(APIFoName, LCase(DocType) & "_breakdown", Array("section_number"))
    ReDim ServerFiles(1 To ServerFolderContents.Count)
    For i = 1 To ServerFolderContents.Count
        ServerFiles(i) = Replace(ServerFolderContents(i)("section_number") & "-" & ServerFolderContents(i)("title"), "_", " ")
    Next
'    ServerFiles = GetAPIFolder(APIFoName, "sow_analysis", Array("section_number"))
    If UBound(ServerFiles) < 1 Then GoTo EmptyServer
    On Error GoTo 0
    If Not IsGoodResponse(CStr(ServerFiles(1))) Then
        WriteLog 2, CurrentMod, "GetServerFolderContents", CStr(ServerFiles(1))
        Set Resp = CreateAPIFolder(DocType & " Analysis", , Description)
        If IsGoodResponse(Resp) Then
            APIFoName = Replace(Resp("@id"), ProjectURLStr, "") & "/"
        Else
            WriteLog 3, CurrentMod, "GetServerFolderContents"
        End If
    End If
EmptyServer:
    On Error GoTo 0
    Set DupColl = New Collection
    Set NewColl = New Collection
    FName = Dir(HTMLPath)
    Do While Len(FName)
        If Not FName Like "*.htmldeliverables.html" Then
            fs = Replace(GetFileName(FName, False), "_", " ")
            If IsEmpty(Match(fs, ServerFiles)) Then
                NewColl.Add fs
            Else
                DupColl.Add fs
            End If
        End If
        FName = Dir
    Loop
End Sub
Private Function DecideDuplicates() As Long
    If DupColl.Count = 0 Then
        DecideDuplicates = IIf(NewColl.Count = 0, -2, 1)
    Else
        If NewColl.Count = 0 Then
            WriteLog 2, CurrentMod, "DecideDuplicates", "All Files are Duplicated"
            frmServerFilesList.btnNewFiles.Visible = False
        Else
            WriteLog 2, CurrentMod, "DecideDuplicates", DupColl.Count & " Files are Duplicated"
            frmServerFilesList.btnNewFiles.Visible = True
        End If
        Select Case frmServerFilesList.ShowList("The server already has files with the same names as " & DupColl.Count & _
                " files to be uploaded" & vbNewLine & "Do you want to upload and create duplicates?", DupColl, NewColl)
        Case "Upload Duplicates": DecideDuplicates = 1
        Case "Upload new files Only": DecideDuplicates = 0
        Case "ReUpload": DecideDuplicates = 2
        Case "Cancel": DecideDuplicates = -1
        End Select
    End If
End Function
Sub StartUploading(DocType As String, OPath As String, HTMLPath As String)
    WriteLog 1, CurrentMod, "StartUploading"
    If Right$(OPath, 1) <> "\" Then OPath = OPath & "\"
    If Right$(HTMLPath, 1) <> "\" Then HTMLPath = HTMLPath & "\"
    DoneFolder = OPath & "Done\"
    CreateDir DoneFolder
    FailedFolder = OPath & "Failed\"
    CreateDir FailedFolder
    Select Case DocType
    Case "RFP"
        APIFoName = "/" & DefaultRFPFolder & "/"
        If Len(RFPURL) = 0 Then UploadRFP SDoc
    Case "Scope"
        APIFoName = "/" & DefaultScopeFolder & "/"
'        If Len(ScopeURL) = 0 Then uploadscope
    Case "Planning Document"
        APIFoName = "/" & DefaultPlanningFolder & "/"
'        If Len(PlanningURL) = 0 Then uploadPlanning
    Case "PMP"
        APIFoName = "/" & DefaultPMPFolder & "/"
'        If Len(PMPURL) = 0 Then uploadpmp
    End Select
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim UploadDuplicates As Long
    Dim FName As String
    On Error GoTo ex
    Dim Resp As WebResponse, oURL As String, SFName As String
    Dim FilesCount As Long, FailedCount As Long, DoneCount As Long, i As Long
Rtry:
    ResetFailed
    GetFilesLists DocType, HTMLPath
    UploadDuplicates = DecideDuplicates
    If UploadDuplicates = -1 Then GoTo can
    If UploadDuplicates = -2 Then MsgBox "Nothing to upload", vbCritical, "": GoTo ex
    If UploadDuplicates = 2 Then
        For i = 1 To ServerFolderContents.Count
            DeleteAPIContent CStr(ServerFolderContents(i)("@id"))
            UploadDuplicates = 1
        Next
    End If
    On Error Resume Next
    ProgressBar.Reset
    ProgressBar.HideApplication = True
    ProgressBar.BarsColor CLng(ProjectColorStr)
    ProgressBar.Dom(1) = IIf(UploadDuplicates = 1, DupColl.Count + NewColl.Count, NewColl.Count)
    ProgressBar.Progress , "Uploading File No. 1", 0
    ProgressBar.Caption = "Exporting Progress"
    ProgressBar.Show
    FilesCount = 0
    DoneCount = 0
    FailedCount = 0
    FName = Dir(HTMLPath)
    Do While Len(FName)
        If Not FName Like "*.htmldeliverables.html" Then
            If Not UploadDuplicates = 1 Then
                SFName = GetFileName(FName, False)
                If Not IsEmpty(Match(SFName, ServerFiles)) Then
                    FSO.MoveFile HTMLPath & FName, DoneFolder & FName
                    FSO.MoveFile HTMLPath & FName & "deliverables.html", DoneFolder & FName & "deliverables.html"
                    GoTo NxtFile
                End If
            End If
            FilesCount = FilesCount + 1
            If ProgressBar.Progress(, "Uploading File No. " & FilesCount) Or Set_Cancelled Then GoTo can
            Set Resp = CreateAPIBreakdown(DocType, HTMLPath & FName, APIFoName)
            If IsGoodResponse(Resp) Then
                FSO.MoveFile HTMLPath & FName, DoneFolder & FName
                FSO.MoveFile HTMLPath & FName & "deliverables.html", DoneFolder & FName & "deliverables.html"
                DoneCount = DoneCount + 1
                WriteLog 1, CurrentMod, "StartUploading", "Uploaded file: " & FName
            Else
                FSO.MoveFile HTMLPath & FName, FailedFolder & FName
                FSO.MoveFile HTMLPath & FName & "deliverables.html", FailedFolder & FName & "deliverables.html"
                FailedCount = FailedCount + 1
                WriteLog 3, CurrentMod, "StartUploading" ', Resp
            End If
        End If
NxtFile:
        FName = Dir(HTMLPath) 'Because the last file was moved
    Loop
    
    oURL = ProjectURLStr & APIFoName 'Mid$(APIFoName, 1, Len(APIFoName) - 1)
    
    FName = "Done uploading." & vbNewLine & vbNewLine & _
            "Total number of files: " & FilesCount & vbNewLine & _
            "  Scope items Created successfully: " & DoneCount
    SFName = "_________________________________" & vbNewLine & _
            "Uploaded files can be found in:"
    
    If FailedCount Then
        WriteLog 3, CurrentMod, "StartUploading", FailedCount & "/" & FilesCount & " Files were not uploaded"
        SavePUploaded "No"
        If frmMsgBox.Display(Array(FName, _
                "  Files not uploaded: " & FailedCount, SFName, oURL), _
                Array("Retry", "Cancel"), Exclamation, Links:=Array("", "", "", oURL)) = "Retry" Then
            WriteLog 2, CurrentMod, "StartUploading", "Rtrying.."
            GoTo Rtry
        End If
    Else
        SavePUploaded "Yes"
        If frmMsgBox.Display(Array(FName, "", SFName, oURL), Array("OK", "Close"), Success, Links:=Array("", "", "", oURL)) = "OK" Then
            On Error Resume Next
            GoToLink oURL
        End If
        WriteLog 1, CurrentMod, "StartUploading", "All Files Uploaded"
    End If
    Unload ProgressBar
    If FailedCount = 0 Then DeleteFolder OPath
    Exit Sub
can:
    MsgBox "Cancelled by user", vbExclamation, ""
    Exit Sub
ex:
    WriteLog 3, CurrentMod, "StartUploading", Err.Number & ":" & Err.Description
    MsgBox "Try again or use help button to send feedback.", vbCritical, "Unknown Error"
'    Stop
'    Resume
End Sub
'Sub test()
'    frmUploadDone.Label1.Caption = "Done uploading." & vbNewLine & _
'            "Total number of files: " & "FilesCount" & vbNewLine & _
'            "SOW Created successfully: " & "DoneCount" & vbNewLine & _
'            "Files not uploaded: " & "FailedCount" & vbNewLine & _
'            "_____________________________________" & vbNewLine & vbNewLine & _
'            "Uploaded files can be found in:"
''    frmUploadDone.Label3.Caption =
'    frmUploadDone.Label2.Caption = ProjectURLStr & "Mid$(APIFoName, 2, Len(APIFoName) - 2)"
'    frmUploadDone.Show
'End Sub
