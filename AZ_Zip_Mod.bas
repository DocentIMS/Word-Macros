Attribute VB_Name = "AZ_Zip_Mod"
Option Explicit

#If Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Private No As Long, ShellApp As Object, FSO As Object
Public LogZip As String
Public AttachZip As String
'Sub Zip(Optional Folders As Variant, Optional Files As Variant)
'    Dim ShellApp As Object, i As Long, FStr As String, initCount As Long
'    CreateZipFile
'    On Error GoTo re
'    Set ShellApp = CreateObject("Shell.Application")
'
'        If Not IsMissing(Folders) Then
'        End If
'        If Not IsMissing(Files) Then
'        End If
'    End With
'    Exit Sub
're:
'    Resume
''    Close #No
'End Sub
Sub ZipFiles(ZipFName As String, Files As Variant)
    Dim i As Long
'    With CreateObject("Shell.Application")
        Select Case TypeName(Files)
        Case "String"
            AddFileToZip ZipFName, CStr(Files)
'            AddFileToZip .Namespace(CStr(ZipFName)), CStr(Files)
        Case Else
            For i = LBound(Files) To UBound(Files)
                AddFileToZip ZipFName, CStr(Files(i))
'                AddFileToZip .Namespace(CStr(ZipFName)), CStr(Files(i))
            Next
        End Select
'    End With
End Sub
Sub ZipFolders(ZipFName As String, Folders As Variant)
    Dim i As Long
'    With CreateObject("Shell.Application")
        Select Case TypeName(Folders)
        Case "String"
            AddFolderToZip ZipFName, CStr(Folders)
'            AddFolderToZip .Namespace(CStr(ZipFName)), CStr(Folders)
        Case Else
            For i = LBound(Folders) To UBound(Folders)
                AddFolderToZip ZipFName, CStr(Folders(i))
'                AddFolderToZip .Namespace(CStr(ZipFName)), CStr(Folders(i))
            Next
        End Select
'    End With
End Sub
Private Sub AddFileToZip(ZipFName As String, FName As String)
    On Error Resume Next
    Dim i As Long
    If ShellApp Is Nothing Then Set ShellApp = CreateObject("Shell.Application")
    With ShellApp
        i = .Namespace(CStr(ZipFName)).Items.Count
        .Namespace(CStr(ZipFName)).CopyHere CStr(FName)
        Do Until .Namespace(CStr(ZipFName)).Items.Count - i = 1
            Sleep 50: DoEvents
        Loop
    End With
End Sub
Private Sub AddFolderToZip(ZipFName As String, FoName As String)
    Dim i As Long, j As Long
    If ShellApp Is Nothing Then Set ShellApp = CreateObject("Shell.Application")
    With ShellApp
        j = .Namespace(FoName).Items.Count
        i = .Namespace(CStr(ZipFName)).Items.Count
        .Namespace(CStr(ZipFName)).CopyHere .Namespace(CStr(FoName)).Items
        Do Until .Namespace(CStr(ZipFName)).Items.Count - i = j
            Sleep 50: DoEvents
        Loop
    End With
End Sub
Sub CreateZipFile(FName As String)
    Select Case Left$(FName, 3)
    Case "Log"
        FName = Environ("temp") & "\DocentWordLog_" & Format(Now, "yy.mm.dd_hh.mm.ss") & ".zip"
        LogZip = FName
    Case "Att"
        FName = Environ("temp") & "\DocentUserAttachement_" & Format(Now, "yy.mm.dd_hh.mm.ss") & ".zip"
        AttachZip = FName
    End Select
'    ZipFullName = Environ("temp") & "\feedback.zip"
    No = FreeFile
    On Error Resume Next
    Open FName For Input As #No
    Close #No
    If Err.Number Then
        Open FName For Output As #No
        Print #No, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
        Close #No
    End If
End Sub
Sub DeleteZip()
    On Error Resume Next
    Kill LogZip
    Kill AttachZip
End Sub
'CreateZipFile("C:\Users\marks\Documents\ZipThisFolder\", "C:\Users\marks\Documents\NameOFZip.zip")
'            typedef enum _SHCONTF {
'              SHCONTF_CHECKING_FOR_CHILDREN = 0x10,
'              SHCONTF_FOLDERS = 0x20,
'              SHCONTF_NONFOLDERS = 0x40,
'              SHCONTF_INCLUDEHIDDEN = 0x80,
'              SHCONTF_INIT_ON_FIRST_NEXT = 0x100,
'              SHCONTF_NETPRINTERSRCH = 0x200,
'              SHCONTF_SHAREABLE = 0x400,
'              SHCONTF_STORAGE = 0x800,
'              SHCONTF_NAVIGATION_ENUM = 0x1000,
'              SHCONTF_FASTITEMS = 0x2000,
'              SHCONTF_FLATLIST = 0x4000,
'              SHCONTF_ENABLE_ASYNC = 0x8000,
'              SHCONTF_INCLUDESUPERHIDDEN = 0x10000
'            } ;
Sub UnzipAFile(ByVal ZipFName As Variant, FoName As Variant, _
            Optional SubFolder As String = "\customUI\images\", _
            Optional FilesLike As String = "", Optional ForceExtenstion As String = "")
    Dim ShellApp As Object, Files As Object, File As Object ' FolderItem  ',NS as she
    Const AddedTxt = "_Temp.zip"
    If Left$(SubFolder, 1) <> "\" Then SubFolder = "\" & SubFolder
    If Len(ForceExtenstion) Then If Left$(ForceExtenstion, 1) <> "." Then ForceExtenstion = "." & ForceExtenstion
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    If ShellApp Is Nothing Then Set ShellApp = CreateObject("Shell.Application")
    CreateDir FoName
    ZipFName = ZipFName & AddedTxt
    FSO.CopyFile Left$(ZipFName, Len(ZipFName) - Len(AddedTxt)), ZipFName
    With ShellApp
        Set Files = .Namespace(ZipFName & SubFolder).Items
        If Len(FilesLike) Then
            Files.Filter &H40, FilesLike
            For Each File In Files
                If Len(ForceExtenstion) Then
                    If Not FSO.FileExists(FoName & GetFileName(File, False) & ForceExtenstion) Then
                        If Not FSO.FileExists(FoName & GetFileName(File.Path)) Then .Namespace(CStr(FoName)).CopyHere File
                        FSO.MoveFile FoName & GetFileName(File.Path), FoName & GetFileName(File, False) & ForceExtenstion
                    End If
                Else
                    If Not FSO.FileExists(FoName & File.Name) Then .Namespace(CStr(FoName)).CopyHere File
                End If
            Next
        Else
            .Namespace(CStr(FoName)).CopyHere .Namespace(ZipFName & SubFolder).Items
        End If
    End With
    Kill ZipFName
End Sub
'UnzipAFile("C:\Users\marks\Documents\ZipHere.zip", "C:\Users\marks\Documents\UnzipHereFolder\")
'Sub Zip_All_Files_in_Folder()
'    Dim FileNameZip, FolderName
'    Dim strDate As String, DefPath As String
'    Dim oApp As Object
'    Dim Fold As Range
'    DefPath = Application.DefaultFilePath
'    If Right(DefPath, 1) <> "\" Then DefPath = DefPath & "\"
'    For Each Fold In Sheet1.Range("A1:A3")
'        FolderName = Fold.Value
'        strDate = Format(Now, " dd-mmm-yy h-mm-ss")
'        FileNameZip = DefPath & "MyFilesZip " & strDate & ".zip"
'        'Create empty Zip File
'        NewZip (FileNameZip)
'        Set oApp = CreateObject("Shell.Application")
'        'Copy the files to the compressed folder
'        oApp.Namespace(FileNameZip).CopyHere oApp.Namespace(FolderName).items
'        'Keep script waiting until Compressing is done
'        On Error Resume Next
'        Do Until oApp.Namespace(FileNameZip).items.Count = _
'           oApp.Namespace(FolderName).items.Count
'            Application.Wait (Now + TimeValue("0:00:01"))
'        Loop
'        On Error GoTo 0
'        MsgBox "You find the zipfile here: " & FileNameZip
'    Next Fold
'End Sub


