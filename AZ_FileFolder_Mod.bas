Attribute VB_Name = "AZ_FileFolder_Mod"
Option Explicit
Option Compare Text
Dim FSO As Object
#If Mac Then
Const PthSep As String = "/"
#Else
Const PthSep As String = "\"
#End If
Public Enum FileType
    FTInvalid = 0 'DEFAULT
    ExcelFiles = 2 ^ 0
    CSVFiles = 2 ^ 1
    ExcelAndCSVFiles = 2 ^ 2
    WordFiles = 2 ^ 3
    PDFFiles = 2 ^ 4
    PPTFiles = 2 ^ 5
    AllFiles = 2 ^ 6
    TxtFiles = 2 ^ 7
End Enum
Public Enum CompareType
    ecOR = 0 'default'
    ecAnd = 1
End Enum
Public Function EnumCompare( _
    theEnum As Variant, _
    enumMember As Variant, _
    Optional ByVal cType As CompareType = CompareType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(cType = CompareType.ecOR, c <> 0, c = enumMember)
End Function
Function FolderExists(ByVal FolderPath As String) As Boolean
    If Right$(FolderPath, 1) <> Application.PathSeparator Then FolderPath = FolderPath & Application.PathSeparator
    On Error Resume Next
    FolderExists = Dir(FolderPath) <> vbNullString
End Function
Function GetFile( _
    Optional Title As String, _
    Optional InitialFileName As String, _
    Optional MultiSelect As Boolean = False, _
    Optional FileType As FileType = ExcelFiles) As Object
    Set GetFile = New Collection
    #If Mac Then
    Set GetFile = GetFileMAC(Title, InitialFileName, MultiSelect, FileType)
    #Else
    Set GetFile = GetFileWIN(Title, InitialFileName, MultiSelect, FileType)
    #End If
End Function
Private Function GetFileWIN( _
        Optional Title As String, _
        Optional InitialFileName As String, _
        Optional MultiSelect As Boolean = False, _
        Optional FileType As FileType = ExcelFiles) As Object
        
    Dim SelectedFormats As String, SelectedCount As Long
    If InitialFileName = vbNullString Then InitialFileName = ThisDocument.Path
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Title = Title
        .InitialFileName = InitialFileName
        .AllowMultiSelect = MultiSelect
        If EnumCompare(FileType, PPTFiles + AllFiles) Then
            .Filters.Add "PowerPoint Files", "*.ppt; *.pptx; *.pptm", 1
            SelectedFormats = SelectedFormats & "; *.ppt; *.pptx; *.pptm"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, PDFFiles + AllFiles) Then
            .Filters.Add "PDF Files", "*.pdf", 1
            SelectedFormats = SelectedFormats & "; *.pdf"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, WordFiles + AllFiles) Then
            .Filters.Add "Word Files", "*.docx; *.docm; *.doc", 1
            SelectedFormats = SelectedFormats & "; *.docx; *.docm; *.doc"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, CSVFiles + AllFiles) Then
            .Filters.Add "Comma separated Files", "*.csv", 1
            SelectedFormats = SelectedFormats & "; *.csv"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, TxtFiles + AllFiles) Then
            .Filters.Add "Text Files", "*.txt", 1
            SelectedFormats = SelectedFormats & "; *.txt"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, ExcelFiles + AllFiles) Then
            .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
            SelectedFormats = SelectedFormats & "; *.xlsx; *.xlsm; *.xls; *.xlsb"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, ExcelAndCSVFiles + AllFiles) Then
            .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb; *.csv", 1
            SelectedFormats = SelectedFormats & "; *.xlsx; *.xlsm; *.xls; *.xlsb; *.csv"
            SelectedCount = SelectedCount + 1
        End If
        If EnumCompare(FileType, AllFiles) Then
            .Filters.Add "All Files", "*.*", 1
            SelectedCount = -1
        End If
        If SelectedCount > 0 Then
            SelectedFormats = Right$(SelectedFormats, Len(SelectedFormats) - 2)
            .Filters.Add "All Supported Files", SelectedFormats, 1
        End If
        If .Show <> -1 Then Exit Function
        Set GetFileWIN = .SelectedItems
    End With
End Function
Private Function GetFileMAC( _
    Optional Title As String, _
    Optional InitialFileName As String, _
    Optional MultiSelect As Boolean = False, _
    Optional FileType As FileType = ExcelFiles) As Object
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016 and higher
    'Ron de Bruin, 20 March 2016
    Dim MyScript As String
    Dim MyFiles As String
    Dim FileFormat As String

    'Dim FName As String ', Get_Path As String
    Dim MySplit() As String
    Dim n As Long

    'Dim mybook As Workbook


    'In this example you can only select xlsx files
    'See my webpage how to use other and more formats.
    FileFormat = MACFileFormat(FileType) ' "{""org.openxmlformats.spreadsheetml.sheet"",""com.adobe.pdf""}"

    On Error Resume Next
    If InitialFileName = vbNullString Then
        InitialFileName = MacScript("return (path to desktop folder) as String")
    Else
        InitialFileName = MACPathToAlias(InitialFileName)
    End If
    'Or use A full path with as separator the :
    'MyPath = "HarddriveName:Users::Desktop:YourFolder:"

    'Building the applescript string, do not change this
    If Val(Application.Version) < 15 Then
        'This is Mac Excel 2011
        If MultiSelect Then
            MyScript = _
                "set applescript's text item delimiters to {ASCII character 10} " & vbNewLine & _
                "try " & vbNewLine & _
                "set theFiles to (choose file " & _
                "of type " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                InitialFileName & """ with multiple selections allowed) as string" & vbNewLine & _
                "set applescript's text item delimiters to """" " & vbNewLine & _
                "return theFiles"
        Else
            MyScript = _
                "set theFile to (choose file " & _
                "of type " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                InitialFileName & """ without multiple selections allowed) as string" & vbNewLine & _
                "return theFile"
        End If
    Else
        'This is Mac Excel 2016 or higher
        If MultiSelect Then
            MyScript = _
                "set theFiles to (choose file " & _
                "of type " & FileFormat & " " & _
                "with prompt ""Please select a file or files"" default location alias """ & _
                InitialFileName & """ with multiple selections allowed)" & vbNewLine & _
                "set thePOSIXFiles to {}" & vbNewLine & _
                "repeat with aFile in theFiles" & vbNewLine & _
                "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
                "end repeat" & vbNewLine & _
                "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
                "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
                "set text item delimiters to TID" & vbNewLine & _
                "return thePOSIXFiles"
        Else
            MyScript = _
                "set theFile to (choose file " & _
                "of type " & FileFormat & " " & _
                "with prompt ""Please select a file"" default location alias """ & _
                InitialFileName & """ without multiple selections allowed) as string" & vbNewLine & _
                "return posix path of theFile"
        End If
    End If
    If FileFormat = vbNullString Then MyScript = Replace(MyScript, "of type " & FileFormat & " ", vbNullString)
    MyFiles = MacScript(MyScript)
    Set GetFileMAC = New Collection
    MySplit = Split(MyFiles, Chr$(10))
    For n = LBound(MySplit) To UBound(MySplit)
        GetFileMAC.Add MySplit(n)
    Next
    On Error GoTo 0
End Function

Function SelectFolder(Optional Title As String, _
    Optional InitialFolderName As String) As String
    Dim FName As String
    On Error Resume Next
    #If Mac Then
    InitialFolderName = MacScript("return (path to desktop folder) as String")
    'Or use RootFolder = "Macintosh HD:Users:YourUserName:Desktop:TestMap:"
    'Note : for a fixed path use : as seperator in 2011 and 2016
    If Val(Application.DocVer) < 15 Then
        FName = "(choose folder with prompt """ & Title & """" & _
            " default location alias """ & InitialFolderName & """) as string"
    Else
        FName = "return posix path of (choose folder with prompt """ & Title & """" & _
            " default location alias """ & InitialFolderName & """) as string"
    End If
    SelectFolder = MacScript(FName)
    #Else
    With Application.FileDialog(4)
        .Title = Title
        .InitialFileName = InitialFolderName
        If .Show <> -1 Then Exit Function
        FName = .SelectedItems(1) & PthSep
        SelectFolder = FName
    End With
    #End If
End Function
Sub CreateShortcut(OfWhat As String, ToWhere As String)
    Dim FName As String
    On Error Resume Next
    FName = GetFileName(OfWhat)
    If Right$(ToWhere, 1) <> "\" Then ToWhere = ToWhere & "\"
    With CreateObject("WScript.Shell").CreateShortcut(ToWhere & FName & ".lnk")
        .TargetPath = OfWhat
        .Description = "Shortcut to " & FName
        .Save
    End With
End Sub
Function CreateDir(ByVal DirStr As String, Optional n As Long) As String
    Dim ParentDir As String, FSO As Object
    If Right$(DirStr, 1) = Application.PathSeparator Then DirStr = Left$(DirStr, Len(DirStr) - 1)
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(DirStr) Then
        ParentDir = Left$(DirStr, InStrRev(DirStr, Application.PathSeparator, Len(DirStr)))
        If FSO.FolderExists(ParentDir) Then
            FSO.CreateFolder DirStr
            CreateDir = DirStr
        Else
            CreateDir ParentDir
            FSO.CreateFolder DirStr
            CreateDir = DirStr
        End If
    ElseIf n > 0 Then
        Dim FoName As String, i As Long
        ParentDir = Left$(DirStr, InStrRev(DirStr, Application.PathSeparator, Len(DirStr)))
        FoName = Right$(DirStr, Len(DirStr) - Len(ParentDir))
        Do
            If FoName Like "*)" Then i = InStrRev(FoName, "(") + 1 Else Exit Do
            If IsNumeric(Mid$(FoName, i, Len(FoName) - i)) Then
                n = Mid$(FoName, i, Len(FoName) - i)
                FoName = Trim$(Left$(FoName, i - 2))
            Else
                Exit Do
            End If
        Loop
        If n = 1 And Not FSO.FolderExists(ParentDir & FoName) Then
            FSO.CreateFolder ParentDir & FoName
            CreateDir = ParentDir & FoName
        Else
            Do While FSO.FolderExists(ParentDir & FoName & " (" & n + 1 & ")"): n = n + 1: Loop
            FSO.CreateFolder ParentDir & FoName & " (" & n + 1 & ")"
            CreateDir = ParentDir & FoName & " (" & n + 1 & ")"
        End If
    Else
        CreateDir = DirStr
    End If
End Function
Function GetParentDir(ByVal Directory As String, Optional Sep As String) As String
    On Error Resume Next
    If Sep = "" Then
        Directory = GetParentDir(Directory, "\")
        Directory = GetParentDir(Directory, "/")
        'Sep = Application.PathSeparator
    End If
    #If Mac Then
        Directory = MACAliasToPath(Directory)
    #End If
    If Right$(Directory, Len(Sep)) = Sep Then Directory = Left$(Directory, Len(Directory) - Len(Sep))
    If InStrRev(Directory, Sep) Then Directory = Left$(Directory, InStrRev(Directory, Sep))
    GetParentDir = Directory
End Function
Function GetFileName(ByVal Pth As String, Optional ByVal KeepExtension As Boolean = True, Optional Sep As String) As String
    If InStr(Pth, "\") + InStr(Pth, "/") > 0 Then
        If Sep = "" Then
            Pth = TrimPath(Pth, "\")
            Pth = TrimPath(Pth, "/")
            'Sep = Application.PathSeparator
        Else
            Pth = TrimPath(Pth, Sep)
        End If
    End If
    GetFileName = IIf(KeepExtension, Pth, Replace(Pth, GetFileExtension(Pth), ""))
End Function
Function GetFileExtension(ByVal FName As String) As String
    If InStr(FName, "\") + InStr(FName, "/") > 0 Then FName = GetFileName(FName)
    If InStr(FName, ".") Then FName = Right$(FName, Len(FName) - InStrRev(FName, ".") + 1) Else FName = ""
    GetFileExtension = FName
End Function
Function GetFileFormat(ByVal FName As String)
    Select Case GetFileExtension(FName)
    Case ".Doc": GetFileFormat = 0 'Word 97-2003 Document
    Case ".docm": GetFileFormat = 13 'Word Macro-Enabled Document
    Case ".docx": GetFileFormat = 12 'Word Document & Strict Open XML Document'24?
    Case ".dot": GetFileFormat = 1 'Word 97-2003 Template
    Case ".dotm": GetFileFormat = 15 'Word Macro-Enabled Template
    Case ".dotx": GetFileFormat = 14 'Word Template
    Case ".htm", ".HTML": GetFileFormat = 0  'web Page & web Page, Filtered
    Case ".mht", ".mhtml": GetFileFormat = 0 'Single File Web Page
    Case ".odt": GetFileFormat = 0 'OpenDocument Text
    Case ".PDF": GetFileFormat = 17 'PDF
    Case ".RTF": GetFileFormat = 0 'Rich Text Format
    Case ".Txt": GetFileFormat = 0 'Plain Text
    Case ".wps": GetFileFormat = 0 'Works 6-9 Document
    Case ".XML": GetFileFormat = 11 'Word 2003 XML Document & Word XML Document
    Case ".XPS": GetFileFormat = 18 'XPS Document
    End Select
End Function
Private Function TrimPath(ByVal Pth As String, Sep As String) As String
    If Right$(Pth, Len(Sep)) = Sep Then Pth = Left$(Pth, Len(Pth) - Len(Sep))
    If Len(Sep) Then If InStrRev(Pth, Sep) Then Pth = Right$(Pth, Len(Pth) - InStrRev(Pth, Sep))
    TrimPath = Pth
End Function
Private Sub DeleteEmptyFolder(FoName As String)
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    If Right$(FoName, 1) = PthSep Then FoName = Left$(FoName, Len(FoName) - 1)
    If Not FilesInFolder(FoName) Then FSO.DeleteFolder FoName
End Sub
Sub DeleteFile(FName As String)
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FileExists(FName) Then
        CloseIfOpen FName
        FSO.DeleteFile FName
    End If
End Sub
Private Sub CloseIfOpen(FName As String)
'    Dim Fo As Object
    Dim wbName As String
    On Error Resume Next
    wbName = Right$(FName, Len(FName) - InStrRev(FName, Application.PathSeparator))
    Open FName For Input Lock Read As #1: Close #1
    If Err.Number = 70 Then
        Err.Clear
'        Set Fo = GetObject(wbName)
'        If Err.Number <> 0 Then Set Fo = Workbooks(wbName)
'        Fo.Close True
'        Err.Clear
    End If
End Sub
Sub DeleteFolder(Optional ByVal FoName As String)
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    If Right$(FoName, 1) = PthSep Then FoName = Left$(FoName, Len(FoName) - 1)
    On Error Resume Next
    FSO.DeleteFolder FoName
'    Dim i As Long
'    Dim File As Object
'    Dim Folder As Object
'    Dim SubFolder As Object
'    If FoName = vbNullString Then FoName = CurDir$
'    If Right$(FoName, 1) <> PthSep Then FoName = FoName & PthSep
'    Set Folder = FSO.GetFolder(FoName)
'    For Each File In FSO.GetFolder(Folder).Files
'        DeleteFile File.Path
'    Next
'    For Each SubFolder In FSO.GetFolder(Folder).subFolders
'        DeleteFolder SubFolder.Path
'    Next
'    DeleteEmptyFolder FoName
End Sub
Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = Dir(filePath) <> vbNullString
End Function
Private Function FilesInFolder(Optional ByVal FoName As String) As Boolean
    Dim Folder As Object
    Dim SubFolder As Object

    If FoName = vbNullString Then FoName = CurDir$
    If Right$(FoName, 1) <> PthSep Then FoName = FoName & PthSep
    If FSO Is Nothing Then Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Folder = FSO.GetFolder(FoName)
    FilesInFolder = FSO.GetFolder(Folder).Files.Count
    If FilesInFolder Then Exit Function
    For Each SubFolder In FSO.GetFolder(Folder).SubFolders
        FilesInFolder = FilesInFolder(SubFolder.Path)
        If FilesInFolder Then Exit Function
    Next
End Function

Private Function MACFileFormat(FileType As FileType) As String
    'xls:  com.microsoft.Excel.xls
    'xlsx:  org.openxmlformats.spreadsheetml.Sheet
    'xlsm:  org.openxmlformats.spreadsheetml.Sheet.macroenabled
    'xlsb:  com.microsoft.Excel.Sheet.binary.macroenabled
    'csv : public.comma-separated-values-text
    'doc:  com.microsoft.word.doc
    'docx:  org.openxmlformats.wordprocessingml.Document
    'docm:  org.openxmlformats.wordprocessingml.Document.macroenabled
    'ppt:  com.microsoft.powerpoint.ppt
    'pptx:  org.openxmlformats.presentationml.presentation
    'pptm:  org.openxmlformats.presentationml.presentation.macroenabled
    'txt : public.plain-text
    'pdf:  com.adobe.pdf
    'jpg : public.jpeg
    'png : public.png
    'QIF:  com.apple.traditional -Mac - plain - Text
    'htm : public.html
    '"{""org.openxmlformats.spreadsheetml.sheet"",""com.microsoft.Excel.xls""}"
    If EnumCompare(FileType, AllFiles) Then
    ElseIf EnumCompare(FileType, ExcelFiles + AllFiles) Then
        MACFileFormat = _
            """com.microsoft.Excel.xls""," & _
            """org.openxmlformats.spreadsheetml.sheet""," & _
            """org.openxmlformats.spreadsheetml.Sheet.macroenabled""," & _
            """com.microsoft.Excel.Sheet.binary.macroenabled"""
    ElseIf EnumCompare(FileType, CSVFiles + AllFiles) Then
        MACFileFormat = _
            """public.comma-separated-values-text"""
    ElseIf EnumCompare(FileType, WordFiles + AllFiles) Then
        MACFileFormat = _
            """com.microsoft.word.doc""," & _
            """org.openxmlformats.wordprocessingml.Document""," & _
            """org.openxmlformats.wordprocessingml.Document.macroenabled"""
    ElseIf EnumCompare(FileType, PDFFiles + AllFiles) Then
        MACFileFormat = _
            """com.adobe.pdf"""
    ElseIf EnumCompare(FileType, PPTFiles + AllFiles) Then
        MACFileFormat = _
            """org.openxmlformats.presentationml.presentation""," & _
            """org.openxmlformats.presentationml.presentation.macroenabled"""
    End If
    If MACFileFormat <> vbNullString Then MACFileFormat = "{" & MACFileFormat & "}"
End Function
Function MACPathToAlias(aliasPath As String) As String
    Dim sMacScript As String
    sMacScript = "set theFilePath to POSIX file """ & aliasPath & """" & vbNewLine & _
        "set theFilePath to theFilePath as alias"
    On Error Resume Next
    MACPathToAlias = MacScript(sMacScript)
    If Err.Number Then
        MACPathToAlias = aliasPath
    Else
        MACPathToAlias = Replace(MACPathToAlias, Chr$(10), vbNullString)
    End If
End Function
Function GetFileContents(FName As String)
    Dim FF As Long, HTML As String
    On Error Resume Next
    FF = FreeFile
    Open FName For Input As #FF
    GetFileContents = Input$(LOF(FF), FF)
    Close #FF
End Function
Function IsWritable(ByVal Pth As String) As Boolean
    Dim No As Long
    On Error Resume Next
    CreateDir Pth
    Pth = GetAValidFileName(Pth & "\Temp.txt") '(Pth & "\tempfile.tmp"
    No = FreeFile
    Err.Clear
    Open Pth For Output As #No
    Close #No
    Kill Pth
    IsWritable = Err.Number = 0
End Function
Function GetAValidFileName(ByVal FName As String, Optional ByVal NewFName As String, _
    Optional NewExtension As String, Optional Suffix As String, Optional Prefex As String) As String
    Dim Extension As String
    Dim FNo As Long
    Dim n As Long
    FName = Replace(FName, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    n = InStrRev(FName, ".")
    If n Then Extension = Right$(FName, Len(FName) - InStrRev(FName, ".") + 1)
    If NewFName = vbNullString Then
        NewFName = Right$(FName, Len(FName) - InStrRev(FName, PthSep))
        NewFName = Prefex & Left$(NewFName, Len(NewFName) - Len(Extension)) & Suffix
        FName = Left$(FName, InStrRev(FName, PthSep)) & NewFName
    Else
        FName = Left$(FName, InStrRev(FName, PthSep)) & Prefex & NewFName & Suffix
    End If
    If NewExtension = vbNullString Then
        NewExtension = Extension
    ElseIf Left$(NewExtension, 1) <> "." Then
        NewExtension = "." & NewExtension
    End If
    GetAValidFileName = FName
    FNo = FreeFile()
    Do
        If FName Like "*)" Then n = InStrRev(FName, "(") + 1 Else Exit Do
        If IsNumeric(Mid$(FName, n, Len(FName) - n)) Then FName = Trim$(Left$(FName, n - 2)) Else Exit Do
    Loop
    n = 2
    On Error Resume Next
    CreateDir GetParentDir(GetAValidFileName)
    If Err.Number = 75 Then
        GetAValidFileName = ""
        Exit Function
    End If
    Err.Clear
    Do
        Open GetAValidFileName & NewExtension For Input Lock Read As #FNo: Close #FNo
        Select Case Err.Number
            Case 0, 70
                GetAValidFileName = FName & " (" & n & ")"
                n = n + 1
                Err.Clear
            Case 53
                Exit Do
            Case 76 ', 75
                CreateDir GetParentDir(GetAValidFileName)
                n = n - 1
            Case Else
                Stop
        End Select
    Loop
    GetAValidFileName = GetAValidFileName & NewExtension
End Function

