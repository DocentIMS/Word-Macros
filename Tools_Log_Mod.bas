Attribute VB_Name = "Tools_Log_Mod"
Option Explicit
Private Const LogFolder = "Word Log\"
Private Const SelectedLevel = 1 '1: INFO & others, 2: Warning & Errors, 3: Errors only, 0: None,
Private Const cLogTo = 1
Private LogTo As Long '0: Immediate Window, 1: LogFile
Private Const KeepCount = 30
Private Const SendCount = 5
Private No As Long
Private LogFullName As String
'%AppData%\Docent IMS LLC\Docent Local\Log Files\Word Log
Public Function CreateLogFile() As Integer
    Dim LogFoPath As String, LogFName As String
    LogFoPath = GetInstallationPath & LogFolder
    CreateDir LogFoPath
    KeepOnly LogFoPath
    LogFName = Format(Now, "yy.mm.dd_hh.mm.ss") & "_Log.csv"
    LogFullName = LogFoPath & LogFName
    No = FreeFile
    On Error Resume Next
    Open LogFullName For Output As #No
    Close #No
    Open LogFoPath & LogFName For Append As #No
    Print #No, "Time,Project Name,Module Name,Sub Name,Log Level,Description"
End Function
Private Sub KeepOnly(LogFoPath As String)
    Dim LogFName As String, i As Long
    LogFName = Dir(LogFoPath & "*.csv")
    On Error Resume Next
    Do While Len(LogFName)
        i = i + 1
        If i > KeepCount Then Kill LogFoPath & LogFName
        LogFName = Dir
    Loop
End Sub
Private Function GenLogLine(ByVal LogLevel As Long, ByVal CurrentMod As String, _
            ByVal CurrentSub As String, Optional ByVal LogText As String) As String
    If InStr(LogText, """") Then LogText = Replace(LogText, """", "'")
    If InStr(LogText, ",") + _
        InStr(LogText, Chr(13)) + _
        InStr(LogText, Chr(10)) > 0 Then LogText = """" & LogText & """"
    Select Case LogLevel
    Case 1: LogText = "INFO," & LogText
    Case 2: LogText = "WARNING," & LogText
    Case 3: LogText = "ERROR," & LogText
    End Select
    LogText = Timer & "," & ProjectNameStr & "," & CurrentMod & "," & CurrentSub & "," & LogText 'Now
    GenLogLine = LogText
End Function
Public Sub WriteLog(ByVal LogLevel As Long, ByVal CurrentMod As String, _
            ByVal CurrentSub As String, Optional ByVal LogText As String)
    Dim LogLine As String
    If LogLevel < SelectedLevel Or SelectedLevel = 0 Then Exit Sub
    If Len(CurrentMod) = 0 Then Exit Sub
    LogLine = GenLogLine(LogLevel, CurrentMod, CurrentSub, LogText)
    LogTo = IIf(Application.UserName = "Abdallah Ali", 0, cLogTo)
    If LogTo = 0 Then
        Debug.Print LogLine
    Else
        On Error Resume Next
RePrint:
        Print #No, GenLogLine(LogLevel, CurrentMod, CurrentSub, LogText)
        If Err.Number Then
            CreateLogFile
            Err.Clear
            GoTo RePrint
        End If
    End If
End Sub
Public Sub CloseLog()
    Close #No
End Sub
Public Function WriteLogLine(ByVal LogLevel As Long, ByVal CurrentMod As String, ByVal CurrentSub As String, Optional ByVal LogText As String)
    If LogLevel < SelectedLevel Then Exit Function
    If Len(CurrentMod) = 0 Then Exit Function
    On Error Resume Next
'    If LogLevel >= SelectedLevel Then
    Select Case LogLevel
    Case 1: LogText = "INFO: " & LogText
    Case 2: LogText = "WARNING: " & LogText
    Case 3: LogText = "ERROR: " & LogText
    End Select
    LogText = "[" & Now & "] {" & CurrentMod & "-" & CurrentSub & "} " & LogText
RePrint:
    Print #No, LogText
    Close #No
    Open LogFullName For Append As #No
    If Err.Number Then
        CreateLogFile
        Err.Clear
        GoTo RePrint
    End If
'    End If
End Function
Sub ZipLog()
    CloseLog
    CreateZipFile "Log"
    Dim LogFoPath As String, LogFName As String, i As Long
    LogFoPath = GetInstallationPath & LogFolder
'    LogFoPath = GetLogPath
    If Len(LogFoPath) = 0 Then LogFoPath = ThisDocument.Path
    If Right$(LogFoPath, 1) <> "\" Then LogFoPath = LogFoPath & "\"
    LogFName = Dir(LogFoPath & "*.csv")
    Do While Len(LogFName)
        ZipFiles LogZip, LogFoPath & LogFName
        LogFName = Dir
        i = i + 1
        If i > SendCount Then Exit Do
    Loop
End Sub
