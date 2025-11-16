Attribute VB_Name = "Tools_GoToLink"
Option Explicit

'=======================================================
' Module: Tools_GoToLink
' Purpose: Handle opening URLs, files, and email links
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Provides wrapper functions for opening various link types:
'   - Web URLs (HTTP/HTTPS)
'   - Local file paths (C:\path\file.ext)
'   - Email addresses (mailto:)
'   - Downloaded API files
'
'   Uses Windows ShellExecute API to open links in their
'   default applications (browser, file manager, email client).
'
' Dependencies:
'   - AB_GlobalVars (for ProjectURLStr)
'   - AD_Upload_mod (for DownloadAPIFile function)
'   - AZ_Log_Mod (for WriteLog)
'   - Windows shell32.dll (ShellExecute API)
'
' Main Procedures:
'   - GoToLocation: Open local file path or URL directly
'   - GoToLink: Smart link handler with automatic formatting
'   - GoToEmail: Open email client with recipient address
'   - OpenFileIn: Download and open file from URL
'
' Usage Examples:
'   GoToLocation "C:\Documents\file.pdf"
'   GoToLink "subfolder/page"  ' Prepends ProjectURLStr
'   GoToLink "example.com"     ' Adds http:// prefix
'   GoToEmail "user@example.com"
'   OpenFileIn "http://server.com/file.pdf", "document.pdf"
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive documentation
'       * Added error handling throughout
'       * Improved parameter validation
'       * Added type safety
'       * Enhanced logging
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Tools_GoToLink"

' ShellExecute return code constants
Private Const SE_ERR_FNF As Long = 2                ' File not found
Private Const SE_ERR_PATH As Long = 3               ' Path not found
Private Const SE_ERR_ACCESSDENIED As Long = 5       ' Access denied
Private Const SE_ERR_OOM As Long = 8                ' Out of memory
Private Const SE_ERR_DLLNOTFOUND As Long = 32       ' DLL not found
Private Const SE_ERR_SHARE As Long = 26             ' Sharing violation
Private Const SE_ERR_ASSOCINCOMPLETE As Long = 27   ' Association incomplete
Private Const SE_ERR_DDETIMEOUT As Long = 28        ' DDE timeout
Private Const SE_ERR_DDEFAIL As Long = 29           ' DDE failed
Private Const SE_ERR_DDEBUSY As Long = 30           ' DDE busy
Private Const SE_ERR_NOASSOC As Long = 31           ' No association
Private Const SUCCESS_THRESHOLD As Long = 32        ' Values > 32 indicate success

'=======================================================
' WINDOWS API DECLARATIONS
'=======================================================

' ShellExecute API function for opening files and URLs
#If VBA7 Then
    ' 64-bit Office
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
    ' 32-bit Office
    Private Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal fileName As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMinimizedFocus _
        ) As Long
#End If

'=======================================================
' PUBLIC PROCEDURES
'=======================================================

'=======================================================
' Sub: GoToLocation
' Purpose: Open a local file path or URL
'
' Parameters:
'   linkPath - File path or URL to open (String)
'
' Description:
'   Opens the specified path using the system's default
'   application. Can handle:
'   - Local file paths (C:\folder\file.ext)
'   - UNC paths (\\server\share\file.ext)
'   - HTTP/HTTPS URLs
'   - Any other protocol Windows recognizes
'
' Example:
'   GoToLocation "C:\Documents\report.pdf"
'   GoToLocation "http://www.example.com"
'   GoToLocation "\\server\share\document.docx"
'
' Error Handling:
'   - Validates input
'   - Checks ShellExecute return code
'   - Provides user-friendly error messages
'   - Logs all operations
'=======================================================
Sub GoToLocation(ByVal linkPath As String)
    Const PROC_NAME As String = "GoToLocation"
    
    Dim result As Long
    Dim errorMessage As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(linkPath) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty link path provided"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Opening location: " & linkPath
    
    ' Execute shell command
    result = ShellExecute(0, "open", linkPath, vbNullString, vbNullString, vbNormalFocus)
    
    ' Check result
    If result <= SUCCESS_THRESHOLD Then
        ' Operation failed - get appropriate error message
        errorMessage = GetShellExecuteError(result)
        
        WriteLog 3, CurrentMod, PROC_NAME, _
                 "ShellExecute failed with code " & result & ": " & errorMessage
        
        MsgBox "Unable to open:" & vbNewLine & vbNewLine & _
               linkPath & vbNewLine & vbNewLine & _
               "Error: " & errorMessage & vbNewLine & vbNewLine & _
               "Please verify the path exists and you have permission to access it.", _
               vbExclamation, "Unable to Open"
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "Location opened successfully"
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred while opening the location:" & vbNewLine & vbNewLine & _
           linkPath & vbNewLine & vbNewLine & _
           "Error: " & Err.Description, _
           vbCritical, "Error Opening Location"
End Sub

'=======================================================
' Sub: GoToLink
' Purpose: Open a project link with automatic URL formatting
'
' Parameters:
'   linkPath - Relative or absolute link path (String)
'
' Description:
'   Smart link handler that automatically formats links:
'
'   Detection rules (in order):
'   1. File paths (e.g., "C:\file.txt") → Opens directly
'   2. Relative paths (no domain) → Prepends ProjectURLStr
'   3. Domains without protocol → Adds "http://"
'   4. Full URLs → Opens as-is
'
' Examples:
'   Input: "C:\docs\file.pdf"
'   Output: Opens local file directly
'
'   Input: "subfolder/page"
'   Output: Opens ProjectURLStr & "subfolder/page"
'
'   Input: "example.com"
'   Output: Opens "http://example.com"
'
'   Input: "http://site.com"
'   Output: Opens "http://site.com" as-is
'
' Error Handling:
'   - Validates input
'   - Logs all transformations
'   - Handles ShellExecute errors
'=======================================================
Sub GoToLink(ByVal linkPath As String)
    Const PROC_NAME As String = "GoToLink"
    
    Dim fullLink As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(linkPath) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty link path provided"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Processing link: " & linkPath
    
    ' Check if it's a local file path (e.g., C:\path\file.ext or D:\...)
    If linkPath Like "?:\*" Then
        WriteLog 1, CurrentMod, PROC_NAME, "Detected as local file path"
        GoToLocation linkPath
        Exit Sub
    End If
    
    fullLink = linkPath
    
    ' Add project URL prefix for relative paths (no domain detected)
    If InStr(fullLink, ".") = 0 Then
        fullLink = ProjectURLStr & fullLink
        WriteLog 1, CurrentMod, PROC_NAME, "Added project URL prefix: " & fullLink
    End If
    
    ' Add http:// prefix if protocol not specified
    If Not (fullLink Like "http://*" Or fullLink Like "https://*") Then
        fullLink = "http://" & fullLink
        WriteLog 1, CurrentMod, PROC_NAME, "Added http:// prefix: " & fullLink
    End If
    
    ' Open the formatted link
    GoToLocation fullLink
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred while processing the link:" & vbNewLine & vbNewLine & _
           linkPath & vbNewLine & vbNewLine & _
           "Error: " & Err.Description, _
           vbCritical, "Error Processing Link"
End Sub

'=======================================================
' Sub: GoToEmail
' Purpose: Open email client with recipient address
'
' Parameters:
'   emailAddress - Email address to open (String)
'
' Description:
'   Opens the default email client with a new message
'   addressed to the specified email address.
'
'   Validates email format before opening (basic check
'   for @ and . characters).
'
' Example:
'   GoToEmail "user@example.com"
'
' Validation:
'   Email must match pattern "*@*.*" (contains @ and dot)
'
' Error Handling:
'   - Validates email format
'   - Shows friendly error for invalid emails
'   - Logs all operations
'=======================================================
Sub GoToEmail(ByVal emailAddress As String)
    Const PROC_NAME As String = "GoToEmail"
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(emailAddress) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty email address provided"
        Exit Sub
    End If
    
    ' Validate email format (basic check)
    If Not (emailAddress Like "*@*.*") Then
        WriteLog 2, CurrentMod, PROC_NAME, "Invalid email format: " & emailAddress
        
        MsgBox "Invalid email address format:" & vbNewLine & vbNewLine & _
               emailAddress & vbNewLine & vbNewLine & _
               "Email addresses must contain @ and a domain.", _
               vbExclamation, "Invalid Email Address"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Opening email client for: " & emailAddress
    
    ' Open email client with mailto: protocol
    GoToLocation "mailto:" & emailAddress
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred while opening the email client:" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description, _
           vbCritical, "Error Opening Email"
End Sub

'=======================================================
' Sub: OpenFileIn
' Purpose: Download and open a file from URL
'
' Parameters:
'   fileURL - URL of file to download (String)
'   localFileName - Suggested local filename (String)
'
' Description:
'   Downloads a file from the specified URL using the
'   DownloadAPIFile function, then opens the downloaded
'   file in the default application.
'
' Example:
'   OpenFileIn "http://server.com/document.pdf", "report.pdf"
'
' Workflow:
'   1. Adds http:// prefix if needed
'   2. Downloads file using DownloadAPIFile
'   3. Opens downloaded file with GoToLocation
'
' Error Handling:
'   - Validates inputs
'   - Checks download success
'   - Provides user feedback on errors
'   - Logs all operations
'=======================================================
Sub OpenFileIn(ByVal fileURL As String, ByVal localFileName As String)
    Const PROC_NAME As String = "OpenFileIn"
    
    Dim downloadedPath As String
    Dim fullURL As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(fileURL) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty file URL provided"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Downloading and opening file from: " & fileURL
    
    fullURL = fileURL
    
    ' Add http:// prefix if needed
    If Not (fullURL Like "http://*" Or fullURL Like "https://*") Then
        fullURL = "http://" & fullURL
        WriteLog 1, CurrentMod, PROC_NAME, "Added http:// prefix"
    End If
    
    ' Download file
    downloadedPath = DownloadAPIFile(fullURL, False, localFileName)
    
    ' Check download success
    If Len(downloadedPath) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Failed to download file from: " & fileURL
        
        MsgBox "Failed to download file from:" & vbNewLine & vbNewLine & _
               fileURL & vbNewLine & vbNewLine & _
               "Please check your internet connection and try again.", _
               vbExclamation, "Download Failed"
        Exit Sub
    End If
    
    ' Open downloaded file
    WriteLog 1, CurrentMod, PROC_NAME, "Opening downloaded file: " & downloadedPath
    GoToLocation downloadedPath
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred while downloading or opening the file:" & vbNewLine & vbNewLine & _
           "URL: " & fileURL & vbNewLine & vbNewLine & _
           "Error: " & Err.Description, _
           vbCritical, "Error Opening File"
End Sub

'=======================================================
' PRIVATE HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: GetShellExecuteError
' Purpose: Get user-friendly error message for ShellExecute codes
'
' Parameters:
'   errorCode - ShellExecute return code (Long)
'
' Returns:
'   User-friendly error message (String)
'
' Description:
'   Translates ShellExecute error codes into readable
'   error messages for display to users.
'=======================================================
Private Function GetShellExecuteError(ByVal errorCode As Long) As String
    Select Case errorCode
        Case SE_ERR_FNF
            GetShellExecuteError = "File not found"
        Case SE_ERR_PATH
            GetShellExecuteError = "Path not found"
        Case SE_ERR_ACCESSDENIED
            GetShellExecuteError = "Access denied"
        Case SE_ERR_OOM
            GetShellExecuteError = "Out of memory"
        Case SE_ERR_DLLNOTFOUND
            GetShellExecuteError = "Required DLL not found"
        Case SE_ERR_SHARE
            GetShellExecuteError = "Sharing violation"
        Case SE_ERR_ASSOCINCOMPLETE
            GetShellExecuteError = "File association incomplete"
        Case SE_ERR_DDETIMEOUT
            GetShellExecuteError = "DDE timeout"
        Case SE_ERR_DDEFAIL
            GetShellExecuteError = "DDE operation failed"
        Case SE_ERR_DDEBUSY
            GetShellExecuteError = "DDE operation busy"
        Case SE_ERR_NOASSOC
            GetShellExecuteError = "No application associated with this file type"
        Case Else
            GetShellExecuteError = "Unknown error (code " & errorCode & ")"
    End Select
End Function

'=======================================================
' END OF MODULE
'=======================================================
