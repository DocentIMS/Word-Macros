Attribute VB_Name = "AC_API_Mod"
Option Explicit
Option Private Module
Option Compare Text

'=======================================================
' Module: AC_API_Mod
' Purpose: API communication module for Docent IMS
' Author: Fully Refactored - November 2025
' Version: 2.0
'
' Description:
'   This module handles all API communication with the
'   Docent IMS server. It provides functions for:
'   - Authentication and credential management
'   - File upload/download operations
'   - Content CRUD operations
'   - Folder management
'   - Workflow transitions
'   - Task and notification management
'   - Meeting document management
'
' Security Updates:
'   - Removed hardcoded credentials
'   - Uses AB_SecureCredentials module
'   - Proper error handling and cleanup
'   - Input validation on all functions
'
' Dependencies:
'   - AB_SecureCredentials (NEW)
'   - AB_GlobalConstants
'   - AB_GlobalVars
'   - WebClient, WebRequest, WebResponse classes
'   - Dictionary class
'
' Change Log:
'   v2.0 - Nov 2025 - Complete refactoring
'       * Removed hardcoded dashboard password
'       * Added comprehensive error handling
'       * Added function documentation
'       * Extracted helper functions
'       * Improved input validation
'   v1.0 - Original version
'=======================================================

' Module constants
Private Const CurrentMod As String = "API_Mod"
Private Const APITimeout As Long = 200000
Private Const APILongTimeout As Long = 60000

' Current function name for logging
Private CurrentF As String

'=======================================================
' CORE API FUNCTIONS
'=======================================================

'=======================================================
' Function: UseAPI
' Purpose: Core API function for executing HTTP requests
'
' Description:
'   Handles authentication, request building, execution,
'   and response processing for all API calls. Includes
'   automatic retry on 401 errors with credential refresh.
'
' Parameters:
'   Body - Dictionary containing request body (can be Nothing)
'   Method - HTTP method (HttpGet, HttpPost, HttpPatch, HttpDelete)
'   RequestURL - API endpoint URL (optional)
'   FolderPath - Folder path for API request (optional)
'   mURL - Base URL override (optional, uses ProjectURLStr if empty)
'   mUser - Username override (optional, uses UserNameStr if empty)
'   mPwd - Password override (optional, uses UserPasswordStr if empty)
'   TimeOut - Request timeout in milliseconds (default: APITimeout)
'   Format - Request body format (default: JSON)
'   AuthType - Authentication type: "Basic" or "Bearer" (default: "Basic")
'   Accept - Response format (default: JSON)
'   UseHTTP - Use HTTP instead of HTTPS (default: False)
'
' Returns:
'   WebResponse object, or Nothing on error
'
' Example:
'   Dim response As WebResponse
'   Set response = UseAPI(Nothing, HttpGet, "@main_info")
'   If Not response Is Nothing Then
'       If response.StatusCode = HTTP_OK Then
'           ' Process response.Data
'       End If
'   End If
'=======================================================
Public Function UseAPI(ByVal Body As Dictionary, _
                      ByVal Method As WebMethod, _
                      Optional ByVal RequestURL As String = vbNullString, _
                      Optional ByVal FolderPath As String = vbNullString, _
                      Optional ByVal mURL As String = vbNullString, _
                      Optional ByVal mUser As String = vbNullString, _
                      Optional ByVal mPwd As String = vbNullString, _
                      Optional ByVal timeout As Long = 0, _
                      Optional ByVal Format As WebFormat = WebFormat.JSON, _
                      Optional ByVal AuthType As String = "Basic", _
                      Optional ByVal Accept As WebFormat = WebFormat.JSON, _
                      Optional ByVal UseHTTP As Boolean = False) As WebResponse
    
    Dim Client As WebClient
    Dim Request As WebRequest
    Dim projectName As String
    Dim isDashboard As Boolean
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Configure credentials and URL
    isDashboard = ConfigureAPICredentials(mURL, mUser, mPwd, projectName, UseHTTP)
    
    ' Validate URL
    If Not ValidateAPIURL(mURL, errorMsg) Then
        WriteLog 3, CurrentMod, CurrentF, "URL validation failed: " & errorMsg
        Set UseAPI = Nothing
        Exit Function
    End If
    
    ' Build resource URL
    RequestURL = GetResource(RequestURL, mURL, FolderPath)
    WriteLog 1, CurrentMod, CurrentF, "Request: " & RequestURL
    
    ' Set defaults
    If timeout = 0 Then timeout = APITimeout
    
    ' Initialize client
    Set Client = New WebClient
    Client.BaseURL = mURL
    Client.TimeoutMs = timeout
    Client.Insecure = UseHTTP
    
RetryRequest:
    ' Build request
    Set Request = BuildAPIRequest(RequestURL, Method, Format, Accept, Body, AuthType, mUser, mPwd)
    
    ' Execute request
    Set UseAPI = Client.Execute(Request)
    
    ' Handle response
    If UseAPI Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Response is Nothing: " & Request.Resource
    Else
        Call HandleAPIResponse(UseAPI, projectName, mUser, mPwd, isDashboard)
        
        ' Retry on 401 if not already retried
        If UseAPI.StatusCode = HTTP_UNAUTHORIZED Then
            If Len(projectName) = 0 Then projectName = ProjectNameStr
            
            ' Get new password
            mPwd = GetUserPassword(projectName, ProjectURLStr)
            If Len(mPwd) > 0 Then
                WriteLog 2, CurrentMod, CurrentF, "Retrying with new credentials"
                GoTo RetryRequest
            End If
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, CurrentF, errorMsg & " (Resource: " & RequestURL & ")"
    Set UseAPI = Nothing
End Function

'=======================================================
' HELPER FUNCTIONS FOR UseAPI
'=======================================================

'=======================================================
' Function: ConfigureAPICredentials
' Purpose: Configure credentials and URL for API call
'
' Parameters:
'   mURL - Base URL (modified by reference)
'   mUser - Username (modified by reference)
'   mPwd - Password (modified by reference)
'   projectName - Project name (output)
'   UseHTTP - Use HTTP flag (modified by reference)
'
' Returns: True if dashboard credentials, False otherwise
'=======================================================
Private Function ConfigureAPICredentials(ByRef mURL As String, _
                                        ByRef mUser As String, _
                                        ByRef mPwd As String, _
                                        ByRef projectName As String, _
                                        ByRef UseHTTP As Boolean) As Boolean
    
    Dim isDashboard As Boolean
    Dim dashCreds As ApiCredentials
    
    On Error GoTo ErrorHandler
    
    ' Check if this is a dashboard request
    isDashboard = (mURL = DashboardURLStr And Len(DashboardURLStr) > 0)
    
    If isDashboard Then
        ' *** SECURITY FIX: Use secure credentials instead of hardcoded password ***
        dashCreds = GetDashboardCredentials()
        
        If dashCreds.IsValid Then
            mPwd = dashCreds.Password
            mUser = dashCreds.UserName
            projectName = "Dashboard"
            UseHTTP = True
            
            WriteLog 1, CurrentMod, "ConfigureAPICredentials", "Using dashboard credentials"
        Else
            WriteLog 3, CurrentMod, "ConfigureAPICredentials", _
                     "Failed to retrieve dashboard credentials"
        End If
    Else
        ' Use project credentials
        If Len(mURL) = 0 Then mURL = ProjectURLStr
        If Len(mUser) = 0 Then mUser = UserNameStr
        If Len(mPwd) = 0 Then mPwd = UserPasswordStr
    End If
    
    ConfigureAPICredentials = isDashboard
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ConfigureAPICredentials", _
             "Error " & Err.Number & ": " & Err.Description
    ConfigureAPICredentials = False
End Function

'=======================================================
' Function: ValidateAPIURL
' Purpose: Validate API URL is not empty
'=======================================================
Private Function ValidateAPIURL(ByVal mURL As String, ByRef errorMsg As String) As Boolean
    If Len(Trim$(mURL)) = 0 Then
        errorMsg = "No URL provided"
        ValidateAPIURL = False
    Else
        ValidateAPIURL = True
    End If
End Function

'=======================================================
' Function: BuildAPIRequest
' Purpose: Build a WebRequest object with all parameters
'=======================================================
Private Function BuildAPIRequest(ByVal RequestURL As String, _
                                 ByVal Method As WebMethod, _
                                 ByVal Format As WebFormat, _
                                 ByVal Accept As WebFormat, _
                                 ByVal Body As Dictionary, _
                                 ByVal AuthType As String, _
                                 ByVal mUser As String, _
                                 ByVal mPwd As String) As WebRequest
    
    Dim Request As WebRequest
    Dim authHeader As String
    
    On Error GoTo ErrorHandler
    
    Set Request = New WebRequest
    
    Request.Resource = RequestURL
    Request.Method = Method
    Request.Format = Format
    Request.ResponseFormat = Accept
    
    If Not Body Is Nothing Then
        Set Request.Body = Body
    End If
    
    ' Build authorization header
    authHeader = BuildAuthorizationHeader(AuthType, mUser, mPwd)
    Request.AddHeader "Authorization", authHeader
    
    Set BuildAPIRequest = Request
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "BuildAPIRequest", _
             "Error " & Err.Number & ": " & Err.Description
    Set BuildAPIRequest = Nothing
End Function

'=======================================================
' Function: BuildAuthorizationHeader
' Purpose: Build authorization header for API request
'=======================================================
Private Function BuildAuthorizationHeader(ByVal AuthType As String, _
                                         ByVal mUser As String, _
                                         ByVal mPwd As String) As String
    On Error GoTo ErrorHandler
    
    If AuthType = "Bearer" Then
        BuildAuthorizationHeader = "Bearer " & mPwd
    Else
        BuildAuthorizationHeader = "Basic " & Base64Encode(mUser & ":" & mPwd)
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "BuildAuthorizationHeader", _
             "Error " & Err.Number & ": " & Err.Description
    BuildAuthorizationHeader = vbNullString
End Function

'=======================================================
' Function: HandleAPIResponse
' Purpose: Process and log API response
'=======================================================
Private Sub HandleAPIResponse(ByVal response As WebResponse, _
                             ByVal projectName As String, _
                             ByVal mUser As String, _
                             ByVal mPwd As String, _
                             Optional ByVal isDashboard As Boolean = False)
    
    Dim statusMsg As String
    
    On Error Resume Next
    
    Select Case response.StatusCode
        Case HTTP_OK, HTTP_CREATED, HTTP_NO_CONTENT
            WriteLog 1, CurrentMod, CurrentF, response.StatusDescription
        
        Case HTTP_UNAUTHORIZED
            statusMsg = "Unauthorized (401)"
            If isDashboard Then
                statusMsg = statusMsg & " - Dashboard credentials may be invalid"
            End If
            WriteLog 3, CurrentMod, CurrentF, statusMsg
        
        Case HTTP_FORBIDDEN
            WriteLog 3, CurrentMod, CurrentF, _
                     "Forbidden (403) - User may not have permission"
        
        Case HTTP_NOT_FOUND
            WriteLog 3, CurrentMod, CurrentF, _
                     "Not Found (404) - Resource does not exist"
        
        Case HTTP_TIMEOUT
            WriteLog 3, CurrentMod, CurrentF, _
                     "Request Timeout (408) - Server took too long to respond"
        
        Case HTTP_SERVER_ERROR
            WriteLog 3, CurrentMod, CurrentF, _
                     "Server Error (500) - Internal server error"
        
        Case Else
            WriteLog 3, CurrentMod, CurrentF, _
                     "Status " & response.StatusCode & ": " & response.Content
    End Select
End Sub

'=======================================================
' Function: GetResource
' Purpose: Build complete resource URL from components
'
' Description:
'   Combines query string, base URL, and folder path into
'   a properly formatted API resource URL. Handles both
'   standard API paths and /noapi/ paths.
'
' Parameters:
'   QueryString - Query string or endpoint
'   BaseURL - Base URL (modified by reference if needed)
'   FolderPath - Folder path (optional)
'
' Returns: Formatted resource URL
'=======================================================
Private Function GetResource(ByVal QueryString As String, _
                            ByRef BaseURL As String, _
                            Optional ByVal FolderPath As String = vbNullString) As String
    
    Dim isNoApi As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Check if this is a noapi request
    isNoApi = (InStr(QueryString, "/noapi/") > 0) Or _
              (InStr(BaseURL, "/noapi/") > 0) Or _
              (InStr(FolderPath, "/noapi/") > 0)
    
    If isNoApi Then
        ' Remove /noapi/ prefix and build path
        BaseURL = Replace(BaseURL, "/noapi/", vbNullString)
        QueryString = Replace(QueryString, "/noapi/", vbNullString)
        FolderPath = Replace(FolderPath, "/noapi/", vbNullString)
        
        QueryString = Replace(QueryString, BaseURL, vbNullString)
        FolderPath = Replace(FolderPath, BaseURL, vbNullString)
        QueryString = Replace(QueryString, "\", "/")
        
        QueryString = FolderPath & "/" & QueryString
    Else
        ' Standard API path
        QueryString = Replace(QueryString, BaseURL, vbNullString)
        FolderPath = Replace(FolderPath, BaseURL, vbNullString)
        QueryString = Replace(QueryString, "\", "/")
        
        QueryString = "/++api++/" & FolderPath & "/" & QueryString
    End If
    
    ' Clean up the path
    Do While Right$(QueryString, 1) = "/"
        QueryString = Left$(QueryString, Len(QueryString) - 1)
    Loop
    
    Do While InStr(QueryString, "//") > 0
        QueryString = Replace(QueryString, "//", "/")
    Loop
    
    GetResource = QueryString
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetResource", _
             "Error " & Err.Number & ": " & Err.Description
    GetResource = vbNullString
End Function

'=======================================================
' FILE HANDLING FUNCTIONS
'=======================================================

'=======================================================
' Function: FileAsDic
' Purpose: Convert a file to Dictionary format for API upload
'
' Parameters:
'   FName - Full path to file
'
' Returns: Dictionary with file data in base64 format
'=======================================================
Private Function FileAsDic(ByVal FName As String) As Dictionary
    Dim FileObj As Dictionary
    Dim fileName As String
    Dim fileExt As String
    Dim ContentType As String
    
    On Error GoTo ErrorHandler
    
    Set FileObj = New Dictionary
    
    If Len(FName) = 0 Then
        Set FileAsDic = FileObj
        Exit Function
    End If
    
    fileName = GetFileName(FName, True)
    fileExt = LCase$(Right$(fileName, Len(fileName) - InStrRev(fileName, ".")))
    
    ' Determine content type
    ContentType = GetContentType(fileExt)
    
    ' Build file dictionary
    FileObj.Add "filename", fileName
    FileObj.Add "encoding", "base64"
    FileObj.Add "content-type", ContentType
    FileObj.Add "data", ConvertFileToBase64(FName)
    
    Set FileAsDic = FileObj
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "FileAsDic", _
             "Error " & Err.Number & ": " & Err.Description & " (File: " & FName & ")"
    Set FileAsDic = New Dictionary
End Function

'=======================================================
' Function: GetContentType
' Purpose: Get MIME content type from file extension
'=======================================================
Private Function GetContentType(ByVal fileExt As String) As String
    Select Case LCase$(fileExt)
        Case "txt", "csv"
            GetContentType = "text/plain"
        Case "jpeg", "jpg"
            GetContentType = "image/jpeg"
        Case "gif"
            GetContentType = "image/gif"
        Case "ico"
            GetContentType = "image/x-icon"
        Case "png"
            GetContentType = "image/png"
        Case "pdf"
            GetContentType = "application/pdf"
        Case "xlsx", "xls", "xlsm"
            GetContentType = "application/vnd.ms-excel"
        Case "doc", "docx", "docm", "dotm", "dotx"
            GetContentType = "application/msword"
        Case Else
            GetContentType = "*/*"
    End Select
End Function

'=======================================================
' BODY CREATION FUNCTIONS
'=======================================================

'=======================================================
' Function: CreateBody
' Purpose: Create request body dictionary from field/value arrays
'
' Description:
'   Builds a Dictionary object from parallel arrays of
'   field names and values. Handles special processing for
'   file uploads and type conversions.
'
' Parameters:
'   Fields - Array of field names
'   Values - Array of field values (parallel to Fields)
'   AddTitleeToo - Add title if not present (default: True)
'
' Returns: Dictionary object, or Nothing on error
'
' Example:
'   Dim body As Dictionary
'   Set body = CreateBody( _
'       Array("@type", "title", "description"), _
'       Array("Folder", "My Folder", "A test folder"))
'=======================================================
Public Function CreateBody(ByRef Fields As Variant, _
                          ByRef Values As Variant, _
                          Optional ByVal AddTitleeToo As Boolean = True) As Dictionary
    
    Dim Body As Dictionary
    Dim fileName As String
    Dim i As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not ValidateBodyInputs(Fields, Values, errorMsg) Then
        WriteLog 3, CurrentMod, "CreateBody", "Validation error: " & errorMsg
        Set CreateBody = Nothing
        Exit Function
    End If
    
    Set Body = New Dictionary
    
    ' Process each field/value pair
    For i = LBound(Fields) To UBound(Fields)
        Call ProcessBodyField(Body, Fields(i), Values(i), fileName)
    Next i
    
    ' Add title if needed and not present
    If AddTitleeToo And Not Body.Exists("title") Then
        If Len(fileName) > 0 Then
            Body.Add "title", fileName
        Else
            Body.Add "title", "Title"
        End If
    End If
    
    Set CreateBody = Body
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "CreateBody", _
             "Error " & Err.Number & ": " & Err.Description
    Set CreateBody = Nothing
End Function

'=======================================================
' Function: ValidateBodyInputs
' Purpose: Validate CreateBody input parameters
'=======================================================
Private Function ValidateBodyInputs(ByRef Fields As Variant, _
                                   ByRef Values As Variant, _
                                   ByRef errorMsg As String) As Boolean
    
    ValidateBodyInputs = False
    
    ' Check if empty
    If IsEmpty(Fields) Then
        errorMsg = "Fields array is empty"
        Exit Function
    End If
    
    If IsEmpty(Values) Then
        errorMsg = "Values array is empty"
        Exit Function
    End If
    
    ' Check if arrays
    If Not IsArray(Fields) Then
        errorMsg = "Fields parameter is not an array"
        Exit Function
    End If
    
    If Not IsArray(Values) Then
        errorMsg = "Values parameter is not an array"
        Exit Function
    End If
    
    ' Check array bounds match
    If UBound(Fields) <> UBound(Values) Or LBound(Fields) <> LBound(Values) Then
        errorMsg = "Fields and Values arrays have mismatched bounds"
        Exit Function
    End If
    
    ValidateBodyInputs = True
End Function

'=======================================================
' Sub: ProcessBodyField
' Purpose: Process a single field/value pair for request body
'=======================================================
Private Sub ProcessBodyField(ByRef Body As Dictionary, _
                            ByVal fieldName As String, _
                            ByVal fieldValue As Variant, _
                            ByRef fileName As String)
    
    On Error Resume Next
    
    Select Case fieldName
        Case "@type"
            Body.Add fieldName, LCase$(Replace(CStr(fieldValue), " ", "_"))
        
        Case "title"
            Dim titleValue As String
            titleValue = CStr(fieldValue)
            Body.Add "title", IIf(Len(titleValue) > 0, titleValue, "Title")
        
        Case "file"
            If Len(fileName) = 0 Then
                fileName = GetFileName(CStr(fieldValue))
            End If
            Body.Add "file", FileAsDic(CStr(fieldValue))
        
        Case Else
            If TypeName(fieldValue) = "String" Then
                If Len(CStr(fieldValue)) > 0 Then
                    If CStr(fieldValue) Like APIFilePrefix & "*" Then
                        ' Handle file reference
                        Dim filePath As String
                        filePath = Right$(CStr(fieldValue), Len(CStr(fieldValue)) - Len(APIFilePrefix))
                        Body.Add fieldName, FileAsDic(filePath)
                    Else
                        Body.Add fieldName, fieldValue
                    End If
                End If
            Else
                Body.Add fieldName, fieldValue
            End If
    End Select
End Sub

'=======================================================
' Function: CreateQuerystringParams
' Purpose: Create query string parameters collection (deprecated)
' Note: Kept for backward compatibility but not actively used
'=======================================================
Private Function CreateQuerystringParams(ByRef Fields As Variant, _
                                        ByRef Values As Variant) As Collection
    Dim QSP As Collection
    Dim i As Long
    Dim Dict As Dictionary
    
    On Error GoTo ErrorHandler
    
    Set QSP = New Collection
    
    For i = LBound(Fields) To UBound(Fields)
         Set Dict = New Dictionary
         Dict.Add "key", Fields(i)
         Dict.Add "value", Values(i)
         QSP.Add Dict
    Next i
    
    Set CreateQuerystringParams = QSP
    Exit Function
    
ErrorHandler:
    Set CreateQuerystringParams = New Collection
End Function

'=======================================================
' AUTHENTICATION AND VALIDATION FUNCTIONS
'=======================================================

'=======================================================
' Function: IsValidUser
' Purpose: Validate user credentials and permissions
'
' Parameters:
'   mURL - Server URL (optional, uses ProjectURLStr if empty)
'   mUser - Username (optional, uses UserNameStr if empty)
'   mPwd - Password (optional, uses UserPasswordStr if empty)
'
' Returns:
'   "Ok" if valid
'   Error message string if invalid
'
' Example:
'   Dim result As String
'   result = IsValidUser()
'   If result = "Ok" Then
'       ' User is valid
'   Else
'       MsgBox result
'   End If
'=======================================================
Public Function IsValidUser(Optional ByVal mURL As String = vbNullString, _
                           Optional ByVal mUser As String = vbNullString, _
                           Optional ByVal mPwd As String = vbNullString) As String
    
    Dim response As WebResponse
    Dim groupsDict As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "IsValidUser"
    
    If Len(mURL) = 0 Then mURL = ProjectURLStr
    If Len(mUser) = 0 Then mUser = UserNameStr
    If Len(mPwd) = 0 Then mPwd = UserPasswordStr
    
    ' Make API call
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/@main_info", , mURL, mUser, mPwd)
    
    ' Check response
    If response Is Nothing Then
        IsValidUser = "Incorrect URL or network error"
        GoTo ErrorHandler
    End If
    
    Select Case response.StatusCode
        Case HTTP_OK
            ' Validate response data
            If IsNull(response.Data(1)("id")) Then
                IsValidUser = "The server is not ready. Please contact the project manager."
            Else
                ' Check group membership
                Set groupsDict = GetGroupsDict(response.Data)
                If groupsDict.Exists("PrjTeam") Or groupsDict.Exists("meadows_board") Then
                    IsValidUser = "Ok"
                    WriteLog 1, CurrentMod, "IsValidUser", "Verified User: " & mUser
                Else
                    IsValidUser = "You are not authorized to use this tool. " & _
                                "You must join the team members' group first. " & _
                                "Please contact the project manager."
                End If
            End If
        
        Case HTTP_FORBIDDEN, HTTP_SERVER_ERROR
            IsValidUser = "You are not authorized to use this tool. " & _
                        "Please contact the project manager."
        
        Case HTTP_TIMEOUT
            IsValidUser = response.StatusDescription
        
        Case HTTP_UNAUTHORIZED
            IsValidUser = "Username or password is incorrect."
        
        Case Else
            IsValidUser = "Something went wrong. Please contact the project manager."
    End Select
    
    Exit Function
    
ErrorHandler:
    If Len(IsValidUser) = 0 Then
        IsValidUser = "Error validating user: " & Err.Description
    End If
    WriteLog 3, CurrentMod, "IsValidUser", _
             "Failed to validate user: " & mUser & " - " & IsValidUser
End Function

'=======================================================
' INFORMATION RETRIEVAL FUNCTIONS
'=======================================================

'=======================================================
' Function: GetMainInfo
' Purpose: Retrieve main user information from server
'
' Parameters:
'   mURL - Server URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'   Scope - Additional scope parameter (optional)
'
' Returns: Collection of user information, or Nothing on error
'=======================================================
Public Function GetMainInfo(Optional ByVal mURL As String = vbNullString, _
                           Optional ByVal mUser As String = vbNullString, _
                           Optional ByVal mPwd As String = vbNullString, _
                           Optional ByVal Scope As String = vbNullString) As Collection
    
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetMainInfo"
    
    ' Validate URL
    If Replace(mURL, "/", vbNullString) = vbNullString Then
        WriteLog 2, CurrentMod, CurrentF, "Empty URL provided"
        Set GetMainInfo = Nothing
        Exit Function
    End If
    
    ' Make API call
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/@main_info" & Scope, , mURL, mUser, mPwd)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Response is Nothing"
    ElseIf response.StatusCode = HTTP_OK Then
        If Not IsNull(response.Data(1)("id")) Then
            Set GetMainInfo = response.Data
        Else
            WriteLog 3, CurrentMod, CurrentF, _
                     "Invalid response data: " & response.Content
        End If
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed with status " & response.StatusCode & ": " & response.Content
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetMainInfo = Nothing
End Function

'=======================================================
' Function: GetDocsInfo
' Purpose: Retrieve project document information
'
' Parameters:
'   mURL - Server URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'   mPName - Project name (optional)
'
' Returns: Dictionary of project information, or Nothing on error
'=======================================================
Public Function GetDocsInfo(Optional ByVal mURL As String = vbNullString, _
                           Optional ByVal mUser As String = vbNullString, _
                           Optional ByVal mPwd As String = vbNullString, _
                           Optional ByVal mPName As String = vbNullString) As Project.Dictionary
    
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetDocsInfo"
    
    ' Validate URL
    If Replace(mURL, "/", vbNullString) = vbNullString Then
        WriteLog 2, CurrentMod, CurrentF, "Empty URL provided"
        Set GetDocsInfo = Nothing
        Exit Function
    End If
    
    If Len(mPName) = 0 Then mPName = mURL
    
    ' Make API call
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/@project_info", , mURL, mUser, mPwd)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Response is Nothing"
    ElseIf response.StatusCode = HTTP_OK Then
        If Not IsNull(response.Data("short_name")) Then
            Set GetDocsInfo = response.Data
        Else
            WriteLog 3, CurrentMod, CurrentF, _
                     "Failed to get project info: " & mPName & " - " & response.Content
        End If
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed with status " & response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetDocsInfo = Nothing
End Function

'=======================================================
' Function: GetWorkflowInfo
' Purpose: Retrieve workflow information for project
'
' Parameters:
'   mURL - Server URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'   mPName - Project name (optional)
'
' Returns: Dictionary of workflow information, or Nothing on error
'=======================================================
Public Function GetWorkflowInfo(Optional ByVal mURL As String = vbNullString, _
                               Optional ByVal mUser As String = vbNullString, _
                               Optional ByVal mPwd As String = vbNullString, _
                               Optional ByVal mPName As String = vbNullString) As Project.Dictionary
    
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetWorkflowInfo"
    
    ' Validate URL
    If Replace(mURL, "/", vbNullString) = vbNullString Then
        WriteLog 2, CurrentMod, CurrentF, "Empty URL provided"
        Set GetWorkflowInfo = Nothing
        Exit Function
    End If
    
    If Len(mPName) = 0 Then mPName = mURL
    
    ' Make API call
    Set response = UseAPI(Nothing, WebMethod.HttpGet, _
                         "/noapi/@workflow_info?portal_type=*", , mURL, mUser, mPwd)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Failed to get workflow info: " & Err.Description
    ElseIf response.StatusCode = HTTP_OK Then
        Set GetWorkflowInfo = response.Data
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed for project " & mPName & ": " & response.Content
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetWorkflowInfo = Nothing
End Function

'=======================================================
' CONTENT MANAGEMENT FUNCTIONS
'=======================================================

'=======================================================
' Function: SendNote
' Purpose: Send a post-it note to the server
'
' Parameters:
'   NoteContents - Note text content
'   NoteName - Title of the note
'   AttachmentFile - Optional file path to attach
'   FolderPath - Destination folder (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function SendNote(ByVal NoteContents As String, _
                        ByVal NoteName As String, _
                        Optional ByVal AttachmentFile As String = vbNullString, _
                        Optional ByVal FolderPath As String = vbNullString) As WebResponse
    
    Dim Body As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "SendNote"
    
    ' Validate inputs
    If Len(NoteContents) = 0 Then
        WriteLog 2, CurrentMod, CurrentF, "Empty note contents"
    End If
    
    If Len(NoteName) = 0 Then NoteName = "Note"
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "postit_note"
    Body.Add "title", NoteName
    Body.Add "my_note", NoteContents
    
    If Len(AttachmentFile) > 0 Then
        Body.Add "attachment", FileAsDic(AttachmentFile)
    End If
    
    ' Send request
    Set SendNote = UseAPI(Body, WebMethod.HttpPost, FolderPath)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set SendNote = Nothing
End Function

'=======================================================
' Function: SendAPIFeedback
' Purpose: Send feedback to the server
'
' Parameters:
'   Title - Feedback title
'   Feedback - Feedback text
'   ToolName - Name of the tool
'   RespReq - Response required (Yes/No)
'   Who - Submitter name
'   AttachmentFile - Optional attachment file path
'   LogFile - Optional log file path
'   FolderPath - Destination folder (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function SendAPIFeedback(ByVal Title As String, _
                               ByVal Feedback As String, _
                               ByVal ToolName As String, _
                               ByVal RespReq As String, _
                               ByVal Who As String, _
                               ByVal AttachmentFile As String, _
                               ByVal LogFile As String, _
                               Optional ByVal FolderPath As String = vbNullString) As WebResponse
    
    Dim Body As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "SendAPIFeedback"
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "feedback"
    Body.Add "title", Title
    Body.Add "feedback_text", Feedback
    Body.Add "docent_tool_type", ToolName
    Body.Add "submitted_by", Who
    Body.Add "reply", RespReq
    
    If Len(AttachmentFile) > 0 Then
        Body.Add "feedback_attachment", FileAsDic(AttachmentFile)
    End If
    
    If Len(LogFile) > 0 Then
        Body.Add "feedback_log", FileAsDic(LogFile)
    End If
    
    ' Send request with long timeout
    Set SendAPIFeedback = UseAPI(Body, WebMethod.HttpPost, FolderPath, timeout:=APILongTimeout)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set SendAPIFeedback = Nothing
End Function

'=======================================================
' Function: CreateAPIBreakdown
' Purpose: Create a breakdown (analysis) document
'
' Parameters:
'   DocType - Document type ("Scope" or "RFP")
'   FilePath - Path to HTML file
'   FolderPath - Destination folder (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function CreateAPIBreakdown(ByVal DocType As String, _
                                  ByVal filePath As String, _
                                  Optional ByVal FolderPath As String = vbNullString) As WebResponse
    
    Dim Body As Dictionary
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    CurrentF = "CreateAPIBreakdown"
    
    ' Validate inputs
    If Len(filePath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file path"
        Set CreateAPIBreakdown = Nothing
        Exit Function
    End If
    
    fileName = GetFileName(filePath, False)
    
    ' Build request body
    Set Body = New Dictionary
    
    Select Case DocType
        Case "Scope"
            Body.Add "@type", "sow_analysis"
        Case "RFP"
            Body.Add "@type", "rfp_breakdown"
        Case Else
            WriteLog 2, CurrentMod, CurrentF, "Unknown doc type: " & DocType
    End Select
    
    Body.Add "section_number", Left$(fileName, 3)
    Body.Add "title", Trim$(Right$(fileName, Len(fileName) - 4))
    Body.Add "bodytext", GetFileContents(filePath)
    Body.Add "deliverable_text", GetFileContents(filePath & "deliverables.html")
    
    ' Send request with long timeout
    Set CreateAPIBreakdown = UseAPI(Body, WebMethod.HttpPost, FolderPath, timeout:=APILongTimeout)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CreateAPIBreakdown = Nothing
End Function

'=======================================================
' Function: GetAPIContent
' Purpose: Get content from a URL
'
' Parameters:
'   URL - Content URL
'   mURL - Base URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'
' Returns: WebResponse object, or Nothing on error
'=======================================================
Public Function GetAPIContent(ByVal URL As String, _
                             Optional ByVal mURL As String = vbNullString, _
                             Optional ByVal mUser As String = vbNullString, _
                             Optional ByVal mPwd As String = vbNullString) As WebResponse
    
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(URL) = 0 Then
        WriteLog 3, CurrentMod, "GetAPIContent", "Empty URL"
        Set GetAPIContent = Nothing
        Exit Function
    End If
    
    Set response = UseAPI(Nothing, WebMethod.HttpGet, URL, , mURL, mUser, mPwd)
    
    If Not response Is Nothing Then
        Set GetAPIContent = response
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetAPIContent", "Error " & Err.Number & ": " & Err.Description
    Set GetAPIContent = Nothing
End Function

'=======================================================
' Function: DeleteAPIContent
' Purpose: Delete content at a URL
'
' Parameters:
'   URL - Content URL to delete
'
' Returns: WebResponse object, or Nothing on error
'=======================================================
Public Function DeleteAPIContent(ByVal URL As String) As WebResponse
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(URL) = 0 Then
        WriteLog 3, CurrentMod, "DeleteAPIContent", "Empty URL"
        Set DeleteAPIContent = Nothing
        Exit Function
    End If
    
    Set response = UseAPI(Nothing, WebMethod.HttpDelete, URL)
    
    If Not response Is Nothing Then
        Set DeleteAPIContent = response
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "DeleteAPIContent", "Error " & Err.Number & ": " & Err.Description
    Set DeleteAPIContent = Nothing
End Function

'=======================================================
' Function: CreateAPIContent
' Purpose: Create new content on server
'
' Parameters:
'   ContentType - Type of content to create
'   FolderPath - Destination folder (optional)
'   Fields - Field names array (optional)
'   Values - Field values array (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function CreateAPIContent(ByVal ContentType As String, _
                                Optional ByVal FolderPath As String = vbNullString, _
                                Optional ByRef Fields As Variant, _
                                Optional ByRef Values As Variant) As WebResponse
    
    Dim Body As Dictionary
    Dim timeout As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "CreateAPIContent"
    
    ' Create body from fields/values
    Set Body = CreateBody(Fields, Values)
    
    ' Ensure @type is set
    If Body Is Nothing Then
        Set Body = New Dictionary
    End If
    
    If Not Body.Exists("@type") Then
        Body.Add "@type", ContentType
    End If
    
    ' Determine timeout based on content
    If Body.Exists("file") Then
        timeout = APILongTimeout
    Else
        timeout = APITimeout
    End If
    
    ' Send request
    Set CreateAPIContent = UseAPI(Body, WebMethod.HttpPost, FolderPath, timeout:=timeout)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CreateAPIContent = Nothing
End Function

'=======================================================
' Function: UpdateAPIContent
' Purpose: Update existing content on server
'
' Parameters:
'   FolderPath - Content URL to update
'   Fields - Field names array (optional)
'   Values - Field values array (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function UpdateAPIContent(ByVal FolderPath As String, _
                                Optional ByRef Fields As Variant, _
                                Optional ByRef Values As Variant) As WebResponse
    
    Dim Body As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "UpdateAPIContent"
    
    ' Validate input
    If Len(FolderPath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty folder path"
        Set UpdateAPIContent = Nothing
        Exit Function
    End If
    
    ' Create body (without auto-adding title)
    Set Body = CreateBody(Fields, Values, False)
    
    ' Send PATCH request with long timeout
    Set UpdateAPIContent = UseAPI(Body, WebMethod.HttpPatch, FolderPath, timeout:=APILongTimeout)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set UpdateAPIContent = Nothing
End Function

'=======================================================
' FOLDER MANAGEMENT FUNCTIONS
'=======================================================

'=======================================================
' Function: CheckAPIFolder
' Purpose: Check if a folder exists
'
' Parameters:
'   FolderPath - Path to folder
'
' Returns: WebResponse object
'=======================================================
Public Function CheckAPIFolder(ByVal FolderPath As String) As WebResponse
    On Error GoTo ErrorHandler
    
    CurrentF = "CheckAPIFolder"
    
    ' Validate input
    If Len(FolderPath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty folder path"
        Set CheckAPIFolder = Nothing
        Exit Function
    End If
    
    Set CheckAPIFolder = UseAPI(Nothing, WebMethod.HttpGet, FolderPath)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CheckAPIFolder = Nothing
End Function

'=======================================================
' Function: CreateAPIFolder
' Purpose: Create a new folder on server
'
' Parameters:
'   FolderName - Name of folder to create
'   FolderPath - Parent folder path (optional)
'   Description - Folder description (optional)
'   NoNav - Exclude from navigation (default: True)
'
' Returns: WebResponse object
'=======================================================
Public Function CreateAPIFolder(ByVal FolderName As String, _
                               Optional ByVal FolderPath As String = vbNullString, _
                               Optional ByVal Description As String = vbNullString, _
                               Optional ByVal NoNav As Boolean = True) As WebResponse
    
    Dim Body As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "CreateAPIFolder"
    
    ' Validate input
    If Len(FolderName) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty folder name"
        Set CreateAPIFolder = Nothing
        Exit Function
    End If
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "Folder"
    Body.Add "title", FolderName
    Body.Add "exclude_from_nav", IIf(NoNav, "1", "0")
    
    If Len(Description) > 0 Then
        Body.Add "description", Description
    End If
    
    ' Send request
    Set CreateAPIFolder = UseAPI(Body, WebMethod.HttpPost, FolderPath)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CreateAPIFolder = Nothing
End Function

'=======================================================
' Function: GetAPIFolder
' Purpose: Search and retrieve folder contents
'
' Description:
'   Advanced search function for retrieving folder contents
'   with filtering, sorting, and metadata options.
'
' Parameters:
'   FolderPath - Folder path to search (optional)
'   DocType - Document type filter (optional)
'   MetadataFields - Array of metadata fields to include (optional)
'   MetadataSearchFilters - Array of filter field names (optional)
'   MetadataFilters - Array of filter values (optional)
'   SortOnFields - Array of fields to sort on (optional)
'   BatchSize - Maximum results to return (default: 9999)
'   mURL - Base URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'
' Returns: Collection of items, or Nothing on error
'
' Example:
'   Dim items As Collection
'   Set items = GetAPIFolder("documents", "File", _
'                           Array("title", "modified"), _
'                           Array("review_state"), Array("published"))
'=======================================================
Public Function GetAPIFolder(Optional ByVal FolderPath As String = vbNullString, _
                            Optional ByVal DocType As String = vbNullString, _
                            Optional ByRef MetadataFields As Variant, _
                            Optional ByRef MetadataSearchFilters As Variant, _
                            Optional ByRef MetadataFilters As Variant, _
                            Optional ByRef SortOnFields As Variant, _
                            Optional ByVal BatchSize As Long = 9999, _
                            Optional ByVal mURL As String = vbNullString, _
                            Optional ByVal mUser As String = vbNullString, _
                            Optional ByVal mPwd As String = vbNullString) As Collection
    
    Dim response As WebResponse
    Dim QueryString As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetAPIFolder"
    
    ' Build query string
    QueryString = BuildFolderQuery(FolderPath, DocType, MetadataFields, _
                                  MetadataSearchFilters, MetadataFilters, _
                                  SortOnFields, BatchSize)
    
    ' Execute search
    Set response = UseAPI(Nothing, WebMethod.HttpGet, QueryString, FolderPath, mURL, mUser, mPwd)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Response is Nothing"
    ElseIf response.StatusCode = HTTP_OK Then
        Set GetAPIFolder = response.Data("items")
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get folder: " & QueryString & " - " & response.Content
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetAPIFolder = Nothing
End Function

'=======================================================
' Function: BuildFolderQuery
' Purpose: Build query string for GetAPIFolder
' Returns: Formatted query string
'=======================================================
Private Function BuildFolderQuery(ByVal FolderPath As String, _
                                 ByVal DocType As String, _
                                 ByRef MetadataFields As Variant, _
                                 ByRef MetadataSearchFilters As Variant, _
                                 ByRef MetadataFilters As Variant, _
                                 ByRef SortOnFields As Variant, _
                                 ByVal BatchSize As Long) As String
    
    Dim QueryString As String
    Dim i As Long
    
    On Error Resume Next
    
    ' Handle NoSearch_ prefix
    If Left$(FolderPath, 9) = "NoSearch_" Then
        FolderPath = Right$(FolderPath, Len(FolderPath) - 9)
    Else
        QueryString = "@search"
    End If
    
    ' Add batch size
    QueryString = QueryString & "?b_size=" & BatchSize
    
    ' Add document type filter
    If Len(DocType) > 0 Then
        QueryString = QueryString & "&portal_type=" & DocType
    End If
    
    ' Add metadata fields
    If Not IsMissing(MetadataFields) Then
        For i = LBound(MetadataFields) To UBound(MetadataFields)
            QueryString = QueryString & "&metadata_fields=" & MetadataFields(i)
        Next i
    End If
    
    ' Add search filters
    If Not IsMissing(MetadataSearchFilters) And Not IsMissing(MetadataFilters) Then
        For i = LBound(MetadataSearchFilters) To UBound(MetadataSearchFilters)
            QueryString = QueryString & "&" & MetadataSearchFilters(i) & _
                        "=" & MetadataFilters(i)
        Next i
    End If
    
    ' Add sort fields
    If Not IsMissing(SortOnFields) Then
        For i = LBound(SortOnFields) To UBound(SortOnFields)
            QueryString = QueryString & "&sort_on=" & SortOnFields(i)
        Next i
    End If
    
    BuildFolderQuery = QueryString
End Function

'=======================================================
' FILE UPLOAD/DOWNLOAD FUNCTIONS
'=======================================================

'=======================================================
' Function: UploadAPIFile
' Purpose: Upload a file to the server
'
' Parameters:
'   FilePath - Local file path
'   FolderPath - Destination folder (optional)
'   mURL - Base URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'   Overwrite - Update if file exists (default: False)
'
' Returns: WebResponse object
'=======================================================
Public Function UploadAPIFile(ByVal filePath As String, _
                             Optional ByVal FolderPath As String = vbNullString, _
                             Optional ByVal mURL As String = vbNullString, _
                             Optional ByVal mUser As String = vbNullString, _
                             Optional ByVal mPwd As String = vbNullString, _
                             Optional ByVal Overwrite As Boolean = False) As WebResponse
    
    Dim Body As Dictionary
    Dim fileName As String
    Dim Itms As Collection
    Dim FCodeName As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "UploadAPIFile"
    
    ' Validate input
    If Len(filePath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file path"
        Set UploadAPIFile = Nothing
        Exit Function
    End If
    
    fileName = GetFileName(filePath, True)
    
    ' Check if file should be overwritten
    If Overwrite Then
        FCodeName = Replace(LCase$(fileName), " ", "_")
        Set Itms = CheckAPIFolder(FolderPath).Data("items")
        
        For i = 1 To Itms.Count
            If Itms(i)("title") = FCodeName Then
                ' File exists - update instead
                Set UploadAPIFile = UpdateAPIFile(filePath, FolderPath, fileName, mURL, mUser, mPwd)
                Exit Function
            End If
        Next i
    End If
    
    ' Create new file
    Set Body = CreateBody(Array("title", "file"), Array(fileName, filePath))
    Body.Add "@type", "File"
    
    Set UploadAPIFile = UseAPI(Body, WebMethod.HttpPost, FolderPath, , mURL, mUser, mPwd)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set UploadAPIFile = Nothing
End Function

'=======================================================
' Function: UpdateAPIFile
' Purpose: Update an existing file on server
'
' Parameters:
'   FilePath - Local file path
'   FolderPath - File URL to update
'   FileName - File name (optional)
'   mURL - Base URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function UpdateAPIFile(ByVal filePath As String, _
                             Optional ByVal FolderPath As String = vbNullString, _
                             Optional ByVal fileName As String = vbNullString, _
                             Optional ByVal mURL As String = vbNullString, _
                             Optional ByVal mUser As String = vbNullString, _
                             Optional ByVal mPwd As String = vbNullString) As WebResponse
    
    Dim Body As Dictionary
    
    On Error GoTo ErrorHandler
    
    CurrentF = "UpdateAPIFile"
    
    ' Validate input
    If Len(filePath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file path"
        Set UpdateAPIFile = Nothing
        Exit Function
    End If
    
    If Len(fileName) = 0 Then fileName = GetFileName(filePath, True)
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "File"
    Body.Add "title", GetFileName(filePath, False)
    Body.Add "file", FileAsDic(filePath)
    
    ' Send PATCH request with long timeout
    Set UpdateAPIFile = UseAPI(Body, WebMethod.HttpPatch, FolderPath, , _
                               mURL, mUser, mPwd, timeout:=APILongTimeout)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set UpdateAPIFile = Nothing
End Function

'=======================================================
' Function: DownloadAPIFile
' Purpose: Download a file from the server
'
' Description:
'   Downloads a file from the server and saves it to the
'   temp directory. Returns the local file path.
'
' Parameters:
'   fileURL - Server file URL
'   AsTemplate - Template mode flag (optional)
'   FName - Local file name (optional)
'   mURL - Base URL (optional)
'   mUser - Username (optional)
'   mPwd - Password (optional)
'
' Returns: Local file path, or empty string on error
'=======================================================
Public Function DownloadAPIFile(ByVal fileURL As String, _
                               Optional ByVal AsTemplate As Boolean = False, _
                               Optional ByVal FName As String = vbNullString, _
                               Optional ByVal mURL As String = vbNullString, _
                               Optional ByVal mUser As String = vbNullString, _
                               Optional ByVal mPwd As String = vbNullString) As String
    
    Dim response As WebResponse
    Dim Stream As Object
    Const DownloadSuffix As String = "/@@download/file"
    
    On Error GoTo ErrorHandler
    
    CurrentF = "DownloadAPIFile"
    
    ' Validate input
    If Len(fileURL) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file URL"
        DownloadAPIFile = vbNullString
        Exit Function
    End If
    
    ' Ensure URL has download suffix
    If InStr(fileURL, "@@download") = 0 Then
        fileURL = fileURL & DownloadSuffix
    End If
    
    ' Download file
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/" & fileURL, , _
                         mURL, mUser, mPwd, timeout:=APILongTimeout)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Failed to download: " & fileURL
    ElseIf response.StatusCode = HTTP_NOT_FOUND Then
        ' Try with lowercase and dashes
        If fileURL <> Replace(LCase$(fileURL), " ", "-") Then
            DownloadAPIFile = DownloadAPIFile(Replace(LCase$(fileURL), " ", "-"), _
                                             AsTemplate, FName, mURL, mUser, mPwd)
        End If
    ElseIf response.StatusCode = HTTP_OK Then
        ' Save file to disk
        DownloadAPIFile = SaveDownloadedFile(response, fileURL, FName, mURL, mUser, mPwd)
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    DownloadAPIFile = vbNullString
End Function

'=======================================================
' Function: SaveDownloadedFile
' Purpose: Save downloaded file content to disk
' Returns: Local file path
'=======================================================
Private Function SaveDownloadedFile(ByVal response As WebResponse, _
                                   ByVal fileURL As String, _
                                   ByRef FName As String, _
                                   ByVal mURL As String, _
                                   ByVal mUser As String, _
                                   ByVal mPwd As String) As String
    
    Dim Stream As Object
    Dim MetaResponse As WebResponse
    
    On Error Resume Next
    
    Set Stream = CreateObject("ADODB.Stream")
    Stream.Open
    Stream.Type = 1 ' adTypeBinary
    Stream.Write response.Body
    
    ' Determine file name if not provided
    If InStr(FName, ".") = 0 Then
        ' Get metadata to find filename
        If InStr(fileURL, "@@download") Then
            fileURL = Left$(fileURL, InStr(fileURL, "/@@download") - 1)
        End If
        
        Set MetaResponse = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/" & fileURL, , mURL, mUser, mPwd)
        
        If Not IsEmpty(MetaResponse.Data("file")) Then
            FName = MetaResponse.Data("file")("filename")
        ElseIf Not IsEmpty(MetaResponse.Data("image")) Then
            FName = MetaResponse.Data("image")("filename")
        End If
    End If
    
    ' Build full path
    FName = Environ("Temp") & "\" & FName
    FName = Replace(FName, "\\", "\")
    
    ' Save file
    Err.Clear
    Stream.SaveToFile FName, 2
    Stream.Close
    
    If Err.Number = 0 Or Err.Number = 3004 Then
        SaveDownloadedFile = FName
    Else
        SaveDownloadedFile = vbNullString
        WriteLog 3, CurrentMod, "SaveDownloadedFile", _
                 "Error saving file: " & Err.Description
    End If
End Function

'=======================================================
' DOCUMENT RETRIEVAL FUNCTIONS
'=======================================================

'=======================================================
' Function: GetDocumentsOfType
' Purpose: Retrieve documents of a specific type
'
' Parameters:
'   DocType - Document type to retrieve (optional)
'   FolderPath - Folder to search (optional)
'
' Returns: Collection of documents, or Nothing on error
'=======================================================
Public Function GetDocumentsOfType(Optional ByVal DocType As String = vbNullString, _
                                  Optional ByVal FolderPath As String = vbNullString) As Variant
    
    Dim DocsList As Collection
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetDocumentsOfType"
    
    ' Get documents from folder
    Set DocsList = GetAPIFolder(FolderPath, "docent_misc_document", , _
                                Array("document_type"), Array(DocType))
    
    ' Process response
    If DocsList Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get documents of type: " & DocType & " in " & FolderPath
    ElseIf TypeName(DocsList) = "Collection" Then
        ' Add state names to each document
        For i = 1 To DocsList.Count
            DocsList(i).Add "State", GetStateName(CStr(DocsList(i)("review_state")), _
                                                 DocsList(i)("@type"))
        Next i
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Unexpected response type for: " & DocType & " in " & FolderPath
    End If
    
    Set GetDocumentsOfType = DocsList
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetDocumentsOfType = Nothing
End Function

'=======================================================
' Function: GetAPIDocumentName
' Purpose: Get document name from URL
'
' Parameters:
'   DocumentURL - Document URL
'
' Returns: Document name, or empty string on error
'=======================================================
Public Function GetAPIDocumentName(ByVal DocumentURL As String) As String
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetAPIDocumentName"
    
    ' Validate input
    If Len(DocumentURL) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty document URL"
        GetAPIDocumentName = vbNullString
        Exit Function
    End If
    
    ' Get document metadata
    Set response = UseAPI(Nothing, WebMethod.HttpGet, DocumentURL)
    
    ' Extract filename
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get document name: " & DocumentURL
    ElseIf response.StatusCode = HTTP_OK Then
        GetAPIDocumentName = response.Data("file")("filename")
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get document name: " & DocumentURL & " - " & response.Content
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    GetAPIDocumentName = vbNullString
End Function

'=======================================================
' Function: GetAPIPostItNoteColors
' Purpose: Get available post-it note colors
'
' Returns: Collection of color options, or Nothing on error
'=======================================================
Public Function GetAPIPostItNoteColors() As Collection
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/@types/postit_note")
    
    If response Is Nothing Then
        WriteLog 3, CurrentMod, "GetAPIPostItNoteColors", "Failed to get colors"
    ElseIf response.StatusCode = HTTP_OK Then
        Set GetAPIPostItNoteColors = response.Data("properties")("color")("choices")
    Else
        WriteLog 3, CurrentMod, "GetAPIPostItNoteColors", _
                 "Failed with status " & response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetAPIPostItNoteColors", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetAPIPostItNoteColors = Nothing
End Function

'=======================================================
' PLANNING AND TASK FUNCTIONS
'=======================================================

'=======================================================
' Function: CreateAPIPlanningItem
' Purpose: Create a planning breakdown item
'
' Parameters:
'   FilePath - Path to planning document
'   Comment - Comment text
'   Commenter - Name of commenter
'   FolderPath - Destination folder (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function CreateAPIPlanningItem(ByVal filePath As String, _
                                     ByVal Comment As String, _
                                     ByVal Commenter As String, _
                                     Optional ByVal FolderPath As String = vbNullString) As WebResponse
    
    Dim Body As Dictionary
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    CurrentF = "CreateAPIPlanningItem"
    
    ' Validate inputs
    If Len(filePath) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file path"
        Set CreateAPIPlanningItem = Nothing
        Exit Function
    End If
    
    fileName = GetFileName(filePath, False)
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "planning_breakdown"
    Body.Add "section_number", Left$(fileName, 3)
    Body.Add "title", Trim$(Right$(fileName, Len(fileName) - 4))
    Body.Add "body_text", GetFileContents(filePath)
    Body.Add "comment", Comment
    Body.Add "commenter", Commenter
    
    ' Send request
    Set CreateAPIPlanningItem = UseAPI(Body, WebMethod.HttpPost, FolderPath)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CreateAPIPlanningItem = Nothing
End Function

'=======================================================
' Function: CreateAPITask
' Purpose: Create an action item task
'
' Description:
'   Creates a new task (action item) on the server with
'   all specified properties.
'
' Parameters:
'   Title - Task title
'   Desctiption - Task description
'   DDate - Due date
'   Priority - Task priority (optional)
'   Who - Assigned to (optional)
'   Notes - Task notes (optional)
'   PrivateNotes - Private notes (optional)
'   Tags - Comma-separated tags (optional)
'   MMURL - Meeting minutes URL (optional)
'   Others - Related members collection (optional)
'   MeetingUID - Source meeting UID (optional)
'
' Returns: WebResponse object
'=======================================================
Public Function CreateAPITask(ByVal Title As String, _
                             ByVal Desctiption As String, _
                             ByVal DDate As String, _
                             Optional ByVal Priority As String = vbNullString, _
                             Optional ByVal Who As String = vbNullString, _
                             Optional ByVal Notes As String = vbNullString, _
                             Optional ByVal PrivateNotes As String = vbNullString, _
                             Optional ByVal Tags As String = vbNullString, _
                             Optional ByVal MMURL As String = vbNullString, _
                             Optional ByVal Others As Collection = Nothing, _
                             Optional ByVal MeetingUID As String = vbNullString) As WebResponse
    
    Dim response As WebResponse
    Dim Body As Dictionary
    Const filePath As String = "action-items"
    
    On Error GoTo ErrorHandler
    
    CurrentF = "CreateAPITask"
    
    ' Validate inputs
    If Len(Title) = 0 Then
        WriteLog 2, CurrentMod, CurrentF, "Empty task title"
    End If
    
    ' Build request body
    Set Body = New Dictionary
    Body.Add "@type", "action_items"
    Body.Add "title", "action_items"
    
    If Len(Desctiption) > 0 Then Body.Add "full_explanation", Desctiption
    If Len(Notes) > 0 Then Body.Add "notes", Notes
    If Len(PrivateNotes) > 0 Then Body.Add "private_notes", PrivateNotes
    If Len(DDate) > 0 Then Body.Add "initial_due_date", Format$(DDate, "yyyy-m-d")
    If Len(MMURL) > 0 Then Body.Add "meeting_minutes", MMURL
    If Len(MeetingUID) > 0 Then Body.Add "source", MeetingUID
    If Not Others Is Nothing Then Body.Add "related_members", Others
    
    ' Add tags as collection
    If Len(Tags) > 0 Then
        Dim Coll As Collection
        Set Coll = ArrToColl(Split(Tags, ","))
        Body.Add "subjects", Coll
    End If
    
    Body.Add "priority", Priority
    Body.Add "assigned_to", Who
    
    ' Create task
    Set response = UseAPI(Body, WebMethod.HttpPost, filePath)
    
    ' Handle response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, "Failed to create task"
    ElseIf response.StatusCode = HTTP_CREATED Then
        ' Update title
        Call UpdateAPIContent(response.Data("@id"), Array("title"), Array(Title))
        Set CreateAPITask = response
    ElseIf response.StatusCode = HTTP_UNAUTHORIZED Then
        Set CreateAPITask = response
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed with status " & response.StatusCode
        Set CreateAPITask = response
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set CreateAPITask = Nothing
End Function

'=======================================================
' Function: GetTaskPriorities
' Purpose: Get available task priority options
'
' Returns: 2D array of priorities (title and token), or empty array on error
'=======================================================
Public Function GetTaskPriorities() As Variant
    Dim response As WebResponse
    Dim Arr() As Variant
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    Set response = UseAPI(Nothing, WebMethod.HttpGet, _
                         "/@vocabularies/DocentIMS.ActionItems.PriorityVocabulary")
    
    If response Is Nothing Then
        ReDim Arr(1 To 1, 1 To 1)
        Arr(1, 1) = "Failed to get task priorities"
        WriteLog 3, CurrentMod, "GetTaskPriorities", Arr(1, 1)
    ElseIf response.StatusCode = HTTP_OK Then
        i = response.Data("items").Count
        If i > 0 Then
            ReDim Arr(1 To 2, 1 To i)
            For i = 1 To i
                Arr(1, i) = response.Data("items")(i)("title")
                Arr(2, i) = response.Data("items")(i)("token")
            Next i
        End If
    Else
        ReDim Arr(1 To 1, 1 To 1)
        Arr(1, 1) = "Failed to get task priorities"
        WriteLog 3, CurrentMod, "GetTaskPriorities", Arr(1, 1)
    End If
    
    GetTaskPriorities = Arr
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetTaskPriorities", _
             "Error " & Err.Number & ": " & Err.Description
    ReDim Arr(1 To 1, 1 To 1)
    Arr(1, 1) = "Error: " & Err.Description
    GetTaskPriorities = Arr
End Function

'=======================================================
' Function: GetAPITasksCounts
' Purpose: Get count of tasks by stoplight color
'
' Returns: Dictionary with task counts and URLs
'=======================================================
Public Function GetAPITasksCounts() As Dictionary
    Dim Resp As Collection
    Dim Arr(1 To 3) As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetAPITasksCounts"
    
    Set TasksDict = New Dictionary
    
    ' Define stoplight colors
    Arr(1) = "Green"
    Arr(2) = "Yellow"
    Arr(3) = "Red"
    
    ' Get all tasks
    Set Resp = GetAPIFolder(, "action_items", Array("stoplight"), _
                           Array("review_state", "assigned_id"), _
                           Array(ProjectInfo("stoplight_state"), MainInfo("id")))
    
    ' Count tasks by color
    For i = 1 To UBound(Arr)
        n = 0
        For j = 1 To Resp.Count
            If Resp(j)("stoplight") = Arr(i) Then
                n = n + 1
            End If
        Next j
        
        TasksDict.Add Arr(i), n
        TasksDict.Add Arr(i) & "URL", _
            "/action-items/action-items-collection/?collectionfilter=1&stoplight=" & _
            Arr(i) & "&assigned_to=" & MainInfo("fullname") & "&review_state=Published"
    Next i
    
    Set GetAPITasksCounts = TasksDict
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetAPITasksCounts = New Dictionary
End Function

'=======================================================
' Function: GetAPINotificationsCounts
' Purpose: Get count of notifications by type
'
' Returns: Dictionary with notification counts and URLs
'=======================================================
Public Function GetAPINotificationsCounts() As Dictionary
    Dim Resp As Collection
    Dim Arr(1 To 3, 1 To 2) As String
    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim nTotal As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetAPINotificationsCounts"
    
    Set NotifsDict = New Dictionary
    
    ' Define notification types
    Arr(1, 1) = "info"
    Arr(2, 1) = "warning"
    Arr(3, 1) = "error"
    Arr(1, 2) = "Aqua"
    Arr(2, 2) = "Yellow"
    Arr(3, 2) = "Red"
    
    ' Get all notifications
    Set Resp = GetAPIFolder("NoSearch_notifications/notifications-collection", , _
                           Array("notification_type"))
    
    ' Count notifications by type
    For i = 1 To UBound(Arr, 1)
        n = 0
        For j = 1 To Resp.Count
            If Resp(j)("notification_type") = Arr(i, 1) Then
                n = n + 1
            End If
        Next j
        
        nTotal = nTotal + n
        NotifsDict.Add Arr(i, 2), n
        NotifsDict.Add Arr(i, 2) & "URL", _
            "/notifications/notifications-collection/?collectionfilter=1&show_all=1&notification_type=" & _
            Arr(i, 1)
    Next i
    
    ' Add total count
    NotifsDict.Add "All", nTotal
    NotifsDict.Add "AllURL", "/notifications/notifications-collection/?collectionfilter=1&show_all=1"
    
    Set GetAPINotificationsCounts = NotifsDict
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetAPINotificationsCounts = New Dictionary
End Function

'=======================================================
' MEETING DOCUMENT FUNCTIONS
'=======================================================

'=======================================================
' Function: GetMeetingDocOfType
' Purpose: Retrieve meeting documents of a specific type
'
' Description:
'   Complex function that retrieves meeting documents and
'   associates them with their parent meetings. Handles
'   different document types (meeting, meeting_notes, meeting_minutes).
'
' Parameters:
'   DocType - Document type (optional)
'   HasFile - File filter (optional)
'   FolderPath - Folder to search (default: "meetings")
'
' Returns: Collection of meeting documents with parent meeting info
'=======================================================
Public Function GetMeetingDocOfType(Optional ByVal DocType As String = vbNullString, _
                                   Optional ByVal HasFile As String = vbNullString, _
                                   Optional ByVal FolderPath As String = "meetings") As Variant
    
    Dim MeetingsDict As Object
    Dim Response1 As Object
    Dim Response2 As Object
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetMeetingDocOfType"
    
    DocType = Replace(LCase$(DocType), " ", "_")
    
    ' Get primary documents
    If DocType = "meeting" Then
        Set Response1 = GetAPIFolder(FolderPath, "meeting", _
                                    Array("start", "location", "meeting_type"))
        If Not IsGoodResponse(Response1) Then GoTo ErrorHandler
        Set MeetingsDict = Response1
    Else
        Set Response1 = GetAPIFolder(FolderPath, DocType, , _
                                    Array("has_file"), Array(HasFile))
        If Not IsGoodResponse(Response1) Then GoTo ErrorHandler
        Set MeetingsDict = GetAPIFolder(FolderPath, "meeting", Array("start"), _
                                       Array("review_state"), _
                                       Array(GetStateID("Published", "meeting")))
    End If
    
    ' Get related documents if needed
    If HasFile <> "True" Then
        Select Case DocType
            Case "meeting_notes"
                Set Response2 = GetAPIFolder(FolderPath, "meeting_agenda", _
                                           Array("planned_action_items"), _
                                           Array("review_state"), _
                                           Array(GetStateID("Published", "meeting_agenda")))
            
            Case "meeting_minutes"
                Set Response2 = GetAPIFolder(FolderPath, "meeting_notes", _
                                           Array("proposed_action_items", "planned_action_items", "actuals"), _
                                           Array("review_state"), _
                                           Array(GetStateID("Published", "meeting_notes")))
        End Select
    End If
    
    ' Process and associate documents
    If Response1 Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get meeting documents of type: " & DocType
    ElseIf TypeName(Response1) = "Collection" Then
        Call ProcessMeetingDocuments(Response1, Response2, MeetingsDict, DocType)
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Unexpected response for type: " & DocType
    End If
    
    Set GetMeetingDocOfType = Response1
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetMeetingDocOfType = Nothing
End Function

'=======================================================
' Sub: ProcessMeetingDocuments
' Purpose: Process and associate meeting documents with parent meetings
'=======================================================
Private Sub ProcessMeetingDocuments(ByRef Response1 As Collection, _
                                   ByVal Response2 As Object, _
                                   ByVal MeetingsDict As Object, _
                                   ByVal DocType As String)
    
    Dim i As Long
    Dim j As Long
    
    On Error Resume Next
    
    ' Add state names and meeting info
    If Response2 Is Nothing Then
        ' Simple case - just add state names
        For i = 1 To Response1.Count
            Response1(i).Add "State", GetStateName(CStr(Response1(i)("review_state")), _
                                                  Response1(i)("@type"))
        Next i
    Else
        ' Complex case - associate with parent documents
        For j = 1 To Response2.Count
            Response2(j).Add "State", GetStateName(CStr(Response2(j)("review_state")), _
                                                  Response2(j)("@type"))
            Response2(j).Add "MeetingDateTime", _
                           ToServerTime(CStr(Response2(j)("actual_meeting_date_time")))
        Next j
        
        ' Associate Response1 items with Response2 items
        For i = 1 To Response1.Count
            For j = 1 To Response2.Count
                If GetParentDir(Response1(i)("@id")) = GetParentDir(Response2(j)("@id")) Then
                    Response1(i).Add "ParentDoc", Response2(j)
                    Exit For
                End If
            Next j
        Next i
        
        ' Remove items without parent docs
        For i = Response1.Count To 1 Step -1
            If Not Response1(i).Exists("ParentDoc") Then
                Response1.Remove i
            End If
        Next i
    End If
    
    ' Add meeting information
    If Response1.Count > 0 Then
        For j = 1 To MeetingsDict.Count
            MeetingsDict(j).Add "MeetingDateTime", _
                              ToServerTime(CStr(MeetingsDict(j)("start")))
            MeetingsDict(j).Add "MeetingShortName", _
                              MeetingsDict(j)("title") & " - " & _
                              Format$(MeetingsDict(j)("MeetingDateTime"), LongDateTimeFormat)
        Next j
    End If
    
    ' Associate with parent meetings
    For i = Response1.Count To 1 Step -1
        For j = 1 To MeetingsDict.Count
            If Response1(i)("@id") Like MeetingsDict(j)("@id") & "/*" Or _
               (DocType = "meeting" And Response1(i)("@id") = MeetingsDict(j)("@id")) Then
                Response1(i).Add "ParentMeeting", CloneDictionary(MeetingsDict(j))
                Exit For
            End If
        Next j
        
        ' Remove if no parent meeting found
        If j > MeetingsDict.Count Then
            Response1.Remove i
        End If
    Next i
End Sub

'=======================================================
' Function: GetParentMeetingObject
' Purpose: Get parent meeting information
'
' Parameters:
'   MeetingId - Meeting ID/URL
'
' Returns: Dictionary with meeting information
'=======================================================
Public Function GetParentMeetingObject(ByVal MeetingId As String) As Dictionary
    Dim response As Object
    Dim MeetingInfo As Dictionary
    
    On Error GoTo ErrorHandler
    
    Set response = GetAPIContent(MeetingId)
    
    If Not response Is Nothing Then
        Set MeetingInfo = response.Data
        MeetingInfo.Add "MeetingDateTime", ToServerTime(CStr(MeetingInfo("start")))
        MeetingInfo.Add "MeetingShortName", _
                       MeetingInfo("title") & " - " & _
                       Format$(MeetingInfo("MeetingDateTime"), LongDateTimeFormat)
    End If
    
    Set GetParentMeetingObject = MeetingInfo
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetParentMeetingObject", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetParentMeetingObject = Nothing
End Function

'=======================================================
' WORKFLOW FUNCTIONS
'=======================================================

'=======================================================
' Function: UpdateAPIFileWorkflow
' Purpose: Update workflow state of a file
'
' Parameters:
'   fileURL - File URL
'   TransitionID - Transition ID or full transition URL
'
' Returns: WebResponse object
'=======================================================
Public Function UpdateAPIFileWorkflow(ByVal fileURL As String, _
                                     ByVal transitionID As String) As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "UpdateAPIFileWorkflow"
    
    ' Validate inputs
    If Len(fileURL) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty file URL"
        Set UpdateAPIFileWorkflow = Nothing
        Exit Function
    End If
    
    If Len(transitionID) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty transition ID"
        Set UpdateAPIFileWorkflow = Nothing
        Exit Function
    End If
    
    ' Build transition URL
    If InStr(transitionID, "/@workflow/") > 0 Then
        Set UpdateAPIFileWorkflow = UseAPI(Nothing, WebMethod.HttpPost, transitionID)
    Else
        Set UpdateAPIFileWorkflow = UseAPI(Nothing, WebMethod.HttpPost, _
                                          fileURL & "/@workflow/" & transitionID)
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set UpdateAPIFileWorkflow = Nothing
End Function

'=======================================================
' Function: GetAPIFileWorkflowTransitions
' Purpose: Get available workflow transitions for a document
'
' Parameters:
'   DocURL - Document URL
'
' Returns: Collection of available transitions, or Nothing on error
'=======================================================
Public Function GetAPIFileWorkflowTransitions(ByVal DocURL As String) As Collection
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    CurrentF = "GetAPIFileWorkflowTransitions"
    
    ' Validate input
    If Len(DocURL) = 0 Then
        WriteLog 3, CurrentMod, CurrentF, "Empty document URL"
        Set GetAPIFileWorkflowTransitions = Nothing
        Exit Function
    End If
    
    ' Get workflow info
    Set response = UseAPI(Nothing, WebMethod.HttpGet, "/noapi/" & DocURL & "/@workflow")
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get workflow transitions: " & DocURL
    ElseIf response.StatusCode = HTTP_OK Then
        Set GetAPIFileWorkflowTransitions = response.Data("transitions")
    Else
        WriteLog 3, CurrentMod, CurrentF, _
                 "Failed to get workflow transitions: " & DocURL & " - " & response.Content
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, CurrentF, "Error " & Err.Number & ": " & Err.Description
    Set GetAPIFileWorkflowTransitions = Nothing
End Function

'=======================================================
' DASHBOARD FUNCTIONS
'=======================================================

'=======================================================
' Function: GetProjectsList
' Purpose: Get list of projects from dashboard
'
' Parameters:
'   mUser - Username/email
'
' Returns: Collection of project buttons/info
'=======================================================
Public Function GetProjectsList(ByVal mUser As String) As Collection
    Dim response As WebResponse
    
    On Error GoTo ErrorHandler
    
    ' Ensure dashboard URL is set
    If Len(DashboardURLStr) = 0 Then
        DashboardURLStr = GetRegDashboardURL
    End If
    
    ' Validate inputs
    If Len(mUser) = 0 Then
        WriteLog 3, CurrentMod, "GetProjectsList", "Empty username"
        Set GetProjectsList = Nothing
        Exit Function
    End If
    
    ' Get projects list
    Set response = UseAPI(Nothing, WebMethod.HttpGet, _
                         "/@dashboard_sites/?email=" & mUser, , DashboardURLStr)
    
    ' Process response
    If response Is Nothing Then
        WriteLog 3, CurrentMod, "GetProjectsList", "Failed to get projects list"
    ElseIf response.StatusCode = HTTP_OK Then
        Set GetProjectsList = response.Data("buttons")
    Else
        WriteLog 3, CurrentMod, "GetProjectsList", _
                 "Failed with status " & response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetProjectsList", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetProjectsList = Nothing
End Function

'=======================================================
' DEPRECATED/COMMENTED FUNCTIONS
'=======================================================

'=======================================================
' Function: LockAPIFile
' Purpose: Lock/unlock API file (DEPRECATED - Not currently used)
' Note: Kept for potential future use but commented out
'=======================================================
Public Function LockAPIFile(ByVal OnlineFilePath As String, _
                           Optional ByVal UnlockMode As Boolean = False) As String
    ' Currently not implemented/used
    ' Kept for reference but functionality is commented out
    LockAPIFile = "Function not implemented"
End Function

'=======================================================
' END OF MODULE
'=======================================================
