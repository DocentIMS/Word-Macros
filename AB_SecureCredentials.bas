Attribute VB_Name = "AB_SecureCredentials"
Option Explicit

'=======================================================
' Module: AB_SecureCredentials
' Purpose: Secure credential management for API access
' Author: Refactored from AC_API_Mod
' Date: November 2025
'
' Description:
'   This module provides secure storage and retrieval of
'   credentials for API access. It removes hardcoded
'   passwords and implements secure storage mechanisms.
'
' Security Features:
'   - No hardcoded passwords
'   - Registry storage with basic encryption
'   - Fallback to user input if credentials not found
'   - Credential validation before use
'
' Usage:
'   Dim creds As ApiCredentials
'   creds = GetDashboardCredentials()
'   If creds.IsValid Then
'       ' Use creds.UserName and creds.Password
'   End If
'=======================================================

Private Const CurrentMod As String = "SecureCredentials"

' Registry paths for credential storage
Private Const CRED_REG_PATH As String = "HKEY_CURRENT_USER\Software\DocentIMS\Credentials\"
Private Const DASHBOARD_USER_REG As String = "DashboardUser"
Private Const DASHBOARD_PWD_REG As String = "DashboardPassword"

' Encryption key (simple XOR - for production, use stronger encryption)
Private Const ENCRYPTION_KEY As String = "DocentIMS_2025_SecureKey_v1"

'=======================================================
' Type: ApiCredentials
' Purpose: Hold API credential information
'=======================================================
Public Type ApiCredentials
    UserName As String
    Password As String
    IsValid As Boolean
End Type

'=======================================================
' Function: GetDashboardCredentials
' Purpose: Retrieve dashboard API credentials securely
'
' Returns: ApiCredentials type with user/password
'          If credentials not found, prompts user to enter them
'
' Example:
'   Dim creds As ApiCredentials
'   creds = GetDashboardCredentials()
'   If creds.IsValid Then
'       ' Use credentials
'   End If
'=======================================================
Public Function GetDashboardCredentials() As ApiCredentials
    Dim creds As ApiCredentials
    Dim storedUser As String
    Dim storedPwd As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetDashboardCredentials", "Retrieving dashboard credentials"
    
    ' Try to get credentials from registry
    storedUser = GetSecureValue(DASHBOARD_USER_REG)
    storedPwd = GetSecureValue(DASHBOARD_PWD_REG)
    
    ' If credentials exist, decrypt and return
    If Len(storedUser) > 0 And Len(storedPwd) > 0 Then
        creds.UserName = storedUser
        creds.Password = DecryptPassword(storedPwd)
        creds.IsValid = True
        
        WriteLog 1, CurrentMod, "GetDashboardCredentials", "Credentials retrieved from secure storage"
    Else
        ' Credentials not found - prompt user
        WriteLog 2, CurrentMod, "GetDashboardCredentials", "Credentials not found in storage"
        creds = PromptForDashboardCredentials()
    End If
    
    GetDashboardCredentials = creds
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetDashboardCredentials", _
             "Error " & Err.Number & ": " & Err.Description
    creds.IsValid = False
    GetDashboardCredentials = creds
End Function

'=======================================================
' Function: GetProjectCredentials
' Purpose: Retrieve project-specific credentials
'
' Parameters:
'   projectName - Name of the project
'
' Returns: ApiCredentials type
'=======================================================
Public Function GetProjectCredentials(ByVal projectName As String) As ApiCredentials
    Dim creds As ApiCredentials
    
    On Error GoTo ErrorHandler
    
    ' For project credentials, use the global variables
    ' This maintains backward compatibility
    creds.UserName = UserNameStr
    creds.Password = UserPasswordStr
    creds.IsValid = (Len(creds.UserName) > 0 And Len(creds.Password) > 0)
    
    If Not creds.IsValid Then
        WriteLog 2, CurrentMod, "GetProjectCredentials", _
                 "Project credentials not set for: " & projectName
    End If
    
    GetProjectCredentials = creds
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetProjectCredentials", _
             "Error " & Err.Number & ": " & Err.Description
    creds.IsValid = False
    GetProjectCredentials = creds
End Function

'=======================================================
' Function: StoreDashboardCredentials
' Purpose: Store dashboard credentials securely
'
' Parameters:
'   userName - Dashboard username
'   password - Dashboard password
'
' Returns: True if stored successfully
'=======================================================
Public Function StoreDashboardCredentials(ByVal UserName As String, _
                                         ByVal Password As String) As Boolean
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "StoreDashboardCredentials", "Storing dashboard credentials"
    
    ' Validate inputs
    If Len(UserName) = 0 Or Len(Password) = 0 Then
        WriteLog 2, CurrentMod, "StoreDashboardCredentials", "Invalid credentials provided"
        StoreDashboardCredentials = False
        Exit Function
    End If
    
    ' Encrypt password before storing
    Dim encryptedPwd As String
    encryptedPwd = EncryptPassword(Password)
    
    ' Store in registry
    Call SetSecureValue(DASHBOARD_USER_REG, UserName)
    Call SetSecureValue(DASHBOARD_PWD_REG, encryptedPwd)
    
    WriteLog 1, CurrentMod, "StoreDashboardCredentials", "Credentials stored successfully"
    StoreDashboardCredentials = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "StoreDashboardCredentials", _
             "Error " & Err.Number & ": " & Err.Description
    StoreDashboardCredentials = False
End Function

'=======================================================
' Function: ClearDashboardCredentials
' Purpose: Remove stored dashboard credentials
'
' Returns: True if cleared successfully
'=======================================================
Public Function ClearDashboardCredentials() As Boolean
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "ClearDashboardCredentials", "Clearing dashboard credentials"
    
    Call DeleteSecureValue(DASHBOARD_USER_REG)
    Call DeleteSecureValue(DASHBOARD_PWD_REG)
    
    WriteLog 1, CurrentMod, "ClearDashboardCredentials", "Credentials cleared successfully"
    ClearDashboardCredentials = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ClearDashboardCredentials", _
             "Error " & Err.Number & ": " & Err.Description
    ClearDashboardCredentials = False
End Function

'=======================================================
' PRIVATE HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: PromptForDashboardCredentials
' Purpose: Prompt user to enter dashboard credentials
'=======================================================
Private Function PromptForDashboardCredentials() As ApiCredentials
    Dim creds As ApiCredentials
    Dim UserName As String
    Dim Password As String
    Dim saveChoice As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    ' Prompt for username
    UserName = InputBox( _
        "Enter Dashboard Username:" & vbCrLf & vbCrLf & _
        "This is required for dashboard API access.", _
        "Dashboard Credentials Required", _
        "vbauser@docentims.com")
    
    If Len(UserName) = 0 Then
        ' User cancelled
        creds.IsValid = False
        PromptForDashboardCredentials = creds
        Exit Function
    End If
    
    ' Prompt for password
    Password = frmInputBox.Display( _
        "Enter Dashboard Password for:" & vbCrLf & _
        UserName & vbCrLf & vbCrLf & _
        "Your password will be stored securely.", _
        "Dashboard Credentials Required", , True)
    
    If Len(Password) = 0 Then
        ' User cancelled
        creds.IsValid = False
        PromptForDashboardCredentials = creds
        Exit Function
    End If
    
    ' Set credentials
    creds.UserName = UserName
    creds.Password = Password
    creds.IsValid = True
    
    ' Ask if user wants to save credentials
    saveChoice = MsgBox( _
        "Would you like to save these credentials for future use?" & vbCrLf & vbCrLf & _
        "They will be stored securely in the Windows Registry.", _
        vbQuestion + vbYesNo, _
        "Save Credentials?")
    
    If saveChoice = vbYes Then
        If StoreDashboardCredentials(UserName, Password) Then
            MsgBox "Credentials saved successfully.", vbInformation, "Success"
        Else
            MsgBox "Warning: Could not save credentials. You may be prompted again.", _
                   vbExclamation, "Warning"
        End If
    End If
    
    PromptForDashboardCredentials = creds
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "PromptForDashboardCredentials", _
             "Error " & Err.Number & ": " & Err.Description
    creds.IsValid = False
    PromptForDashboardCredentials = creds
End Function

'=======================================================
' Function: GetSecureValue
' Purpose: Retrieve value from registry
'=======================================================
Private Function GetSecureValue(ByVal ValueName As String) As String
    On Error Resume Next
    GetSecureValue = GetReg(ValueName, CRED_REG_PATH)
    If Err.Number <> 0 Then
        GetSecureValue = vbNullString
    End If
End Function

'=======================================================
' Function: SetSecureValue
' Purpose: Store value in registry
'=======================================================
Private Sub SetSecureValue(ByVal ValueName As String, ByVal value As String)
    On Error Resume Next
    Call SetReg(ValueName, value, CRED_REG_PATH, REG_SZ)
End Sub

'=======================================================
' Function: DeleteSecureValue
' Purpose: Remove value from registry
'=======================================================
Private Sub DeleteSecureValue(ByVal ValueName As String)
    On Error Resume Next
    Call DelKey(CRED_REG_PATH, ValueName)
End Sub

'=======================================================
' Function: ValidateCredentials
' Purpose: Validate credential format (basic check)
'
' Parameters:
'   creds - Credentials to validate
'
' Returns: True if credentials appear valid
'=======================================================
Public Function ValidateCredentials(ByRef creds As ApiCredentials) As Boolean
    ' Basic validation
    If Len(creds.UserName) = 0 Then
        ValidateCredentials = False
        Exit Function
    End If
    
    If Len(creds.Password) = 0 Then
        ValidateCredentials = False
        Exit Function
    End If
    
    ' Check for valid email format for username
    If InStr(creds.UserName, "@") = 0 Then
        WriteLog 2, CurrentMod, "ValidateCredentials", _
                 "Username does not appear to be a valid email"
    End If
    
    ValidateCredentials = True
End Function

'=======================================================
' ENCRYPTION FUNCTIONS
'=======================================================

'=======================================================
' Function: EncryptPassword
' Purpose: Encrypt password using simple XOR cipher
'
' Parameters:
'   plainText - Password to encrypt
'
' Returns: Encrypted password as Base64 string
'
' Note: This is a basic XOR encryption. For production use,
'       consider implementing stronger encryption methods
'       such as AES or RSA.
'=======================================================
Function EncryptPassword(ByVal plainText As String) As String
    Dim i As Long
    Dim keyLen As Long
    Dim result As String
    Dim charCode As Integer
    Dim keyChar As Integer
    
    On Error GoTo ErrorHandler
    
    If Len(plainText) = 0 Then
        EncryptPassword = vbNullString
        Exit Function
    End If
    
    keyLen = Len(ENCRYPTION_KEY)
    result = vbNullString
    
    ' XOR each character with corresponding key character
    For i = 1 To Len(plainText)
        charCode = Asc(Mid$(plainText, i, 1))
        keyChar = Asc(Mid$(ENCRYPTION_KEY, ((i - 1) Mod keyLen) + 1, 1))
        result = result & Chr$(charCode Xor keyChar)
    Next i
    
    ' Convert to Base64 for safe storage
    EncryptPassword = EncodeBase64(result)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "EncryptPassword", _
             "Error " & Err.Number & ": " & Err.Description
    EncryptPassword = vbNullString
End Function

'=======================================================
' Function: DecryptPassword
' Purpose: Decrypt password using simple XOR cipher
'
' Parameters:
'   encryptedText - Encrypted password (Base64 encoded)
'
' Returns: Decrypted password
'=======================================================
Function DecryptPassword(ByVal encryptedText As String) As String
    Dim i As Long
    Dim keyLen As Long
    Dim result As String
    Dim charCode As Integer
    Dim keyChar As Integer
    Dim decodedText As String
    
    On Error GoTo ErrorHandler
    
    If Len(encryptedText) = 0 Then
        DecryptPassword = vbNullString
        Exit Function
    End If
    
    ' Decode from Base64
    decodedText = DecodeBase64(encryptedText)
    
    keyLen = Len(ENCRYPTION_KEY)
    result = vbNullString
    
    ' XOR each character with corresponding key character (same as encrypt)
    For i = 1 To Len(decodedText)
        charCode = Asc(Mid$(decodedText, i, 1))
        keyChar = Asc(Mid$(ENCRYPTION_KEY, ((i - 1) Mod keyLen) + 1, 1))
        result = result & Chr$(charCode Xor keyChar)
    Next i
    
    DecryptPassword = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "DecryptPassword", _
             "Error " & Err.Number & ": " & Err.Description
    DecryptPassword = vbNullString
End Function

'=======================================================
' Function: EncodeBase64
' Purpose: Encode string to Base64
'
' Parameters:
'   text - Text to encode
'
' Returns: Base64 encoded string
'=======================================================
Private Function EncodeBase64(ByVal text As String) As String
    Dim objXML As Object
    Dim objNode As Object
    Dim byteData() As Byte
    
    On Error GoTo ErrorHandler
    
    ' Use MSXML for Base64 encoding
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    
    ' Convert string to byte array
    byteData = StringToByteArray(text)
    objNode.nodeTypedValue = byteData
    
    EncodeBase64 = objNode.text
    
    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "EncodeBase64", _
             "Error " & Err.Number & ": " & Err.Description
    EncodeBase64 = vbNullString
End Function

'=======================================================
' Function: DecodeBase64
' Purpose: Decode Base64 string
'
' Parameters:
'   base64Text - Base64 encoded text
'
' Returns: Decoded string
'=======================================================
Private Function DecodeBase64(ByVal base64Text As String) As String
    Dim objXML As Object
    Dim objNode As Object
    Dim byteData As Variant
    
    On Error GoTo ErrorHandler
    
    ' Use MSXML for Base64 decoding
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    
    objNode.DataType = "bin.base64"
    objNode.text = base64Text
    
    ' Get byte data as Variant first
    byteData = objNode.nodeTypedValue
    
    ' Convert to string
    DecodeBase64 = ByteArrayToString(byteData)
    
    Set objNode = Nothing
    Set objXML = Nothing
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "DecodeBase64", _
             "Error " & Err.Number & ": " & Err.Description
    DecodeBase64 = vbNullString
End Function

'=======================================================
' Function: StringToByteArray
' Purpose: Convert string to byte array
'
' Parameters:
'   text - String to convert
'
' Returns: Byte array
'=======================================================
Private Function StringToByteArray(ByVal text As String) As Byte()
    Dim i As Long
    Dim byteArray() As Byte
    
    On Error GoTo ErrorHandler
    
    If Len(text) = 0 Then
        ReDim byteArray(0 To 0)
        StringToByteArray = byteArray
        Exit Function
    End If
    
    ReDim byteArray(0 To Len(text) - 1)
    
    For i = 1 To Len(text)
        byteArray(i - 1) = Asc(Mid$(text, i, 1))
    Next i
    
    StringToByteArray = byteArray
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "StringToByteArray", _
             "Error " & Err.Number & ": " & Err.Description
    ReDim byteArray(0 To 0)
    StringToByteArray = byteArray
End Function

'=======================================================
' Function: ByteArrayToString
' Purpose: Convert byte array to string
'
' Parameters:
'   byteArray - Byte array to convert (as Variant to handle MSXML output)
'
' Returns: String
'=======================================================
Private Function ByteArrayToString(ByVal byteArray As Variant) As String
    Dim i As Long
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    result = vbNullString
    
    ' Handle case where byteArray is not an array
    If Not IsArray(byteArray) Then
        ByteArrayToString = vbNullString
        Exit Function
    End If
    
    For i = LBound(byteArray) To UBound(byteArray)
        result = result & Chr$(byteArray(i))
    Next i
    
    ByteArrayToString = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ByteArrayToString", _
             "Error " & Err.Number & ": " & Err.Description
    ByteArrayToString = vbNullString
End Function

'=======================================================
' END OF MODULE
'=======================================================

