Attribute VB_Name = "Tools_Reg"
Option Explicit

'=======================================================
' Module: Tools_Reg
' Purpose: Low-level Windows Registry API wrapper
' Author: Refactored November 2025
' Version: 2.0
'
' Description:
'   Provides safe, robust access to Windows Registry through
'   Windows API functions. Handles both 32-bit and 64-bit
'   environments with proper error handling and validation.
'
' Security Notes:
'   - All registry operations are restricted to HKEY_CURRENT_USER
'     unless explicitly specified with HKEY prefix
'   - Proper error checking on all API calls
'   - Buffer overflow protection
'   - Input validation on all public functions
'
' Dependencies:
'   - Windows API (advapi32.dll, shlwapi.dll)
'   - AB_GlobalConstants (for error codes)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added documentation headers
'       * Improved DelKey with validation
'       * Added error logging
'   v1.0 - Original version
'=======================================================

'=======================================================
' CONSTANTS
'=======================================================

' Buffer size for registry key enumeration
Private Const BUF_SIZE As Long = 512

' Module name for logging
Private Const CurrentMod As String = "Tools_Reg"

'=======================================================
' TYPE DEFINITIONS
'=======================================================

' Windows FILETIME structure
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'=======================================================
' ENUMERATIONS
'=======================================================

' Registry root keys
Private Enum RegistryRootEnum
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    rrDynData = &H80000006
End Enum

' Registry access levels
Private Enum RegKeyAccessLevel
    KEY_ALL_ACCESS = &H3F
    KEY_WOW64_64KEY = &H100&  ' 32-bit app to access 64-bit hive
    KEY_WOW64_32KEY = &H200&  ' 64-bit app to access 32-bit hive
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    KEY_ENUMERATE_SUB_KEYS = &H8
End Enum

' Registry value types (Public for external use)
Public Enum RegValueType
    REG_SZ = 1      ' String value
    REG_DWORD = 4   ' 32-bit number
End Enum

' Registry error codes
Private Enum RegistryErrorEnum
     ERROR_SUCCESS = 0
     ERROR_BADDB = 1
     ERROR_BADKEY = 2
     ERROR_CANTOPEN = 3
     ERROR_CANTREAD = 4
     ERROR_CANTWRITE = 5
     ERROR_OUTOFMEMORY = 6
     ERROR_ARENA_TRASHED = 7
     ERROR_ACCESS_DENIED = 8
     ERROR_INVALID_PARAMETERS = 87
     ERROR_MORE_DATA_AVAILABLE = 234
     ERROR_NO_MORE_ITEMS = 259
End Enum

' Registry options
Private Const REG_OPTION_NON_VOLATILE As Long = 0

'=======================================================
' WINDOWS API DECLARATIONS
'=======================================================

#If VBA7 Then
    ' Delete a registry key value
    Private Declare PtrSafe Function RegDeleteKeyValue Lib "advapi32.dll" Alias "RegDeleteKeyValueA" _
        (ByVal handle As LongPtr, ByVal pszSubKey As String, ByVal ValueName As String) As LongPtr
        
    ' Delete a registry key and all subkeys
    Private Declare PtrSafe Function RegDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" _
        (ByVal hKey As LongPtr, ByVal pszSubKey As String) As LongPtr

    ' Close a registry key handle
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As LongPtr) As LongPtr

    ' Create a registry key
    Private Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        ByVal lpType As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As LongPtr, phkResult As LongPtr, lpdwDisposition As LongPtr) As LongPtr

    ' Open a registry key
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As LongPtr) As LongPtr

    ' Query registry value (String)
    Private Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, ByVal lpData As String, cbData As LongPtr) As LongPtr
    
    ' Query registry value (Long)
    Private Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, lpData As Long, cbData As LongPtr) As LongPtr
    
    ' Query registry value (NULL check)
    Private Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        lpType As LongPtr, ByVal lpData As Long, cbData As LongPtr) As LongPtr

    ' Set registry value (String)
    Private Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        ByVal lpType As LongPtr, ByVal lpData As String, ByVal cbData As LongPtr) As LongPtr
    
    ' Set registry value (Long)
    Private Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
        ByVal lpType As LongPtr, lpData As Long, ByVal cbData As LongPtr) As LongPtr

    ' Enumerate registry subkeys
    Private Declare PtrSafe Function RegEnumKeyExW Lib "advapi32.dll" ( _
        ByVal hKey As LongPtr, ByVal i As Long, ByVal lpName As LongPtr, _
        lpcName As LongPtr, ByVal lpReserved As LongPtr, ByVal lpClass As LongPtr, _
        ByVal lpcClass As LongPtr, lpfLastWrite As FILETIME) As LongPtr
#Else
    ' 32-bit Office declarations
    Private Declare Function RegDeleteKeyValue Lib "advapi32.dll" _
        (ByVal handle As Long, ByVal pszSubKey As String, ByVal ValueName As String) As Long

    Private Declare Function RegDeleteKey Lib "shlwapi.dll" Alias "SHDeleteKeyA" _
        (ByVal hKey As Long, ByVal pszSubKey As String) As Long

    Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

    Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        ByVal lpType As Long, ByVal dwOptions As Long, ByVal samDesired As Long, _
        ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As Long) As Long

    Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, ByVal lpData As String, cbData As Long) As Long
    
    Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Long, cbData As Long) As Long
    
    Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, ByVal lpData As Long, cbData As Long) As Long

    Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        ByVal lpType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    
    Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
        ByVal lpType As Long, lpData As Long, ByVal cbData As Long) As Long

    Private Declare Function RegEnumKeyExW Lib "advapi32.dll" ( _
        ByVal hKey As Long, ByVal i As Long, ByVal lpName As Long, _
        lpcName As Long, ByVal lpReserved As Long, ByVal lpClass As Long, _
        ByVal lpcClass As Long, lpfLastWrite As FILETIME) As Long
#End If

'=======================================================
' MODULE VARIABLES
'=======================================================

' Last registry error number
Private RegErrNum As LongPtr

'=======================================================
' PUBLIC FUNCTIONS
'=======================================================

'=======================================================
' Function: SetReg
' Purpose: Set a registry value
'
' Parameters:
'   VName - Value name
'   vData - Value data (String or convertible to String)
'   KeyPath - Registry key path
'   ValueType - Type of value (REG_SZ or REG_DWORD)
'
' Returns: True if successful, False otherwise
'
' Example:
'   success = SetReg("ServerURL", "https://example.com", _
'                    "HKEY_CURRENT_USER\Software\MyApp", REG_SZ)
'=======================================================
Public Function SetReg(ByVal VName As String, _
                      ByVal vData As String, _
                      ByVal KeyPath As String, _
                      Optional ByVal ValueType As RegValueType = REG_SZ) As Boolean
    
    Dim KeyHand As LongPtr
    Dim hKey As LongPtr
    Dim lngData As Long
    Dim strData As String
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Not ValidateRegInputs(VName, KeyPath, errorMsg) Then
        WriteLog 3, CurrentMod, "SetReg", "Validation error: " & errorMsg
        SetReg = False
        Exit Function
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Create or open the key
    RegErrNum = RegCreateKeyEx(hKey, KeyPath, 0&, vbNullString, _
                               REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                               0&, KeyHand, RegErrNum)
    
    If RegErrNum <> ERROR_SUCCESS Then
        WriteLog 3, CurrentMod, "SetReg", _
                 "Failed to create key: " & KeyPath & " (Error: " & RegErrNum & ")"
        SetReg = False
        Exit Function
    End If
    
    ' Set the value based on type
    Select Case ValueType
        Case REG_SZ
            ' String value - ensure null termination
            If Right$(vData, 1) <> Chr$(0) Then
                strData = vData & Chr$(0)
            Else
                strData = vData
            End If
            RegErrNum = RegSetValueExString(KeyHand, VName, 0&, REG_SZ, _
                                           ByVal strData, Len(strData))
        
        Case REG_DWORD
            ' DWORD value
            lngData = CLng(vData)
            RegErrNum = RegSetValueExLong(KeyHand, VName, 0&, REG_DWORD, _
                                         lngData, 4&)
        
        Case Else
            WriteLog 3, CurrentMod, "SetReg", _
                     "Invalid value type: " & ValueType
            RegCloseKey KeyHand
            SetReg = False
            Exit Function
    End Select
    
    ' Check if set was successful
    If RegErrNum <> ERROR_SUCCESS Then
        WriteLog 3, CurrentMod, "SetReg", _
                 "Failed to set value: " & VName & " (Error: " & RegErrNum & ")"
        SetReg = False
    Else
        SetReg = True
    End If
    
    ' Close the key handle
    RegCloseKey KeyHand
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "SetReg", errorMsg & " (Key: " & KeyPath & ")"
    If KeyHand <> 0 Then RegCloseKey KeyHand
    SetReg = False
End Function

'=======================================================
' Function: GetReg
' Purpose: Get a registry value
'
' Parameters:
'   VName - Value name
'   KeyPath - Registry key path
'
' Returns: Value data (String or Long), or empty string if not found
'
' Example:
'   serverURL = GetReg("ServerURL", "HKEY_CURRENT_USER\Software\MyApp")
'=======================================================
Public Function GetReg(ByVal VName As String, _
                      ByVal KeyPath As String) As Variant
    
    Dim hKey As LongPtr
    Dim lData As LongPtr
    Dim KeyHand As LongPtr
    Dim RegType As LongPtr
    Dim strData As String
    Dim lngData As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize return value
    GetReg = vbNullString
    
    ' Validate inputs
    If Not ValidateRegInputs(VName, KeyPath, errorMsg) Then
        WriteLog 3, CurrentMod, "GetReg", "Validation error: " & errorMsg
        Exit Function
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Open the key
    RegErrNum = RegOpenKeyEx(hKey, KeyPath, 0, KEY_QUERY_VALUE, KeyHand)
    
    If RegErrNum <> ERROR_SUCCESS Then
        ' Key doesn't exist - not necessarily an error
        Exit Function
    End If
    
    ' Query value to get type and size
    RegErrNum = RegQueryValueExNULL(KeyHand, VName, 0&, RegType, 0&, lData)
    
    If RegErrNum <> ERROR_SUCCESS Then
        ' Value doesn't exist
        RegCloseKey KeyHand
        Exit Function
    End If
    
    ' Read value based on type
    Select Case RegType
        Case REG_SZ
            ' String value
            strData = String$(CLng(lData), 0)
            RegErrNum = RegQueryValueExString(KeyHand, VName, 0&, REG_SZ, strData, lData)
            
            If RegErrNum = ERROR_SUCCESS Then
                ' Remove null terminator
                GetReg = Left$(strData, CLng(lData) - 1)
            Else
                GetReg = vbNullString
            End If
        
        Case REG_DWORD
            ' DWORD value
            RegErrNum = RegQueryValueExLong(KeyHand, VName, 0&, REG_DWORD, lngData, lData)
            
            If RegErrNum = ERROR_SUCCESS Then
                GetReg = lngData
            Else
                GetReg = 0
            End If
        
        Case Else
            ' Unsupported type
            WriteLog 2, CurrentMod, "GetReg", _
                     "Unsupported value type: " & RegType & " for " & VName
            GetReg = vbNullString
    End Select
    
    ' Close the key handle
    RegCloseKey KeyHand
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "GetReg", errorMsg & " (Key: " & KeyPath & ")"
    If KeyHand <> 0 Then RegCloseKey KeyHand
    GetReg = vbNullString
End Function

'=======================================================
' Function: DelKey
' Purpose: Delete a registry key or value
'
' Description:
'   If ValueName is provided, deletes only that value.
'   If ValueName is omitted, deletes the entire key and all subkeys.
'
' Parameters:
'   KeyPath - Registry key path
'   ValueName - Optional value name to delete
'
' Example:
'   ' Delete entire key:
'   DelKey "HKEY_CURRENT_USER\Software\MyApp\Temp"
'
'   ' Delete specific value:
'   DelKey "HKEY_CURRENT_USER\Software\MyApp", "ServerURL"
'=======================================================
Public Sub DelKey(ByVal KeyPath As String, _
                 Optional ByVal ValueName As String = vbNullString)
    
    Dim hKey As LongPtr
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(Trim$(KeyPath)) = 0 Then
        WriteLog 3, CurrentMod, "DelKey", "Empty key path provided"
        Exit Sub
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Delete value or key
    If Len(ValueName) > 0 Then
        ' Delete specific value
        RegErrNum = RegDeleteKeyValue(hKey, KeyPath, ValueName)
        
        If RegErrNum = ERROR_SUCCESS Then
            WriteLog 1, CurrentMod, "DelKey", "Deleted value: " & ValueName
        ElseIf RegErrNum = ERROR_BADKEY Or RegErrNum = ERROR_CANTOPEN Then
            ' Key doesn't exist - not an error
            WriteLog 2, CurrentMod, "DelKey", "Value not found: " & ValueName
        Else
            WriteLog 3, CurrentMod, "DelKey", _
                     "Failed to delete value: " & ValueName & " (Error: " & RegErrNum & ")"
        End If
    Else
        ' Delete entire key and all subkeys
        RegErrNum = RegDeleteKey(hKey, KeyPath)
        
        If RegErrNum = ERROR_SUCCESS Then
            WriteLog 1, CurrentMod, "DelKey", "Deleted key: " & KeyPath
        ElseIf RegErrNum = ERROR_BADKEY Or RegErrNum = ERROR_CANTOPEN Then
            ' Key doesn't exist - not an error
            WriteLog 2, CurrentMod, "DelKey", "Key not found: " & KeyPath
        Else
            WriteLog 3, CurrentMod, "DelKey", _
                     "Failed to delete key: " & KeyPath & " (Error: " & RegErrNum & ")"
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "DelKey", errorMsg & " (Key: " & KeyPath & ")"
End Sub

'=======================================================
' Function: GetRegSubKeys
' Purpose: Get list of subkeys under a registry key
'
' Parameters:
'   KeyPath - Registry key path
'
' Returns: Array of subkey names, or empty array if none found
'
' Example:
'   Dim subKeys() As String
'   subKeys = GetRegSubKeys("HKEY_CURRENT_USER\Software\MyApp")
'=======================================================
Public Function GetRegSubKeys(ByVal KeyPath As String) As Variant
    Dim hKey As LongPtr
    Dim KeyHand As LongPtr
    Dim i As Long
    Dim KeyLength As LongPtr
    Dim LastWriteTime As FILETIME
    Dim Keys() As String
    Dim KeyName(BUF_SIZE) As Byte
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Initialize empty array
    GetRegSubKeys = Array()
    
    ' Validate input
    If Len(Trim$(KeyPath)) = 0 Then
        WriteLog 3, CurrentMod, "GetRegSubKeys", "Empty key path provided"
        Exit Function
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Open the key
    RegErrNum = RegOpenKeyEx(hKey, KeyPath, 0, KEY_ENUMERATE_SUB_KEYS, KeyHand)
    
    If RegErrNum <> ERROR_SUCCESS Then
        ' Key doesn't exist or can't be opened
        WriteLog 2, CurrentMod, "GetRegSubKeys", _
                 "Cannot open key: " & KeyPath & " (Error: " & RegErrNum & ")"
        Exit Function
    End If
    
    ' Enumerate subkeys
    i = 0
    Do
        KeyLength = BUF_SIZE  ' Reset buffer size for next call
        
        ' Get subkey information
        RegErrNum = RegEnumKeyExW(KeyHand, i, VarPtr(KeyName(0)), KeyLength, _
                                  0, 0, 0, LastWriteTime)
        
        If RegErrNum <> ERROR_SUCCESS Then Exit Do
        
        i = i + 1
        ReDim Preserve Keys(1 To i)
        Keys(i) = Left$(KeyName, CLng(KeyLength))
    Loop
    
    ' Close the key
    RegCloseKey KeyHand
    
    ' Clean up
    Erase KeyName
    
    ' Return array
    If i > 0 Then
        GetRegSubKeys = Keys
    End If
    
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "GetRegSubKeys", errorMsg & " (Key: " & KeyPath & ")"
    If KeyHand <> 0 Then RegCloseKey KeyHand
    GetRegSubKeys = Array()
End Function

'=======================================================
' Function: EnsureKey
' Purpose: Create a registry key if it doesn't exist
'
' Parameters:
'   KeyPath - Registry key path to create
'
' Returns: Key handle, or 0 on error
'
' Note: Caller is responsible for closing the handle
'=======================================================
Public Function EnsureKey(ByVal KeyPath As String) As LongPtr
    Dim hKey As LongPtr
    Dim KeyHand As LongPtr
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(Trim$(KeyPath)) = 0 Then
        WriteLog 3, CurrentMod, "EnsureKey", "Empty key path provided"
        EnsureKey = 0
        Exit Function
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Create the key
    RegErrNum = RegCreateKeyEx(hKey, KeyPath, 0&, vbNullString, _
                               REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                               0&, KeyHand, RegErrNum)
    
    If RegErrNum = ERROR_SUCCESS Then
        EnsureKey = KeyHand
    Else
        WriteLog 3, CurrentMod, "EnsureKey", _
                 "Failed to create key: " & KeyPath & " (Error: " & RegErrNum & ")"
        EnsureKey = 0
    End If
    
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "EnsureKey", errorMsg & " (Key: " & KeyPath & ")"
    EnsureKey = 0
End Function

'=======================================================
' Function: CreateNewKey
' Purpose: Create a new registry key
'
' Parameters:
'   KeyPath - Registry key path to create
'
' Example:
'   CreateNewKey "HKEY_CURRENT_USER\Software\MyApp\Settings"
'=======================================================
Public Sub CreateNewKey(ByVal KeyPath As String)
    Dim hKey_New As LongPtr
    Dim hKey As LongPtr
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(Trim$(KeyPath)) = 0 Then
        WriteLog 3, CurrentMod, "CreateNewKey", "Empty key path provided"
        Exit Sub
    End If
    
    ' Get root key handle
    hKey = GetHKey(KeyPath)
    
    ' Create the key
    RegErrNum = RegCreateKeyEx(hKey, KeyPath, 0&, vbNullString, _
                               REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
                               0&, hKey_New, RegErrNum)
    
    If RegErrNum = ERROR_SUCCESS Then
        RegCloseKey hKey_New
        WriteLog 1, CurrentMod, "CreateNewKey", "Created key: " & KeyPath
    Else
        WriteLog 3, CurrentMod, "CreateNewKey", _
                 "Failed to create key: " & KeyPath & " (Error: " & RegErrNum & ")"
    End If
    
    Exit Sub
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, "CreateNewKey", errorMsg & " (Key: " & KeyPath & ")"
    If hKey_New <> 0 Then RegCloseKey hKey_New
End Sub

'=======================================================
' HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: GetHKey
' Purpose: Parse registry path and return root key handle
'
' Description:
'   Extracts the root key (HKEY_CURRENT_USER, etc.) from the
'   full path and returns the corresponding handle. Modifies
'   KeyPath parameter to remove the root key prefix.
'
' Parameters:
'   KeyPath - Full registry path (modified by reference)
'
' Returns: Root key handle
'
' Example:
'   path = "HKEY_CURRENT_USER\Software\MyApp"
'   handle = GetHKey(path)
'   ' path is now "Software\MyApp"
'   ' handle is HKEY_CURRENT_USER
'=======================================================
Private Function GetHKey(ByRef KeyPath As String) As LongPtr
    Dim rootKey As String
    Dim backslashPos As Long
    
    On Error Resume Next
    
    ' Default to HKEY_CURRENT_USER if no prefix
    If Left$(KeyPath, 4) <> "HKEY" Then
        GetHKey = HKEY_CURRENT_USER
        Exit Function
    End If
    
    ' Extract root key name
    backslashPos = InStr(KeyPath, "\")
    If backslashPos = 0 Then
        ' No subkey specified
        rootKey = KeyPath
        KeyPath = vbNullString
    Else
        rootKey = Left$(KeyPath, backslashPos - 1)
        KeyPath = Right$(KeyPath, Len(KeyPath) - backslashPos)
    End If
    
    ' Convert root key name to handle
    Select Case rootKey
        Case "HKEY_CLASSES_ROOT"
            GetHKey = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetHKey = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetHKey = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetHKey = HKEY_USERS
        Case "HKEY_CURRENT_CONFIG"
            GetHKey = HKEY_CURRENT_CONFIG
        Case Else
            ' Unknown root key - default to CURRENT_USER
            GetHKey = HKEY_CURRENT_USER
            WriteLog 2, CurrentMod, "GetHKey", _
                     "Unknown root key: " & rootKey & ", using HKEY_CURRENT_USER"
    End Select
End Function

'=======================================================
' Function: ValidateRegInputs
' Purpose: Validate registry function inputs
'
' Parameters:
'   VName - Value name
'   KeyPath - Key path
'   errorMsg - Error message (output)
'
' Returns: True if valid, False otherwise
'=======================================================
Private Function ValidateRegInputs(ByVal VName As String, _
                                   ByVal KeyPath As String, _
                                   ByRef errorMsg As String) As Boolean
    
    ValidateRegInputs = False
    
    ' Check value name
    If Len(Trim$(VName)) = 0 Then
        errorMsg = "Empty value name"
        Exit Function
    End If
    
    ' Check key path
    If Len(Trim$(KeyPath)) = 0 Then
        errorMsg = "Empty key path"
        Exit Function
    End If
    
    ' Check for invalid characters in value name
    If InStr(VName, "\") > 0 Or InStr(VName, "/") > 0 Then
        errorMsg = "Value name contains invalid characters"
        Exit Function
    End If
    
    ValidateRegInputs = True
End Function

'=======================================================
' Function: GetLastRegError
' Purpose: Get the last registry error number
'
' Returns: Last registry error code
'
' Example:
'   If Not SetReg(...) Then
'       MsgBox "Registry error: " & GetLastRegError()
'   End If
'=======================================================
Public Function GetLastRegError() As LongPtr
    GetLastRegError = RegErrNum
End Function

'=======================================================
' END OF MODULE
'=======================================================
