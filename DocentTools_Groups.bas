Attribute VB_Name = "DocentTools_Groups"
Option Explicit
Option Compare Text

'=======================================================
' Module: DocentTools_Groups
' Purpose: Group and member management functions
' Author: Refactored with error handling - November 2025
' Version: 2.0
'
' Description:
'   Provides functions for working with project groups
'   and members including:
'   - Retrieving all groups or user-specific groups
'   - Getting group members
'   - Looking up member information by ID, name, or email
'
' Dependencies:
'   - AB_GlobalVars (ProjectGroupsDict, UserGroupsDict, MembersDict)
'   - Dictionary class
'   - Ribbon_Manager_Mod (RefreshRibbon)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added function documentation
'       * Improved null checking
'       * Added logging
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "DocentTools_Groups"

'=======================================================
' Function: GetAllGroups
' Purpose: Get all project groups, optionally filtered
'
' Parameters:
'   ExceptMask - Optional wildcard pattern to exclude groups (e.g., "*Admin*")
'   ForMapping - If True, also adds groups by ID for easy lookup
'
' Returns: Dictionary of groups (key: title, value: group dictionary)
'          Returns empty Dictionary on error
'
' Example:
'   Dim groups As Dictionary
'   Set groups = GetAllGroups()  ' Get all groups
'   Set groups = GetAllGroups("*Admin*")  ' Exclude admin groups
'=======================================================
Public Function GetAllGroups(Optional ByVal ExceptMask As String = vbNullString, _
                            Optional ByVal ForMapping As Boolean = False) As Dictionary
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetAllGroups", "Retrieving groups"
    
    ' Initialize return value
    Set GetAllGroups = New Dictionary
    
    ' Ensure ProjectGroupsDict is loaded
    If ProjectGroupsDict Is Nothing Then
        WriteLog 2, CurrentMod, "GetAllGroups", "ProjectGroupsDict is Nothing, refreshing ribbon"
        RefreshRibbon
        
        ' Check again after refresh
        If ProjectGroupsDict Is Nothing Then
            WriteLog 3, CurrentMod, "GetAllGroups", "ProjectGroupsDict still Nothing after refresh"
            Exit Function
        End If
    End If
    
    ' Validate dictionary has items
    If ProjectGroupsDict.Count = 0 Then
        WriteLog 2, CurrentMod, "GetAllGroups", "ProjectGroupsDict is empty"
        Exit Function
    End If
    
    ' Build groups dictionary
    For i = 1 To ProjectGroupsDict.Count
        ' Validate group has required fields
        If Not ValidateGroupDictionary(ProjectGroupsDict(i)) Then
            WriteLog 2, CurrentMod, "GetAllGroups", _
                     "Group " & i & " missing required fields, skipping"
            GoTo NextGroup
        End If
        
        ' Apply filter if specified
        If Len(ExceptMask) > 0 Then
            If ProjectGroupsDict.KeyName(i) Like ExceptMask Then
                GoTo NextGroup
            End If
        End If
        
        ' Add by title
        GetAllGroups.Add ProjectGroupsDict(i)("title"), ProjectGroupsDict(i)
        
        ' Also add by ID if requested
        If ForMapping Then
            If ProjectGroupsDict(i).Exists("id") Then
                GetAllGroups.Add ProjectGroupsDict(i)("id"), ProjectGroupsDict(i)
            End If
        End If
        
NextGroup:
    Next i
    
    WriteLog 1, CurrentMod, "GetAllGroups", "Retrieved " & GetAllGroups.Count & " groups"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetAllGroups", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetAllGroups = New Dictionary
End Function

'=======================================================
' Function: GetMyGroups
' Purpose: Get groups that the current user belongs to
'
' Parameters:
'   ExceptMask - Optional wildcard pattern to exclude groups
'   ForMapping - If True, also adds groups by ID for easy lookup
'
' Returns: Dictionary of user's groups (key: title, value: group dictionary)
'          Returns empty Dictionary on error
'=======================================================
Public Function GetMyGroups(Optional ByVal ExceptMask As String = vbNullString, _
                           Optional ByVal ForMapping As Boolean = False) As Dictionary
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMyGroups", "Retrieving user groups"
    
    ' Initialize return value
    Set GetMyGroups = New Dictionary
    
    ' Ensure UserGroupsDict is loaded
    If UserGroupsDict Is Nothing Then
        WriteLog 2, CurrentMod, "GetMyGroups", "UserGroupsDict is Nothing, refreshing ribbon"
        RefreshRibbon
        
        ' Check again after refresh
        If UserGroupsDict Is Nothing Then
            WriteLog 3, CurrentMod, "GetMyGroups", "UserGroupsDict still Nothing after refresh"
            Exit Function
        End If
    End If
    
    ' Validate dictionary has items
    If UserGroupsDict.Count = 0 Then
        WriteLog 2, CurrentMod, "GetMyGroups", "UserGroupsDict is empty"
        Exit Function
    End If
    
    ' Build groups dictionary
    For i = 1 To UserGroupsDict.Count
        ' Validate group has required fields
        If Not ValidateGroupDictionary(UserGroupsDict(i)) Then
            WriteLog 2, CurrentMod, "GetMyGroups", _
                     "Group " & i & " missing required fields, skipping"
            GoTo NextGroup
        End If
        
        ' Apply filter if specified
        If Len(ExceptMask) > 0 Then
            If UserGroupsDict.KeyName(i) Like ExceptMask Then
                GoTo NextGroup
            End If
        End If
        
        ' Add by title
        GetMyGroups.Add UserGroupsDict(i)("title"), UserGroupsDict(i)
        
        ' Also add by ID if requested
        If ForMapping Then
            If UserGroupsDict(i).Exists("id") Then
                GetMyGroups.Add UserGroupsDict(i)("id"), UserGroupsDict(i)
            End If
        End If
        
NextGroup:
    Next i
    
    WriteLog 1, CurrentMod, "GetMyGroups", "Retrieved " & GetMyGroups.Count & " user groups"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMyGroups", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMyGroups = New Dictionary
End Function

'=======================================================
' Function: GetMembersOf
' Purpose: Get members of a specific group
'
' Parameters:
'   GroupName - Name of the group (default: "PrjTeam")
'   GetWhat - Field to retrieve (e.g., "fullname", "email", "id")
'            If empty, returns full member dictionary
'
' Returns: Dictionary of members (key: id, value: requested field or full dict)
'          Returns empty Dictionary on error
'
' Example:
'   Dim members As Dictionary
'   Set members = GetMembersOf("PrjTeam", "fullname")  ' Get names
'   Set members = GetMembersOf("PrjTeam", "")  ' Get full info
'=======================================================
Public Function GetMembersOf(Optional ByVal GroupName As String = "PrjTeam", _
                            Optional ByVal GetWhat As String = "fullname") As Dictionary
    Dim i As Long
    Dim GroupInfo As Dictionary
    Dim MembersList As Collection
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMembersOf", "Getting members of: " & GroupName
    
    ' Initialize return value
    Set GetMembersOf = New Dictionary
    
    ' Validate group name
    If Len(GroupName) = 0 Then
        WriteLog 3, CurrentMod, "GetMembersOf", "Empty group name"
        Exit Function
    End If
    
    ' Ensure ProjectGroupsDict is loaded
    If ProjectGroupsDict Is Nothing Then
        WriteLog 3, CurrentMod, "GetMembersOf", "ProjectGroupsDict is Nothing"
        Exit Function
    End If
    
    ' Check if group exists
    If Not ProjectGroupsDict.Exists(GroupName) Then
        WriteLog 3, CurrentMod, "GetMembersOf", "Group not found: " & GroupName
        Exit Function
    End If
    
    ' Get group info
    Set GroupInfo = ProjectGroupsDict(GroupName)
    
    ' Check if group has members
    If Not GroupInfo.Exists("groupMembers") Then
        WriteLog 2, CurrentMod, "GetMembersOf", "Group has no groupMembers key: " & GroupName
        Exit Function
    End If
    
    Set MembersList = GroupInfo("groupMembers")
    
    If MembersList.Count = 0 Then
        WriteLog 2, CurrentMod, "GetMembersOf", "Group has no members: " & GroupName
        Exit Function
    End If
    
    ' Build members dictionary
    For i = 1 To MembersList.Count
        ' Validate member has ID
        If Not MembersList(i).Exists("id") Then
            WriteLog 2, CurrentMod, "GetMembersOf", _
                     "Member " & i & " missing ID, skipping"
            GoTo NextMember
        End If
        
        ' Add member info
        If Len(GetWhat) > 0 Then
            ' Get specific field
            If MembersList(i).Exists(GetWhat) Then
                GetMembersOf.Add MembersList(i)("id"), MembersList(i)(GetWhat)
            Else
                WriteLog 2, CurrentMod, "GetMembersOf", _
                         "Member " & i & " missing field: " & GetWhat
                GetMembersOf.Add MembersList(i)("id"), vbNullString
            End If
        Else
            ' Get full dictionary
            GetMembersOf.Add MembersList(i)("id"), MembersList(i)
        End If
        
NextMember:
    Next i
    
    WriteLog 1, CurrentMod, "GetMembersOf", _
             "Retrieved " & GetMembersOf.Count & " members"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMembersOf", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMembersOf = New Dictionary
End Function

'=======================================================
' Function: GetAllMembers
' Purpose: Get all project members
'
' Parameters:
'   FromField - Field to use as key (default: "fullname")
'
' Returns: Dictionary of members (key: FromField value, value: member dict)
'          Returns empty Dictionary on error
'=======================================================
Public Function GetAllMembers(Optional ByVal FromField As String = "fullname") As Dictionary
    Dim i As Long
    Dim keyValue As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetAllMembers", "Retrieving all members"
    
    ' Initialize return value
    Set GetAllMembers = New Dictionary
    
    ' Validate MembersDict
    If MembersDict Is Nothing Then
        WriteLog 3, CurrentMod, "GetAllMembers", "MembersDict is Nothing"
        Exit Function
    End If
    
    If MembersDict.Count = 0 Then
        WriteLog 2, CurrentMod, "GetAllMembers", "MembersDict is empty"
        Exit Function
    End If
    
    ' Build members dictionary
    For i = 1 To MembersDict.Count
        ' Validate member has required field
        If MembersDict(i).Exists(FromField) Then
            keyValue = MembersDict(i)(FromField)
            
            ' Skip if key is empty
            If Len(keyValue) = 0 Then
                WriteLog 2, CurrentMod, "GetAllMembers", _
                         "Member " & i & " has empty " & FromField & ", skipping"
                GoTo NextMember
            End If
            
            ' Add member (avoid duplicate key errors)
            If Not GetAllMembers.Exists(keyValue) Then
                GetAllMembers.Add keyValue, MembersDict(i)
            Else
                WriteLog 2, CurrentMod, "GetAllMembers", _
                         "Duplicate key found: " & keyValue
            End If
        Else
            WriteLog 2, CurrentMod, "GetAllMembers", _
                     "Member " & i & " missing field: " & FromField
        End If
        
NextMember:
    Next i
    
    WriteLog 1, CurrentMod, "GetAllMembers", _
             "Retrieved " & GetAllMembers.Count & " members"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetAllMembers", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetAllMembers = New Dictionary
End Function

'=======================================================
' Function: GetMyMembers
' Purpose: Get members relevant to current user (NEEDS TESTING)
'
' Parameters:
'   FromField - Field to use as key (default: "fullname")
'
' Returns: Dictionary of members
'          Returns empty Dictionary on error
'=======================================================
Public Function GetMyMembers(Optional ByVal FromField As String = "fullname") As Dictionary
    Dim i As Long
    Dim LocalMembersDict As Dictionary
    Dim keyValue As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMyMembers", "Retrieving user members"
    
    ' Initialize return value
    Set GetMyMembers = New Dictionary
    
    ' Ensure MainInfo is loaded
    If MainInfo Is Nothing Then
        WriteLog 2, CurrentMod, "GetMyMembers", "MainInfo is Nothing, refreshing ribbon"
        RefreshRibbon
        
        If MainInfo Is Nothing Then
            WriteLog 3, CurrentMod, "GetMyMembers", "MainInfo still Nothing after refresh"
            Exit Function
        End If
    End If
    
    ' Get members dictionary
    Set LocalMembersDict = GetMembersDict(MainInfo)
    
    If LocalMembersDict Is Nothing Then
        WriteLog 3, CurrentMod, "GetMyMembers", "GetMembersDict returned Nothing"
        Exit Function
    End If
    
    If LocalMembersDict.Count = 0 Then
        WriteLog 2, CurrentMod, "GetMyMembers", "LocalMembersDict is empty"
        Exit Function
    End If
    
    ' Build members dictionary
    For i = 1 To LocalMembersDict.Count
        ' Validate member has required field
        If LocalMembersDict(i).Exists(FromField) Then
            keyValue = LocalMembersDict(i)(FromField)
            
            ' Skip if key is empty
            If Len(keyValue) = 0 Then
                GoTo NextMember
            End If
            
            ' Add member (avoid duplicate key errors)
            If Not GetMyMembers.Exists(keyValue) Then
                GetMyMembers.Add keyValue, LocalMembersDict(i)
            End If
        End If
        
NextMember:
    Next i
    
    WriteLog 1, CurrentMod, "GetMyMembers", _
             "Retrieved " & GetMyMembers.Count & " user members"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMyMembers", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMyMembers = New Dictionary
End Function

'=======================================================
' Function: GetMembersNames
' Purpose: Get member names from a group or all members
'
' Parameters:
'   GroupName - Optional group name; if empty, gets all members
'
' Returns: Dictionary of member names (key: fullname, value: fullname)
'          Returns empty Dictionary on error
'=======================================================
Public Function GetMembersNames(Optional ByVal GroupName As String = vbNullString) As Dictionary
    Dim Members As Dictionary
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMembersNames", "Getting member names"
    
    ' Initialize return value
    Set GetMembersNames = New Dictionary
    
    If Len(GroupName) > 0 Then
        ' Get members from specific group
        Set GetMembersNames = GetMembersOf(GroupName)
    Else
        ' Get all members and extract names
        Set Members = GetAllMembers()
        
        If Members Is Nothing Then
            WriteLog 3, CurrentMod, "GetMembersNames", "GetAllMembers returned Nothing"
            Exit Function
        End If
        
        For i = 1 To Members.Count
            If Members(i).Exists("fullname") Then
                Dim memberName As String
                memberName = Members(i)("fullname")
                
                If Len(memberName) > 0 Then
                    If Not GetMembersNames.Exists(memberName) Then
                        GetMembersNames.Add memberName, memberName
                    End If
                End If
            End If
        Next i
    End If
    
    WriteLog 1, CurrentMod, "GetMembersNames", _
             "Retrieved " & GetMembersNames.Count & " member names"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMembersNames", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMembersNames = New Dictionary
End Function

'=======================================================
' Function: GetUserInfo
' Purpose: Get specific user information by lookup
'
' Parameters:
'   What - Value to search for
'   AsWhat - Field to search in (e.g., "fullname", "email", "id")
'   ReturnWhat - Field to return
'
' Returns: Requested field value, or empty string on error
'
' Example:
'   Dim email As String
'   email = GetUserInfo("John Doe", "fullname", "email")
'=======================================================
Public Function GetUserInfo(ByVal What As String, _
                           ByVal AsWhat As String, _
                           ByVal ReturnWhat As String) As String
    Dim AllMembers As Dictionary
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Len(What) = 0 Or Len(AsWhat) = 0 Or Len(ReturnWhat) = 0 Then
        WriteLog 3, CurrentMod, "GetUserInfo", "Empty parameter(s)"
        GetUserInfo = vbNullString
        Exit Function
    End If
    
    ' Get all members indexed by AsWhat field
    Set AllMembers = GetAllMembers(AsWhat)
    
    If AllMembers Is Nothing Then
        GetUserInfo = vbNullString
        Exit Function
    End If
    
    ' Look up user
    If AllMembers.Exists(What) Then
        If AllMembers(What).Exists(ReturnWhat) Then
            GetUserInfo = AllMembers(What)(ReturnWhat)
        Else
            WriteLog 2, CurrentMod, "GetUserInfo", _
                     "User found but missing field: " & ReturnWhat
            GetUserInfo = vbNullString
        End If
    Else
        GetUserInfo = vbNullString
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetUserInfo", _
             "Error " & Err.Number & ": " & Err.Description
    GetUserInfo = vbNullString
End Function

'=======================================================
' Function: GetMemberID
' Purpose: Get member ID by name or email
'
' Parameters:
'   UserNameOrEmail - Member's full name or email address
'
' Returns: Member ID, or empty string if not found
'=======================================================
Public Function GetMemberID(ByVal UserNameOrEmail As String) As String
    On Error GoTo ErrorHandler
    
    If Len(UserNameOrEmail) = 0 Then
        GetMemberID = vbNullString
        Exit Function
    End If
    
    ' Try by fullname first
    GetMemberID = GetUserInfo(UserNameOrEmail, "fullname", "id")
    
    ' If not found, try by email
    If Len(GetMemberID) = 0 Then
        GetMemberID = GetUserInfo(UserNameOrEmail, "email", "id")
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMemberID", _
             "Error " & Err.Number & ": " & Err.Description
    GetMemberID = vbNullString
End Function

'=======================================================
' Function: GetMemberEmail
' Purpose: Get member email by name or ID
'
' Parameters:
'   UserNameOrID - Member's full name or ID
'
' Returns: Member email, or empty string if not found
'=======================================================
Public Function GetMemberEmail(ByVal UserNameOrID As String) As String
    On Error GoTo ErrorHandler
    
    If Len(UserNameOrID) = 0 Then
        GetMemberEmail = vbNullString
        Exit Function
    End If
    
    ' Try by fullname first
    GetMemberEmail = GetUserInfo(UserNameOrID, "fullname", "email")
    
    ' If not found, try by ID
    If Len(GetMemberEmail) = 0 Then
        GetMemberEmail = GetUserInfo(UserNameOrID, "id", "email")
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMemberEmail", _
             "Error " & Err.Number & ": " & Err.Description
    GetMemberEmail = vbNullString
End Function

'=======================================================
' Function: GetMemberName
' Purpose: Get member name by email or ID
'
' Parameters:
'   UserEmailOrID - Member's email or ID
'
' Returns: Member full name, or empty string if not found
'=======================================================
Public Function GetMemberName(ByVal UserEmailOrID As String) As String
    On Error GoTo ErrorHandler
    
    If Len(UserEmailOrID) = 0 Then
        GetMemberName = vbNullString
        Exit Function
    End If
    
    ' Try by email first
    GetMemberName = GetUserInfo(UserEmailOrID, "email", "fullname")
    
    ' If not found, try by ID
    If Len(GetMemberName) = 0 Then
        GetMemberName = GetUserInfo(UserEmailOrID, "id", "fullname")
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMemberName", _
             "Error " & Err.Number & ": " & Err.Description
    GetMemberName = vbNullString
End Function

'=======================================================
' PRIVATE HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: ValidateGroupDictionary
' Purpose: Validate that a group dictionary has required fields
'
' Parameters:
'   GroupDict - Group dictionary to validate
'
' Returns: True if valid, False otherwise
'=======================================================
Private Function ValidateGroupDictionary(ByVal GroupDict As Dictionary) As Boolean
    On Error GoTo ErrorHandler
    
    If GroupDict Is Nothing Then
        ValidateGroupDictionary = False
        Exit Function
    End If
    
    ' Check for required fields
    If Not GroupDict.Exists("title") Then
        ValidateGroupDictionary = False
        Exit Function
    End If
    
    If Not GroupDict.Exists("id") Then
        ValidateGroupDictionary = False
        Exit Function
    End If
    
    ValidateGroupDictionary = True
    Exit Function
    
ErrorHandler:
    ValidateGroupDictionary = False
End Function

'=======================================================
' END OF MODULE
'=======================================================
