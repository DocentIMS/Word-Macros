Attribute VB_Name = "AC1_Meetings_mod"
Option Explicit
Option Compare Text

'=======================================================
' Module: AC1_Meetings_mod
' Purpose: Meeting type management functions
' Author: Refactored with error handling - November 2025
' Version: 2.0
'
' Description:
'   Provides functions for working with meeting types
'   including retrieval and validation of meeting type
'   information from project settings.
'
' Dependencies:
'   - AB_GlobalVars (ProjectInfo)
'   - Dictionary class
'   - Ribbon_Manager_Mod (RefreshRibbon)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added function documentation
'       * Added logging
'       * Enhanced null checking
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Meetings"

'=======================================================
' Function: GetMeetingTypes
' Purpose: Get all available meeting types for the project
'
' Returns: Dictionary of meeting types
'          Key: meeting_type string (e.g., "Daily Standup")
'          Value: meeting type dictionary with full details
'          Returns empty Dictionary on error
'
' Example:
'   Dim meetingTypes As Dictionary
'   Set meetingTypes = GetMeetingTypes()
'   If meetingTypes.Count > 0 Then
'       Debug.Print "First type: " & meetingTypes(1)("meeting_type")
'   End If
'
' Note:
'   Meeting types are stored in ProjectInfo("meeting_types")
'   If not loaded, will attempt to refresh from ribbon/server
'=======================================================
Public Function GetMeetingTypes() As Dictionary
    Dim i As Long
    Dim MeetingTypesCollection As Collection
    Dim meetingType As Dictionary
    Dim keyValue As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMeetingTypes", "Retrieving meeting types"
    
    ' Initialize return value
    Set GetMeetingTypes = New Dictionary
    
    ' Validate ProjectInfo exists
    If ProjectInfo Is Nothing Then
        WriteLog 3, CurrentMod, "GetMeetingTypes", _
                 "ProjectInfo is Nothing, cannot retrieve meeting types"
        Exit Function
    End If
    
    ' Check if meeting_types key exists
    If Not ProjectInfo.Exists("meeting_types") Then
        WriteLog 2, CurrentMod, "GetMeetingTypes", _
                 "meeting_types key not found in ProjectInfo, refreshing ribbon"
        RefreshRibbon
        
        ' Check again after refresh
        If Not ProjectInfo.Exists("meeting_types") Then
            WriteLog 3, CurrentMod, "GetMeetingTypes", _
                     "meeting_types key still not found after refresh"
            Exit Function
        End If
    End If
    
    ' Check if meeting_types is Empty
    If IsEmpty(ProjectInfo("meeting_types")) Then
        WriteLog 2, CurrentMod, "GetMeetingTypes", _
                 "meeting_types is Empty, refreshing ribbon"
        RefreshRibbon
        
        ' Check again after refresh
        If IsEmpty(ProjectInfo("meeting_types")) Then
            WriteLog 3, CurrentMod, "GetMeetingTypes", _
                     "meeting_types still Empty after refresh"
            Exit Function
        End If
    End If
    
    ' Get meeting types collection
    Set MeetingTypesCollection = ProjectInfo("meeting_types")
    
    ' Validate it's actually a collection
    If MeetingTypesCollection Is Nothing Then
        WriteLog 3, CurrentMod, "GetMeetingTypes", "meeting_types collection is Nothing"
        Exit Function
    End If
    
    ' Check if collection has items
    If MeetingTypesCollection.Count = 0 Then
        WriteLog 2, CurrentMod, "GetMeetingTypes", "meeting_types collection is empty"
        Exit Function
    End If
    
    ' Build meeting types dictionary
    For i = 1 To MeetingTypesCollection.Count
        ' Validate meeting type is a dictionary
        If TypeName(MeetingTypesCollection(i)) <> "Dictionary" Then
            WriteLog 2, CurrentMod, "GetMeetingTypes", _
                     "Meeting type " & i & " is not a Dictionary, skipping"
            GoTo NextMeetingType
        End If
        
        Set meetingType = MeetingTypesCollection(i)
        
        ' Validate meeting type has required field
        If Not meetingType.Exists("meeting_type") Then
            WriteLog 2, CurrentMod, "GetMeetingTypes", _
                     "Meeting type " & i & " missing 'meeting_type' field, skipping"
            GoTo NextMeetingType
        End If
        
        keyValue = meetingType("meeting_type")
        
        ' Validate key is not empty
        If Len(keyValue) = 0 Then
            WriteLog 2, CurrentMod, "GetMeetingTypes", _
                     "Meeting type " & i & " has empty 'meeting_type' value, skipping"
            GoTo NextMeetingType
        End If
        
        ' Add to result dictionary (avoid duplicate keys)
        If Not GetMeetingTypes.Exists(keyValue) Then
            GetMeetingTypes.Add keyValue, meetingType
        Else
            WriteLog 2, CurrentMod, "GetMeetingTypes", _
                     "Duplicate meeting type found: " & keyValue
        End If
        
NextMeetingType:
    Next i
    
    WriteLog 1, CurrentMod, "GetMeetingTypes", _
             "Retrieved " & GetMeetingTypes.Count & " meeting types"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMeetingTypes", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMeetingTypes = New Dictionary
End Function

'=======================================================
' Function: GetMeetingTypeByName
' Purpose: Get a specific meeting type by name
'
' Parameters:
'   TypeName - Name of the meeting type to retrieve
'
' Returns: Dictionary with meeting type details, or Nothing if not found
'
' Example:
'   Dim mtgType As Dictionary
'   Set mtgType = GetMeetingTypeByName("Daily Standup")
'   If Not mtgType Is Nothing Then
'       Debug.Print "Found meeting type: " & mtgType("meeting_type")
'   End If
'=======================================================
Public Function GetMeetingTypeByName(ByVal TypeName As String) As Dictionary
    Dim AllTypes As Dictionary
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMeetingTypeByName", "Looking for: " & TypeName
    
    ' Validate input
    If Len(TypeName) = 0 Then
        WriteLog 3, CurrentMod, "GetMeetingTypeByName", "Empty type name"
        Set GetMeetingTypeByName = Nothing
        Exit Function
    End If
    
    ' Get all meeting types
    Set AllTypes = GetMeetingTypes()
    
    If AllTypes Is Nothing Or AllTypes.Count = 0 Then
        WriteLog 3, CurrentMod, "GetMeetingTypeByName", "No meeting types available"
        Set GetMeetingTypeByName = Nothing
        Exit Function
    End If
    
    ' Look up the requested type
    If AllTypes.Exists(TypeName) Then
        Set GetMeetingTypeByName = AllTypes(TypeName)
        WriteLog 1, CurrentMod, "GetMeetingTypeByName", "Meeting type found"
    Else
        WriteLog 2, CurrentMod, "GetMeetingTypeByName", "Meeting type not found: " & TypeName
        Set GetMeetingTypeByName = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMeetingTypeByName", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMeetingTypeByName = Nothing
End Function

'=======================================================
' Function: ValidateMeetingType
' Purpose: Check if a meeting type name is valid
'
' Parameters:
'   TypeName - Name of the meeting type to validate
'
' Returns: True if the meeting type exists, False otherwise
'
' Example:
'   If ValidateMeetingType("Daily Standup") Then
'       ' Meeting type is valid
'   End If
'=======================================================
Public Function ValidateMeetingType(ByVal TypeName As String) As Boolean
    Dim meetingType As Dictionary
    
    On Error GoTo ErrorHandler
    
    If Len(TypeName) = 0 Then
        ValidateMeetingType = False
        Exit Function
    End If
    
    Set meetingType = GetMeetingTypeByName(TypeName)
    ValidateMeetingType = Not (meetingType Is Nothing)
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ValidateMeetingType", _
             "Error " & Err.Number & ": " & Err.Description
    ValidateMeetingType = False
End Function

'=======================================================
' Function: GetMeetingTypesList
' Purpose: Get simple list of meeting type names
'
' Returns: Collection of meeting type names (strings)
'          Returns empty Collection on error
'
' Example:
'   Dim typeNames As Collection
'   Set typeNames = GetMeetingTypesList()
'   Dim i As Long
'   For i = 1 To typeNames.Count
'       Debug.Print typeNames(i)
'   Next i
'=======================================================
Public Function GetMeetingTypesList() As Collection
    Dim AllTypes As Dictionary
    Dim TypeNames As Collection
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetMeetingTypesList", "Getting meeting types list"
    
    ' Initialize return value
    Set TypeNames = New Collection
    
    ' Get all meeting types
    Set AllTypes = GetMeetingTypes()
    
    If AllTypes Is Nothing Then
        WriteLog 3, CurrentMod, "GetMeetingTypesList", "GetMeetingTypes returned Nothing"
        Set GetMeetingTypesList = TypeNames
        Exit Function
    End If
    
    ' Extract type names
    For i = 1 To AllTypes.Count
        If AllTypes(i).Exists("meeting_type") Then
            TypeNames.Add AllTypes(i)("meeting_type")
        End If
    Next i
    
    Set GetMeetingTypesList = TypeNames
    
    WriteLog 1, CurrentMod, "GetMeetingTypesList", _
             "Retrieved " & TypeNames.Count & " meeting type names"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetMeetingTypesList", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetMeetingTypesList = New Collection
End Function

'=======================================================
' END OF MODULE
'=======================================================
