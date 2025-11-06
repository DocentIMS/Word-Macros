Attribute VB_Name = "AA_Zoom"
Option Explicit
Option Private Module
Option Compare Text

'=======================================================
' Module: AA_Zoom
' Purpose: Zoom API integration for meeting management
' Author: Refactored with error handling - November 2025
' Version: 2.0
'
' Description:
'   Handles all Zoom API operations including:
'   - Creating meetings with custom settings
'   - Retrieving meeting details
'   - Managing past meetings
'   - Getting participant lists
'   - Working with meeting templates
'
' Dependencies:
'   - BearerToken class (ZoomToken)
'   - WebClient, WebRequest, WebResponse classes
'   - AC_API_Mod2 (UseAPI function)
'   - AB_CommonTools (DblEncode, CreateBody)
'   - AB_CommonFunctions (ToServerTime)
'   - AB_GlobalVars (ZoomToken)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added function documentation
'       * Removed commented dead code
'       * Added logging
'   v1.0 - Original version
'=======================================================

' Module constants
Private Const CurrentMod As String = "Zoom"
Private Const ZoomAPI As String = "/noapi/https://api.zoom.us/v2/"

' Zoom meeting constants
Private Const ZOOM_DEFAULT_PASSWORD As String = "DocentIMS"
Private Const ZOOM_DEFAULT_TIMEZONE As String = "America/Los_Angeles"
Private Const ZOOM_DEFAULT_SUMMARY_TEMPLATE As String = "654115c4"
Private Const ZOOM_DEFAULT_MEETING_TEMPLATE As String = "5POT1ygMSp-FEMLd8-Uw4Q"

'=======================================================
' Function: CreateZoomMeeting
' Purpose: Create a new Zoom meeting with specified settings
'
' Parameters:
'   MtgTitle - Meeting title/topic
'   MtgStart - Meeting start date/time
'   MtgDuration - Meeting duration in minutes
'   MtgSettings - Dictionary with meeting settings
'
' Returns: WebResponse object, or Nothing on error
'
' Example:
'   Dim response As WebResponse
'   Dim settings As New Dictionary
'   settings.Add "waiting_room", True
'   Set response = CreateZoomMeeting("Team Meeting", Now + 1, 60, settings)
'=======================================================
Public Function CreateZoomMeeting(ByVal MtgTitle As String, _
                                  ByVal MtgStart As Date, _
                                  ByVal MtgDuration As Long, _
                                  ByVal MtgSettings As Dictionary) As WebResponse
    Dim Body As Dictionary
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "CreateZoomMeeting", "Creating meeting: " & MtgTitle
    
    ' Validate inputs
    If Not ValidateCreateMeetingInputs(MtgTitle, MtgStart, MtgDuration, errorMsg) Then
        WriteLog 3, CurrentMod, "CreateZoomMeeting", "Validation failed: " & errorMsg
        Set CreateZoomMeeting = Nothing
        Exit Function
    End If
    
    ' Validate settings dictionary
    If MtgSettings Is Nothing Then
        WriteLog 3, CurrentMod, "CreateZoomMeeting", "MtgSettings is Nothing"
        Set CreateZoomMeeting = Nothing
        Exit Function
    End If
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' Apply default settings
    Call ApplyDefaultMeetingSettings(MtgSettings)
    
    ' Build request body
    Set Body = CreateBody( _
        Array("topic", "type", "start_time", "duration", "timezone", "password", "settings", "template_id"), _
        Array(MtgTitle, 2, ToServerTime(CStr(MtgStart)), MtgDuration, ZOOM_DEFAULT_TIMEZONE, ZOOM_DEFAULT_PASSWORD, MtgSettings, ZOOM_DEFAULT_MEETING_TEMPLATE), _
        False)
    
    ' Execute API call
    Set CreateZoomMeeting = UseAPI(Body, HttpPost, "meetings", "users/me/", _
                                   ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Log result
    If CreateZoomMeeting Is Nothing Then
        WriteLog 3, CurrentMod, "CreateZoomMeeting", "Failed to create meeting"
    ElseIf CreateZoomMeeting.StatusCode = HTTP_CREATED Then
        WriteLog 1, CurrentMod, "CreateZoomMeeting", "Meeting created successfully"
    Else
        WriteLog 2, CurrentMod, "CreateZoomMeeting", _
                 "Meeting creation returned status: " & CreateZoomMeeting.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "CreateZoomMeeting", _
             "Error " & Err.Number & ": " & Err.Description
    Set CreateZoomMeeting = Nothing
End Function

'=======================================================
' Function: GetZoomMtg
' Purpose: Get details of a Zoom meeting by ID
'
' Parameters:
'   ID - Zoom meeting ID
'
' Returns: Dictionary with meeting details, or Nothing on error
'=======================================================
Public Function GetZoomMtg(ByVal ID As String) As Dictionary
    Dim Response As WebResponse
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(ID) = 0 Then
        WriteLog 3, CurrentMod, "GetZoomMtg", "Empty meeting ID"
        Set GetZoomMtg = Nothing
        Exit Function
    End If
    
    WriteLog 1, CurrentMod, "GetZoomMtg", "Getting meeting: " & ID
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' Get meeting details
    Set Response = UseAPI(Nothing, HttpGet, DblEncode(ID), "/meetings/", _
                         ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Process response
    If Response Is Nothing Then
        WriteLog 3, CurrentMod, "GetZoomMtg", "Failed to get meeting"
    ElseIf Response.StatusCode = HTTP_OK Then
        Set GetZoomMtg = Response.Data
        WriteLog 1, CurrentMod, "GetZoomMtg", "Meeting retrieved successfully"
    Else
        WriteLog 3, CurrentMod, "GetZoomMtg", _
                 "Failed with status: " & Response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetZoomMtg", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetZoomMtg = Nothing
End Function

'=======================================================
' Function: GetZoomMtgUUID
' Purpose: Get UUID of a Zoom meeting by ID
'
' Parameters:
'   ID - Zoom meeting ID
'
' Returns: Meeting UUID string, or empty string on error
'=======================================================
Public Function GetZoomMtgUUID(ByVal ID As String) As String
    Dim MeetingInfo As Dictionary
    
    On Error GoTo ErrorHandler
    
    Set MeetingInfo = GetZoomMtg(ID)
    
    If Not MeetingInfo Is Nothing Then
        If MeetingInfo.Exists("uuid") Then
            GetZoomMtgUUID = MeetingInfo("uuid")
        Else
            WriteLog 2, CurrentMod, "GetZoomMtgUUID", "UUID not found in meeting info"
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetZoomMtgUUID", _
             "Error " & Err.Number & ": " & Err.Description
    GetZoomMtgUUID = vbNullString
End Function

'=======================================================
' Function: GetPastMeeting
' Purpose: Get details of a past Zoom meeting
'
' Parameters:
'   ID - Meeting ID or UUID
'
' Returns: Dictionary with past meeting details, or Nothing on error
'=======================================================
Public Function GetPastMeeting(ByVal ID As String) As Dictionary
    Dim Response As WebResponse
    Dim Coll As Collection
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(ID) = 0 Then
        WriteLog 3, CurrentMod, "GetPastMeeting", "Empty meeting ID"
        Set GetPastMeeting = Nothing
        Exit Function
    End If
    
    WriteLog 1, CurrentMod, "GetPastMeeting", "Getting past meeting: " & ID
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' If numeric ID, get the UUID from past instances
    If IsNumeric(ID) Then
        Set Coll = GetPastMeetings(ID)
        If Not Coll Is Nothing Then
            If Coll.Count > 0 Then
                ID = Coll(Coll.Count)("uuid")
            Else
                WriteLog 2, CurrentMod, "GetPastMeeting", "No past instances found"
                Set GetPastMeeting = Nothing
                Exit Function
            End If
        Else
            WriteLog 3, CurrentMod, "GetPastMeeting", "Failed to get past meetings"
            Set GetPastMeeting = Nothing
            Exit Function
        End If
    End If
    
    ' Get past meeting details
    Set Response = UseAPI(Nothing, HttpGet, DblEncode(ID), "/past_meetings/", _
                         ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Process response
    If Response Is Nothing Then
        WriteLog 3, CurrentMod, "GetPastMeeting", "Failed to get past meeting"
    ElseIf Response.StatusCode = HTTP_OK Then
        Set GetPastMeeting = Response.Data
        WriteLog 1, CurrentMod, "GetPastMeeting", "Past meeting retrieved successfully"
    Else
        WriteLog 3, CurrentMod, "GetPastMeeting", _
                 "Failed with status: " & Response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetPastMeeting", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetPastMeeting = Nothing
End Function

'=======================================================
' Function: GetPastMeetings
' Purpose: Get list of past instances of a meeting
'
' Parameters:
'   ID - Meeting ID
'
' Returns: Collection of past meeting instances, or Nothing on error
'=======================================================
Public Function GetPastMeetings(ByVal ID As String) As Collection
    Dim Response As WebResponse
    Dim QueryParams As Dictionary
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(ID) = 0 Then
        WriteLog 3, CurrentMod, "GetPastMeetings", "Empty meeting ID"
        Set GetPastMeetings = Nothing
        Exit Function
    End If
    
    WriteLog 1, CurrentMod, "GetPastMeetings", "Getting past meetings for: " & ID
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' Build query parameters
    Set QueryParams = CreateBody(Array("page_size", "type"), Array(100, "past"))
    
    ' Get past meeting instances
    Set Response = UseAPI(QueryParams, HttpGet, DblEncode(ID) & "/instances", _
                         "/past_meetings/", ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Process response
    If Response Is Nothing Then
        WriteLog 3, CurrentMod, "GetPastMeetings", "Failed to get past meetings"
    ElseIf Response.StatusCode = HTTP_OK Then
        If Response.Data.Exists("meetings") Then
            Set GetPastMeetings = Response.Data("meetings")
            WriteLog 1, CurrentMod, "GetPastMeetings", _
                     "Retrieved " & GetPastMeetings.Count & " past meetings"
        Else
            WriteLog 2, CurrentMod, "GetPastMeetings", "No meetings in response"
            Set GetPastMeetings = New Collection
        End If
    Else
        WriteLog 3, CurrentMod, "GetPastMeetings", _
                 "Failed with status: " & Response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetPastMeetings", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetPastMeetings = Nothing
End Function

'=======================================================
' Function: GetZoomTemplates
' Purpose: Get list of available Zoom meeting templates
'
' Returns: Collection of templates, or Nothing on error
'=======================================================
Public Function GetZoomTemplates() As Collection
    Dim Response As WebResponse
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "GetZoomTemplates", "Getting meeting templates"
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' Get templates
    Set Response = UseAPI(Nothing, HttpGet, "meeting_templates", "/users/me/", _
                         ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Process response
    If Response Is Nothing Then
        WriteLog 3, CurrentMod, "GetZoomTemplates", "Failed to get templates"
    ElseIf Response.StatusCode = HTTP_OK Then
        If Response.Data.Exists("templates") Then
            Set GetZoomTemplates = Response.Data("templates")
            WriteLog 1, CurrentMod, "GetZoomTemplates", _
                     "Retrieved " & GetZoomTemplates.Count & " templates"
        Else
            WriteLog 2, CurrentMod, "GetZoomTemplates", "No templates in response"
            Set GetZoomTemplates = New Collection
        End If
    Else
        WriteLog 3, CurrentMod, "GetZoomTemplates", _
                 "Failed with status: " & Response.StatusCode
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetZoomTemplates", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetZoomTemplates = Nothing
End Function

'=======================================================
' Function: GetZoomAttList
' Purpose: Get list of participants who attended a meeting
'
' Parameters:
'   ID - Meeting UUID (double-encoded)
'
' Returns: Collection of participants who were "in_meeting", or Nothing on error
'=======================================================
Public Function GetZoomAttList(ByVal ID As String) As Collection
    Dim Response As WebResponse
    Dim QueryParams As Dictionary
    Dim Coll As Collection
    Dim AColl As Collection
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(ID) = 0 Then
        WriteLog 3, CurrentMod, "GetZoomAttList", "Empty meeting ID"
        Set GetZoomAttList = Nothing
        Exit Function
    End If
    
    WriteLog 1, CurrentMod, "GetZoomAttList", "Getting attendance list for: " & ID
    
    ' Initialize Zoom token if needed
    If ZoomToken Is Nothing Then Set ZoomToken = New BearerToken
    
    ' Build query parameters
    Set QueryParams = CreateBody(Array("page_size"), Array(100))
    
    ' Get participant report
    Set Response = UseAPI(QueryParams, HttpGet, DblEncode(ID) & "/participants", _
                         "/report/meetings/", ZoomAPI, , ZoomToken.Token, , , "Bearer")
    
    ' Process response
    If Response Is Nothing Then
        WriteLog 3, CurrentMod, "GetZoomAttList", "Failed to get attendance list"
        Set GetZoomAttList = Nothing
        Exit Function
    End If
    
    If Response.StatusCode <> HTTP_OK Then
        WriteLog 3, CurrentMod, "GetZoomAttList", _
                 "Failed with status: " & Response.StatusCode
        Set GetZoomAttList = Nothing
        Exit Function
    End If
    
    ' Filter participants who were in meeting
    If Response.Data.Exists("participants") Then
        Set Coll = Response.Data("participants")
        Set AColl = New Collection
        
        For i = 1 To Coll.Count
            If Coll(i).Exists("status") Then
                If Coll(i)("status") = "in_meeting" Then
                    AColl.Add Coll(i)
                End If
            End If
        Next i
        
        Set GetZoomAttList = AColl
        WriteLog 1, CurrentMod, "GetZoomAttList", _
                 "Retrieved " & AColl.Count & " attendees"
    Else
        WriteLog 2, CurrentMod, "GetZoomAttList", "No participants in response"
        Set GetZoomAttList = New Collection
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetZoomAttList", _
             "Error " & Err.Number & ": " & Err.Description
    Set GetZoomAttList = Nothing
End Function

'=======================================================
' PRIVATE HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: DblEncode
' Purpose: Double URL encode a string for Zoom API
'
' Parameters:
'   ID - String to encode
'
' Returns: Double-encoded string
'=======================================================
Private Function DblEncode(ByVal ID As String) As String
    On Error GoTo ErrorHandler
    DblEncode = UrlEncode(UrlEncode(ID))
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "DblEncode", _
             "Error " & Err.Number & ": " & Err.Description
    DblEncode = ID
End Function

'=======================================================
' Function: ValidateCreateMeetingInputs
' Purpose: Validate inputs for CreateZoomMeeting
'
' Parameters:
'   MtgTitle - Meeting title
'   MtgStart - Meeting start time
'   MtgDuration - Meeting duration in minutes
'   errorMsg - Output parameter for error message
'
' Returns: True if valid, False otherwise
'=======================================================
Private Function ValidateCreateMeetingInputs(ByVal MtgTitle As String, _
                                            ByVal MtgStart As Date, _
                                            ByVal MtgDuration As Long, _
                                            ByRef errorMsg As String) As Boolean
    ' Validate title
    If Len(MtgTitle) = 0 Then
        errorMsg = "Meeting title is empty"
        ValidateCreateMeetingInputs = False
        Exit Function
    End If
    
    ' Validate start time
    If MtgStart < Now Then
        errorMsg = "Meeting start time is in the past"
        ' This is a warning, not an error - allow it but log
        WriteLog 2, CurrentMod, "ValidateCreateMeetingInputs", errorMsg
    End If
    
    ' Validate duration
    If MtgDuration <= 0 Then
        errorMsg = "Meeting duration must be greater than 0"
        ValidateCreateMeetingInputs = False
        Exit Function
    End If
    
    If MtgDuration > 1440 Then ' 24 hours
        errorMsg = "Meeting duration exceeds 24 hours"
        ' This is a warning - allow it but log
        WriteLog 2, CurrentMod, "ValidateCreateMeetingInputs", errorMsg
    End If
    
    ValidateCreateMeetingInputs = True
End Function

'=======================================================
' Sub: ApplyDefaultMeetingSettings
' Purpose: Apply default settings to meeting settings dictionary
'
' Parameters:
'   MtgSettings - Dictionary to apply defaults to (modified)
'=======================================================
Private Sub ApplyDefaultMeetingSettings(ByRef MtgSettings As Dictionary)
    On Error GoTo ErrorHandler
    
    ' Email notification settings
    If Not MtgSettings.Exists("registrants_email_notification") Then
        MtgSettings.Add "registrants_email_notification", False
    End If
    
    If Not MtgSettings.Exists("email_notification") Then
        MtgSettings.Add "email_notification", False
    End If
    
    ' Waiting room settings
    If Not MtgSettings.Exists("waiting_room") Then
        MtgSettings.Add "waiting_room", True
    End If
    
    If Not MtgSettings.Exists("waiting_room_options") Then
        Dim WaitingRoomOpts As New Dictionary
        WaitingRoomOpts.Add "mode", "custom"
        WaitingRoomOpts.Add "who_goes_to_waiting_room", "users_not_on_invite"
        MtgSettings.Add "waiting_room_options", WaitingRoomOpts
    End If
    
    ' Host settings
    If Not MtgSettings.Exists("join_before_host") Then
        MtgSettings.Add "join_before_host", True
    End If
    
    ' Meeting summary settings
    If Not MtgSettings.Exists("auto_start_meeting_summary") Then
        MtgSettings.Add "auto_start_meeting_summary", True
    End If
    
    If Not MtgSettings.Exists("who_will_receive_summary") Then
        MtgSettings.Add "who_will_receive_summary", 1
    End If
    
    If Not MtgSettings.Exists("summary_template_id") Then
        MtgSettings.Add "summary_template_id", ZOOM_DEFAULT_SUMMARY_TEMPLATE
    End If
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ApplyDefaultMeetingSettings", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' END OF MODULE
'=======================================================
