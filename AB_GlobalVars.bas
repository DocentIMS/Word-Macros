Attribute VB_Name = "AB_GlobalVars"
Option Explicit

'=======================================================
' Module: AB_GlobalVars
' Purpose: Global variables used throughout the application
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Central repository for all application-wide variables.
'   Constants have been moved to AB_GlobalConstants2.
'   Procedures have been moved to appropriate functional modules.
'
' Organization:
'   - Ribbon and UI state variables
'   - Settings form variables
'   - Workflow and authorization variables
'   - Document and project variables
'   - User and authentication variables
'   - Dictionary and collection variables
'
' Change Log:
'   v2.0 - Nov 2025
'       * Removed constants (moved to AB_GlobalConstants2)
'       * Removed procedures (moved to AB_SearchHelpers)
'       * Added module documentation
'       * Organized by category
'       * Removed commented dead code
'   v1.0 - Original version
'=======================================================

'=======================================================
' RIBBON AND UI STATE VARIABLES
'=======================================================

' Indicates if project manager features are enabled
Public PrjMgr As Boolean

' Ribbon busy state flag
Public BusyRibbon As Boolean

' Flag indicating code is currently running
Public CodeIsRunning As Boolean

'=======================================================
' DICTIONARY VARIABLES
'=======================================================

' Path to custom dictionary file
Public DocentDictionaryPath As String

' Custom Word dictionary object
Public DocentDictionary As Word.Dictionary

'=======================================================
' SETTINGS FORM VARIABLES
'=======================================================
' Variables used by the settings form for document processing

' Use bookmarks in document processing
Public Set_UseBookmarks As Boolean

' Apply coloring to document elements
Public Set_Coloring As Boolean

' Apply indenting to document elements
Public Set_Indenting As Boolean

' Export processed documents
Public Set_Export As Boolean

' User cancelled the operation
Public Set_Cancelled As Boolean

' Apply bold formatting in addition to other formatting
Public Set_BoldToo As Boolean

' Output directory for exports
Public Set_Odir As String

' Search mode selector
Public Set_SearchMode As Long

' Test limit for processing (0 = no limit)
Public Set_TestLimit As Long

' End position for search range
Public Set_EPos As Long

' Start position for search range
Public Set_SPos As Long

' Active search range in document
Public Set_SearchRange As Range

' Collection of search words
Public Wrds As SearchWords

'=======================================================
' PARSING AND PROCESSING FLAGS
'=======================================================

' Will/Shall parsing mode
Public mWillShall As Boolean

' General parse mode flag
Public mParse As Boolean

'=======================================================
' WORKFLOW AND AUTHORIZATION VARIABLES
'=======================================================

' User is authorized for operations
Public IsAuthorized As Boolean

' Prompt for missing passwords
Public AskForMissingPw As Boolean

' Planning-only mode flag
Public PlanningOnly As Boolean

'=======================================================
' DOCUMENT INFO VARIABLES
'=======================================================

' Document info for currently opening document
Public OpeningDocInfo As DocInfo

' Collection of all document info objects
Public DocsInfo As Collection

'=======================================================
' URL VARIABLES
'=======================================================

' Scope manager URL
Public ScopeURL As String

' RFP manager URL
Public RFPURL As String

' PMP manager URL
Public PMPURL As String

' Dashboard URL
Public DashboardURLStr As String

'=======================================================
' PARSE STATUS FLAGS
'=======================================================

' Scope document has been parsed
Public ScopeParsed As Boolean

' RFP document has been parsed
Public RFPParsed As Boolean

' PMP document has been parsed
Public PMPParsed As Boolean

'=======================================================
' SOW (STATEMENT OF WORK) VARIABLES
'=======================================================

' Collection of all SOWs
Public AllSOWs As SOWs

' Filtered collection of SOWs
Public FilteredSOWs As SOWs

' Active scope document
Public SDoc As Document

' Collection of SOWs by document
Public SOWsColl As Collection

' Collection of items not asked about scope
Public NotScopeAsked As Collection

'=======================================================
' USER INFORMATION ARRAYS
'=======================================================

' Array of user names
Public UserName() As String

' Array of user IDs
Public UserID() As String

'=======================================================
' PROJECT INFORMATION ARRAYS
'=======================================================

' Array of project URLs
Public projectURL() As String

' Array of project names
Public projectName() As String

' Array of planning project names
Public PlanningProjectName() As String

' Array of non-planning project names
Public NoPlanningProjectName() As String

' Array of planning project URLs
Public PlanningProjectURL() As String

' Array of non-planning project URLs
Public NoPlanningProjectURL() As String

' Array of project passwords
Public Password() As String

' Array of project colors
Public ProjectColor() As String

' Array of parse uploaded status
Public ParseUploaded() As String

' Array of project clients
Public ProjectClient() As String

' Array of project contract numbers
Public ProjectContractNumber() As String

' Array of project planning flags
Public ProjectIsPlanning() As String

'=======================================================
' DOCUMENT TYPE ARRAYS
'=======================================================

' Array of document type names
Public documentName() As String

' Array of meeting document names
Public MeetingDocName() As String

' Array of manager document names
Public ManagerDocName() As String

' Array of template names
Public templateName() As String

'=======================================================
' WORKFLOW AND STATE DICTIONARIES
'=======================================================

' Collection of available workflow transitions
Public NextTransitions As Collection

' Workflow information dictionary
Public WorkflowInfo As New Dictionary

' Document types dictionary
Public DocumentsTypes As Dictionary

'=======================================================
' PROJECT AND USER DICTIONARIES
'=======================================================

' Main project information
Public MainInfo As New Dictionary

' Project configuration information
Public ProjectInfo As New Dictionary

' User Plone roles dictionary
Public UserPloneRolesDict As Dictionary

' User team roles dictionary
Public UserTeamRolesDict As Dictionary

' Project members dictionary
Public MembersDict As Dictionary

' Project groups dictionary
Public ProjectGroupsDict As Dictionary

' User groups dictionary
Public UserGroupsDict As Dictionary

'=======================================================
' TIMEZONE VARIABLES
'=======================================================

' Plone server timezone offset
Public PloneTimeZone As Single

' Local machine timezone offset
Public LocalTimeZone As Single

'=======================================================
' CURRENT PROJECT STRING VARIABLES
'=======================================================

' Current user name string
Public UserNameStr As String

' Current user ID string
Public UserIDStr As String

' Current project URL string
Public ProjectURLStr As String

' Current contract number string
Public ContractNumberStr As String

' Current project planning status string
Public ProjectIsPlanningStr As String

' Current project VS name string
Public ProjectVSNameStr As String

' Current project name string
Public ProjectNameStr As String

' Document naming convention string
Public DocumentsNameConvStr As String

' Current user password string
Public UserPasswordStr As String

' Template password string
Public TemplatePasswordStr As String

' Current project color string
Public ProjectColorStr As String

' Current project client string
Public ProjectClientStr As String

' Parse uploaded status string
Public PUploaded As String

'=======================================================
' PROJECT AND DOCUMENT NUMBERS
'=======================================================

' Current project number
Public PNum As Long

' New project number
Public NewPNum As Long

' Current document number
Public DocNum As Long

' Current meeting document number
Public MeetingDocNum As Long

' Current template number
Public TemplateNum As Long

'=======================================================
' PATH VARIABLES
'=======================================================

' Output path for exports
Public OPath As String

' HTML export path
Public HTMLPath As String

' Installation path for add-in
Public InstallationPath As String

' Images path
Public ImagesPath As String

'=======================================================
' TASK AND NOTIFICATION DICTIONARIES
'=======================================================

' Tasks dictionary
Public TasksDict As Dictionary

' Notifications dictionary
Public NotifsDict As Dictionary

'=======================================================
' ZOOM INTEGRATION VARIABLES
'=======================================================

' Zoom API bearer token
Public ZoomToken As BearerToken

'=======================================================
' END OF MODULE
'=======================================================
