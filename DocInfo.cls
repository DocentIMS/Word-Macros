VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'=======================================================
' Class: DocInfo
' Purpose: Store and manage document metadata and properties
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Encapsulates metadata for Docent IMS documents including:
'   - Project information (name, URL, contract number)
'   - Document type and state
'   - Version information (document and template)
'   - Dates (creation, last save, publishing)
'   - Author information
'   - Task information (proposed, planned)
'   - Meeting information (type, URL, actuals)
'
' Usage Example:
'   Dim docInfo As New DocInfo
'   docInfo.Name = "Requirements Document"
'   docInfo.PName = "Project Alpha"
'   docInfo.DocType = "Scope Document"
'   docInfo.DocVer = "1"
'
'   If docInfo.IsValid Then
'       ' Use document info
'   End If
'
' Properties:
'   All properties exposed via Property Get/Let for validation
'   and encapsulation. Property Let procedures trim input values.
'
' Methods:
'   - IsValid: Validates required properties are set
'   - ToString: Returns formatted string representation
'   - Clear: Resets all properties to default values
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive class documentation
'       * Converted all public variables to Property procedures
'       * Removed all commented dead code
'       * Added validation methods
'       * Added initialization and cleanup
'   v1.0 - Original version
'=======================================================

'=======================================================
' PRIVATE MEMBERS
'=======================================================

' Mode and state flags
Private m_UpdateMode As Boolean
Private m_IsDocument As String
Private m_IsTemplate As String
Private m_IsFinalRev As String

' Identity
Private m_Name As String
Private m_Author As String

' Project information
Private m_PName As String
Private m_PURL As String
Private m_ContractNo As String

' Document information
Private m_DocType As String
Private m_DocVer As String
Private m_DocState As String
Private m_DocURL As String

' Dates
Private m_DocLastSave As String
Private m_DocDate As String
Private m_DocCreateDate As String
Private m_publishDate As String

' Template information
Private m_TemplateVer As String
Private m_TemplateDate As String

' Task information
Private m_ProposedTasks As String
Private m_PlannedTasks As String

' Meeting information
Private m_meetingType As String
Private m_Actuals As String
Private m_MeetingURL As String
Private m_OnlineMeetingURL As String

'=======================================================
' PROPERTIES - MODE AND STATE
'=======================================================

'=======================================================
' Property: UpdateMode
' Purpose: Indicates if document is in update mode
'
' Type: Boolean
' Default: False
'=======================================================
Public Property Get UpdateMode() As Boolean
    UpdateMode = m_UpdateMode
End Property

Public Property Let UpdateMode(ByVal value As Boolean)
    m_UpdateMode = value
End Property

'=======================================================
' Property: IsDocument
' Purpose: Indicates if this is a Docent document
'
' Type: String (typically "True" or "False")
' Default: Empty string
'=======================================================
Public Property Get IsDocument() As String
    IsDocument = m_IsDocument
End Property

Public Property Let IsDocument(ByVal value As String)
    m_IsDocument = Trim$(value)
End Property

'=======================================================
' Property: IsTemplate
' Purpose: Indicates if this is a Docent template
'
' Type: String (typically "True" or "False")
' Default: Empty string
'=======================================================
Public Property Get IsTemplate() As String
    IsTemplate = m_IsTemplate
End Property

Public Property Let IsTemplate(ByVal value As String)
    m_IsTemplate = Trim$(value)
End Property

'=======================================================
' Property: IsFinalRev
' Purpose: Indicates if this is the final revision
'
' Type: String (typically "True" or "False")
' Default: Empty string
'=======================================================
Public Property Get IsFinalRev() As String
    IsFinalRev = m_IsFinalRev
End Property

Public Property Let IsFinalRev(ByVal value As String)
    m_IsFinalRev = Trim$(value)
End Property

'=======================================================
' PROPERTIES - IDENTITY
'=======================================================

'=======================================================
' Property: Name
' Purpose: Document name/title
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(ByVal value As String)
    m_Name = Trim$(value)
End Property

'=======================================================
' Property: Author
' Purpose: Last author of the document
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get Author() As String
    Author = m_Author
End Property

Public Property Let Author(ByVal value As String)
    m_Author = Trim$(value)
End Property

'=======================================================
' PROPERTIES - PROJECT INFORMATION
'=======================================================

'=======================================================
' Property: PName
' Purpose: Project name
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get PName() As String
    PName = m_PName
End Property

Public Property Let PName(ByVal value As String)
    m_PName = Trim$(value)
End Property

'=======================================================
' Property: PURL
' Purpose: Project URL on Docent IMS server
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get PURL() As String
    PURL = m_PURL
End Property

Public Property Let PURL(ByVal value As String)
    m_PURL = Trim$(value)
End Property

'=======================================================
' Property: ContractNo
' Purpose: Contract number associated with project
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get ContractNo() As String
    ContractNo = m_ContractNo
End Property

Public Property Let ContractNo(ByVal value As String)
    m_ContractNo = Trim$(value)
End Property

'=======================================================
' PROPERTIES - DOCUMENT INFORMATION
'=======================================================

'=======================================================
' Property: DocType
' Purpose: Type of document (e.g., "Scope Document", "RFP")
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get DocType() As String
    DocType = m_DocType
End Property

Public Property Let DocType(ByVal value As String)
    m_DocType = Trim$(value)
End Property

'=======================================================
' Property: DocVer
' Purpose: Document version number
'
' Type: String
' Default: "1"
'=======================================================
Public Property Get DocVer() As String
    DocVer = m_DocVer
End Property

Public Property Let DocVer(ByVal value As String)
    m_DocVer = Trim$(value)
End Property

'=======================================================
' Property: DocState
' Purpose: Current state of document in workflow
'
' Type: String (e.g., "Draft", "In Review", "Approved")
' Default: Empty string
'=======================================================
Public Property Get DocState() As String
    DocState = m_DocState
End Property

Public Property Let DocState(ByVal value As String)
    m_DocState = Trim$(value)
End Property

'=======================================================
' Property: DocURL
' Purpose: URL of document on Docent IMS server
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get DocURL() As String
    DocURL = m_DocURL
End Property

Public Property Let DocURL(ByVal value As String)
    m_DocURL = Trim$(value)
End Property

'=======================================================
' PROPERTIES - DATES
'=======================================================

'=======================================================
' Property: DocLastSave
' Purpose: Last save date/time of document
'
' Type: String (formatted date)
' Default: Empty string
'=======================================================
Public Property Get DocLastSave() As String
    DocLastSave = m_DocLastSave
End Property

Public Property Let DocLastSave(ByVal value As String)
    m_DocLastSave = Trim$(value)
End Property

'=======================================================
' Property: DocDate
' Purpose: Document date (publication or effective date)
'
' Type: String (formatted date)
' Default: Empty string
'=======================================================
Public Property Get DocDate() As String
    DocDate = m_DocDate
End Property

Public Property Let DocDate(ByVal value As String)
    m_DocDate = Trim$(value)
End Property

'=======================================================
' Property: DocCreateDate
' Purpose: Document creation date
'
' Type: String (formatted date)
' Default: Empty string
'=======================================================
Public Property Get DocCreateDate() As String
    DocCreateDate = m_DocCreateDate
End Property

Public Property Let DocCreateDate(ByVal value As String)
    m_DocCreateDate = Trim$(value)
End Property

'=======================================================
' Property: publishDate
' Purpose: Document publishing date on server
'
' Type: String (formatted date)
' Default: Empty string
'=======================================================
Public Property Get publishDate() As String
    publishDate = m_publishDate
End Property

Public Property Let publishDate(ByVal value As String)
    m_publishDate = Trim$(value)
End Property

'=======================================================
' PROPERTIES - TEMPLATE INFORMATION
'=======================================================

'=======================================================
' Property: TemplateVer
' Purpose: Version of template used for document
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get TemplateVer() As String
    TemplateVer = m_TemplateVer
End Property

Public Property Let TemplateVer(ByVal value As String)
    m_TemplateVer = Trim$(value)
End Property

'=======================================================
' Property: TemplateDate
' Purpose: Date of template used
'
' Type: String (formatted date)
' Default: Empty string
'=======================================================
Public Property Get TemplateDate() As String
    TemplateDate = m_TemplateDate
End Property

Public Property Let TemplateDate(ByVal value As String)
    m_TemplateDate = Trim$(value)
End Property

'=======================================================
' PROPERTIES - TASK INFORMATION
'=======================================================

'=======================================================
' Property: ProposedTasks
' Purpose: List of proposed tasks (typically JSON or delimited)
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get ProposedTasks() As String
    ProposedTasks = m_ProposedTasks
End Property

Public Property Let ProposedTasks(ByVal value As String)
    m_ProposedTasks = Trim$(value)
End Property

'=======================================================
' Property: PlannedTasks
' Purpose: List of planned tasks (typically JSON or delimited)
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get PlannedTasks() As String
    PlannedTasks = m_PlannedTasks
End Property

Public Property Let PlannedTasks(ByVal value As String)
    m_PlannedTasks = Trim$(value)
End Property

'=======================================================
' PROPERTIES - MEETING INFORMATION
'=======================================================

'=======================================================
' Property: meetingType
' Purpose: Type of meeting document represents
'
' Type: String (e.g., "Status Meeting", "Review Meeting")
' Default: Empty string
'=======================================================
Public Property Get meetingType() As String
    meetingType = m_meetingType
End Property

Public Property Let meetingType(ByVal value As String)
    m_meetingType = Trim$(value)
End Property

'=======================================================
' Property: Actuals
' Purpose: Actual meeting information or outcomes
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get Actuals() As String
    Actuals = m_Actuals
End Property

Public Property Let Actuals(ByVal value As String)
    m_Actuals = Trim$(value)
End Property

'=======================================================
' Property: MeetingURL
' Purpose: URL of meeting document on server
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get MeetingURL() As String
    MeetingURL = m_MeetingURL
End Property

Public Property Let MeetingURL(ByVal value As String)
    m_MeetingURL = Trim$(value)
End Property

'=======================================================
' Property: OnlineMeetingURL
' Purpose: URL for online meeting (Zoom, Teams, etc.)
'
' Type: String
' Default: Empty string
'=======================================================
Public Property Get OnlineMeetingURL() As String
    OnlineMeetingURL = m_OnlineMeetingURL
End Property

Public Property Let OnlineMeetingURL(ByVal value As String)
    m_OnlineMeetingURL = Trim$(value)
End Property

'=======================================================
' CLASS METHODS
'=======================================================

'=======================================================
' Function: IsValid
' Purpose: Validate that required properties are set
'
' Returns:
'   True if required properties have values, False otherwise
'
' Description:
'   Checks that minimum required properties are populated:
'   - Name (document name)
'   - DocType (document type)
'   - PName (project name)
'
' Usage:
'   If docInfo.IsValid Then
'       ' Proceed with document operations
'   Else
'       MsgBox "Required document information is missing"
'   End If
'=======================================================
Public Function IsValid() As Boolean
    IsValid = (Len(m_Name) > 0) And _
              (Len(m_DocType) > 0) And _
              (Len(m_PName) > 0)
End Function

'=======================================================
' Function: ToString
' Purpose: Return string representation of DocInfo
'
' Returns:
'   Formatted string with key properties
'
' Description:
'   Provides a human-readable summary of the document info
'   for logging and debugging purposes.
'
' Usage:
'   Debug.Print docInfo.ToString
'   WriteLog 1, "Module", "Proc", docInfo.ToString
'=======================================================
Public Function ToString() As String
    ToString = "DocInfo: " & m_Name & _
               " | Type: " & m_DocType & _
               " | Project: " & m_PName & _
               " | Ver: " & m_DocVer & _
               " | State: " & m_DocState
End Function

'=======================================================
' Sub: Clear
' Purpose: Reset all properties to default values
'
' Description:
'   Clears all property values, resetting the object
'   to its initial state. Useful for reusing an existing
'   DocInfo object for a different document.
'
' Usage:
'   docInfo.Clear
'   docInfo.Name = "New Document"
'   ' ... set other properties
'=======================================================
Public Sub Clear()
    ' Mode and state
    m_UpdateMode = False
    m_IsDocument = ""
    m_IsTemplate = ""
    m_IsFinalRev = ""
    
    ' Identity
    m_Name = ""
    m_Author = ""
    
    ' Project information
    m_PName = ""
    m_PURL = ""
    m_ContractNo = ""
    
    ' Document information
    m_DocType = ""
    m_DocVer = "1"
    m_DocState = ""
    m_DocURL = ""
    
    ' Dates
    m_DocLastSave = ""
    m_DocDate = ""
    m_DocCreateDate = ""
    m_publishDate = ""
    
    ' Template information
    m_TemplateVer = ""
    m_TemplateDate = ""
    
    ' Task information
    m_ProposedTasks = ""
    m_PlannedTasks = ""
    
    ' Meeting information
    m_meetingType = ""
    m_Actuals = ""
    m_MeetingURL = ""
    m_OnlineMeetingURL = ""
End Sub

'=======================================================
' CLASS EVENTS
'=======================================================

'=======================================================
' Sub: Class_Initialize
' Purpose: Initialize default values for new instances
'
' Description:
'   Called automatically when a new DocInfo object is created.
'   Sets default values for all properties.
'=======================================================
Private Sub Class_Initialize()
    ' Set default values
    m_UpdateMode = False
    m_IsDocument = ""
    m_IsTemplate = ""
    m_IsFinalRev = ""
    
    ' Default version to 1
    m_DocVer = "1"
    
    ' All other properties default to empty string
End Sub

'=======================================================
' Sub: Class_Terminate
' Purpose: Cleanup when object is destroyed
'
' Description:
'   Called automatically when the DocInfo object goes out
'   of scope or is explicitly set to Nothing. Currently
'   no cleanup is needed, but provided for completeness.
'=======================================================
Private Sub Class_Terminate()
    ' No cleanup required for string properties
End Sub

'=======================================================
' END OF CLASS MODULE
'=======================================================
