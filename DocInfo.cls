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
Public UpdateMode As Boolean

Public Name As String
'Public URL As String
'Public ProjectName As String
'Public ContractNo As String
'Public DocType As String
'Public DocState As String
'Public Version As Long
'Public TemplateVersion As Long
'Public isDocentDocument As Boolean
'Public isDocentTemplate As Boolean
'Public PublishingDate As String
'Public CreationDate As String
'Public MeetingType As String

Public DocLastSave As String '= -2 ' "Last Save Time"
Public Author As String '= -1 ' "Last Author"
Public IsDocument As String '= 0 '"isDocentDocument"
Public PName As String '= 1 ' "docentProject"
Public PURL As String '= 1 ' "docentProjectURL"
Public DocType As String '= 2 ' "docentDocType"
Public DocVer As String '= 3 ' "docentVersion"
Public DocDate As String '= 4 ' "docentDocDate"
Public DocState As String '= 5 '"docentDocState"
Public DocURL As String '= 6 '"docentDocURL"
Public DocCreateDate As String '= 7 '"docentCreationDate"
Public IsTemplate As String '= 8 ' "isDocentTemplate"
Public IsFinalRev As String '= 8 ' "isDocentTemplate"
Public TemplateVer As String '= 9 ' "docentTemplateVersion"
Public TemplateDate As String '= 10 ' "docentTemplateDate"
Public publishDate As String '= 11 ' "docentPublishingDate"
Public ProposedTasks As String '= 12 '"ProposedTasks
Public PlannedTasks As String '= 13 '"PlannedTasks
Public ContractNo As String '= 14 '"docentContractNo
Public MeetingType As String '= 15 '"docentMeetingType"
Public Actuals As String '= 16 '"docentMeetingActuals"
Public MeetingURL As String '= 17 '"docentMeetingURL"
Public OnlineMeetingURL As String '= 17 '"onlineMeetingURL"

