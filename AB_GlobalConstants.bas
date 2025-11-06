Attribute VB_Name = "AB_GlobalConstants"
Option Explicit

'=======================================================
' Module: AB_GlobalConstants
' Purpose: Global constants used throughout the application
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Central repository for all application-wide constants.
'   Organized by category for easy maintenance.
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added HTTP status code constants
'       * Added documentation headers
'       * Organized constants by category
'   v1.0 - Original version
'=======================================================

'=======================================================
' UI CONSTANTS
'=======================================================

' Number of lines to highlight in document
Public Const HighlightLines As Long = 10

' Maximum length for file names
Public Const FileNameLimit As Long = 40

' Error highlight color (light red)
Public Const ErrorColor As Long = &H8080FF

'=======================================================
' DOCUMENT FORMATTING CONSTANTS
'=======================================================

' Break keyword used in document processing
Public Const BRKwrd As String = "zxzxzxz"

' HTML horizontal rule with gradient styling
Public Const HRTag As String = "<hr style=""height: 3px;" & _
            " margin-top: 0px; margin-bottom: 0px; color: #c06856; background-color: #c06856;" & _
            " background-image: linear-gradient(to right, #ccc, #c06856, #ccc);"" />"

'=======================================================
' DICTIONARY CONSTANTS
'=======================================================

' Docent IMS custom dictionary file name
Public Const DocentDictionaryName As String = "DocentIMS.dic"

'=======================================================
' FOLDER PATH CONSTANTS
'=======================================================

' Default server folder paths
Public Const DefaultDocumentsFolder As String = "documents"
Public Const DefaultMeetingsFolder As String = "meetings"
Public Const DefaultScopeFolder As String = "scope-manager"
Public Const DefaultRFPFolder As String = "rfp-manager"
Public Const DefaultPMPFolder As String = "pmp-manager"
Public Const DefaultPlanningFolder As String = "planning-documents-manager"

'=======================================================
' DATE/TIME FORMAT CONSTANTS
'=======================================================

' Various date and time formatting strings
Public Const DateFormat As String = "m/d/yyyy"
Public Const TimeFormat As String = "h:mm AM/PM"
Public Const DateTimeFormat As String = "m/d/yyyy - h:mm AM/PM"
Public Const APIDateTimeFormat As String = "m/d/yyyy h:nn:ss AM/PM"
Public Const LongDateFormat As String = "mmm d, yyyy"
Public Const LongDateTimeFormat As String = "mmm d, yyyy (h:mm AM/PM)"

'=======================================================
' API CONSTANTS
'=======================================================

' Prefix for file references in API calls
Public Const APIFilePrefix As String = "DocentFile:"

'=======================================================
' FILE EXTENSION CONSTANTS
'=======================================================

' Default image file extension
Public Const ImagesExtension As String = ".jpg"

'=======================================================
' HELP SYSTEM CONSTANTS
'=======================================================

' Help documentation URL
Public Const HelpURL As String = "https://help.docentims.com/help/" '"docent-help/webhelp/index.html"

' Number of help types available
Public Const HelpTypesCount As Long = 4

'=======================================================
' HTTP STATUS CODE CONSTANTS
'=======================================================
' Standard HTTP status codes for API communication
' Reference: https://developer.mozilla.org/en-US/docs/Web/HTTP/Status
'=======================================================

' 2xx Success Status Codes
'-------------------------------------------------------
' The request was successfully received, understood, and accepted

' 200 OK - Request succeeded
Public Const HTTP_OK As Long = 200

' 201 Created - Request succeeded and new resource was created
Public Const HTTP_CREATED As Long = 201

' 202 Accepted - Request accepted for processing but not completed
Public Const HTTP_ACCEPTED As Long = 202

' 204 No Content - Request succeeded but no content to return
Public Const HTTP_NO_CONTENT As Long = 204

' 3xx Redirection Status Codes
'-------------------------------------------------------
' Further action needs to be taken to complete the request

' 301 Moved Permanently - Resource has been permanently moved
Public Const HTTP_MOVED_PERMANENTLY As Long = 301

' 302 Found - Resource temporarily moved
Public Const HTTP_FOUND As Long = 302

' 304 Not Modified - Resource has not been modified
Public Const HTTP_NOT_MODIFIED As Long = 304

' 4xx Client Error Status Codes
'-------------------------------------------------------
' The request contains bad syntax or cannot be fulfilled

' 400 Bad Request - Server cannot process request due to client error
Public Const HTTP_BAD_REQUEST As Long = 400

' 401 Unauthorized - Authentication is required and has failed or not been provided
Public Const HTTP_UNAUTHORIZED As Long = 401

' 403 Forbidden - Server refuses to authorize the request
Public Const HTTP_FORBIDDEN As Long = 403

' 404 Not Found - Server cannot find the requested resource
Public Const HTTP_NOT_FOUND As Long = 404

' 405 Method Not Allowed - Request method not supported for resource
Public Const HTTP_METHOD_NOT_ALLOWED As Long = 405

' 408 Request Timeout - Server timed out waiting for request
Public Const HTTP_TIMEOUT As Long = 408

' 409 Conflict - Request conflicts with current state of server
Public Const HTTP_CONFLICT As Long = 409

' 410 Gone - Resource is no longer available
Public Const HTTP_GONE As Long = 410

' 422 Unprocessable Entity - Request well-formed but unable to process
Public Const HTTP_UNPROCESSABLE_ENTITY As Long = 422

' 429 Too Many Requests - User has sent too many requests
Public Const HTTP_TOO_MANY_REQUESTS As Long = 429

' 5xx Server Error Status Codes
'-------------------------------------------------------
' The server failed to fulfill an apparently valid request

' 500 Internal Server Error - Server encountered unexpected condition
Public Const HTTP_SERVER_ERROR As Long = 500

' 501 Not Implemented - Server does not support functionality required
Public Const HTTP_NOT_IMPLEMENTED As Long = 501

' 502 Bad Gateway - Server received invalid response from upstream server
Public Const HTTP_BAD_GATEWAY As Long = 502

' 503 Service Unavailable - Server is not ready to handle request
Public Const HTTP_SERVICE_UNAVAILABLE As Long = 503

' 504 Gateway Timeout - Server did not get timely response from upstream server
Public Const HTTP_GATEWAY_TIMEOUT As Long = 504

'=======================================================
' DOCUMENT PROCESSING CONSTANTS
'=======================================================

' Header level for document hierarchy (used in TOC and navigation)
Public Const HLevel As Long = 4

' Separator string for delimited data (used in arrays/collections conversion)
Public Const Seperator As String = "|||"

' Duration in milliseconds for brief pauses/delays in processing
Public Const NapDuration As Long = 10

'=======================================================
' END OF MODULE
'=======================================================

