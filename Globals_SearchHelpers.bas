Attribute VB_Name = "Globals_SearchHelpers"
Option Explicit

'=======================================================
' Module: Globals_SearchHelpers
' Purpose: Helper functions for document search operations
' Author: Extracted from AB_GlobalVars - November 2025
' Version: 1.0
'
' Description:
'   This module contains helper procedures for managing
'   search ranges in documents. Previously these were
'   in AB_GlobalVars but have been extracted for better
'   organization.
'
' Dependencies:
'   - AB_GlobalVars (Set_TestLimit, Set_SearchRange, etc.)
'   - AB_GlobalConstants2 (uses constants from there)
'
' Functions:
'   - FindSearchRange: Initialize search range for document
'   - UpdateSearchRange: Update search range boundaries
'   - ResetSetGlobals: Reset all search-related globals
'
' Change Log:
'   v1.0 - Nov 2025
'       * Extracted from AB_GlobalVars
'       * Added error handling
'       * Added documentation
'       * Improved code organization
'=======================================================

Private Const CurrentMod As String = "Globals_SearchHelpers"

'=======================================================
' Sub: FindSearchRange
' Purpose: Find and set the search range in a document
'
' Description:
'   Sets up the search range for document processing.
'   If Set_TestLimit is specified, limits the range to
'   that many pages. Excludes table of contents from
'   the search range.
'
' Global Variables Used:
'   - SDoc: The document to search
'   - Set_TestLimit: Number of pages to limit search (0 = no limit)
'   - Set_SearchRange: Output - the configured search range
'   - Set_EPos: Output - end position of range
'   - Set_SPos: Output - start position of range
'
' Example:
'   Set SDoc = ActiveDocument
'   Set_TestLimit = 5  ' Limit to first 5 pages
'   Call FindSearchRange
'   ' Now Set_SearchRange is configured
'=======================================================
Public Sub FindSearchRange()
    Dim i As Long
    Dim TocEndPos As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "FindSearchRange", "Finding search range"
    
    ' Validate document exists
    If SDoc Is Nothing Then
        WriteLog 3, CurrentMod, "FindSearchRange", "SDoc is Nothing"
        Exit Sub
    End If
    
    ' Start with full document range
    Set Set_SearchRange = SDoc.Range
    
    ' Apply page limit if specified
    If Set_TestLimit > 0 Then
        WriteLog 1, CurrentMod, "FindSearchRange", "Limiting to " & Set_TestLimit & " pages"
        
        For i = 1 To Set_TestLimit
            Set_SearchRange = Set_SearchRange.GoTo(What:=wdGoToBookmark, Name:="\page")
            Set_SearchRange.Move 1
        Next i
    End If
    
    ' Set end position
    Set_EPos = Set_SearchRange.End
    
    ' Set start position (after TOC if present)
    Set_SPos = 0
    
    ' Try to find TOC end
    On Error Resume Next
    TocEndPos = SDoc.TablesOfContents(1).Range.End
    On Error GoTo ErrorHandler
    
    If TocEndPos > 0 Then
        Set_SPos = TocEndPos + 1
        WriteLog 1, CurrentMod, "FindSearchRange", "TOC found, starting after it at position " & Set_SPos
    Else
        WriteLog 1, CurrentMod, "FindSearchRange", "No TOC found, starting at beginning"
    End If
    
    ' Configure the search range
    Set_SearchRange.SetRange Set_SPos, Set_EPos
    
    WriteLog 1, CurrentMod, "FindSearchRange", _
             "Search range set: " & Set_SPos & " to " & Set_EPos
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "FindSearchRange", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: UpdateSearchRange
' Purpose: Update search range boundaries from current range
'
' Description:
'   Updates the global search position variables based on
'   the current Set_SearchRange. Used after modifying
'   the search range.
'
' Global Variables Used:
'   - Set_SearchRange: Input - current search range
'   - Set_SPos: Output - updated start position
'   - Set_EPos: Output - updated end position
'
' Example:
'   ' After modifying Set_SearchRange
'   Set_SearchRange.Collapse wdCollapseEnd
'   Call UpdateSearchRange  ' Update positions
'=======================================================
Public Sub UpdateSearchRange()
    On Error GoTo ErrorHandler
    
    ' Validate search range exists
    If Set_SearchRange Is Nothing Then
        WriteLog 3, CurrentMod, "UpdateSearchRange", "Set_SearchRange is Nothing"
        Exit Sub
    End If
    
    ' Update positions from current range
    Set_SPos = Set_SearchRange.start
    Set_EPos = Set_SearchRange.End
    
    WriteLog 1, CurrentMod, "UpdateSearchRange", _
             "Updated range: " & Set_SPos & " to " & Set_EPos
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "UpdateSearchRange", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: ResetSetGlobals
' Purpose: Reset all search-related global variables
'
' Description:
'   Initializes/resets all global variables related to
'   document searching and processing. Should be called
'   before starting a new search operation.
'
' Global Variables Reset:
'   - All Set_* variables
'   - mWillShall, mParse
'   - SOWsColl, AllSOWs
'   - SDoc (set to ActiveDocument)
'
' Example:
'   Call ResetSetGlobals
'   ' Now ready for new search
'   Set_UseBookmarks = True
'   Call FindSearchRange
'=======================================================
Public Sub ResetSetGlobals()
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, "ResetSetGlobals", "Resetting search globals"
    
    ' Reset boolean flags
    mWillShall = False
    mParse = False
    Set_UseBookmarks = False
    Set_Coloring = False
    Set_Indenting = False
    Set_Export = False
    Set_Cancelled = False
    Set_BoldToo = False
    
    ' Reset string values
    Set_Odir = vbNullString
    
    ' Reset numeric values
    Set_SearchMode = 0
    Set_TestLimit = 0
    Set_EPos = 0
    Set_SPos = 0
    
    ' Set active document
    If Documents.Count > 0 Then
        Set SDoc = ActiveDocument
        WriteLog 1, CurrentMod, "ResetSetGlobals", "SDoc set to: " & SDoc.Name
    Else
        WriteLog 2, CurrentMod, "ResetSetGlobals", "No active document"
        Set SDoc = Nothing
    End If
    
    ' Initialize search range
    If Not SDoc Is Nothing Then
        Call FindSearchRange
    End If
    
    ' Reset SOW collections
    Set SOWsColl = Nothing
    Set SOWsColl = New Collection
    
    ' Try to restore existing SOW for this document
    If Not SDoc Is Nothing Then
        On Error Resume Next
        Set AllSOWs = SOWsColl(SDoc.Name)
        
        ' If not found, create new
        If Err.Number <> 0 Then
            Err.Clear
            Set AllSOWs = New SOWs
            SOWsColl.Add AllSOWs, SDoc.Name
            WriteLog 1, CurrentMod, "ResetSetGlobals", "Created new SOWs collection"
        Else
            WriteLog 1, CurrentMod, "ResetSetGlobals", "Restored existing SOWs collection"
        End If
        On Error GoTo ErrorHandler
    End If
    
    WriteLog 1, CurrentMod, "ResetSetGlobals", "Reset complete"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ResetSetGlobals", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' ADDITIONAL HELPER FUNCTIONS
'=======================================================

'=======================================================
' Function: GetSearchRangeInfo
' Purpose: Get information about current search range
'
' Returns: String with search range details
'
' Example:
'   Debug.Print GetSearchRangeInfo()
'   ' Output: "Search Range: 150-5000 (4850 chars), TestLimit: 0"
'=======================================================
Public Function GetSearchRangeInfo() As String
    Dim info As String
    
    On Error GoTo ErrorHandler
    
    If Set_SearchRange Is Nothing Then
        GetSearchRangeInfo = "Search Range: Not initialized"
        Exit Function
    End If
    
    info = "Search Range: " & Set_SPos & "-" & Set_EPos
    info = info & " (" & (Set_EPos - Set_SPos) & " chars)"
    
    If Set_TestLimit > 0 Then
        info = info & ", TestLimit: " & Set_TestLimit & " pages"
    Else
        info = info & ", TestLimit: None"
    End If
    
    GetSearchRangeInfo = info
    Exit Function
    
ErrorHandler:
    GetSearchRangeInfo = "Error getting search range info: " & Err.Description
End Function

'=======================================================
' Function: IsSearchRangeValid
' Purpose: Check if search range is properly initialized
'
' Returns: True if search range is valid and ready to use
'
' Example:
'   If IsSearchRangeValid() Then
'       ' Proceed with search
'   Else
'       Call ResetSetGlobals
'   End If
'=======================================================
Public Function IsSearchRangeValid() As Boolean
    On Error GoTo ErrorHandler
    
    IsSearchRangeValid = False
    
    ' Check if range exists
    If Set_SearchRange Is Nothing Then
        WriteLog 2, CurrentMod, "IsSearchRangeValid", "Set_SearchRange is Nothing"
        Exit Function
    End If
    
    ' Check if document exists
    If SDoc Is Nothing Then
        WriteLog 2, CurrentMod, "IsSearchRangeValid", "SDoc is Nothing"
        Exit Function
    End If
    
    ' Check if positions are valid
    If Set_EPos <= Set_SPos Then
        WriteLog 2, CurrentMod, "IsSearchRangeValid", _
                 "Invalid positions: Start=" & Set_SPos & ", End=" & Set_EPos
        Exit Function
    End If
    
    ' Check if positions are within document bounds
    If Set_SPos < 0 Or Set_EPos > SDoc.Characters.Count Then
        WriteLog 2, CurrentMod, "IsSearchRangeValid", _
                 "Positions out of document bounds"
        Exit Function
    End If
    
    IsSearchRangeValid = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "IsSearchRangeValid", _
             "Error " & Err.Number & ": " & Err.Description
    IsSearchRangeValid = False
End Function

'=======================================================
' END OF MODULE
'=======================================================
