Attribute VB_Name = "Globals_UIState"
Option Explicit

'=======================================================
' Module: Globals_UIState
' Purpose: Ribbon state management and caching
' Author: Refactored - November 2025
' Version: 1.0
'
' Description:
'   Central management of Ribbon state, caching, and flags.
'   Separates Ribbon concerns from business logic.
'   Consolidates scattered module-level variables from:
'   - Ribbon_Functions_Mod.bas (ErrShown, cProjectName, etc.)
'   - frmSettings.frm (mHeadingCount, etc.)
'
' Usage:
'   Call InitializeRibbonState on startup
'   Call ClearRibbonCache when project changes
'
' Change Log:
'   v1.0 - Nov 2025
'       * Created from scattered variables
'       * Added cache management
'       * Added state validation
'=======================================================

Private Const CurrentMod As String = "Globals_UIState"

'=======================================================
' Ribbon ERROR STATE
'=======================================================

' Tracks if error has been shown to prevent duplicate dialogs
Public RibbonErrorShown As Boolean

'=======================================================
' RIBBON CACHE (Performance Optimization)
'=======================================================

Private Type RibbonCache
    ProjectName As String
    projectURL As String
    documentName As String
    templateName As String
    LastRefresh As Date
    IsValid As Boolean
End Type

Private ribbonCacheData As RibbonCache

'=======================================================
' FORM STATE CACHE
'=======================================================

Private Type FormStateCache
    headingCount As Long
    bookmarksCount As Long
    boldsCount As Long
    DocType As String
    LastUpdate As Date
    IsValid As Boolean
End Type

Private formStateData As FormStateCache

'=======================================================
' REFRESH FLAGS
'=======================================================

' Flag to indicate ribbon refresh is needed
Public RibbonRefreshFlag As Boolean

'=======================================================
' INITIALIZATION AND CLEANUP
'=======================================================

'=======================================================
' Sub: InitializeRibbonState
' Purpose: Initialize Ribbon state management system
'
' Description:
'   Initializes all Ribbon state variables and clears caches.
'   Should be called on application startup.
'
' Called By:
'   - Application startup routine
'   - After major state changes
'=======================================================
Public Sub InitializeRibbonState()
    On Error GoTo ErrorHandler
    
    RibbonErrorShown = False
    RibbonRefreshFlag = False
    
    Call ClearRibbonCache
    Call ClearFormCache
    
    WriteLog 1, CurrentMod, "InitializeRibbonState", "Ribbon state initialized"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "InitializeRibbonState", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' RIBBON CACHE MANAGEMENT
'=======================================================

'=======================================================
' Sub: ClearRibbonCache
' Purpose: Invalidate ribbon cache forcing refresh
'
' Description:
'   Clears all cached ribbon values and marks cache
'   as invalid. Next access will force a refresh.
'=======================================================
Public Sub ClearRibbonCache()
    On Error GoTo ErrorHandler
    
    With ribbonCacheData
        .ProjectName = ""
        .projectURL = ""
        .documentName = ""
        .templateName = ""
        .LastRefresh = 0
        .IsValid = False
    End With
    
    WriteLog 1, CurrentMod, "ClearRibbonCache", "Ribbon cache cleared"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ClearRibbonCache", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Function: GetCachedProjectName
' Purpose: Get cached project name with validation
'
' Returns:
'   String - Cached project name or empty string if invalid
'=======================================================
Public Function GetCachedProjectName() As String
    On Error GoTo ErrorHandler
    
    If ribbonCacheData.IsValid Then
        GetCachedProjectName = ribbonCacheData.ProjectName
    Else
        GetCachedProjectName = ""
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetCachedProjectName", _
             "Error " & Err.Number & ": " & Err.Description
    GetCachedProjectName = ""
End Function

'=======================================================
' Function: GetCachedProjectURL
' Purpose: Get cached project URL with validation
'=======================================================
Public Function GetCachedProjectURL() As String
    On Error GoTo ErrorHandler
    
    If ribbonCacheData.IsValid Then
        GetCachedProjectURL = ribbonCacheData.projectURL
    Else
        GetCachedProjectURL = ""
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetCachedProjectURL", _
             "Error " & Err.Number & ": " & Err.Description
    GetCachedProjectURL = ""
End Function

'=======================================================
' Function: GetCachedDocumentName
' Purpose: Get cached document name with validation
'=======================================================
Public Function GetCachedDocumentName() As String
    On Error GoTo ErrorHandler
    
    If ribbonCacheData.IsValid Then
        GetCachedDocumentName = ribbonCacheData.documentName
    Else
        GetCachedDocumentName = ""
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetCachedDocumentName", _
             "Error " & Err.Number & ": " & Err.Description
    GetCachedDocumentName = ""
End Function

'=======================================================
' Function: GetCachedTemplateName
' Purpose: Get cached template name with validation
'=======================================================
Public Function GetCachedTemplateName() As String
    On Error GoTo ErrorHandler
    
    If ribbonCacheData.IsValid Then
        GetCachedTemplateName = ribbonCacheData.templateName
    Else
        GetCachedTemplateName = ""
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetCachedTemplateName", _
             "Error " & Err.Number & ": " & Err.Description
    GetCachedTemplateName = ""
End Function

'=======================================================
' Sub: UpdateRibbonCache
' Purpose: Update ribbon cache with current values
'
' Parameters:
'   projectName - Current project name
'   projectURL - Current project URL
'   documentName - Current document name
'   templateName - Current template name
'=======================================================
Public Sub UpdateRibbonCache(ByVal ProjectName As String, _
                            ByVal projectURL As String, _
                            ByVal documentName As String, _
                            ByVal templateName As String)
    On Error GoTo ErrorHandler
    
    With ribbonCacheData
        .ProjectName = ProjectName
        .projectURL = projectURL
        .documentName = documentName
        .templateName = templateName
        .LastRefresh = Now
        .IsValid = True
    End With
    
    WriteLog 1, CurrentMod, "UpdateRibbonCache", "Ribbon cache updated"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "UpdateRibbonCache", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Function: IsRibbonCacheValid
' Purpose: Check if ribbon cache is still valid
'
' Parameters:
'   maxAgeMinutes - Maximum cache age in minutes (default: 5)
'
' Returns:
'   Boolean - True if cache is valid and fresh
'=======================================================
Public Function IsRibbonCacheValid(Optional ByVal maxAgeMinutes As Long = 5) As Boolean
    On Error GoTo ErrorHandler
    
    If Not ribbonCacheData.IsValid Then
        IsRibbonCacheValid = False
        Exit Function
    End If
    
    Dim cacheAge As Double
    cacheAge = DateDiff("n", ribbonCacheData.LastRefresh, Now)
    
    IsRibbonCacheValid = (cacheAge <= maxAgeMinutes)
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "IsRibbonCacheValid", _
             "Error " & Err.Number & ": " & Err.Description
    IsRibbonCacheValid = False
End Function

'=======================================================
' FORM STATE MANAGEMENT
'=======================================================

'=======================================================
' Sub: ClearFormCache
' Purpose: Clear form state cache
'=======================================================
Public Sub ClearFormCache()
    On Error GoTo ErrorHandler
    
    With formStateData
        .headingCount = 0
        .bookmarksCount = 0
        .boldsCount = 0
        .DocType = ""
        .LastUpdate = 0
        .IsValid = False
    End With
    
    WriteLog 1, CurrentMod, "ClearFormCache", "Form cache cleared"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ClearFormCache", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: UpdateFormState
' Purpose: Update form state cache
'
' Parameters:
'   headingCount - Number of headings found
'   bookmarksCount - Number of bookmarks found
'   boldsCount - Number of bold items found
'   docType - Document type string
'=======================================================
Public Sub UpdateFormState(ByVal headingCount As Long, _
                          ByVal bookmarksCount As Long, _
                          ByVal boldsCount As Long, _
                          ByVal DocType As String)
    On Error GoTo ErrorHandler
    
    With formStateData
        .headingCount = headingCount
        .bookmarksCount = bookmarksCount
        .boldsCount = boldsCount
        .DocType = DocType
        .LastUpdate = Now
        .IsValid = True
    End With
    
    WriteLog 1, CurrentMod, "UpdateFormState", "Form state updated"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "UpdateFormState", _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Function: GetFormHeadingCount
' Purpose: Get cached heading count
'=======================================================
Public Function GetFormHeadingCount() As Long
    On Error GoTo ErrorHandler
    
    If formStateData.IsValid Then
        GetFormHeadingCount = formStateData.headingCount
    Else
        GetFormHeadingCount = 0
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetFormHeadingCount", _
             "Error " & Err.Number & ": " & Err.Description
    GetFormHeadingCount = 0
End Function

'=======================================================
' Function: GetFormBookmarksCount
' Purpose: Get cached bookmarks count
'=======================================================
Public Function GetFormBookmarksCount() As Long
    On Error GoTo ErrorHandler
    
    If formStateData.IsValid Then
        GetFormBookmarksCount = formStateData.bookmarksCount
    Else
        GetFormBookmarksCount = 0
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetFormBookmarksCount", _
             "Error " & Err.Number & ": " & Err.Description
    GetFormBookmarksCount = 0
End Function

'=======================================================
' Function: GetFormBoldsCount
' Purpose: Get cached bolds count
'=======================================================
Public Function GetFormBoldsCount() As Long
    On Error GoTo ErrorHandler
    
    If formStateData.IsValid Then
        GetFormBoldsCount = formStateData.boldsCount
    Else
        GetFormBoldsCount = 0
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetFormBoldsCount", _
             "Error " & Err.Number & ": " & Err.Description
    GetFormBoldsCount = 0
End Function

'=======================================================
' Function: GetFormDocType
' Purpose: Get cached document type
'=======================================================
Public Function GetFormDocType() As String
    On Error GoTo ErrorHandler
    
    If formStateData.IsValid Then
        GetFormDocType = formStateData.DocType
    Else
        GetFormDocType = ""
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "GetFormDocType", _
             "Error " & Err.Number & ": " & Err.Description
    GetFormDocType = ""
End Function

'=======================================================
' CLEANUP
'=======================================================

'=======================================================
' Sub: ResetRibbonState
' Purpose: Reset all Ribbon state (call on project close/change)
'
' Description:
'   Resets all Ribbon state variables to initial values.
'   Should be called when:
'   - Project is closed
'   - Project is changed
'   - User logs out
'=======================================================
Public Sub ResetRibbonState()
    On Error GoTo ErrorHandler
    
    RibbonErrorShown = False
    RibbonRefreshFlag = False
    
    Call ClearRibbonCache
    Call ClearFormCache
    
    WriteLog 1, CurrentMod, "ResetRibbonState", "Ribbon state reset"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, "ResetRibbonState", _
             "Error " & Err.Number & ": " & Err.Description
End Sub
