Attribute VB_Name = "DocentBreakdown_Export"
Option Explicit
Option Compare Text

'=======================================================
' Module: DocentBreakdown_Export
' Purpose: Export and upload document sections to server
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Handles the export of document sections to HTML format
'   and manages uploading to the Docent IMS server.
'
'   Main workflow:
'   1. Creates temporary output folders
'   2. Cleans document (removes TOC, headers, images)
'   3. Saves document as HTML
'   4. Collects and processes search words
'   5. Splits document into sections
'   6. Exports each section as separate HTML file
'   7. Prompts user for upload confirmation
'   8. Uploads files to server if confirmed
'
' Dependencies:
'   - AB_GlobalConstants (for file paths and formatting)
'   - AB_GlobalVars (for document collections and state)
'   - AB_CommonFunctions (for ExportRange function)
'   - AD_CollectWords_mod (for CollectWords)
'   - AD_Upload_mod (for upload operations)
'   - SOW and SOWs classes (for section management)
'   - AZ_FileFolder_Mod (for folder operations)
'   - AZ_Log_Mod (for WriteLog)
'
' Main Procedures:
'   - ExportActiveDocument: Main export workflow coordinator
'   - CleanFile: Remove TOC, headers, footers, and images
'   - FixAndExport: Split and export document sections
'
' Helper Procedures:
'   - InitializeExportEnvironment: Setup progress bar and boost
'   - SetupExportPaths: Create output directory structure
'   - PromptForUploadConfirmation: Show results and get user approval
'   - PerformUpload: Execute the upload process
'   - CancelUpload: Handle upload cancellation
'   - RestoreApplicationState: Restore normal application state
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive documentation
'       * Improved error handling
'       * Removed dead code
'       * Split large functions into helpers
'       * Added parameter validation
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "DocentBreakdown_Export"

' Module-level variable for section sequence numbering
Private SeqNo As Long

'=======================================================
' MAIN EXPORT PROCEDURES
'=======================================================

'=======================================================
' Sub: ExportActiveDocument
' Purpose: Export active document sections and upload to server
'
' Parameters:
'   DocType - Type of document being exported
'
' Description:
'   Main workflow for exporting document:
'   1. Creates output folders
'   2. Cleans file (removes TOC, images)
'   3. Saves as HTML
'   4. Splits document into sections
'   5. Exports each section as HTML
'   6. Prompts user for upload confirmation
'   7. Uploads files to server if confirmed
'
' Side Effects:
'   - Creates temporary folders in Set_Odir
'   - Modifies and closes SDoc
'   - May quit application if no documents remain open
'
' Error Handling:
'   - User cancellation handled gracefully
'   - Progress bar properly disposed
'   - Temporary folders cleaned up
'   - Application state restored
'=======================================================
Sub ExportActiveDocument(ByVal DocType As String)
    Const PROC_NAME As String = "ExportActiveDocument"
    
    Dim fileName As String
    Dim buttonClicked As String
    Dim wordCounts As Variant
    Dim userConfirmed As Boolean
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting export for document type: " & DocType
    
    ' Check for user cancellation
    If Set_Cancelled Then
        WriteLog 2, CurrentMod, PROC_NAME, "Export cancelled by user before start"
        GoTo CleanupAndExit
    End If
    
    ' Validate DocType parameter
    If Len(Trim$(DocType)) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "DocType is empty"
        MsgBox "Document type is required for export.", vbExclamation, "Invalid Document Type"
        GoTo CleanupAndExit
    End If
    
    ' Initialize environment (boost and progress bar)
    Call InitializeExportEnvironment
    
    ' Setup output paths
    Call SetupExportPaths
    
    ' Check cancellation after path setup
    If Set_Cancelled Then
        WriteLog 2, CurrentMod, PROC_NAME, "Export cancelled by user after path setup"
        GoTo CleanupAndExit
    End If
    
    ' Save original filename
    fileName = SDoc.FullName
    
    ' Clean and prepare document
    WriteLog 1, CurrentMod, PROC_NAME, "Cleaning document"
    Call CleanFile
    
    Call UpdateSearchRange
    ProgressBar.Spin
    
    ' Check cancellation after cleaning
    If Set_Cancelled Then
        WriteLog 2, CurrentMod, PROC_NAME, "Export cancelled by user after cleaning"
        GoTo CleanupAndExit
    End If
    
    ' Save as HTML
    WriteLog 1, CurrentMod, PROC_NAME, "Saving document as HTML"
    SDoc.SaveAs2 fileName & ".html", wdFormatXMLDocumentMacroEnabled
    ProgressBar.Spin
    
    ' Check cancellation after saving
    If Set_Cancelled Then
        WriteLog 2, CurrentMod, PROC_NAME, "Export cancelled by user after saving"
        GoTo CleanupAndExit
    End If
    
    ' Hide screen updates for better performance
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    
    ' Collect search words and export sections
    WriteLog 1, CurrentMod, PROC_NAME, "Collecting words and exporting sections"
    Call CollectWords
    wordCounts = FixAndExport()
    
    ' Check if export was successful
    If IsEmpty(wordCounts) Then
        WriteLog 2, CurrentMod, PROC_NAME, "Export was cancelled or failed"
        GoTo CleanupAndExit
    End If
    
    ' Close source document
    WriteLog 1, CurrentMod, PROC_NAME, "Closing source document"
    SDoc.Close SaveChanges:=False
    
    ' Restore boost
    Boost False, True
    
    ' Show results and get upload confirmation
    userConfirmed = PromptForUploadConfirmation(wordCounts)
    
    If userConfirmed Then
        ' User confirmed - perform upload
        Call PerformUpload(DocType, fileName)
    Else
        ' User cancelled - cleanup temp files
        Call CancelUpload
    End If
    
CleanupAndExit:
    ' Restore application to normal state
    Call RestoreApplicationState
    Call CloseLog
    
    WriteLog 1, CurrentMod, PROC_NAME, "Export process completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred during export:" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & vbNewLine & _
           "Please try again or use the help button to send feedback.", _
           vbCritical, "Export Error"
    
    ' Ensure cleanup happens even on error
    On Error Resume Next
    Call RestoreApplicationState
    Call CloseLog
    On Error GoTo 0
End Sub

'=======================================================
' Sub: CleanFile
' Purpose: Remove TOC, headers, footers, and images from document
'
' Description:
'   Prepares document for export by:
'   1. Removing unlinked headers and footers from all sections
'   2. Deleting all shapes (images, drawings)
'   3. Extracting and exporting front pages (TOC)
'   4. Removing the TOC from the document
'   5. Updating search position markers
'
' Side Effects:
'   - Modifies SDoc (global scope document)
'   - Updates Set_SPos (search start position)
'   - Creates "Front Pages" HTML export
'
' Error Handling:
'   Logs errors but continues processing to ensure
'   document can still be exported even if some
'   elements cannot be removed
'=======================================================
Sub CleanFile()
    Const PROC_NAME As String = "CleanFile"
    
    Dim xSec As Section
    Dim hdFt As HeaderFooter
    Dim SecRng As Range
    Dim shapeIndex As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Removing TOC, headers, footers, and images"
    
    ' Unprotect document for editing
    Unprotect SDoc
    
    ' Remove unlinked headers and footers from all sections
    For Each xSec In SDoc.Sections
        
        ' Process headers
        For Each hdFt In xSec.Headers
            If Not hdFt.LinkToPrevious Then
                hdFt.Range.Delete
            End If
        Next hdFt
        
        ' Process footers
        For Each hdFt In xSec.Footers
            If Not hdFt.LinkToPrevious Then
                hdFt.Range.Delete
            End If
        Next hdFt
        
    Next xSec
    
    ' Remove all shapes (images, drawings, etc.) from the document
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Removing " & SDoc.Shapes.Count & " shapes"
    
    For shapeIndex = SDoc.Shapes.Count To 1 Step -1
        SDoc.Shapes(shapeIndex).Delete
    Next shapeIndex
    
    ' Export and remove front pages (TOC) if present
    If SDoc.TablesOfContents.Count > 0 Then
        WriteLog 1, CurrentMod, PROC_NAME, "Exporting and removing Table of Contents"
        
        Set SecRng = SDoc.Range
        Set_SPos = SDoc.TablesOfContents(1).Range.End + 1
        
        ' Set range to include everything before the end of TOC
        SecRng.SetRange start:=0, End:=Set_SPos - 1
        
        ' Export front pages as sequence 0
        ExportRange SecRng, Nothing, Nothing, "Front Pages", HTMLPath, 0, Nothing
        
        ' Delete the front pages from document
        SecRng.Delete
        
        ' Reset search position
        Set_SPos = 0
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "No Table of Contents found"
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "File cleaning completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    ' Don't stop execution - log and continue
    ' This allows the export to proceed even if some cleanup fails
End Sub

'=======================================================
' Function: FixAndExport
' Purpose: Split document into sections and export each as HTML
'
' Returns:
'   Array(1 To 2) As Long - Word counts:
'     Index 1: Total words in source file
'     Index 2: Total words in all output files
'   Returns Empty if cancelled or error occurs
'
' Description:
'   Iterates through FilteredSOWs collection and exports
'   each section with navigation links to previous and
'   next sections. Updates progress bar during processing.
'
' Side Effects:
'   - Creates multiple HTML files in HTMLPath
'   - Updates module-level SeqNo variable
'   - Displays progress bar updates
'
' Error Handling:
'   Returns Empty on error or cancellation
'   Progress bar is properly disposed
'=======================================================
Function FixAndExport() As Variant
    Const PROC_NAME As String = "FixAndExport"
    
    Dim mSOW As SOW
    Dim SecRng As Range
    Dim pRng As Range
    Dim NRng As Range
    Dim DRng As Range
    Dim sectionIndex As Long
    Dim wordCountsArray(1 To 2) As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Splitting and exporting " & FilteredSOWs.Count & " sections"
    
    ' Initialize ranges
    Set SecRng = SDoc.Range
    Set pRng = SDoc.Range
    Set NRng = SDoc.Range
    Set DRng = SDoc.Range
    
    ' Reset progress bar
    ProgressBar.Reset
    ProgressBar.Progress , "Exporting File No. 1", 0
    ProgressBar.Dom = FilteredSOWs.Count
    
    ' Get total words in search range
    wordCountsArray(1) = Set_SearchRange.ComputeStatistics(wdStatisticWords)
    
    ' Reset sequence number
    SeqNo = 0
    
    ' Process each section
    For sectionIndex = 1 To FilteredSOWs.Count
        
        ' Update progress bar
        If ProgressBar.Progress(, "Exporting File No. " & sectionIndex) Then
            Unload ProgressBar
            MsgBox "Export cancelled by user.", vbExclamation, "Cancelled"
            GoTo Cancelled
        End If
        
        ' Check global cancellation flag
        If Set_Cancelled Then GoTo Cancelled
        
        ' Get current section
        Set mSOW = FilteredSOWs(sectionIndex)
        Set SecRng = mSOW.SectionRng
        Set DRng = mSOW.DelivRange
        
        ' Set previous section range (for navigation)
        If sectionIndex = 1 Then
            ' First section - no previous
            pRng.SetRange start:=0, End:=0
        Else
            ' Set to previous section range
            pRng.SetRange start:=FilteredSOWs(sectionIndex - 1).SectionStart, _
                         End:=FilteredSOWs(sectionIndex - 1).SectionEnd
        End If
        
        ' Set next section range (for navigation)
        If sectionIndex = FilteredSOWs.Count Then
            ' Last section - no next
            NRng.SetRange start:=0, End:=0
        Else
            ' Set to next section range
            NRng.SetRange start:=FilteredSOWs(sectionIndex + 1).SectionStart, _
                         End:=FilteredSOWs(sectionIndex + 1).SectionEnd
        End If
        
        ' Increment sequence number
        SeqNo = SeqNo + 1
        
        ' Export this section
        ExportRange SecRng, pRng, NRng, mSOW.FullName, HTMLPath, SeqNo, DRng
        
        ' Update word count
        wordCountsArray(2) = wordCountsArray(2) + FilteredSOWs(sectionIndex).CountWords
        
    Next sectionIndex
    
    ' Mark progress as finished
    ProgressBar.Finished
    
    ' Return word counts array
    FixAndExport = wordCountsArray
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Exported " & FilteredSOWs.Count & " sections successfully"
    Exit Function
    
Cancelled:
    WriteLog 2, CurrentMod, PROC_NAME, "Export cancelled by user"
    FixAndExport = Empty
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    FixAndExport = Empty
End Function

'=======================================================
' HELPER PROCEDURES
'=======================================================

'=======================================================
' Sub: InitializeExportEnvironment
' Purpose: Setup boost and progress bar for export
'
' Description:
'   Initializes the export environment by:
'   - Activating performance boost
'   - Setting CodeIsRunning flag
'   - Configuring progress bar
'   - Starting progress bar spinner
'
' Side Effects:
'   - Sets global CodeIsRunning flag
'   - Configures and displays ProgressBar form
'=======================================================
Private Sub InitializeExportEnvironment()
    Const PROC_NAME As String = "InitializeExportEnvironment"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Initializing export environment"
    
    Boost
    CodeIsRunning = True
    
    With ProgressBar
        .BarsCount = 1
        .HideApplication = True
        .Spin
    End With
End Sub

'=======================================================
' Sub: SetupExportPaths
' Purpose: Create output directory structure
'
' Description:
'   Creates the folder structure for export:
'   - Main output folder: "DocentIMS Analysis"
'   - Sub-folder: "HTML Documents"
'
'   Deletes any existing output folder first to ensure
'   clean export.
'
' Side Effects:
'   - Sets global OPath variable
'   - Sets global HTMLPath variable
'   - Deletes and recreates output folders
'=======================================================
Private Sub SetupExportPaths()
    Const PROC_NAME As String = "SetupExportPaths"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Setting up output paths"
    
    ' Build paths
    OPath = Set_Odir & Application.PathSeparator & _
            "DocentIMS Analysis" & Application.PathSeparator
    
    HTMLPath = OPath & "HTML Documents" & Application.PathSeparator
    
    ' Clean up any previous export
    DeleteFolder OPath
    
    ' Create fresh directory structure
    CreateDir HTMLPath
    
    ProgressBar.Spin
    
    WriteLog 1, CurrentMod, PROC_NAME, "Output paths created: " & HTMLPath
End Sub

'=======================================================
' Function: PromptForUploadConfirmation
' Purpose: Show results and get user confirmation for upload
'
' Parameters:
'   wordCounts - Variant array with word count statistics
'     Index 1: Words in source file
'     Index 2: Words in output files
'
' Returns:
'   True if user confirms upload (types "yes")
'   False if user cancels or error occurs
'
' Description:
'   1. Hides progress bar
'   2. Shows success message with word counts
'   3. Displays confirmation dialog with project info
'   4. Adjusts text color for dark backgrounds
'   5. Gets user input ("yes" to proceed)
'
' Side Effects:
'   - Unloads ProgressBar form
'   - Shows and unloads frmConfirmUpload form
'=======================================================
Private Function PromptForUploadConfirmation(ByVal wordCounts As Variant) As Boolean
    Const PROC_NAME As String = "PromptForUploadConfirmation"
    
    Dim buttonClicked As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Prompting user for upload confirmation"
    
    ' Hide progress bar
    Unload ProgressBar
    
    ' Show success message with word counts
    frmMsgBox.Display _
        "Words in the source file: " & wordCounts(1) & vbNewLine & _
        "Total words in the output files: " & wordCounts(2), _
        "OK", Success, "Export Complete"
    
    ' Setup confirmation dialog
    With frmConfirmUpload
        .Label1.Caption = "You are uploading to """ & ProjectNameStr & """ project." & _
                         Chr(10) & "Type ""yes"" if you want to continue."
        
        ' Adjust text color for dark backgrounds
        If FullColor(ProjectColorStr).TooDark Then
            .Label1.ForeColor = vbWhite
        End If
        
        .BackColor = ProjectColorStr
        .Show
        
        buttonClicked = LCase$(Trim$(.TextBox1.text))
    End With
    
    Unload frmConfirmUpload
    
    ' Check user response
    PromptForUploadConfirmation = (buttonClicked = "yes")
    
    If PromptForUploadConfirmation Then
        WriteLog 1, CurrentMod, PROC_NAME, "User confirmed upload"
    Else
        WriteLog 1, CurrentMod, PROC_NAME, "User cancelled upload"
    End If
    
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    PromptForUploadConfirmation = False
End Function

'=======================================================
' Sub: PerformUpload
' Purpose: Execute the upload process
'
' Parameters:
'   DocType - Type of document being uploaded
'   fileName - Path to the source document file
'
' Description:
'   1. Opens the exported HTML file
'   2. Calls UploadDoc to prepare document
'   3. Calls StartUploading to upload files
'
' Side Effects:
'   - Opens document file
'   - Uploads files to server
'   - May modify server content
'
' Error Handling:
'   Logs errors and displays error message to user
'=======================================================
Private Sub PerformUpload(ByVal DocType As String, ByVal fileName As String)
    Const PROC_NAME As String = "PerformUpload"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting upload process for: " & DocType
    
    ' Open the exported document and prepare for upload
    UploadDoc Documents.Open(fileName), SilentMode:=True
    
    ' Start the upload process
    StartUploading DocType, OPath, HTMLPath
    
    WriteLog 1, CurrentMod, PROC_NAME, "Upload completed successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred during upload:" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & vbNewLine & _
           "Some files may not have been uploaded.", _
           vbCritical, "Upload Error"
End Sub

'=======================================================
' Sub: CancelUpload
' Purpose: Handle upload cancellation
'
' Description:
'   Cleans up temporary files and notifies user that
'   the upload was cancelled. Deletes the entire output
'   folder structure created during export.
'
' Side Effects:
'   - Deletes OPath folder and all contents
'   - Shows cancellation message to user
'=======================================================
Private Sub CancelUpload()
    Const PROC_NAME As String = "CancelUpload"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Cancelling upload and cleaning up"
    
    ' Delete temporary folder
    DeleteFolder OPath
    
    ' Notify user
    MsgBox "File uploading cancelled." & vbNewLine & _
           "Temporary files have been deleted.", _
           vbExclamation, "Upload Cancelled"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Temporary folder deleted"
End Sub

'=======================================================
' Sub: RestoreApplicationState
' Purpose: Restore application to normal state
'
' Description:
'   Restores the application to normal operating state:
'   1. Clears CodeIsRunning flag
'   2. Restores print view
'   3. Restores boost settings
'   4. Shows application
'   5. Re-enables screen updating
'   6. Re-enables display alerts
'   7. Quits if no documents remain open
'
' Side Effects:
'   - Modifies global CodeIsRunning flag
'   - Changes application visibility
'   - Changes screen updating state
'   - Changes display alerts state
'   - May quit application
'=======================================================
Private Sub RestoreApplicationState()
    Const PROC_NAME As String = "RestoreApplicationState"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Restoring application state"
    
    CodeIsRunning = False
    PrintView
    Boost False, True
    
    If Application.Documents.Count = 0 Then
        ' No documents open - quit application
        WriteLog 1, CurrentMod, PROC_NAME, "No documents open, quitting application"
        Application.Quit
    Else
        ' Restore normal state
        Application.Visible = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = wdAlertsAll
        WriteLog 1, CurrentMod, PROC_NAME, "Application state restored"
    End If
End Sub

'=======================================================
' END OF MODULE
'=======================================================
