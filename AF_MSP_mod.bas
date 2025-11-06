Attribute VB_Name = "AF_MSP_mod"
Option Explicit
Option Private Module

'=======================================================
' Module: AF_MSP_mod
' Purpose: Microsoft Project integration and document generation
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Handles MS Project Excel file import and conversion to
'   Word documents. Manages synchronization between Excel
'   project files and Word documents, including upload/download
'   operations and PDF generation.
'
' Dependencies:
'   - AC_API_Mod (GetAPIContent, GetAPIFolder, CreateAPIContent, UpdateAPIContent)
'   - AC_API_Mod (UploadAPIFile, DownloadAPIFile)
'   - AC_Properties (GetProperty, SetProperty, SetContentControl)
'   - AB_CommonFunctions (Protect, Unprotect, SaveForUpload)
'   - AB_GlobalConstants
'   - AB_GlobalVars
'   - Excel.Application (external dependency)
'   - ProgressBar form
'
' Public Interface:
'   - ImportMSP() - Main entry point for MS Project import
'   - UpdateMSP(SilentMode, xlDt, wrdDt) - Update/generate MS Project document
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive error handling
'       * Added input validation
'       * Added detailed logging
'       * Added resource cleanup
'       * Added function documentation
'       * Improved Excel interop safety
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "AF_MSP_mod"
Private Const MSPTemplatePath As String = "/templates/manager-templates/MS Project.dotx"
Private Const MSPOutputPath As String = "/ms-project/"

'=======================================================
' Sub: ImportMSP
' Purpose: Import MS Project Excel file and generate Word document
'
' Description:
'   Main entry point for MS Project import. Sets busy flags,
'   calls UpdateMSP to perform the actual work, and resets flags.
'
' Error Handling:
'   - Ensures flags are reset even on error
'   - Logs all errors
'   - Displays user-friendly error message
'=======================================================
Sub ImportMSP()
    Const PROC_NAME As String = "ImportMSP"
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting MS Project import"
    
    ' Set busy flags
    BusyRibbon = True
    CodeIsRunning = True
    
    ' Perform import
    Call UpdateMSP(SilentMode:=True)
    
    WriteLog 1, CurrentMod, PROC_NAME, "MS Project import completed successfully"
    
Cleanup:
    ' Always reset flags
    CodeIsRunning = False
    BusyRibbon = False
    Exit Sub
    
ErrorHandler:
    Dim errorMsg As String
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "MS Project import failed: " & errorMsg, vbCritical, "Error"
    Resume Cleanup
End Sub

'=======================================================
' Function: GetXlDate
' Purpose: Get modification date of Excel file from API
'
' Returns:
'   Date - Excel file modification date in server time
'   Returns #12:00:00 AM# if file not found or error
'
' Description:
'   Retrieves the "modified" timestamp from the Excel file's
'   API metadata. Returns midnight (zero date) on error.
'
' Error Handling:
'   - Handles API call failures
'   - Handles missing or invalid date
'   - Logs errors
'=======================================================
Function GetXlDate() As Date
    Const PROC_NAME As String = "GetXlDate"
    
    Dim response As WebResponse
    Dim dateStr As String
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Getting Excel file date"
    
    ' Get file metadata from API
    Set response = GetAPIContent(MSPOutputPath & "ms-project.xlsx")
    
    If Not IsGoodResponse(response) Then
        errorMsg = "API call failed or file not found"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        GetXlDate = #12:00:00 AM#
        Exit Function
    End If
    
    ' Extract and parse date
    On Error Resume Next
    dateStr = CStr(response.Data("modified"))
    If Err.Number <> 0 Or Len(dateStr) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Modified date not found in response"
        GetXlDate = #12:00:00 AM#
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    GetXlDate = TimeFromTFormat(dateStr)
    
    WriteLog 1, CurrentMod, PROC_NAME, "Excel file date: " & GetXlDate
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    GetXlDate = #12:00:00 AM#
End Function

'=======================================================
' Function: GetWrdDate
' Purpose: Get date of last Word document generation
'
' Returns:
'   Date - Word document generation date
'   Returns #12:00:00 AM# if document not found or error
'
' Description:
'   Retrieves the stored date from the ms_project content
'   type in the API. This indicates when the Word document
'   was last generated.
'
' Error Handling:
'   - Handles API call failures
'   - Handles missing date field
'   - Handles invalid date format
'   - Logs errors
'=======================================================
Function GetWrdDate() As Date
    Const PROC_NAME As String = "GetWrdDate"
    
    Dim response As Collection
    Dim dateStr As String
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Getting Word document date"
    
    ' Get ms_project items from API
    Set response = GetAPIFolder(MSPOutputPath, "ms_project", Array("date"))
    
    If Not IsGoodResponse(response) Then
        WriteLog 2, CurrentMod, PROC_NAME, "No ms_project items found"
        GetWrdDate = #12:00:00 AM#
        Exit Function
    End If
    
    If response.Count = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "ms_project collection is empty"
        GetWrdDate = #12:00:00 AM#
        Exit Function
    End If
    
    ' Extract date from first item
    On Error Resume Next
    dateStr = CStr(response(1)("date"))
    If Err.Number <> 0 Or Len(dateStr) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Date field not found or empty"
        GetWrdDate = #12:00:00 AM#
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Parse date
    GetWrdDate = ParseDate(dateStr, APIDateTimeFormat)
    
    WriteLog 1, CurrentMod, PROC_NAME, "Word document date: " & GetWrdDate
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    GetWrdDate = #12:00:00 AM#
End Function

'=======================================================
' Function: UpdateMSP
' Purpose: Update or generate MS Project Word document
'
' Parameters:
'   SilentMode - If True, open existing doc if up-to-date (default: False)
'   xlDt - Excel file date (optional, retrieved if #12:00:00 AM#)
'   wrdDt - Word doc date (optional, retrieved if #12:00:00 AM#)
'
' Returns:
'   Boolean - True if update was needed and performed, False otherwise
'
' Description:
'   Compares Excel file date with Word document date.
'   If Excel is newer, generates new Word document from Excel data.
'   If Word is up-to-date, optionally opens existing document.
'
' Error Handling:
'   - Validates Excel file exists
'   - Handles document generation errors
'   - Handles upload errors
'   - Cleans up temporary files
'   - Logs all steps
'=======================================================
Function UpdateMSP(Optional ByVal SilentMode As Boolean = False, _
                   Optional ByVal xlDt As Date = #12:00:00 AM#, _
                   Optional ByVal wrdDt As Date = #12:00:00 AM#) As Boolean
    Const PROC_NAME As String = "UpdateMSP"
    
    Dim newDoc As Document
    Dim wordPath As String
    Dim pdfPath As String
    Dim needsUpdate As Boolean
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Checking MS Project update status"
    
    ' Get Excel file date if not provided
    If xlDt = #12:00:00 AM# Then
        xlDt = GetXlDate()
    End If
    
    ' Validate Excel file exists
    If xlDt = #12:00:00 AM# Then
        errorMsg = "Excel file is missing or cannot be accessed"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        frmMsgBox.Display "Excel file is missing. Contact the project manager.", , Critical
        UpdateMSP = False
        Exit Function
    End If
    
    ' Get Word document date if not provided
    If wrdDt = #12:00:00 AM# Then
        wrdDt = GetWrdDate()
    End If
    
    ' Compare dates (strip seconds for comparison)
    needsUpdate = (wrdDt < DateSerial(Year(xlDt), Month(xlDt), Day(xlDt)) + _
                           TimeSerial(Hour(xlDt), Minute(xlDt), 0))
    
    If needsUpdate Then
        WriteLog 1, CurrentMod, PROC_NAME, "Update needed. Excel: " & xlDt & ", Word: " & wrdDt
        
        ' Generate new document from Excel
        Set newDoc = CreateNewMSP(MSPOutputPath & "ms-project.xlsx", xlDt)
        
        If newDoc Is Nothing Then
            errorMsg = "Failed to create MS Project document"
            WriteLog 3, CurrentMod, PROC_NAME, errorMsg
            MsgBox errorMsg, vbCritical, "Error"
            UpdateMSP = False
            Exit Function
        End If
        
        ' Save document
        wordPath = SaveForUpload("MS Project", newDoc)
        If Len(wordPath) = 0 Then
            errorMsg = "Failed to save Word document"
            WriteLog 3, CurrentMod, PROC_NAME, errorMsg
            MsgBox errorMsg, vbCritical, "Error"
            UpdateMSP = False
            Exit Function
        End If
        
        ' Generate PDF
        pdfPath = Left$(wordPath, InStrRev(wordPath, ".")) & "pdf"
        
        On Error Resume Next
        newDoc.ExportAsFixedFormat2 pdfPath, wdExportFormatPDF, False, _
                                    wdExportOptimizeForOnScreen, BitmapMissingFonts:=False
        If Err.Number <> 0 Then
            WriteLog 3, CurrentMod, PROC_NAME, "PDF export failed: " & Err.Description
            ' Continue - PDF is optional
        Else
            WriteLog 1, CurrentMod, PROC_NAME, "PDF created: " & pdfPath
            
            ' Upload PDF
            UploadAPIFile pdfPath, MSPOutputPath, Overwrite:=True
            If Err.Number <> 0 Then
                WriteLog 3, CurrentMod, PROC_NAME, "PDF upload failed: " & Err.Description
            End If
        End If
        On Error GoTo ErrorHandler
        
        ' Upload or update API content
        If wrdDt = #12:00:00 AM# Then
            ' Create new content
            WriteLog 1, CurrentMod, PROC_NAME, "Creating new ms_project content"
            CreateAPIContent "ms_project", MSPOutputPath, _
                           Array("file", "date", "default_view"), _
                           Array(wordPath, Format$(xlDt, APIDateTimeFormat), "@@display-file")
        Else
            ' Update existing content
            WriteLog 1, CurrentMod, PROC_NAME, "Updating existing ms_project content"
            UpdateAPIContent MSPOutputPath & "ms-project.docx", _
                           Array("file", "date"), _
                           Array(wordPath, Format$(xlDt, APIDateTimeFormat))
        End If
        
        UpdateMSP = True
        
    Else
        ' Already up-to-date
        WriteLog 1, CurrentMod, PROC_NAME, "MS Project is up-to-date"
        
        If SilentMode Then
            ' Open existing document
            WriteLog 1, CurrentMod, PROC_NAME, "Opening existing MS Project document"
            
            Set OpeningDocInfo = New DocInfo
            With OpeningDocInfo
                .DocType = "MS Project"
                .ContractNo = ContractNumberStr
                .IsDocument = True
                .PName = ProjectNameStr
                .PURL = ProjectURLStr
                .DocCreateDate = Format$(ToServerTime, DateTimeFormat)
            End With
            
            On Error Resume Next
            Documents.Open DownloadAPIFile(MSPOutputPath & "ms-project.docx")
            If Err.Number <> 0 Then
                WriteLog 3, CurrentMod, PROC_NAME, "Failed to open document: " & Err.Description
                MsgBox "Failed to open MS Project document: " & Err.Description, vbExclamation, "Error"
            End If
            On Error GoTo ErrorHandler
        Else
            MsgBox "You are already using the most recent MS Project Excel file", vbInformation, "Up To Date"
        End If
        
        UpdateMSP = False
    End If
    
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    MsgBox "MS Project update failed: " & errorMsg, vbCritical, "Error"
    UpdateMSP = False
End Function

'=======================================================
' Function: CreateNewMSP
' Purpose: Create new MS Project Word document from Excel data
'
' Parameters:
'   ExcelPth - Path to Excel file on API
'   xlDt - Excel file date to embed in document
'
' Returns:
'   Document - Generated Word document, or Nothing on error
'
' Description:
'   Downloads Excel file, extracts data, creates Word document
'   from template, populates table with Excel data, and
'   formats the document. Includes progress bar for user feedback.
'
' Error Handling:
'   - Validates Excel file download
'   - Handles Excel interop errors
'   - Cleans up Excel instance
'   - Cleans up temporary files
'   - Handles user cancellation
'   - Logs all steps
'=======================================================
Function CreateNewMSP(ByVal ExcelPth As String, ByVal xlDt As Date) As Document
    Const PROC_NAME As String = "CreateNewMSP"
    
    Dim excelApp As Object
    Dim workbook As Object
    Dim excelWasOpen As Boolean
    Dim excelData As Variant
    Dim localExcelPath As String
    Dim newDoc As Document
    Dim localWordPath As String
    Dim rowIndex As Long
    Dim dataRow As Long
    Dim colIndex As Long
    Dim errorMsg As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Creating new MS Project document"
    
    ' Download Excel file
    WriteLog 1, CurrentMod, PROC_NAME, "Downloading Excel file"
    localExcelPath = DownloadAPIFile(ExcelPth)
    
    If Len(localExcelPath) = 0 Or Len(Dir$(localExcelPath)) = 0 Then
        errorMsg = "Failed to download Excel file"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    
    ' Open Excel and extract data
    WriteLog 1, CurrentMod, PROC_NAME, "Opening Excel file"
    
    On Error Resume Next
    Set excelApp = GetObject(, "Excel.Application")
    excelWasOpen = Not excelApp Is Nothing
    
    If Not excelWasOpen Then
        Set excelApp = CreateObject("Excel.Application")
        If excelApp Is Nothing Then
            WriteLog 3, CurrentMod, PROC_NAME, "Failed to create Excel application"
            MsgBox "Microsoft Excel is required but could not be started.", vbCritical, "Error"
            Kill localExcelPath
            Set CreateNewMSP = Nothing
            Exit Function
        End If
    End If
    On Error GoTo ErrorHandler
    
    ' Configure Excel
    excelApp.EnableEvents = False
    excelApp.ScreenUpdating = False
    excelApp.Visible = False
    excelApp.DisplayAlerts = False
    
    ' Open workbook and get data
    On Error Resume Next
    Set workbook = excelApp.Workbooks.Open(localExcelPath)
    If Err.Number <> 0 Or workbook Is Nothing Then
        errorMsg = "Failed to open Excel workbook: " & Err.Description
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        Call CleanupExcel(excelApp, Nothing, excelWasOpen)
        Kill localExcelPath
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    
    excelData = workbook.Sheets(1).UsedRange.value
    If Err.Number <> 0 Then
        errorMsg = "Failed to read Excel data: " & Err.Description
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        Call CleanupExcel(excelApp, workbook, excelWasOpen)
        Kill localExcelPath
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Close Excel
    Call CleanupExcel(excelApp, workbook, excelWasOpen)
    
    ' Delete temporary Excel file
    On Error Resume Next
    Kill localExcelPath
    On Error GoTo ErrorHandler
    
    ' Validate data
    If IsEmpty(excelData) Then
        errorMsg = "Excel data is empty"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    
    ' Create Word document from template
    WriteLog 1, CurrentMod, PROC_NAME, "Creating Word document from template"
    
    ' Set document info
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .DocType = "MS Project"
        .ContractNo = ContractNumberStr
        .IsDocument = True
        .PName = ProjectNameStr
        .PURL = ProjectURLStr
    End With
    
    ' Download template
    localWordPath = DownloadAPIFile(MSPTemplatePath)
    
    If Len(localWordPath) = 0 Or Len(Dir$(localWordPath)) = 0 Then
        errorMsg = "Failed to download Word template"
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    
    ' Create document
    On Error Resume Next
    Set newDoc = Documents.Add(localWordPath)
    If Err.Number <> 0 Or newDoc Is Nothing Then
        errorMsg = "Failed to create document from template: " & Err.Description
        WriteLog 3, CurrentMod, PROC_NAME, errorMsg
        Kill localWordPath
        MsgBox errorMsg, vbCritical, "Error"
        Set CreateNewMSP = Nothing
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Unprotect document for editing
    Call Unprotect(newDoc)
    
    ' Set header content controls
    On Error Resume Next
    SetContentControl "MSPDateTime", Format$(ToServerTime(CStr(xlDt)), DateTimeFormat), newDoc
    SetContentControl "Project Name", ProjectNameStr, newDoc
    On Error GoTo ErrorHandler
    
    ' Populate table with Excel data
    WriteLog 1, CurrentMod, PROC_NAME, "Populating table with " & UBound(excelData) & " rows"
    
    rowIndex = 1
    
    ' Show progress bar
    On Error Resume Next
    ProgressBar.BarsCount = 1
    ProgressBar.Reset
    ProgressBar.Progress , "Generating MS Project Word file...", 0
    ProgressBar.Dom = UBound(excelData)
    ProgressBar.Show
    On Error GoTo ErrorHandler
    
    ' Process each row (skip header row 1)
    For dataRow = 2 To UBound(excelData)
        ' Check for user cancellation
        On Error Resume Next
        If ProgressBar.Progress Then
            WriteLog 2, CurrentMod, PROC_NAME, "Cancelled by user"
            MsgBox "Cancelled by user.", vbExclamation, "Cancelled"
            GoTo Cleanup
        End If
        On Error GoTo ErrorHandler
        
        ' Only include active tasks
        If excelData(dataRow, 3) = "Yes" Then
            rowIndex = rowIndex + 1
            
            ' Add new row if needed
            If rowIndex > 2 Then
                On Error Resume Next
                newDoc.Tables(1).Rows.Add
                If Err.Number <> 0 Then
                    WriteLog 3, CurrentMod, PROC_NAME, "Failed to add table row: " & Err.Description
                    GoTo NextRow
                End If
                On Error GoTo ErrorHandler
            End If
            
            ' Populate cells
            Call PopulateTableRow(newDoc, rowIndex, excelData, dataRow)
        End If
        
NextRow:
    Next dataRow
    
    ' Set document properties
    On Error Resume Next
    SetProperty pDocCreateDate, Format$(ToServerTime, DateFormat), newDoc
    On Error GoTo ErrorHandler
    
Cleanup:
    ' Protect document
    Call Protect(newDoc)
    
    ' Save document
    On Error Resume Next
    newDoc.SaveAs2 MSPOutputPath, wdOpenFormatXMLDocument
    newDoc.Saved = True
    On Error GoTo ErrorHandler
    
    ' Restore screen updating
    Application.ScreenUpdating = True
    
    ' Delete temporary template file
    On Error Resume Next
    Kill localWordPath
    Unload ProgressBar
    On Error GoTo ErrorHandler
    
    Set CreateNewMSP = newDoc
    WriteLog 1, CurrentMod, PROC_NAME, "MS Project document created successfully"
    Exit Function
    
ErrorHandler:
    errorMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errorMsg
    
    ' Cleanup
    On Error Resume Next
    Application.ScreenUpdating = True
    Unload ProgressBar
    If Len(localWordPath) > 0 Then Kill localWordPath
    On Error GoTo 0
    
    MsgBox "Failed to create MS Project document: " & errorMsg, vbCritical, "Error"
    Set CreateNewMSP = Nothing
End Function

'=======================================================
' HELPER FUNCTIONS
'=======================================================

'=======================================================
' Sub: CleanupExcel
' Purpose: Safely cleanup Excel application and workbook
'
' Parameters:
'   excelApp - Excel application object
'   workbook - Workbook object (can be Nothing)
'   wasOpen - True if Excel was already running
'
' Description:
'   Closes workbook without saving, restores Excel settings,
'   and quits Excel if we started it.
'=======================================================
Private Sub CleanupExcel(ByVal excelApp As Object, _
                        ByVal workbook As Object, _
                        ByVal wasOpen As Boolean)
    Const PROC_NAME As String = "CleanupExcel"
    
    On Error Resume Next
    
    ' Close workbook
    If Not workbook Is Nothing Then
        workbook.Close SaveChanges:=False
        WriteLog 1, CurrentMod, PROC_NAME, "Workbook closed"
    End If
    
    ' Restore Excel settings
    If Not excelApp Is Nothing Then
        excelApp.EnableEvents = True
        excelApp.ScreenUpdating = True
        excelApp.DisplayAlerts = True
        
        If wasOpen Then
            excelApp.Visible = True
            WriteLog 1, CurrentMod, PROC_NAME, "Excel restored to visible"
        Else
            excelApp.Quit
            WriteLog 1, CurrentMod, PROC_NAME, "Excel application quit"
        End If
    End If
    
    On Error GoTo 0
End Sub

'=======================================================
' Sub: PopulateTableRow
' Purpose: Populate a single table row with Excel data
'
' Parameters:
'   Doc - Word document
'   rowIndex - Table row index (1-based)
'   excelData - Array of Excel data
'   dataRow - Source data row index
'
' Description:
'   Fills table cells with formatted data from Excel.
'   Applies styling based on task outline level.
'=======================================================
Private Sub PopulateTableRow(ByVal Doc As Document, _
                            ByVal rowIndex As Long, _
                            ByRef excelData As Variant, _
                            ByVal dataRow As Long)
    Const PROC_NAME As String = "PopulateTableRow"
    
    Dim colIndex As Long
    Dim outlineLevel As Long
    
    On Error Resume Next
    
    colIndex = 1
    
    ' Column 1: ID
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = excelData(dataRow, 1)
    colIndex = colIndex + 1
    
    ' Column 2: Number
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = excelData(dataRow, 2)
    colIndex = colIndex + 1
    
    ' Column 3: Task Name (with indentation based on outline level)
    outlineLevel = excelData(dataRow, 10)
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = _
        Space$(outlineLevel * 2) & excelData(dataRow, 5)
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Size = 14 - outlineLevel
    
    ' Apply color and bold based on outline level
    Select Case outlineLevel
        Case 0
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Color = wdColorBlack
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Bold = True
        Case 1
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Color = wdColorBlue
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Bold = True
        Case 2
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Color = wdColorDarkGreen
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Bold = False
        Case 3
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Color = wdColorPlum
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Bold = False
        Case 4
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Color = wdColorLightTurquoise
            Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.Font.Bold = False
    End Select
    colIndex = colIndex + 1
    
    ' Column 4: Duration (first element after split)
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = Split(excelData(dataRow, 6))(0)
    colIndex = colIndex + 1
    
    ' Column 5: Start date
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = Format$(excelData(dataRow, 7), "mmm d")
    colIndex = colIndex + 1
    
    ' Column 6: Finish date
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = Format$(excelData(dataRow, 8), "mmm d")
    colIndex = colIndex + 1
    
    ' Column 7: Resource names
    Doc.Tables(1).Rows(rowIndex).Cells(colIndex).Range.text = excelData(dataRow, 11)
    
    If Err.Number <> 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Error populating row " & rowIndex & ": " & Err.Description
    End If
    
    On Error GoTo 0
End Sub
