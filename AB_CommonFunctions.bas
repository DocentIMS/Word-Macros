Attribute VB_Name = "AB_CommonFunctions"
Option Explicit
Option Compare Text

'=======================================================
' Module: AB_CommonFunctions
' Purpose: Common utility functions for document operations
' Author: Refactored - November 2025
' Version: 2.0
'
' Description:
'   Provides shared functionality for document operations
'   including export, protection, time conversion, and
'   validation operations.
'
' Dependencies:
'   - AB_GlobalConstants
'   - AB_GlobalVars
'   - AZ_Log_Mod (for WriteLog)
'   - AC_Properties (for GetProperty/SetProperty)
'
' Change Log:
'   v2.0 - Nov 2025
'       * Fixed module constant (was "Export_mod")
'       * Added proper error handling throughout
'       * Fixed resource leaks in ExportRange
'       * Removed commented debug code
'       * Improved documentation
'       * Split large functions
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "AB_CommonFunctions"

' Module-level collections
Private EnforcedStyleColl As Collection

'=======================================================
' DOCUMENT OPENING FUNCTIONS
'=======================================================

'=======================================================
' Function: OpenAsDocentDocument
' Purpose: Open a file dialog and import document as Docent document
'
' Parameters:
'   DocType - Type of document to open
'
' Returns:
'   Document object if successful, Nothing on cancel/error
'
' Description:
'   Presents file dialog to user, sets up DocInfo metadata,
'   and opens the selected document.
'=======================================================
Function OpenAsDocentDocument(ByVal DocType As String) As Document
    Const PROC_NAME As String = "OpenAsDocentDocument"
    
    Dim fileName As String
    Dim newDoc As Document
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Opening document of type: " & DocType
    
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .ButtonName = "Open"
        .Title = "Browse to " & LCase(DocType)
        .InitialFileName = Environ("userprofile") & "\desktop\"
        .Filters.Add "Word Files", "*.docx; *.docm; *.doc", 1
        
        If .Show Then
            fileName = .SelectedItems(1)
            
            ' Setup document metadata
            Set OpeningDocInfo = New DocInfo
            With OpeningDocInfo
                .ContractNo = ContractNumberStr
                .DocState = ""
                .DocType = DocType
                .IsDocument = True
                .IsTemplate = False
                .Name = GetFileName(fileName)
                .PName = ProjectNameStr
                .PURL = ProjectURLStr
                .DocURL = ""
                .DocVer = 1
            End With
            
            ' Open the document
            Set newDoc = Application.Documents.Open(fileName)
            Set OpenAsDocentDocument = newDoc
            
            WriteLog 1, CurrentMod, PROC_NAME, "Document opened successfully"
        Else
            ' User cancelled
            WriteLog 1, CurrentMod, PROC_NAME, "User cancelled file selection"
            Set OpenAsDocentDocument = Nothing
        End If
    End With
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    Set OpenAsDocentDocument = Nothing
End Function

'=======================================================
' DOCUMENT PROTECTION FUNCTIONS
'=======================================================

'=======================================================
' Sub: Protect
' Purpose: Protect document with optional password
'
' Parameters:
'   Doc - Document to protect (optional, uses ActiveDocument)
'   Password - Protection password (optional)
'
' Description:
'   Protects document for read-only with tracked changes.
'   Preserves enforce style setting from collection.
'=======================================================
Sub Protect(Optional ByVal Doc As Document = Nothing, _
            Optional ByVal Password As String = "")
    Const PROC_NAME As String = "Protect"
    
    Dim EnforceStyleBool As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Use active document if none specified
    If Doc Is Nothing Then Set Doc = ActiveDocument
    
    ' Validate document
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "No document to protect"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Protecting document: " & Doc.Name
    
    ' Unprotect first if already protected
    If Doc.ProtectionType <> wdNoProtection Then
        Call Unprotect(Doc, Password)
    End If
    
    ' Get enforce style setting
    On Error Resume Next
    EnforceStyleBool = EnforcedStyleColl(Doc.Name)
    On Error GoTo ErrorHandler
    
    ' Protect the document
    Doc.Protect wdAllowOnlyReading, False, Password, False, EnforceStyleBool
    Doc.Windows(1).View.ShadeEditableRanges = False
    Application.TaskPanes(wdTaskPaneDocumentProtection).Visible = False
    
    WriteLog 1, CurrentMod, PROC_NAME, "Document protected successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: Unprotect
' Purpose: Unprotect document with optional password
'
' Parameters:
'   Doc - Document to unprotect (optional, uses ActiveDocument)
'   Password - Protection password (optional)
'=======================================================
Sub Unprotect(Optional ByVal Doc As Document = Nothing, _
              Optional ByVal Password As String = "")
    Const PROC_NAME As String = "Unprotect"
    
    On Error GoTo ErrorHandler
    
    ' Use active document if none specified
    If Doc Is Nothing Then Set Doc = ActiveDocument
    
    ' Validate document
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "No document to unprotect"
        Exit Sub
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Unprotecting document: " & Doc.Name
    
    ' Initialize collection if needed
    If EnforcedStyleColl Is Nothing Then
        Set EnforcedStyleColl = New Collection
    End If
    
    ' Save enforce style setting
    On Error Resume Next
    EnforcedStyleColl.Remove Doc.Name
    EnforcedStyleColl.Add Doc.EnforceStyle, Doc.Name
    On Error GoTo ErrorHandler
    
    ' Unprotect the document
    Doc.Unprotect Password
    
    WriteLog 1, CurrentMod, PROC_NAME, "Document unprotected successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' TIME CONVERSION FUNCTIONS
'=======================================================

'=======================================================
' Function: AlreadyServerTime
' Purpose: Convert datetime to server time (already in UTC)
'
' Parameters:
'   DateTime - DateTime string (optional, uses Now if empty)
'
' Returns:
'   Date in server timezone
'=======================================================
Function AlreadyServerTime(Optional ByVal DateTime As String = "") As Date
    Const PROC_NAME As String = "AlreadyServerTime"
    
    On Error GoTo ErrorHandler
    
    If Len(DateTime) > 0 Then
        If InStr(DateTime, "T") > 0 Then
            AlreadyServerTime = TimeFromTFormat(DateTime) + TimeSerial(PloneTimeZone, 0, 0)
        Else
            AlreadyServerTime = ANTIConvertToUtc(DateValue(DateTime) + TimeValue(DateTime)) + TimeSerial(PloneTimeZone, 0, 0)
        End If
    Else
        AlreadyServerTime = ANTIConvertToUtc(Now) + TimeSerial(PloneTimeZone, 0, 0)
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    AlreadyServerTime = Now
End Function

'=======================================================
' Function: ToServerTime
' Purpose: Convert local time to server time
'=======================================================
Function ToServerTime(Optional ByVal DateTime As String = "") As Date
    Const PROC_NAME As String = "ToServerTime"
    
    On Error GoTo ErrorHandler
    
    If Len(DateTime) > 0 Then
        If InStr(DateTime, "T") > 0 Then
            ToServerTime = TimeFromTFormat(DateTime) + TimeSerial(PloneTimeZone, 0, 0)
        Else
            ToServerTime = ConvertToUtc(DateValue(DateTime) + TimeValue(DateTime)) + TimeSerial(PloneTimeZone, 0, 0)
        End If
    Else
        ToServerTime = ConvertToUtc(Now) + TimeSerial(PloneTimeZone, 0, 0)
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ToServerTime = Now
End Function

'=======================================================
' Function: FromServerTime
' Purpose: Convert server time to local time
'=======================================================
Function FromServerTime(Optional ByVal DateTime As String = "") As Date
    Const PROC_NAME As String = "FromServerTime"
    
    On Error GoTo ErrorHandler
    
    If Len(DateTime) > 0 Then
        If InStr(DateTime, "T") > 0 Then
            FromServerTime = TimeFromTFormat(DateTime) - TimeSerial(PloneTimeZone, 0, 0)
        Else
            FromServerTime = ConvertToUtc(DateValue(DateTime) + TimeValue(DateTime)) - TimeSerial(PloneTimeZone, 0, 0)
        End If
    Else
        FromServerTime = ConvertToUtc(Now) - TimeSerial(PloneTimeZone, 0, 0)
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    FromServerTime = Now
End Function

'=======================================================
' DOCUMENT VALIDATION FUNCTIONS
'=======================================================

'=======================================================
' Function: ValidDocument
' Purpose: Validate if active document matches expected type
'
' Parameters:
'   DocType - Expected document type
'
' Returns:
'   Boolean - True if document matches type or user confirms
'=======================================================
Function ValidDocument(ByVal DocType As String) As Boolean
    Const PROC_NAME As String = "ValidDocument"
    
    Dim Rng As Range
    Dim fileName As String
    Dim userResponse As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Validating document type: " & DocType
    
    ' Initialize collection if needed
    If NotScopeAsked Is Nothing Then Set NotScopeAsked = New Collection
    
    ' Check if active document exists
    If ActiveDocument Is Nothing Then
        MsgBox "No document is open", vbCritical, "No Document"
        ValidDocument = False
        Exit Function
    End If
    
    fileName = ActiveDocument.Name
    
    ' Check cache first
    On Error Resume Next
    ValidDocument = NotScopeAsked(fileName)
    If Err.Number = 0 And ValidDocument Then
        WriteLog 1, CurrentMod, PROC_NAME, "Validation cached: True"
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Check document type property
    If IsSameDocType(GetPropertySafe(pDocType), DocType) Then
        ValidDocument = True
        WriteLog 1, CurrentMod, PROC_NAME, "Document type matched by property"
        Exit Function
    End If
    
    ' Check filename patterns
    If ActiveDocument.Name = DocType Then
        ValidDocument = True
        Exit Function
    End If
    
    If ActiveDocument.Name Like "* " & DocType & " *" Then
        ValidDocument = True
        Exit Function
    End If
    
    If ActiveDocument.Name Like DocType & " *" Then
        ValidDocument = True
        Exit Function
    End If
    
    If ActiveDocument.Name Like "* " & DocType Then
        ValidDocument = True
        Exit Function
    End If
    
    ' Search document content
    Set Rng = ActiveDocument.Range
    Set Rng = Rng.GoTo(What:=wdGoToBookmark, Name:="\page")
    
    With Rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .text = DocType
        .Execute
        ValidDocument = .Found
        
        If ValidDocument Then
            WriteLog 1, CurrentMod, PROC_NAME, "Document type found in content"
            Exit Function
        End If
    End With
    
    ' Ask user for confirmation
    On Error Resume Next
    NotScopeAsked.Remove fileName
    On Error GoTo ErrorHandler
    
    userResponse = MsgBox("Are you sure this is a " & DocType & " document?", _
                         vbExclamation + vbYesNo, "Confirm Document Type")
    ValidDocument = (userResponse = vbYes)
    
    If ValidDocument Then
        WriteLog 2, CurrentMod, PROC_NAME, _
                 "User confirmed document type despite no match"
    End If
    
    ' Cache the result
    NotScopeAsked.Add ValidDocument, fileName
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ValidDocument = False
End Function

'=======================================================
' Function: GetPropertySafe
' Purpose: Safely get property with default value on error
'
' Parameters:
'   propertyName - Name of property to retrieve
'
' Returns:
'   Property value or empty string on error
'=======================================================
Private Function GetPropertySafe(ByVal propertyName As DocProperty) As Variant
    On Error Resume Next
    GetPropertySafe = GetProperty(propertyName)
    If Err.Number <> 0 Then GetPropertySafe = ""
    On Error GoTo 0
End Function

'=======================================================
' Function: IsSameDocType
' Purpose: Compare two document types by acronym
'=======================================================
Function IsSameDocType(ByVal DocType1 As String, ByVal DocType2 As String) As Boolean
    IsSameDocType = (GetAcronym(DocType1) = GetAcronym(DocType2))
End Function

'=======================================================
' Function: GetAcronym
' Purpose: Get acronym from multi-word string
'
' Example:
'   "Scope Document" returns "SD"
'   "RFP" returns "RFP"
'=======================================================
Function GetAcronym(ByVal Wrd As String) As String
    Dim s As String
    Dim ss() As String
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If InStr(Wrd, " ") > 0 Then
        ss = Split(Wrd, " ")
        For i = LBound(ss) To UBound(ss)
            s = s & UCase$(Left$(ss(i), 1))
        Next i
        GetAcronym = s
    Else
        GetAcronym = Wrd
    End If
    Exit Function
    
ErrorHandler:
    GetAcronym = Wrd
End Function

'=======================================================
' EXPORT FUNCTIONS
'=======================================================

'=======================================================
' Function: ExportRange
' Purpose: Export document range to HTML with proper resource management
'
' Parameters:
'   SecRng - Section range to export
'   PreviousRng - Previous section range for context
'   NextRng - Next section range for context
'   SecName - Section name for file naming
'   HTMLPath - Output directory path
'   SeqNo - Sequence number for file naming
'   DeliverablesRng - Optional deliverables range
'   HighlightedRng - Optional range to highlight
'
' Returns:
'   String - Path to exported HTML file, empty string on error
'
' Resource Management:
'   - Ensures temporary documents are always closed
'   - Releases all object references
'   - Cleans up on error conditions
'=======================================================
Function ExportRange(ByVal SecRng As Range, _
                    ByVal PreviousRng As Range, _
                    ByVal NextRng As Range, _
                    ByVal SecName As String, _
                    ByVal HTMLPath As String, _
                    ByVal SeqNo As Long, _
                    Optional ByVal DeliverablesRng As Range = Nothing, _
                    Optional ByVal HighlightedRng As Range = Nothing) As String
    Const PROC_NAME As String = "ExportRange"
    
    Dim tempDoc As Document
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Exporting section: " & SecName
    
    ' Validate inputs
    If SecRng Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "SecRng is Nothing"
        ExportRange = ""
        Exit Function
    End If
    
    If Len(HTMLPath) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "HTMLPath is empty"
        ExportRange = ""
        Exit Function
    End If
    
    ' Ensure path ends with backslash
    If Right$(HTMLPath, 1) <> "\" Then HTMLPath = HTMLPath & "\"
    
    ' Build filename
    fileName = HTMLPath & Format$(SeqNo, "000") & "-" & _
               SanitizeFileName(Left$(SecName, 30)) & ".html"
    
    ' Create temporary document
    Set tempDoc = Application.Documents.Add
    If tempDoc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Failed to create temporary document"
        ExportRange = ""
        Exit Function
    End If
    
    ' Insert ranges
    Call InsertRng(tempDoc, PreviousRng, 1, Nothing)
    Call InsertRng(tempDoc, SecRng, 2, HighlightedRng)
    Call InsertRng(tempDoc, NextRng, 3, Nothing)
    
    ' Save as HTML
    tempDoc.SaveAs2 fileName, wdFormatFilteredHTML, , , False
    
    ' Close temporary document
    tempDoc.Close SaveChanges:=False
    Set tempDoc = Nothing
    
    ' Process HTML file
    Call ProcessHTMLFile(fileName)
    
    ' Export deliverables if present
    If Not DeliverablesRng Is Nothing Then
        If Len(Trim$(DeliverablesRng.text)) > 0 Then
            Call ExportDeliverablesSection(DeliverablesRng, fileName)
        End If
    End If
    
    ExportRange = fileName
    WriteLog 1, CurrentMod, PROC_NAME, "Export completed: " & fileName
    Exit Function
    
ErrorHandler:
    Dim errMsg As String
    errMsg = "Error " & Err.Number & ": " & Err.Description
    WriteLog 3, CurrentMod, PROC_NAME, errMsg & " (Section: " & SecName & ")"
    
    ' Cleanup on error
    Call CleanupTempDocument(tempDoc)
    
    ExportRange = ""
End Function

'=======================================================
' Sub: CleanupTempDocument
' Purpose: Safely close and release temporary document
'
' Parameters:
'   Doc - Document to cleanup (can be Nothing)
'=======================================================
Private Sub CleanupTempDocument(ByRef Doc As Document)
    On Error Resume Next
    
    If Not Doc Is Nothing Then
        Doc.Close SaveChanges:=False
        Set Doc = Nothing
    End If
    
    On Error GoTo 0
End Sub

'=======================================================
' Function: SanitizeFileName
' Purpose: Remove invalid characters from filename
'
' Parameters:
'   fileName - Original filename
'
' Returns:
'   String - Sanitized filename safe for file system
'=======================================================
Private Function SanitizeFileName(ByVal fileName As String) As String
    Const INVALID_CHARS As String = "\/:*?""<>|"
    Dim i As Long
    Dim result As String
    
    result = fileName
    
    For i = 1 To Len(INVALID_CHARS)
        result = Replace(result, Mid$(INVALID_CHARS, i, 1), "_")
    Next i
    
    SanitizeFileName = result
End Function

'=======================================================
' Function: ExportDeliverablesSection
' Purpose: Export deliverables range to separate HTML file
'=======================================================
Private Function ExportDeliverablesSection(ByVal DeliverablesRng As Range, _
                                          ByVal baseFileName As String) As Boolean
    Const PROC_NAME As String = "ExportDeliverablesSection"
    
    Dim tempDoc As Document
    Dim delivFileName As String
    
    On Error GoTo ErrorHandler
    
    delivFileName = baseFileName & "deliverables.html"
    
    Set tempDoc = Application.Documents.Add
    If tempDoc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Failed to create temporary document"
        ExportDeliverablesSection = False
        Exit Function
    End If
    
    Call InsertRng(tempDoc, DeliverablesRng, 2, Nothing)
    
    tempDoc.SaveAs2 delivFileName, wdFormatFilteredHTML, , , False
    tempDoc.Close SaveChanges:=False
    Set tempDoc = Nothing
    
    WriteLog 1, CurrentMod, PROC_NAME, "Exported deliverables: " & delivFileName
    ExportDeliverablesSection = True
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    Call CleanupTempDocument(tempDoc)
    ExportDeliverablesSection = False
End Function

'=======================================================
' Sub: ProcessHTMLFile
' Purpose: Post-process HTML file with HR tags
'=======================================================
Private Sub ProcessHTMLFile(ByVal fileName As String)
    Const PROC_NAME As String = "ProcessHTMLFile"
    
    On Error GoTo ErrorHandler
    
    Call AddHRToHTML(fileName)
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Function: AddHRToHTML
' Purpose: Add horizontal rules to HTML file
'=======================================================
Private Function AddHRToHTML(ByVal fileName As String) As String
    Const PROC_NAME As String = "AddHRToHTML"
    
    Dim fileNum As Long
    Dim htmlContent As String
    Dim processedHTML As String
    
    On Error GoTo ErrorHandler
    
    ' Read file
    fileNum = FreeFile
    Open fileName For Input As #fileNum
    htmlContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ' Process HTML
    processedHTML = InsertHRs(htmlContent)
    
    ' Write back if changed
    If Len(processedHTML) <> Len(htmlContent) Then
        fileNum = FreeFile
        Open fileName For Output As #fileNum
        Print #fileNum, processedHTML
        Close #fileNum
        
        WriteLog 1, CurrentMod, PROC_NAME, "Added HR tags to: " & fileName
    End If
    
    AddHRToHTML = processedHTML
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    ' Ensure file is closed
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    
    AddHRToHTML = ""
End Function

'=======================================================
' Function: InsertHRs
' Purpose: Insert horizontal rule tags at break points
'=======================================================
Private Function InsertHRs(ByVal XML As String) As String
    Dim i As Long
    Dim j As Long
    Dim tagEnd As String
    Dim leftPart As String
    
    On Error GoTo ErrorHandler
    
    Do
        ' Find break keyword
        i = InStr(1, XML, BRKwrd, vbTextCompare)
        If i = 0 Then Exit Do
        
        ' Find preceding tag
        j = InStrRev(XML, "<h", i, vbTextCompare)
        i = InStrRev(XML, "<p", i, vbTextCompare)
        
        If i < j Then
            i = j
            tagEnd = "</h"
        Else
            tagEnd = "</p"
        End If
        
        ' Find tag close
        j = InStr(i, XML, tagEnd, vbTextCompare)
        If j > 0 Then
            j = InStr(j, XML, ">", vbTextCompare) + 1
        Else
            Exit Do
        End If
        
        ' Insert HR tag
        leftPart = Left$(XML, i - 1)
        XML = leftPart & HRTag & Right$(XML, Len(XML) - j + 1)
    Loop
    
    InsertHRs = XML
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "InsertHRs", _
             "Error " & Err.Number & ": " & Err.Description
    InsertHRs = XML
End Function

'=======================================================
' Function: RemoveHighlighting
' Purpose: Remove highlighting spans from HTML
'=======================================================
Private Function RemoveHighlighting(ByVal XML As String) As String
    Dim i As Long
    Dim j As Long
    Dim leftPart As String
    
    On Error GoTo ErrorHandler
    
    Do
        i = InStr(1, XML, "<span style='background:", vbTextCompare)
        If i = 0 Then Exit Do
        
        j = InStr(i, XML, ">", vbTextCompare) + 1
        leftPart = Left$(XML, i - 1)
        XML = leftPart & "<span>" & Right$(XML, Len(XML) - j + 1)
    Loop
    
    RemoveHighlighting = XML
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, "RemoveHighlighting", _
             "Error " & Err.Number & ": " & Err.Description
    RemoveHighlighting = XML
End Function

'=======================================================
' Sub: InsertRng
' Purpose: Insert range into document with formatting
'
' Parameters:
'   Doc - Target document
'   sRng - Source range to insert
'   Mode - Insert mode (1=Previous, 2=Main, 3=Next)
'   HighlightedRng - Optional range to highlight
'=======================================================
Private Sub InsertRng(ByVal Doc As Document, _
                     ByVal sRng As Range, _
                     ByVal Mode As Long, _
                     ByVal HighlightedRng As Range)
    Const PROC_NAME As String = "InsertRng"
    
    Dim Rng As Range
    Dim i As Long
    Dim x As Long
    Dim l As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Mode: " & Mode
    
    ' Validate inputs
    If Doc Is Nothing Then Exit Sub
    If sRng Is Nothing Then Exit Sub
    If Len(sRng.text) = 0 Then Exit Sub
    
    ' Get target range
    Set Rng = Doc.Range
    Rng.Collapse wdCollapseStart
    
    If Rng.start > 0 Then
        Rng.text = vbNewLine
        Rng.Collapse wdCollapseStart
    End If
    
    Select Case Mode
        Case 1, 3  ' Previous or Next section
            On Error Resume Next
            Rng.FormattedText = sRng.Paragraphs(1).Range
            
            If Err.Number = 0 Then
                Rng.Move wdParagraph, 10
                Set Rng = Rng.Previous(wdParagraph, 1)
                Rng.Collapse wdCollapseEnd
                x = Rng.start
                
                ' Add section label
                If Mode = 1 Then
                    Rng.text = "Previous section ("
                Else
                    Rng.text = "Next section ("
                End If
                
                Rng.MoveUntil vbNewLine
                Rng.text = ")" & vbNewLine
                
                ' Add highlighted lines
                l = HighlightLines
                For i = 2 To HighlightLines + 1
                    Rng.Collapse wdCollapseStart
                    l = l - RangeLines(sRng.Paragraphs(i).Range)
                    Rng.FormattedText = sRng.Paragraphs(i).Range
                    If l < 0 Then Exit For
                Next i
            End If
            On Error GoTo ErrorHandler
            
            ' Format as gray
            Rng.SetRange x, Rng.End
            Rng.Font.ColorIndex = wdGray25
            Rng.HighlightColorIndex = wdNoHighlight
            
            ' Add break markers
            If Mode = 1 Then
                Rng.Collapse wdCollapseStart
                Rng.text = vbLf & BRKwrd
            ElseIf Mode = 3 Then
                i = Rng.start
                Rng.SetRange i, i
                Rng.text = BRKwrd & vbLf
            End If
            
        Case 2  ' Main section
            i = Rng.start - sRng.start
            Rng.FormattedText = sRng
            
            ' Apply highlighting if specified
            If Not HighlightedRng Is Nothing Then
                x = HighlightedRng.HighlightColorIndex
                Rng.HighlightColorIndex = wdNoHighlight
                Rng.SetRange HighlightedRng.start + i, HighlightedRng.End + i
                Rng.HighlightColorIndex = x
            End If
    End Select
    
    WriteLog 1, CurrentMod, PROC_NAME, "Completed"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Function: RangeLines
' Purpose: Count number of lines in a range
'=======================================================
Private Function RangeLines(ByVal Rng As Range) As Long
    Dim tempRng As Range
    
    On Error GoTo ErrorHandler
    
    If Rng Is Nothing Then
        RangeLines = 0
        Exit Function
    End If
    
    Set tempRng = Rng.Duplicate
    tempRng.Collapse wdCollapseStart
    
    RangeLines = tempRng.Information(wdFirstCharacterLineNumber) - _
                 Rng.Information(wdFirstCharacterLineNumber) + 1
    Exit Function
    
ErrorHandler:
    RangeLines = 1
End Function
