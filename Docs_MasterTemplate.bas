Attribute VB_Name = "Docs_MasterTemplate"
Option Explicit

'=======================================================
' Module: Docs_MasterTemplate
' Purpose: Update document templates from master templates
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Applies master template settings to documents including:
'   - Document theme
'   - Styles and Quick Styles
'   - Page setup and margins
'   - Header and footer content
'   - Images and logos
'   - Content controls (optional)
'
' Dependencies:
'   - AB_CommonFunctions (for Unprotect, Protect, GetFileName)
'   - AZ_Log_Mod (for WriteLog)
'   - Word.Application object model
'
' Main Procedures:
'   - UpdateTemplate: Main template update workflow
'   - CopyStyles: Transfer style settings
'   - CopyImagesToDocument: Transfer header/footer images
'   - UpdateHdFtCC: Update header/footer content controls
'   - UpdateHdFtLogos: Update header/footer logos
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive documentation
'       * Improved error handling
'       * Removed commented dead code
'       * Added validation and logging
'       * Improved resource management
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Docs_MasterTemplate"

'=======================================================
' PUBLIC PROCEDURES
'=======================================================

'=======================================================
' Sub: UpdateTemplate
' Purpose: Apply master template settings to document
'
' Parameters:
'   Doc - Target document to update
'   MasterFName - Full path to master template file
'
' Description:
'   Updates document with all settings from master template:
'   1. Applies document theme
'   2. Copies styles from template
'   3. Updates page setup (margins, size)
'   4. Copies header/footer images and content
'
' Side Effects:
'   - Modifies document theme
'   - Updates document styles
'   - Changes page setup
'   - Replaces header/footer content
'   - Temporarily unprotects document
'
' Error Handling:
'   Ensures master template is closed even on error
'   Restores document protection
'   Logs all errors
'=======================================================
Sub UpdateTemplate(ByVal Doc As Document, ByVal MasterFName As String)
    Const PROC_NAME As String = "UpdateTemplate"
    
    Dim masterDoc As Document
    Dim wasProtected As Boolean
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, "Starting template update from: " & MasterFName
    
    ' Validate inputs
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Target document is Nothing"
        MsgBox "No document specified for template update.", vbExclamation, "Invalid Document"
        Exit Sub
    End If
    
    If Len(MasterFName) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Master template filename is empty"
        MsgBox "Master template filename is required.", vbExclamation, "Invalid Filename"
        Exit Sub
    End If
    
    If Len(Dir$(MasterFName)) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Master template file not found: " & MasterFName
        MsgBox "Master template file not found:" & vbNewLine & vbNewLine & _
               MasterFName & vbNewLine & vbNewLine & _
               "Please verify the file path is correct.", _
               vbExclamation, "Template Not Found"
        Exit Sub
    End If
    
    ' Close master if already open (prevents conflicts)
    Call CloseDocumentIfOpen(GetFileName(MasterFName))
    
    ' Save protection state and unprotect for modification
    wasProtected = (Doc.protectionType <> wdNoProtection)
    Unprotect Doc
    
    WriteLog 1, CurrentMod, PROC_NAME, "Applying document theme"
    
    ' Apply theme from master template
    Doc.ApplyDocumentTheme MasterFName
    
    ' Open master template (read-only for safety)
    WriteLog 1, CurrentMod, PROC_NAME, "Opening master template"
    Set masterDoc = Documents.Open(fileName:=MasterFName, ReadOnly:=True)
    
    If masterDoc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Failed to open master template"
        GoTo Cleanup
    End If
    
    ' Copy styles from template
    WriteLog 1, CurrentMod, PROC_NAME, "Copying styles from template"
    Doc.CopyStylesFromTemplate MasterFName
    Call CopyStyles(Doc, masterDoc)
    
    ' Update page setup
    WriteLog 1, CurrentMod, PROC_NAME, "Updating page setup"
    Call CopyPageSetup(Doc, masterDoc)
    
    ' Copy header/footer images and content controls
    WriteLog 1, CurrentMod, PROC_NAME, "Copying header/footer content"
    Call CopyImagesToDocument(Doc, masterDoc, IncludeContentControls:=True)
    
Cleanup:
    ' Restore protection
    If wasProtected Then Protect Doc
    
    ' Close master template
    If Not masterDoc Is Nothing Then
        masterDoc.Close SaveChanges:=False
        Set masterDoc = Nothing
    End If
    
    WriteLog 1, CurrentMod, PROC_NAME, "Template update completed successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    MsgBox "An error occurred while updating the template:" & vbNewLine & vbNewLine & _
           "Error: " & Err.Description & vbNewLine & vbNewLine & _
           "The template update may be incomplete.", _
           vbCritical, "Template Update Error"
    
    ' Ensure cleanup happens even on error
    On Error Resume Next
    If wasProtected Then Protect Doc
    If Not masterDoc Is Nothing Then
        masterDoc.Close SaveChanges:=False
        Set masterDoc = Nothing
    End If
    On Error GoTo 0
End Sub

'=======================================================
' PRIVATE HELPER PROCEDURES
'=======================================================

'=======================================================
' Sub: CloseDocumentIfOpen
' Purpose: Close a document if it's currently open
'
' Parameters:
'   docName - Name of document to close
'
' Description:
'   Checks if document is open and closes it without saving.
'   Used to prevent conflicts when opening master template.
'=======================================================
Private Sub CloseDocumentIfOpen(ByVal docName As String)
    Const PROC_NAME As String = "CloseDocumentIfOpen"
    
    Dim Doc As Document
    
    On Error Resume Next
    Set Doc = Documents(docName)
    
    If Not Doc Is Nothing Then
        WriteLog 1, CurrentMod, PROC_NAME, "Closing already-open document: " & docName
        Doc.Close SaveChanges:=False
        Set Doc = Nothing
    End If
    
    On Error GoTo 0
End Sub

'=======================================================
' Sub: CopyPageSetup
' Purpose: Copy page setup settings from master to target
'
' Parameters:
'   targetDoc - Document to update
'   masterDoc - Source master template
'
' Description:
'   Copies page dimensions, margins, and header/footer
'   distances from master template to target document.
'=======================================================
Private Sub CopyPageSetup(ByVal targetDoc As Document, _
                         ByVal masterDoc As Document)
    Const PROC_NAME As String = "CopyPageSetup"
    
    On Error GoTo ErrorHandler
    
    With targetDoc.PageSetup
        .PageWidth = masterDoc.PageSetup.PageWidth
        .PageHeight = masterDoc.PageSetup.PageHeight
        .LeftMargin = masterDoc.PageSetup.LeftMargin
        .RightMargin = masterDoc.PageSetup.RightMargin
        .TopMargin = masterDoc.PageSetup.TopMargin
        .BottomMargin = masterDoc.PageSetup.BottomMargin
        .HeaderDistance = masterDoc.PageSetup.HeaderDistance
        .FooterDistance = masterDoc.PageSetup.FooterDistance
    End With
    
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' Sub: CopyStyles
' Purpose: Copy QuickStyle settings from master to target
'
' Parameters:
'   targetDoc - Document to update
'   masterDoc - Source master template
'
' Description:
'   Synchronizes QuickStyle gallery visibility between
'   master and target documents. Styles that are QuickStyles
'   in the master will be made QuickStyles in the target.
'
' Error Handling:
'   Uses On Error Resume Next because some styles may not
'   exist in both documents (built-in vs custom styles)
'=======================================================
Private Sub CopyStyles(ByVal targetDoc As Document, _
                      ByVal masterDoc As Document)
    Const PROC_NAME As String = "CopyStyles"
    
    Dim targetStyle As Style
    Dim isQuickStyle As Boolean
    Dim stylesCopied As Long
    
    On Error Resume Next  ' Expected: Some styles may not exist in both documents
    
    WriteLog 1, CurrentMod, PROC_NAME, "Syncing QuickStyle settings"
    
    For Each targetStyle In targetDoc.Styles
        isQuickStyle = False
        
        ' Check if this style is a QuickStyle in master
        isQuickStyle = masterDoc.Styles(targetStyle.NameLocal).QuickStyle
        
        If Err.Number = 0 Then
            ' Style exists in master - apply its QuickStyle setting
            targetStyle.QuickStyle = isQuickStyle
            If isQuickStyle Then stylesCopied = stylesCopied + 1
        End If
        
        Err.Clear
    Next targetStyle
    
    On Error GoTo 0
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Synced " & stylesCopied & " QuickStyle settings"
End Sub

'=======================================================
' Sub: CopyImagesToDocument
' Purpose: Copy header/footer images and content from master
'
' Parameters:
'   targetDoc - Document to update
'   masterDoc - Source master template
'   IncludeContentControls - If True, copy content controls;
'                            If False, copy only images/logos
'
' Description:
'   For each section in target document:
'   1. Clears existing header images
'   2. Clears existing content controls (if requested)
'   3. Copies new content from master template
'
' Error Handling:
'   Uses On Error Resume Next because:
'   - Some sections may not have headers/footers
'   - Content controls may not exist
'   - Image operations may fail for protected content
'=======================================================
Private Sub CopyImagesToDocument(ByVal targetDoc As Document, _
                                 ByVal masterDoc As Document, _
                                 ByVal IncludeContentControls As Boolean)
    Const PROC_NAME As String = "CopyImagesToDocument"
    
    Dim sectionIndex As Long
    Dim shapeIndex As Long
    Dim headerFooterIndex As Long
    Dim contentControlIndex As Long
    
    On Error Resume Next  ' Expected: Some operations may fail for various sections
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Processing " & targetDoc.Sections.Count & " sections"
    
    For sectionIndex = 1 To targetDoc.Sections.Count
        
        ' Clear old header shapes/images from first header
        For shapeIndex = targetDoc.Sections(sectionIndex).Headers(1).Shapes.Count To 1 Step -1
            targetDoc.Sections(sectionIndex).Headers(1).Shapes(shapeIndex).Delete
        Next shapeIndex
        
        ' Process all three header/footer types (first page, even page, odd page)
        For headerFooterIndex = 1 To 3
            
            If IncludeContentControls Then
                ' Clear old content controls
                For contentControlIndex = targetDoc.Sections(sectionIndex).Headers(headerFooterIndex).Range.ContentControls.Count To 1 Step -1
                    targetDoc.Sections(sectionIndex).Headers(headerFooterIndex).Range.ContentControls(contentControlIndex).Delete
                Next contentControlIndex
                
                ' Copy header and footer content with content controls
                Call UpdateHdFtCC(masterDoc.Sections(1).Headers(headerFooterIndex), _
                                 targetDoc.Sections(sectionIndex).Headers(headerFooterIndex))
                Call UpdateHdFtCC(masterDoc.Sections(1).Footers(headerFooterIndex), _
                                 targetDoc.Sections(sectionIndex).Footers(headerFooterIndex))
            Else
                ' Copy only logos/images (no content controls)
                Call UpdateHdFtLogos(masterDoc.Sections(1).Headers(headerFooterIndex), _
                                    targetDoc.Sections(sectionIndex).Headers(headerFooterIndex))
                Call UpdateHdFtLogos(masterDoc.Sections(1).Footers(headerFooterIndex), _
                                    targetDoc.Sections(sectionIndex).Footers(headerFooterIndex))
            End If
            
        Next headerFooterIndex
        
    Next sectionIndex
    
    On Error GoTo 0
    
    WriteLog 1, CurrentMod, PROC_NAME, "Header/footer content copied"
End Sub

'=======================================================
' Sub: UpdateHdFtCC
' Purpose: Update header/footer with content controls
'
' Parameters:
'   sourceHdFt - Source header/footer from master template
'   targetHdFt - Target header/footer to update
'
' Description:
'   Copies all formatted text including content controls
'   from source to target header/footer. Also preserves
'   table spacing if tables are present.
'
' Note:
'   This is a complete content replacement operation.
'   All existing content in target will be replaced.
'=======================================================
Private Sub UpdateHdFtCC(ByVal sourceHdFt As HeaderFooter, _
                        ByVal targetHdFt As HeaderFooter)
    Const PROC_NAME As String = "UpdateHdFtCC"
    
    Dim tableIndex As Long
    Dim tableCount As Long
    
    On Error Resume Next  ' Expected: Some operations may fail
    
    ' Copy all formatted text (includes content controls)
    targetHdFt.Range.FormattedText = sourceHdFt.Range.FormattedText
    
    ' Preserve table spacing
    tableCount = sourceHdFt.Range.Tables.Count
    For tableIndex = 1 To tableCount
        If tableIndex <= targetHdFt.Range.Tables.Count Then
            targetHdFt.Range.Tables(tableIndex).Range.ParagraphFormat.SpaceAfter = _
                sourceHdFt.Range.Tables(tableIndex).Range.ParagraphFormat.SpaceAfter
            targetHdFt.Range.Tables(tableIndex).Range.ParagraphFormat.SpaceBefore = _
                sourceHdFt.Range.Tables(tableIndex).Range.ParagraphFormat.SpaceBefore
        End If
    Next tableIndex
    
    On Error GoTo 0
End Sub

'=======================================================
' Sub: UpdateHdFtLogos
' Purpose: Copy images/logos from master to target header/footer
'
' Parameters:
'   sourceHdFt - Source header/footer from master template
'   targetHdFt - Target header/footer to update
'
' Description:
'   Copies all shapes (images/logos) from source to target,
'   preserving their positioning. Process:
'   1. Convert shape to inline shape (for copying)
'   2. Copy formatted text to target
'   3. Convert back to shape
'   4. Restore original position
'   5. Delete temporary inline shape
'
' Note:
'   Skips if target header/footer is linked to previous section
'=======================================================
Private Sub UpdateHdFtLogos(ByVal sourceHdFt As HeaderFooter, _
                           ByVal targetHdFt As HeaderFooter)
    Const PROC_NAME As String = "UpdateHdFtLogos"
    
    Dim tempInlineShape As InlineShape
    Dim shapeCount As Long
    Dim shapeIndex As Long
    Dim targetShape As Shape
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim leftPosition As Single
    Dim topPosition As Single
    
    On Error Resume Next  ' Expected: Some operations may fail
    
    ' Skip if this header/footer is linked to previous section
    If targetHdFt.LinkToPrevious Then Exit Sub
    
    shapeCount = sourceHdFt.Range.ShapeRange.Count
    
    ' Process shapes in reverse order to maintain layering
    For shapeIndex = 1 To shapeCount
        
        With sourceHdFt.Range.ShapeRange(shapeCount - shapeIndex + 1)
            ' Save position
            Set sourceRange = sourceHdFt.Range
            .RelativeVerticalPosition = wdRelativeVerticalPositionInnerMarginArea
            topPosition = .Top
            leftPosition = .Left
            
            ' Convert to inline shape for copying
            Set tempInlineShape = .ConvertToInlineShape
        End With
        
        ' Copy to target
        Set targetRange = targetHdFt.Range
        targetRange.Collapse wdCollapseStart
        targetRange.FormattedText = tempInlineShape.Range.FormattedText
        
        ' Convert back to shape in target
        Set targetShape = targetHdFt.Range.InlineShapes(1).ConvertToShape
        Set targetShape = targetHdFt.Shapes(targetHdFt.Shapes.Count)
        
        ' Restore position
        With targetShape
            .RelativeVerticalPosition = wdRelativeVerticalPositionInnerMarginArea
            .Left = leftPosition
            .Top = topPosition
        End With
        
        ' Clean up temporary inline shape
        tempInlineShape.Delete
        Set tempInlineShape = Nothing
        
    Next shapeIndex
    
    On Error GoTo 0
End Sub

'=======================================================
' END OF MODULE
'=======================================================
