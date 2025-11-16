Attribute VB_Name = "DocentTools_CreateBorder"
Option Explicit

'=======================================================
' Module: DocentTools_CreateBorder
' Purpose: Manage document page borders based on project colors
' Author: Updated November 2025
' Version: 2.0
'
' Description:
'   Applies colored page borders to documents based on the
'   current project's assigned color. Borders are applied to
'   all four sides of every page (top, bottom, left, right)
'   and propagated to all document sections.
'
'   Features:
'   - Applies project-colored borders to documents
'   - Removes borders when no project is selected
'   - Only updates borders when necessary (optimization)
'   - Preserves document protection and saved state
'   - Works with multi-section documents
'
' Dependencies:
'   - AB_GlobalVars (for ProjectColorStr)
'   - AB_CommonFunctions (for Unprotect, Protect)
'   - AC_Properties (for GetProperty)
'   - AZ_Log_Mod (for WriteLog)
'
' Main Procedures:
'   - CreateBorder: Apply or remove page borders
'   - BordersChangeNeeded: Check if border update required
'
' Usage Example:
'   ' Apply project color border
'   CreateBorder ActiveDocument, ProjectColorStr
'
'   ' Remove border (pass empty color)
'   CreateBorder ActiveDocument, ""
'
' Change Log:
'   v2.0 - Nov 2025
'       * Added comprehensive documentation
'       * Added error handling
'       * Removed commented code
'       * Improved type safety
'       * Added validation
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "DocentTools_CreateBorder"

' Border constants for clarity
Private Const BORDER_WIDTH_POINTS As Long = 48  ' 6 points (48/8)

'=======================================================
' MAIN PROCEDURES
'=======================================================

'=======================================================
' Sub: CreateBorder
' Purpose: Apply or remove page borders based on project color
'
' Parameters:
'   Doc - Document to modify
'   borderColor - Color to apply (Variant):
'                 - Long value: Apply border with this color
'                 - Empty string: Remove border
'
' Description:
'   Manages document page borders:
'   1. Validates document and preconditions
'   2. Checks if border update is actually needed
'   3. Temporarily unprotects document
'   4. Applies or removes borders on all four sides
'   5. Propagates borders to all sections
'   6. Restores protection and saved state
'
' Side Effects:
'   - Modifies document page borders
'   - Temporarily changes document protection
'   - Preserves document saved state
'   - Applies changes to all document sections
'
' Error Handling:
'   - Validates all inputs
'   - Logs all operations
'   - Ensures cleanup even on error
'   - Restores protection and saved state
'
' Example:
'   CreateBorder ActiveDocument, RGB(192, 104, 86)  ' Apply orange border
'   CreateBorder ActiveDocument, ""                  ' Remove border
'=======================================================
Sub CreateBorder(ByVal Doc As Document, ByVal borderColor As Variant)
    Const PROC_NAME As String = "CreateBorder"
    
    Dim borderIndex As Long
    Dim wasProtected As Boolean
    Dim wasSaved As Boolean
    Dim borderApplied As Boolean
    Dim hasProjectColor As Boolean
    Dim colorValue As Long
    
    On Error GoTo ErrorHandler
    
    WriteLog 1, CurrentMod, PROC_NAME, _
             "Processing border request with color: " & CStr(borderColor)
    
    ' Validate preconditions
    If Documents.Count = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "No documents open"
        Exit Sub
    End If
    
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Document parameter is Nothing"
        Exit Sub
    End If
    
    If Not GetProperty(pIsDocument) Then
        WriteLog 2, CurrentMod, PROC_NAME, "Not a Docent document - skipping border"
        Exit Sub
    End If
    
    If Doc.Sections.Count = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Document has no sections"
        Exit Sub
    End If
    
    ' Convert color parameter
    hasProjectColor = (Len(ProjectColorStr) > 0)
    If hasProjectColor Then
        colorValue = CLng(borderColor)
    End If
    
    ' Check if update is needed
    If Not BordersChangeNeeded(Doc, colorValue) Then
        WriteLog 1, CurrentMod, PROC_NAME, "Borders already correct - no update needed"
        Exit Sub
    End If
    
    ' Save current state
    wasProtected = (Doc.protectionType <> wdNoProtection)
    wasSaved = Doc.Saved
    
    WriteLog 1, CurrentMod, PROC_NAME, "Updating document borders"
    
    ' Unprotect for modification
    Unprotect Doc
    
    ' Disable auto-save during modification
    Doc.AutoSaveOn = False
    
    ' Apply or remove borders
    With Doc.Sections(1)
        ' Process all four borders: Top=-1, Left=-2, Bottom=-3, Right=-4
        For borderIndex = wdBorderTop To wdBorderRight Step -1
            
            With .Borders(borderIndex)
                If hasProjectColor Then
                    ' Apply colored border if not already correct
                    If .Color <> colorValue Or .LineStyle = wdLineStyleNone Then
                        .LineStyle = wdLineStyleSingle
                        .LineWidth = wdLineWidth450pt  ' 6 points
                        .Color = colorValue
                        borderApplied = True
                    End If
                Else
                    ' Remove border if present
                    If .LineStyle <> wdLineStyleNone Then
                        .LineStyle = wdLineStyleNone
                        borderApplied = True
                    End If
                End If
            End With
            
        Next borderIndex
        
        ' Apply to all sections if changes were made
        If borderApplied Then
            .Borders.ApplyPageBordersToAllSections
            WriteLog 1, CurrentMod, PROC_NAME, "Borders applied to all sections"
        End If
    End With
    
    ' Restore state
    If wasProtected Then Protect Doc
    Doc.Saved = wasSaved
    
    WriteLog 1, CurrentMod, PROC_NAME, "Border operation completed successfully"
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    ' Attempt to restore state even on error
    On Error Resume Next
    If wasProtected Then Protect Doc
    Doc.Saved = wasSaved
    On Error GoTo 0
End Sub

'=======================================================
' HELPER PROCEDURES
'=======================================================

'=======================================================
' Function: BordersChangeNeeded
' Purpose: Determine if document borders need updating
'
' Parameters:
'   Doc - Document to check
'   borderColor - Color to apply (Long)
'
' Returns:
'   True if borders need updating, False if already correct
'
' Description:
'   Checks first section borders against desired state:
'
'   If no project color (empty ProjectColorStr):
'     Returns True if any borders are present (need removal)
'     Returns False if borders are already absent
'
'   If project color exists:
'     Returns True if any border has wrong color or is missing
'     Returns False if all borders match the color
'
' Note:
'   Only checks first section - assumes all sections match
'   after using ApplyPageBordersToAllSections
'
' Error Handling:
'   Returns False on error (assume no change needed)
'=======================================================
Private Function BordersChangeNeeded(ByVal Doc As Document, _
                                     ByVal borderColor As Long) As Boolean
    Const PROC_NAME As String = "BordersChangeNeeded"
    
    Dim borderIndex As Long
    Dim currentBorder As Border
    Dim hasProjectColor As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Default to no change needed
    BordersChangeNeeded = False
    
    ' Validate document has sections
    If Doc.Sections.Count = 0 Then Exit Function
    
    ' Check if we have a project color
    hasProjectColor = (Len(ProjectColorStr) > 0)
    
    With Doc.Sections(1)
        ' Check all four borders
        For borderIndex = wdBorderTop To wdBorderRight Step -1
            Set currentBorder = .Borders(borderIndex)
            
            With currentBorder
                If hasProjectColor Then
                    ' Check if border needs to be applied or updated
                    If .Color <> borderColor Or .LineStyle = wdLineStyleNone Then
                        BordersChangeNeeded = True
                        Exit Function
                    End If
                Else
                    ' Check if border needs to be removed
                    If .LineStyle <> wdLineStyleNone Then
                        BordersChangeNeeded = True
                        Exit Function
                    End If
                End If
            End With
            
        Next borderIndex
    End With
    
    ' If we get here, no changes needed
    WriteLog 1, CurrentMod, PROC_NAME, "Borders are already correct"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    
    ' On error, assume no change needed (safe default)
    BordersChangeNeeded = False
End Function

'=======================================================
' END OF MODULE
'=======================================================
