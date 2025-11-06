Attribute VB_Name = "AC_CreateBorder_mod"
Option Explicit
Const CurrentMod = "AC_CreateBorder_mod"
'Public SelectedProjectName As String
Private Function BordersChangeNeeded(Doc As Document, Color) As Boolean
    Dim i As Long
    With Doc.Sections(1)
        For i = -4 To -1
            With .Borders(i)
                If Len(ProjectColorStr) = 0 Then
                    BordersChangeNeeded = .LineStyle = 1
'                    If .LineStyle = 1 Then .LineStyle = 0: Chnged = True
                ElseIf .Color <> Color Or .LineStyle = 0 Then
                    BordersChangeNeeded = True
'                    .LineStyle = 1
'                    .LineWidth = 48
'                    .Color = Color
'                    Chnged = True
                Else
                    Exit Function
                End If
            End With
        Next
'        If Chnged Then .Borders.ApplyPageBordersToAllSections
    End With
End Function
Sub CreateBorder(Doc As Document, Color)
    Dim i As Long, s As Long, Svd As Boolean, Chnged As Boolean, WasProtected As Boolean
    WriteLog 1, CurrentMod, "CreateBorder", Color
    If Documents.Count = 0 Then Exit Sub
    If Not GetProperty(pIsDocument) Then Exit Sub
    If Not BordersChangeNeeded(Doc, Color) Then Exit Sub
    WasProtected = Doc.ProtectionType <> wdNoProtection
    Unprotect Doc
    Svd = Doc.Saved
    Doc.AutoSaveOn = False
    With Doc.Sections(1)
        For i = -4 To -1
            With .Borders(i)
                If Len(ProjectColorStr) = 0 Then
                    If .LineStyle = 1 Then .LineStyle = 0: Chnged = True
                ElseIf .Color <> Color Or .LineStyle = 0 Then
                    .LineStyle = 1
                    .LineWidth = 48
                    .Color = Color
                    Chnged = True
                Else
                    GoTo ex
                End If
            End With
        Next
        If Chnged Then .Borders.ApplyPageBordersToAllSections
    End With
ex:
    If WasProtected Then Protect Doc
    Doc.Saved = Svd
End Sub
