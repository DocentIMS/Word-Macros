Attribute VB_Name = "AT_MasterTemplate_mod"
Option Explicit
'Private MasterFName As String, MasterDoc As Document
'Sub UpdateTemplates()
'    Dim FName As String
'    MasterFName = GetFile("Browse to the master template file", Environ("Userprofile") & "\desktop", False, WordFiles)(1)
'    Set MasterDoc = Nothing
'End Sub
 Sub UpdateTemplate(Doc As Document, MasterFName As String)
    Dim MasterDoc As Document
'    If Doc Is Nothing Then Set Doc = Documents("The Meadows Monthly Board Meeting.docx")
'    MasterFName = "D:\Ongoing\23-06-15 - Wayne Glover (Word)\Meadows Main Template.docx"
    On Error Resume Next
    Set MasterDoc = Documents(GetFileName(MasterFName))
    MasterDoc.Close False
    Unprotect Doc
    On Error GoTo 0
'    On Error GoTo ex
    'Theme
    Doc.ApplyDocumentTheme MasterFName
    Set MasterDoc = Documents.Open(MasterFName) 'ActiveDocument
    'Styles
'    ClearStyles Doc ', MasterDoc
    Doc.CopyStylesFromTemplate MasterFName
    CopyStyles Doc, MasterDoc
    'Page Setup
    With Doc.PageSetup
        .PageWidth = MasterDoc.PageSetup.PageWidth
        .RightMargin = MasterDoc.PageSetup.RightMargin
        .LeftMargin = MasterDoc.PageSetup.LeftMargin
        .PageHeight = MasterDoc.PageSetup.PageHeight
        .TopMargin = MasterDoc.PageSetup.TopMargin
        .BottomMargin = MasterDoc.PageSetup.BottomMargin
        .HeaderDistance = MasterDoc.PageSetup.HeaderDistance
        .FooterDistance = MasterDoc.PageSetup.FooterDistance
    End With
    'Header Logo
    CopyImagesToDocument Doc, MasterDoc, True
'    'Content Controls
'    CopyCCToDocument
    Protect Doc
    MasterDoc.Close False
    Exit Sub
ex:
'    Stop
'    Resume
End Sub
Private Sub CopyStyles(Doc As Document, MasterDoc As Document)
    Dim Stl As Style, Flg As Boolean
    On Error Resume Next
    For Each Stl In Doc.Styles ': Stl.Delete: Next 'i = Doc.Styles.Count To 1 Step -1
        Flg = False
        Flg = MasterDoc.Styles(Stl.NameLocal).QuickStyle
        Stl.QuickStyle = Flg
'        If Not MasterDoc.Styles(Doc.Styles(i).NameLocal).InUse Then Doc.Styles(i).Delete
'        For j = 1 To MasterDoc.Styles.Count
'            If Not MasterDoc.Styles(Doc.Styles(i).NameLocal).InUse Then Doc.Styles(i).Delete
'        Next
'        If j > MasterDoc.Styles.Count Then Doc.Styles(i).Delete
    Next
End Sub
Private Sub ClearStyles(Doc As Document) ', MasterDoc As Document)
    Dim Stl As Style
    On Error Resume Next
    For Each Stl In Doc.Styles: Stl.Delete: Next 'i = Doc.Styles.Count To 1 Step -1
'        If Not MasterDoc.Styles(Stl.NameLocal).InUse Then Stl.Delete
''        If Not MasterDoc.Styles(Doc.Styles(i).NameLocal).InUse Then Doc.Styles(i).Delete
''        For j = 1 To MasterDoc.Styles.Count
''            If Not MasterDoc.Styles(Doc.Styles(i).NameLocal).InUse Then Doc.Styles(i).Delete
''        Next
''        If j > MasterDoc.Styles.Count Then Doc.Styles(i).Delete
'    Next
End Sub
Private Sub CopyImagesToDocument(Doc As Document, MasterDoc As Document, CCToo As Boolean)
    Dim SecNo As Long, HdFt As HeaderFooter, i As Long, j As Long, k As Long
    Dim srcShape As Shape
    Dim newShape As Shape
    On Error Resume Next
    'Clear Old Logos/Images
'    Doc.Windows(1).Activate
    For SecNo = 1 To Doc.Sections.Count
        For i = 1 To Doc.Sections(SecNo).Headers(1).Shapes.Count
            Doc.Sections(SecNo).Headers(1).Shapes(1).Delete
        Next
        For j = 1 To 3
            If CCToo Then
                For k = 1 To Doc.Sections(SecNo).Headers(j).Range.ContentControls.Count
                    Doc.Sections(SecNo).Headers(j).Range.ContentControls(1).Delete
                Next
                UpdateHdFtCC MasterDoc.Sections(1).Headers(j), Doc.Sections(SecNo).Headers(j)
                UpdateHdFtCC MasterDoc.Sections(1).Footers(j), Doc.Sections(SecNo).Footers(j)
            Else
                UpdateHdFtLogos MasterDoc.Sections(1).Headers(j), Doc.Sections(SecNo).Headers(j)
                UpdateHdFtLogos MasterDoc.Sections(1).Footers(j), Doc.Sections(SecNo).Footers(j)
            End If
        Next
    Next
End Sub
Private Sub UpdateHdFtCC(sHdFt As HeaderFooter, dHdFt As HeaderFooter)
    Dim i As Long ', c As Long
    dHdFt.Range.FormattedText = sHdFt.Range.FormattedText
    i = sHdFt.Range.Tables.Count
    For i = 1 To i
        dHdFt.Range.Tables(i).Range.ParagraphFormat.SpaceAfter = sHdFt.Range.Tables(i).Range.ParagraphFormat.SpaceAfter
        dHdFt.Range.Tables(i).Range.ParagraphFormat.SpaceBefore = sHdFt.Range.Tables(i).Range.ParagraphFormat.SpaceBefore
    Next
'    iMax = sHdFt.Range.ContentControls.Count
'    For i = 1 To iMax
'        With sHdFt.Range.ContentControls(i)
'            c = .Range.Start
'            .Copy
'        End With
'        Set Rng = dHdFt.Range
'        Rng.Collapse 1
'
'    Next
End Sub
Private Sub UpdateHdFtLogos(sHdFt As HeaderFooter, dHdFt As HeaderFooter)
    Dim ShIn As InlineShape, iMax As Long, i As Long, sShp As Shape, dShp As Shape, Rng As Range
    Dim Left As Single, Top As Single, Width As Single, Height As Single
    If dHdFt.LinkToPrevious Then Exit Sub
    iMax = sHdFt.Range.ShapeRange.Count
    For i = 1 To iMax
        With sHdFt.Range.ShapeRange(iMax - i + 1)
            Set Rng = sHdFt.Range
            .RelativeVerticalPosition = wdRelativeVerticalPositionInnerMarginArea
            Top = .Top
            Left = .Left
            Set ShIn = .ConvertToInlineShape
        End With
        Set Rng = dHdFt.Range
        Rng.Collapse 1
        Rng.FormattedText = ShIn.Range.FormattedText
        Set dShp = dHdFt.Range.InlineShapes(1).ConvertToShape
        Set dShp = dHdFt.Shapes(dHdFt.Shapes.Count)
        With dShp
            .RelativeVerticalPosition = wdRelativeVerticalPositionInnerMarginArea
            .Left = Left
            .Top = Top
        End With
        ShIn.Delete
    Next
End Sub

