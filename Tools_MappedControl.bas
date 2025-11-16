Attribute VB_Name = "Tools_MappedControl"
Option Explicit
'2024-11-21
'Example: (remove the ' when you copy the line to the immideate window)
'AddContentControlAndMap "Name","DocentMap",wdContentControlText
Sub AddContentControlAndMap(Optional ByVal ContentName As String, _
                            Optional ByVal XMLPatrName As String = "CustomMap", _
                            Optional ByVal ContentType As WdContentControlType)
    Dim oCustomPart As Office.CustomXMLPart
    For Each oCustomPart In ActiveDocument.CustomXMLParts
        If Mid(oCustomPart.XML, 2, Len(XMLPatrName)) = XMLPatrName Then Exit For
    Next
    If oCustomPart Is Nothing Then
        Set oCustomPart = ActiveDocument.CustomXMLParts.Add( _
                    "<" & XMLPatrName & "><" & ContentName & ">" & _
                    "</" & ContentName & "></" & XMLPatrName & ">")
    Else 'If InStr(oCustomPart.XML, "<" & ContentName & ">") = 0 Then
        On Error Resume Next
        oCustomPart.SelectSingleNode("//" & CleanCCName(ContentName)).Delete
        On Error GoTo 0
        oCustomPart.AddNode oCustomPart.SelectSingleNode("/" & XMLPatrName), CleanCCName(ContentName)
    End If
    Dim oCC As Word.ContentControl, OVal As String, OFormat As Range, oCCs As Collection
    Set oCCs = FindCCs(ContentName, ActiveDocument)
    If oCCs.Count = 0 Then
        Set oCC = ActiveDocument.ContentControls.Add(ContentType)
        oCC.Title = ContentName
        oCCs.Add oCC
    End If
    For Each oCC In oCCs
        Set OFormat = oCC.Range.FormattedText
        OVal = oCC.Range.text
        oCC.Appearance = IIf(InStr(ContentName, "Date"), wdContentControlBoundingBox, wdContentControlHidden)
        oCC.XMLMapping.SetMapping "/" & XMLPatrName & "/" & CleanCCName(ContentName) & "[1]"
        If Not OFormat Is Nothing Then
            On Error Resume Next
            oCC.Range.FormattedText = OFormat
            On Error GoTo 0
            oCC.Range.text = OVal
        End If
    Next
End Sub
Private Function CleanCCName(CCName As String) As String
    CleanCCName = Replace(Replace(CCName, " ", ""), "&", "n")
End Function
Function FindCC(ByVal ControlName As String, Optional Doc As Document, Optional ByVal XMLPatrName As String = "DocentIMS", _
                        Optional Mode As String = "CCs") As ContentControl
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Select Case Mode
    Case "CCs": Set FindCC = FindCCs(ControlName, Doc)(1)
    Case "XML": Set FindCC = FindCCXml(ControlName, Doc, XMLPatrName)
    Case Else
    End Select
End Function
Private Function FindCCXml(ByVal ControlName As String, Optional Doc As Document, _
                            Optional ByVal XMLPatrName As String = "DocentIMS") As CustomXMLNode
    Dim oCustomPart As Office.CustomXMLPart
    ControlName = Replace(Replace(ControlName, " ", ""), "&", "n")
    For Each oCustomPart In Doc.CustomXMLParts
        If Mid(oCustomPart.XML, 2, Len(XMLPatrName)) = XMLPatrName Then Exit For
    Next
    Set FindCCXml = oCustomPart.SelectSingleNode("//" & ControlName)
End Function
Private Function FindCCsInRange(Optional ByVal ControlName As String, Optional Rng As Range) As Collection
    Dim CC As ContentControl
    Set FindCCsInRange = New Collection
    For Each CC In Rng.ContentControls
        If Len(ControlName) Then
            Select Case ControlName
            Case CC.Tag, CC.Title: FindCCsInRange.Add CC
            End Select
        Else
            FindCCsInRange.Add CC
        End If
    Next
End Function
Function FindCCs(Optional ByVal ControlName As String, Optional Doc As Document) As Collection
    Dim Sec As Section, CC As ContentControl, HF As HeaderFooter, Shp As Shape
    If Doc Is Nothing Then Set Doc = ActiveDocument
    On Error Resume Next
    Set FindCCs = New Collection
    For Each CC In FindCCsInRange(ControlName, Doc.Range)
        FindCCs.Add CC
    Next
    For Each Shp In Doc.Shapes
        If Shp.TextFrame.HasText Then
            For Each CC In FindCCsInRange(ControlName, Shp.TextFrame.TextRange)
                FindCCs.Add CC
            Next
        End If
    Next
    For Each Sec In Doc.Sections
        For Each HF In Sec.Headers
            For Each CC In FindCCsInRange(ControlName, HF.Range)
                FindCCs.Add CC
            Next
        Next
        For Each HF In Sec.Footers
            For Each CC In FindCCsInRange(ControlName, HF.Range)
                FindCCs.Add CC
            Next
        Next
    Next
End Function
Sub SetControlsViaXML(ByVal ControlName As String, ByVal NewValue As String, Optional Doc As Document, _
                            Optional ByVal XMLPatrName As String = "DocentIMS")
    If Doc Is Nothing Then Set Doc = ActiveDocument
    FindCCXml(ControlName, Doc, XMLPatrName).text = NewValue
End Sub
Function GetControlsViaXML(ByVal ControlName As String, Optional Doc As Document, _
                        Optional ByVal XMLPatrName As String = "DocentIMS") As String
    If Doc Is Nothing Then Set Doc = ActiveDocument
    GetControlsViaXML = FindCCXml(ControlName, Doc, XMLPatrName).text
End Function
Function GetContentControl(Title As String, Optional ByRef Doc As Document)
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Dim CC As ContentControl
    For Each CC In FindCCs(Title, Doc)
        If CC.Type = wdContentControlCheckBox Then
            GetContentControl = CC.Checked
        Else
            If CC.PlaceholderText Is Nothing And Len(CC.Range.text) > 0 Then
                GetContentControl = CC.Range.text
            ElseIf CC.Range.text <> CC.PlaceholderText Then
                GetContentControl = CC.Range.text
            End If
        End If
    Next
End Function
Sub SetContentControl(Title As String, ByVal value As Variant, _
                            Optional ByRef Doc As Document)
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Dim CC As ContentControl
    For Each CC In FindCCs(Title, Doc)
        SetCC CC, value
    Next
End Sub
Private Sub SetCC(CC As ContentControl, value As Variant)
    If CC Is Nothing Then Exit Sub
    If CC.Type = wdContentControlCheckBox Then
        CC.Checked = value
    Else
        CC.Range.text = value
    End If
End Sub



