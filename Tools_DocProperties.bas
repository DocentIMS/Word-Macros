Attribute VB_Name = "Tools_DocProperties"
Option Explicit
'ContentTypeProperties
Function GetDocProperty(ByVal pPName As String, Optional ByRef Doc As Document) As Variant
    If Documents.Count = 0 Then Exit Function
    On Error GoTo ex
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Select Case GetPropertyType(pPName, Doc)
    Case "custom": GetDocProperty = Doc.CustomDocumentProperties(pPName).value
    Case "builtin": GetDocProperty = Doc.BuiltInDocumentProperties(pPName).value
    End Select
ex:
End Function
Function GetCustomDocProperty(ByVal pPName As String, Optional ByRef Doc As Document) As Variant
    If Documents.Count = 0 Then Exit Function
    On Error GoTo ex
    If Doc Is Nothing Then Set Doc = ActiveDocument
    GetCustomDocProperty = Doc.CustomDocumentProperties(pPName).value
ex:
End Function
Function GetBuiltInDocProperty(ByVal pPName As String, Optional ByRef Doc As Document) As Variant
    If Documents.Count = 0 Then Exit Function
    On Error GoTo ex
    If Doc Is Nothing Then Set Doc = ActiveDocument
    GetBuiltInDocProperty = Doc.BuiltInDocumentProperties(pPName).value
ex:
End Function
Sub DeleteDocProperty(pPName As String, Optional ByRef Doc As Document)
    If Doc Is Nothing Then Set Doc = ActiveDocument
    On Error Resume Next
    Doc.BuiltInDocumentProperties(pPName).Delete
    Doc.CustomDocumentProperties(pPName).Delete
End Sub
Sub SetDocProperty(ByVal pPName As String, ByVal PPValue As Variant, Optional ByRef Doc As Document, Optional ByVal PPType As Office.MsoDocProperties = 4)
    If Doc Is Nothing Then Set Doc = ActiveDocument
    On Error Resume Next
    DeleteDocProperty pPName, Doc
    Select Case GetPropertyType(pPName, Doc)
    Case "custom": Doc.CustomDocumentProperties(pPName).value = PPValue
    Case "builtin": Doc.BuiltInDocumentProperties(pPName).value = PPValue
    Case Else: Doc.CustomDocumentProperties.Add pPName, False, PPType, PPValue
    End Select
End Sub
Function GetPropertyType(ByVal pPName As String, Optional ByRef Doc As Document) As String
    If Doc Is Nothing Then Set Doc = ActiveDocument
    On Error Resume Next
    Doc.CustomDocumentProperties(pPName).Name = pPName
    If Err.Number = 0 Then GetPropertyType = "custom": Exit Function
    Err.Clear
    Doc.BuiltInDocumentProperties(pPName).Name = pPName
    If Err.Number = 0 Then GetPropertyType = "builtin": Exit Function
End Function
