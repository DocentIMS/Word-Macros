Attribute VB_Name = "AB_CommonTools"
Option Explicit
Option Compare Text
Public Enum xlNameMode
    VariableName = 0 'DEFAULT
    SheetName = 2 ^ 0
    SectionName = 2 ^ 1
End Enum
Function ParseDate(Dt As String, Frmt As String) As Date
    Dim Seps(0 To 2) As String, i As Long, j As Long
    Dim Yr As Integer, Mo As Integer, Da As Integer, Hr As Integer, Mi As Integer, Se As Integer
    Dim Frmts() As String
    Dim Dts() As String
    On Error GoTo ex
    If Len(Dt) = 0 Then Exit Function
    i = 1
    Do Until i > Len(Frmt)
        Select Case Mid$(Frmt, i, 1)
        Case "y", "m", "d"
            If Len(Seps(0)) = 0 Then
                j = i
                Do Until j > Len(Frmt)
                    Select Case Mid$(Frmt, j, 1)
                    Case Mid$(Frmt, i, 1)
                    Case Else: Seps(0) = Mid$(Frmt, j, 1): Exit Do
                    End Select
                    j = j + 1
                Loop
            End If
        Case "h", "n", "s"
            If Len(Seps(1)) = 0 Then
                j = i
                Do Until j > Len(Frmt)
                    Select Case Mid$(Frmt, j, 1)
                    Case Mid$(Frmt, i, 1)
                    Case Else: Seps(1) = Mid$(Frmt, j, 1): Exit Do
                    End Select
                    j = j + 1
                Loop
            End If
        End Select
        If Len(Seps(1)) > 0 And Len(Seps(0)) > 0 Then Exit Do
        i = i + 1
    Loop
    Seps(2) = ""
    j = 0
    Do
        If Not (j > UBound(Split(Frmt)) Or j > UBound(Split(Dt))) Then
            Frmts = Split(Split(Frmt)(j), Seps(j))
            Dts = Split(Split(Dt)(j), Seps(j))
            For i = LBound(Frmts) To UBound(Frmts)
                Select Case Frmts(i)
                Case "yy", "yyyy": Yr = Dts(i)
                Case "m", "mm": Mo = Dts(i)
                Case "d", "dd": Da = Dts(i)
                Case "h", "hh": Hr = Dts(i)
                Case "n", "nn": Mi = Dts(i)
                Case "s", "ss": Se = Dts(i)
                Case "AM/PM": If Dts(i) = "PM" Then Hr = Hr + 12
                End Select
            Next
        End If
        j = j + 1
        If j > UBound(Dts) Then Exit Do
        If InStr(Frmt, " ") = 0 Or InStr(Dt, " ") = 0 Then Exit Do
    Loop
    ParseDate = DateSerial(Yr, Mo, Da) + TimeSerial(Hr, Mi, Se)
    Exit Function
ex:
'    Stop
'    Resume
End Function
Function IsGoodResponse(Resp, Optional EmptyIsOkay As Boolean = False) As Boolean
    If TypeName(Resp) = "Empty" Then
        IsGoodResponse = EmptyIsOkay 'False
    ElseIf TypeName(Resp) = "String" Then
        IsGoodResponse = Not Resp Like "Failed *" And (Len(Resp) > 0 Or EmptyIsOkay)
'        If Not IsGoodResponse Then WriteLog 3, CurrentMod, "IsGoodResponse", Resp
    ElseIf TypeName(Resp) = "Collection" Then
        IsGoodResponse = True 'Resp.Count > 0
    ElseIf InStr(TypeName(Resp), "()") Then
        Select Case GetDims(Resp)
        Case 1
            IsGoodResponse = LBound(Resp) <> UBound(Resp)
'            IsGoodResponse = IsGoodResponse(Resp(LBound(Resp)))
        Case 0
            IsGoodResponse = EmptyIsOkay
        Case Else
            IsGoodResponse = LBound(Resp) <> UBound(Resp) Or LBound(Resp, 2) <> UBound(Resp, 2)
'            IsGoodResponse = IsGoodResponse(Resp(LBound(Resp), LBound(Resp, 2)))
        End Select
    ElseIf Resp Is Nothing Then
        IsGoodResponse = EmptyIsOkay
    Else
        Select Case Resp.StatusCode
        Case 200, 201, 204: IsGoodResponse = True
        Case Else: IsGoodResponse = False
        End Select
'        IsGoodResponse = (Not Resp Is Nothing) Or EmptyIsOkay
    End If
End Function
Function GetActiveFName(Doc As Document) As String
    If Documents.Count = 0 Or Doc Is Nothing Then
        GetActiveFName = "Nothing"
        Exit Function
    End If
    On Error Resume Next
    If GetFileName(Doc.Name, False) = Doc.Name Then GetActiveFName = "Blank": Exit Function
    GetActiveFName = Doc.Name
End Function
Function GetTableByTitle(Title As String, Optional Doc As Document) As Table
    If Doc Is Nothing Then Set Doc = ActiveDocument
    Dim Tbl As Table
    For Each Tbl In Doc.Tables
        If Tbl.Title = Title Then Set GetTableByTitle = Tbl: Exit Function
    Next
End Function
Function ToStr(Dict As Collection) As String
    Dim i As Long
    For i = 1 To Dict.Count
        ToStr = ToStr & Dict(i) & Seperator
    Next
    If Len(ToStr) > 0 Then ToStr = Left$(ToStr, Len(ToStr) - Len(Seperator))
End Function
Function CollExists(Coll As Collection, sName As String) As Boolean
    Dim i As Long
    For i = 1 To Coll.Count
        If Coll(i) = sName Then CollExists = True: Exit Function
    Next
End Function
Function ContainsWord(ByVal s As String, Wrd As String, Optional ByVal i As Long) As Boolean
    If i = 0 Then i = InStr(s, Wrd)
    If i = 0 Or i > Len(s) Then Exit Function
    If i <> 1 Then If Mid(s, i - 1, 1) <> " " Then i = InStrRev(s, " ", i) + 1
    s = Mid(s, i, Len(s))
    ContainsWord = s = Wrd _
                    Or s Like Wrd & " *" _
                    Or s Like Wrd & "?" _
                    Or s Like Wrd & "." _
                    Or s Like "* " & Wrd & " *" _
                    Or s Like "* " & Wrd _
                    Or s Like "* " & Wrd & "." _
                    Or s Like "* " & Wrd & "?"
End Function
Function ClearFromStr(ByVal s As String, Optional Chars) As String
    Dim i As Long
    Chars = Chars & " "
    For i = 1 To Len(Chars)
        s = Replace(s, Mid(Chars, i, 1), "")
    Next
    ClearFromStr = s
End Function
'Function GetStateFromTrn(ByVal TrnURL As String) As String
'    'Here
'    Dim ss() As String
'    ss = Split(TrnURL, "_")
'    TrnURL = Replace(ss(UBound(ss)), "-", " ")
'    TrnURL = StrConv(TrnURL, vbProperCase)
'    GetStateFromTrn = TrnURL
'End Function
Function IsArrayAllocated(Arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(Arr) And _
                       Not IsError(LBound(Arr, 1)) And _
                       LBound(Arr, 1) <= UBound(Arr, 1)
End Function
Function CellText(ByVal Cell) As String
    Dim i As Long, Txt As String
    Select Case TypeName(Cell)
    Case "Cell": Txt = Cell.Range.text
    Case "Range": Txt = Cell.Txt
    Case "String": Txt = Cell
    End Select
    For i = Len(Txt) To 1 Step -1
    Select Case Asc(Mid$(Txt, i, 1))
    Case 10, 13, 7: Txt = Left$(Txt, Len(Txt) - 1)
    Case Else: Exit For
    End Select
    Next
    CellText = Txt
End Function
Sub HyperlinkWord(Rng As Range, Wrd As String, URL As String)
    With Rng.Find
        .text = Wrd
        .Wrap = wdFindStop
        If .Execute Then Rng.Document.Hyperlinks.Add Rng, URL, , "Click to join"
    End With
End Sub
Sub RemoveSpecificCommandButton(BtnName As String, Doc As Document)
    Dim i As Long, Shp As Shape
    For i = 1 To Doc.InlineShapes.Count
        If Doc.InlineShapes(i).Type = wdInlineShapeOLEControlObject Then
'            If Doc.InlineShapes(i).OLEFormat.Object.Name = BtnName Then ' Replace "CommandButton1" with the actual name of your button
            If Doc.InlineShapes(i).AlternativeText = BtnName Then ' Replace "CommandButton1" with the actual name of your button
                Set Shp = Doc.InlineShapes(i).ConvertToShape 'Doc.InlineShapes(i).Delete
                Shp.Delete
                Exit For ' Exit the loop once the button is found and deleted
            End If
        End If
    Next
End Sub
'@EntryPoint
Function CleanName(Name As String, Optional NameMode As xlNameMode, Optional ReplaceWithChr As Boolean) As String
    'ReplaceWithChr = False : ""
    Dim CharStr As String, i As Long
    CleanName = Name
    Select Case NameMode
    Case VariableName: CharStr = ":?[]/\*+-()&. ^$%"
    Case SheetName: CharStr = ":?[]/\*"
    Case SectionName
        i = 1
        Do Until i > Len(CleanName)
            Select Case Asc(Mid(CleanName, i, 1))
            Case 97 To 122, 65 To 90, 95, 46, 48 To 57, 45, 32
                '"[a-z][A-Z]_.[0-9]- "
            Case Else
                If ReplaceWithChr Then
                    CleanName = Replace(CleanName, Mid(CleanName, i, 1), "chr_" & Asc(Mid(CleanName, i, 1)) & "_")
                Else
                    CleanName = Replace(CleanName, Mid(CleanName, i, 1), vbNullString)
                End If
                i = i - 1
            End Select
            i = i + 1
        Loop
    End Select
    If NameMode <> SectionName Then
        For i = 1 To Len(CharStr)
            If ReplaceWithChr Then
                CleanName = Replace(CleanName, Mid(CharStr, i, 1), "chr_" & Asc(Mid(CharStr, i, 1)) & "_")
            Else
                CleanName = Replace(CleanName, Mid(CharStr, i, 1), vbNullString)
            End If
        Next
    End If
End Function


