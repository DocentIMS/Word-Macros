Attribute VB_Name = "Tools_Utilities"
Option Explicit
Option Compare Text

'=======================================================
' Module: Tools_Utilities
' Purpose: Common utility tools and helper functions
' Author: IMPROVED - November 2025
' Version: 3.0
'
' IMPROVEMENTS APPLIED:
'   ✓ #2: Fixed Inconsistent Error Handling
'       - Replaced all "On Error Resume Next" with structured handlers
'       - Fixed ParseDate's "On Error GoTo ex" pattern
'       - Added proper error restoration
'       - Added logging for all error conditions
'       - Return default/safe values on error
'
' Description:
'   Provides common utility functions for data validation,
'   string manipulation, date parsing, and document operations.
'
' Change Log:
'   v3.0 - Nov 2025 - Applied improvement #2
'       * Fixed all error handling patterns
'       * Improved logging throughout
'       * Added function documentation
'   v2.0 - Previous version
'   v1.0 - Original version
'=======================================================

Private Const CurrentMod As String = "Tools_Utilities"

'=======================================================
' ENUMERATIONS
'=======================================================

Public Enum xlNameMode
    VariableName = 0 'DEFAULT
    SheetName = 2 ^ 0
    SectionName = 2 ^ 1
End Enum

'=======================================================
' IMPROVEMENT #2: Completely refactored with proper error handling
'=======================================================
Function ParseDate(Dt As String, Frmt As String) As Date
    Const PROC_NAME As String = "ParseDate"
    Dim Seps(0 To 2) As String
    Dim i As Long
    Dim j As Long
    Dim Yr As Integer
    Dim Mo As Integer
    Dim Da As Integer
    Dim Hr As Integer
    Dim Mi As Integer
    Dim Se As Integer
    Dim Frmts() As String
    Dim Dts() As String
    Dim result As Date
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Len(Dt) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty date string provided"
        ParseDate = CDate(0)
        Exit Function
    End If
    
    If Len(Frmt) = 0 Then
        WriteLog 2, CurrentMod, PROC_NAME, "Empty format string provided"
        ParseDate = CDate(0)
        Exit Function
    End If
    
    ' Find date separators
    i = 1
    Do Until i > Len(Frmt)
        Select Case Mid$(Frmt, i, 1)
            Case "y", "m", "d"
                If Len(Seps(0)) = 0 Then
                    j = i
                    Do Until j > Len(Frmt)
                        Select Case Mid$(Frmt, j, 1)
                            Case Mid$(Frmt, i, 1)
                                ' Same character, continue
                            Case Else
                                Seps(0) = Mid$(Frmt, j, 1)
                                Exit Do
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
                                ' Same character, continue
                            Case Else
                                Seps(1) = Mid$(Frmt, j, 1)
                                Exit Do
                        End Select
                        j = j + 1
                    Loop
                End If
        End Select
        If Len(Seps(1)) > 0 And Len(Seps(0)) > 0 Then Exit Do
        i = i + 1
    Loop
    
    ' Parse date/time components
    Seps(2) = ""
    j = 0
    Do
        If Not (j > UBound(Split(Frmt)) Or j > UBound(Split(Dt))) Then
            Frmts = Split(Split(Frmt)(j), Seps(j))
            Dts = Split(Split(Dt)(j), Seps(j))
            
            For i = LBound(Frmts) To UBound(Frmts)
                Select Case Frmts(i)
                    Case "yy", "yyyy"
                        Yr = CInt(Dts(i))
                    Case "m", "mm"
                        Mo = CInt(Dts(i))
                    Case "d", "dd"
                        Da = CInt(Dts(i))
                    Case "h", "hh"
                        Hr = CInt(Dts(i))
                    Case "n", "nn"
                        Mi = CInt(Dts(i))
                    Case "s", "ss"
                        Se = CInt(Dts(i))
                    Case "AM/PM"
                        If Dts(i) = "PM" Then Hr = Hr + 12
                End Select
            Next
        End If
        j = j + 1
        If j > UBound(Dts) Then Exit Do
        If InStr(Frmt, " ") = 0 Or InStr(Dt, " ") = 0 Then Exit Do
    Loop
    
    ' Create date result
    result = DateSerial(Yr, Mo, Da) + TimeSerial(Hr, Mi, Se)
    ParseDate = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & _
             " (Dt: '" & Dt & "', Frmt: '" & Frmt & "')"
    ' Return zero date on error
    ParseDate = CDate(0)
End Function

'=======================================================
' IMPROVEMENT #2: Added comprehensive error handling
'=======================================================
Function IsGoodResponse(Resp, Optional EmptyIsOkay As Boolean = False) As Boolean
    Const PROC_NAME As String = "IsGoodResponse"
    Dim responseType As String
    
    On Error GoTo ErrorHandler
    
    responseType = TypeName(Resp)
    
    Select Case responseType
        Case "Empty"
            IsGoodResponse = EmptyIsOkay
            
        Case "String"
            IsGoodResponse = Not Resp Like "Failed *" And (Len(Resp) > 0 Or EmptyIsOkay)
            
        Case "Collection"
            IsGoodResponse = True
            
        Case "Nothing"
            IsGoodResponse = EmptyIsOkay
            
        Case Else
            ' Check if it's an array
            If InStr(responseType, "()") > 0 Then
                Select Case GetDims(Resp)
                    Case 1
                        IsGoodResponse = LBound(Resp) <> UBound(Resp)
                    Case 0
                        IsGoodResponse = EmptyIsOkay
                    Case Else
                        IsGoodResponse = LBound(Resp) <> UBound(Resp) Or _
                                       LBound(Resp, 2) <> UBound(Resp, 2)
                End Select
            ElseIf Resp Is Nothing Then
                IsGoodResponse = EmptyIsOkay
            Else
                ' Assume it's a WebResponse object
                Select Case Resp.StatusCode
                    Case 200, 201, 204
                        IsGoodResponse = True
                    Case Else
                        IsGoodResponse = False
                End Select
            End If
    End Select
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & _
             " (Type: " & TypeName(Resp) & ")"
    IsGoodResponse = EmptyIsOkay
End Function

'=======================================================
' IMPROVEMENT #2: Fixed On Error Resume Next usage
'=======================================================
Function GetActiveFName(Doc As Document) As String
    Const PROC_NAME As String = "GetActiveFName"
    Dim fileName As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Documents.Count = 0 Or Doc Is Nothing Then
        GetActiveFName = "Nothing"
        Exit Function
    End If
    
    ' Get file name
    fileName = GetFileName(Doc.Name, False)
    
    If fileName = Doc.Name Then
        GetActiveFName = "Blank"
    Else
        GetActiveFName = Doc.Name
    End If
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    GetActiveFName = "Nothing"
End Function

'=======================================================
' IMPROVEMENT #2: Added proper error handling
'=======================================================
Function GetTableByTitle(Title As String, Optional Doc As Document) As Table
    Const PROC_NAME As String = "GetTableByTitle"
    Dim Tbl As Table
    
    On Error GoTo ErrorHandler
    
    ' Use active document if none specified
    If Doc Is Nothing Then Set Doc = ActiveDocument
    
    ' Validate document
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "No document available"
        Set GetTableByTitle = Nothing
        Exit Function
    End If
    
    ' Search for table by title
    For Each Tbl In Doc.Tables
        If Tbl.Title = Title Then
            Set GetTableByTitle = Tbl
            Exit Function
        End If
    Next
    
    ' Not found
    WriteLog 2, CurrentMod, PROC_NAME, "Table not found: " & Title
    Set GetTableByTitle = Nothing
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & " (Title: " & Title & ")"
    Set GetTableByTitle = Nothing
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function ToStr(Dict As Collection) As String
    Const PROC_NAME As String = "ToStr"
    Dim i As Long
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Dict Is Nothing Then
        ToStr = ""
        Exit Function
    End If
    
    For i = 1 To Dict.Count
        result = result & Dict(i) & Seperator
    Next
    
    ' Remove trailing separator
    If Len(result) > 0 Then
        result = Left$(result, Len(result) - Len(Seperator))
    End If
    
    ToStr = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ToStr = ""
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function CollExists(Coll As Collection, sName As String) As Boolean
    Const PROC_NAME As String = "CollExists"
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    ' Validate input
    If Coll Is Nothing Then
        CollExists = False
        Exit Function
    End If
    
    For i = 1 To Coll.Count
        If Coll(i) = sName Then
            CollExists = True
            Exit Function
        End If
    Next
    
    CollExists = False
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    CollExists = False
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function ContainsWord(ByVal s As String, Wrd As String, Optional ByVal i As Long) As Boolean
    Const PROC_NAME As String = "ContainsWord"
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Len(s) = 0 Or Len(Wrd) = 0 Then
        ContainsWord = False
        Exit Function
    End If
    
    ' Find word position
    If i = 0 Then i = InStr(s, Wrd)
    If i = 0 Or i > Len(s) Then
        ContainsWord = False
        Exit Function
    End If
    
    ' Ensure word boundary at start
    If i <> 1 Then
        If Mid(s, i - 1, 1) <> " " Then
            i = InStrRev(s, " ", i) + 1
        End If
    End If
    
    ' Extract substring from position
    s = Mid(s, i, Len(s))
    
    ' Check if word matches with various delimiters
    ContainsWord = s = Wrd _
                Or s Like Wrd & " *" _
                Or s Like Wrd & "?" _
                Or s Like Wrd & "." _
                Or s Like "* " & Wrd & " *" _
                Or s Like "* " & Wrd _
                Or s Like "* " & Wrd & "." _
                Or s Like "* " & Wrd & "?"
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ContainsWord = False
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function ClearFromStr(ByVal s As String, Optional Chars) As String
    Const PROC_NAME As String = "ClearFromStr"
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    If IsMissing(Chars) Then Chars = ""
    Chars = Chars & " "
    
    For i = 1 To Len(Chars)
        s = Replace(s, Mid(Chars, i, 1), "")
    Next
    
    ClearFromStr = s
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    ClearFromStr = s
End Function

'=======================================================
' IMPROVEMENT #2: Fixed On Error Resume Next usage
'=======================================================
Function IsArrayAllocated(Arr As Variant) As Boolean
    Const PROC_NAME As String = "IsArrayAllocated"
    Dim lowerBound As Long
    Dim upperBound As Long
    
    On Error GoTo ErrorHandler
    
    ' Check if it's an array
    If Not IsArray(Arr) Then
        IsArrayAllocated = False
        Exit Function
    End If
    
    ' Try to get bounds
    lowerBound = LBound(Arr, 1)
    upperBound = UBound(Arr, 1)
    
    ' If we got here without error, array is allocated
    IsArrayAllocated = (lowerBound <= upperBound)
    Exit Function
    
ErrorHandler:
    ' Error getting bounds means array not allocated
    IsArrayAllocated = False
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function cellText(ByVal Cell) As String
    Const PROC_NAME As String = "CellText"
    Dim i As Long
    Dim Txt As String
    Dim cellType As String
    
    On Error GoTo ErrorHandler
    
    ' Get cell text based on type
    cellType = TypeName(Cell)
    
    Select Case cellType
        Case "Cell"
            Txt = Cell.Range.text
        Case "Range"
            Txt = Cell.text
        Case "String"
            Txt = Cell
        Case Else
            Txt = CStr(Cell)
    End Select
    
    ' Remove trailing control characters
    For i = Len(Txt) To 1 Step -1
        Select Case Asc(Mid$(Txt, i, 1))
            Case 10, 13, 7  ' LF, CR, BEL
                Txt = Left$(Txt, Len(Txt) - 1)
            Case Else
                Exit For
        End Select
    Next
    
    cellText = Txt
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description & " (Type: " & TypeName(Cell) & ")"
    cellText = ""
End Function

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Sub HyperlinkWord(Rng As Range, Wrd As String, URL As String)
    Const PROC_NAME As String = "HyperlinkWord"
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Rng Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Range is Nothing"
        Exit Sub
    End If
    
    If Len(Wrd) = 0 Or Len(URL) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Word or URL is empty"
        Exit Sub
    End If
    
    ' Find and hyperlink the word
    With Rng.Find
        .text = Wrd
        .Wrap = wdFindStop
        
        If .Execute Then
            Rng.Document.Hyperlinks.Add Rng, URL, , "Click to join"
            WriteLog 1, CurrentMod, PROC_NAME, "Hyperlinked word: " & Wrd
        Else
            WriteLog 2, CurrentMod, PROC_NAME, "Word not found: " & Wrd
        End If
    End With
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Sub RemoveSpecificCommandButton(BtnName As String, Doc As Document)
    Const PROC_NAME As String = "RemoveSpecificCommandButton"
    Dim i As Long
    Dim Shp As Shape
    
    On Error GoTo ErrorHandler
    
    ' Validate inputs
    If Doc Is Nothing Then
        WriteLog 3, CurrentMod, PROC_NAME, "Document is Nothing"
        Exit Sub
    End If
    
    If Len(BtnName) = 0 Then
        WriteLog 3, CurrentMod, PROC_NAME, "Button name is empty"
        Exit Sub
    End If
    
    ' Search for and remove button
    For i = 1 To Doc.InlineShapes.Count
        If Doc.InlineShapes(i).Type = wdInlineShapeOLEControlObject Then
            If Doc.InlineShapes(i).AlternativeText = BtnName Then
                Set Shp = Doc.InlineShapes(i).ConvertToShape
                Shp.Delete
                WriteLog 1, CurrentMod, PROC_NAME, "Removed button: " & BtnName
                Exit For
            End If
        End If
    Next
    Exit Sub
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
End Sub

'=======================================================
' IMPROVEMENT #2: Added error handling
'=======================================================
Function CleanName(Name As String, Optional NameMode As xlNameMode, Optional ReplaceWithChr As Boolean) As String
    Const PROC_NAME As String = "CleanName"
    Dim CharStr As String
    Dim i As Long
    Dim result As String
    Dim charCode As Long
    
    On Error GoTo ErrorHandler
    
    result = Name
    
    Select Case NameMode
        Case VariableName
            CharStr = ":?[]/\*+-()&. ^$%"
        Case SheetName
            CharStr = ":?[]/\*"
        Case SectionName
            i = 1
            Do Until i > Len(result)
                charCode = Asc(Mid(result, i, 1))
                
                ' Check if character is allowed: a-z, A-Z, _, ., 0-9, -, space
                If Not ((charCode >= 97 And charCode <= 122) Or _
                       (charCode >= 65 And charCode <= 90) Or _
                       charCode = 95 Or charCode = 46 Or _
                       (charCode >= 48 And charCode <= 57) Or _
                       charCode = 45 Or charCode = 32) Then
                    
                    If ReplaceWithChr Then
                        result = Replace(result, Mid(result, i, 1), "chr_" & charCode & "_")
                    Else
                        result = Replace(result, Mid(result, i, 1), vbNullString)
                    End If
                    i = i - 1
                End If
                i = i + 1
            Loop
    End Select
    
    ' Replace invalid characters for non-SectionName modes
    If NameMode <> SectionName Then
        For i = 1 To Len(CharStr)
            If ReplaceWithChr Then
                result = Replace(result, Mid(CharStr, i, 1), "chr_" & Asc(Mid(CharStr, i, 1)) & "_")
            Else
                result = Replace(result, Mid(CharStr, i, 1), vbNullString)
            End If
        Next
    End If
    
    CleanName = result
    Exit Function
    
ErrorHandler:
    WriteLog 3, CurrentMod, PROC_NAME, _
             "Error " & Err.Number & ": " & Err.Description
    CleanName = Name  ' Return original on error
End Function

'=======================================================
' END OF MODULE
'=======================================================
