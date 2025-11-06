Attribute VB_Name = "AD_CollectWords_mod"
Option Explicit
Option Compare Text
Private Const CurrentMod = "CollectWords_mod"
Private Arr() As Variant, i As Long
Private ExcelApp As Object, rExcel As Long, OPath As String
'Private SDoc As Document
Private Const MaxHdrLvl = -1 '0 = All, -2 = Docent Only
'Private Const Set_SearchMode As Long
'1- The code finds the nearest bold ONLY
'2- The code only finds any "heading" styles
'3- The code finds the nearest bold OR the nearest style named "heading"

Sub CollectWords()
    WriteLog 1, CurrentMod, "CollectWords"
    Dim x As Long, Msg As String
    Dim t As Single, ts As String
    ReDim Arr(1 To 6, 1 To 1)
    i = 0
    t = Timer
    rExcel = 0
    If Not ExcelApp Is Nothing Then
        On Error Resume Next
        ExcelApp.Quit
        Set ExcelApp = Nothing
        On Error GoTo 0
    End If
    System.Cursor = wdCursorWait: DoEvents
    DoEvents
    
'    Set SDoc = ActiveDocument
    
    If mWillShall Then
        WriteLog 1, CurrentMod, "CollectWords", "Creating Will Shall Output Folder"
        If Right$(Set_Odir, 1) <> "\" Then Set_Odir = Set_Odir & "\"
        OPath = Set_Odir & "Docent Command Statements Results\"
        CreateDir OPath
        If Set_Odir <> Environ$("UserProfile") & "\Desktop\" Then CreateShortcut OPath, Environ$("UserProfile") & "\Desktop"
        If Set_Odir <> Environ$("UserProfile") & "\OneDrive\Desktop\" Then CreateShortcut OPath, Environ$("UserProfile") & "\OneDrive\Desktop"
        ProgressBar.Reset
        ProgressBar.BarsCount = 2
        ProgressBar.HideApplication = True
        ProgressBar.Dom(1) = 2 + Wrds.Count
        ProgressBar.Progress 1, "Gathering sections...", 0
        ProgressBar.Show
        FindSearchRange
    End If
    
    For x = 1 To Wrds.Count
        If mWillShall Then
            If ProgressBar.Progress(1, "Collecting """ & Wrds.Item(x).Wrd & """ statements...") Then
                MsgBox "Cancelled by user.", vbExclamation, ""
                GoTo can
            End If
        Else
            If ProgressBar.Spin(, "Coloring """ & Wrds.Item(x).Wrd & """ words...") Then
                MsgBox "Cancelled by user.", vbExclamation, ""
                GoTo can
            End If
        End If
        If Set_Cancelled Then GoTo can
        CollectWord x
    Next
    SDoc.SaveAs2 OPath & SDoc.Name
    If mWillShall Then
        If ProgressBar.Progress(1, "Exporting to Excel...") Then
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        toExcel Arr, SDoc.Name 'OPath &
        If ProgressBar.Canceled Then GoTo can
        If Set_Cancelled Then GoTo can
        System.Cursor = wdCursorNormal: DoEvents
        DoEvents
        For x = 1 To Wrds.Count
            Msg = Msg & vbNewLine & Wrds.Item(x).Count & " """ & Wrds.Item(x).Wrd & """ Statements were found"""
            If x > 1 Then Wrds.Item(1).Count = Wrds.Item(1).Count + Wrds.Item(x).Count
        Next
        t = Timer - t
        If t \ 60 > 0 Then
            ts = t \ 60 & ":" & (t Mod 60) \ 1 & " minutes."
        Else
            ts = (t Mod 60) \ 1 & " seconds."
        End If
        ProgressBar.Finished
        WriteLog 1, CurrentMod, "CollectWords", Wrds.Item(1).Count & " Statements Exported"
        If Wrds.Item(1).Count = rExcel Then
            frmMsgBox.Display Array("Process Completed" & vbNewLine & vbNewLine & Msg & vbNewLine & _
                            "-----------------------------------------" & vbNewLine & _
                            Wrds.Item(1).Count & " Total number of statements to track." & vbNewLine & _
                            rExcel & " Entries made in Excel" & vbNewLine, _
                            "-----------------------------------------" & vbNewLine & _
                            "Total processing time: " & ts), "Open Excel", , "Success", Array(0, 0, 0) '3057486
        Else
            frmMsgBox.Display Array("Process Completed" & vbNewLine & vbNewLine & Msg & vbNewLine & _
                            "-----------------------------------------" & vbNewLine & _
                            Wrds.Item(1).Count & " Total number of statements to track." & vbNewLine & _
                            rExcel & " Entries made in Excel" & vbNewLine, _
                            "-----------------------------------------" & vbNewLine & _
                            "Total processing time: " & ts), "Open Excel", Critical, "Failed", Array(0, 0, 0) '3057486
        End If
'        frmMsgBox.Display Array("Process Completed" & vbNewLine & vbNewLine & Msg & vbNewLine & _
'                "-----------------------------------------" & vbNewLine & _
'                Wrds.Item(1).Count & " Total number of statements to track." & vbNewLine & _
'                rExcel & " Entries made in Excel" & vbNewLine, _
'                IIf(Wrds.Item(1).Count = rExcel, "Success!", vbNewLine), _
'                 vbNewLine & _
'                "-----------------------------------------" & vbNewLine & _
'                "Total processing time: " & ts), "Open Excel)", , "Success", Array(0, 3057486, 0)
'        MsgBox Msg, , "Success"
can:
        On Error Resume Next
        Unload ProgressBar
        ExcelApp.Visible = True
        CloseLog
    End If
'    Debug.Print Timer - t
End Sub
Sub CollectWord(WrdNo As Long)
    Dim Wrd As String, clr As Long, x As Long, mSOW As SOW, PNo As Long
    Wrd = Wrds.Item(WrdNo).Wrd
    WriteLog 1, CurrentMod, "CollectWord", "Coloring/Collecting " & Wrd
    clr = Wrds.Item(WrdNo).wrdClr
    Dim Rng As Range, SntncRng As Range, Hdr As Range
    Dim h As String, hn As String, SecStr As String, si As Long
    Dim ListLevel As Long
    
    Set Rng = SDoc.Range
    Set SntncRng = SDoc.Range
    Set Hdr = SDoc.Range
    
    Rng.SetRange 1, 1
    Hdr.SetRange 1, 1
    Select Case MaxHdrLvl
    Case -2: h = "Docent"
    Case -1: h = "*"
    Case 1: h = "1"
    Case Else: h = "[1-" & MaxHdrLvl & "]"
    End Select
    Do
        PNo = Rng.Information(3)
        If ProgressBar.Spin(, "Collecting """ & Wrd & """ statements. (Page: " & PNo & ") (" & x & ") Statements found...") Then
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        If Set_Cancelled Then GoTo can
        With Rng.Find
            .text = Wrd
            .Wrap = wdFindStop
            .Execute
            If .Found And Rng.start > Set_SPos Then
                If Rng.start > Set_EPos Then Exit Do
                Rng.MoveStartUntil " .?!" & Chr(13) & Chr(10), wdBackward
                Rng.MoveEndUntil " .?!" & Chr(13) & Chr(10)
                If Rng.text = Wrd Then
                    x = x + 1
                    Wrds.Item(WrdNo).Count = Wrds.Item(WrdNo).Count + 1
                    If Set_Coloring Then Rng.HighlightColorIndex = clr
                    If mWillShall Then
                        'Move backward till you reach a header, or a prev. ".?!"
                        Set mSOW = FilteredSOWs.ItemByPosition(Rng.start)
        'If PNo = 22 And x = 202 Then Stop
                        SntncRng.SetRange mSOW.SectionRng.start, Rng.End
                        SecStr = SntncRng.text
'                        SntncRng.SetRange Rng.Start, Rng.End
                        si = InStrRev(SecStr, "?")
                        If si Then SntncRng.MoveStart , si: SecStr = SntncRng.text
                        si = InStrRev(SecStr, ".")
                        If si Then SntncRng.MoveStart , si: SecStr = SntncRng.text
                        si = InStrRev(SecStr, "!")
                        If si Then SntncRng.MoveStart , si: SecStr = SntncRng.text
'                        If InStr(SecStr, "?") + InStr(SecStr, ".") + InStr(SecStr, "!") > 0 Then
'                            SntncRng.MoveStartUntil ".?!", wdBackward
'                        Else
                        If mSOW.SectionRng.start = SntncRng.start Then
                            SntncRng.MoveStartUntil Chr(10) & Chr(13), wdBackward
                        End If
'                        If SntncRng.MoveStartUntil(".?!", wdBackward) = 0 Then SntncRng.MoveStartUntil Chr(10) & Chr(13), wdBackward
                        SntncRng.MoveStartWhile " "
                        Do While SntncRng.Characters(1).Font.Bold: SntncRng.MoveStart 1, 1: Loop
                        SntncRng.MoveEndUntil ".?!" '& Chr(13) & Chr(10)
                        SntncRng.MoveEndWhile ".?!"
                        ListLevel = SntncRng.Paragraphs(1).Range.ListFormat.ListLevelNumber
                        If mSOW.HeadingEnd > SntncRng.start Then SntncRng.SetRange mSOW.HeadingEnd + 1, SntncRng.End
                        Do While Set_Indenting And mSOW.ListLevel > ListLevel
                            Set mSOW = FilteredSOWs.ItemByPosition(mSOW.HeadingStart - 1)
                        Loop
                        i = i + 1
                        ReDim Preserve Arr(1 To 6, 1 To i)
                        Arr(1, i) = i
                        Arr(2, i) = mSOW.Number
                        Arr(3, i) = mSOW.Name
                        Arr(4, i) = SntncRng.text
                        Arr(5, i) = SntncRng.start
                        Arr(6, i) = Wrd
                        On Error Resume Next
                        SDoc.Bookmarks("Docent_" & i).Delete
                        On Error GoTo -1
                        SntncRng.Bookmarks.Add "Docent_" & i
                    End If
                End If
            Else
                Exit Do
            End If
            Rng.Collapse 0
        End With
    Loop
can:
End Sub
Private Sub ExportToExcel(Arr, Optional OFName As String)
    WriteLog 1, CurrentMod, "ExportToExcel", "Exporting to Excel"
    Dim i As Long, c As Long, r As Long, NArr, n As Long
    Dim WB As Object, Sh As Object, FName As String, x As Long
    On Error Resume Next
    ExcelApp.Workbooks(FName).Close False
    On Error GoTo ex
    FName = GetFileName(SDoc.Name, False)
    FName = IIf(Len(OFName) = 0, FName & " - Export.xlsx", FName & ".xlsx")
    Set WB = ExcelApp.Workbooks.Add
    Set Sh = WB.Sheets(1)
    r = UBound(Arr) + 1
    c = 4
    ProgressBar.Dom(2) = UBound(Arr)
    With Sh
        .Cells.NumberFormat = "@"
        .Cells.Font.Size = 12
        .Cells(1, 1).value = "Item No."
        .Cells(1, 2).value = "Section No."
        .Cells(1, 3).value = "Section Title"
        .Cells(1, 4).value = "Section Description" & Chr(10) & "(Click to view the full text)"
        .Cells(2, 1).Resize(UBound(Arr), 4).value = Arr
        ReDim NArr(1 To UBound(Arr), 1 To 1)
        For i = 1 To UBound(NArr)
            NArr(i, 1) = i
        Next
        .Cells(2, 1).Resize(UBound(NArr)).value = NArr
        .Columns(c).ColumnWidth = 105
        .Columns(3).ColumnWidth = 40
        .Columns(c).WrapText = True
        .Rows(1).RowHeight = 41.25
        .Columns(1).HorizontalAlignment = -4108
        .Cells.VerticalAlignment = -4108
        .UsedRange.Rows(1).AutoFilter
        If Len(OFName) > 0 Then
'            Err.Clear
'            On Error GoTo 0
            For i = 1 To UBound(Arr)
                .Hyperlinks.Add .Cells(i + 1, c), OFName, "Docent_" & Arr(i, 1) 'i - 1 'ActiveDocument.FullName
            Next
            .Columns(c).Font.Underline = False
            '"Reviewed"
            c = c + 1
            .Cells(1, c).value = "Responsible" & Chr(10) & "(Click in cell)"
            HighlightWordInCell .Cells(1, c), "(Click in cell)", -1, , True, True, False, 12
            AddValidationOptions .Range(.Cells(2, c), .Cells(r, c)), "John Smith,Jack Ryan"
            .Columns(c).ColumnWidth = 15
            '"Responsible"
            c = c + 1
            .Cells(1, c).value = "Reviewed?" & Chr(10) & "(Click in cell)"
            HighlightWordInCell .Cells(1, c), "(Click in cell)", -1, , True, True, False, 12
            AddValidationOptions .Range(.Cells(2, c), .Cells(r, c)), "Yes,No"
            .Columns(c).ColumnWidth = 15
            
            .Range(.Cells(2, 1), .Cells(r, c)).Locked = False
            
            If Set_Coloring Then
                For i = 1 To UBound(Arr)
                    If ProgressBar.Progress Then
                        MsgBox "Cancelled by user.", vbExclamation, ""
                        GoTo can
                    End If
                    If Set_Cancelled Then GoTo can
                    If i > 1 Then n = IIf(Arr(i, 6) = Arr(i - 1, 6) And Arr(i, 4) = Arr(i - 1, 4), n + 1, 1) Else n = 1
                    HighlightWordInCell .Cells(i + 1, 4), _
                            CStr(Arr(i, 6)), _
                            Wrds.Item(Arr(i, 6)).xlClr, _
                            n, True, False, True, 13
                    .Rows(i + 1).RowHeight = .Rows(i + 1).RowHeight + 30.75
                Next
            End If
        End If
        For i = 7 To 12
            With .Range(.Cells(1, 1), .Cells(r, c)).Borders(i)
               ' .Width = 1
                .LineStyle = 1
            End With
        Next
        With .Range(.Cells(1, 1), .Cells(1, c))
            .Interior.Color = 11184814
            .Font.Bold = True
            .Font.Size = 14
        End With
        .Protect DrawingObjects:=True, contents:=True, Scenarios:=True, AllowFiltering:=True
    End With
    WB.SaveAs OPath & FName, 51
    rExcel = UBound(Arr)
    Exit Sub
ex:
'Stop
'Resume
    WriteLog 3, CurrentMod, "ExportToExcel", Err.Number & ":" & Err.Description
    Exit Sub
can:
    ExcelApp.Quit
End Sub
Private Sub toExcel(Arr, OFName As String)
    WriteLog 1, CurrentMod, "toExcel", "Exporting to Excel"
    Dim NArr
    Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.DisplayAlerts = False
    ExcelApp.EnableEvents = False
    NArr = Sort2DArray(Transpose(Arr), 5)
    If Set_Export Then ExportToExcel NArr
    ExportToExcel NArr, OFName
    On Error Resume Next
'    ExcelApp.Cursor = -4143
    ExcelApp.DisplayAlerts = True
    ExcelApp.EnableEvents = True
    Exit Sub
ex:
    WriteLog 3, CurrentMod, "toExcel", Err.Number & ":" & Err.Description
    Exit Sub
can:
    ExcelApp.Quit
End Sub
Private Sub AddValidationOptions(Cell As Object, OptionsList As String)
    With Cell.validation
        .Delete
        .Add 3, 1, 1, OptionsList
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Private Sub HighlightWordsInCell(Cell As Object)
    Dim i As Long
    For i = 1 To Wrds.Count
        HighlightWordInCell Cell, Wrds.Item(i).Wrd, Wrds.Item(i).xlClr
    Next
End Sub
Private Sub HighlightWordInCell(Cell As Object, Wrd As String, clr As Long, _
        Optional ByVal n As Long = 1, Optional IsBold As Boolean = True, Optional IsItalic As Boolean = False, _
        Optional IsUnderline As Boolean = True, Optional ToSize As Long)
    Dim x As Long, s As String
    s = Cell.value
    x = 0
'    Do
    'If n > 1 Then Stop
        For n = n To 1 Step -1
            x = InStr(x + 1, s, Wrd)
            If x = 0 Then Exit For
        Next
'        If x = 0 Then Exit Do
        If ContainsWord(s, Wrd, x) Then
            With Cell.Characters(x, Len(Wrd)).Font
                If clr > -1 Then .Color = clr
                .Bold = IsBold
                .Italic = IsItalic
                .Underline = IsUnderline
                If ToSize Then .Size = ToSize
            End With
        End If
'    Loop
End Sub

