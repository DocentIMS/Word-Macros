Attribute VB_Name = "Docs_Scope1"
'Option Explicit
'Private Const TasksBookmark = "EOT"
'Private Tbl As Table
'Sub AddMeetingToScopeTask()
'    If CorrectTableSelected Is Nothing Then GoTo InvalidDoc
'    frmAddScopeMeeting.Display
'    Exit Sub
'InvalidDoc:
'    frmMsgBox.Display "Please place cursor inside the ""Related Meetings"" table of the task", , Critical, "Invalid Selection"
'End Sub
'Function CorrectTableSelected() As Object
'    Dim Doc As Document
'    Set Doc = ActiveDocument
'    If GetProperty(pDocType, Doc) <> "Scope" Then GoTo InvalidDoc
'    Dim Tbl As Table
'    On Error Resume Next
'    Set Tbl = Selection.Range.Tables(1)
'    If Tbl Is Nothing Then GoTo InvalidDoc
'    If Not Tbl.Title Like "*Meetings" Then GoTo InvalidDoc
'    Set CorrectTableSelected = Tbl
'InvalidDoc:
'End Function
'Function GetMeetingSummaryTable() As String
'    Dim s As String, Rng As Range, r As Long, c As Long, Cs As String, ss() As String, i As Long
'    Set Rng = ActiveDocument.Range
'    With Rng.Find
'        .text = "Meet* Summary Table"
'        .MatchWildcards = True
'        .Execute
'        If .Found Then
'            Do
'                Rng.MoveUntil Chr(7)
'            Loop Until Rng.Information(wdWithInTable)
'            Set Tbl = Rng.Tables(1)
'        End If
'    End With
'    If Tbl Is Nothing Then Exit Function
'    Set Tbl = ActiveDocument.Tables(ActiveDocument.Tables.Count)
'    s = """Meeting Summary Table"": ["
'    For r = 2 To Tbl.Rows.Count
'        s = s & "{"
'        For c = 1 To Tbl.Rows(r).Cells.Count
'            Cs = Cell(r, c)
'            s = s & """" & Cell(1, c) & """: "
'            If InStr(Cs, Chr(13)) Then
'                ss = Split(Cs, Chr(13))
'                s = s & "["
'                For i = 0 To UBound(ss)
'                    s = s & """" & Trim(ss(i)) & ""","
'                Next
'                s = Left$(s, Len(s) - 1) & "]"
'            Else
'                s = s & """" & Cs & """"
'            End If
'            s = s & ","
'        Next
'        s = Left$(s, Len(s) - 1) & "},"
'    Next
'    GetMeetingSummaryTable = Left$(s, Len(s) - 1) & "]"
'End Function
'Sub FillScopeFields()
'    SetContentControl "Contract Number", ContractNumberStr
'    SetContentControl "Customer Name", ProjectClientStr
'    SetContentControl "Project Name", ProjectNameStr
'    FillFirstDate
''    Application.Application.Run "ActiveDocument!FillFirstDate"
'End Sub
'Sub UnlockDocument() '(ByVal control As IRibbonControl)
'    On Error Resume Next
'    If ActiveDocument.protectionType <> wdNoProtection Then
'        ActiveDocument.Unprotect Application.UserName
'        If ActiveDocument.protectionType <> wdNoProtection Then
'            frmMsgBox.Display "You are not allowed to edit this final version", , Critical
'        Else
'            frmMsgBox.Display "Document is unlocked", , Success
'        End If
'    End If
'End Sub
'Sub FillFirstRevisionDate()
'    On Error Resume Next
'    Set Tbl = GetTableByTitle("Revisions Table") ' ActiveDocument.Tables(1)
'    If Tbl Is Nothing Then Exit Sub
'    If Len(Cell(2, 3)) = 0 Then Tbl.Rows(2).Cells(3).Range.text = Format(ToServerTime, DateFormat) '"mm/dd/yyyy")
'End Sub
'Sub UpdateRevision() '(ByVal control As IRibbonControl)
'    Dim i As Long
'    If ActiveDocument.protectionType <> wdNoProtection Then
'        frmMsgBox.Display "You must unlock the file first", , Critical
'        Exit Sub
'    End If
'    With frmUpdateRev
'        .Show
'        If Not .Cancelled Then
'            If .IsFinal Then
'                AddRev "Final Scope of Work (Clean)", .tbNotes.value, "Final"
'            Else
'                AddRev "Draft Scope of Work", .tbNotes.value, ""
'            End If
'            UpdateFooterVersion
'            If .IsFinal Then
'                For i = 1 To ActiveDocument.Comments.Count
'                    ActiveDocument.Comments(1).DeleteRecursively
'                Next
'                Protect ActiveDocument, Application.UserName
'                ActiveDocument.Protect wdAllowOnlyReading, False, Application.UserName, False, False
'                SetProperty pIsFinalRev, True, ActiveDocument
'                LoadDocInfo ActiveDocument
'                RefreshRibbon
'                frmMsgBox.Display "Comments removed and Password (not really) Set", , Success
'            End If
'        End If
'    End With
'    Unload frmUpdateRev
'End Sub
'Sub UpdateFooterVersion()
'    Dim VName As String
'    VName = getVName
'    SetContentControl "VersionName", VName '"mm/dd/yyyy"
'    SetContentControl "LastSaveTime", IIf(VName = "Final", Format(ToServerTime, DateFormat), Format(ToServerTime, DateTimeFormat)) '"mm/dd/yyyy - hh:mm AM/PM"))
'End Sub
'Private Function getVName() As String
'    Set Tbl = ActiveDocument.Tables(1)
'    getVName = Cell(Tbl.Rows.Count, 1)
'End Function
'Private Sub AddRev(Title As String, Notes As String, Optional VName As String)
'    Dim r As Long
'    Set Tbl = ActiveDocument.Tables(1)
'    With Tbl
'        r = .Rows.Count
'        Unprotect .Range.Document
'        .Rows.Add
'        If Len(VName) = 0 Then
'            For r = r To 2 Step -1
'                If Cell(r, 1) Like "Draft *" Or r = 2 Then
'                    VName = "Draft " & (Val(Replace(Cell(r, 1), "Draft ", "")) + 1)
'                    Exit For
'                End If
'            Next
'        End If
'        .Rows.Last.Cells(1).Range.text = VName
'        .Rows.Last.Cells(2).Range.text = Title
'        .Rows.Last.Cells(3).Range.text = Format(ToServerTime, DateFormat) ' "mm/dd/yyyy")
'        .Rows.Last.Cells(4).Range.text = Notes
'        Protect .Range.Document
'    End With
'End Sub
'Private Function GetDocentHeaderLvl(Rng As Range, LevelRange As String) As Long
'    Dim i As Long, RngStyle As String, ss() As String, j As Long
'    If Not Rng.Paragraphs(1).Range.text Like "Task #*" Then Exit Function
'    RngStyle = Rng.Paragraphs(1).Range.Style
'    ss = Split(RngStyle, ",")
'    For j = LBound(ss) To UBound(ss)
'        If ss(j) Like "Heading " & LevelRange Then
'            GetDocentHeaderLvl = Val(Replace(ss(j), "Heading ", ""))
'            Exit Function
'        End If
'    Next
'End Function
'Sub AddSubLevel()
'    On Error GoTo ex
'    Dim Rng As Range, i As Long, St As scopeTask
'    Set Rng = FindLastHeading(Selection.Range, "[2-4]").Paragraphs(1).Range
'    i = GetDocentHeaderLvl(Rng, "[2-4]")
'    Select Case i
'    Case 0
'        GoTo ex
'    Case 2 To 4
'        Set St = NextTaskNum(GetTaskNumber(Rng, True, i), i + 1, Rng)
'        Rng.Application.UndoRecord.StartCustomRecord "Add " & St.taskNum
'        AddTask St.Range, i + 1, St.taskNum
'        Rng.Document.TablesOfContents(1).Update
'    Case Else
'        frmMsgBox.Display "There is not Heading 6", , Information
'    End Select
'    Exit Sub
'ex:
'    frmMsgBox.Display "Place the cursor on the parent task first", , Information
'End Sub
'Sub AddSameLevel()
'    On Error GoTo ex
'    Dim Rng As Range, i As Long, St As scopeTask
'    Set Rng = FindLastHeading(Selection.Range, "[2-5]").Paragraphs(1).Range
'    i = GetDocentHeaderLvl(Rng, "[2-5]")
'    Select Case i
'    Case 0
'        GoTo ex
'    Case 2
'        AddTopLevel
'    Case Else
'        Set St = NextTaskNum(GetTaskNumber(Rng, False, i), i, Rng)
'        Rng.Application.UndoRecord.StartCustomRecord "Add " & St.taskNum
'        AddTask St.Range, i, St.taskNum
'        Rng.Document.TablesOfContents(1).Update
'    End Select
'    Exit Sub
'ex:
'    frmMsgBox.Display "Place the cursor on the sibling task first", , Information
'End Sub
'Sub AddTopLevel()
'    Dim Rng As Range, TaskName As String, taskNumStr As String, TaskNo As Long
'    TaskName = frmInputBox.Display("Please insert the task name", "Create New Task")
'    If Len(TaskName) = 0 Or TaskName = "Canceled" Then Exit Sub
'    Set Tbl = GetTableByTitle("Tasks Table")
'    If Tbl Is Nothing Then GoTo ex
'    Unprotect Tbl.Range.Document
'    With GetLastRow
'        TaskNo = .Index - 1
'        taskNumStr = "Task " & TaskNo & ":"
'        .Cells(1).Range.text = "Task " & .Index - 1 & ":"
'        .Cells(2).Range.text = TaskName
'    End With
'    Set Rng = ActiveDocument.Range
'        Set Rng = ActiveDocument.Range.GoTo(wdGoToBookmark, Name:=TasksBookmark)
'    If TaskNo = 1 Then
'    Else
'        Set Rng = ActiveDocument.Range.GoTo(wdGoToBookmark, Name:=TaskNo - 1)
'    End If
'    Rng.Application.UndoRecord.StartCustomRecord "Add " & taskNum & " " & TaskName
'    AddTask Rng, 2, , taskNumStr & " " & TaskName
'    Exit Sub
'ex:
'    frmMsgBox.Display "Tasks Table is missing", , Information
'End Sub
'Function FindLastHeading(Rng As Range, HLevels As String) As Range
'    Do Until GetDocentHeaderLvl(Rng, HLevels) > 0  'Rng.Style Like "Heading " & HLevels Or Rng.Style Like "Heading " & HLevels & "*"
'        Rng.Move wdParagraph, -1
'        Rng.Select
'        If Rng.start = 0 Then Exit Do
'    Loop
'    Set FindLastHeading = Rng
'End Function
'Sub RemoveScopeTasksPBreaks(Doc As Document)
'    Dim Rng As Range
'    Unprotect Doc
'    Set Rng = Doc.Range
'    With Rng.Find
'        .ClearAllFuzzyOptions
'        .ClearFormatting
'        .ClearHitHighlight
'        .Replacement.ClearFormatting
'        .text = "^m"
'        .Font.Color = vbRed
'        .Replacement.text = ""
'        .Wrap = wdFindContinue
'        .Forward = True
'        .Execute Replace:=wdReplaceAll
'        .ClearAllFuzzyOptions
'        .ClearFormatting
'        .ClearHitHighlight
'        .Replacement.ClearFormatting
'    End With
'    Protect Doc
'End Sub
'Sub AddEditor(Rng As Range)
'    Dim IsP As Long
'    IsP = Rng.Document.protectionType
'    If Not IsP = wdNoProtection Then Unprotect Rng.Document
'    Rng.Editors.Add wdEditorEveryone
'    If Not IsP = wdNoProtection Then Protect Rng.Document
'End Sub
'Sub AddTask(Rng As Range, Level As Long, Optional TaskNumber, Optional ByVal TaskStr As String)
'    Dim i As Long, Tbl As Table
'    If Len(TaskStr) = 0 Then TaskStr = TaskNumber & " " & frmInputBox.Display("Please insert the task name", "Docent IMS")
'    If Not IsMissing(TaskNumber) Then
'        If TaskStr = TaskNumber & " " Then Exit Sub
'        If TaskStr = TaskNumber & " Canceled" Then Exit Sub
'    End If
'    TaskNumber = Replace(Replace(Left$(TaskStr, InStr(TaskStr, ":") - 1), " ", ""), ".", "_")
'    i = Rng.start
'    With Rng
'        If i <> .Document.Bookmarks(TasksBookmark).Range.start Then
'            .MoveUntil Chr(12)
'        End If
'        If .Characters.Last = Chr(13) Then .Move 1, -1
'        i = .Font.Color
'        Unprotect .Document
'        On Error Resume Next
'        .InsertBreak 7
'        Do While Err.Number = 4605
'            Err.Clear
'            .Move 1, 1
'            AddEditor Rng
'            .InsertBreak 7
'        Loop
'        If Err.Number Then
'            On Error GoTo ex
'            .Collapse 0
'            .InsertBreak 7
'            .MoveStart 1, -1
'        Else
'            .MoveStart 1, -2
'        End If
'        .Font.Color = vbRed
'        .Collapse 0
'        .Font.Color = i
'        .Style = "Heading " & Level
'        .text = TaskStr & Chr(13)
'        On Error Resume Next
'        .Document.Bookmarks.Add TaskNumber, .Paragraphs(1).Range
'        If Err.Number Then Stop
'        If .Paragraphs.Count = 1 Then .InsertParagraphAfter
'        .Move 1, 1
'        .Style = "Normal"
'        .text = "Objectives" & Chr(10) & Chr(10) & Chr(10) & _
'                "Assumptions" & Chr(10) & Chr(10) & Chr(10) & _
'                "Deliverables" & Chr(10) & Chr(10) & Chr(10) & _
'                "Related Meetings" & Chr(10) & Chr(10) & Chr(10)
'        .MoveEnd 1
'        For i = 2 To 8 Step 3
'            .Paragraphs(i).Indent
'            .Paragraphs(i).Range.ContentControls.Add().Range.HighlightColorIndex = wdGray25
'            .Paragraphs(i).Range.ListFormat.ApplyListTemplateWithLevel ListGalleries(wdBulletGallery).ListTemplates(1)
'        Next
'        On Error Resume Next
'        For i = 1 To .Paragraphs.Count Step 3
'            .Paragraphs(i).Range.Font.Bold = True
'            AddEditor .Paragraphs(i + 1).Range
'        Next
'        AddTable TaskStr, "Assumptions", Rng, Array("Assumption", "Cost?")
'        AddTable TaskStr, "Deliverables", Rng, Array("Deliverable", "Date?")
'        AddTable TaskStr, "Related Meetings", Rng, Array("Meeting Type", "Frequency", "Number of Meetings", _
'                "Length (Hours)", "Prep Time (Hours)", "Consultant Attendees", "Which Meeting?")
'        .Document.TablesOfContents(1).Update
'        Protect .Document
'        .Document.Bookmarks(TaskNumber).Range.Select
'    End With
'
'    Exit Sub
'ex:
'    Stop
'    Resume
'End Sub
'Private Sub AddTable(TaskStr As String, TableName As String, Rng As Range, Cols)
'    Dim i As Long, Tbl As Table, PNum As Long
'    With Rng
'        For PNum = 1 To .Paragraphs.Count
'            If cellText(.Paragraphs(PNum).Range.text) = TableName Then
'                PNum = PNum + 1
'                Exit For
'            End If
'        Next
'        Set Tbl = .Tables.Add(.Paragraphs(PNum).Range, 2, UBound(Cols) - LBound(Cols) + 1)
'        Tbl.Title = TaskStr & TableName
'        For i = -6 To -1
'            With Tbl.Borders(i)
'                .LineStyle = Options.DefaultBorderLineStyle
'                .LineWidth = Options.DefaultBorderLineWidth
'                .Color = Options.DefaultBorderColor
'            End With
'        Next
'        Tbl.Rows(1).Shading.BackgroundPatternColor = -603930625 ' = wdGray25
'        Tbl.Rows(1).Range.Bold = True
'        AddEditor Tbl.Rows(2).Range
'        For i = LBound(Cols) To UBound(Cols)
'            Tbl.Cell(1, i - LBound(Cols) + 1).Range.text = Cols(i)
'        Next
'    End With
'End Sub
' Function GetTaskNumber(Rng As Range, IsSub As Boolean, Optional Level As Long) As String
'    Dim Fnd As Range
'    Set Fnd = Rng.Document.Range
'    Fnd.SetRange Rng.start, Rng.End
'    If Fnd.Style = "Heading " & Level Then
'        GetTaskNumber = ExtractTaskNum(Fnd.text, IsSub)
'    Else
'        Dim i As Long
'        For i = 1 To Rng.Paragraphs.Count
'            If GetDocentHeaderLvl(Rng.Paragraphs(i).Range, CStr(Level)) = Level Then
'            GetTaskNumber = ExtractTaskNum(Fnd.text, IsSub)
'            Exit Function
'            End If
'        Next
'        Exit Function
'        With Fnd.Find
'            .ClearAllFuzzyOptions
'            .ClearFormatting
'            .ClearHitHighlight
'            .Replacement.ClearFormatting
'            .Style = "Heading " & Level
'            .text = ""
'            .MatchWildcards = True
'            .Forward = False
'            If Not .Found Then .Style = "Heading " & Level & "*"
'            If .Found Then GetTaskNumber = ExtractTaskNum(Fnd.text, IsSub)
'        End With
'    End If
'End Function
'Function NextTaskNum(ByVal LastFound As String, Level As Long, Optional Fnd As Range) As scopeTask
'    Dim i As Long, St As New scopeTask
'    If Fnd Is Nothing Then Set Fnd = ActiveDocument.Range
'    With Fnd.Find
'        .ClearAllFuzzyOptions
'        .ClearFormatting
'        .ClearHitHighlight
'        .Replacement.ClearFormatting
'        .MatchWildcards = True
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Style = "Heading " & Level
'        Do
'            .text = LastFound
'            .Execute
'            If .Found Then LastFound = ExtractTaskNum(LastFound, False)
'        Loop While .Found
'    End With
'    Set St.Range = Fnd
'    St.taskNum = LastFound
'    Set NextTaskNum = St 'Array(Fnd, LastFound)
'End Function
'Function ExtractTaskNum(TaskName As String, IsSub As Boolean) As String
'    Dim s As String, ss() As String, i As Long, x As Long
'    s = Split(TaskName, ":")(0)
'    If IsSub Then
'        ExtractTaskNum = s & ".1:"
'    Else
'        ss = Split(s, " ")
'        ss = Split(ss(UBound(ss)), "-")
'        ss = Split(ss(UBound(ss)), ".")
'        i = ss(UBound(ss))
'        x = InStrRev(s, i)
'        s = Left$(s, x - 1) & i + 1
'        ExtractTaskNum = s & ":"
'    End If
'End Function
'Private Function Cell(r As Long, c As Long) As String
'    Cell = Trim(Replace(Tbl.Rows(r).Cells(c).Range.text, Chr(13) & Chr(7), ""))
'End Function
'Private Function GetLastRow() As Row
'    Set GetLastRow = Tbl.Rows.Last
'    If Len(Cell(Tbl.Rows.Count, 1)) = 0 Then Exit Function
'    Tbl.Rows.Add
'    Set GetLastRow = Tbl.Rows.Last
'End Function
'Private Function FindOldTask(TaskNo As String) As Range
'    Dim i As Long, Fnd As Range
'    Set Fnd = ActiveDocument.Range
'    With Fnd.Find
'        .ClearAllFuzzyOptions
'        .ClearFormatting
'        .ClearHitHighlight
'        .Replacement.ClearFormatting
'        .MatchWildcards = True
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Style = "Heading 2"
'        .text = TaskNo
'        .Execute
'        If .Found Then
'            Set FindOldTask = Fnd.Paragraphs(1).Range
'            FindOldTask.MoveEndWhile Chr(10) & Chr(13), -1
'        End If
'    End With
'End Function
'
'
