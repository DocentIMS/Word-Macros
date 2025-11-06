VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' Header
'Attribute VB_Name = "ClassOrModuleName"
'Attribute VB_GlobalNameSpace = False ' ignored
'Attribute VB_Creatable = False ' ignored
'Attribute VB_PredeclaredId = False ' a Value of True creates a default global instance
'Attribute VB_Exposed = True ' Controls how the class can be instanced.
'
''Module Scoped Variables
'Attribute variableName.VB_VarUserMemId = 0 ' Zero indicates that this is the default member of the class.
'Attribute variableName.VB_VarDescription = "some string" ' Adds the text to the Object Browser information for this variable.
'
''Procedures
'Attribute procName.VB_Description = "some string" ' Adds the text to the Object Browser information for the procedure.
'Attribute procName.VB_UserMemId = someInteger
'    '  0: Makes the function the default member of the class.
'    ' -4: Specifies that the function returns an Enumerator.
'       Attribute NewEnum.VB_MemberFlags = "40"

Option Explicit
Option Compare Text
Private Const CurrentMod = "PlanItems"

Private mColl As Collection

Private mHighlightsColl As PlanItems
Private mCommentsColl As PlanItems
Private mHighlightsAndCommentsColl As PlanItems
'Private mBookmarksColl As PlanItems

Private mHighlightsCount As Long
Private mCommentsCount As Long
Private mHighlightsAndCommentsCount As Long
'Private mBookmarksCount As Long

Private mIsCounted As Boolean

Private mIsSorted As Boolean
Private mIsColored As Boolean
Private Sorted1DArr() As Long
Private Sorted2DArr() As Long
Public Enum PlanItemSearchingType
    Invalid = 0 'DEFAULT
    Highlights = 2 ^ 0
    Comments = 2 ^ 1
    HighlightsAndComments = 2 ^ 2
    all = 2 ^ 3
End Enum
'Private Enum CompareType
'    ecOR = 0 'default'
'    ecAnd = 1
'End Enum

Public SearchType As PlanItemSearchingType

Private Function EnumCompare( _
    theEnum As Variant, _
    enumMember As Variant, _
    Optional ByVal cType As CompareType = CompareType.ecOR) As Boolean
    Dim c As Long
    c = theEnum And enumMember
    EnumCompare = IIf(cType = CompareType.ecOR, c <> 0, c = enumMember)
End Function

Public Property Get Count() As Long: On Error Resume Next: Count = mColl.Count: End Property
Private Sub Class_Initialize(): Set mColl = New Collection: ReDim mSortingArr(1 To 1): End Sub
Private Sub Class_Terminate(): Set mColl = Nothing: End Sub
'Public Sub Remove(NameOrNumber As String): mColl.Remove NameOrNumber: End Sub
Public Function NewEnum() As IUnknown
    Set NewEnum = mColl.[_NewEnum]
End Function
Property Get Item(NameOrNumber As Variant) As PlanItem
Attribute Item.VB_UserMemId = 0
'Dont Delete This Line
    Set Item = mColl(NameOrNumber)
End Property
Public Function ItemByPosition(NameOrNumber As Variant) As PlanItem
    On Error Resume Next
    Set ItemByPosition = mColl(GetmCollIndex(NameOrNumber))
End Function
Public Sub ReCount()
    mIsCounted = False
End Sub
Private Function ForceUpdateCounts(Optional FirstTime As Boolean = False)
    WriteLog 1, CurrentMod, "ForceUpdateCounts"
    Dim mPlanItem As PlanItem, i As Long
    If mIsCounted And Not FirstTime Then Exit Function
'    mBookmarksCount = 0
    mHighlightsCount = 0
    mCommentsCount = 0
    mHighlightsAndCommentsCount = 0
    For i = 1 To mColl.Count
        Set mPlanItem = mColl(i)
        With mPlanItem
            If .value >= Set_SPos And .value <= Set_EPos Then
'                If .IsBookmark Then mBookmarksCount = mBookmarksCount + 1
                If .IsColored Or FirstTime Then
                    If .IsHighlightStyle Then
                        mHighlightsCount = mHighlightsCount + 1
'                        mHighlightsAndCommentsCount = mHighlightsAndCommentsCount + 1
                    ElseIf .IsComment Then
                        mCommentsCount = mCommentsCount + 1
'                        mHighlightsAndCommentsCount = mHighlightsAndCommentsCount + 1
                    End If
                End If
            End If
        End With
    Next
    mIsCounted = True
End Function
Public Property Get UploadURL() As String
    UploadURL = ProjectURLStr & "/planning-documents"
End Property
Public Property Get boldsCount() As Long
    WriteLog 1, CurrentMod, "BoldsCount"
    If mCommentsCount = 0 Then
        CollectComments
'        mIsCounted = False
        ForceUpdateCounts True
    Else 'If Not mIsCounted Then
        ForceUpdateCounts
    End If
    boldsCount = mCommentsCount
End Property
'Public Property Let BoldsCount(x As Long): mCommentsCount = x: End Property
Public Property Get headingCount() As Long
    WriteLog 1, CurrentMod, "HeadingCount"
    If mHighlightsCount = 0 Then
        CollectHighlights
'        mIsCounted = False
        ForceUpdateCounts True
    Else 'If Not mIsCounted Then
        ForceUpdateCounts
    End If
    headingCount = mHighlightsCount
End Property
'Public Property Let HeadingCount(x As Long): mHighlightsCount = x: End Property
'Public Property Get BookmarksCount() As Long
'    WriteLog 1, CurrentMod, "BookmarksCount"
'    If mBookmarksCount = 0 Then
'        CollectBookmarks
''        mIsCounted = False
'        ForceUpdateCounts True
'    Else 'If Not mIsCounted Then
'        ForceUpdateCounts
'    End If
'    BookmarksCount = mBookmarksCount
'End Property
'Public Property Let BookmarksCount(x As Long): mBookmarksCount = x: End Property
Public Sub Remove(NameOrNumber)
    mIsSorted = False
    mColl(GetmCollIndex(NameOrNumber)).UnColorHeader
End Sub
Public Sub RemoveByPosition(Position As Long)
    mIsSorted = False
    ItemByPosition(Position).UnColorHeader
End Sub
Private Function GetmCollIndex(ByVal NameOrNumber, Optional Mode As Long = 1) As Variant
    Select Case TypeName(NameOrNumber)
    Case "Range": NameOrNumber = Match(NameOrNumber.start, Sorted1DArr, Mode)
    Case "Long", "Integer": NameOrNumber = Match(NameOrNumber, Sorted1DArr, Mode)
    Case "String": If Left$(NameOrNumber, Mode) <> "_" Then NameOrNumber = "_" & NameOrNumber
    End Select
    GetmCollIndex = NameOrNumber
End Function
'Private Sub AddBulk(ByVal Coll As Collection, Mode As PlanItemSearchingType)
'    Dim mPlanItem As PlanItem, i As Long
'    mIsSorted = True
'    Set mColl = Coll
'    Select Case Mode
'    Case PlanItemSearchingType.HighlightsAndComments: Set mHighlightsAndCommentsColl = mColl
'    Case PlanItemSearchingType.Bookmarks: Set mBookmarksColl = mColl
'    Case PlanItemSearchingType.Highlights: Set mHighlightsColl = mColl
'    Case PlanItemSearchingType.Comments: Set mCommentsColl = mColl
'    End Select
'    ReDim Sorted1DArr(1 To mColl.Count)
'    For i = 1 To mColl.Count
'        Set mPlanItem = mColl(i)
'        Sorted1DArr(i) = mPlanItem.HeadingStart
'    Next
'End Sub
Public Function AddPlanItem(ByVal mPlanItem As PlanItem) As PlanItem
'    Set AddPlanItem = mPlanItem
    Set AddPlanItem = New PlanItem
    AddPlanItem.Initialize Me
    Set AddPlanItem.HighlightRng = mPlanItem.HighlightRng
'    Set AddPlanItem.SectionRng = mPlanItem.SectionRng
    AddPlanItem.Comment = mPlanItem.Comment
    AddPlanItem.Commenter = mPlanItem.Commenter
    On Error Resume Next
    mColl.Add AddPlanItem, "_" & mPlanItem.HighlightStart
    If Err.Number = 0 Then
        ReDim Preserve Sorted1DArr(1 To mColl.Count)
        Sorted1DArr(mColl.Count) = mPlanItem.HighlightStart
    End If
End Function
Public Function GetRanges(Mode As PlanItemSearchingType, Optional MustBeColored As Boolean = True) As PlanItems
    Dim mPlanItem As PlanItem, i As Long
    FindSearchRange
    Select Case Mode
    Case PlanItemSearchingType.HighlightsAndComments
        WriteLog 1, CurrentMod, "GetRanges", "Highlights And Comments"
        If mHighlightsAndCommentsColl Is Nothing Then
            Set mHighlightsAndCommentsColl = New PlanItems
        ElseIf Not mHighlightsAndCommentsColl.IsSorted Then
            Set mHighlightsAndCommentsColl = New PlanItems
        Else
            Set GetRanges = mHighlightsAndCommentsColl
            Exit Function
        End If
'    Case PlanItemSearchingType.Bookmarks
'        WriteLog 1, CurrentMod, "GetRanges", "Bookmarks"
'        If mBookmarksColl Is Nothing Then
'            Set mBookmarksColl = New PlanItems
'        ElseIf Not mBookmarksColl.IsSorted Then
'            Set mBookmarksColl = New PlanItems
'        Else
'            Set GetRanges = mBookmarksColl
'            Exit Function
'        End If
    Case PlanItemSearchingType.Highlights
        WriteLog 1, CurrentMod, "GetRanges", "Highlights"
        If mHighlightsColl Is Nothing Then
            Set mHighlightsColl = New PlanItems
        ElseIf Not mHighlightsColl.IsSorted Then
            Set mHighlightsColl = New PlanItems
        Else
            Set GetRanges = mHighlightsColl
            Exit Function
        End If
    Case PlanItemSearchingType.Comments
        WriteLog 1, CurrentMod, "GetRanges", "Comments"
        If mCommentsColl Is Nothing Then
            Set mCommentsColl = New PlanItems
        ElseIf Not mCommentsColl.IsSorted Then
            Set mCommentsColl = New PlanItems
        Else
            Set GetRanges = mCommentsColl
            Exit Function
        End If
    End Select

    On Error GoTo ex
    If Not mIsSorted Then
        ReDim Sorted2DArr(1 To mColl.Count, 1 To 2)
        For i = 1 To mColl.Count
            Set mPlanItem = mColl(i)
            Sorted2DArr(i, 1) = mPlanItem.HighlightStart
            Sorted2DArr(i, 2) = i
        Next
        Sorted2DArr = Sort2DArray(Sorted2DArr, 1)
        mIsSorted = True
    End If
    For i = LBound(Sorted2DArr) To UBound(Sorted2DArr)
        'ProgressBar.Spin
        Set mPlanItem = mColl(Sorted2DArr(i, 2))
        With mPlanItem
            If (.IsColored Or Not MustBeColored) And .value >= Set_SPos And .value <= Set_EPos Then
                Select Case Mode
                Case PlanItemSearchingType.HighlightsAndComments
                    If .IsComment Or .IsHighlightStyle Then mHighlightsAndCommentsColl.AddPlanItem mPlanItem
                Case PlanItemSearchingType.Highlights
                    If .IsHighlightStyle Then mHighlightsColl.AddPlanItem mPlanItem
'                Case PlanItemSearchingType.Bookmarks
'                    If .IsBookmark Then mBookmarksColl.AddPlanItem mPlanItem
                Case PlanItemSearchingType.Comments
                    If .IsComment Then mCommentsColl.AddPlanItem mPlanItem
                End Select
            End If
        End With
    Next
    Select Case Mode
    Case PlanItemSearchingType.HighlightsAndComments: Set GetRanges = mHighlightsAndCommentsColl
    Case PlanItemSearchingType.Comments: Set GetRanges = mCommentsColl
    Case PlanItemSearchingType.Highlights: Set GetRanges = mHighlightsColl
'    Case PlanItemSearchingType.Bookmarks: Set GetRanges = mBookmarksColl
    End Select
    GetRanges.SearchType = Mode
    GetRanges.DefineEnds
    Exit Function
ex:
    WriteLog 3, CurrentMod, "GetRanges", Err.Number & ":" & Err.Description
'    Stop
'    Resume
End Function
Public Function CountWords() As Long
    WriteLog 1, CurrentMod, "CountWords"
    Dim mPlanItem As PlanItem, i As Long
    If Not mColl Is Nothing Then
        For i = 1 To mColl.Count
            Set mPlanItem = mColl(i)
            CountWords = CountWords + mPlanItem.CountWords
        Next
    End If
End Function
Public Sub DefineEnds()
    Dim i As Long, mPlanItem As PlanItem
    For i = 1 To mColl.Count
        'ProgressBar.Spin
        Set mPlanItem = mColl(i)
        If i = mColl.Count Then
            mPlanItem.SectionEnd = Set_EPos
        Else
            mPlanItem.SectionEnd = mColl(i + 1).SectionStart - 1
        End If
        'mPlanItem.SectionRng.Select
    Next
End Sub
Public Property Get IsSorted() As Boolean: IsSorted = mIsSorted: End Property
'Private Sub FindDeliverables()
'    Dim nColl As Collection, i As Long, mPlanItem As PlanItem
'    Select Case SearchType
'    Case PlanItemSearchingType.HighlightsAndComments: Set nColl = mHighlightsAndCommentsColl
'    Case PlanItemSearchingType.Highlights: Set nColl = mHighlightsColl
'    Case PlanItemSearchingType.Comments: Set nColl = mCommentsColl
''    Case PlanItemSearchingType.Bookmarks: Set nColl = mBookmarksColl
'    End Select
'    For i = 1 To mColl.Count
'        ProgressBar.Spin
'        Set mPlanItem = mColl(i)
'        mPlanItem.FindDeliverables
'    Next
'End Sub
Private Function AddComment(ByVal Comm As Comment) As PlanItem
    Set AddComment = AddRange(Comm.Scope, "Comment")
    With AddComment
        .Commenter = Comm.Contact
        .Comment = Comm.Range.text
        If Len(.FullName) = 0 Then Set .HighlightRng = Comm.Range
'        If Len(.FullName) = 0 Then .FullName = Comm.Range.Text
    End With
End Function
Private Function AddRange(ByVal Rng As Range, StyleName As String) As PlanItem
    Dim i As Long
    mIsSorted = False
    i = GetmCollIndex(Rng.start)
    If i Then
        Set AddRange = mColl(i)
        If Rng.start <= AddRange.HighlightEnd + 1 Then
            If AddRange.ExtendRange(Rng) Then
                AddRange.StyleName = StyleName
            Else
                Set AddRange = Nothing
            End If
        Else
            Set AddRange = Nothing
        End If
    End If
    If AddRange Is Nothing Then
        Set AddRange = New PlanItem
        AddRange.Initialize Me
        Set AddRange.HighlightRng = Rng
        mColl.Add AddRange, "_" & Rng.start
        ReDim Preserve Sorted1DArr(1 To mColl.Count)
        Sorted1DArr(mColl.Count) = Rng.start
        AddRange.StyleName = StyleName
    End If
End Function
Public Function Add(ByVal What As Variant, StyleName As String) As PlanItem
    Select Case StyleName
    Case "Highlight": AddRange What, StyleName
    Case "Comment": AddComment What
    End Select
End Function
Public Sub CollectComments()
    WriteLog 1, CurrentMod, "CollectComments"
    Dim oRng As Range, Rng As Range, i As Long
    Set oRng = Selection.Range
    Application.ScreenUpdating = False
    Set Rng = ActiveDocument.Range
    For i = 1 To Rng.Comments.Count
        Me.Add Rng.Comments(i), "Comment"
    Next
    AddBeforeFirst "Highlight"
    oRng.Select
    Application.ScreenUpdating = True
End Sub
Private Sub AddBeforeFirst(StyleName As String)
    On Error GoTo ex
    If Set_SPos = 0 Then Exit Sub
    Dim Rng As Range
    Set Rng = ActiveDocument.Range
    Rng.SetRange Set_SPos, Set_SPos
    Me.Add Rng.Paragraphs(1).Range, StyleName
ex:
End Sub
'Public Sub ColorAll()
'    WriteLog 1, CurrentMod, "ColorAll"
'    Dim i As Long, mPlanItem As PlanItem
'    If mIsColored Then Exit Sub
'    For i = 1 To mColl.Count
'        ProgressBar.Spin , "Coloring Highlights..."
'        Set mPlanItem = mColl(i)
'    'If i = 121 Then Stop
'        mPlanItem.ColorHeader
'    Next
'    mIsColored = True
'End Sub
Public Sub CollectHighlights()
    WriteLog 1, CurrentMod, "CollectHighlights"
    Dim oRng As Range, Rng As Range, RngStart As Long, RngEnd As Long
    Application.ScreenUpdating = False
    Set oRng = Selection.Range
    Set Rng = SDoc.Range
    Rng.SetRange Set_SPos, Set_EPos
    On Error Resume Next
    With Rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
        .text = ""
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        .Highlight = True
        Do While .Execute
            If Rng.start <= RngEnd Then
                Rng.Move 1, 1
            Else
                Me.Add Rng, "Highlight"
                RngEnd = Rng.End
            End If
        Loop
    End With
    AddBeforeFirst "Highlight"
    oRng.Select
    Application.ScreenUpdating = True
'    On Error Resume Next
'    If ProgressBar.BarsCount > 1 Then ProgressBar.Progress 1
    Exit Sub
can:
    WriteLog 3, CurrentMod, "CollectHighlights", "The user cancelled"
End Sub
Private Function IsHeader(ByVal Rng As Range) As Boolean
    Dim RngTxt As String
    If Rng.Information(12) Then Exit Function 'In Table
    If Rng.Characters.First.Style Is Nothing Then Exit Function
    RngTxt = Rng.text
    If ContainsWord(RngTxt, "table") Then Exit Function
    If ContainsWord(RngTxt, "picture") Then Exit Function
    If ContainsWord(RngTxt, "image") Then Exit Function
    If ContainsWord(RngTxt, "figure") Then Exit Function
    If Len(ClearFromStr(RngTxt, Chr(13) & Chr(12) & "., ")) < 3 Then Exit Function
    IsHeader = True
End Function
Private Function HighlightsAsJSON() As String
    Dim PNo As Long, p As PlanItem, x As Long, Arr()
    ReDim Preserve Arr(1 To 5, 0 To x)
    Arr(1, 0) = "Location Start"
    Arr(2, 0) = "Location End"
    Arr(3, 0) = "Text"
    Arr(4, 0) = "Comment"
    Arr(5, 0) = "Commenter"
    For PNo = 1 To mColl.Count
        Set p = mColl(PNo)
        x = x + 1
        ReDim Preserve Arr(1 To 5, 0 To x)
        Arr(1, x) = p.HighlightStart
        Arr(2, x) = p.HighlightEnd
        Arr(3, x) = p.HighlightRng.text
        Arr(4, x) = p.Comment
        Arr(5, x) = p.Commenter
    Next
    HighlightsAsJSON = ArrToJSON(Transpose(Arr), "Highlights", True)
End Function
Public Sub ExportPlanningDocument()
'    Dim SFName As String
'    SFName = CreateAPIFolder(GetFileName(ActiveDocument.Name, False), "documents/planning-documents")
    ActiveDocument.Saved = False
    ActiveDocument.Save
    CreateAPIContent "planning_document", "planning-documents", _
        Array("file", "table"), Array(ActiveDocument.FullName, HighlightsAsJSON)
'    CreateAPIPlanningDocumnet ActiveDocument.FullName, HighlightsAsJSON, DefaultDocumentsFolder 'SFName
    ExportPlanningBreakdowns "planning-documents" 'DefaultDocumentsFolder 'SFName
End Sub
Private Sub ExportPlanningBreakdowns(SFName As String)
    Dim p As PlanItem, PNo As Long, FName As String, Resp As WebResponse, i As Long
    Dim SecRng As Range, pRng As Range, NRng As Range
    ProgressBar.Reset
    ProgressBar.Progress , "Exporting File No. 1", 0
    ProgressBar.Show
    On Error GoTo ex
    ProgressBar.Dom = mColl.Count
    Set pRng = ActiveDocument.Range
    Set NRng = ActiveDocument.Range
    For PNo = 1 To mColl.Count
        If ProgressBar.Progress(, "Exporting File No. " & PNo) Then
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        If Set_Cancelled Then GoTo can
        Set p = mColl(PNo)
        Set SecRng = p.SectionRng
        If PNo = 1 Then
            pRng.SetRange 0, 0
        Else
            pRng.SetRange p.SectionStart, p.SectionStart
            pRng.MoveStart wdParagraph, -HighlightLines
'            PRng.MoveStartUntil Chr(10) & Chr(13), -HighlightLines
'            PRng.SetRange mColl(PNo - 1).SectionStart, mColl(PNo - 1).SectionEnd
        End If
        If PNo = mColl.Count Then
            NRng.SetRange 0, 0
        Else
            NRng.SetRange p.SectionEnd, p.SectionEnd
            NRng.MoveEnd wdParagraph, HighlightLines
            'NRng.MoveEndUntil Chr(10) & Chr(13), HighlightLines
'            NRng.SetRange mColl(PNo + 1).SectionStart, mColl(PNo + 1).SectionEnd
        End If
        FName = ExportRange(SecRng, pRng, NRng, p.FullName, Environ("Temp") & "\", PNo, , p.HighlightRng)
        Set Resp = CreateAPIPlanningItem(FName, p.Comment, p.Commenter, SFName)
        Kill FName
    Next
    ProgressBar.Finished
'    Unload ProgressBar
    Exit Sub
can:
ex:
End Sub

