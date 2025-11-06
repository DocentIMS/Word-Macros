VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SOWs"
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
Private Const CurrentMod = "SOWs"

Private mColl As Collection

Private mHeadingsColl As SOWs
Private mBoldsColl As SOWs
Private mHeadingsAndBoldsColl As SOWs
Private mBookmarksColl As SOWs
Private mALLColl As SOWs

Private mHeadingCount As Long
Private mBoldsCount As Long
Private mHeadingsAndBoldsCount As Long
Private mBookmarksCount As Long
Private mALLCount As Long

Private mIsCounted As Boolean

Private mIsSorted As Boolean
Private mIsColored As Boolean
Private Sorted1DArr() As Long
Private Sorted2DArr() As Long
Public Enum SearchingType
    Invalid = 0 'DEFAULT
    Headings = 2 ^ 0
    Bolds = 2 ^ 1
    HeadingsAndBolds = 2 ^ 2
    Bookmarks = 2 ^ 3
    all = 2 ^ 4
End Enum
'Private Enum CompareType
'    ecOR = 0 'default'
'    ecAnd = 1
'End Enum

Public SearchType As SearchingType

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
Property Get Item(NameOrNumber As Variant) As SOW
Attribute Item.VB_UserMemId = 0
'Dont Delete This Line
    Set Item = mColl(NameOrNumber)
End Property
Public Function ItemByPosition(NameOrNumber As Variant) As SOW
    On Error Resume Next
    Set ItemByPosition = mColl(GetmCollIndex(NameOrNumber))
End Function
Public Sub ReCount()
    mIsCounted = False
End Sub
Private Function ForceUpdateCounts(Optional FirstTime As Boolean = False)
    WriteLog 1, CurrentMod, "ForceUpdateCounts"
    Dim mSOW As SOW, i As Long
    If mIsCounted And Not FirstTime Then Exit Function
    mBookmarksCount = 0
    mHeadingCount = 0
    mBoldsCount = 0
    mHeadingsAndBoldsCount = 0
    For i = 1 To mColl.Count
        Set mSOW = mColl(i)
        With mSOW
            If .value >= Set_SPos And .value <= Set_EPos Then
                If .IsBookmark Then mBookmarksCount = mBookmarksCount + 1
                If .IsColored Or FirstTime Then
                    If .IsHeadingStyle Then
                        mHeadingCount = mHeadingCount + 1
'                        mHeadingsAndBoldsCount = mHeadingsAndBoldsCount + 1
                    ElseIf .IsBold Then
                        mBoldsCount = mBoldsCount + 1
'                        mHeadingsAndBoldsCount = mHeadingsAndBoldsCount + 1
                    End If
                End If
            End If
        End With
    Next
    mIsCounted = True
End Function
Public Property Get boldsCount() As Long
    WriteLog 1, CurrentMod, "BoldsCount"
    If mBoldsCount = 0 Then
        CollectBolds
'        mIsCounted = False
        ForceUpdateCounts True
    Else 'If Not mIsCounted Then
        ForceUpdateCounts
    End If
    boldsCount = mBoldsCount
End Property
'Public Property Let BoldsCount(x As Long): mBoldsCount = x: End Property
Public Property Get headingCount() As Long
    WriteLog 1, CurrentMod, "HeadingCount"
    If mHeadingCount = 0 Then
        CollectHeadings
'        mIsCounted = False
        ForceUpdateCounts True
    Else 'If Not mIsCounted Then
        ForceUpdateCounts
    End If
    headingCount = mHeadingCount
End Property
'Public Property Let HeadingCount(x As Long): mHeadingCount = x: End Property
Public Property Get bookmarksCount() As Long
    WriteLog 1, CurrentMod, "BookmarksCount"
    If mBookmarksCount = 0 Then
        CollectBookmarks
'        mIsCounted = False
        ForceUpdateCounts True
    Else 'If Not mIsCounted Then
        ForceUpdateCounts
    End If
    bookmarksCount = mBookmarksCount
End Property
Public Property Get AllCount() As Long
    WriteLog 1, CurrentMod, "AllCount"
    mALLCount = bookmarksCount + headingCount + boldsCount
    AllCount = mALLCount
End Property
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
'Private Sub AddBulk(ByVal Coll As Collection, Mode As SearchingType)
'    Dim mSOW As SOW, i As Long
'    mIsSorted = True
'    Set mColl = Coll
'    Select Case Mode
'    Case SearchingType.HeadingsAndBolds: Set mHeadingsAndBoldsColl = mColl
'    Case SearchingType.Bookmarks: Set mBookmarksColl = mColl
'    Case SearchingType.Headings: Set mHeadingsColl = mColl
'    Case SearchingType.Bolds: Set mBoldsColl = mColl
'    End Select
'    ReDim Sorted1DArr(1 To mColl.Count)
'    For i = 1 To mColl.Count
'        Set mSOW = mColl(i)
'        Sorted1DArr(i) = mSOW.HeadingStart
'    Next
'End Sub
Public Function AddSOW(ByVal mSOW As SOW) As SOW
    Set AddSOW = New SOW
    AddSOW.Initialize Me
    Set AddSOW.HeadingRng = mSOW.HeadingRng
    On Error Resume Next
    mColl.Add AddSOW, "_" & mSOW.HeadingStart
    If Err.Number = 0 Then
        ReDim Preserve Sorted1DArr(1 To mColl.Count)
        Sorted1DArr(mColl.Count) = mSOW.HeadingStart
    End If
End Function
Public Function GetRanges(Mode As SearchingType, Optional ByVal MustBeColored As Boolean = True) As SOWs
    Dim mSOW As SOW, i As Long
    Select Case Mode
    Case SearchingType.HeadingsAndBolds
        WriteLog 1, CurrentMod, "GetRanges", "Headings And Bolds"
        If mHeadingsAndBoldsColl Is Nothing Then
            Set mHeadingsAndBoldsColl = New SOWs
        ElseIf Not mHeadingsAndBoldsColl.IsSorted Then
            Set mHeadingsAndBoldsColl = New SOWs
        Else
            Set GetRanges = mHeadingsAndBoldsColl
            Exit Function
        End If
    Case SearchingType.Bookmarks
        WriteLog 1, CurrentMod, "GetRanges", "Bookmarks"
        MustBeColored = False
        If mBookmarksColl Is Nothing Then
            Set mBookmarksColl = New SOWs
        ElseIf Not mBookmarksColl.IsSorted Then
            Set mBookmarksColl = New SOWs
        Else
            Set GetRanges = mBookmarksColl
            Exit Function
        End If
    Case SearchingType.Headings
        WriteLog 1, CurrentMod, "GetRanges", "Headings"
        If mHeadingsColl Is Nothing Then
            Set mHeadingsColl = New SOWs
        ElseIf Not mHeadingsColl.IsSorted Then
            Set mHeadingsColl = New SOWs
        Else
            Set GetRanges = mHeadingsColl
            Exit Function
        End If
    Case SearchingType.Bolds
        WriteLog 1, CurrentMod, "GetRanges", "Bolds"
        If mBoldsColl Is Nothing Then
            Set mBoldsColl = New SOWs
        ElseIf Not mBoldsColl.IsSorted Then
            Set mBoldsColl = New SOWs
        Else
            Set GetRanges = mBoldsColl
            Exit Function
        End If
    Case Else
        WriteLog 1, CurrentMod, "GetRanges", "ALL"
        If mALLColl Is Nothing Then
            Set mALLColl = New SOWs
        ElseIf Not mALLColl.IsSorted Then
            Set mALLColl = New SOWs
        Else
            Set GetRanges = mALLColl
            Exit Function
        End If
    End Select
    'Sort them
    On Error GoTo ex
    If Not mIsSorted Then
        ReDim Sorted2DArr(1 To mColl.Count, 1 To 2)
        For i = 1 To mColl.Count
            Set mSOW = mColl(i)
            Sorted2DArr(i, 1) = mSOW.HeadingStart
            Sorted2DArr(i, 2) = i
        Next
        Sorted2DArr = Sort2DArray(Sorted2DArr, 1)
        mIsSorted = True
    End If
    '
    For i = LBound(Sorted2DArr) To UBound(Sorted2DArr)
        ProgressBar.Spin
        Set mSOW = mColl(Sorted2DArr(i, 2))
        With mSOW
            If (.IsColored Or Not MustBeColored) And .value >= Set_SPos And .value <= Set_EPos Then
                Select Case Mode
                Case SearchingType.HeadingsAndBolds
                    If .IsBold Or .IsHeadingStyle Then mHeadingsAndBoldsColl.AddSOW mSOW
                Case SearchingType.Headings
                    If .IsHeadingStyle Then mHeadingsColl.AddSOW mSOW
                Case SearchingType.Bookmarks
                    If .IsBookmark Then mBookmarksColl.AddSOW mSOW
                Case SearchingType.Bolds
                    If .IsBold Then mBoldsColl.AddSOW mSOW
                Case Else
                    mALLColl.AddSOW mSOW
                End Select
            End If
        End With
    Next
    Select Case Mode
    Case SearchingType.HeadingsAndBolds: Set GetRanges = mHeadingsAndBoldsColl
    Case SearchingType.Bolds: Set GetRanges = mBoldsColl
    Case SearchingType.Headings: Set GetRanges = mHeadingsColl
    Case SearchingType.Bookmarks: Set GetRanges = mBookmarksColl
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
    Dim mSOW As SOW, i As Long
    If Not mColl Is Nothing Then
        For i = 1 To mColl.Count
            Set mSOW = mColl(i)
            CountWords = CountWords + mSOW.CountWords
        Next
    End If
End Function
Public Sub DefineEnds()
    Dim i As Long, mSOW As SOW
    For i = 1 To mColl.Count
        ProgressBar.Spin
        Set mSOW = mColl(i)
        If i = mColl.Count Then
            mSOW.SectionEnd = Set_EPos
        Else
            mSOW.SectionEnd = mColl(i + 1).SectionStart - 1
        End If
        'mSOW.SectionRng.Select
    Next
End Sub
Public Property Get IsSorted() As Boolean: IsSorted = mIsSorted: End Property
Private Sub FindDeliverables()
    Dim nColl As Collection, i As Long, mSOW As SOW
    Select Case SearchType
    Case SearchingType.HeadingsAndBolds: Set nColl = mHeadingsAndBoldsColl
    Case SearchingType.Headings: Set nColl = mHeadingsColl
    Case SearchingType.Bolds: Set nColl = mBoldsColl
    Case SearchingType.Bookmarks: Set nColl = mBookmarksColl
    End Select
    For i = 1 To mColl.Count
        ProgressBar.Spin
        Set mSOW = mColl(i)
        mSOW.FindDeliverables
    Next
End Sub
Public Function Add(ByVal Rng As Range, StyleName As String) As SOW
    Dim i As Long
            'If InStr(Rng, "Cost Estimating Requirements") Then Stop
'    If Len(Rng) < 2 And StyleName <> "Bookmark" Then Stop: Exit Function
    mIsSorted = False
    i = GetmCollIndex(Rng.start)
    If i Then
        Set Add = mColl(i)
        If Rng.start <= Add.HeadingEnd + 1 Then
            If Add.ExtendRange(Rng) Then
                If StyleName = "Heading Docent" Then
                    If Not Add.StyleName Like "Heading *" Then Add.StyleName = StyleName
                Else
                    Add.StyleName = StyleName
                End If
            Else
                Set Add = Nothing
            End If
        Else
            Set Add = Nothing
        End If
    End If
    If Add Is Nothing Then
        On Error Resume Next
'        On Error GoTo skip
        Set Add = New SOW
        Add.Initialize Me
        Set Add.HeadingRng = Rng
        mColl.Add Add, "_" & Rng.start
        ReDim Preserve Sorted1DArr(1 To mColl.Count)
        Sorted1DArr(mColl.Count) = Rng.start
        Add.StyleName = StyleName
'skip:
'        On Error GoTo -1
    End If
End Function
Public Sub CollectHeadings()
    WriteLog 1, CurrentMod, "CollectHeadings"
    Dim i As Long, Rng As Range
    Set Rng = Selection.Range
    Application.ScreenUpdating = False
    For i = 1 To HLevel
        CollectHeading "Heading " & i
    Next
    CollectHeading "Heading Docent"
    AddBeforeFirst "Heading Docent"
    Rng.Select
    Application.ScreenUpdating = True
End Sub
Public Sub CollectBolds()
    Dim Rng As Range
    Set Rng = Selection.Range
    Application.ScreenUpdating = False
    WriteLog 1, CurrentMod, "CollectBolds"
    CollectHeading "Bold Text"
    AddBeforeFirst "Bold Text"
    Rng.Select
    Application.ScreenUpdating = True
End Sub
Public Sub CollectBookmarks()
    Dim Rng As Range
    Set Rng = Selection.Range
    Application.ScreenUpdating = False
    WriteLog 1, CurrentMod, "CollectBookmarks"
    Dim i As Long, c As Long, BkMrk As Bookmark
    Dim TRng As Range, s As String
    For i = 1 To SDoc.Bookmarks.Count
        Set BkMrk = SDoc.Bookmarks(i)
        Set TRng = BkMrk.Range.Paragraphs(1).Range
        s = CleanName(BkMrk.Name)
        If ProgressBar.Spin(, "Collecting Bookmarks...") Then
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        If BkMrk.Range.start < Set_EPos Then
            If Not BkMrk.Name Like "Docent_*" Then
                If BkMrk.Range.start = BkMrk.Range.End Then
                    BkMrk.Delete
                    Set BkMrk = SDoc.Bookmarks.Add(s, TRng)
                End If
                Add(BkMrk.Range, "Bookmark").Name = Trim(Replace(BkMrk.Name, "_", " "))
            End If
        End If
    Next
    AddBeforeFirst "Bookmark"
    Rng.Select
    Application.ScreenUpdating = True
can:
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
Public Sub ColorAll()
    WriteLog 1, CurrentMod, "ColorAll"
    Dim i As Long, mSOW As SOW
    If mIsColored Then Exit Sub
    For i = 1 To mColl.Count
        ProgressBar.Spin , "Coloring Headings..."
        Set mSOW = mColl(i)
    'If i = 121 Then Stop
        mSOW.ColorHeader
    Next
    mIsColored = True
End Sub
Private Sub CollectHeading(StyleName As String) ', n As Long, nMax As Long)
'    WriteLog 1, CurrentMod, "CollectHeading"
    Dim Rng As Range, RngStart As Long, RngEnd As Long
    If Len(StyleName) = 0 Then Exit Sub
    Set Rng = SDoc.Range
    Rng.SetRange Set_SPos, Set_EPos
    On Error Resume Next
    With Rng.Find
        .ClearAllFuzzyOptions
        .ClearFormatting
        .ClearHitHighlight
        .Replacement.ClearFormatting
        Select Case StyleName
        Case "Bold Text": .Font.Bold = True
        Case Else: .Style = StyleName
        End Select
        .text = ""
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        Do While .Execute
            If ProgressBar.Spin(, "Collecting " & StyleName & " (Page: " & Rng.Information(3) & ")") Then
                MsgBox "Cancelled by user.", vbExclamation, ""
                GoTo can
            End If
            If Rng.start > Set_EPos Then Exit Do
            If RngStart = Rng.start And RngEnd = Rng.End Then
                Rng.SetRange RngEnd + 1, RngEnd + 1
                If RngStart = Rng.start And RngEnd >= Rng.End Then Exit Do
            End If
            RngStart = Rng.start
            RngEnd = Rng.End
            Rng.MoveStartWhile "., " & Chr(13)
            Do
                Select Case Rng.Characters.Last
                Case ",", " ", Chr(13): Rng.MoveEnd 1, -1
                Case Else: Exit Do
                End Select
            Loop
            If IsHeader(Rng) Then
                Me.Add Rng, StyleName
                Rng.Collapse 0
            Else
                Rng.SetRange RngEnd, RngEnd
            End If
        Loop
    End With
    On Error Resume Next
    If ProgressBar.BarsCount > 1 Then ProgressBar.Progress 1
    Exit Sub
can:
    WriteLog 3, CurrentMod, "CollectHeading", "The user cancelled"
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



