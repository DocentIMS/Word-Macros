VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Private mDelivSearched As Boolean

Private mSectionStart As Long
Private mSectionEnd As Long
Private mHighlightStart As Long
Private mHighlightEnd As Long
'Private mDelivStart As Long
'Private mDelivEnd As Long
'Private mListLevel As Long
'Private mIndentLevel As Single
'Private mHighlightColor As Long

Private mIsColored As Boolean
Private mIsComment As Boolean
Private mIsHighlightStyle As Boolean
'Private mIsBookmark As Boolean

Private mCommenter As String
Private mComment As String
Private mFullName As String
Private mNumber As String
Private mName As String
Private mStyleName As String

Private mParent As PlanItems

Private mLineShape As Shape
Private mSectionRng As Range
Private mHighlightRng As Range
Private mDelivRng As Range

Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
'Dont Delete This Line
    value = mHighlightRng.start
End Property
'Public Property Let Value(v As Single): mValue = v: End Property

Public Property Get Parent() As PlanItems: Set Parent = mParent: End Property
'Public Property Set Parent(Hs As PlanItems): Set mParent = Hs: End Property
Public Sub Initialize(Hs As PlanItems): Set mParent = Hs: End Sub

Public Property Get SectionRng() As Range: Set SectionRng = mSectionRng: End Property
Public Property Get SectionStart() As Long
    On Error Resume Next
    SectionStart = mSectionRng.start
    If SectionStart = 0 Then SectionStart = mSectionStart
End Property
Public Property Let SectionStart(x As Long)
    mSectionStart = x
    Set mSectionRng = mHighlightRng.Paragraphs(1).Range
'    mSectionRng.SetRange mSectionStart, mSectionEnd
End Property
Public Property Get SectionEnd() As Long
    On Error Resume Next
    SectionEnd = mSectionRng.End
End Property
Public Property Let SectionEnd(x As Long)
    mSectionEnd = x
    Set mSectionRng = mHighlightRng.Paragraphs(1).Range
'    mSectionRng.SetRange mSectionStart, mSectionEnd
End Property

Public Property Get HighlightRng() As Range: Set HighlightRng = mHighlightRng: End Property
Public Property Set HighlightRng(ByVal Rng As Range)
    'If Rng.Start = 1442 Then Stop
    mHighlightStart = Rng.start
    mHighlightEnd = Rng.End
    Set mHighlightRng = Rng.Duplicate 'Rng.Parent.Range
'    mHighlightRng.SetRange mHighlightStart, mHighlightEnd
    Set mSectionRng = Rng.Paragraphs(1).Range
    With mHighlightRng
        mName = CleanName(.text)
        mNumber = CleanName(.ListFormat.ListString)
    End With
    SetFullName
'    If mLineShape Is Nothing Then AddLineBefore mHighlightRng
End Property
Public Property Get HighlightStart() As Long
    On Error Resume Next
    HighlightStart = mHighlightRng.start
End Property
Public Property Let HighlightStart(x As Long)
    mHighlightStart = x
    mHighlightRng.SetRange mHighlightStart, mHighlightEnd
End Property
Public Property Get HighlightEnd() As Long
    On Error Resume Next
    HighlightEnd = mHighlightRng.End
End Property
Public Property Let HighlightEnd(x As Long)
    mHighlightEnd = x
    mHighlightRng.SetRange mHighlightStart, mHighlightEnd
    Set HighlightRng = mHighlightRng
End Property

'Public Property Get ListLevel() As Long: ListLevel = mListLevel: End Property
'Public Property Let ListLevel(x As Long): mListLevel = x: End Property
'
'Public Property Get IndentLevel() As Long: IndentLevel = mIndentLevel: End Property
'Public Property Let IndentLevel(x As Long): mIndentLevel = x: End Property

'Public Sub FindDeliverables()
'    If Not mDelivSearched Then
'        Dim Rng As Range
'        Set Rng = mHighlightRng.Parent.Range
'        Rng.SetRange mSectionStart, mSectionEnd
'        With Rng.Find
'            .ClearFormatting
'            .Text = "Deliverables"
'            .Font.Bold = True
'            .Wrap = wdFindStop
'            .Execute
'            If .Found Then
'                mDelivStart = Rng.Start
'                mDelivEnd = Rng.End
'                Set mDelivRng = Rng
'            End If
'        End With
'    End If
'End Sub
'Public Property Get DelivRange() As Range
'    Set DelivRange = mDelivRng
'End Property
'Public Property Get DelivStart() As Long
'    If mDelivRng Is Nothing Then FindDeliverables
'    On Error Resume Next
'    DelivStart = mDelivRng.Start
'End Property
'Public Property Let DelivStart(x As Long)
'    mDelivStart = x
'    If mDelivRng Is Nothing Then Set mDelivRng = mHighlightRng.Parent.Range
'    mDelivRng.SetRange mDelivStart, mDelivEnd
'End Property
'Public Property Get DelivEnd() As Long
'    If mDelivRng Is Nothing Then FindDeliverables
'    On Error Resume Next
'    DelivEnd = mDelivRng.End
'End Property
'Public Property Let DelivEnd(x As Long)
'    mDelivEnd = x
'    If mDelivRng Is Nothing Then Set mDelivRng = mHighlightRng.Parent.Range
'    mDelivRng.SetRange mDelivStart, mDelivEnd
'End Property

Public Property Get IsColored() As Boolean
    Dim ClrIndx
    ClrIndx = mHighlightRng.Font.ColorIndex
    mIsColored = Not ((ClrIndx = wdAuto) Or (ClrIndx = wdBlack))
    IsColored = mIsColored
End Property
'Public Property Let IsColored(x As Boolean): mIsColored = x: End Property

'Public Property Get IsBookmark() As Boolean: IsBookmark = mIsBookmark: End Property
''Public Property Let IsBookMark(x As Boolean): mIsBookMark = x: End Property

Public Property Get IsComment() As Boolean: IsComment = mIsComment: End Property
'Public Property Let IsComment(x As Boolean): mIsComment = x: End Property

Public Property Get IsHighlightStyle() As Boolean: IsHighlightStyle = mIsHighlightStyle: End Property
'Public Property Let IsHighlightStyle(x As Boolean): mIsHighlightStyle = x: End Property

Public Property Get Name() As String:  Name = mName: End Property
Public Property Let Name(s As String)
    mName = s
    mNumber = ""
    SetFullName
End Property

Public Property Get FullName() As String:  FullName = mFullName: End Property
Public Property Let FullName(s As String): mFullName = s: End Property

Public Property Get Comment() As String:  Comment = mComment: End Property
Public Property Let Comment(s As String): mComment = s: End Property

Public Property Get Commenter() As String:  Commenter = mCommenter: End Property
Public Property Let Commenter(s As String): mCommenter = s: End Property

Public Property Get Number() As String:  Number = mNumber: End Property
Public Property Let Number(s As String)
    mNumber = s
    SetFullName
End Property

Public Property Get StyleName() As String:  StyleName = mStyleName: End Property
Public Property Let StyleName(s As String)
    mStyleName = s
    mIsHighlightStyle = mIsHighlightStyle Or (mStyleName = "Highlight")
    mIsComment = mIsComment Or (mStyleName = "Comment")
End Property
'Private Sub UpdateHighlightColor(HRng As Range)
'    'mStyleName = mHighlightRng.Style
''    If Not mIsHighlightStyle And mStyleName Like "Highlight *" Then Stop
'    'mIsHighlightStyle = mStyleName Like "Highlight *"
'    If mStyleName = "Highlight Docent" Then mHighlightColor = xlGreen
'    If mIsComment Then mHighlightColor = 16711935
'    If HRng.Style Is Nothing Then
'        mHighlightColor = xlGreen
'    Else
'        Select Case HRng.Style
'        Case "Highlight 1": mHighlightColor = xlBlue
'        Case "Highlight 2": mHighlightColor = xlRed
'        Case "Highlight 3": mHighlightColor = xlOrange
'        Case "Highlight 4": mHighlightColor = xlYellowGreen
'        Case "Highlight Docent": mHighlightColor = xlGreen
'        End Select
'    End If
'End Sub
'Public Sub ColorHeader()
'    Dim i As Long
'    If mHighlightRng.End > mHighlightRng.Paragraphs(1).Range.End Then
'        For i = 1 To mHighlightRng.Paragraphs.Count
'            UpdateHighlightColor mHighlightRng.Paragraphs(i).Range
'            mHighlightRng.Paragraphs(i).Range.Font.Color = mHighlightColor
'        Next
'    Else
'        UpdateHighlightColor mHighlightRng
'        mHighlightRng.Font.Color = mHighlightColor
'    End If
'    If mLineShape Is Nothing Then AddLineBefore
'    mLineShape.Line.ForeColor = mHighlightColor
'    If Not mIsColored Then
'        Parent.ReCount
'        mIsColored = True
'    End If
'End Sub
Public Sub UnColorHeader()
    mHighlightRng.Font.ColorIndex = wdBlack
    If Not mLineShape Is Nothing Then mLineShape.Delete
    Set mLineShape = Nothing
    'mLineShape.Line.ForeColor = wdBlack
    If mIsColored Then
        Parent.ReCount
        mIsColored = False
    End If
End Sub
Public Function ExtendRange(ByVal Rng As Range) As Boolean
    Dim NRng As Range
    Do
        Select Case Rng.Characters.Last
        Case ",", " ", Chr(13): Rng.MoveEnd 1, -1
        Case Else: Exit Do
        End Select
    Loop
    If mHighlightRng.start <= Rng.start And mHighlightRng.End >= Rng.End Or Rng.End - Rng.start = 0 Then
        ExtendRange = True
    Else
        Set NRng = mHighlightRng.Parent.Range
        NRng.SetRange mHighlightRng.End, mHighlightRng.End + 1
        'IIf(Rng.Start < mHighlightRng.Start, Rng.Start, mHighlightRng.Start), _
        IIf(Rng.End > mHighlightRng.End, Rng.End, mHighlightRng.End)
        ExtendRange = NRng.text <> Chr(13)
    End If
    If Not ExtendRange Then Exit Function
    If Rng.start < mHighlightRng.start Then
        mHighlightRng.start = Rng.start
        mSectionRng.start = Rng.start
    End If
    If Rng.End > mHighlightRng.End Then mHighlightRng.End = Rng.End
End Function
Private Sub SetFullName()
    Dim i As Long
    i = 0
    Do
        i = i + 1
        If i > Len(mName) Then i = 0: Exit Do
        Select Case Mid(mName, i, 1)
        Case 0 To 9, "-", ".", "(", ")", " "
        Case Else
            Exit Do
        End Select
    Loop
    i = i - 1
    If i > 0 Then
        mNumber = IIf(Len(mNumber) = 0, _
                Trim(Left$(mName, i)), _
                mNumber & " " & Trim(Left$(mName, i)))
        mName = Right$(mName, Len(mName) - i)
    End If
    mFullName = IIf(Len(mNumber) = 0, mName, mNumber & " " & mName)
End Sub
Private Function CleanName(ByVal Name As String) As String
    Dim i As Long, cr As String
    CleanName = Trim(Name)
    CleanName = Replace(CleanName, Chr$(9), " ")
    On Error Resume Next
    For i = Len(CleanName) To 1 Step -1
        cr = Mid(CleanName, i, 1)
        Select Case Asc(cr)
        Case 97 To 122, 65 To 90, 95, 46, 48 To 57, 45, 32
        Case Else
            CleanName = Replace(CleanName, cr, vbNullString)
        End Select
    Next
End Function
Public Function CountWords() As Long
'    Debug.Print mSectionRng.End
    CountWords = mSectionRng.ComputeStatistics(0)
End Function

'Private Sub AddLineBefore(Optional Rng As Range)
'    Dim y As Long, xi As Long, xe As Long
'    If Rng Is Nothing Then Set Rng = mHighlightRng
'    With Rng
'        .Select
'        y = .Information(wdVerticalPositionRelativeToPage) - (.Characters(1).Font.Size / 3)
'        xi = .PageSetup.LeftMargin
'        xe = .PageSetup.PageWidth
'    End With
'    xe = xe - xi
'    If xi > 22 Then xi = xi - 22: xe = xe + 22
'    Set mLineShape = ActiveDocument.Shapes.AddLine(xi, y, xe, y)
'    With mLineShape
'        .Line.ForeColor = Rng.Font.Color
'        .Line.DashStyle = msoLineDash
'        .Line.Weight = 1.5
'    End With
'End Sub
'Private Function GetLineBefore(Optional Rng As Range) As Shape
'    If Rng Is Nothing Then Set Rng = mHighlightRng
'    Set GetLineBefore = Rng.Paragraphs(1).Range.ShapeRange(1)
'End Function
