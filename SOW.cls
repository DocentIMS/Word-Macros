VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SOW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private mDelivSearched As Boolean

Private mSectionStart As Long
Private mSectionEnd As Long
Private mHeadingStart As Long
Private mHeadingEnd As Long
Private mDelivStart As Long
Private mDelivEnd As Long
Private mListLevel As Long
Private mIndentLevel As Single
Private mHeadingColor As Long

Private mIsColored As Boolean
Private mIsBold As Boolean
Private mIsHeadingStyle As Boolean
Private mIsBookmark As Boolean

Private mFullName As String
Private mNumber As String
Private mName As String
Private mStyleName As String

Private mParent As SOWs

Private mLineShape As Shape
Private mSectionRng As Range
Private mHeadingRng As Range
Private mDelivRng As Range

Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
'Dont Delete This Line
    value = mHeadingRng.start
End Property
'Public Property Let Value(v As Single): mValue = v: End Property

Public Property Get Parent() As SOWs: Set Parent = mParent: End Property
'Public Property Set Parent(Hs As SOWs): Set mParent = Hs: End Property
Public Sub Initialize(Hs As SOWs): Set mParent = Hs: End Sub

Public Property Get SectionRng() As Range: Set SectionRng = mSectionRng: End Property
Public Property Get SectionStart() As Long
    On Error Resume Next
    SectionStart = mSectionRng.start
    If SectionStart = 0 Then SectionStart = mSectionStart
End Property
Public Property Let SectionStart(x As Long)
    mSectionStart = x
    Set mSectionRng = mHeadingRng.Parent.Range
    mSectionRng.SetRange mSectionStart, mSectionEnd
End Property
Public Property Get SectionEnd() As Long
    On Error Resume Next
    SectionEnd = mSectionRng.End
End Property
Public Property Let SectionEnd(x As Long)
    mSectionEnd = x
    Set mSectionRng = mHeadingRng.Parent.Range
    mSectionRng.SetRange mSectionStart, mSectionEnd
End Property

Public Property Get HeadingRng() As Range: Set HeadingRng = mHeadingRng: End Property
Public Property Set HeadingRng(ByVal Rng As Range)
    'If Rng.Start = 1442 Then Stop
    mHeadingStart = Rng.start
    mHeadingEnd = Rng.End
    Set mHeadingRng = Rng.Parent.Range
    mHeadingRng.SetRange mHeadingStart, mHeadingEnd
    SectionStart = mHeadingRng.start
    SectionEnd = mHeadingRng.End
    With mHeadingRng
        mListLevel = .ListFormat.ListLevelNumber
        mIndentLevel = .Paragraphs(1).LeftIndent
        mName = CleanName(.text)
        mNumber = CleanName(.ListFormat.ListString)
    End With
    SetFullName
    If mLineShape Is Nothing Then AddLineBefore mHeadingRng
End Property
Public Property Get HeadingStart() As Long
    On Error Resume Next
    HeadingStart = mHeadingRng.start
End Property
Public Property Let HeadingStart(x As Long)
    mHeadingStart = x
    mHeadingRng.SetRange mHeadingStart, mHeadingEnd
End Property
Public Property Get HeadingEnd() As Long
    On Error Resume Next
    HeadingEnd = mHeadingRng.End
End Property
Public Property Let HeadingEnd(x As Long)
    mHeadingEnd = x
    mHeadingRng.SetRange mHeadingStart, mHeadingEnd
    Set HeadingRng = mHeadingRng
End Property

Public Property Get ListLevel() As Long: ListLevel = mListLevel: End Property
Public Property Let ListLevel(x As Long): mListLevel = x: End Property

Public Property Get IndentLevel() As Long: IndentLevel = mIndentLevel: End Property
Public Property Let IndentLevel(x As Long): mIndentLevel = x: End Property

Public Sub FindDeliverables()
    If Not mDelivSearched Then
        Dim Rng As Range
        Set Rng = mHeadingRng.Parent.Range
        Rng.SetRange mSectionStart, mSectionEnd
        With Rng.Find
            .ClearFormatting
            .text = "Deliverables"
            .Font.Bold = True
            .Wrap = wdFindStop
            .Execute
            If .Found Then
                mDelivStart = Rng.start
                mDelivEnd = Rng.End
                Set mDelivRng = Rng
            End If
        End With
    End If
End Sub
Public Property Get DelivRange() As Range
    Set DelivRange = mDelivRng
End Property
Public Property Get DelivStart() As Long
    If mDelivRng Is Nothing Then FindDeliverables
    On Error Resume Next
    DelivStart = mDelivRng.start
End Property
Public Property Let DelivStart(x As Long)
    mDelivStart = x
    If mDelivRng Is Nothing Then Set mDelivRng = mHeadingRng.Parent.Range
    mDelivRng.SetRange mDelivStart, mDelivEnd
End Property
Public Property Get DelivEnd() As Long
    If mDelivRng Is Nothing Then FindDeliverables
    On Error Resume Next
    DelivEnd = mDelivRng.End
End Property
Public Property Let DelivEnd(x As Long)
    mDelivEnd = x
    If mDelivRng Is Nothing Then Set mDelivRng = mHeadingRng.Parent.Range
    mDelivRng.SetRange mDelivStart, mDelivEnd
End Property

Public Property Get IsColored() As Boolean
    Dim ClrIndx
    ClrIndx = mHeadingRng.Font.ColorIndex
    mIsColored = Not ((ClrIndx = wdAuto) Or (ClrIndx = wdBlack))
    IsColored = mIsColored
End Property
'Public Property Let IsColored(x As Boolean): mIsColored = x: End Property

Public Property Get IsBookmark() As Boolean: IsBookmark = mIsBookmark: End Property
'Public Property Let IsBookMark(x As Boolean): mIsBookMark = x: End Property

Public Property Get IsBold() As Boolean: IsBold = mIsBold: End Property
'Public Property Let IsBold(x As Boolean): mIsBold = x: End Property

Public Property Get IsHeadingStyle() As Boolean: IsHeadingStyle = mIsHeadingStyle: End Property
'Public Property Let IsHeadingStyle(x As Boolean): mIsHeadingStyle = x: End Property

Public Property Get Name() As String:  Name = mName: End Property
Public Property Let Name(s As String)
    mName = s
    mNumber = ""
    SetFullName
End Property

Public Property Get FullName() As String:  FullName = mFullName: End Property
Public Property Let FullName(s As String): mFullName = s: End Property

Public Property Get Number() As String:  Number = mNumber: End Property
Public Property Let Number(s As String)
    mNumber = s
    SetFullName
End Property

Public Property Get StyleName() As String:  StyleName = mStyleName: End Property
Public Property Let StyleName(s As String)
    mStyleName = s
    mIsBookmark = mIsBookmark Or (mStyleName = "Bookmark")
    mIsHeadingStyle = mIsHeadingStyle Or (mStyleName Like "Heading *")
    mIsBold = mIsBold Or (mStyleName = "Bold Text")
End Property
Private Sub UpdateHeadingColor(HRng As Range)
    'mStyleName = mHeadingRng.Style
'    If Not mIsHeadingStyle And mStyleName Like "Heading *" Then Stop
    'mIsHeadingStyle = mStyleName Like "Heading *"
    If mStyleName = "Heading Docent" Then mHeadingColor = xlGreen
    If mIsBold Then mHeadingColor = 16711935
    If HRng.Style Is Nothing Then
        mHeadingColor = xlGreen
    Else
        Select Case HRng.Style
        Case "Heading 1": mHeadingColor = xlBlue
        Case "Heading 2": mHeadingColor = xlRed
        Case "Heading 3": mHeadingColor = xlOrange
        Case "Heading 4": mHeadingColor = xlYellowGreen
        Case "Heading Docent": mHeadingColor = xlGreen
        End Select
    End If
End Sub
Public Sub ColorHeader()
    Dim i As Long
    On Error GoTo ex
    If mHeadingRng.End > mHeadingRng.Paragraphs(1).Range.End Then
        For i = 1 To mHeadingRng.Paragraphs.Count
            UpdateHeadingColor mHeadingRng.Paragraphs(i).Range
            mHeadingRng.Paragraphs(i).Range.Font.Color = mHeadingColor
        Next
    Else
        UpdateHeadingColor mHeadingRng
        mHeadingRng.Font.Color = mHeadingColor
    End If
    If mLineShape Is Nothing Then AddLineBefore
    mLineShape.Line.ForeColor = mHeadingColor
    If Not mIsColored Then
        Parent.ReCount
        mIsColored = True
    End If
ex:
End Sub
Public Sub UnColorHeader()
    If mLineShape Is Nothing Then Set mLineShape = GetLineBefore
    mHeadingRng.Font.ColorIndex = wdBlack
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
    If mHeadingRng.start <= Rng.start And mHeadingRng.End >= Rng.End Or Rng.End - Rng.start = 0 Then
        ExtendRange = True
    Else
        Set NRng = mHeadingRng.Parent.Range
        NRng.SetRange mHeadingRng.End, mHeadingRng.End + 1
        'IIf(Rng.Start < mHeadingRng.Start, Rng.Start, mHeadingRng.Start), _
        IIf(Rng.End > mHeadingRng.End, Rng.End, mHeadingRng.End)
        ExtendRange = NRng.text <> Chr(13)
    End If
    If Not ExtendRange Then Exit Function
    If Rng.start < mHeadingRng.start Then
        mHeadingRng.start = Rng.start
        mSectionRng.start = Rng.start
    End If
    If Rng.End > mHeadingRng.End Then mHeadingRng.End = Rng.End
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
            '"[a-z][A-Z]_.[0-9]- "
        Case Else
            CleanName = Replace(CleanName, cr, vbNullString)
        End Select
    Next
End Function
Public Function CountWords() As Long
'    Debug.Print mSectionRng.End
    CountWords = mSectionRng.ComputeStatistics(0)
End Function

Private Sub AddLineBefore(Optional Rng As Range)
    Dim Y As Long, xi As Long, xe As Long
    Set mLineShape = GetLineBefore
    If Not mLineShape Is Nothing Then Exit Sub
    On Error Resume Next
    If Rng Is Nothing Then Set Rng = mHeadingRng
    With Rng
        .Select
        Y = .Information(wdVerticalPositionRelativeToPage) - (.Characters(1).Font.Size / 3)
        xi = .PageSetup.LeftMargin
        xe = .PageSetup.PageWidth
    End With
    xe = xe - xi
    If xi > 22 Then xi = xi - 22: xe = xe + 22
    Set mLineShape = ActiveDocument.Shapes.AddLine(xi, Y, xe, Y)
    With mLineShape
        .Line.ForeColor = Rng.Font.Color
        .Line.DashStyle = msoLineDash
        .Line.Weight = 1.5
    End With
End Sub
Private Function GetLineBefore(Optional Rng As Range) As Shape
    On Error Resume Next
    If Rng Is Nothing Then Set Rng = mHeadingRng
    Set GetLineBefore = Rng.Paragraphs(1).Range.ShapeRange(1)
End Function
