VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public mColl As Collection

Public Function NewEnum() As IUnknown 'enable For...Each loops
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mColl.[_NewEnum]
End Function
Property Get Item(NameOrNumber As Variant) As SearchWord
Attribute Item.VB_UserMemId = 0
    Set Item = mColl(NameOrNumber)
End Property

Public Property Get Count() As Long: On Error Resume Next: Count = mColl.Count: End Property
Private Sub Class_Initialize(): Set mColl = New Collection: End Sub
Private Sub Class_Terminate(): Set mColl = Nothing: End Sub

Public Function Add(Word As String, WordType As String) As SearchWord
    Set Add = New SearchWord
    With Add
        .Wrd = Word
        .WrdType = WordType
        Select Case WordType
        Case "Shall": .wrdClr = wdYellow: .xlClr = 255: .wrdClrName = "Red"
        Case "Will": .wrdClr = wdBrightGreen: .xlClr = 3506772: .wrdClrName = "Bright Green"
        Case "Must": .wrdClr = wdPink: .xlClr = 16711935: .wrdClrName = "Pink"
        Case "Others": .wrdClr = wdTurquoise: .xlClr = 7348600: .wrdClrName = "Turquoise"
        End Select
    End With
    mColl.Add Add, Word
End Function
