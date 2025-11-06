VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAttendees 
   Caption         =   "Action Item Selection"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8190
   OleObjectBlob   =   "frmAttendees.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAttendees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ComboBoxesArray() As New ComboBoxEvents
Private mDict As Dictionary
Private mGroupID As String
Private lDict As Dictionary
Private Headers As New Dictionary
Private OldItems As Dictionary
Private Busy As Boolean
Public EEvents As Boolean
Private Const PlaceHolder = "Select an option"
Private Const NoFilters = 1
Private Const HSpacer = 0.5
Private Const VSpacer = 2
'Private Sub UserForm_Initialize(): EEvents = True: End Sub
'Private Sub UserForm_Initialize(): EEvents = True: End Sub
Property Let GroupID(GID As String)
    mGroupID = GID
    DictectOlds
    Set ItemsDict = GetMembersOf(mGroupID)
End Property
Property Set ItemsDict(Dict As Dictionary)
    EEvents = False
    Set mDict = Dict
    Set lDict = Dict
    TableColumns
    PopList
    ResetCBoxes
    EEvents = True
End Property
Private Sub ResetFilters()
    EEvents = False
    Set lDict = mDict
    PopList
    ResetCBoxes
    EEvents = True
End Sub
Private Sub TableColumns()
    Dim i As Long, LastLeft As Single
    Dim CBox As ComboBox, n As Long
    Dim RemainingWidth As Single
    Set Headers = New Dictionary
    Headers.Add , Array(" ", 50, 0, 0) 'Text,Width,Alignmnet,FilterWidth
'    Headers.Add , Array("No.", 30, 2, 0)
    Headers.Add , Array("Name", 220, 0, 170)
'    Headers.Add , Array("Owner", 145, 0, 110)
'    Headers.Add , Array("Priority", 45, 2, 50)
'    Headers.Add , Array("Due Date", 60, 2, 51)
'    Headers.Add , Array("Closed?", 50, 2, 50)
    
    With ListView1
        LastLeft = .Left - HSpacer
        .View = lvwReport
        .Checkboxes = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .Gridlines = True
        RemainingWidth = .Width
'    'to search the whole text use this:
'    Set itm = ListView1.FindItem("some text goes here", lvwSubItem, 2)
'    'to search the partial text use this:
'    Set itm = ListView1.FindItem("some text goes here", lvwSubItem, 2, lvwPartial)
        With .ColumnHeaders
            .Clear
            For i = 1 To Headers.Count
                With .Add()
                    .text = Headers(i)(0)
                    If i = Headers.Count Then
                        .Width = RemainingWidth
                    Else
                        RemainingWidth = RemainingWidth - Headers(i)(1)
                        .Width = Headers(i)(1)
                    End If
                    If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
                End With
'                If (i > NoFilters Or Not ListView1.Checkboxes) And Headers(i)(3) > 0 Then
'                    Set CBox = Controls.Add("Forms.ComboBox.1", "cb" & i - 1 - NoFilters)
'                    n = n + 1
'                    ReDim Preserve ComboBoxesArray(1 To n)
'                    Set ComboBoxesArray(n).Control = CBox
'                    Set ComboBoxesArray(n).Parent = Me
'                    With CBox
'                        .Left = LastLeft + (2 * HSpacer)
'                        .Width = ListView1.ColumnHeaders(i).Width ' Headers(i)(3)
'                        .Font.Size = 11
'                        .Top = ListView1.Top - .Height - (2 * VSpacer)
'                    End With
'                End If
                LastLeft = LastLeft + Headers(i)(1) + HSpacer
            Next
        End With
    End With
End Sub
Sub ReFilter()
    Dim i As Long, j As Long, CBox As ComboBox ', lDict As New Dictionary
    If Not EEvents Then Exit Sub
    EEvents = False
    Set lDict = New Dictionary
    On Error Resume Next
    For i = 1 To mDict.Count
        j = -1
        For j = LBound(mDict(i)) To UBound(mDict(i))
            Set CBox = Me.Controls("cb" & j)
            If Not CBox Is Nothing Then
                If mDict(i)(j) <> CBox.value And CBox.value <> PlaceHolder And mDict(i)(j) Like CBox.value & "*" Then Exit For
            End If
        Next
        If j > UBound(mDict(i)) Then lDict.Add , mDict(i)
    Next
    PopList
    EEvents = True
End Sub
Private Sub RebuildComboboxes()
    Dim i As Long, j As Long, CBox As ComboBox, tVal As String, Coll As Collection
    On Error Resume Next
    For i = LBound(lDict(1)) To UBound(lDict(1)) '- 1
        Set CBox = Me.Controls("cb" & i)
        If Not CBox Is Nothing Then
            tVal = CBox.value
            CBox.Clear
            CBox.AddItem PlaceHolder
            Set Coll = New Collection
            For j = 1 To lDict.Count
                Coll.Add lDict(j)(i), CStr(lDict(j)(i))
            Next
            For j = 1 To Coll.Count
                If Not IsNull(lDict(j)) Then CBox.AddItem Coll(j)
            Next
            CBox.value = tVal 'IIf(Len(tVal) = 0, PlaceHolder, tVal)
        End If
    Next
End Sub
Private Sub ResetCBoxes()
    Dim i As Long, CBox As ComboBox
    On Error Resume Next
    For i = LBound(lDict(1)) To UBound(lDict(1)) - 1
        Set CBox = Me.Controls("cb" & i)
        If Not CBox Is Nothing Then
            CBox.value = PlaceHolder
        End If
    Next
End Sub

Private Sub PopList()
    Dim i As Long, j As Long
    With ListView1
        .ListItems.Clear
        For i = 1 To lDict.Count
            If Not IsNull(lDict(i)) Then
                With .ListItems.Add()
                    Select Case TypeName(lDict(i))
                    Case "Array", "Array()", "Variant()"
                        For j = LBound(lDict(i)) To UBound(lDict(i)) ' - 1
                            .ListSubItems.Add , , lDict(i)(j)
                        Next
                        .Checked = OldItems.Exists("_" & lDict(i)(UBound(lDict(i))))
                    Case "Dictection", "Dictionary"
                        For j = 2 To lDict(i).Count
                            .ListSubItems.Add , , lDict(i)(j)
                            .Checked = OldItems.Exists("_" & lDict(i)(j))
                        Next
                        '.Checked = OldItems.Exists("_" & lDict(i)(lDict(i)(1)))
                    Case "String"
                        .ListSubItems.Add , , lDict(i)
                        .Checked = OldItems.Exists("_" & lDict(i))
                    Case "Null"
                    Case Else
                        Err.Raise 5
                    End Select
                End With
            End If
        Next
        If OldItems.Count = mDict.Count And mDict.Count > 0 Then ckSelectAll.value = True
        .SortOrder = lvwAscending
        .Sorted = True
        .SortKey = 1
    End With
    RebuildComboboxes
End Sub

Private Sub btnCancel_Click(): Unload Me: End Sub

Private Sub btnOk_Click()
    Dim s As String, Item As ListItem
    For Each Item In ListView1.ListItems
        If Item.Checked Then s = s & Chr(13) & Item.SubItems(1)
    Next
    If Len(s) Then s = Right$(s, Len(s) - 1)
    SetContentControl mGroupID & " Attendees", s
    Me.Hide
End Sub

Private Sub btnReset_Click(): ResetFilters: End Sub

'Private Sub ckSelectAll_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
'
'End Sub

Private Sub ckSelectAll_Change()
    Dim Item As ListItem
    If Busy Then Exit Sub
    Busy = True
    If IsNull(ckSelectAll.value) Then ckSelectAll.value = True
    For Each Item In ListView1.ListItems
        Item.Checked = ckSelectAll.value
        ListView1_ItemCheck Item
    Next
    Busy = False
End Sub

'Private Sub ckSelectAll_Click()
'    If Busy Then Exit Sub
'    Busy = True
'    If IsNull(ckSelectAll.Value) Then ckSelectAll.Value = True
'    Busy = False
'End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Busy Then Exit Sub
    If Item.Checked Then
'        btnOk.Enabled = True
        If Not OldItems.Exists("_" & Item.SubItems(1)) Then OldItems.Add "_" & Item.SubItems(1), Item.SubItems(1)
    Else
        If OldItems.Exists("_" & Item.SubItems(1)) Then OldItems.Remove "_" & Item.SubItems(1)
        For Each Item In ListView1.ListItems
            If Item.Checked Then
'                btnOk.Enabled = True
                Exit For
            End If
        Next
    End If
    Busy = True
    
    Select Case mDict.Count
    Case 0: ckSelectAll.value = False
    Case OldItems.Count: ckSelectAll.value = True
    Case Else: ckSelectAll.value = Null
    End Select
    Busy = False
End Sub
Private Function GetURL(ID As String) As String
    ID = Left$(ID, Len(ID) - 2)
    Dim i As Long
    For i = 1 To lDict.Count
        If lDict(i)(0) = ID Then GetURL = lDict(i)(UBound(lDict(i))): Exit Function
    Next
End Function
Private Sub UserForm_Initialize()
    DictectOlds
    CenterUserform Me
End Sub
Private Sub DictectOlds()
    Dim CellTxt As String, ss() As String, i As Long
    CellTxt = GetContentControl(mGroupID & " Attendees")
    CellTxt = Replace(CellTxt, Chr(10), Chr(13))
    CellTxt = Replace(CellTxt, Chr(11), Chr(13))
    ss = Split(CellTxt, Chr(13))
    For i = LBound(ss) To UBound(ss)
        ss(i) = "_" & ss(i)
    Next
    Set OldItems = ArrToDict(ss)
End Sub
