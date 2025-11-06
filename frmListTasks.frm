VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListTasks 
   Caption         =   "Select Task(s)"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11685
   OleObjectBlob   =   "frmListTasks.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ComboBoxesArray() As New ComboBoxEvents
Private mDict As Dictionary
Private lDict As Dictionary
Private Headers As New Dictionary
Private OldItems As Dictionary
Private PArr As Variant ', PDict As New Collection
'Private Busy As Boolean
Public EEvents As Boolean
Private Const PlaceHolder = "Filter..."
Private Const AllPlaceHolder = "Show All"
Private Const NoFilters = 2
Private Const HSpacer = 0.5
Private Const VSpacer = 2
'Private Sub UserForm_Initialize(): EEvents = True: End Sub
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
    Set Headers = New Dictionary
    Headers.Add , Array(" ", 15, 0, 0)
    Headers.Add , Array("No.", 30, 2, 30)
    Headers.Add , Array("Name", 210, 0, 170)
    Headers.Add , Array("Owner", 135, 0, 110)
    Headers.Add , Array("Priority", 45, 2, 50)
    Headers.Add , Array("Due Date", 80, 2, 51)
    Headers.Add , Array("Closed?", 50, 2, 50)
    
    With ListView1
        LastLeft = .Left - HSpacer
        .View = lvwReport
        .LabelEdit = lvwManual
        .Checkboxes = True
        .FullRowSelect = True
        .Gridlines = True
'    'to search the whole text use this:
'    Set itm = ListView1.FindItem("some text goes here", lvwSubItem, 2)
'    'to search the partial text use this:
'    Set itm = ListView1.FindItem("some text goes here", lvwSubItem, 2, lvwPartial)
        With .ColumnHeaders
            .Clear
            For i = 1 To Headers.Count
                With .Add()
                    .text = Headers(i)(0)
                    .Width = Headers(i)(1)
                    If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
                End With
                If i > NoFilters Or Not ListView1.Checkboxes Then
                    Set CBox = Controls.Add("Forms.ComboBox.1", "cb" & i - 2) ' - NoFilters)
                    n = n + 1
                    ReDim Preserve ComboBoxesArray(1 To n)
                    Set ComboBoxesArray(n).control = CBox
                    Set ComboBoxesArray(n).Parent = Me
                    With CBox
                        .Left = LastLeft + (2 * HSpacer)
                        .Width = Headers(i)(1)
                        .Font.Size = 11
                        .Top = ListView1.Top - .Height - (2 * VSpacer)
                    End With
                End If
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
            Set CBox = Nothing
            Set CBox = Me.Controls("cb" & j)
            If Not CBox Is Nothing Then
                If CBox.value = "" And mDict(i)(j) <> "" Then Exit For
                If mDict(i)(j) <> CBox.value And _
                    CBox.value <> PlaceHolder And _
                    CBox.value <> AllPlaceHolder And _
                    InStr(mDict(i)(j), CBox.value) = 0 Then Exit For
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
    For i = LBound(lDict(1)) To UBound(lDict(1)) - 1
        Set CBox = Me.Controls("cb" & i)
        If Not CBox Is Nothing Then
            tVal = CBox.value
            CBox.Clear
            CBox.AddItem AllPlaceHolder 'PlaceHolder
            Set Coll = New Collection
            Select Case Headers(i + 2)(0)
            Case "Priority"
                For j = LBound(PArr, 2) To UBound(PArr, 2)
                    Coll.Add PArr(2, j), PArr(2, j)
                Next
            Case "Closed?"
                Coll.Add "True", "True"
                Coll.Add "False", "False"
            End Select
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
'Private Sub PopList()
'    Dim i As Long, j As Long
'    With ListView1
'        .ListItems.Clear
'        For i = 1 To lDict.Count
'            With .ListItems.Add()
'                For j = LBound(lDict(i)) To UBound(lDict(i)) - 1
'                    .ListSubItems.Add , , lDict(i)(j)
'                Next
'                .Checked = OldItems.Exists("_" & lDict(i)(LBound(lDict(i))))
'            End With
'        Next
'        .SortOrder = lvwAscending
'        .Sorted = True
'        .SortKey = 1
'    End With
'    RebuildComboboxes
'End Sub
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
                        If Not OldItems Is Nothing Then .Checked = OldItems.Exists("_" & lDict(i)(1))
                    Case "Collection", "Dictionary"
                        For j = 2 To lDict(i).Count
                            .ListSubItems.Add , , lDict(i)(j)
                            If Not OldItems Is Nothing Then .Checked = OldItems.Exists("_" & lDict(i)(1))
                        Next
                        '.Checked = OldItems.Exists("_" & lDict(i)(lDict(i)(1)))
                    Case "String"
                        .ListSubItems.Add , , lDict(i)
                        If Not OldItems Is Nothing Then .Checked = OldItems.Exists("_" & lDict(i))
                    Case "Null"
                    Case Else
                        Err.Raise 5
                    End Select
                End With
            End If
        Next
'        If OldItems.Count = mDict.Count And mDict.Count > 0 Then ckSelectAll.Value = True
        .SortOrder = lvwAscending
        .Sorted = True
        .SortKey = 1
    End With
    RebuildComboboxes
End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub

Private Sub btnOk_Click()
    Dim Rng As Range, c As Long, n As Long, CellRng As Range, Item As ListItem, s As String
    Set Rng = ActiveDocument.Range
'    Rng.Move wdStory
'    Rng.InsertBreak wdPageBreak
    'n = 1
    'Rng.Tables.Add Rng, 1, 6
    Unprotect
    s = GetProperty(pPlannedTasks)
    With Rng.Tables(Rng.Tables.Count) ' Rng.Tables(1)
        For n = .Rows.Count To 3 Step -1
            .Rows(n).Delete
        Next
        For c = 1 To Headers.Count - 1
            Set CellRng = .Rows(.Rows.Count).Cells(c).Range
            CellRng.text = ""
        Next
        n = 1
        For Each Item In ListView1.ListItems
            If Item.Checked Then
                If n Then .Rows.Add
                n = n + 1
                s = s & ";"
                For c = 1 To Headers.Count - 1
                    Set CellRng = .Rows(.Rows.Count).Cells(c).Range
                    CellRng.text = Item.SubItems(c)
                    s = s & "," & Item.SubItems(c)
                    If c = 2 Then CellRng.Hyperlinks.Add CellRng, GetURL(.Rows(.Rows.Count).Cells(1).Range.text)
                Next
            End If
        Next
    End With
    SetProperty pPlannedTasks, s
    Protect
    Unload Me
End Sub
Private Sub btnReset_Click(): ResetFilters: End Sub

Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    If Busy Then Exit Sub
    If Item.Checked Then
        btnOk.Enabled = True
        If Not OldItems.Exists("_" & Item.SubItems(1)) Then OldItems.Add "_" & Item.SubItems(1), Item.SubItems(1)
    Else
        If OldItems.Exists("_" & Item.SubItems(1)) Then OldItems.Remove "_" & Item.SubItems(1)
        For Each Item In ListView1.ListItems
            If Item.Checked Then
                btnOk.Enabled = True
                Exit For
            End If
        Next
    End If
'    Busy = True
    
'    Select Case mDict.Count
'    Case 0: ckSelectAll.Value = False
'    Case OldItems.Count: ckSelectAll.Value = True
'    Case Else: ckSelectAll.Value = Null
'    End Select
'    Busy = False
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
    PArr = GetTaskPriorities
    CenterUserform Me
'    If IsGoodResponse(PArr(1, 1)) Then
'        For r = LBound(PArr, 2) To UBound(PArr, 2)
''            cbPriority.AddItem PArr(1, r)
'            PDict.Add PArr(2, r), PArr(1, r)
'        Next
'    End If
End Sub
Private Sub DictectOlds()
    Dim CellTxt As String
    Dim Tbl As Table, r As Long
    Set Tbl = GetTableByTitle("Planned Tasks")
    If Tbl Is Nothing Then Exit Sub

    Set OldItems = New Dictionary
'    Dim Rng As Range ', CellTxt As String
'    Set Rng = ActiveDocument.Range
    With Tbl 'Last Table in the document (?)
        For r = 3 To .Rows.Count
            CellTxt = CellText(.Rows(r).Cells(2).Range.text)
            OldItems.Add "_" & CellTxt, CellTxt
        Next
    End With
        
'    CellTxt = GetContentControl(mGroupID & " Attendees")
'    CellTxt = Replace(CellTxt, Chr(10), Chr(13))
'    CellTxt = Replace(CellTxt, Chr(11), Chr(13))
'    Set OldItems = ArrToDict(Split(CellTxt, Chr(13)))
End Sub

