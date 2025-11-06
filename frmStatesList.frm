VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatesList 
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   OleObjectBlob   =   "frmStatesList.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmStatesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Private SelectedItem As Long
'Private Sub ListView1_BeforeLabelEdit(Cancel As Integer): Cancel = True: End Sub
'
'Private Sub ListView1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean): Cancel = True: End Sub
'
'Private Sub ListView1_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    ListView1.ListItems(SelectedItem).Selected = True
'End Sub
'
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem): ForceSelection: End Sub

Private Sub ListView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS): ForceSelection: End Sub
Private Sub ListView1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS): ForceSelection: End Sub

'Private Sub ListBox1_Change()
'    Dim i As Long
'    For i = 0 To ListBox1.ListCount - 1
'        ListBox1.Selected(i) = False
'    Next
'    ListBox1.Selected(iState) = True
'End Sub
Private Sub UserForm_Initialize()
    CenterUserform Me
    Dim i As Long, j As Long
    Dim States As Dictionary, Transitions As Dictionary
    Dim CurrentState As String, InitialState As String, DocType As String
    CurrentState = GetProperty(pDocState)
    DocType = GetProperty(pDocType)
    InitialState = GetInitalState(DocType)
    Me.Caption = DocType & " Workflow States & Transitions"
    Set States = GetStatesOfDoc(DocType)
    
    With ListView1
        .View = lvwReport
        .LabelWrap = True
        .Checkboxes = False
        .MultiSelect = False
        .FullRowSelect = True
        .LabelEdit = lvwManual
'        .HotTracking = False
        .Gridlines = True
        With .ColumnHeaders
            .Clear
            With .Add()
                .text = "State"
                .Width = 60 '(ListView1.Width / 2) - 1
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Transitions"
                .Width = 60 '(ListView1.Width / 2) - 1
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
            With .Add()
                .text = "Description"
                .Width = ListView1.Width - 120
                'If UBound(Headers(i)) >= 2 Then .Alignment = Headers(i)(2)
            End With
        End With
        .ListItems.Clear
        On Error GoTo ex
        For i = 1 To States.Count
            With .ListItems.Add
                .text = IIf(States(i) = InitialState, "* " & States(i), States(i))
                .ListSubItems.Add
                .ListSubItems.Add , , GetStateDescription(States(i), DocType) ', , GetStateDescription(States(i))
            End With
            If States(i) = CurrentState Then SelectedItem = .ListItems.Count
            Set Transitions = GetStateTransitions(States(i), DocType)
            For j = 1 To Transitions.Count
                With .ListItems.Add
                    .ListSubItems.Add , , GetTransitionName(Transitions(j), DocType)
                    .ListSubItems.Add , , GetTransitionDescription(Transitions(j), DocType) ', , GetTransitionDescription(Transitions(j))
                End With
            Next
            .ListItems.Add
        Next
    End With
ex:
    ForceSelection
End Sub
Private Sub ForceSelection()
    On Error GoTo ex
    If SelectedItem Then
        ListView1.ListItems(SelectedItem).Selected = True
    Else
        ListView1.SelectedItem.Selected = False
    End If
ex:
End Sub
