VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public UserForm As MSForms.UserForm
'Public Controls As New Collection

'Private mParent As MSForms.UserForm
Private mOkBtns As New Collection
Private mRequiredFields As New Collection
'-=-=-=-=-=-=-=-=-=-=-
Private mColl As Collection
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'Attribute NewEnum.VB_UserMemId = -4
'Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mColl.[_NewEnum]
End Function
Property Get Item(NameOrNumber As Variant) As CtrlEvent
Attribute Item.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0
    Set Item = mColl(NameOrNumber)
End Property
'-=-=-=-=-=-=-=-=-=-=-
Public Property Get Count() As Long: On Error Resume Next: Count = mColl.Count: End Property
Public Sub Remove(NameOrNumber As String): mColl.Remove NameOrNumber: End Sub
Private Sub Class_Initialize(): Set mColl = New Collection: End Sub
Private Sub Class_Terminate(): Set mColl = Nothing: End Sub
Public Property Get Controls() As Collection: Set Controls = mColl: End Property
Public Function Add(Ctrl As control) As CtrlEvent
    Set Add = New CtrlEvent
    With Add
        Set .Parent = Me
        Set .control = Ctrl
        .BaseName = GetBaseName(Ctrl.Name)
    End With
    mColl.Add Add, Ctrl.Name
End Function
Sub AddOkButton(OkBtn As MSForms.CommandButton): mOkBtns.Add OkBtn: End Sub
Property Set Parent(Parent As Object) 'MSForms.UserForm)
    CenterUserform Parent
    Set UserForm = Parent
    CollectAllControls
End Property
Private Sub CollectAllControls() 'Register All Conrtols as EventClass
    Dim Ctrl As control
    Set mColl = New Collection
    For Each Ctrl In UserForm.Controls
        Add Ctrl
        If TypeName(Ctrl) = "MultiPage" Then
            Dim xCtrl As control
            For Each xCtrl In Ctrl.Pages
                Add xCtrl
            Next
        End If
    Next
End Sub
Private Sub AddReq(ByVal CtrlName As String, Optional ByVal GroupName As String)
    Dim xColl As Collection
    If Len(GroupName) = 0 Then GroupName = CtrlName
'    If Len(GroupName) Then
        On Error Resume Next
        Set xColl = mRequiredFields(GroupName)
        If xColl Is Nothing Then Set xColl = New Collection
        xColl.Add mColl(CtrlName), CtrlName
        
        mRequiredFields.Remove GroupName
        mRequiredFields.Add xColl, GroupName
'    Else
        
'    End If
End Sub
Sub MakeAllRequired(Optional Typs As String = "tb,cb,lb,op", Optional ValidationType As String = "NotEmpty", Optional InvalidBackColor As Long = vbRed)
    Dim i As Long, ss() As String, j As Long ', Ctrl As MSForms.control
    ss = Split(Typs, ",")
    For i = 0 To UBound(ss)
        For j = 1 To mColl.Count
            If mColl(j).control.Name Like Trim(ss(i)) & "*" Then
                MakeEvRequired mColl(j).control.Name, ValidationType, InvalidBackColor
            End If
        Next
    Next
End Sub
Sub MakeRequired(Ctrls, Optional ValidationType As String = "NotEmpty", Optional InvalidBackColor As Long = vbRed)
    Dim i As Long, Ev As CtrlEvent
    Select Case TypeName(Ctrls)
    Case "String"
        Dim ss() As String
        ss = Split(Ctrls, ",")
        For i = 0 To UBound(ss)
            MakeEvRequired Trim(ss(i)), ValidationType, InvalidBackColor
        Next
    Case "Collection"
        For i = 1 To Ctrls.Count
            MakeRequired Ctrls(i)
        Next
    Case Else
        MakeEvRequired Ctrls.Name, ValidationType, InvalidBackColor
'        AddReq Ctrls.Name
''        mRequiredFields.Add mColl(Ctrls.Name), Ctrls.Name
'        AddMark GetBaseName(Ctrls.Name)
    End Select
End Sub
Private Sub MakeEvRequired(CtrlName As String, Optional ValidationType As String = "NotEmpty", Optional InvalidBackColor As Long = vbRed)
    Dim xCtrl As control, xCtrls As Collection, ReqGroupName As String
    Dim i As Long, j As Long, OkEnabled As Boolean ', IsValid As Long
    AddMark GetBaseName(CtrlName)
    Set xCtrls = GetRelated(CtrlName)
'    ReqMode = 1
'    If CtrlName Like "*Members" Then Stop
    For Each xCtrl In xCtrls
        If TypeName(xCtrl) = "Frame" Then
            ReqGroupName = GetBaseName(xCtrl.Name)
'            Set mColl(xCtrl.Name).Related = xCtrls
'            Set xCtrls = New Collection
'            xCtrls.Add xCtrl
'            mRequiredFields.Add mColl(xCtrl.Name), xCtrl.Name
'            For i = 1 To mOkBtns.Count
'                If mOkBtns(i).Enabled Then mOkBtns(i).Enabled = mColl(xCtrl.Name).IsValid(, -1)
'            Next
            Exit For
        End If
    Next
    For Each xCtrl In xCtrls
        With mColl(xCtrl.Name)
            .MakeRequired ValidationType, InvalidBackColor, ReqGroupName
            Set .Related = xCtrls
        End With
        AddReq xCtrl.Name, ReqGroupName
'        mRequiredFields.Add mColl(xCtrl.Name), xCtrl.Name
        For i = 1 To mOkBtns.Count
            If mOkBtns(i).Enabled Then If mColl(xCtrl.Name).IsValid(, -1) = 0 Then mOkBtns(i).Enabled = False
        Next
    Next
End Sub
Private Sub AddMark(CtrlName As String)
    Dim Ctrl As control, xCtrl As control, FrmMode As Boolean
    Dim Wdth As Single
'    Dim Prtn As Object
'    With
    On Error Resume Next
    Set Ctrl = UserForm.Controls("lb" & CtrlName)
    If Ctrl Is Nothing Then
        Set Ctrl = UserForm.Controls("fr" & CtrlName)
        If Ctrl Is Nothing Then Exit Sub
'        Ctrl.BorderColor = vbRed
'        Ctrl.BorderStyle = 1
        Ctrl.Caption = Ctrl.Caption & " *"
'        FrmMode = True
'        Set Prtn = Ctrl
    Else
    '        Set Prtn = Ctrl.Parent
    '    Prtn.ZOrder 1
    '        Set Prtn = Ctrl.Parent
        Set xCtrl = Ctrl.Parent.Controls.Add("Forms.Label.1", "req" & CtrlName, True)
        Wdth = Ctrl.Width
        Ctrl.AutoSize = True
        With xCtrl
            .ZOrder 1
            .Font.Size = 10
            .Caption = "*"
            .ForeColor = vbRed
            .AutoSize = True
    '        If FrmMode Then
    '            .BorderColor = vbRed
    '
    '            .Top = -6
    '            .Left = 8
    '        Else
            .Top = Ctrl.Top ' + 10
            .Left = Ctrl.Left + Ctrl.Width + 1
'            Ctrl.Left = Ctrl.Left + .Width
    '        End If
            .ZOrder 0
        End With
        Ctrl.AutoSize = False
        Ctrl.Width = Wdth
        Ctrl.ZOrder 1
        Add xCtrl
    End If
'    End With
End Sub
Sub MarkInvalid(ByVal CtrlName As String)
    Dim i As Long, xCtrl As control
    CtrlName = GetBaseName(CtrlName)
'    On Error Resume Next
    For Each xCtrl In GetRelated(CtrlName)
        With mColl(xCtrl.Name)
            .MarkAsValid False
        End With
    Next
    For i = 1 To mOkBtns.Count
        mOkBtns(i).Enabled = False
    Next
End Sub
Sub GrayOut(ByVal CtrlName As String, Optional Activate As Boolean, Optional Enabled As Boolean)
    Dim i As Long, xCtrl As control
    CtrlName = GetBaseName(CtrlName)
'    On Error Resume Next
    For Each xCtrl In GetRelated(CtrlName)
        With mColl(xCtrl.Name)
            .GrayOut Activate, Enabled
        End With
    Next
'    For i = 1 To mOkBtns.Count
'        mOkBtns(i).Enabled = False
'    Next
End Sub
Function IsValid(CtrlName As String) As Boolean
    Dim i As Long, xCtrl As control
    CtrlName = GetBaseName(CtrlName)
    IsValid = True
'    On Error Resume Next
    For Each xCtrl In GetRelated(CtrlName)
        IsValid = IsValid And (mColl(xCtrl.Name).IsValid <> 0)
    Next
End Function
Sub CheckOk()
    Dim i As Long, j As Long, x As Long
    Dim OkEnabled As Boolean, IsValid As Boolean, IsValidResult As Long
    OkEnabled = True
    For i = 1 To mRequiredFields.Count
        IsValid = False
        x = mRequiredFields(i).Count
        For j = 1 To x
'            If x > 1 Then Stop
            With mRequiredFields(i)(j)
                
'                Select Case .ReqMode
'                Case 2
'                Case 1
                    IsValidResult = .IsValid
                    IsValid = IsValid Or (IsValidResult = 1) Or (IsValidResult = -1 And x = 1)
'                End Select
            End With
        Next
'        If IsValid = False Then Stop
'        If x > 1 Then Stop
        For j = 1 To x
            With mRequiredFields(i)(j)
                .MarkAsValid IsValid
            End With
        Next
        OkEnabled = OkEnabled And IsValid
    Next
    For i = 1 To mOkBtns.Count
        mOkBtns(i).Enabled = OkEnabled
    Next
End Sub
Function GetRelated(CtrlName As String) As Collection
    Dim EvsColl As New Collection
    Dim xCtrl As control, BaseName As String, Prnt As Object
    On Error Resume Next
    Set xCtrl = UserForm.Controls(CtrlName)
    If TypeName(xCtrl) = "OptionButton" Then
        BaseName = xCtrl.GroupName
        For Each xCtrl In UserForm.Controls
            Select Case TypeName(xCtrl)
            Case "OptionButton"
                If xCtrl.GroupName = BaseName Then EvsColl.Add xCtrl, xCtrl.Name
            End Select
        Next
    Else
        BaseName = GetBaseName(CtrlName)
        For Each xCtrl In UserForm.Controls
            If xCtrl.Name Like "*" & BaseName Then EvsColl.Add xCtrl, xCtrl.Name
        Next
    End If
    Set xCtrl = UserForm.Controls(CtrlName)
    Set Prnt = xCtrl.Parent
    Do Until Prnt Is UserForm
        EvsColl.Add Prnt, Prnt.Name
        Set Prnt = Prnt.Parent
        If Prnt Is Nothing Then Exit Do
    Loop
    Set GetRelated = EvsColl
End Function
Private Function GetBaseName(CtrlName As String) As String
    Dim i As Long
    For i = 1 To Len(CtrlName)
        Select Case Asc(Mid(CtrlName, i, 1))
        Case 97 To 122
        Case Else: Exit For
        End Select
    Next
    GetBaseName = Mid(CtrlName, i)
End Function
Function GetDateNTime(FldName As String) As Variant
    Dim xCtrl As control, i As Long
    Dim DateArr(1 To 4) As String
    For i = 1 To 4
        DateArr(i) = "Null"
    Next
    For Each xCtrl In GetRelated(GetBaseName(FldName))
        Select Case Mid$(xCtrl.Name, 3, 1)
        Case "d": DateArr(1) = xCtrl.value
        Case "h": DateArr(2) = xCtrl.value
        Case "m": DateArr(3) = xCtrl.value
        Case "t": If Not xCtrl.Name Like "sp*" Then DateArr(4) = xCtrl.value
        End Select
    Next
    For i = 1 To 4
        If DateArr(i) = "" Then Exit Function
        If DateArr(i) = "Null" Then DateArr(i) = ""
    Next
    If Len(DateArr(1)) Then GetDateNTime = DateValue(DateArr(1))
    If Len(DateArr(2)) > 0 And Len(DateArr(3)) > 0 Then
        If DateArr(4) = "PM" Then
            If DateArr(2) <> 12 Then DateArr(2) = DateArr(2) + 12
        ElseIf DateArr(2) = 12 Then
            DateArr(2) = 0
        End If
        GetDateNTime = GetDateNTime + TimeSerial(DateArr(2), DateArr(3), 0)
    End If
End Function



