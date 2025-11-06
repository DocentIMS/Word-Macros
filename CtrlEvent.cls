VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CtrlEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const cbPlaceHolder = "-- Choose "

Private mParent As CtrlEvents
Private mRelatedCtrls As Collection 'CtrlEvents
Private mCtrl As MSForms.control
Private mCtrlName As String
Private mBaseName As String
Private mCtrlType As String
Private mDisableEvents As Boolean
Private mMPLastPage As Long
Private mMPThisPage As Long
'Private mValidationStr As String

Private mForecolor As Long
Private mValidationType As String
Private mInvalidBackColor As Long
Private mValidBackColor As Long
Private mValidForeColor As Long
Private mReqGroupName As String
'Private mIsRequired As Boolean
Private mManuallyEdited As Boolean
Private WithEvents mCheckBox As MSForms.CheckBox 'ck
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mAMPMCheckBox As MSForms.CheckBox 'ampmck
Attribute mAMPMCheckBox.VB_VarHelpID = -1

Private WithEvents mTextBox As MSForms.TextBox 'tb
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mDateTextbox As MSForms.TextBox 'tbd
Attribute mDateTextbox.VB_VarHelpID = -1
Private WithEvents mMinutesTextBox As MSForms.TextBox 'tbm
Attribute mMinutesTextBox.VB_VarHelpID = -1
Private WithEvents mHoursTextBox As MSForms.TextBox 'tbh
Attribute mHoursTextBox.VB_VarHelpID = -1
Private WithEvents mNumbersTextBox As MSForms.TextBox 'tbh
Attribute mNumbersTextBox.VB_VarHelpID = -1
Private WithEvents mDecimalsTextBox As MSForms.TextBox 'tbh
Attribute mDecimalsTextBox.VB_VarHelpID = -1

Private WithEvents mComboBox As MSForms.ComboBox 'cb
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mAMPMComboBox As MSForms.ComboBox 'cbtt
Attribute mAMPMComboBox.VB_VarHelpID = -1

Private WithEvents mSpinButton As MSForms.SpinButton 'sp
Attribute mSpinButton.VB_VarHelpID = -1
Private WithEvents mMinutesSpinButton As MSForms.SpinButton 'spm
Attribute mMinutesSpinButton.VB_VarHelpID = -1
Private WithEvents mHoursSpinButton As MSForms.SpinButton 'sph
Attribute mHoursSpinButton.VB_VarHelpID = -1
Private WithEvents mNumbersSpinButton As MSForms.SpinButton 'sph
Attribute mNumbersSpinButton.VB_VarHelpID = -1
Private WithEvents mDecimalsSpinButton As MSForms.SpinButton 'sph
Attribute mDecimalsSpinButton.VB_VarHelpID = -1
Private WithEvents mAMPMSpinButton As MSForms.SpinButton 'sptt
Attribute mAMPMSpinButton.VB_VarHelpID = -1

Private WithEvents mListBox As MSForms.ListBox 'li
Attribute mListBox.VB_VarHelpID = -1
Private WithEvents mListView As MSComctlLib.ListView 'lv
Attribute mListView.VB_VarHelpID = -1

Private WithEvents mMultiPage As MSForms.MultiPage 'mp
Attribute mMultiPage.VB_VarHelpID = -1

Private WithEvents mCommandButton As MSForms.CommandButton 'btn
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mFrame As MSForms.Frame 'fr
Attribute mFrame.VB_VarHelpID = -1
Private WithEvents mImage As MSForms.image 'img
Attribute mImage.VB_VarHelpID = -1
Private WithEvents mLabel As MSForms.label 'lb
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mOptionButton As MSForms.OptionButton 'op
Attribute mOptionButton.VB_VarHelpID = -1
Private WithEvents mScrollBar As MSForms.ScrollBar 'sc
Attribute mScrollBar.VB_VarHelpID = -1
Private WithEvents mTabStrip As MSForms.TabStrip 'ts
Attribute mTabStrip.VB_VarHelpID = -1
Private WithEvents mToggleButton As MSForms.ToggleButton 'tg
Attribute mToggleButton.VB_VarHelpID = -1
Property Set Parent(Parent As CtrlEvents): Set mParent = Parent: End Property
'Property Set Related(Related As CtrlEvents): Set mRelated = Related: mIsRequired = mRelated.Count > 0: End Property
'Property Get Related() As CtrlEvents: Set Related = mRelated: End Property
Property Set Related(Related As Collection): Set mRelatedCtrls = Related: End Property
Property Get Related() As Collection: Set Related = mRelatedCtrls: End Property
Property Get Name() As String: Name = mCtrl.Name: End Property
Property Let BaseName(BaseName As String): mBaseName = BaseName: End Property
Property Get BaseName() As String: BaseName = mBaseName: End Property
Property Let ManuallyEdited(ManuallyEdited As Boolean): mManuallyEdited = ManuallyEdited: End Property
Property Get ManuallyEdited() As Boolean
    ManuallyEdited = mManuallyEdited
    If Not ManuallyEdited Then
        On Error GoTo ex
        Dim xCtrl As control
        For Each xCtrl In mCtrl.Controls
            If mParent(xCtrl.Name).ManuallyEdited Then
                ManuallyEdited = True
                Exit Property
            End If
        Next
    End If
ex:
End Property
Property Let InvalidBackColor(InvalidBackColor As Long): mInvalidBackColor = InvalidBackColor: End Property
Property Get InvalidBackColor() As Long: InvalidBackColor = mInvalidBackColor: End Property
Property Get ValidBackColor() As Long: ValidBackColor = mValidBackColor: End Property
Property Let ValidationType(ValidationType As String): mValidationType = ValidationType: End Property
Property Get control(): Set control = mCtrl: End Property
Property Set control(control As control)
    Set mCtrl = control
    On Error Resume Next
    mValidBackColor = mCtrl.BackColor
    mValidForeColor = mCtrl.ForeColor
    mCtrlName = mCtrl.Name
    mCtrlType = TypeName(mCtrl)
    If mCtrlType Like "ListView*" Then mCtrlType = "ListView"
    If Len(mValidationType) = 0 Then mValidationType = "NotEmpty"
    Select Case mCtrlType
    Case "CheckBox": Set mCheckBox = mCtrl
    Case "TextBox"
        Set mTextBox = mCtrl
        If Len(mReqGroupName) Then
            Select Case Mid$(mCtrl.Name, 3, 1)
            Case "d": Set mDateTextbox = mCtrl: mValidationType = "Date"
            Case "m": Set mMinutesTextBox = mCtrl: mValidationType = "Minutes"
            Case "h": Set mHoursTextBox = mCtrl: mValidationType = "Hours"
            Case "n": Set mNumbersTextBox = mCtrl: mValidationType = "Numbers"
            Case "d": Set mDecimalsTextBox = mCtrl: mValidationType = "Decimals"
            End Select
        End If
    Case "ComboBox"
        Set mComboBox = mCtrl
'        If mIsRequired Then
        Select Case Mid$(mCtrl.Name, 3, 1)
        Case "t"
            If mCtrl.ListCount = 0 Then
                Set mAMPMComboBox = mCtrl
                mAMPMComboBox.AddItem "AM"
                mAMPMComboBox.AddItem "PM"
            End If
        End Select
'        End If
    Case "SpinButton"
        Set mSpinButton = mCtrl
'        If mIsRequired Then
        Select Case Mid$(mCtrl.Name, 3, 1)
        Case "t"
            Set mAMPMSpinButton = mCtrl
        Case "m"
            Set mMinutesSpinButton = mCtrl
            mCtrl.Min = 0
            mCtrl.Max = 59
            mCtrl.value = 45
            mCtrl.SmallChange = 15
        Case "h"
            Set mHoursSpinButton = mCtrl
            mCtrl.Min = 1
            mCtrl.Max = 12
            mCtrl.value = 11
        Case "n"
            Set mNumbersSpinButton = mCtrl
        Case "d"
            Set mDecimalsSpinButton = mCtrl
        End Select
        mCtrl.Min = mCtrl.Min - 1
        mCtrl.Max = mCtrl.Max + 1
'        End If
    Case "Frame": Set mFrame = mCtrl
    Case "Image": Set mImage = mCtrl
    Case "Label": Set mLabel = mCtrl
    Case "ListBox": Set mListBox = mCtrl
    Case "ListView": Set mListView = mCtrl
    Case "MultiPage": Set mMultiPage = mCtrl
    Case "OptionButton": Set mOptionButton = mCtrl
    Case "ScrollBar": Set mScrollBar = mCtrl
    Case "TabStrip": Set mTabStrip = mCtrl
    Case "ToggleButton": Set mToggleButton = mCtrl
    End Select
    On Error Resume Next
    mForecolor = mCtrl.ForeColor
End Property
Sub MakeRequired(ValidationType As String, InvalidBackColor As Long, Optional ByVal ReqGroupName As String = "Me")
    mValidationType = ValidationType
    If Len(ReqGroupName) = 0 Then ReqGroupName = "Me"
'    If InStr(mCtrl.Name, "tbdStarts") Then Stop
    Select Case mCtrlType
    Case "ComboBox"
        If Mid$(mCtrl.Name, 3, 1) <> "t" Then
            mCtrl.AddItem cbPlaceHolder & SplitWords(mBaseName) & " --", 0
            mCtrl.ListIndex = 0
        End If
    Case "TextBox"
        Select Case Mid$(mCtrl.Name, 3, 1)
        Case "d": Set mDateTextbox = mCtrl: mValidationType = "Date"
        Case "m": Set mMinutesTextBox = mCtrl: mValidationType = "Minutes"
        Case "h": Set mHoursTextBox = mCtrl: mValidationType = "Hours"
        Case "n": Set mNumbersTextBox = mCtrl: mValidationType = "Numbers"
        Case "d": Set mDecimalsTextBox = mCtrl: mValidationType = "Decimals"
        End Select
    End Select
    mInvalidBackColor = InvalidBackColor
'    If mCtrl.Name Like "op*" Then Stop
    mReqGroupName = ReqGroupName 'True'mIsRequired
    mManuallyEdited = False
End Sub
Private Sub mDateTextbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean): mDateTextbox.value = PickDate(mDateTextbox.value): Cancel = True: End Sub
Private Sub mDateTextbox_Enter(): mDateTextbox.value = PickDate(mDateTextbox.value): End Sub
'Private Sub tbDate_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyTab Then PickDate
''    t = Timer
'End Sub
Private Sub mDateTextbox_KeyUp(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If keyCode = vbKeyTab Then mDateTextbox.value = PickDate(mDateTextbox.value)
End Sub
Private Sub mDateTextbox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    mDateTextbox.value = PickDate(mDateTextbox.value)
End Sub
Private Sub mHoursTextBox_Change()
    Dim xCtrl As control
    On Error Resume Next
    Set xCtrl = mRelatedCtrls("tbm" & mBaseName) '.Control
    On Error GoTo 0
    On Error Resume Next
    If xCtrl Is Nothing Then Set xCtrl = mRelatedCtrls("cbm" & mBaseName) '.Control
    If Len(xCtrl.value) = 0 Then xCtrl.value = "00"
End Sub
'Private Sub mNumbersTextBox_Change()
'    Dim xCtrl As control
'    On Error Resume Next
'    Set xCtrl = mRelatedCtrls("tbm" & mBaseName) '.Control
'    On Error GoTo 0
'    On Error Resume Next
'    If xCtrl Is Nothing Then Set xCtrl = mRelatedCtrls("cbm" & mBaseName) '.Control
'    If Len(xCtrl.Value) = 0 Then xCtrl.Value = "00"
'End Sub

Private Sub mSpinButton_SpinDown()
    With mSpinButton
        If .value = .Min Then .value = RoundDownx(.Max - 1, .SmallChange)
    End With
    UpdateNumberTextBox
End Sub
Private Sub mSpinButton_SpinUp()
    With mSpinButton
        If .value = .Max Then .value = RoundDownx(.Min + 1, .SmallChange)
    End With
    UpdateNumberTextBox
End Sub
Private Sub UpdateNumberTextBox()
    Dim s As String
    s = Mid(mCtrl.Name, 3, 1)
    If s = LCase(s) Then
        Dim xCtrl As control
        On Error Resume Next
        Set xCtrl = mRelatedCtrls("tb" & s & mBaseName) '.Control
        On Error GoTo 0
        On Error Resume Next
        xCtrl.value = Format(mCtrl.value, "0")
    End If
End Sub
Private Sub mMinutesSpinButton_Change()
    Dim xCtrl As control
    On Error Resume Next
    Set xCtrl = mRelatedCtrls("tbm" & mBaseName) '.Control
    On Error GoTo 0
    On Error Resume Next
    If xCtrl Is Nothing Then Set xCtrl = mRelatedCtrls("cbm" & mBaseName) '.Control
    xCtrl.value = Format(mMinutesSpinButton.value, "00")
'    If Not mTextBox Is Nothing Then
'        mTextBox.Value = Format(mMinutesSpinButton.Value, "00")
'    ElseIf Not mComboBox Is Nothing Then
'        mComboBox.Value = Format(mMinutesSpinButton.Value, "00")
'    End If
End Sub
Private Sub mHoursSpinButton_Change()
    Dim xCtrl As control
    On Error Resume Next
    Set xCtrl = mRelatedCtrls("tbh" & mBaseName) '.Control
    On Error GoTo 0
    On Error Resume Next
    If xCtrl Is Nothing Then Set xCtrl = mRelatedCtrls("cbh" & mBaseName) '.Control
    xCtrl.value = Format(mHoursSpinButton.value, "0")
'    If Not mTextBox Is Nothing Then
'        mTextBox.Value = Format(mHoursSpinButton.Value, "0")
'    ElseIf Not mComboBox Is Nothing Then
'        mComboBox.Value = Format(mHoursSpinButton.Value, "0")
'    End If
End Sub
Private Sub mAMPMSpinButton_Change()
    Dim xCtrl As control
    On Error Resume Next
    Set xCtrl = mRelatedCtrls("tbtt" & mBaseName) '.Control
    On Error GoTo 0
    On Error Resume Next
    If xCtrl Is Nothing Then Set xCtrl = mRelatedCtrls("cbtt" & mBaseName) '.Control
    xCtrl.value = IIf(xCtrl.value = "PM", "AM", "PM")
'    If Not mTextBox Is Nothing Then
'        mTextBox.Value = Format(mHoursSpinButton.Value, "0")
'    ElseIf Not mComboBox Is Nothing Then
'        mComboBox.Value = Format(mHoursSpinButton.Value, "0")
'    End If
End Sub

Private Sub AllPageChildrenModified()
    Dim xCtrl As control
    For Each xCtrl In mCtrl.Item(mMPLastPage).Controls 'Source
        mParent(xCtrl.Name).ManuallyEdited = True
    Next
End Sub

'Required Controls
Private Sub RequiredUpdated()
    If mDisableEvents Then Exit Sub
    mManuallyEdited = True
'    Dim Prt As Object
'    Set Prt = mCtrl.Parent 'mParent
'    Do Until mParent Is Prt
'        Prt
'    Loop
    If Len(mReqGroupName) Then
        mParent.CheckOk
    End If
End Sub
Private Sub mCheckBox_Change(): RequiredUpdated: End Sub
Private Sub mTextBox_Change(): mTextBox.ControlTipText = mTextBox.value: RequiredUpdated: End Sub
Private Sub mComboBox_Change(): RequiredUpdated: End Sub
Private Sub mSpinButton_Change(): RequiredUpdated: End Sub
'Private Sub mListView_BeforeLabelEdit(Cancel As Integer): RequiredUpdated: End Sub
Private Sub mListView_ItemCheck(ByVal Item As MSComctlLib.ListItem): RequiredUpdated: End Sub
'Private Sub mCommandButton_Change(): RequiredUpdated: End Sub
'Private Sub mFrame_Change(): RequiredUpdated: End Sub
'Private Sub mImage_Change(): RequiredUpdated: End Sub
'Private Sub mLabel_Change(): RequiredUpdated: End Sub
'Private Sub mListBox_Change()
'    RequiredUpdated
'End Sub
'Private Sub mListBox_Click()
'    RequiredUpdated
'End Sub
Private Sub mListBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    RequiredUpdated
End Sub
Private Sub mMultiPage_Change()
    If mMPLastPage <> mMultiPage.SelectedItem.Index Then
        AllPageChildrenModified
        RequiredUpdated
        mMPLastPage = mMultiPage.SelectedItem.Index
    End If
End Sub
Private Sub mMultiPage_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mMPLastPage = mMultiPage.SelectedItem.Index
End Sub
Private Sub mMultiPage_MouseDown(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    mMPLastPage = mMultiPage.SelectedItem.Index
    'AllPageChildrenModified: RequiredUpdated: End Sub
End Sub
Private Sub mOptionButton_Change(): RequiredUpdated: End Sub
Private Sub mScrollBar_Change(): RequiredUpdated: End Sub
'Private Sub mTabStrip_Change(): RequiredUpdated: End Sub
Private Sub mToggleButton_Change(): RequiredUpdated: End Sub
'Function IsValidGroup() As Boolean
'    Dim Clr As Long, xCtrl As Control, Ev As CtrlEvent
''    If Not mRelatedCtrls Is Nothing Then
'    For Each xCtrl In mRelatedCtrls
'        Set Ev = mParent(xCtrl.Name)
'        If Ev.ManuallyEdited Then
'            If IsValid Then
'                Clr = Ev.ValidBackColor
'            Else 'If InvalidBackColor = -1 Then
'                Clr = Ev.InvalidBackColor
'            End If
'            Select Case TypeName(xCtrl)
'            Case "Label", "SpinButton"
'            Case Else: xCtrl.BackColor = Clr
'            End Select
'        End If
'    Next
''    End If
'End Function
Sub MarkAsValid(IsValid As Boolean)
    Dim clr As Long
    If ManuallyEdited Then
        If IsValid Then
            clr = ValidBackColor
        Else 'If InvalidBackColor = -1 Then
            clr = InvalidBackColor
        End If
        BackColor = clr
'        Select Case mCtrlType
'        Case "Label", "SpinButton"
'        Case Else: mCtrl.BackColor = Clr
'        End Select
    End If
End Sub
Private Property Let BackColor(clr)
    Select Case mCtrlType
    Case "Label", "SpinButton"
    Case "ListBox"
        If mCtrl.BackColor <> clr Then
            Dim oldListItems As New Collection, i As Long, ht As Single, li As Long
            ht = mCtrl.Height
            li = mCtrl.ListIndex
            If Not IsNull(mCtrl.List) Then
                For i = 0 To UBound(mCtrl.List)
                    If mCtrl.Selected(i) Then
                        oldListItems.Add i
        '                mCtrl.Selected(i) = False
                    End If
                Next
            End If
    '        mCtrl.Clear
    '!!!!!!!!!!! Changing the backcolor resets the selection, that's ONE thing
    ' BUT IT ALSO RESIZES the HEIGHT?!! WHY MS!!
            mCtrl.BackColor = clr
            
            mDisableEvents = True
            For i = 1 To oldListItems.Count
                mCtrl.Selected(oldListItems(i)) = True
            Next
            mDisableEvents = False
            mCtrl.ListIndex = li
            mCtrl.Height = ht
        End If
    Case "Frame"
        If clr = ValidBackColor Then
            mCtrl.BorderStyle = 0
            mCtrl.SpecialEffect = 3
        Else
            mCtrl.BorderStyle = 1
            mCtrl.BorderColor = clr
        End If
'    Case "ListView"
'        mCtrl.BorderStyle = IIf(Clr = ValidBackColor, 0, 1)
'        mCtrl.BorderColor = Clr
    Case "Page", "MultiPage", "ListView"
    Case Else: mCtrl.BackColor = clr
    End Select
End Property
Sub GrayOut(Optional Activate As Boolean, Optional Enabled As Boolean)
    Dim mFore As Long, mBack As Long, mActive As Boolean
    mActive = Activate Or Enabled
    mBack = mValidBackColor
    mFore = mValidForeColor
'    mBack = IIf(Activate, mValidBackColor, vbInactiveTitleBar)
'    mFore = IIf(Activate, mValidForeColor, vbInactiveTitleBar)
'    If Activate Then Enabled = ActivateBack
    Select Case mCtrlType
    Case "TextBox", "ComboBox"
        mBack = IIf(Activate, mValidBackColor, vbInactiveTitleBar)  'vbInactiveBorder
    Case "CheckBox", "OptionButton"
        mFore = IIf(Activate, mValidForeColor, vbInactiveTitleBar)
    Case "ListBox"
        mBack = IIf(Activate, mValidBackColor, vbInactiveTitleBar)
        mFore = IIf(Activate, mValidForeColor, vbInactiveTitleBar)
    End Select
    
    If mCtrl.BackColor <> mBack Then BackColor = mBack
    If mCtrl.ForeColor <> mFore Then mCtrl.ForeColor = mFore
    If mCtrl.Enabled <> (Activate Or Enabled) Then mCtrl.Enabled = (Activate Or Enabled)
End Sub
Function IsValid(Optional ByVal ValidationType As String = "NotEmpty", Optional ByVal InvalidBackColor As Long = vbRed) As Long
    Dim i As Long, xCtrl As control, Ev As CtrlEvent
'    Dim vvvv As ListView
'    If InStr(mCtrl.Name, "tbmStarts") Then Stop
    If InvalidBackColor <> -1 Then InvalidBackColor = mInvalidBackColor
    If Len(mValidationType) = 0 Then mValidationType = ValidationType Else ValidationType = mValidationType
    IsValid = -1
    Set xCtrl = mCtrl
    If xCtrl.Visible = False Then Exit Function
'    For Each xCtrl In mRelated
        Select Case mCtrlType 'TypeName(xCtrl)
        Case ""
        Case "ListBox"
            For i = 0 To xCtrl.ListCount - 1
                If xCtrl.Selected(i) Then Exit For
            Next
            IsValid = IIf(i < xCtrl.ListCount, 1, 0)
        Case "ListView"
            i = 0
            If xCtrl.Checkboxes Then
                For i = 1 To xCtrl.ListItems.Count
                    If xCtrl.ListItems(i).Checked Then Exit For
                Next
            End If
            IsValid = IIf(i <= xCtrl.ListItems.Count, 1, 0)
'            IsValid = i < xCtrl.ListItems.Count
        Case "ScrollBar"
'            IsValid = xCtrl.Value = xCtrl.Max
        Case "CheckBox", "ToggleButton"
            IsValid = IIf(xCtrl.value, 1, 0)
'            IsValid = IsValid And xCtrl.Value
        Case "OptionButton"
            For i = 1 To mRelatedCtrls.Count
                If TypeName(mRelatedCtrls(i)) = "OptionButton" Then
                    If mRelatedCtrls(i).value Then Exit For
                End If
            Next
            IsValid = IIf(i <= mRelatedCtrls.Count, 1, 0)
'            IsValid = i <= mRelatedCtrls.Count
        Case "TextBox", "ComboBox"
            Select Case Left$(mValidationType, 8)
            Case "Date": IsValid = IIf(Len(xCtrl.value) > 0 And IsDate(xCtrl.value), 1, 0)
            Case "Minutes": IsValid = IIf(xCtrl.value > -1 And xCtrl.value < 60, 1, 0)
            Case "Hours": IsValid = IIf(xCtrl.value > 0 And xCtrl.value < 13, 1, 0)
            Case "Numbers"
                If IsNumeric(xCtrl.value) Then
                    IsValid = IIf(Fix(xCtrl.value) = Val(xCtrl.value), 1, 0)
                Else
                    IsValid = 0
                End If
            Case "Decimals": IsValid = IIf(IsNumeric(xCtrl.value), 1, 0)
            Case "NotEmpty": IsValid = IIf(Len(xCtrl.value) > 0 And Not xCtrl.value Like cbPlaceHolder & "*", 1, 0)
            Case "Email": IsValid = IIf(xCtrl.value Like "*@*.*", 1, 0)
            Case Else: IsValid = IIf(xCtrl.value Like mValidationType, 1, 0)
            End Select
'        Case "Frame"
''            IsValid = False
'            For Each xCtrl In mRelatedCtrls
'                If TypeName(xCtrl) <> "Label" And xCtrl.Name <> mCtrl.Name Then
'                    If mParent.Item(xCtrl.Name).IsValid Then Exit For
'                End If
'            Next
'            IsValid = IIf(i <= mRelatedCtrls.Count, 1, 0)
''            IsValid = i <= mRelatedCtrls.Count
        End Select
'    Next
    'IIf(IsValid, mValidBackColor, IIf(InvalidBackColor = -1, mInvalidBackColor, InvalidBackColor))
''    If mManuallyEdited Then
'        If Not mRelatedCtrls Is Nothing Then
'            For Each xCtrl In mRelatedCtrls 'xCtrl In mRelated
'                Set Ev = mParent(xCtrl.Name)
'                If Ev.ManuallyEdited Then
'    '                Set xCtrl = Ev.Control
'                    If IsValid Then
'                        InvalidBackColor = Ev.ValidBackColor
'                    ElseIf InvalidBackColor = -1 Then
'                        InvalidBackColor = Ev.InvalidBackColor
'                    End If
'                    Select Case TypeName(xCtrl)
'                    Case "Label", "SpinButton"
'                    Case Else: xCtrl.BackColor = InvalidBackColor
'                    End Select
'                End If
'            Next
'        End If
''    End If
'If IsValid = 0 Then Stop
End Function

'Date Pickers
'Private Sub mDateTextbox_DblClick(ByVal Cancel As MSForms.ReturnBoolean): PickDate: Cancel = True: End Sub
'Private Sub mDateTextbox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): PickDate: End Sub
'Private Sub mDateTextbox_Enter(): PickDate: End Sub
'Private Sub mDateTextbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyTab Then PickDate
'End Sub
'Private Sub PickDate()
'    On Error Resume Next
'    Parent.Hide
'    With frmETRcalendar
'        If Len(mDateTextbox.Value) Then .SetDate = mDateTextbox.Value
'        .Show
'        mDateTextbox.Value = .SelectedDate
'    End With
'    Unload frmETRcalendar
'    Parent.Show
'End Sub
'Private Sub mDateComboBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean): PickDate: Cancel = True: End Sub
'Private Sub mDateComboBox_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single): PickDate: End Sub
'Private Sub mDateComboBox_Enter(): PickDate: End Sub
'Private Sub mDateComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = vbKeyTab Then PickDate
'End Sub
Private Function PickDate(Optional OldDate As String) As String
    On Error Resume Next
    mParent.UserForm.Hide
    With frmETRcalendar
        If Len(OldDate) Then .SetDate = OldDate
        .Show
        PickDate = .SelectedDate
    End With
    Unload frmETRcalendar
    mParent.UserForm.Show
End Function
Private Function SplitWords(CtrlName As String) As String
    Dim i As Long, FirstCapFound As Boolean
    For i = 1 To Len(CtrlName)
        Select Case Asc(Mid$(CtrlName, i, 1))
        Case 65 To 90
            If FirstCapFound Then SplitWords = SplitWords & " "
            FirstCapFound = True
        End Select
        If FirstCapFound Then SplitWords = SplitWords & Mid$(CtrlName, i, 1)
    Next
End Function


