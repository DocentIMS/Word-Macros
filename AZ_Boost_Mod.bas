Attribute VB_Name = "AZ_Boost_Mod"
'@Folder("Base")
Option Explicit
'23-06-16 23:59
Private BCount As Long, LastT As Single, DebugMode As Boolean
Sub Boost(Optional Flag As Boolean = True, _
        Optional ByVal ForceUnboost As Boolean = False)
    'Changes settings of Excel to make the code Application.run faster and look better while running.
    ' It also keeps track of levels of boosting.
    ' If you boost 2 times, you would need unboosting 2 times, unless ForceUnboost = True
    
    'Flag: True = Boost, Falst = Unboost. If omitted, Flag = True.
    'ForceUnboost: Ignore previous boost levels & unboost
    Dim Fnd As Range
    On Error Resume Next
    If ForceUnboost Then BCount = 0: Flag = False
    If Flag Then BCount = BCount + 1 Else BCount = BCount - 1
    If BCount < 0 Then BCount = 0
    If Flag And BCount = 1 Then
        Flag = Not Flag
        Application.ScreenUpdating = Flag
        Application.DisplayAlerts = Flag
    ElseIf BCount = 0 Then
        Application.StatusBar = vbNullString
        Flag = Not Flag
        Application.ScreenUpdating = Flag
        Application.DisplayAlerts = Flag
    End If
End Sub
Sub EndAll()
    Boost False, True
    End
End Sub
Sub PrintTimer(status As String)
    If Not DebugMode Then Exit Sub
    If LastT > 0 Then If Len(status) Then Debug.Print Format$(Timer - LastT, "0.00") & ": " & status
    LastT = Timer
End Sub


