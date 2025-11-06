Attribute VB_Name = "Ribbon_Manager_Mod"
Option Explicit

Function IdToggleButtonMgrModeGetVisible(ID As String): IdToggleButtonMgrModeGetVisible = GetVisibleGroup(ID): End Function
Sub IdToggleButtonMgrModeOnAction(Pressed As Boolean)
    PrjMgr = Not PrjMgr
    RefreshRibbon
End Sub
Function IdToggleButtonMgrModeGetPressed(): IdToggleButtonMgrModeGetPressed = True: End Function
Function IdToggleButtonMgrModeGetEnabled(): On Error Resume Next: IdToggleButtonMgrModeGetEnabled = True: End Function
Private Function IdToggleButtonMgrModeGetImage()
    If PrjMgr Then
        Set IdToggleButtonMgrModeGetImage = GetImage("PrjMgr")
    Else
        IdToggleButtonMgrModeGetImage = "MeetingsWorkspace"
    End If
End Function
Private Function IdToggleButtonMgrModeGetLabel(): IdToggleButtonMgrModeGetLabel = IIf(PrjMgr, "Prj Mgr", "Team") & " Mode": End Function
Private Function IdToggleButtonMgrModeGetScreenTip()
    IdToggleButtonMgrModeGetScreenTip = IIf(PrjMgr, "Project Manager", "Team") & " Mode" '     " & _
        "Click to switch to " & IIf(PrjMgr, "Team", "Project Manager") & " Mode"
End Function
Private Function IdToggleButtonMgrModeGetSuperTip()
    If PrjMgr Then
        IdToggleButtonMgrModeGetSuperTip = "Tools used by the project manager at the beginning of each project." & Chr(10) & Chr(10) & "Click to switch to Team Mode"
    Else
        IdToggleButtonMgrModeGetSuperTip = "Tools used by all team members throughout the project" & Chr(10) & Chr(10) & "Click to switch to Project Manager Mode"
    End If
End Function

