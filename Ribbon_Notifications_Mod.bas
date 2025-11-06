Attribute VB_Name = "Ribbon_Notifications_Mod"
Option Explicit

'Notifications
Private Function IdGroupNotificationsGetVisible(ID As String): IdGroupNotificationsGetVisible = GetVisibleGroup(ID): End Function
'Traffic

''IdGroupNotificationsTraffic
'Private Function IdGroupNotificationsTrafficGetVisible(id As String): IdGroupNotificationsTrafficGetVisible = GetVisibleGroup(id): End Function

Private Function IdButtonNotificationsGreenGetImage()
    If Len(ProjectColorStr) Then Set IdButtonNotificationsGreenGetImage = GetNotificationImage("Aqua")
End Function
'Private Function IdButtonNotificationsGreenGetLabel(): IdButtonNotificationsGreenGetLabel = GetNotificationCount("Aqua"): End Function
Private Function IdButtonNotificationsGreenGetVisible(ID As String): IdButtonNotificationsGreenGetVisible = Not PrjMgr: End Function
'Private Function IdButtonNotificationsGreenGetEnabled(): IdButtonNotificationsGreenGetEnabled = GetNotificationCount("Aqua") > 0: End Function
Private Sub IdButtonNotificationsGreenOnAction(): GotoNotificationCollection "Aqua": End Sub
Private Function IdButtonNotificationsGreenGetSupertip(): IdButtonNotificationsGreenGetSupertip = GetNotificationCount("Aqua") & " Informational Notification" & IIf(GetNotificationCount("Aqua") > 1, "s", ""): End Function
'Private Function IdButtonNotificationsGreenGetSupertip(): IdButtonNotificationsGreenGetSupertip = "": End Function

Private Function IdButtonNotificationsYellowGetImage():
    If Len(ProjectColorStr) Then Set IdButtonNotificationsYellowGetImage = GetNotificationImage("Yellow")
End Function
'Private Function IdButtonNotificationsYellowGetLabel(): IdButtonNotificationsYellowGetLabel = GetNotificationCount("Yellow"): End Function
Private Function IdButtonNotificationsYellowGetVisible(ID As String): IdButtonNotificationsYellowGetVisible = Not PrjMgr: End Function
'Private Function IdButtonNotificationsYellowGetEnabled(): IdButtonNotificationsYellowGetEnabled = GetNotificationCount("Yellow") > 0: End Function
Private Sub IdButtonNotificationsYellowOnAction(): GotoNotificationCollection "Yellow": End Sub
Private Function IdButtonNotificationsYellowGetSupertip(): IdButtonNotificationsYellowGetSupertip = GetNotificationCount("Yellow") & " Important Notification" & IIf(GetNotificationCount("Yellow") > 1, "s", ""): End Function
'Private Function IdButtonNotificationsYellowGetSupertip(): IdButtonNotificationsYellowGetSupertip = "": End Function

Private Function IdButtonNotificationsRedGetImage()
    If Len(ProjectColorStr) Then Set IdButtonNotificationsRedGetImage = GetNotificationImage("Red")
End Function
'Private Function IdButtonNotificationsRedGetLabel(): IdButtonNotificationsRedGetLabel = GetNotificationCount("Red"): End Function
Private Function IdButtonNotificationsRedGetVisible(ID As String): IdButtonNotificationsRedGetVisible = Not PrjMgr: End Function
'Private Function IdButtonNotificationsRedGetEnabled(): IdButtonNotificationsRedGetEnabled = GetNotificationCount("Red") > 0: End Function
Private Sub IdButtonNotificationsRedOnAction(): GotoNotificationCollection "Red": End Sub
Private Function IdButtonNotificationsRedGetSupertip(): IdButtonNotificationsRedGetSupertip = GetNotificationCount("Red") & " Critical Notification" & IIf(GetNotificationCount("Red") > 1, "s", ""): End Function
'Private Function IdButtonNotificationsRedGetSupertip(): IdButtonNotificationsRedGetSupertip = "": End Function

'Private Sub IdButtonCreateNotificationOnAction(): On Error Resume Next: frmNotification.Show: End Sub
Private Sub IdButtonCreateNotificationOnAction()
    On Error Resume Next
    If PrjMgr Then frmCreateNotification.Show Else GotoNotificationCollection "All"
End Sub
Private Function IdButtonCreateNotificationGetVisible(ID As String): IdButtonCreateNotificationGetVisible = True: End Function
Private Function IdButtonCreateNotificationGetLabel(): IdButtonCreateNotificationGetLabel = IIf(PrjMgr, "Create Notif.", "Notifications"): End Function
Private Function IdButtonCreateNotificationGetSupertip()
    Dim s As String
    'Currently, we have this (Before I start making edits):
    If PrjMgr Then
        s = "Avoid email and send a notification to any/all team members." & vbNewLine & vbNewLine & _
            "Notifications are displayed as a horizontal bar on the project website"
    Else
        s = "Click to go to your notifications on Plone" '(Without filters) '?"
    End If
        
'    'Previously, we had this:
'    If PrjMgr Then
'        'If this is the project manager, we show this:
'        s = "Notifications are displayed as a horizontal bar on the website."
'    Else
'        'If team member,we show this
'        s = "Only PM can send notifications, but you can read yours."
'    End If
    
    IdButtonCreateNotificationGetSupertip = s
End Function
Private Function IdButtonCreateNotificationGetScreentip()
    IdButtonCreateNotificationGetScreentip = IIf(PrjMgr, "Create Notification", "Notifications")
End Function

'Function IdButtonCreateNotificationGetImage()
'    Dim DarkMode As Boolean
'    DarkMode = IsDarkModeSelected
'    Set IdButtonCreateNotificationGetImage = MLoadPictureGDI.LoadPictureGDI("D:\Ongoing\23-06-15 - Wayne Glover (Word)\Old\Icons\Ribbon Icons\Notification132" & IIf(DarkMode, "B", "W") & ImagesExtension)
'End Function

'''Notifications
'Private Function IdButtonNotificationsGetLabel()
'    Dim i As Long
''    If IsProjectSelected(True) Then
'        i = GetNotificationsCount
'        IdButtonNotificationsGetLabel = i & " Notification" & IIf(i > 1, "s", "")
''    End If
'End Function
'Private Function IdButtonNotificationsGetEnabled()
''    If IsProjectSelected(True) Then
'    IdButtonNotificationsGetEnabled = GetNotificationsCount > 0
'End Function
'Private Function IdButtonNotificationsGetVisible(id As String)
''    If IsProjectSelected(True) Then
'    IdButtonNotificationsGetVisible = GetButtonVisible(1)
'End Function
'Private Sub IdButtonNotificationsOnAction()
''    If IsProjectSelected(True) Then
'    GoToLink ProjectURLStr & "/notifications"
'End Sub
