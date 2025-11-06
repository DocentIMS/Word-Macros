Attribute VB_Name = "Ribbon_Tasks_Mod"
Option Explicit
''IdGroupTasksTraffic
'Private Function IdGroupTasksTrafficGetVisible(id As String): IdGroupTasksTrafficGetVisible = GetVisibleGroup(id): End Function


'Tasks
Function IdGroupTasksGetVisible(ID As String): IdGroupTasksGetVisible = GetVisibleGroup(ID): End Function
'Traffic
Private Function IdButtonTasksGreenGetImage()
    If Len(ProjectColorStr) Then Set IdButtonTasksGreenGetImage = GetTaskImage("Green")
End Function
'Private Function IdButtonTasksGreenGetLabel(): IdButtonTasksGreenGetLabel = GetTaskCount("Green"): End Function
Private Function IdButtonTasksGreenGetEnabled(): IdButtonTasksGreenGetEnabled = GetTaskCount("Green") > 0: End Function
Private Sub IdButtonTasksGreenOnAction(): GotoTaskCollection "Green": End Sub
Private Function IdButtonTasksGreenGetSupertip(): IdButtonTasksGreenGetSupertip = GetTasksTrafficTooltip("Green", "Future"): End Function

Private Function IdButtonTasksYellowGetImage():
    If Len(ProjectColorStr) Then Set IdButtonTasksYellowGetImage = GetTaskImage("Yellow")
End Function
'Private Function IdButtonTasksYellowGetLabel(): IdButtonTasksYellowGetLabel = GetTaskCount("Yellow"): End Function
Private Function IdButtonTasksYellowGetEnabled(): IdButtonTasksYellowGetEnabled = GetTaskCount("Yellow") > 0: End Function
Private Sub IdButtonTasksYellowOnAction(): GotoTaskCollection "Yellow": End Sub
Private Function IdButtonTasksYellowGetSupertip(): IdButtonTasksYellowGetSupertip = GetTasksTrafficTooltip("Yellow", "Soon"): End Function

Private Function IdButtonTasksRedGetImage()
    If Len(ProjectColorStr) Then Set IdButtonTasksRedGetImage = GetTaskImage("Red")
End Function
'Private Function IdButtonTasksRedGetLabel(): IdButtonTasksRedGetLabel = GetTaskCount("Red"): End Function
Private Function IdButtonTasksRedGetEnabled(): IdButtonTasksRedGetEnabled = GetTaskCount("Red") > 0: End Function
Private Sub IdButtonTasksRedOnAction(): GotoTaskCollection "Red": End Sub
Private Function IdButtonTasksRedGetSupertip(): IdButtonTasksRedGetSupertip = GetTasksTrafficTooltip("Red", "Urgent"): End Function

Sub IdButtonCreateTaskOnAction(): On Error Resume Next: frmCreateTask.Display True: End Sub
'Function IdButtonCreateTaskGetImage()
'    Dim DarkMode As Boolean
'    DarkMode = IsDarkModeSelected
'    Set IdButtonCreateTaskGetImage = MLoadPictureGDI.LoadPictureGDI("D:\Ongoing\23-06-15 - Wayne Glover (Word)\Old\Icons\Ribbon Icons\Task132" & IIf(DarkMode, "B", "W") & ImagesExtension)
'End Function

