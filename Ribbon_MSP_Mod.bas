Attribute VB_Name = "Ribbon_MSP_Mod"
Option Explicit

'MSP
Function IdGroupMSPGetVisible(ID As String): IdGroupMSPGetVisible = GetVisibleGroup(ID): End Function
Sub IdButtonMSPOnAction(): ImportMSP: End Sub
Function IdButtonMSPGetVisible(ID As String): IdButtonMSPGetVisible = GetVisibleGroup("IdGroupMSP") And Not GetButtonVisible(6): End Function
Function IdButtonMSPCancelGetVisible(ID As String): IdButtonMSPCancelGetVisible = GetButtonVisible(6): End Function
Sub IdButtonMSPCancelOnAction(): CancelEditingDoc: End Sub
Function IdButtonMSPUpdateGetVisible(ID As String): IdButtonMSPUpdateGetVisible = GetButtonVisible(6): End Function
Sub IdButtonMSPUpdateOnAction(): UpdateMSP False: End Sub

