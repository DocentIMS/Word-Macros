Attribute VB_Name = "AT_Planning_mod"
Option Explicit

Sub UploadPlanningDocument(Doc As Document)
    Dim Ps As New PlanItems, i As Long
    Dim FName As String, FPath As String
    
    FName = GetActiveFName(Doc)
'    If Not ActiveDocument.Saved Then
    If FName = "Blank" Then
        If MsgBox("This document must be saved first." & Chr(10) & _
                "Do you want to save it now?", vbExclamation + vbYesNo, "") <> vbYes Then
            Exit Sub
        Else
            ActiveDocument.Save
        End If
    End If
    If PNum = 0 Then GetSelectedProjectIndex
    Boost
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .PURL = ProjectURLStr
        .PName = ProjectNameStr
        .DocType = "Planning Document"
        .ContractNo = ContractNumberStr
        .DocState = ""
        .IsDocument = True
        .IsTemplate = False
        .Name = Doc.Name
        .DocURL = ""
        .DocVer = 1
    End With
    SetMetaData Doc
'    CodeIsRunning = True
    FPath = ActiveDocument.FullName
    FName = Environ("Temp") & "\" & GetFileName(FPath)
    ActiveDocument.SaveAs2 FName
    ProgressBar.BarsCount = 1
    ProgressBar.HideApplication = True
    ResetSetGlobals
    Set_SPos = 0
    Ps.CollectHighlights
    Ps.CollectComments
    Set Ps = Ps.GetRanges(HighlightsAndComments, False)
    Ps.ExportPlanningDocument
    On Error Resume Next
    ActiveDocument.Close False
    Documents.Open FPath 'FName '
    Kill FName
    Unload ProgressBar
    Boost False, True
    Application.Visible = True
'    CodeIsRunning = False
    FName = "Done uploading." '& vbNewLine & vbNewLine & _
            "Total number of files: " & Ps.Count & vbNewLine & _
            "  Planning items Created successfully: " & Ps.Count
    FPath = "_________________________________" & vbNewLine & _
            "Uploaded files can be found in:"
    Select Case frmMsgBox.Display(Array(FName, "", FPath, Ps.UploadURL), Array("OK", "Close"), Success, Links:=Array("", "", "", Ps.UploadURL))
    Case "OK", ""
        ActiveDocument.Close False
'        On Error Resume Next
'        GoToLink Ps.UploadURL
    End Select
'    Kill FName
'    For i = 1 To Ps.Count
'        Debug.Print Ps(i).HighlightRng.Text
'    Next
End Sub
Sub OpenPlanningDocument(): OpenAsDocentDocument "Planning Document": End Sub
