Attribute VB_Name = "AD_Export_mod"
Option Explicit
Option Compare Text
Private Const CurrentMod = "Export_mod"
Private SeqNo As Long

Sub ExportActiveDocument(DocType As String)
    WriteLog 1, CurrentMod, "ExportActiveDocument", "Creating Output Folder"
    Dim FName As String, BtnClicked As String, Arr As Variant
    Dim Resplit As Boolean
    If Set_Cancelled Then GoTo can
    On Error GoTo ex
    Boost
    CodeIsRunning = True
    ProgressBar.BarsCount = 1
    ProgressBar.HideApplication = True
    ProgressBar.Spin
    OPath = Set_Odir & Application.PathSeparator & "DocentIMS Analysis" & Application.PathSeparator
    HTMLPath = OPath & "HTML Documents" & Application.PathSeparator
    DeleteFolder OPath
    CreateDir HTMLPath
    ProgressBar.Spin
    FName = SDoc.FullName
'    FName = GetFileName(FName, False)
'    FName = OPath & FName & "_AllowsMacros.docm"
    SeqNo = 0
    If Not Resplit Then
        WriteLog 1, CurrentMod, "ExportActiveDocument", "Saving the whole document as HTML"
        CleanFile
        UpdateSearchRange
        ProgressBar.Spin
        If Set_Cancelled Then GoTo can
        SDoc.SaveAs2 FName & ".html", wdFormatXMLDocumentMacroEnabled
        ProgressBar.Spin
        If Set_Cancelled Then GoTo can
    End If
    Application.ScreenUpdating = False
    Application.DisplayAlerts = wdAlertsNone
    CollectWords
    Arr = FixAndExport
    If IsEmpty(Arr) Then GoTo can
    SDoc.Close False
    Boost False, True
    '"Success" & vbNewLine & vbNewLine &
    frmMsgBox.Display "Words in the source file: " & Arr(1) & vbNewLine & _
            "Total words in the output files: " & Arr(2), "OK", Success, "Success"
    Unload ProgressBar
    frmConfirmUpload.Label1.Caption = "You are uploading to """ & ProjectNameStr & """ project." & Chr(10) & _
                "Type ""yes"" if you want to continue."
    If FullColor(ProjectColorStr).TooDark Then frmConfirmUpload.Label1.ForeColor = 16777215
    frmConfirmUpload.BackColor = ProjectColorStr
    frmConfirmUpload.Show
    On Error GoTo ex2
    BtnClicked = LCase(frmConfirmUpload.TextBox1)
    Unload frmConfirmUpload
    On Error GoTo ex
    If BtnClicked = "yes" Then
        UploadDoc Documents.Open(FName), SilentMode:=True
        StartUploading DocType, OPath, HTMLPath
        CloseLog
    Else
        DeleteFolder OPath
        MsgBox "File Uploading Cancelled.", vbExclamation, ""
        WriteLog 1, CurrentMod, "ExportActiveDocument", "Deleting Temp folder"
        CloseLog
    End If
    PrintView
    CodeIsRunning = False
    If Application.Documents.Count = 0 Then
        Application.Quit
    Else
        Application.Visible = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = wdAlertsAll
    End If
    
'    If InputBox("Do you want to upload?" & vbNewLine & "Type ""yes"" to start.", vbNullString) = "yes" Then StartUploading
    Exit Sub
can:
ex2:
    PrintView
    Boost False, True
'    CodeIsRunning = False
    'Application.Visible = True
    Exit Sub
ex:
    PrintView
    Boost False, True
    WriteLog 3, CurrentMod, "ExportActiveDocument", Err.Number & ":" & Err.Description
    MsgBox "Try again or use help button to send feedback.", vbCritical, "Unknown Error"
'    Stop
''    Resume
'    CodeIsRunning = False
End Sub
'Private Function IsHeaderParagraph(ByVal Para As Paragraph) As Boolean
'    Dim s As String
'    s = Para.Style
'    IsHeaderParagraph = s = "Heading Docent"
'    If Not IsHeaderParagraph Then IsHeaderParagraph = s Like "Heading [1-" & HLevel & "]"
'End Function
Sub CleanFile()
    WriteLog 1, CurrentMod, "CleanFile", "Removing Table of contents and Images"
    Dim p As Long
    Dim xSec As Section, HdFt As HeaderFooter
    Dim SecRng As Range
    Unprotect SDoc
    On Error GoTo ex
'    Set SDoc = ActiveDocument
    For Each xSec In SDoc.Sections
        For Each HdFt In xSec.Headers
            If HdFt.LinkToPrevious = False Then HdFt.Range.Delete
        Next
        For Each HdFt In xSec.Footers
            If HdFt.LinkToPrevious = False Then HdFt.Range.Delete
        Next
    Next
    Set SecRng = SDoc.Range
    For p = SDoc.Shapes.Count To 1 Step -1
        SDoc.Shapes(1).Delete
    Next
    Set_SPos = SDoc.TablesOfContents(1).Range.End + 1
    SecRng.SetRange 1, Set_SPos - 1
    ExportRange SecRng, Nothing, Nothing, "Front Pages", HTMLPath, 0, Nothing
    SecRng.Delete
    Set_SPos = 0
    Exit Sub
ex:
    WriteLog 3, CurrentMod, "CleanFile", Err.Number & ":" & Err.Description
''    Stop
''    Resume
End Sub
Function FixAndExport() As Variant
    WriteLog 1, CurrentMod, "FixAndExport", "Splitting Sections and exporting as HTML"
    Dim mSOW As SOW
    Dim SecRng As Range, pRng As Range, NRng As Range, DRng As Range
    Dim n As Long, p As Long, x As Long, nd As Long, i As Long
    Dim Arr(1 To 2) As Long
    Set SecRng = SDoc.Range
    Set pRng = SDoc.Range
    Set NRng = SDoc.Range
    Set DRng = SDoc.Range
    

    ProgressBar.Reset
    ProgressBar.Progress , "Exporting File No. 1", 0
    On Error GoTo ex
    ProgressBar.Dom = FilteredSOWs.Count
    
    Arr(1) = Set_SearchRange.ComputeStatistics(0)
    
    For p = 1 To FilteredSOWs.Count
        'If p = 4 Then Stop
        If ProgressBar.Progress(, "Exporting File No. " & p) Then
            Unload ProgressBar
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        If Set_Cancelled Then GoTo can
        Set mSOW = FilteredSOWs(p)
        Set SecRng = mSOW.SectionRng
        Set DRng = mSOW.DelivRange
        If p = 1 Then
            pRng.SetRange 0, 0
        Else
            pRng.SetRange FilteredSOWs(p - 1).SectionStart, FilteredSOWs(p - 1).SectionEnd
        End If
        If p = FilteredSOWs.Count Then
            NRng.SetRange 0, 0
        Else
            NRng.SetRange FilteredSOWs(p + 1).SectionStart, FilteredSOWs(p + 1).SectionEnd
        End If
        SeqNo = SeqNo + 1
        ExportRange SecRng, pRng, NRng, mSOW.FullName, HTMLPath, SeqNo, DRng
        SecRng.Select
        Arr(2) = Arr(2) + FilteredSOWs(p).CountWords 'SecRng.ComputeStatistics(0)
        'Debug.Print Arr(2)
    Next
    ProgressBar.Finished
    FixAndExport = Arr
can:
    Exit Function
ex:
    WriteLog 3, CurrentMod, "FixAndExport", Err.Number & ":" & Err.Description
End Function


