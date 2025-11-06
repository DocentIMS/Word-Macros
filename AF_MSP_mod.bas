Attribute VB_Name = "AF_MSP_mod"
Option Explicit
Private Const MSPTemplatePath = "/templates/manager-templates/MS Project.dotx"
Private Const MSPOutputPath = "/ms-project/"
'private cons XlPathTest=ThisDocument.Path & "\MSP\sample ms project pmp file.xlsx"
Sub ImportMSP()
    BusyRibbon = True
    CodeIsRunning = True
    UpdateMSP True ', GetXlDate, GetWrdDate
    CodeIsRunning = False
    BusyRibbon = False
End Sub
Function GetXlDate() As Date
    On Error GoTo ex
    GetXlDate = TimeFromTFormat(CStr(GetAPIContent(MSPOutputPath & "ms-project.xlsx").Data("modified"))) 'Ths is already server-time
'    Dim FSO As Object
'    If Len(XlPth) = 0 Then XlPth = ThisDocument.Path & "\MSP\sample ms project pmp file.xlsx"
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    GetXlDate = FSO.GetFile(XlPth).DateCreated
ex:
End Function
Function GetWrdDate() As Date
    On Error GoTo ex
    GetWrdDate = ParseDate(CStr(GetAPIFolder(MSPOutputPath, "ms_project", Array("date"))(1)("date")), APIDateTimeFormat)
'    If Len(MSPOutputPath) = 0 Then MSPOutputPath = ThisDocument.Path & "\MSP\MSP.docx"
'    Dim Doc As Document, DtStr As String
'    If Len(Dir(MSPOutputPath)) Then
'        CodeIsRunning = True
'        Set Doc = Documents.Open(MSPOutputPath)
'        DtStr = GetContentControl("MSPDateTime", Doc)
'        DtStr = Replace(DtStr, "-", "")
'        GetWrdDate = DateValue(DtStr) + TimeValue(DtStr)
'        Doc.Close False
'        CodeIsRunning = False
'    End If
ex:
End Function
Function UpdateMSP(Optional SilentMode As Boolean, Optional xlDt As Date, Optional wrdDt As Date) As Boolean
    Dim Doc As Document
    If xlDt = #12:00:00 AM# Then xlDt = GetXlDate
    If xlDt = #12:00:00 AM# Then
        frmMsgBox.Display "Excel file is missing. Contact the project manager.", , Critical
        Exit Function
    End If
    If wrdDt = #12:00:00 AM# Then wrdDt = GetWrdDate
    UpdateMSP = wrdDt < DateSerial(Year(xlDt), Month(xlDt), Day(xlDt)) + TimeSerial(Hour(xlDt), Minute(xlDt), 0)
    If UpdateMSP Then
        Set Doc = CreateNewMSP(MSPOutputPath & "ms-project.xlsx", xlDt)
        Dim wrdPth As String, pdfPth As String
        wrdPth = SaveForUpload("MS Project", Doc)
        pdfPth = Left$(wrdPth, InStrRev(wrdPth, ".")) & "pdf"
        Doc.ExportAsFixedFormat2 pdfPth, wdExportFormatPDF, False, wdExportOptimizeForOnScreen, BitmapMissingFonts:=False
        'ActiveDocument.SaveAs2 pdfPth, wdFormatPDF, AddToRecentFiles:=False
        UploadAPIFile pdfPth, MSPOutputPath, Overwrite:=True
        If wrdDt = #12:00:00 AM# Then
            CreateAPIContent "ms_project", MSPOutputPath, Array("file", "date", "default_view"), _
                    Array(wrdPth, Format(xlDt, APIDateTimeFormat), "@@display-file")
        Else
            UpdateAPIContent MSPOutputPath & "ms-project.docx", Array("file", "date"), Array(wrdPth, Format(xlDt, APIDateTimeFormat))
        End If
    Else
        If SilentMode Then
            Set OpeningDocInfo = New DocInfo
            With OpeningDocInfo
                .DocType = "MS Project"
                .ContractNo = ContractNumberStr
                .IsDocument = True
                .PName = ProjectNameStr
                .PURL = ProjectURLStr
                .DocCreateDate = Format(ToServerTime, DateTimeFormat)
            End With
            Documents.Open DownloadAPIFile(MSPOutputPath & "ms-project.docx")
        Else
            MsgBox "You are already using the most recent MS Project Excel file"
        End If
    End If
End Function
Function CreateNewMSP(ExcelPth As String, xlDt As Date) As Document
    Dim ExcelApp As Object, WasOpen As Boolean, Data As Variant, WB As Object
    Dim xlPth As String
    xlPth = DownloadAPIFile(ExcelPth)
    On Error Resume Next
    Set ExcelApp = GetObject(, "Excel.Application")
    WasOpen = Not ExcelApp Is Nothing
    If Not WasOpen Then Set ExcelApp = CreateObject("Excel.Application")
    ExcelApp.EnableEvents = False
    ExcelApp.ScreenUpdating = False
    ExcelApp.Visible = False
    ExcelApp.DisplayAlerts = False
    Set WB = ExcelApp.Workbooks.Open(xlPth)
    Data = WB.Sheets(1).UsedRange.value
    WB.Close False
    ExcelApp.EnableEvents = True
    ExcelApp.ScreenUpdating = True
    ExcelApp.DisplayAlerts = True
    If WasOpen Then
        ExcelApp.Visible = True
    Else
        ExcelApp.Quit
    End If
    Kill xlPth
    
    'Application.ScreenUpdating = False
    Dim i As Long, r As Long, Doc As Document, wrdPth As String, c As Long
    ', OData() As String
    Set OpeningDocInfo = New DocInfo
    With OpeningDocInfo
        .DocType = "MS Project"
        .ContractNo = ContractNumberStr
        .IsDocument = True
        .PName = ProjectNameStr
        .PURL = ProjectURLStr
'        .DocCreateDate = Format(ToServerTime, DateFormat)
    End With
    wrdPth = DownloadAPIFile(MSPTemplatePath)
    Set Doc = Documents.Add(wrdPth)
    Unprotect Doc
    SetContentControl "MSPDateTime", Format(ToServerTime(CStr(xlDt)), DateTimeFormat), Doc
    SetContentControl "Project Name", ProjectNameStr, Doc
    r = 1
    ReDim OData(1 To 6, 1 To 1)
    
    ProgressBar.BarsCount = 1
'    ProgressBar.HideApplication = True
    ProgressBar.Reset
    ProgressBar.Progress , "Generating MS Project word file...", 0
'    On Error GoTo ex
    ProgressBar.Dom = UBound(Data)
    ProgressBar.Show
    For i = 2 To UBound(Data)
        If ProgressBar.Progress Then
            MsgBox "Cancelled by user.", vbExclamation, ""
            GoTo can
        End If
        If Data(i, 3) = "Yes" Then 'Active tasks only
            r = r + 1
            If r > 2 Then Doc.Tables(1).Rows.Add
'            ReDim Preserve OData(1 To 6, 1 To r)
            c = 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Data(i, 1)
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Data(i, 2)
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Space$(Data(i, 10) * 2) & Data(i, 5)
            Doc.Tables(1).Rows(r).Cells(c).Range.Font.Size = 14 - Data(i, 10)
            Select Case Data(i, 10)
            Case 0
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Color = wdColorBlack
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Bold = True
            Case 1
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Color = wdColorBlue
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Bold = True
            Case 2
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Color = wdColorDarkGreen
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Bold = False
            Case 3
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Color = wdColorPlum
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Bold = False
            Case 4
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Color = wdColorLightTurquoise
                Doc.Tables(1).Rows(r).Cells(c).Range.Font.Bold = False
            End Select
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Split(Data(i, 6))(0)
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Format(Data(i, 7), "mmm d")
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Format(Data(i, 8), "mmm d")
            c = c + 1
            Doc.Tables(1).Rows(r).Cells(c).Range.text = Data(i, 11)
        End If
    Next
    SetProperty pDocCreateDate, Format(ToServerTime, DateFormat), Doc
can:
    Protect Doc
    Doc.SaveAs2 MSPOutputPath, wdOpenFormatXMLDocument
    Doc.Saved = True
    Application.ScreenUpdating = True
    Kill wrdPth
    Unload ProgressBar
    Set CreateNewMSP = Doc
End Function
