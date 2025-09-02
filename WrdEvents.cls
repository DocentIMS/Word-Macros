VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WrdEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Constants for better maintainability
Private Const DOC_TYPE_DOCENT_HELP As String = "DocentHelp"
Private Const DOC_TYPE_MS_PROJECT As String = "MS Project"
Private Const DOC_TYPE_PROJECT_SCOPE As String = "Project Scope"
Private Const DOCENT_TEMP_PREFIX As String = "DocentIMS_"

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private m_IsBusy As Boolean
Private m_Logger As ILogger ' Interface for logging
Private m_DocumentManager As DocumentManager ' Separate class for document operations

' Property for thread-safe busy state
Public Property Get IsBusy() As Boolean
    IsBusy = m_IsBusy
End Property

Public Property Let IsBusy(ByVal value As Boolean)
    m_IsBusy = value
End Property

' Initialize with dependency injection
Public Sub Initialize(ByVal logger As ILogger, ByVal docManager As DocumentManager)
    Set m_Logger = logger
    Set m_DocumentManager = docManager
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    m_Logger.LogInfo "Document closing: " & Doc.Name
    
    ' Handle different document types with separate methods
    If IsDocentHelpDocument(Doc) Then
        HandleDocentHelpClose Doc, Cancel
    ElseIf IsDocentDocument(Doc) Then
        HandleDocentDocumentClose Doc, Cancel
    End If
    
    ' Clean up document info
    m_DocumentManager.RemoveDocumentInfo Doc.Name
    RefreshRibbon
    
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error in App_DocumentBeforeClose: " & Err.Description
    ' Don't cancel close on error, but ensure cleanup
    m_DocumentManager.RemoveDocumentInfo Doc.Name
End Sub

Private Sub HandleDocentHelpClose(ByVal Doc As Document, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    Dim CC As ContentControl
    Set CC = FindCC("NeverShowAgain", Doc)
    
    If Not CC Is Nothing And CC.Checked Then
        SetNeverHelpAgain 3
    End If
    
    Doc.Saved = True ' Mark as saved to prevent save prompt
    
Cleanup:
    Set CC = Nothing
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error handling Docent Help close: " & Err.Description
    GoTo Cleanup
End Sub

Private Sub HandleDocentDocumentClose(ByVal Doc As Document, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    If CodeIsRunning Then Exit Sub
    
    ' Update server dictionary
    If Not m_DocumentManager.UpdateServerDictionary() Then
        m_Logger.LogWarning "Failed to update server dictionary"
    End If
    
    ' Handle unsaved changes
    If Not Doc.Saved And GetProperty(pDocType) <> DOC_TYPE_MS_PROJECT Then
        Cancel = ShowUnsavedChangesDialog()
        
        If Not Cancel Then
            Doc.Saved = True
            CleanupEmptyDocuments
        End If
    End If
    
    ' Clean up collections
    m_DocumentManager.ClearSOWCollections
    
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error handling Docent document close: " & Err.Description
    ' Continue with cleanup even on error
    m_DocumentManager.ClearSOWCollections
End Sub

Private Function ShowUnsavedChangesDialog() As Boolean
    Dim message As String
    Dim buttons As Variant
    
    message = "Some changes have not been saved." & vbNewLine & _
              "Docent documents can not be saved locally." & vbNewLine & vbNewLine & _
              "You have two options:" & vbNewLine & _
              " 1) Select ""Close without saving"" and lose changes." & vbNewLine & _
              " 2) Select ""Return to Document"" and save your changes using the Docent toolbar."
    
    buttons = Array("Close without saving", "Return to document")
    
    ShowUnsavedChangesDialog = (frmMsgBox.Display(message, buttons, Exclamation, "DocentIMS") = "Return to document")
End Function

Private Sub CleanupEmptyDocuments()
    On Error Resume Next ' OK here as this is cleanup
    
    If Documents.Count = 2 Then
        Dim firstDoc As Document
        Set firstDoc = Documents(1)
        
        If firstDoc.Range.End = 1 And Not firstDoc.Name Like DOCENT_TEMP_PREFIX & "*" Then
            firstDoc.Close False
        End If
    End If
End Sub

Private Sub App_DocumentChange()
    On Error GoTo ErrorHandler
    
    If Application.Documents.Count > 0 Then
        InvalidateRibbonControls
        m_DocumentManager.RefreshDictionary
    End If
    
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error in App_DocumentChange: " & Err.Description
End Sub

Private Sub InvalidateRibbonControls()
    ' Centralized ribbon invalidation
    Dim controlsToInvalidate As Variant
    Dim i As Integer
    
    controlsToInvalidate = Array("FileSaveAs", "FileSave", "TabSave")
    
    For i = 0 To UBound(controlsToInvalidate)
        Invalidate controlsToInvalidate(i), True
    Next i
End Sub

Private Sub App_NewDocument(ByVal Doc As Document)
    On Error GoTo ErrorHandler
    
    App_DocumentOpen Doc
    m_DocumentManager.AddDocumentInfo Doc
    
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error in App_NewDocument: " & Err.Description
End Sub

Private Sub App_DocumentOpen(ByVal Doc As Document)
    On Error GoTo ErrorHandler
    
    IsBusy = True
    m_Logger.LogInfo "Opening document: " & Doc.Name
    
    ' Activate and prepare document
    Doc.Windows(1).Activate
    DoEvents
    
    SetMetaData Doc
    m_DocumentManager.AddDocumentInfo Doc
    RefreshRibbon
    
    ' Small delay for UI updates - consider removing if not needed
    DoEvents
    Application.Wait Now + TimeValue("00:00:00.05")
    
    Exit Sub
    
ErrorHandler:
    m_Logger.LogError "Error in App_DocumentOpen: " & Err.Description
    IsBusy = False
End Sub

Private Sub App_Quit()
    On Error Resume Next ' OK for cleanup
    
    m_Logger.LogInfo "Application shutting down"
    
    SetHelpShown 0
    SaveRibbonID 0
    
    m_DocumentManager.Cleanup
    ClearTempReg
    
    ' Release object references
    Set m_DocumentManager = Nothing
    Set m_Logger = Nothing
End Sub

' Helper methods
Private Function IsDocentHelpDocument(ByVal Doc As Document) As Boolean
    On Error Resume Next
    IsDocentHelpDocument = (GetProperty(pDocType) = DOC_TYPE_DOCENT_HELP)
End Function

Private Function IsDocentDocument(ByVal Doc As Document) As Boolean
    On Error Resume Next
    IsDocentDocument = GetProperty(pIsDocument)
End Function

' Cleanup
Private Sub Class_Terminate()
    Set m_DocumentManager = Nothing
    Set m_Logger = Nothing
End Sub
