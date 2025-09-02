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

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1
Public IsBusy As Boolean

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    Dim isDocentDocument As Boolean
    Dim FName As String
    FName = GetFileName(GetActiveFName(Doc))
    'Here
'    RemRegLastSelectedP FName & "_ProjectName"
    On Error Resume Next
    DocsInfo.Remove Doc.Name
    If GetProperty(pDocType) = "DocentHelp" Then
        Dim CC As ContentControl
        Set CC = FindCC("NeverShowAgain", Doc)
        If CC.Checked Then SetNeverHelpAgain 3
'        Cancel = True
'        Doc.Close False
        Doc.Saved = True
    End If
'    Else
        isDocentDocument = GetProperty(pIsDocument)
        If isDocentDocument Then UpdateServerDictionary
        If isDocentDocument And Not CodeIsRunning Then
            If Not ActiveDocument.Saved And GetProperty(pDocType) <> "MS Project" Then
                
                Cancel = frmMsgBox.Display("Some changes have not been saved." & vbNewLine & _
                            "Docent documents can not be saved locally." & vbNewLine & vbNewLine & _
                            "You have two options:" & vbNewLine & _
                            " 1) Select ""Close without saving"" and loose changes." & vbNewLine & _
                            " 2) Select ""Return to Document"" and save your changes using the Docent toolbar.", _
                            Array("Close without saving", "Return to document"), Exclamation, "DocentIMS") = "Return to document"
                If Not Cancel Then
                    ActiveDocument.Saved = True
                    If Documents.Count = 2 Then
                        If Documents(1).Range.End = 1 And _
                                Not Documents(1).Name Like "DocentIMS_*" Then
                            Documents(1).Close False
                            
                        End If
                    End If
                End If
            End If
            Set FilteredSOWs = Nothing
            Set AllSOWs = Nothing
        End If
'    If Not Cancel Then If Documents.Count = 1 Then ClearTempReg
'    RefreshProject
    RefreshRibbon
    DocsInfo.Remove Doc.Name
End Sub
'Sub CleanEnding()
'    Dim X As Boolean
'    On Error Resume Next
'    X = GetProperty(pDocType) = "Project Scope"
'    On Error GoTo 0
'    If X Then
'        UpdateFooterVersion
'    Else
'        On Error Resume Next
'        X = GetProperty(pIsDocument)
'        On Error GoTo 0
''        If x Then CloseActiveDocument: Exit Sub
'    End If
'End Sub

'Private Sub App_DocumentBeforeSave(ByVal Doc As Document, SaveAsUI As Boolean, Cancel As Boolean)
'
'End Sub

Sub App_DocumentChange()
'    Debug.Print "AAAAAAAAAAAAAAAAAAAAAAA"
    If Application.Documents.Count Then
'        Invalidate
        Invalidate "FileSaveAs", True
        Invalidate "FileSave", True  'contextObject.Document
        Invalidate "TabSave", True  'contextObject.Document
        RefreshDictionary
    End If
'    MsgBox "Changed"
'    On Error Resume Next
'    Dim cc As CommandBarControl
'    Set cc = Application.CommandBars("Standard").Controls("&Undo Highlight")
'    If Not cc Is Nothing Then MsgBox "Caught You Highlighting!"
End Sub

Private Sub App_NewDocument(ByVal Doc As Document)
    App_DocumentOpen Doc
    On Error Resume Next
    DocsInfo.Add LoadDocInfo(Doc), Doc.Name
End Sub
Private Sub App_DocumentOpen(ByVal Doc As Document)
    On Error GoTo ex
    Doc.Windows(1).Activate
    DoEvents
    SetMetaData Doc
    On Error Resume Next
    If DocsInfo Is Nothing Then Set DocsInfo = New Collection
    DocsInfo.Remove Doc.Name
    DocsInfo.Add LoadDocInfo(Doc), Doc.Name
    RefreshRibbon
    DoEvents: Sleep 50
'    If Not IsBusy Then
'        RedefineRibbon True
'        DoEvents: Sleep 50
'        Invalidate
'    Else
'
'    End If
ex:
    IsBusy = False
End Sub

'Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
'    If Not CodeIsRunning Then RefreshRibbon
'End Sub
Private Sub App_Quit()
    SetHelpShown 0
    SaveRibbonID 0
    Set DocsInfo = Nothing
    ClearTempReg
End Sub

'Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
'    RefreshRibbon
'End Sub

'Private Sub App_XMLSelectionChange(ByVal Sel As Selection, ByVal OldXMLNode As XMLNode, ByVal NewXMLNode As XMLNode, Reason As Long)
'    MsgBox Sel
'End Sub
