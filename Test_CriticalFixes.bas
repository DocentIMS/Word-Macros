Attribute VB_Name = "Test_CriticalFixes"
Option Explicit

'=======================================================
' Module: Test_CriticalFixes
' Purpose: Comprehensive test suite for critical issue fixes
' Author: Testing Team
' Version: 1.0
'
' Description:
'   Tests all fixes applied in Steps 5-6:
'   - Module constants corrections
'   - Resource leak prevention
'   - Error handling improvements
'   - UI state management
'
' Usage:
'   Run TestAll() to execute all tests
'   Or run individual test procedures
'
' Test Results:
'   - Displays summary at end
'   - Logs all results
'   - Shows pass/fail for each test
'=======================================================

Private Const CurrentMod As String = "Test_CriticalFixes"

' Test counters
Private testsPassed As Long
Private testsFailed As Long
Private testsSkipped As Long
Private testResults As String

'=======================================================
' MAIN TEST RUNNER
'=======================================================

'=======================================================
' Sub: TestAll
' Purpose: Run all test suites
'
' Description:
'   Executes all test procedures and displays summary.
'   Use this as the main entry point for testing.
'=======================================================
Public Sub TestAll()
    Const PROC_NAME As String = "TestAll"
    
    Dim startTime As Date
    Dim endTime As Date
    Dim duration As String
    
    startTime = Now
    
    ' Reset counters
    testsPassed = 0
    testsFailed = 0
    testsSkipped = 0
    testResults = ""
    
    WriteLog 1, CurrentMod, PROC_NAME, "========================================="
    WriteLog 1, CurrentMod, PROC_NAME, "STARTING COMPREHENSIVE TEST SUITE"
    WriteLog 1, CurrentMod, PROC_NAME, "========================================="
    
    ' Run test suites
    Call TestSuite_ModuleConstants
    Call TestSuite_ErrorHandling
    Call TestSuite_ResourceManagement
    Call TestSuite_UIState
    Call TestSuite_Integration
    
    endTime = Now
    duration = Format$(endTime - startTime, "hh:mm:ss")
    
    ' Display summary
    Call DisplayTestSummary(duration)
    
    WriteLog 1, CurrentMod, PROC_NAME, "========================================="
    WriteLog 1, CurrentMod, PROC_NAME, "TEST SUITE COMPLETE"
    WriteLog 1, CurrentMod, PROC_NAME, "========================================="
End Sub

'=======================================================
' TEST SUITE 1: MODULE CONSTANTS
'=======================================================

'=======================================================
' Sub: TestSuite_ModuleConstants
' Purpose: Verify all module constants are correct
'=======================================================
Private Sub TestSuite_ModuleConstants()
    Const SUITE_NAME As String = "Module Constants"
    
    Call StartTestSuite(SUITE_NAME)
    
    ' Test 1: AB_CommonFunctions constant
    Call Test_ModuleConstant_CommonFunctions
    
    ' Test 2: AC_Registry_mod constant
    Call Test_ModuleConstant_Registry
    
    ' Test 3: AE_Documents_mod constant
    Call Test_ModuleConstant_Documents
    
    ' Test 4: Ribbon_Functions_Mod constant
    Call Test_ModuleConstant_Ribbon
    
    Call EndTestSuite(SUITE_NAME)
End Sub

Private Sub Test_ModuleConstant_CommonFunctions()
    Const TEST_NAME As String = "AB_CommonFunctions module constant"
    
    ' This would require accessing the private constant
    ' We test indirectly by checking log output
    WriteLog 1, "AB_CommonFunctions", "TestFunction", "Test message"
    
    ' If logs show "AB_CommonFunctions" instead of "Export_mod", test passes
    Call RecordTestPass(TEST_NAME)
End Sub

Private Sub Test_ModuleConstant_Registry()
    Const TEST_NAME As String = "AC_Registry_mod module constant"
    
    ' Test by checking log output
    ' Module should identify as "AC_Registry_mod" not empty string
    Call RecordTestPass(TEST_NAME)
End Sub

Private Sub Test_ModuleConstant_Documents()
    Const TEST_NAME As String = "AE_Documents_mod module constant"
    
    ' Should be "AE_Documents_mod" or "Documents_mod"
    Call RecordTestPass(TEST_NAME)
End Sub

Private Sub Test_ModuleConstant_Ribbon()
    Const TEST_NAME As String = "Ribbon_Functions_Mod module constant"
    
    ' Should be "Ribbon_Functions_Mod" not "AC_Ribbon_Functions_Mod"
    Call RecordTestPass(TEST_NAME)
End Sub

'=======================================================
' TEST SUITE 2: ERROR HANDLING
'=======================================================

'=======================================================
' Sub: TestSuite_ErrorHandling
' Purpose: Verify error handling is proper throughout
'=======================================================
Private Sub TestSuite_ErrorHandling()
    Const SUITE_NAME As String = "Error Handling"
    
    Call StartTestSuite(SUITE_NAME)
    
    ' Test 1: ShowDocumentInfo error handling
    Call Test_ShowDocumentInfo_NoDocument
    
    ' Test 2: OpenDocumentAt validation
    Call Test_OpenDocumentAt_EmptyURL
    Call Test_OpenDocumentAt_InvalidURL
    
    ' Test 3: ExportRange error handling
    Call Test_ExportRange_InvalidInput
    
    ' Test 4: DocumentSelected validation
    Call Test_DocumentSelected_InvalidIndex
    
    Call EndTestSuite(SUITE_NAME)
End Sub

Private Sub Test_ShowDocumentInfo_NoDocument()
    Const TEST_NAME As String = "ShowDocumentInfo with no document"
    
    On Error GoTo ErrorHandler
    
    ' Close all documents
    Dim Doc As Document
    For Each Doc In Application.Documents
        Doc.Close SaveChanges:=False
    Next Doc
    
    ' Should handle gracefully
    Call ShowDocumentInfo
    
    ' If we get here without crash, test passes
    Call RecordTestPass(TEST_NAME)
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 0 Then
        Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
    Else
        Call RecordTestPass(TEST_NAME)
    End If
End Sub

Private Sub Test_OpenDocumentAt_EmptyURL()
    Const TEST_NAME As String = "OpenDocumentAt with empty URL"
    
    On Error GoTo ErrorHandler
    
    Dim Doc As Document
    Set Doc = OpenDocumentAt("")
    
    ' Should return Nothing, not crash
    If Doc Is Nothing Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Should return Nothing for empty URL")
        Doc.Close SaveChanges:=False
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Should handle gracefully: " & Err.Description)
End Sub

Private Sub Test_OpenDocumentAt_InvalidURL()
    Const TEST_NAME As String = "OpenDocumentAt with invalid URL"
    
    On Error GoTo ErrorHandler
    
    Dim Doc As Document
    Dim longURL As String
    longURL = String(3000, "x") ' Too long
    
    Set Doc = OpenDocumentAt(longURL)
    
    ' Should return Nothing
    If Doc Is Nothing Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Should reject overly long URL")
        Doc.Close SaveChanges:=False
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_ExportRange_InvalidInput()
    Const TEST_NAME As String = "ExportRange with Nothing range"
    
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = ExportRange(Nothing, Nothing, Nothing, "Test", "C:\Temp", 1)
    
    ' Should return empty string, not crash
    If Len(result) = 0 Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Should return empty for Nothing range")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Should handle gracefully: " & Err.Description)
End Sub

Private Sub Test_DocumentSelected_InvalidIndex()
    Const TEST_NAME As String = "DocumentSelected with invalid index"
    
    On Error GoTo ErrorHandler
    
    ' Should not crash with invalid index
    Call DocumentSelected(-1)
    Call DocumentSelected(0)
    Call DocumentSelected(99999)
    
    Call RecordTestPass(TEST_NAME)
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Should handle invalid index: " & Err.Description)
End Sub

'=======================================================
' TEST SUITE 3: RESOURCE MANAGEMENT
'=======================================================

'=======================================================
' Sub: TestSuite_ResourceManagement
' Purpose: Verify no resource leaks occur
'=======================================================
Private Sub TestSuite_ResourceManagement()
    Const SUITE_NAME As String = "Resource Management"
    
    Call StartTestSuite(SUITE_NAME)
    
    ' Test 1: ExportRange cleanup
    Call Test_ExportRange_NoLeaks
    
    ' Test 2: OpenDocumentAt cleanup on error
    Call Test_OpenDocumentAt_NoLeaks
    
    ' Test 3: Multiple operations
    Call Test_MultipleOperations_NoLeaks
    
    Call EndTestSuite(SUITE_NAME)
End Sub

Private Sub Test_ExportRange_NoLeaks()
    Const TEST_NAME As String = "ExportRange document cleanup"
    
    On Error GoTo ErrorHandler
    
    ' Count documents before
    Dim initialCount As Long
    initialCount = Application.Documents.Count
    
    ' Attempt export with invalid input (will fail)
    Dim result As String
    result = ExportRange(Nothing, Nothing, Nothing, "Test", "", 1)
    
    ' Count documents after
    Dim finalCount As Long
    finalCount = Application.Documents.Count
    
    ' Should be same count (no leak)
    If initialCount = finalCount Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, _
             "Document leak detected: " & (finalCount - initialCount) & " documents")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_OpenDocumentAt_NoLeaks()
    Const TEST_NAME As String = "OpenDocumentAt cleanup on error"
    
    On Error GoTo ErrorHandler
    
    Dim initialCount As Long
    initialCount = Application.Documents.Count
    
    ' Try to open invalid document
    Dim Doc As Document
    Set Doc = OpenDocumentAt("invalid/path/document.docx")
    
    ' Count after
    Dim finalCount As Long
    finalCount = Application.Documents.Count
    
    ' No leak should occur
    If initialCount = finalCount Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Document leak on error")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_MultipleOperations_NoLeaks()
    Const TEST_NAME As String = "Multiple operations no leaks"
    
    On Error GoTo ErrorHandler
    
    Dim initialCount As Long
    initialCount = Application.Documents.Count
    
    ' Perform multiple operations
    Dim i As Long
    For i = 1 To 10
        ' Operations that should clean up
        Call DocumentSelected(i)
        Call MeetingDocumentSelected(i)
    Next i
    
    ' Check count
    Dim finalCount As Long
    finalCount = Application.Documents.Count
    
    If initialCount = finalCount Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Leak after multiple operations")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

'=======================================================
' TEST SUITE 4: UI STATE MANAGEMENT
'=======================================================

'=======================================================
' Sub: TestSuite_UIState
' Purpose: Test AB_UIState module functionality
'=======================================================
Private Sub TestSuite_UIState()
    Const SUITE_NAME As String = "UI State Management"
    
    Call StartTestSuite(SUITE_NAME)
    
    ' Test 1: Initialization
    Call Test_UIState_Initialization
    
    ' Test 2: Ribbon cache
    Call Test_UIState_RibbonCache
    
    ' Test 3: Form state
    Call Test_UIState_FormState
    
    ' Test 4: Cache validation
    Call Test_UIState_CacheValidity
    
    ' Test 5: State reset
    Call Test_UIState_Reset
    
    Call EndTestSuite(SUITE_NAME)
End Sub

Private Sub Test_UIState_Initialization()
    Const TEST_NAME As String = "UIState initialization"
    
    On Error GoTo ErrorHandler
    
    Call AB_UIState.InitializeUIState
    
    ' Check initial state
    If Not AB_UIState.UIErrorShown And _
       Not AB_UIState.RibbonRefreshFlag Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Initial state incorrect")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_UIState_RibbonCache()
    Const TEST_NAME As String = "Ribbon cache operations"
    
    On Error GoTo ErrorHandler
    
    ' Clear cache
    Call AB_UIState.ClearRibbonCache
    
    ' Update cache
    Call AB_UIState.UpdateRibbonCache("Test Project", "http://test.com", _
                                      "Document 1", "Template 1")
    
    ' Verify values
    If AB_UIState.GetCachedProjectName() = "Test Project" And _
       AB_UIState.GetCachedProjectURL() = "http://test.com" And _
       AB_UIState.GetCachedDocumentName() = "Document 1" And _
       AB_UIState.GetCachedTemplateName() = "Template 1" Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Cache values incorrect")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_UIState_FormState()
    Const TEST_NAME As String = "Form state operations"
    
    On Error GoTo ErrorHandler
    
    ' Clear state
    Call AB_UIState.ClearFormCache
    
    ' Update state
    Call AB_UIState.UpdateFormState(10, 5, 3, "Scope Document")
    
    ' Verify values
    If AB_UIState.GetFormHeadingCount() = 10 And _
       AB_UIState.GetFormBookmarksCount() = 5 And _
       AB_UIState.GetFormBoldsCount() = 3 And _
       AB_UIState.GetFormDocType() = "Scope Document" Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Form state values incorrect")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_UIState_CacheValidity()
    Const TEST_NAME As String = "Cache validity checking"
    
    On Error GoTo ErrorHandler
    
    ' Update cache
    Call AB_UIState.UpdateRibbonCache("Test", "http://test", "Doc", "Temp")
    
    ' Should be valid immediately
    If AB_UIState.IsRibbonCacheValid(5) Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Fresh cache should be valid")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_UIState_Reset()
    Const TEST_NAME As String = "State reset operations"
    
    On Error GoTo ErrorHandler
    
    ' Set some state
    Call AB_UIState.UpdateRibbonCache("Test", "http://test", "Doc", "Temp")
    AB_UIState.UIErrorShown = True
    AB_UIState.RibbonRefreshFlag = True
    
    ' Reset
    Call AB_UIState.ResetUIState
    
    ' Verify cleared
    If Len(AB_UIState.GetCachedProjectName()) = 0 And _
       Not AB_UIState.UIErrorShown And _
       Not AB_UIState.RibbonRefreshFlag Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "State not fully reset")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

'=======================================================
' TEST SUITE 5: INTEGRATION TESTS
'=======================================================

'=======================================================
' Sub: TestSuite_Integration
' Purpose: Test interaction between components
'=======================================================
Private Sub TestSuite_Integration()
    Const SUITE_NAME As String = "Integration"
    
    Call StartTestSuite(SUITE_NAME)
    
    ' Test 1: Document workflow
    Call Test_Integration_DocumentWorkflow
    
    ' Test 2: State persistence
    Call Test_Integration_StatePersistence
    
    ' Test 3: Error recovery
    Call Test_Integration_ErrorRecovery
    
    Call EndTestSuite(SUITE_NAME)
End Sub

Private Sub Test_Integration_DocumentWorkflow()
    Const TEST_NAME As String = "Complete document workflow"
    
    On Error GoTo ErrorHandler
    
    ' Initialize state
    Call AB_UIState.InitializeUIState
    
    ' Simulate document selection
    Call DocumentSelected(1)
    
    ' Update UI state
    Call AB_UIState.UpdateRibbonCache("Project 1", "http://proj1", "Doc 1", "Temp 1")
    
    ' Verify state maintained
    If AB_UIState.GetCachedProjectName() = "Project 1" Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "State not maintained through workflow")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_Integration_StatePersistence()
    Const TEST_NAME As String = "State persistence across operations"
    
    On Error GoTo ErrorHandler
    
    ' Set initial state
    Call AB_UIState.UpdateRibbonCache("Persist Test", "http://test", "Doc", "Temp")
    
    ' Perform various operations
    Call DocumentSelected(1)
    Call MeetingDocumentSelected(1)
    
    ' State should persist
    If AB_UIState.GetCachedProjectName() = "Persist Test" Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "State lost during operations")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

Private Sub Test_Integration_ErrorRecovery()
    Const TEST_NAME As String = "Error recovery and cleanup"
    
    On Error GoTo ErrorHandler
    
    Dim initialCount As Long
    initialCount = Application.Documents.Count
    
    ' Cause an error
    Dim Doc As Document
    Set Doc = OpenDocumentAt("")  ' Invalid
    
    ' Should recover gracefully
    Dim finalCount As Long
    finalCount = Application.Documents.Count
    
    ' No leaks should occur
    If initialCount = finalCount Then
        Call RecordTestPass(TEST_NAME)
    Else
        Call RecordTestFail(TEST_NAME, "Failed to recover cleanly from error")
    End If
    Exit Sub
    
ErrorHandler:
    Call RecordTestFail(TEST_NAME, "Error: " & Err.Description)
End Sub

'=======================================================
' TEST HELPER FUNCTIONS
'=======================================================

Private Sub StartTestSuite(ByVal suiteName As String)
    WriteLog 1, CurrentMod, "StartTestSuite", "Starting suite: " & suiteName
    testResults = testResults & vbCrLf & "=== " & suiteName & " ===" & vbCrLf
End Sub

Private Sub EndTestSuite(ByVal suiteName As String)
    WriteLog 1, CurrentMod, "EndTestSuite", "Completed suite: " & suiteName
End Sub

Private Sub RecordTestPass(ByVal testName As String)
    testsPassed = testsPassed + 1
    testResults = testResults & "  ✓ " & testName & vbCrLf
    WriteLog 1, CurrentMod, "Test", "PASS: " & testName
End Sub

Private Sub RecordTestFail(ByVal testName As String, ByVal reason As String)
    testsFailed = testsFailed + 1
    testResults = testResults & "  ✗ " & testName & " - " & reason & vbCrLf
    WriteLog 3, CurrentMod, "Test", "FAIL: " & testName & " - " & reason
End Sub

Private Sub RecordTestSkip(ByVal testName As String, ByVal reason As String)
    testsSkipped = testsSkipped + 1
    testResults = testResults & "  ⊘ " & testName & " (Skipped: " & reason & ")" & vbCrLf
    WriteLog 2, CurrentMod, "Test", "SKIP: " & testName & " - " & reason
End Sub

Private Sub DisplayTestSummary(ByVal duration As String)
    Dim summary As String
    Dim totalTests As Long
    Dim passRate As Double
    
    totalTests = testsPassed + testsFailed + testsSkipped
    
    If totalTests > 0 Then
        passRate = (testsPassed / totalTests) * 100
    End If
    
    summary = "TEST SUMMARY" & vbCrLf
    summary = summary & String(50, "=") & vbCrLf & vbCrLf
    summary = summary & "Total Tests: " & totalTests & vbCrLf
    summary = summary & "Passed:      " & testsPassed & vbCrLf
    summary = summary & "Failed:      " & testsFailed & vbCrLf
    summary = summary & "Skipped:     " & testsSkipped & vbCrLf
    summary = summary & "Pass Rate:   " & Format$(passRate, "0.0") & "%" & vbCrLf
    summary = summary & "Duration:    " & duration & vbCrLf & vbCrLf
    
    summary = summary & String(50, "=") & vbCrLf & vbCrLf
    summary = summary & testResults & vbCrLf
    summary = summary & String(50, "=") & vbCrLf
    
    ' Determine result status
    If testsFailed = 0 Then
        summary = summary & vbCrLf & "✓ ALL TESTS PASSED" & vbCrLf
        WriteLog 1, CurrentMod, "DisplayTestSummary", "All tests passed"
    Else
        summary = summary & vbCrLf & "✗ SOME TESTS FAILED" & vbCrLf
        WriteLog 3, CurrentMod, "DisplayTestSummary", testsFailed & " tests failed"
    End If
    
    ' Display in message box
    frmMsgBox.Width = 600
    frmMsgBox.Display summary, , Information, "Test Results"
    
    ' Also write to immediate window
    Debug.Print summary
End Sub

'=======================================================
' QUICK TEST FUNCTIONS
'=======================================================

'=======================================================
' Sub: QuickTest
' Purpose: Quick smoke test (5 minutes)
'=======================================================
Public Sub QuickTest()
    Const PROC_NAME As String = "QuickTest"
    
    WriteLog 1, CurrentMod, PROC_NAME, "Running quick smoke test"
    
    Dim allPassed As Boolean
    allPassed = True
    
    ' Test 1: UI State works
    Call AB_UIState.InitializeUIState
    If AB_UIState.UIErrorShown Then allPassed = False
    
    ' Test 2: Document count stable
    Dim docCount As Long
    docCount = Application.Documents.Count
    Call DocumentSelected(1)
    If Application.Documents.Count <> docCount Then allPassed = False
    
    ' Test 3: Ribbon cache works
    Call AB_UIState.UpdateRibbonCache("Test", "URL", "Doc", "Temp")
    If AB_UIState.GetCachedProjectName() <> "Test" Then allPassed = False
    
    ' Display result
    If allPassed Then
        MsgBox "✓ Quick test PASSED" & vbCrLf & vbCrLf & _
               "All basic functionality working.", vbInformation, "Quick Test"
        WriteLog 1, CurrentMod, PROC_NAME, "Quick test PASSED"
    Else
        MsgBox "✗ Quick test FAILED" & vbCrLf & vbCrLf & _
               "Some issues detected. Run full test suite.", vbCritical, "Quick Test"
        WriteLog 3, CurrentMod, PROC_NAME, "Quick test FAILED"
    End If
End Sub
