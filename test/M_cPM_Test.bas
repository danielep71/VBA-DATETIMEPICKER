Attribute VB_Name = "M_cPM_Test"

Option Explicit     'Force explicit declaration of all variables

'
'==============================================================================
'
'                           M_cPM_RegressionTests
'
'==============================================================================
' PURPOSE
'   Full regression test suite for cPerformanceManager and its required
'   companion module M_cPM_TimeWasters
'
' WHY THIS EXISTS
'   A simple smoke test is useful, but it is not sufficient for a component
'   whose behavior depends on:
'
'     - multiple timing backends
'     - session-bound validation rules
'     - strict vs non-strict behavior
'     - diagnostic / reporting helpers
'     - explicit environment cleanup
'     - shared Excel Application-state coordination
'
'   This module therefore provides:
'
'     - isolated test cases
'     - reusable assertion helpers
'     - suite-level counters and summary output
'     - a dedicated worksheet log for durable regression evidence
'     - validation of both success paths and error / fallback paths
'
' ENTRY POINT
'   Run:
'
'     Run_cPerformanceManager_RegressionSuite
'
' DEPENDENCIES
'   - cPerformanceManager
'   - M_cPM_TimeWasters
'   - Excel Application object model
'
' NOTES
'   - Place this code in a STANDARD MODULE
'   - The class name must be exactly: cPerformanceManager
'   - The companion TW manager module must already be imported and compiled
'   - Results are written to a dedicated worksheet log and summarized in the
'     Immediate Window
'   - This suite assumes the current class surface where:
'       * MethodName(2) returns "GetTickCount / GetTickCount64"
'       * ElapsedTime supports formatting an already measured elapsed value
'       * OverheadMeasurement_Text supports an optional Iterations argument
'
' VERSION
'   1.1.0
'
' UPDATED
'   2026-04-18
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const cPM_SHEET_LOG    As String = "REGRESSION_cPM"
    
'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    'Snapshot of Excel Application state used by TW regression tests
    Private Type T_AppState
        ScreenUpdating  As Boolean   'Application.ScreenUpdating
        EnableEvents    As Boolean   'Application.EnableEvents
        DisplayAlerts   As Boolean   'Application.DisplayAlerts
        Calculation     As Long      'Application.Calculation
        Cursor          As Long      'Application.Cursor
    End Type

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private m_TotalCases            As Long      'Number of executed test cases
    Private m_TotalAssertions       As Long      'Number of executed assertions
    Private m_TotalFailures         As Long      'Number of failed assertions
    Private m_CurrentCaseName       As String    'Current case name

    Private m_CaseAssertions_Begin  As Long      'Assertion count at case start
    Private m_CaseFailures_Begin    As Long      'Failure count at case start

    Private m_RunTimestamp          As String    'Timestamp shared by the whole run
    Private m_SummaryNextRow        As Long      'Next writable row in the case-summary section
    Private m_DetailNextRow         As Long      'Next writable row in the failure-detail section

Public Sub Run_cPerformanceManager_RegressionSuite()
'
'==============================================================================
'                  RUN CPERFORMANCEMANAGER REGRESSION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the full regression suite for cPerformanceManager
'
' WHY THIS EXISTS
'   This is the single public entry point for the regression module. It resets
'   suite state, prepares the dedicated regression worksheet, runs every test
'   case in a deterministic order, and then prints / writes the final summary
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Forces a clean TW baseline before the suite begins
'   - Resets suite counters and run-level state
'   - Builds or rebuilds the dedicated regression worksheet
'   - Initializes the worksheet log and suite header
'   - Executes all regression cases in a deterministic order
'   - Updates the Excel status bar with live suite progress
'   - Forces a clean TW baseline again at the end of the suite
'   - Prints the suite footer when the logging infrastructure is ready
'   - Restores Application state before re-raising any unexpected runner error
'
' ERROR POLICY
'   Individual cases handle their own unexpected errors and record failures
'
'   This runner captures any unexpected runner-level error, performs centralized
'   cleanup, and then re-raises the original runner-level error
'
' DEPENDENCIES
'   - PM_TW_EndAllSessions
'   - Suite_ResetCounters
'   - DEMO_Begin_FastMode
'   - DEMO_End_FastMode
'   - DEMO_Build_DemoTemplate
'   - Suite_InitLogSheet
'   - Suite_PrintHeader
'   - Suite_PrintFooter
'   - Demo_SB_SetProgress
'   - All private Test_* procedures in this module
'
' NOTES
'   The suite is intentionally ordered from simpler / core behaviors toward
'   shared-state and cleanup behaviors
'
'   TotalSteps must be kept aligned with the number of executed regression cases
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                      As Workbook             'Target workbook
    Dim WS_Test                 As Worksheet            'Dedicated regression worksheet

    Dim FastModeState           As tDemoFastModeState   'Saved Application-state snapshot
    Dim FastModeOn              As Boolean              'TRUE when fast mode was entered
    Dim FooterCanPrint          As Boolean              'TRUE when footer/log infrastructure is ready
    Dim SuppressCleanupErrors   As Boolean              'TRUE when cleanup must not mask an earlier error

    Dim SavedErrNum             As Long                 'Captured error number
    Dim SavedErrSrc             As String               'Captured error source
    Dim SavedErrDesc            As String               'Captured error description

    Dim CurrentStep             As Long                 'Current regression-step counter

    Const TotalSteps            As Long = 40            'Total number of executed regression cases

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo Clean_Fail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Force the shared TW manager to a clean baseline before the suite begins
        PM_TW_EndAllSessions
    'Reset suite-level counters and current-case tracking
        Suite_ResetCounters
    'Resolve the workbook that contains the regression module
        Set WB = ThisWorkbook
    'Capture and apply fast-mode Application settings
        DEMO_Begin_FastMode FastModeState
        FastModeOn = True
    'Show a wait cursor while the regression environment is being prepared
        Application.Cursor = xlWait
    'Initialize status-bar progress
        CurrentStep = 0
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Preparing regression environment"

'------------------------------------------------------------------------------
' PREPARE REGRESSION SHEET
'------------------------------------------------------------------------------
    'Build or rebuild the dedicated regression worksheet
        DEMO_Build_DemoTemplate _
            cPM_SHEET_LOG, _
            "CLASS PERFORMANCE MANAGER", _
            "REGRESSION TESTS", , 6, , , , "C:O", , , , , , , , , , 47
    'Resolve the prepared regression worksheet
        Set WS_Test = WB.Worksheets(cPM_SHEET_LOG)
    'Initialize the dedicated worksheet log
        Suite_InitLogSheet
    'Print the suite header
        Suite_PrintHeader
    'Mark that the footer can safely be printed during centralized cleanup
        FooterCanPrint = True

'------------------------------------------------------------------------------
' RUN CORE / TIMING REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate constructor/default state
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Default constructor state"
        Test_DefaultState

    'Validate valid MethodName mappings
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "MethodName valid indices"
        Test_MethodName_ValidIndices

    'Validate invalid MethodName behavior
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "MethodName invalid indices"
        Test_MethodName_InvalidIndices

    'Validate StartTimer session-state transitions for all methods
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "StartTimer session state"
        Test_StartTimer_SetsSessionState_AllMethods

    'Validate numeric elapsed-time reads for all methods
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ElapsedSeconds across methods"
        Test_ElapsedSeconds_AllMethods

    'Validate formatted elapsed-time reads for all methods
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ElapsedTime across methods"
        Test_ElapsedTime_AllMethods

    'Validate formatting of an already measured elapsed-seconds value
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ElapsedTime format existing seconds"
        Test_ElapsedTime_FormatExistingSeconds

    'Validate aligned-start timing for all methods
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Aligned-start timing"
        Test_AlignedStart_AllMethods

    'Validate raw accessor behavior after a QPC measurement
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Raw/cached accessors after QPC"
        Test_Accessors_QPC

'------------------------------------------------------------------------------
' RUN VALIDATION / FALLBACK REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate strict-mode behavior when elapsed time is requested before StartTimer
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "StrictMode: elapsed before StartTimer"
        Test_StrictMode_ElapsedBeforeStart

    'Validate strict-mode behavior for explicit method mismatch
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "StrictMode: method mismatch"
        Test_StrictMode_MethodMismatch

    'Validate strict-mode behavior for invalid start-method input
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "StrictMode: invalid StartTimer method"
        Test_StrictMode_InvalidStartMethod

    'Validate non-strict fallback for invalid start-method input
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "NonStrictMode: invalid StartTimer fallback"
        Test_NonStrictMode_InvalidStartFallback

    'Validate non-strict fallback for explicit elapsed-method mismatch
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "NonStrictMode: elapsed-method mismatch fallback"
        Test_NonStrictMode_MethodMismatchFallback

'------------------------------------------------------------------------------
' RUN DIAGNOSTIC / OVERHEAD REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate numeric overhead measurement helpers
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "OverheadMeasurement_Seconds"
        Test_OverheadMeasurement_Seconds

    'Validate formatted overhead measurement helpers
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "OverheadMeasurement_Text"
        Test_OverheadMeasurement_Text

    'Validate diagnostic / informational properties
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Diagnostics and informational properties"
        Test_Diagnostics_Properties

'------------------------------------------------------------------------------
' RUN PAUSE REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate Pause method 1
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause method 1"
        Test_Pause_Method1

    'Validate Pause method 2
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause method 2"
        Test_Pause_Method2

    'Validate Pause method 3
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause method 3"
        Test_Pause_Method3

    'Validate Pause method 4
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause method 4"
        Test_Pause_Method4

'------------------------------------------------------------------------------
' RUN TW / CLEANUP REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate blank-key behavior in the shared TW manager
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "TW blank-key validation"
        Test_TW_BlankKeyValidation

    'Validate single-instance TW lifecycle
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "TW single-instance lifecycle"
        Test_TW_SingleInstance

    'Validate overlapping multi-instance TW lifecycle
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "TW overlapping multi-instance lifecycle"
        Test_TW_OverlappingInstances

    'Validate ResetEnvironment idempotence and cleanup behavior
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ResetEnvironment idempotence"
        Test_ResetEnvironment_Idempotent

'------------------------------------------------------------------------------
' RUN CHECKPOINT / REPORTING REGRESSION CASES
'------------------------------------------------------------------------------
    'Validate that Checkpoint raises before StartTimer
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Checkpoint before StartTimer"
        Test_Checkpoint_BeforeStart

    'Validate SetRunLabel before first checkpoint
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "SetRunLabel before first checkpoint"
        Test_SetRunLabel_BeforeFirstCheckpoint

    'Validate SetRunLabel rejection after first checkpoint
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "SetRunLabel after first checkpoint"
        Test_SetRunLabel_AfterFirstCheckpoint

    'Validate default checkpoint naming when the supplied name is blank
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Checkpoint default name when blank"
        Test_Checkpoint_DefaultName_WhenBlank

    'Validate checkpoint count and ReportAsArray structure/content
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "CheckpointCount and ReportAsArray"
        Test_CheckpointCount_And_ReportArray

    'Validate ReportAsText behavior when no checkpoints exist
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ReportAsText when empty"
        Test_ReportAsText_Empty

    'Validate ReportAsText behavior after checkpoint capture
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ReportAsText with checkpoints"
        Test_ReportAsText_WithCheckpoints

    'Validate ClearCheckpoints behavior
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ClearCheckpoints"
        Test_ClearCheckpoints

    'Validate that StartTimer resets checkpoint/report state
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "StartTimer clears checkpoint state"
        Test_StartTimer_ClearsCheckpointState

    'Validate direct empty-array export when no checkpoints exist
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ReportAsArray direct empty export"
        Test_ReportAsArray_Empty_Direct

    'Validate delta/cumulative checkpoint timing semantics
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Checkpoint delta/cumulative semantics"
        Test_Checkpoint_DeltaAndCumulativeSemantics

    'Validate reuse of the same session after ClearCheckpoints
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "ClearCheckpoints then reuse session"
        Test_ClearCheckpoints_ThenReuseSession

    'Validate SetRunLabel after ClearCheckpoints
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "SetRunLabel after ClearCheckpoints"
        Test_SetRunLabel_AfterClearCheckpoints

    'Validate idempotent TW_Turn_ON behavior when no TW session is active
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "TW_Turn_ON when inactive"
        Test_TW_Turn_ON_WhenInactive

    'Validate Pause no-op boundary behavior
        CurrentStep = CurrentStep + 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause no-op boundary behavior"
        Test_Pause_Boundary_NoOpBehavior

Clean_Exit:
'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Determine whether cleanup must be best-effort to preserve an earlier error
        SuppressCleanupErrors = (SavedErrNum <> 0)
    'Do not let cleanup failures mask the original runner-level error
        If SuppressCleanupErrors Then
            On Error Resume Next
        End If

'------------------------------------------------------------------------------
' RESTORE SHARED / APPLICATION STATE
'------------------------------------------------------------------------------
    'Force the shared TW manager to a clean baseline after the suite ends
        PM_TW_EndAllSessions
    'Print the final suite summary when the logging infrastructure is available
        If FooterCanPrint Then
            Suite_PrintFooter
        End If
    'Clear the custom status-bar message
        Application.StatusBar = False
    'Restore the normal cursor
        Application.Cursor = xlDefault
    'Restore the original Excel Application state only when fast mode was entered
        If FastModeOn Then
            DEMO_End_FastMode FastModeState
        End If

'------------------------------------------------------------------------------
' APPLY FINAL PRESENTATION STEP
'------------------------------------------------------------------------------
    'Bring the regression worksheet to the foreground on a best-effort basis
        On Error Resume Next
        If Not WS_Test Is Nothing Then
            WS_Test.Activate
            WS_Test.Range("A1").Select
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR MODE
'------------------------------------------------------------------------------
    'Restore the appropriate error mode after cleanup / presentation work
        If SuppressCleanupErrors Then
            On Error GoTo 0
        Else
            On Error GoTo Clean_Fail
        End If

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original runner-level error after cleanup when needed
        If SavedErrNum <> 0 Then
            Err.Raise SavedErrNum, SavedErrSrc, SavedErrDesc
        End If

    Exit Sub

Clean_Fail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before centralized cleanup
        SavedErrNum = Err.Number
        SavedErrSrc = Err.Source
        SavedErrDesc = Err.Description
    'Continue through the centralized cleanup path
        Resume Clean_Exit

End Sub
'
'==============================================================================
'
'                           SUITE / ASSERT HELPERS
'
'==============================================================================

Private Sub Suite_ResetCounters()
'
'==============================================================================
'                           SUITE RESET COUNTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Resets suite-level counters and run-level tracking state
'
' WHY THIS EXISTS
'   Each regression run should start from a deterministic baseline so that case
'   counts, assertion counts, failure counts, and worksheet-log pointers are
'   not contaminated by prior runs
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' RESET SUITE COUNTERS
'------------------------------------------------------------------------------
    'Reset the executed-case counter
        m_TotalCases = 0
    'Reset the assertion counter
        m_TotalAssertions = 0
    'Reset the failure counter
        m_TotalFailures = 0

'------------------------------------------------------------------------------
' RESET CURRENT-CASE TRACKING
'------------------------------------------------------------------------------
    'Clear the current-case name
        m_CurrentCaseName = vbNullString
    'Reset the per-case assertion baseline
        m_CaseAssertions_Begin = 0
    'Reset the per-case failure baseline
        m_CaseFailures_Begin = 0

'------------------------------------------------------------------------------
' RESET RUN-LEVEL LOG STATE
'------------------------------------------------------------------------------
    'Clear the shared run timestamp
        m_RunTimestamp = vbNullString
    'Reset the next writable row in the case-summary section
        m_SummaryNextRow = 0
    'Reset the next writable row in the failure-detail section
        m_DetailNextRow = 0

End Sub


Private Sub Suite_InitLogSheet()
'
'==============================================================================
'                           SUITE INIT LOG SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the dedicated worksheet log used by the regression suite
'
' WHY THIS EXISTS
'   Debug.Print alone is not a reliable medium for a large suite because older
'   lines become difficult to inspect. A dedicated worksheet preserves the run
'   output in a structured, reviewable format
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Resolves the dedicated regression-log worksheet
'   - Stamps one run timestamp for the whole suite
'   - Adjusts the visible worksheet layout required by the regression log
'   - Rebuilds the case-summary section
'   - Rebuilds the failure-detail section
'   - Applies deterministic widths, alignments, headers, and borders
'   - Initializes row pointers for case summaries and failure details
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Suite_GetOrCreateLogSheet
'   - DEMO_Write_BandHeader
'   - DEMO_Set_RangeBorder
'
' NOTES
'   This routine assumes the regression sheet has just been rebuilt by the
'   suite runner before initialization begins
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet
    Dim SummaryHeaders      As Variant      'Case-summary column headers
    Dim DetailHeaders       As Variant      'Failure-detail column headers
    Dim i                   As Long         'Header loop index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Stamp the current run timestamp once for the whole suite
        m_RunTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()

'------------------------------------------------------------------------------
' DEFINE HEADER SETS
'------------------------------------------------------------------------------
    'Define the case-summary headers
        SummaryHeaders = Array( _
            "RunTimestamp", _
            "CaseNo", _
            "CaseName", _
            "Assertions", _
            "Failures", _
            "Result")
    'Define the failure-detail headers
        DetailHeaders = Array( _
            "RunTimestamp", _
            "CaseNo", _
            "CaseName", _
            "FailureType", _
            "Message", _
            "ErrNumber", _
            "ErrDescription")

'------------------------------------------------------------------------------
' APPLY COLUMN LAYOUT
'------------------------------------------------------------------------------
    'Apply the width layout for the case-summary section
        WS.Columns("C").ColumnWidth = 15
        WS.Columns("D").ColumnWidth = 10
        WS.Columns("E").ColumnWidth = 40
        WS.Columns("F:H").ColumnWidth = 10
    'Apply the width layout for the failure-detail section
        WS.Columns("I").ColumnWidth = 15
        WS.Columns("J").ColumnWidth = 10
        WS.Columns("K").ColumnWidth = 40
        WS.Columns("L").ColumnWidth = 10
        WS.Columns("M").ColumnWidth = 40
        WS.Columns("N").ColumnWidth = 10
        WS.Columns("O").ColumnWidth = 30

'------------------------------------------------------------------------------
' APPLY BASE ALIGNMENT
'------------------------------------------------------------------------------
    'Apply center alignment to the structured log block by default
        With WS.Range("C6:O50")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Size = 9
        End With
    'Restore left alignment for text-heavy columns
        WS.Columns("E").HorizontalAlignment = xlLeft
        WS.Columns("K").HorizontalAlignment = xlLeft
        WS.Columns("M").HorizontalAlignment = xlLeft
        WS.Columns("O").HorizontalAlignment = xlLeft

'------------------------------------------------------------------------------
' BUILD CASE SUMMARY AREA
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' WRITE SECTION HEADER
    '--------------------------------------------------------------------------
        'Write the case-summary section header band
            DEMO_Write_BandHeader WS.Range("C4:H4"), "CASE SUMMARY"

    '--------------------------------------------------------------------------
    ' WRITE COLUMN CAPTIONS
    '--------------------------------------------------------------------------
        'Write the case-summary headers
            For i = LBound(SummaryHeaders) To UBound(SummaryHeaders)
                WS.Cells(5, "C").Offset(0, i).Value = SummaryHeaders(i)
            Next i

    '--------------------------------------------------------------------------
    ' FORMAT HEADER ROW
    '--------------------------------------------------------------------------
        'Apply standard header styling to the case-summary header row
            With WS.Range("C5:H5")
                .Interior.Color = COLOR_SUBHEADER
                .Font.Bold = True
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the visible case-summary block
            DEMO_Set_RangeBorder WS.Range("C4:H45")

'------------------------------------------------------------------------------
' BUILD FAILURE DETAILS AREA
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' WRITE SECTION HEADER
    '--------------------------------------------------------------------------
        'Write the failure-details section header band
            DEMO_Write_BandHeader WS.Range("I4:O4"), "FAILURE DETAILS"

    '--------------------------------------------------------------------------
    ' WRITE COLUMN CAPTIONS
    '--------------------------------------------------------------------------
        'Write the failure-detail headers
            For i = LBound(DetailHeaders) To UBound(DetailHeaders)
                WS.Cells(5, "I").Offset(0, i).Value = DetailHeaders(i)
            Next i

    '--------------------------------------------------------------------------
    ' FORMAT HEADER ROW
    '--------------------------------------------------------------------------
        'Apply standard header styling to the failure-detail header row
            With WS.Range("I5:O5")
                .Interior.Color = COLOR_SUBHEADER
                .Font.Bold = True
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the visible failure-detail block
            DEMO_Set_RangeBorder WS.Range("I4:O45")

'------------------------------------------------------------------------------
' INITIALIZE ROW POINTERS
'------------------------------------------------------------------------------
    'Initialize the first writable summary row
        m_SummaryNextRow = 6
    'Initialize the first writable detail row
        m_DetailNextRow = 6

End Sub
Private Function Suite_GetOrCreateLogSheet() As Worksheet
'
'==============================================================================
'                        SUITE GET OR CREATE LOG SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the dedicated regression-log worksheet, creating it if missing
'
' WHY THIS EXISTS
'   The regression module should remain self-contained and able to resolve its
'   own output worksheet without depending on external sheet-builder logic
'
' INPUTS
'   None
'
' RETURNS
'   Worksheet
'     Existing or newly created regression-log worksheet
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - ThisWorkbook
'   - cPM_SHEET_LOG
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook     'Target workbook
    Dim WS                  As Worksheet    'Worksheet iterator / result
    Dim SheetNameText       As String       'Requested regression sheet name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the workbook that contains the regression module
        Set WB = ThisWorkbook
    'Resolve the requested regression sheet name
        SheetNameText = Trim$(cPM_SHEET_LOG)

'------------------------------------------------------------------------------
' VALIDATE SHEET NAME
'------------------------------------------------------------------------------
    'Reject a blank regression-sheet name
        If Len(SheetNameText) = 0 Then
            Err.Raise vbObjectError + 2400, _
                      "M_cPM_RegressionTests.Suite_GetOrCreateLogSheet", _
                      "Regression sheet name cannot be blank."
        End If
    'Reject names longer than Excel's worksheet-name limit
        If Len(SheetNameText) > 31 Then
            Err.Raise vbObjectError + 2401, _
                      "M_cPM_RegressionTests.Suite_GetOrCreateLogSheet", _
                      "Regression sheet name cannot exceed 31 characters."
        End If
    'Reject worksheet names containing invalid Excel characters
        If InStr(1, SheetNameText, ":", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "\", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "/", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "?", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "*", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "[", vbBinaryCompare) > 0 _
        Or InStr(1, SheetNameText, "]", vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2402, _
                      "M_cPM_RegressionTests.Suite_GetOrCreateLogSheet", _
                      "Regression sheet name contains one or more invalid Excel worksheet-name characters."
        End If

'------------------------------------------------------------------------------
' SEARCH EXISTING SHEETS
'------------------------------------------------------------------------------
    'Search for an existing worksheet with the requested regression-sheet name
        For Each WS In WB.Worksheets
            If StrComp(WS.Name, SheetNameText, vbTextCompare) = 0 Then
                Set Suite_GetOrCreateLogSheet = WS
                Exit Function
            End If
        Next WS

'------------------------------------------------------------------------------
' CREATE SHEET
'------------------------------------------------------------------------------
    'Create a new worksheet because the requested one does not yet exist
        Set WS = WB.Worksheets.Add(After:=WB.Worksheets(WB.Worksheets.Count))
    'Assign the requested regression-sheet name to the new worksheet
        WS.Name = SheetNameText

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the existing or newly created regression worksheet
        Set Suite_GetOrCreateLogSheet = WS

End Function

Private Sub Suite_PrintHeader()
'
'==============================================================================
'                             SUITE PRINT HEADER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the suite-level header to the Immediate Window
'
' WHY THIS EXISTS
'   Makes a regression run easy to identify and separates it from prior
'   Immediate Window output
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT HEADER
'------------------------------------------------------------------------------
    'Print a blank line before the suite header block
        Debug.Print vbNullString
    'Print the opening delimiter
        Debug.Print String$(100, "=")
    'Print the suite title
        Debug.Print "REGRESSION SUITE START : cPerformanceManager"
    'Print the suite timestamp
        Debug.Print "Timestamp              : " & m_RunTimestamp
    'Print the worksheet-log target
        Debug.Print "Worksheet log          : " & cPM_SHEET_LOG
    'Print the closing delimiter
        Debug.Print String$(100, "=")

End Sub


Private Sub Suite_PrintFooter()
'
'==============================================================================
'                             SUITE PRINT FOOTER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints the suite-level summary to the Immediate Window
'
' WHY THIS EXISTS
'   A regression run should end with a concise summary of execution volume and
'   overall pass / fail status
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SUMMARY
'------------------------------------------------------------------------------
    'Print a delimiter line before the summary block
        Debug.Print String$(100, "-")
    'Print the total number of executed cases
        Debug.Print "Total cases           : " & m_TotalCases
    'Print the total number of executed assertions
        Debug.Print "Total assertions      : " & m_TotalAssertions
    'Print the total number of recorded failures
        Debug.Print "Total failures        : " & m_TotalFailures
    'Print the overall suite status
        If m_TotalFailures = 0 Then
            Debug.Print "OVERALL RESULT        : PASS"
        Else
            Debug.Print "OVERALL RESULT        : FAIL"
        End If
    'Print the worksheet-log target
        Debug.Print "Worksheet log         : " & cPM_SHEET_LOG
    'Print the closing delimiter
        Debug.Print String$(100, "=")

End Sub

Private Sub Test_Assert_ApproxDouble( _
    ByVal Expected As Double, _
    ByVal Actual As Double, _
    ByVal Tolerance As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                        TEST ASSERT APPROX DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for approximate Double equality
'
' WHY THIS EXISTS
'   Floating-point values are often better compared within a tolerance than by
'   exact equality
'
' INPUTS
'   Expected
'     Expected Double value
'
'   Actual
'     Actual Double value
'
'   Tolerance
'     Maximum allowed absolute difference
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT APPROXIMATE EQUALITY
'------------------------------------------------------------------------------
    'Assert that the absolute difference is within the requested tolerance
        Test_Assert_True (Abs(Actual - Expected) <= Tolerance), _
                          MessageText & " | expected=" & Format$(Expected, "0.000000000") & _
                          " actual=" & Format$(Actual, "0.000000000") & _
                          " tol=" & Format$(Tolerance, "0.000000000")

End Sub


Private Sub Test_Assert_InRangeDouble( _
    ByVal LowerBound As Double, _
    ByVal UpperBound As Double, _
    ByVal Actual As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                       TEST ASSERT INRANGE DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that a Double lies within a closed interval
'
' WHY THIS EXISTS
'   Pause and elapsed-time checks are often more meaningful as range checks than
'   as exact-value checks
'
' INPUTS
'   LowerBound
'     Minimum acceptable value
'
'   UpperBound
'     Maximum acceptable value
'
'   Actual
'     Actual measured value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT RANGE
'------------------------------------------------------------------------------
    'Assert that the actual value lies within the closed interval
        Test_Assert_True ((Actual >= LowerBound) And (Actual <= UpperBound)), _
                          MessageText & " | range=[" & Format$(LowerBound, "0.000000000") & _
                          ", " & Format$(UpperBound, "0.000000000") & _
                          "] actual=" & Format$(Actual, "0.000000000")

End Sub

Private Sub Test_Assert_NonNegativeDouble( _
    ByVal Actual As Double, _
    ByVal MessageText As String)
'
'==============================================================================
'                     TEST ASSERT NONNEGATIVE DOUBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that a Double is nonnegative
'
' WHY THIS EXISTS
'   Many timing and diagnostic values should never be negative
'
' INPUTS
'   Actual
'     Actual Double value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT NONNEGATIVE
'------------------------------------------------------------------------------
    'Assert that the actual value is nonnegative
        Test_Assert_True (Actual >= 0#), _
                          MessageText & " | actual=" & Format$(Actual, "0.000000000")

End Sub


Private Sub Test_Assert_True( _
    ByVal Condition As Boolean, _
    ByVal MessageText As String)
'
'==============================================================================
'                             TEST ASSERT TRUE
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion based on a Boolean condition
'
' WHY THIS EXISTS
'   Most regression checks reduce naturally to a Boolean predicate. This helper
'   centralizes assertion counting and failure logging
'
' INPUTS
'   Condition
'     Boolean result to evaluate
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' UPDATE ASSERTION COUNT
'------------------------------------------------------------------------------
    'Increment the total assertion count
        m_TotalAssertions = m_TotalAssertions + 1

'------------------------------------------------------------------------------
' RECORD FAILURE ONLY
'------------------------------------------------------------------------------
    'Exit quietly for passing assertions
        If Condition Then Exit Sub
    'Count one failing assertion
        m_TotalFailures = m_TotalFailures + 1
    'Write the failure detail to the worksheet log
        LogFailureDetail "ASSERT", MessageText

End Sub

Private Sub Test_Assert_EqualLong( _
    ByVal Expected As Long, _
    ByVal Actual As Long, _
    ByVal MessageText As String)
'
'==============================================================================
'                          TEST ASSERT EQUAL LONG
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for Long equality
'
' WHY THIS EXISTS
'   Many regression checks compare IDs, counters, and other Long-valued states
'
' INPUTS
'   Expected
'     Expected Long value
'
'   Actual
'     Actual Long value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual Long equals the expected Long
        Test_Assert_True (Actual = Expected), _
                          MessageText & " | expected=" & CStr(Expected) & _
                          " actual=" & CStr(Actual)

End Sub

Private Sub Test_Assert_EqualBoolean( _
    ByVal Expected As Boolean, _
    ByVal Actual As Boolean, _
    ByVal MessageText As String)
'
'==============================================================================
'                        TEST ASSERT EQUAL BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for Boolean equality
'
' WHY THIS EXISTS
'   Many state and lifecycle checks are Boolean in nature
'
' INPUTS
'   Expected
'     Expected Boolean value
'
'   Actual
'     Actual Boolean value
'
'   MessageText
'     Human-readable assertion label
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual Boolean equals the expected Boolean
        Test_Assert_True (Actual = Expected), _
                          MessageText & " | expected=" & CStr(Expected) & _
                          " actual=" & CStr(Actual)

End Sub

Private Sub Test_Assert_EqualString( _
    ByVal Expected As String, _
    ByVal Actual As String, _
    ByVal MessageText As String)
'
'==============================================================================
'                         TEST ASSERT EQUAL STRING
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion for String equality
'
' WHY THIS EXISTS
'   Method labels and text diagnostics often need exact text comparison
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Q                   As String        'Double-quote character

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Prepare a reusable double-quote character
        Q = Chr$(34)

'------------------------------------------------------------------------------
' ASSERT EQUALITY
'------------------------------------------------------------------------------
    'Assert that the actual String equals the expected String
        Test_Assert_True (Actual = Expected), _
                          MessageText & " | expected=" & Q & Expected & Q & _
                          " actual=" & Q & Actual & Q

End Sub

Private Sub Test_Assert_ContainsString( _
    ByVal SourceText As String, _
    ByVal SubText As String, _
    ByVal MessageText As String)
'
'==============================================================================
'                       TEST ASSERT CONTAINS STRING
'------------------------------------------------------------------------------
' PURPOSE
'   Records a pass/fail assertion that one String contains another
'
' WHY THIS EXISTS
'   Many formatted outputs are better validated by required markers than by
'   exact whole-string equality
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Q                   As String        'Double-quote character

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Prepare a reusable double-quote character
        Q = Chr$(34)

'------------------------------------------------------------------------------
' ASSERT CONTAINS
'------------------------------------------------------------------------------
    'Assert that the source text contains the required substring
        Test_Assert_True (InStr(1, SourceText, SubText, vbTextCompare) > 0), _
                          MessageText & " | required=" & Q & SubText & Q

End Sub

Private Sub Case_Begin( _
    ByVal CaseName As String)
'
'==============================================================================
'                                CASE BEGIN
'------------------------------------------------------------------------------
' PURPOSE
'   Marks the start of one regression case
'
' WHY THIS EXISTS
'   Centralizes case counting, current-case naming, and per-case counter
'   baselining
'
' INPUTS
'   CaseName
'     Human-readable regression case name
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' UPDATE CASE COUNTER
'------------------------------------------------------------------------------
    'Increment the total number of executed cases
        m_TotalCases = m_TotalCases + 1

'------------------------------------------------------------------------------
' UPDATE CURRENT-CASE STATE
'------------------------------------------------------------------------------
    'Store the current case name in trimmed form
        m_CurrentCaseName = Trim$(CaseName)

'------------------------------------------------------------------------------
' CAPTURE PER-CASE BASELINES
'------------------------------------------------------------------------------
    'Capture the assertion counter at case start
        m_CaseAssertions_Begin = m_TotalAssertions
    'Capture the failure counter at case start
        m_CaseFailures_Begin = m_TotalFailures

End Sub

Private Sub Case_Finalize()
'
'==============================================================================
'                              CASE FINALIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Finalizes one regression case by writing a compact summary to the worksheet
'   log and to the Immediate Window
'
' WHY THIS EXISTS
'   Keeps the Immediate Window readable while preserving structured case output
'   on the dedicated regression sheet
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet
    Dim SummaryRow          As Long         'Current summary-row target
    Dim CaseAssertions      As Long         'Assertions executed in this case
    Dim CaseFailures        As Long         'Failures recorded in this case
    Dim ResultText          As String       'Per-case result text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()
    'Resolve the current summary-row target
        SummaryRow = m_SummaryNextRow
    'Compute assertions executed in this case
        CaseAssertions = m_TotalAssertions - m_CaseAssertions_Begin
    'Compute failures recorded in this case
        CaseFailures = m_TotalFailures - m_CaseFailures_Begin
    'Resolve the per-case result text
        If CaseFailures = 0 Then
            ResultText = "PASS"
        Else
            ResultText = "FAIL"
        End If

'------------------------------------------------------------------------------
' WRITE WORKSHEET SUMMARY ROW
'------------------------------------------------------------------------------
    'Write the compact worksheet summary row
        With WS
            .Cells(SummaryRow, "C").Value = m_RunTimestamp
            .Cells(SummaryRow, "D").Value = m_TotalCases
            .Cells(SummaryRow, "E").Value = m_CurrentCaseName
            .Cells(SummaryRow, "F").Value = CaseAssertions
            .Cells(SummaryRow, "G").Value = CaseFailures
            .Cells(SummaryRow, "H").Value = ResultText
        End With

'------------------------------------------------------------------------------
' PRINT COMPACT CASE SUMMARY
'------------------------------------------------------------------------------
    'Print one compact case summary line
        If CaseFailures = 0 Then
            Debug.Print "CASE " & Format$(m_TotalCases, "00") & " PASS - " & _
                        m_CurrentCaseName & " | assertions=" & CStr(CaseAssertions)
        Else
            Debug.Print "CASE " & Format$(m_TotalCases, "00") & " FAIL - " & _
                        m_CurrentCaseName & " | assertions=" & CStr(CaseAssertions) & _
                        " | failures=" & CStr(CaseFailures)
        End If

'------------------------------------------------------------------------------
' ADVANCE POINTER
'------------------------------------------------------------------------------
    'Advance the summary-row pointer
        m_SummaryNextRow = SummaryRow + 1

End Sub

Private Sub LogFailureDetail( _
    ByVal FailureType As String, _
    ByVal MessageText As String, _
    Optional ByVal ErrNumberIn As Variant, _
    Optional ByVal ErrDescriptionIn As Variant)
'
'==============================================================================
'                           LOG FAILURE DETAIL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one detailed failure row to the worksheet log
'
' WHY THIS EXISTS
'   Failure details are easier to review and filter on a worksheet than in the
'   Immediate Window alone
'
' INPUTS
'   FailureType
'     Short failure category such as ASSERT or ERROR
'
'   MessageText
'     Human-readable failure message
'
'   ErrNumberIn (optional)
'     Optional error number for unexpected-error cases
'
'   ErrDescriptionIn (optional)
'     Optional error description for unexpected-error cases
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS                  As Worksheet    'Regression-log worksheet
    Dim DetailRow           As Long         'Current detail-row target

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the regression-log worksheet
        Set WS = Suite_GetOrCreateLogSheet()
    'Resolve the current detail-row target
        DetailRow = m_DetailNextRow

'------------------------------------------------------------------------------
' WRITE DETAIL ROW
'------------------------------------------------------------------------------
    'Write the detailed failure row to the worksheet log
        With WS
            .Cells(DetailRow, "I").Value = m_RunTimestamp
            .Cells(DetailRow, "J").Value = m_TotalCases
            .Cells(DetailRow, "K").Value = m_CurrentCaseName
            .Cells(DetailRow, "L").Value = Trim$(FailureType)
            .Cells(DetailRow, "M").Value = MessageText

            'Write the optional error number when supplied
                If Not IsMissing(ErrNumberIn) Then
                    If Not IsEmpty(ErrNumberIn) Then
                        .Cells(DetailRow, "N").Value = ErrNumberIn
                    End If
                End If

            'Write the optional error description when supplied
                If Not IsMissing(ErrDescriptionIn) Then
                    If Not IsEmpty(ErrDescriptionIn) Then
                        .Cells(DetailRow, "O").Value = ErrDescriptionIn
                    End If
                End If
        End With

'------------------------------------------------------------------------------
' ADVANCE POINTER
'------------------------------------------------------------------------------
    'Advance the detail-row pointer
        m_DetailNextRow = DetailRow + 1

End Sub

Private Sub RecordUnexpectedError( _
    ByVal ProcName As String)
'
'==============================================================================
'                          RECORD UNEXPECTED ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Records one unexpected test-case error as a suite failure
'
' WHY THIS EXISTS
'   A regression case may fail before reaching some or all of its explicit
'   assertions. This helper converts that event into one recorded suite failure
'   and stores diagnostic detail in the worksheet log
'
' INPUTS
'   ProcName
'     Name of the regression procedure that encountered the error
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Count one synthetic assertion for the unexpected error event
        m_TotalAssertions = m_TotalAssertions + 1
    'Count one failure for the unexpected error event
        m_TotalFailures = m_TotalFailures + 1

'------------------------------------------------------------------------------
' STORE DIAGNOSTIC
'------------------------------------------------------------------------------
    'Write the unexpected error detail to the worksheet log
        If Len(Trim$(ProcName)) = 0 Then
            LogFailureDetail "ERROR", _
                             "Unexpected error in (unknown procedure)", _
                             Err.Number, _
                             Err.Description
        Else
            LogFailureDetail "ERROR", _
                             "Unexpected error in " & Trim$(ProcName), _
                             Err.Number, _
                             Err.Description
        End If
        
End Sub

Private Sub CaptureAppState( _
    ByRef StateOut As T_AppState)
'
'==============================================================================
'                            CAPTURE APP STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the current Excel Application state used by TW regression tests
'
' WHY THIS EXISTS
'   TW tests need a precise before/after baseline for the Application
'   properties intentionally modified by the shared TW manager
'
' INPUTS
'   StateOut
'     Output structure that receives the Application snapshot
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE STATE
'------------------------------------------------------------------------------
    'Copy the current Excel Application state into the output structure
        With Application
            StateOut.ScreenUpdating = .ScreenUpdating
            StateOut.EnableEvents = .EnableEvents
            StateOut.DisplayAlerts = .DisplayAlerts
            StateOut.Calculation = .Calculation
            StateOut.Cursor = .Cursor
        End With

End Sub

Private Function DelayForTimingMethod( _
    ByVal iMethod As Integer) _
    As Double
'
'==============================================================================
'                         DELAY FOR TIMING METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a practical per-method delay used by timing regression tests
'
' WHY THIS EXISTS
'   Different timing backends have different practical resolution
'   characteristics. In particular, method 6 is much coarser for test purposes
'   than the other timing methods
'
' INPUTS
'   iMethod
'     Timing backend identifier
'
' RETURNS
'   Double
'     Suggested delay in seconds for regression tests
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Use a longer delay for the coarse wall-clock method
        If iMethod = 6 Then
            DelayForTimingMethod = 1.1
            Exit Function
        End If

    'Use a shorter delay for the remaining methods
        DelayForTimingMethod = 0.05

End Function

'
'==============================================================================
'
'                              REGRESSION CASES
'
'==============================================================================

Private Sub Test_DefaultState()
'
'==============================================================================
'                              TEST DEFAULT STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the constructor/default state of a fresh cPerformanceManager
'   instance
'
' WHY THIS EXISTS
'   A deterministic constructor baseline is essential for predictable timing,
'   validation behavior, and TW lifecycle behavior
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Default constructor state"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT DEFAULT STATE
'------------------------------------------------------------------------------
    'Assert the default strict-mode state
        Test_Assert_EqualBoolean True, cPM.StrictMode, "StrictMode defaults to True"
    'Assert that no active timing session exists yet
        Test_Assert_EqualBoolean False, cPM.HasActiveSession, "HasActiveSession defaults to False"
    'Assert that no active method is bound yet
        Test_Assert_EqualLong 0, cPM.ActiveMethodID, "ActiveMethodID defaults to 0"

'------------------------------------------------------------------------------
' ASSERT DEFAULT TIMING CACHE
'------------------------------------------------------------------------------
    'Assert the default raw/cached timing values
        Test_Assert_ApproxDouble 0#, cPM.T1, 0#, "T1 defaults to 0"
    'Assert the default raw/cached timing values
        Test_Assert_ApproxDouble 0#, cPM.T2, 0#, "T2 defaults to 0"
    'Assert the default raw/cached timing values
        Test_Assert_ApproxDouble 0#, cPM.ET, 0#, "ET defaults to 0"

'------------------------------------------------------------------------------
' ASSERT DEFAULT TW STATE
'------------------------------------------------------------------------------
    'Assert that no TW session is active for the new instance
        Test_Assert_EqualBoolean False, cPM.TW_IsActive, "TW_IsActive defaults to False"
    'Assert that the shared TW manager is currently idle
        Test_Assert_EqualLong 0, cPM.TW_ActiveSessionCount, "TW_ActiveSessionCount defaults to 0"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_DefaultState"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_MethodName_ValidIndices()
'
'==============================================================================
'                        TEST METHODNAME VALID INDICES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates exact MethodName mappings for valid indices 1..6
'
' WHY THIS EXISTS
'   The method-name map is both a public diagnostic surface and a documentation
'   contract for the class
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "MethodName valid indices"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT METHOD LABELS
'------------------------------------------------------------------------------
    'Assert each documented method label exactly
        Test_Assert_EqualString "Timer", cPM.MethodName(1), "MethodName(1)"
        Test_Assert_EqualString "GetTickCount / GetTickCount64", cPM.MethodName(2), "MethodName(2)"
        Test_Assert_EqualString "timeGetTime", cPM.MethodName(3), "MethodName(3)"
        Test_Assert_EqualString "timeGetSystemTime", cPM.MethodName(4), "MethodName(4)"
        Test_Assert_EqualString "QPC", cPM.MethodName(5), "MethodName(5)"
        Test_Assert_EqualString "Now()", cPM.MethodName(6), "MethodName(6)"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_MethodName_ValidIndices"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_MethodName_InvalidIndices()
'
'==============================================================================
'                       TEST METHODNAME INVALID INDICES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates MethodName behavior for out-of-range indices
'
' WHY THIS EXISTS
'   The class documents that invalid MethodName indices should return
'   vbNullString rather than raising or returning misleading text
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "MethodName invalid indices"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT INVALID INDICES
'------------------------------------------------------------------------------
    'Assert vbNullString for representative invalid indices
        Test_Assert_EqualString vbNullString, cPM.MethodName(0), "MethodName(0)"
        Test_Assert_EqualString vbNullString, cPM.MethodName(-1), "MethodName(-1)"
        Test_Assert_EqualString vbNullString, cPM.MethodName(7), "MethodName(7)"
        Test_Assert_EqualString vbNullString, cPM.MethodName(99), "MethodName(99)"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_MethodName_InvalidIndices"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_StartTimer_SetsSessionState_AllMethods()
'
'==============================================================================
'                 TEST STARTTIMER SETS SESSION STATE ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates StartTimer session-state transitions for all timing methods
'
' WHY THIS EXISTS
'   StartTimer is the root of the session-bound timing model. A regression in
'   this area undermines elapsed-time validity across the whole class
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StartTimer sets session state for all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT SESSION STATE ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Start a new timing session for the selected backend
                cPM.StartTimer iMethod, False

            'Assert that a session is now active
                Test_Assert_EqualBoolean True, cPM.HasActiveSession, _
                                         "HasActiveSession after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the active method ID matches the requested method
                Test_Assert_EqualLong iMethod, cPM.ActiveMethodID, _
                                      "ActiveMethodID after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the method name is available for the active method
                Test_Assert_True (Len(cPM.MethodName(cPM.ActiveMethodID)) > 0), _
                                 "MethodName available after StartTimer(" & CStr(iMethod) & ")"

            'Assert that the raw start capture is nonnegative
                Test_Assert_NonNegativeDouble cPM.T1, _
                                              "T1 nonnegative after StartTimer(" & CStr(iMethod) & ")"

        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StartTimer_SetsSessionState_AllMethods"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_ElapsedSeconds_AllMethods()
'
'==============================================================================
'                      TEST ELAPSEDSECONDS ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates basic numeric elapsed-time behavior across all timing methods
'
' WHY THIS EXISTS
'   Numeric elapsed-time retrieval is the central timing output of the class and
'   must behave sensibly across all documented backends
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index
    Dim DelayS              As Double                 'Requested delay in seconds
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedSeconds across all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT ELAPSED-TIME BEHAVIOR ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)

            'Start a new timing session
                cPM.StartTimer iMethod, False

            'Perform a deliberate pause so the elapsed value should become positive
                cPM.Pause DelayS, 1

            'Read numeric elapsed time
                ElapsedS = cPM.ElapsedSeconds(iMethod)

            'Assert that the numeric elapsed time is nonnegative
                Test_Assert_NonNegativeDouble ElapsedS, _
                                              "ElapsedSeconds nonnegative for method " & CStr(iMethod)

            'Assert that the measured value is meaningfully positive relative to the delay
                Test_Assert_True (ElapsedS >= (DelayS / 4#)), _
                                 "ElapsedSeconds meaningfully positive for method " & CStr(iMethod)

        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedSeconds_AllMethods"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_ElapsedTime_AllMethods()
'
'==============================================================================
'                        TEST ELAPSEDTIME ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates formatted elapsed-time output across all timing methods
'
' WHY THIS EXISTS
'   ElapsedTime is the public display/reporting companion to ElapsedSeconds and
'   should return readable, semantically complete output for every backend
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index
    Dim DelayS              As Double                 'Requested delay in seconds
    Dim TextOut             As String                 'Formatted elapsed-time output

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedTime across all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT FORMATTED ELAPSED-TIME OUTPUT ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6
            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)
            'Start a new timing session
                cPM.StartTimer iMethod, False
            'Perform a deliberate pause
                cPM.Pause DelayS, 1
            'Read formatted elapsed time
                TextOut = cPM.ElapsedTime(iMethod)
            'Assert that the formatted string is non-empty
                Test_Assert_True (Len(TextOut) > 0), _
                                 "ElapsedTime non-empty for method " & CStr(iMethod)
            'Assert that the formatted string contains each documented unit group
                Test_Assert_ContainsString TextOut, "milliseconds", _
                                           "ElapsedTime contains milliseconds for method " & CStr(iMethod)
                Test_Assert_ContainsString TextOut, "microseconds", _
                                           "ElapsedTime contains microseconds for method " & CStr(iMethod)
                Test_Assert_ContainsString TextOut, "nanoseconds", _
                                           "ElapsedTime contains nanoseconds for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0
    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedTime_AllMethods"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub
Private Sub Test_ElapsedTime_FormatExistingSeconds()
'
'==============================================================================
'                TEST ELAPSEDTIME FORMAT EXISTING SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the ElapsedTime path that formats an already measured
'   elapsed-seconds value without taking a second timing sample
'
' WHY THIS EXISTS
'   This behavior avoids double measurement in callers that already captured
'   ElapsedSeconds separately for logging or further numeric processing
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds
    Dim TextOut             As String                 'Formatted elapsed-time output
    Dim T2Before            As Double                 'Cached raw end timestamp before formatting
    Dim ETBefore            As Double                 'Cached elapsed seconds before formatting

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ElapsedTime formats an existing elapsed-seconds value"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE MEASURED ELAPSED VALUE
'------------------------------------------------------------------------------
    'Start a QPC timing session
        cPM.StartTimer 5, False
    'Perform a short pause
        cPM.Pause 0.05, 1
    'Measure elapsed time numerically once
        ElapsedS = cPM.ElapsedSeconds(5)

'------------------------------------------------------------------------------
' CAPTURE PRE-FORMATTING STATE
'------------------------------------------------------------------------------
    'Capture T2 before formatting-only use
        T2Before = cPM.T2
    'Capture ET before formatting-only use
        ETBefore = cPM.ET

'------------------------------------------------------------------------------
' FORMAT EXISTING ELAPSED VALUE
'------------------------------------------------------------------------------
    'Format the existing elapsed-seconds value directly
        TextOut = cPM.ElapsedTime(, ElapsedS)

'------------------------------------------------------------------------------
' ASSERT FORMATTED OUTPUT
'------------------------------------------------------------------------------
    'Assert that the formatted string is non-empty
        Test_Assert_True (Len(TextOut) > 0), _
                         "ElapsedTime(, ElapsedSecondsIn) returns non-empty text"

    'Assert that the formatted string contains the documented unit groups
        Test_Assert_ContainsString TextOut, "milliseconds", _
                                   "ElapsedTime(, ElapsedSecondsIn) contains milliseconds"
        Test_Assert_ContainsString TextOut, "microseconds", _
                                   "ElapsedTime(, ElapsedSecondsIn) contains microseconds"
        Test_Assert_ContainsString TextOut, "nanoseconds", _
                                   "ElapsedTime(, ElapsedSecondsIn) contains nanoseconds"

'------------------------------------------------------------------------------
' ASSERT NO NEW TIMING SAMPLE
'------------------------------------------------------------------------------
    'Assert that formatting-only use does not take a new timing sample
        Test_Assert_ApproxDouble T2Before, cPM.T2, 0#, _
                                 "T2 unchanged when ElapsedTime formats an existing elapsed value"
    'Assert that formatting-only use does not update the cached elapsed value
        Test_Assert_ApproxDouble ETBefore, cPM.ET, 0#, _
                                 "ET unchanged when ElapsedTime formats an existing elapsed value"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ElapsedTime_FormatExistingSeconds"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_AlignedStart_AllMethods()
'
'==============================================================================
'                       TEST ALIGNEDSTART ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates aligned-start timing behavior across all timing methods
'
' WHY THIS EXISTS
'   AlignToNextTick is a specialized benchmark feature and should still behave
'   sanely across all documented backends
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index
    Dim DelayS              As Double                 'Requested delay in seconds
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "AlignToNextTick across all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT ALIGNED-START BEHAVIOR ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6
            'Choose a practical per-method delay
                DelayS = DelayForTimingMethod(iMethod)
            'Start a new aligned timing session
                cPM.StartTimer iMethod, True
            'Perform a deliberate pause
                cPM.Pause DelayS, 1
            'Read numeric elapsed time
                ElapsedS = cPM.ElapsedSeconds(iMethod)
            'Assert that the aligned elapsed time is nonnegative
                Test_Assert_NonNegativeDouble ElapsedS, _
                                              "Aligned ElapsedSeconds nonnegative for method " & CStr(iMethod)
            'Assert that the aligned elapsed time is meaningfully positive
                Test_Assert_True (ElapsedS >= (DelayS / 4#)), _
                                 "Aligned ElapsedSeconds meaningfully positive for method " & CStr(iMethod)
        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0
    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_AlignedStart_AllMethods"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub
Private Sub Test_Accessors_QPC()
'
'==============================================================================
'                           TEST ACCESSORS QPC
'------------------------------------------------------------------------------
' PURPOSE
'   Validates raw/cached accessor behavior after a QPC measurement
'
' WHY THIS EXISTS
'   T1, T2, and ET are explicit inspection/debugging surfaces and should remain
'   coherent with the underlying elapsed measurement
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Accessors after QPC measurement"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE QPC MEASUREMENT
'------------------------------------------------------------------------------
    'Start a QPC timing session
        cPM.StartTimer 5, False
    'Perform a short pause
        cPM.Pause 0.03, 1
    'Read numeric elapsed time through the public API
        ElapsedS = cPM.ElapsedSeconds(5)

'------------------------------------------------------------------------------
' ASSERT RAW ACCESSORS
'------------------------------------------------------------------------------
    'Assert that the raw captures are nonnegative
        Test_Assert_NonNegativeDouble cPM.T1, _
                                      "T1 nonnegative after QPC measurement"
        Test_Assert_NonNegativeDouble cPM.T2, _
                                      "T2 nonnegative after QPC measurement"
    'Assert that the raw end capture is not earlier than the raw start capture
        Test_Assert_True (cPM.T2 >= cPM.T1), _
                         "T2 >= T1 after QPC measurement"

'------------------------------------------------------------------------------
' ASSERT CACHED ELAPSED VALUE
'------------------------------------------------------------------------------
    'Assert that ET mirrors the cached elapsed value
        Test_Assert_ApproxDouble ElapsedS, cPM.ET, 0.000000001, _
                                 "ET matches ElapsedSeconds after QPC measurement"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0
    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Accessors_QPC"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_StrictMode_ElapsedBeforeStart()
'
'==============================================================================
'                   TEST STRICTMODE ELAPSED BEFORE START
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior when elapsed time is requested before
'   StartTimer
'
' WHY THIS EXISTS
'   Calling ElapsedSeconds before a timing session exists is a fundamental
'   contract violation that strict mode must reject
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Dummy               As Double                 'Throwaway target for the failing call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: ElapsedSeconds before StartTimer"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Force strict mode explicitly for clarity
        cPM.StrictMode = True

'------------------------------------------------------------------------------
' ASSERT STRICT-MODE FAILURE
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid elapsed-time read before StartTimer
        Dummy = cPM.ElapsedSeconds

    'Assert that an error was raised
        Test_Assert_True (Err.Number <> 0), _
                         "Strict mode raises when ElapsedSeconds is called before StartTimer"

    'Assert that the error description mentions StartTimer
        Test_Assert_ContainsString Err.Description, "StartTimer", _
                                   "Strict-mode error text mentions StartTimer"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_ElapsedBeforeStart"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_StrictMode_MethodMismatch()
'
'==============================================================================
'                     TEST STRICTMODE METHOD MISMATCH
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior for explicit elapsed-method mismatch
'
' WHY THIS EXISTS
'   The class is intentionally session-bound. Strict mode must reject attempts
'   to read elapsed time with a method different from the one that started the
'   session
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Dummy               As Double                 'Throwaway target for the failing call

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: explicit elapsed-method mismatch"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Force strict mode explicitly for clarity
        cPM.StrictMode = True
    'Start a session with method 1
        cPM.StartTimer 1, False
    'Perform a short pause so the session is live
        cPM.Pause 0.05, 1

'------------------------------------------------------------------------------
' ASSERT STRICT-MODE FAILURE
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid explicit elapsed read with a mismatched method
        Dummy = cPM.ElapsedSeconds(2)

    'Assert that an error was raised
        Test_Assert_True (Err.Number <> 0), _
                         "Strict mode raises on explicit elapsed-method mismatch"

    'Assert that the error description mentions the method mismatch
        Test_Assert_ContainsString Err.Description, "does not match", _
                                   "Strict-mode mismatch error text is informative"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_MethodMismatch"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_StrictMode_InvalidStartMethod()
'
'==============================================================================
'                  TEST STRICTMODE INVALID START METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Validates strict-mode behavior for invalid start-method input
'
' WHY THIS EXISTS
'   StartTimer should fail fast in strict mode when the caller passes an invalid
'   method identifier
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StrictMode: invalid StartTimer method"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Force strict mode explicitly for clarity
        cPM.StrictMode = True

'------------------------------------------------------------------------------
' ASSERT STRICT-MODE FAILURE
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt an invalid StartTimer call
        cPM.StartTimer 99, False

    'Assert that an error was raised
        Test_Assert_True (Err.Number <> 0), _
                         "Strict mode raises on invalid StartTimer method"

    'Assert that the error description mentions invalid timer method
        Test_Assert_ContainsString Err.Description, "Invalid timer method", _
                                   "Strict-mode invalid-start error text is informative"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StrictMode_InvalidStartMethod"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_NonStrictMode_InvalidStartFallback()
'
'==============================================================================
'               TEST NONSTRICTMODE INVALID START FALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-strict fallback behavior for invalid start-method input
'
' WHY THIS EXISTS
'   In non-strict mode the class documents that invalid start-method inputs are
'   coerced toward a usable backend rather than immediately rejected
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim DelayS              As Double                 'Requested delay in seconds
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "NonStrictMode: invalid StartTimer fallback"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Force non-strict mode
        cPM.StrictMode = False

'------------------------------------------------------------------------------
' APPLY NON-STRICT FALLBACK
'------------------------------------------------------------------------------
    'Call StartTimer with an invalid method
        cPM.StartTimer 99, False

'------------------------------------------------------------------------------
' ASSERT FALLBACK SESSION STATE
'------------------------------------------------------------------------------
    'Assert that a session is active after fallback
        Test_Assert_EqualBoolean True, cPM.HasActiveSession, _
                                 "Non-strict invalid StartTimer still establishes a session"

    'Assert that the resolved active method is valid
        Test_Assert_True ((cPM.ActiveMethodID = 5) Or (cPM.ActiveMethodID = 2)), _
                         "Non-strict invalid StartTimer resolves to method 5 or 2"

    'Assert that the resolved active method has a non-empty name
        Test_Assert_True (Len(cPM.MethodName(cPM.ActiveMethodID)) > 0), _
                         "Resolved fallback method has a valid MethodName"

'------------------------------------------------------------------------------
' ASSERT FALLBACK ELAPSED-TIME PATH
'------------------------------------------------------------------------------
    'Choose a practical delay for the resolved backend
        DelayS = DelayForTimingMethod(cPM.ActiveMethodID)

    'Perform a deliberate pause
        cPM.Pause DelayS, 1

    'Read elapsed time using the active-session path
        ElapsedS = cPM.ElapsedSeconds

    'Assert that the fallback path produces a nonnegative elapsed value
        Test_Assert_NonNegativeDouble ElapsedS, _
                                      "Non-strict fallback path returns nonnegative elapsed time"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_NonStrictMode_InvalidStartFallback"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_NonStrictMode_MethodMismatchFallback()
'
'==============================================================================
'            TEST NONSTRICTMODE METHOD MISMATCH FALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-strict fallback behavior for explicit elapsed-method mismatch
'
' WHY THIS EXISTS
'   In non-strict mode, an explicit elapsed-method mismatch should not raise.
'   Instead, the class should fall back to the active session method where
'   allowed
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "NonStrictMode: explicit elapsed-method mismatch fallback"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Force non-strict mode
        cPM.StrictMode = False

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a session with method 1
        cPM.StartTimer 1, False
    'Perform a short pause
        cPM.Pause 0.05, 1

'------------------------------------------------------------------------------
' APPLY NON-STRICT FALLBACK
'------------------------------------------------------------------------------
    'Request elapsed time with an explicit mismatched method
        ElapsedS = cPM.ElapsedSeconds(2)

'------------------------------------------------------------------------------
' ASSERT FALLBACK BEHAVIOR
'------------------------------------------------------------------------------
    'Assert that the active method remains the original session method
        Test_Assert_EqualLong 1, cPM.ActiveMethodID, _
                              "ActiveMethodID remains unchanged after non-strict mismatch fallback"

    'Assert that the fallback elapsed value is nonnegative
        Test_Assert_NonNegativeDouble ElapsedS, _
                                      "Non-strict mismatch fallback returns nonnegative elapsed time"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_NonStrictMode_MethodMismatchFallback"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_OverheadMeasurement_Seconds()
'
'==============================================================================
'                  TEST OVERHEADMEASUREMENT IN SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates numeric overhead-measurement helpers across all methods
'
' WHY THIS EXISTS
'   Benchmark-support helpers are part of the public API and should remain
'   callable and sane even for coarse timing methods
'
' NOTES
'   Coarse timing methods can legitimately report very small or zero overhead
'   values, so this test asserts nonnegativity rather than strict positivity
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index
    Dim OverheadS           As Double                 'Measured overhead in seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "OverheadMeasurement_Seconds across all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT OVERHEAD HELPERS ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Measure average near-empty timing overhead with a modest iteration count
                OverheadS = cPM.OverheadMeasurement_Seconds(iMethod, 25)

            'Assert that the reported overhead is nonnegative
                Test_Assert_NonNegativeDouble OverheadS, _
                                              "OverheadMeasurement_Seconds nonnegative for method " & CStr(iMethod)

        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_OverheadMeasurement_Seconds"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_OverheadMeasurement_Text()
'
'==============================================================================
'                    TEST OVERHEADMEASUREMENT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates formatted overhead-measurement reporting across all methods
'
' WHY THIS EXISTS
'   The text-reporting companion to the numeric overhead helper should remain
'   readable, non-empty, and semantically informative
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim iMethod             As Integer                'Timing backend index
    Dim TextOut             As String                 'Formatted overhead text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "OverheadMeasurement_Text across all methods"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT OVERHEAD TEXT ACROSS METHODS
'------------------------------------------------------------------------------
    'Iterate over every documented timing backend
        For iMethod = 1 To 6

            'Read formatted overhead text for the current backend using a small explicit iteration count
                TextOut = cPM.OverheadMeasurement_Text(iMethod, 25)

            'Assert that the formatted string is non-empty
                Test_Assert_True (Len(TextOut) > 0), _
                                 "OverheadMeasurement_Text non-empty for method " & CStr(iMethod)

            'Assert that the backend label appears in the formatted string
                Test_Assert_ContainsString TextOut, cPM.MethodName(iMethod), _
                                           "OverheadMeasurement_Text contains backend label for method " & CStr(iMethod)

            'Assert that the formatted string mentions seconds
                Test_Assert_ContainsString TextOut, "seconds", _
                                           "OverheadMeasurement_Text contains seconds for method " & CStr(iMethod)

        Next iMethod

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_OverheadMeasurement_Text"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Diagnostics_Properties()
'
'==============================================================================
'                     TEST DIAGNOSTICS PROPERTIES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates diagnostic and informational properties
'
' WHY THIS EXISTS
'   The class exposes several human-readable and numeric diagnostics that are
'   useful for environment inspection, troubleshooting, and documentation
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim TextOut             As String                 'Diagnostic text
    Dim QpcHz               As Double                 'Numeric QPC frequency

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Diagnostic and informational properties"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT TEXT DIAGNOSTICS
'------------------------------------------------------------------------------
    'Read and validate the nominal system tick-interval text
        TextOut = cPM.Get_SystemTickInterval
        Test_Assert_True (Len(TextOut) > 0), _
                         "Get_SystemTickInterval is non-empty"
        Test_Assert_ContainsString TextOut, "Tick Interval", _
                                   "Get_SystemTickInterval contains label text"

    'Read and validate the QPC tick-interval text
        TextOut = cPM.QPC_Get_SystemTickInterval
        Test_Assert_True (Len(TextOut) > 0), _
                         "QPC_Get_SystemTickInterval is non-empty"
        Test_Assert_ContainsString TextOut, "QPC Tick interval", _
                                   "QPC_Get_SystemTickInterval contains label text"

    'Read and validate the QPC frequency text
        TextOut = cPM.QPC_FrequencyPerSecond
        Test_Assert_True (Len(TextOut) > 0), _
                         "QPC_FrequencyPerSecond is non-empty"
        Test_Assert_ContainsString TextOut, "QPC Tick frequency", _
                                   "QPC_FrequencyPerSecond contains label text"

'------------------------------------------------------------------------------
' ASSERT NUMERIC DIAGNOSTICS
'------------------------------------------------------------------------------
    'Read and validate the numeric QPC frequency
        QpcHz = cPM.QPC_FrequencyPerSecond_Value
        Test_Assert_NonNegativeDouble QpcHz, _
                                      "QPC_FrequencyPerSecond_Value is nonnegative"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Diagnostics_Properties"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Pause_Method1()
'
'==============================================================================
'                           TEST PAUSE METHOD 1
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 1 (Sleep API) using QPC timing
'
' WHY THIS EXISTS
'   Pause method 1 is the lowest-overhead pause path and should produce a delay
'   reasonably close to the requested duration
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 1 (Sleep)"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT PAUSE METHOD 1
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, False
    'Pause for 1 second using method 1
        cPM.Pause 1, 1
    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)
    'Assert that the measured pause lies within a practical tolerance band
        Test_Assert_InRangeDouble 0.8, 1.25, ElapsedS, _
                                  "Pause method 1 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method1"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Pause_Method2()
'
'==============================================================================
'                           TEST PAUSE METHOD 2
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 2 (Timer + DoEvents loop) using QPC timing
'
' WHY THIS EXISTS
'   Pause method 2 is a yielding pause path and should still respect the
'   requested duration within a practical tolerance
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 2 (Timer + DoEvents)"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT PAUSE METHOD 2
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, False
    'Pause for 1 second using method 2
        cPM.Pause 1, 2
    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)
    'Assert that the measured pause lies within a practical tolerance band
        Test_Assert_InRangeDouble 0.8, 1.25, ElapsedS, _
                                  "Pause method 2 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method2"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Pause_Method3()
'
'==============================================================================
'                           TEST PAUSE METHOD 3
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 3 (Application.Wait) using QPC timing
'
' WHY THIS EXISTS
'   Application.Wait is coarse and should not be expected to behave like a
'   fine-grained pause, but it should still produce a reasonable whole-second
'   delay
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 3 (Application.Wait)"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT PAUSE METHOD 3
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, True
    'Pause for 1 second using method 3
        cPM.Pause 1, 3
    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)
    'Assert that the measured pause lies within a broad practical range
        Test_Assert_InRangeDouble 0.8, 1.5, ElapsedS, _
                                  "Pause method 3 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method3"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Pause_Method4()
'
'==============================================================================
'                           TEST PAUSE METHOD 4
'------------------------------------------------------------------------------
' PURPOSE
'   Validates Pause method 4 (Now + DoEvents loop) using QPC timing
'
' WHY THIS EXISTS
'   The Date/Now loop path is coarser and higher-overhead than Sleep or QPC,
'   but it should still approximate the requested delay within a practical
'   range
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedS            As Double                 'Measured elapsed seconds

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause method 4 (Now + DoEvents)"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT PAUSE METHOD 4
'------------------------------------------------------------------------------
    'Start QPC timing
        cPM.StartTimer 5, True
    'Pause for 1 second using method 4
        cPM.Pause 1, 4
    'Measure elapsed time using QPC
        ElapsedS = cPM.ElapsedSeconds(5)
    'Assert that the measured pause lies within a broad practical range
        Test_Assert_InRangeDouble 0.8, 1.5, ElapsedS, _
                                  "Pause method 4 elapsed range"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Method4"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_TW_BlankKeyValidation()
'
'==============================================================================
'                      TEST TW BLANK KEY VALIDATION
'------------------------------------------------------------------------------
' PURPOSE
'   Validates blank-key behavior in the shared TW manager
'
' WHY THIS EXISTS
'   Blank keys must remain rejected because they would otherwise create
'   collisions or ambiguous shared-session bookkeeping
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW manager blank-key validation"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' ASSERT BLANK-KEY BEGINSESSION
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next
    'Attempt an invalid blank-key begin
        PM_TW_BeginSession ""
    'Assert that blank-key begin raises
        Test_Assert_True (Err.Number <> 0), _
                         "PM_TW_BeginSession raises on blank key"
    'Clear the expected error state
        Err.Clear

'------------------------------------------------------------------------------
' ASSERT BLANK-KEY ENDSESSION
'------------------------------------------------------------------------------
    'Attempt an invalid blank-key end
        PM_TW_EndSession ""
    'Assert that blank-key end raises
        Test_Assert_True (Err.Number <> 0), _
                         "PM_TW_EndSession raises on blank key"
    'Clear the expected error state
        Err.Clear

'------------------------------------------------------------------------------
' ASSERT BLANK-KEY ISINSTANCEACTIVE
'------------------------------------------------------------------------------
    'Attempt an invalid blank-key activity query
        Call PM_TW_IsInstanceActive("")
    'Assert that blank-key activity query raises
        Test_Assert_True (Err.Number <> 0), _
                         "PM_TW_IsInstanceActive raises on blank key"
    'Clear the expected error state
        Err.Clear

'------------------------------------------------------------------------------
' RESTORE NORMAL ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Force the shared TW manager to a clean baseline on a best-effort basis
        On Error Resume Next
        PM_TW_EndAllSessions
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_BlankKeyValidation"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_TW_SingleInstance()
'
'==============================================================================
'                        TEST TW SINGLE INSTANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates single-instance TW lifecycle behavior
'
' WHY THIS EXISTS
'   The class publicly exposes TW lifecycle control, and the simplest shared
'   manager behavior must work correctly before overlapping cases are tested
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Baseline            As T_AppState             'Captured Application baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW single-instance lifecycle"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Capture the current Application baseline
        CaptureAppState Baseline

'------------------------------------------------------------------------------
' ASSERT PRECONDITIONS
'------------------------------------------------------------------------------
    'Assert that the instance starts inactive with zero shared TW sessions
        Test_Assert_EqualBoolean False, cPM.TW_IsActive, _
                                 "TW_IsActive before TW_Turn_OFF"
        Test_Assert_EqualLong 0, cPM.TW_ActiveSessionCount, _
                              "TW_ActiveSessionCount before TW_Turn_OFF"

'------------------------------------------------------------------------------
' ACTIVATE TW SUPPRESSION
'------------------------------------------------------------------------------
    'Begin TW suppression for the instance with no exemptions
        cPM.TW_Turn_OFF TW_Enum.None

'------------------------------------------------------------------------------
' ASSERT ACTIVE STATE
'------------------------------------------------------------------------------
    'Assert that the instance is now active
        Test_Assert_EqualBoolean True, cPM.TW_IsActive, _
                                 "TW_IsActive after TW_Turn_OFF"

    'Assert that exactly one shared TW session is active
        Test_Assert_EqualLong 1, cPM.TW_ActiveSessionCount, _
                              "TW_ActiveSessionCount after TW_Turn_OFF"

'------------------------------------------------------------------------------
' ASSERT FORCED APPLICATION STATE
'------------------------------------------------------------------------------
    'Assert forced benchmark/performance-state values
        Test_Assert_EqualBoolean False, Application.ScreenUpdating, _
                                 "ScreenUpdating forced OFF"
        Test_Assert_EqualBoolean False, Application.EnableEvents, _
                                 "EnableEvents forced OFF"
        Test_Assert_EqualBoolean False, Application.DisplayAlerts, _
                                 "DisplayAlerts forced OFF"
        Test_Assert_EqualLong xlCalculationManual, Application.Calculation, _
                              "Calculation forced MANUAL"
        Test_Assert_EqualLong xlWait, Application.Cursor, _
                              "Cursor forced WAIT"

'------------------------------------------------------------------------------
' DEACTIVATE TW SUPPRESSION
'------------------------------------------------------------------------------
    'End TW suppression for the instance
        cPM.TW_Turn_ON

'------------------------------------------------------------------------------
' ASSERT RESTORED STATE
'------------------------------------------------------------------------------
    'Assert that the instance is now inactive
        Test_Assert_EqualBoolean False, cPM.TW_IsActive, _
                                 "TW_IsActive after TW_Turn_ON"

    'Assert that the shared TW manager is idle again
        Test_Assert_EqualLong 0, cPM.TW_ActiveSessionCount, _
                              "TW_ActiveSessionCount after TW_Turn_ON"

'------------------------------------------------------------------------------
' ASSERT BASELINE RESTORATION
'------------------------------------------------------------------------------
    'Assert baseline restoration
        Test_Assert_EqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                                 "ScreenUpdating restored"
        Test_Assert_EqualBoolean Baseline.EnableEvents, Application.EnableEvents, _
                                 "EnableEvents restored"
        Test_Assert_EqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, _
                                 "DisplayAlerts restored"
        Test_Assert_EqualLong Baseline.Calculation, Application.Calculation, _
                              "Calculation restored"
        Test_Assert_EqualLong Baseline.Cursor, Application.Cursor, _
                              "Cursor restored"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_SingleInstance"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_TW_OverlappingInstances()
'
'==============================================================================
'                     TEST TW OVERLAPPING INSTANCES
'------------------------------------------------------------------------------
' PURPOSE
'   Validates overlapping multi-instance TW lifecycle behavior
'
' WHY THIS EXISTS
'   The shared TW manager exists specifically because overlapping class
'   instances must be coordinated safely. This is one of the most important
'   architectural regression surfaces in the project
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM1                As cPerformanceManager    'First class instance
    Dim cPM2                As cPerformanceManager    'Second class instance
    Dim Baseline            As T_AppState             'Captured Application baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW overlapping multi-instance lifecycle"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create both class instances
        Set cPM1 = New cPerformanceManager
        Set cPM2 = New cPerformanceManager
    'Capture the current Application baseline
        CaptureAppState Baseline

'------------------------------------------------------------------------------
' ACTIVATE INSTANCE 1
'------------------------------------------------------------------------------
    'Begin TW suppression on the first instance with no exemptions
        cPM1.TW_Turn_OFF TW_Enum.None

'------------------------------------------------------------------------------
' ASSERT STATE AFTER INSTANCE 1 BEGINS
'------------------------------------------------------------------------------
    'Assert the shared active-session count after instance 1 begins
        Test_Assert_EqualLong 1, cPM1.TW_ActiveSessionCount, _
                              "Shared TW count after instance 1 begins"

'------------------------------------------------------------------------------
' ACTIVATE INSTANCE 2
'------------------------------------------------------------------------------
    'Begin TW suppression on the second instance while exempting ScreenUpdating
        cPM2.TW_Turn_OFF TW_Enum.ScreenUpdating

'------------------------------------------------------------------------------
' ASSERT OVERLAPPING ACTIVE STATE
'------------------------------------------------------------------------------
    'Assert the shared active-session count after instance 2 begins
        Test_Assert_EqualLong 2, cPM2.TW_ActiveSessionCount, _
                              "Shared TW count after instance 2 begins"

    'Assert that ScreenUpdating is still forced OFF because instance 1 still requires it
        Test_Assert_EqualBoolean False, Application.ScreenUpdating, _
                                 "ScreenUpdating remains forced OFF with overlapping sessions"

    'Assert that the remaining shared flags are still forced OFF / MANUAL / WAIT
        Test_Assert_EqualBoolean False, Application.EnableEvents, _
                                 "EnableEvents remains forced OFF with overlapping sessions"
        Test_Assert_EqualBoolean False, Application.DisplayAlerts, _
                                 "DisplayAlerts remains forced OFF with overlapping sessions"
        Test_Assert_EqualLong xlCalculationManual, Application.Calculation, _
                              "Calculation remains MANUAL with overlapping sessions"
        Test_Assert_EqualLong xlWait, Application.Cursor, _
                              "Cursor remains WAIT with overlapping sessions"

'------------------------------------------------------------------------------
' END INSTANCE 1
'------------------------------------------------------------------------------
    'End the first instance's TW participation
        cPM1.TW_Turn_ON

'------------------------------------------------------------------------------
' ASSERT STATE AFTER INSTANCE 1 ENDS
'------------------------------------------------------------------------------
    'Assert the shared active-session count after instance 1 ends
        Test_Assert_EqualLong 1, cPM2.TW_ActiveSessionCount, _
                              "Shared TW count after instance 1 ends"

    'Assert instance-local activity state after instance 1 ends
        Test_Assert_EqualBoolean False, cPM1.TW_IsActive, _
                                 "Instance 1 inactive after TW_Turn_ON"
        Test_Assert_EqualBoolean True, cPM2.TW_IsActive, _
                                 "Instance 2 still active after instance 1 ends"

    'Assert that ScreenUpdating now returns to baseline because the remaining
    'instance exempts that flag
        Test_Assert_EqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                                 "ScreenUpdating restored to baseline when only instance 2 remains"

    'Assert that the remaining flags are still forced by the second instance
        Test_Assert_EqualBoolean False, Application.EnableEvents, _
                                 "EnableEvents still forced OFF by instance 2"
        Test_Assert_EqualBoolean False, Application.DisplayAlerts, _
                                 "DisplayAlerts still forced OFF by instance 2"
        Test_Assert_EqualLong xlCalculationManual, Application.Calculation, _
                              "Calculation still MANUAL by instance 2"
        Test_Assert_EqualLong xlWait, Application.Cursor, _
                              "Cursor still WAIT by instance 2"

'------------------------------------------------------------------------------
' END INSTANCE 2
'------------------------------------------------------------------------------
    'End the second instance's TW participation
        cPM2.TW_Turn_ON

'------------------------------------------------------------------------------
' ASSERT FINAL RESTORED STATE
'------------------------------------------------------------------------------
    'Assert the shared manager is now idle
        Test_Assert_EqualLong 0, cPM2.TW_ActiveSessionCount, _
                              "Shared TW count after instance 2 ends"

    'Assert full baseline restoration
        Test_Assert_EqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                                 "ScreenUpdating restored after final TW session ends"
        Test_Assert_EqualBoolean Baseline.EnableEvents, Application.EnableEvents, _
                                 "EnableEvents restored after final TW session ends"
        Test_Assert_EqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, _
                                 "DisplayAlerts restored after final TW session ends"
        Test_Assert_EqualLong Baseline.Calculation, Application.Calculation, _
                              "Calculation restored after final TW session ends"
        Test_Assert_EqualLong Baseline.Cursor, Application.Cursor, _
                              "Cursor restored after final TW session ends"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the first instance on a best-effort basis
        On Error Resume Next
        If Not cPM1 Is Nothing Then
            cPM1.ResetEnvironment
            Set cPM1 = Nothing
        End If

    'Release any environment changes held by the second instance on a best-effort basis
        If Not cPM2 Is Nothing Then
            cPM2.ResetEnvironment
            Set cPM2 = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_OverlappingInstances"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_ResetEnvironment_Idempotent()
'
'==============================================================================
'                   TEST RESETENVIRONMENT IDEMPOTENT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that ResetEnvironment is safe to call more than once and correctly
'   cleans up active environment changes
'
' WHY THIS EXISTS
'   ResetEnvironment is the explicit cleanup contract for the class. Its
'   idempotence is important for defensive calling patterns and error-handling
'   cleanup paths
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Baseline            As T_AppState             'Captured Application baseline
    Dim Dummy               As Double                 'Throwaway elapsed-time holder

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ResetEnvironment idempotence"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Capture the current Application baseline
        CaptureAppState Baseline

'------------------------------------------------------------------------------
' EXERCISE CLEANUP SURFACES
'------------------------------------------------------------------------------
    'Activate TW suppression for the instance
        cPM.TW_Turn_OFF TW_Enum.None
    'Start method 3 so that timer-resolution activation may occur
        cPM.StartTimer 3, False
    'Perform a short pause and read elapsed time to exercise the method 3 path
        cPM.Pause 0.03, 1
        Dummy = cPM.ElapsedSeconds(3)

'------------------------------------------------------------------------------
' APPLY EXPLICIT CLEANUP TWICE
'------------------------------------------------------------------------------
    'Call the explicit cleanup routine for the first time
        cPM.ResetEnvironment
    'Call the explicit cleanup routine a second time to validate idempotence
        cPM.ResetEnvironment

'------------------------------------------------------------------------------
' ASSERT INSTANCE / SHARED STATE
'------------------------------------------------------------------------------
    'Assert that the instance is no longer active in TW
        Test_Assert_EqualBoolean False, cPM.TW_IsActive, _
                                 "TW_IsActive is False after repeated ResetEnvironment"
    'Assert that the shared TW manager is idle
        Test_Assert_EqualLong 0, cPM.TW_ActiveSessionCount, _
                              "TW_ActiveSessionCount is 0 after repeated ResetEnvironment"

'------------------------------------------------------------------------------
' ASSERT APPLICATION BASELINE RESTORATION
'------------------------------------------------------------------------------
    'Assert Application baseline restoration
        Test_Assert_EqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                                 "ScreenUpdating restored after repeated ResetEnvironment"
        Test_Assert_EqualBoolean Baseline.EnableEvents, Application.EnableEvents, _
                                 "EnableEvents restored after repeated ResetEnvironment"
        Test_Assert_EqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, _
                                 "DisplayAlerts restored after repeated ResetEnvironment"
        Test_Assert_EqualLong Baseline.Calculation, Application.Calculation, _
                              "Calculation restored after repeated ResetEnvironment"
        Test_Assert_EqualLong Baseline.Cursor, Application.Cursor, _
                              "Cursor restored after repeated ResetEnvironment"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any remaining environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ResetEnvironment_Idempotent"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_Checkpoint_BeforeStart()
'
'==============================================================================
'                        TEST CHECKPOINT BEFORE START
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that Checkpoint raises before StartTimer
'
' WHY THIS EXISTS
'   Checkpoint capture is valid only within an active timing session
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Checkpoint before StartTimer"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT CHECKPOINT REJECTION
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt checkpoint capture before a timing session exists
        cPM.Checkpoint "Phase 1"

    'Assert that an error was raised
        Test_Assert_True (Err.Number <> 0), _
                         "Checkpoint raises before StartTimer"

    'Assert that the error text mentions StartTimer
        Test_Assert_ContainsString Err.Description, "StartTimer", _
                                   "Checkpoint-before-start error text mentions StartTimer"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Checkpoint_BeforeStart"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_SetRunLabel_BeforeFirstCheckpoint()
'
'==============================================================================
'                 TEST SETRUNLABEL BEFORE FIRST CHECKPOINT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates SetRunLabel behavior before the first checkpoint
'
' WHY THIS EXISTS
'   Run labels are intended to tag structured checkpoint output for the current
'   timing session
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "SetRunLabel before first checkpoint"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' APPLY RUN LABEL
'------------------------------------------------------------------------------
    'Assign the run label before any checkpoint is captured
        cPM.SetRunLabel "Run A"

'------------------------------------------------------------------------------
' ASSERT RUN-LABEL STATE
'------------------------------------------------------------------------------
    'Assert that the current run label was stored
        Test_Assert_EqualString "Run A", cPM.RunLabel, _
                                "RunLabel stored before first checkpoint"

'------------------------------------------------------------------------------
' CAPTURE ONE CHECKPOINT
'------------------------------------------------------------------------------
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture one named checkpoint
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' ASSERT EXPORTED RUN LABEL
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

    'Assert one captured checkpoint
        Test_Assert_EqualLong 1, cPM.CheckpointCount, _
                              "CheckpointCount after first labeled checkpoint"

    'Assert that the exported run label matches the assigned label
        Test_Assert_EqualString "Run A", CStr(Arr(2, 1)), _
                                "ReportAsArray exports the assigned run label"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_SetRunLabel_BeforeFirstCheckpoint"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_SetRunLabel_AfterFirstCheckpoint()
'
'==============================================================================
'                  TEST SETRUNLABEL AFTER FIRST CHECKPOINT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that SetRunLabel raises after checkpoint capture has begun
'
' WHY THIS EXISTS
'   The class contract requires the run label to be set before the first
'   checkpoint of the current timing session
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "SetRunLabel after first checkpoint"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' ASSERT RUN-LABEL REJECTION
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next

    'Attempt to assign a run label after checkpoint capture has begun
        cPM.SetRunLabel "Late Label"

    'Assert that an error was raised
        Test_Assert_True (Err.Number <> 0), _
                         "SetRunLabel raises after first checkpoint"

    'Assert that the error text mentions the first checkpoint rule
        Test_Assert_ContainsString Err.Description, "first checkpoint", _
                                   "SetRunLabel-after-checkpoint error text is informative"

    'Clear the expected error state
        Err.Clear

    'Restore normal error handling for the remainder of the case
        On Error GoTo CleanFail

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_SetRunLabel_AfterFirstCheckpoint"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_Checkpoint_DefaultName_WhenBlank()
'
'==============================================================================
'                 TEST CHECKPOINT DEFAULT NAME WHEN BLANK
'------------------------------------------------------------------------------
' PURPOSE
'   Validates automatic checkpoint naming when the supplied label is blank
'
' WHY THIS EXISTS
'   Structured reports should remain readable even when the caller omits a
'   checkpoint name
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Checkpoint default name when blank"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Perform a short pause
        cPM.Pause 0.02, 1

'------------------------------------------------------------------------------
' CAPTURE BLANK-NAME CHECKPOINT
'------------------------------------------------------------------------------
    'Capture a checkpoint with a blank label
        cPM.Checkpoint vbNullString

'------------------------------------------------------------------------------
' ASSERT GENERATED NAME
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

    'Assert one captured checkpoint
        Test_Assert_EqualLong 1, cPM.CheckpointCount, _
                              "CheckpointCount after blank-name checkpoint"

    'Assert that the generated checkpoint name is used
        Test_Assert_EqualString "Checkpoint 1", CStr(Arr(2, 3)), _
                                "ReportAsArray uses the default generated checkpoint name"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Checkpoint_DefaultName_WhenBlank"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_CheckpointCount_And_ReportArray()
'
'==============================================================================
'                  TEST CHECKPOINTCOUNT AND REPORTASARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Validates checkpoint count and structured array export
'
' WHY THIS EXISTS
'   ReportAsArray is the main machine-readable reporting surface for structured
'   checkpoints and should remain stable in shape and content
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "CheckpointCount and ReportAsArray"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Assign a run label
        cPM.SetRunLabel "Run A"

'------------------------------------------------------------------------------
' CAPTURE CHECKPOINTS
'------------------------------------------------------------------------------
    'Perform a short pause before the first checkpoint
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1", "First phase"

    'Perform a short pause before the second checkpoint
        cPM.Pause 0.02, 1
    'Capture the second checkpoint
        cPM.Checkpoint "Phase 2"

'------------------------------------------------------------------------------
' EXPORT STRUCTURED REPORT
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

'------------------------------------------------------------------------------
' ASSERT CHECKPOINT COUNT AND ARRAY SHAPE
'------------------------------------------------------------------------------
    'Assert the number of captured checkpoints
        Test_Assert_EqualLong 2, cPM.CheckpointCount, _
                              "CheckpointCount after two checkpoints"

    'Assert export row count including the header row
        Test_Assert_EqualLong 3, UBound(Arr, 1), _
                              "ReportAsArray row count including header"

    'Assert export column count
        Test_Assert_EqualLong 8, UBound(Arr, 2), _
                              "ReportAsArray column count"

'------------------------------------------------------------------------------
' ASSERT HEADER ROW
'------------------------------------------------------------------------------
    'Assert header column 1
        Test_Assert_EqualString "RunLabel", CStr(Arr(1, 1)), "ReportAsArray header column 1"
    'Assert header column 2
        Test_Assert_EqualString "Seq", CStr(Arr(1, 2)), "ReportAsArray header column 2"
    'Assert header column 3
        Test_Assert_EqualString "Checkpoint", CStr(Arr(1, 3)), "ReportAsArray header column 3"
    'Assert header column 4
        Test_Assert_EqualString "Note", CStr(Arr(1, 4)), "ReportAsArray header column 4"
    'Assert header column 5
        Test_Assert_EqualString "MethodID", CStr(Arr(1, 5)), "ReportAsArray header column 5"
    'Assert header column 6
        Test_Assert_EqualString "MethodName", CStr(Arr(1, 6)), "ReportAsArray header column 6"
    'Assert header column 7
        Test_Assert_EqualString "DeltaSeconds", CStr(Arr(1, 7)), "ReportAsArray header column 7"
    'Assert header column 8
        Test_Assert_EqualString "CumulativeSeconds", CStr(Arr(1, 8)), "ReportAsArray header column 8"

'------------------------------------------------------------------------------
' ASSERT DATA ROWS
'------------------------------------------------------------------------------
    'Assert run label export
        Test_Assert_EqualString "Run A", CStr(Arr(2, 1)), "ReportAsArray row 2 run label"
        Test_Assert_EqualString "Run A", CStr(Arr(3, 1)), "ReportAsArray row 3 run label"

    'Assert sequence export
        Test_Assert_EqualLong 1, Arr(2, 2), "ReportAsArray row 2 sequence"
        Test_Assert_EqualLong 2, Arr(3, 2), "ReportAsArray row 3 sequence"

    'Assert checkpoint-name export
        Test_Assert_EqualString "Phase 1", CStr(Arr(2, 3)), "ReportAsArray row 2 checkpoint name"
        Test_Assert_EqualString "Phase 2", CStr(Arr(3, 3)), "ReportAsArray row 3 checkpoint name"

    'Assert note export
        Test_Assert_EqualString "First phase", CStr(Arr(2, 4)), "ReportAsArray row 2 note"
        Test_Assert_EqualString vbNullString, CStr(Arr(3, 4)), "ReportAsArray row 3 note"

    'Assert method metadata export
        Test_Assert_EqualLong 5, Arr(2, 5), "ReportAsArray row 2 method ID"
        Test_Assert_EqualLong 5, Arr(3, 5), "ReportAsArray row 3 method ID"
        Test_Assert_EqualString "QPC", CStr(Arr(2, 6)), "ReportAsArray row 2 method name"
        Test_Assert_EqualString "QPC", CStr(Arr(3, 6)), "ReportAsArray row 3 method name"

'------------------------------------------------------------------------------
' ASSERT TIMING VALUES
'------------------------------------------------------------------------------
    'Assert nonnegative delta/cumulative values
        Test_Assert_NonNegativeDouble CDbl(Arr(2, 7)), "ReportAsArray row 2 delta is nonnegative"
        Test_Assert_NonNegativeDouble CDbl(Arr(3, 7)), "ReportAsArray row 3 delta is nonnegative"
        Test_Assert_NonNegativeDouble CDbl(Arr(2, 8)), "ReportAsArray row 2 cumulative is nonnegative"
        Test_Assert_NonNegativeDouble CDbl(Arr(3, 8)), "ReportAsArray row 3 cumulative is nonnegative"

    'Assert monotonic cumulative timing
        Test_Assert_True (CDbl(Arr(3, 8)) >= CDbl(Arr(2, 8))), _
                         "ReportAsArray cumulative timing is nondecreasing"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_CheckpointCount_And_ReportArray"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub

Private Sub Test_ReportAsText_Empty()
'
'==============================================================================
'                        TEST REPORTASTEXT EMPTY
'------------------------------------------------------------------------------
' PURPOSE
'   Validates ReportAsText behavior when no checkpoints exist
'
' WHY THIS EXISTS
'   The class should return a deterministic readable message even when no
'   checkpoint has been captured
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim TextOut             As String                 'Structured text report

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ReportAsText when no checkpoints exist"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' READ EMPTY REPORT
'------------------------------------------------------------------------------
    'Read the text report before any checkpoint is captured
        TextOut = cPM.ReportAsText

'------------------------------------------------------------------------------
' ASSERT EMPTY REPORT
'------------------------------------------------------------------------------
    'Assert that the empty-report message is returned
        Test_Assert_EqualString "No checkpoints captured.", TextOut, _
                                "ReportAsText returns the empty-report message before checkpoint capture"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ReportAsText_Empty"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_ReportAsText_WithCheckpoints()
'
'==============================================================================
'                 TEST REPORTASTEXT WITH CHECKPOINTS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates ReportAsText behavior after checkpoint capture
'
' WHY THIS EXISTS
'   ReportAsText is the main human-readable reporting surface for structured
'   checkpoint output and should remain readable and semantically complete
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim TextOut             As String                 'Structured text report

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ReportAsText with checkpoints"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Assign a run label
        cPM.SetRunLabel "Run A"

'------------------------------------------------------------------------------
' CAPTURE CHECKPOINTS
'------------------------------------------------------------------------------
    'Perform a short pause before the first checkpoint
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1", "First phase"

    'Perform a short pause before the second checkpoint
        cPM.Pause 0.02, 1
    'Capture the second checkpoint
        cPM.Checkpoint "Phase 2"

'------------------------------------------------------------------------------
' READ TEXT REPORT
'------------------------------------------------------------------------------
    'Read the human-readable checkpoint report
        TextOut = cPM.ReportAsText

'------------------------------------------------------------------------------
' ASSERT TEXT REPORT CONTENT
'------------------------------------------------------------------------------
    'Assert that the report title is present
        Test_Assert_ContainsString TextOut, "CHECKPOINT REPORT", _
                                   "ReportAsText contains the report title"
    'Assert that the run label is present
        Test_Assert_ContainsString TextOut, "RunLabel=Run A", _
                                   "ReportAsText contains the run label"
    'Assert that the checkpoint legend line is present
        Test_Assert_ContainsString TextOut, "Seq | Checkpoint | DeltaSeconds | CumulativeSeconds | MethodName | Note", _
                                   "ReportAsText contains the column legend"
    'Assert that the first checkpoint name is present
        Test_Assert_ContainsString TextOut, "Phase 1", _
                                   "ReportAsText contains the first checkpoint name"
    'Assert that the second checkpoint name is present
        Test_Assert_ContainsString TextOut, "Phase 2", _
                                   "ReportAsText contains the second checkpoint name"
    'Assert that the note text is present
        Test_Assert_ContainsString TextOut, "First phase", _
                                   "ReportAsText contains the first checkpoint note"
    'Assert that the method name is present
        Test_Assert_ContainsString TextOut, "QPC", _
                                   "ReportAsText contains the method name"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ReportAsText_WithCheckpoints"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_ClearCheckpoints()
'
'==============================================================================
'                        TEST CLEARCHECKPOINTS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates ClearCheckpoints behavior
'
' WHY THIS EXISTS
'   Callers should be able to clear structured checkpoint/report state without
'   recreating the class instance
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export
    Dim TextOut             As String                 'Structured text report

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ClearCheckpoints resets checkpoint state"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Assign a run label
        cPM.SetRunLabel "Run A"
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' CLEAR CHECKPOINT STATE
'------------------------------------------------------------------------------
    'Clear the structured checkpoint/report state
        cPM.ClearCheckpoints

'------------------------------------------------------------------------------
' ASSERT CLEARED STATE
'------------------------------------------------------------------------------
    'Assert that the checkpoint counter is reset
        Test_Assert_EqualLong 0, cPM.CheckpointCount, _
                              "CheckpointCount is reset by ClearCheckpoints"
    'Assert that the run label is reset
        Test_Assert_EqualString vbNullString, cPM.RunLabel, _
                                "RunLabel is reset by ClearCheckpoints"

'------------------------------------------------------------------------------
' ASSERT CLEARED EXPORT SURFACES
'------------------------------------------------------------------------------
    'Read the structured export array after clearing
        Arr = cPM.ReportAsArray
    'Read the text report after clearing
        TextOut = cPM.ReportAsText

    'Assert header-only array shape
        Test_Assert_EqualLong 1, UBound(Arr, 1), _
                              "ReportAsArray returns header-only row count after ClearCheckpoints"
        Test_Assert_EqualLong 8, UBound(Arr, 2), _
                              "ReportAsArray column count remains stable after ClearCheckpoints"

    'Assert the empty-report text path
        Test_Assert_EqualString "No checkpoints captured.", TextOut, _
                                "ReportAsText returns the empty-report message after ClearCheckpoints"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ClearCheckpoints"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_StartTimer_ClearsCheckpointState()
'
'==============================================================================
'                TEST STARTTIMER CLEARS CHECKPOINT STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that a new StartTimer session clears checkpoint/report state
'
' WHY THIS EXISTS
'   Checkpoints and run label belong to a timing session and should be reset
'   deterministically when a new session begins
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export
    Dim TextOut             As String                 'Structured text report

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "StartTimer clears checkpoint state"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE FIRST SESSION
'------------------------------------------------------------------------------
    'Start the first timing session
        cPM.StartTimer 5, False
    'Assign a run label for the first session
        cPM.SetRunLabel "Run A"
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture one checkpoint in the first session
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' ASSERT PRE-RESET STATE
'------------------------------------------------------------------------------
    'Assert one checkpoint before starting the new session
        Test_Assert_EqualLong 1, cPM.CheckpointCount, _
                              "CheckpointCount before starting a new session"
    'Assert the run label before starting the new session
        Test_Assert_EqualString "Run A", cPM.RunLabel, _
                                "RunLabel before starting a new session"

'------------------------------------------------------------------------------
' START NEW SESSION
'------------------------------------------------------------------------------
    'Start a new timing session, which should reset checkpoint/report state
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' ASSERT RESET STATE
'------------------------------------------------------------------------------
    'Assert that checkpoint count is reset for the new session
        Test_Assert_EqualLong 0, cPM.CheckpointCount, _
                              "CheckpointCount is reset by StartTimer"
    'Assert that the run label is reset for the new session
        Test_Assert_EqualString vbNullString, cPM.RunLabel, _
                                "RunLabel is reset by StartTimer"

'------------------------------------------------------------------------------
' ASSERT RESET EXPORT SURFACES
'------------------------------------------------------------------------------
    'Read the structured export array after the new session starts
        Arr = cPM.ReportAsArray
    'Read the text report after the new session starts
        TextOut = cPM.ReportAsText

    'Assert header-only array shape after session reset
        Test_Assert_EqualLong 1, UBound(Arr, 1), _
                              "ReportAsArray returns header-only row count after StartTimer resets checkpoint state"
        Test_Assert_EqualLong 8, UBound(Arr, 2), _
                              "ReportAsArray column count remains stable after StartTimer resets checkpoint state"

    'Assert the empty-report text path after session reset
        Test_Assert_EqualString "No checkpoints captured.", TextOut, _
                                "ReportAsText returns the empty-report message after StartTimer resets checkpoint state"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_StartTimer_ClearsCheckpointState"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub
Private Sub Test_ReportAsArray_Empty_Direct()
'
'==============================================================================
'                    TEST REPORTASARRAY EMPTY DIRECT
'------------------------------------------------------------------------------
' PURPOSE
'   Validates ReportAsArray behavior when no checkpoints exist
'
' WHY THIS EXISTS
'   The class should return a deterministic header-only array even when no
'   checkpoint has been captured
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ReportAsArray direct empty export"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' READ EMPTY EXPORT
'------------------------------------------------------------------------------
    'Read the structured report before any checkpoint is captured
        Arr = cPM.ReportAsArray

'------------------------------------------------------------------------------
' ASSERT ARRAY SHAPE
'------------------------------------------------------------------------------
    'Assert header-only row count
        Test_Assert_EqualLong 1, UBound(Arr, 1), _
                              "ReportAsArray returns a single header row when no checkpoints exist"
    'Assert stable column count
        Test_Assert_EqualLong 8, UBound(Arr, 2), _
                              "ReportAsArray returns 8 columns when no checkpoints exist"

'------------------------------------------------------------------------------
' ASSERT HEADER ROW
'------------------------------------------------------------------------------
    'Assert header column 1
        Test_Assert_EqualString "RunLabel", CStr(Arr(1, 1)), _
                                "ReportAsArray empty export header column 1"
    'Assert header column 2
        Test_Assert_EqualString "Seq", CStr(Arr(1, 2)), _
                                "ReportAsArray empty export header column 2"
    'Assert header column 3
        Test_Assert_EqualString "Checkpoint", CStr(Arr(1, 3)), _
                                "ReportAsArray empty export header column 3"
    'Assert header column 4
        Test_Assert_EqualString "Note", CStr(Arr(1, 4)), _
                                "ReportAsArray empty export header column 4"
    'Assert header column 5
        Test_Assert_EqualString "MethodID", CStr(Arr(1, 5)), _
                                "ReportAsArray empty export header column 5"
    'Assert header column 6
        Test_Assert_EqualString "MethodName", CStr(Arr(1, 6)), _
                                "ReportAsArray empty export header column 6"
    'Assert header column 7
        Test_Assert_EqualString "DeltaSeconds", CStr(Arr(1, 7)), _
                                "ReportAsArray empty export header column 7"
    'Assert header column 8
        Test_Assert_EqualString "CumulativeSeconds", CStr(Arr(1, 8)), _
                                "ReportAsArray empty export header column 8"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ReportAsArray_Empty_Direct"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_Checkpoint_DeltaAndCumulativeSemantics()
'
'==============================================================================
'            TEST CHECKPOINT DELTA AND CUMULATIVE SEMANTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates semantic relationships between checkpoint delta and cumulative
'   timing values
'
' WHY THIS EXISTS
'   Structured checkpoint reporting is only useful if exported timing values
'   remain internally coherent
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export
    Dim Delta1              As Double                 'First checkpoint delta
    Dim Delta2              As Double                 'Second checkpoint delta
    Dim Cum1                As Double                 'First checkpoint cumulative
    Dim Cum2                As Double                 'Second checkpoint cumulative

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Checkpoint delta/cumulative semantics"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' CAPTURE CHECKPOINTS
'------------------------------------------------------------------------------
    'Perform a short pause before the first checkpoint
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1"

    'Perform a short pause before the second checkpoint
        cPM.Pause 0.02, 1
    'Capture the second checkpoint
        cPM.Checkpoint "Phase 2"

'------------------------------------------------------------------------------
' EXPORT STRUCTURED REPORT
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

'------------------------------------------------------------------------------
' RESOLVE TIMING VALUES
'------------------------------------------------------------------------------
    'Resolve the first checkpoint delta
        Delta1 = CDbl(Arr(2, 7))
    'Resolve the second checkpoint delta
        Delta2 = CDbl(Arr(3, 7))
    'Resolve the first checkpoint cumulative time
        Cum1 = CDbl(Arr(2, 8))
    'Resolve the second checkpoint cumulative time
        Cum2 = CDbl(Arr(3, 8))

'------------------------------------------------------------------------------
' ASSERT BASIC NONNEGATIVITY
'------------------------------------------------------------------------------
    'Assert nonnegative first checkpoint delta
        Test_Assert_NonNegativeDouble Delta1, _
                                      "First checkpoint delta is nonnegative"
    'Assert nonnegative second checkpoint delta
        Test_Assert_NonNegativeDouble Delta2, _
                                      "Second checkpoint delta is nonnegative"
    'Assert nonnegative first checkpoint cumulative time
        Test_Assert_NonNegativeDouble Cum1, _
                                      "First checkpoint cumulative time is nonnegative"
    'Assert nonnegative second checkpoint cumulative time
        Test_Assert_NonNegativeDouble Cum2, _
                                      "Second checkpoint cumulative time is nonnegative"

'------------------------------------------------------------------------------
' ASSERT FIRST-CHECKPOINT SEMANTICS
'------------------------------------------------------------------------------
    'Assert that the first checkpoint cumulative time matches its delta
        Test_Assert_ApproxDouble Delta1, Cum1, 0.000000001, _
                                 "First checkpoint cumulative time matches first checkpoint delta"

'------------------------------------------------------------------------------
' ASSERT MONOTONIC / DIFFERENCE SEMANTICS
'------------------------------------------------------------------------------
    'Assert that cumulative timing is nondecreasing
        Test_Assert_True (Cum2 >= Cum1), _
                         "Checkpoint cumulative timing is nondecreasing"

    'Assert that each cumulative value is at least as large as its delta
        Test_Assert_True (Cum1 >= Delta1), _
                         "First checkpoint cumulative time is at least as large as first checkpoint delta"
        Test_Assert_True (Cum2 >= Delta2), _
                         "Second checkpoint cumulative time is at least as large as second checkpoint delta"

    'Assert that the second delta equals the incremental cumulative difference
        Test_Assert_ApproxDouble Delta2, (Cum2 - Cum1), 0.000000001, _
                                 "Second checkpoint delta matches the incremental cumulative difference"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Checkpoint_DeltaAndCumulativeSemantics"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_ClearCheckpoints_ThenReuseSession()
'
'==============================================================================
'               TEST CLEARCHECKPOINTS THEN REUSE SESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that the same session can continue normally after ClearCheckpoints
'
' WHY THIS EXISTS
'   Callers may want to clear structured checkpoint/report state and then
'   continue using the same class instance without starting over
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "ClearCheckpoints then reuse session"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE SESSION
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture the first checkpoint
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' CLEAR CHECKPOINT STATE
'------------------------------------------------------------------------------
    'Clear the structured checkpoint/report state
        cPM.ClearCheckpoints

'------------------------------------------------------------------------------
' REUSE SESSION
'------------------------------------------------------------------------------
    'Perform another short pause
        cPM.Pause 0.02, 1
    'Capture a new checkpoint after clearing
        cPM.Checkpoint "Phase 2"

'------------------------------------------------------------------------------
' ASSERT REUSED SESSION OUTPUT
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

    'Assert one captured checkpoint after reuse
        Test_Assert_EqualLong 1, cPM.CheckpointCount, _
                              "CheckpointCount after ClearCheckpoints and reuse"
    'Assert that the surviving exported checkpoint is the new one
        Test_Assert_EqualString "Phase 2", CStr(Arr(2, 3)), _
                                "ReportAsArray contains only the new checkpoint after reuse"
    'Assert that the exported sequence restarts at 1 after clearing
        Test_Assert_EqualLong 1, Arr(2, 2), _
                              "Checkpoint sequence restarts at 1 after ClearCheckpoints"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_ClearCheckpoints_ThenReuseSession"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_SetRunLabel_AfterClearCheckpoints()
'
'==============================================================================
'                TEST SETRUNLABEL AFTER CLEARCHECKPOINTS
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that SetRunLabel is usable again after ClearCheckpoints
'
' WHY THIS EXISTS
'   ClearCheckpoints resets structured reporting state, so the run label should
'   again be assignable before the next checkpoint is captured
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Arr                 As Variant                'Structured report export

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "SetRunLabel after ClearCheckpoints"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' PREPARE INITIAL SESSION STATE
'------------------------------------------------------------------------------
    'Start a fresh timing session
        cPM.StartTimer 5, False
    'Assign the initial run label
        cPM.SetRunLabel "Run A"
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture one checkpoint under the initial run label
        cPM.Checkpoint "Phase 1"

'------------------------------------------------------------------------------
' CLEAR CHECKPOINT STATE
'------------------------------------------------------------------------------
    'Clear the structured checkpoint/report state
        cPM.ClearCheckpoints

'------------------------------------------------------------------------------
' ASSIGN NEW RUN LABEL
'------------------------------------------------------------------------------
    'Assign a new run label after clearing checkpoint state
        cPM.SetRunLabel "Run B"

'------------------------------------------------------------------------------
' CAPTURE NEW CHECKPOINT
'------------------------------------------------------------------------------
    'Perform a short pause
        cPM.Pause 0.02, 1
    'Capture one checkpoint under the new run label
        cPM.Checkpoint "Phase 2"

'------------------------------------------------------------------------------
' ASSERT NEW RUN-LABEL EXPORT
'------------------------------------------------------------------------------
    'Export the structured checkpoint report
        Arr = cPM.ReportAsArray

    'Assert that the current run label is updated
        Test_Assert_EqualString "Run B", cPM.RunLabel, _
                                "RunLabel can be reassigned after ClearCheckpoints"
    'Assert one captured checkpoint after reassignment
        Test_Assert_EqualLong 1, cPM.CheckpointCount, _
                              "CheckpointCount after SetRunLabel following ClearCheckpoints"
    'Assert that the exported run label matches the reassigned value
        Test_Assert_EqualString "Run B", CStr(Arr(2, 1)), _
                                "ReportAsArray exports the reassigned run label after ClearCheckpoints"
    'Assert that the exported checkpoint belongs to the second phase
        Test_Assert_EqualString "Phase 2", CStr(Arr(2, 3)), _
                                "ReportAsArray exports the new checkpoint after reassigned run label"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_SetRunLabel_AfterClearCheckpoints"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_TW_Turn_ON_WhenInactive()
'
'==============================================================================
'                    TEST TW TURN ON WHEN INACTIVE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates idempotent TW_Turn_ON behavior when no TW session is active
'
' WHY THIS EXISTS
'   The TW lifecycle surface should remain safe when callers attempt to end TW
'   participation that does not currently exist
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim Baseline            As T_AppState             'Captured Application baseline

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "TW_Turn_ON when inactive"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions
    'Create a fresh class instance
        Set cPM = New cPerformanceManager
    'Capture the current Application baseline
        CaptureAppState Baseline

'------------------------------------------------------------------------------
' APPLY NO-OP TW TURN ON
'------------------------------------------------------------------------------
    'Attempt to end TW participation when no TW session is active
        cPM.TW_Turn_ON

'------------------------------------------------------------------------------
' ASSERT IDLE STATE
'------------------------------------------------------------------------------
    'Assert that the instance remains inactive
        Test_Assert_EqualBoolean False, cPM.TW_IsActive, _
                                 "TW_IsActive remains False when TW_Turn_ON is called while inactive"
    'Assert that the shared TW manager remains idle
        Test_Assert_EqualLong 0, cPM.TW_ActiveSessionCount, _
                              "TW_ActiveSessionCount remains 0 when TW_Turn_ON is called while inactive"

'------------------------------------------------------------------------------
' ASSERT APPLICATION STATE UNCHANGED
'------------------------------------------------------------------------------
    'Assert that Application state remains unchanged
        Test_Assert_EqualBoolean Baseline.ScreenUpdating, Application.ScreenUpdating, _
                                 "ScreenUpdating remains unchanged when TW_Turn_ON is called while inactive"
        Test_Assert_EqualBoolean Baseline.EnableEvents, Application.EnableEvents, _
                                 "EnableEvents remains unchanged when TW_Turn_ON is called while inactive"
        Test_Assert_EqualBoolean Baseline.DisplayAlerts, Application.DisplayAlerts, _
                                 "DisplayAlerts remains unchanged when TW_Turn_ON is called while inactive"
        Test_Assert_EqualLong Baseline.Calculation, Application.Calculation, _
                              "Calculation remains unchanged when TW_Turn_ON is called while inactive"
        Test_Assert_EqualLong Baseline.Cursor, Application.Cursor, _
                              "Cursor remains unchanged when TW_Turn_ON is called while inactive"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If

    'Force the shared TW manager to a clean baseline
        PM_TW_EndAllSessions
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_TW_Turn_ON_WhenInactive"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub


Private Sub Test_Pause_Boundary_NoOpBehavior()
'
'==============================================================================
'                   TEST PAUSE BOUNDARY NO-OP BEHAVIOR
'------------------------------------------------------------------------------
' PURPOSE
'   Validates no-op boundary behavior for Pause
'
' WHY THIS EXISTS
'   The class documents that non-positive pause durations and excessively large
'   pause durations are treated as no-op exits
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Class under test
    Dim ElapsedZero         As Double                 'Measured elapsed time for Pause 0
    Dim ElapsedNegative     As Double                 'Measured elapsed time for negative pause
    Dim ElapsedTooLarge     As Double                 'Measured elapsed time for oversized pause

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start the regression case
        Case_Begin "Pause no-op boundary behavior"
    'Enable case-level unexpected-error handling
        On Error GoTo CleanFail
    'Create a fresh class instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' ASSERT PAUSE(0)
'------------------------------------------------------------------------------
    'Start QPC timing for the zero-duration pause
        cPM.StartTimer 5, False
    'Invoke Pause with zero seconds
        cPM.Pause 0, 1
    'Measure elapsed time
        ElapsedZero = cPM.ElapsedSeconds(5)

    'Assert near-no-op elapsed range for Pause(0)
        Test_Assert_InRangeDouble 0#, 0.25, ElapsedZero, _
                                  "Pause(0) behaves as a no-op"

'------------------------------------------------------------------------------
' ASSERT PAUSE(-1)
'------------------------------------------------------------------------------
    'Start QPC timing for the negative-duration pause
        cPM.StartTimer 5, False
    'Invoke Pause with a negative duration
        cPM.Pause -1, 1
    'Measure elapsed time
        ElapsedNegative = cPM.ElapsedSeconds(5)

    'Assert near-no-op elapsed range for negative Pause
        Test_Assert_InRangeDouble 0#, 0.25, ElapsedNegative, _
                                  "Pause(-1) behaves as a no-op"

'------------------------------------------------------------------------------
' ASSERT PAUSE(> 3600)
'------------------------------------------------------------------------------
    'Start QPC timing for the oversized-duration pause
        cPM.StartTimer 5, False
    'Invoke Pause with a duration beyond the allowed upper bound
        cPM.Pause 3601, 1
    'Measure elapsed time
        ElapsedTooLarge = cPM.ElapsedSeconds(5)

    'Assert near-no-op elapsed range for oversized Pause
        Test_Assert_InRangeDouble 0#, 0.25, ElapsedTooLarge, _
                                  "Pause(3601) behaves as a no-op"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes held by the instance on a best-effort basis
        On Error Resume Next
        If Not cPM Is Nothing Then
            cPM.ResetEnvironment
            Set cPM = Nothing
        End If
        On Error GoTo 0

    'Finalize the current case
        Case_Finalize

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Record the unexpected case-level error
        RecordUnexpectedError "Test_Pause_Boundary_NoOpBehavior"
    'Continue through centralized cleanup
        Resume CleanExit

End Sub
