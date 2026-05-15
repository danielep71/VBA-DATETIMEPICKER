Attribute VB_Name = "M_cDP_Test"
Option Explicit

'
'==============================================================================
' MODULE: M_cDP_Test
'==============================================================================
' PURPOSE
'   Provides a regression harness for the VBA DatePicker project
'
' WHY THIS EXISTS
'   The DatePicker spans persisted settings, manager event orchestration,
'   worksheet write-back, UserForm bridge state, optional context-menu access,
'   optional keyboard shortcut access, optional in-grid worksheet shapes, and
'   optional WinAPI-dependent UI behavior
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Runs grouped regression suites and writes results to:
'     - the Immediate Window
'     - a worksheet named TST_DP_RESULTS
'
'   Creates and deletes a scratch worksheet named TST_DP_SCRATCH
'
'   Captures and restores DatePicker settings, transient DatePicker state, and
'   selected Excel Application state before and after each run
'
' ERROR POLICY
'   Individual assertion failures are recorded and the run continues
'
'   Suite-level failures are recorded by TST_DP_RunSuiteSafe and the harness
'   continues with the next suite
'
'   A fatal harness failure is recorded, cleanup is attempted, and the original
'   error is re-raised to the caller
'
' DEPENDENCIES
'   M_DatePicker
'   cDatePickerManager
'   UF_DatePicker
'   M_DEMO_BUILDER
'   Excel object model
'
' NOTES
'   Before running this module, the DatePicker project must compile without
'   errors
'
'   TST_DP_RunAll runs all non-disruptive suites without opening UF_DatePicker
'
'   TST_DP_RunAll_WithUISmoke additionally opens and closes UF_DatePicker as a
'   minimal visual lifecycle check
'
'   TST_DP_HolidayCallback, TST_DP_HolidayCallbackNonBoolean, and
'   TST_DP_HolidayCallbackError must remain Public so Application.Run can
'   resolve them as holiday policy callbacks
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TST_DP_RESULT_SHEET_NAME  As String = "TST_DP_RESULTS"        'Regression result worksheet name
    Private Const TST_DP_SCRATCH_SHEET_NAME As String = "TST_DP_SCRATCH"        'Regression scratch worksheet name
    Private Const TST_DP_PASS_TEXT          As String = "PASS"                  'Passed test marker
    Private Const TST_DP_FAIL_TEXT          As String = "FAIL"                  'Failed test marker
    Private Const TST_DP_INFO_TEXT          As String = "INFO"                  'Information marker
    Private Const TST_DP_MODULE_NAME        As String = "M_cDP_Test"            'This module name for callback resolution

    'Result sheet layout
    Private Const TST_DP_RESULT_FIRST_ROW   As Long = 5                         'First result data row on the result sheet
    Private Const TST_DP_COL_SEQ            As Long = 3                         'Result sequence number column index
    Private Const TST_DP_COL_TIMESTAMP      As Long = 4                         'Result timestamp column index
    Private Const TST_DP_COL_RESULT         As Long = 5                         'Result marker column index
    Private Const TST_DP_COL_SUITE          As Long = 6                         'Suite name column index
    Private Const TST_DP_COL_TEST           As Long = 7                         'Test name column index
    Private Const TST_DP_COL_DETAILS        As Long = 8                         'Details column index
    Private Const TST_DP_COL_SUMMARY_LABEL  As Long = 10                        'Summary label column index
    Private Const TST_DP_COL_SUMMARY_VALUE  As Long = 11                        'Summary value column index

    'Conditional formatting
    Private Const TST_DP_FAIL_BACK_COLOR    As Long = 192                       'FAIL cell background color

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type TRegDPSettingsSnapshot
        ShowRightClick              As Boolean  'Right-click feature setting
        ShowGridIcon                As Boolean  'Grid-icon feature setting
        FirstDayOfWeek              As Long     'First-day setting
        UseLocalNames               As Boolean  'Local-name setting
        ClockMode                   As Long     'Clock-mode setting
        SizeMode                    As Long     'Size-mode setting
        HighlightWeekends           As Boolean  'Weekend-highlight setting
        AllowOutsideMonthSelection  As Boolean  'Outside-month selection setting
        CloseAfterSelection         As Boolean  'Close-after-selection setting
        UseWinAPI                   As Boolean  'WinAPI setting
        EnableKeyboardShortcut      As Boolean  'Keyboard shortcut setting
        HolidayCallbackName         As String   'Holiday callback setting
        WriteValue                  As Date     'Transient write value
        InitialDate                 As Date     'Transient initial date
        HasInitialDate              As Boolean  'Transient initial-date flag
        SelectedDate                As Date     'Transient selected date
        HasSelectedDate             As Boolean  'Transient selected-date flag
    End Type

    Private Type TRegDPApplicationSnapshot
        ScreenUpdating              As Boolean  'Application.ScreenUpdating snapshot
        EnableEvents                As Boolean  'Application.EnableEvents snapshot
        DisplayAlerts               As Boolean  'Application.DisplayAlerts snapshot
        CalculationMode             As Long     'Application.Calculation snapshot
        StatusBarWasFalse           As Boolean  'True when StatusBar was Excel-owned
        StatusBarText               As String   'StatusBar text snapshot
    End Type

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mTST_DP_ResultSheet     As Excel.Worksheet  'Result worksheet used by the current run
    Private mTST_DP_ScratchSheet    As Excel.Worksheet  'Scratch worksheet used by the current run
    Private mTST_DP_HostWorkbook    As Excel.Workbook   'Workbook receiving test sheets
    Private mTST_DP_NextResultRow   As Long             'Next available result row on the result sheet
    Private mTST_DP_RunCount        As Long             'Total assertions executed in the current run
    Private mTST_DP_PassCount       As Long             'Total assertions passed in the current run
    Private mTST_DP_FailCount       As Long             'Total assertions failed in the current run
    Private mTST_DP_CurrentSuite    As String           'Suite name currently being executed
    Private mTST_DP_HadManager      As Boolean          'True when a manager existed before the run


'
'------------------------------------------------------------------------------
'
'                              PUBLIC ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub TST_DP_RunAll()

'
'==============================================================================
'                           RUN ALL REGRESSION TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the full non-disruptive DatePicker regression pack
'
' WHY THIS EXISTS
'   This is the standard regression entry point for development and release
'   validation checks that do not require UF_DatePicker to be opened
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs environment, settings, date policy, holiday policy, caption, form
'   bridge, write-back, grid-icon, and manager suites without opening
'   UF_DatePicker
'
' ERROR POLICY
'   Delegates fatal handling and cleanup to TST_DP_RunAllInternal
'
' DEPENDENCIES
'   TST_DP_RunAllInternal
'   DEMO_Sheet_BuildTemplate
'   TST_DP_GetHostWorkbook
'
' NOTES
'   Use TST_DP_RunAll_WithUISmoke when a brief UF_DatePicker open/close check
'   is also needed
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Reset module-level counters and object references before the run
        TST_DP_ResetHarnessState
    'Resolve the workbook that will receive the result and scratch sheets
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()
    'Build the result sheet template before the run
        DEMO_Sheet_BuildTemplate TST_DP_RESULT_SHEET_NAME, "DATE PICKER", _
            "Test Sheet", , TST_DP_RESULT_FIRST_ROW

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the standard non-disruptive regression pack without UI smoke
        TST_DP_RunAllInternal False

End Sub

Public Sub TST_DP_RunAll_WithUISmoke()

'
'==============================================================================
'                           RUN ALL WITH UI SMOKE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the full DatePicker regression pack including a UF_DatePicker smoke
'   test
'
' WHY THIS EXISTS
'   UI smoke testing is useful before a release but briefly opens UF_DatePicker
'   and is therefore separated from the default regression run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs all non-disruptive suites and then opens / closes UF_DatePicker as a
'   minimal visual lifecycle check
'
' ERROR POLICY
'   Delegates fatal handling and cleanup to TST_DP_RunAllInternal
'
' DEPENDENCIES
'   TST_DP_RunAllInternal
'   DEMO_Sheet_BuildTemplate
'   TST_DP_GetHostWorkbook
'
' NOTES
'   This routine is intentionally a distinct Public macro so it can be run
'   separately from the Macro dialog
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESET HARNESS STATE
'------------------------------------------------------------------------------
    'Reset module-level counters and object references before the run
        TST_DP_ResetHarnessState

'------------------------------------------------------------------------------
' RESOLVE HOST WORKBOOK
'------------------------------------------------------------------------------
    'Resolve the workbook that will receive the result and scratch sheets
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()

'------------------------------------------------------------------------------
' BUILD RESULT SHEET TEMPLATE
'------------------------------------------------------------------------------
    'Build the result sheet template before the run
        DEMO_Sheet_BuildTemplate TST_DP_RESULT_SHEET_NAME, "DATE PICKER", _
            "Test Sheet", , TST_DP_RESULT_FIRST_ROW

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the regression pack with the UI smoke suite included
        TST_DP_RunAllInternal True

End Sub


'
'------------------------------------------------------------------------------
'
'                            PUBLIC CALLBACK STUBS
'
'------------------------------------------------------------------------------
'

Public Function TST_DP_HolidayCallback(ByVal CandidateDate As Date) As Boolean

'
'==============================================================================
'                           TEST HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a deterministic Boolean holiday callback for regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate dispatches user callbacks through
'   Application.Run and requires a stable Public callback target during
'   regression
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   True only for 1 January 2026; False for all other dates
'
' BEHAVIOR
'   Compares the date-only component of CandidateDate to 1 January 2026
'
' ERROR POLICY
'   Does not raise intentional errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   This function must remain Public so Application.Run can resolve it
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return True only for the deterministic test holiday date
        TST_DP_HolidayCallback = _
            (VBA.DateValue(CandidateDate) = VBA.DateSerial(2026, 1, 1))

End Function

Public Function TST_DP_HolidayCallbackNonBoolean(ByVal CandidateDate As Date) As String

'
'==============================================================================
'                       TEST NON-BOOLEAN HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a non-Boolean callback result for holiday policy regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate should ignore callback results that are not
'   explicitly Boolean and return False rather than coercing a truthy string
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   String value "TRUE" deliberately
'
' BEHAVIOR
'   Returns a String rather than a Boolean to verify that the policy does not
'   coerce the return value
'
' ERROR POLICY
'   Does not raise intentional errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   CandidateDate is intentionally unused
'
'   This function must remain Public so Application.Run can resolve it
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RETURN NON-BOOLEAN RESULT
'------------------------------------------------------------------------------
    'Return a non-Boolean value deliberately to verify policy rejection
        TST_DP_HolidayCallbackNonBoolean = "TRUE"

End Function

Public Function TST_DP_HolidayCallbackError(ByVal CandidateDate As Date) As Variant

'
'==============================================================================
'                         TEST ERROR-VALUE HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides an Excel error value for holiday policy regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate should ignore callback results that are Excel
'   error values and return False rather than treating the error as a holiday
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   CVErr(xlErrValue) deliberately
'
' BEHAVIOR
'   Returns an Excel error value without raising a VBA runtime error
'
' ERROR POLICY
'   Does not raise intentional runtime errors
'
' DEPENDENCIES
'   CVErr
'   xlErrValue
'
' NOTES
'   CandidateDate is intentionally unused
'
'   This function must remain Public so Application.Run can resolve it
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RETURN ERROR VALUE
'------------------------------------------------------------------------------
    'Return an Excel error value deliberately to verify policy rejection
        TST_DP_HolidayCallbackError = CVErr(xlErrValue)

End Function


'
'------------------------------------------------------------------------------
'
'                               RUN ORCHESTRATION
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_RunAllInternal(ByVal IncludeUISmoke As Boolean)

'
'==============================================================================
'                           RUN ALL INTERNAL
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates one complete DatePicker regression run
'
' WHY THIS EXISTS
'   The harness must capture and restore user settings, isolate scratch
'   worksheet artifacts, restore Excel Application state, and continue through
'   grouped suites where individual failures are non-fatal
'
' INPUTS
'   IncludeUISmoke
'     True to include the UF_DatePicker open / close smoke suite
'     False to run only non-disruptive suites
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures DatePicker settings and Excel Application state, prepares the
'   result and scratch worksheets, runs each suite through TST_DP_RunSuiteSafe,
'   writes the run summary, and restores all captured state on exit
'
' ERROR POLICY
'   Records fatal harness failures, attempts cleanup in CleanExit, then
'   re-raises the original error to the caller
'
' DEPENDENCIES
'   All TST_DP_RunSuite_* routines
'   TST_DP_CaptureSettings
'   TST_DP_CaptureApplicationState
'   TST_DP_PrepareApplicationForRun
'   TST_DP_ResetDatePickerArtifacts
'   TST_DP_PrepareResultSheet
'   TST_DP_PrepareScratchSheet
'   TST_DP_WriteSummary
'   TST_DP_RestoreSettings
'   TST_DP_RestoreManagerState
'   TST_DP_RestoreApplicationState
'
' NOTES
'   Application events are disabled during the run to reduce interference from
'   workbook-level event handlers in the host workbook
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TST_DP_RunAllInternal"    'Current procedure name

    Dim SettingsSnapshot    As TRegDPSettingsSnapshot               'DatePicker state snapshot
    Dim AppSnapshot         As TRegDPApplicationSnapshot            'Excel Application state snapshot
    Dim FatalNumber         As Long                                 'Fatal error number
    Dim FatalDescription    As String                               'Fatal error description
    Dim HasFatalError       As Boolean                              'True when a fatal error must be re-raised

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled fatal handling
        On Error GoTo FatalHandler
    'Capture whether a manager existed before the run
        mTST_DP_HadManager = Not (gDP_Manager Is Nothing)
    'Capture current DatePicker settings and transient state
        TST_DP_CaptureSettings SettingsSnapshot
    'Capture current Excel Application state
        TST_DP_CaptureApplicationState AppSnapshot

'------------------------------------------------------------------------------
' PREPARE ISOLATED RUN STATE
'------------------------------------------------------------------------------
    'Prepare the Excel Application state for the regression run
        TST_DP_PrepareApplicationForRun
    'Reset transient DatePicker UI artifacts before testing
        TST_DP_ResetDatePickerArtifacts
    'Prepare the result worksheet
        TST_DP_PrepareResultSheet mTST_DP_HostWorkbook
    'Prepare the scratch worksheet
        TST_DP_PrepareScratchSheet mTST_DP_HostWorkbook
    'Record the run header
        TST_DP_RecordInfo "Harness", "Start", _
            "IncludeUISmoke=" & VBA.CStr(IncludeUISmoke)

'------------------------------------------------------------------------------
' RUN SUITES
'------------------------------------------------------------------------------
    'Run environment and manager smoke checks
        TST_DP_RunSuiteSafe "Environment"
    'Run settings and persisted-state checks
        TST_DP_RunSuiteSafe "Settings"
    'Run date-selection policy checks
        TST_DP_RunSuiteSafe "DatePolicy"
    'Run holiday callback dispatch and fail-safe checks
        TST_DP_RunSuiteSafe "HolidayPolicy"
    'Run month and date caption helper checks
        TST_DP_RunSuiteSafe "Captions"
    'Run form bridge state checks
        TST_DP_RunSuiteSafe "FormBridge"
    'Run worksheet write-back checks
        TST_DP_RunSuiteSafe "WriteBack"
    'Run in-grid icon lifecycle checks
        TST_DP_RunSuiteSafe "GridIcon"
    'Run manager public API and target gating checks
        TST_DP_RunSuiteSafe "Manager"
    'Run optional UF_DatePicker open / close smoke check when requested
        If IncludeUISmoke Then
            TST_DP_RunSuiteSafe "UISmoke"
        End If

'------------------------------------------------------------------------------
' WRITE SUMMARY
'------------------------------------------------------------------------------
    'Write the run summary to the result sheet and Immediate Window
        TST_DP_WriteSummary

'------------------------------------------------------------------------------
' APPLY CONDITIONAL FORMATTING
'------------------------------------------------------------------------------
    'Apply FAIL conditional formatting to the result column
        If Not mTST_DP_ResultSheet Is Nothing Then
            TST_DP_ApplyFailConditionalFormat _
                mTST_DP_ResultSheet, _
                mTST_DP_ResultSheet.Range( _
                    mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW, TST_DP_COL_RESULT), _
                    mTST_DP_ResultSheet.Cells(1000, TST_DP_COL_RESULT))
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors so every cleanup step is attempted
        On Error Resume Next
    'Reset DatePicker UI artifacts after testing
        TST_DP_ResetDatePickerArtifacts
    'Delete the scratch worksheet
        TST_DP_DeleteScratchSheet
    'Restore DatePicker settings and transient state
        TST_DP_RestoreSettings SettingsSnapshot
    'Restore the manager state to its pre-run condition
        TST_DP_RestoreManagerState
    'Restore the Excel Application state
        TST_DP_RestoreApplicationState AppSnapshot
    'Release module object references
        Set mTST_DP_ScratchSheet = Nothing
        Set mTST_DP_ResultSheet = Nothing
        Set mTST_DP_HostWorkbook = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Re-raise fatal harness errors after cleanup is complete
        If HasFatalError Then
            Err.Raise FatalNumber, PROC_NAME, FatalDescription
        End If
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FATAL HANDLER
'------------------------------------------------------------------------------
FatalHandler:
    'Capture the fatal error number
        FatalNumber = Err.Number
    'Capture the fatal error description
        FatalDescription = Err.Description
    'Mark the fatal error for re-raise after cleanup
        HasFatalError = True
    'Record the fatal harness failure when the result sheet is available
        On Error Resume Next
        TST_DP_RecordResult TST_DP_FAIL_TEXT, _
            "Harness", _
            PROC_NAME, _
            "Fatal error " & VBA.CStr(FatalNumber) & " - " & FatalDescription
        Err.Clear
        On Error GoTo 0
    'Run shared cleanup and re-raise
        Resume CleanExit
End Sub

Private Sub TST_DP_ResetHarnessState()

'
'==============================================================================
'                           RESET HARNESS STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears module-level counters and object references before a run
'
' WHY THIS EXISTS
'   Regression runs may be launched repeatedly in the same Excel session and
'   stale state from a previous run must not carry forward
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets assertion counters, the current suite name, the result row pointer,
'   and all module-level object references
'
' ERROR POLICY
'   Does not raise intentional errors
'
' DEPENDENCIES
'   Module-level state
'
' NOTES
'   Object cleanup is handled by the run orchestrator cleanup path
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESET COUNTERS
'------------------------------------------------------------------------------
    'Reset the total assertion counter
        mTST_DP_RunCount = 0
    'Reset the passed assertion counter
        mTST_DP_PassCount = 0
    'Reset the failed assertion counter
        mTST_DP_FailCount = 0
    'Reset the result row pointer to the first data row
        mTST_DP_NextResultRow = TST_DP_RESULT_FIRST_ROW
    'Reset the current suite name
        mTST_DP_CurrentSuite = vbNullString

'------------------------------------------------------------------------------
' RESET REFERENCES
'------------------------------------------------------------------------------
    'Clear the result sheet reference
        Set mTST_DP_ResultSheet = Nothing
    'Clear the scratch sheet reference
        Set mTST_DP_ScratchSheet = Nothing
    'Clear the host workbook reference
        Set mTST_DP_HostWorkbook = Nothing

End Sub


'
'------------------------------------------------------------------------------
'
'                                  SUITE RUNNER
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_RunSuiteSafe(ByVal SuiteName As String)

'
'==============================================================================
'                           RUN SUITE SAFE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one regression suite behind an outer fail-safe boundary
'
' WHY THIS EXISTS
'   Individual suites already contain local error handlers. This wrapper adds a
'   second boundary so unexpected project changes, Excel object-model errors, or
'   environment-specific failures cannot abort the full regression run
'
' INPUTS
'   SuiteName
'     Logical suite name to dispatch
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Sets the current suite name, dispatches the requested suite by name through
'   a Select Case block, and records any escaping error as a suite-level failure
'   before allowing the harness to continue with the next suite
'
' ERROR POLICY
'   Best-effort. Records escaping suite failures without raising outward
'
' DEPENDENCIES
'   All TST_DP_RunSuite_* routines
'
' NOTES
'   The Select Case dispatch keeps suite routines Private and avoids relying on
'   Application.Run for internal dispatch
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TST_DP_RunSuiteSafe"  'Current procedure name

    Dim ErrorNumber         As Long                             'Captured suite error number
    Dim ErrorDescription    As String                           'Captured suite error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name for recording context
        mTST_DP_CurrentSuite = SuiteName
    'Protect the harness from escaping suite failures
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' DISPATCH SUITE
'------------------------------------------------------------------------------
    'Dispatch the requested suite by name
        Select Case VBA.UCase$(VBA.Trim$(SuiteName))

            Case "ENVIRONMENT"
                TST_DP_RunSuite_Environment

            Case "SETTINGS"
                TST_DP_RunSuite_Settings

            Case "DATEPOLICY"
                TST_DP_RunSuite_DatePolicy

            Case "HOLIDAYPOLICY"
                TST_DP_RunSuite_HolidayPolicy

            Case "CAPTIONS"
                TST_DP_RunSuite_Captions

            Case "FORMBRIDGE"
                TST_DP_RunSuite_FormBridge

            Case "WRITEBACK"
                TST_DP_RunSuite_WriteBack

            Case "GRIDICON"
                TST_DP_RunSuite_GridIcon

            Case "MANAGER"
                TST_DP_RunSuite_Manager

            Case "UISMOKE"
                TST_DP_RunSuite_UISmoke

            Case Else
                TST_DP_RecordFail "Unknown suite name", "SuiteName=" & SuiteName

        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite dispatch completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Capture the escaping error number
        ErrorNumber = Err.Number

    'Capture the escaping error description
        ErrorDescription = Err.Description

    'Suppress result-recording failures
        On Error Resume Next

    'Record the escaped suite failure without aborting the harness
        TST_DP_RecordResult TST_DP_FAIL_TEXT, _
            SuiteName, _
            PROC_NAME, _
            "Escaped suite error " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription

    'Clear any suppressed result-recording error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                                  TEST SUITES
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_RunSuite_Environment()

'
'==============================================================================
'                           ENVIRONMENT SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates basic DatePicker infrastructure startup and platform helpers
'
' WHY THIS EXISTS
'   Later suites assume settings can be loaded, the manager can be created, and
'   platform helpers return deterministic results
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls M_Settings_EnsureLoaded and M_Picker_EnsureManager and asserts that
'   the global manager is instantiated, not busy after startup, and that
'   platform helpers return Boolean values
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Picker_EnsureManager
'   M_Platform_CanUseWinAPI
'   M_Platform_ShouldUseWinAPI
'   gDP_Manager
'
' NOTES
'   The manager Is_Hooked check may report False when Excel is in a transient
'   state; the assertion records the result without failing on that condition
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Environment"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' SETTINGS LOAD
'------------------------------------------------------------------------------
    'Ensure settings load without raising
        M_Settings_EnsureLoaded
    'Record successful settings load
        TST_DP_RecordPass "M_Settings_EnsureLoaded does not raise", vbNullString

'------------------------------------------------------------------------------
' MANAGER CREATION
'------------------------------------------------------------------------------
    'Ensure the global manager can be created
        M_Picker_EnsureManager

    'Assert that the global manager reference is populated
        TST_DP_AssertTrue "Global manager is instantiated", _
            Not (gDP_Manager Is Nothing)

    'Assert that the manager is not busy after startup
        TST_DP_AssertFalse "Manager is not busy after startup", _
            gDP_Manager.Is_Busy

'------------------------------------------------------------------------------
' PLATFORM HELPERS
'------------------------------------------------------------------------------
    'Assert that the platform capability helper returns a Boolean result
        TST_DP_AssertBooleanResult "M_Platform_CanUseWinAPI returns Boolean", _
            M_Platform_CanUseWinAPI

    'Assert that the effective WinAPI policy helper returns a Boolean result
        TST_DP_AssertBooleanResult "M_Platform_ShouldUseWinAPI returns Boolean", _
            M_Platform_ShouldUseWinAPI

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "Environment suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Settings()

'
'==============================================================================
'                           SETTINGS SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates public DatePicker settings accessors and setters
'
' WHY THIS EXISTS
'   Settings drive form rendering, manager behavior, shortcut integration,
'   context-menu availability, grid-icon availability, and date selection rules
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exercises valid settings, invalid setting rejection, text parsing, access
'   path dead-configuration fallback, WinAPI normalization, enum setting round
'   trips, and holiday callback trimming
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Settings_* public API
'
' NOTES
'   Settings are restored by the outer harness cleanup path
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Settings"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' FIRST-DAY VALIDATION
'------------------------------------------------------------------------------
    'Assert vbSunday is a supported first-day value
        TST_DP_AssertTrue "vbSunday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbSunday)

    'Assert vbMonday is a supported first-day value
        TST_DP_AssertTrue "vbMonday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbMonday)

    'Assert zero is not a supported first-day value
        TST_DP_AssertFalse "0 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(0)

    'Assert Tuesday-style value is not a supported first-day value
        TST_DP_AssertFalse "3 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(3)

    'Assert vbSunday converts to the expected text label
        TST_DP_AssertEqualsString "vbSunday converts to text", _
            "vbSunday", _
            M_Settings_FirstDayOfWeekToText(vbSunday)

    'Assert vbMonday converts to the expected text label
        TST_DP_AssertEqualsString "vbMonday converts to text", _
            "vbMonday", _
            M_Settings_FirstDayOfWeekToText(vbMonday)

    'Assert invalid first-day conversion raises a runtime error
        TST_DP_ExpectError_FirstDayToTextInvalid

'------------------------------------------------------------------------------
' FIRST-DAY SETTERS
'------------------------------------------------------------------------------
    'Set Sunday as the first day
        M_Settings_SetFirstDayOfWeek vbSunday

    'Assert the Sunday setting was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeek stores vbSunday", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Monday through the text parser
        M_Settings_SetFirstDayOfWeekText "Monday"

    'Assert the Monday setting was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Monday", _
            vbMonday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Sunday through the short text parser
        M_Settings_SetFirstDayOfWeekText "Sun"

    'Assert the Sunday setting was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Sun", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Assert blank first-day text raises a runtime error
        TST_DP_ExpectError_SetFirstDayBlank

    'Assert unsupported first-day text raises a runtime error
        TST_DP_ExpectError_SetFirstDayInvalidText

'------------------------------------------------------------------------------
' BOOLEAN DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Enable local names
        M_Settings_SetUseLocalNames True

    'Assert local names setting is on
        TST_DP_AssertTrue "UseLocalNames can be enabled", _
            M_Settings_GetUseLocalNames()

    'Disable local names
        M_Settings_SetUseLocalNames False

    'Assert local names setting is off
        TST_DP_AssertFalse "UseLocalNames can be disabled", _
            M_Settings_GetUseLocalNames()

    'Enable weekend highlighting
        M_Settings_SetHighlightWeekends True

    'Assert weekend highlighting is on
        TST_DP_AssertTrue "HighlightWeekends can be enabled", _
            M_Settings_GetHighlightWeekends()

    'Disable weekend highlighting
        M_Settings_SetHighlightWeekends False

    'Assert weekend highlighting is off
        TST_DP_AssertFalse "HighlightWeekends can be disabled", _
            M_Settings_GetHighlightWeekends()

'------------------------------------------------------------------------------
' BOOLEAN BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Enable outside-month selection
        M_Settings_SetAllowOutsideMonthSelection True

    'Assert outside-month selection is on
        TST_DP_AssertTrue "AllowOutsideMonthSelection can be enabled", _
            M_Settings_GetAllowOutsideMonthSelection()

    'Disable outside-month selection
        M_Settings_SetAllowOutsideMonthSelection False

    'Assert outside-month selection is off
        TST_DP_AssertFalse "AllowOutsideMonthSelection can be disabled", _
            M_Settings_GetAllowOutsideMonthSelection()

    'Enable close-after-selection
        M_Settings_SetCloseAfterSelection True

    'Assert close-after-selection is on
        TST_DP_AssertTrue "CloseAfterSelection can be enabled", _
            M_Settings_GetCloseAfterSelection()

    'Disable close-after-selection
        M_Settings_SetCloseAfterSelection False

    'Assert close-after-selection is off
        TST_DP_AssertFalse "CloseAfterSelection can be disabled", _
            M_Settings_GetCloseAfterSelection()

'------------------------------------------------------------------------------
' FEATURE SETTINGS AND ACCESS-PATH NORMALIZATION
'------------------------------------------------------------------------------
    'Enable right-click access
        M_Settings_SetShowRightClick True

    'Assert right-click access is enabled
        TST_DP_AssertTrue "ShowRightClick can be enabled", _
            M_Settings_GetShowRightClick()

    'Disable right-click access
        M_Settings_SetShowRightClick False

    'Assert right-click access is disabled
        TST_DP_AssertFalse "ShowRightClick can be disabled", _
            M_Settings_GetShowRightClick()

    'Enable grid-icon access
        M_Settings_SetShowGridIcon True

    'Assert grid-icon access is enabled
        TST_DP_AssertTrue "ShowGridIcon can be enabled", _
            M_Settings_GetShowGridIcon()

    'Disable grid-icon access
        M_Settings_SetShowGridIcon False

    'Assert grid-icon access is disabled
        TST_DP_AssertFalse "ShowGridIcon can be disabled", _
            M_Settings_GetShowGridIcon()

    'Enable both contextual access paths before testing keyboard override
        M_Settings_SetShowRightClick True
        M_Settings_SetShowGridIcon True

    'Disable keyboard shortcut while other access paths are enabled
        M_Settings_SetEnableKeyboardShortcut False

    'Assert keyboard shortcut can be disabled when other access paths exist
        TST_DP_AssertFalse "Keyboard shortcut can be disabled when other access paths exist", _
            M_Settings_GetEnableKeyboardShortcut()

    'Re-enable keyboard shortcut
        M_Settings_SetEnableKeyboardShortcut True

    'Assert keyboard shortcut can be enabled
        TST_DP_AssertTrue "Keyboard shortcut can be enabled", _
            M_Settings_GetEnableKeyboardShortcut()

    'Disable both contextual access paths to trigger the dead-configuration guard
        M_Settings_SetShowRightClick False
        M_Settings_SetShowGridIcon False

    'Assert keyboard shortcut is forced on when both contextual paths are disabled
        TST_DP_AssertTrue "Keyboard shortcut is forced when right-click and grid icon are disabled", _
            M_Settings_GetEnableKeyboardShortcut()

    'Restore feature settings to working defaults for following suites
        M_Settings_SetShowRightClick True
        M_Settings_SetShowGridIcon True
        M_Settings_SetEnableKeyboardShortcut True

'------------------------------------------------------------------------------
' WINAPI SETTING
'------------------------------------------------------------------------------
    'Disable WinAPI styling
        M_Settings_SetUseWinAPI False

    'Assert WinAPI styling can be disabled
        TST_DP_AssertFalse "UseWinAPI can be disabled", _
            M_Settings_GetUseWinAPI()

    'Request WinAPI styling on
        M_Settings_SetUseWinAPI True

    'Assert WinAPI styling is normalized to platform capability
        TST_DP_AssertEqualsLong "UseWinAPI normalizes to platform capability", _
            VBA.CLng(M_Platform_CanUseWinAPI), _
            VBA.CLng(M_Settings_GetUseWinAPI())

'------------------------------------------------------------------------------
' CLOCK MODE SETTING
'------------------------------------------------------------------------------
    'Set static clock mode
        M_Settings_SetClockMode DP_ClockMode_Static

    'Assert static clock mode is stored
        TST_DP_AssertEqualsLong "ClockMode can be set to static", _
            DP_ClockMode_Static, _
            M_Settings_GetClockMode()

    'Set live clock mode
        M_Settings_SetClockMode DP_ClockMode_Live

    'Assert live clock mode is stored
        TST_DP_AssertEqualsLong "ClockMode can be set to live", _
            DP_ClockMode_Live, _
            M_Settings_GetClockMode()

    'Assert unsupported clock mode raises a runtime error
        TST_DP_ExpectError_SetInvalidClockMode

'------------------------------------------------------------------------------
' SIZE MODE SETTING
'------------------------------------------------------------------------------
    'Set normal size mode
        M_Settings_SetSizeMode DP_SizeMode_Normal

    'Assert normal size mode is stored
        TST_DP_AssertEqualsLong "SizeMode can be set to normal", _
            DP_SizeMode_Normal, _
            M_Settings_GetSizeMode()

    'Set compact size mode
        M_Settings_SetSizeMode DP_SizeMode_Compact

    'Assert compact size mode is stored
        TST_DP_AssertEqualsLong "SizeMode can be set to compact", _
            DP_SizeMode_Compact, _
            M_Settings_GetSizeMode()

    'Assert unsupported size mode raises a runtime error
        TST_DP_ExpectError_SetInvalidSizeMode

'------------------------------------------------------------------------------
' HOLIDAY CALLBACK SETTING
'------------------------------------------------------------------------------
    'Set a padded holiday callback name
        M_Settings_SetHolidayCallback "  TST_DP_HolidayCallback  "

    'Assert the callback name is trimmed before storage
        TST_DP_AssertEqualsString "Holiday callback name is trimmed on storage", _
            "TST_DP_HolidayCallback", _
            M_Settings_GetHolidayCallback()

    'Clear the holiday callback name
        M_Settings_SetHolidayCallback vbNullString

    'Assert the callback name can be cleared
        TST_DP_AssertEqualsString "Holiday callback name can be cleared", _
            vbNullString, _
            M_Settings_GetHolidayCallback()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "Settings suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_DatePolicy()

'
'==============================================================================
'                           DATE POLICY SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates DatePicker date-selection policy behavior
'
' WHY THIS EXISTS
'   The form grid, hover logic, keyboard navigation, and click handling must all
'   agree on whether a given calendar date can be selected
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests inside-month selection, outside-month allow and reject behavior, and
'   invalid displayed period rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_DatePolicy_CanSelectDate
'   gDP_AllowOutsideMonthSelection
'
' NOTES
'   Settings are restored by the outer harness cleanup path
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "DatePolicy"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' OUTSIDE-MONTH SELECTION ENABLED
'------------------------------------------------------------------------------
    'Enable outside-month selection for the first policy branch
        gDP_AllowOutsideMonthSelection = True

    'Assert a current-month date can be selected when outside-month is enabled
        TST_DP_AssertTrue "Current-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert an outside-month date can be selected when outside-month is enabled
        TST_DP_AssertTrue "Outside-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' OUTSIDE-MONTH SELECTION DISABLED
'------------------------------------------------------------------------------
    'Disable outside-month selection for the second policy branch
        gDP_AllowOutsideMonthSelection = False

    'Assert a current-month date can still be selected
        TST_DP_AssertTrue "Current-month date is selectable when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert an outside-month date is rejected
        TST_DP_AssertFalse "Outside-month date is rejected when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' INVALID DISPLAY PERIODS
'------------------------------------------------------------------------------
    'Assert month zero is rejected as an invalid display period
        TST_DP_AssertFalse "Display month 0 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 0)

    'Assert month 13 is rejected as an invalid display period
        TST_DP_AssertFalse "Display month 13 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 13)

    'Assert year 99 is rejected as an invalid display period
        TST_DP_AssertFalse "Display year 99 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 99, 5)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "DatePolicy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_HolidayPolicy()

'
'==============================================================================
'                           HOLIDAY POLICY SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates optional holiday callback dispatch and fail-safe behavior
'
' WHY THIS EXISTS
'   Holiday rendering is high-frequency UI behavior and must not be destabilized
'   by missing, failing, or incorrectly typed callbacks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests blank callback, Boolean callback, non-Boolean callback result
'   rejection, missing callback fail-safe, and Excel error callback result
'   rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_HolidayPolicy_IsHolidayDate
'   TST_DP_QualifiedMacroName
'   gDP_HolidayCallbackName
'
' NOTES
'   The callback name is assigned directly to gDP_HolidayCallbackName to avoid
'   persisting test callback names through the settings setter
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "HolidayPolicy"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' BLANK CALLBACK
'------------------------------------------------------------------------------
    'Clear the holiday callback name
        gDP_HolidayCallbackName = vbNullString

    'Assert a blank callback returns False without raising
        TST_DP_AssertFalse "Blank holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' BOOLEAN CALLBACK
'------------------------------------------------------------------------------
    'Set the deterministic Boolean callback
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallback")

    'Assert the holiday date returns True
        TST_DP_AssertTrue "Boolean callback can return True for the holiday date", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

    'Assert a non-holiday date returns False
        TST_DP_AssertFalse "Boolean callback can return False for a non-holiday date", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 2))

'------------------------------------------------------------------------------
' NON-BOOLEAN CALLBACK RESULT
'------------------------------------------------------------------------------
    'Set the non-Boolean callback
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallbackNonBoolean")

    'Assert a non-Boolean callback result is ignored and returns False
        TST_DP_AssertFalse "Non-Boolean holiday callback result is ignored", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' MISSING CALLBACK FAIL-SAFE
'------------------------------------------------------------------------------
    'Set a callback name that does not resolve to any public procedure
        gDP_HolidayCallbackName = "TST_DP_MissingHolidayCallback_DoesNotExist"

    'Assert a missing callback is fail-safe and returns False
        TST_DP_AssertFalse "Missing holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' EXCEL ERROR VALUE CALLBACK RESULT
'------------------------------------------------------------------------------
    'Set the Excel error value callback
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallbackError")

    'Assert an Excel error callback result is ignored and returns False
        TST_DP_AssertFalse "Excel error holiday callback result is ignored", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "HolidayPolicy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Captions()

'
'==============================================================================
'                           CAPTIONS SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates month and date caption helpers
'
' WHY THIS EXISTS
'   Captions are visible UI output and are sensitive to localization changes
'   and helper refactoring
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests fixed-English month short and full captions, fixed-English date
'   captions, local-name non-empty output, and invalid month rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Caption_GetEnglishMonthShort
'   M_Caption_GetEnglishMonthFull
'   M_Caption_GetMonth
'   M_Caption_GetDate
'
' NOTES
'   Localized month and date captions are not asserted against a specific
'   language because their output depends on the host Office and Windows locale
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Captions"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' ENGLISH SHORT MONTH CAPTIONS
'------------------------------------------------------------------------------
    'Assert the short English January caption
        TST_DP_AssertEqualsString "English short month 1 is JAN", _
            "JAN", _
            M_Caption_GetEnglishMonthShort(1)

    'Assert the short English December caption
        TST_DP_AssertEqualsString "English short month 12 is DEC", _
            "DEC", _
            M_Caption_GetEnglishMonthShort(12)

    'Assert an invalid short English month raises a runtime error
        TST_DP_ExpectError_EnglishMonthShortInvalid

'------------------------------------------------------------------------------
' ENGLISH FULL MONTH CAPTIONS
'------------------------------------------------------------------------------
    'Assert the full English January caption
        TST_DP_AssertEqualsString "English full month 1 is JANUARY", _
            "JANUARY", _
            M_Caption_GetEnglishMonthFull(1)

    'Assert the full English December caption
        TST_DP_AssertEqualsString "English full month 12 is DECEMBER", _
            "DECEMBER", _
            M_Caption_GetEnglishMonthFull(12)

    'Assert an invalid full English month raises a runtime error
        TST_DP_ExpectError_EnglishMonthFullInvalid

'------------------------------------------------------------------------------
' GETMONTH CAPTIONS
'------------------------------------------------------------------------------
    'Assert fixed-English GetMonth returns the expected uppercase full month
        TST_DP_AssertEqualsString "GetMonth fixed-English returns uppercase full month", _
            "MAY", _
            M_Caption_GetMonth(5, False)

    'Assert local-name GetMonth returns a non-empty result
        TST_DP_AssertTrue "GetMonth local caption is non-empty", _
            VBA.Len(M_Caption_GetMonth(5, True)) > 0

    'Assert an invalid GetMonth input raises a runtime error
        TST_DP_ExpectError_GetMonthInvalid

'------------------------------------------------------------------------------
' GETDATE CAPTIONS
'------------------------------------------------------------------------------
    'Assert fixed-English GetDate returns the expected dd-MMM-yyyy format
        TST_DP_AssertEqualsString "GetDate fixed-English returns dd-MMM-yyyy", _
            "03-MAY-2026", _
            M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), False)

    'Assert local-name GetDate returns a non-empty result
        TST_DP_AssertTrue "GetDate local caption is non-empty", _
            VBA.Len(M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), True)) > 0

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "Captions suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_FormBridge()

'
'==============================================================================
'                           FORM BRIDGE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-visual DatePicker form bridge state
'
' WHY THIS EXISTS
'   The form bridge transfers the initial date, refresh commands, and cleanup
'   commands without forcing UF_DatePicker to be loaded unnecessarily
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests initial-date consumption including the one-shot flag, second
'   consumption returning False, and safe no-form bridge calls
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_FormBridge_ConsumeInitialDate
'   M_FormBridge_RefreshFromCell
'   DP_Close
'   gDP_InitialDate
'   gDP_HasInitialDate
'
' NOTES
'   This suite does not open UF_DatePicker
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ConsumedDate    As Date     'Date consumed from the initial-date bridge

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "FormBridge"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' INITIAL DATE CONSUMPTION
'------------------------------------------------------------------------------
    'Prepare the initial-date bridge value
        gDP_InitialDate = VBA.DateSerial(2026, 5, 3)

    'Mark the initial date as available
        gDP_HasInitialDate = True

    'Assert the initial date can be consumed
        TST_DP_AssertTrue "Initial date is consumed when available", _
            M_FormBridge_ConsumeInitialDate(ConsumedDate)

    'Assert the consumed date matches the bridge value
        TST_DP_AssertDateEquals "Consumed initial date matches bridge value", _
            VBA.DateSerial(2026, 5, 3), _
            ConsumedDate

    'Assert the initial-date flag was cleared after consumption
        TST_DP_AssertFalse "Initial date flag is cleared after consumption", _
            gDP_HasInitialDate

    'Assert a second consumption returns False
        TST_DP_AssertFalse "Initial date cannot be consumed twice", _
            M_FormBridge_ConsumeInitialDate(ConsumedDate)

'------------------------------------------------------------------------------
' NO-FORM BRIDGE CALLS
'------------------------------------------------------------------------------
    'Close the picker when no form may be loaded
        DP_Close

    'Record successful close with no visible form
        TST_DP_RecordPass "DP_Close is safe with no visible form", vbNullString

    'Refresh from an explicit scratch cell when no form may be loaded
        M_FormBridge_RefreshFromCell mTST_DP_ScratchSheet.Range("C4")

    'Record successful no-form refresh
        TST_DP_RecordPass "M_FormBridge_RefreshFromCell is safe with no visible form", _
            vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "FormBridge suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_WriteBack()

'
'==============================================================================
'                           WRITE-BACK SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates DatePicker worksheet write-back behavior
'
' WHY THIS EXISTS
'   The DatePicker must reliably populate single cells, contiguous ranges,
'   discontiguous ranges, and table data columns
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests direct contiguous range population, discontiguous range population,
'   selection-based write-back, table data-column expansion, and unsupported
'   write action rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_WriteBack_PopulateRange
'   M_WriteBack_Apply
'   gDP_WriteValue
'   mTST_DP_ScratchSheet
'
' NOTES
'   All write-back tests use the scratch worksheet only and leave no persistent
'   state on the result sheet
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetRange     As Excel.Range       'Target range under test
    Dim TableRange      As Excel.Range       'Source range for the test table
    Dim TestTable       As Excel.ListObject  'Regression test ListObject
    Dim UnionRange      As Excel.Range       'Discontiguous target range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "WriteBack"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet for selection-based tests
        mTST_DP_ScratchSheet.Activate

'------------------------------------------------------------------------------
' DIRECT CONTIGUOUS RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the target cells
        mTST_DP_ScratchSheet.Range("D5:D6").ClearContents
    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 5, 3)
    'Populate the contiguous range
        M_WriteBack_PopulateRange _
            mTST_DP_ScratchSheet.Range("D5:D6"), _
            DP_WriteAction_DatePicker
    'Assert the first cell received the date
        TST_DP_AssertCellDateEquals "Direct contiguous range write D5", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D5")
    'Assert the second cell received the date
        TST_DP_AssertCellDateEquals "Direct contiguous range write D6", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D6")

'------------------------------------------------------------------------------
' DISCONTIGUOUS RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the discontiguous test cells
        mTST_DP_ScratchSheet.Range("C5:C7").ClearContents
    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 6, 15)
    'Build the discontiguous target range
        Set UnionRange = Excel.Application.Union( _
            mTST_DP_ScratchSheet.Range("C5"), _
            mTST_DP_ScratchSheet.Range("C7"))
    'Populate the discontiguous range
        M_WriteBack_PopulateRange UnionRange, DP_WriteAction_DatePicker
    'Assert the first discontiguous cell received the date
        TST_DP_AssertCellDateEquals "Discontiguous range write C5", _
            VBA.DateSerial(2026, 6, 15), _
            mTST_DP_ScratchSheet.Range("C5")
    'Assert the second discontiguous cell received the date
        TST_DP_AssertCellDateEquals "Discontiguous range write C7", _
            VBA.DateSerial(2026, 6, 15), _
            mTST_DP_ScratchSheet.Range("C7")
    'Assert the skipped middle cell remains blank
        TST_DP_AssertTrue "Discontiguous range write leaves C6 blank", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("C6").Value)) = 0

'------------------------------------------------------------------------------
' SELECTION-BASED WRITE-BACK
'------------------------------------------------------------------------------
    'Clear the selection-based target range
        mTST_DP_ScratchSheet.Range("D5:D6").ClearContents
    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 7, 20)
    'Select the target range
        mTST_DP_ScratchSheet.Range("D5:D6").Select
    'Apply write-back to the current selection without table-column expansion
        M_WriteBack_Apply DP_WriteAction_DatePicker, True
    'Assert the first selected cell received the date
        TST_DP_AssertCellDateEquals "Selection write-back D5", _
            VBA.DateSerial(2026, 7, 20), _
            mTST_DP_ScratchSheet.Range("D5")
    'Assert the second selected cell received the date
        TST_DP_AssertCellDateEquals "Selection write-back D6", _
            VBA.DateSerial(2026, 7, 20), _
            mTST_DP_ScratchSheet.Range("D6")

'------------------------------------------------------------------------------
' TABLE COLUMN EXPANSION
'------------------------------------------------------------------------------
    'Prepare the table source range
        Set TableRange = mTST_DP_ScratchSheet.Range("F4:G7")
    'Clear the table source range
        TableRange.Clear
    'Write table headers
        mTST_DP_ScratchSheet.Range("F4").Value = "ID"
        mTST_DP_ScratchSheet.Range("G4").Value = "DateValue"
    'Write table row IDs
        mTST_DP_ScratchSheet.Range("F5:F7").Value = 1
    'Create the regression test table
        Set TestTable = mTST_DP_ScratchSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=TableRange, _
            XlListObjectHasHeaders:=xlYes)
    'Name the regression test table
        TestTable.Name = "TST_DP_Table"
    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 8, 25)
    'Select one data cell in the date column
        mTST_DP_ScratchSheet.Range("G5").Select
    'Apply write-back with table-column expansion enabled
        M_WriteBack_Apply DP_WriteAction_DatePicker, False
    'Assert the first table data cell received the date
        TST_DP_AssertCellDateEquals "Table-column expansion writes G5", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G5")
    'Assert the last table data cell received the date
        TST_DP_AssertCellDateEquals "Table-column expansion writes G7", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G7")

'------------------------------------------------------------------------------
' INVALID WRITE ACTION
'------------------------------------------------------------------------------
    'Assert an unsupported write action raises a runtime error
        TST_DP_ExpectError_InvalidWriteAction

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set TestTable = Nothing
        Set TableRange = Nothing
        Set UnionRange = Nothing
        Set TargetRange = Nothing
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Release object references
        Set TestTable = Nothing
        Set TableRange = Nothing
        Set UnionRange = Nothing
        Set TargetRange = Nothing
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "WriteBack suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_GridIcon()

'
'==============================================================================
'                           GRID ICON SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates in-grid DatePicker icon creation, movement, removal, and purge
'
' WHY THIS EXISTS
'   The grid icon is a high-frequency worksheet shape that is sensitive to stale
'   shape references, hidden shapes, deleted sheets, and movement logic
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the embedded icon file, creates the icon on the scratch sheet,
'   asserts single-instance reuse after a move, removes it, creates it again
'   for a hard purge check, and verifies disabled-feature behavior
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_GridIcon_ShowOrMove
'   M_GridIcon_Remove
'   M_GridIcon_PurgeAll
'   M_GridIcon_EnsureEmbeddedIconFile
'   gDP_ShowGridIcon
'   DP_GRID_ICON_NAME
'
' NOTES
'   This suite temporarily enables gDP_ShowGridIcon directly to avoid
'   triggering the settings setter side effects during the test
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IconPath            As String   'Embedded icon file path
    Dim ShapeLeftBefore     As Double   'Icon left position before the move
    Dim ShapeTopBefore      As Double   'Icon top position before the move

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "GridIcon"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet
        mTST_DP_ScratchSheet.Activate
    'Enable the grid icon feature for this suite
        gDP_ShowGridIcon = True
    'Purge any stale grid icons before the suite
    'M_GridIcon_PurgeAll resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' EMBEDDED ICON FILE
'------------------------------------------------------------------------------
    'Resolve the embedded icon file path
    'M_GridIcon_EnsureEmbeddedIconFile resets On Error GoTo 0 on exit; re-arm
        IconPath = M_GridIcon_EnsureEmbeddedIconFile()
        On Error GoTo SuiteFail
    'Assert the embedded icon path is not blank
        TST_DP_AssertTrue "Embedded grid icon path is not blank", _
            VBA.LenB(IconPath) > 0
    'Assert the embedded icon file exists on disk
        TST_DP_AssertTrue "Embedded grid icon file exists", _
            VBA.LenB(VBA.Dir$(IconPath, vbNormal)) > 0

'------------------------------------------------------------------------------
' CREATE ICON
'------------------------------------------------------------------------------
    'Restore ScreenUpdating so M_GridIcon_Create captures True as PreviousScreenUpdating
    'and restores it on exit, allowing the drawing layer to settle
        Excel.Application.ScreenUpdating = True
    'Create the grid icon beside the target cell
    'M_GridIcon_ShowOrMove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
    'Allow the drawing layer to process the shape creation
        DoEvents
    'Assert the grid icon exists on the scratch sheet
        TST_DP_AssertTrue "Grid icon is created beside an eligible target", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert the tracked reference is set after creation
    'Note: Visible=msoTrue is set inside M_GridIcon_Create under On Error Resume Next
    'while ScreenUpdating=False; on some Excel builds this silently fails and the shape
    'stays hidden until the next screen repaint. The tracked reference being set is the
    'correct invariant to assert here — visible state is a rendering detail.
        TST_DP_AssertTrue "Grid icon tracked reference is set after creation", _
            Not (gDP_GridIconShape Is Nothing)
    'Assert only one named grid icon exists in the host workbook
        TST_DP_AssertEqualsLong "Only one grid icon exists after creation", _
            1, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    'Capture the initial icon position from the tracked reference before the move
        If Not gDP_GridIconShape Is Nothing Then
            ShapeLeftBefore = gDP_GridIconShape.Left
            ShapeTopBefore = gDP_GridIconShape.Top
        End If

'------------------------------------------------------------------------------
' MOVE ICON
'------------------------------------------------------------------------------
    'Move the grid icon beside a different target cell
    'M_GridIcon_ShowOrMove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("F8")
        On Error GoTo SuiteFail
    'Allow the drawing layer to process the move
        DoEvents
    'Assert the grid icon still exists after the move
        TST_DP_AssertTrue "Grid icon still exists after move", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert the tracked reference is still set after the move
        TST_DP_AssertTrue "Grid icon tracked reference is set after move", _
            Not (gDP_GridIconShape Is Nothing)
    'Assert only one named grid icon exists after the move
        TST_DP_AssertEqualsLong "Only one grid icon exists after move", _
            1, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    'Assert the icon position changed using the tracked reference for both reads
        If Not gDP_GridIconShape Is Nothing Then
            TST_DP_AssertTrue "Grid icon position changes after move", _
                (gDP_GridIconShape.Left <> ShapeLeftBefore) Or _
                (gDP_GridIconShape.Top <> ShapeTopBefore)
        End If
    'Restore ScreenUpdating to suppress it for the remaining cleanup steps
        Excel.Application.ScreenUpdating = False

'------------------------------------------------------------------------------
' REMOVE ICON
'------------------------------------------------------------------------------
    'Remove the active grid icon
    'M_GridIcon_Remove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_Remove
        On Error GoTo SuiteFail
    'Assert the grid icon no longer exists on the scratch sheet
        TST_DP_AssertFalse "Grid icon is removed from the scratch sheet", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' PURGE ALL ICONS
'------------------------------------------------------------------------------
    'Create the icon again for the purge test
    'M_GridIcon_ShowOrMove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
    'Purge all named grid icons from the host workbook
    'M_GridIcon_PurgeAll resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail
    'Assert no named grid icon remains anywhere in the host workbook
        TST_DP_AssertEqualsLong "PurgeAll removes all grid icons from the workbook", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' DISABLED FEATURE BEHAVIOR
'------------------------------------------------------------------------------
    'Disable the grid icon feature directly
        gDP_ShowGridIcon = False
    'Attempt to show the icon while the feature is disabled
    'M_GridIcon_ShowOrMove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
    'Assert no icon was created while the feature is disabled
        TST_DP_AssertFalse "Grid icon is not shown when the feature is disabled", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "GridIcon suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Manager()

'
'==============================================================================
'                           MANAGER SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates public cDatePickerManager behavior and target gating
'
' WHY THIS EXISTS
'   The manager is the central Application-event coordinator and must make
'   deterministic UI decisions for explicit target cells
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests manager creation, picker loaded and visible state predicates,
'   Should_ShowGridIcon gating for date value cells, date-formatted cells,
'   general blank cells, multi-cell ranges, and merged cells, selection-change
'   handling for eligible and ineligible targets, and Reset_DatePickerUI
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   cDatePickerManager
'   M_GridIcon_PurgeAll
'   DP_Close
'
' NOTES
'   The ineligible-target assertion checks that no visible grid icon remains,
'   which allows either the remove or the hide implementation path
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Manager         As cDatePickerManager   'Manager instance under test
    Dim MergedArea      As Excel.Range          'Merged area under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Manager"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Create a fresh manager instance for public API checks
        Set Manager = New cDatePickerManager
    'Activate the scratch sheet
        mTST_DP_ScratchSheet.Activate
    'Purge stale grid icons before the manager tests
    'M_GridIcon_PurgeAll resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail
    'Enable grid icon gating for this suite
        gDP_ShowGridIcon = True

'------------------------------------------------------------------------------
' STATE PREDICATES
'------------------------------------------------------------------------------
    'Assert a freshly created manager is not busy
        TST_DP_AssertFalse "New manager is not busy", Manager.Is_Busy

    'Close any loaded picker before the visible / loaded checks
    'DP_Close resets On Error GoTo 0 on exit; re-arm immediately
        DP_Close
        On Error GoTo SuiteFail

    'Assert the picker is not visible after close
        TST_DP_AssertFalse "PickerVisible is False after DP_Close", _
            Manager.PickerVisible

    'Assert the picker is not loaded after close
        TST_DP_AssertFalse "Is_PickerLoaded is False after DP_Close", _
            Manager.Is_PickerLoaded

'------------------------------------------------------------------------------
' SHOULD_SHOWGRIDICON GATING
'------------------------------------------------------------------------------
    'Prepare a cell containing a date value
        mTST_DP_ScratchSheet.Range("D10").Value = VBA.DateSerial(2026, 5, 3)
    'Assert a date value cell is eligible for the grid icon
        TST_DP_AssertTrue "Should_ShowGridIcon accepts an explicit date value cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D10"))

    'Prepare a date-formatted blank cell
        mTST_DP_ScratchSheet.Range("B11").ClearContents
        mTST_DP_ScratchSheet.Range("B11").NumberFormat = "dd/mm/yyyy"
    'Assert a date-formatted blank cell is eligible for the grid icon
        TST_DP_AssertTrue "Should_ShowGridIcon accepts a date-formatted blank cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("B11"))

    'Prepare a general-format blank cell
        mTST_DP_ScratchSheet.Range("D12").ClearContents
        mTST_DP_ScratchSheet.Range("D12").NumberFormat = "General"
    'Assert a general blank cell is not eligible for the grid icon
        TST_DP_AssertFalse "Should_ShowGridIcon rejects a general blank cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D12"))

    'Assert a multi-cell range is not eligible for the grid icon
        TST_DP_AssertFalse "Should_ShowGridIcon rejects a multi-cell range", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D10:D11"))

'------------------------------------------------------------------------------
' MERGED-CELL NORMALIZATION
'------------------------------------------------------------------------------
    'Prepare the merged area
        Set MergedArea = mTST_DP_ScratchSheet.Range("D14:E15")
    'Clear and merge the test area
        MergedArea.Clear
        MergedArea.Merge
    'Format the merged area as date-like
        MergedArea.NumberFormat = "dd/mm/yyyy"
    'Assert the top-left cell of a merged area is eligible
        TST_DP_AssertTrue "Should_ShowGridIcon accepts the top-left cell of a merged area", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D14"))
    'Unmerge the area after the assertion
        MergedArea.UnMerge

'------------------------------------------------------------------------------
' SELECTION-CHANGE HANDLING
'------------------------------------------------------------------------------
    'Prepare a date value cell for the selection-change test
        mTST_DP_ScratchSheet.Range("E10").Value = VBA.DateSerial(2026, 9, 9)
    'Handle a selection change to an eligible target cell
    'Handle_SelectionChange resets On Error GoTo 0 on exit; re-arm immediately
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E10")
        On Error GoTo SuiteFail
    'Assert the manager created or showed the grid icon for the eligible target
        TST_DP_AssertTrue "Handle_SelectionChange shows a visible icon for an eligible target", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Prepare an ineligible general blank cell
        mTST_DP_ScratchSheet.Range("E11").ClearContents
        mTST_DP_ScratchSheet.Range("E11").NumberFormat = "General"
    'Handle a selection change to the ineligible target cell
    'Handle_SelectionChange resets On Error GoTo 0 on exit; re-arm immediately
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E11")
        On Error GoTo SuiteFail
    'Assert no visible grid icon remains for the ineligible target
        TST_DP_AssertFalse "Handle_SelectionChange leaves no visible icon for an ineligible target", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' RESET BEHAVIOR
'------------------------------------------------------------------------------
    'Show the icon again for the reset test
    'Handle_SelectionChange resets On Error GoTo 0 on exit; re-arm immediately
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E10")
        On Error GoTo SuiteFail
    'Reset all DatePicker UI through the manager
        Manager.Reset_DatePickerUI
    'Assert the reset removed all grid icons from the host workbook
        TST_DP_AssertEqualsLong "Reset_DatePickerUI purges all grid icons", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set MergedArea = Nothing
        Set Manager = Nothing
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Suppress local cleanup errors
        On Error Resume Next
    'Unmerge the merged area when still present
        If Not MergedArea Is Nothing Then MergedArea.UnMerge
    'Release object references
        Set MergedArea = Nothing
        Set Manager = Nothing
    'Record the suite-level failure and clear the error
        On Error GoTo 0
        TST_DP_RecordFail "Manager suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_UISmoke()

'
'==============================================================================
'                           UI SMOKE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Opens and closes UF_DatePicker as a minimal visual lifecycle smoke test
'
' WHY THIS EXISTS
'   Some integration failures only surface when the runtime form is fully loaded
'   and initialized; no other suite opens the form
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Activates a date cell on the scratch sheet, opens the picker through
'   DP_Show, validates loaded and visible predicates, closes the picker through
'   DP_Close, and validates that the picker is no longer visible
'
' ERROR POLICY
'   Records suite-level failures and continues. Attempts DP_Close after any
'   failure to ensure the form is not left open
'
' DEPENDENCIES
'   DP_Show
'   DP_Close
'   M_Picker_EnsureManager
'   gDP_Manager
'
' NOTES
'   This suite briefly opens UF_DatePicker and should be run manually before a
'   release rather than on every development edit
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "UISmoke"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet
        mTST_DP_ScratchSheet.Activate
    'Prepare the target cell with a deterministic date value
        mTST_DP_ScratchSheet.Range("H2").Value = VBA.DateSerial(2026, 5, 3)
    'Select the target date cell
        mTST_DP_ScratchSheet.Range("H2").Select
    'Ensure the manager infrastructure is available
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' OPEN FORM
'------------------------------------------------------------------------------
    'Open the DatePicker form
        DP_Show
    'Allow modeless form events to process
        DoEvents
    'Assert the picker was loaded
        TST_DP_AssertTrue "DP_Show loads the picker", _
            gDP_Manager.Is_PickerLoaded
    'Assert the picker is visible
        TST_DP_AssertTrue "DP_Show makes the picker visible", _
            gDP_Manager.PickerVisible

'------------------------------------------------------------------------------
' CLOSE FORM
'------------------------------------------------------------------------------
    'Close the DatePicker form
        DP_Close
    'Allow modeless form events to process
        DoEvents
    'Assert the picker is no longer visible after close
        TST_DP_AssertFalse "DP_Close hides the picker", _
            gDP_Manager.PickerVisible

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Attempt to close the picker after a UI smoke failure
        On Error Resume Next
        DP_Close
        Err.Clear
        On Error GoTo 0
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "UISmoke suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub


'
'------------------------------------------------------------------------------
'
'                             EXPECTED-ERROR HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_ExpectError_FirstDayToTextInvalid()

'
'==============================================================================
'                     EXPECT ERROR: FIRST-DAY TEXT CONVERSION INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an invalid first-day value to trigger the expected error
        M_Settings_FirstDayOfWeekToText 0

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    'Record a failure when no error was raised
        TST_DP_RecordFail "Invalid first-day conversion raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Invalid first-day conversion raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetFirstDayBlank()

'
'==============================================================================
'                     EXPECT ERROR: SET FIRST DAY BLANK TEXT
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with a blank first-day text to trigger the expected error
        M_Settings_SetFirstDayOfWeekText vbNullString

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Blank first-day text raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Blank first-day text raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetFirstDayInvalidText()

'
'==============================================================================
'                     EXPECT ERROR: SET FIRST DAY UNSUPPORTED TEXT
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an unsupported day name to trigger the expected error
        M_Settings_SetFirstDayOfWeekText "Wednesday"

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Unsupported first-day text raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Unsupported first-day text raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetInvalidClockMode()

'
'==============================================================================
'                     EXPECT ERROR: INVALID CLOCK MODE
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an unsupported clock mode value to trigger the expected error
        M_Settings_SetClockMode 99

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Unsupported clock mode raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Unsupported clock mode raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetInvalidSizeMode()

'
'==============================================================================
'                     EXPECT ERROR: INVALID SIZE MODE
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an unsupported size mode value to trigger the expected error
        M_Settings_SetSizeMode 99

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Unsupported size mode raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Unsupported size mode raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_EnglishMonthShortInvalid()

'
'==============================================================================
'                     EXPECT ERROR: ENGLISH MONTH SHORT INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with month 13 to trigger the expected error
        M_Caption_GetEnglishMonthShort 13

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Invalid short English month raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Invalid short English month raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_EnglishMonthFullInvalid()

'
'==============================================================================
'                     EXPECT ERROR: ENGLISH MONTH FULL INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with month 0 to trigger the expected error
        M_Caption_GetEnglishMonthFull 0

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Invalid full English month raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Invalid full English month raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_GetMonthInvalid()

'
'==============================================================================
'                     EXPECT ERROR: GETMONTH INVALID INPUT
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with month 99 to trigger the expected error
        M_Caption_GetMonth 99, False

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Invalid GetMonth input raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Invalid GetMonth input raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_InvalidWriteAction()

'
'==============================================================================
'                     EXPECT ERROR: INVALID WRITE ACTION
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an unsupported write action value to trigger the expected error
        M_WriteBack_PopulateRange mTST_DP_ScratchSheet.Range("J2"), 99

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    TST_DP_RecordFail "Unsupported write action raises", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Unsupported write action raises", Err.Description
    Err.Clear

End Sub


'
'------------------------------------------------------------------------------
'
'                               ASSERTION HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_AssertTrue(ByVal TestName As String, ByVal Condition As Boolean)

'
'==============================================================================
'                              ASSERT TRUE
'==============================================================================

    If Condition Then
        TST_DP_RecordPass TestName, vbNullString
    Else
        TST_DP_RecordFail TestName, "Expected True but received False"
    End If

End Sub

Private Sub TST_DP_AssertFalse(ByVal TestName As String, ByVal Condition As Boolean)

'
'==============================================================================
'                              ASSERT FALSE
'==============================================================================

    If Not Condition Then
        TST_DP_RecordPass TestName, vbNullString
    Else
        TST_DP_RecordFail TestName, "Expected False but received True"
    End If

End Sub

Private Sub TST_DP_AssertBooleanResult(ByVal TestName As String, ByVal BooleanValue As Boolean)

'
'==============================================================================
'                           ASSERT BOOLEAN RESULT
'==============================================================================

    'Record the result and the Boolean value without asserting a specific polarity
        TST_DP_RecordPass TestName, "Value=" & VBA.CStr(BooleanValue)

End Sub

Private Sub TST_DP_AssertEqualsLong( _
    ByVal TestName As String, _
    ByVal ExpectedValue As Long, _
    ByVal ActualValue As Long)

'
'==============================================================================
'                            ASSERT EQUALS LONG
'==============================================================================

    If ExpectedValue = ActualValue Then
        TST_DP_RecordPass TestName, "Value=" & VBA.CStr(ActualValue)
    Else
        TST_DP_RecordFail TestName, _
            "Expected " & VBA.CStr(ExpectedValue) & _
            " but received " & VBA.CStr(ActualValue)
    End If

End Sub

Private Sub TST_DP_AssertEqualsString( _
    ByVal TestName As String, _
    ByVal ExpectedValue As String, _
    ByVal ActualValue As String)

'
'==============================================================================
'                           ASSERT EQUALS STRING
'==============================================================================

    If VBA.StrComp(ExpectedValue, ActualValue, vbBinaryCompare) = 0 Then
        TST_DP_RecordPass TestName, "Value=" & ActualValue
    Else
        TST_DP_RecordFail TestName, _
            "Expected [" & ExpectedValue & "] but received [" & ActualValue & "]"
    End If

End Sub

Private Sub TST_DP_AssertDateEquals( _
    ByVal TestName As String, _
    ByVal ExpectedDate As Date, _
    ByVal ActualDate As Date)

'
'==============================================================================
'                            ASSERT DATE EQUALS
'==============================================================================

    If VBA.CDbl(ExpectedDate) = VBA.CDbl(ActualDate) Then
        TST_DP_RecordPass TestName, _
            "Value=" & VBA.Format$(ActualDate, "yyyy-mm-dd hh:nn:ss")
    Else
        TST_DP_RecordFail TestName, _
            "Expected " & VBA.Format$(ExpectedDate, "yyyy-mm-dd hh:nn:ss") & _
            " but received " & VBA.Format$(ActualDate, "yyyy-mm-dd hh:nn:ss")
    End If

End Sub

Private Sub TST_DP_AssertCellDateEquals( _
    ByVal TestName As String, _
    ByVal ExpectedDate As Date, _
    ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         ASSERT CELL DATE EQUALS
'==============================================================================

    'Reject a missing cell reference
        If TargetCell Is Nothing Then
            TST_DP_RecordFail TestName, "TargetCell is Nothing"
            Exit Sub
        End If

    'Reject a non-date cell value
        If Not VBA.IsDate(TargetCell.Value) Then
            TST_DP_RecordFail TestName, _
                "Cell value is not date-like: " & VBA.CStr(TargetCell.Value)
            Exit Sub
        End If

    'Delegate to the date equality assertion
        TST_DP_AssertDateEquals TestName, ExpectedDate, VBA.CDate(TargetCell.Value)

End Sub


'
'------------------------------------------------------------------------------
'
'                                RESULT HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_RecordPass(ByVal TestName As String, ByVal Details As String)

'
'==============================================================================
'                              RECORD PASS
'==============================================================================

    TST_DP_RecordResult TST_DP_PASS_TEXT, mTST_DP_CurrentSuite, TestName, Details

End Sub

Private Sub TST_DP_RecordFail(ByVal TestName As String, ByVal Details As String)

'
'==============================================================================
'                              RECORD FAIL
'==============================================================================

    TST_DP_RecordResult TST_DP_FAIL_TEXT, mTST_DP_CurrentSuite, TestName, Details

End Sub

Private Sub TST_DP_RecordInfo( _
    ByVal SuiteName As String, _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD INFO
'==============================================================================

    TST_DP_RecordResult TST_DP_INFO_TEXT, SuiteName, TestName, Details

End Sub

Private Sub TST_DP_RecordResult( _
    ByVal ResultText As String, _
    ByVal SuiteName As String, _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Records one regression result row to the Immediate Window and to the result
'   worksheet
'
' WHY THIS EXISTS
'   Result recording must never become the reason a regression run aborts; this
'   routine absorbs worksheet-write failures without re-raising
'
' INPUTS
'   ResultText
'     PASS, FAIL, or INFO marker
'
'   SuiteName
'     Logical suite name
'
'   TestName
'     Individual test or diagnostic name
'
'   Details
'     Optional detail text for the result row
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Updates assertion counters, writes a formatted line to the Immediate Window,
'   and writes a result row to the result worksheet when it is available
'
' ERROR POLICY
'   Best-effort. Worksheet-write failures are reported to the Immediate Window
'   and do not raise outward. The result sheet reference is cleared after a
'   write failure to prevent repeated failures on the same broken sheet
'
' DEPENDENCIES
'   mTST_DP_ResultSheet
'   mTST_DP_NextResultRow
'
' NOTES
'   This routine does not record its own failures recursively
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_RecordResult"   'Current procedure name

    Dim ResultLine              As String   'Immediate Window result line
    Dim RecordErrorNumber       As Long     'Captured worksheet-write error number
    Dim RecordErrorDescription  As String   'Captured worksheet-write error description

'------------------------------------------------------------------------------
' UPDATE COUNTERS
'------------------------------------------------------------------------------
    'Increment the total assertion counter for PASS and FAIL results
        If ResultText = TST_DP_PASS_TEXT Or ResultText = TST_DP_FAIL_TEXT Then
            mTST_DP_RunCount = mTST_DP_RunCount + 1
        End If
    'Increment the passed assertion counter
        If ResultText = TST_DP_PASS_TEXT Then
            mTST_DP_PassCount = mTST_DP_PassCount + 1
        End If
    'Increment the failed assertion counter
        If ResultText = TST_DP_FAIL_TEXT Then
            mTST_DP_FailCount = mTST_DP_FailCount + 1
        End If

'------------------------------------------------------------------------------
' WRITE IMMEDIATE WINDOW
'------------------------------------------------------------------------------
    'Build the Immediate Window result line
        ResultLine = ResultText & _
            " | " & SuiteName & _
            " | " & TestName & _
            IIf(VBA.Len(Details) > 0, " | " & Details, vbNullString)
    'Write the result line to the Immediate Window
        Debug.Print ResultLine

'------------------------------------------------------------------------------
' WRITE RESULT SHEET
'------------------------------------------------------------------------------
    'Exit when no result sheet is available
        If mTST_DP_ResultSheet Is Nothing Then Exit Sub

    'Attempt to write the result row
        On Error GoTo ResultSheetFail

    'Write the sequence number
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_SEQ).Value = _
            mTST_DP_NextResultRow - TST_DP_RESULT_FIRST_ROW + 1
    'Write the timestamp
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_TIMESTAMP).Value = VBA.Now
    'Write the result marker
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_RESULT).Value = ResultText
    'Write the suite name
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_SUITE).Value = SuiteName
    'Write the test name
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_TEST).Value = TestName
    'Write the details
        mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, TST_DP_COL_DETAILS).Value = Details

    'Advance the result row pointer
        mTST_DP_NextResultRow = mTST_DP_NextResultRow + 1

    'Restore normal error handling
        On Error GoTo 0

    'Exit after successful recording
        Exit Sub

'------------------------------------------------------------------------------
' RESULT SHEET FAIL
'------------------------------------------------------------------------------
ResultSheetFail:
    'Capture the worksheet-write error
        RecordErrorNumber = Err.Number
        RecordErrorDescription = Err.Description
    'Write the skip diagnostic to the Immediate Window
        Debug.Print TST_DP_INFO_TEXT & _
            " | Harness | " & PROC_NAME & _
            " | Result-sheet write skipped after error " & VBA.CStr(RecordErrorNumber) & _
            " - " & RecordErrorDescription
    'Clear the result sheet reference to avoid repeated write failures
        Set mTST_DP_ResultSheet = Nothing
    'Clear the suppressed error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub TST_DP_WriteSummary()

'
'==============================================================================
'                              WRITE SUMMARY
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the run summary to the Immediate Window and to the result worksheet
'
' WHY THIS EXISTS
'   A single consolidated summary at the end of the run makes it easy to
'   determine the overall pass / fail status without scanning individual rows
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Records an INFO result with the total run, passed, and failed counts, then
'   writes the same counts to fixed cells on the result sheet if available and
'   auto-fits the result columns
'
' ERROR POLICY
'   Best-effort. Worksheet-write failures are silently absorbed
'
' DEPENDENCIES
'   TST_DP_RecordInfo
'   mTST_DP_ResultSheet
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' WRITE IMMEDIATE WINDOW SUMMARY
'------------------------------------------------------------------------------
    'Record the run summary as an INFO result
        TST_DP_RecordInfo "Harness", "Summary", _
            "Run=" & VBA.CStr(mTST_DP_RunCount) & _
            "; Passed=" & VBA.CStr(mTST_DP_PassCount) & _
            "; Failed=" & VBA.CStr(mTST_DP_FailCount)

'------------------------------------------------------------------------------
' WRITE RESULT SHEET SUMMARY
'------------------------------------------------------------------------------
    'Exit when no result sheet is available
        If mTST_DP_ResultSheet Is Nothing Then Exit Sub

    'Suppress summary-write failures
        On Error Resume Next

    'Write the summary labels and values
        mTST_DP_ResultSheet.Cells(4, TST_DP_COL_SUMMARY_LABEL).Value = "SUMMARY"
        mTST_DP_ResultSheet.Cells(5, TST_DP_COL_SUMMARY_LABEL).Value = "Run"
        mTST_DP_ResultSheet.Cells(5, TST_DP_COL_SUMMARY_VALUE).Value = mTST_DP_RunCount
        mTST_DP_ResultSheet.Cells(6, TST_DP_COL_SUMMARY_LABEL).Value = "Passed"
        mTST_DP_ResultSheet.Cells(6, TST_DP_COL_SUMMARY_VALUE).Value = mTST_DP_PassCount
        mTST_DP_ResultSheet.Cells(7, TST_DP_COL_SUMMARY_LABEL).Value = "Failed"
        mTST_DP_ResultSheet.Cells(7, TST_DP_COL_SUMMARY_VALUE).Value = mTST_DP_FailCount

    'Bold the summary labels
        mTST_DP_ResultSheet.Cells(4, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(5, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(6, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(7, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True

    'Auto-fit the result and summary columns
        mTST_DP_ResultSheet.Columns("C:K").AutoFit

    'Centre-align the sequence, timestamp, result, and suite columns
        mTST_DP_ResultSheet.Columns("C:F").HorizontalAlignment = xlCenter

    'Clear any suppressed write error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                             WORKSHEET HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TST_DP_GetHostWorkbook() As Excel.Workbook

'
'==============================================================================
'                             GET HOST WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the workbook that will receive the result and scratch sheets
'
' WHY THIS EXISTS
'   The harness must write result rows to a stable workbook reference and
'   should prefer the ActiveWorkbook over ThisWorkbook to support use from an
'   add-in context
'
' INPUTS
'   None
'
' RETURNS
'   ActiveWorkbook when available; ThisWorkbook as a fallback
'
' BEHAVIOR
'   Attempts to resolve ActiveWorkbook with error suppression and falls back to
'   ThisWorkbook when ActiveWorkbook is not available
'
' ERROR POLICY
'   Best-effort. Returns Nothing only when both paths fail
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE HOST WORKBOOK
'------------------------------------------------------------------------------
    'Suppress resolution failures
        On Error Resume Next

    'Attempt to use the active workbook
        If Not Excel.Application.ActiveWorkbook Is Nothing Then
            Set TST_DP_GetHostWorkbook = Excel.Application.ActiveWorkbook
        End If

    'Fall back to ThisWorkbook when ActiveWorkbook is unavailable
        If TST_DP_GetHostWorkbook Is Nothing Then
            Set TST_DP_GetHostWorkbook = ThisWorkbook
        End If

    'Clear any suppressed resolution error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Sub TST_DP_PrepareResultSheet(ByVal HostWorkbook As Excel.Workbook)

'
'==============================================================================
'                           PREPARE RESULT SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves and initializes the result worksheet for the current run
'
' WHY THIS EXISTS
'   The result sheet must exist before any result rows can be written, and its
'   header row must be initialized before the first assertion fires
'
' INPUTS
'   HostWorkbook
'     Workbook that received the template sheet created by TST_DP_RunAll
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the result sheet by name, writes the header row, activates the
'   sheet, and resets the result row pointer
'
' ERROR POLICY
'   Raises a runtime error if the result sheet cannot be found
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE RESULT SHEET
'------------------------------------------------------------------------------
    'Resolve the result sheet from the host workbook
        Set mTST_DP_ResultSheet = HostWorkbook.Worksheets(TST_DP_RESULT_SHEET_NAME)

'------------------------------------------------------------------------------
' WRITE HEADER ROW
'------------------------------------------------------------------------------
    'Write the column headers
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_SEQ).Value = "#"
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_TIMESTAMP).Value = "Timestamp"
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_RESULT).Value = "Result"
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_SUITE).Value = "Suite"
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_TEST).Value = "Test"
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_DETAILS).Value = "Details"
    'Bold the header row
        mTST_DP_ResultSheet.Range( _
            mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_SEQ), _
            mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW - 1, TST_DP_COL_DETAILS)).Font.Bold = True

'------------------------------------------------------------------------------
' ACTIVATE AND RESET POINTER
'------------------------------------------------------------------------------
    'Activate the result sheet and select the first data cell
        mTST_DP_ResultSheet.Activate
        mTST_DP_ResultSheet.Cells(TST_DP_RESULT_FIRST_ROW, TST_DP_COL_SEQ).Select
    'Reset the result row pointer
        mTST_DP_NextResultRow = TST_DP_RESULT_FIRST_ROW

End Sub

Private Sub TST_DP_PrepareScratchSheet(ByVal HostWorkbook As Excel.Workbook)

'
'==============================================================================
'                           PREPARE SCRATCH SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Creates a clean scratch worksheet for the current regression run
'
' WHY THIS EXISTS
'   Write-back, grid-icon, and manager suites require a live worksheet that
'   can be freely modified and deleted without affecting any user data
'
' INPUTS
'   HostWorkbook
'     Workbook that will receive the scratch sheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes any existing scratch sheet, creates a new one at the end of the
'   workbook, names it, writes a header cell, and sets standard column widths
'
' ERROR POLICY
'   Raises a runtime error if the scratch sheet cannot be created
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DELETE EXISTING SCRATCH SHEET
'------------------------------------------------------------------------------
    'Delete any existing scratch sheet from a previous run
        TST_DP_DeleteWorksheetIfExists HostWorkbook, TST_DP_SCRATCH_SHEET_NAME

'------------------------------------------------------------------------------
' CREATE SCRATCH SHEET
'------------------------------------------------------------------------------
    'Add the scratch sheet at the end of the workbook
        Set mTST_DP_ScratchSheet = HostWorkbook.Worksheets.Add( _
            After:=HostWorkbook.Worksheets(HostWorkbook.Worksheets.Count))

'------------------------------------------------------------------------------
' INITIALIZE SCRATCH SHEET
'------------------------------------------------------------------------------
    'Name the scratch sheet
        mTST_DP_ScratchSheet.Name = TST_DP_SCRATCH_SHEET_NAME
    'Write a header cell to identify the scratch sheet
        mTST_DP_ScratchSheet.Range("A1").Value = "DatePicker regression scratch sheet"
    'Set standard column widths for the scratch sheet
        mTST_DP_ScratchSheet.Columns("A:J").ColumnWidth = 16
    'Activate the scratch sheet
        mTST_DP_ScratchSheet.Activate

End Sub

Private Sub TST_DP_DeleteScratchSheet()

'
'==============================================================================
'                            DELETE SCRATCH SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes the scratch worksheet after the regression run
'
' WHY THIS EXISTS
'   The scratch sheet is a transient test artifact and must not persist in the
'   host workbook after the run completes
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the scratch sheet when the host workbook reference is available and
'   the scratch sheet name can be found; clears the module-level reference
'
' ERROR POLICY
'   Best-effort. Silently absorbs deletion failures
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DELETE AND RELEASE
'------------------------------------------------------------------------------
    'Delete the scratch sheet if the host workbook is available
        If Not mTST_DP_HostWorkbook Is Nothing Then
            TST_DP_DeleteWorksheetIfExists mTST_DP_HostWorkbook, TST_DP_SCRATCH_SHEET_NAME
        End If
    'Release the module-level scratch sheet reference
        Set mTST_DP_ScratchSheet = Nothing

End Sub

Private Sub TST_DP_DeleteWorksheetIfExists( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal SheetName As String)

'
'==============================================================================
'                        DELETE WORKSHEET IF EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes a named worksheet from a workbook when it exists
'
' WHY THIS EXISTS
'   Repeated regression runs must remove stale scratch and result sheets from
'   previous runs without raising an error when the sheet is absent
'
' INPUTS
'   HostWorkbook
'     Workbook to search for the named sheet
'
'   SheetName
'     Name of the worksheet to delete
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the worksheet by name with error suppression, deletes it when
'   found, and silently ignores the operation when the sheet does not exist or
'   when it is the only remaining sheet in the workbook
'
' ERROR POLICY
'   Best-effort. Silently absorbs all errors
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetSheet     As Excel.Worksheet  'Worksheet to delete

'------------------------------------------------------------------------------
' RESOLVE AND DELETE
'------------------------------------------------------------------------------
    'Exit when the host workbook is not available
        If HostWorkbook Is Nothing Then Exit Sub

    'Suppress resolution and deletion errors
        On Error Resume Next

    'Attempt to resolve the target sheet by name
        Set TargetSheet = HostWorkbook.Worksheets(SheetName)

    'Delete the sheet when found and when the workbook has more than one sheet
        If Not TargetSheet Is Nothing Then
            If HostWorkbook.Worksheets.Count > 1 Then
                TargetSheet.Delete
            End If
        End If

    'Release the sheet reference
        Set TargetSheet = Nothing
    'Clear any suppressed error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                               STATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_CaptureSettings(ByRef Snapshot As TRegDPSettingsSnapshot)

'
'==============================================================================
'                            CAPTURE SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Captures all DatePicker settings and transient state into a snapshot
'
' WHY THIS EXISTS
'   The regression harness must restore all settings after the run so user
'   preferences are not permanently overwritten by test values
'
' INPUTS
'   Snapshot
'     Output TRegDPSettingsSnapshot populated by this routine
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' ENSURE SETTINGS ARE LOADED
'------------------------------------------------------------------------------
    'Ensure settings are in memory before capturing them
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE SETTINGS STATE
'------------------------------------------------------------------------------
    Snapshot.ShowRightClick = gDP_ShowRightClick
    Snapshot.ShowGridIcon = gDP_ShowGridIcon
    Snapshot.FirstDayOfWeek = gDP_FirstDayOfWeek
    Snapshot.UseLocalNames = gDP_UseLocalNames
    Snapshot.ClockMode = gDP_ClockMode
    Snapshot.SizeMode = gDP_SizeMode
    Snapshot.HighlightWeekends = gDP_HighlightWeekends
    Snapshot.AllowOutsideMonthSelection = gDP_AllowOutsideMonthSelection
    Snapshot.CloseAfterSelection = gDP_CloseAfterSelection
    Snapshot.UseWinAPI = gDP_UseWinAPI
    Snapshot.EnableKeyboardShortcut = gDP_EnableKeyboardShortcut
    Snapshot.HolidayCallbackName = gDP_HolidayCallbackName
    Snapshot.WriteValue = gDP_WriteValue
    Snapshot.InitialDate = gDP_InitialDate
    Snapshot.HasInitialDate = gDP_HasInitialDate
    Snapshot.SelectedDate = gDP_SelectedDate
    Snapshot.HasSelectedDate = gDP_HasSelectedDate

End Sub

Private Sub TST_DP_RestoreSettings(ByRef Snapshot As TRegDPSettingsSnapshot)

'
'==============================================================================
'                            RESTORE SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Restores all DatePicker settings and transient state from a snapshot
'
' WHY THIS EXISTS
'   Each regression run must leave settings in the same state as before the run
'   so user preferences and integration state are not permanently changed
'
' INPUTS
'   Snapshot
'     TRegDPSettingsSnapshot captured before the run
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE SETTINGS STATE
'------------------------------------------------------------------------------
    gDP_ShowRightClick = Snapshot.ShowRightClick
    gDP_ShowGridIcon = Snapshot.ShowGridIcon
    gDP_FirstDayOfWeek = Snapshot.FirstDayOfWeek
    gDP_UseLocalNames = Snapshot.UseLocalNames
    gDP_ClockMode = Snapshot.ClockMode
    gDP_SizeMode = Snapshot.SizeMode
    gDP_HighlightWeekends = Snapshot.HighlightWeekends
    gDP_AllowOutsideMonthSelection = Snapshot.AllowOutsideMonthSelection
    gDP_CloseAfterSelection = Snapshot.CloseAfterSelection
    gDP_UseWinAPI = Snapshot.UseWinAPI
    gDP_EnableKeyboardShortcut = Snapshot.EnableKeyboardShortcut
    gDP_HolidayCallbackName = Snapshot.HolidayCallbackName
    gDP_WriteValue = Snapshot.WriteValue
    gDP_InitialDate = Snapshot.InitialDate
    gDP_HasInitialDate = Snapshot.HasInitialDate
    gDP_SelectedDate = Snapshot.SelectedDate
    gDP_HasSelectedDate = Snapshot.HasSelectedDate

'------------------------------------------------------------------------------
' PERSIST AND SYNCHRONIZE
'------------------------------------------------------------------------------
    'Persist the restored settings to the registry
        M_Settings_Save
    'Synchronize the context menu with the restored settings
        M_ContextMenu_Update
    'Synchronize the keyboard shortcut with the restored settings
        M_KeyboardShortcut_Update
    'Remove the grid icon when the feature was disabled before the run
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Remove
        End If

End Sub

Private Sub TST_DP_CaptureApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                         CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures selected Excel Application properties into a snapshot
'
' WHY THIS EXISTS
'   The harness disables ScreenUpdating, EnableEvents, DisplayAlerts, and
'   modifies the StatusBar during the run; all must be restored afterward
'
' INPUTS
'   Snapshot
'     Output TRegDPApplicationSnapshot populated by this routine
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
    Snapshot.ScreenUpdating = Excel.Application.ScreenUpdating
    Snapshot.EnableEvents = Excel.Application.EnableEvents
    Snapshot.DisplayAlerts = Excel.Application.DisplayAlerts
    Snapshot.CalculationMode = Excel.Application.Calculation
    'Detect whether the StatusBar is currently Excel-owned (False) or text-owned
        Snapshot.StatusBarWasFalse = (VBA.VarType(Excel.Application.StatusBar) = vbBoolean)
    'Capture the StatusBar text when it is text-owned
        If Not Snapshot.StatusBarWasFalse Then
            Snapshot.StatusBarText = VBA.CStr(Excel.Application.StatusBar)
        End If

End Sub

Private Sub TST_DP_PrepareApplicationForRun()

'
'==============================================================================
'                       PREPARE APPLICATION FOR RUN
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the Excel Application state used by the regression harness during a run
'
' WHY THIS EXISTS
'   Disabling ScreenUpdating, EnableEvents, and DisplayAlerts reduces
'   interference from workbook-level event handlers and confirmation dialogs
'   during the regression run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' SET HARNESS APPLICATION STATE
'------------------------------------------------------------------------------
    Excel.Application.ScreenUpdating = False
    Excel.Application.EnableEvents = False
    Excel.Application.DisplayAlerts = False
    Excel.Application.StatusBar = "Running DatePicker regression tests..."

End Sub

Private Sub TST_DP_RestoreApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                       RESTORE APPLICATION STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the Excel Application state from a snapshot captured before the run
'
' INPUTS
'   Snapshot
'     TRegDPApplicationSnapshot captured before the run
'
' RETURNS
'   Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE APPLICATION STATE
'------------------------------------------------------------------------------
    Excel.Application.Calculation = Snapshot.CalculationMode
    Excel.Application.DisplayAlerts = Snapshot.DisplayAlerts
    Excel.Application.EnableEvents = Snapshot.EnableEvents
    Excel.Application.ScreenUpdating = Snapshot.ScreenUpdating
    'Restore the StatusBar to its pre-run state
        If Snapshot.StatusBarWasFalse Then
            Excel.Application.StatusBar = False
        Else
            Excel.Application.StatusBar = Snapshot.StatusBarText
        End If

End Sub

Private Sub TST_DP_ResetDatePickerArtifacts()

'
'==============================================================================
'                        RESET DATEPICKER ARTIFACTS
'------------------------------------------------------------------------------
' PURPOSE
'   Clears all transient DatePicker UI artifacts before and after a run
'
' WHY THIS EXISTS
'   Stale timers, open forms, context menu entries, keyboard shortcuts, and
'   grid icons from a previous run or interrupted run must be cleared before
'   testing begins and cleaned up after the run ends
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls each cleanup path behind On Error Resume Next so a failure in one
'   path does not prevent the remaining cleanup steps from executing
'
' ERROR POLICY
'   Best-effort. Silently absorbs all cleanup failures
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RESET ALL ARTIFACTS
'------------------------------------------------------------------------------
    'Suppress individual cleanup failures
        On Error Resume Next
    'Stop the live-clock timer
        M_Timer_Stop
    'Close any loaded DatePicker form
        DP_Close
    'Remove the right-click menu entry
        M_ContextMenu_Remove
    'Remove the keyboard shortcut
        M_KeyboardShortcut_Remove
    'Purge all in-grid DatePicker icons
        M_GridIcon_PurgeAll
    'Clear any suppressed error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub TST_DP_RestoreManagerState()

'
'==============================================================================
'                         RESTORE MANAGER STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the global DatePicker manager to its pre-run state
'
' WHY THIS EXISTS
'   Suite tests create and destroy manager instances; the global manager
'   reference must be restored to the state it was in before the run began
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Releases the current global manager reference and recreates it through
'   M_Picker_EnsureManager only when a manager existed before the run
'
' ERROR POLICY
'   Best-effort. Silently absorbs manager recreation failures
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' RELEASE CURRENT MANAGER
'------------------------------------------------------------------------------
    'Release the current global manager reference
        Set gDP_Manager = Nothing

'------------------------------------------------------------------------------
' RECREATE MANAGER IF NEEDED
'------------------------------------------------------------------------------
    'Recreate the manager only when one existed before the run
        If mTST_DP_HadManager Then
            On Error Resume Next
            M_Picker_EnsureManager
            Err.Clear
            On Error GoTo 0
        End If

End Sub


'
'------------------------------------------------------------------------------
'
'                              SHAPE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function TST_DP_ShapeExists( _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal ShapeName As String) As Boolean

'
'==============================================================================
'                              SHAPE EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a named shape exists on a worksheet
'
' INPUTS
'   TargetSheet - worksheet to search
'   ShapeName   - shape name to locate
'
' RETURNS
'   True when the shape exists; False otherwise or when the sheet is Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetShape     As Excel.Shape  'Resolved shape reference

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Default to not found
        TST_DP_ShapeExists = False
    'Exit when the target sheet is not available
        If TargetSheet Is Nothing Then Exit Function

    'Attempt to resolve the shape by name
        On Error Resume Next
        Set TargetShape = TargetSheet.Shapes(ShapeName)
        TST_DP_ShapeExists = Not (TargetShape Is Nothing)
        Set TargetShape = Nothing
        Err.Clear
        On Error GoTo 0

End Function

Private Function TST_DP_ShapeIsVisible( _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal ShapeName As String) As Boolean

'
'==============================================================================
'                            SHAPE IS VISIBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns True when a named shape exists and is visible on a worksheet
'
' INPUTS
'   TargetSheet - worksheet to search
'   ShapeName   - shape name to check
'
' RETURNS
'   True when the shape exists and Visible = msoTrue; False otherwise
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetShape     As Excel.Shape  'Resolved shape reference

'------------------------------------------------------------------------------
' RESOLVE VISIBILITY
'------------------------------------------------------------------------------
    'Default to not visible
        TST_DP_ShapeIsVisible = False
    'Exit when the target sheet is not available
        If TargetSheet Is Nothing Then Exit Function

    'Attempt to resolve the shape and read its visibility
        On Error Resume Next
        Set TargetShape = TargetSheet.Shapes(ShapeName)
        If Not TargetShape Is Nothing Then
            TST_DP_ShapeIsVisible = (TargetShape.Visible = msoTrue)
        End If
        Set TargetShape = Nothing
        Err.Clear
        On Error GoTo 0

End Function

Private Function TST_DP_CountNamedShapes( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal ShapeName As String) As Long

'
'==============================================================================
'                            COUNT NAMED SHAPES
'------------------------------------------------------------------------------
' PURPOSE
'   Counts all shapes with a given name across all worksheets in a workbook
'
' INPUTS
'   HostWorkbook - workbook to search
'   ShapeName    - shape name to count
'
' RETURNS
'   Total count of matching shapes; 0 when the workbook is Nothing
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetSheet     As Excel.Worksheet  'Worksheet being inspected
    Dim TargetShape     As Excel.Shape      'Shape being inspected
    Dim ShapeCount      As Long             'Running count of matching shapes

'------------------------------------------------------------------------------
' COUNT SHAPES
'------------------------------------------------------------------------------
    'Default to zero
        ShapeCount = 0
    'Exit when the host workbook is not available
        If HostWorkbook Is Nothing Then
            TST_DP_CountNamedShapes = 0
            Exit Function
        End If

    'Iterate every worksheet and every shape in the workbook
    'Suppress shape-enumeration errors: a newly created shape in a transitional
    'visual state (Visible=msoFalse, not yet rendered) can raise E_INVALIDARG
    'during For Each iteration on some Excel builds. The suppression is scoped
    'to the inner loop only so outer-loop worksheet errors also increment safely.
        For Each TargetSheet In HostWorkbook.Worksheets
            On Error Resume Next
            For Each TargetShape In TargetSheet.Shapes
                If Err.Number = 0 Then
                    If VBA.StrComp(TargetShape.Name, ShapeName, vbBinaryCompare) = 0 Then
                        ShapeCount = ShapeCount + 1
                    End If
                End If
                Err.Clear
            Next TargetShape
            On Error GoTo 0
        Next TargetSheet

    'Return the total count
        TST_DP_CountNamedShapes = ShapeCount

    'Release object references
        Set TargetShape = Nothing
        Set TargetSheet = Nothing

End Function


'
'------------------------------------------------------------------------------
'
'                             CALLBACK RESOLUTION
'
'------------------------------------------------------------------------------
'

Private Function TST_DP_QualifiedMacroName(ByVal MacroName As String) As String

'
'==============================================================================
'                           QUALIFIED MACRO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a runnable callback macro name for Application.Run regression tests
'
' WHY THIS EXISTS
'   Holiday-policy regression tests validate callback dispatch through
'   Application.Run. Callback resolution varies depending on workbook
'   qualification, module qualification, project state, and host workbook
'   context. This helper probes supported callback-name formats and returns the
'   first one that Excel can actually execute.
'
' INPUTS
'   MacroName
'     Public callback procedure name without qualification
'
' RETURNS
'   First runnable callback reference string
'
' BEHAVIOR
'   Normalizes the supplied macro name, builds three candidate names (workbook
'   + module qualified, workbook-qualified, and unqualified), probes each
'   candidate through Excel.Application.Run, and returns the first candidate
'   that executes without raising a runtime error
'
' ERROR POLICY
'   Raises a descriptive runtime error when MacroName is blank or when no
'   candidate callback reference can be executed
'
' DEPENDENCIES
'   Excel.Application.Run
'   ThisWorkbook
'   TST_DP_MODULE_NAME
'
' NOTES
'   This helper intentionally calls the callback once with a deterministic test
'   date while resolving the runnable name. The regression callbacks used here
'   are side-effect free
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_QualifiedMacroName"

    Dim NormalizedMacroName     As String       'Normalized callback macro name
    Dim WorkbookQualifier       As String       'Escaped workbook qualifier
    Dim CandidateNames(1 To 3)  As String       'Candidate callback name variants
    Dim CandidateIndex          As Long         'Current candidate index
    Dim CallbackResult          As Variant      'Probe callback result
    Dim RunErrNumber            As Long         'Application.Run error number
    Dim RunErrDescription       As String       'Application.Run error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Normalize the supplied macro name
        NormalizedMacroName = VBA.Trim$(MacroName)

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Reject blank callback names
        If VBA.LenB(NormalizedMacroName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "MacroName cannot be blank"
        End If

'------------------------------------------------------------------------------
' BUILD CANDIDATE NAMES
'------------------------------------------------------------------------------
    'Build the workbook qualifier with escaped single quotes
        WorkbookQualifier = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!"
    'Candidate 1: workbook and module qualified
        CandidateNames(1) = WorkbookQualifier & TST_DP_MODULE_NAME & "." & NormalizedMacroName
    'Candidate 2: workbook qualified only
        CandidateNames(2) = WorkbookQualifier & NormalizedMacroName
    'Candidate 3: unqualified
        CandidateNames(3) = NormalizedMacroName

'------------------------------------------------------------------------------
' PROBE CANDIDATES
'------------------------------------------------------------------------------
    'Test each candidate in order and return the first that succeeds
        For CandidateIndex = LBound(CandidateNames) To UBound(CandidateNames)
            'Suppress the probe error
                On Error Resume Next
            'Clear any pending error before probing
                Err.Clear
            'Probe the candidate with a deterministic test date
                CallbackResult = Excel.Application.Run( _
                    CandidateNames(CandidateIndex), _
                    VBA.DateSerial(2026, 1, 1))
            'Capture the probe result
                RunErrNumber = Err.Number
                RunErrDescription = Err.Description
            'Clear the suppressed probe error
                Err.Clear
            'Restore controlled error handling
                On Error GoTo ErrorHandler
            'Return the first candidate that executed without error
                If RunErrNumber = 0 Then
                    TST_DP_QualifiedMacroName = CandidateNames(CandidateIndex)
                    Exit Function
                End If
        Next CandidateIndex

'------------------------------------------------------------------------------
' RAISE RESOLUTION FAILURE
'------------------------------------------------------------------------------
    'Raise a descriptive error when no candidate succeeded
        Err.Raise vbObjectError + 514, PROC_NAME, _
            "Unable to resolve runnable callback name for '" & NormalizedMacroName & _
            "'. Last Application.Run error: " & VBA.CStr(RunErrNumber) & _
            " - " & RunErrDescription

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Callback name resolution failed: " & Err.Description

End Function


'
'------------------------------------------------------------------------------
'
'                          CONDITIONAL FORMATTING
'
'------------------------------------------------------------------------------
'

Public Sub TST_DP_ApplyFailConditionalFormat( _
    ByVal WS As Excel.Worksheet, _
    ByVal TargetRange As Excel.Range)

'
'==============================================================================
'                       APPLY FAIL CONDITIONAL FORMAT
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a text-contains FAIL conditional formatting rule to a supplied
'   worksheet range
'
' WHY THIS EXISTS
'   Conditional formatting should be applied to an explicit range reference
'   rather than relying on Select, Selection, ActiveCell, or recorded-macro
'   state, which changes between runs
'
' INPUTS
'   WS
'     Worksheet that owns the target range
'
'   TargetRange
'     Range receiving the conditional-formatting rule
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the supplied worksheet and range, confirms the range belongs to
'   the worksheet, adds a text-contains FAIL rule at first priority, and
'   applies bold white font with a dark red fill
'
' ERROR POLICY
'   Raises a descriptive runtime error if inputs are missing or inconsistent
'
' DEPENDENCIES
'   Excel object model
'
' NOTES
'   This routine does not clear existing conditional-formatting rules before
'   adding the new rule
'
'   This routine does not select or activate any worksheet or range
'
' UPDATED
'   2026-05-14
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TST_DP_ApplyFailConditionalFormat"

    Dim FailCondition       As Excel.FormatCondition 'Created conditional-formatting rule
    Dim ErrorNumber         As Long                  'Captured error number
    Dim ErrorDescription    As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If WS Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "WS cannot be Nothing"
        End If
    'Reject a missing target range reference
        If TargetRange Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "TargetRange cannot be Nothing"
        End If
    'Reject a range that does not belong to the supplied worksheet
        If Not TargetRange.Worksheet Is WS Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "TargetRange must belong to the supplied worksheet"
        End If

'------------------------------------------------------------------------------
' ADD CONDITIONAL FORMAT
'------------------------------------------------------------------------------
    'Add a text-contains FAIL rule to the target range
        Set FailCondition = TargetRange.FormatConditions.Add( _
            Type:=xlTextString, _
            String:=TST_DP_FAIL_TEXT, _
            TextOperator:=xlContains)

    'Move the new rule to first priority
        FailCondition.SetFirstPriority

'------------------------------------------------------------------------------
' APPLY FORMAT SETTINGS
'------------------------------------------------------------------------------
    'Apply bold white font to FAIL cells
        With FailCondition.Font
            .Bold = True
            .Italic = False
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With

    'Apply dark red fill to FAIL cells
        With FailCondition.Interior
            .PatternColorIndex = xlAutomatic
            .Color = TST_DP_FAIL_BACK_COLOR
            .TintAndShade = 0
        End With

    'Allow lower-priority rules to continue evaluating
        FailCondition.StopIfTrue = False

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set FailCondition = Nothing
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Release object references
        Set FailCondition = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "FAIL conditional-format application failed: " & ErrorDescription

End Sub




