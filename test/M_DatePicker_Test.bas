Attribute VB_Name = "M_DatePicker_Test"
Option Explicit

'
'==============================================================================
' MODULE: M_REGRESSION_DATEPICKER
'------------------------------------------------------------------------------
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
'   selected Excel Application state
'
' ERROR POLICY
'   The harness records individual assertion failures and continues where safe
'
'   A fatal harness failure is recorded, cleanup is attempted, and the original
'   error is re-raised to the caller
'
' DEPENDENCIES
'   M_DATEPICKER
'   cDatePickerManager
'   UF_DatePicker
'   Excel object model
'
' NOTES
'   Before running this module, the DatePicker project must compile
'
'   TST_DP_RunAll excludes disruptive UI smoke tests by default
'
'   TST_DP_RunAll_WithUISmoke opens the DatePicker form briefly and closes it
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TST_DP_RESULT_SHEET_NAME      As String = "TST_DP_RESULTS"   'Regression result worksheet name
    Private Const TST_DP_SCRATCH_SHEET_NAME     As String = "TST_DP_SCRATCH"   'Regression scratch worksheet name
    Private Const TST_DP_PASS_TEXT              As String = "PASS"             'Passed test marker
    Private Const TST_DP_FAIL_TEXT              As String = "FAIL"             'Failed test marker
    Private Const TST_DP_INFO_TEXT              As String = "INFO"             'Information marker
    Private Const TST_DP_MODULE_NAME            As String = "M_REGRESSION_DATEPICKER" 'Regression standard module name

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type TRegDPSettingsSnapshot
        ShowRightClick             As Boolean    'Right-click feature setting
        ShowGridIcon               As Boolean    'Grid-icon feature setting
        FirstDayOfWeek             As Long       'First-day setting
        UseLocalNames              As Boolean    'Local-name setting
        ClockMode                  As Long       'Clock-mode setting
        SizeMode                   As Long       'Size-mode setting
        HighlightWeekends          As Boolean    'Weekend-highlight setting
        AllowOutsideMonthSelection As Boolean    'Outside-month selection setting
        CloseAfterSelection        As Boolean    'Close-after-selection setting
        UseWinAPI                  As Boolean    'WinAPI setting
        EnableKeyboardShortcut     As Boolean    'Keyboard shortcut setting
        HolidayCallbackName        As String     'Holiday callback setting
        WriteValue                 As Date       'Transient write value
        InitialDate                As Date       'Transient initial date
        HasInitialDate             As Boolean    'Transient initial-date flag
        SelectedDate               As Date       'Transient selected date
        HasSelectedDate            As Boolean    'Transient selected-date flag
    End Type

    Private Type TRegDPApplicationSnapshot
        ScreenUpdating             As Boolean    'Application.ScreenUpdating snapshot
        EnableEvents               As Boolean    'Application.EnableEvents snapshot
        DisplayAlerts              As Boolean    'Application.DisplayAlerts snapshot
        CalculationMode            As Long       'Application.Calculation snapshot
        StatusBarWasFalse          As Boolean    'True when StatusBar was Excel-owned
        StatusBarText              As String     'StatusBar text snapshot
    End Type

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mTST_DP_ResultSheet    As Excel.Worksheet  'Result worksheet used by the current run
    Private mTST_DP_ScratchSheet   As Excel.Worksheet  'Scratch worksheet used by the current run
    Private mTST_DP_HostWorkbook   As Excel.Workbook   'Workbook receiving test sheets
    Private mTST_DP_NextResultRow  As Long             'Next result row
    Private mTST_DP_RunCount       As Long             'Total assertions executed
    Private mTST_DP_PassCount      As Long             'Total assertions passed
    Private mTST_DP_FailCount      As Long             'Total assertions failed
    Private mTST_DP_CurrentSuite   As String           'Current suite name
    Private mTST_DP_HadManager     As Boolean          'True when a manager existed before the run

'
'------------------------------------------------------------------------------
'
'                              PUBLIC ENTRY POINTS
'
'------------------------------------------------------------------------------

Public Sub TST_DP_RunAll()

'
'==============================================================================
'                           RUN ALL REGRESSION TESTS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the full non-disruptive DatePicker regression pack
'
' WHY THIS EXISTS
'   This is the normal regression entry point for development and release checks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs environment, settings, policy, caption, form bridge, write-back,
'   grid-icon, and manager tests without opening the DatePicker UserForm
'
' ERROR POLICY
'   Delegates fatal handling and cleanup to TST_DP_RunAllInternal
'
' DEPENDENCIES
'   TST_DP_RunAllInternal
'
' NOTES
'   Use TST_DP_RunAll_WithUISmoke when you also want a brief UserForm open/close
'   smoke test
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' BUILD TEST SHEET
'------------------------------------------------------------------------------
    'Reset counters and module state
        TST_DP_ResetHarnessState
    'Resolve the workbook that will receive regression sheets
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()
    'Buil template
        DEMO_Sheet_BuildTemplate TST_DP_RESULT_SHEET_NAME, "DATE PICKER", _
        "Test Sheet"
        
'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the standard non-disruptive regression pack
        TST_DP_RunAllInternal True

End Sub

Public Sub TST_DP_RunAll_WithUISmoke()

'
'==============================================================================
'                           RUN ALL WITH UI SMOKE
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the full DatePicker regression pack including a UserForm smoke test
'
' WHY THIS EXISTS
'   UI smoke testing is useful before a release, but it briefly opens the
'   DatePicker form and is therefore separated from the default regression run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the standard regression pack and then opens / closes UF_DatePicker
'
' ERROR POLICY
'   Delegates fatal handling and cleanup to TST_DP_RunAllInternal
'
' DEPENDENCIES
'   TST_DP_RunAllInternal
'
' NOTES
'   This routine is intentionally visible as a separate macro
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the regression pack with the optional UI smoke suite
        TST_DP_RunAllInternal True

End Sub

Public Function TST_DP_HolidayCallback(ByVal CandidateDate As Date) As Boolean

'
'==============================================================================
'                           TEST HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a deterministic Boolean holiday callback for regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate calls user callbacks through Application.Run
'   and needs a stable callback target during regression
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   True only for 1 January 2026; otherwise False
'
' BEHAVIOR
'   Compares the date-only component of CandidateDate to 1 January 2026
'
' ERROR POLICY
'   Raises no intentional errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   This function must remain Public so Application.Run can call it
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return True only for the deterministic holiday date
        TST_DP_HolidayCallback = _
            (VBA.DateValue(CandidateDate) = VBA.DateSerial(2026, 1, 1))

End Function

Public Function TST_DP_HolidayCallbackNonBoolean(ByVal CandidateDate As Date) As String

'
'==============================================================================
'                       TEST NON-BOOLEAN HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a non-Boolean callback result for holiday-policy regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate should accept only explicit Boolean callback
'   results and ignore non-Boolean values
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   String value "TRUE"
'
' BEHAVIOR
'   Returns text rather than Boolean deliberately
'
' ERROR POLICY
'   Raises no intentional errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   CandidateDate is intentionally unused
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return a non-Boolean value deliberately
        TST_DP_HolidayCallbackNonBoolean = "TRUE"

End Function

Public Function TST_DP_HolidayCallbackError(ByVal CandidateDate As Date) As Variant

'
'==============================================================================
'                         TEST ERROR-VALUE HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides an Excel error value for holiday-policy regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate should ignore callback results that are Excel
'   error values and return False rather than treating them as holidays
'
' INPUTS
'   CandidateDate
'     Candidate date supplied by the DatePicker holiday policy
'
' RETURNS
'   CVErr(xlErrValue)
'
' BEHAVIOR
'   Returns an Excel error value deliberately without raising a VBA runtime error
'
' ERROR POLICY
'   Raises no intentional runtime errors
'
' DEPENDENCIES
'   CVErr
'   xlErrValue
'
' NOTES
'   CandidateDate is intentionally unused
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' RETURN ERROR VALUE
'------------------------------------------------------------------------------
    'Return an Excel error value deliberately
        TST_DP_HolidayCallbackError = CVErr(xlErrValue)

End Function

'
'------------------------------------------------------------------------------
'
'                               RUN ORCHESTRATION
'
'------------------------------------------------------------------------------

Private Sub TST_DP_RunAllInternal(ByVal IncludeUISmoke As Boolean)

'
'==============================================================================
'                           RUN ALL INTERNAL
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates one complete DatePicker regression run
'
' WHY THIS EXISTS
'   The harness must preserve user settings, isolate scratch workbook artifacts,
'   restore Excel state, and continue through grouped suites where possible
'
' INPUTS
'   IncludeUISmoke
'     True to include the DatePicker form open/close smoke suite
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures state, prepares scratch assets, runs suites, writes summary, and
'   restores state
'
' ERROR POLICY
'   Records fatal harness failures, attempts cleanup, then re-raises the error
'
' DEPENDENCIES
'   All TST_DP_RunSuite_* routines
'
' NOTES
'   This routine disables Application events during the run to reduce
'   interference from workbook-level event handlers
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_RunAllInternal" 'Current procedure name

    Dim SettingsSnapshot        As TRegDPSettingsSnapshot           'DatePicker state snapshot
    Dim AppSnapshot             As TRegDPApplicationSnapshot        'Excel Application state snapshot
    Dim FatalNumber             As Long                             'Fatal error number
    Dim FatalDescription        As String                           'Fatal error description
    Dim HasFatalError           As Boolean                          'True when fatal error must be re-raised

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled fatal handling
        On Error GoTo FatalHandler
    'Capture whether a manager existed before the run
        mTST_DP_HadManager = Not (gDP_Manager Is Nothing)
    'Capture current DatePicker settings and transient state
        TST_DP_CaptureSettings SettingsSnapshot
    'Capture current Application state
        TST_DP_CaptureApplicationState AppSnapshot

'------------------------------------------------------------------------------
' PREPARE ISOLATED RUN STATE
'------------------------------------------------------------------------------
    'Prepare the Application state used by the harness
        TST_DP_PrepareApplicationForRun
    'Reset DatePicker UI artifacts before testing
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
    'Run date-policy checks
        TST_DP_RunSuiteSafe "DatePolicy"
    'Run holiday-policy callback checks
        TST_DP_RunSuiteSafe "HolidayPolicy"
    'Run caption helper checks
        TST_DP_RunSuiteSafe "Captions"
    'Run form bridge checks
        TST_DP_RunSuiteSafe "FormBridge"
    'Run worksheet write-back checks
        TST_DP_RunSuiteSafe "WriteBack"
    'Run grid-icon shape checks
        TST_DP_RunSuiteSafe "GridIcon"
    'Run manager public API checks
        TST_DP_RunSuiteSafe "Manager"
    'Run optional UserForm smoke checks when requested
        If IncludeUISmoke Then
            TST_DP_RunSuiteSafe "UISmoke"
        End If

'------------------------------------------------------------------------------
' WRITE SUMMARY
'------------------------------------------------------------------------------
    'Write the final run summary
        TST_DP_WriteSummary
'------------------------------------------------------------------------------
' APPLY CONDITION FORMATTING
'------------------------------------------------------------------------------
    TST_DP_ApplyFailConditionalFormat mTST_DP_ResultSheet, mTST_DP_ResultSheet.Range("E5:E1000")

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

    'Restore the manager state according to the pre-run state
        TST_DP_RestoreManagerState

    'Restore the Application state
        TST_DP_RestoreApplicationState AppSnapshot

    'Clear module object references
        Set mTST_DP_ScratchSheet = Nothing
        Set mTST_DP_ResultSheet = Nothing
        Set mTST_DP_HostWorkbook = Nothing

    'Clear any suppressed cleanup error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

    'Re-raise fatal harness errors after cleanup
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

    'Record the fatal harness failure when possible
        On Error Resume Next
        TST_DP_RecordResult TST_DP_FAIL_TEXT, _
            "Harness", _
            PROC_NAME, _
            "Fatal error " & VBA.CStr(FatalNumber) & " - " & FatalDescription
        Err.Clear
        On Error GoTo 0

    'Run shared cleanup
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
'   Regression runs may be launched repeatedly in the same Excel session
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets counters, current suite, result row, and object references
'
' ERROR POLICY
'   Does not raise intentional errors
'
' DEPENDENCIES
'   Module-level state
'
' NOTES
'   Object cleanup itself is handled by the orchestrator cleanup path
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' RESET COUNTERS
'------------------------------------------------------------------------------
    'Reset assertion counters
        mTST_DP_RunCount = 0
    'Reset pass counter
        mTST_DP_PassCount = 0
    'Reset fail counter
        mTST_DP_FailCount = 0
    'Reset result-row pointer
        mTST_DP_NextResultRow = 5
    'Reset current suite name
        mTST_DP_CurrentSuite = vbNullString

'------------------------------------------------------------------------------
' RESET REFERENCES
'------------------------------------------------------------------------------
    'Clear result sheet reference
        Set mTST_DP_ResultSheet = Nothing
    'Clear scratch sheet reference
        Set mTST_DP_ScratchSheet = Nothing
    'Clear host workbook reference
        Set mTST_DP_HostWorkbook = Nothing

End Sub

'
'------------------------------------------------------------------------------
'
'                                  TEST SUITES
'
'------------------------------------------------------------------------------

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
'   second boundary so Excel object-model errors, environment-specific failures,
'   or unexpected project changes cannot abort the full regression run
'
' INPUTS
'   SuiteName
'     Logical suite name to run
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Dispatches the requested suite by name and records any escaping error as a
'   suite-level failure before allowing the harness to continue
'
' ERROR POLICY
'   Best-effort
'   Does not intentionally raise outward
'
' DEPENDENCIES
'   TST_DP_RunSuite_* routines
'
' NOTES
'   This wrapper is deliberately Select Case based so the suite routines can
'   remain Private and do not need Application.Run
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_RunSuiteSafe" 'Current procedure name

    Dim ErrorNumber             As Long                          'Captured suite error number
    Dim ErrorDescription        As String                        'Captured suite error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Normalize the current suite name for diagnostics
        mTST_DP_CurrentSuite = SuiteName
    'Protect the harness from escaping suite failures
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' DISPATCH SUITE
'------------------------------------------------------------------------------
    'Run the requested suite
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
                TST_DP_RecordFail "Unknown suite", "SuiteName=" & SuiteName
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

Private Sub TST_DP_RunSuite_Environment()

'
'==============================================================================
'                           RUN ENVIRONMENT SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates basic DatePicker infrastructure startup and platform helpers
'
' WHY THIS EXISTS
'   Later suites assume settings, manager creation, and platform helpers can run
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Executes non-destructive environment checks
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Picker_EnsureManager
'   M_Platform_CanUseWinAPI
'   M_Platform_ShouldUseWinAPI
'
' NOTES
'   The manager may report Is_Hooked False only if Excel Application hook-up
'   fails in the host environment
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Environment"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' SETTINGS AND MANAGER
'------------------------------------------------------------------------------
    'Ensure settings load successfully
        M_Settings_EnsureLoaded
    'Record successful settings load
        TST_DP_RecordPass "M_Settings_EnsureLoaded does not raise", vbNullString

    'Ensure the global manager can be created
        M_Picker_EnsureManager

    'Assert that the global manager exists
        TST_DP_AssertTrue "Global manager is instantiated", Not (gDP_Manager Is Nothing)

    'Assert that the manager is not busy after startup
        TST_DP_AssertFalse "Manager is not busy after startup", gDP_Manager.Is_Busy

'------------------------------------------------------------------------------
' PLATFORM HELPERS
'------------------------------------------------------------------------------
    'Call the platform capability helper
        TST_DP_AssertBooleanResult "M_Platform_CanUseWinAPI returns Boolean", _
            M_Platform_CanUseWinAPI

    'Call the effective WinAPI helper
        TST_DP_AssertBooleanResult "M_Platform_ShouldUseWinAPI returns Boolean", _
            M_Platform_ShouldUseWinAPI

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Environment suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Settings()

'
'==============================================================================
'                           RUN SETTINGS SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates public DatePicker settings accessors and setters
'
' WHY THIS EXISTS
'   Settings drive form rendering, manager behavior, shortcut integration,
'   context-menu availability, grid-icon availability, and date-selection rules
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exercises valid settings, invalid setting rejection, text parsing, and
'   access-path fallback normalization
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
'   2026-05-08
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
    'Assert vbSunday is supported
        TST_DP_AssertTrue "vbSunday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbSunday)

    'Assert vbMonday is supported
        TST_DP_AssertTrue "vbMonday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbMonday)

    'Assert zero is unsupported
        TST_DP_AssertFalse "0 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(0)

    'Assert Tuesday-style value is unsupported
        TST_DP_AssertFalse "3 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(3)

    'Assert Sunday text conversion
        TST_DP_AssertEqualsString "vbSunday converts to text", _
            "vbSunday", _
            M_Settings_FirstDayOfWeekToText(vbSunday)

    'Assert Monday text conversion
        TST_DP_AssertEqualsString "vbMonday converts to text", _
            "vbMonday", _
            M_Settings_FirstDayOfWeekToText(vbMonday)

    'Assert invalid first-day conversion raises
        TST_DP_ExpectError_FirstDayToTextInvalid

'------------------------------------------------------------------------------
' FIRST-DAY SETTERS
'------------------------------------------------------------------------------
    'Set Sunday as first day
        M_Settings_SetFirstDayOfWeek vbSunday

    'Assert Sunday was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeek stores vbSunday", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Monday through text parser
        M_Settings_SetFirstDayOfWeekText "Monday"

    'Assert Monday was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Monday", _
            vbMonday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Sunday through short text parser
        M_Settings_SetFirstDayOfWeekText "Sun"

    'Assert Sunday was stored
        TST_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Sun", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Assert blank first-day text raises
        TST_DP_ExpectError_SetFirstDayBlank

    'Assert unsupported first-day text raises
        TST_DP_ExpectError_SetFirstDayInvalidText

'------------------------------------------------------------------------------
' BOOLEAN SETTINGS
'------------------------------------------------------------------------------
    'Set local names on
        M_Settings_SetUseLocalNames True

    'Assert local names setting is on
        TST_DP_AssertTrue "UseLocalNames can be enabled", _
            M_Settings_GetUseLocalNames()

    'Set local names off
        M_Settings_SetUseLocalNames False

    'Assert local names setting is off
        TST_DP_AssertFalse "UseLocalNames can be disabled", _
            M_Settings_GetUseLocalNames()

    'Set outside-month selection on
        M_Settings_SetAllowOutsideMonthSelection True

    'Assert outside-month selection is on
        TST_DP_AssertTrue "AllowOutsideMonthSelection can be enabled", _
            M_Settings_GetAllowOutsideMonthSelection()

    'Set outside-month selection off
        M_Settings_SetAllowOutsideMonthSelection False

    'Assert outside-month selection is off
        TST_DP_AssertFalse "AllowOutsideMonthSelection can be disabled", _
            M_Settings_GetAllowOutsideMonthSelection()

    'Set weekend highlighting on
        M_Settings_SetHighlightWeekends True

    'Assert weekend highlighting is on
        TST_DP_AssertTrue "HighlightWeekends can be enabled", _
            M_Settings_GetHighlightWeekends()

    'Set weekend highlighting off
        M_Settings_SetHighlightWeekends False

    'Assert weekend highlighting is off
        TST_DP_AssertFalse "HighlightWeekends can be disabled", _
            M_Settings_GetHighlightWeekends()

    'Set close-after-selection on
        M_Settings_SetCloseAfterSelection True

    'Assert close-after-selection is on
        TST_DP_AssertTrue "CloseAfterSelection can be enabled", _
            M_Settings_GetCloseAfterSelection()

    'Set close-after-selection off
        M_Settings_SetCloseAfterSelection False

    'Assert close-after-selection is off
        TST_DP_AssertFalse "CloseAfterSelection can be disabled", _
            M_Settings_GetCloseAfterSelection()

'------------------------------------------------------------------------------
' FEATURE SETTINGS
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

    'Disable keyboard shortcut while other access paths are available
        M_Settings_SetShowRightClick True
        M_Settings_SetShowGridIcon True
        M_Settings_SetEnableKeyboardShortcut False

    'Assert keyboard shortcut can be disabled when other access paths exist
        TST_DP_AssertFalse "Keyboard shortcut can be disabled when other access paths exist", _
            M_Settings_GetEnableKeyboardShortcut()

    'Re-enable keyboard shortcut
        M_Settings_SetEnableKeyboardShortcut True

    'Assert keyboard shortcut can be enabled
        TST_DP_AssertTrue "Keyboard shortcut can be enabled", _
            M_Settings_GetEnableKeyboardShortcut()

    'Disable both visible/contextual access paths
        M_Settings_SetShowRightClick False
        M_Settings_SetShowGridIcon False

    'Assert keyboard access is forced on when both visible entry points are off
        TST_DP_AssertTrue "Keyboard shortcut is forced when right-click and grid icon are disabled", _
            M_Settings_GetEnableKeyboardShortcut()

    'Restore feature settings for following suites
        M_Settings_SetShowRightClick True
        M_Settings_SetShowGridIcon True
        M_Settings_SetEnableKeyboardShortcut True

'------------------------------------------------------------------------------
' WINAPI SETTING
'------------------------------------------------------------------------------
    'Disable WinAPI setting
        M_Settings_SetUseWinAPI False

    'Assert WinAPI setting can be disabled
        TST_DP_AssertFalse "UseWinAPI can be disabled", _
            M_Settings_GetUseWinAPI()

    'Enable WinAPI setting according to platform capability
        M_Settings_SetUseWinAPI True

    'Assert WinAPI setting is normalized to platform capability
        TST_DP_AssertEqualsLong "UseWinAPI normalizes to platform capability", _
            VBA.CLng(M_Platform_CanUseWinAPI), _
            VBA.CLng(M_Settings_GetUseWinAPI())

'------------------------------------------------------------------------------
' ENUM SETTINGS
'------------------------------------------------------------------------------
    'Set static clock mode
        M_Settings_SetClockMode DP_ClockMode_Static

    'Assert static clock mode is stored
        TST_DP_AssertEqualsLong "ClockMode can be static", _
            DP_ClockMode_Static, _
            M_Settings_GetClockMode()

    'Set live clock mode
        M_Settings_SetClockMode DP_ClockMode_Live

    'Assert live clock mode is stored
        TST_DP_AssertEqualsLong "ClockMode can be live", _
            DP_ClockMode_Live, _
            M_Settings_GetClockMode()

    'Assert unsupported clock mode raises
        TST_DP_ExpectError_SetInvalidClockMode

    'Set normal size mode
        M_Settings_SetSizeMode DP_SizeMode_Normal

    'Assert normal size mode is stored
        TST_DP_AssertEqualsLong "SizeMode can be normal", _
            DP_SizeMode_Normal, _
            M_Settings_GetSizeMode()

    'Set compact size mode
        M_Settings_SetSizeMode DP_SizeMode_Compact

    'Assert compact size mode is stored
        TST_DP_AssertEqualsLong "SizeMode can be compact", _
            DP_SizeMode_Compact, _
            M_Settings_GetSizeMode()

    'Assert unsupported size mode raises
        TST_DP_ExpectError_SetInvalidSizeMode

'------------------------------------------------------------------------------
' HOLIDAY CALLBACK SETTING
'------------------------------------------------------------------------------
    'Set a trimmed holiday callback name
        M_Settings_SetHolidayCallback "  TST_DP_HolidayCallback  "

    'Assert the callback name was trimmed
        TST_DP_AssertEqualsString "Holiday callback name is trimmed", _
            "TST_DP_HolidayCallback", _
            M_Settings_GetHolidayCallback()

    'Clear the holiday callback name
        M_Settings_SetHolidayCallback vbNullString

    'Assert the callback name was cleared
        TST_DP_AssertEqualsString "Holiday callback name can be cleared", _
            vbNullString, _
            M_Settings_GetHolidayCallback()

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Settings suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_DatePolicy()

'
'==============================================================================
'                           RUN DATE POLICY SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates DatePicker date-selection policy behavior
'
' WHY THIS EXISTS
'   The form grid, hover logic, keyboard navigation, and click handling must all
'   agree on whether a calendar date can be selected
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests inside-month selection, outside-month allow / reject behavior, and
'   invalid displayed period rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_DatePolicy_CanSelectDate
'
' NOTES
'   Settings are restored by the outer harness cleanup path
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "DatePolicy"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' OUTSIDE-MONTH ENABLED
'------------------------------------------------------------------------------
    'Enable outside-month selection for the first policy branch
        gDP_AllowOutsideMonthSelection = True

    'Assert current-month date can be selected
        TST_DP_AssertTrue "Current-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert outside-month date can be selected
        TST_DP_AssertTrue "Outside-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' OUTSIDE-MONTH DISABLED
'------------------------------------------------------------------------------
    'Disable outside-month selection for the second policy branch
        gDP_AllowOutsideMonthSelection = False

    'Assert current-month date can still be selected
        TST_DP_AssertTrue "Current-month date is selectable when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert outside-month date is rejected
        TST_DP_AssertFalse "Outside-month date is rejected when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' INVALID DISPLAY PERIODS
'------------------------------------------------------------------------------
    'Assert invalid month is rejected
        TST_DP_AssertFalse "Display month 0 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 0)

    'Assert invalid month is rejected
        TST_DP_AssertFalse "Display month 13 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 13)

    'Assert invalid year is rejected
        TST_DP_AssertFalse "Display year 99 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 99, 5)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Date policy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_HolidayPolicy()

'
'==============================================================================
'                           RUN HOLIDAY POLICY SUITE
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
'   Tests blank callback, Boolean callback, non-Boolean callback, missing
'   callback, and Excel error callback-result behavior
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_HolidayPolicy_IsHolidayDate
'   Application.Run
'
' NOTES
'   The callback name is assigned directly to avoid persisting test callback
'   names through the settings setter
'
' UPDATED
'   2026-05-08
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
    'Clear callback name
        gDP_HolidayCallbackName = vbNullString

    'Assert blank callback returns False
        TST_DP_AssertFalse "Blank holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' BOOLEAN CALLBACK
'------------------------------------------------------------------------------
    'Set deterministic Boolean callback
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallback")

    'Assert matching holiday returns True
        TST_DP_AssertTrue "Boolean callback can return True", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

    'Assert non-matching holiday returns False
        TST_DP_AssertFalse "Boolean callback can return False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 2))

'------------------------------------------------------------------------------
' NON-BOOLEAN CALLBACK
'------------------------------------------------------------------------------
    'Set non-Boolean callback
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallbackNonBoolean")

    'Assert non-Boolean callback result is ignored
        TST_DP_AssertFalse "Non-Boolean holiday callback is ignored", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' MISSING AND ERROR-VALUE CALLBACKS
'------------------------------------------------------------------------------
    'Set missing callback name
        gDP_HolidayCallbackName = "TST_DP_MissingHolidayCallback"

    'Assert missing callback is fail-safe
        TST_DP_AssertFalse "Missing holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

    'Set error-value callback name
        gDP_HolidayCallbackName = TST_DP_QualifiedMacroName("TST_DP_HolidayCallbackError")

    'Assert Excel error callback result is ignored
        TST_DP_AssertFalse "Excel error holiday callback result returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Holiday policy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Captions()

'
'==============================================================================
'                           RUN CAPTION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates month and date caption helpers
'
' WHY THIS EXISTS
'   Captions are visible UI output and are sensitive to localization and helper
'   refactoring
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests fixed-English month captions, fixed-English date captions, local-name
'   non-empty output, and invalid month rejection
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Caption_* public helpers
'
' NOTES
'   Localized month names are not asserted against a specific language because
'   they depend on the host Office / Windows locale
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Captions"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' ENGLISH MONTH CAPTIONS
'------------------------------------------------------------------------------
    'Assert short English January caption
        TST_DP_AssertEqualsString "English short month 1 is JAN", _
            "JAN", _
            M_Caption_GetEnglishMonthShort(1)

    'Assert short English December caption
        TST_DP_AssertEqualsString "English short month 12 is DEC", _
            "DEC", _
            M_Caption_GetEnglishMonthShort(12)

    'Assert full English January caption
        TST_DP_AssertEqualsString "English full month 1 is JANUARY", _
            "JANUARY", _
            M_Caption_GetEnglishMonthFull(1)

    'Assert full English December caption
        TST_DP_AssertEqualsString "English full month 12 is DECEMBER", _
            "DECEMBER", _
            M_Caption_GetEnglishMonthFull(12)

    'Assert fixed-English month helper output
        TST_DP_AssertEqualsString "GetMonth fixed-English returns uppercase full month", _
            "MAY", _
            M_Caption_GetMonth(5, False)

'------------------------------------------------------------------------------
' DATE CAPTIONS
'------------------------------------------------------------------------------
    'Assert fixed-English date caption
        TST_DP_AssertEqualsString "GetDate fixed-English returns dd-MMM-yyyy", _
            "03-MAY-2026", _
            M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), False)

    'Assert local month caption is non-empty
        TST_DP_AssertTrue "GetMonth local caption is non-empty", _
            VBA.Len(M_Caption_GetMonth(5, True)) > 0

    'Assert local date caption is non-empty
        TST_DP_AssertTrue "GetDate local caption is non-empty", _
            VBA.Len(M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), True)) > 0

'------------------------------------------------------------------------------
' INVALID INPUTS
'------------------------------------------------------------------------------
    'Assert invalid short month raises
        TST_DP_ExpectError_EnglishMonthShortInvalid

    'Assert invalid full month raises
        TST_DP_ExpectError_EnglishMonthFullInvalid

    'Assert invalid GetMonth raises
        TST_DP_ExpectError_GetMonthInvalid

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Caption suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_FormBridge()

'
'==============================================================================
'                           RUN FORM BRIDGE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates non-visual DatePicker form bridge state
'
' WHY THIS EXISTS
'   The form bridge transfers initial date state, refresh commands, and cleanup
'   commands without forcing default UserForm creation unnecessarily
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests initial-date consumption and explicit-cell no-form cleanup / refresh
'   calls
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_FormBridge_ConsumeInitialDate
'   M_FormBridge_RefreshFromCell
'   DP_Close
'
' NOTES
'   This suite does not open UF_DatePicker
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InitialDate             As Date     'Consumed initial date

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
    'Prepare initial-date bridge state
        gDP_InitialDate = VBA.DateSerial(2026, 5, 3)

    'Mark the initial date as available
        gDP_HasInitialDate = True

    'Assert initial date can be consumed
        TST_DP_AssertTrue "Initial date is consumed when available", _
            M_FormBridge_ConsumeInitialDate(InitialDate)

    'Assert consumed initial date is correct
        TST_DP_AssertDateEquals "Consumed initial date matches bridge value", _
            VBA.DateSerial(2026, 5, 3), _
            InitialDate

    'Assert initial date flag was cleared after consumption
        TST_DP_AssertFalse "Initial date flag is cleared after consumption", _
            gDP_HasInitialDate

    'Assert second consumption returns False
        TST_DP_AssertFalse "Initial date cannot be consumed twice", _
            M_FormBridge_ConsumeInitialDate(InitialDate)

'------------------------------------------------------------------------------
' NO-FORM BRIDGE CALLS
'------------------------------------------------------------------------------
    'Close picker when no form may be loaded
        DP_Close

    'Record successful close with no visible form
        TST_DP_RecordPass "DP_Close is safe with no visible form", vbNullString

    'Refresh from an explicit scratch cell when no form may be loaded
        M_FormBridge_RefreshFromCell mTST_DP_ScratchSheet.Range("C4")

    'Record successful no-form refresh
        TST_DP_RecordPass "M_FormBridge_RefreshFromCell is safe with explicit cell and no visible form", vbNullString

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Form bridge suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_WriteBack()

'
'==============================================================================
'                           RUN WRITE-BACK SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates DatePicker worksheet write-back behavior
'
' WHY THIS EXISTS
'   The DatePicker must reliably populate single cells, ranges, discontiguous
'   ranges, and table data columns
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests direct range population, selection-based write-back, discontiguous
'   range write-back, and table data-column expansion
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_WriteBack_PopulateRange
'   M_WriteBack_Apply
'
' NOTES
'   This suite uses the scratch worksheet only
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetRange             As Excel.Range       'Target range under test
    Dim TableRange              As Excel.Range       'Table range under test
    Dim TestTable               As Excel.ListObject  'Regression test table
    Dim UnionRange              As Excel.Range       'Discontiguous target range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "WriteBack"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Ensure the scratch sheet is active for selection-based tests
        mTST_DP_ScratchSheet.Activate

'------------------------------------------------------------------------------
' DIRECT RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the target range
        mTST_DP_ScratchSheet.Range("D5:D6").ClearContents

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 5, 3)

    'Populate a contiguous range directly
        M_WriteBack_PopulateRange _
            mTST_DP_ScratchSheet.Range("D5:D6"), _
            DP_WriteAction_DatePicker

    'Assert the first cell was written
        TST_DP_AssertCellDateEquals "Direct range write B2", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D5")

    'Assert the last cell was written
        TST_DP_AssertCellDateEquals "Direct range write B4", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D6")

'------------------------------------------------------------------------------
' DISCONTIGUOUS RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the discontiguous test cells
        mTST_DP_ScratchSheet.Range("C5:C7").ClearContents

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 6, 15)

    'Build a discontiguous target range
        Set UnionRange = Excel.Application.Union( _
            mTST_DP_ScratchSheet.Range("C5"), _
            mTST_DP_ScratchSheet.Range("C7"))

    'Populate the discontiguous range directly
        M_WriteBack_PopulateRange UnionRange, DP_WriteAction_DatePicker

    'Assert the first discontiguous cell was written
        TST_DP_AssertCellDateEquals "Discontiguous write C5", _
            VBA.DateSerial(2026, 6, 15), _
            mTST_DP_ScratchSheet.Range("C5")

    'Assert the second discontiguous cell was written
        TST_DP_AssertCellDateEquals "Discontiguous write C7", _
            VBA.DateSerial(2026, 6, 15), _
            mTST_DP_ScratchSheet.Range("C7")

    'Assert the skipped middle cell remains blank
        TST_DP_AssertTrue "Discontiguous write leaves C3 blank", _
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

    'Apply DatePicker write-back to the current selection without table growth
        M_WriteBack_Apply DP_WriteAction_DatePicker, True

    'Assert the first selected cell was written
        TST_DP_AssertCellDateEquals "Selection write D2", _
            VBA.DateSerial(2026, 7, 20), _
            mTST_DP_ScratchSheet.Range("D5")

    'Assert the second selected cell was written
        TST_DP_AssertCellDateEquals "Selection write D3", _
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

    'Write table IDs
        mTST_DP_ScratchSheet.Range("F5:F7").Value = 1

    'Create the regression table
        Set TestTable = mTST_DP_ScratchSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=TableRange, _
            XlListObjectHasHeaders:=xlYes)

    'Name the regression table
        TestTable.Name = "TST_DP_Table"

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 8, 25)

    'Select one table data cell in the date column
        mTST_DP_ScratchSheet.Range("G5").Select

    'Apply DatePicker write-back with table-column expansion enabled
        M_WriteBack_Apply DP_WriteAction_DatePicker, False

    'Assert the first table data cell was written
        TST_DP_AssertCellDateEquals "Table-column expansion writes G5", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G5")

    'Assert the last table data cell was written
        TST_DP_AssertCellDateEquals "Table-column expansion writes G7", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G7")

'------------------------------------------------------------------------------
' INVALID WRITE ACTION
'------------------------------------------------------------------------------
    'Assert unsupported write action raises
        TST_DP_ExpectError_InvalidWriteAction

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release the table reference
        Set TestTable = Nothing
    'Release the table range reference
        Set TableRange = Nothing
    'Release the union range reference
        Set UnionRange = Nothing
    'Release the target range reference
        Set TargetRange = Nothing
    'Exit after the suite succeeds
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

    'Record the suite-level failure
        TST_DP_RecordFail "Write-back suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_GridIcon()

'
'==============================================================================
'                           RUN GRID ICON SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates in-grid DatePicker icon creation, movement, removal, and purge
'
' WHY THIS EXISTS
'   The grid icon is a high-frequency worksheet shape and is sensitive to stale
'   references, hidden shapes, deleted sheets, and movement logic
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates / moves the icon on the scratch sheet, verifies single-shape reuse,
'   removes it, and checks hard purge behavior
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_GridIcon_ShowOrMove
'   M_GridIcon_Remove
'   M_GridIcon_PurgeAll
'   M_GridIcon_EnsureEmbeddedIconFile
'
' NOTES
'   This suite temporarily enables gDP_ShowGridIcon directly
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IconPath                As String       'Embedded icon file path
    Dim ShapeLeftBefore         As Double       'Icon left position before move
    Dim ShapeTopBefore          As Double       'Icon top position before move

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "GridIcon"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Ensure the scratch sheet is active
        mTST_DP_ScratchSheet.Activate
    'Enable grid icon for this suite
        gDP_ShowGridIcon = True
    'Remove stale grid icons before the suite
        M_GridIcon_PurgeAll

'------------------------------------------------------------------------------
' EMBEDDED ICON FILE
'------------------------------------------------------------------------------
    'Resolve the embedded icon file path
        IconPath = M_GridIcon_EnsureEmbeddedIconFile()
    'Assert the embedded icon path is not blank
        TST_DP_AssertTrue "Embedded grid icon path is not blank", _
            VBA.LenB(IconPath) > 0
    'Assert the embedded icon file exists
        TST_DP_AssertTrue "Embedded grid icon file exists", _
            VBA.LenB(VBA.Dir$(IconPath, vbNormal)) > 0

'------------------------------------------------------------------------------
' CREATE ICON
'------------------------------------------------------------------------------
    'Create or move the grid icon beside B2
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
    'Assert the grid icon exists on the scratch sheet
        TST_DP_AssertTrue "Grid icon is created on eligible target", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert the grid icon is visible after create
        TST_DP_AssertTrue "Grid icon is visible after create", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert only one named grid icon exists in the host workbook
        TST_DP_AssertEqualsLong "Only one grid icon exists after create", _
            1, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    'Capture the initial icon left position
        ShapeLeftBefore = mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Left
    'Capture the initial icon top position
        ShapeTopBefore = mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Top

'------------------------------------------------------------------------------
' MOVE ICON
'------------------------------------------------------------------------------
    'Move the grid icon beside a different cell
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("F8")
    'Assert the grid icon still exists after move
        TST_DP_AssertTrue "Grid icon still exists after move", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert the grid icon is visible after move
        TST_DP_AssertTrue "Grid icon is visible after move", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Assert only one named grid icon exists after move
        TST_DP_AssertEqualsLong "Only one grid icon exists after move", _
            1, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    'Assert the icon moved horizontally or vertically
        TST_DP_AssertTrue "Grid icon position changes after move", _
            (mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Left <> ShapeLeftBefore) Or _
            (mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Top <> ShapeTopBefore)

'------------------------------------------------------------------------------
' REMOVE AND PURGE ICON
'------------------------------------------------------------------------------
    'Remove the current grid icon
        M_GridIcon_Remove
    'Assert the grid icon no longer exists on the active scratch sheet
        TST_DP_AssertFalse "Grid icon is removed from active scratch sheet", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)
    'Create the icon again for purge testing
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
    'Purge all named grid icons
        M_GridIcon_PurgeAll
    'Assert no named grid icon remains in the host workbook
        TST_DP_AssertEqualsLong "Purge removes all grid icons", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' DISABLED FEATURE BEHAVIOR
'------------------------------------------------------------------------------
    'Disable grid icon feature directly
        gDP_ShowGridIcon = False
    'Attempt to show icon while disabled
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
    'Assert disabled feature does not leave an icon
        TST_DP_AssertFalse "Grid icon is not shown when feature is disabled", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the suite-level failure
        TST_DP_RecordFail "Grid icon suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    'Clear the suite error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_Manager()

'
'==============================================================================
'                           RUN MANAGER SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates public cDatePickerManager behavior and target gating
'
' WHY THIS EXISTS
'   The manager is the central Application-event coordinator and should make
'   deterministic UI decisions for explicit target cells
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests manager creation, picker state predicates, Should_ShowGridIcon gating,
'   explicit selection-change handling, and reset behavior
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   cDatePickerManager
'   M_GridIcon_PurgeAll
'
' NOTES
'   The ineligible-target assertion checks that no visible grid icon remains
'
'   This allows either implementation:
'     - remove icon on ineligible selection
'     - hide icon on ineligible selection for faster reuse
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Manager                 As cDatePickerManager   'Manager instance under test
    Dim MergedArea              As Excel.Range          'Merged area under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "Manager"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Create a fresh manager instance for public API checks
        Set Manager = New cDatePickerManager

    'Ensure the scratch sheet is active
        mTST_DP_ScratchSheet.Activate

    'Remove stale grid icons before manager tests
        M_GridIcon_PurgeAll

    'Enable grid icon gating directly
        gDP_ShowGridIcon = True

'------------------------------------------------------------------------------
' STATE PREDICATES
'------------------------------------------------------------------------------
    'Assert new manager is not busy
        TST_DP_AssertFalse "New manager is not busy", Manager.Is_Busy

    'Close picker before loaded / visible checks
        DP_Close

    'Assert picker is not visible after close
        TST_DP_AssertFalse "PickerVisible is False after DP_Close", Manager.PickerVisible

    'Assert picker is not loaded after close
        TST_DP_AssertFalse "Is_PickerLoaded is False after DP_Close", Manager.Is_PickerLoaded

'------------------------------------------------------------------------------
' SHOULD_SHOWGRIDICON GATING
'------------------------------------------------------------------------------
    'Prepare a date value cell
        mTST_DP_ScratchSheet.Range("D10").Value = VBA.DateSerial(2026, 5, 3)

    'Assert date value cell is eligible
        TST_DP_AssertTrue "Should_ShowGridIcon accepts explicit date value cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D10"))

    'Prepare a date-formatted blank cell
        mTST_DP_ScratchSheet.Range("B11").ClearContents
        mTST_DP_ScratchSheet.Range("B11").NumberFormat = "dd/mm/yyyy"

    'Assert date-formatted blank cell is eligible
        TST_DP_AssertTrue "Should_ShowGridIcon accepts date-formatted blank cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("B11"))

    'Prepare a general blank cell
        mTST_DP_ScratchSheet.Range("D12").ClearContents
        mTST_DP_ScratchSheet.Range("D12").NumberFormat = "General"

    'Assert general blank cell is not eligible
        TST_DP_AssertFalse "Should_ShowGridIcon rejects general blank cell", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D12"))

    'Assert multi-cell range is not eligible at manager-gating level
        TST_DP_AssertFalse "Should_ShowGridIcon rejects multi-cell range", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D10:D11"))

'------------------------------------------------------------------------------
' MERGED-CELL NORMALIZATION EXPECTATION
'------------------------------------------------------------------------------
    'Prepare the merged area
        Set MergedArea = mTST_DP_ScratchSheet.Range("D14:E15")

    'Clear the merged area
        MergedArea.Clear

    'Merge the test area
        MergedArea.Merge

    'Format the merged area as date-like
        MergedArea.NumberFormat = "dd/mm/yyyy"

    'Assert the top-left cell of a merged area is eligible
        TST_DP_AssertTrue "Should_ShowGridIcon accepts top-left cell inside merged area", _
            Manager.Should_ShowGridIcon(mTST_DP_ScratchSheet.Range("D14"))

    'Unmerge the area after the assertion
        MergedArea.UnMerge

'------------------------------------------------------------------------------
' SELECTION-CHANGE HANDLING
'------------------------------------------------------------------------------
    'Prepare a date value cell for selection-change handling
        mTST_DP_ScratchSheet.Range("E10").Value = VBA.DateSerial(2026, 9, 9)

    'Handle selection change for an eligible target
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E10")

    'Assert the manager created or showed the grid icon for the eligible target
        TST_DP_AssertTrue "Handle_SelectionChange shows visible icon for eligible target", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Prepare an ineligible target
        mTST_DP_ScratchSheet.Range("E11").ClearContents
        mTST_DP_ScratchSheet.Range("E11").NumberFormat = "General"

    'Handle selection change for an ineligible target
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E11")

    'Assert no visible grid icon remains for the ineligible target
        TST_DP_AssertFalse "Handle_SelectionChange leaves no visible icon for ineligible target", _
            TST_DP_ShapeIsVisible(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' RESET BEHAVIOR
'------------------------------------------------------------------------------
    'Show the icon again for reset testing
        Manager.Handle_SelectionChange mTST_DP_ScratchSheet.Range("E10")

    'Reset all DatePicker UI through the manager
        Manager.Reset_DatePickerUI

    'Assert reset removed all grid icons
        TST_DP_AssertEqualsLong "Reset_DatePickerUI purges all grid icons", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set MergedArea = Nothing

    'Release the test manager
        Set Manager = Nothing

    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Suppress local cleanup errors
        On Error Resume Next

    'Unmerge the merged area when needed
        If Not MergedArea Is Nothing Then
            MergedArea.UnMerge
        End If

    'Release object references
        Set MergedArea = Nothing
        Set Manager = Nothing

    'Record the suite-level failure
        TST_DP_RecordFail "Manager suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub TST_DP_RunSuite_UISmoke()

'
'==============================================================================
'                           RUN UI SMOKE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Opens and closes UF_DatePicker as a minimal visual lifecycle smoke test
'
' WHY THIS EXISTS
'   Some integration failures only appear when the runtime form is loaded and
'   initialized
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Opens the picker for a date cell, validates loaded / visible state, closes
'   the picker, and validates that the form is no longer visible
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   DP_Show
'   DP_Close
'   cDatePickerManager state predicates
'
' NOTES
'   This suite briefly shows the UserForm and should be run manually before
'   release, not necessarily on every development edit
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "UISmoke"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Ensure the scratch sheet is active
        mTST_DP_ScratchSheet.Activate

    'Prepare the active cell with a deterministic date
        mTST_DP_ScratchSheet.Range("H2").Value = VBA.DateSerial(2026, 5, 3)

    'Select the date cell
        mTST_DP_ScratchSheet.Range("H2").Select

    'Ensure manager infrastructure exists
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' OPEN FORM
'------------------------------------------------------------------------------
    'Open the DatePicker form
        DP_Show

    'Allow modeless form events to process
        DoEvents

    'Assert the picker is loaded
        TST_DP_AssertTrue "DP_Show loads the picker", gDP_Manager.Is_PickerLoaded

    'Assert the picker is visible
        TST_DP_AssertTrue "DP_Show shows the picker", gDP_Manager.PickerVisible

'------------------------------------------------------------------------------
' CLOSE FORM
'------------------------------------------------------------------------------
    'Close the DatePicker form
        DP_Close

    'Allow modeless form events to process
        DoEvents

    'Assert the picker is no longer visible
        TST_DP_AssertFalse "DP_Close hides the picker", gDP_Manager.PickerVisible

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite succeeds
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Suppress close failures
        On Error Resume Next

    'Attempt to close the picker after a UI smoke failure
        DP_Close

    'Record the suite-level failure
        TST_DP_RecordFail "UI smoke suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                            EXPECTED-ERROR TESTS
'
'------------------------------------------------------------------------------

Private Sub TST_DP_ExpectError_FirstDayToTextInvalid()

'
'==============================================================================
'                     EXPECT ERROR FIRST DAY TO TEXT INVALID
'==============================================================================

    On Error GoTo ExpectedError

    M_Settings_FirstDayOfWeekToText 0

    TST_DP_RecordFail "Invalid first-day conversion raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Invalid first-day conversion raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetFirstDayBlank()

'
'==============================================================================
'                       EXPECT ERROR SET FIRST DAY BLANK
'==============================================================================

    On Error GoTo ExpectedError

    M_Settings_SetFirstDayOfWeekText vbNullString

    TST_DP_RecordFail "Blank first-day text raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Blank first-day text raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetFirstDayInvalidText()

'
'==============================================================================
'                    EXPECT ERROR SET FIRST DAY INVALID TEXT
'==============================================================================

    On Error GoTo ExpectedError

    M_Settings_SetFirstDayOfWeekText "Wednesday"

    TST_DP_RecordFail "Unsupported first-day text raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Unsupported first-day text raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetInvalidClockMode()

'
'==============================================================================
'                       EXPECT ERROR INVALID CLOCK MODE
'==============================================================================

    On Error GoTo ExpectedError

    M_Settings_SetClockMode 99

    TST_DP_RecordFail "Unsupported clock mode raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Unsupported clock mode raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_SetInvalidSizeMode()

'
'==============================================================================
'                        EXPECT ERROR INVALID SIZE MODE
'==============================================================================

    On Error GoTo ExpectedError

    M_Settings_SetSizeMode 99

    TST_DP_RecordFail "Unsupported size mode raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Unsupported size mode raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_EnglishMonthShortInvalid()

'
'==============================================================================
'                   EXPECT ERROR ENGLISH MONTH SHORT INVALID
'==============================================================================

    On Error GoTo ExpectedError

    M_Caption_GetEnglishMonthShort 13

    TST_DP_RecordFail "Invalid short English month raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Invalid short English month raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_EnglishMonthFullInvalid()

'
'==============================================================================
'                    EXPECT ERROR ENGLISH MONTH FULL INVALID
'==============================================================================

    On Error GoTo ExpectedError

    M_Caption_GetEnglishMonthFull 0

    TST_DP_RecordFail "Invalid full English month raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Invalid full English month raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_GetMonthInvalid()

'
'==============================================================================
'                         EXPECT ERROR GET MONTH INVALID
'==============================================================================

    On Error GoTo ExpectedError

    M_Caption_GetMonth 99, False

    TST_DP_RecordFail "Invalid GetMonth input raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Invalid GetMonth input raises", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectError_InvalidWriteAction()

'
'==============================================================================
'                       EXPECT ERROR INVALID WRITE ACTION
'==============================================================================

    On Error GoTo ExpectedError

    M_WriteBack_PopulateRange mTST_DP_ScratchSheet.Range("J2"), 99

    TST_DP_RecordFail "Unsupported write action raises", "No error was raised"

    Exit Sub

ExpectedError:
    TST_DP_RecordPass "Unsupported write action raises", Err.Description
    Err.Clear

End Sub

'
'------------------------------------------------------------------------------
'
'                              ASSERTION HELPERS
'
'------------------------------------------------------------------------------

Private Sub TST_DP_AssertTrue( _
    ByVal TestName As String, _
    ByVal Condition As Boolean)

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

Private Sub TST_DP_AssertFalse( _
    ByVal TestName As String, _
    ByVal Condition As Boolean)

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

Private Sub TST_DP_AssertBooleanResult( _
    ByVal TestName As String, _
    ByVal BooleanValue As Boolean)

'
'==============================================================================
'                           ASSERT BOOLEAN RESULT
'==============================================================================

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

    If TargetCell Is Nothing Then
        TST_DP_RecordFail TestName, "TargetCell is Nothing"
        Exit Sub
    End If

    If Not VBA.IsDate(TargetCell.Value) Then
        TST_DP_RecordFail TestName, _
            "Cell value is not date-like: " & VBA.CStr(TargetCell.Value)
        Exit Sub
    End If

    TST_DP_AssertDateEquals TestName, ExpectedDate, VBA.CDate(TargetCell.Value)

End Sub

'
'------------------------------------------------------------------------------
'
'                                RESULT HELPERS
'
'------------------------------------------------------------------------------

Private Sub TST_DP_RecordPass( _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD PASS
'==============================================================================

    TST_DP_RecordResult TST_DP_PASS_TEXT, mTST_DP_CurrentSuite, TestName, Details

End Sub

Private Sub TST_DP_RecordFail( _
    ByVal TestName As String, _
    ByVal Details As String)

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
'                             RECORD RESULT
'------------------------------------------------------------------------------
' PURPOSE
'   Records one regression result row
'
' WHY THIS EXISTS
'   Result recording should never become the reason a regression run aborts
'
' INPUTS
'   ResultText
'     PASS, FAIL, or INFO marker
'
'   SuiteName
'     Logical suite name
'
'   TestName
'     Test or diagnostic name
'
'   Details
'     Optional detail text
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Updates counters, writes to the Immediate Window, and writes to the result
'   worksheet when available
'
' ERROR POLICY
'   Best-effort
'   Worksheet-write failures are reported to the Immediate Window and do not
'   raise outward
'
' DEPENDENCIES
'   mTST_DP_ResultSheet
'
' NOTES
'   This routine intentionally avoids recursive failure recording
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_RecordResult" 'Current procedure name

    Dim ResultLine              As String       'Immediate Window result line
    Dim RecordErrorNumber       As Long         'Captured result-write error number
    Dim RecordErrorDescription  As String       'Captured result-write error description

'------------------------------------------------------------------------------
' UPDATE COUNTERS
'------------------------------------------------------------------------------
    If ResultText = TST_DP_PASS_TEXT Or ResultText = TST_DP_FAIL_TEXT Then
        mTST_DP_RunCount = mTST_DP_RunCount + 1
    End If
    If ResultText = TST_DP_PASS_TEXT Then
        mTST_DP_PassCount = mTST_DP_PassCount + 1
    End If
    If ResultText = TST_DP_FAIL_TEXT Then
        mTST_DP_FailCount = mTST_DP_FailCount + 1
    End If

'------------------------------------------------------------------------------
' WRITE IMMEDIATE WINDOW
'------------------------------------------------------------------------------
    ResultLine = ResultText & _
        " | " & SuiteName & _
        " | " & TestName & _
        IIf(VBA.Len(Details) > 0, " | " & Details, vbNullString)

    Debug.Print ResultLine

'------------------------------------------------------------------------------
' WRITE RESULT SHEET
'------------------------------------------------------------------------------
    If mTST_DP_ResultSheet Is Nothing Then Exit Sub

    On Error GoTo ResultSheetFail

    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 3).Value = mTST_DP_NextResultRow - 4
    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 4).Value = VBA.Now
    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 5).Value = ResultText
    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 6).Value = SuiteName
    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 7).Value = TestName
    mTST_DP_ResultSheet.Cells(mTST_DP_NextResultRow, 8).Value = Details

    mTST_DP_NextResultRow = mTST_DP_NextResultRow + 1

    On Error GoTo 0

    Exit Sub

ResultSheetFail:
    RecordErrorNumber = Err.Number
    RecordErrorDescription = Err.Description

    Debug.Print TST_DP_INFO_TEXT & _
        " | Harness | " & PROC_NAME & _
        " | Result-sheet write skipped after error " & VBA.CStr(RecordErrorNumber) & _
        " - " & RecordErrorDescription

    Set mTST_DP_ResultSheet = Nothing

    Err.Clear
    On Error GoTo 0

End Sub

Private Sub TST_DP_WriteSummary()

'
'==============================================================================
'                             WRITE SUMMARY
'==============================================================================

    TST_DP_RecordInfo "Harness", _
        "Summary", _
        "Run=" & VBA.CStr(mTST_DP_RunCount) & _
        "; Passed=" & VBA.CStr(mTST_DP_PassCount) & _
        "; Failed=" & VBA.CStr(mTST_DP_FailCount)

    If mTST_DP_ResultSheet Is Nothing Then
        Exit Sub
    End If

    mTST_DP_ResultSheet.Range("J4").Value = "SUMMARY"
    mTST_DP_ResultSheet.Range("J5").Value = "Run"
    mTST_DP_ResultSheet.Range("K5").Value = mTST_DP_RunCount
    mTST_DP_ResultSheet.Range("J6").Value = "Passed"
    mTST_DP_ResultSheet.Range("K6").Value = mTST_DP_PassCount
    mTST_DP_ResultSheet.Range("J7").Value = "Failed"
    mTST_DP_ResultSheet.Range("K7").Value = mTST_DP_FailCount
    mTST_DP_ResultSheet.Range("J4:J7").Font.Bold = True
    mTST_DP_ResultSheet.Columns("C:K").AutoFit

End Sub

'
'------------------------------------------------------------------------------
'
'                             WORKBOOK HELPERS
'
'------------------------------------------------------------------------------

Private Function TST_DP_GetHostWorkbook() As Excel.Workbook

'
'==============================================================================
'                             GET HOST WORKBOOK
'==============================================================================
    On Error Resume Next
    If Not Excel.Application.ActiveWorkbook Is Nothing Then
        Set TST_DP_GetHostWorkbook = Excel.Application.ActiveWorkbook
    End If
    If TST_DP_GetHostWorkbook Is Nothing Then
        Set TST_DP_GetHostWorkbook = ThisWorkbook
    End If

    Err.Clear
    On Error GoTo 0

End Function

Private Sub TST_DP_PrepareResultSheet(ByVal HostWorkbook As Excel.Workbook)

'
'==============================================================================
'                           PREPARE RESULT SHEET
'==============================================================================


    Set mTST_DP_ResultSheet = HostWorkbook.Worksheets(TST_DP_RESULT_SHEET_NAME)

    mTST_DP_ResultSheet.Name = TST_DP_RESULT_SHEET_NAME

    mTST_DP_ResultSheet.Range("C4:H4").Value = Array( _
        "#", _
        "Timestamp", _
        "Result", _
        "Suite", _
        "Test", _
        "Details")

    mTST_DP_ResultSheet.Range("C4:H4").Font.Bold = True

    mTST_DP_ResultSheet.Activate
    mTST_DP_ResultSheet.Range("C5").Select

    mTST_DP_NextResultRow = 5

End Sub

Private Sub TST_DP_PrepareScratchSheet(ByVal HostWorkbook As Excel.Workbook)

'
'==============================================================================
'                           PREPARE SCRATCH SHEET
'==============================================================================

    TST_DP_DeleteWorksheetIfExists HostWorkbook, TST_DP_SCRATCH_SHEET_NAME

    Set mTST_DP_ScratchSheet = HostWorkbook.Worksheets.Add( _
        After:=HostWorkbook.Worksheets(HostWorkbook.Worksheets.Count))

    mTST_DP_ScratchSheet.Name = TST_DP_SCRATCH_SHEET_NAME
    mTST_DP_ScratchSheet.Range("A1").Value = "DatePicker regression scratch sheet"
    mTST_DP_ScratchSheet.Columns("A:J").ColumnWidth = 16
    mTST_DP_ScratchSheet.Activate

End Sub

Private Sub TST_DP_DeleteScratchSheet()

'
'==============================================================================
'                            DELETE SCRATCH SHEET
'==============================================================================

    If Not mTST_DP_HostWorkbook Is Nothing Then
        TST_DP_DeleteWorksheetIfExists mTST_DP_HostWorkbook, TST_DP_SCRATCH_SHEET_NAME
    End If

    Set mTST_DP_ScratchSheet = Nothing

End Sub

Private Sub TST_DP_DeleteWorksheetIfExists( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal SheetName As String)

'
'==============================================================================
'                        DELETE WORKSHEET IF EXISTS
'==============================================================================

    Dim TargetSheet             As Excel.Worksheet     'Worksheet to delete

    If HostWorkbook Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next

    Set TargetSheet = HostWorkbook.Worksheets(SheetName)

    If Not TargetSheet Is Nothing Then
        If HostWorkbook.Worksheets.Count > 1 Then
            TargetSheet.Delete
        End If
    End If

    Set TargetSheet = Nothing

    Err.Clear
    On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                              STATE HELPERS
'
'------------------------------------------------------------------------------

Private Sub TST_DP_CaptureSettings(ByRef Snapshot As TRegDPSettingsSnapshot)

'
'==============================================================================
'                            CAPTURE SETTINGS
'==============================================================================

    M_Settings_EnsureLoaded

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
'==============================================================================

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

    M_Settings_Save
    M_ContextMenu_Update
    M_KeyboardShortcut_Update

    If Not gDP_ShowGridIcon Then
        M_GridIcon_Remove
    End If

End Sub

Private Sub TST_DP_CaptureApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                         CAPTURE APPLICATION STATE
'==============================================================================

    Snapshot.ScreenUpdating = Excel.Application.ScreenUpdating
    Snapshot.EnableEvents = Excel.Application.EnableEvents
    Snapshot.DisplayAlerts = Excel.Application.DisplayAlerts
    Snapshot.CalculationMode = Excel.Application.Calculation

    Snapshot.StatusBarWasFalse = (VBA.VarType(Excel.Application.StatusBar) = vbBoolean)

    If Not Snapshot.StatusBarWasFalse Then
        Snapshot.StatusBarText = VBA.CStr(Excel.Application.StatusBar)
    End If

End Sub

Private Sub TST_DP_PrepareApplicationForRun()

'
'==============================================================================
'                       PREPARE APPLICATION FOR RUN
'==============================================================================

    Excel.Application.ScreenUpdating = False
    Excel.Application.EnableEvents = False
    Excel.Application.DisplayAlerts = False
    Excel.Application.StatusBar = "Running DatePicker regression tests..."

End Sub

Private Sub TST_DP_RestoreApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                       RESTORE APPLICATION STATE
'==============================================================================

    Excel.Application.Calculation = Snapshot.CalculationMode
    Excel.Application.DisplayAlerts = Snapshot.DisplayAlerts
    Excel.Application.EnableEvents = Snapshot.EnableEvents
    Excel.Application.ScreenUpdating = Snapshot.ScreenUpdating

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
'==============================================================================

    On Error Resume Next

    M_Timer_Stop
    DP_Close
    M_ContextMenu_Remove
    M_KeyboardShortcut_Remove
    M_GridIcon_PurgeAll

    Err.Clear
    On Error GoTo 0

End Sub

Private Sub TST_DP_RestoreManagerState()

'
'==============================================================================
'                         RESTORE MANAGER STATE
'==============================================================================

    Set gDP_Manager = Nothing

    If mTST_DP_HadManager Then
        M_Picker_EnsureManager
    End If

End Sub

'
'------------------------------------------------------------------------------
'
'                              OBJECT HELPERS
'
'------------------------------------------------------------------------------

Private Function TST_DP_ShapeExists( _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal ShapeName As String) As Boolean

'
'==============================================================================
'                              SHAPE EXISTS
'==============================================================================

    Dim TargetShape             As Excel.Shape        'Resolved shape

    TST_DP_ShapeExists = False

    If TargetSheet Is Nothing Then
        Exit Function
    End If

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
'==============================================================================

    Dim TargetShape             As Excel.Shape        'Resolved shape

    TST_DP_ShapeIsVisible = False

    If TargetSheet Is Nothing Then
        Exit Function
    End If

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
'==============================================================================

    Dim TargetSheet             As Excel.Worksheet     'Worksheet being inspected
    Dim TargetShape             As Excel.Shape         'Shape being inspected
    Dim ShapeCount              As Long                'Matched shape count

    ShapeCount = 0

    If HostWorkbook Is Nothing Then
        TST_DP_CountNamedShapes = 0
        Exit Function
    End If

    For Each TargetSheet In HostWorkbook.Worksheets
        For Each TargetShape In TargetSheet.Shapes
            If VBA.StrComp(TargetShape.Name, ShapeName, vbBinaryCompare) = 0 Then
                ShapeCount = ShapeCount + 1
            End If
        Next TargetShape
    Next TargetSheet

    TST_DP_CountNamedShapes = ShapeCount

    Set TargetShape = Nothing
    Set TargetSheet = Nothing

End Function

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
'   Application.Run
'
'   Callback resolution can vary depending on workbook qualification, module
'   qualification, project state, and host workbook context
'
'   This helper probes supported callback-name formats and returns the first one
'   that Excel can actually execute
'
' INPUTS
'   MacroName
'     Public callback procedure name
'
' RETURNS
'   First runnable callback reference
'
' BEHAVIOR
'   Trims the supplied macro name
'   Builds workbook + module qualified, workbook-qualified, and unqualified
'   callback-name candidates
'   Tests each candidate through Excel.Application.Run
'   Returns the first candidate that executes without raising a runtime error
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
'   date while resolving the runnable name
'
'   The regression callbacks used here are side-effect free
'
' UPDATED
'   2026-05-10
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "TST_DP_QualifiedMacroName"

    Dim NormalizedMacroName     As String       'Trimmed callback macro name
    Dim WorkbookQualifier       As String       'Escaped workbook qualifier
    Dim CandidateNames(1 To 3)  As String       'Candidate callback names
    Dim CandidateIndex          As Long         'Candidate loop index
    Dim CallbackResult          As Variant      'Callback probe result
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
    'Build the workbook qualifier
        WorkbookQualifier = "'" & VBA.Replace(ThisWorkbook.Name, "'", "''") & "'!"
    'Build workbook and module qualified callback name
        CandidateNames(1) = WorkbookQualifier & TST_DP_MODULE_NAME & "." & NormalizedMacroName
    'Build workbook-qualified callback name
        CandidateNames(2) = WorkbookQualifier & NormalizedMacroName
    'Build unqualified callback name
        CandidateNames(3) = NormalizedMacroName

'------------------------------------------------------------------------------
' PROBE CANDIDATES
'------------------------------------------------------------------------------
    'Loop through candidate callback names
        For CandidateIndex = LBound(CandidateNames) To UBound(CandidateNames)
            'Suppress candidate probe errors
                On Error Resume Next
            'Clear any pending error before probing the candidate
                Err.Clear
            'Run the candidate callback with a deterministic date
                CallbackResult = Excel.Application.Run( _
                    CandidateNames(CandidateIndex), _
                    VBA.DateSerial(2026, 1, 1))
            'Capture probe error number
                RunErrNumber = Err.Number
            'Capture probe error description
                RunErrDescription = Err.Description
            'Clear the suppressed probe error
                Err.Clear
            'Restore controlled error handling
                On Error GoTo ErrorHandler
            'Return the first runnable callback name
                If RunErrNumber = 0 Then
                    TST_DP_QualifiedMacroName = CandidateNames(CandidateIndex)
                    Exit Function
                End If
        Next CandidateIndex

'------------------------------------------------------------------------------
' RAISE RESOLUTION FAILURE
'------------------------------------------------------------------------------
    'Raise when no candidate could be executed
        Err.Raise vbObjectError + 514, PROC_NAME, _
            "Unable to resolve runnable callback macro name for '" & _
            NormalizedMacroName & "'. Last Application.Run error: " & _
            VBA.CStr(RunErrNumber) & " - " & RunErrDescription

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
            "Regression callback-name resolution failed: " & Err.Description

End Function
Public Sub TST_DP_ApplyFailConditionalFormat( _
    ByVal WS As Excel.Worksheet, _
    ByVal TargetRange As Excel.Range)

'
'------------------------------------------------------------------------------
'                       APPLY FAIL CONDITIONAL FORMAT
'------------------------------------------------------------------------------
' PURPOSE
'   Applies FAIL conditional formatting to a supplied worksheet range
'
' WHY THIS EXISTS
'   Conditional formatting should be applied directly to an explicit target range
'   instead of relying on Select, Selection, ActiveCell, or recorded-macro state
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
'   Validates the supplied worksheet and range
'   Validates that the range belongs to the supplied worksheet
'   Adds a text-contains FAIL conditional-formatting rule
'   Applies bold white font and dark red fill
'
' ERROR POLICY
'   Raises a descriptive runtime error if inputs are missing, inconsistent, or
'   conditional formatting cannot be applied
'
' DEPENDENCIES
'   Excel object model
'
' NOTES
'   This routine does not clear existing conditional-formatting rules
'
'   This routine does not select or activate any worksheet or range
'
' UPDATED
'   2026-05-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ApplyFailConditionalFormat"
    Const FAIL_TEXT             As String = "FAIL"
    Const FAIL_BACK_COLOR       As Long = 192

    Dim FailCondition           As Excel.FormatCondition    'Created conditional-formatting rule
    Dim ErrorNumber             As Long                     'Captured error number
    Dim ErrorDescription        As String                   'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject missing worksheet references
        If WS Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "WS cannot be Nothing"
        End If

    'Reject missing target ranges
        If TargetRange Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "TargetRange cannot be Nothing"
        End If

    'Reject ranges that do not belong to the supplied worksheet
        If Not TargetRange.Worksheet Is WS Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "TargetRange must belong to the supplied worksheet"
        End If

'------------------------------------------------------------------------------
' ADD CONDITIONAL FORMAT
'------------------------------------------------------------------------------
    'Add the FAIL text conditional-formatting rule
        Set FailCondition = TargetRange.FormatConditions.Add( _
            Type:=xlTextString, _
            String:=FAIL_TEXT, _
            TextOperator:=xlContains)

    'Move the new rule to first priority
        FailCondition.SetFirstPriority

'------------------------------------------------------------------------------
' APPLY FORMAT SETTINGS
'------------------------------------------------------------------------------
    With FailCondition.Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    With FailCondition.Interior
        .PatternColorIndex = xlAutomatic
        .Color = FAIL_BACK_COLOR
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Release object references
        Set FailCondition = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "FAIL conditional-format application failed: " & ErrorDescription

End Sub
