Attribute VB_Name = "M_DatePicker_Test"

Option Explicit

'
'==============================================================================
' MODULE: M_REGRESSION_DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a detailed regression harness for the VBA DatePicker project
'
' WHY THIS EXISTS
'   The DatePicker spans persisted settings, manager event orchestration,
'   worksheet write-back, UserForm bridge state, optional context-menu access,
'   optional keyboard shortcut access, and optional in-grid worksheet shapes
'
'   A single regression module makes it easier to validate those behaviors after
'   refactoring the manager, form bridge, label hooks, settings, and grid-icon
'   infrastructure
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
'     - a worksheet named REG_DP_RESULTS
'
'   Creates and deletes a scratch worksheet named REG_DP_SCRATCH
'
'   Captures and restores DatePicker settings and selected Application state
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
'   The uploaded 2026-05-03 file set still contains two compile blockers:
'     - cDatePickerLabelHook.Initialize declares NormalizedHoverMode twice
'     - cDatePickerManager.mExcelApp_WorkbookBeforeClose contains an extra End If
'
'   REG_DP_RunAll excludes disruptive UI smoke tests by default
'
'   REG_DP_RunAll_WithUISmoke opens the DatePicker form briefly and closes it
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const REG_DP_RESULT_SHEET_NAME       As String = "REG_DP_RESULTS"      'Regression result worksheet name
    Private Const REG_DP_SCRATCH_SHEET_NAME      As String = "REG_DP_SCRATCH"      'Regression scratch worksheet name
    Private Const REG_DP_PASS_TEXT               As String = "PASS"                'Passed test marker
    Private Const REG_DP_FAIL_TEXT               As String = "FAIL"                'Failed test marker
    Private Const REG_DP_INFO_TEXT               As String = "INFO"                'Information marker

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
        IconPath                   As String     'Grid icon path setting
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
    Private mREG_DP_ResultSheet     As Worksheet  'Result worksheet used by the current run
    Private mREG_DP_ScratchSheet    As Worksheet  'Scratch worksheet used by the current run
    Private mREG_DP_NextResultRow   As Long       'Next result row
    Private mREG_DP_RunCount        As Long       'Total assertions executed
    Private mREG_DP_PassCount       As Long       'Total assertions passed
    Private mREG_DP_FailCount       As Long       'Total assertions failed
    Private mREG_DP_CurrentSuite    As String     'Current suite name
    Private mREG_DP_HostWorkbook    As Workbook   'Workbook receiving test sheets
    Private mREG_DP_HadManager      As Boolean    'True when a manager existed before the run

'
'==============================================================================
'
'                              PUBLIC ENTRY POINTS
'
'==============================================================================

Public Sub REG_DP_RunAll()

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
'   Runs core, settings, policy, caption, write-back, grid-icon, manager, and
'   bridge tests without opening the DatePicker UserForm
'
' ERROR POLICY
'   Delegates fatal handling and cleanup to REG_DP_RunAllInternal
'
' DEPENDENCIES
'   REG_DP_RunAllInternal
'
' NOTES
'   Use REG_DP_RunAll_WithUISmoke when you also want a brief UserForm open/close
'   smoke test
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the standard non-disruptive regression pack
        REG_DP_RunAllInternal False

End Sub

Public Sub REG_DP_RunAll_WithUISmoke()

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
'   Delegates fatal handling and cleanup to REG_DP_RunAllInternal
'
' DEPENDENCIES
'   REG_DP_RunAllInternal
'
' NOTES
'   This routine is intentionally visible as a separate macro
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Run the regression pack with the optional UI smoke suite
        REG_DP_RunAllInternal True

End Sub

Public Function REG_DP_HolidayCallback(ByVal CandidateDate As Date) As Boolean

'
'==============================================================================
'                           TEST HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a deterministic Boolean holiday callback for regression tests
'
' WHY THIS EXISTS
'   M_HolidayPolicy_IsHolidayDate calls user callbacks through Application.Run
'
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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return True only for the deterministic holiday date
        REG_DP_HolidayCallback = _
            (VBA.DateValue(CandidateDate) = VBA.DateSerial(2026, 1, 1))

End Function

Public Function REG_DP_HolidayCallbackNonBoolean(ByVal CandidateDate As Date) As String

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return a non-Boolean value deliberately
        REG_DP_HolidayCallbackNonBoolean = "TRUE"

End Function

Public Function REG_DP_HolidayCallbackError(ByVal CandidateDate As Date) As Variant

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
'   This callback deliberately avoids Err.Raise because some VBE error-trapping
'   settings stop execution before the caller's fail-safe policy can continue
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RETURN ERROR VALUE
'------------------------------------------------------------------------------
    'Return an Excel error value deliberately
        REG_DP_HolidayCallbackError = CVErr(xlErrValue)

End Function

'
'==============================================================================
'
'                               RUN ORCHESTRATION
'
'==============================================================================

Private Sub REG_DP_RunAllInternal(ByVal IncludeUISmoke As Boolean)

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
'   All REG_DP_RunSuite_* routines
'
' NOTES
'   This routine intentionally disables Application events during the run to
'   reduce interference from workbook-level event handlers
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "REG_DP_RunAllInternal"

    Dim SettingsSnapshot        As TRegDPSettingsSnapshot   'DatePicker settings snapshot
    Dim AppSnapshot             As TRegDPApplicationSnapshot 'Excel application snapshot
    Dim FatalNumber             As Long                     'Fatal error number
    Dim FatalDescription        As String                   'Fatal error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled fatal handling
        On Error GoTo FatalHandler

    'Reset counters and module state
        REG_DP_ResetHarnessState

    'Resolve the workbook that will receive regression sheets
        Set mREG_DP_HostWorkbook = REG_DP_GetHostWorkbook()

    'Capture whether a manager existed before the run
        mREG_DP_HadManager = Not (gDP_Manager Is Nothing)

    'Capture current DatePicker settings and transient state
        REG_DP_CaptureSettings SettingsSnapshot

    'Capture current Application state
        REG_DP_CaptureApplicationState AppSnapshot

'------------------------------------------------------------------------------
' PREPARE ISOLATED RUN STATE
'------------------------------------------------------------------------------
    'Prepare the Application state used by the harness
        REG_DP_PrepareApplicationForRun

    'Reset DatePicker UI artifacts before testing
        REG_DP_ResetDatePickerArtifacts

    'Prepare the result worksheet
        REG_DP_PrepareResultSheet mREG_DP_HostWorkbook

    'Prepare the scratch worksheet
        REG_DP_PrepareScratchSheet mREG_DP_HostWorkbook

    'Record the run header
        REG_DP_RecordInfo "Harness", "Start", _
            "IncludeUISmoke=" & VBA.CStr(IncludeUISmoke)

'------------------------------------------------------------------------------
' RUN SUITES
'------------------------------------------------------------------------------
    'Run environment and manager smoke checks through the defensive suite runner
        REG_DP_RunSuiteSafe "Environment"

    'Run settings and persisted-state checks through the defensive suite runner
        REG_DP_RunSuiteSafe "Settings"

    'Run date-policy checks through the defensive suite runner
        REG_DP_RunSuiteSafe "DatePolicy"

    'Run holiday-policy callback checks through the defensive suite runner
        REG_DP_RunSuiteSafe "HolidayPolicy"

    'Run caption helper checks through the defensive suite runner
        REG_DP_RunSuiteSafe "Captions"

    'Run form bridge checks through the defensive suite runner
        REG_DP_RunSuiteSafe "FormBridge"

    'Run worksheet write-back checks through the defensive suite runner
        REG_DP_RunSuiteSafe "WriteBack"

    'Run grid-icon shape checks through the defensive suite runner
        REG_DP_RunSuiteSafe "GridIcon"

    'Run manager public API checks through the defensive suite runner
        REG_DP_RunSuiteSafe "Manager"

    'Run optional UserForm smoke checks through the defensive suite runner when requested
        If IncludeUISmoke Then REG_DP_RunSuiteSafe "UISmoke"

'------------------------------------------------------------------------------
' WRITE SUMMARY
'------------------------------------------------------------------------------
    'Write the final run summary
        REG_DP_WriteSummary

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors so every cleanup step is attempted
        On Error Resume Next

    'Reset DatePicker UI artifacts after testing
        REG_DP_ResetDatePickerArtifacts

    'Delete the scratch worksheet
        REG_DP_DeleteScratchSheet

    'Restore DatePicker settings and transient state
        REG_DP_RestoreSettings SettingsSnapshot

    'Restore the manager state according to the pre-run state
        REG_DP_RestoreManagerState

    'Restore the Application state
        REG_DP_RestoreApplicationState AppSnapshot

    'Clear module object references
        Set mREG_DP_ScratchSheet = Nothing
        Set mREG_DP_ResultSheet = Nothing
        Set mREG_DP_HostWorkbook = Nothing

    'Clear any suppressed cleanup error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

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

    'Record the fatal harness failure when possible
        On Error Resume Next
        REG_DP_RecordResult REG_DP_FAIL_TEXT, _
            "Harness", _
            PROC_NAME, _
            "Fatal error " & VBA.CStr(FatalNumber) & " - " & FatalDescription
        Err.Clear
        On Error GoTo 0

    'Run shared cleanup
        Resume CleanExit

End Sub

Private Sub REG_DP_ResetHarnessState()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' RESET COUNTERS
'------------------------------------------------------------------------------
    'Reset assertion counters
        mREG_DP_RunCount = 0

    'Reset pass counter
        mREG_DP_PassCount = 0

    'Reset fail counter
        mREG_DP_FailCount = 0

    'Reset result-row pointer
        mREG_DP_NextResultRow = 2

    'Reset current suite name
        mREG_DP_CurrentSuite = vbNullString

'------------------------------------------------------------------------------
' RESET REFERENCES
'------------------------------------------------------------------------------
    'Clear result sheet reference
        Set mREG_DP_ResultSheet = Nothing

    'Clear scratch sheet reference
        Set mREG_DP_ScratchSheet = Nothing

    'Clear host workbook reference
        Set mREG_DP_HostWorkbook = Nothing

End Sub

'
'==============================================================================
'
'                                  TEST SUITES
'
'==============================================================================

Private Sub REG_DP_RunSuiteSafe(ByVal SuiteName As String)

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
'   Best-effort. Does not intentionally raise outward
'
' DEPENDENCIES
'   REG_DP_RunSuite_* routines
'
' NOTES
'   This wrapper is deliberately Select Case based so the suite routines can
'   remain Private and do not need Application.Run
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "REG_DP_RunSuiteSafe"

    Dim ErrorNumber             As Long          'Captured suite error number
    Dim ErrorDescription        As String        'Captured suite error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Normalize the current suite name for diagnostics
        mREG_DP_CurrentSuite = SuiteName

    'Protect the harness from escaping suite failures
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' DISPATCH SUITE
'------------------------------------------------------------------------------
    'Run the requested suite
        Select Case VBA.UCase$(VBA.Trim$(SuiteName))
            Case "ENVIRONMENT"
                REG_DP_RunSuite_Environment

            Case "SETTINGS"
                REG_DP_RunSuite_Settings

            Case "DATEPOLICY"
                REG_DP_RunSuite_DatePolicy

            Case "HOLIDAYPOLICY"
                REG_DP_RunSuite_HolidayPolicy

            Case "CAPTIONS"
                REG_DP_RunSuite_Captions

            Case "FORMBRIDGE"
                REG_DP_RunSuite_FormBridge

            Case "WRITEBACK"
                REG_DP_RunSuite_WriteBack

            Case "GRIDICON"
                REG_DP_RunSuite_GridIcon

            Case "MANAGER"
                REG_DP_RunSuite_Manager

            Case "UISMOKE"
                REG_DP_RunSuite_UISmoke

            Case Else
                REG_DP_RecordFail "Unknown suite", "SuiteName=" & SuiteName
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
        REG_DP_RecordResult REG_DP_FAIL_TEXT, _
            SuiteName, _
            PROC_NAME, _
            "Escaped suite error " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription

    'Clear any suppressed result-recording error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub REG_DP_RunSuite_Environment()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "Environment"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' SETTINGS AND MANAGER
'------------------------------------------------------------------------------
    'Ensure settings load successfully
        M_Settings_EnsureLoaded

    'Record successful settings load
        REG_DP_RecordPass "M_Settings_EnsureLoaded does not raise", vbNullString

    'Ensure the global manager can be created
        M_Picker_EnsureManager

    'Assert that the global manager exists
        REG_DP_AssertTrue "Global manager is instantiated", Not (gDP_Manager Is Nothing)

    'Assert that the manager is not busy after startup
        REG_DP_AssertFalse "Manager is not busy after startup", gDP_Manager.Is_Busy

'------------------------------------------------------------------------------
' PLATFORM HELPERS
'------------------------------------------------------------------------------
    'Call the platform capability helper
        REG_DP_AssertBooleanResult "M_Platform_CanUseWinAPI returns Boolean", M_Platform_CanUseWinAPI

    'Call the effective WinAPI helper
        REG_DP_AssertBooleanResult "M_Platform_ShouldUseWinAPI returns Boolean", M_Platform_ShouldUseWinAPI

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
        REG_DP_RecordFail "Environment suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_Settings()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "Settings"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' FIRST-DAY VALIDATION
'------------------------------------------------------------------------------
    'Assert vbSunday is supported
        REG_DP_AssertTrue "vbSunday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbSunday)

    'Assert vbMonday is supported
        REG_DP_AssertTrue "vbMonday is a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(vbMonday)

    'Assert invalid first-day value is rejected by predicate
        REG_DP_AssertFalse "0 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(0)

    'Assert invalid first-day value is rejected by predicate
        REG_DP_AssertFalse "3 is not a valid first-day setting", _
            M_Settings_IsValidFirstDayOfWeek(3)

    'Assert Sunday text conversion
        REG_DP_AssertEqualsString "vbSunday converts to text", _
            "vbSunday", _
            M_Settings_FirstDayOfWeekToText(vbSunday)

    'Assert Monday text conversion
        REG_DP_AssertEqualsString "vbMonday converts to text", _
            "vbMonday", _
            M_Settings_FirstDayOfWeekToText(vbMonday)

    'Assert invalid first-day conversion raises
        REG_DP_ExpectError_FirstDayToTextInvalid

'------------------------------------------------------------------------------
' FIRST-DAY SETTERS
'------------------------------------------------------------------------------
    'Set Sunday as first day
        M_Settings_SetFirstDayOfWeek vbSunday

    'Assert Sunday was stored
        REG_DP_AssertEqualsLong "SetFirstDayOfWeek stores vbSunday", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Monday through text parser
        M_Settings_SetFirstDayOfWeekText "Monday"

    'Assert Monday was stored
        REG_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Monday", _
            vbMonday, _
            M_Settings_GetFirstDayOfWeek()

    'Set Sunday through short text parser
        M_Settings_SetFirstDayOfWeekText "Sun"

    'Assert Sunday was stored
        REG_DP_AssertEqualsLong "SetFirstDayOfWeekText parses Sun", _
            vbSunday, _
            M_Settings_GetFirstDayOfWeek()

    'Assert blank first-day text raises
        REG_DP_ExpectError_SetFirstDayBlank

    'Assert unsupported first-day text raises
        REG_DP_ExpectError_SetFirstDayInvalidText

'------------------------------------------------------------------------------
' BOOLEAN SETTINGS
'------------------------------------------------------------------------------
    'Set local names on
        M_Settings_SetUseLocalDayNames True

    'Assert local names setting is on
        REG_DP_AssertTrue "UseLocalNames can be enabled", gDP_UseLocalNames

    'Set local names off
        M_Settings_SetUseLocalDayNames False

    'Assert local names setting is off
        REG_DP_AssertFalse "UseLocalNames can be disabled", gDP_UseLocalNames

    'Set outside-month selection on
        M_Settings_SetAllowOutsideMonthDays True

    'Assert outside-month selection is on
        REG_DP_AssertTrue "AllowOutsideMonthDays can be enabled", gDP_AllowOutsideMonthSel

    'Set outside-month selection off
        M_Settings_SetAllowOutsideMonthDays False

    'Assert outside-month selection is off
        REG_DP_AssertFalse "AllowOutsideMonthDays can be disabled", gDP_AllowOutsideMonthSel

    'Set weekend highlighting on
        M_Settings_SetHighlightWeekends True

    'Assert weekend highlighting is on
        REG_DP_AssertTrue "HighlightWeekends can be enabled", gDP_HighlightWeekends

    'Set weekend highlighting off
        M_Settings_SetHighlightWeekends False

    'Assert weekend highlighting is off
        REG_DP_AssertFalse "HighlightWeekends can be disabled", gDP_HighlightWeekends

    'Set close-after-selection on
        M_Settings_SetCloseAfterSelection True

    'Assert close-after-selection is on
        REG_DP_AssertTrue "CloseAfterSelection can be enabled", gDP_CloseAfterSelection

    'Set close-after-selection off
        M_Settings_SetCloseAfterSelection False

    'Assert close-after-selection is off
        REG_DP_AssertFalse "CloseAfterSelection can be disabled", gDP_CloseAfterSelection

'------------------------------------------------------------------------------
' ENUM SETTINGS
'------------------------------------------------------------------------------
    'Set static clock mode
        M_Settings_SetClockMode DP_ClockMode_Static

    'Assert static clock mode is stored
        REG_DP_AssertEqualsLong "ClockMode can be static", _
            DP_ClockMode_Static, _
            gDP_ClockMode

    'Set live clock mode
        M_Settings_SetClockMode DP_ClockMode_Live

    'Assert live clock mode is stored
        REG_DP_AssertEqualsLong "ClockMode can be live", _
            DP_ClockMode_Live, _
            gDP_ClockMode

    'Assert unsupported clock mode raises
        REG_DP_ExpectError_SetInvalidClockMode

    'Set normal size mode
        M_Settings_SetSizeMode DP_SizeMode_Normal

    'Assert normal size mode is stored
        REG_DP_AssertEqualsLong "SizeMode can be normal", _
            DP_SizeMode_Normal, _
            gDP_SizeMode

    'Set compact size mode
        M_Settings_SetSizeMode DP_SizeMode_Compact

    'Assert compact size mode is stored
        REG_DP_AssertEqualsLong "SizeMode can be compact", _
            DP_SizeMode_Compact, _
            gDP_SizeMode

    'Assert unsupported size mode raises
        REG_DP_ExpectError_SetInvalidSizeMode

'------------------------------------------------------------------------------
' ACCESS FALLBACK SETTINGS
'------------------------------------------------------------------------------
    'Disable right-click access
        M_Settings_SetShowRightClick False

    'Disable grid-icon access
        M_Settings_SetShowGridIcon False

    'Assert keyboard access is forced on when both visible entry points are off
        REG_DP_AssertTrue "Keyboard shortcut is forced when right-click and grid icon are disabled", _
            gDP_EnableKeyboardShortcut

    'Enable right-click access again for subsequent suites
        M_Settings_SetShowRightClick True

    'Enable grid-icon access again for subsequent suites
        M_Settings_SetShowGridIcon True

'------------------------------------------------------------------------------
' HOLIDAY CALLBACK SETTING
'------------------------------------------------------------------------------
    'Set a trimmed holiday callback name
        M_Settings_SetHolidayCallback "  REG_DP_HolidayCallback  "

    'Assert the callback name was trimmed
        REG_DP_AssertEqualsString "Holiday callback name is trimmed", _
            "REG_DP_HolidayCallback", _
            gDP_HolidayCallbackName

    'Clear the holiday callback name
        M_Settings_SetHolidayCallback vbNullString

    'Assert the callback name was cleared
        REG_DP_AssertEqualsString "Holiday callback name can be cleared", _
            vbNullString, _
            gDP_HolidayCallbackName

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
        REG_DP_RecordFail "Settings suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_DatePolicy()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "DatePolicy"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' OUTSIDE-MONTH ENABLED
'------------------------------------------------------------------------------
    'Enable outside-month selection for the first policy branch
        gDP_AllowOutsideMonthSel = True

    'Assert current-month date can be selected
        REG_DP_AssertTrue "Current-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert outside-month date can be selected
        REG_DP_AssertTrue "Outside-month date is selectable when outside-month is enabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' OUTSIDE-MONTH DISABLED
'------------------------------------------------------------------------------
    'Disable outside-month selection for the second policy branch
        gDP_AllowOutsideMonthSel = False

    'Assert current-month date can still be selected
        REG_DP_AssertTrue "Current-month date is selectable when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 5)

    'Assert outside-month date is rejected
        REG_DP_AssertFalse "Outside-month date is rejected when outside-month is disabled", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 4, 30), 2026, 5)

'------------------------------------------------------------------------------
' INVALID DISPLAY PERIODS
'------------------------------------------------------------------------------
    'Assert invalid month is rejected
        REG_DP_AssertFalse "Display month 0 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 0)

    'Assert invalid month is rejected
        REG_DP_AssertFalse "Display month 13 is rejected", _
            M_DatePolicy_CanSelectDate(VBA.DateSerial(2026, 5, 3), 2026, 13)

    'Assert invalid year is rejected
        REG_DP_AssertFalse "Display year 99 is rejected", _
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
        REG_DP_RecordFail "Date policy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_HolidayPolicy()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "HolidayPolicy"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' BLANK CALLBACK
'------------------------------------------------------------------------------
    'Clear callback name
        gDP_HolidayCallbackName = vbNullString

    'Assert blank callback returns False
        REG_DP_AssertFalse "Blank holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' BOOLEAN CALLBACK
'------------------------------------------------------------------------------
    'Set deterministic Boolean callback
        gDP_HolidayCallbackName = REG_DP_QualifiedMacroName("REG_DP_HolidayCallback")

    'Assert matching holiday returns True
        REG_DP_AssertTrue "Boolean callback can return True", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

    'Assert non-matching holiday returns False
        REG_DP_AssertFalse "Boolean callback can return False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 2))

'------------------------------------------------------------------------------
' NON-BOOLEAN CALLBACK
'------------------------------------------------------------------------------
    'Set non-Boolean callback
        gDP_HolidayCallbackName = REG_DP_QualifiedMacroName("REG_DP_HolidayCallbackNonBoolean")

    'Assert non-Boolean callback result is ignored
        REG_DP_AssertFalse "Non-Boolean holiday callback is ignored", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

'------------------------------------------------------------------------------
' MISSING AND ERROR-VALUE CALLBACKS
'------------------------------------------------------------------------------
    'Set missing callback name
        gDP_HolidayCallbackName = "REG_DP_MissingHolidayCallback"

    'Assert missing callback is fail-safe
        REG_DP_AssertFalse "Missing holiday callback returns False", _
            M_HolidayPolicy_IsHolidayDate(VBA.DateSerial(2026, 1, 1))

    'Set error-value callback name
        gDP_HolidayCallbackName = REG_DP_QualifiedMacroName("REG_DP_HolidayCallbackError")

    'Assert Excel error callback result is ignored
        REG_DP_AssertFalse "Excel error holiday callback result returns False", _
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
        REG_DP_RecordFail "Holiday policy suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_Captions()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "Captions"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' ENGLISH MONTH CAPTIONS
'------------------------------------------------------------------------------
    'Assert short English January caption
        REG_DP_AssertEqualsString "English short month 1 is JAN", _
            "JAN", _
            M_Caption_GetEnglishMonthShort(1)

    'Assert short English December caption
        REG_DP_AssertEqualsString "English short month 12 is DEC", _
            "DEC", _
            M_Caption_GetEnglishMonthShort(12)

    'Assert full English January caption
        REG_DP_AssertEqualsString "English full month 1 is JANUARY", _
            "JANUARY", _
            M_Caption_GetEnglishMonthFull(1)

    'Assert full English December caption
        REG_DP_AssertEqualsString "English full month 12 is DECEMBER", _
            "DECEMBER", _
            M_Caption_GetEnglishMonthFull(12)

    'Assert fixed-English month helper output
        REG_DP_AssertEqualsString "GetMonth fixed-English returns uppercase full month", _
            "MAY", _
            M_Caption_GetMonth(5, False)

'------------------------------------------------------------------------------
' DATE CAPTIONS
'------------------------------------------------------------------------------
    'Assert fixed-English date caption
        REG_DP_AssertEqualsString "GetDate fixed-English returns dd-MMM-yyyy", _
            "03-MAY-2026", _
            M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), False)

    'Assert local month caption is non-empty
        REG_DP_AssertTrue "GetMonth local caption is non-empty", _
            VBA.Len(M_Caption_GetMonth(5, True)) > 0

    'Assert local date caption is non-empty
        REG_DP_AssertTrue "GetDate local caption is non-empty", _
            VBA.Len(M_Caption_GetDate(VBA.DateSerial(2026, 5, 3), True)) > 0

'------------------------------------------------------------------------------
' INVALID INPUTS
'------------------------------------------------------------------------------
    'Assert invalid short month raises
        REG_DP_ExpectError_EnglishMonthShortInvalid

    'Assert invalid full month raises
        REG_DP_ExpectError_EnglishMonthFullInvalid

    'Assert invalid GetMonth raises
        REG_DP_ExpectError_GetMonthInvalid

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
        REG_DP_RecordFail "Caption suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_FormBridge()

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
'   Tests initial-date consumption and explicit-cell no-form cleanup / refresh calls
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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim InitialDate             As Date     'Consumed initial date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "FormBridge"

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
        REG_DP_AssertTrue "Initial date is consumed when available", _
            M_FormBridge_ConsumeInitialDate(InitialDate)

    'Assert consumed initial date is correct
        REG_DP_AssertDateEquals "Consumed initial date matches bridge value", _
            VBA.DateSerial(2026, 5, 3), _
            InitialDate

    'Assert initial date flag was cleared after consumption
        REG_DP_AssertFalse "Initial date flag is cleared after consumption", _
            gDP_HasInitialDate

    'Assert second consumption returns False
        REG_DP_AssertFalse "Initial date cannot be consumed twice", _
            M_FormBridge_ConsumeInitialDate(InitialDate)

'------------------------------------------------------------------------------
' NO-FORM BRIDGE CALLS
'------------------------------------------------------------------------------
    'Close picker when no form may be loaded
        DP_Close

    'Record successful close with no visible form
        REG_DP_RecordPass "DP_Close is safe with no visible form", vbNullString

    'Refresh from an explicit scratch cell when no form may be loaded
        M_FormBridge_RefreshFromCell mREG_DP_ScratchSheet.Range("A1")

    'Record successful no-form refresh
        REG_DP_RecordPass "M_FormBridge_RefreshFromCell is safe with explicit cell and no visible form", vbNullString

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
        REG_DP_RecordFail "Form bridge suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_WriteBack()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetRange             As Range        'Target range under test
    Dim TableRange              As Range        'Table range under test
    Dim TestTable               As ListObject   'Regression test table
    Dim UnionRange              As Range        'Discontiguous target range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "WriteBack"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Ensure the scratch sheet is active for selection-based tests
        mREG_DP_ScratchSheet.Activate

'------------------------------------------------------------------------------
' DIRECT RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the target range
        mREG_DP_ScratchSheet.Range("B2:B4").ClearContents

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 5, 3)

    'Populate a contiguous range directly
        M_WriteBack_PopulateRange mREG_DP_ScratchSheet.Range("B2:B4"), Date_Picker

    'Assert the first cell was written
        REG_DP_AssertCellDateEquals "Direct range write B2", _
            VBA.DateSerial(2026, 5, 3), _
            mREG_DP_ScratchSheet.Range("B2")

    'Assert the last cell was written
        REG_DP_AssertCellDateEquals "Direct range write B4", _
            VBA.DateSerial(2026, 5, 3), _
            mREG_DP_ScratchSheet.Range("B4")

'------------------------------------------------------------------------------
' DISCONTIGUOUS RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the discontiguous test cells
        mREG_DP_ScratchSheet.Range("C2:C4").ClearContents

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 6, 15)

    'Build a discontiguous target range
        Set UnionRange = Application.Union( _
            mREG_DP_ScratchSheet.Range("C2"), _
            mREG_DP_ScratchSheet.Range("C4"))

    'Populate the discontiguous range directly
        M_WriteBack_PopulateRange UnionRange, Date_Picker

    'Assert the first discontiguous cell was written
        REG_DP_AssertCellDateEquals "Discontiguous write C2", _
            VBA.DateSerial(2026, 6, 15), _
            mREG_DP_ScratchSheet.Range("C2")

    'Assert the second discontiguous cell was written
        REG_DP_AssertCellDateEquals "Discontiguous write C4", _
            VBA.DateSerial(2026, 6, 15), _
            mREG_DP_ScratchSheet.Range("C4")

    'Assert the skipped middle cell remains blank
        REG_DP_AssertTrue "Discontiguous write leaves C3 blank", _
            VBA.LenB(VBA.CStr(mREG_DP_ScratchSheet.Range("C3").Value)) = 0

'------------------------------------------------------------------------------
' SELECTION-BASED WRITE-BACK
'------------------------------------------------------------------------------
    'Clear the selection-based target range
        mREG_DP_ScratchSheet.Range("D2:D3").ClearContents

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 7, 20)

    'Select the target range
        mREG_DP_ScratchSheet.Range("D2:D3").Select

    'Apply DatePicker write-back to the current selection without table growth
        M_WriteBack_Apply Date_Picker, True

    'Assert the first selected cell was written
        REG_DP_AssertCellDateEquals "Selection write D2", _
            VBA.DateSerial(2026, 7, 20), _
            mREG_DP_ScratchSheet.Range("D2")

    'Assert the second selected cell was written
        REG_DP_AssertCellDateEquals "Selection write D3", _
            VBA.DateSerial(2026, 7, 20), _
            mREG_DP_ScratchSheet.Range("D3")

'------------------------------------------------------------------------------
' TABLE COLUMN EXPANSION
'------------------------------------------------------------------------------
    'Prepare the table source range
        Set TableRange = mREG_DP_ScratchSheet.Range("F1:G4")

    'Clear the table source range
        TableRange.Clear

    'Write table headers
        mREG_DP_ScratchSheet.Range("F1").Value = "ID"
        mREG_DP_ScratchSheet.Range("G1").Value = "DateValue"

    'Write table IDs
        mREG_DP_ScratchSheet.Range("F2:F4").Value = 1

    'Create the regression table
        Set TestTable = mREG_DP_ScratchSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=TableRange, _
            XlListObjectHasHeaders:=xlYes)

    'Name the regression table
        TestTable.Name = "REG_DP_Table"

    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 8, 25)

    'Select one table data cell in the date column
        mREG_DP_ScratchSheet.Range("G2").Select

    'Apply DatePicker write-back with table-column expansion enabled
        M_WriteBack_Apply Date_Picker, False

    'Assert the first table data cell was written
        REG_DP_AssertCellDateEquals "Table-column expansion writes G2", _
            VBA.DateSerial(2026, 8, 25), _
            mREG_DP_ScratchSheet.Range("G2")

    'Assert the last table data cell was written
        REG_DP_AssertCellDateEquals "Table-column expansion writes G4", _
            VBA.DateSerial(2026, 8, 25), _
            mREG_DP_ScratchSheet.Range("G4")

'------------------------------------------------------------------------------
' INVALID WRITE ACTION
'------------------------------------------------------------------------------
    'Assert unsupported write action raises
        REG_DP_ExpectError_InvalidWriteAction

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
        REG_DP_RecordFail "Write-back suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_GridIcon()

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
'   2026-05-03
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
        mREG_DP_CurrentSuite = "GridIcon"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Ensure the scratch sheet is active
        mREG_DP_ScratchSheet.Activate

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
        REG_DP_AssertTrue "Embedded grid icon path is not blank", _
            VBA.LenB(IconPath) > 0

    'Assert the embedded icon file exists
        REG_DP_AssertTrue "Embedded grid icon file exists", _
            VBA.LenB(VBA.Dir$(IconPath, vbNormal)) > 0

'------------------------------------------------------------------------------
' CREATE ICON
'------------------------------------------------------------------------------
    'Create or move the grid icon beside B2
        M_GridIcon_ShowOrMove mREG_DP_ScratchSheet.Range("B2")

    'Assert the grid icon exists on the scratch sheet
        REG_DP_AssertTrue "Grid icon is created on eligible target", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Assert only one named grid icon exists in the host workbook
        REG_DP_AssertEqualsLong "Only one grid icon exists after create", _
            1, _
            REG_DP_CountNamedShapes(mREG_DP_HostWorkbook, DP_GRID_ICON_NAME)

    'Capture the initial icon left position
        ShapeLeftBefore = mREG_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Left

    'Capture the initial icon top position
        ShapeTopBefore = mREG_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Top

'------------------------------------------------------------------------------
' MOVE ICON
'------------------------------------------------------------------------------
    'Move the grid icon beside D5
        M_GridIcon_ShowOrMove mREG_DP_ScratchSheet.Range("D5")

    'Assert the grid icon still exists after move
        REG_DP_AssertTrue "Grid icon still exists after move", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Assert only one named grid icon exists after move
        REG_DP_AssertEqualsLong "Only one grid icon exists after move", _
            1, _
            REG_DP_CountNamedShapes(mREG_DP_HostWorkbook, DP_GRID_ICON_NAME)

    'Assert the icon moved horizontally or vertically
        REG_DP_AssertTrue "Grid icon position changes after move", _
            (mREG_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Left <> ShapeLeftBefore) Or _
            (mREG_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Top <> ShapeTopBefore)

'------------------------------------------------------------------------------
' REMOVE AND PURGE ICON
'------------------------------------------------------------------------------
    'Remove the current grid icon
        M_GridIcon_Remove

    'Assert the grid icon no longer exists on the active scratch sheet
        REG_DP_AssertFalse "Grid icon is removed from active scratch sheet", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Create the icon again for purge testing
        M_GridIcon_ShowOrMove mREG_DP_ScratchSheet.Range("B2")

    'Purge all named grid icons
        M_GridIcon_PurgeAll

    'Assert no named grid icon remains in the host workbook
        REG_DP_AssertEqualsLong "Purge removes all grid icons", _
            0, _
            REG_DP_CountNamedShapes(mREG_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' DISABLED FEATURE BEHAVIOR
'------------------------------------------------------------------------------
    'Disable grid icon feature directly
        gDP_ShowGridIcon = False

    'Attempt to show icon while disabled
        M_GridIcon_ShowOrMove mREG_DP_ScratchSheet.Range("B2")

    'Assert disabled feature does not leave an icon
        REG_DP_AssertFalse "Grid icon is not shown when feature is disabled", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

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
        REG_DP_RecordFail "Grid icon suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

End Sub

Private Sub REG_DP_RunSuite_Manager()

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
'   Tests manager creation, picker state predicates, Should_ShowDP gating,
'   explicit selection-change handling, reset behavior, and known merged-cell
'   policy expectations
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   cDatePickerManager
'   M_GridIcon_PurgeAll
'
' NOTES
'   The merged-cell assertion encodes the documented desired behavior that a
'   single cell inside a merged area should normalize to the top-left cell
'
'   The current uploaded manager rejects that case, so this test will fail until
'   Normalize_TargetCell is corrected
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Manager                 As cDatePickerManager   'Manager instance under test
    Dim MergedArea              As Range                'Merged area under test

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "Manager"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Create a fresh manager instance for public API checks
        Set Manager = New cDatePickerManager

    'Ensure the scratch sheet is active
        mREG_DP_ScratchSheet.Activate

    'Remove stale grid icons before manager tests
        M_GridIcon_PurgeAll

    'Enable grid icon gating directly
        gDP_ShowGridIcon = True

'------------------------------------------------------------------------------
' STATE PREDICATES
'------------------------------------------------------------------------------
    'Assert new manager is not busy
        REG_DP_AssertFalse "New manager is not busy", Manager.Is_Busy

    'Close picker before loaded / visible checks
        DP_Close

    'Assert picker is not visible after close
        REG_DP_AssertFalse "PickerVisible is False after DP_Close", Manager.PickerVisible

    'Assert picker is not loaded after close
        REG_DP_AssertFalse "Is_PickerLoaded is False after DP_Close", Manager.Is_PickerLoaded

'------------------------------------------------------------------------------
' SHOULD_SHOWDP GATING
'------------------------------------------------------------------------------
    'Prepare a date value cell
        mREG_DP_ScratchSheet.Range("B10").Value = VBA.DateSerial(2026, 5, 3)

    'Assert date value cell is eligible
        REG_DP_AssertTrue "Should_ShowDP accepts explicit date value cell", _
            Manager.Should_ShowDP(mREG_DP_ScratchSheet.Range("B10"))

    'Prepare a date-formatted blank cell
        mREG_DP_ScratchSheet.Range("B11").ClearContents
        mREG_DP_ScratchSheet.Range("B11").NumberFormat = "dd/mm/yyyy"

    'Assert date-formatted blank cell is eligible
        REG_DP_AssertTrue "Should_ShowDP accepts date-formatted blank cell", _
            Manager.Should_ShowDP(mREG_DP_ScratchSheet.Range("B11"))

    'Prepare a general blank cell
        mREG_DP_ScratchSheet.Range("B12").ClearContents
        mREG_DP_ScratchSheet.Range("B12").NumberFormat = "General"

    'Assert general blank cell is not eligible
        REG_DP_AssertFalse "Should_ShowDP rejects general blank cell", _
            Manager.Should_ShowDP(mREG_DP_ScratchSheet.Range("B12"))

    'Assert multi-cell range is not eligible
        REG_DP_AssertFalse "Should_ShowDP rejects multi-cell range", _
            Manager.Should_ShowDP(mREG_DP_ScratchSheet.Range("B10:B11"))

'------------------------------------------------------------------------------
' MERGED-CELL NORMALIZATION EXPECTATION
'------------------------------------------------------------------------------
    'Prepare the merged area
        Set MergedArea = mREG_DP_ScratchSheet.Range("B14:C15")

    'Clear the merged area
        MergedArea.Clear

    'Merge the test area
        MergedArea.Merge

    'Format the merged area as date-like
        MergedArea.NumberFormat = "dd/mm/yyyy"

    'Assert a single cell inside a merged area is eligible by documented policy
        REG_DP_AssertTrue "Should_ShowDP accepts top-left cell inside merged area", _
            Manager.Should_ShowDP(mREG_DP_ScratchSheet.Range("B14"))

    'Unmerge the area after the assertion
        MergedArea.UnMerge

'------------------------------------------------------------------------------
' SELECTION-CHANGE HANDLING
'------------------------------------------------------------------------------
    'Prepare a date value cell for selection-change handling
        mREG_DP_ScratchSheet.Range("E10").Value = VBA.DateSerial(2026, 9, 9)

    'Handle selection change for an eligible target
        Manager.Handle_SelectionChange mREG_DP_ScratchSheet.Range("E10")

    'Assert the manager created the grid icon for the eligible target
        REG_DP_AssertTrue "Handle_SelectionChange shows icon for eligible target", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Prepare an ineligible target
        mREG_DP_ScratchSheet.Range("E11").ClearContents
        mREG_DP_ScratchSheet.Range("E11").NumberFormat = "General"

    'Handle selection change for an ineligible target
        Manager.Handle_SelectionChange mREG_DP_ScratchSheet.Range("E11")

    'Assert the manager removed the grid icon for the ineligible target
        REG_DP_AssertFalse "Handle_SelectionChange removes icon for ineligible target", _
            REG_DP_ShapeExists(mREG_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' RESET BEHAVIOR
'------------------------------------------------------------------------------
    'Show the icon again for reset testing
        Manager.Handle_SelectionChange mREG_DP_ScratchSheet.Range("E10")

    'Reset all DatePicker UI through the manager
        Manager.Reset_DatePickerUI

    'Assert reset removed all grid icons
        REG_DP_AssertEqualsLong "Reset_DatePickerUI purges all grid icons", _
            0, _
            REG_DP_CountNamedShapes(mREG_DP_HostWorkbook, DP_GRID_ICON_NAME)

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
        If Not MergedArea Is Nothing Then MergedArea.UnMerge

    'Release object references
        Set MergedArea = Nothing
        Set Manager = Nothing

    'Record the suite-level failure
        REG_DP_RecordFail "Manager suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub REG_DP_RunSuite_UISmoke()

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
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mREG_DP_CurrentSuite = "UISmoke"

    'Enable suite-level error handling
        On Error GoTo SuiteFail

    'Ensure the scratch sheet is active
        mREG_DP_ScratchSheet.Activate

    'Prepare the active cell with a deterministic date
        mREG_DP_ScratchSheet.Range("H2").Value = VBA.DateSerial(2026, 5, 3)

    'Select the date cell
        mREG_DP_ScratchSheet.Range("H2").Select

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
        REG_DP_AssertTrue "DP_Show loads the picker", gDP_Manager.Is_PickerLoaded

    'Assert the picker is visible
        REG_DP_AssertTrue "DP_Show shows the picker", gDP_Manager.PickerVisible

'------------------------------------------------------------------------------
' CLOSE FORM
'------------------------------------------------------------------------------
    'Close the DatePicker form
        DP_Close

    'Allow modeless form events to process
        DoEvents

    'Assert the picker is no longer visible
        REG_DP_AssertFalse "DP_Close hides the picker", gDP_Manager.PickerVisible

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
        REG_DP_RecordFail "UI smoke suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description

    'Clear the suite error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'==============================================================================
'
'                            EXPECTED-ERROR TESTS
'
'==============================================================================

Private Sub REG_DP_ExpectError_FirstDayToTextInvalid()

'
'==============================================================================
'                     EXPECT ERROR FIRST DAY TO TEXT INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the converter with an invalid first-day value
        M_Settings_FirstDayOfWeekToText 0

    'Record failure when no error is raised
        REG_DP_RecordFail "Invalid first-day conversion raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Invalid first-day conversion raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_SetFirstDayBlank()

'
'==============================================================================
'                       EXPECT ERROR SET FIRST DAY BLANK
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the setter with blank text
        M_Settings_SetFirstDayOfWeekText vbNullString

    'Record failure when no error is raised
        REG_DP_RecordFail "Blank first-day text raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Blank first-day text raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_SetFirstDayInvalidText()

'
'==============================================================================
'                    EXPECT ERROR SET FIRST DAY INVALID TEXT
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the setter with unsupported text
        M_Settings_SetFirstDayOfWeekText "Wednesday"

    'Record failure when no error is raised
        REG_DP_RecordFail "Unsupported first-day text raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Unsupported first-day text raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_SetInvalidClockMode()

'
'==============================================================================
'                       EXPECT ERROR INVALID CLOCK MODE
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the setter with an unsupported clock mode
        M_Settings_SetClockMode 99

    'Record failure when no error is raised
        REG_DP_RecordFail "Unsupported clock mode raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Unsupported clock mode raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_SetInvalidSizeMode()

'
'==============================================================================
'                        EXPECT ERROR INVALID SIZE MODE
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the setter with an unsupported size mode
        M_Settings_SetSizeMode 99

    'Record failure when no error is raised
        REG_DP_RecordFail "Unsupported size mode raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Unsupported size mode raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_EnglishMonthShortInvalid()

'
'==============================================================================
'                   EXPECT ERROR ENGLISH MONTH SHORT INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the helper with an invalid month number
        M_Caption_GetEnglishMonthShort 13

    'Record failure when no error is raised
        REG_DP_RecordFail "Invalid short English month raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Invalid short English month raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_EnglishMonthFullInvalid()

'
'==============================================================================
'                    EXPECT ERROR ENGLISH MONTH FULL INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the helper with an invalid month number
        M_Caption_GetEnglishMonthFull 0

    'Record failure when no error is raised
        REG_DP_RecordFail "Invalid full English month raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Invalid full English month raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_GetMonthInvalid()

'
'==============================================================================
'                         EXPECT ERROR GET MONTH INVALID
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call the helper with an invalid month number
        M_Caption_GetMonth 99, False

    'Record failure when no error is raised
        REG_DP_RecordFail "Invalid GetMonth input raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Invalid GetMonth input raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

Private Sub REG_DP_ExpectError_InvalidWriteAction()

'
'==============================================================================
'                       EXPECT ERROR INVALID WRITE ACTION
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable expected-error handling
        On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' RUN INVALID CALL
'------------------------------------------------------------------------------
    'Call write-back with an unsupported action
        M_WriteBack_PopulateRange mREG_DP_ScratchSheet.Range("J2"), 99

    'Record failure when no error is raised
        REG_DP_RecordFail "Unsupported write action raises", "No error was raised"

    'Exit after unexpected success
        Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    'Record the expected error
        REG_DP_RecordPass "Unsupported write action raises", Err.Description

    'Clear the expected error
        Err.Clear

End Sub

'
'==============================================================================
'
'                              ASSERTION HELPERS
'
'==============================================================================

Private Sub REG_DP_AssertTrue( _
    ByVal TestName As String, _
    ByVal Condition As Boolean)

'
'==============================================================================
'                              ASSERT TRUE
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass when the condition is True
        If Condition Then
            REG_DP_RecordPass TestName, vbNullString

    'Record a fail when the condition is False
        Else
            REG_DP_RecordFail TestName, "Expected True but received False"
        End If

End Sub

Private Sub REG_DP_AssertFalse( _
    ByVal TestName As String, _
    ByVal Condition As Boolean)

'
'==============================================================================
'                              ASSERT FALSE
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass when the condition is False
        If Not Condition Then
            REG_DP_RecordPass TestName, vbNullString

    'Record a fail when the condition is True
        Else
            REG_DP_RecordFail TestName, "Expected False but received True"
        End If

End Sub

Private Sub REG_DP_AssertBooleanResult( _
    ByVal TestName As String, _
    ByVal BooleanValue As Boolean)

'
'==============================================================================
'                           ASSERT BOOLEAN RESULT
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass because the Boolean expression evaluated successfully
        REG_DP_RecordPass TestName, "Value=" & VBA.CStr(BooleanValue)

End Sub

Private Sub REG_DP_AssertEqualsLong( _
    ByVal TestName As String, _
    ByVal ExpectedValue As Long, _
    ByVal ActualValue As Long)

'
'==============================================================================
'                            ASSERT EQUALS LONG
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass when the values match
        If ExpectedValue = ActualValue Then
            REG_DP_RecordPass TestName, "Value=" & VBA.CStr(ActualValue)

    'Record a fail when the values differ
        Else
            REG_DP_RecordFail TestName, _
                "Expected " & VBA.CStr(ExpectedValue) & _
                " but received " & VBA.CStr(ActualValue)
        End If

End Sub

Private Sub REG_DP_AssertEqualsString( _
    ByVal TestName As String, _
    ByVal ExpectedValue As String, _
    ByVal ActualValue As String)

'
'==============================================================================
'                           ASSERT EQUALS STRING
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass when the strings match exactly
        If VBA.StrComp(ExpectedValue, ActualValue, vbBinaryCompare) = 0 Then
            REG_DP_RecordPass TestName, "Value=" & ActualValue

    'Record a fail when the strings differ
        Else
            REG_DP_RecordFail TestName, _
                "Expected [" & ExpectedValue & "] but received [" & ActualValue & "]"
        End If

End Sub

Private Sub REG_DP_AssertDateEquals( _
    ByVal TestName As String, _
    ByVal ExpectedDate As Date, _
    ByVal ActualDate As Date)

'
'==============================================================================
'                            ASSERT DATE EQUALS
'==============================================================================

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Record a pass when the date values match
        If VBA.CDbl(ExpectedDate) = VBA.CDbl(ActualDate) Then
            REG_DP_RecordPass TestName, "Value=" & VBA.Format$(ActualDate, "yyyy-mm-dd hh:nn:ss")

    'Record a fail when the date values differ
        Else
            REG_DP_RecordFail TestName, _
                "Expected " & VBA.Format$(ExpectedDate, "yyyy-mm-dd hh:nn:ss") & _
                " but received " & VBA.Format$(ActualDate, "yyyy-mm-dd hh:nn:ss")
        End If

End Sub

Private Sub REG_DP_AssertCellDateEquals( _
    ByVal TestName As String, _
    ByVal ExpectedDate As Date, _
    ByVal TargetCell As Range)

'
'==============================================================================
'                         ASSERT CELL DATE EQUALS
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE CELL
'------------------------------------------------------------------------------
    'Fail when no target cell is supplied
        If TargetCell Is Nothing Then
            REG_DP_RecordFail TestName, "TargetCell is Nothing"
            Exit Sub
        End If

    'Fail when the target cell does not contain a date
        If Not VBA.IsDate(TargetCell.Value) Then
            REG_DP_RecordFail TestName, "Cell value is not date-like: " & VBA.CStr(TargetCell.Value)
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Assert the date value in the cell
        REG_DP_AssertDateEquals TestName, ExpectedDate, VBA.CDate(TargetCell.Value)

End Sub

'
'==============================================================================
'
'                                RESULT HELPERS
'
'==============================================================================

Private Sub REG_DP_RecordPass( _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD PASS
'==============================================================================

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Record one passed assertion
        REG_DP_RecordResult REG_DP_PASS_TEXT, mREG_DP_CurrentSuite, TestName, Details

End Sub

Private Sub REG_DP_RecordFail( _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD FAIL
'==============================================================================

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Record one failed assertion
        REG_DP_RecordResult REG_DP_FAIL_TEXT, mREG_DP_CurrentSuite, TestName, Details

End Sub

Private Sub REG_DP_RecordInfo( _
    ByVal SuiteName As String, _
    ByVal TestName As String, _
    ByVal Details As String)

'
'==============================================================================
'                              RECORD INFO
'==============================================================================

'------------------------------------------------------------------------------
' RECORD RESULT
'------------------------------------------------------------------------------
    'Record one informational row
        REG_DP_RecordResult REG_DP_INFO_TEXT, SuiteName, TestName, Details

End Sub

Private Sub REG_DP_RecordResult( _
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
'   Best-effort. Worksheet-write failures are reported to the Immediate Window
'   and do not raise outward
'
' DEPENDENCIES
'   mREG_DP_ResultSheet
'
' NOTES
'   This routine intentionally avoids recursive failure recording
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "REG_DP_RecordResult"

    Dim ResultLine              As String       'Immediate Window result line
    Dim RecordErrorNumber       As Long         'Captured result-write error number
    Dim RecordErrorDescription  As String       'Captured result-write error description

'------------------------------------------------------------------------------
' UPDATE COUNTERS
'------------------------------------------------------------------------------
    'Increment assertion counters only for pass / fail rows
        If ResultText = REG_DP_PASS_TEXT Or ResultText = REG_DP_FAIL_TEXT Then
            mREG_DP_RunCount = mREG_DP_RunCount + 1
        End If

    'Increment pass counter
        If ResultText = REG_DP_PASS_TEXT Then mREG_DP_PassCount = mREG_DP_PassCount + 1

    'Increment fail counter
        If ResultText = REG_DP_FAIL_TEXT Then mREG_DP_FailCount = mREG_DP_FailCount + 1

'------------------------------------------------------------------------------
' WRITE IMMEDIATE WINDOW
'------------------------------------------------------------------------------
    'Build one diagnostic result line
        ResultLine = ResultText & _
            " | " & SuiteName & _
            " | " & TestName & _
            IIf(VBA.Len(Details) > 0, " | " & Details, vbNullString)

    'Write the diagnostic result line
        Debug.Print ResultLine

'------------------------------------------------------------------------------
' WRITE RESULT SHEET
'------------------------------------------------------------------------------
    'Exit if the result sheet is not available
        If mREG_DP_ResultSheet Is Nothing Then Exit Sub

    'Protect the harness from result-sheet write failures
        On Error GoTo ResultSheetFail

    'Write the sequence number
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 1).Value = mREG_DP_NextResultRow - 1

    'Write the timestamp
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 2).Value = VBA.Now

    'Write the result text
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 3).Value = ResultText

    'Write the suite name
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 4).Value = SuiteName

    'Write the test name
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 5).Value = TestName

    'Write the details
        mREG_DP_ResultSheet.Cells(mREG_DP_NextResultRow, 6).Value = Details

    'Increment the next result row
        mREG_DP_NextResultRow = mREG_DP_NextResultRow + 1

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

    'Exit after recording the result
        Exit Sub

'------------------------------------------------------------------------------
' RESULT SHEET FAIL
'------------------------------------------------------------------------------
ResultSheetFail:
    'Capture the result-write error number
        RecordErrorNumber = Err.Number

    'Capture the result-write error description
        RecordErrorDescription = Err.Description

    'Write the result-sheet failure to the Immediate Window only
        Debug.Print REG_DP_INFO_TEXT & _
            " | Harness | " & PROC_NAME & _
            " | Result-sheet write skipped after error " & VBA.CStr(RecordErrorNumber) & _
            " - " & RecordErrorDescription

    'Clear the result-sheet reference to prevent repeated write failures
        Set mREG_DP_ResultSheet = Nothing

    'Clear the suppressed result-write error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub REG_DP_WriteSummary()

'
'==============================================================================
'                             WRITE SUMMARY
'==============================================================================

'------------------------------------------------------------------------------
' RECORD SUMMARY
'------------------------------------------------------------------------------
    'Record one final summary row
        REG_DP_RecordInfo "Harness", _
            "Summary", _
            "Run=" & VBA.CStr(mREG_DP_RunCount) & _
            "; Passed=" & VBA.CStr(mREG_DP_PassCount) & _
            "; Failed=" & VBA.CStr(mREG_DP_FailCount)

'------------------------------------------------------------------------------
' FORMAT SUMMARY
'------------------------------------------------------------------------------
    'Exit if result sheet is missing
        If mREG_DP_ResultSheet Is Nothing Then Exit Sub

    'Write summary title
        mREG_DP_ResultSheet.Range("H1").Value = "SUMMARY"

    'Write run count label
        mREG_DP_ResultSheet.Range("H2").Value = "Run"

    'Write run count value
        mREG_DP_ResultSheet.Range("I2").Value = mREG_DP_RunCount

    'Write pass count label
        mREG_DP_ResultSheet.Range("H3").Value = "Passed"

    'Write pass count value
        mREG_DP_ResultSheet.Range("I3").Value = mREG_DP_PassCount

    'Write fail count label
        mREG_DP_ResultSheet.Range("H4").Value = "Failed"

    'Write fail count value
        mREG_DP_ResultSheet.Range("I4").Value = mREG_DP_FailCount

    'Apply bold summary labels
        mREG_DP_ResultSheet.Range("H1:H4").Font.Bold = True

    'Auto-fit result columns
        mREG_DP_ResultSheet.Columns("A:I").AutoFit

End Sub

'
'==============================================================================
'
'                             WORKBOOK HELPERS
'
'==============================================================================

Private Function REG_DP_GetHostWorkbook() As Workbook

'
'==============================================================================
'                             GET HOST WORKBOOK
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable safe workbook resolution
        On Error Resume Next

'------------------------------------------------------------------------------
' RESOLVE ACTIVE WORKBOOK
'------------------------------------------------------------------------------
    'Use the active workbook when available
        If Not Application.ActiveWorkbook Is Nothing Then
            Set REG_DP_GetHostWorkbook = Application.ActiveWorkbook
        End If

'------------------------------------------------------------------------------
' FALL BACK TO THISWORKBOOK
'------------------------------------------------------------------------------
    'Fall back to ThisWorkbook when ActiveWorkbook is unavailable
        If REG_DP_GetHostWorkbook Is Nothing Then
            Set REG_DP_GetHostWorkbook = ThisWorkbook
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed workbook-resolution error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Function

Private Sub REG_DP_PrepareResultSheet(ByVal HostWorkbook As Workbook)

'
'==============================================================================
'                           PREPARE RESULT SHEET
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Delete any previous result sheet
        REG_DP_DeleteWorksheetIfExists HostWorkbook, REG_DP_RESULT_SHEET_NAME

    'Add a fresh result sheet
        Set mREG_DP_ResultSheet = HostWorkbook.Worksheets.Add( _
            After:=HostWorkbook.Worksheets(HostWorkbook.Worksheets.Count))

    'Name the result sheet
        mREG_DP_ResultSheet.Name = REG_DP_RESULT_SHEET_NAME

'------------------------------------------------------------------------------
' WRITE HEADERS
'------------------------------------------------------------------------------
    'Write result headers
        mREG_DP_ResultSheet.Range("A1:F1").Value = Array( _
            "#", _
            "Timestamp", _
            "Result", _
            "Suite", _
            "Test", _
            "Details")

    'Format result headers
        mREG_DP_ResultSheet.Range("A1:F1").Font.Bold = True

    'Freeze the result header row
        mREG_DP_ResultSheet.Activate
        mREG_DP_ResultSheet.Range("A2").Select
        ActiveWindow.FreezePanes = True

    'Reset the next result row
        mREG_DP_NextResultRow = 2

End Sub

Private Sub REG_DP_PrepareScratchSheet(ByVal HostWorkbook As Workbook)

'
'==============================================================================
'                           PREPARE SCRATCH SHEET
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Delete any previous scratch sheet
        REG_DP_DeleteWorksheetIfExists HostWorkbook, REG_DP_SCRATCH_SHEET_NAME

    'Add a fresh scratch sheet
        Set mREG_DP_ScratchSheet = HostWorkbook.Worksheets.Add( _
            After:=HostWorkbook.Worksheets(HostWorkbook.Worksheets.Count))

    'Name the scratch sheet
        mREG_DP_ScratchSheet.Name = REG_DP_SCRATCH_SHEET_NAME

'------------------------------------------------------------------------------
' BUILD SCRATCH LAYOUT
'------------------------------------------------------------------------------
    'Write scratch title
        mREG_DP_ScratchSheet.Range("A1").Value = "DatePicker regression scratch sheet"

    'Set useful column widths
        mREG_DP_ScratchSheet.Columns("A:J").ColumnWidth = 16

    'Activate the scratch sheet
        mREG_DP_ScratchSheet.Activate

End Sub

Private Sub REG_DP_DeleteScratchSheet()

'
'==============================================================================
'                            DELETE SCRATCH SHEET
'==============================================================================

'------------------------------------------------------------------------------
' DELETE SHEET
'------------------------------------------------------------------------------
    'Delete the scratch sheet when available
        If Not mREG_DP_HostWorkbook Is Nothing Then
            REG_DP_DeleteWorksheetIfExists mREG_DP_HostWorkbook, REG_DP_SCRATCH_SHEET_NAME
        End If

    'Clear scratch sheet reference
        Set mREG_DP_ScratchSheet = Nothing

End Sub

Private Sub REG_DP_DeleteWorksheetIfExists( _
    ByVal HostWorkbook As Workbook, _
    ByVal SheetName As String)

'
'==============================================================================
'                        DELETE WORKSHEET IF EXISTS
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetSheet             As Worksheet     'Worksheet to delete

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Exit if no workbook was supplied
        If HostWorkbook Is Nothing Then Exit Sub

    'Suppress missing-sheet and delete errors
        On Error Resume Next

'------------------------------------------------------------------------------
' DELETE SHEET
'------------------------------------------------------------------------------
    'Resolve the target sheet
        Set TargetSheet = HostWorkbook.Worksheets(SheetName)

    'Delete the target sheet when found and safe
        If Not TargetSheet Is Nothing Then
            If HostWorkbook.Worksheets.Count > 1 Then TargetSheet.Delete
        End If

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the target sheet reference
        Set TargetSheet = Nothing

    'Clear any suppressed worksheet error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'==============================================================================
'
'                              STATE HELPERS
'
'==============================================================================

Private Sub REG_DP_CaptureSettings(ByRef Snapshot As TRegDPSettingsSnapshot)

'
'==============================================================================
'                            CAPTURE SETTINGS
'==============================================================================

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before snapshotting
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE SETTINGS
'------------------------------------------------------------------------------
    'Capture the right-click feature setting
        Snapshot.ShowRightClick = gDP_ShowRightClick

    'Capture the grid-icon feature setting
        Snapshot.ShowGridIcon = gDP_ShowGridIcon

    'Capture the first-day setting
        Snapshot.FirstDayOfWeek = gDP_FirstDayOfWeek

    'Capture the local-name setting
        Snapshot.UseLocalNames = gDP_UseLocalNames

    'Capture the clock-mode setting
        Snapshot.ClockMode = gDP_ClockMode

    'Capture the size-mode setting
        Snapshot.SizeMode = gDP_SizeMode

    'Capture the weekend-highlight setting
        Snapshot.HighlightWeekends = gDP_HighlightWeekends

    'Capture the outside-month selection setting
        Snapshot.AllowOutsideMonthSelection = gDP_AllowOutsideMonthSel

    'Capture the close-after-selection setting
        Snapshot.CloseAfterSelection = gDP_CloseAfterSelection

    'Capture the WinAPI setting
        Snapshot.UseWinAPI = gDP_UseWinAPI

    'Capture the keyboard shortcut setting
        Snapshot.EnableKeyboardShortcut = gDP_EnableKeyboardShortcut

    'Capture the holiday callback name
        Snapshot.HolidayCallbackName = gDP_HolidayCallbackName

    'Capture the grid icon path
        Snapshot.IconPath = gDP_IconPath

    'Capture the transient write value
        Snapshot.WriteValue = gDP_WriteValue

    'Capture the transient initial date
        Snapshot.InitialDate = gDP_InitialDate

    'Capture the transient initial-date flag
        Snapshot.HasInitialDate = gDP_HasInitialDate

    'Capture the transient selected date
        Snapshot.SelectedDate = gDP_SelectedDate

    'Capture the transient selected-date flag
        Snapshot.HasSelectedDate = gDP_HasSelectedDate

End Sub

Private Sub REG_DP_RestoreSettings(ByRef Snapshot As TRegDPSettingsSnapshot)

'
'==============================================================================
'                            RESTORE SETTINGS
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE SETTINGS
'------------------------------------------------------------------------------
    'Restore the right-click feature setting
        gDP_ShowRightClick = Snapshot.ShowRightClick

    'Restore the grid-icon feature setting
        gDP_ShowGridIcon = Snapshot.ShowGridIcon

    'Restore the first-day setting
        gDP_FirstDayOfWeek = Snapshot.FirstDayOfWeek

    'Restore the local-name setting
        gDP_UseLocalNames = Snapshot.UseLocalNames

    'Restore the clock-mode setting
        gDP_ClockMode = Snapshot.ClockMode

    'Restore the size-mode setting
        gDP_SizeMode = Snapshot.SizeMode

    'Restore the weekend-highlight setting
        gDP_HighlightWeekends = Snapshot.HighlightWeekends

    'Restore the outside-month selection setting
        gDP_AllowOutsideMonthSel = Snapshot.AllowOutsideMonthSelection

    'Restore the close-after-selection setting
        gDP_CloseAfterSelection = Snapshot.CloseAfterSelection

    'Restore the WinAPI setting
        gDP_UseWinAPI = Snapshot.UseWinAPI

    'Restore the keyboard shortcut setting
        gDP_EnableKeyboardShortcut = Snapshot.EnableKeyboardShortcut

    'Restore the holiday callback name
        gDP_HolidayCallbackName = Snapshot.HolidayCallbackName

    'Restore the grid icon path
        gDP_IconPath = Snapshot.IconPath

    'Restore the transient write value
        gDP_WriteValue = Snapshot.WriteValue

    'Restore the transient initial date
        gDP_InitialDate = Snapshot.InitialDate

    'Restore the transient initial-date flag
        gDP_HasInitialDate = Snapshot.HasInitialDate

    'Restore the transient selected date
        gDP_SelectedDate = Snapshot.SelectedDate

    'Restore the transient selected-date flag
        gDP_HasSelectedDate = Snapshot.HasSelectedDate

'------------------------------------------------------------------------------
' PERSIST AND SYNCHRONIZE
'------------------------------------------------------------------------------
    'Persist restored settings
        M_Settings_Save

    'Synchronize context menu according to restored settings
        M_ContextMenu_Update

    'Synchronize keyboard shortcut according to restored settings
        M_KeyboardShortcut_Update

    'Remove stale grid icon when the restored setting disables it
        If Not gDP_ShowGridIcon Then M_GridIcon_Remove

End Sub

Private Sub REG_DP_CaptureApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                         CAPTURE APPLICATION STATE
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE CORE STATE
'------------------------------------------------------------------------------
    'Capture ScreenUpdating
        Snapshot.ScreenUpdating = Application.ScreenUpdating

    'Capture EnableEvents
        Snapshot.EnableEvents = Application.EnableEvents

    'Capture DisplayAlerts
        Snapshot.DisplayAlerts = Application.DisplayAlerts

    'Capture calculation mode
        Snapshot.CalculationMode = Application.Calculation

'------------------------------------------------------------------------------
' CAPTURE STATUS BAR
'------------------------------------------------------------------------------
    'Capture whether Excel owns the status bar
        Snapshot.StatusBarWasFalse = (VBA.VarType(Application.StatusBar) = vbBoolean)

    'Capture custom status-bar text when present
        If Not Snapshot.StatusBarWasFalse Then Snapshot.StatusBarText = VBA.CStr(Application.StatusBar)

End Sub

Private Sub REG_DP_PrepareApplicationForRun()

'
'==============================================================================
'                       PREPARE APPLICATION FOR RUN
'==============================================================================

'------------------------------------------------------------------------------
' SET APPLICATION STATE
'------------------------------------------------------------------------------
    'Disable screen updating during the regression run
        Application.ScreenUpdating = False

    'Disable Excel events during the regression run
        Application.EnableEvents = False

    'Disable display alerts for scratch sheet cleanup
        Application.DisplayAlerts = False

    'Set status-bar text for visibility during long runs
        Application.StatusBar = "Running DatePicker regression tests..."

End Sub

Private Sub REG_DP_RestoreApplicationState(ByRef Snapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                       RESTORE APPLICATION STATE
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE CORE STATE
'------------------------------------------------------------------------------
    'Restore calculation mode
        Application.Calculation = Snapshot.CalculationMode

    'Restore display alerts
        Application.DisplayAlerts = Snapshot.DisplayAlerts

    'Restore events
        Application.EnableEvents = Snapshot.EnableEvents

    'Restore screen updating
        Application.ScreenUpdating = Snapshot.ScreenUpdating

'------------------------------------------------------------------------------
' RESTORE STATUS BAR
'------------------------------------------------------------------------------
    'Return the status bar to Excel when Excel owned it before the run
        If Snapshot.StatusBarWasFalse Then
            Application.StatusBar = False

    'Otherwise restore the previous custom status-bar text
        Else
            Application.StatusBar = Snapshot.StatusBarText
        End If

End Sub

Private Sub REG_DP_ResetDatePickerArtifacts()

'
'==============================================================================
'                        RESET DATEPICKER ARTIFACTS
'==============================================================================

'------------------------------------------------------------------------------
' RESET UI
'------------------------------------------------------------------------------
    'Suppress reset failures so cleanup can continue
        On Error Resume Next

    'Stop the DatePicker timer
        M_Timer_Stop

    'Close the DatePicker form
        DP_Close

    'Remove context-menu item
        M_ContextMenu_Remove

    'Remove keyboard shortcut assignment
        M_KeyboardShortcut_Remove

    'Purge all grid icons
        M_GridIcon_PurgeAll

    'Clear any suppressed reset error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub REG_DP_RestoreManagerState()

'
'==============================================================================
'                         RESTORE MANAGER STATE
'==============================================================================

'------------------------------------------------------------------------------
' RESET MANAGER
'------------------------------------------------------------------------------
    'Release the manager created or used during tests
        Set gDP_Manager = Nothing

'------------------------------------------------------------------------------
' RESTORE PRE-RUN POLICY
'------------------------------------------------------------------------------
    'Recreate the manager when one existed before the run
        If mREG_DP_HadManager Then M_Picker_EnsureManager

End Sub

'
'==============================================================================
'
'                              OBJECT HELPERS
'
'==============================================================================

Private Function REG_DP_ShapeExists( _
    ByVal TargetSheet As Worksheet, _
    ByVal ShapeName As String) As Boolean

'
'==============================================================================
'                              SHAPE EXISTS
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetShape             As Shape        'Resolved shape

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not found
        REG_DP_ShapeExists = False

    'Suppress missing-shape errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Resolve the shape by name
        Set TargetShape = TargetSheet.Shapes(ShapeName)

    'Return whether the shape was found
        REG_DP_ShapeExists = Not (TargetShape Is Nothing)

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release the shape reference
        Set TargetShape = Nothing

    'Clear any suppressed shape error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Function

Private Function REG_DP_CountNamedShapes( _
    ByVal HostWorkbook As Workbook, _
    ByVal ShapeName As String) As Long

'
'==============================================================================
'                            COUNT NAMED SHAPES
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetSheet             As Worksheet     'Worksheet being inspected
    Dim TargetShape             As Shape         'Shape being inspected
    Dim ShapeCount              As Long          'Matched shape count

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default the count to zero
        ShapeCount = 0

    'Exit if no workbook was supplied
        If HostWorkbook Is Nothing Then
            REG_DP_CountNamedShapes = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' COUNT SHAPES
'------------------------------------------------------------------------------
    'Loop through worksheets in the host workbook
        For Each TargetSheet In HostWorkbook.Worksheets
            'Loop through shapes in the worksheet
                For Each TargetShape In TargetSheet.Shapes
                    'Count matching shape names
                        If VBA.StrComp(TargetShape.Name, ShapeName, vbBinaryCompare) = 0 Then
                            ShapeCount = ShapeCount + 1
                        End If
                Next TargetShape
        Next TargetSheet

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the matched shape count
        REG_DP_CountNamedShapes = ShapeCount

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release object references
        Set TargetShape = Nothing
        Set TargetSheet = Nothing

End Function

Private Function REG_DP_QualifiedMacroName(ByVal MacroName As String) As String

'
'==============================================================================
'                           QUALIFIED MACRO NAME
'==============================================================================

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return a workbook-qualified macro name for Application.Run
        REG_DP_QualifiedMacroName = "'" & ThisWorkbook.Name & "'!" & MacroName

End Function


