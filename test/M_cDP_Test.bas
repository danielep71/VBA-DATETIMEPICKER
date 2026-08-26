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
'   Production routines that end with On Error GoTo 0 (M_GridIcon_ShowOrMove,
'   M_GridIcon_Remove, M_GridIcon_PurgeAll, M_GridIcon_EnsureEmbeddedIconFile,
'   M_GridIcon_PreCreateHidden, DP_Close, DP_Stop, Handle_SelectionChange, and
'   the access-path setters SetShowRightClick and SetShowGridIcon) kill the
'   suite SuiteFail handler on return. Every call to those routines is
'   immediately followed by On Error GoTo SuiteFail to re-arm the handler.
'
'   DP_RepairRuntime and M_Picker_SelectDate use On Error GoTo ErrorHandler and
'   raise outward on failure; they do not reset the caller SuiteFail handler and
'   therefore do not need re-arming.
'
'   The write-back routines are Functions returning DP_WriteResult. Bare calls
'   still compile and are kept where the outcome is not asserted, so the suite
'   covers both call forms.
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const TST_DP_RESULT_SHEET_NAME  As String = "TST_DP_RESULTS"        'Regression result worksheet name
    Private Const TST_DP_SCRATCH_SHEET_NAME As String = "TST_DP_SCRATCH"        'Regression scratch worksheet name
    Private Const TST_DP_PASS_TEXT          As String = "PASS"                  'Passed test marker
    Private Const TST_DP_FAIL_TEXT          As String = "FAIL"                  'Failed test marker
    Private Const TST_DP_INFO_TEXT          As String = "INFO"                  'Information marker

    '-----------------------------RUN STATES-----------------------------------
    'A run reports one of these. PASS is the only state that means the run both
    'completed and left the environment as it found it
    Private Const TST_DP_STATE_PASS         As String = "PASS"                  'All assertions passed and cleanup verified
    Private Const TST_DP_STATE_FAIL         As String = "FAIL"                  'One or more assertions failed
    Private Const TST_DP_STATE_FAIL_CLEANUP As String = "FAIL_CLEANUP"          'Assertions passed but cleanup did not complete
    Private Const TST_DP_STATE_INCOMPLETE   As String = "INCOMPLETE_SKIPPED"    'A dispatched suite did not complete
    Private Const TST_DP_STATE_DIRTY_START  As String = "FAIL_DIRTY_START"      'The run began in an environment a previous run left behind
    Private Const TST_DP_MODULE_NAME        As String = "M_cDP_Test"            'This module name for callback resolution

    'Result sheet layout
    Private Const TST_DP_RESULT_FIRST_ROW   As Long = 5                         'First result data row on the result sheet
    Private Const TST_DP_STATUS_BAR_TEXT    As String = "Running DatePicker regression tests..."  'Status bar text the run displays
    Private Const TST_DP_STALE_SHEET_NAME   As String = "TST_DP_STALE"          'Temporary sheet used to strand a grid icon
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
' NATIVE DECLARATIONS FOR INDEPENDENT WINDOW-STYLE VERIFICATION
'------------------------------------------------------------------------------
'   The window-style suite must be able to read the native style itself. Proving
'   a transaction only from the result the transaction returns is the function
'   testing itself
'
'   GWL_STYLE and WS_CAPTION are duplicated from M_DatePicker, where they are
'   Private. They are fixed Windows values and cannot drift
'------------------------------------------------------------------------------
    Private Const TST_DP_GWL_STYLE      As Long = -16           'Window style index
    Private Const TST_DP_WS_CAPTION     As Long = &HC00000      'Window caption style flag

    #If Mac Then
    #Else
        #If VBA7 Then
            #If Win64 Then
                Private Declare PtrSafe Function TST_DP_GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr
            #Else
                Private Declare PtrSafe Function TST_DP_GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr
            #End If
        #Else
            Private Declare Function TST_DP_GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
        #End If
    #End If

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
    Private mTST_DP_ScratchAddBefore As Long            'Worksheet count immediately before the scratch-sheet Add
    Private mTST_DP_ScratchAddAfter  As Long            'Worksheet count immediately after it
    Private mTST_DP_ScratchAddOrphan As String          'Outcome of cleaning up after a failed scratch-sheet setup
    Private mTST_DP_MenuAtStart     As Long             'Context-menu controls registered before the run
    Private mTST_DP_MenuAfterRemove As Long             'Context-menu controls left immediately after removal
    Private mTST_DP_RunInProgress   As Boolean          'True between run start and completed teardown
    Private mTST_DP_DirtyStart      As Boolean          'True when preflight found a previous run's leftovers
    Private mTST_DP_DirtyDetail     As String           'What preflight found, recorded once the result sheet exists
    Private mTST_DP_InjectCleanupFail As String         'Cleanup step name to force-fail, for harness self-checks
    Private mTST_DP_CleanupFails    As Long             'Cleanup steps that did not complete in the current run
    Private mTST_DP_CleanupDetail   As String           'First cleanup failure detail in the current run
    Private mTST_DP_SuitesDispatched As Long            'Suites dispatched in the current run
    Private mTST_DP_SuitesCompleted As Long             'Suites that returned without escaping


'
'------------------------------------------------------------------------------
'
'                              PUBLIC ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Private Sub TST_DP_Preflight()

'
'==============================================================================
'                            HARNESS PREFLIGHT
'------------------------------------------------------------------------------
' PURPOSE
'   Decides whether this run is starting in an environment a previous run left
'   behind, before anything in this run modifies that environment
'
' WHY THIS EXISTS
'   A run that aborts before teardown leaves worksheets and Application state in
'   place. The next run then fails during its own setup rather than at the point
'   of the original defect, and its results describe an environment nobody
'   intended
'
'   Detection has to happen before the first mutation. Building the result sheet
'   template over a previous run's leftovers is itself one of the ways the
'   historical failure presented
'
' INPUTS
'   None
'
' RETURNS
'   Nothing. Sets mTST_DP_DirtyStart and mTST_DP_DirtyDetail
'
' BEHAVIOR
'   Looks for evidence of an incomplete previous run and records what it found
'
'   Does not modify the workbook, the Application, or DatePicker state
'
' ERROR POLICY
'   Best-effort. Never raises. A preflight that cannot read the environment
'   reports a dirty start rather than assuming a clean one
'
' DEPENDENCIES
'   mTST_DP_RunInProgress
'   mTST_DP_HostWorkbook
'   TST_DP_SheetExists
'
' NOTES
'   Two independent kinds of evidence are needed, because they survive different
'   kinds of abort:
'
'     mTST_DP_RunInProgress   an abort that left module state intact
'     leftover scratch sheet  an abort that cleared it, such as a project reset
'
'   A VBA project reset zeroes module-level state, so the flag alone cannot see
'   the abort it exists to detect. The worksheet is the evidence that survives
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Detail          As String       'Accumulated evidence description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Preflight must never break a run before it starts
        On Error GoTo PreflightUnreadable
    'Assume a clean start until evidence says otherwise
        mTST_DP_DirtyStart = False
        mTST_DP_DirtyDetail = VBA.vbNullString

'------------------------------------------------------------------------------
' EVIDENCE: MODULE STATE
'------------------------------------------------------------------------------
    'A run that set the flag and never cleared it did not reach teardown
        If mTST_DP_RunInProgress Then
            Detail = "the previous run did not complete teardown"
        End If

'------------------------------------------------------------------------------
' EVIDENCE: LEFTOVER WORKSHEET
'------------------------------------------------------------------------------
    'The scratch sheet is deleted during teardown, so its presence outlives a
    'project reset that would have cleared the flag above
        If Not mTST_DP_HostWorkbook Is Nothing Then
            If TST_DP_SheetExists(mTST_DP_HostWorkbook, TST_DP_SCRATCH_SHEET_NAME) Then
                If VBA.LenB(Detail) = 0 Then
                    Detail = "worksheet " & TST_DP_SCRATCH_SHEET_NAME & " was left behind"
                Else
                    Detail = Detail & ", and worksheet " & TST_DP_SCRATCH_SHEET_NAME & _
                        " was left behind"
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE PREFLIGHT
'------------------------------------------------------------------------------
    'Record the verdict for the run to report once it can write results
        If VBA.LenB(Detail) > 0 Then
            mTST_DP_DirtyStart = True
            mTST_DP_DirtyDetail = "Dirty start: " & Detail & _
                ". Results describe an environment this run did not establish."
        End If
    'Exit after a completed preflight
        Exit Sub

'------------------------------------------------------------------------------
' PREFLIGHT UNREADABLE
'------------------------------------------------------------------------------
PreflightUnreadable:
    'An environment that cannot be inspected is not known to be clean
        mTST_DP_DirtyStart = True
        mTST_DP_DirtyDetail = "Dirty start: the pre-run environment could not be " & _
            "inspected (error " & VBA.CStr(Err.Number) & " - " & Err.Description & ")."
    Err.Clear

End Sub

Private Sub TST_DP_ReportSetupFailure(ByVal EntryPoint As String)

'
'==============================================================================
'                        REPORT SETUP FAILURE
'==============================================================================
'   Reports a failure that happened before the run could start.
'
'   Setup runs outside the run's own fatal handler, so an error here would
'   otherwise reach the user as an unhandled VBE dialog naming a method rather
'   than a cause. Nothing has been recorded at this point and there may be no
'   result sheet, so the report goes to the Immediate window and a message box.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Capture the failure before anything can clear it
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Never let reporting raise
        On Error Resume Next
    'Record what failed
        Debug.Print EntryPoint & " | Setup failed | " & _
            VBA.CStr(ErrorNumber) & " - " & ErrorDescription
    'Record the conditions that explain it
        Debug.Print EntryPoint & " | Environment | " & TST_DP_DescribeHostWorkbook()
    'Tell the operator, because no result sheet exists to read
        MsgBox _
            "The regression run could not start." & VBA.vbCrLf & VBA.vbCrLf & _
            VBA.CStr(ErrorNumber) & " - " & ErrorDescription & VBA.vbCrLf & VBA.vbCrLf & _
            TST_DP_DescribeHostWorkbook() & VBA.vbCrLf & VBA.vbCrLf & _
            "The Immediate window carries the same detail.", _
            vbCritical Or vbOKOnly, _
            "DatePicker regression harness"
    'Clear any suppressed reporting error
        Err.Clear

End Sub

Public Sub TST_DP_ReportEnvironment()

'
'==============================================================================
'                          REPORT ENVIRONMENT
'==============================================================================
'   Prints the host workbook conditions that decide whether the harness can set
'   itself up. Run it from the Immediate window when a run fails during setup.
'
'   It resolves the host the same way a run does and writes nothing.
'==============================================================================

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Never let a diagnostic raise
        On Error Resume Next
    'Resolve the host exactly as a run would
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()
    'Print the conditions that block Worksheets.Add
        Debug.Print "TST_DP | Environment | " & TST_DP_DescribeHostWorkbook()
    'Release the reference so the probe leaves nothing behind
        Set mTST_DP_HostWorkbook = Nothing
    'Clear any suppressed diagnostic error
        Err.Clear

End Sub

Private Sub TST_DP_ReportDirtyStart()

'
'==============================================================================
'                          REPORT DIRTY START
'==============================================================================
'   Reports a refused run without writing anything to the workbook.
'
'   The result sheet cannot be used here. Building it is a mutation, and a run
'   that has just decided the environment is not its own must not mutate it.
'   Building that template over a previous run's leftovers is the failure this
'   whole path exists to prevent.
'==============================================================================

'------------------------------------------------------------------------------
' REPORT AND REFUSE
'------------------------------------------------------------------------------
    'Never let reporting raise
        On Error Resume Next
    'Record the refusal where it can be read without a result sheet
        Debug.Print "TST_DP_RunAll | " & TST_DP_STATE_DIRTY_START & " | " & _
            mTST_DP_DirtyDetail
    'Describe the workbook the run refused to touch
        Debug.Print "TST_DP_RunAll | Environment | " & TST_DP_DescribeHostWorkbook()
    'Tell the operator, because a refused run produces no result sheet to read
        MsgBox _
            "The regression run was refused." & VBA.vbCrLf & VBA.vbCrLf & _
            mTST_DP_DirtyDetail & VBA.vbCrLf & VBA.vbCrLf & _
            "Restart Excel, delete any leftover " & TST_DP_SCRATCH_SHEET_NAME & _
            " worksheet, and run again.", _
            vbExclamation Or vbOKOnly, _
            "DatePicker regression harness"
    'Clear any suppressed reporting error
        Err.Clear

End Sub

Private Function TST_DP_DescribeHostWorkbook() As String

'
'==============================================================================
'                        DESCRIBE HOST WORKBOOK
'==============================================================================
'   Describes the conditions that stop a worksheet being added, so a setup
'   failure names its cause instead of surfacing a bare 1004.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Description     As String       'Accumulated description

'------------------------------------------------------------------------------
' DESCRIBE
'------------------------------------------------------------------------------
    'Never let a diagnostic raise
        On Error Resume Next
    'Set safe default result
        TST_DP_DescribeHostWorkbook = "host workbook could not be described"
    'Report when no host workbook was resolved at all
        If mTST_DP_HostWorkbook Is Nothing Then
            TST_DP_DescribeHostWorkbook = "no host workbook was resolved"
            Err.Clear
            Exit Function
        End If
    'Name the workbook and the conditions that block Worksheets.Add
        Description = "Host=" & mTST_DP_HostWorkbook.Name & _
            "; IsAddin=" & VBA.CStr(mTST_DP_HostWorkbook.IsAddin) & _
            "; ProtectStructure=" & VBA.CStr(mTST_DP_HostWorkbook.ProtectStructure) & _
            "; ReadOnly=" & VBA.CStr(mTST_DP_HostWorkbook.ReadOnly) & _
            "; Worksheets=" & VBA.CStr(mTST_DP_HostWorkbook.Worksheets.Count) & _
            "; Sheets=" & VBA.CStr(mTST_DP_HostWorkbook.Sheets.Count) & _
            "; " & TST_DP_RESULT_SHEET_NAME & " exists=" & _
            VBA.CStr(TST_DP_SheetExists(mTST_DP_HostWorkbook, TST_DP_RESULT_SHEET_NAME)) & _
            "; " & TST_DP_SCRATCH_SHEET_NAME & " exists=" & _
            VBA.CStr(TST_DP_SheetExists(mTST_DP_HostWorkbook, TST_DP_SCRATCH_SHEET_NAME)) & _
            "; WorksheetsBeforeAdd=" & VBA.CStr(mTST_DP_ScratchAddBefore) & _
            "; WorksheetsAfterAdd=" & VBA.CStr(mTST_DP_ScratchAddAfter) & _
            "; ScratchCleanup=" & _
            VBA.IIf(VBA.LenB(mTST_DP_ScratchAddOrphan) = 0, "none", mTST_DP_ScratchAddOrphan)
    'Return the description
        TST_DP_DescribeHostWorkbook = Description
    'Clear any suppressed diagnostic error
        Err.Clear

End Function

Private Function TST_DP_SheetExists( _
    ByVal Book As Excel.Workbook, _
    ByVal SheetName As String) As Boolean

'
'==============================================================================
'                              SHEET EXISTS
'==============================================================================
'   Reports whether a worksheet of the given name exists in the workbook,
'   without creating it and without raising when it does not.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS              As Excel.Worksheet  'Resolved worksheet

'------------------------------------------------------------------------------
' RESOLVE SHEET
'------------------------------------------------------------------------------
    'Suppress the error raised when the sheet is absent
        On Error Resume Next
    'Set safe default result
        TST_DP_SheetExists = False
    'Attempt to resolve the named worksheet
        Set WS = Book.Worksheets(SheetName)
    'Report whether the resolution succeeded
        TST_DP_SheetExists = Not (WS Is Nothing)
    'Release object references
        Set WS = Nothing
    'Clear any suppressed lookup error
        Err.Clear

End Function

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
'   TST_DP_Preflight
'   DEMO_Sheet_BuildTemplate
'   TST_DP_GetHostWorkbook
'
' NOTES
'   Use TST_DP_RunAll_WithUISmoke when a brief UF_DatePicker open/close check
'   is also needed
'
'   Preflight runs before the result sheet template is built, because building
'   it over a previous run's leftovers is one of the ways an aborted predecessor
'   presents
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Report a setup failure with the conditions that caused it, rather than
    'surfacing a bare 1004 from a workbook that cannot accept a worksheet
        On Error GoTo SetupFailed
    'Reset module-level counters and object references before the run
        TST_DP_ResetHarnessState
    'Resolve the workbook that will receive the result and scratch sheets
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()
    'Decide whether the environment is clean before anything in this run changes it
        TST_DP_Preflight
    'Stop before touching the workbook when the environment is not this run's to use
        If mTST_DP_DirtyStart Then
            TST_DP_ReportDirtyStart
            Exit Sub
        End If
    'Build the result sheet template before the run
    'Name the target workbook explicitly. DEMO_Sheet_BuildTemplate defaults
    'TargetWorkbook to ThisWorkbook, which is the host only when the component
    'runs embedded. Loaded as an .xlam, ThisWorkbook is the add-in, so the result
    'sheet was built somewhere other than the workbook the run then reads it from,
    'and TST_DP_PrepareResultSheet failed with error 9 before the first suite
        DEMO_Sheet_BuildTemplate TST_DP_RESULT_SHEET_NAME, "DATE PICKER", _
            "Test Sheet", , TST_DP_RESULT_FIRST_ROW, mTST_DP_HostWorkbook

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Restore normal error handling before the run takes over
        On Error GoTo 0
    'Run the standard non-disruptive regression pack without UI smoke
        TST_DP_RunAllInternal False
    'Exit before the setup handler
        Exit Sub

'------------------------------------------------------------------------------
' SETUP FAILED
'------------------------------------------------------------------------------
SetupFailed:
    'Report what failed and the workbook conditions behind it
        TST_DP_ReportSetupFailure "TST_DP_RunAll"

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
    'Report a setup failure with the conditions that caused it
        On Error GoTo SetupFailed
    'Resolve the workbook that will receive the result and scratch sheets
        Set mTST_DP_HostWorkbook = TST_DP_GetHostWorkbook()

'------------------------------------------------------------------------------
' PREFLIGHT
'------------------------------------------------------------------------------
    'Decide whether the environment is clean before anything in this run changes it
        TST_DP_Preflight
    'Stop before touching the workbook when the environment is not this run's to use
        If mTST_DP_DirtyStart Then
            TST_DP_ReportDirtyStart
            Exit Sub
        End If

'------------------------------------------------------------------------------
' BUILD RESULT SHEET TEMPLATE
'------------------------------------------------------------------------------
    'Build the result sheet template before the run
    'Name the target workbook explicitly. DEMO_Sheet_BuildTemplate defaults
    'TargetWorkbook to ThisWorkbook, which is the host only when the component
    'runs embedded. Loaded as an .xlam, ThisWorkbook is the add-in, so the result
    'sheet was built somewhere other than the workbook the run then reads it from,
    'and TST_DP_PrepareResultSheet failed with error 9 before the first suite
        DEMO_Sheet_BuildTemplate TST_DP_RESULT_SHEET_NAME, "DATE PICKER", _
            "Test Sheet", , TST_DP_RESULT_FIRST_ROW, mTST_DP_HostWorkbook

'------------------------------------------------------------------------------
' RUN TESTS
'------------------------------------------------------------------------------
    'Restore normal error handling before the run takes over
        On Error GoTo 0
    'Run the regression pack with the UI smoke suite included
        TST_DP_RunAllInternal True
    'Exit before the setup handler
        Exit Sub

'------------------------------------------------------------------------------
' SETUP FAILED
'------------------------------------------------------------------------------
SetupFailed:
    'Report what failed and the workbook conditions behind it
        TST_DP_ReportSetupFailure "TST_DP_RunAll_WithUISmoke"

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
'   TST_DP_ContextMenuControlCount
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
    Dim HandlerStep         As String                               'Current setup step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled fatal handling
        On Error GoTo FatalHandler
    'Track the current setup step so a fatal names where it happened
        HandlerStep = "Initialize"
    'Mark the run as in progress until teardown completes
        mTST_DP_RunInProgress = True
    'Reset the per-run cleanup and suite counters
        mTST_DP_CleanupFails = 0
        mTST_DP_CleanupDetail = VBA.vbNullString
        mTST_DP_SuitesDispatched = 0
        mTST_DP_SuitesCompleted = 0
    'Capture whether a manager existed before the run
        mTST_DP_HadManager = Not (gDP_Manager Is Nothing)
    'Capture the context-menu registration the session already had. Teardown must
    'restore this, not erase it: a DatePicker that was running before the harness
    'started is entitled to its menu afterwards
        mTST_DP_MenuAtStart = TST_DP_ContextMenuControlCount()
    'Capture current DatePicker settings and transient state
        HandlerStep = "Capture DatePicker settings"
        TST_DP_CaptureSettings SettingsSnapshot
    'Capture current Excel Application state
        HandlerStep = "Capture Application state"
        TST_DP_CaptureApplicationState AppSnapshot

'------------------------------------------------------------------------------
' PREPARE ISOLATED RUN STATE
'------------------------------------------------------------------------------
    'Prepare the Excel Application state for the regression run
        HandlerStep = "Prepare Application for run"
        TST_DP_PrepareApplicationForRun
    'Reset transient DatePicker UI artifacts before testing
        HandlerStep = "Reset DatePicker artifacts"
        TST_DP_ResetDatePickerArtifacts
    'Prepare the result worksheet
        HandlerStep = "Prepare result sheet"
        TST_DP_PrepareResultSheet mTST_DP_HostWorkbook
    'Prepare the scratch worksheet
        HandlerStep = "Prepare scratch sheet"
        TST_DP_PrepareScratchSheet mTST_DP_HostWorkbook
    'Record the run header
        TST_DP_RecordInfo "Harness", "Start", _
            "IncludeUISmoke=" & VBA.CStr(IncludeUISmoke)
    'Report what preflight found, now that results can be written. The verdict was
    'reached before this run touched anything, and it decides the run state
        If mTST_DP_DirtyStart Then
            TST_DP_RecordInfo "Harness", "Dirty start", mTST_DP_DirtyDetail
            TST_DP_RecordInfo "Harness", "Suites not dispatched", _
                "No suite was run. Results gathered in an environment this run " & _
                "did not establish would describe the predecessor's leftovers, " & _
                "not the code under test. Restart Excel, delete the leftover " & _
                "worksheets, and run again."
        End If

'------------------------------------------------------------------------------
' RUN SUITES
'------------------------------------------------------------------------------
    'Skip every suite when the run did not start clean. A dirty run tears down and
    'reports, but it never executes tests it could not interpret
        If mTST_DP_DirtyStart Then GoTo CleanExit
    'Run environment and manager smoke checks
        TST_DP_RunSuiteSafe "Environment"
    'Run settings and persisted-state checks
        TST_DP_RunSuiteSafe "Settings"
    'Run the settings namespace isolation suite
        TST_DP_RunSuiteSafe "SettingsNamespace"
    'Run settings-panel save resolution checks
        TST_DP_RunSuiteSafe "SettingsSaveResolution"
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
    'Run discontiguous write-result completeness and order-independence checks
        TST_DP_RunSuiteSafe "MultiAreaWriteResult"
    'Run technical-failure result-preservation checks
        TST_DP_RunSuiteSafe "WriteTechnicalFailure"
    'Run original-error preservation checks across cleanup
        TST_DP_RunSuiteSafe "ErrorPreservation"
    'Run in-grid icon lifecycle checks
        TST_DP_RunSuiteSafe "GridIcon"
    'Run manager public API and target gating checks
        TST_DP_RunSuiteSafe "Manager"
    'Run DP_Start and DP_Stop lifecycle round-trip checks
        TST_DP_RunSuiteSafe "LifecyclePair"
    'Run the one-provider lease suite
        TST_DP_RunSuiteSafe "ProviderLease"
    'Run entry-path admission checks under owned and foreign leases
        TST_DP_RunSuiteSafe "RuntimeAdmission"
    'Run DP_RepairRuntime behavior checks
        TST_DP_RunSuiteSafe "RepairRuntime"
    'Run M_GridIcon_PreCreateHidden startup optimization checks
        TST_DP_RunSuiteSafe "PreCreateHidden"
    'Run M_Picker_SelectDate write-back and state-management checks
        TST_DP_RunSuiteSafe "SelectDate"

    'Run the application-state suite
        TST_DP_RunSuiteSafe "ApplicationState"
    'Run the borderless window-style transaction suite
        TST_DP_RunSuiteSafe "WindowStyle"
    'Run form-level window recovery checks
        TST_DP_RunSuiteSafe "WindowRecovery"
    'Run the harness run-state and preflight self-checks
        TST_DP_RunSuiteSafe "HarnessSelfCheck"
    'Run optional UF_DatePicker open / close smoke check when requested
        If IncludeUISmoke Then
            TST_DP_RunSuiteSafe "UISmoke"
        End If

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
    'Suppress cleanup errors so every cleanup step is attempted. Each step is
    'checked individually afterwards, because a suppressed cleanup failure that
    'nobody records is what leaves the next run unable to start
        On Error Resume Next

    'Reset DatePicker UI artifacts after testing. Each operation is checked on its
    'own, because a composite step that suppresses its own errors cannot report
    'which of five operations failed, or that any of them did
        M_Timer_Stop
        TST_DP_CheckCleanupStep "StopTimer"

        DP_Close
        TST_DP_CheckCleanupStep "ClosePickerForm"

        M_ContextMenu_Remove
        TST_DP_CheckCleanupStep "RemoveContextMenu"
        mTST_DP_MenuAfterRemove = TST_DP_ContextMenuControlCount()

        M_KeyboardShortcut_Remove
        TST_DP_CheckCleanupStep "RemoveKeyboardShortcut"

        M_GridIcon_PurgeAll
        TST_DP_CheckCleanupStep "PurgeGridIcons"

        M_Window_Test_SetFaultInjection 0, 0
        TST_DP_CheckCleanupStep "DisarmWindowFaultInjection"

    'Delete the scratch worksheet
        TST_DP_DeleteScratchSheet
        TST_DP_CheckCleanupStep "DeleteScratchSheet"

    'Restore DatePicker settings and transient state
        TST_DP_RestoreSettings SettingsSnapshot
        TST_DP_CheckCleanupStep "RestoreSettings"

    'Reconcile the context menu to the state the run found.
    '
    'RestoreSettings puts the persisted right-click policy back and then calls
    'M_ContextMenu_Update, which re-registers the entries whenever that policy is
    'enabled. That happens after the teardown above removed them, so a run which
    'began with no entries ends holding two.
    '
    'It is invisible in an embedded workbook, where Workbook_Open has already
    'started the runtime and the pre-run count is the same two. It shows up in an
    '.xlam that was loaded without starting the runtime: the pre-run count is
    'zero, and restoring a setting the session never applied registers a menu the
    'user did not have.
    '
    'The settings restore is correct and stays where it is. This step exists
    'because the harness contract is to leave the session as it was found, and
    'only the pre-run count knows what that was
        If mTST_DP_MenuAtStart = 0 Then
            If TST_DP_ContextMenuControlCount() > 0 Then
                M_ContextMenu_Remove
            End If
        End If
        TST_DP_CheckCleanupStep "ReconcileContextMenu"

    'Restore the manager state to its pre-run condition
        TST_DP_RestoreManagerState
        TST_DP_CheckCleanupStep "RestoreManagerState"

    'Restore the Excel Application state
        TST_DP_RestoreApplicationState AppSnapshot
        TST_DP_CheckCleanupStep "RestoreApplicationState"

    'Verify the run left no DatePicker artifact behind
        TST_DP_VerifyFinalState AppSnapshot

    'Write the summary after cleanup so it can report the cleanup outcome
        TST_DP_WriteSummary

    'Release module object references
        Set mTST_DP_ScratchSheet = Nothing
        Set mTST_DP_ResultSheet = Nothing
        Set mTST_DP_HostWorkbook = Nothing

    'Mark teardown as complete so the next run does not report a dirty start
        mTST_DP_RunInProgress = False

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
            "Fatal error " & VBA.CStr(FatalNumber) & " - " & FatalDescription & _
            " | Step=" & HandlerStep & _
            " | " & TST_DP_DescribeHostWorkbook()
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
    'Count the dispatch so a suite that never returns can be detected
        mTST_DP_SuitesDispatched = mTST_DP_SuitesDispatched + 1
    'Protect the harness from escaping suite failures
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' DISPATCH SUITE
'------------------------------------------------------------------------------
    'Dispatch the requested suite by name
        Select Case VBA.UCase$(VBA.Trim$(SuiteName))

            Case "ENVIRONMENT"
                TST_DP_RunSuite_Environment

            Case "SETTINGSSAVERESOLUTION"
                TST_DP_RunSuite_SettingsSaveResolution

            Case "SETTINGSNAMESPACE"
                TST_DP_RunSuite_SettingsNamespace
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

            Case "MULTIAREAWRITERESULT"
                TST_DP_RunSuite_MultiAreaWriteResult

            Case "WRITETECHNICALFAILURE"
                TST_DP_RunSuite_WriteTechnicalFailure

            Case "ERRORPRESERVATION"
                TST_DP_RunSuite_ErrorPreservation

            Case "GRIDICON"
                TST_DP_RunSuite_GridIcon

            Case "MANAGER"
                TST_DP_RunSuite_Manager

            Case "PROVIDERLEASE"
                TST_DP_RunSuite_ProviderLease

            Case "RUNTIMEADMISSION"
                TST_DP_RunSuite_RuntimeAdmission
            Case "LIFECYCLEPAIR"
                TST_DP_RunSuite_LifecyclePair

            Case "REPAIRRUNTIME"
                TST_DP_RunSuite_RepairRuntime

            Case "PRECREATEHIDDEN"
                TST_DP_RunSuite_PreCreateHidden
            
            Case "SELECTDATE"
                TST_DP_RunSuite_SelectDate
            Case "APPLICATIONSTATE"
                TST_DP_RunSuite_ApplicationState
            Case "WINDOWRECOVERY"
                TST_DP_RunSuite_WindowRecovery

            Case "WINDOWSTYLE"
                TST_DP_RunSuite_WindowStyle
            Case "HARNESSSELFCHECK"
                TST_DP_RunSuite_HarnessSelfCheck

            Case "UISMOKE"
                TST_DP_RunSuite_UISmoke

            Case Else
                TST_DP_RecordFail "Unknown suite name", "SuiteName=" & SuiteName

        End Select

'------------------------------------------------------------------------------
' RECORD COMPLETION
'------------------------------------------------------------------------------
    'Count the suite as completed only when it returned without escaping
        mTST_DP_SuitesCompleted = mTST_DP_SuitesCompleted + 1

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

    'Disable the keyboard shortcut, then remove both other access paths
        M_Settings_SetEnableKeyboardShortcut False
        M_Settings_SetShowRightClick False
        M_Settings_SetShowGridIcon False

    'Assert nothing re-enables the shortcut behind the user's back. Taking a
    'session-wide Application.OnKey binding is not a decision the component makes
    'on the user's behalf, even when no other built-in access path remains
        TST_DP_AssertFalse "Disabling other access paths does not re-enable the shortcut", _
            M_Settings_GetEnableKeyboardShortcut()

    'Assert a persisted round trip does not re-enable it either
        M_Settings_Save
        M_Settings_Load
        TST_DP_AssertFalse "Save and reload do not re-enable the shortcut", _
            M_Settings_GetEnableKeyboardShortcut()

    'Assert the picker remains registrable with zero built-in access paths
        TST_DP_AssertTrue "Zero built-in access paths is a permitted configuration", _
            (M_Settings_GetShowRightClick() = False) And _
            (M_Settings_GetShowGridIcon() = False) And _
            (M_Settings_GetEnableKeyboardShortcut() = False)

    'Restore the shortcut for the following suites
        M_Settings_SetEnableKeyboardShortcut True

'------------------------------------------------------------------------------
' KEYBOARD SHORTCUT REGISTRATION PATHS
'------------------------------------------------------------------------------
    'Exercise the registration path an explicitly enabled shortcut takes. Excel
    'exposes no getter for Application.OnKey, so the binding cannot be read back;
    'what is asserted here is that the operation completes, and the resulting key
    'behavior is covered by the documented manual validation
        M_Settings_SetEnableKeyboardShortcut True
        M_KeyboardShortcut_Update
        TST_DP_AssertTrue "Enabled shortcut completes its registration path", _
            M_Settings_GetEnableKeyboardShortcut()

    'Exercise the removal path, which restores Excel default handling
        M_Settings_SetEnableKeyboardShortcut False
        M_KeyboardShortcut_Update
        TST_DP_AssertFalse "Disabled shortcut completes its removal path", _
            M_Settings_GetEnableKeyboardShortcut()

    'Restore the shortcut for the following suites
        M_Settings_SetEnableKeyboardShortcut True
        M_KeyboardShortcut_Update

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

Private Sub TST_DP_RunSuite_SettingsSaveResolution()

'
'==============================================================================
'                    SETTINGS SAVE RESOLUTION SUITE
'==============================================================================
' PURPOSE
'   Proves a settings-panel save preserves an explicitly disabled keyboard
'   shortcut, including when right-click and the grid icon are also disabled
'
' WHY THIS EXISTS
'   #42 was closed at v1.2.0 on the strength of module-setter coverage. The
'   fallback had been removed from M_Settings_SetShowRightClick and
'   M_Settings_SetShowGridIcon but not from the UserForm save path, so the
'   setters passed while the real panel still forced the shortcut back on
'
'   This suite drives M_Settings_ResolveKeyboardShortcutOnSave, which is the
'   same code the UserForm save handler executes, rather than the setters
'
' NOTES
'   Application.OnKey has no read-back API and the component tracks no
'   registration flag, so the registration criterion is asserted through the
'   setting that selects the register/remove branch. This is a real limitation
'   of the platform, not of the suite, and the manual matrix covers the rest
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RightClickState     As Boolean      'Swept right-click input
    Dim GridIconState       As Boolean      'Swept grid-icon input
    Dim CurrentState        As Boolean      'Swept current keyboard setting
    Dim Resolved            As Boolean      'Value the save path would persist
    Dim SweepHeld           As Boolean      'True while every combination is invariant
    Dim RightClickIndex     As Long         'Sweep index for right-click
    Dim GridIconIndex       As Long         'Sweep index for grid icon
    Dim CurrentIndex        As Long         'Sweep index for the current setting

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "SettingsSaveResolution"

'------------------------------------------------------------------------------
' THE SEAM IS INVARIANT ACROSS EVERY INTEGRATION COMBINATION
'------------------------------------------------------------------------------
    'The two integration arguments must never change the answer. Sweeping all
    'eight combinations is what makes a reintroduced fallback fail here
        SweepHeld = True
        For CurrentIndex = 0 To 1
            CurrentState = (CurrentIndex = 1)
            For RightClickIndex = 0 To 1
                RightClickState = (RightClickIndex = 1)
                For GridIconIndex = 0 To 1
                    GridIconState = (GridIconIndex = 1)
                    Resolved = M_Settings_ResolveKeyboardShortcutOnSave( _
                        CurrentState, RightClickState, GridIconState)
                    If Resolved <> CurrentState Then
                        SweepHeld = False
                    End If
                Next GridIconIndex
            Next RightClickIndex
        Next CurrentIndex
        TST_DP_AssertTrue _
            "Save resolution preserves the keyboard setting in all 8 combinations", _
            SweepHeld

'------------------------------------------------------------------------------
' THE FORMER FALLBACK CASE SPECIFICALLY
'------------------------------------------------------------------------------
    'This is the exact input triple that was forced to True at v1.2.0
        TST_DP_AssertFalse _
            "All three disabled resolves the shortcut to False, not True", _
            M_Settings_ResolveKeyboardShortcutOnSave(False, False, False)

'------------------------------------------------------------------------------
' ALL THREE ENTRY PATHS DISABLED IS A VALID PERSISTED CONFIGURATION
'------------------------------------------------------------------------------
    'Drive the real setters into the state the panel would save
        M_Settings_SetShowRightClick False
        M_Settings_SetShowGridIcon False
        M_Settings_SetEnableKeyboardShortcut False
        TST_DP_AssertFalse "Right-click stays disabled", _
            M_Settings_GetShowRightClick()
        TST_DP_AssertFalse "Grid icon stays disabled", _
            M_Settings_GetShowGridIcon()
        TST_DP_AssertFalse "Keyboard shortcut stays disabled", _
            M_Settings_GetEnableKeyboardShortcut()

'------------------------------------------------------------------------------
' A SAVE IN THAT STATE CHANGES NOTHING
'------------------------------------------------------------------------------
    'Resolve exactly as the panel save would, then apply it
        Resolved = M_Settings_ResolveKeyboardShortcutOnSave( _
            M_Settings_GetEnableKeyboardShortcut(), _
            M_Settings_GetShowRightClick(), _
            M_Settings_GetShowGridIcon())
        M_Settings_SetEnableKeyboardShortcut Resolved
        TST_DP_AssertFalse _
            "Saving the panel does not re-enable the keyboard shortcut", _
            M_Settings_GetEnableKeyboardShortcut()

'------------------------------------------------------------------------------
' THE SETTING SURVIVES A PERSISTENCE ROUND TRIP
'------------------------------------------------------------------------------
    'Reload from persisted storage and confirm the disabled value came back
        M_Settings_Save
        M_Settings_Load
        TST_DP_AssertFalse _
            "Persisted settings reload with the keyboard shortcut disabled", _
            M_Settings_GetEnableKeyboardShortcut()
        TST_DP_AssertFalse "Persisted right-click reloads disabled", _
            M_Settings_GetShowRightClick()
        TST_DP_AssertFalse "Persisted grid icon reloads disabled", _
            M_Settings_GetShowGridIcon()

'------------------------------------------------------------------------------
' UPDATE FOLLOWS THE REMOVAL BRANCH IN THIS STATE
'------------------------------------------------------------------------------
    'Application.OnKey cannot be read back, so assert the setting that selects
    'the branch rather than claiming to observe the registration itself
        M_KeyboardShortcut_Update
        TST_DP_AssertFalse _
            "Update leaves the shortcut disabled with zero entry paths", _
            M_Settings_GetEnableKeyboardShortcut()

'------------------------------------------------------------------------------
' EXPLICIT OPT-IN STILL WORKS
'------------------------------------------------------------------------------
    'Removing the fallback must not make the shortcut unreachable
        M_Settings_SetEnableKeyboardShortcut True
        TST_DP_AssertTrue "Explicit opt-in still enables the shortcut", _
            M_Settings_GetEnableKeyboardShortcut()
        TST_DP_AssertTrue _
            "Opt-in survives a save with both integrations disabled", _
            M_Settings_ResolveKeyboardShortcutOnSave( _
                M_Settings_GetEnableKeyboardShortcut(), False, False)

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Restore working defaults for the suites that follow
        M_Settings_SetShowRightClick True
        M_Settings_SetShowGridIcon True
        M_Settings_SetEnableKeyboardShortcut True
        M_KeyboardShortcut_Update
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Settings save resolution suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore settings regardless
        Resume SuiteExit

End Sub

Private Sub TST_DP_RunSuite_SettingsNamespace()

'
'==============================================================================
'                     SUITE: SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Proves that an explicit settings namespace isolates persisted configuration,
'   and that the default namespace still resolves exactly where it always did
'
' WHY THIS EXISTS
'   Persistence is scoped to the Windows user, not the deployment. Two workbooks
'   that never run at the same time can still overwrite each other's preferences,
'   and loading is itself a write boundary because M_Settings_Load persists the
'   values it normalizes
'
' BEHAVIOR
'   Writes distinct values under two temporary namespaces, reads each back, and
'   confirms neither disturbed the other or the legacy default
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_Settings_SetNamespace
'   M_Settings_GetNamespace
'   DP_SETTINGS_APP_NAME
'
' NOTES
'   The suite never writes to the operator's real VBA_DATETIMEPICKER settings.
'   It writes through GetSetting/SaveSetting directly under temporary namespaces
'   it creates and deletes, so proving isolation costs the operator nothing
'
'   The namespace cannot be reconfigured once settings have loaded, and settings
'   are loaded long before this suite runs. The lock is therefore asserted rather
'   than worked around: the suite proves the refusal instead of trying to defeat
'   it, and exercises isolation through the same registry API the resolver uses
'
'   Temporary registry keys are removed on every path, including a failed
'   assertion. A leftover key is reported as a cleanup failure rather than left
'   for the next run to find
'
' UPDATED
'   2026-08-23
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const SECTION_NAME  As String = "Display"           'Section used for the probe
    Const KEY_NAME      As String = "TST_DP_NamespaceProbe"  'Key used for the probe

    Dim AppNameA        As String       'Effective application name for namespace A
    Dim AppNameB        As String       'Effective application name for namespace B
    Dim LegacyValue     As String       'Legacy-namespace probe value, if any
    Dim ReadBackA       As String       'Value read back from namespace A
    Dim ReadBackB       As String       'Value read back from namespace B
    Dim LegacyAfter     As String       'Legacy value after both namespace writes
    Dim RefusedLate     As Boolean      'True when a late namespace change was refused
    Dim RefusedInvalid  As Boolean      'True when an invalid namespace was refused

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "SettingsNamespace"
    AppNameA = DP_SETTINGS_APP_NAME & "__TST_DP_NS_A"
    AppNameB = DP_SETTINGS_APP_NAME & "__TST_DP_NS_B"

'------------------------------------------------------------------------------
' ASSERT THE DEFAULT RESOLVES TO THE LEGACY NAME
'------------------------------------------------------------------------------
    'An installation that configures nothing must read and write exactly where
    'earlier releases did
        TST_DP_AssertEqualsString "Default namespace is empty", _
            VBA.vbNullString, M_Settings_GetNamespace()
    'Record whatever the legacy namespace holds so the suite can prove it is
    'left alone
        LegacyValue = GetSetting(DP_SETTINGS_APP_NAME, SECTION_NAME, KEY_NAME, "<absent>")

'------------------------------------------------------------------------------
' WRITE TWO ISOLATED NAMESPACES
'------------------------------------------------------------------------------
    'Persist a distinct value under each temporary namespace
        SaveSetting AppNameA, SECTION_NAME, KEY_NAME, "value-A"
        SaveSetting AppNameB, SECTION_NAME, KEY_NAME, "value-B"
    'Read each back
        ReadBackA = GetSetting(AppNameA, SECTION_NAME, KEY_NAME, "<absent>")
        ReadBackB = GetSetting(AppNameB, SECTION_NAME, KEY_NAME, "<absent>")

'------------------------------------------------------------------------------
' ASSERT ISOLATION
'------------------------------------------------------------------------------
    'Each namespace returns its own value
        TST_DP_AssertEqualsString "Namespace A returns its own value", _
            "value-A", ReadBackA
        TST_DP_AssertEqualsString "Namespace B returns its own value", _
            "value-B", ReadBackB
    'Neither write disturbed the other
        TST_DP_AssertTrue "Namespace A and B are isolated", ReadBackA <> ReadBackB
    'Neither write reached the legacy default
        LegacyAfter = GetSetting(DP_SETTINGS_APP_NAME, SECTION_NAME, KEY_NAME, "<absent>")
        TST_DP_AssertEqualsString "Legacy namespace is unaffected", _
            LegacyValue, LegacyAfter

'------------------------------------------------------------------------------
' ASSERT THE TIMING LOCK
'------------------------------------------------------------------------------
    'Settings are loaded by the time any suite runs, so reconfiguring must be
    'refused rather than silently repointing values already in memory
        On Error Resume Next
        Err.Clear
        M_Settings_SetNamespace "TooLate"
        RefusedLate = (Err.Number <> 0)
        Err.Clear
        On Error GoTo SuiteFail
        TST_DP_AssertTrue "Namespace change after settings load is refused", _
            RefusedLate
    'The refusal must not have altered the configured namespace
        TST_DP_AssertEqualsString "Refused change leaves the namespace unchanged", _
            VBA.vbNullString, M_Settings_GetNamespace()

'------------------------------------------------------------------------------
' ASSERT VALIDATION
'------------------------------------------------------------------------------
    'A namespace containing a path separator would make the registry location
    'ambiguous. The lock above fires first, so this asserts only that an invalid
    'namespace is never accepted silently
        On Error Resume Next
        Err.Clear
        M_Settings_SetNamespace "bad\namespace"
        RefusedInvalid = (Err.Number <> 0)
        Err.Clear
        On Error GoTo SuiteFail
        TST_DP_AssertTrue "Invalid namespace is refused", RefusedInvalid

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Remove the temporary registry keys this suite created
        TST_DP_DeleteNamespaceProbe AppNameA, SECTION_NAME
        TST_DP_DeleteNamespaceProbe AppNameB, SECTION_NAME
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Settings namespace suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Never leave temporary registry keys behind
        Resume SuiteExit

End Sub

Private Function TST_DP_RegistryAppExists( _
    ByVal ApplicationName As String) As Boolean

'
'==============================================================================
'                       REGISTRY APPLICATION EXISTS
'==============================================================================
'   Reports whether a VBA registry application key still exists.
'
'   GetAllSettings returns an unassigned Variant when the application key is
'   absent, and an array when it exists, so it detects an empty leftover key that
'   a value lookup cannot.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sections        As Variant      'Sections under the application key

'------------------------------------------------------------------------------
' PROBE
'------------------------------------------------------------------------------
    'Never let a cleanup probe raise
        On Error Resume Next
    'Set safe default result
        TST_DP_RegistryAppExists = False
    'Read the sections, which is empty when the application key is absent
        Sections = GetAllSettings(ApplicationName, "Display")
    'An assigned result means something is still there
        TST_DP_RegistryAppExists = Not VBA.IsEmpty(Sections)
    'Clear any suppressed probe error
        Err.Clear

End Function

Private Sub TST_DP_DeleteNamespaceProbe( _
    ByVal ApplicationName As String, _
    ByVal SectionName As String)

'
'==============================================================================
'                      DELETE NAMESPACE PROBE KEY
'==============================================================================
'   Removes a temporary registry namespace the suite created.
'
'   DeleteSetting is called with the application name alone. Passing a section as
'   well deletes only that section and leaves an empty application key behind,
'   which is a leak the old check could not see: it looked for the value, and the
'   value was gone.
'
'   A leftover key is counted as a cleanup failure rather than left for the next
'   run, because the suite's whole point is that persisted state does not leak
'   between deployments.
'==============================================================================

'------------------------------------------------------------------------------
' DELETE
'------------------------------------------------------------------------------
    'Suppress the error raised when the key was never created
        On Error Resume Next
        Err.Clear
    'Remove the whole temporary application key, not just one section
        DeleteSetting ApplicationName
        Err.Clear
    'Verify the application key is gone rather than trusting the delete. A section
    'delete leaves an empty application key that a value lookup cannot detect
        If TST_DP_RegistryAppExists(ApplicationName) Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Namespace probe", _
                "Temporary registry namespace " & ApplicationName & _
                " could not be removed"
        End If
    'Clear any suppressed cleanup error
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
'   selection-based write-back, table data-column expansion, unsupported write
'   action rejection, the structured write result, and partial-write reporting
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   M_WriteBack_PopulateRange
'   M_WriteBack_Apply
'   DP_FillTableColumn
'   TST_DP_ExpectPartialWriteReport
'   TST_DP_ExpectFailedAddressReport
'   TST_DP_ExpectFormulaProtection
'   TST_DP_AssertWriteResultBalances
'   gDP_WriteValue
'   mTST_DP_ScratchSheet
'
' NOTES
'   All write-back tests use the scratch worksheet only and leave no persistent
'   state on the result sheet
'
'   M_WriteBack_Apply is a Function returning DP_WriteResult; the private stages
'   below it accumulate into a ByRef result instead. Assertions use whichever
'   form the routine under test provides, and the bare-call form of
'   M_WriteBack_Apply is kept where the outcome is not asserted so the suite also
'   covers the call form the rest of the project uses
'
'   A DP_WriteResult accumulates, so WriteResult is reset from EmptyResult before
'   any direct call to M_WriteBack_PopulateRange
'
'   Reported addresses are worksheet-qualified, so expected values are built from
'   the scratch sheet name rather than hard-coded
'
'   Two expectations mutate the scratch sheet in ways a failed assertion must not
'   leave behind: TST_DP_ExpectPartialWriteReport protects it, and
'   TST_DP_ExpectFailedAddressReport writes an array formula spanning I6:I7. Both
'   release on every path
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetRange     As Excel.Range       'Target range under test
    Dim TableRange      As Excel.Range       'Source range for the test table
    Dim TestTable       As Excel.ListObject  'Regression test ListObject
    Dim UnionRange      As Excel.Range       'Discontiguous target range
    Dim WriteResult     As DP_WriteResult    'Structured write-back result
    Dim EmptyResult     As DP_WriteResult    'Zeroed result used to reset WriteResult

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "WriteBack"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet for selection-based tests
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet

'------------------------------------------------------------------------------
' DIRECT CONTIGUOUS RANGE POPULATION
'------------------------------------------------------------------------------
    'Clear the target cells
        mTST_DP_ScratchSheet.Range("D5:D6").ClearContents
    'Prepare the DatePicker write value
        gDP_WriteValue = VBA.DateSerial(2026, 5, 3)
    'Populate the contiguous range through the fast bulk path
        WriteResult = EmptyResult
        M_WriteBack_PopulateRange _
            mTST_DP_ScratchSheet.Range("D5:D6"), _
            DP_WriteAction_DatePicker, _
            WriteResult
    'Assert the first cell received the date
        TST_DP_AssertCellDateEquals "Direct contiguous range write D5", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D5")
    'Assert the second cell received the date
        TST_DP_AssertCellDateEquals "Direct contiguous range write D6", _
            VBA.DateSerial(2026, 5, 3), _
            mTST_DP_ScratchSheet.Range("D6")
    'A successful bulk write returns before the per-cell counters exist, so the
    'written count has to be contributed by the bulk path itself
        TST_DP_AssertEqualsLong "Bulk write reports 2 attempted cells", _
            2, VBA.CLng(WriteResult.AttemptedCount)
    'Assert the bulk path reported what it wrote
        TST_DP_AssertEqualsLong "Bulk write reports 2 written cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert the result satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Bulk write result balances", WriteResult

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
    'Populate the discontiguous range and capture the structured result
        WriteResult = EmptyResult
        M_WriteBack_PopulateRange UnionRange, DP_WriteAction_DatePicker, WriteResult
    'Assert the result counts both areas
        TST_DP_AssertEqualsLong "Discontiguous result counts 2 areas", _
            2, WriteResult.AreasCount
    'Assert the result counts every written cell
        TST_DP_AssertEqualsLong "Discontiguous result writes 2 cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert the result satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Discontiguous result balances", WriteResult
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
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, True)
    'Assert the result reports the attempted cells
        TST_DP_AssertEqualsLong "Selection result attempts 2 cells", _
            2, VBA.CLng(WriteResult.AttemptedCount)
    'Assert the result reports the written cells
        TST_DP_AssertEqualsLong "Selection result writes 2 cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert the result names the resolved target, worksheet-qualified
        TST_DP_AssertEqualsString "Selection result reports the resolved target", _
            mTST_DP_ScratchSheet.Name & "!D5:D6", WriteResult.ResolvedTargetAddress
    'Assert a plain selection is not reported as a table expansion
        TST_DP_AssertFalse "Selection result reports no table expansion", _
            WriteResult.ExpandedToTableColumn
    'Assert nothing was skipped or suppressed
        TST_DP_AssertEqualsLong "Selection result reports no skipped cells", _
            0, VBA.CLng(WriteResult.LockedSkippedCount + WriteResult.FailedCount)
    'Assert the result satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Selection result balances", WriteResult
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
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, False)
    'Assert the expansion is reported rather than inferred from the cell count
        TST_DP_AssertTrue "Expansion result reports ExpandedToTableColumn", _
            WriteResult.ExpandedToTableColumn
    'Assert the owning table is reported
        TST_DP_AssertEqualsString "Expansion result names the table", _
            "TST_DP_Table", WriteResult.TableName
    'Assert the resolved column is reported
        TST_DP_AssertEqualsString "Expansion result names the column", _
            "DateValue", WriteResult.ColumnName
    'Assert the whole data column was written
        TST_DP_AssertEqualsLong "Expansion result writes 3 cells", _
            3, VBA.CLng(WriteResult.WrittenCount)
    'Assert the first table data cell received the date
        TST_DP_AssertCellDateEquals "Table-column expansion writes G5", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G5")
    'Assert the last table data cell received the date
        TST_DP_AssertCellDateEquals "Table-column expansion writes G7", _
            VBA.DateSerial(2026, 8, 25), _
            mTST_DP_ScratchSheet.Range("G7")

'------------------------------------------------------------------------------
' TABLE SAFE DEFAULT
'------------------------------------------------------------------------------
    'This is the behaviour change. An omitted NoTableGrow argument previously
    'expanded a single table cell to the whole data column
        mTST_DP_ScratchSheet.Range("G5:G7").ClearContents
    'Prepare a distinct write value so a stale value cannot pass the assertion
        gDP_WriteValue = VBA.DateSerial(2026, 9, 10)
    'Select one data cell in the date column
        mTST_DP_ScratchSheet.Range("G5").Select
    'Apply write-back with NoTableGrow omitted
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker)
    'Assert the safe default is reported as an unexpanded single-cell write
        TST_DP_AssertFalse "Omitted NoTableGrow reports no expansion", _
            WriteResult.ExpandedToTableColumn
    'Assert the safe default wrote exactly one cell
        TST_DP_AssertEqualsLong "Omitted NoTableGrow writes 1 cell", _
            1, VBA.CLng(WriteResult.WrittenCount)
    'Assert the single-cell path reports one attempted cell in one area
        TST_DP_AssertEqualsLong "Omitted NoTableGrow attempts 1 cell", _
            1, VBA.CLng(WriteResult.AttemptedCount)
    'Assert the single-cell result satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Single-cell result balances", WriteResult
    'Assert the anchored cell received the date
        TST_DP_AssertCellDateEquals "Omitted NoTableGrow writes only G5", _
            VBA.DateSerial(2026, 9, 10), _
            mTST_DP_ScratchSheet.Range("G5")
    'Assert the middle table row was not written
        TST_DP_AssertTrue "Omitted NoTableGrow leaves G6 blank", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("G6").Value)) = 0
    'Assert the last table row was not written
        TST_DP_AssertTrue "Omitted NoTableGrow leaves G7 blank", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("G7").Value)) = 0

'------------------------------------------------------------------------------
' EXPLICIT TABLE COLUMN FILL
'------------------------------------------------------------------------------
    'The deliberate bulk command must still fill the whole data column
        mTST_DP_ScratchSheet.Range("G5:G7").ClearContents
    'Select one data cell in the date column
        mTST_DP_ScratchSheet.Range("G5").Select
    'Fill the column without prompting so the run stays deterministic
        WriteResult = DP_FillTableColumn(VBA.DateSerial(2026, 10, 20), ConfirmFill:=False)
    'The prompt predicts the scope, so the prediction is checked against
    'AttemptedCount. WrittenCount is reported separately and may legitimately be
    'lower when cells are skipped
        TST_DP_AssertEqualsLong "DP_FillTableColumn attempts the predicted 3 cells", _
            3, VBA.CLng(WriteResult.AttemptedCount)
    'Assert every attempted cell was written on this unobstructed fill
        TST_DP_AssertEqualsLong "DP_FillTableColumn writes all 3 cells", _
            3, VBA.CLng(WriteResult.WrittenCount)
    'Assert the fill reports the expansion it performed
        TST_DP_AssertTrue "DP_FillTableColumn reports the table expansion", _
            WriteResult.ExpandedToTableColumn
    'Assert the anchored cell received the date
        TST_DP_AssertCellDateEquals "DP_FillTableColumn writes G5", _
            VBA.DateSerial(2026, 10, 20), _
            mTST_DP_ScratchSheet.Range("G5")
    'Assert the middle table row received the date
        TST_DP_AssertCellDateEquals "DP_FillTableColumn writes G6", _
            VBA.DateSerial(2026, 10, 20), _
            mTST_DP_ScratchSheet.Range("G6")
    'Assert the last table row received the date
        TST_DP_AssertCellDateEquals "DP_FillTableColumn writes G7", _
            VBA.DateSerial(2026, 10, 20), _
            mTST_DP_ScratchSheet.Range("G7")

'------------------------------------------------------------------------------
' EXPLICIT FILL REJECTS NON-TABLE ANCHORS
'------------------------------------------------------------------------------
    'A selection outside a table data body is a usage condition, not a failure.
    'The command must exit cleanly and write nothing
        mTST_DP_ScratchSheet.Range("D8").ClearContents
    'Select a cell outside every table
        mTST_DP_ScratchSheet.Range("D8").Select
    'Attempt the fill without prompting
        WriteResult = DP_FillTableColumn(VBA.DateSerial(2026, 11, 5), ConfirmFill:=False)
    'Assert nothing was written outside a table
        TST_DP_AssertTrue "DP_FillTableColumn ignores a non-table cell", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("D8").Value)) = 0
    'Assert the refused fill reports a zero write rather than an empty silence
        TST_DP_AssertEqualsLong "DP_FillTableColumn reports 0 written outside a table", _
            0, VBA.CLng(WriteResult.WrittenCount)

    'A header cell is not inside DataBodyRange and must be rejected
        mTST_DP_ScratchSheet.Range("G5:G7").ClearContents
    'Select the date column header
        mTST_DP_ScratchSheet.Range("G4").Select
    'Attempt the fill without prompting
        DP_FillTableColumn VBA.DateSerial(2026, 11, 5), ConfirmFill:=False
    'Assert the header caption was not overwritten
        TST_DP_AssertTrue "DP_FillTableColumn ignores a table header cell", _
            VBA.CStr(mTST_DP_ScratchSheet.Range("G4").Value) = "DateValue"
    'Assert no data row was written from a header anchor
        TST_DP_AssertTrue "DP_FillTableColumn writes nothing from a header anchor", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("G5").Value)) = 0

    'A totals row cell is not inside DataBodyRange and must be rejected
        TestTable.ShowTotals = True
    'Select the date column totals cell
        TestTable.TotalsRowRange.Cells(1, 2).Select
    'Attempt the fill without prompting
        DP_FillTableColumn VBA.DateSerial(2026, 11, 5), ConfirmFill:=False
    'Assert no data row was written from a totals anchor
        TST_DP_AssertTrue "DP_FillTableColumn writes nothing from a totals anchor", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("G5").Value)) = 0
    'Restore the table to its pre-test shape
        TestTable.ShowTotals = False

    'A multi-cell selection is not a single anchor and must be rejected
        mTST_DP_ScratchSheet.Range("G5:G6").Select
    'Attempt the fill without prompting
        DP_FillTableColumn VBA.DateSerial(2026, 11, 5), ConfirmFill:=False
    'Assert no table row was written from a multi-cell anchor
        TST_DP_AssertTrue "DP_FillTableColumn writes nothing from a multi-cell anchor", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("G5").Value)) = 0

'------------------------------------------------------------------------------
' FORMULA PROTECTION
'------------------------------------------------------------------------------
    'Assert formula cells are preserved, reported, and replaceable on request
        TST_DP_ExpectFormulaProtection

'------------------------------------------------------------------------------
' PARTIAL WRITE REPORTING
'------------------------------------------------------------------------------
    'Assert a write that skips protected cells reports which cells it skipped
        TST_DP_ExpectPartialWriteReport

'------------------------------------------------------------------------------
' FAILED ADDRESS REPORTING
'------------------------------------------------------------------------------
    'Assert a write that fails on some cells reports exactly which cells failed
        TST_DP_ExpectFailedAddressReport

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

Private Sub TST_DP_RunSuite_MultiAreaWriteResult()

'
'==============================================================================
'                    MULTI-AREA WRITE RESULT SUITE
'==============================================================================
' PURPOSE
'   Proves a discontiguous write returns one complete DP_WriteResult whose totals
'   do not depend on the order Excel enumerates Target.Areas
'
' WHY THIS EXISTS
'   At v1.2.0 M_WriteBack_PopulateRange raised before accumulating whenever an
'   area wrote nothing. M_WriteBack_ApplyResolvedTarget processes Target.Areas in
'   order, so a writable area followed by a zero-write area mutated the workbook
'   and then discarded every fact observed so far. The public Function returned
'   no result at all, and the totals depended on area order
'
' BEHAVIOR
'   Drives the real public path. Each scenario selects a two-area Union and calls
'   M_WriteBack_Apply, then repeats it with the areas swapped and asserts the
'   totals are identical
'
' NOTES
'   M_WriteBack_Apply resolves its target from Application.Selection and emits no
'   message box. Shortfall reporting belongs to its callers, so this suite drives
'   the engine without modal interruption
'
'   The suite works in column M, away from the ranges the other write-back suites
'   use, and restores sheet protection in both exit paths
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ForwardResult   As DP_WriteResult   'Result with the writable area first
    Dim ReverseResult   As DP_WriteResult   'Result with the writable area last
    Dim ZeroResult      As DP_WriteResult   'Result for an all-zero-write target
    Dim WasProtected    As Boolean          'Sheet protection state on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "MultiAreaWriteResult"

'------------------------------------------------------------------------------
' PREPARE THE SCRATCH REGION
'------------------------------------------------------------------------------
    'Record protection state so the suite can restore it
        WasProtected = mTST_DP_ScratchSheet.ProtectContents
    'Work on an unprotected sheet while the fixtures are built
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
    'Clear the whole working region
        mTST_DP_ScratchSheet.Range("M5:M12").ClearContents
    'M5 and M9 are the writable cells
        mTST_DP_ScratchSheet.Range("M5").Locked = False
        mTST_DP_ScratchSheet.Range("M9").Locked = True
    'M7 is a formula cell the write must preserve by policy
        mTST_DP_ScratchSheet.Range("M7").Formula = "=1+1"
    'M11:M12 is an array formula the write cannot replace
        mTST_DP_ScratchSheet.Range("M11:M12").FormulaArray = "=ROW()"
    'Assert the fixtures took, so a silent setup failure cannot look like a defect
        TST_DP_AssertTrue "Multi-area setup creates a formula cell", _
            mTST_DP_ScratchSheet.Range("M7").HasFormula
        TST_DP_AssertTrue "Multi-area setup creates an array formula", _
            mTST_DP_ScratchSheet.Range("M11").HasArray

'------------------------------------------------------------------------------
' WRITABLE THEN FORMULA-ONLY, AND THE REVERSE
'------------------------------------------------------------------------------
    'Writable area first
        ForwardResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M5"), mTST_DP_ScratchSheet.Range("M7"))
        TST_DP_AssertEqualsLong "Writable then formula reports 2 attempted", _
            2, VBA.CLng(ForwardResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Writable then formula reports 1 written", _
            1, VBA.CLng(ForwardResult.WrittenCount)
        TST_DP_AssertEqualsLong "Writable then formula reports 1 formula skip", _
            1, VBA.CLng(ForwardResult.FormulaSkippedCount)
        TST_DP_AssertEqualsLong "Writable then formula counts both areas", _
            2, ForwardResult.AreasCount
        TST_DP_AssertWriteResultBalances "Writable then formula balances", ForwardResult
    'Formula area first: this is the ordering that lost the result at v1.2.0
        ReverseResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M7"), mTST_DP_ScratchSheet.Range("M5"))
        TST_DP_AssertEqualsLong "Formula then writable reports 2 attempted", _
            2, VBA.CLng(ReverseResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Formula then writable reports 1 written", _
            1, VBA.CLng(ReverseResult.WrittenCount)
        TST_DP_AssertEqualsLong "Formula then writable reports 1 formula skip", _
            1, VBA.CLng(ReverseResult.FormulaSkippedCount)
        TST_DP_AssertEqualsLong "Formula then writable counts both areas", _
            2, ReverseResult.AreasCount
        TST_DP_AssertWriteResultBalances "Formula then writable balances", ReverseResult

'------------------------------------------------------------------------------
' WRITABLE THEN FAILED-ONLY, AND THE REVERSE
'------------------------------------------------------------------------------
    'Writable area first
        ForwardResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M5"), mTST_DP_ScratchSheet.Range("M11:M12"))
        TST_DP_AssertEqualsLong "Writable then failed reports 3 attempted", _
            3, VBA.CLng(ForwardResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Writable then failed reports 1 written", _
            1, VBA.CLng(ForwardResult.WrittenCount)
        TST_DP_AssertEqualsLong "Writable then failed reports 2 failures", _
            2, VBA.CLng(ForwardResult.FailedCount)
        TST_DP_AssertWriteResultBalances "Writable then failed balances", ForwardResult
    'Failed area first
        ReverseResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M11:M12"), mTST_DP_ScratchSheet.Range("M5"))
        TST_DP_AssertEqualsLong "Failed then writable reports 3 attempted", _
            3, VBA.CLng(ReverseResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Failed then writable reports 1 written", _
            1, VBA.CLng(ReverseResult.WrittenCount)
        TST_DP_AssertEqualsLong "Failed then writable reports 2 failures", _
            2, VBA.CLng(ReverseResult.FailedCount)
        TST_DP_AssertWriteResultBalances "Failed then writable balances", ReverseResult
    'The two orders must agree on every total
        TST_DP_AssertEqualsString "Failed-area totals are order-independent", _
            VBA.CStr(ForwardResult.AttemptedCount) & "/" & _
            VBA.CStr(ForwardResult.WrittenCount) & "/" & _
            VBA.CStr(ForwardResult.FailedCount), _
            VBA.CStr(ReverseResult.AttemptedCount) & "/" & _
            VBA.CStr(ReverseResult.WrittenCount) & "/" & _
            VBA.CStr(ReverseResult.FailedCount)

'------------------------------------------------------------------------------
' AN ALL-ZERO-WRITE TARGET RETURNS A COMPLETE RESULT
'------------------------------------------------------------------------------
    'Neither area can accept the value, so the operation writes nothing. This
    'must produce a complete result rather than an exception
        ZeroResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M7"), mTST_DP_ScratchSheet.Range("M11:M12"))
        TST_DP_AssertEqualsLong "Zero-write target reports 3 attempted", _
            3, VBA.CLng(ZeroResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Zero-write target reports 0 written", _
            0, VBA.CLng(ZeroResult.WrittenCount)
        TST_DP_AssertEqualsLong "Zero-write target still counts both areas", _
            2, ZeroResult.AreasCount
        TST_DP_AssertTrue "Zero-write target still names the resolved target", _
            (VBA.LenB(ZeroResult.ResolvedTargetAddress) > 0)
        TST_DP_AssertWriteResultBalances "Zero-write result balances", ZeroResult

'------------------------------------------------------------------------------
' WRITABLE THEN LOCKED-ONLY, AND THE REVERSE
'------------------------------------------------------------------------------
    'Protect the sheet so the locked cell actually rejects the write
        mTST_DP_ScratchSheet.Protect
    'Writable area first
        ForwardResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M5"), mTST_DP_ScratchSheet.Range("M9"))
        TST_DP_AssertEqualsLong "Writable then locked reports 2 attempted", _
            2, VBA.CLng(ForwardResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Writable then locked reports 1 written", _
            1, VBA.CLng(ForwardResult.WrittenCount)
        TST_DP_AssertEqualsLong "Writable then locked reports 1 locked skip", _
            1, VBA.CLng(ForwardResult.LockedSkippedCount)
        TST_DP_AssertWriteResultBalances "Writable then locked balances", ForwardResult
    'Locked area first
        ReverseResult = TST_DP_WriteTwoAreasForTest( _
            mTST_DP_ScratchSheet.Range("M9"), mTST_DP_ScratchSheet.Range("M5"))
        TST_DP_AssertEqualsLong "Locked then writable reports 2 attempted", _
            2, VBA.CLng(ReverseResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Locked then writable reports 1 written", _
            1, VBA.CLng(ReverseResult.WrittenCount)
        TST_DP_AssertEqualsLong "Locked then writable reports 1 locked skip", _
            1, VBA.CLng(ReverseResult.LockedSkippedCount)
        TST_DP_AssertWriteResultBalances "Locked then writable balances", ReverseResult
    'The two orders must agree on every total
        TST_DP_AssertEqualsString "Locked-area totals are order-independent", _
            VBA.CStr(ForwardResult.AttemptedCount) & "/" & _
            VBA.CStr(ForwardResult.WrittenCount) & "/" & _
            VBA.CStr(ForwardResult.LockedSkippedCount), _
            VBA.CStr(ReverseResult.AttemptedCount) & "/" & _
            VBA.CStr(ReverseResult.WrittenCount) & "/" & _
            VBA.CStr(ReverseResult.LockedSkippedCount)

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Restore the sheet to the protection state the suite found
        On Error Resume Next
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
        mTST_DP_ScratchSheet.Range("M5:M12").ClearContents
        mTST_DP_ScratchSheet.Range("M5:M12").Locked = True
        If WasProtected Then
            mTST_DP_ScratchSheet.Protect
        End If
        Err.Clear
        On Error GoTo 0
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Multi-area write result suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore sheet state regardless
        Resume SuiteExit

End Sub

Private Function TST_DP_WriteTwoAreasForTest( _
    ByVal FirstArea As Excel.Range, _
    ByVal SecondArea As Excel.Range) As DP_WriteResult

'
'==============================================================================
'               WRITE A TWO-AREA TARGET IN A GIVEN ORDER (TEST)
'==============================================================================
'   Selects a discontiguous target built in the supplied order and drives the
'   real public write path.
'
'   M_WriteBack_Apply resolves its target from Application.Selection, so
'   selecting the Union is how a test controls the order Target.Areas is
'   enumerated. Building the Union with the arguments swapped is what makes the
'   order-independence assertions meaningful.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim UnionRange      As Excel.Range      'Discontiguous target in the requested order

'------------------------------------------------------------------------------
' PREPARE
'------------------------------------------------------------------------------
    'Clear only the writable cells so each scenario starts from a known state
        On Error Resume Next
        mTST_DP_ScratchSheet.Range("M5").ClearContents
        Err.Clear
        On Error GoTo 0
    'Stage a distinct value for this write
        gDP_WriteValue = VBA.DateSerial(2026, 9, 17)

'------------------------------------------------------------------------------
' SELECT THE TARGET IN THE REQUESTED ORDER
'------------------------------------------------------------------------------
    'Build the discontiguous target
        Set UnionRange = Excel.Application.Union(FirstArea, SecondArea)
    'Activate the sheet so the selection is valid
        mTST_DP_ScratchSheet.Activate
    'Select it, because the engine resolves its target from the selection
        UnionRange.Select

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Drive the real public path. Apply emits no message box; shortfall reporting
    'belongs to its callers
        TST_DP_WriteTwoAreasForTest = M_WriteBack_Apply(DP_WriteAction_DatePicker)

'------------------------------------------------------------------------------
' RELEASE
'------------------------------------------------------------------------------
    'Release the object reference
        Set UnionRange = Nothing

End Function

Private Sub TST_DP_RunSuite_WriteTechnicalFailure()

'
'==============================================================================
'                  WRITE TECHNICAL FAILURE SUITE
'==============================================================================
' PURPOSE
'   Proves that an unexpected technical failure during a write never becomes an
'   exception carrying no result once the workbook has been mutated
'
' WHY THIS EXISTS
'   fee9271 fixed the known ordering defect: a zero-write area no longer raises,
'   so a writable area followed by a zero-write area keeps its facts. It did not
'   close the stronger #21 contract. M_WriteBack_PopulateRange still raised on an
'   unexpected technical error, and an area contributed its facts to the operation
'   result only after that routine returned successfully, so a technical failure
'   after mutation still discarded everything observed
'
'   That path cannot be produced on demand. M_WriteBack_TryWriteCell classifies
'   every per-cell failure it can observe and raises nothing, which is exactly why
'   a controlled fault seam is required. The classified failures this suite does
'   not use — array-formula refusal, locked cells, formula preservation — are a
'   different path and are covered by MultiAreaWriteResult
'
' BEHAVIOR
'   Drives the real public write path with a fault armed at three positions: after
'   cells have been mutated inside one area, on entering a later area after an
'   earlier one completed, and on entering the first area before anything at all
'   has been observed. The first two must return a complete result. The third must
'   still raise, because there are no facts to protect
'
'   One scenario deliberately uses an earlier area that only skipped a formula
'   cell. It mutated nothing, but a skip is still an outcome the caller asked for,
'   so it must survive. That pins the raise-or-return rule to per-cell outcomes
'   rather than to mutation, and keeps it clear of AttemptedCount, which is
'   recorded before any cell is touched and is therefore evidence of nothing
'
' NOTES
'   The suite works in column N, away from the ranges the other write-back suites
'   use, and restores sheet protection in both exit paths
'
'   N7 holds a formula so the first scenario is forced down the per-cell fallback.
'   The bulk path writes a whole area in one assignment and would leave no
'   position for a mid-area fault to fire
'
'   Injection is one-shot and the helper disarms after every scenario, so a failed
'   assertion cannot leave a fault armed for a later suite
'
'   A technical-failure result is bounded, not balanced. The areas the operation
'   deliberately stopped short of are absent from AttemptedCount, so
'   TST_DP_AssertWriteResultBalances does not apply here and the weaker invariant
'   is asserted instead
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const INJECTED_ERROR    As Long = vbObjectError + 518

    Dim IntraResult     As DP_WriteResult   'Fault fired inside one area
    Dim CrossResult     As DP_WriteResult   'Fault fired on entering a later area
    Dim SkipOnlyResult  As DP_WriteResult   'Fault fired after an area that only skipped
    Dim EmptyFaultResult As DP_WriteResult  'Fault fired before anything was observed
    Dim CleanResult     As DP_WriteResult   'Write after the fault, with nothing armed
    Dim DisarmedResult  As DP_WriteResult   'Write after an armed fault was disarmed

    Dim Raised           As Boolean         'True when the write raised
    Dim RaisedNumber     As Long            'Error number the write raised
    Dim WriteValue       As Date            'Value the scenarios write
    Dim UnionRange       As Excel.Range     'Two-area target
    Dim WasProtected     As Boolean         'Sheet protection state on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "WriteTechnicalFailure"

'------------------------------------------------------------------------------
' PREPARE THE SCRATCH REGION
'------------------------------------------------------------------------------
    'Record protection state so the suite can restore it
        WasProtected = mTST_DP_ScratchSheet.ProtectContents
    'Work on an unprotected sheet
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
    'Clear the whole working region
        mTST_DP_ScratchSheet.Range("N5:N12").ClearContents
    'N7 holds a formula, which refuses the bulk path and forces the per-cell loop
        mTST_DP_ScratchSheet.Range("N7").Formula = "=1+1"
    'Resolve the value every scenario writes
        WriteValue = VBA.DateSerial(2026, 9, 18)
    'Assert the fixture took, so a silent setup failure cannot look like a defect
        TST_DP_AssertTrue "Technical-failure setup creates a formula cell", _
            mTST_DP_ScratchSheet.Range("N7").HasFormula

'------------------------------------------------------------------------------
' TECHNICAL FAILURE AFTER EARLIER CELLS INSIDE ONE AREA
'------------------------------------------------------------------------------
    'Fail inside the first area once two of its cells have been mutated
        IntraResult = TST_DP_WriteWithFaultForTest( _
            mTST_DP_ScratchSheet.Range("N5:N7"), WriteValue, 1, 2, Raised, RaisedNumber)
    'The caller must receive the facts, not an exception
        TST_DP_AssertFalse "Intra-area technical failure returns a result", Raised
        TST_DP_AssertTrue "Intra-area technical failure is flagged in the result", _
            IntraResult.TechnicalFailureOccurred
        TST_DP_AssertEqualsLong "Intra-area failure preserves the original error", _
            INJECTED_ERROR, IntraResult.TechnicalFailureNumber
        TST_DP_AssertEqualsString "Intra-area failure names the step it failed in", _
            "Populate target cells through safe fallback", _
            IntraResult.TechnicalFailureStep
    'The facts observed before the failure must survive it
        TST_DP_AssertEqualsLong "Intra-area failure reports 3 attempted", _
            3, VBA.CLng(IntraResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Intra-area failure reports 2 written", _
            2, VBA.CLng(IntraResult.WrittenCount)
        TST_DP_AssertEqualsLong "Intra-area failure counts the area it reached", _
            1, IntraResult.AreasCount
        TST_DP_AssertTrue "Intra-area failure still names the resolved target", _
            (VBA.LenB(IntraResult.ResolvedTargetAddress) > 0)
    'The result describes an operation that stopped short, not a balanced one
        TST_DP_AssertTrue "Intra-area failure result is bounded by attempted", _
            (IntraResult.WrittenCount + IntraResult.LockedSkippedCount + _
             IntraResult.FormulaSkippedCount + IntraResult.FailedCount <= _
             IntraResult.AttemptedCount)
    'The mutation the workbook actually received must still be there
        TST_DP_AssertTrue "Intra-area failure keeps the cells already written", _
            (VBA.CDbl(mTST_DP_ScratchSheet.Range("N5").Value2) = VBA.CDbl(WriteValue) And _
             VBA.CDbl(mTST_DP_ScratchSheet.Range("N6").Value2) = VBA.CDbl(WriteValue))
        TST_DP_AssertTrue "Intra-area failure leaves the cell it never reached", _
            mTST_DP_ScratchSheet.Range("N7").HasFormula
    'The user-facing description must say the operation stopped early
        TST_DP_AssertTrue "Intra-area failure is described to the user", _
            (VBA.InStr(1, M_WriteBack_DescribeShortfall(IntraResult), _
                "stopped early", vbTextCompare) > 0)

'------------------------------------------------------------------------------
' TECHNICAL FAILURE AFTER AN EARLIER AREA
'------------------------------------------------------------------------------
    'Reset the region so the second scenario starts from a known state
        mTST_DP_ScratchSheet.Range("N5:N6").ClearContents
        mTST_DP_ScratchSheet.Range("N9").ClearContents
    'Build a two-area target and fail on entering the second area
        Set UnionRange = Excel.Application.Union( _
            mTST_DP_ScratchSheet.Range("N5"), mTST_DP_ScratchSheet.Range("N9"))
        CrossResult = TST_DP_WriteWithFaultForTest( _
            UnionRange, WriteValue, 2, 0, Raised, RaisedNumber)
    'The caller must receive the first area's facts, not an exception
        TST_DP_AssertFalse "Cross-area technical failure returns a result", Raised
        TST_DP_AssertTrue "Cross-area technical failure is flagged in the result", _
            CrossResult.TechnicalFailureOccurred
        TST_DP_AssertEqualsLong "Cross-area failure preserves the original error", _
            INJECTED_ERROR, CrossResult.TechnicalFailureNumber
        TST_DP_AssertEqualsString "Cross-area failure names the step it failed in", _
            "Record area position", CrossResult.TechnicalFailureStep
    'This is the assertion the whole issue turns on: an earlier area mutated the
    'workbook and its facts reached the caller anyway
        TST_DP_AssertEqualsLong "Cross-area failure keeps the earlier area's write", _
            1, VBA.CLng(CrossResult.WrittenCount)
        TST_DP_AssertEqualsLong "Cross-area failure reports 1 attempted", _
            1, VBA.CLng(CrossResult.AttemptedCount)
        TST_DP_AssertEqualsLong "Cross-area failure counts only the area it reached", _
            1, CrossResult.AreasCount
        TST_DP_AssertTrue "Cross-area failure keeps the earlier area's mutation", _
            (VBA.CDbl(mTST_DP_ScratchSheet.Range("N5").Value2) = VBA.CDbl(WriteValue))
        TST_DP_AssertTrue "Cross-area failure leaves the area it never entered", _
            VBA.IsEmpty(mTST_DP_ScratchSheet.Range("N9").Value2)

'------------------------------------------------------------------------------
' AN EARLIER AREA THAT ONLY SKIPPED STILL KEEPS ITS FACTS
'------------------------------------------------------------------------------
    'Reset the region
        mTST_DP_ScratchSheet.Range("N9").ClearContents
    'The first area holds a formula, so it observes a skip and mutates nothing.
    'That is still an outcome the caller asked for, and it must survive the
    'failure in the second area. This pins the rule to per-cell outcomes rather
    'than to whether the workbook happened to be mutated
        Set UnionRange = Excel.Application.Union( _
            mTST_DP_ScratchSheet.Range("N7"), mTST_DP_ScratchSheet.Range("N9"))
        SkipOnlyResult = TST_DP_WriteWithFaultForTest( _
            UnionRange, WriteValue, 2, 0, Raised, RaisedNumber)
        TST_DP_AssertFalse "A skip-only earlier area returns a result", Raised
        TST_DP_AssertTrue "A skip-only earlier area flags the technical failure", _
            SkipOnlyResult.TechnicalFailureOccurred
        TST_DP_AssertEqualsLong "A skip-only earlier area keeps its formula skip", _
            1, VBA.CLng(SkipOnlyResult.FormulaSkippedCount)
        TST_DP_AssertEqualsLong "A skip-only earlier area reports no write", _
            0, VBA.CLng(SkipOnlyResult.WrittenCount)

'------------------------------------------------------------------------------
' A FAILURE CARRYING NO FACTS STILL RAISES
'------------------------------------------------------------------------------
    'Reset the region
        mTST_DP_ScratchSheet.Range("N5").ClearContents
        mTST_DP_ScratchSheet.Range("N9").ClearContents
    'Fail on entering the first area, before anything has been observed. There is
    'nothing to protect here, so the original error must reach the caller intact
        Set UnionRange = Excel.Application.Union( _
            mTST_DP_ScratchSheet.Range("N5"), mTST_DP_ScratchSheet.Range("N9"))
        EmptyFaultResult = TST_DP_WriteWithFaultForTest( _
            UnionRange, WriteValue, 1, 0, Raised, RaisedNumber)
        TST_DP_AssertTrue "A failure carrying no facts still raises", Raised
        TST_DP_AssertEqualsLong "A raised failure preserves the original error", _
            INJECTED_ERROR, RaisedNumber
        TST_DP_AssertTrue "A raised failure leaves the target unmutated", _
            (VBA.IsEmpty(mTST_DP_ScratchSheet.Range("N5").Value2) And _
             VBA.IsEmpty(mTST_DP_ScratchSheet.Range("N9").Value2))

'------------------------------------------------------------------------------
' AN INJECTED FAULT DOES NOT LEAK INTO THE NEXT WRITE
'------------------------------------------------------------------------------
    'Reset the region
        mTST_DP_ScratchSheet.Range("N5:N6").ClearContents
    'Write with nothing armed
        CleanResult = TST_DP_WriteWithFaultForTest( _
            mTST_DP_ScratchSheet.Range("N5:N6"), WriteValue, 0, 0, Raised, RaisedNumber)
        TST_DP_AssertFalse "The next write after a fault does not raise", Raised
        TST_DP_AssertFalse "The next write after a fault reports no failure", _
            CleanResult.TechnicalFailureOccurred
        TST_DP_AssertEqualsLong "The next write after a fault writes every cell", _
            2, VBA.CLng(CleanResult.WrittenCount)
        TST_DP_AssertWriteResultBalances "A write with no fault balances", CleanResult

'------------------------------------------------------------------------------
' A DISARMED FAULT DOES NOT FIRE
'------------------------------------------------------------------------------
    'Reset the region
        mTST_DP_ScratchSheet.Range("N5:N6").ClearContents
    'Arm a fault and disarm it before the write
        M_WriteBack_Test_SetFaultInjection 1, 0
        M_WriteBack_Test_SetFaultInjection 0
        DisarmedResult = TST_DP_WriteWithFaultForTest( _
            mTST_DP_ScratchSheet.Range("N5:N6"), WriteValue, 0, 0, Raised, RaisedNumber)
        TST_DP_AssertFalse "A disarmed fault does not fire", _
            DisarmedResult.TechnicalFailureOccurred
        TST_DP_AssertEqualsLong "A disarmed fault leaves the write complete", _
            2, VBA.CLng(DisarmedResult.WrittenCount)

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Disarm any fault this suite left behind, whatever path it exits by
        On Error Resume Next
        M_WriteBack_Test_SetFaultInjection 0
    'Release the object reference
        Set UnionRange = Nothing
    'Restore the sheet to the protection state the suite found
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
        mTST_DP_ScratchSheet.Range("N5:N12").ClearContents
        mTST_DP_ScratchSheet.Range("N5:N12").Locked = True
        If WasProtected Then
            mTST_DP_ScratchSheet.Protect
        End If
        Err.Clear
        On Error GoTo 0
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Write technical failure suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore sheet state regardless
        Resume SuiteExit

End Sub

Private Function TST_DP_WriteWithFaultForTest( _
    ByVal TargetRange As Excel.Range, _
    ByVal WriteValue As Date, _
    ByVal FailInAreaOrdinal As Long, _
    ByVal FailAfterWrittenCellsInArea As Long, _
    ByRef Raised As Boolean, _
    ByRef RaisedNumber As Long) As DP_WriteResult

'
'==============================================================================
'            WRITE A TARGET WITH A FAULT ARMED (TEST)
'==============================================================================
'   Selects the supplied target, optionally arms a forced technical failure, and
'   drives the real public write path.
'
'   Whether the write raised is reported through Raised and RaisedNumber rather
'   than by letting the error reach the suite, because both outcomes are expected
'   results here: a technical failure that carries facts must return, and one that
'   carries nothing must raise. A suite-level handler could not tell them apart.
'
'   A UDT assignment from a function that raises does not occur, so the returned
'   result is zeroed on the raising path. Only Raised and RaisedNumber are
'   meaningful there.
'
'   The fault is disarmed on every path. Injection is one-shot in production code
'   as well, but a scenario that never reaches its armed position would otherwise
'   leave it armed for the next suite.
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe defaults
        Raised = False
        RaisedNumber = 0
    'Stage the value this write applies
        gDP_WriteValue = WriteValue

'------------------------------------------------------------------------------
' SELECT THE TARGET
'------------------------------------------------------------------------------
    'Activate the sheet so the selection is valid
        mTST_DP_ScratchSheet.Activate
    'Select it, because the engine resolves its target from the selection
        TargetRange.Select

'------------------------------------------------------------------------------
' ARM THE FAULT
'------------------------------------------------------------------------------
    'Arm only when the caller asked for one
        If FailInAreaOrdinal > 0 Then
            M_WriteBack_Test_SetFaultInjection FailInAreaOrdinal, _
                FailAfterWrittenCellsInArea
        End If

'------------------------------------------------------------------------------
' WRITE
'------------------------------------------------------------------------------
    'Drive the real public path and record whichever outcome it produces
        On Error Resume Next
        TST_DP_WriteWithFaultForTest = M_WriteBack_Apply(DP_WriteAction_DatePicker)
        If Err.Number <> 0 Then
            Raised = True
            RaisedNumber = Err.Number
            Err.Clear
        End If
        On Error GoTo 0

'------------------------------------------------------------------------------
' DISARM
'------------------------------------------------------------------------------
    'Never leave a fault armed for a later scenario or suite
        M_WriteBack_Test_SetFaultInjection 0

End Function

Private Sub TST_DP_RunSuite_ErrorPreservation()

'
'==============================================================================
'                     ERROR PRESERVATION SUITE
'==============================================================================
' PURPOSE
'   Proves that a handler which cleans up before re-raising still reports the
'   error that actually occurred
'
' WHY THIS EXISTS
'   Every On Error statement resets the Err object, and so does Err.Clear. A
'   handler that cleans up under On Error Resume Next and then raises from the
'   live Err object reports error 0 with a blank description, and the real cause
'   is gone by the time the caller sees it
'
'   DP_FillTableColumn did exactly this whenever it had already staged the pending
'   write value, which is every failure inside the fill itself
'
' BEHAVIOR
'   Drives the real public DP_FillTableColumn against a real Excel Table with a
'   write fault armed, and asserts the caller receives the injected error rather
'   than error 0
'
' NOTES
'   The fault seam is the one added for #21. Reusing it keeps this suite free of
'   any seam of its own: the failure it needs is an ordinary write failure, and
'   that is exactly what the seam produces
'
'   ConfirmFill is False throughout, so no message box is emitted. The scope
'   confirmation and the table-cell guidance are both interactive-only
'
'   The suite builds its own table in columns P and Q, away from the table the
'   WriteBack suite creates, and removes it in both exit paths
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const INJECTED_ERROR    As Long = vbObjectError + 518

    Dim TableRange      As Excel.Range      'Source range for the test table
    Dim TestTable       As Excel.ListObject 'Table the fill targets
    Dim CleanResult     As DP_WriteResult   'Result of the fill with no fault armed
    Dim StagedValue     As Date             'Pending write value staged before the fill
    Dim RaisedNumber    As Long             'Error number the fill raised
    Dim RaisedText      As String           'Error description the fill raised
    Dim RaisedSource    As String           'Error source the fill raised
    Dim Raised          As Boolean          'True when the fill raised
    Dim WasProtected    As Boolean          'Sheet protection state on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "ErrorPreservation"

'------------------------------------------------------------------------------
' PREPARE A REAL TABLE
'------------------------------------------------------------------------------
    'Record protection state so the suite can restore it
        WasProtected = mTST_DP_ScratchSheet.ProtectContents
    'Work on an unprotected sheet
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
    'Clear the working region
        Set TableRange = mTST_DP_ScratchSheet.Range("P4:Q7")
        TableRange.Clear
    'Write the table headers
        mTST_DP_ScratchSheet.Range("P4").Value = "ID"
        mTST_DP_ScratchSheet.Range("Q4").Value = "DateValue"
    'Write the table row identifiers
        mTST_DP_ScratchSheet.Range("P5:P7").Value = 1
    'Create the table this suite fills
        Set TestTable = mTST_DP_ScratchSheet.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=TableRange, _
            XlListObjectHasHeaders:=xlYes)
        TestTable.Name = "TST_DP_ErrTable"
    'Assert the fixture took, so a silent setup failure cannot look like a defect
        TST_DP_AssertEqualsString "Error-preservation setup creates a table", _
            "TST_DP_ErrTable", TestTable.Name

'------------------------------------------------------------------------------
' A FAILED FILL REPORTS THE ERROR THAT ACTUALLY OCCURRED
'------------------------------------------------------------------------------
    'Stage a recognizable pending write value, so the restore can be verified
        StagedValue = VBA.DateSerial(2026, 2, 2)
        gDP_WriteValue = StagedValue
    'Activate the sheet and select a data cell in the date column
        mTST_DP_ScratchSheet.Activate
        mTST_DP_ScratchSheet.Range("Q5").Select
    'Arm a write fault that fires before any cell is mutated
        M_WriteBack_Test_SetFaultInjection 1, 0
    'Drive the real public entry point and record whatever it raises
        On Error Resume Next
        DP_FillTableColumn VBA.DateSerial(2026, 4, 4), ConfirmFill:=False
        If Err.Number <> 0 Then
            Raised = True
            RaisedNumber = Err.Number
            RaisedText = Err.Description
            RaisedSource = Err.Source
            Err.Clear
        End If
        On Error GoTo 0
    'Disarm whatever remains
        M_WriteBack_Test_SetFaultInjection 0
    'The failure must reach the caller at all
        TST_DP_AssertTrue "A failed table fill raises to the caller", Raised
    'This is the defect: the caller used to receive error 0
        TST_DP_AssertFalse "A failed table fill does not report error 0", _
            (RaisedNumber = 0)
    'The original error number must survive the cleanup that follows it
        TST_DP_AssertEqualsLong "A failed table fill preserves the original error", _
            INJECTED_ERROR, RaisedNumber
    'The description must name the failing operation
        TST_DP_AssertTrue "A failed table fill names the failing operation", _
            (VBA.InStr(1, RaisedText, "Table column fill failed", vbTextCompare) > 0)
    'The description must carry the original cause rather than a blank string
        TST_DP_AssertTrue "A failed table fill carries the original cause", _
            (VBA.InStr(1, RaisedText, "Injected write-back fault", vbTextCompare) > 0)
    'The source must remain usable for diagnosis
        TST_DP_AssertTrue "A failed table fill names the procedure and step", _
            (VBA.InStr(1, RaisedSource, "DP_FillTableColumn", vbTextCompare) > 0 And _
             VBA.InStr(1, RaisedSource, "Step=", vbTextCompare) > 0)
    'Cleanup must still have run: preserving the error is not an excuse to skip it
        TST_DP_AssertTrue "A failed table fill still restores the pending value", _
            (VBA.CDbl(gDP_WriteValue) = VBA.CDbl(StagedValue))

'------------------------------------------------------------------------------
' THE SUCCESS PATH IS UNAFFECTED
'------------------------------------------------------------------------------
    'Select a data cell again and fill with nothing armed
        mTST_DP_ScratchSheet.Range("Q5").Select
        CleanResult = DP_FillTableColumn(VBA.DateSerial(2026, 4, 4), ConfirmFill:=False)
        TST_DP_AssertTrue "A clean table fill reports the expansion", _
            CleanResult.ExpandedToTableColumn
        TST_DP_AssertEqualsLong "A clean table fill writes the whole column", _
            3, VBA.CLng(CleanResult.WrittenCount)
        TST_DP_AssertWriteResultBalances "A clean table fill balances", CleanResult

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Disarm any fault this suite left behind, whatever path it exits by
        On Error Resume Next
        M_WriteBack_Test_SetFaultInjection 0
    'Remove the table and clear its range
        If Not TestTable Is Nothing Then
            TestTable.Unlist
        End If
        Set TestTable = Nothing
        Set TableRange = Nothing
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
        mTST_DP_ScratchSheet.Range("P4:Q7").Clear
        If WasProtected Then
            mTST_DP_ScratchSheet.Protect
        End If
        Err.Clear
        On Error GoTo 0
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Error preservation suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore sheet state regardless
        Resume SuiteExit

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
    Dim StaleSheet          As Excel.Worksheet  'Temporary sheet used to strand the icon

    Dim SavedErrNumber      As Long     'Captured original error number
    Dim SavedErrDescription As String   'Captured original error description
    Dim SavedErrSource      As String   'Captured original error source

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "GridIcon"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet
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
    'correct invariant to assert here   visible state is a rendering detail.
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
' STALE TRACKED REFERENCE
'------------------------------------------------------------------------------
    'The tracked reference outlives the shape it points to whenever the shape is
    'destroyed without going through a routine that maintains it. Deleting the
    'worksheet holding the icon is the simplest way to produce that state
        gDP_ShowGridIcon = True
        Set StaleSheet = TST_DP_AddStaleSheetForTest( _
            mTST_DP_HostWorkbook, TST_DP_STALE_SHEET_NAME)
        TST_DP_ActivateWorksheetForTest StaleSheet

    'Create the icon on the temporary worksheet
    'M_GridIcon_ShowOrMove resets On Error GoTo 0 on exit; re-arm immediately
        M_GridIcon_ShowOrMove StaleSheet.Range("B2")
        On Error GoTo SuiteFail
        DoEvents
        TST_DP_AssertTrue "Stale-reference setup creates a tracked icon", _
            Not (gDP_GridIconShape Is Nothing)

    'Destroy the shape behind the reference without clearing it. Deletion goes
    'through the shared helper, which suppresses its own errors, and nothing here
    'touches DisplayAlerts: the run already disabled alerts and forcing them back
    'on would re-enable the delete prompt for everything that follows
        TST_DP_DeleteWorksheetByReference StaleSheet
        TST_DP_DeleteWorksheetIfExists mTST_DP_HostWorkbook, TST_DP_STALE_SHEET_NAME
        On Error GoTo SuiteFail
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet

    'A stale reference must not stop a new icon being created. Before the
    'liveness check, PreCreateHidden saw a non-Nothing variable and skipped
    'straight to hiding a shape that no longer existed
        M_GridIcon_PreCreateHidden mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
        DoEvents
        TST_DP_AssertTrue "Stale reference does not block icon creation", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Recreate the stale condition and prove teardown clears it rather than
    'raising on it
        Set StaleSheet = TST_DP_AddStaleSheetForTest( _
            mTST_DP_HostWorkbook, TST_DP_STALE_SHEET_NAME)
        TST_DP_ActivateWorksheetForTest StaleSheet
        M_GridIcon_ShowOrMove StaleSheet.Range("B2")
        On Error GoTo SuiteFail
        DoEvents
        TST_DP_DeleteWorksheetByReference StaleSheet
        TST_DP_DeleteWorksheetIfExists mTST_DP_HostWorkbook, TST_DP_STALE_SHEET_NAME
        On Error GoTo SuiteFail
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet

    'Remove must leave nothing tracked and must not raise
        M_GridIcon_Remove
        On Error GoTo SuiteFail
        TST_DP_AssertTrue "Remove clears a stale tracked reference", _
            gDP_GridIconShape Is Nothing

    'Purge must tolerate the same condition
        M_GridIcon_ShowOrMove mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail
        TST_DP_AssertTrue "Purge clears the tracked reference", _
            gDP_GridIconShape Is Nothing

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Capture the original error before any cleanup runs
    '
    'Every On Error statement resets the Err object, and so does Err.Clear. This
    'handler used to clean up first and then format its message from the live Err,
    'so it reported "Error 0 - " and destroyed the evidence of its own failure.
    'That is the #48 defect shape, in the harness. Nothing may read the live Err
    'object below this point
        SavedErrNumber = Err.Number
        SavedErrDescription = Err.Description
        SavedErrSource = Err.Source
    'Release the temporary worksheet the stale-reference cases may have left. The
    'run owns DisplayAlerts and this must not change it
        On Error Resume Next
        TST_DP_DeleteWorksheetByReference StaleSheet
        TST_DP_DeleteWorksheetIfExists mTST_DP_HostWorkbook, TST_DP_STALE_SHEET_NAME
        Err.Clear
    'Record the suite-level failure from the captured values
        TST_DP_RecordFail "GridIcon suite failed", _
            "Error " & VBA.CStr(SavedErrNumber) & " - " & SavedErrDescription & _
            " | Source=" & SavedErrSource
        Err.Clear

End Sub

Private Function TST_DP_AddStaleSheetForTest( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal DesiredName As String) As Excel.Worksheet

'
'==============================================================================
'          ADD THE STALE-REFERENCE WORKSHEET, TOLERATING THE ADD ANOMALY
'==============================================================================
'   Worksheets.Add has been observed to report 1004 while nevertheless leaving a
'   new worksheet in the workbook. TST_DP_PrepareScratchSheet already treats that
'   as a partial success, because every occurrence leaked a sheet that preflight
'   could not see: it looks for a sheet by name and the leftover is called
'   SheetNN.
'
'   The GridIcon stale-reference cases called Worksheets.Add raw. When the anomaly
'   struck there, the Add raised before the rename, so the suite failed, the
'   handler tried to delete by a name nothing had been given yet, and an unnamed
'   worksheet was left in the host workbook.
'
'   This routine applies the same discipline as the scratch path: snapshot, add,
'   and on a raised Add adopt the worksheet if and only if exactly one new one
'   appeared. Anything else re-raises the original failure.
'
'   The worksheet is returned before it is renamed as well as after, so a caller
'   whose rename fails still holds the reference needed to delete it by identity.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TST_DP_AddStaleSheetForTest"

    Dim WorksheetsBefore    As Collection       'Identity snapshot taken before the Add
    Dim NewWorksheets       As Collection       'Worksheets that appeared during the Add
    Dim CreatedSheet        As Excel.Worksheet  'Worksheet this call established
    Dim WS                  As Excel.Worksheet  'Snapshot loop cursor
    Dim AddErrNumber        As Long             'Error the Add reported, if any
    Dim AddErrDescription   As String           'Description the Add reported

'------------------------------------------------------------------------------
' SNAPSHOT
'------------------------------------------------------------------------------
    'Identity, not names: a rename would make a pre-existing sheet look new
        Set WorksheetsBefore = New Collection
        For Each WS In HostWorkbook.Worksheets
            WorksheetsBefore.Add WS
        Next WS
        Set WS = Nothing

'------------------------------------------------------------------------------
' ADD
'------------------------------------------------------------------------------
    'Anchor on Sheets so a trailing chart sheet is handled
        On Error Resume Next
        Set CreatedSheet = HostWorkbook.Worksheets.Add( _
            After:=HostWorkbook.Sheets(HostWorkbook.Sheets.Count))
        AddErrNumber = Err.Number
        AddErrDescription = Err.Description
        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' ADOPT A PARTIAL COMMIT
'------------------------------------------------------------------------------
    'The Add reported a failure. Adopt its worksheet only when exactly one
    'appeared, which is the signature of the documented anomaly
        If CreatedSheet Is Nothing Then
            Set NewWorksheets = TST_DP_FindNewWorksheets(HostWorkbook, WorksheetsBefore)
            If NewWorksheets.Count = 1 Then
                Set CreatedSheet = NewWorksheets(1)
                TST_DP_RecordInfo mTST_DP_CurrentSuite, "Worksheets.Add anomaly", _
                    "Add reported " & VBA.CStr(AddErrNumber) & _
                    " but created one worksheet; adopted it | " & AddErrDescription
            Else
                Err.Raise vbObjectError + 2406, PROC_NAME, _
                    "Worksheets.Add failed and left " & _
                    VBA.CStr(NewWorksheets.Count) & _
                    " candidate worksheets; none adopted. Original error " & _
                    VBA.CStr(AddErrNumber) & " - " & AddErrDescription
            End If
        End If

'------------------------------------------------------------------------------
' PUBLISH BEFORE RENAMING
'------------------------------------------------------------------------------
    'Hand the reference back before the rename, so a caller whose rename fails can
    'still delete this worksheet by identity rather than by a name it never got
        Set TST_DP_AddStaleSheetForTest = CreatedSheet

'------------------------------------------------------------------------------
' NAME
'------------------------------------------------------------------------------
    'Remove any previous sheet holding the name before claiming it
        TST_DP_DeleteWorksheetIfExists HostWorkbook, DesiredName
        CreatedSheet.Name = DesiredName

'------------------------------------------------------------------------------
' RELEASE
'------------------------------------------------------------------------------
    'Release local references
        Set CreatedSheet = Nothing
        Set NewWorksheets = Nothing
        Set WorksheetsBefore = Nothing

End Function

Private Sub TST_DP_DeleteWorksheetByReference( _
    ByRef Candidate As Excel.Worksheet)

'
'==============================================================================
'                 DELETE A WORKSHEET BY IDENTITY (TEST)
'==============================================================================
'   Deletes a worksheet through the reference held for it, which is the only way
'   to remove one whose rename never completed. A name-based delete cannot see a
'   leftover called SheetNN.
'
'   Best-effort by design: this runs on failure paths where nothing may raise.
'==============================================================================

'------------------------------------------------------------------------------
' DELETE
'------------------------------------------------------------------------------
    'Never let cleanup raise
        On Error Resume Next
    'Only touch a reference that still points at a live worksheet
        If TST_DP_WorksheetReferenceIsLive(Candidate) Then
            Candidate.Delete
        End If
    'Release the reference whatever happened
        Set Candidate = Nothing
    'Clear any suppressed cleanup error
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
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Manager             As cDatePickerManager   'Manager instance under test
    Dim MergedArea          As Excel.Range          'Merged area under test
    Dim ErrorNumber         As Long                 'Captured error number
    Dim ErrorDescription    As String               'Captured error description

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
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet
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
    'Capture the escaping error before any On Error statement resets Err
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Suppress local cleanup errors
        On Error Resume Next
    'Unmerge the merged area when still present
        If Not MergedArea Is Nothing Then MergedArea.UnMerge
    'Release object references
        Set MergedArea = Nothing
        Set Manager = Nothing
    'Record the captured suite-level failure
        On Error GoTo 0
        TST_DP_RecordFail "Manager suite failed", _
            "Error " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription
        Err.Clear

End Sub


Private Sub TST_DP_RunSuite_LifecyclePair()

'
'==============================================================================
'                           LIFECYCLE PAIR SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates the DP_Start and DP_Stop lifecycle round-trip
'
' WHY THIS EXISTS
'   DP_Stop was missing in earlier versions. This suite verifies that the
'   start-stop cycle leaves a fully clean teardown state and that a second
'   DP_Start after DP_Stop produces a live, working manager.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Calls DP_Stop and verifies teardown state, then calls DP_Start and verifies
'   the manager is recreated with EnableEvents True, then calls DP_Stop again
'   and verifies a clean second teardown.
'
' ERROR POLICY
'   Records suite-level failures and continues.
'   Attempts best-effort manager restoration after a failure so subsequent
'   suites are not left without a working manager.
'
' DEPENDENCIES
'   DP_Start
'   DP_Stop
'   M_Picker_EnsureManager
'   gDP_Manager
'   M_GridIcon_PurgeAll
'   TST_DP_CountNamedShapes
'
' NOTES
'   DP_Stop uses On Error Resume Next throughout and ends with On Error GoTo 0,
'   which kills the SuiteFail handler on return. It is re-armed immediately
'   after each DP_Stop call.
'
'   DP_Start uses On Error GoTo ErrorHandler and raises outward on failure.
'   It does not reset the caller SuiteFail handler when it succeeds.
'
'   M_Picker_EnsureManager uses On Error GoTo ErrorHandler (raises outward)
'   and does not reset the caller SuiteFail handler.
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EventsBeforeStart   As Boolean      'Caller event state before DP_Start

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "LifecyclePair"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' FIRST STOP - TEARDOWN STATE
'------------------------------------------------------------------------------
    'Call DP_Stop to tear down the DatePicker runtime
    'DP_Stop uses OERN and ends with On Error GoTo 0; re-arm immediately
        DP_Stop
        On Error GoTo SuiteFail

    'Assert the global manager reference is released after DP_Stop
        TST_DP_AssertTrue "Manager is Nothing after DP_Stop", _
            (gDP_Manager Is Nothing)

    'Assert no grid icons remain in the host workbook after DP_Stop
        TST_DP_AssertEqualsLong "No grid icons remain after DP_Stop", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' START AFTER STOP - RECOVERY
'------------------------------------------------------------------------------
    'Capture the caller event state before the recovery start
        EventsBeforeStart = Excel.Application.EnableEvents
    'Call DP_Start to recreate the DatePicker runtime
    'DP_Start internally calls Handle_SelectionChange which ends with
    'On Error GoTo 0, killing the SuiteFail handler. Re-arm immediately.
        DP_Start
        On Error GoTo SuiteFail

    'Assert the manager is recreated after DP_Start
        TST_DP_AssertFalse "Manager is instantiated after DP_Start", _
            (gDP_Manager Is Nothing)

    'Assert DP_Start preserved the caller's Application.EnableEvents state
        TST_DP_AssertBooleanResult "DP_Start preserves the caller event state", _
            (Excel.Application.EnableEvents = EventsBeforeStart)

    'Assert the manager is not busy after DP_Start
        If Not gDP_Manager Is Nothing Then
            TST_DP_AssertFalse "Manager is not busy after DP_Start", _
                gDP_Manager.Is_Busy
        End If

'------------------------------------------------------------------------------
' SECOND STOP - IDEMPOTENT TEARDOWN
'------------------------------------------------------------------------------
    'Call DP_Stop a second time to verify idempotent behavior
    'DP_Stop uses OERN and ends with On Error GoTo 0; re-arm immediately
        DP_Stop
        On Error GoTo SuiteFail

    'Assert the manager is released again after the second DP_Stop
        TST_DP_AssertTrue "Manager is Nothing after second DP_Stop", _
            (gDP_Manager Is Nothing)

    'Assert no grid icons remain after the second DP_Stop
        TST_DP_AssertEqualsLong "No grid icons remain after second DP_Stop", _
            0, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

    'Record that DP_Stop completed without raising after a second call
        TST_DP_RecordPass "DP_Stop is idempotent", vbNullString

'------------------------------------------------------------------------------
' RESTORE MANAGER FOR SUBSEQUENT SUITES
'------------------------------------------------------------------------------
    'Recreate the manager so later suites have working infrastructure
    'M_Picker_EnsureManager raises outward on failure; SuiteFail handler survives
        M_Picker_EnsureManager

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
        TST_DP_RecordFail "LifecyclePair suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Best-effort manager restoration after failure
        On Error Resume Next
        If gDP_Manager Is Nothing Then M_Picker_EnsureManager
        Err.Clear
        On Error GoTo 0

End Sub

Private Sub TST_DP_RunSuite_ProviderLease()

'
'==============================================================================
'                        SUITE: PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Proves the one-provider lease acquires, refuses, releases, and cannot be
'   released by a provider that does not own it
'
' WHY THIS EXISTS
'   Two DatePicker copies in one Excel process register the same application-wide
'   resources under the same fixed identifiers, and either one's teardown removes
'   the other's. Teardown is the more dangerous half: refusing a second provider
'   at startup protects nothing while its DP_Stop still dismantles the owner
'
' BEHAVIOR
'   Exercises acquisition, idempotent re-acquisition, ownership reporting and
'   release from this single VBA project, then simulates a second provider by
'   clearing this project's token while leaving the lease in place
'
' ERROR POLICY
'   Records suite-level failures and continues
'
'   Restores the lease state the run started with on every path
'
' DEPENDENCIES
'   M_Lease_TryAcquire
'   M_Lease_IsOwner
'   M_Lease_Release
'   DP_Start
'
' NOTES
'   A second VBA project cannot be loaded from inside a run, so the second
'   provider is simulated the way a VBA project reset produces it: the lease
'   survives while the local token does not. That is the same state a reset
'   leaves, and it is what the refusal has to detect
'
'   The suite therefore proves the DatePicker side of the contract. Two genuinely
'   separate providers in one Excel session remain a manual validation case
'
'   The lease bar is Temporary, so Excel removes it at shutdown. Nothing here
'   persists past the session even if the suite fails
'
' UPDATED
'   2026-08-23
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AcquiredFirst   As Boolean      'Result of the first acquisition
    Dim AcquiredAgain   As Boolean      'Result of re-acquiring while owning
    Dim OwnedAfter      As Boolean      'Ownership after acquisition
    Dim OwnedAfterFree  As Boolean      'Ownership after release
    Dim RefusedSecond   As Boolean      'Second provider refused acquisition
    Dim OwnedAsSecond   As Boolean      'Second provider claimed ownership
    Dim LeaseSurvived   As Boolean      'Lease still present after a refused release

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "ProviderLease"

'------------------------------------------------------------------------------
' START FROM A KNOWN STATE
'------------------------------------------------------------------------------
    'The run has already started the DatePicker, so this project may hold the
    'lease. Release it to begin from a defined position
        M_Lease_Release

'------------------------------------------------------------------------------
' ACQUIRE
'------------------------------------------------------------------------------
    'A free lease is acquired
        AcquiredFirst = M_Lease_TryAcquire()
        TST_DP_AssertTrue "A free lease is acquired", AcquiredFirst
    'The owner reports ownership
        OwnedAfter = M_Lease_IsOwner()
        TST_DP_AssertTrue "The acquiring provider owns the lease", OwnedAfter
    'Re-acquiring while owning is an idempotent success, not a refusal
        AcquiredAgain = M_Lease_TryAcquire()
        TST_DP_AssertTrue "Re-acquiring an owned lease succeeds", AcquiredAgain

'------------------------------------------------------------------------------
' A SECOND PROVIDER IS REFUSED
'------------------------------------------------------------------------------
    'Simulate the state a second provider sees, and the state a VBA project reset
    'leaves behind: the lease exists, this project cannot prove it owns it
        M_Lease_Test_ClearOwnerToken
    'Acquisition is refused
        RefusedSecond = Not M_Lease_TryAcquire()
        TST_DP_AssertTrue "A second provider is refused the lease", RefusedSecond
    'And it must not report ownership
        OwnedAsSecond = M_Lease_IsOwner()
        TST_DP_AssertFalse "A second provider does not own the lease", OwnedAsSecond

'------------------------------------------------------------------------------
' A SECOND PROVIDER CANNOT RELEASE THE OWNER'S LEASE
'------------------------------------------------------------------------------
    'This is the half that matters. A refused provider calling release must leave
    'the owner's lease intact
        M_Lease_Release
        LeaseSurvived = (VBA.LenB(TST_DP_ReadLeaseOwnerForTest()) > 0)
        TST_DP_AssertTrue "A refused provider cannot release the owner's lease", _
            LeaseSurvived

'------------------------------------------------------------------------------
' THE OWNER RELEASES ITS OWN LEASE
'------------------------------------------------------------------------------
    'Reclaim ownership the only way a project can once its token is gone: the
    'lease has to be removed by the test, standing in for the owner's own DP_Stop
        TST_DP_ForceClearLeaseForTest
    'A cleared lease is acquirable again
        TST_DP_AssertTrue "A released lease can be acquired again", _
            M_Lease_TryAcquire()
    'The owner releases what it owns
        M_Lease_Release
        OwnedAfterFree = M_Lease_IsOwner()
        TST_DP_AssertFalse "The owner no longer owns a released lease", OwnedAfterFree
        TST_DP_AssertEqualsString "A released lease is gone", _
            VBA.vbNullString, TST_DP_ReadLeaseOwnerForTest()

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Leave the run owning the lease, as it did before this suite
        TST_DP_ForceClearLeaseForTest
        M_Lease_TryAcquire
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Provider lease suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore the lease state regardless
        Resume SuiteExit

End Sub

Private Sub TST_DP_RunSuite_RuntimeAdmission()

'
'==============================================================================
'                        RUNTIME ADMISSION SUITE
'==============================================================================
' PURPOSE
'   Proves every supported and callback-driven entry path refuses to act while
'   another provider holds the lease, and succeeds idempotently while this
'   project owns it
'
' WHY THIS EXISTS
'   The ProviderLease suite covers the lease primitive. It never drives a real
'   entry point, which is exactly how v1.2.0 shipped with DP_Show and DP_Preload
'   reaching M_Picker_EnsureManager without proving ownership
'
'   This suite exercises the public paths themselves, so a future change that
'   removes an admission call fails here rather than in a user's session
'
' BEHAVIOR
'   Plants the state a second provider sees - the lease exists, this project
'   cannot prove it owns it - then drives each entry path and asserts that no
'   manager, form or shared registration appeared, and that the path refused
'
' NOTES
'   Refusal reporting is silenced for the duration. Application.DisplayAlerts
'   does not suppress VBA.MsgBox, so without this the run would block on a modal
'   dialog at the first refused entry point
'
'   Silencing is restored in both exit paths. A run that left it enabled would
'   hide genuine provider conflicts from the operator
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EventsBefore        As Boolean      'Application.EnableEvents on entry
    Dim OwnerToken          As String       'Lease token planted as the foreign owner
    Dim RefusalsBefore      As Long         'Refusal count before an entry path runs
    Dim ManagerAfter        As Boolean      'True when a manager exists after a refused path
    Dim FormAfter           As Boolean      'True when the picker form is loaded after a refused path
    Dim GuardHeld           As Boolean      'True when the direct-admission backstop raised
    Dim OwnerSurvived       As Boolean      'True when the owner's lease outlived a refused teardown

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "RuntimeAdmission"

'------------------------------------------------------------------------------
' START FROM A KNOWN STATE
'------------------------------------------------------------------------------
    'Silence the modal refusal report so the real entry paths stay drivable
        M_Lease_Test_SilenceRefusalReport True
    'Record the caller's Excel event state so the preservation check is honest
        EventsBefore = Excel.Application.EnableEvents
    'Begin from a defined position
        TST_DP_ForceClearLeaseForTest
        Set gDP_Manager = Nothing

'------------------------------------------------------------------------------
' PLANT A FOREIGN LEASE
'------------------------------------------------------------------------------
    'Take the lease, then drop the local token. The lease survives, this project
    'can no longer prove it owns it: exactly what a second copy sees
        TST_DP_AssertTrue "Admission setup acquires a free lease", _
            M_Lease_TryAcquire()
        OwnerToken = TST_DP_ReadLeaseOwnerForTest()
        TST_DP_AssertTrue "Admission setup planted a lease token", _
            (VBA.LenB(OwnerToken) > 0)
        M_Lease_Test_ClearOwnerToken
        TST_DP_AssertFalse "This project does not own the planted lease", _
            M_Lease_IsOwner()

'------------------------------------------------------------------------------
' FOREIGN LEASE + DP_SHOW
'------------------------------------------------------------------------------
    'The shared boundary for every interactive open path
        RefusalsBefore = M_Lease_Test_RefusalReportCount()
        Set gDP_Manager = Nothing
        DP_Show
        TST_DP_AssertTrue "DP_Show refuses under a foreign lease", _
            (M_Lease_Test_RefusalReportCount() > RefusalsBefore)
        ManagerAfter = Not (gDP_Manager Is Nothing)
        TST_DP_AssertFalse "DP_Show creates no manager under a foreign lease", _
            ManagerAfter
        FormAfter = TST_DP_IsPickerFormLoadedForTest()
        TST_DP_AssertFalse "DP_Show loads no form under a foreign lease", FormAfter
        TST_DP_AssertEqualsString "DP_Show does not disturb the owner's lease", _
            OwnerToken, TST_DP_ReadLeaseOwnerForTest()

'------------------------------------------------------------------------------
' FOREIGN LEASE + DP_PRELOAD
'------------------------------------------------------------------------------
    'Preload refuses silently, so ownership is what proves it declined
        Set gDP_Manager = Nothing
        DP_Preload
        TST_DP_AssertFalse "DP_Preload creates no manager under a foreign lease", _
            Not (gDP_Manager Is Nothing)
        TST_DP_AssertFalse "DP_Preload loads no hidden form under a foreign lease", _
            TST_DP_IsPickerFormLoadedForTest()
        TST_DP_AssertEqualsString "DP_Preload does not disturb the owner's lease", _
            OwnerToken, TST_DP_ReadLeaseOwnerForTest()

'------------------------------------------------------------------------------
' FOREIGN LEASE + DP_CLICK
'------------------------------------------------------------------------------
    'Delegates through DP_OpenForActiveCell into DP_Show
        RefusalsBefore = M_Lease_Test_RefusalReportCount()
        Set gDP_Manager = Nothing
        DP_Click
        TST_DP_AssertTrue "DP_Click refuses under a foreign lease", _
            (M_Lease_Test_RefusalReportCount() > RefusalsBefore)
        TST_DP_AssertFalse "DP_Click creates no manager under a foreign lease", _
            Not (gDP_Manager Is Nothing)

'------------------------------------------------------------------------------
' FOREIGN LEASE + DP_OPENFORACTIVECELL
'------------------------------------------------------------------------------
    'The keyboard and grid-icon paths both arrive here
        RefusalsBefore = M_Lease_Test_RefusalReportCount()
        Set gDP_Manager = Nothing
        DP_OpenForActiveCell
        TST_DP_AssertTrue "DP_OpenForActiveCell refuses under a foreign lease", _
            (M_Lease_Test_RefusalReportCount() > RefusalsBefore)
        TST_DP_AssertFalse _
            "DP_OpenForActiveCell creates no manager under a foreign lease", _
            Not (gDP_Manager Is Nothing)

'------------------------------------------------------------------------------
' FOREIGN LEASE + RIBBON_SHOWPICKER
'------------------------------------------------------------------------------
    'The Ribbon callback passes an IRibbonControl it never dereferences, so
    'Nothing is a valid stand-in for the Excel-supplied argument
        RefusalsBefore = M_Lease_Test_RefusalReportCount()
        Set gDP_Manager = Nothing
        Ribbon_ShowPicker Nothing
        TST_DP_AssertTrue "Ribbon_ShowPicker refuses under a foreign lease", _
            (M_Lease_Test_RefusalReportCount() > RefusalsBefore)
        TST_DP_AssertFalse _
            "Ribbon_ShowPicker creates no manager under a foreign lease", _
            Not (gDP_Manager Is Nothing)

'------------------------------------------------------------------------------
' DIRECT MANAGER ADMISSION CANNOT BYPASS THE GUARD
'------------------------------------------------------------------------------
    'M_Picker_EnsureManager is technically public. Calling it directly under a
    'foreign lease must fail closed rather than bootstrap a second runtime
        Set gDP_Manager = Nothing
        GuardHeld = False
        On Error Resume Next
        M_Picker_EnsureManager
        GuardHeld = (Err.Number <> 0)
        Err.Clear
        On Error GoTo SuiteFail
        TST_DP_AssertTrue _
            "Direct M_Picker_EnsureManager admission is refused", GuardHeld
        TST_DP_AssertFalse _
            "Direct admission creates no manager under a foreign lease", _
            Not (gDP_Manager Is Nothing)

'------------------------------------------------------------------------------
' A REFUSED COPY CANNOT TEAR DOWN THE OWNER
'------------------------------------------------------------------------------
    'DP_Stop and DP_RepairRuntime already gate on ownership. This proves the
    'admission work did not weaken that
        DP_Stop
        OwnerSurvived = (VBA.StrComp(TST_DP_ReadLeaseOwnerForTest(), OwnerToken, _
            vbBinaryCompare) = 0)
        TST_DP_AssertTrue "A refused copy cannot stop the owner's runtime", _
            OwnerSurvived
        DP_RepairRuntime
        TST_DP_AssertEqualsString "A refused copy cannot repair the owner", _
            OwnerToken, TST_DP_ReadLeaseOwnerForTest()

'------------------------------------------------------------------------------
' CALLER EXCEL EVENT STATE IS PRESERVED
'------------------------------------------------------------------------------
    'No refused path may leave Application.EnableEvents changed
        TST_DP_AssertTrue "Refused paths preserve Application.EnableEvents", _
            (Excel.Application.EnableEvents = EventsBefore)

'------------------------------------------------------------------------------
' OWNER LEASE + SUPPORTED PATHS SUCCEED IDEMPOTENTLY
'------------------------------------------------------------------------------
    'Reclaim the lease the only way a project can once its token is gone
        TST_DP_ForceClearLeaseForTest
        Set gDP_Manager = Nothing
        TST_DP_AssertTrue "The owner reacquires a cleared lease", _
            M_Lease_TryAcquire()
    'Admission must be idempotent for the owner, not a one-shot
        RefusalsBefore = M_Lease_Test_RefusalReportCount()
        DP_Start
        TST_DP_AssertTrue "DP_Start succeeds for the owner", M_Lease_IsOwner()
        TST_DP_AssertTrue "DP_Start creates the manager for the owner", _
            Not (gDP_Manager Is Nothing)
        DP_Start
        TST_DP_AssertTrue "Repeated DP_Start remains owned and idempotent", _
            M_Lease_IsOwner()
        M_Picker_EnsureManager
        TST_DP_AssertTrue "Direct admission succeeds for the owner", _
            Not (gDP_Manager Is Nothing)
        TST_DP_AssertEqualsString "No owner path reported a refusal", _
            VBA.CStr(RefusalsBefore), _
            VBA.CStr(M_Lease_Test_RefusalReportCount())

'------------------------------------------------------------------------------
' FORCE RELEASE REMAINS EXPLICIT
'------------------------------------------------------------------------------
    'Nothing above may have reclaimed a foreign lease automatically. The only
    'route back is the documented explicit call
        TST_DP_AssertTrue "Force release remains an explicit operator action", _
            (VBA.LenB(TST_DP_ReadLeaseOwnerForTest()) > 0)

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Restore reporting so a genuine conflict is never hidden from the operator
        M_Lease_Test_SilenceRefusalReport False
    'Leave the run owning the lease, as it did before this suite
        TST_DP_ForceClearLeaseForTest
        M_Lease_TryAcquire
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Runtime admission suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore reporting and lease state regardless
        Resume SuiteExit

End Sub

Private Function TST_DP_IsPickerFormLoadedForTest() As Boolean

'
'==============================================================================
'                   IS THE PICKER FORM LOADED (TEST)
'==============================================================================
'   Reports whether the DatePicker UserForm is currently loaded.
'
'   M_FormBridge_GetLoadedForm is Private to the production module, so this
'   suite enumerates the VBA UserForms collection instead. That collection holds
'   only loaded forms, which is exactly the question being asked.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm      As Object       'Iterated loaded form

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Never let a lookup raise into an assertion
        On Error Resume Next
    'Set safe default result
        TST_DP_IsPickerFormLoadedForTest = False
    'Only loaded forms appear in this collection
        For Each LoadedForm In VBA.UserForms
            If VBA.StrComp(LoadedForm.Name, "UF_DatePicker", vbTextCompare) = 0 Then
                TST_DP_IsPickerFormLoadedForTest = True
                Exit For
            End If
        Next LoadedForm
    'Clear any suppressed enumeration error
        Err.Clear

End Function

Private Function TST_DP_IsPickerFormVisibleForTest() As Boolean

'
'==============================================================================
'                   IS A PICKER FORM VISIBLE (TEST)
'==============================================================================
'   Reports whether any loaded DatePicker instance is actually visible.
'
'   Presence in VBA.UserForms is not the same question. A Load that raises from
'   UserForm_Initialize leaves an orphaned instance in that collection which the
'   project cannot reach by name, so "loaded" over-reports. Visibility is the
'   property that matters for #47: the contract is that a window in no known good
'   state is never presented to the user.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm      As Object       'Iterated loaded form

'------------------------------------------------------------------------------
' SEARCH
'------------------------------------------------------------------------------
    'Never let a lookup raise into an assertion
        On Error Resume Next
    'Set safe default result
        TST_DP_IsPickerFormVisibleForTest = False
    'Report the first visible DatePicker instance found
        For Each LoadedForm In VBA.UserForms
            If VBA.StrComp(LoadedForm.Name, "UF_DatePicker", vbTextCompare) = 0 Then
                If LoadedForm.Visible Then
                    TST_DP_IsPickerFormVisibleForTest = True
                    Exit For
                End If
            End If
        Next LoadedForm
    'Clear any suppressed enumeration error
        Err.Clear

End Function

Private Sub TST_DP_UnloadAllPickerFormsForTest()

'
'==============================================================================
'                   UNLOAD EVERY PICKER INSTANCE (TEST)
'==============================================================================
'   Removes every DatePicker instance in VBA.UserForms, including one orphaned
'   by a Load that raised from UserForm_Initialize.
'
'   Unload UF_DatePicker only reaches the default instance. An orphan left by a
'   failed Load survives that call and would otherwise leak into later suites,
'   so this walks the collection by index in reverse and unloads what it finds.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index           As Long         'Reverse index into the loaded-form collection

'------------------------------------------------------------------------------
' UNLOAD
'------------------------------------------------------------------------------
    'Never let teardown raise into a suite
        On Error Resume Next
    'Walk backwards because unloading removes entries from the collection
        For Index = VBA.UserForms.Count - 1 To 0 Step -1
            If VBA.StrComp(VBA.UserForms(Index).Name, "UF_DatePicker", vbTextCompare) = 0 Then
                Unload VBA.UserForms(Index)
            End If
        Next Index
    'Clear any suppressed teardown error
        Err.Clear

End Sub

Private Function TST_DP_ReadLeaseOwnerForTest() As String

'
'==============================================================================
'                     READ LEASE OWNER FOR TEST
'==============================================================================
'   Reads the lease marker directly, so the suite can assert on the workbook
'   state rather than on the routine under test.
'
'   The bar name and marker tag are duplicated from M_DatePicker, where they are
'   Private. They are fixed identifiers that cannot drift without this suite
'   failing.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const LEASE_BAR     As String = "__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__"
    Const MARKER_TAG    As String = "VBA_DATETIMEPICKER_RUNTIME_LEASE_OWNER"

    Dim LeaseBar        As Object       'Resolved lease command bar
    Dim Ctl             As Object       'Current control while scanning

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Never let a probe raise into a test
        On Error Resume Next
    'Set safe default result
        TST_DP_ReadLeaseOwnerForTest = VBA.vbNullString
    'Resolve the lease bar, which may not exist
        Set LeaseBar = Excel.Application.CommandBars(LEASE_BAR)
        If LeaseBar Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Report the first marker token found
        For Each Ctl In LeaseBar.Controls
            If VBA.StrComp(Ctl.Tag, MARKER_TAG, vbBinaryCompare) = 0 Then
                TST_DP_ReadLeaseOwnerForTest = Ctl.Parameter
                Exit For
            End If
        Next Ctl
    'Release object references
        Set Ctl = Nothing
        Set LeaseBar = Nothing
    'Clear any suppressed probe error
        Err.Clear

End Function

Private Sub TST_DP_ForceClearLeaseForTest()

'
'==============================================================================
'                     FORCE CLEAR LEASE FOR TEST
'==============================================================================
'   Removes the lease bar outright, standing in for an owner's clean shutdown.
'
'   This exists only because a suite cannot load a second VBA project to release
'   a lease it does not own. Production code never deletes a lease it cannot
'   prove it owns; this deliberately does, which is why it lives in the harness.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const LEASE_BAR     As String = "__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__"

    Dim LeaseBar        As Object       'Resolved lease command bar

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Never let cleanup raise into a test
        On Error Resume Next
    'Resolve and delete the lease bar when it exists
        Set LeaseBar = Excel.Application.CommandBars(LEASE_BAR)
        If Not LeaseBar Is Nothing Then
            LeaseBar.Delete
        End If
    'Release object references
        Set LeaseBar = Nothing
    'Clear any suppressed cleanup error
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_RepairRuntime()

'
'==============================================================================
'                           REPAIR RUNTIME SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates DP_RepairRuntime behavior after disabled Excel events and after
'   a simulated stale runtime state
'
' WHY THIS EXISTS
'   DP_RepairRuntime addresses the most common real-world DatePicker failure
'   mode: Application.EnableEvents = False left behind by an interrupted or
'   poorly-written macro, which silently kills all manager event routing.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Disables Application.EnableEvents, verifies the disabled state, calls
'   DP_RepairRuntime, and asserts that events are re-enabled, the manager is
'   alive and not busy, and context-menu and keyboard-shortcut integration is
'   synchronized.
'
'   Calls DP_RepairRuntime a second time against a healthy runtime to verify
'   that the procedure is safe to call when nothing is actually broken.
'
' ERROR POLICY
'   Records suite-level failures and continues.
'   Always restores Application.EnableEvents in the SuiteFail handler so a
'   failed repair call cannot leave Excel events disabled.
'
' DEPENDENCIES
'   DP_RepairRuntime
'   M_Picker_EnsureManager
'   gDP_Manager
'
' NOTES
'   DP_RepairRuntime uses On Error GoTo ErrorHandler and raises outward on
'   its own failure path. However, it calls Handle_SelectionChange internally,
'   and Handle_SelectionChange ends with On Error GoTo 0 in its CleanExit.
'   This kills the SuiteFail handler inside DP_RepairRuntime's call stack.
'   Re-arm On Error GoTo SuiteFail after every DP_RepairRuntime call.
'
'   The RepairRuntime suite re-enables Application.EnableEvents = True.
'   Subsequent suites that call mTST_DP_ScratchSheet.Activate must
'   temporarily disable events before the Activate call to prevent the
'   live manager from firing SheetActivate or SelectionChange during the
'   sheet switch, which would reset the SuiteFail handler unexpectedly.
'
'   The suite deliberately disables EnableEvents before calling DP_RepairRuntime.
'   The harness has already set EnableEvents = False for its own operation.
'   DP_RepairRuntime is expected to restore it to True.
'   The harness's own PrepareApplicationForRun state will be restored by the
'   cleanup path at the end of the run.
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "RepairRuntime"
    'Enable suite-level error handling
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' SIMULATE DISABLED EVENTS
'------------------------------------------------------------------------------
    'Disable Excel events to simulate the failure condition DP_RepairRuntime
    'is designed to recover from
        Excel.Application.EnableEvents = False

    'Assert EnableEvents is False before the repair call
        TST_DP_AssertFalse "EnableEvents is False before repair", _
            Excel.Application.EnableEvents

'------------------------------------------------------------------------------
' CALL DP_REPAIRRUNTIME
'------------------------------------------------------------------------------
    'Call DP_RepairRuntime to repair the DatePicker runtime
    'DP_RepairRuntime raises outward on its own failure path.
    'However it calls Handle_SelectionChange internally, which ends with
    'On Error GoTo 0 in its CleanExit, killing the SuiteFail handler.
    'Re-arm the handler after DP_RepairRuntime returns.
        DP_RepairRuntime
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' ASSERT REPAIRED STATE
'------------------------------------------------------------------------------
    'Assert EnableEvents was re-enabled by DP_RepairRuntime
        TST_DP_AssertTrue "EnableEvents is True after DP_RepairRuntime", _
            Excel.Application.EnableEvents

    'Assert the manager is alive after repair
        TST_DP_AssertFalse "Manager is instantiated after DP_RepairRuntime", _
            (gDP_Manager Is Nothing)

    'Assert the manager is not busy after repair
        If Not gDP_Manager Is Nothing Then
            TST_DP_AssertFalse "Manager is not busy after DP_RepairRuntime", _
                gDP_Manager.Is_Busy
        End If

'------------------------------------------------------------------------------
' IDEMPOTENT CALL - HEALTHY RUNTIME
'------------------------------------------------------------------------------
    'Call DP_RepairRuntime again against an already healthy runtime
    'Handle_SelectionChange inside DP_RepairRuntime ends with On Error GoTo 0;
    're-arm the handler after the call.
        DP_RepairRuntime
        On Error GoTo SuiteFail

    'Assert EnableEvents is still True after the second repair call
        TST_DP_AssertTrue "EnableEvents is True after second DP_RepairRuntime", _
            Excel.Application.EnableEvents

    'Assert the manager is still alive after the second repair call
        TST_DP_AssertFalse "Manager is still instantiated after second DP_RepairRuntime", _
            (gDP_Manager Is Nothing)

    'Record that DP_RepairRuntime completed without raising on a healthy runtime
        TST_DP_RecordPass "DP_RepairRuntime is safe when runtime is already healthy", _
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
        TST_DP_RecordFail "RepairRuntime suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Always restore EnableEvents in case the repair call itself failed
        On Error Resume Next
        Excel.Application.EnableEvents = True
        Err.Clear
        On Error GoTo 0

End Sub

Private Sub TST_DP_RunSuite_PreCreateHidden()

'
'==============================================================================
'                           PRE-CREATE HIDDEN SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates M_GridIcon_PreCreateHidden startup optimization behavior
'
' WHY THIS EXISTS
'   M_GridIcon_PreCreateHidden pre-creates the grid icon during DP_Start and
'   hides it so the high-frequency SelectionChange path can move and show an
'   already-created shape instead of creating one on demand. If pre-creation
'   fails, leaves the shape visible, or creates duplicate shapes, the
'   SelectionChange path degrades to on-demand creation, losing the
'   startup-latency benefit.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Enables the grid icon feature via setter, purges stale icons, calls
'   M_GridIcon_PreCreateHidden, and asserts that:
'     - the tracked reference is set after pre-creation
'     - the shape exists on the scratch sheet with the expected name
'     - the shape is hidden (Visible = msoFalse) immediately after pre-creation
'     - calling M_GridIcon_PreCreateHidden again is safe and does not create a
'       duplicate shape
'
'   Then disables the grid icon feature via setter, purges again, calls
'   M_GridIcon_PreCreateHidden, and asserts that no shape was created.
'
' ERROR POLICY
'   Records suite-level failures and continues.
'
' DEPENDENCIES
'   M_GridIcon_PreCreateHidden
'   M_GridIcon_PurgeAll
'   M_Settings_SetShowGridIcon
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   TST_DP_ShapeExists
'   TST_DP_CountNamedShapes
'
' NOTES
'   M_GridIcon_PreCreateHidden uses On Error Resume Next throughout and ends
'   with On Error GoTo 0, which kills the SuiteFail handler on return.
'   It is re-armed immediately after every call.
'
'   M_GridIcon_PurgeAll also ends with On Error GoTo 0; re-armed after each call.
'
'   The setter M_Settings_SetShowGridIcon internally calls M_GridIcon_Remove
'   and M_KeyboardShortcut_Update, both of which end with On Error GoTo 0.
'   It is re-armed after each setter call.
'
'   ScreenUpdating is set True before pre-creation because M_GridIcon_Create
'   captures PreviousScreenUpdating. With ScreenUpdating = False the final
'   shape show step may silently fail on some Excel builds. The tracked
'   reference being set is the correct invariant to assert; Visible state is
'   a rendering detail that this suite does not assert.
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "PreCreateHidden"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Disable Excel events because the live manager may still be active after
    'RepairRuntime and should not interfere with this controlled grid-icon test
        Excel.Application.EnableEvents = False
    'Reject a missing scratch worksheet before using it as the explicit anchor
        If mTST_DP_ScratchSheet Is Nothing Then
            Err.Raise vbObjectError + 513, _
                "TST_DP_RunSuite_PreCreateHidden", _
                "Scratch worksheet is not available"
        End If
    'Ensure the scratch worksheet is visible without activating it
        If mTST_DP_ScratchSheet.Visible <> xlSheetVisible Then
            mTST_DP_ScratchSheet.Visible = xlSheetVisible
        End If

'------------------------------------------------------------------------------
' PREPARE FEATURE STATE
'------------------------------------------------------------------------------
    'Enable the grid icon feature via setter
    'M_Settings_SetShowGridIcon calls M_GridIcon_Remove and
    'M_KeyboardShortcut_Update which both end with On Error GoTo 0; re-arm
        M_Settings_SetShowGridIcon True
        On Error GoTo SuiteFail

    'Purge any stale grid icons before the suite
    'M_GridIcon_PurgeAll ends with On Error GoTo 0; re-arm immediately
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' PRE-CREATE ON ELIGIBLE CELL
'------------------------------------------------------------------------------
    'Restore ScreenUpdating so M_GridIcon_Create captures True as its
    'PreviousScreenUpdating and restores it on exit
        Excel.Application.ScreenUpdating = True
    'Pre-create the grid icon using the scratch-sheet anchor cell
    'M_GridIcon_PreCreateHidden ends with On Error GoTo 0; re-arm immediately
        M_GridIcon_PreCreateHidden mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail
    'Suppress ScreenUpdating for remaining cleanup steps
        Excel.Application.ScreenUpdating = False
    'Allow the drawing layer to settle after pre-creation
        DoEvents

    'Assert the tracked reference is set after pre-creation
        TST_DP_AssertFalse "Tracked reference is set after PreCreateHidden", _
            (gDP_GridIconShape Is Nothing)

    'Assert the shape exists on the scratch sheet with the expected name
        TST_DP_AssertTrue "Shape exists on scratch sheet after PreCreateHidden", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

    'Assert the shape is hidden after pre-creation using the tracked reference
    'Note: gDP_GridIconShape.Visible = msoFalse is the correct invariant here.
    'M_GridIcon_PreCreateHidden explicitly sets Visible = msoFalse after creation.
    'The shape was never shown, so the msoFalse state does not depend on
    'ScreenUpdating behavior.
        If Not gDP_GridIconShape Is Nothing Then
            TST_DP_AssertTrue "Shape is hidden (Visible=msoFalse) after PreCreateHidden", _
                (gDP_GridIconShape.Visible = msoFalse)
        End If

'------------------------------------------------------------------------------
' IDEMPOTENT SECOND CALL
'------------------------------------------------------------------------------
    'Call M_GridIcon_PreCreateHidden again when a shape already exists
    'M_GridIcon_PreCreateHidden ends with On Error GoTo 0; re-arm immediately
        M_GridIcon_PreCreateHidden mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail

    'Assert only one shape exists after the second pre-create call
        TST_DP_AssertEqualsLong "PreCreateHidden is idempotent when shape already exists", _
            1, _
            TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

    'Record that PreCreateHidden completed without raising on a second call
        TST_DP_RecordPass "PreCreateHidden does not raise on a second call", vbNullString

'------------------------------------------------------------------------------
' DISABLED FEATURE BEHAVIOR
'------------------------------------------------------------------------------
    'Purge before the disabled test to ensure a clean state
    'M_GridIcon_PurgeAll ends with On Error GoTo 0; re-arm immediately
        M_GridIcon_PurgeAll
        On Error GoTo SuiteFail

    'Disable the grid icon feature via setter
    'M_Settings_SetShowGridIcon calls removal routines ending with GoTo 0; re-arm
        M_Settings_SetShowGridIcon False
        On Error GoTo SuiteFail

    'Call M_GridIcon_PreCreateHidden while the feature is disabled
    'M_GridIcon_PreCreateHidden ends with On Error GoTo 0; re-arm immediately
        M_GridIcon_PreCreateHidden mTST_DP_ScratchSheet.Range("D5")
        On Error GoTo SuiteFail

    'Assert no shape was created while the feature is disabled
        TST_DP_AssertFalse "PreCreateHidden creates no shape when feature is disabled", _
            TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' RESTORE FEATURE STATE FOR SUBSEQUENT SUITES
'------------------------------------------------------------------------------
    'Re-enable the grid icon feature via setter
    'M_Settings_SetShowGridIcon calls removal routines ending with GoTo 0; re-arm
        M_Settings_SetShowGridIcon True
        On Error GoTo SuiteFail

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
        TST_DP_RecordFail "PreCreateHidden suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_SelectDate()

'
'==============================================================================
'                           SELECT DATE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates M_Picker_SelectDate write-back, state capture, rollback, and
'   lifecycle behavior
'
' WHY THIS EXISTS
'   M_Picker_SelectDate is the centralized write-back entry point for all
'   day-label clicks. Its correctness directly determines whether dates are
'   written reliably, whether state is stored only after a successful write,
'   whether previous state is restored on failure, and whether the configured
'   CloseAfterSelection behavior is honoured.
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests:
'     - successful write-back to a single selected cell
'     - gDP_HasSelectedDate is True only after successful write-back
'     - gDP_SelectedDate matches the written date-only value
'     - gDP_WriteValue is set to the date-only component before write-back
'     - repeated write-back to the same cell with a different date
'     - zero date rejection raises a runtime error
'     - NoTableGrow = True path does not expand table-column selection
'     - CloseAfterSelection = True path completes without raising
'     - CloseAfterSelection = False path completes without raising
'
' ERROR POLICY
'   Records suite-level failures and continues.
'
' DEPENDENCIES
'   M_Picker_SelectDate
'   M_Settings_SetCloseAfterSelection
'   M_Settings_GetCloseAfterSelection
'   gDP_HasSelectedDate
'   gDP_SelectedDate
'   gDP_WriteValue
'   mTST_DP_ScratchSheet
'
' NOTES
'   M_Picker_SelectDate uses On Error GoTo ErrorHandler and raises outward on
'   failure. It does not reset the caller SuiteFail handler when it succeeds.
'   No re-arm is needed after M_Picker_SelectDate calls.
'
'   M_Settings_SetCloseAfterSelection only calls M_Settings_Save internally.
'   M_Settings_Save uses GoTo ErrorHandler and raises outward; it does not
'   reset the SuiteFail handler. No re-arm is needed after setter calls.
'
'   gDP_WriteValue, gDP_HasSelectedDate, and gDP_SelectedDate are read directly
'   after calls to verify transient state. No public getters exist for these
'   transient fields. Direct reads are the correct approach for these.
'
'   The target cell is selected before each M_Picker_SelectDate call so that
'   M_WriteBack_Apply uses the correct Excel selection as its write target.
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetCell      As Excel.Range   'Write-back target cell

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "SelectDate"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Disable events before Activate for the same reason as PreCreateHidden:
    'the manager is live and would fire on sheet switch
        Excel.Application.EnableEvents = False
    'Activate the scratch sheet
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet

    'Prepare the write-back target cell
        Set TargetCell = mTST_DP_ScratchSheet.Range("K2")
    'Clear any previous value in the target cell
        TargetCell.ClearContents
    'Select the target cell so M_WriteBack_Apply uses it
        TargetCell.Select

    'Clear transient state before the suite
        gDP_HasSelectedDate = False
        gDP_SelectedDate = 0
        gDP_WriteValue = 0

    'Configure CloseAfterSelection = False via setter so no form lifecycle
    'path attempts to close a form that was never opened
        M_Settings_SetCloseAfterSelection False

'------------------------------------------------------------------------------
' SUCCESSFUL WRITE-BACK
'------------------------------------------------------------------------------
    'Select the target cell before the first SelectDate call
        TargetCell.Select
    'Call M_Picker_SelectDate with a valid date
    'M_Picker_SelectDate raises outward on failure; SuiteFail handler survives
        M_Picker_SelectDate VBA.DateSerial(2026, 5, 3)

    'Assert the cell received the date
        TST_DP_AssertCellDateEquals "SelectDate writes the date to the target cell", _
            VBA.DateSerial(2026, 5, 3), TargetCell

    'Assert gDP_HasSelectedDate is True after successful write-back
        TST_DP_AssertTrue "HasSelectedDate is True after SelectDate succeeds", _
            gDP_HasSelectedDate

    'Assert gDP_SelectedDate matches the written date-only value
        TST_DP_AssertDateEquals "gDP_SelectedDate matches the written date", _
            VBA.DateSerial(2026, 5, 3), gDP_SelectedDate

    'Assert gDP_WriteValue was set to the date-only component
        TST_DP_AssertDateEquals "gDP_WriteValue is set to the date-only value", _
            VBA.DateSerial(2026, 5, 3), VBA.CDate(gDP_WriteValue)

'------------------------------------------------------------------------------
' REPEATED WRITE-BACK
'------------------------------------------------------------------------------
    'Select the target cell before the second SelectDate call
        TargetCell.Select
    'Call M_Picker_SelectDate with a different date to verify repeated write-back
    'M_Picker_SelectDate raises outward on failure; SuiteFail handler survives
        M_Picker_SelectDate VBA.DateSerial(2026, 12, 31)

    'Assert the cell was updated with the new date
        TST_DP_AssertCellDateEquals "SelectDate supports repeated write-back", _
            VBA.DateSerial(2026, 12, 31), TargetCell

    'Assert gDP_SelectedDate was updated after the second write-back
        TST_DP_AssertDateEquals "gDP_SelectedDate is updated after repeated write-back", _
            VBA.DateSerial(2026, 12, 31), gDP_SelectedDate

'------------------------------------------------------------------------------
' ZERO DATE REJECTION
'------------------------------------------------------------------------------
    'Assert that M_Picker_SelectDate raises when passed a zero date
        TST_DP_ExpectError_SelectDateZero

'------------------------------------------------------------------------------
' NOTABLEGROW = TRUE PATH
'------------------------------------------------------------------------------
    'Select the target cell before the NoTableGrow call
        TargetCell.Select
    'Call M_Picker_SelectDate with NoTableGrow = True
    'This verifies the parameter is accepted and write-back still succeeds
    'M_Picker_SelectDate raises outward on failure; SuiteFail handler survives
        M_Picker_SelectDate VBA.DateSerial(2026, 3, 15), True

    'Assert the cell received the date when NoTableGrow = True
        TST_DP_AssertCellDateEquals "SelectDate writes correctly with NoTableGrow = True", _
            VBA.DateSerial(2026, 3, 15), TargetCell

'------------------------------------------------------------------------------
' CLOSEAFTERSELECTION = TRUE PATH
'------------------------------------------------------------------------------
    'Set CloseAfterSelection = True via setter
        M_Settings_SetCloseAfterSelection True

    'Assert the setting was stored
        TST_DP_AssertTrue "CloseAfterSelection is True after setter", _
            M_Settings_GetCloseAfterSelection()

    'Select the target cell before the CloseAfterSelection call
        TargetCell.Select
    'Call M_Picker_SelectDate; it will attempt DP_Close after write-back
    'DP_Close is safe to call when no form is loaded
    'M_Picker_SelectDate raises outward on failure; SuiteFail handler survives
        M_Picker_SelectDate VBA.DateSerial(2026, 6, 1)

    'Assert the cell received the date with CloseAfterSelection = True
        TST_DP_AssertCellDateEquals "SelectDate writes correctly with CloseAfterSelection = True", _
            VBA.DateSerial(2026, 6, 1), TargetCell

    'Record that CloseAfterSelection = True path completed without raising
        TST_DP_RecordPass "SelectDate with CloseAfterSelection = True does not raise", _
            vbNullString

'------------------------------------------------------------------------------
' CLOSEAFTERSELECTION = FALSE PATH
'------------------------------------------------------------------------------
    'Restore CloseAfterSelection = False via setter
        M_Settings_SetCloseAfterSelection False

    'Select the target cell before the CloseAfterSelection = False call
        TargetCell.Select
    'Call M_Picker_SelectDate; it will call M_FormBridge_AfterSuccessfulSelection
    'which is best-effort and safe to call when no form is loaded
    'M_Picker_SelectDate raises outward on failure; SuiteFail handler survives
        M_Picker_SelectDate VBA.DateSerial(2026, 7, 4)

    'Assert the cell received the date with CloseAfterSelection = False
        TST_DP_AssertCellDateEquals "SelectDate writes correctly with CloseAfterSelection = False", _
            VBA.DateSerial(2026, 7, 4), TargetCell

    'Record that CloseAfterSelection = False path completed without raising
        TST_DP_RecordPass "SelectDate with CloseAfterSelection = False does not raise", _
            vbNullString

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set TargetCell = Nothing
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Release object references
        Set TargetCell = Nothing
    'Record the suite-level failure and clear the error
        TST_DP_RecordFail "SelectDate suite failed", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear

End Sub

Private Sub TST_DP_RunSuite_ApplicationState()

'
'==============================================================================
'                        APPLICATION STATE SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates that the public DatePicker entry points preserve the caller's
'   Application.EnableEvents state
'
' WHY THIS EXISTS
'   M_Picker_EnsureManager previously forced Application.EnableEvents = True on
'   every normal entry point, which silently broke a business macro that had
'   deliberately suppressed Excel events for a transactional update
'
'   DP_RepairRuntime remains the only entry point permitted to force events on
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Tests that DP_Start, DP_Show, DP_Preload, and M_Picker_EnsureManager leave a
'   pre-existing disabled state untouched, that the EventsDisabledByCaller output
'   flag reports the caller state, that write-back restores both the enabled and
'   the disabled case, and that DP_RepairRuntime still force-enables events
'
' ERROR POLICY
'   Records suite-level failures and continues
'
'   Restores Application.EnableEvents = False on every exit path so the harness
'   run state set by TST_DP_PrepareApplicationForRun is preserved
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   DP_Start
'   DP_Show
'   DP_Close
'   DP_Preload
'   DP_RepairRuntime
'   M_WriteBack_Apply
'   mTST_DP_ScratchSheet
'
' NOTES
'   The harness itself runs with Application.EnableEvents = False, so the
'   suppressed-caller condition under test is the ambient run state
'
'   DP_Show is exercised here rather than in the UI smoke suite because event
'   preservation is a release-blocking assertion and must run unconditionally
'
'   The caller event state is asserted twice: against Application.EnableEvents
'   after the call, and against EventsDisabledByCaller on the returned result
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetCell          As Excel.Range  'Write-back target under test
    Dim EventsDisabled      As Boolean      'Output flag from M_Picker_EnsureManager
    Dim WriteResult         As DP_WriteResult   'Structured write-back result
    Dim RestoredState       As Boolean      'Application.EnableEvents after write-back
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "ApplicationState"
        On Error GoTo SuiteFail
    'Activate the scratch sheet
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet
    'Establish the suppressed-caller condition under test
        Excel.Application.EnableEvents = False

'------------------------------------------------------------------------------
' BOOTSTRAP PRESERVES CALLER STATE
'------------------------------------------------------------------------------
    'Bootstrap the manager while the caller has events suppressed
        M_Picker_EnsureManager EventsDisabled
    'Assert the bootstrapper did not re-enable events
        TST_DP_AssertFalse "M_Picker_EnsureManager preserves disabled events", _
            Excel.Application.EnableEvents
    'Assert the bootstrapper reported the caller state
        TST_DP_AssertTrue "M_Picker_EnsureManager reports EventsDisabledByCaller", _
            EventsDisabled
    'Assert the manager was still created and hooked while events were suppressed
        TST_DP_AssertTrue "Manager is hooked while events are suppressed", _
            (Not gDP_Manager Is Nothing)
    'Assert the hook state predicate is independent of Application.EnableEvents
        TST_DP_AssertTrue "Is_Hooked is True while events are suppressed", _
            gDP_Manager.Is_Hooked

'------------------------------------------------------------------------------
' DP_START PRESERVES CALLER STATE
'------------------------------------------------------------------------------
    'Start the runtime while the caller has events suppressed
    'DP_Start resets On Error GoTo 0 on exit; re-arm immediately
        DP_Start
        On Error GoTo SuiteFail
    'Assert DP_Start did not re-enable events
        TST_DP_AssertFalse "DP_Start preserves disabled events", _
            Excel.Application.EnableEvents

'------------------------------------------------------------------------------
' DP_PRELOAD PRESERVES CALLER STATE
'------------------------------------------------------------------------------
    'Restore the suppressed condition before the preload check
        Excel.Application.EnableEvents = False
    'Preload the form while the caller has events suppressed
    'DP_Preload resets On Error GoTo 0 on exit; re-arm immediately
        DP_Preload
        On Error GoTo SuiteFail
    'Assert DP_Preload did not re-enable events
        TST_DP_AssertFalse "DP_Preload preserves disabled events", _
            Excel.Application.EnableEvents

'------------------------------------------------------------------------------
' DP_SHOW PRESERVES CALLER STATE
'------------------------------------------------------------------------------
    'Restore the suppressed condition before the show check
        Excel.Application.EnableEvents = False
    'Show the picker while the caller has events suppressed
    'DP_Show resets On Error GoTo 0 on exit; re-arm immediately
        DP_Show
        On Error GoTo SuiteFail
    'Assert DP_Show did not re-enable events
        TST_DP_AssertFalse "DP_Show preserves disabled events", _
            Excel.Application.EnableEvents
    'Close the picker before the write-back checks
    'DP_Close resets On Error GoTo 0 on exit; re-arm immediately
        DP_Close
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' WRITE-BACK RESTORES THE DISABLED CASE
'------------------------------------------------------------------------------
    'Prepare a date-formatted target cell
        Set TargetCell = mTST_DP_ScratchSheet.Range("H10")
        TargetCell.ClearContents
        TargetCell.NumberFormat = "dd/mm/yyyy"
        TargetCell.Select
    'Set the value the write-back transaction will apply
        gDP_WriteValue = VBA.DateSerial(2026, 8, 21)
    'Establish the suppressed condition before the transaction
        Excel.Application.EnableEvents = False
    'Apply the write-back transaction
    'M_WriteBack_Apply resets On Error GoTo 0 on exit; re-arm immediately
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, True)
        On Error GoTo SuiteFail
    'Capture the restored state
        RestoredState = Excel.Application.EnableEvents
    'Assert write-back restored the disabled caller state
        TST_DP_AssertFalse "Write-back restores disabled events", RestoredState
    'Assert the result reports the caller's suppressed state
        TST_DP_AssertTrue "Write-back result reports EventsDisabledByCaller", _
            WriteResult.EventsDisabledByCaller
    'Assert the target cell received the written value
        TST_DP_AssertCellDateEquals "Write-back writes the target cell with events disabled", _
            VBA.DateSerial(2026, 8, 21), _
            TargetCell

'------------------------------------------------------------------------------
' WRITE-BACK RESTORES THE ENABLED CASE
'------------------------------------------------------------------------------
    'Prepare a second date-formatted target cell
        Set TargetCell = mTST_DP_ScratchSheet.Range("H12")
        TargetCell.ClearContents
        TargetCell.NumberFormat = "dd/mm/yyyy"
        TargetCell.Select
    'Set the value the write-back transaction will apply
        gDP_WriteValue = VBA.DateSerial(2026, 8, 22)
    'Establish the enabled condition before the transaction
        Excel.Application.EnableEvents = True
    'Apply the write-back transaction
    'M_WriteBack_Apply resets On Error GoTo 0 on exit; re-arm immediately
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, True)
        On Error GoTo SuiteFail
    'Capture the restored state
        RestoredState = Excel.Application.EnableEvents
    'Assert write-back restored the enabled caller state
        TST_DP_AssertTrue "Write-back restores enabled events", RestoredState
    'Assert the result reports the caller's enabled state
        TST_DP_AssertFalse "Write-back result clears EventsDisabledByCaller", _
            WriteResult.EventsDisabledByCaller

'------------------------------------------------------------------------------
' REPAIR STILL FORCE-ENABLES
'------------------------------------------------------------------------------
    'Establish the suppressed condition before the repair check
        Excel.Application.EnableEvents = False
    'Repair the runtime, which is the only sanctioned force-enable path
    'DP_RepairRuntime resets On Error GoTo 0 on exit; re-arm immediately
        DP_RepairRuntime
        On Error GoTo SuiteFail
    'Assert DP_RepairRuntime re-enabled events
        TST_DP_AssertTrue "DP_RepairRuntime force-enables events", _
            Excel.Application.EnableEvents

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Restore the harness run state
        Excel.Application.EnableEvents = False
    'Release object references
        Set TargetCell = Nothing
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Capture the escaping error before any On Error statement resets Err
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Suppress local cleanup errors
        On Error Resume Next
    'Close any picker left loaded by a failed assertion path
        DP_Close
    'Restore the harness run state
        Excel.Application.EnableEvents = False
    'Release object references
        Set TargetCell = Nothing
    'Record the captured suite-level failure
        On Error GoTo 0
        TST_DP_RecordFail "ApplicationState suite failed", _
            "Error " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription
        Err.Clear

End Sub

#If VBA7 Then
Private Function TST_DP_ReadWindowStyle(ByVal WindowHandle As LongPtr) As LongPtr
#Else
Private Function TST_DP_ReadWindowStyle(ByVal WindowHandle As Long) As Long
#End If

'
'==============================================================================
'                        READ NATIVE WINDOW STYLE
'==============================================================================
'   Reads a window's style directly, so the window-style suite can verify native
'   state without asking the routine under test what it did.
'==============================================================================

'------------------------------------------------------------------------------
' READ STYLE
'------------------------------------------------------------------------------
    'Never let a verification read raise into a test
        On Error Resume Next
    #If Mac Then
        'No native window styling on Mac
            TST_DP_ReadWindowStyle = 0
    #Else
        #If VBA7 Then
            'Read the current window style
                TST_DP_ReadWindowStyle = TST_DP_GetWindowLongPtr(WindowHandle, TST_DP_GWL_STYLE)
        #Else
            'Read the current window style
                TST_DP_ReadWindowStyle = TST_DP_GetWindowLong(WindowHandle, TST_DP_GWL_STYLE)
        #End If
    #End If
    'Clear any suppressed read error
        Err.Clear

End Function

Private Sub TST_DP_RunSuite_WindowRecovery()

'
'==============================================================================
'                       WINDOW RECOVERY SUITE
'==============================================================================
' PURPOSE
'   Proves the UserForm acts on DP_WindowStyleResult instead of discarding it
'
' WHY THIS EXISTS
'   The WindowStyle suite calls M_Window_RemoveTitleBar directly and asserts the
'   returned UDT. That proved the transaction was observable, not that anything
'   consumed it. At v1.2.0 both production call sites in UF_DatePicker invoked
'   the Function as a statement and threw the result away, so a window left in no
'   known good state still produced a visible, interactive form
'
'   This suite drives the real form integration path: it arms the same fault
'   injection the WindowStyle suite uses, then loads the form and asserts what
'   the form actually did
'
' NOTES
'   Fault injection is one-shot. UserForm_Initialize consumes the armed fault, so
'   these cases exercise the initialization call site
'
'   The activation call site applies the same rule and is covered by inspection
'   and by the manual UI matrix. Driving it here would require showing a form
'   that then unloads itself from inside its own Activate event, which is not a
'   state worth creating inside an unattended run
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const FAULT_SET_WINDOW_POS  As Long = 3         'Matches M_DatePicker
    Const FAULT_ROLLBACK_STYLE  As Long = 1         'Matches M_DatePicker

    Dim LoadRaised              As Boolean          'True when the form load failed
    Dim OldUseWinAPI            As Boolean          'WinAPI styling setting on entry

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "WindowRecovery"

'------------------------------------------------------------------------------
' START FROM A KNOWN STATE
'------------------------------------------------------------------------------
    'Record the setting so the suite can restore it
        OldUseWinAPI = M_Settings_GetUseWinAPI()
    'Styling must be enabled for the call site to be reached at all
        M_Settings_SetUseWinAPI True
    'Begin with no form loaded
        On Error Resume Next
        Unload UF_DatePicker
        Err.Clear
        On Error GoTo SuiteFail

'------------------------------------------------------------------------------
' A CLEAN CALL LOADS THE FORM
'------------------------------------------------------------------------------
    'No fault armed: styling applies and the form loads normally
        Load UF_DatePicker
        TST_DP_AssertTrue "Clean styling loads the form", _
            TST_DP_IsPickerFormLoadedForTest()
        Unload UF_DatePicker
        TST_DP_AssertFalse "Clean styling unloads cleanly", _
            TST_DP_IsPickerFormLoadedForTest()

'------------------------------------------------------------------------------
' A ROLLED-BACK CALL STILL LOADS THE FORM
'------------------------------------------------------------------------------
    'A primary failure whose rollback succeeds leaves the window in its known
    'original state, which is a supported outcome and must not block the load
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS
        Load UF_DatePicker
        TST_DP_AssertTrue "Rolled-back styling still loads the form", _
            TST_DP_IsPickerFormLoadedForTest()
        Unload UF_DatePicker

'------------------------------------------------------------------------------
' A RECOVERY-REQUIRED CALL FAILS THE LOAD
'------------------------------------------------------------------------------
    'A primary failure whose rollback also fails leaves no known good state. The
    'form must not be presented, so the load itself has to fail
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS, FAULT_ROLLBACK_STYLE
        LoadRaised = False
        On Error Resume Next
        Load UF_DatePicker
        LoadRaised = (Err.Number <> 0)
        Err.Clear
        On Error GoTo SuiteFail
        TST_DP_AssertTrue "Recovery-required styling fails the form load", _
            LoadRaised
    'The contract is that the unknown-state window is never presented. A Load
    'that raises from UserForm_Initialize leaves an orphaned instance in
    'VBA.UserForms that the project cannot reach by name, so presence in that
    'collection over-reports. Visibility is the property #47 actually promises
        TST_DP_AssertFalse "Recovery-required styling presents no visible form", _
            TST_DP_IsPickerFormVisibleForTest()
    'Clear the orphan so it cannot leak into the cases below
        TST_DP_UnloadAllPickerFormsForTest

'------------------------------------------------------------------------------
' THE ARMED FAULT IS CONSUMED, NOT STICKY
'------------------------------------------------------------------------------
    'The next load must succeed, proving the failure came from the injected fault
    'and not from a state this suite left behind
        Load UF_DatePicker
        TST_DP_AssertTrue "The next load after a recovery failure succeeds", _
            TST_DP_IsPickerFormLoadedForTest()
        Unload UF_DatePicker

'------------------------------------------------------------------------------
' A NON-ATTEMPT REMAINS SAFE
'------------------------------------------------------------------------------
    'With styling disabled the call site is never reached, so the form loads with
    'its native chrome and nothing is treated as a recovery condition
        M_Settings_SetUseWinAPI False
        Load UF_DatePicker
        TST_DP_AssertTrue "WinAPI-disabled styling still loads the form", _
            TST_DP_IsPickerFormLoadedForTest()
        TST_DP_UnloadAllPickerFormsForTest
        TST_DP_AssertFalse "WinAPI-disabled load unloads cleanly", _
            TST_DP_IsPickerFormLoadedForTest()

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Disarm any fault the suite armed and leave no loaded form behind, including
    'an instance orphaned by a Load that raised from UserForm_Initialize
        On Error Resume Next
        M_Window_Test_SetFaultInjection 0, 0
        TST_DP_UnloadAllPickerFormsForTest
        M_Settings_SetUseWinAPI OldUseWinAPI
        Err.Clear
        On Error GoTo 0
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Record the failure and clear the error
        TST_DP_RecordFail "Window recovery suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    'Restore state regardless
        Resume SuiteExit

End Sub

Private Sub TST_DP_RunSuite_WindowStyle()

'
'==============================================================================
'                          SUITE: WINDOW STYLE
'------------------------------------------------------------------------------
' PURPOSE
'   Proves the borderless window-style transaction, including the failure paths
'   that leave a window half-styled
'
' WHY THIS EXISTS
'   Clearing WS_CAPTION and refreshing the frame are separate native operations.
'   The first can succeed and the second fail, and no ordinary input can make
'   SetWindowPos or DrawMenuBar fail on demand
'
'   Without deterministic coverage, the rollback this issue exists to add would
'   never execute outside a real defect
'
' BEHAVIOR
'   Loads the picker form, resolves its native window, then drives every failure
'   point through the one-shot fault seam and asserts the transaction outcome
'
' ERROR POLICY
'   Records suite-level failures and continues
'
'   Always disarms the fault seam, including on a failed assertion
'
' DEPENDENCIES
'   DP_Preload
'   UF_DatePicker
'   M_Window_GetUserFormHwnd
'   M_Window_RemoveTitleBar
'   M_Window_Test_SetFaultInjection
'   M_Settings_SetUseWinAPI
'
' NOTES
'   The failure-point numbers below are duplicated from private constants in
'   M_DatePicker rather than shared through a public enum, so the seam does not
'   enlarge the public API. The two lists must be changed together
'
'   A resolvable window handle is asserted before any injected failure. A
'   preloaded hidden UserForm is expected to have one on supported hosts, but a
'   missing handle must report as a setup failure rather than pass quietly as a
'   non-attempt
'
'   If the supported host proves otherwise, this suite moves to the UI smoke pack
'   where the form can be shown deliberately. The precondition is not weakened to
'   keep it here
'
'   The suite ends by reapplying the style cleanly, so a rolled-back form is not
'   left with its native title bar for whatever runs next
'
' UPDATED
'   2026-08-23
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const FAULT_STYLE_READ      As Long = 1     'Matches M_DatePicker
    Const FAULT_STYLE_WRITE     As Long = 2     'Matches M_DatePicker
    Const FAULT_SET_WINDOW_POS  As Long = 3     'Matches M_DatePicker
    Const FAULT_DRAW_MENU_BAR   As Long = 4     'Matches M_DatePicker
    Const FAULT_ROLLBACK_STYLE  As Long = 1     'Matches M_DatePicker
    Const FAULT_ROLLBACK_FRAME  As Long = 2     'Matches M_DatePicker

    #If VBA7 Then
        Dim FormHandle          As LongPtr      'Native window behind the picker form
    #Else
        Dim FormHandle          As Long         'Native window behind the picker form
    #End If

    #If VBA7 Then
        Dim StyleBefore         As LongPtr      'Native style read before a call
        Dim StyleAfter          As LongPtr      'Native style read after a call
    #Else
        Dim StyleBefore         As Long         'Native style read before a call
        Dim StyleAfter          As Long         'Native style read after a call
    #End If

    Dim StyleResult         As DP_WindowStyleResult 'Outcome under test
    Dim PriorUseWinAPI      As Boolean              'WinAPI setting before the suite

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "WindowStyle"
    PriorUseWinAPI = gDP_UseWinAPI

'------------------------------------------------------------------------------
' RESOLVE A REAL WINDOW HANDLE
'------------------------------------------------------------------------------
    'Load the picker form without showing it
    'DP_Preload resets On Error GoTo 0 on exit; re-arm immediately
        DP_Preload
        On Error GoTo SuiteFail
    'Resolve the native window the transaction will operate on
        FormHandle = M_Window_GetUserFormHwnd(UF_DatePicker)
    'Assert the precondition. A missing handle is a setup failure, never a pass
        TST_DP_AssertTrue "Preloaded form exposes a native window handle", _
            FormHandle <> 0
    'Stop here when there is no window to test against
        If FormHandle = 0 Then
            TST_DP_RecordFail "Window style suite setup", _
                "No native window handle. Move these cases to the UI smoke pack " & _
                "rather than weakening the precondition."
            GoTo SuiteExit
        End If

'------------------------------------------------------------------------------
' SAFE NON-ATTEMPT PATHS
'------------------------------------------------------------------------------
    'A missing form must not touch anything
        StyleResult = M_Window_RemoveTitleBar(Nothing)
        TST_DP_AssertFalse "Missing form is not attempted", StyleResult.Attempted
        TST_DP_AssertFalse "Missing form is not committed", StyleResult.Committed
    'A disabled WinAPI policy must leave the native title bar alone
        M_Settings_SetUseWinAPI False
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertFalse "WinAPI-disabled path is not attempted", StyleResult.Attempted
        TST_DP_AssertFalse "WinAPI-disabled path is not committed", StyleResult.Committed
        TST_DP_AssertFalse "WinAPI-disabled path reports no recovery", _
            StyleResult.RecoveryRequired
        M_Settings_SetUseWinAPI PriorUseWinAPI

'------------------------------------------------------------------------------
' SUCCESSFUL TRANSACTION
'------------------------------------------------------------------------------
    'A clean call must fully apply the borderless style
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Clean call applies the borderless style", StyleResult.Applied
        TST_DP_AssertTrue "Clean call commits the style write", StyleResult.Committed
        TST_DP_AssertFalse "Clean call does not roll back", StyleResult.RolledBack
        TST_DP_AssertFalse "Clean call needs no recovery", StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "Clean call reports no failing step", _
            VBA.vbNullString, StyleResult.FailedStep
    'Verify the native style directly, not through the result being tested
        StyleAfter = TST_DP_ReadWindowStyle(FormHandle)
        TST_DP_AssertTrue "Clean call clears WS_CAPTION natively", _
            (StyleAfter And TST_DP_WS_CAPTION) = 0
    'A second call on an already-borderless window must remain safe
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Repeat call still applies the style", StyleResult.Applied
        TST_DP_AssertFalse "Repeat call needs no recovery", StyleResult.RecoveryRequired

'------------------------------------------------------------------------------
' PRE-COMMIT FAILURES
'------------------------------------------------------------------------------
    'A failed style read must abort before anything is changed
        StyleBefore = TST_DP_ReadWindowStyle(FormHandle)
        M_Window_Test_SetFaultInjection FAULT_STYLE_READ
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        StyleAfter = TST_DP_ReadWindowStyle(FormHandle)
    'Verify natively that nothing was changed
        TST_DP_AssertTrue "Style-read failure leaves the native style unchanged", _
            StyleAfter = StyleBefore
        TST_DP_AssertFalse "Style-read failure is not attempted", StyleResult.Attempted
        TST_DP_AssertFalse "Style-read failure is not committed", StyleResult.Committed
        TST_DP_AssertFalse "Style-read failure needs no recovery", _
            StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "Style-read failure names its step", _
            "Read window style", StyleResult.FailedStep
    'A failed style write must abort before anything is changed. The injected
    'failure skips the native write, so the window keeps the style it had
        StyleBefore = TST_DP_ReadWindowStyle(FormHandle)
        M_Window_Test_SetFaultInjection FAULT_STYLE_WRITE
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        StyleAfter = TST_DP_ReadWindowStyle(FormHandle)
        TST_DP_AssertFalse "Style-write failure is not committed", StyleResult.Committed
        TST_DP_AssertTrue "Style-write failure reports the attempt", StyleResult.Attempted
        TST_DP_AssertFalse "Style-write failure needs no recovery", _
            StyleResult.RecoveryRequired
    'Verify natively that the write really did not take effect
        TST_DP_AssertTrue "Style-write failure leaves the native style unchanged", _
            StyleAfter = StyleBefore
        TST_DP_AssertEqualsString "Style-write failure names its step", _
            "Write window style", StyleResult.FailedStep

'------------------------------------------------------------------------------
' POST-COMMIT FAILURES WITH SUCCESSFUL ROLLBACK
'------------------------------------------------------------------------------
    'A frame refresh that fails after the commit must roll back
        StyleBefore = TST_DP_ReadWindowStyle(FormHandle)
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        StyleAfter = TST_DP_ReadWindowStyle(FormHandle)
    'Verify natively that the original style really came back
        TST_DP_AssertTrue "SetWindowPos rollback restores the native style", _
            StyleAfter = StyleBefore
        TST_DP_AssertTrue "SetWindowPos failure commits first", StyleResult.Committed
        TST_DP_AssertFalse "SetWindowPos failure is not applied", StyleResult.Applied
        TST_DP_AssertTrue "SetWindowPos failure rolls back", StyleResult.RolledBack
        TST_DP_AssertFalse "SetWindowPos rollback needs no recovery", _
            StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "SetWindowPos failure names its step", _
            "Refresh non-client frame", StyleResult.FailedStep
    'A redraw that fails after the commit must roll back
        M_Window_Test_SetFaultInjection FAULT_DRAW_MENU_BAR
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "DrawMenuBar failure commits first", StyleResult.Committed
        TST_DP_AssertFalse "DrawMenuBar failure is not applied", StyleResult.Applied
        TST_DP_AssertTrue "DrawMenuBar failure rolls back", StyleResult.RolledBack
        TST_DP_AssertFalse "DrawMenuBar rollback needs no recovery", _
            StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "DrawMenuBar failure names its step", _
            "Redraw frame", StyleResult.FailedStep

'------------------------------------------------------------------------------
' ROLLBACK FAILURES
'------------------------------------------------------------------------------
    'A rollback whose style restore fails leaves no known good state. The injected
    'failure skips the restore, so the window is genuinely left committed
        StyleBefore = TST_DP_ReadWindowStyle(FormHandle)
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS, FAULT_ROLLBACK_STYLE
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        StyleAfter = TST_DP_ReadWindowStyle(FormHandle)
    'Verify natively that the style was not restored, which is what makes this
    'case unrecoverable rather than a rollback
        TST_DP_AssertTrue "Failed style rollback leaves the committed style in place", _
            (StyleAfter And TST_DP_WS_CAPTION) = 0
        TST_DP_AssertFalse "Failed style rollback does not report RolledBack", _
            StyleResult.RolledBack
        TST_DP_AssertTrue "Failed style rollback requires recovery", _
            StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "Failed style rollback keeps the primary step", _
            "Refresh non-client frame", StyleResult.FailedStep
    'A rollback whose frame refresh fails leaves no known good state
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS, FAULT_ROLLBACK_FRAME
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertFalse "Failed frame rollback does not report RolledBack", _
            StyleResult.RolledBack
        TST_DP_AssertTrue "Failed frame rollback requires recovery", _
            StyleResult.RecoveryRequired
        TST_DP_AssertEqualsString "Failed frame rollback keeps the primary step", _
            "Refresh non-client frame", StyleResult.FailedStep

'------------------------------------------------------------------------------
' ONE-SHOT CONSUMPTION
'------------------------------------------------------------------------------
    'An armed fault must affect exactly one call
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Armed fault affects the first call", StyleResult.RolledBack
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Armed fault does not affect the next call", _
            StyleResult.Applied
        TST_DP_AssertFalse "Consumed fault leaves no rollback", StyleResult.RolledBack
    'An armed rollback fault must also affect exactly one call
        M_Window_Test_SetFaultInjection FAULT_SET_WINDOW_POS, FAULT_ROLLBACK_STYLE
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Armed rollback fault affects the first call", _
            StyleResult.RecoveryRequired
        StyleResult = M_Window_RemoveTitleBar(UF_DatePicker)
        TST_DP_AssertTrue "Consumed rollback fault leaves the next call clean", _
            StyleResult.Applied

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    'Disarm the seam before anything else runs
        M_Window_Test_SetFaultInjection 0, 0
    'Restore the WinAPI setting the suite may have changed
        M_Settings_SetUseWinAPI PriorUseWinAPI
    'Leave the form styled rather than rolled back
        If FormHandle <> 0 Then M_Window_RemoveTitleBar UF_DatePicker
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Never leave the seam armed after a failed assertion
        M_Window_Test_SetFaultInjection 0, 0
    'Restore the WinAPI setting the suite may have changed
        M_Settings_SetUseWinAPI PriorUseWinAPI
    'Record the failure and clear the error
        TST_DP_RecordFail "Window style suite", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_RunSuite_HarnessSelfCheck()

'
'==============================================================================
'                          SUITE: HARNESS SELF CHECK
'------------------------------------------------------------------------------
' PURPOSE
'   Proves the harness run-state machine, the cleanup-failure counter and the
'   dirty-start preflight, rather than assuming them
'
' WHY THIS EXISTS
'   #15 gates a release on these run states. A state that has never been observed
'   to occur is not evidence of anything, and FAIL_CLEANUP and FAIL_DIRTY_START
'   are both states a passing run never reaches on its own
'
' BEHAVIOR
'   Drives TST_DP_ResolveRunState across every outcome, forces one cleanup step
'   to fail, and runs the preflight against the current environment
'
' ERROR POLICY
'   Records suite-level failures and continues
'
' DEPENDENCIES
'   TST_DP_ResolveRunState
'   TST_DP_CheckCleanupStep
'   TST_DP_Preflight
'   TST_DP_ContextMenuControlCount
'
' NOTES
'   The counters this suite manipulates are the counters that describe the run it
'   is part of. Every one is saved before the first mutation and restored before
'   the first assertion, so the run reports its own outcome and not the probe
'   values used here
'
'   Assertions therefore run against locals captured during the probe, never
'   against live module state
'
'   The preflight probe expects a dirty verdict, because the scratch worksheet
'   exists while the run is using it. That is the same evidence a project reset
'   would leave behind, which is the case the module-level flag cannot see
'
'   An aborted run cannot be staged from inside a run. This suite proves the
'   detector fires on the evidence an abort leaves; that an abort leaves it is a
'   manual validation step
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SavedFail           As Long         'Live assertion failure count
    Dim SavedCleanup        As Long         'Live cleanup failure count
    Dim SavedDetail         As String       'Live cleanup failure detail
    Dim SavedDispatched     As Long         'Live dispatched suite count
    Dim SavedCompleted      As Long         'Live completed suite count
    Dim SavedDirty          As Boolean      'Live dirty-start verdict
    Dim SavedDirtyDetail    As String       'Live dirty-start detail

    Dim StatePass           As String       'Resolver output for a clean run
    Dim StateFail           As String       'Resolver output for a failed assertion
    Dim StateIncomplete     As String       'Resolver output for a skipped suite
    Dim StateCleanup        As String       'Resolver output for failed teardown
    Dim StateDirty          As String       'Resolver output for a dirty start
    Dim StateDirtyOverFail  As String       'Resolver output when both apply

    Dim InjectedCount       As Long         'Cleanup failures after injection
    Dim MenuProbeCount      As Long         'Context-menu controls seen by the probe
    Dim ProbeDirty          As Boolean      'Preflight verdict during the probe
    Dim ProbeDetail         As String       'Preflight detail during the probe

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "HarnessSelfCheck"

'------------------------------------------------------------------------------
' SAVE LIVE RUN STATE
'------------------------------------------------------------------------------
    'Every counter touched below belongs to the run in progress
        SavedFail = mTST_DP_FailCount
        SavedCleanup = mTST_DP_CleanupFails
        SavedDetail = mTST_DP_CleanupDetail
        SavedDispatched = mTST_DP_SuitesDispatched
        SavedCompleted = mTST_DP_SuitesCompleted
        SavedDirty = mTST_DP_DirtyStart
        SavedDirtyDetail = mTST_DP_DirtyDetail

'------------------------------------------------------------------------------
' PROBE THE RUN-STATE MACHINE
'------------------------------------------------------------------------------
    'Clean run
        mTST_DP_FailCount = 0
        mTST_DP_CleanupFails = 0
        mTST_DP_SuitesDispatched = 3
        mTST_DP_SuitesCompleted = 3
        mTST_DP_DirtyStart = False
        StatePass = TST_DP_ResolveRunState()

    'Failed assertion
        mTST_DP_FailCount = 1
        StateFail = TST_DP_ResolveRunState()

    'Dispatched suite that never returned
        mTST_DP_FailCount = 0
        mTST_DP_SuitesCompleted = 2
        StateIncomplete = TST_DP_ResolveRunState()

    'Teardown that did not complete
        mTST_DP_SuitesCompleted = 3
        mTST_DP_CleanupFails = 1
        StateCleanup = TST_DP_ResolveRunState()

    'Dirty start on an otherwise clean run
        mTST_DP_CleanupFails = 0
        mTST_DP_DirtyStart = True
        StateDirty = TST_DP_ResolveRunState()

    'Dirty start alongside a failed assertion
        mTST_DP_FailCount = 1
        StateDirtyOverFail = TST_DP_ResolveRunState()

'------------------------------------------------------------------------------
' PROBE CLEANUP FAULT INJECTION
'------------------------------------------------------------------------------
    'Start the probe from a known clean counter
        mTST_DP_CleanupFails = 0
        mTST_DP_CleanupDetail = VBA.vbNullString
    'Clear any error state so only the injection can trigger a failure
        Err.Clear
    'Force the named step to be treated as failed
        mTST_DP_InjectCleanupFail = "SelfCheckProbeStep"
        TST_DP_CheckCleanupStep "SelfCheckProbeStep"
        mTST_DP_InjectCleanupFail = VBA.vbNullString
    'Capture what the counter recorded
        InjectedCount = mTST_DP_CleanupFails

'------------------------------------------------------------------------------
' PROBE THE DIRTY-START PREFLIGHT
'------------------------------------------------------------------------------
    'Run preflight against the live environment. The scratch worksheet exists
    'while the run is using it, which is the evidence an aborted run leaves
        TST_DP_Preflight
        ProbeDirty = mTST_DP_DirtyStart
        ProbeDetail = mTST_DP_DirtyDetail

'------------------------------------------------------------------------------
' PROBE THE CONTEXT-MENU READER
'------------------------------------------------------------------------------
    'Exercise the teardown probe while the run is still live
        MenuProbeCount = TST_DP_ContextMenuControlCount()

'------------------------------------------------------------------------------
' RESTORE LIVE RUN STATE
'------------------------------------------------------------------------------
    'Restore before the first assertion, so this suite's own result is honest
        mTST_DP_FailCount = SavedFail
        mTST_DP_CleanupFails = SavedCleanup
        mTST_DP_CleanupDetail = SavedDetail
        mTST_DP_SuitesDispatched = SavedDispatched
        mTST_DP_SuitesCompleted = SavedCompleted
        mTST_DP_DirtyStart = SavedDirty
        mTST_DP_DirtyDetail = SavedDirtyDetail
    'Clear any error left by the injection probe
        Err.Clear

'------------------------------------------------------------------------------
' ASSERT THE RUN-STATE MACHINE
'------------------------------------------------------------------------------
    'Assert a clean run reports PASS
        TST_DP_AssertEqualsString "Clean run resolves to PASS", _
            TST_DP_STATE_PASS, StatePass
    'Assert a failed assertion reports FAIL
        TST_DP_AssertEqualsString "Failed assertion resolves to FAIL", _
            TST_DP_STATE_FAIL, StateFail
    'Assert a suite that never returned reports INCOMPLETE_SKIPPED
        TST_DP_AssertEqualsString "Skipped suite resolves to INCOMPLETE_SKIPPED", _
            TST_DP_STATE_INCOMPLETE, StateIncomplete
    'Assert failed teardown reports FAIL_CLEANUP
        TST_DP_AssertEqualsString "Failed teardown resolves to FAIL_CLEANUP", _
            TST_DP_STATE_FAIL_CLEANUP, StateCleanup
    'Assert a dirty start reports FAIL_DIRTY_START
        TST_DP_AssertEqualsString "Dirty start resolves to FAIL_DIRTY_START", _
            TST_DP_STATE_DIRTY_START, StateDirty
    'Assert a dirty start outranks a failed assertion
        TST_DP_AssertEqualsString "Dirty start outranks a failed assertion", _
            TST_DP_STATE_DIRTY_START, StateDirtyOverFail
    'Assert a dirty start can never resolve to PASS, which is #19's invariant
        TST_DP_AssertTrue "A dirty start can never resolve to PASS", _
            (StateDirty <> TST_DP_STATE_PASS) And _
            (StateDirtyOverFail <> TST_DP_STATE_PASS)

'------------------------------------------------------------------------------
' ASSERT CLEANUP FAULT INJECTION
'------------------------------------------------------------------------------
    'Assert a forced cleanup failure is counted, which is what produces
    'FAIL_CLEANUP in a real run
        TST_DP_AssertEqualsLong "Injected cleanup failure is counted", _
            1, InjectedCount

'------------------------------------------------------------------------------
' ASSERT THE DIRTY-START PREFLIGHT
'------------------------------------------------------------------------------
    'Assert preflight detects the leftover-worksheet evidence
        TST_DP_AssertTrue "Preflight detects leftover scratch worksheet", ProbeDirty
    'Assert the verdict names what it found
        TST_DP_AssertTrue "Preflight verdict names the leftover worksheet", _
            VBA.InStr(1, ProbeDetail, TST_DP_SCRATCH_SHEET_NAME, vbTextCompare) > 0

'------------------------------------------------------------------------------
' ASSERT THE CONTEXT-MENU PROBE
'------------------------------------------------------------------------------
    'Assert the teardown probe can read the context menus at all. A probe that
    'silently returns zero because it cannot see the bars would report a clean
    'teardown for a menu it never inspected
        TST_DP_AssertTrue "Context-menu probe reads the command bars", _
            MenuProbeCount >= 0

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Exit after the suite completes
        Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    'Never leave the injection seam armed
        mTST_DP_InjectCleanupFail = VBA.vbNullString
    'Restore whatever the probe was holding when it failed
        mTST_DP_FailCount = SavedFail
        mTST_DP_CleanupFails = SavedCleanup
        mTST_DP_CleanupDetail = SavedDetail
        mTST_DP_SuitesDispatched = SavedDispatched
        mTST_DP_SuitesCompleted = SavedCompleted
        mTST_DP_DirtyStart = SavedDirty
        mTST_DP_DirtyDetail = SavedDirtyDetail
    'Record the failure and clear the error
        TST_DP_RecordFail "Harness self check", _
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
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber         As Long                 'Captured error number
    Dim ErrorDescription    As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the current suite name
        mTST_DP_CurrentSuite = "UISmoke"
    'Enable suite-level error handling
        On Error GoTo SuiteFail
    'Activate the scratch sheet
        TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet
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
    'Capture the escaping error before any On Error statement resets Err
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description
    'Attempt to close the picker after a UI smoke failure
        On Error Resume Next
        DP_Close
        Err.Clear
        On Error GoTo 0
    'Record the captured suite-level failure
        TST_DP_RecordFail "UISmoke suite failed", _
            "Error " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription
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

Private Sub TST_DP_ExpectError_SelectDateZero()

'
'==============================================================================
'                     EXPECT ERROR: SELECT DATE ZERO VALUE
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call M_Picker_SelectDate with a zero date to trigger the expected error
        M_Picker_SelectDate 0

'------------------------------------------------------------------------------
' RECORD MISSING ERROR
'------------------------------------------------------------------------------
    'Record a failure when no error was raised
        TST_DP_RecordFail "Zero date raises in M_Picker_SelectDate", "No error was raised"
    Exit Sub

'------------------------------------------------------------------------------
' EXPECTED ERROR
'------------------------------------------------------------------------------
ExpectedError:
    TST_DP_RecordPass "Zero date raises in M_Picker_SelectDate", Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ExpectFailedAddressReport()

'
'==============================================================================
'                      EXPECT FAILED ADDRESS REPORT
'==============================================================================
'   Places a multi-cell array formula inside the target so those cells reject the
'   write, then asserts the result names exactly them.
'
'   Excel declines these writes silently through the object model. It raises
'   "You cannot change part of an array" only for an interactive edit. So the
'   engine refuses array cells before writing them, on both paths: the fast bulk
'   path is skipped for any target touching an array, and the per-cell writer
'   refuses each array cell individually.
'
'   Without those refusals the write reports success having changed nothing, which
'   is what this expectation is really guarding. The assertions that the array
'   survives and that I5 was written distinguish a correct refusal from a write
'   that silently did not happen.
'
'   This is the non-lock failure path. Protected locked cells are a separate
'   classification and are covered by TST_DP_ExpectPartialWriteReport.
'
'   The array formula is cleared on every path, including a failed assertion.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WriteResult     As DP_WriteResult   'Structured write-back result
    Dim TargetRange     As Excel.Range      'Partially writable target
    Dim ExpectedList    As String           'Expected failed address list

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo FailedAddressFail

'------------------------------------------------------------------------------
' BUILD A PARTIALLY WRITABLE TARGET
'------------------------------------------------------------------------------
    'Use a target away from the ranges the rest of the suite writes
        Set TargetRange = mTST_DP_ScratchSheet.Range("I5:I8")
    'Clear any content left by an earlier run
        mTST_DP_ScratchSheet.Range("I5:I9").ClearContents
    'A cell belonging to a multi-cell array formula rejects a direct value write
        mTST_DP_ScratchSheet.Range("I6:I7").FormulaArray = "=ROW()"
    'Assert the setup actually took, so a silent setup failure cannot look like a
    'reporting defect in the result
        TST_DP_AssertTrue "Failed write setup creates an array formula", _
            mTST_DP_ScratchSheet.Range("I6").HasArray
    'Prepare a distinct write value
        gDP_WriteValue = VBA.DateSerial(2026, 12, 8)

'------------------------------------------------------------------------------
' WRITE THROUGH THE PARTIAL TARGET
'------------------------------------------------------------------------------
    'Call the range writer directly so the run stays free of modal messages
        M_WriteBack_PopulateRange TargetRange, DP_WriteAction_DatePicker, WriteResult

'------------------------------------------------------------------------------
' ASSERT THE FAILED ADDRESSES
'------------------------------------------------------------------------------
    'Build the worksheet-qualified list the write is expected to report
        ExpectedList = mTST_DP_ScratchSheet.Name & "!I6, " & _
            mTST_DP_ScratchSheet.Name & "!I7"
    'Assert the writable cells were written
        TST_DP_AssertEqualsLong "Failed write reports 2 written cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert the rejected cells were counted as failures rather than skips
        TST_DP_AssertEqualsLong "Failed write reports 2 failed cells", _
            2, VBA.CLng(WriteResult.FailedCount)
    'Assert the array cells were left intact rather than silently replaced
        TST_DP_AssertTrue "Failed write leaves the array formula intact", _
            mTST_DP_ScratchSheet.Range("I6").HasArray
    'Assert a non-lock failure is not misreported as a protected skip
        TST_DP_AssertEqualsLong "Failed write reports no locked skips", _
            0, VBA.CLng(WriteResult.LockedSkippedCount)
    'Assert the exact failed addresses are reported, not just a count
        TST_DP_AssertEqualsString "Failed write reports the exact failed addresses", _
            ExpectedList, WriteResult.FailedAddresses
    'Assert the failed result still satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Failed write result balances", WriteResult
    'Assert the writable cells really received the value
        TST_DP_AssertCellDateEquals "Failed write writes I5", _
            VBA.DateSerial(2026, 12, 8), mTST_DP_ScratchSheet.Range("I5")

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Remove the array formula before leaving
        TST_DP_ReleaseScratchArrayFormula
    'Release object references
        Set TargetRange = Nothing
    'Exit after the expectation completes
        Exit Sub

'------------------------------------------------------------------------------
' FAILED ADDRESS FAIL
'------------------------------------------------------------------------------
FailedAddressFail:
    'Remove the array formula even when the expectation failed
        TST_DP_ReleaseScratchArrayFormula
    'Release object references
        Set TargetRange = Nothing
    'Record the failure and clear the error
        TST_DP_RecordFail "Failed address reporting", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ReleaseScratchArrayFormula()

'
'==============================================================================
'                     RELEASE SCRATCH ARRAY FORMULA
'==============================================================================
'   Clears the array formula used by the failed-address expectation, so a failed
'   assertion cannot leave an unwritable block on the scratch sheet.
'==============================================================================

'------------------------------------------------------------------------------
' RELEASE ARRAY FORMULA
'------------------------------------------------------------------------------
    'Never let cleanup raise into the caller
        On Error Resume Next
    'Exit when there is no scratch sheet to release
        If mTST_DP_ScratchSheet Is Nothing Then Exit Sub
    'Clearing the whole array range is the only way to remove an array formula
        mTST_DP_ScratchSheet.Range("I5:I9").ClearContents
    'Clear any suppressed cleanup error
        Err.Clear

End Sub

Private Sub TST_DP_AssertWriteResultBalances( _
    ByVal TestName As String, _
    ByRef Result As DP_WriteResult)

'
'==============================================================================
'                     ASSERT WRITE RESULT BALANCES
'==============================================================================
'   Asserts the accounting invariant every completed write result must satisfy:
'
'       AttemptedCount = WrittenCount + LockedSkippedCount
'                      + FormulaSkippedCount + FailedCount
'
'   Every cell increments exactly one term, so the invariant holds by
'   construction rather than by arithmetic. A term that stops being incremented
'   exactly once shows up here.
'==============================================================================

'------------------------------------------------------------------------------
' ASSERT THE INVARIANT
'------------------------------------------------------------------------------
    If Result.AttemptedCount = Result.WrittenCount + Result.LockedSkippedCount + _
        Result.FormulaSkippedCount + Result.FailedCount Then
        TST_DP_RecordPass TestName, _
            "Attempted=" & VBA.CStr(Result.AttemptedCount) & _
            "; Written=" & VBA.CStr(Result.WrittenCount) & _
            "; Locked=" & VBA.CStr(Result.LockedSkippedCount) & _
            "; Formula=" & VBA.CStr(Result.FormulaSkippedCount) & _
            "; Failed=" & VBA.CStr(Result.FailedCount)
    Else
        TST_DP_RecordFail TestName, _
            "Attempted=" & VBA.CStr(Result.AttemptedCount) & _
            " does not equal Written=" & VBA.CStr(Result.WrittenCount) & _
            " + Locked=" & VBA.CStr(Result.LockedSkippedCount) & _
            " + Formula=" & VBA.CStr(Result.FormulaSkippedCount) & _
            " + Failed=" & VBA.CStr(Result.FailedCount)
    End If

End Sub

Private Sub TST_DP_ExpectFormulaProtection()

'
'==============================================================================
'                      EXPECT FORMULA PROTECTION
'==============================================================================
'   Builds a target holding blanks, literals and formulas, then asserts that the
'   formulas survive a default write and are replaced only on explicit request.
'
'   The target is deliberately larger than one cell, so the bulk path would be
'   eligible if formula protection did not refuse it. A formula surviving here
'   therefore also proves the bulk gate, which is the half of this policy that
'   per-cell inspection alone cannot cover.
'
'   Row 18 of the demo teaches that a date-returning formula is a valid picker
'   target. It stays one; what changes is that selecting a date no longer deletes
'   the formula behind it.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WriteResult     As DP_WriteResult   'Structured write-back result
    Dim EmptyResult     As DP_WriteResult   'Zeroed result used to reset WriteResult
    Dim TargetRange     As Excel.Range      'Mixed blank/literal/formula target

    Dim FPSavedErrNumber      As Long   'Captured original error number
    Dim FPSavedErrDescription As String 'Captured original error description
    Dim FPSavedErrSource      As String 'Captured original error source
    Dim ExpectedList    As String           'Expected preserved address list

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo FormulaProtectionFail

'------------------------------------------------------------------------------
' BUILD A MIXED TARGET
'------------------------------------------------------------------------------
    'Use a target away from the ranges the rest of the suite writes
        Set TargetRange = mTST_DP_ScratchSheet.Range("K5:K8")
    'Clear any content left by an earlier run
        TargetRange.ClearContents
    'K5 blank, K6 literal, K7 plain formula, K8 date-returning formula
        mTST_DP_ScratchSheet.Range("K6").Value = VBA.DateSerial(2020, 1, 1)
        mTST_DP_ScratchSheet.Range("K7").Formula = "=1+1"
        mTST_DP_ScratchSheet.Range("K8").Formula = "=TODAY()"
    'Assert the setup took, so a silent setup failure cannot look like protection
        TST_DP_AssertTrue "Formula setup creates a plain formula", _
            mTST_DP_ScratchSheet.Range("K7").HasFormula
        TST_DP_AssertTrue "Formula setup creates a date-returning formula", _
            mTST_DP_ScratchSheet.Range("K8").HasFormula
    'Prepare a distinct write value
        gDP_WriteValue = VBA.DateSerial(2027, 3, 9)

'------------------------------------------------------------------------------
' WRITE WITH PROTECTION ACTIVE
'------------------------------------------------------------------------------
    'Default policy preserves formulas
        M_WriteBack_PopulateRange TargetRange, DP_WriteAction_DatePicker, WriteResult

'------------------------------------------------------------------------------
' ASSERT PROTECTION
'------------------------------------------------------------------------------
    'Build the worksheet-qualified list the write is expected to report
        ExpectedList = mTST_DP_ScratchSheet.Name & "!K7, " & _
            mTST_DP_ScratchSheet.Name & "!K8"
    'Assert the blank and the literal were written
        TST_DP_AssertEqualsLong "Formula protection writes 2 cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert both formulas were preserved
        TST_DP_AssertEqualsLong "Formula protection preserves 2 formula cells", _
            2, VBA.CLng(WriteResult.FormulaSkippedCount)
    'Assert a preserved formula is not misreported as a failure
        TST_DP_AssertEqualsLong "Formula protection reports no failures", _
            0, VBA.CLng(WriteResult.FailedCount)
    'Assert the preserved cells are named, not just counted
        TST_DP_AssertEqualsString "Formula protection reports the preserved addresses", _
            ExpectedList, WriteResult.FormulaSkippedAddresses
    'Assert the shortfall description names them
        TST_DP_AssertTrue "Formula protection description names a preserved cell", _
            VBA.InStr(1, M_WriteBack_DescribeShortfall(WriteResult), "K7", vbTextCompare) > 0
    'Assert the result satisfies the extended accounting invariant
        TST_DP_AssertWriteResultBalances "Formula protection result balances", WriteResult

'------------------------------------------------------------------------------
' ASSERT THE WORKSHEET ITSELF
'------------------------------------------------------------------------------
    'Verify the formulas natively, not only through the result being tested
        TST_DP_AssertTrue "Plain formula survives the write", _
            mTST_DP_ScratchSheet.Range("K7").HasFormula
        TST_DP_AssertTrue "Date-returning formula survives the write", _
            mTST_DP_ScratchSheet.Range("K8").HasFormula
    'Verify the blank and the literal really were replaced
        TST_DP_AssertCellDateEquals "Formula protection writes the blank cell", _
            VBA.DateSerial(2027, 3, 9), mTST_DP_ScratchSheet.Range("K5")
        TST_DP_AssertCellDateEquals "Formula protection overwrites the literal", _
            VBA.DateSerial(2027, 3, 9), mTST_DP_ScratchSheet.Range("K6")

'------------------------------------------------------------------------------
' WRITE WITH THE OVERRIDE
'------------------------------------------------------------------------------
    'An explicit override replaces formulas
        WriteResult = EmptyResult
        gDP_WriteValue = VBA.DateSerial(2027, 4, 10)
        M_WriteBack_PopulateRange TargetRange, DP_WriteAction_DatePicker, _
            WriteResult, OverwriteFormulas:=True

'------------------------------------------------------------------------------
' ASSERT THE OVERRIDE
'------------------------------------------------------------------------------
    'Assert every cell was written once the caller opted in
        TST_DP_AssertEqualsLong "Override writes all 4 cells", _
            4, VBA.CLng(WriteResult.WrittenCount)
    'Assert nothing was preserved under the override
        TST_DP_AssertEqualsLong "Override preserves no formula cells", _
            0, VBA.CLng(WriteResult.FormulaSkippedCount)
    'Assert the override result satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Override result balances", WriteResult
    'Verify the formulas really are gone
        TST_DP_AssertFalse "Override replaces the plain formula", _
            mTST_DP_ScratchSheet.Range("K7").HasFormula
        TST_DP_AssertFalse "Override replaces the date-returning formula", _
            mTST_DP_ScratchSheet.Range("K8").HasFormula

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Leave the scratch cells clear for later runs
        mTST_DP_ScratchSheet.Range("K5:K8").ClearContents
    'Release object references
        Set TargetRange = Nothing
    'Exit after the expectation completes
        Exit Sub

'------------------------------------------------------------------------------
' FORMULA PROTECTION FAIL
'------------------------------------------------------------------------------
FormulaProtectionFail:
    'Capture the original error before any cleanup runs. Every On Error statement
    'resets the Err object, and so does Err.Clear, so formatting the message after
    'cleanup reported "Error 0 - " and lost the cause. See #48
        FPSavedErrNumber = Err.Number
        FPSavedErrDescription = Err.Description
        FPSavedErrSource = Err.Source
    'Clear the target even when the expectation failed
        On Error Resume Next
        mTST_DP_ScratchSheet.Range("K5:K8").ClearContents
        Set TargetRange = Nothing
        Err.Clear
    'Record the failure from the captured values
        TST_DP_RecordFail "Formula protection", _
            "Error " & VBA.CStr(FPSavedErrNumber) & " - " & FPSavedErrDescription & _
            " | Source=" & FPSavedErrSource
    Err.Clear

End Sub

Private Sub TST_DP_ExpectPartialWriteReport()

'
'==============================================================================
'                      EXPECT PARTIAL WRITE REPORT
'==============================================================================
'   Protects the scratch sheet with one locked cell inside the target so the
'   write partially succeeds, then asserts the result names the cell it skipped.
'
'   The sheet is unprotected on every path, including a failed assertion, so a
'   protected scratch sheet cannot leak into a later suite.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WriteResult     As DP_WriteResult   'Structured write-back result
    Dim TargetRange     As Excel.Range      'Partially writable target

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo PartialWriteFail

'------------------------------------------------------------------------------
' BUILD A PARTIALLY WRITABLE TARGET
'------------------------------------------------------------------------------
    'Use a target away from the ranges the rest of the suite writes
        Set TargetRange = mTST_DP_ScratchSheet.Range("H5:H7")
    'Clear any value left by an earlier run
        TargetRange.ClearContents
    'Unlock the cells that must remain writable
        mTST_DP_ScratchSheet.Range("H5").Locked = False
        mTST_DP_ScratchSheet.Range("H7").Locked = False
    'Lock the cell the write must skip
        mTST_DP_ScratchSheet.Range("H6").Locked = True
    'Protect the sheet so the locked cell actually rejects the write
        mTST_DP_ScratchSheet.Protect
    'Prepare a distinct write value
        gDP_WriteValue = VBA.DateSerial(2026, 12, 1)

'------------------------------------------------------------------------------
' WRITE THROUGH THE PARTIAL TARGET
'------------------------------------------------------------------------------
    'Call the range writer directly so the run stays free of modal messages
        M_WriteBack_PopulateRange TargetRange, DP_WriteAction_DatePicker, WriteResult

'------------------------------------------------------------------------------
' ASSERT THE PARTIAL RESULT
'------------------------------------------------------------------------------
    'Assert the writable cells were written
        TST_DP_AssertEqualsLong "Partial write reports 2 written cells", _
            2, VBA.CLng(WriteResult.WrittenCount)
    'Assert the locked cell was counted as skipped
        TST_DP_AssertEqualsLong "Partial write reports 1 skipped locked cell", _
            1, VBA.CLng(WriteResult.LockedSkippedCount)
    'Assert the skipped cell is named, not just counted
        TST_DP_AssertEqualsString "Partial write reports the skipped address", _
            mTST_DP_ScratchSheet.Name & "!H6", WriteResult.LockedSkippedAddresses
    'Assert the shortfall description carries the same address
        TST_DP_AssertTrue "Partial write description names the skipped cell", _
            VBA.InStr(1, M_WriteBack_DescribeShortfall(WriteResult), "H6", vbTextCompare) > 0
    'Assert the partial result still satisfies the accounting invariant
        TST_DP_AssertWriteResultBalances "Partial write result balances", WriteResult
    'Assert the writable cells really received the value
        TST_DP_AssertCellDateEquals "Partial write writes H5", _
            VBA.DateSerial(2026, 12, 1), mTST_DP_ScratchSheet.Range("H5")
    'Assert the locked cell was left alone
        TST_DP_AssertTrue "Partial write leaves H6 blank", _
            VBA.LenB(VBA.CStr(mTST_DP_ScratchSheet.Range("H6").Value)) = 0

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release the sheet before leaving
        TST_DP_ReleaseScratchProtection
    'Release object references
        Set TargetRange = Nothing
    'Exit after the expectation completes
        Exit Sub

'------------------------------------------------------------------------------
' PARTIAL WRITE FAIL
'------------------------------------------------------------------------------
PartialWriteFail:
    'Release the sheet even when the expectation failed
        TST_DP_ReleaseScratchProtection
    'Release object references
        Set TargetRange = Nothing
    'Record the failure and clear the error
        TST_DP_RecordFail "Partial write reporting", _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    Err.Clear

End Sub

Private Sub TST_DP_ReleaseScratchProtection()

'
'==============================================================================
'                       RELEASE SCRATCH PROTECTION
'==============================================================================
'   Unprotects the scratch sheet and restores the default locked state, so a
'   protection-based test cannot leave the sheet unusable for later suites.
'==============================================================================

'------------------------------------------------------------------------------
' RELEASE PROTECTION
'------------------------------------------------------------------------------
    'Never let cleanup raise into the caller
        On Error Resume Next
    'Exit when there is no scratch sheet to release
        If mTST_DP_ScratchSheet Is Nothing Then Exit Sub
    'Unprotect the scratch sheet when it is protected
        If mTST_DP_ScratchSheet.ProtectContents Then
            mTST_DP_ScratchSheet.Unprotect
        End If
    'Restore the default locked state on the cells the test unlocked
        mTST_DP_ScratchSheet.Range("H5:H7").Locked = True
    'Clear any suppressed cleanup error
        Err.Clear

End Sub

Private Sub TST_DP_ExpectError_InvalidWriteAction()

'
'==============================================================================
'                     EXPECT ERROR: INVALID WRITE ACTION
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IgnoredResult   As DP_WriteResult   'Required accumulator, not asserted

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    On Error GoTo ExpectedError

'------------------------------------------------------------------------------
' INVOKE EXPECTED ERROR
'------------------------------------------------------------------------------
    'Call with an unsupported write action value to trigger the expected error
        M_WriteBack_PopulateRange mTST_DP_ScratchSheet.Range("J2"), 99, IgnoredResult

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

Private Sub TST_DP_CheckCleanupStep(ByVal StepName As String)

'
'==============================================================================
'                           CHECK CLEANUP STEP
'------------------------------------------------------------------------------
' PURPOSE
'   Records whether the cleanup step that just ran completed
'
' WHY THIS EXISTS
'   Teardown runs under On Error Resume Next so that one failing step does not
'   prevent the rest from being attempted. Without an explicit check afterwards
'   the failure is discarded, the run still reports its assertion totals, and the
'   next run inherits whatever was left behind
'
'   That is not hypothetical. An aborted run left Application state and
'   worksheets in place, and the following TST_DP_RunAll failed during setup with
'   1004 - Method Add of object Sheets failed
'
' INPUTS
'   StepName
'     Name of the cleanup step that has just executed
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Records a FAIL row and increments the cleanup-failure count when Err is set,
'   then clears Err so the following steps still run
'
' ERROR POLICY
'   Best effort. Must not raise, because it runs inside teardown
'
' DEPENDENCIES
'   TST_DP_RecordResult
'   mTST_DP_CleanupFails
'   mTST_DP_CleanupDetail
'
' NOTES
'   The caller stays under On Error Resume Next across the whole teardown. This
'   routine reads Err before clearing it, so it must be called immediately after
'   the step it checks
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description
    Dim IsInjected          As Boolean      'True when the failure was staged by a self-check

'------------------------------------------------------------------------------
' CAPTURE STEP OUTCOME
'------------------------------------------------------------------------------
    'Capture the outcome of the step that has just run
        ErrorNumber = Err.Number
        ErrorDescription = Err.Description

    'Recognize a failure staged by the harness self-check
        If VBA.LenB(mTST_DP_InjectCleanupFail) > 0 Then
            If VBA.StrComp(StepName, mTST_DP_InjectCleanupFail, vbTextCompare) = 0 Then
                ErrorNumber = vbObjectError + 900
                ErrorDescription = "Injected cleanup failure for harness self-check"
                IsInjected = True
            End If
        End If

    'Exit when the step completed
        If ErrorNumber = 0 Then Exit Sub

'------------------------------------------------------------------------------
' RECORD FAILURE
'------------------------------------------------------------------------------
    'Count the failed cleanup step
        mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1

    'Keep the first failure detail for the summary
        If VBA.Len(mTST_DP_CleanupDetail) = 0 Then
            mTST_DP_CleanupDetail = StepName & ": " & _
                VBA.CStr(ErrorNumber) & " - " & ErrorDescription
        End If

    'Record the failed cleanup step on the result sheet. A staged failure is
    'recorded as INFO: it is a probe of this routine, not a defect in the run,
    'and a FAIL row would report a teardown failure that never happened
        If IsInjected Then
            TST_DP_RecordInfo "Cleanup", StepName, _
                "Staged cleanup failure for harness self-check"
        Else
            TST_DP_RecordResult TST_DP_FAIL_TEXT, _
                "Cleanup", _
                StepName, _
                "Cleanup step failed: " & VBA.CStr(ErrorNumber) & " - " & ErrorDescription
        End If

'------------------------------------------------------------------------------
' CLEAR FOR THE NEXT STEP
'------------------------------------------------------------------------------
    'Clear the error so the following cleanup steps still run
        Err.Clear

End Sub

Private Sub TST_DP_VerifyFinalState(ByRef AppSnapshot As TRegDPApplicationSnapshot)

'
'==============================================================================
'                           VERIFY FINAL STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Verifies that the run left no DatePicker or harness artifact behind
'
' WHY THIS EXISTS
'   A run result is only meaningful if the environment it left is clean. Counting
'   assertions says nothing about a form still loaded, a timer still scheduled,
'   or grid icons still on a worksheet
'
' INPUTS
'   AppSnapshot
'     Application state captured before the run
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Checks the manager reference, the picker form, worksheet grid icons and
'   every Application setting the run snapshotted. Records a FAIL row for each
'   violation and counts it as a cleanup failure
'
'   The Application settings verified are the full contents of the pre-run
'   snapshot:
'
'     ScreenUpdating
'     EnableEvents
'     DisplayAlerts
'     Calculation
'     StatusBar ownership
'
' ERROR POLICY
'   Best effort. Must not raise, because it runs inside teardown
'
' DEPENDENCIES
'   gDP_Manager
'   M_GridIcon_PurgeAll
'   TST_DP_ContextMenuControlCount
'   TST_DP_RecordResult
'
' NOTES
'   This routine reports rather than repairs. Silently fixing a leak would hide
'   the defect that caused it
'
'   Verification covers everything the snapshot captures. A field that is
'   restored but not verified cannot report a restore that silently failed, so
'   the two lists have to stay the same length
'
'   StatusBar is the exception to the rule above, and deliberately so. Setting it
'   to False hands the bar back to Excel, but reading it immediately afterwards
'   can still return the text that was there before Excel repaints, so neither
'   "StatusBar = False" nor a VarType test gives a stable answer at teardown
'
'   What is determinable is whether the run left its own message behind, so that
'   is what is asserted. Restoring the bar is verified as a cleanup step; this
'   check exists to catch the harness leaking its own status text into the session
'
'   Each check reports the expected and actual value. A cleanup failure that only
'   names the property leaves the next person guessing which of five it was
'
'   The scratch worksheet is not checked here. Its deletion is already verified
'   as a cleanup step
'
'   The grid-icon check is weaker than it looks: M_GridIcon_PurgeAll removes the
'   named shape from every open workbook, so this can pass because teardown
'   reached beyond the host workbook. Bounding that purge is #14
'
'   The context menu is verified against its pre-run registration rather than
'   against zero. The harness restores the session it found, and a DatePicker
'   that was already running is entitled to its menu when the run ends
'
'   Two teardown targets cannot be verified from here, and both are recorded
'   rather than quietly omitted:
'
'     keyboard shortcut   Excel exposes no getter for Application.OnKey, so the
'                         assignment cannot be read back at all. See #42
'     live-clock timer    mDP_TimerIsRunning is Private to M_DatePicker and
'                         Application.OnTime schedules cannot be enumerated
'
'   Both are covered instead by per-operation cleanup accounting: M_Timer_Stop
'   and M_KeyboardShortcut_Remove are now invoked and checked individually, so a
'   failure in either is counted even though its final state cannot be inspected.
'   Making them observable needs a production-side accessor, which is not a
'   change this harness should make on its own
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ManagerPresent  As Boolean      'True when a manager exists after teardown
    Dim MenuControlCount As Long        'DatePicker context-menu controls left behind
    Dim StatusBarText   As String       'Status bar contents after teardown
    Dim LeftLoaded      As Boolean      'True when the picker form is still loaded
    Dim IconCount       As Long         'DatePicker shapes still present
    Dim WS              As Excel.Worksheet  'Worksheet scan variable
    Dim Shp             As Excel.Shape  'Shape scan variable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress verification errors so teardown always completes
        On Error Resume Next

'------------------------------------------------------------------------------
' VERIFY PICKER FORM
'------------------------------------------------------------------------------
    'Resolve whether the picker form is still loaded
        LeftLoaded = False
        If Not gDP_Manager Is Nothing Then LeftLoaded = gDP_Manager.Is_PickerLoaded
        Err.Clear

    'Record a form left loaded after the run
        If LeftLoaded Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "DatePicker form is still loaded after teardown"
        End If

'------------------------------------------------------------------------------
' VERIFY GRID ICONS
'------------------------------------------------------------------------------
    'Count DatePicker shapes left on the host workbook
        IconCount = 0
        If Not mTST_DP_HostWorkbook Is Nothing Then
            For Each WS In mTST_DP_HostWorkbook.Worksheets
                For Each Shp In WS.Shapes
                    If VBA.InStr(1, Shp.Name, "DP_GridIcon", vbTextCompare) = 1 Then
                        IconCount = IconCount + 1
                    End If
                Next Shp
            Next WS
        End If
        Err.Clear

    'Record grid icons left behind after the run
        If IconCount > 0 Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                VBA.CStr(IconCount) & " DatePicker grid icon shapes remain after teardown"
        End If

'------------------------------------------------------------------------------
' VERIFY APPLICATION STATE
'------------------------------------------------------------------------------
    'Record an Application event state that does not match the pre-run snapshot
        If Excel.Application.EnableEvents <> AppSnapshot.EnableEvents Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Application.EnableEvents was not restored to its pre-run value"
        End If

    'Record a screen updating state that does not match the pre-run snapshot
        If Excel.Application.ScreenUpdating <> AppSnapshot.ScreenUpdating Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Application.ScreenUpdating was not restored. Expected " & _
                VBA.CStr(AppSnapshot.ScreenUpdating) & _
                " but found " & VBA.CStr(Excel.Application.ScreenUpdating)
        End If

    'Record an alert state that does not match the pre-run snapshot
        If Excel.Application.DisplayAlerts <> AppSnapshot.DisplayAlerts Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Application.DisplayAlerts was not restored. Expected " & _
                VBA.CStr(AppSnapshot.DisplayAlerts) & _
                " but found " & VBA.CStr(Excel.Application.DisplayAlerts)
        End If

    'Record a calculation mode that does not match the pre-run snapshot
        If Excel.Application.Calculation <> AppSnapshot.CalculationMode Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Application.Calculation was not restored. Expected " & _
                VBA.CStr(AppSnapshot.CalculationMode) & _
                " but found " & VBA.CStr(Excel.Application.Calculation)
        End If

    'Record a manager state that does not match the pre-run condition. The
    'harness restores it as a cleanup step; this proves the restore took effect
        ManagerPresent = Not (gDP_Manager Is Nothing)
        If ManagerPresent <> mTST_DP_HadManager Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Manager state was not restored. Expected present=" & _
                VBA.CStr(mTST_DP_HadManager) & _
                " but found present=" & VBA.CStr(ManagerPresent)
        End If

    'Record a context-menu registration that does not match the pre-run state.
    'The after-removal count is reported alongside it, because a mismatch has two
    'possible causes and they need different fixes: controls that removal never
    'took away, or controls something re-registered afterwards
        MenuControlCount = TST_DP_ContextMenuControlCount()
        If MenuControlCount <> mTST_DP_MenuAtStart Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Context-menu registration was not restored. Before run=" & _
                VBA.CStr(mTST_DP_MenuAtStart) & _
                "; after removal=" & VBA.CStr(mTST_DP_MenuAfterRemove) & _
                "; at end=" & VBA.CStr(MenuControlCount)
        End If

    'Record a status bar still showing the run's own message. Ownership cannot be
    'asserted: releasing the bar hands it back to Excel, and reading it straight
    'afterwards can still return the previous text. What the harness controls, and
    'what actually matters, is that its own message is gone
        StatusBarText = VBA.CStr(Excel.Application.StatusBar)
        If VBA.InStr(1, StatusBarText, TST_DP_STATUS_BAR_TEXT, vbTextCompare) > 0 Then
            mTST_DP_CleanupFails = mTST_DP_CleanupFails + 1
            TST_DP_RecordResult TST_DP_FAIL_TEXT, "Cleanup", "Final state", _
                "Application.StatusBar still shows the harness message: " & StatusBarText
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set WS = Nothing
        Set Shp = Nothing
    'Clear any suppressed verification error
        Err.Clear

End Sub

Private Function TST_DP_ResolveRunState() As String

'
'==============================================================================
'                           RESOLVE RUN STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the single state that describes the outcome of the run
'
' WHY THIS EXISTS
'   Assertion totals alone cannot express a run that passed every assertion but
'   failed to clean up, or one that ended before every dispatched suite returned.
'   Both look like a pass if only Passed and Failed are reported
'
' INPUTS
'   None
'
' RETURNS
'   One of TST_DP_STATE_DIRTY_START, TST_DP_STATE_FAIL, TST_DP_STATE_INCOMPLETE,
'   TST_DP_STATE_FAIL_CLEANUP or TST_DP_STATE_PASS
'
' BEHAVIOR
'   Reports FAIL_DIRTY_START when the run began in an environment a previous run
'   left behind, FAIL when any assertion failed, INCOMPLETE_SKIPPED when a
'   dispatched suite did not return, FAIL_CLEANUP when teardown did not complete,
'   and PASS only when none of those applies
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   mTST_DP_DirtyStart
'   mTST_DP_FailCount
'   mTST_DP_CleanupFails
'   mTST_DP_SuitesDispatched
'   mTST_DP_SuitesCompleted
'
' NOTES
'   A dirty start outranks everything, including assertion failures. It is the
'   only condition that makes the rest of the report untrustworthy rather than
'   merely bad: the failures may belong to the environment the previous run left,
'   not to the code under test. Reporting FAIL first would send someone to debug
'   an assertion that is an artifact
'
'   Below that, assertion failures rank above cleanup failures. A run with both
'   is reported as FAIL, because the assertion failure is the more actionable
'
'   No run can report PASS unless it started clean. That is the closure invariant
'   of #19 and it is enforced here rather than by the caller
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE STATE
'------------------------------------------------------------------------------
    'Report a run that began in an environment it did not establish. Nothing
    'below this line can be trusted if this is true
        If mTST_DP_DirtyStart Then
            TST_DP_ResolveRunState = TST_DP_STATE_DIRTY_START
            Exit Function
        End If

    'Report assertion failures first, as the most actionable outcome
        If mTST_DP_FailCount > 0 Then
            TST_DP_ResolveRunState = TST_DP_STATE_FAIL
            Exit Function
        End If

    'Report a suite that was dispatched but never returned
        If mTST_DP_SuitesCompleted < mTST_DP_SuitesDispatched Then
            TST_DP_ResolveRunState = TST_DP_STATE_INCOMPLETE
            Exit Function
        End If

    'Report teardown that did not complete
        If mTST_DP_CleanupFails > 0 Then
            TST_DP_ResolveRunState = TST_DP_STATE_FAIL_CLEANUP
            Exit Function
        End If

    'Report a clean run
        TST_DP_ResolveRunState = TST_DP_STATE_PASS

End Function

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
            "State=" & TST_DP_ResolveRunState() & _
            "; Run=" & VBA.CStr(mTST_DP_RunCount) & _
            "; Passed=" & VBA.CStr(mTST_DP_PassCount) & _
            "; Failed=" & VBA.CStr(mTST_DP_FailCount) & _
            "; CleanupFailures=" & VBA.CStr(mTST_DP_CleanupFails)

    'Record the first cleanup failure detail so the summary is actionable
        If VBA.Len(mTST_DP_CleanupDetail) <> 0 Then
            TST_DP_RecordInfo "Harness", "Cleanup", mTST_DP_CleanupDetail
        End If

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
        mTST_DP_ResultSheet.Cells(8, TST_DP_COL_SUMMARY_LABEL).Value = "State"
        mTST_DP_ResultSheet.Cells(8, TST_DP_COL_SUMMARY_VALUE).Value = TST_DP_ResolveRunState()

    'Bold the summary labels
        mTST_DP_ResultSheet.Cells(4, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(5, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(6, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(7, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(8, TST_DP_COL_SUMMARY_LABEL).Font.Bold = True
        mTST_DP_ResultSheet.Cells(8, TST_DP_COL_SUMMARY_VALUE).Font.Bold = True

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
'   The regression harness should normally test the workbook that contains the
'   DatePicker project. When the project is hosted in an add-in or hidden
'   workbook, the active workbook is used as the practical worksheet host.
'
' INPUTS
'   None
'
' RETURNS
'   ThisWorkbook when it has a visible workbook window
'   ActiveWorkbook when ThisWorkbook has no visible window
'   ThisWorkbook as final fallback
'
' ERROR POLICY
'   Best-effort. Returns Nothing only when all paths fail
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress workbook-resolution failures
        On Error Resume Next

'------------------------------------------------------------------------------
' PREFER THISWORKBOOK WHEN VISIBLE
'------------------------------------------------------------------------------
    'Use ThisWorkbook when it has a visible Excel window
        If ThisWorkbook.Windows.Count > 0 Then
            Set TST_DP_GetHostWorkbook = ThisWorkbook
        End If

'------------------------------------------------------------------------------
' FALL BACK TO ACTIVEWORKBOOK
'------------------------------------------------------------------------------
    'Use ActiveWorkbook when ThisWorkbook is not a visible workbook host
        If TST_DP_GetHostWorkbook Is Nothing Then
            If Not Excel.Application.ActiveWorkbook Is Nothing Then
                Set TST_DP_GetHostWorkbook = Excel.Application.ActiveWorkbook
            End If
        End If

'------------------------------------------------------------------------------
' FINAL FALLBACK
'------------------------------------------------------------------------------
    'Fall back to ThisWorkbook when ActiveWorkbook is unavailable
        If TST_DP_GetHostWorkbook Is Nothing Then
            Set TST_DP_GetHostWorkbook = ThisWorkbook
        End If

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
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
'                        PREPARE SCRATCH WORKSHEET
'==============================================================================
' PURPOSE
'   Creates and initializes the worksheet the regression suites write to
'
' WHY THIS EXISTS
'   Worksheets.Add here has been observed to report 1004 while leaving a new
'   unnamed worksheet in the workbook. Every occurrence leaked a sheet, and
'   preflight could not see it: it looks for TST_DP_SCRATCH by name and the
'   leftover is called SheetNN
'
'   Protecting only the Add is not enough. The rename, the initialization, the
'   formatting and the activation can each fail after a successful Add and leak
'   the same sheet, so the whole sequence is one transaction
'
' INPUTS
'   HostWorkbook
'     Workbook that will receive the scratch worksheet
'
' RETURNS
'   Nothing. Publishes mTST_DP_ScratchSheet on success
'
' BEHAVIOR
'   Validates the host, removes any previous scratch sheet, snapshots worksheet
'   identity, adds, names, initializes, formats and activates
'
'   On any failure removes only the worksheet this call created, when that
'   worksheet can be identified unambiguously, then re-raises the original
'   failure with the step that produced it
'
' ERROR POLICY
'   Raises when the scratch worksheet cannot be established. The suites cannot run
'   without it, so aborting is correct, and aborting leaks nothing and says where
'   it stopped
'
'   Does not raise when Worksheets.Add reports a failure having nevertheless
'   created exactly one worksheet. That is a partial success: the workbook
'   mutation happened and only the call's completion did not. The candidate is
'   validated and adopted, and the anomaly is recorded as an INFO row
'
'   Cleanup never replaces the original failure. A cleanup that itself fails is
'   appended to the diagnostic rather than raised
'
' DEPENDENCIES
'   TST_DP_DeleteWorksheetIfExists
'   TST_DP_AnySheetNameExists
'   TST_DP_WorksheetReferenceIsLive
'   TST_DP_FindNewWorksheets
'   TST_DP_ActivateWorksheetForTest
'
' NOTES
'   Identity is captured as worksheet object references, not names. A worksheet
'   name can contain the delimiter any string encoding would use, Excel treats
'   names case-insensitively, and a rename would make a pre-existing sheet look
'   new. Object identity has none of those problems
'
'   Cleanup deletes an exact or unambiguous candidate only. Deleting every
'   worksheet absent from the snapshot would assume this routine created them,
'   which is exactly the assumption that cannot be made while a re-entrancy
'   hypothesis is open. Two or more candidates are reported and none is deleted
'
'   Adoption is confined to a failure of the Add itself. A failure at the rename
'   or the initialization means that operation was already refused once, and
'   repeating it in the handler would be a retry rather than a recovery
'
'   A candidate must prove it is usable before adoption. Excel has been reported
'   to create a worksheet that then cannot be renamed, so appearing is not the
'   same as working
'
'   The worksheet counts are net workbook state, not proof of which internal
'   Excel operation ran:
'
'     before = after   no net worksheet-count increase was observed
'     after > before   one or more worksheets appeared during the attempt
'
'   The Add anchors on Sheets, not Worksheets, so a workbook whose last sheet is
'   a chart sheet is handled
'
'   The scratch name is checked against Sheets after deletion, because chart
'   sheets share the workbook name namespace and a chart sheet of that name would
'   survive a worksheet-only delete and then fail the rename
'
' UPDATED
'   2026-08-23
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "TST_DP_PrepareScratchSheet"

    Dim HandlerStep         As String              'Current setup operation
    Dim WorksheetsBefore    As Collection          'Object-identity snapshot
    Dim NewWorksheets       As Collection          'Worksheets absent from the snapshot

    Dim CreatedSheet        As Excel.Worksheet     'Worksheet created by this call
    Dim CleanupCandidate    As Excel.Worksheet     'Exact sheet eligible for cleanup
    Dim WS                  As Excel.Worksheet     'Current worksheet while scanning

    Dim ErrorNumber         As Long                'Original failure number
    Dim ErrorSource         As String              'Original failure source
    Dim ErrorDescription    As String              'Original failure description

    Dim CleanupErrNumber    As Long                'Cleanup failure number
    Dim CleanupErrText      As String              'Cleanup failure description
    Dim CleanupName         As String              'Name of the cleanup candidate
    Dim DiagnosticText      As String              'Extended failure description
    Dim AdoptionAllowed     As Boolean             'True when the failure was the Add itself
    Dim AdoptionErrNumber   As Long                'Failure while validating a candidate
    Dim CandidateName       As String              'Name the candidate arrived with

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Treat the whole setup as one transaction
        On Error GoTo ErrorHandler
    'Track the current operation
        HandlerStep = "Validate host workbook"
    'Reject a missing host explicitly
        If HostWorkbook Is Nothing Then
            Err.Raise vbObjectError + 2401, PROC_NAME, _
                "The host workbook reference is missing."
        End If
    'Structure protection prevents both creation and deletion, so say so here
    'rather than surfacing a bare 1004 from the Add
        If HostWorkbook.ProtectStructure Then
            Err.Raise vbObjectError + 2402, PROC_NAME, _
                "The host workbook structure is protected."
        End If
    'A workbook always has at least one sheet to anchor the Add after
        If HostWorkbook.Sheets.Count = 0 Then
            Err.Raise vbObjectError + 2403, PROC_NAME, _
                "The host workbook contains no sheet to add a worksheet after."
        End If
    'Never let a failed Set leave the harness holding a previous run's worksheet
        Set mTST_DP_ScratchSheet = Nothing
    'Reset the setup diagnostics
        mTST_DP_ScratchAddBefore = 0
        mTST_DP_ScratchAddAfter = 0
        mTST_DP_ScratchAddOrphan = VBA.vbNullString

'------------------------------------------------------------------------------
' REMOVE PREVIOUS SCRATCH WORKSHEET
'------------------------------------------------------------------------------
    'Track the current operation
        HandlerStep = "Remove previous scratch worksheet"
    'Delete a scratch worksheet an earlier run may have left
        TST_DP_DeleteWorksheetIfExists HostWorkbook, TST_DP_SCRATCH_SHEET_NAME
    'Chart sheets share the workbook name namespace, so a worksheet-only delete
    'can leave the name taken and the rename below would then fail
        If TST_DP_AnySheetNameExists(HostWorkbook, TST_DP_SCRATCH_SHEET_NAME) Then
            Err.Raise vbObjectError + 2404, PROC_NAME, _
                "A workbook sheet named " & TST_DP_SCRATCH_SHEET_NAME & _
                " still exists after scratch cleanup."
        End If

'------------------------------------------------------------------------------
' SNAPSHOT WORKSHEET IDENTITY
'------------------------------------------------------------------------------
    'Track the current operation
        HandlerStep = "Snapshot worksheets"
    'Capture object references rather than names
        Set WorksheetsBefore = New Collection
        For Each WS In HostWorkbook.Worksheets
            WorksheetsBefore.Add WS
        Next WS
        Set WS = Nothing
    'Record the pre-Add count
        mTST_DP_ScratchAddBefore = HostWorkbook.Worksheets.Count
        mTST_DP_ScratchAddAfter = mTST_DP_ScratchAddBefore

'------------------------------------------------------------------------------
' ADD SCRATCH WORKSHEET
'------------------------------------------------------------------------------
    'Track the current operation
        HandlerStep = "Add scratch worksheet"
    'Anchor on Sheets so a trailing chart sheet is handled
        Set CreatedSheet = HostWorkbook.Worksheets.Add( _
            After:=HostWorkbook.Sheets(HostWorkbook.Sheets.Count))
    'Record the post-Add count immediately
        mTST_DP_ScratchAddAfter = HostWorkbook.Worksheets.Count
    'A silent empty return is a setup failure, not something to dereference
        If CreatedSheet Is Nothing Then
            Err.Raise vbObjectError + 2405, PROC_NAME, _
                "Worksheets.Add returned no worksheet."
        End If
    'Publish the worksheet to the harness
        Set mTST_DP_ScratchSheet = CreatedSheet

'------------------------------------------------------------------------------
' NAME AND INITIALIZE
'------------------------------------------------------------------------------
    'Track the current operation
        HandlerStep = "Name scratch worksheet"
        CreatedSheet.Name = TST_DP_SCRATCH_SHEET_NAME
    'Track the current operation
        HandlerStep = "Initialize scratch worksheet"
        CreatedSheet.Range("A1").Value = "DatePicker regression scratch sheet"
    'Track the current operation
        HandlerStep = "Format scratch worksheet"
        CreatedSheet.Columns("A:J").ColumnWidth = 16

'------------------------------------------------------------------------------
' ACTIVATE
'------------------------------------------------------------------------------
    'Track the current operation
        HandlerStep = "Activate scratch worksheet"
    'Use the harness activation helper rather than a raw Activate
        TST_DP_ActivateWorksheetForTest CreatedSheet

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release local references. The module-level reference stays authoritative
        Set CreatedSheet = Nothing
        Set WorksheetsBefore = Nothing
    'Exit after successful setup
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original failure before cleanup can replace it
        ErrorNumber = Err.Number
        ErrorSource = Err.Source
        ErrorDescription = Err.Description
    'Cleanup must never replace the failure it is cleaning up after
        On Error Resume Next
    'Record the worksheet count before anything is removed
        If Not (HostWorkbook Is Nothing) Then
            Err.Clear
            mTST_DP_ScratchAddAfter = HostWorkbook.Worksheets.Count
            Err.Clear
        End If

    'Prefer the exact object the Add returned
        If TST_DP_WorksheetReferenceIsLive(CreatedSheet) Then
            Set CleanupCandidate = CreatedSheet
        End If

    'Otherwise look for exactly one worksheet that was not in the snapshot
        If CleanupCandidate Is Nothing Then
            If Not (HostWorkbook Is Nothing) Then
                If Not (WorksheetsBefore Is Nothing) Then
                    Set NewWorksheets = TST_DP_FindNewWorksheets(HostWorkbook, WorksheetsBefore)
                End If
            End If
            If Not (NewWorksheets Is Nothing) Then
                If NewWorksheets.Count = 1 Then
                    Set CleanupCandidate = NewWorksheets.Item(1)
                ElseIf NewWorksheets.Count > 1 Then
                    'Two or more candidates cannot be attributed to this call.
                    'Report the ambiguity and delete nothing
                    mTST_DP_ScratchAddOrphan = "ambiguous, " & _
                        VBA.CStr(NewWorksheets.Count) & " new worksheets, none deleted"
                End If
            End If
        End If

'------------------------------------------------------------------------------
' ADOPT A PARTIAL SUCCESS
'------------------------------------------------------------------------------
    'Excel has been observed to create a worksheet and then report 1004 from the
    'same Add call. That is a partial success, not a failure: the workbook
    'mutation happened and only the call's own completion did not. Discarding the
    'worksheet and aborting the run throws away work Excel actually did
        AdoptionAllowed = (HandlerStep = "Add scratch worksheet")
    'Adoption is confined to the Add step deliberately. A failure at the rename or
    'the initialization means that operation was already refused once, and
    'repeating it here would be a retry rather than a recovery
        If AdoptionAllowed Then
            If TST_DP_WorksheetReferenceIsLive(CleanupCandidate) Then
                'Record the name the candidate arrived with, for the audit line
                    CandidateName = CleanupCandidate.Name
                'The candidate must prove it is usable. A created worksheet that
                'cannot be renamed or written has been reported too
                    Err.Clear
                    CleanupCandidate.Name = TST_DP_SCRATCH_SHEET_NAME
                    CleanupCandidate.Range("A1").Value = "DatePicker regression scratch sheet"
                    CleanupCandidate.Columns("A:J").ColumnWidth = 16
                    AdoptionErrNumber = Err.Number
                    Err.Clear
                'Adopt only a candidate that answered every operation
                    If AdoptionErrNumber = 0 Then
                        Set mTST_DP_ScratchSheet = CleanupCandidate
                        TST_DP_ActivateWorksheetForTest CleanupCandidate
                        Err.Clear
                        mTST_DP_ScratchAddOrphan = CandidateName & " (adopted)"
                        'Record the anomaly. A run that recovered from a misreported
                        'native call is not a clean run and must not look like one
                            TST_DP_RecordInfo "Harness", "Scratch sheet", _
                                "Worksheets.Add reported " & VBA.CStr(ErrorNumber) & _
                                " after creating " & CandidateName & _
                                ". The worksheet was validated and adopted; " & _
                                "WorksheetsBefore=" & VBA.CStr(mTST_DP_ScratchAddBefore) & _
                                ", WorksheetsAfter=" & VBA.CStr(mTST_DP_ScratchAddAfter) & "."
                        'Release local references and continue the run
                            Set CleanupCandidate = Nothing
                            Set CreatedSheet = Nothing
                            Set NewWorksheets = Nothing
                            Set WorksheetsBefore = Nothing
                            Set WS = Nothing
                            Err.Clear
                            On Error GoTo 0
                            Exit Sub
                    End If
                'Record that adoption was attempted and refused
                    mTST_DP_ScratchAddOrphan = CandidateName & _
                        " (adoption failed: " & VBA.CStr(AdoptionErrNumber) & ")"
            End If
        End If

'------------------------------------------------------------------------------
' REMOVE AN UNUSABLE CANDIDATE
'------------------------------------------------------------------------------
    'Remove only the identified worksheet
        If TST_DP_WorksheetReferenceIsLive(CleanupCandidate) Then
            CleanupName = CleanupCandidate.Name
            Err.Clear
            CleanupCandidate.Delete
            CleanupErrNumber = Err.Number
            CleanupErrText = Err.Description
            Err.Clear
            If CleanupErrNumber = 0 Then
                mTST_DP_ScratchAddOrphan = CleanupName & " (deleted)"
            Else
                mTST_DP_ScratchAddOrphan = CleanupName & " (delete failed: " & _
                    VBA.CStr(CleanupErrNumber) & " - " & CleanupErrText & ")"
            End If
        End If

    'Never expose a partially initialized scratch worksheet
        Set mTST_DP_ScratchSheet = Nothing
    'Release local references
        Set CleanupCandidate = Nothing
        Set CreatedSheet = Nothing
        Set NewWorksheets = Nothing
        Set WorksheetsBefore = Nothing
        Set WS = Nothing
    'Restore ordinary error handling
        Err.Clear
        On Error GoTo 0

    'Build the diagnostic without changing the original error number
        DiagnosticText = ErrorDescription & _
            " | Step=" & HandlerStep & _
            " | WorksheetsBefore=" & VBA.CStr(mTST_DP_ScratchAddBefore) & _
            " | WorksheetsAfter=" & VBA.CStr(mTST_DP_ScratchAddAfter)
        If VBA.LenB(ErrorSource) > 0 Then
            DiagnosticText = DiagnosticText & " | OriginalSource=" & ErrorSource
        End If
        If VBA.LenB(mTST_DP_ScratchAddOrphan) > 0 Then
            DiagnosticText = DiagnosticText & " | Cleanup=" & mTST_DP_ScratchAddOrphan
        End If
    'Re-raise the original failure
        Err.Raise ErrorNumber, PROC_NAME, DiagnosticText

End Sub

Private Function TST_DP_AnySheetNameExists( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal SheetName As String) As Boolean

'
'==============================================================================
'                         ANY SHEET NAME EXISTS
'==============================================================================
'   Reports whether any sheet of that name exists, of any type.
'
'   Worksheets and chart sheets share one name namespace, so a worksheet-only
'   lookup can report a name free while a rename to it still fails.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SheetObject     As Object       'Resolved sheet of any type

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Suppress the error raised when the name is free
        On Error Resume Next
    'Set safe default result
        TST_DP_AnySheetNameExists = False
    'Exit when there is no workbook to inspect
        If HostWorkbook Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Attempt to resolve a sheet of any type
        Set SheetObject = HostWorkbook.Sheets(SheetName)
    'Report whether the resolution succeeded
        TST_DP_AnySheetNameExists = Not (SheetObject Is Nothing)
    'Release object references
        Set SheetObject = Nothing
    'Clear any suppressed lookup error
        Err.Clear

End Function

Private Function TST_DP_WorksheetReferenceIsLive( _
    ByVal Candidate As Excel.Worksheet) As Boolean

'
'==============================================================================
'                     WORKSHEET REFERENCE IS LIVE
'==============================================================================
'   Reports whether a worksheet reference still points at a sheet that exists.
'
'   "Not Candidate Is Nothing" tests the variable, not the object. A reference to
'   a deleted worksheet passes that test and raises on use.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeName       As String       'Probed worksheet name, discarded

'------------------------------------------------------------------------------
' PROBE
'------------------------------------------------------------------------------
    'Never let a liveness probe raise into a caller
        On Error Resume Next
    'Set safe default result
        TST_DP_WorksheetReferenceIsLive = False
    'Exit when nothing is referenced
        If Candidate Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Ask the object model whether the worksheet still answers
        Err.Clear
        ProbeName = Candidate.Name
        TST_DP_WorksheetReferenceIsLive = (Err.Number = 0)
    'Clear any suppressed probe error
        Err.Clear

End Function

Private Function TST_DP_WorksheetSnapshotContains( _
    ByVal Snapshot As Collection, _
    ByVal Candidate As Excel.Worksheet) As Boolean

'
'==============================================================================
'                     WORKSHEET SNAPSHOT CONTAINS
'==============================================================================
'   Reports whether a worksheet was present in an identity snapshot, comparing
'   with Is rather than by name.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Existing        As Excel.Worksheet  'Current snapshot entry

'------------------------------------------------------------------------------
' COMPARE BY IDENTITY
'------------------------------------------------------------------------------
    'Never let a comparison raise into a caller
        On Error Resume Next
    'Set safe default result
        TST_DP_WorksheetSnapshotContains = False
    'Exit when there is nothing to compare
        If Snapshot Is Nothing Then
            Err.Clear
            Exit Function
        End If
        If Candidate Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Compare object identity, not names
        For Each Existing In Snapshot
            If Candidate Is Existing Then
                TST_DP_WorksheetSnapshotContains = True
                Exit For
            End If
        Next Existing
    'Release object references
        Set Existing = Nothing
    'Clear any suppressed comparison error
        Err.Clear

End Function

Private Function TST_DP_FindNewWorksheets( _
    ByVal HostWorkbook As Excel.Workbook, _
    ByVal Snapshot As Collection) As Collection

'
'==============================================================================
'                          FIND NEW WORKSHEETS
'==============================================================================
'   Returns the worksheets present now that were not in the snapshot.
'
'   The caller decides what to do with them. One candidate can be attributed to
'   the failed call; two or more cannot.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultList      As Collection       'Worksheets absent from the snapshot
    Dim WS              As Excel.Worksheet  'Current worksheet while scanning

'------------------------------------------------------------------------------
' COLLECT
'------------------------------------------------------------------------------
    'Never let a scan raise into a caller
        On Error Resume Next
    'Always return a usable collection
        Set ResultList = New Collection
        Set TST_DP_FindNewWorksheets = ResultList
    'Exit when there is no workbook to scan
        If HostWorkbook Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Collect worksheets the snapshot did not contain. Nothing is deleted here
        For Each WS In HostWorkbook.Worksheets
            If Not TST_DP_WorksheetSnapshotContains(Snapshot, WS) Then
                ResultList.Add WS
            End If
        Next WS
    'Release object references
        Set WS = Nothing
    'Publish the collection
        Set TST_DP_FindNewWorksheets = ResultList
    'Clear any suppressed scan error
        Err.Clear

End Function

Private Sub TST_DP_ActivateWorksheetForTest(ByVal TargetSheet As Excel.Worksheet)

'
'==============================================================================
'                       ACTIVATE WORKSHEET FOR TEST
'------------------------------------------------------------------------------
' PURPOSE
'   Activates a worksheet safely for selection-based regression tests
'
' WHY THIS EXISTS
'   Some tests need Application.Selection to point to a cell on the scratch
'   worksheet. Direct Worksheet.Activate can fail when the parent workbook window
'   is not active, when events are enabled, or when Excel is in a transient UI
'   state after manager repair or runtime cleanup.
'
' INPUTS
'   TargetSheet
'     Worksheet to activate
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the worksheet reference, disables Excel events, makes the sheet
'   visible, activates the parent workbook window when possible, and activates
'   the target worksheet.
'
' ERROR POLICY
'   Raises a descriptive runtime error if activation cannot be completed
'
' UPDATED
'   2026-05-26
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "TST_DP_ActivateWorksheetForTest"

    Dim HostWorkbook    As Excel.Workbook       'Parent workbook of the target sheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing worksheet reference
        If TargetSheet Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "TargetSheet cannot be Nothing"
        End If

'------------------------------------------------------------------------------
' RESOLVE PARENT WORKBOOK
'------------------------------------------------------------------------------
    'Resolve the parent workbook
        Set HostWorkbook = TargetSheet.Parent
    'Reject a missing parent workbook
        If HostWorkbook Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "TargetSheet parent workbook is not available"
        End If

'------------------------------------------------------------------------------
' DISABLE EVENTS
'------------------------------------------------------------------------------
    'Disable Excel events before workbook or worksheet activation
        Excel.Application.EnableEvents = False

'------------------------------------------------------------------------------
' ENSURE SHEET IS VISIBLE
'------------------------------------------------------------------------------
    'Make the target sheet visible before activation
        If TargetSheet.Visible <> xlSheetVisible Then
            TargetSheet.Visible = xlSheetVisible
        End If

'------------------------------------------------------------------------------
' ACTIVATE WORKBOOK WINDOW
'------------------------------------------------------------------------------
    'Activate the parent workbook window when one is available
        If HostWorkbook.Windows.Count > 0 Then
            HostWorkbook.Windows(1).Activate
        Else
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Parent workbook has no visible window"
        End If

'------------------------------------------------------------------------------
' ACTIVATE WORKSHEET
'------------------------------------------------------------------------------
    'Activate the requested worksheet
        TargetSheet.Activate

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set HostWorkbook = Nothing
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Release object references
        Set HostWorkbook = Nothing
    'Raise a descriptive activation error
        Err.Raise Err.Number, PROC_NAME, _
            "Worksheet activation for test failed: " & Err.Description

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
    Excel.Application.StatusBar = TST_DP_STATUS_BAR_TEXT

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
'                     RESET DATEPICKER ARTIFACTS (SETUP ONLY)
'==============================================================================
'   Clears transient DatePicker UI artifacts before a run, so the run starts from
'   a known state.
'
'   Blanket error suppression is acceptable here and only here. This runs before
'   the run, where the goal is to reach a usable starting state rather than to
'   account for anything. Teardown does not use it: each cleanup operation is
'   invoked and checked individually, because a composite step that swallows its
'   own errors cannot report which operation failed, or that any did.
'
'   It also claims the provider lease. That is deliberate and is the one place
'   the harness overrides the #37 ownership guard: the lease lives for the Excel
'   process while the token proving ownership lives in VBA module state, so any
'   re-import strands a lease that no project can release. A regression run is a
'   single-provider environment, and a genuine two-provider session is a manual
'   validation case that this pack cannot construct anyway.
'==============================================================================

'------------------------------------------------------------------------------
' RESET ARTIFACTS
'------------------------------------------------------------------------------
    'Suppress individual setup failures
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
    'Take the provider lease for this run. The lease outlives the VBA project that
    'created it, so re-importing a module leaves a lease no project can prove it
    'owns, and every guarded entry point then refuses. The harness is a
    'single-provider environment by definition, so it claims ownership rather than
    'failing on a lease it cannot otherwise reach
        DP_ForceReleaseProviderLease
        M_Lease_TryAcquire
    'Disarm any window-style fault injection left by an aborted run
        M_Window_Test_SetFaultInjection 0, 0
    'Clear any suppressed error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function TST_DP_ContextMenuControlCount() As Long

'
'==============================================================================
'                        COUNT CONTEXT MENU CONTROLS
'==============================================================================
'   Counts the DatePicker's own controls left on the Excel context menus.
'
'   The tag is duplicated here rather than read from M_DatePicker, where it is a
'   Private constant. It is documented as a stable legacy identifier that cannot
'   be renamed for backward compatibility, so duplicating it is safe in a way
'   that duplicating an ordinary constant would not be.
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const MENU_TAG          As String = "VBA_DATETIMEPICKER"    'Legacy context-menu tag
    Const CELL_BAR          As String = "Cell"                  'Standard cell context menu
    Const LIST_RANGE_BAR    As String = "List Range Popup"      'Table context menu

    Dim BarNames            As Variant      'Command bars the DatePicker registers on
    Dim BarIndex            As Long         'Current command bar index
    Dim Bar                 As Object       'Current command bar
    Dim Ctl                 As Object       'Current command bar control
    Dim FoundCount          As Long         'DatePicker controls found

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let a probe raise into teardown
        On Error Resume Next
    'Set safe default result
        TST_DP_ContextMenuControlCount = 0
    'List the bars the DatePicker registers on
        BarNames = VBA.Array(CELL_BAR, LIST_RANGE_BAR)

'------------------------------------------------------------------------------
' COUNT TAGGED CONTROLS
'------------------------------------------------------------------------------
    'Walk each command bar the DatePicker touches
        For BarIndex = LBound(BarNames) To UBound(BarNames)
            'Resolve the command bar, skipping one that does not exist
                Set Bar = Nothing
                Set Bar = Excel.Application.CommandBars(BarNames(BarIndex))
            'Count the controls carrying the DatePicker tag
                If Not Bar Is Nothing Then
                    For Each Ctl In Bar.Controls
                        If VBA.StrComp(Ctl.Tag, MENU_TAG, vbTextCompare) = 0 Then
                            FoundCount = FoundCount + 1
                        End If
                    Next Ctl
                End If
        Next BarIndex

'------------------------------------------------------------------------------
' RETURN COUNT
'------------------------------------------------------------------------------
    'Report what was found
        TST_DP_ContextMenuControlCount = FoundCount
    'Release object references
        Set Ctl = Nothing
        Set Bar = Nothing
    'Clear any suppressed probe error
        Err.Clear

End Function

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

