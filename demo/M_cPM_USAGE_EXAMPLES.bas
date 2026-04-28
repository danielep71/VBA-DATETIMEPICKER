Attribute VB_Name = "M_cPM_USAGE_EXAMPLES"
'==============================================================================
' MODULE: M_cPM_USAGE_EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Provides a compact, user-facing set of meaningful usage examples for
'   cPerformanceManager
'
' WHY THIS EXISTS
'   Once the demo workbook and regression suite exist, a large example module
'   becomes partly redundant. What remains genuinely useful is a smaller set of
'   examples that show:
'
'     - the normal recommended timing pattern
'     - how to format an already measured elapsed value
'     - how strict mode behaves on invalid usage
'     - how non-strict mode behaves on the same invalid usage
'     - how to benchmark with shared TW suppression
'     - how to structure real-world cleanup safely
'     - how to capture and export structured checkpoints
'
'   This module therefore keeps only the examples that still add real teaching
'   value beyond the demo sheets and tests
'
' INPUTS
'   None at module level
'
' RETURNS
'   None at module level
'
' BEHAVIOR
'   - Public launcher procedures run curated example groups
'   - Individual public procedures demonstrate one meaningful usage pattern
'   - Output is written primarily to the Immediate Window
'   - The checkpoint example can also write a structured report to a worksheet
'
' ERROR POLICY
'   Individual examples use local cleanup and re-raise normally unless the
'   example intentionally demonstrates expected invalid usage
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Excel Application object model
'
' NOTES
'   - Place this code in a STANDARD MODULE
'   - Results are printed primarily to the Immediate Window
'   - Press Ctrl+G in the VBA editor to open the Immediate Window
'   - Run worksheet-writing examples from a safe workbook / worksheet
'
' UPDATED
'   2026-04-18
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force explicit declaration of all variables

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const cPM_USAGE_SHEET_DATA           As String = "DATA_cPM"
    Private Const cPM_USAGE_CHECKPOINT_TOPLEFT   As String = "D3"

'
'==============================================================================
'
'                           PUBLIC: MASTER LAUNCHERS
'
'==============================================================================

Public Sub Run_All_UsageExamples()
'
'==============================================================================
'                          RUN ALL USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes all retained usage examples in a logical sequence
'
' WHY THIS EXISTS
'   Provides one compact entry point for reviewing the most meaningful usage
'   patterns without running the full demo workbook or regression suite
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Prints a module-level start banner
'   - Runs all retained example groups in teaching order
'   - Prints a module-level completion banner
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Run_CoreUsageExamples
'   - Run_ValidationUsageExamples
'   - Run_TimeWasterUsageExamples
'   - Run_SafePatternUsageExamples
'   - Run_CheckpointUsageExamples
'
' NOTES
'   This is the best launcher when you want the full compact walkthrough
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT MODULE START BANNER
'------------------------------------------------------------------------------
    'Print the overall suite start banner
        PrintModuleBanner "RUN ALL cPerformanceManager USAGE EXAMPLES - START"

'------------------------------------------------------------------------------
' RUN EXAMPLE GROUPS
'------------------------------------------------------------------------------
    'Run the core usage examples
        Run_CoreUsageExamples
    'Run the validation-behavior examples
        Run_ValidationUsageExamples
    'Run the TW-related usage examples
        Run_TimeWasterUsageExamples
    'Run the structured safe-pattern example
        Run_SafePatternUsageExamples
    'Run the checkpoint/report example
        Run_CheckpointUsageExamples

'------------------------------------------------------------------------------
' PRINT MODULE END BANNER
'------------------------------------------------------------------------------
    'Print the overall suite completion banner
        PrintModuleBanner "RUN ALL cPerformanceManager USAGE EXAMPLES - COMPLETE"

End Sub

Public Sub Run_CoreUsageExamples()
'
'==============================================================================
'                         RUN CORE USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the core, day-to-day usage demonstrations
'
' WHY THIS EXISTS
'   These are the most immediately useful examples for ordinary adoption of the
'   class in client code
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_BasicTiming_DefaultQPC
'     - Example_ElapsedTime_FromMeasuredSeconds
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_BasicTiming_DefaultQPC
'   - Example_ElapsedTime_FromMeasuredSeconds
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "CORE USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic default-QPC timing example
        PrintExampleBanner "Example_BasicTiming_DefaultQPC"
        Example_BasicTiming_DefaultQPC
    'Run the formatting-from-existing-seconds example
        PrintExampleBanner "Example_ElapsedTime_FromMeasuredSeconds"
        Example_ElapsedTime_FromMeasuredSeconds

End Sub

Public Sub Run_ValidationUsageExamples()
'
'==============================================================================
'                      RUN VALIDATION USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the examples that contrast strict and non-strict validation
'
' WHY THIS EXISTS
'   Correct session usage is one of the defining design choices of the class,
'   and these examples show how the caller can choose fail-fast or forgiving
'   behavior
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_StrictMode
'     - Example_NonStrictMode
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_StrictMode
'   - Example_NonStrictMode
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "VALIDATION USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the strict-mode example
        PrintExampleBanner "Example_StrictMode"
        Example_StrictMode
    'Run the non-strict-mode example
        PrintExampleBanner "Example_NonStrictMode"
        Example_NonStrictMode

End Sub

Public Sub Run_TimeWasterUsageExamples()
'
'==============================================================================
'                      RUN TIME-WASTER USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the TW-related benchmark usage example
'
' WHY THIS EXISTS
'   Shared Excel TW suppression is one of the most practically useful features
'   when benchmarking worksheet or application work
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_TimeWasters_Basic
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_TimeWasters_Basic
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "TIME-WASTER USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the basic TW-suppression example
        PrintExampleBanner "Example_TimeWasters_Basic"
        Example_TimeWasters_Basic

End Sub

Public Sub Run_SafePatternUsageExamples()
'
'==============================================================================
'                    RUN SAFE-PATTERN USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the recommended structured cleanup example
'
' WHY THIS EXISTS
'   This is the most important integration example when using the class in
'   real procedures
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_SafePattern
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_SafePattern
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "SAFE-PATTERN USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the structured cleanup pattern example
        PrintExampleBanner "Example_SafePattern"
        Example_SafePattern

End Sub

Public Sub Run_CheckpointUsageExamples()
'
'==============================================================================
'                     RUN CHECKPOINT USAGE EXAMPLES
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the structured checkpoint/reporting example
'
' WHY THIS EXISTS
'   Checkpoints are one of the newer high-value surfaces of the class and
'   deserve an explicit standalone example
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Runs:
'     - Example_CheckpointReport
'
' ERROR POLICY
'   Raises errors normally unless a called example handles errors internally
'
' DEPENDENCIES
'   - Example_CheckpointReport
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT SECTION BANNER
'------------------------------------------------------------------------------
    'Print the group banner
        PrintSectionBanner "CHECKPOINT USAGE EXAMPLES"

'------------------------------------------------------------------------------
' RUN EXAMPLES
'------------------------------------------------------------------------------
    'Run the checkpoint/report example
        PrintExampleBanner "Example_CheckpointReport"
        Example_CheckpointReport

End Sub

'
'==============================================================================
'
'                       PRIVATE: OUTPUT / BANNER HELPERS
'
'==============================================================================

Private Sub PrintModuleBanner( _
    ByVal Title As String)
'
'==============================================================================
'                           PRINT MODULE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a visually distinct top-level banner to the Immediate Window
'
' WHY THIS EXISTS
'   Improves readability when running groups of examples
'
' INPUTS
'   Title
'     Banner title text
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints a blank line, a delimiter, the title, and a closing delimiter
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print a blank line before the banner
        Debug.Print vbNullString
    'Print the opening delimiter
        Debug.Print String$(78, "=")
    'Print the banner title
        Debug.Print Title
    'Print the closing delimiter
        Debug.Print String$(78, "=")

End Sub

Private Sub PrintSectionBanner( _
    ByVal Title As String)
'
'==============================================================================
'                           PRINT SECTION BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a section-level banner to the Immediate Window
'
' WHY THIS EXISTS
'   Makes grouped example runs easier to read and review
'
' INPUTS
'   Title
'     Section title text
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints a blank line, a delimiter, the section title, and a closing
'   delimiter
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print a blank line before the section banner
        Debug.Print vbNullString
    'Print the opening delimiter
        Debug.Print String$(78, "-")
    'Print the section title
        Debug.Print Title
    'Print the closing delimiter
        Debug.Print String$(78, "-")

End Sub

Private Sub PrintExampleBanner( _
    ByVal ProcName As String)
'
'==============================================================================
'                           PRINT EXAMPLE BANNER
'------------------------------------------------------------------------------
' PURPOSE
'   Prints a small marker identifying the example about to run
'
' WHY THIS EXISTS
'   Helps the user associate Immediate Window output with the procedure that
'   produced it
'
' INPUTS
'   ProcName
'     Name of the example procedure being executed
'
' RETURNS
'   None
'
' BEHAVIOR
'   Prints one compact marker line
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' PRINT BANNER
'------------------------------------------------------------------------------
    'Print the example marker
        Debug.Print ">>> " & ProcName

End Sub

Private Function cPM_Usage_GetDataSheet() As Worksheet
'
'==============================================================================
'                           GET DATA SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the standard DATA_cPM worksheet used by the usage examples
'
' WHY THIS EXISTS
'   Centralizes the workbook-qualified worksheet lookup used by several
'   examples and avoids repeated unqualified Worksheets(...) calls
'
' INPUTS
'   None
'
' RETURNS
'   Worksheet
'     DATA_cPM worksheet from ThisWorkbook
'
' ERROR POLICY
'   Raises errors normally if the expected worksheet is missing
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the standard data sheet from the host workbook
        Set cPM_Usage_GetDataSheet = ThisWorkbook.Worksheets(cPM_USAGE_SHEET_DATA)

End Function

'
'==============================================================================
'
'                     PUBLIC: CORE TIMING USAGE EXAMPLES
'
'==============================================================================

Public Sub Example_BasicTiming_DefaultQPC()
'
'==============================================================================
'                       EXAMPLE: BASIC TIMING (DEFAULT QPC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the simplest recommended usage pattern for cPerformanceManager
'
' WHY THIS EXISTS
'   This is the canonical "getting started" example:
'     - instantiate the class
'     - start timing
'     - perform work
'     - read elapsed seconds
'     - clean up deterministically
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a new cPerformanceManager instance
'   - Starts timing with the default method, which is method 5 (QPC)
'   - Writes a constant into a worksheet range
'   - Reads numeric elapsed time in seconds
'   - Prints the result to the Immediate Window
'   - Restores environment state and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This is the preferred example to show the normal benchmark path
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ElapsedS            As Double                 'Elapsed time in seconds

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with the default method (5 = QPC)
        cPM.StartTimer

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Write a constant value into a worksheet range
        WS.Range("I6:I10006").Value = 7

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read the numeric elapsed time
        ElapsedS = cPM.ElapsedSeconds
    'Print the measured elapsed seconds
        Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Example_ElapsedTime_FromMeasuredSeconds()
'
'==============================================================================
'               EXAMPLE: ELAPSEDTIME FROM MEASURED SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how to format an already measured elapsed-seconds value
'
' WHY THIS EXISTS
'   In real code, callers often need both:
'     - a numeric elapsed value for comparisons / storage
'     - a formatted elapsed string for reporting
'
'   This example shows the recommended pattern:
'     - measure once with ElapsedSeconds
'     - format that same value through ElapsedTime(, ElapsedSecondsIn)
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Starts timing explicitly with method 5 (QPC)
'   - Performs a calculation workload
'   - Reads numeric elapsed seconds once
'   - Formats that existing value without taking a second timing sample
'   - Prints both results
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Application.Calculate
'   - Debug.Print
'
' NOTES
'   This is the recommended pattern when you want both numeric and
'   display-oriented output without double measurement
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim ElapsedS            As Double                 'Elapsed time in seconds
    Dim ElapsedT            As String                 'Formatted elapsed time

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing explicitly with QPC
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Force a workbook/application calculation pass
        Application.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read elapsed seconds once
        ElapsedS = cPM.ElapsedSeconds
    'Format the already measured elapsed-seconds value
        ElapsedT = cPM.ElapsedTime(, ElapsedS)
    'Print the numeric result
        Debug.Print "Elapsed seconds: " & Format$(ElapsedS, "0.000000000")
    'Print the formatted result
        Debug.Print "Elapsed time   : " & ElapsedT

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

'
'==============================================================================
'
'                      PUBLIC: VALIDATION USAGE EXAMPLES
'
'==============================================================================

Public Sub Example_StrictMode()
'
'==============================================================================
'                           EXAMPLE: STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how StrictMode enforces correct method/session usage
'
' WHY THIS EXISTS
'   A key design feature of cPerformanceManager is that elapsed reads are
'   session-bound. In strict mode, an invalid elapsed-method request raises an
'   error rather than being silently coerced
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Enables strict mode
'   - Starts a session using method 5
'   - Intentionally requests elapsed time with method 2
'   - Captures and prints the raised error
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Uses local expected-error handling to demonstrate the raised error safely
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'   - Err
'
' NOTES
'   This example is intentionally invalid by design
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim Dummy               As Double                 'Throwaway target for the failing call
    Dim ExpectedErrNum      As Long                   'Captured expected error number
    Dim ExpectedErrDesc     As String                 'Captured expected error description

    Dim SavedErrNumber      As Long                   'Captured unexpected error number
    Dim SavedErrSource      As String                 'Captured unexpected error source
    Dim SavedErrDescription As String                 'Captured unexpected error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Enable strict validation behavior
        cPM.StrictMode = True
    'Start timing with method 5
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' TRIGGER INTENTIONAL INVALID USAGE
'------------------------------------------------------------------------------
    'Switch to local expected-error handling
        On Error Resume Next
    'This is intentionally invalid because the active session uses method 5
        Dummy = cPM.ElapsedSeconds(2)
    'Capture the expected error information
        ExpectedErrNum = Err.Number
        ExpectedErrDesc = Err.Description
    'Clear the local expected-error state
        Err.Clear
    'Restore normal error handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' PRINT RESULT
'------------------------------------------------------------------------------
    'Print the captured expected error number
        Debug.Print "Error number: " & ExpectedErrNum
    'Print the captured expected error text
        Debug.Print "Error text  : " & ExpectedErrDesc

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original unexpected error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original unexpected error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

Public Sub Example_NonStrictMode()
'
'==============================================================================
'                         EXAMPLE: NON-STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates how the class behaves in non-strict mode when the caller asks
'   for an elapsed method that does not match the active session
'
' WHY THIS EXISTS
'   Non-strict mode is the forgiving mode of the class. Instead of raising, the
'   class may coerce the request and continue using the active session method
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Disables strict mode
'   - Starts a session with method 5
'   - Waits briefly
'   - Requests elapsed time using the wrong method identifier
'   - Prints the returned value and active method
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This example contrasts directly with Example_StrictMode
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim ElapsedS            As Double                 'Elapsed time returned after fallback

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE
'------------------------------------------------------------------------------
    'Disable strict validation behavior
        cPM.StrictMode = False
    'Start timing with method 5
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' APPLY SMALL DELAY
'------------------------------------------------------------------------------
    'Pause briefly to make the elapsed reading visible
        cPM.Pause 0.03, 1

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'In non-strict mode this falls back to the active session method
        ElapsedS = cPM.ElapsedSeconds(2)
    'Print the active method identifier
        Debug.Print "ActiveMethodID : " & cPM.ActiveMethodID
    'Print the active method name
        Debug.Print "MethodName     : " & cPM.MethodName(cPM.ActiveMethodID)
    'Print the returned elapsed time
        Debug.Print "ElapsedSeconds : " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

'
'==============================================================================
'
'               PUBLIC: SHARED TIME-WASTER SUPPRESSION EXAMPLES
'
'==============================================================================

Public Sub Example_TimeWasters_Basic()
'
'==============================================================================
'                     EXAMPLE: TIMEWASTERS (BASIC)
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates basic shared Excel "time-waster" suppression for a benchmark
'   run
'
' WHY THIS EXISTS
'   Excel application behaviors such as ScreenUpdating, events, alerts,
'   calculation mode, and cursor changes can add noise to benchmarks badly.
'   This example shows the normal suppression pattern
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Turns off all supported TW settings for this shared scope
'   - Starts QPC timing
'   - Performs a worksheet workload
'   - Prints elapsed seconds
'   - Ends the TW session
'   - Cleans up and releases the instance
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   TW control is shared/global in effect, so this relies on the shared
'   manager model rather than direct instance-local restoration
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ElapsedS            As Double                 'Measured elapsed seconds

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' SUPPRESS TIME-WASTERS
'------------------------------------------------------------------------------
    'Disable all supported TW settings for this instance's shared session
        cPM.TW_Turn_OFF

'------------------------------------------------------------------------------
' START TIMING
'------------------------------------------------------------------------------
    'Start timing with QPC
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a worksheet workload for benchmarking
        WS.Range("A1:A50000").Formula = "=ROW()"

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Read the numeric elapsed seconds
        ElapsedS = cPM.ElapsedSeconds
    'Print elapsed seconds with TW suppression in effect
        Debug.Print "Elapsed seconds with TW off: " & Format$(ElapsedS, "0.000000000")

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any environment changes made by this instance
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

'
'==============================================================================
'
'                     PUBLIC: RECOMMENDED SAFETY PATTERN
'
'==============================================================================

Public Sub Example_SafePattern()
'
'==============================================================================
'                         EXAMPLE: SAFE CLEANUP PATTERN
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates the recommended structured pattern for using
'   cPerformanceManager safely in real procedures
'
' WHY THIS EXISTS
'   Benchmarks and TW suppression can modify environment state. A structured
'   cleanup block ensures that environment restoration still happens when the
'   workload raises an error
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a timing manager instance
'   - Starts shared TW suppression
'   - Starts QPC timing
'   - Performs a workload
'   - Prints the formatted elapsed time
'   - Uses cleanup labels to ensure ResetEnvironment and object release happen
'     even if an error occurs
'
' ERROR POLICY
'   Uses structured local error handling with CleanFail / CleanExit labels
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'
' NOTES
'   This is the best example to follow when integrating the class into real
'   project code
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route runtime failures to the cleanup-aware failure block
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Create a new timing manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' CONFIGURE BENCHMARK ENVIRONMENT
'------------------------------------------------------------------------------
    'Start shared TW suppression for this instance
        cPM.TW_Turn_OFF
    'Start timing with QPC
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' APPLY WORKLOAD
'------------------------------------------------------------------------------
    'Execute a worksheet workload
        WS.UsedRange.Calculate

'------------------------------------------------------------------------------
' READ RESULT
'------------------------------------------------------------------------------
    'Print the formatted elapsed-time report
        Debug.Print "Elapsed: " & cPM.ElapsedTime

CleanExit:
'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release environment changes and the instance if the object exists
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
    'Exit normally after cleanup
        Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' FAILURE EXIT
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Print the error information for diagnostics
        Debug.Print "Error " & SavedErrNumber & " - " & SavedErrDescription
    'Always route through the normal cleanup block
        Resume CleanExit

End Sub

Public Sub Example_CheckpointReport()
'
'==============================================================================
'                       EXAMPLE: CHECKPOINT REPORT
'------------------------------------------------------------------------------
' PURPOSE
'   Demonstrates structured checkpoint capture and export
'
' WHY THIS EXISTS
'   Checkpoints are useful when one elapsed value is not enough and the caller
'   wants structured sub-measurements such as load / write / recalculate
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates a fresh performance manager instance
'   - Assigns a run label
'   - Starts a fresh QPC timing session
'   - Runs three simple benchmark phases
'   - Captures a checkpoint after each phase
'   - Prints the text report to the Immediate Window
'   - Writes the array report to DATA_cPM starting at D3
'
' ERROR POLICY
'   Restores the class environment before re-raising unexpected errors
'
' DEPENDENCIES
'   - cPerformanceManager
'   - Debug.Print
'   - ReportAsText
'   - ReportAsArray
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager    'Performance manager instance
    Dim WS                  As Worksheet              'Target worksheet
    Dim ReportArr           As Variant                'Structured report array
    Dim TopLeft             As Range                  'Top-left output anchor

    Dim SavedErrNumber      As Long                   'Captured error number
    Dim SavedErrSource      As String                 'Captured error source
    Dim SavedErrDescription As String                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Resolve the target worksheet
        Set WS = cPM_Usage_GetDataSheet()
    'Resolve the top-left report output anchor
        Set TopLeft = WS.Range(cPM_USAGE_CHECKPOINT_TOPLEFT)
    'Create a fresh performance manager instance
        Set cPM = New cPerformanceManager

'------------------------------------------------------------------------------
' INITIALIZE TIMING SESSION
'------------------------------------------------------------------------------
    'Assign a run label
        cPM.SetRunLabel "ImportWorkflow"
    'Start a fresh timing session
        cPM.StartTimer 5, False

'------------------------------------------------------------------------------
' CHECKPOINTED WORKLOAD
'------------------------------------------------------------------------------
    'Simulate phase 1
        WS.Range("A1:A10000").Value = 1
    'Capture the first checkpoint
        cPM.Checkpoint "LoadValues"
    'Simulate phase 2
        WS.Range("B1:B10000").Formula = "=ROW()"
    'Capture the second checkpoint
        cPM.Checkpoint "WriteFormulas"
    'Simulate phase 3
        Application.Calculate
    'Capture the third checkpoint
        cPM.Checkpoint "Recalculate", "Full workbook calculation pass"

'------------------------------------------------------------------------------
' OUTPUT TO IMMEDIATE WINDOW
'------------------------------------------------------------------------------
    'Print the readable checkpoint report
        Debug.Print cPM.ReportAsText

'------------------------------------------------------------------------------
' OUTPUT TO WORKSHEET
'------------------------------------------------------------------------------
    'Resolve the structured report array
        ReportArr = cPM.ReportAsArray
    'Clear the target write area sized to the current report
        TopLeft.Resize(UBound(ReportArr, 1), UBound(ReportArr, 2)).ClearContents
    'Write the structured report array to the worksheet
        TopLeft.Resize(UBound(ReportArr, 1), UBound(ReportArr, 2)).Value = ReportArr

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Route through centralized cleanup
        Resume CleanExit

End Sub

