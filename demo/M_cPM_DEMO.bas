Attribute VB_Name = "M_cPM_DEMO"
'==============================================================================
' MODULE: M_cPM_DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the cPerformanceManager demo workbook layout and provides the
'   executable demo actions used by that layout
'
' WHY THIS EXISTS
'   cPerformanceManager is easier to understand and validate when users can:
'
'     - configure benchmark parameters from a visible control panel
'     - run predefined actions from dedicated buttons
'     - inspect structured timing outputs in a results log
'     - use a dedicated DATA sheet for configured workload anchors
'     - consult an embedded HELP sheet without leaving the workbook
'
'   This module centralizes both:
'
'     - sheet-building logic
'     - demo-action logic
'
'   while keeping the implementation in a single module
'
' WHAT THIS MODULE DOES
'   - Rebuilds DEMO_cPM, DATA_cPM, and HELP_cPM
'   - Creates the control panel, action buttons, and log table
'   - Creates the DATA sheet workload-anchor setup
'   - Creates the HELP sheet guidance sections
'   - Runs the demo actions used by the buttons
'   - Appends structured log rows to the demo log table
'
' DESIGN NOTES
'   - Generic sheet creation / reset / template formatting is delegated to
'     M_DEMO_BUILDER
'   - This module keeps one coherent cPM-specific surface in a single place
'   - Target ranges are configured through workbook-level names that store the
'     text address of the workload anchor cell
'   - Effective workload height is controlled by cPM_RangeRows
'   - All demo actions release environment changes deterministically through
'     ResetEnvironment
'
' ENTRY POINTS
'   - cPM_CreateDemoSheets
'   - cPM_RunBasicTiming
'   - cPM_RunAllMethods
'   - cPM_RunAlignedDemo
'   - cPM_RunDiagnostics
'   - cPM_RunOverheadDemo
'   - cPM_RunPauseDemo
'   - cPM_RunTWComparison
'   - cPM_ClearResultsLog
'
' DEPENDENCIES
'   - cPerformanceManager
'   - M_DEMO_BUILDER
'   - M_cPM_TimeWasters
'
' NOTES
'   - Place this code in a STANDARD MODULE
'   - Public routines are intentionally hidden from Excel's Macro dialog via
'     Option Private Module
'
' UPDATED
'   2026-04-15
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force all variable declarations
    Option Private Module   'Public routines are not visible from Excel

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const cPM_SHEET_DEMO            As String = "DEMO_cPM"
    Private Const cPM_SHEET_DATA            As String = "DATA_cPM"
    Private Const cPM_SHEET_HELP            As String = "HELP_cPM"

    Private Const cPM_TABLE_LOG             As String = "Tbl_cPM_Log"

    Private Const cPM_DEMO_TITLE            As String = "CLASS PERFORMANCE MANAGER"

    Private Const cPM_SUBTITLE_DEMO         As String = _
                    "Demo Sheet: Interactive control panel, run buttons, and structured results log"

    Private Const cPM_SUBTITLE_DATA         As String = _
                    "Data Sheet: Sample areas for worksheet-write, formula-fill, and calculation demonstrations"

    Private Const cPM_SUBTITLE_HELP         As String = _
                    "Help Sheet: Quick in-workbook guidance for the demo layout"

    Private Const cPM_NAME_METHOD           As String = "cPM_MethodID"
    Private Const cPM_NAME_ALIGN            As String = "cPM_AlignToNextTick"
    Private Const cPM_NAME_STRICT           As String = "cPM_StrictMode"
    Private Const cPM_NAME_PAUSESEC         As String = "cPM_PauseSeconds"
    Private Const cPM_NAME_PAUSEMETH        As String = "cPM_PauseMethod"
    Private Const cPM_NAME_ITER             As String = "cPM_Iterations"
    Private Const cPM_NAME_ROWS             As String = "cPM_RangeRows"
    Private Const cPM_NAME_TWMODE           As String = "cPM_TWMode"

    Private Const cPM_NAME_VALUEFILL        As String = "cPM_ValueFill"
    Private Const cPM_NAME_VALUEFILLVALUE   As String = "cPM_ValueFillValue"
    Private Const cPM_NAME_FORMULAFILL      As String = "cPM_FormulaFill"
    Private Const cPM_NAME_FORMULAFILLEXPR  As String = "cPM_FormulaFillExpr"

Public Sub cPM_CreateDemoSheets()
'
'==============================================================================
'                           CREATE DEMO SHEETS
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or rebuilds the demo sheets used to showcase cPerformanceManager
'
' WHY THIS EXISTS
'   This routine provides a single orchestration entry point for rebuilding the
'   cPerformanceManager demo environment using the shared demo-builder
'   infrastructure together with the three sheet-specific build routines in
'   this module
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Delegates generic sheet preparation to DEMO_Build_DemoTemplate
'   - Resolves the prepared worksheets from ThisWorkbook
'   - Builds the sheet-specific content for DEMO_cPM, DATA_cPM, and HELP_cPM
'   - Activates the main demo sheet at the end
'
' ERROR POLICY
'   Restores Application state and then re-raises the original error
'
' DEPENDENCIES
'   - DEMO_Build_DemoTemplate
'   - DEMO_Begin_FastMode
'   - DEMO_End_FastMode
'   - cPM_BuildDemoSheet
'   - cPM_BuildDataSheet
'   - cPM_BuildHelpSheet
'
' NOTES
'   Generic worksheet lifecycle concerns such as create / get / reset are
'   delegated to DEMO_Build_DemoTemplate and are therefore not repeated here
'
'   This routine rebuilds the three demo sheets and then activates DEMO_cPM
'   as the user-facing landing sheet
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WB                  As Workbook             'Target workbook
    Dim WS_Demo             As Worksheet            'Demo / control sheet
    Dim WS_Data             As Worksheet            'Data / benchmark sheet
    Dim WS_Help             As Worksheet            'Embedded help sheet

    Dim FastModeState       As tDemoFastModeState   'Saved Application-state snapshot
    Dim FastModeOn          As Boolean              'TRUE when fast mode was entered

    Dim SavedErrNumber      As Long                 'Captured error number
    Dim SavedErrSource      As String               'Captured error source
    Dim SavedErrDescription As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo Clean_Fail
    'Target the workbook that contains this module
        Set WB = ThisWorkbook
    'Capture and apply fast-mode Application settings
        DEMO_Begin_FastMode FastModeState
        FastModeOn = True
    'Show the wait cursor while rebuilding the demo workbook
        Application.Cursor = xlWait

'------------------------------------------------------------------------------
' PREPARE GENERIC SHEETS
'------------------------------------------------------------------------------
    'Build or rebuild the generic template for the main demo sheet
        DEMO_Build_DemoTemplate _
            cPM_SHEET_DEMO, _
            cPM_DEMO_TITLE, _
            cPM_SUBTITLE_DEMO, , 17
    'Build or rebuild the generic template for the data sheet
        DEMO_Build_DemoTemplate _
            cPM_SHEET_DATA, _
            cPM_DEMO_TITLE, _
            cPM_SUBTITLE_DATA
    'Build or rebuild the generic template for the help sheet
        DEMO_Build_DemoTemplate _
            cPM_SHEET_HELP, _
            cPM_DEMO_TITLE, _
            cPM_SUBTITLE_HELP

'------------------------------------------------------------------------------
' RESOLVE PREPARED SHEETS
'------------------------------------------------------------------------------
    'Resolve the main demo sheet after template preparation
        Set WS_Demo = WB.Worksheets(cPM_SHEET_DEMO)
    'Resolve the data sheet after template preparation
        Set WS_Data = WB.Worksheets(cPM_SHEET_DATA)
    'Resolve the help sheet after template preparation
        Set WS_Help = WB.Worksheets(cPM_SHEET_HELP)

'------------------------------------------------------------------------------
' BUILD SHEET-SPECIFIC CONTENT
'------------------------------------------------------------------------------
    'Build the main demo / control sheet content
        cPM_BuildDemoSheet WB, WS_Demo
    'Build the supporting data / benchmark sheet content
        cPM_BuildDataSheet WB, WS_Data
    'Build the embedded help / documentation sheet content
        cPM_BuildHelpSheet WS_Help

Clean_Exit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Restore the normal cursor
        Application.Cursor = xlDefault
    'Restore the original Excel Application state only when fast mode was entered
        If FastModeOn Then
            DEMO_End_FastMode FastModeState
        End If
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    Exit Sub

Clean_Fail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Capture the original error details before cleanup
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description
    'Continue through the centralized cleanup path
        Resume Clean_Exit

End Sub
Private Sub cPM_BuildDemoSheet( _
    ByVal WB As Workbook, _
    ByVal WS As Worksheet)
'
'==============================================================================
'                            BUILD DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the DEMO_cPM worksheet
'
' WHY THIS EXISTS
'   This sheet acts as the visible control panel and results dashboard for the
'   cPerformanceManager demo workbook
'
' INPUTS
'   WB
'     Target workbook
'
'   WS
'     DEMO worksheet to populate
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Builds the control panel
'   - Builds the action-button area
'   - Builds the structured results-log table
'   - Applies lightweight number-format refinements to the results log
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Prepare_LabeledInputSection
'   - DEMO_Write_NamedInputRow
'   - DEMO_Write_BandHeader
'   - DEMO_Add_ButtonGrid
'   - DEMO_Set_RangeBorder
'   - DEMO_Create_TableSection
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lo                  As ListObject       'Results log table
    Dim ButtonSpecs         As Variant          'Button name / caption / macro specification
    Dim LogHeaders          As Variant          'Results-log column captions

'------------------------------------------------------------------------------
' ADJUST COLUMNS
'------------------------------------------------------------------------------
    'Adjust the main visible columns for the demo sheet layout
        WS.Columns("C:M").ColumnWidth = 18
    'Widen the elapsed-time text column
        WS.Columns("I").ColumnWidth = 50
    'Widen the notes column
        WS.Columns("M").ColumnWidth = 40

'------------------------------------------------------------------------------
' BUILD CONTROL PANEL
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' PREPARE SECTION FRAME
    '--------------------------------------------------------------------------
        'Apply the standard section / label / input formatting
            DEMO_Prepare_LabeledInputSection _
                WS, _
                WS.Range("C4:D4"), _
                "CONTROL PANEL", _
                WS.Range("C5:C12"), _
                WS.Range("D5:D12")

    '--------------------------------------------------------------------------
    ' WRITE INPUT ROWS
    '--------------------------------------------------------------------------
        'Write the timing-method input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C5"), WS.Range("D5"), _
                "Timing Method ID", 5, _
                cPM_NAME_METHOD, _
                DemoInputValidationList, _
                "1,2,3,4,5,6"

        'Write the align-to-next-tick input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C6"), WS.Range("D6"), _
                "AlignToNextTick", "FALSE", _
                cPM_NAME_ALIGN, _
                DemoInputValidationBoolean

        'Write the strict-mode input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C7"), WS.Range("D7"), _
                "StrictMode", "TRUE", _
                cPM_NAME_STRICT, _
                DemoInputValidationBoolean

        'Write the pause-seconds input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C8"), WS.Range("D8"), _
                "Pause Seconds", 0.25, _
                cPM_NAME_PAUSESEC, _
                DemoInputValidationNumeric, _
                "", _
                0, 100, _
                True, _
                "0.00"

        'Write the pause-method input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C9"), WS.Range("D9"), _
                "Pause Method", 1, _
                cPM_NAME_PAUSEMETH, _
                DemoInputValidationList, _
                "1,2,3,4"

        'Write the iterations input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C10"), WS.Range("D10"), _
                "Iterations", 100, _
                cPM_NAME_ITER, _
                DemoInputValidationNumeric, _
                "", _
                0, 100000, _
                False, _
                "#,##0"

        'Write the range-rows input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C11"), WS.Range("D11"), _
                "Range Rows", 100000, _
                cPM_NAME_ROWS, _
                DemoInputValidationNumeric, _
                "", _
                0, 1000000, _
                False, _
                "#,##0"

        'Write the time-waster mode input row
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C12"), WS.Range("D12"), _
                "TW Mode", "All Off", _
                cPM_NAME_TWMODE, _
                DemoInputValidationList, _
                "None,All Off,Keep ScreenUpdating,Keep EnableEvents," & _
                "Keep ScreenUpdating+EnableEvents"

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the full control-panel area
            DEMO_Set_RangeBorder WS.Range("C4:D12")

'------------------------------------------------------------------------------
' BUILD ACTION BUTTONS
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' PREPARE ACTION SECTION
    '--------------------------------------------------------------------------
        'Write the actions section header band
            DEMO_Write_BandHeader WS.Range("J4:L4"), "DEMO ACTIONS"

    '--------------------------------------------------------------------------
    ' DEFINE BUTTON GRID
    '--------------------------------------------------------------------------
        'Define the button grid specification
            ButtonSpecs = Array( _
                Array("btn_cPM_Basic", "Run Basic Timing", "cPM_RunBasicTiming"), _
                Array("btn_cPM_AllMethods", "Run All Methods", "cPM_RunAllMethods"), _
                Array("btn_cPM_Aligned", "Run Aligned Demo", "cPM_RunAlignedDemo"), _
                Array("btn_cPM_Diagnostics", "Run Diagnostics", "cPM_RunDiagnostics"), _
                Array("btn_cPM_Overhead", "Run Overhead Demo", "cPM_RunOverheadDemo"), _
                Array("btn_cPM_Pause", "Run Pause Demo", "cPM_RunPauseDemo"), _
                Array("btn_cPM_TW", "Run TW Comparison", "cPM_RunTWComparison"), _
                Array("btn_cPM_ClearLog", "Clear Results Log", "cPM_ClearResultsLog"))

    '--------------------------------------------------------------------------
    ' CREATE BUTTONS
    '--------------------------------------------------------------------------
        'Create the standard two-column demo button grid
            DEMO_Add_ButtonGrid _
                WS, _
                WS.Range("J5"), _
                ButtonSpecs, _
                2, _
                135, _
                25, _
                18, _
                15, _
                4, _
                5

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the full action-button area
            DEMO_Set_RangeBorder WS.Range("J4:L12")
    
    '--------------------------------------------------------------------------
    ' CREATE REGRESSION TEST BUTTON
    '--------------------------------------------------------------------------
        DEMO_Add_DemoButton WS, "Btn_cPM_Regression", "Run Regression Tests", _
            WS.Range("M8").Left + 40, WS.Range("I8").Top + 10, _
            135, 25, "Run_cPerformanceManager_RegressionSuite"
    
'------------------------------------------------------------------------------
' BUILD RESULTS LOG
'------------------------------------------------------------------------------
    'Define the ordered results-log headers
        LogHeaders = Array( _
            "Run ID", _
            "Timestamp", _
            "Scenario", _
            "Method ID", _
            "Method Name", _
            "Elapsed Seconds", _
            "Elapsed Time", _
            "T1", _
            "T2", _
            "ET", _
            "Notes")

    'Create the titled results-log table section
        Set Lo = DEMO_Create_TableSection( _
                    WS, _
                    WS.Range("C15:M15"), _
                    "RESULTS LOG", _
                    WS.Range("C16"), _
                    LogHeaders, _
                    cPM_TABLE_LOG, _
                    "TableStyleMedium6")

'------------------------------------------------------------------------------
' APPLY TABLE FORMATTING
'------------------------------------------------------------------------------
    'Apply lightweight number formatting to the created log table
        On Error Resume Next

        'Format the run identifier column as an integer
            Lo.ListColumns("Run ID").DataBodyRange.NumberFormat = "0"
        'Format the timestamp column as date + time
            Lo.ListColumns("Timestamp").DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm:ss"
        'Format the method identifier column as an integer
            Lo.ListColumns("Method ID").DataBodyRange.NumberFormat = "0"
        'Format elapsed / timer metrics with higher precision
            Lo.ListColumns("Elapsed Seconds").DataBodyRange.NumberFormat = "0.000000"
            Lo.ListColumns("T1").DataBodyRange.NumberFormat = "0.000000"
            Lo.ListColumns("T2").DataBodyRange.NumberFormat = "0.000000"
            Lo.ListColumns("ET").DataBodyRange.NumberFormat = "0.000000"

        On Error GoTo 0

End Sub


Private Sub cPM_BuildDataSheet( _
    ByVal WB As Workbook, _
    ByVal WS As Worksheet)
'
'==============================================================================
'                            BUILD DATA SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the DATA_cPM worksheet
'
' WHY THIS EXISTS
'   This sheet provides pre-positioned workload anchors and benchmark-oriented
'   worksheet areas for cPerformanceManager demos
'
' INPUTS
'   WB
'     Target workbook
'
'   WS
'     DATA worksheet to populate
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Builds the value-fill setup panel
'   - Builds the formula-fill setup panel
'   - Builds a visible benchmark-target descriptor area
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Prepare_LabeledInputSection
'   - DEMO_Write_NamedInputRow
'   - DEMO_Write_BandHeader
'   - DEMO_Set_RangeBorder
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ADDITIONAL FORMATTING
'------------------------------------------------------------------------------
    'Adjust a few specific worksheet columns for the benchmark layout
        WS.Columns("E").ColumnWidth = 5
        WS.Columns("H").ColumnWidth = 5
        WS.Columns("I:J").HorizontalAlignment = xlCenter

'------------------------------------------------------------------------------
' BUILD VALUE-FILL TEST PANEL
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' PREPARE SECTION FRAME
    '--------------------------------------------------------------------------
        'Apply the standard section / label / input formatting
            DEMO_Prepare_LabeledInputSection _
                WS, _
                WS.Range("C4:D4"), _
                "VALUE FILL TEST", _
                WS.Range("C5:C6"), _
                WS.Range("D5:D6")

    '--------------------------------------------------------------------------
    ' WRITE INPUT ROWS
    '--------------------------------------------------------------------------
        'Write the target-anchor input row for the value-fill scenario
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C5"), WS.Range("D5"), _
                "Target Range", "I6", _
                cPM_NAME_VALUEFILL

        'Write the fill-value input row for the value-fill scenario
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("C6"), WS.Range("D6"), _
                "Fill Value", 1, _
                cPM_NAME_VALUEFILLVALUE, _
                DemoInputValidationNumeric, _
                "", _
                -1000000000, 1000000000, _
                True, _
                "0.00"

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the full value-fill panel
            DEMO_Set_RangeBorder WS.Range("C4:D6")

'------------------------------------------------------------------------------
' BUILD FORMULA-FILL TEST PANEL
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' PREPARE SECTION FRAME
    '--------------------------------------------------------------------------
        'Apply the standard section / label / input formatting
            DEMO_Prepare_LabeledInputSection _
                WS, _
                WS.Range("F4:G4"), _
                "FORMULA FILL TEST", _
                WS.Range("F5:F6"), _
                WS.Range("G5:G6")

    '--------------------------------------------------------------------------
    ' WRITE INPUT ROWS
    '--------------------------------------------------------------------------
        'Write the target-anchor input row for the formula-fill scenario
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("F5"), WS.Range("G5"), _
                "Target Range", "J6", _
                cPM_NAME_FORMULAFILL

        'Force the formula-expression cell to text before writing the default
            WS.Range("G6").NumberFormat = "@"

        'Write the formula-text input row for the formula-fill scenario
            DEMO_Write_NamedInputRow _
                WB, WS, _
                WS.Range("F6"), WS.Range("G6"), _
                "Formula", "=RAND()", _
                cPM_NAME_FORMULAFILLEXPR

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the full formula-fill panel
            DEMO_Set_RangeBorder WS.Range("F4:G6")

'------------------------------------------------------------------------------
' BUILD BENCHMARK TARGET AREA
'------------------------------------------------------------------------------
    '--------------------------------------------------------------------------
    ' WRITE SECTION HEADER
    '--------------------------------------------------------------------------
        'Write the benchmark target-area section header
            DEMO_Write_BandHeader WS.Range("I4:J4"), "BENCHMARK TARGET AREA"

    '--------------------------------------------------------------------------
    ' WRITE COLUMN CAPTIONS
    '--------------------------------------------------------------------------
        'Write the value-fill target caption
            WS.Range("I5").Value = "Value Fill Target"
        'Write the formula-fill target caption
            WS.Range("J5").Value = "Formula Fill Target"

    '--------------------------------------------------------------------------
    ' FORMAT HEADER ROW
    '--------------------------------------------------------------------------
        'Apply standard header styling to the benchmark-area captions
            With WS.Range("I5:J5")
                .Interior.Color = COLOR_SUBHEADER
                .Font.Bold = True
                .Font.Color = vbWhite
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With

    '--------------------------------------------------------------------------
    ' APPLY VISUAL FRAME
    '--------------------------------------------------------------------------
        'Apply a border around the visible benchmark-area descriptor block
            DEMO_Set_RangeBorder WS.Range("I4:J5")

End Sub


Private Sub cPM_BuildHelpSheet(ByVal WS As Worksheet)
'
'==============================================================================
'                            BUILD HELP SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the HELP_cPM worksheet
'
' WHY THIS EXISTS
'   A demo workbook is easier to use when it contains short embedded guidance
'   rather than assuming users will read external documentation first
'
' INPUTS
'   WS
'     HELP worksheet to populate
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates titled help sections
'   - Explains what each workbook sheet contains
'   - Summarizes the recommended first steps
'   - Documents the timing-method map
'   - Documents the supported pause methods
'   - Notes the meaning of TW mode
'   - Records a few practical reminders
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - DEMO_Write_BandHeader
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' WHAT THIS WORKBOOK CONTAINS
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C4:D4"), "WHAT THIS WORKBOOK CONTAINS", , , , , , False

    'Write the workbook-overview guidance lines
        WS.Range("D5").Value = "DEMO_cPM contains the control panel, action buttons, and the results log."
        WS.Range("D6").Value = "DATA_cPM contains worksheet areas suitable for write / fill / calculate demonstrations."
        WS.Range("D7").Value = "HELP_cPM contains a short embedded guide."

'------------------------------------------------------------------------------
' RECOMMENDED FIRST STEPS
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C9:D9"), "RECOMMENDED FIRST STEPS", , , , , , False

    'Write the recommended-first-step guidance lines
        WS.Range("D10").Value = "1. Go to DEMO_cPM."
        WS.Range("D11").Value = "2. Leave Method ID = 5 unless you are explicitly comparing methods."
        WS.Range("D12").Value = "3. Use the control panel to adjust AlignToNextTick, StrictMode, Pause Seconds, Pause Method, and TW Mode."
        WS.Range("D13").Value = "4. Use the demo buttons to run the benchmark actions."

'------------------------------------------------------------------------------
' METHOD MAP
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C15:D15"), "METHOD MAP", , , , , , False

    'Write the timer-method map lines
        WS.Range("D16").Value = "1 = Timer"
        WS.Range("D17").Value = "2 = GetTickCount / GetTickCount64"
        WS.Range("D18").Value = "3 = timeGetTime"
        WS.Range("D19").Value = "4 = timeGetSystemTime"
        WS.Range("D20").Value = "5 = QPC (recommended default)"
        WS.Range("D21").Value = "6 = Now() * 86400"

'------------------------------------------------------------------------------
' PAUSE METHODS
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C23:D23"), "PAUSE METHODS", , , , , , False

    'Write the pause-method guidance lines
        WS.Range("D24").Value = "1 = Sleep API: millisecond-style kernel wait; lowest overhead; does not yield via DoEvents."
        WS.Range("D25").Value = "2 = Timer + DoEvents loop: coarser wait; yields to Excel / UI; higher overhead."
        WS.Range("D26").Value = "3 = Application.Wait: whole-second granularity; simple and deterministic; may overshoot slightly."
        WS.Range("D27").Value = "4 = Now + DoEvents loop: Date-based yielding wait; useful as a fallback / comparison path."

'------------------------------------------------------------------------------
' TW MODE
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C29:D29"), "TW MODE", , , , , , False

    'Write the TW guidance lines
        WS.Range("D30").Value = "TW means Time-Wasters suppression."
        WS.Range("D31").Value = "The demo control supports modes such as All Off and selected exemptions."
        WS.Range("D32").Value = "The current project design requires the companion module M_cPM_TimeWasters.bas."

'------------------------------------------------------------------------------
' IMPORTANT NOTES
'------------------------------------------------------------------------------
    'Write the section header
        DEMO_Write_BandHeader WS.Range("C34:D34"), "IMPORTANT NOTES", , , , , , False

    'Write the practical reminder lines
        WS.Range("D35").Value = "Use method 5 for the main benchmark path."
        WS.Range("D36").Value = "Use ElapsedSeconds for numeric comparisons and ElapsedTime for readable output."
        WS.Range("D37").Value = "Always call ResetEnvironment in real demo / benchmark macros."

End Sub

Public Sub cPM_RunBasicTiming()
'
'==============================================================================
'                             RUN BASIC TIMING
'------------------------------------------------------------------------------
' PURPOSE
'   Runs one basic worksheet-write timing scenario using the current control
'   panel settings
'
' WHY THIS EXISTS
'   This is the simplest end-to-end demo path:
'     - read the current timing controls
'     - resolve the configured value-fill target
'     - write the configured fill value
'     - repeat the workload for the configured number of iterations
'     - log the measured result
'
'   This version also supports optional in-loop progress in the Excel status
'   bar, enabled only when the iteration count is high enough to justify it
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected method and alignment flag
'   - Resolves the configured value-fill target range
'   - Reads the configured fill value
'   - Reads the configured iteration count
'   - Executes the timed value-fill workload repeatedly
'   - Updates the status bar periodically during the timed loop only when
'     Iterations >= 100
'   - Appends one log row
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising the
'   original error
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetValueTargetRange
'   - cPM_Demo_GetValueFillValue
'   - cPM_Demo_GetIterations
'   - cPM_Demo_ApplyTWMode
'   - cPM_Demo_GetTWMode
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'
' NOTES
'   Showing progress inside the timed loop adds UI overhead and therefore
'   slightly contaminates the measurement. To reduce that effect:
'     - in-loop progress is shown only when Iterations >= 100
'     - the status bar is updated periodically rather than on every iteration
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance
    Dim MethodID            As Integer               'Selected timing method
    Dim AlignFlag           As Boolean               'Selected alignment flag
    Dim Target              As Range                 'Worksheet target range
    Dim FillValue           As Variant               'Configured fill value
    Dim Iterations          As Long                  'Configured iteration count
    Dim TWMode              As String                'Selected TW mode
    Dim i                   As Long                  'Loop counter
    Dim ElapsedS            As Double                'Measured elapsed seconds
    Dim ElapsedTxt          As String                'Formatted elapsed-time text
    Dim Notes               As String                'Scenario note text
    Dim ProgressEvery       As Long                  'Status-bar update frequency
    Dim ProgressEnabled     As Boolean               'TRUE when in-loop progress should be shown

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Read the selected timing method
        MethodID = cPM_Demo_GetMethodID()
    'Read the selected alignment flag
        AlignFlag = cPM_Demo_GetAlignFlag()
    'Resolve the selected TW mode once
        TWMode = cPM_Demo_GetTWMode()
    'Prepare the timer instance from the current control-panel settings
        Set cPM = cPM_Demo_PrepareInstance()
    'Resolve the target worksheet range for the value-fill workload
        Set Target = cPM_Demo_GetValueTargetRange()
    'Read the configured fill value
        FillValue = cPM_Demo_GetValueFillValue()
    'Read the configured iteration count
        Iterations = cPM_Demo_GetIterations()

'------------------------------------------------------------------------------
' RESOLVE PROGRESS POLICY
'------------------------------------------------------------------------------
    'Enable in-loop progress only for sufficiently large iteration counts
        ProgressEnabled = (Iterations >= 100)

    'Resolve the status-bar update interval only when progress is enabled
        If ProgressEnabled Then
            ProgressEvery = Iterations \ 100
            If ProgressEvery < 1 Then
                ProgressEvery = 1
            End If
        End If

'------------------------------------------------------------------------------
' APPLY OPTIONAL TW MODE
'------------------------------------------------------------------------------
    'Apply the selected TW mode before the measurement starts
        cPM_Demo_ApplyTWMode cPM, TWMode

'------------------------------------------------------------------------------
' PREPARE WORKLOAD
'------------------------------------------------------------------------------
    'Clear the target range before starting the timed block
        Target.ClearContents
    'Write an initial status-bar message before timing starts when enabled
        If ProgressEnabled Then
            Demo_SB_SetProgress 0, Iterations, "Basic Timing - " & cPM.MethodName(MethodID)
        End If

'------------------------------------------------------------------------------
' RUN WORKLOAD
'------------------------------------------------------------------------------
    'Start the timing session using the selected method and alignment
        cPM.StartTimer MethodID, AlignFlag
    'Repeat the value-fill workload for the configured number of iterations
        For i = 1 To Iterations
            'Run the value-fill workload
                Target.Value = FillValue
            'Update the status bar periodically during the timed loop when enabled
                If ProgressEnabled Then
                    If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                        Demo_SB_SetProgress i, Iterations, "Basic Timing - " & cPM.MethodName(MethodID)
                    End If
                End If
        Next i

'------------------------------------------------------------------------------
' READ RESULTS
'------------------------------------------------------------------------------
    'Read the numeric elapsed time
        ElapsedS = cPM.ElapsedSeconds()
    'Read the formatted elapsed time without taking a second timing sample
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
    'Build the scenario note text
        Notes = "Rows=" & CStr(Target.Rows.Count) & _
                " | Iterations=" & CStr(Iterations) & _
                " | Align=" & CStr(AlignFlag) & _
                " | FillValue=" & CStr(FillValue) & _
                " | TW=" & TWMode

'------------------------------------------------------------------------------
' LOG RESULT
'------------------------------------------------------------------------------
    'Append one log row for the measured workload
        cPM_Demo_AppendLog "Basic Timing", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           Notes

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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
    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub
Public Sub cPM_RunAllMethods()
'
'==============================================================================
'                             RUN ALL METHODS
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the same worksheet-write workload across all six timing backends
'
' WHY THIS EXISTS
'   This routine provides a side-by-side comparison of the supported timing
'   backends under the same workload and current control settings
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected alignment flag once
'   - Resolves the configured value-fill target range once
'   - Reads the configured fill value once
'   - Reads the configured iteration count once
'   - Reads the selected TW mode once
'   - Loops through timing methods 1 to 6
'   - Updates the Excel status bar with method-level progress
'   - Optionally updates the Excel status bar inside the iterations loop when
'     Iterations >= 100
'   - Executes the timed workload repeatedly for each method
'   - Appends one log row per method
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising errors
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetAlignFlag
'   - cPM_Demo_GetValueTargetRange
'   - cPM_Demo_GetValueFillValue
'   - cPM_Demo_GetIterations
'   - cPM_Demo_GetTWMode
'   - cPM_Demo_ApplyTWMode
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'   - cPM_Demo_MethodName
'
' NOTES
'   Showing progress inside the timed loop adds UI overhead and therefore
'   slightly contaminates the measurement. To reduce that effect:
'     - in-loop progress is shown only when Iterations >= 100
'     - the status bar is updated periodically rather than on every iteration
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description

    Dim MethodID            As Integer               'Looped timing method
    Dim AlignFlag           As Boolean               'Selected alignment flag
    Dim Target              As Range                 'Worksheet target range
    Dim FillValue           As Variant               'Configured fill value
    Dim Iterations          As Long                  'Configured iteration count
    Dim TWMode              As String                'Selected TW mode
    Dim i                   As Long                  'Workload loop counter
    Dim ElapsedS            As Double                'Measured elapsed seconds
    Dim ElapsedTxt          As String                'Formatted elapsed-time text
    Dim Notes               As String                'Scenario note text

    Dim CurrentStep         As Long                  'Current method-level progress step
    Dim ProgressEnabled     As Boolean               'TRUE when in-loop progress should be shown
    Dim ProgressEvery       As Long                  'In-loop status-bar update frequency

    Const TotalSteps As Long = 6                     'Total number of timing methods

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Read the selected alignment flag once
        AlignFlag = cPM_Demo_GetAlignFlag()
    'Resolve the target worksheet range once
        Set Target = cPM_Demo_GetValueTargetRange()
    'Read the configured fill value once
        FillValue = cPM_Demo_GetValueFillValue()
    'Read the configured iteration count once
        Iterations = cPM_Demo_GetIterations()
    'Read the selected TW mode once
        TWMode = cPM_Demo_GetTWMode()

'------------------------------------------------------------------------------
' RESOLVE PROGRESS POLICY
'------------------------------------------------------------------------------
    'Enable in-loop progress only for sufficiently large iteration counts
        ProgressEnabled = (Iterations >= 100)
    'Resolve the in-loop update interval only when progress is enabled
        If ProgressEnabled Then
            ProgressEvery = Iterations \ 100
            If ProgressEvery < 1 Then
                ProgressEvery = 1
            End If
        End If
    'Initialize method-level progress
        CurrentStep = 0
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Preparing all-methods comparison"

'------------------------------------------------------------------------------
' RUN WORKLOADS
'------------------------------------------------------------------------------
    'Loop through all documented timing backends
        For MethodID = 1 To 6
            'Advance method-level progress and update the Excel status bar
                CurrentStep = CurrentStep + 1
                Demo_SB_SetProgress CurrentStep, _
                                    TotalSteps, _
                                    "Running method " & CStr(MethodID) & " - " & cPM_Demo_MethodName(MethodID)

            'Create and configure a fresh timer instance for this method
                Set cPM = cPM_Demo_PrepareInstance()
            'Apply the selected TW mode
                cPM_Demo_ApplyTWMode cPM, TWMode
            'Clear the target range before the timed block
                Target.ClearContents
            'Write an initial in-loop status-bar message when enabled
                If ProgressEnabled Then
                    Demo_SB_SetProgress 0, Iterations, _
                                        "Method " & CStr(MethodID) & " - " & cPM_Demo_MethodName(MethodID)
                End If
            'Start the timing session
                cPM.StartTimer MethodID, AlignFlag
            'Repeat the value-fill workload for the configured number of iterations
                For i = 1 To Iterations
                    'Run the value-fill workload
                        Target.Value = FillValue
                    'Update the status bar periodically during the timed loop when enabled
                        If ProgressEnabled Then
                            If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                                Demo_SB_SetProgress i, Iterations, _
                                                    "Method " & CStr(MethodID) & " - " & cPM_Demo_MethodName(MethodID)
                            End If
                        End If
                Next i
            'Read numeric elapsed time
                ElapsedS = cPM.ElapsedSeconds()
            'Read formatted elapsed time without taking a second timing sample
                ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
            'Build the scenario note text
                Notes = "Rows=" & CStr(Target.Rows.Count) & _
                        " | Iterations=" & CStr(Iterations) & _
                        " | Align=" & CStr(AlignFlag) & _
                        " | FillValue=" & CStr(FillValue) & _
                        " | TW=" & TWMode
            'Append one result row for the current method
                cPM_Demo_AppendLog "All Methods", _
                                   MethodID, _
                                   cPM.MethodName(MethodID), _
                                   ElapsedS, _
                                   ElapsedTxt, _
                                   cPM.T1, _
                                   cPM.T2, _
                                   cPM.ET, _
                                   Notes
            'Release the current method instance cleanly before the next loop
                cPM.ResetEnvironment
                Set cPM = Nothing
        Next MethodID

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub
Public Sub cPM_RunAlignedDemo()
'
'==============================================================================
'                             RUN ALIGNED DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Compares non-aligned and aligned start timing for the currently selected
'   timing backend
'
' WHY THIS EXISTS
'   Alignment effects are easier to understand when the same workload is run
'   twice under the same method:
'     - once without alignment
'     - once with alignment
'
'   This version also supports optional in-loop status-bar progress for both
'   passes, enabled only when the iteration count is high enough to justify it
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected timing method
'   - Resolves the configured value-fill target range
'   - Reads the configured fill value
'   - Reads the configured iteration count
'   - Reads the selected TW mode once
'   - Runs one non-aligned pass and one aligned pass
'   - Updates the Excel status bar periodically inside each iterations loop
'     only when Iterations >= 100
'   - Appends one log row for each pass
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising errors
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetMethodID
'   - cPM_Demo_GetValueTargetRange
'   - cPM_Demo_GetValueFillValue
'   - cPM_Demo_GetIterations
'   - cPM_Demo_GetTWMode
'   - cPM_Demo_ApplyTWMode
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'   - cPM_Demo_MethodName
'
' NOTES
'   Showing progress inside the timed loops adds UI overhead and therefore
'   slightly contaminates the measurement. To reduce that effect:
'     - progress is shown only when Iterations >= 100
'     - the status bar is updated periodically rather than on every iteration
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description
    
    Dim MethodID            As Integer               'Selected timing method
    Dim Target              As Range                 'Worksheet target range
    Dim FillValue           As Variant               'Configured fill value
    Dim Iterations          As Long                  'Configured iteration count
    Dim TWMode              As String                'Selected TW mode
    Dim i                   As Long                  'Loop counter
    Dim ElapsedS            As Double                'Measured elapsed seconds
    Dim ElapsedTxt          As String                'Formatted elapsed-time text
    Dim ProgressEvery       As Long                  'Status-bar update frequency
    Dim ProgressEnabled     As Boolean               'TRUE when in-loop progress should be shown

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Read the selected timing method
        MethodID = cPM_Demo_GetMethodID()
    'Resolve the target worksheet range
        Set Target = cPM_Demo_GetValueTargetRange()
    'Read the configured fill value
        FillValue = cPM_Demo_GetValueFillValue()
    'Read the configured iteration count
        Iterations = cPM_Demo_GetIterations()
    'Read the selected TW mode once
        TWMode = cPM_Demo_GetTWMode()

'------------------------------------------------------------------------------
' RESOLVE PROGRESS POLICY
'------------------------------------------------------------------------------
    'Enable in-loop progress only for sufficiently large iteration counts
        ProgressEnabled = (Iterations >= 100)
    'Resolve the status-bar update interval only when progress is enabled
        If ProgressEnabled Then
            ProgressEvery = Iterations \ 100
            If ProgressEvery < 1 Then
                ProgressEvery = 1
            End If
        End If

'------------------------------------------------------------------------------
' RUN NON-ALIGNED PASS
'------------------------------------------------------------------------------
    'Write an initial status-bar message for the non-aligned pass when enabled
        If ProgressEnabled Then
            Demo_SB_SetProgress 0, Iterations, _
                                "Aligned Demo - Non-Aligned Pass - " & cPM_Demo_MethodName(MethodID)
        End If

    'Create and configure a timer instance for the non-aligned pass
        Set cPM = cPM_Demo_PrepareInstance()
    'Apply the selected TW mode
        cPM_Demo_ApplyTWMode cPM, TWMode
    'Clear the target range before the timed block
        Target.ClearContents
    'Start the non-aligned timing session
        cPM.StartTimer MethodID, False
    'Repeat the value-fill workload for the configured number of iterations
        For i = 1 To Iterations
            'Run the value-fill workload
                Target.Value = FillValue

            'Update the status bar periodically during the timed loop when enabled
                If ProgressEnabled Then
                    If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                        Demo_SB_SetProgress i, Iterations, _
                                            "Aligned Demo - Non-Aligned Pass - " & cPM_Demo_MethodName(MethodID)
                    End If
                End If
        Next i
    'Read the non-aligned results
        ElapsedS = cPM.ElapsedSeconds()
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
    'Log the non-aligned measurement
        cPM_Demo_AppendLog "Aligned Demo", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           "Iterations=" & CStr(Iterations) & _
                           " | Align=False | FillValue=" & CStr(FillValue)
    'Release the first pass cleanly
        cPM.ResetEnvironment
        Set cPM = Nothing

'------------------------------------------------------------------------------
' RUN ALIGNED PASS
'------------------------------------------------------------------------------
    'Write an initial status-bar message for the aligned pass when enabled
        If ProgressEnabled Then
            Demo_SB_SetProgress 0, Iterations, _
                                "Aligned Demo - Aligned Pass - " & cPM_Demo_MethodName(MethodID)
        End If

    'Create and configure a timer instance for the aligned pass
        Set cPM = cPM_Demo_PrepareInstance()
    'Apply the selected TW mode
        cPM_Demo_ApplyTWMode cPM, TWMode
    'Clear the target range before the timed block
        Target.ClearContents
    'Start the aligned timing session
        cPM.StartTimer MethodID, True
    'Repeat the value-fill workload for the configured number of iterations
        For i = 1 To Iterations
            'Run the value-fill workload
                Target.Value = FillValue

            'Update the status bar periodically during the timed loop when enabled
                If ProgressEnabled Then
                    If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                        Demo_SB_SetProgress i, Iterations, _
                                            "Aligned Demo - Aligned Pass - " & cPM_Demo_MethodName(MethodID)
                    End If
                End If
        Next i
    'Read the aligned results
        ElapsedS = cPM.ElapsedSeconds()
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
    'Log the aligned measurement
        cPM_Demo_AppendLog "Aligned Demo", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           "Iterations=" & CStr(Iterations) & _
                           " | Align=True | FillValue=" & CStr(FillValue)

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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
    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub
Public Sub cPM_RunDiagnostics()
'
'==============================================================================
'                             RUN DIAGNOSTICS
'------------------------------------------------------------------------------
' PURPOSE
'   Logs the main diagnostic and informational properties of the current
'   environment
'
' WHY THIS EXISTS
'   Demo users benefit from a quick diagnostic view of the timer environment
'   and QPC-related characteristics without needing to inspect the class code
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates one cPerformanceManager instance
'   - Logs system tick information
'   - Logs QPC tick / frequency information
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising errors
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_AppendLog
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Prepare a timer instance using the current strict-mode setting
        Set cPM = cPM_Demo_PrepareInstance()

'------------------------------------------------------------------------------
' LOG DIAGNOSTICS
'------------------------------------------------------------------------------
    'Append the nominal system tick interval
        cPM_Demo_AppendLog "Diagnostics", 0, "", 0#, "", 0#, 0#, 0#, _
                           cPM.Get_SystemTickInterval
    'Append the QPC tick interval
        cPM_Demo_AppendLog "Diagnostics", 0, "", 0#, "", 0#, 0#, 0#, _
                           cPM.QPC_Get_SystemTickInterval
    'Append the QPC frequency text
        cPM_Demo_AppendLog "Diagnostics", 0, "", 0#, "", 0#, 0#, 0#, _
                           cPM.QPC_FrequencyPerSecond
    'Append the QPC frequency numeric value
        cPM_Demo_AppendLog "Diagnostics", 0, "", 0#, "", 0#, 0#, 0#, _
                           "QPC Frequency Value = " & CStr(cPM.QPC_FrequencyPerSecond_Value)

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
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

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub


Public Sub cPM_RunOverheadDemo()
'
'==============================================================================
'                             RUN OVERHEAD DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Measures and logs benchmark-overhead information for the currently selected
'   timing method
'
' WHY THIS EXISTS
'   A benchmark result is easier to interpret when users can also inspect the
'   approximate measurement overhead of the selected backend
'
'   This version assumes the corrected class design where:
'     - OverheadMeasurement_Seconds performs the numeric measurement
'     - OverheadMeasurement_Text can format an already measured overhead value
'       without taking a second measurement
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected timing method
'   - Reads the selected iteration count
'   - Shows coarse stage progress in the status bar
'   - Measures the numeric timing overhead once
'   - Formats the text output from the same measured value
'   - Appends one log row
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising errors
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetMethodID
'   - cPM_Demo_GetIterations
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'
' NOTES
'   No progress is shown inside the measured loop because that would distort the
'   overhead measurement itself
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description
    
    Dim MethodID            As Integer               'Selected timing method
    Dim Iterations          As Long                  'Selected iteration count
    Dim OverheadS           As Double                'Numeric overhead measurement
    Dim OverheadTxt         As String                'Formatted overhead text

    Dim CurrentStep         As Long                  'Current stage progress step
    Const TotalSteps As Long = 3                     'Total number of visible execution stages

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Initialize stage progress
        CurrentStep = 0
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Preparing overhead demo"
    'Read the selected timing method
        MethodID = cPM_Demo_GetMethodID()
    'Read the selected iteration count
        Iterations = cPM_Demo_GetIterations()
    'Prepare a timer instance using the current strict-mode setting
        Set cPM = cPM_Demo_PrepareInstance()

'------------------------------------------------------------------------------
' RUN NUMERIC MEASUREMENT
'------------------------------------------------------------------------------
    'Advance progress before numeric overhead measurement
        CurrentStep = 1
        Demo_SB_SetProgress CurrentStep, _
                            TotalSteps, _
                            "Measuring numeric overhead - " & cPM.MethodName(MethodID)
    'Measure numeric timing overhead once
        OverheadS = cPM.OverheadMeasurement_Seconds(MethodID, Iterations)

'------------------------------------------------------------------------------
' FORMAT SAME MEASURED VALUE
'------------------------------------------------------------------------------
    'Advance progress before text formatting
        CurrentStep = 2
        Demo_SB_SetProgress CurrentStep, _
                            TotalSteps, _
                            "Formatting overhead result - " & cPM.MethodName(MethodID)
    'Format the same measured numeric value without taking a second measurement
        OverheadTxt = cPM.OverheadMeasurement_Text(MethodID, Iterations, OverheadS)

'------------------------------------------------------------------------------
' LOG RESULT
'------------------------------------------------------------------------------
    'Append one results row for the overhead demo
        cPM_Demo_AppendLog "Overhead Demo", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           OverheadS, _
                           OverheadTxt, _
                           0#, _
                           0#, _
                           0#, _
                           "Iterations=" & CStr(Iterations)

    'Mark completion in the status bar
        CurrentStep = 3
        Demo_SB_SetProgress CurrentStep, _
                            TotalSteps, _
                            "Overhead demo complete - " & cPM.MethodName(MethodID)

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub
Public Sub cPM_RunPauseDemo()
'
'==============================================================================
'                               RUN PAUSE DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Measures the requested Pause() scenario using QPC
'
' WHY THIS EXISTS
'   The demo workbook exposes pause controls separately from write / fill
'   workloads so users can inspect pause behavior and overshoot more directly
'
'   This version also shows coarse progress in the Excel status bar across the
'   main execution stages of the routine
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected pause seconds and pause method
'   - Uses QPC as the measurement backend
'   - Shows stage-level progress in the status bar
'   - Runs the selected pause scenario
'   - Appends one log row
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising errors
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetPauseSeconds
'   - cPM_Demo_GetPauseMethod
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'
' NOTES
'   This routine cannot show true live pause progress because the waiting logic
'   runs inside cPM.Pause
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description
    
    Dim PauseS              As Double                'Requested pause seconds
    Dim PauseMethod         As Integer               'Requested pause method
    Dim ElapsedS            As Double                'Measured elapsed seconds
    Dim ElapsedTxt          As String                'Formatted elapsed text

    Dim CurrentStep         As Long                  'Current stage progress step
    Const TotalSteps As Long = 4                     'Total number of visible execution stages

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Initialize stage progress
        CurrentStep = 0
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Preparing pause demo"
    'Read the selected pause seconds
        PauseS = cPM_Demo_GetPauseSeconds()
    'Read the selected pause method
        PauseMethod = cPM_Demo_GetPauseMethod()
    'Prepare a timer instance using QPC and current strict-mode setting
        Set cPM = cPM_Demo_PrepareInstance()

'------------------------------------------------------------------------------
' RUN PAUSE
'------------------------------------------------------------------------------
    'Advance progress before starting the timing session
        CurrentStep = 1
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Starting QPC timing session"
    'Start a QPC timing session
        cPM.StartTimer 5, False
    'Advance progress before running the pause scenario
        CurrentStep = 2
        Demo_SB_SetProgress CurrentStep, _
                            TotalSteps, _
                            "Running pause scenario - Method " & CStr(PauseMethod) & _
                            " | PauseSeconds=" & CStr(PauseS)
    'Run the selected pause scenario
        cPM.Pause PauseS, PauseMethod

'------------------------------------------------------------------------------
' READ RESULTS
'------------------------------------------------------------------------------
    'Advance progress before reading the results
        CurrentStep = 3
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Reading pause results"
    'Read numeric elapsed time
        ElapsedS = cPM.ElapsedSeconds()
    'Read formatted elapsed time without taking a second timing sample
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)

'------------------------------------------------------------------------------
' LOG RESULT
'------------------------------------------------------------------------------
    'Append one results row for the pause demo
        cPM_Demo_AppendLog "Pause Demo", _
                           5, _
                           cPM.MethodName(5), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           "PauseSeconds=" & CStr(PauseS) & " | PauseMethod=" & CStr(PauseMethod)
    'Mark completion in the status bar
        CurrentStep = 4
        Demo_SB_SetProgress CurrentStep, TotalSteps, "Pause demo complete"

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release environment changes deterministically
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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

    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub

Public Sub cPM_RunTWComparison()
'
'==============================================================================
'                            RUN TW COMPARISON
'------------------------------------------------------------------------------
' PURPOSE
'   Compares the same worksheet workload with TW inactive and with the selected
'   TW mode active
'
' WHY THIS EXISTS
'   This routine lets the user compare the same formula-fill workload under two
'   conditions:
'     - normal Excel environment
'     - selected Time-Wasters suppression mode
'
'   This version also supports optional in-loop progress in the Excel status
'   bar for both passes, enabled only when the iteration count is high enough
'   to justify it
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the selected timing method, alignment flag, and TW mode
'   - Resolves the configured formula-fill target range
'   - Reads the configured formula text
'   - Reads the configured iteration count
'   - Runs one normal pass and one TW-managed pass
'   - Updates the status bar periodically inside each iterations loop only when
'     Iterations >= 100
'   - Appends one structured log row for each pass
'
' ERROR POLICY
'   Restores the cPerformanceManager environment before re-raising the
'   original error
'
' DEPENDENCIES
'   - Btn_Click
'   - cPM_Demo_PrepareInstance
'   - cPM_Demo_GetMethodID
'   - cPM_Demo_GetAlignFlag
'   - cPM_Demo_GetTWMode
'   - cPM_Demo_GetFormulaTargetRange
'   - cPM_Demo_GetFormulaFillExpr
'   - cPM_Demo_GetIterations
'   - cPM_Demo_ApplyTWMode
'   - cPM_Demo_AppendLog
'   - Demo_SB_SetProgress
'   - cPM_Demo_MethodName
'
' NOTES
'   The configured formula text is expected to be stored in invariant Excel /
'   VBA formula syntax suitable for Range.Formula assignment, for example:
'
'     =RAND()
'     =ROW()*2
'     =SUM(A1:A10)
'
'   This routine intentionally uses Range.Formula rather than FormulaLocal
'   and therefore does not expect localized function names or locale-specific
'   list separators in the configured formula text
'
'   Showing progress inside the timed loops adds UI overhead and therefore
'   slightly contaminates the measurement. To reduce that effect:
'     - in-loop progress is shown only when Iterations >= 100
'     - the status bar is updated periodically rather than on every iteration
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Timer instance
    Dim MethodID            As Integer               'Selected timing method
    Dim AlignFlag           As Boolean               'Selected alignment flag
    Dim TWMode              As String                'Selected TW mode
    Dim Target              As Range                 'Worksheet target range
    Dim FormulaText         As String                'Configured formula text
    Dim Iterations          As Long                  'Configured iteration count
    Dim i                   As Long                  'Loop counter
    Dim ElapsedS            As Double                'Measured elapsed seconds
    Dim ElapsedTxt          As String                'Formatted elapsed text
    Dim ProgressEnabled     As Boolean               'TRUE when in-loop progress should be shown
    Dim ProgressEvery       As Long                  'Status-bar update frequency

    Dim SavedErrNumber      As Long                  'Captured error number
    Dim SavedErrSource      As String                'Captured error source
    Dim SavedErrDescription As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail
    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click
    'Read the selected timing method
        MethodID = cPM_Demo_GetMethodID()
    'Read the selected alignment flag
        AlignFlag = cPM_Demo_GetAlignFlag()
    'Read the selected TW mode
        TWMode = cPM_Demo_GetTWMode()
    'Resolve the target worksheet range
        Set Target = cPM_Demo_GetFormulaTargetRange()
    'Read the configured formula text
        FormulaText = cPM_Demo_GetFormulaFillExpr()
    'Read the configured iteration count
        Iterations = cPM_Demo_GetIterations()

'------------------------------------------------------------------------------
' RESOLVE PROGRESS POLICY
'------------------------------------------------------------------------------
    'Enable in-loop progress only for sufficiently large iteration counts
        ProgressEnabled = (Iterations >= 100)
    'Resolve the in-loop update interval only when progress is enabled
        If ProgressEnabled Then
            ProgressEvery = Iterations \ 100
            If ProgressEvery < 1 Then
                ProgressEvery = 1
            End If
        End If

'------------------------------------------------------------------------------
' RUN NORMAL PASS
'------------------------------------------------------------------------------
    'Write an initial status-bar message for the normal pass when enabled
        If ProgressEnabled Then
            Demo_SB_SetProgress 0, Iterations, _
                                "TW Comparison - Normal Pass - " & cPM_Demo_MethodName(MethodID)
        End If

    'Create and configure a timer instance for the normal pass
        Set cPM = cPM_Demo_PrepareInstance()
    'Clear the target range before the timed block
        Target.ClearContents
    'Start the normal timing session
        cPM.StartTimer MethodID, AlignFlag
    'Repeat the formula-fill workload for the configured number of iterations
        For i = 1 To Iterations
            'Apply the configured invariant formula text to the target range
                Target.Formula = FormulaText
            'Update the status bar periodically during the timed loop when enabled
                If ProgressEnabled Then
                    If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                        Demo_SB_SetProgress i, Iterations, _
                                            "TW Comparison - Normal Pass - " & cPM_Demo_MethodName(MethodID)
                    End If
                End If
        Next i
    'Read the normal-pass results
        ElapsedS = cPM.ElapsedSeconds()
    'Read the formatted elapsed time without taking a second timing sample
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
    'Append the normal-pass result row
        cPM_Demo_AppendLog "TW Comparison", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           "Rows=" & CStr(Target.Rows.Count) & _
                           " | Iterations=" & CStr(Iterations) & _
                           " | Align=" & CStr(AlignFlag) & _
                           " | Formula=" & FormulaText & _
                           " | TW=Normal"
    'Release the normal pass cleanly before starting the TW-managed pass
        cPM.ResetEnvironment
        Set cPM = Nothing

'------------------------------------------------------------------------------
' RUN TW PASS
'------------------------------------------------------------------------------
    'Write an initial status-bar message for the TW pass when enabled
        If ProgressEnabled Then
            Demo_SB_SetProgress 0, Iterations, _
                                "TW Comparison - TW Pass - " & cPM_Demo_MethodName(MethodID)
        End If

    'Create and configure a timer instance for the TW pass
        Set cPM = cPM_Demo_PrepareInstance()
    'Apply the selected TW mode
        cPM_Demo_ApplyTWMode cPM, TWMode
    'Clear the target range before the timed block
        Target.ClearContents
    'Start the TW timing session
        cPM.StartTimer MethodID, AlignFlag
    'Repeat the formula-fill workload for the configured number of iterations
        For i = 1 To Iterations
            'Apply the configured invariant formula text to the target range
                Target.Formula = FormulaText
            'Update the status bar periodically during the timed loop when enabled
                If ProgressEnabled Then
                    If (i = 1) Or ((i Mod ProgressEvery) = 0) Or (i = Iterations) Then
                        Demo_SB_SetProgress i, Iterations, _
                                            "TW Comparison - TW Pass - " & cPM_Demo_MethodName(MethodID)
                    End If
                End If
        Next i
    'Read the TW-pass results
        ElapsedS = cPM.ElapsedSeconds()
    'Read the formatted elapsed time without taking a second timing sample
        ElapsedTxt = cPM.ElapsedTime(, ElapsedS)
    'Append the TW-pass result row
        cPM_Demo_AppendLog "TW Comparison", _
                           MethodID, _
                           cPM.MethodName(MethodID), _
                           ElapsedS, _
                           ElapsedTxt, _
                           cPM.T1, _
                           cPM.T2, _
                           cPM.ET, _
                           "Rows=" & CStr(Target.Rows.Count) & _
                           " | Iterations=" & CStr(Iterations) & _
                           " | Align=" & CStr(AlignFlag) & _
                           " | Formula=" & FormulaText & _
                           " | TW=" & TWMode

CleanExit:
'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release any remaining environment changes
        If Not cPM Is Nothing Then
            On Error Resume Next
            cPM.ResetEnvironment
            Set cPM = Nothing
            On Error GoTo 0
        End If
    'Return control of the status bar to Excel
        Application.StatusBar = False
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
    'Continue through the centralized cleanup path
        Resume CleanExit

End Sub

Public Sub cPM_ClearResultsLog()
'
'==============================================================================
'                            CLEAR RESULTS LOG
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the data rows from tbl_cPM_Log while preserving the table structure
'
' WHY THIS EXISTS
'   Demo users often need to reset the visible results area without rebuilding
'   the whole workbook layout
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   Deletes the current DataBodyRange when rows are present
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lo                  As ListObject       'Results-log table

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail

    'Simulate a pressed button when the routine was launched by a shape
        Btn_Click

    'Resolve the results-log table on the demo sheet
        Set Lo = ThisWorkbook.Worksheets(cPM_SHEET_DEMO).ListObjects(cPM_TABLE_LOG)

'------------------------------------------------------------------------------
' CLEAR BODY
'------------------------------------------------------------------------------
    'Delete the table body when rows are present
        If Not Lo.DataBodyRange Is Nothing Then
            Lo.DataBodyRange.Delete
        End If

CleanExit:
'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    Exit Sub

CleanFail:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'
'==============================================================================
'                  PRIVATE: DEMO INSTANCE / CONTROL READERS
'==============================================================================

Private Function cPM_Demo_MethodName( _
    ByVal MethodID As Integer) _
    As String
'
'==============================================================================
'                           DEMO METHOD NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display name of one timing method for demo/status-bar use
'
' WHY THIS EXISTS
'   Some demo/status-bar messages need a method label before a timer instance
'   has been created for the current loop iteration
'
' INPUTS
'   MethodID
'     Timing backend identifier
'
' RETURNS
'   String
'     Human-readable method label, or vbNullString when the ID is invalid
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the human-readable method label
        Select Case MethodID
            Case 1
                cPM_Demo_MethodName = "Timer"
            Case 2
                cPM_Demo_MethodName = "GetTickCount / GetTickCount64"
            Case 3
                cPM_Demo_MethodName = "timeGetTime"
            Case 4
                cPM_Demo_MethodName = "timeGetSystemTime"
            Case 5
                cPM_Demo_MethodName = "QPC"
            Case 6
                cPM_Demo_MethodName = "Now()"
            Case Else
                cPM_Demo_MethodName = vbNullString
        End Select

End Function

Private Function cPM_Demo_PrepareInstance() As cPerformanceManager
'
'==============================================================================
'                           PREPARE DEMO INSTANCE
'------------------------------------------------------------------------------
' PURPOSE
'   Creates and configures one cPerformanceManager instance from the current
'   control-panel settings
'
' WHY THIS EXISTS
'   Every public demo action needs the same base class setup:
'     - create a fresh instance
'     - apply the current strict-mode setting
'
'   Centralizing that setup avoids duplication and prevents inconsistent
'   configuration across demo routines
'
' INPUTS
'   None
'
' RETURNS
'   cPerformanceManager
'     Newly created and configured instance
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim cPM                 As cPerformanceManager   'Configured class instance

'------------------------------------------------------------------------------
' CREATE / CONFIGURE INSTANCE
'------------------------------------------------------------------------------
    'Create a fresh timer instance
        Set cPM = New cPerformanceManager
    'Apply the selected strict-mode flag
        cPM.StrictMode = cPM_Demo_GetStrictFlag()

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the configured class instance
        Set cPM_Demo_PrepareInstance = cPM

End Function


Private Function cPM_Demo_GetControlValue( _
    ByVal NameText As String) _
    As Variant
'
'==============================================================================
'                           GET CONTROL VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current value of one workbook-level named control cell
'
' WHY THIS EXISTS
'   The demo workbook exposes its controls through workbook-level names so demo
'   logic can remain independent from hard-coded worksheet addresses
'
' INPUTS
'   NameText
'     Workbook-level defined name to read
'
' RETURNS
'   Variant
'     The value stored in the named cell
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the value from the workbook-level named control cell
        cPM_Demo_GetControlValue = ThisWorkbook.Names(NameText).RefersToRange.Value

End Function


Private Function cPM_Demo_GetMethodID() As Integer
'
'==============================================================================
'                             GET METHOD ID
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and normalizes the selected timing method from the demo control panel
'
' WHY THIS EXISTS
'   The timing-method control should always resolve into the documented method
'   range even when users type unexpected values manually
'
' RETURNS
'   Integer
'     Normalized method identifier between 1 and 6 inclusive
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant          'Raw control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the raw control value
        V = cPM_Demo_GetControlValue(cPM_NAME_METHOD)

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Coerce the control value into the documented method range
        If Val(V) < 1 Then
            cPM_Demo_GetMethodID = 1
        ElseIf Val(V) > 6 Then
            cPM_Demo_GetMethodID = 6
        Else
            cPM_Demo_GetMethodID = CInt(Val(V))
        End If

End Function


Private Function cPM_Demo_GetAlignFlag() As Boolean
'
'==============================================================================
'                            GET ALIGN FLAG
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the selected AlignToNextTick flag from the demo control panel
'
' RETURNS
'   Boolean
'     TRUE when the control contains TRUE, otherwise FALSE
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Convert the control value into a Boolean flag
        cPM_Demo_GetAlignFlag = _
            (UCase$(Trim$(CStr(cPM_Demo_GetControlValue(cPM_NAME_ALIGN)))) = "TRUE")

End Function


Private Function cPM_Demo_GetStrictFlag() As Boolean
'
'==============================================================================
'                           GET STRICT FLAG
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the selected StrictMode flag from the demo control panel
'
' RETURNS
'   Boolean
'     TRUE when the control contains TRUE, otherwise FALSE
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Convert the control value into a Boolean flag
        cPM_Demo_GetStrictFlag = _
            (UCase$(Trim$(CStr(cPM_Demo_GetControlValue(cPM_NAME_STRICT)))) = "TRUE")

End Function


Private Function cPM_Demo_GetPauseSeconds() As Double
'
'==============================================================================
'                          GET PAUSE SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and normalizes the Pause Seconds control value
'
' RETURNS
'   Double
'     Pause seconds clamped to zero or above
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant          'Raw control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the raw control value
        V = cPM_Demo_GetControlValue(cPM_NAME_PAUSESEC)

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Clamp negative values to zero
        If CDbl(V) < 0# Then
            cPM_Demo_GetPauseSeconds = 0#
        Else
            cPM_Demo_GetPauseSeconds = CDbl(V)
        End If

End Function


Private Function cPM_Demo_GetPauseMethod() As Integer
'
'==============================================================================
'                           GET PAUSE METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and normalizes the Pause Method control value
'
' RETURNS
'   Integer
'     Pause method clamped to the documented range 1 to 4 inclusive
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant          'Raw control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the raw control value
        V = cPM_Demo_GetControlValue(cPM_NAME_PAUSEMETH)

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Coerce the pause method into the documented range
        If Val(V) < 1 Then
            cPM_Demo_GetPauseMethod = 1
        ElseIf Val(V) > 4 Then
            cPM_Demo_GetPauseMethod = 4
        Else
            cPM_Demo_GetPauseMethod = CInt(Val(V))
        End If

End Function


Private Function cPM_Demo_GetIterations() As Long
'
'==============================================================================
'                            GET ITERATIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and normalizes the Iterations control value
'
' RETURNS
'   Long
'     Iteration count clamped to a minimum of 1
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant          'Raw control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the raw control value
        V = cPM_Demo_GetControlValue(cPM_NAME_ITER)

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Clamp the iteration count to a minimum of 1
        If CLng(V) < 1 Then
            cPM_Demo_GetIterations = 1
        Else
            cPM_Demo_GetIterations = CLng(V)
        End If

End Function


Private Function cPM_Demo_GetRangeRows() As Long
'
'==============================================================================
'                             GET RANGE ROWS
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and normalizes the Range Rows control value
'
' RETURNS
'   Long
'     Row count clamped to a minimum of 1
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V                   As Variant          'Raw control value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Read the raw control value
        V = cPM_Demo_GetControlValue(cPM_NAME_ROWS)

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Clamp the row count to a minimum of 1
        If CLng(V) < 1 Then
            cPM_Demo_GetRangeRows = 1
        Else
            cPM_Demo_GetRangeRows = CLng(V)
        End If

End Function


Private Function cPM_Demo_GetTWMode() As String
'
'==============================================================================
'                              GET TW MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Reads the selected TW mode text from the demo control panel
'
' RETURNS
'   String
'     Trimmed TW mode text
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the selected TW mode text
        cPM_Demo_GetTWMode = Trim$(CStr(cPM_Demo_GetControlValue(cPM_NAME_TWMODE)))

End Function

'
'==============================================================================
'                 PRIVATE: TW DISPATCH / RANGE / VALUE HELPERS
'==============================================================================

Private Sub cPM_Demo_ApplyTWMode( _
    ByVal cPM As cPerformanceManager, _
    ByVal TWMode As String)
'
'==============================================================================
'                             APPLY TW MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the selected demo TW mode to one cPerformanceManager instance
'
' WHY THIS EXISTS
'   The demo workbook exposes user-friendly TW mode text, while the class
'   expects the corresponding TW enumeration flags
'
' INPUTS
'   cPM
'     Target cPerformanceManager instance
'
'   TWMode
'     Selected TW mode text from the control panel
'
' RETURNS
'   None
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case UCase$(TWMode)
        Case "", "NONE"
            'Do nothing when TW is not requested
        Case "ALL OFF"
            'Disable all supported TW settings
                cPM.TW_Turn_OFF TW_Enum.None
        Case "KEEP SCREENUPDATING"
            'Keep ScreenUpdating ON
                cPM.TW_Turn_OFF TW_Enum.ScreenUpdating
        Case "KEEP ENABLEEVENTS"
            'Keep EnableEvents ON
                cPM.TW_Turn_OFF TW_Enum.EnableEvents
        Case "KEEP SCREENUPDATING+ENABLEEVENTS"
            'Keep both ScreenUpdating and EnableEvents ON
                cPM.TW_Turn_OFF (TW_Enum.ScreenUpdating Or TW_Enum.EnableEvents)
        Case Else
            'Fallback: treat unknown text as no TW action
    End Select

End Sub


Private Function cPM_Demo_GetValueTargetRange() As Range
'
'==============================================================================
'                       GET VALUE TARGET RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the value-fill target range configured on DATA_cPM
'
' WHY THIS EXISTS
'   The demo control stores the value target anchor as text in a workbook-level
'   named input cell. This helper reads that configured text address, resolves
'   it to a real worksheet range, and resizes it to the requested row count
'
' RETURNS
'   Range
'     Resolved one-column value-fill target range
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - cPM_Demo_GetConfiguredTargetRange
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Resolve and return the configured value-fill target range
        Set cPM_Demo_GetValueTargetRange = cPM_Demo_GetConfiguredTargetRange(cPM_NAME_VALUEFILL)

End Function


Private Function cPM_Demo_GetFormulaTargetRange() As Range
'
'==============================================================================
'                        GET FORMULA TARGET RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the formula-fill target range configured on DATA_cPM
'
' WHY THIS EXISTS
'   The demo control stores the formula target anchor as text in a workbook-level
'   named input cell. This helper reads that configured text address, resolves
'   it to a real worksheet range, and resizes it to the requested row count
'
' RETURNS
'   Range
'     Resolved one-column formula-fill target range
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - cPM_Demo_GetConfiguredTargetRange
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Resolve and return the configured formula-fill target range
        Set cPM_Demo_GetFormulaTargetRange = cPM_Demo_GetConfiguredTargetRange(cPM_NAME_FORMULAFILL)

End Function


Private Function cPM_Demo_GetConfiguredTargetRange( _
    ByVal AddressNameText As String) _
    As Range
'
'==============================================================================
'                      GET CONFIGURED TARGET RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one configured DATA_cPM target range from a workbook-level named
'   cell containing the text address of the workload anchor cell
'
' WHY THIS EXISTS
'   Both value-fill and formula-fill workloads follow the same pattern:
'     - read a text address from a named cell
'     - resolve the anchor cell on DATA_cPM
'     - resize the anchor to the requested number of rows
'
'   Centralizing that logic avoids duplication and keeps both target-resolution
'   paths consistent
'
' INPUTS
'   AddressNameText
'     Workbook-level name whose cell contains the text address to resolve
'
' RETURNS
'   Range
'     Resolved one-column target range
'
' ERROR POLICY
'   Raises when:
'     - the configured address is blank
'     - the configured address cannot be resolved on DATA_cPM
'     - the configured address resolves to more than one cell
'     - the requested resized range would exceed worksheet row limits
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS_Data             As Worksheet         'DATA worksheet
    Dim TargetAddress       As String            'Configured anchor-cell address
    Dim AnchorCell          As Range             'Top-left anchor cell of the target range
    Dim RowCount            As Long              'Requested target row count
    Dim LastRequiredRow     As Long              'Last row required by the resized target

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the DATA worksheet
        Set WS_Data = ThisWorkbook.Worksheets(cPM_SHEET_DATA)

'------------------------------------------------------------------------------
' READ CONFIGURED TARGET ADDRESS
'------------------------------------------------------------------------------
    'Read the text address stored in the workbook-level named input cell
        TargetAddress = _
            Trim$(CStr(ThisWorkbook.Names(AddressNameText).RefersToRange.Value))
    'Reject blank configured addresses
        If Len(TargetAddress) = 0 Then
            Err.Raise vbObjectError + 2100, "cPM_Demo_GetConfiguredTargetRange", _
                "The configured target address is blank for name '" & _
                AddressNameText & "'"
        End If

'------------------------------------------------------------------------------
' RESOLVE TARGET ANCHOR
'------------------------------------------------------------------------------
    'Resolve the configured target anchor on DATA_cPM
        On Error GoTo InvalidAddress
        Set AnchorCell = WS_Data.Range(TargetAddress)
        On Error GoTo 0
    'Reject multi-cell configured addresses
        If AnchorCell.Cells.CountLarge <> 1 Then
            Err.Raise vbObjectError + 2101, "cPM_Demo_GetConfiguredTargetRange", _
                "The configured target address must resolve to exactly one cell on '" _
                & cPM_SHEET_DATA & "'"
        End If

'------------------------------------------------------------------------------
' VALIDATE REQUESTED SIZE
'------------------------------------------------------------------------------
    'Read the requested target row count
        RowCount = cPM_Demo_GetRangeRows()
    'Resolve the last worksheet row required by the resized target
        LastRequiredRow = AnchorCell.Row + RowCount - 1
    'Reject targets that would run past the worksheet row limit
        If LastRequiredRow > WS_Data.Rows.Count Then
            Err.Raise vbObjectError + 2102, _
                      "cPM_Demo_GetConfiguredTargetRange", _
                      "The configured target range exceeds the available rows on '" & cPM_SHEET_DATA & "'"
        End If

'------------------------------------------------------------------------------
' RETURN RESIZED RANGE
'------------------------------------------------------------------------------
    'Return the target range resized to the requested number of rows
        Set cPM_Demo_GetConfiguredTargetRange = AnchorCell.Resize(RowCount, 1)

    Exit Function

InvalidAddress:
'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
    'Raise a clearer error when the configured address cannot be resolved
        Err.Raise vbObjectError + 2103, _
                  "cPM_Demo_GetConfiguredTargetRange", _
                  "The configured target address '" & TargetAddress & "' is invalid on '" & cPM_SHEET_DATA & "'"

End Function


Private Function cPM_Demo_GetValueFillValue() As Variant
'
'==============================================================================
'                         GET VALUE FILL VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the configured value used by value-fill demo workloads
'
' RETURNS
'   Variant
'     Value stored in cPM_ValueFillValue
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the configured value-fill value
        cPM_Demo_GetValueFillValue = ThisWorkbook.Names(cPM_NAME_VALUEFILLVALUE).RefersToRange.Value

End Function


Private Function cPM_Demo_GetFormulaFillExpr() As String
'
'==============================================================================
'                       GET FORMULA FILL EXPRESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the configured formula text used by formula-fill demo workloads
'
' RETURNS
'   String
'     Formula text stored in cPM_FormulaFillExpr
'
' ERROR POLICY
'   Raises errors normally
'
' NOTES
'   The configured formula text is expected to be stored in invariant Excel /
'   VBA formula syntax suitable for Range.Formula assignment, for example:
'
'     =RAND()
'     =ROW()*2
'     =SUM(A1:A10)
'
'   This demo does not use FormulaLocal and therefore does not expect localized
'   function names or locale-specific list separators in the configured formula
'   text
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the configured formula-fill expression text
        cPM_Demo_GetFormulaFillExpr = CStr(ThisWorkbook.Names(cPM_NAME_FORMULAFILLEXPR).RefersToRange.Value)

End Function

'
'==============================================================================
'                         PRIVATE: LOG APPEND HELPERS
'==============================================================================

Private Sub cPM_Demo_AppendLog( _
    ByVal ScenarioText As String, _
    ByVal MethodID As Integer, _
    ByVal MethodNameText As String, _
    ByVal ElapsedSecondsValue As Double, _
    ByVal ElapsedTimeText As String, _
    ByVal T1Value As Double, _
    ByVal T2Value As Double, _
    ByVal ETValue As Double, _
    ByVal NotesText As String)
'
'==============================================================================
'                              APPEND LOG ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one result row to Tbl_cPM_Log on DEMO_cPM
'
' WHY THIS EXISTS
'   Every demo action should leave a persistent structured result so users can
'   compare runs directly in the workbook instead of relying only on the
'   Immediate Window
'
' INPUTS
'   ScenarioText
'     Scenario label
'
'   MethodID
'     Timing method identifier
'
'   MethodNameText
'     Timing method descriptive name
'
'   ElapsedSecondsValue
'     Numeric elapsed-seconds result
'
'   ElapsedTimeText
'     Human-readable elapsed-time result
'
'   T1Value
'     Logged T1 value
'
'   T2Value
'     Logged T2 value
'
'   ETValue
'     Logged ET value
'
'   NotesText
'     Free-text notes for the run
'
' RETURNS
'   None
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Lo                  As ListObject       'Results-log table
    Dim LR                  As ListRow          'Newly added log row
    Dim RunID               As Long             'Sequential run ID

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the results-log table on the demo sheet
        Set Lo = ThisWorkbook.Worksheets(cPM_SHEET_DEMO).ListObjects(cPM_TABLE_LOG)
    'Add one new row to the table
        Set LR = Lo.ListRows.Add
    'Compute the sequential run ID from the current body-row count
        RunID = Lo.ListRows.Count

'------------------------------------------------------------------------------
' WRITE LOG FIELDS
'------------------------------------------------------------------------------
    'Write the run ID
        LR.Range(1, 1).Value = RunID
        LR.Range(1, 1).HorizontalAlignment = xlCenter
    'Write the timestamp
        LR.Range(1, 2).Value = Now
        LR.Range(1, 2).HorizontalAlignment = xlCenter
    'Write the scenario label
        LR.Range(1, 3).Value = ScenarioText
        LR.Range(1, 3).HorizontalAlignment = xlCenter
    'Write the method ID
        LR.Range(1, 4).Value = MethodID
        LR.Range(1, 4).HorizontalAlignment = xlCenter
    'Write the method name
        LR.Range(1, 5).Value = MethodNameText
        LR.Range(1, 5).HorizontalAlignment = xlCenter
    'Write the numeric elapsed seconds
        LR.Range(1, 6).Value = ElapsedSecondsValue
        LR.Range(1, 6).HorizontalAlignment = xlCenter
    'Write the formatted elapsed time
        LR.Range(1, 7).Value = ElapsedTimeText
    'Write T1
        LR.Range(1, 8).Value = T1Value
        LR.Range(1, 8).HorizontalAlignment = xlCenter
    'Write T2
        LR.Range(1, 9).Value = T2Value
        LR.Range(1, 9).HorizontalAlignment = xlCenter
    'Write ET
        LR.Range(1, 10).Value = ETValue
        LR.Range(1, 10).HorizontalAlignment = xlCenter
    'Write notes
        LR.Range(1, 11).Value = NotesText

'------------------------------------------------------------------------------
' FORMAT FIELDS
'------------------------------------------------------------------------------
    'Format the timestamp
        LR.Range(1, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    'Format the numeric timing outputs
        LR.Range(1, 6).NumberFormat = "0.000000000"
        LR.Range(1, 8).NumberFormat = "0.000000000"
        LR.Range(1, 9).NumberFormat = "0.000000000"
        LR.Range(1, 10).NumberFormat = "0.000000000"

End Sub

