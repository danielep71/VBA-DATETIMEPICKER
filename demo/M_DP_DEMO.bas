Attribute VB_Name = "M_DP_DEMO"
Option Explicit

'
'==============================================================================
' M_DP_DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the DatePicker demo worksheet from code, so the demo surface can be
'   recreated on demand in any workbook rather than existing only as content
'   inside a tracked binary
'
' WHY THIS EXISTS
'   The demo sheet was previously hand-built and lived only inside
'   demo/DATEPICKER.xlsm. That had two consequences:
'
'     - Ribbon_Demo resolves the sheet from ThisWorkbook, so the Ribbon demo
'       button works when the code is embedded in the demo workbook and fails
'       in the .xlam, which has no worksheets at all
'
'     - the demo had no reviewable source. A change to it appeared in a pull
'       request as a binary diff and nothing else
'
'   This module removes both. The demo is now composed from the M_DEMO_BUILDER
'   primitives into a caller-supplied workbook, which means the add-in can build
'   it into whatever the user has open
'
' PUBLIC SURFACE
'   - DP_Demo_CreateDemoSheet
'   - DP_Demo_EnsureDemoSheet
'
' EXPECTED DEMO SHEET
'   - Worksheet name: DATE PICKER DEMO
'
' DEMO SEMANTICS
'   Green input cells are the surfaces the DatePicker is meant to write into.
'   Each section states what the picker should do when one of its cells is
'   selected, so the sheet doubles as a manual acceptance checklist
'
' TARGET WORKBOOK
'   Every routine takes an optional target workbook and defaults to
'   ActiveWorkbook, never ThisWorkbook. In the .xlam, ThisWorkbook is the
'   add-in, which cannot hold a demo sheet
'
' DEPENDENCIES
'   M_DEMO_BUILDER
'     DEMO_FastMode_Begin / DEMO_FastMode_End
'     DEMO_Sheet_BuildTemplate
'     DEMO_Write_BandHeader
'     DEMO_Format_Labels
'     DEMO_Format_InputCell
'     DEMO_Format_OutputCell
'     DEMO_SetRangeBorder
'     DEMO_Create_TableSection
'     DEMO_Table_AppendRow
'
' NOTES
'   Sample dates are literal rather than derived from the current date. The
'   Format Showcase section demonstrates how specific number formats render, and
'   a shifting sample would make the rendered values inconsistent with the format
'   strings beside them
'
'   The locale-dependent long-date sample renders in the user's Excel language.
'   That is intentional: it demonstrates the UseLocalNames setting
'
' UPDATED
'   2026-08-22
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------

    '-------------------------------SHEET--------------------------------------
    Private Const DEMO_SHEET_NAME               As String = "DATE PICKER DEMO"  'Demo worksheet name
    Private Const DEMO_TITLE                    As String = "DATETIME PICKER"   'Title band caption
    Private Const DEMO_SUBTITLE                 As String = "Demo"              'Subtitle band caption

    '-------------------------------LAYOUT-------------------------------------
    Private Const DEMO_CONTENT_COLUMNS          As String = "C:K"               'Content columns
    Private Const DEMO_CONTENT_WIDTH            As Double = 22                  'Content column width
    Private Const DEMO_SEPARATOR_COLUMNS        As String = "L:L"               'Right separator column
    Private Const DEMO_BODY_ROW_HEIGHT          As Double = 18                  'Standard body row height
    Private Const DEMO_ZOOM_PERCENT             As Long = 100                   'Demo sheet zoom

    '------------------------------SECTIONS------------------------------------
    Private Const DEMO_NOTES_RANGE              As String = "C4:K9"             'Instruction block
    Private Const DEMO_SINGLE_HEADER            As String = "C11:F11"           'Single-cell section band
    Private Const DEMO_SINGLE_COLUMNS           As String = "C12:F12"           'Single-cell column headers
    Private Const DEMO_SINGLE_BODY              As String = "C13:F19"           'Single-cell body
    Private Const DEMO_SINGLE_INPUTS            As String = "D13:D19"           'Single-cell input column

    Private Const DEMO_FORMAT_HEADER            As String = "H11:K11"           'Format showcase band
    Private Const DEMO_FORMAT_COLUMNS           As String = "H12:K12"           'Format showcase column headers
    Private Const DEMO_FORMAT_BODY              As String = "H13:K19"           'Format showcase body
    Private Const DEMO_FORMAT_SAMPLES           As String = "I13:I19"           'Format showcase sample column

    Private Const DEMO_MULTI_HEADER             As String = "C21:F21"           'Multi-cell section band
    Private Const DEMO_MULTI_COLUMNS            As String = "C22:F22"           'Multi-cell column headers
    Private Const DEMO_MULTI_BODY               As String = "C23:F27"           'Multi-cell body
    Private Const DEMO_MULTI_INPUTS             As String = "D23:D27"           'Multi-cell input column

    Private Const DEMO_TABLE_HEADER             As String = "H21:K21"           'Table section band
    Private Const DEMO_TABLE_TOPLEFT            As String = "H22"               'Table header top-left cell
    Private Const DEMO_TABLE_NAME               As String = "tblDatePickerDemo" 'ListObject name
    Private Const DEMO_TABLE_STYLE              As String = "TableStyleMedium2" 'ListObject style
    Private Const DEMO_TABLE_ROWS               As Long = 5                     'Sample rows appended

    '-------------------------------COLOURS------------------------------------
    'The builder has no green. COLOR_INPUT is pale yellow and is kept for cells
    'that must NOT qualify, so the two states are distinguishable at a glance
    Private Const DEMO_COLOR_QUALIFY            As Long = 14806246              'Pale green - cell should qualify
    Private Const DEMO_COLOR_NOTES              As Long = 15921906              'Light grey - instruction block

    '------------------------------COLUMN WIDTHS-------------------------------
    Private Const DEMO_WIDTH_SCENARIO           As Double = 26                  'Scenario caption columns
    Private Const DEMO_WIDTH_VALUE              As Double = 16                  'Input and sample columns
    Private Const DEMO_WIDTH_EXPECTED           As Double = 17                  'Expected verdict columns
    Private Const DEMO_WIDTH_NOTES              As Double = 40                  'Notes column
    Private Const DEMO_WIDTH_FORMAT             As Double = 22                  'Format string column

    '-------------------------------FORMATS------------------------------------
    Private Const DEMO_FORMAT_DATE              As String = "dd/mm/yyyy"        'Standard demo date format
    Private Const DEMO_FORMAT_DATETIME          As String = "dd/mm/yyyy hh:mm:ss" 'Standard demo datetime format

'
'------------------------------------------------------------------------------
'
'                               PUBLIC ENTRY POINTS
'
'------------------------------------------------------------------------------
'

Public Sub DP_Demo_CreateDemoSheet( _
    Optional ByVal TargetWorkbook As Workbook = Nothing)

'
'------------------------------------------------------------------------------
'                           CREATE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds or rebuilds the DatePicker demo worksheet in the target workbook
'
' WHY THIS EXISTS
'   The demo surface needs a single repeatable builder so it can be recreated
'   from the add-in, from an embedded copy, or from a fresh workbook, without
'   depending on a tracked binary that has no reviewable diff
'
' INPUTS
'   TargetWorkbook
'     Workbook to build into. Defaults to ActiveWorkbook when omitted
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Enters a fast-mode scope, rebuilds the demo template, writes the instruction
'   block and the four demo sections, activates the sheet, and restores
'   Application state through a single cleanup path
'
'   An existing demo sheet is rebuilt in place rather than duplicated
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
'   Cleanup is best-effort and must not overwrite the original error
'
'   Cursor positioning is cosmetic and is suppressed. A build that produced the
'   sheet correctly must not fail because the sheet could not be activated
'
' DEPENDENCIES
'   DP_Demo_ResolveTargetWorkbook
'   DEMO_FastMode_Begin / DEMO_FastMode_End
'   DEMO_Sheet_BuildTemplate
'   DP_Demo_WriteInstructions
'   DP_Demo_BuildSingleCellSection
'   DP_Demo_BuildFormatShowcaseSection
'   DP_Demo_BuildMultiCellSection
'   DP_Demo_BuildTableSection
'
' NOTES
'   The target workbook is resolved from ActiveWorkbook, never ThisWorkbook, so
'   the routine behaves identically whether the code is embedded or loaded as an
'   add-in
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_DP_DEMO.DP_Demo_CreateDemoSheet"

    Dim Wb                  As Workbook                 'Resolved target workbook
    Dim WS                  As Worksheet                'Demo worksheet
    Dim FastModeState       As tDEMOFastModeState       'Saved Application-state snapshot
    Dim FastModeOn          As Boolean                  'True once fast mode was entered
    Dim SavedErrNumber      As Long                     'Captured error number
    Dim SavedErrDescription As String                   'Captured error description
    Dim HandlerStep         As String                   'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo CleanFail
        HandlerStep = "Resolve target workbook"

    'Resolve the workbook that will receive the demo sheet
        Set Wb = DP_Demo_ResolveTargetWorkbook(TargetWorkbook)

    'Capture and apply fast-mode Application settings
        HandlerStep = "Enter fast mode"
        DEMO_FastMode_Begin FastModeState
        FastModeOn = True

    'Show the wait cursor while the sheet is rebuilt
        Application.Cursor = xlWait

'------------------------------------------------------------------------------
' BUILD TEMPLATE
'------------------------------------------------------------------------------
    'Build or rebuild the standard demo template
        HandlerStep = "Build sheet template"
    'Named arguments are used so only the values that differ from the builder
    'defaults appear here. Passing all twenty-five positionally exceeds the VBA
    'limit of twenty-five line continuations in a single statement
        DEMO_Sheet_BuildTemplate _
            WS_Name:=DEMO_SHEET_NAME, Title:=DEMO_TITLE, SubTitle:=DEMO_SUBTITLE, _
            TargetWorkbook:=Wb, _
            ContentColumns:=DEMO_CONTENT_COLUMNS, ContentColumnWidth:=DEMO_CONTENT_WIDTH, _
            SeparatorColumns:=DEMO_SEPARATOR_COLUMNS, _
            BodyRowHeight:=DEMO_BODY_ROW_HEIGHT, ZoomPercent:=DEMO_ZOOM_PERCENT

    'Resolve the demo sheet after template preparation
        HandlerStep = "Resolve demo sheet"
        Set WS = Wb.Worksheets(DEMO_SHEET_NAME)

    'Apply per-column widths. The template applies one width across the content
    'block, which leaves the Notes column too narrow to read
        HandlerStep = "Apply column widths"
        WS.Columns("C").ColumnWidth = DEMO_WIDTH_SCENARIO
        WS.Columns("D").ColumnWidth = DEMO_WIDTH_VALUE
        WS.Columns("E").ColumnWidth = DEMO_WIDTH_EXPECTED
        WS.Columns("F").ColumnWidth = DEMO_WIDTH_NOTES
        WS.Columns("H").ColumnWidth = DEMO_WIDTH_SCENARIO
        WS.Columns("I").ColumnWidth = DEMO_WIDTH_VALUE
        WS.Columns("J").ColumnWidth = DEMO_WIDTH_FORMAT
        WS.Columns("K").ColumnWidth = DEMO_WIDTH_EXPECTED

'------------------------------------------------------------------------------
' BUILD SECTIONS
'------------------------------------------------------------------------------
    'Write the instruction block
        HandlerStep = "Write instructions"
        DP_Demo_WriteInstructions WS

    'Build the basic single-cell scenario section
        HandlerStep = "Build single-cell section"
        DP_Demo_BuildSingleCellSection WS

    'Build the number-format showcase section
        HandlerStep = "Build format showcase section"
        DP_Demo_BuildFormatShowcaseSection WS

    'Build the multi-cell selection section
        HandlerStep = "Build multi-cell section"
        DP_Demo_BuildMultiCellSection WS

    'Build the Excel Table section
        HandlerStep = "Build table section"
        DP_Demo_BuildTableSection WS

'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Position the cursor on the first demo input cell. A Range cannot be selected
    'unless its sheet is active, and the sheet cannot be activated unless the
    'workbook has a window, so both are checked before selecting
        HandlerStep = "Finalize"
        If Wb.Windows.Count > 0 Then
            'Suppress selection failures because cursor position is cosmetic and
            'must not fail a build that has already completed
                On Error Resume Next
                Wb.Activate
                WS.Activate
                WS.Range("D13").Select
                Err.Clear
                On Error GoTo CleanFail
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup failures so they cannot mask a real error
        On Error Resume Next
    'Restore the default cursor
        Application.Cursor = xlDefault
    'Restore the captured Application state
        If FastModeOn Then DEMO_FastMode_End FastModeState
    'Release object references
        Set WS = Nothing
        Set Wb = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
        On Error GoTo 0

    'Re-raise the original error when the build failed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, _
                PROC_NAME & " | Step=" & HandlerStep, _
                "Demo sheet creation failed: " & SavedErrDescription
        End If

    'Exit after a successful build
        Exit Sub

'------------------------------------------------------------------------------
' CLEAN FAIL
'------------------------------------------------------------------------------
CleanFail:
    'Capture the escaping error before any On Error statement resets Err
        SavedErrNumber = Err.Number
        SavedErrDescription = Err.Description
    'Route through the single cleanup path
        Resume CleanExit

End Sub

Public Function DP_Demo_EnsureDemoSheet( _
    Optional ByVal TargetWorkbook As Workbook = Nothing) As Worksheet

'
'------------------------------------------------------------------------------
'                           ENSURE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the demo worksheet, building it first when it does not exist
'
' WHY THIS EXISTS
'   The Ribbon demo button should work in the add-in, where no demo sheet exists
'   until one is created. Callers need a single routine that resolves the sheet
'   without having to decide whether a build is required
'
' INPUTS
'   TargetWorkbook
'     Workbook to search and, if necessary, build into. Defaults to
'     ActiveWorkbook when omitted
'
' RETURNS
'   The demo worksheet
'
' BEHAVIOR
'   Returns an existing demo sheet unchanged. Builds one when it is missing, then
'   returns it
'
' ERROR POLICY
'   Raises a descriptive runtime error when no workbook is available or the sheet
'   cannot be built
'
' DEPENDENCIES
'   DP_Demo_ResolveTargetWorkbook
'   DP_Demo_TryGetDemoSheet
'   DP_Demo_CreateDemoSheet
'
' NOTES
'   An existing sheet is never rebuilt. A user who has entered values into the
'   demo would lose them, and this routine is reached from a Ribbon button where
'   that would be unexpected
'
'   Call DP_Demo_CreateDemoSheet directly to force a rebuild
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_EnsureDemoSheet"

    Dim Wb              As Workbook     'Resolved target workbook
    Dim WS              As Worksheet    'Resolved demo worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

    'Resolve the workbook that should hold the demo sheet
        Set Wb = DP_Demo_ResolveTargetWorkbook(TargetWorkbook)

'------------------------------------------------------------------------------
' RETURN EXISTING SHEET
'------------------------------------------------------------------------------
    'Return the existing demo sheet unchanged when it is present
        Set WS = DP_Demo_TryGetDemoSheet(Wb)
        If Not WS Is Nothing Then
            Set DP_Demo_EnsureDemoSheet = WS
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD MISSING SHEET
'------------------------------------------------------------------------------
    'Build the demo sheet into the resolved workbook
        DP_Demo_CreateDemoSheet Wb

    'Resolve the newly built demo sheet
        Set WS = DP_Demo_TryGetDemoSheet(Wb)
    'Reject a build that did not produce the expected sheet
        If WS Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Demo sheet '" & DEMO_SHEET_NAME & "' was not created"
        End If

'------------------------------------------------------------------------------
' RETURN SHEET
'------------------------------------------------------------------------------
    'Return the demo worksheet
        Set DP_Demo_EnsureDemoSheet = WS

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Demo sheet resolution failed: " & Err.Description

End Function

'
'------------------------------------------------------------------------------
'
'                                 SECTION BUILDERS
'
'------------------------------------------------------------------------------
'

Public Sub DP_Demo_FillTableColumn()

'
'------------------------------------------------------------------------------
'                       DEMO FILL TABLE COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Demo-sheet entry point for the explicit table-column fill
'
' WHY THIS EXISTS
'   DP_FillTableColumn takes the date to write, so a worksheet button cannot call
'   it directly. This wrapper supplies today's date and leaves confirmation on,
'   which is the behaviour the demo exists to show
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Fills the table column containing the current selection with today's date,
'   after the confirmation prompt has described the resolved scope
'
' ERROR POLICY
'   Best effort. Reports failures through a message box rather than raising,
'   because this is reached from a worksheet button
'
' DEPENDENCIES
'   DP_FillTableColumn
'
' NOTES
'   A selection outside a table data body is handled by DP_FillTableColumn
'   itself, which reports what is required and exits cleanly
'
'   The demo deliberately uses today's date rather than opening the picker. The
'   point being demonstrated is write scope, not date selection
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_FillTableColumn"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' FILL COLUMN
'------------------------------------------------------------------------------
    'Fill the column containing the selection, with confirmation left on
        DP_FillTableColumn VBA.Date

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Report the failure without raising out of a worksheet button
        VBA.MsgBox _
            "The table column could not be filled." & VBA.vbCrLf & VBA.vbCrLf & _
            "Error " & VBA.CStr(Err.Number) & " - " & Err.Description, _
            vbExclamation, _
            PROC_NAME

End Sub

Private Sub DP_Demo_WriteInstructions(ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                           WRITE INSTRUCTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the instruction block displayed above the demo sections
'
' WHY THIS EXISTS
'   The demo is used both as a showcase and as a manual acceptance checklist. It
'   needs to state what the reader is expected to do and to record the one
'   behaviour that surprises people, which is the in-grid icon on a protected
'   sheet
'
' INPUTS
'   WS
'     Demo worksheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the instruction lines, applies a bordered block, and italicises the
'   opening line
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
' DEPENDENCIES
'   DEMO_SetRangeBorder
'
' NOTES
'   The protected-sheet note is not incidental. The in-grid icon is a worksheet
'   Shape, so sheet protection can suppress it even when the target cell is
'   unlocked, and the workaround is not obvious
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_WriteInstructions"

    Dim NotesBlock      As Range        'Instruction block range

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
        Set NotesBlock = WS.Range(DEMO_NOTES_RANGE)

'------------------------------------------------------------------------------
' WRITE INSTRUCTION LINES
'------------------------------------------------------------------------------
    'Write the block heading
        WS.Range("C4").Value = _
            "Use this sheet to validate date, datetime, multi-cell, and table scenarios."
    'Write the instruction steps
        WS.Range("C5").Value = "Instructions:"
        WS.Range("C6").Value = "1) Select a demo input cell."
        WS.Range("C7").Value = "2) Invoke the DatePicker using the in-grid icon or right-click menu."
        WS.Range("C8").Value = _
            "3) Table write scope: picking a date in one Expiry Date cell writes " & _
            "that cell only. To fill the whole column, click Fill Table Column " & _
            "and confirm the reported scope."
    'Record the protected-sheet behaviour
        WS.Range("C9").Value = _
            "Protected-sheet note: the in-grid DatePicker icon is a worksheet Shape, " & _
            "so on a protected sheet it may not appear even when the target cell is " & _
            "unlocked. For the icon to remain available under protection, the sheet " & _
            "must allow drawing objects (for example DrawingObjects:=False, " & _
            "optionally UserInterfaceOnly:=True for VBA operations)."

'------------------------------------------------------------------------------
' FORMAT BLOCK
'------------------------------------------------------------------------------
    'Emphasise the opening line
        With WS.Range("C4").Font
            .Bold = True
            .Italic = True
        End With
    'Keep the instruction text left-aligned and unwrapped
        With NotesBlock
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .WrapText = False
            .Interior.Color = DEMO_COLOR_NOTES
        End With
    'Compact the instruction rows so the block reads as one panel
        WS.Rows("4:9").RowHeight = 15
    'Draw the block outline without inner gridlines
        DEMO_SetRangeBorder NotesBlock, 0, xlThin, False

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Instruction block creation failed: " & Err.Description

End Sub

Private Sub DP_Demo_BuildSingleCellSection(ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                       BUILD SINGLE-CELL SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the basic single-cell scenario section
'
' WHY THIS EXISTS
'   Single-cell eligibility is the core behaviour of the in-grid icon. The
'   section enumerates the cases that should qualify and the cases that should
'   not, so a tester can confirm both without reading the source
'
' INPUTS
'   WS
'     Demo worksheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the section band, column headers and scenario rows, formats the input
'   column as demo input cells, applies the required number formats, and outlines
'   the section
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
' DEPENDENCIES
'   DEMO_Write_BandHeader
'   DEMO_Format_Labels
'   DEMO_Format_InputCell
'   DEMO_SetRangeBorder
'
' NOTES
'   The formula row uses a live formula rather than a literal so that the
'   date-like detection path is exercised against a formula result
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_BuildSingleCellSection"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' WRITE SECTION HEADER
'------------------------------------------------------------------------------
    'Write the section band
        DEMO_Write_BandHeader WS.Range(DEMO_SINGLE_HEADER), "Basic Single-Cell Scenarios"
    'Write the column headers
        WS.Range("C12").Value = "Scenario"
        WS.Range("D12").Value = "Input Cell"
        WS.Range("E12").Value = "Expected"
        WS.Range("F12").Value = "Notes"
    'Format the column header row
        DEMO_Format_Labels WS.Range(DEMO_SINGLE_COLUMNS)
    'Centre the header captions over their columns
        WS.Range(DEMO_SINGLE_COLUMNS).HorizontalAlignment = xlCenter

'------------------------------------------------------------------------------
' WRITE SCENARIO ROWS
'------------------------------------------------------------------------------
    'Empty date-formatted cell
        WS.Range("C13").Value = "Empty date-formatted cell"
        WS.Range("E13").Value = "Positive"
        WS.Range("F13").Value = "Icon should appear."
    'Pre-filled date
        WS.Range("C14").Value = "Pre-filled date"
        WS.Range("D14").Value = VBA.DateSerial(2026, 4, 5)
        WS.Range("E14").Value = "Positive"
        WS.Range("F14").Value = "Uses explicit date value."
    'Pre-filled datetime
        WS.Range("C15").Value = "Pre-filled datetime"
        WS.Range("D15").Value = VBA.DateSerial(2026, 4, 6) + VBA.TimeSerial(20, 21, 13)
        WS.Range("E15").Value = "Positive"
        WS.Range("F15").Value = "Tests datetime handling."
    'Empty general-format cell
        WS.Range("C16").Value = "Empty general-format cell"
        WS.Range("E16").Value = "Usually negative"
        WS.Range("F16").Value = "General format should not qualify."
    'Text cell
        WS.Range("C17").Value = "Text cell"
        WS.Range("D17").Value = "ABC123"
        WS.Range("E17").Value = "Negative"
        WS.Range("F17").Value = "Text should not qualify."
    'Formula returning a date
        WS.Range("C18").Value = "Formula returning date"
        WS.Range("D18").Formula = "=TODAY()"
        WS.Range("E18").Value = "Positive"
        WS.Range("F18").Value = "Date-like result; formula preserved on write."
    'Empty datetime-formatted cell
        WS.Range("C19").Value = "Empty datetime-formatted cell"
        WS.Range("E19").Value = "Positive"
        WS.Range("F19").Value = "Formatted for date+time entry."

'------------------------------------------------------------------------------
' APPLY NUMBER FORMATS
'------------------------------------------------------------------------------
    'Apply the date format to the date-formatted rows
        WS.Range("D13,D14,D18").NumberFormat = DEMO_FORMAT_DATE
    'Apply the datetime format to the datetime rows
        WS.Range("D15,D19").NumberFormat = DEMO_FORMAT_DATETIME
    'Leave the general-format row deliberately unformatted
        WS.Range("D16").NumberFormat = "General"
    'Force the text row to text so it cannot be coerced to a date
        WS.Range("D17").NumberFormat = "@"

'------------------------------------------------------------------------------
' FORMAT SECTION
'------------------------------------------------------------------------------
    'Format the whole input column as demo input cells
        DEMO_Format_InputCell WS.Range(DEMO_SINGLE_INPUTS)
    'Mark the cells that should qualify for the picker in green
        WS.Range("D13,D14,D15,D18,D19").Interior.Color = DEMO_COLOR_QUALIFY
    'Leave the two non-qualifying cells in the builder's input colour, so the
    'positive and negative cases are distinguishable without reading the row
        WS.Range("D16,D17").Interior.Color = COLOR_INPUT
    'Outline the section body
        DEMO_SetRangeBorder WS.Range(DEMO_SINGLE_COLUMNS & ":" & "F19")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Single-cell section creation failed: " & Err.Description

End Sub

Private Sub DP_Demo_BuildFormatShowcaseSection(ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                       BUILD FORMAT SHOWCASE SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the number-format showcase section
'
' WHY THIS EXISTS
'   Date-cell detection depends on the cell number format, not only on the value.
'   The section renders the same underlying date through several formats so a
'   tester can confirm each one is recognised
'
' INPUTS
'   WS
'     Demo worksheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the section band, column headers and one row per format, applies each
'   format to its sample cell, and outlines the section
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
' DEPENDENCIES
'   DEMO_Write_BandHeader
'   DEMO_Format_Labels
'   DEMO_Format_OutputCell
'   DEMO_SetRangeBorder
'
' NOTES
'   Sample cells are formatted as output rather than input. They demonstrate
'   rendering and are not the surfaces the picker is meant to write into
'
'   The long-date sample renders in the user's Excel language, which is what
'   makes it useful for exercising the UseLocalNames setting
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_BuildFormatShowcaseSection"

    Dim SampleDate      As Date         'Shared sample date
    Dim SampleTime      As Date         'Shared sample time

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'Use one underlying value so every row differs only by format
        SampleDate = VBA.DateSerial(2026, 4, 5)
        SampleTime = VBA.TimeSerial(14, 35, 20)

'------------------------------------------------------------------------------
' WRITE SECTION HEADER
'------------------------------------------------------------------------------
    'Write the section band
        DEMO_Write_BandHeader WS.Range(DEMO_FORMAT_HEADER), "Format Showcase"
    'Write the column headers
        WS.Range("H12").Value = "Scenario"
        WS.Range("I12").Value = "Sample Cell"
        WS.Range("J12").Value = "Format"
        WS.Range("K12").Value = "Expected"
    'Format the column header row
        DEMO_Format_Labels WS.Range(DEMO_FORMAT_COLUMNS)
    'Centre the header captions over their columns
        WS.Range(DEMO_FORMAT_COLUMNS).HorizontalAlignment = xlCenter

'------------------------------------------------------------------------------
' WRITE FORMAT ROWS
'------------------------------------------------------------------------------
    'Standard short date
        DP_Demo_WriteFormatRow WS, 13, "dd/mm/yyyy", SampleDate, "dd/mm/yyyy"
    'Abbreviated month name
        DP_Demo_WriteFormatRow WS, 14, "dd-mmm-yyyy", SampleDate, "dd-mmm-yyyy"
    'Locale-dependent long date
        DP_Demo_WriteFormatRow WS, 15, "dddd, dd mmmm yyyy", SampleDate, "dddd, dd mmmm yyyy"
    'Date with time
        DP_Demo_WriteFormatRow WS, 16, "dd/mm/yyyy hh:mm", SampleDate + SampleTime, "dd/mm/yyyy hh:mm"
    'Time only
        DP_Demo_WriteFormatRow WS, 17, "hh:mm:ss", SampleTime, "hh:mm:ss"
    'US-ordered date
        DP_Demo_WriteFormatRow WS, 18, "mm/dd/yyyy", SampleDate, "mm/dd/yyyy"
    'ISO-ordered date
        DP_Demo_WriteFormatRow WS, 19, "yyyy-mm-dd", SampleDate, "yyyy-mm-dd"

'------------------------------------------------------------------------------
' FORMAT SECTION
'------------------------------------------------------------------------------
    'Format the sample column as demo output cells
        DEMO_Format_OutputCell WS.Range(DEMO_FORMAT_SAMPLES)
    'Every sample is a genuine date value and should qualify
        With WS.Range(DEMO_FORMAT_SAMPLES)
            .Interior.Color = DEMO_COLOR_QUALIFY
            .HorizontalAlignment = xlCenter
        End With
    'Outline the section body
        DEMO_SetRangeBorder WS.Range(DEMO_FORMAT_COLUMNS & ":" & "K19")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Format showcase section creation failed: " & Err.Description

End Sub

Private Sub DP_Demo_BuildMultiCellSection(ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                       BUILD MULTI-CELL SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the multi-cell selection section
'
' WHY THIS EXISTS
'   Write-back scope is the highest-risk behaviour in the component. The section
'   gives a tester contiguous and non-contiguous ranges to select so the written
'   scope can be observed directly
'
' INPUTS
'   WS
'     Demo worksheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the section band, column headers and scenario rows, formats the input
'   column, and outlines the section
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
' DEPENDENCIES
'   DEMO_Write_BandHeader
'   DEMO_Format_Labels
'   DEMO_Format_InputCell
'   DEMO_SetRangeBorder
'
' NOTES
'   Three rows are left empty deliberately. A tester needs spare adjacent cells
'   to build a multi-cell selection with
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_BuildMultiCellSection"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' WRITE SECTION HEADER
'------------------------------------------------------------------------------
    'Write the section band
        DEMO_Write_BandHeader WS.Range(DEMO_MULTI_HEADER), "Multi-Cell Selection"
    'Write the column headers
        WS.Range("C22").Value = "Scenario"
        WS.Range("D22").Value = "Input Range"
        WS.Range("E22").Value = "Expected"
        WS.Range("F22").Value = "Notes"
    'Format the column header row
        DEMO_Format_Labels WS.Range(DEMO_MULTI_COLUMNS)
    'Centre the header captions over their columns
        WS.Range(DEMO_MULTI_COLUMNS).HorizontalAlignment = xlCenter

'------------------------------------------------------------------------------
' WRITE SCENARIO ROWS
'------------------------------------------------------------------------------
    'Contiguous multi-cell write-back
        WS.Range("C23").Value = "Multi-cell write-back"
        WS.Range("E23").Value = "Positive"
        WS.Range("F23").Value = "All selected cells should receive the picked date."
    'Non-contiguous selection
        WS.Range("C24").Value = "Select D25:D27 together"
        WS.Range("E24").Value = "Negative"
        WS.Range("F24").Value = "Discontiguous selection is not a single target."

'------------------------------------------------------------------------------
' APPLY NUMBER FORMATS
'------------------------------------------------------------------------------
    'Format the whole input column for date entry
        WS.Range(DEMO_MULTI_INPUTS).NumberFormat = DEMO_FORMAT_DATE

'------------------------------------------------------------------------------
' FORMAT SECTION
'------------------------------------------------------------------------------
    'Format the input column as demo input cells
        DEMO_Format_InputCell WS.Range(DEMO_MULTI_INPUTS)
    'Every cell in this column is a valid write-back target
        WS.Range(DEMO_MULTI_INPUTS).Interior.Color = DEMO_COLOR_QUALIFY
    'Outline the section body
        DEMO_SetRangeBorder WS.Range(DEMO_MULTI_COLUMNS & ":" & "F27")

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Multi-cell section creation failed: " & Err.Description

End Sub

Private Sub DP_Demo_BuildTableSection(ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                       BUILD EXCEL TABLE SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the Excel Table demo section as a real ListObject
'
' WHY THIS EXISTS
'   Write-back into a structured table is the scenario behind the current
'   table-column write-scope defect. The section must be a genuine ListObject,
'   not a formatted range, or it does not exercise that path at all
'
' INPUTS
'   WS
'     Demo worksheet
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates the section band and ListObject, appends sample rows, and applies
'   date and datetime formats to the table columns
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller
'
' DEPENDENCIES
'   DEMO_Create_TableSection
'   DEMO_Table_AppendRow
'
' NOTES
'   The Expiry Date column is left empty on purpose. It is the shortest route to
'   comparing the two write scopes: picking a date in one of its cells writes
'   that cell only, while the Fill Table Column button fills the column after
'   reporting how many cells it would affect
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_BuildTableSection"

    Dim DemoTable       As ListObject   'Created demo table
    Dim RowIndex        As Long         'Sample row index
    Dim TradeDate       As Date         'Sample trade date
    Dim SettlementDate  As Date         'Sample settlement date

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler
    'Use fixed sample dates so the demo renders identically on every build
        TradeDate = VBA.DateSerial(2026, 4, 14)
        SettlementDate = VBA.DateSerial(2026, 5, 16)

'------------------------------------------------------------------------------
' CREATE TABLE
'------------------------------------------------------------------------------
    'Create the section band and the demo ListObject
        Set DemoTable = DEMO_Create_TableSection( _
            WS, _
            WS.Range(DEMO_TABLE_HEADER), _
            "Excel Table Demo", _
            WS.Range(DEMO_TABLE_TOPLEFT), _
            Array("Trade Date", "Settlement Date", "Expiry Date", "Timestamp"), _
            DEMO_TABLE_NAME, _
            DEMO_TABLE_STYLE)

    'Reject a section that did not produce a table
        If DemoTable Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Demo table '" & DEMO_TABLE_NAME & "' was not created"
        End If

'------------------------------------------------------------------------------
' APPEND SAMPLE ROWS
'------------------------------------------------------------------------------
    'Append the sample rows, leaving the expiry column empty
        For RowIndex = 1 To DEMO_TABLE_ROWS
            DEMO_Table_AppendRow DemoTable, _
                Array( _
                    TradeDate, _
                    SettlementDate, _
                    VBA.vbNullString, _
                    VBA.DateSerial(2026, 5, 13) + VBA.TimeSerial(0, 0, 0))
        Next RowIndex

'------------------------------------------------------------------------------
' APPLY COLUMN FORMATS
'------------------------------------------------------------------------------
    'Format the three date columns
        DemoTable.ListColumns("Trade Date").DataBodyRange.NumberFormat = DEMO_FORMAT_DATE
        DemoTable.ListColumns("Settlement Date").DataBodyRange.NumberFormat = DEMO_FORMAT_DATE
        DemoTable.ListColumns("Expiry Date").DataBodyRange.NumberFormat = DEMO_FORMAT_DATE
    'Format the timestamp column for date and time
        DemoTable.ListColumns("Timestamp").DataBodyRange.NumberFormat = DEMO_FORMAT_DATETIME

'------------------------------------------------------------------------------
' ADD EXPLICIT FILL BUTTON
'------------------------------------------------------------------------------
    'Give the deliberate bulk operation a visible control, so the two write
    'scopes can be compared without leaving the sheet
        DEMO_Btn_Add WS, "Btn_FillTableColumn", "Fill Table Column", _
            WS.Range("H29").Left, WS.Range("H29").Top, 110, 24, _
            "DP_Demo_FillTableColumn"

'------------------------------------------------------------------------------
' TIDY TABLE
'------------------------------------------------------------------------------
    'Remove the filter buttons. They add no value to a five-row demo and they
    'obscure the header captions at this column width
        DemoTable.ShowAutoFilter = False
    'A newly created ListObject carries one empty data row. Appending leaves it
    'stranded above the sample rows, so it is removed once the rows exist
        If DemoTable.ListRows.Count > DEMO_TABLE_ROWS Then
            If VBA.Len(VBA.CStr(DemoTable.ListRows(1).Range.Cells(1, 1).Value)) = 0 Then
                DemoTable.ListRows(1).Delete
            End If
        End If
    'Centre the table body
        DemoTable.DataBodyRange.HorizontalAlignment = xlCenter

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set DemoTable = Nothing
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Table section creation failed: " & Err.Description

End Sub

'
'------------------------------------------------------------------------------
'
'                                 PRIVATE HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub DP_Demo_WriteFormatRow( _
    ByVal WS As Worksheet, _
    ByVal RowNumber As Long, _
    ByVal ScenarioText As String, _
    ByVal SampleValue As Date, _
    ByVal FormatText As String)

'
'------------------------------------------------------------------------------
'                           WRITE FORMAT ROW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one row of the number-format showcase
'
' WHY THIS EXISTS
'   Seven rows differ only by format string and sample value. A helper keeps the
'   section builder readable and makes the rows structurally identical
'
' INPUTS
'   WS
'     Demo worksheet
'   RowNumber
'     Worksheet row to write
'   ScenarioText
'     Scenario caption
'   SampleValue
'     Underlying date or time value
'   FormatText
'     Number format applied to the sample cell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the scenario, sample value, format string and expected verdict, then
'   applies the format to the sample cell
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   None
'
' NOTES
'   The format string is written twice: once as the applied number format, once
'   as visible text, so the reader can compare the rendering against the format
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' WRITE ROW
'------------------------------------------------------------------------------
    'Write the scenario caption
        WS.Cells(RowNumber, "H").Value = ScenarioText
    'Write the sample value
        WS.Cells(RowNumber, "I").Value = SampleValue
    'Write the format string as visible text
        WS.Cells(RowNumber, "J").Value = FormatText
    'Write the expected verdict
        WS.Cells(RowNumber, "K").Value = "Positive"

'------------------------------------------------------------------------------
' APPLY FORMAT
'------------------------------------------------------------------------------
    'Apply the demonstrated format to the sample cell
        WS.Cells(RowNumber, "I").NumberFormat = FormatText

End Sub

Private Function DP_Demo_ResolveTargetWorkbook( _
    ByVal TargetWorkbook As Workbook) As Workbook

'
'------------------------------------------------------------------------------
'                       RESOLVE TARGET WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the workbook that should receive the demo sheet
'
' WHY THIS EXISTS
'   The demo builder must work identically whether the code is embedded in a
'   workbook or loaded as an add-in. ThisWorkbook resolves to the add-in, which
'   cannot hold a demo sheet, so it must never be used as the fallback
'
' INPUTS
'   TargetWorkbook
'     Explicit target, or Nothing to use ActiveWorkbook
'
' RETURNS
'   Resolved workbook
'
' BEHAVIOR
'   Returns the supplied workbook when present. Otherwise returns ActiveWorkbook
'
' ERROR POLICY
'   Raises a descriptive runtime error when no workbook is available or the
'   resolved workbook cannot accept a new worksheet
'
' DEPENDENCIES
'   None
'
' NOTES
'   An add-in has no worksheets and reports ThisWorkbook.IsAddin as True. That
'   case is rejected explicitly rather than allowed to fail later with an opaque
'   subscript error
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_DP_DEMO.DP_Demo_ResolveTargetWorkbook"

    Dim ResolvedBook    As Workbook     'Resolved target workbook

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE WORKBOOK
'------------------------------------------------------------------------------
    'Use the supplied workbook when one was provided
        If Not TargetWorkbook Is Nothing Then
            Set ResolvedBook = TargetWorkbook
        Else
            Set ResolvedBook = Excel.Application.ActiveWorkbook
        End If

'------------------------------------------------------------------------------
' VALIDATE WORKBOOK
'------------------------------------------------------------------------------
    'Reject the case where no workbook is open
        If ResolvedBook Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "No workbook is open. Open or create a workbook before building the demo sheet."
        End If
    'Reject an add-in, which cannot hold a demo worksheet
        If ResolvedBook.IsAddin Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "'" & ResolvedBook.Name & "' is an add-in and cannot hold the demo sheet."
        End If
    'Reject a workbook whose structure is protected
        If ResolvedBook.ProtectStructure Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Workbook structure of '" & ResolvedBook.Name & "' is protected."
        End If

'------------------------------------------------------------------------------
' RETURN WORKBOOK
'------------------------------------------------------------------------------
    'Return the validated workbook
        Set DP_Demo_ResolveTargetWorkbook = ResolvedBook

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

Private Function DP_Demo_TryGetDemoSheet( _
    ByVal Wb As Workbook) As Worksheet

'
'------------------------------------------------------------------------------
'                       TRY GET DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the demo worksheet when it exists in the supplied workbook
'
' WHY THIS EXISTS
'   Callers need to distinguish a missing demo sheet from a failed lookup without
'   treating absence as an error
'
' INPUTS
'   Wb
'     Workbook to search
'
' RETURNS
'   The demo worksheet, or Nothing when it is not present
'
' BEHAVIOR
'   Probes the Worksheets collection by name and suppresses the lookup error
'
' ERROR POLICY
'   Does not raise. Absence is a normal outcome
'
' DEPENDENCIES
'   None
'
' NOTES
'   The lookup is by name, so a demo sheet renamed by the user is not found and
'   a new one is built alongside it
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FoundSheet      As Worksheet    'Demo worksheet when present

'------------------------------------------------------------------------------
' PROBE FOR SHEET
'------------------------------------------------------------------------------
    'Suppress the lookup error raised when the sheet does not exist
        On Error Resume Next
        Set FoundSheet = Wb.Worksheets(DEMO_SHEET_NAME)
        Err.Clear
        On Error GoTo 0

'------------------------------------------------------------------------------
' RETURN RESULT
'------------------------------------------------------------------------------
    'Return the demo worksheet, or Nothing when it was not found
        Set DP_Demo_TryGetDemoSheet = FoundSheet

End Function
