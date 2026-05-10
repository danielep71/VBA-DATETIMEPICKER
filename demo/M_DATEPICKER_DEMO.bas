Attribute VB_Name = "M_DatePicker_Demo"
'------------------------------------------------------------------------------
' MODULE: M_DATEPICKER_DEMO
'------------------------------------------------------------------------------
' PURPOSE
'   Builds and saves the demo worksheet for the DatePicker project
'
' WHY THIS EXISTS
'   The DatePicker project needs one repeatable demo surface that can be rebuilt
'   deterministically and used to exercise the main user-facing configuration
'   options without manually editing hidden registry-backed settings
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - demo-sheet rebuild
'     - demo named-range cleanup
'     - settings control-panel creation
'     - demo button creation
'     - current setting value synchronization
'     - settings parsing and persistence from the demo sheet
'     - best-effort refresh of right-click menu and grid icon state
'
' ERROR POLICY
'   Public routines raise descriptive runtime errors after deterministic cleanup
'   when cleanup is required
'
' DEPENDENCIES
'   Excel object model
'   M_DATEPICKER
'   M_DEMO_BUILDER
'   cDatePickerManager
'
' NOTES
'   This module deliberately avoids Option Private Module because demo buttons
'   and shape callbacks may need to call public routines from the Excel UI
'
'   Named ranges are always accessed through ThisWorkbook to avoid accidental
'   dependency on ActiveWorkbook or ActiveSheet
'
'   The DatePicker itself writes selected calendar dates as date-only values.
'   The demo exposes Today and Now through footer shortcuts in the form, not as
'   a date-selection mode
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

Option Explicit

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------

    '-------------------------------SHEET--------------------------------------
    Private Const DEMO_SHEET_NAME                    As String = "DATE PICKER DEMO"         'Demo worksheet name
    Private Const DEMO_TITLE_TEXT                    As String = "DATE PICKER DEMO"         'Demo sheet title
    Private Const DEMO_SUBTITLE_TEXT                 As String = "DATE PICKER"              'Demo sheet subtitle
    Private Const DEMO_TEMPLATE_CONTEXT              As String = "Demo Sheet"               'Template context label

    '-----------------------------NAMED RANGES---------------------------------
    Private Const DEMO_NM_FIRST_DAY                 As String = "M_FirstDayOfWeek"        'First-day named range
    Private Const DEMO_NM_LOCAL_NAMES               As String = "M_LocalNames"            'Local-name named range
    Private Const DEMO_NM_LIVE_CLOCK                As String = "M_LiveClock"             'Live-clock named range
    Private Const DEMO_NM_COMPACT_LAYOUT            As String = "M_CompactLayout"         'Compact-layout named range
    Private Const DEMO_NM_HIGHLIGHT_WEEKENDS        As String = "M_HighlightWeekends"     'Highlight-weekends named range
    Private Const DEMO_NM_ALLOW_OUTSIDE              As String = "M_AllowOutside"         'Outside-month named range

    Private Const DEMO_NM_RIGHT_CLICK               As String = "M_RightClick"            'Right-click named range
    Private Const DEMO_NM_GRID_ICON                 As String = "M_IconInGrid"            'Grid-icon named range
    Private Const DEMO_NM_USE_WINAPI                As String = "M_UseWinAPI"             'WinAPI setting named range
    Private Const DEMO_NM_HOLIDAY_CALLBACK          As String = "M_HolidayCallback"       'Holiday callback named range

    Private Const DEMO_NM_CLOSE_AFTER_SELECTION     As String = "M_CloseAfterSelection"   'Close-after-selection named range
    
    '-------------------------------INPUTS-------------------------------------
    Private Const DEMO_FIRST_DAY_LIST               As String = "vbMonday,vbSunday"         'First-day validation list
    Private Const DEMO_BOOLEAN_LIST                 As String = "TRUE,FALSE"                'Boolean validation list

    '-------------------------------BUTTONS------------------------------------
    Private Const DEMO_BTN_SHOW_NAME                As String = "Btn_ShowDP"                'Show button name
    Private Const DEMO_BTN_SHOW_CAPTION             As String = "SHOW DATE PICKER"          'Show button caption
    Private Const DEMO_BTN_SAVE_NAME                As String = "Btn_SaveSettings"          'Save button name
    Private Const DEMO_BTN_SAVE_CAPTION             As String = "SAVE SETTINGS"             'Save button caption

Public Sub Demo_CreateDemoSheet()

'
'------------------------------------------------------------------------------
'                           CREATE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Builds or rebuilds the DatePicker demo sheet and its main demo assets
'
' WHY THIS EXISTS
'   The DatePicker project needs one repeatable entry point that prepares a clean
'   demo surface, exposes the main runnable scenario, and provides grouped
'   settings panels for user-configurable DatePicker preferences
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Targets ThisWorkbook
'   - Loads the current DatePicker settings
'   - Enters fast mode
'   - Rebuilds the demo sheet template
'   - Deletes stale demo named ranges
'   - Builds three settings sections:
'       - Display settings
'       - Behavior settings
'       - Integration settings
'   - Adds demo buttons
'   - Writes current setting values to the input cells
'   - Selects the first input cell after rebuild to show the validation dropdown
'
' ERROR POLICY
'   Structured cleanup:
'     - captures the original error
'     - restores cursor and fast mode during cleanup
'     - re-raises the original error after cleanup
'
' DEPENDENCIES
'   DEMO_FastMode_Begin
'   DEMO_FastMode_End
'   DEMO_Sheet_BuildTemplate
'   Demo_DeleteDemoNames
'   Demo_BuildDisplaySettingsSection
'   Demo_BuildBehaviorSettingsSection
'   Demo_BuildIntegrationSettingsSection
'   Demo_AddDemoButtons
'   Demo_WriteCurrentSettings
'   M_Settings_Load
'
' NOTES
'   Named ranges are workbook-qualified throughout the routine
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_CreateDemoSheet"        'Current procedure name

    Dim Wb                     As Workbook                               'Target workbook
    Dim WS                     As Worksheet                              'Demo worksheet
    Dim FastModeState          As tDEMOFastModeState                     'Saved Application-state snapshot
    Dim FastModeOn             As Boolean                                'True when fast mode was entered
    Dim SavedErrNumber         As Long                                   'Captured error number
    Dim SavedErrSource         As String                                 'Captured error source
    Dim SavedErrDescription    As String                                 'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo Clean_Fail
    'Target the workbook that contains this module
        Set Wb = ThisWorkbook
    'Load current DatePicker settings before building the control panel
        M_Settings_Load
    'Capture and apply fast-mode Application settings
        DEMO_FastMode_Begin FastModeState
    'Mark that fast mode was entered successfully
        FastModeOn = True
    'Show the wait cursor while rebuilding the demo sheet
        Application.Cursor = xlWait

'------------------------------------------------------------------------------
' BUILD DEMO SHEET
'------------------------------------------------------------------------------
    'Build or rebuild the generic template for the demo sheet
        DEMO_Sheet_BuildTemplate _
            DEMO_SHEET_NAME, _
            DEMO_TITLE_TEXT, _
            DEMO_TEMPLATE_CONTEXT

    'Resolve the demo worksheet after template preparation
        Set WS = Wb.Worksheets(DEMO_SHEET_NAME)
    'Delete stale demo named ranges before recreating them
        Demo_DeleteDemoNames Wb

'------------------------------------------------------------------------------
' LOCAL FORMAT ADJUSTMENTS
'------------------------------------------------------------------------------
    'Set the display-setting label column width
        WS.Columns("C:C").ColumnWidth = 30
    'Set the display-setting input column width
        WS.Columns("D:D").ColumnWidth = 18
    'Set the behavior-setting label column width
        WS.Columns("F:F").ColumnWidth = 30
    'Set the behavior-setting input column width
        WS.Columns("G:G").ColumnWidth = 18
    'Set the integration-setting label column width
        WS.Columns("I:I").ColumnWidth = 30
    'Set the integration-setting input column width
        WS.Columns("J:J").ColumnWidth = 18

'------------------------------------------------------------------------------
' BUILD SETTINGS SECTIONS
'------------------------------------------------------------------------------
    'Build the display settings section
        Demo_BuildDisplaySettingsSection Wb, WS
    'Build the behavior settings section
        Demo_BuildBehaviorSettingsSection Wb, WS
    'Build the integration settings section
        Demo_BuildIntegrationSettingsSection Wb, WS

'------------------------------------------------------------------------------
' ADD DEMO BUTTONS
'------------------------------------------------------------------------------
    'Add the demo action buttons
        Demo_AddDemoButtons Wb, WS

'------------------------------------------------------------------------------
' WRITE CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Write the current DatePicker settings to the demo input cells
        Demo_WriteCurrentSettings Wb

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
Clean_Exit:

    'Protect cleanup so it cannot overwrite the original error
        On Error Resume Next
    'Restore the normal cursor
        Application.Cursor = xlDefault
    'Restore the original Excel Application state only when fast mode was entered
        If FastModeOn Then
            DEMO_FastMode_End FastModeState
        End If
    'Activate the demo sheet and select the first input after a successful rebuild
        If SavedErrNumber = 0 Then
            If Not WS Is Nothing Then
                WS.Activate
                Wb.Names(DEMO_NM_FIRST_DAY).RefersToRange.Select
                DoEvents
            End If
        End If
    'Restore normal error handling before any re-raise
        On Error GoTo 0
    'Re-raise the original error after cleanup when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
    'Normal termination point
        Exit Sub

'------------------------------------------------------------------------------
' CLEAN FAIL
'------------------------------------------------------------------------------
Clean_Fail:

    'Capture the original error number before cleanup
        SavedErrNumber = Err.Number
    'Capture the original error source before cleanup
        If Len(Err.Source) > 0 Then
            SavedErrSource = Err.Source
        Else
            SavedErrSource = PROC_NAME
        End If
    'Capture the original error description before cleanup
        SavedErrDescription = Err.Description
    'Continue through the centralized cleanup path
        Resume Clean_Exit

End Sub
Private Sub Demo_BuildDisplaySettingsSection( _
    ByVal Wb As Workbook, _
    ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                           BUILD DISPLAY SETTINGS SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the display settings block on the DatePicker demo sheet
'
' WHY THIS EXISTS
'   Display settings control how the DatePicker appears and how captions are
'   rendered, without changing core write-back behavior
'
' INPUTS
'   Wb
'     Workbook that owns the demo named ranges
'
'   WS
'     Demo worksheet being built
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates labeled input rows for:
'     - first day of week
'     - local names
'     - live clock
'     - compact layout
'
' ERROR POLICY
'   Propagates errors raised by the demo builder routines
'
' DEPENDENCIES
'   DEMO_Prepare_LabeledInputSection
'   DEMO_Write_NamedInputRow
'
' NOTES
'   Live clock is stored as TRUE / FALSE and mapped to DP_ClockMode on save
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD SECTION FRAME
'------------------------------------------------------------------------------
    'Apply the standard section, label, and input formatting
        DEMO_Prepare_LabeledInputSection _
            WS, _
            WS.Range("C4:D4"), _
            "DISPLAY SETTINGS", _
            WS.Range("C5:C9"), _
            WS.Range("D5:D9")

'------------------------------------------------------------------------------
' WRITE DISPLAY SETTINGS ROWS
'------------------------------------------------------------------------------
    'Write the first-day-of-week input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("C5"), WS.Range("D5"), _
            "First day of week", vbNullString, DEMO_NM_FIRST_DAY, _
            DemoInputValidationList, DEMO_FIRST_DAY_LIST
    'Write the local-name input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("C6"), WS.Range("D6"), _
            "Use local names", vbNullString, DEMO_NM_LOCAL_NAMES, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the live-clock input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("C7"), WS.Range("D7"), _
            "Live clock", vbNullString, DEMO_NM_LIVE_CLOCK, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the compact-layout input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("C8"), WS.Range("D8"), _
            "Compact layout", vbNullString, DEMO_NM_COMPACT_LAYOUT, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the weekend-highlight input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("C9"), WS.Range("D9"), _
            "Highlight weekends", vbNullString, DEMO_NM_HIGHLIGHT_WEEKENDS, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST

End Sub
Private Sub Demo_BuildBehaviorSettingsSection( _
    ByVal Wb As Workbook, _
    ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                           BUILD BEHAVIOR SETTINGS SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the behavior settings block on the DatePicker demo sheet
'
' WHY THIS EXISTS
'   Behavior settings control how the DatePicker responds to user selections
'
' INPUTS
'   Wb
'     Workbook that owns the demo named ranges
'
'   WS
'     Demo worksheet being built
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates labeled input rows for behavior-related settings currently supported
'   by the companion module
'
' ERROR POLICY
'   Propagates errors raised by the demo builder routines
'
' DEPENDENCIES
'   DEMO_Prepare_LabeledInputSection
'   DEMO_Write_NamedInputRow
'
' NOTES
'   Additional future behavior settings can be added here, for example:
'     - allow weekend selection
'     - close after selection
'     - expand table column
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD SECTION FRAME
'------------------------------------------------------------------------------
    'Apply the standard section, label, and input formatting
        DEMO_Prepare_LabeledInputSection _
            WS, _
            WS.Range("F4:G4"), _
            "BEHAVIOR SETTINGS", _
            WS.Range("F5:F6"), _
            WS.Range("G5:G6")

'------------------------------------------------------------------------------
' WRITE BEHAVIOR SETTINGS ROWS
'------------------------------------------------------------------------------
    'Write the outside-month selection input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("F5"), WS.Range("G5"), _
            "Allow outside-month selection", vbNullString, DEMO_NM_ALLOW_OUTSIDE, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the close-after-selection input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("F6"), WS.Range("G6"), _
            "Close after selection", vbNullString, DEMO_NM_CLOSE_AFTER_SELECTION, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST

End Sub
Private Sub Demo_BuildIntegrationSettingsSection( _
    ByVal Wb As Workbook, _
    ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                           BUILD INTEGRATION SETTINGS SECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the integration settings block on the DatePicker demo sheet
'
' WHY THIS EXISTS
'   Integration settings control how the DatePicker connects to Excel UI surfaces
'   and platform-specific behavior
'
' INPUTS
'   Wb
'     Workbook that owns the demo named ranges
'
'   WS
'     Demo worksheet being built
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates labeled input rows for:
'     - right-click menu
'     - in-grid icon
'     - WinAPI styling
'     - holiday callback
'
' ERROR POLICY
'   Propagates errors raised by the demo builder routines
'
' DEPENDENCIES
'   DEMO_Prepare_LabeledInputSection
'   DEMO_Write_NamedInputRow
'
' NOTES
'   Holiday callback is exposed as free text and may be left blank
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BUILD SECTION FRAME
'------------------------------------------------------------------------------
    'Apply the standard section, label, and input formatting
        DEMO_Prepare_LabeledInputSection _
            WS, _
            WS.Range("I4:J4"), _
            "INTEGRATION SETTINGS", _
            WS.Range("I5:I8"), _
            WS.Range("J5:J8")

'------------------------------------------------------------------------------
' WRITE INTEGRATION SETTINGS ROWS
'------------------------------------------------------------------------------
    'Write the right-click input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("I5"), WS.Range("J5"), _
            "Right-click menu", vbNullString, DEMO_NM_RIGHT_CLICK, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the grid-icon input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("I6"), WS.Range("J6"), _
            "In-grid icon", vbNullString, DEMO_NM_GRID_ICON, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the WinAPI styling input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("I7"), WS.Range("J7"), _
            "Use WinAPI styling", vbNullString, DEMO_NM_USE_WINAPI, _
            DemoInputValidationList, DEMO_BOOLEAN_LIST
    'Write the holiday callback free-text input row
        DEMO_Write_NamedInputRow Wb, WS, WS.Range("I8"), WS.Range("J8"), _
            "Holiday callback", vbNullString, DEMO_NM_HOLIDAY_CALLBACK

End Sub
Private Sub Demo_AddDemoButtons( _
    ByVal Wb As Workbook, _
    ByVal WS As Worksheet)

'
'------------------------------------------------------------------------------
'                           ADD DEMO BUTTONS
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the DatePicker demo action buttons
'
' WHY THIS EXISTS
'   The demo sheet needs explicit buttons to launch the DatePicker and save the
'   visible settings panel
'
' INPUTS
'   Wb
'     Workbook that owns the callback macros
'
'   WS
'     Demo worksheet receiving the buttons
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds:
'     - SHOW DATE PICKER
'     - SAVE SETTINGS
'
' ERROR POLICY
'   Propagates errors raised by DEMO_Btn_Add
'
' DEPENDENCIES
'   DEMO_Btn_Add
'   Demo_GetQualifiedMacroName
'
' NOTES
'   Button callbacks are workbook-qualified
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ADD BUTTONS
'------------------------------------------------------------------------------
    'Add the show DatePicker demo button
        DEMO_Btn_Add _
            WS, _
            DEMO_BTN_SHOW_NAME, _
            DEMO_BTN_SHOW_CAPTION, _
            WS.Range("F8").Left, _
            WS.Range("F8").Top, _
            150, _
            25, _
            Demo_GetQualifiedMacroName(Wb, "DP_Show")

    'Add the save settings button
        DEMO_Btn_Add _
            WS, _
            DEMO_BTN_SAVE_NAME, _
            DEMO_BTN_SAVE_CAPTION, _
            WS.Range("F10").Left, _
            WS.Range("F10").Top, _
            150, _
            25, _
            Demo_GetQualifiedMacroName(Wb, "Demo_SaveSettings")

End Sub
Public Sub Demo_SaveSettings()

'
'------------------------------------------------------------------------------
'                           SAVE DEMO SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Saves DatePicker settings selected in the demo worksheet control panel
'
' WHY THIS EXISTS
'   The demo sheet exposes user-editable settings through named input cells.
'   This routine reads the selected dropdown values, updates the shared
'   DatePicker settings, persists them, and applies feature changes where needed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads and applies:
'     - first day of week
'     - local names
'     - static / live clock
'     - outside-month selection
'     - right-click integration
'     - in-grid icon
'     - normal / compact size
'     - WinAPI enabled / Mac-safe mode
'
' ERROR POLICY
'   Raises a descriptive runtime error if a named range is missing, contains an
'   unsupported value, or if settings cannot be saved/applied
'
' DEPENDENCIES
'   M_Settings_Load
'   M_Settings_SetFirstDayOfWeek
'   M_Settings_SetUseLocalNames
'   M_Settings_SetClockMode
'   M_Settings_SetAllowOutsideMonthSelection
'   M_Settings_SetShowRightClick
'   M_Settings_SetShowGridIcon
'   M_Settings_SetSizeMode
'   M_Settings_Save
'   DEMO_Btn_PlayFeedback
'
' NOTES
'   Calendar date selection remains date-only. The Now shortcut writes a
'   date-time value directly and is not controlled by a value-mode setting
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_SaveSettings"           'Current procedure name

    Dim Wb                     As Workbook                               'Target workbook
    
    Dim FirstDayOfWeek         As Long                                   'Parsed first-day setting
    Dim UseLocalNames          As Boolean                                'Parsed local-name setting
    Dim UseLiveClock           As Boolean                                'Parsed live-clock setting
    Dim UseCompactLayout       As Boolean                                'Parsed compact-layout setting
    Dim AllowOutsideMonth      As Boolean                                'Parsed outside-month setting
    Dim ShowRightClick         As Boolean                                'Parsed right-click setting
    Dim ShowGridIcon           As Boolean                                'Parsed grid-icon setting
    Dim UseWinAPIStyling       As Boolean                                'Parsed WinAPI setting
    Dim HolidayCallbackName    As String                                 'Parsed holiday callback name
    Dim HighlightWeekends      As Boolean                                'Parsed weekend-highlight setting
    Dim CloseAfterSelection    As Boolean                                'Parsed close-after-selection setting
    
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Target the workbook that contains this module
        Set Wb = ThisWorkbook
    'Load current settings before applying demo changes
        M_Settings_Load

'------------------------------------------------------------------------------
' READ AND PARSE DEMO SETTINGS
'------------------------------------------------------------------------------
    'Parse the selected first-day-of-week value
        FirstDayOfWeek = Demo_ReadFirstDayOfWeek(Wb, DEMO_NM_FIRST_DAY)
    'Parse the selected local-name value
        UseLocalNames = Demo_ReadNamedBoolean(Wb, DEMO_NM_LOCAL_NAMES)
    'Parse the selected live-clock value
        UseLiveClock = Demo_ReadNamedBoolean(Wb, DEMO_NM_LIVE_CLOCK)
    'Parse the selected compact-layout value
        UseCompactLayout = Demo_ReadNamedBoolean(Wb, DEMO_NM_COMPACT_LAYOUT)
    'Parse the selected weekend-highlight value
        HighlightWeekends = Demo_ReadNamedBoolean(Wb, DEMO_NM_HIGHLIGHT_WEEKENDS)
    'Parse the selected outside-month value
        AllowOutsideMonth = Demo_ReadNamedBoolean(Wb, DEMO_NM_ALLOW_OUTSIDE)
    'Parse the selected close-after-selection value
        CloseAfterSelection = Demo_ReadNamedBoolean(Wb, DEMO_NM_CLOSE_AFTER_SELECTION)
    'Parse the selected right-click value
        ShowRightClick = Demo_ReadNamedBoolean(Wb, DEMO_NM_RIGHT_CLICK)
    'Parse the selected grid-icon value
        ShowGridIcon = Demo_ReadNamedBoolean(Wb, DEMO_NM_GRID_ICON)
    'Parse the selected WinAPI value
        UseWinAPIStyling = Demo_ReadNamedBoolean(Wb, DEMO_NM_USE_WINAPI)
    'Read the optional holiday callback name
        HolidayCallbackName = Demo_GetNamedText(Wb, DEMO_NM_HOLIDAY_CALLBACK)

'------------------------------------------------------------------------------
' APPLY DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Apply the first-day-of-week setting
        M_Settings_SetFirstDayOfWeek FirstDayOfWeek
    'Apply the local-name setting
        M_Settings_SetUseLocalNames UseLocalNames
    'Apply the clock mode setting
        If UseLiveClock Then
            M_Settings_SetClockMode DP_ClockMode_Live
        Else
            M_Settings_SetClockMode DP_ClockMode_Static
        End If
    'Apply the size-mode setting
        If UseCompactLayout Then
            M_Settings_SetSizeMode DP_SizeMode_Compact
        Else
            M_Settings_SetSizeMode DP_SizeMode_Normal
        End If

'------------------------------------------------------------------------------
' APPLY BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Apply the outside-month selection setting
        M_Settings_SetAllowOutsideMonthSelection AllowOutsideMonth
    'Apply the close-after-selection setting
        M_Settings_SetCloseAfterSelection CloseAfterSelection
    'Apply the weekend-highlight setting
        M_Settings_SetHighlightWeekends HighlightWeekends
        
'------------------------------------------------------------------------------
' APPLY INTEGRATION SETTINGS
'------------------------------------------------------------------------------
    'Apply the right-click integration setting
        M_Settings_SetShowRightClick ShowRightClick
    'Apply the grid-icon setting
        M_Settings_SetShowGridIcon ShowGridIcon
    'Apply the WinAPI styling setting
        gDP_UseWinAPI = UseWinAPIStyling
    'Apply the holiday callback setting
        M_Settings_SetHolidayCallback HolidayCallbackName

    'Persist the final combined settings
        M_Settings_Save

'------------------------------------------------------------------------------
' REFRESH DEMO / FEATURE STATE
'------------------------------------------------------------------------------
    'Reload normalized settings after save
        M_Settings_Load
    'Write normalized values back to the demo sheet
        Demo_WriteCurrentSettings Wb
    'Refresh the grid icon state when the manager is available
        Demo_TryRefreshGridIconState
    'Play demo button feedback after successful save
        DEMO_Btn_PlayFeedback

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

'------------------------------------------------------------------------------
' PRIVATE SETTINGS HELPERS
'------------------------------------------------------------------------------

Private Sub Demo_WriteCurrentSettings(ByVal Wb As Workbook)

'
'------------------------------------------------------------------------------
'                           WRITE CURRENT SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the currently persisted DatePicker settings to the demo sheet inputs
'
' WHY THIS EXISTS
'   When the demo sheet is rebuilt, its input cells must reflect the actual
'   DatePicker settings currently loaded from the registry, not hard-coded
'   defaults or stale in-memory values
'
' INPUTS
'   Wb
'     Workbook that owns the demo named ranges
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads the current DatePicker settings, then writes all supported demo
'   settings to their workbook-level named ranges
'
' ERROR POLICY
'   Raises a descriptive runtime error if Wb is missing, if settings cannot be
'   loaded, or if one or more demo named ranges cannot be written
'
' DEPENDENCIES
'   M_Settings_Load
'   Demo_SetNamedValue
'   Demo_BoolToText
'
' NOTES
'   This routine must be called after the demo input rows and named ranges have
'   been created
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_WriteCurrentSettings"   'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing workbook
        If Wb Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Wb cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Load persisted DatePicker settings before writing them to the sheet
        M_Settings_Load

'------------------------------------------------------------------------------
' WRITE DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Write the current first-day-of-week value
        Demo_SetNamedValue Wb, DEMO_NM_FIRST_DAY, _
            IIf(gDP_FirstDayOfWeek = vbSunday, "vbSunday", "vbMonday")

    'Write the current local-name value
        Demo_SetNamedValue Wb, DEMO_NM_LOCAL_NAMES, _
            Demo_BoolToText(gDP_UseLocalNames)

    'Write the current live-clock value
        Demo_SetNamedValue Wb, DEMO_NM_LIVE_CLOCK, _
            Demo_BoolToText(gDP_ClockMode = DP_ClockMode_Live)

    'Write the current compact-layout value
        Demo_SetNamedValue Wb, DEMO_NM_COMPACT_LAYOUT, _
            Demo_BoolToText(gDP_SizeMode = DP_SizeMode_Compact)

    'Write the current weekend-highlight value
        Demo_SetNamedValue Wb, DEMO_NM_HIGHLIGHT_WEEKENDS, _
            Demo_BoolToText(gDP_HighlightWeekends)

'------------------------------------------------------------------------------
' WRITE BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Write the current outside-month selection value
        Demo_SetNamedValue Wb, DEMO_NM_ALLOW_OUTSIDE, _
            Demo_BoolToText(gDP_AllowOutsideMonthSelection)

    'Write the current close-after-selection value
        Demo_SetNamedValue Wb, DEMO_NM_CLOSE_AFTER_SELECTION, _
            Demo_BoolToText(gDP_CloseAfterSelection)

'------------------------------------------------------------------------------
' WRITE INTEGRATION SETTINGS
'------------------------------------------------------------------------------
    'Write the current right-click integration value
        Demo_SetNamedValue Wb, DEMO_NM_RIGHT_CLICK, _
            Demo_BoolToText(gDP_ShowRightClick)

    'Write the current grid-icon value
        Demo_SetNamedValue Wb, DEMO_NM_GRID_ICON, _
            Demo_BoolToText(gDP_ShowGridIcon)

    'Write the current WinAPI styling value
        Demo_SetNamedValue Wb, DEMO_NM_USE_WINAPI, _
            Demo_BoolToText(gDP_UseWinAPI)

    'Write the current holiday callback value
        Demo_SetNamedValue Wb, DEMO_NM_HOLIDAY_CALLBACK, _
            Trim$(gDP_HolidayCallbackName)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub
Private Sub Demo_DeleteDemoNames(ByVal Wb As Workbook)

'
'------------------------------------------------------------------------------
'                           DELETE DEMO NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes workbook-level named ranges owned by the DatePicker demo sheet
'
' WHY THIS EXISTS
'   A full demo-sheet rebuild should not leave stale named-range definitions
'
' INPUTS
'   Wb
'     Workbook whose demo names should be deleted
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Attempts to delete each known demo input name
'
' ERROR POLICY
'   Suppresses missing-name errors because cleanup is best-effort
'
' DEPENDENCIES
'   Demo_DeleteWorkbookNameIfExists
'
' NOTES
'   The names are recreated later by DEMO_Write_NamedInputRow
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELETE DISPLAY NAMES
'------------------------------------------------------------------------------
    'Delete the first-day-of-week name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_FIRST_DAY
    'Delete the local-name name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_LOCAL_NAMES
    'Delete the live-clock name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_LIVE_CLOCK
    'Delete the compact-layout name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_COMPACT_LAYOUT
    'Delete the weekend-highlight name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_HIGHLIGHT_WEEKENDS

'------------------------------------------------------------------------------
' DELETE BEHAVIOR NAMES
'------------------------------------------------------------------------------
    'Delete the outside-month name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_ALLOW_OUTSIDE
    'Delete the close-after-selection name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_CLOSE_AFTER_SELECTION

'------------------------------------------------------------------------------
' DELETE INTEGRATION NAMES
'------------------------------------------------------------------------------
    'Delete the right-click name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_RIGHT_CLICK
    'Delete the grid-icon name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_GRID_ICON
    'Delete the WinAPI name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_USE_WINAPI
    'Delete the holiday-callback name
        Demo_DeleteWorkbookNameIfExists Wb, DEMO_NM_HOLIDAY_CALLBACK

End Sub
Private Sub Demo_DeleteWorkbookNameIfExists( _
    ByVal Wb As Workbook, _
    ByVal NameText As String)

'
'------------------------------------------------------------------------------
'                           DELETE WORKBOOK NAME IF EXISTS
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes one workbook-level name when present
'
' WHY THIS EXISTS
'   Demo rebuild cleanup should be tolerant of names that do not yet exist
'
' INPUTS
'   Wb
'     Workbook that owns the name
'
'   NameText
'     Name to delete
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the name if it exists and silently ignores missing names
'
' ERROR POLICY
'   Safe cleanup. Does not raise outward
'
' DEPENDENCIES
'   Excel Names collection
'
' NOTES
'   This routine is intentionally best-effort
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAN UP NAME
'------------------------------------------------------------------------------
    'Suppress missing-name errors
        On Error Resume Next

    'Delete the requested workbook name
        Wb.Names(NameText).Delete

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub Demo_SetNamedValue( _
    ByVal Wb As Workbook, _
    ByVal NameText As String, _
    ByVal ValueToWrite As Variant)

'
'------------------------------------------------------------------------------
'                           SET NAMED VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a value to a workbook-level named range
'
' WHY THIS EXISTS
'   Demo settings should be written through workbook-qualified names only
'
' INPUTS
'   Wb
'     Workbook that owns the named range
'
'   NameText
'     Workbook-level name to write
'
'   ValueToWrite
'     Value to write to the target range
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes ValueToWrite to the named range's RefersToRange
'
' ERROR POLICY
'   Raises a descriptive runtime error if the name cannot be resolved or written
'
' DEPENDENCIES
'   Excel Names collection
'
' NOTES
'   This helper intentionally does not use ActiveWorkbook or ActiveSheet
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_SetNamedValue"          'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing workbook
        If Wb Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Wb cannot be Nothing."
        End If

    'Reject an empty name
        If Len(Trim$(NameText)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "NameText cannot be empty."
        End If

'------------------------------------------------------------------------------
' WRITE VALUE
'------------------------------------------------------------------------------
    'Write the value to the named range
        Wb.Names(NameText).RefersToRange.Value = ValueToWrite

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Unable to write demo named range '" & NameText & "'. " & Err.Description

End Sub

Private Function Demo_GetNamedText( _
    ByVal Wb As Workbook, _
    ByVal NameText As String) As String

'
'------------------------------------------------------------------------------
'                           GET NAMED TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the trimmed text value from a workbook-level named range
'
' WHY THIS EXISTS
'   Demo setting parsing should read from workbook-qualified names only
'
' INPUTS
'   Wb
'     Workbook that owns the named range
'
'   NameText
'     Workbook-level name to read
'
' RETURNS
'   Trimmed text value
'
' BEHAVIOR
'   Reads the named range's current value and converts it to trimmed text
'
' ERROR POLICY
'   Raises a descriptive runtime error if the name cannot be resolved or read
'
' DEPENDENCIES
'   Excel Names collection
'
' NOTES
'   This helper intentionally does not use Range("name") because that can bind
'   to the active workbook or active sheet
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_GetNamedText"           'Current procedure name

    Dim RawValue               As Variant                                'Raw named-range value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing workbook
        If Wb Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Wb cannot be Nothing."
        End If

    'Reject an empty name
        If Len(Trim$(NameText)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "NameText cannot be empty."
        End If

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Read the named range value
        RawValue = Wb.Names(NameText).RefersToRange.Value

    'Return the normalized text value
        Demo_GetNamedText = Trim$(CStr(RawValue))

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
            "Unable to read demo named range '" & NameText & "'. " & Err.Description

End Function

Private Function Demo_ReadNamedBoolean( _
    ByVal Wb As Workbook, _
    ByVal NameText As String) As Boolean

'
'------------------------------------------------------------------------------
'                           READ NAMED BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and parses a Boolean value from a workbook-level named range
'
' WHY THIS EXISTS
'   Demo settings are stored as visible TRUE / FALSE dropdown values, but parsing
'   should be tolerant and explicit
'
' INPUTS
'   Wb
'     Workbook that owns the named range
'
'   NameText
'     Workbook-level name to parse
'
' RETURNS
'   Parsed Boolean value
'
' BEHAVIOR
'   Accepts common true/false text and numeric values
'
' ERROR POLICY
'   Raises a descriptive runtime error if the value cannot be parsed
'
' DEPENDENCIES
'   Demo_GetNamedText
'   Demo_TryParseBoolean
'
' NOTES
'   Accepted true values include TRUE, 1, YES, ON, SI, S , and VERO
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_ReadNamedBoolean"       'Current procedure name

    Dim RawText                As String                                 'Raw named-range text value
    Dim ParsedValue            As Boolean                                'Parsed Boolean value

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Read the named text value
        RawText = Demo_GetNamedText(Wb, NameText)

'------------------------------------------------------------------------------
' PARSE VALUE
'------------------------------------------------------------------------------
    'Parse the named text value as Boolean
        If Not Demo_TryParseBoolean(RawText, ParsedValue) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Named range '" & NameText & "' must contain TRUE or FALSE."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the parsed Boolean value
        Demo_ReadNamedBoolean = ParsedValue

End Function

Private Function Demo_ReadFirstDayOfWeek( _
    ByVal Wb As Workbook, _
    ByVal NameText As String) As Long

'
'------------------------------------------------------------------------------
'                           READ FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Reads and parses the first-day-of-week setting from a named range
'
' WHY THIS EXISTS
'   The demo sheet stores first-day values as worksheet-friendly text
'
' INPUTS
'   Wb
'     Workbook that owns the named range
'
'   NameText
'     Workbook-level name to parse
'
' RETURNS
'   vbSunday or vbMonday
'
' BEHAVIOR
'   Parses supported Sunday and Monday values
'
' ERROR POLICY
'   Raises a descriptive runtime error if the value cannot be parsed
'
' DEPENDENCIES
'   Demo_GetNamedText
'   Demo_TryParseFirstDayOfWeek
'
' NOTES
'   Accepted values include vbSunday, Sunday, Sun, 1, vbMonday, Monday, Mon, 2
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_ReadFirstDayOfWeek"     'Current procedure name

    Dim RawText                As String                                 'Raw named-range text value
    Dim ParsedValue            As Long                                   'Parsed first-day value

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Read the named text value
        RawText = Demo_GetNamedText(Wb, NameText)

'------------------------------------------------------------------------------
' PARSE VALUE
'------------------------------------------------------------------------------
    'Parse the named text value as first day of week
        If Not Demo_TryParseFirstDayOfWeek(RawText, ParsedValue) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Named range '" & NameText & "' must contain vbSunday or vbMonday."
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the parsed first-day value
        Demo_ReadFirstDayOfWeek = ParsedValue

End Function

Private Function Demo_TryParseBoolean( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY PARSE BOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a worksheet-friendly Boolean value
'
' WHY THIS EXISTS
'   Demo settings may be stored as text or numeric values depending on workbook
'   editing behavior and Excel regional settings
'
' INPUTS
'   RawValue
'     Raw text value to parse
'
'   ParsedValue
'     Output parsed Boolean value
'
' RETURNS
'   True when parsing succeeds; otherwise False
'
' BEHAVIOR
'   Accepts common English and Italian true/false values
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   None
'
' NOTES
'   This parser is intentionally independent from the companion module's private
'   settings parser
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedValue        As String                                 'Normalized Boolean text

'------------------------------------------------------------------------------
' NORMALIZE VALUE
'------------------------------------------------------------------------------
    'Normalize the supplied value
        NormalizedValue = UCase$(Trim$(RawValue))

'------------------------------------------------------------------------------
' PARSE VALUE
'------------------------------------------------------------------------------
    'Parse true values
        Select Case NormalizedValue

            Case "1", "-1", "TRUE", "YES", "Y", "ON", "SI", "S ", "VERO"
                ParsedValue = True
                Demo_TryParseBoolean = True
                Exit Function

            Case "0", "FALSE", "NO", "N", "OFF", "FALSO"
                ParsedValue = False
                Demo_TryParseBoolean = True
                Exit Function

        End Select

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return False when parsing fails
        Demo_TryParseBoolean = False

End Function

Private Function Demo_TryParseFirstDayOfWeek( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY PARSE FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a worksheet-friendly first-day-of-week value
'
' WHY THIS EXISTS
'   Demo sheets usually expose first-day choices as readable text rather than raw
'   numeric VBA constants
'
' INPUTS
'   RawValue
'     Raw text value to parse
'
'   ParsedValue
'     Output parsed first-day value
'
' RETURNS
'   True when parsing succeeds; otherwise False
'
' BEHAVIOR
'   Supports Sunday-start and Monday-start values
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   VBA weekday constants
'
' NOTES
'   The DatePicker currently supports only vbSunday and vbMonday
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedValue        As String                                 'Normalized first-day text

'------------------------------------------------------------------------------
' NORMALIZE VALUE
'------------------------------------------------------------------------------
    'Normalize the supplied value
        NormalizedValue = UCase$(Trim$(RawValue))

'------------------------------------------------------------------------------
' PARSE VALUE
'------------------------------------------------------------------------------
    'Parse Sunday values
        Select Case NormalizedValue

            Case "1", "VBSUNDAY", "SUNDAY", "SUN"
                ParsedValue = vbSunday
                Demo_TryParseFirstDayOfWeek = True
                Exit Function

            Case "2", "VBMONDAY", "MONDAY", "MON"
                ParsedValue = vbMonday
                Demo_TryParseFirstDayOfWeek = True
                Exit Function

        End Select

'------------------------------------------------------------------------------
' RETURN FALLBACK
'------------------------------------------------------------------------------
    'Return False when parsing fails
        Demo_TryParseFirstDayOfWeek = False

End Function

Private Function Demo_BoolToText(ByVal Value As Boolean) As String

'
'------------------------------------------------------------------------------
'                           BOOLEAN TO TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts a Boolean value to demo-sheet text
'
' WHY THIS EXISTS
'   The demo sheet uses visible TRUE / FALSE dropdown values
'
' INPUTS
'   Value
'     Boolean value to convert
'
' RETURNS
'   TRUE or FALSE
'
' BEHAVIOR
'   Returns uppercase worksheet-friendly Boolean text
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine writes text deliberately, not Boolean values, to match the
'   validation-list presentation
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return TRUE when the value is True
        If Value Then
            Demo_BoolToText = "TRUE"
            Exit Function
        End If

    'Return FALSE when the value is False
        Demo_BoolToText = "FALSE"

End Function

Private Function Demo_GetQualifiedMacroName( _
    ByVal Wb As Workbook, _
    ByVal MacroName As String) As String

'
'------------------------------------------------------------------------------
'                           GET QUALIFIED MACRO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a workbook-qualified macro name for shape and button callbacks
'
' WHY THIS EXISTS
'   Demo buttons should call macros in ThisWorkbook even when another workbook is
'   active
'
' INPUTS
'   Wb
'     Workbook that owns the callback macro
'
'   MacroName
'     Public macro name to qualify
'
' RETURNS
'   Workbook-qualified macro reference
'
' BEHAVIOR
'   Builds a quoted workbook macro reference suitable for Shape.OnAction
'
' ERROR POLICY
'   Raises a descriptive runtime error if inputs are invalid
'
' DEPENDENCIES
'   None
'
' NOTES
'   Apostrophes in workbook names are escaped for safer callback routing
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "Demo_GetQualifiedMacroName"  'Current procedure name

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject a missing workbook
        If Wb Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Wb cannot be Nothing."
        End If

    'Reject an empty macro name
        If Len(Trim$(MacroName)) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "MacroName cannot be empty."
        End If

'------------------------------------------------------------------------------
' RETURN MACRO NAME
'------------------------------------------------------------------------------
    'Return the workbook-qualified macro name
        Demo_GetQualifiedMacroName = "'" & Replace(Wb.Name, "'", "''") & "'!" & Trim$(MacroName)

End Function

Private Sub Demo_TryRefreshGridIconState()

'
'------------------------------------------------------------------------------
'                           TRY REFRESH GRID ICON STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Best-effort refresh of the DatePicker grid icon after demo settings change
'
' WHY THIS EXISTS
'   Enabling or disabling the in-grid icon from the demo sheet should be visible
'   immediately when the manager can safely handle the current workbook context
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the DatePicker manager exists and asks it to refresh workbook context
'
' ERROR POLICY
'   Suppresses refresh errors because settings have already been saved
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   gDP_Manager.Handle_WorkbookContextChange
'
' NOTES
'   This routine is intentionally best-effort because the manager may be absent,
'   reset, or operating in a workbook context where no icon should be shown
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REFRESH GRID ICON STATE
'------------------------------------------------------------------------------
    'Suppress manager-refresh errors
        On Error Resume Next

    'Ensure the DatePicker manager exists
        M_Picker_EnsureManager

    'Refresh the workbook context through the manager when available
        If Not gDP_Manager Is Nothing Then
            gDP_Manager.Handle_WorkbookContextChange
        End If

    'Restore normal error handling
        On Error GoTo 0

End Sub


