Attribute VB_Name = "M_DATEPICKER"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the shared companion module for the DatePicker project
'
' WHY THIS EXISTS
'   The DatePicker coordinates the modeless runtime UserForm, worksheet write-
'   back, right-click integration, optional in-grid icon, optional WinAPI
'   positioning / borderless styling, optional footer clock, and persisted
'   settings
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - public state and enums
'     - display, behavior, feature, and advanced settings
'     - form show / close and initial-date bridge
'     - date-only calendar selection write-back
'     - direct Today and Now write-back commands
'     - single-cell, multi-cell, and table data-column write-back
'     - Application.OnTime footer clock
'     - Mac-safe WinAPI helpers
'     - right-click menu and in-grid icon integration
'     - optional holiday callback dispatch
'
' ERROR POLICY
'   Public routines raise descriptive runtime errors unless explicitly documented
'   as best-effort or safe-default routines
'
' DEPENDENCIES
'   Excel object model
'   Office CommandBars object model
'   MSForms / VBA.UserForms
'   cDatePickerManager
'   UF_DatePicker
'
' NOTES
'   This module deliberately does not use Option Private Module because Excel UI
'   callbacks such as CommandBars, Shape.OnAction, Application.OnTime, and
'   RibbonX may need public procedures in this module
'
'   Registry and context-menu identifiers retain VBA_DATETIMEPICKER as stable
'   legacy names for backward compatibility
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type POINTAPI
        X As Long                   'Mouse or screen X position in pixels
        Y As Long                   'Mouse or screen Y position in pixels
    End Type
    
    Private Type RECT
        Left As Long                'Rectangle left coordinate in pixels
        Top As Long                 'Rectangle top coordinate in pixels
        Right As Long               'Rectangle right coordinate in pixels
        Bottom As Long              'Rectangle bottom coordinate in pixels
    End Type
    
    Private Type MONITORINFO
        cbSize As Long              'Structure size in bytes
        rcMonitor As RECT           'Full monitor rectangle in pixels
        rcWork As RECT              'Work-area rectangle in pixels
        dwFlags As Long             'Monitor flags
    End Type

'------------------------------------------------------------------------------
' WINDOWS API DECLARATIONS
'------------------------------------------------------------------------------
    #If Mac Then
    #Else
        #If VBA7 Then
            Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
                ByRef lpPoint As POINTAPI) As Long
    
            Private Declare PtrSafe Function MonitorFromRect Lib "user32" ( _
                ByRef lprc As RECT, _
                ByVal dwFlags As Long) As LongPtr
    
            Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
                ByVal hMonitor As LongPtr, _
                ByRef lpmi As MONITORINFO) As Long
    
            Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As LongPtr
    
            Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As LongPtr, _
                ByRef lpRect As RECT) As Long
    
            Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
                ByVal hWnd As LongPtr, _
                ByVal hWndInsertAfter As LongPtr, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal uFlags As Long) As Long
    
            Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
                ByVal hWnd As LongPtr) As Long
    
            #If Win64 Then
                Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr
    
                Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As LongPtr) As LongPtr
            #Else
                Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr
    
                Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As LongPtr) As LongPtr
            #End If
        #Else
            Private Declare Function GetCursorPos Lib "user32" ( _
                ByRef lpPoint As POINTAPI) As Long
    
            Private Declare Function MonitorFromRect Lib "user32" ( _
                ByRef lprc As RECT, _
                ByVal dwFlags As Long) As Long
    
            Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
                ByVal hMonitor As Long, _
                ByRef lpmi As MONITORINFO) As Long
    
            Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long
    
            Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                ByRef lpRect As RECT) As Long
    
            Private Declare Function SetWindowPos Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal uFlags As Long) As Long
    
            Private Declare Function DrawMenuBar Lib "user32" ( _
                ByVal hWnd As Long) As Long
    
            Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long
    
            Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
        #End If
    #End If

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_FORM_NAME                     As String = "UF_DatePicker"          'DatePicker UserForm name

    Private Const DP_SETTINGS_SECTION_DISPLAY      As String = "Display"                 'Display settings section
    Private Const DP_SETTINGS_SECTION_FEATURES     As String = "Features"                'Feature settings section
    Private Const DP_SETTINGS_SECTION_ADVANCED     As String = "Advanced"                'Advanced settings section

    Private Const DP_SETTING_FIRST_DAY_OF_WEEK     As String = "FirstDayOfWeek"          'First-day setting key
    Private Const DP_SETTING_USE_LOCAL_NAMES       As String = "UseLocalNames"           'Local names setting key
    Private Const DP_SETTING_ALLOW_OUTSIDE_MONTH   As String = "AllowOutsideMonth"       'Outside-month setting key
    Private Const DP_SETTING_HIGHLIGHT_WEEKENDS    As String = "HighlightWeekends"       'Highlight weekends setting key
    Private Const DP_SETTING_CLOSE_AFTER_SELECTION As String = "CloseAfterSelection"     'Close-after-selection setting key
    Private Const DP_SETTING_CLOCK_MODE            As String = "ClockMode"               'Clock mode setting key
    Private Const DP_SETTING_SIZE_MODE             As String = "SizeMode"                'Size mode setting key
    Private Const DP_SETTING_SHOW_RIGHT_CLICK      As String = "ShowRightClick"          'Right-click setting key
    Private Const DP_SETTING_SHOW_GRID_ICON        As String = "ShowGridIcon"            'Grid-icon setting key
    Private Const DP_SETTING_USE_WINAPI            As String = "UseWinAPI"               'WinAPI setting key
    Private Const DP_SETTING_HOLIDAY_CALLBACK      As String = "HolidayCallback"         'Holiday callback setting key

    Private Const DP_CONTEXT_MENU_TAG              As String = "VBA_DATETIMEPICKER"      'Legacy context-menu tag
    Private Const DP_CONTEXT_MENU_CAPTION          As String = "Date Picker"             'Context-menu caption
    Private Const DP_CONTEXT_MENU_FACEID           As Long = 1992                        'Context-menu icon FaceId
    Private Const DP_CONTEXT_MENU_BEFORE           As Long = 1                           'Context-menu insertion position

    Private Const DP_TIMER_SECONDS                 As Long = 1                           'Live clock interval in seconds

    Private Const GWL_STYLE                        As Long = -16                         'Window style index
    Private Const MONITOR_DEFAULTTONEAREST         As Long = 2                           'Monitor nearest flag
    Private Const SWP_NOSIZE                       As Long = &H1                         'SetWindowPos no-size flag
    Private Const SWP_NOMOVE                       As Long = &H2                         'SetWindowPos no-move flag
    Private Const SWP_NOZORDER                     As Long = &H4                         'SetWindowPos no-z-order flag
    Private Const SWP_NOACTIVATE                   As Long = &H10                        'SetWindowPos no-activate flag
    Private Const SWP_FRAMECHANGED                 As Long = &H20                        'SetWindowPos frame-changed flag
    Private Const SWP_SHOWWINDOW                   As Long = &H40                        'SetWindowPos show-window flag

    Private Const DP_GRID_ICON_FONT_NAME           As String = "Segoe MDL2 Assets"       'Grid icon font name
    Private Const DP_GRID_ICON_CALENDAR_CODEPOINT  As Long = &HE8BF                      'Segoe MDL2 calendar glyph
    Private Const DP_GRID_ICON_FONT_SIZE           As Single = 10                        'Grid icon font size
    Private Const DP_GRID_ICON_FORE_COLOR          As Long = vbWhite                     'Grid icon foreground color

    #If VBA7 Then
        Private Const WS_CAPTION                   As LongPtr = &HC00000                 'Window caption style flag
    #Else
        Private Const WS_CAPTION                   As Long = &HC00000                    'Window caption style flag
    #End If

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------

    Public Const DP_SETTINGS_APP_NAME              As String = "VBA_DATETIMEPICKER"      'Legacy registry application name
    Public Const DP_GRID_ICON_NAME                 As String = "DP_GridIcon"             'Worksheet grid icon shape name
    Public Const DP_MSGBOX_TITLE                   As String = "Date Picker"             'Message-box title

    Public Const DP_DEFAULT_FIRST_DAY_OF_WEEK      As Long = vbMonday                    'Default first day of week
    Public Const DP_DEFAULT_USE_LOCAL_NAMES        As Boolean = True                     'Default local names flag
    Public Const DP_DEFAULT_ALLOW_OUTSIDE_MONTH    As Boolean = True                     'Default outside-month selection flag
    Public Const DP_DEFAULT_HIGHLIGHT_WEEKENDS     As Boolean = True                     'Default weekend highlight flag
    Public Const DP_DEFAULT_CLOSE_AFTER_SELECTION  As Boolean = True                     'Default close-after-selection flag
    Public Const DP_DEFAULT_SHOW_RIGHT_CLICK       As Boolean = True                     'Default right-click feature flag
    Public Const DP_DEFAULT_SHOW_GRID_ICON         As Boolean = True                     'Default grid-icon feature flag
    Public Const DP_DEFAULT_HOLIDAY_CALLBACK       As String = vbNullString              'Default holiday callback name

'------------------------------------------------------------------------------
' PUBLIC THEME CONSTANTS
'------------------------------------------------------------------------------

    Public Const DP_THEME_COLOR_HEADER             As Long = 9985057                     'Header background
    Public Const DP_THEME_COLOR_HEADER_FORE        As Long = vbWhite                     'Header foreground
    Public Const DP_THEME_COLOR_FORM_BACK          As Long = vbWhite                     'Form background
    Public Const DP_THEME_COLOR_TEXT               As Long = vbButtonText                'Primary text
    Public Const DP_THEME_COLOR_TEXT_MUTED         As Long = &H808080                    'Muted text
    Public Const DP_THEME_COLOR_BORDER_LIGHT       As Long = &HE6E6E6                    'Light border
    Public Const DP_THEME_COLOR_HOVER              As Long = &HF2F2F2                    'Hover background
    Public Const DP_THEME_COLOR_FOOTER             As Long = &HF7F7F7                    'Footer background
    Public Const DP_THEME_COLOR_SELECTED_BACK      As Long = 9985057                     'Selected background
    Public Const DP_THEME_COLOR_SELECTED_FORE      As Long = vbWhite                     'Selected foreground
    Public Const DP_THEME_COLOR_TODAY_BACK         As Long = &HF2F7FC                    'Today background
    Public Const DP_THEME_COLOR_HOLIDAY_FORE       As Long = &H8080FF                    'Holiday foreground

'------------------------------------------------------------------------------
' PUBLIC ENUMS
'------------------------------------------------------------------------------
    Public Enum DP_WriteAction
        Date_Picker = 0
    End Enum
    
    Public Enum DP_ClockMode
        DP_ClockMode_Static = 0
        DP_ClockMode_Live = 1
    End Enum
    
    Public Enum DP_SizeMode
        DP_SizeMode_Normal = 0
        DP_SizeMode_Compact = 1
    End Enum

'------------------------------------------------------------------------------
' PUBLIC STATE
'------------------------------------------------------------------------------
    Public gDP_Manager                  As cDatePickerManager   'DatePicker manager/controller
    Public gDP_WriteValue               As Date                 'Picked value written to Excel
    Public gDP_ShowRightClick           As Boolean              'Show right-click menu entry
    Public gDP_ShowGridIcon             As Boolean              'Show in-grid icon
    Public gDP_GridIconShape            As Shape                'In-grid icon shape reference
    Public gDP_IconPath                 As String               'Optional icon file path

    Public gDP_FirstDayOfWeek           As Long                 'vbSunday or vbMonday
    Public gDP_UseLocalNames            As Boolean              'Use local day/month captions
    Public gDP_ClockMode                As DP_ClockMode         'Static or live clock
    Public gDP_SizeMode                 As DP_SizeMode          'Normal or compact layout
    Public gDP_HighlightWeekends        As Boolean              'Highlight weekend days
    Public gDP_AllowOutsideMonthSel     As Boolean              'Allow outside-month day selection
    Public gDP_CloseAfterSelection      As Boolean              'Close picker after successful write-back
    Public gDP_EnableRightClickMenu     As Boolean
    Public gDP_ShowInGridIcon           As Boolean
    Public gDP_UseWinAPI                As Boolean              'Allow WinAPI features
    Public gDP_HolidayCallbackName      As String               'Optional holiday callback name

    Public gDP_InitialDate              As Date                 'Initial date for next form instance
    Public gDP_HasInitialDate           As Boolean              'True when initial date is available
    Public gDP_SelectedDate             As Date                 'Currently selected date
    Public gDP_HasSelectedDate          As Boolean              'True when selected date is available

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mSettingsLoaded             As Boolean              'Settings loaded flag
    Private mDP_NextTickTime            As Date                 'Next OnTime tick
    Private mDP_TimerIsRunning          As Boolean              'Timer running flag
    Private mDP_TimerProcedureName      As String               'Qualified OnTime timer procedure name

'------------------------------------------------------------------------------
' EMBEDDED GRID ICON
'------------------------------------------------------------------------------
    Private Const DP_GRID_ICON_TEMP_FILE_NAME       As String = "VBA_DatePicker_GridIcon_v1_64.png" 'Embedded icon temp-file name
    Private Const DP_GRID_ICON_TEMP_MIN_BYTES       As Long = 500                                   'Minimum valid embedded icon size
'
'------------------------------------------------------------------------------
'
'                                   SETTINGS
'
'------------------------------------------------------------------------------

Public Sub M_Settings_Load()

'
'------------------------------------------------------------------------------
'                           LOAD DATEPICKER SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Loads persisted DatePicker settings into public project state
'
' WHY THIS EXISTS
'   User preferences should persist across Excel sessions and should be available
'   before the DatePicker form is initialized
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads display, behavior, feature, and advanced DatePicker settings from the
'   current user's registry hive, normalizes invalid values to defaults, and
'   saves the normalized state back to the registry
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or saved
'
' DEPENDENCIES
'   GetSetting
'   M_Settings_Save
'
' NOTES
'   Boolean values are parsed from stable 1 / 0 strings and common textual forms
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_Load"          'Current procedure name

    Dim RawValue        As String       'Raw setting value
    Dim ParsedLong      As Long         'Parsed numeric setting value
    Dim ParsedBoolean   As Boolean      'Parsed Boolean setting value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Read the first-day-of-week setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_FIRST_DAY_OF_WEEK, CStr(DP_DEFAULT_FIRST_DAY_OF_WEEK))

    'Store the parsed first-day setting or its default
        If M_Settings_TryParseFirstDayOfWeek(RawValue, ParsedLong) Then
            gDP_FirstDayOfWeek = ParsedLong
        Else
            gDP_FirstDayOfWeek = DP_DEFAULT_FIRST_DAY_OF_WEEK
        End If

    'Read the local-name setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_USE_LOCAL_NAMES, M_Settings_BooleanToStorageValue(DP_DEFAULT_USE_LOCAL_NAMES))

    'Store the parsed local-name setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_UseLocalNames = ParsedBoolean
        Else
            gDP_UseLocalNames = DP_DEFAULT_USE_LOCAL_NAMES
        End If

    'Read the outside-month setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH, M_Settings_BooleanToStorageValue(DP_DEFAULT_ALLOW_OUTSIDE_MONTH))

    'Store the parsed outside-month setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_AllowOutsideMonthSel = ParsedBoolean
        Else
            gDP_AllowOutsideMonthSel = DP_DEFAULT_ALLOW_OUTSIDE_MONTH
        End If

    'Read the weekend-highlight setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_HIGHLIGHT_WEEKENDS, M_Settings_BooleanToStorageValue(DP_DEFAULT_HIGHLIGHT_WEEKENDS))

    'Store the parsed weekend-highlight setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_HighlightWeekends = ParsedBoolean
        Else
            gDP_HighlightWeekends = DP_DEFAULT_HIGHLIGHT_WEEKENDS
        End If

'------------------------------------------------------------------------------
' LOAD BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Read the close-after-selection setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_CLOSE_AFTER_SELECTION, M_Settings_BooleanToStorageValue(DP_DEFAULT_CLOSE_AFTER_SELECTION))

    'Store the parsed close-after-selection setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_CloseAfterSelection = ParsedBoolean
        Else
            gDP_CloseAfterSelection = DP_DEFAULT_CLOSE_AFTER_SELECTION
        End If

    'Read the clock mode setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_CLOCK_MODE, CStr(CLng(DP_ClockMode_Static)))

    'Store the parsed clock mode or its default
        If M_Settings_TryParseLong(RawValue, ParsedLong) Then
            If M_Settings_IsValidClockMode(ParsedLong) Then
                gDP_ClockMode = ParsedLong
            Else
                gDP_ClockMode = DP_ClockMode_Static
            End If
        Else
            gDP_ClockMode = DP_ClockMode_Static
        End If

    'Read the size mode setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_SIZE_MODE, CStr(CLng(DP_SizeMode_Normal)))

    'Store the parsed size mode or its default
        If M_Settings_TryParseLong(RawValue, ParsedLong) Then
            If M_Settings_IsValidSizeMode(ParsedLong) Then
                gDP_SizeMode = ParsedLong
            Else
                gDP_SizeMode = DP_SizeMode_Normal
            End If
        Else
            gDP_SizeMode = DP_SizeMode_Normal
        End If

'------------------------------------------------------------------------------
' LOAD FEATURE SETTINGS
'------------------------------------------------------------------------------
    'Read the right-click setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_RIGHT_CLICK, M_Settings_BooleanToStorageValue(DP_DEFAULT_SHOW_RIGHT_CLICK))

    'Store the parsed right-click setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_ShowRightClick = ParsedBoolean
        Else
            gDP_ShowRightClick = DP_DEFAULT_SHOW_RIGHT_CLICK
        End If

    'Read the grid-icon setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_GRID_ICON, M_Settings_BooleanToStorageValue(DP_DEFAULT_SHOW_GRID_ICON))

    'Store the parsed grid-icon setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_ShowGridIcon = ParsedBoolean
        Else
            gDP_ShowGridIcon = DP_DEFAULT_SHOW_GRID_ICON
        End If

'------------------------------------------------------------------------------
' LOAD ADVANCED SETTINGS
'------------------------------------------------------------------------------
    'Read the WinAPI setting
        RawValue = GetSetting(DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_USE_WINAPI, M_Settings_BooleanToStorageValue(M_Platform_CanUseWinAPI))

    'Store the parsed WinAPI setting or its platform default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_UseWinAPI = ParsedBoolean
        Else
            gDP_UseWinAPI = M_Platform_CanUseWinAPI
        End If

    'Read the holiday callback setting
        gDP_HolidayCallbackName = Trim$(GetSetting(DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_ADVANCED, DP_SETTING_HOLIDAY_CALLBACK, _
            DP_DEFAULT_HOLIDAY_CALLBACK))

'------------------------------------------------------------------------------
' FINALIZE SETTINGS LOAD
'------------------------------------------------------------------------------
    'Mark settings as loaded before saving normalized values
        mSettingsLoaded = True

    'Save normalized settings back to the registry
        M_Settings_Save

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

Public Sub M_Settings_Save()

'
'------------------------------------------------------------------------------
'                           SAVE DATEPICKER SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Saves current DatePicker settings for the current Windows user
'
' WHY THIS EXISTS
'   User preferences should persist between DatePicker sessions
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates and persists current DatePicker settings
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings are invalid or cannot be saved
'
' DEPENDENCIES
'   SaveSetting
'
' NOTES
'   Boolean values are saved as 1 / 0 for stable parsing
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_Save"     'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' INITIALIZE DEFAULTS IF NEEDED
'------------------------------------------------------------------------------
    'Initialize default settings if settings have not yet been loaded
        If Not mSettingsLoaded Then M_Settings_InitializeDefaults

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "gDP_FirstDayOfWeek must be vbSunday or vbMonday."
        End If

    'Reject unsupported clock modes
        If Not M_Settings_IsValidClockMode(gDP_ClockMode) Then
            Err.Raise vbObjectError + 514, PROC_NAME, "gDP_ClockMode is unsupported."
        End If

    'Reject unsupported size modes
        If Not M_Settings_IsValidSizeMode(gDP_SizeMode) Then
            Err.Raise vbObjectError + 515, PROC_NAME, "gDP_SizeMode is unsupported."
        End If

'------------------------------------------------------------------------------
' SAVE DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Save the first-day-of-week setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_FIRST_DAY_OF_WEEK, CStr(gDP_FirstDayOfWeek)

    'Save the local-name setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_USE_LOCAL_NAMES, M_Settings_BooleanToStorageValue(gDP_UseLocalNames)

    'Save the outside-month setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH, M_Settings_BooleanToStorageValue(gDP_AllowOutsideMonthSel)

    'Save the weekend-highlight setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_HIGHLIGHT_WEEKENDS, M_Settings_BooleanToStorageValue(gDP_HighlightWeekends)

'------------------------------------------------------------------------------
' SAVE BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Save the close-after-selection setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_CLOSE_AFTER_SELECTION, M_Settings_BooleanToStorageValue(gDP_CloseAfterSelection)

    'Save the clock mode setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_CLOCK_MODE, CStr(CLng(gDP_ClockMode))

    'Save the size mode setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_SIZE_MODE, CStr(CLng(gDP_SizeMode))

'------------------------------------------------------------------------------
' SAVE FEATURE SETTINGS
'------------------------------------------------------------------------------
    'Save the right-click setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_RIGHT_CLICK, M_Settings_BooleanToStorageValue(gDP_ShowRightClick)

    'Save the grid-icon setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_GRID_ICON, M_Settings_BooleanToStorageValue(gDP_ShowGridIcon)

'------------------------------------------------------------------------------
' SAVE ADVANCED SETTINGS
'------------------------------------------------------------------------------
    'Save the WinAPI setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_USE_WINAPI, M_Settings_BooleanToStorageValue(gDP_UseWinAPI)

    'Save the holiday callback setting
        SaveSetting DP_SETTINGS_APP_NAME, DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_HOLIDAY_CALLBACK, Trim$(gDP_HolidayCallbackName)

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


Private Sub M_Settings_InitializeDefaults()

'
'------------------------------------------------------------------------------
'                           INITIALIZE DEFAULT SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes DatePicker settings to module defaults
'
' WHY THIS EXISTS
'   M_Settings_Save may be called before settings have been loaded from the
'   registry. Defaults must be applied before validation and save
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns default values to all persisted settings
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'
' NOTES
'   This routine changes in-memory settings only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE DEFAULTS
'------------------------------------------------------------------------------
    'Initialize the first-day setting
        gDP_FirstDayOfWeek = DP_DEFAULT_FIRST_DAY_OF_WEEK

    'Initialize the local-name setting
        gDP_UseLocalNames = DP_DEFAULT_USE_LOCAL_NAMES

    'Initialize the outside-month setting
        gDP_AllowOutsideMonthSel = DP_DEFAULT_ALLOW_OUTSIDE_MONTH

    'Initialize the weekend-highlight setting
        gDP_HighlightWeekends = DP_DEFAULT_HIGHLIGHT_WEEKENDS

    'Initialize the close-after-selection setting
        gDP_CloseAfterSelection = DP_DEFAULT_CLOSE_AFTER_SELECTION

    'Initialize the clock mode
        gDP_ClockMode = DP_ClockMode_Static

    'Initialize the size mode
        gDP_SizeMode = DP_SizeMode_Normal

    'Initialize the right-click feature setting
        gDP_ShowRightClick = DP_DEFAULT_SHOW_RIGHT_CLICK

    'Initialize the grid-icon feature setting
        gDP_ShowGridIcon = DP_DEFAULT_SHOW_GRID_ICON

    'Initialize the WinAPI feature setting
        gDP_UseWinAPI = M_Platform_CanUseWinAPI

    'Initialize the holiday callback name
        gDP_HolidayCallbackName = DP_DEFAULT_HOLIDAY_CALLBACK

'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Mark settings as loaded
        mSettingsLoaded = True

End Sub

Private Sub M_Settings_EnsureLoaded()

'
'------------------------------------------------------------------------------
'                           ENSURE SETTINGS LOADED
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures DatePicker settings have been loaded before they are read or changed
'
' WHY THIS EXISTS
'   Public getters and setters should be safe to call before the caller has
'   explicitly loaded settings
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings only if they have not already been loaded
'
' ERROR POLICY
'   Propagates errors from M_Settings_Load
'
' DEPENDENCIES
'   M_Settings_Load
'
' NOTES
'   This helper avoids overwriting persisted settings with default values
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LOAD SETTINGS IF NEEDED
'------------------------------------------------------------------------------
    'Load settings if they have not been initialized
        If Not mSettingsLoaded Then M_Settings_Load

End Sub

Public Sub M_Settings_SetFirstDayOfWeek(ByVal FirstDayOfWeek As Long)

'
'------------------------------------------------------------------------------
'                           SET FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves the DatePicker first-day-of-week preference
'
' WHY THIS EXISTS
'   Caller code, demo sheets, and settings panels need to choose whether the
'   calendar starts from Sunday or Monday
'
' INPUTS
'   FirstDayOfWeek
'     First day displayed in the calendar
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, updates the first-day setting, saves it, and
'   refreshes the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if FirstDayOfWeek is unsupported
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Supported values are vbSunday and vbMonday
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetFirstDayOfWeek" 'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "FirstDayOfWeek must be vbSunday or vbMonday."
        End If

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Store the requested first-day setting
        gDP_FirstDayOfWeek = FirstDayOfWeek

    'Persist the updated settings
        M_Settings_Save

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Refresh settings-dependent captions if the form is loaded
        M_FormBridge_RefreshSettings

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

Public Sub M_Settings_SetFirstDayOfWeekText(ByVal FirstDayOfWeekText As String)

'
'------------------------------------------------------------------------------
'                           SET FIRST DAY OF WEEK TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the first-day-of-week preference from worksheet-friendly text
'
' WHY THIS EXISTS
'   Settings sheets often store selections as text such as vbMonday and vbSunday
'
' INPUTS
'   FirstDayOfWeekText
'     Text value resolving to vbSunday or vbMonday
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Parses text and delegates to M_Settings_SetFirstDayOfWeek
'
' ERROR POLICY
'   Raises a descriptive runtime error if text cannot be parsed
'
' DEPENDENCIES
'   M_Settings_TryParseFirstDayOfWeek
'   M_Settings_SetFirstDayOfWeek
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
    Const PROC_NAME         As String = "M_Settings_SetFirstDayOfWeekText" 'Current procedure name

    Dim ParsedValue         As Long     'Parsed first-day value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse the supplied first-day text
        If Not M_Settings_TryParseFirstDayOfWeek(FirstDayOfWeekText, ParsedValue) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "FirstDayOfWeekText must resolve to vbSunday or vbMonday."
        End If

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Save the parsed first-day setting
        M_Settings_SetFirstDayOfWeek ParsedValue

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

Public Sub M_Settings_SetUseLocalDayNames(ByVal UseLocalDayNames As Boolean)

'
'------------------------------------------------------------------------------
'                           SET USE LOCAL NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether the DatePicker uses local day and month names
'
' WHY THIS EXISTS
'   The DatePicker supports both locale-dependent captions and fixed English
'   captions for deterministic cross-locale workbooks
'
' INPUTS
'   UseLocalDayNames
'     True to use local captions
'     False to use fixed English captions
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, updates the setting, saves it, and refreshes the
'   loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if saving or refresh fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   The variable name is preserved for compatibility, but it controls both day
'   and month captions
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetUseLocalDayNames" 'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Store the requested local-name setting
        gDP_UseLocalNames = UseLocalDayNames

    'Persist the updated settings
        M_Settings_Save

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Refresh settings-dependent captions if the form is loaded
        M_FormBridge_RefreshSettings

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

Public Sub M_Settings_SetAllowOutsideMonthDays(ByVal AllowOutsideMonthDays As Boolean)

'
'------------------------------------------------------------------------------
'                           SET ALLOW OUTSIDE MONTH DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether outside-month days can be selected
'
' WHY THIS EXISTS
'   Some workflows allow quick adjacent-month selection from leading or trailing
'   calendar cells, while others restrict users to the displayed month
'
' INPUTS
'   AllowOutsideMonthDays
'     True to allow outside-month day selection
'     False to disable outside-month day selection
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, updates the setting, saves it, and refreshes the
'   loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if saving or refresh fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   The form should call M_DatePolicy_CanSelectDate before accepting a day click
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetAllowOutsideMonthDays" 'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Store the outside-month setting
        gDP_AllowOutsideMonthSel = AllowOutsideMonthDays

    'Persist the updated settings
        M_Settings_Save

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Refresh settings-dependent captions if the form is loaded
        M_FormBridge_RefreshSettings

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

Public Sub M_Settings_SetHighlightWeekends(ByVal HighlightWeekends As Boolean)

'
'------------------------------------------------------------------------------
'                           SET HIGHLIGHT WEEKENDS
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether weekend days are visually highlighted in the
'   DatePicker calendar grid
'
' WHY THIS EXISTS
'   Some users prefer weekends to be visually differentiated, while others prefer
'   a neutral calendar grid
'
' INPUTS
'   HighlightWeekends
'     True to highlight weekend days
'     False to display weekends with normal day-label font weight
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, updates the weekend-highlight setting, saves it,
'   and refreshes the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if saving or refresh fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   This setting controls weekend visual highlighting only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetHighlightWeekends" 'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Store the weekend-highlight setting
        gDP_HighlightWeekends = HighlightWeekends

    'Persist the updated settings
        M_Settings_Save

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Refresh settings-dependent captions and grid if the form is loaded
        M_FormBridge_RefreshSettings

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

Public Sub M_Settings_SetCloseAfterSelection(ByVal CloseAfterSelection As Boolean)

'
'------------------------------------------------------------------------------
'                           SET CLOSE AFTER SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether the DatePicker closes after a successful write-back
'
' WHY THIS EXISTS
'   Some workflows prefer the picker to close immediately after a date is
'   selected, while others keep the modeless picker open for repeated use
'
' INPUTS
'   CloseAfterSelection
'     True to close the picker after successful write-back
'     False to keep the picker open after successful write-back
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, updates the close-after-selection setting, and
'   saves it
'
' ERROR POLICY
'   Raises a descriptive runtime error if saving fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   This setting controls form lifecycle after successful write-back only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetCloseAfterSelection" 'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Store the close-after-selection setting
        gDP_CloseAfterSelection = CloseAfterSelection

    'Persist the updated settings
        M_Settings_Save

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

Public Sub M_Settings_SetClockMode(ByVal ClockMode As DP_ClockMode)

'
'------------------------------------------------------------------------------
'                           SETTINGS SET CLOCKMODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and persists the DatePicker settings clockmode setting
'
' WHY THIS EXISTS
'   DatePicker settings must be updateable through controlled public entry points so demo sheets, Ribbon callbacks, and host workbooks do not mutate public state inconsistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded
'   - Validates the supplied value when applicable
'   - Updates the related in-memory setting
'   - Persists the updated settings
'   - Refreshes open DatePicker UI when applicable
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_Settings_SetClockMode"          'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Reject unsupported clock modes
        If Not M_Settings_IsValidClockMode(ClockMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ClockMode is unsupported."
        End If

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

    'Store the clock mode
        gDP_ClockMode = ClockMode

    'Persist the updated settings
        M_Settings_Save

    'Start or stop the live clock timer according to the mode
        M_Timer_ApplyClockMode

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Settings_SetSizeMode(ByVal SizeMode As DP_SizeMode)

'
'------------------------------------------------------------------------------
'                           SETTINGS SET SIZEMODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and persists the DatePicker settings sizemode setting
'
' WHY THIS EXISTS
'   DatePicker settings must be updateable through controlled public entry points so demo sheets, Ribbon callbacks, and host workbooks do not mutate public state inconsistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded
'   - Validates the supplied value when applicable
'   - Updates the related in-memory setting
'   - Persists the updated settings
'   - Refreshes open DatePicker UI when applicable
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_Settings_SetSizeMode"           'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Reject unsupported size modes
        If Not M_Settings_IsValidSizeMode(SizeMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "SizeMode is unsupported."
        End If

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

    'Store the size mode
        gDP_SizeMode = SizeMode

    'Persist the updated settings
        M_Settings_Save

    'Refresh settings-dependent captions if the form is loaded
        M_FormBridge_RefreshSettings

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Settings_SetHolidayCallback(ByVal HolidayCallbackName As String)

'
'------------------------------------------------------------------------------
'                           SETTINGS SET HOLIDAYCALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and persists the DatePicker settings holidaycallback setting
'
' WHY THIS EXISTS
'   DatePicker settings must be updateable through controlled public entry points so demo sheets, Ribbon callbacks, and host workbooks do not mutate public state inconsistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded
'   - Validates the supplied value when applicable
'   - Updates the related in-memory setting
'   - Persists the updated settings
'   - Refreshes open DatePicker UI when applicable
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_Settings_SetHolidayCallback"    'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

    'Store the callback name
        gDP_HolidayCallbackName = Trim$(HolidayCallbackName)

    'Persist the updated settings
        M_Settings_Save

    'Refresh settings-dependent captions if the form is loaded
        M_FormBridge_RefreshSettings

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Settings_SetShowRightClick(ByVal ShowRightClick As Boolean)

'
'------------------------------------------------------------------------------
'                           SETTINGS SET SHOWRIGHTCLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and persists the DatePicker settings showrightclick setting
'
' WHY THIS EXISTS
'   DatePicker settings must be updateable through controlled public entry points so demo sheets, Ribbon callbacks, and host workbooks do not mutate public state inconsistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded
'   - Validates the supplied value when applicable
'   - Updates the related in-memory setting
'   - Persists the updated settings
'   - Refreshes open DatePicker UI when applicable
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_Settings_SetShowRightClick"     'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

    'Store the feature setting
        gDP_ShowRightClick = ShowRightClick

    'Persist the updated settings
        M_Settings_Save

    'Synchronize right-click menus
        M_ContextMenu_Update

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Settings_SetShowGridIcon(ByVal ShowGridIcon As Boolean)

'
'------------------------------------------------------------------------------
'                           SETTINGS SET SHOWGRIDICON
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and persists the DatePicker settings showgridicon setting
'
' WHY THIS EXISTS
'   DatePicker settings must be updateable through controlled public entry points so demo sheets, Ribbon callbacks, and host workbooks do not mutate public state inconsistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded
'   - Validates the supplied value when applicable
'   - Updates the related in-memory setting
'   - Persists the updated settings
'   - Refreshes open DatePicker UI when applicable
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "M_Settings_SetShowGridIcon"       'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

    'Store the feature setting
        gDP_ShowGridIcon = ShowGridIcon

    'Persist the updated settings
        M_Settings_Save

    'Remove any stale icon when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Remove
        End If

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Function M_Settings_GetFirstDayOfWeek() As Long

'
'------------------------------------------------------------------------------
'                  SETTINGS GET FIRST DAY OF THE WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings get firstdayofweek value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Load settings if needed
        M_Settings_EnsureLoaded

    'Return the current first-day setting
        M_Settings_GetFirstDayOfWeek = gDP_FirstDayOfWeek
End Function

Public Function M_Settings_GetUseLocalDayNames() As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS GET USELOCALDAYNAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings get uselocaldaynames value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Load settings if needed
        M_Settings_EnsureLoaded

    'Return the current local-name setting
        M_Settings_GetUseLocalDayNames = gDP_UseLocalNames
End Function

Public Function M_Settings_GetFirstDayOfWeekText() As String

'
'------------------------------------------------------------------------------
'                           SETTINGS GET FIRSTDAYOFWEEKTEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings get firstdayofweektext value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Load settings if needed
        M_Settings_EnsureLoaded

    'Return the current first-day setting as text
        M_Settings_GetFirstDayOfWeekText = M_Settings_FirstDayOfWeekToText(gDP_FirstDayOfWeek)
End Function

Public Function M_Settings_FirstDayOfWeekToText(ByVal FirstDayOfWeek As Long) As String

'
'------------------------------------------------------------------------------
'                           SETTINGS FIRSTDAYOFWEEKTOTEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings firstdayofweektotext value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, "M_Settings_FirstDayOfWeekToText", _
                "FirstDayOfWeek must be vbSunday or vbMonday."
        End If

    'Return the text representation
        Select Case FirstDayOfWeek
            Case vbSunday
                M_Settings_FirstDayOfWeekToText = "vbSunday"
            Case vbMonday
                M_Settings_FirstDayOfWeekToText = "vbMonday"
        End Select
End Function

Public Function M_Settings_IsValidFirstDayOfWeek(ByVal FirstDayOfWeek As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS ISVALIDFIRSTDAYOFWEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings isvalidfirstdayofweek value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Return whether the first-day value is supported
        M_Settings_IsValidFirstDayOfWeek = _
            (FirstDayOfWeek = vbSunday Or FirstDayOfWeek = vbMonday)
End Function

Private Function M_Settings_IsValidClockMode(ByVal ClockMode As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS ISVALIDCLOCKMODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings isvalidclockmode value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Return whether the clock mode is supported
        M_Settings_IsValidClockMode = _
            (ClockMode = DP_ClockMode_Static Or ClockMode = DP_ClockMode_Live)
End Function

Private Function M_Settings_IsValidSizeMode(ByVal SizeMode As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS ISVALIDSIZEMODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the DatePicker settings isvalidsizemode value
'
' WHY THIS EXISTS
'   A controlled accessor keeps external callers away from directly relying on mutable module state where possible
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Ensures settings are loaded when needed
'   - Returns the requested setting or validation result
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Return whether the size mode is supported
        M_Settings_IsValidSizeMode = _
            (SizeMode = DP_SizeMode_Normal Or SizeMode = DP_SizeMode_Compact)
End Function

Private Function M_Settings_TryParseBoolean( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS TRYPARSEBOOLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be parsed defensively before being trusted
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Normalizes the raw setting value
'   - Returns True when parsing succeeds
'   - Returns False when parsing fails
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   None
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim NormalizedValue As String                                     'Normalized setting value

    'Normalize the raw setting value
        NormalizedValue = UCase$(Trim$(RawValue))

    'Parse supported Boolean values
        Select Case NormalizedValue
            Case "1", "-1", "TRUE", "YES", "Y", "ON", "SI", "S", "VERO"
                ParsedValue = True
                M_Settings_TryParseBoolean = True
            Case "0", "FALSE", "NO", "N", "OFF", "FALSO"
                ParsedValue = False
                M_Settings_TryParseBoolean = True
            Case Else
                M_Settings_TryParseBoolean = False
        End Select
End Function

Private Function M_Settings_TryParseFirstDayOfWeek( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS TRYPARSEFIRSTDAYOFWEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be parsed defensively before being trusted
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Normalizes the raw setting value
'   - Returns True when parsing succeeds
'   - Returns False when parsing fails
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   None
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim NormalizedValue As String                                     'Normalized first-day setting

    'Normalize the raw first-day value
        NormalizedValue = UCase$(Trim$(RawValue))

    'Parse supported first-day values
        Select Case NormalizedValue
            Case "1", "VBSUNDAY", "SUNDAY", "SUN"
                ParsedValue = vbSunday
                M_Settings_TryParseFirstDayOfWeek = True
            Case "2", "VBMONDAY", "MONDAY", "MON"
                ParsedValue = vbMonday
                M_Settings_TryParseFirstDayOfWeek = True
            Case Else
                M_Settings_TryParseFirstDayOfWeek = False
        End Select
End Function

Private Function M_Settings_TryParseLong( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           SETTINGS TRYPARSELONG
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be parsed defensively before being trusted
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Normalizes the raw setting value
'   - Returns True when parsing succeeds
'   - Returns False when parsing fails
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   None
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim NormalizedValue As String                                     'Normalized numeric value

    'Set safe default
        M_Settings_TryParseLong = False

    'Suppress parse failures
        On Error GoTo ParseFail

    'Normalize the raw numeric value
        NormalizedValue = Trim$(RawValue)

    'Exit when the value is blank
        If Len(NormalizedValue) = 0 Then Exit Function

    'Exit when the value is not numeric
        If Not IsNumeric(NormalizedValue) Then Exit Function

    'Reject decimal-looking values for integer settings
        If InStr(1, NormalizedValue, ".", vbBinaryCompare) > 0 Then Exit Function
        If InStr(1, NormalizedValue, ",", vbBinaryCompare) > 0 Then Exit Function

    'Parse the value as Long
        ParsedValue = CLng(NormalizedValue)

    'Return success
        M_Settings_TryParseLong = True

    'Exit after successful parse
        Exit Function

ParseFail:
    'Return safe default
        M_Settings_TryParseLong = False
End Function

Private Function M_Settings_BooleanToStorageValue(ByVal Value As Boolean) As String
    'Return the settings representation
        If Value Then
            M_Settings_BooleanToStorageValue = "1"
        Else
            M_Settings_BooleanToStorageValue = "0"
        End If
End Function


'
'------------------------------------------------------------------------------
'
'                                  FORM BRIDGE
'
'------------------------------------------------------------------------------

Public Sub M_Picker_EnsureManager()

'
'------------------------------------------------------------------------------
'                          ENSURE DATEPICKER MANAGER
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures the global DatePicker manager object is instantiated
'
' WHY THIS EXISTS
'   The manager hooks Excel application events and coordinates context-sensitive
'   DatePicker UI behavior
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Creates gDP_Manager only when it is currently Nothing
'
' ERROR POLICY
'   Propagates manager instantiation errors
'
' DEPENDENCIES
'   cDatePickerManager
'
' NOTES
'   This is a lazy loader for the manager/controller object
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ENSURE MANAGER
'------------------------------------------------------------------------------
    'Instantiate the manager when missing
        If gDP_Manager Is Nothing Then
            Set gDP_Manager = New cDatePickerManager
        End If

End Sub

Public Sub DP_Show()

'
'------------------------------------------------------------------------------
'                           SHOW DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the DatePicker UserForm modelessly using ActiveCell as the initial
'   date context
'
' WHY THIS EXISTS
'   The DatePicker should open on the date already present in ActiveCell when
'   ActiveCell contains a valid date. Otherwise, it should open on today's date
'
'   When the form is shown from the in-grid icon, the UserForm should also be
'   positioned close to the current mouse position
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - Loads persisted DatePicker settings
'   - Resolves the initial display date from ActiveCell
'   - Stores the initial-date bridge state consumed by UF_DatePicker
'   - Stores selected-date state only when ActiveCell contains a valid date
'   - Unloads any existing DatePicker instance so UserForm_Initialize runs again
'   - Shows UF_DatePicker modelessly
'   - Resolves the loaded form instance
'   - Moves the loaded form close to the current mouse position
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker cannot be shown or
'   positioned
'
' DEPENDENCIES
'   M_Settings_Load
'   M_FormBridge_UnloadLoadedPicker
'   M_FormBridge_GetLoadedForm
'   UF_DatePicker
'   M_Window_MoveFormToMouse
'
' NOTES
'   The form is positioned after Show because M_Window_MoveFormToMouse relies on the
'   native UserForm window handle, which is more reliable after the form has been
'   created
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "DP_Show"                    'Current procedure name

    Dim CellValue              As Variant                               'ActiveCell value snapshot
    Dim InitialDate            As Date                                  'Resolved initial display date
    Dim HasCellDate            As Boolean                               'True when ActiveCell contains a date
    Dim LoadedForm             As Object                                'Loaded DatePicker form instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Load DatePicker settings before opening the form
        M_Settings_Load

    'Default the initial date to today
        InitialDate = VBA.Date

    'Default selected-date availability to False
        HasCellDate = False

'------------------------------------------------------------------------------
' RESOLVE INITIAL DATE FROM ACTIVE CELL
'------------------------------------------------------------------------------
    'Suppress ActiveCell access errors
        On Error Resume Next

    'Read ActiveCell value safely
        CellValue = Application.ActiveCell.Value

    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Use ActiveCell date only when the value is not an Excel error
        If Not IsError(CellValue) Then
            If VBA.IsDate(CellValue) Then
                InitialDate = VBA.DateValue(CDate(CellValue))
                HasCellDate = True
            End If
        End If

'------------------------------------------------------------------------------
' STORE FORM BRIDGE STATE
'------------------------------------------------------------------------------
    'Store the initial date for the next UF_DatePicker instance
        gDP_InitialDate = InitialDate

    'Mark the initial date as available
        gDP_HasInitialDate = True

    'Store selected-date state when ActiveCell contains a valid date
        If HasCellDate Then
            gDP_SelectedDate = InitialDate
            gDP_HasSelectedDate = True
        Else
            gDP_SelectedDate = 0
            gDP_HasSelectedDate = False
        End If

'------------------------------------------------------------------------------
' RESET EXISTING FORM INSTANCE
'------------------------------------------------------------------------------
    'Unload any existing DatePicker instance so Initialize runs again
        M_FormBridge_UnloadLoadedPicker

'------------------------------------------------------------------------------
' SHOW FORM
'------------------------------------------------------------------------------
    'Show the DatePicker form modelessly
        UF_DatePicker.Show vbModeless

'------------------------------------------------------------------------------
' POSITION LOADED FORM
'------------------------------------------------------------------------------
    'Resolve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

    'Move the loaded form top-left corner to the current mouse position
        If Not LoadedForm Is Nothing Then
            M_Window_MoveFormToMouse LoadedForm, 0, 0, False
        End If

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

Public Sub DP_Click()

'
'------------------------------------------------------------------------------
'                           DATEPICKER CLICK ENTRY POINT
'------------------------------------------------------------------------------
' PURPOSE
'   Public callback entry point for right-click menu and grid-icon actions
'
' WHY THIS EXISTS
'   Excel UI elements such as CommandBars and Shapes need a public macro entry
'   point to launch the DatePicker
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to DP_Show
'
' ERROR POLICY
'   Propagates errors from DP_Show
'
' DEPENDENCIES
'   DP_Show
'
' NOTES
'   This routine intentionally remains public
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' SHOW DATEPICKER
'------------------------------------------------------------------------------
    'Show the DatePicker form
        DP_Show

End Sub

Public Sub DP_Close()

'
'------------------------------------------------------------------------------
'                           CLOSE DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Closes the DatePicker form when it is loaded and stops transient UI activity
'
' WHY THIS EXISTS
'   The DatePicker should tear down timers and modeless form state in a
'   controlled way when dismissed or refreshed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stops the live clock timer and unloads UF_DatePicker only when it is loaded
'
' ERROR POLICY
'   Best-effort cleanup. Does not intentionally raise outward
'
' DEPENDENCIES
'   M_Timer_Stop
'   M_FormBridge_GetLoadedForm
'
' NOTES
'   This routine avoids direct default-instance references until a loaded form is
'   explicitly found
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm             As Object                                'Loaded DatePicker form instance

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

    'Stop any active live clock timer
        M_Timer_Stop

    'Retrieve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

    'Unload the loaded form if present
        If Not LoadedForm Is Nothing Then
            Unload LoadedForm
        End If

    'Clear initial-date bridge state
        gDP_HasInitialDate = False

    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Function M_FormBridge_ConsumeInitialDate(ByRef InitialDate As Date) As Boolean

'
'------------------------------------------------------------------------------
'                         CONSUME FORM INITIAL DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Supplies the initial date to UF_DatePicker during initialization
'
' WHY THIS EXISTS
'   DP_Show resolves the initial date from ActiveCell before the form is shown.
'   UF_DatePicker then consumes that value during UserForm_Initialize
'
' INPUTS
'   InitialDate
'     Output date consumed by UF_DatePicker
'
' RETURNS
'   True if an initial date was available; otherwise False
'
' BEHAVIOR
'   Returns the pending initial date once and clears the availability flag
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   gDP_InitialDate
'   gDP_HasInitialDate
'
' NOTES
'   This routine is intentionally consumed once per form initialization
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN INITIAL DATE
'------------------------------------------------------------------------------
    'Return the pending initial date if available
        If gDP_HasInitialDate Then
            InitialDate = gDP_InitialDate
            M_FormBridge_ConsumeInitialDate = True
        Else
            InitialDate = VBA.Date
            M_FormBridge_ConsumeInitialDate = False
        End If

'------------------------------------------------------------------------------
' CLEAR INITIAL DATE FLAG
'------------------------------------------------------------------------------
    'Clear the bridge flag after consumption
        gDP_HasInitialDate = False

End Function

'
'------------------------------------------------------------------------------
'
'                                  WRITE-BACK
'
'------------------------------------------------------------------------------

Public Sub M_Picker_SelectDate( _
    ByVal SelectedDate As Date, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                           SELECT DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Stores the selected DatePicker date, writes it to the Excel target, and then
'   closes or refreshes the picker according to the configured lifecycle setting
'
' WHY THIS EXISTS
'   Day-label click handling in UF_DatePicker should delegate write-back to the
'   companion module so single-cell, multi-cell, and table-column behavior stays
'   centralized
'
' INPUTS
'   SelectedDate
'     Date selected by the user
'
'   NoTableGrow
'     True to prevent single table-cell selections from expanding to the full
'     table column
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes a date-only value to the current Excel target and then:
'     - closes the form when gDP_CloseAfterSelection is True
'     - refreshes the open form when gDP_CloseAfterSelection is False
'
' ERROR POLICY
'   Raises a descriptive runtime error if write-back fails
'
' DEPENDENCIES
'   M_WriteBack_Apply
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   Calendar day selection is intentionally date-only
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Picker_SelectDate"             'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' STORE SELECTED DATE
'------------------------------------------------------------------------------
    'Store the selected date without time
        gDP_SelectedDate = VBA.DateValue(SelectedDate)

    'Mark the selected date as available
        gDP_HasSelectedDate = True

'------------------------------------------------------------------------------
' BUILD WRITE VALUE
'------------------------------------------------------------------------------
    'Store the date-only value to write
        gDP_WriteValue = VBA.DateValue(SelectedDate)

'------------------------------------------------------------------------------
' WRITE TO EXCEL
'------------------------------------------------------------------------------
    'Write the selected value to the current Excel target
        M_WriteBack_Apply Date_Picker, NoTableGrow

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Close the DatePicker after successful selection when configured
        If gDP_CloseAfterSelection Then
            DP_Close
        Else
            M_FormBridge_AfterSuccessfulSelection SelectedDate
        End If

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

Public Sub DP_Today()

'
'------------------------------------------------------------------------------
'                           WRITE TODAY
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current system date to the current Excel selection
'
' WHY THIS EXISTS
'   Footer, Ribbon, menu, or macro callers may need a direct Today command
'   without selecting a calendar day
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes VBA.Date according to normal DatePicker write-back rules, then closes
'   or refreshes the form according to gDP_CloseAfterSelection
'
' ERROR POLICY
'   Propagates write-back errors from M_WriteBack_Apply
'
' DEPENDENCIES
'   M_WriteBack_Apply
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   This command always writes a date-only value
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE DATE
'------------------------------------------------------------------------------
    'Store the current system date
        gDP_WriteValue = VBA.Date

'------------------------------------------------------------------------------
' WRITE TO EXCEL
'------------------------------------------------------------------------------
    'Apply the date to the current selection
        M_WriteBack_Apply Date_Picker, False

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Close or refresh the DatePicker according to the configured lifecycle mode
        If gDP_CloseAfterSelection Then
            DP_Close
        Else
            M_FormBridge_AfterSuccessfulSelection VBA.Date
        End If

End Sub

Public Sub DP_Now()

'
'------------------------------------------------------------------------------
'                           WRITE NOW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current system date-time to the current Excel selection
'
' WHY THIS EXISTS
'   Footer, Ribbon, menu, or macro callers may need a direct Now command without
'   selecting a calendar day
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes VBA.Now according to normal DatePicker write-back rules, then closes
'   or refreshes the form according to gDP_CloseAfterSelection
'
' ERROR POLICY
'   Propagates write-back errors from M_WriteBack_Apply
'
' DEPENDENCIES
'   M_WriteBack_Apply
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   This command always writes a date-time value
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE DATE-TIME
'------------------------------------------------------------------------------
    'Store the current system date-time
        gDP_WriteValue = VBA.Now

'------------------------------------------------------------------------------
' WRITE TO EXCEL
'------------------------------------------------------------------------------
    'Apply the date-time to the current selection
        M_WriteBack_Apply Date_Picker, False

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Close or refresh the DatePicker according to the configured lifecycle mode
        If gDP_CloseAfterSelection Then
            DP_Close
        Else
            M_FormBridge_AfterSuccessfulSelection VBA.Date
        End If

End Sub

Public Sub M_WriteBack_Apply( _
    ByVal iType As DP_WriteAction, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                           DO ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a DatePicker action to the current Excel selection
'
' WHY THIS EXISTS
'   UI click handlers and entry points need a shared routine that suppresses
'   worksheet events during write-back and restores them deterministically
'
' INPUTS
'   iType
'     Action to apply
'
'   NoTableGrow
'     True to keep a single table-cell target as a single cell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Temporarily disables Excel events, dispatches target shaping and write-back,
'   then restores the previous event state
'
' ERROR POLICY
'   Restores Application.EnableEvents and re-raises the original error
'
' DEPENDENCIES
'   M_WriteBack_ResolveAndApplyTarget
'
' NOTES
'   This routine does not change calculation or screen updating
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_WriteBack_Apply"                   'Current procedure name

    Dim PreviousEvents         As Boolean                               'Prior Application.EnableEvents state
    Dim SavedErrNumber         As Long                                  'Captured error number
    Dim SavedErrSource         As String                                'Captured error source
    Dim SavedErrDescription    As String                                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable structured cleanup on failure
        On Error GoTo CleanFail

    'Capture the current Excel events state
        PreviousEvents = Application.EnableEvents

'------------------------------------------------------------------------------
' SUPPRESS EVENTS
'------------------------------------------------------------------------------
    'Disable events during write-back
        Application.EnableEvents = False

'------------------------------------------------------------------------------
' DISPATCH ACTION
'------------------------------------------------------------------------------
    'Apply the requested action to the current selection
        M_WriteBack_ResolveAndApplyTarget iType, NoTableGrow

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Protect cleanup from masking the original error
        On Error Resume Next

    'Restore the previous Excel events state
        Application.EnableEvents = PreviousEvents

    'Restore normal error handling
        On Error GoTo 0

    'Re-raise the original error when needed
        If SavedErrNumber <> 0 Then
            Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If

    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
CleanFail:
    'Capture the original error number
        SavedErrNumber = Err.Number

    'Capture the original error source
        If Len(Err.Source) > 0 Then
            SavedErrSource = Err.Source
        Else
            SavedErrSource = PROC_NAME
        End If

    'Capture the original error description
        SavedErrDescription = Err.Description

    'Resume through cleanup
        Resume CleanExit

End Sub

Private Sub M_WriteBack_ResolveAndApplyTarget( _
    ByVal iType As DP_WriteAction, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                           M_WriteBack_ResolveAndApplyTarget
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker M_WriteBack_ResolveAndApplyTarget routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim Target                 As Range                                 'Resolved target range
    Dim Block                  As Range                                 'Single target area
    Dim TargetTable            As ListObject                            'Worksheet table being inspected
    Dim ColumnIndex            As Long                                  'Resolved table column index

    'Proceed only when the current selection is a Range
        If TypeName(Selection) <> "Range" Then
            Exit Sub
        End If

    'Store the current selection as initial target
        Set Target = Selection

    'Consider table expansion only for one selected cell when allowed
        If Target.Cells.CountLarge = 1 And Not NoTableGrow Then

            'Loop through tables on the target worksheet
                For Each TargetTable In Target.Worksheet.ListObjects

                    'Continue only when the table has a data body
                        If Not TargetTable.DataBodyRange Is Nothing Then

                            'Expand only when the selected cell is inside the data body
                                If Not Intersect(Target, TargetTable.DataBodyRange) Is Nothing Then

                                    'Resolve the table column index
                                        ColumnIndex = Target.Column - TargetTable.DataBodyRange.Column + 1

                                    'Expand to the table data column when valid
                                        If ColumnIndex >= 1 And ColumnIndex <= TargetTable.ListColumns.Count Then
                                            Set Target = TargetTable.ListColumns(ColumnIndex).DataBodyRange
                                        End If

                                    'Stop after resolving the owning table
                                        Exit For

                                End If

                        End If

                Next TargetTable

        End If

    'Loop through each discontiguous area
        For Each Block In Target.Areas

            'Populate this target area
                M_WriteBack_PopulateRange Block, iType

        Next Block
End Sub

Public Sub M_WriteBack_PopulateRange( _
    ByVal oRange As Range, _
    ByVal iType As DP_WriteAction)

'
'------------------------------------------------------------------------------
'                       M_WriteBack_PopulateRange
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker M_WriteBack_PopulateRange routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep
'   UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim Cell                   As Range                                 'Current target cell
    Dim LockedCount            As Long                                  'Protected locked cells skipped
    Dim FailedCount            As Long                                  'Other write failures suppressed
    Dim WriteValue             As Variant                               'Resolved write value

    'Exit when no range is supplied
        If oRange Is Nothing Then
            Exit Sub
        End If

    'Resolve the value to write
        Select Case iType
            Case DP_WriteAction.Date_Picker
                WriteValue = M_WriteBack_GetPickedValue
            Case Else
                Exit Sub
        End Select

    'Loop through each target cell
        For Each Cell In oRange.Cells

            'Attempt to write the value
                Call M_WriteBack_TryWriteCell(Cell, WriteValue, LockedCount, FailedCount)

        Next Cell

    'Show one summary message for protected locked cells
        If LockedCount > 0 Then
            MsgBox CStr(LockedCount) & " protected locked cell(s) were skipped.", _
                vbInformation Or vbOKOnly, DP_MSGBOX_TITLE
        End If

    'Write non-lock failures to the Immediate Window
        If FailedCount > 0 Then
            Debug.Print "M_DATEPICKER.M_WriteBack_PopulateRange: " & _
                CStr(FailedCount) & " cell write failure(s) were suppressed."
        End If
End Sub

Private Function M_WriteBack_TryWriteCell( _
    ByVal TargetCell As Range, _
    ByVal WriteValue As Variant, _
    ByRef LockedCount As Long, _
    ByRef FailedCount As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           TRYWRITECELLVALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker trywritecellvalue routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Set safe default result
        M_WriteBack_TryWriteCell = False

    'Enable per-cell fail-safe handling
        On Error GoTo WriteFail

    'Exit when no target cell is supplied
        If TargetCell Is Nothing Then Exit Function

    'Skip locked cells on protected sheets
        If TargetCell.Locked Then
            If TargetCell.Worksheet.ProtectContents Then
                LockedCount = LockedCount + 1
                Exit Function
            End If
        End If

    'Write the value to the target cell
        TargetCell.Value = WriteValue

    'Return success
        M_WriteBack_TryWriteCell = True

    'Exit after successful write
        Exit Function

WriteFail:
    'Count and suppress the write failure
        FailedCount = FailedCount + 1

    'Write diagnostics to the Immediate Window
        Debug.Print "M_DATEPICKER.M_WriteBack_TryWriteCell: suppressed error " & _
            Err.Number & " - " & Err.Description
End Function

Private Function M_WriteBack_GetPickedValue() As Date

'
'------------------------------------------------------------------------------
'                           M_WriteBack_GetPickedValue
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker M_WriteBack_GetPickedValue routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Return the current picked value
        M_WriteBack_GetPickedValue = gDP_WriteValue
End Function

Public Function M_DatePolicy_CanSelectDate( _
    ByVal CandidateDate As Date, _
    ByVal DisplayYear As Long, _
    ByVal DisplayMonth As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           CANSELECTDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker canselectdate routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Reject invalid display years
        If DisplayYear < 100 Or DisplayYear > 9999 Then
            M_DatePolicy_CanSelectDate = False
            Exit Function
        End If

    'Reject invalid display months
        If DisplayMonth < 1 Or DisplayMonth > 12 Then
            M_DatePolicy_CanSelectDate = False
            Exit Function
        End If

    'Allow all dates when outside-month days are enabled
        If gDP_AllowOutsideMonthSel Then
            M_DatePolicy_CanSelectDate = True
            Exit Function
        End If

    'Allow only dates belonging to the displayed month and year
        M_DatePolicy_CanSelectDate = _
            (VBA.Year(CandidateDate) = DisplayYear And VBA.Month(CandidateDate) = DisplayMonth)
End Function

Public Function M_HolidayPolicy_IsHolidayDate(ByVal CandidateDate As Date) As Boolean

'
'------------------------------------------------------------------------------
'                           ISHOLIDAYDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker isholidaydate routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim CallbackResult As Variant                                     'Callback return value

    'Set safe default
        M_HolidayPolicy_IsHolidayDate = False

    'Exit if no callback is configured
        If Len(Trim$(gDP_HolidayCallbackName)) = 0 Then Exit Function

    'Suppress callback failures
        On Error Resume Next

    'Run the custom holiday callback
        CallbackResult = Application.Run(gDP_HolidayCallbackName, CandidateDate)

    'Return callback result when it is Boolean
        If VarType(CallbackResult) = vbBoolean Then
            M_HolidayPolicy_IsHolidayDate = CBool(CallbackResult)
        End If

    'Restore normal error handling
        On Error GoTo 0
End Function


'
'------------------------------------------------------------------------------
'
'                                  TEXT HELPERS
'
'------------------------------------------------------------------------------

Public Function M_Caption_GetMonth( _
    ByVal MonthNumber As Long, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET MONTHCAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a DatePicker caption value
'
' WHY THIS EXISTS
'   Caption generation should remain centralized so local-name and fixed-English display modes stay consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates inputs
'   - Builds the requested caption
'   - Returns the normalized display text
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA date formatting functions
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_Caption_GetMonth"                 'Current procedure name
    Dim MonthCaption As String                                         'Resolved month caption

    'Reject invalid month numbers
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

    'Resolve the month caption according to the local-name setting
        If UseLocalNames Then
            MonthCaption = MonthName(MonthNumber, False)
        Else
            MonthCaption = M_Caption_GetEnglishMonthFull(MonthNumber)
        End If

    'Return the normalized caption
        M_Caption_GetMonth = UCase$(Trim$(MonthCaption))
End Function

Public Function M_Caption_GetDate( _
    ByVal DateValue As Date, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET DATECAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a DatePicker caption value
'
' WHY THIS EXISTS
'   Caption generation should remain centralized so local-name and fixed-English display modes stay consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates inputs
'   - Builds the requested caption
'   - Returns the normalized display text
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA date formatting functions
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_Caption_GetDate"                  'Current procedure name
    Dim DayText As String                                              'Day component
    Dim MonthText As String                                            'Month component
    Dim YearText As String                                             'Year component

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Build the day component
        DayText = Format$(DateValue, "dd")

    'Build the month component
        If UseLocalNames Then
            MonthText = UCase$(Format$(DateValue, "mmm"))
        Else
            MonthText = M_Caption_GetEnglishMonthShort(VBA.Month(DateValue))
        End If

    'Build the year component
        YearText = Format$(DateValue, "yyyy")

    'Return the final caption
        M_Caption_GetDate = DayText & "-" & MonthText & "-" & YearText

    'Exit before the error handler
        Exit Function

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Function

Public Function M_Caption_GetEnglishMonthShort(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISHMONTHSHORTCAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a DatePicker caption value
'
' WHY THIS EXISTS
'   Caption generation should remain centralized so local-name and fixed-English display modes stay consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates inputs
'   - Builds the requested caption
'   - Returns the normalized display text
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA date formatting functions
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Reject invalid month numbers
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, "M_Caption_GetEnglishMonthShort", _
                "MonthNumber must be between 1 and 12."
        End If

    'Return the fixed English short month caption
        Select Case MonthNumber
            Case 1: M_Caption_GetEnglishMonthShort = "JAN"
            Case 2: M_Caption_GetEnglishMonthShort = "FEB"
            Case 3: M_Caption_GetEnglishMonthShort = "MAR"
            Case 4: M_Caption_GetEnglishMonthShort = "APR"
            Case 5: M_Caption_GetEnglishMonthShort = "MAY"
            Case 6: M_Caption_GetEnglishMonthShort = "JUN"
            Case 7: M_Caption_GetEnglishMonthShort = "JUL"
            Case 8: M_Caption_GetEnglishMonthShort = "AUG"
            Case 9: M_Caption_GetEnglishMonthShort = "SEP"
            Case 10: M_Caption_GetEnglishMonthShort = "OCT"
            Case 11: M_Caption_GetEnglishMonthShort = "NOV"
            Case 12: M_Caption_GetEnglishMonthShort = "DEC"
        End Select
End Function

Public Function M_Caption_GetEnglishMonthFull(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISHMONTHCAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a DatePicker caption value
'
' WHY THIS EXISTS
'   Caption generation should remain centralized so local-name and fixed-English display modes stay consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates inputs
'   - Builds the requested caption
'   - Returns the normalized display text
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA date formatting functions
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Reject invalid month numbers
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, "M_Caption_GetEnglishMonthFull", _
                "MonthNumber must be between 1 and 12."
        End If

    'Return the fixed English full month caption
        Select Case MonthNumber
            Case 1: M_Caption_GetEnglishMonthFull = "JANUARY"
            Case 2: M_Caption_GetEnglishMonthFull = "FEBRUARY"
            Case 3: M_Caption_GetEnglishMonthFull = "MARCH"
            Case 4: M_Caption_GetEnglishMonthFull = "APRIL"
            Case 5: M_Caption_GetEnglishMonthFull = "MAY"
            Case 6: M_Caption_GetEnglishMonthFull = "JUNE"
            Case 7: M_Caption_GetEnglishMonthFull = "JULY"
            Case 8: M_Caption_GetEnglishMonthFull = "AUGUST"
            Case 9: M_Caption_GetEnglishMonthFull = "SEPTEMBER"
            Case 10: M_Caption_GetEnglishMonthFull = "OCTOBER"
            Case 11: M_Caption_GetEnglishMonthFull = "NOVEMBER"
            Case 12: M_Caption_GetEnglishMonthFull = "DECEMBER"
        End Select
End Function

'
'------------------------------------------------------------------------------
'
'                                    TIMER
'
'------------------------------------------------------------------------------

Public Sub M_Timer_ApplyClockMode()

'
'------------------------------------------------------------------------------
'                           TIMER APPLYCLOCKMODE
'------------------------------------------------------------------------------
' PURPOSE
'   Manages the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot and requires controlled start, stop, tick, and cleanup behavior
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Updates the loaded form clock when applicable
'   - Schedules or cancels the next timer tick
'   - Stops safely when the form is no longer loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.OnTime
'   VBA.UserForms
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_Timer_ApplyClockMode"             'Current procedure name
    Dim LoadedForm As Object                                           'Loaded UserForm instance

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure settings are loaded before reading gDP_ClockMode
        M_Settings_EnsureLoaded

    'Stop any existing DatePicker timer before applying the current mode
        M_Timer_Stop

    'Loop through loaded UserForms
        For Each LoadedForm In VBA.UserForms
            'Update the loaded DatePicker clock once
                If TypeName(LoadedForm) = DP_FORM_NAME Then
                    LoadedForm.UF_DP_UpdateLiveClock
                    Exit For
                End If
        Next LoadedForm

    'Start the live timer only when live clock mode is enabled
        If gDP_ClockMode = DP_ClockMode_Live Then
            M_Timer_Start
        End If

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Stop the timer after an unexpected clock-mode error
        M_Timer_Stop

    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Timer_Start()

'
'------------------------------------------------------------------------------
'                           STARTTIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Manages the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot and requires controlled start, stop, tick, and cleanup behavior
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Updates the loaded form clock when applicable
'   - Schedules or cancels the next timer tick
'   - Stops safely when the form is no longer loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.OnTime
'   VBA.UserForms
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_Timer_Start"                        'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Exit if the timer is already running
        If mDP_TimerIsRunning Then
            Exit Sub
        End If

    'Build the workbook-qualified timer procedure name
        mDP_TimerProcedureName = M_GetQualifiedMacroName("M_Timer_Tick")

    'Mark the timer as running
        mDP_TimerIsRunning = True

    'Calculate the next timer tick
        mDP_NextTickTime = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)

    'Schedule the next timer tick
        Application.OnTime _
            EarliestTime:=mDP_NextTickTime, _
            Procedure:=mDP_TimerProcedureName, _
            Schedule:=True

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Clear timer state after scheduling failure
        mDP_TimerIsRunning = False

    'Clear next tick time after scheduling failure
        mDP_NextTickTime = 0

    'Clear timer procedure name after scheduling failure
        mDP_TimerProcedureName = vbNullString

    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_Timer_Stop()

'
'------------------------------------------------------------------------------
'                           STOPTIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Manages the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot and requires controlled start, stop, tick, and cleanup behavior
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Updates the loaded form clock when applicable
'   - Schedules or cancels the next timer tick
'   - Stops safely when the form is no longer loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.OnTime
'   VBA.UserForms
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Suppress cancellation errors
        On Error Resume Next

    'Build the timer procedure name if it is missing
        If Len(mDP_TimerProcedureName) = 0 Then
            mDP_TimerProcedureName = M_GetQualifiedMacroName("M_Timer_Tick")
        End If

    'Cancel the next scheduled timer tick when active
        If mDP_TimerIsRunning Then
            If mDP_NextTickTime <> 0 Then
                Application.OnTime _
                    EarliestTime:=mDP_NextTickTime, _
                    Procedure:=mDP_TimerProcedureName, _
                    Schedule:=False
            End If
        End If

    'Clear timer state
        mDP_TimerIsRunning = False

    'Clear next tick time
        mDP_NextTickTime = 0

    'Clear timer procedure name
        mDP_TimerProcedureName = vbNullString

    'Restore normal error handling
        On Error GoTo 0
End Sub

Public Sub M_Timer_Tick()

'
'------------------------------------------------------------------------------
'                           TIMERTICK
'------------------------------------------------------------------------------
' PURPOSE
'   Manages the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot and requires controlled start, stop, tick, and cleanup behavior
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Updates the loaded form clock when applicable
'   - Schedules or cancels the next timer tick
'   - Stops safely when the form is no longer loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.OnTime
'   VBA.UserForms
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim LoadedForm As Object                                           'Loaded UserForm instance
    Dim FormWasFound As Boolean                                        'True if DatePicker form is loaded

    'Enable fail-safe error handling
        On Error GoTo FailSafe

    'Exit if the timer is no longer running
        If Not mDP_TimerIsRunning Then
            Exit Sub
        End If

    'Loop through loaded UserForms
        For Each LoadedForm In VBA.UserForms
            'Update the loaded DatePicker form when found
                If TypeName(LoadedForm) = DP_FORM_NAME Then
                    LoadedForm.UF_DP_UpdateLiveClock
                    FormWasFound = True
                    Exit For
                End If
        Next LoadedForm

    'Stop the timer if the DatePicker form is no longer loaded
        If Not FormWasFound Then
            M_Timer_Stop
            Exit Sub
        End If

    'Build the workbook-qualified timer procedure name
        mDP_TimerProcedureName = M_GetQualifiedMacroName("M_Timer_Tick")

    'Calculate the next timer tick
        mDP_NextTickTime = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)

    'Schedule the next timer tick
        Application.OnTime _
            EarliestTime:=mDP_NextTickTime, _
            Procedure:=mDP_TimerProcedureName, _
            Schedule:=True

    'Exit after successful scheduling
        Exit Sub

FailSafe:
    'Stop the timer after an unexpected callback error
        M_Timer_Stop

    'Write diagnostics to the Immediate Window
        Debug.Print "M_DATEPICKER.M_Timer_Tick: timer stopped after error " & _
            Err.Number & " - " & Err.Description
End Sub

'
'------------------------------------------------------------------------------
'
'                                WINAPI HELPERS
'
'------------------------------------------------------------------------------

Public Function M_Platform_CanUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM CANUSEWINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Provides Mac-safe WinAPI support for DatePicker form behavior
'
' WHY THIS EXISTS
'   Borderless styling and mouse-based positioning require native Windows APIs but must degrade safely on Mac and when disabled
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Exits safely when WinAPI is unavailable
'   - Uses native handles and screen coordinates where required
'   - Suppresses best-effort UI positioning/styling errors
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Windows User32 APIs
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
#If Mac Then
    'Return False on Mac
        M_Platform_CanUseWinAPI = False
#Else
    'Return True on Windows
        M_Platform_CanUseWinAPI = True
#End If
End Function

Public Function M_Platform_ShouldUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM SHOULDUSEWINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Provides Mac-safe WinAPI support for DatePicker form behavior
'
' WHY THIS EXISTS
'   Borderless styling and mouse-based positioning require native Windows APIs but must degrade safely on Mac and when disabled
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Exits safely when WinAPI is unavailable
'   - Uses native handles and screen coordinates where required
'   - Suppresses best-effort UI positioning/styling errors
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Windows User32 APIs
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
    'Return whether WinAPI features should be used
        M_Platform_ShouldUseWinAPI = _
            (M_Platform_CanUseWinAPI And gDP_UseWinAPI)
End Function

Public Sub M_Window_RemoveTitleBar(ByVal Frm As Object)

'
'------------------------------------------------------------------------------
'                           TITLEBAR REMOVE
'------------------------------------------------------------------------------
' PURPOSE
'   Provides Mac-safe WinAPI support for DatePicker form behavior
'
' WHY THIS EXISTS
'   Borderless styling and mouse-based positioning require native Windows APIs but must degrade safely on Mac and when disabled
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Exits safely when WinAPI is unavailable
'   - Uses native handles and screen coordinates where required
'   - Suppresses best-effort UI positioning/styling errors
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Windows User32 APIs
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
#If Mac Then
    'Do nothing on Mac
        Exit Sub
#Else
    #If VBA7 Then
        Dim hWndForm As LongPtr                                       'UserForm window handle
        Dim WindowStyle As LongPtr                                    'Window style bits
    #Else
        Dim hWndForm As Long                                          'UserForm window handle
        Dim WindowStyle As Long                                       'Window style bits
    #End If

    Dim WindowFlags As Long                                           'SetWindowPos flags

    'Suppress borderless styling errors
        On Error GoTo CleanExit

    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Sub

    'Exit if WinAPI features should not be used
        If Not M_Platform_ShouldUseWinAPI Then Exit Sub

    'Resolve the form window handle
        hWndForm = M_Window_GetUserFormHwnd(Frm)

    'Exit if the form window handle is unavailable
        If hWndForm = 0 Then Exit Sub

    #If VBA7 Then
        'Read current window style
            WindowStyle = GetWindowLongPtr(hWndForm, GWL_STYLE)
    #Else
        'Read current window style
            WindowStyle = GetWindowLong(hWndForm, GWL_STYLE)
    #End If

    'Exit if the style cannot be read
        If WindowStyle = 0 Then Exit Sub

    'Remove the caption style bit
        WindowStyle = (WindowStyle And Not WS_CAPTION)

    #If VBA7 Then
        'Write the updated window style
            Call SetWindowLongPtr(hWndForm, GWL_STYLE, WindowStyle)
    #Else
        'Write the updated window style
            Call SetWindowLong(hWndForm, GWL_STYLE, WindowStyle)
    #End If

    'Build non-client refresh flags
        WindowFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
            SWP_NOACTIVATE Or SWP_FRAMECHANGED

    'Force Windows to recalculate the frame
        Call SetWindowPos(hWndForm, 0, 0, 0, 0, 0, WindowFlags)

    'Redraw menu bar and non-client elements
        Call DrawMenuBar(hWndForm)

CleanExit:
    'Best-effort routine
#End If
End Sub

#If VBA7 Then
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As LongPtr
#Else
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As Long
#End If
'
'------------------------------------------------------------------------------
'                           GETUSERFORMHWND
'------------------------------------------------------------------------------
' PURPOSE
'   Executes the DatePicker getuserformhwnd routine
'
' WHY THIS EXISTS
'   This DatePicker operation is centralized in the companion module to keep UserForm rendering, Excel integration, and write-back behavior consistent
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Validates required inputs
'   - Applies the requested DatePicker operation
'   - Exits through the documented error policy
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   DatePicker companion module state
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------
#If Mac Then
    'Return zero on Mac
        M_Window_GetUserFormHwnd = 0
#Else
    Dim FormCaption As String                                         'Current form caption

    'Set safe default
        M_Window_GetUserFormHwnd = 0

    'Suppress handle lookup errors
        On Error GoTo FailSafe

    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Function

    'Read the current form caption
        FormCaption = CStr(Frm.Caption)

    'Exit if the caption is blank
        If LenB(FormCaption) = 0 Then Exit Function

    'Try the most common UserForm window class
        M_Window_GetUserFormHwnd = FindWindow("ThunderDFrame", FormCaption)

    'Try the alternate UserForm window class
        If M_Window_GetUserFormHwnd = 0 Then
            M_Window_GetUserFormHwnd = FindWindow("ThunderXFrame", FormCaption)
        End If

    'Exit after lookup
        Exit Function

FailSafe:
    'Return safe default
        M_Window_GetUserFormHwnd = 0
#End If
End Function

Public Sub M_Window_MoveFormToMouse( _
    ByVal Frm As Object, _
    Optional ByVal OffsetXPx As Long = 0, _
    Optional ByVal OffsetYPx As Long = 12, _
    Optional ByVal CenterOnMouse As Boolean = False)

'
'------------------------------------------------------------------------------
'                           MOVEFORMTOMOUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Provides Mac-safe WinAPI support for DatePicker form behavior
'
' WHY THIS EXISTS
'   Borderless styling and mouse-based positioning require native Windows APIs but must degrade safely on Mac and when disabled
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Exits safely when WinAPI is unavailable
'   - Uses native handles and screen coordinates where required
'   - Suppresses best-effort UI positioning/styling errors
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Windows User32 APIs
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

#If Mac Then
    'Do nothing on Mac
        Exit Sub
#Else
    Dim CursorPoint As POINTAPI                                       'Mouse cursor position
    Dim CursorRect As RECT                                            'One-pixel cursor rectangle
    Dim FormRect As RECT                                              'Current form window rectangle
    Dim MonitorData As MONITORINFO                                    'Nearest monitor information

    #If VBA7 Then
        Dim hWndForm As LongPtr                                       'UserForm window handle
        Dim hMonitor As LongPtr                                       'Nearest monitor handle
    #Else
        Dim hWndForm As Long                                          'UserForm window handle
        Dim hMonitor As Long                                          'Nearest monitor handle
    #End If

    Dim FormWidthPx As Long                                           'Form width in pixels
    Dim FormHeightPx As Long                                          'Form height in pixels
    Dim WorkWidthPx As Long                                           'Monitor work-area width in pixels
    Dim WorkHeightPx As Long                                          'Monitor work-area height in pixels
    Dim TargetX As Long                                               'Target X position in pixels
    Dim TargetY As Long                                               'Target Y position in pixels
    Dim MoveFlags As Long                                             'SetWindowPos movement flags

    'Suppress positioning errors
        On Error GoTo CleanExit

    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Sub

    'Exit if WinAPI features should not be used
        If Not M_Platform_ShouldUseWinAPI Then Exit Sub

    'Exit if the cursor position cannot be read
        If GetCursorPos(CursorPoint) = 0 Then Exit Sub

    'Resolve the native UserForm window handle
        hWndForm = M_Window_GetUserFormHwnd(Frm)

    'Exit if the form window handle is unavailable
        If hWndForm = 0 Then Exit Sub

    'Exit if the form window rectangle cannot be read
        If GetWindowRect(hWndForm, FormRect) = 0 Then Exit Sub

    'Calculate the form width and height in pixels
        FormWidthPx = FormRect.Right - FormRect.Left
        FormHeightPx = FormRect.Bottom - FormRect.Top

    'Exit if resolved form size is invalid
        If FormWidthPx <= 0 Or FormHeightPx <= 0 Then Exit Sub

    'Build a one-pixel rectangle around the cursor
        CursorRect.Left = CursorPoint.X
        CursorRect.Top = CursorPoint.Y
        CursorRect.Right = CursorPoint.X + 1
        CursorRect.Bottom = CursorPoint.Y + 1

    'Resolve the nearest monitor handle
        hMonitor = MonitorFromRect(CursorRect, MONITOR_DEFAULTTONEAREST)

    'Exit if the monitor handle is unavailable
        If hMonitor = 0 Then Exit Sub

    'Initialize the monitor-info structure size
        MonitorData.cbSize = LenB(MonitorData)

    'Exit if monitor information cannot be read
        If GetMonitorInfo(hMonitor, MonitorData) = 0 Then Exit Sub

    'Calculate the monitor work-area width and height
        WorkWidthPx = MonitorData.rcWork.Right - MonitorData.rcWork.Left
        WorkHeightPx = MonitorData.rcWork.Bottom - MonitorData.rcWork.Top

    'Exit if monitor work area is invalid
        If WorkWidthPx <= 0 Or WorkHeightPx <= 0 Then Exit Sub

    'Use the mouse position as the default top-left anchor
        TargetX = CursorPoint.X + OffsetXPx
        TargetY = CursorPoint.Y + OffsetYPx

    'Center the form on the mouse when requested
        If CenterOnMouse Then
            TargetX = CursorPoint.X - (FormWidthPx \ 2) + OffsetXPx
            TargetY = CursorPoint.Y - (FormHeightPx \ 2) + OffsetYPx
        End If

    'Clamp the X position to the current monitor work area
        If FormWidthPx >= WorkWidthPx Then
            TargetX = MonitorData.rcWork.Left
        ElseIf TargetX < MonitorData.rcWork.Left Then
            TargetX = MonitorData.rcWork.Left
        ElseIf TargetX + FormWidthPx > MonitorData.rcWork.Right Then
            TargetX = MonitorData.rcWork.Right - FormWidthPx
        End If

    'Clamp the Y position to the current monitor work area
        If FormHeightPx >= WorkHeightPx Then
            TargetY = MonitorData.rcWork.Top
        ElseIf TargetY < MonitorData.rcWork.Top Then
            TargetY = MonitorData.rcWork.Top
        ElseIf TargetY + FormHeightPx > MonitorData.rcWork.Bottom Then
            TargetY = MonitorData.rcWork.Bottom - FormHeightPx
        End If

    'Build movement flags
        MoveFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW

    'Move the form without resizing, changing z-order, or activating it
        Call SetWindowPos(hWndForm, 0, TargetX, TargetY, 0, 0, MoveFlags)

CleanExit:
    'Best-effort routine
#End If
End Sub


'
'------------------------------------------------------------------------------
'
'                              RIGHT-CLICK MENU
'
'------------------------------------------------------------------------------

Public Sub M_ContextMenu_Update()

'
'------------------------------------------------------------------------------
'                           RIGHTCLICKMENU UPDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_ContextMenu_Update"            'Current procedure name

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure settings are loaded before reading feature flags
        M_Settings_EnsureLoaded

    'Add or remove right-click entries according to the setting
        If gDP_ShowRightClick Then
            M_ContextMenu_Add
        Else
            M_ContextMenu_Remove
        End If

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Public Sub M_ContextMenu_Add()

'
'------------------------------------------------------------------------------
'                           ADDRIGHTCLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Add to the standard cell context menu
        M_ContextMenu_AddToCommandBar M_ContextMenu_GetCommandBar("Cell")

    'Add to the table/list range context menu
        M_ContextMenu_AddToCommandBar M_ContextMenu_GetCommandBar("List Range Popup")
End Sub

Public Sub M_ContextMenu_Remove()

'
'------------------------------------------------------------------------------
'                           REMOVERIGHTCLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Remove from the standard cell context menu
        M_ContextMenu_RemoveFromCommandBar M_ContextMenu_GetCommandBar("Cell")

    'Remove from the table/list range context menu
        M_ContextMenu_RemoveFromCommandBar M_ContextMenu_GetCommandBar("List Range Popup")
End Sub

Private Function M_ContextMenu_GetCommandBar(ByVal CommandBarName As String) As CommandBar

'
'------------------------------------------------------------------------------
'                           GETCOMMANDBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    'Set safe default
        Set M_ContextMenu_GetCommandBar = Nothing

    'Suppress command-bar lookup errors
        On Error Resume Next

    'Retrieve the requested command bar
        Set M_ContextMenu_GetCommandBar = Application.CommandBars(CommandBarName)

    'Restore normal error handling
        On Error GoTo 0
End Function

Private Sub M_ContextMenu_AddToCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           ADDTOCOMMANDBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim ButtonControl As CommandBarButton                              'Created command-bar button

    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Sub

    'Exit when the DatePicker control is already present
        If M_ContextMenu_ContainsDatePicker(TargetCommandBar) Then Exit Sub

    'Add the DatePicker button
        Set ButtonControl = TargetCommandBar.Controls.Add( _
            Type:=msoControlButton, _
            Before:=DP_CONTEXT_MENU_BEFORE, _
            Temporary:=True)

    'Configure the DatePicker button
        With ButtonControl
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .FaceId = DP_CONTEXT_MENU_FACEID
            .Caption = DP_CONTEXT_MENU_CAPTION
            .Tag = DP_CONTEXT_MENU_TAG
            .BeginGroup = True
        End With
End Sub

Private Sub M_ContextMenu_RemoveFromCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           REMOVEFROMCOMMANDBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim Index As Long                                                  'CommandBar control index

    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Sub

    'Loop backward through controls
        For Index = TargetCommandBar.Controls.Count To 1 Step -1

            'Delete DatePicker controls by tag
                If TargetCommandBar.Controls(Index).Tag = DP_CONTEXT_MENU_TAG Then
                    TargetCommandBar.Controls(Index).Delete
                End If

        Next Index
End Sub

Private Function M_ContextMenu_ContainsDatePicker(ByVal TargetCommandBar As CommandBar) As Boolean

'
'------------------------------------------------------------------------------
'                           ISONCOMMANDBAR
'------------------------------------------------------------------------------
' PURPOSE
'   Manages DatePicker right-click menu integration
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently using a stable tag
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Retrieves target command bars safely
'   - Adds missing DatePicker controls
'   - Removes DatePicker controls by tag
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Application.CommandBars
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim ControlItem As CommandBarControl                               'CommandBar control being inspected

    'Set safe default
        M_ContextMenu_ContainsDatePicker = False

    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Function

    'Loop through controls
        For Each ControlItem In TargetCommandBar.Controls

            'Return True when the DatePicker tag is found
                If ControlItem.Tag = DP_CONTEXT_MENU_TAG Then
                    M_ContextMenu_ContainsDatePicker = True
                    Exit Function
                End If

        Next ControlItem
End Function

'
'------------------------------------------------------------------------------
'
'                                  GRID ICON
'
'------------------------------------------------------------------------------

Public Sub M_GridIcon_SetPath(ByVal IconPath As String)

    'Store the optional icon path
        gDP_IconPath = Trim$(IconPath)
End Sub

Public Sub M_GridIcon_Create(Optional ByVal TargetCell As Range)

'
'------------------------------------------------------------------------------
'                         CREATE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or recreates the in-grid DatePicker entry-point shape
'
' WHY THIS EXISTS
'   Users can invoke the DatePicker directly from the worksheet grid without
'   using the right-click menu or Ribbon
'
' INPUTS
'   TargetCell
'     Optional single-cell anchor. ActiveCell is used when omitted
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - resolves the target cell
'   - exits and removes the icon when the grid-icon feature is disabled
'   - resolves the embedded temporary icon file before touching the worksheet
'   - creates a fully formatted temporary shape while ScreenUpdating is disabled
'   - applies the embedded picture icon when available
'   - falls back to the built-in font glyph when the embedded icon is unavailable
'   - replaces the old icon only after the new icon is fully prepared
'   - restores Application.ScreenUpdating deterministically
'
' ERROR POLICY
'   Best-effort UI routine. Suppresses creation failures, removes partial
'   temporary shapes, and restores Application state
'
' DEPENDENCIES
'   Excel Shapes object model
'   M_Settings_EnsureLoaded
'   M_GridIcon_Remove
'   M_GridIcon_EnsureEmbeddedIconFile
'   M_GetQualifiedMacroName
'   DP_GRID_ICON_NAME
'   DP_GRID_ICON_FONT_NAME
'   DP_GRID_ICON_CALENDAR_CODEPOINT
'   DP_GRID_ICON_FONT_SIZE
'   DP_GRID_ICON_FORE_COLOR
'   DP_THEME_COLOR_HEADER
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'
' NOTES
'   The routine deliberately does not call M_GridIcon_Remove at the beginning of
'   the valid create path, because deleting the old icon before the new icon is
'   ready creates visible flicker
'
'   The old icon is replaced only after the new temporary icon has been created,
'   formatted, and filled
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_Create"           'Current procedure name
    Const ICON_SIZE            As Double = 24#                           'Icon width and height
    Const ICON_GAP             As Double = 5#                            'Gap between target cell and icon

    Dim AnchorCell             As Range                                  'Resolved anchor cell
    Dim TargetSheet            As Worksheet                              'Worksheet receiving the icon
    Dim NewIconShape           As Shape                                  'New temporary icon shape
    Dim IconLeft               As Double                                 'Icon left position
    Dim IconTop                As Double                                 'Icon top position
    Dim IconFilePath           As String                                 'Resolved embedded icon file path
    Dim TempShapeName          As String                                 'Temporary icon shape name
    Dim UsedPictureIcon        As Boolean                                'True when picture fill succeeds
    Dim PreviousScreenUpdating As Boolean                                'Previous ScreenUpdating state
    Dim HasScreenState         As Boolean                                'True after ScreenUpdating is captured

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

    'Load settings before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Remove any stale icon and exit when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Remove
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESOLVE EMBEDDED ICON FILE
'------------------------------------------------------------------------------
    'Suppress embedded-icon creation failures and allow fallback glyph rendering
        On Error Resume Next

    'Resolve or create the embedded temporary icon file
        IconFilePath = M_GridIcon_EnsureEmbeddedIconFile()

    'Clear any suppressed embedded-icon error
        Err.Clear

    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
        Else

            'Suppress ActiveCell access failures
                On Error Resume Next

            'Use ActiveCell when no explicit target is supplied
                Set AnchorCell = Application.ActiveCell

            'Clear any suppressed ActiveCell error
                Err.Clear

            'Restore fail-safe error handling
                On Error GoTo FailSafe

        End If

    'Remove stale icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Remove
            Exit Sub
        End If

    'Remove stale icon and exit when the target is not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Remove
            Exit Sub
        End If

    'Normalize merged cells to the top-left cell
        If AnchorCell.MergeCells Then
            Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
        End If

    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet

    'Exit when no worksheet is available
        If TargetSheet Is Nothing Then
            M_GridIcon_Remove
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CALCULATE ICON POSITION
'------------------------------------------------------------------------------
    'Position the icon to the right of the anchor cell
        IconLeft = AnchorCell.Left + AnchorCell.Width + ICON_GAP

    'Vertically center the icon against the anchor cell
        IconTop = AnchorCell.Top + ((AnchorCell.Height - ICON_SIZE) / 2)

    'Build the temporary shape name
        TempShapeName = DP_GRID_ICON_NAME & "_Pending"

'------------------------------------------------------------------------------
' CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
    'Capture the current ScreenUpdating state
        PreviousScreenUpdating = Application.ScreenUpdating

    'Mark ScreenUpdating as captured
        HasScreenState = True

    'Suppress worksheet repaint during icon creation
        Application.ScreenUpdating = False

'------------------------------------------------------------------------------
' REMOVE STALE TEMPORARY SHAPE
'------------------------------------------------------------------------------
    'Suppress stale temporary-shape cleanup errors
        On Error Resume Next

    'Delete a stale temporary icon from a previous interrupted run
        TargetSheet.Shapes(TempShapeName).Delete

    'Clear any suppressed cleanup error
        Err.Clear

    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CREATE TEMPORARY ICON SHAPE
'------------------------------------------------------------------------------
    'Create the new icon with a temporary name
        Set NewIconShape = TargetSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            IconLeft, _
            IconTop, _
            ICON_SIZE, _
            ICON_SIZE)

    'Assign the temporary shape name
        NewIconShape.Name = TempShapeName

    'Hide the shape while it is being formatted
        NewIconShape.Visible = msoFalse

'------------------------------------------------------------------------------
' APPLY BASE SHAPE SETTINGS
'------------------------------------------------------------------------------
    'Apply base icon behavior and callback settings
        With NewIconShape
            .Placement = xlMove
            .AlternativeText = "DatePicker Grid Entry Point"
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .LockAspectRatio = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
            .Line.Visible = msoFalse
        End With

'------------------------------------------------------------------------------
' APPLY EMBEDDED PICTURE ICON
'------------------------------------------------------------------------------
    'Start from fallback mode
        UsedPictureIcon = False

    'Try to apply the embedded picture icon when available
        If LenB(IconFilePath) <> 0 Then
            If LenB(Dir$(IconFilePath, vbNormal)) <> 0 Then

                'Suppress picture-fill failures and fall back to the glyph
                    On Error Resume Next

                'Apply the embedded picture as the shape fill
                    NewIconShape.Fill.UserPicture IconFilePath

                'Store whether the picture fill succeeded
                    UsedPictureIcon = (Err.Number = 0)

                'Clear any suppressed picture-fill error
                    Err.Clear

                'Restore fail-safe error handling
                    On Error GoTo FailSafe

            End If
        End If

'------------------------------------------------------------------------------
' APPLY FALLBACK GLYPH ICON
'------------------------------------------------------------------------------
    'Apply the built-in fallback glyph when no picture icon was applied
        If Not UsedPictureIcon Then

            'Apply fallback shape fill
                With NewIconShape
                    .Fill.Visible = msoTrue
                    .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
                    .Line.Visible = msoFalse
                End With

            'Suppress TextFrame2 formatting failures
                On Error Resume Next

            'Clear any existing text
                NewIconShape.TextFrame2.TextRange.Text = vbNullString

            'Apply the calendar glyph
                NewIconShape.TextFrame2.TextRange.Text = ChrW$(DP_GRID_ICON_CALENDAR_CODEPOINT)

            'Apply the glyph font name
                NewIconShape.TextFrame2.TextRange.Font.Name = DP_GRID_ICON_FONT_NAME

            'Apply the glyph font size
                NewIconShape.TextFrame2.TextRange.Font.Size = DP_GRID_ICON_FONT_SIZE

            'Apply the glyph foreground color
                NewIconShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = DP_GRID_ICON_FORE_COLOR

            'Remove text margins
                NewIconShape.TextFrame2.MarginLeft = 0
                NewIconShape.TextFrame2.MarginRight = 0
                NewIconShape.TextFrame2.MarginTop = 0
                NewIconShape.TextFrame2.MarginBottom = 0

            'Vertically center the glyph
                NewIconShape.TextFrame2.VerticalAnchor = msoAnchorMiddle

            'Horizontally center the glyph
                NewIconShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

            'Clear any suppressed TextFrame2 error
                Err.Clear

            'Restore fail-safe error handling
                On Error GoTo FailSafe

        End If

'------------------------------------------------------------------------------
' ATOMIC REPLACE OLD ICON
'------------------------------------------------------------------------------
    'Suppress old-icon cleanup errors
        On Error Resume Next

    'Delete the tracked old icon when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Delete
        End If

    'Delete any same-named icon on the target sheet
        TargetSheet.Shapes(DP_GRID_ICON_NAME).Delete

    'Clear any suppressed cleanup error
        Err.Clear

    'Restore fail-safe error handling
        On Error GoTo FailSafe

    'Promote the temporary icon to the stable DatePicker icon name
        NewIconShape.Name = DP_GRID_ICON_NAME

    'Store the new icon reference
        Set gDP_GridIconShape = NewIconShape

    'Show the fully formatted icon
        gDP_GridIconShape.Visible = msoTrue

    'Bring the icon to the front
        gDP_GridIconShape.ZOrder msoBringToFront

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress state-restore failures
        On Error Resume Next
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Application.ScreenUpdating = PreviousScreenUpdating
        End If
    'Restore normal error handling
        On Error GoTo 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Suppress cleanup failures
        On Error Resume Next

    'Delete the partially created temporary icon
        If Not NewIconShape Is Nothing Then
            NewIconShape.Delete
        End If
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Application.ScreenUpdating = PreviousScreenUpdating
        End If

    'Restore normal error handling
        On Error GoTo 0

End Sub
Public Sub M_GridIcon_Remove()

'
'------------------------------------------------------------------------------
'                         REMOVE CURRENT GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the currently tracked DatePicker in-grid icon
'
' WHY THIS EXISTS
'   Selection-change and worksheet-change events can fire very frequently. The
'   normal icon-removal path must therefore be lightweight and must not scan all
'   worksheets or workbooks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the tracked DatePicker icon shape when available
'
'   Also attempts to delete a same-named icon from the active worksheet as a
'   cheap fallback for stale references
'
' ERROR POLICY
'   Best-effort cleanup. Suppresses shape-deletion errors
'
' DEPENDENCIES
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   Excel Shapes object model
'
' NOTES
'   This routine is intended for high-frequency UI refresh paths
'
'   Use M_GridIcon_PurgeAll for hard cleanup boundaries such as save, close,
'   print, teardown, or regression reset
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ActiveSheetObject       As Object           'Current active sheet object
    Dim ActiveWorksheet         As Worksheet        'Current active worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' DELETE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Delete the tracked grid icon shape when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Delete
        End If

    'Clear the tracked shape reference
        Set gDP_GridIconShape = Nothing

'------------------------------------------------------------------------------
' DELETE ACTIVE-SHEET FALLBACK
'------------------------------------------------------------------------------
    'Capture the active sheet object safely
        Set ActiveSheetObject = ActiveSheet

    'Use the active sheet only when it is a worksheet
        If TypeOf ActiveSheetObject Is Excel.Worksheet Then
            Set ActiveWorksheet = ActiveSheetObject
        End If

    'Delete a same-named icon from the active worksheet as a cheap stale cleanup
        If Not ActiveWorksheet Is Nothing Then
            ActiveWorksheet.Shapes(DP_GRID_ICON_NAME).Delete
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub
Public Function M_GridIcon_EnsureEmbeddedIconFile() As String

'
'------------------------------------------------------------------------------
'                       ENSURE EMBEDDED GRID ICON FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the path of the embedded DatePicker grid-icon PNG stored in the
'   current user's temporary directory
'
' WHY THIS EXISTS
'   Shape.Fill.UserPicture requires a file path. The project should not
'   distribute a separate icon file, so the icon is embedded as Base64 text in
'   VBA code and materialized to the user's temp folder on demand
'
' INPUTS
'   None
'
' RETURNS
'   Full path of the temporary PNG file
'
' BEHAVIOR
'   - Resolves the user's temp directory
'   - Builds a stable icon file path
'   - Reuses the file if it already exists and appears valid
'   - Recreates the file from embedded Base64 when missing or invalid
'   - Returns the final valid file path
'
' ERROR POLICY
'   Raises a descriptive runtime error if the temp path cannot be resolved, the
'   embedded payload cannot be decoded, or the icon file cannot be written
'
' DEPENDENCIES
'   M_GridIcon_EmbeddedIconBytes
'   M_GridIcon_WriteBytesToFile
'   M_GridIcon_IconFileIsUsable
'
' NOTES
'   The temp file is intentionally not deleted during normal cleanup
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_EnsureEmbeddedIconFile" 'Current procedure name

    Dim TempFolder             As String       'Resolved temp folder
    Dim IconPath               As String       'Resolved icon file path
    Dim IconBytes()            As Byte         'Decoded embedded PNG bytes

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE TEMP PATH
'------------------------------------------------------------------------------
    'Resolve the user's temp folder
        TempFolder = M_GridIcon_GetTempFolder()

    'Reject an empty temp folder
        If LenB(TempFolder) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Temporary folder could not be resolved."
        End If

    'Build the full icon path
        IconPath = M_GridIcon_CombinePath(TempFolder, DP_GRID_ICON_TEMP_FILE_NAME)

'------------------------------------------------------------------------------
' REUSE EXISTING FILE
'------------------------------------------------------------------------------
    'Return the existing temp file when it appears usable
        If M_GridIcon_IconFileIsUsable(IconPath) Then
            M_GridIcon_EnsureEmbeddedIconFile = IconPath
            Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE TEMP FILE
'------------------------------------------------------------------------------
    'Decode the embedded PNG payload
        IconBytes = M_GridIcon_EmbeddedIconBytes()

    'Write the decoded bytes to the temp file
        M_GridIcon_WriteBytesToFile IconPath, IconBytes

'------------------------------------------------------------------------------
' VALIDATE WRITTEN FILE
'------------------------------------------------------------------------------
    'Reject the file if it still does not appear usable
        If Not M_GridIcon_IconFileIsUsable(IconPath) Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Embedded icon file was written but failed validation."
        End If

'------------------------------------------------------------------------------
' RETURN PATH
'------------------------------------------------------------------------------
    'Return the valid icon path
        M_GridIcon_EnsureEmbeddedIconFile = IconPath

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

Private Function M_GridIcon_GetTempFolder() As String

'
'------------------------------------------------------------------------------
'                           GET TEMP FOLDER
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the current user's temporary folder
'
' WHY THIS EXISTS
'   The embedded grid icon needs a writable per-user location because
'   Shape.Fill.UserPicture requires a real file path
'
' INPUTS
'   None
'
' RETURNS
'   Normalized temp folder path without a trailing path separator
'
' BEHAVIOR
'   Tries TEMP first, then TMP
'
' ERROR POLICY
'   Safe default. Returns an empty string on failure
'
' DEPENDENCIES
'   Environ$
'
' NOTES
'   The Windows temp folder should already exist
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress temp-resolution failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE TEMP FOLDER
'------------------------------------------------------------------------------
    'Use TEMP when available
        M_GridIcon_GetTempFolder = Trim$(Environ$("TEMP"))

    'Use TMP when TEMP is unavailable
        If LenB(M_GridIcon_GetTempFolder) = 0 Then
            M_GridIcon_GetTempFolder = Trim$(Environ$("TMP"))
        End If

'------------------------------------------------------------------------------
' NORMALIZE PATH
'------------------------------------------------------------------------------
    'Remove trailing backslash when present
        If Right$(M_GridIcon_GetTempFolder, 1) = "\" Then
            M_GridIcon_GetTempFolder = Left$(M_GridIcon_GetTempFolder, Len(M_GridIcon_GetTempFolder) - 1)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Return safe default
        M_GridIcon_GetTempFolder = vbNullString

End Function

Private Function M_GridIcon_CombinePath( _
    ByVal FolderPath As String, _
    ByVal FileName As String) As String

'
'------------------------------------------------------------------------------
'                           COMBINE PATH
'------------------------------------------------------------------------------
' PURPOSE
'   Combines a folder path and file name using a Windows path separator
'
' WHY THIS EXISTS
'   Repeated temp-file path construction should remain deterministic and simple
'
' INPUTS
'   FolderPath
'     Folder path
'
'   FileName
'     File name
'
' RETURNS
'   Combined file path
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper assumes a Windows Excel environment
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN COMBINED PATH
'------------------------------------------------------------------------------
    'Return the combined path
        If Right$(FolderPath, 1) = "\" Then
            M_GridIcon_CombinePath = FolderPath & FileName
        Else
            M_GridIcon_CombinePath = FolderPath & "\" & FileName
        End If

End Function

Private Function M_GridIcon_IconFileIsUsable(ByVal IconPath As String) As Boolean

'
'------------------------------------------------------------------------------
'                         ICON FILE IS USABLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate embedded grid-icon file appears usable
'
' WHY THIS EXISTS
'   A previous temp-file write may have failed or produced a truncated file. The
'   DatePicker should recreate the file when it is missing or invalid
'
' INPUTS
'   IconPath
'     Candidate PNG path
'
' RETURNS
'   True if the file exists, has a plausible size, and has a PNG signature
'
' ERROR POLICY
'   Safe default. Returns False on failure
'
' DEPENDENCIES
'   Dir$
'   FileLen
'
' NOTES
'   This is a lightweight validation, not a full PNG parser
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber             As Integer      'Free file number
    Dim Signature(1 To 8)       As Byte         'PNG signature bytes

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to unusable
        M_GridIcon_IconFileIsUsable = False

'------------------------------------------------------------------------------
' VALIDATE PATH
'------------------------------------------------------------------------------
    'Exit when the path is blank
        If LenB(Trim$(IconPath)) = 0 Then
            Exit Function
        End If

    'Exit when the file does not exist
        If LenB(Dir$(IconPath, vbNormal)) = 0 Then
            Exit Function
        End If

    'Exit when the file is implausibly small
        If FileLen(IconPath) < DP_GRID_ICON_TEMP_MIN_BYTES Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' READ SIGNATURE
'------------------------------------------------------------------------------
    'Get a free file number
        FileNumber = FreeFile

    'Open the file for binary reading
        Open IconPath For Binary Access Read As #FileNumber

    'Read the first eight bytes
        Get #FileNumber, 1, Signature

    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' VALIDATE PNG SIGNATURE
'------------------------------------------------------------------------------
    'Return True only when the PNG signature matches
        M_GridIcon_IconFileIsUsable = _
            (Signature(1) = &H89 And _
             Signature(2) = &H50 And _
             Signature(3) = &H4E And _
             Signature(4) = &H47 And _
             Signature(5) = &HD And _
             Signature(6) = &HA And _
             Signature(7) = &H1A And _
             Signature(8) = &HA)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Suppress close failure
        On Error Resume Next

    'Close the file if it was opened
        If FileNumber <> 0 Then Close #FileNumber

    'Return safe default
        M_GridIcon_IconFileIsUsable = False

    'Restore normal error handling
        On Error GoTo 0

End Function

Private Sub M_GridIcon_WriteBytesToFile( _
    ByVal IconPath As String, _
    ByRef IconBytes() As Byte)

'
'------------------------------------------------------------------------------
'                         WRITE BYTES TO FILE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes decoded embedded icon bytes to a binary file
'
' WHY THIS EXISTS
'   The embedded PNG is stored as Base64 text in code, but Excel requires a
'   physical file path for Shape.Fill.UserPicture
'
' INPUTS
'   IconPath
'     Destination file path
'
'   IconBytes
'     Decoded PNG byte array
'
' RETURNS
'   Nothing
'
' ERROR POLICY
'   Raises a descriptive runtime error if the file cannot be written
'
' DEPENDENCIES
'   FreeFile
'
' NOTES
'   Existing temp files are overwritten
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_WriteBytesToFile"  'Current procedure name

    Dim FileNumber             As Integer      'Free file number

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty destination path
        If LenB(Trim$(IconPath)) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "IconPath cannot be empty."
        End If

    'Reject an empty byte payload
        If M_GridIcon_ByteArrayLength(IconBytes) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "IconBytes cannot be empty."
        End If

'------------------------------------------------------------------------------
' REPLACE EXISTING FILE
'------------------------------------------------------------------------------
    'Delete any existing file before writing
        If LenB(Dir$(IconPath, vbNormal)) <> 0 Then
            Kill IconPath
        End If

'------------------------------------------------------------------------------
' WRITE FILE
'------------------------------------------------------------------------------
    'Get a free file number
        FileNumber = FreeFile

    'Open the target file for binary writing
        Open IconPath For Binary Access Write As #FileNumber

    'Write the byte array to disk
        Put #FileNumber, 1, IconBytes

    'Close the file
        Close #FileNumber

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Suppress close failure
        On Error Resume Next

    'Close the file if it was opened
        If FileNumber <> 0 Then Close #FileNumber

    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Sub

Private Function M_GridIcon_ByteArrayLength(ByRef SourceBytes() As Byte) As Long

'
'------------------------------------------------------------------------------
'                         BYTE ARRAY LENGTH
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of bytes in a byte array
'
' WHY THIS EXISTS
'   Dynamic arrays may be uninitialized. A defensive length helper keeps binary
'   write validation safe
'
' INPUTS
'   SourceBytes
'     Byte array to inspect
'
' RETURNS
'   Number of bytes in the array
'
' ERROR POLICY
'   Safe default. Returns zero for uninitialized arrays
'
' DEPENDENCIES
'   LBound
'   UBound
'
' NOTES
'   This helper assumes a one-dimensional byte array
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return zero for uninitialized arrays
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETURN LENGTH
'------------------------------------------------------------------------------
    'Return byte count
        M_GridIcon_ByteArrayLength = UBound(SourceBytes) - LBound(SourceBytes) + 1

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Return safe default
        M_GridIcon_ByteArrayLength = 0

End Function

Private Function M_GridIcon_EmbeddedIconBytes() As Byte()

'
'------------------------------------------------------------------------------
'                         EMBEDDED ICON BYTES
'------------------------------------------------------------------------------
' PURPOSE
'   Decodes the embedded Base64 PNG payload into a byte array
'
' WHY THIS EXISTS
'   The project stores the grid icon directly in VBA code and materializes it to
'   a temporary file only when needed
'
' INPUTS
'   None
'
' RETURNS
'   Decoded PNG bytes
'
' ERROR POLICY
'   Raises a descriptive runtime error if Base64 decoding fails
'
' DEPENDENCIES
'   MSXML2.DOMDocument.6.0
'
' NOTES
'   Late binding avoids requiring a VBA reference
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_EmbeddedIconBytes" 'Current procedure name

    Dim XmlDoc                 As Object       'MSXML document
    Dim XmlNode                As Object       'MSXML Base64 node
    Dim EncodedText            As String       'Base64 encoded PNG

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' BUILD BASE64 TEXT
'------------------------------------------------------------------------------
    'Load the embedded Base64 payload
        EncodedText = M_GridIcon_EmbeddedIconBase64()

    'Reject an empty embedded payload
        If LenB(EncodedText) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Embedded icon payload is empty."
        End If

'------------------------------------------------------------------------------
' DECODE BASE64
'------------------------------------------------------------------------------
    'Create the MSXML document
        Set XmlDoc = CreateObject("MSXML2.DOMDocument.6.0")

    'Create a Base64 node
        Set XmlNode = XmlDoc.createElement("Base64Data")

    'Configure the node as Base64 binary data
        XmlNode.DataType = "bin.base64"

    'Assign the encoded payload
        XmlNode.Text = EncodedText

    'Return the decoded bytes
        M_GridIcon_EmbeddedIconBytes = XmlNode.NodeTypedValue

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
        Err.Raise Err.Number, PROC_NAME, Err.Description

End Function

Private Function M_GridIcon_EmbeddedIconBase64() As String

'
'------------------------------------------------------------------------------
'                         EMBEDDED ICON BASE64
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the embedded Base64 PNG payload used by the DatePicker grid icon
'
' WHY THIS EXISTS
'   The DatePicker should not require a distributed external icon file, but
'   Shape.Fill.UserPicture requires a real file path. This payload is decoded
'   into the user's temp folder on demand
'
' INPUTS
'   None
'
' RETURNS
'   Base64 text representing a PNG icon
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   None
'
' NOTES
'   Keep chunks short enough to avoid line-continuation limits
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EncodedText            As String       'Base64 payload

'------------------------------------------------------------------------------
' BUILD PAYLOAD
'------------------------------------------------------------------------------
    'Append Base64 chunks
        EncodedText = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAJm0lEQVR42u1ae4xcVRn/fefce2dm59GdZR/dtrSlDxaBAKb8"
        EncodedText = EncodedText & "AcFaDEoMIRKig/5n9A+hgiIYgmBkWDQxKhWiJqQlGAJo4g5gUHxQaEgVH4BVa215bFu6ZeljX/OeO3vvOefzj5ld16a7OzP7"
        EncodedText = EncodedText & "rvslJ3sz3+Z3z/nO9/id7x5gWWZVCMy0hPEbl0RPjxx77pnwvFTwZyRc3ZVEIiG//oOnwuO7tQTwxUwBkkkWRMTfevSJm6+6"
        EncodedText = EncodedText & "4XP7z2+L9CZ/9NP7K/OeubvONf4s7AzT/d/Z2fnIk6nckRNDfGIoyz//7R/5mz/cef1M3XWu8WfsAalUSgDEnqMvWbVyZbR1"
        EncodedText = EncodedText & "RVQ1R8PeBeevNFpjGwAcbGujxYo/QwMwHTzYRgATMWutNftKCV9poZQWgsgDmE6+GyWAGxr14TeWE0SdPknbkq9alZcRP9T9"
        EncodedText = EncodedText & "MQUQC+FpqgiIACICsWaA+PHbrvQB4kZGffjgRKJHIpmsa01WHXVIgkjvBRQAXH7nq80hMSpOjrgm6w40n2EolI0Mr/v8q83t"
        EncodedText = EncodedText & "nUoOuAXdyO60K5YDBdLT4Xtuq97ffEUu1U16fK6pW/TsGSDJAt2kL7r9+fM0rfgqATeWfG+jyyARDNO/0m38Ub8MAMIwhPaK"
        EncodedText = EncodedText & "OJKJ3oGgui2bU+SYEDdigAwxiaDEVPi5rIahId48tOcU3fHKn4zv7Ti884ZDlbAgnnkIVBZvNm1/6XqD2BEh7QdY+1fAcJCk"
        EncodedText = EncodedText & "cJRWppAdeJ+EGMvcLKQNv5w+VS4XClI6NoAAqM4BBISQ9nT4wrIdGBWEUZsA+qIMhA9eePvLdwPESExfIaxp3b6b9KZbf/MJ"
        EncodedText = EncodedText & "GQi/JGGeCwYDvxTRoNYaVjgWHv3nm69fHQ25WxkEMBio/DFs9NDAid0b16/eM6rZAPVVK6N9EY7FasQXBhpgCTVaLF6lRWjH"
        EncodedText = EncodedText & "5lt/39y785MPJBI9MjVFOFhTZXn0wKz98otxQc6zZEZfaF3TuaPkeq0WANsBK8+1y+ViuxUTJKUAUSURSyHZlpK0cluyI8PB"
        EncodedText = EncodedText & "5o61Hyi/bJOQNYcCOVbd+CyIYp0dzw2dHshwKNK9+UsvvpLadeMfpsoJk4dADwSIOGCsu8gOxMKxFT2e67XCwNZKWZqNHC37"
        EncodedText = EncodedText & "ASk4kC2jOJLJY9TzCADSmSxlSlwgIqm1ChpjpNawtOLaRwP4rEjki/nOWHP8r6zcE5DyewATLk5MavjJ/bLnQWzDtVa+pB5l"
        EncodedText = EncodedText & "rVqDkfCzEFJDM5EQzNoIJxR1C8Mn155Ol5ts8pUuDZ93rP8Dennf0VO73yr32pZwV63veh0STCSqJay20Sg+wwgpLc/N5beS"
        EncodedText = EncodedText & "kB+OX/7OEyOPXJgFkgLYyzWGABOI+L3tL0YdyFXSCVEpm7u6tX3VcyVVaGeAIQgCWsQ7zv9bufT2+l/vL7y3//1jA0GbrN5B"
        EncodedText = EncodedText & "lRHgcFvHqj870eaSX8wGUYf7V7MA1YvP2ohQqCldyKUvZNBGkg6M43cB6EfiEkKqgTJIBsRakzL6CyPpwXQ4Gn+DBDMAKK2p"
        EncodedText = EncodedText & "Y/UF77HxfzF08vg1x9PlZm2ASNAuN7e07Vl70WUHtO9Ly2kqNlIG68Fno0nYwhTzhc2lons3EUJgAxjmhquA8UoEO0pgBQbH"
        EncodedText = EncodedText & "PLf8oFfq7yWCyyABZgaIpAx67Z2b7LhX7gDYsuzAgJR21+DxD5wqUoMk1aA2fIz/L7P5EElbsva0kLaE0TQDA7gEO1phvqwN"
        EncodedText = EncodedText & "syGSzmYmATBX5sgAgwDJcIJWtUiJFkOV36ZLNVOLrBF/wjuUD1Yej50NmA3NnAmONyCIWPmTbidj7HxOjDmQmvAJlXrJXNMc"
        EncodedText = EncodedText & "rLpnQVOxR8Jsd4PmGt+q5ZUEVEnI4uxHTuItNc12WgP4plKYfU8t2obsWRcmBSxJ0KoBA1RyO3Dd6hyaVkg0hZpw03VXI+A4"
        EncodedText = EncodedText & "4EXuB8wMIsLu197EqZHTgHLxGIBEAkjVywNWx4BItIxIxMK2S9tmkM3nX/reBSyvBMHVFgFwNh40fQiMGoKjCIWSj1BQLBkP"
        EncodedText = EncodedText & "KPuAZwQsrWchCRIgBUEKWgIGqLbMCLOTBM9mEKU0VNWyQhAc2/5fCrvA+rnpCY5VBaVwanAY2lTKDDOjrSWOcFOwqtc4NTi0"
        EncodedText = EncodedText & "YPo5b4trbWAMQxBBVBsU/oQ401ovmJ7nwwBnEqTJaNlC6+csBESVZusq1TaGx1tVqDLGhdTPuQFsx0bbeXEopaoGEQg3hcb1"
        EncodedText = EncodedText & "gQXUUwNhYDVitXAo2LCeGQgFg2c/4swCPubDADM6z9Hi4hHzxgPGCFS+5OPpPYeRLfkVYmUYn732AqzviMIwIOgc5QGhUBBE"
        EncodedText = EncodedText & "wFN7DuPOJ99CezwMZsZIwcOB41k8c89W0CzwjDk3wEQeQAA086R1eqJ+LI3lSz7a4mF0toahNSMUcuB6Zjw8GsVfMjxAymoZ"
        EncodedText = EncodedText & "MwzDDGMYUtD/Dw8whpEueAiHnPHnUV+M7+A5ywPGpviZretxoC8L1zOQguD5Al/5VNf4Ufac5wEbO2N45p6tk3rXfPMAgXkW"
        EncodedText = EncodedText & "Zobhic3Lihuf8zxgYgwrpeEt9wOWSD+AzvgKs9T6AUJM/YVoSgMEHIcYLMdcdyn1A8bmyyysBg3AdDJ/oiRIDLqjHkayOWb+"
        EncodedText = EncodedText & "r9tpY6CMmbxOL4S+utm+rzCcyRIx+0XXPQkAicTZb4lYZz+xEff0sLzllu5S94+f3MskNr554B1z03XXCMuy0NrSDKU0iCav"
        EncodedText = EncodedText & "076v5l2vjYElJQ4d6TMjuSLZgo4cLg68nUwmBRGZujzgYKJCzDLp4UclWP/9UC96j/WzlAKRphCaYxGsiEYQjTRVrqYwj49w"
        EncodedText = EncodedText & "KDjvegCwpESuUMTu195QkUiEPG90R6q72wOuFVNVtUklmUxa3d3d6t7v/+QbHSvXfDefS+uPbLkMV17aJWKR8KL6PDDq+Th0"
        EncodedText = EncodedText & "+JjZ85d9CnbAyWeGf7X3+Wc+3dXVRbt27fIbMgAAuvjihH3oUIrve/ixb0ejK+4dVQohW2Jla4uZCQefbRnO5EQ6X0Q4HEG5"
        EncodedText = EncodedText & "mHuh998Htve/84/B9vZ2TqVSulEDAIBcs2aN09/fX95+30M3d65e+zUmXMkQocWz/wQ2SttSvpvJDO96+P67Hgfgb9myhfft"
        EncodedText = EncodedText & "2+dPR+xqIX/Whg0bmo4ePVoGEP34jYmLVq1ft0H5vhAL/M2chTBOIMDZkeHTv3v2Zwdd1x1saWkJRKNRr6+vz8M0F5RqnbwA"
        EncodedText = EncodedText & "IFs2bQr5p0/b+Xx+4uWfhbw5wWPHibHT9Lp163RfX5+Lyq12Xcvu1sMaJQArHo87wWDQNsYIY4xYUOcnYiGEcV3Xz2azPgC/"
        EncodedText = EncodedText & "unBTy+mY6gq0yhATxmw1Zhrd/TO9wExYONe6qNnuhi1UKKCehS/LsizLsiwLgP8AgyHWpjCmt5QAAAAASUVORK5CYII="

'------------------------------------------------------------------------------
' RETURN PAYLOAD
'------------------------------------------------------------------------------
    'Return the Base64 payload
        M_GridIcon_EmbeddedIconBase64 = EncodedText

End Function

Public Sub M_GridIcon_PurgeAll()

'
'------------------------------------------------------------------------------
'                         PURGE ALL GRID ICONS
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker in-grid icon shapes from all open workbooks
'
' WHY THIS EXISTS
'   Hard cleanup boundaries should remove stale DatePicker icons even when the
'   tracked shape reference has been lost after VBA reset, workbook activation,
'   add-in reload, or unexpected UI interruption
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the tracked grid icon shape when available, then scans all open
'   workbooks and deletes shapes named DP_GRID_ICON_NAME
'
' ERROR POLICY
'   Best-effort cleanup. Suppresses workbook, worksheet, protection, and shape
'   deletion errors
'
' DEPENDENCIES
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   M_GridIcon_DeleteNamedShapeAcrossWorkbook
'   Excel Workbooks / Worksheets / Shapes object model
'
' NOTES
'   This routine is intentionally heavier than M_GridIcon_Remove
'
'   Do not call this routine from high-frequency selection-change paths
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurWorkbook             As Workbook        'Workbook being scanned

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' DELETE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Delete the tracked grid icon shape when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Delete
        End If

    'Clear the tracked shape reference
        Set gDP_GridIconShape = Nothing

'------------------------------------------------------------------------------
' PURGE OPEN WORKBOOKS
'------------------------------------------------------------------------------
    'Loop through all open workbooks
        For Each CurWorkbook In Application.Workbooks

            'Delete same-named grid icon shapes from this workbook
                M_GridIcon_DeleteNamedShapeAcrossWorkbook CurWorkbook, DP_GRID_ICON_NAME

        Next CurWorkbook

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub M_GridIcon_DeleteNamedShapeAcrossWorkbook( _
    ByVal TargetWorkbook As Workbook, _
    ByVal TargetShapeName As String)

'
'------------------------------------------------------------------------------
'                           DELETENAMEDSHAPEACROSSWORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Manages the DatePicker in-grid icon
'
' WHY THIS EXISTS
'   The worksheet icon is a transient entry point and must be created, formatted, and removed consistently
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Resolves the target cell
'   - Removes stale icons
'   - Creates or deletes the DatePicker icon
'   - Applies fallback visual formatting
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   Excel Shapes object model
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim CurWorksheet As Worksheet                                     'Worksheet being scanned

    'Suppress cleanup errors
        On Error Resume Next

    'Exit when no workbook is supplied
        If TargetWorkbook Is Nothing Then Exit Sub

    'Exit when shape name is blank
        If LenB(TargetShapeName) = 0 Then Exit Sub

    'Loop through all worksheets
        For Each CurWorksheet In TargetWorkbook.Worksheets

            'Delete the named shape when found
                CurWorksheet.Shapes(TargetShapeName).Delete

        Next CurWorksheet

    'Restore normal error handling
        On Error GoTo 0
End Sub

'
'------------------------------------------------------------------------------
'
'                             LOADED FORM HELPERS
'
'------------------------------------------------------------------------------

Private Function M_FormBridge_GetLoadedForm(ByVal UserFormName As String) As Object

'
'------------------------------------------------------------------------------
'                           GET LOADED FORM
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates the loaded DatePicker UserForm from the companion module
'
' WHY THIS EXISTS
'   The companion module owns write-back and settings while the UserForm owns rendering,
'   so loaded-form access must avoid instantiating the default form accidentally
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Scans loaded UserForms
'   - Calls the required public form method when available
'   - Falls back safely when the form is not loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA.UserForms
'   Frm_DatePicker
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim LoadedForm As Object                                           'Loaded UserForm instance

    'Set safe default
        Set M_FormBridge_GetLoadedForm = Nothing

    'Suppress lookup errors
        On Error GoTo FailSafe

    'Exit if form name is blank
        If Len(Trim$(UserFormName)) = 0 Then Exit Function

    'Loop through loaded UserForms
        For Each LoadedForm In VBA.UserForms

            'Return the matching loaded form
                If StrComp(CStr(LoadedForm.Name), Trim$(UserFormName), vbTextCompare) = 0 Then
                    Set M_FormBridge_GetLoadedForm = LoadedForm
                    Exit Function
                End If

        Next LoadedForm

    'Exit after scan
        Exit Function

FailSafe:
    'Return safe default
        Set M_FormBridge_GetLoadedForm = Nothing
End Function

Public Sub M_FormBridge_RefreshFromActiveCell()

'
'------------------------------------------------------------------------------
'                           REFRESH FORM FROM ACTIVE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the loaded DatePicker form from the current ActiveCell
'
' WHY THIS EXISTS
'   When close-after-selection is disabled, the DatePicker can remain open while
'   the user selects another worksheet cell. The form must then reflect the new
'   cell context instead of closing or keeping stale selected-date state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   - reads ActiveCell safely
'   - if ActiveCell contains a date, stores it as the selected date
'   - otherwise clears the selected-date state
'   - refreshes the loaded DatePicker form when present
'
' ERROR POLICY
'   Safe-default behavior for ActiveCell access
'
'   Raises a descriptive runtime error only if the loaded form refresh fails
'
' DEPENDENCIES
'   Application.ActiveCell
'   M_FormBridge_GetLoadedForm
'   UF_DatePicker.UF_DP_RefreshFromExternalSelection
'
' NOTES
'   This routine avoids direct default-instance references
'
'   It is intended to be called by cDatePickerManager when the user changes
'   worksheet selection while the DatePicker remains open
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_FormBridge_RefreshFromActiveCell"       'Current procedure name
    Dim LoadedForm As Object                                           'Loaded DatePicker form instance
    Dim CellValue As Variant                                           'ActiveCell value snapshot
    Dim ResolvedDate As Date                                           'Resolved ActiveCell date
    Dim HasDate As Boolean                                             'True when ActiveCell contains a date

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Default the resolved date to today
        ResolvedDate = VBA.Date

    'Suppress ActiveCell access errors
        On Error Resume Next

    'Read the current ActiveCell value
        CellValue = Application.ActiveCell.Value

    'Restore controlled error handling
        On Error GoTo ErrorHandler

    'Resolve an ActiveCell date only when the value is usable
        If Not IsError(CellValue) Then
            If VBA.IsDate(CellValue) Then
                ResolvedDate = VBA.DateValue(CDate(CellValue))
                HasDate = True
            End If
        End If

    'Store selected-date state from ActiveCell
        If HasDate Then
            gDP_SelectedDate = ResolvedDate
            gDP_HasSelectedDate = True
        Else
            gDP_SelectedDate = 0
            gDP_HasSelectedDate = False
        End If

    'Retrieve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

    'Refresh the loaded form from the external worksheet selection
        If Not LoadedForm Is Nothing Then
            LoadedForm.UF_DP_RefreshFromExternalSelection ResolvedDate, HasDate
        End If

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Private Sub M_FormBridge_AfterSuccessfulSelection(ByVal SelectedDate As Date)

'
'------------------------------------------------------------------------------
'                           FORM AFTERSUCCESSFULSELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates the loaded DatePicker UserForm from the companion module
'
' WHY THIS EXISTS
'   The companion module owns write-back and settings while the UserForm owns rendering, so loaded-form access must avoid instantiating the default form accidentally
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Scans loaded UserForms
'   - Calls the required public form method when available
'   - Falls back safely when the form is not loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA.UserForms
'   UF_DatePicker
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_FormBridge_AfterSuccessfulSelection"    'Current procedure name
    Dim LoadedForm As Object                                           'Loaded DatePicker form instance

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Retrieve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

    'Notify the loaded form after successful selection
        If Not LoadedForm Is Nothing Then
            LoadedForm.UF_DP_AfterSuccessfulSelection SelectedDate
        End If

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Private Sub M_FormBridge_RefreshSettings()

'
'------------------------------------------------------------------------------
'                           FORM REFRESHSETTINGSDEPENDENTCAPTIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates the loaded DatePicker UserForm from the companion module
'
' WHY THIS EXISTS
'   The companion module owns write-back and settings while the UserForm owns rendering, so loaded-form access must avoid instantiating the default form accidentally
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Scans loaded UserForms
'   - Calls the required public form method when available
'   - Falls back safely when the form is not loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA.UserForms
'   UF_DatePicker
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "M_FormBridge_RefreshSettings" 'Current procedure name
    Dim LoadedForm As Object                                           'Loaded UserForm instance

    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Loop through loaded UserForms
        For Each LoadedForm In VBA.UserForms

            'Refresh the DatePicker form when it is loaded
                If TypeName(LoadedForm) = DP_FORM_NAME Then
                    LoadedForm.UF_DP_RefreshSettings
                    Exit For
                End If

        Next LoadedForm

    'Exit before the error handler
        Exit Sub

ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub

Private Sub M_FormBridge_UnloadLoadedPicker()

'
'------------------------------------------------------------------------------
'                           FORM UNLOADLOADEDDATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Coordinates the loaded DatePicker UserForm from the companion module
'
' WHY THIS EXISTS
'   The companion module owns write-back and settings while the UserForm owns rendering, so loaded-form access must avoid instantiating the default form accidentally
'
' INPUTS
'   See procedure signature
'
' RETURNS
'   See procedure type
'
' BEHAVIOR
'   - Scans loaded UserForms
'   - Calls the required public form method when available
'   - Falls back safely when the form is not loaded
'
' ERROR POLICY
'   Uses the local documented error-handling path or safe-default behavior
'
' DEPENDENCIES
'   VBA.UserForms
'   UF_DatePicker
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'
' UPDATED
'   2026-04-29
'------------------------------------------------------------------------------

    Dim Index As Long                                                  'Loaded UserForms reverse-loop index
    Dim LoadedForm As Object                                           'Loaded UserForm instance

    'Suppress cleanup errors
        On Error Resume Next

    'Stop any active timer before unloading the form
        M_Timer_Stop

    'Loop backwards because unloading changes the UserForms collection
        For Index = VBA.UserForms.Count - 1 To 0 Step -1

            'Get the loaded form instance
                Set LoadedForm = VBA.UserForms.Item(Index)

            'Unload the DatePicker form when found
                If TypeName(LoadedForm) = DP_FORM_NAME Then
                    Unload LoadedForm
                End If

        Next Index

    'Restore normal error handling
        On Error GoTo 0
End Sub

'
'------------------------------------------------------------------------------
'
'                         COMPATIBILITY / STARTUP WRAPPERS
'
'------------------------------------------------------------------------------



Public Sub M_Settings_UseLocalNames(ByVal UseLocalDayNames As Boolean)

    'Delegate to the preferred setting routine
        M_Settings_SetUseLocalDayNames UseLocalDayNames
End Sub

'
'------------------------------------------------------------------------------
'
'                                SHARED HELPERS
'
'------------------------------------------------------------------------------

Private Function M_GetQualifiedMacroName(ByVal ProcedureName As String) As String

'
'------------------------------------------------------------------------------
'                           GET QUALIFIED MACRO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a workbook-qualified macro name suitable for Excel callbacks
'
' WHY THIS EXISTS
'   Shape.OnAction, CommandBarButton.OnAction, and Application.OnTime callbacks
'   should resolve back to the workbook that contains this module
'
' INPUTS
'   ProcedureName
'     Public procedure name to qualify
'
' RETURNS
'   Workbook-qualified macro name
'
' BEHAVIOR
'   Quotes ThisWorkbook.Name and appends the supplied procedure name
'
' ERROR POLICY
'   Raises a descriptive runtime error when ProcedureName is blank
'
' DEPENDENCIES
'   ThisWorkbook
'
' NOTES
'   The routine intentionally returns a workbook-level macro reference rather
'   than a module-qualified reference for broad callback compatibility
'
' UPDATED
'   2026-04-28
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject an empty procedure name
        If Len(Trim$(ProcedureName)) = 0 Then
            Err.Raise vbObjectError + 513, "M_GetQualifiedMacroName", _
                "ProcedureName cannot be empty."
        End If

'------------------------------------------------------------------------------
' RETURN QUALIFIED NAME
'------------------------------------------------------------------------------
    'Return the workbook-qualified macro name
        M_GetQualifiedMacroName = _
            "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & Trim$(ProcedureName)

End Function


