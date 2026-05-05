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
    Private Const DP_SETTINGS_SECTION_BEHAVIOR     As String = "Behavior"                'Behavior settings section
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
    Private Const DP_SETTING_ENABLE_KEYBOARD       As String = "EnableKeyboardShortcut"  'Keyboard shortcut setting key
    Private Const DP_KEYBOARD_SHORTCUT_KEY         As String = "^+d"                     'Ctrl + Shift + D
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
    Public Const DP_DEFAULT_ENABLE_KEYBOARD        As Boolean = True                     'Default keyboard shortcut feature flag
    Public Const DP_KEYBOARD_SHORTCUT_TEXT         As String = "Ctrl + Shift + D"        'Displayed keyboard shortcut
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
    Public gDP_UseWinAPI                As Boolean              'Allow WinAPI features
    Public gDP_EnableKeyboardShortcut   As Boolean              'Enable keyboard shortcut fallback
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
'   before the DatePicker form, manager, context menu, grid icon, or keyboard
'   shortcut infrastructure is initialized
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads Display, Behavior, Feature, and Advanced DatePicker settings from the
'   current user's registry hive
'
'   Falls back to legacy Display-section keys for settings that were moved to
'   the Behavior section in later builds
'
'   Parses all Boolean settings from stable storage values, normalizes invalid
'   or unsupported values to defaults, prevents a dead access configuration, and
'   saves the normalized state back to the registry
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, normalized,
'   or saved
'
' DEPENDENCIES
'   GetSetting
'   M_Settings_Save
'   M_Settings_TryParseFirstDayOfWeek
'   M_Settings_TryParseBoolean
'   M_Settings_TryParseLong
'   M_Settings_IsValidClockMode
'   M_Settings_IsValidSizeMode
'   M_Settings_BooleanToStorageValue
'   M_Platform_CanUseWinAPI
'
' NOTES
'   Boolean values are parsed from stable 1 / 0 strings and common textual forms
'
'   Behavior settings are read from the Behavior section first and from the
'   legacy Display section second, so existing user preferences are preserved
'   after upgrading from earlier DatePicker builds
'
'   Settings are saved after load so invalid, missing, legacy, or unsupported
'   values are normalized in persistent storage
'
'   If right-click integration and in-grid icon integration are both disabled,
'   keyboard shortcut access is forced on to avoid a configuration with no
'   practical manual DatePicker entry point
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_Load"

    Dim RawValue                As String           'Raw setting value read from storage
    Dim ParsedLong              As Long             'Parsed numeric setting value
    Dim ParsedBoolean           As Boolean          'Parsed Boolean setting value
    Dim PlatformCanUseWinAPI    As Boolean          'True when WinAPI helpers are supported
    Dim HandlerStep             As String           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"
    'Resolve platform support once for this load pass
        PlatformCanUseWinAPI = M_Platform_CanUseWinAPI

'------------------------------------------------------------------------------
' LOAD DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Load first-day-of-week setting"

    'Read the first-day-of-week setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_FIRST_DAY_OF_WEEK, _
            VBA.CStr(DP_DEFAULT_FIRST_DAY_OF_WEEK))
    'Store the parsed first-day setting or its default
        If M_Settings_TryParseFirstDayOfWeek(RawValue, ParsedLong) Then
            gDP_FirstDayOfWeek = ParsedLong
        Else
            gDP_FirstDayOfWeek = DP_DEFAULT_FIRST_DAY_OF_WEEK
        End If
    'Track the current handler step
        HandlerStep = "Load local-name setting"
    'Read the local-name setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_USE_LOCAL_NAMES, _
            M_Settings_BooleanToStorageValue(DP_DEFAULT_USE_LOCAL_NAMES))
    'Store the parsed local-name setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_UseLocalNames = ParsedBoolean
        Else
            gDP_UseLocalNames = DP_DEFAULT_USE_LOCAL_NAMES
        End If

    'Track the current handler step
        HandlerStep = "Load weekend-highlight setting"

    'Read the weekend-highlight setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_HIGHLIGHT_WEEKENDS, _
            M_Settings_BooleanToStorageValue(DP_DEFAULT_HIGHLIGHT_WEEKENDS))
    'Store the parsed weekend-highlight setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_HighlightWeekends = ParsedBoolean
        Else
            gDP_HighlightWeekends = DP_DEFAULT_HIGHLIGHT_WEEKENDS
        End If

'------------------------------------------------------------------------------
' LOAD BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Load outside-month selection setting"

    'Read the outside-month setting from the Behavior section
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                DP_SETTINGS_APP_NAME, _
                DP_SETTINGS_SECTION_DISPLAY, _
                DP_SETTING_ALLOW_OUTSIDE_MONTH, _
                M_Settings_BooleanToStorageValue(DP_DEFAULT_ALLOW_OUTSIDE_MONTH))
        End If
    'Store the parsed outside-month setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_AllowOutsideMonthSel = ParsedBoolean
        Else
            gDP_AllowOutsideMonthSel = DP_DEFAULT_ALLOW_OUTSIDE_MONTH
        End If

    'Track the current handler step
        HandlerStep = "Load close-after-selection setting"
        
    'Read the close-after-selection setting from the Behavior section
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOSE_AFTER_SELECTION, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                DP_SETTINGS_APP_NAME, _
                DP_SETTINGS_SECTION_DISPLAY, _
                DP_SETTING_CLOSE_AFTER_SELECTION, _
                M_Settings_BooleanToStorageValue(DP_DEFAULT_CLOSE_AFTER_SELECTION))
        End If
    'Store the parsed close-after-selection setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_CloseAfterSelection = ParsedBoolean
        Else
            gDP_CloseAfterSelection = DP_DEFAULT_CLOSE_AFTER_SELECTION
        End If

    'Track the current handler step
        HandlerStep = "Load clock-mode setting"

    'Read the clock mode setting from the Behavior section
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOCK_MODE, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                DP_SETTINGS_APP_NAME, _
                DP_SETTINGS_SECTION_DISPLAY, _
                DP_SETTING_CLOCK_MODE, _
                VBA.CStr(VBA.CLng(DP_ClockMode_Static)))
        End If
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

    'Track the current handler step
        HandlerStep = "Load size-mode setting"

    'Read the size mode setting from the Behavior section
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_SIZE_MODE, _
            VBA.vbNullString)

    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                DP_SETTINGS_APP_NAME, _
                DP_SETTINGS_SECTION_DISPLAY, _
                DP_SETTING_SIZE_MODE, _
                VBA.CStr(VBA.CLng(DP_SizeMode_Normal)))
        End If
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
    'Track the current handler step
        HandlerStep = "Load right-click integration setting"

    'Read the right-click setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_RIGHT_CLICK, _
            M_Settings_BooleanToStorageValue(DP_DEFAULT_SHOW_RIGHT_CLICK))
    'Store the parsed right-click setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_ShowRightClick = ParsedBoolean
        Else
            gDP_ShowRightClick = DP_DEFAULT_SHOW_RIGHT_CLICK
        End If

    'Track the current handler step
        HandlerStep = "Load in-grid icon setting"

    'Read the grid-icon setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_GRID_ICON, _
            M_Settings_BooleanToStorageValue(DP_DEFAULT_SHOW_GRID_ICON))
    'Store the parsed grid-icon setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_ShowGridIcon = ParsedBoolean
        Else
            gDP_ShowGridIcon = DP_DEFAULT_SHOW_GRID_ICON
        End If

    'Track the current handler step
        HandlerStep = "Load keyboard shortcut setting"

    'Read the keyboard shortcut setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_ENABLE_KEYBOARD, _
            M_Settings_BooleanToStorageValue(DP_DEFAULT_ENABLE_KEYBOARD))
    'Store the parsed keyboard shortcut setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_EnableKeyboardShortcut = ParsedBoolean
        Else
            gDP_EnableKeyboardShortcut = DP_DEFAULT_ENABLE_KEYBOARD
        End If

'------------------------------------------------------------------------------
' LOAD ADVANCED SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Load WinAPI setting"

    'Read the WinAPI setting
        RawValue = GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_USE_WINAPI, _
            M_Settings_BooleanToStorageValue(PlatformCanUseWinAPI))
    'Store the parsed WinAPI setting or its platform default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_UseWinAPI = CBool(ParsedBoolean And PlatformCanUseWinAPI)
        Else
            gDP_UseWinAPI = PlatformCanUseWinAPI
        End If

    'Track the current handler step
        HandlerStep = "Load holiday callback setting"

    'Read the holiday callback setting
        gDP_HolidayCallbackName = VBA.Trim$(GetSetting( _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_HOLIDAY_CALLBACK, _
            DP_DEFAULT_HOLIDAY_CALLBACK))

'------------------------------------------------------------------------------
' NORMALIZE ACCESS SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize access settings"

    'Force keyboard access when both contextual and visible entry points are disabled
        If Not gDP_ShowRightClick Then
            If Not gDP_ShowGridIcon Then
                gDP_EnableKeyboardShortcut = True
            End If
        End If

'------------------------------------------------------------------------------
' FINALIZE SETTINGS LOAD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Finalize settings load"
    'Mark settings as loaded before saving normalized values
        mSettingsLoaded = True
    'Track the current handler step
        HandlerStep = "Save normalized settings"
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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker settings load failed: " & Err.Description

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
'   User preferences should persist between DatePicker sessions and should be
'   stored in one canonical settings layout after load-time migration and
'   normalization have been applied
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Initializes default settings when no in-memory settings state is available,
'   validates current DatePicker settings, normalizes settings that must remain
'   safe or platform-compatible, and persists the canonical Display, Behavior,
'   Feature, and Advanced registry sections
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings are invalid, cannot be
'   normalized, or cannot be saved
'
' DEPENDENCIES
'   SaveSetting
'   M_Settings_InitializeDefaults
'   M_Settings_IsValidFirstDayOfWeek
'   M_Settings_IsValidClockMode
'   M_Settings_IsValidSizeMode
'   M_Settings_BooleanToStorageValue
'   M_Platform_CanUseWinAPI
'
' NOTES
'   Boolean values are saved as 1 / 0 for stable parsing
'
'   Behavior settings are persisted under the Behavior section. The load path
'   still supports the previous Display-section location for migration purposes
'
'   If right-click menu and in-grid icon access are both disabled, keyboard
'   shortcut access is forced on to avoid a configuration with no practical
'   manual DatePicker entry point
'
'   WinAPI styling is saved as disabled when the current platform does not
'   support WinAPI helpers
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_Save"
    
    Dim PlatformCanUseWinAPI    As Boolean      'True when WinAPI helpers are supported
    Dim HandlerStep             As String       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' INITIALIZE DEFAULTS IF NEEDED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Initialize defaults if needed"

    'Initialize default settings if settings have not yet been loaded
        If Not mSettingsLoaded Then
            M_Settings_InitializeDefaults
        End If

'------------------------------------------------------------------------------
' NORMALIZE PLATFORM SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize platform settings"

    'Resolve platform support once for this save pass
        PlatformCanUseWinAPI = M_Platform_CanUseWinAPI
    'Disable WinAPI styling when the current platform does not support it
        If Not PlatformCanUseWinAPI Then
            gDP_UseWinAPI = False
        End If

'------------------------------------------------------------------------------
' NORMALIZE ACCESS SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize access settings"

    'Keep the keyboard shortcut enabled when all visible/contextual entry points are disabled
        If Not gDP_ShowRightClick Then
            If Not gDP_ShowGridIcon Then
                gDP_EnableKeyboardShortcut = True
            End If
        End If

'------------------------------------------------------------------------------
' NORMALIZE TEXT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize text settings"

    'Normalize the holiday callback name before persistence
        gDP_HolidayCallbackName = VBA.Trim$(gDP_HolidayCallbackName)

'------------------------------------------------------------------------------
' VALIDATE SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate first-day setting"

    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If
        
    'Track the current handler step
        HandlerStep = "Validate clock mode"

    'Reject unsupported clock modes
        If Not M_Settings_IsValidClockMode(gDP_ClockMode) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "gDP_ClockMode is unsupported"
        End If

    'Track the current handler step
        HandlerStep = "Validate size mode"

    'Reject unsupported size modes
        If Not M_Settings_IsValidSizeMode(gDP_SizeMode) Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "gDP_SizeMode is unsupported"
        End If

'------------------------------------------------------------------------------
' SAVE DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Save first-day-of-week setting"

    'Save the first-day-of-week setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_FIRST_DAY_OF_WEEK, _
            VBA.CStr(gDP_FirstDayOfWeek)

    'Track the current handler step
        HandlerStep = "Save local-name setting"

    'Save the local-name setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_USE_LOCAL_NAMES, _
            M_Settings_BooleanToStorageValue(gDP_UseLocalNames)

    'Track the current handler step
        HandlerStep = "Save weekend-highlight setting"

    'Save the weekend-highlight setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_HIGHLIGHT_WEEKENDS, _
            M_Settings_BooleanToStorageValue(gDP_HighlightWeekends)

'------------------------------------------------------------------------------
' SAVE BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Save outside-month selection setting"

    'Save the outside-month setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH, _
            M_Settings_BooleanToStorageValue(gDP_AllowOutsideMonthSel)

    'Track the current handler step
        HandlerStep = "Save close-after-selection setting"

    'Save the close-after-selection setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOSE_AFTER_SELECTION, _
            M_Settings_BooleanToStorageValue(gDP_CloseAfterSelection)

    'Track the current handler step
        HandlerStep = "Save clock-mode setting"

    'Save the clock mode setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOCK_MODE, _
            VBA.CStr(VBA.CLng(gDP_ClockMode))

    'Track the current handler step
        HandlerStep = "Save size-mode setting"

    'Save the size mode setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_SIZE_MODE, _
            VBA.CStr(VBA.CLng(gDP_SizeMode))

'------------------------------------------------------------------------------
' SAVE FEATURE SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Save right-click integration setting"

    'Save the right-click setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_RIGHT_CLICK, _
            M_Settings_BooleanToStorageValue(gDP_ShowRightClick)

    'Track the current handler step
        HandlerStep = "Save in-grid icon setting"

    'Save the grid-icon setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_GRID_ICON, _
            M_Settings_BooleanToStorageValue(gDP_ShowGridIcon)

    'Track the current handler step
        HandlerStep = "Save keyboard shortcut setting"

    'Save the keyboard shortcut setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_ENABLE_KEYBOARD, _
            M_Settings_BooleanToStorageValue(gDP_EnableKeyboardShortcut)

'------------------------------------------------------------------------------
' SAVE ADVANCED SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Save WinAPI setting"

    'Save the WinAPI setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_USE_WINAPI, _
            M_Settings_BooleanToStorageValue(gDP_UseWinAPI)

    'Track the current handler step
        HandlerStep = "Save holiday callback setting"

    'Save the holiday callback setting
        SaveSetting _
            DP_SETTINGS_APP_NAME, _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_HOLIDAY_CALLBACK, _
            gDP_HolidayCallbackName

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
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker settings save failed: " & Err.Description

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
'   Assigns default values to all persisted settings, applies safe platform
'   defaults, and prevents a dead access configuration
'
' ERROR POLICY
'   Best-effort initialization
'
'   If platform capability detection fails, WinAPI styling is initialized as
'   disabled because native MSForms rendering is the safest fallback
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'
' NOTES
'   This routine changes in-memory settings only
'
'   The keyboard shortcut is forced on if both right-click menu and in-grid icon
'   access are disabled
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PlatformCanUseWinAPI    As Boolean      'True when WinAPI helpers are supported

'------------------------------------------------------------------------------
' INITIALIZE PLATFORM DEFAULTS
'------------------------------------------------------------------------------
    'Suppress platform-detection errors
        On Error Resume Next
    'Resolve whether WinAPI helpers can be used on the current platform
        PlatformCanUseWinAPI = M_Platform_CanUseWinAPI
    'Clear any suppressed platform-detection error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' INITIALIZE DISPLAY SETTINGS
'------------------------------------------------------------------------------
    'Initialize the first-day setting
        gDP_FirstDayOfWeek = DP_DEFAULT_FIRST_DAY_OF_WEEK
    'Initialize the local-name setting
        gDP_UseLocalNames = DP_DEFAULT_USE_LOCAL_NAMES
    'Initialize the weekend-highlight setting
        gDP_HighlightWeekends = DP_DEFAULT_HIGHLIGHT_WEEKENDS

'------------------------------------------------------------------------------
' INITIALIZE BEHAVIOR SETTINGS
'------------------------------------------------------------------------------
    'Initialize the outside-month setting
        gDP_AllowOutsideMonthSel = DP_DEFAULT_ALLOW_OUTSIDE_MONTH
    'Initialize the close-after-selection setting
        gDP_CloseAfterSelection = DP_DEFAULT_CLOSE_AFTER_SELECTION
    'Initialize the clock mode
        gDP_ClockMode = DP_ClockMode_Static
    'Initialize the size mode
        gDP_SizeMode = DP_SizeMode_Normal

'------------------------------------------------------------------------------
' INITIALIZE FEATURE SETTINGS
'------------------------------------------------------------------------------
    'Initialize the right-click feature setting
        gDP_ShowRightClick = DP_DEFAULT_SHOW_RIGHT_CLICK
    'Initialize the grid-icon feature setting
        gDP_ShowGridIcon = DP_DEFAULT_SHOW_GRID_ICON
    'Initialize the keyboard shortcut feature setting
        gDP_EnableKeyboardShortcut = DP_DEFAULT_ENABLE_KEYBOARD
    'Force keyboard access when both visual/contextual entry points are disabled
        If Not gDP_ShowRightClick Then
            If Not gDP_ShowGridIcon Then
                gDP_EnableKeyboardShortcut = True
            End If
        End If

'------------------------------------------------------------------------------
' INITIALIZE ADVANCED SETTINGS
'------------------------------------------------------------------------------
    'Initialize the WinAPI styling setting from the safe platform capability result
        gDP_UseWinAPI = PlatformCanUseWinAPI
    'Initialize the holiday callback name
        gDP_HolidayCallbackName = VBA.Trim$(DP_DEFAULT_HOLIDAY_CALLBACK)

'------------------------------------------------------------------------------
' FINALIZE
'------------------------------------------------------------------------------
    'Mark settings as loaded
        mSettingsLoaded = True

End Sub
Public Sub M_Settings_EnsureLoaded()

'
'------------------------------------------------------------------------------
'                           ENSURE SETTINGS LOADED
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures DatePicker settings have been loaded before they are read or changed
'
' WHY THIS EXISTS
'   Public callers, form lifecycle routines, demo sheets, and settings panels
'   should be safe to call before settings have been explicitly loaded
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits immediately when settings are already loaded
'
'   Loads settings from the persisted settings store only when they have not
'   already been initialized
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_Load
'
' NOTES
'   This helper avoids overwriting persisted settings with default values
'
'   The routine is Public because it is intentionally used by UF_DatePicker and
'   other project modules
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_EnsureLoaded"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF ALREADY LOADED
'------------------------------------------------------------------------------
    'Exit when settings have already been initialized
        If mSettingsLoaded Then Exit Sub

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Load persisted settings into project state
        M_Settings_Load

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
        Err.Raise Err.Number, PROC_NAME, "Settings load check failed: " & Err.Description

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
'   Loads settings if needed, validates the requested first-day setting, exits
'   immediately when the requested value is already active, updates the in-memory
'   setting, persists it, and refreshes the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if FirstDayOfWeek is unsupported, if
'   settings cannot be loaded, if settings cannot be saved, or if the loaded
'   DatePicker form cannot be refreshed
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory value is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Supported values are vbSunday and vbMonday
'
'   The form is refreshed only after a successful persistence step
'
'   If the form refresh fails after persistence succeeds, the saved preference is
'   not rolled back because the user setting has already been committed
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_SetFirstDayOfWeek"

    Dim OldFirstDay            As Long                  'Previous first-day setting
    Dim SettingChanged         As Boolean               'True when the requested value differs
    Dim SettingMutated         As Boolean               'True after in-memory state is changed
    Dim SettingPersisted       As Boolean               'True after settings are saved successfully
    Dim ErrorNumber            As Long                  'Captured error number
    Dim ErrorDescription       As String                'Captured error description

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
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current first-day setting
        OldFirstDay = gDP_FirstDayOfWeek
    'Resolve whether the requested setting is different
        SettingChanged = (FirstDayOfWeek <> OldFirstDay)

'------------------------------------------------------------------------------
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Exit when no setting change is required
        If Not SettingChanged Then Exit Sub

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested first-day setting
        gDP_FirstDayOfWeek = FirstDayOfWeek
    'Mark the in-memory setting as mutated
        SettingMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Persist the updated settings
        M_Settings_Save
    'Mark the setting as persisted
        SettingPersisted = True

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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting when persistence failed after mutation
        If SettingMutated And Not SettingPersisted Then
            On Error Resume Next
            gDP_FirstDayOfWeek = OldFirstDay
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "First-day-of-week update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetFirstDayOfWeekText(ByVal FirstDayOfWeekText As String)

'
'------------------------------------------------------------------------------
'                       SET FIRST DAY OF WEEK FROM TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Sets the DatePicker first-day-of-week preference from worksheet-friendly text
'
' WHY THIS EXISTS
'   Settings sheets, demo sheets, and configuration tables often store selections
'   as text such as vbMonday, Monday, Mon, vbSunday, Sunday, or Sun rather than
'   as VBA weekday constants
'
' INPUTS
'   FirstDayOfWeekText
'     Text value resolving to vbSunday or vbMonday
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Normalizes the supplied text, parses it into a supported first-day-of-week
'   value, and delegates validation, persistence, rollback, and open-form refresh
'   to M_Settings_SetFirstDayOfWeek
'
' ERROR POLICY
'   Raises a descriptive runtime error if FirstDayOfWeekText is blank, if it
'   cannot be parsed, or if the delegated first-day setting update fails
'
' DEPENDENCIES
'   M_Settings_TryParseFirstDayOfWeek
'   M_Settings_SetFirstDayOfWeek
'
' NOTES
'   Accepted values depend on M_Settings_TryParseFirstDayOfWeek and should include
'   vbSunday, Sunday, Sun, 1, vbMonday, Monday, Mon, and 2
'
'   This routine intentionally delegates persistence and form refresh so there is
'   only one canonical first-day update path
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetFirstDayOfWeekText"

    Dim EffectiveText       As String           'Normalized input text
    Dim ParsedValue         As Long             'Parsed first-day value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied first-day text
        EffectiveText = Trim$(FirstDayOfWeekText)
    'Reject blank first-day text
        If Len(EffectiveText) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "FirstDayOfWeekText cannot be blank"
        End If

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse the supplied first-day text
        If Not M_Settings_TryParseFirstDayOfWeek(EffectiveText, ParsedValue) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "FirstDayOfWeekText must resolve to vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' SAVE SETTING
'------------------------------------------------------------------------------
    'Save the parsed first-day setting through the canonical setter
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
        Err.Raise Err.Number, PROC_NAME, _
            "First-day-of-week text update failed: " & Err.Description

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
'   Loads settings if needed, exits when the requested value is already active,
'   updates the in-memory setting, persists the updated settings, and refreshes
'   the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, saved, or
'   applied to the loaded DatePicker form
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   The argument name is preserved for compatibility, but this setting controls
'   both day and month captions
'
'   The unchanged-value exit avoids unnecessary registry writes and unnecessary
'   UserForm refresh work
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetUseLocalDayNames"

    Dim OldUseLocalNames    As Boolean          'Previous local-name setting
    Dim SettingsMutated     As Boolean          'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean          'True after settings are saved successfully
    Dim ErrorNumber         As Long             'Captured error number
    Dim ErrorDescription    As String           'Captured error description

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
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested setting is already active
        If gDP_UseLocalNames = UseLocalDayNames Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current local-name setting for rollback
        OldUseLocalNames = gDP_UseLocalNames

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested local-name setting
        gDP_UseLocalNames = UseLocalDayNames
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_UseLocalNames = OldUseLocalNames
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "Use-local-names setting update failed: " & ErrorDescription

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
'   Loads settings if needed, exits when the requested value is already active,
'   updates the in-memory setting, persists the updated settings, and refreshes
'   the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, saved, or
'   applied to the loaded DatePicker form
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   The form should call M_DatePolicy_CanSelectDate before accepting a day click
'
'   The unchanged-value exit avoids unnecessary registry writes and unnecessary
'   UserForm refresh work
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetAllowOutsideMonthDays"

    Dim OldAllowOutside     As Boolean      'Previous outside-month setting
    Dim SettingsMutated     As Boolean      'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean      'True after settings are saved successfully
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description

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
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested setting is already active
        If gDP_AllowOutsideMonthSel = AllowOutsideMonthDays Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current outside-month setting for rollback
        OldAllowOutside = gDP_AllowOutsideMonthSel

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested outside-month setting
        gDP_AllowOutsideMonthSel = AllowOutsideMonthDays
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_AllowOutsideMonthSel = OldAllowOutside
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "Allow-outside-month-days setting update failed: " & ErrorDescription

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
'   Loads settings if needed, exits when the requested value is already active,
'   updates the in-memory weekend-highlight setting, persists the updated
'   settings, and refreshes the loaded DatePicker form if present
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, saved, or
'   applied to the loaded DatePicker form
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   This setting controls weekend visual highlighting only
'
'   The unchanged-value exit avoids unnecessary registry writes and unnecessary
'   UserForm refresh work
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetHighlightWeekends"

    Dim OldHighlight        As Boolean      'Previous weekend-highlight setting
    Dim SettingsMutated     As Boolean      'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean      'True after settings are saved successfully
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description

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
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested setting is already active
        If gDP_HighlightWeekends = HighlightWeekends Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current weekend-highlight setting for rollback
        OldHighlight = gDP_HighlightWeekends

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested weekend-highlight setting
        gDP_HighlightWeekends = HighlightWeekends
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_HighlightWeekends = OldHighlight
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "Weekend-highlight setting update failed: " & ErrorDescription

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
'   Loads settings if needed, exits when the requested value is already active,
'   updates the in-memory close-after-selection setting, and persists the
'   updated settings
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or saved
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   This setting controls form lifecycle after successful write-back only
'
'   It does not require a visual refresh of the loaded DatePicker form because
'   the value is consumed at the next successful selection event
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetCloseAfterSelection"

    Dim OldCloseAfter       As Boolean      'Previous close-after-selection setting
    Dim SettingsMutated     As Boolean      'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean      'True after settings are saved successfully
    Dim ErrorNumber         As Long         'Captured error number
    Dim ErrorDescription    As String       'Captured error description

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
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested setting is already active
        If gDP_CloseAfterSelection = CloseAfterSelection Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current setting for rollback
        OldCloseAfter = gDP_CloseAfterSelection

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested close-after-selection setting
        gDP_CloseAfterSelection = CloseAfterSelection
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_CloseAfterSelection = OldCloseAfter
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "Close-after-selection setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetClockMode(ByVal ClockMode As DP_ClockMode)

'
'------------------------------------------------------------------------------
'                           SET CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves the DatePicker clock-mode preference
'
' WHY THIS EXISTS
'   The DatePicker footer can display either a static time value or a live clock.
'   Caller code, demo sheets, Ribbon callbacks, settings panels, and host
'   workbooks need one controlled public entry point to update this setting
'   consistently
'
' INPUTS
'   ClockMode
'     Requested DatePicker clock mode
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the requested clock mode, ensures current settings are loaded,
'   avoids unnecessary registry writes when the setting is unchanged, updates the
'   in-memory clock-mode setting when required, persists the updated settings,
'   and synchronizes the live-clock timer with the active mode
'
' ERROR POLICY
'   Raises a descriptive runtime error if ClockMode is unsupported, settings
'   cannot be loaded, settings cannot be saved, or the clock timer cannot be
'   synchronized
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory clock mode is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_IsValidClockMode
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Timer_ApplyClockMode
'
' NOTES
'   This routine intentionally applies the timer mode even when the requested
'   value is already active, because the live-clock timer may need to be
'   resynchronized after form lifecycle changes
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetClockMode"

    Dim OldClockMode        As DP_ClockMode     'Previous clock-mode setting
    Dim SettingsMutated     As Boolean          'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean          'True after settings are saved successfully
    Dim ErrorNumber         As Long             'Captured error number
    Dim ErrorDescription    As String           'Captured error description
    Dim HandlerStep         As String           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate clock mode"

    'Reject unsupported clock modes
        If Not M_Settings_IsValidClockMode(ClockMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ClockMode is unsupported"
        End If

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' EXIT IF SETTING IS UNCHANGED
'------------------------------------------------------------------------------
    'Synchronize the timer and exit when the requested clock mode is already active
        If gDP_ClockMode = ClockMode Then
            M_Timer_ApplyClockMode
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current setting"

    'Capture the current clock mode for rollback
        OldClockMode = gDP_ClockMode

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory setting"

    'Store the requested clock mode
        gDP_ClockMode = ClockMode
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist setting"

    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' APPLY TIMER SIDE EFFECT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply clock mode"

    'Start or stop the live-clock timer according to the active mode
        M_Timer_ApplyClockMode

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_ClockMode = OldClockMode
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Clock-mode setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetSizeMode(ByVal SizeMode As DP_SizeMode)

'
'------------------------------------------------------------------------------
'                           SET SIZE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves the DatePicker size-mode preference
'
' WHY THIS EXISTS
'   The DatePicker supports alternative layout modes, such as normal and compact
'   display. Caller code, demo sheets, Ribbon callbacks, settings panels, and
'   host workbooks need one controlled public entry point to update this setting
'   consistently
'
' INPUTS
'   SizeMode
'     Requested DatePicker size mode
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the requested size mode, ensures current settings are loaded,
'   avoids unnecessary registry writes when the setting is unchanged, updates the
'   in-memory size-mode setting when required, persists the updated settings, and
'   refreshes the loaded DatePicker form when applicable
'
' ERROR POLICY
'   Raises a descriptive runtime error if SizeMode is unsupported, settings
'   cannot be loaded, settings cannot be saved, or the loaded DatePicker form
'   cannot be refreshed
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory size mode is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_IsValidSizeMode
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Size mode is a structural UI setting. If the DatePicker form is already
'   loaded, a simple caption refresh is not sufficient unless
'   M_FormBridge_RefreshSettings also resizes, unloads, or rebuilds the form
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetSizeMode"

    Dim OldSizeMode         As DP_SizeMode      'Previous size-mode setting
    Dim SettingsMutated     As Boolean          'True after in-memory setting is changed
    Dim SettingsPersisted   As Boolean          'True after settings are saved successfully
    Dim ErrorNumber         As Long             'Captured error number
    Dim ErrorDescription    As String           'Captured error description
    Dim HandlerStep         As String           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate size mode"

    'Reject unsupported size modes
        If Not M_Settings_IsValidSizeMode(SizeMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "SizeMode is unsupported"
        End If

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' EXIT IF SETTING IS UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested size mode is already active
        If gDP_SizeMode = SizeMode Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current setting"

    'Capture the current size mode for rollback
        OldSizeMode = gDP_SizeMode

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory setting"

    'Store the requested size mode
        gDP_SizeMode = SizeMode
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist setting"

    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' REFRESH LOADED FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh loaded DatePicker form"

    'Refresh or rebuild the loaded DatePicker form when applicable
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            On Error Resume Next
            gDP_SizeMode = OldSizeMode
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Size-mode setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetHolidayCallback(ByVal HolidayCallbackName As String)

'
'------------------------------------------------------------------------------
'                           SET HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves the DatePicker holiday-callback preference
'
' WHY THIS EXISTS
'   The DatePicker can delegate holiday / non-business-day logic to a caller
'   supplied VBA callback. Caller code, demo sheets, Ribbon callbacks, settings
'   panels, and host workbooks need one controlled public entry point to update
'   that callback name consistently
'
' INPUTS
'   HolidayCallbackName
'     Name of the VBA callback used by the DatePicker holiday policy
'
'     Blank clears the callback and disables callback-based holiday logic
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded, normalizes the supplied callback name,
'   avoids unnecessary registry writes when the setting is unchanged, updates the
'   in-memory callback name when required, persists the updated settings, and
'   refreshes the loaded DatePicker form when applicable
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, settings
'   cannot be saved, or the loaded DatePicker form cannot be refreshed
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory callback name is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   This routine intentionally does not execute or validate the callback by
'   calling it. A callback may belong to a host workbook that is not currently in
'   the expected runtime state
'
'   Callback execution and callback-signature validation belong in the date
'   policy layer, not in the settings persistence layer
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetHolidayCallback"

    Dim NewCallbackName         As String       'Normalized requested callback name
    Dim OldCallbackName         As String       'Previous callback name
    Dim SettingsMutated         As Boolean      'True after in-memory setting is changed
    Dim SettingsPersisted       As Boolean      'True after settings are saved successfully
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description
    Dim HandlerStep             As String       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize callback name"

    'Normalize the supplied callback name
        NewCallbackName = VBA.Trim$(HolidayCallbackName)

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' EXIT IF SETTING IS UNCHANGED
'------------------------------------------------------------------------------
    'Exit when the requested callback name is already active
        If VBA.StrComp(gDP_HolidayCallbackName, NewCallbackName, vbBinaryCompare) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current setting"

    'Capture the current callback name for rollback
        OldCallbackName = gDP_HolidayCallbackName

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory setting"

    'Store the normalized callback name
        gDP_HolidayCallbackName = NewCallbackName
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist setting"

    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' REFRESH LOADED FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh loaded DatePicker form"

    'Refresh settings-dependent DatePicker state when the form is loaded
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore the previous in-memory setting if persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous callback name
                gDP_HolidayCallbackName = OldCallbackName
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Holiday-callback setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetShowRightClick(ByVal ShowRightClick As Boolean)

'
'------------------------------------------------------------------------------
'                           SET SHOW RIGHT-CLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether the DatePicker is available from the Excel right-click
'   context menu
'
' WHY THIS EXISTS
'   Caller code, demo sheets, Ribbon callbacks, settings panels, and host
'   workbooks need one controlled public entry point to update right-click menu
'   integration without mutating public DatePicker state inconsistently
'
' INPUTS
'   ShowRightClick
'     True to enable right-click menu integration
'     False to disable right-click menu integration
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded, applies the requested right-click menu
'   setting, keeps keyboard shortcut access enabled when both visual entry
'   points are disabled, persists the updated settings when needed, synchronizes
'   the Excel right-click menu, and updates keyboard shortcut integration when it
'   was changed as a fallback
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, settings
'   cannot be saved, the right-click menu cannot be synchronized, or keyboard
'   shortcut integration cannot be updated
'
'   If persistence fails after in-memory settings were changed, the previous
'   in-memory settings are restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'
' NOTES
'   Disabling both right-click menu access and in-grid icon access would leave
'   the user without a visible/manual DatePicker entry point unless the keyboard
'   shortcut remains enabled
'
'   The right-click menu is synchronized even when the persisted setting is
'   unchanged, because Excel context menus may be reset by Excel, another add-in,
'   workbook activation, or application-level cleanup
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetShowRightClick"

    Dim OldShowRightClick       As Boolean      'Previous right-click menu setting
    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting

    Dim RightClickChanged       As Boolean      'True when right-click menu setting changed
    Dim KeyboardChanged         As Boolean      'True when keyboard shortcut fallback changed
    Dim SettingsMutated         As Boolean      'True after in-memory settings are changed
    Dim SettingsPersisted       As Boolean      'True after settings are saved successfully

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description
    Dim HandlerStep             As String       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current settings"

    'Capture the current right-click menu setting
        OldShowRightClick = gDP_ShowRightClick
    'Capture the current keyboard shortcut setting
        OldEnableKeyboard = gDP_EnableKeyboardShortcut
    'Initialize the resolved keyboard shortcut setting
        NewEnableKeyboard = gDP_EnableKeyboardShortcut

'------------------------------------------------------------------------------
' PROTECT MANUAL ACCESS PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve access fallback"

    'Keep keyboard access enabled when both visible entry points are disabled
        If Not ShowRightClick Then
            If Not gDP_ShowGridIcon Then
                If Not NewEnableKeyboard Then
                    NewEnableKeyboard = True
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE CHANGE FLAGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flags"

    'Resolve whether the right-click menu setting changed
        RightClickChanged = (ShowRightClick <> OldShowRightClick)
    'Resolve whether the keyboard shortcut setting changed
        KeyboardChanged = (NewEnableKeyboard <> OldEnableKeyboard)

'------------------------------------------------------------------------------
' SYNCHRONIZE ONLY WHEN SETTINGS ARE UNCHANGED
'------------------------------------------------------------------------------
    'Synchronize the right-click menu and exit when no persisted setting changed
        If Not RightClickChanged Then
            If Not KeyboardChanged Then
                'Track the current handler step
                    HandlerStep = "Synchronize right-click menu"
                'Synchronize right-click menus with the current setting
                    M_ContextMenu_Update
                'Exit because no registry write is required
                    Exit Sub
            End If
        End If

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory settings"

    'Store the requested right-click menu setting
        gDP_ShowRightClick = ShowRightClick
    'Store the resolved keyboard shortcut setting
        gDP_EnableKeyboardShortcut = NewEnableKeyboard
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist settings"

    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' SYNCHRONIZE RIGHT-CLICK MENU
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize right-click menu"

    'Synchronize right-click menus with the saved setting
        M_ContextMenu_Update

'------------------------------------------------------------------------------
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"

    'Synchronize keyboard shortcut integration only when the fallback changed
        If KeyboardChanged Then
            M_KeyboardShortcut_Update
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Restore previous in-memory settings when persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous right-click menu setting
                gDP_ShowRightClick = OldShowRightClick
            'Restore the previous keyboard shortcut setting
                gDP_EnableKeyboardShortcut = OldEnableKeyboard
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Right-click menu setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetShowGridIcon(ByVal ShowGridIcon As Boolean)

'
'------------------------------------------------------------------------------
'                           SET SHOW GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether the DatePicker in-grid worksheet icon is enabled
'
' WHY THIS EXISTS
'   Caller code, demo sheets, Ribbon callbacks, settings panels, and host
'   workbooks need one controlled public entry point to update worksheet icon
'   integration without mutating public DatePicker state inconsistently
'
' INPUTS
'   ShowGridIcon
'     True to enable in-grid icon integration
'     False to disable in-grid icon integration
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded, applies the requested in-grid icon
'   setting, keeps keyboard shortcut access enabled when both visual entry
'   points are disabled, persists the updated settings when needed, synchronizes
'   keyboard shortcut integration when it was changed as a fallback, and removes
'   any stale in-grid icon when the feature is disabled
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, settings
'   cannot be saved, keyboard shortcut integration cannot be synchronized, or a
'   stale in-grid icon cannot be removed
'
'   If persistence fails after in-memory settings were changed, the previous
'   in-memory settings are restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_KeyboardShortcut_Update
'   M_GridIcon_Remove
'
' NOTES
'   Disabling both right-click menu access and in-grid icon access would leave
'   the user without a visible/manual DatePicker entry point unless the keyboard
'   shortcut remains enabled
'
'   The stale grid icon is removed even when the setting was already disabled,
'   because worksheet shapes may survive prior failures, workbook activation, or
'   partial UI cleanup
'
'   Enabling the in-grid icon is persisted immediately. Actual icon display is
'   expected to occur through the manager selection / context refresh path
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetShowGridIcon"

    Dim OldShowGridIcon         As Boolean      'Previous in-grid icon setting
    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting

    Dim GridIconChanged         As Boolean      'True when in-grid icon setting changed
    Dim KeyboardChanged         As Boolean      'True when keyboard shortcut fallback changed
    Dim SettingsMutated         As Boolean      'True after in-memory settings are changed
    Dim SettingsPersisted       As Boolean      'True after settings are saved successfully

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description
    Dim HandlerStep             As String       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"

    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current settings"

    'Capture the current in-grid icon setting
        OldShowGridIcon = gDP_ShowGridIcon
    'Capture the current keyboard shortcut setting
        OldEnableKeyboard = gDP_EnableKeyboardShortcut
    'Initialize the resolved keyboard shortcut setting
        NewEnableKeyboard = gDP_EnableKeyboardShortcut

'------------------------------------------------------------------------------
' PROTECT MANUAL ACCESS PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve access fallback"

    'Keep keyboard access enabled when both visible entry points are disabled
        If Not ShowGridIcon Then
            If Not gDP_ShowRightClick Then
                If Not NewEnableKeyboard Then
                    NewEnableKeyboard = True
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE CHANGE FLAGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flags"

    'Resolve whether the in-grid icon setting changed
        GridIconChanged = (ShowGridIcon <> OldShowGridIcon)
    'Resolve whether the keyboard shortcut setting changed
        KeyboardChanged = (NewEnableKeyboard <> OldEnableKeyboard)

'------------------------------------------------------------------------------
' CLEAN UP ONLY WHEN SETTINGS ARE UNCHANGED
'------------------------------------------------------------------------------
    'Remove stale grid icon and exit when no persisted setting changed
        If Not GridIconChanged Then
            If Not KeyboardChanged Then
                'Remove any stale in-grid icon when the feature is disabled
                    If Not ShowGridIcon Then M_GridIcon_Remove
                'Exit because no registry write is required
                    Exit Sub
            End If
        End If

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory settings"

    'Store the requested in-grid icon setting
        gDP_ShowGridIcon = ShowGridIcon
    'Store the resolved keyboard shortcut setting
        gDP_EnableKeyboardShortcut = NewEnableKeyboard
    'Mark settings as mutated
        SettingsMutated = True

'------------------------------------------------------------------------------
' PERSIST SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Persist settings"

    'Persist the updated settings
        M_Settings_Save
    'Mark settings as persisted
        SettingsPersisted = True

'------------------------------------------------------------------------------
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"

    'Synchronize keyboard shortcut integration only when the fallback changed
        If KeyboardChanged Then M_KeyboardShortcut_Update

'------------------------------------------------------------------------------
' REMOVE STALE GRID ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Remove stale grid icon"

    'Remove any stale in-grid icon when the feature is disabled
        If Not gDP_ShowGridIcon Then M_GridIcon_Remove

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Restore previous in-memory settings when persistence failed after mutation
        If SettingsMutated And Not SettingsPersisted Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous in-grid icon setting
                gDP_ShowGridIcon = OldShowGridIcon
            'Restore the previous keyboard shortcut setting
                gDP_EnableKeyboardShortcut = OldEnableKeyboard
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "In-grid icon setting update failed: " & ErrorDescription

End Sub
Public Function M_Settings_GetFirstDayOfWeek() As Long

'
'------------------------------------------------------------------------------
'                           GET FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current DatePicker first-day-of-week setting
'
' WHY THIS EXISTS
'   External callers, demo sheets, settings panels, and host workbooks should
'   read DatePicker settings through controlled accessors instead of depending
'   directly on mutable public module state
'
' INPUTS
'   None
'
' RETURNS
'   Current first-day-of-week setting:
'     - vbSunday
'     - vbMonday
'
' BEHAVIOR
'   Ensures settings are loaded, validates the in-memory first-day setting, and
'   returns the current value
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or if the
'   loaded in-memory first-day setting is unsupported
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   gDP_FirstDayOfWeek
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_GetFirstDayOfWeek"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the current value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' VALIDATE SETTING
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' RETURN SETTING
'------------------------------------------------------------------------------
    'Return the current first-day setting
        M_Settings_GetFirstDayOfWeek = gDP_FirstDayOfWeek

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
            "First-day-of-week setting retrieval failed: " & Err.Description

End Function
Public Function M_Settings_GetUseLocalDayNames() As Boolean

'
'------------------------------------------------------------------------------
'                           GET USE LOCAL DAY NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker uses local day and month captions
'
' WHY THIS EXISTS
'   External callers, demo sheets, settings panels, and host workbooks should
'   read DatePicker settings through controlled accessors instead of depending
'   directly on mutable public module state
'
' INPUTS
'   None
'
' RETURNS
'   True when the DatePicker uses local day and month names
'
'   False when the DatePicker uses fixed English day and month names
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current local-name setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_UseLocalNames
'
' NOTES
'   The historical procedure name is preserved for compatibility
'
'   Although the name refers to day names, the setting also controls month
'   captions where applicable
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_GetUseLocalDayNames"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the current value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' RETURN SETTING
'------------------------------------------------------------------------------
    'Return the current local-name setting
        M_Settings_GetUseLocalDayNames = gDP_UseLocalNames

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
            "Use-local-day-names setting retrieval failed: " & Err.Description

End Function
Public Function M_Settings_GetFirstDayOfWeekText() As String

'
'------------------------------------------------------------------------------
'                       GET FIRST DAY OF WEEK TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current DatePicker first-day-of-week setting as text
'
' WHY THIS EXISTS
'   External callers, demo sheets, settings panels, and host workbooks may need a
'   worksheet-friendly textual representation of the first-day setting without
'   reading mutable public module state directly
'
' INPUTS
'   None
'
' RETURNS
'   Text representation of the current first-day-of-week setting
'
' BEHAVIOR
'   Ensures settings are loaded, validates the in-memory first-day setting, and
'   returns the corresponding text value
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, if the
'   first-day setting is unsupported, or if the value cannot be converted to text
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   M_Settings_FirstDayOfWeekToText
'   gDP_FirstDayOfWeek
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_GetFirstDayOfWeekText"
    
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the current value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' VALIDATE SETTING
'------------------------------------------------------------------------------
    'Reject unsupported first-day settings
        If Not M_Settings_IsValidFirstDayOfWeek(gDP_FirstDayOfWeek) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "gDP_FirstDayOfWeek must be vbSunday or vbMonday"
        End If

'------------------------------------------------------------------------------
' RETURN SETTING
'------------------------------------------------------------------------------
    'Return the current first-day setting as text
        M_Settings_GetFirstDayOfWeekText = _
            M_Settings_FirstDayOfWeekToText(gDP_FirstDayOfWeek)

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
            "First-day-of-week text retrieval failed: " & Err.Description

End Function
Public Function M_Settings_FirstDayOfWeekToText(ByVal FirstDayOfWeek As Long) As String

'
'------------------------------------------------------------------------------
'                       CONVERT FIRST DAY OF WEEK TO TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts a supported DatePicker first-day-of-week value to stable text
'
' WHY THIS EXISTS
'   Settings sheets, diagnostics, demos, and public accessors may need a stable
'   textual representation of the first-day setting without relying on numeric
'   VBA weekday constants
'
' INPUTS
'   FirstDayOfWeek
'     First-day-of-week value to convert
'
' RETURNS
'   "vbSunday" when FirstDayOfWeek is vbSunday
'
'   "vbMonday" when FirstDayOfWeek is vbMonday
'
' BEHAVIOR
'   Maps supported first-day values to their canonical text representation
'
' ERROR POLICY
'   Raises a descriptive runtime error if FirstDayOfWeek is unsupported
'
' DEPENDENCIES
'   VBA weekday constants
'
' NOTES
'   This is a pure conversion helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
'   Returned values are intentionally aligned with
'   M_Settings_TryParseFirstDayOfWeek
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Settings_FirstDayOfWeekToText"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETURN TEXT VALUE
'------------------------------------------------------------------------------
    'Return the canonical text value for supported first-day settings
        Select Case FirstDayOfWeek
            Case vbSunday
                M_Settings_FirstDayOfWeekToText = "vbSunday"
            Case vbMonday
                M_Settings_FirstDayOfWeekToText = "vbMonday"
            Case Else
                Err.Raise vbObjectError + 513, PROC_NAME, _
                    "FirstDayOfWeek must be vbSunday or vbMonday"
        End Select

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
            "First-day-of-week text conversion failed: " & Err.Description

End Function
Public Function M_Settings_IsValidFirstDayOfWeek(ByVal FirstDayOfWeek As Long) As Boolean

'
'------------------------------------------------------------------------------
'                       VALIDATE FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a supplied first-day-of-week value is supported by the
'   DatePicker
'
' WHY THIS EXISTS
'   The DatePicker intentionally supports only two first-day policies:
'     - Sunday-start calendars
'     - Monday-start calendars
'
'   Centralizing the validation keeps settings loading, saving, parsing, and
'   public setter routines aligned to the same policy
'
' INPUTS
'   FirstDayOfWeek
'     First-day-of-week value to validate
'
' RETURNS
'   True when FirstDayOfWeek is vbSunday or vbMonday
'
'   False for all other values
'
' BEHAVIOR
'   Compares the supplied value against the supported VBA weekday constants
'
' ERROR POLICY
'   Does not raise errors
'
'   Unsupported values return False so callers can decide whether to default,
'   reject, or raise a higher-level error
'
' DEPENDENCIES
'   VBA weekday constants
'
' NOTES
'   This is a pure validation helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALIDATION RESULT
'------------------------------------------------------------------------------
    'Return whether the first-day value is supported
        Select Case FirstDayOfWeek
            Case vbSunday, vbMonday
                M_Settings_IsValidFirstDayOfWeek = True
            Case Else
                M_Settings_IsValidFirstDayOfWeek = False
        End Select

End Function
Private Function M_Settings_IsValidClockMode(ByVal ClockMode As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           VALIDATE CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a supplied DatePicker clock mode is supported
'
' WHY THIS EXISTS
'   The DatePicker supports a controlled set of footer-clock behaviors:
'     - static time caption
'     - live time caption
'
'   Centralizing the validation keeps settings loading, saving, parsing, and
'   public setter routines aligned to the same clock-mode policy
'
' INPUTS
'   ClockMode
'     Clock-mode value to validate
'
' RETURNS
'   True when ClockMode is DP_ClockMode_Static or DP_ClockMode_Live
'
'   False for all other values
'
' BEHAVIOR
'   Compares the supplied value against the supported DatePicker clock-mode
'   constants
'
' ERROR POLICY
'   Does not raise errors
'
'   Unsupported values return False so callers can decide whether to default,
'   reject, or raise a higher-level error
'
' DEPENDENCIES
'   DP_ClockMode_Static
'   DP_ClockMode_Live
'
' NOTES
'   This is a pure validation helper
'
'   This routine does not load settings, save settings, read the registry, start
'   timers, stop timers, or mutate DatePicker state
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALIDATION RESULT
'------------------------------------------------------------------------------
    'Return whether the clock mode is supported
        Select Case ClockMode
            Case DP_ClockMode_Static, DP_ClockMode_Live
                M_Settings_IsValidClockMode = True
            Case Else
                M_Settings_IsValidClockMode = False
        End Select

End Function
Private Function M_Settings_IsValidSizeMode(ByVal SizeMode As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           VALIDATE SIZE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a supplied DatePicker size mode is supported
'
' WHY THIS EXISTS
'   The DatePicker supports a controlled set of layout modes:
'     - normal layout
'     - compact layout
'
'   Centralizing the validation keeps settings loading, saving, parsing, and
'   public setter routines aligned to the same size-mode policy
'
' INPUTS
'   SizeMode
'     Size-mode value to validate
'
' RETURNS
'   True when SizeMode is DP_SizeMode_Normal or DP_SizeMode_Compact
'
'   False for all other values
'
' BEHAVIOR
'   Compares the supplied value against the supported DatePicker size-mode
'   constants
'
' ERROR POLICY
'   Does not raise errors
'
'   Unsupported values return False so callers can decide whether to default,
'   reject, or raise a higher-level error
'
' DEPENDENCIES
'   DP_SizeMode_Normal
'   DP_SizeMode_Compact
'
' NOTES
'   This is a pure validation helper
'
'   This routine does not load settings, save settings, read the registry,
'   resize the form, rebuild controls, or mutate DatePicker state
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALIDATION RESULT
'------------------------------------------------------------------------------
    'Return whether the size mode is supported
        Select Case SizeMode
            Case DP_SizeMode_Normal, DP_SizeMode_Compact
                M_Settings_IsValidSizeMode = True
            Case Else
                M_Settings_IsValidSizeMode = False
        End Select

End Function
Private Function M_Settings_TryParseBoolean( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY PARSE BOOLEAN SETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted Boolean DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be normalized before being
'   trusted as Boolean settings
'
'   A dedicated parser keeps settings load logic clean, defensive, and aligned
'   to one common interpretation of stored Boolean values
'
' INPUTS
'   RawValue
'     Raw string value read from persisted settings
'
'   ParsedValue
'     Output Boolean value populated when parsing succeeds
'
' RETURNS
'   True when RawValue is recognized as a supported Boolean representation
'
'   False when RawValue is blank, unsupported, or cannot be parsed safely
'
' BEHAVIOR
'   Normalizes RawValue, recognizes canonical persisted values and common
'   textual Boolean forms, assigns ParsedValue deterministically, and returns the
'   parse result
'
' ERROR POLICY
'   Does not raise errors
'
'   Any unexpected parsing failure returns False and resets ParsedValue to False
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is a pure parsing helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
'   Canonical persisted Boolean values should remain 1 and 0
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedValue     As String   'Normalized setting value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe parsing
        On Error GoTo ParseFail
    'Initialize the parsed value to a deterministic failure state
        ParsedValue = False
    'Initialize the function result to parsing failure
        M_Settings_TryParseBoolean = False

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the raw setting value
        NormalizedValue = VBA.UCase$(VBA.Trim$(RawValue))
    'Return False for blank values
        If VBA.LenB(NormalizedValue) = 0 Then Exit Function

'------------------------------------------------------------------------------
' PARSE BOOLEAN VALUE
'------------------------------------------------------------------------------
    'Parse the normalized Boolean value
        Select Case NormalizedValue
            Case "1", "-1", "TRUE", "T", "YES", "Y", "ON"
                'Store the parsed Boolean value
                    ParsedValue = True
                'Return successful parsing
                    M_Settings_TryParseBoolean = True

            Case "0", "FALSE", "F", "NO", "N", "OFF"
                'Store the parsed Boolean value
                    ParsedValue = False
                'Return successful parsing
                    M_Settings_TryParseBoolean = True

            Case Else
                'Reset the parsed value to a deterministic failure state
                    ParsedValue = False
                'Return parsing failure
                    M_Settings_TryParseBoolean = False
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Function

'------------------------------------------------------------------------------
' PARSE FAIL
'------------------------------------------------------------------------------
ParseFail:
    'Suppress secondary cleanup errors
        On Error Resume Next
    'Reset the parsed value to a deterministic failure state
        ParsedValue = False
    'Return parsing failure
        M_Settings_TryParseBoolean = False
    'Clear the parsing error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function
Private Function M_Settings_TryParseFirstDayOfWeek( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Long) As Boolean

'
'------------------------------------------------------------------------------
'                       TRY PARSE FIRST DAY OF WEEK SETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted first-day-of-week DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be normalized before being
'   trusted as DatePicker settings
'
'   A dedicated parser keeps settings load logic clean, defensive, and aligned
'   to one common interpretation of supported first-day values
'
' INPUTS
'   RawValue
'     Raw string value read from persisted settings
'
'   ParsedValue
'     Output first-day-of-week value populated when parsing succeeds
'
' RETURNS
'   True when RawValue resolves to vbSunday or vbMonday
'
'   False when RawValue is blank, unsupported, or cannot be parsed safely
'
' BEHAVIOR
'   Normalizes RawValue, recognizes canonical persisted values, VBA-style
'   constant names, English weekday names, and common Italian weekday names,
'   assigns ParsedValue deterministically, and returns the parse result
'
' ERROR POLICY
'   Does not raise errors
'
'   Any unexpected parsing failure returns False and resets ParsedValue to zero
'
' DEPENDENCIES
'   VBA weekday constants
'
' NOTES
'   This is a pure parsing helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
'   Canonical persisted first-day values should remain 1 and 2
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedValue     As String   'Normalized first-day setting value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe parsing
        On Error GoTo ParseFail
    'Initialize the parsed value to a deterministic failure state
        ParsedValue = 0
    'Initialize the function result to parsing failure
        M_Settings_TryParseFirstDayOfWeek = False

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the raw first-day value
        NormalizedValue = VBA.UCase$(VBA.Trim$(RawValue))
    'Return False for blank values
        If VBA.LenB(NormalizedValue) = 0 Then Exit Function

'------------------------------------------------------------------------------
' PARSE FIRST-DAY VALUE
'------------------------------------------------------------------------------
    'Parse the normalized first-day value
        Select Case NormalizedValue
            Case "1", "VBSUNDAY", "SUNDAY", "SUN"
                'Store the parsed first-day value
                    ParsedValue = vbSunday
                'Return successful parsing
                    M_Settings_TryParseFirstDayOfWeek = True

            Case "2", "VBMONDAY", "MONDAY", "MON"
                'Store the parsed first-day value
                    ParsedValue = vbMonday
                'Return successful parsing
                    M_Settings_TryParseFirstDayOfWeek = True

            Case Else
                'Reset the parsed value to a deterministic failure state
                    ParsedValue = 0
                'Return parsing failure
                    M_Settings_TryParseFirstDayOfWeek = False
        End Select

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Function

'------------------------------------------------------------------------------
' PARSE FAIL
'------------------------------------------------------------------------------
ParseFail:
    'Suppress secondary cleanup errors
        On Error Resume Next
    'Reset the parsed value to a deterministic failure state
        ParsedValue = 0
    'Return parsing failure
        M_Settings_TryParseFirstDayOfWeek = False
    'Clear the parsing error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function
Private Function M_Settings_TryParseLong( _
    ByVal RawValue As String, _
    ByRef ParsedValue As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY PARSE LONG SETTING
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to parse a persisted Long DatePicker setting value
'
' WHY THIS EXISTS
'   Persisted registry values are strings and must be normalized before being
'   trusted as numeric DatePicker settings
'
'   A dedicated strict parser avoids accepting ambiguous numeric formats such as
'   decimals, currency values, scientific notation, hexadecimal notation, or
'   locale-dependent numeric strings
'
' INPUTS
'   RawValue
'     Raw string value read from persisted settings
'
'   ParsedValue
'     Output Long value populated when parsing succeeds
'
' RETURNS
'   True when RawValue is a strict base-10 integer inside the VBA Long range
'
'   False when RawValue is blank, unsupported, out of range, or cannot be parsed
'   safely
'
' BEHAVIOR
'   Trims RawValue, accepts an optional leading sign, rejects non-digit content,
'   validates the Long range while parsing, assigns ParsedValue deterministically,
'   and returns the parse result
'
' ERROR POLICY
'   Does not raise errors
'
'   Any unexpected parsing failure returns False and resets ParsedValue to zero
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is a pure parsing helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
'   This routine intentionally does not use IsNumeric because IsNumeric accepts
'   formats that are not appropriate for persisted integer settings
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const LONG_MAX_VALUE         As Double = 2147483647#    'Maximum VBA Long value
    Const LONG_MIN_ABS_VALUE     As Double = 2147483648#    'Absolute minimum VBA Long value

    Dim NormalizedValue          As String                  'Trimmed raw numeric value
    Dim NumericText              As String                  'Numeric portion without sign
    Dim FirstCharacter           As String                  'First character of the normalized value
    Dim CurrentCharacter         As String                  'Current parsed digit
    Dim CharacterIndex           As Long                    'Character loop index
    Dim SignMultiplier           As Long                    'Parsed sign multiplier
    Dim AccumulatedValue         As Double                  'Accumulated absolute numeric value
    Dim SignedValue              As Double                  'Signed numeric value before Long conversion

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe parsing
        On Error GoTo ParseFail
    'Initialize the parsed value to a deterministic failure state
        ParsedValue = 0
    'Initialize the function result to parsing failure
        M_Settings_TryParseLong = False

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the raw numeric value
        NormalizedValue = VBA.Trim$(RawValue)
    'Return False for blank values
        If VBA.LenB(NormalizedValue) = 0 Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' RESOLVE SIGN
'------------------------------------------------------------------------------
    'Initialize the sign multiplier
        SignMultiplier = 1
    'Read the first character
        FirstCharacter = VBA.Left$(NormalizedValue, 1)
    'Handle an explicit positive sign
        If FirstCharacter = "+" Then
            NumericText = VBA.Mid$(NormalizedValue, 2)
    'Handle an explicit negative sign
        ElseIf FirstCharacter = "-" Then
            SignMultiplier = -1
            NumericText = VBA.Mid$(NormalizedValue, 2)
    'Use the full value when no explicit sign is supplied
        Else
            NumericText = NormalizedValue
        End If
    'Return False when only a sign was supplied
        If VBA.LenB(NumericText) = 0 Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE DIGITS
'------------------------------------------------------------------------------
    'Reject any non-digit character
        If NumericText Like "*[!0-9]*" Then Exit Function

'------------------------------------------------------------------------------
' PARSE NUMERIC VALUE
'------------------------------------------------------------------------------
    'Loop through the numeric characters
        For CharacterIndex = 1 To VBA.Len(NumericText)
            'Read the current digit
                CurrentCharacter = VBA.Mid$(NumericText, CharacterIndex, 1)
            'Accumulate the absolute numeric value
                AccumulatedValue = (AccumulatedValue * 10#) + VBA.CDbl(CurrentCharacter)
            'Reject positive values outside the Long range
                If SignMultiplier = 1 Then
                    If AccumulatedValue > LONG_MAX_VALUE Then Exit Function
                End If
            'Reject negative values outside the Long range
                If SignMultiplier = -1 Then
                    If AccumulatedValue > LONG_MIN_ABS_VALUE Then Exit Function
                End If
        Next CharacterIndex

'------------------------------------------------------------------------------
' CONVERT TO LONG
'------------------------------------------------------------------------------
    'Apply the parsed sign
        SignedValue = AccumulatedValue * SignMultiplier
    'Store the parsed Long value
        ParsedValue = VBA.CLng(SignedValue)
    'Return successful parsing
        M_Settings_TryParseLong = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Function

'------------------------------------------------------------------------------
' PARSE FAIL
'------------------------------------------------------------------------------
ParseFail:
    'Suppress secondary cleanup errors
        On Error Resume Next
    'Reset the parsed value to a deterministic failure state
        ParsedValue = 0
    'Return parsing failure
        M_Settings_TryParseLong = False
    'Clear the parsing error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function
Private Function M_Settings_BooleanToStorageValue(ByVal Value As Boolean) As String

'
'------------------------------------------------------------------------------
'                       CONVERT BOOLEAN TO STORAGE VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Converts a Boolean DatePicker setting to its canonical persisted string value
'
' WHY THIS EXISTS
'   VBA represents True internally as -1, but persisted settings should use a
'   stable and language-independent representation
'
'   Saving Boolean settings as 1 / 0 keeps registry values compact, predictable,
'   and aligned with M_Settings_TryParseBoolean
'
' INPUTS
'   Value
'     Boolean value to convert
'
' RETURNS
'   "1" when Value is True
'
'   "0" when Value is False
'
' BEHAVIOR
'   Maps the supplied Boolean value to the DatePicker canonical storage format
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is a pure conversion helper
'
'   This routine does not load settings, save settings, read the registry, or
'   mutate DatePicker state
'
'   Do not replace this with CStr(Value) or CStr(CLng(Value)), because VBA would
'   persist True as "True" or "-1" instead of the canonical "1"
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STORAGE VALUE
'------------------------------------------------------------------------------
    'Return the canonical storage value for True
        If Value Then
            M_Settings_BooleanToStorageValue = "1"
    'Return the canonical storage value for False
        Else
            M_Settings_BooleanToStorageValue = "0"
        End If

End Function



Public Sub M_Picker_EnsureManager()

'
'==============================================================================
'                          ENSURE DATEPICKER MANAGER
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures the global DatePicker manager object is instantiated and hooked
'
' WHY THIS EXISTS
'   The DatePicker manager owns the Excel Application event hooks used to
'   coordinate context-sensitive DatePicker behavior
'
'   After VBA reset, project import, workbook reload, or partial teardown, the
'   global manager reference may be missing or may exist without a valid
'   Application event hook
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures persisted settings are loaded, creates the manager when missing, and
'   recreates it when the existing manager is not hooked
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or if the
'   manager cannot be instantiated / re-instantiated
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   cDatePickerManager
'   gDP_Manager
'
' NOTES
'   This routine is the canonical manager / controller bootstrapper
'
'   The Is_Hooked check prevents a stale manager object from silently disabling
'   selection-change behavior
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Picker_EnsureManager"

    Dim ManagerNeedsCreate      As Boolean          'True when manager must be created
    Dim ManagerNeedsRecreate    As Boolean          'True when manager exists but is not hooked
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure persisted DatePicker settings are available before manager startup
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' RESOLVE MANAGER STATE
'------------------------------------------------------------------------------
    'Detect missing manager
        ManagerNeedsCreate = (gDP_Manager Is Nothing)
    'Check hook state only when a manager already exists
        If Not ManagerNeedsCreate Then
            ManagerNeedsRecreate = Not gDP_Manager.Is_Hooked
        End If

'------------------------------------------------------------------------------
' RECREATE STALE MANAGER IF NEEDED
'------------------------------------------------------------------------------
    'Release stale manager when the Application event hook is not active
        If ManagerNeedsRecreate Then
            Set gDP_Manager = Nothing
            ManagerNeedsCreate = True
        End If

'------------------------------------------------------------------------------
' ENSURE MANAGER
'------------------------------------------------------------------------------
    'Instantiate the manager when missing or after stale-manager release
        If ManagerNeedsCreate Then
            Set gDP_Manager = New cDatePickerManager
        End If

'------------------------------------------------------------------------------
' VALIDATE MANAGER
'------------------------------------------------------------------------------
    'Reject failed manager creation
        If gDP_Manager Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "DatePicker manager was not created"
        End If
    'Reject manager creation without Application event hook
        If Not gDP_Manager.Is_Hooked Then
            Err.Raise vbObjectError + 514, PROC_NAME, "DatePicker manager was created but Application events are not hooked"
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "DatePicker manager initialization failed: " & ErrorDescription

End Sub
'

Public Sub DP_Start()

'
'==============================================================================
'                           START DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the DatePicker manager and immediately refreshes the current UI state
'
' WHY THIS EXISTS
'   The manager is event-driven. After workbook open, VBA reset, code import, or
'   add-in reload, the manager must be explicitly bootstrapped before Excel
'   selection-change events can move or remove the in-grid icon
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the manager is alive and hooked, then evaluates the current ActiveCell
'   so stale icons are cleaned and the correct in-grid icon state is shown
'
' ERROR POLICY
'   Raises a descriptive runtime error if startup fails
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   gDP_Manager.Handle_SelectionChange
'
' NOTES
'   Call this from Workbook_Open, Auto_Open, add-in startup, or manually after
'   importing the project into a workbook
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "DP_Start"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ENSURE MANAGER
'------------------------------------------------------------------------------
    'Ensure the DatePicker manager exists and Application events are hooked
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' REFRESH CURRENT UI
'------------------------------------------------------------------------------
    'Force one initial current-context refresh
        gDP_Manager.Handle_SelectionChange

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive startup error
        Err.Raise Err.Number, PROC_NAME, "DatePicker startup failed: " & Err.Description

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
'   ActiveCell contains a valid date
'
'   Otherwise, it should open on today's date
'
'   The form is explicitly loaded and positioned before it is shown so the first
'   visible paint occurs close to the mouse instead of flashing briefly in the
'   top-left corner of the screen
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures DatePicker infrastructure is available, resolves the initial display
'   date from ActiveCell when possible, unloads any existing DatePicker instance,
'   stores bridge state for the next form instance, loads UF_DatePicker, positions
'   the loaded form close to the mouse, shows it modelessly, and applies one
'   final post-show positioning correction
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker cannot be prepared,
'   loaded, resolved, or shown
'
'   Mouse positioning is best-effort and does not prevent the picker from opening
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   M_FormBridge_UnloadLoadedPicker
'   M_FormBridge_GetLoadedForm
'   M_Window_MoveFormToMouse
'   UF_DatePicker
'
' NOTES
'   Load UF_DatePicker is intentional
'
'   Loading the form before Show runs UserForm_Initialize while the form is still
'   hidden. This allows the form to be positioned before the first visible paint
'
'   The post-show positioning call is retained as a safety correction for host
'   environments where the native UserForm window position is finalized only
'   after Show
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "DP_Show"    'Current procedure name

    Const FORM_MOUSE_OFFSET_XPX     As Long = 10             'Mouse-position X offset in pixels
    Const FORM_MOUSE_OFFSET_YPX     As Long = 0              'Mouse-position Y offset in pixels
    Const FORM_CENTER_ON_MOUSE      As Boolean = False       'True to center form on mouse

    Dim CellValue                   As Variant               'ActiveCell value snapshot
    Dim InitialDate                 As Date                  'Resolved initial display date
    Dim HasCellDate                 As Boolean               'True when ActiveCell contains a usable date
    Dim HasActiveCellValue          As Boolean               'True when ActiveCell value was read successfully
    Dim LoadedForm                  As Object                'Loaded DatePicker form instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Ensure DatePicker settings and manager infrastructure are available
        M_Picker_EnsureManager
    'Default the initial date to today
        InitialDate = VBA.Date
    'Default selected-date availability to False
        HasCellDate = False
    'Default ActiveCell value availability to False
        HasActiveCellValue = False

'------------------------------------------------------------------------------
' READ ACTIVE CELL VALUE
'------------------------------------------------------------------------------
    'Suppress ActiveCell access errors
        On Error Resume Next
    'Read ActiveCell value safely
        CellValue = Application.ActiveCell.Value
    'Store whether ActiveCell value was read successfully
        HasActiveCellValue = (Err.Number = 0)
    'Clear any suppressed ActiveCell access error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE INITIAL DATE FROM ACTIVE CELL
'------------------------------------------------------------------------------
    'Use ActiveCell only when a value was read successfully
        If HasActiveCellValue Then
            'Ignore Excel error values
                If Not IsError(CellValue) Then
                    'Evaluate only date-like values
                        If VBA.IsDate(CellValue) Then
                            'Suppress conversion errors for unusual date-like values
                                On Error Resume Next
                            'Resolve the ActiveCell date without time
                                InitialDate = VBA.DateValue(VBA.CDate(CellValue))
                            'Store whether the conversion succeeded
                                HasCellDate = (Err.Number = 0)
                            'Clear any suppressed conversion error
                                Err.Clear
                            'Restore controlled error handling
                                On Error GoTo ErrorHandler
                        End If
                End If
        End If

    'Restore the default initial date when no valid ActiveCell date was resolved
        If Not HasCellDate Then InitialDate = VBA.Date

'------------------------------------------------------------------------------
' RESET EXISTING FORM INSTANCE
'------------------------------------------------------------------------------
    'Unload any existing DatePicker instance so Initialize runs again
        M_FormBridge_UnloadLoadedPicker

'------------------------------------------------------------------------------
' STORE FORM BRIDGE STATE
'------------------------------------------------------------------------------
    'Store the initial date for the next UF_DatePicker instance
        gDP_InitialDate = InitialDate
    'Mark the initial date as available
        gDP_HasInitialDate = True
    'Store selected-date state when ActiveCell contains a valid date
        If HasCellDate Then
            'Store the selected date
                gDP_SelectedDate = InitialDate
            'Mark the selected date as available
                gDP_HasSelectedDate = True
        Else
            'Clear the selected date
                gDP_SelectedDate = 0
            'Mark the selected date as unavailable
                gDP_HasSelectedDate = False
        End If

'------------------------------------------------------------------------------
' LOAD FORM INSTANCE
'------------------------------------------------------------------------------
    'Load the DatePicker form while it is still hidden
        Load UF_DatePicker
    'Resolve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Fallback to the default instance if the bridge did not resolve it
        If LoadedForm Is Nothing Then Set LoadedForm = UF_DatePicker
    'Reject unresolved DatePicker form instance
        If LoadedForm Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "Unable to resolve loaded DatePicker form instance"
        End If

'------------------------------------------------------------------------------
' PRE-POSITION HIDDEN FORM
'------------------------------------------------------------------------------
    'Suppress best-effort pre-show positioning errors
        On Error Resume Next
    'Move the hidden loaded form close to the current mouse position before first paint
        M_Window_MoveFormToMouse _
            LoadedForm, _
            FORM_MOUSE_OFFSET_XPX, _
            FORM_MOUSE_OFFSET_YPX, _
            FORM_CENTER_ON_MOUSE
    'Clear any suppressed positioning error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SHOW FORM
'------------------------------------------------------------------------------
    'Show the loaded DatePicker form modelessly
        LoadedForm.Show vbModeless

'------------------------------------------------------------------------------
' FINAL POSITION CORRECTION
'------------------------------------------------------------------------------
    'Suppress best-effort post-show positioning errors
        On Error Resume Next
    'Apply one final position correction after the native UserForm window is visible
        M_Window_MoveFormToMouse _
            LoadedForm, _
            FORM_MOUSE_OFFSET_XPX, _
            FORM_MOUSE_OFFSET_YPX, _
            FORM_CENTER_ON_MOUSE
    'Clear any suppressed positioning error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

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
        Err.Raise Err.Number, PROC_NAME, "DatePicker show failed: " & Err.Description

End Sub
Public Sub DP_Click()

'
'------------------------------------------------------------------------------
'                       DATEPICKER CLICK ENTRY POINT
'------------------------------------------------------------------------------
' PURPOSE
'   Public macro callback entry point for DatePicker launch actions
'
' WHY THIS EXISTS
'   Excel UI elements such as CommandBars, worksheet shapes, and in-grid icons
'   require a public Sub that can be assigned as an action macro
'
'   Keeping this routine as a thin wrapper avoids duplicating DatePicker launch
'   logic across UI integration points
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to the canonical ActiveCell-based DatePicker open routine
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker cannot be opened
'
' DEPENDENCIES
'   DP_OpenForActiveCell
'
' NOTES
'   This routine intentionally remains Public so Excel UI surfaces can call it
'
'   Do not add business logic here. Keep DatePicker launch behavior centralized
'   in DP_OpenForActiveCell
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_Click"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' OPEN DATEPICKER
'------------------------------------------------------------------------------
    'Open the DatePicker for the current ActiveCell
        DP_OpenForActiveCell

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
        Err.Raise Err.Number, PROC_NAME, "DatePicker click entry point failed: " _
            & Err.Description

End Sub
Public Sub DP_OpenForActiveCell()

'
'------------------------------------------------------------------------------
'                       OPEN DATEPICKER FOR ACTIVE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Opens the DatePicker for the current ActiveCell
'
' WHY THIS EXISTS
'   The DatePicker needs a stable manual entry point that remains available even
'   when optional UI integrations such as the in-grid icon and right-click menu
'   are disabled
'
'   Public callback surfaces such as keyboard shortcuts, Ribbon buttons, QAT
'   buttons, worksheet buttons, and assigned shape macros should delegate to one
'   canonical DatePicker launch path instead of duplicating launch logic
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to the standard DatePicker show routine, which resolves ActiveCell
'   context, prepares bridge state, loads the form, positions it, and shows it
'   modelessly
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker cannot be opened
'
' DEPENDENCIES
'   DP_Show
'
' NOTES
'   This routine intentionally remains Public
'
'   Keep this routine as a thin wrapper. Do not add ActiveCell parsing, settings
'   loading, form positioning, or write-back logic here
'
'   This routine is intended for Application.OnKey, RibbonX, QAT, worksheet
'   button callbacks, and assigned Shape.OnAction macros
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_OpenForActiveCell"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' OPEN PICKER
'------------------------------------------------------------------------------
    'Show the DatePicker for the current ActiveCell
        DP_Show

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
            "DatePicker open-for-active-cell failed: " & Err.Description

End Sub

Public Sub DP_RepairRuntime()

'
'==============================================================================
'                           REPAIR DATEPICKER RUNTIME
'------------------------------------------------------------------------------
' PURPOSE
'   Repairs the interactive DatePicker runtime after interrupted macros, VBA
'   reset, disabled Excel events, stale manager state, or stale grid icons
'
' WHY THIS EXISTS
'   The DatePicker in-grid icon is event-driven. If Application.EnableEvents is
'   False, Excel will not raise SheetSelectionChange and the icon cannot move,
'   disappear, or refresh when the active cell changes
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Re-enables Excel events, purges stale grid icons, recreates the manager, and
'   refreshes the current active-cell context
'
' ERROR POLICY
'   Raises a descriptive runtime error if the repair cannot be completed
'
' DEPENDENCIES
'   Application.EnableEvents
'   M_GridIcon_PurgeAll
'   M_Picker_EnsureManager
'   gDP_Manager
'
' NOTES
'   This routine is intended for interactive repair, startup, testing, and demo
'   scenarios
'
'   It intentionally forces Application.EnableEvents = True because the
'   DatePicker cannot operate interactively while Excel events are disabled
'
'   Do not call this inside a business macro that deliberately suppresses Excel
'   events unless that macro is ready for events to be re-enabled
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "DP_RepairRuntime"

    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RE-ENABLE EXCEL EVENTS
'------------------------------------------------------------------------------
    'Re-enable Excel events required by the DatePicker manager
        Application.EnableEvents = True

'------------------------------------------------------------------------------
' CLEAR STALE GRID ICONS
'------------------------------------------------------------------------------
    'Purge stale worksheet icon artifacts
        M_GridIcon_PurgeAll

'------------------------------------------------------------------------------
' RECREATE MANAGER
'------------------------------------------------------------------------------
    'Release the current manager reference
        Set gDP_Manager = Nothing
    'Recreate and hook the DatePicker manager
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' REFRESH CURRENT CONTEXT
'------------------------------------------------------------------------------
    'Refresh the DatePicker UI for the current active-cell context
        If Not gDP_Manager Is Nothing Then
            gDP_Manager.Handle_SelectionChange
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
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Raise a descriptive repair error
        Err.Raise ErrorNumber, PROC_NAME, _
            "DatePicker runtime repair failed: " & ErrorDescription

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
'   The DatePicker uses a modeless UserForm and optional live-clock timer
'
'   Closing the picker should therefore stop timer activity, unload the current
'   form instance when available, and clear transient bridge state without
'   raising unnecessary cleanup errors to the user
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stops the live clock timer, resolves the loaded DatePicker form without
'   forcing a default-instance load, hides and unloads the form when present,
'   clears initial-date bridge state, and releases the local form reference
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Suppresses cleanup errors because closing the picker should never interrupt
'   the user workflow
'
' DEPENDENCIES
'   M_Timer_Stop
'   M_FormBridge_GetLoadedForm
'   DP_FORM_NAME
'
' NOTES
'   This routine intentionally avoids direct default-instance references until a
'   loaded form has been explicitly resolved
'
'   Selected-date state is not cleared here because it is owned by the selection
'   / write-back flow and is refreshed when the picker is opened again
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm      As Object        'Loaded DatePicker form instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Stop any active live clock timer
        M_Timer_Stop

'------------------------------------------------------------------------------
' UNLOAD FORM
'------------------------------------------------------------------------------
    'Retrieve the loaded DatePicker form instance without forcing default-instance creation
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Hide and unload the loaded form when present
        If Not LoadedForm Is Nothing Then
            'Hide the form before unloading to reduce visible teardown artifacts
                LoadedForm.Visible = False
            'Unload the loaded form instance
                Unload LoadedForm
        End If

'------------------------------------------------------------------------------
' CLEAR TRANSIENT BRIDGE STATE
'------------------------------------------------------------------------------
    'Clear the initial-date bridge value
        gDP_InitialDate = 0
    'Clear initial-date bridge availability
        gDP_HasInitialDate = False

'------------------------------------------------------------------------------
' RELEASE LOCAL REFERENCES
'------------------------------------------------------------------------------
    'Release the local loaded-form reference
        Set LoadedForm = Nothing

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed cleanup error
        Err.Clear
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



Public Sub M_Picker_SelectDate( _
    ByVal SelectedDate As Date, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                              SELECT DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Stores the selected DatePicker date, writes it to the current Excel target,
'   and applies the configured DatePicker lifecycle behavior
'
' WHY THIS EXISTS
'   Day-label click handling in UF_DatePicker should delegate write-back to the
'   companion module so single-cell, multi-cell, and table-column write-back
'   behavior remains centralized
'
'   The selected-date state should represent a date that was successfully written
'   to Excel, not merely a date that was clicked in the UI
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
'   Ensures settings are loaded, captures the previous transient selection state,
'   normalizes the selected date to a date-only value, prepares the write-back
'   value, writes it to the current Excel target, stores the selected-date state
'   only after successful write-back, and then closes or refreshes the picker
'   according to gDP_CloseAfterSelection
'
' ERROR POLICY
'   Raises a descriptive runtime error if the selected date is invalid, if
'   settings cannot be loaded, or if write-back fails
'
'   Restores the previous selected-date and write-value state if write-back fails
'
'   Post-write-back form close / refresh is best-effort so a successful Excel
'   write-back is not converted into a user-facing error by a visual refresh
'   issue
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_WriteBack_Apply
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   Calendar day selection is intentionally date-only
'
'   SelectedDate is normalized once and the normalized value is reused for
'   write-back, selected-state storage, and optional open-form refresh
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Picker_SelectDate"

    Dim SelectedDateOnly    As Date         'Selected date without time
    Dim OldSelectedDate     As Date         'Previous selected date
    Dim OldHasSelectedDate  As Boolean      'Previous selected-date availability
    Dim OldWriteValue       As Variant      'Previous transient write value
    Dim StateCaptured       As Boolean      'True when rollback state is available
    Dim ErrorNumber         As Long         'Captured runtime error number
    Dim ErrorDescription    As String       'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure DatePicker settings are available before lifecycle behavior is evaluated
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE PREVIOUS STATE
'------------------------------------------------------------------------------
    'Capture the previous selected date
        OldSelectedDate = gDP_SelectedDate
    'Capture previous selected-date availability
        OldHasSelectedDate = gDP_HasSelectedDate
    'Capture the previous transient write value
        OldWriteValue = gDP_WriteValue
    'Mark rollback state as available
        StateCaptured = True

'------------------------------------------------------------------------------
' NORMALIZE SELECTED DATE
'------------------------------------------------------------------------------
    'Normalize the selected date to a date-only value
        SelectedDateOnly = VBA.DateValue(SelectedDate)

'------------------------------------------------------------------------------
' VALIDATE SELECTED DATE
'------------------------------------------------------------------------------
    'Reject an empty date value
        If SelectedDateOnly = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "SelectedDate cannot be zero"
        End If

'------------------------------------------------------------------------------
' PREPARE WRITE VALUE
'------------------------------------------------------------------------------
    'Store the date-only value to write
        gDP_WriteValue = SelectedDateOnly

'------------------------------------------------------------------------------
' WRITE TO EXCEL
'------------------------------------------------------------------------------
    'Write the selected date to the current Excel target
        M_WriteBack_Apply Date_Picker, NoTableGrow

'------------------------------------------------------------------------------
' STORE SELECTED DATE AFTER SUCCESSFUL WRITE-BACK
'------------------------------------------------------------------------------
    'Store the selected date only after write-back succeeds
        gDP_SelectedDate = SelectedDateOnly
    'Mark the selected date as available
        gDP_HasSelectedDate = True

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Suppress post-write-back visual lifecycle errors
        On Error Resume Next
    'Close the DatePicker after successful selection when configured
        If gDP_CloseAfterSelection Then
            'Close the DatePicker form
                DP_Close
        Else
            'Refresh the open DatePicker form after successful selection
                M_FormBridge_AfterSuccessfulSelection SelectedDateOnly
        End If
    'Clear any suppressed visual lifecycle error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Restore previous transient state when available
        If StateCaptured Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous selected date
                gDP_SelectedDate = OldSelectedDate
            'Restore previous selected-date availability
                gDP_HasSelectedDate = OldHasSelectedDate
            'Restore the previous transient write value
                gDP_WriteValue = OldWriteValue
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, "DatePicker date selection failed: " & ErrorDescription

End Sub
Public Sub DP_Today()

'
'------------------------------------------------------------------------------
'                               WRITE TODAY
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current system date to the current Excel target
'
' WHY THIS EXISTS
'   Footer labels, Ribbon callbacks, context-menu actions, keyboard shortcuts, or
'   public macro callers may need a direct Today command without requiring the
'   user to select a day from the calendar grid
'
'   Today should follow the same write-back, selected-state, rollback, and
'   close-or-refresh lifecycle used by normal day-cell selection
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the current system date once and delegates the date-only write-back
'   to M_Picker_SelectDate
'
' ERROR POLICY
'   Raises a descriptive runtime error if the Today command cannot be completed
'
' DEPENDENCIES
'   M_Picker_SelectDate
'
' NOTES
'   This command always writes a date-only value
'
'   Write-back state management is intentionally centralized in
'   M_Picker_SelectDate so Today, calendar-day clicks, and other date-only
'   selection paths remain behaviorally consistent
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_Today"

    Dim TodayDate       As Date     'Current system date without time

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE TODAY
'------------------------------------------------------------------------------
    'Store the current system date once for consistent write-back and refresh state
        TodayDate = VBA.Date

'------------------------------------------------------------------------------
' WRITE TODAY
'------------------------------------------------------------------------------
    'Delegate date-only write-back to the canonical DatePicker selection routine
        M_Picker_SelectDate TodayDate, False

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
        Err.Raise Err.Number, PROC_NAME, "DatePicker Today command failed: " & Err.Description

End Sub
Public Sub DP_Now()

'
'------------------------------------------------------------------------------
'                               WRITE NOW
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current system date-time to the current Excel target
'
' WHY THIS EXISTS
'   Footer labels, Ribbon callbacks, context-menu actions, keyboard shortcuts, or
'   public macro callers may need a direct Now command without requiring the user
'   to select a day from the calendar grid
'
'   Now is intentionally different from normal calendar-day selection because it
'   writes a date-time value while the DatePicker selected-day visual state uses
'   only the date portion
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded, captures the current system timestamp once,
'   stores the timestamp as the write-back value, writes it to the current Excel
'   target, commits the selected-date state only after successful write-back, and
'   then closes or refreshes the picker according to gDP_CloseAfterSelection
'
' ERROR POLICY
'   Raises a descriptive runtime error if the Now command cannot be completed
'
'   Restores the previous selected-date and write-value state if write-back fails
'
'   Post-write-back form close / refresh is best-effort so a successful Excel
'   write-back is not converted into a user-facing error by a visual refresh
'   issue
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_WriteBack_Apply
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   This command writes a date-time value
'
'   Selected-date highlighting uses only the date portion of the timestamp
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "DP_Now"

    Dim NowValue            As Date                 'Current system date-time
    Dim NowDate             As Date                 'Date-only part of current timestamp
    Dim OldSelectedDate     As Date                 'Previous selected date
    Dim OldHasSelectedDate  As Boolean              'Previous selected-date availability
    Dim OldWriteValue       As Variant              'Previous transient write value
    Dim StateCaptured       As Boolean              'True when rollback state is available
    Dim ErrorNumber         As Long                 'Captured runtime error number
    Dim ErrorDescription    As String               'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure DatePicker settings are available before lifecycle behavior is evaluated
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' CAPTURE PREVIOUS STATE
'------------------------------------------------------------------------------
    'Capture the previous selected date
        OldSelectedDate = gDP_SelectedDate
    'Capture previous selected-date availability
        OldHasSelectedDate = gDP_HasSelectedDate
    'Capture the previous transient write value
        OldWriteValue = gDP_WriteValue
    'Mark rollback state as available
        StateCaptured = True

'------------------------------------------------------------------------------
' RESOLVE CURRENT TIMESTAMP
'------------------------------------------------------------------------------
    'Store the current system timestamp once for consistent write-back and refresh state
        NowValue = VBA.Now
    'Resolve the date-only part of the current timestamp
        NowDate = VBA.DateValue(NowValue)

'------------------------------------------------------------------------------
' PREPARE WRITE VALUE
'------------------------------------------------------------------------------
    'Store the current system timestamp as the write-back value
        gDP_WriteValue = NowValue

'------------------------------------------------------------------------------
' WRITE TO EXCEL
'------------------------------------------------------------------------------
    'Apply the date-time value to the current Excel target
        M_WriteBack_Apply Date_Picker, False

'------------------------------------------------------------------------------
' STORE SELECTED DATE AFTER SUCCESSFUL WRITE-BACK
'------------------------------------------------------------------------------
    'Store the date-only part as the active selected date
        gDP_SelectedDate = NowDate
    'Mark selected-date state as available
        gDP_HasSelectedDate = True

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Suppress post-write-back visual lifecycle errors
        On Error Resume Next
    'Close the DatePicker after successful write-back when configured
        If gDP_CloseAfterSelection Then
            'Close the DatePicker form
                DP_Close
        Else
            'Refresh the open DatePicker form using the date-only part
                M_FormBridge_AfterSuccessfulSelection NowDate
        End If
    'Clear any suppressed visual lifecycle error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
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
    'Restore previous transient state when available
        If StateCaptured Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous selected date
                gDP_SelectedDate = OldSelectedDate
            'Restore previous selected-date availability
                gDP_HasSelectedDate = OldHasSelectedDate
            'Restore the previous transient write value
                gDP_WriteValue = OldWriteValue
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, "DatePicker Now command failed: " & ErrorDescription

End Sub
Public Sub M_WriteBack_Apply( _
    ByVal iType As DP_WriteAction, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                           APPLY WRITE-BACK
'------------------------------------------------------------------------------
' PURPOSE
'   Applies a DatePicker write-back action to the current Excel selection
'
' WHY THIS EXISTS
'   UI click handlers, footer actions, keyboard shortcuts, context-menu actions,
'   and public entry points need one shared routine that writes DatePicker values
'   to Excel while suppressing worksheet events safely
'
' INPUTS
'   iType
'     DatePicker write action to apply
'
'   NoTableGrow
'     True to keep a single table-cell target as a single cell
'
'     False to allow single table-cell selections to expand to the full table
'     data column when applicable
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the requested write action, captures the current Excel event state,
'   disables events during write-back, delegates target resolution and write-back
'   to M_WriteBack_ResolveAndApplyTarget, and restores the previous Excel event
'   state before exiting
'
' ERROR POLICY
'   Restores Application.EnableEvents when the previous state was captured
'
'   Re-raises the original write-back error after cleanup
'
'   Raises a cleanup error only when event restoration fails and no original
'   write-back error exists
'
' DEPENDENCIES
'   M_WriteBack_ResolveAndApplyTarget
'   Application.EnableEvents
'
' NOTES
'   This routine intentionally does not change calculation mode
'
'   This routine intentionally does not change screen updating
'
'   Unsupported write actions are rejected explicitly instead of silently doing
'   nothing
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_WriteBack_Apply"

    Dim PreviousEvents          As Boolean      'Prior Application.EnableEvents state
    Dim EventsStateCaptured     As Boolean      'True when PreviousEvents is available

    Dim SavedErrNumber          As Long         'Captured original error number
    Dim SavedErrSource          As String       'Captured original error source
    Dim SavedErrDescription     As String       'Captured original error description

    Dim CleanupErrNumber        As Long         'Captured cleanup error number
    Dim CleanupErrDescription   As String       'Captured cleanup error description

    Dim HandlerStep             As String       'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE WRITE ACTION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate write action"

    'Validate the requested write action
        Select Case iType
            Case DP_WriteAction.Date_Picker
                'Supported DatePicker write action
            Case Else
                'Reject unsupported write actions
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Unsupported DatePicker write action: " & VBA.CStr(VBA.CLng(iType))
        End Select

'------------------------------------------------------------------------------
' CAPTURE EXCEL EVENT STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture Excel event state"

    'Capture the current Excel events state
        PreviousEvents = Application.EnableEvents
    'Mark the Excel event state as captured
        EventsStateCaptured = True

'------------------------------------------------------------------------------
' SUPPRESS EVENTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Suppress Excel events"

    'Disable events only when they are currently enabled
        If PreviousEvents Then Application.EnableEvents = False

'------------------------------------------------------------------------------
' DISPATCH WRITE-BACK
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve and apply write-back target"

    'Apply the requested action to the current selection
        M_WriteBack_ResolveAndApplyTarget iType, NoTableGrow

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Protect cleanup from masking the original error
        On Error Resume Next
    'Restore the previous Excel event state when it was captured
        If EventsStateCaptured Then Application.EnableEvents = PreviousEvents
    'Capture cleanup failure when no original error exists
        If Err.Number <> 0 Then
            CleanupErrNumber = Err.Number
            CleanupErrDescription = Err.Description
            Err.Clear
        End If
    'Restore normal error handling
        On Error GoTo 0
    'Re-raise the original write-back error when needed
        If SavedErrNumber <> 0 Then
            'Append cleanup diagnostics when cleanup also failed
                If CleanupErrNumber <> 0 Then
                    SavedErrDescription = SavedErrDescription & _
                        " Cleanup also failed while restoring Application.EnableEvents: " & _
                        CleanupErrDescription
                End If
            'Raise the original error after best-effort cleanup
                Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription
        End If
    'Raise cleanup failure when there was no original write-back error
        If CleanupErrNumber <> 0 Then
            Err.Raise CleanupErrNumber, PROC_NAME, _
                "DatePicker write-back cleanup failed while restoring Application.EnableEvents: " & _
                CleanupErrDescription
        End If
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        SavedErrNumber = Err.Number
    'Capture the original error source
        SavedErrSource = PROC_NAME & " | Step=" & HandlerStep
    'Capture the original error description
        SavedErrDescription = "DatePicker write-back failed: " & Err.Description
    'Resume through cleanup
        Resume CleanExit

End Sub
Private Sub M_WriteBack_ResolveAndApplyTarget( _
    ByVal iType As DP_WriteAction, _
    Optional ByVal NoTableGrow As Boolean = False)

'
'------------------------------------------------------------------------------
'                       RESOLVE AND APPLY WRITE-BACK TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the current Excel write-back target and applies the requested
'   DatePicker write action
'
' WHY THIS EXISTS
'   DatePicker UI handlers should not write directly to Excel
'
'   This routine centralizes target shaping so single-cell, multi-cell,
'   discontiguous-range, and table-column write-back behavior remains consistent
'   across calendar clicks, Today, Now, keyboard shortcuts, context-menu actions,
'   and public macro entry points
'
' INPUTS
'   iType
'     DatePicker write action to apply
'
'   NoTableGrow
'     True to keep a single selected table cell as a single-cell target
'
'     False to expand a single selected table data-body cell to the full table
'     data column
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the requested write action, resolves the current Excel selection,
'   rejects non-range selections, optionally expands a single table data-body
'   cell to its full ListObject data column, and writes to each target area
'
' ERROR POLICY
'   Raises a descriptive runtime error if the write action is unsupported, if
'   the current Excel selection is not a Range, if table target expansion fails,
'   or if range population fails
'
' DEPENDENCIES
'   M_WriteBack_PopulateRange
'   Application.Selection
'   Excel.Range
'   Excel.ListObject
'
' NOTES
'   This routine does not suppress Application events
'
'   Application.EnableEvents is managed by M_WriteBack_Apply
'
'   This routine intentionally raises on non-Range selections so callers do not
'   treat a no-op as a successful write-back
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_WriteBack_ResolveAndApplyTarget"

    Dim SelectedObject      As Object           'Current Excel selection object
    Dim Target              As Range            'Resolved target range
    Dim Block               As Range            'Single target area
    Dim TargetTable         As ListObject       'Worksheet table being inspected
    Dim ColumnIndex         As Long             'Resolved table column index
    Dim HandlerStep         As String           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE WRITE ACTION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate write action"

    'Validate the requested write action
        Select Case iType
            Case Date_Picker
                'Supported DatePicker write action
            Case Else
                'Reject unsupported write actions
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Unsupported DatePicker write action: " & VBA.CStr(VBA.CLng(iType))
        End Select

'------------------------------------------------------------------------------
' RESOLVE CURRENT SELECTION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve current Excel selection"

    'Suppress selection access errors temporarily
        On Error Resume Next
    'Capture the current Excel selection object
        Set SelectedObject = Application.Selection
    'Clear any suppressed selection access error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler
    'Reject missing selection objects
        If SelectedObject Is Nothing Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Current Excel selection is not available"
        End If
    'Reject non-range selections
        If VBA.TypeName(SelectedObject) <> "Range" Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "Current Excel selection must be a Range. Current selection type is '" & _
                VBA.TypeName(SelectedObject) & "'"
        End If
    'Use the current selection as the initial write-back target
        Set Target = SelectedObject

'------------------------------------------------------------------------------
' EXPAND SINGLE TABLE CELL WHEN ALLOWED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve optional table-column expansion"

    'Consider table expansion only for one selected cell when allowed
        If Target.Cells.CountLarge = 1 Then
            If Not NoTableGrow Then
                'Loop through worksheet tables
                    For Each TargetTable In Target.Worksheet.ListObjects
                        'Continue only when the table has a data body
                            If Not TargetTable.DataBodyRange Is Nothing Then
                                'Expand only when the selected cell is inside the table data body
                                    If Not Application.Intersect(Target, TargetTable.DataBodyRange) Is Nothing Then
                                        'Resolve the selected table column index
                                            ColumnIndex = Target.Column - TargetTable.DataBodyRange.Column + 1
                                        'Expand to the table data column when the column index is valid
                                            If ColumnIndex >= 1 Then
                                                If ColumnIndex <= TargetTable.ListColumns.Count Then
                                                    Set Target = TargetTable.ListColumns(ColumnIndex).DataBodyRange
                                                End If
                                            End If
                                        'Stop after resolving the owning table
                                            Exit For
                                    End If
                            End If
                    Next TargetTable
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE RESOLVED TARGET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate resolved write-back target"

    'Reject a missing resolved target
        If Target Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, "Unable to resolve DatePicker write-back target"
        End If
    'Reject empty resolved targets
        If Target.Cells.CountLarge = 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, "Resolved DatePicker write-back target is empty"
        End If

'------------------------------------------------------------------------------
' POPULATE TARGET AREAS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Populate target areas"

    'Loop through each discontiguous target area
        For Each Block In Target.Areas
            'Populate this target area
                M_WriteBack_PopulateRange Block, iType
        Next Block

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set Block = Nothing
    'Release object references
        Set TargetTable = Nothing
    'Release object references
        Set Target = Nothing
    'Release object references
        Set SelectedObject = Nothing
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker write-back target resolution failed: " & Err.Description

End Sub
Public Sub M_WriteBack_PopulateRange( _
    ByVal oRange As Range, _
    ByVal iType As DP_WriteAction)

'
'------------------------------------------------------------------------------
'                           POPULATE RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current DatePicker value to every cell in a resolved Excel range
'
' WHY THIS EXISTS
'   DatePicker write-back can target:
'     - one cell
'     - multiple selected cells
'     - discontiguous areas
'     - a resolved table data column
'
'   This routine centralizes the final cell-level population logic so write-back
'   behavior remains consistent across calendar-day selection, Today, Now,
'   keyboard shortcuts, context-menu actions, and public macro entry points
'
' INPUTS
'   oRange
'     Resolved Excel target range to populate
'
'   iType
'     DatePicker write action used to resolve the value to write
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates the target range and write action, resolves the DatePicker write
'   value, writes the value cell by cell through M_WriteBack_TryWriteCell, counts
'   protected locked cells and other failures, reports partial protected-cell
'   skips once, logs partial non-lock failures to the Immediate Window, and
'   raises an error when no cell was successfully written
'
' ERROR POLICY
'   Raises a descriptive runtime error if the target range is missing, the write
'   action is unsupported, the write value cannot be resolved, or no cell can be
'   written successfully
'
'   Protected locked cells may be skipped when at least one target cell is
'   written successfully
'
'   Other cell-level failures may be suppressed by M_WriteBack_TryWriteCell when
'   at least one target cell is written successfully, but they are logged for
'   diagnostics
'
' DEPENDENCIES
'   M_WriteBack_GetPickedValue
'   M_WriteBack_TryWriteCell
'   DP_MSGBOX_TITLE
'
' NOTES
'   This routine intentionally writes cell by cell instead of assigning the whole
'   range in one operation
'
'   Cell-by-cell write-back is slower for very large ranges but safer for
'   protected sheets, locked cells, validation failures, and partially writable
'   selections
'
'   This routine assumes M_WriteBack_TryWriteCell increments LockedCount and
'   FailedCount only when a cell is not written
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_WriteBack_PopulateRange"
    
    Dim Cell            As Range            'Current target cell
    Dim LockedCount     As Long             'Protected locked cells skipped
    Dim FailedCount     As Long             'Other write failures suppressed
    Dim AttemptedCount  As Double           'Total target cells attempted
    Dim WrittenCount    As Double           'Total target cells successfully written
    Dim WriteValue      As Variant          'Resolved write value
    Dim HandlerStep     As String           'Current handler step for diagnostics

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE TARGET RANGE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate target range"

    'Reject missing target ranges
        If oRange Is Nothing Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Target range cannot be Nothing"
        End If
    'Capture the number of target cells
        AttemptedCount = oRange.Cells.CountLarge
    'Reject empty target ranges
        If AttemptedCount <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Target range does not contain writable cells"
        End If

'------------------------------------------------------------------------------
' RESOLVE WRITE VALUE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve write value"

    'Resolve the value to write from the requested DatePicker action
        Select Case iType
            Case Date_Picker
                'Resolve the current DatePicker write value
                    WriteValue = M_WriteBack_GetPickedValue
            Case Else
                'Reject unsupported write actions
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Unsupported DatePicker write action: " & VBA.CStr(VBA.CLng(iType))
        End Select

'------------------------------------------------------------------------------
' POPULATE TARGET CELLS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Populate target cells"

    'Loop through each target cell
        For Each Cell In oRange.Cells
            'Attempt to write the resolved value to the current cell
                M_WriteBack_TryWriteCell Cell, WriteValue, LockedCount, FailedCount
        Next Cell

'------------------------------------------------------------------------------
' RESOLVE WRITE RESULT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve write result"

    'Calculate the number of successfully written cells
        WrittenCount = AttemptedCount - LockedCount - FailedCount
    'Reject write-back attempts that did not write any cell
        If WrittenCount <= 0 Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DatePicker write-back did not write any cell. Target cells: " & _
                VBA.CStr(AttemptedCount) & "; protected locked cells skipped: " & _
                VBA.CStr(LockedCount) & "; other failures: " & VBA.CStr(FailedCount)
        End If

'------------------------------------------------------------------------------
' REPORT PARTIAL PROTECTED-CELL SKIPS
'------------------------------------------------------------------------------
    'Show one summary message for protected locked cells that were skipped
        If LockedCount > 0 Then
            MsgBox VBA.CStr(LockedCount) & " protected locked cell(s) were skipped.", _
                vbInformation Or vbOKOnly, DP_MSGBOX_TITLE
        End If

'------------------------------------------------------------------------------
' LOG PARTIAL NON-LOCK FAILURES
'------------------------------------------------------------------------------
    'Write non-lock failures to the Immediate Window for diagnostics
        If FailedCount > 0 Then
            Debug.Print PROC_NAME & ": " & VBA.CStr(FailedCount) & _
                " cell write failure(s) were suppressed."
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set Cell = Nothing
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Release object references
        Set Cell = Nothing
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker range population failed: " & Err.Description

End Sub
Private Function M_WriteBack_TryWriteCell( _
    ByVal TargetCell As Range, _
    ByVal WriteValue As Variant, _
    ByRef LockedCount As Long, _
    ByRef FailedCount As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY WRITE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to write one DatePicker value to one Excel cell
'
' WHY THIS EXISTS
'   Range write-back can involve many target cells
'
'   A single protected, locked, invalid, or otherwise failing cell should not
'   necessarily stop the whole DatePicker write-back when other cells remain
'   writable
'
'   This routine centralizes safe per-cell write behavior and reports the result
'   through a Boolean return value plus failure counters
'
' INPUTS
'   TargetCell
'     Single Excel cell to populate
'
'   WriteValue
'     DatePicker value to write
'
'   LockedCount
'     Counter incremented when a protected locked cell is skipped
'
'   FailedCount
'     Counter incremented when another cell-level write failure occurs
'
' RETURNS
'   True when the value was successfully written to TargetCell
'
'   False when TargetCell was missing, invalid, protected locked, or could not be
'   written
'
' BEHAVIOR
'   Validates the target cell, skips protected locked cells, writes the supplied
'   value to writable cells, returns True only after a successful write, and logs
'   suppressed write failures to the Immediate Window
'
' ERROR POLICY
'   Best-effort per-cell write
'
'   Does not raise outward. Cell-level failures are counted and suppressed so the
'   caller can complete the range write-back and decide whether the aggregate
'   result is acceptable
'
' DEPENDENCIES
'   Excel.Range
'   Excel.Worksheet.ProtectContents
'
' NOTES
'   This routine intentionally uses Range.Value rather than Range.Value2 so VBA
'   Date and DateTime variants are written through Excel's normal date handling
'
'   Application.EnableEvents is managed by M_WriteBack_Apply, not by this routine
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_WriteBack_TryWriteCell"

    Dim TargetAddress   As String       'Diagnostic target address
    Dim TargetSheetName As String       'Diagnostic worksheet name
    Dim ErrorNumber     As Long         'Captured runtime error number
    Dim ErrorText       As String       'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_WriteBack_TryWriteCell = False
    'Enable per-cell fail-safe handling
        On Error GoTo WriteFail

'------------------------------------------------------------------------------
' VALIDATE TARGET CELL
'------------------------------------------------------------------------------
    'Count and exit when no target cell is supplied
        If TargetCell Is Nothing Then
            FailedCount = FailedCount + 1
            Debug.Print PROC_NAME & ": skipped missing target cell"
            Exit Function
        End If
    'Count and exit when a non-single-cell range is supplied unexpectedly
        If TargetCell.Cells.CountLarge <> 1 Then
            FailedCount = FailedCount + 1
            Debug.Print PROC_NAME & ": skipped non-single-cell target"
            Exit Function
        End If

'------------------------------------------------------------------------------
' CAPTURE DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Capture the target worksheet name for diagnostics
        TargetSheetName = TargetCell.Worksheet.Name
    'Capture the target cell address for diagnostics
        TargetAddress = TargetCell.Address(False, False)

'------------------------------------------------------------------------------
' SKIP PROTECTED LOCKED CELLS
'------------------------------------------------------------------------------
    'Skip locked cells on protected sheets
        If TargetCell.Worksheet.ProtectContents Then
            If TargetCell.Locked Then
                LockedCount = LockedCount + 1
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' WRITE CELL VALUE
'------------------------------------------------------------------------------
    'Write the resolved DatePicker value to the target cell
        TargetCell.Value = WriteValue

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return success only after the value has been written
        M_WriteBack_TryWriteCell = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful write
        Exit Function

'------------------------------------------------------------------------------
' WRITE FAIL
'------------------------------------------------------------------------------
WriteFail:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorText = Err.Description
    'Suppress diagnostic failures
        On Error Resume Next
    'Increment the non-lock failure counter
        FailedCount = FailedCount + 1
    'Refresh diagnostic context if it was not captured before the failure
        If Len(TargetAddress) = 0 Then
            If Not TargetCell Is Nothing Then
                TargetAddress = TargetCell.Address(False, False)
            End If
        End If
    'Refresh diagnostic context if it was not captured before the failure
        If Len(TargetSheetName) = 0 Then
            If Not TargetCell Is Nothing Then
                TargetSheetName = TargetCell.Worksheet.Name
            End If
        End If
    'Write diagnostics to the Immediate Window
        Debug.Print PROC_NAME & ": suppressed error " & VBA.CStr(ErrorNumber) & _
            " while writing " & TargetSheetName & "!" & TargetAddress & _
            " - " & ErrorText
    'Return safe failure result
        M_WriteBack_TryWriteCell = False
    'Restore normal error handling
        On Error GoTo 0

End Function
Private Function M_WriteBack_GetPickedValue() As Date

'
'------------------------------------------------------------------------------
'                           GET PICKED WRITE VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current DatePicker value prepared for Excel write-back
'
' WHY THIS EXISTS
'   DatePicker write-back routines should retrieve the value to write through one
'   controlled helper instead of reading transient module state directly
'
'   This provides one validation point before the value is written to Excel
'
' INPUTS
'   None
'
' RETURNS
'   DatePicker write value as a VBA Date
'
'   The returned value may contain a time component when the caller prepared a
'   date-time value, such as through DP_Now
'
' BEHAVIOR
'   Validates that gDP_WriteValue contains a usable date or date-time value,
'   converts it to a VBA Date, validates the supported DatePicker year range,
'   and returns the validated value
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker write value is empty,
'   Null, an Excel error value, non-date, or outside the supported DatePicker
'   year range
'
' DEPENDENCIES
'   gDP_WriteValue
'   DP_MIN_YEAR
'   DP_MAX_YEAR
'
' NOTES
'   This routine intentionally uses CDate rather than DateValue so date-time
'   values retain their time component
'
'   Calendar day selection and Today prepare date-only values upstream
'
'   Now prepares a date-time value upstream
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_WriteBack_GetPickedValue"
    
    Dim PickedValue     As Date     'Resolved DatePicker write value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE TRANSIENT WRITE VALUE
'------------------------------------------------------------------------------
    'Reject Null write values
        If IsNull(gDP_WriteValue) Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DatePicker write value cannot be Null"
        End If
    'Reject Excel error values
        If IsError(gDP_WriteValue) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DatePicker write value cannot be an Excel error value"
        End If
    'Reject empty write values
        If IsEmpty(gDP_WriteValue) Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DatePicker write value has not been initialized"
        End If
    'Reject non-date write values
        If Not VBA.IsDate(gDP_WriteValue) Then
            Err.Raise vbObjectError + 516, PROC_NAME, _
                "DatePicker write value must be a date or date-time value"
        End If

'------------------------------------------------------------------------------
' CONVERT WRITE VALUE
'------------------------------------------------------------------------------
    'Convert the transient write value to a VBA Date while preserving time
        PickedValue = VBA.CDate(gDP_WriteValue)

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the validated DatePicker write value
        M_WriteBack_GetPickedValue = PickedValue

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
            "DatePicker write value resolution failed: " & Err.Description

End Function
Public Function M_DatePolicy_CanSelectDate( _
    ByVal CandidateDate As Date, _
    ByVal DisplayYear As Long, _
    ByVal DisplayMonth As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           CAN SELECT DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate calendar date can be selected in the DatePicker
'
' WHY THIS EXISTS
'   The UserForm should not decide selection policy directly
'
'   Date selection rules belong in the companion module so calendar rendering,
'   hover behavior, keyboard selection, and click handling all use the same
'   policy
'
' INPUTS
'   CandidateDate
'     Date being evaluated for selection
'
'   DisplayYear
'     Year currently displayed by the DatePicker grid
'
'   DisplayMonth
'     Month currently displayed by the DatePicker grid
'
' RETURNS
'   True when the candidate date can be selected
'
'   False when the display period is invalid or when outside-month selection is
'   disabled and the candidate date does not belong to the displayed month
'
' BEHAVIOR
'   Validates the displayed year and month, normalizes the candidate date to a
'   date-only value, allows all valid candidate dates when outside-month
'   selection is enabled, and otherwise allows only dates belonging to the
'   displayed month and year
'
' ERROR POLICY
'   Fail-closed policy routine
'
'   Returns False for invalid display state or unexpected runtime errors
'
'   Does not raise outward because this routine is used by high-frequency UI
'   paths such as grid rendering, hover handling, and keyboard navigation
'
' DEPENDENCIES
'   gDP_AllowOutsideMonthSel
'
' NOTES
'   This routine intentionally does not call M_Settings_EnsureLoaded
'
'   Settings are loaded by public entry points and UserForm initialization before
'   this routine is used in normal runtime paths
'
'   CandidateDate is normalized with DateValue so date-time values are evaluated
'   by their date component only
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const MIN_SUPPORTED_YEAR    As Long = 100       'Minimum supported DatePicker year
    Const MAX_SUPPORTED_YEAR    As Long = 9999      'Maximum supported DatePicker year

    Dim CandidateDateOnly       As Date             'Candidate date without time
    Dim CandidateYear           As Long             'Candidate date year
    Dim CandidateMonth          As Long             'Candidate date month

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_DatePolicy_CanSelectDate = False
    'Enable fail-closed error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Reject invalid display years
        If DisplayYear < MIN_SUPPORTED_YEAR Or DisplayYear > MAX_SUPPORTED_YEAR Then
            Exit Function
        End If
    'Reject invalid display months
        If DisplayMonth < 1 Or DisplayMonth > 12 Then Exit Function


'------------------------------------------------------------------------------
' NORMALIZE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Normalize the candidate date to its date-only component
        CandidateDateOnly = VBA.DateValue(CandidateDate)
    'Resolve the candidate year once
        CandidateYear = VBA.Year(CandidateDateOnly)
    'Resolve the candidate month once
        CandidateMonth = VBA.Month(CandidateDateOnly)

'------------------------------------------------------------------------------
' VALIDATE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Reject candidate dates outside the supported DatePicker year range
        If CandidateYear < MIN_SUPPORTED_YEAR Or CandidateYear > MAX_SUPPORTED_YEAR Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY OUTSIDE-MONTH POLICY
'------------------------------------------------------------------------------
    'Allow valid candidate dates when outside-month selection is enabled
        If gDP_AllowOutsideMonthSel Then
            M_DatePolicy_CanSelectDate = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY DISPLAYED-MONTH POLICY
'------------------------------------------------------------------------------
    'Allow only dates belonging to the displayed month and year
        M_DatePolicy_CanSelectDate = _
            (CandidateYear = DisplayYear And CandidateMonth = DisplayMonth)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after policy evaluation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default on unexpected policy errors
        M_DatePolicy_CanSelectDate = False

End Function
Public Function M_HolidayPolicy_IsHolidayDate(ByVal CandidateDate As Date) As Boolean

'
'------------------------------------------------------------------------------
'                           IS HOLIDAY DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate date is classified as a holiday by the optional
'   DatePicker holiday callback
'
' WHY THIS EXISTS
'   Holiday logic is workbook-specific
'
'   The DatePicker companion module should provide one controlled policy point
'   that can call a user-defined callback without hard-coding holiday calendars
'   into the UserForm rendering layer
'
' INPUTS
'   CandidateDate
'     Date to evaluate
'
' RETURNS
'   True when the configured callback explicitly returns True
'
'   False when no callback is configured, when the callback fails, when the
'   callback does not return a Boolean, or when the candidate date is outside
'   the supported DatePicker range
'
' BEHAVIOR
'   Trims the configured callback name, exits safely when no callback exists,
'   normalizes CandidateDate to a date-only value, validates the supported year
'   range, invokes the callback through Application.Run, and accepts only a
'   Boolean callback result
'
' ERROR POLICY
'   Fail-open policy routine
'
'   Does not raise outward because this routine may be used by high-frequency UI
'   rendering and selection-policy paths
'
'   Unexpected callback errors are suppressed and interpreted as not holiday
'
' DEPENDENCIES
'   Application.Run
'   gDP_HolidayCallbackName
'
' NOTES
'   The holiday callback should accept one Date argument and return Boolean
'
'   The callback receives the date-only component of CandidateDate
'
'   Non-Boolean callback results are ignored deliberately to keep the callback
'   contract strict and predictable
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const MIN_SUPPORTED_YEAR    As Long = 100       'Minimum supported DatePicker year
    Const MAX_SUPPORTED_YEAR    As Long = 9999      'Maximum supported DatePicker year

    Dim CallbackName            As String           'Configured holiday callback name
    Dim CandidateDateOnly       As Date             'Candidate date without time
    Dim CandidateYear           As Long             'Candidate date year
    Dim CallbackResult          As Variant          'Callback return value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_HolidayPolicy_IsHolidayDate = False
    'Enable fail-open error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' RESOLVE CALLBACK NAME
'------------------------------------------------------------------------------
    'Normalize the configured callback name
        CallbackName = VBA.Trim$(gDP_HolidayCallbackName)
    'Exit when no callback is configured
        If VBA.Len(CallbackName) = 0 Then Exit Function

'------------------------------------------------------------------------------
' NORMALIZE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Normalize the candidate date to its date-only component
        CandidateDateOnly = VBA.DateValue(CandidateDate)
    'Resolve the candidate year once
        CandidateYear = VBA.Year(CandidateDateOnly)

'------------------------------------------------------------------------------
' VALIDATE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Reject dates before the supported DatePicker year range
        If CandidateYear < MIN_SUPPORTED_YEAR Then Exit Function
    'Reject dates after the supported DatePicker year range
        If CandidateYear > MAX_SUPPORTED_YEAR Then Exit Function

'------------------------------------------------------------------------------
' RUN HOLIDAY CALLBACK
'------------------------------------------------------------------------------
    'Run the configured holiday callback with the normalized candidate date
        CallbackResult = Application.Run(CallbackName, CandidateDateOnly)

'------------------------------------------------------------------------------
' INTERPRET CALLBACK RESULT
'------------------------------------------------------------------------------
    'Ignore Excel error callback results
        If VBA.IsError(CallbackResult) Then Exit Function
    'Ignore Null callback results
        If VBA.IsNull(CallbackResult) Then Exit Function
    'Return the callback result only when it is explicitly Boolean
        If VBA.VarType(CallbackResult) = vbBoolean Then
            M_HolidayPolicy_IsHolidayDate = VBA.CBool(CallbackResult)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after holiday-policy evaluation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Suppress callback or policy failures
        On Error Resume Next

    'Return safe default
        M_HolidayPolicy_IsHolidayDate = False

    'Clear any pending error
        Err.Clear

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

    'Exit only when native WinAPI calls are unavailable
        If Not M_Platform_CanUseWinAPI Then Exit Sub

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
'                              KEYBOARD SHORTCUT
'
'------------------------------------------------------------------------------
'

Public Sub M_KeyboardShortcut_Update()

'
'------------------------------------------------------------------------------
'                         UPDATE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Registers or removes the DatePicker keyboard shortcut according to settings
'
' WHY THIS EXISTS
'   The keyboard shortcut is the safe fallback entry point when optional visual
'   integrations such as the in-grid icon and right-click menu are disabled
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Loads settings if needed, registers Ctrl + Shift + D when enabled, and
'   restores the key to Excel when disabled
'
' ERROR POLICY
'   Raises a descriptive runtime error if shortcut synchronization fails
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_KeyboardShortcut_Register
'   M_KeyboardShortcut_Remove
'   gDP_EnableKeyboardShortcut
'
' NOTES
'   Application.OnKey is application-wide for the current Excel session, so it
'   must be removed during teardown
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_KeyboardShortcut_Update"     'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

    'Ensure settings are available before reading the feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SYNCHRONIZE SHORTCUT
'------------------------------------------------------------------------------
    'Register or remove the keyboard shortcut according to the setting
        If gDP_EnableKeyboardShortcut Then
            M_KeyboardShortcut_Register
        Else
            M_KeyboardShortcut_Remove
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

Public Sub M_KeyboardShortcut_Register()

'
'------------------------------------------------------------------------------
'                         REGISTER KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Registers the DatePicker keyboard shortcut for the current Excel session
'
' WHY THIS EXISTS
'   Application.OnKey requires an explicit public macro callback and the
'   assignment is not persisted by Excel across sessions
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns Ctrl + Shift + D to DP_OpenForActiveCell
'
' ERROR POLICY
'   Raises a descriptive runtime error if the shortcut cannot be registered
'
' DEPENDENCIES
'   Application.OnKey
'   M_GetQualifiedMacroName
'   DP_OpenForActiveCell
'
' NOTES
'   The callback is workbook-qualified so the shortcut resolves to this project
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_KeyboardShortcut_Register"   'Current procedure name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REGISTER SHORTCUT
'------------------------------------------------------------------------------
    'Assign Ctrl + Shift + D to the DatePicker public launcher
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY, M_GetQualifiedMacroName("DP_OpenForActiveCell")

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

Public Sub M_KeyboardShortcut_Remove()

'
'------------------------------------------------------------------------------
'                         REMOVE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the DatePicker keyboard shortcut to Excel
'
' WHY THIS EXISTS
'   Application.OnKey assignments are application-wide. If the workbook or add-in
'   is closed, the shortcut must not remain bound to a macro that may no longer
'   exist
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores Ctrl + Shift + D to Excel's default handling
'
' ERROR POLICY
'   Best-effort cleanup. Does not raise outward
'
' DEPENDENCIES
'   Application.OnKey
'
' NOTES
'   Calling Application.OnKey with only the key argument restores normal behavior
'
' UPDATED
'   2026-05-02
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REMOVE SHORTCUT
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

    'Restore the key to Excel
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY

    'Restore normal error handling
        On Error GoTo 0

End Sub


'
'------------------------------------------------------------------------------
'
'                                  GRID ICON
'
'------------------------------------------------------------------------------
'

Public Sub M_GridIcon_SetPath(ByVal IconPath As String)

    'Store the optional icon path
        gDP_IconPath = Trim$(IconPath)
End Sub

Public Sub M_GridIcon_ShowOrMove(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         SHOW OR MOVE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Shows the DatePicker in-grid icon by reusing the existing worksheet shape
'   whenever possible
'
' WHY THIS EXISTS
'   SelectionChange can fire very frequently. Deleting and recreating the
'   worksheet icon on every eligible cell selection creates visible latency
'   because Excel must rebuild and repaint the drawing-layer shape
'
' INPUTS
'   TargetCell
'     Optional single-cell anchor. ActiveCell is used when omitted
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the target cell, exits and removes the icon when the grid-icon
'   feature is disabled, reuses a valid existing icon on the target worksheet,
'   deletes a tracked icon from a previous worksheet when necessary, and creates
'   the icon only when no reusable target-sheet shape exists
'
' ERROR POLICY
'   Best-effort UI routine. Suppresses failures, falls back to the cold-path
'   creation routine where possible, and writes diagnostics to the Immediate
'   Window when the fallback path is used
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_GridIcon_Remove
'   M_GridIcon_Create
'   M_GetQualifiedMacroName
'   DP_GRID_ICON_NAME
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'
' NOTES
'   This is the preferred high-frequency SelectionChange path
'
'   M_GridIcon_Create remains the cold-path creation routine
'
'   The manager refresh path must call this routine, not M_GridIcon_Create
'   directly, otherwise the reuse / move optimization is bypassed
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_ShowOrMove"      'Current procedure name
    Const ICON_SIZE             As Double = 24#                          'Icon width and height
    Const ICON_GAP              As Double = 5#                           'Gap between target cell and icon

    Dim AnchorCell              As Excel.Range                           'Resolved anchor cell
    Dim TargetSheet             As Excel.Worksheet                       'Worksheet receiving the icon
    Dim CandidateShape          As Excel.Shape                           'Existing reusable icon candidate
    Dim IconLeft                As Double                                'Icon left position
    Dim IconTop                 As Double                                'Icon top position
    Dim HasReusableIcon         As Boolean                               'True when an existing shape can be moved
    Dim ErrorNumber             As Long                                  'Captured error number
    Dim ErrorDescription        As String                                'Captured error description

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
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell

    'Otherwise use ActiveCell safely
        Else
            On Error Resume Next
            Set AnchorCell = Application.ActiveCell
            Err.Clear
            On Error GoTo FailSafe

        End If

    'Remove stale icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Remove
            Exit Sub
        End If

'------------------------------------------------------------------------------
' NORMALIZE MERGED CELL TARGETS
'------------------------------------------------------------------------------
    'Normalize merged cells before rejecting multi-cell merged ranges
        If AnchorCell.MergeCells Then
            Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
        End If

    'Remove stale icon and exit when the target is still not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Remove
            Exit Sub
        End If

    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet

    'Remove stale icon and exit when no worksheet is available
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

'------------------------------------------------------------------------------
' RESOLVE TRACKED ICON
'------------------------------------------------------------------------------
    'Suppress stale shape-reference failures
        On Error Resume Next

    'Start from the tracked icon reference
        Set CandidateShape = gDP_GridIconShape

    'Clear invalid tracked references
        If Err.Number <> 0 Then
            Err.Clear
            Set CandidateShape = Nothing
            Set gDP_GridIconShape = Nothing
        End If

    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' HANDLE TRACKED ICON FROM ANOTHER WORKSHEET
'------------------------------------------------------------------------------
    'Validate the tracked icon parent when a candidate exists
        If Not CandidateShape Is Nothing Then

            'Suppress stale shape-parent failures
                On Error Resume Next

            'Reuse only when the tracked icon already belongs to the target sheet
                HasReusableIcon = (CandidateShape.Parent Is TargetSheet)

            'Delete a valid tracked icon that belongs to a previous worksheet
                If Err.Number = 0 Then
                    If Not HasReusableIcon Then
                        CandidateShape.Delete
                        Set CandidateShape = Nothing
                        Set gDP_GridIconShape = Nothing
                    End If
                End If

            'Clear stale candidate references
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CandidateShape = Nothing
                    Set gDP_GridIconShape = Nothing
                    HasReusableIcon = False
                End If

            'Restore fail-safe error handling
                On Error GoTo FailSafe

        End If

'------------------------------------------------------------------------------
' RESOLVE SAME-SHEET FALLBACK ICON
'------------------------------------------------------------------------------
    'Try to reuse a same-named shape on the target worksheet
        If CandidateShape Is Nothing Then

            'Suppress missing-shape failures
                On Error Resume Next

            'Resolve the existing same-named shape from the target worksheet
                Set CandidateShape = TargetSheet.Shapes(DP_GRID_ICON_NAME)

            'Use the same-named shape when it exists
                HasReusableIcon = Not (CandidateShape Is Nothing)

            'Clear missing-shape errors
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CandidateShape = Nothing
                    HasReusableIcon = False
                End If

            'Restore fail-safe error handling
                On Error GoTo FailSafe

        End If

'------------------------------------------------------------------------------
' MOVE EXISTING ICON
'------------------------------------------------------------------------------
    'Move and show the existing icon when it can be reused
        If HasReusableIcon Then

            With CandidateShape
                .Left = IconLeft
                .Top = IconTop
                .Width = ICON_SIZE
                .Height = ICON_SIZE
                .Placement = xlMove
                .AlternativeText = "DatePicker Grid Entry Point"
                .OnAction = M_GetQualifiedMacroName("DP_Click")
                .Visible = msoTrue
                .ZOrder msoBringToFront
            End With

            'Store the reusable icon reference
                Set gDP_GridIconShape = CandidateShape

            'Exit after moving the icon
                Exit Sub

        End If

'------------------------------------------------------------------------------
' CREATE ICON ONLY WHEN NEEDED
'------------------------------------------------------------------------------
    'Create the icon only when no reusable shape exists
        M_GridIcon_Create AnchorCell

'------------------------------------------------------------------------------
' NORMAL EXIT
'------------------------------------------------------------------------------
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the error number
        ErrorNumber = Err.Number

    'Capture the error description
        ErrorDescription = Err.Description

    'Suppress fallback failures
        On Error Resume Next

    'Clear invalid tracked references
        Set gDP_GridIconShape = Nothing

    'Fall back to cold-path creation when an anchor cell is available
        If Not AnchorCell Is Nothing Then
            M_GridIcon_Create AnchorCell
        End If

    'Write diagnostics without interrupting worksheet interaction
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription

    'Clear any suppressed fallback error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

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

Public Sub M_FormBridge_RefreshFromCell(ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                           FORM BRIDGE REFRESH FROM CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the visible DatePicker form from an explicit worksheet cell
'
' WHY THIS EXISTS
'   The manager already resolves the authoritative target cell. The form bridge
'   must therefore refresh from that same cell instead of independently reading
'   Excel.Application.ActiveCell
'
' INPUTS
'   TargetCell
'     Explicit worksheet cell to use as the DatePicker context
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Refreshes the existing visible DatePicker form from the supplied cell
'
' ERROR POLICY
'   Best-effort. Does not raise outward
'
' DEPENDENCIES
'   Existing M_FormBridge form-refresh implementation
'
' NOTES
'   This routine is the correct entry point for cDatePickerManager
'
'   Do not read Excel.Application.ActiveCell inside this routine
'
' UPDATED
'   2026-05-03
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_FormBridge_RefreshFromCell"

    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect caller from bridge failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit if no target cell was supplied
        If TargetCell Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Move the existing implementation of M_FormBridge_RefreshFromCell here
    'and replace every ActiveCell reference with TargetCell
        '
        ' Example of the intended internal policy:
        '
        '   - use TargetCell.Worksheet as the write-back worksheet
        '   - use TargetCell.Address as the write-back address
        '   - use TargetCell.Value / Value2 as the current cell value
        '   - use TargetCell.NumberFormat as the display/date-format context
        '   - refresh the already loaded DatePicker form from that explicit cell
        '
        ' Do not call Excel.Application.ActiveCell here

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number

    'Capture the error description
        ErrorDescription = Err.Description

    'Write bridge diagnostics without interrupting caller flow
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription

    'Clear the suppressed bridge error
        Err.Clear

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
'==============================================================================
'                       FORM BRIDGE UNLOAD LOADED PICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Unloads any already-loaded DatePicker UserForm instance before DP_Show loads
'   a fresh instance
'
' WHY THIS EXISTS
'   DP_Show relies on UserForm_Initialize to rebuild the runtime UI and consume
'   the pending initial-date bridge state. If a previous DatePicker instance is
'   already loaded, it should be removed before the new instance is loaded
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Scans the loaded VBA.UserForms collection, collects matching DatePicker form
'   instances, stops the DatePicker timer only when a picker instance exists, and
'   unloads matching forms on a best-effort basis
'
' ERROR POLICY
'   Best-effort cleanup. Does not raise outward
'
' DEPENDENCIES
'   VBA.UserForms
'   DP_FORM_NAME
'   M_Timer_Stop
'
' NOTES
'   This routine deliberately avoids direct references to UF_DatePicker to avoid
'   accidentally loading the default form instance
'
'   The scan and the unload phases are separated because unloading a form changes
'   the VBA.UserForms collection
'
'   DP_Show must not fail just because an old transient form instance cannot be
'   unloaded cleanly
'
' UPDATED
'   2026-05-04
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_FormBridge_UnloadLoadedPicker"

    Dim CurForm                 As Object       'Current loaded UserForm instance
    Dim LoadedForm              As Object       'DatePicker form selected for unload
    Dim FormsToUnload           As Collection   'Matching DatePicker forms to unload
    Dim Index                   As Long         'Collection reverse-loop index
    Dim FormTypeName            As String       'Loaded form class name
    Dim FormRuntimeName         As String       'Loaded form runtime name
    Dim StepName                As String       'Current diagnostic step
    Dim StepErrNumber           As Long         'Captured non-fatal cleanup error number
    Dim StepErrDescription      As String       'Captured non-fatal cleanup error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect DP_Show from all unload helper failures
        On Error GoTo FailSafe

    'Create the collection used to decouple scan and unload phases
        Set FormsToUnload = New Collection

'------------------------------------------------------------------------------
' SCAN LOADED USERFORMS
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Scan loaded UserForms"

    'Loop through currently loaded UserForms without unloading during enumeration
        For Each CurForm In VBA.UserForms

            'Reset form metadata
                FormTypeName = vbNullString
                FormRuntimeName = vbNullString

            'Suppress metadata-read errors for unusual transient form states
                On Error Resume Next

            'Capture the loaded form class name
                FormTypeName = VBA.TypeName(CurForm)

            'Capture the loaded form runtime name
                FormRuntimeName = VBA.CStr(CurForm.Name)

            'Capture metadata-read failure if any
                StepErrNumber = Err.Number
                StepErrDescription = Err.Description

            'Clear any suppressed metadata error
                Err.Clear

            'Restore fail-safe handling
                On Error GoTo FailSafe

            'Write metadata diagnostics only when a read failed
                If StepErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=" & StepName & _
                        " | Error=" & VBA.CStr(StepErrNumber) & _
                        " | " & StepErrDescription
                End If

            'Collect matching DatePicker forms by class name or runtime name
                If VBA.StrComp(FormTypeName, DP_FORM_NAME, vbTextCompare) = 0 Or _
                   VBA.StrComp(FormRuntimeName, DP_FORM_NAME, vbTextCompare) = 0 Then
                    FormsToUnload.Add CurForm
                End If

        Next CurForm

'------------------------------------------------------------------------------
' EXIT WHEN NO PICKER IS LOADED
'------------------------------------------------------------------------------
    'Exit cleanly when there is no loaded DatePicker form
        If FormsToUnload.Count = 0 Then GoTo CleanExit

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Stop timer"

    'Suppress timer-stop errors because Application.OnTime cancellation can fail
        On Error Resume Next

    'Stop any active DatePicker timer before unloading the form
        M_Timer_Stop

    'Capture timer-stop failure if any
        StepErrNumber = Err.Number
        StepErrDescription = Err.Description

    'Clear any suppressed timer-stop error
        Err.Clear

    'Restore fail-safe handling
        On Error GoTo FailSafe

    'Write timer-stop diagnostics only when stop failed
        If StepErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=" & StepName & _
                " | Error=" & VBA.CStr(StepErrNumber) & _
                " | " & StepErrDescription
        End If

'------------------------------------------------------------------------------
' UNLOAD MATCHING FORMS
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Unload DatePicker forms"

    'Unload collected DatePicker forms in reverse order
        For Index = FormsToUnload.Count To 1 Step -1

            'Resolve the collected form reference
                Set LoadedForm = FormsToUnload.Item(Index)

            'Suppress individual unload errors
                On Error Resume Next

            'Hide the form first to reduce visual flicker
                LoadedForm.Hide

            'Clear any non-fatal hide error
                Err.Clear

            'Unload the DatePicker form
                Unload LoadedForm

            'Capture unload failure if any
                StepErrNumber = Err.Number
                StepErrDescription = Err.Description

            'Clear any suppressed unload error
                Err.Clear

            'Restore fail-safe handling
                On Error GoTo FailSafe

            'Write unload diagnostics only when unload failed
                If StepErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=" & StepName & _
                        " | Index=" & VBA.CStr(Index) & _
                        " | Error=" & VBA.CStr(StepErrNumber) & _
                        " | " & StepErrDescription
                End If

            'Release the loop form reference
                Set LoadedForm = Nothing

        Next Index

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release the current form reference
        Set CurForm = Nothing

    'Release the loaded form reference
        Set LoadedForm = Nothing

    'Release the collection reference
        Set FormsToUnload = Nothing

    'Clear any suppressed cleanup error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Write unload-helper diagnostics without interrupting DP_Show
        Debug.Print PROC_NAME & _
            " | Step=" & StepName & _
            " | Error=" & VBA.CStr(Err.Number) & _
            " | " & Err.Description

    'Continue through clean exit
        Resume CleanExit

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




