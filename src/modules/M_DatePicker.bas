Attribute VB_Name = "M_DatePicker"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_DatePicker
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
'   2026-05-06
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
                
            Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long

            Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As LongPtr, _
                ByVal wMsg As Long, _
                ByVal wParam As LongPtr, _
                ByVal lParam As LongPtr) As LongPtr
    
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
            
            Private Declare Function ReleaseCapture Lib "user32" () As Long

            Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
        #End If
    #End If

    #If Mac Then
    #Else
        #If VBA7 Then
    
            Private Declare PtrSafe Sub SetLastError Lib "kernel32" ( _
                ByVal dwErrCode As Long)
        #Else
    
            Private Declare Sub SetLastError Lib "kernel32" ( _
                ByVal dwErrCode As Long)
        #End If
    #End If

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_SETTINGS_SECTION_DISPLAY      As String = "Display"                 'Display settings section
    Private Const DP_SETTINGS_SECTION_BEHAVIOR     As String = "Behavior"                'Behavior settings section
    Private Const DP_SETTINGS_SECTION_FEATURES     As String = "Features"                'Feature settings section
    Private Const DP_SETTINGS_SECTION_ADVANCED     As String = "Advanced"                'Advanced settings section

    Private Const DP_SETTING_FIRST_DAY_OF_WEEK     As String = "FirstDayOfWeek"          'First-day setting key
    Private Const DP_SETTING_USE_LOCAL_NAMES       As String = "UseLocalNames"           'Local names setting key
    Private Const DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION   As String = "AllowOutsideMonth"       'Outside-month setting key
    Private Const DP_SETTING_HIGHLIGHT_WEEKENDS    As String = "HighlightWeekends"       'Highlight weekends setting key
    Private Const DP_SETTING_CLOSE_AFTER_SELECTION As String = "CloseAfterSelection"     'Close-after-selection setting key
    Private Const DP_SETTING_CLOCK_MODE            As String = "ClockMode"               'Clock mode setting key
    Private Const DP_SETTING_SIZE_MODE             As String = "SizeMode"                'Size mode setting key
    Private Const DP_SETTING_SHOW_RIGHT_CLICK      As String = "ShowRightClick"          'Right-click setting key
    Private Const DP_SETTING_SHOW_GRID_ICON        As String = "ShowGridIcon"            'Grid-icon setting key
    Private Const DP_SETTING_ENABLE_KEYBOARD       As String = "EnableKeyboardShortcut"  'Keyboard shortcut setting key
    Private Const DP_SETTING_USE_WINAPI            As String = "UseWinAPI"               'WinAPI setting key
    Private Const DP_SETTING_HOLIDAY_CALLBACK      As String = "HolidayCallback"         'Holiday callback setting key
    
    Private Const DP_KEYBOARD_SHORTCUT_KEY         As String = "^+d"                     'Ctrl + Shift + D

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

    Private Const WS_CAPTION                       As Long = &HC00000                    'Window caption style flag
    Private Const DP_DEMO_SHEET_NAME               As String = "DATE PICKER DEMO"        'Demo worksheet name
    
    Private Const DP_WM_NCLBUTTONDOWN              As Long = &HA1                        'Non-client left-button down message
    Private Const DP_HTCAPTION                     As Long = 2                          'Title-bar hit-test code

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const DP_FORM_NAME                      As String = "UF_DatePicker"           'DatePicker UserForm name

    Public Const DP_MIN_YEAR                       As Long = 100                         'Minimum supported DatePicker year
    Public Const DP_MAX_YEAR                       As Long = 9999                        'Maximum supported DatePicker year


    Public Const DP_SETTINGS_APP_NAME              As String = "VBA_DATETIMEPICKER"      'Legacy registry application name
    Public Const DP_GRID_ICON_NAME                 As String = "DP_GridIcon"             'Worksheet grid icon shape name
    Public Const DP_MSGBOX_TITLE                   As String = "Date / Time Picker"      'Message-box title

    Public Const DP_DEFAULT_FIRST_DAY_OF_WEEK      As Long = vbMonday                    'Default first day of week
    Public Const DP_DEFAULT_USE_LOCAL_NAMES        As Boolean = True                     'Default local names flag
    Public Const DP_DEFAULT_ALLOW_OUTSIDE_MONTH_SELECTION    As Boolean = True           'Default outside-month selection flag
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
        DP_WriteAction_DatePicker = 0
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
    Public gDP_FirstDayOfWeek           As Long                 'vbSunday or vbMonday
    Public gDP_UseLocalNames            As Boolean              'Use local day/month captions
    Public gDP_ClockMode                As DP_ClockMode         'Static or live clock
    Public gDP_SizeMode                 As DP_SizeMode          'Normal or compact layout
    Public gDP_HighlightWeekends        As Boolean              'Highlight weekend days
    Public gDP_AllowOutsideMonthSelection As Boolean            'Allow outside-month day selection
    
    Public gDP_CloseAfterSelection      As Boolean              'Close picker after successful write-back
    Public gDP_UseWinAPI                As Boolean              'Allow optional WinAPI-dependent native behavior
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
    
    Private mQualifiedMacroNameCache    As Object               'Scripting.Dictionary: ProcedureName -> qualified macro name
    Private mQualifiedMacroWorkbookName As String               'Workbook name used by the qualified macro-name cache
    
    Private mDP_GridIconLastAnchorKey       As String           'Last grid-icon target key
    Private mDP_GridIconLastLeft            As Double           'Last grid-icon left position
    Private mDP_GridIconLastTop             As Double           'Last grid-icon top position
    Private mDP_TimerProcedureWorkbookName  As String           'Workbook name used for cached timer callback

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
            DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                DP_SETTINGS_APP_NAME, _
                DP_SETTINGS_SECTION_DISPLAY, _
                DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION, _
                M_Settings_BooleanToStorageValue(DP_DEFAULT_ALLOW_OUTSIDE_MONTH_SELECTION))
        End If
    'Store the parsed outside-month setting or its default
        If M_Settings_TryParseBoolean(RawValue, ParsedBoolean) Then
            gDP_AllowOutsideMonthSelection = ParsedBoolean
        Else
            gDP_AllowOutsideMonthSelection = DP_DEFAULT_ALLOW_OUTSIDE_MONTH_SELECTION
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
            DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION, _
            M_Settings_BooleanToStorageValue(gDP_AllowOutsideMonthSelection)

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
        gDP_AllowOutsideMonthSelection = DP_DEFAULT_ALLOW_OUTSIDE_MONTH_SELECTION
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

Public Sub M_Settings_SetUseLocalNames(ByVal UseLocalDayNames As Boolean)

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
    Const PROC_NAME         As String = "M_Settings_SetUseLocalNames"

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

Public Sub M_Settings_SetEnableKeyboardShortcut(ByVal EnableKeyboardShortcut As Boolean)

'
'------------------------------------------------------------------------------
'                       SET ENABLE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether the DatePicker keyboard shortcut is enabled
'
' WHY THIS EXISTS
'   The keyboard shortcut is the safe fallback entry point when optional visual
'   integrations such as the in-grid icon and right-click menu are disabled
'
'   Caller code, demo sheets, settings panels, and host workbooks need one
'   controlled public entry point to update keyboard shortcut integration without
'   mutating public DatePicker state inconsistently
'
' INPUTS
'   EnableKeyboardShortcut
'     True to enable keyboard shortcut integration
'     False to disable keyboard shortcut integration
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded
'   Applies the requested keyboard shortcut setting
'   Forces keyboard shortcut access on when both right-click menu access and
'   in-grid icon access are disabled
'   Persists the updated settings when needed
'   Synchronizes Application.OnKey with the final setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, settings
'   cannot be saved, or keyboard shortcut integration cannot be synchronized
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
'   If synchronization fails after persistence succeeds, the saved setting is not
'   rolled back because the user preference has already been committed
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_KeyboardShortcut_Update
'   gDP_EnableKeyboardShortcut
'   gDP_ShowRightClick
'   gDP_ShowGridIcon
'
' NOTES
'   Disabling right-click menu access, in-grid icon access, and keyboard shortcut
'   access at the same time would leave the user without a practical manual
'   DatePicker entry point
'
'   For this reason, a request to disable the keyboard shortcut is normalized
'   back to True when both other manual access paths are disabled
'
'   The keyboard shortcut is synchronized even when the setting is unchanged
'   because Application.OnKey is session state and may have been reset by Excel,
'   another add-in, workbook activation, or teardown logic
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetEnableKeyboardShortcut"

    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting
    Dim SettingsChanged         As Boolean      'True when the resolved setting differs
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
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current setting"
    'Capture the current keyboard shortcut setting
        OldEnableKeyboard = gDP_EnableKeyboardShortcut
    'Initialize the resolved keyboard shortcut setting from the requested value
        NewEnableKeyboard = EnableKeyboardShortcut

'------------------------------------------------------------------------------
' PROTECT MANUAL ACCESS PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve access fallback"
    'Keep keyboard access enabled when both visual entry points are disabled
        If Not NewEnableKeyboard Then
            If Not gDP_ShowRightClick Then
                If Not gDP_ShowGridIcon Then
                    NewEnableKeyboard = True
                End If
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE CHANGE FLAG
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flag"
    'Resolve whether the final keyboard shortcut setting changed
        SettingsChanged = (NewEnableKeyboard <> OldEnableKeyboard)

'------------------------------------------------------------------------------
' SYNCHRONIZE ONLY WHEN SETTING IS UNCHANGED
'------------------------------------------------------------------------------
    'Synchronize Application.OnKey and exit when no persisted setting changed
        If Not SettingsChanged Then
            'Track the current handler step
                HandlerStep = "Synchronize keyboard shortcut"
            'Synchronize keyboard shortcut integration with the current setting
                M_KeyboardShortcut_Update
            'Exit because no registry write is required
                Exit Sub
        End If

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory setting"
    'Store the resolved keyboard shortcut setting
        gDP_EnableKeyboardShortcut = NewEnableKeyboard
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
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"
    'Synchronize keyboard shortcut integration with the saved setting
        M_KeyboardShortcut_Update

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
        If SettingsMutated And Not SettingsPersisted Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous keyboard shortcut setting
                gDP_EnableKeyboardShortcut = OldEnableKeyboard
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard shortcut setting update failed: " & ErrorDescription

End Sub

Public Sub M_Settings_SetUseWinAPI(ByVal UseWinAPI As Boolean)

'
'------------------------------------------------------------------------------
'                           SET USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether optional WinAPI-dependent native behavior is enabled
'
' WHY THIS EXISTS
'   Some DatePicker features, such as borderless styling, native window-handle
'   operations, and mouse-based positioning, require Windows API calls
'
'   Caller code, demo sheets, settings panels, and host workbooks need one
'   controlled public entry point to update the WinAPI policy without mutating
'   public DatePicker state inconsistently
'
' INPUTS
'   UseWinAPI
'     True to enable optional WinAPI-dependent native behavior when the platform
'     supports it
'
'     False to disable optional WinAPI-dependent native behavior
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded
'   Resolves current platform WinAPI capability
'   Normalizes the requested setting to False when WinAPI is not supported on the
'   current platform
'   Exits when the normalized setting is already active
'   Updates the in-memory WinAPI setting
'   Persists the updated settings
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
'   M_Platform_CanUseWinAPI
'   gDP_UseWinAPI
'
' NOTES
'   This setter controls the user / project preference only
'
'   The effective runtime policy remains M_Platform_ShouldUseWinAPI, which
'   combines platform capability and gDP_UseWinAPI
'
'   A request to enable WinAPI on Mac is normalized to False rather than saved as
'   True, because the current platform cannot execute the native API calls
'
'   This routine does not directly restyle, reposition, unload, or rebuild an
'   already-loaded DatePicker form
'
'   Visual WinAPI-dependent effects are applied by the relevant window helper
'   routines when they are next called
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetUseWinAPI"

    Dim OldUseWinAPI            As Boolean      'Previous WinAPI setting
    Dim NewUseWinAPI            As Boolean      'Resolved WinAPI setting
    Dim PlatformCanUseWinAPI    As Boolean      'True when WinAPI helpers are supported
    Dim SettingsChanged         As Boolean      'True when the resolved setting differs
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
' LOAD CURRENT SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Ensure existing settings are loaded before changing one value
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' RESOLVE PLATFORM CAPABILITY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve platform capability"
    'Resolve whether WinAPI helpers can be used on the current platform
        PlatformCanUseWinAPI = M_Platform_CanUseWinAPI

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture current setting"
    'Capture the current WinAPI setting
        OldUseWinAPI = gDP_UseWinAPI

'------------------------------------------------------------------------------
' NORMALIZE REQUESTED SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize requested setting"
    'Initialize the resolved WinAPI setting from the requested value
        NewUseWinAPI = UseWinAPI
    'Force WinAPI off when the current platform does not support it
        If Not PlatformCanUseWinAPI Then NewUseWinAPI = False

'------------------------------------------------------------------------------
' EXIT IF UNCHANGED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flag"
    'Resolve whether the final WinAPI setting changed
        SettingsChanged = (NewUseWinAPI <> OldUseWinAPI)
    'Exit when no setting change is required
        If Not SettingsChanged Then Exit Sub

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Update in-memory setting"
    'Store the resolved WinAPI setting
        gDP_UseWinAPI = NewUseWinAPI
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
        If SettingsMutated And Not SettingsPersisted Then
            'Suppress rollback errors
                On Error Resume Next
            'Restore the previous WinAPI setting
                gDP_UseWinAPI = OldUseWinAPI
            'Restore normal error handling
                On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "WinAPI setting update failed: " & ErrorDescription

End Sub
Public Sub M_Settings_SetAllowOutsideMonthSelection(ByVal AllowOutsideMonthSelection As Boolean)

'
'------------------------------------------------------------------------------
'                       SET ALLOW OUTSIDE MONTH SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Sets and saves whether outside-month days can be selected
'
' WHY THIS EXISTS
'   Some workflows allow quick adjacent-month selection from leading or trailing
'   calendar cells, while others restrict users to the displayed month
'
' INPUTS
'   AllowOutsideMonthSelection
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
    Const PROC_NAME         As String = "M_Settings_SetAllowOutsideMonthSelection"

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
        If gDP_AllowOutsideMonthSelection = AllowOutsideMonthSelection Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE CURRENT SETTING
'------------------------------------------------------------------------------
    'Capture the current outside-month setting for rollback
        OldAllowOutside = gDP_AllowOutsideMonthSelection

'------------------------------------------------------------------------------
' UPDATE IN-MEMORY SETTING
'------------------------------------------------------------------------------
    'Store the requested outside-month setting
        gDP_AllowOutsideMonthSelection = AllowOutsideMonthSelection
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
            gDP_AllowOutsideMonthSelection = OldAllowOutside
            On Error GoTo 0
        End If
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "Allow-outside-month-selection setting update failed: " & ErrorDescription

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
'   Caller code, demo sheets, Ribbon callbacks, settings panels, and host
'   workbooks need one controlled public entry point to update this setting
'   consistently
'
' INPUTS
'   HighlightWeekends
'     True to highlight weekend days
'     False to display weekends with normal day-cell styling
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures current settings are loaded
'   Exits when the requested value is already active
'   Updates the in-memory weekend-highlight setting
'   Persists the updated settings
'   Delegates loaded-form synchronization to M_FormBridge_RefreshSettings
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, saved, or
'   synchronized with the loaded DatePicker form
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory setting is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'   UF_DatePicker.UF_DP_RefreshSettings
'
' NOTES
'   This routine owns validation, persistence, and rollback only
'
'   This routine does not directly repaint calendar cells, rebuild the day grid,
'   or apply weekend styling to the loaded UserForm
'
'   Loaded-form synchronization is delegated through M_FormBridge_RefreshSettings
'
'   The actual weekend-highlight repaint must be implemented inside
'   UF_DatePicker.UF_DP_RefreshSettings
'
'   UF_DP_RefreshSettings must therefore apply any required visual refresh for
'   gDP_HighlightWeekends, including repainting or rebuilding day labels as
'   appropriate
'
'   If UF_DP_RefreshSettings only refreshes captions, then changing
'   HighlightWeekends while the form is already loaded will be persisted
'   correctly but may not be visually reflected until the form is reopened
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_SetHighlightWeekends" 'Current procedure name

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
' SYNCHRONIZE LOADED FORM
'------------------------------------------------------------------------------
    'Delegate loaded-form synchronization to the form bridge
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
'   display
'
'   Caller code, demo sheets, Ribbon callbacks, settings panels, and host
'   workbooks need one controlled public entry point to update this setting
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
'   Validates the requested size mode
'   Ensures current settings are loaded
'   Avoids unnecessary registry writes when the setting is unchanged
'   Updates the in-memory size-mode setting when required
'   Persists the updated settings
'   Delegates loaded-form synchronization to M_FormBridge_RefreshSettings
'
' ERROR POLICY
'   Raises a descriptive runtime error if SizeMode is unsupported, settings
'   cannot be loaded, settings cannot be saved, or the loaded DatePicker form
'   cannot be synchronized
'
'   If persistence fails after the in-memory setting was changed, the previous
'   in-memory size mode is restored before the error is re-raised
'
' DEPENDENCIES
'   M_Settings_IsValidSizeMode
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'   UF_DatePicker.UF_DP_RefreshSettings
'
' NOTES
'   This routine owns validation, persistence, and rollback only
'
'   This routine does not directly resize, unload, rebuild, or repaint the
'   DatePicker UserForm
'
'   Loaded-form synchronization is delegated through M_FormBridge_RefreshSettings
'
'   The actual size-mode application must be implemented inside
'   UF_DatePicker.UF_DP_RefreshSettings
'
'   UF_DP_RefreshSettings must therefore apply any required layout changes for
'   gDP_SizeMode, including resizing, rebuilding, hiding, showing, or repainting
'   controls as appropriate
'
'   If UF_DP_RefreshSettings only refreshes captions, then changing SizeMode
'   while the form is already loaded will be persisted correctly but may not be
'   visually reflected until the form is reopened
'
' UPDATED
'   2026-05-06
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
' SYNCHRONIZE LOADED FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize loaded DatePicker form"
    'Delegate loaded-form synchronization to the form bridge
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

Public Function M_Settings_GetUseLocalNames() As Boolean

'
'------------------------------------------------------------------------------
'                           GET USE LOCAL NAMES
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
    Const PROC_NAME            As String = "M_Settings_GetUseLocalNames"

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
        M_Settings_GetUseLocalNames = gDP_UseLocalNames

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

Public Function M_Settings_GetAllowOutsideMonthSelection() As Boolean

'
'------------------------------------------------------------------------------
'                   GET ALLOW OUTSIDE MONTH SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether outside-month days can be selected in the DatePicker
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
'   True when outside-month day selection is enabled
'
'   False when outside-month day selection is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current outside-month selection
'   setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_AllowOutsideMonthSelection
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_GetAllowOutsideMonthSelection"

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
    'Return the current outside-month selection setting
        M_Settings_GetAllowOutsideMonthSelection = gDP_AllowOutsideMonthSelection

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
            "Outside-month selection setting retrieval failed: " & Err.Description

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


Public Function M_Settings_GetShowRightClick() As Boolean

'
'------------------------------------------------------------------------------
'                           GET SHOW RIGHT-CLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether DatePicker right-click menu integration is enabled
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
'   True when right-click menu integration is enabled
'
'   False when right-click menu integration is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current right-click menu setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_ShowRightClick
'
' NOTES
'   This routine does not read the registry directly
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_GetShowRightClick"

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
    'Return the current right-click menu setting
        M_Settings_GetShowRightClick = gDP_ShowRightClick

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
            "Right-click menu setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetShowGridIcon() As Boolean

'
'------------------------------------------------------------------------------
'                           GET SHOW GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether DatePicker in-grid worksheet icon integration is enabled
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
'   True when in-grid icon integration is enabled
'
'   False when in-grid icon integration is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current in-grid icon setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_ShowGridIcon
'
' NOTES
'   This routine does not create, move, or remove the worksheet icon
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_GetShowGridIcon"

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
    'Return the current in-grid icon setting
        M_Settings_GetShowGridIcon = gDP_ShowGridIcon

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
            "In-grid icon setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetEnableKeyboardShortcut() As Boolean

'
'------------------------------------------------------------------------------
'                       GET ENABLE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker keyboard shortcut is enabled
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
'   True when keyboard shortcut integration is enabled
'
'   False when keyboard shortcut integration is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current keyboard shortcut setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_EnableKeyboardShortcut
'
' NOTES
'   This routine reports the setting only
'
'   It does not register or remove the Application.OnKey binding
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_GetEnableKeyboardShortcut"

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
    'Return the current keyboard shortcut setting
        M_Settings_GetEnableKeyboardShortcut = gDP_EnableKeyboardShortcut

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
            "Keyboard shortcut setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           GET USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether optional WinAPI-dependent native behavior is enabled by
'   DatePicker settings
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
'   True when optional WinAPI-dependent behavior is enabled by setting
'
'   False when optional WinAPI-dependent behavior is disabled by setting or by
'   platform normalization
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current WinAPI setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_UseWinAPI
'
' NOTES
'   This routine returns the normalized setting state
'
'   Use M_Platform_ShouldUseWinAPI when the caller needs the effective runtime
'   policy combining platform capability and setting state
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_GetUseWinAPI"

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
    'Return the current WinAPI setting
        M_Settings_GetUseWinAPI = gDP_UseWinAPI

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
            "WinAPI setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetClockMode() As DP_ClockMode

'
'------------------------------------------------------------------------------
'                           GET CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current DatePicker clock-mode setting
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
'   Current DatePicker clock mode
'
' BEHAVIOR
'   Ensures settings are loaded
'   Validates the in-memory clock-mode setting
'   Returns the current clock-mode value
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or if the
'   loaded in-memory clock mode is unsupported
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidClockMode
'   gDP_ClockMode
'
' NOTES
'   This routine does not start, stop, or resynchronize the live-clock timer
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_GetClockMode"

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
    'Reject unsupported clock modes
        If Not M_Settings_IsValidClockMode(gDP_ClockMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "gDP_ClockMode is unsupported"
        End If

'------------------------------------------------------------------------------
' RETURN SETTING
'------------------------------------------------------------------------------
    'Return the current clock-mode setting
        M_Settings_GetClockMode = gDP_ClockMode

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
            "Clock-mode setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetSizeMode() As DP_SizeMode

'
'------------------------------------------------------------------------------
'                           GET SIZE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current DatePicker size-mode setting
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
'   Current DatePicker size mode
'
' BEHAVIOR
'   Ensures settings are loaded
'   Validates the in-memory size-mode setting
'   Returns the current size-mode value
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded or if the
'   loaded in-memory size mode is unsupported
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidSizeMode
'   gDP_SizeMode
'
' NOTES
'   This routine reports the setting only
'
'   It does not resize, rebuild, or repaint the loaded DatePicker form
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_GetSizeMode"
    
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
    'Reject unsupported size modes
        If Not M_Settings_IsValidSizeMode(gDP_SizeMode) Then
            Err.Raise vbObjectError + 513, PROC_NAME, "gDP_SizeMode is unsupported"
        End If

'------------------------------------------------------------------------------
' RETURN SETTING
'------------------------------------------------------------------------------
    'Return the current size-mode setting
        M_Settings_GetSizeMode = gDP_SizeMode

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
            "Size-mode setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetHighlightWeekends() As Boolean

'
'------------------------------------------------------------------------------
'                           GET HIGHLIGHT WEEKENDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether weekend days are visually highlighted in the DatePicker
'   calendar grid
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
'   True when weekend highlighting is enabled
'
'   False when weekend highlighting is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current weekend-highlight setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_HighlightWeekends
'
' NOTES
'   This routine reports the setting only
'
'   It does not repaint or rebuild the loaded DatePicker form
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_GetHighlightWeekends"

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
    'Return the current weekend-highlight setting
        M_Settings_GetHighlightWeekends = gDP_HighlightWeekends

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
            "Weekend-highlight setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetCloseAfterSelection() As Boolean

'
'------------------------------------------------------------------------------
'                         GET CLOSE AFTER SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker closes after a successful write-back
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
'   True when the picker closes after successful write-back
'
'   False when the picker remains open after successful write-back
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current close-after-selection
'   setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_CloseAfterSelection
'
' NOTES
'   This routine reports the setting only
'
'   The setting is consumed during the successful write-back lifecycle
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_GetCloseAfterSelection"

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
    'Return the current close-after-selection setting
        M_Settings_GetCloseAfterSelection = gDP_CloseAfterSelection

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
            "Close-after-selection setting retrieval failed: " & Err.Description

End Function

Public Function M_Settings_GetHolidayCallback() As String

'
'------------------------------------------------------------------------------
'                           GET HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the configured DatePicker holiday-callback name
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
'   Current holiday-callback name
'
'   Blank when callback-based holiday logic is disabled
'
' BEHAVIOR
'   Ensures settings are loaded and returns the current holiday-callback setting
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_HolidayCallbackName
'
' NOTES
'   This routine reports the configured callback name only
'
'   It does not execute, validate, or inspect the callback procedure
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_GetHolidayCallback"

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
    'Return the current holiday-callback setting
        M_Settings_GetHolidayCallback = VBA.Trim$(gDP_HolidayCallbackName)

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
            "Holiday-callback setting retrieval failed: " & Err.Description

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
    Const PROC_NAME         As String = "M_Settings_FirstDayOfWeekToText"

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
'   False when RawValue is blank, unsupported, or cannot be parsed safely
'
' BEHAVIOR
'   Normalizes RawValue, recognizes canonical persisted values, VBA-style
'   constant names, English weekday names, assigns ParsedValue deterministically,
'   and returns the parse result
'
' ERROR POLICY
'   Does not raise errors
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

Public Sub M_Picker_EnsureManager( _
    Optional ByRef EventsDisabledByCaller As Boolean)

'
'------------------------------------------------------------------------------
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
'   EventsDisabledByCaller
'       Optional output flag. Set True when Application.EnableEvents was already
'       False on entry. Left False otherwise
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures persisted settings are loaded, reports the caller's Excel event state
'   without modifying it, creates the manager when missing, and recreates it when
'   the existing manager is not hooked
'
' ERROR POLICY
'   Raises a descriptive runtime error if settings cannot be loaded, or if the
'   manager cannot be instantiated or re-instantiated
'
'   Does not raise when Excel events are disabled. That is a valid caller state
'
' DEPENDENCIES
'   Excel.Application.EnableEvents
'   M_Settings_EnsureLoaded
'   cDatePickerManager
'   gDP_Manager
'
' NOTES
'   This routine is the canonical manager bootstrapper
'
'   The Is_Hooked check prevents a stale manager object from silently disabling
'   Application event handling
'
'   This routine does not modify Application.EnableEvents. A caller that
'   deliberately suppresses Excel events keeps that state across this call
'
'   A manager hooked while events are disabled is a valid, self-correcting state.
'   The WithEvents reference remains live and Excel resumes dispatching as soon
'   as the caller restores Application.EnableEvents = True
'
'   Is_Hooked reports the manager's own hook state and does not consult
'   Application.EnableEvents, so a suppressed-event session does not trigger
'   repeated manager recreation
'
'   DP_RepairRuntime is the only entry point that force-enables Excel events
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Picker_EnsureManager"

    Dim ManagerNeedsCreate      As Boolean      'True when manager must be created
    Dim ManagerNeedsRecreate    As Boolean      'True when manager exists but is not hooked
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure persisted DatePicker settings are available before manager startup
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' OBSERVE EXCEL EVENT STATE
'------------------------------------------------------------------------------
    'Report the caller's Excel event state without altering it
        EventsDisabledByCaller = Not Excel.Application.EnableEvents

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
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
        ErrorNumber = Err.Number

        ErrorDescription = Err.Description

        Err.Raise ErrorNumber, PROC_NAME, _
            "DatePicker manager initialization failed: " & ErrorDescription

End Sub

Public Sub DP_Preload()

'
'------------------------------------------------------------------------------
'                           PRELOAD DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Loads the DatePicker UserForm once and keeps it hidden for fast later display
'
' WHY THIS EXISTS
'   The first UserForm load is the slowest DatePicker interaction because the
'   runtime controls, hooks, fonts, settings panel, picker panel, and calendar
'   grid are created during UserForm_Initialize
'
'   Preloading the form during workbook startup moves that cost away from the
'   first user click so Ribbon, keyboard, right-click, and in-grid activation feel
'   immediate
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures DatePicker settings and manager infrastructure are available
'   Exits when a DatePicker form is already loaded
'   Initializes bridge state from today's date
'   Loads UF_DatePicker while hidden
'   Keeps the loaded form hidden
'   Stops any live-clock timer after preload so hidden forms do not keep ticking
'
' ERROR POLICY
'   Best-effort startup optimization
'   Does not raise outward
'   Writes diagnostics to the Immediate Window when preload fails
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   M_FormBridge_GetLoadedForm
'   M_Timer_Stop
'   UF_DatePicker
'   DP_FORM_NAME
'
' NOTES
'   This routine is an optimization only
'
'   DP_Show must still work if this preload routine was never called or failed
'
'   The form is loaded but not shown, so UserForm_Activate does not run
'
'   The visible-positioning logic remains inside DP_Show
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_Preload"

    Dim LoadedForm      As Object       'Already-loaded DatePicker form instance
    Dim ErrorNumber     As Long         'Captured preload error number
    Dim ErrorDescription As String      'Captured preload error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress preload failures through the fail-safe path
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' ENSURE RUNTIME
'------------------------------------------------------------------------------
    'Ensure settings and manager infrastructure are available
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' EXIT IF ALREADY LOADED
'------------------------------------------------------------------------------
    'Resolve an already-loaded DatePicker form without creating a new instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Exit when the DatePicker form is already loaded
        If Not LoadedForm Is Nothing Then GoTo CleanExit

'------------------------------------------------------------------------------
' INITIALIZE BRIDGE STATE
'------------------------------------------------------------------------------
    'Use today as the preload display date
        gDP_InitialDate = VBA.Date
    'Mark the preload initial date as available
        gDP_HasInitialDate = True
    'Clear selected-date state during preload
        gDP_SelectedDate = 0
    'Mark selected-date state as unavailable during preload
        gDP_HasSelectedDate = False

'------------------------------------------------------------------------------
' LOAD FORM HIDDEN
'------------------------------------------------------------------------------
    'Load the DatePicker form while keeping it hidden
        Load UF_DatePicker
    'Resolve the loaded form after initialization
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Hide the form defensively when it was loaded successfully
        If Not LoadedForm Is Nothing Then LoadedForm.Hide

'------------------------------------------------------------------------------
' STOP HIDDEN TIMER
'------------------------------------------------------------------------------
    'Stop timer activity after preload because the form is not visible yet
        M_Timer_Stop

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release local object reference
        Set LoadedForm = Nothing
    'Clear non-fatal preload errors
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the preload error number
        ErrorNumber = Err.Number
    'Capture the preload error description
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Stop timer activity if preload partially initialized it
        M_Timer_Stop
    'Release local object reference
        Set LoadedForm = Nothing
    'Write preload diagnostics without blocking workbook open
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the handled preload error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub DP_Start()

'
'------------------------------------------------------------------------------
'                           START DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the DatePicker manager and synchronizes the interactive Excel UI
'   integration points
'
' WHY THIS EXISTS
'   The manager is event-driven. After workbook open, VBA reset, code import, or
'   add-in reload, the manager must be explicitly bootstrapped before Excel
'   Application events can move or remove the in-grid icon
'
'   Startup must also synchronize right-click menu and keyboard shortcut
'   integration so all configured DatePicker entry points are available in the
'   current Excel session
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures the manager is alive and hooked
'   Synchronizes the Excel right-click menu according to settings
'   Synchronizes the keyboard shortcut according to settings
'   Evaluates the current ActiveCell so stale icons are cleaned and the correct
'   in-grid icon state is shown
'
' ERROR POLICY
'   Raises a descriptive runtime error if startup fails
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'   gDP_Manager.Handle_SelectionChange
'
' NOTES
'   Call this from Workbook_Open, Auto_Open, add-in startup, or manually after
'   importing the project into a workbook
'
'   M_ContextMenu_Update and M_KeyboardShortcut_Update are intentionally called
'   here because those integrations are session/UI state, not only persisted
'   settings state
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "DP_Start" 'Current procedure name

    Dim HandlerStep            As String       'Current handler step for diagnostics
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' ENSURE MANAGER
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure manager"
    'Ensure the DatePicker manager exists and Application events are hooked
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' SYNCHRONIZE RIGHT-CLICK MENU
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize right-click menu"
    'Synchronize the DatePicker right-click menu with current settings
        M_ContextMenu_Update

'------------------------------------------------------------------------------
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"
    'Synchronize the DatePicker keyboard shortcut with current settings
        M_KeyboardShortcut_Update

'------------------------------------------------------------------------------
' PRE-CREATE GRID ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Pre-create hidden grid icon"
    'Pre-create the grid icon for fast selection-change reuse
        M_GridIcon_PreCreateHidden
        
'------------------------------------------------------------------------------
' REFRESH CURRENT UI
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh current selection context"
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Raise a descriptive startup error
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker startup failed: " & ErrorDescription

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
'   The StepName diagnostic is intentionally retained because UserForm loading
'   can fail inside nested Initialize routines and otherwise surface only as a
'   generic DP_Show failure
'
' UPDATED
'   2026-05-04
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
    Dim StepName                    As String                'Current diagnostic step
    Dim ErrorNumber                 As Long                  'Captured error number
    Dim ErrorSource                 As String                'Captured error source
    Dim ErrorDescription            As String                'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Track the current step
        StepName = "Ensure manager"
'------------------------------------------------------------------------------
' ENSURE DATEPICKER INFRASTRUCTURE
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Ensure manager"
    'Ensure DatePicker settings and manager infrastructure are available
        M_Picker_EnsureManager

'------------------------------------------------------------------------------
' REUSE EXISTING VISIBLE FORM
'------------------------------------------------------------------------------
    'Reuse the already-visible DatePicker form when possible
        If M_FormBridge_TryReuseLoadedPickerFromActiveCell( _
            FORM_MOUSE_OFFSET_XPX, _
            FORM_MOUSE_OFFSET_YPX, _
            FORM_CENTER_ON_MOUSE) Then
            'Exit because the visible picker was refreshed instead of rebuilt
                Exit Sub
        End If
    'Default the initial date to today
        InitialDate = VBA.Date
    'Default selected-date availability to False
        HasCellDate = False
    'Default ActiveCell value availability to False
        HasActiveCellValue = False

'------------------------------------------------------------------------------
' READ ACTIVE CELL VALUE
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Read ActiveCell value"
    'Suppress ActiveCell access errors
        On Error Resume Next
    'Read ActiveCell value safely
        CellValue = Excel.Application.ActiveCell.Value
    'Store whether ActiveCell value was read successfully
        HasActiveCellValue = (Err.Number = 0)
    'Clear any suppressed ActiveCell access error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE INITIAL DATE FROM ACTIVE CELL
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Resolve initial date"
    'Use ActiveCell only when a value was read successfully
        If HasActiveCellValue Then
            'Ignore Excel error values
                If Not VBA.IsError(CellValue) Then
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
' RESOLVE EXISTING FORM INSTANCE
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Resolve existing DatePicker form"
    'Resolve an already-loaded DatePicker form without forcing default-instance creation
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

'------------------------------------------------------------------------------
' STORE FORM BRIDGE STATE
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Store bridge state"
    'Store the initial date for the next refresh or form instance
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
' LOAD FORM ONLY WHEN NEEDED
'------------------------------------------------------------------------------
    'Load the DatePicker form only when no reusable instance exists
        If LoadedForm Is Nothing Then
            'Track the current step
                StepName = "Load UF_DatePicker"
            'Load the DatePicker form while it is still hidden
                Load UF_DatePicker
            'Track the current step
                StepName = "Resolve loaded DatePicker form"
            'Resolve the loaded DatePicker form instance
                Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
            'Fallback to the default instance if the bridge did not resolve it
                If LoadedForm Is Nothing Then Set LoadedForm = UF_DatePicker
            'Reject unresolved DatePicker form instance
                If LoadedForm Is Nothing Then
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Unable to resolve loaded DatePicker form instance"
                End If
        Else
            'Track the current step
                StepName = "Refresh reused DatePicker form"
            'Refresh the already-loaded form from the current ActiveCell context
                LoadedForm.UF_DP_RefreshFromExternalSelection InitialDate, HasCellDate
        End If

'------------------------------------------------------------------------------
' SHOW FORM
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Show modeless form"
    'Show the loaded DatePicker form modelessly
        LoadedForm.Show vbModeless

'------------------------------------------------------------------------------
' APPLY CLOCK MODE
'------------------------------------------------------------------------------
    'Apply the configured clock mode after the form is loaded and visible
        M_Timer_ApplyClockMode

'------------------------------------------------------------------------------
' FINAL POSITION CORRECTION
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Final position correction"
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
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error source
        ErrorSource = Err.Source
    'Capture the error description
        ErrorDescription = Err.Description
    'Write a detailed diagnostic before re-raising
        Debug.Print PROC_NAME & _
            " | Step=" & StepName & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | Source=" & ErrorSource & _
            " | " & ErrorDescription
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "DatePicker show failed" & _
            " | Step=" & StepName & _
            " | Source=" & ErrorSource & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription

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
'------------------------------------------------------------------------------
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
'------------------------------------------------------------------------------

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
' SYNCHRONIZE RIGHT-CLICK MENU
'------------------------------------------------------------------------------
    'Synchronize the DatePicker right-click menu with current settings
        M_ContextMenu_Update

'------------------------------------------------------------------------------
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Synchronize the DatePicker keyboard shortcut with current settings
        M_KeyboardShortcut_Update

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

Public Sub DP_Hide()

'
'------------------------------------------------------------------------------
'                           HIDE DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the DatePicker form without unloading it
'
' WHY THIS EXISTS
'   Preloaded / reusable UserForm mode should keep the runtime controls alive so
'   subsequent DatePicker launches are immediate
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stops live-clock timer activity and hides the already-loaded picker form
'   without unloading it
'
' ERROR POLICY
'   Best-effort UI cleanup
'   Suppresses errors because hiding should not interrupt the user workflow
'
' DEPENDENCIES
'   M_Timer_Stop
'   M_FormBridge_GetLoadedForm
'   DP_FORM_NAME
'
' NOTES
'   Use DP_Hide when you want fast reuse
'
'   Use DP_Close when you want to unload and release the form
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm      As Object        'Loaded DatePicker form instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress hide errors
        On Error Resume Next

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Stop live-clock timer while the form is hidden
        M_Timer_Stop

'------------------------------------------------------------------------------
' HIDE FORM
'------------------------------------------------------------------------------
    'Retrieve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Hide the form when it is loaded
        If Not LoadedForm Is Nothing Then LoadedForm.Hide

'------------------------------------------------------------------------------
' RELEASE REFERENCES
'------------------------------------------------------------------------------
    'Release the local loaded-form reference
        Set LoadedForm = Nothing

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed hide error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub DP_Stop()

'
'------------------------------------------------------------------------------
'                           STOP DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Stops DatePicker session-level integrations and clears transient UI artifacts
'
' WHY THIS EXISTS
'   The DatePicker uses application-wide and workbook-level transient surfaces:
'     - Excel Application event manager
'     - right-click command-bar entries
'     - keyboard shortcut assignment
'     - modeless UserForm
'     - live-clock timer
'     - worksheet grid icon shapes
'
'   Workbook close / add-in unload must remove those artifacts explicitly so
'   nothing survives after the host project is closed
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Releases the global DatePicker manager, removes right-click menu entries,
'   removes the keyboard shortcut, closes the DatePicker form, stops timer
'   activity, and purges in-grid icon shapes from open workbooks
'
' ERROR POLICY
'   Best-effort teardown
'
'   Suppresses cleanup errors because workbook shutdown must not be interrupted
'   by missing forms, missing command bars, protected sheets, or stale shapes
'
' DEPENDENCIES
'   gDP_Manager
'   M_ContextMenu_Remove
'   M_KeyboardShortcut_Remove
'   DP_Close
'   M_Timer_Stop
'   M_GridIcon_PurgeAll
'
' NOTES
'   Releasing gDP_Manager triggers cDatePickerManager.Class_Terminate when the
'   manager exists
'
'   M_ContextMenu_Remove is called explicitly because right-click menu ownership
'   is not delegated to cDatePickerManager teardown
'
'   This routine is safe to call more than once
'
' UPDATED
'   2026-05-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress shutdown errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RELEASE MANAGER
'------------------------------------------------------------------------------
    'Release the Application event manager and trigger its teardown path
        Set gDP_Manager = Nothing

'------------------------------------------------------------------------------
' REMOVE APPLICATION-WIDE ENTRY POINTS
'------------------------------------------------------------------------------
    'Remove DatePicker right-click command-bar entries
        M_ContextMenu_Remove
    'Remove the DatePicker keyboard shortcut assignment
        M_KeyboardShortcut_Remove

'------------------------------------------------------------------------------
' CLEAR TRANSIENT UI
'------------------------------------------------------------------------------
    'Stop any active live-clock timer
        M_Timer_Stop
    'Close any loaded DatePicker form
        DP_Close
    'Purge all DatePicker grid icons from open workbooks
        M_GridIcon_PurgeAll

'------------------------------------------------------------------------------
' CLEAR CALLBACK CACHE
'------------------------------------------------------------------------------
    'Clear cached workbook-qualified callback names
        M_GetQualifiedMacroName_ClearCache

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed teardown error
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
        M_WriteBack_Apply DP_WriteAction_DatePicker, NoTableGrow

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
            'Hide the DatePicker form
                DP_Hide
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
'   hide-or-refresh lifecycle used by normal day-cell selection
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
'   When close-after-selection is enabled, the final visual lifecycle is now
'   handled by DP_Hide inside M_Picker_SelectDate so the preloaded form remains
'   available for fast reuse
'
' UPDATED
'   2026-05-17
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
        M_WriteBack_Apply DP_WriteAction_DatePicker, False

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
            'Hide the DatePicker form
                DP_Hide
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
'   Validates the requested write action
'   Validates that the prepared DatePicker write value is not the zero-date
'   Captures the current Excel event state
'   Disables events during write-back
'   Delegates target resolution and write-back to M_WriteBack_ResolveAndApplyTarget
'   Restores the previous Excel event state before exiting
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
'   gDP_WriteValue
'   M_WriteBack_ResolveAndApplyTarget
'   Application.EnableEvents
'
' NOTES
'   gDP_WriteValue is declared as Date, not Variant
'
'   Because a Date variable cannot be Null, Empty, or an Excel error value, this
'   routine treats zero-date as the uninitialized write-value sentinel
'
'   This routine intentionally does not change calculation mode
'
'   This routine intentionally does not change screen updating
'
'   Unsupported write actions are rejected explicitly instead of silently doing
'   nothing
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_WriteBack_Apply" 'Current procedure name

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
            Case DP_WriteAction_DatePicker
                'Supported DatePicker write action
            Case Else
                'Reject unsupported write actions
                    Err.Raise vbObjectError + 513, PROC_NAME, _
                        "Unsupported DatePicker write action: " & VBA.CStr(VBA.CLng(iType))
        End Select

'------------------------------------------------------------------------------
' VALIDATE PREPARED WRITE VALUE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate prepared write value"
    'Reject the zero-date sentinel because gDP_WriteValue is a Date, not a Variant
        If gDP_WriteValue = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DatePicker write value has not been initialized."
        End If

'------------------------------------------------------------------------------
' CAPTURE EXCEL EVENT STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture Excel event state"
    'Capture the current Excel events state
        PreviousEvents = Excel.Application.EnableEvents
    'Mark the Excel event state as captured
        EventsStateCaptured = True

'------------------------------------------------------------------------------
' SUPPRESS EVENTS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Suppress Excel events"
    'Disable events only when they are currently enabled
        If PreviousEvents Then
            Excel.Application.EnableEvents = False
        End If

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
        If EventsStateCaptured Then
            Excel.Application.EnableEvents = PreviousEvents
        End If
    'Capture cleanup failure
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
            Case DP_WriteAction_DatePicker
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
'   This routine centralizes the final population logic so write-back behavior
'   remains consistent across calendar-day selection, Today, Now, keyboard
'   shortcuts, context-menu actions, and public macro entry points
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
'   value, attempts a fast bulk write for multi-cell ranges, and falls back to
'   safe cell-by-cell write-back when the bulk write cannot be completed
'
'   Protected locked cells may be skipped by the fallback path when at least one
'   target cell is written successfully
'
' ERROR POLICY
'   Raises a descriptive runtime error if the target range is missing, the write
'   action is unsupported, the write value cannot be resolved, or no cell can be
'   written successfully
'
'   Bulk-write failures are not raised directly because they are expected in
'   mixed protected, validated, or partially writable ranges. The routine falls
'   back to the existing per-cell write policy
'
' DEPENDENCIES
'   M_WriteBack_GetPickedValue
'   M_WriteBack_TryBulkWriteRange
'   M_WriteBack_TryWriteCell
'   DP_MSGBOX_TITLE
'
' NOTES
'   Bulk write is used only as a fast path
'
'   The fallback preserves the existing robust behavior for protected sheets,
'   locked cells, validation failures, and partially writable selections
'
'   This routine intentionally uses Range.Value rather than Range.Value2 so VBA
'   Date and DateTime values are written through Excel's normal date handling
'
' UPDATED
'   2026-05-17
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
            Case DP_WriteAction_DatePicker
                'Resolve the current DatePicker write value
                    WriteValue = M_WriteBack_GetPickedValue
            Case Else
                'Reject unsupported write actions
                    Err.Raise vbObjectError + 515, PROC_NAME, _
                        "Unsupported DatePicker write action: " & VBA.CStr(VBA.CLng(iType))
        End Select

'------------------------------------------------------------------------------
' ATTEMPT FAST BULK WRITE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Attempt fast bulk write"
    'Use the bulk path only when it can provide a meaningful benefit
        If AttemptedCount > 1 Then
            'Exit immediately when the fast bulk write succeeds
                If M_WriteBack_TryBulkWriteRange(oRange, WriteValue) Then GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' FALL BACK TO SAFE CELL-BY-CELL WRITE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Populate target cells through safe fallback"
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

Private Function M_WriteBack_TryBulkWriteRange( _
    ByVal TargetRange As Range, _
    ByVal WriteValue As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY BULK WRITE RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to write one DatePicker value to an entire Excel range in one
'   operation
'
' WHY THIS EXISTS
'   Cell-by-cell write-back is robust but slow for large ranges, table columns,
'   and multi-cell selections
'
'   A bulk write is much faster when the entire target is writable. When it is
'   not writable, the caller can fall back to the safer per-cell path
'
' INPUTS
'   TargetRange
'     Excel range to populate
'
'   WriteValue
'     DatePicker value to write
'
' RETURNS
'   True when the bulk write succeeds
'
'   False when the range is missing, empty, protected, validated, partially
'   unwritable, or otherwise cannot be written in one operation
'
' BEHAVIOR
'   Validates the range reference, attempts one Range.Value assignment, and
'   returns a Boolean success flag
'
' ERROR POLICY
'   Safe-default helper
'
'   Does not raise outward. Any failure returns False so the caller can use the
'   safe fallback path
'
' DEPENDENCIES
'   Excel.Range
'
' NOTES
'   This helper intentionally uses Range.Value rather than Range.Value2 so VBA
'   Date and DateTime values are written through Excel's normal date handling
'
'   This helper does not suppress Application events. Event suppression is owned
'   by M_WriteBack_Apply
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_WriteBack_TryBulkWriteRange = False
    'Enable safe-default error handling
        On Error GoTo BulkWriteFail

'------------------------------------------------------------------------------
' VALIDATE TARGET RANGE
'------------------------------------------------------------------------------
    'Exit when no target range is supplied
        If TargetRange Is Nothing Then Exit Function
    'Exit when the target range has no cells
        If TargetRange.Cells.CountLarge <= 0 Then Exit Function

'------------------------------------------------------------------------------
' APPLY BULK WRITE
'------------------------------------------------------------------------------
    'Write the DatePicker value to the whole target range in one operation
        TargetRange.Value = WriteValue

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return success after the bulk write has completed
        M_WriteBack_TryBulkWriteRange = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful bulk write
        Exit Function

'------------------------------------------------------------------------------
' BULK WRITE FAIL
'------------------------------------------------------------------------------
BulkWriteFail:
    'Suppress the bulk failure and let the caller fall back to cell-by-cell write
        On Error Resume Next
    'Return safe default
        M_WriteBack_TryBulkWriteRange = False
    'Clear the suppressed bulk-write error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function
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
'   Validates that gDP_WriteValue has been initialized
'   Preserves any time component already stored in gDP_WriteValue
'   Validates the supported DatePicker year range
'   Returns the validated DatePicker write value
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker write value is the
'   zero-date sentinel or outside the supported DatePicker year range
'
' DEPENDENCIES
'   gDP_WriteValue
'   DP_MIN_YEAR
'   DP_MAX_YEAR
'
' NOTES
'   gDP_WriteValue is declared as Date, not Variant
'
'   Therefore this routine does not test for Null, Empty, Excel error values, or
'   non-date values because those states cannot exist in a Date variable
'
'   The zero-date value is treated as the uninitialized write-value sentinel
'
'   Calendar day selection and Today prepare date-only values upstream
'
'   Now prepares a date-time value upstream
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_WriteBack_GetPickedValue" 'Current procedure name

    Dim PickedValue            As Date         'Resolved DatePicker write value
    Dim PickedYear             As Long         'Resolved DatePicker write year
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE WRITE VALUE
'------------------------------------------------------------------------------
    'Read the prepared DatePicker write value
        PickedValue = gDP_WriteValue

'------------------------------------------------------------------------------
' VALIDATE INITIALIZATION STATE
'------------------------------------------------------------------------------
    'Reject the zero-date sentinel
        If PickedValue = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, _
                "DatePicker write value has not been initialized."
        End If

'------------------------------------------------------------------------------
' VALIDATE SUPPORTED YEAR RANGE
'------------------------------------------------------------------------------
    'Resolve the write-value year
        PickedYear = VBA.Year(PickedValue)
    'Reject dates before the supported DatePicker year range
        If PickedYear < DP_MIN_YEAR Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "DatePicker write value year is before the supported minimum year: " & _
                VBA.CStr(DP_MIN_YEAR)
        End If
    'Reject dates after the supported DatePicker year range
        If PickedYear > DP_MAX_YEAR Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DatePicker write value year is after the supported maximum year: " & _
                VBA.CStr(DP_MAX_YEAR)
        End If

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
    'Capture the error number
        ErrorNumber = Err.Number

    'Capture the error description
        ErrorDescription = Err.Description

    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, _
            "DatePicker write value resolution failed: " & ErrorDescription

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
'   gDP_AllowOutsideMonthSelection
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
        If DisplayYear < DP_MIN_YEAR Or DisplayYear > DP_MAX_YEAR Then
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
        If CandidateYear < DP_MIN_YEAR Or CandidateYear > DP_MAX_YEAR Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY OUTSIDE-MONTH POLICY
'------------------------------------------------------------------------------
    'Allow valid candidate dates when outside-month selection is enabled
        If gDP_AllowOutsideMonthSelection Then
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
'   Fail-open from a date-selection perspective
'   Callback failures are interpreted as not holiday
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
        If CandidateYear < DP_MIN_YEAR Then Exit Function
    'Reject dates after the supported DatePicker year range
        If CandidateYear > DP_MAX_YEAR Then Exit Function

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
'

Public Function M_Caption_GetMonth( _
    ByVal MonthNumber As Long, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET MONTH CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   logic keeps local-name mode and fixed-English mode consistent across the
'   form, picker panels, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
'   UseLocalNames
'     True to use the local VBA month name
'     False to use the fixed-English DatePicker caption
'
' RETURNS
'   Uppercase trimmed month caption
'
' BEHAVIOR
'   Validates the month number, resolves the caption using either local VBA
'   month names or the fixed-English helper, then returns a normalized uppercase
'   caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
'   Other unexpected failures are allowed to propagate naturally so the caller
'   receives the original VBA runtime error
'
' DEPENDENCIES
'   VBA.MonthName
'   M_Caption_GetEnglishMonthFull
'
' NOTES
'   Fixed-English mode is delegated to M_Caption_GetEnglishMonthFull so English
'   month naming remains controlled by the DatePicker companion-module API
'
'   Local-name mode uses VBA.MonthName and is therefore driven by the host
'   Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Caption_GetMonth"

    Dim MonthCaption        As String       'Resolved month caption

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RESOLVE MONTH CAPTION
'------------------------------------------------------------------------------
    'Resolve the month caption using local host names when requested
        If UseLocalNames Then
            MonthCaption = VBA.MonthName(MonthNumber, False)
    'Otherwise resolve the month caption using the fixed-English helper
        Else
            MonthCaption = M_Caption_GetEnglishMonthFull(MonthNumber)
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized display caption
        M_Caption_GetMonth = VBA.UCase$(VBA.Trim$(MonthCaption))

End Function


Public Function M_Caption_GetDate( _
    ByVal DateValue As Date, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET DATE CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display caption for a DatePicker date value
'
' WHY THIS EXISTS
'   DatePicker date captions are used in several UI locations. Centralizing the
'   logic keeps local-name mode and fixed-English mode consistent across the
'   form, picker panels, and regression tests
'
' INPUTS
'   DateValue
'     Date value to convert into display text
'
'   UseLocalNames
'     True to use the local abbreviated month name
'     False to use the fixed-English DatePicker abbreviated month caption
'
' RETURNS
'   Date caption in dd-MMM-yyyy style
'
' BEHAVIOR
'   Builds the day, month, and year components from DateValue, resolves the month
'   component using either local VBA formatting or the fixed-English helper, and
'   returns the final normalized caption
'
' ERROR POLICY
'   Raises unexpected failures back to the caller with this procedure name as
'   the error source
'
' DEPENDENCIES
'   VBA.Format$
'   VBA.Month
'   M_Caption_GetEnglishMonthShort
'
' NOTES
'   Local-name mode uses VBA.Format$(DateValue, "mmm") and is therefore driven
'   by the host Office / Windows locale
'
'   Fixed-English mode is delegated to M_Caption_GetEnglishMonthShort so English
'   month naming remains controlled by the DatePicker companion-module API
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetDate"

    Dim DayText                 As String        'Day component
    Dim MonthText               As String        'Month component
    Dim YearText                As String        'Year component
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' BUILD DAY COMPONENT
'------------------------------------------------------------------------------
    'Build the two-digit day component
        DayText = VBA.Format$(DateValue, "dd")

'------------------------------------------------------------------------------
' BUILD MONTH COMPONENT
'------------------------------------------------------------------------------
    'Build the local abbreviated month component when requested
        If UseLocalNames Then
            MonthText = VBA.UCase$(VBA.Format$(DateValue, "mmm"))
    'Otherwise build the fixed-English abbreviated month component
        Else
            MonthText = M_Caption_GetEnglishMonthShort(VBA.Month(DateValue))
        End If

'------------------------------------------------------------------------------
' BUILD YEAR COMPONENT
'------------------------------------------------------------------------------
    'Build the four-digit year component
        YearText = VBA.Format$(DateValue, "yyyy")

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the final date caption
        M_Caption_GetDate = DayText & "-" & MonthText & "-" & YearText

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Function

Public Function M_Caption_GetEnglishMonthShort(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISH MONTH SHORT CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed-English abbreviated month caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   fixed-English caption map keeps display behavior consistent across the form,
'   picker panels, caption helpers, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
' RETURNS
'   Fixed-English uppercase three-letter month caption
'
' BEHAVIOR
'   Validates the month number and returns the corresponding fixed-English short
'   month caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper deliberately does not use VBA.Format$ or VBA.MonthName because
'   fixed-English mode must be independent from the host Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetEnglishMonthShort"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RETURN FIXED-ENGLISH SHORT MONTH CAPTION
'------------------------------------------------------------------------------
    'Return the fixed-English short caption for the requested month
        Select Case MonthNumber
            Case 1
                M_Caption_GetEnglishMonthShort = "JAN"
            Case 2
                M_Caption_GetEnglishMonthShort = "FEB"
            Case 3
                M_Caption_GetEnglishMonthShort = "MAR"
            Case 4
                M_Caption_GetEnglishMonthShort = "APR"
            Case 5
                M_Caption_GetEnglishMonthShort = "MAY"
            Case 6
                M_Caption_GetEnglishMonthShort = "JUN"
            Case 7
                M_Caption_GetEnglishMonthShort = "JUL"
            Case 8
                M_Caption_GetEnglishMonthShort = "AUG"
            Case 9
                M_Caption_GetEnglishMonthShort = "SEP"
            Case 10
                M_Caption_GetEnglishMonthShort = "OCT"
            Case 11
                M_Caption_GetEnglishMonthShort = "NOV"
            Case 12
                M_Caption_GetEnglishMonthShort = "DEC"
        End Select

End Function

Public Function M_Caption_GetEnglishMonthFull(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISH MONTH FULL CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed-English full month caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   fixed-English caption map keeps display behavior consistent across the form,
'   picker panels, caption helpers, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
' RETURNS
'   Fixed-English uppercase full month caption
'
' BEHAVIOR
'   Validates the month number and returns the corresponding fixed-English full
'   month caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper deliberately does not use VBA.Format$ or VBA.MonthName because
'   fixed-English mode must be independent from the host Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetEnglishMonthFull"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RETURN FIXED-ENGLISH FULL MONTH CAPTION
'------------------------------------------------------------------------------
    'Return the fixed-English full caption for the requested month
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
'

Private Function M_Timer_GetProcedureName() As String

'
'------------------------------------------------------------------------------
'                         TIMER GET PROCEDURE NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the cached workbook-qualified DatePicker timer procedure name
'
' WHY THIS EXISTS
'   Application.OnTime needs a workbook-qualified callback name. Rebuilding the
'   same qualified macro name on every timer tick is unnecessary.
'
' INPUTS
'   None
'
' RETURNS
'   Workbook-qualified M_Timer_Tick procedure name
'
' BEHAVIOR
'   Reuses the cached procedure name when it was already built for the current
'   ThisWorkbook.Name. Rebuilds it when the cache is empty or when the workbook
'   name changed during the session.
'
' ERROR POLICY
'   Raises a descriptive runtime error if the qualified procedure name cannot be
'   resolved
'
' DEPENDENCIES
'   M_GetQualifiedMacroName
'   ThisWorkbook.Name
'
' NOTES
'   The workbook-name check protects Save As / rename scenarios during a live
'   Excel session
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "M_Timer_GetProcedureName"
    Const TIMER_TICK_PROCEDURE      As String = "M_Timer_Tick"

    Dim CurrentWorkbookName         As String       'Current host workbook name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE CURRENT WORKBOOK NAME
'------------------------------------------------------------------------------
    'Read the current host workbook name
        CurrentWorkbookName = VBA.CStr(ThisWorkbook.Name)

'------------------------------------------------------------------------------
' REBUILD CACHE WHEN NEEDED
'------------------------------------------------------------------------------
    'Rebuild the cached procedure name when empty or stale
        If VBA.LenB(mDP_TimerProcedureName) = 0 _
        Or VBA.StrComp(mDP_TimerProcedureWorkbookName, CurrentWorkbookName, vbBinaryCompare) <> 0 Then
            'Build the workbook-qualified timer callback
                mDP_TimerProcedureName = M_GetQualifiedMacroName(TIMER_TICK_PROCEDURE)
            'Store the workbook name used for this cache entry
                mDP_TimerProcedureWorkbookName = CurrentWorkbookName
        End If

'------------------------------------------------------------------------------
' RETURN PROCEDURE NAME
'------------------------------------------------------------------------------
    'Return the cached timer procedure name
        M_Timer_GetProcedureName = mDP_TimerProcedureName

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear stale timer procedure cache
        mDP_TimerProcedureName = vbNullString
    'Clear stale workbook-name cache
        mDP_TimerProcedureWorkbookName = vbNullString
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, PROC_NAME, _
            "Timer procedure-name resolution failed: " & Err.Description

End Function


Public Sub M_Timer_ApplyClockMode()

'
'==============================================================================
'                           TIMER APPLY CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the current DatePicker clock mode to the already-loaded picker form
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot
'
'   The DatePicker therefore needs one controlled routine that stops any existing
'   clock timer, refreshes the loaded form clock once, and restarts the timer
'   only when live-clock mode is enabled and a DatePicker form is actually loaded
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Stops any existing DatePicker timer
'   Resolves the already-loaded DatePicker form through the shared form bridge
'   Exits without starting a new timer when no DatePicker form is loaded
'   Refreshes the loaded DatePicker form clock when a form is available
'   Starts the next timer tick only when gDP_ClockMode is DP_ClockMode_Live
'
' ERROR POLICY
'   Stops the DatePicker timer on unexpected failure and re-raises the original
'   error with this procedure name as the source
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_FormBridge_GetLoadedForm
'   M_Timer_Stop
'   M_Timer_Start
'   DP_FORM_NAME
'   gDP_ClockMode
'   DP_ClockMode_Live
'   UF_DatePicker.UF_DP_UpdateLiveClock
'
' NOTES
'   This routine deliberately avoids direct UF_DatePicker default-instance
'   references
'
'   Loaded-form lookup is delegated to M_FormBridge_GetLoadedForm so timer,
'   refresh, and unload routines use the same loaded-form matching policy
'
'   This routine intentionally does not start Application.OnTime when the
'   DatePicker form is not loaded
'
'   If live-clock mode is enabled while the form is closed, the timer should be
'   started later by the form open / initialize lifecycle when the picker becomes
'   loaded
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Timer_ApplyClockMode"
    
    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description
    Dim StopErrNumber          As Long         'Captured timer-stop error number
    Dim StopErrDescription     As String       'Captured timer-stop error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the clock-mode flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' STOP EXISTING TIMER
'------------------------------------------------------------------------------
    'Stop any existing DatePicker timer before applying the current mode
        M_Timer_Stop

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Resolve the already-loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

'------------------------------------------------------------------------------
' EXIT WHEN NO FORM IS LOADED
'------------------------------------------------------------------------------
    'Exit without starting a timer when the DatePicker form is not loaded
        If LoadedForm Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' REFRESH LOADED FORM CLOCK
'------------------------------------------------------------------------------
    'Update the loaded DatePicker clock once
        LoadedForm.UF_DP_UpdateLiveClock

'------------------------------------------------------------------------------
' START LIVE TIMER WHEN ENABLED
'------------------------------------------------------------------------------
    'Start the live timer only when live clock mode is enabled
        If gDP_ClockMode = DP_ClockMode_Live Then M_Timer_Start

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release the loaded form reference
        Set LoadedForm = Nothing
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
    'Suppress timer-stop errors while preserving the original failure
        On Error Resume Next
    'Stop the timer after an unexpected clock-mode error
        M_Timer_Stop
    'Capture timer-stop error number
        StopErrNumber = Err.Number
    'Capture timer-stop error description
        StopErrDescription = Err.Description
    'Clear any suppressed timer-stop error
        Err.Clear
    'Release the loaded form reference
        Set LoadedForm = Nothing
    'Write timer-stop diagnostics only when cleanup failed
        If StopErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=M_Timer_Stop" & _
                " | Error=" & VBA.CStr(StopErrNumber) & _
                " | " & StopErrDescription
        End If
    'Restore normal error handling
        On Error GoTo 0
    'Re-raise the original error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub
Public Sub M_Timer_Start()

'
'------------------------------------------------------------------------------
'                           START TIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot. The DatePicker therefore needs a controlled
'   start routine that builds the workbook-qualified tick procedure name,
'   records the next scheduled tick, and schedules the next timer callback
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when the timer is already running, resolves the workbook-qualified
'   timer tick procedure name, marks the timer as running, calculates the next
'   tick time, and schedules the next Application.OnTime callback
'
' ERROR POLICY
'   Clears timer state after a scheduling failure and re-raises the original
'   error with this procedure name as the source
'
' DEPENDENCIES
'   Application.OnTime
'   M_GetQualifiedMacroName
'   M_Timer_Tick
'   mDP_TimerIsRunning
'   mDP_TimerProcedureName
'   mDP_NextTickTime
'   DP_TIMER_SECONDS
'
' NOTES
'   This routine does not update the visible clock directly
'
'   Clock refresh is handled by M_Timer_Tick and by M_Timer_ApplyClockMode
'
'   The procedure name is workbook-qualified so OnTime can find the callback
'   reliably when multiple workbooks or add-ins are open
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Timer_Start"

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' EXIT IF ALREADY RUNNING
'------------------------------------------------------------------------------
    'Exit if the timer is already running
        If mDP_TimerIsRunning Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE TIMER PROCEDURE
'------------------------------------------------------------------------------
    'Build the workbook-qualified timer procedure name
        mDP_TimerProcedureName = M_Timer_GetProcedureName
        
'------------------------------------------------------------------------------
' MARK TIMER RUNNING
'------------------------------------------------------------------------------
    'Mark the timer as running before scheduling the OnTime callback
        mDP_TimerIsRunning = True

'------------------------------------------------------------------------------
' CALCULATE NEXT TICK
'------------------------------------------------------------------------------
    'Calculate the next timer tick
        mDP_NextTickTime = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)

'------------------------------------------------------------------------------
' SCHEDULE NEXT TICK
'------------------------------------------------------------------------------
    'Schedule the next timer tick
        Application.OnTime _
            EarliestTime:=mDP_NextTickTime, _
            Procedure:=mDP_TimerProcedureName, _
            Schedule:=True

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
    'Clear timer running state after scheduling failure
        mDP_TimerIsRunning = False
    'Clear next tick time after scheduling failure
        mDP_NextTickTime = 0
    'Re-raise the original error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub
Public Sub M_Timer_Stop()

'
'------------------------------------------------------------------------------
'                           STOP TIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Stops the DatePicker live footer clock timer
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot and cancellation requires the same scheduled
'   time and procedure name used when the callback was registered
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Rebuilds the qualified timer procedure name when needed, attempts to cancel
'   the next scheduled timer tick when a timer is active, and always clears the
'   internal timer state
'
' ERROR POLICY
'   Best-effort cleanup. Cancellation errors are suppressed because OnTime
'   cancellation can fail when the callback already fired, Excel is shutting
'   down, the project was reset, or the procedure name cannot be resolved
'
' DEPENDENCIES
'   Application.OnTime
'   M_GetQualifiedMacroName
'   M_Timer_Tick
'   mDP_TimerIsRunning
'   mDP_TimerProcedureName
'   mDP_NextTickTime
'
' NOTES
'   This routine intentionally does not raise outward
'
'   Timer state is cleared even when the OnTime cancellation attempt fails
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Timer_Stop"

    Dim CancelErrNumber         As Long          'Captured cancellation error number
    Dim CancelErrDescription    As String        'Captured cancellation error description
    Dim NameErrNumber           As Long          'Captured procedure-name error number
    Dim NameErrDescription      As String        'Captured procedure-name error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress timer-stop errors because this is a best-effort cleanup routine
        On Error Resume Next

'------------------------------------------------------------------------------
' RESOLVE TIMER PROCEDURE NAME
'------------------------------------------------------------------------------
    'Build the timer procedure name only when cancellation may need it
        If mDP_TimerIsRunning Then
            If mDP_NextTickTime <> 0 Then
                If VBA.LenB(mDP_TimerProcedureName) = 0 Then
                    mDP_TimerProcedureName = M_Timer_GetProcedureName
                    NameErrNumber = Err.Number
                    NameErrDescription = Err.Description
                    Err.Clear
                End If
            End If
        End If

'------------------------------------------------------------------------------
' CANCEL SCHEDULED TICK
'------------------------------------------------------------------------------
    'Cancel the next scheduled timer tick when active and cancellable
        If mDP_TimerIsRunning Then
            If mDP_NextTickTime <> 0 Then
                If VBA.Len(mDP_TimerProcedureName) > 0 Then
                    Application.OnTime _
                        EarliestTime:=mDP_NextTickTime, _
                        Procedure:=mDP_TimerProcedureName, _
                        Schedule:=False
                    CancelErrNumber = Err.Number
                    CancelErrDescription = Err.Description
                    Err.Clear
                End If
            End If
        End If

'------------------------------------------------------------------------------
' CLEAR TIMER STATE
'------------------------------------------------------------------------------
    'Clear timer running state
        mDP_TimerIsRunning = False
    'Clear next tick time
        mDP_NextTickTime = 0
    'Keep the cached timer procedure name for the next timer start

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write procedure-name diagnostics only when resolution failed
        If NameErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=ResolveTimerProcedure" & _
                " | Error=" & VBA.CStr(NameErrNumber) & _
                " | " & NameErrDescription
        End If
    'Write cancellation diagnostics only when cancellation failed
        If CancelErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=CancelOnTime" & _
                " | Error=" & VBA.CStr(CancelErrNumber) & _
                " | " & CancelErrDescription
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed timer-stop error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
Public Sub M_Timer_Tick()

'
'------------------------------------------------------------------------------
'                           TIMER TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Handles one DatePicker live footer clock timer callback
'
' WHY THIS EXISTS
'   Application.OnTime is one-shot. Each timer callback must update the loaded
'   DatePicker form when available and explicitly schedule the next callback
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when the timer is no longer marked as running
'   Resolves the already-loaded DatePicker form through the shared form bridge
'   Updates the loaded DatePicker form clock when a form is available
'   Stops the timer when no DatePicker form is loaded
'   Rebuilds the workbook-qualified timer procedure name
'   Records the next scheduled tick time
'   Schedules the next Application.OnTime callback
'
' ERROR POLICY
'   Fail-safe
'   Stops the timer after any unexpected callback failure
'   Writes diagnostics to the Immediate Window
'
' DEPENDENCIES
'   Application.OnTime
'   M_FormBridge_GetLoadedForm
'   M_GetQualifiedMacroName
'   M_Timer_Stop
'   UF_DatePicker.UF_DP_UpdateLiveClock
'   DP_FORM_NAME
'   DP_TIMER_SECONDS
'   mDP_TimerIsRunning
'   mDP_TimerProcedureName
'   mDP_NextTickTime
'
' NOTES
'   This routine must remain Public because Application.OnTime calls it by name
'
'   This routine deliberately avoids direct UF_DatePicker default-instance
'   references
'
'   Loaded-form lookup is delegated to M_FormBridge_GetLoadedForm so timer,
'   refresh, and unload routines use the same loaded-form matching policy
'
'   Application.OnTime is one-shot, so successful ticks must schedule the next
'   tick explicitly
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Timer_Tick" 'Current procedure name

    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description
    Dim StopErrNumber          As Long         'Captured timer-stop error number
    Dim StopErrDescription     As String       'Captured timer-stop error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EXIT IF TIMER IS NOT RUNNING
'------------------------------------------------------------------------------
    'Exit if the timer is no longer running
        If Not mDP_TimerIsRunning Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Resolve the already-loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

'------------------------------------------------------------------------------
' STOP TIMER WHEN FORM IS NOT LOADED
'------------------------------------------------------------------------------
    'Stop the timer if the DatePicker form is no longer loaded
        If LoadedForm Is Nothing Then
            M_Timer_Stop
            GoTo ExitProcedure
        End If

'------------------------------------------------------------------------------
' REFRESH LOADED FORM CLOCK
'------------------------------------------------------------------------------
    'Update the loaded DatePicker form clock
        LoadedForm.UF_DP_UpdateLiveClock

'------------------------------------------------------------------------------
' RESOLVE TIMER PROCEDURE
'------------------------------------------------------------------------------
    'Build the workbook-qualified timer procedure name
        mDP_TimerProcedureName = M_Timer_GetProcedureName
        
'------------------------------------------------------------------------------
' CALCULATE NEXT TICK
'------------------------------------------------------------------------------
    'Calculate the next timer tick
        mDP_NextTickTime = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)

'------------------------------------------------------------------------------
' SCHEDULE NEXT TICK
'------------------------------------------------------------------------------
    'Schedule the next timer tick
        Excel.Application.OnTime _
            EarliestTime:=mDP_NextTickTime, _
            Procedure:=mDP_TimerProcedureName, _
            Schedule:=True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release the loaded form reference
        Set LoadedForm = Nothing
    'Exit after successful handling
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the original error number before cleanup can alter Err
        ErrorNumber = Err.Number
    'Capture the original error description before cleanup can alter Err
        ErrorDescription = Err.Description
    'Suppress timer-stop errors while preserving the original callback failure
        On Error Resume Next
    'Stop the timer after an unexpected callback error
        M_Timer_Stop
    'Capture timer-stop error number
        StopErrNumber = Err.Number
    'Capture timer-stop error description
        StopErrDescription = Err.Description
    'Clear any suppressed timer-stop error
        Err.Clear
    'Release the loaded form reference
        Set LoadedForm = Nothing
    'Write timer-stop diagnostics only when cleanup failed
        If StopErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=M_Timer_Stop" & _
                " | Error=" & VBA.CStr(StopErrNumber) & _
                " | " & StopErrDescription
        End If
    'Write callback diagnostics to the Immediate Window
        Debug.Print PROC_NAME & _
            " | Step=TimerCallback" & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Restore normal error handling
        On Error GoTo 0

End Sub
'
'------------------------------------------------------------------------------
'
'                                WINAPI HELPERS
'
'------------------------------------------------------------------------------
'


Public Function M_Platform_CanUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM CAN USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the current platform can use Windows API calls
'
' WHY THIS EXISTS
'   Settings normalization and platform-safe startup logic need to know whether
'   WinAPI calls are technically available before applying user configuration
'
' INPUTS
'   None
'
' RETURNS
'   True when the project is running on Windows
'   False when the project is running on Mac
'
' BEHAVIOR
'   Uses conditional compilation to return a platform-safe Boolean without
'   calling any native API
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   VBA conditional compilation constant Mac
'
' NOTES
'   This routine answers platform capability only
'
'   It deliberately does not inspect gDP_UseWinAPI
'
'   Runtime policy belongs in M_Platform_ShouldUseWinAPI
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

#If Mac Then
    'Return False when running on Mac
        M_Platform_CanUseWinAPI = False
#Else
    'Return True when running on Windows
        M_Platform_CanUseWinAPI = True
#End If

End Function
Public Function M_Platform_ShouldUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM SHOULD USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether optional DatePicker WinAPI styling should be enabled
'
' WHY THIS EXISTS
'   Some DatePicker visual behavior, especially native title-bar removal, should
'   be controlled by both:
'     - platform capability
'     - user / project configuration
'
'   This helper provides the effective runtime policy for optional WinAPI styling
'   behavior. It is not the general WinAPI capability check.
'
' INPUTS
'   None
'
' RETURNS
'   True when WinAPI is available on the current platform and styling is enabled
'   False otherwise
'
' BEHAVIOR
'   Returns True only when both conditions are satisfied:
'     - the current platform can use WinAPI
'     - gDP_UseWinAPI is True
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   gDP_UseWinAPI
'
' NOTES
'   Use this helper for optional styling features such as title-bar removal.
'
'   Use M_Platform_CanUseWinAPI for capability-only operations that should remain
'   available on Windows even when optional WinAPI styling is disabled.
'
'   Do not call M_Settings_EnsureLoaded from this helper.
'
'   Settings load / save routines already use M_Platform_CanUseWinAPI to avoid
'   circular policy dependencies while normalizing gDP_UseWinAPI.
'
' UPDATED
'   2026-05-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN POLICY DECISION
'------------------------------------------------------------------------------
    'Return whether optional WinAPI-dependent behavior should be used
        M_Platform_ShouldUseWinAPI = (M_Platform_CanUseWinAPI And gDP_UseWinAPI)

End Function

Public Sub M_Window_RemoveTitleBar(ByVal Frm As Object)

'
'==============================================================================
'                           WINDOW REMOVE TITLE BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the native title bar from a DatePicker UserForm on Windows
'
' WHY THIS EXISTS
'   The DatePicker can use a borderless visual style
'
'   Removing the native UserForm title bar requires Windows API calls and must
'   therefore degrade safely on Mac or when WinAPI behavior is disabled by
'   settings
'
' INPUTS
'   Frm
'     UserForm instance whose native title bar should be removed
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits safely on Mac
'   Exits safely when no form is supplied
'   Exits safely when WinAPI usage is disabled
'   Exits safely when the form window handle cannot be resolved
'   Reads the current native window style
'   Removes the WS_CAPTION style bit
'   Writes the updated native window style
'   Refreshes the non-client frame
'   Redraws the menu bar / frame area
'   Logs WinAPI return-code failures to the Immediate Window
'
' ERROR POLICY
'   Best-effort UI styling
'   Suppresses unexpected WinAPI or form-handle errors
'   Diagnoses both VBA runtime errors and WinAPI return-code failures
'   Does not raise outward
'
' DEPENDENCIES
'   M_Platform_ShouldUseWinAPI
'   M_Window_GetUserFormHwnd
'   GetWindowLongPtr / GetWindowLong
'   SetWindowLongPtr / SetWindowLong
'   SetWindowPos
'   DrawMenuBar
'   Err.LastDllError
'   SetLastError
'   GWL_STYLE
'   WS_CAPTION
'   SWP_NOMOVE
'   SWP_NOSIZE
'   SWP_NOZORDER
'   SWP_NOACTIVATE
'   SWP_FRAMECHANGED
'
' NOTES
'   This routine intentionally does nothing on Mac
'
'   SetWindowLongPtr / SetWindowLong return the previous value, not a Boolean
'   success flag
'
'   Because zero can theoretically be either a previous value or a failure, the
'   routine clears the WinAPI last-error state before the call and then inspects
'   Err.LastDllError when the return value is zero
'
'   SetWindowPos and DrawMenuBar return zero on failure
'
'   The routine does not raise outward because title-bar removal is visual polish,
'   not a functional requirement for date selection
'
' UPDATED
'   2026-08-21
'==============================================================================

#If Mac Then

'------------------------------------------------------------------------------
' MAC SAFE EXIT
'------------------------------------------------------------------------------
    'Do nothing on Mac
        Exit Sub

#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_RemoveTitleBar" 'Current procedure name

    #If VBA7 Then
        Dim hWndForm            As LongPtr       'UserForm window handle
        Dim WindowStyle         As LongPtr       'Window style bits
        Dim SetStyleResult      As LongPtr       'Previous window style returned by SetWindowLongPtr
    #Else
        Dim hWndForm            As Long          'UserForm window handle
        Dim WindowStyle         As Long          'Window style bits
        Dim SetStyleResult      As Long          'Previous window style returned by SetWindowLong
    #End If

    Dim WindowFlags             As Long          'SetWindowPos flags
    Dim ApiResult               As Long          'Generic WinAPI Boolean-style result
    Dim LastApiError            As Long          'WinAPI last-error code
    Dim HandlerStep             As String        'Current handler step for diagnostics
    Dim ErrorNumber             As Long          'Captured VBA error number
    Dim ErrorDescription        As String        'Captured VBA error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress borderless styling errors through the local cleanup path
        On Error GoTo CleanExit
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate input"
    'Exit if no form was supplied
        If Frm Is Nothing Then GoTo CleanExit

'------------------------------------------------------------------------------
' CHECK WINAPI POLICY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check WinAPI policy"
    'Exit if WinAPI features should not be used
        If Not M_Platform_ShouldUseWinAPI Then GoTo CleanExit


'------------------------------------------------------------------------------
' RESOLVE FORM HANDLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve form handle"
    'Resolve the form window handle
        hWndForm = M_Window_GetUserFormHwnd(Frm)
    'Exit if the form window handle is unavailable
        If hWndForm = 0 Then GoTo CleanExit


'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Read window style"

    #If VBA7 Then
        'Read current window style
            WindowStyle = GetWindowLongPtr(hWndForm, GWL_STYLE)
    #Else
        'Read current window style
            WindowStyle = GetWindowLong(hWndForm, GWL_STYLE)
    #End If

    'Exit if the style cannot be read
        If WindowStyle = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=GetWindowLong" & _
                " | LastError=" & VBA.CStr(LastApiError)
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' REMOVE TITLE BAR STYLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Remove title bar style"
    'Remove the caption style bit
        WindowStyle = (WindowStyle And Not WS_CAPTION)

'------------------------------------------------------------------------------
' WRITE WINDOW STYLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Write window style"
    'Clear the WinAPI last-error state before SetWindowLong
        SetLastError 0

    #If VBA7 Then
        'Write the updated window style
            SetStyleResult = SetWindowLongPtr(hWndForm, GWL_STYLE, WindowStyle)
    #Else
        'Write the updated window style
            SetStyleResult = SetWindowLong(hWndForm, GWL_STYLE, WindowStyle)
    #End If

    'Diagnose SetWindowLong failure when return is zero and LastError is non-zero
        If SetStyleResult = 0 Then
            LastApiError = Err.LastDllError
            If LastApiError <> 0 Then
                Debug.Print PROC_NAME & _
                    " | Step=" & HandlerStep & _
                    " | Api=SetWindowLong" & _
                    " | LastError=" & VBA.CStr(LastApiError)
                GoTo CleanExit
            End If
        End If

'------------------------------------------------------------------------------
' REFRESH NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh non-client frame"
    'Build non-client refresh flags
        WindowFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
            SWP_NOACTIVATE Or SWP_FRAMECHANGED
    'Clear the WinAPI last-error state before SetWindowPos
        SetLastError 0
    'Force Windows to recalculate the frame
        ApiResult = SetWindowPos(hWndForm, 0, 0, 0, 0, 0, WindowFlags)
    'Diagnose SetWindowPos return-code failure
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=SetWindowPos" & _
                " | LastError=" & VBA.CStr(LastApiError)
        End If

'------------------------------------------------------------------------------
' REDRAW FRAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Redraw frame"
    'Clear the WinAPI last-error state before DrawMenuBar
        SetLastError 0
    'Redraw menu bar and non-client elements
        ApiResult = DrawMenuBar(hWndForm)
    'Diagnose DrawMenuBar return-code failure
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=DrawMenuBar" & _
                " | LastError=" & VBA.CStr(LastApiError)
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Capture any suppressed VBA error number
        ErrorNumber = Err.Number
    'Capture any suppressed VBA error description
        ErrorDescription = Err.Description
    'Write diagnostics only when VBA raised during borderless styling
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear any suppressed styling error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

#End If

End Sub

Public Sub M_Window_BeginUserFormDrag(ByVal TargetForm As Object)

'
'------------------------------------------------------------------------------
'                       BEGIN USERFORM DRAG
'------------------------------------------------------------------------------
' PURPOSE
'   Starts native Windows drag movement for a borderless UserForm
'
' WHY THIS EXISTS
'   Removing the native title bar removes the normal Windows drag surface. This
'   helper lets a runtime label, header banner, or other custom surface behave
'   like the missing title bar
'
' INPUTS
'   TargetForm
'     UserForm instance to move
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the UserForm window handle from its Caption, releases current mouse
'   capture, and sends a title-bar drag message to the UserForm window
'
' ERROR POLICY
'   Best-effort helper
'
'   Silently exits when WinAPI capability is unavailable, when TargetForm is
'   missing, or when the UserForm window handle cannot be resolved
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   M_Window_GetUserFormHwnd
'   ReleaseCapture
'   SendMessage
'   DP_WM_NCLBUTTONDOWN
'   DP_HTCAPTION
'
' NOTES
'   The UserForm Caption should remain stable and sufficiently unique while the
'   form is open. The title bar may be hidden, but the Caption property can still
'   be used internally to resolve the window handle
'
'   This routine uses capability only, not the optional styling setting. Mouse
'   movement is safe on Windows even when optional WinAPI styling is disabled
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then
    'Exit because this helper is Windows-only
        Exit Sub
#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    #If VBA7 Then
        Dim FormHandle As LongPtr       'Resolved UserForm window handle
    #Else
        Dim FormHandle As Long          'Resolved UserForm window handle
    #End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress drag helper errors
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit when WinAPI calls are not available
        If Not M_Platform_CanUseWinAPI Then Exit Sub
    'Exit when the target form reference is missing
        If TargetForm Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE USERFORM WINDOW
'------------------------------------------------------------------------------
    'Resolve the handle through the single module resolver
        FormHandle = M_Window_GetUserFormHwnd(TargetForm)
    'Exit when the UserForm window cannot be resolved
        If FormHandle = 0 Then GoTo CleanExit

'------------------------------------------------------------------------------
' START NATIVE DRAG
'------------------------------------------------------------------------------
    'Release current mouse capture
        ReleaseCapture
    'Ask Windows to move the form as if the title bar were being dragged
        SendMessage FormHandle, DP_WM_NCLBUTTONDOWN, DP_HTCAPTION, 0

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Clear any suppressed drag error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

#End If

End Sub


#If VBA7 Then
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As LongPtr
#Else
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As Long
#End If

'
'------------------------------------------------------------------------------
'                           WINDOW GET USERFORM HWND
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the native window handle for a loaded UserForm
'
' WHY THIS EXISTS
'   Some DatePicker visual and positioning operations require the native Windows
'   window handle of the UserForm. This lookup must remain centralized so all
'   WinAPI-dependent behavior uses the same Mac-safe and fail-safe handle policy
'
' INPUTS
'   Frm
'     UserForm instance whose native window handle should be resolved
'
' RETURNS
'   Native UserForm window handle on Windows
'   Zero when running on Mac, when WinAPI behavior is disabled by setting, when
'   no form is supplied, when the caption is blank, when the handle cannot be
'   found, or when lookup fails
'
' BEHAVIOR
'   Returns zero on Mac. On Windows, validates the supplied form, reads its
'   caption, then attempts to find the UserForm window using the common
'   ThunderDFrame class and the alternate ThunderXFrame class
'
' ERROR POLICY
'   Safe default. Returns zero on lookup failure and writes diagnostics to the
'   Immediate Window
'
' DEPENDENCIES
'   FindWindow
'   ThunderDFrame UserForm window class
'   ThunderXFrame UserForm window class
'   M_Platform_ShouldUseWinAPI
'
' NOTES
'   This routine deliberately avoids raising outward because native handle lookup
'   is required only for optional WinAPI-dependent UI behavior
'
'   The lookup is caption-based because MSForms UserForms do not expose their
'   native window handle directly through the object model
'
'   If multiple loaded UserForms share the same caption, Windows may return the
'   first matching window. DatePicker captions should therefore remain unique
'   while WinAPI behavior is enabled
'
'   This is the single handle resolver for the module. Callers must not perform
'   their own FindWindow lookup, so the caption policy and the blank-caption
'   guard apply to every WinAPI-dependent routine
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then

'------------------------------------------------------------------------------
' MAC SAFE DEFAULT
'------------------------------------------------------------------------------
    'Return zero on Mac
        M_Window_GetUserFormHwnd = 0

#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_GetUserFormHwnd"

    Dim FormCaption             As String        'Current form caption
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return the safe default unless lookup succeeds
        M_Window_GetUserFormHwnd = 0
    'Suppress handle lookup errors through the local fail-safe path
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' CHECK WINAPI CAPABILITY
'------------------------------------------------------------------------------
    'Exit when the current platform cannot use WinAPI calls
        If Not M_Platform_CanUseWinAPI Then Exit Function

'------------------------------------------------------------------------------
' READ FORM CAPTION
'------------------------------------------------------------------------------
    'Read the current form caption
        FormCaption = VBA.CStr(Frm.Caption)
    'Exit if the caption is blank
        If VBA.LenB(FormCaption) = 0 Then Exit Function

'------------------------------------------------------------------------------
' TRY PRIMARY USERFORM WINDOW CLASS
'------------------------------------------------------------------------------
    'Try the most common MSForms UserForm window class
        M_Window_GetUserFormHwnd = FindWindow("ThunderDFrame", FormCaption)

'------------------------------------------------------------------------------
' TRY ALTERNATE USERFORM WINDOW CLASS
'------------------------------------------------------------------------------
    'Try the alternate MSForms UserForm window class
        If M_Window_GetUserFormHwnd = 0 Then
            M_Window_GetUserFormHwnd = FindWindow("ThunderXFrame", FormCaption)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after lookup
        Exit Function

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the lookup error number
        ErrorNumber = Err.Number
    'Capture the lookup error description
        ErrorDescription = Err.Description
    'Return the safe default
        M_Window_GetUserFormHwnd = 0
    'Write diagnostics only when lookup failed with an error
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear the suppressed lookup error
        Err.Clear

#End If

End Function

Public Sub M_Window_MoveFormToMouse( _
    ByVal Frm As Object, _
    Optional ByVal OffsetXPx As Long = 0, _
    Optional ByVal OffsetYPx As Long = 12, _
    Optional ByVal CenterOnMouse As Boolean = False)

'
'------------------------------------------------------------------------------
'                           WINDOW MOVE FORM TO MOUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a DatePicker UserForm near the current mouse position on Windows
'
' WHY THIS EXISTS
'   The DatePicker may be shown as a compact modeless popup. Mouse-based
'   positioning keeps the form close to the user's interaction point while
'   clamping the window inside the active monitor work area
'
' INPUTS
'   Frm
'     UserForm instance to move
'
'   OffsetXPx
'     Horizontal offset in pixels from the mouse position
'
'   OffsetYPx
'     Vertical offset in pixels from the mouse position
'
'   CenterOnMouse
'     True to center the form on the mouse position before applying offsets
'     False to use the mouse position as the top-left anchor before offsets
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits safely on Mac, when no form is supplied, when WinAPI is unavailable,
'   when the mouse position cannot be read, when the form handle cannot be
'   resolved, when the form rectangle cannot be read, or when the monitor work
'   area cannot be resolved
'
'   On Windows, calculates the desired form position from the mouse position,
'   applies the requested offsets, clamps the result to the nearest monitor work
'   area, and moves the form without resizing, changing z-order, or activating it
'
' ERROR POLICY
'   Best-effort positioning. Unexpected WinAPI, monitor, or form-positioning
'   errors are suppressed and written to the Immediate Window for diagnostics
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   M_Window_GetUserFormHwnd
'   GetCursorPos
'   GetWindowRect
'   MonitorFromRect
'   GetMonitorInfo
'   SetWindowPos
'   POINTAPI
'   RECT
'   MONITORINFO
'   MONITOR_DEFAULTTONEAREST
'   SWP_NOSIZE
'   SWP_NOZORDER
'   SWP_NOACTIVATE
'   SWP_SHOWWINDOW
'
' NOTES
'   This routine intentionally does nothing on Mac
'
'   This routine uses M_Platform_CanUseWinAPI, not M_Platform_ShouldUseWinAPI,
'   because mouse positioning is a capability-only behavior. Disabling optional
'   WinAPI styling should disable title-bar removal, not mouse positioning.
'
'   Positioning is pixel-based because WinAPI cursor, window, and monitor APIs
'   operate in screen pixels
'
'   The nearest monitor work area is used so the picker remains visible in
'   multi-monitor setups
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then

'------------------------------------------------------------------------------
' MAC SAFE EXIT
'------------------------------------------------------------------------------
    'Do nothing on Mac
        Exit Sub

#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_MoveFormToMouse"

    Dim CursorPoint             As POINTAPI      'Mouse cursor position
    Dim CursorRect              As RECT          'One-pixel cursor rectangle
    Dim FormRect                As RECT          'Current form window rectangle
    Dim MonitorInfoData         As MONITORINFO   'Nearest monitor information

    #If VBA7 Then
        Dim hWndForm            As LongPtr       'UserForm window handle
        Dim hMonitor            As LongPtr       'Nearest monitor handle
    #Else
        Dim hWndForm            As Long          'UserForm window handle
        Dim hMonitor            As Long          'Nearest monitor handle
    #End If

    Dim FormWidthPx             As Long          'Form width in pixels
    Dim FormHeightPx            As Long          'Form height in pixels
    Dim WorkWidthPx             As Long          'Monitor work-area width in pixels
    Dim WorkHeightPx            As Long          'Monitor work-area height in pixels
    Dim TargetX                 As Long          'Target X position in pixels
    Dim TargetY                 As Long          'Target Y position in pixels
    
    Dim MoveFlags               As Long          'SetWindowPos movement flags
    
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description
    
    Dim ApiResult               As Long          'Generic WinAPI Boolean-style result
    Dim LastApiError            As Long          'WinAPI last-error code

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress positioning errors through the local cleanup path
        On Error GoTo CleanExit

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CHECK WINAPI CAPABILITY
'------------------------------------------------------------------------------
    'Exit when the current platform cannot use WinAPI calls
        If Not M_Platform_CanUseWinAPI Then Exit Sub

'------------------------------------------------------------------------------
' READ MOUSE POSITION
'------------------------------------------------------------------------------
    'Exit if the cursor position cannot be read
        If GetCursorPos(CursorPoint) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE FORM HANDLE
'------------------------------------------------------------------------------
    'Resolve the native UserForm window handle
        hWndForm = M_Window_GetUserFormHwnd(Frm)
    'Exit if the form window handle is unavailable
        If hWndForm = 0 Then Exit Sub

'------------------------------------------------------------------------------
' READ FORM RECTANGLE
'------------------------------------------------------------------------------
    'Exit if the form window rectangle cannot be read
        If GetWindowRect(hWndForm, FormRect) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE FORM SIZE
'------------------------------------------------------------------------------
    'Calculate the form width in pixels
        FormWidthPx = FormRect.Right - FormRect.Left
    'Calculate the form height in pixels
        FormHeightPx = FormRect.Bottom - FormRect.Top
    'Exit if resolved form size is invalid
        If FormWidthPx <= 0 Or FormHeightPx <= 0 Then Exit Sub

'------------------------------------------------------------------------------
' BUILD CURSOR RECTANGLE
'------------------------------------------------------------------------------
    'Set cursor rectangle left edge
        CursorRect.Left = CursorPoint.X
    'Set cursor rectangle top edge
        CursorRect.Top = CursorPoint.Y
    'Set cursor rectangle right edge
        CursorRect.Right = CursorPoint.X + 1
    'Set cursor rectangle bottom edge
        CursorRect.Bottom = CursorPoint.Y + 1

'------------------------------------------------------------------------------
' RESOLVE NEAREST MONITOR
'------------------------------------------------------------------------------
    'Resolve the nearest monitor handle
        hMonitor = MonitorFromRect(CursorRect, MONITOR_DEFAULTTONEAREST)
    'Exit if the monitor handle is unavailable
        If hMonitor = 0 Then Exit Sub

'------------------------------------------------------------------------------
' READ MONITOR WORK AREA
'------------------------------------------------------------------------------
    'Initialize the monitor-info structure size
        MonitorInfoData.cbSize = Len(MonitorInfoData)
    'Exit if monitor information cannot be read
        If GetMonitorInfo(hMonitor, MonitorInfoData) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE WORK AREA SIZE
'------------------------------------------------------------------------------
    'Calculate the monitor work-area width
        WorkWidthPx = MonitorInfoData.rcWork.Right - MonitorInfoData.rcWork.Left
    'Calculate the monitor work-area height
        WorkHeightPx = MonitorInfoData.rcWork.Bottom - MonitorInfoData.rcWork.Top
    'Exit if monitor work area is invalid
        If WorkWidthPx <= 0 Or WorkHeightPx <= 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE TARGET POSITION
'------------------------------------------------------------------------------
    'Use the mouse position as the default top-left anchor
        TargetX = CursorPoint.X + OffsetXPx
    'Use the mouse position as the default top-left anchor
        TargetY = CursorPoint.Y + OffsetYPx
    'Center the form on the mouse when requested
        If CenterOnMouse Then
            TargetX = CursorPoint.X - (FormWidthPx \ 2) + OffsetXPx
            TargetY = CursorPoint.Y - (FormHeightPx \ 2) + OffsetYPx
        End If

'------------------------------------------------------------------------------
' CLAMP TARGET X POSITION
'------------------------------------------------------------------------------
    'Anchor X to the work-area left edge when the form is wider than the work area
        If FormWidthPx >= WorkWidthPx Then
            TargetX = MonitorInfoData.rcWork.Left
    'Clamp X to the work-area left edge
        ElseIf TargetX < MonitorInfoData.rcWork.Left Then
            TargetX = MonitorInfoData.rcWork.Left
    'Clamp X to the work-area right edge
        ElseIf TargetX + FormWidthPx > MonitorInfoData.rcWork.Right Then
            TargetX = MonitorInfoData.rcWork.Right - FormWidthPx
        End If

'------------------------------------------------------------------------------
' CLAMP TARGET Y POSITION
'------------------------------------------------------------------------------
    'Anchor Y to the work-area top edge when the form is taller than the work area
        If FormHeightPx >= WorkHeightPx Then
            TargetY = MonitorInfoData.rcWork.Top
    'Clamp Y to the work-area top edge
        ElseIf TargetY < MonitorInfoData.rcWork.Top Then
            TargetY = MonitorInfoData.rcWork.Top
    'Clamp Y to the work-area bottom edge
        ElseIf TargetY + FormHeightPx > MonitorInfoData.rcWork.Bottom Then
            TargetY = MonitorInfoData.rcWork.Bottom - FormHeightPx
        End If

'------------------------------------------------------------------------------
' MOVE FORM
'------------------------------------------------------------------------------
    'Build movement flags
        MoveFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    'Clear the WinAPI last-error state before moving the form
        SetLastError 0
    'Move the form without resizing, changing z-order, or activating it
        ApiResult = SetWindowPos(hWndForm, 0, TargetX, TargetY, 0, 0, MoveFlags)
    'Diagnose SetWindowPos return-code failure
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=Move form" & _
                " | Api=SetWindowPos" & _
                " | LastError=" & VBA.CStr(LastApiError)
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Capture any suppressed error number
        ErrorNumber = Err.Number
    'Capture any suppressed error description
        ErrorDescription = Err.Description
    'Write diagnostics only when mouse positioning failed
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear any suppressed positioning error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

#End If

End Sub

'
'------------------------------------------------------------------------------
'
'                              RIGHT-CLICK MENU
'
'------------------------------------------------------------------------------
'


Public Sub M_ContextMenu_Update()

'
'------------------------------------------------------------------------------
'                           RIGHT-CLICK MENU UPDATE
'------------------------------------------------------------------------------
' PURPOSE
'   Synchronizes DatePicker right-click menu integration with the current setting
'
' WHY THIS EXISTS
'   Context-menu entries should be added, detected, and removed consistently
'   using the DatePicker stable tag. Centralizing the update policy keeps menu
'   state aligned with persisted settings and avoids duplicate or stale command
'   bar controls
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded, then adds DatePicker right-click entries when
'   gDP_ShowRightClick is True and removes them when gDP_ShowRightClick is False
'
' ERROR POLICY
'   Raises unexpected failures back to the caller with this procedure name as
'   the error source
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_ContextMenu_Add
'   M_ContextMenu_Remove
'   gDP_ShowRightClick
'   Application.CommandBars
'
' NOTES
'   This routine owns the high-level policy decision only
'
'   The detailed add / remove mechanics remain delegated to M_ContextMenu_Add
'   and M_ContextMenu_Remove
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_Update"

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are loaded before reading the right-click feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' APPLY CONTEXT-MENU POLICY
'------------------------------------------------------------------------------
    'Add right-click entries when the feature is enabled
        If gDP_ShowRightClick Then
            M_ContextMenu_Add
    'Otherwise remove DatePicker right-click entries
        Else
            M_ContextMenu_Remove
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
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub

Private Sub M_ContextMenu_Add()

'
'==============================================================================
'                           ADD RIGHT-CLICK MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Adds DatePicker entries to the supported Excel right-click menus
'
' WHY THIS EXISTS
'   DatePicker right-click integration targets more than one Excel context menu
'
'   Centralizing the add operation keeps the supported command-bar list
'   consistent and avoids duplicating menu wiring logic across the project
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Adds the DatePicker right-click entry to the standard cell context menu
'   Adds the DatePicker right-click entry to the table / list range context menu
'   Delegates duplicate-control prevention to M_ContextMenu_AddToCommandBar
'   Delegates safe command-bar lookup to M_ContextMenu_GetCommandBar
'
' ERROR POLICY
'   Raises a descriptive runtime error if right-click menu synchronization fails
'   Preserves the original error number and description
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBar
'   M_ContextMenu_AddToCommandBar
'   Application.CommandBars
'
' NOTES
'   Duplicate-control prevention should remain inside
'   M_ContextMenu_AddToCommandBar because that routine owns command-bar mutation
'   logic and stable-tag checks
'
'   This routine owns only the list of supported right-click command bars
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                    As String = "M_ContextMenu_Add"  'Current procedure name
    Const CELL_COMMAND_BAR_NAME        As String = "Cell"               'Standard cell context menu
    Const LIST_RANGE_COMMAND_BAR_NAME  As String = "List Range Popup"   'Table / list range context menu

    Dim HandlerStep                    As String                        'Current handler step for diagnostics
    Dim ErrorNumber                    As Long                          'Captured error number
    Dim ErrorDescription               As String                        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' ADD STANDARD CELL CONTEXT MENU ENTRY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Add standard cell context menu entry"
    'Add to the standard cell context menu
        M_ContextMenu_AddToCommandBar M_ContextMenu_GetCommandBar(CELL_COMMAND_BAR_NAME)

'------------------------------------------------------------------------------
' ADD TABLE / LIST RANGE CONTEXT MENU ENTRY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Add table / list range context menu entry"
    'Add to the table / list range context menu
        M_ContextMenu_AddToCommandBar M_ContextMenu_GetCommandBar(LIST_RANGE_COMMAND_BAR_NAME)

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
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Right-click menu add failed: " & ErrorDescription

End Sub
Public Sub M_ContextMenu_Remove()

'
'------------------------------------------------------------------------------
'                           REMOVE RIGHT-CLICK MENU
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker entries from the supported Excel right-click menus
'
' WHY THIS EXISTS
'   DatePicker right-click integration targets more than one Excel context menu.
'   Centralizing the remove operation keeps the supported command-bar list
'   consistent and avoids leaving stale DatePicker controls after settings
'   changes, workbook close, add-in unload, or manual reset
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Removes the DatePicker right-click entry from:
'     - the standard cell context menu
'     - the table / list range context menu
'
' ERROR POLICY
'   Delegates best-effort removal to M_ContextMenu_RemoveFromCommandBar
'   Does not normally raise for missing, protected, or stale command-bar controls
'
' DEPENDENCIES
'   M_ContextMenu_GetCommandBar
'   M_ContextMenu_RemoveFromCommandBar
'   Application.CommandBars
'
' NOTES
'   Stable-tag based deletion should remain inside
'   M_ContextMenu_RemoveFromCommandBar because that routine owns the command-bar
'   mutation logic and DatePicker control identification policy
'
'   This routine owns only the list of supported right-click command bars
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const CELL_COMMAND_BAR_NAME         As String = "Cell"
    Const LIST_RANGE_COMMAND_BAR_NAME   As String = "List Range Popup"

'------------------------------------------------------------------------------
' REMOVE STANDARD CELL CONTEXT MENU ENTRY
'------------------------------------------------------------------------------
    'Remove from the standard cell context menu
        M_ContextMenu_RemoveFromCommandBar M_ContextMenu_GetCommandBar(CELL_COMMAND_BAR_NAME)

'------------------------------------------------------------------------------
' REMOVE TABLE / LIST RANGE CONTEXT MENU ENTRY
'------------------------------------------------------------------------------
    'Remove from the table / list range context menu
        M_ContextMenu_RemoveFromCommandBar M_ContextMenu_GetCommandBar(LIST_RANGE_COMMAND_BAR_NAME)

End Sub

Private Function M_ContextMenu_GetCommandBar(ByVal CommandBarName As String) As CommandBar

'
'------------------------------------------------------------------------------
'                           GET COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Safely resolves an Excel command bar by name
'
' WHY THIS EXISTS
'   DatePicker right-click integration depends on Excel context menus. Command
'   bars may be unavailable, renamed, hidden, unsupported in a given host state,
'   or inaccessible in some Excel configurations, so lookup must fail safely
'
' INPUTS
'   CommandBarName
'     Name of the Excel command bar to retrieve
'
' RETURNS
'   Matching CommandBar object when available
'   Nothing when the command bar cannot be resolved
'
' BEHAVIOR
'   Attempts to retrieve Application.CommandBars(CommandBarName) and returns a
'   safe Nothing reference when lookup fails
'
' ERROR POLICY
'   Safe default. Command-bar lookup errors are suppressed and reported to the
'   Immediate Window for diagnostics
'
' DEPENDENCIES
'   Application.CommandBars
'   CommandBar
'
' NOTES
'   This routine does not create command bars
'
'   This routine deliberately does not raise outward because missing context
'   menus should not break DatePicker startup, settings synchronization, or
'   teardown
'
'   Add / remove routines must tolerate a Nothing command-bar reference
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_GetCommandBar"

    Dim LookupErrNumber         As Long          'Captured command-bar lookup error number
    Dim LookupErrDescription    As String        'Captured command-bar lookup error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default
        Set M_ContextMenu_GetCommandBar = Nothing
    'Suppress command-bar lookup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RETRIEVE COMMAND BAR
'------------------------------------------------------------------------------
    'Retrieve the requested command bar
        Set M_ContextMenu_GetCommandBar = Application.CommandBars(CommandBarName)
    'Capture command-bar lookup error number
        LookupErrNumber = Err.Number
    'Capture command-bar lookup error description
        LookupErrDescription = Err.Description
    'Clear any suppressed lookup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when command-bar lookup failed with an error
        If LookupErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | CommandBarName=" & CommandBarName & _
                " | Error=" & VBA.CStr(LookupErrNumber) & _
                " | " & LookupErrDescription
        End If

End Function

Private Sub M_ContextMenu_AddToCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           ADD TO COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Adds the DatePicker command to one Excel command bar
'
' WHY THIS EXISTS
'   DatePicker right-click integration can target multiple Excel context menus.
'   The actual command-bar mutation must remain centralized so duplicate
'   detection, stable tagging, button placement, and button configuration stay
'   consistent across all supported menus
'
' INPUTS
'   TargetCommandBar
'     CommandBar object that should receive the DatePicker entry
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when no command bar is supplied, exits when the DatePicker control is
'   already present, otherwise adds a temporary command-bar button and configures
'   its action, icon, caption, tag, and grouping
'
' ERROR POLICY
'   Raises unexpected command-bar mutation failures back to the caller with this
'   procedure name as the error source
'
' DEPENDENCIES
'   M_ContextMenu_ContainsDatePicker
'   M_GetQualifiedMacroName
'   CommandBar.Controls.Add
'   CommandBarButton
'   msoControlButton
'   DP_CONTEXT_MENU_BEFORE
'   DP_CONTEXT_MENU_FACEID
'   DP_CONTEXT_MENU_CAPTION
'   DP_CONTEXT_MENU_TAG
'   DP_Click
'
' NOTES
'   Duplicate prevention is based on M_ContextMenu_ContainsDatePicker
'
'   The control is added as Temporary so Excel does not persist it permanently
'   into the command-bar customization layer
'
'   The stable tag is the authoritative identifier used later for cleanup
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_AddToCommandBar"

    Dim ButtonControl           As CommandBarButton     'Created command-bar button
    Dim ErrorNumber             As Long                 'Captured error number
    Dim ErrorDescription        As String               'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' AVOID DUPLICATE CONTROL
'------------------------------------------------------------------------------
    'Exit when the DatePicker control is already present
        If M_ContextMenu_ContainsDatePicker(TargetCommandBar) Then Exit Sub

'------------------------------------------------------------------------------
' ADD BUTTON
'------------------------------------------------------------------------------
    'Add the DatePicker button
        Set ButtonControl = TargetCommandBar.Controls.Add( _
            Type:=msoControlButton, _
            Before:=DP_CONTEXT_MENU_BEFORE, _
            Temporary:=True)

'------------------------------------------------------------------------------
' CONFIGURE BUTTON
'------------------------------------------------------------------------------
    'Configure the DatePicker button
        With ButtonControl
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .FaceId = DP_CONTEXT_MENU_FACEID
            .Caption = DP_CONTEXT_MENU_CAPTION
            .Tag = DP_CONTEXT_MENU_TAG
            .BeginGroup = True
        End With

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release the command-bar button reference
        Set ButtonControl = Nothing
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
    'Release the command-bar button reference
        Set ButtonControl = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub
Private Sub M_ContextMenu_RemoveFromCommandBar(ByVal TargetCommandBar As CommandBar)

'
'------------------------------------------------------------------------------
'                           REMOVE FROM COMMAND BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker controls from one Excel command bar
'
' WHY THIS EXISTS
'   DatePicker right-click menu entries are transient Excel UI controls. They
'   must be removable by stable tag so menu cleanup remains reliable after
'   settings changes, workbook close, add-in unload, or runtime repair
'
' INPUTS
'   TargetCommandBar
'     CommandBar object to clean
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when no command bar is supplied
'   Reads the command-bar control count defensively
'   Exits safely when command-bar controls cannot be inspected
'   Scans the command-bar controls backwards
'   Deletes every control whose Tag matches DP_CONTEXT_MENU_TAG
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Missing, protected, stale, or otherwise non-readable command-bar controls
'   are skipped and logged to the Immediate Window
'
'   Missing, protected, stale, or otherwise non-removable command-bar controls
'   are skipped and logged to the Immediate Window
'
'   Does not raise outward
'
' DEPENDENCIES
'   CommandBar
'   CommandBarControl
'   DP_CONTEXT_MENU_TAG
'
' NOTES
'   The backward scan is intentional because controls are deleted while the
'   collection is being traversed
'
'   Matching by Tag is preferred over matching by Caption because captions may
'   change with localization or UI wording
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_RemoveFromCommandBar" 'Current procedure name

    Dim Index                   As Long       'CommandBar control index
    Dim ControlCount            As Long       'CommandBar control count
    Dim ControlTag              As String     'Current control tag
    Dim CountErrNumber          As Long       'Captured controls-count error number
    Dim CountErrDescription     As String     'Captured controls-count error description
    Dim DeleteErrNumber         As Long       'Captured delete error number
    Dim DeleteErrDescription    As String     'Captured delete error description
    Dim ReadErrNumber           As Long       'Captured tag-read error number
    Dim ReadErrDescription      As String     'Captured tag-read error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors because this is a best-effort cleanup routine
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' READ CONTROL COUNT
'------------------------------------------------------------------------------
    'Clear any pending error before reading the controls count
        Err.Clear
    'Read the command-bar controls count defensively
        ControlCount = TargetCommandBar.Controls.Count
    'Capture controls-count error number
        CountErrNumber = Err.Number
    'Capture controls-count error description
        CountErrDescription = Err.Description
    'Clear any suppressed controls-count error
        Err.Clear
    'Exit safely when the controls collection cannot be inspected
        If CountErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=ReadControlsCount" & _
                " | Error=" & VBA.CStr(CountErrNumber) & _
                " | " & CountErrDescription
            GoTo ExitProcedure
        End If

'------------------------------------------------------------------------------
' REMOVE DATEPICKER CONTROLS
'------------------------------------------------------------------------------
    'Loop backward through controls so deletion does not disturb pending indexes
        For Index = ControlCount To 1 Step -1
            'Reset the current control tag
                ControlTag = vbNullString
            'Reset tag-read diagnostics
                ReadErrNumber = 0
            'Reset tag-read diagnostic description
                ReadErrDescription = vbNullString
            'Clear any pending error before reading the tag
                Err.Clear
            'Read the current control tag
                ControlTag = VBA.CStr(TargetCommandBar.Controls(Index).Tag)
            'Capture tag-read error number
                ReadErrNumber = Err.Number
            'Capture tag-read error description
                ReadErrDescription = Err.Description
            'Clear any suppressed tag-read error
                Err.Clear
            'Write diagnostics only when tag read failed
                If ReadErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=ReadTag" & _
                        " | Index=" & VBA.CStr(Index) & _
                        " | Error=" & VBA.CStr(ReadErrNumber) & _
                        " | " & ReadErrDescription
                End If
            'Delete DatePicker controls by stable tag
                If VBA.StrComp(ControlTag, DP_CONTEXT_MENU_TAG, vbBinaryCompare) = 0 Then
                    'Reset delete diagnostics
                        DeleteErrNumber = 0
                    'Reset delete diagnostic description
                        DeleteErrDescription = vbNullString
                    'Clear any pending error before deleting the control
                        Err.Clear
                    'Delete the matching DatePicker control
                        TargetCommandBar.Controls(Index).Delete
                    'Capture delete error number
                        DeleteErrNumber = Err.Number
                    'Capture delete error description
                        DeleteErrDescription = Err.Description
                    'Clear any suppressed delete error
                        Err.Clear
                    'Write diagnostics only when deletion failed
                        If DeleteErrNumber <> 0 Then
                            Debug.Print PROC_NAME & _
                                " | Step=DeleteControl" & _
                                " | Index=" & VBA.CStr(Index) & _
                                " | Error=" & VBA.CStr(DeleteErrNumber) & _
                                " | " & DeleteErrDescription
                        End If
                End If
        Next Index

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
Private Function M_ContextMenu_ContainsDatePicker(ByVal TargetCommandBar As CommandBar) As Boolean

'
'------------------------------------------------------------------------------
'                           CONTAINS DATEPICKER CONTROL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether one Excel command bar already contains a DatePicker control
'
' WHY THIS EXISTS
'   DatePicker right-click menu integration must avoid duplicate controls when
'   menus are refreshed, settings are reapplied, workbooks are activated, or the
'   runtime is repaired
'
' INPUTS
'   TargetCommandBar
'     CommandBar object to inspect
'
' RETURNS
'   True when a control tagged with DP_CONTEXT_MENU_TAG is found
'   False when no matching control is found or the command bar cannot be scanned
'
' BEHAVIOR
'   Exits safely when no command bar is supplied, scans each command-bar control,
'   reads the control Tag defensively, and returns True as soon as the stable
'   DatePicker tag is found
'
' ERROR POLICY
'   Safe-default predicate
'
'   Unexpected command-bar or control-inspection errors return False and are
'   written to the Immediate Window for diagnostics
'
' DEPENDENCIES
'   CommandBar
'   CommandBarControl
'   DP_CONTEXT_MENU_TAG
'
' NOTES
'   Matching by Tag is preferred over matching by Caption because captions may
'   change with localization or UI wording
'
'   This function does not mutate the command bar
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_ContextMenu_ContainsDatePicker"

    Dim ControlItem             As CommandBarControl     'CommandBar control being inspected
    Dim ControlTag              As String                'Current control tag
    Dim ReadErrNumber           As Long                  'Captured tag-read error number
    Dim ReadErrDescription      As String                'Captured tag-read error description
    Dim ErrorNumber             As Long                  'Captured runtime error number
    Dim ErrorDescription        As String                'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default
        M_ContextMenu_ContainsDatePicker = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no command bar is supplied
        If TargetCommandBar Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' SCAN COMMAND-BAR CONTROLS
'------------------------------------------------------------------------------
    'Loop through command-bar controls
        For Each ControlItem In TargetCommandBar.Controls
            'Reset current control tag
                ControlTag = vbNullString
            'Reset tag-read diagnostics
                ReadErrNumber = 0
            'Reset tag-read diagnostic description
                ReadErrDescription = vbNullString
            'Suppress tag-read errors for unusual command-bar controls
                On Error Resume Next
            'Read the current control tag
                ControlTag = VBA.CStr(ControlItem.Tag)
            'Capture tag-read error number
                ReadErrNumber = Err.Number
            'Capture tag-read error description
                ReadErrDescription = Err.Description
            'Clear any suppressed tag-read error
                Err.Clear
            'Restore safe-default error handling
                On Error GoTo ErrorHandler
            'Write diagnostics only when tag read failed
                If ReadErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=ReadTag" & _
                        " | Error=" & VBA.CStr(ReadErrNumber) & _
                        " | " & ReadErrDescription
                End If
            'Return True when the DatePicker stable tag is found
                If VBA.StrComp(ControlTag, DP_CONTEXT_MENU_TAG, vbBinaryCompare) = 0 Then
                    M_ContextMenu_ContainsDatePicker = True
                    Exit Function
                End If
        Next ControlItem

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release object reference
        Set ControlItem = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Release object reference
        Set ControlItem = Nothing
    'Return the safe default
        M_ContextMenu_ContainsDatePicker = False
    'Write diagnostics without interrupting context-menu synchronization
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed predicate error
        Err.Clear

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
'   This routine owns only the setting-driven synchronization policy
'
'   The actual Application.OnKey assignment and removal remain delegated to the
'   dedicated register / remove routines
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Update"

    Dim HandlerStep             As String       'Current handler step for diagnostics
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Ensure settings are available before reading the keyboard shortcut flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' SYNCHRONIZE SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Synchronize keyboard shortcut"
    'Register the keyboard shortcut when the feature is enabled
        If gDP_EnableKeyboardShortcut Then
            M_KeyboardShortcut_Register
    'Otherwise restore the key to Excel default handling
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard shortcut synchronization failed: " & ErrorDescription

End Sub


Private Sub M_KeyboardShortcut_Register()

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
'   Assigns Ctrl + Shift + D to the workbook-qualified DP_OpenForActiveCell
'   callback
'
' ERROR POLICY
'   Raises a descriptive runtime error if the shortcut cannot be registered
'
' DEPENDENCIES
'   Application.OnKey
'   M_GetQualifiedMacroName
'   DP_OpenForActiveCell
'   DP_KEYBOARD_SHORTCUT_KEY
'
' NOTES
'   The callback is workbook-qualified so the shortcut resolves to this project
'
'   Application.OnKey is application-wide for the current Excel session
'
'   The shortcut should be removed during teardown through
'   M_KeyboardShortcut_Remove
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Register"

    Dim CallbackMacroName       As String       'Workbook-qualified callback macro name
    Dim HandlerStep             As String       'Current handler step for diagnostics
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RESOLVE CALLBACK MACRO NAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve callback macro name"
    'Build the workbook-qualified DatePicker launcher macro name
        CallbackMacroName = M_GetQualifiedMacroName("DP_OpenForActiveCell")

'------------------------------------------------------------------------------
' REGISTER SHORTCUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Register keyboard shortcut"
    'Assign Ctrl + Shift + D to the DatePicker public launcher
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY, CallbackMacroName

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
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Keyboard shortcut registration failed: " & ErrorDescription

End Sub

Public Sub M_KeyboardShortcut_Remove()

'
'------------------------------------------------------------------------------
'                         REMOVE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the DatePicker keyboard shortcut to Excel default handling
'
' WHY THIS EXISTS
'   Application.OnKey assignments are application-wide. If the workbook or add-in
'   is closed, reset, or unloaded, the shortcut must not remain bound to a macro
'   that may no longer exist
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
'   Shortcut-removal failures are suppressed and written to the Immediate Window
'   for diagnostics
'
' DEPENDENCIES
'   Application.OnKey
'   DP_KEYBOARD_SHORTCUT_KEY
'
' NOTES
'   Calling Application.OnKey with only the key argument restores normal Excel
'   behavior for that key combination
'
'   This routine intentionally does not call M_Settings_EnsureLoaded because it
'   is a teardown-safe cleanup routine
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_KeyboardShortcut_Remove"

    Dim RemoveErrNumber         As Long         'Captured shortcut-removal error number
    Dim RemoveErrDescription    As String       'Captured shortcut-removal error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors because shortcut removal is best-effort
        On Error Resume Next

'------------------------------------------------------------------------------
' REMOVE SHORTCUT
'------------------------------------------------------------------------------
    'Restore the key to Excel default handling
        Application.OnKey DP_KEYBOARD_SHORTCUT_KEY

'------------------------------------------------------------------------------
' CAPTURE DIAGNOSTICS
'------------------------------------------------------------------------------
    'Capture shortcut-removal error number
        RemoveErrNumber = Err.Number
    'Capture shortcut-removal error description
        RemoveErrDescription = Err.Description
    'Clear any suppressed shortcut-removal error
        Err.Clear

'------------------------------------------------------------------------------
' WRITE DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write diagnostics only when shortcut removal failed
        If RemoveErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=Application.OnKey" & _
                " | Key=" & DP_KEYBOARD_SHORTCUT_KEY & _
                " | Error=" & VBA.CStr(RemoveErrNumber) & _
                " | " & RemoveErrDescription
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
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



Private Function M_GridIcon_BuildAnchorKey(ByVal AnchorCell As Excel.Range) As String

'
'------------------------------------------------------------------------------
'                           BUILD GRID ICON ANCHOR KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Builds a stable key for the current grid-icon anchor cell
'
' WHY THIS EXISTS
'   The in-grid icon can receive repeated selection-change refreshes for the same
'   cell. A compact target key allows the high-frequency path to detect that the
'   icon is already shown at the correct target and exit early.
'
' INPUTS
'   AnchorCell
'     Normalized one-cell anchor range
'
' RETURNS
'   External cell address key, or vbNullString when the anchor cannot be resolved
'
' BEHAVIOR
'   Returns the external A1 address of the supplied anchor cell
'
' ERROR POLICY
'   Safe-default helper. Returns vbNullString on failure
'
' DEPENDENCIES
'   Excel.Range.Address
'
' NOTES
'   The caller is responsible for normalizing merged cells before calling this
'   helper
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return a safe default on failure
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no anchor cell is supplied
        If AnchorCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' RETURN ANCHOR KEY
'------------------------------------------------------------------------------
    'Return a workbook-qualified worksheet address
        M_GridIcon_BuildAnchorKey = AnchorCell.Address( _
            RowAbsolute:=False, _
            ColumnAbsolute:=False, _
            ReferenceStyle:=xlA1, _
            External:=True)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful key creation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default
        M_GridIcon_BuildAnchorKey = vbNullString

End Function

Private Sub M_GridIcon_ClearLastTarget()

'
'------------------------------------------------------------------------------
'                         CLEAR GRID ICON LAST TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the cached last grid-icon target
'
' WHY THIS EXISTS
'   The same-target short-circuit must be invalidated when the tracked shape is
'   deleted, purged, or found stale
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the cached anchor key and cached icon position
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR CACHE
'------------------------------------------------------------------------------
    'Clear the cached anchor key
        mDP_GridIconLastAnchorKey = vbNullString
    'Clear the cached left position
        mDP_GridIconLastLeft = 0#
    'Clear the cached top position
        mDP_GridIconLastTop = 0#

End Sub

Private Sub M_GridIcon_RememberTarget( _
    ByVal AnchorKey As String, _
    ByVal IconLeft As Double, _
    ByVal IconTop As Double)

'
'------------------------------------------------------------------------------
'                         REMEMBER GRID ICON TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Stores the last successfully positioned grid-icon target
'
' WHY THIS EXISTS
'   The cached target allows M_GridIcon_ShowOrMove to skip unnecessary shape
'   property writes when Excel fires repeated refreshes for the same cell
'
' INPUTS
'   AnchorKey
'     Stable anchor-cell key
'
'   IconLeft
'     Icon left position in points
'
'   IconTop
'     Icon top position in points
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the supplied key and position when the key is not blank
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' STORE CACHE
'------------------------------------------------------------------------------
    'Clear cache when the supplied key is blank
        If VBA.LenB(AnchorKey) = 0 Then
            M_GridIcon_ClearLastTarget
            Exit Sub
        End If

    'Store the anchor key
        mDP_GridIconLastAnchorKey = AnchorKey
    'Store the icon left position
        mDP_GridIconLastLeft = IconLeft
    'Store the icon top position
        mDP_GridIconLastTop = IconTop

End Sub

Private Function M_GridIcon_IsSameVisibleTarget( _
    ByVal AnchorKey As String, _
    ByVal TargetSheet As Excel.Worksheet, _
    ByVal IconLeft As Double, _
    ByVal IconTop As Double, _
    ByVal IconSize As Double) As Boolean

'
'------------------------------------------------------------------------------
'                       GRID ICON IS SAME VISIBLE TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the tracked grid icon is already visible at the requested
'   target
'
' WHY THIS EXISTS
'   SelectionChange can fire repeatedly while the DatePicker icon is already
'   positioned correctly. In that case, moving the shape, resetting properties,
'   and bringing it to front again is unnecessary work.
'
' INPUTS
'   AnchorKey
'     Stable anchor-cell key
'
'   TargetSheet
'     Worksheet that should own the icon
'
'   IconLeft
'     Expected icon left position
'
'   IconTop
'     Expected icon top position
'
'   IconSize
'     Expected icon width and height
'
' RETURNS
'   True when the tracked icon is already valid, visible, and correctly placed
'
'   False otherwise
'
' BEHAVIOR
'   Checks the cached target key first, then validates the tracked shape, parent
'   worksheet, visibility, position, and size
'
' ERROR POLICY
'   Safe-default predicate. Returns False and clears stale tracking state on
'   failure
'
' DEPENDENCIES
'   gDP_GridIconShape
'   M_GridIcon_ClearLastTarget
'
' NOTES
'   Position comparisons use a small tolerance to avoid point-rounding noise
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const POSITION_TOLERANCE    As Double = 0.05       'Point comparison tolerance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return safe default unless all checks pass
        M_GridIcon_IsSameVisibleTarget = False
    'Enable safe-default error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE CACHE
'------------------------------------------------------------------------------
    'Exit when the requested key is blank
        If VBA.LenB(AnchorKey) = 0 Then Exit Function
    'Exit when the target sheet is missing
        If TargetSheet Is Nothing Then Exit Function
    'Exit when the tracked icon is missing
        If gDP_GridIconShape Is Nothing Then Exit Function
    'Exit when the cached target key differs
        If VBA.StrComp(mDP_GridIconLastAnchorKey, AnchorKey, vbBinaryCompare) <> 0 Then Exit Function
    'Exit when the cached left position differs
        If VBA.Abs(mDP_GridIconLastLeft - IconLeft) > POSITION_TOLERANCE Then Exit Function
    'Exit when the cached top position differs
        If VBA.Abs(mDP_GridIconLastTop - IconTop) > POSITION_TOLERANCE Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Exit when the tracked icon belongs to another worksheet
        If Not (gDP_GridIconShape.Parent Is TargetSheet) Then Exit Function
    'Exit when the tracked icon is hidden
        If gDP_GridIconShape.Visible <> msoTrue Then Exit Function
    'Exit when the tracked icon left position differs
        If VBA.Abs(gDP_GridIconShape.Left - IconLeft) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon top position differs
        If VBA.Abs(gDP_GridIconShape.Top - IconTop) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon width differs
        If VBA.Abs(gDP_GridIconShape.Width - IconSize) > POSITION_TOLERANCE Then Exit Function
    'Exit when the tracked icon height differs
        If VBA.Abs(gDP_GridIconShape.Height - IconSize) > POSITION_TOLERANCE Then Exit Function

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return True because the existing visible icon already matches the target
        M_GridIcon_IsSameVisibleTarget = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful validation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Suppress stale-shape cleanup errors
        On Error Resume Next
    'Clear the tracked shape reference
        Set gDP_GridIconShape = Nothing
    'Clear the cached target
        M_GridIcon_ClearLastTarget
    'Return safe default
        M_GridIcon_IsSameVisibleTarget = False
    'Restore normal error handling
        On Error GoTo 0

End Function

Public Sub M_GridIcon_ShowOrMove(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         SHOW OR MOVE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or moves the DatePicker in-grid icon by reusing an existing worksheet
'   shape whenever possible
'
' WHY THIS EXISTS
'   SelectionChange can fire very frequently
'
'   Deleting and recreating the worksheet icon on every eligible cell selection
'   creates visible latency because Excel must rebuild and repaint the drawing-
'   layer shape
'
'   This routine is the optimized high-frequency path: it moves a reusable icon
'   when possible and creates a new icon only when needed
'
' INPUTS
'   TargetCell
'     Optional anchor range
'
'     If omitted, Excel.Application.ActiveCell is used
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Removes the icon and exits when the grid-icon feature is disabled
'   Resolves the target cell from TargetCell or ActiveCell
'   Normalizes the target to the first physical cell
'   Normalizes merged cells to the top-left cell of the merge area
'   Resolves the target worksheet
'   Calculates the icon position
'   Reuses the tracked icon when it belongs to the target worksheet
'   Deletes the tracked icon when it belongs to another worksheet
'   Reuses a same-named target-sheet icon when available
'   Updates only changed icon properties where practical
'   Creates the icon only when no reusable target-sheet shape exists
'
' ERROR POLICY
'   Best-effort UI routine
'
'   Suppresses failures, clears stale tracked references, falls back to the cold
'   creation path when possible, and writes diagnostics to the Immediate Window
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_GridIcon_Hide
'   M_GridIcon_Create
'   M_GetQualifiedMacroName
'   DP_GRID_ICON_NAME
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'   Excel Shapes object model
'
' NOTES
'   This is the preferred high-frequency SelectionChange path
'
'   M_GridIcon_Create remains the cold-path creation routine
'
'   The manager refresh path should call this routine, not M_GridIcon_Create
'   directly, otherwise the reuse / move optimization is bypassed
'
'   Target normalization is intentionally aligned with M_GridIcon_Create
'
'   When a multi-cell range is supplied, the first physical cell is used as the
'   icon anchor
'
'   The icon is still brought to front on each move because other worksheet
'   shapes may have changed z-order since the last selection event
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_ShowOrMove"
    Const ICON_SIZE             As Double = 24#                     'Grid icon size in points
    Const ICON_GAP              As Double = 5#                      'Gap between target cell and icon
    Const ICON_ALT_TEXT         As String = "DatePicker Grid Entry Point" 'Grid icon alternative text

    Dim TargetSheet             As Excel.Worksheet   'Worksheet receiving the icon
    Dim AnchorCell              As Excel.Range       'Resolved anchor cell
    Dim AnchorKey               As String            'Stable same-target cache key
    
    Dim CandidateShape          As Excel.Shape       'Existing reusable icon candidate
    Dim IconLeft                As Double            'Icon left position
    Dim IconTop                 As Double            'Icon top position
    Dim HasReusableIcon         As Boolean           'True when an existing shape can be moved
    
    Dim MergeState              As Variant           'Target merge-state snapshot
    Dim CallbackMacroName       As String            'Workbook-qualified icon callback
    
    Dim HandlerStep             As String            'Current handler step for diagnostics
    Dim ErrorNumber             As Long              'Captured error number
    Dim ErrorDescription        As String            'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Load settings before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check grid-icon feature flag"
    'Remove any stale icon and exit when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve anchor cell"
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
    'Otherwise use ActiveCell safely
        Else
            On Error Resume Next
            Set AnchorCell = Excel.Application.ActiveCell
            Err.Clear
            On Error GoTo FailSafe
        End If
    'Remove stale icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' NORMALIZE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize anchor cell"
    'Use the first physical cell defensively
        Set AnchorCell = AnchorCell.Cells(1, 1)
    'Capture merge state safely
        MergeState = AnchorCell.MergeCells
    'Normalize merged targets to the top-left cell of the merge area
        If Not VBA.IsError(MergeState) Then
            If Not VBA.IsNull(MergeState) Then
                If VBA.CBool(MergeState) Then
                    Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
                End If
            End If
        End If
    'Remove stale icon and exit when the target is not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE TARGET WORKSHEET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve target worksheet"
    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet
    'Remove stale icon and exit when no worksheet is available
        If TargetSheet Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CALCULATE ICON POSITION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Calculate icon position"
    'Position the icon to the right of the anchor cell
        IconLeft = AnchorCell.Left + AnchorCell.Width + ICON_GAP
    'Vertically center the icon against the anchor cell
        IconTop = AnchorCell.Top + ((AnchorCell.Height - ICON_SIZE) / 2)

'------------------------------------------------------------------------------
' SHORT-CIRCUIT SAME TARGET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check same-target grid icon"
    'Build the stable anchor key
        AnchorKey = M_GridIcon_BuildAnchorKey(AnchorCell)
    'Exit when the tracked icon is already visible at the requested target
        If M_GridIcon_IsSameVisibleTarget( _
            AnchorKey, _
            TargetSheet, _
            IconLeft, _
            IconTop, _
            ICON_SIZE) Then
            GoTo CleanExit
        End If
        
'------------------------------------------------------------------------------
' RESOLVE CALLBACK MACRO
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve callback macro"
    'Build the workbook-qualified icon callback once
        CallbackMacroName = M_GetQualifiedMacroName("DP_Click")

'------------------------------------------------------------------------------
' RESOLVE TRACKED ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve tracked icon"
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
' VALIDATE TRACKED ICON PARENT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate tracked icon parent"
    'Validate the tracked icon parent when a candidate exists
        If Not CandidateShape Is Nothing Then
            'Suppress stale parent-reference failures
                On Error Resume Next
            'Resolve whether the tracked icon belongs to the target worksheet
                HasReusableIcon = (CandidateShape.Parent Is TargetSheet)
            'Clear invalid tracked references when parent inspection fails
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CandidateShape = Nothing
                    Set gDP_GridIconShape = Nothing
                    HasReusableIcon = False
            'Delete a tracked icon from another worksheet
                ElseIf Not HasReusableIcon Then
                    CandidateShape.Delete
                    Err.Clear
                    Set CandidateShape = Nothing
                    Set gDP_GridIconShape = Nothing
                End If
            'Restore fail-safe error handling
                On Error GoTo FailSafe
        End If

'------------------------------------------------------------------------------
' RESOLVE SAME-SHEET FALLBACK ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve same-sheet icon"
    'Try to reuse a same-named shape on the target worksheet
        If CandidateShape Is Nothing Then
            'Suppress missing-shape lookup errors
                On Error Resume Next
            'Resolve a same-named shape from the target worksheet
                Set CandidateShape = TargetSheet.Shapes(DP_GRID_ICON_NAME)
            'Resolve whether a reusable same-sheet icon was found
                HasReusableIcon = Not (CandidateShape Is Nothing)
            'Clear lookup failures
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
    'Track the current handler step
        HandlerStep = "Move reusable icon"
    'Move and show the existing icon when it can be reused
        If HasReusableIcon Then
            With CandidateShape
                If .Left <> IconLeft Then .Left = IconLeft
                If .Top <> IconTop Then .Top = IconTop
                If .Width <> ICON_SIZE Then .Width = ICON_SIZE
                If .Height <> ICON_SIZE Then .Height = ICON_SIZE
                If .Placement <> xlMove Then .Placement = xlMove
                If .AlternativeText <> ICON_ALT_TEXT Then .AlternativeText = ICON_ALT_TEXT
                If .OnAction <> CallbackMacroName Then .OnAction = CallbackMacroName
                If .Visible <> msoTrue Then .Visible = msoTrue
                .ZOrder msoBringToFront
            End With
            'Store the reusable icon reference
                Set gDP_GridIconShape = CandidateShape
            'Remember the successfully positioned target
                M_GridIcon_RememberTarget AnchorKey, IconLeft, IconTop
            'Exit after moving the icon
                GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CREATE ICON ONLY WHEN NEEDED
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Create icon"
    'Create the icon only when no reusable shape exists
        M_GridIcon_Create AnchorCell
    'Remember the target when cold creation succeeded
        If Not gDP_GridIconShape Is Nothing Then
            M_GridIcon_RememberTarget AnchorKey, IconLeft, IconTop
        End If
        
'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release local references
        Set CandidateShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any non-fatal pending error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
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
    'Clear the cached last target after a grid-icon failure
        M_GridIcon_ClearLastTarget
    'Fall back to cold-path creation when an anchor cell is available
        If Not AnchorCell Is Nothing Then
            M_GridIcon_Create AnchorCell
        End If
    'Write diagnostics without interrupting worksheet interaction
        Debug.Print PROC_NAME & _
            " | Step=" & HandlerStep & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Release local references
        Set CandidateShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed fallback error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Sub M_GridIcon_Hide()

'
'==============================================================================
'                           HIDE CURRENT GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the currently tracked DatePicker in-grid icon without deleting it
'
' WHY THIS EXISTS
'   SelectionChange can fire very frequently
'
'   When the active cell is not eligible for the DatePicker, deleting the icon
'   forces the next eligible date-cell selection to recreate the worksheet shape
'
'   Hiding the existing icon keeps the drawing-layer shape available for fast
'   reuse when the user moves back to an eligible date cell
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Hides the tracked grid icon shape when available
'   Leaves the tracked shape reference intact when hiding succeeds
'   Clears the tracked shape reference only when the reference is stale
'   Does not scan worksheets or workbooks
'   Does not delete shapes
'
' ERROR POLICY
'   Best-effort high-frequency UI routine
'   Suppresses stale-shape and visibility errors
'   Does not raise outward
'
' DEPENDENCIES
'   gDP_GridIconShape
'
' NOTES
'   Use this routine from high-frequency selection-change paths when the current
'   target is not DatePicker-eligible
'
'   Use M_GridIcon_Remove or M_GridIcon_PurgeAll for hard cleanup boundaries
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress high-frequency UI cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' HIDE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Hide the tracked icon when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Visible = msoFalse
            'Clear stale tracked references only when hiding failed
                If Err.Number <> 0 Then
                    Err.Clear
                    Set gDP_GridIconShape = Nothing
                End If
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed visibility error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub M_GridIcon_Create(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         CREATE GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Creates or recreates the in-grid DatePicker entry-point shape
'
' WHY THIS EXISTS
'   Users can invoke the DatePicker directly from the worksheet grid without
'   using the right-click menu, keyboard shortcut, or Ribbon
'
' INPUTS
'   TargetCell
'     Optional anchor range
'
'     If omitted, Excel.Application.ActiveCell is used when available
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Hides the current icon and exits when the grid-icon feature is disabled
'   Resolves and normalizes the anchor cell
'   Calculates the icon position
'   Resolves the embedded PNG file as the primary icon source
'   Creates a hidden temporary worksheet shape
'   Applies the embedded PNG picture icon when available
'   Falls back to the built-in font glyph only when the PNG path fails
'   Replaces the previous stable icon only after the new icon is prepared
'   Restores Application.ScreenUpdating deterministically
'
' ERROR POLICY
'   Best-effort UI routine
'
'   Suppresses creation failures, removes partial temporary shapes before stable
'   promotion, restores Application state, clears the tracked icon reference only
'   when stable replacement started but did not finish, and writes diagnostics to
'   the Immediate Window
'
' DEPENDENCIES
'   Excel Shapes object model
'   M_Settings_EnsureLoaded
'   M_GridIcon_Hide
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
'   The embedded PNG is the primary icon source
'
'   The Segoe MDL2 glyph is only a fallback when the embedded PNG cannot be
'   resolved, decoded, written, found, or applied to the shape
'
'   The routine deliberately does not delete the old icon at the beginning of
'   the valid create path because deleting the old icon before the new icon is
'   ready creates visible flicker
'
'   The old icon is replaced only after the new temporary icon has been created,
'   formatted, and filled
'
'   M_GridIcon_ShowOrMove remains the preferred high-frequency selection-change
'   path
'
'   This routine is the cold creation / recreation path
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_Create"
    Const ICON_SIZE             As Double = 24#
    Const ICON_GAP              As Double = 5#
    Const ICON_ALT_TEXT         As String = "DatePicker Grid Entry Point"

    Dim AnchorCell              As Excel.Range      'Resolved anchor cell
    Dim TargetSheet             As Excel.Worksheet  'Worksheet receiving the icon
    Dim NewIconShape            As Excel.Shape      'New temporary icon shape
    Dim IconLeft                As Double           'Icon left position
    Dim IconTop                 As Double           'Icon top position
    Dim IconFilePath            As String           'Resolved embedded icon file path
    Dim TempShapeName           As String           'Temporary icon shape name
    Dim UsedPictureIcon         As Boolean          'True when picture fill succeeds
    Dim StableReplaceStarted    As Boolean          'True once old stable icon replacement begins
    Dim StableIconPromoted      As Boolean          'True once the new icon is named and tracked
    Dim MergeState              As Variant          'Anchor merge-state snapshot
    Dim PreviousScreenUpdating  As Boolean          'Previous ScreenUpdating state
    Dim HasScreenState          As Boolean          'True after ScreenUpdating is captured
    Dim IconErrNumber           As Long             'Embedded icon resolution error number
    Dim IconErrDescription      As String           'Embedded icon resolution error description
    Dim PictureErrNumber        As Long             'Picture-fill error number
    Dim PictureErrDescription   As String           'Picture-fill error description
    Dim HandlerStep             As String           'Current handler step for diagnostics
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable fail-safe error handling
        On Error GoTo FailSafe
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Ensure settings loaded"
    'Load settings before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check grid-icon feature flag"
    'Hide any reusable icon and exit when the feature is disabled
        If Not gDP_ShowGridIcon Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve anchor cell"
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
        Else
            On Error Resume Next
            Set AnchorCell = Excel.Application.ActiveCell
            Err.Clear
            On Error GoTo FailSafe
        End If

    'Hide any reusable icon and exit when no anchor cell is available
        If AnchorCell Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' NORMALIZE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Normalize anchor cell"
    'Use the first physical cell defensively
        Set AnchorCell = AnchorCell.Cells(1, 1)
    'Capture merge state safely
        MergeState = AnchorCell.MergeCells
    'Normalize merged cells to the top-left cell
        If Not VBA.IsError(MergeState) Then
            If Not VBA.IsNull(MergeState) Then
                If VBA.CBool(MergeState) Then
                    Set AnchorCell = AnchorCell.MergeArea.Cells(1, 1)
                End If
            End If
        End If
    'Hide any reusable icon and exit when the target is not a single cell
        If AnchorCell.Cells.CountLarge <> 1 Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE TARGET WORKSHEET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve target worksheet"
    'Store the target worksheet
        Set TargetSheet = AnchorCell.Worksheet
    'Hide any reusable icon and exit when no worksheet is available
        If TargetSheet Is Nothing Then
            M_GridIcon_Hide
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CALCULATE ICON POSITION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Calculate icon position"
    'Position the icon to the right of the anchor cell
        IconLeft = AnchorCell.Left + AnchorCell.Width + ICON_GAP
    'Vertically center the icon against the anchor cell
        IconTop = AnchorCell.Top + ((AnchorCell.Height - ICON_SIZE) / 2)
    'Build the temporary shape name
        TempShapeName = DP_GRID_ICON_NAME & "_Pending"

'------------------------------------------------------------------------------
' RESOLVE EMBEDDED ICON FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve embedded icon file"
    'Start with no resolved picture icon
        IconFilePath = vbNullString
    'Suppress embedded-icon materialization errors and allow fallback only if needed
        On Error Resume Next
    'Resolve or create the embedded temporary icon file
        IconFilePath = M_GridIcon_EnsureEmbeddedIconFile()
    'Write diagnostics when embedded-icon resolution failed
        If Err.Number <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Error=" & VBA.CStr(Err.Number) & _
                " | " & Err.Description
            Err.Clear
        End If
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CAPTURE APPLICATION STATE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Capture Application state"
    'Capture the current ScreenUpdating state
        PreviousScreenUpdating = Excel.Application.ScreenUpdating
    'Mark ScreenUpdating as captured
        HasScreenState = True
    'Suppress worksheet repaint during icon creation
        Excel.Application.ScreenUpdating = False

'------------------------------------------------------------------------------
' REMOVE STALE TEMPORARY SHAPE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Remove stale temporary shape"
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
    'Track the current handler step
        HandlerStep = "Create temporary icon shape"
    'Create the new icon with a temporary name
        Set NewIconShape = TargetSheet.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            IconLeft, _
            IconTop, _
            ICON_SIZE, _
            ICON_SIZE)

'------------------------------------------------------------------------------
' APPLY BASE SHAPE SETTINGS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply base shape settings"
    'Apply base icon behavior and callback settings
        With NewIconShape
            .Name = TempShapeName
            .Placement = xlMove
            .AlternativeText = ICON_ALT_TEXT
            .OnAction = M_GetQualifiedMacroName("DP_Click")
            .LockAspectRatio = msoTrue
            .Fill.Visible = msoTrue
            .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
            .Line.Visible = msoFalse
            .Visible = msoFalse
        End With

'------------------------------------------------------------------------------
' APPLY EMBEDDED PICTURE ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply embedded picture icon"
    'Start from fallback mode
        UsedPictureIcon = False
    'Try to apply the embedded picture icon when a usable file path exists
        If VBA.LenB(IconFilePath) <> 0 Then
            If VBA.LenB(VBA.Dir$(IconFilePath, vbNormal)) <> 0 Then
                'Suppress picture-fill errors so glyph remains an emergency fallback
                    On Error Resume Next
                'Clear any fallback text before applying the picture
                    NewIconShape.TextFrame2.TextRange.Text = vbNullString
                'Apply the embedded PNG as the icon fill
                    NewIconShape.Fill.UserPicture IconFilePath
                'Store whether picture application succeeded
                    UsedPictureIcon = (Err.Number = 0)
                'Write diagnostics when picture application failed
                    If Err.Number <> 0 Then
                        Debug.Print PROC_NAME & _
                            " | Step=" & HandlerStep & _
                            " | IconFilePath=" & IconFilePath & _
                            " | Error=" & VBA.CStr(Err.Number) & _
                            " | " & Err.Description
                        Err.Clear
                    End If
                'Restore fail-safe error handling
                    On Error GoTo FailSafe
            Else
                'Log missing materialized icon file
                    Debug.Print PROC_NAME & _
                        " | Step=" & HandlerStep & _
                        " | Embedded icon file not found: " & IconFilePath
            End If
        Else
            'Log empty icon path so fallback is explainable
                Debug.Print PROC_NAME & _
                    " | Step=" & HandlerStep & _
                    " | Embedded icon path is empty"
        End If

'------------------------------------------------------------------------------
' APPLY FALLBACK GLYPH ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply fallback glyph icon"
    'Apply the built-in fallback glyph only when no picture icon was applied
        If Not UsedPictureIcon Then
            'Apply the base fallback shape appearance
                With NewIconShape
                    .Fill.Visible = msoTrue
                    .Fill.ForeColor.RGB = DP_THEME_COLOR_HEADER
                    .Line.Visible = msoFalse
                End With
            'Suppress glyph-formatting errors because the shape itself remains usable
                On Error Resume Next
            'Clear existing text
                NewIconShape.TextFrame2.TextRange.Text = vbNullString
            'Apply the calendar glyph
                NewIconShape.TextFrame2.TextRange.Text = VBA.ChrW$(DP_GRID_ICON_CALENDAR_CODEPOINT)
            'Apply the glyph font name
                NewIconShape.TextFrame2.TextRange.Font.Name = DP_GRID_ICON_FONT_NAME
            'Apply the glyph font size
                NewIconShape.TextFrame2.TextRange.Font.Size = DP_GRID_ICON_FONT_SIZE
            'Apply the glyph foreground color
                NewIconShape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = DP_GRID_ICON_FORE_COLOR
            'Remove left text margin
                NewIconShape.TextFrame2.MarginLeft = 0
            'Remove right text margin
                NewIconShape.TextFrame2.MarginRight = 0
            'Remove top text margin
                NewIconShape.TextFrame2.MarginTop = 0
            'Remove bottom text margin
                NewIconShape.TextFrame2.MarginBottom = 0
            'Vertically center the glyph
                NewIconShape.TextFrame2.VerticalAnchor = msoAnchorMiddle
            'Horizontally center the glyph
                NewIconShape.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            'Clear any suppressed glyph-formatting error
                Err.Clear
            'Restore fail-safe error handling
                On Error GoTo FailSafe
        End If

'------------------------------------------------------------------------------
' ATOMIC REPLACE OLD ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Replace old icon"
    'Mark that stable icon replacement has started
        StableReplaceStarted = True
    'Suppress old-icon cleanup errors
        On Error Resume Next
    'Delete the tracked old icon when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Delete
        End If
    'Delete any same-named icon on the target sheet
        TargetSheet.Shapes(DP_GRID_ICON_NAME).Delete
    'Clear any suppressed old-icon cleanup error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe
    'Promote the temporary icon to the stable DatePicker icon name
        NewIconShape.Name = DP_GRID_ICON_NAME
    'Store the new icon reference
        Set gDP_GridIconShape = NewIconShape
    'Mark the new icon as safely promoted
        StableIconPromoted = True
    'Remember the successfully promoted icon target
        M_GridIcon_RememberTarget _
            M_GridIcon_BuildAnchorKey(AnchorCell), _
            IconLeft, _
            IconTop

'------------------------------------------------------------------------------
' SHOW FINAL ICON
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Show final icon"
    'Suppress final visual-order failures because the icon has already been created
        On Error Resume Next
    'Show the fully formatted icon
        gDP_GridIconShape.Visible = msoTrue
    'Bring the icon to the front
        gDP_GridIconShape.ZOrder msoBringToFront
    'Clear any suppressed final visual-order error
        Err.Clear
    'Restore fail-safe error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress state-restore failures
        On Error Resume Next
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Excel.Application.ScreenUpdating = PreviousScreenUpdating
        End If
    'Release local references
        Set NewIconShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
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
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Suppress cleanup failures
        On Error Resume Next
    'Delete the partially created temporary icon only before stable promotion
        If Not StableIconPromoted Then
            If Not NewIconShape Is Nothing Then NewIconShape.Delete
        End If
    'Clear the tracked icon reference only when replacement started but did not finish
        If StableReplaceStarted And Not StableIconPromoted Then
            Set gDP_GridIconShape = Nothing
        End If
    'Restore ScreenUpdating when it was captured
        If HasScreenState Then
            Excel.Application.ScreenUpdating = PreviousScreenUpdating
        End If

    'Write diagnostics without interrupting worksheet interaction
        Debug.Print PROC_NAME & _
            " | Step=" & HandlerStep & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Release local references
        Set NewIconShape = Nothing
    'Release local references
        Set TargetSheet = Nothing
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
Public Sub M_GridIcon_PreCreateHidden(Optional ByVal TargetCell As Excel.Range)

'
'==============================================================================
'                         PRE-CREATE HIDDEN GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Pre-creates the DatePicker in-grid icon and hides it for later fast reuse
'
' WHY THIS EXISTS
'   Creating worksheet shapes during SelectionChange can create visible latency
'
'   Pre-creating the icon during startup allows the high-frequency selection path
'   to move and show an existing shape instead of creating one on demand
'
' INPUTS
'   TargetCell
'     Optional initial anchor cell
'
'     If omitted, Excel.Application.ActiveCell is used when available
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures settings are loaded
'   Exits when the grid-icon feature is disabled
'   Creates the icon through the normal cold creation path when no tracked icon
'   exists
'   Hides the icon immediately after creation
'
' ERROR POLICY
'   Best-effort startup optimization
'   Does not raise outward
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_GridIcon_Create
'   gDP_ShowGridIcon
'   gDP_GridIconShape
'
' NOTES
'   This routine is an optimization only
'
'   M_GridIcon_ShowOrMove should still keep M_GridIcon_Create as a fallback
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AnchorCell              As Excel.Range  'Initial icon anchor cell

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress startup optimization errors
        On Error Resume Next

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Ensure settings are available before reading the grid-icon feature flag
        M_Settings_EnsureLoaded

'------------------------------------------------------------------------------
' FEATURE GATE
'------------------------------------------------------------------------------
    'Exit when the grid-icon feature is disabled
        If Not gDP_ShowGridIcon Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' EXIT IF ICON ALREADY EXISTS
'------------------------------------------------------------------------------
    'Exit when a tracked icon already exists
        If Not gDP_GridIconShape Is Nothing Then GoTo HideIcon

'------------------------------------------------------------------------------
' RESOLVE INITIAL ANCHOR
'------------------------------------------------------------------------------
    'Use the supplied target cell when provided
        If Not TargetCell Is Nothing Then
            Set AnchorCell = TargetCell
        Else
            Set AnchorCell = Application.ActiveCell
        End If

    'Exit when no anchor cell is available
        If AnchorCell Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' CREATE ICON
'------------------------------------------------------------------------------
    'Create the icon through the normal cold creation path
        M_GridIcon_Create AnchorCell

'------------------------------------------------------------------------------
' HIDE ICON
'------------------------------------------------------------------------------
HideIcon:
    'Hide the icon after creation so startup does not display it prematurely
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Visible = msoFalse
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release local references
        Set AnchorCell = Nothing
    'Clear any suppressed startup optimization error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub
Public Sub M_GridIcon_Remove()

'
'==============================================================================
'                         REMOVE CURRENT GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the currently tracked DatePicker in-grid icon
'
' WHY THIS EXISTS
'   Selection-change and worksheet-change events can fire very frequently
'
'   The normal icon-removal path must therefore be lightweight and must not scan
'   all worksheets or workbooks
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the tracked DatePicker icon shape when available
'   Clears the tracked icon reference regardless of deletion result
'   Attempts to resolve the active worksheet
'   Attempts to resolve a same-named icon from the active worksheet
'   Deletes the active-sheet fallback icon when found
'   Logs non-trivial cleanup failures to the Immediate Window
'
' ERROR POLICY
'   Best-effort cleanup
'   Suppresses shape-deletion and ActiveSheet-resolution errors
'   Logs tracked-shape deletion failures
'   Logs ActiveSheet resolution failures
'   Logs active-sheet fallback-shape deletion failures
'   Does not log the expected case where no same-named active-sheet shape exists
'   Does not raise outward
'
' DEPENDENCIES
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   Excel Shapes object model
'
' NOTES
'   This routine is intended for high-frequency UI refresh paths
'
'   Missing active-sheet fallback icons are normal and are intentionally not
'   logged
'
'   Use M_GridIcon_PurgeAll for hard cleanup boundaries such as save, close,
'   print, teardown, or regression reset
'
' UPDATED
'   2026-05-06
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "M_GridIcon_Remove" 'Current procedure name

    Dim ActiveSheetObject           As Object       'Current active sheet object
    Dim ActiveWorksheet             As Worksheet    'Current active worksheet
    Dim FallbackShape               As Shape        'Same-named active-sheet fallback shape

    Dim TrackedErrNumber            As Long         'Tracked-shape deletion error number
    Dim TrackedErrDescription       As String       'Tracked-shape deletion error description
    Dim ActiveSheetErrNumber        As Long         'ActiveSheet resolution error number
    Dim ActiveSheetErrDescription   As String       'ActiveSheet resolution error description
    Dim FallbackDeleteErrNumber     As Long         'Fallback-shape deletion error number
    Dim FallbackDeleteErrDescription As String      'Fallback-shape deletion error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' DELETE TRACKED SHAPE
'------------------------------------------------------------------------------
    'Clear any pending error before deleting the tracked shape
        Err.Clear
    'Delete the tracked grid icon shape when available
        If Not gDP_GridIconShape Is Nothing Then
            gDP_GridIconShape.Delete
            TrackedErrNumber = Err.Number
            TrackedErrDescription = Err.Description
            Err.Clear
        End If
    'Clear the tracked shape reference regardless of deletion result
        Set gDP_GridIconShape = Nothing
    'Clear the cached last target
        M_GridIcon_ClearLastTarget

'------------------------------------------------------------------------------
' RESOLVE ACTIVE WORKSHEET
'------------------------------------------------------------------------------
    'Clear any pending error before resolving ActiveSheet
        Err.Clear
    'Capture the active sheet object safely
        Set ActiveSheetObject = Excel.Application.ActiveSheet
    'Capture ActiveSheet resolution failure if any
        ActiveSheetErrNumber = Err.Number
    'Capture ActiveSheet resolution failure description if any
        ActiveSheetErrDescription = Err.Description
    'Clear any suppressed ActiveSheet error
        Err.Clear
    'Use the active sheet only when it is a worksheet
        If Not ActiveSheetObject Is Nothing Then
            If TypeOf ActiveSheetObject Is Excel.Worksheet Then
                Set ActiveWorksheet = ActiveSheetObject
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE ACTIVE-SHEET FALLBACK SHAPE
'------------------------------------------------------------------------------
    'Resolve a same-named icon from the active worksheet when available
        If Not ActiveWorksheet Is Nothing Then
            Set FallbackShape = ActiveWorksheet.Shapes(DP_GRID_ICON_NAME)
            Err.Clear
        End If

'------------------------------------------------------------------------------
' DELETE ACTIVE-SHEET FALLBACK SHAPE
'------------------------------------------------------------------------------
    'Delete the fallback shape only when it was resolved
        If Not FallbackShape Is Nothing Then
            'Clear any pending error before deleting the fallback shape
                Err.Clear
            'Delete the same-named active-sheet fallback icon
                FallbackShape.Delete
            'Capture fallback-shape deletion failure if any
                FallbackDeleteErrNumber = Err.Number
            'Capture fallback-shape deletion failure description if any
                FallbackDeleteErrDescription = Err.Description
            'Clear any suppressed fallback deletion error
                Err.Clear
        End If

'------------------------------------------------------------------------------
' DIAGNOSTICS
'------------------------------------------------------------------------------
    'Write tracked-shape diagnostics only when a tracked deletion failed
        If TrackedErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=DeleteTrackedShape" & _
                " | Error=" & VBA.CStr(TrackedErrNumber) & _
                " | " & TrackedErrDescription
        End If
    'Write ActiveSheet diagnostics only when ActiveSheet resolution failed
        If ActiveSheetErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=ResolveActiveSheet" & _
                " | Error=" & VBA.CStr(ActiveSheetErrNumber) & _
                " | " & ActiveSheetErrDescription
        End If
    'Write fallback-shape diagnostics only when deletion of a resolved shape failed
        If FallbackDeleteErrNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Step=DeleteActiveSheetFallbackShape" & _
                " | Error=" & VBA.CStr(FallbackDeleteErrNumber) & _
                " | " & FallbackDeleteErrDescription
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release local object references
        Set FallbackShape = Nothing
    'Release local object references
        Set ActiveWorksheet = Nothing
    'Release local object references
        Set ActiveSheetObject = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
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
'   Resolves the user's temp directory, builds a stable icon file path, reuses
'   the file if it already exists and appears valid, recreates the file from the
'   embedded Base64 payload when missing or invalid, validates the written file,
'   and returns the final valid file path
'
' ERROR POLICY
'   Raises a descriptive runtime error if the temp path cannot be resolved, the
'   embedded payload cannot be decoded, the icon file cannot be written, or the
'   written icon file fails validation
'
' DEPENDENCIES
'   M_GridIcon_GetTempFolder
'   M_GridIcon_CombinePath
'   M_GridIcon_IconFileIsUsable
'   M_GridIcon_EmbeddedIconBytes
'   M_GridIcon_WriteBytesToFile
'   DP_GRID_ICON_TEMP_FILE_NAME
'
' NOTES
'   The temp file is intentionally not deleted during normal cleanup
'
'   The routine returns an existing valid temp file without decoding the embedded
'   payload again
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_EnsureEmbeddedIconFile"

    Dim TempFolder              As String       'Resolved temp folder
    Dim IconPath                As String       'Resolved icon file path
    Dim IconBytes()             As Byte         'Decoded embedded PNG bytes
    Dim HandlerStep             As String       'Current handler step for diagnostics
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RESOLVE TEMP PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve temp folder"
    'Resolve the user's temp folder
        TempFolder = M_GridIcon_GetTempFolder()
    'Reject an empty temp folder
        If VBA.LenB(TempFolder) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "Temporary folder could not be resolved."
        End If

'------------------------------------------------------------------------------
' BUILD ICON PATH
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Build icon path"
    'Build the full icon path
        IconPath = M_GridIcon_CombinePath(TempFolder, DP_GRID_ICON_TEMP_FILE_NAME)

'------------------------------------------------------------------------------
' REUSE EXISTING FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate existing icon file"
    'Return the existing temp file when it appears usable
        If M_GridIcon_IconFileIsUsable(IconPath) Then
            M_GridIcon_EnsureEmbeddedIconFile = IconPath
            Exit Function
        End If

'------------------------------------------------------------------------------
' DECODE EMBEDDED PAYLOAD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Decode embedded icon payload"
    'Decode the embedded PNG payload
        IconBytes = M_GridIcon_EmbeddedIconBytes()

'------------------------------------------------------------------------------
' WRITE TEMP FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Write embedded icon file"
    'Write the decoded bytes to the temp file
        M_GridIcon_WriteBytesToFile IconPath, IconBytes

'------------------------------------------------------------------------------
' VALIDATE WRITTEN FILE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate written icon file"
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
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release decoded byte-array storage
        Erase IconBytes
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Release decoded byte-array storage
        Erase IconBytes
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            ErrorDescription

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
'   Reads TEMP first, falls back to TMP when TEMP is unavailable, trims the
'   resolved value, removes trailing path separators, and returns the normalized
'   folder path
'
' ERROR POLICY
'   Safe default. Returns an empty string on failure and writes diagnostics to
'   the Immediate Window only when an unexpected runtime error occurs
'
' DEPENDENCIES
'   VBA.Environ$
'
' NOTES
'   The Windows temp folder should already exist
'
'   This routine only resolves the path. It does not validate write permission
'   and does not create folders
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_GridIcon_GetTempFolder"

    Dim TempFolder              As String       'Resolved temporary folder path
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the safe default
        M_GridIcon_GetTempFolder = vbNullString
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE TEMP FOLDER
'------------------------------------------------------------------------------
    'Read TEMP first
        TempFolder = VBA.Trim$(VBA.Environ$("TEMP"))
    'Read TMP when TEMP is unavailable
        If VBA.LenB(TempFolder) = 0 Then
            TempFolder = VBA.Trim$(VBA.Environ$("TMP"))
        End If

'------------------------------------------------------------------------------
' NORMALIZE PATH
'------------------------------------------------------------------------------
    'Remove trailing path separators when present
        Do While VBA.LenB(TempFolder) > 0
            If VBA.Right$(TempFolder, 1) = "\" Or VBA.Right$(TempFolder, 1) = "/" Then
                TempFolder = VBA.Left$(TempFolder, VBA.Len(TempFolder) - 1)
            Else
                Exit Do
            End If
        Loop

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized temp folder path
        M_GridIcon_GetTempFolder = TempFolder

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Return the safe default
        M_GridIcon_GetTempFolder = vbNullString
    'Write diagnostics only when an unexpected error occurred
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear the suppressed error
        Err.Clear

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
'   Repeated temp-file path construction should remain deterministic, readable,
'   and isolated in one helper
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
' BEHAVIOR
'   If FolderPath already ends with a backslash, appends FileName directly
'   If FolderPath does not end with a backslash, inserts one backslash separator
'   Does not trim, validate, normalize, expand, or inspect either input
'
' ERROR POLICY
'   Does not raise custom errors
'   Lets native VBA string behavior apply
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper assumes a Windows Excel environment
'   Empty FolderPath behavior is intentionally preserved from the original code
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const C_PATH_SEPARATOR As String = "\"        'Windows path separator

'------------------------------------------------------------------------------
' RETURN COMBINED PATH
'------------------------------------------------------------------------------
    'Return the direct concatenation when the folder already ends with a separator
        If Right$(FolderPath, Len(C_PATH_SEPARATOR)) = C_PATH_SEPARATOR Then
            M_GridIcon_CombinePath = FolderPath & FileName
        Else
            M_GridIcon_CombinePath = FolderPath & C_PATH_SEPARATOR & FileName
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
'   A previous temp-file write may have failed, produced a missing file, or
'   produced a truncated file
'
'   The DatePicker should recreate the embedded grid-icon file when the existing
'   candidate file is missing, too small, unreadable, or not recognizable as PNG
'
' INPUTS
'   IconPath
'     Candidate PNG path
'
' RETURNS
'   True if the file exists, has a plausible size, and starts with the PNG
'   signature bytes
'
'   False otherwise
'
' BEHAVIOR
'   Rejects blank paths after Trim$
'   Rejects missing files
'   Rejects files smaller than DP_GRID_ICON_TEMP_MIN_BYTES
'   Opens the candidate file in binary read mode
'   Reads the first eight bytes
'   Accepts the file only when the standard PNG signature matches
'
' ERROR POLICY
'   Safe-default predicate
'   Returns False on any failure
'   Suppresses cleanup errors while attempting to close the file handle
'
' DEPENDENCIES
'   DP_GRID_ICON_TEMP_MIN_BYTES
'   Dir$
'   FileLen
'   FreeFile
'
' NOTES
'   This is a lightweight validation, not a full PNG parser
'   The routine intentionally does not validate the full PNG structure
'   The routine intentionally preserves the original path handling behavior
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FileNumber             As Integer      'Free file number
    Dim Signature(1 To 8)       As Byte         'PNG signature bytes

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable safe-default error handling
        On Error GoTo ErrorHandler
    'Default to unusable
        M_GridIcon_IconFileIsUsable = False

'------------------------------------------------------------------------------
' VALIDATE PATH
'------------------------------------------------------------------------------
    'Exit when the path is blank
        If LenB(Trim$(IconPath)) = 0 Then Exit Function
    'Exit when the file does not exist
        If LenB(Dir$(IconPath, vbNormal)) = 0 Then Exit Function
    'Exit when the file is implausibly small
        If FileLen(IconPath) < DP_GRID_ICON_TEMP_MIN_BYTES Then Exit Function

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
    'Return True only when the standard PNG signature matches
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
    'Suppress cleanup errors
        On Error Resume Next
    'Close the file number when one has been assigned
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
'   This helper isolates the disk-write step so the grid-icon creation logic
'   remains focused on path resolution, icon validation, and shape rendering
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
' BEHAVIOR
'   Rejects blank destination paths after Trim$
'   Rejects empty byte payloads
'   Deletes an existing target file before writing
'   Opens the target file in binary write mode
'   Writes the full byte array starting at byte position 1
'   Closes the file handle after writing
'
' ERROR POLICY
'   Raises a descriptive runtime error if validation fails or the file cannot
'   be written
'
'   Attempts to close the file handle before re-raising the original error
'
' DEPENDENCIES
'   M_GridIcon_ByteArrayLength
'   Dir$
'   Kill
'   FreeFile
'
' NOTES
'   Existing temp files are overwritten
'   Parent-folder creation is intentionally not handled here
'   The caller is responsible for passing a valid writable destination folder
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_WriteBytesToFile"

    Dim FileNumber             As Integer      'Free file number
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

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
        If LenB(Dir$(IconPath, vbNormal)) <> 0 Then Kill IconPath

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
    'Capture the original error before cleanup
        ErrorNumber = Err.Number
    'Capture the original error description before cleanup
        ErrorDescription = Err.Description
    'Suppress cleanup errors
        On Error Resume Next
    'Close the file number when one has been assigned
        If FileNumber <> 0 Then Close #FileNumber
    'Restore normal error handling
        On Error GoTo 0
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

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
'   Dynamic byte arrays may be uninitialized
'
'   A defensive length helper keeps binary-write validation safe and avoids
'   repeating fragile LBound / UBound checks around the grid-icon file writer
'
' INPUTS
'   SourceBytes
'     Byte array to inspect
'
' RETURNS
'   Number of bytes in the array
'
'   Zero when the array bounds cannot be read
'
' BEHAVIOR
'   Uses LBound and UBound to calculate the byte count
'   Returns zero when SourceBytes is uninitialized
'   Returns zero when array-bound inspection fails for any reason
'
' ERROR POLICY
'   Safe-default predicate
'   Returns zero on failure
'   Does not raise custom errors
'
' DEPENDENCIES
'   LBound
'   UBound
'
' NOTES
'   This helper assumes a one-dimensional byte array
'   The routine intentionally performs no content validation
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return zero when array bounds cannot be read
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETURN LENGTH
'------------------------------------------------------------------------------
    'Return the byte count from the declared array bounds
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
    'Return the safe default
        M_GridIcon_ByteArrayLength = 0

End Function

Private Function M_GridIcon_Base64LooksValid(ByVal EncodedText As String) As Boolean

'
'==============================================================================
'                         BASE64 LOOKS VALID
'------------------------------------------------------------------------------
' PURPOSE
'   Performs lightweight validation of an embedded Base64 payload
'
' WHY THIS EXISTS
'   MSXML raises a generic parsing error when an invalid Base64 string is assigned
'   to a bin.base64 node
'
'   This helper catches obvious payload corruption before the decode step so the
'   caller can fail with a clearer diagnostic
'
' INPUTS
'   EncodedText
'     Base64 text to validate
'
' RETURNS
'   True when the text looks like a valid Base64 payload
'
'   False when the text is blank, has invalid length, contains unsupported
'   characters, or contains padding in an invalid position
'
' BEHAVIOR
'   Removes common whitespace characters
'   Checks that the remaining length is a multiple of four
'   Allows only A-Z, a-z, 0-9, plus, slash, and equals
'   Allows equals padding only in the final two characters
'
' ERROR POLICY
'   Safe-default predicate
'
'   Returns False on any unexpected validation failure
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is not a full decoder
'
'   It is intentionally a fast pre-flight check before MSXML decoding
'
' UPDATED
'   2026-05-08
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CleanText       As String       'Whitespace-normalized Base64 text
    Dim TextLength      As Long         'Normalized text length
    Dim CharIndex       As Long         'Character loop index
    Dim CurrentChar     As String       'Current character being inspected
    Dim PaddingPosition As Long         'First equals padding position

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_GridIcon_Base64LooksValid = False

    'Enable safe-default validation
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Trim the supplied payload
        CleanText = VBA.Trim$(EncodedText)

    'Remove carriage returns
        CleanText = VBA.Replace(CleanText, vbCr, vbNullString)

    'Remove line feeds
        CleanText = VBA.Replace(CleanText, vbLf, vbNullString)

    'Remove tab characters
        CleanText = VBA.Replace(CleanText, vbTab, vbNullString)

    'Remove spaces
        CleanText = VBA.Replace(CleanText, " ", vbNullString)

'------------------------------------------------------------------------------
' VALIDATE LENGTH
'------------------------------------------------------------------------------
    'Capture normalized text length
        TextLength = VBA.Len(CleanText)

    'Reject an empty payload
        If TextLength = 0 Then Exit Function

    'Reject payloads whose length is not divisible by four
        If TextLength Mod 4 <> 0 Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE CHARACTERS
'------------------------------------------------------------------------------
    'Loop through each Base64 character
        For CharIndex = 1 To TextLength
            'Read the current character
                CurrentChar = VBA.Mid$(CleanText, CharIndex, 1)
            'Validate the current character
                Select Case CurrentChar
                    Case "A" To "Z", "a" To "z", "0" To "9", "+", "/"
                        'Valid non-padding Base64 character
                    Case "="
                        'Store the first padding position
                            If PaddingPosition = 0 Then PaddingPosition = CharIndex
                        'Reject padding before the final two characters
                            If CharIndex < TextLength - 1 Then Exit Function
                    Case Else
                        'Reject unsupported characters
                            Exit Function
                End Select
        Next CharIndex

'------------------------------------------------------------------------------
' VALIDATE PADDING ORDER
'------------------------------------------------------------------------------
    'Reject non-padding characters after padding begins
        If PaddingPosition > 0 Then
            For CharIndex = PaddingPosition To TextLength
                If VBA.Mid$(CleanText, CharIndex, 1) <> "=" Then Exit Function
            Next CharIndex
        End If

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return successful validation
        M_GridIcon_Base64LooksValid = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after successful validation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default on unexpected validation failure
        M_GridIcon_Base64LooksValid = False

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
'   The grid icon should use the embedded PNG as the primary visual asset
'
'   The worksheet glyph remains only a fallback when the PNG cannot be decoded,
'   written, or applied to the shape
'
' INPUTS
'   None
'
' RETURNS
'   Decoded PNG bytes
'
' BEHAVIOR
'   Loads the embedded Base64 payload
'   Normalizes the Base64 text by removing whitespace and copy/paste artifacts
'   Validates the Base64 character set and padding position before decoding
'   Decodes the payload through a late-bound MSXML Base64 node
'
' ERROR POLICY
'   Raises a compact descriptive runtime error when the embedded payload is
'   blank, structurally invalid, or cannot be decoded
'
'   Does not raise the full MSXML error text because that can contain the entire
'   Base64 payload and flood the Immediate Window or result sheet
'
' DEPENDENCIES
'   M_GridIcon_EmbeddedIconBase64
'   M_GridIcon_NormalizeBase64
'   M_GridIcon_Base64LooksValid
'   MSXML2.DOMDocument.6.0
'
' NOTES
'   This routine keeps the embedded PNG as the primary grid-icon source
'
'   If this routine fails, M_GridIcon_Create should fall back to the glyph path
'
' UPDATED
'   2026-05-08
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_GridIcon_EmbeddedIconBytes"

    Dim XmlDoc                 As Object       'MSXML document
    Dim XmlNode                As Object       'MSXML Base64 node
    Dim EncodedText            As String       'Normalized Base64 encoded PNG
    Dim PayloadLength          As Long         'Normalized Base64 payload length
    Dim PayloadRemainder       As Long         'Payload length remainder modulo 4
    Dim HandlerStep            As String       'Current handler step for diagnostics
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' LOAD AND NORMALIZE PAYLOAD
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Load embedded Base64 payload"
    'Load and normalize the embedded Base64 payload
        EncodedText = M_GridIcon_NormalizeBase64(M_GridIcon_EmbeddedIconBase64())
    'Capture normalized payload length
        PayloadLength = VBA.Len(EncodedText)
    'Capture normalized payload modulo 4
        PayloadRemainder = PayloadLength Mod 4

'------------------------------------------------------------------------------
' VALIDATE BASE64 PAYLOAD
'------------------------------------------------------------------------------
    'Reject an invalid embedded payload before MSXML decoding
        If Not M_GridIcon_Base64LooksValid(EncodedText) Then
            Err.Raise vbObjectError + 514, PROC_NAME, _
                "Embedded icon payload is not valid Base64."
        End If

'------------------------------------------------------------------------------
' CREATE DECODER
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Create MSXML decoder"
    'Create the MSXML document
        Set XmlDoc = VBA.CreateObject("MSXML2.DOMDocument.6.0")
    'Create a temporary Base64 node
        Set XmlNode = XmlDoc.createElement("Base64Data")

'------------------------------------------------------------------------------
' DECODE BASE64
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Decode Base64 payload"
    'Configure the node as Base64 binary data
        XmlNode.DataType = "bin.base64"
    'Assign the normalized Base64 text
        XmlNode.Text = EncodedText
    'Return the decoded byte array
        M_GridIcon_EmbeddedIconBytes = XmlNode.NodeTypedValue

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set XmlNode = Nothing
    'Release object references
        Set XmlDoc = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Release object references
        Set XmlNode = Nothing
    'Release object references
        Set XmlDoc = Nothing
    'Raise a compact descriptive error to the caller
        Err.Raise ErrorNumber, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Embedded grid-icon PNG decode failed. " & _
            "Length=" & VBA.CStr(PayloadLength) & _
            "; Length Mod 4=" & VBA.CStr(PayloadRemainder) & _
            "; Native error=" & VBA.CStr(ErrorNumber) & _
            "; " & ErrorDescription

End Function

Private Function M_GridIcon_NormalizeBase64(ByVal EncodedText As String) As String

'
'------------------------------------------------------------------------------
'                         NORMALIZE BASE64 TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Removes whitespace and common copy/paste artifacts from Base64 text
'
' WHY THIS EXISTS
'   Embedded Base64 stored in VBA can be damaged by line wrapping, copied hidden
'   characters, non-breaking spaces, or accidental formatting artifacts
'
' INPUTS
'   EncodedText
'     Raw Base64 text
'
' RETURNS
'   Normalized Base64 text
'
' BEHAVIOR
'   Removes carriage returns, line feeds, tabs, normal spaces, non-breaking
'   spaces, zero-width spaces, and byte-order-mark characters
'
' ERROR POLICY
'   Does not raise custom errors
'
' DEPENDENCIES
'   VBA.Replace
'   VBA.ChrW$
'
' NOTES
'   This routine does not validate Base64 correctness
'
'   Validation belongs in M_GridIcon_Base64LooksValid
'
' UPDATED
'   2026-05-08
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResultText             As String       'Normalized Base64 text

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Start from the supplied text
        ResultText = EncodedText

'------------------------------------------------------------------------------
' REMOVE STANDARD WHITESPACE
'------------------------------------------------------------------------------
    'Remove carriage returns
        ResultText = VBA.Replace(ResultText, vbCr, vbNullString)
    'Remove line feeds
        ResultText = VBA.Replace(ResultText, vbLf, vbNullString)
    'Remove tabs
        ResultText = VBA.Replace(ResultText, vbTab, vbNullString)
    'Remove normal spaces
        ResultText = VBA.Replace(ResultText, " ", vbNullString)

'------------------------------------------------------------------------------
' REMOVE COMMON COPY / PASTE ARTIFACTS
'------------------------------------------------------------------------------
    'Remove non-breaking spaces
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(160), vbNullString)
    'Remove zero-width spaces
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(8203), vbNullString)
    'Remove byte-order marks
        ResultText = VBA.Replace(ResultText, VBA.ChrW$(65279), vbNullString)

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized Base64 text
        M_GridIcon_NormalizeBase64 = ResultText

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
'   Shape.Fill.UserPicture requires a real file path
'
' INPUTS
'   None
'
' RETURNS
'   Base64 text representing a 64 x 64 transparent PNG icon
'
' BEHAVIOR
'   Builds the Base64 payload by concatenating short string chunks
'   Returns the concatenated Base64 payload as a String
'
' ERROR POLICY
'   Does not raise custom errors
'
' DEPENDENCIES
'   None
'
' NOTES
'   Generated from DP_GridIcon_64.png
'   Base64 length: 6228
'   Base64 length Mod 4: 0
'
' UPDATED
'   2026-05-08
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EncodedText            As String       'Base64 payload

'------------------------------------------------------------------------------
' BUILD PAYLOAD
'------------------------------------------------------------------------------
    'Append the embedded Base64 chunks
        EncodedText = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAASBklEQVR42uVbfYwd1XX/nXNn5r19u+u1jXHBBttFoSStEtSK"
        EncodedText = EncodedText & "RoikUIgUVY3U0lYmfzVRm0aVQC1/RFHbKNLalRopqhoV8gFKE0ILIek6QAOtBUEE07pQWoiBLCFAUooB8+EPvPv2fc3Mvad/"
        EncodedText = EncodedText & "zJ07987MW5uoSKQZaeX1ezP349xzfud3fmcW+Bm/6M3eLyIgIvnanff//llnbn7P0eMrz9196+tf3bfvKi0iRETyVi64nGNx"
        EncodedText = EncodedText & "8YHo/IvUH2zbuvn8F48c/f5Hf+fyW+x3ACBvhaFoaWlJAcC+/QdvPNafiDYiqyMt37rn4B2Liw9ES0tLCiJU3v9//SMitCjC"
        EncodedText = EncodedText & "i4uL0e33HLyrP9YiInJsdSL79h/8PAAUa3Rr+MmvxcVFvmxxMcLiIvve8tdfXvr5/3ryGRGRfDAYpiKS/uC/X5bP3njbReVz"
        EncodedText = EncodedText & "b9Xpl2N/8R/uvOS5w6+JaJ3aNeQPH3o6/+Rnv7Kt9gRfdtli9GbXRNhdnLTvdiLLiYjwl7+5/+If/OhFMx6PzcrqmhmPJ/rH"
        EncodedText = EncodedText & "h1/VNy39yxUiwsvLy4mIqOJnyfu3/F1qn7V9N+X7ZUlEhL9+132/9+JrJ/R4NM5XVtdkMknN40//WF//1dsvFBEWkVgKT6yu"
        EncodedText = EncodedText & "Yk+0PgaIEGwM/9Ynb7lgLcWHc03vA/EOEYlAZLJMdz/zh7+649JfPg/9YYq5mQRPPHcEf3L9Ay+BoyFTMaaIQAAQuQ8A8qYj"
        EncodedText = EncodedText & "ghEAYsAggMndI/Zz2OfL+2FAUGyUmcx+4dortl+w6yyMxhnmZzv47qM/kr/4yiMvdJI4hRgGU6aIj8Dk/zYTp9+897o/esZu"
        EncodedText = EncodedText & "koAKp6L65peXl5Nr/u7QZ06Mu1cnc2fMJCqBFoERQDEhHY2hRUFEYIqVYpILdG/7ObOzsxBj3EYp9CL7SfEMEUNAxa0i7jzE"
        EncodedText = EncodedText & "fg+pbd7eIwCyyQi5KT4zIoAYjI0izO/Y1ZntQGsDxQAg79LZ5AP9wfE//7Wr//66v9y949OXX065b4TIbX7PHvqbhw53P/6F"
        EncodedText = EncodedText & "B/6pu+ncD8Iw1sY6z82QREDF/ILJJEVuDIsYiBFoYwAQxplIujIQd8j2frEbIw+ag1uIQBBQ4S8Q95xnRPswEUGLIGJDRETG"
        EncodedText = EncodedText & "SOFpIhADrK6NjclzlKYSQBLF0pvb2gXMn336H1+4cP/+/Vf+5iN7MuwVACQMAIt79hDt3Wu+dfN3buxs2vHBScbpytpIsjyL"
        EncodedText = EncodedText & "CFAEMJFwpBSngzdYjwcQMIwxABhmMsDo+EsUxxErAjMRM8H9EIGJwcxgxcTsfoe9H1wEgTCRMFP1ObP9nomJhOM44tEbr1E6"
        EncodedText = EncodedText & "GhRhZARCxRrStaMcxREzCVPxvMq0jlYGE8mNSuOFc37jr+569fO0d6/ZvfsqBgDevXtJ7d271/z61Tdfis7mj4xTyYfjcRLZ"
        EncodedText = EncodedText & "lfonIRCsHn0BJAbaCIwYGDHQ2qB/4mUYk5f+DhDssyUG1LKzVP8V5/5U+QjZSBIfNgg6T4s1EGDK0xdA5xn6J45ALHaU4xII"
        EncodedText = EncodedText & "RET9wShOc+SIFz7+/j+++b379u3Tu3cvKd5nbx5puVrNbJJxmiFigpBzSut/DK0zpHmONDdgAnJd7EKIkWuNdDwEMYdY54Fe"
        EncodedText = EncodedText & "yVGobo3SGOLNWDNWAYqEPB0h1xpSQCdybUAE5EaQTYYweQq2wEue3ZmZRpMcPLMgucmvAYDXX3+KGPuu0tddt78DFV8ySTUR"
        EncodedText = EncodedText & "iMtAJodHBIhBHMUgFeGOB5/CcDjCpvkOjM6w7/4nkGpBksxAjHEnCqBAeru56sTJbVjE8wPPVcqTrUBRYIxGp9PDKBN8497v"
        EncodedText = EncodedText & "AaKxab6L1dU13PnAMqK4A1axndOagAhEbAGVeJwKiUouffTRR+MHH9ybRwBwaHVlq9G8VQtAJFS5Y+WcxcXYfs4u3P3wM3j1"
        EncodedText = EncodedText & "U1/HJe/ZhceffRn/+uRL2HbueSBWEKNtKhM3hoj9TeAyRGHU0BNcRiCyC5YgRCACYoWzt+/Ezfcs4/kjJ/Hu87fh4e8fxmPP"
        EncodedText = EncodedText & "vobtu94JIQUgh+dQLvlBDGkhaI3Ne257bAHAsQgA0oGJtUAV6clDbbeGwn21ztHbeDZ2MvDk4Zfw0A8fwUyng3N3vgPdTecg"
        EncodedText = EncodedText & "z3IvdaFIHUE4VChPVEG88wIRL2NIGcRVPBPBiMHC1p2IFOHg0y/j/ideRq/bxc7zLkAyfxa0zmzadGYunjdSegFyrfHKK/2K"
        EncodedText = EncodedText & "B7DSQuDqwXLhdmFE5HK21hqdhbPxjk1nw+gUrBLkhpFlKYjZHSiDIFy6uQ9k5SalBhQe2AacgZx3kPWMPNeYPWMHLthyDsRk"
        EncodedText = EncodedText & "IE6QGUae2wOQCjQJ4rZTgUt1KlGN8sKur1iEx8zKIq9AYo2cCEQJJDUQ0WBiVBm8IDQk5P7nxqfCs+pM1VaZDrSodMFwvRBT"
        EncodedText = EncodedText & "rCnPNTIAoASiBYTckkmBFHm7wJDS+j7h8mwfBUxNvEkgpdktBlXkhIndZGXIVDFNDeLj6K0HdD4zFg8oqxRC3rFVX5eGdazR"
        EncodedText = EncodedText & "GLBdpxgTbtCBLNm0bNkl1QwwGRefMRPIA0DyFu9nJaKC3Uh5PFLl3CK+yW2yJJogDwcggLL3U3VKbj6QJariuEQ1OYPgsUx/"
        EncodedText = EncodedText & "YYp9AhDgSzkvK4KhWgjEc4qGRyc01sMA+4nIEQvnnlLGIqpTFQqYDpUsRsSRYReH1KTDzm+s2/uUuRy6DK0SZEXEM7L4SOnR"
        EncodedText = EncodedText & "h/KZimewiqAmGeZ9A1C2Nvrdi3dM3v2ud/a0zl3+dOFDFSRBPO8sF2LTXrFxcWhPHoJX35eRIMHmqzEpDHuq0qWxSF4eRsiW"
        EncodedText = EncodedText & "Ku8TrxapQrS4X0URln+o6KUXBvRYaYCNc2l6+YVnZldcsutnQgfcgFfkvrVuVQ12JgmtjTVpI9BaVycVxNOU/78drtNck4hA"
        EncodedText = EncodedText & "KYW1cQ70vRBYQ6FHKC5yFNHbfcc/+cVMxT7DNNgPqigf7X1i4YdbkGMBm4ra9VgjUhvTs6tMH9Pd26LzhmOiWuc6Y7bQgBoP"
        EncodedText = EncodedText & "qG1+kmYYjychK7P77M10EEXKDdofFFUaAUGCj5RCr9d1Y2ZZjuF4gsa+RDDT7SBJYgeYw9EYWZZb0aS6n4kwOzvjEqTWBoPh"
        EncodedText = EncodedText & "uBmtIuh0EnQ7SbA/aSdCzZPr9weYpJmluGI3TzDGwBjBpoU5gIA0zbDSH3hoTi6DiwBxHCFJYgDAYDjGYDgGc5jujDHIco0t"
        EncodedText = EncodedText & "mxeK8jbXWO0PbGUnAS4ZbRBFCjPdDgBgOJqgPxhZbhJmqEmao5NElS5hZ1xro8ItMeAkGbGEpzhcRpXhy3qbHP2sMftGVDBT"
        EncodedText = EncodedText & "JZbaJTOz5R2CqrdBUFxuvzIqmFtpNBEFHuATspoei7mGAWh628j3DqlVWaGoIaeEzDD7Sx2mPXHUqwg9M8p0iA8FRDccBXBe"
        EncodedText = EncodedText & "/43X65E1H6jYl0y11np9KWmZjZr035NUKuNXAkdAKU8xR1to+37CoRpxOsmWpvfOiFpbDtSYWqakV2rso76stvMs5TYiqmvO"
        EncodedText = EncodedText & "rdsiAkouzOs6q9TjV1p/dV4hTRMJJEDgQspuiOO2CjXBsMZIEHZuHiMNDGh6ZVho+TMaIyERavcVQieJYYyx4ORZj9ihuogg"
        EncodedText = EncodedText & "UhHiKII2uqjWvL0pFSGKlCt1O0mMLM/BzJ7wQWA26HQSB4RKMTpJXKRWQlAERZEKx4wjjBV7pa7VIw2QJHE1V5lFBPUsMN/w"
        EncodedText = EncodedText & "FRHB/FwPs7MzjdN2BYm1ilKEMzZvcCcW3mu1A2vk2V4XMzOd1thUyitfibB50wbbe6DG/MxlC03QdfyhGbLMFFaLduy5OhM0"
        EncodedText = EncodedText & "xrRTx5qA0IYB5cTM7TEduLV4SnHtTq+r5qVMbh9T2tLgFLSQ6RgQVR/yVGBrw27TwhwprEyDUtg/ZTUVeiTwolB3CIduY65t"
        EncodedText = EncodedText & "x+PfK57B1vq1EGigLRVNhzTNA7AqBZFuJ64mJWCSZshzHWQoESCKFDpJBGMEzIwT/RQPLh8LExYBeW7w/l86E2dt7Nh7C4aZ"
        EncodedText = EncodedText & "5TqQyCEAW3zwT3g8mTSMIgJ0kthS9tCwc81iqAmCK6sDjCepJ38VnxtjsGGuh4X5WQiANMtx/I2VqvkhlfbGxDhj0wawUiAi"
        EncodedText = EncodedText & "fO7bz+Gm+57HQi8qOku2El0dZrjy4u244ZqLAAi0Njhxsg9tjBM1SpIhALZsXkAniUEA+sMRTq6uFeEScH5BksTYsnkBtRKl"
        EncodedText = EncodedText & "XgvMT6+fuRQ5q1yrAWjjUWFjySpzrdFBro2u7Own+xNsno2x0IugTeUB3ZjxxtrEYUmWWYBlqhEhFD1JY5y3GSuMluJoPVRK"
        EncodedText = EncodedText & "jCiXZnKNNZsHufSAaeSmkYNdN0dO+1UrHwojRchyDW1b67kxMFKEWxyxJ7PB0/yotava6CGiWpusQ95IEebsofOpuF8borbS"
        EncodedText = EncodedText & "SzoV+pIjIdTKEkOK2qJxtk1W39qUXbQDZmCABgbIm32jLhCrW6avqClxjeOLBFloGqL74vup19KsQqWF9UdtVgne7SA6BeeH"
        EncodedText = EncodedText & "1+SkoiVWc9f6giMmRIpgpCAzzIRIGadKiUgznzpMCTfQ1IWn+RPWF0QgaHUTrXWwYSI4eTp4ccGYoB8wvUgjHFudQGuD3OZ8"
        EncodedText = EncodedText & "Vgr94QTAhuCcxXV2JFDIpcVLdEDkvHKXVesh4nQEkQ1zPYyi5gCwklh5wnEcYeOGOSuJhfVapBhxHDkAveZD5+PnFrogeC9T"
        EncodedText = EncodedText & "EUFrwYfeu831JCOlsDA/i8w2PMVrnLJiJHHsPKU30w3A09WcVhIjanbdmmmQmikwSWJ0OvFpCZhzvZmmlC4VayRbmf3C9nl8"
        EncodedText = EncodedText & "6sO/ONXorviCoNfrtguoHhMVAFHEWNgw21q9B+uUZnBEVRqkVh4g0hIi05TflnvqMZtr3cr5CyywtUfNGG2VONWelSlh16Y6"
        EncodedText = EncodedText & "iEizFsAUWbyEQ5I6ilKDXrYtgpmC12MUc6HzNbh8rZlJRX3SOiaFY9absb6zsLeuQFxq6AH1jEqEwXCM4WgUYoA1xPxcz8nN"
        EncodedText = EncodedText & "IoKTK2vIc416U0kphY0bZl0JPR4XCi7V+oIiwOxMF71e13neSn8NWZqFdb4UNcXChlkoLl7qSNMMq/1BYxvGFFL7/FwvNKSc"
        EncodedText = EncodedText & "JggOR2OkWR6qvTYLjEYTzHQSu4BC6/etXRotTTPM9rqueBmOJ5ikmSudKxS3gNbrFv2DPMdwOA4KrjKcTJZjJutgptsBETCa"
        EncodedText = EncodedText & "pHZMDgJERDAYjTFrx2y7ovVoINk3HIvTk8Dd/DIZtq1WhEH4aku9ni/FDK4tqDhMqqUw8trtCIhUXQx28r347yjBvblSx8e1"
        EncodedText = EncodedText & "UzLB2tuNJQUWaUfB8DuawuKmy6XrAay0HU4LNRafe8r0ef1ymNcVhWWdVZ2iepDT5OnSqmbTNJlyisxdb8a0P98m7FQGqEtP"
        EncodedText = EncodedText & "9VfcCC2y0/oGkilGlcYUza7OKewbZlyR9jR4Gk1tPq0JPfd2oeDevoB7iSqUsL3PxW90SMMI5SZM7RU2n4f4b402lR8JQtDR"
        EncodedText = EncodedText & "Z2mXxaVugCyNqU05TJJC9iIue3dcCA+lJGUfiaIIcRyBrApLDjgLmhxFyk3csVIal708r6lRKjyFEFNQaNcM8qTxKFKIS1kc"
        EncodedText = EncodedText & "hfStFLvx/LFLWbzWGaJOEpHLAju2bsiMER1ggQjm52Yxa1vbdd9STI6OKkXYEsji3qtWtsHq8/Zup9MaNszsxmQmbNm8YIss"
        EncodedText = EncodedText & "tGaSUhaf6XaQxPHUMSXs2EJEJslcIT9F9k/NVr52x32H09xsTCJljDGqRFSlONDjUNPVyBNKlaKakkGt8am4ev9wWk+xfCRS"
        EncodedText = EncodedText & "KvySmqKMk+4a/DrEcVbKpLnhteHohb3XXrsqIsQHDhxQAKTfH93z2vFV64GmwbNLqt+YvJYGw/tkKrTUx5Optbs395RmZzCm"
        EncodedText = EncodedText & "tI9pm27yyvEVWhuM7wWAAwcOKD5w4IABADMef+k/Di0PDcBKRWba4t8kyL4trqJ9p0xuwP/5vacGufAN1gDFSZd/EHnDN/75"
        EncodedText = EncodedText & "Y//++HMiIrmI5Hmey/+Hn3I/Bw89K1+89dsf8/fsrvKD62+9+xPfeegJ6Y9SEREjItlP+Y/pD1O576En5Eu33f2J+uapboSr"
        EncodedText = EncodedText & "rrpK/+0td37gzI2b9m7besb7ztuxPSgmqldmZXooBLIYVdzVgqP/brDUKKr7e0O0vpw29a0F/92kMv/3ByM8/+IRvHr0xMFj"
        EncodedText = EncodedText & "J04u/ulHfvu75R6njunf8Lmbbv+V3szMRUTYaYzEdeWmAJew4OEAeELCYfxvTLMBylw0SKu77H3F33fZMUw1ov3Ofg0GFWBn"
        EncodedText = EncodedText & "AGY2eW7+J9fZI9d+9MpDALB7aUnt8zY/9VpaWlKNPz39Kb5EhBoxb6//BTrnfQKAxO/pAAAAAElFTkSuQmCC"

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
'   This routine provides a broader cleanup pass than the normal single-icon
'   removal path
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Suppresses cleanup errors
'   Deletes the tracked grid icon shape when available
'   Clears the tracked grid icon shape reference
'   Scans all open workbooks
'   Deletes shapes named DP_GRID_ICON_NAME from each open workbook
'   Restores normal error handling before exit
'
' ERROR POLICY
'   Best-effort cleanup
'   Suppresses workbook, worksheet, protection, and shape-deletion errors
'   Does not raise custom errors
'
' DEPENDENCIES
'   gDP_GridIconShape
'   DP_GRID_ICON_NAME
'   M_GridIcon_DeleteNamedShapeAcrossWorkbook
'   Application.Workbooks
'   Excel Workbook / Worksheet / Shape object model
'
' NOTES
'   This routine is intentionally heavier than M_GridIcon_Remove
'   Do not call this routine from high-frequency selection-change paths
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurWorkbook            As Workbook     'Workbook being scanned

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
    'Clear the cached last target
        M_GridIcon_ClearLastTarget

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
'                    DELETE NAMED SHAPE ACROSS WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Deletes same-named DatePicker in-grid icon shapes from every worksheet in a
'   target workbook
'
' WHY THIS EXISTS
'   DatePicker in-grid icons are transient worksheet shapes
'
'   During workbook activation, add-in reload, VBA reset, or interrupted UI
'   flows, a stale icon may remain on a worksheet even when the tracked shape
'   reference has already been lost
'
'   This helper provides the workbook-level cleanup pass used by the broader
'   grid-icon purge routine
'
' INPUTS
'   TargetWorkbook
'     Workbook to scan
'
'   TargetShapeName
'     Name of the worksheet shape to delete
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Suppresses cleanup errors
'   Exits when no workbook is supplied
'   Exits when the target shape name is blank
'   Loops through every worksheet in the workbook
'   Attempts to delete the same-named shape from each worksheet
'   Restores normal error handling before exit
'
' ERROR POLICY
'   Best-effort cleanup
'   Suppresses workbook, worksheet, protection, and shape-deletion errors
'   Does not raise custom errors
'
' DEPENDENCIES
'   Excel Workbook / Worksheet / Shape object model
'
' NOTES
'   This routine intentionally deletes by shape name only
'   This routine intentionally does not validate whether the named shape belongs
'   to the DatePicker beyond the supplied shape name
'   This routine intentionally does not trim TargetShapeName
'   Keep this helper aligned with M_GridIcon_PurgeAll
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurWorksheet           As Worksheet    'Worksheet being scanned

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit when no workbook is supplied
        If TargetWorkbook Is Nothing Then GoTo ExitProcedure
    'Exit when the target shape name is blank
        If LenB(TargetShapeName) = 0 Then GoTo ExitProcedure


'------------------------------------------------------------------------------
' DELETE NAMED SHAPES
'------------------------------------------------------------------------------
    'Loop through all worksheets in the target workbook
        For Each CurWorksheet In TargetWorkbook.Worksheets
            'Delete the named shape when found
                CurWorksheet.Shapes(TargetShapeName).Delete
        Next CurWorksheet

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release local object references
        Set CurWorksheet = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
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
'                         FORM BRIDGE GET LOADED FORM
'------------------------------------------------------------------------------
' PURPOSE
'   Returns an already-loaded UserForm instance by class name or runtime name
'
' WHY THIS EXISTS
'   Direct references to a UserForm default instance can instantiate it
'
'   This helper scans only the currently loaded VBA.UserForms collection so the
'   lookup does not create a new default UserForm instance as a side effect
'
' INPUTS
'   UserFormName
'     UserForm class name or runtime name to locate
'
' RETURNS
'   Matching loaded UserForm object
'
'   Nothing when no matching form is currently loaded
'
' BEHAVIOR
'   Trims the requested form name
'   Returns Nothing when the requested form name is blank
'   Scans only currently loaded UserForms
'   Compares both TypeName and the runtime Name property case-insensitively
'   Prefers a visible matching instance
'   Falls back to the first hidden matching instance when no visible match exists
'
' ERROR POLICY
'   Safe-default lookup
'   Returns Nothing on error
'   Suppresses metadata-read errors while inspecting each loaded form
'
' DEPENDENCIES
'   VBA.UserForms
'   VBA.TypeName
'
' NOTES
'   TypeName is preferred because the runtime Name property can theoretically be
'   changed
'
'   The runtime Name property is still checked as a practical fallback for
'   callers that identify a form by its current instance name
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LoadedForm             As Object       'Loaded UserForm instance
    Dim HiddenMatch            As Object       'First hidden matching UserForm
    Dim NormalizedFormName     As String       'Trimmed requested form name
    Dim FormTypeName           As String       'Loaded form class name
    Dim FormRuntimeName        As String       'Loaded form runtime name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set the safe default
        Set M_FormBridge_GetLoadedForm = Nothing
    'Enable safe-default lookup handling
        On Error GoTo FailSafe
    'Normalize the requested form name
        NormalizedFormName = VBA.Trim$(UserFormName)

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no form name was supplied
        If VBA.Len(NormalizedFormName) = 0 Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' SCAN LOADED FORMS
'------------------------------------------------------------------------------
    'Loop through loaded UserForms without instantiating default instances
        For Each LoadedForm In VBA.UserForms
            'Reset current metadata
                FormTypeName = vbNullString
            'Reset current metadata
                FormRuntimeName = vbNullString
            'Suppress metadata-read errors for unusual transient form states
                On Error Resume Next
            'Capture the loaded form class name
                FormTypeName = VBA.TypeName(LoadedForm)
            'Capture the loaded form runtime name
                FormRuntimeName = VBA.CStr(LoadedForm.Name)
            'Clear any suppressed metadata-read error
                Err.Clear
            'Restore safe-default lookup handling
                On Error GoTo FailSafe
            'Check whether this loaded form matches the requested name
                If VBA.StrComp(FormTypeName, NormalizedFormName, vbTextCompare) = 0 Or _
                   VBA.StrComp(FormRuntimeName, NormalizedFormName, vbTextCompare) = 0 Then
                    'Return immediately when a visible match is found
                        If VBA.CBool(LoadedForm.Visible) Then
                            Set M_FormBridge_GetLoadedForm = LoadedForm
                            GoTo ExitProcedure
                        End If
                    'Keep the first hidden match as a fallback
                        If HiddenMatch Is Nothing Then
                            Set HiddenMatch = LoadedForm
                        End If
                End If
        Next LoadedForm

'------------------------------------------------------------------------------
' RETURN HIDDEN FALLBACK
'------------------------------------------------------------------------------
    'Return the first hidden match when no visible match was found
        If Not HiddenMatch Is Nothing Then
            Set M_FormBridge_GetLoadedForm = HiddenMatch
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Restore normal error handling
        On Error GoTo 0
    'Exit before the fail-safe handler
        Exit Function

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Return the safe default
        Set M_FormBridge_GetLoadedForm = Nothing
    'Restore normal error handling
        On Error GoTo 0

End Function

Private Function M_FormBridge_TryReuseLoadedPickerFromActiveCell( _
    ByVal OffsetXPx As Long, _
    ByVal OffsetYPx As Long, _
    ByVal CenterOnMouse As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'               FORM BRIDGE TRY REUSE LOADED PICKER FROM ACTIVE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to reuse an already-visible DatePicker form instead of unloading and
'   rebuilding it
'
' WHY THIS EXISTS
'   DP_Show currently unloads and rebuilds the UserForm every time it is called.
'   That is robust, but slower than necessary when the picker is already visible.
'
'   Reusing the visible form avoids rebuilding runtime labels, hooks, panels,
'   fonts, settings controls, and day-grid infrastructure.
'
' INPUTS
'   OffsetXPx
'     Horizontal mouse-position offset in pixels
'
'   OffsetYPx
'     Vertical mouse-position offset in pixels
'
'   CenterOnMouse
'     True to center the form on the mouse position before applying offsets
'
' RETURNS
'   True when an already-visible picker was refreshed and reused
'
'   False when no reusable visible picker exists or when reuse failed
'
' BEHAVIOR
'   Resolves an already-loaded DatePicker form without creating a default
'   instance, exits when no visible picker exists, refreshes the visible picker
'   from the current ActiveCell when available, reapplies clock mode, repositions
'   the form near the mouse, and returns True
'
' ERROR POLICY
'   Safe fallback helper
'
'   Does not raise outward. If reuse fails, returns False so DP_Show can continue
'   through the existing unload / load path
'
' DEPENDENCIES
'   DP_FORM_NAME
'   M_FormBridge_GetLoadedForm
'   M_FormBridge_RefreshFromCell
'   M_Timer_ApplyClockMode
'   M_Window_MoveFormToMouse
'
' NOTES
'   This helper deliberately avoids direct UF_DatePicker references
'
'   The existing DP_Show cold-load path remains the fallback
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_FormBridge_TryReuseLoadedPickerFromActiveCell"

    Dim LoadedForm          As Object           'Already-loaded DatePicker form instance
    Dim ActiveCellRef       As Excel.Range      'Current ActiveCell reference
    Dim HandlerStep         As String           'Current handler step for diagnostics
    Dim ErrorNumber         As Long             'Captured runtime error number
    Dim ErrorDescription    As String           'Captured runtime error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe fallback result
        M_FormBridge_TryReuseLoadedPickerFromActiveCell = False
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve loaded DatePicker form"
    'Resolve the already-loaded DatePicker form without creating a default instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Exit when no DatePicker form is currently loaded
        If LoadedForm Is Nothing Then GoTo CleanExit
    'Exit when the DatePicker form is loaded but not visible
        If Not VBA.CBool(LoadedForm.Visible) Then GoTo CleanExit

'------------------------------------------------------------------------------
' RESOLVE ACTIVE CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve ActiveCell"
    'Suppress ActiveCell resolution errors
        On Error Resume Next
    'Resolve the current ActiveCell
        Set ActiveCellRef = Excel.Application.ActiveCell
    'Clear any suppressed ActiveCell error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REFRESH VISIBLE FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh visible picker from ActiveCell"
    'Refresh the visible picker from the current ActiveCell when available
        If Not ActiveCellRef Is Nothing Then
            M_FormBridge_RefreshFromCell ActiveCellRef
        End If

'------------------------------------------------------------------------------
' REAPPLY CLOCK MODE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply clock mode"
    'Reapply clock mode in case timer state was stale or reset
        M_Timer_ApplyClockMode

'------------------------------------------------------------------------------
' REPOSITION FORM
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Move visible picker to mouse"
    'Suppress best-effort positioning errors
        On Error Resume Next
    'Move the visible picker close to the current mouse position
        M_Window_MoveFormToMouse _
            LoadedForm, _
            OffsetXPx, _
            OffsetYPx, _
            CenterOnMouse
    'Clear any suppressed positioning error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RETURN SUCCESS
'------------------------------------------------------------------------------
    'Return success because the existing visible picker was reused
        M_FormBridge_TryReuseLoadedPickerFromActiveCell = True

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors
        On Error Resume Next
    'Release the ActiveCell reference
        Set ActiveCellRef = Nothing
    'Release the loaded form reference
        Set LoadedForm = Nothing
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0
    'Exit the function
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error number
        ErrorNumber = Err.Number
    'Capture the original error description
        ErrorDescription = Err.Description
    'Return False so DP_Show can use the cold-load fallback path
        M_FormBridge_TryReuseLoadedPickerFromActiveCell = False
    'Write diagnostics without interrupting DP_Show
        Debug.Print PROC_NAME & _
            " | Step=" & HandlerStep & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Continue through cleanup
        Resume CleanExit

End Function
Public Sub M_FormBridge_RefreshFromCell(ByVal TargetCell As Excel.Range)

'
'------------------------------------------------------------------------------
'                           FORM BRIDGE REFRESH FROM CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes the visible DatePicker form from an explicit worksheet cell
'
' WHY THIS EXISTS
'   The manager already resolves the authoritative target cell
'
'   The form bridge must therefore refresh from that same cell instead of
'   independently reading Excel.Application.ActiveCell
'
' INPUTS
'   TargetCell
'     Explicit worksheet cell to use as the DatePicker context
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits when no target cell is supplied
'   Normalizes the target to one physical cell
'   Normalizes merged cells to the top-left cell of the merge area
'   Captures workbook, worksheet, and target-address diagnostics
'   Resolves a selected date only from non-error date-like values
'   Locates the already-loaded DatePicker form
'   Exits when the DatePicker form is not loaded
'   Exits when the DatePicker form is loaded but not visible
'   Delegates the refresh to UF_DP_RefreshFromExternalSelection
'
' ERROR POLICY
'   Best-effort bridge routine
'   Does not raise errors outward
'   Writes diagnostics to the Immediate Window when refresh fails
'   Releases local object references before exit
'
' DEPENDENCIES
'   DP_FORM_NAME
'   M_FormBridge_GetLoadedForm
'   UF_DatePicker.UF_DP_RefreshFromExternalSelection
'   Excel Range / Worksheet / Workbook object model
'
' NOTES
'   This routine deliberately avoids Excel.Application.ActiveCell
'
'   This routine does not create or show the DatePicker form
'
'   If no visible DatePicker form is already loaded, the refresh is a safe no-op
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_FormBridge_RefreshFromCell" 'Current procedure name

    Dim EffectiveCell          As Excel.Range  'Normalized explicit target cell
    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim CellValue              As Variant      'Target cell value snapshot
    Dim SelectedDate           As Date         'Resolved selected date
    Dim HasSelectedDate        As Boolean      'True when target contains a date-like value
    Dim WorkbookName           As String       'Diagnostic workbook name
    Dim WorksheetName          As String       'Diagnostic worksheet name
    Dim TargetAddress          As String       'Diagnostic target address
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect caller from bridge failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit when no target cell was supplied
        If TargetCell Is Nothing Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' NORMALIZE TARGET
'------------------------------------------------------------------------------
    'Use the first physical cell defensively
        Set EffectiveCell = TargetCell.Cells(1, 1)
    'Normalize merged cells to the top-left cell of the merge area
        If EffectiveCell.MergeCells Then
            Set EffectiveCell = EffectiveCell.MergeArea.Cells(1, 1)
        End If

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Capture workbook name for diagnostics
        WorkbookName = VBA.CStr(EffectiveCell.Worksheet.Parent.Name)
    'Capture worksheet name for diagnostics
        WorksheetName = VBA.CStr(EffectiveCell.Worksheet.Name)
    'Capture target address for diagnostics
        TargetAddress = VBA.CStr(EffectiveCell.Address(External:=True))

'------------------------------------------------------------------------------
' RESOLVE TARGET DATE STATE
'------------------------------------------------------------------------------
    'Capture the target cell value once
        CellValue = EffectiveCell.Value
    'Default to no selected date
        HasSelectedDate = False
    'Resolve a date only from non-error date-like values
        If Not VBA.IsError(CellValue) Then
            If VBA.IsDate(CellValue) Then
                SelectedDate = VBA.DateValue(VBA.CDate(CellValue))
                HasSelectedDate = True
            End If
        End If

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Resolve the already-loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    'Exit when the picker is not loaded
        If LoadedForm Is Nothing Then GoTo ExitProcedure
    'Exit when the picker is loaded but not visible
        If Not VBA.CBool(LoadedForm.Visible) Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' REFRESH FORM
'------------------------------------------------------------------------------
    'Refresh the visible picker from the explicit target-cell date state
        LoadedForm.UF_DP_RefreshFromExternalSelection SelectedDate, HasSelectedDate

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Release object references
        Set EffectiveCell = Nothing
    'Release object references
        Set LoadedForm = Nothing
    'Restore normal error handling
        On Error GoTo 0
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
    'Suppress diagnostic and cleanup errors
        On Error Resume Next
    'Write bridge diagnostics without interrupting caller flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Worksheet=" & WorksheetName & _
            " | Target=" & TargetAddress & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Release object references
        Set EffectiveCell = Nothing
    'Release object references
        Set LoadedForm = Nothing
    'Clear the suppressed bridge error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub M_FormBridge_AfterSuccessfulSelection(ByVal SelectedDate As Date)

'
'------------------------------------------------------------------------------
'                     FORM BRIDGE AFTER SUCCESSFUL SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Notifies the already-loaded DatePicker UserForm after a date has been
'   successfully selected
'
' WHY THIS EXISTS
'   The companion module owns write-back, settings, and integration flow
'
'   The UserForm owns rendering and post-selection UI state
'
'   This bridge lets the companion module notify the loaded form without
'   directly referencing the default UserForm instance and accidentally
'   instantiating it
'
' INPUTS
'   SelectedDate
'     Date selected by the user
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the already-loaded DatePicker form instance
'   Calls UF_DP_AfterSuccessfulSelection when a loaded form is available
'   Performs no action when the DatePicker form is not loaded
'   Does not create, show, or activate the DatePicker form
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller when notification fails
'   Preserves the original error number and description
'
' DEPENDENCIES
'   DP_FORM_NAME
'   M_FormBridge_GetLoadedForm
'   UF_DatePicker.UF_DP_AfterSuccessfulSelection
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'   This routine intentionally relies on M_FormBridge_GetLoadedForm to avoid
'   default-instance creation
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_FormBridge_AfterSuccessfulSelection" 'Current procedure name

    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Retrieve the loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

'------------------------------------------------------------------------------
' NOTIFY FORM
'------------------------------------------------------------------------------
    'Notify the loaded form after successful selection
        If Not LoadedForm Is Nothing Then
            LoadedForm.UF_DP_AfterSuccessfulSelection SelectedDate
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release object references
        Set LoadedForm = Nothing
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
        Set LoadedForm = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub

Private Sub M_FormBridge_RefreshSettings()

'
'------------------------------------------------------------------------------
'                         FORM BRIDGE REFRESH SETTINGS
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes settings-dependent state on an already-loaded DatePicker UserForm
'
' WHY THIS EXISTS
'   The companion module owns write-back, settings, and integration flow
'
'   The UserForm owns rendering and settings-dependent UI state
'
'   This bridge lets the companion module notify the loaded form without
'   directly referencing the default UserForm instance and accidentally
'   instantiating it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the already-loaded DatePicker form through the shared form bridge
'   Performs no action when the DatePicker form is not loaded
'   Calls UF_DP_RefreshSettings when a loaded form is available
'   Does not create, show, activate, or otherwise instantiate the DatePicker form
'
' ERROR POLICY
'   Raises a descriptive runtime error to the caller when refresh fails
'   Preserves the original error number and description
'
' DEPENDENCIES
'   DP_FORM_NAME
'   M_FormBridge_GetLoadedForm
'   UF_DatePicker.UF_DP_RefreshSettings
'
' NOTES
'   Keep this routine aligned with the DatePicker companion-module API
'   This routine intentionally relies on M_FormBridge_GetLoadedForm so all form
'   bridge routines use the same loaded-form lookup policy
'
'   M_FormBridge_GetLoadedForm checks both TypeName and runtime Name, prefers a
'   visible matching instance, and falls back to a hidden matching instance
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_FormBridge_RefreshSettings"
    
    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim ErrorNumber            As Long         'Captured error number
    Dim ErrorDescription       As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE LOADED FORM
'------------------------------------------------------------------------------
    'Resolve the already-loaded DatePicker form instance
        Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)

'------------------------------------------------------------------------------
' REFRESH SETTINGS
'------------------------------------------------------------------------------
    'Refresh the DatePicker form when it is loaded
        If Not LoadedForm Is Nothing Then
            LoadedForm.UF_DP_RefreshSettings
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release object references
        Set LoadedForm = Nothing
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
        Set LoadedForm = Nothing
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Sub
Private Sub M_FormBridge_UnloadLoadedPicker()

'
'------------------------------------------------------------------------------
'                       FORM BRIDGE UNLOAD LOADED PICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Unloads any already-loaded DatePicker UserForm instance before DP_Show loads
'   a fresh instance
'
' WHY THIS EXISTS
'   DP_Show relies on UserForm_Initialize to rebuild the runtime UI and consume
'   the pending initial-date bridge state
'
'   If a previous DatePicker instance is already loaded, it should be removed
'   before the new instance is loaded
'
'   The helper must avoid direct references to the DatePicker default instance
'   because such references can instantiate the form accidentally
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Scans the currently loaded VBA.UserForms collection
'   Collects loaded DatePicker instances by class name or runtime name
'   Separates the scan phase from the unload phase
'   Stops any active or stale DatePicker timer after the loaded-form scan
'   Exits after timer cleanup when no DatePicker instance is loaded
'   Hides each collected picker before unload to reduce visual flicker
'   Unloads collected picker instances in reverse collection order
'   Writes non-fatal diagnostics to the Immediate Window
'   Does not create, show, activate, or instantiate the DatePicker form
'
' ERROR POLICY
'   Best-effort cleanup
'   Does not raise errors outward
'   Suppresses metadata-read, timer-stop, hide, and unload failures where needed
'   Continues through cleanup after unexpected helper failures
'
' DEPENDENCIES
'   DP_FORM_NAME
'   M_Timer_Stop
'   VBA.UserForms
'   VBA.TypeName
'
' NOTES
'   This routine deliberately avoids direct references to UF_DatePicker
'
'   The scan and unload phases are separated because unloading a form changes
'   the VBA.UserForms collection
'
'   DP_Show must not fail just because an old transient form instance or stale
'   timer cannot be cleaned up fully
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_FormBridge_UnloadLoadedPicker"

    Dim CurForm                As Object       'Current loaded UserForm instance
    Dim LoadedForm             As Object       'DatePicker form selected for unload
    Dim FormsToUnload          As Collection   'Matching DatePicker forms to unload
    Dim Index                  As Long         'Collection reverse-loop index
    Dim FormTypeName           As String       'Loaded form class name
    Dim FormRuntimeName        As String       'Loaded form runtime name
    Dim FormMatchesPicker      As Boolean      'True when loaded form is a DatePicker instance
    Dim StepName               As String       'Current diagnostic step
    Dim StepErrNumber          As Long         'Captured non-fatal cleanup error number
    Dim StepErrDescription     As String       'Captured non-fatal cleanup error description

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
            'Reset form metadata
                FormRuntimeName = vbNullString
            'Reset match state
                FormMatchesPicker = False
            'Reset captured step error
                StepErrNumber = 0
            'Reset captured step error
                StepErrDescription = vbNullString
            'Suppress metadata-read errors for unusual transient form states
                On Error Resume Next
            'Capture the loaded form class name
                FormTypeName = VBA.TypeName(CurForm)
            'Capture the loaded form runtime name
                FormRuntimeName = VBA.CStr(CurForm.Name)
            'Capture metadata-read failure if any
                StepErrNumber = Err.Number
            'Capture metadata-read failure if any
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
            'Evaluate whether the loaded form is a DatePicker instance
                FormMatchesPicker = _
                    (VBA.StrComp(FormTypeName, DP_FORM_NAME, vbTextCompare) = 0 Or _
                     VBA.StrComp(FormRuntimeName, DP_FORM_NAME, vbTextCompare) = 0)
            'Collect matching DatePicker forms by class name or runtime name
                If FormMatchesPicker Then
                    FormsToUnload.Add CurForm
                End If
        Next CurForm

'------------------------------------------------------------------------------
' STOP TIMER
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Stop timer"
    'Reset captured step error
        StepErrNumber = 0
    'Reset captured step error
        StepErrDescription = vbNullString
    'Suppress timer-stop errors because Application.OnTime cancellation can fail
        On Error Resume Next
    'Stop any active or stale DatePicker timer before continuing
        M_Timer_Stop
    'Capture timer-stop failure if any
        StepErrNumber = Err.Number
    'Capture timer-stop failure if any
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
' EXIT WHEN NO PICKER IS LOADED
'------------------------------------------------------------------------------
    'Exit cleanly after stale-timer cleanup when there is no loaded DatePicker form
        If FormsToUnload.Count = 0 Then GoTo CleanExit

'------------------------------------------------------------------------------
' UNLOAD MATCHING FORMS
'------------------------------------------------------------------------------
    'Track the current step
        StepName = "Unload DatePicker forms"
    'Unload collected DatePicker forms in reverse order
        For Index = FormsToUnload.Count To 1 Step -1
            'Resolve the collected form reference
                Set LoadedForm = FormsToUnload.Item(Index)
            'Reset captured step error
                StepErrNumber = 0
            'Reset captured step error
                StepErrDescription = vbNullString
            'Suppress individual hide and unload errors
                On Error Resume Next
            'Hide the form first to reduce visual flicker
                LoadedForm.Hide
            'Clear any non-fatal hide error
                Err.Clear
            'Unload the DatePicker form
                Unload LoadedForm
            'Capture unload failure if any
                StepErrNumber = Err.Number
            'Capture unload failure if any
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
'


Public Sub M_Settings_UseLocalNames(ByVal UseLocalDayNames As Boolean)

'
'------------------------------------------------------------------------------
'                           SETTINGS USE LOCAL NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Compatibility wrapper for setting whether DatePicker day names use local
'   language names
'
' WHY THIS EXISTS
'   Older calling code may still call M_Settings_UseLocalNames
'
'   The preferred settings routine is M_Settings_SetUseLocalNames, so this
'   wrapper preserves the legacy entry point while centralizing the actual
'   settings update in the preferred routine
'
' INPUTS
'   UseLocalDayNames
'     True to use local day names
'     False to use the standard configured day-name behavior
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates directly to M_Settings_SetUseLocalNames
'   Does not validate, transform, or reinterpret the input value
'   Does not perform any additional settings logic locally
'
' ERROR POLICY
'   Does not raise custom errors
'   Does not suppress errors
'   Lets M_Settings_SetUseLocalNames error behavior propagate unchanged
'
' DEPENDENCIES
'   M_Settings_SetUseLocalNames
'
' NOTES
'   Keep this routine as a thin compatibility wrapper
'   Do not add duplicate settings logic here
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE SETTING UPDATE
'------------------------------------------------------------------------------
    'Delegate to the preferred setting routine
        M_Settings_SetUseLocalNames UseLocalDayNames

End Sub


Private Function M_GetQualifiedMacroName(ByVal ProcedureName As String) As String

'
'------------------------------------------------------------------------------
'                           GET QUALIFIED MACRO NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a workbook-qualified macro name suitable for Excel callbacks
'
' WHY THIS EXISTS
'   Shape.OnAction, CommandBarButton.OnAction, Application.OnTime, Application.
'   OnKey, and Ribbon-related callback paths should resolve back to the workbook
'   that contains this module
'
'   This helper also caches resolved callback names because it is used by repeated
'   UI integration paths such as grid-icon movement, timer scheduling, keyboard
'   shortcut registration, and context-menu creation
'
' INPUTS
'   ProcedureName
'     Public procedure name to qualify
'
' RETURNS
'   Workbook-qualified macro name
'
' BEHAVIOR
'   Trims the supplied procedure name
'   Rejects a blank procedure name
'   Initializes a late-bound dictionary cache when needed
'   Invalidates the cache automatically when ThisWorkbook.Name changes
'   Returns a cached qualified name when available
'   Otherwise builds, stores, and returns the qualified macro name
'
' ERROR POLICY
'   Raises a descriptive runtime error when ProcedureName is blank or when the
'   callback name cannot be resolved
'
' DEPENDENCIES
'   ThisWorkbook
'   Scripting.Dictionary through late binding
'   VBA.Replace
'   VBA.Trim$
'
' NOTES
'   Expected output format:
'     'WorkbookName.xlsm'!ProcedureName
'
'   The workbook name is escaped by doubling apostrophes
'
'   The cache is workbook-name aware, so Save As / rename scenarios are handled
'   without requiring a manual reset
'
'   The routine intentionally remains private because callers should not depend
'   on the cache implementation
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME                 As String = "M_GetQualifiedMacroName"

    Dim NormalizedProcedureName     As String       'Trimmed callback procedure name
    Dim CurrentWorkbookName         As String       'Current host workbook name
    Dim QualifiedMacroName          As String       'Resolved workbook-qualified macro name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Normalize the supplied procedure name once
        NormalizedProcedureName = VBA.Trim$(ProcedureName)
    'Reject an empty procedure name
        If VBA.Len(NormalizedProcedureName) = 0 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "ProcedureName cannot be empty"
        End If

'------------------------------------------------------------------------------
' RESOLVE WORKBOOK NAME
'------------------------------------------------------------------------------
    'Read the current host workbook name
        CurrentWorkbookName = VBA.CStr(ThisWorkbook.Name)
    'Reject an empty workbook name
        If VBA.Len(CurrentWorkbookName) = 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "ThisWorkbook.Name cannot be empty"
        End If

'------------------------------------------------------------------------------
' INITIALIZE OR REFRESH CACHE
'------------------------------------------------------------------------------
    'Create the cache when it does not exist
        If mQualifiedMacroNameCache Is Nothing Then
            Set mQualifiedMacroNameCache = VBA.CreateObject("Scripting.Dictionary")
            mQualifiedMacroNameCache.CompareMode = vbTextCompare
            mQualifiedMacroWorkbookName = CurrentWorkbookName
        End If
    'Recreate the cache when the workbook name changed
        If VBA.StrComp(mQualifiedMacroWorkbookName, CurrentWorkbookName, vbBinaryCompare) <> 0 Then
            Set mQualifiedMacroNameCache = VBA.CreateObject("Scripting.Dictionary")
            mQualifiedMacroNameCache.CompareMode = vbTextCompare
            mQualifiedMacroWorkbookName = CurrentWorkbookName
        End If

'------------------------------------------------------------------------------
' RETURN CACHED VALUE
'------------------------------------------------------------------------------
    'Return the cached qualified macro name when available
        If mQualifiedMacroNameCache.Exists(NormalizedProcedureName) Then
            M_GetQualifiedMacroName = VBA.CStr(mQualifiedMacroNameCache(NormalizedProcedureName))
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD QUALIFIED NAME
'------------------------------------------------------------------------------
    'Build the workbook-qualified macro name
        QualifiedMacroName = _
            "'" & VBA.Replace(CurrentWorkbookName, "'", "''") & "'!" & NormalizedProcedureName

'------------------------------------------------------------------------------
' CACHE QUALIFIED NAME
'------------------------------------------------------------------------------
    'Store the resolved callback name for reuse
        mQualifiedMacroNameCache.Add NormalizedProcedureName, QualifiedMacroName

'------------------------------------------------------------------------------
' RETURN QUALIFIED NAME
'------------------------------------------------------------------------------
    'Return the resolved callback name
        M_GetQualifiedMacroName = QualifiedMacroName

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
            "Qualified macro-name resolution failed: " & Err.Description

End Function

Private Sub M_GetQualifiedMacroName_ClearCache()

'
'------------------------------------------------------------------------------
'                       CLEAR QUALIFIED MACRO NAME CACHE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears cached workbook-qualified macro names
'
' WHY THIS EXISTS
'   Callback names are session-level helper values. They should be releasable
'   during DatePicker teardown, runtime repair, or workbook lifecycle cleanup
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Releases the late-bound dictionary cache and clears the workbook-name marker
'
' ERROR POLICY
'   Best-effort cleanup
'
'   Suppresses errors because cache cleanup should never interrupt workbook
'   close, DatePicker stop, or runtime repair logic
'
' DEPENDENCIES
'   mQualifiedMacroNameCache
'   mQualifiedMacroWorkbookName
'
' NOTES
'   The cache also self-invalidates when ThisWorkbook.Name changes, so this
'   helper is cleanup-oriented rather than required for correctness
'
' UPDATED
'   2026-05-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' CLEAR CACHE
'------------------------------------------------------------------------------
    'Release the qualified macro-name cache
        Set mQualifiedMacroNameCache = Nothing
    'Clear the cached workbook-name marker
        mQualifiedMacroWorkbookName = vbNullString

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'------------------------------------------------------------------------------
'
'                               RIBBON CALLBACKS
'
'------------------------------------------------------------------------------
'



Public Sub Ribbon_ShowPicker(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON SHOW PICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Opens or refreshes the Date / Time Picker from a Ribbon button
'
' WHY THIS EXISTS
'   RibbonX onAction callbacks require public procedures in a standard module.
'   This routine provides a clean Ribbon entry point without exposing Ribbon
'   logic inside the DatePicker manager, UserForm, or label-hook classes
'
' INPUTS
'   Control
'     Ribbon control that triggered the callback
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates DatePicker display to DP_Show
'
' ERROR POLICY
'   Catches runtime errors and reports them through a user-facing message box
'   so Ribbon callback failures do not fail silently
'
' DEPENDENCIES
'   IRibbonControl
'   DP_Show
'   DP_MSGBOX_TITLE
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_ShowPicker"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SHOW PICKER
'------------------------------------------------------------------------------
    'Open or refresh the Date / Time Picker
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
    'Report the Ribbon callback failure
        Ribbon_ReportError PROC_NAME, Err.Number, Err.Description

End Sub

Public Sub Ribbon_Reset(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON RESET DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Repairs the DatePicker runtime from a Ribbon button
'
' WHY THIS EXISTS
'   A visible Ribbon repair action is useful when Excel events, transient UI,
'   keyboard shortcuts, modeless form state, or in-grid icons need to be
'   refreshed during a live workbook session
'
' INPUTS
'   Control
'     Ribbon control that triggered the callback
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates runtime repair to DP_RepairRuntime and displays a confirmation
'   message when repair completes successfully
'
' ERROR POLICY
'   Catches runtime errors and reports them through Ribbon_ReportError
'
' DEPENDENCIES
'   IRibbonControl
'   DP_RepairRuntime
'   Ribbon_ReportInfo
'   Ribbon_ReportError
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_Reset"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REPAIR RUNTIME
'------------------------------------------------------------------------------
    'Repair DatePicker runtime state
        DP_RepairRuntime

'------------------------------------------------------------------------------
' CONFIRM SUCCESS
'------------------------------------------------------------------------------
    'Report successful runtime repair to the user
        Ribbon_ReportInfo _
            "Date / Time Picker runtime repair completed successfully." & _
            VBA.vbCrLf & VBA.vbCrLf & _
            "You can now select a date cell or click Show Picker again."

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Report the Ribbon callback failure
        Ribbon_ReportError PROC_NAME, Err.Number, Err.Description

End Sub


Public Sub Ribbon_Demo(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON TOGGLE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Shows or hides the DatePicker demo worksheet from a Ribbon button
'
' WHY THIS EXISTS
'   The demo worksheet should normally remain very hidden during normal workbook
'   use, but the user should be able to open and close it explicitly from the
'   Ribbon without relying on save events or AutoSave-sensitive cleanup logic
'
' INPUTS
'   Control
'     Ribbon control that triggered the callback
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   If the demo sheet is currently visible, hides it using xlSheetVeryHidden
'   If the demo sheet is currently hidden or very hidden, shows and activates it
'
' ERROR POLICY
'   Catches runtime errors and reports them through Ribbon_ReportError
'
' DEPENDENCIES
'   IRibbonControl
'   ThisWorkbook
'   DP_DEMO_SHEET_NAME
'   DP_DemoSheet_Show
'   DP_DemoSheet_HideVeryHidden
'   Ribbon_ReportError
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
'   This routine uses ThisWorkbook rather than ActiveWorkbook so the Ribbon
'   callback always targets the DatePicker host workbook
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_Demo"

    Dim DemoSheet   As Excel.Worksheet       'DatePicker demo worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE DEMO SHEET
'------------------------------------------------------------------------------
    'Retrieve the demo worksheet from the host workbook
        Set DemoSheet = ThisWorkbook.Worksheets(DP_DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' TOGGLE DEMO SHEET
'------------------------------------------------------------------------------
    'Hide the demo sheet when it is already visible
        With DemoSheet
            If .Visible = xlSheetVisible Then
                DP_DemoSheet_HideVeryHidden
            Else
                .Visible = xlSheetVisible
                .Activate
            End If
        End With

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Report the Ribbon callback failure
        Ribbon_ReportError PROC_NAME, Err.Number, Err.Description

End Sub


Public Sub Ribbon_HideDemo(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON HIDE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the DatePicker demo worksheet from a Ribbon button
'
' WHY THIS EXISTS
'   The demo sheet can be shown explicitly through the Ribbon. This companion
'   action lets the user hide it again without relying on save events, which may
'   be triggered automatically by AutoSave
'
' INPUTS
'   Control
'     Ribbon control that triggered the callback
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates demo-sheet hiding to DP_DemoSheet_HideVeryHidden
'
' ERROR POLICY
'   Catches runtime errors and reports them through Ribbon_ReportError
'
' DEPENDENCIES
'   IRibbonControl
'   DP_DemoSheet_HideVeryHidden
'   Ribbon_ReportError
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_HideDemo"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' HIDE DEMO
'------------------------------------------------------------------------------
    'Hide the demo worksheet
        DP_DemoSheet_HideVeryHidden

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Report the Ribbon callback failure
        Ribbon_ReportError PROC_NAME, Err.Number, Err.Description

End Sub

Private Sub Ribbon_ReportError( _
    ByVal ProcedureName As String, _
    ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String)

'
'------------------------------------------------------------------------------
'                           REPORT RIBBON ERROR
'------------------------------------------------------------------------------
' PURPOSE
'   Reports Ribbon callback failures consistently
'
' WHY THIS EXISTS
'   Ribbon callbacks often fail silently or show poor diagnostics when errors
'   are not handled explicitly. This helper gives the user a clear message and
'   writes the same diagnostic to the Immediate Window
'
' INPUTS
'   ProcedureName
'     Name of the Ribbon callback that failed
'
'   ErrorNumber
'     Captured runtime error number
'
'   ErrorDescription
'     Captured runtime error description
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Builds a diagnostic message, writes it to the Immediate Window, and displays
'   it in a message box
'
' ERROR POLICY
'   Suppresses reporting errors so the error reporter cannot raise a secondary
'   Ribbon callback failure
'
' DEPENDENCIES
'   DP_MSGBOX_TITLE
'
' NOTES
'   Keep this helper private. RibbonX should call only the public callback
'   procedures
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim MessageText As String       'User-facing diagnostic message

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress secondary reporting errors
        On Error Resume Next

'------------------------------------------------------------------------------
' BUILD MESSAGE
'------------------------------------------------------------------------------
    'Build the diagnostic message
        MessageText = _
            "Date / Time Picker Ribbon action failed." & VBA.vbCrLf & VBA.vbCrLf & _
            "Procedure: " & ProcedureName & VBA.vbCrLf & _
            "Error: " & VBA.CStr(ErrorNumber) & VBA.vbCrLf & _
            "Description: " & ErrorDescription

'------------------------------------------------------------------------------
' REPORT MESSAGE
'------------------------------------------------------------------------------
    'Write the diagnostic message to the Immediate Window
        Debug.Print MessageText

    'Display the diagnostic message to the user
        VBA.MsgBox MessageText, vbExclamation, DP_MSGBOX_TITLE

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed reporting error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub



Public Sub DP_DemoSheet_Show()

'
'------------------------------------------------------------------------------
'                           SHOW DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Shows and activates the DatePicker demo worksheet
'
' WHY THIS EXISTS
'   The demo worksheet is intentionally hidden during normal workbook startup
'   and save lifecycle. A Ribbon action needs a controlled way to make it visible
'   again for showcase or validation purposes
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the demo worksheet from ThisWorkbook, makes it visible, activates
'   the workbook window when possible, and activates the demo sheet
'
' ERROR POLICY
'   Raises a descriptive runtime error if the demo worksheet cannot be resolved
'   or activated
'
' DEPENDENCIES
'   ThisWorkbook
'   Excel.Worksheet
'   DP_DEMO_SHEET_NAME
'
' NOTES
'   Uses ThisWorkbook rather than ActiveWorkbook so the callback always targets
'   the DatePicker host workbook
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_DemoSheet_Show"

    Dim DemoSheet       As Excel.Worksheet       'DatePicker demo worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE DEMO SHEET
'------------------------------------------------------------------------------
    'Retrieve the demo worksheet from the host workbook
        Set DemoSheet = ThisWorkbook.Worksheets(DP_DEMO_SHEET_NAME)

'------------------------------------------------------------------------------
' SHOW DEMO SHEET
'------------------------------------------------------------------------------
    'Make the demo sheet visible
        DemoSheet.Visible = xlSheetVisible

'------------------------------------------------------------------------------
' ACTIVATE HOST WORKBOOK WINDOW
'------------------------------------------------------------------------------
    'Suppress window activation errors for hidden or add-in-like contexts
        On Error Resume Next
    'Activate the first workbook window when available
        If ThisWorkbook.Windows.Count > 0 Then ThisWorkbook.Windows(1).Activate
    'Clear any suppressed window activation error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' ACTIVATE DEMO SHEET
'------------------------------------------------------------------------------
    'Activate the demo worksheet
        DemoSheet.Activate

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
            "Unable to show the DatePicker demo sheet '" & DP_DEMO_SHEET_NAME & "'. " & Err.Description

End Sub

Public Sub DP_DemoSheet_HideVeryHidden()

'
'------------------------------------------------------------------------------
'                       HIDE DEMO SHEET VERY HIDDEN
'------------------------------------------------------------------------------
' PURPOSE
'   Hides the DatePicker demo worksheet using xlSheetVeryHidden
'
' WHY THIS EXISTS
'   The demo worksheet is useful for showcase and validation, but it should not
'   remain visible in normal user workflows or be accidentally saved as the
'   active workbook surface
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the demo worksheet, activates a safe visible non-demo worksheet
'   when needed, and sets the demo worksheet to xlSheetVeryHidden
'
' ERROR POLICY
'   Best-effort cleanup. Suppresses errors because workbook open, save, and
'   close lifecycle should not be interrupted by a missing demo sheet, protected
'   workbook structure, or unavailable workbook window
'
' DEPENDENCIES
'   ThisWorkbook
'   Excel.Worksheet
'   DP_DEMO_SHEET_NAME
'   DP_DemoSheet_GetSafeVisibleSheet
'
' NOTES
'   Excel requires at least one visible worksheet. This routine therefore exits
'   safely when no visible non-demo worksheet is available
'
'   Safe to call repeatedly
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim DemoSheet       As Excel.Worksheet       'DatePicker demo worksheet
    Dim SafeSheet       As Excel.Worksheet       'Visible non-demo worksheet

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress lifecycle cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' RESOLVE DEMO SHEET
'------------------------------------------------------------------------------
    'Retrieve the demo worksheet from the host workbook
        Set DemoSheet = ThisWorkbook.Worksheets(DP_DEMO_SHEET_NAME)

    'Exit if the demo worksheet is not available
        If DemoSheet Is Nothing Then GoTo CleanExit

'------------------------------------------------------------------------------
' EXIT IF ALREADY VERY HIDDEN
'------------------------------------------------------------------------------
    'Exit if the demo worksheet is already very hidden
        If DemoSheet.Visible = xlSheetVeryHidden Then GoTo CleanExit

'------------------------------------------------------------------------------
' HANDLE NON-VISIBLE DEMO SHEET
'------------------------------------------------------------------------------
    'Convert ordinary hidden state to very hidden without changing active sheet
        If DemoSheet.Visible <> xlSheetVisible Then
            DemoSheet.Visible = xlSheetVeryHidden
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' RESOLVE SAFE VISIBLE SHEET
'------------------------------------------------------------------------------
    'Find a visible non-demo worksheet before hiding the demo worksheet
        Set SafeSheet = DP_DemoSheet_GetSafeVisibleSheet(DemoSheet)

    'Exit if no safe visible worksheet exists
        If SafeSheet Is Nothing Then GoTo CleanExit

'------------------------------------------------------------------------------
' MOVE ACTIVE SHEET WHEN NEEDED
'------------------------------------------------------------------------------
    'Activate the safe sheet when the demo sheet is currently active
        If Excel.Application.ActiveSheet Is DemoSheet Then SafeSheet.Activate

'------------------------------------------------------------------------------
' HIDE DEMO SHEET
'------------------------------------------------------------------------------
    'Set the demo worksheet to very hidden
        DemoSheet.Visible = xlSheetVeryHidden

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Clear any suppressed lifecycle error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Function DP_DemoSheet_GetSafeVisibleSheet( _
    ByVal DemoSheet As Excel.Worksheet) As Excel.Worksheet

'
'------------------------------------------------------------------------------
'                       GET SAFE VISIBLE SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a visible non-demo worksheet that can remain visible while the demo
'   worksheet is hidden
'
' WHY THIS EXISTS
'   Excel does not allow all worksheets in a workbook to be hidden. Before the
'   demo worksheet is made very hidden, the workbook must have another visible
'   worksheet available
'
' INPUTS
'   DemoSheet
'     Demo worksheet that should be excluded from the search
'
' RETURNS
'   First visible non-demo worksheet found in ThisWorkbook
'   Nothing when no suitable worksheet exists
'
' BEHAVIOR
'   Scans ThisWorkbook.Worksheets and returns the first visible worksheet that
'   is not the supplied demo worksheet
'
' ERROR POLICY
'   Best-effort lookup. Returns Nothing on error
'
' DEPENDENCIES
'   ThisWorkbook
'   Excel.Worksheet
'
' NOTES
'   This helper does not create sheets and does not change visibility
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim WS As Excel.Worksheet       'Worksheet scan variable

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress lookup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' FIND SAFE SHEET
'------------------------------------------------------------------------------
    'Loop through worksheets in the host workbook
        For Each WS In ThisWorkbook.Worksheets
            'Return the first visible worksheet that is not the demo sheet
                If Not WS Is DemoSheet Then
                    If WS.Visible = xlSheetVisible Then
                        Set DP_DemoSheet_GetSafeVisibleSheet = WS
                        Exit Function
                    End If
                End If
        Next WS

'------------------------------------------------------------------------------
' FALLBACK
'------------------------------------------------------------------------------
    'Return Nothing when no safe visible sheet is available
        Set DP_DemoSheet_GetSafeVisibleSheet = Nothing

End Function

Private Sub Ribbon_ReportInfo(ByVal MessageText As String)

'
'------------------------------------------------------------------------------
'                           REPORT RIBBON INFO
'------------------------------------------------------------------------------
' PURPOSE
'   Displays a standard informational message for successful Ribbon actions
'
' WHY THIS EXISTS
'   Some Ribbon actions, such as runtime repair, complete without a visible UI
'   change. A confirmation message reassures the user that the action completed
'
' INPUTS
'   MessageText
'     Message to display to the user
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Writes the message to the Immediate Window and shows a standard information
'   message box
'
' ERROR POLICY
'   Suppresses reporting errors so the info reporter cannot raise a secondary
'   Ribbon callback failure
'
' DEPENDENCIES
'   DP_MSGBOX_TITLE
'
' NOTES
'   Keep this helper private. RibbonX should call only the public callbacks
'
' UPDATED
'   2026-05-15
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress secondary reporting errors
        On Error Resume Next

'------------------------------------------------------------------------------
' REPORT MESSAGE
'------------------------------------------------------------------------------
    'Write the message to the Immediate Window
        Debug.Print MessageText

    'Show the message to the user
        VBA.MsgBox MessageText, vbInformation, DP_MSGBOX_TITLE

'------------------------------------------------------------------------------
' EXIT
'------------------------------------------------------------------------------
    'Clear any suppressed reporting error
        Err.Clear

    'Restore normal error handling
        On Error GoTo 0

End Sub



