Attribute VB_Name = "M_DP_Core"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Holds the DatePicker implementation behind the M_DatePicker facade
'
' WHY THIS EXISTS
'   Option Private Module keeps the implementation out of the Alt+F8 macro
'   dialog and out of other projects, while the facade keeps the supported
'   names stable. This module is transitional: #56 to #59 split it into the
'   M_DP_* responsibility modules. #56 moved captions, date policy, qualified
'   macro names and native window services out; #57 moved the timer, context
'   menu, keyboard shortcut, grid icon and grid-icon asset
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Manages:
'     - settings, persistence namespace, and runtime state
'     - form show / close and initial-date bridge
'     - date-only calendar selection write-back
'     - direct Today and Now write-back commands
'     - single-cell, multi-cell, and table data-column write-back
'     - provider lease and runtime lifecycle, which calls the verified
'       teardown of each integration module
'
' ERROR POLICY
'   Public routines raise descriptive runtime errors unless explicitly documented
'   as best-effort or safe-default routines
'
' DEPENDENCIES
'   Excel object model
'   Office CommandBars object model
'   MSForms / VBA.UserForms
'   M_DatePicker (public enums and types only)
'   M_DP_Captions
'   M_DP_DatePolicy
'   M_DP_MacroNames
'   M_DP_Window
'   M_DP_Timer
'   M_DP_ContextMenu
'   M_DP_Keyboard
'   M_DP_GridIcon
'   cDatePickerManager
'   UF_DatePicker
'
' NOTES
'   Each supported or callback procedure of M_DatePicker has its implementation
'   here as <Name>_Impl. Internal code calls those procedures and never the
'   facade wrappers. Callback names registered with Excel are the facade names,
'   because Excel cannot run a procedure of a private module from outside it
'
'   Registry and context-menu identifiers retain VBA_DATETIMEPICKER as stable
'   legacy names for backward compatibility
'
'   DP_WriteResult is the one structured outcome of a write-back. Later write
'   policies extend it rather than adding a parallel reporting mechanism
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

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

    Private Const DP_DEMO_SHEET_NAME               As String = "DATE PICKER DEMO"        'Demo worksheet name
    Private Const DP_WRITEBACK_ADDRESS_LIMIT        As Long = 25                          'Maximum structured addresses retained per outcome category and operation

'------------------------------------------------------------------------------
' PUBLIC CONSTANTS
'------------------------------------------------------------------------------
    Public Const DP_FORM_NAME                      As String = "UF_DatePicker"           'DatePicker UserForm name

    Public Const DP_MIN_YEAR                       As Long = 100                         'Minimum supported DatePicker year
    Public Const DP_MAX_YEAR                       As Long = 9999                        'Maximum supported DatePicker year

    Public Const DP_SETTINGS_APP_NAME              As String = "VBA_DATETIMEPICKER"      'Legacy registry application name
    Private Const DP_SETTINGS_NAMESPACE_SEPARATOR  As String = "__"                      'Separates the legacy name from an optional namespace
    Private Const DP_LEASE_BAR_NAME                As String = "__VBA_DATETIMEPICKER_RUNTIME_PROVIDER_LEASE__"  'Application-wide provider lease
    Private Const DP_LEASE_MARKER_TAG              As String = "VBA_DATETIMEPICKER_RUNTIME_LEASE_OWNER"         'Identifies the lease marker control
    Private Const DP_LEASE_AMBIGUOUS               As String = "?"                       'Sentinel for a lease that cannot be read
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
' TEST FAULT INJECTION STATE
'------------------------------------------------------------------------------
    'Armed target area ordinal for a forced write-back failure, zero when disarmed
    Private mDP_TestWriteFailArea           As Long
    'Cells to write inside that area before the forced failure fires
    Private mDP_TestWriteFailAfterCells     As Long
    'Target areas populated so far in the current write operation
    Private mDP_TestWriteAreaOrdinal        As Long

'------------------------------------------------------------------------------
' PUBLIC STATE
'------------------------------------------------------------------------------
    Public gDP_Manager                  As cDatePickerManager   'DatePicker manager/controller
    Public gDP_WriteValue               As Date                 'Picked value written to Excel

    Public gDP_ShowRightClick           As Boolean              'Show right-click menu entry
    Public gDP_ShowGridIcon             As Boolean              'Show in-grid icon
    Public gDP_FirstDayOfWeek           As Long                 'vbSunday or vbMonday
    Public gDP_UseLocalNames            As Boolean              'Use local day/month captions
    Public gDP_ClockMode                As DP_ClockMode         'Static or live clock
    Public gDP_SizeMode                 As DP_SizeMode          'Normal or compact layout
    Public gDP_HighlightWeekends        As Boolean              'Highlight weekend days
    Public gDP_AllowOutsideMonthSelection As Boolean            'Allow outside-month day selection

    Public gDP_CloseAfterSelection      As Boolean              'Close picker after successful write-back
    Public gDP_UseWinAPI                As Boolean              'Allow optional WinAPI-dependent native behavior
    Public gDP_EnableKeyboardShortcut   As Boolean              'Ctrl+Shift+D shortcut enabled
    Public gDP_HolidayCallbackName      As String               'Optional holiday callback name

    Public gDP_InitialDate              As Date                 'Initial date for next form instance
    Public gDP_HasInitialDate           As Boolean              'True when initial date is available
    Public gDP_SelectedDate             As Date                 'Currently selected date
    Public gDP_HasSelectedDate          As Boolean              'True when selected date is available

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mSettingsLoaded             As Boolean              'Settings loaded flag
    Private mDP_SettingsNamespace       As String               'Optional persistence namespace, empty for the legacy default
    Private mDP_RuntimeOwnerId          As String               'Ephemeral lease token, non-empty only while this project owns the lease
    Private mDP_LeaseRefusalSilenced    As Boolean              'Test-only: suppress the modal refusal report so entry paths stay drivable
    Private mDP_LeaseRefusalCount       As Long                 'Test-only: refusal reports raised while silenced

    'Issue #50 lifecycle transaction / deterministic fault state
    Private mDP_LifecycleFaultStep          As String           'One-shot lifecycle fault step, blank when disarmed
    Private mDP_LifecycleFaultNumber        As Long             'One-shot lifecycle fault number
    Private mDP_LifecycleLastTrace          As String           'Ordered cleanup trace from the last lifecycle transaction
    Private mDP_LifecycleLastCleanupDetail  As String           'First cleanup failure from the last lifecycle transaction
    Private mDP_LifecycleLastCriticalClean  As Boolean          'True when every critical cleanup step proved clean
    Private mDP_LifecycleLastLeaseReleased  As Boolean          'True when the last transaction verified lease release
    Private mDP_LifecycleLastPrimaryNumber  As Long             'Primary lifecycle failure number preserved across cleanup
    Private mDP_LifecycleLastPrimaryStep    As String           'Primary lifecycle failure step preserved across cleanup
    Private mDP_LifecycleLastPrimaryDescription As String       'Primary lifecycle failure description preserved across cleanup
    Private mDP_LifecycleLastOperation      As String           'Lifecycle entry point represented by the current diagnostic state
    Private mDP_LifecycleLastSucceeded      As Boolean          'True only when the lifecycle entry point completed its contract
    Private mDP_LifecycleLastCleanupAttempted As Boolean        'True once the lifecycle cleanup transaction was entered
    Private mDP_LifecycleLastCleanupFailureCount As Long        'Number of cleanup steps that actually failed
    Private mDP_LifecycleLastLeaseWasAlreadyOwned As Boolean    'True when the provider lease pre-dated the lifecycle call
    Private mDP_LifecycleLastLeaseAcquiredThisCall As Boolean   'True when the lifecycle call acquired the provider lease


'------------------------------------------------------------------------------
'
'                                   SETTINGS
'
'------------------------------------------------------------------------------

Private Function M_Settings_GetEffectiveAppName() As String

'
'------------------------------------------------------------------------------
'                     RESOLVE EFFECTIVE APPLICATION NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the VBA registry application name every settings read and write uses
'
' WHY THIS EXISTS
'   Every deployment persisted to one application name, so two workbooks that
'   never ran at the same time could still overwrite each other's preferences.
'   Loading is itself a write boundary here, because M_Settings_Load persists the
'   normalized values it just read
'
'   Resolution is centralized so no call site constructs the name itself. A
'   namespace applied at 28 call sites would be wrong at one of them eventually
'
' INPUTS
'   None
'
' RETURNS
'   The legacy application name when no namespace is configured
'
'   The namespaced application name otherwise
'
' BEHAVIOR
'   Appends the configured namespace to the legacy identifier
'
' ERROR POLICY
'   Does not raise. The namespace was validated when it was configured
'
' DEPENDENCIES
'   DP_SETTINGS_APP_NAME
'   mDP_SettingsNamespace
'
' NOTES
'   DP_SETTINGS_APP_NAME is unchanged and still means the default application
'   name. An installation that configures no namespace reads and writes exactly
'   where earlier releases did
'
'   The separator is a double underscore, which the namespace validator rejects
'   inside a namespace, so two different namespaces cannot resolve to one
'   effective name
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Use the legacy application name when nothing is configured
        If VBA.LenB(mDP_SettingsNamespace) = 0 Then
            M_Settings_GetEffectiveAppName = DP_SETTINGS_APP_NAME
        Else
            M_Settings_GetEffectiveAppName = DP_SETTINGS_APP_NAME & _
                DP_SETTINGS_NAMESPACE_SEPARATOR & mDP_SettingsNamespace
        End If

End Function

Public Sub M_Settings_SetNamespace_Impl(ByVal NamespaceName As String)

'
'------------------------------------------------------------------------------
'                        SET SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetNamespace
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetNamespace
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetNamespace
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetNamespace
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetNamespace. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   mSettingsLoaded
'   M_Settings_ValidateNamespace
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Settings_SetNamespace"

    Dim Normalized      As String       'Validated namespace

'------------------------------------------------------------------------------
' REJECT A LATE CHANGE
'------------------------------------------------------------------------------
    'Settings already in memory came from the current namespace. Repointing now
    'would write them into a different one
        If mSettingsLoaded Then
            Err.Raise vbObjectError + 2601, PROC_NAME, _
                "The settings namespace cannot be changed after settings have " & _
                "been loaded. Configure it before any DatePicker entry point runs."
        End If

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Resolve the namespace this deployment will persist under
        Normalized = M_Settings_ValidateNamespace(NamespaceName)

'------------------------------------------------------------------------------
' STORE
'------------------------------------------------------------------------------
    'Apply the validated namespace to every later read and write
        mDP_SettingsNamespace = Normalized

End Sub

Public Function M_Settings_GetNamespace_Impl() As String

'
'------------------------------------------------------------------------------
'                        GET SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetNamespace
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetNamespace
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetNamespace
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetNamespace
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetNamespace. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   mDP_SettingsNamespace
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN NAMESPACE
'------------------------------------------------------------------------------
    'Report the configured namespace
        M_Settings_GetNamespace_Impl = mDP_SettingsNamespace

End Function

Private Function M_Settings_ValidateNamespace( _
    ByVal NamespaceName As String) As String

'
'------------------------------------------------------------------------------
'                       VALIDATE SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Normalizes a caller-supplied namespace and rejects one that cannot be used
'
' WHY THIS EXISTS
'   The namespace becomes part of a registry application name. A value containing
'   a path separator, a control character or the reserved separator would either
'   fail unpredictably or let two different namespaces resolve to the same
'   effective name, which would defeat the isolation entirely
'
' INPUTS
'   NamespaceName
'     Caller-supplied namespace
'
' RETURNS
'   The normalized namespace
'
'   An empty string when the caller supplied only whitespace or nothing, which
'   selects the legacy default
'
' BEHAVIOR
'   Trims surrounding whitespace, then rejects a value that is too long or
'   contains a character that would make the effective name ambiguous
'
' ERROR POLICY
'   Raises a descriptive error naming the offending condition
'
' DEPENDENCIES
'   DP_SETTINGS_NAMESPACE_SEPARATOR
'
' NOTES
'   Validation is deterministic: the same input always resolves to the same
'   namespace, and two different valid namespaces never resolve to one effective
'   application name
'
'   An empty namespace is not an error. It is how a caller returns to the legacy
'   default before settings load
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Settings_ValidateNamespace"
    Const MAX_LENGTH        As Long = 64            'Longest supported namespace

    Dim Normalized          As String       'Trimmed namespace
    Dim CharIndex           As Long         'Current character position
    Dim CharCode            As Long         'Current character code

'------------------------------------------------------------------------------
' NORMALIZE
'------------------------------------------------------------------------------
    'Trim surrounding whitespace so " Books " and "Books" are one namespace
        Normalized = VBA.Trim$(NamespaceName)
    'An empty namespace selects the legacy default and is not an error
        If VBA.LenB(Normalized) = 0 Then
            M_Settings_ValidateNamespace = VBA.vbNullString
            Exit Function
        End If

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Bound the length so the effective application name stays usable
        If VBA.Len(Normalized) > MAX_LENGTH Then
            Err.Raise vbObjectError + 2602, PROC_NAME, _
                "The settings namespace cannot exceed " & VBA.CStr(MAX_LENGTH) & _
                " characters."
        End If
    'Reject the reserved separator, which would let two namespaces collide
        If VBA.InStr(1, Normalized, DP_SETTINGS_NAMESPACE_SEPARATOR, vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2603, PROC_NAME, _
                "The settings namespace cannot contain " & _
                DP_SETTINGS_NAMESPACE_SEPARATOR & "."
        End If
    'Reject characters that would make the registry location ambiguous
        For CharIndex = 1 To VBA.Len(Normalized)
            CharCode = VBA.Asc(VBA.Mid$(Normalized, CharIndex, 1))
            If CharCode < 32 Then
                Err.Raise vbObjectError + 2604, PROC_NAME, _
                    "The settings namespace cannot contain control characters."
            End If
            Select Case VBA.Mid$(Normalized, CharIndex, 1)
                Case "\", "/", ":", "*", "?", """", "<", ">", "|"
                    Err.Raise vbObjectError + 2605, PROC_NAME, _
                        "The settings namespace cannot contain the character " & _
                        VBA.Mid$(Normalized, CharIndex, 1) & "."
            End Select
        Next CharIndex

'------------------------------------------------------------------------------
' RETURN NORMALIZED NAMESPACE
'------------------------------------------------------------------------------
    'Return the validated namespace
        M_Settings_ValidateNamespace = Normalized

End Function

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
'   or unsupported values to defaults, and saves the normalized state back to
'   the registry. Valid access-path settings are kept as stored, even when no
'   built-in entry point remains enabled (see NOTES)
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
'   Access-path settings reflect explicit configuration. Disabling right-click
'   and in-grid icon access no longer re-enables the keyboard shortcut, because
'   Application.OnKey is a session-wide binding and the DatePicker must not take
'   it on behalf of a user who did not choose it
'
'   Zero built-in interactive access paths is a permitted configuration. The
'   picker is then opened through DP_Show, a Ribbon integration, or a
'   caller-supplied entry point
'
' UPDATED
'   2026-11-09
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOSE_AFTER_SELECTION, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOCK_MODE, _
            VBA.vbNullString)
    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_SIZE_MODE, _
            VBA.vbNullString)

    'Fall back to the legacy Display section when needed
        If VBA.LenB(RawValue) = 0 Then
            RawValue = GetSetting( _
                M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_HOLIDAY_CALLBACK, _
            DP_DEFAULT_HOLIDAY_CALLBACK))

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
'   Access-path settings reflect explicit configuration. Disabling right-click
'   and in-grid icon access no longer re-enables the keyboard shortcut, because
'   Application.OnKey is a session-wide binding and the DatePicker must not take
'   it on behalf of a user who did not choose it
'
'   Zero built-in interactive access paths is a permitted configuration. The
'   picker is then opened through DP_Show, a Ribbon integration, or a
'   caller-supplied entry point
'
'   WinAPI styling is saved as disabled when the current platform does not
'   support WinAPI helpers
'
' UPDATED
'   2026-08-23
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_FIRST_DAY_OF_WEEK, _
            VBA.CStr(gDP_FirstDayOfWeek)

    'Track the current handler step
        HandlerStep = "Save local-name setting"

    'Save the local-name setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_DISPLAY, _
            DP_SETTING_USE_LOCAL_NAMES, _
            M_Settings_BooleanToStorageValue(gDP_UseLocalNames)

    'Track the current handler step
        HandlerStep = "Save weekend-highlight setting"

    'Save the weekend-highlight setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_ALLOW_OUTSIDE_MONTH_SELECTION, _
            M_Settings_BooleanToStorageValue(gDP_AllowOutsideMonthSelection)

    'Track the current handler step
        HandlerStep = "Save close-after-selection setting"

    'Save the close-after-selection setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOSE_AFTER_SELECTION, _
            M_Settings_BooleanToStorageValue(gDP_CloseAfterSelection)

    'Track the current handler step
        HandlerStep = "Save clock-mode setting"

    'Save the clock mode setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_BEHAVIOR, _
            DP_SETTING_CLOCK_MODE, _
            VBA.CStr(VBA.CLng(gDP_ClockMode))

    'Track the current handler step
        HandlerStep = "Save size-mode setting"

    'Save the size mode setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_RIGHT_CLICK, _
            M_Settings_BooleanToStorageValue(gDP_ShowRightClick)

    'Track the current handler step
        HandlerStep = "Save in-grid icon setting"

    'Save the grid-icon setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_FEATURES, _
            DP_SETTING_SHOW_GRID_ICON, _
            M_Settings_BooleanToStorageValue(gDP_ShowGridIcon)

    'Track the current handler step
        HandlerStep = "Save keyboard shortcut setting"

    'Save the keyboard shortcut setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
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
            M_Settings_GetEffectiveAppName(), _
            DP_SETTINGS_SECTION_ADVANCED, _
            DP_SETTING_USE_WINAPI, _
            M_Settings_BooleanToStorageValue(gDP_UseWinAPI)

    'Track the current handler step
        HandlerStep = "Save holiday callback setting"

    'Save the holiday callback setting
        SaveSetting _
            M_Settings_GetEffectiveAppName(), _
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
'   Assigns default values to all persisted settings and applies safe platform
'   defaults. Access-path settings take their defaults and are not adjusted to
'   keep an entry point available (see NOTES)
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
'   Access-path settings reflect explicit configuration. Disabling right-click
'   and in-grid icon access no longer re-enables the keyboard shortcut, because
'   Application.OnKey is a session-wide binding and the DatePicker must not take
'   it on behalf of a user who did not choose it
'
'   Zero built-in interactive access paths is a permitted configuration. The
'   picker is then opened through DP_Show, a Ribbon integration, or a
'   caller-supplied entry point
'
' UPDATED
'   2026-11-09
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

Public Sub M_Settings_SetFirstDayOfWeek_Impl(ByVal FirstDayOfWeek As Long)

'
'------------------------------------------------------------------------------
'                           SET FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetFirstDayOfWeek
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeek
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeek
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeek
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeek. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetFirstDayOfWeekText_Impl(ByVal FirstDayOfWeekText As String)

'
'------------------------------------------------------------------------------
'                       SET FIRST DAY OF WEEK FROM TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetFirstDayOfWeekText
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeekText
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeekText
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeekText
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetFirstDayOfWeekText. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_TryParseFirstDayOfWeek
'   M_Settings_SetFirstDayOfWeek
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_SetFirstDayOfWeek_Impl ParsedValue

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

Public Sub M_Settings_SetUseLocalNames_Impl(ByVal UseLocalDayNames As Boolean)

'
'------------------------------------------------------------------------------
'                           SET USE LOCAL NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetUseLocalNames
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetUseLocalNames
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetUseLocalNames
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetUseLocalNames
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetUseLocalNames. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Function M_Settings_ResolveKeyboardShortcutOnSave( _
    ByVal CurrentEnableKeyboard As Boolean, _
    ByVal ResolvedShowRightClick As Boolean, _
    ByVal ResolvedShowGridIcon As Boolean) As Boolean

'
'==============================================================================
'                 RESOLVE KEYBOARD SHORTCUT ON SETTINGS SAVE
'==============================================================================
' PURPOSE
'   Resolves the keyboard shortcut setting a settings save should persist
'
' WHY THIS EXISTS
'   The settings panel has no keyboard checkbox. The shortcut is set through
'   M_Settings_SetEnableKeyboardShortcut, and a panel save must carry the
'   current value forward untouched
'
'   At v1.2.0 the UserForm save path instead forced the value back to True
'   whenever right-click and the grid icon were both disabled, then persisted it
'   and registered Application.OnKey. A user who had deliberately disabled all
'   three built-in entry points could not keep that configuration
'
'   The equivalent block was already removed from M_Settings_SetShowRightClick
'   and M_Settings_SetShowGridIcon, which is why those setters still carry an
'   empty PROTECT MANUAL ACCESS PATH banner. The UserForm copy was missed
'
' INPUTS
'   CurrentEnableKeyboard
'       The keyboard shortcut setting as currently resolved
'
'   ResolvedShowRightClick
'       The right-click setting the save is about to persist
'
'   ResolvedShowGridIcon
'       The grid-icon setting the save is about to persist
'
' RETURNS
'   CurrentEnableKeyboard, unchanged, for every combination of the other two
'
' BEHAVIOR
'   Zero built-in entry paths is a valid configuration. The DatePicker remains
'   reachable through DP_Show, DP_Click and the Ribbon, so no combination of
'   integration settings justifies overriding an explicit user choice
'
' ERROR POLICY
'   Pure. Does not raise, read global state or mutate anything
'
' NOTES
'   The two integration arguments are accepted deliberately even though they do
'   not affect the result. They are the exact inputs the removed fallback
'   consulted, so this signature states the contract that they must never affect
'   it, and lets a regression sweep every combination
'
'   This is the seam the settings-save regression drives. It is the same code
'   the real UserForm save handler executes, not a module-level setter standing
'   in for it. It is classified internal in src/public-surface.json
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Preserve the explicitly resolved setting exactly as the user chose it
        M_Settings_ResolveKeyboardShortcutOnSave = CurrentEnableKeyboard

End Function

Public Sub M_Settings_SetEnableKeyboardShortcut_Impl(ByVal EnableKeyboardShortcut As Boolean)

'
'------------------------------------------------------------------------------
'                       SET ENABLE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetEnableKeyboardShortcut
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetEnableKeyboardShortcut
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetEnableKeyboardShortcut
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetEnableKeyboardShortcut
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetEnableKeyboardShortcut. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Lease_IsOwner
'   M_Lease_ReportSettingNotApplied
'   M_KeyboardShortcut_Update
'   gDP_EnableKeyboardShortcut
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetEnableKeyboardShortcut"

    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting
    Dim SettingsChanged         As Boolean      'True when the resolved setting differs
    Dim OwnsRuntime             As Boolean      'True when this copy owns the provider lease
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
    'Deliberately empty. An earlier design forced the keyboard shortcut back on
    'when right-click and the grid icon were both disabled. Zero built-in entry
    'paths is a valid configuration, so the resolved value above is carried
    'forward unchanged. Do not reintroduce a fallback here: see #42
'------------------------------------------------------------------------------
' RESOLVE CHANGE FLAG
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve change flag"
    'Resolve whether the final keyboard shortcut setting changed
        SettingsChanged = (NewEnableKeyboard <> OldEnableKeyboard)

'------------------------------------------------------------------------------
' RESOLVE PROVIDER OWNERSHIP
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve provider ownership"
    'Only the copy that owns the provider lease may touch Application.OnKey (#194)
        OwnsRuntime = M_Lease_IsOwner()

'------------------------------------------------------------------------------
' SYNCHRONIZE ONLY WHEN SETTING IS UNCHANGED
'------------------------------------------------------------------------------
    'Synchronize Application.OnKey and exit when no persisted setting changed
        If Not SettingsChanged Then
            'Leave the shared binding to the owning copy and report why
                If Not OwnsRuntime Then
                    HandlerStep = "Report setting not applied"
                    M_Lease_ReportSettingNotApplied PROC_NAME
                    Exit Sub
                End If
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
' DEFER TO THE OWNING COPY
'------------------------------------------------------------------------------
    'Leave the shared binding to the owning copy and report why. The saved value
    'is applied by that copy, at the latest when DP_Start reads it
        If Not OwnsRuntime Then
            HandlerStep = "Report setting not applied"
            M_Lease_ReportSettingNotApplied PROC_NAME
            Exit Sub
        End If

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

Public Sub M_Settings_SetUseWinAPI_Impl(ByVal UseWinAPI As Boolean)

'
'------------------------------------------------------------------------------
'                           SET USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetUseWinAPI
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetUseWinAPI
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetUseWinAPI
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetUseWinAPI
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetUseWinAPI. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Platform_CanUseWinAPI
'   gDP_UseWinAPI
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetAllowOutsideMonthSelection_Impl(ByVal AllowOutsideMonthSelection As Boolean)

'
'------------------------------------------------------------------------------
'                       SET ALLOW OUTSIDE MONTH SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetAllowOutsideMonthSelection
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetAllowOutsideMonthSelection
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetAllowOutsideMonthSelection
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetAllowOutsideMonthSelection
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetAllowOutsideMonthSelection. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetHighlightWeekends_Impl(ByVal HighlightWeekends As Boolean)

'
'------------------------------------------------------------------------------
'                           SET HIGHLIGHT WEEKENDS
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetHighlightWeekends
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetHighlightWeekends
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetHighlightWeekends
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetHighlightWeekends
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetHighlightWeekends. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'   UF_DatePicker.UF_DP_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetCloseAfterSelection_Impl(ByVal CloseAfterSelection As Boolean)

'
'------------------------------------------------------------------------------
'                           SET CLOSE AFTER SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetCloseAfterSelection
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetCloseAfterSelection
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetCloseAfterSelection
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetCloseAfterSelection
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetCloseAfterSelection. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetClockMode_Impl(ByVal ClockMode As DP_ClockMode)

'
'------------------------------------------------------------------------------
'                           SET CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetClockMode
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetClockMode
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetClockMode
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetClockMode
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetClockMode. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_IsValidClockMode
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Timer_ApplyClockMode
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetSizeMode_Impl(ByVal SizeMode As DP_SizeMode)

'
'------------------------------------------------------------------------------
'                           SET SIZE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetSizeMode
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetSizeMode
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetSizeMode
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetSizeMode
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetSizeMode. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_IsValidSizeMode
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'   UF_DatePicker.UF_DP_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetHolidayCallback_Impl(ByVal HolidayCallbackName As String)

'
'------------------------------------------------------------------------------
'                           SET HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetHolidayCallback
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetHolidayCallback
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetHolidayCallback
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetHolidayCallback
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetHolidayCallback. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_FormBridge_RefreshSettings
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub M_Settings_SetShowRightClick_Impl(ByVal ShowRightClick As Boolean)

'
'------------------------------------------------------------------------------
'                           SET SHOW RIGHT-CLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetShowRightClick
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetShowRightClick
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetShowRightClick
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetShowRightClick
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetShowRightClick. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Lease_IsOwner
'   M_Lease_ReportSettingNotApplied
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetShowRightClick"

    Dim OldShowRightClick       As Boolean      'Previous right-click menu setting
    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting

    Dim RightClickChanged       As Boolean      'True when right-click menu setting changed
    Dim KeyboardChanged         As Boolean      'True when the keyboard shortcut setting changed
    Dim OwnsRuntime             As Boolean      'True when this copy owns the provider lease
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
    'Deliberately empty. An earlier design forced the keyboard shortcut back on
    'when right-click and the grid icon were both disabled. Zero built-in entry
    'paths is a valid configuration, so the resolved value above is carried
    'forward unchanged. Do not reintroduce a fallback here: see #42
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
' RESOLVE PROVIDER OWNERSHIP
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve provider ownership"

    'Only the copy that owns the provider lease may touch the shared menus (#194)
        OwnsRuntime = M_Lease_IsOwner()

'------------------------------------------------------------------------------
' SYNCHRONIZE ONLY WHEN SETTINGS ARE UNCHANGED
'------------------------------------------------------------------------------
    'Synchronize the right-click menu and exit when no persisted setting changed
        If Not RightClickChanged Then
            If Not KeyboardChanged Then
                'Leave the shared menus to the owning copy and report why
                    If Not OwnsRuntime Then
                        HandlerStep = "Report setting not applied"
                        M_Lease_ReportSettingNotApplied PROC_NAME
                        Exit Sub
                    End If
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
' DEFER TO THE OWNING COPY
'------------------------------------------------------------------------------
    'Leave the shared menus and binding to the owning copy and report why. The
    'saved value is applied by that copy, at the latest when DP_Start reads it
        If Not OwnsRuntime Then
            HandlerStep = "Report setting not applied"
            M_Lease_ReportSettingNotApplied PROC_NAME
            Exit Sub
        End If

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

Public Sub M_Settings_SetShowGridIcon_Impl(ByVal ShowGridIcon As Boolean)

'
'------------------------------------------------------------------------------
'                           SET SHOW GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_SetShowGridIcon
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_SetShowGridIcon
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_SetShowGridIcon
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_SetShowGridIcon
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_SetShowGridIcon. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_Save
'   M_Lease_IsOwner
'   M_Lease_ReportSettingNotApplied
'   M_KeyboardShortcut_Update
'   M_GridIcon_Remove
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Settings_SetShowGridIcon"

    Dim OldShowGridIcon         As Boolean      'Previous in-grid icon setting
    Dim OldEnableKeyboard       As Boolean      'Previous keyboard shortcut setting
    Dim NewEnableKeyboard       As Boolean      'Resolved keyboard shortcut setting

    Dim GridIconChanged         As Boolean      'True when in-grid icon setting changed
    Dim KeyboardChanged         As Boolean      'True when the keyboard shortcut setting changed
    Dim OwnsRuntime             As Boolean      'True when this copy owns the provider lease
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
    'Deliberately empty. An earlier design forced the keyboard shortcut back on
    'when right-click and the grid icon were both disabled. Zero built-in entry
    'paths is a valid configuration, so the resolved value above is carried
    'forward unchanged. Do not reintroduce a fallback here: see #42
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
' RESOLVE PROVIDER OWNERSHIP
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve provider ownership"

    'Only the copy that owns the provider lease may touch the grid icons (#194)
        OwnsRuntime = M_Lease_IsOwner()

'------------------------------------------------------------------------------
' CLEAN UP ONLY WHEN SETTINGS ARE UNCHANGED
'------------------------------------------------------------------------------
    'Remove stale grid icon and exit when no persisted setting changed
        If Not GridIconChanged Then
            If Not KeyboardChanged Then
                'Leave the grid icons to the owning copy and report why
                    If Not OwnsRuntime Then
                        HandlerStep = "Report setting not applied"
                        M_Lease_ReportSettingNotApplied PROC_NAME
                        Exit Sub
                    End If
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
' DEFER TO THE OWNING COPY
'------------------------------------------------------------------------------
    'Leave the grid icons and binding to the owning copy and report why. The
    'saved value is applied by that copy, at the latest when DP_Start reads it
        If Not OwnsRuntime Then
            HandlerStep = "Report setting not applied"
            M_Lease_ReportSettingNotApplied PROC_NAME
            Exit Sub
        End If

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

Public Function M_Settings_GetFirstDayOfWeek_Impl() As Long

'
'------------------------------------------------------------------------------
'                           GET FIRST DAY OF WEEK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetFirstDayOfWeek
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeek
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeek
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeek
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeek. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   gDP_FirstDayOfWeek
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetFirstDayOfWeek_Impl = gDP_FirstDayOfWeek

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

Public Function M_Settings_GetUseLocalNames_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                           GET USE LOCAL NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetUseLocalNames
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetUseLocalNames
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetUseLocalNames
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetUseLocalNames
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetUseLocalNames. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_UseLocalNames
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetUseLocalNames_Impl = gDP_UseLocalNames

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

Public Function M_Settings_GetAllowOutsideMonthSelection_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                   GET ALLOW OUTSIDE MONTH SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetAllowOutsideMonthSelection
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetAllowOutsideMonthSelection
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetAllowOutsideMonthSelection
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetAllowOutsideMonthSelection
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetAllowOutsideMonthSelection. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_AllowOutsideMonthSelection
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetAllowOutsideMonthSelection_Impl = gDP_AllowOutsideMonthSelection

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

Public Function M_Settings_GetFirstDayOfWeekText_Impl() As String

'
'------------------------------------------------------------------------------
'                       GET FIRST DAY OF WEEK TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetFirstDayOfWeekText
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeekText
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeekText
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeekText
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetFirstDayOfWeekText. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidFirstDayOfWeek
'   M_Settings_FirstDayOfWeekToText
'   gDP_FirstDayOfWeek
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetFirstDayOfWeekText_Impl = _
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

Public Function M_Settings_GetShowRightClick_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                           GET SHOW RIGHT-CLICK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetShowRightClick
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetShowRightClick
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetShowRightClick
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetShowRightClick
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetShowRightClick. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_ShowRightClick
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetShowRightClick_Impl = gDP_ShowRightClick

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

Public Function M_Settings_GetShowGridIcon_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                           GET SHOW GRID ICON
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetShowGridIcon
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetShowGridIcon
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetShowGridIcon
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetShowGridIcon
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetShowGridIcon. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_ShowGridIcon
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetShowGridIcon_Impl = gDP_ShowGridIcon

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

Public Function M_Settings_GetEnableKeyboardShortcut_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                       GET ENABLE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetEnableKeyboardShortcut
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetEnableKeyboardShortcut
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetEnableKeyboardShortcut
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetEnableKeyboardShortcut
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetEnableKeyboardShortcut. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_EnableKeyboardShortcut
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetEnableKeyboardShortcut_Impl = gDP_EnableKeyboardShortcut

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

Public Function M_Settings_GetUseWinAPI_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                           GET USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetUseWinAPI
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetUseWinAPI
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetUseWinAPI
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetUseWinAPI
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetUseWinAPI. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_UseWinAPI
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetUseWinAPI_Impl = gDP_UseWinAPI

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

Public Function M_Settings_GetClockMode_Impl() As DP_ClockMode

'
'------------------------------------------------------------------------------
'                           GET CLOCK MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetClockMode
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetClockMode
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetClockMode
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetClockMode
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetClockMode. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidClockMode
'   gDP_ClockMode
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetClockMode_Impl = gDP_ClockMode

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

Public Function M_Settings_GetSizeMode_Impl() As DP_SizeMode

'
'------------------------------------------------------------------------------
'                           GET SIZE MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetSizeMode
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetSizeMode
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetSizeMode
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetSizeMode
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetSizeMode. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_Settings_IsValidSizeMode
'   gDP_SizeMode
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetSizeMode_Impl = gDP_SizeMode

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

Public Function M_Settings_GetHighlightWeekends_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                           GET HIGHLIGHT WEEKENDS
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetHighlightWeekends
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetHighlightWeekends
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetHighlightWeekends
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetHighlightWeekends
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetHighlightWeekends. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_HighlightWeekends
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetHighlightWeekends_Impl = gDP_HighlightWeekends

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

Public Function M_Settings_GetCloseAfterSelection_Impl() As Boolean

'
'------------------------------------------------------------------------------
'                         GET CLOSE AFTER SELECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetCloseAfterSelection
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetCloseAfterSelection
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetCloseAfterSelection
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetCloseAfterSelection
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetCloseAfterSelection. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_CloseAfterSelection
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetCloseAfterSelection_Impl = gDP_CloseAfterSelection

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

Public Function M_Settings_GetHolidayCallback_Impl() As String

'
'------------------------------------------------------------------------------
'                           GET HOLIDAY CALLBACK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_GetHolidayCallback
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_GetHolidayCallback
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_GetHolidayCallback
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_GetHolidayCallback
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_GetHolidayCallback. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   gDP_HolidayCallbackName
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        M_Settings_GetHolidayCallback_Impl = VBA.Trim$(gDP_HolidayCallbackName)

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

Public Sub M_Picker_EnsureManager()

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
'   None. The parameterless signature deliberately preserves eligibility for
'   Alt+F8, assigned controls and other Excel macro entry points.
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Ensures persisted settings are loaded, preserves the caller's Excel event
'   state, creates the manager when missing, and recreates it when the existing
'   manager is not hooked
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
'   2026-09-03 - Restored a parameterless Excel macro entry point.
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
' VERIFY RUNTIME ADMISSION
'------------------------------------------------------------------------------
    'Prove ownership before any side effect. This is the backstop for direct
    'calls to this technically public bootstrapper: supported entry paths admit
    'themselves first and exit quietly, so reaching here unadmitted means the
    'guard was bypassed rather than a user action being refused
        If Not M_Lease_TryAcquire() Then
            Err.Raise vbObjectError + 515, PROC_NAME, _
                "DatePicker manager admission refused: another provider owns " & _
                "this Excel session, or ownership could not be verified"
        End If

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

Public Sub DP_Preload_Impl()

'
'------------------------------------------------------------------------------
'                           PRELOAD DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Preload
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Preload
'
' RETURNS
'   As documented on M_DatePicker.DP_Preload
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Preload
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Preload. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Picker_EnsureManager
'   M_FormBridge_GetLoadedForm
'   M_Timer_Stop
'   UF_DatePicker
'   DP_FORM_NAME
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
' ADMIT THIS PROVIDER
'------------------------------------------------------------------------------
    'Prove ownership before loading a hidden form. Preload is a background
    'optimization, so a refused copy declines silently rather than interrupting
    'workbook open with a second message box after DP_Start already reported
        If Not M_Lease_EnsureAdmitted(PROC_NAME, ReportToUser:=False) Then
            GoTo CleanExit
        End If

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

Public Sub DP_ForceReleaseProviderLease_Impl()

'
'------------------------------------------------------------------------------
'                     FORCE RELEASE PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_ForceReleaseProviderLease
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_ForceReleaseProviderLease
'
' RETURNS
'   As documented on M_DatePicker.DP_ForceReleaseProviderLease
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_ForceReleaseProviderLease
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_ForceReleaseProviderLease. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Lease_GetBar
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "DP_ForceReleaseProviderLease"

    Dim LeaseBar        As Object       'Lease command bar

'------------------------------------------------------------------------------
' FORCE RELEASE
'------------------------------------------------------------------------------
    'Never let a recovery action raise
        On Error Resume Next
    'Delete the lease whoever owns it
        Set LeaseBar = M_Lease_GetBar()
        If Not LeaseBar Is Nothing Then
            LeaseBar.Delete
            Debug.Print PROC_NAME & " | Provider lease released by operator request"
        Else
            Debug.Print PROC_NAME & " | No provider lease was present"
        End If
    'Drop this project's claim as well
        mDP_RuntimeOwnerId = VBA.vbNullString
    'Release object references
        Set LeaseBar = Nothing
    'Clear any suppressed release error
        Err.Clear

End Sub

Public Sub M_Lease_Test_ClearOwnerToken()

'
'==============================================================================
'                     CLEAR LOCAL OWNER TOKEN (TEST)
'==============================================================================
'   Drops this project's lease token without touching the lease itself.
'
'   THIS IS INTERNAL TEST INFRASTRUCTURE. It is not supported DatePicker API and
'   is classified internal in src/public-surface.json.
'
'   It exists because a regression suite cannot load a second VBA project. This
'   reproduces exactly the state a second provider sees, and the state a VBA
'   project reset leaves behind:
'
'       the lease survives, the local token does not
'
'   It never deletes, creates or rewrites a lease. Clearing a token can only ever
'   make this project less privileged, so a stray call cannot seize ownership or
'   destroy another provider's registrations.
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR
'------------------------------------------------------------------------------
    'Drop this project's claim, leaving the lease untouched
        mDP_RuntimeOwnerId = VBA.vbNullString

End Sub

Public Sub M_Lease_Test_SilenceRefusalReport(ByVal Silence As Boolean)

'
'==============================================================================
'                   SILENCE THE REFUSAL REPORT (TEST)
'==============================================================================
'   Suppresses the operator-facing refusal message box and counts refusals
'   instead.
'
'   THIS IS INTERNAL TEST INFRASTRUCTURE. It is not supported DatePicker API and
'   is classified internal in src/public-surface.json.
'
'   It exists because #37 requires automated coverage of the real entry points
'   under a refused lease. Those paths report refusal through a modal MsgBox,
'   and Application.DisplayAlerts does not suppress VBA.MsgBox, so an unattended
'   regression run would block forever waiting for a click.
'
'   Silencing only affects reporting. Admission itself is unchanged: a refused
'   provider is still refused, and no shared state is mutated either way.
'
'   Enabling resets the counter so each suite measures its own refusals. A run
'   that leaves this True would hide genuine conflicts from the operator, so the
'   suite restores it in its exit path.
'==============================================================================

'------------------------------------------------------------------------------
' SET
'------------------------------------------------------------------------------
    'Apply the requested reporting mode
        mDP_LeaseRefusalSilenced = Silence
    'Reset the counter so each measurement window starts at zero
        mDP_LeaseRefusalCount = 0

End Sub

Public Function M_Lease_Test_RefusalReportCount() As Long

'
'==============================================================================
'                   READ THE REFUSAL COUNT (TEST)
'==============================================================================
'   Reports how many refusals were raised since reporting was silenced.
'
'   THIS IS INTERNAL TEST INFRASTRUCTURE. It is not supported DatePicker API and
'   is classified internal in src/public-surface.json.
'
'   A refusal that is silently skipped and a refusal that never happened look
'   identical from outside. This lets a suite prove the entry path actually
'   refused, rather than merely proving that nothing visible occurred.
'==============================================================================

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Report refusals observed while silenced
        M_Lease_Test_RefusalReportCount = mDP_LeaseRefusalCount

End Function

Private Sub M_Lease_ReportRefusal(ByVal EntryPoint As String)

'
'==============================================================================
'                          REPORT LEASE REFUSAL
'==============================================================================
'   Tells the operator why a DatePicker entry point declined to act.
'
'   A refused provider produces no visible change, so without this it would look
'   like the DatePicker simply failed to start.
'==============================================================================

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Never let reporting raise into a refusal
        On Error Resume Next
    'Record the refusal for diagnostics
        Debug.Print EntryPoint & " | Refused | Another DatePicker provider owns " & _
            "this Excel session, or ownership could not be verified"
    'Count and return without the modal when a regression run has silenced
    'reporting. An automated suite has to drive the real entry points, and a
    'modal message box would block the run waiting for a human
        If mDP_LeaseRefusalSilenced Then
            mDP_LeaseRefusalCount = mDP_LeaseRefusalCount + 1
            Err.Clear
            Exit Sub
        End If
    'Tell the operator, because nothing else in the session will
        MsgBox _
            "Another copy of the Date / Time Picker is already active in this Excel " & _
            "session, or its ownership could not be verified." & VBA.vbCrLf & VBA.vbCrLf & _
            "Running two copies at once is not supported: they share the same " & _
            "keyboard shortcut, right-click entry and grid icons, and either " & _
            "one's shutdown would remove the other's." & VBA.vbCrLf & VBA.vbCrLf & _
            "Close the other copy and restart Excel." & VBA.vbCrLf & VBA.vbCrLf & _
            "If you are certain no other copy is running - for example after " & _
            "resetting the VBA project - run:" & VBA.vbCrLf & VBA.vbCrLf & _
            "    DP_ForceReleaseProviderLease", _
            vbExclamation Or vbOKOnly, _
            DP_MSGBOX_TITLE
    'Clear any suppressed reporting error
        Err.Clear

End Sub

Private Sub M_Lease_ReportSettingNotApplied(ByVal EntryPoint As String)

'
'------------------------------------------------------------------------------
'                         REPORT SETTING NOT APPLIED
'------------------------------------------------------------------------------
' PURPOSE
'   Explains why a settings setter saved its value but left the shared Excel
'   registration it controls untouched
'
' WHY THIS EXISTS
'   The keyboard shortcut, right-click and grid-icon setters change
'   application-wide Excel state. Only the copy that owns the provider lease
'   may do that, so a setter in any other copy saves the value and stops (#194).
'   Whoever called it still needs to know whether that was a conflict with
'   another copy or a runtime that has not been started yet
'
' INPUTS
'   EntryPoint
'     Name of the calling setter, for the report and diagnostics
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Reads the lease. When a lease exists - held by another copy, or unreadable -
'   reports through M_Lease_ReportRefusal, exactly as a refused entry point
'   does. When no lease exists, writes one line to the Immediate Window and
'   shows nothing
'
' ERROR POLICY
'   Does not raise. The setting has already been saved, and reporting must not
'   turn that into a failure
'
' DEPENDENCIES
'   M_Lease_ReadOwner
'   M_Lease_ReportRefusal
'
' NOTES
'   Call it only after M_Lease_IsOwner returned False. It never acquires,
'   reclaims or releases a lease: a setter is not a start path
'
'   No lease is the ordinary state before DP_Start, for example in a host that
'   sets its integration preferences in Workbook_Open and then starts the
'   runtime. That is not a conflict, so it shows no message box. DP_Start reads
'   the saved value and applies it
'
'   A copy whose lease bar vanished while it held its token (#113) also sees no
'   lease here. The setter does not reclaim it; the right-click re-ensure and
'   DP_Stop do
'
' UPDATED
'   2026-11-14
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LeaseOwner      As String       'Lease owner as read now

'------------------------------------------------------------------------------
' CLASSIFY
'------------------------------------------------------------------------------
    'Never let reporting raise into a setter that has already saved its value
        On Error Resume Next
    'Read the lease. No bar is the only free answer; an unreadable lease is held
        LeaseOwner = M_Lease_ReadOwner()

'------------------------------------------------------------------------------
' REPORT
'------------------------------------------------------------------------------
    'Another copy holds the lease, or ownership cannot be verified: refuse
    'visibly, through the same report the entry points use
        If VBA.LenB(LeaseOwner) > 0 Then
            M_Lease_ReportRefusal EntryPoint
    'No copy runs the DatePicker yet: note it without interrupting the caller
        Else
            Debug.Print EntryPoint & " | Not applied | Setting saved; no copy " & _
                "holds the provider lease, so the copy that runs DP_Start " & _
                "applies it"
        End If
    'Clear any suppressed reporting error
        Err.Clear

End Sub

Private Function M_Lease_EnsureAdmitted( _
    ByVal EntryPoint As String, _
    Optional ByVal ReportToUser As Boolean = True) As Boolean

'
'==============================================================================
'                            ENSURE RUNTIME ADMISSION
'==============================================================================
' PURPOSE
'   Single admission gate for every runtime entry path. Proves this project owns
'   the provider lease before any manager creation, form load or application-wide
'   registration is attempted
'
' WHY THIS EXISTS
'   The lease primitive already refuses correctly, but at v1.2.0 only DP_Start
'   consulted it. DP_Show and DP_Preload reached M_Picker_EnsureManager without
'   proving ownership, so a refused second copy could create a manager, load
'   settings, register Application.OnKey, hook Application events and later
'   remove the true owner's registrations during its own teardown
'
' INPUTS
'   EntryPoint
'       Diagnostic name of the calling entry path
'
'   ReportToUser
'       True to show the operator-facing refusal message. False for best-effort
'       background paths that must stay silent
'
' RETURNS
'   True when this project owns the lease, including when it already did
'
'   False when another provider owns it, or ownership cannot be verified
'
' BEHAVIOR
'   Delegates to M_Lease_TryAcquire, which is idempotent for the current owner
'   and fails closed for a foreign or unverifiable lease. Reports refusal once
'   at the refusing entry point
'
' ERROR POLICY
'   Does not raise. Refusal is an ordinary answer the caller acts on
'
' DEPENDENCIES
'   M_Lease_TryAcquire
'   M_Lease_ReportRefusal
'
' NOTES
'   Refusal performs no shared-state mutation. M_Lease_TryAcquire only creates
'   the lease bar on the free-lease path, so a refused copy leaves the owner's
'   registrations and the caller's Application.EnableEvents state untouched
'
'   Delegating entry points must not call this a second time. DP_Click and
'   Ribbon_ShowPicker reach DP_Show, which is the shared boundary, so a single
'   refusal produces a single report
'
' UPDATED
'   2026-08-25
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Lease_EnsureAdmitted"

'------------------------------------------------------------------------------
' ADMIT
'------------------------------------------------------------------------------
    'Acquire the lease, or confirm this project already holds it
        If M_Lease_TryAcquire() Then
            M_Lease_EnsureAdmitted = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' REFUSE
'------------------------------------------------------------------------------
    'Refusal is fail-closed and leaves no shared state changed
        M_Lease_EnsureAdmitted = False
    'Tell the operator unless this is a silent background path
        If ReportToUser Then
            M_Lease_ReportRefusal EntryPoint
        Else
            'Record the silent refusal so diagnostics still show the cause
                On Error Resume Next
                Debug.Print EntryPoint & _
                    " | Refused | Provider lease not held; entry path declined"
                Err.Clear
        End If

End Function

Public Function M_Lease_NewOwnerId() As String

'
'------------------------------------------------------------------------------
'                          NEW RUNTIME OWNER ID
'------------------------------------------------------------------------------
' PURPOSE
'   Generates the ephemeral token that identifies this provider while it holds
'   the runtime lease
'
' WHY THIS EXISTS
'   The lease has to distinguish "I own this" from "someone else owns this", and
'   the answer must not survive the VBA project that produced it
'
' INPUTS
'   None
'
' RETURNS
'   A token for this acquisition attempt
'
' BEHAVIOR
'   Combines the clock, the sub-second timer and a random component
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
'
' NOTES
'   The token deliberately carries no workbook identity. ThisWorkbook.Name and
'   FullName change on rename, move and save-as, and none of those should alter
'   ownership during a live Excel session
'
'   A GUID would be stronger, but CoCreateGuid is WinAPI and the lease must work
'   with WinAPI disabled by setting or by platform. Two providers starting within
'   the same millisecond would need the random component to collide as well,
'   and a collision fails closed: the second provider reads a marker equal to its
'   own token, believes it already owns the lease, and stops. It does not seize
'   anything
'
'   Public only so M_DP_GridIcon can mint its icon owner token from the same
'   generator; Option Private Module keeps it inside the project (#57)
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' GENERATE
'------------------------------------------------------------------------------
    'Never let token generation raise
        On Error Resume Next
    'Seed from the system timer so two projects do not share a sequence
        Randomize
    'Combine clock, sub-second timer and a random component
        M_Lease_NewOwnerId = VBA.Format$(VBA.Now, "yyyymmddhhnnss") & "-" & _
            VBA.Format$(VBA.Timer * 1000, "00000000") & "-" & _
            VBA.Hex$(VBA.Int(VBA.Rnd() * 2147483647#))
    'Clear any suppressed generation error
        Err.Clear

End Function

Public Function M_Lease_LocalOwnerId() As String

'
'------------------------------------------------------------------------------
'                          GET LOCAL OWNER ID
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the lease token this project holds, or an empty string
'
' WHY THIS EXISTS
'   M_DP_GridIcon stamps its icons with the live lease token when there is
'   one. Reading it through this function keeps the lease state owned by
'   M_DP_Core (#57)
'
' INPUTS
'   None
'
' RETURNS
'   The local owner token, or vbNullString when this project holds no lease
'
' BEHAVIOR
'   Returns mDP_RuntimeOwnerId unchanged
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   mDP_RuntimeOwnerId
'
' NOTES
'   A non-empty token is the local record only; M_Lease_IsOwner is what proves
'   ownership against the lease marker
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the local owner token
        M_Lease_LocalOwnerId = mDP_RuntimeOwnerId

End Function

Private Function M_Lease_GetBar() As Object

'
'==============================================================================
'                            GET LEASE BAR
'==============================================================================
'   Returns the lease command bar, or Nothing when no bar of that name exists.
'
'   Resolving by name never creates anything: absence is an ordinary answer.
'==============================================================================

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Suppress the error raised when the bar does not exist
        On Error Resume Next
    'Set safe default result
        Set M_Lease_GetBar = Nothing
    'Attempt to resolve the lease bar
        Set M_Lease_GetBar = Excel.Application.CommandBars(DP_LEASE_BAR_NAME)
    'Clear any suppressed lookup error
        Err.Clear

End Function

Private Function M_Lease_ReadOwner() As String

'
'------------------------------------------------------------------------------
'                          READ LEASE OWNER
'------------------------------------------------------------------------------
' PURPOSE
'   Reports which provider currently holds the runtime lease
'
' WHY THIS EXISTS
'   Every ownership decision, on acquisition and on release, needs one answer to
'   this question, and it must distinguish "free" from "held by someone" from
'   "cannot tell"
'
' INPUTS
'   None
'
' RETURNS
'   An empty string when no lease exists
'
'   The owner token when exactly one well-formed marker is present
'
'   DP_LEASE_AMBIGUOUS when a bar exists whose marker is missing, malformed or
'   duplicated
'
' BEHAVIOR
'   Scans the lease bar for controls carrying the marker tag and reports what it
'   finds
'
' ERROR POLICY
'   Does not raise. An unreadable lease is reported as ambiguous, never as free
'
' DEPENDENCIES
'   M_Lease_GetBar
'
' NOTES
'   Ambiguous is not the same as free, and the difference is the whole safety
'   property. A bar this component cannot interpret belongs to something, and
'   treating it as garbage to clear would be exactly the destruction #37 exists
'   to prevent
'
'   Two markers mean two writers, which is unverifiable rather than resolvable
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim LeaseBar        As Object       'Resolved lease command bar
    Dim Ctl             As Object       'Current control while scanning
    Dim FoundToken      As String       'Token read from the marker
    Dim FoundCount      As Long         'Markers found

'------------------------------------------------------------------------------
' READ
'------------------------------------------------------------------------------
    'Never let a lease read raise into a caller
        On Error Resume Next
    'Set safe default result
        M_Lease_ReadOwner = VBA.vbNullString
    'Resolve the lease bar
        Set LeaseBar = M_Lease_GetBar()
    'No bar means no lease, which is the only "free" answer
        If LeaseBar Is Nothing Then
            Err.Clear
            Exit Function
        End If
    'Count the marker controls and capture the token
        For Each Ctl In LeaseBar.Controls
            If VBA.StrComp(Ctl.Tag, DP_LEASE_MARKER_TAG, vbBinaryCompare) = 0 Then
                FoundCount = FoundCount + 1
                FoundToken = Ctl.Parameter
            End If
        Next Ctl
        Set Ctl = Nothing
    'A bar exists, so something owns it. Anything but one readable marker is
    'unverifiable rather than free
        If Err.Number <> 0 Then
            M_Lease_ReadOwner = DP_LEASE_AMBIGUOUS
        ElseIf FoundCount <> 1 Then
            M_Lease_ReadOwner = DP_LEASE_AMBIGUOUS
        ElseIf VBA.LenB(FoundToken) = 0 Then
            M_Lease_ReadOwner = DP_LEASE_AMBIGUOUS
        Else
            M_Lease_ReadOwner = FoundToken
        End If
    'Release object references
        Set LeaseBar = Nothing
    'Clear any suppressed read error
        Err.Clear

End Function

Public Function M_Lease_TryReclaimVanished(ByVal EntryPoint As String) As Boolean

'
'------------------------------------------------------------------------------
'                       RECLAIM A VANISHED PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores ownership for a project whose lease bar disappeared while it still
'   held its owner token
'
' WHY THIS EXISTS
'   The lease is a temporary command bar. Loaded from Excel's alternate startup
'   folder, the add-in started and acquired it, yet the bar was gone shortly
'   after, together with the right-click entry, while the token remained (#113).
'   Without this, the owner could neither restore its menu nor stop cleanly
'
' INPUTS
'   EntryPoint
'     Name of the calling entry point, for diagnostics
'
' RETURNS
'   True when the lease was free and is now held by this project
'
' BEHAVIOR
'   Does nothing unless this project holds a token and no lease exists at all.
'   It then acquires the free lease through the silent admission path. A lease
'   held by another provider, or unreadable, is never taken
'
' ERROR POLICY
'   Never raises. Returns False on any failure
'
' DEPENDENCIES
'   mDP_RuntimeOwnerId
'   M_Lease_ReadOwner
'   M_Lease_EnsureAdmitted
'
' NOTES
'   This is the acquisition any copy may make of a free lease, restricted to
'   the copy that last held it. A copy that never held a token never claims
'   one here, so a right-click or a shutdown cannot start a copy that was never
'   started
'
'   Public only so M_DP_ContextMenu can reclaim at right-click; Option
'   Private Module keeps it inside the project (#57)
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECIDE
'------------------------------------------------------------------------------
    'Never let a reclaim raise into a caller
        On Error Resume Next
    'Set safe default result
        M_Lease_TryReclaimVanished = False
    'Only a project that held the lease may reclaim it
        If VBA.LenB(mDP_RuntimeOwnerId) = 0 Then Exit Function
    'Only a lease that is absent is free; held or unreadable leases are not
        If VBA.LenB(M_Lease_ReadOwner()) > 0 Then
            Err.Clear
            Exit Function
        End If

'------------------------------------------------------------------------------
' RECLAIM
'------------------------------------------------------------------------------
    'Acquire the free lease without any user-facing report
        If M_Lease_EnsureAdmitted(EntryPoint, False) Then
            M_Lease_TryReclaimVanished = True
            Debug.Print EntryPoint & " | Reclaimed a vanished provider lease"
        End If
    'Clear any suppressed error
        Err.Clear

End Function

Public Function M_Lease_IsOwner() As Boolean

'
'==============================================================================
'                              IS LEASE OWNER
'==============================================================================
'   Reports whether this VBA project currently holds the runtime lease.
'
'   Both halves must agree: this project must hold a token, and the lease must
'   still carry that same token. A project reset clears the first, and another
'   provider's lease fails the second.
'==============================================================================

'------------------------------------------------------------------------------
' COMPARE
'------------------------------------------------------------------------------
    'Never let an ownership check raise into a caller
        On Error Resume Next
    'Set safe default result
        M_Lease_IsOwner = False
    'A project holding no token cannot be the owner
        If VBA.LenB(mDP_RuntimeOwnerId) = 0 Then
            Err.Clear
            Exit Function
        End If
    'The lease must still carry this project's token
        M_Lease_IsOwner = (VBA.StrComp(M_Lease_ReadOwner(), mDP_RuntimeOwnerId, _
            vbBinaryCompare) = 0)
    'Clear any suppressed comparison error
        Err.Clear

End Function

Public Function M_Lease_TryAcquire() As Boolean

'
'------------------------------------------------------------------------------
'                          ACQUIRE PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Claims the one-provider runtime lease, or reports that another provider holds
'   it
'
' WHY THIS EXISTS
'   Two DatePicker copies in one Excel process register the same application-wide
'   resources under the same fixed identifiers, and either one's teardown removes
'   the other's. Detection has to happen before the first shared registration
'
' INPUTS
'   None
'
' RETURNS
'   True when this project owns the lease, including when it already did
'
'   False when another provider owns it, or when ownership cannot be verified
'
' BEHAVIOR
'   Reads the lease, and when it is free creates the bar, writes the marker, then
'   re-reads to confirm the marker is this project's before claiming ownership
'
' ERROR POLICY
'   Does not raise. Refusal is an ordinary answer that the caller reports
'
'   Fails closed: an unverifiable lease refuses rather than assuming it is free
'
' DEPENDENCIES
'   M_Lease_ReadOwner
'   M_Lease_GetBar
'   M_Lease_NewOwnerId
'
' NOTES
'   The bar and its control are Temporary, so Excel deletes them when it closes.
'   The lease therefore cannot survive into the next Excel process, which is what
'   makes a registry-backed lease the wrong shape: that one would outlive the
'   process and block startup permanently
'
'   A VBA project reset clears mDP_RuntimeOwnerId while the bar survives, so the
'   former owner can no longer prove ownership and startup refuses. That is the
'   specified policy: fail closed, and restart Excel to recover. Automatic
'   reclamation of a stale lease belongs to #14
'
'   The bar name is fixed rather than carrying the token, which gives the object
'   model a uniqueness point during acquisition. Owner identity lives on the
'   control
'
'   Creation is verified by re-reading. If something appeared between the lookup
'   and the create, the re-read classifies it rather than the code retrying
'
'   A CommandBar is not indestructible: other VBA in the process can delete
'   custom bars, and published examples do exactly that while sweeping invisible
'   ones. That is external tampering. What this component guarantees is that it
'   never deletes a lease it cannot prove it owns
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_Lease_TryAcquire"

    Dim CurrentOwner    As String       'Owner token currently on the lease
    Dim Candidate       As String       'Token this attempt would claim
    Dim LeaseBar        As Object       'Lease command bar
    Dim Marker          As Object       'Marker control
    Dim CreatedBar      As Boolean      'True when this attempt created the bar

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let acquisition raise into startup
        On Error Resume Next
    'Set safe default result
        M_Lease_TryAcquire = False

'------------------------------------------------------------------------------
' CLASSIFY THE EXISTING LEASE
'------------------------------------------------------------------------------
    'Read what currently holds the lease
        CurrentOwner = M_Lease_ReadOwner()
    'An existing lease carrying this project's token is an idempotent success
        If VBA.LenB(mDP_RuntimeOwnerId) > 0 Then
            If VBA.StrComp(CurrentOwner, mDP_RuntimeOwnerId, vbBinaryCompare) = 0 Then
                M_Lease_TryAcquire = True
                Err.Clear
                Exit Function
            End If
        End If
    'A lease held by anyone else refuses this provider
        If VBA.LenB(CurrentOwner) > 0 Then
            Debug.Print PROC_NAME & _
                " | Refused | Lease held by " & _
                VBA.IIf(CurrentOwner = DP_LEASE_AMBIGUOUS, _
                    "an unverifiable owner", "another provider")
            Err.Clear
            Exit Function
        End If

'------------------------------------------------------------------------------
' CREATE THE LEASE
'------------------------------------------------------------------------------
    'Generate the token this attempt will claim
        Candidate = M_Lease_NewOwnerId()
    'Create a temporary, hidden, application-wide bar
        Err.Clear
        Set LeaseBar = Excel.Application.CommandBars.Add( _
            Name:=DP_LEASE_BAR_NAME, Temporary:=True)
    'Something may have created the lease between the read and this call
        If Err.Number <> 0 Or LeaseBar Is Nothing Then
            Err.Clear
            Debug.Print PROC_NAME & " | Refused | Lease appeared during acquisition"
            Set LeaseBar = Nothing
            Exit Function
        End If
        CreatedBar = True
    'Keep the lease invisible. A new custom bar is hidden by default; this is
    'explicit so a later change cannot surface it
        LeaseBar.Visible = False

'------------------------------------------------------------------------------
' WRITE THE MARKER
'------------------------------------------------------------------------------
    'Add one hidden temporary control to carry ownership
        Set Marker = LeaseBar.Controls.Add(Temporary:=True)
        If Not Marker Is Nothing Then
            Marker.Tag = DP_LEASE_MARKER_TAG
            Marker.Parameter = Candidate
            Marker.Visible = False
        End If
        Set Marker = Nothing

'------------------------------------------------------------------------------
' VERIFY BEFORE CLAIMING
'------------------------------------------------------------------------------
    'Re-read rather than trusting the writes above
        Err.Clear
        If VBA.StrComp(M_Lease_ReadOwner(), Candidate, vbBinaryCompare) = 0 Then
            mDP_RuntimeOwnerId = Candidate
            M_Lease_TryAcquire = True
        Else
            'Remove only the bar this attempt created, and only when it is still
            'the object this attempt created
                If CreatedBar Then
                    If Not LeaseBar Is Nothing Then
                        LeaseBar.Delete
                    End If
                End If
                mDP_RuntimeOwnerId = VBA.vbNullString
                Debug.Print PROC_NAME & " | Refused | Lease could not be verified after creation"
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set LeaseBar = Nothing
    'Clear any suppressed acquisition error
        Err.Clear

End Function

Private Sub M_Lifecycle_ResetObservation()

'
'------------------------------------------------------------------------------
'                        RESET LIFECYCLE OBSERVATION
'------------------------------------------------------------------------------
' PURPOSE
'   Clears every field describing the previous lifecycle operation
'
' WHY THIS EXISTS
'   The observation fields are module-level and outlive the call that wrote them.
'   A later operation that failed before reaching a given step would otherwise be
'   read through the previous operation's values, which is worse than having none
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Zeroes the trace, the cleanup detail and count, the primary failure, the
'   operation name, and every lease and success flag
'
' ERROR POLICY
'   Cannot raise. Assignment to module-level scalars only
'
' DEPENDENCIES
'   The mDP_Lifecycle* observation fields
'
' NOTES
'   Called at the top of DP_Start, DP_Stop and DP_RepairRuntime, before anything
'   those routines do can fail. Every entry point must reset before it observes,
'   or a refused call reports the last successful one
'
'   This clears observation only. It never touches the provider lease, the
'   manager, the timer or any Excel state
'
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    mDP_LifecycleLastTrace = VBA.vbNullString
    mDP_LifecycleLastCleanupDetail = VBA.vbNullString
    mDP_LifecycleLastCriticalClean = False
    mDP_LifecycleLastLeaseReleased = False
    mDP_LifecycleLastPrimaryNumber = 0
    mDP_LifecycleLastPrimaryStep = VBA.vbNullString
    mDP_LifecycleLastPrimaryDescription = VBA.vbNullString

    mDP_LifecycleLastOperation = VBA.vbNullString
    mDP_LifecycleLastSucceeded = False
    mDP_LifecycleLastCleanupAttempted = False
    mDP_LifecycleLastCleanupFailureCount = 0
    mDP_LifecycleLastLeaseWasAlreadyOwned = False
    mDP_LifecycleLastLeaseAcquiredThisCall = False
End Sub

Private Sub M_Lifecycle_SetPrimaryFailure( _
    ByVal ErrorNumber As Long, _
    ByVal StepName As String, _
    ByVal ErrorDescription As String)

'
'------------------------------------------------------------------------------
'                        RECORD THE PRIMARY FAILURE
'------------------------------------------------------------------------------
' PURPOSE
'   Records the failure that caused a lifecycle operation to stop
'
' WHY THIS EXISTS
'   Cleanup runs after the primary failure and can fail in its own right. Keeping
'   the primary cause in dedicated fields is what stops a cleanup failure from
'   overwriting the reason the operation stopped
'
' INPUTS
'   ErrorNumber
'     Number of the primary failure
'
'   StepName
'     Transaction step that produced it
'
'   ErrorDescription
'     Description of the primary failure
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the three values in the primary-failure observation fields
'
' ERROR POLICY
'   Cannot raise. Assignment to module-level scalars only
'
' DEPENDENCIES
'   mDP_LifecycleLastPrimaryNumber
'   mDP_LifecycleLastPrimaryStep
'   mDP_LifecycleLastPrimaryDescription
'
' NOTES
'   The caller captures Err before any cleanup runs and passes the captured
'   values here. Reading the live Err object after cleanup would report whatever
'   the last suppressed cleanup operation left behind
'
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    mDP_LifecycleLastPrimaryNumber = ErrorNumber
    mDP_LifecycleLastPrimaryStep = StepName
    mDP_LifecycleLastPrimaryDescription = ErrorDescription

End Sub

Private Function M_Lifecycle_TryConsumeFault( _
    ByVal StepName As String, _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                       CONSUME AN INJECTED LIFECYCLE FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a fault is armed for the named transaction step, and consumes
'   it when it is
'
' WHY THIS EXISTS
'   Startup and shutdown have to be proven to roll back from a failure at every
'   boundary, and none of those boundaries can be made to fail on demand by any
'   ordinary input. Without an injection seam the transaction logic would only
'   ever execute during a real defect
'
' INPUTS
'   StepName
'     Boundary being entered, for example Cleanup.Timer or Lease.Delete
'
'   ErrorNumber
'     Receives the armed error number when the fault fires
'
'   ErrorDescription
'     Receives a description naming the step
'
' RETURNS
'   True when a fault was armed for this step and has now been consumed
'
' BEHAVIOR
'   Matches the armed step name exactly, hands back the armed number, and disarms
'
' ERROR POLICY
'   Cannot raise. Comparison and assignment only
'
' DEPENDENCIES
'   mDP_LifecycleFaultStep
'   mDP_LifecycleFaultNumber
'
' NOTES
'   One-shot by construction: the armed step and number are cleared as the fault
'   is handed back, so an armed fault can never affect a second boundary or leak
'   into a later suite
'
'   The comparison is binary, so a step name that differs only in case does not
'   match. Callers pass literals that must equal the literals the tests arm
'
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString
    M_Lifecycle_TryConsumeFault = False

    If mDP_LifecycleFaultNumber = 0 Then Exit Function
    If VBA.StrComp(mDP_LifecycleFaultStep, StepName, vbBinaryCompare) <> 0 Then Exit Function

    ErrorNumber = mDP_LifecycleFaultNumber
    ErrorDescription = "Injected lifecycle fault at " & StepName
    mDP_LifecycleFaultStep = VBA.vbNullString
    mDP_LifecycleFaultNumber = 0
    M_Lifecycle_TryConsumeFault = True

End Function

Private Sub M_Lifecycle_RaiseIfFault(ByVal StepName As String)

'
'------------------------------------------------------------------------------
'                        RAISE AN INJECTED LIFECYCLE FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Raises an armed fault outward, for boundaries that report failure by raising
'
' WHY THIS EXISTS
'   The cleanup helpers report failure through a Boolean and ByRef error fields,
'   but the startup steps run under a live error handler and fail by raising.
'   Both shapes need the same seam, so the seam is offered in both forms
'
' INPUTS
'   StepName
'     Boundary being entered
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Consumes a fault armed for the step and raises it. Returns silently when no
'   fault is armed for that step
'
' ERROR POLICY
'   Deliberately raises outward when a fault fires. The caller's own handler
'   decides what that means for the transaction
'
' DEPENDENCIES
'   M_Lifecycle_TryConsumeFault
'
' NOTES
'   The raised Source is M_Lifecycle_TestFault rather than the calling routine,
'   so an injected failure is never mistaken for a real one in a trace
'
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim ErrorNumber As Long
    Dim ErrorDescription As String

    If M_Lifecycle_TryConsumeFault(StepName, ErrorNumber, ErrorDescription) Then
        Err.Raise ErrorNumber, "M_Lifecycle_TestFault", ErrorDescription
    End If

End Sub

Private Sub M_Lifecycle_RecordCleanupStep( _
    ByVal StepName As String, _
    ByVal Succeeded As Boolean, _
    ByVal ErrorNumber As Long, _
    ByVal ErrorDescription As String, _
    Optional ByVal CountFailure As Boolean = True)

'
'------------------------------------------------------------------------------
'                          RECORD ONE CLEANUP STEP
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one step outcome to the cleanup trace and accounts for its failure
'
' WHY THIS EXISTS
'   Teardown attempts every step even after one fails, so the outcome of a
'   shutdown is a sequence rather than a single verdict. Before this, a
'   suppressed cleanup failure left nothing behind at all and the operation still
'   reported success
'
' INPUTS
'   StepName
'     Cleanup boundary being recorded
'
'   Succeeded
'     True when the step completed
'
'   ErrorNumber
'     Number the step reported, or zero
'
'   ErrorDescription
'     Description the step reported
'
'   CountFailure
'     False to record the step in the trace without counting it as a failure
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Appends StepName=PASS or StepName=FAIL(number) to the trace, increments the
'   failure count unless the caller opted out, and keeps the first failure detail
'
' ERROR POLICY
'   Cannot raise. String building and assignment only
'
' DEPENDENCIES
'   mDP_LifecycleLastTrace
'   mDP_LifecycleLastCleanupDetail
'   mDP_LifecycleLastCleanupFailureCount
'
' NOTES
'   Only the first failure detail is kept. A later failure is visible in the
'   trace but does not overwrite the detail, so the earliest cause stays readable
'
'   CountFailure exists for steps that are recorded for completeness but are not
'   part of the critical set, so an optional step cannot make a clean teardown
'   look incomplete
'
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim StepText As String

    If Succeeded Then
        StepText = StepName & "=PASS"
    Else
        StepText = StepName & "=FAIL"
        If ErrorNumber <> 0 Then StepText = StepText & "(" & VBA.CStr(ErrorNumber) & ")"
    End If

    If VBA.LenB(mDP_LifecycleLastTrace) = 0 Then
        mDP_LifecycleLastTrace = StepText
    Else
        mDP_LifecycleLastTrace = mDP_LifecycleLastTrace & " > " & StepText
    End If

    If Not Succeeded Then
        If CountFailure Then
            mDP_LifecycleLastCleanupFailureCount = _
                mDP_LifecycleLastCleanupFailureCount + 1
        End If
        If VBA.LenB(mDP_LifecycleLastCleanupDetail) = 0 Then
            mDP_LifecycleLastCleanupDetail = StepName
            If ErrorNumber <> 0 Then
                mDP_LifecycleLastCleanupDetail = mDP_LifecycleLastCleanupDetail & _
                    " | Error=" & VBA.CStr(ErrorNumber)
            End If
            If VBA.LenB(ErrorDescription) > 0 Then
                mDP_LifecycleLastCleanupDetail = mDP_LifecycleLastCleanupDetail & _
                    " | " & ErrorDescription
            End If
        End If
    End If

End Sub

Public Sub M_Lifecycle_Test_ArmFault(ByVal StepName As String, ByVal ErrorNumber As Long)

'
'------------------------------------------------------------------------------
'                            ARM A LIFECYCLE FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a one-shot fault for one named transaction boundary
'
' WHY THIS EXISTS
'   The regression matrix has to fail startup and shutdown at every boundary,
'   and none of them can be made to fail by any ordinary input
'
' INPUTS
'   StepName
'     Boundary to fail, for example Cleanup.Timer or Start.AfterManager
'
'   ErrorNumber
'     Error number the boundary reports; zero disarms
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns module-level state only
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. It takes an
'   argument, so it does not appear in the macro dialog, and #25 classifies it as
'   internal rather than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    mDP_LifecycleFaultStep = VBA.Trim$(StepName)
    mDP_LifecycleFaultNumber = ErrorNumber

End Sub

Public Sub M_Lifecycle_Test_Reset(ByVal ResetState As Boolean)

'
'------------------------------------------------------------------------------
'                          RESET LIFECYCLE TEST STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Disarms any armed fault and clears the observation fields
'
' WHY THIS EXISTS
'   A fault or an observation left behind would be read by the next suite as
'   though it belonged to that suite's own operation
'
' INPUTS
'   ResetState
'     False makes the call a no-op, so a suite can reset conditionally without
'     branching at the call site
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Assigns module-level state only
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. It takes an
'   argument, so it does not appear in the macro dialog, and #25 classifies it as
'   internal rather than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    If Not ResetState Then Exit Sub

    mDP_LifecycleFaultStep = VBA.vbNullString
    mDP_LifecycleFaultNumber = 0
    M_Lifecycle_ResetObservation

End Sub

Public Function M_Lifecycle_Test_LastTrace() As String

'
'------------------------------------------------------------------------------
'                               READ LAST TRACE
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the step-by-step trace of the last operation
'
' WHY THIS EXISTS
'   Every cleanup step appends PASS or FAIL(number) here, so a shutdown that
'   failed halfway is readable as a sequence rather than a single verdict
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastTrace = mDP_LifecycleLastTrace

End Function

Public Function M_Lifecycle_Test_LastCleanupDetail() As String

'
'------------------------------------------------------------------------------
'                           READ LAST CLEANUP DETAIL
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the first cleanup failure detail of the last operation
'
' WHY THIS EXISTS
'   The first failure is the actionable one. Later failures stay visible in the
'   trace without displacing it
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastCleanupDetail = mDP_LifecycleLastCleanupDetail

End Function

Public Function M_Lifecycle_Test_LastCriticalClean() As Boolean

'
'------------------------------------------------------------------------------
'                           READ LAST CRITICAL CLEAN
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether every critical cleanup step succeeded
'
' WHY THIS EXISTS
'   This is the conjunction that authorises releasing the provider lease, so a
'   test has to be able to read it independently of the operation result
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastCriticalClean = mDP_LifecycleLastCriticalClean

End Function

Public Function M_Lifecycle_Test_LastLeaseReleased() As Boolean

'
'------------------------------------------------------------------------------
'                           READ LAST LEASE RELEASED
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether the provider lease was actually released
'
' WHY THIS EXISTS
'   Release is conditional on critical cleanliness. Attempting it and achieving
'   it are different facts and are reported separately
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastLeaseReleased = mDP_LifecycleLastLeaseReleased

End Function

Public Function M_Lifecycle_Test_HasLocalOwnerToken() As Boolean

'
'------------------------------------------------------------------------------
'                          READ HAS LOCAL OWNER TOKEN
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether this project still holds its lease ownership token
'
' WHY THIS EXISTS
'   The #50 defect was discarding this token while the lease survived. A test
'   must be able to prove the token was retained when release failed
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_HasLocalOwnerToken = (VBA.LenB(mDP_RuntimeOwnerId) > 0)

End Function

Public Function M_Lifecycle_Test_LastPrimaryNumber() As Long

'
'------------------------------------------------------------------------------
'                           READ LAST PRIMARY NUMBER
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the error number that stopped the last operation
'
' WHY THIS EXISTS
'   Cleanup runs after the primary failure and can fail in its own right. The
'   primary cause is reported separately so cleanup cannot displace it
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastPrimaryNumber = mDP_LifecycleLastPrimaryNumber

End Function

Public Function M_Lifecycle_Test_LastPrimaryStep() As String

'
'------------------------------------------------------------------------------
'                            READ LAST PRIMARY STEP
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the transaction step that produced the primary failure
'
' WHY THIS EXISTS
'   Knowing a startup failed is not enough to prove rollback covered the right
'   steps; the boundary it failed at is what the matrix asserts against
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastPrimaryStep = mDP_LifecycleLastPrimaryStep

End Function

Public Function M_Lifecycle_Test_LastPrimaryDescription() As String

'
'------------------------------------------------------------------------------
'                        READ LAST PRIMARY DESCRIPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the description of the primary failure
'
' WHY THIS EXISTS
'   Preserving the description across cleanup is the same contract #48 imposed
'   on the write-back handlers
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastPrimaryDescription = mDP_LifecycleLastPrimaryDescription

End Function

Public Function M_Lifecycle_Test_LastOperation() As String

'
'------------------------------------------------------------------------------
'                             READ LAST OPERATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the name of the last lifecycle operation attempted
'
' WHY THIS EXISTS
'   Start, stop and repair share the observation fields, so a test reading them
'   must be able to confirm which operation wrote them
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastOperation = mDP_LifecycleLastOperation

End Function

Public Function M_Lifecycle_Test_LastSucceeded() As Boolean

'
'------------------------------------------------------------------------------
'                             READ LAST SUCCEEDED
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether the last lifecycle operation succeeded
'
' WHY THIS EXISTS
'   The operation result is deliberately distinct from critical cleanliness. A
'   refused call fails without any cleanup having been attempted at all
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastSucceeded = mDP_LifecycleLastSucceeded

End Function

Public Function M_Lifecycle_Test_LastCleanupAttempted() As Boolean

'
'------------------------------------------------------------------------------
'                         READ LAST CLEANUP ATTEMPTED
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a cleanup transaction ran at all
'
' WHY THIS EXISTS
'   A refused operation and a failed cleanup both report failure. Only this
'   distinguishes them, and they need different fixes
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastCleanupAttempted = mDP_LifecycleLastCleanupAttempted

End Function

Public Function M_Lifecycle_Test_LastCleanupFailureCount() As Long

'
'------------------------------------------------------------------------------
'                       READ LAST CLEANUP FAILURE COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Reports how many cleanup steps failed
'
' WHY THIS EXISTS
'   Teardown attempts every step, so the count is the difference between one
'   bad boundary and a comprehensively failed shutdown
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastCleanupFailureCount = mDP_LifecycleLastCleanupFailureCount

End Function

Public Function M_Lifecycle_Test_LastLeaseWasAlreadyOwned() As Boolean

'
'------------------------------------------------------------------------------
'                      READ LAST LEASE WAS ALREADY OWNED
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether the lease pre-existed the last operation
'
' WHY THIS EXISTS
'   A repeated start must never release a lease it did not take. This is the
'   fact that distinguishes a fresh start from a repeat
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastLeaseWasAlreadyOwned = mDP_LifecycleLastLeaseWasAlreadyOwned

End Function

Public Function M_Lifecycle_Test_LastLeaseAcquiredThisCall() As Boolean

'
'------------------------------------------------------------------------------
'                      READ LAST LEASE ACQUIRED THIS CALL
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether the last operation acquired the lease itself
'
' WHY THIS EXISTS
'   Rollback releases only a lease this call acquired. Asserting that requires
'   reading the classification the operation made, not inferring it
'
' INPUTS
'   None
'
' RETURNS
'   The recorded value
'
' BEHAVIOR
'   Reads one observation field. Free of side effects
'
' ERROR POLICY
'   Cannot raise
'
' DEPENDENCIES
'   The mDP_Lifecycle* fields
'
' NOTES
'   Public only because the regression harness is a separate module. Functions do
'   not appear in the macro dialog, and #25 classifies this as internal rather
'   than supported API
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    M_Lifecycle_Test_LastLeaseAcquiredThisCall = mDP_LifecycleLastLeaseAcquiredThisCall

End Function

Private Function M_Lifecycle_TryReleaseManager( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                        RELEASE THE MANAGER, VERIFIED
'------------------------------------------------------------------------------
' PURPOSE
'   Releases the global manager reference and proves it is gone
'
' WHY THIS EXISTS
'   Setting a reference to Nothing is not evidence that it is Nothing. Every step
'   in this transaction verifies its own outcome rather than assuming the
'   assignment took, because the lease is released only once the critical steps
'   are proven clean
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the failure when the reference survives or the release raises
'
' RETURNS
'   True when the manager reference is Nothing afterwards
'
' BEHAVIOR
'   Clears gDP_Manager and re-tests it
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   gDP_Manager
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    Set gDP_Manager = Nothing
    M_Lifecycle_TryReleaseManager = (gDP_Manager Is Nothing)
    If Not M_Lifecycle_TryReleaseManager Then
        ErrorNumber = vbObjectError + 2710
        ErrorDescription = "DatePicker manager reference remained live after release"
    End If
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Err.Clear
    M_Lifecycle_TryReleaseManager = False

End Function

Private Function M_Lifecycle_TryClosePickerForm( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                       CLOSE THE PICKER FORM, VERIFIED
'------------------------------------------------------------------------------
' PURPOSE
'   Unloads the picker form, clears the initial-date bridge, and proves no form
'   remains loaded
'
' WHY THIS EXISTS
'   Unload can leave an instance behind when initialization raised, and that
'   orphan is unreachable by name. Re-resolving after the unload is the only way
'   to tell an unload that worked from one that only appeared to
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the failure when a form is still loaded afterwards
'
' RETURNS
'   True when no form resolves after the unload
'
' BEHAVIOR
'   Hides and unloads any loaded form, clears the initial-date bridge state, then
'   re-resolves to confirm
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_FormBridge_GetLoadedForm
'   DP_FORM_NAME
'   gDP_InitialDate
'   gDP_HasInitialDate
'
' NOTES
'   The bridge state is cleared between the unload and the verification, so a
'   stale initial date cannot survive into the next runtime
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim LoadedForm As Object

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    If Not LoadedForm Is Nothing Then
        LoadedForm.Visible = False
        Unload LoadedForm
    End If
    Set LoadedForm = Nothing

    gDP_InitialDate = 0
    gDP_HasInitialDate = False

    Set LoadedForm = M_FormBridge_GetLoadedForm(DP_FORM_NAME)
    M_Lifecycle_TryClosePickerForm = (LoadedForm Is Nothing)
    If Not M_Lifecycle_TryClosePickerForm Then
        ErrorNumber = vbObjectError + 2713
        ErrorDescription = "A DatePicker form remained loaded after unload"
    End If
    Set LoadedForm = Nothing
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Set LoadedForm = Nothing
    Err.Clear
    M_Lifecycle_TryClosePickerForm = False

End Function

Private Function M_Lifecycle_TryRestoreEnableEvents( _
    ByVal DesiredValue As Boolean, _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                     RESTORE APPLICATION EVENTS, VERIFIED
'------------------------------------------------------------------------------
' PURPOSE
'   Restores Application.EnableEvents to a caller-chosen value and proves it took
'
' WHY THIS EXISTS
'   The component must leave the caller's event state as it found it. A business
'   macro that deliberately suppressed events has to still have them suppressed
'   when the DatePicker is done
'
' INPUTS
'   DesiredValue
'     State to restore, captured by the caller before the operation began
'
'   ErrorNumber, ErrorDescription
'     Receive the failure when the value does not read back
'
' RETURNS
'   True when Application.EnableEvents reads back as DesiredValue
'
' BEHAVIOR
'   Assigns the value and re-reads it
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   Excel.Application.EnableEvents
'
' NOTES
'   DP_RepairRuntime is the documented exception to caller preservation: it
'   forces events on deliberately, because a caller that left them off is the
'   condition repair exists to fix. That decision belongs to the caller, which is
'   why the desired value is a parameter here rather than a policy
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    Excel.Application.EnableEvents = DesiredValue
    If Excel.Application.EnableEvents <> DesiredValue Then
        ErrorNumber = vbObjectError + 2722
        ErrorDescription = "Application.EnableEvents did not restore to the caller state"
        Exit Function
    End If

    M_Lifecycle_TryRestoreEnableEvents = True
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Err.Clear
    M_Lifecycle_TryRestoreEnableEvents = False

End Function

Private Function M_Lifecycle_TryReleaseLease( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                     RELEASE THE PROVIDER LEASE, VERIFIED
'------------------------------------------------------------------------------
' PURPOSE
'   Releases the provider lease only when its deletion can be verified, and
'   retains the local ownership token whenever it cannot
'
' WHY THIS EXISTS
'   This is the defect #50 was filed for. Release cleared the local token whether
'   or not the lease bar was actually deleted, so a failed teardown surrendered
'   the proof of ownership while the lease survived. No project could then
'   release it, and every guarded entry point refused
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the failure when release cannot be verified
'
' RETURNS
'   True only when the lease is verified free, was already free, or this project
'   never held it
'
' BEHAVIOR
'   Treats a missing local token and an already-free lease as released. Refuses
'   when the current owner no longer matches the retained token. Otherwise
'   deletes the bar and re-reads the owner, clearing the local token only once
'   the lease reads as free
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Lease_ReadOwner
'   M_Lease_GetBar
'   mDP_RuntimeOwnerId
'
' NOTES
'   The ambiguous-owner refusal matters as much as the verification. A lease
'   whose owner token has changed belongs to someone else, and deleting it would
'   be the same class of defect as deleting an unowned grid-icon shape
'
'   Every failure path leaves mDP_RuntimeOwnerId intact, so a later retry still
'   has the proof it needs
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim CurrentOwner As String
    Dim LeaseBar As Object
    Dim FaultErrNumber As Long
    Dim FaultErrDescription As String

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    If VBA.LenB(mDP_RuntimeOwnerId) = 0 Then
        M_Lifecycle_TryReleaseLease = True
        Exit Function
    End If

    CurrentOwner = M_Lease_ReadOwner()
    If VBA.LenB(CurrentOwner) = 0 Then
        mDP_RuntimeOwnerId = VBA.vbNullString
        M_Lifecycle_TryReleaseLease = True
        Exit Function
    End If

    If VBA.StrComp(CurrentOwner, mDP_RuntimeOwnerId, vbBinaryCompare) <> 0 Then
        ErrorNumber = vbObjectError + 2717
        ErrorDescription = "Provider lease owner is ambiguous or no longer matches the retained local token"
        Exit Function
    End If

    If M_Lifecycle_TryConsumeFault("Lease.Delete", FaultErrNumber, FaultErrDescription) Then
        ErrorNumber = FaultErrNumber
        ErrorDescription = FaultErrDescription
        Exit Function
    End If

    Set LeaseBar = M_Lease_GetBar()
    If LeaseBar Is Nothing Then
        CurrentOwner = M_Lease_ReadOwner()
        If VBA.LenB(CurrentOwner) = 0 Then
            mDP_RuntimeOwnerId = VBA.vbNullString
            M_Lifecycle_TryReleaseLease = True
        Else
            ErrorNumber = vbObjectError + 2718
            ErrorDescription = "Provider lease could not be resolved for verified deletion"
        End If
        Exit Function
    End If

    LeaseBar.Delete
    Set LeaseBar = Nothing

    CurrentOwner = M_Lease_ReadOwner()
    If VBA.LenB(CurrentOwner) = 0 Then
        mDP_RuntimeOwnerId = VBA.vbNullString
        M_Lifecycle_TryReleaseLease = True
    Else
        ErrorNumber = vbObjectError + 2719
        ErrorDescription = "Provider lease deletion did not produce a verified free lease"
    End If
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Set LeaseBar = Nothing
    Err.Clear
    M_Lifecycle_TryReleaseLease = False

End Function

Private Function M_Lifecycle_Cleanup( _
    ByVal ReleaseLease As Boolean, _
    ByVal EntryPoint As String, _
    Optional ByVal RestoreApplicationEvents As Boolean = False, _
    Optional ByVal DesiredEnableEvents As Boolean = True) As Boolean

'
'------------------------------------------------------------------------------
'                       THE SHUTDOWN CLEANUP TRANSACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Runs every teardown boundary, records each outcome, and releases the provider
'   lease only when the critical steps are proven clean
'
' WHY THIS EXISTS
'   Teardown used to suppress its own errors and release the lease regardless, so
'   a shutdown that failed halfway reported success and freed ownership while
'   application-wide state survived. A new provider could then start on top of it
'
' INPUTS
'   ReleaseLease
'     True to attempt lease release once the critical steps are evaluated
'
'   EntryPoint
'     Operation this cleanup belongs to, for diagnostics
'
'   RestoreApplicationEvents
'     True to restore Application.EnableEvents as part of the transaction
'
'   DesiredEnableEvents
'     Value to restore when RestoreApplicationEvents is True
'
' RETURNS
'   True when every critical step succeeded and, where requested, the lease was
'   verifiably released
'
' BEHAVIOR
'   Resets the observation fields, then attempts manager, timer, form, grid,
'   context-menu and keyboard cleanup in that order. Every step is attempted
'   regardless of earlier failures and each outcome is appended to the trace.
'   Critical cleanliness is the conjunction of the critical step results, and the
'   lease is released only after that is known
'
' ERROR POLICY
'   Never raises outward. Each step reports through its own return value, and the
'   per-step fault seam is consulted before each one
'
' DEPENDENCIES
'   M_Lifecycle_ResetObservation
'   M_Lifecycle_TryConsumeFault
'   M_Lifecycle_RecordCleanupStep
'   the M_Lifecycle_Try* cleanup helpers
'
' NOTES
'   Attempting every step even after one fails is the point. Stopping at the
'   first failure would leave the remaining application-wide state registered
'   with no owner able to remove it
'
'   Order matters. The timer is stopped before the form is unloaded so a pending
'   tick cannot fire against a form that is going away, and the lease is last
'   because it is the ownership that authorises all of it
'
'   The keyboard step contributes to critical cleanliness but cannot verify its
'   own outcome; see M_KeyboardShortcut_TryRemove
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim StepSucceeded As Boolean
    Dim StepErrNumber As Long
    Dim StepErrDescription As String
    Dim CriticalClean As Boolean

    mDP_LifecycleLastTrace = VBA.vbNullString
    mDP_LifecycleLastCleanupDetail = VBA.vbNullString
    mDP_LifecycleLastCriticalClean = False
    mDP_LifecycleLastLeaseReleased = False
    mDP_LifecycleLastCleanupAttempted = True
    mDP_LifecycleLastCleanupFailureCount = 0
    CriticalClean = True

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.Manager", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_Lifecycle_TryReleaseManager(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "Manager", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.Timer", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_Timer_TryStop(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "Timer", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.Form", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_Lifecycle_TryClosePickerForm(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "Form", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.Grid", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_GridIcon_TryPurgeAll(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "Grid", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.ContextMenu", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_ContextMenu_TryRemove(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "ContextMenu", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    StepErrNumber = 0
    StepErrDescription = VBA.vbNullString
    If M_Lifecycle_TryConsumeFault("Cleanup.Keyboard", StepErrNumber, StepErrDescription) Then
        StepSucceeded = False
    Else
        StepSucceeded = M_KeyboardShortcut_TryRemove(StepErrNumber, StepErrDescription)
    End If
    M_Lifecycle_RecordCleanupStep "Keyboard", StepSucceeded, StepErrNumber, StepErrDescription
    CriticalClean = CriticalClean And StepSucceeded

    On Error Resume Next
    Err.Clear
    M_GetQualifiedMacroName_ClearCache
    StepErrNumber = Err.Number
    StepErrDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    StepSucceeded = (StepErrNumber = 0)
    M_Lifecycle_RecordCleanupStep "CallbackCache", StepSucceeded, StepErrNumber, StepErrDescription

    If RestoreApplicationEvents Then
        StepErrNumber = 0
        StepErrDescription = VBA.vbNullString
        If M_Lifecycle_TryConsumeFault("Cleanup.EnableEvents", StepErrNumber, StepErrDescription) Then
            StepSucceeded = False
        Else
            StepSucceeded = M_Lifecycle_TryRestoreEnableEvents( _
                DesiredEnableEvents, StepErrNumber, StepErrDescription)
        End If
        M_Lifecycle_RecordCleanupStep "EnableEvents", StepSucceeded, StepErrNumber, StepErrDescription
        CriticalClean = CriticalClean And StepSucceeded
    End If

    mDP_LifecycleLastCriticalClean = CriticalClean

    If ReleaseLease Then
        If CriticalClean Then
            StepErrNumber = 0
            StepErrDescription = VBA.vbNullString
            If M_Lifecycle_TryConsumeFault("Cleanup.Lease", StepErrNumber, StepErrDescription) Then
                StepSucceeded = False
            Else
                StepSucceeded = M_Lifecycle_TryReleaseLease(StepErrNumber, StepErrDescription)
            End If
            M_Lifecycle_RecordCleanupStep "Lease", StepSucceeded, StepErrNumber, StepErrDescription
            mDP_LifecycleLastLeaseReleased = StepSucceeded
        Else
            M_Lifecycle_RecordCleanupStep "Lease", False, _
                vbObjectError + 2720, "Lease retained because critical cleanup is incomplete", _
                CountFailure:=False
            mDP_LifecycleLastLeaseReleased = False
        End If
    Else
        If VBA.LenB(mDP_LifecycleLastTrace) > 0 Then
            mDP_LifecycleLastTrace = mDP_LifecycleLastTrace & " > Lease=RETAINED"
        Else
            mDP_LifecycleLastTrace = "Lease=RETAINED"
        End If
    End If

    M_Lifecycle_Cleanup = CriticalClean
    If ReleaseLease Then M_Lifecycle_Cleanup = (CriticalClean And mDP_LifecycleLastLeaseReleased)

    If Not M_Lifecycle_Cleanup Then
        Debug.Print EntryPoint & " | Incomplete cleanup | " & _
            mDP_LifecycleLastTrace & " | " & mDP_LifecycleLastCleanupDetail
    End If

End Function

Public Sub M_Lease_Release()

'
'------------------------------------------------------------------------------
'                          RELEASE THE PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Supported wrapper that attempts to release the provider lease
'
' WHY THIS EXISTS
'   The lease release is part of the shutdown transaction, but the operation is
'   also reachable on its own. This keeps the supported name while delegating the
'   verified release to the transaction helper
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to M_Lifecycle_TryReleaseLease and writes a diagnostic when the
'   lease is retained rather than released
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Lifecycle_TryReleaseLease
'
' NOTES
'   Retention is not failure. A lease that could not be verifiably deleted is
'   deliberately kept, along with the local ownership token, so a later retry
'   still has the proof it needs. Before #50 this cleared the token regardless,
'   which stranded the lease beyond any project's reach
'
'   The outcome is observable through M_Lifecycle_Test_HasLocalOwnerToken rather
'   than from this routine, which reports nothing to its caller
' UPDATED
'   2026-09-05
'------------------------------------------------------------------------------

    Dim ErrorNumber As Long
    Dim ErrorDescription As String

    If Not M_Lifecycle_TryReleaseLease(ErrorNumber, ErrorDescription) Then
        Debug.Print "M_Lease_Release | Retained | Error=" & _
            VBA.CStr(ErrorNumber) & " | " & ErrorDescription
    End If

End Sub

Public Sub DP_Start_Impl()

'
'------------------------------------------------------------------------------
'                               START DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Start
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Start
'
' RETURNS
'   As documented on M_DatePicker.DP_Start
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Start
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Start. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Lease_EnsureAdmitted
'   M_Picker_EnsureManager
'   M_ContextMenu_Update
'   M_KeyboardShortcut_Update
'   M_GridIcon_PreCreateHidden
'   M_Lifecycle_Cleanup
'   M_Lifecycle_RaiseIfFault
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "DP_Start"

    Dim HandlerStep As String
    Dim ErrorNumber As Long
    Dim ErrorDescription As String
    Dim ErrorSource As String
    Dim CallerEnableEvents As Boolean
    Dim HasCallerEnableEvents As Boolean
    Dim PreOwned As Boolean
    Dim AcquiredThisCall As Boolean

    On Error GoTo ErrorHandler
    M_Lifecycle_ResetObservation

    mDP_LifecycleLastOperation = PROC_NAME
    HandlerStep = "Capture caller event state"
    CallerEnableEvents = Excel.Application.EnableEvents
    HasCallerEnableEvents = True

    HandlerStep = "Classify existing provider lease"
    PreOwned = M_Lease_IsOwner()

    mDP_LifecycleLastLeaseWasAlreadyOwned = PreOwned
    HandlerStep = "Acquire provider lease"
    If Not M_Lease_EnsureAdmitted(PROC_NAME) Then GoTo CleanExit
    AcquiredThisCall = (Not PreOwned And M_Lease_IsOwner())

    mDP_LifecycleLastLeaseAcquiredThisCall = AcquiredThisCall
    HandlerStep = "After provider admission"
    M_Lifecycle_RaiseIfFault "Start.AfterAdmission"

    HandlerStep = "Ensure manager"
    M_Picker_EnsureManager
    M_Lifecycle_RaiseIfFault "Start.AfterManager"

    HandlerStep = "Synchronize right-click menu"
    M_ContextMenu_Update
    M_Lifecycle_RaiseIfFault "Start.AfterContextMenu"

    HandlerStep = "Synchronize keyboard shortcut"
    M_KeyboardShortcut_Update
    M_Lifecycle_RaiseIfFault "Start.AfterKeyboard"

    HandlerStep = "Pre-create hidden grid icon"
    M_GridIcon_PreCreateHidden
    M_Lifecycle_RaiseIfFault "Start.AfterGrid"

    HandlerStep = "Refresh current selection context"
    gDP_Manager.Handle_SelectionChange
    M_Lifecycle_RaiseIfFault "Start.AfterRefresh"

CleanExit:
    HandlerStep = "Restore caller event state"
    If HasCallerEnableEvents Then Excel.Application.EnableEvents = CallerEnableEvents
    mDP_LifecycleLastSucceeded = M_Lease_IsOwner()
    Exit Sub

ErrorHandler:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    ErrorSource = Err.Source
    M_Lifecycle_SetPrimaryFailure ErrorNumber, HandlerStep, ErrorDescription

    mDP_LifecycleLastSucceeded = False
    On Error Resume Next
    If AcquiredThisCall Then
        M_Lifecycle_Cleanup True, PROC_NAME & ".Rollback", _
            HasCallerEnableEvents, CallerEnableEvents
    ElseIf PreOwned Then
        mDP_LifecycleLastTrace = "PreOwnedRuntime=PRESERVED"
        mDP_LifecycleLastCriticalClean = True
        mDP_LifecycleLastLeaseReleased = False
    End If
    If Not AcquiredThisCall Then
        If HasCallerEnableEvents Then Excel.Application.EnableEvents = CallerEnableEvents
    End If
    Err.Clear
    On Error GoTo 0

    Err.Raise ErrorNumber, _
        PROC_NAME & " | Step=" & HandlerStep & " | Source=" & ErrorSource, _
        "DatePicker startup failed: " & ErrorDescription

End Sub

Public Sub DP_Show_Impl()

'
'------------------------------------------------------------------------------
'                           SHOW DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Show
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Show
'
' RETURNS
'   As documented on M_DatePicker.DP_Show
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Show
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Show. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Lease_EnsureAdmitted
'   M_Picker_EnsureManager
'   M_FormBridge_TryReuseLoadedPickerFromActiveCell
'   M_FormBridge_GetLoadedForm
'   M_Timer_ApplyClockMode
'   M_Window_MoveFormToMouse
'   UF_DatePicker
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        StepName = "Admit this provider"

'------------------------------------------------------------------------------
' ADMIT THIS PROVIDER
'------------------------------------------------------------------------------
    'This is the shared admission boundary for every interactive open path.
    'DP_Click, DP_OpenForActiveCell and Ribbon_ShowPicker all delegate here, so
    'admitting once refuses once and reports once. Nothing above this point has
    'touched the manager, the form or any application-wide registration
        If Not M_Lease_EnsureAdmitted(PROC_NAME) Then
            Exit Sub
        End If

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

Public Sub DP_Click_Impl()

'
'------------------------------------------------------------------------------
'                       DATEPICKER CLICK ENTRY POINT
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Click
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Click
'
' RETURNS
'   As documented on M_DatePicker.DP_Click
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Click
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Click. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   DP_OpenForActiveCell
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        DP_OpenForActiveCell_Impl

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

Public Sub DP_OpenForActiveCell_Impl()

'
'------------------------------------------------------------------------------
'                       OPEN DATEPICKER FOR ACTIVE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_OpenForActiveCell
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_OpenForActiveCell
'
' RETURNS
'   As documented on M_DatePicker.DP_OpenForActiveCell
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_OpenForActiveCell
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_OpenForActiveCell. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   DP_Show
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        DP_Show_Impl

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

Public Sub DP_RepairRuntime_Impl()

'
'------------------------------------------------------------------------------
'                          REPAIR DATEPICKER RUNTIME
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_RepairRuntime
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_RepairRuntime
'
' RETURNS
'   As documented on M_DatePicker.DP_RepairRuntime
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_RepairRuntime
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_RepairRuntime. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Lifecycle_Cleanup
'   M_Picker_EnsureManager
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Const PROC_NAME As String = "DP_RepairRuntime"

    Dim HandlerStep As String
    Dim ErrorNumber As Long
    Dim ErrorDescription As String
    Dim ErrorSource As String

    On Error GoTo ErrorHandler
    M_Lifecycle_ResetObservation

    mDP_LifecycleLastOperation = PROC_NAME
    mDP_LifecycleLastLeaseWasAlreadyOwned = M_Lease_IsOwner()
    mDP_LifecycleLastLeaseAcquiredThisCall = False
    HandlerStep = "Verify provider ownership"
    If Not M_Lease_IsOwner() Then
        M_Lease_ReportRefusal PROC_NAME
        Exit Sub
    End If

    HandlerStep = "Re-enable Excel events"
    Excel.Application.EnableEvents = True

    HandlerStep = "Clean existing runtime"
    If Not M_Lifecycle_Cleanup(False, PROC_NAME & ".Prepare") Then
        Err.Raise vbObjectError + 2721, PROC_NAME, _
            "DatePicker runtime repair refused to rebuild over incomplete cleanup: " & _
            mDP_LifecycleLastCleanupDetail
    End If

    HandlerStep = "Recreate manager"
    M_Picker_EnsureManager
    M_Lifecycle_RaiseIfFault "Repair.AfterManager"

    HandlerStep = "Synchronize right-click menu"
    M_ContextMenu_Update
    M_Lifecycle_RaiseIfFault "Repair.AfterContextMenu"

    HandlerStep = "Synchronize keyboard shortcut"
    M_KeyboardShortcut_Update
    M_Lifecycle_RaiseIfFault "Repair.AfterKeyboard"

    HandlerStep = "Refresh current selection context"
    If Not gDP_Manager Is Nothing Then gDP_Manager.Handle_SelectionChange
    M_Lifecycle_RaiseIfFault "Repair.AfterRefresh"


    mDP_LifecycleLastSucceeded = True
    Exit Sub

ErrorHandler:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    ErrorSource = Err.Source
    M_Lifecycle_SetPrimaryFailure ErrorNumber, HandlerStep, ErrorDescription

    mDP_LifecycleLastSucceeded = False
    On Error Resume Next
    If M_Lease_IsOwner() Then
        If VBA.StrComp(HandlerStep, "Clean existing runtime", vbBinaryCompare) <> 0 Then
            M_Lifecycle_Cleanup False, PROC_NAME & ".Rollback"
        End If
    End If
    Excel.Application.EnableEvents = True
    Err.Clear
    On Error GoTo 0

    Err.Raise ErrorNumber, _
        PROC_NAME & " | Step=" & HandlerStep & " | Source=" & ErrorSource, _
        "DatePicker runtime repair failed: " & ErrorDescription

End Sub

Public Sub DP_Close_Impl()

'
'------------------------------------------------------------------------------
'                           CLOSE DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Close
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Close
'
' RETURNS
'   As documented on M_DatePicker.DP_Close
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Close
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Close. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Timer_Stop
'   M_FormBridge_GetLoadedForm
'   DP_FORM_NAME
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub DP_Hide_Impl()

'
'------------------------------------------------------------------------------
'                           HIDE DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Hide
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Hide
'
' RETURNS
'   As documented on M_DatePicker.DP_Hide
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Hide
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Hide. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Timer_Stop
'   M_FormBridge_GetLoadedForm
'   DP_FORM_NAME
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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

Public Sub DP_Stop_Impl()

'
'------------------------------------------------------------------------------
'                               STOP DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Stop
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Stop
'
' RETURNS
'   As documented on M_DatePicker.DP_Stop
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Stop
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Stop. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Lease_IsOwner
'   M_Lease_TryReclaimVanished
'   M_Lifecycle_Cleanup
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    Dim CallerEnableEvents As Boolean
    Dim HasCallerEnableEvents As Boolean
    Dim OwnedOnEntry As Boolean

    On Error Resume Next
    M_Lifecycle_ResetObservation
    mDP_LifecycleLastOperation = "DP_Stop"

    CallerEnableEvents = Excel.Application.EnableEvents
    HasCallerEnableEvents = (Err.Number = 0)
    Err.Clear

    OwnedOnEntry = M_Lease_IsOwner()
    mDP_LifecycleLastLeaseWasAlreadyOwned = OwnedOnEntry
    mDP_LifecycleLastLeaseAcquiredThisCall = False
    If Not OwnedOnEntry Then
        OwnedOnEntry = M_Lease_TryReclaimVanished("DP_Stop")
        mDP_LifecycleLastLeaseAcquiredThisCall = OwnedOnEntry
    End If

    If Not OwnedOnEntry Then
        Debug.Print "DP_Stop | Skipped | This copy cannot prove runtime ownership, " & _
            "so it touches nothing"
        mDP_LifecycleLastSucceeded = False
        GoTo CleanExit
    End If

    mDP_LifecycleLastSucceeded = M_Lifecycle_Cleanup( _
        True, "DP_Stop", HasCallerEnableEvents, CallerEnableEvents)

CleanExit:
    Err.Clear
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
    Optional ByVal NoTableGrow As Boolean = True)

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
'   M_WriteBack_ReportShortfall
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   Calendar day selection is intentionally date-only
'
'   SelectedDate is normalized once and the normalized value is reused for
'   write-back, selected-state storage, and optional open-form refresh
'
'   This is an interactive entry point, so it reports a partial write once, after
'   the whole operation. The write engine collects the facts and displays nothing
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Picker_SelectDate"

    Dim SelectedDateOnly    As Date         'Selected date without time
    Dim WriteResult         As DP_WriteResult   'Structured write-back outcome
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
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, NoTableGrow)

'------------------------------------------------------------------------------
' STORE SELECTED DATE AFTER SUCCESSFUL WRITE-BACK
'------------------------------------------------------------------------------
    'Store the selected date only when at least one cell actually received it.
    'Before #21 a zero-write raised out of the engine and never reached here, so
    'the guard preserves that contract now that a zero-write returns a complete
    'result instead of an exception
        If WriteResult.WrittenCount > 0 Then
            'Store the selected date
                gDP_SelectedDate = SelectedDateOnly
            'Mark the selected date as available
                gDP_HasSelectedDate = True
        End If

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Suppress post-write-back visual lifecycle errors
        On Error Resume Next
    'Close the DatePicker after successful selection when configured
        If gDP_CloseAfterSelection Then
            'Hide the DatePicker form
                DP_Hide_Impl
        Else
            'Refresh the open DatePicker form after successful selection
                M_FormBridge_AfterSuccessfulSelection SelectedDateOnly
        End If
    'Clear any suppressed visual lifecycle error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REPORT PARTIAL WRITE
'------------------------------------------------------------------------------
    'Report once for the whole operation when some cells were not written
        M_WriteBack_ReportShortfall WriteResult

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

Public Sub DP_Today_Impl()

'
'------------------------------------------------------------------------------
'                               WRITE TODAY
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Today
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Today
'
' RETURNS
'   As documented on M_DatePicker.DP_Today
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Today
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Today. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Picker_SelectDate
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
    'Delegate date-only write-back to the canonical DatePicker selection routine.
    'NoTableGrow is deliberately omitted so Today inherits the safe single-cell
    'default rather than opting into table-column expansion
        M_Picker_SelectDate TodayDate

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

Public Sub DP_Now_Impl()

'
'------------------------------------------------------------------------------
'                               WRITE NOW
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_Now
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_Now
'
' RETURNS
'   As documented on M_DatePicker.DP_Now
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_Now
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_Now. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_EnsureLoaded
'   M_WriteBack_Apply
'   M_WriteBack_ReportShortfall
'   DP_Close
'   M_FormBridge_AfterSuccessfulSelection
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "DP_Now"

    Dim NowValue            As Date                 'Current system date-time
    Dim NowDate             As Date                 'Date-only part of current timestamp
    Dim WriteResult         As DP_WriteResult       'Structured write-back outcome
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
    'Apply the date-time value to the current Excel target. NoTableGrow is
    'deliberately omitted so Now inherits the safe single-cell default rather
    'than opting into table-column expansion
        WriteResult = M_WriteBack_Apply(DP_WriteAction_DatePicker)

'------------------------------------------------------------------------------
' STORE SELECTED DATE AFTER SUCCESSFUL WRITE-BACK
'------------------------------------------------------------------------------
    'Store the date only when at least one cell actually received it. Before #21
    'a zero-write raised out of the engine and never reached here, so the guard
    'preserves that contract now that a zero-write returns a complete result
        If WriteResult.WrittenCount > 0 Then
            'Store the date-only part as the active selected date
                gDP_SelectedDate = NowDate
            'Mark selected-date state as available
                gDP_HasSelectedDate = True
        End If

'------------------------------------------------------------------------------
' CLOSE OR REFRESH FORM
'------------------------------------------------------------------------------
    'Suppress post-write-back visual lifecycle errors
        On Error Resume Next
    'Close the DatePicker after successful write-back when configured
        If gDP_CloseAfterSelection Then
            'Hide the DatePicker form
                DP_Hide_Impl
        Else
            'Refresh the open DatePicker form using the date-only part
                M_FormBridge_AfterSuccessfulSelection NowDate
        End If
    'Clear any suppressed visual lifecycle error
        Err.Clear
    'Restore controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REPORT PARTIAL WRITE
'------------------------------------------------------------------------------
    'Report once for the whole operation when some cells were not written
        M_WriteBack_ReportShortfall WriteResult

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

Public Function DP_FillTableColumn_Impl( _
    ByVal ValueToWrite As Date, _
    Optional ByVal ConfirmFill As Boolean = True, _
    Optional ByVal OverwriteFormulas As Boolean = False) As DP_WriteResult

'
'------------------------------------------------------------------------------
'                           FILL TABLE COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.DP_FillTableColumn
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.DP_FillTableColumn
'
' RETURNS
'   As documented on M_DatePicker.DP_FillTableColumn
'
' BEHAVIOR
'   As documented on M_DatePicker.DP_FillTableColumn
'
' ERROR POLICY
'   As documented on M_DatePicker.DP_FillTableColumn. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_WriteBack_TryResolveTableColumn
'   M_WriteBack_Apply
'   M_WriteBack_ReportShortfall
'   gDP_WriteValue
'   DP_MSGBOX_TITLE
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "DP_FillTableColumn"

    Dim FillResult          As DP_WriteResult   'Structured outcome of the fill
    Dim TargetColumn        As Range        'Resolved table data column
    Dim TableName           As String       'Owning table name
    Dim ColumnName          As String       'Resolved column name
    Dim CellCount           As Long         'Cells the fill would affect
    Dim OldWriteValue       As Date         'Previous pending write value
    Dim ValueApplied        As Boolean      'True once the write value was replaced
    Dim PromptText          As String       'Confirmation prompt text
    Dim HandlerStep         As String       'Current handler step for diagnostics

    Dim SavedErrNumber          As Long     'Captured original error number
    Dim SavedErrSource          As String   'Captured original error source
    Dim SavedErrDescription     As String   'Captured original error description

    Dim CleanupErrNumber        As Long     'Captured cleanup error number
    Dim CleanupErrDescription   As String   'Captured cleanup error description

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
        HandlerStep = "Validate value"
    'Reject a zero date, which the write engine also rejects
        If ValueToWrite = 0 Then
            Err.Raise vbObjectError + 530, PROC_NAME, "ValueToWrite cannot be zero"
        End If

'------------------------------------------------------------------------------
' RESOLVE TABLE COLUMN
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve table column"
    'Report and exit when the selection is not a table data cell
        If Not M_WriteBack_TryResolveTableColumn(TargetColumn, TableName, ColumnName) Then
            'Tell the user what is required, but only on the interactive path
                If ConfirmFill Then
                    VBA.MsgBox _
                        "A table data cell is required." & VBA.vbCrLf & VBA.vbCrLf & _
                        "Select a cell in the Excel Table column you want to fill, " & _
                        "then try again.", _
                        vbExclamation, _
                        DP_MSGBOX_TITLE
                End If
            'Exit cleanly. This is a usage condition, not a failure
                GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' CONFIRM SCOPE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Confirm fill scope"
    'Resolve how many cells the fill is expected to affect
        CellCount = VBA.CLng(TargetColumn.Cells.CountLarge)
    'Describe the resolved scope before writing it
        If ConfirmFill Then
            'Build the confirmation text from the resolved target
                PromptText = "Fill " & VBA.CStr(CellCount) & " cells in " & _
                    TableName & "[" & ColumnName & "] with " & _
                    VBA.Format$(ValueToWrite, "dd-mmm-yyyy") & "?"
            'Exit when the user declines the described scope
                If VBA.MsgBox(PromptText, vbQuestion Or vbYesNo, DP_MSGBOX_TITLE) <> vbYes Then
                    GoTo CleanExit
                End If
        End If

'------------------------------------------------------------------------------
' APPLY FILL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Apply table column fill"
    'Preserve the pending write value so it can be restored on failure
        OldWriteValue = gDP_WriteValue
    'Stage the value the write engine will apply
        gDP_WriteValue = ValueToWrite
        ValueApplied = True
    'Write through the normal engine with table expansion explicitly enabled
        FillResult = M_WriteBack_Apply(DP_WriteAction_DatePicker, _
            NoTableGrow:=False, OverwriteFormulas:=OverwriteFormulas)
    'Publish the structured outcome
        DP_FillTableColumn_Impl = FillResult

'------------------------------------------------------------------------------
' CHECK PREDICTED SCOPE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check predicted fill scope"
    'The prediction describes the scope, so it is checked against AttemptedCount
        If FillResult.AttemptedCount <> CellCount Then
            'Tell the user the described scope is not the scope that was written
                If ConfirmFill Then
                    VBA.MsgBox _
                        "The fill scope changed after it was described." & _
                        VBA.vbCrLf & VBA.vbCrLf & _
                        "Described: " & VBA.CStr(CellCount) & " cells in " & _
                        TableName & "[" & ColumnName & "]" & VBA.vbCrLf & _
                        "Written to: " & VBA.CStr(FillResult.AttemptedCount) & " cells in " & _
                        FillResult.ResolvedTargetAddress, _
                        vbExclamation, _
                        DP_MSGBOX_TITLE
                End If
            'Record the inconsistency on every path, including the harness
                Debug.Print PROC_NAME & ": predicted " & VBA.CStr(CellCount) & _
                    " cells but the engine targeted " & VBA.CStr(FillResult.AttemptedCount)
        End If

'------------------------------------------------------------------------------
' REPORT PARTIAL WRITE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Report partial fill"
    'A legitimate partial write is reported once, and only on the interactive path
        If ConfirmFill Then
            M_WriteBack_ReportShortfall FillResult
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set TargetColumn = Nothing
    'Exit the procedure
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the original error before any cleanup runs
    '
    'Cleanup below restores gDP_WriteValue under On Error Resume Next and clears
    'Err. At v1.2.0 the raise that followed read Err.Number and Err.Description
    'after that clear, so whenever ValueApplied was True the caller was handed
    'error 0 with a blank cause instead of the failure that actually occurred.
    'Nothing may read the live Err object below this point: see #48
        SavedErrNumber = Err.Number
        SavedErrSource = PROC_NAME & " | Step=" & HandlerStep
        SavedErrDescription = "Table column fill failed: " & Err.Description
    'Restore the previous pending write value when the write failed
        If ValueApplied Then
            On Error Resume Next
            gDP_WriteValue = OldWriteValue
            'Capture a cleanup failure separately rather than letting it replace
            'the primary failure
                If Err.Number <> 0 Then
                    CleanupErrNumber = Err.Number
                    CleanupErrDescription = Err.Description
                End If
            Err.Clear
            On Error GoTo 0
        End If
    'Release object references
        Set TargetColumn = Nothing
    'Append cleanup diagnostics when cleanup also failed
        If CleanupErrNumber <> 0 Then
            SavedErrDescription = SavedErrDescription & _
                " Cleanup also failed while restoring the pending write value: " & _
                CleanupErrDescription
        End If
    'Raise the original error after best-effort cleanup
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Function

Public Function M_WriteBack_Apply( _
    ByVal iType As DP_WriteAction, _
    Optional ByVal NoTableGrow As Boolean = True, _
    Optional ByVal OverwriteFormulas As Boolean = False) As DP_WriteResult

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
'   OverwriteFormulas
'     False, the default, preserves a cell holding a formula and counts it in
'     FormulaSkippedCount. True replaces it
'
' RETURNS
'   DP_WriteResult describing the completed write-back
'
'     AttemptedCount, WrittenCount
'     LockedSkippedCount, LockedSkippedAddresses
'     FormulaSkippedCount, FormulaSkippedAddresses
'     FailedCount, FailedAddresses
'     ResolvedTargetAddress, ExpandedToTableColumn, TableName, ColumnName
'     AreasCount, EventsDisabledByCaller
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
'   A raised error means nothing was written. A returned result with
'   WrittenCount below AttemptedCount means a partial write
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
'   This is a Function so callers can inspect the outcome, but bare-call syntax
'   still compiles. A caller that ignores the result behaves exactly as before
'
'   A ByRef output was rejected: VBA does not permit a user-defined type as an
'   Optional parameter, so the argument would have to be required and every
'   existing caller would need editing
'
'   This routine displays nothing. Programmatic callers must be able to consume
'   the result without being taken through modal UI, so any interactive summary
'   belongs to the entry point above, which calls M_WriteBack_ReportShortfall
'   once for the whole operation
'
'   The returned result satisfies:
'     AttemptedCount = WrittenCount + LockedSkippedCount
'                    + FormulaSkippedCount + FailedCount
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_WriteBack_Apply" 'Current procedure name

    Dim Result                  As DP_WriteResult   'Structured write-back outcome

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
        M_WriteBack_ResolveAndApplyTarget iType, NoTableGrow, OverwriteFormulas, Result
    'Record the caller's event state in the same result
        Result.EventsDisabledByCaller = Not PreviousEvents
    'Publish the structured outcome
        M_WriteBack_Apply = Result

'------------------------------------------------------------------------------
' LOG PARTIAL WRITE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Log partial write"
    'Record a partial write for the developer without interrupting the caller
        If Result.WrittenCount < Result.AttemptedCount Or Result.TechnicalFailureOccurred Then
            Debug.Print PROC_NAME & ": wrote " & VBA.CStr(Result.WrittenCount) & _
                " of " & VBA.CStr(Result.AttemptedCount) & " cells - " & _
                M_WriteBack_DescribeShortfall(Result)
        End If

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
        Exit Function

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

End Function

Private Sub M_WriteBack_AppendAddress( _
    ByRef AddressList As String, _
    ByVal RecordedCount As Double, _
    ByVal AddressText As String)

'
'------------------------------------------------------------------------------
'                          APPEND RESULT ADDRESS
'------------------------------------------------------------------------------
' PURPOSE
'   Appends one real worksheet-qualified address to a structured write-result
'   address list while enforcing the operation-level reporting budget
'
' WHY THIS EXISTS
'   A partial write has to identify cells that were not written, but an operation
'   over a long or discontiguous target must not build unbounded diagnostic text
'
' INPUTS
'   AddressList
'     Structured address list, modified in place
'
'   RecordedCount
'     Operation-level classified count for this category, including this item
'
'   AddressText
'     Worksheet-qualified address to append
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Retains at most DP_WRITEBACK_ADDRESS_LIMIT addresses for one outcome category
'   across the complete write operation. A new area does not reset that budget.
'   Structured fields contain addresses only: no ellipsis or human truncation
'   sentinel is ever appended
'
' ERROR POLICY
'   Best-effort. Never raises, because it runs inside the per-cell write loop
'
' DEPENDENCIES
'   DP_WRITEBACK_ADDRESS_LIMIT
'
' NOTES
'   The caller supplies worksheet-qualified addresses in SheetName!A1 form.
'   Non-address diagnostics such as "(no cell)" are deliberately excluded from
'   structured address fields. Exact category totals remain uncapped elsewhere;
'   M_WriteBack_DescribeShortfall reports the exact omitted count to humans
'
' UPDATED
'   2026-09-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let diagnostics break a write
        On Error Resume Next

'------------------------------------------------------------------------------
' ENFORCE OPERATION-LEVEL ADDRESS BUDGET
'------------------------------------------------------------------------------
    'Once this category has consumed its operation-level budget, retain only the
    'exact classification count and leave the structured address list unchanged
        If RecordedCount > DP_WRITEBACK_ADDRESS_LIMIT Then Exit Sub

'------------------------------------------------------------------------------
' REQUIRE A STRUCTURED WORKSHEET ADDRESS
'------------------------------------------------------------------------------
    'Structured address fields must contain only worksheet-qualified addresses
        If VBA.InStr(1, AddressText, "!", vbBinaryCompare) <= 1 Then Exit Sub

'------------------------------------------------------------------------------
' APPEND ADDRESS
'------------------------------------------------------------------------------
    'Start the list or extend it with the real address only
        If VBA.LenB(AddressList) = 0 Then
            AddressList = AddressText
        Else
            AddressList = AddressList & ", " & AddressText
        End If

End Sub

Private Function M_WriteBack_CountRetainedAddresses( _
    ByVal AddressList As String) As Double

'
'------------------------------------------------------------------------------
'                    COUNT RETAINED WRITE ADDRESSES
'------------------------------------------------------------------------------
' PURPOSE
'   Counts the structured worksheet addresses actually retained in one result
'   field so human omission reporting can be derived from observed output
'
' INPUTS
'   AddressList
'     Comma-space-delimited structured address field built by AppendAddress
'
' RETURNS
'   Number of retained address entries; zero for a blank or unreadable list
'
' ERROR POLICY
'   Best-effort. Returns zero on an unexpected parsing failure
'
' NOTES
'   This counts the representation emitted by M_WriteBack_AppendAddress. It does
'   not consult the classification total or the cap, so omitted count remains
'   exactly: classified count minus addresses actually retained
'
' UPDATED
'   2026-09-17
'------------------------------------------------------------------------------

    Dim AddressParts As Variant

    On Error GoTo SafeExit
    If VBA.LenB(VBA.Trim$(AddressList)) = 0 Then Exit Function
    AddressParts = VBA.Split(AddressList, ", ")
    M_WriteBack_CountRetainedAddresses = _
        UBound(AddressParts) - LBound(AddressParts) + 1

SafeExit:
    Err.Clear

End Function

Public Sub M_WriteBack_Test_SetFaultInjection( _
    ByVal FailInAreaOrdinal As Long, _
    Optional ByVal FailAfterWrittenCellsInArea As Long = 0)

'
'------------------------------------------------------------------------------
'                    ARM WRITE-BACK FAULT INJECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a single forced technical failure inside the next write-back operation
'
' WHY THIS EXISTS
'   #21 promises that partial mutation is never reported as an exception carrying
'   no result. That promise is only meaningful on the unexpected-technical-error
'   path, and that path cannot be produced on demand: M_WriteBack_TryWriteCell
'   classifies every per-cell failure it can observe and raises nothing, so a
'   genuine technical error inside the area loop is exactly the case a test has
'   no way to arrange
'
'   Classified cell failures such as an array-formula refusal are already covered
'   by the MultiAreaWriteResult suite. They become FailedCount and are a
'   different path from this one
'
'   The regression module is a separate VBA module and cannot assign private
'   state in M_DP_Core, so the setter has to be technically Public
'
' INPUTS
'   FailInAreaOrdinal
'     1-based target area to fail in, counted within one write operation.
'     Zero disarms
'
'   FailAfterWrittenCellsInArea
'     Cells to write inside that area before failing. Zero fails on entering the
'     area, before any of its cells are mutated
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the requested failure point for the next write operation only
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
'
' NOTES
'   THIS IS INTERNAL TEST INFRASTRUCTURE. It is not supported DatePicker API,
'   is classified internal under #25, and must not appear in the README public
'   API table
'
'   Injection is one-shot. The fault disarms itself as it fires, so an armed
'   test cannot leak into a later real write. A test that arms a fault it never
'   reaches must disarm it explicitly in cleanup
'
'   The required first argument keeps this out of the Alt+F8 macro list
'
'   The forced error is raised through the ordinary Err.Raise path inside
'   M_WriteBack_PopulateRange, so it exercises the real handler rather than a
'   test-only branch of it
'
'   No state is persisted. Nothing is written to the registry, a workbook, or
'   any Excel object
'
' UPDATED
'   2026-08-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ARM INJECTION
'------------------------------------------------------------------------------
    'Arm the requested target area for the next write operation only
        mDP_TestWriteFailArea = FailInAreaOrdinal
    'Arm how far into that area the failure fires
        mDP_TestWriteFailAfterCells = FailAfterWrittenCellsInArea
    'Reset the per-operation area counter so arming is independent of history
        mDP_TestWriteAreaOrdinal = 0

End Sub

Private Function M_WriteBack_TestFaultShouldFire( _
    ByVal AreaOrdinal As Long, _
    ByVal WrittenInArea As Double) As Boolean

'
'------------------------------------------------------------------------------
'                   RESOLVE ARMED WRITE-BACK FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether an armed write-back fault fires at this point
'
' WHY THIS EXISTS
'   The decision is consulted from two places inside one area, and both have to
'   agree about one-shot consumption
'
' INPUTS
'   AreaOrdinal
'     1-based ordinal of the area currently being populated
'
'   WrittenInArea
'     Cells written inside that area so far
'
' RETURNS
'   True when the caller must raise the forced technical failure
'
' BEHAVIOR
'   Fires at most once, and disarms itself as it fires
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
'
' NOTES
'   Disarmed state is the normal state. Production never arms this
'
' UPDATED
'   2026-08-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Set safe default result
        M_WriteBack_TestFaultShouldFire = False
    'Exit when no fault is armed
        If mDP_TestWriteFailArea <= 0 Then Exit Function
    'Exit when this is not the armed area
        If AreaOrdinal <> mDP_TestWriteFailArea Then Exit Function
    'Exit until the armed number of cells has been written inside that area
        If WrittenInArea < mDP_TestWriteFailAfterCells Then Exit Function

'------------------------------------------------------------------------------
' CONSUME
'------------------------------------------------------------------------------
    'Disarm before reporting, so the fault fires exactly once
        mDP_TestWriteFailArea = 0
        mDP_TestWriteFailAfterCells = 0
    'Report that the caller must raise
        M_WriteBack_TestFaultShouldFire = True

End Function

Private Function M_WriteBack_HasCellOutcomes( _
    ByRef Result As DP_WriteResult) As Boolean

'
'------------------------------------------------------------------------------
'                   REPORT WHETHER A RESULT HOLDS CELL OUTCOMES
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a write result records anything that actually happened to a
'   cell
'
' WHY THIS EXISTS
'   Deciding what an unexpected technical failure means turns on one question:
'   is there anything to lose? That is answered by the per-cell outcomes, not by
'   AttemptedCount
'
'   AttemptedCount is the size of the target. It is recorded before the first cell
'   is touched, so treating it as evidence would classify a validation failure
'   that did nothing as a failure worth protecting, and would silently swallow the
'   programming errors this engine is supposed to raise loudly
'
' INPUTS
'   Result
'     Write result to inspect. Either an area result or an operation result
'
' RETURNS
'   True when at least one cell was written, skipped or failed
'
' BEHAVIOR
'   Pure. Reads four counters and returns
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
'
' NOTES
'   A written cell means the workbook was mutated. A skipped or failed cell means
'   an observation was made that the caller asked for. Both are facts #21 promises
'   not to discard
'
' UPDATED
'   2026-08-25
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE
'------------------------------------------------------------------------------
    'Report whether any cell produced an outcome
        M_WriteBack_HasCellOutcomes = _
            (Result.WrittenCount > 0) Or _
            (Result.LockedSkippedCount > 0) Or _
            (Result.FormulaSkippedCount > 0) Or _
            (Result.FailedCount > 0)

End Function

Public Sub M_WriteBack_ReportShortfall( _
    ByRef Result As DP_WriteResult)

'
'------------------------------------------------------------------------------
'                          REPORT WRITE SHORTFALL
'------------------------------------------------------------------------------
' PURPOSE
'   Shows one consolidated message describing the cells a write-back did not
'   write
'
' WHY THIS EXISTS
'   The write engine collects facts and returns them. Deciding whether a human is
'   told about a partial write belongs to the entry point the human invoked
'
'   Reporting from inside the engine produced one dialog per target area, and
'   forced every programmatic caller through modal UI
'
' INPUTS
'   Result
'     Completed DP_WriteResult to report
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Shows nothing when every attempted cell was written
'
'   Otherwise shows one message naming the counts and the addresses behind them
'
' ERROR POLICY
'   Best-effort. Never raises, because a reporting failure must not turn a
'   successful partial write into an error
'
' DEPENDENCIES
'   M_WriteBack_DescribeShortfall
'   DP_MSGBOX_TITLE
'
' NOTES
'   Interactive entry points call this once, after the whole operation. A
'   programmatic caller inspects the result and never calls it
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let reporting break a completed write
        On Error Resume Next

'------------------------------------------------------------------------------
' SKIP A COMPLETE WRITE
'------------------------------------------------------------------------------
    'Say nothing when every attempted cell was written
    '
    'A technical failure is always reported, even when the cells that were
    'attempted all succeeded. The areas the operation never reached are not in
    'AttemptedCount, so the counts alone can look complete while the operation
    'actually stopped early
        If Result.WrittenCount >= Result.AttemptedCount Then
            If Not Result.TechnicalFailureOccurred Then Exit Sub
        End If

'------------------------------------------------------------------------------
' REPORT THE SHORTFALL
'------------------------------------------------------------------------------
    'Describe the whole operation in one message
        MsgBox _
            "Wrote " & VBA.CStr(Result.WrittenCount) & " of " & _
            VBA.CStr(Result.AttemptedCount) & " cells in " & _
            Result.ResolvedTargetAddress & "." & VBA.vbCrLf & VBA.vbCrLf & _
            M_WriteBack_DescribeShortfall(Result), _
            vbInformation Or vbOKOnly, _
            DP_MSGBOX_TITLE
    'Clear any suppressed reporting error
        Err.Clear

End Sub

Public Function M_WriteBack_DescribeShortfall( _
    ByRef Result As DP_WriteResult) As String

'
'------------------------------------------------------------------------------
'                          DESCRIBE WRITE SHORTFALL
'------------------------------------------------------------------------------
' PURPOSE
'   Builds one human-readable description of the cells a write-back did not write
'
' WHY THIS EXISTS
'   Structured result fields stay machine-readable and bounded. Human reporting
'   adds the exact number of classified cells whose addresses were omitted by the
'   operation-level cap without embedding a sentinel in the structured fields
'
' INPUTS
'   Result
'     Completed DP_WriteResult to describe
'
' RETURNS
'   Description of skipped and failed cells; empty when there is no shortfall
'
' BEHAVIOR
'   Describes protected locked cells, preserved formulas and failures using the
'   exact uncapped category totals plus the retained structured addresses. When a
'   category exceeds DP_WRITEBACK_ADDRESS_LIMIT, reports the exact excess count
'   as human text derived from total minus the addresses actually retained
'
' ERROR POLICY
'   Best-effort. Never raises, because it is called while reporting an outcome
'
' DEPENDENCIES
'   DP_WRITEBACK_ADDRESS_LIMIT
'
' NOTES
'   Structured address fields contain only worksheet-qualified addresses. Human
'   omission text is derived here and is never stored back into DP_WriteResult
'
' UPDATED
'   2026-09-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Description     As String       'Accumulated description
    Dim OmittedCount    As Double       'Exact category count omitted by the cap

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Never let reporting break a caller
        On Error Resume Next
    'Set safe default result
        M_WriteBack_DescribeShortfall = VBA.vbNullString

'------------------------------------------------------------------------------
' DESCRIBE PROTECTED LOCKED CELLS
'------------------------------------------------------------------------------
    'Describe the protected locked cells that were skipped
        If Result.LockedSkippedCount > 0 Then
            Description = VBA.CStr(Result.LockedSkippedCount) & " protected locked: " & _
                Result.LockedSkippedAddresses
            OmittedCount = Result.LockedSkippedCount - _
                M_WriteBack_CountRetainedAddresses(Result.LockedSkippedAddresses)
            If OmittedCount > 0 Then
                Description = Description & " (" & VBA.CStr(OmittedCount) & _
          " additional classified cells omitted from address list)"
            End If
        End If

'------------------------------------------------------------------------------
' DESCRIBE PRESERVED FORMULA CELLS
'------------------------------------------------------------------------------
    'Describe the formula cells policy left in place
        If Result.FormulaSkippedCount > 0 Then
            If VBA.LenB(Description) > 0 Then Description = Description & VBA.vbCrLf
            Description = Description & VBA.CStr(Result.FormulaSkippedCount) & _
                " formula cells preserved: " & Result.FormulaSkippedAddresses
            OmittedCount = Result.FormulaSkippedCount - _
                M_WriteBack_CountRetainedAddresses(Result.FormulaSkippedAddresses)
            If OmittedCount > 0 Then
                Description = Description & " (" & VBA.CStr(OmittedCount) & _
          " additional classified cells omitted from address list)"
            End If
        End If

'------------------------------------------------------------------------------
' DESCRIBE SUPPRESSED FAILURES
'------------------------------------------------------------------------------
    'Describe the cells that failed for another reason
        If Result.FailedCount > 0 Then
            If VBA.LenB(Description) > 0 Then Description = Description & VBA.vbCrLf
            Description = Description & VBA.CStr(Result.FailedCount) & " failed: " & _
                Result.FailedAddresses
            OmittedCount = Result.FailedCount - _
                M_WriteBack_CountRetainedAddresses(Result.FailedAddresses)
            If OmittedCount > 0 Then
                Description = Description & " (" & VBA.CStr(OmittedCount) & _
          " additional classified cells omitted from address list)"
            End If
        End If

'------------------------------------------------------------------------------
' DESCRIBE AN UNEXPECTED TECHNICAL FAILURE
'------------------------------------------------------------------------------
    'Describe an unexpected error that stopped the operation. Cells the operation
    'never reached are absent from every classified count above
        If Result.TechnicalFailureOccurred Then
            If VBA.LenB(Description) > 0 Then Description = Description & VBA.vbCrLf
            Description = Description & _
                "The operation stopped early after an unexpected error at step """ & _
                Result.TechnicalFailureStep & """. Any remaining cells were not attempted."
        End If

'------------------------------------------------------------------------------
' RETURN DESCRIPTION
'------------------------------------------------------------------------------
    'Return the accumulated description
        M_WriteBack_DescribeShortfall = Description

End Function

Private Function M_WriteBack_TryResolveTableColumn( _
    ByRef TargetColumn As Range, _
    ByRef TableName As String, _
    ByRef ColumnName As String) As Boolean

'
'------------------------------------------------------------------------------
'                   TRY RESOLVE SELECTED TABLE COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the table data column owning the current selection, reporting
'   whether one was found rather than raising when none was
'
' WHY THIS EXISTS
'   A selection outside a table is an ordinary usage condition, not a failure.
'   DP_FillTableColumn needs to tell the user that plainly, and a raised runtime
'   error is the wrong instrument for it
'
'   Separating the expected negative from a genuine fault also keeps the two
'   distinguishable when the structured write result is added
'
' INPUTS
'   TargetColumn
'     Receives the resolved ListColumn data body range
'
'   TableName
'     Receives the owning ListObject name
'
'   ColumnName
'     Receives the resolved column name
'
' RETURNS
'   True when the selection is a single cell inside a table data body
'   False when it is a valid Range but not a table data cell
'
' BEHAVIOR
'   Requires a single-cell selection inside ListObject.DataBodyRange. A header
'   cell, a totals row cell, a multi-cell selection, or a cell outside any table
'   all return False
'
' ERROR POLICY
'   Raises only on genuinely abnormal conditions such as an unavailable or
'   non-Range selection. An ordinary non-table selection returns False
'
' DEPENDENCIES
'   Application.Selection
'   Excel.ListObject
'
' NOTES
'   A header cell is deliberately rejected. The command fills a
'   ListColumn.DataBodyRange, so the anchor must itself be inside that range
'
'   This mirrors the expansion rule used by M_WriteBack_ResolveTarget, which
'   also intersects against DataBodyRange rather than the whole table
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_WriteBack_TryResolveTableColumn"

    Dim SelectedObject      As Object           'Current Excel selection object
    Dim Anchor              As Range            'Single selected cell
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
    'Return the negative result unless a table column is resolved
        M_WriteBack_TryResolveTableColumn = False
        Set TargetColumn = Nothing
        TableName = VBA.vbNullString
        ColumnName = VBA.vbNullString

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
    'Use the selection as the candidate anchor
        Set Anchor = SelectedObject

'------------------------------------------------------------------------------
' REQUIRE A SINGLE ANCHOR CELL
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate anchor cell"
    'Return the negative result for a multi-cell selection
        If Anchor.Cells.CountLarge <> 1 Then Exit Function

'------------------------------------------------------------------------------
' RESOLVE OWNING TABLE COLUMN
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve owning table column"
    'Loop through worksheet tables
        For Each TargetTable In Anchor.Worksheet.ListObjects
            'Consider only tables that have a data body
                If Not TargetTable.DataBodyRange Is Nothing Then
                    'Consider only an anchor inside the table data body
                        If Not Application.Intersect(Anchor, TargetTable.DataBodyRange) Is Nothing Then
                            'Resolve the anchored table column index
                                ColumnIndex = Anchor.Column - TargetTable.DataBodyRange.Column + 1
                            'Return the resolved column when the index is valid
                                If ColumnIndex >= 1 Then
                                    If ColumnIndex <= TargetTable.ListColumns.Count Then
                                        Set TargetColumn = TargetTable.ListColumns(ColumnIndex).DataBodyRange
                                        TableName = TargetTable.Name
                                        ColumnName = TargetTable.ListColumns(ColumnIndex).Name
                                        M_WriteBack_TryResolveTableColumn = True
                                    End If
                                End If
                            'Stop after resolving the owning table
                                Exit For
                        End If
                End If
        Next TargetTable

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set TargetTable = Nothing
        Set Anchor = Nothing
        Set SelectedObject = Nothing
    'Exit the function
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "Table column resolution failed: " & Err.Description

End Function

Private Sub M_WriteBack_ResolveTarget( _
    ByRef ResolvedTarget As Range, _
    ByRef ExpandedToTableColumn As Boolean, _
    ByRef TableName As String, _
    ByRef ColumnName As String, _
    Optional ByVal NoTableGrow As Boolean = True)

'
'------------------------------------------------------------------------------
'                       RESOLVE WRITE-BACK TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the range the DatePicker would write to, without writing anything
'
' WHY THIS EXISTS
'   Resolving the target and mutating it were previously one routine, so nothing
'   could ask what the target would be before it was written
'
'   Separating them gives callers a seam: a confirmation prompt can describe the
'   resolved scope, and a structured write result can report it, without either
'   feature needing to re-derive the target
'
' INPUTS
'   ResolvedTarget
'     Receives the range that would be written
'
'   ExpandedToTableColumn
'     Receives True when a single selected table cell was expanded to its full
'     data column
'
'   TableName
'     Receives the owning ListObject name when expansion occurred
'
'   ColumnName
'     Receives the expanded column name when expansion occurred
'
'   NoTableGrow
'     True to keep a single selected table cell as a single-cell target
'
'     False to expand a single selected table data-body cell to the full table
'     data column
'
' RETURNS
'   Nothing. Results are returned through the ByRef arguments
'
' BEHAVIOR
'   Resolves the current Excel selection, rejects non-range selections,
'   optionally expands a single table data-body cell to its full ListObject data
'   column, and reports what it resolved
'
' ERROR POLICY
'   Raises a descriptive runtime error if the current Excel selection is not a
'   Range, or if the resolved target is missing or empty
'
' DEPENDENCIES
'   Application.Selection
'   Excel.Range
'   Excel.ListObject
'
' NOTES
'   This routine does not write, does not suppress Application events, and does
'   not prompt. It is safe to call to find out what a write would affect
'
'   It intentionally raises on non-Range selections so callers do not treat a
'   no-op as a successful resolution
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_WriteBack_ResolveTarget"

    Dim SelectedObject      As Object           'Current Excel selection object
    Dim Target              As Range            'Resolved target range
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
    'Return the safe defaults unless expansion occurs
        Set ResolvedTarget = Nothing
        ExpandedToTableColumn = False
        TableName = VBA.vbNullString
        ColumnName = VBA.vbNullString

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
                                                    ExpandedToTableColumn = True
                                                    TableName = TargetTable.Name
                                                    ColumnName = TargetTable.ListColumns(ColumnIndex).Name
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
' RETURN RESOLVED TARGET
'------------------------------------------------------------------------------
    'Return the resolved target to the caller
        Set ResolvedTarget = Target

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set TargetTable = Nothing
        Set Target = Nothing
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

Private Sub M_WriteBack_ApplyResolvedTarget( _
    ByVal Target As Range, _
    ByVal iType As DP_WriteAction, _
    ByVal OverwriteFormulas As Boolean, _
    ByRef Result As DP_WriteResult)

'
'------------------------------------------------------------------------------
'                       APPLY RESOLVED WRITE-BACK TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the requested DatePicker value to an already-resolved target and
'   accumulates one result describing the whole target
'
' WHY THIS EXISTS
'   Mutation is separated from resolution so a caller can inspect or confirm the
'   target before anything is written
'
' INPUTS
'   Target
'     Range resolved by M_WriteBack_ResolveTarget
'
'   iType
'     DatePicker write action to apply
'
'   OverwriteFormulas
'     False preserves a cell holding a formula. True replaces it. Passed through
'     to M_WriteBack_PopulateRange
'
'   Result
'     Accumulating DP_WriteResult, populated across every target area
'
' RETURNS
'   Nothing. The outcome is accumulated into Result
'
' BEHAVIOR
'   Validates the requested write action and the supplied target, writes to each
'   discontiguous target area, and reports the whole target address
'
' ERROR POLICY
'   Raises a descriptive runtime error if the write action is unsupported or the
'   target is missing or empty. Both are checked before any area is touched
'
'   An unexpected failure during area population is decided by
'   M_WriteBack_PopulateRange, which raises when no cell has produced an outcome
'   and returns a result carrying the failure when one has. A technical failure
'   therefore reaches this routine only when there is something worth returning,
'   and this routine returns it rather than converting it back into an exception
'
'   Population stops at the first technical failure. The remaining areas are
'   deliberately not attempted, so a technical-failure result is bounded
'   (Written + Locked + Formula + Failed <= Attempted) rather than balanced
'
' DEPENDENCIES
'   M_WriteBack_PopulateRange
'
' NOTES
'   This routine does not suppress Application events
'
'   Application.EnableEvents is managed by M_WriteBack_Apply
'
'   The write action is validated here rather than during resolution, so that
'   resolving a target for inspection does not require a valid write action
'
'   One result is accumulated across every area, so a discontiguous target
'   reports the whole write rather than its last area
'
' UPDATED
'   2026-11-09
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_WriteBack_ApplyResolvedTarget"

    Dim Block               As Range            'Single target area
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
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate supplied write-back target"
    'Reject a missing target
        If Target Is Nothing Then
            Err.Raise vbObjectError + 516, PROC_NAME, "Unable to resolve DatePicker write-back target"
        End If
    'Reject empty targets
        If Target.Cells.CountLarge = 0 Then
            Err.Raise vbObjectError + 517, PROC_NAME, "Resolved DatePicker write-back target is empty"
        End If

'------------------------------------------------------------------------------
' REPORT RESOLVED TARGET
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Report resolved target address"
    'Report the whole target rather than the last area written
        Result.ResolvedTargetAddress = Target.Worksheet.Name & "!" & _
            Target.Address(False, False)

'------------------------------------------------------------------------------
' POPULATE TARGET AREAS
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Populate target areas"
    'Reset the per-operation area counter that positions an injected fault. This
    'is test infrastructure and is inert unless a fault is armed
        mDP_TestWriteAreaOrdinal = 0
    'Loop through each discontiguous target area. Every area contributes its
    'attempted, written, skipped, failed and address data to Result before the
    'operation outcome is decided, so the totals do not depend on the order
    'Excel enumerates Target.Areas
        For Each Block In Target.Areas
            'Populate this target area into the accumulating result
                M_WriteBack_PopulateRange Block, iType, Result, OverwriteFormulas
            'Stop the operation on an unexpected technical failure. Continuing to
            'mutate the workbook after an error nobody classified would enlarge the
            'damage, so the remaining areas are deliberately not attempted. They are
            'therefore absent from AttemptedCount, which is what makes a
            'technical-failure result bounded rather than balanced
                If Result.TechnicalFailureOccurred Then Exit For
        Next Block

'------------------------------------------------------------------------------
' FINALIZE OPERATION OUTCOME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Finalize operation outcome"
    'Record a technical failure that survived to the operation level
    '
    'A failure reaches this point only when at least one cell had already produced
    'an outcome. M_WriteBack_PopulateRange raises instead when nothing had, so the
    'result being returned here always carries something worth returning. The
    'failure is not hidden: it travels in the result, is described by
    'M_WriteBack_DescribeShortfall, and is reported by M_WriteBack_ReportShortfall
        If Result.TechnicalFailureOccurred Then
            Debug.Print PROC_NAME & _
                " | Technical failure | Step=" & Result.TechnicalFailureStep & _
                "; Error=" & VBA.CStr(Result.TechnicalFailureNumber) & _
                "; Attempted=" & VBA.CStr(Result.AttemptedCount) & _
                "; Written=" & VBA.CStr(Result.WrittenCount) & _
                "; Areas=" & VBA.CStr(Result.AreasCount)
        End If
    'The outcome is decided here, once, on complete data. A legitimate zero-write
    'operation returns a complete DP_WriteResult rather than an exception, so a
    'caller can never be handed "exception with no result" after the workbook has
    'already been mutated by an earlier area. Callers report the shortfall through
    'M_WriteBack_ReportShortfall, which emits at most one operation-level message
        If Result.WrittenCount <= 0 Then
            'Record the zero-write operation for diagnostics
                Debug.Print PROC_NAME & _
                    " | Zero-write operation | Areas=" & VBA.CStr(Result.AreasCount) & _
                    "; Attempted=" & VBA.CStr(Result.AttemptedCount) & _
                    "; LockedSkipped=" & VBA.CStr(Result.LockedSkippedCount) & _
                    "; FormulaSkipped=" & VBA.CStr(Result.FormulaSkippedCount) & _
                    "; Failed=" & VBA.CStr(Result.FailedCount)
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Release object references
        Set Block = Nothing
    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Raise a descriptive error to the caller
        Err.Raise Err.Number, _
            PROC_NAME & " | Step=" & HandlerStep, _
            "DatePicker write-back failed: " & Err.Description

End Sub

Private Sub M_WriteBack_ResolveAndApplyTarget( _
    ByVal iType As DP_WriteAction, _
    ByVal NoTableGrow As Boolean, _
    ByVal OverwriteFormulas As Boolean, _
    ByRef Result As DP_WriteResult)

'
'------------------------------------------------------------------------------
'                       RESOLVE AND APPLY WRITE-BACK TARGET
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the current Excel write-back target, applies the requested DatePicker
'   write action, and reports both through one accumulating result
'
' WHY THIS EXISTS
'   Most callers want resolution and mutation as one step. This routine keeps
'   that convenience while the two stages remain separately callable
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
'   OverwriteFormulas
'     False preserves a cell holding a formula. True replaces it. Passed through
'     to M_WriteBack_ApplyResolvedTarget
'
'   Result
'     Accumulating DP_WriteResult, populated with the resolver metadata and the
'     write counts
'
' RETURNS
'   Nothing. The outcome is accumulated into Result
'
' BEHAVIOR
'   Delegates to M_WriteBack_ResolveTarget and then to
'   M_WriteBack_ApplyResolvedTarget, then attaches the resolver metadata to the
'   result rather than discarding it
'
' ERROR POLICY
'   Raises whatever the two stages raise, without adding a further wrapper
'
' DEPENDENCIES
'   M_WriteBack_ResolveTarget
'   M_WriteBack_ApplyResolvedTarget
'
' NOTES
'   The resolver signature is unchanged. This routine consumes the metadata the
'   resolver already produces instead of re-deriving it
'
'   NoTableGrow is required here because the only caller always supplies it. The
'   safe default lives on the public entry points, not on this private stage
'
' UPDATED
'   2026-11-09
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Target              As Range            'Resolved target range
    Dim Expanded            As Boolean          'True when expansion occurred
    Dim TableName           As String           'Owning table name when expanded
    Dim ColumnName          As String           'Expanded column name when expanded

'------------------------------------------------------------------------------
' RESOLVE TARGET
'------------------------------------------------------------------------------
    'Resolve the range that will receive the value
        M_WriteBack_ResolveTarget Target, Expanded, TableName, ColumnName, NoTableGrow

'------------------------------------------------------------------------------
' ATTACH RESOLVER METADATA
'------------------------------------------------------------------------------
    'Report whether the selection was expanded to a table data column
        Result.ExpandedToTableColumn = Expanded
    'Report the owning table when expansion occurred
        Result.TableName = TableName
    'Report the resolved column when expansion occurred
        Result.ColumnName = ColumnName

'------------------------------------------------------------------------------
' APPLY TO TARGET
'------------------------------------------------------------------------------
    'Write the value to the resolved range
        M_WriteBack_ApplyResolvedTarget Target, iType, OverwriteFormulas, Result

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
    'Release object references
        Set Target = Nothing

End Sub

Public Sub M_WriteBack_PopulateRange( _
    ByVal oRange As Range, _
    ByVal iType As DP_WriteAction, _
    ByRef Result As DP_WriteResult, _
    Optional ByVal OverwriteFormulas As Boolean = False)

'
'------------------------------------------------------------------------------
'                           POPULATE RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the current DatePicker value to every cell in a resolved Excel range
'   and accumulates the outcome into the supplied write result
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
'   Result
'     Accumulating DP_WriteResult. Counts and addresses for this range are added
'     to whatever the result already carries, so one result can describe a target
'     written one area at a time
'
'   OverwriteFormulas
'     False, the default, preserves a cell holding a formula and counts it in
'     FormulaSkippedCount, and refuses the bulk path when the range holds any
'     formula. True replaces formulas
'
' RETURNS
'   Nothing. The outcome is accumulated into Result
'
' BEHAVIOR
'   Validates the target range and write action, resolves the DatePicker write
'   value, attempts a fast bulk write for multi-cell ranges, and falls back to
'   safe cell-by-cell write-back when the bulk write cannot be completed
'
'   Protected locked cells may be skipped by the fallback path when at least one
'   target cell is written successfully
'
'   Formula cells are preserved unless the caller opted into replacing them
'
' ERROR POLICY
'   An area that writes no cell is an outcome, not an exception, and is accumulated
'   like any other
'
'   An unexpected error is decided by one rule: whether any cell, in this area or
'   in an earlier one, has already produced an outcome
'
'   If none has, the original error is raised. Nothing is destroyed by raising, and
'   an unsupported write action or a missing target reaching this routine is a
'   programming error that a direct caller of this Public routine must be told
'   about loudly
'
'   If one has, the error is recorded in TechnicalFailureOccurred,
'   TechnicalFailureStep, TechnicalFailureNumber and TechnicalFailureDescription,
'   the facts observed so far are accumulated, and the routine returns. A caller
'   whose workbook has already been mutated is never handed an exception carrying
'   no result. See #21
'
'   AttemptedCount is deliberately not part of that rule. It is recorded before the
'   first cell is touched, so it describes the size of the target rather than
'   anything that happened to it
'
'   Bulk-write failures are not raised directly because they are expected in
'   mixed protected, validated, or partially writable ranges. The routine falls
'   back to the existing per-cell write policy
'
'   A failed bulk write is not counted as a failed cell. The result describes the
'   outcome of the logical write, not the optimization attempts behind it
'
' DEPENDENCIES
'   M_WriteBack_GetPickedValue
'   M_WriteBack_TryBulkWriteRange
'   M_WriteBack_TryWriteCell
'
' NOTES
'   Counts are accumulated locally and added to Result once, so a raised error
'   cannot leave Result holding half of this range
'
'   WrittenCount counts the cells that reported a successful write. It is not
'   derived by subtracting the skips and failures from the attempted count,
'   because that derivation treats anything that did not raise as a success, and
'   Excel declines some writes without raising
'
'   The bulk path contributes its full target count explicitly. It returns before
'   the per-cell counters exist, so a result fed only by the fallback would report
'   a successful bulk write as nothing written
'
'   The bulk path is refused when any target cell belongs to an array formula.
'   A range assignment that overlaps an array neither raises nor writes, and a
'   range assignment that covers one replaces it silently, so the fast path would
'   report cells written that were not written. Range.HasArray returns Null for a
'   mixed target, and an unreadable array state is also treated as a refusal
'
'   The bulk path is refused again when formula protection is active and the
'   target holds any formula, because the bulk write never reaches per-cell
'   inspection and would destroy them. Range.HasFormula follows the same
'   True/False/Null convention, and Null is the mixed case that must refuse. A
'   single formula therefore disables the fast path for the whole target: that is
'   deliberate, and cheaper than a partition that could disagree with the write
'
'   This routine reports through Result and displays nothing. Deciding whether a
'   human is told about a partial write belongs to the entry point that was
'   invoked, which sees the whole operation rather than one area
'
'   This routine intentionally uses Range.Value rather than Range.Value2 so VBA
'   Date and DateTime values are written through Excel's normal date handling
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "M_WriteBack_PopulateRange"

    Dim AreaResult      As DP_WriteResult   'Outcome for this range only
    Dim Cell            As Range            'Current target cell
    Dim WriteValue      As Variant          'Resolved write value
    Dim ArrayState      As Variant          'Range.HasArray for the target
    Dim FormulaState    As Variant          'Range.HasFormula for the target
    Dim BulkAllowed     As Boolean          'True when the fast path may be used
    Dim AreaOrdinal     As Long             'Position of this area in the operation
    Dim HandlerStep     As String           'Current handler step for diagnostics
    Dim PriorLockedSkippedCount As Double   'Operation locked count before this area
    Dim PriorFormulaSkippedCount As Double  'Operation formula count before this area
    Dim PriorFailedCount As Double          'Operation failure count before this area

    Dim SavedErrNumber      As Long         'Captured original error number
    Dim SavedErrDescription As String       'Captured original error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler
    'Initialize diagnostic step
        HandlerStep = "Initialize"
    'Seed category counters from the operation result before this area starts.
    'The per-cell classifier can then keep using AreaResult while its RecordedCount
    'is operation-global, so a new target area cannot restart the 25-address budget.
        PriorLockedSkippedCount = Result.LockedSkippedCount
        PriorFormulaSkippedCount = Result.FormulaSkippedCount
        PriorFailedCount = Result.FailedCount
        AreaResult.LockedSkippedCount = PriorLockedSkippedCount
        AreaResult.FormulaSkippedCount = PriorFormulaSkippedCount
        AreaResult.FailedCount = PriorFailedCount

'------------------------------------------------------------------------------
' RECORD AREA POSITION
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Record area position"
    'Record where this area sits in the operation. The counter is reset by
    'M_WriteBack_ApplyResolvedTarget before the area loop, so the ordinal is
    'per-operation rather than per-session
        mDP_TestWriteAreaOrdinal = mDP_TestWriteAreaOrdinal + 1
        AreaOrdinal = mDP_TestWriteAreaOrdinal
    'Raise an armed fault that fires on entering this area, before anything about
    'it has been recorded and before any of its cells are mutated. This is the one
    'position from which an operation can still end up carrying no facts at all,
    'which is the case that must still raise
    '
    'Disarmed in production: see M_WriteBack_Test_SetFaultInjection
        If M_WriteBack_TestFaultShouldFire(AreaOrdinal, 0) Then
            Err.Raise vbObjectError + 518, PROC_NAME, _
                "Injected write-back fault on entering area " & VBA.CStr(AreaOrdinal)
        End If

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
        AreaResult.AttemptedCount = oRange.Cells.CountLarge
    'Reject empty target ranges
        If AreaResult.AttemptedCount <= 0 Then
            Err.Raise vbObjectError + 514, PROC_NAME, "Target range does not contain writable cells"
        End If
    'Capture the discontiguous areas this range covers
        AreaResult.AreasCount = oRange.Areas.Count

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
        BulkAllowed = (AreaResult.AttemptedCount > 1)
    'Resolve whether the target touches an array formula
        If BulkAllowed Then
            'Treat an unreadable array state as a reason to refuse the fast path
                On Error Resume Next
                ArrayState = oRange.HasArray
                If Err.Number <> 0 Then
                    ArrayState = Null
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            'Refuse the fast path when any target cell belongs to an array formula
                If VBA.IsNull(ArrayState) Then
                    BulkAllowed = False
                ElseIf VBA.CBool(ArrayState) Then
                    BulkAllowed = False
                End If
        End If
    'Resolve whether the target holds formulas that policy protects
        If BulkAllowed And Not OverwriteFormulas Then
            'Treat an unreadable formula state as a reason to refuse the fast path
                On Error Resume Next
                FormulaState = oRange.HasFormula
                If Err.Number <> 0 Then
                    FormulaState = Null
                    Err.Clear
                End If
                On Error GoTo ErrorHandler
            'Refuse the fast path when any target cell holds a formula. Null is the
            'mixed case and must refuse, not be coerced to False
                If VBA.IsNull(FormulaState) Then
                    BulkAllowed = False
                ElseIf VBA.CBool(FormulaState) Then
                    BulkAllowed = False
                End If
        End If
    'Account for the whole range and exit when the fast bulk write succeeds
        If BulkAllowed Then
            If M_WriteBack_TryBulkWriteRange(oRange, WriteValue) Then
                'The bulk write covered every target cell
                    AreaResult.WrittenCount = AreaResult.AttemptedCount
                'Skip the per-cell fallback
                    GoTo AccumulateResult
            End If
        End If

'------------------------------------------------------------------------------
' FALL BACK TO SAFE CELL-BY-CELL WRITE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Populate target cells through safe fallback"
    'Loop through each target cell
        For Each Cell In oRange.Cells
            'Count only the cells that actually received the value
                If M_WriteBack_TryWriteCell(Cell, WriteValue, OverwriteFormulas, AreaResult) Then
                    AreaResult.WrittenCount = AreaResult.WrittenCount + 1
                End If
            'Raise an armed fault once the requested number of cells in this area
            'has been mutated. Disarmed in production: see
            'M_WriteBack_Test_SetFaultInjection
                If M_WriteBack_TestFaultShouldFire(AreaOrdinal, AreaResult.WrittenCount) Then
                    Err.Raise vbObjectError + 518, PROC_NAME, _
                        "Injected write-back fault after " & _
                        VBA.CStr(AreaResult.WrittenCount) & _
                        " written cells in area " & VBA.CStr(AreaOrdinal)
                End If
        Next Cell

'------------------------------------------------------------------------------
' RESOLVE WRITE RESULT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve write result"
    'A zero-write area is an outcome, not an exception, and is accumulated like
    'any other. At v1.2.0 this raised before the accumulation below, so an area
    'that wrote nothing discarded its own classification counts and every area
    'already accumulated ahead of it. A caller that had just mutated an earlier
    'area received an exception carrying no DP_WriteResult at all, and the totals
    'depended on the order Excel happened to enumerate Target.Areas
    '
    'The operation-level outcome is decided by M_WriteBack_ApplyResolvedTarget
    'once every area has been accounted for. Do not reinstate a per-area raise
    'here: see #21
        If AreaResult.WrittenCount <= 0 Then
            'Record the zero-write area for diagnostics without interrupting the
            'operation or discarding what the other areas observed
                Debug.Print PROC_NAME & _
                    " | Zero-write area | Attempted=" & VBA.CStr(AreaResult.AttemptedCount) & _
                    "; LockedSkipped=" & VBA.CStr(AreaResult.LockedSkippedCount - PriorLockedSkippedCount) & _
                    "; FormulaSkipped=" & VBA.CStr(AreaResult.FormulaSkippedCount - PriorFormulaSkippedCount) & _
                    "; Failed=" & VBA.CStr(AreaResult.FailedCount - PriorFailedCount)
        End If

'------------------------------------------------------------------------------
' ACCUMULATE RESULT
'------------------------------------------------------------------------------
AccumulateResult:
    'Track the current handler step
        HandlerStep = "Accumulate write result"
    'Add this range to the running totals
        Result.AttemptedCount = Result.AttemptedCount + AreaResult.AttemptedCount
        Result.WrittenCount = Result.WrittenCount + AreaResult.WrittenCount
        Result.LockedSkippedCount = AreaResult.LockedSkippedCount
        Result.FormulaSkippedCount = AreaResult.FormulaSkippedCount
        Result.FailedCount = AreaResult.FailedCount
        Result.AreasCount = Result.AreasCount + AreaResult.AreasCount
    'Join the skipped locked addresses
        If VBA.LenB(AreaResult.LockedSkippedAddresses) > 0 Then
            If VBA.LenB(Result.LockedSkippedAddresses) = 0 Then
                Result.LockedSkippedAddresses = AreaResult.LockedSkippedAddresses
            Else
                Result.LockedSkippedAddresses = Result.LockedSkippedAddresses & ", " & _
                    AreaResult.LockedSkippedAddresses
            End If
        End If
    'Join the preserved formula addresses
        If VBA.LenB(AreaResult.FormulaSkippedAddresses) > 0 Then
            If VBA.LenB(Result.FormulaSkippedAddresses) = 0 Then
                Result.FormulaSkippedAddresses = AreaResult.FormulaSkippedAddresses
            Else
                Result.FormulaSkippedAddresses = Result.FormulaSkippedAddresses & ", " & _
                    AreaResult.FormulaSkippedAddresses
            End If
        End If
    'Join the failed addresses
        If VBA.LenB(AreaResult.FailedAddresses) > 0 Then
            If VBA.LenB(Result.FailedAddresses) = 0 Then
                Result.FailedAddresses = AreaResult.FailedAddresses
            Else
                Result.FailedAddresses = Result.FailedAddresses & ", " & _
                    AreaResult.FailedAddresses
            End If
        End If
    'Carry an unexpected technical failure into the operation result. The first
    'failure wins, so the original error survives any later area
        If AreaResult.TechnicalFailureOccurred Then
            If Not Result.TechnicalFailureOccurred Then
                Result.TechnicalFailureOccurred = True
                Result.TechnicalFailureStep = AreaResult.TechnicalFailureStep
                Result.TechnicalFailureNumber = AreaResult.TechnicalFailureNumber
                Result.TechnicalFailureDescription = AreaResult.TechnicalFailureDescription
            End If
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
    'Capture the original error before anything can disturb it. The identity of
    'this error is preserved all the way to the operation-level decision, so a
    'caller that is told about a technical failure is told which one
        SavedErrNumber = Err.Number
        SavedErrDescription = "DatePicker range population failed: " & Err.Description
    'Release object references
        Set Cell = Nothing
    'Refuse to re-enter. Accumulation below is arithmetic and string joining and
    'cannot realistically fail, but a handler that resumes into a block able to
    'raise back into the same handler would spin. A second entry raises instead
        If AreaResult.TechnicalFailureOccurred Then
            Err.Raise SavedErrNumber, _
                PROC_NAME & " | Step=" & HandlerStep, SavedErrDescription
        End If
    'Raise when there is nothing to lose
    '
    'One rule decides this, and it turns on whether any cell has produced an
    'outcome. If none has, no work is destroyed by raising, and the caller is
    'better served by a loud failure: an unsupported write action or a missing
    'target reaching this routine is a programming error, and this is the only
    'signal a direct caller of this Public routine would get
    '
    'AttemptedCount is deliberately not consulted. It is recorded before the first
    'cell is touched, so it says how large the target was, not that anything
    'happened to it
        If Not M_WriteBack_HasCellOutcomes(Result) Then
            If Not M_WriteBack_HasCellOutcomes(AreaResult) Then
                Err.Raise SavedErrNumber, _
                    PROC_NAME & " | Step=" & HandlerStep, SavedErrDescription
            End If
        End If
    'Record the failure as an observed fact about this area rather than losing it
    '
    'Once a cell has produced an outcome this handler does not raise. At v1.2.0 it
    'always did, and because the accumulation block below sits after it, an
    'unexpected error discarded every cell this area had already written and every
    'area accumulated before it. A caller whose workbook had just been mutated
    'received an exception carrying no DP_WriteResult at all
    '
    'Do not reinstate an unconditional raise here: see #21
        AreaResult.TechnicalFailureOccurred = True
        AreaResult.TechnicalFailureStep = HandlerStep
        AreaResult.TechnicalFailureNumber = SavedErrNumber
        AreaResult.TechnicalFailureDescription = SavedErrDescription
    'Record the failure for the developer
        Debug.Print PROC_NAME & _
            " | Technical failure | Step=" & HandlerStep & _
            "; Error=" & VBA.CStr(SavedErrNumber) & _
            "; WrittenInArea=" & VBA.CStr(AreaResult.WrittenCount) & _
            "; " & SavedErrDescription
    'Accumulate what this area observed before returning to the caller
        Resume AccumulateResult

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
    ByVal OverwriteFormulas As Boolean, _
    ByRef Result As DP_WriteResult) As Boolean

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
'   through a Boolean return value plus the accumulating write result
'
' INPUTS
'   TargetCell
'     Single Excel cell to populate
'
'   WriteValue
'     DatePicker value to write
'
'   OverwriteFormulas
'     False preserves a cell holding a formula. True replaces it
'
'   Result
'     Accumulating DP_WriteResult. Its skip and failure counters are incremented
'     in place and the corresponding cell addresses are recorded
'
' RETURNS
'   True when the value was successfully written to TargetCell
'
'   False when TargetCell was missing, invalid, protected locked, or could not be
'   written
'
' BEHAVIOR
'   Validates the target cell, skips protected locked cells, refuses cells that
'   belong to an array formula, preserves formula cells unless the caller opted
'   into replacing them, writes the supplied value to writable cells, returns True
'   only after a successful write, records the address behind every skip and
'   failure, and logs suppressed write failures to the Immediate Window
'
' ERROR POLICY
'   Best-effort per-cell write
'
'   Does not raise outward. Cell-level failures are counted and suppressed so the
'   caller can complete the range write-back and decide whether the aggregate
'   result is acceptable
'
'   Every cell increments exactly one of the written, locked or failed counts, so
'   the caller's result satisfies its accounting invariant by construction
'
' DEPENDENCIES
'   Excel.Range
'   Excel.Worksheet.ProtectContents
'   M_WriteBack_AppendAddress
'
' NOTES
'   This routine intentionally uses Range.Value rather than Range.Value2 so VBA
'   Date and DateTime variants are written through Excel's normal date handling
'
'   Application.EnableEvents is managed by M_WriteBack_Apply, not by this routine
'
'   Recorded addresses are worksheet-qualified, in the form SheetName!A1, so a
'   reported address stays unambiguous and is stable enough to assert against
'
'   A cell belonging to an array formula is refused before the write rather than
'   after it. Excel raises "You cannot change part of an array" for an
'   interactive edit but declines the same assignment silently through the object
'   model, so an attempted write would return success having changed nothing
'
'   The array gate runs before the formula gate deliberately. An array cell cannot
'   be written at all, which is a stronger and non-overridable condition than a
'   formula the caller could choose to replace. Reporting it as a failure stays
'   correct whichever way OverwriteFormulas is set
'
'   Every cell increments exactly one of the written, locked, formula-skipped or
'   failed counts, so the caller's accounting invariant holds by construction
'
'   Counters stay exact. Only the reported address lists are capped
'
' UPDATED
'   2026-08-23
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
            Result.FailedCount = Result.FailedCount + 1
            M_WriteBack_AppendAddress Result.FailedAddresses, Result.FailedCount, "(no cell)"
            Debug.Print PROC_NAME & ": skipped missing target cell"
            Exit Function
        End If
    'Count and exit when a non-single-cell range is supplied unexpectedly
        If TargetCell.Cells.CountLarge <> 1 Then
            Result.FailedCount = Result.FailedCount + 1
            M_WriteBack_AppendAddress Result.FailedAddresses, Result.FailedCount, _
                TargetCell.Worksheet.Name & "!" & TargetCell.Address(False, False)
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
                Result.LockedSkippedCount = Result.LockedSkippedCount + 1
                M_WriteBack_AppendAddress Result.LockedSkippedAddresses, _
                    Result.LockedSkippedCount, TargetSheetName & "!" & TargetAddress
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' SKIP CELLS BELONGING TO AN ARRAY FORMULA
'------------------------------------------------------------------------------
    'Excel accepts a value assignment to an array cell without raising and without
    'writing, so attempting one here would be counted as a successful write
        If TargetCell.HasArray Then
            Result.FailedCount = Result.FailedCount + 1
            M_WriteBack_AppendAddress Result.FailedAddresses, Result.FailedCount, _
                TargetSheetName & "!" & TargetAddress
            Debug.Print PROC_NAME & ": skipped array-formula cell " & _
                TargetSheetName & "!" & TargetAddress
            Exit Function
        End If

'------------------------------------------------------------------------------
' PRESERVE FORMULA CELLS
'------------------------------------------------------------------------------
    'Leave a formula in place unless the caller explicitly opted into replacing it.
    'A formula that evaluates to a date is still a formula: the user may mean to
    'replace the displayed date without meaning to delete what produced it
        If Not OverwriteFormulas Then
            If TargetCell.HasFormula Then
                Result.FormulaSkippedCount = Result.FormulaSkippedCount + 1
                M_WriteBack_AppendAddress Result.FormulaSkippedAddresses, _
                    Result.FormulaSkippedCount, TargetSheetName & "!" & TargetAddress
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
        Result.FailedCount = Result.FailedCount + 1
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
    'Record the address behind this failure
        M_WriteBack_AppendAddress Result.FailedAddresses, Result.FailedCount, _
            TargetSheetName & "!" & TargetAddress
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


'
'------------------------------------------------------------------------------
'
'                             LOADED FORM HELPERS
'
'------------------------------------------------------------------------------

Public Function M_FormBridge_GetLoadedForm(ByVal UserFormName As String) As Object

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
'   Public only so M_DP_Timer can reach the loaded picker; Option Private
'   Module keeps it inside the project (#57)
'
' UPDATED
'   2026-11-21
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
'   Before this reuse path, DP_Show unloaded and rebuilt the UserForm on every
'   call. That was robust, but slower than necessary when the picker is already
'   visible.
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
'   Unloads any already-loaded DatePicker UserForm instance
'
'   CURRENTLY UNUSED. Nothing in this project calls it. DP_Show, which it was
'   written for, now reuses or refreshes a loaded form instead of unloading it,
'   and the routine is Private, so nothing outside this module can reach it
'
' WHY THIS EXISTS
'   DP_Show used to rely on UserForm_Initialize to rebuild the runtime UI and
'   consume the pending initial-date bridge state, so a previously loaded
'   instance had to be removed before a new one was loaded
'
'   The helper must avoid direct references to the DatePicker default instance
'   because such references can instantiate the form accidentally
'
'   It is kept unchanged while its removal is open in #182
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
'   A caller must not fail just because an old transient form instance or stale
'   timer cannot be cleaned up fully
'
' UPDATED
'   2026-11-09
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


Public Sub M_Settings_UseLocalNames_Impl(ByVal UseLocalDayNames As Boolean)

'
'------------------------------------------------------------------------------
'                           SETTINGS USE LOCAL NAMES
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Settings_UseLocalNames
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Settings_UseLocalNames
'
' RETURNS
'   As documented on M_DatePicker.M_Settings_UseLocalNames
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Settings_UseLocalNames
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Settings_UseLocalNames. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   M_Settings_SetUseLocalNames
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE SETTING UPDATE
'------------------------------------------------------------------------------
    'Delegate to the preferred setting routine
        M_Settings_SetUseLocalNames_Impl UseLocalDayNames

End Sub

'
'------------------------------------------------------------------------------
'
'                               RIBBON CALLBACKS
'
'------------------------------------------------------------------------------
'


Public Sub Ribbon_ShowPicker_Impl(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON SHOW PICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.Ribbon_ShowPicker
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.Ribbon_ShowPicker
'
' RETURNS
'   As documented on M_DatePicker.Ribbon_ShowPicker
'
' BEHAVIOR
'   As documented on M_DatePicker.Ribbon_ShowPicker
'
' ERROR POLICY
'   As documented on M_DatePicker.Ribbon_ShowPicker. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   IRibbonControl
'   DP_Show
'   Ribbon_ReportError
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
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
        DP_Show_Impl

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

Public Sub Ribbon_Reset_Impl(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON RESET DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.Ribbon_Reset
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.Ribbon_Reset
'
' RETURNS
'   As documented on M_DatePicker.Ribbon_Reset
'
' BEHAVIOR
'   As documented on M_DatePicker.Ribbon_Reset
'
' ERROR POLICY
'   As documented on M_DatePicker.Ribbon_Reset. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   IRibbonControl
'   DP_Start
'   DP_RepairRuntime
'   M_Ribbon_ResetShouldStart
'   M_Ribbon_LifecycleSucceeded
'   Ribbon_ReportInfo
'   Ribbon_ReportError
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_Reset"

    Dim StartInstead    As Boolean      'True to start rather than repair
    Dim OperationName   As String       'Lifecycle operation this callback ran

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' START OR REPAIR RUNTIME
'------------------------------------------------------------------------------
    'Start when no provider holds the lease; repair otherwise
        StartInstead = M_Ribbon_ResetShouldStart( _
            M_Lease_IsOwner(), (VBA.LenB(M_Lease_ReadOwner()) = 0))
        If StartInstead Then
            OperationName = "DP_Start"
            DP_Start_Impl
        Else
            OperationName = "DP_RepairRuntime"
            DP_RepairRuntime_Impl
        End If

'------------------------------------------------------------------------------
' CONFIRM SUCCESS
'------------------------------------------------------------------------------
    'Say nothing more after a refusal: it has already been reported
        If Not M_Ribbon_LifecycleSucceeded(OperationName, _
            mDP_LifecycleLastOperation, mDP_LifecycleLastSucceeded) Then
            Exit Sub
        End If
    'Report the successful start or repair to the user
        If StartInstead Then
            Ribbon_ReportInfo _
                "Date / Time Picker runtime started." & _
                VBA.vbCrLf & VBA.vbCrLf & _
                "You can now select a date cell or click Show Picker again."
        Else
            Ribbon_ReportInfo _
                "Date / Time Picker runtime repair completed successfully." & _
                VBA.vbCrLf & VBA.vbCrLf & _
                "You can now select a date cell or click Show Picker again."
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
    'Report the Ribbon callback failure
        Ribbon_ReportError PROC_NAME, Err.Number, Err.Description

End Sub

Public Function M_Ribbon_ResetShouldStart( _
    ByVal IsLeaseOwner As Boolean, _
    ByVal LeaseIsFree As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                  DECIDE WHETHER RIBBON RESET STARTS OR REPAIRS
'------------------------------------------------------------------------------
' PURPOSE
'   Decides whether the Ribbon Reset command starts the runtime or repairs it
'
' WHY THIS EXISTS
'   DP_RepairRuntime refuses unless this project owns the lease, so Reset could
'   not recover a copy that had never started. When no provider holds the lease,
'   starting is exactly what DP_Start would admit anyway (#113)
'
' INPUTS
'   IsLeaseOwner
'     True when this project owns the provider lease
'
'   LeaseIsFree
'     True when no provider lease exists
'
' RETURNS
'   True to start with DP_Start, False to repair with DP_RepairRuntime
'
' BEHAVIOR
'   Starts only when this project is not the owner and the lease is free. An
'   owner repairs. A lease held by anyone else goes to repair, which refuses it
'
' ERROR POLICY
'   Cannot raise. Boolean logic only
'
' DEPENDENCIES
'   None
'
' NOTES
'   Ribbon_Reset is the only production consumer. Public only because the
'   regression harness is a separate module. It takes arguments, so it does not
'   appear in the macro dialog, and it is internal rather than supported API
'
' UPDATED
'   2026-10-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECIDE
'------------------------------------------------------------------------------
    'Start only a session that no provider holds
        M_Ribbon_ResetShouldStart = (Not IsLeaseOwner) And LeaseIsFree

End Function

Public Function M_Ribbon_LifecycleSucceeded( _
    ByVal ExpectedOperation As String, _
    ByVal ObservedOperation As String, _
    ByVal ObservedSucceeded As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'               DECIDE WHETHER A RIBBON LIFECYCLE ACTION SUCCEEDED
'------------------------------------------------------------------------------
' PURPOSE
'   Decides whether a Ribbon command may report success for the lifecycle
'   operation it just ran
'
' WHY THIS EXISTS
'   DP_Start and DP_RepairRuntime return normally when they refuse, so "the call
'   returned" is not "the call succeeded". Ribbon_Reset reported success after a
'   refused repair (#89). Both operations already record their outcome
'
' INPUTS
'   ExpectedOperation
'     Name of the lifecycle operation the caller ran
'
'   ObservedOperation
'     Operation name recorded by the lifecycle observation fields
'
'   ObservedSucceeded
'     Success flag recorded by the lifecycle observation fields
'
' RETURNS
'   True only when the recorded operation is the expected one and it succeeded
'
' BEHAVIOR
'   Trusts the success flag only after confirming it belongs to the operation
'   the caller ran, so a stale record from an earlier operation cannot pass
'
' ERROR POLICY
'   Cannot raise. String comparison and Boolean logic only
'
' DEPENDENCIES
'   None
'
' NOTES
'   Ribbon_Reset is the only production consumer. Public only because the
'   regression harness is a separate module. It takes arguments, so it does not
'   appear in the macro dialog, and it is internal rather than supported API
'
' UPDATED
'   2026-10-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECIDE
'------------------------------------------------------------------------------
    'A success flag counts only for the operation the caller actually ran
        M_Ribbon_LifecycleSucceeded = ObservedSucceeded And _
            (VBA.StrComp(ObservedOperation, ExpectedOperation, vbBinaryCompare) = 0)

End Function

Public Function M_DemoSheet_ResolveShowOnToggle( _
    ByVal ExistedBefore As Boolean, _
    ByVal VisibleBefore As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                     RESOLVE THE DEMO SHEET TOGGLE ACTION
'------------------------------------------------------------------------------
' PURPOSE
'   Decides whether the Ribbon demo command should show the demo sheet or hide it
'
' WHY THIS EXISTS
'   Ribbon_Demo read the sheet's visibility after ensuring it existed. Ensuring it
'   builds it visible on first use, so the very first click built the demo sheet
'   and immediately hid it again, and the command appeared to do nothing
'
'   The decision depends on state that exists only before the sheet is ensured,
'   so it cannot be read from the sheet afterwards. Isolating it here gives the
'   regression a deterministic target without pulling the demo builder into the
'   harness
'
' INPUTS
'   ExistedBefore
'     True when the demo sheet already existed before this command ran
'
'   VisibleBefore
'     True when that pre-existing sheet was visible. Ignored when ExistedBefore
'     is False
'
' RETURNS
'   True to show and activate the sheet, False to hide it
'
' BEHAVIOR
'   A sheet that did not exist has just been built and is always shown. A sheet
'   that existed toggles on the visibility it had before the command ran
'
' ERROR POLICY
'   Cannot raise. Boolean logic only
'
' DEPENDENCIES
'   None
'
' NOTES
'   xlSheetHidden and xlSheetVeryHidden are the same pre-existing state for this
'   decision. The caller collapses both to VisibleBefore = False, so a demo sheet
'   hidden either way is shown by the next click
'
'   Ribbon_Demo is the only production consumer. The regression asserts through
'   this routine rather than restating the condition, so a change here cannot
'   pass a test that no longer describes the callback
'
'   Public only because the regression harness is a separate module. It takes
'   arguments, so it does not appear in the macro dialog, and #25 classifies it
'   as internal rather than supported API
'
' UPDATED
'   2026-09-17
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESOLVE THE ACTION
'------------------------------------------------------------------------------
    'A sheet this command just built is always shown, never hidden
        If Not ExistedBefore Then
            M_DemoSheet_ResolveShowOnToggle = True
            Exit Function
        End If
    'A pre-existing sheet toggles on the visibility it had beforehand
        M_DemoSheet_ResolveShowOnToggle = Not VisibleBefore

End Function

Public Sub Ribbon_Demo_Impl(ByVal control As IRibbonControl)

'
'------------------------------------------------------------------------------
'                           RIBBON TOGGLE DEMO SHEET
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.Ribbon_Demo
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.Ribbon_Demo
'
' RETURNS
'   As documented on M_DatePicker.Ribbon_Demo
'
' BEHAVIOR
'   As documented on M_DatePicker.Ribbon_Demo
'
' ERROR POLICY
'   As documented on M_DatePicker.Ribbon_Demo. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
'
' DEPENDENCIES
'   IRibbonControl
'   DP_DemoSheet_ResolveHostWorkbook
'   DP_Demo_EnsureDemoSheet
'   DP_DEMO_SHEET_NAME
'   M_DemoSheet_ResolveShowOnToggle
'   DP_DemoSheet_HideVeryHidden
'   Ribbon_ReportError
'
' NOTES
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME As String = "Ribbon_Demo"

    Dim HostBook    As Excel.Workbook        'Workbook that holds the demo sheet
    Dim DemoSheet   As Excel.Worksheet       'DatePicker demo worksheet
    Dim ExistingSheet As Excel.Worksheet     'Demo sheet as it stood before this command
    Dim ExistedBefore As Boolean             'True when the demo sheet already existed
    Dim VisibleBefore As Boolean             'True when that pre-existing sheet was visible

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE HOST WORKBOOK
'------------------------------------------------------------------------------
    'Resolve the workbook that should hold the demo sheet, creating one when the
    'component runs as an add-in and no open workbook already holds it
        Set HostBook = DP_DemoSheet_ResolveHostWorkbook(True)

    'Reject the case where no host workbook could be resolved
        If HostBook Is Nothing Then
            Err.Raise vbObjectError + 540, PROC_NAME, _
                "No workbook is available to hold the demo sheet."
        End If

'------------------------------------------------------------------------------
' RESOLVE DEMO SHEET
'------------------------------------------------------------------------------
    'Record whether the demo sheet already existed, and whether it was visible,
    'before anything can create it. Ensuring the sheet builds it visible, so a
    'visibility read taken afterwards cannot distinguish a sheet this command
    'just built from one the user had already opened
        On Error Resume Next
        Set ExistingSheet = HostBook.Worksheets(DP_DEMO_SHEET_NAME)
        Err.Clear
    'Restore this procedure's own handler after its own error-mode change
        On Error GoTo ErrorHandler
        ExistedBefore = Not (ExistingSheet Is Nothing)
    'Collapse xlSheetHidden and xlSheetVeryHidden to one hidden state
        If ExistedBefore Then
            VisibleBefore = (ExistingSheet.Visible = xlSheetVisible)
        End If
        Set ExistingSheet = Nothing
    'Build the demo sheet on first use, then return it
        Set DemoSheet = DP_Demo_EnsureDemoSheet(HostBook)

'------------------------------------------------------------------------------
' TOGGLE DEMO SHEET
'------------------------------------------------------------------------------
    'Apply the action resolved from the state captured before the sheet was
    'ensured, so a sheet this command just built is shown rather than hidden
        If M_DemoSheet_ResolveShowOnToggle(ExistedBefore, VisibleBefore) Then
            DemoSheet.Visible = xlSheetVisible
            DemoSheet.Activate
        Else
            DP_DemoSheet_HideVeryHidden
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
'   DP_DemoSheet_ResolveHostWorkbook
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
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HostBook        As Excel.Workbook        'Workbook that holds the demo sheet
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
    'Resolve the workbook that holds the demo sheet
        Set HostBook = DP_DemoSheet_ResolveHostWorkbook(False)

    'Exit if no open workbook holds a demo sheet
        If HostBook Is Nothing Then GoTo CleanExit

    'Retrieve the demo worksheet from the resolved workbook
        Set DemoSheet = HostBook.Worksheets(DP_DEMO_SHEET_NAME)

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

Private Function DP_DemoSheet_ResolveHostWorkbook( _
    ByVal CreateWhenMissing As Boolean) As Excel.Workbook

'
'------------------------------------------------------------------------------
'                       RESOLVE DEMO HOST WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the workbook that holds, or should hold, the DatePicker demo sheet
'
' WHY THIS EXISTS
'   The demo routines previously resolved the sheet from ThisWorkbook. That is
'   correct when the component is embedded in the demo workbook and wrong when it
'   is loaded as an add-in, where ThisWorkbook is the add-in and has no
'   worksheets at all
'
'   Centralizing the decision keeps the four demo routines in agreement about
'   which workbook they are operating on
'
' INPUTS
'   CreateWhenMissing
'     True to add a new workbook when running as an add-in and no open workbook
'     already holds the demo sheet
'
' RETURNS
'   Workbook that holds or will hold the demo sheet, or Nothing
'
' BEHAVIOR
'   Returns ThisWorkbook when the component is embedded
'
'   When running as an add-in, returns the first open workbook that already holds
'   the demo sheet. Adds a new workbook when none does and CreateWhenMissing is
'   True. Otherwise returns Nothing
'
' ERROR POLICY
'   Safe default. Returns Nothing rather than raising when no workbook can be
'   resolved, so lifecycle cleanup paths can exit quietly
'
' DEPENDENCIES
'   DP_DEMO_SHEET_NAME
'
' NOTES
'   The add-in deliberately does not build the demo into whichever workbook
'   happens to be active. Adding an unrequested sheet to a user's live workbook
'   is a worse outcome than opening a new one
'
'   An add-in, a workbook with protected structure, and a workbook already
'   holding the demo are all distinguished here rather than left to fail later
'
' UPDATED
'   2026-08-22
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CandidateBook   As Excel.Workbook       'Workbook scan variable
    Dim ProbeSheet      As Excel.Worksheet      'Demo sheet probe

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return the safe default unless a workbook can be resolved
        Set DP_DemoSheet_ResolveHostWorkbook = Nothing
    'Suppress resolution errors through the local fail-safe path
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' EMBEDDED DEPLOYMENT
'------------------------------------------------------------------------------
    'Use the host workbook when the component is embedded in one
        If Not ThisWorkbook.IsAddin Then
            Set DP_DemoSheet_ResolveHostWorkbook = ThisWorkbook
            Exit Function
        End If

'------------------------------------------------------------------------------
' REUSE AN OPEN DEMO WORKBOOK
'------------------------------------------------------------------------------
    'Return the first open workbook that already holds the demo sheet
        For Each CandidateBook In Excel.Application.Workbooks
            'Skip add-ins, which cannot hold a worksheet
                If Not CandidateBook.IsAddin Then
                    'Probe for the demo sheet without raising when it is absent
                        Set ProbeSheet = Nothing
                        On Error Resume Next
                        Set ProbeSheet = CandidateBook.Worksheets(DP_DEMO_SHEET_NAME)
                        Err.Clear
                        On Error GoTo FailSafe
                    'Return the workbook that already holds the demo sheet
                        If Not ProbeSheet Is Nothing Then
                            Set DP_DemoSheet_ResolveHostWorkbook = CandidateBook
                            Exit Function
                        End If
                End If
        Next CandidateBook

'------------------------------------------------------------------------------
' CREATE A DEMO WORKBOOK
'------------------------------------------------------------------------------
    'Exit with the safe default when the caller does not want one created
        If Not CreateWhenMissing Then Exit Function

    'Add a workbook to hold the demo rather than writing into the user's own
        Set DP_DemoSheet_ResolveHostWorkbook = Excel.Application.Workbooks.Add

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the fail-safe handler
        Exit Function

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Return the safe default
        Set DP_DemoSheet_ResolveHostWorkbook = Nothing
    'Clear the suppressed resolution error
        Err.Clear

End Function

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
'   First visible non-demo worksheet in the workbook that owns the demo sheet
'   Nothing when no suitable worksheet exists
'
' BEHAVIOR
'   Scans the worksheets of the workbook that owns the supplied demo sheet and
'   returns the first visible worksheet that
'   is not the supplied demo worksheet
'
' ERROR POLICY
'   Best-effort lookup. Returns Nothing on error
'
' DEPENDENCIES
'   Worksheet.Parent
'   Excel.Worksheet
'
' NOTES
'   This helper does not create sheets and does not change visibility
'
' UPDATED
'   2026-08-22
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
    'Loop through worksheets in the workbook that owns the demo sheet
        For Each WS In DemoSheet.Parent.Worksheets
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
