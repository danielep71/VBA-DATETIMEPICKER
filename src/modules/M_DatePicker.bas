Attribute VB_Name = "M_DatePicker"

Option Explicit

'
'------------------------------------------------------------------------------
' MODULE: M_DatePicker
'------------------------------------------------------------------------------
' PURPOSE
'   Stable public facade of the DatePicker: the supported enums and types, the
'   supported API and the procedures Excel calls back by name
'
' WHY THIS EXISTS
'   Hosts and Excel bind to these names. Keeping them in one thin, non-private
'   module lets the implementation move behind Option Private Module without
'   changing a supported name, signature, default or callback
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Every procedure delegates at once to the matching <Name>_Impl procedure,
'   qualified with the name of the private module that holds it: M_DP_Core,
'   M_DP_Timer for M_Timer_Tick, M_DP_GridIcon for M_GridIcon_PurgeAll
'
' ERROR POLICY
'   Wrappers have no error handler: errors raised by the implementation reach
'   the caller with their original number, source and description
'
' DEPENDENCIES
'   M_DP_Core
'   M_DP_Timer
'   M_DP_GridIcon
'
' NOTES
'   src/public-surface.json classifies every public symbol; tools/
'   check_public_surface.py keeps this module, the manifest and the README in
'   agreement
'
'   Dependency rule: calls flow from this facade to the private modules, never
'   back. Internal code calls the <Name>_Impl procedures; only the callback
'   names registered with Excel (Ribbon, OnAction, OnKey, OnTime) point here
'
'   This module must contain no WinAPI declaration, no module-level variable
'   and no implementation logic
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

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
' PUBLIC TYPES
'------------------------------------------------------------------------------
    'Structured outcome of one borderless-window styling operation
    Public Type DP_WindowStyleResult
        Attempted               As Boolean      'True when the native style was actually touched
        Applied                 As Boolean      'True when the borderless style is fully in effect
        Committed               As Boolean      'True once the style write succeeded
        RolledBack              As Boolean      'True when the original style was restored
        RecoveryRequired        As Boolean      'True when the window is in no known good state
        FailedStep              As String       'Step that failed, empty on success
        LastApiError            As Long         'WinAPI last-error behind the failure
    End Type

    'Structured outcome of one DatePicker write-back operation
    Public Type DP_WriteResult
        AttemptedCount          As Double       'Cells the write-back targeted
        WrittenCount            As Double       'Cells that received the value
        LockedSkippedCount      As Double       'Protected locked cells skipped
        LockedSkippedAddresses  As String       'Addresses of the skipped locked cells
        FormulaSkippedCount     As Double       'Formula cells preserved by policy
        FormulaSkippedAddresses As String       'Addresses of the preserved formula cells
        FailedCount             As Double       'Other suppressed write failures
        FailedAddresses         As String       'Addresses of the failed cells
        ResolvedTargetAddress   As String       'Address of the resolved target
        ExpandedToTableColumn   As Boolean      'True when a table column was resolved
        TableName               As String       'Owning table when expanded
        ColumnName              As String       'Resolved column when expanded
        AreasCount              As Long         'Discontiguous areas in the target
        EventsDisabledByCaller  As Boolean      'True when events were already off
        TechnicalFailureOccurred    As Boolean  'True when an unexpected error stopped the operation
        TechnicalFailureStep        As String   'Handler step the unexpected error occurred in
        TechnicalFailureNumber      As Long     'Original error number, preserved
        TechnicalFailureDescription As String   'Original error description, preserved
    End Type


'
'------------------------------------------------------------------------------
'
'                                   SETTINGS
'
'------------------------------------------------------------------------------
'

Public Sub M_Settings_SetNamespace(ByVal NamespaceName As String)

'
'------------------------------------------------------------------------------
'                        SET SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Isolates this deployment's persisted settings from every other deployment
'   running under the same Windows user
'
' WHY THIS EXISTS
'   The default persistence scope is the Windows user, not the workbook. Some
'   settings are plausible user preferences; others describe one deployment's
'   integration, and HolidayCallback is a callback name this component later
'   executes through Application.Run. A callback chosen for one workbook should
'   not be inherited by another merely because both persist to the same place
'
' INPUTS
'   NamespaceName
'     Caller-supplied persistence scope. An empty string restores the legacy
'     default
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Validates and stores the namespace, which every later settings read and write
'   resolves through
'
' ERROR POLICY
'   Raises when called after settings have been loaded, and when the supplied
'   namespace cannot be used
'
' DEPENDENCIES
'   M_DP_Core.M_Settings_SetNamespace_Impl, which holds the implementation
'
' NOTES
'   Call this before anything that loads settings, which includes ordinary
'   DatePicker startup. Changing the namespace after a load would leave values
'   read from one namespace being written into another, so it is refused rather
'   than silently reinterpreted
'
'   The namespace is chosen by the caller. It is never derived from
'   ThisWorkbook.Name, FullName or any path: those change when a file is renamed,
'   moved or copied, and settings would appear to vanish. Nor is the release
'   version appended, which would make every upgrade look like a settings reset
'
'   This is a persistence scope, not a runtime owner token. The one-provider
'   lease in #37 is runtime state that must not reach the registry, because a
'   persisted lease would survive an Excel restart and block startup forever
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetNamespace_Impl NamespaceName

End Sub

Public Function M_Settings_GetNamespace() As String

'
'------------------------------------------------------------------------------
'                        GET SETTINGS NAMESPACE
'------------------------------------------------------------------------------
' PURPOSE
'   Reports the configured settings namespace
'
' WHY THIS EXISTS
'   A caller that configured a namespace, and a diagnostic reading where settings
'   came from, both need to see it
'
' INPUTS
'   None
'
' RETURNS
'   The configured namespace, or an empty string when the legacy default is in
'   use
'
' BEHAVIOR
'   Returns the stored value without resolving it to an application name
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   M_DP_Core.M_Settings_GetNamespace_Impl, which holds the implementation
'
' NOTES
'   An empty result means the deployment persists exactly where earlier releases
'   did
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetNamespace = M_DP_Core.M_Settings_GetNamespace_Impl()

End Function

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
'   M_DP_Core.M_Settings_SetFirstDayOfWeek_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetFirstDayOfWeek_Impl FirstDayOfWeek

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
'   M_DP_Core.M_Settings_SetFirstDayOfWeekText_Impl, which holds the implementation
'
' NOTES
'   Accepted values depend on M_Settings_TryParseFirstDayOfWeek and should include
'   vbSunday, Sunday, Sun, 1, vbMonday, Monday, Mon, and 2
'
'   This routine intentionally delegates persistence and form refresh so there is
'   only one canonical first-day update path
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetFirstDayOfWeekText_Impl FirstDayOfWeekText

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
'   M_DP_Core.M_Settings_SetUseLocalNames_Impl, which holds the implementation
'
' NOTES
'   The argument name is preserved for compatibility, but this setting controls
'   both day and month captions
'
'   The unchanged-value exit avoids unnecessary registry writes and unnecessary
'   UserForm refresh work
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetUseLocalNames_Impl UseLocalDayNames

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
'   The keyboard shortcut is one of three interactive entry points, alongside the
'   in-grid icon and the right-click menu. Disabling the other two no longer
'   enables it: Application.OnKey is a session-wide binding, and the component
'   does not take it on behalf of a user who did not ask for it
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
'   Applies the requested keyboard shortcut setting unchanged, whatever the
'   right-click menu and in-grid icon settings are
'   Persists the updated settings when needed
'   Synchronizes Application.OnKey with the final setting when this copy owns
'   the provider lease
'   Otherwise leaves Application.OnKey untouched: reports the refusal when
'   another copy holds the lease, or writes one Immediate Window line when no
'   copy holds it yet
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
'   M_DP_Core.M_Settings_SetEnableKeyboardShortcut_Impl, which holds the implementation
'
' NOTES
'   Disabling right-click menu access, in-grid icon access, and keyboard shortcut
'   access at the same time is a valid configuration. A request to disable the
'   keyboard shortcut is honoured even when both other access paths are
'   disabled; the picker remains reachable through DP_Show, DP_Click and the
'   Ribbon (#42)
'
'   The keyboard shortcut is synchronized even when the setting is unchanged
'   because Application.OnKey is session state and may have been reset by Excel,
'   another add-in, workbook activation, or teardown logic
'
'   Application.OnKey is shared by every copy in the Excel session. This setter
'   used to rebind it without consulting the provider lease, so a copy that did
'   not own the runtime could take Ctrl + Shift + D from the owner or release
'   it under the owner (#194). The value is still validated, stored and saved
'   in every copy, because it is a user preference; only the owning copy
'   applies it, here or when DP_Start reads it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetEnableKeyboardShortcut_Impl EnableKeyboardShortcut

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
'   M_DP_Core.M_Settings_SetUseWinAPI_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetUseWinAPI_Impl UseWinAPI

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
'   M_DP_Core.M_Settings_SetAllowOutsideMonthSelection_Impl, which holds the implementation
'
' NOTES
'   The form should call M_DatePolicy_CanSelectDate before accepting a day click
'
'   The unchanged-value exit avoids unnecessary registry writes and unnecessary
'   UserForm refresh work
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetAllowOutsideMonthSelection_Impl AllowOutsideMonthSelection

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
'   M_DP_Core.M_Settings_SetHighlightWeekends_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetHighlightWeekends_Impl HighlightWeekends

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
'   M_DP_Core.M_Settings_SetCloseAfterSelection_Impl, which holds the implementation
'
' NOTES
'   This setting controls form lifecycle after successful write-back only
'
'   It does not require a visual refresh of the loaded DatePicker form because
'   the value is consumed at the next successful selection event
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetCloseAfterSelection_Impl CloseAfterSelection

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
'   M_DP_Core.M_Settings_SetClockMode_Impl, which holds the implementation
'
' NOTES
'   This routine intentionally applies the timer mode even when the requested
'   value is already active, because the live-clock timer may need to be
'   resynchronized after form lifecycle changes
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetClockMode_Impl ClockMode

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
'   M_DP_Core.M_Settings_SetSizeMode_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetSizeMode_Impl SizeMode

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
'   M_DP_Core.M_Settings_SetHolidayCallback_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetHolidayCallback_Impl HolidayCallbackName

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
'   setting, persists the updated settings when needed, and synchronizes the
'   Excel right-click menu when this copy owns the provider lease. Otherwise the
'   menus are left untouched: the refusal is reported when another copy holds
'   the lease, or one Immediate Window line is written when no copy holds it
'   yet. The keyboard shortcut setting is carried forward unchanged: disabling
'   both visual entry points does not re-enable it
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
'   M_DP_Core.M_Settings_SetShowRightClick_Impl, which holds the implementation
'
' NOTES
'   Disabling both right-click menu access and in-grid icon access leaves the
'   keyboard shortcut as the user set it. Zero built-in entry points is a valid
'   configuration: the picker remains reachable through DP_Show, DP_Click and
'   the Ribbon (#42)
'
'   The keyboard change flag and the M_KeyboardShortcut_Update call below are
'   left over from the removed fallback. The resolved keyboard value always
'   equals the current one, so that branch is not taken
'
'   The right-click menu is synchronized even when the persisted setting is
'   unchanged, because Excel context menus may be reset by Excel, another add-in,
'   workbook activation, or application-level cleanup
'
'   The right-click entry sits on command bars shared by every copy in the
'   Excel session. This setter used to add or remove it without consulting the
'   provider lease, so a copy that did not own the runtime could remove the
'   owner's entry (#194). The value is still validated, stored and saved in
'   every copy; only the owning copy applies it, here or when DP_Start reads it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetShowRightClick_Impl ShowRightClick

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
'   setting, persists the updated settings when needed, and removes any stale
'   in-grid icon when the feature is disabled and this copy owns the provider
'   lease. Otherwise no icon is touched: the refusal is reported when another
'   copy holds the lease, or one Immediate Window line is written when no copy
'   holds it yet. The keyboard shortcut setting is carried forward unchanged:
'   disabling both visual entry points does not re-enable it
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
'   M_DP_Core.M_Settings_SetShowGridIcon_Impl, which holds the implementation
'
' NOTES
'   Disabling both right-click menu access and in-grid icon access leaves the
'   keyboard shortcut as the user set it. Zero built-in entry points is a valid
'   configuration: the picker remains reachable through DP_Show, DP_Click and
'   the Ribbon (#42)
'
'   The keyboard change flag and the M_KeyboardShortcut_Update call below are
'   left over from the removed fallback. The resolved keyboard value always
'   equals the current one, so that branch is not taken
'
'   The stale grid icon is removed even when the setting was already disabled,
'   because worksheet shapes may survive prior failures, workbook activation, or
'   partial UI cleanup
'
'   Enabling the in-grid icon is persisted immediately. Actual icon display is
'   expected to occur through the manager selection / context refresh path
'
'   The grid icons belong to the copy that owns the runtime. This setter used
'   to remove them without consulting the provider lease, so a copy that did
'   not own the runtime could delete the owner's icon (#194). The value is
'   still validated, stored and saved in every copy; only the owning copy
'   applies it, here or when DP_Start reads it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_SetShowGridIcon_Impl ShowGridIcon

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
'   M_DP_Core.M_Settings_GetFirstDayOfWeek_Impl, which holds the implementation
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetFirstDayOfWeek = M_DP_Core.M_Settings_GetFirstDayOfWeek_Impl()

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
'   M_DP_Core.M_Settings_GetUseLocalNames_Impl, which holds the implementation
'
' NOTES
'   The historical procedure name is preserved for compatibility
'
'   Although the name refers to day names, the setting also controls month
'   captions where applicable
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetUseLocalNames = M_DP_Core.M_Settings_GetUseLocalNames_Impl()

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
'   M_DP_Core.M_Settings_GetAllowOutsideMonthSelection_Impl, which holds the implementation
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetAllowOutsideMonthSelection = M_DP_Core.M_Settings_GetAllowOutsideMonthSelection_Impl()

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
'   M_DP_Core.M_Settings_GetFirstDayOfWeekText_Impl, which holds the implementation
'
' NOTES
'   This routine does not read the registry directly
'
'   Registry loading, migration, normalization, and defaulting are handled by
'   M_Settings_Load through M_Settings_EnsureLoaded
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetFirstDayOfWeekText = M_DP_Core.M_Settings_GetFirstDayOfWeekText_Impl()

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
'   M_DP_Core.M_Settings_GetShowRightClick_Impl, which holds the implementation
'
' NOTES
'   This routine does not read the registry directly
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetShowRightClick = M_DP_Core.M_Settings_GetShowRightClick_Impl()

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
'   M_DP_Core.M_Settings_GetShowGridIcon_Impl, which holds the implementation
'
' NOTES
'   This routine does not create, move, or remove the worksheet icon
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetShowGridIcon = M_DP_Core.M_Settings_GetShowGridIcon_Impl()

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
'   M_DP_Core.M_Settings_GetEnableKeyboardShortcut_Impl, which holds the implementation
'
' NOTES
'   This routine reports the setting only
'
'   It does not register or remove the Application.OnKey binding
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetEnableKeyboardShortcut = M_DP_Core.M_Settings_GetEnableKeyboardShortcut_Impl()

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
'   M_DP_Core.M_Settings_GetUseWinAPI_Impl, which holds the implementation
'
' NOTES
'   This routine returns the normalized setting state
'
'   Use M_Platform_ShouldUseWinAPI when the caller needs the effective runtime
'   policy combining platform capability and setting state
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetUseWinAPI = M_DP_Core.M_Settings_GetUseWinAPI_Impl()

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
'   M_DP_Core.M_Settings_GetClockMode_Impl, which holds the implementation
'
' NOTES
'   This routine does not start, stop, or resynchronize the live-clock timer
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetClockMode = M_DP_Core.M_Settings_GetClockMode_Impl()

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
'   M_DP_Core.M_Settings_GetSizeMode_Impl, which holds the implementation
'
' NOTES
'   This routine reports the setting only
'
'   It does not resize, rebuild, or repaint the loaded DatePicker form
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetSizeMode = M_DP_Core.M_Settings_GetSizeMode_Impl()

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
'   M_DP_Core.M_Settings_GetHighlightWeekends_Impl, which holds the implementation
'
' NOTES
'   This routine reports the setting only
'
'   It does not repaint or rebuild the loaded DatePicker form
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetHighlightWeekends = M_DP_Core.M_Settings_GetHighlightWeekends_Impl()

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
'   M_DP_Core.M_Settings_GetCloseAfterSelection_Impl, which holds the implementation
'
' NOTES
'   This routine reports the setting only
'
'   The setting is consumed during the successful write-back lifecycle
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetCloseAfterSelection = M_DP_Core.M_Settings_GetCloseAfterSelection_Impl()

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
'   M_DP_Core.M_Settings_GetHolidayCallback_Impl, which holds the implementation
'
' NOTES
'   This routine reports the configured callback name only
'
'   It does not execute, validate, or inspect the callback procedure
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_Settings_GetHolidayCallback = M_DP_Core.M_Settings_GetHolidayCallback_Impl()

End Function


'
'------------------------------------------------------------------------------
'
'                            LIFECYCLE AND COMMANDS
'
'------------------------------------------------------------------------------
'

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
'   M_DP_Core.DP_Preload_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Preload_Impl

End Sub

Public Sub DP_ForceReleaseProviderLease()

'
'------------------------------------------------------------------------------
'                     FORCE RELEASE PROVIDER LEASE
'------------------------------------------------------------------------------
' PURPOSE
'   Removes the one-provider lease regardless of who owns it, so an operator can
'   recover a session the automatic policy has locked out
'
' WHY THIS EXISTS
'   The lease lives for the Excel process; the token proving ownership lives in
'   VBA module state. A VBA project reset destroys the token and leaves the lease,
'   after which the provider that created it can no longer prove ownership and
'   every entry point refuses
'
'   That is correct for a crashed provider and unusable during development, where
'   re-importing a module resets the project several times an hour. Restarting
'   Excel is a poor answer to a self-inflicted lockout
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Deletes the lease and clears this project''s token, whether or not this
'   project owned it
'
' ERROR POLICY
'   Does not raise. A lease that is absent or already gone is not an error
'
' DEPENDENCIES
'   M_DP_Core.DP_ForceReleaseProviderLease_Impl, which holds the implementation
'
' NOTES
'   This is the one place the component deliberately breaks its own rule that a
'   lease is never deleted unless ownership can be proved. That is safe only
'   because it is an explicit operator action, never called automatically and
'   never called by DP_Start, DP_Stop or DP_RepairRuntime
'
'   Call it only when no other copy of the DatePicker is running. Calling it
'   while a live provider owns the lease will let a second provider acquire one,
'   which is exactly the unsupported configuration #37 exists to prevent
'
'   Restarting Excel achieves the same thing without the risk, because the lease
'   is Temporary. Prefer that when in doubt
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_ForceReleaseProviderLease_Impl

End Sub

Public Sub DP_Start()

'
'------------------------------------------------------------------------------
'                               START DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts and synchronizes the DatePicker runtime as a transaction
'
' WHY THIS EXISTS
'   Startup registers application-wide state: a provider lease, Application
'   event hooks, a context menu, a keyboard binding and a worksheet shape. A
'   failure partway through used to leave some of that registered with nothing
'   owning it, and a lease acquired by the failed attempt was surrendered anyway
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the caller's event state, classifies whether the lease was already
'   owned, then admits, creates the manager, synchronizes the menu, the keyboard
'   shortcut and the hidden grid icon, and refreshes the selection context.
'   Restores the caller's event state on every path
'
' ERROR POLICY
'   Preserves the primary startup failure and re-raises it with the step that
'   produced it. A fresh-start failure rolls back every DatePicker-owned resource
'   and releases the lease only when critical cleanup proves complete. A repeated
'   start never releases a lease that pre-existed the call
'
' DEPENDENCIES
'   M_DP_Core.DP_Start_Impl, which holds the implementation
'
' NOTES
'   The pre-owned and freshly-acquired cases are deliberately asymmetric. A
'   failed fresh start rolls back, because nothing was working before it. A
'   failed repeated start preserves the existing runtime and records
'   PreOwnedRuntime=PRESERVED, because rolling back would dismantle a runtime
'   that was already serving the user
'
'   Refused admission exits through CleanExit rather than the error handler. A
'   refusal is not a failure of this call; it is another provider owning the
'   session, and the refusal is reported by the admission boundary
'
'   Every step is followed by a fault seam, so the regression matrix can fail
'   startup at each boundary and assert what rollback did
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Start_Impl

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
'   The form is explicitly loaded before it is shown, so its runtime UI is built
'   while it is still hidden. Positioning near the mouse is applied once the
'   native window exists: UserForm_Activate positions it on first activation,
'   and DP_Show applies one final correction after Show
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Admits this provider through the shared admission boundary and exits when
'   admission is refused. Ensures DatePicker infrastructure is available, and
'   reuses an already-visible picker and exits when possible. Otherwise resolves
'   the initial display date from ActiveCell when possible, stores bridge state,
'   refreshes an already-loaded form in place or loads UF_DatePicker when none
'   is loaded, shows the form modelessly, applies the configured clock mode, and
'   applies one final post-show positioning correction
'
' ERROR POLICY
'   Raises a descriptive runtime error if the DatePicker cannot be prepared,
'   loaded, resolved, or shown
'
'   Mouse positioning is best-effort and does not prevent the picker from opening
'
' DEPENDENCIES
'   M_DP_Core.DP_Show_Impl, which holds the implementation
'
' NOTES
'   Load UF_DatePicker is intentional
'
'   An existing instance is not unloaded. A loaded form is reused or refreshed
'   in place, so M_FormBridge_UnloadLoadedPicker is not called
'
'   The StepName diagnostic is intentionally retained because UserForm loading
'   can fail inside nested Initialize routines and otherwise surface only as a
'   generic DP_Show failure
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Show_Impl

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
'   M_DP_Core.DP_Click_Impl, which holds the implementation
'
' NOTES
'   This routine intentionally remains Public so Excel UI surfaces can call it
'
'   Do not add business logic here. Keep DatePicker launch behavior centralized
'   in DP_OpenForActiveCell
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Click_Impl

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
'   M_DP_Core.DP_OpenForActiveCell_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_OpenForActiveCell_Impl

End Sub

Public Sub DP_RepairRuntime()

'
'------------------------------------------------------------------------------
'                          REPAIR DATEPICKER RUNTIME
'------------------------------------------------------------------------------
' PURPOSE
'   Rebuilds the interactive runtime, but only after the previous owned runtime
'   has been proven clean
'
' WHY THIS EXISTS
'   Repair addresses the most common real-world failure: Application.EnableEvents
'   left False by an unrelated macro, which silently kills all event routing.
'   Rebuilding on top of state that could not be cleaned would produce a second
'   set of registrations rather than a repair
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the cleanup transaction without releasing the lease, and rebuilds only
'   when critical cleanup succeeded
'
' ERROR POLICY
'   Never releases the provider lease. Incomplete cleanup blocks the rebuild. A
'   rebuild failure keeps the primary error and attempts another lease-retaining
'   cleanup
'
' DEPENDENCIES
'   M_DP_Core.DP_RepairRuntime_Impl, which holds the implementation
'
' NOTES
'   Repair is the one documented exception to preserving the caller's event
'   state: it leaves Application.EnableEvents True on purpose, because a caller
'   that left them False is the condition being repaired
'
'   The lease is never released here. Repair means the same provider continues to
'   own the session, so surrendering ownership would turn a repair into an
'   unannounced shutdown
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_RepairRuntime_Impl

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
'   M_DP_Core.DP_Close_Impl, which holds the implementation
'
' NOTES
'   This routine intentionally avoids direct default-instance references until a
'   loaded form has been explicitly resolved
'
'   Selected-date state is not cleared here because it is owned by the selection
'   / write-back flow and is refreshed when the picker is opened again
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Close_Impl

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
'   M_DP_Core.DP_Hide_Impl, which holds the implementation
'
' NOTES
'   Use DP_Hide when you want fast reuse
'
'   Use DP_Close when you want to unload and release the form
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Hide_Impl

End Sub

Public Sub DP_Stop()

'
'------------------------------------------------------------------------------
'                               STOP DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Tears down the DatePicker runtime as a transaction and releases the provider
'   lease only when critical cleanup is proven clean
'
' WHY THIS EXISTS
'   Teardown used to suppress its own errors and release the lease regardless, so
'   a shutdown that failed halfway reported success while application-wide state
'   survived. A new provider could then start on top of it
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Captures the caller's event state and runs the cleanup transaction with lease
'   release enabled when this project owns the lease. When it does not, it
'   touches nothing, records an unsuccessful stop and writes one Immediate
'   Window line. It never shows a dialog
'
' ERROR POLICY
'   Never raises outward. The outcome is reported through the observation fields
'
' DEPENDENCIES
'   M_DP_Core.DP_Stop_Impl, which holds the implementation
'
' NOTES
'   Ownership is checked before anything is touched. A refused copy must not be
'   able to dismantle the owner's registrations, which is the defect #37 closed
'   for the open paths and this closes for teardown
'
'   Retaining ownership when cleanup is incomplete is the point of the issue. A
'   lease released over surviving shared state is worse than a lease held too
'   long, because the next provider starts on top of it
'
'   DP_Stop usually runs from Workbook_BeforeClose, at shutdown. A copy that
'   never started (for example an add-in whose Workbook_Open did not run) used
'   to show the "another copy is already active" message on every exit, although
'   no other copy existed (#115). After a startup-folder launch the lease could
'   also be lost or replaced while Excel ran, so the message still appeared at
'   exit (#113). A stop runs at shutdown, where the operator can do nothing
'   with it, so a non-owner stop is silent in every case. Genuine provider
'   conflicts are still reported by the start and open paths
'
'   A copy that holds its token but whose lease bar vanished first reclaims the
'   free lease and then stops as the owner. A lease held by another provider or
'   unreadable is left untouched
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Stop_Impl

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
'   M_DP_Core.DP_Today_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Today_Impl

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
'   M_DP_Core.DP_Now_Impl, which holds the implementation
'
' NOTES
'   This command writes a date-time value
'
'   Selected-date highlighting uses only the date portion of the timestamp
'
'   This is an interactive entry point, so it reports a partial write once, after
'   the whole operation. The write engine collects the facts and displays nothing
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.DP_Now_Impl

End Sub

Public Function DP_FillTableColumn( _
    ByVal ValueToWrite As Date, _
    Optional ByVal ConfirmFill As Boolean = True, _
    Optional ByVal OverwriteFormulas As Boolean = False) As DP_WriteResult
'
'------------------------------------------------------------------------------
'                           FILL TABLE COLUMN
'------------------------------------------------------------------------------
' PURPOSE
'   Writes one date to every cell of the Excel Table data column containing the
'   current selection
'
' WHY THIS EXISTS
'   Filling a table column is a legitimate operation, but it used to happen
'   implicitly: selecting one cell inside a table and picking a date wrote the
'   whole column, with nothing to indicate the scope
'
'   That default is now single-cell. This routine is the deliberate way to ask
'   for the broad write, so the scope comes from the command the user invoked
'   rather than from hidden state
'
' INPUTS
'   ValueToWrite
'     Date to write to every cell of the resolved table column
'
'   ConfirmFill
'     True to describe the resolved scope and require confirmation, and to report
'     a non-table selection with a message
'
'     False for a non-interactive call. Suppresses both prompts, which is what
'     makes the routine usable from the regression harness
'
'   OverwriteFormulas
'     False, the default, preserves a cell holding a formula and counts it in
'     FormulaSkippedCount. True replaces formulas in the filled column
'
' RETURNS
'   DP_WriteResult describing the fill
'
'   A zeroed result when the selection was not a table data cell or the user
'   declined the described scope. WrittenCount is then zero and nothing was
'   written
'
'   A result with AttemptedCount above zero and WrittenCount zero when the fill
'   reached the column but wrote no cell
'
' BEHAVIOR
'   Resolves the table column owning the selection. Reports and exits when the
'   selection is not a table data cell. Otherwise confirms the scope, writes
'   through the normal write-back engine with table expansion explicitly enabled,
'   checks the predicted scope against AttemptedCount, and reports a partial write
'   once for the whole operation
'
' ERROR POLICY
'   Raises a descriptive runtime error for a zero date and for genuine write
'   failures
'
'   A selection outside a table data body is an ordinary usage condition, not an
'   error. It reports and exits cleanly
'
'   A partial write is reported, not raised. That includes a fill that wrote no
'   cell at all: since #21 the engine returns a zero-write as a complete result
'   rather than raising
'
'   A predicted scope that does not match AttemptedCount is a detectable
'   inconsistency rather than a partial write, and is reported as one
'
' DEPENDENCIES
'   M_DP_Core.DP_FillTableColumn_Impl, which holds the implementation
'
' NOTES
'   The write itself is delegated to M_WriteBack_Apply with NoTableGrow:=False
'   rather than writing the resolved range directly. That reuses the existing
'   event suppression, value validation and rollback rather than duplicating
'   them, and keeps one write engine
'
'   The resolved column is used for the confirmation text. The engine resolves it
'   again when it writes, which is cheap and avoids two routines disagreeing
'   about the target
'
'   The predicted cell count is resolved on both paths, not only when prompting,
'   so the non-interactive path can make the same comparison. That is what turns
'   the confirmation prompt into a check
'
'   The prediction is checked against AttemptedCount, not WrittenCount. A fill
'   that legitimately skips protected cells writes fewer cells than it predicted
'   and is still a correct prediction:
'
'     Predicted scope / AttemptedCount   247
'     WrittenCount                       244
'     LockedSkippedCount                   3
'
'   A prediction that does not match AttemptedCount means the target changed
'   between preview and application, or that the two resolution paths diverged
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        DP_FillTableColumn = M_DP_Core.DP_FillTableColumn_Impl( _
            ValueToWrite, _
            ConfirmFill, _
            OverwriteFormulas)

End Function


'
'------------------------------------------------------------------------------
'
'                                    TIMER
'
'------------------------------------------------------------------------------
'

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
'   M_DP_Timer.M_Timer_Tick_Impl, which holds the implementation
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
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Timer.M_Timer_Tick_Impl

End Sub


'
'------------------------------------------------------------------------------
'
'                                  GRID ICON
'
'------------------------------------------------------------------------------
'

Public Sub M_GridIcon_PurgeAll()

'
'------------------------------------------------------------------------------
'                         PURGE ALL GRID ICONS
'------------------------------------------------------------------------------
' PURPOSE
'   Removes DatePicker-owned canonical and pending grid-icon Shapes from every
'   open workbook while leaving foreign same-named Shapes untouched
'
' WHY THIS EXISTS
'   The documented host template calls it before save and print, so no
'   DatePicker grid icon is stored in or printed from a workbook
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Runs the transactional grid-icon purge and, when it is incomplete, prints
'   one diagnostic line to the Immediate window
'
' ERROR POLICY
'   Best-effort public wrapper. Transactional lifecycle callers use the internal
'   Boolean result so protected or otherwise undeletable owned Shapes remain an
'   explicit incomplete-cleanup condition
'
' DEPENDENCIES
'   M_DP_GridIcon.M_GridIcon_PurgeAll_Impl, which holds the implementation
'
' NOTES
'   None
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_GridIcon.M_GridIcon_PurgeAll_Impl

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
'   M_DP_Core.M_Settings_UseLocalNames_Impl, which holds the implementation
'
' NOTES
'   Keep this routine as a thin compatibility wrapper
'   Do not add duplicate settings logic here
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.M_Settings_UseLocalNames_Impl UseLocalDayNames

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
'   Catches runtime errors and reports them through Ribbon_ReportError, which
'   shows a user-facing message box, so Ribbon callback failures do not fail
'   silently
'
' DEPENDENCIES
'   M_DP_Core.Ribbon_ShowPicker_Impl, which holds the implementation
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.Ribbon_ShowPicker_Impl control

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
'   Starts the runtime with DP_Start when no provider holds the lease, so Reset
'   also recovers a copy that never started. Otherwise delegates to
'   DP_RepairRuntime. Displays a confirmation only when the recorded lifecycle
'   outcome of that exact operation is a success
'
' ERROR POLICY
'   Catches runtime errors and reports them through Ribbon_ReportError
'
'   A refused start or repair does not raise. The refusal is reported at the
'   admission boundary and this callback adds nothing after it
'
' DEPENDENCIES
'   M_DP_Core.Ribbon_Reset_Impl, which holds the implementation
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
'   Returning without an error is not proof of success. DP_RepairRuntime
'   returns normally when it refuses, and this callback used to follow the
'   refusal message with "repair completed successfully" (#89)
'
'   Starting instead of repairing applies only to a free lease, which is the
'   same condition under which DP_Start itself admits a provider. It never
'   takes over a lease another provider holds. This gives a copy whose
'   runtime never started a one-click recovery
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.Ribbon_Reset_Impl control

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
'   M_DP_Core.Ribbon_Demo_Impl, which holds the implementation
'
' NOTES
'   The Control argument is required by the RibbonX callback signature even when
'   this routine does not use it directly
'
'   The host workbook is resolved through DP_DemoSheet_ResolveHostWorkbook
'   rather than assumed to be ThisWorkbook. ThisWorkbook is the add-in when the
'   component is loaded as an .xlam, and an add-in has no worksheets, so the
'   previous assumption made this button fail on every click in that deployment
'
'   The demo sheet is built on first use, so the button works from a session
'   that has never opened the demo workbook
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE TO IMPLEMENTATION
'------------------------------------------------------------------------------
    'Run the private implementation; its errors reach the caller unchanged
        M_DP_Core.Ribbon_Demo_Impl control

End Sub
