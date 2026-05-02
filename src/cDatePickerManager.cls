VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatePickerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'==============================================================================
' CLASS: cDatePickerManager
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the central Excel Application event manager for the DatePicker
'   project
'
' WHY THIS EXISTS
'   The DatePicker spans several Excel / VBA UI surfaces:
'     - modeless runtime UserForm
'     - in-grid DatePicker icon
'     - right-click menu integration
'     - workbook / worksheet / window event transitions
'
'   Centralizing the Application event flow in one manager class keeps the
'   project predictable, avoids duplicate event wiring, and prevents stale UI
'   artifacts from surviving workbook, worksheet, save, print, resize, close, or
'   activation transitions
'
' INPUTS
'   None at class level
'
' RETURNS
'   Nothing at class level
'
' BEHAVIOR
'   Manages:
'     - Application event hook lifecycle
'     - guarded selection-change refreshes
'     - guarded worksheet-change refreshes
'     - workbook / sheet / window context refreshes
'     - light grid-icon removal during high-frequency refresh paths
'     - hard grid-icon purge during save / close / print / teardown paths
'     - loaded DatePicker form cleanup
'     - open-form refresh when close-after-selection is disabled
'     - safe reset of transient DatePicker UI
'
' ERROR POLICY
'   Event callbacks are best-effort and must not raise errors into Excel
'
'   Public predicates return safe defaults on error
'
'   Refresh paths use a re-entrancy guard to avoid recursive UI churn
'
' DEPENDENCIES
'   Excel.Application
'   VBA.UserForms
'   DP_Close
'   M_GridIcon_Remove
'   M_GridIcon_Create
'   M_GridIcon_PurgeAll
'   M_Settings_Load
'   M_FormBridge_RefreshFromActiveCell
'   gDP_ShowGridIcon
'   gDP_CloseAfterSelection
'
' NOTES
'   Runtime UserForm control / label event wiring does not belong here. It is
'   handled by cDatePickerLabelHook and UF_DatePicker
'
'   The manager distinguishes between:
'     - the picker being loaded in memory
'     - the picker being visible on screen
'
'   When gDP_CloseAfterSelection is False and the picker is visible, selection
'   changes keep the form open and refresh its selected-date state from the new
'   ActiveCell instead of closing it
'
'   Date-format detection is intentionally heuristic and is used only for UI
'   gating, not for full semantic format parsing
'
'   M_GridIcon_Remove should be the lightweight active-context cleanup routine
'
'   M_GridIcon_PurgeAll should be the heavier all-workbook cleanup routine used
'   only for hard lifecycle boundaries
'
' UPDATED
'   2026-04-30
'==============================================================================

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_FORM_NAME                    As String = "UF_DatePicker"    'DatePicker UserForm name

'------------------------------------------------------------------------------
' EVENTS
'------------------------------------------------------------------------------
    Private WithEvents mExcelApp                  As Excel.Application            'Hooked Excel Application instance
Attribute mExcelApp.VB_VarHelpID = -1

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mIsBusy                         As Boolean          'Re-entrancy guard for refresh / teardown
    Private mIsHooked                       As Boolean          'True when Application events are hooked
    Private mEmbeddedGridIconPath           As String           'Cached temp-file path for embedded grid icon

'
'==============================================================================
'
'                               CLASS LIFECYCLE
'
'==============================================================================
'

Private Sub Class_Initialize()

'
'------------------------------------------------------------------------------
'                           INITIALIZE MANAGER
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the DatePicker manager and hooks Excel Application events
'
' WHY THIS EXISTS
'   The manager must receive Excel Application callbacks in order to coordinate
'   DatePicker UI cleanup, in-grid icon lifecycle, and open-form refresh behavior
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Initializes internal state flags, loads DatePicker settings on a best-effort
'   basis, hooks the Excel Application object, and records whether hook-up
'   succeeded
'
' ERROR POLICY
'   Best-effort initialization. Startup should not hard-fail if Excel is in a
'   transient state during workbook load or VBA reset
'
' DEPENDENCIES
'   M_Settings_Load
'   Excel.Application
'
' NOTES
'   Settings are loaded here so event gating can safely use persisted feature
'   flags such as gDP_ShowGridIcon and gDP_CloseAfterSelection
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE STATE
'------------------------------------------------------------------------------
    'Suppress startup errors
        On Error Resume Next
    'Clear the busy guard
        mIsBusy = False
    'Clear the hook-state flag
        mIsHooked = False

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Load persisted DatePicker settings before events start using them
        M_Settings_Load

'------------------------------------------------------------------------------
' PREPARE EMBEDDED GRID ICON
'------------------------------------------------------------------------------
    'Prepare the embedded grid icon file on a best-effort basis
        mEmbeddedGridIconPath = M_GridIcon_EnsureEmbeddedIconFile

'------------------------------------------------------------------------------
' HOOK APPLICATION EVENTS
'------------------------------------------------------------------------------
    'Hook the Excel Application object
        Set mExcelApp = Excel.Application

    'Store whether the hook was created
        mIsHooked = Not mExcelApp Is Nothing

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

Private Sub Class_Terminate()

'
'------------------------------------------------------------------------------
'                           TERMINATE MANAGER
'------------------------------------------------------------------------------
' PURPOSE
'   Releases the Application event hook and clears transient DatePicker UI
'
' WHY THIS EXISTS
'   A manager instance may be released during workbook close, project reset, or
'   explicit teardown. Cleanup here prevents orphaned forms or worksheet shapes
'   from surviving beyond the manager lifetime
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Marks the manager as busy, performs hard cleanup, releases the Application
'   event hook, and clears internal state flags
'
' ERROR POLICY
'   Best-effort teardown. Termination must not raise outward
'
' DEPENDENCIES
'   Clear_AllDatePickerUI
'
' NOTES
'   Teardown uses the hard purge path because it is a lifecycle boundary, not a
'   high-frequency selection refresh
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Suppress teardown errors
        On Error Resume Next
    'Block nested refresh work during teardown
        mIsBusy = True
    'Clear cached embedded grid-icon path
        mEmbeddedGridIconPath = vbNullString
    'Clear all transient UI artifacts using the hard-purge path
        Clear_AllDatePickerUI True, True
    'Release the Application event hook
        Set mExcelApp = Nothing

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Clear the hook-state flag
        mIsHooked = False

    'Clear the busy guard
        mIsBusy = False

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

'
'==============================================================================
'
'                                  PROPERTIES
'
'==============================================================================

Public Property Get Is_PickerLoaded() As Boolean

'
'------------------------------------------------------------------------------
'                           IS PICKER LOADED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker UserForm is currently loaded in memory
'
' WHY THIS EXISTS
'   Loaded state and visible state are not the same thing. Callers may need to
'   know whether the form instance exists without assuming it is visible
'
' INPUTS
'   None
'
' RETURNS
'   True if UF_DatePicker is loaded; otherwise False
'
' BEHAVIOR
'   Scans the loaded VBA.UserForms collection without touching the default form
'   instance
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Is_UserFormLoaded
'
' NOTES
'   This property deliberately avoids direct references to UF_DatePicker
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return whether the DatePicker form is currently loaded
        Is_PickerLoaded = Is_UserFormLoaded(DP_FORM_NAME)

End Property

Public Property Get PickerVisible() As Boolean

'
'------------------------------------------------------------------------------
'                           PICKER VISIBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker UserForm is currently visible on screen
'
' WHY THIS EXISTS
'   A UserForm can be loaded but hidden. This property derives actual visibility
'   instead of relying on a mirrored mutable flag
'
' INPUTS
'   None
'
' RETURNS
'   True if UF_DatePicker is loaded and visible; otherwise False
'
' BEHAVIOR
'   Scans the loaded VBA.UserForms collection and reads the Visible property of
'   the matching DatePicker form instance
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Is_UserFormVisible
'
' NOTES
'   This property deliberately avoids direct references to UF_DatePicker
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return whether the DatePicker form is visible
        PickerVisible = Is_UserFormVisible(DP_FORM_NAME)

End Property

Public Property Get Is_Hooked() As Boolean

'
'------------------------------------------------------------------------------
'                           IS HOOKED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the manager believes the Application event hook is active
'
' WHY THIS EXISTS
'   This lightweight diagnostic property is useful during startup validation,
'   tests, and troubleshooting
'
' INPUTS
'   None
'
' RETURNS
'   True if the manager believes Application events are hooked; otherwise False
'
' BEHAVIOR
'   Returns the internal hook-state flag
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   mIsHooked
'
' NOTES
'   This is a diagnostic flag, not a guarantee that Excel will fire events
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return the hook-state flag
        Is_Hooked = mIsHooked

End Property

Public Property Get Is_Busy() As Boolean

'
'------------------------------------------------------------------------------
'                           IS BUSY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the manager is currently inside a guarded refresh or teardown
'
' WHY THIS EXISTS
'   This lightweight diagnostic property helps validate re-entrancy behavior and
'   troubleshoot event recursion
'
' INPUTS
'   None
'
' RETURNS
'   True if guarded work is in progress; otherwise False
'
' BEHAVIOR
'   Returns the internal busy-state flag
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   mIsBusy
'
' NOTES
'   The busy flag is intentionally simple and local to this manager instance
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return the busy-state flag
        Is_Busy = mIsBusy

End Property

'
'==============================================================================
'
'                               PUBLIC COMMANDS
'
'==============================================================================

Public Sub Reset_DatePickerUI()

'
'------------------------------------------------------------------------------
'                           RESET DATEPICKER UI
'------------------------------------------------------------------------------
' PURPOSE
'   Resets all transient DatePicker UI artifacts
'
' WHY THIS EXISTS
'   Host code, demo sheets, and regression tests need a simple command that
'   restores the DatePicker UI to a clean baseline state
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Closes the DatePicker form when loaded and purges all in-grid icon artifacts
'
' ERROR POLICY
'   Best-effort cleanup. Does not intentionally raise outward
'
' DEPENDENCIES
'   Clear_AllDatePickerUI
'
' NOTES
'   This is a public convenience wrapper and intentionally uses the hard-purge
'   path because manual reset is not a high-frequency event
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET UI
'------------------------------------------------------------------------------
    'Reset all transient DatePicker UI artifacts using the hard-purge path
        Clear_AllDatePickerUI True, True

End Sub

Public Sub Handle_SelectionChange(Optional ByVal vTarget As Variant)

'
'------------------------------------------------------------------------------
'                           HANDLE SELECTION CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Reacts to an Excel cell / selection change
'
' WHY THIS EXISTS
'   Selection movement is the primary trigger for deciding whether the in-grid
'   DatePicker entry point should be shown, moved, hidden, or refreshed
'
' INPUTS
'   vTarget
'     Optional event Target Range. If omitted, the routine falls back to the
'     current ActiveCell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Enters the manager re-entrancy guard, refreshes DatePicker UI for the new
'   selection context, and releases the guard deterministically
'
' ERROR POLICY
'   Best-effort event handling. Failures trigger cleanup and do not raise into
'   Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Release_Busy
'
' NOTES
'   When gDP_CloseAfterSelection is False and the picker is visible, this routine
'   keeps the form open and refreshes it from the new ActiveCell
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GuardWasEntered        As Boolean        'True when this routine owns the busy guard

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from recursive or failing UI event handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RE-ENTRANCY GUARD
'------------------------------------------------------------------------------
    'Attempt to enter the busy guard
        GuardWasEntered = Try_BeginBusy()

    'Exit if another refresh is already running
        If Not GuardWasEntered Then Exit Sub

'------------------------------------------------------------------------------
' REFRESH UI
'------------------------------------------------------------------------------
    'Refresh DatePicker UI for the new selection context
        If IsMissing(vTarget) Then
            Refresh_CurrentUI
        Else
            Refresh_CurrentUI vTarget
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy

    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False

    'Leave through the shared guard-release path
        Resume SafeExit

End Sub

Public Sub Handle_WorksheetChange(Optional ByVal vTarget As Variant)

'
'------------------------------------------------------------------------------
'                           HANDLE WORKSHEET CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Reacts to changes in worksheet contents
'
' WHY THIS EXISTS
'   A cell edit can invalidate in-grid icon positioning, selection relevance, or
'   the selected-date state of an already open picker
'
' INPUTS
'   vTarget
'     Optional event Target Range. If omitted, the routine falls back to the
'     current ActiveCell
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Uses the same guarded refresh policy used for selection changes
'
' ERROR POLICY
'   Best-effort event handling. Failures trigger cleanup and do not raise into
'   Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Release_Busy
'
' NOTES
'   This routine is intentionally thin because the refresh policy is centralized
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GuardWasEntered        As Boolean        'True when this routine owns the busy guard

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from recursive or failing UI event handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RE-ENTRANCY GUARD
'------------------------------------------------------------------------------
    'Attempt to enter the busy guard
        GuardWasEntered = Try_BeginBusy()

    'Exit if another refresh is already running
        If Not GuardWasEntered Then Exit Sub

'------------------------------------------------------------------------------
' REFRESH UI
'------------------------------------------------------------------------------
    'Refresh DatePicker UI for the changed worksheet context
        If IsMissing(vTarget) Then
            Refresh_CurrentUI
        Else
            Refresh_CurrentUI vTarget
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy

    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False

    'Leave through the shared guard-release path
        Resume SafeExit

End Sub

Public Sub Handle_WorkbookContextChange()

'
'------------------------------------------------------------------------------
'                           HANDLE WORKBOOK CONTEXT CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Reacts to workbook-level or sheet-level context changes
'
' WHY THIS EXISTS
'   Workbook activation, sheet activation, and new workbook creation can change
'   the active context and invalidate the current DatePicker UI decision
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Uses the guarded refresh policy against the current ActiveCell context
'
' ERROR POLICY
'   Best-effort event handling. Failures trigger cleanup and do not raise into
'   Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Release_Busy
'
' NOTES
'   This routine is intentionally thin because the refresh policy is centralized
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim GuardWasEntered        As Boolean        'True when this routine owns the busy guard

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from recursive or failing UI event handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' RE-ENTRANCY GUARD
'------------------------------------------------------------------------------
    'Attempt to enter the busy guard
        GuardWasEntered = Try_BeginBusy()

    'Exit if another refresh is already running
        If Not GuardWasEntered Then Exit Sub

'------------------------------------------------------------------------------
' REFRESH UI
'------------------------------------------------------------------------------
    'Refresh DatePicker UI against the current active-cell context
        Refresh_CurrentUI

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy

    'Exit the procedure
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False

    'Leave through the shared guard-release path
        Resume SafeExit

End Sub

Public Sub Clear_AllDatePickerUI( _
    Optional ByVal ClosePicker As Boolean = True, _
    Optional ByVal PurgeGridIcons As Boolean = False)

'
'------------------------------------------------------------------------------
'                           CLEAR ALL DATEPICKER UI
'------------------------------------------------------------------------------
' PURPOSE
'   Clears transient DatePicker UI artifacts
'
' WHY THIS EXISTS
'   The DatePicker uses multiple UI surfaces. Centralized cleanup prevents stale
'   forms, grid icons, or worksheet shapes from surviving context changes
'
' INPUTS
'   ClosePicker
'     True to close the loaded DatePicker form
'     False to keep the loaded DatePicker form open and clear only grid artifacts
'
'   PurgeGridIcons
'     True to use the hard all-workbook grid-icon purge path
'     False to use the lightweight active-context grid-icon removal path
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Optionally closes the DatePicker UserForm and removes in-grid DatePicker icon
'   artifacts through either the light or hard cleanup path
'
' ERROR POLICY
'   Best-effort cleanup. Missing artifacts are safe no-ops
'
' DEPENDENCIES
'   DP_Close
'   M_GridIcon_Remove
'   M_GridIcon_PurgeAll
'
' NOTES
'   Use PurgeGridIcons = False for high-frequency selection refreshes
'
'   Use PurgeGridIcons = True for hard boundaries such as save, close, print,
'   explicit reset, and manager teardown
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress cleanup errors
        On Error Resume Next

'------------------------------------------------------------------------------
' CLOSE PICKER IF REQUESTED
'------------------------------------------------------------------------------
    'Close the DatePicker form when requested
        If ClosePicker Then DP_Close

'------------------------------------------------------------------------------
' REMOVE OR PURGE GRID ICONS
'------------------------------------------------------------------------------
    'Use the requested grid-icon cleanup policy
        If PurgeGridIcons Then
            M_GridIcon_PurgeAll
        Else
            M_GridIcon_Remove
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Function Should_ShowDP(Optional ByVal vTarget As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                           SHOULD SHOW DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the DatePicker entry point should be shown for a target cell
'
' WHY THIS EXISTS
'   Public callers may need the manager's full gating decision without
'   duplicating target resolution and cell checks
'
' INPUTS
'   vTarget
'     Optional Range-like target. If omitted, ActiveCell is used
'
' RETURNS
'   True if the in-grid DatePicker entry point should be shown; otherwise False
'
' BEHAVIOR
'   Resolves a single target cell and applies cell-level gating
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Resolve_TargetCell
'   Should_ShowDP_ForCell
'
' NOTES
'   This routine does not create UI
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetCell             As Excel.Range        'Resolved single target cell

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not showing the DatePicker entry point
        Should_ShowDP = False

'------------------------------------------------------------------------------
' RESOLVE TARGET CELL
'------------------------------------------------------------------------------
    'Resolve target from ActiveCell or explicit target
        If IsMissing(vTarget) Then
            If Not Resolve_TargetCell(TargetCell) Then Exit Function
        Else
            If Not Resolve_TargetCell(TargetCell, vTarget) Then Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY CELL GATING
'------------------------------------------------------------------------------
    'Return the cell-level gating decision
        Should_ShowDP = Should_ShowDP_ForCell(TargetCell)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Safe default
        Should_ShowDP = False

End Function

'
'==============================================================================
'
'                               REFRESH HELPERS
'
'==============================================================================

Private Sub Refresh_CurrentUI(Optional ByVal vTarget As Variant)

'
'------------------------------------------------------------------------------
'                           REFRESH CURRENT UI
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the manager's current-context DatePicker UI refresh policy
'
' WHY THIS EXISTS
'   Selection, edit, workbook, sheet, and window events all need consistent UI
'   lifecycle handling
'
' INPUTS
'   vTarget
'     Optional event Target Range
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Decides whether a visible picker should remain open, clears stale grid icon
'   artifacts through the lightweight path, resolves the current target cell,
'   refreshes the open picker from ActiveCell when appropriate, and recreates the
'   in-grid icon when the target context allows it
'
' ERROR POLICY
'   Best-effort. Failures trigger cleanup and do not raise outward
'
' DEPENDENCIES
'   Should_KeepPickerOpenOnContextChange
'   Clear_AllDatePickerUI
'   Resolve_TargetCell
'   M_FormBridge_RefreshFromActiveCell
'   Should_ShowDP_ForCell
'   M_GridIcon_Create
'
' NOTES
'   This routine is the key path that supports close-after-selection = False
'   across worksheet cell navigation
'
'   This routine intentionally uses lightweight grid-icon removal because it can
'   run on every selection change
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetCell             As Excel.Range        'Resolved target cell
    Dim KeepPickerOpen         As Boolean            'True to keep open form alive

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Keep refresh isolated from callers
        On Error GoTo FailSafe

    'Resolve whether the visible picker should stay open
        KeepPickerOpen = Should_KeepPickerOpenOnContextChange()

'------------------------------------------------------------------------------
' CLEAR FIRST
'------------------------------------------------------------------------------
    'Clear transient UI while respecting open-picker lifecycle
        Clear_AllDatePickerUI Not KeepPickerOpen, False

'------------------------------------------------------------------------------
' RESOLVE TARGET CELL
'------------------------------------------------------------------------------
    'Resolve the target cell from ActiveCell or explicit event target
        If IsMissing(vTarget) Then
            If Not Resolve_TargetCell(TargetCell) Then
                If KeepPickerOpen Then M_FormBridge_RefreshFromActiveCell
                Exit Sub
            End If
        Else
            If Not Resolve_TargetCell(TargetCell, vTarget) Then
                If KeepPickerOpen Then M_FormBridge_RefreshFromActiveCell
                Exit Sub
            End If
        End If

'------------------------------------------------------------------------------
' REFRESH OPEN FORM IF KEPT ALIVE
'------------------------------------------------------------------------------
    'Refresh the loaded DatePicker form from the current ActiveCell
        If KeepPickerOpen Then M_FormBridge_RefreshFromActiveCell

'------------------------------------------------------------------------------
' RECREATE GRID ICON IF NEEDED
'------------------------------------------------------------------------------
    'Create the in-grid icon when current context allows it
        If Should_ShowDP_ForCell(TargetCell) Then M_GridIcon_Create TargetCell

'------------------------------------------------------------------------------
' NORMAL EXIT
'------------------------------------------------------------------------------
    'No further action required
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Clear transient UI after a failed refresh without using the expensive purge
        Clear_AllDatePickerUI True, False

End Sub

Private Function Should_KeepPickerOpenOnContextChange() As Boolean

'
'------------------------------------------------------------------------------
'                           SHOULD KEEP PICKER OPEN
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether an already visible DatePicker form should remain open during
'   a context refresh
'
' WHY THIS EXISTS
'   When close-after-selection is disabled, users should be able to keep the
'   DatePicker open and select another worksheet cell as the next write target
'
' INPUTS
'   None
'
' RETURNS
'   True if the visible picker should remain open; otherwise False
'
' BEHAVIOR
'   Keeps the picker open only when it is currently visible and
'   gDP_CloseAfterSelection is False
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   PickerVisible
'   gDP_CloseAfterSelection
'
' NOTES
'   This routine does not inspect the active target cell
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to closing / not preserving the picker
        Should_KeepPickerOpenOnContextChange = False

'------------------------------------------------------------------------------
' RETURN DECISION
'------------------------------------------------------------------------------
    'Keep the visible picker open only when close-after-selection is disabled
        Should_KeepPickerOpenOnContextChange = _
            (PickerVisible And Not gDP_CloseAfterSelection)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Safe default
        Should_KeepPickerOpenOnContextChange = False

End Function

Private Sub Clear_UI_IfIdle( _
    Optional ByVal ClosePicker As Boolean = True, _
    Optional ByVal PurgeGridIcons As Boolean = False)

'
'------------------------------------------------------------------------------
'                           CLEAR UI IF IDLE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears transient DatePicker UI only when no guarded refresh is in progress
'
' WHY THIS EXISTS
'   Some Application events are cleanup boundaries. Guarding those callbacks
'   avoids re-entrant cleanup churn if teardown triggers more Excel events
'
' INPUTS
'   ClosePicker
'     True to close the loaded DatePicker form
'     False to keep the loaded DatePicker form open
'
'   PurgeGridIcons
'     True to use the hard all-workbook grid-icon purge path
'     False to use the lightweight active-context grid-icon removal path
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits if the manager is busy, otherwise delegates to Clear_AllDatePickerUI
'
' ERROR POLICY
'   Best-effort. Does not intentionally raise outward
'
' DEPENDENCIES
'   Clear_AllDatePickerUI
'
' NOTES
'   Save, print, close, and teardown should generally pass PurgeGridIcons = True
'
'   Selection, resize, and window deactivation should generally keep
'   PurgeGridIcons = False for responsiveness
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' GUARD RE-ENTRANT CLEANUP
'------------------------------------------------------------------------------
    'Do nothing if guarded work is already running
        If mIsBusy Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear transient DatePicker UI
        Clear_AllDatePickerUI ClosePicker, PurgeGridIcons

End Sub

'
'==============================================================================
'
'                                UI GATING
'
'==============================================================================

Private Function Should_ShowDP_ForCell(ByVal TargetCell As Excel.Range) As Boolean

'
'------------------------------------------------------------------------------
'                           SHOULD SHOW FOR CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Applies DatePicker in-grid icon gating to one resolved worksheet cell
'
' WHY THIS EXISTS
'   Public and private refresh paths can share the same cell-level gating logic
'
' INPUTS
'   TargetCell
'     Single worksheet cell to evaluate
'
' RETURNS
'   True if the in-grid DatePicker icon should be shown; otherwise False
'
' BEHAVIOR
'   Shows the entry point when the feature is enabled and the cell either:
'     - contains a date-like value
'     - uses a date/time-like number format
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   gDP_ShowGridIcon
'   Is_CellDateValue
'   Is_CellDateFormatted
'
' NOTES
'   This is UI gating only and does not validate write-back semantics
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not showing the DatePicker icon
        Should_ShowDP_ForCell = False

'------------------------------------------------------------------------------
' FEATURE FLAG
'------------------------------------------------------------------------------
    'Exit if the in-grid icon feature is disabled
        If Not gDP_ShowGridIcon Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Exit if no cell was supplied
        If TargetCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' VALUE-BASED GATING
'------------------------------------------------------------------------------
    'Show the entry point for date-like values
        If Is_CellDateValue(TargetCell) Then
            Should_ShowDP_ForCell = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' FORMAT-BASED GATING
'------------------------------------------------------------------------------
    'Show the entry point for date/time-like formats
        If Is_CellDateFormatted(TargetCell) Then
            Should_ShowDP_ForCell = True
            Exit Function
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
    'Safe default
        Should_ShowDP_ForCell = False

End Function

Private Function Is_CellDateValue(ByVal TargetCell As Excel.Range) As Boolean

'
'------------------------------------------------------------------------------
'                           IS CELL DATE VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a cell contains a date-like value
'
' WHY THIS EXISTS
'   A cell may justify showing the DatePicker because its current value is
'   already date-like, even before considering number format
'
' INPUTS
'   TargetCell
'     Single worksheet cell to evaluate
'
' RETURNS
'   True if the cell value is date-like; otherwise False
'
' BEHAVIOR
'   Recognizes explicit Date subtypes and non-empty strings that parse as dates
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   VBA.IsDate
'
' NOTES
'   Numeric serials are handled by format-based gating instead
'
'   Range.Value is used intentionally because Excel may expose formatted date
'   cells as Date variants, which is useful for UI gating
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CellValue              As Variant        'Cell value snapshot

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not date-like
        Is_CellDateValue = False

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Exit if no cell was supplied
        If TargetCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Capture the cell value once
        CellValue = TargetCell.Value

'------------------------------------------------------------------------------
' REJECT NON-USABLE VALUES
'------------------------------------------------------------------------------
    'Reject error values
        If IsError(CellValue) Then Exit Function

    'Reject empty values
        If IsEmpty(CellValue) Then Exit Function

'------------------------------------------------------------------------------
' TEST DATE SUBTYPE
'------------------------------------------------------------------------------
    'Recognize explicit Date variants
        If VarType(CellValue) = vbDate Then
            Is_CellDateValue = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' TEST PARSEABLE STRING
'------------------------------------------------------------------------------
    'Recognize non-empty strings that parse as dates
        If VarType(CellValue) = vbString Then
            If LenB(Trim$(CStr(CellValue))) > 0 Then
                If VBA.IsDate(CStr(CellValue)) Then
                    Is_CellDateValue = True
                    Exit Function
                End If
            End If
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
    'Safe default
        Is_CellDateValue = False

End Function

Private Function Is_CellDateFormatted(ByVal TargetCell As Excel.Range) As Boolean

'
'------------------------------------------------------------------------------
'                           IS CELL DATE FORMATTED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a cell uses a date/time-like number format
'
' WHY THIS EXISTS
'   A cell may be empty or numeric but still intended for date input because of
'   its number format
'
' INPUTS
'   TargetCell
'     Single worksheet cell to evaluate
'
' RETURNS
'   True if the format appears date/time-like; otherwise False
'
' BEHAVIOR
'   Tests both invariant and localized number format strings
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Is_DateFormatString
'
' NOTES
'   This is heuristic and intended only for UI gating
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FormatCode             As String         'Invariant format code
    Dim LocalFormatCode        As String         'Localized format code

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not date/time-formatted
        Is_CellDateFormatted = False

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Exit if no cell was supplied
        If TargetCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' READ FORMATS
'------------------------------------------------------------------------------
    'Read the invariant number format
        FormatCode = CStr(TargetCell.NumberFormat)

    'Read the localized number format
        LocalFormatCode = CStr(TargetCell.NumberFormatLocal)

'------------------------------------------------------------------------------
' TEST INVARIANT FORMAT
'------------------------------------------------------------------------------
    'Return True when the invariant format looks date/time-like
        If Is_DateFormatString(FormatCode) Then
            Is_CellDateFormatted = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' TEST LOCALIZED FORMAT
'------------------------------------------------------------------------------
    'Return True when the localized format looks date/time-like
        If LenB(LocalFormatCode) <> 0 Then
            If Is_DateFormatString(LocalFormatCode) Then
                Is_CellDateFormatted = True
                Exit Function
            End If
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
    'Safe default
        Is_CellDateFormatted = False

End Function

Private Function Is_DateFormatString(ByVal FormatCode As String) As Boolean

'
'------------------------------------------------------------------------------
'                           IS DATE FORMAT STRING
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether an Excel number-format code appears date/time-like
'
' WHY THIS EXISTS
'   Excel number formats can contain literals, locale tags, colors, conditions,
'   and escaped characters. This helper provides a tolerant UI-gating heuristic
'
' INPUTS
'   FormatCode
'     Raw Excel number-format code
'
' RETURNS
'   True if the format looks date/time-like; otherwise False
'
' BEHAVIOR
'   Normalizes the format and scans for conservative date/time token patterns
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Normalize_FormatCodeForDateScan
'
' NOTES
'   This routine is not a full Excel number-format parser
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedFormat       As String         'Normalized format code
    Dim HasDayToken            As Boolean        'Day token detected
    Dim HasMonthToken          As Boolean        'Month or minute token detected
    Dim HasYearToken           As Boolean        'Year token detected
    Dim HasHourToken           As Boolean        'Hour token detected
    Dim HasSecondToken         As Boolean        'Second token detected
    Dim HasAMPMToken           As Boolean        'AM/PM token detected
    Dim HasDateSeparator       As Boolean        'Date separator detected
    Dim HasTimeSeparator       As Boolean        'Time separator detected

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not date/time-like
        Is_DateFormatString = False

'------------------------------------------------------------------------------
' NORMALIZE FORMAT
'------------------------------------------------------------------------------
    'Normalize the raw format string for token scanning
        NormalizedFormat = Normalize_FormatCodeForDateScan(FormatCode)

    'Normalize whitespace and case
        NormalizedFormat = LCase$(Trim$(NormalizedFormat))

'------------------------------------------------------------------------------
' REJECT NON-USABLE FORMATS
'------------------------------------------------------------------------------
    'Reject empty formats
        If LenB(NormalizedFormat) = 0 Then Exit Function

    'Reject General format
        If NormalizedFormat = "general" Then Exit Function

    'Reject Text format
        If NormalizedFormat = "@" Then Exit Function

'------------------------------------------------------------------------------
' DETECT TOKENS
'------------------------------------------------------------------------------
    'Detect day-like tokens
        HasDayToken = _
            (InStr(1, NormalizedFormat, "d", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, "g", vbBinaryCompare) > 0)

    'Detect month / minute token family
        HasMonthToken = _
            (InStr(1, NormalizedFormat, "m", vbBinaryCompare) > 0)

    'Detect year tokens
        HasYearToken = _
            (InStr(1, NormalizedFormat, "y", vbBinaryCompare) > 0)

    'Detect AM/PM tokens
        HasAMPMToken = _
            (InStr(1, NormalizedFormat, "am/pm", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, "a/p", vbBinaryCompare) > 0)

    'Detect hour tokens
        HasHourToken = _
            (InStr(1, NormalizedFormat, "h", vbBinaryCompare) > 0) Or _
            HasAMPMToken

    'Detect second tokens
        HasSecondToken = _
            (InStr(1, NormalizedFormat, "s", vbBinaryCompare) > 0)

'------------------------------------------------------------------------------
' DETECT SEPARATORS
'------------------------------------------------------------------------------
    'Detect common date separators
        HasDateSeparator = _
            (InStr(1, NormalizedFormat, "/", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, "-", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, ".", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, ",", vbBinaryCompare) > 0) Or _
            (InStr(1, NormalizedFormat, " ", vbBinaryCompare) > 0)

    'Detect common time separator
        HasTimeSeparator = _
            (InStr(1, NormalizedFormat, ":", vbBinaryCompare) > 0)

'------------------------------------------------------------------------------
' TIME FORMAT POSITIVES
'------------------------------------------------------------------------------
    'Time separator plus time tokens qualifies as date/time-like
        If HasTimeSeparator Then
            If HasHourToken Or HasMonthToken Or HasSecondToken Then
                Is_DateFormatString = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' DATE FORMAT POSITIVES
'------------------------------------------------------------------------------
    'Visible date separator plus any two date-token families qualifies
        If HasDateSeparator Then
            If (HasDayToken And HasMonthToken) _
            Or (HasDayToken And HasYearToken) _
            Or (HasMonthToken And HasYearToken) Then
                Is_DateFormatString = True
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' LONG DATE POSITIVES
'------------------------------------------------------------------------------
    'All three date-token families qualify even without explicit separators
        If HasDayToken And HasMonthToken And HasYearToken Then
            Is_DateFormatString = True
            Exit Function
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
    'Safe default
        Is_DateFormatString = False

End Function

Private Function Normalize_FormatCodeForDateScan(ByVal FormatCode As String) As String

'
'------------------------------------------------------------------------------
'                    NORMALIZE FORMAT CODE FOR DATE SCAN
'------------------------------------------------------------------------------
' PURPOSE
'   Removes number-format sections that create false positives in date scanning
'
' WHY THIS EXISTS
'   Excel number formats may contain quoted text, bracketed locale/color sections,
'   escaped characters, fill instructions, and spacing instructions. These parts
'   are useful for rendering but noisy for date-intent detection
'
' INPUTS
'   FormatCode
'     Raw Excel number-format code
'
' RETURNS
'   Simplified format string suitable for heuristic scanning
'
' BEHAVIOR
'   Removes quoted literals, removes bracketed sections, skips escaped
'   characters, skips layout payload after underscore and asterisk, and returns
'   retained visible scan characters
'
' ERROR POLICY
'   Safe default. Returns an empty string on error
'
' DEPENDENCIES
'   None
'
' NOTES
'   This routine intentionally does not parse every Excel number-format feature
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index                  As Long           'Character index
    Dim CurChar                As String         'Current character
    Dim OutputText             As String         'Normalized output
    Dim InQuotes               As Boolean        'True inside quoted literal text
    Dim InBrackets             As Boolean        'True inside bracketed sections
    Dim EscapedNext            As Boolean        'True when next character is escaped
    Dim SkipNext               As Boolean        'True when next character is layout payload

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default helper
        On Error GoTo ErrorHandler

    'Default to empty output
        Normalize_FormatCodeForDateScan = vbNullString

'------------------------------------------------------------------------------
' SCAN FORMAT CODE
'------------------------------------------------------------------------------
    'Loop through each character in the format code
        For Index = 1 To Len(FormatCode)

            'Read the current character
                CurChar = Mid$(FormatCode, Index, 1)

            'Skip escaped character payload
                If EscapedNext Then
                    EscapedNext = False

            'Skip layout character payload
                ElseIf SkipNext Then
                    SkipNext = False

            'Backslash escapes the next character
                ElseIf CurChar = "\" Then
                    EscapedNext = True

            'Underscore reserves space equal to the next character
                ElseIf CurChar = "_" Then
                    SkipNext = True

            'Asterisk repeats the next character as fill
                ElseIf CurChar = "*" Then
                    SkipNext = True

            'Double quote toggles quoted literal mode
                ElseIf CurChar = """" Then
                    InQuotes = Not InQuotes

            'Ignore quoted literal content
                ElseIf InQuotes Then
                    'Do nothing

            'Opening bracket starts bracketed section
                ElseIf CurChar = "[" Then
                    InBrackets = True

            'Closing bracket ends bracketed section
                ElseIf CurChar = "]" Then
                    InBrackets = False

            'Ignore bracketed section content
                ElseIf InBrackets Then
                    'Do nothing

            'Retain visible scan character
                Else
                    OutputText = OutputText & CurChar
                End If

        Next Index

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized format code
        Normalize_FormatCodeForDateScan = OutputText

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Safe default
        Normalize_FormatCodeForDateScan = vbNullString

End Function

'
'==============================================================================
'
'                              TARGET RESOLUTION
'
'==============================================================================

Private Function Resolve_TargetCell( _
    ByRef TargetCell As Excel.Range, _
    Optional ByVal vTarget As Variant) As Boolean

'
'------------------------------------------------------------------------------
'                           RESOLVE TARGET CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet cell from an optional event target or from ActiveCell
'
' WHY THIS EXISTS
'   Event handlers should use the real Target range when supplied, while public
'   commands and context refreshes can fall back to ActiveCell
'
' INPUTS
'   TargetCell
'     Output Range receiving the resolved single cell
'
'   vTarget
'     Optional Range-like event target
'
' RETURNS
'   True if a valid single cell was resolved; otherwise False
'
' BEHAVIOR
'   Falls back to ActiveCell when no explicit target is available, requires an
'   explicit target to be a Range object, rejects multi-cell selections, and
'   normalizes merged cells to the merge area's top-left cell
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Try_GetActiveCell
'   Normalize_TargetCell
'
' NOTES
'   Multi-cell selections are intentionally rejected for in-grid icon placement
'
'   A single merged range is accepted and normalized to its top-left cell
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetObject           As Object         'Late-bound event target object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default resolver
        On Error GoTo ErrorHandler

    'Default to unresolved
        Resolve_TargetCell = False

    'Clear output reference
        Set TargetCell = Nothing

'------------------------------------------------------------------------------
' RESOLVE SOURCE
'------------------------------------------------------------------------------
    'Use ActiveCell when no explicit target was supplied
        If IsMissing(vTarget) Then
            If Not Try_GetActiveCell(TargetCell) Then Exit Function

    'Capture explicit object targets
        ElseIf IsObject(vTarget) Then
            Set TargetObject = vTarget

            If TargetObject Is Nothing Then Exit Function
            If Not TypeOf TargetObject Is Excel.Range Then Exit Function

            Set TargetCell = TargetObject

    'Treat Empty as no explicit target supplied
        ElseIf IsEmpty(vTarget) Then
            If Not Try_GetActiveCell(TargetCell) Then Exit Function

    'Treat Null as no explicit target supplied
        ElseIf IsNull(vTarget) Then
            If Not Try_GetActiveCell(TargetCell) Then Exit Function

    'Reject all other explicit target values
        Else
            Exit Function
        End If

'------------------------------------------------------------------------------
' NORMALIZE AND VALIDATE CELL
'------------------------------------------------------------------------------
    'Normalize and validate the resolved target cell
        If Not Normalize_TargetCell(TargetCell) Then Exit Function

'------------------------------------------------------------------------------
' SUCCESS
'------------------------------------------------------------------------------
    'Signal successful target resolution
        Resolve_TargetCell = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear output reference on failure
        Set TargetCell = Nothing

    'Safe default
        Resolve_TargetCell = False

End Function

Private Function Try_GetActiveCell(ByRef TargetCell As Excel.Range) As Boolean

'
'------------------------------------------------------------------------------
'                           TRY GET ACTIVE CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Safely resolves Application.ActiveCell
'
' WHY THIS EXISTS
'   ActiveCell access can fail during no-window states, workbook teardown, or
'   transient activation boundaries
'
' INPUTS
'   TargetCell
'     Output Range receiving ActiveCell when available
'
' RETURNS
'   True if ActiveCell resolved to a single cell; otherwise False
'
' BEHAVIOR
'   Reads Application.ActiveCell, validates the target, and normalizes merged
'   cells to the top-left cell
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Excel.Application.ActiveCell
'   Normalize_TargetCell
'
' NOTES
'   This routine does not validate DatePicker relevance
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default resolver
        On Error GoTo ErrorHandler

    'Default to unresolved
        Try_GetActiveCell = False

    'Clear output reference
        Set TargetCell = Nothing

'------------------------------------------------------------------------------
' READ ACTIVE CELL
'------------------------------------------------------------------------------
    'Capture the current ActiveCell
        Set TargetCell = Excel.Application.ActiveCell

'------------------------------------------------------------------------------
' NORMALIZE AND VALIDATE CELL
'------------------------------------------------------------------------------
    'Normalize and validate the resolved active cell
        If Not Normalize_TargetCell(TargetCell) Then Exit Function

'------------------------------------------------------------------------------
' SUCCESS
'------------------------------------------------------------------------------
    'Signal successful ActiveCell resolution
        Try_GetActiveCell = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear output reference on failure
        Set TargetCell = Nothing

    'Safe default
        Try_GetActiveCell = False

End Function

Private Function Normalize_TargetCell(ByRef TargetCell As Excel.Range) As Boolean

'
'------------------------------------------------------------------------------
'                           NORMALIZE TARGET CELL
'------------------------------------------------------------------------------
' PURPOSE
'   Validates and normalizes a resolved DatePicker target range
'
' WHY THIS EXISTS
'   Target resolution and ActiveCell resolution both need the same policy for
'   multi-cell ranges and merged cells
'
' INPUTS
'   TargetCell
'     Range reference to validate and normalize
'
' RETURNS
'   True when TargetCell is a valid single worksheet cell; otherwise False
'
' BEHAVIOR
'   Accepts normal single cells and single merged ranges. Rejects multi-cell
'   non-merged selections. Normalizes merged cells to the top-left cell
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   Excel.Range
'
' NOTES
'   The normalized reference is returned through the ByRef TargetCell argument
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default resolver
        On Error GoTo ErrorHandler

    'Default to invalid
        Normalize_TargetCell = False

'------------------------------------------------------------------------------
' VALIDATE PRESENCE
'------------------------------------------------------------------------------
    'Exit if no range was supplied
        If TargetCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' HANDLE MERGED CELLS
'------------------------------------------------------------------------------
    'Normalize merged cells to the merge area's top-left cell
        If TargetCell.MergeCells Then
            Set TargetCell = TargetCell.MergeArea.Cells(1, 1)
            Normalize_TargetCell = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' REJECT MULTI-CELL NON-MERGED SELECTIONS
'------------------------------------------------------------------------------
    'Reject multi-cell selections that are not one merged-cell context
        If TargetCell.Cells.CountLarge <> 1 Then Exit Function

'------------------------------------------------------------------------------
' SUCCESS
'------------------------------------------------------------------------------
    'Signal successful target normalization
        Normalize_TargetCell = True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear output reference on failure
        Set TargetCell = Nothing

    'Safe default
        Normalize_TargetCell = False

End Function

'
'==============================================================================
'
'                           USERFORM STATE HELPERS
'
'==============================================================================

Private Function Is_UserFormLoaded(ByVal FormName As String) As Boolean

'
'------------------------------------------------------------------------------
'                           IS USERFORM LOADED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a named UserForm is currently loaded
'
' WHY THIS EXISTS
'   Directly touching a default form instance can instantiate it. This helper
'   inspects only the already-loaded UserForms collection
'
' INPUTS
'   FormName
'     Name of the UserForm to find
'
' RETURNS
'   True if the form is loaded; otherwise False
'
' BEHAVIOR
'   Scans VBA.UserForms and compares names case-insensitively
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   VBA.UserForms
'
' NOTES
'   This routine does not inspect visibility
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurForm                As Object         'Loaded UserForm instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not loaded
        Is_UserFormLoaded = False

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit for blank form names
        If Len(Trim$(FormName)) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SCAN LOADED FORMS
'------------------------------------------------------------------------------
    'Loop through currently loaded UserForms
        For Each CurForm In VBA.UserForms

            'Return True when the requested form is found
                If StrComp(CStr(CurForm.Name), Trim$(FormName), vbTextCompare) = 0 Then
                    Is_UserFormLoaded = True
                    Exit Function
                End If

        Next CurForm

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Safe default
        Is_UserFormLoaded = False

End Function

Private Function Is_UserFormVisible(ByVal FormName As String) As Boolean

'
'------------------------------------------------------------------------------
'                           IS USERFORM VISIBLE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a named loaded UserForm is visible
'
' WHY THIS EXISTS
'   Loaded and visible states are distinct. This helper derives actual visible
'   state without forcing a default form instance into existence
'
' INPUTS
'   FormName
'     Name of the UserForm to find
'
' RETURNS
'   True if the form is loaded and visible; otherwise False
'
' BEHAVIOR
'   Scans VBA.UserForms and reads Visible on the matching loaded form
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   VBA.UserForms
'
' NOTES
'   This routine does not instantiate forms
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurForm                As Object         'Loaded UserForm instance

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Safe-default predicate
        On Error GoTo ErrorHandler

    'Default to not visible
        Is_UserFormVisible = False

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit for blank form names
        If Len(Trim$(FormName)) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SCAN LOADED FORMS
'------------------------------------------------------------------------------
    'Loop through currently loaded UserForms
        For Each CurForm In VBA.UserForms

            'Return actual visibility when the requested form is found
                If StrComp(CStr(CurForm.Name), Trim$(FormName), vbTextCompare) = 0 Then
                    Is_UserFormVisible = CBool(CurForm.Visible)
                    Exit Function
                End If

        Next CurForm

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Safe default
        Is_UserFormVisible = False

End Function

'
'==============================================================================
'
'                                BUSY GUARD
'
'==============================================================================

Private Function Try_BeginBusy() As Boolean

'
'------------------------------------------------------------------------------
'                           TRY BEGIN BUSY
'------------------------------------------------------------------------------
' PURPOSE
'   Attempts to enter the manager re-entrancy guard
'
' WHY THIS EXISTS
'   Excel UI actions can indirectly trigger more Application events. Without a
'   guard, DatePicker refreshes may recurse, flicker, or partially overlap
'
' INPUTS
'   None
'
' RETURNS
'   True if the caller entered the busy state; otherwise False
'
' BEHAVIOR
'   Sets mIsBusy only when no guarded work is already in progress
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   mIsBusy
'
' NOTES
'   Callers that receive True must call Release_Busy
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CHECK CURRENT STATE
'------------------------------------------------------------------------------
    'Reject nested guarded work
        If mIsBusy Then Exit Function

'------------------------------------------------------------------------------
' ENTER BUSY STATE
'------------------------------------------------------------------------------
    'Mark guarded work as in progress
        mIsBusy = True

    'Signal successful guard entry
        Try_BeginBusy = True

End Function

Private Sub Release_Busy()

'
'------------------------------------------------------------------------------
'                           RELEASE BUSY
'------------------------------------------------------------------------------
' PURPOSE
'   Leaves the manager re-entrancy guard
'
' WHY THIS EXISTS
'   Guarded refresh paths must reliably release the busy state on both success
'   and failure paths
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resets mIsBusy to False
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   mIsBusy
'
' NOTES
'   This routine is intentionally simple
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LEAVE BUSY STATE
'------------------------------------------------------------------------------
    'Clear the busy guard
        mIsBusy = False

End Sub

'
'==============================================================================
'
'                               APPLICATION EVENTS
'
'==============================================================================

Private Sub mExcelApp_SheetSelectionChange( _
    ByVal Sh As Object, _
    ByVal Target As Excel.Range)

'
'------------------------------------------------------------------------------
'                           APP SHEET SELECTION CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Handles Excel sheet-selection changes
'
' WHY THIS EXISTS
'   Selection movement is the main trigger for showing, moving, hiding, or
'   refreshing the DatePicker entry point
'
' INPUTS
'   Sh
'     Sheet raising the event
'
'   Target
'     New selection range
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_SelectionChange using the real event Target
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Handle_SelectionChange
'
' NOTES
'   Excel supplies Target as a Range
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Delegate selection-change handling
        Handle_SelectionChange Target

End Sub

Private Sub mExcelApp_SheetChange( _
    ByVal Sh As Object, _
    ByVal Target As Excel.Range)

'
'------------------------------------------------------------------------------
'                           APP SHEET CHANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Handles Excel worksheet content changes
'
' WHY THIS EXISTS
'   Edits can invalidate DatePicker relevance or icon placement
'
' INPUTS
'   Sh
'     Sheet raising the event
'
'   Target
'     Changed range
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_WorksheetChange using the real event Target
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Handle_WorksheetChange
'
' NOTES
'   DatePicker write-back suppresses worksheet events, so this routine mainly
'   handles user edits and non-DatePicker changes
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Delegate worksheet-change handling
        Handle_WorksheetChange Target

End Sub

Private Sub mExcelApp_NewWorkbook(ByVal Wb As Excel.Workbook)

'
'------------------------------------------------------------------------------
'                           APP NEW WORKBOOK
'------------------------------------------------------------------------------
' PURPOSE
'   Handles new workbook creation
'
' WHY THIS EXISTS
'   Workbook creation changes the active Excel context and can invalidate the
'   current DatePicker UI decision
'
' INPUTS
'   Wb
'     Newly created workbook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_WorkbookContextChange
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The workbook argument is not used directly
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the new workbook context
        Handle_WorkbookContextChange

End Sub

Private Sub mExcelApp_WorkbookBeforePrint( _
    ByVal Wb As Excel.Workbook, _
    Cancel As Boolean)

'
'------------------------------------------------------------------------------
'                           APP WORKBOOK BEFORE PRINT
'------------------------------------------------------------------------------
' PURPOSE
'   Clears DatePicker UI before printing
'
' WHY THIS EXISTS
'   Transient UI elements should not interfere with print layout or printed
'   output
'
' INPUTS
'   Wb
'     Workbook being printed
'
'   Cancel
'     Excel print-cancel flag
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears transient UI when the manager is idle
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Printing is a hard cleanup boundary, so the picker is closed and grid icons
'   are purged
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before print using the hard-purge path
        Clear_UI_IfIdle True, True

End Sub

Private Sub mExcelApp_WorkbookBeforeSave( _
    ByVal Wb As Excel.Workbook, _
    ByVal SaveAsUI As Boolean, _
    Cancel As Boolean)

'
'------------------------------------------------------------------------------
'                           APP WORKBOOK BEFORE SAVE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears DatePicker UI before saving
'
' WHY THIS EXISTS
'   Save operations should not preserve stale transient UI artifacts beyond the
'   interactive session
'
' INPUTS
'   Wb
'     Workbook being saved
'
'   SaveAsUI
'     True when Save As UI is shown
'
'   Cancel
'     Excel save-cancel flag
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears transient UI when the manager is idle
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Saving is a hard cleanup boundary, so the picker is closed and grid icons are
'   purged
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before save using the hard-purge path
        Clear_UI_IfIdle True, True

End Sub

Private Sub mExcelApp_WorkbookBeforeClose( _
    ByVal Wb As Excel.Workbook, _
    Cancel As Boolean)

'
'------------------------------------------------------------------------------
'                           APP WORKBOOK BEFORE CLOSE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears DatePicker UI before workbook close
'
' WHY THIS EXISTS
'   Workbook-close transitions are a common source of orphaned UserForms and
'   worksheet shapes in event-driven Excel projects
'
' INPUTS
'   Wb
'     Workbook being closed
'
'   Cancel
'     Excel close-cancel flag
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears transient UI when the manager is idle
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Closing is a hard cleanup boundary, so the picker is closed and grid icons
'   are purged
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before close using the hard-purge path
        Clear_UI_IfIdle True, True

End Sub

Private Sub mExcelApp_WindowResize( _
    ByVal Wb As Excel.Workbook, _
    ByVal Wn As Excel.Window)

'
'------------------------------------------------------------------------------
'                           APP WINDOW RESIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears worksheet-level DatePicker artifacts after Excel window resizing
'
' WHY THIS EXISTS
'   Resizing can invalidate worksheet-shape positioning. It should not
'   necessarily close a modeless picker that the user is actively using
'
' INPUTS
'   Wb
'     Workbook owning the resized window
'
'   Wn
'     Resized Excel window
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears grid artifacts only when the manager is idle
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Resize uses the lightweight cleanup path for responsiveness
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR GRID ARTIFACTS ONLY
'------------------------------------------------------------------------------
    'Clear transient grid UI without closing the modeless picker
        Clear_UI_IfIdle False, False

End Sub

Private Sub mExcelApp_WindowDeactivate( _
    ByVal Wb As Excel.Workbook, _
    ByVal Wn As Excel.Window)

'
'------------------------------------------------------------------------------
'                           APP WINDOW DEACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Removes worksheet-level DatePicker artifacts when an Excel window is
'   deactivated without forcibly closing the modeless picker
'
' WHY THIS EXISTS
'   Window deactivation can occur during legitimate modeless UserForm
'   interaction. Closing the picker here may make the UI feel unstable
'
' INPUTS
'   Wb
'     Workbook owning the deactivated window
'
'   Wn
'     Deactivated Excel window
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears grid artifacts only and keeps the picker form alive
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   WorkbookBeforeClose remains the hard cleanup boundary
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR GRID ARTIFACTS ONLY
'------------------------------------------------------------------------------
    'Clear transient grid UI without closing the modeless picker
        Clear_UI_IfIdle False, False

End Sub

Private Sub mExcelApp_SheetActivate(ByVal Sh As Object)

'
'------------------------------------------------------------------------------
'                           APP SHEET ACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Handles sheet activation
'
' WHY THIS EXISTS
'   Activating a sheet changes the active context and may require a fresh
'   DatePicker UI decision
'
' INPUTS
'   Sh
'     Activated sheet object
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_WorkbookContextChange
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The sheet argument is not used directly
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the newly activated sheet context
        Handle_WorkbookContextChange

End Sub

Private Sub mExcelApp_WorkbookActivate(ByVal Wb As Excel.Workbook)

'
'------------------------------------------------------------------------------
'                           APP WORKBOOK ACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Handles workbook activation
'
' WHY THIS EXISTS
'   Activating a workbook changes the active DatePicker context and may justify
'   recreating the in-grid entry point
'
' INPUTS
'   Wb
'     Activated workbook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_WorkbookContextChange
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The workbook argument is not used directly
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the newly activated workbook context
        Handle_WorkbookContextChange

End Sub

Private Sub mExcelApp_WorkbookDeactivate(ByVal Wb As Excel.Workbook)

'
'------------------------------------------------------------------------------
'                           APP WORKBOOK DEACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears DatePicker UI when a workbook is deactivated
'
' WHY THIS EXISTS
'   Leaving a workbook is a natural boundary for removing transient UI artifacts
'   from no-longer-active contexts
'
' INPUTS
'   Wb
'     Deactivated workbook
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears transient UI when the manager is idle
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Workbook deactivation is treated as a cleanup boundary
'
' UPDATED
'   2026-04-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI on workbook deactivation using the hard-purge path
        Clear_UI_IfIdle True, True

End Sub


