VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDatePickerManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'
'------------------------------------------------------------------------------
' CLASS: cDatePickerManager
'------------------------------------------------------------------------------
' PURPOSE
'   Provides the central Excel Application event manager for the DatePicker
'   project
'
' WHY THIS EXISTS
'   The DatePicker uses several transient Excel / VBA UI surfaces:
'     - modeless runtime UserForm
'     - in-grid DatePicker icon
'     - workbook, worksheet, window, save, print, close, and activation events
'     - optional keyboard shortcut fallback
'
'   Centralizing the Application event flow in one manager class keeps the
'   project predictable, avoids duplicate event wiring, and prevents stale UI
'   artifacts from surviving normal Excel context transitions
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
'     - explicit target-cell resolution
'     - single-source target synchronization between the form and the grid icon
'     - in-grid icon show / move / remove decisions
'     - hard grid-icon purge during save / close / print / teardown paths
'     - loaded DatePicker form cleanup
'     - open-form refresh when close-after-selection is disabled
'     - keyboard shortcut synchronization and teardown
'     - safe reset of transient DatePicker UI
'
' ERROR POLICY
'   Event callbacks are best-effort and must not raise errors into Excel
'
'   Public predicates return safe defaults on error
'
'   Refresh paths use a re-entrancy guard to avoid recursive UI churn
'
'   Cleanup paths suppress expected teardown failures because transient UI
'   artifacts may already have been removed by Excel or by another event path
'
' DEPENDENCIES
'   Excel.Application
'   Excel.Range
'   Excel.Workbook
'   Excel.Window
'   VBA.UserForms
'   DP_Close
'   M_Settings_Load
'   M_KeyboardShortcut_Update
'   M_KeyboardShortcut_Remove
'   M_GridIcon_ShowOrMove
'   M_GridIcon_Hide
'   M_GridIcon_Remove
'   M_GridIcon_PurgeAll
'   M_FormBridge_RefreshFromCell
'   gDP_ShowGridIcon
'   gDP_CloseAfterSelection
'
' NOTES
'   Runtime UserForm control / label event wiring does not belong here. It is
'   handled by cDatePickerLabelHook and UF_DatePicker
'
'   Right-click menu integration is not owned by this class unless a dedicated
'   SheetBeforeRightClick handler is added
'
'   The manager distinguishes between:
'     - the picker being loaded in memory
'     - the picker being visible on screen
'
'   When gDP_CloseAfterSelection is False and the picker is visible, context
'   refreshes keep the form open and refresh it from the same resolved TargetCell
'   used for the in-grid icon
'
'   The resolved TargetCell is the authoritative UI context. Once resolved, the
'   refresh path must not fall back to Excel.Application.ActiveCell
'
'   Date-format detection is intentionally heuristic and is used only for UI
'   gating, not for full semantic format parsing
'
'   M_GridIcon_ShowOrMove is the normal high-frequency eligible-cell path
'
'   M_GridIcon_Hide is the normal high-frequency non-eligible-cell path
'
'   M_GridIcon_Remove is the lightweight active-context cleanup path used when
'   the icon feature is disabled or the current UI context is explicitly cleared
'
'   M_GridIcon_PurgeAll is the heavier all-workbook cleanup path used only for
'   hard lifecycle boundaries
'
'   Application events are active only while this manager instance is kept alive
'   by an external module-level or global object reference
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' APPLICATION EVENT SOURCE
'------------------------------------------------------------------------------
    Private WithEvents mExcelApp        As Excel.Application            'Hooked Excel Application instance
Attribute mExcelApp.VB_VarHelpID = -1

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mIsBusy                     As Boolean                      'True while guarded refresh / teardown is running
    Private mIsHooked                   As Boolean                      'True when the Application event hook was created


'
'------------------------------------------------------------------------------
'
'                               CLASS LIFECYCLE
'
'------------------------------------------------------------------------------
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
'   basis, synchronizes the keyboard shortcut fallback, hooks the Excel
'   Application object, and records whether hook-up succeeded
'
' ERROR POLICY
'   Best-effort initialization. Startup should not hard-fail if Excel is in a
'   transient state during workbook load, add-in load, or VBA reset
'
'   Individual startup failures are written to the Immediate Window and do not
'   prevent later initialization steps from being attempted
'
' DEPENDENCIES
'   M_Settings_Load
'   M_KeyboardShortcut_Update
'   Excel.Application
'
' NOTES
'   Settings are loaded before Application events are hooked so event gating can
'   safely use persisted feature flags such as gDP_ShowGridIcon and
'   gDP_CloseAfterSelection
'
'   This routine assumes that DatePicker setting globals have safe compile-time
'   or module-level defaults in case persisted settings cannot be loaded
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Class_Initialize"

    Dim StepErrNumber           As Long         'Captured startup-step error number
    Dim StepErrDescription      As String       'Captured startup-step error description

'------------------------------------------------------------------------------
' INITIALIZE STATE
'------------------------------------------------------------------------------
    'Clear the busy guard
        mIsBusy = False
    'Clear the hook-state flag
        mIsHooked = False
    'Release any existing Excel Application reference defensively
        Set mExcelApp = Nothing

'------------------------------------------------------------------------------
' ENABLE BEST-EFFORT STARTUP
'------------------------------------------------------------------------------
    'Suppress startup errors while each step is isolated and inspected
        On Error Resume Next

'------------------------------------------------------------------------------
' LOAD SETTINGS
'------------------------------------------------------------------------------
    'Clear any previous startup error
        Err.Clear
    'Load persisted DatePicker settings before events start using them
        M_Settings_Load
    'Log settings-load errors without stopping manager startup
        If Err.Number <> 0 Then
            StepErrNumber = Err.Number
            StepErrDescription = Err.Description
            Debug.Print PROC_NAME & _
                " | Step=SettingsLoad" & _
                " | Error=" & VBA.CStr(StepErrNumber) & _
                " | " & StepErrDescription

            Err.Clear
        End If

'------------------------------------------------------------------------------
' SYNCHRONIZE KEYBOARD SHORTCUT
'------------------------------------------------------------------------------
    'Clear any previous startup error
        Err.Clear
    'Synchronize the keyboard shortcut fallback
        M_KeyboardShortcut_Update
    'Log keyboard-shortcut errors without stopping manager startup
        If Err.Number <> 0 Then
            StepErrNumber = Err.Number
            StepErrDescription = Err.Description
            Debug.Print PROC_NAME & _
                " | Step=KeyboardShortcutUpdate" & _
                " | Error=" & VBA.CStr(StepErrNumber) & _
                " | " & StepErrDescription
            Err.Clear
        End If

'------------------------------------------------------------------------------
' HOOK APPLICATION EVENTS
'------------------------------------------------------------------------------
    'Clear any previous startup error
        Err.Clear
    'Hook the Excel Application object
        Set mExcelApp = Excel.Application
    'Handle Application hook errors without hard-failing startup
        If Err.Number <> 0 Then
            StepErrNumber = Err.Number
            StepErrDescription = Err.Description
            Set mExcelApp = Nothing
            mIsHooked = False
            Debug.Print PROC_NAME & _
                " | Step=ApplicationHook" & _
                " | Error=" & VBA.CStr(StepErrNumber) & _
                " | " & StepErrDescription
            Err.Clear
        Else
            mIsHooked = Not (mExcelApp Is Nothing)
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed startup error
        Err.Clear
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
'   explicit teardown. Cleanup here prevents orphaned forms, worksheet shapes,
'   and application-wide keyboard shortcuts from surviving beyond the manager
'   lifetime
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Marks the manager as busy, releases the Application event hook, removes the
'   keyboard shortcut fallback, performs hard UI cleanup, and resets internal
'   state flags
'
' ERROR POLICY
'   Best-effort teardown. Termination must not raise outward
'
' DEPENDENCIES
'   Clear_AllDatePickerUI
'   M_KeyboardShortcut_Remove
'
' NOTES
'   Teardown uses the hard purge path because it is a lifecycle boundary, not a
'   high-frequency selection refresh
'
'   The Application event hook is released before hard UI cleanup so shape
'   deletion, form closing, or workbook state changes cannot route back into this
'   manager through Excel Application events
'
'   The keyboard shortcut is removed because Application.OnKey assignments are
'   application-wide for the current Excel session
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Suppress teardown errors
        On Error Resume Next
    'Block nested refresh work during teardown
        mIsBusy = True
    'Release the Application event hook before hard UI cleanup
        Set mExcelApp = Nothing
    'Clear the hook-state flag immediately after releasing the event hook
        mIsHooked = False
    'Remove the application-wide keyboard shortcut fallback
        M_KeyboardShortcut_Remove
    'Clear all transient UI artifacts using the hard-purge path
        Clear_AllDatePickerUI True, True

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Clear the busy guard
        mIsBusy = False

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed teardown error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

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
'   Delegates to the shared UserForm-loaded helper, which scans the loaded
'   VBA.UserForms collection without touching the default form instance
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   DP_FORM_NAME
'   Is_UserFormLoaded
'
' NOTES
'   This property deliberately avoids direct references to UF_DatePicker
'
'   Avoiding the default form instance prevents accidental UserForm creation
'   while checking loaded state
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return a safe default unless the loaded-state check succeeds
        Is_PickerLoaded = False
    'Enable safe-default error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return whether the DatePicker form is currently loaded
        Is_PickerLoaded = Is_UserFormLoaded(M_DatePicker.DP_FORM_NAME)
        
'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Clear any suppressed property error
        Err.Clear

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
'   Delegates to the shared UserForm-visible helper, which checks the loaded
'   VBA.UserForms collection without touching the default form instance
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   DP_FORM_NAME
'   Is_UserFormVisible
'
' NOTES
'   This property deliberately avoids direct references to UF_DatePicker
'
'   Avoiding the default form instance prevents accidental UserForm creation
'   while checking visible state
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return a safe default unless the visible-state check succeeds
        PickerVisible = False
    'Enable safe-default error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return whether the DatePicker form is currently visible
        PickerVisible = Is_UserFormVisible(M_DatePicker.DP_FORM_NAME)
    
'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Clear any suppressed property error
        Err.Clear

End Property

Public Property Get Is_Hooked() As Boolean

'
'------------------------------------------------------------------------------
'                           IS HOOKED
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the manager currently has an active Application event hook
'
' WHY THIS EXISTS
'   This lightweight diagnostic property is useful during startup validation,
'   tests, and troubleshooting
'
' INPUTS
'   None
'
' RETURNS
'   True if the manager hook flag is set and the Excel Application reference is
'   available; otherwise False
'
' BEHAVIOR
'   Returns a derived hook state based on both the internal hook-state flag and
'   the stored Excel Application reference
'
' ERROR POLICY
'   Safe default. Returns False on error
'
' DEPENDENCIES
'   mIsHooked
'   mExcelApp
'
' NOTES
'   This is a diagnostic state, not a guarantee that Excel will fire events
'
'   A True result means the manager holds an Application reference and believes
'   the event hook was created successfully
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return a safe default unless the hook-state check succeeds
        Is_Hooked = False
    'Enable safe-default error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return the derived hook-state flag
        Is_Hooked = (mIsHooked And Not (mExcelApp Is Nothing))

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
SafeExit:
    'Clear any suppressed property error
        Err.Clear

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
'   No local error handler is required because this property only reads a
'   private Boolean member and performs no object access or external calls
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN STATE
'------------------------------------------------------------------------------
    'Return the busy-state flag
        Is_Busy = mIsBusy

End Property

'
'------------------------------------------------------------------------------
'
'                               PUBLIC COMMANDS
'
'------------------------------------------------------------------------------

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
'   Cleanup failures are written to the Immediate Window for diagnostics but are
'   not raised to the caller
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME     As String = "cDatePickerManager.Reset_DatePickerUI"
    
'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable best-effort cleanup handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' RESET UI
'------------------------------------------------------------------------------
    'Reset all transient DatePicker UI artifacts using the hard-purge path
        Clear_AllDatePickerUI True, True

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the cleanup handler
        Exit Sub

'------------------------------------------------------------------------------
' CLEAN FAIL
'------------------------------------------------------------------------------
CleanFail:
    'Write cleanup diagnostics without interrupting the caller
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(Err.Number) & _
            " | " & Err.Description

    'Clear the suppressed cleanup error
        Err.Clear

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
'   Best-effort event handling. Failures trigger lightweight cleanup and do not
'   raise into Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Clear_AllDatePickerUI
'   Release_Busy
'
' NOTES
'   When gDP_CloseAfterSelection is False and the picker is visible, this routine
'   keeps the form open and refreshes it from the resolved TargetCell
'
'   This routine is a public event-facing wrapper. The detailed UI decision tree
'   remains centralized in Refresh_CurrentUI
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Handle_SelectionChange"

    Dim GuardWasEntered         As Boolean      'True when this routine owns the busy guard
    Dim TargetWasProvided       As Boolean      'True when an explicit target was supplied
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from recursive or failing UI event handling
        On Error GoTo FailSafe
    'Capture whether the caller supplied an explicit target
        TargetWasProvided = VBA.IsMissing(vTarget) = False

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
    'Refresh DatePicker UI for the supplied selection target
        If TargetWasProvided Then
            Refresh_CurrentUI vTarget
        Else
            Refresh_CurrentUI
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors while releasing the guard
        On Error Resume Next
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy
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
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress failure-handling errors
        On Error Resume Next
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | TargetWasProvided=" & VBA.CStr(TargetWasProvided) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear any suppressed failure-handling error
        Err.Clear
    'Leave through the shared guard-release path
        Resume CleanExit

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
'   Enters the manager re-entrancy guard, refreshes DatePicker UI for the changed
'   worksheet context, and releases the guard deterministically
'
' ERROR POLICY
'   Best-effort event handling. Failures trigger lightweight cleanup and do not
'   raise into Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Clear_AllDatePickerUI
'   Release_Busy
'
' NOTES
'   This routine is intentionally thin because the detailed refresh policy is
'   centralized in Refresh_CurrentUI
'
'   The same guarded policy is used for selection changes and worksheet changes
'   so UI cleanup, icon refresh, and open-form synchronization remain consistent
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Handle_WorksheetChange"

    Dim GuardWasEntered         As Boolean      'True when this routine owns the busy guard
    Dim TargetWasProvided       As Boolean      'True when an explicit target was supplied
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from recursive or failing UI event handling
        On Error GoTo FailSafe
    'Capture whether the caller supplied an explicit target
        TargetWasProvided = VBA.IsMissing(vTarget) = False

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
    'Refresh DatePicker UI for the supplied worksheet-change target
        If TargetWasProvided Then
            Refresh_CurrentUI vTarget
        Else
            Refresh_CurrentUI
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors while releasing the guard
        On Error Resume Next
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy
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
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress failure-handling errors
        On Error Resume Next
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | TargetWasProvided=" & VBA.CStr(TargetWasProvided) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear any suppressed failure-handling error
        Err.Clear
    'Leave through the shared guard-release path
        Resume CleanExit

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
'   Enters the manager re-entrancy guard, lets Refresh_CurrentUI resolve the
'   current active Excel context into an authoritative TargetCell, refreshes
'   DatePicker UI from that TargetCell, and releases the guard deterministically
'
' ERROR POLICY
'   Best-effort event handling. Failures trigger lightweight cleanup and do not
'   raise into Excel event flow
'
' DEPENDENCIES
'   Try_BeginBusy
'   Refresh_CurrentUI
'   Clear_AllDatePickerUI
'   Release_Busy
'
' NOTES
'   This routine is intentionally thin because the detailed refresh policy is
'   centralized in Refresh_CurrentUI
'
'   Workbook and worksheet context changes can occur when no explicit event
'   Target range is available. Refresh_CurrentUI resolves the current active
'   context when possible and then uses the resolved TargetCell as the single
'   source of truth
'
'   Cases where no valid TargetCell can be resolved are left to Refresh_CurrentUI
'   and the fail-safe cleanup path
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Handle_WorkbookContextChange"
    
    Dim GuardWasEntered         As Boolean          'True when this routine owns the busy guard
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

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
    'Refresh DatePicker UI by resolving the current active context into a TargetCell
        Refresh_CurrentUI

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Suppress cleanup errors while releasing the guard
        On Error Resume Next
    'Release the busy guard only when this routine acquired it
        If GuardWasEntered Then Release_Busy
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
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress failure-handling errors
        On Error Resume Next
    'Clear transient UI after a failed refresh without expensive purge
        Clear_AllDatePickerUI True, False
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear any suppressed failure-handling error
        Err.Clear
    'Leave through the shared guard-release path
        Resume CleanExit

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
'   Best-effort cleanup. Missing artifacts are safe no-ops and cleanup errors are
'   intentionally suppressed
'
' DEPENDENCIES
'   DP_Close
'   M_GridIcon_Remove
'   M_GridIcon_PurgeAll
'
' NOTES
'   Use PurgeGridIcons = False for scoped cleanup of the currently tracked
'   grid icon when the DatePicker UI context is explicitly cleared
'
'   Use PurgeGridIcons = True for hard boundaries such as save, close, print,
'   explicit reset, and manager teardown
'
'   Do not use this routine for normal high-frequency eligible / non-eligible
'   selection refreshes. Refresh_CurrentUI should use M_GridIcon_ShowOrMove
'   for eligible cells and M_GridIcon_Hide for non-eligible cells
'
'   This routine does not manage the manager busy guard directly. Callers that
'   invoke it from event flow are responsible for guard ownership
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Clear_AllDatePickerUI"

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
    'Use the hard all-workbook grid-icon purge path when requested
        If PurgeGridIcons Then
            M_GridIcon_PurgeAll
    'Otherwise use the lightweight active-context grid-icon removal path
        Else
            M_GridIcon_Remove
        End If

'------------------------------------------------------------------------------
' RESTORE ERROR HANDLING
'------------------------------------------------------------------------------
    'Clear any suppressed cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Sub

Public Function Should_ShowGridIcon(Optional ByVal vTarget As Variant) As Boolean

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
'   Should_ShowGridIcon_ForCell
'
' NOTES
'   This routine does not create, move, remove, or refresh UI
'
'   This routine does not require Application.EnableEvents to be True because it
'   is a direct predicate, not an Excel event callback
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Should_ShowGridIcon"
    
    Dim TargetCell              As Excel.Range      'Resolved single target cell
    Dim TargetWasProvided       As Boolean          'True when an explicit target was supplied

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not showing the DatePicker entry point
        Should_ShowGridIcon = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler
    'Capture whether an explicit target was supplied
        TargetWasProvided = Not VBA.IsMissing(vTarget)

'------------------------------------------------------------------------------
' RESOLVE TARGET CELL
'------------------------------------------------------------------------------
    'Resolve the target cell from the explicit target when supplied
        If TargetWasProvided Then
            If Not Resolve_TargetCell(TargetCell, vTarget) Then
                Exit Function
            End If
    'Otherwise resolve the target cell from ActiveCell
        Else
            If Not Resolve_TargetCell(TargetCell) Then
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' APPLY CELL GATING
'------------------------------------------------------------------------------
    'Return the cell-level gating decision
        Should_ShowGridIcon = Should_ShowGridIcon_ForCell(TargetCell)

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
        Should_ShowGridIcon = False
    'Clear any suppressed predicate error
        Err.Clear

End Function

Private Sub Refresh_CurrentUI(Optional ByVal vTarget As Variant)

'
'------------------------------------------------------------------------------
'                           REFRESH CURRENT UI
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the manager's resolved-target DatePicker UI refresh policy
'
' WHY THIS EXISTS
'   Selection, edit, workbook, sheet, and window events all need consistent UI
'   lifecycle handling without deleting and recreating the in-grid icon
'   unnecessarily
'
' INPUTS
'   vTarget
'     Optional event Target Range
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves one authoritative target cell, keeps or closes the visible picker
'   according to the close-after-selection policy, refreshes the open form from
'   that same resolved target cell, and either moves / shows or hides the
'   in-grid icon
'
'   When the form bridge fails, the failure is logged and the picker is closed
'   on a best-effort basis to avoid leaving stale form state on screen
'
' ERROR POLICY
'   Best-effort refresh. Failures trigger lightweight cleanup and do not raise
'   outward
'
' DEPENDENCIES
'   Should_KeepPickerOpenOnContextChange
'   Resolve_TargetCell
'   DP_Close
'   M_FormBridge_RefreshFromCell
'   Should_ShowGridIcon_ForCell
'   M_GridIcon_ShowOrMove
'   M_GridIcon_Hide
'   Clear_AllDatePickerUI
'
' NOTES
'   This routine deliberately does not clear the grid icon before deciding
'   whether the new target should show it
'
'   For eligible cells, the high-frequency path moves or reuses the existing
'   icon instead of deleting and recreating it
'
'   For non-eligible cells, the existing icon is hidden rather than deleted
'   so selection refresh remains cheap and flicker-resistant
'
'   The resolved TargetCell is the single source of truth for both the picker
'   form and the in-grid icon
'
'   This routine no longer falls back to ActiveCell after target resolution
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Refresh_CurrentUI"

    Dim TargetCell              As Excel.Range      'Resolved authoritative target cell
    
    Dim KeepPickerOpen          As Boolean          'True to keep open form alive
    Dim TargetWasProvided       As Boolean          'True when an explicit target was supplied
    Dim TargetWasResolved       As Boolean          'True when a valid target cell was resolved
    Dim ShowGridIcon            As Boolean          'True when the icon should be shown
    Dim TargetAddress           As String           'Diagnostic target address
    Dim BridgeErrNumber         As Long             'Captured form-bridge error number
    Dim BridgeErrDescription    As String           'Captured form-bridge error description
    Dim CloseErrNumber          As Long             'Captured picker-close error number
    Dim CloseErrDescription     As String           'Captured picker-close error description
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Keep refresh isolated from callers
        On Error GoTo FailSafe
    'Capture whether an explicit target was supplied
        TargetWasProvided = (VBA.IsMissing(vTarget) = False)
    'Resolve whether the visible picker should stay open
        KeepPickerOpen = Should_KeepPickerOpenOnContextChange()

'------------------------------------------------------------------------------
' RESOLVE TARGET CELL
'------------------------------------------------------------------------------
    'Resolve the target cell from the explicit event target when supplied
        If TargetWasProvided Then
            TargetWasResolved = Resolve_TargetCell(TargetCell, vTarget)
    'Otherwise resolve the target cell from the current ActiveCell
        Else
            TargetWasResolved = Resolve_TargetCell(TargetCell)
        End If

'------------------------------------------------------------------------------
' HANDLE UNRESOLVED TARGET
'------------------------------------------------------------------------------
    'Close transient UI when no authoritative target cell is available
        If Not TargetWasResolved Then
            Clear_AllDatePickerUI True, False
            'Exit after unresolved-target cleanup
                Exit Sub
        End If

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC TARGET
'------------------------------------------------------------------------------
    'Suppress target-address diagnostic failures
        On Error Resume Next
    'Capture the resolved target address for diagnostics
        TargetAddress = VBA.CStr(TargetCell.Address(External:=True))
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore refresh error handling
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' DECIDE GRID ICON VISIBILITY
'------------------------------------------------------------------------------
    'Evaluate whether the resolved target should show the in-grid icon
        ShowGridIcon = Should_ShowGridIcon_ForCell(TargetCell)

'------------------------------------------------------------------------------
' UPDATE PICKER FORM LIFECYCLE
'------------------------------------------------------------------------------
    'Close the picker when the current policy does not keep it open
        If Not KeepPickerOpen Then
            'Reset close diagnostics
                CloseErrNumber = 0
            'Reset close diagnostic description
                CloseErrDescription = vbNullString
            'Suppress close errors while capturing diagnostics
                On Error Resume Next
            'Close the DatePicker form
                DP_Close
            'Capture picker-close error number
                CloseErrNumber = Err.Number
            'Capture picker-close error description
                CloseErrDescription = Err.Description
            'Clear any suppressed picker-close error
                Err.Clear
            'Restore refresh error handling
                On Error GoTo FailSafe
            'Log picker-close failure without interrupting refresh flow
                If CloseErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=DP_Close" & _
                        " | Target=" & TargetAddress & _
                        " | Error=" & VBA.CStr(CloseErrNumber) & _
                        " | " & CloseErrDescription
                End If
    'Otherwise refresh the open picker from the same resolved target cell
        Else
            'Reset bridge diagnostics
                BridgeErrNumber = 0
            'Reset bridge diagnostic description
                BridgeErrDescription = vbNullString
            'Suppress bridge errors while capturing diagnostics
                On Error Resume Next
            'Refresh the visible picker from the resolved target cell
                M_FormBridge_RefreshFromCell TargetCell
            'Capture form-bridge error number
                BridgeErrNumber = Err.Number
            'Capture form-bridge error description
                BridgeErrDescription = Err.Description
            'Clear any suppressed form-bridge error
                Err.Clear
            'Restore refresh error handling
                On Error GoTo FailSafe
            'Log form-bridge failure without interrupting Excel event flow
                If BridgeErrNumber <> 0 Then
                    Debug.Print PROC_NAME & _
                        " | Step=M_FormBridge_RefreshFromCell" & _
                        " | Target=" & TargetAddress & _
                        " | Error=" & VBA.CStr(BridgeErrNumber) & _
                        " | " & BridgeErrDescription
                    'Reset close diagnostics
                        CloseErrNumber = 0
                    'Reset close diagnostic description
                        CloseErrDescription = vbNullString
                    'Suppress close errors while capturing diagnostics
                        On Error Resume Next
                    'Close the picker to avoid stale visible form state
                        DP_Close
                    'Capture picker-close error number
                        CloseErrNumber = Err.Number
                    'Capture picker-close error description
                        CloseErrDescription = Err.Description
                    'Clear any suppressed picker-close error
                        Err.Clear
                    'Restore refresh error handling
                        On Error GoTo FailSafe
                    'Log picker-close failure after bridge failure
                        If CloseErrNumber <> 0 Then
                            Debug.Print PROC_NAME & _
                                " | Step=DP_CloseAfterBridgeFailure" & _
                                " | Target=" & TargetAddress & _
                                " | Error=" & VBA.CStr(CloseErrNumber) & _
                                " | " & CloseErrDescription
                        End If
                End If
        End If

'------------------------------------------------------------------------------
' UPDATE GRID ICON
'------------------------------------------------------------------------------
    'Move or create the icon when the resolved target allows it
        If ShowGridIcon Then
            M_GridIcon_ShowOrMove TargetCell
    'Otherwise hide the current icon
        Else
            M_GridIcon_Hide
        End If

'------------------------------------------------------------------------------
' NORMAL EXIT
'------------------------------------------------------------------------------
    'Exit after successful refresh
        Exit Sub

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Suppress failure-cleanup errors
        On Error Resume Next
    'Clear transient UI after a failed refresh without using the expensive purge
        Clear_AllDatePickerUI True, False
    'Write refresh diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | TargetWasProvided=" & VBA.CStr(TargetWasProvided) & _
            " | TargetWasResolved=" & VBA.CStr(TargetWasResolved) & _
            " | Target=" & TargetAddress & _
            " | KeepPickerOpen=" & VBA.CStr(KeepPickerOpen) & _
            " | ShowGridIcon=" & VBA.CStr(ShowGridIcon) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription

    'Clear any suppressed failure-cleanup error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

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
'   This routine does not depend on Application.EnableEvents because it is a
'   direct internal predicate, not an Excel event callback
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim IsVisiblePicker         As Boolean        'True when the DatePicker form is visible
    Dim CloseAfterSelection     As Boolean        'Current close-after-selection setting

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to closing / not preserving the picker
        Should_KeepPickerOpenOnContextChange = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CAPTURE STATE
'------------------------------------------------------------------------------
    'Capture the current picker visibility state
        IsVisiblePicker = PickerVisible
    'Exit when there is no visible picker to preserve
        If Not IsVisiblePicker Then Exit Function
    'Capture the close-after-selection setting
        CloseAfterSelection = gDP_CloseAfterSelection

'------------------------------------------------------------------------------
' RETURN DECISION
'------------------------------------------------------------------------------
    'Keep the picker open only when close-after-selection is disabled
        Should_KeepPickerOpenOnContextChange = (CloseAfterSelection = False)

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
        Should_KeepPickerOpenOnContextChange = False
    'Clear any suppressed predicate error
        Err.Clear

End Function

Private Sub Clear_UI_IfIdle( _
    Optional ByVal ClosePicker As Boolean = True, _
    Optional ByVal PurgeGridIcons As Boolean = False)

'
'------------------------------------------------------------------------------
'                           CLEAR UI IF IDLE
'------------------------------------------------------------------------------
' PURPOSE
'   Clears transient DatePicker UI only when no guarded refresh or teardown work
'   is in progress
'
' WHY THIS EXISTS
'   Some Application events request cleanup at lifecycle boundaries. However,
'   cleanup should not interrupt an active guarded refresh because that can
'   create re-entrant form closing, shape deletion, or partially completed UI
'   transitions
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
'   Exits when the manager is busy, otherwise delegates to Clear_AllDatePickerUI
'
' ERROR POLICY
'   Best-effort. Does not intentionally raise outward
'
' DEPENDENCIES
'   Clear_AllDatePickerUI
'
' NOTES
'   Save, print, close, and workbook deactivation may request hard cleanup with
'   PurgeGridIcons = True
'
'   A hard cleanup request is not forced when mIsBusy = True
'
'   Skipping cleanup while busy is intentional because the active guarded path is
'   already responsible for reaching a safe UI state or fail-safe cleanup
'
'   Class_Terminate remains the true forced teardown boundary because it calls
'   Clear_AllDatePickerUI directly after releasing the Application event hook
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Clear_UI_IfIdle"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable best-effort cleanup handling
        On Error GoTo CleanFail

'------------------------------------------------------------------------------
' GUARD RE-ENTRANT CLEANUP
'------------------------------------------------------------------------------
    'Skip cleanup while guarded work is already running
        If mIsBusy Then
            Debug.Print PROC_NAME & _
                " | Step=SkippedBusy" & _
                " | ClosePicker=" & VBA.CStr(ClosePicker) & _
                " | PurgeGridIcons=" & VBA.CStr(PurgeGridIcons)
            Exit Sub
        End If

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear transient DatePicker UI
        Clear_AllDatePickerUI ClosePicker, PurgeGridIcons

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the cleanup handler
        Exit Sub

'------------------------------------------------------------------------------
' CLEAN FAIL
'------------------------------------------------------------------------------
CleanFail:
    'Write cleanup diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | ClosePicker=" & VBA.CStr(ClosePicker) & _
            " | PurgeGridIcons=" & VBA.CStr(PurgeGridIcons) & _
            " | Error=" & VBA.CStr(Err.Number) & _
            " | " & Err.Description

    'Clear the suppressed cleanup error
        Err.Clear

End Sub

'
'------------------------------------------------------------------------------
'
'                                UI GATING
'
'------------------------------------------------------------------------------
'


Private Function Should_ShowGridIcon_ForCell(ByVal TargetCell As Excel.Range) As Boolean

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
'   The caller should normally pass a resolved single cell. This function still
'   rejects invalid or multi-cell targets defensively
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not showing the DatePicker icon
        Should_ShowGridIcon_ForCell = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

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
    'Exit if the supplied target is not a single cell
        If TargetCell.Cells.CountLarge <> 1 Then Exit Function

'------------------------------------------------------------------------------
' VALUE-BASED GATING
'------------------------------------------------------------------------------
    'Show the entry point for date-like values
        If Is_CellDateValue(TargetCell) Then
            Should_ShowGridIcon_ForCell = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' FORMAT-BASED GATING
'------------------------------------------------------------------------------
    'Show the entry point for date/time-like formats
        If Is_CellDateFormatted(TargetCell) Then
            Should_ShowGridIcon_ForCell = True
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
    'Return the safe default
        Should_ShowGridIcon_ForCell = False
    'Clear any suppressed predicate error
        Err.Clear

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
'   Numeric serials are intentionally handled by format-based gating instead
'
'   Range.Value is used intentionally because Excel may expose formatted date
'   cells as Date variants, which is useful for UI gating
'
'   String parsing uses VBA.IsDate and is therefore locale-sensitive. This is
'   acceptable for UI gating but should not be treated as strict write-back
'   validation
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CellValue               As Variant        'Cell value snapshot
    Dim CellType                As VbVarType      'Cell value variant subtype
    Dim CellText                As String         'Trimmed text value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not date-like
        Is_CellDateValue = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Exit if no cell was supplied
        If TargetCell Is Nothing Then Exit Function
    'Exit if the supplied target is not a single cell
        If TargetCell.Cells.CountLarge <> 1 Then Exit Function

'------------------------------------------------------------------------------
' READ VALUE
'------------------------------------------------------------------------------
    'Capture the cell value once
        CellValue = TargetCell.Value

'------------------------------------------------------------------------------
' REJECT NON-USABLE VALUES
'------------------------------------------------------------------------------
    'Reject array values
        If VBA.IsArray(CellValue) Then Exit Function
    'Reject error values
        If VBA.IsError(CellValue) Then Exit Function
    'Reject empty values
        If VBA.IsEmpty(CellValue) Then Exit Function
    'Reject Null values
        If VBA.IsNull(CellValue) Then Exit Function
    'Capture the variant subtype once
        CellType = VBA.VarType(CellValue)
    'Reject Boolean values
        If CellType = vbBoolean Then Exit Function

'------------------------------------------------------------------------------
' TEST DATE SUBTYPE
'------------------------------------------------------------------------------
    'Recognize explicit Date variants
        If CellType = vbDate Then
            Is_CellDateValue = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' TEST PARSEABLE STRING
'------------------------------------------------------------------------------
    'Test only string values through locale-aware date parsing
        If CellType = vbString Then
            'Trim the cell text once
                CellText = VBA.Trim$(VBA.CStr(CellValue))
            'Exit when the string is blank after trimming
                If VBA.Len(CellText) = 0 Then Exit Function
            'Recognize non-empty strings that parse as dates
                If VBA.IsDate(CellText) Then
                    Is_CellDateValue = True
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
    'Return the safe default
        Is_CellDateValue = False
    'Clear any suppressed predicate error
        Err.Clear

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
'   Tests the invariant number format first and, only if needed, the localized
'   number format
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
'   NumberFormat and NumberFormatLocal are read as Variant first because Excel
'   can return Null for heterogeneous ranges. This function still rejects
'   multi-cell targets defensively, but the Variant handling keeps the routine
'   safer if called outside the normal resolved-cell path
'
'   The localized format check improves behavior in non-English Excel
'   installations, but final detection quality depends on Is_DateFormatString
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FormatValue             As Variant        'Invariant number format value
    Dim LocalFormatValue        As Variant        'Localized number format value
    Dim FormatCode              As String         'Invariant format code
    Dim LocalFormatCode         As String         'Localized format code

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not date/time-formatted
        Is_CellDateFormatted = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE TARGET
'------------------------------------------------------------------------------
    'Exit if no cell was supplied
        If TargetCell Is Nothing Then Exit Function
    'Exit if the supplied target is not a single cell
        If TargetCell.Cells.CountLarge <> 1 Then Exit Function

'------------------------------------------------------------------------------
' READ INVARIANT FORMAT
'------------------------------------------------------------------------------
    'Read the invariant number format as Variant
        FormatValue = TargetCell.NumberFormat
    'Skip the invariant format when Excel returns an unusable value
        If Not VBA.IsError(FormatValue) Then
            If Not VBA.IsNull(FormatValue) Then
                If Not VBA.IsEmpty(FormatValue) Then
                    'Convert the invariant format to text
                        FormatCode = VBA.CStr(FormatValue)
                    'Return True when the invariant format looks date/time-like
                        If Is_DateFormatString(FormatCode) Then
                            Is_CellDateFormatted = True
                            Exit Function
                        End If
                End If
            End If
        End If

'------------------------------------------------------------------------------
' READ LOCALIZED FORMAT
'------------------------------------------------------------------------------
    'Read the localized number format as Variant
        LocalFormatValue = TargetCell.NumberFormatLocal
    'Skip the localized format when Excel returns an unusable value
        If Not VBA.IsError(LocalFormatValue) Then
            If Not VBA.IsNull(LocalFormatValue) Then
                If Not VBA.IsEmpty(LocalFormatValue) Then
                    'Convert the localized format to text
                        LocalFormatCode = VBA.CStr(LocalFormatValue)
                    'Return True when the localized format looks date/time-like
                        If VBA.Len(LocalFormatCode) > 0 Then
                            If Is_DateFormatString(LocalFormatCode) Then
                                Is_CellDateFormatted = True
                                Exit Function
                            End If
                        End If
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
    'Return the safe default
        Is_CellDateFormatted = False
    'Clear any suppressed predicate error
        Err.Clear

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
'   The result is intended only for UI gating. It should not be used as strict
'   semantic validation that a value is a date or time
'
'   The scan intentionally combines token detection with separator detection to
'   reduce false positives from ordinary numeric, text, or accounting formats
'
'   The token scan includes common localized date tokens because
'   NumberFormatLocal can differ from invariant NumberFormat
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NormalizedFormat        As String        'Normalized format code
    Dim HasDayToken             As Boolean       'Day token detected
    Dim HasMonthToken           As Boolean       'Month or minute token detected
    Dim HasYearToken            As Boolean       'Year token detected
    Dim HasHourToken            As Boolean       'Hour token detected
    Dim HasSecondToken          As Boolean       'Second token detected
    Dim HasAMPMToken            As Boolean       'AM/PM token detected
    Dim HasDateSeparator        As Boolean       'Date separator detected
    Dim HasTimeSeparator        As Boolean       'Time separator detected

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not date/time-like
        Is_DateFormatString = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE FORMAT
'------------------------------------------------------------------------------
    'Normalize the raw format string for token scanning
        NormalizedFormat = Normalize_FormatCodeForDateScan(FormatCode)
    'Normalize whitespace and case
        NormalizedFormat = VBA.LCase$(VBA.Trim$(NormalizedFormat))

'------------------------------------------------------------------------------
' REJECT NON-USABLE FORMATS
'------------------------------------------------------------------------------
    'Reject empty formats
        If VBA.Len(NormalizedFormat) = 0 Then Exit Function
    'Reject General format
        If NormalizedFormat = "general" Then Exit Function
    'Reject Text format
        If NormalizedFormat = "@" Then Exit Function

'------------------------------------------------------------------------------
' DETECT DATE TOKENS
'------------------------------------------------------------------------------
    'Detect day-like tokens from invariant and common localized formats
        HasDayToken = _
            (VBA.InStr(1, NormalizedFormat, "d", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "g", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "t", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "j", vbBinaryCompare) > 0)
    'Detect month / minute token family
        HasMonthToken = _
            (VBA.InStr(1, NormalizedFormat, "m", vbBinaryCompare) > 0)
    'Detect year-like tokens from invariant and common localized formats
        HasYearToken = _
            (VBA.InStr(1, NormalizedFormat, "y", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "aa", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "jj", vbBinaryCompare) > 0)

'------------------------------------------------------------------------------
' DETECT TIME TOKENS
'------------------------------------------------------------------------------
    'Detect AM/PM tokens
        HasAMPMToken = _
            (VBA.InStr(1, NormalizedFormat, "am/pm", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "a/p", vbBinaryCompare) > 0)
    'Detect hour tokens
        HasHourToken = _
            (VBA.InStr(1, NormalizedFormat, "h", vbBinaryCompare) > 0) Or _
            HasAMPMToken
    'Detect second tokens
        HasSecondToken = _
            (VBA.InStr(1, NormalizedFormat, "s", vbBinaryCompare) > 0)

'------------------------------------------------------------------------------
' DETECT SEPARATORS
'------------------------------------------------------------------------------
    'Detect common date separators
        HasDateSeparator = _
            (VBA.InStr(1, NormalizedFormat, "/", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, "-", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, ".", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, ",", vbBinaryCompare) > 0) Or _
            (VBA.InStr(1, NormalizedFormat, " ", vbBinaryCompare) > 0)
    'Detect common time separator
        HasTimeSeparator = _
            (VBA.InStr(1, NormalizedFormat, ":", vbBinaryCompare) > 0)

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
' COMPACT DATE FORMAT POSITIVES
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
    'Return the safe default
        Is_DateFormatString = False
    'Clear any suppressed predicate error
        Err.Clear

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
'   Simplified format string suitable for heuristic date/time scanning
'
' BEHAVIOR
'   Removes quoted literals, removes non-time bracketed sections, skips escaped
'   literal characters, skips layout payload after underscore and asterisk, and
'   returns retained scan characters
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
'   Bracketed elapsed-time tokens such as [h], [m], and [s] are preserved because
'   they indicate time-like intent
'
'   Other bracketed sections are removed because they usually contain colors,
'   conditions, locale tags, currency tags, or other rendering metadata
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Index                   As Long          'Character index
    Dim CurChar                 As String        'Current character
    Dim OutputText              As String        'Normalized output
    Dim BracketText             As String        'Captured bracket payload
    Dim BracketToken            As String        'Normalized bracket payload
    Dim InQuotes                As Boolean       'True inside quoted literal text
    Dim InBrackets              As Boolean       'True inside bracketed sections
    Dim EscapedNext             As Boolean       'True when next character is escaped
    Dim SkipNext                As Boolean       'True when next character is layout payload

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to empty output
        Normalize_FormatCodeForDateScan = vbNullString
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SCAN FORMAT CODE
'------------------------------------------------------------------------------
    'Loop through each character in the format code
        For Index = 1 To VBA.Len(FormatCode)
            'Read the current character
                CurChar = VBA.Mid$(FormatCode, Index, 1)
            'Skip escaped character payload
                If EscapedNext Then
                    EscapedNext = False
            'Skip layout character payload
                ElseIf SkipNext Then
                    SkipNext = False
            'Backslash escapes the next literal character
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
            'Opening bracket starts bracketed section capture
                ElseIf CurChar = "[" Then
                    InBrackets = True
                    BracketText = vbNullString
            'Closing bracket ends bracketed section capture
                ElseIf CurChar = "]" Then
                    InBrackets = False
                    'Normalize the bracket payload
                        BracketToken = VBA.LCase$(VBA.Trim$(BracketText))
                    'Preserve elapsed-time hour tokens
                        If BracketToken = "h" Or BracketToken = "hh" Then
                            OutputText = OutputText & "h"
                    'Preserve elapsed-time minute tokens
                        ElseIf BracketToken = "m" Or BracketToken = "mm" Then
                            OutputText = OutputText & "m"
                    'Preserve elapsed-time second tokens
                        ElseIf BracketToken = "s" Or BracketToken = "ss" Then
                            OutputText = OutputText & "s"
                    End If
                    'Clear the bracket payload
                        BracketText = vbNullString
            'Capture bracketed section content
                ElseIf InBrackets Then
                    BracketText = BracketText & CurChar
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
    'Return the safe default
        Normalize_FormatCodeForDateScan = vbNullString
    'Clear any suppressed normalization error
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                              TARGET RESOLUTION
'
'------------------------------------------------------------------------------

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
'   explicit object target to be an Excel.Range, rejects multi-cell selections,
'   and normalizes merged cells to the merge area's top-left cell
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
'   Empty and Null targets are treated like an omitted target and therefore fall
'   back to ActiveCell
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim TargetObject            As Object        'Late-bound event target object

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to unresolved
        Resolve_TargetCell = False
    'Clear output reference
        Set TargetCell = Nothing
    'Enable safe-default resolver handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESOLVE SOURCE
'------------------------------------------------------------------------------
    'Use ActiveCell when no explicit target was supplied
        If VBA.IsMissing(vTarget) Then
            If Not Try_GetActiveCell(TargetCell) Then
                Exit Function
            End If
    'Capture explicit object targets
        ElseIf VBA.IsObject(vTarget) Then
            Set TargetObject = vTarget
            'Reject a missing object target
                If TargetObject Is Nothing Then Exit Function
            'Reject object targets that are not Excel ranges
                If Not TypeOf TargetObject Is Excel.Range Then Exit Function
            'Use the explicit Excel range target
                Set TargetCell = TargetObject
    'Treat Empty as no explicit target supplied
        ElseIf VBA.IsEmpty(vTarget) Then
            If Not Try_GetActiveCell(TargetCell) Then Exit Function
    'Treat Null as no explicit target supplied
        ElseIf VBA.IsNull(vTarget) Then
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
    'Clear local object reference on failure
        Set TargetObject = Nothing
    'Return the safe default
        Resolve_TargetCell = False
    'Clear any suppressed resolver error
        Err.Clear

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
'   ActiveCell access can fail during no-window states, workbook teardown, chart
'   sheet activation, or transient activation boundaries
'
' INPUTS
'   TargetCell
'     Output Range receiving ActiveCell when available
'
' RETURNS
'   True if ActiveCell resolved to a valid single worksheet cell; otherwise False
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
'   ActiveCell is normally a single cell, but validation is still delegated to
'   Normalize_TargetCell so merged cells, invalid contexts, and unexpected
'   states are handled consistently
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Try_GetActiveCell"

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to unresolved
        Try_GetActiveCell = False
    'Clear output reference
        Set TargetCell = Nothing
    'Enable safe-default resolver handling
        On Error GoTo ErrorHandler

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
    'Return the safe default
        Try_GetActiveCell = False
    'Clear any suppressed resolver error
        Err.Clear

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
'   True when TargetCell is a valid DatePicker target cell; otherwise False
'
' BEHAVIOR
'   Accepts:
'     - normal single worksheet cells
'     - a single physical cell inside a merged range
'     - the full merged range supplied by Excel selection events
'
'   Rejects:
'     - missing targets
'     - multi-area ranges
'     - multi-cell non-merged ranges
'     - partial merged ranges that are neither one physical cell nor the full
'       MergeArea
'
'   Normalizes accepted merged-cell contexts to the top-left cell of the
'   MergeArea
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
'   MergeArea is resolved from the first physical cell because Excel's MergeArea
'   property is safest when applied to a single-cell range
'
'   This routine intentionally accepts a single selected cell inside a merged
'   area because ActiveCell commonly resolves that way
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Normalize_TargetCell"

    Dim FirstCell               As Excel.Range      'First physical cell in the supplied range
    Dim MergeArea               As Excel.Range      'Merged area containing the first cell
    Dim TargetCellCount         As Double           'Number of cells in supplied target
    Dim IsFullMergeArea         As Boolean          'True when supplied target equals full merge area

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to invalid
        Normalize_TargetCell = False
    'Enable safe-default resolver handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' VALIDATE PRESENCE
'------------------------------------------------------------------------------
    'Exit if no range was supplied
        If TargetCell Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' VALIDATE AREA COUNT
'------------------------------------------------------------------------------
    'Reject multi-area selections
        If TargetCell.Areas.Count <> 1 Then Exit Function

'------------------------------------------------------------------------------
' CAPTURE FIRST CELL
'------------------------------------------------------------------------------
    'Capture the first physical cell for safe merged-cell inspection
        Set FirstCell = TargetCell.Cells(1, 1)
    'Capture the target cell count once
        TargetCellCount = TargetCell.Cells.CountLarge

'------------------------------------------------------------------------------
' HANDLE MERGED CELLS
'------------------------------------------------------------------------------
    'Handle targets whose first physical cell belongs to a merged range
        If FirstCell.MergeCells Then
            'Resolve the merge area from the first physical cell
                Set MergeArea = FirstCell.MergeArea
            'Accept one physical cell inside the merged range
                If TargetCellCount = 1 Then
                    Set TargetCell = MergeArea.Cells(1, 1)
                    Normalize_TargetCell = True
                    Exit Function
                End If
            'Reject targets on a different worksheet
                If Not TargetCell.Worksheet Is MergeArea.Worksheet Then Exit Function
            'Test whether the supplied target is exactly the full merge area
                IsFullMergeArea = _
                    (TargetCell.Row = MergeArea.Row) And _
                    (TargetCell.Column = MergeArea.Column) And _
                    (TargetCell.Rows.CountLarge = MergeArea.Rows.CountLarge) And _
                    (TargetCell.Columns.CountLarge = MergeArea.Columns.CountLarge)
            'Reject partial merged ranges
                If Not IsFullMergeArea Then Exit Function
            'Normalize the full merged range to its top-left cell
                Set TargetCell = MergeArea.Cells(1, 1)
            'Signal successful target normalization
                Normalize_TargetCell = True
            'Exit after merged-cell normalization
                Exit Function
        End If

'------------------------------------------------------------------------------
' REJECT MULTI-CELL NON-MERGED SELECTIONS
'------------------------------------------------------------------------------
    'Reject multi-cell selections that are not a merged-cell context
        If TargetCellCount <> 1 Then Exit Function

'------------------------------------------------------------------------------
' NORMALIZE NORMAL SINGLE CELL
'------------------------------------------------------------------------------
    'Normalize the output reference to the single physical cell
        Set TargetCell = FirstCell

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
    'Clear local object references on failure
        Set FirstCell = Nothing
        Set MergeArea = Nothing
    'Return the safe default
        Normalize_TargetCell = False
    'Clear any suppressed resolver error
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                           USERFORM STATE HELPERS
'
'------------------------------------------------------------------------------
'

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
'     UserForm class name or runtime Name property to find
'
' RETURNS
'   True if a matching form is loaded; otherwise False
'
' BEHAVIOR
'   Scans VBA.UserForms and compares both:
'     - TypeName(CurForm)
'     - CurForm.Name
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
'   This routine deliberately avoids references such as UF_DatePicker.Visible,
'   because referencing a default form instance may load it
'
'   TypeName is preferred because the runtime Name property can be changed
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Is_UserFormLoaded"

    Dim CurForm                 As Object        'Loaded UserForm instance
    Dim NormalizedFormName      As String        'Trimmed form name to find
    Dim CurTypeName             As String        'Current form class name
    Dim CurRuntimeName          As String        'Current form runtime name

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not loaded
        Is_UserFormLoaded = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Normalize the requested form name once
        NormalizedFormName = VBA.Trim$(FormName)

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit for blank form names
        If VBA.Len(NormalizedFormName) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SCAN LOADED FORMS
'------------------------------------------------------------------------------
    'Loop through currently loaded UserForms
        For Each CurForm In VBA.UserForms
            'Reset current form metadata
                CurTypeName = vbNullString
            'Reset current runtime name
                CurRuntimeName = vbNullString
            'Suppress metadata-read errors for unusual transient form states
                On Error Resume Next
            'Capture the form class name
                CurTypeName = VBA.TypeName(CurForm)
            'Capture the form runtime name
                CurRuntimeName = VBA.CStr(CurForm.Name)
            'Clear any suppressed metadata error
                Err.Clear
            'Restore safe-default error handling
                On Error GoTo ErrorHandler
            'Return True when either the class name or runtime name matches
                If VBA.StrComp(CurTypeName, NormalizedFormName, vbTextCompare) = 0 Or _
                   VBA.StrComp(CurRuntimeName, NormalizedFormName, vbTextCompare) = 0 Then
                    Is_UserFormLoaded = True
                    Exit Function
                End If
        Next CurForm

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release local object reference
        Set CurForm = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear local object reference on failure
        Set CurForm = Nothing
    'Return the safe default
        Is_UserFormLoaded = False
    'Clear any suppressed predicate error
        Err.Clear

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
'     UserForm class name or runtime Name property to find
'
' RETURNS
'   True if any matching loaded form is visible; otherwise False
'
' BEHAVIOR
'   Scans VBA.UserForms and compares both:
'     - TypeName(CurForm)
'     - CurForm.Name
'
'   Returns True if at least one matching loaded instance is visible
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
'   This routine deliberately avoids references such as UF_DatePicker.Visible,
'   because referencing a default form instance may load it
'
'   If multiple matching instances are loaded, this routine returns True when
'   any matching instance is visible
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.Is_UserFormVisible"

    Dim CurForm                 As Object        'Loaded UserForm instance
    Dim NormalizedFormName      As String        'Trimmed form name to find
    Dim CurTypeName             As String        'Current form class name
    Dim CurRuntimeName          As String        'Current form runtime name
    Dim CurIsVisible            As Boolean       'Current form visibility state
    Dim IsMatchingForm          As Boolean       'True when current form matches requested form

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to not visible
        Is_UserFormVisible = False
    'Enable safe-default error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Normalize the requested form name once
        NormalizedFormName = VBA.Trim$(FormName)

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit for blank form names
        If VBA.Len(NormalizedFormName) = 0 Then Exit Function

'------------------------------------------------------------------------------
' SCAN LOADED FORMS
'------------------------------------------------------------------------------
    'Loop through currently loaded UserForms
        For Each CurForm In VBA.UserForms
            'Reset current form metadata
                CurTypeName = vbNullString
            'Reset current runtime name
                CurRuntimeName = vbNullString
            'Reset current visibility state
                CurIsVisible = False
            'Reset current match flag
                IsMatchingForm = False
            'Suppress metadata-read errors for unusual transient form states
                On Error Resume Next
            'Capture the form class name
                CurTypeName = VBA.TypeName(CurForm)
            'Capture the form runtime name
                CurRuntimeName = VBA.CStr(CurForm.Name)
            'Capture the form visibility state
                CurIsVisible = VBA.CBool(CurForm.Visible)
            'Clear any suppressed metadata error
                Err.Clear
            'Restore safe-default error handling
                On Error GoTo ErrorHandler
            'Test whether either the class name or runtime name matches
                IsMatchingForm = _
                    (VBA.StrComp(CurTypeName, NormalizedFormName, vbTextCompare) = 0) Or _
                    (VBA.StrComp(CurRuntimeName, NormalizedFormName, vbTextCompare) = 0)
            'Return True as soon as a visible matching instance is found
                If IsMatchingForm Then
                    If CurIsVisible Then
                        Is_UserFormVisible = True
                        Exit Function
                    End If
                End If
        Next CurForm

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Release local object reference
        Set CurForm = Nothing
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Clear local object reference on failure
        Set CurForm = Nothing
    'Return the safe default
        Is_UserFormVisible = False
    'Clear any suppressed predicate error
        Err.Clear

End Function

'
'------------------------------------------------------------------------------
'
'                                BUSY GUARD
'
'------------------------------------------------------------------------------

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
'   No local error handler is required because this routine only reads and writes
'   a private Boolean member
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default to failed guard entry
        Try_BeginBusy = False

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
'   No local error handler is required because this routine only writes a
'   private Boolean member
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' LEAVE BUSY STATE
'------------------------------------------------------------------------------
    'Clear the busy guard
        mIsBusy = False

End Sub

'
'------------------------------------------------------------------------------
'
'                               APPLICATION EVENTS
'
'------------------------------------------------------------------------------

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
'   Delegates to Handle_SelectionChange using the real Excel event Target
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_SelectionChange
'
' NOTES
'   Excel supplies Target as a Range
'
'   This event only fires when Application.EnableEvents is True
'
'   Sh is retained in the signature because Excel supplies it, but the manager
'   delegates the target-resolution policy to Handle_SelectionChange
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_SheetSelectionChange"

    Dim SheetName               As String       'Diagnostic sheet name
    Dim TargetAddress           As String       'Diagnostic target address
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the sheet name when available
        If Not Sh Is Nothing Then SheetName = VBA.CStr(Sh.Name)
    'Capture the target address when available
        If Not Target Is Nothing Then TargetAddress = Target.Address(False, False)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Delegate selection-change handling
        Handle_SelectionChange Target

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Sheet=" & SheetName & _
            " | Target=" & TargetAddress & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Edits can invalidate DatePicker relevance, icon placement, or the selected
'   date state of an already open picker
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
'   Delegates to Handle_WorksheetChange using the real Excel event Target
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_WorksheetChange
'
' NOTES
'   DatePicker write-back suppresses worksheet events, so this routine mainly
'   handles user edits and non-DatePicker changes
'
'   This event only fires when Application.EnableEvents is True
'
'   Sh is retained in the signature because Excel supplies it, but the manager
'   delegates the target-resolution policy to Handle_WorksheetChange
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_SheetChange"
    
    Dim SheetName               As String       'Diagnostic sheet name
    Dim TargetAddress           As String       'Diagnostic target address
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the sheet name when available
        If Not Sh Is Nothing Then SheetName = VBA.CStr(Sh.Name)
    'Capture the target address when available
        If Not Target Is Nothing Then TargetAddress = Target.Address(False, False)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Delegate worksheet-change handling
        Handle_WorksheetChange Target

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Sheet=" & SheetName & _
            " | Target=" & TargetAddress & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The workbook argument is not used for routing, but its name is captured for
'   diagnostics when available
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_NewWorkbook"

    Dim WorkbookName            As String       'Diagnostic workbook name
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the new workbook context
        Handle_WorkbookContextChange

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Clears transient UI when the manager is idle, closes the picker, and purges
'   grid icons before Excel proceeds with printing
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Printing is a hard cleanup boundary, so the picker is closed and grid icons
'   are purged
'
'   This routine does not modify Cancel. Any existing print-cancel decision made
'   by Excel or another handler is preserved
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WorkbookBeforePrint"

    Dim WorkbookName            As String       'Diagnostic workbook name
    Dim CancelWasSet            As Boolean      'Diagnostic print-cancel state
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next

    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture the current cancel state without modifying it
        CancelWasSet = Cancel
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before print using the hard-purge path
        Clear_UI_IfIdle True, True

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | CancelWasSet=" & VBA.CStr(CancelWasSet) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Clears transient UI when the manager is idle, closes the picker, and purges
'   grid icons before Excel proceeds with saving
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Saving is a hard cleanup boundary, so the picker is closed and grid icons are
'   purged
'
'   This routine does not modify Cancel. Any existing save-cancel decision made
'   by Excel or another handler is preserved
'
'   If Save As UI is later cancelled by the user, the transient UI has still
'   been cleared. This is acceptable because DatePicker UI artifacts should not
'   be treated as workbook content
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WorkbookBeforeSave"

    Dim WorkbookName            As String           'Diagnostic workbook name
    Dim SaveAsUIWasShown        As Boolean          'Diagnostic Save As UI state
    Dim CancelWasSet            As Boolean          'Diagnostic save-cancel state
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture whether Save As UI is being shown
        SaveAsUIWasShown = SaveAsUI
    'Capture the current cancel state without modifying it
        CancelWasSet = Cancel
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before save using the hard-purge path
        Clear_UI_IfIdle True, True

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | SaveAsUI=" & VBA.CStr(SaveAsUIWasShown) & _
            " | CancelWasSet=" & VBA.CStr(CancelWasSet) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Clears transient UI when the manager is idle and the close operation has not
'   already been cancelled
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Closing is a hard cleanup boundary, so the picker is closed and grid icons
'   are purged
'
'   This routine does not modify Cancel. Any existing close-cancel decision made
'   by Excel or another handler is preserved
'
'   If close has already been cancelled when this handler runs, cleanup is
'   skipped because the workbook is remaining open
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WorkbookBeforeClose"

    Dim WorkbookName            As String       'Diagnostic workbook name
    Dim CancelWasSet            As Boolean      'Diagnostic close-cancel state
    Dim ErrorNumber             As Long         'Captured error number
    Dim ErrorDescription        As String       'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture the current cancel state without modifying it
        CancelWasSet = Cancel
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' RESPECT PRE-EXISTING CLOSE CANCELLATION
'------------------------------------------------------------------------------
    'Exit without cleanup when close has already been cancelled
        If CancelWasSet Then Exit Sub

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI before close using the hard-purge path
        Clear_UI_IfIdle True, True

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | CancelWasSet=" & VBA.CStr(CancelWasSet) & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Clears grid artifacts only when the manager is idle and keeps any loaded
'   modeless DatePicker form open
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Resize uses the lightweight cleanup path for responsiveness
'
'   This routine intentionally does not call Handle_WorkbookContextChange because
'   that path may close the picker depending on the close-after-selection policy
'
'   If immediate icon repositioning after resize is required, add a dedicated
'   grid-icon-only refresh path rather than using the full picker refresh path
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WindowResize"

    Dim WorkbookName            As String        'Diagnostic workbook name
    Dim WindowCaption           As String        'Diagnostic window caption
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture the window caption when available
        If Not Wn Is Nothing Then WindowCaption = VBA.CStr(Wn.Caption)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CLEAR GRID ARTIFACTS ONLY
'------------------------------------------------------------------------------
    'Clear transient grid UI without closing the modeless picker
        Clear_UI_IfIdle False, False

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Window=" & WindowCaption & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

End Sub
'


Private Sub mExcelApp_WindowActivate( _
    ByVal Wb As Excel.Workbook, _
    ByVal Wn As Excel.Window)

'
'------------------------------------------------------------------------------
'                           APP WINDOW ACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Refreshes DatePicker UI when an Excel window is activated
'
' WHY THIS EXISTS
'   Window deactivation may remove transient grid artifacts without closing the
'   modeless picker. When the Excel window becomes active again, the manager
'   should re-evaluate the current active Excel context, resolve it into a
'   TargetCell when possible, and restore the in-grid icon when appropriate
'
' INPUTS
'   Wb
'     Workbook owning the activated window
'
'   Wn
'     Activated Excel window
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to Handle_WorkbookContextChange, which resolves the current active
'   context into an authoritative TargetCell when possible
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   This handler complements mExcelApp_WindowDeactivate
'
'   The workbook and window arguments are used only for diagnostic context
'
'   This event only fires when Application.EnableEvents is True and the manager
'   instance already exists
'
' UPDATED
'   2026-05-04
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WindowActivate"

    Dim WorkbookName            As String        'Diagnostic workbook name
    Dim WindowCaption           As String        'Diagnostic window caption
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture the window caption when available
        If Not Wn Is Nothing Then WindowCaption = VBA.CStr(Wn.Caption)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REFRESH ACTIVE CONTEXT
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the newly activated window context
        Handle_WorkbookContextChange

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Window=" & WindowCaption & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Clears grid artifacts only when the manager is idle and keeps any loaded
'   modeless DatePicker form alive
'
' ERROR POLICY
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   WorkbookBeforeClose remains the hard cleanup boundary
'
'   This routine intentionally does not call Handle_WorkbookContextChange because
'   that path may close the picker depending on the close-after-selection policy
'
'   If immediate icon restoration is required when the Excel window is activated
'   again, use the matching activation/context-change handler to resolve the
'   current active context into a TargetCell when possible
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WindowDeactivate"

    Dim WorkbookName            As String           'Diagnostic workbook name
    Dim WindowCaption           As String           'Diagnostic window caption
    Dim ErrorNumber             As Long             'Captured error number
    Dim ErrorDescription        As String           'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook name when available
        If Not Wb Is Nothing Then WorkbookName = VBA.CStr(Wb.Name)
    'Capture the window caption when available
        If Not Wn Is Nothing Then WindowCaption = VBA.CStr(Wn.Caption)
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CLEAR GRID ARTIFACTS ONLY
'------------------------------------------------------------------------------
    'Clear transient grid UI without closing the modeless picker
        Clear_UI_IfIdle False, False

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Window=" & WindowCaption & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

End Sub

Private Sub mExcelApp_SheetActivate(ByVal Sh As Object)

'
'------------------------------------------------------------------------------
'                           APP SHEET ACTIVATE
'------------------------------------------------------------------------------
' PURPOSE
'   Handles worksheet or chart-sheet activation
'
' WHY THIS EXISTS
'   Activating a sheet changes the active Excel context and may require a fresh
'   DatePicker UI decision for the newly active cell, sheet, and workbook
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
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The sheet argument is used only for diagnostic context
'
'   The refresh uses the current active Excel context only as an input source.
'   Refresh_CurrentUI then normalizes that context into a resolved TargetCell
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_SheetActivate"

    Dim WorkbookName            As String        'Diagnostic workbook name
    Dim SheetName               As String        'Diagnostic sheet name
    Dim SheetType               As String        'Diagnostic sheet object type
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the sheet context when available
        If Not Sh Is Nothing Then
            SheetName = VBA.CStr(Sh.Name)
            SheetType = VBA.TypeName(Sh)
            WorkbookName = VBA.CStr(Sh.Parent.Name)
        End If
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REFRESH ACTIVE CONTEXT
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the newly activated sheet context
        Handle_WorkbookContextChange

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | Sheet=" & SheetName & _
            " | SheetType=" & SheetType & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Activating a workbook changes the active Excel context and may require a
'   fresh DatePicker UI decision for the active sheet and active cell
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
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Handle_WorkbookContextChange
'
' NOTES
'   The workbook argument is used only for diagnostic context
'
'   The refresh uses the current active Excel context only as an input source.
'   Refresh_CurrentUI then normalizes that context into a resolved TargetCell
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WorkbookActivate"

    Dim WorkbookName            As String        'Diagnostic workbook name
    Dim WorkbookFullName        As String        'Diagnostic workbook full name
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook context when available
        If Not Wb Is Nothing Then
            WorkbookName = VBA.CStr(Wb.Name)
            WorkbookFullName = VBA.CStr(Wb.FullName)
        End If
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' REFRESH ACTIVE CONTEXT
'------------------------------------------------------------------------------
    'Re-evaluate DatePicker UI under the newly activated workbook context
        Handle_WorkbookContextChange

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | FullName=" & WorkbookFullName & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

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
'   Leaving a workbook is a natural boundary for removing transient DatePicker
'   UI artifacts from a no-longer-active workbook context
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
'   Best-effort. Does not raise into Excel event flow
'
' DEPENDENCIES
'   Clear_UI_IfIdle
'
' NOTES
'   Workbook deactivation is treated as a cleanup boundary
'
'   This routine keeps the existing hard-purge behavior by requesting full
'   picker close and global grid-icon cleanup through Clear_UI_IfIdle
'
'   This event only fires when Application.EnableEvents is True
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "cDatePickerManager.mExcelApp_WorkbookDeactivate"

    Dim WorkbookName            As String        'Diagnostic workbook name
    Dim WorkbookFullName        As String        'Diagnostic workbook full name
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Protect Excel from event-handler failures
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' SNAPSHOT DIAGNOSTIC CONTEXT
'------------------------------------------------------------------------------
    'Suppress diagnostic snapshot failures
        On Error Resume Next
    'Capture the workbook context when available
        If Not Wb Is Nothing Then
            WorkbookName = VBA.CStr(Wb.Name)
            WorkbookFullName = VBA.CStr(Wb.FullName)
        End If
    'Clear any suppressed diagnostic error
        Err.Clear
    'Restore event-handler error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' CLEAR UI
'------------------------------------------------------------------------------
    'Clear DatePicker UI on workbook deactivation using the hard-purge path
        Clear_UI_IfIdle True, True

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
    'Write event diagnostics without interrupting Excel event flow
        Debug.Print PROC_NAME & _
            " | Workbook=" & WorkbookName & _
            " | FullName=" & WorkbookFullName & _
            " | Error=" & VBA.CStr(ErrorNumber) & _
            " | " & ErrorDescription
    'Clear the suppressed event error
        Err.Clear

End Sub



