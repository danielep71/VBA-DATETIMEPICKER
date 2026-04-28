Attribute VB_Name = "M_cPM_TimeWasters"
'==============================================================================
' MODULE: M_cPM_TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Shared, process-wide manager for Excel "time-waster" suppression used by
'   cPerformanceManager
'
' WHY THIS EXISTS
'   The Excel Application properties controlled here are global process state,
'   not instance-local state:
'
'     - ScreenUpdating
'     - EnableEvents
'     - DisplayAlerts
'     - Calculation
'     - Cursor
'
'   Therefore:
'     - multiple cPerformanceManager instances can overlap
'     - each instance can request a different exemption mask
'     - restore logic must be coordinated globally
'     - the final restore must happen exactly once when the last active session
'       ends
'
' REQUIRED BY
'   cPerformanceManager
'
' COMPILE-TIME CONTRACT
'   This module is required by cPerformanceManager because the class directly
'   references the following procedures exposed here:
'
'     - PM_TW_BeginSession
'     - PM_TW_EndSession
'     - PM_TW_ActiveCount
'
'   Additional diagnostic / recovery procedures exposed here are:
'
'     - PM_TW_IsInstanceActive
'     - PM_TW_EndAllSessions
'
' IMPORT REQUIREMENT
'   Import this module together with cPerformanceManager.cls
'
' DESIGN
'   - The first active session captures the original Application baseline
'   - Each active instance registers its own disable-mask
'   - The effective disable-mask is the OR of all active instance masks
'   - Whenever a session begins, updates, or ends, the effective state is
'     recomputed
'   - When the final session ends, the original baseline is restored exactly
'     once and the shared store is released
'
' DEPENDENCIES
'   - Excel Application object model
'   - Late-bound Scripting.Dictionary via CreateObject("Scripting.Dictionary")
'
' REFERENCE POLICY
'   No manual reference to "Microsoft Scripting Runtime" is required
'
' ERROR POLICY
'   - Public begin/end operations raise errors normally for invalid calling
'     flow, invalid inputs, or Application-state failures
'   - Idempotent no-op paths remain intentional, for example when ending an
'     already-idle shared store or an inactive instance key
'   - Internal helpers assume a valid Excel host and valid calling flow
'   - PM_TW_EndAllSessions is an emergency / recovery helper and also raises
'     errors normally
'
' NOTES
'   - This module should not appear as a user-runnable macro surface.
'     Therefore Option Private Module is used
'   - Cursor suppression uses xlWait while active to force a deterministic
'     benchmark-time cursor state and avoid ordinary cursor-state churn during
'     benchmark runs
'
' VERSION
'   1.0.0
'
' UPDATED
'   2026-04-18
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit         'Force explicit declaration of all variables
    Option Private Module   'Hide internal support procedures from Macro dialog

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    'Known TW mask bits
        Private Const PM_TW_MASK_NONE               As Long = 0
        Private Const PM_TW_MASK_SCREENUPDATING     As Long = 1
        Private Const PM_TW_MASK_ENABLEEVENTS       As Long = 2
        Private Const PM_TW_MASK_DISPLAYALERTS      As Long = 4
        Private Const PM_TW_MASK_CALCULATION        As Long = 8
        Private Const PM_TW_MASK_CURSOR             As Long = 16
        Private Const PM_TW_MASK_ALL                As Long = 31

'------------------------------------------------------------------------------
' PRIVATE SHARED STATE
'------------------------------------------------------------------------------
    'Dictionary:
    '   key   = instance key (String)
    '   item  = disable-mask (Long)
        Private g_TW_Sessions               As Object

    'TRUE once the baseline Application state has been captured
        Private g_TW_BaselineSaved          As Boolean

    'Saved baseline Application state
        Private g_TW_SCREENUPDATING         As Boolean
        Private g_TW_ENABLEEVENTS           As Boolean
        Private g_TW_DISPLAYALERTS          As Boolean
        Private g_TW_CALCULATION            As XlCalculation
        Private g_TW_CURSOR                 As XlMousePointer

'
'==============================================================================
'
'                                  PUBLIC API
'
'==============================================================================

Public Sub PM_TW_BeginSession( _
    ByVal InstanceKey As String, _
    Optional ByVal ExceptMask As Long = 0)
'
'==============================================================================
'                             PM_TW_BEGINSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Starts or updates a shared TW suppression session for one class instance
'
' WHY THIS EXISTS
'   The calling class cannot safely manage Excel Application TW state
'   independently because that state is global across the Excel process.
'   Therefore the class registers its request here and this manager computes the
'   aggregate effective state across all active instances
'
' INPUTS
'   InstanceKey
'     Unique key identifying the calling class instance
'
'   ExceptMask (optional)
'     Bitmask of TW flags to EXEMPT
'     Any flag present in ExceptMask remains at the original baseline state for
'     this instance's request
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Ensures the shared session store exists
'   - Captures the original Application baseline when the first active session
'     begins
'   - Converts the supplied exemption mask into a disable-mask
'   - Registers or updates this instance in the shared session store
'   - Recomputes and applies the aggregate effective state
'   - Rolls back the registration/update if effective-state application fails
'   - Best-effort reapplies the previous effective state after rollback
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_EnsureStore
'   - PM_TW_SaveBaseline
'   - PM_TW_DisableMaskFromExcept
'   - PM_TW_AggregateDisableMask
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_ResetSharedState
'
' NOTES
'   This routine is instance-idempotent:
'     - first call for an instance => begin/register
'     - later calls for same instance => update requested mask
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HadKeyBefore            As Boolean   'TRUE when the instance was already registered
    Dim PrevDisableMask         As Long      'Previously stored disable-mask for this instance
    Dim WasFirstSession         As Boolean   'TRUE when this call began the first shared session

    Dim SavedErrNumber          As Long      'Captured original error number
    Dim SavedErrSource          As String    'Captured original error source
    Dim SavedErrDescription     As String    'Captured original error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise vbObjectError + 2200, _
                      "M_cPM_TimeWasters.PM_TW_BeginSession", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists
        PM_TW_EnsureStore
    'Capture whether the key already exists before modification
        HadKeyBefore = g_TW_Sessions.Exists(InstanceKey)
    'Capture whether this call is opening the first shared session
        WasFirstSession = (g_TW_Sessions.Count = 0)
    'Capture the previous disable-mask when this is an update
        If HadKeyBefore Then
            PrevDisableMask = CLng(g_TW_Sessions(InstanceKey))
        End If

'------------------------------------------------------------------------------
' CAPTURE BASELINE (FIRST ACTIVE SESSION ONLY)
'------------------------------------------------------------------------------
    'Capture the original Application state only when the first shared session begins
        If WasFirstSession Then
            PM_TW_SaveBaseline
        End If

'------------------------------------------------------------------------------
' REGISTER / UPDATE INSTANCE MASK
'------------------------------------------------------------------------------
    'Store this instance's requested disable-mask
        g_TW_Sessions(InstanceKey) = PM_TW_DisableMaskFromExcept(ExceptMask)

'------------------------------------------------------------------------------
' APPLY EFFECTIVE SHARED STATE
'------------------------------------------------------------------------------
    'Recompute the aggregate disable-mask and apply it
        On Error GoTo ApplyFail
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
        On Error GoTo 0

    Exit Sub

ApplyFail:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before rollback
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

'------------------------------------------------------------------------------
' ROLLBACK REGISTRATION / UPDATE
'------------------------------------------------------------------------------
    'Rollback must be best-effort so that we can preserve the original error
        On Error Resume Next

    'Restore the prior registration state
        If HadKeyBefore Then
            g_TW_Sessions(InstanceKey) = PrevDisableMask
        ElseIf g_TW_Sessions.Exists(InstanceKey) Then
            g_TW_Sessions.Remove InstanceKey
        End If

'------------------------------------------------------------------------------
' BEST-EFFORT REAPPLY PREVIOUS EFFECTIVE STATE
'------------------------------------------------------------------------------
    'If rollback returned the manager to a true idle state, try to restore the
    'baseline and then clear shared state
        If g_TW_Sessions.Count = 0 Then
            If g_TW_BaselineSaved Then
                Err.Clear
                PM_TW_ApplyEffectiveState PM_TW_MASK_NONE

                If Err.Number = 0 Then
                    PM_TW_ResetSharedState
                End If
            Else
                PM_TW_ResetSharedState
            End If
    'Otherwise best-effort reapply the prior aggregate effective state
        Else
            Err.Clear
            PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
        End If

        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Sub
Public Sub PM_TW_EndSession( _
    ByVal InstanceKey As String)
'
'==============================================================================
'                              PM_TW_ENDSESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Ends a shared TW suppression session for one class instance
'
' WHY THIS EXISTS
'   TW suppression must be removed by deregistering the calling instance from
'   the shared manager, not by blindly restoring Excel state locally. This keeps
'   overlapping instances safe and ensures the effective state is recomputed
'   correctly from the remaining active sessions
'
' INPUTS
'   InstanceKey
'     Unique key identifying the calling class instance
'
' RETURNS
'   None
'
' BEHAVIOR
'   - If no shared store exists, exits immediately
'   - Removes the specified instance if present
'   - If no sessions remain:
'       * restores the original baseline exactly once when available
'       * clears the baseline-saved flag
'       * releases the shared dictionary
'   - Otherwise recomputes and reapplies the remaining effective state
'   - Rolls back the removal if effective-state application fails
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_AggregateDisableMask
'   - PM_TW_ResetSharedState
'
' NOTES
'   This routine is idempotent for unknown / inactive instance keys:
'     - if the instance is not present, no removal occurs
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim HadKeyBefore            As Boolean   'TRUE when the instance was registered before removal
    Dim PrevDisableMask         As Long      'Previously stored disable-mask for this instance

    Dim SavedErrNumber          As Long      'Captured original error number
    Dim SavedErrSource          As String    'Captured original error source
    Dim SavedErrDescription     As String    'Captured original error description

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise vbObjectError + 2201, _
                      "M_cPM_TimeWasters.PM_TW_EndSession", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'If no shared store exists yet, there is nothing to end
        If g_TW_Sessions Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CAPTURE PRE-REMOVAL STATE
'------------------------------------------------------------------------------
    'Capture whether the instance is currently registered
        HadKeyBefore = g_TW_Sessions.Exists(InstanceKey)

    'Capture the previous disable-mask when present
        If HadKeyBefore Then
            PrevDisableMask = CLng(g_TW_Sessions(InstanceKey))
        End If

'------------------------------------------------------------------------------
' REMOVE INSTANCE (IF PRESENT)
'------------------------------------------------------------------------------
    'Remove the calling instance from the active session set
        If HadKeyBefore Then
            g_TW_Sessions.Remove InstanceKey
        End If

'------------------------------------------------------------------------------
' RESTORE OR REAPPLY
'------------------------------------------------------------------------------
        On Error GoTo ApplyFail

    'If no sessions remain, restore the original baseline
        If g_TW_Sessions.Count = 0 Then
            'Restore the original Application state only if a baseline was actually captured
                If g_TW_BaselineSaved Then
                    PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
                End If

            'Return to a clean idle shared-state baseline
                PM_TW_ResetSharedState

            On Error GoTo 0
            Exit Sub
        End If

    'Otherwise recompute and apply the remaining aggregate disable-mask
        PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
        On Error GoTo 0

    Exit Sub

ApplyFail:
'------------------------------------------------------------------------------
' CAPTURE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Capture the original error before rollback
        SavedErrNumber = Err.Number
        SavedErrSource = Err.Source
        SavedErrDescription = Err.Description

'------------------------------------------------------------------------------
' ROLLBACK REMOVAL
'------------------------------------------------------------------------------
    'Rollback must be best-effort so that we can preserve the original error
        On Error Resume Next
    'Restore the removed instance registration when it existed before the call
        If HadKeyBefore Then
            g_TW_Sessions(InstanceKey) = PrevDisableMask
        End If

'------------------------------------------------------------------------------
' BEST-EFFORT REAPPLY PREVIOUS EFFECTIVE STATE
'------------------------------------------------------------------------------
    'Best-effort reapply the prior aggregate effective state
        If Not g_TW_Sessions Is Nothing Then
            If g_TW_Sessions.Count > 0 Then
                Err.Clear
                PM_TW_ApplyEffectiveState PM_TW_AggregateDisableMask()
            End If
        End If

        On Error GoTo 0

'------------------------------------------------------------------------------
' RE-RAISE ORIGINAL ERROR
'------------------------------------------------------------------------------
    'Re-raise the original error
        Err.Raise SavedErrNumber, SavedErrSource, SavedErrDescription

End Sub

Public Function PM_TW_ActiveCount() As Long
'
'==============================================================================
'                             PM_TW_ACTIVECOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of currently active shared TW sessions
'
' WHY THIS EXISTS
'   Useful for diagnostics, assertions, and visibility when verifying shared TW
'   lifecycle behavior across multiple cPerformanceManager instances
'
' INPUTS
'   None.
'
' RETURNS
'   Long
'     Number of currently active shared TW sessions
'
' BEHAVIOR
'   - Returns 0 when no shared store currently exists
'   - Otherwise returns the dictionary count
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   - g_TW_Sessions
'
' NOTES
'   This routine does not create the shared store on an idle read path
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'Return 0 when no shared store currently exists
        If g_TW_Sessions Is Nothing Then
            PM_TW_ActiveCount = 0
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the active session count
        PM_TW_ActiveCount = g_TW_Sessions.Count

End Function

Public Function PM_TW_IsInstanceActive( _
    ByVal InstanceKey As String) _
    As Boolean
'
'==============================================================================
'                          PM_TW_ISINSTANCEACTIVE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if the specified class instance currently has an active shared
'   TW session registered in the global manager
'
' WHY THIS EXISTS
'   Useful for diagnostics, regression testing, and troubleshooting of shared
'   TW registration / update / end behavior
'
' INPUTS
'   InstanceKey
'     Unique key identifying the class instance to inspect
'
' RETURNS
'   Boolean
'     TRUE  => the instance is currently registered
'     FALSE => the instance is not currently registered
'
' BEHAVIOR
'   - Returns FALSE when no shared store currently exists
'   - Otherwise queries the dictionary for the supplied key
'
' ERROR POLICY
'   Raises on a blank instance key
'
' DEPENDENCIES
'   - g_TW_Sessions
'
' NOTES
'   This routine does not create the shared store on an idle read path
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a blank instance key
        If Len(Trim$(InstanceKey)) = 0 Then
            Err.Raise vbObjectError + 2202, _
                      "M_cPM_TimeWasters.PM_TW_IsInstanceActive", _
                      "InstanceKey cannot be blank."
        End If

'------------------------------------------------------------------------------
' VALIDATE STORE
'------------------------------------------------------------------------------
    'Return FALSE when no shared store currently exists
        If g_TW_Sessions Is Nothing Then
            PM_TW_IsInstanceActive = False
            Exit Function
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return instance activity state
        PM_TW_IsInstanceActive = g_TW_Sessions.Exists(InstanceKey)

End Function

'
'==============================================================================
'
'                                PRIVATE HELPERS
'
'==============================================================================

Private Sub PM_TW_EnsureStore()
'
'==============================================================================
'                             PM_TW_ENSURESTORE
'------------------------------------------------------------------------------
' PURPOSE
'   Lazily creates the shared session dictionary
'
' WHY THIS EXISTS
'   The shared TW store should exist only when needed. This helper centralizes
'   the lazy-creation logic and keeps begin/update paths simple
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Creates the shared dictionary only when it does not already exist
'   - Uses late binding so no external reference is required
'   - Sets binary key comparison explicitly
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - CreateObject("Scripting.Dictionary")
'
' NOTES
'   Read-only status helpers intentionally avoid calling this routine so that an
'   idle project remains in a true idle state
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Create the shared dictionary only once
        If g_TW_Sessions Is Nothing Then
            Set g_TW_Sessions = CreateObject("Scripting.Dictionary")
            g_TW_Sessions.CompareMode = vbBinaryCompare
        End If

End Sub

Private Sub PM_TW_SaveBaseline()
'
'==============================================================================
'                             PM_TW_SAVEBASELINE
'------------------------------------------------------------------------------
' PURPOSE
'   Captures the original Application state before any TW suppression is
'   applied
'
' WHY THIS EXISTS
'   Shared TW suppression must restore the exact original Excel baseline when
'   the final active session ends. Therefore that baseline must be captured once
'   at the beginning of the first shared session
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the current Application state into shared baseline variables
'   - Marks the baseline as captured
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Excel Application object model
'
' NOTES
'   This routine should only be called when the first shared session begins
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' CAPTURE BASELINE
'------------------------------------------------------------------------------
    'Read the current Application state into shared baseline variables
        With Application
            g_TW_SCREENUPDATING = .ScreenUpdating
            g_TW_ENABLEEVENTS = .EnableEvents
            g_TW_DISPLAYALERTS = .DisplayAlerts
            g_TW_CALCULATION = .Calculation
            g_TW_CURSOR = .Cursor
        End With

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Mark the baseline as captured
        g_TW_BaselineSaved = True

End Sub

Private Sub PM_TW_ResetSharedState()
'
'==============================================================================
'                           PM_TW_RESETSHAREDSTATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the shared TW manager to a clean idle state
'
' WHY THIS EXISTS
'   Several paths need to clear the same shared-state variables:
'     - successful end of the final active session
'     - explicit emergency reset
'     - rollback after a failed first-session begin
'
'   Centralizing that reset logic avoids partial cleanup drift
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Releases the shared session dictionary
'   - Clears the baseline-saved flag
'   - Clears the cached baseline values
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR SHARED STORE
'------------------------------------------------------------------------------
    'Release the shared session dictionary
        Set g_TW_Sessions = Nothing

'------------------------------------------------------------------------------
' CLEAR BASELINE FLAGS / VALUES
'------------------------------------------------------------------------------
    'Mark that no baseline is currently cached
        g_TW_BaselineSaved = False

    'Clear cached baseline values deterministically
        g_TW_SCREENUPDATING = False
        g_TW_ENABLEEVENTS = False
        g_TW_DISPLAYALERTS = False
        g_TW_CALCULATION = xlCalculationAutomatic
        g_TW_CURSOR = xlDefault

End Sub

Private Function PM_TW_DisableMaskFromExcept( _
    ByVal ExceptMask As Long) _
    As Long
'
'==============================================================================
'                        PM_TW_DISABLEMASKFROMEXCEPT
'------------------------------------------------------------------------------
' PURPOSE
'   Converts an exemption mask into a disable-mask
'
' WHY THIS EXISTS
'   Callers express their request as "leave these flags alone." Internally the
'   shared manager works more naturally with a disable-mask, because aggregate
'   shared state is computed by OR-ing together every active instance's disabled
'   flags
'
' INPUTS
'   ExceptMask
'     Bitmask of TW flags to exempt
'
' RETURNS
'   Long
'     Disable-mask representing the supported flags that should be forced into
'     their benchmark/performance state
'
' BEHAVIOR
'   - Clamps the input to known bits
'   - Inverts those bits within the supported TW mask universe
'
' ERROR POLICY
'   Does not raise errors
'
' DEPENDENCIES
'   - PM_TW_MASK_ALL
'
' EXAMPLE
'   ExceptMask = ScreenUpdating Or EnableEvents
'   => disable all supported TW flags except those two
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Clamp to known bits and invert within the supported TW mask set
        PM_TW_DisableMaskFromExcept = (PM_TW_MASK_ALL And Not ExceptMask)

End Function

Private Function PM_TW_AggregateDisableMask() As Long
'
'==============================================================================
'                        PM_TW_AGGREGATEDISABLEMASK
'------------------------------------------------------------------------------
' PURPOSE
'   ORs together the disable-masks of all active shared TW sessions
'
' WHY THIS EXISTS
'   The effective TW state is the union of every currently active instance's
'   requested disable-mask. This helper centralizes that aggregation
'
' INPUTS
'   None.
'
' RETURNS
'   Long
'     Aggregate disable-mask across all active sessions
'
' BEHAVIOR
'   - Ensures the shared store exists
'   - ORs together every active instance's stored disable-mask
'   - Returns the final aggregate mask
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_EnsureStore
'   - g_TW_Sessions
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K    As Variant    'Dictionary key
    Dim Mask As Long       'Accumulated disable-mask

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Ensure the shared dictionary exists
        PM_TW_EnsureStore

'------------------------------------------------------------------------------
' AGGREGATE
'------------------------------------------------------------------------------
    'OR together every active instance's disable-mask
        For Each K In g_TW_Sessions.Keys
            Mask = (Mask Or CLng(g_TW_Sessions(K)))
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the aggregate disable-mask
        PM_TW_AggregateDisableMask = Mask

End Function

Private Sub PM_TW_ApplyEffectiveState( _
    ByVal DisableMask As Long)
'
'==============================================================================
'                        PM_TW_APPLYEFFECTIVESTATE
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the effective shared TW state using:
'     - the original saved baseline, and
'     - the aggregate disable-mask of all active sessions
'
' WHY THIS EXISTS
'   Shared TW control should never restore or force individual flags in an
'   ad-hoc per-instance way. Instead, each flag must be derived from:
'
'     - the original baseline captured at the first session, and
'     - whether any currently active instance wants that flag disabled
'
' INPUTS
'   DisableMask
'     Aggregate disable-mask to apply
'
' RETURNS
'   None
'
' BEHAVIOR
'   For each supported flag:
'     - if disabled by any active session => force benchmark/performance state
'     - otherwise => restore original baseline state
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Excel Application object model
'   - g_TW_SCREENUPDATING
'   - g_TW_ENABLEEVENTS
'   - g_TW_DISPLAYALERTS
'   - g_TW_CALCULATION
'   - g_TW_CURSOR
'
' NOTES
'   - This routine assumes the baseline Application state has already been
'     captured when restoration semantics are required
'   - Cursor uses xlWait when disabled to force a deterministic benchmark-time
'     cursor state
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' APPLY EFFECTIVE STATE
'------------------------------------------------------------------------------
    'Apply the effective state flag-by-flag
        With Application
            'ScreenUpdating
                If (DisableMask And PM_TW_MASK_SCREENUPDATING) <> 0 Then
                    .ScreenUpdating = False
                Else
                    .ScreenUpdating = g_TW_SCREENUPDATING
                End If
            'EnableEvents
                If (DisableMask And PM_TW_MASK_ENABLEEVENTS) <> 0 Then
                    .EnableEvents = False
                Else
                    .EnableEvents = g_TW_ENABLEEVENTS
                End If
            'DisplayAlerts
                If (DisableMask And PM_TW_MASK_DISPLAYALERTS) <> 0 Then
                    .DisplayAlerts = False
                Else
                    .DisplayAlerts = g_TW_DISPLAYALERTS
                End If
            'Calculation
                If (DisableMask And PM_TW_MASK_CALCULATION) <> 0 Then
                    .Calculation = xlCalculationManual
                Else
                    .Calculation = g_TW_CALCULATION
                End If
            'Cursor
                If (DisableMask And PM_TW_MASK_CURSOR) <> 0 Then
                    .Cursor = xlWait
                Else
                    .Cursor = g_TW_CURSOR
                End If
        End With

End Sub

Public Sub PM_TW_EndAllSessions()
'
'==============================================================================
'                            PM_TW_ENDALLSESSIONS
'------------------------------------------------------------------------------
' PURPOSE
'   Emergency / global reset for development, recovery, or test-cleanup
'   scenarios
'
' WHY THIS EXISTS
'   In normal operation each instance should end only its own session through
'   PM_TW_EndSession. However, during development, test teardown, or recovery
'   from interrupted flows it can be useful to force a full shared reset
'
' INPUTS
'   None.
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Restores the original Application baseline when available
'   - Clears the baseline-saved flag
'   - Releases all shared session bookkeeping
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - PM_TW_ApplyEffectiveState
'   - PM_TW_ResetSharedState
'
' NOTES
'   This is not the normal lifecycle path.
'   Normal callers should use PM_TW_EndSession for the specific active instance
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' RESTORE BASELINE (IF AVAILABLE)
'------------------------------------------------------------------------------
    'Restore the original Application baseline if available
        If g_TW_BaselineSaved Then
            PM_TW_ApplyEffectiveState PM_TW_MASK_NONE
        End If

'------------------------------------------------------------------------------
' CLEAR SHARED STATE
'------------------------------------------------------------------------------
    'Release all active-session bookkeeping and baseline state
        PM_TW_ResetSharedState

End Sub

'
'------------------------------------------------------------------------------
'
'                       HELPERS FOR WORKSHEET OUTPUT
'
'------------------------------------------------------------------------------
'

Public Sub cPM_Report_WriteToRange( _
    ByVal cPM As cPerformanceManager, _
    ByVal TargetTopLeft As Range, _
    Optional ByVal ClearOutputArea As Boolean = False)
'
'==============================================================================
'                         REPORT WRITE TO RANGE
'------------------------------------------------------------------------------
' PURPOSE
'   Writes the structured checkpoint report to a worksheet range
'
' WHY THIS EXISTS
'   The class should own timing and structured capture, while worksheet output
'   belongs more naturally in a standard helper module
'
' INPUTS
'   cPM
'     Source cPerformanceManager instance
'
'   TargetTopLeft
'     Top-left output cell for the report
'
'   ClearOutputArea (optional)
'     TRUE  => clears the CurrentRegion of TargetTopLeft first
'     FALSE => writes only into the resized target report block
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Reads the report as a 2D array
'   - Writes it to the target worksheet
'   - Applies lightweight formatting to headers and numeric columns
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Data                As Variant              'Structured report array
    Dim RowCount            As Long                 'Report row count
    Dim ColCount            As Long                 'Report column count
    Dim OutputRange         As Range                'Resolved output range

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Reject a missing class instance
        If cPM Is Nothing Then
            Err.Raise vbObjectError + 2600, _
                      "M_cPM_ReportHelpers.cPM_Report_WriteToRange", _
                      "cPerformanceManager instance cannot be Nothing."
        End If
    'Reject a missing output anchor cell
        If TargetTopLeft Is Nothing Then
            Err.Raise vbObjectError + 2601, _
                      "M_cPM_ReportHelpers.cPM_Report_WriteToRange", _
                      "TargetTopLeft cannot be Nothing."
        End If

'------------------------------------------------------------------------------
' READ REPORT ARRAY
'------------------------------------------------------------------------------
    'Read the structured checkpoint report as a 2D array
        Data = cPM.ReportAsArray
    'Resolve the report dimensions
        RowCount = UBound(Data, 1)
        ColCount = UBound(Data, 2)

'------------------------------------------------------------------------------
' RESOLVE OUTPUT RANGE
'------------------------------------------------------------------------------
    'Resolve the resized output range
        Set OutputRange = TargetTopLeft.Resize(RowCount, ColCount)

'------------------------------------------------------------------------------
' OPTIONAL CLEAR
'------------------------------------------------------------------------------
    'Clear the surrounding output area when requested
        If ClearOutputArea Then
            TargetTopLeft.CurrentRegion.Clear
        End If

'------------------------------------------------------------------------------
' WRITE REPORT
'------------------------------------------------------------------------------
    'Write the structured report array into the worksheet
        OutputRange.Value = Data

'------------------------------------------------------------------------------
' FORMAT OUTPUT
'------------------------------------------------------------------------------
    'Format the header row
        With OutputRange.Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    'Format numeric timing columns
        OutputRange.Columns(7).NumberFormat = "0.000000000"
        OutputRange.Columns(8).NumberFormat = "0.000000000"
    'Apply lightweight autofit
        OutputRange.EntireColumn.AutoFit

End Sub
