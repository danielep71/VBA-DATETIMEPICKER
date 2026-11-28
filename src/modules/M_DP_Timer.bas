Attribute VB_Name = "M_DP_Timer"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_Timer
'------------------------------------------------------------------------------
' PURPOSE
'   Runs the picker's live footer clock through Application.OnTime
'
' WHY THIS EXISTS
'   Excel keeps an OnTime registration after the code that made it has
'   forgotten it. One module owns every registration and its bookkeeping,
'   so a stop, a retry or a stale callback is decided in one place (#57)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - start, stop, tick and clock-mode application
'     - recovery of a registration whose cancellation failed
'     - M_Timer_TryStop, the verified stop used by lifecycle cleanup
'     - test seams for the scheduling state
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   Application.OnTime
'   M_DP_MacroNames (M_GetQualifiedMacroName)
'   M_DP_Core (gDP_ClockMode, DP_FORM_NAME, M_Settings_EnsureLoaded,
'     M_FormBridge_GetLoadedForm)
'   UF_DatePicker
'
' NOTES
'   Every registration names the facade callback M_Timer_Tick, which
'   delegates to M_Timer_Tick_Impl here
'
'   Option Private Module keeps every procedure out of other VBA projects
'   and out of the macro dialog. Procedures are Public only where another
'   module of the project calls them
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const DP_TIMER_SECONDS        As Long = 1   'Live clock interval in seconds
    Private Const DP_TIMER_LATEST_SECONDS As Long = 30  'Bounded OnTime delivery window in seconds

'------------------------------------------------------------------------------
' PRIVATE STATE
'------------------------------------------------------------------------------
    Private mDP_NextTickTime               As Date      'Next OnTime tick
    Private mDP_TimerIsRunning             As Boolean   'Timer running flag
    Private mDP_TimerProcedureName         As String    'Qualified OnTime timer procedure name
    Private mDP_TimerProcedureWorkbookName As String    'Workbook name used for cached timer callback

    Private mDP_TimerLatestTime            As Date      'LatestTime of the current registration
    Private mDP_TimerLastEarliest          As Date      'EarliestTime of the last scheduling call
    Private mDP_TimerLastLatest            As Date      'LatestTime of the last scheduling call
    Private mDP_TimerLastProcedure         As String    'Procedure of the last scheduling call
    Private mDP_TimerLastSchedule          As Boolean   'Schedule flag of the last scheduling call
    Private mDP_TimerLastErrNumber         As Long      'Error number from the last scheduling call
    Private mDP_TimerLastErrDescription    As String    'Error description from the last scheduling call
    Private mDP_TimerCallCount             As Long      'Number of scheduling calls made
    Private mDP_TimerFaultNumber           As Long      'One-shot injected scheduling failure, 0 when disarmed

    Private mDP_TimerUnresolved            As Boolean   'True while a failed cancellation leaves a registration outstanding
    Private mDP_TimerUnresolvedEarliest    As Date      'EarliestTime of the unresolved registration
    Private mDP_TimerUnresolvedLatest      As Date      'LatestTime of the unresolved registration
    Private mDP_TimerUnresolvedProcedure   As String    'Procedure of the unresolved registration
    Private mDP_TimerRefusalReported       As Boolean   'Diagnostic-only: restart refusal already reported


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

Private Function M_Timer_ApplySchedule( _
    ByVal EarliestTime As Date, _
    ByVal LatestTime As Date, _
    ByVal ProcedureName As String, _
    ByVal ScheduleFlag As Boolean) As Boolean

'
'------------------------------------------------------------------------------
'                        APPLY A TIMER SCHEDULING CALL
'------------------------------------------------------------------------------
' PURPOSE
'   The single point through which every Application.OnTime call is made, and the
'   recorder of what was asked for
'
' WHY THIS EXISTS
'   Cancellation needs the exact EarliestTime and Procedure that were scheduled,
'   and nothing could previously observe either. Routing every call site
'   through one routine records the identity of every registration and gives the
'   regression a deterministic way to make scheduling or cancellation fail
'   without waiting for wall-clock time
'
' INPUTS
'   EarliestTime
'     Registration time, and for a cancellation the time to match
'
'   LatestTime
'     Bounded delivery deadline; ignored for a cancellation
'
'   ProcedureName
'     Workbook-qualified callback name
'
'   ScheduleFlag
'     True to schedule, False to cancel
'
' RETURNS
'   True when the call was accepted
'
' BEHAVIOR
'   Records the requested registration, consumes a one-shot injected fault when
'   armed, otherwise calls Application.OnTime and records the outcome
'
' ERROR POLICY
'   Never raises outward. Callers read the recorded error and decide
'
' DEPENDENCIES
'   Application.OnTime
'
' NOTES
'   LatestTime is passed only when scheduling. Application.OnTime matches a
'   cancellation on EarliestTime and Procedure alone
'
'   The injected fault is one-shot and is consumed whether or not it is observed,
'   so an armed fault can never leak into a later test
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim FaultNumber     As Long         'One-shot injected failure being consumed

'------------------------------------------------------------------------------
' RECORD THE REQUEST
'------------------------------------------------------------------------------
    'Never let a scheduling call raise into its caller
        On Error Resume Next
    'Record what this call asked for, before it is attempted
        mDP_TimerLastEarliest = EarliestTime
        mDP_TimerLastLatest = LatestTime
        mDP_TimerLastProcedure = ProcedureName
        mDP_TimerLastSchedule = ScheduleFlag
        mDP_TimerLastErrNumber = 0
        mDP_TimerLastErrDescription = VBA.vbNullString
        mDP_TimerCallCount = mDP_TimerCallCount + 1
    'Assume the call fails until it is accepted
        M_Timer_ApplySchedule = False

'------------------------------------------------------------------------------
' CONSUME AN INJECTED FAULT
'------------------------------------------------------------------------------
    'Consume a one-shot injected failure instead of calling Excel
        If mDP_TimerFaultNumber <> 0 Then
            FaultNumber = mDP_TimerFaultNumber
            mDP_TimerFaultNumber = 0
            mDP_TimerLastErrNumber = FaultNumber
            mDP_TimerLastErrDescription = "Injected timer scheduling failure"
            Err.Clear
            Exit Function
        End If

'------------------------------------------------------------------------------
' CALL EXCEL
'------------------------------------------------------------------------------
    'Schedule with a bounded delivery window, or cancel by matching identity
        Err.Clear
        If ScheduleFlag Then
            Excel.Application.OnTime _
                EarliestTime:=EarliestTime, _
                Procedure:=ProcedureName, _
                LatestTime:=LatestTime, _
                Schedule:=True
        Else
            Excel.Application.OnTime _
                EarliestTime:=EarliestTime, _
                Procedure:=ProcedureName, _
                Schedule:=False
        End If
    'Record the outcome
        If Err.Number = 0 Then
            M_Timer_ApplySchedule = True
        Else
            mDP_TimerLastErrNumber = Err.Number
            mDP_TimerLastErrDescription = Err.Description
        End If
    'Clear any suppressed scheduling error
        Err.Clear

End Function

Private Sub M_Timer_RetainUnresolved( _
    ByVal EarliestTime As Date, _
    ByVal LatestTime As Date, _
    ByVal ProcedureName As String)

'
'------------------------------------------------------------------------------
'                      RETAIN AN UNRESOLVED REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Records the exact registration a failed cancellation left outstanding
'
' WHY THIS EXISTS
'   A Boolean running flag cannot distinguish generations. If an old callback
'   arrives after a restart, the flag is True again and the callback cannot know
'   which generation invoked it. Retaining the exact registration lets a restart
'   be refused until the old one is provably gone
'
' INPUTS
'   EarliestTime, LatestTime, ProcedureName
'     Identity of the registration that could not be cancelled
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the registration and marks it unresolved
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   None
'
' NOTES
'   This is deliberately distinct from a healthy current registration. An
'   unresolved registration must never be treated as a running timer
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETAIN THE REGISTRATION
'------------------------------------------------------------------------------
    'Never let bookkeeping raise into a caller
        On Error Resume Next
    'Store the exact identity the cancellation failed on
        mDP_TimerUnresolvedEarliest = EarliestTime
        mDP_TimerUnresolvedLatest = LatestTime
        mDP_TimerUnresolvedProcedure = ProcedureName
        mDP_TimerUnresolved = True
    'Clear any suppressed bookkeeping error
        Err.Clear

End Sub

Private Sub M_Timer_ResolveUnresolved()

'
'------------------------------------------------------------------------------
'                     RESOLVE AN UNRESOLVED REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Drains the outstanding registration once it can no longer fire
'
' WHY THIS EXISTS
'   The restart barrier has to lift. Draining is centralized so the three valid
'   drains cannot diverge
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Clears the retained registration and the refusal diagnostic flag
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   None
'
' NOTES
'   Called by all three drains: a stale callback arriving, a successful retry
'   cancellation, and expiry of the retained LatestTime
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DRAIN THE REGISTRATION
'------------------------------------------------------------------------------
    'Never let bookkeeping raise into a caller
        On Error Resume Next
    'Forget the retained registration
        mDP_TimerUnresolved = False
        mDP_TimerUnresolvedEarliest = 0
        mDP_TimerUnresolvedLatest = 0
        mDP_TimerUnresolvedProcedure = VBA.vbNullString
    'Allow a later refusal to report again
        mDP_TimerRefusalReported = False
    'Clear any suppressed bookkeeping error
        Err.Clear

End Sub

Private Function M_Timer_TryDrainUnresolved() As Boolean

'
'------------------------------------------------------------------------------
'                       TRY TO DRAIN THE REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a fresh registration may be scheduled
'
' WHY THIS EXISTS
'   Two of the three drains can be attempted on demand: retrying the cancellation
'   with the retained identity, and observing that the retained LatestTime has
'   passed. The third, a stale callback arriving, drains itself
'
' INPUTS
'   None
'
' RETURNS
'   True when nothing is outstanding and scheduling may proceed
'
' BEHAVIOR
'   Returns True when no registration is outstanding. Otherwise retries the
'   cancellation using the exact retained identity, and failing that checks
'   whether the retained LatestTime has passed
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Timer_ApplySchedule
'   M_Timer_ResolveUnresolved
'
' NOTES
'   Expiry is sound only because every tick is scheduled with an explicit
'   LatestTime. With LatestTime omitted, Excel may run a callback well after its
'   EarliestTime, so elapsed time alone would prove nothing
'
'   The retry runs on every attempt. Only the refusal diagnostic is
'   de-duplicated; suppressing the retry would strand the barrier
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ATTEMPT THE DRAINS
'------------------------------------------------------------------------------
    'Never let a drain attempt raise into a caller
        On Error Resume Next
    'Nothing outstanding means scheduling may proceed
        If Not mDP_TimerUnresolved Then
            M_Timer_TryDrainUnresolved = True
            Err.Clear
            Exit Function
        End If
    'Retry the cancellation against the exact retained identity
        If VBA.LenB(mDP_TimerUnresolvedProcedure) > 0 Then
            If M_Timer_ApplySchedule(mDP_TimerUnresolvedEarliest, 0, _
                mDP_TimerUnresolvedProcedure, False) Then
                M_Timer_ResolveUnresolved
                M_Timer_TryDrainUnresolved = True
                Err.Clear
                Exit Function
            End If
        End If
    'Treat the registration as expired once its bounded window has passed
        If mDP_TimerUnresolvedLatest <> 0 Then
            If VBA.Now > mDP_TimerUnresolvedLatest Then
                M_Timer_ResolveUnresolved
                M_Timer_TryDrainUnresolved = True
                Err.Clear
                Exit Function
            End If
        End If
    'Refuse a fresh registration while the old one may still fire
        M_Timer_TryDrainUnresolved = False
    'Clear any suppressed drain error
        Err.Clear

End Function

Private Sub M_Timer_ReportRestartRefusal(ByVal EntryPoint As String)

'
'------------------------------------------------------------------------------
'                        REPORT A RESTART REFUSAL
'------------------------------------------------------------------------------
' PURPOSE
'   Writes a diagnostic once while restarts are refused
'
' WHY THIS EXISTS
'   A refusal can repeat on every clock-mode change and every form activation.
'   Silence would leave a static clock with no explanation
'
' INPUTS
'   EntryPoint
'     Name of the refusing routine
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Prints only when a refusal has not already been reported since the last drain
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_TimerRefusalReported
'
' NOTES
'   Diagnostic only. It never suppresses the cancellation retry and never affects
'   whether a registration is outstanding
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REPORT THE REFUSAL
'------------------------------------------------------------------------------
    'Never let a diagnostic raise into a caller
        On Error Resume Next
    'Report only once per outstanding registration
        If mDP_TimerRefusalReported Then
            Err.Clear
            Exit Sub
        End If
    'Remember that this refusal was reported
        mDP_TimerRefusalReported = True
    'Write the refusal diagnostic
        Debug.Print EntryPoint & _
            " | Refused | A previous timer registration could not be cancelled" & _
            " and may still fire; the live clock stays inactive until it drains"
    'Clear any suppressed diagnostic error
        Err.Clear

End Sub

Public Sub M_Timer_EnsureHealthy(ByVal EntryPoint As String)

'
'------------------------------------------------------------------------------
'                        ENSURE THE LIVE CLOCK IS HEALTHY
'------------------------------------------------------------------------------
' PURPOSE
'   The only cross-module entry point to the timer health check
'
' WHY THIS EXISTS
'   UF_DatePicker must be able to repair a dropped clock registration when the
'   form is reactivated, and M_Timer_CheckHealth is Private
'
'   The form previously reached it through M_Timer_ApplyClockMode. That does not
'   work. ApplyClockMode calls M_Timer_Stop unconditionally before anything else,
'   which clears the running flag, so the health check inside M_Timer_Start then
'   exits immediately and the recovery path was unreachable
'
'   In the case the call exists for it was worse than useless. With the
'   registration already dropped, that Stop attempted to cancel a registration
'   Excel no longer held; the cancellation failed, the unresolved barrier was
'   retained, and the following start was refused, leaving the clock dead until
'   the retained LatestTime expired
'
' INPUTS
'   EntryPoint
'     Name of the routine asking for the check, for diagnostics
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Delegates to M_Timer_CheckHealth and does nothing else
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Timer_CheckHealth
'
' NOTES
'   This routine must never stop the timer, load settings, reapply the clock
'   mode, resolve the form or refresh the UI
'
'   A healthy registration produces no scheduling call at all. An expired current
'   registration may produce exactly one replacement. A registration retained by
'   a failed cancellation stays governed by the barrier and its drains
'
'   Public only because the form is a separate module. It takes an argument, so
'   it does not appear in the macro dialog, and #25 classifies it as internal
'   rather than supported API
'
' UPDATED
'   2026-09-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DELEGATE
'------------------------------------------------------------------------------
    'Never let the bridge raise into a caller
        On Error Resume Next
    'Run the health check and nothing else
        M_Timer_CheckHealth EntryPoint
    'Clear any suppressed error
        Err.Clear

End Sub

Private Sub M_Timer_CheckHealth(ByVal EntryPoint As String)

'
'------------------------------------------------------------------------------
'                          CHECK TIMER HEALTH
'------------------------------------------------------------------------------
' PURPOSE
'   Repairs a live clock whose registration was dropped by Excel
'
' WHY THIS EXISTS
'   Every tick is scheduled with an explicit LatestTime, so a tick that cannot
'   run inside its window is dropped rather than delayed. Because each tick
'   schedules the next one, one dropped tick ends the chain: M_Timer_Tick never
'   runs, nothing reschedules, and the running flag stays True while no
'   registration exists
'
'   Nothing independent can observe that absence. This is opportunistic
'   self-healing driven by the next DatePicker interaction, not a watchdog
'
' INPUTS
'   EntryPoint
'     Name of the routine asking for the check, for diagnostics
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   When the timer is logically active, holds a current registration that is not
'   an unresolved one, and the retained LatestTime has passed, clears the stale
'   state and schedules exactly one fresh registration
'
' ERROR POLICY
'   Never raises outward. A failed recovery leaves the timer honestly inactive
'
' DEPENDENCIES
'   M_Timer_ApplySchedule
'   M_Timer_GetProcedureName
'
' NOTES
'   A healthy registration still inside its delivery window produces no
'   scheduling call at all
'
'   An unresolved registration left by a failed cancellation is explicitly not
'   healed here. That case is a barrier, not a dropped tick, and is drained by
'   M_Timer_TryDrainUnresolved
'
' UPDATED
'   2026-09-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim NextEarliest    As Date         'Registration time of the replacement tick
    Dim NextLatest      As Date         'Bounded delivery deadline of the replacement tick

'------------------------------------------------------------------------------
' DECIDE WHETHER RECOVERY IS DUE
'------------------------------------------------------------------------------
    'Never let a health check raise into a caller
        On Error Resume Next
    'A logically inactive timer has nothing to repair
        If Not mDP_TimerIsRunning Then GoTo ExitProcedure
    'An unresolved registration is a barrier, not a dropped tick
        If mDP_TimerUnresolved Then GoTo ExitProcedure
    'A registration with no bounded window cannot be proven expired
        If mDP_TimerLatestTime = 0 Then GoTo ExitProcedure
    'A registration still inside its delivery window is healthy
        If VBA.Now <= mDP_TimerLatestTime Then GoTo ExitProcedure

'------------------------------------------------------------------------------
' REPLACE THE EXPIRED REGISTRATION
'------------------------------------------------------------------------------
    'Clear the stale active-registration state before rescheduling
        mDP_TimerIsRunning = False
        mDP_NextTickTime = 0
        mDP_TimerLatestTime = 0
    'Rebuild the qualified callback name
        mDP_TimerProcedureName = M_Timer_GetProcedureName
        Err.Clear
    'Calculate exactly one replacement registration
        NextEarliest = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)
        NextLatest = NextEarliest + VBA.TimeSerial(0, 0, DP_TIMER_LATEST_SECONDS)
    'Arm the replacement only when the schedule call is accepted
        If M_Timer_ApplySchedule(NextEarliest, NextLatest, mDP_TimerProcedureName, True) Then
            mDP_TimerIsRunning = True
            mDP_NextTickTime = NextEarliest
            mDP_TimerLatestTime = NextLatest
        Else
            Debug.Print EntryPoint & " | M_Timer_CheckHealth" & _
                " | Step=RescheduleDroppedTick" & _
                " | Error=" & VBA.CStr(mDP_TimerLastErrNumber) & _
                " | " & mDP_TimerLastErrDescription
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
ExitProcedure:
    'Clear any suppressed health-check error
        Err.Clear

End Sub

Public Sub M_Timer_Test_ArmScheduleFault(ByVal ErrorNumber As Long)

'
'------------------------------------------------------------------------------
'                      ARM A TIMER SCHEDULING FAULT
'------------------------------------------------------------------------------
' PURPOSE
'   Makes the next scheduling or cancellation call fail deterministically
'
' WHY THIS EXISTS
'   Cancellation failure, restart refusal and dropped-tick recovery cannot be
'   exercised by waiting. The regression must be able to make one call fail
'
' INPUTS
'   ErrorNumber
'     Error number the next call reports; zero disarms
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Arms a one-shot fault consumed by the next M_Timer_ApplySchedule call
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_TimerFaultNumber
'
' NOTES
'   Public only because the regression harness is a separate module. It takes an
'   argument, so it does not appear in the macro dialog. It is classified
'   internal in src/public-surface.json
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ARM THE FAULT
'------------------------------------------------------------------------------
    'Never let arming raise into a caller
        On Error Resume Next
    'Store the one-shot fault
        mDP_TimerFaultNumber = ErrorNumber
    'Clear any suppressed arming error
        Err.Clear

End Sub

Public Sub M_Timer_Test_LastRegistration( _
    ByRef EarliestTime As Date, _
    ByRef LatestTime As Date, _
    ByRef ProcedureName As String, _
    ByRef ScheduleFlag As Boolean, _
    ByRef ErrorNumber As Long, _
    ByRef CallCount As Long)

'
'------------------------------------------------------------------------------
'                      READ THE LAST TIMER REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reports exactly what the last scheduling call asked Excel for
'
' WHY THIS EXISTS
'   The issue requires proof that start records one exact scheduled time and
'   qualified callback, and that stop cancels using the same values. Nothing
'   could observe either before
'
' INPUTS
'   EarliestTime, LatestTime, ProcedureName, ScheduleFlag, ErrorNumber, CallCount
'     Receive the recorded registration and the running call count
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Copies the recorded values out
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   None
'
' NOTES
'   Reading is free of side effects. It never drains, arms or reschedules
'   anything
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' COPY THE RECORDED VALUES
'------------------------------------------------------------------------------
    'Never let a read raise into a caller
        On Error Resume Next
    'Copy the recorded registration
        EarliestTime = mDP_TimerLastEarliest
        LatestTime = mDP_TimerLastLatest
        ProcedureName = mDP_TimerLastProcedure
        ScheduleFlag = mDP_TimerLastSchedule
        ErrorNumber = mDP_TimerLastErrNumber
        CallCount = mDP_TimerCallCount
    'Clear any suppressed read error
        Err.Clear

End Sub

Public Function M_Timer_Test_IsUnresolved() As Boolean

'
'------------------------------------------------------------------------------
'                   REPORT AN UNRESOLVED TIMER REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether a failed cancellation left a registration outstanding
'
' WHY THIS EXISTS
'   A cancellation failure previously reached only Debug.Print, so the state that
'   refuses a restart was unobservable. Whether that state should also make
'   shutdown incomplete and retain the lease is #50's contract, not this one
'
' INPUTS
'   None
'
' RETURNS
'   True while a registration is outstanding
'
' BEHAVIOR
'   Reports the retained state without attempting any drain
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_TimerUnresolved
'
' NOTES
'   Deliberately does not retry the cancellation, so a test can observe the
'   barrier before deciding to drain it
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REPORT THE STATE
'------------------------------------------------------------------------------
    'Never let a read raise into a caller
        On Error Resume Next
    'Report whether a registration is outstanding
        M_Timer_Test_IsUnresolved = mDP_TimerUnresolved
    'Clear any suppressed read error
        Err.Clear

End Function

Public Function M_Timer_Test_IsRunning() As Boolean

'
'------------------------------------------------------------------------------
'                       REPORT THE LOGICAL TIMER STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Reports whether the timer is logically active
'
' WHY THIS EXISTS
'   Several cases turn on the timer being honestly inactive after a failure, and
'   the running flag is private
'
' INPUTS
'   None
'
' RETURNS
'   True while the timer is logically active
'
' BEHAVIOR
'   Reports the flag without running the health check
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_TimerIsRunning
'
' NOTES
'   Deliberately does not call M_Timer_CheckHealth, so a test can observe a
'   dropped registration before recovery is triggered
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' REPORT THE STATE
'------------------------------------------------------------------------------
    'Never let a read raise into a caller
        On Error Resume Next
    'Report the logical timer state
        M_Timer_Test_IsRunning = mDP_TimerIsRunning
    'Clear any suppressed read error
        Err.Clear

End Function

Public Sub M_Timer_Test_ExpireRegistration()

'
'------------------------------------------------------------------------------
'                    EXPIRE THE RETAINED REGISTRATION
'------------------------------------------------------------------------------
' PURPOSE
'   Backdates the retained delivery deadlines so expiry can be tested
'
' WHY THIS EXISTS
'   The bounded window is thirty seconds. Both the expiry drain and dropped-tick
'   recovery depend on it having passed, and a regression cannot wait
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Moves the current and unresolved LatestTime values into the past
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   mDP_TimerLatestTime
'   mDP_TimerUnresolvedLatest
'
' NOTES
'   Backdates only the deadlines. It never cancels, schedules or drains anything,
'   so what the production code then does is the thing under test
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' BACKDATE THE DEADLINES
'------------------------------------------------------------------------------
    'Never let a test seam raise into a caller
        On Error Resume Next
    'Backdate the current registration deadline when one exists
        If mDP_TimerLatestTime <> 0 Then
            mDP_TimerLatestTime = VBA.Now - VBA.TimeSerial(0, 1, 0)
        End If
    'Backdate the unresolved registration deadline when one exists
        If mDP_TimerUnresolvedLatest <> 0 Then
            mDP_TimerUnresolvedLatest = VBA.Now - VBA.TimeSerial(0, 1, 0)
        End If
    'Clear any suppressed test-seam error
        Err.Clear

End Sub

Public Sub M_Timer_Test_Reset()

'
'------------------------------------------------------------------------------
'                        RESET TIMER TEST STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the recorder, the injected fault and the barrier to a known state
'
' WHY THIS EXISTS
'   Timer cases must not leak an armed fault, a recorded registration or an
'   outstanding barrier into the suites that follow
'
' INPUTS
'   None
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Disarms the fault, clears the recorder and drains any retained registration
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Timer_ResolveUnresolved
'
' NOTES
'   This clears bookkeeping only. It does not cancel a live registration, so a
'   suite must still stop the timer through the supported path
'
' UPDATED
'   2026-08-30
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RESET THE STATE
'------------------------------------------------------------------------------
    'Never let a reset raise into a caller
        On Error Resume Next
    'Disarm any injected fault
        mDP_TimerFaultNumber = 0
    'Clear the recorded registration
        mDP_TimerLastEarliest = 0
        mDP_TimerLastLatest = 0
        mDP_TimerLastProcedure = VBA.vbNullString
        mDP_TimerLastSchedule = False
        mDP_TimerLastErrNumber = 0
        mDP_TimerLastErrDescription = VBA.vbNullString
        mDP_TimerCallCount = 0
    'Drain any retained registration
        M_Timer_ResolveUnresolved
    'Clear any suppressed reset error
        Err.Clear

End Sub

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
'   2026-09-01
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Timer_Start"

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description
    Dim NextEarliest            As Date          'Registration time of the next tick
    Dim NextLatest              As Date          'Bounded delivery deadline of the next tick

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' HEAL A DROPPED REGISTRATION
'------------------------------------------------------------------------------
    'Repair a chain Excel dropped before deciding the timer is already running.
    'This runs before the guard below because a dropped registration leaves the
    'running flag True with nothing scheduled
        M_Timer_CheckHealth PROC_NAME

'------------------------------------------------------------------------------
' EXIT IF ALREADY RUNNING
'------------------------------------------------------------------------------
    'Exit if the timer is already running
        If mDP_TimerIsRunning Then Exit Sub

'------------------------------------------------------------------------------
' DRAIN AN UNRESOLVED REGISTRATION
'------------------------------------------------------------------------------
    'Refuse a fresh registration while an uncancelled one may still fire. The
    'retry runs on every attempt; only the diagnostic is de-duplicated
        If Not M_Timer_TryDrainUnresolved() Then
            M_Timer_ReportRestartRefusal PROC_NAME
            Exit Sub
        End If

'------------------------------------------------------------------------------
' RESOLVE TIMER PROCEDURE
'------------------------------------------------------------------------------
    'Build the workbook-qualified timer procedure name
        mDP_TimerProcedureName = M_Timer_GetProcedureName

'------------------------------------------------------------------------------
' CALCULATE NEXT TICK
'------------------------------------------------------------------------------
    'Calculate the next tick and its bounded delivery deadline
        NextEarliest = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)
        NextLatest = NextEarliest + VBA.TimeSerial(0, 0, DP_TIMER_LATEST_SECONDS)

'------------------------------------------------------------------------------
' SCHEDULE NEXT TICK
'------------------------------------------------------------------------------
    'Schedule the next timer tick and raise when Excel refuses it
        If Not M_Timer_ApplySchedule(NextEarliest, NextLatest, _
            mDP_TimerProcedureName, True) Then
            Err.Raise mDP_TimerLastErrNumber, PROC_NAME, mDP_TimerLastErrDescription
        End If

'------------------------------------------------------------------------------
' ARM THE TIMER
'------------------------------------------------------------------------------
    'Establish the armed state only after the schedule call was accepted, so a
    'refused schedule can never leave a false running state behind
        mDP_TimerIsRunning = True
        mDP_NextTickTime = NextEarliest
        mDP_TimerLatestTime = NextLatest
        mDP_TimerRefusalReported = False

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
    'Clear the delivery deadline after scheduling failure
        mDP_TimerLatestTime = 0
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
'   Rebuilds the qualified timer procedure name when needed and attempts to
'   cancel the next scheduled timer tick when a timer is active
'
'   A successful cancellation resolves the registration and leaves nothing
'   outstanding. A failed cancellation retains the exact registration identity
'   as unresolved, because that registration may still fire
'
'   The logical timer state is cleared either way, so an outstanding
'   registration is never mistaken for a running timer
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
'   Clearing the logical state is not the same as discarding the registration.
'   Before #27 a failed cancellation reached only Debug.Print while the state
'   claimed clean teardown, so a registration that could still fire was simply
'   forgotten. It is now retained through M_Timer_RetainUnresolved and refuses
'   a restart until one of its drains lifts the barrier
'
'   Whether an unresolved registration should also make shutdown incomplete and
'   retain the provider lease is #50's contract, not this routine's
'
' UPDATED
'   2026-09-01
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
                    If M_Timer_ApplySchedule(mDP_NextTickTime, 0, _
                        mDP_TimerProcedureName, False) Then
                        M_Timer_ResolveUnresolved
                    Else
                        CancelErrNumber = mDP_TimerLastErrNumber
                        CancelErrDescription = mDP_TimerLastErrDescription
                        M_Timer_RetainUnresolved mDP_NextTickTime, _
                            mDP_TimerLatestTime, mDP_TimerProcedureName
                    End If
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
    'Clear the delivery deadline. A registration that could not be cancelled is
    'retained separately as unresolved, never as a running timer
        mDP_TimerLatestTime = 0
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

Public Sub M_Timer_Tick_Impl()

'
'------------------------------------------------------------------------------
'                           TIMER TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Implements the supported entry point M_DatePicker.M_Timer_Tick
'
' WHY THIS EXISTS
'   The facade keeps the supported name and signature; the implementation
'   lives in this private module so internal code never calls the facade
'
' INPUTS
'   As documented on M_DatePicker.M_Timer_Tick
'
' RETURNS
'   As documented on M_DatePicker.M_Timer_Tick
'
' BEHAVIOR
'   As documented on M_DatePicker.M_Timer_Tick
'
' ERROR POLICY
'   As documented on M_DatePicker.M_Timer_Tick. PROC_NAME keeps the supported name, so the
'   error source a caller sees is unchanged
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
'   Internal code calls this procedure, never the facade wrapper
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME            As String = "M_Timer_Tick" 'Current procedure name

    Dim LoadedForm             As Object       'Loaded DatePicker form instance
    Dim NextEarliest           As Date         'Registration time of the next tick
    Dim NextLatest             As Date         'Bounded delivery deadline of the next tick
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
    'A callback arriving while the timer is inactive is the stale one a failed
    'cancellation left outstanding. Drain it, do no work and do not reschedule
        If Not mDP_TimerIsRunning Then
            M_Timer_ResolveUnresolved
            Exit Sub
        End If

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
    'Calculate the next tick and its bounded delivery deadline
        NextEarliest = VBA.Now + VBA.TimeSerial(0, 0, DP_TIMER_SECONDS)
        NextLatest = NextEarliest + VBA.TimeSerial(0, 0, DP_TIMER_LATEST_SECONDS)

'------------------------------------------------------------------------------
' SCHEDULE NEXT TICK
'------------------------------------------------------------------------------
    'Keep the chain alive only while Excel accepts the next registration
        If M_Timer_ApplySchedule(NextEarliest, NextLatest, _
            mDP_TimerProcedureName, True) Then
            mDP_NextTickTime = NextEarliest
            mDP_TimerLatestTime = NextLatest
        Else
            mDP_TimerIsRunning = False
            mDP_NextTickTime = 0
            mDP_TimerLatestTime = 0
        End If

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
'                              VERIFIED TEARDOWN
'
'------------------------------------------------------------------------------
'

Public Function M_Timer_TryStop( _
    ByRef ErrorNumber As Long, _
    ByRef ErrorDescription As String) As Boolean

'
'------------------------------------------------------------------------------
'                           STOP THE TIMER, VERIFIED
'------------------------------------------------------------------------------
' PURPOSE
'   Stops the live clock and proves nothing is left pending
'
' WHY THIS EXISTS
'   #27 made a failed cancellation observable: the registration is retained as
'   unresolved rather than forgotten. Teardown must consume that state, because a
'   retained callback can still fire after the runtime is gone
'
' INPUTS
'   ErrorNumber, ErrorDescription
'     Receive the failure when the timer stays active or stays unresolved
'
' RETURNS
'   True only when the timer is inactive and no registration is outstanding
'
' BEHAVIOR
'   Calls M_Timer_Stop, then rejects both a still-running timer and an unresolved
'   registration
'
' ERROR POLICY
'   Never raises outward
'
' DEPENDENCIES
'   M_Timer_Stop
'   mDP_TimerIsRunning
'   mDP_TimerUnresolved
'
' NOTES
'   The two rejections are separate errors on purpose. A timer that is still
'   running is a different defect from one that stopped but could not cancel
'
'   Moved from M_DP_Core by #57, where it was named
'   M_Lifecycle_TryStopTimer, so the module that owns this
'   integration's state also verifies its teardown
'
'   Public only so the lifecycle cleanup in M_DP_Core can call it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

    On Error GoTo Failed
    ErrorNumber = 0
    ErrorDescription = VBA.vbNullString

    M_Timer_Stop
    If mDP_TimerIsRunning Then
        ErrorNumber = vbObjectError + 2711
        ErrorDescription = "Timer remained logically active after stop"
        Exit Function
    End If
    If mDP_TimerUnresolved Then
        ErrorNumber = vbObjectError + 2712
        ErrorDescription = "Timer cancellation is unresolved; the retained callback may still fire"
        Exit Function
    End If

    M_Timer_TryStop = True
    Exit Function

Failed:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    Err.Clear
    M_Timer_TryStop = False

End Function
