from pathlib import Path

path = Path('test/M_cDP_Test.bas')
text = path.read_text(encoding='utf-8')

def replace_once(old, new, label):
    global text
    count = text.count(old)
    if count != 1:
        raise SystemExit(f'{label}: expected 1 anchor, found {count}')
    text = text.replace(old, new, 1)

replace_once(
'''    'Run DP_RepairRuntime behavior checks
        TST_DP_RunSuiteSafe "RepairRuntime"
    'Run deterministic live-clock timer registration and drain checks
        TST_DP_RunSuiteSafe "Timer"
''',
'''    'Run DP_RepairRuntime behavior checks
        TST_DP_RunSuiteSafe "RepairRuntime"
    'Run transactional startup / shutdown / repair fault matrices
        TST_DP_RunSuiteSafe "LifecycleTransaction"
    'Run deterministic live-clock timer registration and drain checks
        TST_DP_RunSuiteSafe "Timer"
''',
'dispatch list')

replace_once(
'''            Case "REPAIRRUNTIME"
                TST_DP_RunSuite_RepairRuntime

            Case "TIMER"
                TST_DP_RunSuite_Timer
''',
'''            Case "REPAIRRUNTIME"
                TST_DP_RunSuite_RepairRuntime

            Case "LIFECYCLETRANSACTION"
                TST_DP_RunSuite_LifecycleTransaction

            Case "TIMER"
                TST_DP_RunSuite_Timer
''',
'suite switch')

suite = r'''Private Sub TST_DP_Lifecycle_ForceCleanForTest()

'
'==============================================================================
'                 FORCE CLEAN LIFECYCLE STATE FOR TEST
'==============================================================================
'   Establishes a deterministic blank DatePicker runtime between #50 fault cases.
'   This is harness-only cleanup and may use the explicit force-clear lease helper.
'==============================================================================

    On Error Resume Next
    M_Lifecycle_Test_ArmFault VBA.vbNullString, 0
    M_Timer_Test_ArmScheduleFault 0
    M_Timer_Stop
    M_Timer_Test_Reset
    TST_DP_UnloadAllPickerFormsForTest
    Set gDP_Manager = Nothing
    M_ContextMenu_Remove
    M_KeyboardShortcut_Remove
    M_GridIcon_PurgeAll
    TST_DP_ForceClearLeaseForTest
    M_Lifecycle_Test_Reset True
    Err.Clear
    On Error GoTo 0

End Sub

Private Sub TST_DP_RunSuite_LifecycleTransaction()

'
'==============================================================================
'                    LIFECYCLE TRANSACTION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Proves #50 transactional startup, shutdown, repair, and lease-release rules
'
' BEHAVIOR
'   Exercises every post-mutation startup fault point, pre-owned startup failure,
'   every critical shutdown fault boundary, verified lease-deletion failure, a
'   compound manager + unresolved-timer failure, protected owned-grid deletion,
'   repair refusal on incomplete cleanup, and every post-mutation repair fault
'
' ERROR POLICY
'   Captures intentionally injected errors locally, records assertion failures,
'   and restores a clean owned runtime for following suites
'
' NOTES
'   The suite contains no package-certification claim. It is source-host evidence
'   only. #63 remains the final .xlsm/.xlam certification boundary.
'
' UPDATED
'   2026-09-05
'==============================================================================

    Const INJECT_BASE As Long = vbObjectError + 3600

    Dim StartupFaults As Variant
    Dim StopFaults As Variant
    Dim RepairFaults As Variant
    Dim FaultIndex As Long
    Dim FaultName As String
    Dim StepName As String
    Dim InjectedError As Long
    Dim EscapedError As Long
    Dim LeaseToken As String
    Dim TraceText As String
    Dim SavedEvents As Boolean
    Dim SavedShowGrid As Boolean
    Dim SavedShowRightClick As Boolean
    Dim SavedKeyboard As Boolean
    Dim SavedClockMode As DP_ClockMode
    Dim OwnedShape As Excel.Shape
    Dim ArmedEarliest As Date
    Dim ArmedLatest As Date
    Dim ArmedProcedure As String
    Dim ArmedSchedule As Boolean
    Dim ArmedError As Long
    Dim ArmedCalls As Long

'------------------------------------------------------------------------------
' INITIALIZE AND CAPTURE CALLER STATE
'------------------------------------------------------------------------------
    mTST_DP_CurrentSuite = "LifecycleTransaction"
    On Error GoTo SuiteFail

    SavedEvents = Excel.Application.EnableEvents
    SavedShowGrid = M_Settings_GetShowGridIcon()
    SavedShowRightClick = M_Settings_GetShowRightClick()
    SavedKeyboard = M_Settings_GetEnableKeyboardShortcut()
    SavedClockMode = M_Settings_GetClockMode()

    M_Settings_SetShowGridIcon True
    M_Settings_SetShowRightClick True
    M_Settings_SetEnableKeyboardShortcut True
    M_Settings_SetClockMode DP_ClockMode_Static

    TST_DP_ActivateWorksheetForTest mTST_DP_ScratchSheet
    mTST_DP_ScratchSheet.Range("D5").Value = VBA.DateSerial(2026, 9, 5)
    mTST_DP_ScratchSheet.Range("D5").NumberFormat = "dd/mm/yyyy"
    mTST_DP_ScratchSheet.Range("D5").Select

'------------------------------------------------------------------------------
' FRESH START: EVERY POST-MUTATION BOUNDARY ROLLS BACK
'------------------------------------------------------------------------------
    StartupFaults = Array( _
        "Start.AfterAdmission", _
        "Start.AfterManager", _
        "Start.AfterContextMenu", _
        "Start.AfterKeyboard", _
        "Start.AfterGrid", _
        "Start.AfterRefresh")

    For FaultIndex = LBound(StartupFaults) To UBound(StartupFaults)
        FaultName = VBA.CStr(StartupFaults(FaultIndex))
        InjectedError = INJECT_BASE + FaultIndex

        TST_DP_Lifecycle_ForceCleanForTest
        Excel.Application.EnableEvents = False
        M_Lifecycle_Test_Reset True
        M_Lifecycle_Test_ArmFault FaultName, InjectedError

        EscapedError = 0
        On Error Resume Next
        Err.Clear
        DP_Start
        EscapedError = Err.Number
        Err.Clear
        On Error GoTo SuiteFail

        TraceText = M_Lifecycle_Test_LastTrace()

        TST_DP_AssertEqualsLong "Fresh start preserves primary error | " & FaultName, _
            InjectedError, EscapedError
        TST_DP_AssertEqualsLong "Diagnostic primary error matches | " & FaultName, _
            InjectedError, M_Lifecycle_Test_LastPrimaryNumber()
        TST_DP_AssertEqualsString "Operation is DP_Start | " & FaultName, _
            "DP_Start", M_Lifecycle_Test_LastOperation()
        TST_DP_AssertFalse "Fresh failed start is not successful | " & FaultName, _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertFalse "Fresh start did not pre-own lease | " & FaultName, _
            M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
        TST_DP_AssertTrue "Fresh start acquired lease this call | " & FaultName, _
            M_Lifecycle_Test_LastLeaseAcquiredThisCall()
        TST_DP_AssertTrue "Fresh start attempted rollback | " & FaultName, _
            M_Lifecycle_Test_LastCleanupAttempted()
        TST_DP_AssertTrue "Fresh rollback reaches critical-clean state | " & FaultName, _
            M_Lifecycle_Test_LastCriticalClean()
        TST_DP_AssertEqualsLong "Fresh rollback has no cleanup failures | " & FaultName, _
            0, M_Lifecycle_Test_LastCleanupFailureCount()
        TST_DP_AssertTrue "Fresh rollback releases newly acquired lease | " & FaultName, _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertFalse "No lease remains after fresh rollback | " & FaultName, _
            M_Lease_IsOwner()
        TST_DP_AssertFalse "No local lease token remains after fresh rollback | " & FaultName, _
            M_Lifecycle_Test_HasLocalOwnerToken()
        TST_DP_AssertTrue "Fresh rollback trace reaches lease | " & FaultName, _
            VBA.InStr(1, TraceText, "Lease=PASS", vbBinaryCompare) > 0
        TST_DP_AssertTrue "Fresh rollback releases manager | " & FaultName, _
            (gDP_Manager Is Nothing)
        TST_DP_AssertFalse "Fresh rollback unloads picker | " & FaultName, _
            TST_DP_IsPickerFormLoadedForTest()
        TST_DP_AssertEqualsLong "Fresh rollback removes owned grid icon | " & FaultName, _
            0, TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
        TST_DP_AssertFalse "Fresh rollback restores disabled caller events | " & FaultName, _
            Excel.Application.EnableEvents
    Next FaultIndex

'------------------------------------------------------------------------------
' PRE-OWNED START FAILURE MUST NOT RELEASE THE EXISTING RUNTIME
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ForceCleanForTest
    Excel.Application.EnableEvents = False
    DP_Start
    LeaseToken = TST_DP_ReadLeaseOwnerForTest()

    M_Lifecycle_Test_Reset True
    InjectedError = INJECT_BASE + 20
    M_Lifecycle_Test_ArmFault "Start.AfterContextMenu", InjectedError

    EscapedError = 0
    On Error Resume Next
    Err.Clear
    DP_Start
    EscapedError = Err.Number
    Err.Clear
    On Error GoTo SuiteFail

    TST_DP_AssertEqualsLong "Repeated start preserves its primary error", _
        InjectedError, EscapedError
    TST_DP_AssertTrue "Repeated start classifies lease as pre-owned", _
        M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
    TST_DP_AssertFalse "Repeated start acquired no new lease", _
        M_Lifecycle_Test_LastLeaseAcquiredThisCall()
    TST_DP_AssertFalse "Repeated start does not destructively roll back pre-owned runtime", _
        M_Lifecycle_Test_LastCleanupAttempted()
    TST_DP_AssertFalse "Repeated start never reports lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Repeated start keeps exact pre-owned lease token", _
        LeaseToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Repeated start keeps ownership", M_Lease_IsOwner()
    TST_DP_AssertFalse "Repeated start keeps manager alive", (gDP_Manager Is Nothing)
    TST_DP_AssertTrue "Repeated start records preserved pre-owned runtime", _
        VBA.InStr(1, M_Lifecycle_Test_LastTrace(), _
            "PreOwnedRuntime=PRESERVED", vbBinaryCompare) > 0

    M_Lifecycle_Test_Reset True
    DP_Stop

'------------------------------------------------------------------------------
' SHUTDOWN: EVERY CRITICAL BOUNDARY CONTINUES, RETAINS, THEN RETRIES
'------------------------------------------------------------------------------
    StopFaults = Array( _
        "Cleanup.Manager", _
        "Cleanup.Timer", _
        "Cleanup.Form", _
        "Cleanup.Grid", _
        "Cleanup.ContextMenu", _
        "Cleanup.Keyboard", _
        "Cleanup.EnableEvents", _
        "Cleanup.Lease")

    For FaultIndex = LBound(StopFaults) To UBound(StopFaults)
        FaultName = VBA.CStr(StopFaults(FaultIndex))
        StepName = VBA.Mid$(FaultName, VBA.Len("Cleanup.") + 1)
        InjectedError = INJECT_BASE + 40 + FaultIndex

        TST_DP_Lifecycle_ForceCleanForTest
        Excel.Application.EnableEvents = False
        DP_Start
        LeaseToken = TST_DP_ReadLeaseOwnerForTest()

        M_Lifecycle_Test_Reset True
        M_Lifecycle_Test_ArmFault FaultName, InjectedError
        DP_Stop
        TraceText = M_Lifecycle_Test_LastTrace()

        TST_DP_AssertEqualsString "Stop operation is observable | " & StepName, _
            "DP_Stop", M_Lifecycle_Test_LastOperation()
        TST_DP_AssertFalse "Faulted stop is not successful | " & StepName, _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertTrue "Faulted stop records cleanup attempted | " & StepName, _
            M_Lifecycle_Test_LastCleanupAttempted()
        TST_DP_AssertEqualsLong "Faulted stop counts one failed cleanup | " & StepName, _
            1, M_Lifecycle_Test_LastCleanupFailureCount()
        TST_DP_AssertFalse "Faulted stop does not release lease | " & StepName, _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertTrue "Faulted stop retains local ownership token | " & StepName, _
            M_Lifecycle_Test_HasLocalOwnerToken()
        TST_DP_AssertTrue "Faulted stop retains verified ownership | " & StepName, _
            M_Lease_IsOwner()
        TST_DP_AssertEqualsString "Faulted stop retains exact lease token | " & StepName, _
            LeaseToken, TST_DP_ReadLeaseOwnerForTest()
        TST_DP_AssertTrue "Fault trace records failed step | " & StepName, _
            VBA.InStr(1, TraceText, StepName & "=FAIL", vbBinaryCompare) > 0
        TST_DP_AssertTrue "Fault trace reaches lease decision | " & StepName, _
            VBA.InStr(1, TraceText, "Lease=", vbBinaryCompare) > 0

        If VBA.StrComp(StepName, "Lease", vbBinaryCompare) = 0 Then
            TST_DP_AssertTrue "Lease-only fault follows critical-clean teardown", _
                M_Lifecycle_Test_LastCriticalClean()
        Else
            TST_DP_AssertFalse "Pre-lease critical fault marks teardown unclean | " & StepName, _
                M_Lifecycle_Test_LastCriticalClean()
        End If

        M_Lifecycle_Test_Reset True
        DP_Stop
        TST_DP_AssertTrue "Retry stop succeeds | " & StepName, _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertTrue "Retry releases lease | " & StepName, _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertEqualsString "Retry leaves no lease | " & StepName, _
            VBA.vbNullString, TST_DP_ReadLeaseOwnerForTest()
    Next FaultIndex

'------------------------------------------------------------------------------
' VERIFIED LEASE DELETE FAILURE RETAINS BOTH LIVE LEASE AND LOCAL PROOF
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ForceCleanForTest
    Excel.Application.EnableEvents = False
    DP_Start
    LeaseToken = TST_DP_ReadLeaseOwnerForTest()

    M_Lifecycle_Test_Reset True
    M_Lifecycle_Test_ArmFault "Lease.Delete", INJECT_BASE + 70
    DP_Stop

    TST_DP_AssertTrue "Lease-delete fault follows critical-clean teardown", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertEqualsLong "Lease-delete fault counts one cleanup failure", _
        1, M_Lifecycle_Test_LastCleanupFailureCount()
    TST_DP_AssertFalse "Lease-delete fault does not report release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertTrue "Lease-delete fault retains local owner token", _
        M_Lifecycle_Test_HasLocalOwnerToken()
    TST_DP_AssertEqualsString "Lease-delete fault retains exact live lease", _
        LeaseToken, TST_DP_ReadLeaseOwnerForTest()

    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Lease-delete retry releases cleanly", _
        M_Lifecycle_Test_LastLeaseReleased()

'------------------------------------------------------------------------------
' COMPOUND FAILURE: MANAGER CLEANUP FAULT + REAL #27 UNRESOLVED TIMER
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ForceCleanForTest
    Excel.Application.EnableEvents = False
    DP_Start

    M_Timer_Stop
    M_Timer_Test_Reset
    M_Timer_Start
    M_Timer_Test_LastRegistration ArmedEarliest, ArmedLatest, ArmedProcedure, _
        ArmedSchedule, ArmedError, ArmedCalls

    M_Timer_Test_ArmScheduleFault 1004
    M_Timer_Stop
    TST_DP_AssertTrue "Compound setup creates unresolved timer registration", _
        M_Timer_Test_IsUnresolved()

    'Keep the unresolved retry failing when the lifecycle timer cleanup reaches it.
    'Cleanup.Manager is injected too, so Class_Terminate cannot consume this timer fault first.
    M_Timer_Test_ArmScheduleFault 1004
    M_Lifecycle_Test_Reset True
    M_Lifecycle_Test_ArmFault "Cleanup.Manager", INJECT_BASE + 80
    DP_Stop
    TraceText = M_Lifecycle_Test_LastTrace()

    TST_DP_AssertEqualsLong "Compound cleanup records two independent failures", _
        2, M_Lifecycle_Test_LastCleanupFailureCount()
    TST_DP_AssertTrue "Compound trace records manager failure", _
        VBA.InStr(1, TraceText, "Manager=FAIL", vbBinaryCompare) > 0
    TST_DP_AssertTrue "Compound trace consumes #27 unresolved timer as failure", _
        VBA.InStr(1, TraceText, "Timer=FAIL", vbBinaryCompare) > 0
    TST_DP_AssertTrue "Compound cleanup continues after both failures", _
        VBA.InStr(1, TraceText, "Keyboard=PASS", vbBinaryCompare) > 0
    TST_DP_AssertFalse "Compound failure retains lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertTrue "Compound failure retains ownership", M_Lease_IsOwner()

    TST_DP_CancelRegistrationForTest ArmedEarliest, ArmedProcedure
    M_Timer_Test_ArmScheduleFault 0
    M_Timer_Test_Reset
    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Compound-failure retry releases lease", _
        M_Lifecycle_Test_LastLeaseReleased()

'------------------------------------------------------------------------------
' NATURAL #53 GRID FAILURE: OWNED SHAPE ON PROTECTED SHEET
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ForceCleanForTest
    Excel.Application.EnableEvents = False
    DP_Start
    M_GridIcon_PurgeAll

    Set OwnedShape = TST_DP_AddNamedShapeForTest( _
        mTST_DP_ScratchSheet, _
        DP_GRID_ICON_NAME, _
        "DatePicker Grid Entry Point | dp-owner-v1=20200101010101-00000000-DEADBEEF")
    mTST_DP_ScratchSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    M_Lifecycle_Test_Reset True
    DP_Stop
    TraceText = M_Lifecycle_Test_LastTrace()

    TST_DP_AssertTrue "Protected owned grid shape is reported as cleanup failure", _
        VBA.InStr(1, TraceText, "Grid=FAIL", vbBinaryCompare) > 0
    TST_DP_AssertEqualsLong "Protected grid case counts one cleanup failure", _
        1, M_Lifecycle_Test_LastCleanupFailureCount()
    TST_DP_AssertFalse "Protected owned grid blocks lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertTrue "Protected owned grid keeps lease ownership", M_Lease_IsOwner()
    TST_DP_AssertEqualsLong "Protected owned grid shape remains in place", _
        1, TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

    mTST_DP_ScratchSheet.Unprotect
    Set OwnedShape = Nothing
    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Unprotected owned grid retries and releases lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsLong "Retry reclaims formerly protected owned grid", _
        0, TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)

'------------------------------------------------------------------------------
' REPAIR PREPARE FAILURE REFUSES TO REBUILD OVER INCOMPLETE CLEANUP
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ForceCleanForTest
    Excel.Application.EnableEvents = False
    DP_Start
    LeaseToken = TST_DP_ReadLeaseOwnerForTest()

    M_Lifecycle_Test_Reset True
    M_Lifecycle_Test_ArmFault "Cleanup.Grid", INJECT_BASE + 90
    EscapedError = 0
    On Error Resume Next
    Err.Clear
    DP_RepairRuntime
    EscapedError = Err.Number
    Err.Clear
    On Error GoTo SuiteFail

    TST_DP_AssertTrue "Repair prepare failure raises to caller", EscapedError <> 0
    TST_DP_AssertEqualsString "Repair prepare operation is observable", _
        "DP_RepairRuntime", M_Lifecycle_Test_LastOperation()
    TST_DP_AssertFalse "Repair does not report success over incomplete cleanup", _
        M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertTrue "Repair prepare records cleanup attempted", _
        M_Lifecycle_Test_LastCleanupAttempted()
    TST_DP_AssertFalse "Repair prepare does not release lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Repair prepare retains exact lease", _
        LeaseToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Repair prepare retains ownership", M_Lease_IsOwner()
    TST_DP_AssertTrue "Repair remains the EnableEvents=True exception", _
        Excel.Application.EnableEvents
    TST_DP_AssertTrue "Repair does not rebuild manager over incomplete cleanup", _
        (gDP_Manager Is Nothing)

    M_Lifecycle_Test_Reset True
    DP_RepairRuntime
    TST_DP_AssertTrue "Repair retry succeeds after cleanup is available", _
        M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertFalse "Successful repair rebuilds manager", _
        (gDP_Manager Is Nothing)

'------------------------------------------------------------------------------
' REPAIR: EVERY POST-MUTATION REBUILD FAULT CLEANS AGAIN AND RETAINS LEASE
'------------------------------------------------------------------------------
    RepairFaults = Array( _
        "Repair.AfterManager", _
        "Repair.AfterContextMenu", _
        "Repair.AfterKeyboard", _
        "Repair.AfterRefresh")

    For FaultIndex = LBound(RepairFaults) To UBound(RepairFaults)
        FaultName = VBA.CStr(RepairFaults(FaultIndex))
        InjectedError = INJECT_BASE + 100 + FaultIndex

        TST_DP_Lifecycle_ForceCleanForTest
        Excel.Application.EnableEvents = False
        DP_Start
        LeaseToken = TST_DP_ReadLeaseOwnerForTest()

        M_Lifecycle_Test_Reset True
        M_Lifecycle_Test_ArmFault FaultName, InjectedError
        EscapedError = 0
        On Error Resume Next
        Err.Clear
        DP_RepairRuntime
        EscapedError = Err.Number
        Err.Clear
        On Error GoTo SuiteFail

        TST_DP_AssertEqualsLong "Repair preserves rebuild primary error | " & FaultName, _
            InjectedError, EscapedError
        TST_DP_AssertEqualsLong "Repair diagnostic keeps primary error | " & FaultName, _
            InjectedError, M_Lifecycle_Test_LastPrimaryNumber()
        TST_DP_AssertEqualsString "Repair operation remains observable | " & FaultName, _
            "DP_RepairRuntime", M_Lifecycle_Test_LastOperation()
        TST_DP_AssertFalse "Faulted repair is not successful | " & FaultName, _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertTrue "Faulted repair classifies lease as pre-owned | " & FaultName, _
            M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
        TST_DP_AssertFalse "Faulted repair acquires no new lease | " & FaultName, _
            M_Lifecycle_Test_LastLeaseAcquiredThisCall()
        TST_DP_AssertTrue "Faulted repair attempts rollback cleanup | " & FaultName, _
            M_Lifecycle_Test_LastCleanupAttempted()
        TST_DP_AssertTrue "Faulted repair rollback is critical-clean | " & FaultName, _
            M_Lifecycle_Test_LastCriticalClean()
        TST_DP_AssertEqualsLong "Faulted repair rollback has no cleanup failures | " & FaultName, _
            0, M_Lifecycle_Test_LastCleanupFailureCount()
        TST_DP_AssertFalse "Faulted repair retains lease | " & FaultName, _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertEqualsString "Faulted repair retains exact lease | " & FaultName, _
            LeaseToken, TST_DP_ReadLeaseOwnerForTest()
        TST_DP_AssertTrue "Faulted repair keeps ownership | " & FaultName, _
            M_Lease_IsOwner()
        TST_DP_AssertTrue "Faulted repair leaves events enabled | " & FaultName, _
            Excel.Application.EnableEvents
        TST_DP_AssertTrue "Faulted repair leaves no partial manager | " & FaultName, _
            (gDP_Manager Is Nothing)

        M_Lifecycle_Test_Reset True
        DP_RepairRuntime
        TST_DP_AssertTrue "Repair retry succeeds | " & FaultName, _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertFalse "Repair retry rebuilds manager | " & FaultName, _
            (gDP_Manager Is Nothing)
    Next FaultIndex

'------------------------------------------------------------------------------
' SUITE EXIT
'------------------------------------------------------------------------------
SuiteExit:
    On Error Resume Next
    If Not mTST_DP_ScratchSheet Is Nothing Then mTST_DP_ScratchSheet.Unprotect
    Set OwnedShape = Nothing
    TST_DP_CancelRegistrationForTest ArmedEarliest, ArmedProcedure
    M_Timer_Test_ArmScheduleFault 0
    M_Timer_Test_Reset
    TST_DP_Lifecycle_ForceCleanForTest

    M_Settings_SetShowGridIcon SavedShowGrid
    M_Settings_SetShowRightClick SavedShowRightClick
    M_Settings_SetEnableKeyboardShortcut SavedKeyboard
    M_Settings_SetClockMode SavedClockMode
    Excel.Application.EnableEvents = SavedEvents

    'Restore a normal owned runtime for the suites that follow.
    DP_Start
    Excel.Application.EnableEvents = SavedEvents
    Err.Clear
    On Error GoTo 0
    Exit Sub

'------------------------------------------------------------------------------
' SUITE FAIL
'------------------------------------------------------------------------------
SuiteFail:
    TST_DP_RecordFail "LifecycleTransaction suite failed", _
        "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    Err.Clear
    Resume SuiteExit

End Sub

'''

anchor = 'Private Sub TST_DP_RunSuite_Timer()\n'
if anchor not in text:
    raise SystemExit('Timer suite anchor not found')
text = text.replace(anchor, suite + anchor, 1)

# Static guardrails for topology and accidental duplication.
if text.count('TST_DP_RunSuiteSafe "LifecycleTransaction"') != 1:
    raise SystemExit('LifecycleTransaction dispatch count invalid')
if text.count('Case "LIFECYCLETRANSACTION"') != 1:
    raise SystemExit('LifecycleTransaction switch count invalid')
if text.count('Private Sub TST_DP_RunSuite_LifecycleTransaction()') != 1:
    raise SystemExit('LifecycleTransaction procedure count invalid')
if text.count('Private Sub TST_DP_Lifecycle_ForceCleanForTest()') != 1:
    raise SystemExit('Lifecycle helper count invalid')

path.write_text(text, encoding='utf-8')
