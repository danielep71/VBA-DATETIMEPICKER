from pathlib import Path
import re

path = Path('test/M_cDP_Test.bas')
text = path.read_text(encoding='utf-8')

# Register the dedicated suite immediately after RepairRuntime.
old = '''    'Run DP_RepairRuntime behavior checks
        TST_DP_RunSuiteSafe "RepairRuntime"
    'Run deterministic live-clock timer registration and drain checks
        TST_DP_RunSuiteSafe "Timer"
'''
new = '''    'Run DP_RepairRuntime behavior checks
        TST_DP_RunSuiteSafe "RepairRuntime"
    'Run transactional lifecycle rollback / teardown / repair checks
        TST_DP_RunSuiteSafe "LifecycleTransaction"
    'Run deterministic live-clock timer registration and drain checks
        TST_DP_RunSuiteSafe "Timer"
'''
if text.count(old) != 1:
    raise SystemExit('dispatcher registration anchor mismatch')
text = text.replace(old, new, 1)

old = '''            Case "REPAIRRUNTIME"
                TST_DP_RunSuite_RepairRuntime

            Case "TIMER"
                TST_DP_RunSuite_Timer
'''
new = '''            Case "REPAIRRUNTIME"
                TST_DP_RunSuite_RepairRuntime

            Case "LIFECYCLETRANSACTION"
                TST_DP_RunSuite_LifecycleTransaction

            Case "TIMER"
                TST_DP_RunSuite_Timer
'''
if text.count(old) != 1:
    raise SystemExit('safe dispatcher anchor mismatch')
text = text.replace(old, new, 1)

# Insert before the Timer suite.
anchor = 'Private Sub TST_DP_RunSuite_Timer()\n'
if text.count(anchor) != 1:
    raise SystemExit('timer suite anchor mismatch')

suite = r'''Private Sub TST_DP_RunSuite_LifecycleTransaction()

'
'==============================================================================
'                    LIFECYCLE TRANSACTION SUITE
'------------------------------------------------------------------------------
' PURPOSE
'   Validates #50 startup rollback, teardown continuation / lease retention,
'   retry, repair, and diagnostic state under deterministic injected failures.
'
' NOTES
'   This suite owns the transaction contract rather than a feature domain.
'   It deliberately keeps Application.EnableEvents disabled except where repair
'   is being asserted, and restores the run's lease/runtime state on every exit.
'
'   Package / second-provider certification remains #63.
' UPDATED
'   2026-09-05
'==============================================================================

    Const START_ERR As Long = vbObjectError + 3250
    Const CLEAN_ERR As Long = vbObjectError + 3251
    Const REPAIR_ERR As Long = vbObjectError + 3252

    Dim StartSteps As Variant
    Dim StepIndex As Long
    Dim RaisedNumber As Long
    Dim RaisedDescription As String
    Dim OwnerToken As String
    Dim TraceText As String

    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "LifecycleTransaction"
    M_Lease_Test_SilenceRefusalReport True

'------------------------------------------------------------------------------
' FRESH-START ROLLBACK MATRIX
'------------------------------------------------------------------------------
    StartSteps = Array( _
        "Start.AfterAdmission", _
        "Start.AfterManager", _
        "Start.AfterContextMenu", _
        "Start.AfterKeyboard", _
        "Start.AfterGrid", _
        "Start.AfterRefresh")

    For StepIndex = LBound(StartSteps) To UBound(StartSteps)
        TST_DP_Lifecycle_ResetToFreeState
        Excel.Application.EnableEvents = False
        M_Lifecycle_Test_ArmFault VBA.CStr(StartSteps(StepIndex)), START_ERR + StepIndex

        RaisedNumber = 0
        RaisedDescription = VBA.vbNullString
        On Error Resume Next
        DP_Start
        RaisedNumber = Err.Number
        RaisedDescription = Err.Description
        Err.Clear
        On Error GoTo SuiteFail

        TST_DP_AssertEqualsLong "Fresh start raises injected primary failure at " & VBA.CStr(StartSteps(StepIndex)), _
            START_ERR + StepIndex, RaisedNumber
        TST_DP_AssertEqualsLong "Primary failure number survives rollback at " & VBA.CStr(StartSteps(StepIndex)), _
            START_ERR + StepIndex, M_Lifecycle_Test_LastPrimaryNumber()
        TST_DP_AssertTrue "Primary failure step is retained at " & VBA.CStr(StartSteps(StepIndex)), _
            (VBA.LenB(M_Lifecycle_Test_LastPrimaryStep()) > 0)
        TST_DP_AssertTrue "Primary failure description survives rollback at " & VBA.CStr(StartSteps(StepIndex)), _
            (VBA.LenB(M_Lifecycle_Test_LastPrimaryDescription()) > 0 And VBA.LenB(RaisedDescription) > 0)
        TST_DP_AssertEqualsString "Fresh-start diagnostic operation is DP_Start", _
            "DP_Start", M_Lifecycle_Test_LastOperation()
        TST_DP_AssertFalse "Failed fresh start is not reported successful", _
            M_Lifecycle_Test_LastSucceeded()
        TST_DP_AssertFalse "Fresh-start lease was not pre-owned", _
            M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
        TST_DP_AssertTrue "Fresh-start lease was acquired by the failing call", _
            M_Lifecycle_Test_LastLeaseAcquiredThisCall()
        TST_DP_AssertTrue "Fresh-start rollback was attempted", _
            M_Lifecycle_Test_LastCleanupAttempted()
        TST_DP_AssertTrue "Fresh-start rollback proves critical state clean", _
            M_Lifecycle_Test_LastCriticalClean()
        TST_DP_AssertTrue "Fresh-start rollback releases its newly acquired lease", _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertFalse "Fresh-start rollback leaves no owned lease", M_Lease_IsOwner()
        TST_DP_AssertEqualsString "Fresh-start rollback removes the live lease", _
            VBA.vbNullString, TST_DP_ReadLeaseOwnerForTest()
        TST_DP_AssertTrue "Fresh-start rollback releases the manager", (gDP_Manager Is Nothing)
        TST_DP_AssertFalse "Fresh-start rollback leaves no loaded picker form", _
            TST_DP_IsPickerFormLoadedForTest()
        TST_DP_AssertEqualsLong "Fresh-start rollback leaves no owned canonical grid icon", _
            0, TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
        TST_DP_AssertFalse "Fresh-start rollback preserves disabled caller events", _
            Excel.Application.EnableEvents
    Next StepIndex

'------------------------------------------------------------------------------
' PRE-OWNED START FAILURE MUST RETAIN OWNERSHIP
'------------------------------------------------------------------------------
    TST_DP_Lifecycle_ResetToFreeState
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Pre-owned setup owns the provider lease", (VBA.LenB(OwnerToken) > 0)

    M_Lifecycle_Test_ArmFault "Start.AfterManager", START_ERR + 100
    RaisedNumber = 0
    On Error Resume Next
    DP_Start
    RaisedNumber = Err.Number
    Err.Clear
    On Error GoTo SuiteFail

    TST_DP_AssertEqualsLong "Repeated start raises its injected primary failure", _
        START_ERR + 100, RaisedNumber
    TST_DP_AssertTrue "Repeated start classifies the lease as pre-owned", _
        M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
    TST_DP_AssertFalse "Repeated start does not classify lease as acquired this call", _
        M_Lifecycle_Test_LastLeaseAcquiredThisCall()
    TST_DP_AssertFalse "Repeated start does not release the pre-owned lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Repeated start preserves the exact pre-owned lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Repeated start retains provider ownership", M_Lease_IsOwner()

'------------------------------------------------------------------------------
' SHUTDOWN FAILURE CONTINUES CLEANUP, RETAINS LEASE, AND RETRIES
'------------------------------------------------------------------------------
    M_Lifecycle_Test_Reset True
    M_Lifecycle_Test_ArmFault "Cleanup.ContextMenu", CLEAN_ERR
    DP_Stop

    TraceText = M_Lifecycle_Test_LastTrace()
    TST_DP_AssertEqualsString "Stop diagnostic operation is DP_Stop", _
        "DP_Stop", M_Lifecycle_Test_LastOperation()
    TST_DP_AssertFalse "Stop with cleanup failure is not reported successful", _
        M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertTrue "Stop entered the cleanup transaction", _
        M_Lifecycle_Test_LastCleanupAttempted()
    TST_DP_AssertTrue "Stop records at least one cleanup failure", _
        (M_Lifecycle_Test_LastCleanupFailureCount() >= 1)
    TST_DP_AssertFalse "Critical cleanup failure is not reported clean", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertFalse "Critical cleanup failure does not release the lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertTrue "Critical cleanup failure retains local ownership", M_Lease_IsOwner()
    TST_DP_AssertEqualsString "Critical cleanup failure retains the exact lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Stop continues to keyboard cleanup after context-menu failure", _
        (VBA.InStr(1, TraceText, "Keyboard=", vbBinaryCompare) > 0)
    TST_DP_AssertTrue "Stop continues to grid cleanup after context-menu failure", _
        (VBA.InStr(1, TraceText, "Grid=", vbBinaryCompare) > 0)
    TST_DP_AssertTrue "Stop records retained lease after critical failure", _
        (VBA.InStr(1, TraceText, "Lease=", vbBinaryCompare) > 0)

    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Retry stop proves critical state clean", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertTrue "Retry stop releases the lease last", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertFalse "Retry stop leaves no provider ownership", M_Lease_IsOwner()
    TST_DP_AssertEqualsString "Retry stop removes the lease", _
        VBA.vbNullString, TST_DP_ReadLeaseOwnerForTest()

'------------------------------------------------------------------------------
' COMPOUND CLEANUP FAILURE RETAINS MULTIPLE FAILURES
'------------------------------------------------------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Cleanup.ContextMenu", CLEAN_ERR + 10
    M_Lifecycle_Test_ArmFault "Cleanup.Keyboard", CLEAN_ERR + 11
    DP_Stop
    TST_DP_AssertTrue "Compound stop represents multiple cleanup failures", _
        (M_Lifecycle_Test_LastCleanupFailureCount() >= 2)
    TST_DP_AssertFalse "Compound cleanup failure retains the lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Compound cleanup retains exact ownership proof", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Compound-failure retry releases after clean retry", _
        M_Lifecycle_Test_LastLeaseReleased()

'------------------------------------------------------------------------------
' #27 UNRESOLVED TIMER OUTCOME BLOCKS LEASE RELEASE
'------------------------------------------------------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Timer_Test_Reset
    M_Timer_Test_ArmFault "Cancel", CLEAN_ERR + 20
    M_Timer_Start
    M_Timer_Stop
    TST_DP_AssertTrue "Timer setup retains an unresolved registration", _
        M_Timer_Test_IsUnresolved()
    DP_Stop
    TST_DP_AssertFalse "Unresolved timer prevents critical-clean shutdown", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertFalse "Unresolved timer prevents lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Unresolved timer retains exact provider lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    M_Timer_Test_Reset
    M_Lifecycle_Test_Reset True
    DP_Stop

'------------------------------------------------------------------------------
' #53 OWNED GRID DELETION FAILURE BLOCKS LEASE RELEASE; FOREIGN SHAPE DOES NOT
'------------------------------------------------------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Cleanup.Grid", CLEAN_ERR + 30
    DP_Stop
    TST_DP_AssertFalse "Owned-grid cleanup failure is not reported clean", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertFalse "Owned-grid cleanup failure blocks lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Owned-grid cleanup failure retains exact lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_Reset True
    DP_Stop

    Set gDP_GridIconShape = Nothing
    TST_DP_AddNamedShapeForTest( _
        mTST_DP_ScratchSheet, DP_GRID_ICON_NAME, "Unrelated user shape").Visible = msoTrue
    Excel.Application.EnableEvents = False
    DP_Start
    DP_Stop
    TST_DP_AssertTrue "Foreign same-name shape does not block clean stop", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertTrue "Foreign same-name shape does not block lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsLong "Foreign same-name shape survives lifecycle cleanup", _
        1, TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Delete

'------------------------------------------------------------------------------
' REPAIR: CLEAN FIRST, RETAIN LEASE, REBUILD; FAILURE NEVER RELEASES OWNERSHIP
'------------------------------------------------------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Repair.AfterContextMenu", REPAIR_ERR
    RaisedNumber = 0
    On Error Resume Next
    DP_RepairRuntime
    RaisedNumber = Err.Number
    Err.Clear
    On Error GoTo SuiteFail

    TST_DP_AssertEqualsLong "Repair raises injected rebuild failure", REPAIR_ERR, RaisedNumber
    TST_DP_AssertEqualsString "Repair diagnostic operation is DP_RepairRuntime", _
        "DP_RepairRuntime", M_Lifecycle_Test_LastOperation()
    TST_DP_AssertFalse "Failed repair is not reported successful", _
        M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertFalse "Repair never reports lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Failed repair retains exact provider lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Failed repair retains provider ownership", M_Lease_IsOwner()
    TST_DP_AssertTrue "Repair keeps its deliberate EnableEvents=True exception", _
        Excel.Application.EnableEvents

    M_Lifecycle_Test_Reset True
    DP_RepairRuntime
    TST_DP_AssertTrue "Healthy repair reports success", M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertTrue "Healthy repair retains provider ownership", M_Lease_IsOwner()
    TST_DP_AssertEqualsString "Healthy repair retains exact provider lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Healthy repair leaves events enabled", Excel.Application.EnableEvents

'------------------------------------------------------------------------------
' NON-OWNER STOP / REPAIR FAIL CLOSED
'------------------------------------------------------------------------------
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lease_Test_ClearOwnerToken
    M_Lease_Test_SilenceRefusalReport True
    DP_Stop
    TST_DP_AssertEqualsString "Non-owner stop does not mutate foreign lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    DP_RepairRuntime
    TST_DP_AssertEqualsString "Non-owner repair does not mutate foreign lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()

SuiteExit:
    On Error Resume Next
    M_Lifecycle_Test_Reset True
    M_Timer_Test_Reset
    M_Lease_Test_SilenceRefusalReport False
    TST_DP_ForceClearLeaseForTest
    Set gDP_Manager = Nothing
    M_GridIcon_PurgeAll
    If TST_DP_ShapeExists(mTST_DP_ScratchSheet, DP_GRID_ICON_NAME) Then
        mTST_DP_ScratchSheet.Shapes(DP_GRID_ICON_NAME).Delete
    End If
    M_Lease_TryAcquire
    Excel.Application.EnableEvents = False
    Err.Clear
    On Error GoTo 0
    Exit Sub

SuiteFail:
    TST_DP_RecordFail "LifecycleTransaction suite failed", _
        "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
   