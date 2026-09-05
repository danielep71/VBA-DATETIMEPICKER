from pathlib import Path

path = Path('test/M_cDP_Test.bas')
text = path.read_text(encoding='utf-8')

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
assert text.count(old) == 1, 'dispatcher registration anchor'
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
assert text.count(old) == 1, 'safe dispatcher anchor'
text = text.replace(old, new, 1)

anchor = 'Private Sub TST_DP_RunSuite_Timer()\n'
assert text.count(anchor) == 1, 'timer suite anchor'

suite = '''Private Sub TST_DP_RunSuite_LifecycleTransaction()

'==============================================================================
' LIFECYCLE TRANSACTION SUITE (#50)
' Deterministic rollback, cleanup continuation, lease retention/retry and repair.
' UPDATED 2026-09-05
'==============================================================================
    Const START_ERR As Long = vbObjectError + 3250
    Const CLEAN_ERR As Long = vbObjectError + 3260
    Const REPAIR_ERR As Long = vbObjectError + 3270
    Dim StartSteps As Variant
    Dim I As Long
    Dim RaisedNumber As Long
    Dim OwnerToken As String
    Dim TraceText As String
    Dim ForeignShape As Excel.Shape

    On Error GoTo SuiteFail
    mTST_DP_CurrentSuite = "LifecycleTransaction"
    M_Lease_Test_SilenceRefusalReport True

'--- fresh-start rollback after each mutating boundary -------------------------
    StartSteps = Array("Start.AfterAdmission", "Start.AfterManager", _
        "Start.AfterContextMenu", "Start.AfterKeyboard", "Start.AfterGrid", _
        "Start.AfterRefresh")

    For I = LBound(StartSteps) To UBound(StartSteps)
        TST_DP_Lifecycle_ResetToFreeState
        Excel.Application.EnableEvents = False
        M_Lifecycle_Test_ArmFault VBA.CStr(StartSteps(I)), START_ERR + I
        RaisedNumber = 0
        On Error Resume Next
        DP_Start
        RaisedNumber = Err.Number
        Err.Clear
        On Error GoTo SuiteFail

        TST_DP_AssertEqualsLong "Fresh start raises at " & VBA.CStr(StartSteps(I)), _
            START_ERR + I, RaisedNumber
        TST_DP_AssertEqualsLong "Primary error survives rollback at " & VBA.CStr(StartSteps(I)), _
            START_ERR + I, M_Lifecycle_Test_LastPrimaryNumber()
        TST_DP_AssertEqualsString "Fresh-start operation is DP_Start", _
            "DP_Start", M_Lifecycle_Test_LastOperation()
        TST_DP_AssertTrue "Fresh start acquired the lease", _
            M_Lifecycle_Test_LastLeaseAcquiredThisCall()
        TST_DP_AssertTrue "Fresh-start rollback attempted cleanup", _
            M_Lifecycle_Test_LastCleanupAttempted()
        TST_DP_AssertTrue "Fresh-start rollback is critically clean", _
            M_Lifecycle_Test_LastCriticalClean()
        TST_DP_AssertTrue "Fresh-start rollback released its lease", _
            M_Lifecycle_Test_LastLeaseReleased()
        TST_DP_AssertEqualsString "Fresh-start rollback leaves no lease", _
            VBA.vbNullString, TST_DP_ReadLeaseOwnerForTest()
        TST_DP_AssertTrue "Fresh-start rollback releases manager", (gDP_Manager Is Nothing)
        TST_DP_AssertFalse "Fresh-start rollback preserves disabled events", _
            Excel.Application.EnableEvents
    Next I

'--- pre-owned start failure retains ownership --------------------------------
    TST_DP_Lifecycle_ResetToFreeState
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Start.AfterManager", START_ERR + 20
    On Error Resume Next
    DP_Start
    RaisedNumber = Err.Number
    Err.Clear
    On Error GoTo SuiteFail
    TST_DP_AssertEqualsLong "Repeated start raises injected failure", START_ERR + 20, RaisedNumber
    TST_DP_AssertTrue "Repeated start classifies pre-owned lease", _
        M_Lifecycle_Test_LastLeaseWasAlreadyOwned()
    TST_DP_AssertFalse "Repeated start did not acquire lease this call", _
        M_Lifecycle_Test_LastLeaseAcquiredThisCall()
    TST_DP_AssertEqualsString "Repeated start retains exact lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()

'--- cleanup failure continues, retains lease, and clean retry releases --------
    M_Lifecycle_Test_Reset True
    M_Lifecycle_Test_ArmFault "Cleanup.ContextMenu", CLEAN_ERR
    DP_Stop
    TraceText = M_Lifecycle_Test_LastTrace()
    TST_DP_AssertEqualsString "Stop operation is DP_Stop", "DP_Stop", _
        M_Lifecycle_Test_LastOperation()
    TST_DP_AssertTrue "Stop attempted cleanup", M_Lifecycle_Test_LastCleanupAttempted()
    TST_DP_AssertTrue "Stop records cleanup failure", _
        (M_Lifecycle_Test_LastCleanupFailureCount() >= 1)
    TST_DP_AssertFalse "Failed cleanup is not critically clean", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertFalse "Failed cleanup retains lease", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Failed cleanup retains exact lease", _
        OwnerToken, TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Cleanup continues after context-menu failure", _
        (VBA.InStr(1, TraceText, "Keyboard=", vbBinaryCompare) > 0 And _
         VBA.InStr(1, TraceText, "Grid=", vbBinaryCompare) > 0)

    M_Lifecycle_Test_Reset True
    DP_Stop
    TST_DP_AssertTrue "Retry stop becomes critically clean", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertTrue "Retry stop releases lease", M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsString "Retry stop removes lease", VBA.vbNullString, _
        TST_DP_ReadLeaseOwnerForTest()

'--- compound cleanup failure --------------------------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Cleanup.ContextMenu", CLEAN_ERR + 1
    M_Lifecycle_Test_ArmFault "Cleanup.Keyboard", CLEAN_ERR + 2
    DP_Stop
    TST_DP_AssertTrue "Compound cleanup records multiple failures", _
        (M_Lifecycle_Test_LastCleanupFailureCount() >= 2)
    TST_DP_AssertEqualsString "Compound cleanup retains lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_Reset True
    DP_Stop

'--- #27 unresolved timer prevents lease release -------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Timer_Test_Reset
    M_Timer_Test_ArmFault "Cancel", CLEAN_ERR + 3
    M_Timer_Start
    M_Timer_Stop
    TST_DP_AssertTrue "Timer setup is unresolved", M_Timer_Test_IsUnresolved()
    DP_Stop
    TST_DP_AssertFalse "Unresolved timer blocks critical-clean stop", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertEqualsString "Unresolved timer retains lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()
    M_Timer_Test_Reset
    M_Lifecycle_Test_Reset True
    DP_Stop

'--- #53 owned cleanup failure blocks; foreign same-name shape does not --------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Cleanup.Grid", CLEAN_ERR + 4
    DP_Stop
    TST_DP_AssertFalse "Owned-grid cleanup failure blocks clean stop", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertEqualsString "Owned-grid cleanup failure retains lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_Reset True
    DP_Stop

    Set gDP_GridIconShape = Nothing
    Set ForeignShape = TST_DP_AddNamedShapeForTest( _
        mTST_DP_ScratchSheet, DP_GRID_ICON_NAME, "Unrelated user shape")
    Excel.Application.EnableEvents = False
    DP_Start
    DP_Stop
    TST_DP_AssertTrue "Foreign same-name shape does not block clean stop", _
        M_Lifecycle_Test_LastCriticalClean()
    TST_DP_AssertTrue "Foreign same-name shape does not block lease release", _
        M_Lifecycle_Test_LastLeaseReleased()
    TST_DP_AssertEqualsLong "Foreign same-name shape survives stop", 1, _
        TST_DP_CountNamedShapes(mTST_DP_HostWorkbook, DP_GRID_ICON_NAME)
    ForeignShape.Delete
    Set ForeignShape = Nothing

'--- repair retains lease on failure and success -------------------------------
    Excel.Application.EnableEvents = False
    DP_Start
    OwnerToken = TST_DP_ReadLeaseOwnerForTest()
    M_Lifecycle_Test_ArmFault "Repair.AfterContextMenu", REPAIR_ERR
    On Error Resume Next
    DP_RepairRuntime
    RaisedNumber = Err.Number
    Err.Clear
    On Error GoTo SuiteFail
    TST_DP_AssertEqualsLong "Repair raises injected failure", REPAIR_ERR, RaisedNumber
    TST_DP_AssertEqualsString "Repair operation is DP_RepairRuntime", _
        "DP_RepairRuntime", M_Lifecycle_Test_LastOperation()
    TST_DP_AssertFalse "Failed repair is not successful", M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertEqualsString "Failed repair retains lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()
    TST_DP_AssertTrue "Repair keeps EnableEvents=True exception", _
        Excel.Application.EnableEvents

    M_Lifecycle_Test_Reset True
    DP_RepairRuntime
    TST_DP_AssertTrue "Healthy repair reports success", M_Lifecycle_Test_LastSucceeded()
    TST_DP_AssertEqualsString "Healthy repair retains lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()

'--- non-owner paths fail closed -----------------------------------------------
    M_Lease_Test_ClearOwnerToken
    M_Lease_Test_SilenceRefusalReport True
    DP_Stop
    TST_DP_AssertEqualsString "Non-owner stop leaves foreign lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()
    DP_RepairRuntime
    TST_DP_AssertEqualsString "Non-owner repair leaves foreign lease", OwnerToken, _
        TST_DP_ReadLeaseOwnerForTest()

SuiteExit:
    On Error Resume Next
    Set ForeignShape = Nothing
    M_Lifecycle_Test_Reset True
    M_Timer_Test_Reset
    M_Lease_Test_SilenceRefusalReport False
    TST_DP_ForceClearLeaseForTest
    Set gDP_Manager = Nothing
    M_GridIcon_PurgeAll
    M_Lease_TryAcquire
    Excel.Application.EnableEvents = False
    Err.Clear
    On Error GoTo 0
    Exit Sub

SuiteFail:
    TST_DP_RecordFail "LifecycleTransaction suite failed", _
        "Error " & VBA.CStr(Err.Number) & " - " & Err.Description
    Err.Clear
    Resume SuiteExit

End Sub

Private Sub TST_DP_Lifecycle_ResetToFreeState()
    On Error Resume Next
    M_Lifecycle_Test_Reset True
    M_Timer_Test_Reset
    M_Lease_Test_SilenceRefusalReport True
    DP_Stop
    TST_DP_ForceClearLeaseForTest
    Set gDP_Manager = Nothing
    M_GridIcon_PurgeAll
    Excel.Application.EnableEvents = False
    Err.Clear
    On Error GoTo 0
End Sub

'''

text = text.replace(anchor, suite + anchor, 1)
path.write_text(text, encoding='utf-8')
