from pathlib import Path
import re

path = Path('src/modules/M_DatePicker.bas')
text = path.read_text(encoding='utf-8')

def sub_once(pattern, repl, label, flags=0):
    global text
    text, count = re.subn(pattern, repl, text, count=1, flags=flags)
    if count != 1:
        raise SystemExit(f'{label}: expected 1 match, found {count}')

def patch_proc(name, patcher):
    global text
    pat = re.compile(rf'(?ms)^Public Sub {re.escape(name)}\(.*?^End Sub\s*\n')
    m = pat.search(text)
    if not m:
        raise SystemExit(f'{name}: procedure not found')
    block = m.group(0)
    new = patcher(block)
    text = text[:m.start()] + new + text[m.end():]

# Diagnostic state fields.
sub_once(
    r"(^\s*Private mDP_LifecycleLastPrimaryDescription As String[^\n]*\n)",
    r"\1    Private mDP_LifecycleLastOperation      As String           'Lifecycle entry point represented by the current diagnostic state\n"
    r"    Private mDP_LifecycleLastSucceeded      As Boolean          'True only when the lifecycle entry point completed its contract\n"
    r"    Private mDP_LifecycleLastCleanupAttempted As Boolean        'True once the lifecycle cleanup transaction was entered\n"
    r"    Private mDP_LifecycleLastCleanupFailureCount As Long        'Number of cleanup steps that actually failed\n"
    r"    Private mDP_LifecycleLastLeaseWasAlreadyOwned As Boolean    'True when the provider lease pre-dated the lifecycle call\n"
    r"    Private mDP_LifecycleLastLeaseAcquiredThisCall As Boolean   'True when the lifecycle call acquired the provider lease\n",
    'diagnostic state fields', re.M)

# Reset all observable lifecycle state.
sub_once(
    r"(mDP_LifecycleLastPrimaryDescription = VBA\.vbNullString\s*\n)",
    r"\1    mDP_LifecycleLastOperation = VBA.vbNullString\n"
    r"    mDP_LifecycleLastSucceeded = False\n"
    r"    mDP_LifecycleLastCleanupAttempted = False\n"
    r"    mDP_LifecycleLastCleanupFailureCount = 0\n"
    r"    mDP_LifecycleLastLeaseWasAlreadyOwned = False\n"
    r"    mDP_LifecycleLastLeaseAcquiredThisCall = False\n",
    'reset diagnostics')

# Cleanup recorder: count actual failed attempts.
sub_once(
    r"(Private Sub M_Lifecycle_RecordCleanupStep\( _\s*\n\s*ByVal StepName As String, _\s*\n\s*ByVal Succeeded As Boolean, _\s*\n\s*ByVal ErrorNumber As Long, _\s*\n\s*ByVal ErrorDescription As String)\)",
    r"\1, _\n    Optional ByVal CountFailure As Boolean = True)",
    'cleanup recorder signature')
sub_once(
    r"(If Not Succeeded Then\s*\n)",
    r"\1        If CountFailure Then\n"
    r"            mDP_LifecycleLastCleanupFailureCount = _\n"
    r"                mDP_LifecycleLastCleanupFailureCount + 1\n"
    r"        End If\n",
    'cleanup failure count')

# Query seams appended after the already-existing primary-description seam.
query_block = '''Public Function M_Lifecycle_Test_LastOperation() As String

    M_Lifecycle_Test_LastOperation = mDP_LifecycleLastOperation

End Function

Public Function M_Lifecycle_Test_LastSucceeded() As Boolean

    M_Lifecycle_Test_LastSucceeded = mDP_LifecycleLastSucceeded

End Function

Public Function M_Lifecycle_Test_LastCleanupAttempted() As Boolean

    M_Lifecycle_Test_LastCleanupAttempted = mDP_LifecycleLastCleanupAttempted

End Function

Public Function M_Lifecycle_Test_LastCleanupFailureCount() As Long

    M_Lifecycle_Test_LastCleanupFailureCount = mDP_LifecycleLastCleanupFailureCount

End Function

Public Function M_Lifecycle_Test_LastLeaseWasAlreadyOwned() As Boolean

    M_Lifecycle_Test_LastLeaseWasAlreadyOwned = mDP_LifecycleLastLeaseWasAlreadyOwned

End Function

Public Function M_Lifecycle_Test_LastLeaseAcquiredThisCall() As Boolean

    M_Lifecycle_Test_LastLeaseAcquiredThisCall = mDP_LifecycleLastLeaseAcquiredThisCall

End Function

'''
sub_once(
    r"(Public Function M_Lifecycle_Test_LastPrimaryDescription\(\) As String\s*\n\s*\n\s*M_Lifecycle_Test_LastPrimaryDescription = mDP_LifecycleLastPrimaryDescription\s*\n\s*\nEnd Function\s*\n\s*\n)",
    lambda m: m.group(1) + query_block,
    'query seams')

# Scope cleanup bookkeeping to M_Lifecycle_Cleanup.
func_pat = re.compile(r'(?ms)^Private Function M_Lifecycle_Cleanup\(.*?^End Function\s*\n')
m = func_pat.search(text)
if not m:
    raise SystemExit('M_Lifecycle_Cleanup not found')
cleanup = m.group(0)
cleanup, c = re.subn(
    r"(mDP_LifecycleLastLeaseReleased = False\s*\n)",
    r"\1    mDP_LifecycleLastCleanupAttempted = True\n"
    r"    mDP_LifecycleLastCleanupFailureCount = 0\n",
    cleanup, count=1)
if c != 1:
    raise SystemExit(f'cleanup bookkeeping: {c}')
cleanup, c = re.subn(
    r'(M_Lifecycle_RecordCleanupStep "Lease", False, _\s*\n\s*vbObjectError \+ 2720, "Lease retained because critical cleanup is incomplete")',
    r'\1, _\n      CountFailure:=False',
    cleanup, count=1)
if c != 1:
    raise SystemExit(f'withheld lease: {c}')
text = text[:m.start()] + cleanup + text[m.end():]

# DP_Start diagnostics and classification.
def patch_start(block):
    block, c = re.subn(r'(M_Lifecycle_ResetObservation\s*\n)', r'\1    mDP_LifecycleLastOperation = PROC_NAME\n', block, count=1)
    if c != 1: raise SystemExit(f'DP_Start operation: {c}')
    block, c = re.subn(r'(PreOwned = M_Lease_IsOwner\(\)\s*\n)', r'\1    mDP_LifecycleLastLeaseWasAlreadyOwned = PreOwned\n', block, count=1)
    if c != 1: raise SystemExit(f'DP_Start preowned: {c}')
    block, c = re.subn(r'(AcquiredThisCall = \(Not PreOwned And M_Lease_IsOwner\(\)\)\s*\n)', r'\1    mDP_LifecycleLastLeaseAcquiredThisCall = AcquiredThisCall\n', block, count=1)
    if c != 1: raise SystemExit(f'DP_Start acquired: {c}')
    block, c = re.subn(r'(If HasCallerEnableEvents Then Excel\.Application\.EnableEvents = CallerEnableEvents\s*\n\s*Exit Sub)', r'If HasCallerEnableEvents Then Excel.Application.EnableEvents = CallerEnableEvents\n    mDP_LifecycleLastSucceeded = M_Lease_IsOwner()\n    Exit Sub', block, count=1)
    if c != 1: raise SystemExit(f'DP_Start success: {c}')
    block, c = re.subn(r'(M_Lifecycle_SetPrimaryFailure ErrorNumber, HandlerStep, ErrorDescription\s*\n)', r'\1    mDP_LifecycleLastSucceeded = False\n', block, count=1)
    if c != 1: raise SystemExit(f'DP_Start failure: {c}')
    return block
patch_proc('DP_Start', patch_start)

# DP_Stop uses the common cleanup transaction and records ownership classification.
def replace_stop(_block):
    return '''Public Sub DP_Stop()

'------------------------------------------------------------------------------
'                           STOP DATEPICKER
'------------------------------------------------------------------------------
' PURPOSE
'   Tears down DatePicker-owned runtime state and releases the provider lease only
'   after every critical cleanup boundary is proven clean
'
' ERROR POLICY
'   Best-effort outward behavior is preserved. Incomplete cleanup is diagnostic,
'   retains ownership, and is retryable by a later DP_Stop call
'
' UPDATED
'   2026-09-05
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
        M_Lease_ReportRefusal "DP_Stop"
        mDP_LifecycleLastSucceeded = False
        GoTo CleanExit
    End If

    mDP_LifecycleLastSucceeded = M_Lifecycle_Cleanup( _
        True, "DP_Stop", HasCallerEnableEvents, CallerEnableEvents)

CleanExit:
    Err.Clear
    On Error GoTo 0

End Sub
'''
patch_proc('DP_Stop', replace_stop)

# DP_RepairRuntime diagnostics; repair remains the EnableEvents=True exception.
def patch_repair(block):
    block, c = re.subn(r'(M_Lifecycle_ResetObservation\s*\n)', r'\1    mDP_LifecycleLastOperation = PROC_NAME\n    mDP_LifecycleLastLeaseWasAlreadyOwned = M_Lease_IsOwner()\n    mDP_LifecycleLastLeaseAcquiredThisCall = False\n', block, count=1)
    if c != 1: raise SystemExit(f'Repair operation: {c}')
    block, c = re.subn(r'(M_Lifecycle_RaiseIfFault "Repair\.AfterRefresh"\s*\n)', r'\1\n    mDP_LifecycleLastSucceeded = True\n', block, count=1)
    if c != 1: raise SystemExit(f'Repair success: {c}')
    block, c = re.subn(r'(M_Lifecycle_SetPrimaryFailure ErrorNumber, HandlerStep, ErrorDescription\s*\n)', r'\1    mDP_LifecycleLastSucceeded = False\n', block, count=1)
    if c != 1: raise SystemExit(f'Repair failure: {c}')
    return block
patch_proc('DP_RepairRuntime', patch_repair)

path.write_text(text, encoding='utf-8')

required = [
    'M_Lifecycle_Test_LastOperation',
    'M_Lifecycle_Test_LastSucceeded',
    'M_Lifecycle_Test_LastCleanupAttempted',
    'M_Lifecycle_Test_LastCleanupFailureCount',
    'M_Lifecycle_Test_LastLeaseWasAlreadyOwned',
    'M_Lifecycle_Test_LastLeaseAcquiredThisCall',
    'CountFailure:=False',
]
missing = [x for x in required if x not in text]
if missing:
    raise SystemExit(f'missing diagnostics: {missing}')
if text.count('Public Sub DP_Stop()') != 1:
    raise SystemExit('DP_Stop count invalid')
