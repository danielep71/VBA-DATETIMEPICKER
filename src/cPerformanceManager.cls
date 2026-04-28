VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPerformanceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CLASS: cPerformanceManager
'------------------------------------------------------------------------------
' PURPOSE
'   High-precision timing and benchmark-support utility for VBA on Windows.
'
'   The class provides:
'     - multiple timing backends under a session-bound interface
'     - low-overhead numeric elapsed-time measurement
'     - human-readable elapsed-time diagnostics
'     - optional formatting of an already measured elapsed-seconds value
'     - optional shared Excel "time-waster" suppression for benchmark runs and
'       for general Excel/VBA performance improvement in heavy procedures
'     - structured named checkpoints with numeric delta/cumulative reporting
'
' WHY THIS EXISTS
'   - VBA Timer() has limited resolution and rolls over at midnight.
'   - Windows provides monotonic and high-resolution timing sources that are
'     better suited to instrumentation and benchmarking.
'   - A single, consistent API reduces timing mistakes, duplicated code, and
'     backend-specific handling across modules.
'   - Excel benchmark runs and heavy workbook automation often benefit from
'     temporarily suppressing expensive Application behaviors such as
'     ScreenUpdating, EnableEvents, DisplayAlerts, Calculation, and Cursor
'     changes.
'
' COMPATIBILITY
'   - Windows only for API-backed timing methods (2..5).
'   - On non-Windows hosts, only Timer()/Now()-based methods are conceptually
'     portable, if retained.
'   - Requires appropriate VBA7 / Win64 conditional compilation for 32-bit and
'     64-bit Office.
'
' TIMER METHODS (iMethod)
'   1 = Timer
'       Seconds since midnight; rolls over every 24 hours.
'
'   2 = GetTickCount / GetTickCount64
'       Milliseconds since boot; the 32-bit variant wraps at approximately
'       49.7 days.
'
'   3 = timeGetTime
'       Millisecond counter with 32-bit rollover semantics; can benefit from
'       a 1ms multimedia timer resolution request.
'
'   4 = timeGetSystemTime (MMTIME / TIME_MS)
'       Millisecond counter source treated with 32-bit millisecond-counter
'       rollover semantics.
'
'   5 = QueryPerformanceCounter (QPC)   [DEFAULT / RECOMMENDED]
'       High-resolution monotonic counter for benchmark-grade timing.
'
'   6 = Now() * 86400
'       Wall-clock seconds; useful for diagnostics, not recommended as the
'       primary monotonic benchmark source.
'
' PERFORMANCE GUIDANCE
'   - Prefer ElapsedSeconds() for low-overhead numeric measurement.
'   - Prefer method 5 (QPC) for the primary benchmark path.
'   - ElapsedTime() is presentation-oriented:
'       * when called without ElapsedSecondsIn, it measures and formats
'       * when called with ElapsedSecondsIn, it formats an existing value
'         without taking a second timing sample
'   - AlignToNextTick is benchmark-specific and should be used only when its
'     extra polling overhead is justified.
'
' IMPLEMENTATION NOTES
'   - Timing is session-bound:
'       StartTimer establishes the active backend and elapsed-time reads are
'       validated against that same backend.
'   - QPC ticks and frequency are stored as Currency to preserve stable tick
'     precision while keeping the implementation efficient in VBA.
'   - 32-bit millisecond-counter sources require rollover-safe elapsed-time
'     arithmetic.
'   - timeBeginPeriod(1) for method 3 affects system-wide timer resolution and
'     can increase power usage; it is enabled only when needed and released
'     through ResetEnvironment, with Class_Terminate as a fallback safety net.
'   - Excel "time-waster" suppression is managed through a shared scope model so
'     overlapping instances can coexist without corrupting Application state.
'
' UPDATED
'   2026-04-18
'
' AUTHOR
'   Daniele Penza
'
' VERSION: 1.1.0
'==============================================================================

'------------------------------------------------------------------------------
' CLASS SETTINGS
'------------------------------------------------------------------------------
    Option Explicit     'Force explicit declaration of all variables

'------------------------------------------------------------------------------
' ENUMS
'------------------------------------------------------------------------------
    'Bitmask flags used to exempt specific Excel Application settings from
    'shared TW suppression.
    '
    'Example:
    '   Except:=ScreenUpdating Or EnableEvents
    Public Enum TW_Enum
        None = 0
        ScreenUpdating = 1
        EnableEvents = 2
        DisplayAlerts = 4
        Calculation = 8
        Cursor = 16
    End Enum

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    'SMPTE is retained only to preserve the expected MMTIME structure layout.
    Private Type SMPTE
        Hour        As Byte
        Min         As Byte
        Sec         As Byte
        Frame       As Byte
        Fps         As Byte
        Dummy       As Byte
        Pad(2)      As Byte
    End Type

    'MMTIME is used by timeGetSystemTime.
    'The class configures wType = 1 (TIME_MS), so Units is interpreted as
    'milliseconds.
    Private Type MMTIME
        wType       As Long     'Set to 1 to interpret Units as milliseconds
        Units       As Long     'Milliseconds when wType = 1
        smpteVal    As SMPTE    'Unused here; retained for structure layout
        songPtrPos  As Long
    End Type

    'Structured checkpoint row captured within one active timing session
    Private Type tCheckpointRow
        Seq                     As Long      'Checkpoint sequence number
        NameText                As String    'Checkpoint label
        NoteText                As String    'Optional checkpoint note
        DeltaSecondsValue       As Double    'Elapsed seconds since previous checkpoint
        ElapsedSecondsValue     As Double    'Cumulative elapsed seconds since StartTimer
        MethodID                As Integer   'Active method identifier
        MethodNameText          As String    'Active method descriptive name
        RunLabelText            As String    'Optional run label
    End Type

'
'------------------------------------------------------------------------------
'
'                               API DECLARATIONS
'
'------------------------------------------------------------------------------
'
    'Win32 / Win64 API declarations used for:
    '  - timing backends
    '  - timer-resolution management
    '  - coarse sleep / scheduler diagnostics
    '
    'Conditional compilation handles:
    '  - VBA7 vs pre-VBA7
    '  - Win64 vs Win32

    '---------------------------------------------------------------------------
    ' HIGH-RESOLUTION TIMING (KERNEL32)
    '   QueryPerformanceFrequency / QueryPerformanceCounter
    '---------------------------------------------------------------------------
        'QPC provides a high-resolution monotonic counter:
        ' - QueryPerformanceFrequency returns ticks per second
        ' - QueryPerformanceCounter returns the current tick value
        '
        'Currency is used as efficient fixed-point 64-bit storage in VBA and is
        'a common, stable representation for QPC values in VBA implementations.

        #If VBA7 Then
            Private Declare PtrSafe Function QueryPerformanceFrequency Lib _
                "kernel32" (ByRef Frequency As Currency) As Long

            Private Declare PtrSafe Function QueryPerformanceCounter Lib _
                "kernel32" (ByRef Counter As Currency) As Long
        #Else
            Private Declare Function QueryPerformanceFrequency Lib _
                "kernel32" (ByRef Frequency As Currency) As Long

            Private Declare Function QueryPerformanceCounter Lib "kernel32" _
                (ByRef Counter As Currency) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MILLISECOND UPTIME COUNTERS (KERNEL32)
    '   GetTickCount / GetTickCount64
    '---------------------------------------------------------------------------
        'System uptime counter in milliseconds since boot:
        ' - Win64: GetTickCount64 is effectively non-rolling for practical use
        ' - Win32: GetTickCount has 32-bit rollover semantics

        #If VBA7 Then
            #If Win64 Then
                Private Declare PtrSafe Function GetTickCount64 Lib _
                    "kernel32" () As LongLong
            #Else
                Private Declare PtrSafe Function GetTickCount Lib "kernel32" _
                    () As Long
            #End If
        #Else
            Private Declare Function GetTickCount Lib "kernel32" () As Long
        #End If

    '---------------------------------------------------------------------------
    ' SLEEP / SYSTEM CLOCK INFO (KERNEL32)
    '   Sleep / GetSystemTimeAdjustment
    '---------------------------------------------------------------------------
        'Sleep suspends the current thread for the requested number of
        'milliseconds and is used for pause/coarse-wait behavior.

        #If VBA7 Then
            Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal _
                dwMilliseconds As Long)
        #Else
            Private Declare Sub Sleep Lib "kernel32" (ByVal _
                dwMilliseconds As Long)
        #End If

        'GetSystemTimeAdjustment exposes the nominal system clock increment.
        'lpTimeIncrement is used for coarse tick-interval diagnostics.

        #If VBA7 Then
            Private Declare PtrSafe Function GetSystemTimeAdjustment Lib _
                "kernel32" (lpTimeAdjustment As Long, lpTimeIncrement As Long, _
                lpTimeAdjustmentDisabled As Long) As Long
        #Else
            Private Declare Function GetSystemTimeAdjustment Lib "kernel32" _
                (lpTimeAdjustment As Long, lpTimeIncrement As Long, _
                lpTimeAdjustmentDisabled As Long) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MULTIMEDIA TIMERS (WINMM)
    '   timeGetTime / timeBeginPeriod / timeEndPeriod
    '---------------------------------------------------------------------------
        'Multimedia timer APIs:
        ' - timeGetTime       : millisecond counter with 32-bit rollover semantics
        ' - timeBeginPeriod(1): requests 1ms timer resolution (global effect)
        ' - timeEndPeriod(1)  : releases the requested timer resolution

        #If VBA7 Then
            Private Declare PtrSafe Function timeGetTime Lib "winmm.dll" () _
                As Long

            Private Declare PtrSafe Function timeBeginPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Long) As Long

            Private Declare PtrSafe Function timeEndPeriod Lib "winmm.dll" _
                (ByVal uPeriod As Long) As Long
        #Else
            Private Declare Function timeGetTime Lib "winmm.dll" () As Long

            Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Long) As Long

            Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal _
                uPeriod As Long) As Long
        #End If

    '---------------------------------------------------------------------------
    ' MULTIMEDIA SYSTEM TIME (WINMM)
    '   timeGetSystemTime using MMTIME
    '---------------------------------------------------------------------------
        'timeGetSystemTime returns its value through an MMTIME structure.
        'The class configures wType = 1 (TIME_MS), so Units is interpreted as
        'milliseconds.

        #If VBA7 Then
            Private Declare PtrSafe Function timeGetSystemTime Lib _
                "winmm.dll" (lpTime As MMTIME, ByVal uSize As Long) As Long
        #Else
            Private Declare Function timeGetSystemTime Lib "winmm.dll" _
                (lpTime As MMTIME, ByVal uSize As Long) As Long
        #End If

'------------------------------------------------------------------------------
' STATE
'------------------------------------------------------------------------------
    'Internal class state for:
    '  - active timing sessions and elapsed-time computation
    '  - QPC and non-QPC timing infrastructure
    '  - shared TW participation for this instance

    '--------------------------------------------------------------------------
    ' METHOD NAMES (1..6)
    '--------------------------------------------------------------------------
        Private m_MethodName(1 To 6) As String   'Human-readable timing backend labels

    '--------------------------------------------------------------------------
    ' GENERAL TIMING STATE (METHODS 1 / 2 / 3 / 4 / 6)
    '--------------------------------------------------------------------------
        Private m_T1    As Double    'Start timestamp for the active session
        Private m_T2    As Double    'End timestamp captured during elapsed-time evaluation
        Private m_ET    As Double    'Last computed elapsed time in seconds

    '--------------------------------------------------------------------------
    ' ACTIVE TIMER SESSION STATE
    '--------------------------------------------------------------------------
        Private m_ActiveMethodID  As Integer   'Backend currently bound to the active session
        Private m_HasActiveSession  As Boolean   'TRUE once StartTimer has established a session
        Private m_QpcAvailable  As Boolean   'TRUE when QPC is available on this machine

        'Strict mode:
        ' - TRUE  => raise on invalid / mismatched timer usage
        ' - FALSE => coerce / fallback where allowed
        Private m_StrictMode    As Boolean

    '--------------------------------------------------------------------------
    ' QPC TIMING STATE (METHOD 5)
    '--------------------------------------------------------------------------
        'QPC values are stored as Currency to preserve stable tick precision in
        'VBA while keeping the implementation efficient.
        Private m_qpcStart  As Currency      'QPC start tick
        Private m_qpcEnd    As Currency      'QPC end tick
        Private m_qpcFreq   As Currency      'QPC frequency (ticks per second)

    '--------------------------------------------------------------------------
    ' MMTIME INSTANCE FOR timeGetSystemTime
    '--------------------------------------------------------------------------
        'Reusable structure configured once and reused for method 4.
        Private m_Mmt       As MMTIME

    '--------------------------------------------------------------------------
    ' timeBeginPeriod(1) LIFECYCLE GUARD (METHOD 3)
    '--------------------------------------------------------------------------
        'Tracks whether this instance successfully requested 1ms multimedia timer
        'resolution and therefore must later balance it with timeEndPeriod(1).
        Private m_TimePeriodOn As Boolean

    '--------------------------------------------------------------------------
    ' SHARED "TIME WASTER" SESSION STATE
    '--------------------------------------------------------------------------
        Private m_TW_Active As Boolean   'TRUE when this instance is registered in shared TW suppression
        Private m_Except    As TW_Enum   'Current exemption mask requested by this instance
    
    '--------------------------------------------------------------------------
    ' CHECKPOINT / STRUCTURED REPORT STATE
    '--------------------------------------------------------------------------
        Private m_RunLabel              As String                'Optional run label for the current timing session
        Private m_CheckpointCount       As Long                  'Number of captured checkpoints
        Private m_LastCheckpointS       As Double                'Last cumulative checkpoint elapsed seconds
        Private m_Checkpoints()         As tCheckpointRow        'Checkpoint storage array

'
'------------------------------------------------------------------------------
'
'                                   LIFECYCLE
'
'------------------------------------------------------------------------------
'

Private Sub Class_Initialize()
'
'==============================================================================
'                             CLASS INITIALIZE
'------------------------------------------------------------------------------
' PURPOSE
'   Initializes the performance manager class state:
'     - Captures and validates the QPC frequency once.
'     - Initializes the MMTIME structure for timeGetSystemTime.
'     - Populates the method name table.
'     - Resets lifecycle/session flags deterministically.
'
' WHY THIS EXISTS
'   This routine establishes a known, reproducible baseline for each new
'   cPerformanceManager instance so that timing sessions and TW lifecycle
'   handling always begin from a clean state.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Queries QueryPerformanceFrequency once and records availability.
'   - Configures MMTIME to return milliseconds.
'   - Loads the timer method labels used by MethodName().
'   - Resets timing-session state.
'   - Resets TW lifecycle flags for this instance.
'   - Initializes structured checkpoint / report state.
'
' ERROR POLICY
'   Does not raise an error if QPC is unavailable.
'   Instead, records availability in m_QpcAvailable for later validation.
'
' DEPENDENCIES
'   - QueryPerformanceFrequency
'   - MMTIME
'   - TW_Enum
'
' NOTES
'   None.
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC      As Long         'Return code from QueryPerformanceFrequency

'------------------------------------------------------------------------------
' INITIALIZE (HIGH-RES COUNTER + TIMEBASE CONFIG)
'------------------------------------------------------------------------------
    'Query the high-resolution performance counter frequency once
        RC = QueryPerformanceFrequency(m_qpcFreq)
    'Record whether QPC is available and usable
        m_QpcAvailable = (RC <> 0 And m_qpcFreq <> 0)
    'Configure timeGetSystemTime's MMTIME format to milliseconds
        m_Mmt.wType = 1

'------------------------------------------------------------------------------
' INITIALIZE (METHOD LABELS)
'------------------------------------------------------------------------------
    'Method 1: VBA Timer()
        m_MethodName(1) = "Timer"
    'Method 2: kernel32 GetTickCount / GetTickCount64
        m_MethodName(2) = "GetTickCount / GetTickCount64"
    'Method 3: winmm timeGetTime
        m_MethodName(3) = "timeGetTime"
    'Method 4: winmm timeGetSystemTime
        m_MethodName(4) = "timeGetSystemTime"
    'Method 5: kernel32 QueryPerformanceCounter
        m_MethodName(5) = "QPC"
    'Method 6: VBA Now()
        m_MethodName(6) = "Now()"

'------------------------------------------------------------------------------
' INITIALIZE (STATE FLAGS / DEFAULTS)
'------------------------------------------------------------------------------
    'Guard for timeBeginPeriod/timeEndPeriod lifecycle
        m_TimePeriodOn = False
    'Default exemption mask for TW_Turn_OFF
        m_Except = TW_Enum.None
    'Initialize active timer session state
        m_ActiveMethodID = 0
    'Mark that no timer session is active yet
        m_HasActiveSession = False
    'Default strict-mode policy
        m_StrictMode = True
    'Clear cached generic timing values deterministically
        m_T1 = 0#
        m_T2 = 0#
        m_ET = 0#
    'Clear cached QPC timing values deterministically
        m_qpcStart = 0@
        m_qpcEnd = 0@
    'Indicates whether this instance currently owns an active shared TW session
        m_TW_Active = False
'------------------------------------------------------------------------------
' INITIALIZE (CHECKPOINT / REPORT STATE)
'------------------------------------------------------------------------------
    'Clear the optional run label
        m_RunLabel = vbNullString
    'Initialize the checkpoint counter
        m_CheckpointCount = 0
    'Initialize the last checkpoint elapsed-seconds marker
        m_LastCheckpointS = 0#
    'Initialize checkpoint storage with a small first capacity
        ReDim m_Checkpoints(1 To 1)
        
End Sub

Private Sub Class_Terminate()
'
'==============================================================================
'                              CLASS TERMINATE
'------------------------------------------------------------------------------
' PURPOSE
'   Final safety-net cleanup for environment changes made by this instance.
'
' WHY THIS EXISTS
'   Class_Terminate is the last opportunity for this instance to release
'   environment changes if the caller forgot to invoke ResetEnvironment
'   explicitly.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   Calls ResetEnvironment as a best-effort cleanup step during object teardown.
'
' ERROR POLICY
'   Suppresses teardown errors to avoid destabilizing object termination.
'
' DEPENDENCIES
'   - ResetEnvironment
'
' NOTES
'   ResetEnvironment remains the primary cleanup contract.
'   Class_Terminate is only the final fallback.
'
' UPDATED
'   2026-03-28
'==============================================================================

    'Destructor cleanup must be best-effort and non-fatal
        On Error Resume Next
    'Use the explicit cleanup routine as the termination safety net
        ResetEnvironment
    'Return to normal error handling semantics
        On Error GoTo 0
End Sub

'
'------------------------------------------------------------------------------
'
'                           PUBLIC: CORE TIMING API
'
'------------------------------------------------------------------------------
'


Public Sub StartTimer(Optional ByVal iMethod As Integer = 5, _
                      Optional ByVal AlignToNextTick As Boolean = False)
'
'==============================================================================
'                                  START TIMER
'------------------------------------------------------------------------------
' PURPOSE
'   Starts a timer using the selected backend and binds that backend to the
'   current timing session.
'
' WHY THIS EXISTS
'   Timing must begin from a well-defined backend and that backend must remain
'   bound to the current session so that later elapsed-time reads are validated
'   against the same method.
'
' INPUTS
'   iMethod (optional)
'     Requested timing backend.
'
'   AlignToNextTick (optional)
'     TRUE  => wait for the next observable tick before capturing the start
'     FALSE => capture the start timestamp immediately
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Normalizes and validates the requested method.
'   - Binds the validated method to the current timing session.
'   - Marks the session as started.
'   - Clears cached end/elapsed state from any prior session.
'   - Starts a fresh structured checkpoint session for the new timing run.
'   - Captures the start timestamp using the selected backend.
'   - Handles QPC separately because it uses Currency-based tick storage.
'
' ERROR POLICY
'   Validation behavior is delegated to Method_NormalizeStart.
'   If strict mode is enabled, invalid/unavailable methods may raise there.
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - QPC_Get_NextTimestampTick
'   - QueryPerformanceCounter
'   - Get_NextTimestampSeconds
'   - Get_CurrentTimestampSeconds
'
' NOTES
'   - AlignToNextTick is primarily useful for certain benchmark/diagnostic
'     scenarios and is intentionally FALSE by default.
'   - Method 5 (QPC) is handled on a dedicated path because its tick values are
'     stored in Currency fields rather than Double-based wall/tick seconds.
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' NORMALIZE INPUTS
'------------------------------------------------------------------------------
    'Normalize and validate the requested method
        iMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' INITIALIZE SESSION STATE
'------------------------------------------------------------------------------
    'Bind the selected method to the active timing session
        m_ActiveMethodID = iMethod
    'Mark the timer session as started
        m_HasActiveSession = True
    'Clear cached end/elapsed values for this new session
        m_T2 = 0#
        m_ET = 0#
        m_qpcEnd = 0@
    'Start a fresh checkpoint session for the new timing run
        Checkpoint_ClearState
        
'------------------------------------------------------------------------------
' METHOD 5: QPC (HIGH RESOLUTION)
'------------------------------------------------------------------------------
    'QPC is handled separately because it uses Currency fields for tick precision
        If iMethod = 5 Then
            'Optionally align to the next observable QPC tick
                If AlignToNextTick Then
                    m_qpcStart = QPC_Get_NextTimestampTick
                Else
                    QueryPerformanceCounter m_qpcStart
                End If
            'Mirror the raw start tick into the generic start slot
                m_T1 = CDbl(m_qpcStart)
            Exit Sub
        End If

'------------------------------------------------------------------------------
' METHODS 1/2/3/4/6: DOUBLE-BASED CLOCKS
'------------------------------------------------------------------------------
    'Optionally align to the next tick; otherwise sample immediately
        If AlignToNextTick Then
            m_T1 = Get_NextTimestampSeconds(iMethod)
        Else
            m_T1 = Get_CurrentTimestampSeconds(iMethod)
        End If
End Sub

Public Property Get ElapsedSeconds(Optional ByVal iMethod As Integer = 0) As Double
'
'==============================================================================
'                              ELAPSED SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns low-overhead numeric elapsed time in seconds, bound to the active
'   timer session started by StartTimer().
'
' WHY THIS EXISTS
'   This property is the core elapsed-time reader for the class. It resolves the
'   backend against the active session and computes elapsed seconds using the
'   correct timing source and rollover logic.
'
' INPUTS
'   iMethod (optional)
'     0   => use the active session method
'     1-6 => explicit method to validate against the active session
'
' RETURNS
'   Double
'     Elapsed time in seconds.
'
' BEHAVIOR
'   - Resolves the requested method against the active timing session.
'   - Dispatches to the correct backend-specific elapsed-time path.
'   - Applies rollover correction where needed.
'   - Stores the computed elapsed value in m_ET and returns it.
'
' ERROR POLICY
'   Method/session validation behavior is delegated to Method_ResolveElapsed.
'   A defensive fallback is included for impossible resolved-method states.
'
' DEPENDENCIES
'   - Method_ResolveElapsed
'   - Get_TickCountSeconds
'   - RolloverSeconds
'   - EnsureTimerResolution1ms
'   - timeGetTime
'   - Get_SystemTimeMs
'   - QueryPerformanceCounter
'   - Timer
'   - Now
'
' NOTES
'   - Method 5 uses Currency-based tick storage; the ratio
'       (m_qpcEnd - m_qpcStart) / m_qpcFreq
'     is intentional and valid because the Currency scaling cancels.
'   - This property is designed to be numeric and low-overhead.
'     Human-readable formatting belongs in ElapsedTime().
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ResolvedMethod  As Integer      'Validated method bound to the active session

'------------------------------------------------------------------------------
' RESOLVE METHOD
'------------------------------------------------------------------------------
    'Resolve and validate the method against the active session
        ResolvedMethod = Method_ResolveElapsed(iMethod)

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case ResolvedMethod

        '----------------------------------------------------------------------
        ' METHOD 1: VBA TIMER
        '----------------------------------------------------------------------
        Case 1
            'Capture end timestamp
                m_T2 = Timer
            'Compute elapsed
                m_ET = m_T2 - m_T1
            'Correct midnight rollover
                If m_ET < 0# Then m_ET = m_ET + 86400#
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' METHOD 2: GetTickCount / GetTickCount64
        '----------------------------------------------------------------------
        Case 2
            'Capture end timestamp in seconds
                m_T2 = Get_TickCountSeconds()
            'Compute elapsed
                m_ET = m_T2 - m_T1
            'Correct rollover only when applicable
                If m_ET < 0# Then m_ET = m_ET + RolloverSeconds(2)
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' METHOD 3: timeGetTime
        '----------------------------------------------------------------------
        Case 3
            'Ensure 1ms timer resolution is enabled.
                EnsureTimerResolution1ms
            'Capture end timestamp in seconds
                m_T2 = timeGetTime() / 1000#
            'Compute elapsed
                m_ET = m_T2 - m_T1
            'Correct 32-bit millisecond rollover
                If m_ET < 0# Then m_ET = m_ET + RolloverSeconds(3)
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' METHOD 4: timeGetSystemTime
        '----------------------------------------------------------------------
        Case 4
            'Capture end timestamp in seconds
                m_T2 = Get_SystemTimeMs() / 1000#
            'Compute elapsed.
                m_ET = m_T2 - m_T1
            'Correct 32-bit millisecond rollover
                If m_ET < 0# Then m_ET = m_ET + RolloverSeconds(4)
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' METHOD 5: QueryPerformanceCounter
        '----------------------------------------------------------------------
        Case 5
            'Capture end tick
                QueryPerformanceCounter m_qpcEnd
            'Compute elapsed seconds from ticks/frequency
                m_ET = CDbl((m_qpcEnd - m_qpcStart) / m_qpcFreq)
            'Mirror the raw end tick into the generic end slot
                m_T2 = CDbl(m_qpcEnd)
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' METHOD 6: Now()
        '----------------------------------------------------------------------
        Case 6
            'Capture end timestamp in seconds
                m_T2 = Now() * 86400#
            'Compute elapsed
                m_ET = m_T2 - m_T1
            'Return elapsed seconds
                ElapsedSeconds = m_ET

        '----------------------------------------------------------------------
        ' DEFENSIVE FALLBACK
        '----------------------------------------------------------------------
        Case Else
            'Defensive fallback for impossible/invalid resolved methods
                If m_StrictMode Then
                    Err.Raise vbObjectError + 1005, _
                              "cPerformanceManager.ElapsedSeconds", _
                              "Resolved timer method is invalid."
                Else
                    m_ET = 0#
                    ElapsedSeconds = 0#
                End If

    End Select
End Property

Public Function ElapsedTime( _
    Optional ByVal iMethod As Integer = 0, _
    Optional ByVal ElapsedSecondsIn As Variant) _
    As String
'
'==============================================================================
'                                ELAPSED TIME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a human-readable elapsed time string:
'     H:nn:ss - mmm milliseconds - uuu microseconds - nnn nanoseconds
'
' WHY THIS EXISTS
'   This routine supports two related use cases:
'
'     - when no numeric elapsed value is supplied, it measures elapsed time
'       through ElapsedSeconds()
'
'     - when a numeric elapsed value is already available, it formats that
'       value directly without taking a second timing sample
'
'   This avoids the double-measurement problem in callers that already captured
'   ElapsedSeconds() separately for logging or further numeric processing
'
' INPUTS
'   iMethod (optional)
'     0   => use the active timer session method
'     1-6 => explicit method to validate against the active session
'
'   ElapsedSecondsIn (optional)
'     Optional numeric elapsed-seconds value to format directly
'
' RETURNS
'   String
'     The formatted elapsed time string
'
' BEHAVIOR
'   - If ElapsedSecondsIn is omitted, measures elapsed time through
'     ElapsedSeconds(iMethod)
'   - If ElapsedSecondsIn is supplied, formats that value directly
'   - Splits elapsed time into whole seconds and fractional seconds
'   - Decomposes the fractional part into millisecond / microsecond /
'     nanosecond display groups
'   - Formats the whole-second duration manually so hours do not wrap at 24
'
' ERROR POLICY
'   Delegates timer/session validation to ElapsedSeconds() when measurement is
'   requested
'
' DEPENDENCIES
'   - ElapsedSeconds (only when ElapsedSecondsIn is omitted)
'
' NOTES
'   - Nanosecond output is presentation precision only and does not imply actual
'     measurement resolution
'   - When ElapsedSecondsIn is supplied, no new timing sample is taken
'
' UPDATED
'   2026-04-18
'==============================================================================
    
'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Sec         As Double   'Elapsed seconds to format
    Dim WholeS      As Long     'Whole seconds component (floor)
    Dim Frac        As Double   'Fractional seconds component in [0,1)
    Dim NS_Total    As Long     'Fractional part scaled to nanoseconds
    Dim Ms          As Long     'Milliseconds component
    Dim Us          As Long     'Microseconds component
    Dim NN          As Long     'Nanoseconds component
    Dim Hours       As Long     'Whole hours in the elapsed duration
    Dim Minutes     As Long     'Minute component
    Dim Seconds     As Long     'Second component
    Dim sMS         As String   'Zero-padded milliseconds text
    Dim sUS         As String   'Zero-padded microseconds text
    Dim sNN         As String   'Zero-padded nanoseconds text

'------------------------------------------------------------------------------
' RESOLVE ELAPSED SECONDS TO FORMAT
'------------------------------------------------------------------------------
    'Use the supplied elapsed-seconds value when provided
        If Not IsMissing(ElapsedSecondsIn) Then
            Sec = CDbl(ElapsedSecondsIn)
    'Otherwise measure elapsed seconds now
        Else
            Sec = ElapsedSeconds(iMethod)
        End If

'------------------------------------------------------------------------------
' NORMALIZE INPUT
'------------------------------------------------------------------------------
    'Defensively clamp negative elapsed values
        If Sec < 0# Then Sec = 0#

'------------------------------------------------------------------------------
' SPLIT WHOLE / FRACTIONAL SECONDS
'------------------------------------------------------------------------------
    'Extract whole seconds and the fractional remainder
        WholeS = CLng(Fix(Sec))
        Frac = Sec - CDbl(WholeS)
    'Defensive clamp for negative fractional jitter
        If Frac < 0# Then Frac = 0#
        
'------------------------------------------------------------------------------
' CONVERT FRACTIONAL SECONDS TO ns/ms/us/nnn
'------------------------------------------------------------------------------
    'Convert fractional seconds to integer nanoseconds using floor semantics
        NS_Total = CLng(Fix(Frac * 1000000000#))
    'Defensively clamp unexpected negative values
        If NS_Total < 0 Then
            NS_Total = 0
        End If
    'Defensively clamp unexpected overflow to one-second boundary
        If NS_Total > 999999999 Then
            NS_Total = 999999999
        End If
    'Decompose nanoseconds into mmm / uuu / nnn display groups
        Ms = NS_Total \ 1000000
        Us = (NS_Total \ 1000) Mod 1000
        NN = NS_Total Mod 1000
        
'------------------------------------------------------------------------------
' DECOMPOSE WHOLE-SECOND DURATION
'------------------------------------------------------------------------------
    'Compute duration hours/minutes/seconds without 24-hour wraparound
        Hours = WholeS \ 3600
        Minutes = (WholeS \ 60) Mod 60
        Seconds = WholeS Mod 60

'------------------------------------------------------------------------------
' FORMAT COMPONENTS
'------------------------------------------------------------------------------
    'Format each sub-second group as exactly three digits
        sMS = Right$("000" & CStr(Ms), 3)
        sUS = Right$("000" & CStr(Us), 3)
        sNN = Right$("000" & CStr(NN), 3)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the final formatted duration string
        ElapsedTime = CStr(Hours) & ":" & _
                      Right$("00" & CStr(Minutes), 2) & ":" & _
                      Right$("00" & CStr(Seconds), 2) & " - " & _
                      sMS & " milliseconds - " & _
                      sUS & " microseconds - " & _
                      sNN & " nanoseconds"

End Function

'
'------------------------------------------------------------------------------
'
'                       PUBLIC: SESSION / STATE INSPECTION
'
'------------------------------------------------------------------------------
'

Public Property Get T1() As Double
'
'==============================================================================
'                                      T1
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last captured raw start timestamp for the most recent timing
'   session established by StartTimer().
'
' WHY THIS EXISTS
'   Useful for diagnostics, verification, and unit tests when the caller wants
'   to inspect the internal raw start capture without triggering formatting or
'   recomputation logic.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_T1 value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the raw start timestamp without modification.
'   - The meaning of the value depends on the active timing backend used when
'     StartTimer() was called:
'       * Method 1 (Timer)             : seconds since midnight
'       * Method 2 (GetTickCount / GetTickCount64) : seconds since boot (wrapper-normalized)
'       * Method 3 (timeGetTime)       : seconds since boot (wrapper-normalized)
'       * Method 4 (timeGetSystemTime) : seconds from a millisecond source
'       * Method 5 (QPC)               : CDbl(m_qpcStart), i.e. raw QPC ticks
'       * Method 6 (Now)               : Now() * 86400 as stored by the class
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_T1
'
' NOTES
'   - Prefer ElapsedSeconds() for the canonical numeric elapsed-time result.
'   - Prefer ElapsedTime() for display-oriented diagnostics.
'   - T1 is an inspection/debugging surface, not the primary timing API.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the raw start timestamp.
        T1 = m_T1
End Property

Public Property Get T2() As Double
'
'==============================================================================
'                                      T2
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last captured raw end timestamp for the most recent timing
'   session as evaluated by ElapsedSeconds() or by ElapsedTime() when that
'   function internally performs a timing read
'
' WHY THIS EXISTS
'   Useful for diagnostics, verification, and unit tests when the caller wants
'   to inspect the internal raw end capture directly
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_T2 value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the raw end timestamp without modification.
'   - m_T2 is updated when ElapsedSeconds() is evaluated.
'   - m_T2 is also updated when ElapsedTime() is called without
'     ElapsedSecondsIn, because that path internally measures elapsed time.
'   - When ElapsedTime() is called with ElapsedSecondsIn, no new timing sample
'     is taken and m_T2 is not updated by that formatting step.
'   - The meaning of the value depends on the timing backend:
'       * Methods 1 / 2 / 3 / 4 / 6 : stored in seconds
'       * Method 5 (QPC)             : stored as CDbl(m_qpcEnd), i.e. raw QPC ticks
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_T2
'
' NOTES
'   - If ElapsedSeconds() has not yet been called after StartTimer(), T2 may
'     still reflect a prior session or its reset value.
'   - T2 is an inspection/debugging surface, not the primary timing API.
'
' UPDATED
'   2026-04-15
'==============================================================================
'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the raw end timestamp.
        T2 = m_T2
End Property

Public Property Get ET() As Double
'
'==============================================================================
'                                      ET
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last computed elapsed time in seconds for the most recent
'   timing session
'
' WHY THIS EXISTS
'   Allows callers, diagnostics, and tests to inspect the cached elapsed value
'   without re-running timing logic or invoking formatting code
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     The internal m_ET value exactly as currently stored.
'
' BEHAVIOR
'   - Returns the cached elapsed-seconds value without modification.
'   - m_ET is set by ElapsedSeconds().
'   - m_ET is also updated when ElapsedTime() is called without
'     ElapsedSecondsIn, because that path internally measures elapsed time.
'   - When ElapsedTime() is called with ElapsedSecondsIn, the supplied value is
'     formatted directly and m_ET is not updated by that formatting step.
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - m_ET
'
' NOTES
'   - For the canonical numeric timing result, prefer calling
'     ElapsedSeconds() directly.
'   - ET is an inspection/debugging surface and cache exposure, not the
'     primary timing API.
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the cached elapsed seconds.
        ET = m_ET
End Property

Public Property Get StrictMode() As Boolean
'------------------------------------------------------------------------------
' STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the class raises errors for invalid/mismatched timer usage.
'------------------------------------------------------------------------------
    'Return the current strict-mode flag.
        StrictMode = m_StrictMode
End Property

Public Property Let StrictMode(ByVal Value As Boolean)
'------------------------------------------------------------------------------
' STRICT MODE
'------------------------------------------------------------------------------
' PURPOSE
'   Sets whether the class raises errors for invalid/mismatched timer usage.
'------------------------------------------------------------------------------
    'Persist the strict-mode flag.
        m_StrictMode = Value
End Property

Public Property Get ActiveMethodID() As Integer
'------------------------------------------------------------------------------
'                               ACTIVE METHOD ID
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the active timer method selected by the most recent StartTimer().
'------------------------------------------------------------------------------
    'Return the current active method.
        ActiveMethodID = m_ActiveMethodID
End Property

Public Property Get HasActiveSession() As Boolean
'------------------------------------------------------------------------------
'                              HAS ACTIVE SESSION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE when a timer session is currently active.
'------------------------------------------------------------------------------
    'Return the current session-start flag.
        HasActiveSession = m_HasActiveSession
End Property

Public Property Get MethodName( _
    ByVal Idx As Integer) _
    As String
'
'==============================================================================
'                                 METHOD NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the human-readable name of a timing backend by method index.
'
' WHY THIS EXISTS
'   Provides a stable mapping from the numeric timing-method identifier used by
'   StartTimer() / ElapsedSeconds() to a readable backend label for diagnostics,
'   reporting, and UI/debug output.
'
' INPUTS
'   Idx
'     Timing method index:
'       1 = "Timer"
'       2 = "GetTickCount / GetTickCount64"
'       3 = "timeGetTime"
'       4 = "timeGetSystemTime"
'       5 = "QPC"
'       6 = "Now()"
'
' RETURNS
'   String
'     The corresponding method label when Idx is valid (1..6);
'     otherwise vbNullString.
'
' BEHAVIOR
'   - Validates the supplied method index.
'   - Returns the stored label from m_MethodName() when the index is valid.
'   - Returns vbNullString for out-of-range indices.
'
' ERROR POLICY
'   Does not raise errors.
'   Invalid indices return vbNullString.
'
' DEPENDENCIES
'   - m_MethodName
'
' NOTES
'   This property is intended for display/diagnostic use and does not perform
'   timing-method normalization.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Guard against invalid indices to avoid subscript errors
        If (Idx < 1) Or (Idx > 6) Then
            MethodName = vbNullString
            Exit Property
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the corresponding method label.
        MethodName = m_MethodName(Idx)
End Property

Public Property Get RunLabel() As String
'
'==============================================================================
'                                  RUN LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current run label stored on this instance
'
' WHY THIS EXISTS
'   Structured reporting helpers and diagnostics may need to inspect the
'   current run label without exporting the full checkpoint report
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Current run label
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the current run label
        RunLabel = m_RunLabel

End Property

'
'------------------------------------------------------------------------------
'
'                   PUBLIC: DIAGNOSTICS / BENCHMARK HELPERS
'
'------------------------------------------------------------------------------
'

Public Function OverheadMeasurement_Seconds( _
    Optional ByVal iMethod As Integer = 5, _
    Optional ByVal Iterations As Long = 1000) _
    As Double
'
'==============================================================================
'                        OVERHEAD MEASUREMENT IN SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Estimates the average near-empty measurement overhead for the selected timer
'   backend across multiple iterations.
'
' WHY THIS EXISTS
'   This function provides the numeric benchmark primitive behind
'   OverheadMeasurement_Text().
'
'   It is intended for programmatic comparison of the approximate cost of a
'   minimal timing cycle through cPerformanceManager.
'
' INPUTS
'   iMethod (optional)
'     Timing backend to benchmark.
'
'   Iterations (optional)
'     Number of near-empty timing cycles to average.
'     Values less than 1 are coerced to 1.
'
' RETURNS
'   Double
'     Average near-empty measurement overhead in seconds.
'
' BEHAVIOR
'   - Normalizes and validates the requested timer method.
'   - Warms up the timing path once before the measured loop.
'   - Repeats a minimal StartTimer/ElapsedSeconds cycle for the requested
'     number of iterations.
'   - Returns the arithmetic mean of the measured elapsed values.
'
' ERROR POLICY
'   Validation behavior is delegated to Method_NormalizeStart().
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - StartTimer
'   - ElapsedSeconds
'
' NOTES
'   - This measures the approximate overhead of a minimal class timing cycle,
'     not just the raw underlying backend call in isolation.
'   - Results depend on the selected backend, host state, and general runtime
'     noise.
'   - Method 3 may incur one-time initialization effects on first use because
'     EnsureTimerResolution1ms can activate 1ms timer resolution.
'   - Performs the measurement on an isolated temporary cPerformanceManager
'     instance so that the caller's current checkpoint / report state is not
'     disturbed.
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim K               As Long                  'Loop counter
    Dim TotalS          As Double                'Accumulated elapsed seconds
    Dim ResolvedMethod  As Integer               'Validated timer method
    Dim WarmUpS         As Double                'Throwaway elapsed value used only for warm-up
    Dim cPM_Work        As cPerformanceManager   'Isolated worker instance

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Enforce a minimum of one iteration
        If Iterations < 1 Then Iterations = 1
    'Normalize and validate the timer method using the current instance policy
        ResolvedMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' INITIALIZE ISOLATED WORKER
'------------------------------------------------------------------------------
    'Create an isolated worker instance so checkpoint/report state on the
    'current instance is not disturbed
        Set cPM_Work = New cPerformanceManager
    'Mirror the current strict-mode policy into the isolated worker
        cPM_Work.StrictMode = m_StrictMode

'------------------------------------------------------------------------------
' WARM UP
'------------------------------------------------------------------------------
    'Warm up the timing path once before measurement
        cPM_Work.StartTimer ResolvedMethod, False
        WarmUpS = cPM_Work.ElapsedSeconds()

'------------------------------------------------------------------------------
' MEASURE
'------------------------------------------------------------------------------
    'Run a repeated near-empty timing loop and accumulate the measured elapsed values
        For K = 1 To Iterations
            cPM_Work.StartTimer ResolvedMethod, False
            TotalS = TotalS + cPM_Work.ElapsedSeconds()
        Next K

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the average measurement overhead in seconds
        OverheadMeasurement_Seconds = TotalS / CDbl(Iterations)

'------------------------------------------------------------------------------
' CLEANUP
'------------------------------------------------------------------------------
    'Release isolated worker environment changes
        cPM_Work.ResetEnvironment
    'Release the isolated worker instance
        Set cPM_Work = Nothing

End Function

Public Property Get OverheadMeasurement_Text( _
    Optional ByVal iMethod As Integer = 5, _
    Optional ByVal Iterations As Long = 1000, _
    Optional ByVal OverheadSecondsIn As Variant) _
    As String
'
'==============================================================================
'                         OVERHEAD MEASUREMENT TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a human-readable report of the average near-empty measurement
'   overhead for the selected timer backend
'
' WHY THIS EXISTS
'   OverheadMeasurement_Seconds() provides the numeric benchmark primitive
'
'   This property exists as the presentation-oriented companion for diagnostics,
'   logs, and quick inspection
'
'   When an already measured overhead value is supplied, this routine formats
'   that value directly without taking a second measurement
'
' INPUTS
'   iMethod (optional)
'     Timing backend whose near-empty average measurement overhead should be
'     reported
'
'   Iterations (optional)
'     Number of near-empty timing cycles to average
'     Values less than 1 are coerced to 1 by OverheadMeasurement_Seconds()
'
'   OverheadSecondsIn (optional)
'     Optional already measured overhead value to format directly
'
' RETURNS
'   String
'     A human-readable measurement-overhead report
'
' BEHAVIOR
'   - Normalizes and validates the requested timer method
'   - If OverheadSecondsIn is omitted, measures average near-empty overhead
'   - If OverheadSecondsIn is supplied, formats that value directly
'   - Returns a readable string containing the backend name and average
'     measurement overhead in seconds
'
' ERROR POLICY
'   Validation behavior is delegated to Method_NormalizeStart()
'
' DEPENDENCIES
'   - Method_NormalizeStart
'   - OverheadMeasurement_Seconds
'   - MethodName
'
' NOTES
'   This property mirrors the design of ElapsedTime:
'     - measure when no numeric value is supplied
'     - format directly when a numeric value is supplied
'
' UPDATED
'   2026-04-19
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim AvgS            As Double    'Average measurement overhead in seconds
    Dim ResolvedMethod  As Integer   'Validated timer method

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Normalize and validate the timer method
        ResolvedMethod = Method_NormalizeStart(iMethod)

'------------------------------------------------------------------------------
' RESOLVE NUMERIC OVERHEAD VALUE
'------------------------------------------------------------------------------
    'Use the supplied numeric overhead value when provided
        If Not IsMissing(OverheadSecondsIn) Then
            AvgS = CDbl(OverheadSecondsIn)
    'Otherwise measure average near-empty overhead now
        Else
            AvgS = OverheadMeasurement_Seconds(ResolvedMethod, Iterations)
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a readable diagnostic string
        OverheadMeasurement_Text = "Average measurement overhead - " & _
                                   MethodName(ResolvedMethod) & " - " & _
                                   Format$(AvgS, "0.000000000") & " seconds"

End Property
Public Property Get Get_SystemTickInterval() As String
'
'==============================================================================
'                           GET SYSTEM TICK INTERVAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the nominal system clock tick interval in milliseconds, derived from
'   GetSystemTimeAdjustment.
'
' WHY THIS EXISTS
'   This property provides a coarse diagnostic of the system clock increment
'   that can affect low-resolution timing sources and general scheduler
'   granularity.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable tick-interval text.
'
' BEHAVIOR
'   - Calls GetSystemTimeAdjustment to retrieve the nominal system clock
'     increment.
'   - Converts lpTimeIncrement from 100-nanosecond units to milliseconds.
'   - Returns a deterministic "(unknown)" string if the API call fails.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when the underlying API call fails.
'
' DEPENDENCIES
'   - GetSystemTimeAdjustment
'
' NOTES
'   - lpTimeIncrement from GetSystemTimeAdjustment is expressed in
'     100-nanosecond units, hence division by 10,000 to obtain milliseconds.
'   - This is not the same as QPC resolution and does not guarantee actual
'     scheduling precision.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long     'Return code from GetSystemTimeAdjustment
    Dim lA As Long     'lpTimeAdjustment (unused here)
    Dim LB As Long     'lpTimeIncrement (100ns units)
    Dim lC As Long     'lpTimeAdjustmentDisabled (unused here)

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Query system time adjustment parameters.
        RC = GetSystemTimeAdjustment(lA, LB, lC)

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when the API call fails.
        If RC = 0 Then
            Get_SystemTickInterval = "Tick Interval = (unknown)"
            Exit Property
        End If

    'Convert tick increment (100ns units) to milliseconds.
        Get_SystemTickInterval = "Tick Interval = " & (LB / 10000#) & " milliseconds"
End Property

Public Property Get QPC_Get_SystemTickInterval() As String
'
'==============================================================================
'                       QPC GET SYSTEM TICK INTERVAL
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC tick interval in seconds per tick.
'
' WHY THIS EXISTS
'   This property provides a human-readable diagnostic for the effective
'   granularity of the QueryPerformanceCounter timebase.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable QPC tick interval text.
'
' BEHAVIOR
'   - Fetches the true QPC frequency via QPC_GetFrequencyHz().
'   - Returns "(unknown)" when frequency is unavailable.
'   - Otherwise returns the reciprocal of the QPC frequency in seconds/tick.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when QPC frequency is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   This is a presentation-oriented diagnostic property.
'   Use numeric helpers when machine-readable values are required.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Hz As Double      'True QPC frequency in ticks/second

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Fetch the true QPC frequency
        Hz = QPC_GetFrequencyHz()

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when frequency is unavailable
        If Hz = 0# Then
            QPC_Get_SystemTickInterval = "QPC Tick interval is: (unknown)"
            Exit Property
        End If

    'Return seconds-per-tick
        QPC_Get_SystemTickInterval = "QPC Tick interval is: " & (1# / Hz)
End Property

Public Property Get QPC_FrequencyPerSecond() As String
'
'==============================================================================
'                          QPC FREQUENCY PER SECOND
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC frequency as a human-readable string.
'
' WHY THIS EXISTS
'   This property provides a presentation-oriented diagnostic for the effective
'   QueryPerformanceCounter frequency used by the class.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     Human-readable QPC frequency text.
'
' BEHAVIOR
'   - Fetches the true QPC frequency via QPC_GetFrequencyHz().
'   - Returns "(unknown)" when frequency is unavailable.
'   - Otherwise returns ticks-per-second as readable text.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns a deterministic string when QPC frequency is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   This is a presentation-oriented diagnostic property.
'   Use a numeric helper/property when machine-readable frequency values are
'   required.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Hz As Double      'True QPC frequency in ticks/second

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Fetch the true QPC frequency
        Hz = QPC_GetFrequencyHz()

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a deterministic string when frequency is unavailable
        If Hz = 0# Then
            QPC_FrequencyPerSecond = "QPC Tick frequency is: (unknown)"
            Exit Property
        End If

    'Return ticks-per-second as text
        QPC_FrequencyPerSecond = "QPC Tick frequency is: " & Hz & " ticks/second"
End Property

Public Property Get QPC_FrequencyPerSecond_Value() As Double
'
'==============================================================================
'                       QPC FREQUENCY PER SECOND VALUE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QueryPerformanceCounter frequency as a numeric Double.
'
' WHY THIS EXISTS
'   This property is the machine-readable companion to
'   QPC_FrequencyPerSecond(). It provides the true QPC frequency in
'   ticks/second for calculations, assertions, diagnostics, or external
'   reporting that should not depend on formatted text.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     True QPC frequency in ticks/second.
'     Returns 0 when QPC frequency is unavailable.
'
' BEHAVIOR
'   Delegates directly to QPC_GetFrequencyHz().
'
' ERROR POLICY
'   Does not raise errors.
'   Returns 0 when QPC frequency information is unavailable.
'
' DEPENDENCIES
'   - QPC_GetFrequencyHz
'
' NOTES
'   Use QPC_FrequencyPerSecond() when human-readable text is desired.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the true QPC frequency in ticks/second.
        QPC_FrequencyPerSecond_Value = QPC_GetFrequencyHz()
End Property

'
'------------------------------------------------------------------------------
'
'                   PUBLIC: EXECUTION CONTROL / ENVIRONMENT
'
'------------------------------------------------------------------------------
'

Public Sub Pause( _
    ByVal dSeconds As Double, _
    Optional ByVal iMethod As Integer = 1)
'
'==============================================================================
'                                     PAUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Pauses execution for dSeconds (limited to <= 3600 seconds) using one of the
'   supported waiting mechanisms.
'
' WHY THIS EXISTS
'   This routine provides a centralized pause utility with multiple waiting
'   strategies so callers can choose between lower overhead, better UI
'   responsiveness, and simpler blocking behavior.
'
' INPUTS
'   dSeconds
'     Desired pause duration in seconds.
'     - <= 0   : no-op (returns immediately)
'     - > 3600 : rejected (returns immediately) to avoid excessive blocking
'
'   iMethod (optional)
'     Selects the pause mechanism:
'       1 => Sleep API (kernel32) (millisecond granularity; no DoEvents)
'       2 => Timer + DoEvents loop (coarser; yields to Excel/UI)
'       3 => Application.Wait for coarse waiting + Timer guard
'       4 => Now + DoEvents loop + Timer guard
'       Else => treated like method 4
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Method 1 is the lowest-overhead pause mechanism and does not call
'     DoEvents.
'   - Method 2 uses Timer() delta logic and yields via DoEvents.
'   - Method 3 uses Application.Wait as the coarse wait engine, then applies a
'     Timer()-based guard loop so the total pause cannot undershoot.
'   - Method 4 uses a Date/Now-flavored yielding loop and also applies a
'     Timer()-based guard loop so the total pause cannot undershoot.
'
' ERROR POLICY
'   Invalid, non-positive, or excessively large pause durations are treated as
'   no-op exits.
'
' DEPENDENCIES
'   - Sleep
'   - Timer
'   - Application.Wait
'   - Now
'   - DoEvents
'
' NOTES
'   - Method 2 is midnight-safe because it measures elapsed delta rather than
'     comparing Timer() against an absolute target.
'   - Methods 3 and 4 use a Timer()-based guard loop specifically because
'     Application.Wait and Now are too coarse to guarantee the requested pause
'     duration accurately on their own.
'   - This routine uses local elapsed-time variables and does not overwrite the
'     class-level m_ET timing cache.
'
' UPDATED
'   2026-03-30
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim PauseTime   As Double   'Requested pause duration in seconds
    Dim StartS      As Double   'Timer() snapshot at pause start
    Dim ElapsedS    As Double   'Elapsed seconds since StartS
    Dim WaitTill    As Date     'Target timestamp for coarse Date/Wait paths
    Dim Ms          As Long     'Milliseconds for Sleep (method 1)
    Dim WholeS      As Long     'Whole-second coarse-wait portion (method 3)
    Dim FracS       As Double   'Fractional remainder after whole seconds (method 3)
    Dim K           As Long     'Loop counter for repeated Wait calls (method 3)

'------------------------------------------------------------------------------
' VALIDATE / CONTROL
'------------------------------------------------------------------------------
    'Reject non-positive pauses (no-op)
        If dSeconds <= 0# Then Exit Sub
    'Hard cap to avoid accidental long blocking pauses
        If dSeconds > 3600# Then Exit Sub

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Persist the requested pause duration locally
        PauseTime = dSeconds

'------------------------------------------------------------------------------
' DISPATCH
'------------------------------------------------------------------------------
    Select Case iMethod

        Case 1
            'Sleep takes milliseconds; coerce seconds to Long milliseconds
                Ms = CLng(PauseTime * 1000#)

            'Sleep(0) yields the remainder of the thread's timeslice; allow it explicitly
                Sleep Ms

        Case 2
            'Capture the Timer() start snapshot
                StartS = Timer

            'Loop until the requested pause duration has elapsed
                Do
                    'Compute elapsed seconds since the start snapshot
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

        Case 3
            'Application.Wait is coarse and second-based, so use it only for the
            'whole-second portion and then finish with a Timer()-based guard loop
                WholeS = CLng(Fix(PauseTime))
                FracS = PauseTime - CDbl(WholeS)

            'Capture the Timer() start snapshot so we can guarantee the total
            'requested elapsed time even if Application.Wait undershoots
                StartS = Timer

            'Apply the coarse whole-second wait portion using repeated 1-second
            'Application.Wait calls
                For K = 1 To WholeS
                    Application.Wait Now + TimeSerial(0, 0, 1)
                Next K

            'If there is no whole-second portion, do nothing here and let the
            'Timer()-based guard handle the entire requested delay
                If FracS < 0# Then FracS = 0#

            'Guard against undershoot using Timer() delta logic
                Do
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the total requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

        Case Else
            'Preserve the Date/Now-flavored yielding path, but add a Timer()-based
            'guard loop because Now is too coarse to guarantee the requested pause
            'duration accurately by itself
                WaitTill = Now + (PauseTime / 86400#)

            'Capture the Timer() start snapshot so we can enforce the true
            'requested pause duration
                StartS = Timer

            'Run the coarse Date-based yielding loop first
                Do While Now < WaitTill
                    DoEvents
                Loop

            'Guard against undershoot using Timer() delta logic
                Do
                    'Compute elapsed seconds since the pause started
                        ElapsedS = Timer - StartS

                    'Correct midnight rollover when Timer() wraps at 24 hours
                        If ElapsedS < 0# Then ElapsedS = ElapsedS + 86400#

                    'Exit once the total requested pause duration has been reached
                        If ElapsedS >= PauseTime Then Exit Do

                    'Yield to Excel/UI while waiting
                        DoEvents
                Loop

    End Select
End Sub
Public Sub ResetEnvironment()
'
'==============================================================================
'                              RESET ENVIRONMENT
'------------------------------------------------------------------------------
' PURPOSE
'   Explicitly restores environment changes made by this instance:
'     - releases timeBeginPeriod(1) if enabled
'     - ends this instance's shared TW session if active
'
' WHY THIS EXISTS
'   This is the primary explicit cleanup routine for cPerformanceManager.
'   Callers should use this routine when they want deterministic release of any
'   environment changes made by the instance, rather than relying only on
'   Class_Terminate as a fallback.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - If this instance enabled 1ms multimedia timer resolution, balances it with
'     timeEndPeriod 1 and clears the lifecycle flag.
'   - If this instance is active in the shared TW manager, ends only this
'     instance's shared TW session.
'   - Clears this instance's TW activity flag and exemption mask after the
'     shared session is ended successfully.
'
' ERROR POLICY
'   Raises errors normally.
'   This routine is the primary cleanup contract and should fail loudly if an
'   explicit cleanup action cannot be completed.
'
' DEPENDENCIES
'   - timeEndPeriod
'   - PM_TW_EndSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   - This routine is intentionally idempotent:
'       * if m_TimePeriodOn = False, no timer-resolution cleanup is performed
'       * if m_TW_Active = False, no TW cleanup is performed
'   - Class_Terminate may call this routine as a final safety net, but
'     Class_Terminate should handle teardown errors defensively.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' BALANCE timeBeginPeriod(1) IF IT WAS ENABLED
'------------------------------------------------------------------------------
    'Release the previously requested 1ms multimedia timer resolution
        If m_TimePeriodOn Then
            timeEndPeriod 1
            m_TimePeriodOn = False
        End If

'------------------------------------------------------------------------------
' END THIS INSTANCE'S SHARED TW SESSION IF ACTIVE
'------------------------------------------------------------------------------
    'End only this instance's TW participation; the shared manager handles
    'restore/recompute correctly depending on whether other instances remain
        If m_TW_Active Then
            PM_TW_EndSession TW_InstanceKey()
            m_TW_Active = False
            m_Except = TW_Enum.None
        End If
End Sub

'
'------------------------------------------------------------------------------
'
'                      PUBLIC: SHARED TIME-WASTER CONTROL
'
'------------------------------------------------------------------------------
'

Public Sub TW_Turn_OFF( _
    Optional ByVal Except As TW_Enum = TW_Enum.None)
'
'==============================================================================
'                           TURN OFF TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Starts or updates this instance's shared Excel "time waster" suppression
'   session.
'
' WHY THIS EXISTS
'   Application-level TW settings are global to Excel, not instance-local.
'   Therefore this routine registers the current instance with the shared
'   module-level TW scope manager rather than trying to manage Application state
'   independently inside the class.
'
' INPUTS
'   Except (optional)
'     Bitmask of TW_Enum flags that this instance does NOT want disabled.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - Persists this instance's exemption mask.
'   - Registers this instance in the shared TW manager if not already active.
'   - Updates this instance's effective mask if already active.
'   - Marks this instance as actively participating in the shared TW scope.
'
' ERROR POLICY
'   Raises errors normally.
'   Registration/update failures in the shared TW manager should not be silently
'   ignored during the normal execution path.
'
' DEPENDENCIES
'   - PM_TW_BeginSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   - This routine is effectively idempotent at the instance level:
'       * first call => begins this instance's shared session
'       * subsequent calls => update this instance's exemption mask
'   - The effective Excel state is determined by the shared manager from the
'     aggregate disable-mask of all active instances.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Persist this instance's current exemption mask
        m_Except = Except

'------------------------------------------------------------------------------
' BEGIN / UPDATE SHARED SESSION
'------------------------------------------------------------------------------
    'Register or update this instance in the shared TW manager
        PM_TW_BeginSession TW_InstanceKey(), CLng(m_Except)

'------------------------------------------------------------------------------
' UPDATE INSTANCE STATE
'------------------------------------------------------------------------------
    'Mark this instance as actively participating in the shared TW scope.
        m_TW_Active = True
End Sub

Public Sub TW_Turn_ON()
'
'==============================================================================
'                              TURN ON TIMEWASTERS
'------------------------------------------------------------------------------
' PURPOSE
'   Ends this instance's shared TW suppression session.
'
' WHY THIS EXISTS
'   TW suppression is managed through a shared module-level scope manager
'   because the underlying Excel Application settings are global. Therefore
'   "turning TW back on" means removing only this instance's participation and
'   letting the shared manager determine the correct resulting effective state.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - No-op if this instance is not currently active.
'   - If this instance is active, removes it from the shared TW manager.
'   - If other instances remain active, the shared manager recomputes the
'     effective state from the remaining aggregate disable-mask.
'   - If this is the final active instance, the original baseline is restored.
'   - Clears this instance's local TW activity state and exemption mask after
'     successful removal from the shared manager.
'
' ERROR POLICY
'   Raises errors normally.
'   Explicit shared-session teardown should fail loudly during the normal
'   execution path if it cannot be completed.
'
' DEPENDENCIES
'   - PM_TW_EndSession
'   - TW_InstanceKey
'   - TW_Enum
'
' NOTES
'   - This routine is idempotent:
'       * if m_TW_Active = False, nothing happens
'   - This routine does not restore Excel state directly.
'     That responsibility belongs to the shared TW manager.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' END SHARED SESSION (IF ACTIVE)
'------------------------------------------------------------------------------
    'Only end a session if this instance is currently active in the shared manager
        If m_TW_Active Then
            PM_TW_EndSession TW_InstanceKey()

            m_TW_Active = False
            m_Except = TW_Enum.None
        End If
End Sub

Public Property Get TW_IsActive() As Boolean
'------------------------------------------------------------------------------
' TW_IsActive
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if this instance currently has an active shared TW session.
'------------------------------------------------------------------------------
    'Return this instance's TW activity state.
        TW_IsActive = m_TW_Active
End Property

Public Property Get TW_ActiveSessionCount() As Long
'------------------------------------------------------------------------------
' TW_ACTIVESESSIONCOUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of active shared TW sessions across all instances.
'------------------------------------------------------------------------------
    'Return the global active shared TW session count.
        TW_ActiveSessionCount = PM_TW_ActiveCount()
End Property


'
'------------------------------------------------------------------------------
'
'                 PUBLIC: CHECKPOINTS / STRUCTURED REPORTING
'
'------------------------------------------------------------------------------
'

Public Sub SetRunLabel( _
    ByVal RunLabel As String)
'
'==============================================================================
'                                SET RUN LABEL
'------------------------------------------------------------------------------
' PURPOSE
'   Stores an optional run label used by later checkpoint captures
'
' WHY THIS EXISTS
'   Structured checkpoint output is easier to interpret when rows can be tagged
'   with a user-supplied run label such as:
'
'     - ImportWorkflow
'     - ScenarioA
'     - Baseline
'     - TW_Off
'
'   To keep report output unambiguous, the run label must be set before the
'   first checkpoint is captured in the current timing session
'
' INPUTS
'   RunLabel
'     User-defined label for the current timing run
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Stores the supplied text as the current run label
'   - Raises if the caller attempts to change the run label after one or more
'     checkpoints have already been captured in the current session
'
' ERROR POLICY
'   Raises when the caller attempts to change the run label after checkpoint
'   capture has already begun in the current timing session
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE SESSION CHECKPOINT STATE
'------------------------------------------------------------------------------
    'Reject run-label changes after checkpoint capture has already begun
        If m_CheckpointCount > 0 Then
            Err.Raise vbObjectError + 1011, _
                      "cPerformanceManager.SetRunLabel", _
                      "RunLabel must be set before the first checkpoint of the current timing session."
        End If

'------------------------------------------------------------------------------
' ASSIGN STATE
'------------------------------------------------------------------------------
    'Store the supplied run label
        m_RunLabel = RunLabel

End Sub
Public Sub ClearCheckpoints()
'
'==============================================================================
'                              CLEAR CHECKPOINTS
'------------------------------------------------------------------------------
' PURPOSE
'   Clears the structured checkpoint state for the current instance
'
' WHY THIS EXISTS
'   Although a new StartTimer session already resets checkpoint state, callers
'   may also want to clear the current structured checkpoint / report state
'   explicitly without recreating the object
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Clears the current run label
'   - Clears the checkpoint counter
'   - Resets the last checkpoint elapsed-seconds marker
'   - Reinitializes checkpoint storage
'
' ERROR POLICY
'   Raises errors normally
'
' DEPENDENCIES
'   - Checkpoint_ClearState
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' CLEAR STATE
'------------------------------------------------------------------------------
    'Clear all checkpoint / report state explicitly
        Checkpoint_ClearState

End Sub
Public Sub Checkpoint( _
    ByVal CheckpointName As String, _
    Optional ByVal NoteText As String = vbNullString)
'
'==============================================================================
'                                  CHECKPOINT
'------------------------------------------------------------------------------
' PURPOSE
'   Captures one named checkpoint within the current timing session
'
' WHY THIS EXISTS
'   A single elapsed-time value is useful, but many real workflows need
'   structured sub-measurements such as:
'
'     - LoadData
'     - BuildArray
'     - Transform
'     - WriteBack
'
'   This routine records both:
'
'     - cumulative elapsed seconds since StartTimer
'     - delta seconds since the previous checkpoint
'
' INPUTS
'   CheckpointName
'     User-defined checkpoint label
'
'   NoteText (optional)
'     Optional descriptive note stored with the checkpoint row
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Requires an active timing session
'   - Captures cumulative elapsed seconds once
'   - Computes delta seconds versus the previous checkpoint
'   - Appends one structured checkpoint row
'   - Updates the last checkpoint marker
'
' ERROR POLICY
'   Raises when no timing session is active
'   Otherwise delegates timing validation behavior to ElapsedSeconds
'
' DEPENDENCIES
'   - ElapsedSeconds
'   - MethodName
'   - Checkpoint_EnsureCapacity
'   - Checkpoint_DefaultName
'
' NOTES
'   This routine intentionally captures numeric timing only
'   It does not call ElapsedTime and therefore avoids formatting overhead
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CurrentS            As Double       'Current cumulative elapsed seconds
    Dim DeltaS              As Double       'Elapsed seconds since previous checkpoint
    Dim NextSeq             As Long         'Next checkpoint sequence number
    Dim EffectiveName       As String       'Normalized checkpoint label

'------------------------------------------------------------------------------
' VALIDATE SESSION STATE
'------------------------------------------------------------------------------
    'Reject checkpoint capture when no timing session is active
        If Not m_HasActiveSession Then
            Err.Raise vbObjectError + 1010, _
                      "cPerformanceManager.Checkpoint", _
                      "Checkpoint cannot be captured before StartTimer."
        End If

'------------------------------------------------------------------------------
' RESOLVE CURRENT ELAPSED TIME
'------------------------------------------------------------------------------
    'Read the current cumulative elapsed seconds for the active session
        CurrentS = Me.ElapsedSeconds

'------------------------------------------------------------------------------
' RESOLVE CHECKPOINT NAME
'------------------------------------------------------------------------------
    'Normalize the checkpoint label and auto-generate one when blank
        EffectiveName = Trim$(CheckpointName)
    'Use a generated label when the caller supplied a blank checkpoint name
        If Len(EffectiveName) = 0 Then
            EffectiveName = Checkpoint_DefaultName(m_CheckpointCount + 1)
        End If

'------------------------------------------------------------------------------
' COMPUTE DELTA / SEQUENCE
'------------------------------------------------------------------------------
    'Resolve the next checkpoint sequence number
        NextSeq = m_CheckpointCount + 1

    'Compute elapsed seconds since the previous checkpoint
        DeltaS = CurrentS - m_LastCheckpointS

    'Defensively clamp negative jitter
        If DeltaS < 0# Then
            DeltaS = 0#
        End If

'------------------------------------------------------------------------------
' ENSURE STORAGE CAPACITY
'------------------------------------------------------------------------------
    'Ensure checkpoint storage can hold the new row
        Checkpoint_EnsureCapacity NextSeq

'------------------------------------------------------------------------------
' APPEND CHECKPOINT ROW
'------------------------------------------------------------------------------
    'Write the structured checkpoint row
        With m_Checkpoints(NextSeq)
            .Seq = NextSeq
            .NameText = EffectiveName
            .NoteText = NoteText
            .DeltaSecondsValue = DeltaS
            .ElapsedSecondsValue = CurrentS
            .MethodID = m_ActiveMethodID
            .MethodNameText = MethodName(m_ActiveMethodID)
            .RunLabelText = m_RunLabel
        End With

'------------------------------------------------------------------------------
' UPDATE STATE
'------------------------------------------------------------------------------
    'Persist the new checkpoint count
        m_CheckpointCount = NextSeq
    'Persist the most recent cumulative checkpoint elapsed seconds
        m_LastCheckpointS = CurrentS

End Sub
Public Property Get CheckpointCount() As Long
'
'==============================================================================
'                              CHECKPOINT COUNT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of captured structured checkpoints
'
' WHY THIS EXISTS
'   Callers often need a lightweight way to know whether checkpoint rows exist
'   and how many were captured
'
' INPUTS
'   None
'
' RETURNS
'   Long
'     Number of captured checkpoints
'
' ERROR POLICY
'   Does not raise errors
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the number of captured checkpoints
        CheckpointCount = m_CheckpointCount

End Property

Public Function ReportAsArray() As Variant
'
'==============================================================================
'                               REPORT AS ARRAY
'------------------------------------------------------------------------------
' PURPOSE
'   Exports the structured checkpoint report as a 2D Variant array
'
' WHY THIS EXISTS
'   A 2D array is the most reusable machine-readable report surface because it
'   can be:
'
'     - written directly to a worksheet range
'     - consumed by helper modules
'     - post-processed by calling code
'
' INPUTS
'   None
'
' RETURNS
'   Variant
'     Two-dimensional 1-based array including a header row
'
' BEHAVIOR
'   - Returns a 2D array with one header row
'   - Returns one additional row per checkpoint
'   - Uses this column order:
'
'       1 = RunLabel
'       2 = Seq
'       3 = Checkpoint
'       4 = Note
'       5 = MethodID
'       6 = MethodName
'       7 = DeltaSeconds
'       8 = CumulativeSeconds
'
'   - When no checkpoints exist, still returns a header-only 2D array
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
    Dim Arr()       As Variant      'Export array
    Dim RowCount    As Long         'Total export rows including the header
    Dim i           As Long         'Checkpoint loop index

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the total export row count including the header row
        RowCount = m_CheckpointCount + 1
    'Create the export array
        ReDim Arr(1 To RowCount, 1 To 8)

'------------------------------------------------------------------------------
' WRITE HEADER ROW
'------------------------------------------------------------------------------
    'Write the report header row
        Arr(1, 1) = "RunLabel"
        Arr(1, 2) = "Seq"
        Arr(1, 3) = "Checkpoint"
        Arr(1, 4) = "Note"
        Arr(1, 5) = "MethodID"
        Arr(1, 6) = "MethodName"
        Arr(1, 7) = "DeltaSeconds"
        Arr(1, 8) = "CumulativeSeconds"

'------------------------------------------------------------------------------
' WRITE DATA ROWS
'------------------------------------------------------------------------------
    'Write one export row per captured checkpoint
        For i = 1 To m_CheckpointCount
            Arr(i + 1, 1) = m_Checkpoints(i).RunLabelText
            Arr(i + 1, 2) = m_Checkpoints(i).Seq
            Arr(i + 1, 3) = m_Checkpoints(i).NameText
            Arr(i + 1, 4) = m_Checkpoints(i).NoteText
            Arr(i + 1, 5) = m_Checkpoints(i).MethodID
            Arr(i + 1, 6) = m_Checkpoints(i).MethodNameText
            Arr(i + 1, 7) = m_Checkpoints(i).DeltaSecondsValue
            Arr(i + 1, 8) = m_Checkpoints(i).ElapsedSecondsValue
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the structured export array
        ReportAsArray = Arr

End Function

Public Function ReportAsText() As String
'
'==============================================================================
'                                REPORT AS TEXT
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the structured checkpoint report as a readable multiline text block
'
' WHY THIS EXISTS
'   This provides a quick developer-facing report surface for:
'
'     - Immediate Window output
'     - ad hoc diagnostics
'     - lightweight reporting without worksheet output
'
' INPUTS
'   None
'
' RETURNS
'   String
'     Multiline text report
'
' BEHAVIOR
'   - Returns a short message when no checkpoints exist
'   - Otherwise returns a readable multiline checkpoint report
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
    Dim S       As String       'Accumulated report text
    Dim i       As Long         'Checkpoint loop index

'------------------------------------------------------------------------------
' HANDLE EMPTY REPORT
'------------------------------------------------------------------------------
    'Return a short message when no checkpoints were captured
        If m_CheckpointCount = 0 Then
            ReportAsText = "No checkpoints captured."
            Exit Function
        End If

'------------------------------------------------------------------------------
' BUILD HEADER
'------------------------------------------------------------------------------
    'Start the report title
        S = "CHECKPOINT REPORT"
    'Append the optional run label when present
        If Len(m_RunLabel) > 0 Then
            S = S & " | RunLabel=" & m_RunLabel
        End If
    'Append the column legend line
        S = S & vbCrLf & _
            "Seq | Checkpoint | DeltaSeconds | CumulativeSeconds | MethodName | Note"

'------------------------------------------------------------------------------
' BUILD DATA LINES
'------------------------------------------------------------------------------
    'Append one readable line per checkpoint
        For i = 1 To m_CheckpointCount
            S = S & vbCrLf & _
                CStr(m_Checkpoints(i).Seq) & " | " & _
                m_Checkpoints(i).NameText & " | " & _
                Format$(m_Checkpoints(i).DeltaSecondsValue, "0.000000000") & " | " & _
                Format$(m_Checkpoints(i).ElapsedSecondsValue, "0.000000000") & " | " & _
                m_Checkpoints(i).MethodNameText & " | " & _
                m_Checkpoints(i).NoteText
        Next i

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the readable structured report text
        ReportAsText = S

End Function

'
'------------------------------------------------------------------------------
'
'                   PRIVATE: SESSION / VALIDATION HELPERS
'
'------------------------------------------------------------------------------
'

Private Function Method_NormalizeStart( _
    ByVal iMethod As Integer) _
    As Integer
'
'==============================================================================
'                          NORMALIZE START METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Normalizes the method requested by StartTimer and validates QPC availability.
'
' WHY THIS EXISTS
'   StartTimer should not contain repeated validation logic for method range and
'   backend availability. This helper centralizes that validation and returns a
'   method that is safe for use by the timing session.
'
' INPUTS
'   iMethod
'     Requested timer method identifier.
'
' RETURNS
'   Integer
'     A validated timer method identifier.
'
' BEHAVIOR
'   - Accepts methods 1..6 as valid.
'   - In non-strict mode:
'       * invalid method values default to 5 (QPC)
'       * if QPC is unavailable, method 5 falls back to 2
'   - In strict mode:
'       * invalid method values raise an error
'       * requesting QPC when unavailable raises an error
'
' ERROR POLICY
'   - Raises on invalid method values when m_StrictMode = True
'   - Raises when QPC is requested but unavailable and m_StrictMode = True
'
' DEPENDENCIES
'   - m_QpcAvailable
'   - m_StrictMode
'
' NOTES
'   Method 2 is the non-strict fallback when QPC is unavailable because it is a
'   practical general-purpose monotonic timing source.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE METHOD RANGE
'------------------------------------------------------------------------------
    'Reject invalid method values in strict mode
        If (iMethod < 1) Or (iMethod > 6) Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1000, _
                          "cPerformanceManager.Method_NormalizeStart", _
                          "Invalid timer method passed to StartTimer."
            Else
                iMethod = 5
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE QPC AVAILABILITY
'------------------------------------------------------------------------------
    'If QPC is requested but unavailable, either raise or fallback
        If iMethod = 5 And Not m_QpcAvailable Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1001, _
                          "cPerformanceManager.Method_NormalizeStart", _
                          "QPC is unavailable on this system."
            Else
                iMethod = 2
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the validated method
        Method_NormalizeStart = iMethod
End Function

Private Function Method_ResolveElapsed( _
    ByVal iMethod As Integer) _
    As Integer
'
'==============================================================================
'                          RESOLVE ELAPSED METHOD
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves which timer method ElapsedSeconds should use:
'     - 0 => use the active session method
'     - 1..6 => validate against the active session method
'
' WHY THIS EXISTS
'   Once StartTimer establishes a timing session, elapsed-time reads must remain
'   bound to that same backend. This helper centralizes the validation rules
'   that prevent cross-method misuse.
'
' INPUTS
'   iMethod
'     Requested elapsed-time method:
'       0   => use active session method
'       1-6 => explicit method to validate against the active session
'
' RETURNS
'   Integer
'     The validated method that ElapsedSeconds should use.
'
' BEHAVIOR
'   - Raises or fails fast if no session has been started.
'   - Resolves method 0 to the active session method.
'   - Rejects invalid explicit method values.
'   - Rejects explicit methods that do not match the active session method.
'
' ERROR POLICY
'   - Raises when ElapsedSeconds is called before StartTimer if m_StrictMode = True
'   - Raises on invalid explicit method values if m_StrictMode = True
'   - Raises on explicit method/session mismatch if m_StrictMode = True
'   - In non-strict mode, returns the active method when possible; otherwise 0
'
' DEPENDENCIES
'   - m_HasActiveSession
'   - m_ActiveMethodID
'   - m_StrictMode
'
' NOTES
'   This routine is one of the key safeguards that turns the class from a loose
'   timer wrapper into a real timing-session manager.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' VALIDATE SESSION STATE
'------------------------------------------------------------------------------
    'Elapsed timing is invalid before StartTimer has established a session.
        If Not m_HasActiveSession Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1002, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "ElapsedSeconds called before StartTimer."
            Else
                Method_ResolveElapsed = 0
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' USE ACTIVE METHOD WHEN OMITTED
'------------------------------------------------------------------------------
    'Method 0 means: use the currently active session method.
        If iMethod = 0 Then
            Method_ResolveElapsed = m_ActiveMethodID
            Exit Function
        End If

'------------------------------------------------------------------------------
' VALIDATE EXPLICIT METHOD RANGE
'------------------------------------------------------------------------------
    'Reject invalid explicit method values.
        If (iMethod < 1) Or (iMethod > 6) Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1003, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "Invalid timer method passed to ElapsedSeconds."
            Else
                Method_ResolveElapsed = m_ActiveMethodID
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' VALIDATE EXPLICIT METHOD MATCH
'------------------------------------------------------------------------------
    'Reject mismatches between StartTimer's method and ElapsedSeconds' method.
        If iMethod <> m_ActiveMethodID Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1004, _
                          "cPerformanceManager.Method_ResolveElapsed", _
                          "ElapsedSeconds method does not match the active StartTimer method."
            Else
                Method_ResolveElapsed = m_ActiveMethodID
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the validated explicit method.
        Method_ResolveElapsed = iMethod
End Function

Private Function TW_InstanceKey() As String
'
'==============================================================================
'                               TW INSTANCE KEY
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a stable per-instance key used by the shared TW scope manager.
'
' WHY THIS EXISTS
'   The shared TW manager must distinguish overlapping cPerformanceManager
'   instances so it can maintain the correct aggregate disable-mask and remove
'   the correct instance when a shared TW session ends.
'
' INPUTS
'   None.
'
' RETURNS
'   String
'     A string representation of the current object's pointer value.
'
' BEHAVIOR
'   - Uses ObjPtr(Me) as the per-instance identity source.
'   - Uses CLngPtr on VBA7 builds for pointer-width safety.
'   - Falls back to ObjPtr(Me) directly on pre-VBA7 builds.
'
' ERROR POLICY
'   Raises errors normally.
'
' DEPENDENCIES
'   - ObjPtr
'   - CLngPtr (VBA7 only)
'
' NOTES
'   - The returned key is intended only for live in-memory instance tracking.
'   - It is stable for the lifetime of the object instance.
'   - It is not intended as a persisted or cross-session identifier.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    #If VBA7 Then
        'Use the current object's pointer value as the unique instance key
            TW_InstanceKey = CStr(CLngPtr(ObjPtr(Me)))
    #Else
        'Pre-VBA7 fallback
            TW_InstanceKey = CStr(ObjPtr(Me))
    #End If
End Function

'
'------------------------------------------------------------------------------
'
'                       PRIVATE: TIMING SOURCE HELPERS
'
'------------------------------------------------------------------------------
'

Private Function QPC_GetFrequencyHz() As Double
'
'==============================================================================
'                             QPC FREQUENCY HZ
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the QPC frequency in true ticks/second.
'
' WHY THIS EXISTS
'   The class stores m_qpcFreq as Currency for compatibility with the API usage
'   pattern in the QPC routines. This helper centralizes the conversion back to
'   true frequency units so that QPC diagnostics and related calculations remain
'   correct and consistent.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     True QPC frequency in ticks/second.
'     Returns 0 when QPC frequency is unavailable.
'
' BEHAVIOR
'   - Returns 0 when the stored QPC frequency is unavailable.
'   - Otherwise undoes Currency's 10,000 scaling and returns the true frequency.
'
' ERROR POLICY
'   Does not raise errors.
'   Returns 0 when frequency information is unavailable.
'
' DEPENDENCIES
'   - m_qpcFreq
'
' NOTES
'   m_qpcFreq is stored as Currency, so we must undo Currency's 10,000 scaling.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Return 0 when frequency is unavailable
        If m_qpcFreq = 0 Then
            QPC_GetFrequencyHz = 0#
            Exit Function
        End If

    'Undo Currency scaling to obtain the true QPC frequency in Hz
        QPC_GetFrequencyHz = CDbl(m_qpcFreq) * 10000#
End Function

Private Function Get_TickCountSeconds() As Double
'
'==============================================================================
'                              GET TICKCOUNT SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the system uptime tick counter in seconds as a Double.
'
'   - Win64: uses GetTickCount64 (no practical rollover)
'   - Win32: uses GetTickCount (32-bit unsigned ms counter; wraps at 2^32 ms)
'
' WHY THIS EXISTS
'   VBA Long is signed. On 32-bit, GetTickCount returns a signed Long even
'   though the underlying counter is logically unsigned. When the high bit is
'   set, the returned Long is negative. This helper corrects that by remapping
'   the value into the unsigned 32-bit range in Double space.
'
' INPUTS
'   None.
'
' RETURNS
'   Double
'     Seconds since boot (subject to rollover on 32-bit).
'
' BEHAVIOR
'   - On Win64, uses GetTickCount64 and converts milliseconds to seconds.
'   - On Win32/pre-VBA7, uses GetTickCount and corrects signed Long values into
'     the logical unsigned 32-bit range before converting to seconds.
'
' ERROR POLICY
'   Does not raise errors.
'
' DEPENDENCIES
'   - GetTickCount
'   - GetTickCount64 (Win64/VBA7)
'   - Conditional compilation symbols: VBA7, Win64
'
' NOTES
'   - Uses only a small number of branches and avoids Currency conversion here.
'   - 4294967296# is the exact 2^32 constant expressed as a Double literal.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim tc32    As Long     'Signed 32-bit tick value from GetTickCount (Win32 only)
    Dim u32     As Double   'Unsigned 32-bit tick value mapped into Double space (Win32 only)

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
#If VBA7 Then
    #If Win64 Then
        '64-bit: GetTickCount64 returns a 64-bit millisecond counter; convert to seconds
            Get_TickCountSeconds = CDbl(GetTickCount64()) / 1000#
    #Else
        '32-bit: GetTickCount returns a signed Long; map to unsigned range [0..2^32-1]
            tc32 = GetTickCount()

            If tc32 < 0 Then
                u32 = 4294967296# + CDbl(tc32)
            Else
                u32 = CDbl(tc32)
            End If

            Get_TickCountSeconds = u32 / 1000#
    #End If
#Else
        'Pre-VBA7: 32-bit only; same signed-to-unsigned correction
            tc32 = GetTickCount()

            If tc32 < 0 Then
                u32 = 4294967296# + CDbl(tc32)
            Else
                u32 = CDbl(tc32)
            End If

            Get_TickCountSeconds = u32 / 1000#
#End If
End Function

Private Function Get_SystemTimeMs() As Long
'
'==============================================================================
'                           GET CURRENT TIME MS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current timeGetSystemTime millisecond counter
'   (MMTIME.Units).
'
' WHY THIS EXISTS
'   timeGetSystemTime returns its value through an MMTIME structure.
'   Class_Initialize configures m_Mmt.wType = 1 (TIME_MS), so m_Mmt.Units is
'   interpreted as milliseconds.
'
' INPUTS
'   None.
'
' RETURNS
'   Long
'     Milliseconds as reported by timeGetSystemTime in m_Mmt.Units.
'
' BEHAVIOR
'   - Calls timeGetSystemTime using the reusable class-level MMTIME structure.
'   - Returns m_Mmt.Units when the API call succeeds.
'   - In strict mode, raises if the API call fails.
'   - In non-strict mode, returns 0 if the API call fails.
'
' ERROR POLICY
'   - Raises in strict mode when timeGetSystemTime fails.
'   - Returns 0 in non-strict mode when timeGetSystemTime fails.
'
' DEPENDENCIES
'   - timeGetSystemTime
'   - m_Mmt
'   - m_StrictMode
'
' NOTES
'   - This is a millisecond counter source and may have rollover semantics
'     similar to other 32-bit millisecond sources.
'   - m_Mmt is a member to avoid repeated structure initialization.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long    'Return code from timeGetSystemTime

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Populate m_Mmt with the current time in the configured format
    '(wType = 1 => milliseconds).
        RC = timeGetSystemTime(m_Mmt, LenB(m_Mmt))

'------------------------------------------------------------------------------
' VALIDATE API RESULT
'------------------------------------------------------------------------------
    'Raise or return a deterministic fallback when the API call fails.
        If RC <> 0 Then
            If m_StrictMode Then
                Err.Raise vbObjectError + 1006, _
                          "cPerformanceManager.Get_SystemTimeMs", _
                          "timeGetSystemTime failed."
            Else
                Get_SystemTimeMs = 0&
                Exit Function
            End If
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return millisecond units.
        Get_SystemTimeMs = m_Mmt.Units
End Function
Private Function RolloverSeconds( _
    ByVal iMethod As Integer) _
    As Double
'
'==============================================================================
'                              ROLLOVER SECONDS
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the rollover period (in seconds) for the selected timing backend.
'
' WHY THIS EXISTS
'   Methods 2, 3, and 4 do not share the same rollover behavior on Win64:
'     - Method 2 uses GetTickCount64 on Win64 => effectively no rollover
'     - Methods 3 and 4 still behave like 32-bit millisecond counters
'
' INPUTS
'   iMethod
'     Timing backend identifier.
'
' RETURNS
'   Double
'     Rollover period in seconds for the specified backend.
'     Returns 0 when rollover correction is not applicable here.
'
' BEHAVIOR
'   - For method 2:
'       * returns 0 on Win64/VBA7 because GetTickCount64 is used
'       * returns 2^32 milliseconds expressed in seconds otherwise
'   - For methods 3 and 4:
'       * returns 2^32 milliseconds expressed in seconds
'   - For all other methods:
'       * returns 0
'
' ERROR POLICY
'   Does not raise errors.
'   Unsupported/non-rollover methods simply return 0.
'
' DEPENDENCIES
'   - Conditional compilation symbols: VBA7, Win64
'
' NOTES
'   - The rollover value 2^32 milliseconds equals 4294967296# / 1000# seconds.
'   - This helper is intentionally limited to the backends whose rollover
'     correction is handled centrally in ElapsedSeconds().
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    Select Case iMethod
        Case 2
            #If VBA7 Then
                #If Win64 Then
                    'GetTickCount64 has no practical rollover for measurement use
                        RolloverSeconds = 0#
                #Else
                    '32-bit GetTickCount wraps at 2^32 milliseconds
                        RolloverSeconds = 4294967296# / 1000#
                #End If
            #Else
                'Pre-VBA7: 32-bit GetTickCount wraps at 2^32 milliseconds
                    RolloverSeconds = 4294967296# / 1000#
            #End If
        Case 3, 4
            'timeGetTime and timeGetSystemTime behave like 32-bit millisecond counters
                RolloverSeconds = 4294967296# / 1000#
        Case Else
            'All other methods either have custom correction or no rollover logic here
                RolloverSeconds = 0#

    End Select
End Function

'
'------------------------------------------------------------------------------
'
'                 PRIVATE: TIMING CAPTURE / ALIGNMENT HELPERS
'
'------------------------------------------------------------------------------
'

Private Sub EnsureTimerResolution1ms()
'
'==============================================================================
'                             ENSURE TIME PERIOD ON
'------------------------------------------------------------------------------
' PURPOSE
'   Requests 1ms multimedia timer resolution exactly once for this instance
'   when using timeGetTime-based timing (method 3).
'
' WHY THIS EXISTS
'   timeGetTime() depends on timer-resolution behavior. This helper centralizes
'   the one-time request logic so callers do not repeatedly invoke
'   timeBeginPeriod(1) in hot paths.
'
' INPUTS
'   None.
'
' RETURNS
'   None.
'
' BEHAVIOR
'   - If 1ms timer resolution has not yet been requested by this instance,
'     calls timeBeginPeriod 1.
'   - Marks m_TimePeriodOn = True only when the request succeeds.
'   - Does nothing on subsequent calls once the request is active.
'
' ERROR POLICY
'   Silent on failure by default.
'   If stricter behavior is desired, this routine can be changed to raise when
'   timeBeginPeriod(1) fails.
'
' DEPENDENCIES
'   - timeBeginPeriod
'   - m_TimePeriodOn
'
' NOTES
'   - This routine is intentionally lightweight and guarded because it may be
'     reached from timing helpers used in repeated measurements.
'   - ResetEnvironment is responsible for balancing a successful request with
'     timeEndPeriod 1.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim RC As Long    'Return code from timeBeginPeriod

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Request 1ms timer resolution only once for this instance
        If Not m_TimePeriodOn Then
            RC = timeBeginPeriod(1)
            If RC = 0 Then
                m_TimePeriodOn = True
            End If
        End If
End Sub

Private Function Get_CurrentTimestampSeconds( _
    ByVal iMethod As Integer) _
    As Double
'
'==============================================================================
'                                  NOW TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the current timestamp for non-QPC methods as a Double in seconds.
'
' WHY THIS EXISTS
'   StartTimer and related timing helpers need a shared abstraction for
'   obtaining a consistent seconds-based timestamp for the supported non-QPC
'   backends, without duplicating dispatch logic.
'
' INPUTS
'   iMethod
'     Timing backend selector:
'       1 = VBA Timer()                         (seconds, rolls at midnight)
'       2 = GetTickCount / GetTickCount64       (seconds since boot)
'       3 = timeGetTime                         (seconds since boot; 32-bit rollover)
'       4 = timeGetSystemTime (MMTIME.Units)    (seconds from millisecond source)
'       6 = VBA Now() * 86400                   (seconds; wall clock)
'
' RETURNS
'   Double
'     Timestamp in seconds for the requested backend.
'
' BEHAVIOR
'   - Dispatches to the requested non-QPC timing backend.
'   - Returns all supported backends in seconds.
'   - Ensures 1ms timer resolution is enabled before using method 3.
'   - Falls back to Timer for unexpected method values.
'
' ERROR POLICY
'   Does not raise errors.
'   Invalid method values are treated as a defensive fallback to Timer.
'
' DEPENDENCIES
'   - Get_TickCountSeconds
'   - EnsureTimerResolution1ms
'   - timeGetTime
'   - Get_SystemTimeMs
'   - Timer
'   - Now
'
' NOTES
'   - Method 5 (QPC) is intentionally excluded and handled elsewhere.
'   - This helper assumes callers have normally already validated the method.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Dispatch to the requested timing backend and return a seconds-based timestamp.
        Select Case iMethod
            Case 1
                Get_CurrentTimestampSeconds = Timer

            Case 2
                Get_CurrentTimestampSeconds = Get_TickCountSeconds()

            Case 3
                EnsureTimerResolution1ms
                Get_CurrentTimestampSeconds = timeGetTime() / 1000#

            Case 4
                Get_CurrentTimestampSeconds = Get_SystemTimeMs() / 1000#

            Case 6
                Get_CurrentTimestampSeconds = Now() * 86400#

            Case Else
                Get_CurrentTimestampSeconds = Timer
        End Select
End Function

Private Function Get_NextTimestampSeconds( _
    ByVal iMethod As Integer) _
    As Double
'
'==============================================================================
'                                   NEXT TICK
'------------------------------------------------------------------------------
' PURPOSE
'   Aligns the start time to the next observable "tick" for the selected timing
'   backend.
'
' WHY THIS EXISTS
'   For certain micro-benchmark scenarios, aligning the start capture to the
'   next observable counter change can reduce start-time quantization jitter.
'   This helper centralizes that alignment logic for the supported non-QPC
'   backends.
'
' INPUTS
'   iMethod
'     Timing backend selector:
'       1 = VBA Timer()                         (seconds, rolls at midnight)
'       2 = GetTickCount / GetTickCount64       (seconds since boot)
'       3 = timeGetTime                         (ms counter; ensure 1ms resolution)
'       4 = timeGetSystemTime (MMTIME.Units)    (ms counter via m_Mmt.Units)
'       6 = VBA Now() * 86400                   (seconds; wall clock)
'
' RETURNS
'   Double
'     A seconds-based timestamp aligned to the next observed change in the
'     underlying counter, where meaningful.
'
' BEHAVIOR
'   - Spins in a tight loop until the selected counter advances.
'   - Returns the first observed value after the change.
'   - Converts millisecond-based sources to seconds before returning.
'   - Returns current wall-clock seconds for method 6 because no meaningful
'     tick-alignment abstraction exists there.
'   - Falls back to Get_CurrentTimestampSeconds(iMethod) for unexpected method values.
'
' ERROR POLICY
'   Does not raise errors.
'   Unexpected method values fall back defensively to Get_CurrentTimestampSeconds(iMethod).
'
' DEPENDENCIES
'   - Get_TickCountSeconds
'   - EnsureTimerResolution1ms
'   - timeGetTime
'   - Get_SystemTimeMs
'   - Get_CurrentTimestampSeconds
'   - Timer
'   - Now
'
' NOTES
'   - This routine uses tight polling loops and is therefore intentionally more
'     expensive than a simple timestamp read.
'   - It is best suited to benchmark alignment scenarios, not general-purpose
'     production timing.
'   - Method 1 alignment can still be coarse because Timer() granularity is
'     platform-dependent.
'
' UPDATED
'   2026-03-28
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim A As Double   'Initial observed timestamp for the selected backend
    Dim B As Double   'Next observed timestamp; loop exits when B <> A

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Dispatch by timing backend and spin until the counter advances
        Select Case iMethod

            Case 1
                'Timer() returns seconds since midnight
                    A = Timer
                    Do
                        B = Timer
                    Loop Until B <> A
                    Get_NextTimestampSeconds = B

            Case 2
                'Get_TickCountSeconds returns seconds since boot
                    A = Get_TickCountSeconds()
                    Do
                        B = Get_TickCountSeconds()
                    Loop Until B <> A
                    Get_NextTimestampSeconds = B

            Case 3
                'timeGetTime returns milliseconds; ensure resolution then convert to seconds
                    EnsureTimerResolution1ms
                    A = CDbl(timeGetTime())
                    Do
                        B = CDbl(timeGetTime())
                    Loop Until B <> A
                    Get_NextTimestampSeconds = B / 1000#

            Case 4
                'timeGetSystemTime via Get_SystemTimeMs returns milliseconds; convert to seconds
                    A = CDbl(Get_SystemTimeMs())
                    Do
                        B = CDbl(Get_SystemTimeMs())
                    Loop Until B <> A
                    Get_NextTimestampSeconds = B / 1000#

            Case 6
                'Now() is wall-clock time; no meaningful tick alignment is available
                    Get_NextTimestampSeconds = Now() * 86400#

            Case Else
                'Fallback: return a current timestamp for the requested method
                    Get_NextTimestampSeconds = Get_CurrentTimestampSeconds(iMethod)

        End Select
End Function

Private Function QPC_Get_NextTimestampTick() As Currency
'
'==============================================================================
'                                 NEXT TICK QPC
'------------------------------------------------------------------------------
' PURPOSE
'   Aligns to the next QueryPerformanceCounter tick
'
' WHY THIS EXISTS
'   In some micro-benchmark scenarios, aligning the QPC start capture to the
'   next observed counter increment can reduce start-time quantization jitter
'
' INPUTS
'   None.
'
' RETURNS
'   Currency
'     The first observed QPC counter value strictly greater than the initial
'     read
'
' BEHAVIOR
'   - Captures an initial QPC value
'   - Spins in a tight loop until a strictly greater QPC value is observed
'   - Returns that aligned raw QPC tick
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   - QueryPerformanceCounter
'
' NOTES
'   - Tight polling loop by design; do not use in ordinary production paths
'   - QPC is high resolution, so this loop typically exits quickly, but it still
'     adds overhead and should only be used when explicitly requested
'
' UPDATED
'   2026-04-15
'==============================================================================

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim T1  As Currency     'Initial counter tick
    Dim T2  As Currency     'Next counter tick; loop exits when T2 > T1

'------------------------------------------------------------------------------
' ROUTINE
'------------------------------------------------------------------------------
    'Read initial tick
        QueryPerformanceCounter T1

    'Spin until the counter advances strictly beyond the initial read
        Do
            QueryPerformanceCounter T2
        Loop Until T2 > T1

    'Return the aligned tick
        QPC_Get_NextTimestampTick = T2
        
End Function

'
'------------------------------------------------------------------------------
'
'                PRIVATE: CHECKPOINT / REPORT HELPER ROUTINES
'
'------------------------------------------------------------------------------
'

Private Sub Checkpoint_ClearState()
'
'==============================================================================
'                           CHECKPOINT CLEAR STATE
'------------------------------------------------------------------------------
' PURPOSE
'   Resets the structured checkpoint state for a fresh timing session
'
' WHY THIS EXISTS
'   Checkpoints and run label belong to a timing session. Starting a new session
'   should therefore reset the structured reporting state deterministically.
'
'   This same helper is also used by ClearCheckpoints(), so that explicit
'   checkpoint clearing resets the full structured report state, including the
'   current run label.
'
' INPUTS
'   None
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Clears the current run label
'   - Clears the checkpoint count
'   - Clears the last checkpoint marker
'   - Reinitializes checkpoint storage to a small first capacity
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' RESET STATE
'------------------------------------------------------------------------------
    'Clear the current run label
        m_RunLabel = vbNullString
    'Clear the checkpoint counter
        m_CheckpointCount = 0
    'Clear the last checkpoint elapsed-seconds marker
        m_LastCheckpointS = 0#
    'Reinitialize checkpoint storage with a small first capacity
        ReDim m_Checkpoints(1 To 1)

End Sub

Private Sub Checkpoint_EnsureCapacity( _
    ByVal RequiredCount As Long)
'
'==============================================================================
'                        CHECKPOINT ENSURE CAPACITY
'------------------------------------------------------------------------------
' PURPOSE
'   Ensures that checkpoint storage can hold the requested number of rows
'
' WHY THIS EXISTS
'   Checkpoints are appended dynamically. This helper centralizes array growth
'   logic and avoids repeated ReDim Preserve on every single row
'
' INPUTS
'   RequiredCount
'     Minimum required checkpoint capacity
'
' RETURNS
'   None
'
' BEHAVIOR
'   - Leaves storage unchanged when current capacity is already sufficient
'   - Otherwise grows storage using a simple doubling strategy
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
    Dim CurrentCapacity     As Long                 'Current checkpoint storage capacity
    Dim NewCapacity         As Long                 'New checkpoint storage capacity

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Resolve the current checkpoint storage capacity
        CurrentCapacity = UBound(m_Checkpoints)

'------------------------------------------------------------------------------
' EXIT EARLY WHEN CAPACITY IS ALREADY SUFFICIENT
'------------------------------------------------------------------------------
    'Leave storage unchanged when the requested capacity is already available
        If RequiredCount <= CurrentCapacity Then
            Exit Sub
        End If

'------------------------------------------------------------------------------
' COMPUTE NEW CAPACITY
'------------------------------------------------------------------------------
    'Start from a simple doubling growth rule
        NewCapacity = CurrentCapacity * 2
    'Ensure the new capacity is never smaller than the required count
        If NewCapacity < RequiredCount Then
            NewCapacity = RequiredCount
        End If

'------------------------------------------------------------------------------
' GROW STORAGE
'------------------------------------------------------------------------------
    'Resize checkpoint storage while preserving existing rows
        ReDim Preserve m_Checkpoints(1 To NewCapacity)

End Sub

Private Function Checkpoint_DefaultName( _
    ByVal Seq As Long) _
    As String
'
'==============================================================================
'                         CHECKPOINT DEFAULT NAME
'------------------------------------------------------------------------------
' PURPOSE
'   Returns a generated fallback checkpoint name for blank checkpoint labels
'
' WHY THIS EXISTS
'   Structured reports are easier to read when every checkpoint row has a
'   non-blank label, even when the caller omitted one
'
' INPUTS
'   Seq
'     Sequence number that will be embedded in the generated label
'
' RETURNS
'   String
'     Generated checkpoint label
'
' ERROR POLICY
'   Raises errors normally
'
' UPDATED
'   2026-04-18
'==============================================================================

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return a generated checkpoint label
        Checkpoint_DefaultName = "Checkpoint " & CStr(Seq)

End Function

