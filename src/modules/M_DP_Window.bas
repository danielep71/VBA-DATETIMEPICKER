Attribute VB_Name = "M_DP_Window"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_Window
'------------------------------------------------------------------------------
' PURPOSE
'   Owns the production Windows API declarations and the native window
'   operations of the picker form
'
' WHY THIS EXISTS
'   Native declarations and their 32-bit, 64-bit and Mac variants are the
'   riskiest code in the project. One owner keeps them in a single place
'   and keeps them out of the settings and write-back code (#56)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - platform predicates: M_Platform_CanUseWinAPI and
'       M_Platform_ShouldUseWinAPI
'     - borderless styling as one transaction: M_Window_RemoveTitleBar
'     - native drag, handle lookup and positioning near the mouse
'     - the test-only window-style fault seam
'
' ERROR POLICY
'   Each procedure documents its own policy. Native styling reports its
'   outcome through DP_WindowStyleResult instead of raising
'
' DEPENDENCIES
'   user32 and kernel32 on Windows
'   M_DatePicker (DP_WindowStyleResult)
'   M_DP_Core (gDP_UseWinAPI)
'
' NOTES
'   Every Declare is Private, so no other module calls the API directly. The
'   regression harness keeps its own independent declarations to read real
'   window state
'
'   DP_WindowStyleResult stays in the M_DatePicker facade because #25
'   classifies it as a supported result type
'
'   Option Private Module keeps the helpers out of other VBA projects and out
'   of the macro dialog
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' PRIVATE TYPES
'------------------------------------------------------------------------------
    Private Type POINTAPI
        X As Long                   'Mouse or screen X position in pixels
        Y As Long                   'Mouse or screen Y position in pixels
    End Type

    Private Type RECT
        Left As Long                'Rectangle left coordinate in pixels
        Top As Long                 'Rectangle top coordinate in pixels
        Right As Long               'Rectangle right coordinate in pixels
        Bottom As Long              'Rectangle bottom coordinate in pixels
    End Type

    Private Type MONITORINFO
        cbSize As Long              'Structure size in bytes
        rcMonitor As RECT           'Full monitor rectangle in pixels
        rcWork As RECT              'Work-area rectangle in pixels
        dwFlags As Long             'Monitor flags
    End Type

'------------------------------------------------------------------------------
' WINDOWS API DECLARATIONS
'------------------------------------------------------------------------------
    #If Mac Then
    #Else
        #If VBA7 Then
            Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
                ByRef lpPoint As POINTAPI) As Long

            Private Declare PtrSafe Function MonitorFromRect Lib "user32" ( _
                ByRef lprc As RECT, _
                ByVal dwFlags As Long) As LongPtr

            Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
                ByVal hMonitor As LongPtr, _
                ByRef lpmi As MONITORINFO) As Long

            Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As LongPtr

            Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As LongPtr, _
                ByRef lpRect As RECT) As Long

            Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
                ByVal hWnd As LongPtr, _
                ByVal hWndInsertAfter As LongPtr, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal uFlags As Long) As Long

            Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
                ByVal hWnd As LongPtr) As Long

            Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long

            Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As LongPtr, _
                ByVal wMsg As Long, _
                ByVal wParam As LongPtr, _
                ByVal lParam As LongPtr) As LongPtr

            #If Win64 Then
                Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr

                Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As LongPtr) As LongPtr
            #Else
                Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long) As LongPtr

                Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" ( _
                    ByVal hWnd As LongPtr, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As LongPtr) As LongPtr
            #End If
        #Else
            Private Declare Function GetCursorPos Lib "user32" ( _
                ByRef lpPoint As POINTAPI) As Long

            Private Declare Function MonitorFromRect Lib "user32" ( _
                ByRef lprc As RECT, _
                ByVal dwFlags As Long) As Long

            Private Declare Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" ( _
                ByVal hMonitor As Long, _
                ByRef lpmi As MONITORINFO) As Long

            Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                ByVal lpClassName As String, _
                ByVal lpWindowName As String) As Long

            Private Declare Function GetWindowRect Lib "user32" ( _
                ByVal hWnd As Long, _
                ByRef lpRect As RECT) As Long

            Private Declare Function SetWindowPos Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cx As Long, _
                ByVal cy As Long, _
                ByVal uFlags As Long) As Long

            Private Declare Function DrawMenuBar Lib "user32" ( _
                ByVal hWnd As Long) As Long

            Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long

            Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long

            Private Declare Function ReleaseCapture Lib "user32" () As Long

            Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hWnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
        #End If
    #End If

    #If Mac Then
    #Else
        #If VBA7 Then

            Private Declare PtrSafe Sub SetLastError Lib "kernel32" ( _
                ByVal dwErrCode As Long)
        #Else

            Private Declare Sub SetLastError Lib "kernel32" ( _
                ByVal dwErrCode As Long)
        #End If
    #End If

'------------------------------------------------------------------------------
' PRIVATE CONSTANTS
'------------------------------------------------------------------------------
    Private Const GWL_STYLE                As Long = -16        'Window style index
    Private Const MONITOR_DEFAULTTONEAREST As Long = 2          'Monitor nearest flag
    Private Const SWP_NOSIZE               As Long = &H1        'SetWindowPos no-size flag
    Private Const SWP_NOMOVE               As Long = &H2        'SetWindowPos no-move flag
    Private Const SWP_NOZORDER             As Long = &H4        'SetWindowPos no-z-order flag
    Private Const SWP_NOACTIVATE           As Long = &H10       'SetWindowPos no-activate flag
    Private Const SWP_FRAMECHANGED         As Long = &H20       'SetWindowPos frame-changed flag
    Private Const SWP_SHOWWINDOW           As Long = &H40       'SetWindowPos show-window flag

    Private Const WS_CAPTION               As Long = &HC00000   'Window caption style flag
    Private Const DP_WM_NCLBUTTONDOWN      As Long = &HA1       'Non-client left-button down message
    Private Const DP_HTCAPTION             As Long = 2          'Title-bar hit-test code

'------------------------------------------------------------------------------
' TEST FAULT INJECTION STATE
'------------------------------------------------------------------------------
    'Armed failure point for the next borderless-styling call, zero when disarmed
    Private mDP_TestWindowPrimaryFailure    As Long
    'Armed rollback failure point for that same call, zero when disarmed
    Private mDP_TestWindowRollbackFailure   As Long


'
'------------------------------------------------------------------------------
'
'                               WINDOW SERVICES
'
'------------------------------------------------------------------------------
'

Public Function M_Platform_CanUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM CAN USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether the current platform can use Windows API calls
'
' WHY THIS EXISTS
'   Settings normalization and platform-safe startup logic need to know whether
'   WinAPI calls are technically available before applying user configuration
'
' INPUTS
'   None
'
' RETURNS
'   True when the project is running on Windows
'   False when the project is running on Mac
'
' BEHAVIOR
'   Uses conditional compilation to return a platform-safe Boolean without
'   calling any native API
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   VBA conditional compilation constant Mac
'
' NOTES
'   This routine answers platform capability only
'
'   It deliberately does not inspect gDP_UseWinAPI
'
'   Runtime policy belongs in M_Platform_ShouldUseWinAPI
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

#If Mac Then
    'Return False when running on Mac
        M_Platform_CanUseWinAPI = False
#Else
    'Return True when running on Windows
        M_Platform_CanUseWinAPI = True
#End If

End Function

Public Function M_Platform_ShouldUseWinAPI() As Boolean

'
'------------------------------------------------------------------------------
'                           PLATFORM SHOULD USE WINAPI
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether optional DatePicker WinAPI styling should be enabled
'
' WHY THIS EXISTS
'   Some DatePicker visual behavior, especially native title-bar removal, should
'   be controlled by both:
'     - platform capability
'     - user / project configuration
'
'   This helper provides the effective runtime policy for optional WinAPI styling
'   behavior. It is not the general WinAPI capability check.
'
' INPUTS
'   None
'
' RETURNS
'   True when WinAPI is available on the current platform and styling is enabled
'   False otherwise
'
' BEHAVIOR
'   Returns True only when both conditions are satisfied:
'     - the current platform can use WinAPI
'     - gDP_UseWinAPI is True
'
' ERROR POLICY
'   Does not raise errors directly
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   gDP_UseWinAPI
'
' NOTES
'   Use this helper for optional styling features such as title-bar removal.
'
'   Use M_Platform_CanUseWinAPI for capability-only operations that should remain
'   available on Windows even when optional WinAPI styling is disabled.
'
'   Do not call M_Settings_EnsureLoaded from this helper.
'
'   Settings load / save routines already use M_Platform_CanUseWinAPI to avoid
'   circular policy dependencies while normalizing gDP_UseWinAPI.
'
' UPDATED
'   2026-05-10
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' RETURN POLICY DECISION
'------------------------------------------------------------------------------
    'Return whether optional WinAPI-dependent behavior should be used
        M_Platform_ShouldUseWinAPI = (M_Platform_CanUseWinAPI And gDP_UseWinAPI)

End Function

#If VBA7 Then
Private Sub M_Window_RollbackStyle(ByVal hWndForm As LongPtr, ByVal OriginalStyle As LongPtr, ByVal FailedStep As String, ByVal FailedApiError As Long, ByVal RollbackFault As Long, ByRef Result As DP_WindowStyleResult)
#Else
Private Sub M_Window_RollbackStyle(ByVal hWndForm As Long, ByVal OriginalStyle As Long, ByVal FailedStep As String, ByVal FailedApiError As Long, ByVal RollbackFault As Long, ByRef Result As DP_WindowStyleResult)
#End If

'
'------------------------------------------------------------------------------
'                        ROLLBACK WINDOW STYLE
'------------------------------------------------------------------------------
' PURPOSE
'   Restores the window style captured before a committed change, after a later
'   step in the same operation failed
'
' WHY THIS EXISTS
'   Clearing WS_CAPTION succeeds long before the frame is refreshed. A failure in
'   between leaves a window whose style and frame disagree, which no later call
'   detects and nothing repairs
'
' INPUTS
'   hWndForm
'     Native window handle being restored
'
'   OriginalStyle
'     Style read before the change
'
'   FailedStep
'     Step whose failure triggered this rollback
'
'   FailedApiError
'     WinAPI error behind that failure
'
'   RollbackFault
'     Regression fault-injection point, zero in normal use
'
'   Result
'     Outcome to populate, modified in place
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Restores the original style and refreshes the frame
'
'   Reports RolledBack when the window is back as it was found, and
'   RecoveryRequired when it is not
'
' ERROR POLICY
'   Does not raise. A rollback that fails is reported, not thrown, because the
'   caller needs the original failure as well as this one
'
' DEPENDENCIES
'   SetWindowLong / SetWindowLongPtr
'   SetWindowPos
'
' NOTES
'   The original failure is what the caller has to act on, so FailedStep and
'   LastApiError describe that failure and not this rollback. A rollback that
'   fails is reported through RecoveryRequired
'
'   A restored style still needs a frame refresh. Restoring the bits without it
'   reproduces the same half-applied state in the opposite direction
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_RollbackStyle"

    Const FAULT_ROLLBACK_STYLE  As Long = 1     'Fail the rollback style restore
    Const FAULT_ROLLBACK_FRAME  As Long = 2     'Fail the rollback frame refresh

    #If VBA7 Then
        Dim RestoreResult       As LongPtr       'Previous style returned by the restore
    #Else
        Dim RestoreResult       As Long          'Previous style returned by the restore
    #End If

    Dim WindowFlags             As Long          'SetWindowPos flags
    Dim ApiResult               As Long          'Generic WinAPI Boolean-style result
    Dim LastApiError            As Long          'WinAPI last-error code

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'A rollback must never raise into the operation it is recovering
        On Error Resume Next
    'Report the failure that caused the rollback, not the rollback itself
        Result.FailedStep = FailedStep
        Result.LastApiError = FailedApiError

'------------------------------------------------------------------------------
' RESTORE ORIGINAL STYLE
'------------------------------------------------------------------------------
    'Clear the WinAPI last-error state before restoring
        SetLastError 0
    'Skip the restore entirely when a regression test has armed this point.
    'Restoring and then reporting failure would leave the window back in its
    'original state while the result claimed it could not be recovered
        If RollbackFault = FAULT_ROLLBACK_STYLE Then
            Debug.Print PROC_NAME & _
                " | Step=Restore original style" & _
                " | Injected failure, native restore skipped"
            Result.RecoveryRequired = True
            Err.Clear
            Exit Sub
        End If
    #If VBA7 Then
        'Restore the style captured before the change
            RestoreResult = SetWindowLongPtr(hWndForm, GWL_STYLE, OriginalStyle)
    #Else
        'Restore the style captured before the change
            RestoreResult = SetWindowLong(hWndForm, GWL_STYLE, OriginalStyle)
    #End If
    'Report an unrecoverable window when the style cannot be put back
        If RestoreResult = 0 Then
            LastApiError = Err.LastDllError
            If LastApiError <> 0 Then
                Debug.Print PROC_NAME & _
                    " | Step=Restore original style" & _
                    " | Api=SetWindowLong" & _
                    " | LastError=" & VBA.CStr(LastApiError)
                Result.RecoveryRequired = True
                Err.Clear
                Exit Sub
            End If
        End If

'------------------------------------------------------------------------------
' REFRESH RESTORED FRAME
'------------------------------------------------------------------------------
    'Build non-client refresh flags
        WindowFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
            SWP_NOACTIVATE Or SWP_FRAMECHANGED
    'Clear the WinAPI last-error state before SetWindowPos
        SetLastError 0
    'Skip the refresh entirely when a regression test has armed this point, so the
    'injected state really is a restored style whose frame was never refreshed
        If RollbackFault = FAULT_ROLLBACK_FRAME Then
            ApiResult = 0
        Else
            'Make the restored style visible in the frame
                ApiResult = SetWindowPos(hWndForm, 0, 0, 0, 0, 0, WindowFlags)
        End If
    'Report an unrecoverable window when the restored frame cannot be refreshed
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=Refresh restored frame" & _
                " | Api=SetWindowPos" & _
                " | LastError=" & VBA.CStr(LastApiError)
            Result.RecoveryRequired = True
            Err.Clear
            Exit Sub
        End If

'------------------------------------------------------------------------------
' REPORT ROLLBACK
'------------------------------------------------------------------------------
    'The window is back as it was found, so the native title bar is usable
        Result.RolledBack = True
    'Clear any suppressed rollback error
        Err.Clear

End Sub

Public Sub M_Window_Test_SetFaultInjection( _
    ByVal PrimaryFailurePoint As Long, _
    Optional ByVal RollbackFailurePoint As Long = 0)

'
'------------------------------------------------------------------------------
'                    ARM WINDOW-STYLE FAULT INJECTION
'------------------------------------------------------------------------------
' PURPOSE
'   Arms a single forced failure inside the next borderless-styling call
'
' WHY THIS EXISTS
'   The failure paths this seam covers cannot be produced on demand. A window
'   that refuses SetWindowPos after accepting a style write is not something a
'   test can arrange, and those are exactly the paths where a partially applied
'   style is possible
'
'   The regression module is a separate VBA module and cannot assign private
'   state in M_DP_Window, so the setter has to be technically Public
'
' INPUTS
'   PrimaryFailurePoint
'     1 style read, 2 style write, 3 frame refresh, 4 redraw. Zero disarms
'
'   RollbackFailurePoint
'     1 rollback style restore, 2 rollback frame refresh. Zero disarms
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Stores the requested failure points for the next call only
'
' ERROR POLICY
'   Does not raise
'
' DEPENDENCIES
'   None
'
' NOTES
'   THIS IS INTERNAL TEST INFRASTRUCTURE. It is not supported DatePicker API,
'   is classified internal under #25, and must not appear in the README public
'   API table
'
'   Injection is one-shot. M_Window_RemoveTitleBar copies these values and
'   clears them before touching the window, so an armed test cannot leak into a
'   later real call
'
'   The required first argument keeps this out of the Alt+F8 macro list
'
'   Failure-point numbers are duplicated as private constants in the regression
'   module rather than shared through a public enum, so this seam does not
'   enlarge the public surface. The two lists must be changed together
'
'   No state is persisted. Nothing is written to the registry, a workbook, or
'   any Excel object
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ARM INJECTION
'------------------------------------------------------------------------------
    'Arm the requested primary failure point for the next call only
        mDP_TestWindowPrimaryFailure = PrimaryFailurePoint
    'Arm the requested rollback failure point for the next call only
        mDP_TestWindowRollbackFailure = RollbackFailurePoint

End Sub

Public Function M_Window_RemoveTitleBar(ByVal Frm As Object) As DP_WindowStyleResult

'
'------------------------------------------------------------------------------
'                      REMOVE USERFORM TITLE BAR
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the borderless DatePicker window style, as one transaction that
'   either completes or leaves the window as it was found
'
' WHY THIS EXISTS
'   Clearing WS_CAPTION and refreshing the frame are two separate native
'   operations. The first can succeed and the second fail, leaving a window whose
'   style says borderless and whose frame still shows a title bar
'
'   Reporting through Debug.Print alone made that state invisible: complete
'   success, a safe abort before any change, and a half-applied style were
'   indistinguishable to the caller
'
' INPUTS
'   Frm
'     UserForm whose native window should lose its title bar
'
' RETURNS
'   DP_WindowStyleResult describing the outcome
'
'     Attempted           the native style was actually touched
'     Applied             the borderless style is fully in effect
'     Committed           the style write succeeded
'     RolledBack          the original style was restored after a later failure
'     RecoveryRequired    the window is in no known good state
'     FailedStep          the step that failed
'     LastApiError        the WinAPI error behind it
'
' BEHAVIOR
'   Captures the original style, clears WS_CAPTION, and refreshes the frame
'
'   A failure before the style write leaves the window untouched
'
'   A failure after the style write restores the original style and refreshes the
'   frame again, so the form falls back to its native title bar
'
'   A rollback that itself fails reports RecoveryRequired, so the caller can
'   unload and rebuild the form rather than continue against a window in an
'   unknown state
'
' ERROR POLICY
'   Does not raise. Every outcome is reported through the returned result
'
'   Diagnostics are still written to the Immediate Window, naming the failing
'   step and the WinAPI error. They supplement the result, they do not replace it
'
' DEPENDENCIES
'   M_Platform_ShouldUseWinAPI
'   M_Window_GetUserFormHwnd
'   GetWindowLong / SetWindowLong
'   SetWindowPos
'   DrawMenuBar
'
' NOTES
'   The original style is kept in its own variable. The masked value goes
'   somewhere else, because rollback is impossible once the two share storage
'
'   There is deliberately no shortcut for a caption bit that is already clear.
'   The bit being clear proves a previous style write succeeded; it proves
'   nothing about the frame refresh that should have followed. Skipping the
'   refresh on that basis would make a half-applied window permanently
'   unrepairable, because every retry would see the bit clear and report success
'
'   Repeating the refresh on an already-borderless window is harmless, which is
'   what makes a second call a valid recovery
'
'   The SetWindowLong zero return is ambiguous: it also means the previous style
'   was zero. Err.LastDllError disambiguates, and that check is preserved
'
'   This is a Function so callers can inspect the outcome, but bare-call syntax
'   still compiles. Both UF_DatePicker call sites are unchanged
'
'   Fault injection is consumed one-shot at entry. See
'   M_Window_Test_SetFaultInjection
'
'   An injected failure skips the native call it is failing. Performing the call
'   and then overwriting its result would leave the window in the state of a
'   success while the result described a failure, which is the opposite of what
'   these paths exist to reproduce
'
' UPDATED
'   2026-08-23
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_RemoveTitleBar"

    Const FAULT_STYLE_READ      As Long = 1     'Fail the style read
    Const FAULT_STYLE_WRITE     As Long = 2     'Fail the style write
    Const FAULT_SET_WINDOW_POS  As Long = 3     'Fail the frame refresh after commit
    Const FAULT_DRAW_MENU_BAR   As Long = 4     'Fail the redraw after commit
    Const FAULT_ROLLBACK_STYLE  As Long = 1     'Fail the rollback style restore
    Const FAULT_ROLLBACK_FRAME  As Long = 2     'Fail the rollback frame refresh

    #If VBA7 Then
        Dim hWndForm            As LongPtr       'UserForm window handle
        Dim OriginalStyle       As LongPtr       'Window style as found
        Dim WindowStyle         As LongPtr       'Window style being written
        Dim SetStyleResult      As LongPtr       'Previous window style returned by SetWindowLongPtr
    #Else
        Dim hWndForm            As Long          'UserForm window handle
        Dim OriginalStyle       As Long          'Window style as found
        Dim WindowStyle         As Long          'Window style being written
        Dim SetStyleResult      As Long          'Previous window style returned by SetWindowLong
    #End If

    Dim Result                  As DP_WindowStyleResult  'Structured outcome
    Dim PrimaryFault            As Long          'Armed primary failure point
    Dim RollbackFault           As Long          'Armed rollback failure point
    Dim WindowFlags             As Long          'SetWindowPos flags
    Dim ApiResult               As Long          'Generic WinAPI Boolean-style result
    Dim LastApiError            As Long          'WinAPI last-error code
    Dim HandlerStep             As String        'Current handler step for diagnostics
    Dim ErrorNumber             As Long          'Captured VBA error number
    Dim ErrorDescription        As String        'Captured VBA error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress borderless styling errors through the local cleanup path
        On Error GoTo CleanExit
    'Initialize diagnostic step
        HandlerStep = "Initialize"
    'Consume any armed fault injection immediately, so a test call cannot leave
    'this routine poisoned for a later real one
        PrimaryFault = mDP_TestWindowPrimaryFailure
        RollbackFault = mDP_TestWindowRollbackFailure
        mDP_TestWindowPrimaryFailure = 0
        mDP_TestWindowRollbackFailure = 0

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Validate input"
    'Exit if no form was supplied
        If Frm Is Nothing Then GoTo CleanExit

'------------------------------------------------------------------------------
' CHECK WINAPI POLICY
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Check WinAPI policy"
    'Leave the native title bar in place when WinAPI use is disabled
        If Not M_Platform_ShouldUseWinAPI Then GoTo CleanExit

'------------------------------------------------------------------------------
' RESOLVE FORM HANDLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Resolve form handle"
    'Resolve the native window behind the UserForm
        hWndForm = M_Window_GetUserFormHwnd(Frm)
    'Exit when the window cannot be resolved
        If hWndForm = 0 Then GoTo CleanExit

'------------------------------------------------------------------------------
' READ WINDOW STYLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Read window style"
    'Clear the WinAPI last-error state before reading
        SetLastError 0
    'Skip the read entirely when a regression test has armed this point, so the
    'injected failure reproduces the native state of a real failure rather than
    'overwriting the result of a call that already succeeded
        If PrimaryFault = FAULT_STYLE_READ Then
            OriginalStyle = 0
        Else
            #If VBA7 Then
                'Read the current window style
                    OriginalStyle = GetWindowLongPtr(hWndForm, GWL_STYLE)
            #Else
                'Read the current window style
                    OriginalStyle = GetWindowLong(hWndForm, GWL_STYLE)
            #End If
        End If
    'Abort before any change when the style cannot be read
        If OriginalStyle = 0 Then
            LastApiError = Err.LastDllError
            Result.FailedStep = HandlerStep
            Result.LastApiError = LastApiError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=GetWindowLong" & _
                " | LastError=" & VBA.CStr(LastApiError)
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' WRITE WINDOW STYLE
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Write window style"
    'Build the borderless style without losing the original
        WindowStyle = (OriginalStyle And Not WS_CAPTION)
    'Record that the native style path was entered. The write below either takes
    'effect or does not, and Committed is what distinguishes the two
        Result.Attempted = True
    'Clear the WinAPI last-error state before writing
        SetLastError 0
    'Skip the write entirely when a regression test has armed this point. Writing
    'and then reporting failure would leave the window borderless while the result
    'said nothing had happened, and no rollback would run
        If PrimaryFault = FAULT_STYLE_WRITE Then
            SetStyleResult = 0
            Result.FailedStep = HandlerStep
            Result.LastApiError = 0
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=SetWindowLong" & _
                " | Injected failure, native write skipped"
            GoTo CleanExit
        End If
    #If VBA7 Then
        'Write the updated window style
            SetStyleResult = SetWindowLongPtr(hWndForm, GWL_STYLE, WindowStyle)
    #Else
        'Write the updated window style
            SetStyleResult = SetWindowLong(hWndForm, GWL_STYLE, WindowStyle)
    #End If
    'Diagnose SetWindowLong failure when return is zero and LastError is non-zero
        If SetStyleResult = 0 Then
            LastApiError = Err.LastDllError
            If LastApiError <> 0 Then
                Result.FailedStep = HandlerStep
                Result.LastApiError = LastApiError
                Debug.Print PROC_NAME & _
                    " | Step=" & HandlerStep & _
                    " | Api=SetWindowLong" & _
                    " | LastError=" & VBA.CStr(LastApiError)
                GoTo CleanExit
            End If
        End If
    'The style is now committed. Everything after this point must recover
        Result.Committed = True

'------------------------------------------------------------------------------
' REFRESH NON-CLIENT FRAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Refresh non-client frame"
    'Build non-client refresh flags
        WindowFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or _
            SWP_NOACTIVATE Or SWP_FRAMECHANGED
    'Clear the WinAPI last-error state before SetWindowPos
        SetLastError 0
    'Skip the refresh entirely when a regression test has armed this point, so the
    'window really is left with a committed style and an unrefreshed frame
        If PrimaryFault = FAULT_SET_WINDOW_POS Then
            ApiResult = 0
        Else
            'Force Windows to recalculate the frame
                ApiResult = SetWindowPos(hWndForm, 0, 0, 0, 0, 0, WindowFlags)
        End If
    'Recover when the frame could not be refreshed after the style was committed
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=SetWindowPos" & _
                " | LastError=" & VBA.CStr(LastApiError)
            M_Window_RollbackStyle hWndForm, OriginalStyle, HandlerStep, _
                LastApiError, RollbackFault, Result
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' REDRAW FRAME
'------------------------------------------------------------------------------
    'Track the current handler step
        HandlerStep = "Redraw frame"
    'Clear the WinAPI last-error state before DrawMenuBar
        SetLastError 0
    'Skip the redraw entirely when a regression test has armed this point
        If PrimaryFault = FAULT_DRAW_MENU_BAR Then
            ApiResult = 0
        Else
            'Redraw menu bar and non-client elements
                ApiResult = DrawMenuBar(hWndForm)
        End If
    'Recover when the redraw could not be completed after the style was committed
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Api=DrawMenuBar" & _
                " | LastError=" & VBA.CStr(LastApiError)
            M_Window_RollbackStyle hWndForm, OriginalStyle, HandlerStep, _
                LastApiError, RollbackFault, Result
            GoTo CleanExit
        End If

'------------------------------------------------------------------------------
' REPORT SUCCESS
'------------------------------------------------------------------------------
    'The borderless style is fully in effect
        Result.Applied = True

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Capture any suppressed VBA error number
        ErrorNumber = Err.Number
    'Capture any suppressed VBA error description
        ErrorDescription = Err.Description
    'Write diagnostics only when VBA raised during borderless styling
        If ErrorNumber <> 0 Then
            If VBA.LenB(Result.FailedStep) = 0 Then Result.FailedStep = HandlerStep
            Debug.Print PROC_NAME & _
                " | Step=" & HandlerStep & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'A committed style that was neither applied nor rolled back needs recovery
        If Result.Committed And Not Result.Applied And Not Result.RolledBack Then
            Result.RecoveryRequired = True
        End If
    'Publish the outcome
        M_Window_RemoveTitleBar = Result
    'Clear any suppressed styling error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function

Public Sub M_Window_BeginUserFormDrag(ByVal TargetForm As Object)

'
'------------------------------------------------------------------------------
'                       BEGIN USERFORM DRAG
'------------------------------------------------------------------------------
' PURPOSE
'   Starts native Windows drag movement for a borderless UserForm
'
' WHY THIS EXISTS
'   Removing the native title bar removes the normal Windows drag surface. This
'   helper lets a runtime label, header banner, or other custom surface behave
'   like the missing title bar
'
' INPUTS
'   TargetForm
'     UserForm instance to move
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Resolves the UserForm window handle from its Caption, releases current mouse
'   capture, and sends a title-bar drag message to the UserForm window
'
' ERROR POLICY
'   Best-effort helper
'
'   Silently exits when WinAPI capability is unavailable, when TargetForm is
'   missing, or when the UserForm window handle cannot be resolved
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   M_Window_GetUserFormHwnd
'   ReleaseCapture
'   SendMessage
'   DP_WM_NCLBUTTONDOWN
'   DP_HTCAPTION
'
' NOTES
'   The UserForm Caption should remain stable and sufficiently unique while the
'   form is open. The title bar may be hidden, but the Caption property can still
'   be used internally to resolve the window handle
'
'   This routine uses capability only, not the optional styling setting. Mouse
'   movement is safe on Windows even when optional WinAPI styling is disabled
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then
    'Exit because this helper is Windows-only
        Exit Sub
#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    #If VBA7 Then
        Dim FormHandle As LongPtr       'Resolved UserForm window handle
    #Else
        Dim FormHandle As Long          'Resolved UserForm window handle
    #End If

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress drag helper errors
        On Error Resume Next

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Exit when WinAPI calls are not available
        If Not M_Platform_CanUseWinAPI Then Exit Sub
    'Exit when the target form reference is missing
        If TargetForm Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE USERFORM WINDOW
'------------------------------------------------------------------------------
    'Resolve the handle through the single module resolver
        FormHandle = M_Window_GetUserFormHwnd(TargetForm)
    'Exit when the UserForm window cannot be resolved
        If FormHandle = 0 Then GoTo CleanExit

'------------------------------------------------------------------------------
' START NATIVE DRAG
'------------------------------------------------------------------------------
    'Release current mouse capture
        ReleaseCapture
    'Ask Windows to move the form as if the title bar were being dragged
        SendMessage FormHandle, DP_WM_NCLBUTTONDOWN, DP_HTCAPTION, 0

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Clear any suppressed drag error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

#End If

End Sub


#If VBA7 Then
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As LongPtr
#Else
Public Function M_Window_GetUserFormHwnd(ByVal Frm As Object) As Long
#End If

'
'------------------------------------------------------------------------------
'                           WINDOW GET USERFORM HWND
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves the native window handle for a loaded UserForm
'
' WHY THIS EXISTS
'   Some DatePicker visual and positioning operations require the native Windows
'   window handle of the UserForm. This lookup must remain centralized so all
'   WinAPI-dependent behavior uses the same Mac-safe and fail-safe handle policy
'
' INPUTS
'   Frm
'     UserForm instance whose native window handle should be resolved
'
' RETURNS
'   Native UserForm window handle on Windows
'   Zero when running on Mac, when WinAPI behavior is disabled by setting, when
'   no form is supplied, when the caption is blank, when the handle cannot be
'   found, or when lookup fails
'
' BEHAVIOR
'   Returns zero on Mac. On Windows, validates the supplied form, reads its
'   caption, then attempts to find the UserForm window using the common
'   ThunderDFrame class and the alternate ThunderXFrame class
'
' ERROR POLICY
'   Safe default. Returns zero on lookup failure and writes diagnostics to the
'   Immediate Window
'
' DEPENDENCIES
'   FindWindow
'   ThunderDFrame UserForm window class
'   ThunderXFrame UserForm window class
'   M_Platform_ShouldUseWinAPI
'
' NOTES
'   This routine deliberately avoids raising outward because native handle lookup
'   is required only for optional WinAPI-dependent UI behavior
'
'   The lookup is caption-based because MSForms UserForms do not expose their
'   native window handle directly through the object model
'
'   If multiple loaded UserForms share the same caption, Windows may return the
'   first matching window. DatePicker captions should therefore remain unique
'   while WinAPI behavior is enabled
'
'   This is the single handle resolver for the module. Callers must not perform
'   their own FindWindow lookup, so the caption policy and the blank-caption
'   guard apply to every WinAPI-dependent routine
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then

'------------------------------------------------------------------------------
' MAC SAFE DEFAULT
'------------------------------------------------------------------------------
    'Return zero on Mac
        M_Window_GetUserFormHwnd = 0

#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_GetUserFormHwnd"

    Dim FormCaption             As String        'Current form caption
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Return the safe default unless lookup succeeds
        M_Window_GetUserFormHwnd = 0
    'Suppress handle lookup errors through the local fail-safe path
        On Error GoTo FailSafe

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Function

'------------------------------------------------------------------------------
' CHECK WINAPI CAPABILITY
'------------------------------------------------------------------------------
    'Exit when the current platform cannot use WinAPI calls
        If Not M_Platform_CanUseWinAPI Then Exit Function

'------------------------------------------------------------------------------
' READ FORM CAPTION
'------------------------------------------------------------------------------
    'Read the current form caption
        FormCaption = VBA.CStr(Frm.Caption)
    'Exit if the caption is blank
        If VBA.LenB(FormCaption) = 0 Then Exit Function

'------------------------------------------------------------------------------
' TRY PRIMARY USERFORM WINDOW CLASS
'------------------------------------------------------------------------------
    'Try the most common MSForms UserForm window class
        M_Window_GetUserFormHwnd = FindWindow("ThunderDFrame", FormCaption)

'------------------------------------------------------------------------------
' TRY ALTERNATE USERFORM WINDOW CLASS
'------------------------------------------------------------------------------
    'Try the alternate MSForms UserForm window class
        If M_Window_GetUserFormHwnd = 0 Then
            M_Window_GetUserFormHwnd = FindWindow("ThunderXFrame", FormCaption)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after lookup
        Exit Function

'------------------------------------------------------------------------------
' FAIL-SAFE
'------------------------------------------------------------------------------
FailSafe:
    'Capture the lookup error number
        ErrorNumber = Err.Number
    'Capture the lookup error description
        ErrorDescription = Err.Description
    'Return the safe default
        M_Window_GetUserFormHwnd = 0
    'Write diagnostics only when lookup failed with an error
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear the suppressed lookup error
        Err.Clear

#End If

End Function

Public Sub M_Window_MoveFormToMouse( _
    ByVal Frm As Object, _
    Optional ByVal OffsetXPx As Long = 0, _
    Optional ByVal OffsetYPx As Long = 12, _
    Optional ByVal CenterOnMouse As Boolean = False)

'
'------------------------------------------------------------------------------
'                           WINDOW MOVE FORM TO MOUSE
'------------------------------------------------------------------------------
' PURPOSE
'   Moves a DatePicker UserForm near the current mouse position on Windows
'
' WHY THIS EXISTS
'   The DatePicker may be shown as a compact modeless popup. Mouse-based
'   positioning keeps the form close to the user's interaction point while
'   clamping the window inside the active monitor work area
'
' INPUTS
'   Frm
'     UserForm instance to move
'
'   OffsetXPx
'     Horizontal offset in pixels from the mouse position
'
'   OffsetYPx
'     Vertical offset in pixels from the mouse position
'
'   CenterOnMouse
'     True to center the form on the mouse position before applying offsets
'     False to use the mouse position as the top-left anchor before offsets
'
' RETURNS
'   Nothing
'
' BEHAVIOR
'   Exits safely on Mac, when no form is supplied, when WinAPI is unavailable,
'   when the mouse position cannot be read, when the form handle cannot be
'   resolved, when the form rectangle cannot be read, or when the monitor work
'   area cannot be resolved
'
'   On Windows, calculates the desired form position from the mouse position,
'   applies the requested offsets, clamps the result to the nearest monitor work
'   area, and moves the form without resizing, changing z-order, or activating it
'
' ERROR POLICY
'   Best-effort positioning. Unexpected WinAPI, monitor, or form-positioning
'   errors are suppressed and written to the Immediate Window for diagnostics
'
' DEPENDENCIES
'   M_Platform_CanUseWinAPI
'   M_Window_GetUserFormHwnd
'   GetCursorPos
'   GetWindowRect
'   MonitorFromRect
'   GetMonitorInfo
'   SetWindowPos
'   POINTAPI
'   RECT
'   MONITORINFO
'   MONITOR_DEFAULTTONEAREST
'   SWP_NOSIZE
'   SWP_NOZORDER
'   SWP_NOACTIVATE
'   SWP_SHOWWINDOW
'
' NOTES
'   This routine intentionally does nothing on Mac
'
'   This routine uses M_Platform_CanUseWinAPI, not M_Platform_ShouldUseWinAPI,
'   because mouse positioning is a capability-only behavior. Disabling optional
'   WinAPI styling should disable title-bar removal, not mouse positioning.
'
'   Positioning is pixel-based because WinAPI cursor, window, and monitor APIs
'   operate in screen pixels
'
'   The nearest monitor work area is used so the picker remains visible in
'   multi-monitor setups
'
' UPDATED
'   2026-08-21
'------------------------------------------------------------------------------

#If Mac Then

'------------------------------------------------------------------------------
' MAC SAFE EXIT
'------------------------------------------------------------------------------
    'Do nothing on Mac
        Exit Sub

#Else

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Window_MoveFormToMouse"

    Dim CursorPoint             As POINTAPI      'Mouse cursor position
    Dim CursorRect              As RECT          'One-pixel cursor rectangle
    Dim FormRect                As RECT          'Current form window rectangle
    Dim MonitorInfoData         As MONITORINFO   'Nearest monitor information

    #If VBA7 Then
        Dim hWndForm            As LongPtr       'UserForm window handle
        Dim hMonitor            As LongPtr       'Nearest monitor handle
    #Else
        Dim hWndForm            As Long          'UserForm window handle
        Dim hMonitor            As Long          'Nearest monitor handle
    #End If

    Dim FormWidthPx             As Long          'Form width in pixels
    Dim FormHeightPx            As Long          'Form height in pixels
    Dim WorkWidthPx             As Long          'Monitor work-area width in pixels
    Dim WorkHeightPx            As Long          'Monitor work-area height in pixels
    Dim TargetX                 As Long          'Target X position in pixels
    Dim TargetY                 As Long          'Target Y position in pixels

    Dim MoveFlags               As Long          'SetWindowPos movement flags

    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

    Dim ApiResult               As Long          'Generic WinAPI Boolean-style result
    Dim LastApiError            As Long          'WinAPI last-error code

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Suppress positioning errors through the local cleanup path
        On Error GoTo CleanExit

'------------------------------------------------------------------------------
' VALIDATE INPUT
'------------------------------------------------------------------------------
    'Exit if no form was supplied
        If Frm Is Nothing Then Exit Sub

'------------------------------------------------------------------------------
' CHECK WINAPI CAPABILITY
'------------------------------------------------------------------------------
    'Exit when the current platform cannot use WinAPI calls
        If Not M_Platform_CanUseWinAPI Then Exit Sub

'------------------------------------------------------------------------------
' READ MOUSE POSITION
'------------------------------------------------------------------------------
    'Exit if the cursor position cannot be read
        If GetCursorPos(CursorPoint) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' RESOLVE FORM HANDLE
'------------------------------------------------------------------------------
    'Resolve the native UserForm window handle
        hWndForm = M_Window_GetUserFormHwnd(Frm)
    'Exit if the form window handle is unavailable
        If hWndForm = 0 Then Exit Sub

'------------------------------------------------------------------------------
' READ FORM RECTANGLE
'------------------------------------------------------------------------------
    'Exit if the form window rectangle cannot be read
        If GetWindowRect(hWndForm, FormRect) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE FORM SIZE
'------------------------------------------------------------------------------
    'Calculate the form width in pixels
        FormWidthPx = FormRect.Right - FormRect.Left
    'Calculate the form height in pixels
        FormHeightPx = FormRect.Bottom - FormRect.Top
    'Exit if resolved form size is invalid
        If FormWidthPx <= 0 Or FormHeightPx <= 0 Then Exit Sub

'------------------------------------------------------------------------------
' BUILD CURSOR RECTANGLE
'------------------------------------------------------------------------------
    'Set cursor rectangle left edge
        CursorRect.Left = CursorPoint.X
    'Set cursor rectangle top edge
        CursorRect.Top = CursorPoint.Y
    'Set cursor rectangle right edge
        CursorRect.Right = CursorPoint.X + 1
    'Set cursor rectangle bottom edge
        CursorRect.Bottom = CursorPoint.Y + 1

'------------------------------------------------------------------------------
' RESOLVE NEAREST MONITOR
'------------------------------------------------------------------------------
    'Resolve the nearest monitor handle
        hMonitor = MonitorFromRect(CursorRect, MONITOR_DEFAULTTONEAREST)
    'Exit if the monitor handle is unavailable
        If hMonitor = 0 Then Exit Sub

'------------------------------------------------------------------------------
' READ MONITOR WORK AREA
'------------------------------------------------------------------------------
    'Initialize the monitor-info structure size
        MonitorInfoData.cbSize = Len(MonitorInfoData)
    'Exit if monitor information cannot be read
        If GetMonitorInfo(hMonitor, MonitorInfoData) = 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE WORK AREA SIZE
'------------------------------------------------------------------------------
    'Calculate the monitor work-area width
        WorkWidthPx = MonitorInfoData.rcWork.Right - MonitorInfoData.rcWork.Left
    'Calculate the monitor work-area height
        WorkHeightPx = MonitorInfoData.rcWork.Bottom - MonitorInfoData.rcWork.Top
    'Exit if monitor work area is invalid
        If WorkWidthPx <= 0 Or WorkHeightPx <= 0 Then Exit Sub

'------------------------------------------------------------------------------
' CALCULATE TARGET POSITION
'------------------------------------------------------------------------------
    'Use the mouse position as the default top-left anchor
        TargetX = CursorPoint.X + OffsetXPx
    'Use the mouse position as the default top-left anchor
        TargetY = CursorPoint.Y + OffsetYPx
    'Center the form on the mouse when requested
        If CenterOnMouse Then
            TargetX = CursorPoint.X - (FormWidthPx \ 2) + OffsetXPx
            TargetY = CursorPoint.Y - (FormHeightPx \ 2) + OffsetYPx
        End If

'------------------------------------------------------------------------------
' CLAMP TARGET X POSITION
'------------------------------------------------------------------------------
    'Anchor X to the work-area left edge when the form is wider than the work area
        If FormWidthPx >= WorkWidthPx Then
            TargetX = MonitorInfoData.rcWork.Left
    'Clamp X to the work-area left edge
        ElseIf TargetX < MonitorInfoData.rcWork.Left Then
            TargetX = MonitorInfoData.rcWork.Left
    'Clamp X to the work-area right edge
        ElseIf TargetX + FormWidthPx > MonitorInfoData.rcWork.Right Then
            TargetX = MonitorInfoData.rcWork.Right - FormWidthPx
        End If

'------------------------------------------------------------------------------
' CLAMP TARGET Y POSITION
'------------------------------------------------------------------------------
    'Anchor Y to the work-area top edge when the form is taller than the work area
        If FormHeightPx >= WorkHeightPx Then
            TargetY = MonitorInfoData.rcWork.Top
    'Clamp Y to the work-area top edge
        ElseIf TargetY < MonitorInfoData.rcWork.Top Then
            TargetY = MonitorInfoData.rcWork.Top
    'Clamp Y to the work-area bottom edge
        ElseIf TargetY + FormHeightPx > MonitorInfoData.rcWork.Bottom Then
            TargetY = MonitorInfoData.rcWork.Bottom - FormHeightPx
        End If

'------------------------------------------------------------------------------
' MOVE FORM
'------------------------------------------------------------------------------
    'Build movement flags
        MoveFlags = SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
    'Clear the WinAPI last-error state before moving the form
        SetLastError 0
    'Move the form without resizing, changing z-order, or activating it
        ApiResult = SetWindowPos(hWndForm, 0, TargetX, TargetY, 0, 0, MoveFlags)
    'Diagnose SetWindowPos return-code failure
        If ApiResult = 0 Then
            LastApiError = Err.LastDllError
            Debug.Print PROC_NAME & _
                " | Step=Move form" & _
                " | Api=SetWindowPos" & _
                " | LastError=" & VBA.CStr(LastApiError)
        End If

'------------------------------------------------------------------------------
' CLEAN EXIT
'------------------------------------------------------------------------------
CleanExit:
    'Capture any suppressed error number
        ErrorNumber = Err.Number
    'Capture any suppressed error description
        ErrorDescription = Err.Description
    'Write diagnostics only when mouse positioning failed
        If ErrorNumber <> 0 Then
            Debug.Print PROC_NAME & _
                " | Error=" & VBA.CStr(ErrorNumber) & _
                " | " & ErrorDescription
        End If
    'Clear any suppressed positioning error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

#End If

End Sub
