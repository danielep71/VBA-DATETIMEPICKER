Attribute VB_Name = "M_DP_DatePolicy"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_DatePolicy
'------------------------------------------------------------------------------
' PURPOSE
'   Decides whether a calendar date can be selected and whether the
'   optional holiday callback classifies it as a holiday
'
' WHY THIS EXISTS
'   The form must not own selection or holiday policy. Rendering, hover,
'   keyboard selection and click handling all ask this one module (#56)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - M_DatePolicy_CanSelectDate, which applies the supported year
'       range and the outside-month selection setting
'     - M_HolidayPolicy_IsHolidayDate, which runs the configured holiday
'       callback and accepts only a Boolean result
'
' ERROR POLICY
'   Each procedure documents its own policy. The holiday callback fails
'   open: a failing callback means not a holiday
'
' DEPENDENCIES
'   Application.Run
'   M_DP_Core (DP_MIN_YEAR, DP_MAX_YEAR,
'     gDP_AllowOutsideMonthSelection, gDP_HolidayCallbackName)
'
' NOTES
'   Option Private Module keeps these helpers out of other VBA projects and
'   out of the macro dialog. The policy reads the settings state that
'   M_DP_Core owns and never changes it
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                 DATE POLICY
'
'------------------------------------------------------------------------------
'

Public Function M_DatePolicy_CanSelectDate( _
    ByVal CandidateDate As Date, _
    ByVal DisplayYear As Long, _
    ByVal DisplayMonth As Long) As Boolean

'
'------------------------------------------------------------------------------
'                           CAN SELECT DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate calendar date can be selected in the DatePicker
'
' WHY THIS EXISTS
'   The UserForm should not decide selection policy directly
'
'   Date selection rules belong in the companion module so calendar rendering,
'   hover behavior, keyboard selection, and click handling all use the same
'   policy
'
' INPUTS
'   CandidateDate
'     Date being evaluated for selection
'
'   DisplayYear
'     Year currently displayed by the DatePicker grid
'
'   DisplayMonth
'     Month currently displayed by the DatePicker grid
'
' RETURNS
'   True when the candidate date can be selected
'
'   False when the display period is invalid or when outside-month selection is
'   disabled and the candidate date does not belong to the displayed month
'
' BEHAVIOR
'   Validates the displayed year and month, normalizes the candidate date to a
'   date-only value, allows all valid candidate dates when outside-month
'   selection is enabled, and otherwise allows only dates belonging to the
'   displayed month and year
'
' ERROR POLICY
'   Fail-closed policy routine
'
'   Returns False for invalid display state or unexpected runtime errors
'
'   Does not raise outward because this routine is used by high-frequency UI
'   paths such as grid rendering, hover handling, and keyboard navigation
'
' DEPENDENCIES
'   gDP_AllowOutsideMonthSelection
'
' NOTES
'   This routine intentionally does not call M_Settings_EnsureLoaded
'
'   Settings are loaded by public entry points and UserForm initialization before
'   this routine is used in normal runtime paths
'
'   CandidateDate is normalized with DateValue so date-time values are evaluated
'   by their date component only
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CandidateDateOnly       As Date             'Candidate date without time
    Dim CandidateYear           As Long             'Candidate date year
    Dim CandidateMonth          As Long             'Candidate date month

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_DatePolicy_CanSelectDate = False
    'Enable fail-closed error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' VALIDATE DISPLAY PERIOD
'------------------------------------------------------------------------------
    'Reject invalid display years
        If DisplayYear < DP_MIN_YEAR Or DisplayYear > DP_MAX_YEAR Then
            Exit Function
        End If
    'Reject invalid display months
        If DisplayMonth < 1 Or DisplayMonth > 12 Then Exit Function


'------------------------------------------------------------------------------
' NORMALIZE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Normalize the candidate date to its date-only component
        CandidateDateOnly = VBA.DateValue(CandidateDate)
    'Resolve the candidate year once
        CandidateYear = VBA.Year(CandidateDateOnly)
    'Resolve the candidate month once
        CandidateMonth = VBA.Month(CandidateDateOnly)

'------------------------------------------------------------------------------
' VALIDATE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Reject candidate dates outside the supported DatePicker year range
        If CandidateYear < DP_MIN_YEAR Or CandidateYear > DP_MAX_YEAR Then
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY OUTSIDE-MONTH POLICY
'------------------------------------------------------------------------------
    'Allow valid candidate dates when outside-month selection is enabled
        If gDP_AllowOutsideMonthSelection Then
            M_DatePolicy_CanSelectDate = True
            Exit Function
        End If

'------------------------------------------------------------------------------
' APPLY DISPLAYED-MONTH POLICY
'------------------------------------------------------------------------------
    'Allow only dates belonging to the displayed month and year
        M_DatePolicy_CanSelectDate = _
            (CandidateYear = DisplayYear And CandidateMonth = DisplayMonth)

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after policy evaluation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Return safe default on unexpected policy errors
        M_DatePolicy_CanSelectDate = False

End Function

Public Function M_HolidayPolicy_IsHolidayDate(ByVal CandidateDate As Date) As Boolean

'
'------------------------------------------------------------------------------
'                           IS HOLIDAY DATE
'------------------------------------------------------------------------------
' PURPOSE
'   Returns whether a candidate date is classified as a holiday by the optional
'   DatePicker holiday callback
'
' WHY THIS EXISTS
'   Holiday logic is workbook-specific
'
'   The DatePicker companion module should provide one controlled policy point
'   that can call a user-defined callback without hard-coding holiday calendars
'   into the UserForm rendering layer
'
' INPUTS
'   CandidateDate
'     Date to evaluate
'
' RETURNS
'   True when the configured callback explicitly returns True
'
'   False when no callback is configured, when the callback fails, when the
'   callback does not return a Boolean, or when the candidate date is outside
'   the supported DatePicker range
'
' BEHAVIOR
'   Trims the configured callback name, exits safely when no callback exists,
'   normalizes CandidateDate to a date-only value, validates the supported year
'   range, invokes the callback through Application.Run, and accepts only a
'   Boolean callback result
'
' ERROR POLICY
'   Fail-open from a date-selection perspective
'   Callback failures are interpreted as not holiday
'
' DEPENDENCIES
'   Application.Run
'   gDP_HolidayCallbackName
'
' NOTES
'   The holiday callback should accept one Date argument and return Boolean
'
'   The callback receives the date-only component of CandidateDate
'
'   Non-Boolean callback results are ignored deliberately to keep the callback
'   contract strict and predictable
'
' UPDATED
'   2026-05-03
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim CallbackName            As String           'Configured holiday callback name
    Dim CandidateDateOnly       As Date             'Candidate date without time
    Dim CandidateYear           As Long             'Candidate date year
    Dim CallbackResult          As Variant          'Callback return value

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Set safe default result
        M_HolidayPolicy_IsHolidayDate = False
    'Enable fail-open error handling
        On Error GoTo SafeExit

'------------------------------------------------------------------------------
' RESOLVE CALLBACK NAME
'------------------------------------------------------------------------------
    'Normalize the configured callback name
        CallbackName = VBA.Trim$(gDP_HolidayCallbackName)
    'Exit when no callback is configured
        If VBA.Len(CallbackName) = 0 Then Exit Function

'------------------------------------------------------------------------------
' NORMALIZE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Normalize the candidate date to its date-only component
        CandidateDateOnly = VBA.DateValue(CandidateDate)
    'Resolve the candidate year once
        CandidateYear = VBA.Year(CandidateDateOnly)

'------------------------------------------------------------------------------
' VALIDATE CANDIDATE DATE
'------------------------------------------------------------------------------
    'Reject dates before the supported DatePicker year range
        If CandidateYear < DP_MIN_YEAR Then Exit Function
    'Reject dates after the supported DatePicker year range
        If CandidateYear > DP_MAX_YEAR Then Exit Function

'------------------------------------------------------------------------------
' RUN HOLIDAY CALLBACK
'------------------------------------------------------------------------------
    'Run the configured holiday callback with the normalized candidate date
        CallbackResult = Application.Run(CallbackName, CandidateDateOnly)

'------------------------------------------------------------------------------
' INTERPRET CALLBACK RESULT
'------------------------------------------------------------------------------
    'Ignore Excel error callback results
        If VBA.IsError(CallbackResult) Then Exit Function
    'Ignore Null callback results
        If VBA.IsNull(CallbackResult) Then Exit Function
    'Return the callback result only when it is explicitly Boolean
        If VBA.VarType(CallbackResult) = vbBoolean Then
            M_HolidayPolicy_IsHolidayDate = VBA.CBool(CallbackResult)
        End If

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit after holiday-policy evaluation
        Exit Function

'------------------------------------------------------------------------------
' SAFE EXIT
'------------------------------------------------------------------------------
SafeExit:
    'Suppress callback or policy failures
        On Error Resume Next
    'Return safe default
        M_HolidayPolicy_IsHolidayDate = False
    'Clear any pending error
        Err.Clear
    'Restore normal error handling
        On Error GoTo 0

End Function
