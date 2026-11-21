Attribute VB_Name = "M_DP_Captions"

Option Explicit
Option Private Module

'
'------------------------------------------------------------------------------
' MODULE: M_DP_Captions
'------------------------------------------------------------------------------
' PURPOSE
'   Builds the month and date captions the picker displays, including the
'   fixed-English month names used when local names are off
'
' WHY THIS EXISTS
'   Caption text is pure formatting with no state of its own. Keeping it
'   apart from settings, write-back and lifecycle code gives the form and
'   the regression harness one source of captions (#56)
'
' INPUTS
'   None at module level
'
' RETURNS
'   Nothing at module level
'
' BEHAVIOR
'   Provides:
'     - M_Caption_GetMonth and M_Caption_GetDate, which honor the
'       UseLocalNames argument
'     - M_Caption_GetEnglishMonthShort and M_Caption_GetEnglishMonthFull,
'       which never depend on the Windows locale
'
' ERROR POLICY
'   Each procedure documents its own policy
'
' DEPENDENCIES
'   VBA date and string functions
'
' NOTES
'   Option Private Module keeps these helpers out of other VBA projects and
'   out of the macro dialog. The callers pass every input as an argument, so
'   the module reads no shared state
'
' UPDATED
'   2026-11-21
'------------------------------------------------------------------------------


'
'------------------------------------------------------------------------------
'
'                                   CAPTIONS
'
'------------------------------------------------------------------------------
'

Public Function M_Caption_GetMonth( _
    ByVal MonthNumber As Long, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET MONTH CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   logic keeps local-name mode and fixed-English mode consistent across the
'   form, picker panels, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
'   UseLocalNames
'     True to use the local VBA month name
'     False to use the fixed-English DatePicker caption
'
' RETURNS
'   Uppercase trimmed month caption
'
' BEHAVIOR
'   Validates the month number, resolves the caption using either local VBA
'   month names or the fixed-English helper, then returns a normalized uppercase
'   caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
'   Other unexpected failures are allowed to propagate naturally so the caller
'   receives the original VBA runtime error
'
' DEPENDENCIES
'   VBA.MonthName
'   M_Caption_GetEnglishMonthFull
'
' NOTES
'   Fixed-English mode is delegated to M_Caption_GetEnglishMonthFull so English
'   month naming remains controlled by the DatePicker companion-module API
'
'   Local-name mode uses VBA.MonthName and is therefore driven by the host
'   Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME         As String = "M_Caption_GetMonth"

    Dim MonthCaption        As String       'Resolved month caption

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RESOLVE MONTH CAPTION
'------------------------------------------------------------------------------
    'Resolve the month caption using local host names when requested
        If UseLocalNames Then
            MonthCaption = VBA.MonthName(MonthNumber, False)
    'Otherwise resolve the month caption using the fixed-English helper
        Else
            MonthCaption = M_Caption_GetEnglishMonthFull(MonthNumber)
        End If

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the normalized display caption
        M_Caption_GetMonth = VBA.UCase$(VBA.Trim$(MonthCaption))

End Function

Public Function M_Caption_GetDate( _
    ByVal DateValue As Date, _
    ByVal UseLocalNames As Boolean) As String

'
'------------------------------------------------------------------------------
'                           GET DATE CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the display caption for a DatePicker date value
'
' WHY THIS EXISTS
'   DatePicker date captions are used in several UI locations. Centralizing the
'   logic keeps local-name mode and fixed-English mode consistent across the
'   form, picker panels, and regression tests
'
' INPUTS
'   DateValue
'     Date value to convert into display text
'
'   UseLocalNames
'     True to use the local abbreviated month name
'     False to use the fixed-English DatePicker abbreviated month caption
'
' RETURNS
'   Date caption in dd-MMM-yyyy style
'
' BEHAVIOR
'   Builds the day, month, and year components from DateValue, resolves the month
'   component using either local VBA formatting or the fixed-English helper, and
'   returns the final normalized caption
'
' ERROR POLICY
'   Raises unexpected failures back to the caller with this procedure name as
'   the error source
'
' DEPENDENCIES
'   VBA.Format$
'   VBA.Month
'   M_Caption_GetEnglishMonthShort
'
' NOTES
'   Local-name mode uses VBA.Format$(DateValue, "mmm") and is therefore driven
'   by the host Office / Windows locale
'
'   Fixed-English mode is delegated to M_Caption_GetEnglishMonthShort so English
'   month naming remains controlled by the DatePicker companion-module API
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetDate"

    Dim DayText                 As String        'Day component
    Dim MonthText               As String        'Month component
    Dim YearText                As String        'Year component
    Dim ErrorNumber             As Long          'Captured error number
    Dim ErrorDescription        As String        'Captured error description

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Enable controlled error handling
        On Error GoTo ErrorHandler

'------------------------------------------------------------------------------
' BUILD DAY COMPONENT
'------------------------------------------------------------------------------
    'Build the two-digit day component
        DayText = VBA.Format$(DateValue, "dd")

'------------------------------------------------------------------------------
' BUILD MONTH COMPONENT
'------------------------------------------------------------------------------
    'Build the local abbreviated month component when requested
        If UseLocalNames Then
            MonthText = VBA.UCase$(VBA.Format$(DateValue, "mmm"))
    'Otherwise build the fixed-English abbreviated month component
        Else
            MonthText = M_Caption_GetEnglishMonthShort(VBA.Month(DateValue))
        End If

'------------------------------------------------------------------------------
' BUILD YEAR COMPONENT
'------------------------------------------------------------------------------
    'Build the four-digit year component
        YearText = VBA.Format$(DateValue, "yyyy")

'------------------------------------------------------------------------------
' RETURN VALUE
'------------------------------------------------------------------------------
    'Return the final date caption
        M_Caption_GetDate = DayText & "-" & MonthText & "-" & YearText

'------------------------------------------------------------------------------
' EXIT PROCEDURE
'------------------------------------------------------------------------------
    'Exit before the error handler
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
ErrorHandler:
    'Capture the error number
        ErrorNumber = Err.Number
    'Capture the error description
        ErrorDescription = Err.Description
    'Raise a descriptive error to the caller
        Err.Raise ErrorNumber, PROC_NAME, ErrorDescription

End Function

Public Function M_Caption_GetEnglishMonthShort(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISH MONTH SHORT CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed-English abbreviated month caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   fixed-English caption map keeps display behavior consistent across the form,
'   picker panels, caption helpers, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
' RETURNS
'   Fixed-English uppercase three-letter month caption
'
' BEHAVIOR
'   Validates the month number and returns the corresponding fixed-English short
'   month caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper deliberately does not use VBA.Format$ or VBA.MonthName because
'   fixed-English mode must be independent from the host Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetEnglishMonthShort"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RETURN FIXED-ENGLISH SHORT MONTH CAPTION
'------------------------------------------------------------------------------
    'Return the fixed-English short caption for the requested month
        Select Case MonthNumber
            Case 1
                M_Caption_GetEnglishMonthShort = "JAN"
            Case 2
                M_Caption_GetEnglishMonthShort = "FEB"
            Case 3
                M_Caption_GetEnglishMonthShort = "MAR"
            Case 4
                M_Caption_GetEnglishMonthShort = "APR"
            Case 5
                M_Caption_GetEnglishMonthShort = "MAY"
            Case 6
                M_Caption_GetEnglishMonthShort = "JUN"
            Case 7
                M_Caption_GetEnglishMonthShort = "JUL"
            Case 8
                M_Caption_GetEnglishMonthShort = "AUG"
            Case 9
                M_Caption_GetEnglishMonthShort = "SEP"
            Case 10
                M_Caption_GetEnglishMonthShort = "OCT"
            Case 11
                M_Caption_GetEnglishMonthShort = "NOV"
            Case 12
                M_Caption_GetEnglishMonthShort = "DEC"
        End Select

End Function

Public Function M_Caption_GetEnglishMonthFull(ByVal MonthNumber As Long) As String

'
'------------------------------------------------------------------------------
'                           GET ENGLISH MONTH FULL CAPTION
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the fixed-English full month caption for a month number
'
' WHY THIS EXISTS
'   DatePicker month captions are used in several UI locations. Centralizing the
'   fixed-English caption map keeps display behavior consistent across the form,
'   picker panels, caption helpers, and regression tests
'
' INPUTS
'   MonthNumber
'     Month number to convert
'     Must be between 1 and 12
'
' RETURNS
'   Fixed-English uppercase full month caption
'
' BEHAVIOR
'   Validates the month number and returns the corresponding fixed-English full
'   month caption
'
' ERROR POLICY
'   Raises vbObjectError + 513 when MonthNumber is outside the supported range
'
' DEPENDENCIES
'   None
'
' NOTES
'   This helper deliberately does not use VBA.Format$ or VBA.MonthName because
'   fixed-English mode must be independent from the host Office / Windows locale
'
' UPDATED
'   2026-05-06
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Const PROC_NAME             As String = "M_Caption_GetEnglishMonthFull"

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Reject month numbers outside the supported calendar range
        If MonthNumber < 1 Or MonthNumber > 12 Then
            Err.Raise vbObjectError + 513, PROC_NAME, "MonthNumber must be between 1 and 12."
        End If

'------------------------------------------------------------------------------
' RETURN FIXED-ENGLISH FULL MONTH CAPTION
'------------------------------------------------------------------------------
    'Return the fixed-English full caption for the requested month
        Select Case MonthNumber
            Case 1: M_Caption_GetEnglishMonthFull = "JANUARY"
            Case 2: M_Caption_GetEnglishMonthFull = "FEBRUARY"
            Case 3: M_Caption_GetEnglishMonthFull = "MARCH"
            Case 4: M_Caption_GetEnglishMonthFull = "APRIL"
            Case 5: M_Caption_GetEnglishMonthFull = "MAY"
            Case 6: M_Caption_GetEnglishMonthFull = "JUNE"
            Case 7: M_Caption_GetEnglishMonthFull = "JULY"
            Case 8: M_Caption_GetEnglishMonthFull = "AUGUST"
            Case 9: M_Caption_GetEnglishMonthFull = "SEPTEMBER"
            Case 10: M_Caption_GetEnglishMonthFull = "OCTOBER"
            Case 11: M_Caption_GetEnglishMonthFull = "NOVEMBER"
            Case 12: M_Caption_GetEnglishMonthFull = "DECEMBER"
        End Select

End Function
