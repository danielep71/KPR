'==============================================================================
' MODULE: KPR_DATES_DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   Finance-oriented, worksheet-facing date primitives, date arithmetic, weekday
'   locators, and pillar labeling utilities for the KPR toolkit.
'
'   Public functions are designed for direct use in Excel formulas
'   (Function Wizard, tooltips, scalar usage) with deterministic input handling
'   and native Excel error values on failure.
'
' WHY THIS EXISTS
'   Date logic for pricing, scheduling, accrual boundaries, reporting cutoffs,
'   and tenor labeling appears repeatedly across financial workbook models.
'   This module centralizes those recurring calendar primitives behind a single,
'   consistent contract so callers do not need to re-implement:
'     - date parsing / normalization policy
'     - supported-range gating
'     - month-end and leap-year edge handling
'     - EOM-aware month arithmetic
'
' SCOPE / PUBLIC SURFACE
'   - Day primitives
'       * KPR_Dates_DayOfWeek
'       * KPR_Dates_DaysInMonth
'       * KPR_Dates_DaysInYear
'
'   - Month primitives
'       * KPR_Dates_BeginOfMonth
'       * KPR_Dates_EndOfMonth
'       * KPR_Dates_IsMonthEnd
'       * KPR_Dates_IsQuarterEnd
'
'   - Year primitives
'       * KPR_Dates_IsLeapYear
'       * KPR_Dates_IsYearEnd
'
'   - Date arithmetic (EOM-aware)
'       * KPR_Dates_AddWeeks
'       * KPR_Dates_AddMonths       (Optional Opt_KeepEOM As Boolean)
'       * KPR_Dates_AddYears        (Optional Opt_KeepEOM As Boolean)
'
'   - Weekday locators
'       * KPR_Dates_NthWeekdayOfMonth
'       * KPR_Dates_LastWeekdayOfMonth
'
'   - Pillar / tenor formatting
'       * KPR_Dates_PillarFromDates
'       * KPR_Dates_DatesFromPillar
'
' SHAPE CONTRACT (THIS REVISION)
'   - All public functions are SCALAR ONLY.
'   - Parse_Date still unwraps a single-cell Range and a 1x1 wrapper array, so
'     VBA callers and legacy single-cell references keep working.
'   - Multi-cell / multi-element inputs are rejected as #VALUE!.
'   - Spill (2D) support is deliberately deferred; when it is added it wraps
'     these scalar cores rather than duplicating them.
'
' DESIGN / INPUT NORMALIZATION
'   - Every DateIn-style argument is a Variant funnelled through Parse_Date,
'     which owns the whole acceptance policy:
'       * deterministic object / array unwrapping
'       * date-only normalization (time component stripped)
'       * rejection of blanks / Excel errors / Boolean
'       * acceptance of date-like strings and numeric Excel serials
'       * a SINGLE supported-range gate applied to every accepted path
'
'   - The range gate is the module contract, not a per-function check:
'         KPR_MIN_DATE = 1900-03-01
'         KPR_MAX_DATE = 9999-12-31
'     Because the gate is applied after parsing rather than only to numeric
'     serials, a native VBA Date of 1850-01-01 is now rejected exactly like the
'     equivalent serial. Individual functions therefore carry no year gate.
'
'   - The lower bound is 1900-03-01 (serial 61) by design: it excludes the
'     Excel 1900-system region where worksheet serials and VBA dates disagree
'     because of the fictitious 29-Feb-1900.
'
' ERROR POLICY (USER FACING)
'   Public UDFs return either a valid scalar result or a native Excel error.
'   No public function returns a message string.
'
'       #VALUE!  the input cannot be interpreted
'                (unparseable date, non-scalar input, malformed pillar,
'                 month outside 1..12, weekday outside 1..7, occurrence
'                 outside 1..5, unexpected runtime error)
'
'       #NUM!    the input is well formed but the answer does not exist or
'                falls outside the supported date window
'                (no such weekday occurrence in the month, arithmetic or
'                 pillar shift beyond KPR_MIN_DATE / KPR_MAX_DATE)
'
'   Rationale: message strings are invisible to IFERROR / ISERROR, break any
'   downstream arithmetic that references the cell, and cannot be aggregated or
'   filtered. Native error values are the only return that composes.
'
' INTERNAL ERROR POLICY
'   - Private cores are Try-style (Boolean return, result ByRef) or plain
'     computation on already-validated inputs.
'   - Public wrappers own all worksheet-facing error behavior.
'
' DEPENDENCIES / INTEGRATION
'   - Private helpers in this module:
'       * Parse_Date
'       * Array_Rank_1Or2
'       * EndOfMonth_Core
'       * IsLeapYear_Core
'       * TryAddMonths_Core
'       * TryPillar_Parse
'       * Pillar_Format_Nearest
'
'   - No external module dependencies.
'
' NOTES / LIMITATIONS
'   - Date math follows Excel / VBA DateSerial, DateAdd, DateDiff, and Weekday
'     semantics.
'   - Gregorian leap-year logic is used throughout.
'   - AddMonths / AddYears EOM handling is controlled by Opt_KeepEOM:
'       FALSE => clip day-of-month to target month length
'       TRUE  => if input is EOM, return target EOM
'   - AddYears delegates to the single month shifter so year-roll edge cases
'     (29-Feb, short months, EOM) cannot drift between the two functions.
'   - Pillar formatting uses TRUE NEAREST rounding via Pillar_Format_Nearest.
'   - String date parsing follows host locale rules (CDate / IsDate).
'
' UPDATED
'   2026-08-29
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit  'Force explicit variable declarations

'------------------------------------------------------------------------------
' MODULE CONSTANTS
'------------------------------------------------------------------------------
    'Supported date window (single source of truth for every parse / result gate)
        Private Const KPR_MIN_DATE  As Date = #3/1/1900#     'Serial 61; excludes the Excel 1900 leap-bug region
        Private Const KPR_MAX_DATE  As Date = #12/31/9999#   'Serial 2958465; upper bound of the Excel date system

'
'------------------------------------------------------------------------------
'
'                           PUBLIC API - DAY PRIMITIVES
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_DayOfWeek( _
    ByVal DateIn As Variant, _
    Optional ByVal Opt_WeekBaseMonday As Boolean = True) _
    As Variant
'
'==============================================================================
' KPR_Dates_DayOfWeek
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the weekday index for DateIn, with configurable week base:
'     - Opt_WeekBaseMonday = TRUE  -> 1..7 where Mon=1 .. Sun=7
'     - Opt_WeekBaseMonday = FALSE -> 1..7 where Sun=1 .. Sat=7
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
'   Opt_WeekBaseMonday (optional)
'     Weekday() base selector:
'       TRUE  => vbMonday (Mon=1..Sun=7)
'       FALSE => vbSunday (Sun=1..Sat=7)
'
' RETURNS
'   Variant
'     Long (1..7) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'
' NOTES
'   - The week base is exposed as Boolean so the worksheet API stays explicit:
'         TRUE  => Monday-based numbering (ISO habit)
'         FALSE => Sunday-based numbering (US habit)
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date          'Parsed date-only value (per Parse_Date policy)
    Dim WkBase          As VbDayOfWeek   'Weekday() base selector (vbMonday or vbSunday)
    Dim FailErr         As Variant       'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If Opt_WeekBaseMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return weekday index 1..7 under the selected base
        KPR_Dates_DayOfWeek = Weekday(ParsedDate, WkBase)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DayOfWeek = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_DaysInMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_DaysInMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of calendar days in the month containing DateIn.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Long (28..31) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - EndOfMonth_Core
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Month length is the day number of the containing month-end
        KPR_Dates_DaysInMonth = CLng(Day(EndOfMonth_Core(ParsedDate)))
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DaysInMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_DaysInYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_DaysInYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the number of calendar days in the year containing DateIn:
'     - 365 for common years
'     - 366 for leap years (Gregorian rule)
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Long (365 / 366) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - IsLeapYear_Core
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return 366 for leap years, 365 otherwise
        If IsLeapYear_Core(Year(ParsedDate)) Then
            KPR_Dates_DaysInYear = 366&
        Else
            KPR_Dates_DaysInYear = 365&
        End If

        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DaysInYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                          PUBLIC API - MONTH PRIMITIVES
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_BeginOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_BeginOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the first calendar day of the month containing DateIn.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'
' NOTES
'   - Only the year / month components of the parsed input survive; the result
'     is always day 1.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the first day of the containing month
        KPR_Dates_BeginOfMonth = DateSerial(Year(ParsedDate), Month(ParsedDate), 1)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_BeginOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_EndOfMonth( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_EndOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last calendar day of the month containing DateIn.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - EndOfMonth_Core
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the last day of the containing month
        KPR_Dates_EndOfMonth = EndOfMonth_Core(ParsedDate)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_EndOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_IsMonthEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_IsMonthEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if DateIn is the last calendar day of its month.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - EndOfMonth_Core
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Compare the parsed date to its own month-end
        KPR_Dates_IsMonthEnd = (ParsedDate = EndOfMonth_Core(ParsedDate))
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsMonthEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_IsQuarterEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_IsQuarterEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if DateIn is a quarter-end date:
'     - month is one of {Mar, Jun, Sep, Dec}
'     - AND DateIn is the last calendar day of that month
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - EndOfMonth_Core
'
' NOTES
'   - Quarter months are detected as Month Mod 3 = 0, which is exactly the set
'     {3, 6, 9, 12}. This is a calendar quarter, not a fiscal quarter; a fiscal
'     variant belongs in a separate function with an explicit anchor argument.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Quarter-end = month-end AND quarter month
        KPR_Dates_IsQuarterEnd = (ParsedDate = EndOfMonth_Core(ParsedDate)) _
                                 And ((Month(ParsedDate) Mod 3) = 0)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsQuarterEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                           PUBLIC API - YEAR PRIMITIVES
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_IsYearEnd( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_IsYearEnd
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if DateIn is a year-end date (31 December).
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'
' NOTES
'   - December always has 31 days, so no month-end computation is required:
'     the test is simply month 12 and day 31.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Year-end is 31 December
        KPR_Dates_IsYearEnd = ((Month(ParsedDate) = 12) And (Day(ParsedDate) = 31))
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsYearEnd = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_IsLeapYear( _
    ByVal DateIn As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_IsLeapYear
'------------------------------------------------------------------------------
' PURPOSE
'   Returns TRUE if the year containing DateIn is a leap year under the
'   Gregorian rule.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - IsLeapYear_Core
'
' NOTES
'   - The previous per-function year gate is gone: Parse_Date already refuses
'     anything before KPR_MIN_DATE, so no year reaching this point can be out
'     of the supported range.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Apply the Gregorian leap-year rule to the containing year
        KPR_Dates_IsLeapYear = IsLeapYear_Core(Year(ParsedDate))
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_IsLeapYear = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                               DATE ARITHMETICS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_AddWeeks( _
    ByVal DateIn As Variant, _
    ByVal nWeeks As Long) _
    As Variant
'
'==============================================================================
' KPR_Dates_AddWeeks
'------------------------------------------------------------------------------
' PURPOSE
'   Adds nWeeks to DateIn using calendar-day arithmetic:
'       Result = DateIn + (7 * nWeeks) days
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
'   nWeeks
'     Number of weeks to add. Negative values are allowed.
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'   #NUM!    result falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
' DEPENDENCIES
'   - Parse_Date
'
' NOTES
'   - The shift is computed in Double and range-gated BEFORE coercion back to
'     Date, so a large nWeeks returns #NUM! rather than raising an overflow.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim ResultD         As Double    'Shifted date serial before range gating
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Compute the shifted serial in Double to avoid overflow on coercion
        ResultD = CDbl(ParsedDate) + (7# * CDbl(nWeeks))

    'Gate the result to the supported date window
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the shifted date
        KPR_Dates_AddWeeks = CDate(ResultD)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddWeeks = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_AddMonths( _
    ByVal DateIn As Variant, _
    ByVal nMonths As Long, _
    Optional ByVal Opt_KeepEOM As Boolean = False) _
    As Variant
'
'==============================================================================
' KPR_Dates_AddMonths
'------------------------------------------------------------------------------
' PURPOSE
'   Adds nMonths to DateIn with deterministic end-of-month (EOM) handling:
'     - Opt_KeepEOM = FALSE : clip day-of-month to the target month length
'     - Opt_KeepEOM = TRUE  : if DateIn is EOM, return the target EOM
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
'   nMonths
'     Number of calendar months to add. Negative values are allowed.
'
'   Opt_KeepEOM (optional)
'     TRUE  => EOM in, EOM out
'     FALSE => clip day-of-month when the target month is shorter
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'   #NUM!    result falls outside KPR_MIN_DATE .. KPR_MAX_DATE
'
' DEPENDENCIES
'   - Parse_Date
'   - TryAddMonths_Core
'
' NOTES
'   - Worked examples, clip mode (Opt_KeepEOM = FALSE):
'         31-Jan-2026 + 1M  => 28-Feb-2026
'         30-Apr-2026 + 1M  => 30-May-2026
'   - Same inputs, preserve mode (Opt_KeepEOM = TRUE):
'         31-Jan-2026 + 1M  => 28-Feb-2026
'         30-Apr-2026 + 1M  => 31-May-2026
'     The two modes differ only when the input is itself month-end.
'   - This is the single month shifter in the module; AddYears and the pillar
'     resolver both route through it.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim ResultDate      As Date      'Shifted date returned by the month core
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths_Core(ParsedDate, nMonths, Opt_KeepEOM, ResultDate) Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the shifted date
        KPR_Dates_AddMonths = ResultDate
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddMonths = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_AddYears( _
    ByVal DateIn As Variant, _
    ByVal nYears As Long, _
    Optional ByVal Opt_KeepEOM As Boolean = False) _
    As Variant
'
'==============================================================================
' KPR_Dates_AddYears
'------------------------------------------------------------------------------
' PURPOSE
'   Adds nYears to DateIn by converting years to months (12 * nYears) and
'   applying the same EOM semantics as KPR_Dates_AddMonths.
'
' INPUTS
'   DateIn
'     Date candidate handled by the module Parse_Date policy.
'
'   nYears
'     Number of calendar years to add. Negative values are allowed.
'
'   Opt_KeepEOM (optional)
'     TRUE  => EOM in, EOM out
'     FALSE => clip day-of-month when the target month is shorter
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'   #NUM!    month delta or result falls outside the supported range
'
' DEPENDENCIES
'   - Parse_Date
'   - TryAddMonths_Core
'
' NOTES
'   - Delegating to the month core keeps every year-roll edge case identical to
'     AddMonths. The one that matters:
'         29-Feb-2024 + 1Y => 28-Feb-2025 in both EOM modes
'   - The years -> months conversion is done in Double so a large nYears cannot
'     overflow the Long month delta before it is gated.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per Parse_Date policy)
    Dim MonthsD         As Double    'Year -> month conversion before Long gating
    Dim ResultDate      As Date      'Shifted date returned by the month core
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Parse to date-only under module policy (shape + value + range)
        If Not Parse_Date(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' CONVERT YEARS TO MONTHS
'------------------------------------------------------------------------------
    'Convert in Double to avoid intermediate overflow
        MonthsD = 12# * CDbl(nYears)

    'Gate to Long range before coercion
        If (MonthsD < -2147483648#) Or (MonthsD > 2147483647#) Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths_Core(ParsedDate, CLng(MonthsD), Opt_KeepEOM, ResultDate) Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the shifted date
        KPR_Dates_AddYears = ResultDate
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_AddYears = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                               WEEKDAY LOCATORS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_NthWeekdayOfMonth( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WdIndex As Long, _
    ByVal n As Long, _
    Optional ByVal Opt_WeekBaseMonday As Boolean = True) _
    As Variant
'
'==============================================================================
' KPR_Dates_NthWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the N-th occurrence of a requested weekday within a given
'   Year / Month.
'
' INPUTS
'   YearIn
'     Calendar year within the module supported range.
'
'   MonthIn
'     Calendar month number in 1..12.
'
'   WdIndex
'     Weekday index in 1..7 under the selected Weekday() base:
'       Opt_WeekBaseMonday = TRUE  => 1 = Mon .. 7 = Sun
'       Opt_WeekBaseMonday = FALSE => 1 = Sun .. 7 = Sat
'
'   n
'     Occurrence number in 1..5.
'
'   Opt_WeekBaseMonday (optional)
'     TRUE  => vbMonday base (default)
'     FALSE => vbSunday base
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  YearIn, MonthIn, WdIndex or n outside their accepted domains,
'            or unexpected runtime error
'   #NUM!    arguments are valid but the occurrence does not exist in that
'            month (typically a requested 5th weekday)
'
' DEPENDENCIES
'   - VBA.DateSerial
'   - VBA.Weekday
'
' NOTES
'   - The #VALUE! / #NUM! split is deliberate: "weekday 9" is a caller mistake,
'     "no 5th Friday in February" is a legitimate question with no answer.
'   - WdIndex must be passed consistently with Opt_WeekBaseMonday.
'   - Capping n at 5 is sufficient: no calendar month holds six occurrences of
'     the same weekday.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim D0              As Date         'First calendar day of the target month
    Dim Off             As Long         'Days from D0 to the first occurrence of WdIndex (0..6)
    Dim dOut            As Date         'Computed date for the requested occurrence
    Dim WkBase          As VbDayOfWeek  'Weekday() base selector (vbMonday or vbSunday)
    Dim FailErr         As Variant      'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a year inside the module supported window
        If (YearIn < Year(KPR_MIN_DATE)) Or (YearIn > Year(KPR_MAX_DATE)) Then GoTo Fail

    'Require a valid calendar month
        If (MonthIn < 1) Or (MonthIn > 12) Then GoTo Fail

    'Require a weekday index in 1..7 under the selected base
        If (WdIndex < 1) Or (WdIndex > 7) Then GoTo Fail

    'Require an occurrence number in 1..5
        If (n < 1) Or (n > 5) Then GoTo Fail

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If Opt_WeekBaseMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' COMPUTE REQUESTED OCCURRENCE
'------------------------------------------------------------------------------
    'Anchor at the first day of the requested month
        D0 = DateSerial(YearIn, MonthIn, 1)

    'Forward offset (0..6) from the anchor to the first requested weekday
        Off = (WdIndex - Weekday(D0, WkBase) + 7) Mod 7

    'Step forward to the requested occurrence
        dOut = D0 + Off + (7& * (n - 1))

    'Reject an occurrence that has spilled into the following month
        If Month(dOut) <> MonthIn Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the requested occurrence
        KPR_Dates_NthWeekdayOfMonth = dOut
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_NthWeekdayOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_LastWeekdayOfMonth( _
    ByVal YearIn As Long, _
    ByVal MonthIn As Long, _
    ByVal WdIndex As Long, _
    Optional ByVal Opt_WeekBaseMonday As Boolean = True) _
    As Variant
'
'==============================================================================
' KPR_Dates_LastWeekdayOfMonth
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last occurrence of a requested weekday within a given
'   Year / Month.
'
' INPUTS
'   YearIn
'     Calendar year within the module supported range.
'
'   MonthIn
'     Calendar month number in 1..12.
'
'   WdIndex
'     Weekday index in 1..7 under the selected Weekday() base:
'       Opt_WeekBaseMonday = TRUE  => 1 = Mon .. 7 = Sun
'       Opt_WeekBaseMonday = FALSE => 1 = Sun .. 7 = Sat
'
'   Opt_WeekBaseMonday (optional)
'     TRUE  => vbMonday base (default)
'     FALSE => vbSunday base
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  YearIn, MonthIn or WdIndex outside their accepted domains, or
'            unexpected runtime error
'
' DEPENDENCIES
'   - VBA.DateSerial
'   - VBA.Weekday
'
' NOTES
'   - A last occurrence always exists, so this function has no #NUM! path.
'   - WdIndex must be passed consistently with Opt_WeekBaseMonday.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim EOM             As Date         'End-of-month date for the target Year / Month
    Dim Diff            As Long         'Backward offset in days from EOM to the requested weekday (0..6)
    Dim WkBase          As VbDayOfWeek  'Weekday() base selector (vbMonday or vbSunday)
    Dim FailErr         As Variant      'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' VALIDATE INPUTS
'------------------------------------------------------------------------------
    'Require a year inside the module supported window
        If (YearIn < Year(KPR_MIN_DATE)) Or (YearIn > Year(KPR_MAX_DATE)) Then GoTo Fail

    'Require a valid calendar month
        If (MonthIn < 1) Or (MonthIn > 12) Then GoTo Fail

    'Require a weekday index in 1..7 under the selected base
        If (WdIndex < 1) Or (WdIndex > 7) Then GoTo Fail

'------------------------------------------------------------------------------
' RESOLVE WEEKDAY BASE
'------------------------------------------------------------------------------
    'Resolve Weekday() base once
        If Opt_WeekBaseMonday Then
            WkBase = vbMonday
        Else
            WkBase = vbSunday
        End If

'------------------------------------------------------------------------------
' COMPUTE LAST OCCURRENCE
'------------------------------------------------------------------------------
    'Anchor at end-of-month
        EOM = DateSerial(YearIn, MonthIn + 1, 0)

    'Backward offset (0..6) from the anchor to the requested weekday
        Diff = (Weekday(EOM, WkBase) - WdIndex + 7) Mod 7

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the last requested weekday in the target month
        KPR_Dates_LastWeekdayOfMonth = EOM - Diff
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_LastWeekdayOfMonth = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                                    PILLARS
'
'------------------------------------------------------------------------------
'

Public Function KPR_Dates_PillarFromDates( _
    ByVal StartDate As Variant, _
    ByVal EndDate As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_PillarFromDates
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the yield-curve bucket label ("pillar") spanning StartDate to
'   EndDate, using TRUE NEAREST rounding.
'
' INPUTS
'   StartDate
'     Date candidate handled by the module Parse_Date policy.
'
'   EndDate
'     Date candidate handled by the module Parse_Date policy.
'
' RETURNS
'   Variant
'     String pillar token on success (e.g. "3W", "5M", "2Y4M", "-1W"),
'     else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  either date not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - Parse_Date
'   - Pillar_Format_Nearest
'
' NOTES
'   - The label is a String by design; it is a category, not a quantity, and is
'     meant to be grouped or matched rather than computed on.
'   - EndDate before StartDate yields a "-" prefixed token.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedStart     As Date      'Parsed start date (per Parse_Date policy)
    Dim ParsedEnd       As Date      'Parsed end date   (per Parse_Date policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUTS
'------------------------------------------------------------------------------
    'Parse the start date under module policy
        If Not Parse_Date(StartDate, ParsedStart) Then GoTo Fail

    'Parse the end date under module policy
        If Not Parse_Date(EndDate, ParsedEnd) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Format the nearest-rounded pillar token
        KPR_Dates_PillarFromDates = Pillar_Format_Nearest(ParsedStart, ParsedEnd)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_PillarFromDates = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

Public Function KPR_Dates_DatesFromPillar( _
    ByVal StartDate As Variant, _
    ByVal Pillar As Variant) _
    As Variant
'
'==============================================================================
' KPR_Dates_DatesFromPillar
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the calendar date obtained by applying a pillar token to StartDate.
'
' INPUTS
'   StartDate
'     Date candidate handled by the module Parse_Date policy.
'
'   Pillar
'     Tenor token. Accepted grammar:
'       - optional leading sign: "+" or "-", applied to the WHOLE token
'       - one or more [integer][unit] components, no spaces inside
'       - units: Y, M, W, D (case insensitive)
'       - whole-token aliases: "ON" / "O/N" => 1D, "TN" / "T/N" => 2D
'
'     Examples: "1W", "3M", "25Y", "2Y4M", "1Y6M2W", "-6D", "ON", "T/N"
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  StartDate not parseable, Pillar not scalar text, or Pillar does
'            not match the accepted grammar, or unexpected runtime error
'   #NUM!    the pillar is well formed but shifts the date outside
'            KPR_MIN_DATE .. KPR_MAX_DATE
'
' DEPENDENCIES
'   - Parse_Date
'   - TryPillar_Parse
'   - TryAddMonths_Core
'
' NOTES
'   - Y and M components are aggregated into one calendar-month shift, applied
'     first with CLIP semantics (never preserve-EOM). W and D are then applied
'     as an exact calendar-day delta. Order matters: "1M1D" from 31-Jan-2026 is
'     28-Feb + 1D = 01-Mar-2026.
'   - Duplicate units accumulate: "1Y2Y3M" is 3Y3M.
'   - Parsing is strict. Nothing is silently stripped, so a malformed token can
'     never be reinterpreted as a different valid pillar.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedStart     As Date      'Parsed StartDate (per Parse_Date policy)
    Dim WorkDate        As Date      'Intermediate date after the month shift
    Dim ResultD         As Double    'Final serial after the day shift, before range gating

    Dim TotalMonths     As Double    'Signed month delta parsed from the pillar
    Dim TotalDays       As Double    'Signed day delta parsed from the pillar

    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = CVErr(xlErrValue)

'------------------------------------------------------------------------------
' PARSE INPUTS
'------------------------------------------------------------------------------
    'Parse the start date under module policy
        If Not Parse_Date(StartDate, ParsedStart) Then GoTo Fail

    'Reject a non-scalar pillar payload before touching its text
        If IsObject(Pillar) Or IsArray(Pillar) Then GoTo Fail

    'Parse the pillar token into signed month / day deltas
        If Not TryPillar_Parse(Pillar, TotalMonths, TotalDays) Then GoTo Fail

'------------------------------------------------------------------------------
' APPLY MONTH SHIFT
'------------------------------------------------------------------------------
    'Start from the parsed date
        WorkDate = ParsedStart

    'Apply the aggregated Y / M components with clip semantics
        If TotalMonths <> 0# Then

            'Gate the month delta to Long range before coercion
                If (TotalMonths < -2147483648#) Or (TotalMonths > 2147483647#) Then
                    FailErr = CVErr(xlErrNum)
                    GoTo Fail
                End If

            'Delegate to the single month shifter (never preserve-EOM here)
                If Not TryAddMonths_Core(WorkDate, CLng(TotalMonths), False, WorkDate) Then
                    FailErr = CVErr(xlErrNum)
                    GoTo Fail
                End If

        End If

'------------------------------------------------------------------------------
' APPLY DAY SHIFT
'------------------------------------------------------------------------------
    'Apply the aggregated W / D components as an exact calendar-day delta
        ResultD = CDbl(WorkDate) + TotalDays

    'Gate the result to the supported date window
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            FailErr = CVErr(xlErrNum)
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Return the resolved date
        KPR_Dates_DatesFromPillar = CDate(ResultD)
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return the worksheet-facing error value
        KPR_Dates_DatesFromPillar = FailErr
        Exit Function

'------------------------------------------------------------------------------
' ERROR HANDLER
'------------------------------------------------------------------------------
Err_Handler:
    'Unexpected runtime error => #VALUE!
        FailErr = CVErr(xlErrValue)
        Resume Fail

End Function

'
'******************************************************************************
'
'                                 DATE HELPERS
'
'******************************************************************************
'

Private Function Parse_Date( _
    ByVal DateIn As Variant, _
    ByRef ParsedDate As Date) _
    As Boolean
'
'==============================================================================
'                                  Parse_Date
'------------------------------------------------------------------------------
' PURPOSE
'   Parse a single Excel date candidate (Variant) into a normalized VBA Date
'   (date-only), returning a Boolean success flag.
'
' WHY THIS EXISTS
'   Worksheet-facing APIs receive dates in many shapes and runtime types. This
'   routine is the single point where the acceptance policy lives, so that:
'     - callers do not duplicate fragile type / shape logic
'     - no unhandled runtime error can reach the worksheet
'     - the supported-range gate is applied identically to every input form
'
' SIGNATURE
'   Parse_Date(DateIn, ParsedDate) -> Boolean
'
' INPUTS
'   DateIn
'     Variant date candidate. Accepted:
'       - vbDate values
'       - numeric Excel serials
'       - date-like strings (host locale rules)
'       - numeric-looking strings (treated as serials)
'       - single-cell Range (unwrapped via Value2)
'       - 1x1 Variant array (unwrapped to its scalar)
'
' OUTPUTS
'   ParsedDate (ByRef)
'     Assigned ONLY on success, normalized to date-only.
'
' RETURNS
'   Boolean
'     TRUE  => parsing succeeded; ParsedDate assigned
'     FALSE => parsing failed; ParsedDate untouched
'
' BEHAVIOR
'   - Unwraps Range / array wrappers first, then classifies the scalar.
'   - Rejects non-Range objects, multi-cell ranges, arrays other than 1x1.
'   - Rejects Excel error values, Empty / Null / blanks, and Boolean.
'   - Every accepted path lands on a single candidate value, which is then
'     gated once against KPR_MIN_DATE / KPR_MAX_DATE.
'
' ERROR POLICY
'   - Does not propagate VBA runtime errors to callers.
'   - Signals failure only through the Boolean return.
'
' DEPENDENCIES
'   - Array_Rank_1Or2
'
' NOTES
'   - Boolean is rejected deliberately: TRUE coerces to serial -1 in VBA, which
'     is neither a date the user meant nor a value inside the supported range.
'   - The single terminal gate is the reason this revision drops the per-
'     function year checks. Previously a native Date bypassed the numeric
'     serial range check entirely, so 01-Jan-1850 was accepted as a Date but
'     rejected as a serial. Both are now rejected.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim V           As Variant      'Working value after unwrapping Range / array wrappers
    Dim VT          As VbVarType    'Cached VarType taken only after unwrapping
    Dim Rng         As Range        'Bound Range when DateIn is a Range object

    Dim Candidate   As Date         'Parsed candidate before the terminal range gate
    Dim X           As Double       'Numeric working value for serial inputs
    Dim S           As String       'Trimmed string token

    Dim ArrRank     As Long         'Array shape discriminator (1 or 2)
    Dim LB1         As Long         'Lower bound, dimension 1
    Dim UB1         As Long         'Upper bound, dimension 1
    Dim LB2         As Long         'Lower bound, dimension 2
    Dim UB2         As Long         'Upper bound, dimension 2

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime coercion errors and convert them to a FALSE return
        On Error GoTo Fail

    'Default return is failure unless we explicitly succeed
        Parse_Date = False

    'Work on a local copy to avoid repeated Variant indirection
        V = DateIn

'------------------------------------------------------------------------------
' UNWRAP OBJECTS
'------------------------------------------------------------------------------
    'If an object arrived, support only a single-cell Range
        If IsObject(V) Then

            'Reject non-Range objects
                If Not TypeOf V Is Range Then GoTo Fail

            'Bind strongly typed
                Set Rng = V

            'Require exactly one cell
                If Rng.CountLarge <> 1 Then GoTo Fail

            'Unwrap the cell value
                V = Rng.Value2

        End If

'------------------------------------------------------------------------------
' UNWRAP ARRAYS
'------------------------------------------------------------------------------
    'If an array arrived, accept only a 1x1 wrapper (scalar contract)
        If IsArray(V) Then

            'Classify the array shape once
                ArrRank = Array_Rank_1Or2(V)

            'Two-dimensional wrapper
                If ArrRank = 2 Then

                    'Capture bounds
                        LB1 = LBound(V, 1): UB1 = UBound(V, 1)
                        LB2 = LBound(V, 2): UB2 = UBound(V, 2)

                    'Reject anything larger than 1x1
                        If (UB1 <> LB1) Or (UB2 <> LB2) Then GoTo Fail

                    'Unwrap the scalar payload
                        V = V(LB1, LB2)

            'One-dimensional wrapper
                Else

                    'Capture bounds
                        LB1 = LBound(V): UB1 = UBound(V)

                    'Reject vectors with more than one element
                        If UB1 <> LB1 Then GoTo Fail

                    'Unwrap the scalar payload
                        V = V(LB1)

                End If

        End If

'------------------------------------------------------------------------------
' CLASSIFY SCALAR
'------------------------------------------------------------------------------
    'Take VarType only after unwrapping
        VT = VarType(V)

    Select Case VT

        Case vbError, vbEmpty, vbNull, vbBoolean
            'Excel errors, blanks and Boolean are never dates
                GoTo Fail

        Case vbDate
            'Native Date => strip any time component
                Candidate = DateValue(CDate(V))

        Case vbString
            'Trim once for deterministic checks
                S = Trim$(CStr(V))

            'Blank string is not a date
                If Len(S) = 0 Then GoTo Fail

            'Numeric-looking string => treat as an Excel serial
                If IsNumeric(S) Then
                    X = CDbl(S)
                    Candidate = CDate(Fix(X))

            'Otherwise require a locale-recognized date string
                Else
                    If Not IsDate(S) Then GoTo Fail
                    Candidate = DateValue(CDate(S))
                End If

        Case Else
            'Remaining scalars are acceptable only if numeric (Excel serials)
                If Not IsNumeric(V) Then GoTo Fail

            'Coerce once and strip the fractional day
                X = CDbl(V)
                Candidate = CDate(Fix(X))

    End Select

'------------------------------------------------------------------------------
' TERMINAL RANGE GATE
'------------------------------------------------------------------------------
    'One gate for every accepted path
        If (Candidate < KPR_MIN_DATE) Or (Candidate > KPR_MAX_DATE) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate

    'Contract: TRUE only when ParsedDate was assigned
        Parse_Date = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving ParsedDate untouched
        Parse_Date = False

End Function

Private Function Array_Rank_1Or2( _
    ByVal ArrayIn As Variant) _
    As Long
'
'==============================================================================
'                               Array_Rank_1Or2
'------------------------------------------------------------------------------
' PURPOSE
'   Project-specific array shape classifier.
'
'   Returns only:
'     - 1 => dimension 2 is NOT addressable
'     - 2 => dimension 2 IS addressable
'
' WHY THIS EXISTS
'   Callers do not need the exact mathematical rank of an array, only whether
'   it can be indexed as a matrix. Keeping the probe here means no caller has
'   to hold its own On Error Resume Next block.
'
' SIGNATURE
'   Array_Rank_1Or2(ArrayIn) -> Long
'
' INPUTS
'   ArrayIn
'     Variant that may or may not hold an array payload.
'
' RETURNS
'   Long
'     1 => non-array, true 1D array, or uninitialized dynamic array
'     2 => 2D or higher
'
' ERROR POLICY
'   - Does NOT raise.
'   - Restores normal error handling before exit.
'
' NOTES
'   - Callers treating 1 as addressable 1D must still guard against
'     uninitialized dynamic arrays before calling LBound / UBound.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ProbeDim2   As Long      'Probe target; only success / failure matters

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default is 1 unless the dimension-2 probe succeeds
        Array_Rank_1Or2 = 1

'------------------------------------------------------------------------------
' VALIDATE
'------------------------------------------------------------------------------
    'Non-array input is never dimension-2 addressable
        If Not IsArray(ArrayIn) Then Exit Function

'------------------------------------------------------------------------------
' DIMENSION-2 PROBE
'------------------------------------------------------------------------------
    'Dimension 2 may be unavailable, so probe rather than test
        On Error Resume Next
        Err.Clear
        ProbeDim2 = LBound(ArrayIn, 2)

    'A clean probe means 2D or higher
        If Err.Number = 0 Then Array_Rank_1Or2 = 2

    'Restore normal error handling
        Err.Clear
        On Error GoTo 0

End Function

Private Function EndOfMonth_Core( _
    ByVal DateIn As Date) _
    As Date
'
'==============================================================================
'                               EndOfMonth_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Returns the last calendar day of the month containing DateIn.
'
' INPUTS
'   DateIn
'     Already-parsed, date-only VBA Date.
'
' RETURNS
'   Date
'
' ERROR POLICY
'   None. The caller owns validation; inputs reaching here are already parsed.
'
' NOTES
'   - Day 0 of the following month is the last day of the current one, and
'     DateSerial rolls month 13 into January of the next year on its own.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Day 0 of the following month
        EndOfMonth_Core = DateSerial(Year(DateIn), Month(DateIn) + 1, 0)

End Function

Private Function IsLeapYear_Core( _
    ByVal YearIn As Long) _
    As Boolean
'
'==============================================================================
'                               IsLeapYear_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Applies the Gregorian leap-year rule to a calendar year.
'
' INPUTS
'   YearIn
'     Calendar year, already inside the module supported range.
'
' RETURNS
'   Boolean
'
' ERROR POLICY
'   None. The caller owns validation.
'
' NOTES
'   - Rule: divisible by 4 => leap, unless divisible by 100, unless also
'     divisible by 400. So 1900 is not a leap year and 2000 is.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Gregorian rule
        IsLeapYear_Core = ((YearIn Mod 4 = 0) And ((YearIn Mod 100 <> 0) Or (YearIn Mod 400 = 0)))

End Function

Private Function TryAddMonths_Core( _
    ByVal DateIn As Date, _
    ByVal nMonths As Long, _
    ByVal KeepEOM As Boolean, _
    ByRef DateOut As Date) _
    As Boolean
'
'==============================================================================
'                               TryAddMonths_Core
'------------------------------------------------------------------------------
' PURPOSE
'   Single month shifter for the module. Adds nMonths to an already-parsed
'   date with explicit end-of-month semantics and full range safety.
'
' WHY THIS EXISTS
'   AddMonths, AddYears and the pillar resolver all need the same shift with
'   the same edge-case behavior. Duplicating it three times is how the three
'   drift apart.
'
' SIGNATURE
'   TryAddMonths_Core(DateIn, nMonths, KeepEOM, DateOut) -> Boolean
'
' INPUTS
'   DateIn
'     Already-parsed, date-only VBA Date.
'
'   nMonths
'     Signed month delta.
'
'   KeepEOM
'     TRUE  => if DateIn is month-end, return the target month-end
'     FALSE => clip day-of-month to the target month length
'
' OUTPUTS
'   DateOut (ByRef)
'     Assigned ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => shift succeeded and the result is inside the supported window
'     FALSE => the target month or the result falls outside it
'
' BEHAVIOR
'   - Converts year / month into an absolute month index, shifts it, and gates
'     the index before it is ever handed to DateSerial. This is why an absurd
'     nMonths returns FALSE instead of raising an overflow.
'   - Resolves the day-of-month by the KeepEOM rule, then re-gates the result.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'
' NOTES
'   - DateOut is safe to pass as the same variable as an input date at the call
'     site, because it is written only once, at the end.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim Y               As Long      'Year component of DateIn
    Dim M               As Long      'Month component of DateIn
    Dim DayIn           As Long      'Day-of-month of DateIn
    Dim IsEOM           As Boolean   'TRUE when DateIn is its own month-end

    Dim MonthIndex      As Double    'Absolute 0-based month index after the shift
    Dim TargetY         As Long      'Resolved target year
    Dim TargetM         As Long      'Resolved target month
    Dim TargetDim       As Long      'Days in the target month
    Dim dd              As Long      'Resolved day-of-month in the target month

    Dim Candidate       As Date      'Shifted date before the terminal range gate

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail

    'Default return is failure unless we explicitly succeed
        TryAddMonths_Core = False

'------------------------------------------------------------------------------
' CACHE INPUT COMPONENTS
'------------------------------------------------------------------------------
    'Split the input date once
        Y = Year(DateIn)
        M = Month(DateIn)
        DayIn = Day(DateIn)

    'Detect the EOM anchor before shifting
        IsEOM = (DayIn = Day(DateSerial(Y, M + 1, 0)))

'------------------------------------------------------------------------------
' RESOLVE TARGET MONTH
'------------------------------------------------------------------------------
    'Absolute 0-based month index, shifted, computed in Double
        MonthIndex = (CDbl(Y) * 12# + CDbl(M) - 1#) + CDbl(nMonths)

    'Gate the index before any coercion so DateSerial never sees a wild value
        If (MonthIndex < (CDbl(Year(KPR_MIN_DATE)) * 12#)) Or _
           (MonthIndex > (CDbl(Year(KPR_MAX_DATE)) * 12# + 11#)) Then GoTo Fail

    'Split the gated index back into year and month
        TargetY = CLng(Int(MonthIndex / 12#))
        TargetM = CLng(MonthIndex - (CDbl(TargetY) * 12#)) + 1

'------------------------------------------------------------------------------
' RESOLVE DAY OF MONTH
'------------------------------------------------------------------------------
    'Target month length
        TargetDim = Day(DateSerial(TargetY, TargetM + 1, 0))

    'Preserve EOM when asked and the input was EOM; otherwise clip
        If KeepEOM And IsEOM Then
            dd = TargetDim
        Else
            dd = DayIn
            If dd > TargetDim Then dd = TargetDim
        End If

'------------------------------------------------------------------------------
' TERMINAL RANGE GATE
'------------------------------------------------------------------------------
    'Build the candidate and gate it against the supported window
        Candidate = DateSerial(TargetY, TargetM, dd)
        If (Candidate < KPR_MIN_DATE) Or (Candidate > KPR_MAX_DATE) Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        DateOut = Candidate

    'Contract: TRUE only when DateOut was assigned
        TryAddMonths_Core = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving DateOut untouched
        TryAddMonths_Core = False

End Function

'
'******************************************************************************
'
'                                PILLAR HELPERS
'
'******************************************************************************
'

Private Function TryPillar_Parse( _
    ByVal PillarIn As Variant, _
    ByRef TotalMonths As Double, _
    ByRef TotalDays As Double) _
    As Boolean
'
'==============================================================================
'                               TryPillar_Parse
'------------------------------------------------------------------------------
' PURPOSE
'   Parses a pillar token into a signed month delta and a signed day delta.
'
' SIGNATURE
'   TryPillar_Parse(PillarIn, TotalMonths, TotalDays) -> Boolean
'
' INPUTS
'   PillarIn
'     Scalar Variant expected to hold pillar text.
'
' OUTPUTS
'   TotalMonths (ByRef)
'     Signed aggregate of the Y and M components: sign * (12Y + M).
'
'   TotalDays (ByRef)
'     Signed aggregate of the W and D components: sign * (7W + D).
'
'   Both are assigned ONLY on success.
'
' RETURNS
'   Boolean
'     TRUE  => the token matched the accepted grammar
'     FALSE => blank, malformed, unknown unit, or non-text payload
'
' BEHAVIOR
'   - Trims and upper-cases, consumes an optional leading sign, then checks the
'     whole-token aliases ON / O/N / TN / T/N.
'   - Otherwise scans repeated [digits][unit] pairs to the end of the string.
'   - Any character that is not a digit in a numeric position or a known unit
'     in a unit position fails the parse.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'
' NOTES
'   - No characters are stripped from the body, so "1 M" and "1/M" fail rather
'     than being silently reinterpreted as "1M". The aliases are the only
'     tokens containing "/", and they are matched whole.
'   - Fractional tenors are not part of the grammar: "1.5M" fails. Half months
'     are not a market convention, and admitting them would force a rounding
'     policy into a parser.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim S               As String    'Upper-cased, trimmed pillar text
    Dim SBody           As String    'Pillar body after the optional leading sign
    Dim Ch              As String    'Current character while scanning

    Dim p               As Long      'Current scan position in SBody
    Dim TokenStart      As Long      'Start position of the current numeric token
    Dim TokenCount      As Long      'Number of components parsed so far

    Dim SignMul         As Double    'Global sign multiplier (+1 or -1)
    Dim QtyD            As Double    'Current component quantity

    Dim YearsD          As Double    'Accumulated Y components
    Dim MonthsD         As Double    'Accumulated M components
    Dim WeeksD          As Double    'Accumulated W components
    Dim DaysD           As Double    'Accumulated D components

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Trap runtime errors and convert them to a FALSE return
        On Error GoTo Fail

    'Default return is failure unless we explicitly succeed
        TryPillar_Parse = False

    'Default sign is positive
        SignMul = 1#

'------------------------------------------------------------------------------
' NORMALIZE PILLAR TEXT
'------------------------------------------------------------------------------
    'Require a text payload
        If VarType(PillarIn) <> vbString Then GoTo Fail

    'Trim and upper-case once
        S = UCase$(Trim$(CStr(PillarIn)))

    'Reject a blank token
        If Len(S) = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' CONSUME OPTIONAL LEADING SIGN
'------------------------------------------------------------------------------
    'A sign applies to the whole token
        Select Case Left$(S, 1)
            Case "+"
                SBody = Mid$(S, 2)
            Case "-"
                SignMul = -1#
                SBody = Mid$(S, 2)
            Case Else
                SBody = S
        End Select

    'Reject a sign with no body
        If Len(SBody) = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' WHOLE-TOKEN ALIASES
'------------------------------------------------------------------------------
    'Overnight and tom-next are matched whole, never as generic tokens
        Select Case SBody
            Case "ON", "O/N"
                DaysD = 1#
                TokenCount = 1
            Case "TN", "T/N"
                DaysD = 2#
                TokenCount = 1
        End Select

'------------------------------------------------------------------------------
' GENERIC INTEGER + UNIT SCANNER
'------------------------------------------------------------------------------
    'Only run the scanner when no alias matched
        If TokenCount = 0 Then

            'Start at the first character of the body
                p = 1

            'Consume [digits][unit] pairs until the body is exhausted
                Do While p <= Len(SBody)

                    'Mark the start of the numeric token
                        TokenStart = p

                    'Advance through contiguous digits
                        Do While p <= Len(SBody)
                            Ch = Mid$(SBody, p, 1)
                            If (Ch < "0") Or (Ch > "9") Then Exit Do
                            p = p + 1
                        Loop

                    'Reject a component with no digits
                        If TokenStart = p Then GoTo Fail

                    'Coerce the numeric token once
                        QtyD = CDbl(Mid$(SBody, TokenStart, p - TokenStart))

                    'Reject a quantity with no trailing unit
                        If p > Len(SBody) Then GoTo Fail

                    'Accumulate by unit
                        Select Case Mid$(SBody, p, 1)
                            Case "Y"
                                YearsD = YearsD + QtyD
                            Case "M"
                                MonthsD = MonthsD + QtyD
                            Case "W"
                                WeeksD = WeeksD + QtyD
                            Case "D"
                                DaysD = DaysD + QtyD
                            Case Else
                                GoTo Fail
                        End Select

                    'Count the component and step past the unit
                        TokenCount = TokenCount + 1
                        p = p + 1

                Loop

        End If

    'Defensive: require at least one parsed component
        If TokenCount = 0 Then GoTo Fail

'------------------------------------------------------------------------------
' ASSIGN RESULTS
'------------------------------------------------------------------------------
    'Aggregate Y / M into a signed month delta
        TotalMonths = SignMul * ((12# * YearsD) + MonthsD)

    'Aggregate W / D into a signed day delta
        TotalDays = SignMul * ((7# * WeeksD) + DaysD)

    'Contract: TRUE only when both outputs were assigned
        TryPillar_Parse = True
        Exit Function

'------------------------------------------------------------------------------
' FAIL
'------------------------------------------------------------------------------
Fail:
    'Return FALSE per contract, leaving the outputs untouched
        TryPillar_Parse = False

End Function

Private Function Pillar_Format_Nearest( _
    ByVal DtStart As Date, _
    ByVal DtEnd As Date) _
    As String
'
'==============================================================================
'                            Pillar_Format_Nearest
'------------------------------------------------------------------------------
' PURPOSE
'   Formats the interval between two dates as a market-style tenor label using
'   TRUE NEAREST rounding with calendar-true month handling.
'
' INPUTS
'   DtStart
'     Already-parsed start date.
'
'   DtEnd
'     Already-parsed end date.
'
' RETURNS
'   String
'     Labels such as "0D", "6D", "3W", "5M", "2Y4M", prefixed with "-" when
'     DtEnd precedes DtStart.
'
' BEHAVIOR
'   - Zero days           => "0D"
'   - Fewer than 7 days   => exact "nD"
'   - Otherwise:
'       * build the nearest whole-week anchor from the start date
'       * build the nearest calendar-month anchor from the start date
'       * measure both against the actual end date and emit the closer one
'       * ties resolve to the month-style label
'
' ROUNDING POLICY
'   - Weeks: nearest whole week on calendar days, half up.
'   - Months: compare the floor and ceiling month anchors, take the closer,
'     ties round up. Anchors are built with DateAdd so month lengths are real.
'   - Family choice: whichever anchor sits closer to the true end date.
'
' ERROR POLICY
'   None. The caller owns parsing and worksheet messaging.
'
' DEPENDENCIES
'   - VBA.DateDiff
'   - VBA.DateAdd
'
' NOTES
'   - Deliberately coarse: no mixed tokens like "1M2W". A pillar is a curve
'     bucket label, not a precise interval.
'   - Day tokens are reserved for intervals under a week so longer intervals
'     never collapse into raw day counts.
'
' UPDATED
'   2026-08-29
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim SignPrefix      As String   '"-" when DtEnd precedes DtStart
    Dim D0              As Date     'Normalized earlier date
    Dim D1              As Date     'Normalized later date

    Dim TotalDays       As Long     'Calendar days between D0 and D1

    Dim nWeeks          As Long     'Nearest whole week count (half up)
    Dim WeekDate        As Date     'Anchor at D0 + 7 * nWeeks
    Dim DistWeek        As Long     'Day distance from D1 to the week anchor

    Dim FloorMonths     As Long     'Largest month count not overshooting D1
    Dim CeilMonths      As Long     'Next month count after FloorMonths
    Dim NearestMonths   As Long     'Chosen month count
    Dim DistMonth       As Long     'Day distance from D1 to the chosen month anchor
    Dim DistFloor       As Long     'Day distance to the floor month anchor
    Dim DistCeil        As Long     'Day distance to the ceiling month anchor

    Dim nYears          As Long     'Year component of NearestMonths
    Dim nMonths         As Long     'Residual month component of NearestMonths
    Dim OutTok          As String   'Unsigned output token

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Normalize the interval order and keep the direction as a prefix
        If DtEnd >= DtStart Then
            SignPrefix = vbNullString
            D0 = DtStart
            D1 = DtEnd
        Else
            SignPrefix = "-"
            D0 = DtEnd
            D1 = DtStart
        End If

    'Measure the interval once
        TotalDays = DateDiff("d", D0, D1)

'------------------------------------------------------------------------------
' SHORT INTERVALS
'------------------------------------------------------------------------------
    'Same date is explicit rather than empty
        If TotalDays = 0 Then
            Pillar_Format_Nearest = SignPrefix & "0D"
            Exit Function
        End If

    'Sub-week intervals stay exact in days
        If TotalDays < 7 Then
            Pillar_Format_Nearest = SignPrefix & CStr(TotalDays) & "D"
            Exit Function
        End If

'------------------------------------------------------------------------------
' NEAREST WEEK ANCHOR
'------------------------------------------------------------------------------
    'Nearest whole week, half up
        nWeeks = (TotalDays + 3) \ 7

    'Defensive clamp; TotalDays >= 7 here so this should already hold
        If nWeeks < 1 Then nWeeks = 1

    'Build the anchor and measure it against the true end date
        WeekDate = DateAdd("d", 7& * nWeeks, D0)
        DistWeek = Abs(DateDiff("d", WeekDate, D1))

'------------------------------------------------------------------------------
' NEAREST MONTH ANCHOR
'------------------------------------------------------------------------------
    'Crossed month boundaries, then step back if adding them overshoots
        FloorMonths = DateDiff("m", D0, D1)
        If FloorMonths > 0 Then
            If DateAdd("m", FloorMonths, D0) > D1 Then FloorMonths = FloorMonths - 1
        End If

    'Defensive clamp
        If FloorMonths < 0 Then FloorMonths = 0

    'Below one month the only meaningful month tenor is 1M
        If FloorMonths = 0 Then
            NearestMonths = 1
            DistMonth = Abs(DateDiff("d", DateAdd("m", 1, D0), D1))
        Else
            'Measure both bracketing anchors
                CeilMonths = FloorMonths + 1
                DistFloor = Abs(DateDiff("d", DateAdd("m", FloorMonths, D0), D1))
                DistCeil = Abs(DateDiff("d", DateAdd("m", CeilMonths, D0), D1))

            'Take the closer anchor; ties round up
                If DistCeil <= DistFloor Then
                    NearestMonths = CeilMonths
                    DistMonth = DistCeil
                Else
                    NearestMonths = FloorMonths
                    DistMonth = DistFloor
                End If
        End If

'------------------------------------------------------------------------------
' CHOOSE TENOR FAMILY
'------------------------------------------------------------------------------
    'A strictly closer week anchor wins outright
        If DistWeek < DistMonth Then
            Pillar_Format_Nearest = SignPrefix & CStr(nWeeks) & "W"
            Exit Function
        End If

    'Otherwise fall through to month-style output; ties prefer months
        nYears = NearestMonths \ 12
        nMonths = NearestMonths Mod 12

'------------------------------------------------------------------------------
' BUILD MONTH-STYLE TOKEN
'------------------------------------------------------------------------------
    'Emit years when present
        If nYears > 0 Then OutTok = OutTok & CStr(nYears) & "Y"

    'Emit residual months when present
        If nMonths > 0 Then OutTok = OutTok & CStr(nMonths) & "M"

    'Defensive fallback; unreachable for a positive month tenor
        If Len(OutTok) = 0 Then OutTok = "0D"

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Apply the direction prefix and return
        Pillar_Format_Nearest = SignPrefix & OutTok

End Function
