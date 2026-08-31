Attribute VB_Name = "KPR_DATES_DAYS"
'==============================================================================
' MODULE: KPR_DATES_DAYS
'------------------------------------------------------------------------------
' PURPOSE
'   The worksheet-facing date facade for the KPR toolkit: date primitives, date
'   arithmetic, weekday locators and pillar labeling.
'
'   Public functions are designed for direct use in Excel formulas (Function
'   Wizard, tooltips, scalar usage) with deterministic input handling and native
'   Excel error values on failure.
'
' WHY THE NAME
'   Facade components are named in upper case to mark them as worksheet-facing
'   rather than internal, and this one is suffixed _DAYS because
'   KPR_DATES_CALENDARS and KPR_DATES_BUSINESSDAYS follow in later milestones.
'   Internal modules keep mixed-case KPR_Core_* names. The convention is
'   documented rather than enforced by a static casing rule.
'
' SCOPE / PUBLIC SURFACE (THIS REVISION)
'   Sixteen migrated functions. The frozen contract specifies twenty-two;
'   issue #15 adds the six missing names and applies the final signatures.
'
'   - Day primitives
'       * KPR_Dates_DayOfWeek           (Optional Opt_WeekBaseMonday)
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
'       * KPR_Dates_AddMonths           (Optional Opt_KeepEOM)
'       * KPR_Dates_AddYears            (Optional Opt_KeepEOM)
'
'   - Weekday locators
'       * KPR_Dates_NthWeekdayOfMonth   (Optional Opt_WeekBaseMonday)
'       * KPR_Dates_LastWeekdayOfMonth  (Optional Opt_WeekBaseMonday)
'
'   - Pillar / tenor formatting
'       * KPR_Dates_PillarFromDates
'       * KPR_Dates_DatesFromPillar
'
'   KPR_Dates_DatesFromPillar returns exactly one date; the plural name is a
'   baseline defect. The contract names it KPR_Dates_DateFromPillar, and the
'   rename lands with issue #15 so that this migration stays a pure move.
'
' ARCHITECTURE
'   This module is the only place a supported KPR_Dates_* function is declared.
'   It owns dispatch and worksheet-facing error mapping, and delegates every
'   computation:
'
'       KPR_Core_Array   shape: single-cell Range and 1x1 unwrapping
'       KPR_Core_Parse   value: scalar -> date-only VBA Date
'       KPR_Core_Dates   calendar: window, arithmetic, boundaries, pillars
'       KPR_Core_Err     construction of native Excel error values
'
'   The four core modules declare Option Private Module and never depend on this
'   facade, on registration, on UI, on tests or on demo code. This module
'   declares no Option Private Module, because its members must reach the
'   worksheet.
'
' SHAPE CONTRACT (THIS REVISION)
'   - All public functions are SCALAR ONLY.
'   - A single-cell Range and a 1x1 wrapper are accepted and unwrapped, so VBA
'     callers and legacy single-cell references keep working.
'   - Multi-cell / multi-element inputs are rejected as #VALUE!.
'   - Array traversal is deliberately deferred to issue #16. When it arrives it
'     wraps these scalar cores rather than duplicating them.
'
' DESIGN / INPUT NORMALIZATION
'   - Every DateIn-style argument is a Variant funnelled through TryResolveDate,
'     which composes the three boundaries in a fixed order:
'         unwrap shape -> parse value -> apply the supported window
'   - The window gate is applied exactly once on every accepted path, as in the
'     baseline. What changed is ownership, not the invariant: the window is a
'     calendar fact declared by KPR_Core_Dates, so KPR_Core_Parse does not need
'     to depend on it.
'
' ERROR POLICY (USER FACING)
'   Public UDFs return either a valid scalar result or a native Excel error.
'   No public function returns a message string.
'
'       #VALUE!  the input cannot be interpreted
'                (unparseable date, non-scalar input, out-of-window date,
'                 malformed pillar, month outside 1..12, weekday outside 1..7,
'                 occurrence outside 1..5, unexpected runtime error)
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
' KNOWN DIVERGENCES FROM THE FROZEN CONTRACT
'   This revision migrates behaviour unchanged, so the following remain
'   contract-invalid until their owning issues land. They are recorded here
'   rather than left for a reader to discover:
'
'     - An out-of-window date returns #VALUE!. The contract classifies this as
'       DATE_WINDOW and requires #NUM!.                             (#12, #13)
'     - An incoming Excel error is rejected as #VALUE! rather than propagated
'       verbatim.                                                   (#12)
'     - Text dates are parsed under host locale rules, and numeric-looking text
'       is reinterpreted as a serial.                               (#12)
'     - Integer arguments are declared As Long and are silently coerced.  (#12)
'     - Optional controls are declared As Boolean and are silently coerced.
'                                                                   (#13, #16)
'     - Duplicate pillar units accumulate rather than being rejected.    (#15)
'     - Pillar rounding is NEAREST only; Opt_Rounding is absent.         (#14)
'     - No host date-system detection exists, so a 1904 workbook is answered
'       with silently shifted values.                               (#17)
'     - Multi-cell inputs are rejected rather than traversed.            (#16)
'
' INTERNAL ERROR POLICY
'   - Core routines are Try-style (Boolean return, result ByRef) or plain
'     computation on already-validated inputs.
'   - This facade owns all worksheet-facing error behavior.
'   - Defensive catch-all handlers are containment only. They must be
'     unreachable in contract-conforming execution; an activation is a defect,
'     never an expected outcome.
'
' DEPENDENCIES / INTEGRATION
'   - KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates, KPR_Core_Err
'   - Private helper in this module: TryResolveDate
'   - No dependency on registration, UI, tests or demo code.
'
' UPDATED
'   2026-08-31
'
' AUTHOR
'   Daniele Penza
'==============================================================================

'------------------------------------------------------------------------------
' MODULE SETTINGS
'------------------------------------------------------------------------------
    Option Explicit  'Force explicit variable declarations

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
'     Date candidate handled by the facade TryResolveDate policy.
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
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'
' NOTES
'   - TRUE selects Monday-based numbering (ISO habit); FALSE selects
'     Sunday-based numbering (US habit).
'   - The Boolean declaration is NOT a strictness guarantee. Excel coerces 0, 1,
'     "TRUE" and even a date into a Boolean before this function is entered, so
'     an unintended value is accepted silently rather than rejected. The frozen
'     contract makes optional controls Variant and rejects everything that is
'     not a native Boolean; that change is owned by issues #13 and #16.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date          'Parsed date-only value (per TryResolveDate policy)
    Dim WkBase          As VbDayOfWeek   'Weekday() base selector (vbMonday or vbSunday)
    Dim FailErr         As Variant       'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Long (28..31) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.EndOfMonth_Core
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Long (365 / 366) on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.IsLeapYear_Core
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                        PUBLIC API - MONTH PRIMITIVES                         
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'
' NOTES
'   - Only the year / month components of the parsed input survive; the result
'     is always day 1.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Date on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.EndOfMonth_Core
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.EndOfMonth_Core
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.EndOfMonth_Core
'
' NOTES
'   - Quarter months are detected as Month Mod 3 = 0, which is exactly the set
'     {3, 6, 9, 12}. This is a calendar quarter, not a fiscal quarter; a fiscal
'     variant belongs in a separate function with an explicit anchor argument.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                         PUBLIC API - YEAR PRIMITIVES                         
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'
' NOTES
'   - December always has 31 days, so no month-end computation is required:
'     the test is simply month 12 and day 31.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
'
' RETURNS
'   Variant
'     Boolean on success, else a native Excel error value
'
' ERROR POLICY (USER FACING)
'   #VALUE!  DateIn not parseable / not scalar / unexpected runtime error
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.IsLeapYear_Core
'
' NOTES
'   - The previous per-function year gate is gone: TryResolveDate already
'     refuses
'     anything before KPR_MIN_DATE, so no year reaching this point can be out
'     of the supported range.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                         PUBLIC API - DATE ARITHMETIC                         
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
'     Date candidate handled by the facade TryResolveDate policy.
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
'   The integer argument is declared As Long, so Excel coerces a fractional
'   value before entry, using round-half-to-even. A fractional argument is
'   therefore silently rounded rather than rejected. That is contract-invalid:
'   INTEGER_FRACTION requires #VALUE!, and issue #12 replaces this with strict
'   Variant parsing.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'
' NOTES
'   - The shift is computed in Double and range-gated BEFORE coercion back to
'     Date, so a large nWeeks returns #NUM! rather than raising an overflow.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim ResultD         As Double    'Shifted date serial before range gating
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Compute the shifted serial in Double to avoid overflow on coercion
        ResultD = CDbl(ParsedDate) + (7# * CDbl(nWeeks))

    'Gate the result to the supported date window
        If (ResultD < CDbl(KPR_MIN_DATE)) Or (ResultD > CDbl(KPR_MAX_DATE)) Then
            FailErr = ErrNum()
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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
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
'   The integer argument is declared As Long, so Excel coerces a fractional
'   value before entry, using round-half-to-even. A fractional argument is
'   therefore silently rounded rather than rejected. That is contract-invalid:
'   INTEGER_FRACTION requires #VALUE!, and issue #12 replaces this with strict
'   Variant parsing.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.TryAddMonths_Core
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
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim ResultDate      As Date      'Shifted date returned by the month core
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths_Core(ParsedDate, nMonths, Opt_KeepEOM, ResultDate) Then
            FailErr = ErrNum()
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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
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
'   Integer arguments are declared As Long, so Excel coerces a fractional value
'   before entry, using round-half-to-even. A fractional argument is therefore
'   silently rounded rather than rejected. That is contract-invalid:
'   INTEGER_FRACTION requires #VALUE!, and issue #12 replaces this with strict
'   Variant parsing.
'
' DEPENDENCIES
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.TryAddMonths_Core
'
' NOTES
'   - Delegating to the month core keeps every year-roll edge case identical to
'     AddMonths. The one that matters:
'         29-Feb-2024 + 1Y => 28-Feb-2025 in both EOM modes
'   - The years -> months conversion is done in Double so a large nYears cannot
'     overflow the Long month delta before it is gated.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedDate      As Date      'Parsed date-only value (per TryResolveDate policy)
    Dim MonthsD         As Double    'Year -> month conversion before Long gating
    Dim ResultDate      As Date      'Shifted date returned by the month core
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUT
'------------------------------------------------------------------------------
    'Unwrap, parse and window-gate through the core boundaries
        If Not TryResolveDate(DateIn, ParsedDate) Then GoTo Fail

'------------------------------------------------------------------------------
' CONVERT YEARS TO MONTHS
'------------------------------------------------------------------------------
    'Convert in Double to avoid intermediate overflow
        MonthsD = 12# * CDbl(nYears)

    'Gate to Long range before coercion
        If (MonthsD < -2147483648#) Or (MonthsD > 2147483647#) Then
            FailErr = ErrNum()
            GoTo Fail
        End If

'------------------------------------------------------------------------------
' COMPUTE SHIFT
'------------------------------------------------------------------------------
    'Delegate to the single month shifter; failure here is always range-related
        If Not TryAddMonths_Core(ParsedDate, CLng(MonthsD), Opt_KeepEOM, ResultDate) Then
            FailErr = ErrNum()
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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                        PUBLIC API - WEEKDAY LOCATORS                         
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
'            month (typically a requested 5th weekday)''
'   Integer arguments are declared As Long, so Excel coerces a fractional value
'   before entry, using round-half-to-even. A fractional argument is therefore
'   silently rounded rather than rejected. That is contract-invalid:
'   INTEGER_FRACTION requires #VALUE!, and issue #12 replaces this with strict
'   Variant parsing.
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
'   2026-08-31
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
        FailErr = ErrValue()

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
            FailErr = ErrNum()
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
        FailErr = ErrValue()
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
'            unexpected runtime error''
'   Integer arguments are declared As Long, so Excel coerces a fractional value
'   before entry, using round-half-to-even. A fractional argument is therefore
'   silently rounded rather than rejected. That is contract-invalid:
'   INTEGER_FRACTION requires #VALUE!, and issue #12 replaces this with strict
'   Variant parsing.
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
'   2026-08-31
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
        FailErr = ErrValue()

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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                        PUBLIC API - PILLAR FORMATTING                        
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
'     Date candidate handled by the facade TryResolveDate policy.
'
'   EndDate
'     Date candidate handled by the facade TryResolveDate policy.
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
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.Pillar_Format_Nearest
'
' NOTES
'   - The label is a String by design; it is a category, not a quantity, and is
'     meant to be grouped or matched rather than computed on.
'   - EndDate before StartDate yields a "-" prefixed token.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedStart     As Date      'Parsed start date (per TryResolveDate policy)
    Dim ParsedEnd       As Date      'Parsed end date   (per TryResolveDate policy)
    Dim FailErr         As Variant   'Worksheet-facing error value returned on failure

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Route unexpected runtime errors to handler
        On Error GoTo Err_Handler

    'Default failure is #VALUE!
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUTS
'------------------------------------------------------------------------------
    'Parse the start date under module policy
        If Not TryResolveDate(StartDate, ParsedStart) Then GoTo Fail

    'Parse the end date under module policy
        If Not TryResolveDate(EndDate, ParsedEnd) Then GoTo Fail

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
        FailErr = ErrValue()
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
'     Date candidate handled by the facade TryResolveDate policy.
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
'   - TryResolveDate (KPR_Core_Array, KPR_Core_Parse, KPR_Core_Dates)
'   - KPR_Core_Dates.TryPillar_Parse
'   - KPR_Core_Dates.TryAddMonths_Core
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
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ParsedStart     As Date      'Parsed StartDate (per TryResolveDate policy)
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
        FailErr = ErrValue()

'------------------------------------------------------------------------------
' PARSE INPUTS
'------------------------------------------------------------------------------
    'Parse the start date under module policy
        If Not TryResolveDate(StartDate, ParsedStart) Then GoTo Fail

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
                    FailErr = ErrNum()
                    GoTo Fail
                End If

            'Delegate to the single month shifter (never preserve-EOM here)
                If Not TryAddMonths_Core(WorkDate, CLng(TotalMonths), False, WorkDate) Then
                    FailErr = ErrNum()
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
            FailErr = ErrNum()
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
        FailErr = ErrValue()
        Resume Fail

End Function

'
'------------------------------------------------------------------------------
'
'                              PRIVATE COMPOSITION
'
'------------------------------------------------------------------------------
'

Private Function TryResolveDate( _
    ByVal DateIn As Variant, _
    ByRef ParsedDate As Date) _
    As Boolean
'
'==============================================================================
'                                TryResolveDate
'------------------------------------------------------------------------------
' PURPOSE
'   Resolves one worksheet date argument into a normalized, in-window VBA Date
'   by composing the three core boundaries in a fixed order.
'
' WHY THIS EXISTS
'   The pre-split baseline combined shape, value and window inside a single
'   Parse_Date routine. The dependency matrix does not allow that: KPR_Core_Parse
'   may depend on KPR_Core_Err alone, so it can call neither the shape helper in
'   KPR_Core_Array nor the window constants in KPR_Core_Dates. Composition
'   therefore belongs to the facade, which is allowed to call all four cores.
'
' SIGNATURE
'   TryResolveDate(DateIn, ParsedDate) -> Boolean
'
' INPUTS
'   DateIn
'     Any Variant date argument, including a single-cell Range or 1x1 wrapper.
'
' OUTPUTS
'   ParsedDate (ByRef)
'     Assigned ONLY on success, normalized to date-only and inside the window.
'
' RETURNS
'   Boolean
'     TRUE  => resolution succeeded; ParsedDate assigned
'     FALSE => shape, value or window rejected the input
'
' BEHAVIOR
'   1. KPR_Core_Array.TryUnwrapScalar reduces wrappers to one scalar.
'   2. KPR_Core_Parse.TryParseDateScalar converts that scalar to a date.
'   3. KPR_Core_Dates.IsDateInWindow applies the supported window once.
'
' ERROR POLICY
'   - Does not raise. All failure paths return FALSE.
'   - The caller maps FALSE to a worksheet error value.
'
' DEPENDENCIES
'   - KPR_Core_Array.TryUnwrapScalar
'   - KPR_Core_Parse.TryParseDateScalar
'   - KPR_Core_Dates.IsDateInWindow
'
' NOTES
'   - The three failure causes are indistinguishable to the caller in this
'     revision, exactly as they were in the baseline. The contract separates
'     them into distinct conditions with distinct errors; issues #12 and #13
'     replace this Boolean with a classified outcome.
'
' UPDATED
'   2026-08-31
'==============================================================================
'

'------------------------------------------------------------------------------
' DECLARE
'------------------------------------------------------------------------------
    Dim ScalarValue     As Variant   'Payload after shape unwrapping
    Dim Candidate       As Date      'Parsed candidate before the window gate

'------------------------------------------------------------------------------
' INITIALIZE
'------------------------------------------------------------------------------
    'Default return is failure unless every stage succeeds
        TryResolveDate = False

'------------------------------------------------------------------------------
' RESOLVE SHAPE
'------------------------------------------------------------------------------
    'Reduce a single-cell Range or 1x1 wrapper to its scalar payload
        If Not TryUnwrapScalar(DateIn, ScalarValue) Then Exit Function

'------------------------------------------------------------------------------
' RESOLVE VALUE
'------------------------------------------------------------------------------
    'Convert the scalar to a date-only VBA Date
        If Not TryParseDateScalar(ScalarValue, Candidate) Then Exit Function

'------------------------------------------------------------------------------
' APPLY WINDOW
'------------------------------------------------------------------------------
    'One gate for every accepted path
        If Not IsDateInWindow(Candidate) Then Exit Function

'------------------------------------------------------------------------------
' ASSIGN RESULT
'------------------------------------------------------------------------------
    'Assign the output only once the value is fully validated
        ParsedDate = Candidate

    'Contract: TRUE only when ParsedDate was assigned
        TryResolveDate = True

End Function
